#!/usr/bin/env python3
"""
My SoftBank 料金明細PDF自動ダウンロードスクリプト

前提:
  - アカウント情報はGoogleスプレッドシート（サービスアカウント認証）から取得
  - PDFはGoogle Drive API経由で直接アップロード（またはローカルパスに保存）
  - 2段階認証(セキュリティ番号)はターミナルのinput()で手動入力
"""

import json
import os
import re
import subprocess
import sys
import time
import shutil
import webbrowser
import logging
import tempfile
from pathlib import Path
from datetime import datetime, timedelta
from dataclasses import dataclass, field
from typing import Any

import gspread
import pandas as pd
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# ─── 設定 ───────────────────────────────────────────────

def _bootstrap_env_from_gsheet() -> None:
    """
    .env が存在しない場合、同ディレクトリの *.gsheet から SPREADSHEET_URL を自動生成する。
    .gsheet は Google Drive が生成するショートカットファイルで doc_id を含む。
    """
    script_dir = Path(__file__).parent
    env_path = script_dir / ".env"
    if env_path.exists():
        return

    gsheet_files = list(script_dir.glob("*.gsheet"))
    if not gsheet_files:
        return

    try:
        data = json.loads(gsheet_files[0].read_text(encoding="utf-8"))
        doc_id = data.get("doc_id", "").strip()
        if not doc_id:
            return
    except Exception:
        return

    url = f"https://docs.google.com/spreadsheets/d/{doc_id}/edit?gid=0"
    env_path.write_text(
        "# ─── My SoftBank 料金明細ダウンロード 環境変数 ───\n"
        "# .gsheet ファイルから自動生成されました\n\n"
        f"SPREADSHEET_URL={url}\n\n"
        "# BASE_SAVE_PATH=\n"
        "# TARGET_MONTH=202602\n"
        "HEADLESS=true\n"
        "SECURITY_CODE_TIMEOUT=300\n",
        encoding="utf-8",
    )
    logging.getLogger(__name__).info(f".env を自動生成しました: {env_path}")

_bootstrap_env_from_gsheet()
load_dotenv()

# スプレッドシートURL or ID（環境変数 SPREADSHEET_URL または SPREADSHEET_ID で設定）
def _resolve_spreadsheet_id() -> str:
    url = os.environ.get("SPREADSHEET_URL", "").strip()
    if url:
        m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
        if m:
            return m.group(1)
    sid = os.environ.get("SPREADSHEET_ID", "").strip()
    if sid:
        return sid
    logging.getLogger(__name__).error(
        "スプレッドシートが設定されていません。.env に以下のいずれかを設定してください:\n"
        "  SPREADSHEET_URL=https://docs.google.com/spreadsheets/d/xxxxx/edit\n"
        "  SPREADSHEET_ID=xxxxx"
    )
    sys.exit(1)

SPREADSHEET_ID = _resolve_spreadsheet_id()
BASE_SAVE_PATH = os.environ.get("BASE_SAVE_PATH")
# ダウンロード対象月 (YYYYMM)。未指定時は前月
TARGET_MONTH = os.environ.get("TARGET_MONTH")
# 電話番号 → 運用端末名のマップ（同スプシ月シート読み込み後に自動設定）
_phone_device_map: dict[str, str] = {}
# ヘッドレスモード (デフォルト: true)
HEADLESS = os.environ.get("HEADLESS", "true").lower() in ("true", "1", "yes")
DRY_RUN = os.environ.get("DRY_RUN", "false").lower() == "true"
RETRY_PHONES = [p.strip() for p in os.environ.get("RETRY_PHONES", "").split(",") if p.strip()]
# セキュリティ番号入力のタイムアウト(秒)
SECURITY_CODE_TIMEOUT = int(os.environ.get("SECURITY_CODE_TIMEOUT", "300"))

LOGIN_URL = "https://my.softbank.jp/msb/d/webLink/doSend/MSB010000"
# ログイン後の実際のドメイン
WCO_BASE = "https://bl11.my.softbank.jp/wco"
# 書面発行ページ
CERTIFICATE_URL = f"{WCO_BASE}/certificate/WCO250"
# 自分で印刷する（無料）ページ
BILL_PDF_URL = f"{WCO_BASE}/external/goBillInfoPdf"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

CARRIER_NAME = "SoftBank"

# Drive APIモード時のコンテキスト（ローカルパスモード時はNone）
_drive_ctx: "DriveContext | None" = None
# Drive APIモード時のPDF一時保存ディレクトリ
_temp_save_dir: Path | None = None


# ─── Google Drive API ─────────────────────────────────────

@dataclass
class DriveContext:
    """Google Drive API を使ったフォルダ作成・存在確認・アップロードを担うクラス"""
    base_folder_id: str
    service: Any = field(init=False)
    _folder_cache: dict = field(default_factory=dict)

    def __post_init__(self):
        self.service = _get_drive_service()

    def _get_or_create_folder(self, parent_id: str, name: str) -> str:
        """指定名のサブフォルダを取得または作成してIDを返す"""
        key = f"{parent_id}/{name}"
        if key in self._folder_cache:
            return self._folder_cache[key]
        q = (
            f"name='{name}' and '{parent_id}' in parents "
            f"and mimeType='application/vnd.google-apps.folder' and trashed=false"
        )
        res = self.service.files().list(q=q, fields="files(id)").execute()
        files = res.get("files", [])
        if files:
            fid = files[0]["id"]
        else:
            meta = {
                "name": name,
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [parent_id],
            }
            fid = self.service.files().create(body=meta, fields="id").execute()["id"]
            log.info(f"  Driveフォルダを作成しました: {name}")
        self._folder_cache[key] = fid
        return fid

    def get_folder_id(self, year: str, month: str) -> str:
        """base_folder_id/YYYY/MM/SoftBank のフォルダIDを返す（なければ作成）"""
        year_id = self._get_or_create_folder(self.base_folder_id, year)
        month_id = self._get_or_create_folder(year_id, month)
        return self._get_or_create_folder(month_id, CARRIER_NAME)

    def file_exists(self, folder_id: str, year: str, month: str, phone: str) -> bool:
        """指定フォルダ内に対象ファイルが既に存在するか確認する"""
        prefix = f"{year}{month}_SoftBank_{phone}_"
        q = (
            f"name contains '{prefix}' and '{folder_id}' in parents "
            f"and mimeType='application/pdf' and trashed=false"
        )
        res = self.service.files().list(q=q, fields="files(name)").execute()
        files = res.get("files", [])
        if files:
            log.info(f"  既にDriveにアップロード済み: {files[0]['name']}  → スキップ")
            return True
        return False

    def upload(self, local_path: Path, folder_id: str) -> bool:
        """ローカルファイルをDriveフォルダにアップロードする。成功時True。"""
        try:
            media = MediaFileUpload(str(local_path), mimetype="application/pdf")
            meta = {"name": local_path.name, "parents": [folder_id]}
            self.service.files().create(
                body=meta, media_body=media, fields="id"
            ).execute()
            log.info(f"  Driveにアップロード完了: {local_path.name}")
            return True
        except Exception as e:
            if "storageQuotaExceeded" in str(e):
                return self._upload_local_fallback(local_path)
            log.error(f"  Driveアップロード失敗: {e}")
            return False

    def _upload_local_fallback(self, local_path: Path) -> bool:
        """サービスアカウントのストレージ制限時にプロジェクトフォルダへ直接コピーする"""
        try:
            # local_path 構造: {tmp}/{year}/{month}/{carrier}/{file}
            carrier = local_path.parent.name
            month   = local_path.parent.parent.name
            year    = local_path.parent.parent.parent.name
            dest_dir = Path(__file__).resolve().parent.parent / year / month / carrier
            dest_dir.mkdir(parents=True, exist_ok=True)
            dest = dest_dir / local_path.name
            shutil.copy2(local_path, dest)
            log.info(f"  ローカル保存（Google Drive同期）: {dest}")
            return True
        except Exception as e:
            log.error(f"  ローカル保存フォールバックも失敗: {e}")
            return False


_DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]


def _find_client_secrets() -> "Path | None":
    """client_secrets.json を探す（同ディレクトリ → 既知のキャリアフォルダ）"""
    script_dir = Path(__file__).resolve().parent
    parent = script_dir.parent
    for p in [
        script_dir / "client_secrets.json",
        parent / "SoftBank" / "client_secrets.json",
        parent / "Ymobile" / "client_secrets.json",
    ]:
        if p.exists():
            return p
    return None


def _bootstrap_client_secrets() -> None:
    """
    client_secrets.json が存在しない場合、サービスアカウントでアクセス可能な
    Driveから自動ダウンロードを試みる。
    事前に client_secrets.json を PDF保存先フォルダにアップロードしておくこと。
    """
    if _find_client_secrets():
        return
    try:
        json_path = _find_service_account_json()
    except FileNotFoundError:
        return
    try:
        import io
        from google.oauth2.service_account import Credentials as SACreds
        from googleapiclient.http import MediaIoBaseDownload
        creds = SACreds.from_service_account_file(
            str(json_path),
            scopes=["https://www.googleapis.com/auth/drive.readonly"],
        )
        svc = build("drive", "v3", credentials=creds)
        results = svc.files().list(
            q="name='client_secrets.json' and trashed=false",
            fields="files(id)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files = results.get("files", [])
        if not files:
            return
        dest = Path(__file__).resolve().parent / "client_secrets.json"
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, svc.files().get_media(fileId=files[0]["id"]))
        done = False
        while not done:
            _, done = downloader.next_chunk()
        dest.write_bytes(fh.getvalue())
        log.info(f"client_secrets.jsonをDriveから自動ダウンロードしました: {dest}")
    except Exception as e:
        log.debug(f"client_secrets.json自動ダウンロード失敗（手動配置が必要）: {e}")


def _get_drive_service():
    """Drive APIサービスを返す。client_secrets.jsonがあればOAuth、なければサービスアカウント。"""
    _bootstrap_client_secrets()
    secrets_path = _find_client_secrets()
    if secrets_path:
        return _get_drive_service_oauth(secrets_path)
    json_path = _find_service_account_json()
    creds = Credentials.from_service_account_file(
        str(json_path),
        scopes=_DRIVE_SCOPES,
    )
    return build("drive", "v3", credentials=creds)


def _get_drive_service_oauth(secrets_path: Path):
    """OAuth2認証でDrive APIサービスを返す（初回のみブラウザ認証）。"""
    from google.oauth2.credentials import Credentials as OAuthCreds
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow

    token_path = secrets_path.parent / "drive_oauth_token.json"
    creds = None

    if token_path.exists():
        creds = OAuthCreds.from_authorized_user_file(str(token_path), _DRIVE_SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            log.info("OAuthトークンを自動更新しました")
        else:
            log.info("ブラウザでGoogleアカウントの認証を行います...")
            flow = InstalledAppFlow.from_client_secrets_file(str(secrets_path), _DRIVE_SCOPES)
            creds = flow.run_local_server(port=0)
            log.info("OAuth認証が完了しました")
        token_path.write_text(creds.to_json(), encoding="utf-8")
        log.info(f"OAuthトークンを保存しました: {token_path}")

    return build("drive", "v3", credentials=creds)


# ─── ユーティリティ ──────────────────────────────────────

def get_target_month() -> tuple[str, str]:
    """対象の年月を (year, month) 文字列で返す"""
    if TARGET_MONTH:
        ym = TARGET_MONTH.strip()
        return ym[:4], ym[4:6]
    # 未指定なら前月
    today = datetime.today()
    first = today.replace(day=1)
    prev = first - timedelta(days=1)
    return str(prev.year), f"{prev.month:02d}"


def strip_hyphens(phone: str) -> str:
    """電話番号からハイフンを除去する (例: 090-4769-5015 → 09047695015)"""
    return re.sub(r"[-\s\u2010-\u2015\u2212\uFF0D]", "", phone)


def _find_service_account_json() -> Path:
    """service_account.json を探す（同ディレクトリ → 既知のキャリアフォルダ）"""
    script_dir = Path(__file__).resolve().parent
    parent = script_dir.parent
    candidates = [
        script_dir / "service_account.json",
        parent / "SoftBank" / "service_account.json",
        parent / "Ymobile" / "service_account.json",
    ]
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError(
        "service_account.json が見つかりません。\n"
        "SoftBank/ または Ymobile/ フォルダに配置してください。"
    )


def _guide_spreadsheet_sharing(spreadsheet_id: str) -> None:
    """スプレッドシートへの権限エラー時にサービスアカウントのメールをクリップボードにコピー＋ブラウザを開く"""
    try:
        sa_email = json.loads(_find_service_account_json().read_text())["client_email"]
    except Exception:
        sa_email = "（service_account.json を確認してください）"

    try:
        subprocess.run(["pbcopy"], input=sa_email.encode(), check=True)
        clipboard_msg = "（クリップボードにコピー済み）"
    except Exception:
        clipboard_msg = ""

    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
    print("\n" + "=" * 60)
    print("スプレッドシートの共有設定が必要です")
    print(f"\nサービスアカウントのメールアドレス{clipboard_msg}:")
    print(f"  {sa_email}")
    print("\nブラウザでスプレッドシートを開きます...")
    print("右上の「共有」ボタンをクリックし、上記メールを編集者として追加してください。")
    print("=" * 60)
    webbrowser.open(url)
    input("\n共有設定が完了したらEnterを押してください: ")


def get_gspread_client() -> gspread.Client:
    """サービスアカウント認証済みの gspread クライアントを返す"""
    json_path = _find_service_account_json()
    creds = Credentials.from_service_account_file(
        str(json_path),
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"],
    )
    return gspread.authorize(creds)


def _open_sheet(gc: gspread.Client) -> gspread.Spreadsheet:
    """スプレッドシートを開く。権限エラー時は共有案内を表示してリトライする。"""
    try:
        return gc.open_by_key(SPREADSHEET_ID)
    except (PermissionError, gspread.exceptions.APIError) as e:
        is_403 = isinstance(e, PermissionError) or (
            hasattr(e, "response") and e.response.status_code == 403
        )
        if is_403:
            _guide_spreadsheet_sharing(SPREADSHEET_ID)
            return gc.open_by_key(SPREADSHEET_ID)
        raise


def resolve_save_path() -> str:
    """PDF保存先パスを取得する。

    Google Drive URLの場合はDrive APIモードを初期化し "drive://{folder_id}" を返す。
    ローカルパスの場合はそのまま返す。

    優先順:
      1. スプレッドシート「設定」シートの「PDF保存先フォルダ」行
         - Google DriveのURL → Drive APIで直接アップロード
         - ローカルパス（/で始まる）→ ローカル保存
      2. 環境変数 BASE_SAVE_PATH
    """
    global _drive_ctx, _temp_save_dir

    # ── 1. 設定シートから取得 ──
    try:
        gc = get_gspread_client()
        sh = _open_sheet(gc)
        ws = sh.worksheet("設定")
        for row in ws.get_all_records():
            if str(row.get("設定名", "")).strip() == "PDF保存先フォルダ":
                save_path = str(row.get("値", "")).strip()
                if save_path and save_path.lower() != "nan":
                    if save_path.startswith("https://"):
                        m = re.search(r"/folders/([a-zA-Z0-9_-]+)", save_path)
                        if m:
                            folder_id = m.group(1)
                            log.info(f"Drive APIモードで初期化: フォルダID={folder_id}")
                            try:
                                _drive_ctx = DriveContext(folder_id)
                                _temp_save_dir = Path(tempfile.mkdtemp(prefix="softbank_pdf_"))
                                log.info(f"一時保存ディレクトリ: {_temp_save_dir}")
                                return f"drive://{folder_id}"
                            except Exception as e:
                                log.error(
                                    f"Google Drive APIの初期化に失敗しました: {e}\n"
                                    f"サービスアカウントにフォルダへのアクセス権限があるか確認してください。\n"
                                    f"フォルダID: {folder_id}\n"
                                    f"service_account.json のメールアドレスをDriveフォルダの共有設定に追加してください。"
                                )
                                sys.exit(1)
                        else:
                            log.warning(
                                f"Google DriveのURLからフォルダIDを抽出できませんでした\n"
                                f"  URL: {save_path}"
                            )
                    elif Path(save_path).is_absolute():
                        log.info(f"ローカル保存モード: {save_path}")
                        return save_path
    except Exception as e:
        log.error(f"設定シートからの保存先取得に失敗しました: {e}")
        sys.exit(1)

    # ── 2. 環境変数 ──
    if BASE_SAVE_PATH:
        return BASE_SAVE_PATH

    log.error(
        "PDF保存先が設定されていません。以下のいずれかを設定してください:\n"
        "  1. スプレッドシートの「設定」シートの「PDF保存先フォルダ」に記入\n"
        "     - Google DriveフォルダのURL（推奨）: https://drive.google.com/drive/folders/xxxxx\n"
        "       ※ service_account.json のメールアドレスをフォルダの共有設定（編集者）に追加してください\n"
        "     - Macのローカル絶対パス: /Users/.../マイドライブ/確定申告系/携帯領収書管理\n"
        "  2. 環境変数 BASE_SAVE_PATH を設定"
    )
    sys.exit(1)


def parse_pdf_types(raw: str) -> set[str]:
    """PDFの種類列の値をパースしてセットで返す。
    例: "電話番号別,一括" → {"電話番号別", "一括"}
    未指定・不正値の場合はデフォルト {"電話番号別"} を返す。
    """
    valid = {"電話番号別", "一括", "機種別"}
    if not raw or str(raw).strip().lower() in ("nan", ""):
        return {"電話番号別"}
    types = {t.strip() for t in str(raw).split(",") if t.strip()}
    result = types & valid
    if not result:
        log.warning(f"  PDFの種類 '{raw}' に有効な値がありません。デフォルト '電話番号別' を使用します。")
        return {"電話番号別"}
    return result


def _load_password_from_settings(sh) -> "str | None":
    """設定シートから共通パスワードを取得する。"""
    try:
        ws = sh.worksheet("設定")
        for row in ws.get_all_records():
            if str(row.get("設定名", "")).strip() == "パスワード":
                pw = str(row.get("値", "")).strip()
                if pw:
                    return pw
    except Exception:
        pass
    return None


def load_accounts() -> pd.DataFrame:
    """スプレッドシートからアカウント情報を読み込む。
    「認証情報」シートからキャリア列でフィルタ。パスワードは設定シートから取得。
    認証情報シートの列: 電話番号 | キャリア | PDFの種類 | 運用端末
    """
    log.info("スプレッドシートからアカウント情報を読み込み中...")
    gc = get_gspread_client()
    sh = _open_sheet(gc)

    # 設定シートから共通パスワードを取得
    common_password = _load_password_from_settings(sh)
    if not common_password:
        log.error(
            "パスワードが設定されていません。\n"
            "  設定シートの「パスワード」行にログインパスワードを入力してください。"
        )
        sys.exit(1)

    # 「認証情報」シートを読み込み
    df: "pd.DataFrame | None" = None
    df_all: "pd.DataFrame | None" = None
    try:
        ws = sh.worksheet("認証情報")
        records = ws.get_all_records()
        df_all = pd.DataFrame(records)
        df_all.columns = df_all.columns.str.strip()
        if "電話番号" in df_all.columns and "キャリア" in df_all.columns:
            df = df_all[df_all["キャリア"].str.strip() == CARRIER_NAME].reset_index(drop=True)
            log.info(f"  「認証情報」シートから {CARRIER_NAME} アカウントを読み込み")
        else:
            df = None
    except gspread.exceptions.WorksheetNotFound:
        df = None

    if df is None:
        log.error("「認証情報」シートが見つからないか、電話番号・キャリア列がありません。")
        sys.exit(1)

    if "電話番号" not in df.columns:
        log.error(f"スプレッドシートに「電話番号」カラムがありません。現在のカラム: {list(df.columns)}")
        sys.exit(1)

    # パスワードは設定シートから一律設定
    df["パスワード"] = common_password
    # 電話番号のハイフンを除去
    df["電話番号"] = df["電話番号"].astype(str).apply(strip_hyphens)
    # PDFの種類列がなければデフォルト値を補完
    if "PDFの種類" not in df.columns:
        df["PDFの種類"] = "電話番号別"
    # 解約済行を除外（GASが状態列に「解約済」と書き込む）
    if "状態" in df.columns:
        before = len(df)
        df = df[df["状態"].astype(str).str.strip() != "解約済"].reset_index(drop=True)
        cancelled = before - len(df)
        if cancelled > 0:
            log.info(f"  解約済 {cancelled} 件を除外")
    log.info(f"  {len(df)} 件のアカウントを読み込みました")

    # 認証情報シートの「運用端末」列から _phone_device_map を構築
    # （GASが回線管理表から自動同期した値をそのまま使用）
    global _phone_device_map
    _phone_device_map = {}
    if df_all is not None and "運用端末" in df_all.columns:
        for _, row in df_all.iterrows():
            phone = strip_hyphens(str(row.get("電話番号", "")))
            device = str(row.get("運用端末", "")).strip()
            if re.match(r'^\d{10,13}$', phone) and device:
                _phone_device_map[phone] = device
        if _phone_device_map:
            log.info(f"  運用端末マップ: {len(_phone_device_map)} 件")

    return df


def _write_download_history(results: list[tuple[str, bool]], year: str, month: str) -> None:
    """ダウンロード結果をスプレッドシートの「ダウンロード履歴」シートに記録する。"""
    try:
        gc = get_gspread_client()
        sh = _open_sheet(gc)
        try:
            ws = sh.worksheet("ダウンロード履歴")
        except gspread.exceptions.WorksheetNotFound:
            log.debug("ダウンロード履歴シートが見つかりません（スキップ）")
            return
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        rows = []
        for phone, ok in results:
            rows.append([now, CARRIER_NAME, phone, f"{year}{month}", "", "成功" if ok else "失敗"])
        if rows:
            ws.append_rows(rows, value_input_option="USER_ENTERED")
            log.info(f"  ダウンロード履歴を記録しました（{len(rows)}件）")
    except Exception as e:
        log.warning(f"ダウンロード履歴の記録に失敗（処理は続行）: {e}")


_TMPDIR = Path(tempfile.gettempdir())
CODE_FILE = _TMPDIR / "softbank_security_code.txt"


def _session_file(phone_number: str) -> Path:
    """電話番号ごとに独立したセッションファイルパスを返す（アカウント干渉防止）"""
    return _TMPDIR / f"softbank_session_{phone_number}.json"


def ask_security_code(phone_number: str) -> str | None:
    """セキュリティ番号を取得する。
    インタラクティブ環境: terminal input()
    非インタラクティブ環境: /tmp/softbank_security_code.txt をポーリング
    """
    # 念のため古いファイルを削除
    if CODE_FILE.exists():
        CODE_FILE.unlink()

    device = _phone_device_map.get(phone_number, "")
    print("\n" + "=" * 60)
    print(f"  📱 SMS認証が必要です")
    print(f"  電話番号: {phone_number}")
    if device:
        print(f"  端末    : {device}")
    print(f"  SMSに届いた3桁のセキュリティ番号を入力してください")
    print(f"  ターミナル入力 または 以下のコマンドで渡してください:")
    print(f"    echo '123' > {CODE_FILE}")
    print("=" * 60)

    import sys
    if sys.stdin.isatty():
        # インタラクティブ環境: 直接入力
        try:
            code = input("  セキュリティ番号（3桁）: ").strip()
            if code:
                return code
        except (EOFError, KeyboardInterrupt):
            pass
    else:
        # 非インタラクティブ環境 (Bashツール等): ファイルをポーリング
        log.info(f"  非インタラクティブ環境を検出。ファイルの出現を待機中...")
        log.info(f"  別ターミナルで: echo '123' > {CODE_FILE}")

    # ファイルのポーリング (最大SECURITY_CODE_TIMEOUT秒)
    deadline = time.time() + SECURITY_CODE_TIMEOUT
    last_log = time.time()
    while time.time() < deadline:
        if CODE_FILE.exists():
            try:
                code = CODE_FILE.read_text(encoding="utf-8").strip()
                CODE_FILE.unlink()
                if code:
                    log.info(f"  セキュリティ番号をファイルから取得しました")
                    return code
            except Exception:
                pass
        now = time.time()
        if now - last_log >= 15:
            remaining = int(deadline - now)
            log.info(f"  セキュリティ番号待機中... 残り{remaining}秒")
            last_log = now
        time.sleep(1)

    log.error("  セキュリティ番号の入力がタイムアウトしました")
    return None


def sanitize_amount(text: str) -> str:
    """請求金額テキストから数字だけ抽出して「○○円」形式にする"""
    digits = re.sub(r"[^\d]", "", text)
    if digits:
        return f"{int(digits)}円"
    return ""


def build_filename(year: str, month: str, phone: str, amount: str = "") -> str:
    """電子帳簿保存法準拠のファイル名を生成する"""
    base = f"{year}{month}_SoftBank_{phone}"
    if amount:
        base += f"_{amount}"
    else:
        base += "_利用料金明細"
    return base + ".pdf"


def _save_debug_screenshot(
    page, save_dir: Path, phone: str, year: str, month: str,
    error_type: str, error_detail: str,
) -> None:
    """デバッグ用スクリーンショットをローカルに保存する。
    Drive APIモード時はスクリプトの隣の debug/ フォルダに保存。
    ローカルモード時は save_dir の3階層上の debug/ フォルダに保存。
    """
    try:
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        if _drive_ctx is not None:
            # Drive APIモード: スクリプトと同じディレクトリの debug/ に保存
            base_path = Path(__file__).parent
        else:
            base_path = save_dir.parent.parent.parent  # base_path/YYYY/MM/SoftBank → base_path
        debug_dir = base_path / "debug" / f"debug_{now}"
        debug_dir.mkdir(parents=True, exist_ok=True)

        # スクリーンショット保存
        safe_type = re.sub(r'[\\/:*?"<>|]', "_", error_type)
        png_name = f"{phone}_{year}{month}_{safe_type}.png"
        png_path = debug_dir / png_name
        page.screenshot(path=str(png_path), full_page=True)
        log.error(f"  デバッグスクリーンショットを保存: {png_path}")

        # エラー詳細をテキストに記録
        detail_path = debug_dir / f"{phone}_{year}{month}_{safe_type}.txt"
        current_url = ""
        try:
            current_url = page.url
        except Exception:
            pass
        detail_path.write_text(
            f"電話番号: {phone}\n"
            f"対象月: {year}年{month}月\n"
            f"エラー種別: {error_type}\n"
            f"エラー詳細: {error_detail}\n"
            f"発生日時: {now}\n"
            f"ページURL: {current_url}\n",
            encoding="utf-8",
        )
    except Exception:
        pass


def ensure_save_dir(base_path: str, year: str, month: str) -> Path:
    """保存先ディレクトリを作成して返す (base_path/year/month/キャリア名)"""
    if base_path.startswith("drive://") and _temp_save_dir is not None:
        # Drive APIモード: 一時ディレクトリ内に同一構造を作成
        save_dir = _temp_save_dir / year / month / CARRIER_NAME
    else:
        save_dir = Path(base_path) / year / month / CARRIER_NAME
    save_dir.mkdir(parents=True, exist_ok=True)
    return save_dir


# ─── メイン処理 ──────────────────────────────────────────

def check_already_downloaded(save_dir: Path, year: str, month: str, phone: str) -> bool:
    """指定の電話番号・月のPDFが既にダウンロード済み（またはアップロード済み）かチェックする"""
    if _drive_ctx is not None:
        # Drive APIモード: Drive上で確認
        folder_id = _drive_ctx.get_folder_id(year, month)
        return _drive_ctx.file_exists(folder_id, year, month, phone)
    # ローカル保存モード
    pattern = f"{year}{month}_SoftBank_{phone}_*.pdf"
    existing = list(save_dir.glob(pattern))
    if existing:
        log.info(f"  既にダウンロード済み: {existing[0].name}  → スキップ")
        return True
    return False


def _js_click(page, selector: str) -> bool:
    """JavaScriptでクリックする（装飾divのオーバーレイ対策）"""
    try:
        page.evaluate(f'document.querySelector("{selector}")?.click()')
        return True
    except Exception:
        return False


def _click_any_button(page, locator, label: str = "ボタン", text_hint: str = "") -> bool:
    """ボタンをクリックする。force=True → テキスト検索 → JS click の順に試す。
    SoftBankのボタンは <a> タグ / <input> / <button> / <div> など多様なので全パターン対応。
    """
    # 1) locator で force click
    try:
        locator.first.click(force=True)
        log.info(f"  {label}をクリックしました (force)")
        return True
    except Exception as e:
        log.info(f"  {label} force click 失敗: {e}")

    # 2) テキストヒントがあればテキストで探してクリック
    if text_hint:
        try:
            text_elem = page.get_by_text(text_hint, exact=False).first
            if text_elem.is_visible(timeout=3000):
                text_elem.click(force=True)
                log.info(f"  {label}をクリックしました (テキスト検索: {text_hint})")
                return True
        except Exception:
            pass

    # 3) JSフォールバック: 幅広い要素を探す
    try:
        hint_js = text_hint.replace('"', '\\"') if text_hint else ""
        clicked = page.evaluate(f"""() => {{
            // input[type="submit"] or button[type="submit"]
            let btn = document.querySelector('input[type="submit"]')
                   || document.querySelector('button[type="submit"]');
            if (btn) {{ btn.click(); return 'submit'; }}
            // テキストヒントで <a> タグを探す
            const hint = "{hint_js}";
            if (hint) {{
                const links = document.querySelectorAll('a');
                for (const a of links) {{
                    if (a.textContent && a.textContent.includes(hint)) {{
                        a.click(); return 'link';
                    }}
                }}
                // 全要素から探す
                const all = document.querySelectorAll('*');
                for (const el of all) {{
                    if (el.textContent && el.textContent.trim() === hint) {{
                        el.click(); return 'element';
                    }}
                }}
            }}
            // フォームを直接submit
            const form = document.querySelector('form');
            if (form) {{ form.submit(); return 'form'; }}
            return null;
        }}""")
        if clicked:
            log.info(f"  {label}をクリックしました (JS: {clicked})")
            return True
    except Exception as e:
        log.error(f"  {label}のクリックに失敗: {e}")
    return False


def _wait_for_page_stable(page, timeout: int = 10) -> None:
    """ページが安定するまで待つ（networkidle + sleep）"""
    try:
        page.wait_for_load_state("networkidle", timeout=timeout * 1000)
    except Exception:
        pass
    time.sleep(2)


def _get_page_text(page) -> str:
    """ページのテキストを安全に取得する"""
    try:
        return page.evaluate("() => document.body.innerText || ''")
    except Exception:
        return ""


def _click_send_button(page) -> str | None:
    """「送信する」ボタンをJSでクリックする。クリックした要素の説明を返す。"""
    try:
        return page.evaluate("""() => {
            // <a>タグで「送信する」を含むもの
            for (const a of document.querySelectorAll('a')) {
                if ((a.textContent || '').includes('送信する')) {
                    a.click(); return 'a: ' + a.textContent.trim();
                }
            }
            // input[type="submit"] で「ログイン」以外
            const submit = document.querySelector('input[type="submit"]');
            if (submit && !(submit.value || '').includes('ログイン')) {
                submit.click(); return 'input-submit: ' + submit.value;
            }
            // button[type="submit"]
            for (const btn of document.querySelectorAll('button[type="submit"]')) {
                if (!(btn.textContent || '').includes('ログイン')) {
                    btn.click(); return 'button: ' + btn.textContent.trim();
                }
            }
            // セキュリティ関連フォームをsubmit
            for (const form of document.querySelectorAll('form')) {
                const action = form.action || '';
                if (action.includes('security') || action.includes('sendTel') || action.includes('confirm')) {
                    form.submit(); return 'form-submit: ' + action;
                }
            }
            return null;
        }""")
    except Exception:
        return None


def _handle_security_code_flow(page, phone_number: str, password: str = "") -> bool:
    """セキュリティ番号の送信→入力→確認の一連のフローを処理する。
    認証ページ(id.my.softbank.jp)にいる場合に呼ばれる。
    ループで各ページ状態を判定しながら処理する。成功時 True を返す。

    対応パターン:
      A: ラジオ選択 → 送信 → 確認ページ → 送信 → 入力 → 本人確認
      B: ラジオなし確認ページ → 送信 → 入力 → 本人確認
      C: 直接セキュリティ番号入力ページ
    """
    log.info("セキュリティ番号フロー開始...")
    last4 = phone_number[-4:]

    for attempt in range(8):  # 最大8ステップ
        _wait_for_page_stable(page)
        current_url = page.url
        log.info(f"  [ステップ{attempt+1}] URL: {current_url}")

        # 認証ページを離れていれば完了
        if "id.my.softbank.jp" not in current_url:
            log.info("  認証フロー完了！")
            return True

        page_text = _get_page_text(page)

        # ── パターンC: セキュリティ番号入力欄がある ──
        try:
            sec_input = (
                page.locator('input[maxlength="3"]')
                .or_(page.locator('input[name="securityNumber"]'))
                .or_(page.locator('input[name="security_number"]'))
                .or_(page.locator('input[name="confirmationNumber"]'))
            )
            if sec_input.first.is_visible(timeout=2000):
                log.info("  セキュリティ番号入力欄を検出")
                code = ask_security_code(phone_number)
                if not code:
                    log.warning("  セキュリティ番号が入力されませんでした")
                    return False
                sec_input.first.fill(code)
                log.info("  セキュリティ番号を入力しました")
                # 本人確認ボタンをクリック
                # NOTE: <a>タグは検索しない。FAQへの案内リンク（「本人確認について」等）を
                # 誤クリックしてFAQページに飛ぶバグを防ぐため、submitボタンのみを対象にする。
                try:
                    result = page.evaluate("""() => {
                        // input[type="submit"]（ログイン以外）を最優先
                        const s = document.querySelector('input[type="submit"]');
                        if (s && !(s.value || '').includes('ログイン')) {
                            s.click(); return 'submit: ' + s.value;
                        }
                        // button[type="submit"]
                        for (const btn of document.querySelectorAll('button[type="submit"]')) {
                            if (!(btn.textContent || '').includes('ログイン')) {
                                btn.click(); return 'button: ' + btn.textContent.trim();
                            }
                        }
                        // フォームをsubmit（最終手段）
                        const f = document.querySelector('form');
                        if (f) { f.submit(); return 'form-submit'; }
                        return null;
                    }""")
                    log.info(f"  本人確認ボタン: {result}")
                except Exception as e:
                    log.warning(f"  本人確認ボタンクリック失敗: {e}")
                _wait_for_page_stable(page, timeout=15)
                log.info(f"  本人確認後URL: {page.url}")
                # 認証ページを離れていれば完了
                if "id.my.softbank.jp" not in page.url:
                    log.info("  本人確認完了！認証フロー完了！")
                    return True
                continue
        except Exception:
            pass

        # ── ページ情報をログ出力（デバッグ用）──
        try:
            info = page.evaluate("""() => {
                const rows = [];
                document.querySelectorAll('form').forEach((f, i) =>
                    rows.push('form[' + i + ']: ' + f.action + ' ' + f.method));
                document.querySelectorAll('a, input[type="submit"], button').forEach(el => {
                    const t = (el.textContent || el.value || '').trim();
                    if (t) rows.push(el.tagName + ': ' + t.substring(0, 40));
                });
                return rows.join('\\n');
            }""")
            log.info(f"  ページ要素:\n{info}")
        except Exception:
            pass

        # ── パターンA: ラジオボタンがある（送付先選択） ──
        try:
            radios = page.locator('input[type="radio"]')
            radio_count = radios.count()
            if radio_count > 0:
                log.info(f"  ラジオボタン {radio_count}個 を検出（送付先選択）")
                selected = False
                for i in range(radio_count):
                    try:
                        label = radios.nth(i).locator("..").text_content() or ""
                        log.info(f"    ラジオ{i}: {label.strip()}")
                        if last4 in label:
                            radios.nth(i).check(force=True)
                            log.info(f"    末尾{last4}に一致 → 選択")
                            selected = True
                            break
                    except Exception:
                        pass
                if not selected:
                    log.info("    一致なし → 先頭を選択")
                    try:
                        radios.first.check(force=True)
                    except Exception:
                        pass
                time.sleep(1)
        except Exception:
            pass

        # ── 「送信する」ボタンをクリック（確認ページ・ラジオ後・ラジオなし共通） ──
        if ("セキュリティ番号" in page_text or "送付先" in page_text
                or "連絡先" in page_text or "本人確認" in page_text):
            result = _click_send_button(page)
            if result:
                log.info(f"  送信ボタンクリック: {result}")
                _wait_for_page_stable(page, timeout=12)
                continue
            else:
                log.warning("  送信ボタンが見つかりませんでした")

        # ── ログインフォームにいる場合: 資格情報を入力してログインボタンをクリック ──
        elif "id.my.softbank.jp" in current_url:
            try:
                pw_field = page.locator('input[type="password"]')
                if pw_field.first.is_visible(timeout=1000):
                    log.info("  ログインフォームを検出 → 認証情報を入力してログイン")
                    phone_field = page.locator('input[name="telnum"]')
                    if phone_field.first.is_visible(timeout=1000):
                        phone_field.first.fill(phone_number)
                    if password:
                        pw_field.first.fill(password)
                    page.evaluate("""() => {
                        const s = document.querySelector('input[type="submit"]');
                        if (s) { s.click(); return; }
                        const f = document.querySelector('form');
                        if (f) f.submit();
                    }""")
                    _wait_for_page_stable(page, timeout=15)
                    continue
            except Exception:
                pass
            log.info("  ページ遷移待ち中...")
            time.sleep(3)
            continue

        # ループ上限に達しそうなら終了
        if attempt >= 7:
            break

    if "id.my.softbank.jp" in page.url:
        log.error("  まだ認証ページにいます")
        return False

    log.info("  認証フロー完了！")
    return True


def _is_on_auth_page(page) -> bool:
    """現在のページが認証ページ(id.my.softbank.jp)かどうか判定する"""
    return "id.my.softbank.jp" in page.url


def do_login_and_navigate(page, phone_number: str, password: str) -> bool:
    """ログイン → 2FA → PDFダウンロードページまで一気に遷移する。

    戦略: PDFページ(BILL_PDF_URL)に直接アクセスし、SoftBankが自動で
    認証ページにリダイレクトするのを利用する。認証完了後にPDFページに
    自動リダイレクトされるため、ログインとページ遷移を一度に処理できる。
    """
    # ── Step 1: PDFページに直接アクセス（認証ページにリダイレクトされる） ──
    log.info("PDFページに直接アクセス中（認証ページにリダイレクトされるはず）...")
    page.goto(BILL_PDF_URL, wait_until="networkidle")
    page.wait_for_load_state("networkidle")
    time.sleep(2)
    log.info(f"  リダイレクト先URL: {page.url}")

    # 認証不要で直接PDFページに到達した場合（既存セッションがある場合）
    if not _is_on_auth_page(page):
        log.info("  認証なしでPDFページに到達しました")
        return True

    # ── Step 2: ページ状態を確認してログインまたは2FAフローへ ──
    page_text = _get_page_text(page)

    # セキュリティ番号ページに既にいる場合（セッション再利用時）はログインをスキップ
    if "セキュリティ番号" in page_text or "送付先" in page_text:
        log.info("  セキュリティ番号ページを検出（セッション再利用）→ ログインをスキップして2FAフローへ")
        if not _handle_security_code_flow(page, phone_number, password):
            return False
    else:
        log.info("ログイン情報を入力中...")

        # ログインフォームがあるか確認
        phone_input = (
            page.locator('input[name="telnum"]')
            .or_(page.locator('input.sbid-msn-check'))
            .or_(page.locator('input[name="username"]'))
            .or_(page.locator('input[name="msn"]'))
            .or_(page.locator('input[name="loginId"]'))
        )

        try:
            phone_input.first.wait_for(state="visible", timeout=10000)
            phone_input.first.fill(phone_number)
            log.info("  電話番号を入力しました")

            pw_input = page.locator('input[type="password"]')
            pw_input.first.wait_for(state="visible", timeout=10000)
            pw_input.first.fill(password)
            log.info("  パスワードを入力しました")

            login_btn = (
                page.locator('input[type="submit"]')
                .or_(page.get_by_text("ログインする", exact=False))
                .or_(page.get_by_role("link", name=re.compile(r"ログイン")))
                .or_(page.get_by_role("button", name=re.compile(r"ログイン")))
                .or_(page.locator('button[type="submit"]'))
            )
            _click_any_button(page, login_btn, "ログインボタン", text_hint="ログイン")
            # ログインフォーム（電話番号入力欄）が消えるまで待つ
            try:
                page.wait_for_function(
                    "() => !document.querySelector('input[name=\"telnum\"]')",
                    timeout=15000,
                )
                log.info("  ログインフォームが消えました（認証処理完了）")
            except Exception:
                pass
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except Exception:
                pass
            time.sleep(3)
            log.info(f"  ログイン後のURL: {page.url}")
        except (PlaywrightTimeout, Exception) as e:
            log.error(f"  ログインフォームの操作に失敗: {e}")
            return False

    # ── Step 3: ログイン後にエラーがないか確認 ──
    if _is_on_auth_page(page):
        # エラーメッセージ確認
        error_msg = page.locator(".err-area, .error, .alert-error, .sbid-error")
        if error_msg.count() > 0:
            try:
                err_text = error_msg.first.text_content()
                log.error(f"  ログインエラー: {err_text}")
            except Exception:
                pass

        # ログインは成功したが、セキュリティ番号が必要（1回目の2FA）
        log.info("  認証ページにいます → セキュリティ番号フローを処理します（1回目）")
        if not _handle_security_code_flow(page, phone_number, password):
            return False

    log.info(f"  1回目の認証後のURL: {page.url}")

    # ── Step 4: PDFページに到達したか確認 ──
    # 1回目の認証後にMy SoftBankトップに飛ばされることがある
    # その場合、書面発行ページ→PDFページへの遷移が追加の認証(acr_value=2)を要求する
    if not _is_on_auth_page(page):
        # combobox（月選択）があればPDFページに到達
        try:
            combobox = page.get_by_role("combobox").first
            if combobox.is_visible(timeout=3000):
                log.info("  PDFダウンロードページに到達しました！")
                return True
        except Exception:
            pass

        # PDFページではない場合、書面発行ページへ遷移を試みる
        log.info("  PDFページではありません。書面発行ページへ遷移します...")
        try:
            # まずリンクを探す
            cert_link = page.locator('a[href*="/wco/certificate/WCO250"]')
            if cert_link.count() > 0:
                cert_link.first.click()
            else:
                page.goto(CERTIFICATE_URL, wait_until="networkidle")
            page.wait_for_load_state("networkidle")
            time.sleep(2)
            log.info(f"  書面発行ページ遷移後のURL: {page.url}")
        except Exception:
            pass

        # 自分で印刷するページへ
        try:
            pdf_page_link = page.locator('a[href*="goBillInfoPdf"]')
            if pdf_page_link.count() > 0:
                pdf_page_link.first.click()
            else:
                print_link = page.get_by_text(re.compile(r"自分で印刷"))
                if print_link.count() > 0:
                    print_link.first.click()
                else:
                    page.goto(BILL_PDF_URL, wait_until="networkidle")
            page.wait_for_load_state("networkidle")
            time.sleep(2)
            log.info(f"  自分で印刷ページ遷移後のURL: {page.url}")
        except Exception:
            pass

    # ── Step 5: 2回目の認証が必要な場合 ──
    if _is_on_auth_page(page):
        log.info("  2回目の認証が必要です → セキュリティ番号フローを処理します")

        # 2回目はログインフォームではなく直接セキュリティ番号画面のはず
        # ただしログインフォームが出る場合もある
        try:
            phone_field = page.locator('input[name="telnum"]')
            if phone_field.first.is_visible(timeout=3000):
                log.info("  2回目の認証にもログインが必要です")
                phone_field.first.fill(phone_number)
                pw_field = page.locator('input[type="password"]')
                pw_field.first.fill(password)
                login_btn2 = page.locator('input[type="submit"]').or_(page.get_by_text("ログインする", exact=False))
                _click_any_button(page, login_btn2, "ログインボタン(2回目)", text_hint="ログイン")
                page.wait_for_load_state("networkidle")
                time.sleep(3)
        except Exception:
            pass

        # セキュリティ番号フロー（2回目）
        if _is_on_auth_page(page):
            if not _handle_security_code_flow(page, phone_number, password):
                return False

    # ── 最終確認 ──
    final_url = page.url
    log.info(f"  最終URL: {final_url}")

    if _is_on_auth_page(page):
        log.error("  認証を完了できませんでした")
        return False

    # combobox確認
    try:
        combobox = page.get_by_role("combobox").first
        if combobox.is_visible(timeout=5000):
            log.info("  PDFダウンロードページに到達しました！")
            return True
    except Exception:
        pass

    # PDFリンク確認
    try:
        pdf_link = page.locator('a[href*="doPrint"]')
        if pdf_link.count() > 0:
            log.info("  PDFダウンロードページに到達しました（PDFリンク検出）")
            return True
    except Exception:
        pass

    # WCOドメインにはいるが、PDFページではない → リンクから遷移を試みる
    if "bl11.my.softbank.jp/wco" in final_url:
        log.info("  WCOドメインにいますが、PDFページではありません。再遷移を試みます...")
        try:
            page.goto(BILL_PDF_URL, wait_until="networkidle")
            time.sleep(2)
            combobox = page.get_by_role("combobox").first
            if combobox.is_visible(timeout=5000):
                log.info("  PDFダウンロードページに到達しました！")
                return True
        except Exception:
            pass

    log.error(f"  PDFダウンロードページへの到達に失敗 (URL: {final_url})")
    return False


def select_target_month(page, year: str, month: str) -> bool:
    """PDFページで対象月をコンボボックスから選択する。"""
    target_value = f"{year}{month}"  # 例: "202602"
    target_label = f"{year}年{month}月"
    log.info(f"対象月を選択中: {target_label} (value={target_value})")

    try:
        combobox = page.get_by_role("combobox").first
        combobox.wait_for(state="visible", timeout=10000)

        # value属性で選択を試みる
        try:
            combobox.select_option(value=target_value)
            log.info(f"  月を選択しました: {target_value}")
        except Exception:
            # ラベルで選択
            combobox.select_option(label=re.compile(target_label))
            log.info(f"  月をラベルで選択しました: {target_label}")

        # 月選択後にページが更新される場合を待つ
        page.wait_for_load_state("networkidle")
        time.sleep(2)
        return True
    except Exception as e:
        log.error(f"対象月の選択に失敗しました: {e}")
        return False


def _download_single_pdf(page, link, label: str, save_dir: Path,
                          year: str, month: str, phone: str, amount: str) -> bool:
    """PDFリンクを1件ダウンロードして保存（またはDriveにアップロード）する。成功時True。"""
    try:
        link_text = link.text_content() or ""
        href = link.get_attribute("href") or ""
        log.info(f"  [{label}] PDFリンク: {link_text.strip()} ({href})")
        with page.expect_download(timeout=60000) as download_info:
            link.click()
        download = download_info.value
        filename = build_filename(year, month, phone, amount)
        # 同名ファイルが既にある場合はサフィックスを付ける
        dest = save_dir / filename
        if dest.exists():
            stem = dest.stem
            dest = save_dir / f"{stem}_{label}.pdf"
        download.save_as(str(dest))

        if _drive_ctx is not None:
            # Drive APIモード: アップロード後にローカル一時ファイルを削除
            folder_id = _drive_ctx.get_folder_id(year, month)
            ok = _drive_ctx.upload(dest, folder_id)
            try:
                dest.unlink()
            except Exception:
                pass
            return ok

        log.info(f"  [{label}] 保存完了: {dest}")
        return True
    except Exception as e:
        log.warning(f"  [{label}] ダウンロード失敗: {e}")
        return False


def download_pdf_from_page(
    page, save_dir: Path, year: str, month: str, phone: str,
    pdf_types: set[str] | None = None,
) -> bool:
    """PDFダウンロードページから指定種別のPDFをダウンロードする。

    pdf_types: {"電話番号別", "一括", "機種別"} の部分集合。
      電話番号別 … 電話番号別PDF (doPrintMsn idx=2)
      一括       … 一括印刷用PDF (doPrintSbmAll)
      機種別     … 機種別PDF    (doPrintMsn idx=3)
    未指定時はデフォルト {"電話番号別"}。
    """
    if pdf_types is None:
        pdf_types = {"電話番号別"}

    log.info(f"ダウンロード対象: {', '.join(sorted(pdf_types))}")

    # ── 請求金額の取得 ──
    amount = ""
    try:
        amount_el = page.locator("text=/[\\d,]+円/").first
        if amount_el.is_visible(timeout=3000):
            raw_amount = amount_el.text_content()
            if raw_amount:
                amount = sanitize_amount(raw_amount)
                log.info(f"  請求金額を取得: {amount}")
    except Exception:
        log.info("  請求金額の取得をスキップしました")

    any_success = False

    # ── 一括印刷用PDF ──
    if "一括" in pdf_types:
        log.info("一括印刷用PDFを確認中...")
        bulk_link = page.locator('a[href*="doPrintSbmAll"]')
        if bulk_link.count() > 0 and bulk_link.first.is_visible():
            if _download_single_pdf(page, bulk_link.first, "一括", save_dir, year, month, phone, amount):
                any_success = True
        else:
            log.info("  一括印刷用PDFリンクが見つかりませんでした")

    # ── 電話番号別PDF / 機種別PDF ──
    # idx番号はSoftBankのページ変更で入れ替わることがあるため、リンクテキストで種別を判定する
    if "電話番号別" in pdf_types or "機種別" in pdf_types:
        msn_links = page.locator('a[href*="doPrintMsn"]')
        if msn_links.count() > 0:
            for i in range(msn_links.count()):
                link = msn_links.nth(i)
                link_text = (link.text_content() or "").strip()
                log.info(f"  doPrintMsnリンク[{i}]: {link_text}")
                if "電話番号別" in link_text and "電話番号別" in pdf_types:
                    if _download_single_pdf(page, link, "電話番号別", save_dir, year, month, phone, amount):
                        any_success = True
                elif "機種別" in link_text and "機種別" in pdf_types:
                    if _download_single_pdf(page, link, "機種別", save_dir, year, month, phone, amount):
                        any_success = True
        else:
            log.info("  電話番号別/機種別PDFリンクが見つかりませんでした")

    if not any_success:
        log.error("指定した種別のPDFがダウンロードできませんでした")
    return any_success


def download_billing_pdf(
    phone_number: str,
    password: str,
    year: str,
    month: str,
    save_dir: Path,
    pdf_types: set[str] | None = None,
) -> bool:
    """1アカウント分のPDFダウンロード処理。
    ログイン → 書面発行 → 自分で印刷する → 月選択 → PDFダウンロード
    """
    if pdf_types is None:
        pdf_types = {"電話番号別"}
    log.info(f"=== {phone_number} の処理を開始 (PDFの種類: {', '.join(sorted(pdf_types))}) ===")

    # 既にダウンロード済みならスキップ
    if check_already_downloaded(save_dir, year, month, phone_number):
        return True

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)

        # セッションファイルがあれば読み込んで再利用（毎回SMS認証を不要にする）
        ctx_kwargs = dict(
            accept_downloads=True,
            locale="ja-JP",
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
        )
        session_file = _session_file(phone_number)
        if session_file.exists():
            log.info(f"  保存済みセッションを読み込み: {session_file}")
            ctx_kwargs["storage_state"] = str(session_file)
        context = browser.new_context(**ctx_kwargs)
        page = context.new_page()
        page.set_default_timeout(30000)

        success = False
        try:
            # ── ログイン → PDFページまで一気に遷移 ──
            if not do_login_and_navigate(page, phone_number, password):
                _save_debug_screenshot(
                    page, save_dir, phone_number, year, month,
                    "ログインまたはページ遷移失敗",
                    f"PDFダウンロードページに到達できませんでした。最終URL: {page.url}",
                )
                return False

            # ── セッション保存（次回以降の認証をスキップするため） ──
            try:
                context.storage_state(path=str(session_file))
                log.info(f"  セッションを保存しました: {session_file}")
            except Exception as e:
                log.warning(f"  セッション保存に失敗: {e}")

            # ── 対象月を選択 ──
            if not select_target_month(page, year, month):
                _save_debug_screenshot(
                    page, save_dir, phone_number, year, month,
                    "月選択失敗", f"対象月 {year}{month} の選択に失敗しました",
                )
                return False

            # ── PDFダウンロード ──
            success = download_pdf_from_page(page, save_dir, year, month, phone_number, pdf_types)
            if not success:
                _save_debug_screenshot(
                    page, save_dir, phone_number, year, month,
                    "PDFダウンロード失敗", "PDFリンクが見つからないかダウンロードに失敗しました",
                )

        except PlaywrightTimeout as e:
            log.error(f"タイムアウトエラー: {e}")
            _save_debug_screenshot(page, save_dir, phone_number, year, month, "タイムアウト", str(e))
        except Exception as e:
            log.error(f"エラーが発生しました: {e}")
            _save_debug_screenshot(page, save_dir, phone_number, year, month, "エラー", str(e))
        finally:
            context.close()
            browser.close()

    return success


def main():
    """メイン関数"""
    log.info("My SoftBank 料金明細PDFダウンロードを開始します")

    # 対象月の決定（優先順: 環境変数 TARGET_MONTH → 設定シート → 前月自動）
    year, month = get_target_month()
    if not TARGET_MONTH:
        try:
            gc = get_gspread_client()
            sh = _open_sheet(gc)
            ws = sh.worksheet("設定")
            for row in ws.get_all_records():
                if str(row.get("設定名", "")).strip() == "対象月":
                    val = str(row.get("値", "")).strip()
                    # YYYYMM形式
                    if re.match(r"^\d{6}$", val):
                        year, month = val[:4], val[4:6]
                        log.info(f"対象月（設定シートから取得）: {year}年{month}月")
                    # 「YYYY年M月」形式（ドロップダウン選択値）
                    elif m2 := re.match(r"^(\d{4})年(\d+)月$", val):
                        year, month = m2.group(1), m2.group(2).zfill(2)
                        log.info(f"対象月（設定シートから取得）: {year}年{month}月")
                    # 「自動（前月）」または空欄はデフォルト（前月）のまま
                    break
        except Exception as e:
            log.warning(f"設定シートから対象月の取得に失敗: {e}")
    log.info(f"対象月: {year}年{month}月")

    # 保存先の決定（Drive APIモード or ローカルモード）
    base_path = resolve_save_path()
    save_dir = ensure_save_dir(base_path, year, month)
    mode_label = "Drive APIモード" if _drive_ctx else "ローカルモード"
    log.info(f"保存先: {save_dir} ({mode_label})")

    # アカウント情報の読み込み
    accounts = load_accounts()

    if DRY_RUN:
        log.info("=== ドライランモード ===")
        log.info(f"  保存先: {save_dir} ({mode_label})")
        log.info(f"  対象月: {year}年{month}月")
        log.info(f"  対象アカウント: {len(accounts)} 件")
        for _, row in accounts.iterrows():
            phone = str(row["電話番号"]).strip()
            pdf_types = parse_pdf_types(row.get("PDFの種類", "電話番号別"))
            log.info(f"    {phone} ({', '.join(pdf_types)})")
        log.info("ドライランのため実際のダウンロードは行いません。")
        return

    if RETRY_PHONES:
        before = len(accounts)
        accounts = accounts[accounts["電話番号"].isin(RETRY_PHONES)].reset_index(drop=True)
        log.info(f"  リトライモード: {len(accounts)}/{before} 件に絞り込み")

    # 各アカウントについてPDFをダウンロード
    results = []
    for _, row in accounts.iterrows():
        phone = str(row["電話番号"]).strip()
        pw = str(row["パスワード"]).strip()
        pdf_types = parse_pdf_types(row.get("PDFの種類", "電話番号別"))

        ok = download_billing_pdf(phone, pw, year, month, save_dir, pdf_types)
        results.append((phone, ok))

    # 結果サマリー
    log.info("=" * 50)
    log.info("処理結果サマリー:")
    for phone, ok in results:
        status = "✅ 成功" if ok else "❌ 失敗"
        log.info(f"  {phone}: {status}")

    succeeded = sum(1 for _, ok in results if ok)
    log.info(f"  合計: {succeeded}/{len(results)} 件成功")
    log.info("=" * 50)

    # ダウンロード履歴をスプレッドシートに記録
    _write_download_history(results, year, month)

    # Drive APIモード: 一時ディレクトリのクリーンアップ
    if _temp_save_dir is not None and _temp_save_dir.exists():
        try:
            shutil.rmtree(str(_temp_save_dir))
            log.info(f"一時ディレクトリを削除しました: {_temp_save_dir}")
        except Exception as e:
            log.warning(f"一時ディレクトリの削除に失敗: {e}")

    failed = [p for p, ok in results if not ok]
    if failed:
        log.warning(f"{len(failed)} 件の失敗がありました")
        sys.exit(1)
    else:
        log.info("全アカウントの処理が正常に完了しました")


if __name__ == "__main__":
    main()
