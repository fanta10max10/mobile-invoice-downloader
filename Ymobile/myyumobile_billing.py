#!/usr/bin/env python3
"""
My Y!mobile 料金明細PDF自動ダウンロードスクリプト

前提:
  - アカウント情報はGoogleスプレッドシート（サービスアカウント認証）から取得
  - PDFはGoogle Drive API経由で直接アップロード（またはローカルパスに保存）
  - 2段階認証(セキュリティ番号)はターミナルのinput()で手動入力
  - 認証基盤はSoftBank IDと共通（id.my.ymobile.jp）
"""

import json
import os
import re
import subprocess
import sys
import time
import webbrowser
import shutil
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
        "# ─── My Y!mobile 料金明細ダウンロード 環境変数 ───\n"
        "# .gsheet ファイルから自動生成されました\n\n"
        f"SPREADSHEET_URL={url}\n\n"
        "# BASE_SAVE_PATH=\n"
        "# TARGET_MONTH=202602\n"
        "HEADLESS=true\n"
        "SECURITY_CODE_TIMEOUT=300\n",
        encoding="utf-8",
    )
    logging.getLogger(__name__).info(f".env を自動生成しました: {env_path}")

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
        logging.getLogger(__name__).info(f"client_secrets.jsonをDriveから自動ダウンロードしました: {dest}")
    except Exception as e:
        logging.getLogger(__name__).debug(f"client_secrets.json自動ダウンロード失敗（手動配置が必要）: {e}")


_bootstrap_env_from_gsheet()
load_dotenv()

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
TARGET_MONTH = os.environ.get("TARGET_MONTH")
HEADLESS = os.environ.get("HEADLESS", "true").lower() in ("true", "1", "yes")
SECURITY_CODE_TIMEOUT = int(os.environ.get("SECURITY_CODE_TIMEOUT", "300"))
# 電話番号 → 運用端末名のマップ（同スプシ月シート読み込み後に自動設定）
_phone_device_map: dict[str, str] = {}

# ─── Y!mobile URL定義 ───────────────────────────────────
# MySoftBankとの主な違い:
#   認証: id.my.ymobile.jp (SoftBankは id.my.softbank.jp)
#   ポータル: my.ymobile.jp (SoftBankは my.softbank.jp)
#   WCOシステム: bl61.my.ymobile.jp (SoftBankは bl11.my.softbank.jp)

YMB_PORTAL = "https://my.ymobile.jp"
AUTH_DOMAIN = "id.my.ymobile.jp"

# 書面発行ページへのディープリンク（ログイン後にWCOシステムへ誘導）
LOGIN_URL = f"{YMB_PORTAL}/muc/d/webLink/doSend/WCO010023"

WCO_BASE = "https://bl61.my.ymobile.jp/wco"
CERTIFICATE_URL = f"{WCO_BASE}/certificate/WCO250"
BILL_PDF_URL = f"{WCO_BASE}/external/goBillInfoPdf"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

CARRIER_NAME = "Ymobile"

_drive_ctx: "DriveContext | None" = None
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
        year_id = self._get_or_create_folder(self.base_folder_id, year)
        month_id = self._get_or_create_folder(year_id, month)
        return self._get_or_create_folder(month_id, CARRIER_NAME)

    def file_exists(self, folder_id: str, year: str, month: str, phone: str) -> bool:
        prefix = f"{year}{month}_Ymobile_{phone}_"
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
    if TARGET_MONTH:
        ym = TARGET_MONTH.strip()
        return ym[:4], ym[4:6]
    today = datetime.today()
    first = today.replace(day=1)
    prev = first - timedelta(days=1)
    return str(prev.year), f"{prev.month:02d}"


def strip_hyphens(phone: str) -> str:
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
    global _drive_ctx, _temp_save_dir

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
                                _temp_save_dir = Path(tempfile.mkdtemp(prefix="ymobile_pdf_"))
                                log.info(f"一時保存ディレクトリ: {_temp_save_dir}")
                                return f"drive://{folder_id}"
                            except Exception as e:
                                log.error(
                                    f"Google Drive APIの初期化に失敗しました: {e}\n"
                                    f"service_account.json のメールアドレスをDriveフォルダの共有設定（編集者）に追加してください。\n"
                                    f"フォルダID: {folder_id}"
                                )
                                sys.exit(1)
                        else:
                            log.warning(f"Google DriveのURLからフォルダIDを抽出できませんでした\n  URL: {save_path}")
                    elif Path(save_path).is_absolute():
                        log.info(f"ローカル保存モード: {save_path}")
                        return save_path
    except RecursionError:
        raise
    except Exception as e:
        log.warning(f"設定シートからの保存先取得に失敗: {type(e).__name__}: {e}")

    if BASE_SAVE_PATH:
        return BASE_SAVE_PATH

    log.error(
        "PDF保存先が設定されていません。以下のいずれかを設定してください:\n"
        "  1. スプレッドシートの「設定」シートの「PDF保存先フォルダ」に記入\n"
        "     - Google DriveフォルダのURL（推奨）: https://drive.google.com/drive/folders/xxxxx\n"
        "       ※ service_account.json のメールアドレスをフォルダの共有設定（編集者）に追加してください\n"
        "  2. 環境変数 BASE_SAVE_PATH を設定"
    )
    sys.exit(1)


def parse_pdf_types(raw: str) -> set[str]:
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
    df["電話番号"] = df["電話番号"].astype(str).apply(strip_hyphens)
    if "PDFの種類" not in df.columns:
        df["PDFの種類"] = "電話番号別"
    # 解約済行を除外（GASがPDFの種類列に「解約済」と書き込む）
    before = len(df)
    df = df[df["PDFの種類"].str.strip() != "解約済"].reset_index(drop=True)
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


_TMPDIR = Path(tempfile.gettempdir())
CODE_FILE = _TMPDIR / "ymobile_security_code.txt"


def _session_file(phone_number: str) -> Path:
    return _TMPDIR / f"ymobile_session_{phone_number}.json"


def ask_security_code(phone_number: str) -> str | None:
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
        try:
            code = input("  セキュリティ番号（3桁）: ").strip()
            if code:
                return code
        except (EOFError, KeyboardInterrupt):
            pass
    else:
        log.info(f"  非インタラクティブ環境を検出。ファイルの出現を待機中...")
        log.info(f"  別ターミナルで: echo '123' > {CODE_FILE}")

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
    digits = re.sub(r"[^\d]", "", text)
    if digits:
        return f"{int(digits)}円"
    return ""


def build_filename(year: str, month: str, phone: str, amount: str = "") -> str:
    base = f"{year}{month}_Ymobile_{phone}"
    if amount:
        base += f"_{amount}"
    else:
        base += "_利用料金明細"
    return base + ".pdf"


def _save_debug_screenshot(
    page, save_dir: Path, phone: str, year: str, month: str,
    error_type: str, error_detail: str,
) -> None:
    try:
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        if _drive_ctx is not None:
            base_path = Path(__file__).parent
        else:
            base_path = save_dir.parent.parent.parent
        debug_dir = base_path / "debug" / f"debug_{now}"
        debug_dir.mkdir(parents=True, exist_ok=True)

        safe_type = re.sub(r'[\\/:*?"<>|]', "_", error_type)
        png_name = f"{phone}_{year}{month}_{safe_type}.png"
        png_path = debug_dir / png_name
        page.screenshot(path=str(png_path), full_page=True)
        log.error(f"  デバッグスクリーンショットを保存: {png_path}")

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
    if base_path.startswith("drive://") and _temp_save_dir is not None:
        save_dir = _temp_save_dir / year / month / CARRIER_NAME
    else:
        save_dir = Path(base_path) / year / month / CARRIER_NAME
    save_dir.mkdir(parents=True, exist_ok=True)
    return save_dir


# ─── メイン処理 ──────────────────────────────────────────

def check_already_downloaded(save_dir: Path, year: str, month: str, phone: str) -> bool:
    if _drive_ctx is not None:
        folder_id = _drive_ctx.get_folder_id(year, month)
        return _drive_ctx.file_exists(folder_id, year, month, phone)
    pattern = f"{year}{month}_Ymobile_{phone}_*.pdf"
    existing = list(save_dir.glob(pattern))
    if existing:
        log.info(f"  既にダウンロード済み: {existing[0].name}  → スキップ")
        return True
    return False


def _click_any_button(page, locator, label: str = "ボタン", text_hint: str = "") -> bool:
    try:
        locator.first.click(force=True)
        log.info(f"  {label}をクリックしました (force)")
        return True
    except Exception as e:
        log.info(f"  {label} force click 失敗: {e}")

    if text_hint:
        try:
            text_elem = page.get_by_text(text_hint, exact=False).first
            if text_elem.is_visible(timeout=3000):
                text_elem.click(force=True)
                log.info(f"  {label}をクリックしました (テキスト検索: {text_hint})")
                return True
        except Exception:
            pass

    try:
        hint_js = text_hint.replace('"', '\\"') if text_hint else ""
        clicked = page.evaluate(f"""() => {{
            let btn = document.querySelector('input[type="submit"]')
                   || document.querySelector('button[type="submit"]');
            if (btn) {{ btn.click(); return 'submit'; }}
            const hint = "{hint_js}";
            if (hint) {{
                const links = document.querySelectorAll('a');
                for (const a of links) {{
                    if (a.textContent && a.textContent.includes(hint)) {{
                        a.click(); return 'link';
                    }}
                }}
            }}
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
    try:
        page.wait_for_load_state("networkidle", timeout=timeout * 1000)
    except Exception:
        pass
    time.sleep(2)


def _get_page_text(page) -> str:
    try:
        return page.evaluate("() => document.body.innerText || ''")
    except Exception:
        return ""


def _click_send_button(page) -> str | None:
    try:
        return page.evaluate("""() => {
            for (const a of document.querySelectorAll('a')) {
                if ((a.textContent || '').includes('送信する')) {
                    a.click(); return 'a: ' + a.textContent.trim();
                }
            }
            const submit = document.querySelector('input[type="submit"]');
            if (submit && !(submit.value || '').includes('ログイン')) {
                submit.click(); return 'input-submit: ' + submit.value;
            }
            for (const btn of document.querySelectorAll('button[type="submit"]')) {
                if (!(btn.textContent || '').includes('ログイン')) {
                    btn.click(); return 'button: ' + btn.textContent.trim();
                }
            }
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


def _is_on_auth_page(page) -> bool:
    return AUTH_DOMAIN in page.url


def _handle_security_code_flow(page, phone_number: str, password: str = "") -> bool:
    """セキュリティ番号の送信→入力→確認の一連のフローを処理する。"""
    log.info("セキュリティ番号フロー開始...")
    last4 = phone_number[-4:]

    for attempt in range(8):
        _wait_for_page_stable(page)
        current_url = page.url
        log.info(f"  [ステップ{attempt+1}] URL: {current_url}")

        if AUTH_DOMAIN not in current_url:
            log.info("  認証フロー完了！")
            return True

        page_text = _get_page_text(page)

        # セキュリティ番号入力欄
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
                try:
                    result = page.evaluate("""() => {
                        const s = document.querySelector('input[type="submit"]');
                        if (s && !(s.value || '').includes('ログイン')) {
                            s.click(); return 'submit: ' + s.value;
                        }
                        for (const btn of document.querySelectorAll('button[type="submit"]')) {
                            if (!(btn.textContent || '').includes('ログイン')) {
                                btn.click(); return 'button: ' + btn.textContent.trim();
                            }
                        }
                        const f = document.querySelector('form');
                        if (f) { f.submit(); return 'form-submit'; }
                        return null;
                    }""")
                    log.info(f"  本人確認ボタン: {result}")
                except Exception as e:
                    log.warning(f"  本人確認ボタンクリック失敗: {e}")
                _wait_for_page_stable(page, timeout=15)
                log.info(f"  本人確認後URL: {page.url}")
                if AUTH_DOMAIN not in page.url:
                    log.info("  本人確認完了！認証フロー完了！")
                    return True
                continue
        except Exception:
            pass

        # ページ情報ログ
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

        # ラジオボタン（SMS送付先選択）
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

        if ("セキュリティ番号" in page_text or "送付先" in page_text
                or "連絡先" in page_text or "本人確認" in page_text):
            result = _click_send_button(page)
            if result:
                log.info(f"  送信ボタンクリック: {result}")
                _wait_for_page_stable(page, timeout=12)
                continue
            else:
                log.warning("  送信ボタンが見つかりませんでした")

        elif AUTH_DOMAIN in current_url:
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

        if attempt >= 7:
            break

    if AUTH_DOMAIN in page.url:
        log.error("  まだ認証ページにいます")
        return False

    log.info("  認証フロー完了！")
    return True


def do_login_and_navigate(page, phone_number: str, password: str) -> bool:
    """ログイン → 2FA → PDFダウンロードページまで遷移する。"""
    # Step 1: PDFページに直接アクセス（認証ページにリダイレクトされる）
    log.info("PDFページに直接アクセス中...")
    page.goto(BILL_PDF_URL, wait_until="networkidle")
    page.wait_for_load_state("networkidle")
    time.sleep(2)
    log.info(f"  リダイレクト先URL: {page.url}")

    if not _is_on_auth_page(page):
        log.info("  認証なしでPDFページに到達しました")
        return True

    page_text = _get_page_text(page)

    if "セキュリティ番号" in page_text or "送付先" in page_text:
        log.info("  セキュリティ番号ページを検出（セッション再利用）")
        if not _handle_security_code_flow(page, phone_number, password):
            return False
    else:
        log.info("ログイン情報を入力中...")
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
            try:
                page.wait_for_function(
                    "() => !document.querySelector('input[name=\"telnum\"]')",
                    timeout=15000,
                )
                log.info("  ログインフォームが消えました")
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

    # Step 3: ログイン後のエラー確認
    if _is_on_auth_page(page):
        error_msg = page.locator(".err-area, .error, .alert-error, .sbid-error")
        if error_msg.count() > 0:
            try:
                err_text = error_msg.first.text_content()
                log.error(f"  ログインエラー: {err_text}")
            except Exception:
                pass
        log.info("  認証ページにいます → セキュリティ番号フローを処理します")
        if not _handle_security_code_flow(page, phone_number, password):
            return False

    log.info(f"  認証後のURL: {page.url}")

    # Step 4: PDFページへの到達確認・ナビゲーション
    if not _is_on_auth_page(page):
        # combobox（月選択）があればPDFページに到達
        try:
            combobox = page.get_by_role("combobox").first
            if combobox.is_visible(timeout=3000):
                log.info("  PDFダウンロードページに到達しました！")
                return True
        except Exception:
            pass

        # ポータルにいる場合: 書面発行ページへ誘導
        log.info("  PDFページへの遷移を試みます...")

        # 書面発行ディープリンクを試す
        try:
            log.info(f"  書面発行ディープリンクへアクセス: {LOGIN_URL}")
            page.goto(LOGIN_URL, wait_until="networkidle")
            time.sleep(2)
            log.info(f"  ディープリンク後のURL: {page.url}")
        except Exception:
            pass

        # 2回目の認証が必要な場合
        if _is_on_auth_page(page):
            log.info("  2回目の認証が必要です")
            try:
                phone_field = page.locator('input[name="telnum"]')
                if phone_field.first.is_visible(timeout=3000):
                    phone_field.first.fill(phone_number)
                    pw_field = page.locator('input[type="password"]')
                    pw_field.first.fill(password)
                    login_btn2 = page.locator('input[type="submit"]').or_(page.get_by_text("ログインする", exact=False))
                    _click_any_button(page, login_btn2, "ログインボタン(2回目)", text_hint="ログイン")
                    page.wait_for_load_state("networkidle")
                    time.sleep(3)
            except Exception:
                pass
            if _is_on_auth_page(page):
                if not _handle_security_code_flow(page, phone_number, password):
                    return False

        # WCOシステム上のPDFページへ遷移
        _navigate_to_pdf_page(page)

    # Step 5: 2回目の認証（WCOアクセス時）
    if _is_on_auth_page(page):
        log.info("  WCOアクセスで2回目の認証が必要です")
        if not _handle_security_code_flow(page, phone_number, password):
            return False

    final_url = page.url
    log.info(f"  最終URL: {final_url}")

    if _is_on_auth_page(page):
        log.error("  認証を完了できませんでした")
        return False

    # 最終確認
    try:
        combobox = page.get_by_role("combobox").first
        if combobox.is_visible(timeout=5000):
            log.info("  PDFダウンロードページに到達しました！")
            return True
    except Exception:
        pass

    try:
        pdf_link = page.locator('a[href*="doPrint"]')
        if pdf_link.count() > 0:
            log.info("  PDFダウンロードページに到達しました（PDFリンク検出）")
            return True
    except Exception:
        pass

    if "bl61.my.ymobile.jp/wco" in final_url:
        log.info("  WCOドメインにいます。再遷移を試みます...")
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


def _navigate_to_pdf_page(page) -> None:
    """WCOシステム上のPDFダウンロードページへ遷移する"""
    # 書面発行ページ（WCO250 または同等のY!mobile向けページ）
    try:
        cert_link = page.locator('a[href*="/wco/certificate/"]')
        if cert_link.count() > 0:
            cert_link.first.click()
        else:
            page.goto(CERTIFICATE_URL, wait_until="networkidle")
        page.wait_for_load_state("networkidle")
        time.sleep(2)
        log.info(f"  書面発行ページ後のURL: {page.url}")
    except Exception:
        pass

    # PDFダウンロードページ（goBillInfoPdf）
    try:
        pdf_page_link = page.locator('a[href*="goBillInfoPdf"]')
        if pdf_page_link.count() > 0:
            pdf_page_link.first.click()
        else:
            print_link = page.get_by_text(re.compile(r"自分で印刷|PDFファイル"))
            if print_link.count() > 0:
                print_link.first.click()
            else:
                page.goto(BILL_PDF_URL, wait_until="networkidle")
        page.wait_for_load_state("networkidle")
        time.sleep(2)
        log.info(f"  PDF印刷ページ後のURL: {page.url}")
    except Exception:
        pass


def select_target_month(page, year: str, month: str) -> bool:
    target_value = f"{year}{month}"
    target_label = f"{year}年{month}月"
    log.info(f"対象月を選択中: {target_label}")

    try:
        combobox = page.get_by_role("combobox").first
        combobox.wait_for(state="visible", timeout=10000)
        try:
            combobox.select_option(value=target_value)
            log.info(f"  月を選択しました: {target_value}")
        except Exception:
            combobox.select_option(label=re.compile(target_label))
            log.info(f"  月をラベルで選択しました: {target_label}")
        page.wait_for_load_state("networkidle")
        time.sleep(2)
        return True
    except Exception as e:
        log.error(f"対象月の選択に失敗しました: {e}")
        return False


def _download_single_pdf(page, link, label: str, save_dir: Path,
                          year: str, month: str, phone: str, amount: str) -> bool:
    try:
        link_text = link.text_content() or ""
        href = link.get_attribute("href") or ""
        log.info(f"  [{label}] PDFリンク: {link_text.strip()} ({href})")
        with page.expect_download(timeout=60000) as download_info:
            link.click()
        download = download_info.value
        filename = build_filename(year, month, phone, amount)
        dest = save_dir / filename
        if dest.exists():
            stem = dest.stem
            dest = save_dir / f"{stem}_{label}.pdf"
        download.save_as(str(dest))

        if _drive_ctx is not None:
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
    if pdf_types is None:
        pdf_types = {"電話番号別"}

    log.info(f"ダウンロード対象: {', '.join(sorted(pdf_types))}")

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

    if "一括" in pdf_types:
        log.info("一括印刷用PDFを確認中...")
        bulk_link = page.locator('a[href*="doPrintSbmAll"]')
        if bulk_link.count() > 0 and bulk_link.first.is_visible():
            if _download_single_pdf(page, bulk_link.first, "一括", save_dir, year, month, phone, amount):
                any_success = True
        else:
            log.info("  一括印刷用PDFリンクが見つかりませんでした")

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
    if pdf_types is None:
        pdf_types = {"電話番号別"}
    log.info(f"=== {phone_number} の処理を開始 (PDFの種類: {', '.join(sorted(pdf_types))}) ===")

    if check_already_downloaded(save_dir, year, month, phone_number):
        return True

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
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
            if not do_login_and_navigate(page, phone_number, password):
                _save_debug_screenshot(
                    page, save_dir, phone_number, year, month,
                    "ログインまたはページ遷移失敗",
                    f"PDFダウンロードページに到達できませんでした。最終URL: {page.url}",
                )
                return False

            try:
                context.storage_state(path=str(session_file))
                log.info(f"  セッションを保存しました: {session_file}")
            except Exception as e:
                log.warning(f"  セッション保存に失敗: {e}")

            if not select_target_month(page, year, month):
                _save_debug_screenshot(
                    page, save_dir, phone_number, year, month,
                    "月選択失敗", f"対象月 {year}{month} の選択に失敗しました",
                )
                return False

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
    log.info("My Y!mobile 料金明細PDFダウンロードを開始します")

    year, month = get_target_month()
    if not TARGET_MONTH:
        try:
            gc = get_gspread_client()
            sh = _open_sheet(gc)
            ws = sh.worksheet("設定")
            for row in ws.get_all_records():
                if str(row.get("設定名", "")).strip() == "対象月":
                    val = str(row.get("値", "")).strip()
                    if re.match(r"^\d{6}$", val):
                        year, month = val[:4], val[4:6]
                        log.info(f"対象月（設定シートから取得）: {year}年{month}月")
                    elif m2 := re.match(r"^(\d{4})年(\d+)月$", val):
                        year, month = m2.group(1), m2.group(2).zfill(2)
                        log.info(f"対象月（設定シートから取得）: {year}年{month}月")
                    break
        except RecursionError:
            raise
        except Exception as e:
            log.warning(f"設定シートから対象月の取得に失敗: {type(e).__name__}: {e}")
    log.info(f"対象月: {year}年{month}月")

    base_path = resolve_save_path()
    save_dir = ensure_save_dir(base_path, year, month)
    mode_label = "Drive APIモード" if _drive_ctx else "ローカルモード"
    log.info(f"保存先: {save_dir} ({mode_label})")

    accounts = load_accounts()

    results = []
    for _, row in accounts.iterrows():
        phone = str(row["電話番号"]).strip()
        pw = str(row["パスワード"]).strip()
        pdf_types = parse_pdf_types(row.get("PDFの種類", "電話番号別"))
        ok = download_billing_pdf(phone, pw, year, month, save_dir, pdf_types)
        results.append((phone, ok))

    log.info("=" * 50)
    log.info("処理結果サマリー:")
    for phone, ok in results:
        status = "✅ 成功" if ok else "❌ 失敗"
        log.info(f"  {phone}: {status}")

    succeeded = sum(1 for _, ok in results if ok)
    log.info(f"  合計: {succeeded}/{len(results)} 件成功")
    log.info("=" * 50)

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
