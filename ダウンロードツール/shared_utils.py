"""
携帯領収書管理 共通ユーティリティ

SoftBank・Ymobile・au・UQmobile から共通で使用する関数群。
CarrierConfig / BillingContext を中心にすべてのロジックを集約し、
各キャリアスクリプトは薄いラッパーとして動作する。

carrier_family による分岐:
  - "softbank": SoftBank / Ymobile（WCOシステム経由のPDFダウンロード）
  - "au": au / UQmobile（WEB de 請求書経由のPDFダウンロード）
"""

import json
import os
import random
import re
import shutil
import subprocess
import sys
import logging
import tempfile
import time
import webbrowser
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

log = logging.getLogger(__name__)

# ── User-Agent定数 ──

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)


# ══════════════════════════════════════════════════════════
#  CarrierConfig / BillingContext
# ══════════════════════════════════════════════════════════

@dataclass(frozen=True)
class CarrierConfig:
    """キャリア固有の定数群"""
    carrier_name: str           # "SoftBank", "Ymobile", "au", "UQmobile"
    display_name: str           # "My SoftBank", "My Y!mobile", "My au", "My UQ mobile"
    auth_domain: str            # "id.my.softbank.jp", "id.my.ymobile.jp", "connect.auone.jp"
    login_url: str
    code_file_prefix: str       # "softbank", "ymobile", "au", "uqmobile"
    session_file_prefix: str
    temp_dir_prefix: str        # "softbank_pdf_", "ymobile_pdf_", "au_pdf_", "uqmobile_pdf_"
    carrier_family: str = "softbank"  # "softbank" or "au" — ログイン・PDF取得フローの分岐に使用
    # SoftBank/Ymobile 固有（WCOシステム）
    wco_base: str = ""
    certificate_url: str = ""
    bill_pdf_url: str = ""
    # au/UQ 固有
    au_billing_top_url: str = ""      # WEB de 請求書トップページ
    au_pin_setting_name: str = ""     # 設定シートの暗証番号キー名
    password_setting_name: str = "パスワード"  # 設定シートのパスワードキー名


@dataclass
class BillingContext:
    """実行時の状態をすべて保持するコンテキスト"""
    config: CarrierConfig
    script_dir: Path
    spreadsheet_id: str
    base_save_path: str | None
    target_month: str | None
    headless: bool
    dry_run: bool
    retry_phones: list[str]
    security_code_timeout: int
    drive_ctx: "DriveContext | None" = None
    temp_save_dir: Path | None = None
    phone_device_map: dict[str, str] = field(default_factory=dict)
    phone_login_id_map: dict[str, str] = field(default_factory=dict)  # 電話番号 → au ID


# ══════════════════════════════════════════════════════════
#  リトライユーティリティ
# ══════════════════════════════════════════════════════════

def retry_with_backoff(fn, max_retries=3, base_delay=2.0, max_delay=30.0,
                       retryable_exceptions=(Exception,), logger=None):
    """指数バックオフ付きリトライ。fn() を最大 max_retries 回実行する。"""
    last_exc = None
    for attempt in range(max_retries):
        try:
            return fn()
        except retryable_exceptions as e:
            last_exc = e
            if attempt < max_retries - 1:
                delay = min(base_delay * (2 ** attempt) + random.uniform(0, 1), max_delay)
                if logger:
                    logger.warning(f"  リトライ {attempt+1}/{max_retries}: {e} → {delay:.1f}秒後に再試行")
                time.sleep(delay)
    raise last_exc


# ══════════════════════════════════════════════════════════
#  Driveフォールバックパス
# ══════════════════════════════════════════════════════════

def get_drive_fallback_path() -> Path:
    """Drive容量超過時のローカル保存先を返す。環境変数 DRIVE_FALLBACK_PATH で上書き可能。"""
    env = os.environ.get("DRIVE_FALLBACK_PATH", "").strip()
    if env:
        return Path(env)
    return Path(__file__).resolve().parent.parent


# ══════════════════════════════════════════════════════════
#  電話番号処理
# ══════════════════════════════════════════════════════════

def strip_hyphens(phone: str) -> str:
    """電話番号からハイフンを除去し、先頭0が消えていたら補完する。"""
    cleaned = re.sub(r"[-\s\u2010-\u2015\u2212\uFF0D]", "", phone)
    if cleaned and cleaned[0] != "0" and len(cleaned) == 10:
        cleaned = "0" + cleaned
    return cleaned


# ══════════════════════════════════════════════════════════
#  PDFの種類パース
# ══════════════════════════════════════════════════════════

def parse_pdf_types(raw: str, carrier_family: str = "softbank") -> set[str]:
    """PDFの種類列の値をパースしてセットで返す。"""
    if carrier_family == "au":
        valid = {"請求書", "領収書", "支払証明書"}
        default = {"請求書"}
    else:
        valid = {"電話番号別", "一括", "機種別"}
        default = {"電話番号別"}
    if not raw or str(raw).strip().lower() in ("nan", ""):
        return default
    types = {t.strip() for t in str(raw).split(",") if t.strip()}
    result = types & valid
    if not result:
        log.warning(f"  PDFの種類 '{raw}' に有効な値がありません。デフォルト '{', '.join(default)}' を使用します。")
        return default
    return result


# ══════════════════════════════════════════════════════════
#  サービスアカウント・認証
# ══════════════════════════════════════════════════════════

def find_service_account_json() -> Path:
    """service_account.json を探す（ダウンロードツール/）"""
    tool_dir = Path(__file__).resolve().parent
    p = tool_dir / "service_account.json"
    if p.exists():
        return p
    raise FileNotFoundError(
        "service_account.json が見つかりません。\n"
        "ダウンロードツール/ フォルダに配置してください。"
    )


def find_client_secrets() -> "Path | None":
    """client_secrets.json を探す（ダウンロードツール/）"""
    p = Path(__file__).resolve().parent / "client_secrets.json"
    return p if p.exists() else None


def guide_spreadsheet_sharing(spreadsheet_id: str) -> None:
    """スプレッドシートへの権限エラー時にサービスアカウントのメールをクリップボードにコピー＋ブラウザを開く"""
    try:
        sa_email = json.loads(find_service_account_json().read_text())["client_email"]
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
    json_path = find_service_account_json()
    creds = Credentials.from_service_account_file(
        str(json_path),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return gspread.authorize(creds)


def open_sheet(gc: gspread.Client, spreadsheet_id: str) -> gspread.Spreadsheet:
    """スプレッドシートを開く。権限エラー時は共有案内を表示してリトライする。"""
    try:
        return gc.open_by_key(spreadsheet_id)
    except (PermissionError, gspread.exceptions.APIError) as e:
        is_403 = isinstance(e, PermissionError) or (
            hasattr(e, "response") and e.response.status_code == 403
        )
        if is_403:
            guide_spreadsheet_sharing(spreadsheet_id)
            return gc.open_by_key(spreadsheet_id)
        raise


# ══════════════════════════════════════════════════════════
#  設定シート読み取り
# ══════════════════════════════════════════════════════════

def load_password_from_settings(sh, setting_name: str = "パスワード") -> "str | None":
    """設定シートからパスワードを取得する。setting_name で取得するキーを指定可能。"""
    try:
        ws = sh.worksheet("設定")
        for row in ws.get_all_records():
            if str(row.get("設定名", "")).strip() == setting_name:
                pw = str(row.get("値", "")).strip()
                if pw:
                    return pw
    except Exception:
        pass
    return None


# ══════════════════════════════════════════════════════════
#  ダウンロード履歴
# ══════════════════════════════════════════════════════════

def write_download_history(
    spreadsheet_id: str,
    carrier_name: str,
    results: list[tuple[str, bool]],
    year: str,
    month: str,
    save_dir: "str | Path | None" = None,
    downloaded_filenames: "dict[str, list[str]] | None" = None,
) -> None:
    """ダウンロード結果をスプレッドシートの「ダウンロード履歴」シートに記録する。"""
    try:
        gc = get_gspread_client()
        sh = gc.open_by_key(spreadsheet_id)
        try:
            ws = sh.worksheet("ダウンロード履歴")
        except gspread.exceptions.WorksheetNotFound:
            log.debug("ダウンロード履歴シートが見つかりません（スキップ）")
            return

        # ファイル名マップが渡されなかった場合、ローカルディレクトリから検索（フォールバック）
        if downloaded_filenames is None:
            downloaded_filenames = {}
            if save_dir:
                try:
                    sd = Path(str(save_dir))
                    if sd.exists():
                        for f in sd.glob("*.pdf"):
                            m = re.match(rf"^\d{{6}}_{re.escape(carrier_name)}_(\d+)", f.name)
                            if m:
                                phone = m.group(1)
                                if phone not in downloaded_filenames:
                                    downloaded_filenames[phone] = []
                                downloaded_filenames[phone].append(f.name)
                except Exception:
                    pass

        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        rows = []
        for phone, ok in results:
            filenames = ", ".join(downloaded_filenames.get(phone, []))
            rows.append([now, carrier_name, phone, f"{year}{month}", filenames, "成功" if ok else "失敗"])
        if rows:
            ws.append_rows(rows, value_input_option="USER_ENTERED")
            log.info(f"  ダウンロード履歴を記録しました（{len(rows)}件）")
    except Exception as e:
        log.warning(f"ダウンロード履歴の記録に失敗（処理は続行）: {e}")


# ══════════════════════════════════════════════════════════
#  Google Drive API
# ══════════════════════════════════════════════════════════

_DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]


@dataclass
class DriveContext:
    """Google Drive API を使ったフォルダ作成・存在確認・アップロードを担うクラス"""
    base_folder_id: str
    carrier_name: str
    drive_service_factory: Any  # callable returning drive service
    service: Any = field(init=False)
    _folder_cache: dict = field(default_factory=dict)

    def __post_init__(self):
        self.service = self.drive_service_factory()

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
        """base_folder_id/YYYY/MM/{carrier_name} のフォルダIDを返す（なければ作成）"""
        year_id = self._get_or_create_folder(self.base_folder_id, year)
        month_id = self._get_or_create_folder(year_id, month)
        return self._get_or_create_folder(month_id, self.carrier_name)

    def file_exists(self, folder_id: str, year: str, month: str, phone: str) -> bool:
        """指定フォルダ内に対象ファイルが既に存在するか確認する"""
        prefix = f"{year}{month}_{self.carrier_name}_{phone}_"
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
            carrier = local_path.parent.name
            month   = local_path.parent.parent.name
            year    = local_path.parent.parent.parent.name
            dest_dir = get_drive_fallback_path() / year / month / carrier
            dest_dir.mkdir(parents=True, exist_ok=True)
            dest = dest_dir / local_path.name
            shutil.copy2(local_path, dest)
            log.info(f"  ローカル保存（Google Drive同期）: {dest}")
            return True
        except Exception as e:
            log.error(f"  ローカル保存フォールバックも失敗: {e}")
            return False


def _bootstrap_client_secrets(script_dir: Path) -> None:
    """
    client_secrets.json が存在しない場合、サービスアカウントでアクセス可能な
    Driveから自動ダウンロードを試みる。
    """
    if find_client_secrets():
        return
    try:
        json_path = find_service_account_json()
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


def _get_drive_service(script_dir: Path):
    """Drive APIサービスを返す。client_secrets.jsonがあればOAuth、なければサービスアカウント。"""
    _bootstrap_client_secrets(script_dir)
    secrets_path = find_client_secrets()
    if secrets_path:
        return _get_drive_service_oauth(secrets_path)
    json_path = find_service_account_json()
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


# ══════════════════════════════════════════════════════════
#  環境初期化
# ══════════════════════════════════════════════════════════

def bootstrap_env_from_gsheet(script_dir: Path, display_name: str) -> None:
    """
    .env が存在しない場合、同ディレクトリの *.gsheet から SPREADSHEET_URL を自動生成する。
    """
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
        f"# ─── {display_name} 料金明細ダウンロード 環境変数 ───\n"
        "# .gsheet ファイルから自動生成されました\n\n"
        f"SPREADSHEET_URL={url}\n\n"
        "# BASE_SAVE_PATH=\n"
        "# TARGET_MONTH=202602\n"
        "HEADLESS=true\n"
        "SECURITY_CODE_TIMEOUT=60\n",
        encoding="utf-8",
    )
    log.info(f".env を自動生成しました: {env_path}")


def _resolve_spreadsheet_id() -> str:
    url = os.environ.get("SPREADSHEET_URL", "").strip()
    if url:
        m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
        if m:
            return m.group(1)
    sid = os.environ.get("SPREADSHEET_ID", "").strip()
    if sid:
        return sid
    log.error(
        "スプレッドシートが設定されていません。.env に以下のいずれかを設定してください:\n"
        "  SPREADSHEET_URL=https://docs.google.com/spreadsheets/d/xxxxx/edit\n"
        "  SPREADSHEET_ID=xxxxx"
    )
    sys.exit(1)


def create_billing_context(config: CarrierConfig, script_dir: Path) -> BillingContext:
    """CarrierConfig と script_dir から BillingContext を初期化して返す。"""
    bootstrap_env_from_gsheet(script_dir, config.display_name)
    load_dotenv(script_dir / ".env")

    spreadsheet_id = _resolve_spreadsheet_id()
    base_save_path = os.environ.get("BASE_SAVE_PATH")
    target_month = os.environ.get("TARGET_MONTH")
    headless = os.environ.get("HEADLESS", "true").lower() in ("true", "1", "yes")
    dry_run = os.environ.get("DRY_RUN", "false").lower() == "true"
    retry_phones = [p.strip() for p in os.environ.get("RETRY_PHONES", "").split(",") if p.strip()]
    security_code_timeout = int(os.environ.get("SECURITY_CODE_TIMEOUT", "60"))

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )

    return BillingContext(
        config=config,
        script_dir=script_dir,
        spreadsheet_id=spreadsheet_id,
        base_save_path=base_save_path,
        target_month=target_month,
        headless=headless,
        dry_run=dry_run,
        retry_phones=retry_phones,
        security_code_timeout=security_code_timeout,
    )


# ══════════════════════════════════════════════════════════
#  ユーティリティ関数
# ══════════════════════════════════════════════════════════

def get_target_month(ctx: BillingContext) -> tuple[str, str]:
    """対象の年月を (year, month) 文字列で返す"""
    if ctx.target_month:
        ym = ctx.target_month.strip()
        return ym[:4], ym[4:6]
    today = datetime.today()
    first = today.replace(day=1)
    prev = first - timedelta(days=1)
    return str(prev.year), f"{prev.month:02d}"


def resolve_save_path(ctx: BillingContext) -> str:
    """PDF保存先パスを取得する。

    Google Drive URLの場合はDrive APIモードを初期化し "drive://{folder_id}" を返す。
    ローカルパスの場合はそのまま返す。
    """
    # ── 1. 設定シートから取得 ──
    try:
        gc = get_gspread_client()
        sh = open_sheet(gc, ctx.spreadsheet_id)
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
                                ctx.drive_ctx = DriveContext(
                                    base_folder_id=folder_id,
                                    carrier_name=ctx.config.carrier_name,
                                    drive_service_factory=lambda: _get_drive_service(ctx.script_dir),
                                )
                                ctx.temp_save_dir = Path(tempfile.mkdtemp(prefix=ctx.config.temp_dir_prefix))
                                log.info(f"一時保存ディレクトリ: {ctx.temp_save_dir}")
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
    except RecursionError:
        raise
    except Exception as e:
        log.error(f"設定シートからの保存先取得に失敗しました: {e}")
        sys.exit(1)

    # ── 2. 環境変数 ──
    if ctx.base_save_path:
        return ctx.base_save_path

    log.error(
        "PDF保存先が設定されていません。以下のいずれかを設定してください:\n"
        "  1. スプレッドシートの「設定」シートの「PDF保存先フォルダ」に記入\n"
        "     - Google DriveフォルダのURL（推奨）: https://drive.google.com/drive/folders/xxxxx\n"
        "       ※ service_account.json のメールアドレスをフォルダの共有設定（編集者）に追加してください\n"
        "     - Macのローカル絶対パス: /Users/.../マイドライブ/確定申告系/携帯領収書管理\n"
        "  2. 環境変数 BASE_SAVE_PATH を設定"
    )
    sys.exit(1)


def load_accounts(ctx: BillingContext) -> pd.DataFrame:
    """スプレッドシートから回線情報を読み込む。"""
    log.info("スプレッドシートから回線情報を読み込み中...")
    gc = get_gspread_client()
    sh = open_sheet(gc, ctx.spreadsheet_id)

    common_password = load_password_from_settings(sh, ctx.config.password_setting_name)
    if not common_password:
        log.error(
            f"パスワードが設定されていません。\n"
            f"  設定シートの「{ctx.config.password_setting_name}」行にログインパスワードを入力してください。"
        )
        sys.exit(1)

    df: "pd.DataFrame | None" = None
    df_all: "pd.DataFrame | None" = None
    try:
        ws = sh.worksheet("認証情報")
        records = ws.get_all_records()
        df_all = pd.DataFrame(records)
        df_all.columns = df_all.columns.str.strip()
        if "電話番号" in df_all.columns and "キャリア" in df_all.columns:
            df = df_all[df_all["キャリア"].str.strip() == ctx.config.carrier_name].reset_index(drop=True)
            log.info(f"  「認証情報」シートから {ctx.config.carrier_name} 回線を読み込み")
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

    df["パスワード"] = common_password
    df["電話番号"] = df["電話番号"].astype(str).apply(strip_hyphens)
    if "PDFの種類" not in df.columns:
        df["PDFの種類"] = "電話番号別"
    if "状態" in df.columns:
        cancelled_count = (df["状態"].astype(str).str.strip() == "解約済").sum()
        if cancelled_count > 0:
            log.info(f"  うち解約済 {cancelled_count} 件を含む")
    log.info(f"  {len(df)} 件の回線を読み込みました")

    ctx.phone_device_map = {}
    ctx.phone_login_id_map = {}
    if df_all is not None:
        for _, row in df_all.iterrows():
            phone = strip_hyphens(str(row.get("電話番号", "")))
            if not re.match(r'^\d{10,13}$', phone):
                continue
            device = str(row.get("運用端末", "")).strip()
            if device:
                ctx.phone_device_map[phone] = device
            login_id = str(row.get("ログインID", "") or row.get("au ID", "")).strip()
            if login_id:
                ctx.phone_login_id_map[phone] = login_id
        if ctx.phone_device_map:
            log.info(f"  運用端末マップ: {len(ctx.phone_device_map)} 件")
        if ctx.phone_login_id_map:
            log.info(f"  ログインIDマップ: {len(ctx.phone_login_id_map)} 件")

    return df


# ══════════════════════════════════════════════════════════
#  セッション・セキュリティコード
# ══════════════════════════════════════════════════════════

_TMPDIR = Path(tempfile.gettempdir())


def _code_file(ctx: BillingContext) -> Path:
    return _TMPDIR / f"{ctx.config.code_file_prefix}_security_code.txt"


def _session_file(ctx: BillingContext, phone_number: str) -> Path:
    return _TMPDIR / f"{ctx.config.session_file_prefix}_session_{phone_number}.json"


def ask_security_code(ctx: BillingContext, phone_number: str) -> str | None:
    """セキュリティ番号を取得する。"""
    code_f = _code_file(ctx)
    if code_f.exists():
        code_f.unlink()

    device = ctx.phone_device_map.get(phone_number, "")
    print("\n" + "=" * 60)
    print(f"  \U0001f4f1 SMS認証が必要です")
    print(f"  電話番号: {phone_number}")
    if device:
        print(f"  端末    : {device}")
    print(f"  SMSに届いた3桁のセキュリティ番号を入力してください")
    print(f"  ターミナル入力 または 以下のコマンドで渡してください:")
    print(f"    echo '123' > {code_f}")
    print("=" * 60)

    if sys.stdin.isatty():
        try:
            code = input("  セキュリティ番号（3桁）: ").strip()
            if code:
                return code
        except (EOFError, KeyboardInterrupt):
            pass
    else:
        log.info(f"  非インタラクティブ環境を検出。ファイルの出現を待機中...")
        log.info(f"  別ターミナルで: echo '123' > {code_f}")

    timeout = ctx.security_code_timeout
    deadline = time.time() + timeout
    last_log = time.time()
    while time.time() < deadline:
        if code_f.exists():
            try:
                code = code_f.read_text(encoding="utf-8").strip()
                code_f.unlink()
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

    log.error(
        f"  セキュリティ番号の入力がタイムアウトしました（{timeout}秒）\n"
        f"  考えられる原因:\n"
        f"    1. SMSが届いていない → 電話番号が正しいか確認してください\n"
        f"    2. 入力が間に合わなかった → SECURITY_CODE_TIMEOUT の値を増やしてください（現在: {timeout}秒）\n"
        f"    3. ファイル経由の場合 → echo '123' > {code_f}"
    )
    return None


# ══════════════════════════════════════════════════════════
#  ファイル名・保存関連
# ══════════════════════════════════════════════════════════

def sanitize_amount(text: str) -> str:
    """請求金額テキストから数字だけ抽出して「○○円」形式にする"""
    digits = re.sub(r"[^\d]", "", text)
    if digits:
        return f"{int(digits)}円"
    return ""


def build_filename(ctx: BillingContext, year: str, month: str, phone: str, amount: str = "") -> str:
    """電子帳簿保存法準拠のファイル名を生成する"""
    base = f"{year}{month}_{ctx.config.carrier_name}_{phone}"
    if amount:
        base += f"_{amount}"
    else:
        base += "_利用料金明細"
    return base + ".pdf"


def _save_debug_screenshot(
    ctx: BillingContext, page, save_dir: Path, phone: str, year: str, month: str,
    error_type: str, error_detail: str,
) -> None:
    """デバッグ用スクリーンショットをローカルに保存する。"""
    try:
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        if ctx.drive_ctx is not None:
            base_path = ctx.script_dir
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


def _cleanup_old_debug_screenshots(ctx: BillingContext, max_age_days: int = 7) -> None:
    """7日以上前のデバッグスクリーンショットを自動削除する"""
    try:
        debug_base = ctx.script_dir / "debug"
        if not debug_base.exists():
            return
        cutoff = datetime.now() - timedelta(days=max_age_days)
        for d in debug_base.iterdir():
            if d.is_dir() and d.name.startswith("debug_"):
                try:
                    ts = datetime.strptime(d.name[6:21], "%Y%m%d_%H%M%S")
                    if ts < cutoff:
                        shutil.rmtree(str(d))
                        log.info(f"  古いデバッグフォルダを削除: {d.name}")
                except (ValueError, OSError):
                    pass
    except Exception:
        pass


def ensure_save_dir(ctx: BillingContext, base_path: str, year: str, month: str) -> Path:
    """保存先ディレクトリを作成して返す (base_path/year/month/キャリア名)"""
    if base_path.startswith("drive://") and ctx.temp_save_dir is not None:
        save_dir = ctx.temp_save_dir / year / month / ctx.config.carrier_name
    else:
        save_dir = Path(base_path) / year / month / ctx.config.carrier_name
    save_dir.mkdir(parents=True, exist_ok=True)
    return save_dir


# ══════════════════════════════════════════════════════════
#  ダウンロード判定
# ══════════════════════════════════════════════════════════

def check_already_downloaded(ctx: BillingContext, save_dir: Path, year: str, month: str, phone: str,
                             pdf_types: set[str] | None = None) -> tuple[bool, set[str]]:
    """指定の電話番号・月のPDFが既にダウンロード済みかチェックする。
    返却: (all_downloaded, remaining_types) - 全種類DL済みならTrue、未DLの種類セット
    """
    if pdf_types is None:
        pdf_types = {"電話番号別"}

    # au/UQ: 種類ごとにサフィックスが異なる
    type_suffixes = {"請求書": "", "領収書": "_領収書", "支払証明書": "_支払証明書",
                     "電話番号別": "", "一括": "_一括", "機種別": "_機種別"}

    remaining = set()
    prefix = f"{year}{month}_{ctx.config.carrier_name}_{phone}_"

    for pt in pdf_types:
        suffix = type_suffixes.get(pt, "")
        if ctx.drive_ctx is not None:
            folder_id = ctx.drive_ctx.get_folder_id(year, month)
            # サフィックス付きで検索
            if suffix:
                q = (f"name contains '{prefix}' and name contains '{suffix}.pdf' "
                     f"and '{folder_id}' in parents and mimeType='application/pdf' and trashed=false")
            else:
                # 基本ファイル（サフィックスなし）: _領収書/_支払証明書/_一括/_機種別 を除外
                q = (f"name contains '{prefix}' and '{folder_id}' in parents "
                     f"and mimeType='application/pdf' and trashed=false")
                # 除外サフィックスを持つファイルは基本ファイルではない
            res = ctx.drive_ctx.service.files().list(q=q, fields="files(name)").execute()
            files = res.get("files", [])
            if suffix:
                found = any(f["name"].endswith(f"{suffix}.pdf") for f in files)
            else:
                exclude = {"_領収書.pdf", "_支払証明書.pdf", "_一括.pdf", "_機種別.pdf"}
                found = any(not any(f["name"].endswith(ex) for ex in exclude) for f in files)
            if found:
                log.info(f"  既にDriveにアップロード済み: {pt}  → スキップ")
            else:
                remaining.add(pt)
        else:
            if suffix:
                pattern = f"{prefix}*{suffix}.pdf"
            else:
                pattern = f"{prefix}*.pdf"
            existing = list(save_dir.glob(pattern))
            if suffix:
                found = len(existing) > 0
            else:
                exclude = {"_領収書.pdf", "_支払証明書.pdf", "_一括.pdf", "_機種別.pdf"}
                found = any(not any(str(f).endswith(ex) for ex in exclude) for f in existing)
            if found:
                log.info(f"  既にダウンロード済み: {pt}  → スキップ")
            else:
                remaining.add(pt)

    return (len(remaining) == 0, remaining)


# ══════════════════════════════════════════════════════════
#  ブラウザ操作ヘルパー
# ══════════════════════════════════════════════════════════

def _click_any_button(page, locator, label: str = "ボタン", text_hint: str = "") -> bool:
    """ボタンをクリックする。force=True → テキスト検索 → JS click の順に試す。"""
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

    # 3) JSフォールバック: 幅広い要素を探す (SoftBank版の全要素検索を含む)
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
    """「送信する」ボタンをJSでクリックする。"""
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


def _is_on_auth_page(ctx: BillingContext, page) -> bool:
    """現在のページが認証ページかどうか判定する"""
    return ctx.config.auth_domain in page.url


# ══════════════════════════════════════════════════════════
#  セキュリティ番号フロー
# ══════════════════════════════════════════════════════════

def _handle_security_code_flow(ctx: BillingContext, page, phone_number: str, password: str = "") -> bool:
    """セキュリティ番号の送信→入力→確認の一連のフローを処理する。"""
    log.info("セキュリティ番号フロー開始...")
    last4 = phone_number[-4:]
    auth_domain = ctx.config.auth_domain

    for attempt in range(8):
        _wait_for_page_stable(page)
        current_url = page.url
        log.info(f"  [ステップ{attempt+1}] URL: {current_url}")

        # 認証ページを離れていれば完了
        if auth_domain not in current_url:
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
                code = ask_security_code(ctx, phone_number)
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
                if auth_domain not in page.url:
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

        # 「送信する」ボタンをクリック
        if ("セキュリティ番号" in page_text or "送付先" in page_text
                or "連絡先" in page_text or "本人確認" in page_text):
            result = _click_send_button(page)
            if result:
                log.info(f"  送信ボタンクリック: {result}")
                _wait_for_page_stable(page, timeout=12)
                continue
            else:
                log.warning("  送信ボタンが見つかりませんでした")

        # ログインフォームにいる場合
        elif auth_domain in current_url:
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

    if auth_domain in page.url:
        log.error("  まだ認証ページにいます")
        return False

    log.info("  認証フロー完了！")
    return True


# ══════════════════════════════════════════════════════════
#  PDF ページ遷移
# ══════════════════════════════════════════════════════════

def _navigate_to_pdf_page(ctx: BillingContext, page) -> None:
    """WCOシステム上のPDFダウンロードページへ遷移する"""
    # 書面発行ページ
    try:
        cert_link = page.locator('a[href*="/wco/certificate/"]')
        if cert_link.count() > 0:
            cert_link.first.click()
        else:
            page.goto(ctx.config.certificate_url, wait_until="networkidle")
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
                page.goto(ctx.config.bill_pdf_url, wait_until="networkidle")
        page.wait_for_load_state("networkidle")
        time.sleep(2)
        log.info(f"  PDF印刷ページ後のURL: {page.url}")
    except Exception:
        pass


# ══════════════════════════════════════════════════════════
#  ログイン・ナビゲーション
# ══════════════════════════════════════════════════════════

def do_login_and_navigate(ctx: BillingContext, page, phone_number: str, password: str,
                          is_cancelled: bool = False) -> bool:
    """ログイン → 2FA → PDFダウンロードページまで一気に遷移する。"""
    if ctx.config.carrier_family == "au":
        return _do_au_login_and_navigate(ctx, page, phone_number, password)
    bill_pdf_url = ctx.config.bill_pdf_url
    wco_base_domain = ctx.config.wco_base.replace("https://", "")

    # Step 1: PDFページに直接アクセス
    log.info("PDFページに直接アクセス中（認証ページにリダイレクトされるはず）...")
    retry_with_backoff(
        lambda: page.goto(bill_pdf_url, wait_until="networkidle"),
        max_retries=3, retryable_exceptions=(PlaywrightTimeout, ConnectionError), logger=log,
    )
    page.wait_for_load_state("networkidle")
    time.sleep(2)
    log.info(f"  リダイレクト先URL: {page.url}")

    if not _is_on_auth_page(ctx, page):
        log.info("  認証なしでPDFページに到達しました")
        return True

    # Step 2: ページ状態を確認
    page_text = _get_page_text(page)

    if "セキュリティ番号" in page_text or "送付先" in page_text:
        log.info("  セキュリティ番号ページを検出（セッション再利用）→ ログインをスキップして2FAフローへ")
        if not _handle_security_code_flow(ctx, page, phone_number, password):
            return False
    else:
        # SoftBank ID対応: ログインIDがあればそちらを使用
        login_id = ctx.phone_login_id_map.get(phone_number)
        if is_cancelled and not login_id:
            log.error(
                f"  解約済回線 {phone_number} にSoftBank IDが設定されていません。\n"
                f"  解約後は電話番号でログインできないため、回線管理スプレッドシートの「ID」列に"
                f"SoftBank IDを入力してください。"
            )
            return False
        login_value = login_id if login_id else phone_number
        login_label = f"SoftBank ID: {login_id[:4]}***" if login_id else f"電話番号: {phone_number}"
        log.info(f"ログイン情報を入力中...（{login_label}）")
        phone_input = (
            page.locator('input[name="telnum"]')
            .or_(page.locator('input.sbid-msn-check'))
            .or_(page.locator('input[name="username"]'))
            .or_(page.locator('input[name="msn"]'))
            .or_(page.locator('input[name="loginId"]'))
        )

        try:
            phone_input.first.wait_for(state="visible", timeout=10000)
            phone_input.first.fill(login_value)
            log.info(f"  {'SoftBank ID' if login_id else '電話番号'}を入力しました")

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

    # Step 3: 解約済SoftBank回線の場合、authページでもセキュリティ番号フローに入らない
    # SoftBank IDログイン成功後、redirect_uriがmy.softbank.jpを経由するためエラーになるが
    # WCO書面発行ページに直接遷移すればSoftBank IDのセッションクッキーで認証が通る
    if is_cancelled and login_id and _is_on_auth_page(ctx, page):
        log.info("  解約済回線: authページにいますが、form.submit()でリダイレクトを試みます")
        try:
            page.evaluate("document.querySelector('form').submit()")
            time.sleep(5)
            log.info(f"  form.submit()後のURL: {page.url}")
        except Exception:
            pass

        # エラーページまたはまだauthの場合、WCO書面発行ページに直接遷移
        current_url = page.url
        if "error" in current_url or _is_on_auth_page(ctx, page):
            log.info("  → WCO書面発行ページに直接遷移します")
            try:
                page.goto(ctx.config.bill_pdf_url, wait_until="networkidle")
                time.sleep(3)
                log.info(f"  WCO遷移後のURL: {page.url}")
                # WCOアクセスでauthにリダイレクトされた場合、SoftBank IDで再ログイン
                if _is_on_auth_page(ctx, page):
                    phone_input2 = page.locator('input[name="telnum"]').or_(page.locator('input[name="loginId"]')).or_(page.locator('input[name="username"]'))
                    try:
                        if phone_input2.first.is_visible(timeout=3000):
                            phone_input2.first.fill(login_id)
                            pw_input2 = page.locator('input[type="password"]')
                            pw_input2.first.fill(password)
                            page.evaluate("document.querySelector('form').submit()")
                            time.sleep(5)
                            log.info(f"  WCO再ログイン後のURL: {page.url}")
                    except Exception:
                        pass
            except Exception as e:
                log.error(f"  WCO直接遷移に失敗: {e}")

            # 最終確認
            final = page.url
            if wco_base_domain in final:
                log.info("  WCOドメインに到達しました")
                try:
                    combobox = page.get_by_role("combobox").first
                    if combobox.is_visible(timeout=5000):
                        log.info("  PDFダウンロードページに到達しました！（解約済WCO直接遷移）")
                        return True
                except Exception:
                    pass
                pdf_link = page.locator('a[href*="doPrint"]')
                if pdf_link.count() > 0:
                    log.info("  PDFダウンロードページに到達しました！（解約済WCO直接遷移・PDFリンク検出）")
                    return True

            log.error(f"  解約済回線のWCO直接遷移でもPDFページに到達できませんでした (URL: {page.url})")
            return False

    # Step 3b: 通常のセキュリティ番号フロー
    if _is_on_auth_page(ctx, page):
        error_msg = page.locator(".err-area, .error, .alert-error, .sbid-error")
        if error_msg.count() > 0:
            try:
                err_text = error_msg.first.text_content()
                log.error(f"  ログインエラー: {err_text}")
            except Exception:
                pass

        log.info("  認証ページにいます → セキュリティ番号フローを処理します（1回目）")
        if not _handle_security_code_flow(ctx, page, phone_number, password):
            return False

    log.info(f"  1回目の認証後のURL: {page.url}")

    # Step 4: PDFページへの到達確認・ナビゲーション
    if not _is_on_auth_page(ctx, page):
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
            log.info(f"  書面発行ディープリンクへアクセス: {ctx.config.login_url}")
            page.goto(ctx.config.login_url, wait_until="networkidle")
            time.sleep(2)
            log.info(f"  ディープリンク後のURL: {page.url}")
        except Exception:
            pass

        # 2回目の認証が必要な場合
        if _is_on_auth_page(ctx, page):
            log.info("  2回目の認証が必要です")
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

            if _is_on_auth_page(ctx, page):
                if not _handle_security_code_flow(ctx, page, phone_number, password):
                    return False

        # WCOシステム上のPDFページへ遷移
        _navigate_to_pdf_page(ctx, page)

    # Step 5: 2回目の認証（WCOアクセス時）
    if _is_on_auth_page(ctx, page):
        log.info("  WCOアクセスで2回目の認証が必要です")
        if not _handle_security_code_flow(ctx, page, phone_number, password):
            return False

    # 最終確認
    final_url = page.url
    log.info(f"  最終URL: {final_url}")

    if _is_on_auth_page(ctx, page):
        log.error("  認証を完了できませんでした")
        return False

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

    # WCOドメインにはいるが、PDFページではない
    if wco_base_domain in final_url:
        log.info("  WCOドメインにいますが、PDFページではありません。再遷移を試みます...")
        try:
            page.goto(bill_pdf_url, wait_until="networkidle")
            time.sleep(2)
            combobox = page.get_by_role("combobox").first
            if combobox.is_visible(timeout=5000):
                log.info("  PDFダウンロードページに到達しました！")
                return True
        except Exception:
            pass

    log.error(f"  PDFダウンロードページへの到達に失敗 (URL: {final_url})")
    return False


# ══════════════════════════════════════════════════════════
#  月選択・PDFダウンロード
# ══════════════════════════════════════════════════════════

def select_target_month(ctx: BillingContext, page, year: str, month: str) -> bool:
    """PDFページで対象月を選択する。"""
    if ctx.config.carrier_family == "au":
        return _au_select_target_month(ctx, page, year, month)
    target_value = f"{year}{month}"
    target_label = f"{year}年{month}月"
    log.info(f"対象月を選択中: {target_label} (value={target_value})")

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


def _download_single_pdf(ctx: BillingContext, page, link, label: str, save_dir: Path,
                          year: str, month: str, phone: str, amount: str) -> bool:
    """PDFリンクを1件ダウンロードして保存（またはDriveにアップロード）する。最大3回リトライ。"""
    dest = None
    for attempt in range(3):
        try:
            link_text = link.text_content() or ""
            href = link.get_attribute("href") or ""
            if attempt == 0:
                log.info(f"  [{label}] PDFリンク: {link_text.strip()} ({href})")
            with page.expect_download(timeout=60000) as download_info:
                link.click()
            download = download_info.value
            filename = build_filename(ctx, year, month, phone, amount)
            dest = save_dir / filename
            if dest.exists():
                stem = dest.stem
                dest = save_dir / f"{stem}_{label}.pdf"
            download.save_as(str(dest))

            if ctx.drive_ctx is not None:
                folder_id = ctx.drive_ctx.get_folder_id(year, month)
                ok = ctx.drive_ctx.upload(dest, folder_id)
                try:
                    dest.unlink()
                except Exception:
                    pass
                return ok

            log.info(f"  [{label}] 保存完了: {dest}")
            return True
        except Exception as e:
            if dest and dest.exists():
                try:
                    dest.unlink()
                except Exception:
                    pass
            if attempt < 2:
                delay = 2 * (2 ** attempt)
                log.warning(f"  [{label}] ダウンロード失敗 (試行{attempt+1}/3): {e} → {delay}秒後にリトライ")
                time.sleep(delay)
            else:
                log.warning(f"  [{label}] ダウンロード失敗 (全3回試行): {e}")
    return False


def download_pdf_from_page(
    ctx: BillingContext, page, save_dir: Path, year: str, month: str, phone: str,
    pdf_types: set[str] | None = None,
) -> bool:
    """PDFダウンロードページから指定種別のPDFをダウンロードする。"""
    if ctx.config.carrier_family == "au":
        return _au_download_pdf_from_page(ctx, page, save_dir, year, month, phone, pdf_types)
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
            if _download_single_pdf(ctx, page, bulk_link.first, "一括", save_dir, year, month, phone, amount):
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
                    if _download_single_pdf(ctx, page, link, "電話番号別", save_dir, year, month, phone, amount):
                        any_success = True
                elif "機種別" in link_text and "機種別" in pdf_types:
                    if _download_single_pdf(ctx, page, link, "機種別", save_dir, year, month, phone, amount):
                        any_success = True
        else:
            log.info("  電話番号別/機種別PDFリンクが見つかりませんでした")

    if not any_success:
        log.error("指定した種別のPDFがダウンロードできませんでした")
    return any_success


# ══════════════════════════════════════════════════════════
#  1回線分のダウンロード
# ══════════════════════════════════════════════════════════

def download_billing_pdf(
    ctx: BillingContext,
    phone_number: str,
    password: str,
    year: str,
    month: str,
    save_dir: Path,
    pdf_types: set[str] | None = None,
    is_cancelled: bool = False,
) -> str:
    """1回線分のPDFダウンロード処理。"success"/"skipped"/"failed" を返す。"""
    if pdf_types is None:
        pdf_types = {"電話番号別"}
    status_label = "（解約済）" if is_cancelled else ""
    log.info(f"=== {phone_number}{status_label} の処理を開始 (PDFの種類: {', '.join(sorted(pdf_types))}) ===")

    all_done, remaining_types = check_already_downloaded(ctx, save_dir, year, month, phone_number, pdf_types)
    if all_done:
        return "skipped"
    if remaining_types != pdf_types:
        log.info(f"  未ダウンロード: {', '.join(sorted(remaining_types))}")
    pdf_types = remaining_types

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=ctx.headless)

        ctx_kwargs = dict(
            accept_downloads=True,
            locale="ja-JP",
            user_agent=USER_AGENT,
        )
        session_f = _session_file(ctx, phone_number)
        session_loaded = False
        if session_f.exists():
            log.info(f"  保存済みセッションを読み込み: {session_f}")
            ctx_kwargs["storage_state"] = str(session_f)
            session_loaded = True
        context = browser.new_context(**ctx_kwargs)
        page = context.new_page()
        page.set_default_timeout(30000)

        success = False
        try:
            if not do_login_and_navigate(ctx, page, phone_number, password, is_cancelled):
                if session_loaded:
                    log.info("  セッションが無効の可能性 → セッションを削除して再試行")
                    session_f.unlink(missing_ok=True)
                    context.close()
                    clean_kwargs = {k: v for k, v in ctx_kwargs.items() if k != "storage_state"}
                    context = browser.new_context(**clean_kwargs)
                    page = context.new_page()
                    page.set_default_timeout(30000)
                    if not do_login_and_navigate(ctx, page, phone_number, password, is_cancelled):
                        _save_debug_screenshot(
                            ctx, page, save_dir, phone_number, year, month,
                            "ログインまたはページ遷移失敗",
                            f"PDFダウンロードページに到達できませんでした。最終URL: {page.url}",
                        )
                        return "failed"
                else:
                    _save_debug_screenshot(
                        ctx, page, save_dir, phone_number, year, month,
                        "ログインまたはページ遷移失敗",
                        f"PDFダウンロードページに到達できませんでした。最終URL: {page.url}",
                    )
                    return "failed"

            try:
                context.storage_state(path=str(session_f))
                log.info(f"  セッションを保存しました: {session_f}")
            except Exception as e:
                log.warning(f"  セッション保存に失敗: {e}")

            # au系は _au_download_pdf_from_page 内で月選択も行うためスキップ
            if ctx.config.carrier_family != "au":
                if not select_target_month(ctx, page, year, month):
                    _save_debug_screenshot(
                        ctx, page, save_dir, phone_number, year, month,
                        "月選択失敗", f"対象月 {year}{month} の選択に失敗しました",
                    )
                    return "failed"

            success = download_pdf_from_page(ctx, page, save_dir, year, month, phone_number, pdf_types)
            if not success:
                _save_debug_screenshot(
                    ctx, page, save_dir, phone_number, year, month,
                    "PDFダウンロード失敗", "PDFリンクが見つからないかダウンロードに失敗しました",
                )

        except PlaywrightTimeout as e:
            log.error(f"タイムアウトエラー: {e}")
            _save_debug_screenshot(ctx, page, save_dir, phone_number, year, month, "タイムアウト", str(e))
        except Exception as e:
            log.error(f"エラーが発生しました: {e}")
            _save_debug_screenshot(ctx, page, save_dir, phone_number, year, month, "エラー", str(e))
        finally:
            if not success and ctx.drive_ctx is None:
                for partial in save_dir.glob(f"{year}{month}_{ctx.config.carrier_name}_{phone_number}_*.pdf"):
                    try:
                        if partial.stat().st_size == 0:
                            partial.unlink()
                            log.info(f"  空の一時ファイルを削除: {partial.name}")
                    except Exception:
                        pass
            context.close()
            browser.close()

    return "success" if success else "failed"


# ══════════════════════════════════════════════════════════
#  メインエントリポイント
# ══════════════════════════════════════════════════════════

def run_main(ctx: BillingContext) -> None:
    """メイン関数: 全回線のPDFダウンロードを実行する。"""
    log.info(f"{ctx.config.display_name} 料金明細PDFダウンロードを開始します")
    _cleanup_old_debug_screenshots(ctx)

    # 対象月の決定
    year, month = get_target_month(ctx)
    if not ctx.target_month:
        try:
            gc = get_gspread_client()
            sh = open_sheet(gc, ctx.spreadsheet_id)
            ws = sh.worksheet("設定")
            for row in ws.get_all_records():
                if str(row.get("設定名", "")).strip() == "対象月":
                    raw_val = row.get("値", "")
                    val = str(raw_val).strip()
                    if re.match(r"^\d{6}$", val):
                        year, month = val[:4], val[4:6]
                        log.info(f"対象月（設定シートから取得）: {year}年{month}月")
                    elif m2 := re.match(r"^(\d{4})年(\d+)月$", val):
                        year, month = m2.group(1), m2.group(2).zfill(2)
                        log.info(f"対象月（設定シートから取得）: {year}年{month}月")
                    elif isinstance(raw_val, datetime):
                        # スプレッドシートがDate型で返す場合（「2026年1月」→ 2026-01-01）
                        year, month = str(raw_val.year), str(raw_val.month).zfill(2)
                        log.info(f"対象月（設定シートから取得・Date型）: {year}年{month}月")
                    else:
                        log.warning(f"対象月の値を解析できません: '{val}' (型: {type(raw_val).__name__})")
                    break
        except RecursionError:
            raise
        except Exception as e:
            log.warning(f"設定シートから対象月の取得に失敗: {type(e).__name__}: {e}")
    log.info(f"対象月: {year}年{month}月")

    # 保存先の決定
    base_path = resolve_save_path(ctx)
    save_dir = ensure_save_dir(ctx, base_path, year, month)
    mode_label = "Drive APIモード" if ctx.drive_ctx else "ローカルモード"
    log.info(f"保存先: {save_dir} ({mode_label})")

    # 回線情報の読み込み
    accounts = load_accounts(ctx)

    if len(accounts) == 0:
        log.info(f"{ctx.config.display_name} の対象回線がないためスキップします")
        return

    if ctx.dry_run:
        log.info("=== ドライランモード（接続テスト） ===")
        log.info(f"  保存先: {save_dir} ({mode_label})")
        log.info(f"  対象月: {year}年{month}月")
        log.info(f"  対象回線: {len(accounts)} 件")
        for _, row in accounts.iterrows():
            phone = str(row["電話番号"]).strip()
            pdf_types_r = parse_pdf_types(row.get("PDFの種類", ""), ctx.config.carrier_family)
            log.info(f"    {phone} ({', '.join(pdf_types_r)})")
        if ctx.drive_ctx:
            try:
                folder_id = ctx.drive_ctx.get_folder_id(year, month)
                log.info(f"  Drive接続テスト: OK (フォルダID: {folder_id})")
            except Exception as e:
                log.error(f"  Drive接続テスト: 失敗 ({e})")
        if len(accounts) > 0:
            phone = str(accounts.iloc[0]["電話番号"]).strip()
            log.info(f"  ページアクセステスト: {phone} でログインページにアクセス中...")
            try:
                with sync_playwright() as pw:
                    br = pw.chromium.launch(headless=True)
                    bctx = br.new_context(locale="ja-JP", user_agent=USER_AGENT)
                    pg = bctx.new_page()
                    test_url = ctx.config.bill_pdf_url or ctx.config.au_billing_top_url or ctx.config.login_url
                    pg.goto(test_url, wait_until="networkidle", timeout=15000)
                    log.info(f"  ページアクセステスト: OK (URL: {pg.url})")
                    bctx.close()
                    br.close()
            except Exception as e:
                log.error(f"  ページアクセステスト: 失敗 ({e})")
        log.info("ドライランモード完了。実際のダウンロードは行いません。")
        return

    if ctx.retry_phones:
        before = len(accounts)
        accounts = accounts[accounts["電話番号"].isin(ctx.retry_phones)].reset_index(drop=True)
        log.info(f"  リトライモード: {len(accounts)}/{before} 件に絞り込み")

    # 各回線についてPDFをダウンロード
    results = []
    downloaded_filenames = {}  # phone -> [filename, ...]
    for _, row in accounts.iterrows():
        phone = str(row["電話番号"]).strip()
        pw = str(row["パスワード"]).strip()
        pdf_types = parse_pdf_types(row.get("PDFの種類", ""), ctx.config.carrier_family)
        is_cancelled = str(row.get("状態", "")).strip() == "解約済"

        result = download_billing_pdf(ctx, phone, pw, year, month, save_dir, pdf_types, is_cancelled)
        results.append((phone, result))
        if result == "success":
            # ファイル名を構築して記録（Drive APIモードではファイルが既に削除済みのため）
            filename = build_filename(ctx, year, month, phone)
            downloaded_filenames[phone] = [filename]

    # ダウンロード履歴をスプレッドシートに記録（スキップ分は記録しない）
    history_results = [(p, r == "success") for p, r in results if r != "skipped"]
    if history_results:
        write_download_history(ctx.spreadsheet_id, ctx.config.carrier_name, history_results, year, month, save_dir, downloaded_filenames)

    # Drive APIモード: 一時ディレクトリのクリーンアップ
    if ctx.temp_save_dir is not None and ctx.temp_save_dir.exists():
        try:
            shutil.rmtree(str(ctx.temp_save_dir))
        except Exception as e:
            log.warning(f"一時ディレクトリの削除に失敗: {e}")

    # キャリア名付きの結果を返す（download.pyで集約表示）
    return [(ctx.config.carrier_name, phone, result) for phone, result in results]


# ══════════════════════════════════════════════════════════
#  au / UQmobile 固有ロジック
# ══════════════════════════════════════════════════════════

def _is_on_au_auth_page(ctx: BillingContext, page) -> bool:
    """au の認証ページ上にいるか判定する"""
    url = page.url
    return "connect.auone.jp" in url or "id.auone.jp" in url


def _do_au_login_and_navigate(ctx: BillingContext, page, phone_number: str, password: str) -> bool:
    """au ID でログイン → WEB de 請求書ページまで遷移する。"""

    # Step 1: au IDログインページに直接アクセス（targeturl付きでセッション生成）
    import urllib.parse
    target = ctx.config.au_billing_top_url or "https://id.auone.jp/index.html"
    login_url = (
        "https://connect.auone.jp/net/vwc/cca_lg_eu_nets/login"
        f"?targeturl={urllib.parse.quote(target, safe='')}"
    )
    log.info(f"au IDログインページにアクセス中...")
    retry_with_backoff(
        lambda: page.goto(login_url, wait_until="networkidle"),
        max_retries=3, retryable_exceptions=(PlaywrightTimeout, ConnectionError), logger=log,
    )
    page.wait_for_load_state("networkidle")
    time.sleep(2)
    log.info(f"  現在のURL: {page.url}")

    if not _is_on_au_auth_page(ctx, page):
        log.info("  認証済みのためログインをスキップ → 請求書ページへ遷移")
        return True

    # Step 2: au ID ログイン（2ステップ: ID入力 → パスワード入力）
    login_id = ctx.phone_login_id_map.get(phone_number, phone_number)
    if login_id != phone_number:
        log.info(f"au ID ログイン情報を入力中...（au ID: {login_id[:3]}***）")
    else:
        log.info("au ID ログイン情報を入力中...（au ID未設定のため電話番号を使用）")
    try:
        # au ID入力
        id_input = (
            page.locator('#loginAliasId')
            .or_(page.locator('input[name="loginAliasId"]'))
            .or_(page.locator('input[name="username"]'))
            .or_(page.locator('input[type="text"]').first)
        )
        id_input.first.wait_for(state="visible", timeout=10000)
        id_input.first.fill(login_id)
        log.info("  au IDを入力しました")

        # 「次へ」ボタンをクリック（2ステップログインの場合のみ）
        # パスワード欄がすでに表示されていればシングルステップなのでスキップ
        pw_visible = page.locator('input[type="password"]').first.is_visible(timeout=1000) if True else False
        try:
            pw_visible = page.locator('input[type="password"]').first.is_visible(timeout=1000)
        except Exception:
            pw_visible = False

        if not pw_visible:
            next_btn = (
                page.locator('#btn_idInput')
                .or_(page.get_by_text("次へ", exact=True))
            )
            try:
                if next_btn.first.is_visible(timeout=3000):
                    next_btn.first.click()
                    page.wait_for_load_state("networkidle")
                    time.sleep(2)
                    log.info("  「次へ」をクリックしました")
            except Exception:
                pass
        else:
            log.info("  シングルステップログイン（パスワード欄が表示済み）")

        # パスワード入力
        pw_input = (
            page.locator('#loginAuonePwd')
            .or_(page.locator('input[name="loginAuonePwd"]'))
            .or_(page.locator('input[type="password"]'))
        )
        pw_input.first.wait_for(state="visible", timeout=10000)
        pw_input.first.fill(password)
        log.info("  パスワードを入力しました")

        # ログインボタンクリック
        login_btn = (
            page.locator('#btn_pwdLogin')
            .or_(page.get_by_role("button", name=re.compile(r"^ログイン$")))
            .or_(page.locator('button[type="submit"]'))
            .or_(page.locator('input[type="submit"]'))
        )
        _click_any_button(page, login_btn, "ログインボタン", text_hint="ログイン")

        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass
        time.sleep(3)
        log.info(f"  ログイン後のURL: {page.url}")

    except (PlaywrightTimeout, Exception) as e:
        log.error(f"  au IDログインフォームの操作に失敗: {e}")
        return False

    # Step 3: 2段階認証（SMS確認コードまたはワンタイムURL）
    if _is_on_au_auth_page(ctx, page):
        page_text = _get_page_text(page)

        # SMS確認コード入力
        if "確認コード" in page_text or "セキュリティ" in page_text or "認証" in page_text:
            log.info("  au 2段階認証を検出しました")
            if not _handle_au_2fa(ctx, page, phone_number):
                return False
        else:
            # ワンタイムURL方式の場合、ユーザーにSMSのURLをタップするよう案内
            log.info("  au認証ページにいます。ワンタイムURL方式の可能性があります")
            log.info("  SMSに届いたURLをタップして認証を完了してください")
            # URLタップ後のページ遷移を待機
            try:
                page.wait_for_url(lambda url: "connect.auone.jp" not in url and "id.auone.jp" not in url,
                                  timeout=ctx.security_code_timeout * 1000)
                log.info(f"  認証完了: {page.url}")
            except PlaywrightTimeout:
                log.error("  au 2段階認証がタイムアウトしました")
                return False

    # Step 4: 暗証番号（4桁PIN）が必要な場合
    page_text = _get_page_text(page)
    if "暗証番号" in page_text:
        log.info("  暗証番号の入力を検出しました")
        pin = _load_au_pin(ctx)
        if not pin:
            log.error(
                "au暗証番号が設定されていません。\n"
                f"  設定シートの「{ctx.config.au_pin_setting_name}」行に4桁の暗証番号を入力してください。"
            )
            return False
        pin_input = (
            page.locator('input[type="password"]')
            .or_(page.locator('input[type="tel"]'))
            .or_(page.locator('input[maxlength="4"]'))
        )
        try:
            pin_input.first.wait_for(state="visible", timeout=5000)
            pin_input.first.fill(pin)
            submit_btn = page.locator('input[type="submit"]').or_(page.locator('button[type="submit"]'))
            _click_any_button(page, submit_btn, "暗証番号送信ボタン")
            page.wait_for_load_state("networkidle")
            time.sleep(2)
            log.info("  暗証番号を入力しました")
        except Exception as e:
            log.error(f"  暗証番号の入力に失敗: {e}")
            return False

    # Step 5: 請求書ダウンロードページへ遷移
    final_url = page.url
    log.info(f"  最終URL: {final_url}")

    if _is_on_au_auth_page(ctx, page):
        log.error("  au認証を完了できませんでした")
        return False

    # 「請求書/領収書/支払証明書の保存・印刷」リンクをクリック
    billing_link_texts = [
        "請求書/領収書/支払証明書の保存・印刷",
        "請求書/領収書/支払証明書",
        "保存・印刷",
        "請求書一括印刷",
    ]
    for text in billing_link_texts:
        try:
            link = page.get_by_text(text, exact=False).first
            if link.is_visible(timeout=3000):
                link.click()
                page.wait_for_load_state("networkidle")
                time.sleep(2)
                log.info(f"  「{text}」をクリック → {page.url}")
                return True
        except Exception:
            continue

    # リンクが見つからない場合、現在のページで続行
    log.info(f"  請求ページに到達: {page.url}")
    return True


def _handle_au_2fa(ctx: BillingContext, page, phone_number: str) -> bool:
    """au 2段階認証の確認コード入力を処理する。"""
    for attempt in range(8):
        page_text = _get_page_text(page)

        if not _is_on_au_auth_page(ctx, page):
            log.info("  au 2段階認証完了")
            return True

        # 確認コード入力欄を検出
        code_input = (
            page.locator('#confirmcode')
            .or_(page.locator('input[name="confirmcode"]'))
            .or_(page.locator('input[name="otp"]'))
            .or_(page.locator('input[maxlength="6"]'))
            .or_(page.locator('input[maxlength="4"]'))
        )
        try:
            if code_input.first.is_visible(timeout=3000):
                code = ask_security_code(ctx, phone_number)
                if not code:
                    return False
                code_input.first.fill(code)
                log.info("  確認コードを入力しました")

                # 「次へ」ボタンをクリック（au 2段階認証ページ）
                # JavaScriptでフォーム送信するボタンのため、複数の方法を試す
                before_url = page.url
                clicked = False

                # 方法1: フォームを直接submit
                try:
                    result = page.evaluate("""() => {
                        const form = document.querySelector('form');
                        if (form) { form.submit(); return 'form-submit'; }
                        return null;
                    }""")
                    if result:
                        clicked = True
                        log.info(f"  確認コード送信: {result}")
                except Exception:
                    pass

                # 方法2: 「次へ」ボタンをクリック（form submitが効かなかった場合）
                if not clicked:
                    for selector in ['a:has-text("次へ")', 'button:has-text("次へ")',
                                     '#btn_submit', 'input[type="submit"]', 'button[type="submit"]']:
                        try:
                            el = page.locator(selector).first
                            if el.is_visible(timeout=1000):
                                el.click()
                                clicked = True
                                log.info(f"  確認コード送信: {selector} をクリック")
                                break
                        except Exception:
                            continue

                if not clicked:
                    try:
                        page.get_by_text("次へ").first.click()
                        clicked = True
                        log.info("  確認コード送信: get_by_text('次へ') をクリック")
                    except Exception:
                        log.warning("  確認コード送信ボタンが見つかりませんでした")

                # クリック後のページ遷移を待つ
                try:
                    page.wait_for_load_state("networkidle", timeout=15000)
                except Exception:
                    pass
                time.sleep(3)
                log.info(f"  確認コード送信後のURL: {page.url}")

                if not _is_on_au_auth_page(ctx, page):
                    log.info("  au 2段階認証完了")
                    return True
                # まだ認証ページにいる場合、ページテキストを確認
                pt = _get_page_text(page)
                if "確認コード" not in pt and "認証" not in pt and "2段階" not in pt:
                    log.info("  認証ドメインにいるが認証ページではない → 完了とみなす")
                    return True
                log.info("  まだ認証ページにいます。リトライ...")
                continue
        except Exception:
            pass

        # 「許可する」「次へ」等のボタンを検出
        try:
            allow_btn = (
                page.get_by_text("許可する", exact=False)
                .or_(page.get_by_text("許可", exact=True))
                .or_(page.get_by_text("次へ", exact=False))
            )
            if allow_btn.first.is_visible(timeout=2000):
                allow_btn.first.click()
                page.wait_for_load_state("networkidle")
                time.sleep(2)
                continue
        except Exception:
            pass

        time.sleep(3)

    log.error("  au 2段階認証: 最大試行回数に到達")
    return False


def _load_au_pin(ctx: BillingContext) -> "str | None":
    """設定シートからau暗証番号を取得する。4桁でなければエラー。"""
    if not ctx.config.au_pin_setting_name:
        return None
    try:
        gc = get_gspread_client()
        sh = open_sheet(gc, ctx.spreadsheet_id)
        pin = load_password_from_settings(sh, ctx.config.au_pin_setting_name)
        if pin and len(pin) != 4:
            log.error(f"  au暗証番号が4桁ではありません（{len(pin)}桁）。設定シートの書式がテキストになっているか確認してください。")
            return None
        return pin
    except Exception:
        return None


def _au_select_target_month(ctx: BillingContext, page, year: str, month: str) -> bool:
    """au WEB de 請求書で対象月のラジオボタンを選択する。
    ラジオボタンのvalue形式: "{index}_{YYYYMM}" (例: "1_202602")
    """
    target_ym = f"{year}{month}"
    log.info(f"対象月を選択中: {year}年{month}月 (au WEB de 請求書)")

    try:
        # ラジオボタンから対象月を探す（value末尾が _YYYYMM のもの）
        radios = page.locator('input[type="radio"][name="bill"]')
        for i in range(radios.count()):
            radio = radios.nth(i)
            value = radio.get_attribute("value") or ""
            if value.endswith(f"_{target_ym}"):
                radio.check()
                log.info(f"  月を選択しました: value={value}")
                time.sleep(1)
                return True

        log.error(f"  対象月 {target_ym} のラジオボタンが見つかりませんでした")
        return False

    except Exception as e:
        log.error(f"対象月の選択に失敗しました: {e}")
        return False


def _au_download_pdf_from_page(
    ctx: BillingContext, page, save_dir: Path, year: str, month: str, phone: str,
    pdf_types: set[str] | None = None,
) -> bool:
    """au WEB de 請求書からPDFをダウンロードする。
    pdf_types: {"請求書"}, {"領収書"}, {"支払証明書"} のいずれか
    各種別ごとにダウンロードページのURLパラメータが異なる:
      請求書: DlCals=01, 領収書: DlCals=02, 支払証明書: DlCals=03
    """
    if pdf_types is None:
        pdf_types = {"請求書"}

    log.info(f"ダウンロード対象: {', '.join(sorted(pdf_types))}")

    dl_params = {"請求書": "01", "領収書": "02", "支払証明書": "03"}
    any_success = False

    for pdf_type in sorted(pdf_types):
        dl_cal = dl_params.get(pdf_type)
        if not dl_cal:
            log.warning(f"  不明なPDFの種類: {pdf_type}")
            continue

        log.info(f"{pdf_type}のダウンロードを試みます...")

        # 対応するダウンロードページへ遷移（種別ごとにURLが異なる）
        agdt = "2" if pdf_type == "支払証明書" else "1"
        dl_url = f"https://my.au.com/aus/seikyu/download?agdt={agdt}&DlCals={dl_cal}"
        try:
            page.goto(dl_url, wait_until="domcontentloaded", timeout=60000)
            time.sleep(3)
            log.info(f"  ダウンロードページ: {page.url}")
        except Exception as e:
            log.error(f"  ダウンロードページへの遷移に失敗: {e}")
            continue

        # 回線選択画面が表示される場合（複数契約が同一ログインIDに紐づく場合）
        number_radios = page.locator('input[type="radio"][name="number"]')
        if number_radios.count() > 0:
            log.info(f"  回線選択画面を検出（{number_radios.count()}件の契約）")
            # 電話番号にマッチする回線を選択
            phone_digits = re.sub(r'\D', '', phone)
            selected = False
            for i in range(number_radios.count()):
                radio = number_radios.nth(i)
                # ラジオボタンの隣接テキストから電話番号を取得
                label_text = ""
                try:
                    label_text = radio.evaluate("""el => {
                        const label = el.closest('label') || el.parentElement;
                        return label ? label.textContent.trim() : '';
                    }""")
                except Exception:
                    pass
                label_digits = re.sub(r'\D', '', label_text.split('\n')[0] if label_text else "")
                log.info(f"    契約[{i}]: label={label_text.strip()[:60]}")
                if phone_digits in label_digits or label_digits in phone_digits:
                    radio.check(force=True)
                    selected = True
                    log.info(f"    → 選択しました")
                    break
            if not selected:
                log.error(f"  電話番号 {phone} にマッチする回線が見つかりません")
                continue

            # 「選択」ボタンをクリック
            select_btn = page.get_by_text("選択", exact=True).or_(page.locator('button:has-text("選択")'))
            try:
                select_btn.first.click()
                page.wait_for_load_state("networkidle")
                time.sleep(3)
                log.info(f"  回線選択後のURL: {page.url}")
            except Exception as e:
                log.error(f"  回線選択ボタンのクリックに失敗: {e}")
                continue

        # 対象月を選択（請求書/領収書: ラジオボタン、支払証明書: チェックボックス）
        target_ym = f"{year}{month}"
        month_selected = False

        # まずラジオボタン（請求書・領収書）を試す
        radios = page.locator('input[type="radio"][name="bill"]')
        if radios.count() > 0:
            for i in range(radios.count()):
                radio = radios.nth(i)
                value = radio.get_attribute("value") or ""
                if value.endswith(f"_{target_ym}"):
                    radio.check(force=True)
                    log.info(f"  月を選択（ラジオ）: value={value}")
                    month_selected = True
                    break

        # ラジオボタンがない場合、チェックボックス（支払証明書）を試す
        if not month_selected:
            checkboxes = page.locator('input[type="checkbox"]')
            cb_count = checkboxes.count()
            if cb_count > 0:
                log.info(f"  チェックボックス形式を検出（{cb_count}件）")
                # マッチパターン: value末尾 "_YYYYMM"、value含む "YYYYMM"、
                # またはラベルテキストに "YYYY年M月" を含む
                target_year_month_text = f"{int(year)}年{int(month)}月"  # "2026年1月"
                target_found = False
                for i in range(cb_count):
                    cb = checkboxes.nth(i)
                    value = cb.get_attribute("value") or ""
                    label_text = ""
                    try:
                        label_text = cb.evaluate("""el => {
                            const label = el.closest('label') || el.parentElement;
                            return label ? label.textContent.trim() : '';
                        }""")
                    except Exception:
                        pass
                    is_target = (
                        value.endswith(f"_{target_ym}")
                        or target_ym in value
                        or target_year_month_text in label_text
                    )
                    if is_target:
                        if not cb.is_checked():
                            cb.check(force=True)
                        log.info(f"  対象月をチェック: value={value}, label={label_text[:40]}")
                        target_found = True
                    else:
                        if cb.is_checked():
                            cb.uncheck(force=True)
                if target_found:
                    month_selected = True
                else:
                    # デバッグ: 実際のチェックボックス値を列挙
                    for i in range(min(cb_count, 10)):
                        cb = checkboxes.nth(i)
                        v = cb.get_attribute("value") or ""
                        lbl = ""
                        try:
                            lbl = cb.evaluate("el => (el.closest('label') || el.parentElement)?.textContent?.trim()?.substring(0, 50) || ''")
                        except Exception:
                            pass
                        log.info(f"    cb[{i}]: value={v}, label={lbl}")

        if not month_selected:
            log.error(f"  対象月 {target_ym} の選択肢が見つかりません (URL: {page.url})")
            continue

        # 金額取得を試みる（回線選択後のページから対象電話番号の金額を取得）
        amount = ""
        try:
            phone_formatted = f"{phone[:3]}-{phone[3:7]}-{phone[7:]}" if len(phone) >= 10 else phone
            # ページテキストから電話番号の近くにある金額を探す
            page_text = page.inner_text("body")
            # 電話番号を含む行の近くにある ( 金額 ) パターン
            idx = page_text.find(phone_formatted)
            if idx == -1:
                idx = page_text.find(phone)
            if idx >= 0:
                # 電話番号の前後200文字から金額を探す
                start = max(0, idx - 50)
                end = min(len(page_text), idx + 200)
                nearby = page_text[start:end]
                # 括弧付き金額 ( 1,851 )
                m_bracket = re.search(r'\(\s*([\d,]+)\s*\)', nearby)
                if m_bracket:
                    amount = sanitize_amount(m_bracket.group(1))
                else:
                    # 「円」付き金額
                    m_yen = re.search(r'([\d,]+)\s*円', nearby)
                    if m_yen:
                        amount = sanitize_amount(m_yen.group(1))
            if not amount:
                # フォールバック: ページ上の金額テキスト
                amount_els = page.locator("text=/[\\d,]+円/").all()
                for el in amount_els:
                    raw = el.text_content() or ""
                    a = sanitize_amount(raw)
                    if a:
                        amount = a
                        break
            if amount:
                log.info(f"  請求金額: {amount}")
        except Exception as e:
            log.info(f"  請求金額の取得をスキップ: {e}")

        # 「選択した期間の請求書のダウンロード」ボタンをクリック
        download_btn = page.get_by_text(re.compile(r"選択した期間の.+ダウンロード"), exact=False)
        try:
            if not download_btn.first.is_visible(timeout=3000):
                download_btn = page.get_by_text("ダウンロード", exact=False)
        except Exception:
            download_btn = page.get_by_text("ダウンロード", exact=False)

        for attempt in range(3):
            try:
                page.on("dialog", lambda dialog: dialog.accept())
                with page.expect_download(timeout=60000) as download_info:
                    download_btn.first.click()
                download = download_info.value
                filename = build_filename(ctx, year, month, phone, amount)
                if pdf_type != "請求書":
                    stem = Path(filename).stem
                    filename = f"{stem}_{pdf_type}.pdf"
                dest = save_dir / filename

                download.save_as(str(dest))

                if ctx.drive_ctx is not None:
                    folder_id = ctx.drive_ctx.get_folder_id(year, month)
                    ok = ctx.drive_ctx.upload(dest, folder_id)
                    try:
                        dest.unlink()
                    except Exception:
                        pass
                    if ok:
                        any_success = True
                    break
                else:
                    log.info(f"  [{pdf_type}] 保存完了: {dest}")
                    any_success = True
                    break
            except Exception as e:
                if attempt < 2:
                    log.warning(f"  [{pdf_type}] ダウンロード失敗 (試行{attempt+1}/3): {e}")
                    time.sleep(2 * (2 ** attempt))
                else:
                    log.warning(f"  [{pdf_type}] ダウンロード失敗 (全3回試行): {e}")

    if not any_success:
        log.error("指定した種別のPDFがダウンロードできませんでした")
    return any_success
