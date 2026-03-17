"""
携帯領収書管理 共通ユーティリティ

SoftBank・Ymobile 両スクリプトから共通で使用する関数群。
"""

import json
import os
import re
import subprocess
import sys
import logging
import webbrowser
from pathlib import Path
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials

log = logging.getLogger(__name__)


# ── 電話番号処理 ──

def strip_hyphens(phone: str) -> str:
    """電話番号からハイフンを除去し、先頭0が消えていたら補完する。
    例: 090-4769-5015 → 09047695015, 7043941930 → 07043941930
    """
    cleaned = re.sub(r"[-\s\u2010-\u2015\u2212\uFF0D]", "", phone)
    if cleaned and cleaned[0] != "0" and len(cleaned) == 10:
        cleaned = "0" + cleaned
    return cleaned


# ── PDFの種類パース ──

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


# ── サービスアカウント・認証 ──

def find_service_account_json() -> Path:
    """service_account.json を探す（呼び出し元スクリプトのディレクトリ → 既知のキャリアフォルダ）"""
    # sys.modules の呼び出し元から推定、またはプロジェクトルートから検索
    project_root = Path(__file__).resolve().parent
    candidates = [
        project_root / "SoftBank" / "service_account.json",
        project_root / "Ymobile" / "service_account.json",
    ]
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError(
        "service_account.json が見つかりません。\n"
        "SoftBank/ または Ymobile/ フォルダに配置してください。"
    )


def find_client_secrets() -> "Path | None":
    """client_secrets.json を探す"""
    project_root = Path(__file__).resolve().parent
    for p in [
        project_root / "SoftBank" / "client_secrets.json",
        project_root / "Ymobile" / "client_secrets.json",
    ]:
        if p.exists():
            return p
    return None


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


# ── 設定シート読み取り ──

def load_password_from_settings(sh) -> "str | None":
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


# ── ダウンロード履歴 ──

def write_download_history(
    spreadsheet_id: str,
    carrier_name: str,
    results: list[tuple[str, bool]],
    year: str,
    month: str,
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
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        rows = []
        for phone, ok in results:
            rows.append([now, carrier_name, phone, f"{year}{month}", "", "成功" if ok else "失敗"])
        if rows:
            ws.append_rows(rows, value_input_option="USER_ENTERED")
            log.info(f"  ダウンロード履歴を記録しました（{len(rows)}件）")
    except Exception as e:
        log.warning(f"ダウンロード履歴の記録に失敗（処理は続行）: {e}")
