"""バックグラウンドワーカー群

PhoneListLoader  : 認証情報シートから電話番号一覧を読み込む
DownloadWorker   : ダウンロード処理をバックグラウンドで実行
UpdateAmountsWorker: 金額更新をバックグラウンドで実行
SpreadsheetReader: 設定・履歴を読み込む
"""

import io
import logging
import os
import re
import sys
import traceback
from pathlib import Path

from PySide6.QtCore import QThread, Signal

# キャリア名 → キャリアファミリー
CARRIER_FAMILY_MAP = {
    "SoftBank": "softbank",
    "Ymobile":  "softbank",
    "au":       "au",
    "UQmobile": "au",
    "docomo":   "docomo",
}


class StdoutRedirector(io.TextIOBase):
    """stdout をリダイレクトして Signal を発火。SMS認証プロンプトも検出。"""

    def __init__(self, signal, sms_callback=None):
        super().__init__()
        self._signal = signal
        self._sms_callback = sms_callback
        self._sms_state = None

    def write(self, text):
        if not text:
            return 0
        self._signal.emit(text)
        if self._sms_callback:
            for line in text.splitlines():
                self._detect_sms(line)
        return len(text)

    def _detect_sms(self, line):
        if "SMS認証が必要です" in line:
            self._sms_state = {"phone": "", "device": "", "code_file": "", "digits": 3}
        elif self._sms_state is not None:
            if m := re.search(r"電話番号:\s*(.+)", line):
                self._sms_state["phone"] = m.group(1).strip()
            elif m := re.search(r"端末\s*:\s*(.+)", line):
                self._sms_state["device"] = m.group(1).strip()
            elif m := re.search(r"(\d+)桁", line):
                self._sms_state["digits"] = int(m.group(1))
            elif m := re.search(r"echo .+ > (.+)", line):
                self._sms_state["code_file"] = m.group(1).strip()
                if self._sms_callback and self._sms_state["phone"]:
                    self._sms_callback(dict(self._sms_state))
                self._sms_state = None

    def flush(self):
        pass

    def isatty(self):
        return False


class SignalLogHandler(logging.Handler):
    """logging の出力を Signal に変換するハンドラ。"""

    def __init__(self, signal):
        super().__init__()
        self._signal = signal

    def emit(self, record):
        self._signal.emit(self.format(record) + "\n")


class PhoneListLoader(QThread):
    """認証情報シートから電話番号一覧を取得する。

    Signal:
        phones_loaded(dict): {family: [{"phone", "carrier", "status", "device"}, ...]}
    """

    phones_loaded = Signal(dict)
    error_signal  = Signal(str)

    def __init__(self, script_dir: Path):
        super().__init__()
        self._script_dir = script_dir

    def run(self):
        try:
            from dotenv import load_dotenv
            from shared_utils import (
                bootstrap_env_from_gsheet, get_gspread_client, open_sheet,
            )
            import re as _re

            bootstrap_env_from_gsheet(self._script_dir, "GUI")
            load_dotenv(self._script_dir / ".env")

            url = os.environ.get("SPREADSHEET_URL", "").strip()
            m = _re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
            spreadsheet_id = m.group(1) if m else os.environ.get("SPREADSHEET_ID", "").strip()
            if not spreadsheet_id:
                self.error_signal.emit("スプレッドシートIDが見つかりません")
                return

            gc = get_gspread_client()
            sh = open_sheet(gc, spreadsheet_id)
            ws = sh.worksheet("認証情報")
            records = ws.get_all_records()

            result: dict[str, list] = {"softbank": [], "au": [], "docomo": []}
            for row in records:
                carrier = str(row.get("キャリア", "")).strip()
                family  = CARRIER_FAMILY_MAP.get(carrier)
                if not family:
                    continue
                phone  = str(row.get("電話番号", "")).strip().replace("-", "").replace(" ", "")
                if not phone:
                    continue
                status = str(row.get("状態", "契約中")).strip()
                device = str(row.get("運用端末", "")).strip()
                result[family].append({
                    "phone":   phone,
                    "carrier": carrier,
                    "status":  status,
                    "device":  device,
                })

            self.phones_loaded.emit(result)

        except Exception as e:
            self.error_signal.emit(f"電話番号の読み込みに失敗しました: {e}")


class DownloadWorker(QThread):
    """ダウンロード処理をバックグラウンドで実行するワーカー。"""

    log_signal      = Signal(str)
    sms_signal      = Signal(dict)
    finished_signal = Signal(list)   # [(carrier, phone, result), ...]
    error_signal    = Signal(str)
    progress_signal = Signal(str)

    def __init__(self, carriers, script_dir: Path, env_overrides: dict = None,
                 selected_phones: list = None):
        super().__init__()
        self._carriers        = carriers
        self._script_dir      = script_dir
        self._env_overrides   = env_overrides or {}
        self._selected_phones = selected_phones or []

    def run(self):
        saved_env = {}

        # selected_phones が指定されていれば RETRY_PHONES に設定
        overrides = dict(self._env_overrides)
        if self._selected_phones:
            overrides["RETRY_PHONES"] = ",".join(self._selected_phones)
        else:
            overrides["RETRY_PHONES"] = None

        for k, v in overrides.items():
            saved_env[k] = os.environ.get(k)
            if v is not None:
                os.environ[k] = v
            elif k in os.environ:
                del os.environ[k]

        log_handler = SignalLogHandler(self.log_signal)
        log_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
        )
        root_logger = logging.getLogger()
        root_logger.addHandler(log_handler)
        root_logger.setLevel(logging.INFO)

        old_stdout = sys.stdout
        old_stderr = sys.stderr
        old_stdin  = sys.stdin

        def sms_cb(info):
            self.sms_signal.emit(info)

        sys.stdout = StdoutRedirector(self.log_signal, sms_cb)
        sys.stderr = StdoutRedirector(self.log_signal)
        sys.stdin  = io.StringIO()

        try:
            from shared_utils import create_billing_context, run_main

            all_results = []
            for config in self._carriers:
                self.progress_signal.emit(f"実行中: {config.display_name}")
                ctx = create_billing_context(config, script_dir=self._script_dir)
                carrier_results = run_main(ctx) or []
                all_results.extend(carrier_results)

            self.finished_signal.emit(all_results)

        except SystemExit:
            self.error_signal.emit("処理が中断されました")
        except Exception as e:
            self.error_signal.emit(f"{e}\n{traceback.format_exc()}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            sys.stdin  = old_stdin
            root_logger.removeHandler(log_handler)
            for k, v in saved_env.items():
                if v is not None:
                    os.environ[k] = v
                elif k in os.environ:
                    del os.environ[k]


class UpdateAmountsWorker(QThread):
    """金額更新をバックグラウンドで実行する。"""

    log_signal      = Signal(str)
    finished_signal = Signal()
    error_signal    = Signal(str)

    def __init__(self, script_dir: Path):
        super().__init__()
        self._script_dir = script_dir

    def run(self):
        log_handler = SignalLogHandler(self.log_signal)
        log_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
        )
        root_logger = logging.getLogger()
        root_logger.addHandler(log_handler)
        root_logger.setLevel(logging.INFO)

        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = StdoutRedirector(self.log_signal)
        sys.stderr = StdoutRedirector(self.log_signal)

        try:
            from download import update_amounts
            update_amounts()
            self.finished_signal.emit()
        except SystemExit:
            self.error_signal.emit("金額更新が中断されました")
        except Exception as e:
            self.error_signal.emit(f"{e}\n{traceback.format_exc()}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            root_logger.removeHandler(log_handler)


class SpreadsheetReader(QThread):
    """設定・履歴をバックグラウンドで読み込む。"""

    settings_loaded      = Signal(list)   # [(key, value), ...]
    history_loaded       = Signal(list)   # [dict, ...]
    target_month_loaded  = Signal(str)
    error_signal         = Signal(str)

    def __init__(self, script_dir: Path):
        super().__init__()
        self._script_dir = script_dir

    def run(self):
        try:
            from dotenv import load_dotenv
            from shared_utils import (
                bootstrap_env_from_gsheet, get_gspread_client, open_sheet,
            )
            import re as _re

            bootstrap_env_from_gsheet(self._script_dir, "GUI")
            load_dotenv(self._script_dir / ".env")

            url = os.environ.get("SPREADSHEET_URL", "").strip()
            m = _re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
            spreadsheet_id = m.group(1) if m else os.environ.get("SPREADSHEET_ID", "").strip()
            if not spreadsheet_id:
                self.error_signal.emit("スプレッドシートIDが見つかりません")
                return

            gc = get_gspread_client()
            sh = open_sheet(gc, spreadsheet_id)

            # 設定シート
            try:
                ws = sh.worksheet("設定")
                settings = []
                for row in ws.get_all_records():
                    name  = str(row.get("設定名", "")).strip()
                    value = str(row.get("値", "")).strip()
                    if name:
                        settings.append((name, value))
                        if name == "対象月":
                            self.target_month_loaded.emit(value)
                self.settings_loaded.emit(settings)
            except Exception as e:
                self.error_signal.emit(f"設定シート読み込みエラー: {e}")

            # 履歴シート
            try:
                ws = sh.worksheet("ダウンロード履歴")
                self.history_loaded.emit(ws.get_all_records())
            except Exception as e:
                self.error_signal.emit(f"履歴シート読み込みエラー: {e}")

        except Exception as e:
            self.error_signal.emit(f"スプレッドシート接続エラー: {e}")


# ─── PDF種類定数 ───
PDF_TYPES_SB = [
    "電話番号別", "一括", "機種別",
    "電話番号別,一括", "電話番号別,機種別", "一括,機種別", "電話番号別,一括,機種別",
]
PDF_TYPES_AU = [
    "請求書", "領収書", "支払証明書",
    "請求書,領収書", "請求書,支払証明書", "領収書,支払証明書", "請求書,領収書,支払証明書",
]
PDF_TYPES_DOCOMO_REP = ["一括請求", "利用内訳", "一括請求,利用内訳"]
PDF_TYPES_DOCOMO     = ["利用内訳"]


def _normalize_carrier(text: str) -> str | None:
    """キャリア名を正規化する（GASの _normalizeCarrierName_ 相当）。"""
    s = text.strip().lower()
    if not s:
        return None
    if s in ("softbank",) or "ソフトバンク" in s:
        return "SoftBank"
    if s in ("ymobile", "y!mobile") or "ワイモバイル" in s:
        return "Ymobile"
    if s in ("au", "kddi") or "エーユー" in s:
        return "au"
    if s in ("uqmobile", "uq mobile", "uq") or "ユーキュー" in s:
        return "UQmobile"
    if s in ("docomo", "nttdocomo") or "ドコモ" in s:
        return "docomo"
    return None


def _normalize_phone(phone: str) -> str:
    """電話番号を正規化（ハイフン・スペース除去、全角→半角）。"""
    import unicodedata
    s = unicodedata.normalize("NFKC", str(phone or "")).strip()
    s = s.replace("-", "").replace(" ", "").replace("\u3000", "")
    if len(s) == 10 and not s.startswith("0"):
        s = "0" + s
    return s


def get_default_pdf_type(carrier: str, phone: str, docomo_rep: str) -> str:
    """キャリアと電話番号からデフォルトのPDFの種類を返す。"""
    if carrier == "docomo":
        return "一括請求" if (not docomo_rep or phone == docomo_rep) else "利用内訳"
    if carrier in ("au", "UQmobile"):
        return "請求書,支払証明書"
    return "電話番号別"


def get_pdf_types_for_carrier(carrier: str, phone: str, docomo_rep: str) -> list:
    """キャリアと電話番号から選択可能なPDFの種類リストを返す。"""
    if carrier in ("au", "UQmobile"):
        return PDF_TYPES_AU
    if carrier == "docomo":
        return PDF_TYPES_DOCOMO_REP if (not docomo_rep or phone == docomo_rep) else PDF_TYPES_DOCOMO
    return PDF_TYPES_SB


def _parse_target_month_from_settings(sh) -> tuple:
    """設定シートから対象月 (year, month) を取得する。未設定は前月。"""
    import re as _re
    from datetime import datetime, timedelta

    try:
        ws = sh.worksheet("設定")
        for row in ws.get_all_records():
            if str(row.get("設定名", "")).strip() == "対象月":
                raw = row.get("値", "")
                val = str(raw).strip()
                m1 = _re.match(r"^(\d{4})(\d{2})$", val)
                if m1:
                    return int(m1.group(1)), int(m1.group(2))
                m2 = _re.match(r"^(\d{4})年(\d{1,2})月$", val)
                if m2:
                    return int(m2.group(1)), int(m2.group(2))
                # Dateオブジェクト（gspreadがdatetimeとして返す場合）
                if hasattr(raw, "year"):
                    return raw.year, raw.month
                break
    except Exception:
        pass

    # 前月
    today = datetime.today().replace(day=1)
    prev = today - timedelta(days=1)
    return prev.year, prev.month


def _parse_month_sheet_num(name: str) -> int:
    """月別シート名を数値に変換（ソート用）。"""
    import re as _re
    m1 = _re.match(r"\((\d{4})\)(\d+)月", name)
    if m1:
        return int(m1.group(1)) * 100 + int(m1.group(2))
    m2 = _re.search(r"(\d{4})年(\d+)月$", name)
    if m2:
        return int(m2.group(1)) * 100 + int(m2.group(2))
    m3 = _re.search(r"(\d+)月$", name)
    if m3:
        from datetime import datetime
        return datetime.today().year * 100 + int(m3.group(1))
    return 0


def _load_phones_from_month_sheets(ss, target_year: int, target_month: int) -> dict:
    """
    回線管理スプシの月別シートから電話番号を収集する。
    GASの _getAllPhonesFromMonthSheets_() のPython版。
    """
    result = {k: [] for k in ("SoftBank", "Ymobile", "au", "UQmobile", "docomo")}

    # 月別シートを特定（".*\\d+月$" にマッチ）
    import re as _re
    all_sheets = ss.worksheets()
    month_sheets = [ws for ws in all_sheets if _re.search(r"\d+月$", ws.title)]
    if not month_sheets:
        return result

    month_sheets.sort(key=lambda ws: _parse_month_sheet_num(ws.title))

    # 対象月シートを探す
    target_num = target_year * 100 + target_month
    target_sheet = None
    for ws in month_sheets:
        if _parse_month_sheet_num(ws.title) == target_num:
            target_sheet = ws
            break

    # なければ最新のデータあり月にフォールバック
    if not target_sheet:
        for ws in reversed(month_sheets):
            if ws.row_count > 1:
                target_sheet = ws
                break

    if not target_sheet:
        return result

    rows = target_sheet.get_all_values()
    cols = None
    section_carrier = None

    for row in rows:
        # セクションラベル行の判定（非空セルが1〜3個で「電話番号」を含まない）
        non_empty = [c for c in row if str(c).strip()]
        if 0 < len(non_empty) <= 3 and "電話番号" not in [c.strip() for c in non_empty]:
            for cell in non_empty:
                c = _normalize_carrier(cell)
                if c:
                    section_carrier = c
                    cols = None
                    break
            continue

        # ヘッダー行の判定
        stripped = [str(c).strip() for c in row]
        if "電話番号" in stripped:
            cols = {v: i for i, v in enumerate(stripped) if v}
            continue

        if cols is None or "電話番号" not in cols:
            continue

        phone = _normalize_phone(row[cols["電話番号"]])
        if not phone or not _re.match(r"^\d{10,13}$", phone):
            continue

        # キャリア判定
        carrier = None
        if "キャリア" in cols:
            carrier = _normalize_carrier(str(row[cols["キャリア"]]))
        if not carrier:
            carrier = section_carrier
        if not carrier or carrier not in result:
            continue

        cancelled = False
        if "解約済" in cols:
            cancelled = str(row[cols["解約済"]]).strip().upper() == "TRUE"

        device = ""
        if "運用端末" in cols:
            device = str(row[cols["運用端末"]]).strip()

        name = ""
        for key in ("名義", "契約者名"):
            if key in cols:
                name = str(row[cols[key]]).strip()
                break

        login_id = ""
        if "ID" in cols:
            login_id = str(row[cols["ID"]]).strip()

        # 重複チェック
        if not any(p["phone"] == phone for p in result[carrier]):
            result[carrier].append({
                "phone":     phone,
                "cancelled": cancelled,
                "device":    device,
                "name":      name,
                "loginId":   login_id,
            })

    return result


class PhoneManagerLoader(QThread):
    """回線管理スプシ + 認証情報シートから電話番号一覧・選択状態を読み込む。

    GASの getPhoneManagerData() 相当。

    Signal data_loaded の引数:
      {
        "phones": {carrier: [{"phone", "cancelled", "device", "name", "loginId"}]},
        "selections": {carrier: {phone: {"pdfType": str}}},
        "docomo_rep": str,
        "target_month": str,  # "2026年2月" 形式
      }
    """

    data_loaded  = Signal(dict)
    error_signal = Signal(str)

    def __init__(self, script_dir: Path):
        super().__init__()
        self._script_dir = script_dir

    def run(self):
        try:
            import re as _re
            from dotenv import load_dotenv
            from shared_utils import bootstrap_env_from_gsheet, get_gspread_client, open_sheet

            bootstrap_env_from_gsheet(self._script_dir, "GUI")
            load_dotenv(self._script_dir / ".env")

            url = os.environ.get("SPREADSHEET_URL", "").strip()
            m = _re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
            spreadsheet_id = m.group(1) if m else os.environ.get("SPREADSHEET_ID", "").strip()
            if not spreadsheet_id:
                self.error_signal.emit("スプレッドシートIDが見つかりません")
                return

            gc = get_gspread_client()
            sh = open_sheet(gc, spreadsheet_id)

            # 設定シートから各種設定を取得
            docomo_rep = ""
            mgmt_url   = ""
            try:
                ws = sh.worksheet("設定")
                for row in ws.get_all_records():
                    key = str(row.get("設定名", "")).strip()
                    val = str(row.get("値", "")).strip()
                    if key == "docomo代表回線":
                        docomo_rep = _normalize_phone(val)
                    elif key == "回線管理スプレッドシート":
                        mgmt_url = val
            except Exception:
                pass

            # 対象月
            target_year, target_month = _parse_target_month_from_settings(sh)
            target_month_str = f"{target_year}年{target_month}月"

            # 回線管理スプシから電話番号を読み込む
            phones = {k: [] for k in ("SoftBank", "Ymobile", "au", "UQmobile", "docomo")}
            if mgmt_url:
                m2 = _re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", mgmt_url)
                if m2:
                    try:
                        mgmt_sh = gc.open_by_key(m2.group(1))
                        phones = _load_phones_from_month_sheets(mgmt_sh, target_year, target_month)
                    except Exception as e:
                        self.error_signal.emit(f"回線管理スプシの読み込みに失敗: {e}")

            # phones が空なら認証情報シートからフォールバック
            total_phones = sum(len(v) for v in phones.values())
            if total_phones == 0:
                try:
                    ws = sh.worksheet("認証情報")
                    for row in ws.get_all_records():
                        carrier = str(row.get("キャリア", "")).strip()
                        phone   = _normalize_phone(str(row.get("電話番号", "")))
                        if carrier not in phones or not phone:
                            continue
                        status = str(row.get("状態", "契約中")).strip()
                        phones[carrier].append({
                            "phone":     phone,
                            "cancelled": status == "解約済",
                            "device":    str(row.get("運用端末", "")).strip(),
                            "name":      "",
                            "loginId":   str(row.get("ログインID", "")).strip(),
                        })
                except Exception:
                    pass

            # 現在の選択状態（認証情報シート）
            selections = {k: {} for k in ("SoftBank", "Ymobile", "au", "UQmobile", "docomo")}
            try:
                ws = sh.worksheet("認証情報")
                for row in ws.get_all_records():
                    carrier  = str(row.get("キャリア", "")).strip()
                    phone    = _normalize_phone(str(row.get("電話番号", "")))
                    pdf_type = str(row.get("PDFの種類", "")).strip()
                    status   = str(row.get("状態", "契約中")).strip()
                    if carrier in selections and phone and status != "解約済":
                        selections[carrier][phone] = {
                            "pdfType": pdf_type or get_default_pdf_type(carrier, phone, docomo_rep)
                        }
            except Exception:
                pass

            self.data_loaded.emit({
                "phones":       phones,
                "selections":   selections,
                "docomo_rep":   docomo_rep,
                "target_month": target_month_str,
            })

        except Exception as e:
            self.error_signal.emit(f"回線データの読み込みに失敗しました: {e}\n{traceback.format_exc()}")


class PhoneManagerSaver(QThread):
    """選択状態を認証情報シートに保存する。

    GASの savePhoneSelections() 相当。

    Args:
        script_dir:   スクリプトディレクトリ
        phones_data:  {carrier: [{"phone", "cancelled", "device", "name", "loginId"}]}
        selections:   {carrier: {phone: {"pdfType": str}}}  ← checkedなphonesのみ
        docomo_rep:   docomo代表回線電話番号
    """

    saved        = Signal(str)
    error_signal = Signal(str)

    def __init__(self, script_dir: Path, phones_data: dict,
                 selections: dict, docomo_rep: str):
        super().__init__()
        self._script_dir  = script_dir
        self._phones_data = phones_data
        self._selections  = selections
        self._docomo_rep  = docomo_rep

    def run(self):
        try:
            import re as _re
            from dotenv import load_dotenv
            from shared_utils import bootstrap_env_from_gsheet, get_gspread_client, open_sheet

            bootstrap_env_from_gsheet(self._script_dir, "GUI")
            load_dotenv(self._script_dir / ".env")

            url = os.environ.get("SPREADSHEET_URL", "").strip()
            m = _re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
            spreadsheet_id = m.group(1) if m else os.environ.get("SPREADSHEET_ID", "").strip()
            if not spreadsheet_id:
                self.error_signal.emit("スプレッドシートIDが見つかりません")
                return

            gc = get_gspread_client()
            sh = open_sheet(gc, spreadsheet_id)

            # device / loginId マップを構築
            device_map    = {}
            login_id_map  = {}
            cancelled_set = set()
            ALL_CARRIERS  = ["SoftBank", "Ymobile", "au", "UQmobile", "docomo"]

            for carrier in ALL_CARRIERS:
                for p in self._phones_data.get(carrier, []):
                    phone = p["phone"]
                    if p.get("cancelled"):
                        cancelled_set.add(phone)
                    if p.get("device"):
                        device_map[phone] = p["device"]
                    if p.get("loginId"):
                        login_id_map[phone] = p["loginId"]

            # 行データを構築
            active_rows    = []
            cancelled_rows = []

            for carrier in ALL_CARRIERS:
                sel = self._selections.get(carrier, {})
                for phone, info in sel.items():
                    pdf_type = info.get("pdfType") or get_default_pdf_type(
                        carrier, phone, self._docomo_rep
                    )
                    status = "解約済" if phone in cancelled_set else "契約中"
                    row = [
                        phone,
                        carrier,
                        pdf_type,
                        device_map.get(phone, ""),
                        status,
                        login_id_map.get(phone, ""),
                    ]
                    if status == "解約済":
                        cancelled_rows.append(row)
                    else:
                        active_rows.append(row)

            all_rows = active_rows + cancelled_rows

            # 認証情報シートに書き込む
            ws = sh.worksheet("認証情報")

            # 既存データをクリア（ヘッダー行以外）
            existing_rows = ws.row_count
            if existing_rows > 1:
                ws.batch_clear([f"A2:F{max(existing_rows, len(all_rows) + 10)}"])

            # 新データを書き込む
            if all_rows:
                ws.update(f"A2:F{len(all_rows) + 1}", all_rows, value_input_option="RAW")

            n_active    = len(active_rows)
            n_cancelled = len(cancelled_rows)
            msg = f"保存しました（契約中 {n_active} 件"
            if n_cancelled:
                msg += f"、解約済 {n_cancelled} 件"
            msg += f"、合計 {len(all_rows)} 行）"
            self.saved.emit(msg)

        except Exception as e:
            self.error_signal.emit(f"保存に失敗しました: {e}\n{traceback.format_exc()}")


class MonthSaver(QThread):
    """対象月をスプレッドシートの設定シートに保存する。

    Signal:
        saved():        保存完了
        error_signal(str): エラーメッセージ
    """

    saved        = Signal()
    error_signal = Signal(str)

    def __init__(self, script_dir: Path, month_str: str):
        super().__init__()
        self._script_dir = script_dir
        self._month_str  = month_str   # "YYYY年MM月"

    def run(self):
        try:
            import os
            from dotenv import load_dotenv
            from shared_utils import (
                bootstrap_env_from_gsheet, get_gspread_client, open_sheet,
            )

            bootstrap_env_from_gsheet(self._script_dir, "GUI")
            load_dotenv(self._script_dir / ".env")

            url = os.environ.get("SPREADSHEET_URL", "").strip()
            m   = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
            spreadsheet_id = (
                m.group(1) if m else os.environ.get("SPREADSHEET_ID", "").strip()
            )

            if not spreadsheet_id:
                self.error_signal.emit("スプレッドシートIDが見つかりません")
                return

            gc = get_gspread_client()
            sh = open_sheet(gc, spreadsheet_id)
            ws = sh.worksheet("設定")

            # 「対象月」行を探して値を更新
            records = ws.get_all_records()
            for i, row in enumerate(records, start=2):   # ヘッダー = 1行目
                if str(row.get("設定名", "")).strip() == "対象月":
                    ws.update_cell(i, 2, self._month_str)
                    self.saved.emit()
                    return

            self.error_signal.emit("設定シートに「対象月」行が見つかりません")

        except Exception as e:
            self.error_signal.emit(f"対象月の保存に失敗しました: {e}")
