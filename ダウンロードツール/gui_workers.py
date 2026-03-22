"""バックグラウンドワーカーとログキャプチャ

DownloadWorker: キャリアごとのPDFダウンロードをQThreadで実行
UpdateAmountsWorker: 金額更新をQThreadで実行
SpreadsheetReader: スプレッドシートから設定・履歴を読み込む
"""

import io
import logging
import os
import re
import sys
import traceback
from pathlib import Path

from PySide6.QtCore import QThread, Signal


class StdoutRedirector(io.TextIOBase):
    """sys.stdoutをリダイレクトしてSignalを発火する。
    SMS認証プロンプトの検出も行う。
    """

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
        """ask_security_code() の print() 出力からSMS認証情報を抽出する。

        検出する出力パターン:
          📱 SMS認証が必要です
          電話番号: 09012345678
          端末    : iPhoneAir
          SMSに届いた3桁のセキュリティ番号を入力してください
          echo '123' > /tmp/softbank_security_code.txt
        """
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
    """logging の出力を Signal に変換するハンドラ。
    DownloadWorker が root logger に追加して使用する。
    """

    def __init__(self, signal):
        super().__init__()
        self._signal = signal

    def emit(self, record):
        msg = self.format(record)
        self._signal.emit(msg + "\n")


class DownloadWorker(QThread):
    """ダウンロード処理をバックグラウンドで実行するワーカー。

    shared_utils.run_main() を QThread 内で呼び出す。
    stdout/stderr をリダイレクトし、ログとSMS認証プロンプトをSignalで通知する。
    """

    log_signal = Signal(str)
    sms_signal = Signal(dict)       # {phone, device, code_file, digits}
    finished_signal = Signal(list)  # [(carrier, phone, result), ...]
    error_signal = Signal(str)
    progress_signal = Signal(str)

    def __init__(self, carriers, script_dir, env_overrides=None):
        super().__init__()
        self._carriers = carriers
        self._script_dir = script_dir
        self._env_overrides = env_overrides or {}

    def run(self):
        # 環境変数を設定（終了時に復元）
        saved_env = {}
        for k, v in self._env_overrides.items():
            saved_env[k] = os.environ.get(k)
            if v is not None:
                os.environ[k] = v
            elif k in os.environ:
                del os.environ[k]

        # ログハンドラを設定（create_billing_context の basicConfig より先に）
        log_handler = SignalLogHandler(self.log_signal)
        log_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
        )
        root_logger = logging.getLogger()
        root_logger.addHandler(log_handler)
        root_logger.setLevel(logging.INFO)

        # stdout/stdin をリダイレクト
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        old_stdin = sys.stdin

        def sms_cb(info):
            self.sms_signal.emit(info)

        sys.stdout = StdoutRedirector(self.log_signal, sms_cb)
        sys.stderr = StdoutRedirector(self.log_signal)
        sys.stdin = io.StringIO()  # 非インタラクティブ → ファイル待機モード

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
            sys.stdin = old_stdin
            root_logger.removeHandler(log_handler)
            for k, v in saved_env.items():
                if v is not None:
                    os.environ[k] = v
                elif k in os.environ:
                    del os.environ[k]


class UpdateAmountsWorker(QThread):
    """金額更新（download.py --update-amounts 相当）をバックグラウンドで実行する。"""

    log_signal = Signal(str)
    finished_signal = Signal()
    error_signal = Signal(str)

    def __init__(self, script_dir):
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
    """スプレッドシートから設定・履歴を読み込む。
    GUI起動時と「更新」ボタン押下時に実行。
    """

    settings_loaded = Signal(list)      # [(key, value), ...]
    history_loaded = Signal(list)       # [dict, ...]
    target_month_loaded = Signal(str)   # "2026年02月" 等
    error_signal = Signal(str)

    def __init__(self, script_dir):
        super().__init__()
        self._script_dir = script_dir

    def run(self):
        try:
            from dotenv import load_dotenv
            from shared_utils import (
                bootstrap_env_from_gsheet, get_gspread_client, open_sheet,
            )

            bootstrap_env_from_gsheet(self._script_dir, "GUI")
            load_dotenv(self._script_dir / ".env")

            # スプレッドシートIDを取得
            url = os.environ.get("SPREADSHEET_URL", "").strip()
            m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
            spreadsheet_id = m.group(1) if m else os.environ.get("SPREADSHEET_ID", "").strip()

            if not spreadsheet_id:
                self.error_signal.emit("スプレッドシートIDが見つかりません")
                return

            gc = get_gspread_client()
            sh = open_sheet(gc, spreadsheet_id)

            # 設定シート読み込み
            try:
                ws = sh.worksheet("設定")
                settings = []
                for row in ws.get_all_records():
                    name = str(row.get("設定名", "")).strip()
                    value = str(row.get("値", "")).strip()
                    if name:
                        settings.append((name, value))
                        if name == "対象月":
                            self.target_month_loaded.emit(value)
                self.settings_loaded.emit(settings)
            except Exception as e:
                self.error_signal.emit(f"設定シート読み込みエラー: {e}")

            # ダウンロード履歴シート読み込み
            try:
                ws = sh.worksheet("ダウンロード履歴")
                rows = ws.get_all_records()
                self.history_loaded.emit(rows)
            except Exception as e:
                self.error_signal.emit(f"履歴シート読み込みエラー: {e}")

        except Exception as e:
            self.error_signal.emit(f"スプレッドシート接続エラー: {e}")
