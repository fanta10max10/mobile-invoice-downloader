#!/usr/bin/env python3
"""携帯領収書管理 GUIアプリ

既存の download.py / shared_utils.py をGUIから実行するデスクトップアプリ。
起動方法: python3 gui_app.py
"""

import os
import sys
from pathlib import Path

# スクリプトディレクトリをPythonパスに追加（shared_utils, download のインポート用）
SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

try:
    from PySide6.QtCore import Qt, QTimer
    from PySide6.QtGui import QColor, QFont, QTextCharFormat, QTextCursor
    from PySide6.QtWidgets import (
        QApplication, QCheckBox, QGroupBox, QHBoxLayout, QHeaderView,
        QLabel, QLineEdit, QMainWindow, QPushButton, QTabWidget,
        QTableWidget, QTableWidgetItem, QTextEdit, QVBoxLayout, QWidget,
    )
except ImportError:
    print("PySide6 がインストールされていません。")
    print("  pip install PySide6")
    sys.exit(1)

from download import ALL_CARRIERS
from gui_sms_dialog import SmsCodeDialog
from gui_workers import DownloadWorker, SpreadsheetReader, UpdateAmountsWorker

# キャリアアイコン（要件定義書より）
CARRIER_ICONS = {
    "SoftBank": "🐶",
    "Ymobile": "🐱",
    "au": "🍊",
    "UQmobile": "👸",
    "docomo": "🍄",
}


class MainWindow(QMainWindow):
    """メインウィンドウ"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("携帯領収書管理")
        self.setMinimumSize(750, 620)
        self._worker = None
        self._update_worker = None
        self._reader = None

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(8)

        # ── 上部: 対象月 + ステータス ──
        top = QHBoxLayout()
        top.addWidget(QLabel("対象月:"))
        self._month_label = QLabel("読み込み中...")
        self._month_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        top.addWidget(self._month_label)
        top.addStretch()
        self._status_label = QLabel("⏳ 待機中")
        self._status_label.setStyleSheet("font-size: 14px;")
        top.addWidget(self._status_label)
        main_layout.addLayout(top)

        # ── オプション ──
        opts = QGroupBox("オプション")
        opts_layout = QHBoxLayout(opts)
        self._headless_cb = QCheckBox("ブラウザ表示 (デバッグ用)")
        self._dryrun_cb = QCheckBox("テストモード (DRY_RUN)")
        opts_layout.addWidget(self._headless_cb)
        opts_layout.addWidget(self._dryrun_cb)
        opts_layout.addWidget(QLabel("リトライ番号:"))
        self._retry_input = QLineEdit()
        self._retry_input.setPlaceholderText("09012345678,08012345678")
        self._retry_input.setMaximumWidth(250)
        opts_layout.addWidget(self._retry_input)
        main_layout.addWidget(opts)

        # ── キャリアボタン ──
        dl_group = QGroupBox("ダウンロード")
        dl_layout = QVBoxLayout(dl_group)

        # 個別キャリアボタン
        btn_row = QHBoxLayout()
        self._carrier_buttons = []
        for config in ALL_CARRIERS:
            icon = CARRIER_ICONS.get(config.carrier_name, "📱")
            btn = QPushButton(f"{icon} {config.carrier_name}")
            btn.setMinimumHeight(40)
            btn.clicked.connect(lambda checked, c=config: self._start_download([c]))
            btn_row.addWidget(btn)
            self._carrier_buttons.append(btn)
        dl_layout.addLayout(btn_row)

        # アクションボタン
        action_row = QHBoxLayout()
        self._all_btn = QPushButton("📥 全キャリア一括実行")
        self._all_btn.setMinimumHeight(45)
        self._all_btn.setStyleSheet("font-weight: bold;")
        self._all_btn.clicked.connect(
            lambda: self._start_download(list(ALL_CARRIERS))
        )
        action_row.addWidget(self._all_btn)

        self._amount_btn = QPushButton("💰 金額更新")
        self._amount_btn.setMinimumHeight(45)
        self._amount_btn.clicked.connect(self._start_update_amounts)
        action_row.addWidget(self._amount_btn)
        dl_layout.addLayout(action_row)

        main_layout.addWidget(dl_group)

        # ── タブ: ログ / 設定 / ダウンロード履歴 ──
        self._tabs = QTabWidget()

        # ログタブ
        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(0, 4, 0, 0)
        clear_btn = QPushButton("🗑 ログクリア")
        clear_btn.setMaximumWidth(120)
        clear_btn.clicked.connect(lambda: self._log_text.clear())
        log_layout.addWidget(clear_btn, alignment=Qt.AlignRight)
        self._log_text = QTextEdit()
        self._log_text.setReadOnly(True)
        self._log_text.setFont(QFont("Menlo", 11))
        log_layout.addWidget(self._log_text)
        self._tabs.addTab(log_widget, "ログ")

        # 設定タブ
        settings_widget = QWidget()
        settings_layout = QVBoxLayout(settings_widget)
        settings_layout.setContentsMargins(0, 4, 0, 0)
        refresh_btn = QPushButton("🔄 スプレッドシートから再読み込み")
        refresh_btn.clicked.connect(self._load_spreadsheet)
        settings_layout.addWidget(refresh_btn, alignment=Qt.AlignLeft)
        self._settings_table = QTableWidget(0, 2)
        self._settings_table.setHorizontalHeaderLabels(["設定名", "値"])
        self._settings_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch
        )
        settings_layout.addWidget(self._settings_table)
        self._tabs.addTab(settings_widget, "設定")

        # ダウンロード履歴タブ
        history_widget = QWidget()
        history_layout = QVBoxLayout(history_widget)
        history_layout.setContentsMargins(0, 4, 0, 0)
        refresh_btn2 = QPushButton("🔄 スプレッドシートから再読み込み")
        refresh_btn2.clicked.connect(self._load_spreadsheet)
        history_layout.addWidget(refresh_btn2, alignment=Qt.AlignLeft)
        self._history_table = QTableWidget(0, 0)
        history_layout.addWidget(self._history_table)
        self._tabs.addTab(history_widget, "ダウンロード履歴")

        main_layout.addWidget(self._tabs, 1)  # stretch=1 で残り領域をすべて使う

        # 起動時にスプレッドシートをバックグラウンド読み込み
        QTimer.singleShot(500, self._load_spreadsheet)

    # ─── オプション → 環境変数 ───

    def _get_env_overrides(self) -> dict:
        """GUIのオプション設定を環境変数のオーバーライド辞書に変換する。"""
        overrides = {}
        # チェック ON = ブラウザ表示 = HEADLESS=false
        overrides["HEADLESS"] = "false" if self._headless_cb.isChecked() else "true"
        overrides["DRY_RUN"] = "true" if self._dryrun_cb.isChecked() else None
        retry = self._retry_input.text().strip()
        overrides["RETRY_PHONES"] = retry if retry else None
        return overrides

    # ─── ボタン制御 ───

    def _set_running(self, running: bool):
        """実行中はボタンを無効化し、ログタブに切り替える。"""
        for btn in self._carrier_buttons:
            btn.setEnabled(not running)
        self._all_btn.setEnabled(not running)
        self._amount_btn.setEnabled(not running)
        if running:
            self._status_label.setText("🔄 実行中...")
            self._tabs.setCurrentIndex(0)

    # ─── ダウンロード実行 ───

    def _start_download(self, carriers):
        if self._worker and self._worker.isRunning():
            return
        self._set_running(True)
        self._log_text.clear()

        self._worker = DownloadWorker(
            carriers, SCRIPT_DIR, self._get_env_overrides()
        )
        self._worker.log_signal.connect(self._on_log)
        self._worker.sms_signal.connect(self._on_sms_needed)
        self._worker.progress_signal.connect(self._on_progress)
        self._worker.finished_signal.connect(self._on_download_finished)
        self._worker.error_signal.connect(self._on_error)
        self._worker.start()

    def _start_update_amounts(self):
        if (self._worker and self._worker.isRunning()) or \
           (self._update_worker and self._update_worker.isRunning()):
            return
        self._set_running(True)
        self._log_text.clear()

        self._update_worker = UpdateAmountsWorker(SCRIPT_DIR)
        self._update_worker.log_signal.connect(self._on_log)
        self._update_worker.finished_signal.connect(
            lambda: self._on_download_finished([])
        )
        self._update_worker.error_signal.connect(self._on_error)
        self._update_worker.start()

    # ─── ログ表示 ───

    def _on_log(self, text):
        cursor = self._log_text.textCursor()
        cursor.movePosition(QTextCursor.End)

        fmt = QTextCharFormat()
        if "[ERROR]" in text or "❌" in text:
            fmt.setForeground(QColor("#FF4444"))
        elif "[WARNING]" in text or "⚠" in text:
            fmt.setForeground(QColor("#CC8800"))
        elif "✅" in text or "成功" in text:
            fmt.setForeground(QColor("#44AA44"))
        else:
            fmt.setForeground(self.palette().text().color())

        cursor.setCharFormat(fmt)
        cursor.insertText(text)
        self._log_text.setTextCursor(cursor)
        self._log_text.ensureCursorVisible()

    # ─── SMS認証 ───

    def _on_sms_needed(self, info):
        """SMS認証プロンプトを検出したらダイアログを表示する。"""
        timeout = int(os.environ.get("SECURITY_CODE_TIMEOUT", "60"))
        dialog = SmsCodeDialog(
            phone=info["phone"],
            device=info.get("device", ""),
            code_file=info["code_file"],
            digits=info.get("digits", 3),
            timeout=timeout,
            parent=self,
        )
        dialog.exec()

    # ─── 完了・エラー ───

    def _on_progress(self, text):
        self._status_label.setText(f"🔄 {text}")

    def _on_download_finished(self, results):
        self._set_running(False)
        if results:
            n_success = sum(1 for *_, r in results if r == "success")
            n_skipped = sum(1 for *_, r in results if r == "skipped")
            n_failed = sum(1 for *_, r in results if r == "failed")
            self._status_label.setText(
                f"✅ 完了: {n_success}件成功 / {n_skipped}件スキップ / {n_failed}件失敗"
            )
        else:
            self._status_label.setText("✅ 完了")

    def _on_error(self, error):
        self._set_running(False)
        self._status_label.setText("❌ エラー")
        self._on_log(f"\n❌ エラー: {error}\n")

    # ─── スプレッドシート読み込み ───

    def _load_spreadsheet(self):
        """スプレッドシートから設定・履歴をバックグラウンドで読み込む。"""
        if self._reader and self._reader.isRunning():
            return
        self._reader = SpreadsheetReader(SCRIPT_DIR)
        self._reader.settings_loaded.connect(self._on_settings_loaded)
        self._reader.history_loaded.connect(self._on_history_loaded)
        self._reader.target_month_loaded.connect(self._on_target_month)
        self._reader.error_signal.connect(lambda e: self._on_log(f"⚠️ {e}\n"))
        self._reader.start()

    def _on_target_month(self, month_str):
        self._month_label.setText(month_str if month_str else "自動（前月）")

    def _on_settings_loaded(self, settings):
        self._settings_table.setRowCount(len(settings))
        for i, (key, value) in enumerate(settings):
            self._settings_table.setItem(i, 0, QTableWidgetItem(key))
            self._settings_table.setItem(i, 1, QTableWidgetItem(value))

    def _on_history_loaded(self, rows):
        if not rows:
            self._history_table.setRowCount(0)
            return
        headers = list(rows[0].keys())
        self._history_table.setColumnCount(len(headers))
        self._history_table.setHorizontalHeaderLabels(headers)
        # 最新を上に表示
        rows_reversed = list(reversed(rows))
        self._history_table.setRowCount(len(rows_reversed))
        for i, row in enumerate(rows_reversed):
            for j, key in enumerate(headers):
                self._history_table.setItem(
                    i, j, QTableWidgetItem(str(row.get(key, "")))
                )
        self._history_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeToContents
        )


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("携帯領収書管理")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
