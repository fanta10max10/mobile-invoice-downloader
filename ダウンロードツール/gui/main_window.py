"""メインウィンドウ"""

import os
import socket
from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QGraphicsDropShadowEffect, QHBoxLayout, QLabel, QMainWindow,
    QPushButton, QStatusBar, QVBoxLayout, QWidget,
)

from .styles import (
    ACCENT, AMBER, BG_CARD, BG_DEEPER, BG_MAIN, BG_SURFACE, BG_HOVER,
    BORDER, BORDER_LIGHT, GREEN, RED, TEXT, TEXT_MUTED, TEXT_FAINT, TEXT_DIM,
    ACCENT_SOFT, ACCENT_CHIP,
    ToggleSwitch, MonthSelector,
)
from .carrier_tabs import CarrierTabs
from .log_view     import LogView
from .settings_view import SettingsView
from .history_view  import HistoryView
from .sms_dialog    import SmsCodeDialog
from .workers       import (
    DownloadWorker, PhoneListLoader, PhoneManagerLoader, PhoneManagerSaver,
    MonthSaver, SpreadsheetReader, UpdateAmountsWorker,
)

VERSION = "1.0.0"


class MainWindow(QMainWindow):
    """携帯領収書管理メインウィンドウ。"""

    def __init__(self, script_dir: Path):
        super().__init__()
        self._script_dir    = script_dir
        self._worker        = None
        self._update_worker = None
        self._reader        = None
        self._phone_loader  = None
        self._saver         = None
        self._month_saver   = None

        self.setWindowTitle("携帯領収書管理")
        self.setMinimumSize(840, 640)
        self.resize(1000, 740)

        # ── 中央ウィジェット ──
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # ── ヘッダー ──
        root_layout.addWidget(self._build_header())

        # ── メインコンテンツ（オーロラ背景）──
        content_wrapper = QWidget()
        content_wrapper.setStyleSheet(f"""
            QWidget {{
                background: qradialgradient(
                    cx:0.75, cy:0.0, radius:0.9,
                    fx:0.75, fy:0.0,
                    stop:0   rgba(124, 58, 237, 18),
                    stop:0.4 rgba(99, 102, 241, 8),
                    stop:1   rgba(8, 8, 18, 255)
                );
            }}
        """)
        cw_layout = QVBoxLayout(content_wrapper)
        cw_layout.setContentsMargins(20, 16, 20, 16)
        cw_layout.setSpacing(14)

        # キャリアタブ（カードシャドウ）
        self._carrier_tabs = CarrierTabs()
        card_shadow = QGraphicsDropShadowEffect()
        card_shadow.setBlurRadius(20)
        card_shadow.setOffset(0, 6)
        card_shadow.setColor(QColor(0, 0, 0, 100))
        self._carrier_tabs.setGraphicsEffect(card_shadow)
        self._carrier_tabs.run_requested.connect(self._start_download)
        self._carrier_tabs.save_requested.connect(self._on_save_requested)
        self._carrier_tabs.save_and_run_requested.connect(self._on_save_and_run_requested)
        cw_layout.addWidget(self._carrier_tabs)

        # アクションボタン行
        cw_layout.addLayout(self._build_action_buttons())

        # コンテンツタブバー + コンテンツ
        cw_layout.addLayout(self._build_content_tabs())

        root_layout.addWidget(content_wrapper, 1)

        # ── ステータスバー ──
        self._setup_status_bar()

        # ── 起動時の読み込み ──
        QTimer.singleShot(300, self._load_all)

    # ─────────────────────────────── ヘッダー ───────────────────────────────

    def _build_header(self) -> QWidget:
        header = QWidget()
        header.setFixedHeight(64)
        header.setStyleSheet(f"""
            QWidget {{
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 {BG_SURFACE}, stop:1 #0c0c1e
                );
                border-bottom: 1px solid rgba(139, 92, 246, 0.15);
            }}
        """)
        layout = QHBoxLayout(header)
        layout.setContentsMargins(24, 0, 24, 0)
        layout.setSpacing(0)

        # ── 左: アプリタイトル ──
        title_col = QVBoxLayout()
        title_col.setSpacing(2)
        title_col.setAlignment(Qt.AlignVCenter)

        title_lbl = QLabel("携帯領収書管理")
        title_lbl.setStyleSheet(f"""
            color: {TEXT};
            font-size: 15px;
            font-weight: 700;
            letter-spacing: 0.5px;
            background: transparent;
        """)
        subtitle_lbl = QLabel("Mobile Invoice Manager")
        subtitle_lbl.setStyleSheet(f"""
            color: rgba(139, 92, 246, 0.75);
            font-size: 10px;
            letter-spacing: 1.5px;
            font-weight: 500;
            background: transparent;
        """)
        title_col.addWidget(title_lbl)
        title_col.addWidget(subtitle_lbl)
        layout.addLayout(title_col)

        layout.addStretch()

        # ── 中央: 対象月セレクター ──
        month_col = QVBoxLayout()
        month_col.setSpacing(4)
        month_col.setAlignment(Qt.AlignVCenter)

        month_title = QLabel("対象月")
        month_title.setStyleSheet(f"""
            color: {TEXT_MUTED};
            font-size: 10px;
            letter-spacing: 1px;
            font-weight: 500;
            background: transparent;
        """)
        month_title.setAlignment(Qt.AlignCenter)

        self._month_selector = MonthSelector()
        self._month_selector.month_changed.connect(self._on_month_changed)

        month_col.addWidget(month_title)
        month_col.addWidget(self._month_selector)
        layout.addLayout(month_col)

        layout.addStretch()

        # ── 右: トグルスイッチ群 ──
        toggles_row = QHBoxLayout()
        toggles_row.setSpacing(20)
        toggles_row.setAlignment(Qt.AlignVCenter)

        for label, attr in [
            ("デバッグ表示", "_debug_toggle"),
            ("DRY RUN",    "_dryrun_toggle"),
        ]:
            col = QVBoxLayout()
            col.setSpacing(4)
            col.setAlignment(Qt.AlignCenter)

            toggle = ToggleSwitch()
            setattr(self, attr, toggle)
            col.addWidget(toggle, alignment=Qt.AlignCenter)

            lbl = QLabel(label)
            lbl.setStyleSheet(f"""
                color: {TEXT_MUTED};
                font-size: 10px;
                letter-spacing: 0.5px;
                background: transparent;
            """)
            col.addWidget(lbl, alignment=Qt.AlignCenter)

            toggles_row.addLayout(col)

        layout.addLayout(toggles_row)

        return header

    # ──────────────────────────── アクションボタン ────────────────────────────

    def _build_action_buttons(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(10)

        self._all_btn = QPushButton("🚀  全キャリア一括実行")
        self._all_btn.setObjectName("accentBtn")
        self._all_btn.clicked.connect(self._start_all)
        # バイオレットグロウシャドウ
        glow = QGraphicsDropShadowEffect()
        glow.setBlurRadius(28)
        glow.setOffset(0, 4)
        glow.setColor(QColor(124, 58, 237, 90))
        self._all_btn.setGraphicsEffect(glow)
        row.addWidget(self._all_btn)

        self._amount_btn = QPushButton("💰  金額更新")
        self._amount_btn.setMinimumHeight(48)
        self._amount_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                color: {TEXT_MUTED};
                border: 1px solid rgba(245, 158, 11, 0.25);
                border-radius: 12px;
                font-size: 13px;
                font-weight: 500;
                min-height: 48px;
                padding: 0 22px;
            }}
            QPushButton:hover {{
                background: rgba(245, 158, 11, 0.08);
                color: #fcd34d;
                border-color: rgba(245, 158, 11, 0.45);
            }}
            QPushButton:disabled {{
                color: {TEXT_FAINT};
                border-color: {BORDER};
            }}
        """)
        self._amount_btn.clicked.connect(self._start_update_amounts)
        row.addWidget(self._amount_btn)

        return row

    # ────────────────────────── コンテンツタブ ──────────────────────────────

    def _build_content_tabs(self) -> QVBoxLayout:
        outer = QVBoxLayout()
        outer.setSpacing(0)

        # タブバー（下線スタイル）
        tab_bar = QWidget()
        tab_bar.setStyleSheet(f"""
            QWidget {{
                background: {BG_MAIN};
                border-bottom: 1px solid {BORDER};
            }}
        """)
        tab_bar_layout = QHBoxLayout(tab_bar)
        tab_bar_layout.setContentsMargins(0, 0, 0, 0)
        tab_bar_layout.setSpacing(0)

        self._content_tab_btns: list[QPushButton] = []
        self._content_pages: list[QWidget] = []

        # ── ログビュー ──
        self._log_view = LogView()

        # ── 設定ビュー ──
        self._settings_view = SettingsView()
        self._settings_view.reload_btn.clicked.connect(self._load_spreadsheet)

        # ── 履歴ビュー ──
        self._history_view = HistoryView()
        self._history_view.reload_btn.clicked.connect(self._load_spreadsheet)

        pages = [
            ("📋  ログ",            self._log_view),
            ("⚙️   設定",           self._settings_view),
            ("🕒  ダウンロード履歴", self._history_view),
        ]

        for i, (label, page) in enumerate(pages):
            btn = QPushButton(label)
            btn.setObjectName("contentTabBtn")
            btn.clicked.connect(lambda _, idx=i: self._switch_content_tab(idx))
            tab_bar_layout.addWidget(btn)
            self._content_tab_btns.append(btn)

            page.setVisible(i == 0)
            self._content_pages.append(page)

        tab_bar_layout.addStretch()
        outer.addWidget(tab_bar)

        for page in self._content_pages:
            outer.addWidget(page, 1)

        self._switch_content_tab(0)
        return outer

    def _switch_content_tab(self, idx: int):
        for i, (btn, page) in enumerate(
            zip(self._content_tab_btns, self._content_pages)
        ):
            active = (i == idx)
            btn.setProperty("active", "true" if active else "false")
            btn.style().unpolish(btn)
            btn.style().polish(btn)
            page.setVisible(active)

    # ──────────────────────────── ステータスバー ────────────────────────────

    def _setup_status_bar(self):
        sb = QStatusBar()
        self.setStatusBar(sb)

        self._conn_label = QLabel("● 接続確認中")
        self._conn_label.setStyleSheet(
            f"color: {AMBER}; font-family: 'Menlo'; font-size: 11px;"
        )
        sb.addWidget(self._conn_label)

        hostname = socket.gethostname().split(".")[0]
        self._sb_info = QLabel(f"— · {hostname}")
        self._sb_info.setStyleSheet(
            f"color: {TEXT_FAINT}; font-family: 'Menlo'; font-size: 11px;"
        )
        sb.addWidget(self._sb_info)

        sb.addPermanentWidget(QLabel(f"v{VERSION}"))

    # ──────────────────────────── データ読み込み ────────────────────────────

    def _load_all(self):
        self._load_phones()
        self._load_spreadsheet()

    def _load_phones(self):
        if self._phone_loader and self._phone_loader.isRunning():
            return
        self._phone_loader = PhoneManagerLoader(self._script_dir)
        self._phone_loader.data_loaded.connect(self._on_phone_manager_data)
        self._phone_loader.error_signal.connect(
            lambda e: self._log_view.append(f"⚠️ 電話番号読み込みエラー: {e}\n")
        )
        self._phone_loader.start()

    def _on_phone_manager_data(self, data: dict):
        """PhoneManagerLoader のデータ受信。キャリアタブと月セレクターに反映。"""
        self._carrier_tabs.load_data(data)
        target_month = data.get("target_month", "")
        if target_month:
            self._month_selector.set_month(target_month)
            self._update_status_month(target_month)

    def _load_spreadsheet(self):
        if self._reader and self._reader.isRunning():
            return
        self._reader = SpreadsheetReader(self._script_dir)
        self._reader.settings_loaded.connect(self._settings_view.load_settings)
        self._reader.history_loaded.connect(self._history_view.load_history)
        self._reader.target_month_loaded.connect(self._on_target_month)
        self._reader.error_signal.connect(
            lambda e: self._log_view.append(f"⚠️ {e}\n")
        )
        self._reader.finished.connect(self._on_spreadsheet_loaded)
        self._reader.start()

    def _on_target_month(self, month_str: str):
        """SpreadsheetReader から対象月を受け取る。"""
        if month_str:
            self._month_selector.set_month(month_str)
            self._update_status_month(month_str)

    def _update_status_month(self, month_str: str):
        hostname = socket.gethostname().split(".")[0]
        self._sb_info.setText(f"· {month_str} · {hostname}")

    def _on_spreadsheet_loaded(self):
        self._conn_label.setText("● 接続 OK")
        self._conn_label.setStyleSheet(
            f"color: {GREEN}; font-family: 'Menlo'; font-size: 11px;"
        )

    # ──────────────────────── 対象月の変更・保存 ─────────────────────────────

    def _on_month_changed(self, month_str: str):
        """月セレクターが変更されたら設定シートに保存してデータ再読み込み。"""
        if self._month_saver and self._month_saver.isRunning():
            return
        self._month_saver = MonthSaver(self._script_dir, month_str)
        self._month_saver.saved.connect(
            lambda: (
                self.statusBar().showMessage(f"📅 対象月を {month_str} に変更しました", 5000),
                QTimer.singleShot(200, self._load_phones),
            )
        )
        self._month_saver.error_signal.connect(
            lambda e: self._log_view.append(f"\n⚠️ 対象月保存エラー: {e}\n")
        )
        self._month_saver.start()
        self._update_status_month(month_str)

    # ──────────────────────────── ダウンロード実行 ───────────────────────────

    def _get_env_overrides(self) -> dict:
        return {
            "HEADLESS": "false" if self._debug_toggle.isChecked() else "true",
            "DRY_RUN":  "true"  if self._dryrun_toggle.isChecked() else None,
        }

    def _start_download(self, carriers: list, selected_phones: list):
        if self._worker and self._worker.isRunning():
            return
        self._set_running(True)
        self._log_view.clear()
        self._switch_content_tab(0)

        self._worker = DownloadWorker(
            carriers, self._script_dir,
            self._get_env_overrides(), selected_phones
        )
        self._worker.log_signal.connect(self._log_view.append)
        self._worker.sms_signal.connect(self._on_sms_needed)
        self._worker.progress_signal.connect(self._on_progress)
        self._worker.finished_signal.connect(self._on_download_finished)
        self._worker.error_signal.connect(self._on_error)
        self._worker.start()

    def _start_all(self):
        carriers, selected = self._carrier_tabs.get_all_selected()
        self._start_download(carriers, selected)

    def _start_update_amounts(self):
        if (self._worker and self._worker.isRunning()) or \
           (self._update_worker and self._update_worker.isRunning()):
            return
        self._set_running(True)
        self._log_view.clear()
        self._switch_content_tab(0)

        self._update_worker = UpdateAmountsWorker(self._script_dir)
        self._update_worker.log_signal.connect(self._log_view.append)
        self._update_worker.finished_signal.connect(
            lambda: self._on_download_finished([])
        )
        self._update_worker.error_signal.connect(self._on_error)
        self._update_worker.start()

    def _set_running(self, running: bool):
        self._all_btn.setEnabled(not running)
        self._amount_btn.setEnabled(not running)
        self._carrier_tabs.set_enabled(not running)

    # ──────────────────────────── SMS認証 ───────────────────────────────────

    def _on_sms_needed(self, info: dict):
        timeout = int(os.environ.get("SECURITY_CODE_TIMEOUT", "60"))
        dlg = SmsCodeDialog(
            phone=info["phone"],
            device=info.get("device", ""),
            code_file=info["code_file"],
            digits=info.get("digits", 3),
            timeout=timeout,
            parent=self,
        )
        dlg.exec()

    # ──────────────────────────── 完了・エラー ───────────────────────────────

    def _on_progress(self, text: str):
        self.statusBar().showMessage(f"🔄  {text}", 0)

    def _on_download_finished(self, results: list):
        self._set_running(True)
        self._set_running(False)
        if results:
            n_ok   = sum(1 for *_, r in results if r == "success")
            n_skip = sum(1 for *_, r in results if r == "skipped")
            n_fail = sum(1 for *_, r in results if r == "failed")
            msg    = f"✅  完了: {n_ok} 件成功 / {n_skip} 件スキップ / {n_fail} 件失敗"
        else:
            msg = "✅  完了"
        self.statusBar().showMessage(msg, 10000)
        self._log_view.append(f"\n{msg}\n")
        QTimer.singleShot(1000, self._load_spreadsheet)

    # ──────────────────────────── 保存処理 ──────────────────────────────────

    def _on_save_requested(self, phones_data: dict, selections: dict, docomo_rep: str):
        if self._saver and self._saver.isRunning():
            return
        self._saver = PhoneManagerSaver(
            self._script_dir, phones_data, selections, docomo_rep
        )
        self._saver.saved.connect(
            lambda msg: self.statusBar().showMessage(f"✅  {msg}", 8000)
        )
        self._saver.error_signal.connect(
            lambda e: self._log_view.append(f"\n❌  保存エラー: {e}\n")
        )
        self._saver.start()

    def _on_save_and_run_requested(self, phones_data: dict, selections: dict,
                                   docomo_rep: str, carrier_configs: list):
        if self._saver and self._saver.isRunning():
            return
        selected_phones = []
        for carrier_sel in selections.values():
            selected_phones.extend(carrier_sel.keys())

        self._saver = PhoneManagerSaver(
            self._script_dir, phones_data, selections, docomo_rep
        )
        self._saver.saved.connect(
            lambda msg, cfgs=carrier_configs, phones=selected_phones: (
                self.statusBar().showMessage(f"✅  {msg}", 5000),
                self._start_download(cfgs, phones),
            )
        )
        self._saver.error_signal.connect(
            lambda e: self._log_view.append(f"\n❌  保存エラー: {e}\n")
        )
        self._saver.start()

    def _on_error(self, error: str):
        self._set_running(False)
        self.statusBar().showMessage("❌  エラーが発生しました", 10000)
        self._log_view.append(f"\n❌  エラー: {error}\n")
