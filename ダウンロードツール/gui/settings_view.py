"""設定タブ（読み取り専用カード表示）"""

import os
import webbrowser

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QHBoxLayout, QLabel, QPushButton,
    QScrollArea, QVBoxLayout, QWidget,
)

from .styles import BG_CARD, BG_HOVER, BORDER, TEXT, TEXT_DIM, TEXT_MUTED, ACCENT

# パスワード系の設定名（マスク表示する）
PASSWORD_KEYS = {
    "パスワード", "au/UQパスワード", "dアカウントパスワード",
    "au暗証番号", "PIN", "暗証番号",
}


class _SettingCard(QWidget):
    """設定項目1件のカード。"""

    def __init__(self, name: str, value: str, parent=None):
        super().__init__(parent)
        self._value     = value
        self._masked    = name in PASSWORD_KEYS
        self._showing   = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(4)

        # ラベル（設定名）
        lbl = QLabel(name)
        lbl.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 11px;")
        layout.addWidget(lbl)

        # 値の行
        val_row = QHBoxLayout()
        val_row.setContentsMargins(0, 0, 0, 0)
        val_row.setSpacing(8)

        self._val_label = QLabel(self._masked_text())
        self._val_label.setStyleSheet(f"color: {TEXT}; font-size: 14px;")
        self._val_label.setWordWrap(True)
        val_row.addWidget(self._val_label, 1)

        # パスワード系は表示トグルボタン
        if self._masked:
            eye_btn = QPushButton("👁")
            eye_btn.setFixedSize(28, 28)
            eye_btn.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    border: none;
                    font-size: 14px;
                }
                QPushButton:hover { background: rgba(255,255,255,0.05); border-radius: 4px; }
            """)
            eye_btn.clicked.connect(self._toggle_mask)
            val_row.addWidget(eye_btn)

        layout.addLayout(val_row)

        self.setStyleSheet(f"""
            QWidget {{
                background: {BG_CARD};
                border: 1px solid rgba(255,255,255,0.06);
                border-radius: 10px;
            }}
        """)

    def _masked_text(self) -> str:
        if self._masked and not self._showing:
            return "••••••••"
        return self._value or "（未設定）"

    def _toggle_mask(self):
        self._showing = not self._showing
        self._val_label.setText(self._masked_text())


class SettingsView(QWidget):
    """設定タブの全体ウィジェット。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 8, 0, 0)
        layout.setSpacing(8)

        # ボタン行
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        reload_btn = QPushButton("🔄 スプレッドシートから再読み込み")
        reload_btn.setFixedHeight(34)
        btn_row.addWidget(reload_btn)

        open_btn = QPushButton("📋 スプレッドシートを開く")
        open_btn.setFixedHeight(34)
        open_btn.clicked.connect(self._open_spreadsheet)
        btn_row.addWidget(open_btn)

        btn_row.addStretch()
        layout.addLayout(btn_row)

        self._reload_btn = reload_btn

        # スクロールエリア
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("background: transparent; border: none;")

        self._content = QWidget()
        self._content.setStyleSheet("background: transparent;")
        self._cards_layout = QVBoxLayout(self._content)
        self._cards_layout.setContentsMargins(0, 0, 0, 0)
        self._cards_layout.setSpacing(6)
        self._cards_layout.addStretch()

        scroll.setWidget(self._content)
        layout.addWidget(scroll)

    @property
    def reload_btn(self) -> QPushButton:
        return self._reload_btn

    def load_settings(self, settings: list[tuple[str, str]]):
        """設定データをカード一覧として表示する。"""
        # 既存カードをクリア
        while self._cards_layout.count() > 1:
            item = self._cards_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        for name, value in settings:
            card = _SettingCard(name, value)
            self._cards_layout.insertWidget(self._cards_layout.count() - 1, card)

    def _open_spreadsheet(self):
        url = os.environ.get("SPREADSHEET_URL", "").strip()
        if url:
            webbrowser.open(url)
