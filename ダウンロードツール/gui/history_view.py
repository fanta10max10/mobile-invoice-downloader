"""ダウンロード履歴タブ（カード表示）"""

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QHBoxLayout, QLabel, QPushButton,
    QScrollArea, QVBoxLayout, QWidget,
)

from .styles import BG_CARD, BG_HOVER, BORDER, BORDER_LIGHT, GREEN, AMBER, RED, TEXT, TEXT_DIM, TEXT_MUTED, TEXT_FAINT, ACCENT_CHIP

CARRIER_ICONS = {
    "SoftBank": "🐶",
    "Ymobile":  "🐱",
    "au":       "🍊",
    "UQmobile": "👸",
    "docomo":   "🍄",
}


class _HistoryCard(QWidget):
    """履歴1件のカード。"""

    def __init__(self, row: dict, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(4)

        # 上段: キャリア + 結果バッジ + 金額
        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        carrier = str(row.get("キャリア", "")).strip()
        icon    = CARRIER_ICONS.get(carrier, "📱")
        carrier_lbl = QLabel(f"{icon} {carrier}")
        carrier_lbl.setStyleSheet(f"color: {TEXT}; font-size: 13px; font-weight: bold;")
        top_row.addWidget(carrier_lbl)

        top_row.addStretch()

        # 金額（ファイル名から取得できる場合）
        filename = str(row.get("ファイル名", "")).strip()
        amount   = _extract_amount(filename)
        if amount:
            amt_lbl = QLabel(amount)
            amt_lbl.setStyleSheet(f"color: {TEXT}; font-size: 16px; font-weight: bold;")
            top_row.addWidget(amt_lbl)

        # 結果バッジ
        result = str(row.get("結果", "")).strip()
        if result in ("success", "成功", "✅"):
            badge_text  = "✅ 成功"
            badge_color = GREEN
            badge_bg    = "rgba(52,211,153,0.12)"
        elif result in ("skipped", "スキップ", "⏭"):
            badge_text  = "⏭ スキップ"
            badge_color = TEXT_MUTED
            badge_bg    = "rgba(100,116,139,0.12)"
        else:
            badge_text  = "❌ 失敗"
            badge_color = RED
            badge_bg    = "rgba(248,113,113,0.12)"

        badge = QLabel(badge_text)
        badge.setStyleSheet(f"""
            background: {badge_bg};
            color: {badge_color};
            border-radius: 4px;
            padding: 2px 8px;
            font-size: 11px;
            font-weight: bold;
        """)
        top_row.addWidget(badge)
        layout.addLayout(top_row)

        # 電話番号
        phone = str(row.get("電話番号", "")).strip()
        if phone:
            phone_lbl = QLabel(phone)
            phone_lbl.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")
            layout.addWidget(phone_lbl)

        # 日時
        dt = str(row.get("日時", "") or row.get("ダウンロード日時", "")).strip()
        if dt:
            dt_lbl = QLabel(dt)
            dt_lbl.setStyleSheet(f"color: {TEXT_FAINT}; font-size: 11px;")
            layout.addWidget(dt_lbl)

        # ファイル名
        if filename:
            fn_lbl = QLabel(filename)
            fn_lbl.setStyleSheet(f"color: {TEXT_FAINT}; font-size: 10px; font-family: 'Menlo';")
            fn_lbl.setWordWrap(True)
            layout.addWidget(fn_lbl)

        self.setStyleSheet(f"""
            QWidget {{
                background: {BG_CARD};
                border: 1px solid rgba(255,255,255,0.06);
                border-radius: 10px;
            }}
        """)


def _extract_amount(filename: str) -> str:
    """ファイル名から金額を抽出する（例: _4980円(税抜).pdf → ¥4,980）。"""
    import re
    m = re.search(r"_(\d+)円", filename)
    if m:
        amount = int(m.group(1))
        return f"¥{amount:,}"
    return ""


class HistoryView(QWidget):
    """ダウンロード履歴タブの全体ウィジェット。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 8, 0, 0)
        layout.setSpacing(8)

        # 再読み込みボタン
        btn_row = QHBoxLayout()
        reload_btn = QPushButton("🔄 スプレッドシートから再読み込み")
        reload_btn.setFixedHeight(34)
        btn_row.addWidget(reload_btn)
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

    def load_history(self, rows: list[dict]):
        """履歴データをカード一覧として表示する（最新が上）。"""
        while self._cards_layout.count() > 1:
            item = self._cards_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        for row in reversed(rows):
            card = _HistoryCard(row)
            self._cards_layout.insertWidget(self._cards_layout.count() - 1, card)
