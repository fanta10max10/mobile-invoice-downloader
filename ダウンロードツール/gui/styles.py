"""スタイル定数・QSSテーマ・カスタムウィジェット定義

デザインコンセプト: "Dark Matter"
 - Linear / Arc / Vercel / Raycast から着想
 - 深宇宙ネイビー + エレクトリックバイオレットアクセント
 - ピルタブ / グラデーションCTA / モノスペース技術テキスト
"""

import re
from datetime import date

from PySide6.QtCore import (
    Qt, QPropertyAnimation, QEasingCurve, Property, QRectF, QSize, Signal,
)
from PySide6.QtGui import QColor, QPainter, QPen
from PySide6.QtWidgets import (
    QAbstractButton, QHBoxLayout, QLabel, QPushButton, QWidget,
)

# ─── カラーパレット（Dark Matter） ───
BG_MAIN      = "#080812"   # 深宇宙ブラック（メイン背景）
BG_SURFACE   = "#0e0e1c"   # サーフェス（ヘッダー・タブバー）
BG_HOVER     = "#14142a"   # ホバー状態
BG_CARD      = "#0c0c1a"   # カード背景
BG_DEEPER    = "#04040b"   # ログ・深い背景
BG_ELEVATED  = "#181830"   # 浮き上がった要素（ドロップダウン等）

BORDER       = "rgba(255, 255, 255, 0.05)"
BORDER_LIGHT = "rgba(255, 255, 255, 0.10)"
BORDER_FOCUS = "rgba(139, 92, 246, 0.55)"

# アクセント: バイオレット系
ACCENT       = "#8b5cf6"   # violet-500
ACCENT_DARK  = "#7c3aed"   # violet-600
ACCENT2      = "#6366f1"   # indigo-500
ACCENT_SOFT  = "rgba(139, 92, 246, 0.12)"
ACCENT_GLOW  = "rgba(139, 92, 246, 0.22)"
ACCENT_CHIP  = "rgba(139, 92, 246, 0.14)"

# ステータスカラー
GREEN        = "#10b981"   # emerald-500
GREEN_SOFT   = "rgba(16, 185, 129, 0.10)"
GREEN_DIM    = "#059669"

AMBER        = "#f59e0b"   # amber-500
AMBER_SOFT   = "rgba(245, 158, 11, 0.12)"

RED          = "#ef4444"   # red-500
RED_SOFT     = "rgba(239, 68, 68, 0.12)"

BLUE         = "#3b82f6"   # blue-500
BLUE_SOFT    = "rgba(59, 130, 246, 0.12)"

# テキスト階層
TEXT         = "#f1f5f9"   # slate-100 (主テキスト)
TEXT_DIM     = "#cbd5e1"   # slate-300
TEXT_MUTED   = "#64748b"   # slate-500
TEXT_FAINT   = "#334155"   # slate-700

# ─── QSSアプリケーションスタイル ───
APP_QSS = f"""
/* ── ベース ── */
QWidget {{
    background-color: {BG_MAIN};
    color: {TEXT};
    font-family: ".AppleSystemUIFont", "Helvetica Neue";
    font-size: 13px;
}}
QMainWindow {{
    background-color: {BG_MAIN};
}}

/* ── スクロールバー（細く控えめに）── */
QScrollBar:vertical {{
    background: transparent;
    width: 5px;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background: rgba(139, 92, 246, 0.25);
    border-radius: 2px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{
    background: rgba(139, 92, 246, 0.45);
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
    background: transparent;
}}
QScrollBar:horizontal {{
    background: transparent;
    height: 5px;
}}
QScrollBar::handle:horizontal {{
    background: rgba(139, 92, 246, 0.25);
    border-radius: 2px;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0;
}}

/* ── ラベル ── */
QLabel {{
    background: transparent;
    color: {TEXT};
}}

/* ── プッシュボタン（デフォルト）── */
QPushButton {{
    background-color: transparent;
    color: {TEXT_MUTED};
    border: 1px solid {BORDER_LIGHT};
    border-radius: 8px;
    padding: 7px 16px;
    font-size: 13px;
}}
QPushButton:hover {{
    background-color: {BG_HOVER};
    color: {TEXT_DIM};
    border-color: rgba(255,255,255,0.15);
}}
QPushButton:pressed {{
    background-color: {BG_MAIN};
    color: {TEXT};
}}
QPushButton:disabled {{
    color: {TEXT_FAINT};
    background-color: transparent;
    border-color: {BORDER};
}}

/* ── アクセントボタン（🚀 全キャリア一括実行）── */
QPushButton#accentBtn {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #6d28d9, stop:0.5 {ACCENT}, stop:1 {ACCENT2});
    color: white;
    border: none;
    border-radius: 12px;
    font-size: 14px;
    font-weight: 700;
    min-height: 48px;
    letter-spacing: 0.3px;
}}
QPushButton#accentBtn:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #7c3aed, stop:0.5 #9b72fb, stop:1 #7475f5);
}}
QPushButton#accentBtn:pressed {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 #5b21b6, stop:1 #4f46e5);
}}
QPushButton#accentBtn:disabled {{
    background: {BG_HOVER};
    color: {TEXT_FAINT};
    border: 1px solid {BORDER};
}}

/* ── キャリアタブボタン（ピル型）── */
QPushButton#tabBtn {{
    background: transparent;
    color: {TEXT_MUTED};
    border: none;
    border-radius: 20px;
    padding: 7px 20px;
    font-size: 13px;
    font-weight: 500;
    min-width: 100px;
}}
QPushButton#tabBtn:hover {{
    background: rgba(255, 255, 255, 0.04);
    color: {TEXT_DIM};
}}
QPushButton#tabBtn[active="true"] {{
    background: {ACCENT_CHIP};
    color: #c4b5fd;
    border: 1px solid rgba(139, 92, 246, 0.28);
    font-weight: 600;
}}
QPushButton#tabBtn[active="true"]:hover {{
    background: rgba(139, 92, 246, 0.20);
}}

/* ── 実行ボタン（キャリアタブ内）── */
QPushButton#runBtn {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 {ACCENT_DARK}, stop:1 {ACCENT2});
    color: white;
    border: none;
    border-radius: 9px;
    font-size: 13px;
    font-weight: 600;
    min-height: 38px;
    padding: 8px 22px;
}}
QPushButton#runBtn:hover {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #8b5cf6, stop:1 #7475f5);
}}
QPushButton#runBtn:disabled {{
    background: {BG_HOVER};
    color: {TEXT_FAINT};
}}

/* ── コンテンツタブ（ログ/設定/履歴）── */
QPushButton#contentTabBtn {{
    background: transparent;
    color: {TEXT_MUTED};
    border: none;
    border-bottom: 2px solid transparent;
    border-radius: 0;
    padding: 10px 28px;
    font-size: 13px;
    font-weight: 500;
}}
QPushButton#contentTabBtn:hover {{
    color: {TEXT_DIM};
    background: rgba(255,255,255,0.02);
}}
QPushButton#contentTabBtn[active="true"] {{
    color: #c4b5fd;
    border-bottom: 2px solid {ACCENT};
    font-weight: 600;
}}

/* ── ナビゲーションボタン（月選択の矢印等）── */
QPushButton#navBtn {{
    background: transparent;
    color: {TEXT_MUTED};
    border: none;
    border-radius: 12px;
    font-size: 18px;
    font-weight: 300;
    padding: 0;
    min-width: 28px;
    min-height: 28px;
    max-width: 28px;
    max-height: 28px;
}}
QPushButton#navBtn:hover {{
    background: {ACCENT_SOFT};
    color: #c4b5fd;
}}
QPushButton#navBtn:pressed {{
    background: {ACCENT_CHIP};
}}

/* ── チェックボックス ── */
QCheckBox {{
    background: transparent;
    color: {TEXT};
    spacing: 10px;
}}
QCheckBox::indicator {{
    width: 17px;
    height: 17px;
    border: 1.5px solid rgba(255,255,255,0.18);
    border-radius: 5px;
    background: {BG_MAIN};
}}
QCheckBox::indicator:hover {{
    border-color: {ACCENT};
    background: {ACCENT_SOFT};
}}
QCheckBox::indicator:checked {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
        stop:0 {ACCENT_DARK}, stop:1 {ACCENT});
    border-color: transparent;
    image: none;
}}

/* ── テキストエリア（ログ）── */
QTextEdit {{
    background-color: {BG_DEEPER};
    color: {TEXT_DIM};
    border: 1px solid {BORDER};
    border-radius: 10px;
    padding: 14px;
    font-family: "Menlo", "Monaco";
    font-size: 12px;
    line-height: 1.7;
    selection-background-color: {ACCENT_DARK};
}}

/* ── ステータスバー ── */
QStatusBar {{
    background-color: {BG_DEEPER};
    color: {TEXT_FAINT};
    border-top: 1px solid {BORDER};
    font-size: 11px;
    font-family: "Menlo";
}}
QStatusBar QLabel {{
    color: {TEXT_FAINT};
    font-size: 11px;
    font-family: "Menlo";
    background: transparent;
    padding: 0 8px;
}}

/* ── ラインエディット ── */
QLineEdit {{
    background-color: {BG_SURFACE};
    color: {TEXT};
    border: 1.5px solid {BORDER_LIGHT};
    border-radius: 8px;
    padding: 6px 10px;
    font-size: 13px;
}}
QLineEdit:focus {{
    border-color: {ACCENT};
    background: rgba(139, 92, 246, 0.05);
}}
QLineEdit:disabled {{
    color: {TEXT_FAINT};
    border-color: {BORDER};
}}

/* ── スクロールエリア ── */
QScrollArea {{
    background: transparent;
    border: none;
}}
QScrollArea > QWidget > QWidget {{
    background: transparent;
}}

/* ── ダイアログ ── */
QDialog {{
    background-color: {BG_SURFACE};
}}

/* ── コンボボックス（PDFの種類）── */
QComboBox {{
    background: {BG_SURFACE};
    color: {TEXT};
    border: 1px solid {BORDER_LIGHT};
    border-radius: 7px;
    padding: 3px 10px;
    font-size: 11px;
    min-height: 26px;
}}
QComboBox:hover {{
    border-color: rgba(139, 92, 246, 0.35);
    background: {BG_HOVER};
}}
QComboBox:focus {{
    border-color: {ACCENT};
}}
QComboBox::drop-down {{
    border: none;
    width: 18px;
}}
QComboBox::down-arrow {{
    width: 8px;
    height: 8px;
}}
QComboBox QAbstractItemView {{
    background: {BG_ELEVATED};
    color: {TEXT};
    selection-background-color: {ACCENT_DARK};
    border: 1px solid rgba(139, 92, 246, 0.20);
    border-radius: 8px;
    padding: 4px;
    outline: none;
}}
QComboBox QAbstractItemView::item {{
    min-height: 26px;
    padding: 3px 8px;
    border-radius: 4px;
}}
"""


class ToggleSwitch(QAbstractButton):
    """滑らかなトグルスイッチ（ON=バイオレット / OFF=グレー）"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setFixedSize(42, 22)
        self._thumb_pos = 0.0

        self._anim = QPropertyAnimation(self, b"thumb_pos", self)
        self._anim.setDuration(160)
        self._anim.setEasingCurve(QEasingCurve.InOutCubic)
        self.toggled.connect(self._on_toggled)

    def _on_toggled(self, checked: bool):
        self._anim.stop()
        self._anim.setStartValue(self._thumb_pos)
        self._anim.setEndValue(1.0 if checked else 0.0)
        self._anim.start()

    def get_thumb_pos(self) -> float:
        return self._thumb_pos

    def set_thumb_pos(self, pos: float):
        self._thumb_pos = pos
        self.update()

    thumb_pos = Property(float, get_thumb_pos, set_thumb_pos)

    def paintEvent(self, _event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)

        w, h = self.width(), self.height()
        r = h / 2

        # トラック
        if self.isChecked():
            # ON: violet gradient
            track_color = QColor(ACCENT_DARK)
            # 徐々に色を変えるためにthumb_posを使う
            c1 = QColor("#4c1d95")
            c2 = QColor(ACCENT)
            t = self._thumb_pos
            track_color = QColor(
                int(c1.red()   * (1-t) + c2.red()   * t),
                int(c1.green() * (1-t) + c2.green() * t),
                int(c1.blue()  * (1-t) + c2.blue()  * t),
            )
        else:
            track_color = QColor("#1e1e38")
        p.setPen(Qt.NoPen)
        p.setBrush(track_color)
        p.drawRoundedRect(0, 0, w, h, r, r)

        # サム（白い円）
        pad = 3
        travel = w - h
        x = pad + travel * self._thumb_pos
        thumb_size = h - pad * 2
        p.setBrush(QColor("white"))
        p.drawEllipse(QRectF(x, pad, thumb_size, thumb_size))

    def sizeHint(self):
        return QSize(42, 22)


def _prev_month(year: int, month: int) -> tuple:
    month -= 1
    if month < 1:
        month = 12
        year -= 1
    return year, month


def _next_month(year: int, month: int) -> tuple:
    month += 1
    if month > 12:
        month = 1
        year += 1
    return year, month


class MonthSelector(QWidget):
    """対象月の選択ウィジェット。

    ‹  2026年03月  ›  の形式で表示。
    矢印クリックで月を変更し、month_changed シグナルを発火。
    """

    month_changed = Signal(str)   # "YYYY年MM月" フォーマット

    def __init__(self, parent=None):
        super().__init__(parent)
        # デフォルト: 前月
        today = date.today()
        self._year, self._month = _prev_month(today.year, today.month)
        self._auto = True  # スプレッドシートからまだ読んでいない状態

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(3)

        self._prev_btn = QPushButton("‹")
        self._prev_btn.setObjectName("navBtn")
        self._prev_btn.setFixedSize(28, 28)
        self._prev_btn.clicked.connect(self._on_prev)

        self._label = QLabel(self._format())
        self._label.setAlignment(Qt.AlignCenter)
        self._label.setFixedWidth(136)
        self._label.setStyleSheet(f"""
            QLabel {{
                background: {ACCENT_CHIP};
                color: #c4b5fd;
                border: 1px solid rgba(139, 92, 246, 0.28);
                border-radius: 14px;
                padding: 5px 14px;
                font-size: 13px;
                font-weight: 600;
            }}
        """)

        self._next_btn = QPushButton("›")
        self._next_btn.setObjectName("navBtn")
        self._next_btn.setFixedSize(28, 28)
        self._next_btn.clicked.connect(self._on_next)

        layout.addWidget(self._prev_btn)
        layout.addWidget(self._label)
        layout.addWidget(self._next_btn)

    def _format(self) -> str:
        return f"{self._year}年{self._month:02d}月"

    def current_month_str(self) -> str:
        """現在の月を "YYYY年MM月" 形式で返す。"""
        return self._format()

    def set_month(self, month_str: str):
        """スプレッドシートから読み込んだ月をセットする。変更シグナルは発火しない。"""
        if not month_str or month_str in ("自動", "前月", "自動（前月）"):
            # 自動モード: 前月のまま表示
            self._auto = True
            self._label.setText(self._format())
            return

        m = re.match(r"^(\d{4})年(\d+)月$", month_str)
        if m:
            self._year, self._month = int(m.group(1)), int(m.group(2))
            self._auto = False
            self._label.setText(self._format())
            return

        m = re.match(r"^(\d{4})(\d{2})$", month_str)
        if m:
            self._year, self._month = int(m.group(1)), int(m.group(2))
            self._auto = False
            self._label.setText(self._format())
            return

    def _on_prev(self):
        self._auto = False
        self._year, self._month = _prev_month(self._year, self._month)
        self._label.setText(self._format())
        self.month_changed.emit(self._format())

    def _on_next(self):
        self._auto = False
        self._year, self._month = _next_month(self._year, self._month)
        self._label.setText(self._format())
        self.month_changed.emit(self._format())
