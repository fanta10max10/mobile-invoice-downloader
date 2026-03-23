"""ログ表示ウィジェット（色分け・自動スクロール）"""

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import QTextEdit, QWidget, QVBoxLayout, QPushButton

from .styles import GREEN, AMBER, RED, ACCENT, TEXT_MUTED, TEXT


class LogView(QWidget):
    """カラーコーディング付きログ表示ウィジェット。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 8, 0, 0)
        layout.setSpacing(4)

        # クリアボタン
        clear_btn = QPushButton("🗑 クリア")
        clear_btn.setObjectName("clearBtn")
        clear_btn.setFixedWidth(80)
        clear_btn.setFixedHeight(28)
        clear_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #64748b;
                border: 1px solid rgba(255,255,255,0.1);
                border-radius: 6px;
                font-size: 11px;
                padding: 2px 8px;
            }
            QPushButton:hover {
                color: #e2e8f0;
                border-color: rgba(255,255,255,0.2);
            }
        """)
        layout.addWidget(clear_btn, alignment=Qt.AlignRight)

        self._text = QTextEdit()
        self._text.setReadOnly(True)
        self._text.setFont(QFont("Menlo", 11))
        layout.addWidget(self._text)

        clear_btn.clicked.connect(self._text.clear)

    def append(self, text: str):
        """テキストを追記する。内容に応じて色を変える。"""
        cursor = self._text.textCursor()
        cursor.movePosition(QTextCursor.End)

        fmt = QTextCharFormat()
        if "[ERROR]" in text or "❌" in text or "失敗" in text:
            fmt.setForeground(QColor(RED))
        elif "[WARNING]" in text or "⚠" in text or "SMS認証待ち" in text or "⏳" in text:
            fmt.setForeground(QColor(AMBER))
        elif "✅" in text or "成功" in text or "完了" in text:
            fmt.setForeground(QColor(GREEN))
        elif "===" in text or "開始" in text or "[INFO]" in text and "=" in text:
            fmt.setForeground(QColor(ACCENT))
        elif "[INFO]" in text:
            fmt.setForeground(QColor(TEXT_MUTED))
        else:
            fmt.setForeground(QColor(TEXT_MUTED))

        cursor.setCharFormat(fmt)
        cursor.insertText(text)
        self._text.setTextCursor(cursor)
        self._text.ensureCursorVisible()

    def clear(self):
        self._text.clear()
