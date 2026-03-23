"""SMS認証コード入力ダイアログ（1文字1ボックス方式）"""

from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QDialog, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QVBoxLayout, QWidget,
)

from .styles import (
    ACCENT, AMBER, BG_MAIN, BG_HOVER, TEXT, TEXT_MUTED, RED,
)


class _DigitBox(QLineEdit):
    """1文字入力用ボックス。入力後に次のボックスへ自動フォーカス移動。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._next_box: "_DigitBox | None" = None
        self._prev_box: "_DigitBox | None" = None
        self.setMaxLength(1)
        self.setAlignment(Qt.AlignCenter)
        self.setFixedSize(48, 56)
        self.setFont(QFont("Menlo", 22, QFont.Bold))
        self.setStyleSheet(f"""
            QLineEdit {{
                background: {BG_MAIN};
                color: {TEXT};
                border: 2px solid rgba(255,255,255,0.1);
                border-radius: 6px;
                font-size: 22px;
                font-weight: bold;
                text-align: center;
            }}
            QLineEdit:focus {{
                border-color: {ACCENT};
            }}
        """)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Backspace:
            if not self.text() and self._prev_box:
                self._prev_box.setFocus()
                self._prev_box.clear()
            else:
                super().keyPressEvent(event)
        elif event.key() in (Qt.Key_Left,):
            if self._prev_box:
                self._prev_box.setFocus()
        elif event.key() in (Qt.Key_Right,):
            if self._next_box:
                self._next_box.setFocus()
        else:
            super().keyPressEvent(event)

    def textChanged_handler(self, text):
        """1文字入力されたら次のボックスへ移動。"""
        if text and self._next_box:
            self._next_box.setFocus()
            self._next_box.selectAll()


class SmsCodeDialog(QDialog):
    """SMS認証コード入力ダイアログ。

    Args:
        phone     : 電話番号
        device    : 端末名（SMS受信端末）
        code_file : コードを書き込むファイルパス
        digits    : 桁数（3 or 6）
        timeout   : タイムアウト秒数
    """

    def __init__(
        self,
        phone: str,
        device: str,
        code_file: str,
        digits: int = 3,
        timeout: int = 60,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("SMS認証")
        self.setModal(True)
        self.setFixedWidth(420)
        self._code_file = Path(code_file)
        self._remaining = timeout
        self._digits    = digits
        self._boxes: list[_DigitBox] = []

        self.setStyleSheet(f"""
            QDialog {{
                background: {BG_HOVER};
                border: 1px solid {AMBER};
                border-radius: 12px;
            }}
            QLabel {{
                background: transparent;
                color: {TEXT};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        # タイトル
        title = QLabel("📱 SMS認証コードを入力")
        title.setStyleSheet(f"font-size: 16px; font-weight: bold; color: {TEXT};")
        layout.addWidget(title)

        # 電話番号・端末名
        phone_label = QLabel(f"<b>{phone}</b>")
        phone_label.setStyleSheet(f"font-size: 15px; color: {TEXT};")
        layout.addWidget(phone_label)

        device_text = device if device else "（端末未登録）"
        dev_label = QLabel(f"送信先端末: {device_text}")
        dev_label.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")
        layout.addWidget(dev_label)

        # 桁数の説明
        hint = QLabel(f"SMSに届いた {digits} 桁のコードを入力してください")
        hint.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")
        layout.addWidget(hint)

        # 入力ボックス群
        boxes_widget = QWidget()
        boxes_layout = QHBoxLayout(boxes_widget)
        boxes_layout.setSpacing(8)
        boxes_layout.setContentsMargins(0, 0, 0, 0)
        boxes_layout.addStretch()

        for i in range(digits):
            box = _DigitBox()
            box.textChanged.connect(box.textChanged_handler)
            self._boxes.append(box)
            boxes_layout.addWidget(box)

        # ボックス間のリンク設定
        for i, box in enumerate(self._boxes):
            if i > 0:
                box._prev_box = self._boxes[i - 1]
            if i < len(self._boxes) - 1:
                box._next_box = self._boxes[i + 1]

        boxes_layout.addStretch()
        layout.addWidget(boxes_widget)

        # カウントダウン
        self._timer_label = QLabel(f"残り {self._remaining} 秒")
        self._timer_label.setAlignment(Qt.AlignCenter)
        self._timer_label.setStyleSheet(f"color: {TEXT_MUTED}; font-size: 12px;")
        layout.addWidget(self._timer_label)

        # ボタン
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        self._ok_btn = QPushButton("送信")
        self._ok_btn.setEnabled(False)
        self._ok_btn.setFixedHeight(40)
        self._ok_btn.setStyleSheet(f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #4f46e5, stop:1 #6366f1);
                color: white;
                border: none;
                border-radius: 8px;
                font-size: 14px;
                font-weight: bold;
            }}
            QPushButton:disabled {{
                background: #2a2a40;
                color: {TEXT_MUTED};
            }}
            QPushButton:hover:!disabled {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #5b52f0, stop:1 #7173f5);
            }}
        """)
        self._ok_btn.clicked.connect(self._submit)

        cancel_btn = QPushButton("キャンセル")
        cancel_btn.setFixedHeight(40)
        cancel_btn.clicked.connect(self.reject)

        btn_row.addWidget(self._ok_btn)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

        # 各ボックスの変更を監視して送信ボタンの有効/無効を切り替え
        for box in self._boxes:
            box.textChanged.connect(self._check_complete)

        # Enter キーで送信
        self._ok_btn.setDefault(True)

        # カウントダウンタイマー
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(1000)

        # 最初のボックスにフォーカス
        if self._boxes:
            self._boxes[0].setFocus()

    def _check_complete(self):
        """全ボックスが埋まったら送信ボタンを有効化。"""
        complete = all(b.text() for b in self._boxes)
        self._ok_btn.setEnabled(complete)

    def _tick(self):
        self._remaining -= 1
        self._timer_label.setText(f"残り {self._remaining} 秒")
        if self._remaining <= 10:
            self._timer_label.setStyleSheet(f"color: {RED}; font-weight: bold; font-size: 12px;")
        if self._remaining <= 0:
            self._timer.stop()
            self.reject()

    def _submit(self):
        code = "".join(b.text() for b in self._boxes)
        if len(code) != self._digits:
            return
        self._timer.stop()
        try:
            self._code_file.write_text(code, encoding="utf-8")
        except Exception:
            pass
        self.accept()
