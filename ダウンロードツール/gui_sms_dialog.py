"""SMS認証コード入力ダイアログ

ask_security_code() がファイル待機モードに入った際に表示し、
ユーザーが入力したコードを /tmp/{carrier}_security_code.txt に書き込む。
"""

from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QIntValidator
from PySide6.QtWidgets import (
    QDialog, QHBoxLayout, QLabel, QLineEdit, QPushButton, QVBoxLayout,
)


class SmsCodeDialog(QDialog):
    """SMS認証コード入力ダイアログ。

    表示内容:
      - 電話番号
      - 端末名（SMSが届くデバイス）
      - 桁数に応じた入力フィールド（3桁 or 6桁）
      - カウントダウンタイマー
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
        self.setWindowTitle("SMS認証コード入力")
        self.setModal(True)
        self.setMinimumWidth(380)
        self._code_file = Path(code_file)
        self._remaining = timeout
        self._code = None

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # 情報表示
        title = QLabel("📱 SMS認証が必要です")
        title.setFont(QFont("", 16, QFont.Bold))
        layout.addWidget(title)

        layout.addWidget(QLabel(f"電話番号: {phone}"))
        if device:
            layout.addWidget(QLabel(f"端末: {device}"))

        hint = QLabel(f"SMSに届いた{digits}桁のセキュリティ番号を入力してください")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        # コード入力フィールド
        self._input = QLineEdit()
        self._input.setMaxLength(digits)
        self._input.setValidator(QIntValidator())
        self._input.setPlaceholderText("0" * digits)
        self._input.setAlignment(Qt.AlignCenter)
        self._input.setFont(QFont("Menlo", 28))
        self._input.setMinimumHeight(50)
        layout.addWidget(self._input)

        # カウントダウン
        self._timer_label = QLabel(f"残り {self._remaining} 秒")
        self._timer_label.setAlignment(Qt.AlignCenter)
        self._timer_label.setStyleSheet("color: gray;")
        layout.addWidget(self._timer_label)

        # ボタン
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("送信")
        ok_btn.setDefault(True)
        ok_btn.setMinimumHeight(36)
        ok_btn.clicked.connect(self._submit)
        cancel_btn = QPushButton("キャンセル")
        cancel_btn.setMinimumHeight(36)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        # Enter で送信
        self._input.returnPressed.connect(self._submit)

        # カウントダウンタイマー
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._tick)
        self._timer.start(1000)

        # 入力フィールドにフォーカス
        self._input.setFocus()

    def _tick(self):
        self._remaining -= 1
        self._timer_label.setText(f"残り {self._remaining} 秒")
        if self._remaining <= 10:
            self._timer_label.setStyleSheet("color: red; font-weight: bold;")
        if self._remaining <= 0:
            self._timer.stop()
            self.reject()

    def _submit(self):
        code = self._input.text().strip()
        if not code:
            return
        self._code = code
        self._timer.stop()
        # コードをファイルに書き込む（ask_security_code のポーリングが読み取る）
        try:
            self._code_file.write_text(code, encoding="utf-8")
        except Exception:
            pass
        self.accept()

    def get_code(self) -> str | None:
        """入力されたコードを返す。キャンセル/タイムアウト時はNone。"""
        return self._code
