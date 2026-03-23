#!/usr/bin/env python3
"""携帯領収書管理 GUIアプリ エントリーポイント

起動方法: python3 gui_app.py
"""

import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

try:
    from PySide6.QtWidgets import QApplication
    from PySide6.QtCore import Qt
except ImportError:
    print("PySide6 がインストールされていません。")
    print("  pip install PySide6")
    sys.exit(1)

from gui.styles import APP_QSS
from gui.main_window import MainWindow


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("携帯領収書管理")
    # AA_UseHighDpiPixmaps is deprecated in Qt6 (enabled by default)
    # app.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app.setStyleSheet(APP_QSS)

    window = MainWindow(SCRIPT_DIR)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
