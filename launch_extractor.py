"""
Convenience launcher - opens the main app on the Extract & Refresh tab (Tab 4).
"""

import sys
import os

sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def main():
    try:
        from PyQt5.QtWidgets import QApplication
        from gui.pyqt_main_window import GSTOrganizerWindow

        app = QApplication(sys.argv)
        window = GSTOrganizerWindow()
        window.tabs.setCurrentIndex(3)  # Switch to Tab 4 (Extract & Refresh)
        window.show()
        sys.exit(app.exec_())

    except Exception as e:
        print(f"Failed to launch: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
