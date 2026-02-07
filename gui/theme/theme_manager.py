"""
Theme manager for PyQt5 - replaces the 212-line dark_mode_manager.py.
Uses QSS stylesheets instead of recursive widget walking.
"""

from PyQt5.QtWidgets import QApplication


LIGHT_STYLESHEET = """
QMainWindow {
    background-color: #f0f0f0;
}
QWidget {
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: 10pt;
}
QTabWidget::pane {
    border: 1px solid #cccccc;
    background: white;
}
QTabBar::tab {
    background: #e0e0e0;
    padding: 8px 20px;
    margin-right: 2px;
    border: 1px solid #cccccc;
    border-bottom: none;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
}
QTabBar::tab:selected {
    background: white;
    border-bottom: 2px solid #0078D4;
}
QTabBar::tab:hover:!selected {
    background: #d0d0d0;
}
QGroupBox {
    font-weight: bold;
    border: 1px solid #cccccc;
    border-radius: 4px;
    margin-top: 8px;
    padding-top: 16px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 5px;
}
QLineEdit {
    padding: 6px;
    border: 1px solid #cccccc;
    border-radius: 3px;
    background: white;
}
QLineEdit:focus {
    border: 1px solid #0078D4;
}
QPushButton {
    padding: 6px 16px;
    border: 1px solid #cccccc;
    border-radius: 3px;
    background: #f0f0f0;
    min-height: 24px;
}
QPushButton:hover {
    background: #e0e0e0;
    border-color: #0078D4;
}
QPushButton:pressed {
    background: #d0d0d0;
}
QPushButton:disabled {
    color: #999999;
    background: #f5f5f5;
}
QPushButton#primary {
    background: #0078D4;
    color: white;
    border: none;
    font-weight: bold;
}
QPushButton#primary:hover {
    background: #006cbd;
}
QPushButton#primary:pressed {
    background: #005a9e;
}
QPushButton#danger {
    background: #d83b01;
    color: white;
    border: none;
}
QPushButton#danger:hover {
    background: #c43400;
}
QTreeView {
    border: 1px solid #cccccc;
    alternate-background-color: #f7f7f7;
    selection-background-color: #0078D4;
    selection-color: white;
}
QTreeView::item {
    padding: 4px;
}
QHeaderView::section {
    background-color: #366092;
    color: white;
    padding: 6px;
    border: 1px solid #2d5280;
    font-weight: bold;
}
QTextEdit {
    border: 1px solid #cccccc;
    border-radius: 3px;
    background: white;
    font-family: "Consolas", "Courier New", monospace;
    font-size: 9pt;
}
QProgressBar {
    border: 1px solid #cccccc;
    border-radius: 3px;
    text-align: center;
    height: 22px;
}
QProgressBar::chunk {
    background-color: #0078D4;
    border-radius: 2px;
}
QSpinBox {
    padding: 4px;
    border: 1px solid #cccccc;
    border-radius: 3px;
}
QCheckBox {
    spacing: 6px;
}
QRadioButton {
    spacing: 6px;
}
QStatusBar {
    background: #f0f0f0;
    border-top: 1px solid #cccccc;
}
QScrollBar:vertical {
    width: 12px;
}
QScrollBar:horizontal {
    height: 12px;
}
"""

DARK_STYLESHEET = """
QMainWindow {
    background-color: #1e1e1e;
}
QWidget {
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: 10pt;
    color: #d4d4d4;
    background-color: #1e1e1e;
}
QTabWidget::pane {
    border: 1px solid #3c3c3c;
    background: #252526;
}
QTabBar::tab {
    background: #2d2d2d;
    color: #cccccc;
    padding: 8px 20px;
    margin-right: 2px;
    border: 1px solid #3c3c3c;
    border-bottom: none;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
}
QTabBar::tab:selected {
    background: #252526;
    color: white;
    border-bottom: 2px solid #0078D4;
}
QTabBar::tab:hover:!selected {
    background: #383838;
}
QGroupBox {
    font-weight: bold;
    color: #cccccc;
    border: 1px solid #3c3c3c;
    border-radius: 4px;
    margin-top: 8px;
    padding-top: 16px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 5px;
}
QLineEdit {
    padding: 6px;
    border: 1px solid #3c3c3c;
    border-radius: 3px;
    background: #2d2d2d;
    color: #d4d4d4;
}
QLineEdit:focus {
    border: 1px solid #0078D4;
}
QPushButton {
    padding: 6px 16px;
    border: 1px solid #3c3c3c;
    border-radius: 3px;
    background: #2d2d2d;
    color: #d4d4d4;
    min-height: 24px;
}
QPushButton:hover {
    background: #383838;
    border-color: #0078D4;
}
QPushButton:pressed {
    background: #1e1e1e;
}
QPushButton:disabled {
    color: #666666;
    background: #2a2a2a;
}
QPushButton#primary {
    background: #0078D4;
    color: white;
    border: none;
    font-weight: bold;
}
QPushButton#primary:hover {
    background: #006cbd;
}
QPushButton#primary:pressed {
    background: #005a9e;
}
QPushButton#danger {
    background: #d83b01;
    color: white;
    border: none;
}
QPushButton#danger:hover {
    background: #c43400;
}
QTreeView {
    border: 1px solid #3c3c3c;
    background: #1e1e1e;
    alternate-background-color: #252526;
    selection-background-color: #0078D4;
    selection-color: white;
    color: #d4d4d4;
}
QTreeView::item {
    padding: 4px;
}
QHeaderView::section {
    background-color: #2d5280;
    color: white;
    padding: 6px;
    border: 1px solid #1e4570;
    font-weight: bold;
}
QTextEdit {
    border: 1px solid #3c3c3c;
    border-radius: 3px;
    background: #1e1e1e;
    color: #d4d4d4;
    font-family: "Consolas", "Courier New", monospace;
    font-size: 9pt;
}
QProgressBar {
    border: 1px solid #3c3c3c;
    border-radius: 3px;
    text-align: center;
    height: 22px;
    background: #2d2d2d;
    color: white;
}
QProgressBar::chunk {
    background-color: #0078D4;
    border-radius: 2px;
}
QSpinBox {
    padding: 4px;
    border: 1px solid #3c3c3c;
    border-radius: 3px;
    background: #2d2d2d;
    color: #d4d4d4;
}
QCheckBox {
    spacing: 6px;
    color: #d4d4d4;
}
QRadioButton {
    spacing: 6px;
    color: #d4d4d4;
}
QStatusBar {
    background: #1e1e1e;
    color: #d4d4d4;
    border-top: 1px solid #3c3c3c;
}
QLabel {
    color: #d4d4d4;
}
QScrollBar:vertical {
    width: 12px;
    background: #1e1e1e;
}
QScrollBar::handle:vertical {
    background: #3c3c3c;
    border-radius: 4px;
}
QScrollBar:horizontal {
    height: 12px;
    background: #1e1e1e;
}
QScrollBar::handle:horizontal {
    background: #3c3c3c;
    border-radius: 4px;
}
"""


class ThemeManager:
    """Manages light/dark theme toggling via QSS stylesheets"""

    def __init__(self):
        self.is_dark = False

    def apply_theme(self, dark_mode: bool):
        """Apply theme to entire application"""
        self.is_dark = dark_mode
        app = QApplication.instance()
        if app:
            app.setStyleSheet(DARK_STYLESHEET if dark_mode else LIGHT_STYLESHEET)

    def toggle(self):
        """Toggle between light and dark themes"""
        self.apply_theme(not self.is_dark)
        return self.is_dark
