"""
Processing View - Tab 3: Progress bar, start/stop, and log output.
"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QProgressBar, QTextEdit
)
from PyQt5.QtCore import pyqtSignal, Qt
from PyQt5.QtGui import QTextCursor, QColor


class ProcessingView(QWidget):
    """Tab 3: processing progress and log"""

    start_requested = pyqtSignal()
    stop_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)

        # Progress section
        progress_layout = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)

        progress_layout.addWidget(self.progress_bar)
        layout.addLayout(progress_layout)

        # Status label
        self.status_label = QLabel("Ready to process")
        layout.addWidget(self.status_label)

        # Control buttons
        btn_layout = QHBoxLayout()

        self.btn_start = QPushButton("Start Processing")
        self.btn_start.setObjectName("primary")
        self.btn_start.clicked.connect(self.start_requested.emit)

        self.btn_stop = QPushButton("Stop")
        self.btn_stop.setObjectName("danger")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_requested.emit)

        self.btn_clear_log = QPushButton("Clear Log")
        self.btn_clear_log.clicked.connect(self._clear_log)

        btn_layout.addWidget(self.btn_start)
        btn_layout.addWidget(self.btn_stop)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_clear_log)

        layout.addLayout(btn_layout)

        # Log output
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #d4d4d4;
                font-family: "Consolas", "Courier New", monospace;
                font-size: 9pt;
            }
        """)
        layout.addWidget(self.log_text)

    # ── Public API ────────────────────────────────────────────

    def set_processing_state(self, processing: bool):
        """Update UI for processing/idle state"""
        self.btn_start.setEnabled(not processing)
        self.btn_stop.setEnabled(processing)

    def update_progress(self, percent: float, message: str):
        """Update progress bar and status label"""
        self.progress_bar.setValue(int(percent))
        self.status_label.setText(message)

    def log_message(self, message: str, level: str = 'normal'):
        """Append a colored message to the log"""
        color_map = {
            'normal': '#d4d4d4',
            'info': '#569cd6',
            'success': '#4ec9b0',
            'warning': '#dcdcaa',
            'error': '#f44747',
        }
        color = color_map.get(level, '#d4d4d4')

        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertHtml(f'<span style="color:{color};">{message}</span><br>')
        self.log_text.setTextCursor(cursor)
        self.log_text.ensureCursorVisible()

    def reset(self):
        """Reset to idle state"""
        self.progress_bar.setValue(0)
        self.status_label.setText("Ready to process")
        self.set_processing_state(False)

    def _clear_log(self):
        self.log_text.clear()
