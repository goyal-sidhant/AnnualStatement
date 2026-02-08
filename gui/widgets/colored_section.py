"""
Colored section widget - replicates the old Tkinter create_colored_section.
A white card with a colored header banner.
"""

from PyQt5.QtWidgets import QFrame, QVBoxLayout, QLabel, QWidget
from PyQt5.QtCore import Qt


class ColoredSection(QFrame):
    """A card widget with a colored header banner and white body."""

    def __init__(self, title: str, description: str, color: str, parent=None):
        super().__init__(parent)
        self.setFrameStyle(QFrame.Box | QFrame.Plain)
        self.setLineWidth(1)
        self.setStyleSheet(f"""
            ColoredSection {{
                background: white;
                border: 2px solid {color};
                border-radius: 6px;
            }}
        """)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # Colored header banner
        header = QLabel(title)
        header.setAlignment(Qt.AlignCenter)
        header.setStyleSheet(f"""
            QLabel {{
                background-color: {color};
                color: white;
                font-size: 12pt;
                font-weight: bold;
                padding: 8px 12px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }}
        """)
        outer.addWidget(header)

        # Description
        if description:
            desc_label = QLabel(description)
            desc_label.setAlignment(Qt.AlignCenter)
            desc_label.setWordWrap(True)
            desc_label.setStyleSheet("""
                QLabel {
                    background: white;
                    color: #323130;
                    font-size: 10pt;
                    padding: 8px 12px 4px 12px;
                }
            """)
            outer.addWidget(desc_label)

        # Content area
        self._content = QWidget()
        self._content.setStyleSheet("QWidget { background: white; }")
        self._content_layout = QVBoxLayout(self._content)
        self._content_layout.setContentsMargins(12, 8, 12, 12)
        outer.addWidget(self._content)

    def content_layout(self) -> QVBoxLayout:
        """Get the layout for adding child widgets."""
        return self._content_layout

    def addWidget(self, widget):
        self._content_layout.addWidget(widget)

    def addLayout(self, layout):
        self._content_layout.addLayout(layout)
