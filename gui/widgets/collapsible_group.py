"""
Collapsible QGroupBox widget - replaces the Tkinter CollapsibleFrame.
"""

from PyQt5.QtWidgets import QGroupBox, QVBoxLayout, QWidget
from PyQt5.QtCore import Qt, QPropertyAnimation, QParallelAnimationGroup


class CollapsibleGroup(QGroupBox):
    """A QGroupBox that can be collapsed/expanded by clicking its title"""

    def __init__(self, title: str = '', parent=None):
        super().__init__(title, parent)
        self.setCheckable(True)
        self.setChecked(True)
        self.toggled.connect(self._on_toggled)

        self._content = QWidget()
        self._layout = QVBoxLayout()
        self._layout.setContentsMargins(8, 4, 8, 4)
        self._content.setLayout(self._layout)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(4, 8, 4, 4)
        main_layout.addWidget(self._content)
        super().setLayout(main_layout)

    def content_layout(self) -> QVBoxLayout:
        """Get the layout where child widgets should be added"""
        return self._layout

    def addWidget(self, widget):
        """Convenience method to add widgets to the content area"""
        self._layout.addWidget(widget)

    def addLayout(self, layout):
        """Convenience method to add layouts to the content area"""
        self._layout.addLayout(layout)

    def _on_toggled(self, checked: bool):
        """Show/hide content when checkbox is toggled"""
        self._content.setVisible(checked)
