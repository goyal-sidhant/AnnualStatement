"""
Validation View - Tab 2: Client tree, selection, and scan summary.
Two-column layout matching the original Tkinter UI.
"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTreeView, QTextEdit, QHeaderView, QFrame, QScrollArea, QSplitter
)
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QColor
from PyQt5.QtCore import Qt, pyqtSignal
from gui.widgets.colored_section import ColoredSection


class ValidationView(QWidget):
    """Tab 2: client tree + selection controls (two-column layout)"""

    dry_run_requested = pyqtSignal()
    start_processing_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        # Horizontal splitter for resizable left/right columns
        splitter = QSplitter(Qt.Horizontal)
        splitter.setStyleSheet("""
            QSplitter { background: #F8F9FA; }
            QSplitter::handle { background: #cccccc; width: 3px; }
            QSplitter::handle:hover { background: #0078D4; }
        """)

        # Left panel (scrollable) - instructions, summary, actions
        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.NoFrame)
        left_scroll.setStyleSheet("QScrollArea { background: #F8F9FA; border: none; }")

        left_widget = QWidget()
        left_widget.setStyleSheet("QWidget { background: #F8F9FA; }")
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(20, 20, 10, 20)
        left_layout.setSpacing(12)
        left_layout.addWidget(self._build_instructions_section())
        left_layout.addWidget(self._build_summary_section())
        left_layout.addWidget(self._build_action_section())
        left_layout.addStretch()

        left_scroll.setWidget(left_widget)

        # Right panel - client selection tree
        right_widget = QWidget()
        right_widget.setStyleSheet("QWidget { background: #F8F9FA; }")
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(10, 20, 20, 20)
        right_layout.setSpacing(0)
        right_layout.addWidget(self._build_client_selection_section())

        splitter.addWidget(left_scroll)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 3)

        outer.addWidget(splitter)

    # ── Left Column ────────────────────────────────────────────

    def _build_instructions_section(self):
        frame = QFrame()
        frame.setStyleSheet("""
            QFrame {
                background: #FFF3E0;
                border: 2px solid #FFB74D;
                border-radius: 6px;
            }
        """)
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(15, 12, 15, 12)

        title = QLabel("VALIDATION INSTRUCTIONS")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                color: #E65100;
                font-size: 12pt;
                font-weight: bold;
                background: #FFF3E0;
            }
        """)
        layout.addWidget(title)

        instructions = [
            "Review the scan summary below",
            "Check client list - green = complete, orange = missing files",
            "Select clients using checkboxes",
            "Try 'Dry Run' to preview without processing",
            "Click 'Start Processing' when ready",
        ]
        for text in instructions:
            lbl = QLabel(f"  {text}")
            lbl.setStyleSheet("""
                QLabel {
                    color: #E65100;
                    font-size: 10pt;
                    background: #FFF3E0;
                    padding: 2px 0px;
                }
            """)
            layout.addWidget(lbl)

        return frame

    def _build_summary_section(self):
        section = ColoredSection(
            "SCAN SUMMARY",
            "",
            "#0078D4"  # primary blue
        )

        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        self.summary_text.setMinimumHeight(160)
        self.summary_text.setPlaceholderText("Scan summary will appear here after scanning files...")
        self.summary_text.setStyleSheet("""
            QTextEdit {
                background: #f8f9fa;
                color: #323130;
                font-family: "Consolas", "Courier New", monospace;
                font-size: 9pt;
                border: 1px solid #cccccc;
                border-radius: 3px;
            }
        """)
        section.addWidget(self.summary_text)

        return section

    def _build_action_section(self):
        frame = QFrame()
        frame.setStyleSheet("""
            QFrame {
                background: white;
                border: 2px solid #107C10;
                border-radius: 6px;
            }
        """)
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(20, 15, 20, 15)

        title = QLabel("READY TO PROCESS?")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                color: #107C10;
                font-size: 14pt;
                font-weight: bold;
                background: white;
            }
        """)
        layout.addWidget(title)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        self.btn_dry_run = QPushButton("DRY RUN")
        self.btn_dry_run.setCursor(Qt.PointingHandCursor)
        self.btn_dry_run.setStyleSheet("""
            QPushButton {
                background: #FF8C00;
                color: white;
                font-weight: bold;
                font-size: 10pt;
                padding: 6px 15px;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover { background: #e07d00; }
            QPushButton:pressed { background: #c46e00; }
        """)
        self.btn_dry_run.clicked.connect(self.dry_run_requested.emit)

        self.btn_export = QPushButton("EXPORT LIST")
        self.btn_export.setCursor(Qt.PointingHandCursor)
        self.btn_export.setStyleSheet("""
            QPushButton {
                background: #0099BC;
                color: white;
                font-weight: bold;
                font-size: 10pt;
                padding: 6px 15px;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover { background: #0088a8; }
            QPushButton:pressed { background: #007794; }
        """)

        self.btn_start = QPushButton("START")
        self.btn_start.setCursor(Qt.PointingHandCursor)
        self.btn_start.setStyleSheet("""
            QPushButton {
                background: #107C10;
                color: white;
                font-weight: bold;
                font-size: 10pt;
                padding: 6px 15px;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover { background: #0e6b0e; }
            QPushButton:pressed { background: #0c5c0c; }
        """)
        self.btn_start.clicked.connect(self.start_processing_requested.emit)

        btn_row.addStretch()
        btn_row.addWidget(self.btn_dry_run)
        btn_row.addWidget(self.btn_export)
        btn_row.addWidget(self.btn_start)
        btn_row.addStretch()

        layout.addLayout(btn_row)

        return frame

    # ── Right Column ───────────────────────────────────────────

    def _build_client_selection_section(self):
        frame = QFrame()
        frame.setStyleSheet("""
            QFrame#clientSection {
                background: white;
                border: 2px solid #107C10;
                border-radius: 6px;
            }
        """)
        frame.setObjectName("clientSection")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Green header banner
        header = QLabel("CLIENT SELECTION")
        header.setAlignment(Qt.AlignCenter)
        header.setStyleSheet("""
            QLabel {
                background-color: #107C10;
                color: white;
                font-size: 12pt;
                font-weight: bold;
                padding: 8px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
        """)
        layout.addWidget(header)

        # Keyboard hint bar
        hint = QLabel("Use checkboxes to select clients, or use the buttons below")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet("""
            QLabel {
                background: #E3F2FD;
                color: #0078D4;
                font-size: 9pt;
                padding: 6px;
                border: 1px solid #90CAF9;
            }
        """)
        layout.addWidget(hint)

        # Selection buttons
        btn_frame = QWidget()
        btn_frame.setStyleSheet("QWidget { background: white; }")
        btn_row = QHBoxLayout(btn_frame)
        btn_row.setContentsMargins(15, 10, 15, 10)
        btn_row.setSpacing(8)

        self.btn_select_all = QPushButton("Select All")
        self.btn_select_all.setCursor(Qt.PointingHandCursor)
        self.btn_select_all.setStyleSheet(self._sel_btn_style("#107C10", "#0e6b0e"))
        self.btn_select_all.clicked.connect(self._select_all)

        self.btn_clear_all = QPushButton("Clear All")
        self.btn_clear_all.setCursor(Qt.PointingHandCursor)
        self.btn_clear_all.setStyleSheet(self._sel_btn_style("#323130", "#252525"))
        self.btn_clear_all.clicked.connect(self._clear_all)

        self.btn_select_complete = QPushButton("Complete Only")
        self.btn_select_complete.setCursor(Qt.PointingHandCursor)
        self.btn_select_complete.setStyleSheet(self._sel_btn_style("#0078D4", "#006cbd"))
        self.btn_select_complete.clicked.connect(self._select_complete)

        self.lbl_count = QLabel("0 clients found")
        self.lbl_count.setStyleSheet("""
            QLabel {
                color: #366092;
                font-weight: bold;
                font-size: 10pt;
                background: white;
            }
        """)

        btn_row.addWidget(self.btn_select_all)
        btn_row.addWidget(self.btn_clear_all)
        btn_row.addWidget(self.btn_select_complete)
        btn_row.addStretch()
        btn_row.addWidget(self.lbl_count)

        layout.addWidget(btn_frame)

        # Client tree
        tree_container = QWidget()
        tree_container.setStyleSheet("QWidget { background: white; }")
        tree_layout = QVBoxLayout(tree_container)
        tree_layout.setContentsMargins(15, 0, 15, 15)

        self.tree = QTreeView()
        self.tree.setAlternatingRowColors(True)
        self.tree.setRootIsDecorated(False)
        self.tree.setSortingEnabled(True)

        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels([
            'Select', 'Client', 'State', 'Status', 'Files',
            'Missing', 'Extra', 'Folder Name', 'Last Run'
        ])
        self.tree.setModel(self.model)

        header_view = self.tree.header()
        header_view.setStretchLastSection(True)
        header_view.setSectionResizeMode(0, QHeaderView.Fixed)
        header_view.resizeSection(0, 50)
        header_view.setSectionResizeMode(1, QHeaderView.Stretch)

        tree_layout.addWidget(self.tree)
        layout.addWidget(tree_container, stretch=1)

        return frame

    # ── Helpers ─────────────────────────────────────────────────

    @staticmethod
    def _sel_btn_style(bg, hover_bg):
        return f"""
            QPushButton {{
                background: {bg};
                color: white;
                font-weight: bold;
                font-size: 10pt;
                padding: 5px 15px;
                border: none;
                border-radius: 3px;
            }}
            QPushButton:hover {{ background: {hover_bg}; }}
        """

    # ── Public API ────────────────────────────────────────────

    def populate_clients(self, client_data, run_history=None):
        """Populate tree with client data from scan results"""
        self.model.removeRows(0, self.model.rowCount())

        for client_key, info in sorted(client_data.items()):
            select_item = QStandardItem()
            select_item.setCheckable(True)
            select_item.setCheckState(Qt.Checked)
            select_item.setData(client_key, Qt.UserRole)

            client_item = QStandardItem(info.get('client', ''))
            state_item = QStandardItem(info.get('state', ''))
            status_item = QStandardItem(info.get('status', ''))
            files_item = QStandardItem(str(info.get('file_count', 0)))
            missing_item = QStandardItem(str(len(info.get('missing_files', []))))
            extra_item = QStandardItem(str(len(info.get('extra_files', []))))
            folder_item = QStandardItem(client_key)

            # Last run info
            last_run = ''
            if run_history:
                last_run = run_history.get_last_run_timestamp_for_client(client_key)
            last_run_item = QStandardItem(last_run)

            # Color coding by status
            if info.get('status') == 'Complete':
                color = QColor('#107C10')
            elif 'Missing' in info.get('status', ''):
                color = QColor('#FF8C00')
            else:
                color = QColor('#D83B01')

            for item in [client_item, state_item, status_item, files_item,
                         missing_item, extra_item, folder_item, last_run_item]:
                item.setForeground(color)
                item.setEditable(False)

            select_item.setEditable(False)

            self.model.appendRow([
                select_item, client_item, state_item, status_item,
                files_item, missing_item, extra_item, folder_item, last_run_item
            ])

        total = self.model.rowCount()
        complete = sum(1 for i in range(total)
                       if self.model.item(i, 3) and self.model.item(i, 3).text() == 'Complete')
        self.lbl_count.setText(f"{total} clients ({complete} complete)")

    def get_selected_client_keys(self):
        """Get list of selected client keys"""
        selected = []
        for row in range(self.model.rowCount()):
            item = self.model.item(row, 0)
            if item and item.checkState() == Qt.Checked:
                key = item.data(Qt.UserRole)
                if key:
                    selected.append(key)
        return selected

    def set_summary(self, html: str):
        """Set the summary text"""
        self.summary_text.setHtml(html)

    # ── Selection helpers ─────────────────────────────────────

    def _select_all(self):
        for row in range(self.model.rowCount()):
            item = self.model.item(row, 0)
            if item:
                item.setCheckState(Qt.Checked)

    def _clear_all(self):
        for row in range(self.model.rowCount()):
            item = self.model.item(row, 0)
            if item:
                item.setCheckState(Qt.Unchecked)

    def _select_complete(self):
        for row in range(self.model.rowCount()):
            status_item = self.model.item(row, 3)
            check_item = self.model.item(row, 0)
            if status_item and check_item:
                is_complete = status_item.text() == 'Complete'
                check_item.setCheckState(Qt.Checked if is_complete else Qt.Unchecked)
