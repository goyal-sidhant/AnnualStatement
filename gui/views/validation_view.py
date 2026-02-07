"""
Validation View - Tab 2: Client tree, selection, and scan summary.
"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTreeView, QTextEdit, QSplitter, QHeaderView
)
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QColor, QFont
from PyQt5.QtCore import Qt, pyqtSignal


class ValidationView(QWidget):
    """Tab 2: client tree + selection controls"""

    dry_run_requested = pyqtSignal()
    start_processing_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)

        # Selection buttons
        layout.addWidget(self._build_selection_buttons())

        # Splitter: tree (top) + summary (bottom)
        splitter = QSplitter(Qt.Vertical)

        # Client tree
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

        header = self.tree.header()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.resizeSection(0, 50)
        header.setSectionResizeMode(1, QHeaderView.Stretch)

        splitter.addWidget(self.tree)

        # Summary text
        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        self.summary_text.setMaximumHeight(150)
        self.summary_text.setPlaceholderText("Scan summary will appear here after scanning files...")
        splitter.addWidget(self.summary_text)

        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)

        layout.addWidget(splitter)

        # Action buttons
        layout.addWidget(self._build_action_buttons())

    def _build_selection_buttons(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 4)

        self.btn_select_all = QPushButton("Select All")
        self.btn_select_all.clicked.connect(self._select_all)

        self.btn_clear_all = QPushButton("Clear All")
        self.btn_clear_all.clicked.connect(self._clear_all)

        self.btn_select_complete = QPushButton("Select Complete Only")
        self.btn_select_complete.clicked.connect(self._select_complete)

        self.lbl_count = QLabel("0 clients found")

        layout.addWidget(self.btn_select_all)
        layout.addWidget(self.btn_clear_all)
        layout.addWidget(self.btn_select_complete)
        layout.addStretch()
        layout.addWidget(self.lbl_count)

        return widget

    def _build_action_buttons(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 4, 0, 0)

        self.btn_dry_run = QPushButton("Dry Run")
        self.btn_dry_run.clicked.connect(self.dry_run_requested.emit)

        self.btn_export = QPushButton("Export List")

        self.btn_start = QPushButton("Start Processing")
        self.btn_start.setObjectName("primary")
        self.btn_start.clicked.connect(self.start_processing_requested.emit)

        layout.addWidget(self.btn_dry_run)
        layout.addWidget(self.btn_export)
        layout.addStretch()
        layout.addWidget(self.btn_start)

        return widget

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

            # Color coding
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
