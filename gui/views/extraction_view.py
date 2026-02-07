"""
Extraction View - Tab 4: PQ refresh + cell extraction (merged PQ Extractor).
Auto-populates from Step 3 processing results.
"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QSpinBox, QLineEdit, QCheckBox, QProgressBar,
    QTextEdit, QTreeView, QHeaderView, QFrame
)
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QTextCursor
from PyQt5.QtCore import Qt, pyqtSignal


class ExtractionView(QWidget):
    """Tab 4: PQ refresh and data extraction"""

    start_extraction_requested = pyqtSignal()
    stop_extraction_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)

        # Top: options + client list
        top_layout = QHBoxLayout()
        top_layout.addWidget(self._build_options_group())
        top_layout.addWidget(self._build_client_list_group(), stretch=1)
        layout.addLayout(top_layout)

        # Control buttons
        layout.addWidget(self._build_control_buttons())

        # Progress
        progress_layout = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        layout.addLayout(progress_layout)

        self.status_label = QLabel("Ready for extraction")
        layout.addWidget(self.status_label)

        # Log
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

    def _build_options_group(self):
        group = QGroupBox("Extraction Options")
        layout = QVBoxLayout(group)

        # Wait time
        row = QHBoxLayout()
        row.addWidget(QLabel("Refresh wait time (seconds):"))
        self.spin_wait_time = QSpinBox()
        self.spin_wait_time.setRange(5, 120)
        self.spin_wait_time.setValue(10)
        row.addWidget(self.spin_wait_time)
        row.addStretch()
        layout.addLayout(row)

        # Suffix pattern
        row = QHBoxLayout()
        row.addWidget(QLabel("File suffix:"))
        self.suffix_input = QLineEdit("_Refreshed_{timestamp}")
        row.addWidget(self.suffix_input)
        layout.addLayout(row)

        # Skip refresh
        self.chk_skip_refresh = QCheckBox("Skip refresh (use existing refreshed files)")
        layout.addWidget(self.chk_skip_refresh)

        layout.addStretch()
        return group

    def _build_client_list_group(self):
        group = QGroupBox("Clients")
        layout = QVBoxLayout(group)

        # Selection buttons
        btn_row = QHBoxLayout()

        self.btn_itc_only = QPushButton("ITC Only")
        self.btn_itc_only.clicked.connect(self._select_itc_only)
        self.btn_sales_only = QPushButton("Sales Only")
        self.btn_sales_only.clicked.connect(self._select_sales_only)
        self.btn_select_all = QPushButton("Select All")
        self.btn_select_all.clicked.connect(self._select_all)
        self.btn_deselect_all = QPushButton("Deselect All")
        self.btn_deselect_all.clicked.connect(self._deselect_all)

        btn_row.addWidget(self.btn_itc_only)
        btn_row.addWidget(self.btn_sales_only)
        btn_row.addWidget(self.btn_select_all)
        btn_row.addWidget(self.btn_deselect_all)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # Client tree
        self.client_tree = QTreeView()
        self.client_tree.setAlternatingRowColors(True)
        self.client_tree.setRootIsDecorated(False)

        self.client_model = QStandardItemModel()
        self.client_model.setHorizontalHeaderLabels([
            'Client', 'ITC', 'Sales', 'Last Refresh'
        ])
        self.client_tree.setModel(self.client_model)

        header = self.client_tree.header()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        header.resizeSection(1, 50)
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        header.resizeSection(2, 50)

        layout.addWidget(self.client_tree)

        return group

    def _build_control_buttons(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 4, 0, 4)

        self.btn_start = QPushButton("Start Extraction")
        self.btn_start.setObjectName("primary")
        self.btn_start.clicked.connect(self.start_extraction_requested.emit)

        self.btn_stop = QPushButton("Stop")
        self.btn_stop.setObjectName("danger")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_extraction_requested.emit)

        self.btn_clear_log = QPushButton("Clear Log")
        self.btn_clear_log.clicked.connect(lambda: self.log_text.clear())

        layout.addWidget(self.btn_start)
        layout.addWidget(self.btn_stop)
        layout.addStretch()
        layout.addWidget(self.btn_clear_log)

        return widget

    # ── Public API ────────────────────────────────────────────

    def populate_clients(self, clients, run_history=None):
        """Populate client list from processing results or scan data.

        Args:
            clients: list of dicts with keys: name, client_key, has_itc, has_sales
            run_history: optional RunHistory for last refresh timestamps
        """
        self.client_model.removeRows(0, self.client_model.rowCount())

        for client in clients:
            name_item = QStandardItem(client.get('name', ''))
            name_item.setData(client.get('client_key', ''), Qt.UserRole)
            name_item.setEditable(False)

            itc_item = QStandardItem()
            itc_item.setCheckable(True)
            itc_item.setCheckState(Qt.Checked if client.get('has_itc', False) else Qt.Unchecked)
            itc_item.setEditable(False)

            sales_item = QStandardItem()
            sales_item.setCheckable(True)
            sales_item.setCheckState(Qt.Checked if client.get('has_sales', False) else Qt.Unchecked)
            sales_item.setEditable(False)

            last_refresh = ''
            if run_history:
                record = run_history.get_last_run_for_client(client.get('client_key', ''))
                if record and record.get('pq_refresh_timestamp'):
                    last_refresh = record['pq_refresh_timestamp']
            refresh_item = QStandardItem(last_refresh)
            refresh_item.setEditable(False)

            self.client_model.appendRow([name_item, itc_item, sales_item, refresh_item])

    def get_selected_clients(self):
        """Get list of selected clients with their ITC/Sales selections."""
        selected = []
        for row in range(self.client_model.rowCount()):
            name_item = self.client_model.item(row, 0)
            itc_item = self.client_model.item(row, 1)
            sales_item = self.client_model.item(row, 2)

            process_itc = itc_item.checkState() == Qt.Checked if itc_item else False
            process_sales = sales_item.checkState() == Qt.Checked if sales_item else False

            if process_itc or process_sales:
                selected.append({
                    'client_key': name_item.data(Qt.UserRole),
                    'name': name_item.text(),
                    'process_itc': process_itc,
                    'process_sales': process_sales,
                })
        return selected

    def set_extraction_state(self, extracting: bool):
        """Update UI for extracting/idle state"""
        self.btn_start.setEnabled(not extracting)
        self.btn_stop.setEnabled(extracting)

    def update_progress(self, percent: float, message: str):
        """Update progress bar and status"""
        self.progress_bar.setValue(int(percent))
        self.status_label.setText(message)

    def log_message(self, message: str, level: str = 'normal'):
        """Append message to log"""
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

    # ── Selection helpers ─────────────────────────────────────

    def _select_all(self):
        for row in range(self.client_model.rowCount()):
            for col in [1, 2]:
                item = self.client_model.item(row, col)
                if item:
                    item.setCheckState(Qt.Checked)

    def _deselect_all(self):
        for row in range(self.client_model.rowCount()):
            for col in [1, 2]:
                item = self.client_model.item(row, col)
                if item:
                    item.setCheckState(Qt.Unchecked)

    def _select_itc_only(self):
        for row in range(self.client_model.rowCount()):
            itc = self.client_model.item(row, 1)
            sales = self.client_model.item(row, 2)
            if itc:
                itc.setCheckState(Qt.Checked)
            if sales:
                sales.setCheckState(Qt.Unchecked)

    def _select_sales_only(self):
        for row in range(self.client_model.rowCount()):
            itc = self.client_model.item(row, 1)
            sales = self.client_model.item(row, 2)
            if itc:
                itc.setCheckState(Qt.Unchecked)
            if sales:
                sales.setCheckState(Qt.Checked)
