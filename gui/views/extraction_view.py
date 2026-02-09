"""
Extraction View - Tab 4: PQ refresh + cell extraction.
Tree view driven by disk scanning: Client (parent) → Version (child).
"""

from pathlib import Path
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QGroupBox, QSpinBox, QLineEdit, QCheckBox, QProgressBar,
    QTextEdit, QTreeView, QHeaderView, QFrame
)
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QTextCursor, QColor, QFont
from PyQt5.QtCore import Qt, pyqtSignal


class ExtractionView(QWidget):
    """Tab 4: PQ refresh and data extraction"""

    start_extraction_requested = pyqtSignal()
    stop_extraction_requested = pyqtSignal()
    refresh_list_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._level1_folder = None
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        # Header banner
        header = QLabel("EXTRACT & REFRESH")
        header.setAlignment(Qt.AlignCenter)
        header.setStyleSheet("""
            QLabel {
                background-color: #0099BC;
                color: white;
                font-size: 13pt;
                font-weight: bold;
                padding: 8px;
                border-radius: 4px;
            }
        """)
        layout.addWidget(header)

        # Top: options + client list
        top_layout = QHBoxLayout()
        top_layout.addWidget(self._build_options_group())
        top_layout.addWidget(self._build_client_list_group(), stretch=1)
        layout.addLayout(top_layout)

        # Control buttons
        layout.addWidget(self._build_control_buttons())

        # Progress
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #cccccc;
                border-radius: 4px;
                text-align: center;
                height: 26px;
                font-weight: bold;
                background: #f0f0f0;
            }
            QProgressBar::chunk {
                background-color: #0099BC;
                border-radius: 3px;
            }
        """)
        layout.addWidget(self.progress_bar)

        self.status_label = QLabel("Ready for extraction")
        self.status_label.setStyleSheet("""
            QLabel {
                color: #323130;
                font-size: 10pt;
                font-weight: bold;
                padding: 2px 0px;
            }
        """)
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
                border: 2px solid #3c3c3c;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        layout.addWidget(self.log_text)

    def _build_options_group(self):
        group = QGroupBox("Extraction Options")
        group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #0099BC;
                border-radius: 6px;
                margin-top: 8px;
                padding-top: 16px;
            }
            QGroupBox::title {
                color: #0099BC;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
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
        group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #0099BC;
                border-radius: 6px;
                margin-top: 8px;
                padding-top: 16px;
            }
            QGroupBox::title {
                color: #0099BC;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        layout = QVBoxLayout(group)

        # Level 1 folder label + refresh button
        source_row = QHBoxLayout()
        self._level1_label = QLabel("No data loaded")
        self._level1_label.setStyleSheet("""
            QLabel {
                color: #0078D4;
                font-size: 9pt;
                padding: 2px 4px;
                background: #E8F4FD;
                border-radius: 3px;
            }
        """)
        source_row.addWidget(self._level1_label, stretch=1)

        btn_refresh = QPushButton("Refresh List")
        btn_refresh.setCursor(Qt.PointingHandCursor)
        btn_refresh.setStyleSheet("""
            QPushButton {
                background: #0078D4;
                color: white;
                font-size: 8pt;
                padding: 4px 12px;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover { background: #106EBE; }
        """)
        btn_refresh.clicked.connect(self.refresh_list_requested.emit)
        source_row.addWidget(btn_refresh)
        layout.addLayout(source_row)

        # Selection buttons
        btn_row = QHBoxLayout()

        self.btn_itc_only = QPushButton("ITC Only")
        self.btn_itc_only.setObjectName("primary")
        self.btn_itc_only.setCursor(Qt.PointingHandCursor)
        self.btn_itc_only.clicked.connect(self._select_itc_only)

        self.btn_sales_only = QPushButton("Sales Only")
        self.btn_sales_only.setObjectName("warning")
        self.btn_sales_only.setCursor(Qt.PointingHandCursor)
        self.btn_sales_only.clicked.connect(self._select_sales_only)

        self.btn_select_all = QPushButton("Select All")
        self.btn_select_all.setObjectName("success")
        self.btn_select_all.setCursor(Qt.PointingHandCursor)
        self.btn_select_all.clicked.connect(self._select_all)

        self.btn_deselect_all = QPushButton("Deselect All")
        self.btn_deselect_all.setCursor(Qt.PointingHandCursor)
        self.btn_deselect_all.clicked.connect(self._deselect_all)

        btn_row.addWidget(self.btn_itc_only)
        btn_row.addWidget(self.btn_sales_only)
        btn_row.addWidget(self.btn_select_all)
        btn_row.addWidget(self.btn_deselect_all)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # Client tree (parent/child: Client → Version)
        self.client_tree = QTreeView()
        self.client_tree.setAlternatingRowColors(True)
        self.client_tree.setRootIsDecorated(True)

        self.client_model = QStandardItemModel()
        self.client_model.setHorizontalHeaderLabels([
            'Name', 'ITC', 'Sales', 'Refresh Status'
        ])
        self.client_tree.setModel(self.client_model)

        header = self.client_tree.header()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Fixed)
        header.resizeSection(1, 50)
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        header.resizeSection(2, 50)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)

        layout.addWidget(self.client_tree)

        return group

    def _build_control_buttons(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 4, 0, 4)

        self.btn_start = QPushButton("Start Extraction")
        self.btn_start.setCursor(Qt.PointingHandCursor)
        self.btn_start.setStyleSheet("""
            QPushButton {
                background: #0099BC;
                color: white;
                font-size: 11pt;
                font-weight: bold;
                padding: 8px 24px;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover { background: #0088a8; }
            QPushButton:pressed { background: #007794; }
            QPushButton:disabled { background: #a0a0a0; }
        """)
        self.btn_start.clicked.connect(self.start_extraction_requested.emit)

        self.btn_stop = QPushButton("Stop")
        self.btn_stop.setObjectName("danger")
        self.btn_stop.setCursor(Qt.PointingHandCursor)
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_extraction_requested.emit)

        self.btn_clear_log = QPushButton("Clear Log")
        self.btn_clear_log.setCursor(Qt.PointingHandCursor)
        self.btn_clear_log.clicked.connect(lambda: self.log_text.clear())

        layout.addWidget(self.btn_start)
        layout.addWidget(self.btn_stop)
        layout.addStretch()
        layout.addWidget(self.btn_clear_log)

        return widget

    # ── Public API ────────────────────────────────────────────

    def populate_tree(self, tree_data, level1_folder):
        """Populate tree with client→version hierarchy from disk scan.

        Args:
            tree_data: List of dicts with 'name', 'path', 'versions' keys.
            level1_folder: Path to the Level 1 folder being displayed.
        """
        self.client_model.removeRows(0, self.client_model.rowCount())
        self._level1_folder = level1_folder
        self._level1_label.setText(f"Source: {Path(level1_folder).name}")

        bold_font = QFont()
        bold_font.setBold(True)

        for client in tree_data:
            # Parent row: client folder name
            client_item = QStandardItem(client['name'])
            client_item.setEditable(False)
            client_item.setFont(bold_font)
            client_item.setData(str(client['path']), Qt.UserRole)

            # Placeholder columns for parent row
            parent_itc = QStandardItem("")
            parent_itc.setEditable(False)
            parent_sales = QStandardItem("")
            parent_sales.setEditable(False)
            parent_status = QStandardItem(f"{len(client['versions'])} version(s)")
            parent_status.setEditable(False)
            parent_status.setForeground(QColor('#666666'))

            # Add version child rows
            for version in client['versions']:
                ver_name = QStandardItem(f"  {version['name']}")
                ver_name.setEditable(False)
                ver_name.setData(str(version['path']), Qt.UserRole)

                ver_itc = QStandardItem()
                ver_itc.setCheckable(version['has_itc'])
                ver_itc.setCheckState(Qt.Unchecked)
                ver_itc.setEditable(False)
                if not version['has_itc']:
                    ver_itc.setForeground(QColor('#cccccc'))

                ver_sales = QStandardItem()
                ver_sales.setCheckable(version['has_sales'])
                ver_sales.setCheckState(Qt.Unchecked)
                ver_sales.setEditable(False)
                if not version['has_sales']:
                    ver_sales.setForeground(QColor('#cccccc'))

                status_text = version['refresh_status']
                if version['refresh_timestamp']:
                    status_text += f" ({version['refresh_timestamp']})"
                ver_status = QStandardItem(status_text)
                ver_status.setEditable(False)

                if version['refresh_status'] == 'Refreshed':
                    ver_status.setForeground(QColor('#107C10'))
                else:
                    ver_status.setForeground(QColor('#FF8C00'))

                client_item.appendRow([ver_name, ver_itc, ver_sales, ver_status])

            self.client_model.appendRow([client_item, parent_itc, parent_sales, parent_status])

        self.client_tree.expandAll()

    def show_no_data_message(self, message):
        """Show a message when no data is available."""
        self.client_model.removeRows(0, self.client_model.rowCount())
        self._level1_folder = None
        self._level1_label.setText(message)

    def get_level1_folder(self):
        """Return the stored Level 1 folder path."""
        return str(self._level1_folder) if self._level1_folder else ''

    def get_selected_clients(self):
        """Get list of selected versions with their ITC/Sales selections."""
        selected = []
        root = self.client_model.invisibleRootItem()
        for client_row in range(root.rowCount()):
            client_item = root.child(client_row, 0)
            for ver_row in range(client_item.rowCount()):
                ver_name_item = client_item.child(ver_row, 0)
                ver_itc_item = client_item.child(ver_row, 1)
                ver_sales_item = client_item.child(ver_row, 2)

                process_itc = ver_itc_item.checkState() == Qt.Checked if ver_itc_item else False
                process_sales = ver_sales_item.checkState() == Qt.Checked if ver_sales_item else False

                if process_itc or process_sales:
                    selected.append({
                        'name': f"{client_item.text()} / {ver_name_item.text().strip()}",
                        'client_key': client_item.text(),
                        'version_folder': ver_name_item.data(Qt.UserRole) or '',
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

    # ── Selection helpers (parent→child iteration) ────────────

    def _iter_version_items(self):
        """Yield (itc_item, sales_item) for every version child row."""
        root = self.client_model.invisibleRootItem()
        for client_row in range(root.rowCount()):
            client_item = root.child(client_row, 0)
            for ver_row in range(client_item.rowCount()):
                itc = client_item.child(ver_row, 1)
                sales = client_item.child(ver_row, 2)
                yield itc, sales

    def _select_all(self):
        for itc, sales in self._iter_version_items():
            if itc and itc.isCheckable():
                itc.setCheckState(Qt.Checked)
            if sales and sales.isCheckable():
                sales.setCheckState(Qt.Checked)

    def _deselect_all(self):
        for itc, sales in self._iter_version_items():
            if itc and itc.isCheckable():
                itc.setCheckState(Qt.Unchecked)
            if sales and sales.isCheckable():
                sales.setCheckState(Qt.Unchecked)

    def _select_itc_only(self):
        for itc, sales in self._iter_version_items():
            if itc and itc.isCheckable():
                itc.setCheckState(Qt.Checked)
            if sales and sales.isCheckable():
                sales.setCheckState(Qt.Unchecked)

    def _select_sales_only(self):
        for itc, sales in self._iter_version_items():
            if itc and itc.isCheckable():
                itc.setCheckState(Qt.Unchecked)
            if sales and sales.isCheckable():
                sales.setCheckState(Qt.Checked)
