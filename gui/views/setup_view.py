"""
Setup View - Tab 1: Source folder, template, and target selection.
"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QGroupBox, QRadioButton, QButtonGroup, QCheckBox, QSpinBox,
    QFileDialog, QTextEdit, QFrame
)
from PyQt5.QtCore import pyqtSignal


class SetupView(QWidget):
    """Tab 1: folder/template selection and processing options"""

    # Signals emitted by this view
    scan_requested = pyqtSignal()
    rescan_requested = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)

        # Left column - inputs
        left = QVBoxLayout()
        left.addWidget(self._build_file_selection_group())
        left.addWidget(self._build_processing_options_group())
        left.addWidget(self._build_action_buttons())
        left.addStretch()

        # Right column - help
        right = QVBoxLayout()
        right.addWidget(self._build_help_panel())
        right.addStretch()

        layout.addLayout(left, stretch=3)
        layout.addLayout(right, stretch=2)

    def _build_file_selection_group(self):
        group = QGroupBox("File Selection")
        layout = QVBoxLayout(group)

        # Source folder
        layout.addWidget(QLabel("Source Folder (contains GST Excel files):"))
        row = QHBoxLayout()
        self.source_folder = QLineEdit()
        self.source_folder.setPlaceholderText("Select folder containing GST files...")
        self.btn_browse_source = QPushButton("Browse...")
        self.btn_browse_source.clicked.connect(self._browse_source)
        row.addWidget(self.source_folder)
        row.addWidget(self.btn_browse_source)
        layout.addLayout(row)

        # ITC Template
        layout.addWidget(QLabel("ITC Report Template:"))
        row = QHBoxLayout()
        self.itc_template = QLineEdit()
        self.itc_template.setPlaceholderText("Select ITC Excel template...")
        self.btn_browse_itc = QPushButton("Browse...")
        self.btn_browse_itc.clicked.connect(self._browse_itc)
        row.addWidget(self.itc_template)
        row.addWidget(self.btn_browse_itc)
        layout.addLayout(row)

        # Sales Template
        layout.addWidget(QLabel("Sales Report Template:"))
        row = QHBoxLayout()
        self.sales_template = QLineEdit()
        self.sales_template.setPlaceholderText("Select Sales Excel template...")
        self.btn_browse_sales = QPushButton("Browse...")
        self.btn_browse_sales.clicked.connect(self._browse_sales)
        row.addWidget(self.sales_template)
        row.addWidget(self.btn_browse_sales)
        layout.addLayout(row)

        # Target folder
        layout.addWidget(QLabel("Target Folder (output location):"))
        row = QHBoxLayout()
        self.target_folder = QLineEdit()
        self.target_folder.setPlaceholderText("Select target folder for organized files...")
        self.btn_browse_target = QPushButton("Browse...")
        self.btn_browse_target.clicked.connect(self._browse_target)
        row.addWidget(self.target_folder)
        row.addWidget(self.btn_browse_target)
        layout.addLayout(row)

        return group

    def _build_processing_options_group(self):
        group = QGroupBox("Processing Options")
        layout = QVBoxLayout(group)

        # Processing mode
        mode_label = QLabel("Processing Mode:")
        layout.addWidget(mode_label)

        self.mode_group = QButtonGroup(self)
        self.radio_fresh = QRadioButton("Fresh - Create new folder structure")
        self.radio_rerun = QRadioButton("Re-Run - Overwrite existing files")
        self.radio_resume = QRadioButton("Resume - Skip already processed")
        self.radio_fresh.setChecked(True)

        self.mode_group.addButton(self.radio_fresh, 0)
        self.mode_group.addButton(self.radio_rerun, 1)
        self.mode_group.addButton(self.radio_resume, 2)

        layout.addWidget(self.radio_fresh)
        layout.addWidget(self.radio_rerun)
        layout.addWidget(self.radio_resume)

        # Separator
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        layout.addWidget(line)

        # Client name in folders
        row = QHBoxLayout()
        self.chk_client_name = QCheckBox("Include client name in folder names")
        row.addWidget(self.chk_client_name)

        row.addWidget(QLabel("Max length:"))
        self.spin_max_length = QSpinBox()
        self.spin_max_length.setRange(5, 100)
        self.spin_max_length.setValue(35)
        self.spin_max_length.setEnabled(False)
        row.addWidget(self.spin_max_length)
        row.addStretch()

        self.chk_client_name.toggled.connect(self.spin_max_length.setEnabled)
        layout.addLayout(row)

        return group

    def _build_action_buttons(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(0, 8, 0, 0)

        self.btn_scan = QPushButton("Scan Files")
        self.btn_scan.setObjectName("primary")
        self.btn_scan.clicked.connect(self.scan_requested.emit)

        self.btn_rescan = QPushButton("Re-Scan")
        self.btn_rescan.clicked.connect(self.rescan_requested.emit)
        self.btn_rescan.setEnabled(False)

        layout.addWidget(self.btn_scan)
        layout.addWidget(self.btn_rescan)
        layout.addStretch()

        return widget

    def _build_help_panel(self):
        group = QGroupBox("Help & File Patterns")
        layout = QVBoxLayout(group)

        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setHtml("""
        <h3>Expected File Naming Patterns</h3>
        <p>Place all GST files in a single folder. The app recognizes:</p>
        <ul>
            <li><b>GSTR-2B Reco:</b> GSTR-2B-Reco-ClientName-State-Period.xlsx</li>
            <li><b>IMS Reco:</b> ImsReco-ClientName-State-DDMMYYYY.xlsx</li>
            <li><b>GSTR-3B:</b> GSTR3B-ClientName-State-Month.xlsx</li>
            <li><b>Sales:</b> Sales-ClientName-State-StartMonth-EndMonth.xlsx</li>
            <li><b>Sales Reco:</b> SalesReco-ClientName-State-Period.xlsx</li>
            <li><b>Annual Report:</b> AnnualReport-ClientName-State-Year.xlsx</li>
        </ul>
        <h3>Processing Modes</h3>
        <ul>
            <li><b>Fresh:</b> Creates a new timestamped folder structure</li>
            <li><b>Re-Run:</b> Overwrites files in existing structure</li>
            <li><b>Resume:</b> Skips clients already processed</li>
        </ul>
        <h3>Steps</h3>
        <ol>
            <li>Select source folder with GST files</li>
            <li>Select ITC and Sales Excel templates</li>
            <li>Select target output folder</li>
            <li>Click "Scan Files" to analyze</li>
            <li>Review clients in Validation tab</li>
            <li>Start processing in Processing tab</li>
        </ol>
        """)
        layout.addWidget(help_text)
        return group

    # ── Browse handlers ───────────────────────────────────────

    def _browse_source(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Source Folder")
        if folder:
            self.source_folder.setText(folder)

    def _browse_itc(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select ITC Template", "", "Excel Files (*.xlsx *.xlsm)"
        )
        if path:
            self.itc_template.setText(path)

    def _browse_sales(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Sales Template", "", "Excel Files (*.xlsx *.xlsm)"
        )
        if path:
            self.sales_template.setText(path)

    def _browse_target(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Target Folder")
        if folder:
            self.target_folder.setText(folder)

    # ── Public API ────────────────────────────────────────────

    def get_processing_mode(self) -> str:
        """Get selected processing mode"""
        checked = self.mode_group.checkedId()
        return {0: 'fresh', 1: 'rerun', 2: 'resume'}.get(checked, 'fresh')

    def set_values(self, source='', itc='', sales='', target='',
                   mode='fresh', include_name=False, max_length=35):
        """Populate fields from settings"""
        self.source_folder.setText(source)
        self.itc_template.setText(itc)
        self.sales_template.setText(sales)
        self.target_folder.setText(target)

        mode_map = {'fresh': self.radio_fresh, 'rerun': self.radio_rerun, 'resume': self.radio_resume}
        btn = mode_map.get(mode, self.radio_fresh)
        btn.setChecked(True)

        self.chk_client_name.setChecked(include_name)
        self.spin_max_length.setValue(max_length)
