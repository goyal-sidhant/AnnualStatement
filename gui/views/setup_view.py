"""
Setup View - Tab 1: Source folder, template, and target selection.
Colorful card-based layout matching the original Tkinter UI.
"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QRadioButton, QButtonGroup, QCheckBox, QSpinBox,
    QFileDialog, QFrame, QScrollArea, QSizePolicy, QSplitter,
    QTextEdit, QGroupBox
)
from PyQt5.QtCore import pyqtSignal, Qt
from PyQt5.QtGui import QTextCursor
from gui.widgets.colored_section import ColoredSection


class CollapsibleCard(QFrame):
    """A colored card with a clickable header that collapses/expands its content."""

    def __init__(self, title, bg_color, border_color, text_color, parent=None):
        super().__init__(parent)
        self._collapsed = False
        self._bg = bg_color
        self._border = border_color
        self._text_color = text_color
        self._title_text = title

        self.setStyleSheet(f"""
            CollapsibleCard {{
                background: {bg_color};
                border: 2px solid {border_color};
                border-radius: 6px;
            }}
        """)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # Clickable header
        self._header = QPushButton(f"  {title}")
        self._header.setCursor(Qt.PointingHandCursor)
        self._header.setStyleSheet(f"""
            QPushButton {{
                background: {border_color};
                color: white;
                font-size: 11pt;
                font-weight: bold;
                padding: 6px 12px;
                border: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                text-align: left;
            }}
            QPushButton:hover {{ background: {text_color}; }}
        """)
        self._header.clicked.connect(self._toggle)
        outer.addWidget(self._header)

        # Content area
        self._content = QWidget()
        self._content.setStyleSheet(f"QWidget {{ background: {bg_color}; }}")
        self._content_layout = QVBoxLayout(self._content)
        self._content_layout.setContentsMargins(15, 8, 15, 12)
        outer.addWidget(self._content)

        self._update_header()

    def _toggle(self):
        self._collapsed = not self._collapsed
        self._content.setVisible(not self._collapsed)
        self._update_header()

    def _update_header(self):
        arrow = ">" if self._collapsed else "v"
        self._header.setText(f"  {arrow}  {self._title_text}")

    def addWidget(self, widget):
        self._content_layout.addWidget(widget)

    def addLayout(self, layout):
        self._content_layout.addLayout(layout)


class SetupView(QWidget):
    """Tab 1: folder/template selection and processing options"""

    scan_requested = pyqtSignal()
    rescan_requested = pyqtSignal()

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

        # Left panel (scrollable) - all input sections
        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.NoFrame)
        left_scroll.setStyleSheet("QScrollArea { background: #F8F9FA; border: none; }")

        left_widget = QWidget()
        left_widget.setStyleSheet("QWidget { background: #F8F9FA; }")
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(20, 20, 10, 20)
        left_layout.setSpacing(12)
        left_layout.addWidget(self._build_welcome_banner())
        left_layout.addWidget(self._build_source_section())
        left_layout.addWidget(self._build_templates_section())
        left_layout.addWidget(self._build_target_section())
        left_layout.addWidget(self._build_processing_mode_section())
        left_layout.addWidget(self._build_action_section())
        left_layout.addStretch()

        left_scroll.setWidget(left_widget)

        # Right panel - collapsible help cards + scan log
        right_widget = QWidget()
        right_widget.setStyleSheet("QWidget { background: #F8F9FA; }")
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(10, 20, 20, 20)
        right_layout.setSpacing(10)

        right_layout.addWidget(self._build_instructions_card())
        right_layout.addWidget(self._build_patterns_card())
        right_layout.addWidget(self._build_scan_log())

        splitter.addWidget(left_scroll)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)

        outer.addWidget(splitter)

    # ── Welcome Banner ─────────────────────────────────────────

    def _build_welcome_banner(self):
        frame = QFrame()
        frame.setStyleSheet("""
            QFrame {
                background: white;
                border: 2px solid #0078D4;
                border-radius: 6px;
            }
        """)
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(20, 15, 20, 15)

        title = QLabel("Welcome to GST File Organizer!")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                color: #0078D4;
                font-size: 16pt;
                font-weight: bold;
                background: white;
            }
        """)
        layout.addWidget(title)

        subtitle = QLabel(
            "This tool helps you organize GST files and create Excel reports automatically.\n"
            "Let's get started!"
        )
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("""
            QLabel {
                color: #323130;
                font-size: 10pt;
                background: white;
            }
        """)
        layout.addWidget(subtitle)

        return frame

    # ── Source Folder Section ──────────────────────────────────

    def _build_source_section(self):
        section = ColoredSection(
            "SOURCE FOLDER",
            "Select the folder containing your GST Excel files",
            "#107C10"
        )

        row = QHBoxLayout()
        self.source_folder = QLineEdit()
        self.source_folder.setPlaceholderText("Select folder containing GST files...")
        self.source_folder.setStyleSheet("QLineEdit { background: white; color: #323130; }")

        self.btn_browse_source = QPushButton("Browse")
        self.btn_browse_source.setMinimumWidth(90)
        self.btn_browse_source.setObjectName("success")
        self.btn_browse_source.setCursor(Qt.PointingHandCursor)
        self.btn_browse_source.clicked.connect(self._browse_source)

        row.addWidget(self.source_folder)
        row.addWidget(self.btn_browse_source)
        section.addLayout(row)

        return section

    # ── Templates Section ──────────────────────────────────────

    def _build_templates_section(self):
        section = ColoredSection(
            "EXCEL TEMPLATES",
            "Select your Excel template files (.xlsx or .xlsm)",
            "#0078D4"
        )

        # ITC template
        lbl = QLabel("ITC Report Template:")
        lbl.setStyleSheet("QLabel { font-weight: bold; background: white; color: #323130; }")
        section.addWidget(lbl)

        row = QHBoxLayout()
        self.itc_template = QLineEdit()
        self.itc_template.setPlaceholderText("Select ITC Excel template...")
        self.itc_template.setStyleSheet("QLineEdit { background: white; color: #323130; }")
        self.btn_browse_itc = QPushButton("Browse")
        self.btn_browse_itc.setMinimumWidth(90)
        self.btn_browse_itc.setObjectName("primary")
        self.btn_browse_itc.setCursor(Qt.PointingHandCursor)
        self.btn_browse_itc.clicked.connect(self._browse_itc)
        row.addWidget(self.itc_template)
        row.addWidget(self.btn_browse_itc)
        section.addLayout(row)

        # Sales template
        lbl2 = QLabel("Sales Report Template:")
        lbl2.setStyleSheet("QLabel { font-weight: bold; background: white; color: #323130; }")
        section.addWidget(lbl2)

        row2 = QHBoxLayout()
        self.sales_template = QLineEdit()
        self.sales_template.setPlaceholderText("Select Sales Excel template...")
        self.sales_template.setStyleSheet("QLineEdit { background: white; color: #323130; }")
        self.btn_browse_sales = QPushButton("Browse")
        self.btn_browse_sales.setMinimumWidth(90)
        self.btn_browse_sales.setObjectName("primary")
        self.btn_browse_sales.setCursor(Qt.PointingHandCursor)
        self.btn_browse_sales.clicked.connect(self._browse_sales)
        row2.addWidget(self.sales_template)
        row2.addWidget(self.btn_browse_sales)
        section.addLayout(row2)

        return section

    # ── Target Folder Section ──────────────────────────────────

    def _build_target_section(self):
        section = ColoredSection(
            "TARGET FOLDER - IMPORTANT!",
            "Select where the organized files will be saved",
            "#D83B01"
        )

        # Emphasis warning
        warn_frame = QFrame()
        warn_frame.setStyleSheet("""
            QFrame {
                background: #FFEBEE;
                border: 1px solid #EF9A9A;
                border-radius: 4px;
            }
        """)
        warn_layout = QVBoxLayout(warn_frame)
        warn_layout.setContentsMargins(10, 6, 10, 6)
        warn_label = QLabel("This is where your organized files will be created!")
        warn_label.setAlignment(Qt.AlignCenter)
        warn_label.setStyleSheet("""
            QLabel {
                color: #C62828;
                font-weight: bold;
                font-size: 10pt;
                background: #FFEBEE;
            }
        """)
        warn_layout.addWidget(warn_label)
        section.addWidget(warn_frame)

        row = QHBoxLayout()
        self.target_folder = QLineEdit()
        self.target_folder.setPlaceholderText("Select target folder for organized files...")
        self.target_folder.setStyleSheet("""
            QLineEdit {
                background: white;
                color: #323130;
                border: 2px solid #D83B01;
                padding: 8px;
            }
        """)

        self.btn_browse_target = QPushButton("BROWSE")
        self.btn_browse_target.setMinimumWidth(100)
        self.btn_browse_target.setCursor(Qt.PointingHandCursor)
        self.btn_browse_target.setStyleSheet("""
            QPushButton {
                background: #D83B01;
                color: white;
                font-weight: bold;
                font-size: 10pt;
                padding: 8px 20px;
                border: none;
                border-radius: 3px;
            }
            QPushButton:hover { background: #c43400; }
            QPushButton:pressed { background: #b02e00; }
        """)
        self.btn_browse_target.clicked.connect(self._browse_target)

        row.addWidget(self.target_folder)
        row.addWidget(self.btn_browse_target)
        section.addLayout(row)

        return section

    # ── Processing Mode Section ────────────────────────────────

    def _build_processing_mode_section(self):
        section = ColoredSection(
            "PROCESSING MODE",
            "Choose how to organize files",
            "#0099BC"
        )

        self.mode_group = QButtonGroup(self)

        modes = [
            ("Fresh - Create new folder structure (recommended for first time)", "fresh"),
            ("Re-Run - Add new version folders to existing structure", "rerun"),
            ("Resume - Continue previously interrupted processing", "resume"),
        ]

        for i, (label, key) in enumerate(modes):
            radio = QRadioButton(label)
            radio.setStyleSheet("QRadioButton { background: white; color: #323130; spacing: 8px; }")
            if i == 0:
                radio.setChecked(True)
            self.mode_group.addButton(radio, i)
            section.addWidget(radio)

        self.radio_fresh = self.mode_group.button(0)
        self.radio_rerun = self.mode_group.button(1)
        self.radio_resume = self.mode_group.button(2)

        # Separator
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("QFrame { background: #cccccc; max-height: 1px; }")
        section.addWidget(line)

        # Client name option
        row = QHBoxLayout()
        self.chk_client_name = QCheckBox("Include client name in brackets for Level 4 folders")
        self.chk_client_name.setStyleSheet("QCheckBox { background: white; color: #323130; }")
        row.addWidget(self.chk_client_name)

        row.addWidget(self._styled_label("Max length:"))
        self.spin_max_length = QSpinBox()
        self.spin_max_length.setRange(15, 100)
        self.spin_max_length.setValue(35)
        self.spin_max_length.setEnabled(False)
        self.spin_max_length.setStyleSheet("QSpinBox { background: white; color: #323130; }")
        row.addWidget(self.spin_max_length)
        row.addStretch()

        self.chk_client_name.toggled.connect(self.spin_max_length.setEnabled)
        section.addLayout(row)

        note = QLabel("Names >10 chars will show warning. This overrides individual client settings if checked.")
        note.setStyleSheet("QLabel { color: #FF8C00; font-size: 9pt; font-style: italic; background: white; }")
        note.setWordWrap(True)
        section.addWidget(note)

        return section

    # ── Action Buttons Section ─────────────────────────────────

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

        ready_label = QLabel("READY TO SCAN?")
        ready_label.setAlignment(Qt.AlignCenter)
        ready_label.setStyleSheet("""
            QLabel {
                color: #107C10;
                font-size: 14pt;
                font-weight: bold;
                background: white;
            }
        """)
        layout.addWidget(ready_label)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(15)

        self.btn_scan = QPushButton("SCAN FILES")
        self.btn_scan.setCursor(Qt.PointingHandCursor)
        self.btn_scan.setStyleSheet("""
            QPushButton {
                background: #107C10;
                color: white;
                font-size: 12pt;
                font-weight: bold;
                padding: 10px 30px;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover { background: #0e6b0e; }
            QPushButton:pressed { background: #0c5c0c; }
            QPushButton:disabled { background: #a0a0a0; }
        """)
        self.btn_scan.clicked.connect(self.scan_requested.emit)

        self.btn_rescan = QPushButton("RE-SCAN")
        self.btn_rescan.setCursor(Qt.PointingHandCursor)
        self.btn_rescan.setStyleSheet("""
            QPushButton {
                background: #FF8C00;
                color: white;
                font-size: 12pt;
                font-weight: bold;
                padding: 10px 30px;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover { background: #e07d00; }
            QPushButton:pressed { background: #c46e00; }
            QPushButton:disabled { background: #a0a0a0; }
        """)
        self.btn_rescan.clicked.connect(self.rescan_requested.emit)
        self.btn_rescan.setEnabled(False)

        btn_row.addStretch()
        btn_row.addWidget(self.btn_scan)
        btn_row.addWidget(self.btn_rescan)
        btn_row.addStretch()

        layout.addLayout(btn_row)

        return frame

    # ── Right Panel: Collapsible Cards + Scan Log ──────────────

    def _build_instructions_card(self):
        card = CollapsibleCard(
            "INSTRUCTIONS", "#E3F2FD", "#90CAF9", "#1565C0"
        )
        steps = [
            "1. Select folder with GST files",
            "2. Choose Excel templates",
            "3. Select TARGET FOLDER",
            "4. Choose processing mode",
            "5. Click 'Scan Files'",
            "6. Review clients in Validation tab",
            "7. Start processing in Processing tab",
        ]
        for step in steps:
            lbl = QLabel(step)
            lbl.setStyleSheet("QLabel { color: #1565C0; font-size: 10pt; background: #E3F2FD; padding: 1px 0; }")
            card.addWidget(lbl)
        return card

    def _build_patterns_card(self):
        card = CollapsibleCard(
            "EXPECTED FILE NAMES / MODES", "#F3E5F5", "#CE93D8", "#7B1FA2"
        )
        patterns = [
            "GSTR-2B-Reco-Client-State-Period",
            "ImsReco-Client-State-DDMMYYYY",
            "GSTR3B-Client-State-Month",
            "Sales-Client-State-Start-End",
            "SalesReco-Client-State-Period",
            "AnnualReport-Client-State-Year",
        ]
        for p in patterns:
            lbl = QLabel(f"  {p}")
            lbl.setStyleSheet("QLabel { color: #4A148C; font-size: 9pt; background: #F3E5F5; padding: 1px 0; }")
            card.addWidget(lbl)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("QFrame { background: #CE93D8; max-height: 1px; }")
        card.addWidget(sep)

        for name, desc in [("Fresh:", "New timestamped structure"),
                           ("Re-Run:", "Overwrite existing"),
                           ("Resume:", "Skip already processed")]:
            lbl = QLabel(f"  {name} {desc}")
            lbl.setStyleSheet("QLabel { color: #4A148C; font-size: 9pt; background: #F3E5F5; padding: 1px 0; }")
            card.addWidget(lbl)

        return card

    def _build_scan_log(self):
        """Scan activity log - fills remaining space in right panel."""
        frame = QFrame()
        frame.setStyleSheet("""
            QFrame#scanLogFrame {
                background: #1e1e1e;
                border: 2px solid #3c3c3c;
                border-radius: 6px;
            }
        """)
        frame.setObjectName("scanLogFrame")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        header = QLabel("SCAN LOG")
        header.setAlignment(Qt.AlignCenter)
        header.setStyleSheet("""
            QLabel {
                background: #366092;
                color: white;
                font-size: 11pt;
                font-weight: bold;
                padding: 6px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
        """)
        layout.addWidget(header)

        self.scan_log = QTextEdit()
        self.scan_log.setReadOnly(True)
        self.scan_log.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #d4d4d4;
                font-family: "Consolas", "Courier New", monospace;
                font-size: 9pt;
                border: none;
                padding: 6px;
            }
        """)
        self.scan_log.setPlaceholderText("Scan output will appear here...")
        layout.addWidget(self.scan_log)

        # Make scan log expand to fill space
        frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        return frame

    # ── Public: Scan Log Methods ───────────────────────────────

    def log_scan_message(self, message: str, level: str = 'normal'):
        """Append a message to the scan log."""
        color_map = {
            'normal': '#d4d4d4',
            'info': '#569cd6',
            'success': '#4ec9b0',
            'warning': '#dcdcaa',
            'error': '#f44747',
        }
        color = color_map.get(level, '#d4d4d4')
        cursor = self.scan_log.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertHtml(f'<span style="color:{color};">{message}</span><br>')
        self.scan_log.setTextCursor(cursor)
        self.scan_log.ensureCursorVisible()

    def clear_scan_log(self):
        self.scan_log.clear()

    # ── Helpers ─────────────────────────────────────────────────

    def _styled_label(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet("QLabel { background: white; color: #323130; }")
        return lbl

    # ── Browse handlers ────────────────────────────────────────

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

    # ── Public API ─────────────────────────────────────────────

    def get_processing_mode(self) -> str:
        checked = self.mode_group.checkedId()
        return {0: 'fresh', 1: 'rerun', 2: 'resume'}.get(checked, 'fresh')

    def set_values(self, source='', itc='', sales='', target='',
                   mode='fresh', include_name=False, max_length=35):
        self.source_folder.setText(source)
        self.itc_template.setText(itc)
        self.sales_template.setText(sales)
        self.target_folder.setText(target)

        mode_map = {'fresh': self.radio_fresh, 'rerun': self.radio_rerun, 'resume': self.radio_resume}
        btn = mode_map.get(mode, self.radio_fresh)
        btn.setChecked(True)

        self.chk_client_name.setChecked(include_name)
        self.spin_max_length.setValue(max_length)
