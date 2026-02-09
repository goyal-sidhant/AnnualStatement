"""
Main Window - Thin QMainWindow that wires controllers to views.
Replaces the old Tkinter GSTOrganizerApp god object.
"""

import os
import platform
import logging
from datetime import datetime
from pathlib import Path
from PyQt5.QtWidgets import (
    QMainWindow, QTabWidget, QStatusBar, QMessageBox, QAction, QMenuBar
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QColor

from models.app_state import AppSettings, ScanResult, ProcessingState
from models.run_history import RunHistory
from models.cache_manager import CacheManager
from gui.theme.theme_manager import ThemeManager
from gui.views.setup_view import SetupView
from gui.views.validation_view import ValidationView
from gui.views.processing_view import ProcessingView
from gui.views.extraction_view import ExtractionView
from gui.controllers.scan_controller import ScanController
from gui.controllers.processing_controller import ProcessingController
from gui.controllers.extraction_controller import ExtractionController

logger = logging.getLogger(__name__)


class GSTOrganizerWindow(QMainWindow):
    """Main application window with 4 tabs"""

    def __init__(self):
        super().__init__()

        # State
        self.settings = AppSettings()
        self.scan_result = ScanResult()
        self.processing_state = ProcessingState()
        self.cache_manager = CacheManager()
        self.run_history = self.cache_manager.load_run_history()
        self.theme_manager = ThemeManager()

        # Migrate old cache if needed
        self.cache_manager.migrate_from_old_cache()

        # Load saved settings
        saved = self.cache_manager.load_settings()
        self.settings = saved

        # Controllers
        self.scan_controller = ScanController(self)
        self.processing_controller = ProcessingController(self)
        self.extraction_controller = ExtractionController(self)

        # Build UI
        self._setup_window()
        self._create_views()
        self._create_menu()
        self._wire_signals()
        self._restore_settings()

        # Apply theme
        self.theme_manager.apply_theme(self.settings.dark_mode)

        logger.info("GST File Organizer v4.0 initialized")

    def _setup_window(self):
        self.setWindowTitle("GST File Organizer v4.0")
        self.setMinimumSize(1200, 800)
        self.resize(1400, 900)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

    def _create_views(self):
        """Create the 4-tab layout"""
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.setup_view = SetupView()
        self.validation_view = ValidationView()
        self.processing_view = ProcessingView()
        self.extraction_view = ExtractionView()

        self.tabs.addTab(self.setup_view, "Step 1: Setup")
        self.tabs.addTab(self.validation_view, "Step 2: Validation")
        self.tabs.addTab(self.processing_view, "Step 3: Processing")
        self.tabs.addTab(self.extraction_view, "Step 4: Extract && Refresh")

    def _create_menu(self):
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("&File")

        save_action = QAction("&Save Settings", self)
        save_action.triggered.connect(self._save_settings)
        file_menu.addAction(save_action)

        file_menu.addSeparator()

        exit_action = QAction("E&xit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # View menu
        view_menu = menubar.addMenu("&View")

        self.dark_mode_action = QAction("&Dark Mode", self)
        self.dark_mode_action.setCheckable(True)
        self.dark_mode_action.setChecked(self.settings.dark_mode)
        self.dark_mode_action.triggered.connect(self._toggle_dark_mode)
        view_menu.addAction(self.dark_mode_action)

    def _wire_signals(self):
        """Connect all signals between controllers and views"""

        # Setup view -> scan
        self.setup_view.scan_requested.connect(self._on_scan_requested)
        self.setup_view.rescan_requested.connect(self._on_scan_requested)

        # Scan controller -> setup view log + validation view
        self.scan_controller.scan_started.connect(self._on_scan_started)
        self.scan_controller.scan_progress.connect(self._on_scan_progress)
        self.scan_controller.scan_completed.connect(self._on_scan_completed)
        self.scan_controller.scan_error.connect(self._on_scan_error)

        # Validation view -> processing
        self.validation_view.start_processing_requested.connect(self._on_start_processing)
        self.validation_view.dry_run_requested.connect(self._on_dry_run)

        # Processing controller -> processing view
        self.processing_controller.processing_started.connect(
            lambda: self.processing_view.set_processing_state(True))
        self.processing_controller.processing_progress.connect(
            self.processing_view.update_progress)
        self.processing_controller.processing_log.connect(
            self.processing_view.log_message)
        self.processing_controller.processing_completed.connect(
            self._on_processing_completed)
        self.processing_controller.processing_error.connect(
            self._on_processing_error)

        # Processing view buttons
        self.processing_view.start_requested.connect(self._on_start_processing)
        self.processing_view.stop_requested.connect(
            self.processing_controller.stop_processing)

        # Extraction controller -> extraction view
        self.extraction_controller.extraction_started.connect(
            lambda: self.extraction_view.set_extraction_state(True))
        self.extraction_controller.extraction_progress.connect(
            self.extraction_view.update_progress)
        self.extraction_controller.extraction_log.connect(
            self.extraction_view.log_message)
        self.extraction_controller.extraction_completed.connect(
            self._on_extraction_completed)
        self.extraction_controller.extraction_error.connect(
            self._on_extraction_error)

        # Extraction view buttons
        self.extraction_view.start_extraction_requested.connect(
            self._on_start_extraction)
        self.extraction_view.stop_extraction_requested.connect(
            self.extraction_controller.stop_extraction)
        self.extraction_view.refresh_list_requested.connect(
            self._populate_extraction_from_disk)

    def _restore_settings(self):
        """Populate views from saved settings"""
        self.setup_view.set_values(
            source=self.settings.source_folder,
            itc=self.settings.itc_template,
            sales=self.settings.sales_template,
            target=self.settings.target_folder,
            mode=self.settings.processing_mode,
            include_name=self.settings.include_client_name_in_folders,
            max_length=self.settings.client_name_max_length,
        )
        self.extraction_view.spin_wait_time.setValue(self.settings.pq_wait_time)
        self.extraction_view.suffix_input.setText(self.settings.pq_suffix_pattern)
        self.extraction_view.chk_skip_refresh.setChecked(self.settings.pq_skip_refresh)

        # Populate extraction tab from disk (latest Level 1 folder)
        self._populate_extraction_from_disk()

    # ── Settings ──────────────────────────────────────────────

    def _collect_settings(self) -> AppSettings:
        """Collect current settings from all views"""
        return AppSettings(
            source_folder=self.setup_view.source_folder.text(),
            itc_template=self.setup_view.itc_template.text(),
            sales_template=self.setup_view.sales_template.text(),
            target_folder=self.setup_view.target_folder.text(),
            processing_mode=self.setup_view.get_processing_mode(),
            include_client_name_in_folders=self.setup_view.chk_client_name.isChecked(),
            client_name_max_length=self.setup_view.spin_max_length.value(),
            dark_mode=self.settings.dark_mode,
            pq_wait_time=self.extraction_view.spin_wait_time.value(),
            pq_suffix_pattern=self.extraction_view.suffix_input.text(),
            pq_skip_refresh=self.extraction_view.chk_skip_refresh.isChecked(),
        )

    def _save_settings(self):
        """Save current settings to cache"""
        self.settings = self._collect_settings()
        self.cache_manager.save_settings(self.settings)
        self.status_bar.showMessage("Settings saved", 3000)

    # ── Scan ──────────────────────────────────────────────────

    def _on_scan_requested(self):
        folder = self.setup_view.source_folder.text()
        if not folder:
            QMessageBox.warning(self, "Warning", "Please select a source folder first.")
            return
        if not Path(folder).is_dir():
            QMessageBox.warning(self, "Warning", f"Folder not found: {folder}")
            return

        self._save_settings()
        self.setup_view.collapse_info_cards()
        self.setup_view.clear_scan_log()
        self.setup_view.log_scan_message(f"Scanning: {folder}", 'info')
        self.setup_view.btn_scan.setEnabled(False)
        self.scan_controller.scan_files(folder)

    def _on_scan_started(self):
        self.status_bar.showMessage("Scanning files...")
        self.setup_view.log_scan_message("Scan started...", 'info')

    def _on_scan_progress(self, current, total, message):
        self.setup_view.log_scan_message(message, 'normal')
        if total > 0:
            pct = int(current / total * 100)
            self.status_bar.showMessage(f"Scanning... {pct}% ({current}/{total})")

    def _on_scan_completed(self, result: ScanResult):
        try:
            self.scan_result = result
            stats = result.statistics

            # Log detailed stats to setup view scan log
            self.setup_view.log_scan_message("", 'normal')
            self.setup_view.log_scan_message("SCAN RESULTS", 'success')
            self.setup_view.log_scan_message("=" * 40, 'success')
            self.setup_view.log_scan_message(f"  Files found: {stats.get('total_files', 0)}", 'info')
            self.setup_view.log_scan_message(f"  Parsed: {stats.get('parsed_files', 0)} ({stats.get('parsing_rate', 0):.1f}%)", 'info')
            self.setup_view.log_scan_message(f"  Unparsed: {stats.get('unparsed_files', 0)}", 'warning' if stats.get('unparsed_files', 0) else 'info')
            self.setup_view.log_scan_message(f"  Total size: {stats.get('total_size_formatted', '0 B')}", 'info')
            self.setup_view.log_scan_message("", 'normal')
            self.setup_view.log_scan_message(f"  Clients: {stats.get('total_clients', 0)}", 'info')
            self.setup_view.log_scan_message(f"  Complete: {stats.get('complete_clients', 0)} ({stats.get('completion_rate', 0):.1f}%)", 'success')
            self.setup_view.log_scan_message(f"  Incomplete: {stats.get('incomplete_clients', 0)}", 'warning' if stats.get('incomplete_clients', 0) else 'info')
            if stats.get('variations', 0):
                self.setup_view.log_scan_message(f"  Variations: {stats['variations']}", 'warning')
            # File type distribution
            dist = stats.get('file_type_distribution', {})
            if dist:
                self.setup_view.log_scan_message("", 'normal')
                self.setup_view.log_scan_message("FILE TYPES:", 'info')
                for ftype, count in sorted(dist.items()):
                    self.setup_view.log_scan_message(f"  {ftype}: {count}", 'normal')

            self.setup_view.btn_scan.setEnabled(True)

            # Populate validation view
            self.validation_view.populate_clients(result.client_data, self.run_history)

            # Build rich summary HTML for validation tab
            dist_rows = ""
            for ftype, count in sorted(dist.items()):
                dist_rows += f"<tr><td style='padding:1px 8px;'>{ftype}</td><td style='padding:1px 8px;'><b>{count}</b></td></tr>"

            summary_html = f"""
            <h3 style="color:#0078D4; margin-bottom:4px;">Scan Summary</h3>
            <table style="margin-bottom:6px;">
              <tr><td><b>Total files:</b></td><td>{stats.get('total_files', 0)}</td></tr>
              <tr><td><b>Parsed:</b></td><td>{stats.get('parsed_files', 0)} ({stats.get('parsing_rate', 0):.1f}%)</td></tr>
              <tr><td><b>Unparsed:</b></td><td>{stats.get('unparsed_files', 0)}</td></tr>
              <tr><td><b>Total size:</b></td><td>{stats.get('total_size_formatted', '0 B')}</td></tr>
            </table>
            <h4 style="color:#107C10; margin-bottom:4px;">Clients</h4>
            <table style="margin-bottom:6px;">
              <tr><td><b>Total clients:</b></td><td>{stats.get('total_clients', 0)}</td></tr>
              <tr><td><b>Complete:</b></td><td>{stats.get('complete_clients', 0)} ({stats.get('completion_rate', 0):.1f}%)</td></tr>
              <tr><td><b>Incomplete:</b></td><td>{stats.get('incomplete_clients', 0)}</td></tr>
              <tr><td><b>Variations:</b></td><td>{stats.get('variations', 0)}</td></tr>
            </table>
            <h4 style="color:#FF8C00; margin-bottom:4px;">File Types</h4>
            <table>{dist_rows}</table>
            """
            self.validation_view.set_summary(summary_html)

            # Enable rescan
            self.setup_view.btn_rescan.setEnabled(True)

            # Update tab state indicators and switch to validation
            self._update_tab_state(0, 'complete')
            self._update_tab_state(1, 'ready')
            self.tabs.setCurrentIndex(1)
            self.status_bar.showMessage(
                f"Scan complete: {stats.get('total_clients', 0)} clients, "
                f"{stats.get('total_files', 0)} files, "
                f"{stats.get('total_size_formatted', '')}"
            )
        except Exception as e:
            logger.error(f"Error in scan completion handler: {e}", exc_info=True)
            self.setup_view.log_scan_message(f"Error displaying results: {e}", 'error')
            self.setup_view.btn_scan.setEnabled(True)

    def _on_scan_error(self, error_msg):
        self.setup_view.log_scan_message(f"SCAN ERROR: {error_msg}", 'error')
        self.setup_view.btn_scan.setEnabled(True)
        QMessageBox.critical(self, "Scan Error", error_msg)
        self.status_bar.showMessage("Scan failed")

    # ── Processing ────────────────────────────────────────────

    def _on_dry_run(self):
        selected = self.validation_view.get_selected_client_keys()
        if not selected:
            QMessageBox.warning(self, "Warning", "No clients selected.")
            return

        self.tabs.setCurrentIndex(2)
        self.processing_view.log_message("🧪 ════════ DRY RUN PREVIEW ════════", 'info')
        for key in selected:
            info = self.scan_result.client_data.get(key, {})
            self.processing_view.log_message(f"  🏢 {info.get('client', key)} - {info.get('state', '')}", 'info')
            for ft, files in info.get('files', {}).items():
                for f in files:
                    self.processing_view.log_message(f"    📄 Would organize: {f.get('name', '')}", 'normal')
            self.processing_view.log_message("    📊 Would create: ITC Report", 'success')
            self.processing_view.log_message("    💰 Would create: Sales Report", 'success')
        self.processing_view.log_message("🧪 ════════ DRY RUN COMPLETE ════════", 'success')
        self.processing_view.log_message("💡 No files were actually moved or created", 'warning')

    def _on_start_processing(self):
        selected = self.validation_view.get_selected_client_keys()
        if not selected:
            QMessageBox.warning(self, "Warning", "No clients selected.")
            return

        # Validate inputs
        settings = self._collect_settings()
        if not settings.itc_template or not Path(settings.itc_template).exists():
            QMessageBox.warning(self, "Warning", "Please select a valid ITC template.")
            return
        if not settings.sales_template or not Path(settings.sales_template).exists():
            QMessageBox.warning(self, "Warning", "Please select a valid Sales template.")
            return
        if not settings.target_folder:
            QMessageBox.warning(self, "Warning", "Please select a target folder.")
            return

        self.settings = settings
        self.tabs.setCurrentIndex(2)

        self.processing_controller.start_processing(
            settings=self.settings,
            client_data=self.scan_result.client_data,
            selected_keys=selected,
            client_folder_settings=self.processing_state.client_folder_settings,
        )

    def _on_processing_completed(self, summary_data):
        self.processing_view.set_processing_state(False)
        self.status_bar.showMessage("Processing complete")

        # Save run history
        client_records = summary_data.pop('client_records', [])
        self.run_history.add_run(summary_data, client_records)
        self.cache_manager.save_run_history(self.run_history)
        self._save_settings()

        # Refresh extraction tab from disk (includes just-processed clients)
        self._populate_extraction_from_disk()

        # Show completion
        successful = summary_data.get('successful_clients', 0)
        total = summary_data.get('total_clients', 0)

        # Update tab state indicators
        self._update_tab_state(2, 'complete')
        self._update_tab_state(3, 'ready')

        QMessageBox.information(
            self, "Processing Complete",
            f"Successfully processed {successful}/{total} clients\n\n"
            f"Files organized in:\n{self.settings.target_folder}"
        )

        # Navigate to extraction tab
        self.tabs.setCurrentIndex(3)

        # Open output folder
        level1 = summary_data.get('level1_folder', '')
        if level1 and Path(level1).exists():
            try:
                if platform.system() == 'Windows':
                    os.startfile(level1)
            except Exception as e:
                logger.warning(f"Could not open folder: {e}")

    def _on_processing_error(self, error_msg):
        self.processing_view.set_processing_state(False)
        QMessageBox.critical(self, "Processing Error", error_msg)
        self.status_bar.showMessage("Processing failed")

    # ── Extraction ────────────────────────────────────────────

    def _populate_extraction_from_disk(self):
        """Scan target folder on disk and populate extraction tab with tree data."""
        target_str = self.settings.target_folder
        if not target_str:
            self.extraction_view.show_no_data_message("No target folder configured.")
            return
        target = Path(target_str)
        if not target.is_dir():
            self.extraction_view.show_no_data_message("Target folder does not exist.")
            return

        # Find latest Level 1 folder (Annual Statement-*)
        try:
            level1_folders = sorted(
                [f for f in target.iterdir()
                 if f.is_dir() and f.name.startswith('Annual Statement-')],
                key=lambda p: p.name, reverse=True
            )
        except OSError as e:
            logger.warning(f"Error scanning target folder: {e}")
            self.extraction_view.show_no_data_message(f"Error scanning: {e}")
            return

        if not level1_folders:
            self.extraction_view.show_no_data_message(
                "No 'Annual Statement-...' folders found in target."
            )
            return

        latest_l1 = level1_folders[0]

        # Build tree data: scan Level 2 (client) → Version folders
        tree_data = []
        try:
            for level2 in sorted(latest_l1.iterdir()):
                if not level2.is_dir():
                    continue
                versions = []
                for version_dir in sorted(level2.iterdir(), reverse=True):
                    if not version_dir.is_dir():
                        continue
                    # Check for ITC/Sales reports
                    has_itc = any(version_dir.glob("ITC_Report_*.xlsx"))
                    has_sales = any(version_dir.glob("Sales_Report_*.xlsx"))
                    if not has_itc and not has_sales:
                        continue  # Skip version folders with no reports

                    # Check for refreshed files
                    refreshed_files = [
                        f for f in version_dir.glob("*_Refreshed*")
                        if f.suffix.lower() in {'.xlsx', '.xls', '.xlsm'}
                        and not f.name.startswith('~$')
                    ]
                    refresh_status = 'Pending'
                    refresh_timestamp = ''
                    if refreshed_files:
                        newest = max(refreshed_files, key=lambda f: f.stat().st_mtime)
                        refresh_timestamp = datetime.fromtimestamp(
                            newest.stat().st_mtime
                        ).strftime('%d/%m/%y %H:%M')
                        refresh_status = 'Refreshed'

                    versions.append({
                        'name': version_dir.name,
                        'path': version_dir,
                        'has_itc': has_itc,
                        'has_sales': has_sales,
                        'refresh_status': refresh_status,
                        'refresh_timestamp': refresh_timestamp,
                    })

                if versions:
                    tree_data.append({
                        'name': level2.name,
                        'path': level2,
                        'versions': versions,
                    })
        except OSError as e:
            logger.warning(f"Error scanning client folders: {e}")

        if tree_data:
            self.extraction_view.populate_tree(tree_data, latest_l1)
        else:
            self.extraction_view.show_no_data_message(
                f"No processed clients found in {latest_l1.name}"
            )

    def _on_start_extraction(self):
        selected = self.extraction_view.get_selected_clients()
        if not selected:
            QMessageBox.warning(self, "Warning", "No versions selected for extraction.")
            return

        level1_folder = self.extraction_view.get_level1_folder()
        if not level1_folder:
            QMessageBox.warning(self, "Warning", "No Level 1 folder available.")
            return

        # Build client data — version paths come directly from tree items
        extraction_clients = []
        for sel in selected:
            vf = sel.get('version_folder', '')
            if vf and Path(vf).is_dir():
                extraction_clients.append({
                    'name': sel['name'],
                    'client_key': sel['client_key'],
                    'latest_version': Path(vf),
                    'process_itc': sel.get('process_itc', True),
                    'process_sales': sel.get('process_sales', True),
                })
            else:
                logger.warning(f"Version folder not found: {vf}")

        if not extraction_clients:
            QMessageBox.warning(
                self, "Warning",
                "Could not find version folders for the selected items.\n\n"
                "Make sure clients have been processed first (Step 3)."
            )
            return

        self.extraction_controller.start_extraction(
            clients=extraction_clients,
            output_folder=self.settings.target_folder,
            level1_folder=level1_folder,
            wait_time=self.extraction_view.spin_wait_time.value(),
            suffix_pattern=self.extraction_view.suffix_input.text(),
            skip_refresh=self.extraction_view.chk_skip_refresh.isChecked(),
        )

    def _on_extraction_completed(self, results):
        self.extraction_view.set_extraction_state(False)
        self._update_tab_state(3, 'complete')
        self.status_bar.showMessage("Extraction complete")

        # Save extraction results to run history
        for result in results.get('results', []):
            client_key = result.get('client', '')
            self.run_history.add_extraction_result(client_key, result)
        self.cache_manager.save_run_history(self.run_history)

        # Re-scan disk to update refresh status in tree
        self._populate_extraction_from_disk()

    def _on_extraction_error(self, error_msg):
        self.extraction_view.set_extraction_state(False)
        QMessageBox.critical(self, "Extraction Error", error_msg)

    # ── Tab State Indicators ─────────────────────────────────

    def _make_dot_icon(self, color_hex):
        """Create a small colored circle icon for tab state."""
        size = 12
        pixmap = QPixmap(size, size)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QColor(color_hex))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(1, 1, size - 2, size - 2)
        painter.end()
        return QIcon(pixmap)

    def _update_tab_state(self, tab_index, state):
        """Update tab icon: 'pending' = gray, 'ready' = orange, 'complete' = green."""
        colors = {
            'pending': '#a0a0a0',
            'ready': '#FF8C00',
            'complete': '#107C10',
        }
        color = colors.get(state, '#a0a0a0')
        self.tabs.setTabIcon(tab_index, self._make_dot_icon(color))

    # ── Theme ─────────────────────────────────────────────────

    def _toggle_dark_mode(self, checked):
        self.settings.dark_mode = checked
        self.theme_manager.apply_theme(checked)
        self._save_settings()

    # ── Lifecycle ─────────────────────────────────────────────

    def closeEvent(self, event):
        """Save state on close"""
        self._save_settings()
        self.cache_manager.save_session_state({
            'active_tab': self.tabs.currentIndex(),
            'window_size': [self.width(), self.height()],
        })
        event.accept()
