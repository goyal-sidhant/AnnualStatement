"""
Main Window - Thin QMainWindow that wires controllers to views.
Replaces the old Tkinter GSTOrganizerApp god object.
"""

import os
import platform
import logging
from pathlib import Path
from PyQt5.QtWidgets import (
    QMainWindow, QTabWidget, QStatusBar, QMessageBox, QAction, QMenuBar
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon

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

        # Populate extraction tab from all previously processed clients
        self._populate_extraction_from_history()

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

            # Log to setup view
            total = result.total_files
            clients = result.total_clients
            complete = result.complete_clients
            variations = len(result.variations)

            self.setup_view.log_scan_message("", 'normal')
            self.setup_view.log_scan_message(f"Scan complete!", 'success')
            self.setup_view.log_scan_message(f"  Files found: {total}", 'success')
            self.setup_view.log_scan_message(f"  Clients: {clients} ({complete} complete)", 'success')
            if variations:
                self.setup_view.log_scan_message(f"  Variations/issues: {variations}", 'warning')
            self.setup_view.btn_scan.setEnabled(True)

            # Populate validation view
            self.validation_view.populate_clients(result.client_data, self.run_history)

            summary_html = f"""
            <h3>Scan Summary</h3>
            <p><b>Total files:</b> {total}</p>
            <p><b>Clients found:</b> {clients} ({complete} complete)</p>
            <p><b>Variations/issues:</b> {variations}</p>
            """
            self.validation_view.set_summary(summary_html)

            # Enable rescan
            self.setup_view.btn_rescan.setEnabled(True)

            # Switch to validation tab
            self.tabs.setCurrentIndex(1)
            self.status_bar.showMessage(f"Scan complete: {clients} clients, {total} files")
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

        # Refresh extraction tab with ALL processed clients (including this run)
        self._populate_extraction_from_history()

        # Show completion
        successful = summary_data.get('successful_clients', 0)
        total = summary_data.get('total_clients', 0)

        QMessageBox.information(
            self, "Processing Complete",
            f"Successfully processed {successful}/{total} clients\n\n"
            f"Files organized in:\n{self.settings.target_folder}"
        )

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

    def _populate_extraction_from_history(self):
        """Populate extraction tab from ALL processed clients in run history"""
        clients = []
        for client_key, records in self.run_history.client_history.items():
            if not records:
                continue
            latest = records[-1]

            name = latest.get('client_name', '')
            state = latest.get('state', '')
            display_name = f"{name} - {state}" if name and state else client_key

            # Find latest version folder from the output_folder (level2)
            output_folder = latest.get('output_folder', '')
            version_folder = ''
            if output_folder and Path(output_folder).is_dir():
                try:
                    for sub in sorted(Path(output_folder).iterdir(), reverse=True):
                        if sub.is_dir():
                            version_folder = str(sub)
                            break
                except OSError:
                    pass

            clients.append({
                'name': display_name,
                'client_key': client_key,
                'version_folder': version_folder,
                'has_itc': latest.get('itc_report_created', False),
                'has_sales': latest.get('sales_report_created', False),
            })

        # Sort by client name for consistent display
        clients.sort(key=lambda c: c['name'])
        self.extraction_view.populate_clients(clients, self.run_history)

    def _on_start_extraction(self):
        selected = self.extraction_view.get_selected_clients()
        if not selected:
            QMessageBox.warning(self, "Warning", "No clients selected for extraction.")
            return

        if not self.settings.target_folder:
            QMessageBox.warning(self, "Warning", "Please set a target folder in Setup tab.")
            return

        # Build client data for extraction (needs file paths)
        extraction_clients = self._build_extraction_clients(selected)

        if not extraction_clients:
            QMessageBox.warning(
                self, "Warning",
                "Could not find output folders for the selected clients.\n\n"
                "Make sure you have processed these clients first (Step 3)."
            )
            return

        self.extraction_controller.start_extraction(
            clients=extraction_clients,
            output_folder=self.settings.target_folder,
            wait_time=self.extraction_view.spin_wait_time.value(),
            suffix_pattern=self.extraction_view.suffix_input.text(),
            skip_refresh=self.extraction_view.chk_skip_refresh.isChecked(),
        )

    def _build_extraction_clients(self, selected):
        """Build client data dicts for the extraction worker"""
        clients = []
        target = Path(self.settings.target_folder)

        for sel in selected:
            key = sel['client_key']
            name = sel['name']
            version_folder = sel.get('version_folder', '')

            # Use stored version folder path if available
            if version_folder and Path(version_folder).is_dir():
                clients.append({
                    'name': name,
                    'client_key': key,
                    'latest_version': Path(version_folder),
                    'process_itc': sel.get('process_itc', True),
                    'process_sales': sel.get('process_sales', True),
                })
                continue

            # Fallback: search for the client's folder in the target directory
            latest_version = self._find_latest_version_folder(target, key)
            if latest_version:
                clients.append({
                    'name': name,
                    'client_key': key,
                    'latest_version': latest_version,
                    'process_itc': sel.get('process_itc', True),
                    'process_sales': sel.get('process_sales', True),
                })
            else:
                logger.warning(f"Could not find version folder for {key}")

        return clients

    def _find_latest_version_folder(self, target: Path, client_key: str):
        """Search for the latest version folder for a client key"""
        try:
            # Search only top-level subfolders (Level 1 date folders)
            for level1 in sorted(target.iterdir(), reverse=True):
                if not level1.is_dir():
                    continue
                # Look for client folder inside level1
                for level2 in sorted(level1.iterdir(), reverse=True):
                    if not level2.is_dir():
                        continue
                    if client_key in level2.name:
                        # Find latest version subfolder
                        for version in sorted(level2.iterdir(), reverse=True):
                            if version.is_dir():
                                return version
        except OSError as e:
            logger.warning(f"Error searching for {client_key}: {e}")
        return None

    def _on_extraction_completed(self, results):
        self.extraction_view.set_extraction_state(False)
        self.status_bar.showMessage("Extraction complete")

        # Save extraction results to run history
        for result in results.get('results', []):
            client_key = result.get('client', '')
            self.run_history.add_extraction_result(client_key, result)
        self.cache_manager.save_run_history(self.run_history)

    def _on_extraction_error(self, error_msg):
        self.extraction_view.set_extraction_state(False)
        QMessageBox.critical(self, "Extraction Error", error_msg)

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
