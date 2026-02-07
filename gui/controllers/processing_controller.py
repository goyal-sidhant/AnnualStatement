"""
Processing Controller - handles file organization and report creation in a worker thread.
"""

import logging
from datetime import datetime
from pathlib import Path
from PyQt5.QtCore import QObject, QThread, pyqtSignal

from core.file_organizer import FileOrganizer
from core.excel_handler import ExcelHandler
from models.app_state import AppSettings
from models.run_history import ClientRunRecord
from utils.helpers import get_timestamp, sanitize_filename, get_state_code, ProgressTracker

logger = logging.getLogger(__name__)


class ProcessingWorker(QObject):
    """Worker that runs file processing in a background thread"""
    progress = pyqtSignal(float, str)
    log = pyqtSignal(str, str)
    finished = pyqtSignal(dict)  # summary_data
    error = pyqtSignal(str)

    def __init__(self, settings: AppSettings, client_data: dict,
                 selected_keys: list, client_folder_settings: dict):
        super().__init__()
        self.settings = settings
        self.client_data = client_data
        self.selected_keys = selected_keys
        self.client_folder_settings = client_folder_settings
        self.stop_requested = False

    def run(self):
        try:
            self._process(self.selected_keys)
        except Exception as e:
            logger.error(f"Processing error: {e}", exc_info=True)
            self.error.emit(str(e))

    def _process(self, selected_clients):
        start_time = datetime.now()
        timestamp = get_timestamp()

        organizer = FileOrganizer(
            self.settings.target_folder,
            self.settings.processing_mode,
            include_client_name=self.settings.include_client_name_in_folders,
            client_folder_settings=self.client_folder_settings,
            client_name_max_length=self.settings.client_name_max_length
        )

        excel_handler = ExcelHandler()
        total_clients = len(selected_clients)
        tracker = ProgressTracker(total_clients)

        summary_data = {
            'timestamp': timestamp,
            'total_clients': total_clients,
            'successful_clients': 0,
            'failed_clients': 0,
            'total_files': 0,
            'reports_generated': 0,
            'itc_reports_created': 0,
            'sales_reports_created': 0,
            'report_errors': 0,
            'processing_mode': self.settings.processing_mode,
            'target_folder': self.settings.target_folder,
            'clients': [],
            'file_mappings': [],
            'errors': [],
            'variations': [],
            'include_client_name': self.settings.include_client_name_in_folders,
            'client_records': [],  # For run history
        }

        level1_folder = None

        for idx, client_key in enumerate(selected_clients):
            if self.stop_requested:
                self.log.emit("Processing stopped by user", 'warning')
                break

            client_info = self.client_data[client_key]
            tracker.update(0 if idx == 0 else 1, f"Processing {client_info['client']}")
            time_remaining = tracker.get_eta() if idx > 0 else "Calculating..."

            pct = (idx / total_clients) * 100
            self.progress.emit(pct, f"Client {idx+1}/{total_clients}: {client_info['client']} | ETA: {time_remaining}")

            try:
                self.log.emit(f"\n🏢 Processing {client_info['client']} - {client_info['state']}", 'info')

                itc_success = False
                sales_success = False

                # Create folders
                folders = organizer.create_client_structure(client_info)
                if level1_folder is None:
                    level1_folder = folders['level1']

                self.log.emit("  ✓ Created folder structure", 'success')

                # Organize files
                file_results = organizer.organize_files(client_info, folders)
                successful_files = sum(1 for r in file_results if r['status'] == 'Success')
                summary_data['total_files'] += successful_files
                self.log.emit(f"  ✓ Organized {successful_files} files", 'success')

                # Prepare safe names
                safe_client = sanitize_filename(client_info['client'])
                safe_state = client_info.get('state_code', get_state_code(client_info['state']))
                safe_ts = timestamp.replace(' ', '_')

                # ITC Report
                itc_success = self._create_report(
                    excel_handler, client_info, folders,
                    'ITC', self.settings.itc_template,
                    f"ITC_Report_{safe_client}_{safe_state}_{safe_ts}",
                    summary_data
                )

                # Sales Report
                sales_success = self._create_report(
                    excel_handler, client_info, folders,
                    'Sales', self.settings.sales_template,
                    f"Sales_Report_{safe_client}_{safe_state}_{safe_ts}",
                    summary_data
                )

                # Organization report
                organizer.create_organization_report(folders, client_info)

                summary_data['successful_clients'] += 1
                summary_data['clients'].append({
                    'client': client_info['client'],
                    'state': client_info['state'],
                    'status': 'Success',
                    'file_count': client_info['file_count'],
                    'missing_files': client_info['missing_files'],
                    'completeness': client_info.get('completeness', 0),
                    'itc_report_created': itc_success,
                    'sales_report_created': sales_success
                })

                # Track for run history
                summary_data['client_records'].append(ClientRunRecord(
                    client_key=client_key,
                    client_name=client_info['client'],
                    state=client_info['state'],
                    run_timestamp=datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    processing_mode=self.settings.processing_mode,
                    files_organized=successful_files,
                    itc_report_created=itc_success,
                    sales_report_created=sales_success,
                    output_folder=str(folders.get('level2', '')),
                ))

                for result in file_results:
                    summary_data['file_mappings'].append({
                        'filename': result['filename'],
                        'client': client_key,
                        'type': result['file_type'],
                        'destination': result.get('destination', ''),
                        'status': result['status']
                    })

                self.log.emit(f"  ✅ Completed {client_info['client']}", 'success')

            except Exception as e:
                summary_data['failed_clients'] += 1
                summary_data['errors'].append({
                    'time': datetime.now().strftime('%H:%M:%S'),
                    'client': client_key,
                    'type': 'Processing Error',
                    'message': str(e)
                })
                self.log.emit(f"  ❌ Failed {client_key}: {str(e)}", 'error')

        # Summary report
        if summary_data['successful_clients'] > 0 and level1_folder:
            self.progress.emit(95, "Creating summary report...")
            self.log.emit("\n📊 Creating summary report...", 'info')
            safe_ts = timestamp.replace(' ', '_')
            summary_path = level1_folder / f"GST_Processing_Summary_{safe_ts}.xlsx"
            try:
                if excel_handler.create_summary_report(summary_path, summary_data):
                    self.log.emit(f"  ✅ Summary report created", 'success')
            except Exception as e:
                self.log.emit(f"  ❌ Summary error: {e}", 'error')

        # Completion
        elapsed = tracker.get_elapsed_time()
        mins, secs = int(elapsed // 60), int(elapsed % 60)
        self.progress.emit(100, f"Complete | {mins}m {secs}s")
        self.log.emit(f"\n🎉 Processing complete: {summary_data['successful_clients']}/{total_clients} clients", 'success')

        summary_data['level1_folder'] = str(level1_folder) if level1_folder else ''
        self.finished.emit(summary_data)

    def _create_report(self, excel_handler, client_info, folders,
                       report_type, template_path, report_name, summary_data):
        """Create a single report (ITC or Sales). Returns True on success."""
        label = "📊" if report_type == 'ITC' else "💰"
        self.log.emit(f"  {label} Creating {report_type} report...", 'info')

        if not template_path:
            self.log.emit(f"    ❌ {report_type} template path is empty!", 'error')
            summary_data['report_errors'] += 1
            return False

        if not Path(template_path).exists():
            self.log.emit(f"    ❌ {report_type} template not found", 'error')
            summary_data['report_errors'] += 1
            return False

        output_path = folders['version'] / f"{report_name}.xlsx"
        self.log.emit(f"    Output: {output_path.name}", 'normal')

        try:
            data = excel_handler.prepare_template_data(client_info, folders, report_type)
            if excel_handler.create_report_from_template(template_path, output_path, data, report_type):
                summary_data['reports_generated'] += 1
                key = f'{report_type.lower()}_reports_created'
                summary_data[key] = summary_data.get(key, 0) + 1
                self.log.emit(f"    ✅ Success: {report_name}.xlsx", 'success')
                return True
            else:
                summary_data['report_errors'] += 1
                self.log.emit(f"    ❌ Failed to create {report_type} report", 'error')
                return False
        except Exception as e:
            summary_data['report_errors'] += 1
            self.log.emit(f"    ❌ {report_type} Report Error: {e}", 'error')
            logger.error(f"{report_type} Report Error: {e}", exc_info=True)
            return False


class ProcessingController(QObject):
    """Controls file processing operations"""

    processing_started = pyqtSignal()
    processing_progress = pyqtSignal(float, str)
    processing_log = pyqtSignal(str, str)
    processing_completed = pyqtSignal(dict)
    processing_error = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._thread = None
        self._worker = None

    def start_processing(self, settings: AppSettings, client_data: dict,
                         selected_keys: list, client_folder_settings: dict):
        """Start processing in a worker thread"""
        if self._thread and self._thread.isRunning():
            logger.warning("Processing already in progress")
            return

        self.processing_started.emit()

        self._thread = QThread()
        self._worker = ProcessingWorker(settings, client_data, selected_keys, client_folder_settings)
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self.processing_progress.emit)
        self._worker.log.connect(self.processing_log.emit)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)

        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.finished.connect(self._cleanup)

        self._thread.start()

    def stop_processing(self):
        """Request stop"""
        if self._worker:
            self._worker.stop_requested = True

    def _on_finished(self, summary_data):
        self.processing_completed.emit(summary_data)

    def _on_error(self, error_msg):
        self.processing_error.emit(error_msg)

    def _cleanup(self):
        if self._worker:
            self._worker.deleteLater()
            self._worker = None
        if self._thread:
            self._thread.deleteLater()
            self._thread = None
