"""
Extraction Controller - handles PQ refresh and data extraction in a worker thread.
"""

import logging
from datetime import datetime
from pathlib import Path
from PyQt5.QtCore import QObject, QThread, pyqtSignal

from power_query_extractor.core.report_processor import ReportProcessor
from power_query_extractor.core.data_consolidator import DataConsolidator

logger = logging.getLogger(__name__)


class ExtractionWorker(QObject):
    """Worker that runs PQ refresh + extraction in a background thread"""
    progress = pyqtSignal(float, str)
    log = pyqtSignal(str, str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, clients, output_folder, wait_time=10,
                 suffix_pattern='_Refreshed_{timestamp}', skip_refresh=False,
                 level1_folder=None):
        super().__init__()
        self.clients = clients
        self.output_folder = Path(output_folder)
        self.level1_folder = Path(level1_folder) if level1_folder else self.output_folder
        self.wait_time = wait_time
        self.suffix_pattern = suffix_pattern
        self.skip_refresh = skip_refresh
        self.stop_requested = False
        self._log_lines = []

    def run(self):
        try:
            self._process()
        except Exception as e:
            logger.error(f"Extraction error: {e}", exc_info=True)
            self.error.emit(str(e))

    def _log(self, msg, level='info'):
        """Emit log message and capture for log file."""
        self._log_lines.append(msg)
        self.log.emit(msg, level)

    def _process(self):
        processor = ReportProcessor()
        processor.wait_time = self.wait_time
        processor.suffix_pattern = self.suffix_pattern

        # Configure skip refresh
        from power_query_extractor.config.cell_mappings import EXTRACTION_OPTIONS
        EXTRACTION_OPTIONS['skip_refresh'] = self.skip_refresh

        total = len(self.clients)
        results = []

        for idx, client in enumerate(self.clients):
            if self.stop_requested:
                self._log("Extraction stopped by user", 'warning')
                break

            pct = (idx / total) * 100
            name = client.get('name', 'Unknown')
            self.progress.emit(pct, f"Processing {idx+1}/{total}: {name}")
            self._log(f"\n Processing {name}...", 'info')

            def log_callback(msg, level='info'):
                self._log(msg, level)

            result = processor.process_client(client, log_callback=log_callback)
            results.append(result)

            # Log result summary
            itc_ok = result.get('itc', {}).get('status', {}).get('success', False)
            sales_ok = result.get('sales', {}).get('status', {}).get('success', False)
            status = []
            if client.get('process_itc', True):
                status.append(f"ITC: {'OK' if itc_ok else 'FAIL'}")
            if client.get('process_sales', True):
                status.append(f"Sales: {'OK' if sales_ok else 'FAIL'}")
            self._log(f"  Result: {', '.join(status)}", 'success' if (itc_ok or sales_ok) else 'error')

        # Create consolidated report in Level 1 folder
        if results:
            self._log("\n Creating consolidated report...", 'info')
            self.progress.emit(95, "Creating report...")
            try:
                consolidator = DataConsolidator()
                report_path = consolidator.create_report(results, self.level1_folder)
                self._log(f"  Report saved: {report_path.name}", 'success')
            except Exception as e:
                self._log(f"  Report error: {e}", 'error')

        self.progress.emit(100, "Extraction complete")
        self._log(f"\n Extraction complete: {len(results)}/{total} clients processed", 'success')

        # Save extraction log to Level 1 folder
        self._save_log_file()

        self.finished.emit({
            'results': results,
            'total': total,
            'processed': len(results),
        })

    def _save_log_file(self):
        """Save extraction log as text file in Level 1 folder."""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M')
            log_path = self.level1_folder / f"Extraction_Log_{timestamp}.txt"
            with open(log_path, 'w', encoding='utf-8') as f:
                f.write(f"PQ Extraction Log - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 60 + "\n\n")
                for line in self._log_lines:
                    f.write(line + "\n")
            logger.info(f"Extraction log saved: {log_path}")
        except OSError as e:
            logger.warning(f"Could not save extraction log: {e}")


class ExtractionController(QObject):
    """Controls PQ extraction operations"""

    extraction_started = pyqtSignal()
    extraction_progress = pyqtSignal(float, str)
    extraction_log = pyqtSignal(str, str)
    extraction_completed = pyqtSignal(dict)
    extraction_error = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._thread = None
        self._worker = None

    def start_extraction(self, clients, output_folder, wait_time=10,
                         suffix_pattern='_Refreshed_{timestamp}', skip_refresh=False,
                         level1_folder=None):
        """Start extraction in a worker thread"""
        if self._thread and self._thread.isRunning():
            logger.warning("Extraction already in progress")
            return

        self.extraction_started.emit()

        self._thread = QThread()
        self._worker = ExtractionWorker(
            clients, output_folder, wait_time, suffix_pattern, skip_refresh,
            level1_folder=level1_folder
        )
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self.extraction_progress.emit)
        self._worker.log.connect(self.extraction_log.emit)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)

        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.finished.connect(self._cleanup)

        self._thread.start()

    def stop_extraction(self):
        """Request stop"""
        if self._worker:
            self._worker.stop_requested = True

    def _on_finished(self, results):
        self.extraction_completed.emit(results)

    def _on_error(self, error_msg):
        self.extraction_error.emit(error_msg)

    def _cleanup(self):
        if self._worker:
            self._worker.deleteLater()
            self._worker = None
        if self._thread:
            self._thread.deleteLater()
            self._thread = None
