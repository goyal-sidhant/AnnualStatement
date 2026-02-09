"""
Scan Controller - handles file scanning in a worker thread.
"""

import logging
from PyQt5.QtCore import QObject, QThread, pyqtSignal

from core.file_parser import FileParser
from models.app_state import ScanResult

logger = logging.getLogger(__name__)


class ScanWorker(QObject):
    """Worker that runs file scanning in a background thread"""
    finished = pyqtSignal(object)  # ScanResult
    progress = pyqtSignal(int, int, str)
    error = pyqtSignal(str)

    def __init__(self, folder_path: str):
        super().__init__()
        self.folder_path = folder_path

    def run(self):
        try:
            parser = FileParser()
            scanned_files, client_data, variations = parser.scan_folder(
                self.folder_path,
                progress_callback=lambda cur, tot, msg: self.progress.emit(cur, tot, msg)
            )
            statistics = parser.get_statistics()
            result = ScanResult(
                scanned_files=scanned_files,
                client_data=client_data,
                variations=variations,
                statistics=statistics,
            )
            self.finished.emit(result)
        except Exception as e:
            logger.error(f"Scan error: {e}", exc_info=True)
            self.error.emit(str(e))


class ScanController(QObject):
    """Controls file scanning operations"""

    scan_started = pyqtSignal()
    scan_progress = pyqtSignal(int, int, str)
    scan_completed = pyqtSignal(object)  # ScanResult
    scan_error = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._thread = None
        self._worker = None

    def scan_files(self, folder_path: str):
        """Start scanning files in the given folder"""
        if self._thread and self._thread.isRunning():
            logger.warning("Scan already in progress")
            return

        self.scan_started.emit()

        self._thread = QThread()
        self._worker = ScanWorker(folder_path)
        self._worker.moveToThread(self._thread)

        # Wire signals
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_finished)
        self._worker.progress.connect(self.scan_progress.emit)
        self._worker.error.connect(self._on_error)

        # Cleanup
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.finished.connect(self._cleanup)

        self._thread.start()

    def _on_finished(self, result):
        self.scan_completed.emit(result)

    def _on_error(self, error_msg):
        self.scan_error.emit(error_msg)

    def _cleanup(self):
        if self._worker:
            self._worker.deleteLater()
            self._worker = None
        if self._thread:
            self._thread.deleteLater()
            self._thread = None
