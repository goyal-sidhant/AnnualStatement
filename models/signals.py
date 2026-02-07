"""
Central signal hub for the PyQt5 application.
All inter-component communication goes through these signals.
"""

from PyQt5.QtCore import QObject, pyqtSignal


class AppSignals(QObject):
    """Central signal hub - replaces direct self.app.xxx references"""

    # ── Settings ──────────────────────────────────────────────
    settings_changed = pyqtSignal()

    # ── Scan ──────────────────────────────────────────────────
    scan_started = pyqtSignal()
    scan_progress = pyqtSignal(int, int, str)       # current, total, message
    scan_completed = pyqtSignal(object)              # ScanResult
    scan_error = pyqtSignal(str)

    # ── Processing ────────────────────────────────────────────
    processing_started = pyqtSignal()
    processing_progress = pyqtSignal(float, str)     # percent, message
    processing_log = pyqtSignal(str, str)            # message, level
    processing_completed = pyqtSignal(dict)          # summary_data
    processing_error = pyqtSignal(str)
    processing_stopped = pyqtSignal()

    # ── PQ Extraction ─────────────────────────────────────────
    extraction_started = pyqtSignal()
    extraction_progress = pyqtSignal(float, str)     # percent, message
    extraction_log = pyqtSignal(str, str)            # message, level
    extraction_completed = pyqtSignal(dict)          # results
    extraction_error = pyqtSignal(str)

    # ── UI ────────────────────────────────────────────────────
    status_message = pyqtSignal(str)
    switch_tab = pyqtSignal(int)                     # tab index
