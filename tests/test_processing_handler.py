"""Tests for ProcessingHandler guard logic (no GUI/threads exercised).

Uses a minimal fake app so we can assert the re-entrancy guard in
start_processing() without spinning up Tkinter or a worker thread.
"""
from gui.handlers.processing_handler import ProcessingHandler


class FakeFileHandler:
    def __init__(self):
        self.validate_called = False

    def validate_processing_inputs(self):
        self.validate_called = True
        return False   # stop start_processing right after the guard/validation


class FakeApp:
    def __init__(self, is_processing):
        self.is_processing = is_processing
        self.file_handler = FakeFileHandler()


def test_start_processing_blocked_while_already_running():
    """A click while a run is in progress must not proceed past the guard."""
    app = FakeApp(is_processing=True)
    ProcessingHandler(app).start_processing()
    # Guard returned before even validating inputs -> no second run started.
    assert app.file_handler.validate_called is False


def test_start_processing_proceeds_when_idle():
    """When idle, the guard is transparent and normal flow continues."""
    app = FakeApp(is_processing=False)
    ProcessingHandler(app).start_processing()
    assert app.file_handler.validate_called is True
