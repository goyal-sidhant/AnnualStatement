"""Tests for the workflow mode: organise-only vs full pipeline.

The mode decides whether Step 3 hands over to Step 4 automatically. It must
NEVER disable Step 4 - that tab stays usable on its own so an older batch can
be refreshed without re-running everything.
"""
import pytest

from gui.handlers.processing_handler import ProcessingHandler
from utils.constants import WORKFLOW_MODES


class FakeVar:
    def __init__(self, value=''):
        self._value = value

    def get(self):
        return self._value

    def set(self, value):
        self._value = value


class FakePanel:
    def __init__(self):
        self.folder_path = FakeVar('')
        self.scanned = 0

    def scan_folder(self):
        self.scanned += 1


class FakeExtractTab:
    def __init__(self):
        self.panel = FakePanel()


class FakeNotebook:
    def __init__(self):
        self.selected = None

    def select(self, index):
        self.selected = index


class FakeRoot:
    """after() runs immediately so the hand-off is observable in the test."""
    def after(self, _delay, callback):
        callback()


class FakeApp:
    def __init__(self, mode='organize', with_extract_tab=True):
        self.root = FakeRoot()
        self.notebook = FakeNotebook()
        self.workflow_mode = FakeVar(mode)
        self.target_folder = FakeVar(r'C:\out')
        self.extract_tab = FakeExtractTab() if with_extract_tab else None
        self.logs = []

    def log_message(self, message, level='normal'):
        self.logs.append((message, level))


def texts(app):
    return " ".join(m for m, _ in app.logs)


class TestWorkflowModes:
    def test_two_modes_are_defined(self):
        assert set(WORKFLOW_MODES) == {'organize', 'full'}
        for info in WORKFLOW_MODES.values():
            assert info['name'] and info['description']


class TestHandoffToStep4:
    def test_full_pipeline_opens_and_scans_step_4(self):
        app = FakeApp(mode='full')
        ProcessingHandler(app)._maybe_continue_to_extraction()

        assert app.notebook.selected == 3, "should switch to Step 4"
        assert app.extract_tab.panel.scanned == 1, "should scan so the list is ready"
        assert app.extract_tab.panel.folder_path.get() == r'C:\out'

    def test_full_pipeline_does_not_start_the_refresh(self):
        """Refreshing drives Excel for minutes per report, so it stays an
        explicit press - the hand-off only prepares the tab."""
        app = FakeApp(mode='full')
        ProcessingHandler(app)._maybe_continue_to_extraction()
        assert 'press Process' in texts(app)

    def test_organize_only_does_not_switch_tabs(self):
        app = FakeApp(mode='organize')
        ProcessingHandler(app)._maybe_continue_to_extraction()

        assert app.notebook.selected is None
        assert app.extract_tab.panel.scanned == 0

    def test_organize_only_still_points_at_step_4(self):
        """Step 4 remains available - the user is told, not blocked."""
        app = FakeApp(mode='organize')
        ProcessingHandler(app)._maybe_continue_to_extraction()
        assert 'Step 4' in texts(app)

    def test_missing_extract_tab_is_survivable(self):
        """The standalone/older layout has no Step 4; this must not raise."""
        app = FakeApp(mode='full', with_extract_tab=False)
        ProcessingHandler(app)._maybe_continue_to_extraction()   # must not raise
        assert app.notebook.selected is None

    def test_a_broken_panel_does_not_kill_processing(self):
        """A failure handing over must not surface as a processing failure."""
        app = FakeApp(mode='full')

        def boom():
            raise RuntimeError("scan exploded")
        app.extract_tab.panel.scan_folder = boom

        ProcessingHandler(app)._maybe_continue_to_extraction()   # must not raise
        assert 'Could not open Step 4' in texts(app)
