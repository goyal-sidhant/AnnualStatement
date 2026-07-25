"""Smoke tests that build the real SetupTab in a real Tk window.

These catch wiring mistakes that pure-logic tests cannot: missing app
attributes, bad widget options, broken command= references. Skipped
automatically where no display/Tk is available.
"""
import tkinter as tk

import pytest

from gui.tabs.setup_tab import SetupTab
from gui.utils import setup_validation as sv


@pytest.fixture(scope="module")
def root():
    """One Tk root for the whole module.

    Creating and destroying a Tk() per test exhausts Tk resources and makes the
    suite flakily skip, so the root is shared; each test still gets a fresh
    container frame and a fresh FakeApp (with its own variables and traces).
    """
    try:
        r = tk.Tk()
    except tk.TclError:
        pytest.skip("no display available for Tk")
    r.withdraw()
    yield r
    try:
        r.destroy()
    except tk.TclError:
        pass


class FakeApp:
    """Minimal stand-in exposing exactly what SetupTab touches."""

    def __init__(self, root):
        self.root = root
        self.source_folder = tk.StringVar()
        self.itc_template = tk.StringVar()
        self.sales_template = tk.StringVar()
        self.target_folder = tk.StringVar()
        self.processing_mode = tk.StringVar(value='fresh')
        self.include_client_name_in_folders = tk.BooleanVar(value=False)
        self.client_name_max_length = tk.IntVar(value=35)
        self.calls = []

    # callbacks SetupTab wires to
    def browse_source_folder(self): self.calls.append('browse_source')
    def browse_itc_template(self): self.calls.append('browse_itc')
    def browse_sales_template(self): self.calls.append('browse_sales')
    def browse_target_folder(self): self.calls.append('browse_target')
    def scan_files(self): self.calls.append('scan')
    def rescan_files(self): self.calls.append('rescan')
    def update_global_folder_setting(self): self.calls.append('global_setting')
    def save_cache(self): self.calls.append('save_cache')


@pytest.fixture
def built(root):
    notebook = tk.Frame(root)          # ttk.Notebook not required; .add is
    notebook.add = lambda *a, **k: None
    app = FakeApp(root)
    tab = SetupTab(notebook, app)
    return tab, app


def test_tab_builds_without_error(built):
    tab, app = built
    assert tab.tab_frame is not None
    assert app.scan_btn is not None
    assert app.client_name_check is not None   # kept for compatibility


def test_scan_button_starts_disabled_when_nothing_selected(built):
    tab, app = built
    assert str(app.scan_btn['state']) == 'disabled'


def test_scan_button_enables_once_source_is_valid(built, tmp_path):
    tab, app = built
    app.source_folder.set(str(tmp_path))
    tab.refresh_status(force=True)            # bypass the debounce timer
    assert str(app.scan_btn['state']) == 'normal'


def test_scan_button_disabled_again_for_a_bad_path(built, tmp_path):
    tab, app = built
    app.source_folder.set(str(tmp_path))
    tab.refresh_status(force=True)
    app.source_folder.set(str(tmp_path / "does_not_exist"))
    tab.refresh_status(force=True)
    assert str(app.scan_btn['state']) == 'disabled'


def test_checklist_reflects_states(built, tmp_path):
    tab, app = built
    tmpl = tmp_path / "itc.xltx"
    tmpl.write_bytes(b"x")
    app.source_folder.set(str(tmp_path))
    app.itc_template.set(str(tmpl))
    tab.refresh_status(force=True)

    assert tab._rows['source']['icon']['text'] == '✓'
    assert tab._rows['itc']['icon']['text'] == '✓'
    assert tab._rows['sales']['icon']['text'] == '○'      # not chosen yet
    assert 'Not selected yet' in tab._rows['target']['detail']['text']


def test_summary_updates(built, tmp_path):
    tab, app = built
    app.source_folder.set(str(tmp_path))
    tab.refresh_status(force=True)
    assert tab.summary_label['text'] == sv.summary_line(
        sv.evaluate_setup(str(tmp_path), '', '', ''))


def test_scan_command_is_wired(built, tmp_path):
    tab, app = built
    app.source_folder.set(str(tmp_path))
    tab.refresh_status(force=True)
    app.scan_btn.invoke()                     # only works via command=
    assert 'scan' in app.calls


def test_length_hint_uses_the_configured_limit(built):
    tab, app = built
    app.client_name_max_length.set(50)
    tab._update_length_hint()
    assert '50' in tab.length_hint['text']
