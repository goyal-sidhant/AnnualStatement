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
    """Build the tab and actually lay it out.

    The widgets must be packed and updated, otherwise the canvas has a height of
    1px and anything that depends on real geometry (yview, scrolling) is
    meaningless.
    """
    root.geometry('900x600')
    notebook = tk.Frame(root)          # ttk.Notebook not required; .add is
    notebook.add = lambda *a, **k: None
    notebook.pack(fill='both', expand=True)
    app = FakeApp(root)
    tab = SetupTab(notebook, app)
    tab.tab_frame.pack(fill='both', expand=True)
    root.update()
    yield tab, app
    notebook.destroy()
    root.update_idletasks()


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


class FakeWheel:
    """Stand-in for a Tk MouseWheel event."""
    def __init__(self, widget, delta):
        self.widget = widget
        self.delta = delta


def _make_scrollable(tab, root):
    """Ensure the scrollregion covers the (taller-than-viewport) content."""
    tab._canvas.configure(scrollregion=tab._canvas.bbox('all'))
    root.update_idletasks()
    first, last = tab._canvas.yview()
    assert last < 1.0, "content should overflow the viewport for these tests"


def _deep_child(tab):
    """A widget nested inside the scroll area (a checklist label)."""
    return tab._rows['source']['detail']


class TestScrolling:
    def test_wheel_over_a_child_widget_scrolls(self, built, root):
        """Regression: <Leave> fires on a container when the pointer moves onto
        its children, so an Enter/Leave binding switched scrolling off over the
        cards - i.e. over most of the tab."""
        tab, app = built
        _make_scrollable(tab, root)
        child = _deep_child(tab)
        assert tab._is_in_scroll_area(child)

        before = tab._canvas.yview()[0]
        tab._on_mousewheel(FakeWheel(child, -120))
        root.update_idletasks()
        assert tab._canvas.yview()[0] > before

    def test_small_delta_still_scrolls(self, built, root):
        """Regression: int(delta/120) truncates a precision touchpad's small
        deltas to 0, making the wheel look dead. Deltas must accumulate."""
        tab, app = built
        _make_scrollable(tab, root)
        child = _deep_child(tab)

        before = tab._canvas.yview()[0]
        for _ in range(6):                       # 6 x 40 == two full notches
            tab._on_mousewheel(FakeWheel(child, -40))
        root.update_idletasks()
        assert tab._canvas.yview()[0] > before

    def test_wheel_outside_the_tab_is_ignored(self, built, root):
        """The binding is application-level, so it must not steal wheel events
        belonging to other widgets."""
        tab, app = built
        _make_scrollable(tab, root)
        outsider = tk.Label(root, text='elsewhere')

        assert tab._is_in_scroll_area(outsider) is False
        before = tab._canvas.yview()[0]
        assert tab._on_mousewheel(FakeWheel(outsider, -120)) is None
        root.update_idletasks()
        assert tab._canvas.yview()[0] == before
        outsider.destroy()

    def test_no_scroll_when_everything_fits(self, built, root):
        """With nothing to scroll the handler must not swallow the event."""
        tab, app = built
        # A scrollregion shorter than the viewport => nothing to scroll.
        tab._canvas.configure(scrollregion=(0, 0, 800, 50))
        root.update_idletasks()
        assert tab._canvas.yview()[1] >= 1.0

        assert tab._on_mousewheel(FakeWheel(_deep_child(tab), -120)) is None

    def test_scrolling_up_at_the_top_stays_put(self, built, root):
        tab, app = built
        _make_scrollable(tab, root)
        tab._on_mousewheel(FakeWheel(_deep_child(tab), 120))   # wheel up
        root.update_idletasks()
        assert tab._canvas.yview()[0] == 0.0
