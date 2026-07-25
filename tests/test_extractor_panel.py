"""Tests for the ExtractorPanel / PowerQueryExtractorApp split.

The panel must behave identically whether it is hosted by its own standalone
window or embedded in another app's notebook, so the same code serves both.
"""
import tkinter as tk
from tkinter import ttk

import pytest

from power_query_extractor.gui.extractor_window import (
    ExtractorPanel, PowerQueryExtractorApp,
)

XL = b'PK\x03\x04' + b'\0' * 2048


@pytest.fixture
def root(tk_root):
    """The session-wide Tk root (see tests/conftest.py)."""
    return tk_root


@pytest.fixture
def output_tree(tmp_path):
    """A realistic organiser output tree with one client and two reports."""
    version = tmp_path / "Annual Statement-250726 1000" / "ABC-DL" / "Version-250726 1000"
    version.mkdir(parents=True)
    (version / "ITC_Report_ABC_DL.xlsx").write_bytes(XL)
    (version / "Sales_Report_ABC_DL.xlsx").write_bytes(XL)
    return tmp_path, version


@pytest.fixture
def embedded(root, output_tree):
    tmp_path, version = output_tree
    container = ttk.Frame(root)
    container.pack(fill='both', expand=True)
    panel = ExtractorPanel(container, target_folder=str(tmp_path),
                           show_header=False, show_status_bar=False)
    panel.pack(fill='both', expand=True)
    root.update_idletasks()
    yield panel, version
    container.destroy()
    root.update_idletasks()


class TestEmbeddedPanel:
    def test_builds_inside_an_arbitrary_parent(self, embedded):
        panel, _ = embedded
        assert panel.winfo_exists() == 1
        assert isinstance(panel, ttk.Frame)

    def test_finds_clients_from_the_target_folder(self, embedded):
        panel, _ = embedded
        assert 'ABC-DL' in panel.client_vars

    def test_host_chrome_is_suppressed(self, embedded):
        """Embedded, the host supplies the banner and status bar."""
        panel, _ = embedded
        assert panel.show_header is False
        assert panel.show_status_bar is False

    def test_has_its_own_setup_and_processing_tabs(self, embedded):
        panel, _ = embedded
        tabs = [panel.notebook.tab(i)['text']
                for i in range(panel.notebook.index('end'))]
        assert tabs == ['Setup', 'Processing']

    def test_logging_and_progress_work(self, embedded, root):
        panel, _ = embedded
        panel.log_message("embedded message", 'info')
        panel.update_progress(50, "halfway")
        root.update()                       # drain the scheduled UI callbacks
        assert "embedded message" in panel.log_text.get('1.0', 'end')
        assert panel.progress_var.get() == 50


class TestRefreshStatusUpdates:
    """Regression: the 'last refreshed' column never updated after a run."""

    def test_status_starts_as_never(self, embedded):
        panel, _ = embedded
        assert panel.client_vars['ABC-DL']['status_label'].cget('text') == 'Never refreshed'

    def test_status_updates_in_place_after_processing(self, embedded, root):
        panel, version = embedded
        (version / "ITC_Report_ABC_DL_Refreshed_250726_1130.xlsx").write_bytes(XL)

        panel.refresh_status_display()
        root.update_idletasks()

        text = panel.client_vars['ABC-DL']['status_label'].cget('text')
        assert text.startswith('ITC: ')
        assert 'Never' not in text.split('|')[0]

    def test_selections_survive_the_status_update(self, embedded, root):
        """Updating in place (not re-scanning) must keep the user's ticks."""
        panel, version = embedded
        panel.client_vars['ABC-DL']['itc_var'].set(False)
        panel.client_vars['ABC-DL']['sales_var'].set(True)

        (version / "ITC_Report_ABC_DL_Refreshed_250726_1130.xlsx").write_bytes(XL)
        panel.refresh_status_display()
        root.update_idletasks()

        assert panel.client_vars['ABC-DL']['itc_var'].get() is False
        assert panel.client_vars['ABC-DL']['sales_var'].get() is True
