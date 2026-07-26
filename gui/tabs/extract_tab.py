# gui/tabs/extract_tab.py
"""Step 4: Extract - hosts the Power Query Extractor inside the main app.

The extractor's UI lives in power_query_extractor.gui.extractor_window.
ExtractorPanel, which is deliberately parent-agnostic so the same code serves
both this tab and the standalone launch_extractor.py window.

Imported from extractor_window directly, NOT from extractor_main: that module
calls logging.basicConfig() at import time, which would hijack the main app's
logging configuration.
"""

import tkinter as tk
from tkinter import ttk

from utils.constants import GUI_CONFIG
from power_query_extractor.gui.extractor_window import ExtractorPanel

COLORS = GUI_CONFIG['colors']


class ExtractTab:
    """Wraps ExtractorPanel as the fourth step of the main flow."""

    def __init__(self, notebook, app_instance):
        self.app = app_instance
        self.notebook = notebook
        # Tracks the folder value this tab pushed in from Step 1, so a value the
        # user typed or browsed to is never overwritten.
        self._pushed_folder = None
        self.create_tab()
        self._attach_target_folder_sync()

    def create_tab(self):
        self.tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_frame, text="🔄 Step 4: Extract")

        # A short strip explaining what this step does, since the panel's own
        # banner is suppressed below.
        header = tk.Frame(self.tab_frame, bg='white', relief='solid', borderwidth=1)
        header.pack(fill='x', padx=20, pady=(12, 0))
        tk.Label(header, text="🔄  Step 4 — Refresh Power Query & extract data",
                 font=('Segoe UI', 12, 'bold'), bg='white',
                 fg=COLORS['primary'], anchor='w').pack(fill='x', padx=14, pady=(10, 2))
        tk.Label(header,
                 text="Opens the ITC and Sales reports produced in Step 3, refreshes their "
                      "Power Query connections, and collects the results into one workbook. "
                      "Press Scan to list the clients in the folder below.",
                 font=('Segoe UI', 9), bg='white', fg='#5F6368',
                 anchor='w', justify='left', wraplength=980).pack(
                     fill='x', padx=14, pady=(0, 10))

        # The extractor itself. Header and status bar are suppressed because the
        # main window already provides both. auto_load=False so building this tab
        # at startup does not trigger a network scan.
        self.panel = ExtractorPanel(
            self.tab_frame,
            target_folder=self._target_folder_value(),
            show_header=False,
            show_status_bar=False,
            auto_load=False,
        )
        self.panel.pack(fill='both', expand=True, padx=20, pady=(8, 12))
        self._pushed_folder = self.panel.folder_path.get()

    # ------------------------------------------------------------------ folder
    def _target_folder_value(self):
        try:
            return self.app.target_folder.get() or None
        except Exception:
            return None

    def _attach_target_folder_sync(self):
        """Follow Step 1's target folder until the user overrides it here."""
        try:
            self.app.target_folder.trace_add('write', self._on_target_folder_changed)
        except Exception:
            pass

    def _on_target_folder_changed(self, *_args):
        current = self.panel.folder_path.get()
        # Only push the new value through if the box still holds what we put
        # there - never clobber a folder the user chose in this tab.
        if current and current != self._pushed_folder:
            return
        new_value = self._target_folder_value() or ''
        self.panel.folder_path.set(new_value)
        self._pushed_folder = new_value
