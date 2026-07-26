# gui/tabs/extract_tab.py
"""Step 4: Extract - hosts the Power Query Extractor inside the main app.

The extractor's UI lives in power_query_extractor.gui.extractor_window.
ExtractorPanel, which is deliberately parent-agnostic so the same code serves
both this tab and the standalone launch_extractor.py window.

Imported from extractor_window directly, NOT from extractor_main: that module
calls logging.basicConfig() at import time, which would hijack the main app's
logging configuration.
"""

import logging
import tkinter as tk
from tkinter import ttk

from utils.constants import GUI_CONFIG

logger = logging.getLogger(__name__)

# The extractor pulls in pywin32 for Excel COM. That is optional - the organizer
# itself works without it - so a failure here must NEVER stop the app starting.
# Step 4 shows an explanation instead.
try:
    from power_query_extractor.gui.extractor_window import ExtractorPanel
    EXTRACTOR_IMPORT_ERROR = None
except Exception as _import_error:          # pragma: no cover - platform dependent
    ExtractorPanel = None
    EXTRACTOR_IMPORT_ERROR = _import_error
    logger.warning("Power Query Extractor unavailable: %s", _import_error)

COLORS = GUI_CONFIG['colors']


class ExtractTab:
    """Wraps ExtractorPanel as the fourth step of the main flow."""

    def __init__(self, notebook, app_instance):
        self.app = app_instance
        self.notebook = notebook
        # Tracks the folder value this tab pushed in from Step 1, so a value the
        # user typed or browsed to is never overwritten.
        self._pushed_folder = None
        self.panel = None          # stays None if the extractor cannot be loaded
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

        if ExtractorPanel is None:
            self._create_unavailable_notice()
            return

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

    def _create_unavailable_notice(self):
        """Explain why Step 4 is inactive, instead of taking the app down."""
        card = tk.Frame(self.tab_frame, bg='#FFF4E5', relief='solid', borderwidth=1)
        card.pack(fill='x', padx=20, pady=20)
        tk.Label(card, text="⚠️  Power Query refresh is unavailable",
                 font=('Segoe UI', 11, 'bold'), bg='#FFF4E5', fg='#8A5300',
                 anchor='w').pack(fill='x', padx=14, pady=(12, 4))
        tk.Label(card,
                 text=("This step needs pywin32 (Excel automation), which is not "
                       "installed for the Python running this app.\n\n"
                       "Install it with:    pip install pywin32\n\n"
                       "Note that 'py' and 'python' can be different interpreters - "
                       "install it for the one you launch the app with.\n\n"
                       f"Details: {EXTRACTOR_IMPORT_ERROR}"),
                 font=('Segoe UI', 9), bg='#FFF4E5', fg='#5F6368',
                 anchor='w', justify='left', wraplength=900).pack(
                     fill='x', padx=14, pady=(0, 12))
        tk.Label(self.tab_frame,
                 text="Steps 1-3 are unaffected: organising files and creating the "
                      "reports works without pywin32.",
                 font=('Segoe UI', 9, 'italic'), bg=COLORS['light'],
                 fg='#5F6368', anchor='w').pack(fill='x', padx=20)

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
        if self.panel is None:
            return
        current = self.panel.folder_path.get()
        # Only push the new value through if the box still holds what we put
        # there - never clobber a folder the user chose in this tab.
        if current and current != self._pushed_folder:
            return
        new_value = self._target_folder_value() or ''
        self.panel.folder_path.set(new_value)
        self._pushed_folder = new_value
