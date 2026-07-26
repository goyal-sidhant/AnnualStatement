# gui/tabs/setup_tab.py
"""Setup Tab - Step 1 of GST File Organizer.

Shows a live readiness checklist: each required input reports its own state as
you pick it, and the Scan button enables itself as soon as the source folder is
valid (scanning does not need the templates - those gate Step 3).

Validation logic lives in gui/utils/setup_validation.py so it can be tested
without Tkinter. Existence checks are debounced and cached because on a network
share each one is a real round-trip.
"""

import tkinter as tk
from pathlib import Path
from tkinter import ttk

from utils.constants import (
    FILE_PATTERNS, GUI_CONFIG, PROCESSING_MODES, WORKFLOW_MODES,
)
from ..utils.ui_helpers import UIHelpers
from ..utils import setup_validation as sv
from ..widgets.tooltip import Tooltip

COLORS = GUI_CONFIG['colors']

# Per-state presentation for a checklist row: icon, text colour, row tint
STATE_STYLE = {
    sv.OK:      ('✓', COLORS['success'], '#E8F5E8'),
    sv.MISSING: ('○', '#8A8F98', '#F1F3F4'),
    sv.INVALID: ('✕', COLORS['danger'], '#FDECEA'),
}

# Accent colour + emoji per input, so each row is visually distinct
INPUT_ACCENT = {
    'source': (COLORS['success'], '📂'),
    'itc':    (COLORS['primary'], '📊'),
    'sales':  (COLORS['info'], '💰'),
    'target': (COLORS['warning'], '🎯'),
}

CARD_BORDER = '#DADCE0'

REVALIDATE_DELAY_MS = 350   # debounce: avoid stat'ing a UNC path on every keystroke

SCROLL_STEP_PX = 22         # pixels per scroll unit
SCROLL_UNITS_PER_NOTCH = 3  # a wheel notch moves 3 units (~= 3 text lines)


class SetupTab:
    """Setup tab for folder and template selection"""

    def __init__(self, notebook, app_instance):
        self.app = app_instance
        self.notebook = notebook
        self._rows = {}            # key -> dict of row widgets
        self._path_tooltips = {}   # key -> Tooltip for the entry
        self._revalidate_job = None
        self._exists_cache = {}
        self._canvas = None
        self._wheel_accum = 0.0
        self.create_tab()
        self._attach_traces()
        self.refresh_status(force=True)

    # ------------------------------------------------------------------ setup
    def create_tab(self):
        self.tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_frame, text="📁 Step 1: Setup")

        canvas = tk.Canvas(self.tab_frame, bg=COLORS['light'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.tab_frame, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg=COLORS['light'])

        scrollable.bind('<Configure>',
                        lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        window_id = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfigure(window_id, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        body = tk.Frame(scrollable, bg=COLORS['light'])
        body.pack(fill='both', expand=True, padx=20, pady=16)

        self._create_header(body)
        self._create_checklist(body)
        self._create_workflow(body)
        self._create_inputs(body)
        self._create_options(body)
        self._create_actions(body)
        self._create_reference(body)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self._canvas = canvas
        # A predictable step size: without this, 'units' scrolling depends on an
        # implementation-defined default and feels inconsistent.
        canvas.configure(yscrollincrement=SCROLL_STEP_PX)

        # Single application-level binding, filtered by whether the pointer is
        # actually inside this tab (see _on_mousewheel).
        #
        # NOT Enter/Leave on the canvas: Tk delivers <Leave> to a container when
        # the pointer moves onto one of its CHILDREN, so an Enter/Leave pair
        # would switch scrolling off over the cards - i.e. over almost the whole
        # tab. And NOT unbind_all, which would drop every other widget's
        # application-level wheel binding too.
        canvas.bind_all('<MouseWheel>', self._on_mousewheel, add='+')

    # --------------------------------------------------------------- scrolling
    def _is_in_scroll_area(self, widget):
        """True if `widget` is the scroll canvas or a descendant of it."""
        node = widget
        while node is not None:
            if node is self._canvas:
                return True
            node = getattr(node, 'master', None)
        return False

    def _on_mousewheel(self, event):
        canvas = self._canvas
        if canvas is None:
            return None
        try:
            if not canvas.winfo_exists():
                return None
        except tk.TclError:
            return None

        # Application-level binding: ignore wheel events aimed at other tabs or
        # at widgets that scroll themselves (the log box, comboboxes, ...).
        if not self._is_in_scroll_area(event.widget):
            return None

        # Nothing to scroll - let the event through rather than swallowing it.
        first, last = canvas.yview()
        if first <= 0.0 and last >= 1.0:
            return None

        # Accumulate fractional notches. High-resolution wheels and precision
        # touchpads send deltas smaller than 120, which int(delta/120) would
        # truncate to 0 - making the wheel appear dead.
        self._wheel_accum += -event.delta / 120.0 * SCROLL_UNITS_PER_NOTCH
        steps = int(self._wheel_accum)
        if steps:
            self._wheel_accum -= steps
            canvas.yview_scroll(steps, 'units')
        return 'break'

    def _accent_card(self, parent, accent, pady=(0, 12)):
        """A white card with a coloured accent bar down its left edge."""
        outer = tk.Frame(parent, bg=CARD_BORDER, relief='flat', borderwidth=0,
                         highlightbackground=CARD_BORDER, highlightthickness=1)
        outer.pack(fill='x', pady=pady)
        tk.Frame(outer, bg=accent, width=5).pack(side='left', fill='y')
        inner = tk.Frame(outer, bg='white')
        inner.pack(side='left', fill='both', expand=True)
        return inner

    def _create_header(self, parent):
        card = self._accent_card(parent, COLORS['primary'])

        row = tk.Frame(card, bg='white')
        row.pack(fill='x', padx=16, pady=(12, 4))

        tk.Label(row, text="📁  Step 1 — Setup", font=('Segoe UI', 15, 'bold'),
                 bg='white', fg=COLORS['primary']).pack(side='left')

        # Coloured readiness pill, updated by refresh_status
        self.summary_label = tk.Label(row, text="", font=('Segoe UI', 9, 'bold'),
                                      bg='#F1F3F4', fg='#5F6368', padx=10, pady=3)
        self.summary_label.pack(side='right')

        tk.Label(card,
                 text="Choose where your GST files are, the two report templates, "
                      "and where the organised output should go.",
                 font=('Segoe UI', 9), bg='white', fg='#5F6368',
                 justify='left', anchor='w').pack(fill='x', padx=16, pady=(0, 12))

    def _create_checklist(self, parent):
        card = self._accent_card(parent, COLORS['info'])

        tk.Label(card, text="CHECKLIST", font=('Segoe UI', 9, 'bold'),
                 bg='white', fg=COLORS['info'], anchor='w').pack(
                     fill='x', padx=16, pady=(10, 6))

        holder = tk.Frame(card, bg='white')
        holder.pack(fill='x', padx=16, pady=(0, 12))

        for key, label in (('source', 'Source folder'), ('itc', 'ITC template'),
                           ('sales', 'Sales template'), ('target', 'Target folder')):
            accent, emoji = INPUT_ACCENT[key]

            # Tinted strip per row - the tint itself carries the status colour
            row = tk.Frame(holder, bg='#F1F3F4')
            row.pack(fill='x', pady=2)

            icon = tk.Label(row, text='○', font=('Segoe UI', 12, 'bold'),
                            bg='#F1F3F4', fg='#8A8F98', width=3)
            icon.pack(side='left', pady=5)

            name = tk.Label(row, text=f"{emoji}  {label}", font=('Segoe UI', 10, 'bold'),
                            bg='#F1F3F4', fg=accent, width=20, anchor='w')
            name.pack(side='left', pady=5)

            detail = tk.Label(row, text='', font=('Segoe UI', 9), bg='#F1F3F4',
                              fg='#5F6368', anchor='w')
            detail.pack(side='left', fill='x', expand=True, pady=5)

            self._rows[key] = {'row': row, 'icon': icon, 'name': name, 'detail': detail}

    def _create_workflow(self, parent):
        """How far the run should go: reports only, or the full pipeline."""
        card = self._accent_card(parent, COLORS['warning'])

        tk.Label(card, text="WHAT SHOULD THIS RUN DO?", font=('Segoe UI', 9, 'bold'),
                 bg='white', fg='#8A5300', anchor='w').pack(
                     fill='x', padx=16, pady=(10, 6))

        holder = tk.Frame(card, bg='white')
        holder.pack(fill='x', padx=16, pady=(0, 12))

        for key, info in WORKFLOW_MODES.items():
            block = tk.Frame(holder, bg='white')
            block.pack(fill='x', pady=2)
            tk.Radiobutton(block, text=info['name'],
                           variable=self.app.workflow_mode, value=key,
                           font=('Segoe UI', 10, 'bold'), bg='white',
                           fg=COLORS['dark'], anchor='w',
                           command=self._on_workflow_change).pack(anchor='w')
            tk.Label(block, text=f"     {info['description']}",
                     font=('Segoe UI', 9), bg='white', fg='#5F6368',
                     anchor='w').pack(fill='x')

        tk.Label(holder,
                 text="Step 4 stays available either way - this only decides whether "
                      "Step 3 continues into it automatically.",
                 font=('Segoe UI', 8, 'italic'), bg='white', fg='#5F6368',
                 anchor='w', justify='left').pack(fill='x', pady=(6, 0))

    def _on_workflow_change(self):
        try:
            self.app.save_cache()
        except Exception:
            pass

    def _create_inputs(self, parent):
        section = UIHelpers.create_colored_section(
            parent, "📂 FOLDERS & TEMPLATES",
            "Pick each item - the checklist above updates as you go",
            COLORS['primary'])

        self._path_row(section, 'source', "Source folder (files to organise)",
                       self.app.source_folder, self.app.browse_source_folder)
        self._path_row(section, 'itc', "ITC report template",
                       self.app.itc_template, self.app.browse_itc_template)
        self._path_row(section, 'sales', "Sales report template",
                       self.app.sales_template, self.app.browse_sales_template)
        self._path_row(section, 'target', "Target folder (where output is created)",
                       self.app.target_folder, self.app.browse_target_folder)

    def _path_row(self, parent, key, label, variable, command):
        frame = tk.Frame(parent, bg='white')
        frame.pack(fill='x', padx=15, pady=(6, 2))

        tk.Label(frame, text=label, font=('Segoe UI', 9, 'bold'),
                 bg='white', fg=COLORS['dark'], anchor='w').pack(fill='x')

        line = tk.Frame(frame, bg='white')
        line.pack(fill='x', pady=(3, 0))

        entry = tk.Entry(line, textvariable=variable, font=('Segoe UI', 9),
                         relief='solid', borderwidth=1)
        entry.pack(side='left', fill='x', expand=True, ipady=4)
        self._path_tooltips[key] = Tooltip(entry, variable.get())

        # command= (not bind) so state='disabled' is honoured and the keyboard works
        ttk.Button(line, text="Browse…", command=command).pack(side='right', padx=(8, 0))

        problem = tk.Label(frame, text='', font=('Segoe UI', 8), bg='white',
                           fg=COLORS['danger'], anchor='w')
        problem.pack(fill='x')
        self._rows[key]['problem'] = problem

    def _create_options(self, parent):
        section = UIHelpers.create_colored_section(
            parent, "⚙️ PROCESSING MODE", "How the output folders are organised",
            COLORS['info'])

        holder = tk.Frame(section, bg='white')
        holder.pack(fill='x', padx=15, pady=10)

        for mode_key, mode_info in PROCESSING_MODES.items():
            block = tk.Frame(holder, bg='white')
            block.pack(fill='x', pady=2)
            tk.Radiobutton(block, text=mode_info['name'],
                           variable=self.app.processing_mode, value=mode_key,
                           font=('Segoe UI', 10, 'bold'), bg='white',
                           fg=COLORS['dark'], anchor='w').pack(anchor='w')
            tk.Label(block, text=f"     {mode_info['description']}",
                     font=('Segoe UI', 9), bg='white', fg='#5F6368',
                     anchor='w').pack(fill='x')

        opts = tk.Frame(holder, bg='white')
        opts.pack(fill='x', pady=(10, 0))

        self.app.client_name_check = tk.Checkbutton(
            opts, text="Include client name in Level 4 folder names",
            variable=self.app.include_client_name_in_folders,
            font=('Segoe UI', 10), bg='white',
            command=self.app.update_global_folder_setting)
        self.app.client_name_check.pack(anchor='w')

        tk.Label(opts, text="Overrides the per-client setting on the next tab.",
                 font=('Segoe UI', 8), bg='white', fg='#5F6368',
                 anchor='w').pack(fill='x', padx=22)

        length_row = tk.Frame(opts, bg='white')
        length_row.pack(fill='x', pady=(8, 0))
        tk.Label(length_row, text="Max client folder name length:",
                 font=('Segoe UI', 9), bg='white', fg=COLORS['dark']).pack(side='left')

        def on_length_change():
            if self.app.client_name_max_length.get() <= 0:
                self.app.client_name_max_length.set(35)
            self._update_length_hint()
            self.app.save_cache()

        tk.Spinbox(length_row, from_=15, to=100, width=5,
                   textvariable=self.app.client_name_max_length,
                   command=on_length_change, font=('Segoe UI', 9)).pack(side='left', padx=(8, 6))

        # Message reflects the ACTUAL configured limit (the old text hard-coded
        # "10 chars" while the default was 35).
        self.length_hint = tk.Label(length_row, text='', font=('Segoe UI', 8, 'italic'),
                                    bg='white', fg='#5F6368')
        self.length_hint.pack(side='left')
        self._update_length_hint()

    def _create_actions(self, parent):
        card = self._accent_card(parent, COLORS['success'], pady=(12, 12))

        inner = tk.Frame(card, bg='white')
        inner.pack(fill='x', padx=16, pady=12)

        self.app.scan_btn = ttk.Button(inner, text="🔍  Scan Files",
                                       command=self.app.scan_files)
        self.app.scan_btn.pack(side='left')

        self.app.rescan_btn = ttk.Button(inner, text="🔄  Re-scan",
                                         command=self.app.rescan_files)
        self.app.rescan_btn.pack(side='left', padx=8)

        # Status strip whose tint tracks readiness (set in _update_actions)
        self.action_bar = tk.Frame(card, bg='#F1F3F4')
        self.action_bar.pack(fill='x')
        self.action_hint = tk.Label(self.action_bar, text='', font=('Segoe UI', 9, 'bold'),
                                    bg='#F1F3F4', fg='#5F6368', anchor='w',
                                    padx=16, pady=7)
        self.action_hint.pack(fill='x')

    def _create_reference(self, parent):
        """Expected filename patterns, read from FILE_PATTERNS so this card can
        never drift from the regexes the parser actually uses."""
        card = self._accent_card(parent, '#7B1FA2', pady=(0, 20))

        tk.Label(card, text="📝  EXPECTED FILE NAMES", font=('Segoe UI', 9, 'bold'),
                 bg='white', fg='#7B1FA2', anchor='w').pack(
                     fill='x', padx=16, pady=(10, 2))
        tk.Label(card, text="Files in the source folder must follow these patterns "
                            "to be recognised.",
                 font=('Segoe UI', 8), bg='white', fg='#5F6368',
                 anchor='w').pack(fill='x', padx=16, pady=(0, 6))

        holder = tk.Frame(card, bg='white')
        holder.pack(fill='x', padx=16, pady=(0, 12))

        # Tint alternate rows so the list is easy to scan
        for idx, info in enumerate(FILE_PATTERNS.values()):
            bg = '#FAF5FC' if idx % 2 == 0 else 'white'
            row = tk.Frame(holder, bg=bg)
            row.pack(fill='x')
            tk.Label(row, text=info['type'], font=('Segoe UI', 8, 'bold'),
                     bg=bg, fg='#7B1FA2', width=16, anchor='w').pack(
                         side='left', padx=(6, 0), pady=3)
            tk.Label(row, text=info['description'], font=('Consolas', 8),
                     bg=bg, fg=COLORS['dark'], anchor='w').pack(
                         side='left', fill='x', expand=True, pady=3)

    # ------------------------------------------------------------- validation
    def _attach_traces(self):
        """Re-validate (debounced) whenever any path changes."""
        for var in (self.app.source_folder, self.app.itc_template,
                    self.app.sales_template, self.app.target_folder):
            var.trace_add('write', self._schedule_refresh)

    def _schedule_refresh(self, *_args):
        if self._revalidate_job:
            try:
                self.app.root.after_cancel(self._revalidate_job)
            except Exception:
                pass
        self._revalidate_job = self.app.root.after(REVALIDATE_DELAY_MS,
                                                   self.refresh_status)

    def _exists(self, path_str):
        """Existence check with a small cache - one round-trip per distinct path
        instead of one per repaint."""
        if path_str not in self._exists_cache:
            try:
                self._exists_cache[path_str] = Path(path_str).exists()
            except (OSError, ValueError):
                self._exists_cache[path_str] = False
        return self._exists_cache[path_str]

    def refresh_status(self, force=False):
        """Repaint the checklist and button states from current values."""
        self._revalidate_job = None
        if force:
            self._exists_cache.clear()

        checks = sv.evaluate_setup(
            self.app.source_folder.get(), self.app.itc_template.get(),
            self.app.sales_template.get(), self.app.target_folder.get(),
            exists=self._exists)

        for check in checks:
            row = self._rows.get(check.key)
            if not row:
                continue
            icon, colour, tint = STATE_STYLE[check.state]
            accent, emoji = INPUT_ACCENT[check.key]
            row['icon'].config(text=icon, fg=colour, bg=tint)
            row['detail'].config(text=check.message, bg=tint)
            row['name'].config(bg=tint, fg=accent)
            row['row'].config(bg=tint)
            if 'problem' in row:
                row['problem'].config(
                    text=check.message if check.state == sv.INVALID else '')

        for key, var in (('source', self.app.source_folder),
                         ('itc', self.app.itc_template),
                         ('sales', self.app.sales_template),
                         ('target', self.app.target_folder)):
            tip = self._path_tooltips.get(key)
            if tip:
                tip.set_text(var.get())

        # Readiness pill: green when everything is set, amber while incomplete
        if sv.can_process(checks):
            self.summary_label.config(text=sv.summary_line(checks),
                                      bg='#E8F5E8', fg=COLORS['success'])
        else:
            self.summary_label.config(text=sv.summary_line(checks),
                                      bg='#FFF4E5', fg='#8A5300')

        self._update_actions(checks)

    def _update_actions(self, checks):
        ready = sv.can_scan(checks)
        state = 'normal' if ready else 'disabled'
        for btn in (getattr(self.app, 'scan_btn', None),
                    getattr(self.app, 'rescan_btn', None)):
            if btn is not None:
                btn.config(state=state)

        if not ready:
            blocking = sv.blocking_labels(checks, sv.FOR_SCAN)
            text, fg, tint = (f"Select a valid {blocking[0].lower()} to scan",
                              COLORS['danger'], '#FDECEA')
        elif sv.can_process(checks):
            text, fg, tint = ("Ready to scan, and to process afterwards",
                              COLORS['success'], '#E8F5E8')
        else:
            pending = sv.blocking_labels(checks, sv.FOR_PROCESSING)
            text, fg, tint = ("Ready to scan · still needed for Step 3: "
                              + ", ".join(pending), '#8A5300', '#FFF4E5')

        self.action_hint.config(text=text, fg=fg, bg=tint)
        if getattr(self, 'action_bar', None) is not None:
            self.action_bar.config(bg=tint)

    def _update_length_hint(self):
        try:
            limit = self.app.client_name_max_length.get()
        except Exception:
            limit = 35
        self.length_hint.config(text=f"(default 35 · names over {limit} are shortened)")
