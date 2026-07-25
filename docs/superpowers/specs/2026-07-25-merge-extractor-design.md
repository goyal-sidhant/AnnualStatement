# Merge the Power Query Extractor into the main app

**Date:** 2026-07-25
**Status:** Approved, staged implementation

## Context

The tool ships as two separate programs. `main.py` opens the GST File Organizer
(organise files, create ITC/Sales reports). `launch_extractor.py` opens a second
window, the Power Query Extractor, which refreshes Power Query in those reports
and extracts values into a consolidated workbook. They communicate only through
`~/.gst_organizer_cache.json`.

Running them as two apps makes the normal end-to-end job feel disjointed: finish
one program, open another, re-pick the folder. The goal is a single continuous
flow, plus an explicit choice for users who only want files organised and reports
created without the (slow) Power Query refresh.

A hard constraint: **the app works today and must not break.** Every step below
is separately verifiable, and the first two steps are pure refactors with no
behaviour change.

## Decisions

| Decision | Choice |
|---|---|
| Layout | Add a fourth tab, "Step 4: Extract" |
| Workflow choice | Picked up front on the Setup tab |
| Step 4 independence | Always accessible, with its own folder picker |
| Standalone launcher | `launch_extractor.py` keeps working |

## Architecture

`PowerQueryExtractorApp` currently **is** a `tk.Tk` window, so it cannot be a tab.
Split it in two:

| Piece | What it is | Used by |
|---|---|---|
| `ExtractorPanel(ttk.Frame)` | All extractor UI + logic, parent-agnostic | both |
| `PowerQueryExtractorApp(tk.Tk)` | Thin shell: root window, hosts the panel | standalone |
| `gui/tabs/extract_tab.py` | "Step 4: Extract" tab, hosts the panel | main app |

One implementation, two homes.

### Window-level concerns move to the standalone shell

These are currently inside the extractor and would corrupt the merged app:

1. `title` / `geometry` / `minsize` (`extractor_window.py:109-111`).
2. `ttk.Style().theme_use('clam')` plus global `TNotebook` / `TNotebook.Tab`
   configuration (`extractor_window.py:118-124`). ttk styles are
   application-global, so embedding as-is would silently re-theme the whole
   main window.
3. `logging.basicConfig(... FileHandler('pq_extractor.log'))` at
   `extractor_main.py:12-19`, which runs on **import** and would hijack the main
   app's logging configuration. It moves into the standalone entry point.

### Thread safety (mandatory, not optional)

`_process_clients_thread` (`extractor_window.py:801`) updates the UI directly
from the worker thread: `log_message` (`:894`) inserts into the Text widget and
`update_progress` (`:887`) sets variables and calls `update_idletasks()`. As a
separate process with its own Tk loop this survives; sharing the main window's
Tk interpreter it is a genuine cause of freezes and crashes.

All worker-thread UI updates route through `root.after(0, ...)` with values bound
at schedule time (the pattern already proven in
`gui/handlers/processing_handler.py::_log_async`), and `update_idletasks()` is
removed from the worker thread.

## Workflow mode

New setting `workflow_mode`, chosen on the Setup tab and persisted in the cache:

- `organize` — organise files and create reports, then stop.
- `full` — also refresh Power Query and extract.

**The mode does not disable Step 4.** Because Step 4 must stay independently
usable, it is always accessible; the mode only decides whether Step 3
*automatically continues* into extraction. In `organize` mode Step 3 finishes
with a "Continue to extraction" button, so nothing is ever locked away.

## Step 4 behaviour

Its own folder picker, pre-filled from the Setup tab's target folder and
following it while the user has not overridden it. The common case needs no extra
clicks; pointing at an `Annual Statement-...` folder from a previous run still
works, which is why the two programs were separate in the first place (see
`PROJECT_AUDIT_ANNUALSTATEMENT.md`, decision D3).

## Staging

Each step is committed and verified on its own.

1. **Thread-safety fix** in the extractor, still standalone. Verify it works
   alone before anything moves.
2. **Extract `ExtractorPanel`**; the standalone shell uses it. Pure refactor.
3. **Add the Step 4 tab** hosting the panel in the main app. This is where the
   two apps actually meet.
4. **Workflow mode** on Setup, plus auto-advance from Step 3.
5. **Dark mode** rebuilt, once there is a single app to theme (see below).

Steps 1 and 2 change no behaviour and are characterised accordingly. If any step
proves risky it can stop there, leaving a working app.

## Testing

Existing coverage: `_classify_version_reports` and `_latest_refresh_times` are
unit-tested (`tests/test_extractor_helpers` equivalents live in the suite).

Added by this work:

- Smoke test: `ExtractorPanel` builds inside a plain `ttk.Frame`.
- Smoke test: the standalone shell still builds and hosts the panel.
- Smoke test: the main app builds with four tabs and Step 4 is reachable.
- Unit tests: workflow-mode gating (which steps run for each mode) as pure logic,
  no Tkinter.
- Thread-safety: assert worker-thread updates are scheduled rather than applied
  directly (same approach as the existing `_log_async` tests).

Manual check per step: launch both the merged app and `launch_extractor.py`.

## Out of scope

- Rewriting the Power Query refresh itself (the SendKeys approach and its fixed
  `time.sleep` waits stay as they are).
- The `process_files_thread` god-method refactor.
- Any change to how reports are produced.

## Deferred: dark mode

Dark mode is currently broken by design and was the original request that led
here. It is deliberately sequenced last, because theming one merged app is
simpler than theming two. Known defects, to be addressed then:

- `_restore_widget_colors` (`dark_mode_manager.py:172`) forces every Frame and
  Canvas to `white` rather than restoring the original; originals are only
  captured for Labels and Buttons, so one toggle permanently flattens every
  coloured surface (accent bars, section headers, Setup checklist state tints).
- `update_tree_tags` (`:198-206`): the `else` binds to `if hasattr(...)` instead
  of `if is_dark_mode`, so light mode never restores tree tags, and a widget
  without `tag_configure` would raise.
- Buttons are restored with `fg='white'` unconditionally (`:182`, `:184`).
- Labels are only darkened when their background is in a four-item whitelist
  (`:143`), so coloured labels stay light-on-light in dark mode.
- Only `TFrame`/`TLabel`/`TNotebook`/`Treeview` ttk styles are handled; not
  `TButton`, `TScrollbar` or `TEntry`.
- Widgets created after a toggle are never themed; both failure paths use bare
  `except: pass`.
