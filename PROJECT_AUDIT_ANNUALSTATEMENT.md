# PROJECT AUDIT: GST File Organizer & Report Generator

**Audit Date:** 2026-03-05
**Auditor:** Claude Opus 4.6 (AI-assisted reverse engineering)
**Project Root:** `d:\OneDrive 2\OneDrive\Python 2026\AnnualStatement`
**Current Version:** v3.0 (Production Ready) — Tkinter GUI on `main` branch
**Active Branch:** `main` at commit `fc279af`
**Unmerged Branch:** `copilot-worktree-2026-03-04T11-52-56` — attempted PyQt5 migration (3 commits ahead of main, NOT merged)

---

## SECTION 1: EXECUTIVE SUMMARY

This tool is a desktop application built for **CA/GST consultants** (Chartered Accountants working with India's Goods and Services Tax system). It solves a specific, time-consuming problem: when a consultant manages annual GST filings for dozens of clients, they receive multiple Excel files per client (GSTR-2B reconciliation, IMS reconciliation, GSTR-3B exports, Sales data, Sales reconciliation, Annual Reports) that must be organized into a standardized folder structure, linked to Excel report templates, and optionally processed through Power Query to extract key differences and totals across all clients. Without this tool, a consultant would manually create folders, copy files, open each Excel template, type in file paths, refresh Power Query connections, and copy out cell values — for every single client. This tool automates that entire workflow.

The core technology stack is: **Python 3.7+** for the language, **Tkinter** (Python's built-in GUI library) for the desktop interface, **openpyxl** for creating and reading Excel files, and **win32com/pywin32** for automating Microsoft Excel on Windows (required for Power Query refresh). The application has two entry points: `main.py` launches the main organizer GUI, and `launch_extractor.py` launches a separate Power Query Extractor tool.

The project is **actively maintained and production-used**. The git history shows real-world bug fixes (circular imports, threading issues, long path handling) suggesting it has been used with actual client data. A PyQt5 migration was attempted (2026-03-04) on a separate branch but **was never merged to main** and the branch was subsequently abandoned — the production code remains Tkinter-based. The most recent commit on `main` (fc279af) added the initial project audit document.

Current state: **Functional but with significant issues**. The main organizer works end-to-end on Windows with Tkinter. The Power Query Extractor works but has hardcoded Windows dependencies (`import pythoncom`, `import win32com.client`) without graceful fallback at the module level, meaning it will crash on import on non-Windows platforms. There is a stale version reference in `main.py` docstring (says v3.0) but the UI says v3.0 too, which is consistent. The `requirements.txt` is incomplete — it lists `openpyxl` and `lxml` but omits `pywin32`, which is critical for the application's core Power Query functionality.

---

## SECTION 2: PROJECT STRUCTURE MAP

```
AnnualStatement/                         [ROOT - Entry point directory]
├── main.py                              [CORE] Entry point - launches main GUI app
├── launch_extractor.py                  [CORE] Separate entry point for Power Query Extractor
├── requirements.txt                     [CONFIG] Python dependencies (INCOMPLETE - missing pywin32)
├── README.md                            [DOCS] Comprehensive usage documentation (364 lines)
├── PROJECT_AUDIT_ANNUALSTATEMENT.md     [DOCS] This audit document
├── .gitignore                           [CONFIG] Git exclusion rules (79 lines)
│
├── .claude/                             [TOOLING] Claude Code AI assistant settings
│   └── settings.local.json              [TOOLING] Permits `python -m py_compile` via Bash
│
├── core/                                [CORE] Business logic layer
│   ├── __init__.py                      [CORE] Package init - exports FileParser, FileOrganizer, ExcelHandler
│   ├── file_parser.py                   [CORE] Scans folders, matches filenames to GST patterns, groups by client
│   ├── file_organizer.py                [CORE] Creates folder structure, copies files to organized locations
│   └── excel_handler.py                 [CORE] Creates Excel reports from templates, writes summary reports
│
├── utils/                               [UTIL] Shared utilities and constants
│   ├── __init__.py                      [UTIL] Package init (minimal, just docstring)
│   ├── constants.py                     [CONFIG] All regex patterns, folder templates, GUI config, state codes
│   └── helpers.py                       [UTIL] File operations, path utilities, validation, progress tracking
│
├── gui/                                 [UI] Tkinter GUI layer
│   ├── __init__.py                      [UI] Package init - exports GSTOrganizerApp from main_window
│   ├── main_window.py                   [UI] Main application window - ties all components together
│   ├── handlers/                        [UI] Event handler modules
│   │   ├── __init__.py                  [UI] Package init (empty docstring)
│   │   ├── file_handler.py              [UI] Browse dialogs, scan validation
│   │   ├── client_handler.py            [UI] Client tree selection, keyboard navigation, export
│   │   └── processing_handler.py        [UI] Processing thread, report creation, progress updates
│   ├── tabs/                            [UI] Tab panel modules
│   │   ├── __init__.py                  [UI] Package init (empty docstring)
│   │   ├── setup_tab.py                 [UI] Step 1 - folder/template selection
│   │   ├── validation_tab.py            [UI] Step 2 - client review and selection
│   │   └── processing_tab.py            [UI] Step 3 - progress display and log
│   ├── utils/                           [UI] GUI utility modules
│   │   ├── __init__.py                  [UI] Package init (empty docstring)
│   │   ├── cache_manager.py             [UI] JSON-based settings persistence (~/.gst_organizer_cache.json)
│   │   ├── dark_mode_manager.py         [UI] Theme switching between light and dark
│   │   ├── status_bar.py                [UI] Bottom status bar widget
│   │   └── ui_helpers.py                [UI] Static methods for log messages, progress, colored sections
│   └── widgets/                         [UI] Reusable widget components
│       ├── __init__.py                  [UI] Package init (empty docstring)
│       ├── collapsible_frame.py         [UI] Expandable/collapsible frame widget
│       └── title_bar.py                 [UI] Title bar with dark mode toggle
│
└── power_query_extractor/               [CORE] Separate sub-application for PQ operations
    ├── __init__.py                      [CORE] Package init (version only)
    ├── extractor_main.py                [CORE] Entry point class for PQ extractor
    ├── config/                          [CONFIG] Extraction configuration
    │   ├── __init__.py                  [CONFIG] Package init (empty)
    │   └── cell_mappings.py             [CONFIG] Master cell mapping configuration (editable by user)
    ├── core/                            [CORE] Processing logic
    │   ├── __init__.py                  [CORE] Package init (empty)
    │   ├── report_processor.py          [CORE] Excel COM automation for Power Query refresh
    │   └── data_consolidator.py         [CORE] Creates consolidated Excel reports from extracted data
    └── gui/                             [UI] Extractor GUI
        ├── __init__.py                  [UI] Package init (empty)
        └── extractor_window.py          [UI] Tkinter window for PQ extractor

```

### Entry Points

| File | Command | Purpose |
|------|---------|---------|
| `main.py` | `python main.py` | Launches the main GST File Organizer GUI |
| `launch_extractor.py` | `python launch_extractor.py` | Launches the Power Query Extractor GUI |

### Dead Code Analysis

| Type | Location | Evidence |
|------|----------|----------|
| Dead method | `helpers.py:422` `calculate_file_hash()` | Never called from any file in the project. Grep for `calculate_file_hash` returns only its definition. |
| Dead method | `helpers.py:446` `extract_filename_without_extension()` | This IS called from `excel_handler.py:576,588,598,614,631,646` via import. NOT dead. |
| Dead method | `data_consolidator.py:250` `_write_section()` | Never called from anywhere. The `create_report()` method was rewritten to use `EXTRACTION_CONFIG` directly, making this legacy method dead. |
| Dead method | `data_consolidator.py:354` `_create_details_sheet()` | Never called from anywhere. Was likely used in an earlier version before `EXTRACTION_CONFIG` was introduced. |
| Dead constant | `constants.py:180` `TEMPLATE_EXTENSIONS` | Never referenced outside its definition. |
| Dead constant | `constants.py:182` `EXCEL_SIGNATURES` | Never referenced outside its definition. The actual signature checking in `helpers.py:240-241` uses hardcoded byte values instead of this constant. |
| Dead method | `cache_manager.py:50-55` `get_cached_value()` and `set_cached_value()` | Never called from any file. Only `load_cache()` and `save_cache()` are used. |
| Dead parameter | `file_parser.py:308` `get_statistics()` return value `'parsing_rate'` | Computed but displayed in `main_window.py:280` — actually IS used. NOT dead. |
| Dead class reference | `main.py:72` references `gui/app.py` in error message | The file `gui/app.py` does not exist. The actual file is `gui/main_window.py`. This is a stale error message from an earlier structure. |
| Stale comment | `helpers.py:15` says `# ADD THIS LINE` | AI-generated instruction comment left in production code. |
| Stale comment | `file_parser.py:16` says `# ADD THIS` | Same issue — leftover AI instruction. |
| Stale comment | `file_organizer.py:15` says `# ADD THESE` | Same issue. |
| Stale comment | `excel_handler.py:39` says `# ADD THESE TWO` | Same issue. |

---

## SECTION 3: FEATURE INVENTORY

### F1: File Scanning and Pattern Matching

**What it does:** Scans a user-selected folder for Excel files, matches each filename against 6 predefined regex patterns (GSTR-2B Reco, IMS Reco, GSTR-3B, Sales, Sales Reco, Annual Report), extracts client name and state from the filename, and groups files by client.

**Where it lives:** `core/file_parser.py` (FileParser class), `utils/constants.py` (FILE_PATTERNS dict, lines 13-62)

**Input → Output:**
- Input: A folder path containing Excel files with standardized names
- Output: Three data structures — `scanned_files` (dict of all parsed files), `client_data` (dict keyed by "ClientName-StateCode"), `variations` (list of unparseable files)

**User interaction:** User clicks "Scan Files" button on Setup tab. Progress updates shown in status bar.

**Concrete trace:** User has folder with files: `GSTR3B-Reliance-Maharashtra-Apr24.xlsx`, `Sales-Reliance-Maharashtra-Apr-Jun.xlsx`, `AnnualReport-Reliance-Maharashtra-2024.xlsx`.
1. `scan_folder()` calls `find_excel_files()` which globs for *.xlsx/*.xls/*.xlsm
2. Each file is validated via `validate_excel_file()` (checks extension, minimum 1KB size, ZIP/OLE signature)
3. `parse_filename()` tries each regex pattern — GSTR3B pattern matches `GSTR3B-Reliance-Maharashtra-Apr24.xlsx`, extracts client="Reliance", state="Maharashtra"
4. `_add_to_client_data()` converts "Maharashtra" to "MH" via `get_state_code()`, creates key "Reliance-MH"
5. `_analyze_client_completeness()` checks found types vs EXPECTED_FILE_TYPES — finds 3 of 6, marks as "Missing 3 files"
6. The GSTR3B regex `^(?:\(\d+\)\s*)?GSTR3B-` also handles `(1) GSTR3B-...` prefix (Windows duplicate download naming)

**Bug found during trace:** The `_analyze_client_completeness()` method at line 278 checks `if not missing and not extras` for "Complete" status. But `extras` only contains strings about duplicates, not about unexpected extra files — so a client with one file of each type but also a seventh unrecognized file would still be "Complete". This is arguably correct behavior since unrecognized files go to `variations`, not to client data.

### F2: Folder Structure Creation

**What it does:** Creates a hierarchical folder structure in the target directory for each client being processed.

**Where it lives:** `core/file_organizer.py` (FileOrganizer class)

**Input → Output:**
- Input: Target folder path, processing mode (fresh/rerun/resume), client info dict
- Output: Dict mapping folder keys to Path objects:
  - `level1`: `Annual Statement-DDMMYY HHMM/`
  - `level2`: `ClientName-StateCode/`
  - `version`: `Version-DDMMYY HHMM/`
  - `gstr3b`: `GSTR-3B Exports (ClientName)/` or `GSTR-3B Exports/`
  - `itc`: `Other ITC related files (ClientName)/` or `Other ITC related files/`
  - `sales`: `Sales related files (ClientName)/` or `Sales related files/`

**User interaction:** Automatic during processing. The user controls whether client names appear in parentheses via a checkbox.

**Concrete trace:** Fresh run for client "Reliance-MH", include_client_name=False:
1. `_get_level1_folder()` creates `Annual Statement-050326 1430/` (current timestamp)
2. `create_client_state_key("Reliance", "Maharashtra", max_length=35)` → "Reliance-MH"
3. Level 2 folder: `Annual Statement-050326 1430/Reliance-MH/`
4. Version folder: `Annual Statement-050326 1430/Reliance-MH/Version-050326 1430/`
5. Category folders (without client name): `GSTR-3B Exports/`, `Other ITC related files/`, `Sales related files/`

### F3: File Organization (Copy and Sanitize)

**What it does:** Copies source Excel files into the appropriate category folders, sanitizing filenames (replacing "Private" with "Pvt", "Limited" with "Ltd", removing invalid characters).

**Where it lives:** `core/file_organizer.py` (organize_files method, lines 181-231)

**Input → Output:**
- Input: Client info with file lists, folder structure dict
- Output: Files physically copied to destination folders. Returns list of operation results (Success/Failed/Skipped).

**Key behavior:** After copying, the method updates `client_info['files'][file_type][i]['name']` with the sanitized filename. This is critical because the ExcelHandler later reads these names to populate template cells. If the original filename had characters like apostrophes, the sanitized version is what gets written to the template.

**Concrete trace:** Source file `Sales-Reliance Private Limited-Maharashtra-Apr-Jun.xlsx`:
1. `sanitize_filename()` replaces "Private" → "Pvt", "Limited" → "Ltd"
2. Result: `Sales-Reliance_Pvt_Ltd-Maharashtra-Apr-Jun.xlsx` (spaces become underscores via regex)
3. File copied to `Sales related files/Sales-Reliance_Pvt_Ltd-Maharashtra-Apr-Jun.xlsx`
4. Size verification: source and destination sizes compared; if mismatch, dest deleted and returns False
5. `client_info['files']['Sales'][0]['name']` updated to `Sales-Reliance_Pvt_Ltd-Maharashtra-Apr-Jun.xlsx`

### F4: Excel Report Generation (ITC and Sales)

**What it does:** Creates ITC and Sales Excel reports by copying template files and writing folder paths and filenames into specific cells of a "Links" sheet. On Windows, uses Excel COM automation (win32com) to preserve Power Query connections. Falls back to openpyxl on other platforms (loses Power Query).

**Where it lives:** `core/excel_handler.py` (ExcelHandler class, lines 65-274)

**Input → Output:**
- Input: Template file path, output path, data mappings (folder paths and filenames), report type ("ITC" or "Sales")
- Output: New Excel file at output path with Links sheet cells populated

**Template cell mappings** (from `constants.py:96-120`):

ITC Template Links sheet:
| Cell | Data Key | Value Written |
|------|----------|---------------|
| B2 | gstr3b_folder | Windows path to GSTR-3B folder |
| B4 | annual_folder | Windows path to Version folder |
| B5 | annual_filename | Annual report filename without extension |
| B7 | gstr2b_folder | Windows path to ITC folder |
| B8 | gstr2b_filename | GSTR-2B Reco filename without extension |
| B10 | ims_folder | Windows path to ITC folder |
| B11 | ims_filename | IMS Reco filename without extension |

Sales Template Links sheet:
| Cell | Data Key | Value Written |
|------|----------|---------------|
| B2 | sales_folder | Windows path to Sales folder |
| B3 | sales_filename | Sales filename without extension |
| B5 | annual_folder | Windows path to Version folder |
| B6 | annual_filename | Annual report filename without extension |
| B8 | sales_reco_folder | Windows path to Sales folder |
| B9 | sales_reco_filename | Sales Reco filename without extension |

**Win32COM path (lines 103-228):**
1. Calls `pythoncom.CoInitialize()`
2. Tries `GetActiveObject("Excel.Application")` first (reuses running Excel), falls back to `Dispatch`
3. Copies template to output location via `shutil.copy2`
4. Opens copied file via `_safe_open_workbook()` which handles long paths (>218 chars) with 8.3 short path fallback
5. Finds "Links" sheet (case-insensitive, also tries partial match on "link")
6. Writes values to cells via `ws.Range(cell_ref).Value = str(value)`
7. Saves and closes
8. Calls `pythoncom.CoUninitialize()` in finally block

**Concrete trace:** ITC report for Reliance-MH:
1. Template copied to `Version-050326 1430/ITC_Report_Reliance_MH_050326_1430.xlsx`
2. Excel COM opens the copy
3. Finds "Links" sheet
4. Writes: B2 = `D:\target\Annual Statement-050326 1430\Reliance-MH\Version-050326 1430\GSTR-3B Exports`
5. Writes: B4 = version folder path, B5 = `AnnualReport-Reliance-MH-2024` (no extension)
6. Saves and closes

### F5: Summary Report Generation

**What it does:** Creates a comprehensive Excel summary report at the Level 1 folder (Annual Statement root) containing 5 sheets: Summary, Client Status, File Mapping, Errors, and Variations.

**Where it lives:** `core/excel_handler.py` (create_summary_report method, lines 319-357), sheets created by private methods `_create_summary_sheet`, `_create_client_status_sheet`, etc.

**Input → Output:**
- Input: Output path, report_data dict containing all processing results
- Output: `GST_Processing_Summary_DDMMYY_HHMM.xlsx` with 5 formatted sheets

**User interaction:** Automatic at end of processing. Not user-triggered separately.

### F6: Dry Run Preview

**What it does:** Shows what would happen during processing without actually copying files or creating reports.

**Where it lives:** `gui/handlers/processing_handler.py` (dry_run method, lines 23-52)

**Input → Output:**
- Input: Selected client list
- Output: Log messages showing which files would be organized and which reports would be created

**Bug found:** Lines 41-48 use `lambda` inside a loop but capture `client_info` by reference, not by value. Since `lambda: self.app.log_message(...)` captures the variable `client_info`, and the loop reassigns it, all deferred `root.after` calls will log the LAST client's info, not each client's. The same bug exists for `file_info` on line 46. This means the dry run log will show incorrect filenames — it will repeat the last file's name for all entries.

> **This is a confirmed silent bug**: The dry run output looks correct in structure but shows wrong file/client names. The actual processing (F8) has the same pattern but with a different impact — since processing creates real files, the visible output (folders, files) is correct, but the log messages displayed to the user during processing show stale/incorrect values for the same reason.

### F7: Client Selection and Management

**What it does:** Displays clients in a treeview with checkboxes, supports keyboard navigation (arrow keys + space), selection helpers (Select All, Clear All, Complete Only), and per-client folder name settings.

**Where it lives:** `gui/handlers/client_handler.py` (ClientHandler class), `gui/tabs/validation_tab.py` (creates the treeview)

**Input → Output:**
- Input: Scanned client data
- Output: Visual tree display with selection state

**User interaction:** Click tree column #0 to toggle selection, click FolderName column (#7) to toggle per-client folder name inclusion. Space key toggles selection on focused item.

**Concrete trace:** 10 clients scanned, user clicks "Complete Only":
1. `select_complete_clients()` iterates all tree items
2. Checks if column index 2 (Status) contains "Complete"
3. Sets checkbox to ☑ for complete clients, ☐ for others
4. Updates status bar: "Selected 6 complete clients"

### F8: Full Processing Pipeline

**What it does:** The main processing operation that ties everything together — creates folders, copies files, generates ITC and Sales reports for each selected client, creates organization reports, and produces a master summary.

**Where it lives:** `gui/handlers/processing_handler.py` (process_files_thread method, lines 119-463)

**Input → Output:**
- Input: List of selected client keys
- Output: Complete organized folder structure with all files and reports

**User interaction:** "Start Processing" button. Runs in a background daemon thread. Stop button available. Progress bar and detailed log updated in real-time via `root.after(0, ...)`.

**Thread safety note:** All UI updates from the processing thread use `self.app.root.after(0, lambda: ...)` to marshal calls to the main thread. This was added in commit `6237597` to fix threading bugs. However, the lambda variable capture bug (see F6) means log messages may show stale values.

### F9: Dark Mode

**What it does:** Toggles between light and dark themes across the entire application.

**Where it lives:** `gui/utils/dark_mode_manager.py` (DarkModeManager class)

**Input → Output:**
- Input: Dark mode checkbox toggle
- Output: All widgets re-themed. Setting persisted to cache.

**Implementation:** Recursively walks all widgets via `winfo_children()`, stores original colors on first run, applies dark palette (`#2b2b2b` bg, `#ffffff` fg, `#3b3b3b` widget bg) or restores originals.

### F10: Settings Persistence (Cache)

**What it does:** Saves and loads application settings (source folder, template paths, target folder, processing mode, dark mode state, client name settings) to a JSON file in the user's home directory.

**Where it lives:** `gui/utils/cache_manager.py` (CacheManager class)

**Cache file location:** `~/.gst_organizer_cache.json` (user's home directory)

**Shared state verification:**
- **Writer:** `main_window.py:328-340` calls `save_cache()` with keys: `source_folder`, `itc_template`, `sales_template`, `target_folder`, `processing_mode`, `include_client_name`, `dark_mode`, `client_name_max_length`
- **Reader:** `main_window.py:342-356` calls `apply_cached_values()` reading the same keys
- **Cross-app reader:** `power_query_extractor/extractor_main.py:36-60` and `extractor_window.py:473-485` both read `target_folder` from `gst_organizer_cache.json`
- **Mismatch found:** The main app writes cache to `Path.home() / '.gst_organizer_cache.json'` (lines cache_manager.py:15). But `extractor_main.py:39` looks for `gst_organizer_cache.json` (in current directory, no dot prefix, no home directory). And `extractor_window.py:28` looks for `gst_organizer_cache.json` (also current directory). **This means the PQ Extractor will NOT find the main app's cache** unless the current working directory is the user's home directory OR a separate `gst_organizer_cache.json` exists in the project directory. This is a **confirmed bug** — the cache paths are inconsistent between the two apps.

### F11: Power Query Refresh and Data Extraction

**What it does:** Opens Excel reports via COM automation, triggers Power Query refresh (via SendKeys Alt+A, R, A), waits for completion, validates results, then extracts specific cell values from designated sheets.

**Where it lives:** `power_query_extractor/core/report_processor.py` (ReportProcessor class)

**Input → Output:**
- Input: Client data with paths to report files
- Output: Refreshed Excel files (saved with suffix), extracted cell values dict

**Key implementation details:**
1. Creates a copy of the original report with configurable suffix (default: `_Refreshed_{timestamp}`)
2. Opens Excel COM instance (visible, to avoid COM issues)
3. Sends keyboard shortcut Alt+A → R → A (Refresh All) via `WScript.Shell.SendKeys`
4. Waits configurable time (default 10 seconds, range 5-60)
5. Checks for error dialogs by enumerating Windows windows
6. Validates refresh by checking cell BB2 in "Info" sheet for non-empty/non-zero/non-error value
7. Saves and closes
8. Opens a NEW Excel instance (hidden) for data extraction
9. Reads cells according to `CELL_MAPPINGS` config

**Bug: Unreachable code** at line 342-345 in `_refresh_power_query_simple()`: There's a `return {'success': True, 'error': None}` after the try block's except clause. This return is unreachable because the try block (starting at line 217) always returns in either the try body (line 329) or the except body (line 337-340). The unreachable return at 342 never executes.

### F12: Data Consolidation Report

**What it does:** Creates a multi-sheet Excel report consolidating extracted values from all processed clients, with one sheet per configured extraction (GSTR-3B Difference, GSTR-1 Difference, Turnover, Reco Diff, etc.).

**Where it lives:** `power_query_extractor/core/data_consolidator.py` (DataConsolidator class)

**Input → Output:**
- Input: List of processing results (one per client), output folder path
- Output: `PQ_Extraction_Report_YYYYMMDD_HHMM.xlsx` with sheets for each EXTRACTION_CONFIG entry plus a Summary sheet

### F13: Export Client List

**What it does:** Exports selected client names to a text file and opens it.

**Where it lives:** `gui/handlers/client_handler.py` (export_client_list method, lines 221-253)

**Platform concern:** Uses `os.startfile(filename)` at line 251, which is Windows-only. Will crash on macOS/Linux.

---

## SECTION 4: DATA FLOW

### Main Application Flow

```
User selects source folder
         ↓
    [SCAN FILES]
         ↓
Source Folder → find_excel_files() → list of .xlsx/.xls paths
         ↓
For each file → validate_excel_file() → check signature (PK\x03\x04 or \xd0\xcf\x11\xe0)
         ↓
For each valid file → parse_filename() → regex match against 6 patterns
         ↓
Matched files → _add_to_client_data() → grouped by "ClientName-StateCode" key
         ↓
Unmatched files → variations list (with suggested correct pattern)
         ↓
_analyze_client_completeness() → marks each client Complete/Missing N files
         ↓
Results displayed in GUI: summary text + client treeview
         ↓
    [USER SELECTS CLIENTS]
         ↓
    [START PROCESSING]
         ↓
For each selected client:
    ├── FileOrganizer.create_client_structure() → creates folder hierarchy
    ├── FileOrganizer.organize_files() → copies files to category folders
    │   └── sanitize_filename() → clean file names
    │   └── safe_copy_file() → copy with size verification
    ├── ExcelHandler.prepare_template_data() → build cell→value mappings
    ├── ExcelHandler.create_report_from_template('ITC') → ITC report
    │   ├── [Windows] win32com: copy template, open, write Links sheet, save
    │   └── [Other] openpyxl: copy template, open, write Links sheet, save
    ├── ExcelHandler.create_report_from_template('Sales') → Sales report
    └── FileOrganizer.create_organization_report() → text log file
         ↓
ExcelHandler.create_summary_report() → master summary Excel at Level 1
         ↓
os.startfile(level1_folder) → opens output in file explorer
```

### Power Query Extractor Flow

```
User launches launch_extractor.py
         ↓
Loads target_folder from cache (BUG: wrong path - see F10)
         ↓
Scans for "Annual Statement-*" folders → finds latest
         ↓
For each client subfolder:
    └── Finds latest "Version-*" subfolder
    └── Checks for ITC_Report_*.xlsx and Sales_Report_*.xlsx
         ↓
Displays client list with ITC/Sales checkboxes
         ↓
    [START PROCESSING]
         ↓
For each selected client:
    ├── Copy report → add "_Refreshed_{timestamp}" suffix
    ├── Open in Excel COM (visible)
    ├── SendKeys: Alt+A, R, A (Refresh All)
    ├── Wait configurable seconds
    ├── Check for error dialogs (enumerate windows)
    ├── Validate: check Info sheet cell BB2
    ├── Save and close
    ├── Open NEW Excel instance (hidden)
    ├── Extract cells per CELL_MAPPINGS config
    └── Close
         ↓
DataConsolidator.create_report() → PQ_Extraction_Report_*.xlsx at Level 1
         ↓
Offer to open report
```

### Settings/Cache Data Flow

```
Main App (gui/utils/cache_manager.py)
    WRITES TO: ~/.gst_organizer_cache.json
    Keys: source_folder, itc_template, sales_template, target_folder,
          processing_mode, include_client_name, dark_mode,
          client_name_max_length, recent_folders[]

PQ Extractor (power_query_extractor/)
    READS FROM: ./gst_organizer_cache.json  ← MISMATCH! Different path!
    Keys read: target_folder
```

---

## SECTION 5: BUSINESS RULES AND DOMAIN LOGIC

### File Naming Patterns

| Pattern Key | Regex | File Type | Category |
|-------------|-------|-----------|----------|
| GSTR-2B-Reco | `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | GSTR-2B Reco | ITC |
| ImsReco | `^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$` | IMS Reco | ITC |
| GSTR3B | `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | GSTR-3B Export | GSTR3B |
| Sales | `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | Sales | Sales |
| SalesReco | `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | Sales Reco | Sales |
| AnnualReport | `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$` | Annual Report | Annual |

Key rules:
- All patterns are case-insensitive (`re.IGNORECASE`) — `constants.py:15-61`
- GSTR3B pattern handles Windows duplicate prefix `(1) GSTR3B-...` — `constants.py:32`
- ImsReco requires exactly 8-digit date format — `constants.py:24`
- Sales requires exactly 4 groups (client, state, start_month, end_month) — `constants.py:39-44`
- The `[^-]+` group means hyphens WITHIN client names will break parsing. A client named "ABC-XYZ Ltd" would be mis-parsed.

### Expected File Types per Client

A "complete" client must have ALL of these file types (`constants.py:68-75`):
1. GSTR-2B Reco
2. IMS Reco
3. GSTR-3B Export
4. Sales
5. Sales Reco
6. Annual Report

**Special rule:** GSTR-3B Export allows unlimited duplicates (monthly filings) — `file_parser.py:267-268`. All other types flag duplicates as warnings.

### State Code Mapping

Full Indian state/UT name → 2-letter code mapping at `constants.py:234-274`. Includes alternative names:
- "orissa" → "OD" (old name for Odisha)
- "pondicherry" → "PY" (old name for Puducherry)
- "national capital territory of delhi" → "DL"
- "dadra and nagar haveli and daman and diu" → "DD" (merged UT)

Fallback for unknown states (`helpers.py:502-510`): takes first letters of words (2-word → first letter of each, 1-word → first 3 chars).

### Filename Sanitization Rules

Applied at `helpers.py:66-112`:
1. "Private" → "Pvt", "Limited" → "Ltd" (case-sensitive!)
2. Invalid chars `<>:"/\|?*[]{}+=!@#$%^,;'"\`~` replaced with `_`
3. Control characters (ASCII 0-31, 127-159) replaced with `_`
4. Hyphens `-` are PRESERVED (deliberately excluded from invalid chars)
5. Multiple spaces/underscores collapsed to single `_`
6. Leading/trailing `_`, `.`, spaces stripped
7. Max length: 200 chars (default), leaving room for extension

### Client Name Length Limit

`create_client_state_key()` at `helpers.py:513-537`:
- Default max length: 35 characters (configurable via GUI spinbox, range 15-100)
- If key exceeds max, client name is truncated to fit while preserving state code and hyphen
- Minimum client name length after truncation: 5 characters

### Folder Structure Template

Defined at `constants.py:81-90`:
```
Level 1: Annual Statement-{timestamp}         (DDMMYY HHMM format)
Level 2: {client}-{state}                     (using state CODE, not full name)
Level 3: Version-{timestamp}
Level 4 categories:
  - GSTR-3B Exports ({client})               (client name in parens if enabled)
  - Other ITC related files ({client})
  - Sales related files ({client})
```

### File Type → Destination Folder Mapping

Defined at `file_organizer.py:288-295`:
| File Type | Destination Folder Key |
|-----------|----------------------|
| GSTR-3B Export | gstr3b |
| GSTR-2B Reco | itc |
| IMS Reco | itc |
| Sales | sales |
| Sales Reco | sales |
| Annual Report | version (root of Version folder) |

### Excel Validation Rules

At `helpers.py:210-247`:
- Must exist and be a file
- Extension must be one of: `.xlsx`, `.xls`, `.xlsm`, `.xltx`, `.xltm`
- File size must be ≥ 1024 bytes (1KB)
- File header must start with `PK\x03\x04` (ZIP/XLSX) or `\xd0\xcf\x11\xe0` (OLE/XLS)

### Client Name Validation

At `helpers.py:321-342`:
- Cannot be empty/whitespace
- Minimum 2 characters
- Maximum 100 characters
- Cannot contain: `<>:"/\|?*`

### Processing Modes

At `constants.py:157-173`:
- **fresh**: Create new timestamped folder structure. If destination file exists, create backup.
- **rerun**: Find latest existing "Annual Statement-" folder, create new Version subfolder.
- **resume**: Find latest existing folder. If destination file already exists, skip it.

### Power Query Refresh Validation

At `report_processor.py:597-644`:
- Checks "Info" sheet exists in workbook
- Reads cell BB2
- Fails if: None, empty string, "0", "0.0", or any Excel error value (#REF!, #VALUE!, #NAME?, #NULL!, #DIV/0!, #N/A, #NUM!, #ERROR!)
- Success: BB2 contains a non-empty, non-zero, non-error string (expected to be client name)

### Operational Constants

| Constant | Value | Location | Purpose |
|----------|-------|----------|---------|
| MAX_FILE_SIZE_MB | 100 | constants.py:181 | Never actually enforced — dead constant |
| max_retries | 3 | constants.py:218 | Excel save retries — not used in current code |
| retry_delay | 1 sec | constants.py:219 | Delay between retries — not used |
| temp_prefix | "temp_gst_" | constants.py:220 | Temp file naming — not used |
| verification_size_min | 10240 (10KB) | constants.py:222 | Min file size for verification — not used |
| wait_time | 10 sec | report_processor.py:27 | Default Power Query refresh wait |
| Long path threshold | 218 chars | excel_handler.py:304, report_processor.py:676 | Triggers 8.3 short path fallback |
| Unique path safety limit | 100 | helpers.py:363 | Max attempts to create unique filename |
| Short path fallback | win32api.GetShortPathName | helpers.py:476-478 | Only on Windows, with graceful except |

---

## SECTION 6: EDGE CASES AND DEFENSIVE CODE

### File System Edge Cases

| Edge Case | Handling | Location | Rating |
|-----------|----------|----------|--------|
| Source folder doesn't exist | Raises FileNotFoundError | file_parser.py:154 | ROBUST |
| Source folder is a file, not directory | Raises ValueError | file_parser.py:156 | ROBUST |
| No Excel files in folder | Returns empty results, logs warning | file_parser.py:159-161 | ROBUST |
| Corrupted Excel file (bad signature) | Excluded from processing, added to variations | file_parser.py:186-194 | ROBUST |
| File smaller than 1KB | Treated as corrupted, excluded | helpers.py:232 | ROBUST |
| Destination file already exists (fresh mode) | Creates backup with timestamp suffix | file_organizer.py:260-265 | ROBUST |
| Destination file already exists (resume mode) | Skipped | file_organizer.py:266-269 | ROBUST |
| Long Windows paths (>218 chars) | Falls back to 8.3 short path via win32api | excel_handler.py:296-317 | MODERATE |
| Network drive disconnection | Not handled — will throw generic Exception | — | FRAGILE |
| File locked by another process | Not explicitly handled — relies on generic Exception catch | — | FRAGILE |

### Data Edge Cases

| Edge Case | Handling | Location | Rating |
|-----------|----------|----------|--------|
| Empty client name from parsing | Logged as warning, still processed | file_parser.py:114-116 | MODERATE |
| Unknown state name | Fallback abbreviation generated (first letters) | helpers.py:502-510 | ROBUST |
| Client name > 100 chars | validate_client_name returns error | helpers.py:334-335 | MODERATE (warning only, doesn't block) |
| Empty filename | Returns "unnamed" | helpers.py:70-71 | ROBUST |
| Unicode in filenames | Control characters replaced with underscore | helpers.py:85-88 | MODERATE |
| Hyphen in client name | Breaks regex parsing (treated as field separator) | constants.py patterns | FRAGILE |
| Duplicate client-state combinations | Files merged into same client_data entry | file_parser.py:226-246 | ROBUST |

### Error Handling Patterns

| Pattern | Used in | Rating |
|---------|---------|--------|
| Try-except with logging + re-raise | file_parser.py scan_folder | ROBUST |
| Try-except with fallback value | helpers.py validate_excel_file | ROBUST |
| Bare except (catches everything silently) | excel_handler.py:62,124,222-228,263,545 | FRAGILE |
| Thread exception → messagebox | processing_handler.py:455-458 | MODERATE |
| COM cleanup in finally block | excel_handler.py:206-228, report_processor.py:353-374 | ROBUST |

**Overall defensive coding rating: MODERATE** — Common paths are well-handled, but bare except clauses hide errors, and several Windows-specific operations have no cross-platform fallback.

---

## SECTION 7: EDGE CASES NOT HANDLED (GAPS AND VULNERABILITIES)

### Critical Gaps

1. **Cache path mismatch (F10):** Main app writes cache to `~/.gst_organizer_cache.json` but PQ Extractor reads from `./gst_organizer_cache.json`. The extractor will fail to find the main app's settings unless manually configured.

2. **Lambda variable capture in loops (F6):** Multiple locations in `processing_handler.py` use `lambda` inside for loops without capturing the loop variable by value. Lines 41-48 (dry_run), 178, 181, 189, 195-200, 206, 216, 221, 237-239, etc. All log messages from the processing thread show the LAST iteration's values, not the current one. Fix: Use `lambda info=client_info: self.app.log_message(...)`.

3. **Platform crashes in PQ Extractor:** `report_processor.py` imports `pythoncom`, `win32com.client`, `win32api`, `win32con`, `win32gui` at module level (lines 11-14) without any try/except guard. Importing this module on macOS/Linux will crash immediately with ImportError. Compare with `excel_handler.py:20-28` which does this correctly with a platform check.

4. **os.startfile() calls without platform check:** `client_handler.py:251` uses `os.startfile()` which is Windows-only. Compare with `processing_handler.py:444-449` which correctly checks `platform.system()`. The client handler will crash on non-Windows when exporting a client list.

5. **requirements.txt missing pywin32:** The file lists only `openpyxl` and `lxml`. The README correctly states `pywin32>=305` should be installed, but `requirements.txt` doesn't include it. A `pip install -r requirements.txt` on a fresh machine won't install pywin32.

6. **Thread safety for shared mutable state:** `processing_handler.py` runs in a daemon thread and accesses `self.app.client_data`, `self.app.client_folder_settings`, and other shared state without locks. While Tkinter's GIL provides some protection, dictionary mutations during iteration could theoretically cause issues if the user navigates/rescans while processing.

### Silent Logic Bugs

7. **Sanitize filename case sensitivity:** `helpers.py:78-79` replaces "Private" and "Limited" with exact case. A filename with "PRIVATE" or "limited" won't be abbreviated. Compare with `helpers.py:517` (`create_client_state_key`) which uses `re.IGNORECASE` — inconsistent.

8. **validate_client_name used for state validation:** `file_parser.py:113` calls `validate_client_name(result['state'])` to validate the state name. This works but is semantically wrong — a state name of "MH" (2 chars) passes, but a state name of "M" (1 char) would fail with "Client name too short".

### Scenario-Based Failure Analysis

9. **Different working directory:** `main.py:36` creates `logs/` and `temp/` directories relative to `__file__`, which is correct. But `extractor_main.py:39` looks for cache files relative to CWD, which varies depending on how the script is launched.

10. **Second run on same data (fresh mode):** Creates a NEW `Annual Statement-{timestamp}` folder. Previous run's output is untouched. Safe but may confuse users about which is the latest.

11. **Two instances running simultaneously:** Both would write to the same cache file. Both would try to create folders with the same timestamp. Both would try to use Excel COM — which can conflict if both open the same workbook.

12. **Excel already open:** `excel_handler.py:122-123` tries `GetActiveObject` first, reusing a running Excel instance. If the user has files open in Excel, the COM automation might interact with those files unexpectedly. The `_safe_open_workbook` method (line 285) does check if the specific file is already open and reuses that workbook, which helps but doesn't prevent all conflicts.

13. **Path with spaces in OneDrive:** The project itself lives at `d:\OneDrive 2\OneDrive\Python 2026\AnnualStatement` — a path with spaces. This works because Path objects handle spaces correctly, but the 218-character path limit check could trigger more often with long OneDrive paths.

### Missing Validations

14. **No template structure validation:** The code doesn't verify that template files actually contain a "Links" sheet before processing begins. The error only surfaces during report creation, after folders are already created and files copied.

15. **No disk space check:** No verification that sufficient disk space exists before copying potentially hundreds of files.

16. **No concurrent access protection:** No file locking on the cache file or output directories.

---

## SECTION 8: DESIGN DECISIONS (INFERRED)

### D1: Tkinter over other GUI frameworks
**Confidence: HIGH** — Tkinter is Python's built-in GUI library, requiring zero additional installation. For a non-programmer using AI assistance to build a tool, this is the path of least resistance. The PyQt5 migration attempt (branch `copilot-worktree-2026-03-04T11-52-56`) confirms this was reconsidered later but abandoned.

### D2: Win32COM for Excel operations (not openpyxl alone)
**Confidence: HIGH** — The code explicitly states why at `excel_handler.py:3`: "WINDOWS VERSION - Uses Excel COM to preserve Power Query and all Excel features". openpyxl cannot refresh Power Query connections or preserve them in saved files. Using Excel COM via win32com is the ONLY way to preserve Power Query while programmatically modifying Excel files. The openpyxl fallback (lines 230-273) exists for non-Windows but explicitly warns "Power Query may not be preserved".

### D3: Separate Power Query Extractor app
**Confidence: HIGH** — The PQ Extractor is a post-processing step. The main organizer creates files and writes paths into templates. The PQ Extractor then opens those files, refreshes Power Query (which requires Excel to actually query external data sources), and extracts the results. These are logically sequential but temporally separated — the user may want to run the extractor days later or multiple times. Having it as a separate app makes sense.

### D4: File-based cache (JSON in home directory) instead of database
**Confidence: HIGH** — Simple, portable, no dependencies. For a single-user desktop app saving 8 settings, a database would be over-engineering. The JSON cache at `~/.gst_organizer_cache.json` is appropriate.

### D5: Regex-based filename parsing instead of folder-based organization
**Confidence: MEDIUM** — The tool requires specific filename patterns. This is likely because the consultant's existing workflow already produces files with these names (from the GST portal's export functionality or from the consultant's own naming convention). The tool adapts to existing naming rather than requiring a new folder structure as input.

### D6: State code abbreviation for folder names
**Confidence: HIGH** — Using "MH" instead of "Maharashtra" keeps folder names shorter, which matters because Windows has a 260-character path limit. The code at `file_organizer.py:87` explicitly comments "Use state code instead of full name". The `get_short_path()` fallback and 218-character threshold check further confirm path length was a real problem encountered in production.

### D7: Processing in a daemon thread
**Confidence: HIGH** — Tkinter's main loop blocks on `mainloop()`. Processing multiple clients with Excel COM operations can take minutes. Running in a daemon thread with `root.after(0, ...)` for UI updates is the standard Tkinter approach for long-running operations.

### D8: Modular handler architecture
**Confidence: MEDIUM** — The GUI is split into tabs, handlers, widgets, and utils. This is a common AI-generated code pattern where separation of concerns is applied somewhat mechanically. The handlers (file_handler, client_handler, processing_handler) each take `app_instance` as constructor argument and access shared state through it. This avoids circular imports but creates tight coupling.

### D9: Client name in folder brackets is optional
**Confidence: HIGH** — The `include_client_name_in_folders` checkbox and per-client FolderName toggle were added to handle the trade-off between descriptive folder names (useful when browsing) and path length (shorter is safer for Windows).

### Pattern Consistency Check

| Pattern | Established By | Consistent? |
|---------|---------------|-------------|
| Platform check for win32 imports | `excel_handler.py:20-28` (try/except with platform check) | **INCONSISTENT**: `report_processor.py:11-14` imports without any guard |
| os.startfile cross-platform | `processing_handler.py:444-449` (platform.system() check) | **INCONSISTENT**: `client_handler.py:251` uses os.startfile without check |
| Thread-safe UI updates | `processing_handler.py` (root.after for all UI calls) | **CONSISTENT** across processing_handler |
| Logging pattern | All modules use `logger = logging.getLogger(__name__)` | **CONSISTENT** |
| Error handling in COM operations | `excel_handler.py` uses try/finally with cleanup | **CONSISTENT** with `report_processor.py` |
| Cache file path | Main app: `Path.home() / '.gst_organizer_cache.json'` | **INCONSISTENT**: PQ Extractor: `Path("gst_organizer_cache.json")` (CWD) |
| Case sensitivity in filename abbreviations | `helpers.py:78-79` case-sensitive | **INCONSISTENT**: `helpers.py:517` uses re.IGNORECASE |

---

## SECTION 9: DEPENDENCIES AND ENVIRONMENT

### External Libraries

| Library | Version | Purpose | Required? |
|---------|---------|---------|-----------|
| openpyxl | ≥3.1.2 | Excel file creation/reading (non-Power Query) | YES |
| lxml | ≥4.9.0 | Faster XML parsing for openpyxl | Recommended |
| pywin32 | Not specified | Excel COM automation for Power Query | YES on Windows (MISSING from requirements.txt) |

### Python Version

- Minimum: 3.7 (stated in requirements.txt comment and README)
- Uses: f-strings (3.6+), `Path` (3.4+), `typing` (3.5+), `defaultdict` (2.5+), `dataclass` not used
- No 3.10+ features (no match/case, no `|` union types)

### OS-Specific Dependencies

| Module | Used In | Platform | Fallback |
|--------|---------|----------|----------|
| `win32com.client` | excel_handler.py, report_processor.py | Windows only | openpyxl (excel_handler.py), NONE (report_processor.py) |
| `pythoncom` | excel_handler.py, report_processor.py | Windows only | Same as above |
| `win32api` | helpers.py, report_processor.py | Windows only | Returns original path (helpers.py), CRASHES (report_processor.py) |
| `win32con` | report_processor.py | Windows only | CRASHES on import |
| `win32gui` | report_processor.py | Windows only | CRASHES on import |
| `os.startfile()` | processing_handler.py, client_handler.py, extractor_window.py | Windows only | platform check (processing_handler only) |
| `tkinter` | All GUI modules | All platforms | Built-in with Python |

### File System Assumptions

- Windows path separators used for Excel paths (`clean_windows_path()`)
- 260-character Windows path limit considered (218 threshold for Excel COM)
- OneDrive paths expected (based on project location)
- Home directory writable (for cache file)
- Temp directory writable (for Excel temp files: `%TEMP%/gst_excel_temp/`)

### Setup Steps (Fresh Machine)

1. Install Python 3.7+ with "Add to PATH" checked
2. `pip install openpyxl>=3.1.2`
3. `pip install lxml>=4.9.0`
4. `pip install pywin32>=305` (Windows only)
5. Run `python Scripts/pywin32_postinstall.py -install` as admin (Windows, for COM registration)
6. Download project files
7. `python main.py` to launch organizer
8. `python launch_extractor.py` to launch PQ extractor

---

## SECTION 10: UI/INTERFACE DOCUMENTATION

### Main Application Window (GSTOrganizerApp)

**Window size:** 1200×800, minimum 1000×700
**Theme:** Clam ttk theme, custom colors (primary blue #0078D4, success green #107C10)
**Layout:** Title bar at top, tabbed notebook in center, status bar at bottom

#### Title Bar
- Blue (#0078D4) background, 70px height
- Title: "🏢 GST FILE ORGANIZER & REPORT GENERATOR" in 18pt bold white
- Subtitle: "Organize files and generate Excel reports automatically" in 11pt
- Dark Mode checkbox (🌙 Dark Mode) in top-right corner

#### Tab 1: Setup (📁 Step 1: Setup)
Two-column layout (60/40 split):

**Left Column:**
1. Welcome banner (white card with blue text)
2. **Source Folder** section (green header): Entry field + Browse button
3. **Excel Templates** section (blue header): ITC template entry + browse, Sales template entry + browse
4. **Target Folder** section (red header, emphasized): Entry field + Browse button with "🚨 IMPORTANT!" warning
5. **Processing Mode** section (teal header): 3 radio buttons (Fresh/Re-Run/Resume), "Include client name" checkbox, client name max length spinbox (15-100, default 35)
6. **Action Buttons** section: "🔍 SCAN FILES" (green) + "🔄 RE-SCAN" (orange)

**Right Column:**
1. Instructions card (blue): 5-step numbered guide
2. File patterns card (purple): 6 expected file name formats

#### Tab 2: Validation (✅ Step 2: Validation)
Two-column layout (40/60 split):

**Left Side:**
1. Collapsible instructions (orange)
2. Scan Summary (scrolled text widget, Consolas 9pt)
3. Action buttons: "🧪 DRY RUN" (orange) + "📋 EXPORT LIST" (teal) + "🚀 START" (green)

**Right Side:**
- Client Selection panel with:
  - Keyboard hint: "⌨️ Use ↑↓ arrows + SPACE to select"
  - Buttons: "☑️ Select All", "☐ Clear All", "✅ Complete Only"
  - Treeview with columns: ✓, Client Name, State, Status, Files, Missing, Extra, FolderName
  - Color coding: green (#E8F5E8) for complete, orange (#FFF4E5) for incomplete

#### Tab 3: Processing (🚀 Step 3: Processing)
Vertical layout:
1. Progress section: Progress bar + percentage label + current operation label
2. Control buttons: "🚀 START PROCESSING" (green) + "⏹ STOP" (red, initially disabled)
3. Log section: Dark-themed text widget (Consolas 9pt, #1E1E1E background), color-coded tags (success=green, warning=orange, error=red, info=blue)

#### Status Bar
- Dark (#323130) background, 30px height
- Left: Status message with 💡 prefix
- Right: "v3.0 | Production Ready"

### Power Query Extractor Window (PowerQueryExtractorApp)

**Window size:** 1000×700, minimum 900×600
**Theme:** Clam ttk theme

#### Title Bar
- Blue (#0078D4) background, 60px height
- "🔄 Power Query Report Extractor" in 18pt bold white

#### Tab 1: Setup
Two-column layout:

**Left Column (400px fixed):**
1. Target Folder: entry + Browse + Scan buttons
2. Processing Options:
   - Wait time spinbox (5-60 seconds, default 10)
   - Suffix pattern entry (default: `_Refreshed_{timestamp}`)
   - "Skip Refresh" checkbox
3. "🚀 Start Processing" button (full width, initially disabled)

**Right Column:**
- Client Selection with header labels: Client Name | ITC | Sales | Last Refresh Status
- Buttons: ✓ Select All, ✗ Deselect All, ITC Only, Sales Only
- Scrollable checkbox list per client with ITC/Sales individual toggles

#### Tab 2: Processing
- Progress label + progress bar
- Log text widget (Consolas 9pt, light background)
- Color tags: info=black, success=green, warning=yellow, error=red

#### Status Bar
- Gray (#e0e0e0) background, 25px height

---

## SECTION 11: RECONSTRUCTION BLUEPRINT

### 1. Foundation

**Environment setup:**
- Python 3.8+ (for better typing support)
- `pip install openpyxl>=3.1.2 pywin32>=305` (on Windows)
- Project structure: `core/`, `utils/`, `gui/`, `power_query_extractor/`

**First file to create:** `utils/constants.py` — all patterns, mappings, and config in one place.
**Second file:** `utils/helpers.py` — all utility functions.
**Third file:** `core/file_parser.py` — the scanning/parsing engine.

### 2. Core Logic Build Order

```
1. utils/constants.py          (no dependencies)
2. utils/helpers.py            (depends on constants)
3. core/file_parser.py         (depends on constants, helpers)
4. core/file_organizer.py      (depends on constants, helpers)
5. core/excel_handler.py       (depends on constants, helpers)
6. gui/ modules                (depends on all core modules)
7. power_query_extractor/      (depends on openpyxl, win32com)
```

### 3. Feature Priority

| Priority | Feature | Complexity | Notes |
|----------|---------|------------|-------|
| P0 | File scanning & pattern matching (F1) | MODERATE | Core value proposition |
| P0 | Folder structure creation (F2) | SIMPLE | Straightforward Path operations |
| P0 | File organization/copy (F3) | SIMPLE | shutil.copy2 with verification |
| P0 | ITC/Sales report generation (F4) | COMPLEX | Win32COM is tricky |
| P1 | Summary report (F5) | MODERATE | openpyxl formatting |
| P1 | Client selection UI (F7) | MODERATE | Treeview with keyboard support |
| P1 | Full processing pipeline (F8) | COMPLEX | Threading, progress, error handling |
| P2 | Settings persistence (F10) | SIMPLE | JSON read/write |
| P2 | Dark mode (F9) | MODERATE | Recursive widget theming |
| P2 | Dry run preview (F6) | SIMPLE | Log-only version of F8 |
| P3 | PQ Refresh & Extract (F11) | COMPLEX | COM automation with SendKeys |
| P3 | Consolidation report (F12) | MODERATE | Multi-sheet openpyxl |
| P3 | Export client list (F13) | SIMPLE | Text file write |

### 4. Known Improvements for Rebuild

> Following Section 8 constraints — not recommending replacing Tkinter (D1 justifies it) or Win32COM (D2 requires it).

1. **Fix cache path mismatch** (Section 7, item 1): Centralize cache path in a shared constant. Both apps should use `Path.home() / '.gst_organizer_cache.json'`.

2. **Fix lambda capture bug** (Section 7, item 2): In all `root.after(0, lambda: ...)` calls inside loops, use default argument binding: `lambda ci=client_info: self.app.log_message(ci['client'])`.

3. **Guard PQ Extractor imports** (Section 7, item 3): Add try/except around win32 imports in `report_processor.py`, matching the pattern in `excel_handler.py:20-28`.

4. **Fix os.startfile calls** (Section 7, item 4): Add platform check to `client_handler.py:251`.

5. **Complete requirements.txt** (Section 7, item 5): Add `pywin32>=305; sys_platform == 'win32'`.

6. **Case-insensitive abbreviations** (Section 7, item 7): Use `re.sub(r'\bPrivate\b', 'Pvt', ..., flags=re.IGNORECASE)` in `sanitize_filename()`.

7. **Validate templates before processing** (Section 7, item 14): Check for "Links" sheet in templates during `validate_processing_inputs()`.

8. **Remove dead code**: Delete `calculate_file_hash()`, `_write_section()`, `_create_details_sheet()`, unused constants (`TEMPLATE_EXTENSIONS`, `EXCEL_SIGNATURES`, `MAX_FILE_SIZE_MB`), stale `# ADD THIS` comments.

9. **Clean up excessive DEBUG logging**: `excel_handler.py` and `processing_handler.py` contain many `[DEBUG]` log statements that were clearly added during troubleshooting and never removed. These clutter the log file.

10. **Add template validation at scan time**: Before switching to the Validation tab, verify that selected templates are valid and contain a "Links" sheet. This catches issues early.

### 5. Estimated Complexity

| Component | Complexity | Reason |
|-----------|------------|--------|
| Constants/Config | SIMPLE | Pure data definitions |
| Helpers/Utilities | SIMPLE | Standard Python operations |
| File Parser | MODERATE | Regex patterns, data structure design |
| File Organizer | MODERATE | Path logic, processing modes |
| Excel Handler (openpyxl) | MODERATE | Cell formatting, multiple sheets |
| Excel Handler (COM) | COMPLEX | Win32COM is brittle, error handling critical |
| GUI Layout (Tkinter) | MODERATE | Many widgets, scrollable areas |
| Threading/Progress | COMPLEX | Thread safety, progress reporting |
| Dark Mode | MODERATE | Recursive widget traversal |
| PQ Refresh | COMPLEX | SendKeys, window enumeration, timing |
| Data Consolidation | MODERATE | Multi-sheet report from config |
| Cache/Settings | SIMPLE | JSON serialize/deserialize |

### 6. Suggested Claude Code Skills

| Feature Area | Skill Description |
|--------------|-------------------|
| Excel COM Automation | Skill for opening, modifying, and saving Excel files via win32com while preserving Power Query, with proper COM initialization/cleanup pattern |
| Tkinter Threaded Processing | Skill for running long operations in background thread with thread-safe UI updates via root.after |
| Regex-based File Classification | Skill for scanning directories and classifying files by regex patterns into grouped categories |
| Excel Report Generation | Skill for creating formatted multi-sheet Excel reports with headers, borders, auto-fit columns using openpyxl |

---

## SECTION 12: CHANGELOG ARCHAEOLOGY (FROM GIT HISTORY)

### Branch Map

```
main (production):    edac6a3 → 6237597 → 9c56a1c → d674386 → 5ea77c2 → 1cb878b → 01862a0
                                    ↑           ↓
                      (GitHub)  cbfa4a5 ────────┘

                      → 1c48722 → 4c42758 → bd176bf → 47a1492 → fc279af (HEAD)

copilot-worktree (unmerged):  47a1492 → c258047 → a62d3cc → c21f788
```

### 12A. Complete Development Timeline

#### Phase 1: Initial Release (edac6a3 — 2025-10-29)

**Commit edac6a3** — "Initial commit"
- **Branch:** main
- **What changed:** Entire codebase added — 110 files, 12,916 insertions. Includes all core modules, GUI, utilities, power_query_extractor, and accidentally commits `__pycache__/` directories and `gst_organizer.log` (5,689 lines of debug output).
- **Why:** First commit of a working application. The included log file reveals extensive prior development/testing outside version control.
- **What it reveals:** The project was developed substantially before being put into git. The log file shows real-world processing of GST files with actual client names and file operations. The codebase arrived in a relatively mature state — modular architecture, comprehensive error handling, configurable processing modes — suggesting significant AI-assisted iteration before the first commit.

#### Phase 2: Cleanup (cbfa4a5, 6237597, 9c56a1c — 2025-10-29)

**Commit cbfa4a5** — "Delete gst_organizer.log"
- **Branch:** main (remote)
- **What changed:** Removed 5,689-line log file accidentally committed.
- **Why:** The log contained debug output and potentially sensitive client information.

**Commit 6237597** — "Fix critical import and threading bugs in GST Organizer"
- **Branch:** main (local)
- **What changed:** 4 files, 59 insertions, 45 deletions:
  1. `core/file_parser.py`: Added `self.progress_callback = None` initialization and `_update_progress()` safe wrapper
  2. `gui/handlers/file_handler.py`: Fixed `self.app.file_parser.scanned_files` to `self.app.scanned_files` (the scanned data is stored on the app instance, not the parser)
  3. `gui/handlers/processing_handler.py`: Wrapped ALL UI update calls in `self.app.root.after(0, lambda: ...)` for thread safety; improved long client name validation with configurable max_len
  4. `utils/helpers.py`: Fixed `from utils.constants import STATE_CODE_MAPPING` to `from .constants import STATE_CODE_MAPPING` (relative import fix)
- **Why:** These are bugs discovered during first real-world use:
  - The circular import error (helpers → utils.constants) would crash on startup
  - The thread safety issue would cause random Tkinter crashes during processing
  - The scanned_files attribute error would crash on scan completion
- **What it reveals:** The initial commit was from a development environment where absolute imports worked (maybe running from project root). The circular import and thread safety issues are classic "it worked in testing but not in production" bugs.

**Commit 9c56a1c** — "Merge branch 'main'"
- **Branch:** main
- **What changed:** Merge commit combining local fixes (6237597) with remote log deletion (cbfa4a5).
- **Why:** The developer had both local changes and a remote change (GitHub UI deletion of the log file), requiring a merge.

#### Phase 3: Documentation and Hygiene (d674386, 5ea77c2, 1cb878b, 01862a0 — 2025-10-29)

**Commit d674386** — "Update README.md for clarity and formatting improvements"
- **Branch:** main
- **What changed:** Major README overhaul — added Excel Corruption Prevention section, table-formatted file patterns, PQ Extractor documentation, cell extraction config example, troubleshooting table. 239 insertions, 135 deletions.
- **Why:** Documentation maturation after initial deployment. The "Excel Corruption Prevention" section suggests corruption was a real problem in pre-git development.

**Commit 5ea77c2** — "Remove shebang line from main.py"
- **Branch:** main
- **What changed:** Removed `# !/usr/bin/env python3` from line 1.
- **Why:** The shebang is unnecessary on Windows and had a space (`# !` instead of `#!`), making it invalid anyway.

**Commit 1cb878b** — "Add .gitignore file"
- **Branch:** main
- **What changed:** Added 79-line .gitignore covering Python bytecode, virtual environments, IDE files, logs, Excel temps, output directories, Windows files, Power Query cache.
- **Why:** Prevents future accidental commits of log files, __pycache__, and output data.

**Commit 01862a0** — "Remove __pycache__ directories from Git tracking"
- **Branch:** main
- **What changed:** Removed 70 .pyc binary files that were committed in the initial commit.
- **Why:** Cleanup now that .gitignore prevents re-addition.

#### Phase 4: Production Bug Fixes (1c48722 — 2025-10-29)

**Commit 1c48722** — "Remove redundant text color handling in client tree display"
- **Branch:** main
- **What changed:** Removed 6 lines from `client_handler.py` that redundantly re-configured tree tag colors and force-iterated columns.
- **Why:** The dark mode manager already handles tree tag coloring. This redundant code was likely causing visual glitches (colors not matching dark mode state).

#### Phase 5: PQ Extractor Enhancement (4c42758, bd176bf, 47a1492 — 2025-10-30)

**Commit 4c42758** — "Add mouse wheel scrolling support for main and log text areas"
- **Branch:** main
- **What changed:** Added `_on_mousewheel` handler for canvas (with Enter/Leave focus binding) and `_on_log_mousewheel` for log text in `extractor_window.py`.
- **Why:** Users couldn't scroll the client list or log with the mouse wheel. Focus-based binding prevents scroll conflicts when multiple scrollable areas exist.

**Commit bd176bf** — "Enhance report processing UI with configurable wait time and file suffix options"
- **Branch:** main
- **What changed:** 3 files, 412 insertions, 162 deletions:
  1. `.claude/settings.local.json`: New file (Claude Code tooling)
  2. `report_processor.py`: Added configurable `wait_time` and `suffix_pattern`; made `process_client()` selective (process_itc/process_sales flags)
  3. `extractor_window.py`: Major UI restructuring — two-column layout, wait time spinbox (5-60s), suffix pattern entry, skip refresh checkbox, ITC/Sales individual selection per client, ITC Only/Sales Only buttons, refresh status display
- **Why:** Users needed to:
  - Adjust Power Query wait time (some reports take longer to refresh)
  - Customize the refreshed file suffix
  - Skip refresh when re-extracting from already-refreshed files
  - Process only ITC or only Sales reports for specific clients
- **What it reveals:** This is a significant feature enhancement driven by real-world usage. The default 10-second wait was probably too short for some reports, and re-running the extractor after fixing one client shouldn't require re-refreshing all.

**Commit 47a1492** — "Add new ITC report mappings"
- **Branch:** main
- **What changed:** 2 files, 126 insertions, 63 deletions:
  - Deleted `cell_mappings copy.py` (duplicate file)
  - Added 6 new extraction configs: Reco Diff, Table 6J Match, Table 8D Limits, IMS Reco, Extra In Purchase, Extra In GSTR2B
- **Why:** The consultant needed to extract more data from ITC reports. Each new config maps specific cells from specific sheets (e.g., "Reco" sheet cell C37 → "IGST Diff").
- **What it reveals:** The extraction configuration grew organically as the consultant discovered which values they needed. The deletion of the "copy" file suggests the file was duplicated during development for backup.

#### Phase 6: Project Audit (fc279af — 2026-03-04)

**Commit fc279af** — "Project Audit"
- **Branch:** main
- **What changed:** Added `PROJECT_AUDIT_ANNUALSTATEMENT.md` — 856-line comprehensive audit document.
- **Why:** The project owner (a non-coder) needed documentation of the codebase for maintenance and future development.

#### Unmerged Branch: PyQt5 Migration (c258047, a62d3cc, c21f788 — 2026-03-04)

**Commit c258047** — "Refactor GST File Organizer v3.0 -> v4.0: PyQt5 migration, unified app, enhanced caching"
- **Branch:** copilot-worktree-2026-03-04T11-52-56 (NOT on main)
- **What changed:** 33 files, 3,070 insertions, 541 deletions. Massive refactoring:
  - New `models/` directory (app_state.py, cache_manager.py, run_history.py, signals.py)
  - New `gui/pyqt_main_window.py`
  - New `gui/views/` (setup_view.py, validation_view.py, processing_view.py, extraction_view.py)
  - New `gui/controllers/` (scan_controller.py, processing_controller.py, extraction_controller.py)
  - New `gui/theme/` (theme_manager.py)
  - New `gui/widgets/collapsible_group.py`
  - Modified core files (cleanup of debug logging, bare excepts)
  - Updated main.py and launch_extractor.py for PyQt5
  - Added PyQt5 to requirements.txt
- **Why:** Attempted modernization from Tkinter to PyQt5 for better widget quality, layout system, and signal/slot architecture.

**Commit a62d3cc** — "Remove old Tkinter GUI files replaced by PyQt5 in v4.0"
- **Branch:** copilot-worktree (NOT on main)
- **What changed:** Deleted 17 Tkinter GUI files (main_window.py, all handlers, tabs, utils, widgets, extractor_window.py).
- **Why:** Cleanup after PyQt5 replacement.

**Commit c21f788** — "Fix gui/__init__.py to import PyQt5 main window instead of deleted Tkinter module"
- **Branch:** copilot-worktree (NOT on main)
- **What changed:** Updated `gui/__init__.py` to import `GSTOrganizerWindow` from `.pyqt_main_window` instead of `GSTOrganizerApp` from `.main_window`.
- **Why:** The old import was broken after Tkinter files were deleted.

> **Important:** This branch was NEVER merged to main. The current production code on `main` remains Tkinter-based. The PyQt5 code exists only on the unmerged branch.

### 12B. File Evolution Map

| File | Created | Last Modified | Evolution |
|------|---------|---------------|-----------|
| `main.py` | edac6a3 | 5ea77c2 | Shebang removed. Otherwise stable. |
| `launch_extractor.py` | edac6a3 | — | Never modified after initial commit. |
| `requirements.txt` | edac6a3 | — | Never modified on main. (Modified on unmerged branch to add PyQt5.) |
| `core/file_parser.py` | edac6a3 | 6237597 | Added progress callback safety wrapper. |
| `core/file_organizer.py` | edac6a3 | — | Never modified after initial commit. |
| `core/excel_handler.py` | edac6a3 | — | Never modified after initial commit. |
| `utils/constants.py` | edac6a3 | — | Never modified after initial commit. |
| `utils/helpers.py` | edac6a3 | 6237597 | Fixed relative import for STATE_CODE_MAPPING. |
| `gui/main_window.py` | edac6a3 | — | Never modified on main. |
| `gui/handlers/file_handler.py` | edac6a3 | 6237597 | Fixed scanned_files attribute access. |
| `gui/handlers/client_handler.py` | edac6a3 | 1c48722 | Removed redundant dark mode code. |
| `gui/handlers/processing_handler.py` | edac6a3 | 6237597 | Added thread-safe UI updates, configurable max length. |
| `gui/tabs/*.py` | edac6a3 | — | Never modified. |
| `gui/utils/*.py` | edac6a3 | — | Never modified. |
| `gui/widgets/*.py` | edac6a3 | — | Never modified. |
| `pq_extractor/gui/extractor_window.py` | edac6a3 | bd176bf | Major UI restructuring for ITC/Sales selection. |
| `pq_extractor/core/report_processor.py` | edac6a3 | bd176bf | Added configurable wait time, suffix, selective processing. |
| `pq_extractor/config/cell_mappings.py` | edac6a3 | 47a1492 | Added 6 new ITC extraction configs. |
| `pq_extractor/core/data_consolidator.py` | edac6a3 | — | Never modified. |

### 12C. Logic and Rule Evolution

| Rule | First Appeared | Changed? | Notes |
|------|---------------|----------|-------|
| 6 file patterns (regex) | edac6a3 | No | Stable since creation |
| State code mapping | edac6a3 | No | Complete from initial commit |
| GSTR-3B unlimited duplicates | edac6a3 | No | Always allowed |
| Client name max length (35) | edac6a3 | 6237597 | Made configurable via GUI |
| Thread-safe UI updates | 6237597 | No | Added to fix crash bugs |
| ITC/Sales selective processing | bd176bf | No | Added for PQ Extractor |
| Configurable PQ wait time | bd176bf | No | Added based on user need |
| Skip refresh option | bd176bf | No | Added for re-extraction |
| ITC cell mappings (GSTR-3B Diff) | edac6a3 | 47a1492 | 6 new configs added |
| Sales cell mappings (GSTR-1 Diff) | edac6a3 | No | Stable |

### 12D. Edge Cases Discovered in Production

| Commit | Bug | Fix | Unhandled Siblings |
|--------|-----|-----|-------------------|
| 6237597 | Circular import: `from utils.constants` fails when running as package | Changed to relative import `from .constants` | None — all other imports already use relative paths |
| 6237597 | Thread crash: Tkinter widget access from background thread | Wrapped in `root.after(0, lambda: ...)` | Lambda capture bug introduced — log shows wrong values |
| 6237597 | AttributeError: `file_parser.scanned_files` | Changed to `self.app.scanned_files` | None |
| 1c48722 | Dark mode tree colors incorrect after toggle | Removed redundant color code conflicting with DarkModeManager | None |
| 4c42758 | Mouse wheel not working in PQ Extractor lists | Added mousewheel event handlers with focus binding | Setup tab canvas in main app also binds `bind_all` which may conflict |
| bd176bf | PQ refresh too short for complex reports | Made wait time configurable (5-60 seconds) | Extended wait at line 285 still hardcoded as `max(10, self.wait_time)` |

### 12E. Abandoned Approaches and Reversals

1. **PyQt5 Migration (branch `copilot-worktree-2026-03-04T11-52-56`):** 3 commits adding 3,070 lines of PyQt5 code and deleting 3,298 lines of Tkinter code. Never merged to main. The branch represents a complete architectural overhaul that was abandoned. Reason likely: the migration was complex, and the existing Tkinter code works.

2. **cell_mappings copy.py:** A duplicate of the cell mappings config file was committed in the initial commit and deleted in 47a1492. This suggests the developer made a backup copy of the config before editing it — a manual versioning approach before git was used.

3. **Dead methods in data_consolidator.py:** `_write_section()` (line 250) and `_create_details_sheet()` (line 354) are remnants of an earlier report format. The `create_report()` method was rewritten to use `EXTRACTION_CONFIG` directly, making these methods dead. They should be removed.

4. **EXCEL_SAFETY constants (constants.py:217-228):** Defined `max_retries`, `retry_delay`, `temp_prefix`, `backup_prefix`, `verification_size_min`, and `save_options`. Only `save_options` is referenced (by `excel_handler.py:54` which stores it but never uses it for actual save options). The rest are dead. These suggest an earlier design for more robust file operations that was simplified or never implemented.

### 12F. Dependency Evolution

| Commit | Change | Why |
|--------|--------|-----|
| edac6a3 | requirements.txt created with openpyxl ≥3.1.2, lxml ≥4.9.0 | Initial setup |
| — | pywin32 NOT added to requirements.txt | Oversight. Used extensively but never declared. |
| c258047 (unmerged) | PyQt5 added to requirements.txt | For PyQt5 migration. Never reached main. |

**Dependencies used but missing from requirements.txt:**
- `pywin32` — Used via `import win32com.client`, `import pythoncom`, `import win32api`, `import win32con`, `import win32gui`. Critical for Power Query operations. Listed in README but not in requirements.txt.

---

## Self-Consistency Review Notes

1. **Section 3 (F6) vs Section 7 (item 2):** Both correctly identify the lambda capture bug. Consistent.

2. **Section 8 (D1, Tkinter justified) vs Section 11 (Reconstruction Blueprint):** Blueprint does NOT recommend replacing Tkinter, respecting D1. Consistent.

3. **Section 8 (D2, Win32COM required) vs Section 11:** Blueprint does NOT recommend replacing Win32COM. Consistent.

4. **Section 2 (dead code analysis) vs Section 12E (abandoned approaches):** Both identify `_write_section()` and `_create_details_sheet()` as dead. Consistent.

5. **Section 10 (cache path) vs Section 7 (item 1):** Both identify the cache path mismatch. Consistent.

6. **Section 12B ("never modified" claims) verified:** All "never modified" files checked against current code — no imports reference functions added in later commits. Verified consistent.

7. [Corrected during self-review] Section 2 initially listed `extract_filename_without_extension()` as potentially dead. Verified it IS called from `excel_handler.py` lines 576, 588, 598, 614, 631, 646. Removed from dead code list.

8. [Corrected during self-review] Section 2 initially listed `parsing_rate` as dead. Verified it IS displayed in `main_window.py:280`. Removed from dead code list.
