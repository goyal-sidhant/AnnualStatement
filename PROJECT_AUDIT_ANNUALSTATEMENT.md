# PROJECT AUDIT: GST File Organizer & Report Generator

**Audit Date:** 2026-03-04
**Auditor:** Claude Opus 4.6 (AI-assisted reverse engineering)
**Project Root:** `d:\OneDrive 2\OneDrive\Python 2026\AnnualStatement`
**Current Version:** v3.0 (Production Ready) — Tkinter GUI
**Active Branch:** `main` at commit `47a1492`

---

## SECTION 1: EXECUTIVE SUMMARY

This tool is a desktop application built for **CA/GST consultants** (Chartered Accountants working with India's Goods and Services Tax system). It solves a specific, time-consuming problem: when a consultant manages annual GST filings for dozens of clients, they receive multiple Excel files per client (GSTR-2B reconciliation, IMS reconciliation, GSTR-3B exports, Sales data, Sales reconciliation, Annual Reports) that must be organized into a standardized folder structure, linked to Excel report templates, and optionally processed through Power Query to extract key differences and totals across all clients. Without this tool, a consultant would manually create folders, copy files, open each Excel template, type in file paths, refresh Power Query connections, and copy out cell values — for every single client. This tool automates that entire workflow.

The core technology stack is: **Python 3.7+** for the language, **Tkinter** (Python's built-in GUI library) for the desktop interface, **openpyxl** for creating and reading Excel files, and **win32com/pywin32** for automating Microsoft Excel on Windows (required for Power Query refresh). The application has two entry points: `main.py` launches the main organizer GUI, and `launch_extractor.py` launches a separate Power Query Extractor tool.

The project appears **actively maintained and production-used**. The git history shows real-world bug fixes (circular imports, threading issues, long path handling) that suggest it has been used with actual client data. A PyQt5 migration was recently attempted (2026-02-07) on a separate branch but has **not been merged to main** — the production code remains Tkinter-based. The cell mappings configuration was last expanded on 2025-10-30, adding 6 new ITC report extraction sheets, suggesting the tool is growing with the consultant's needs.

---

## SECTION 2: PROJECT STRUCTURE MAP

```
AnnualStatement/                         [ROOT - Entry point directory]
├── main.py                              [CORE] Entry point - launches main GUI app
├── launch_extractor.py                  [CORE] Separate entry point for Power Query Extractor
├── requirements.txt                     [CONFIG] Python dependencies
├── README.md                            [DOCS] Comprehensive usage documentation
├── .gitignore                           [CONFIG] Git exclusion rules
│
├── core/                                [CORE] Business logic layer
│   ├── __init__.py                      [CORE] Package init - exports FileParser, FileOrganizer, ExcelHandler
│   ├── file_parser.py                   [CORE] Scans folders, matches filenames to GST patterns, groups by client
│   ├── file_organizer.py                [CORE] Creates folder structure, copies files to organized locations
│   └── excel_handler.py                 [CORE] Creates Excel reports from templates, writes summary reports
│
├── utils/                               [UTIL] Shared utilities and constants
│   ├── __init__.py                      [UTIL] Package init - re-exports everything
│   ├── constants.py                     [CONFIG] All regex patterns, folder templates, GUI config, state codes
│   └── helpers.py                       [UTIL] File operations, path utilities, validation, progress tracking
│
├── gui/                                 [UI] Tkinter GUI layer
│   ├── __init__.py                      [UI] Package init - exports GSTOrganizerApp
│   ├── main_window.py                   [UI] Main application window - ties all components together
│   ├── handlers/                        [UI] Event handler modules
│   │   ├── __init__.py                  [UI] Package init
│   │   ├── file_handler.py              [UI] Browse dialogs, scan validation
│   │   ├── client_handler.py            [UI] Client tree selection, keyboard navigation, export
│   │   └── processing_handler.py        [UI] Processing thread, report creation, progress updates
│   ├── tabs/                            [UI] Tab panel modules
│   │   ├── __init__.py                  [UI] Package init
│   │   ├── setup_tab.py                 [UI] Step 1: folder/template selection UI
│   │   ├── validation_tab.py            [UI] Step 2: client review/selection UI
│   │   └── processing_tab.py            [UI] Step 3: progress/log display UI
│   ├── utils/                           [UI] GUI utility modules
│   │   ├── __init__.py                  [UI] Package init
│   │   ├── cache_manager.py             [UI] Save/load settings to JSON file
│   │   ├── dark_mode_manager.py         [UI] Dark/light theme switching
│   │   ├── status_bar.py               [UI] Bottom status bar widget
│   │   └── ui_helpers.py                [UI] Shared UI creation helpers
│   └── widgets/                         [UI] Custom widget modules
│       ├── __init__.py                  [UI] Package init
│       ├── collapsible_frame.py         [UI] Expandable/collapsible section widget
│       └── title_bar.py                 [UI] Application title bar with dark mode toggle
│
└── power_query_extractor/               [CORE] Separate sub-application for PQ operations
    ├── __init__.py                      [CORE] Package init
    ├── extractor_main.py                [CORE] Launcher class - loads cache, starts GUI
    ├── config/                          [CONFIG] Extraction configuration
    │   ├── __init__.py                  [CONFIG] Package init
    │   └── cell_mappings.py             [CONFIG] Master config: which cells to extract from which sheets
    ├── core/                            [CORE] Processing logic
    │   ├── __init__.py                  [CORE] Package init
    │   ├── report_processor.py          [CORE] Excel COM automation: refresh Power Query, extract cell values
    │   └── data_consolidator.py         [CORE] Creates consolidated extraction report across all clients
    └── gui/                             [UI] Extractor GUI
        ├── __init__.py                  [UI] Package init
        └── extractor_window.py          [UI] Tkinter window for PQ extractor with two tabs
```

### Entry Points
- **Primary:** `python main.py` — Launches the main GST File Organizer GUI
- **Secondary:** `python launch_extractor.py` — Launches the Power Query Extractor GUI separately

### Dead/Orphan Files
- None detected in current codebase. All files are imported and used.
- Note: A file `cell_mappings copy.py` existed in the initial commit (a backup copy of cell_mappings.py) but was removed in commit `47a1492`.

---

## SECTION 3: FEATURE INVENTORY

### F1: File Scanning and Pattern Matching
- **What it does:** Scans a user-selected folder for Excel files, matches each filename against 6 predefined regex patterns (GSTR-2B Reco, IMS Reco, GSTR-3B, Sales, Sales Reco, Annual Report), extracts client name, state, and period information from the filename, and groups files by client-state combinations.
- **Where it lives:** `core/file_parser.py` (FileParser class), `utils/constants.py` (FILE_PATTERNS dict, lines 13-62)
- **Input -> Output:** A folder path containing Excel files -> A dictionary of client data (keyed by "ClientName-StateCode"), a dict of scanned files, and a list of "variations" (files that didn't match any pattern)
- **User interaction:** User selects a source folder in Setup tab, clicks "Scan Files" button

### F2: Client Completeness Analysis
- **What it does:** After scanning, checks each client against the 6 expected file types. Marks clients as "Complete" (all 6 types present), "Missing N files" (some types absent), or "Has duplicates". Calculates a completeness percentage. Special rule: GSTR-3B can have multiple files (monthly filings) without being flagged as duplicate.
- **Where it lives:** `core/file_parser.py`, method `_analyze_client_completeness()` (lines 252-286)
- **Input -> Output:** Parsed client data -> Same data enriched with status, missing files list, completeness percentage
- **User interaction:** Automatically runs after scanning; results shown in Validation tab's client tree

### F3: File Organization (Folder Structure Creation)
- **What it does:** Creates a hierarchical folder structure and copies files into it. Structure: `Target/Annual Statement-DDMMYY HHMM/ClientName-StateCode/Version-DDMMYY HHMM/[category folders]`. Category folders: "GSTR-3B Exports", "Other ITC related files", "Sales related files". Optionally includes the client name in parentheses in category folder names.
- **Where it lives:** `core/file_organizer.py` (FileOrganizer class)
- **Input -> Output:** Client data + target folder + processing mode -> Created folder tree with files copied into correct locations
- **User interaction:** User selects clients, clicks "Start Processing" in Validation or Processing tab. Three modes: Fresh (new structure), Re-Run (add version folder to existing), Resume (skip existing files)

### F4: Excel Report Generation (ITC and Sales)
- **What it does:** Copies an Excel template file to the client's version folder, then writes file paths and filenames into specific cells on a "Links" sheet. This allows the template's Power Query connections to find the data files. On Windows, uses Excel COM automation (win32com) to preserve Power Query; on other OS, falls back to openpyxl (which may lose Power Query).
- **Where it lives:** `core/excel_handler.py` (ExcelHandler class), `utils/constants.py` (EXCEL_TEMPLATE_MAPPING, lines 96-120)
- **Input -> Output:** Template file + client's organized files + folder paths -> ITC_Report_Client_State_Timestamp.xlsx and Sales_Report_Client_State_Timestamp.xlsx
- **User interaction:** Automatic during processing. User provides template files in Setup tab.

### F5: Processing Summary Report
- **What it does:** Creates a master Excel summary with 5 sheets: Summary (statistics), Client Status (per-client status), File Mapping (source-to-destination for every file), Errors (any errors encountered), Variations (files that didn't match patterns). Uses professional formatting with headers, borders, and color coding.
- **Where it lives:** `core/excel_handler.py`, method `create_summary_report()` (lines 319-549)
- **Input -> Output:** Processing results data -> GST_Processing_Summary_Timestamp.xlsx saved in Level 1 folder
- **User interaction:** Automatic at end of processing

### F6: Organization Text Report
- **What it does:** Creates a plain text file per client documenting what folders were created, which files were organized, and statistics about the operation.
- **Where it lives:** `core/file_organizer.py`, method `create_organization_report()` (lines 300-362)
- **Input -> Output:** Client info + folders -> _Organization_Report_Client_State.txt
- **User interaction:** Automatic during processing

### F7: Power Query Refresh
- **What it does:** Opens Excel reports using COM automation, sends keyboard shortcuts (Alt+A, R, A = "Refresh All") to trigger Power Query refresh, waits a configurable number of seconds, checks for error dialogs, validates refresh by checking cell BB2 in an "Info" sheet, then saves the refreshed copy with a suffix like "_Refreshed_DDMMYY_HHMM".
- **Where it lives:** `power_query_extractor/core/report_processor.py` (ReportProcessor class)
- **Input -> Output:** Excel report file -> Refreshed copy of the file with Power Query data populated
- **User interaction:** User launches PQ Extractor, selects clients and ITC/Sales checkboxes, configures wait time, clicks "Start Processing"

### F8: Cell Value Extraction
- **What it does:** Opens refreshed Excel files, reads specific cell values from specific sheets as defined in cell_mappings.py, and stores them for consolidation. Supports 9 extraction configurations: GSTR-3B Difference, GSTR-1 Difference, Turnover GSTR-1, Reco Diff, Table 6J Match, Table 8D Limits, IMS Reco, Extra In Purchase, Extra In GSTR2B.
- **Where it lives:** `power_query_extractor/core/report_processor.py` method `_extract_data()` (lines 486-595), `power_query_extractor/config/cell_mappings.py`
- **Input -> Output:** Refreshed Excel files + cell mapping config -> Dictionary of extracted values per client
- **User interaction:** Automatic after Power Query refresh

### F9: Consolidated Extraction Report
- **What it does:** Creates a master Excel workbook with one sheet per extraction configuration (9 sheets), plus a Summary sheet. Each sheet shows client names in rows and extracted values in columns, with success/failure indicators and number formatting.
- **Where it lives:** `power_query_extractor/core/data_consolidator.py` (DataConsolidator class)
- **Input -> Output:** All client extraction results -> PQ_Extraction_Report_Timestamp.xlsx in the Annual Statement folder
- **User interaction:** Automatic at end of PQ Extractor processing

### F10: Session Persistence (Cache)
- **What it does:** Saves and restores user settings (source folder, template paths, target folder, processing mode, dark mode preference, client name options) to a JSON file at `~/.gst_organizer_cache.json`. Also maintains a list of last 5 source folders used.
- **Where it lives:** `gui/utils/cache_manager.py` (CacheManager class), `gui/main_window.py` methods `save_cache()` and `apply_cached_values()`
- **Input -> Output:** Application state -> JSON file, and vice versa on startup
- **User interaction:** Automatic on every settings change and application startup

### F11: Dark Mode
- **What it does:** Toggles between light and dark color themes by recursively walking all Tkinter widgets and changing their background/foreground colors. Stores original colors on initialization for restoration. Updates tree view tag colors for dark mode.
- **Where it lives:** `gui/utils/dark_mode_manager.py` (DarkModeManager class)
- **Input -> Output:** Boolean toggle -> All widget colors changed
- **User interaction:** Checkbox in title bar labeled "Dark Mode"

### F12: Client List Export
- **What it does:** Exports the list of selected clients to a text file with timestamps.
- **Where it lives:** `gui/handlers/client_handler.py`, method `export_client_list()` (lines 221-253)
- **Input -> Output:** Selected clients -> Text file + auto-opens in default editor
- **User interaction:** "Export List" button in Validation tab

### F13: Dry Run Preview
- **What it does:** Shows what would happen during processing without actually copying files or creating reports. Logs each client and their files to the processing log.
- **Where it lives:** `gui/handlers/processing_handler.py`, method `dry_run()` (lines 23-52)
- **Input -> Output:** Selected clients -> Log messages showing planned operations
- **User interaction:** "Dry Run" button in Validation tab

### F14: State Name to Code Conversion
- **What it does:** Converts full Indian state names (e.g., "Maharashtra") to two-letter codes (e.g., "MH") for shorter folder names. Includes 30+ state/territory mappings plus alternative names (e.g., "Orissa" -> "OD", "Pondicherry" -> "PY"). Falls back to creating abbreviations from first letters for unknown states.
- **Where it lives:** `utils/constants.py` (STATE_CODE_MAPPING, lines 234-275), `utils/helpers.py` method `get_state_code()` (lines 485-511)
- **Input -> Output:** Full state name string -> 2-3 letter code string
- **User interaction:** Automatic during folder creation. Affects folder names (e.g., "ClientName-MH" instead of "ClientName-Maharashtra")

### F15: Configurable Client Folder Names
- **What it does:** Lets users control two aspects of folder naming: (1) a global checkbox to include client names in category folder names (e.g., "GSTR-3B Exports (ClientName)" vs "GSTR-3B Exports"), and (2) a per-client toggle in the tree view's FolderName column. Also provides a configurable maximum length for client folder names (default 35, range 15-100). Warns when client names exceed 10 characters.
- **Where it lives:** `gui/tabs/setup_tab.py` (processing mode section), `gui/handlers/client_handler.py` (toggle methods), `core/file_organizer.py` (create_client_structure, lines 107-135)
- **Input -> Output:** User preferences -> Folder names with or without client name suffixes
- **User interaction:** Checkbox in Setup tab + per-client toggle in Validation tab tree + spinbox for max length

---

## SECTION 4: DATA FLOW

### Main Application Flow

```
User selects Source Folder
        ↓
[file_parser.py: scan_folder()]
        ↓
Find all .xlsx/.xls/.xlsm files in folder
        ↓
For each file: validate Excel signature (ZIP or OLE header) + check size > 1KB
        ↓
For each valid file: match filename against 6 regex patterns
        ↓
Extract client name, state, period from regex groups
        ↓
Group by "ClientName-StateCode" key → client_data dictionary
        ↓
Analyze completeness (6 expected file types per client)
        ↓
Display in GUI: summary statistics + client tree with status icons
        ↓
User selects clients → clicks "Start Processing"
        ↓
[file_organizer.py: create_client_structure()]
        ↓
Create: Target/Annual Statement-DDMMYY HHMM/ClientName-StateCode/Version-DDMMYY HHMM/
        ↓
Create category subfolders: GSTR-3B Exports, Other ITC related files, Sales related files
        ↓
[file_organizer.py: organize_files()]
        ↓
Copy each file to its category folder (with filename sanitization + size verification)
        ↓
Update client_info with sanitized filenames (critical for Excel template paths)
        ↓
[excel_handler.py: create_report_from_template()]
        ↓
Copy ITC template → write folder paths + filenames to Links sheet cells (B2-B11)
Copy Sales template → write folder paths + filenames to Links sheet cells (B2-B9)
        ↓
[excel_handler.py: create_summary_report()]
        ↓
Create master summary Excel with 5 sheets
        ↓
Open output folder in file explorer
```

### Power Query Extractor Flow

```
User launches launch_extractor.py (or standalone)
        ↓
Load target folder from cache (~/.gst_organizer_cache.json)
        ↓
Scan Annual Statement folder for client subfolders
        ↓
For each client: find latest Version folder, check for ITC_Report_*.xlsx and Sales_Report_*.xlsx
        ↓
Display client list with ITC/Sales checkboxes and last refresh timestamps
        ↓
User selects clients/reports → clicks "Start Processing"
        ↓
For each selected client+report:
    ↓
    Copy report file → add "_Refreshed_DDMMYY_HHMM" suffix
    ↓
    Open copy in Excel via COM automation (visible window)
    ↓
    Send keyboard shortcuts: Alt → A → R → A (Refresh All)
    ↓
    Wait configurable seconds (default 10) + extended wait (at least 10 more)
    ↓
    Check for Excel error dialogs → close them if found
    ↓
    Validate: read cell BB2 in "Info" sheet (should have client name)
    ↓
    Save and close workbook
    ↓
    Open refreshed file in new hidden Excel instance
    ↓
    Read configured cell values from configured sheets (cell_mappings.py)
    ↓
    Store extracted values
        ↓
[data_consolidator.py: create_report()]
        ↓
Create PQ_Extraction_Report_YYYYMMDD_HHMM.xlsx with 9 data sheets + Summary
```

---

## SECTION 5: BUSINESS RULES AND DOMAIN LOGIC

### File Naming Patterns (constants.py, lines 13-62)

| File Type | Regex Pattern | Example |
|-----------|--------------|---------|
| GSTR-2B Reco | `^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | GSTR-2B-Reco-ABCLtd-Maharashtra-Apr24.xlsx |
| IMS Reco | `^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$` | ImsReco-ABCLtd-Maharashtra-30042024.xlsx |
| GSTR-3B | `^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | GSTR3B-ABCLtd-Maharashtra-Apr24.xlsx |
| Sales | `^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$` | Sales-ABCLtd-Maharashtra-Apr-Jun.xlsx |
| Sales Reco | `^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$` | SalesReco-ABCLtd-Maharashtra-Q1.xlsx |
| Annual Report | `^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$` | AnnualReport-ABCLtd-Maharashtra-2024.xlsx |

> **Important:** The GSTR-3B pattern includes an optional `(\d+)` prefix to handle browser download numbering like "(1) GSTR3B-..." — this was a real-world edge case (constants.py, line 31).

### File Categorization Rules (constants.py, lines 13-62; file_organizer.py, lines 285-298)

| File Type | Category | Destination Folder |
|-----------|----------|-------------------|
| GSTR-3B Export | GSTR3B | `GSTR-3B Exports (ClientName)/` |
| GSTR-2B Reco | ITC | `Other ITC related files (ClientName)/` |
| IMS Reco | ITC | `Other ITC related files (ClientName)/` |
| Sales | Sales | `Sales related files (ClientName)/` |
| Sales Reco | Sales | `Sales related files (ClientName)/` |
| Annual Report | Annual | Version folder root (not in a subfolder) |

### Completeness Rules (file_parser.py, lines 252-286)
- A client is "Complete" when all 6 file types are present AND no unexpected duplicates exist
- GSTR-3B is exempt from duplicate checking (multiple monthly filings are expected)
- All other types flag as duplicates if more than 1 file exists

### Excel Template Cell Mappings (constants.py, lines 96-120)

**ITC Template "Links" sheet:**
| Cell | Value Written |
|------|--------------|
| B2 | GSTR-3B folder path (Windows backslash format) |
| B4 | Annual Report folder path |
| B5 | Annual Report filename (without extension) |
| B7 | GSTR-2B Reco folder path |
| B8 | GSTR-2B Reco filename (without extension) |
| B10 | IMS Reco folder path |
| B11 | IMS Reco filename (without extension) |

**Sales Template "Links" sheet:**
| Cell | Value Written |
|------|--------------|
| B2 | Sales file folder path |
| B3 | Sales filename (without extension) |
| B5 | Annual Report folder path |
| B6 | Annual Report filename (without extension) |
| B8 | Sales Reco folder path |
| B9 | Sales Reco filename (without extension) |

### Data Extraction Cell Mappings (cell_mappings.py, lines 33-242)

| Output Sheet | Source Sheet | Cell | Output Column | Report Type |
|-------------|-------------|------|---------------|-------------|
| GSTR-3B Difference | Diff GSTR-3B | J15 | Diff IGST | ITC |
| GSTR-3B Difference | Diff GSTR-3B | K15 | Diff CGST | ITC |
| GSTR-3B Difference | Diff GSTR-3B | L15 | Diff SGST | ITC |
| GSTR-1 Difference | Diff - GSTR-1 | N1 | Diff Taxable Value | Sales |
| GSTR-1 Difference | Diff - GSTR-1 | O1 | Diff IGST | Sales |
| GSTR-1 Difference | Diff - GSTR-1 | P1 | Diff CGST | Sales |
| GSTR-1 Difference | Diff - GSTR-1 | Q1 | Diff SGST | Sales |
| Turnover GSTR-1 | Diff - GSTR-1 | I1 | Taxable Value | Sales |
| Reco Diff | Reco | C37, D37, E37 | IGST/CGST/SGST Diff | ITC |
| Table 6J Match | T6 of R9 | K29 | Table 6J Match | ITC |
| Table 8D Limits | T8 of R9 | G6 | Table 8D Limits | ITC |
| IMS Reco | IMS_Reco | AA1, AD1, AG1 | IGST/CGST/SGST | ITC |
| Extra In Purchase | ExtraInPurchase | O1, P1, Q1 | IGST/CGST/SGST | ITC |
| Extra In GSTR2B | ExtraInGSTR2B | S1, T1, U1 | IGST/CGST/SGST | ITC |

### Filename Sanitization Rules (helpers.py, lines 66-112)
- "Private" → "Pvt", "Limited" → "Ltd" (business abbreviations applied first)
- Characters `<>:"/\|?*[]{}+=!@#$%^,;'"\`~` are replaced with underscore
- Control characters (ASCII 0-31, 127-159) replaced with underscore
- Hyphens are deliberately PRESERVED (not in the invalid characters list)
- Multiple spaces/underscores collapsed to single underscore
- Leading/trailing underscores and dots stripped
- Maximum length: 200 characters (minus extension length minus 10)

### Client-State Key Rules (helpers.py, lines 513-537)
- Format: `ClientName-StateCode` (e.g., "ABC Pvt Ltd-MH")
- "Private" → "Pvt", "Limited" → "Ltd" (case-insensitive)
- State name converted to 2-letter code via STATE_CODE_MAPPING
- Maximum key length: configurable (default 35 characters)
- If key exceeds max length, client name is truncated (minimum 5 characters preserved)

### Excel File Validation Rules (helpers.py, lines 210-247)
- File must exist and be a regular file
- Extension must be one of: .xlsx, .xls, .xlsm, .xltx, .xltm
- File size must be at least 1,024 bytes (1 KB)
- File must start with either `PK\x03\x04` (ZIP/xlsx) or `\xD0\xCF\x11\xE0` (OLE/xls)

### Power Query Refresh Validation (report_processor.py, lines 597-644)
- After refresh, checks cell BB2 in an "Info" sheet
- Fails if: cell is None, blank, "0", "0.0", or any Excel error value (#REF!, #VALUE!, #NAME?, #NULL!, #DIV/0!, #N/A, #NUM!, #ERROR!)
- Success means BB2 contains a non-empty, non-zero, non-error value (expected to be client name)

### Long Path Handling (report_processor.py, lines 646-755; excel_handler.py, lines 275-317)
- Excel has a ~218 character path limit
- If path exceeds 218 chars, tries Windows 8.3 short path names via `win32api.GetShortPathName()`
- Also tries Windows extended path format (`\\?\` prefix)
- Multiple retry attempts with different path strategies

---

## SECTION 6: EDGE CASES AND DEFENSIVE CODE

### File Operations
- **Empty files:** Files under 1KB are rejected during Excel validation (helpers.py, line 232) — treats them as corrupted
- **Missing columns/sheets:** When the "Links" sheet is not found in a template, the code tries to find any sheet with "link" in its name, and if that fails, creates a new sheet (excel_handler.py, lines 147-164)
- **Copy verification:** After every file copy, source and destination file sizes are compared; mismatched copies are deleted (helpers.py, lines 172-177)
- **Existing files in fresh mode:** Creates backups with timestamp suffix before overwriting (file_organizer.py, lines 260-265)
- **Existing files in resume mode:** Skips files that already exist (file_organizer.py, lines 266-269)
- **Unique path generation:** If a file/folder already exists, appends _1, _2, etc. up to _100 (helpers.py, lines 344-363)
- **Empty folders on error:** Cleans up empty version folders if processing fails in fresh mode (file_organizer.py, lines 364-390)

### Excel COM Automation
- **Excel already running:** Attempts to reuse existing Excel instance before creating new one (excel_handler.py, lines 121-126)
- **Workbook already open:** Checks if file is already open and reuses it (excel_handler.py, lines 285-288; report_processor.py, lines 665-668)
- **COM initialization/cleanup:** Every COM operation bracket with `pythoncom.CoInitialize()` and `CoUninitialize()` in try/finally blocks
- **Error dialogs during refresh:** Scans for visible error windows by window class and title text, attempts to click OK/Close buttons, falls back to sending WM_CLOSE message (report_processor.py, lines 376-484)
- **Excel properties retry:** Setting Excel.Visible and other properties retries up to 5 times with 2-second delays (report_processor.py, lines 181-193)
- **"Call was rejected by callee" error:** Retries workbook opening 3 times with 3-second delays (report_processor.py, lines 198-208)

### GUI Thread Safety
- Processing runs in a daemon thread (processing_handler.py, line 107-112)
- GUI updates from processing thread use `self.app.root.after(0, lambda: ...)` to schedule on main thread
- Processing can be stopped via `self.app.stop_requested` flag checked at each client iteration

### Data Handling
- **Empty client/state names:** Validation in file_parser.py flags these as errors (lines 343-356)
- **Client name length validation:** Names must be 2-100 characters, no filesystem-invalid characters (helpers.py, lines 321-342)
- **Unknown state names:** Falls back to abbreviation from first letters of words (helpers.py, lines 502-511)
- **Alternative state names:** "Orissa" maps to same code as "Odisha", "Pondicherry" maps to same as "Puducherry" (constants.py, lines 255, 274)

### Overall Rating: **MODERATE**

The code handles many common edge cases (file validation, path issues, COM errors) but has gaps in some areas (see Section 7).

---

## SECTION 7: EDGE CASES NOT HANDLED (GAPS AND VULNERABILITIES)

### Missing Validations
1. **No concurrent access protection:** If two instances run simultaneously targeting the same folder, files could be corrupted or overwritten without detection
2. **No disk space check:** The tool doesn't verify sufficient disk space before starting a potentially large copy operation
3. **Template validation is shallow:** Only checks for "Links" sheet existence, doesn't verify that the expected cells (B2, B4, etc.) exist or that Power Query connections are present
4. **No GSTIN validation:** Despite being a GST tool, there is no validation of GSTIN numbers (15-digit alphanumeric codes that follow a specific format with check digits)

### Error Conditions That Could Crash
1. **Network path disconnection:** If source or target is on a network drive that disconnects mid-operation, the error handling may not gracefully recover
2. **Excel process zombie:** If Excel crashes during COM automation, the COM object may become stale and subsequent operations will fail. The cleanup code uses bare try/except which swallows errors
3. **Memory with very large files:** No file size limit enforcement during copy (MAX_FILE_SIZE_MB = 100 is defined in constants but never checked)
4. **Regex backtracking:** The `[^-]+` patterns in file regex could be slow with unusual filenames containing many special characters

### Input Variations Not Handled
1. **Files in subfolders:** Only scans the top level of the source folder (uses `folder_path.glob(pattern)` not `rglob`)
2. **Mixed case extensions:** While patterns use `re.IGNORECASE`, the file finding uses hardcoded lowercase patterns (`*.xlsx`, `*.xls`, `*.xlsm`) which should work on Windows (case-insensitive FS) but may miss files on Linux
3. **Unicode in filenames:** The sanitization removes control characters but doesn't handle full Unicode normalization — two visually identical names using different Unicode forms would be treated as different clients
4. **Filenames with multiple hyphens in client name:** Since the regex uses `[^-]+` for client name capture, a client name like "ABC-DEF Corp" would only capture "ABC" as the client name and "DEF Corp" as the state

### Performance Concerns
5. **Sequential processing:** Clients are processed one at a time. With 50+ clients, each requiring Excel COM operations, this could take hours
6. **No progress persistence:** If the application crashes mid-processing in "fresh" mode, there's no way to resume — must start over (the "resume" mode only works if the folder structure was already created)
7. **Excel COM is fragile:** Using keyboard simulation (SendKeys) for Power Query refresh is inherently unreliable — any popup, notification, or focus change can break the sequence

### Security Concerns
8. **No file path sanitization for directory traversal:** The `sanitize_filename` function handles invalid characters but doesn't check for `..` path traversal in client names extracted from filenames
9. **`os.startfile()` call:** After processing, automatically opens the output folder and export files using `os.startfile()` which could be exploited if filenames were crafted maliciously (though the risk is low since files come from the user's own folders)
10. **Cache file in home directory:** Settings including file paths are stored in plaintext JSON in the user's home directory

### Usability Gaps
11. **No undo:** Once processing is complete, there's no way to undo the file organization
12. **No progress for individual file copies:** The progress bar updates per-client, not per-file, so for a client with many files, the bar may appear stuck
13. **Logging is excessive for production:** Many `[DEBUG]` log lines remain in production code (especially in excel_handler.py and processing_handler.py), though they've been partially cleaned up

---

## SECTION 8: DESIGN DECISIONS (INFERRED)

### Why Tkinter over PyQt5/wxPython?
Tkinter is Python's built-in GUI library — no additional installation required. For a tool distributed to CA/GST consultants who may not be technical, minimizing dependencies is critical. **Confidence: HIGH** (the README specifically notes minimal dependencies, and a PyQt5 migration was later attempted but not deployed, suggesting the simplicity of Tkinter was valued)

### Why Two Separate Applications?
The main organizer and Power Query Extractor are separate programs with separate entry points. This is because they serve different stages of the workflow: the organizer is run first (once per batch), then the extractor is run after the user has verified the organized files. Separating them also reduces the risk of the PQ Extractor's heavy Excel COM operations affecting the main app. **Confidence: HIGH** (the launch_extractor.py file specifically says "Run this file separately after GST Organizer processing")

### Why Excel COM Instead of Pure openpyxl?
Power Query connections cannot be refreshed by openpyxl — they require the actual Excel application to execute. The win32com approach was chosen specifically to preserve Power Query functionality, which is the core value proposition for the extraction workflow. The openpyxl fallback exists for non-Windows platforms but explicitly warns that Power Query will be lost. **Confidence: HIGH** (excel_handler.py line 3: "WINDOWS VERSION - Uses Excel COM to preserve Power Query and all Excel features")

### Why Keyboard Simulation for Power Query Refresh?
The Excel COM API does have a `RefreshAll()` method, but it's known to be unreliable with complex Power Query connections — it may return before the refresh is complete, or not trigger all connection types. The SendKeys approach (Alt+A, R, A) uses the same UI path a human would use, which is more reliable but fragile. **Confidence: MEDIUM** (the extensive retry logic, error window detection, and configurable wait times all suggest this approach was arrived at through trial and error)

### Why State Codes Instead of Full Names?
Windows has a 260-character path limit. With nested folders (Annual Statement/Client-State/Version/Category/filename.xlsx), long state names like "Dadra and Nagar Haveli and Daman and Diu" would quickly exceed the limit. State codes keep paths short. **Confidence: HIGH** (the commit `6237597` added state code functions alongside fixes for path-related issues, and report_processor.py has extensive long-path handling code)

### Why Copy Files Instead of Move/Link?
The README explicitly states "Original files are never modified." Copying preserves the original files as a safety measure — if something goes wrong with the organized copies, the originals are untouched. **Confidence: HIGH** (the safe_copy_file function includes verification, and backups are created when overwriting)

### Why Sanitize Filenames But Preserve Hyphens?
Hyphens are part of the GST file naming convention (e.g., "GSTR-2B-Reco-Client-State-Period"). Removing them would break the naming pattern. Other special characters are removed because they could cause issues with Excel formulas, Windows paths, or the regex parsing. **Confidence: HIGH** (helpers.py line 81: "Define invalid characters WITHOUT including hyphen" — the comment is explicit)

### Why a Configurable Client Name Max Length?
This was added to handle the Windows path length limit problem. Some client names are very long ("ABC Private Limited Industries"), and when combined with folder nesting, they exceed 260 characters. The default of 35 characters was likely chosen through experimentation. **Confidence: HIGH** (the feature was added in commit `c258047` alongside other path-length fixes)

### Evidence of Iteration and Refinement
- **Commented-out code:** None found in current codebase — it was cleaned up
- **Debug logging:** Extensive `[DEBUG]` log lines in excel_handler.py and processing_handler.py suggest these areas were particularly difficult to get right and were debugged extensively
- **The `cell_mappings copy.py` file** in the initial commit was a backup copy that was removed later — suggests the mappings were being actively developed and the backup was a safety measure
- **The Sales report creation code** in processing_handler.py has 6 numbered "debug checkpoints" (lines 268-343) — this suggests the Sales report generation was failing in production and required systematic debugging
- **Multiple retry patterns** in report_processor.py (5 retries for Excel properties, 3 for workbook open, multiple path strategies) suggest each was added after encountering a specific failure in real use

---

## SECTION 9: DEPENDENCIES AND ENVIRONMENT

### External Libraries

| Library | Version | Purpose | Required? |
|---------|---------|---------|-----------|
| openpyxl | >= 3.1.2 | Read/write Excel .xlsx files, create summary reports | **Yes** |
| lxml | >= 4.9.0 | Faster XML parsing for openpyxl (Excel files are XML inside ZIP) | Optional, recommended |
| pywin32 | (not in requirements.txt) | Excel COM automation for Power Query refresh | **Yes for PQ features** |

> **Note:** The requirements.txt lists only openpyxl and lxml. pywin32 is mentioned in the README but not in requirements.txt. The code handles its absence gracefully (falls back to openpyxl-only mode).

### Python Version
- Minimum: Python 3.7 (uses f-strings, Path objects, typing module)
- Recommended: Python 3.8+ (as stated in README)
- Evidence of Python 3.12 and 3.13 usage: the initial commit included __pycache__ files compiled with both versions

### OS-Specific Dependencies
- **Windows:** Required for Power Query features (win32com, pythoncom, win32gui, win32api, win32con)
- **macOS/Linux:** Can run the file organizer but cannot do Power Query refresh or extraction
- **Windows path format:** The `clean_windows_path()` function always converts to backslashes for Excel compatibility

### File System Assumptions
- Source folder contains Excel files at the top level (no recursive scanning)
- Target folder must be writable
- Cache file stored at `~/.gst_organizer_cache.json` (user's home directory)
- Log files: `gst_organizer.log` (main app), `pq_extractor.log` (extractor) — written to current working directory
- Temp files: `%TEMP%/gst_excel_temp/` — used by ExcelHandler for temporary Excel operations
- Windows 260-character path limit is actively worked around

### Setup Steps for Fresh Machine
1. Install Python 3.8+ from python.org (check "Add to PATH")
2. Download/clone this repository
3. Open terminal in project directory
4. `pip install openpyxl lxml`
5. For Power Query features on Windows: `pip install pywin32` then run `python Scripts/pywin32_postinstall.py -install` as administrator
6. `python main.py` to launch

---

## SECTION 10: UI/INTERFACE DOCUMENTATION

### Main Application Window (main_window.py)

**Window Properties:**
- Title: "GST File Organizer v3.0"
- Size: 1200x800 pixels, minimum 1000x700
- Theme: 'clam' ttk style
- Centered on screen at launch

**Title Bar (title_bar.py):**
- Blue banner (#0078D4) with white text
- Title: "GST FILE ORGANIZER & REPORT GENERATOR"
- Subtitle: "Organize files and generate Excel reports automatically"
- Dark Mode checkbox in top-right corner

**Tab 1: Setup (setup_tab.py)**
- Two-column layout: left 60% for inputs, right 40% for help
- Left column sections (top to bottom):
  1. Welcome banner with green border
  2. SOURCE FOLDER: Green header, text entry + Browse button
  3. EXCEL TEMPLATES: Blue header, ITC template entry + Browse, Sales template entry + Browse
  4. TARGET FOLDER: Red header (emphasized with warning banner), text entry + large Browse button
  5. PROCESSING MODE: Cyan header, three radio buttons (Fresh/Re-Run/Resume), "Include client name" checkbox, max length spinbox
  6. Action section: "SCAN FILES" green button + "RE-SCAN" orange button
- Right column: Instructions card (blue) + Expected File Names card (purple)
- Scrollable with mouse wheel

**Tab 2: Validation (validation_tab.py)**
- Two-column layout: left 40% info, right 60% client list
- Left column:
  1. Collapsible instructions section (orange)
  2. Scan Summary (blue header, Consolas font, read-only text showing statistics)
  3. Action buttons: DRY RUN (orange), EXPORT LIST (cyan), START (green)
- Right column:
  1. CLIENT SELECTION header (green)
  2. Keyboard hint: "Use arrows + SPACE to select"
  3. Select All / Clear All / Complete Only buttons
  4. TreeView with columns: Checkbox, Client Name, State, Status, Files, Missing, Extra, FolderName
  5. Color coding: green background for complete, orange for incomplete
  6. Horizontal and vertical scrollbars

**Tab 3: Processing (processing_tab.py)**
- Progress section (blue header):
  - Progress bar (determinate mode)
  - Progress percentage label
  - Current operation label
  - START PROCESSING (green) and STOP (red) buttons
- Log section (green header):
  - Dark-themed ScrolledText (black background #1E1E1E)
  - Color-coded log levels: green=success, orange=warning, red=error, blue=info, white=normal
  - Consolas 9pt font
  - Timestamps on every line

**Status Bar:**
- Dark gray bar (#323130) at bottom
- Status message on left with lightbulb emoji
- "v3.0 | Production Ready" on right

### Power Query Extractor Window (extractor_window.py)

**Window Properties:**
- Title: "Power Query Extractor"
- Size: 1000x700 pixels, minimum 900x600

**Title Bar:**
- Blue banner with "Power Query Report Extractor"

**Tab 1: Setup**
- Two-column layout:
- Left column (fixed 400px):
  1. Target Folder: path entry + Browse + Scan buttons
  2. Processing Options: wait time spinbox (5-60 seconds), file suffix pattern entry, skip refresh checkbox
  3. Start Processing button (disabled until clients found)
- Right column:
  1. Client list with per-client ITC and Sales checkboxes
  2. Buttons: Select All, Deselect All, ITC Only, Sales Only
  3. Last Refresh Status shown per client
  4. Scrollable with mouse wheel

**Tab 2: Processing**
- Progress label + progress bar
- Processing log (light background, Consolas font, color-coded)

### Typical User Workflow
1. Launch `python main.py`
2. Setup tab: Browse to source folder, select ITC + Sales templates, select target folder
3. Choose "Fresh Run" mode
4. Click "Scan Files" → automatically switches to Validation tab
5. Review client list, select desired clients (or "Select All")
6. Optionally click "Dry Run" to preview
7. Click "Start Processing" → switches to Processing tab
8. Watch progress bar and log until completion
9. Output folder opens automatically
10. Launch `python launch_extractor.py`
11. Extractor auto-loads target folder from cache
12. Select clients and ITC/Sales checkboxes
13. Configure wait time if needed
14. Click "Start Processing"
15. Wait for Power Query refresh + extraction
16. Consolidated report opens automatically

---

## SECTION 11: RECONSTRUCTION BLUEPRINT

### 1. Foundation

**Environment Setup:**
- Python 3.8+ with pip
- `pip install openpyxl lxml pywin32` (on Windows)
- Project structure: `core/`, `utils/`, `gui/`, `power_query_extractor/`

**Build Order:**
1. `utils/constants.py` — Define all patterns, mappings, config (no dependencies)
2. `utils/helpers.py` — File operations, sanitization, path utilities
3. `core/file_parser.py` — File scanning and pattern matching
4. `core/file_organizer.py` — Folder creation and file copying
5. `core/excel_handler.py` — Excel report creation
6. GUI layer (can be built independently of core)
7. `power_query_extractor/` — Separate sub-application

### 2. Feature Priority

| Priority | Feature | Reason |
|----------|---------|--------|
| **Essential** | F1: File Scanning | Nothing works without this |
| **Essential** | F3: File Organization | Core value proposition |
| **Essential** | F4: Report Generation | Clients need the Excel reports |
| **Essential** | F14: State Codes | Required for folder naming |
| **High** | F2: Completeness Analysis | Users need to know what's missing |
| **High** | F5: Summary Report | Audit trail for the consultant |
| **High** | F10: Session Persistence | Saves time on repeated use |
| **Medium** | F7: PQ Refresh | High value but Windows-only and fragile |
| **Medium** | F8/F9: Extraction | Depends on F7 working |
| **Medium** | F13: Dry Run | Safety feature |
| **Low** | F11: Dark Mode | Cosmetic |
| **Low** | F12: Client Export | Niche use case |
| **Low** | F6: Org Report | Supplementary documentation |
| **Low** | F15: Configurable Names | Power user feature |

### 3. Known Improvements for Rebuild

1. **Use QThread (PyQt5) instead of threading.Thread** — Better integration with GUI event loop, no need for `root.after()` hacks
2. **Use Excel COM `RefreshAll()` with proper wait** instead of SendKeys — More reliable, less fragile
3. **Add recursive folder scanning** — Real-world users may have files in subfolders
4. **Add GSTIN validation** — Extract and validate the 15-digit GSTIN from filenames or file contents
5. **Add progress persistence** — Save processing state so interrupted runs can truly resume
6. **Separate the PQ Extractor into the main app** as a fourth tab (the PyQt5 branch attempted this)
7. **Replace flat cache file with structured config** — Use proper config directory (`~/.gst_organizer/`) with separate files for settings, history, and state
8. **Add file size check before processing** — Verify sufficient disk space for all planned copies
9. **Handle client names with hyphens** — Current regex breaks on "ABC-DEF Corp" style names

### 4. Estimated Complexity

| Feature | Complexity | Notes |
|---------|-----------|-------|
| F1: File Scanning | MODERATE | Regex patterns are the tricky part |
| F2: Completeness | SIMPLE | Straightforward set comparison |
| F3: File Organization | MODERATE | Many edge cases in folder creation |
| F4: Report Generation | COMPLEX | Excel COM is always complex |
| F5: Summary Report | MODERATE | openpyxl styling takes time |
| F7: PQ Refresh | COMPLEX | Most fragile part of the system |
| F8/F9: Extraction | MODERATE | Depends on stable COM foundation |
| F10: Cache | SIMPLE | JSON read/write |
| F11: Dark Mode | MODERATE | Recursive widget walking is tricky in Tkinter |
| F14: State Codes | SIMPLE | Just a lookup dictionary |
| F15: Configurable Names | SIMPLE | UI wiring + string truncation |
| GUI (complete) | COMPLEX | Many widgets, events, threading |

### 5. Suggested Claude Code Skills

| Feature Area | Skill Description |
|-------------|-------------------|
| Excel COM Automation | Skill for safely opening/editing/saving Excel files via win32com with retry logic, path handling, and cleanup |
| Tkinter Tabbed App | Skill for creating multi-tab Tkinter apps with handler delegation, dark mode, and thread-safe updates |
| File Pattern Organizer | Skill for scanning folders, matching regex patterns, and creating organized folder structures |
| Config-Driven Data Extraction | Skill for reading specific cells from Excel files based on a declarative configuration |

---

## SECTION 12: CHANGELOG ARCHAEOLOGY (FROM GIT HISTORY)

### 12A. Complete Development Timeline

#### Phase 1: Initial Creation (2025-08-31)

**Commit `edac6a3` — "Initial commit" — 2025-08-31**
- **What changed:** Massive initial commit with 110 files, 12,916 lines of code. The entire application was committed at once — core logic, GUI, Power Query Extractor, utilities, README, and even compiled __pycache__ files and a 5,689-line log file.
- **Why:** This was not an incremental development — the application was built in AI-assisted conversations and then committed as a complete working unit. The presence of both Python 3.12 and 3.13 __pycache__ files suggests it was tested on two Python versions.
- **What it reveals:** The app was substantially complete from day one. The log file (5,689 lines) shows it had been run extensively before the initial commit, processing real client data. The README was already comprehensive, suggesting the AI assistant generated it as part of the development process.

**Commit `cbfa4a5` — "Delete gst_organizer.log" — 2025-08-31 (3 minutes later)**
- **What changed:** Removed the accidentally committed 5,689-line log file.
- **Why:** The log contained real processing data that shouldn't be in version control.
- **What it reveals:** The developer was new to git — committing everything and then immediately removing sensitive files.

#### Phase 2: Critical Bug Fixes (2025-10-29, ~2 months later)

**Commit `6237597` — "Fix critical import and threading bugs" — 2025-10-29**
- **What changed:** 4 files, 59 insertions, 45 deletions. Fixed circular import in helpers.py (STATE_CODE_MAPPING), fixed scanned_files attribute access in file_handler.py, fixed thread safety in processing_handler.py GUI updates, fixed indentation bug in client name validation, corrected reference to non-existent `app.py` in error messages, improved progress callback handling.
- **Why:** The app was crashing on launch and during processing. These are the bugs you find when you actually try to use the software — circular imports, threading race conditions, incorrect attribute names.
- **What it reveals:** The initial AI-generated code had structural bugs that only appeared at runtime. The 2-month gap suggests the developer tried to use the tool, hit these bugs, and came back for fixes.

**Commit `9c56a1c` — "Merge branch 'main'" — 2025-10-29**
- **What changed:** Merge commit reconciling local bug fixes with remote (which had the log file deletion).
- **Why:** Standard git workflow — local and remote had diverged.

**Commit `d674386` — "Update README.md" — 2025-10-29**
- **What changed:** 239 insertions, 135 deletions to README.
- **Why:** README was updated to match the actual state of the application after bug fixes.
- **What it reveals:** The README was being kept as living documentation.

**Commit `5ea77c2` — "Remove shebang line from main.py" — 2025-10-29**
- **What changed:** Removed `#!/usr/bin/env python3` from main.py.
- **Why:** Unnecessary on Windows (the target platform). Cleanup.

**Commit `1cb878b` — "Add .gitignore" — 2025-10-29**
- **What changed:** 79-line .gitignore covering Python caches, IDE files, logs, temp files, Excel temp files, output directories, user data, and Power Query cache files.
- **Why:** Proper project hygiene after the initial log file mistake.
- **What it reveals:** The developer learned from the log file incident. The .gitignore is comprehensive and domain-specific (e.g., excludes `~$*.xlsx` Excel temp files, `temp_gst_*` prefix files).

**Commit `01862a0` — "Remove __pycache__ directories" — 2025-10-29**
- **What changed:** Removed 70 compiled .pyc files that were accidentally tracked.
- **Why:** These should never have been in git. Now excluded by .gitignore.

**Commit `1c48722` — "Remove redundant text color handling" — 2025-10-29**
- **What changed:** Removed 6 lines from client_handler.py.
- **Why:** Text color was being set redundantly in the client tree display.
- **What it reveals:** Fine-tuning of the UI after the major bug fixes.

**Commit `4c42758` — "Add mouse wheel scrolling" — 2025-10-29**
- **What changed:** Added 30 lines to extractor_window.py for mouse wheel scrolling on both the main client list and the processing log.
- **Why:** The scrollable areas didn't respond to mouse wheel — a usability issue discovered during real use.
- **What it reveals:** The PQ Extractor was being actively used and polished.

#### Phase 3: Feature Enhancement (2025-10-30)

**Commit `bd176bf` — "Enhance report processing UI" — 2025-10-30**
- **What changed:** 412 insertions, 162 deletions across 3 files. Major redesign of the PQ Extractor window: added two-column layout, configurable wait time (5-60 seconds spinbox), file suffix pattern configuration, skip refresh checkbox, per-client ITC/Sales report checkboxes (instead of all-or-nothing), refresh status display per client.
- **Why:** The original PQ Extractor was inflexible — fixed wait time, no way to skip refresh, no way to process only ITC or only Sales. Real-world use revealed these needs.
- **What it reveals:** The consultant needed more control over the PQ refresh process. The wait time becoming configurable (5-60 seconds) shows different reports took different amounts of time to refresh. The ITC/Sales selection shows some clients only needed one type of report.

**Commit `47a1492` — "Add new ITC report mappings" — 2025-10-30 (HEAD of main)**
- **What changed:** Added 126 lines of new extraction configurations and removed the `cell_mappings copy.py` backup file. New extraction sheets: Reco Diff (cells C37/D37/E37 from 'Reco' sheet), Table 6J Match (K29 from 'T6 of R9'), Table 8D Limits (G6 from 'T8 of R9'), IMS Reco (AA1/AD1/AG1 from 'IMS_Reco'), Extra In Purchase (O1/P1/Q1 from 'ExtraInPurchase'), Extra In GSTR2B (S1/T1/U1 from 'ExtraInGSTR2B').
- **Why:** The consultant needed to extract more data points from the ITC reports. These are specific GST reconciliation values that need to be consolidated across all clients.
- **What it reveals:** The extraction configuration is growing organically as the consultant discovers more data points they need. The cell references (AA1, AD1, AG1) suggest these are in wide, complex Excel sheets with many columns. The "T6 of R9" and "T8 of R9" sheet names refer to specific tables in the GSTR-9 annual return form.

#### Phase 4: PyQt5 Migration Attempt (2026-02-07, NOT MERGED)

> **Important:** The following 3 commits exist on a branch called `copilot-worktree-2026-03-04T11-52-56` and have NOT been merged to main. They represent an attempted but incomplete migration.

**Commit `c258047` — "Refactor v3.0 -> v4.0: PyQt5 migration" — 2026-02-07**
- **What changed:** 3,070 insertions, 541 deletions across 33 files. Complete architectural overhaul: added models/ layer (dataclasses for state), PyQt5 views replacing Tkinter tabs, QThread-based controllers, QSS stylesheet theming, merged PQ Extractor into main app as Tab 4, run history tracking, and significant code cleanup (removed 50 debug lines, fixed 26 bare except clauses, broke up 344-line function into 7 methods).
- **Why:** The Tkinter GUI had accumulated technical debt (recursive widget walking for dark mode, thread-unsafe GUI updates, god-object main window). PyQt5 offers better threading, styling, and widget capabilities.
- **What it reveals:** This was a major investment — 20 new files, complete re-architecture. Co-authored with Claude Opus 4.6, suggesting another AI-assisted development session.

**Commit `a62d3cc` — "Remove old Tkinter GUI files" — 2026-02-07**
- **What changed:** Deleted 17 Tkinter-based files (3,298 lines removed).
- **Why:** Cleanup after PyQt5 migration — old files no longer needed.

**Commit `c21f788` — "Fix gui/__init__.py import" — 2026-02-07**
- **What changed:** Fixed the gui/__init__.py to import from PyQt5 main window instead of deleted Tkinter module.
- **Why:** The previous commit deleted main_window.py but didn't update the import.
- **What it reveals:** The migration was done hastily — the import fix was needed immediately after the deletion.

> **Current Status:** These commits are on a worktree branch, not main. The main branch remains at the Tkinter v3.0 code. The PyQt5 migration may have been abandoned, deferred, or is awaiting testing.

### 12B. File Evolution Map

| File | Created | Last Modified | Evolution |
|------|---------|---------------|-----------|
| main.py | edac6a3 (initial) | 5ea77c2 (shebang removal) | Minimal changes — stable entry point |
| core/file_parser.py | edac6a3 | 6237597 (bug fixes) | Progress callback fix, import fix |
| core/file_organizer.py | edac6a3 | Never changed on main | Stable from day one |
| core/excel_handler.py | edac6a3 | Never changed on main | Stable from day one |
| utils/helpers.py | edac6a3 | 6237597 (import fix) | One-line circular import fix |
| utils/constants.py | edac6a3 | Never changed | Stable from day one |
| gui/handlers/processing_handler.py | edac6a3 | 6237597 (thread safety) | Major thread-safety fixes |
| gui/handlers/file_handler.py | edac6a3 | 6237597 (attribute fix) | One-line fix |
| gui/handlers/client_handler.py | edac6a3 | 1c48722 (color fix) | Minor UI cleanup |
| PQ extractor_window.py | edac6a3 | bd176bf (major UI redesign) | Most evolved file — 3 significant changes |
| PQ report_processor.py | edac6a3 | bd176bf (configurable) | Wait time and suffix made configurable |
| PQ cell_mappings.py | edac6a3 | 47a1492 (6 new sheets) | Actively growing configuration |

### 12C. Logic and Rule Evolution

- **State code mapping:** Added in 6237597 to fix circular import — the mapping existed from the start but the import path was wrong, causing crashes
- **Cell extraction configs:** Started with 3 sheets (GSTR-3B Diff, GSTR-1 Diff, Turnover) and grew to 9 sheets by 47a1492. Rules were added incrementally based on real-world needs.
- **Processing handler:** Thread safety was significantly improved in 6237597, suggesting real-world crashes during processing
- **No rules were removed or modified** — all changes were additive, suggesting the business rules were correct from the start

### 12D. Edge Cases Discovered in Production

| Commit | Bug | Fix | Similar Unhandled Cases? |
|--------|-----|-----|------------------------|
| 6237597 | Circular import: helpers.py importing from constants.py | Changed import to lazy import inside function | Other circular import risks exist if modules grow |
| 6237597 | `scanned_files` accessed before scan completed | Fixed attribute reference | Other attributes could have similar timing issues |
| 6237597 | GUI updates from processing thread causing crashes | Wrapped in `root.after(0, ...)` | Some updates may still not be thread-safe |
| 6237597 | Indentation bug in client name validation | Fixed indentation | Suggests the code wasn't being linted |
| 4c42758 | Mouse wheel not working in scrollable areas | Added explicit mousewheel bindings | Other scrollable areas might need similar treatment |
| bd176bf | Fixed wait time being too short for some reports | Made configurable (5-60 seconds) | Very complex reports might need >60 seconds |

### 12E. Abandoned Approaches and Reversals

1. **`cell_mappings copy.py`** — A backup copy of cell_mappings.py existed from the initial commit and was removed in 47a1492. This suggests the developer was manually backing up config files before making changes, a practice that became unnecessary once git was being used properly.

2. **PyQt5 Migration** — A comprehensive migration to PyQt5 was attempted (c258047) but exists only on a worktree branch, not merged to main. This is the most significant abandoned approach. It included 20 new files and a complete architectural overhaul. The fact that it wasn't merged suggests either: (a) testing revealed issues, (b) the developer decided Tkinter was good enough, or (c) the migration is still in progress.

3. **`gui/app.py`** — The initial commit's main.py error messages referenced a `gui/app.py` file that never existed in the repository. Fixed in 6237597. This suggests the application was originally planned as a single-file GUI before being modularized.

### 12F. Dependency Evolution

| Commit | Change | Reason |
|--------|--------|--------|
| edac6a3 | Initial requirements.txt: openpyxl>=3.1.2, lxml>=4.9.0, optional dev deps | Foundation dependencies |
| c258047 (unmerged) | Added PyQt5>=5.15.0, structured sections | PyQt5 migration attempt |

The dependencies have been remarkably stable. No library was ever replaced or removed on the main branch. pywin32 is used in the code but was never added to requirements.txt — this is a documentation gap.

---

*End of audit. Generated by Claude Opus 4.6 on 2026-03-04.*
