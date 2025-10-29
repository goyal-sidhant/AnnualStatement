# 📊 GST File Organizer v3.0 - Production Ready

A robust, enterprise-grade Python application for organizing GST (Goods and Services Tax) files and generating automated reports with Power Query support.

## ⚠️ IMPORTANT: Prevents Excel Corruption

This version specifically includes multiple safety mechanisms to prevent Excel file corruption - a common issue in earlier versions. All file operations use temporary files and verification steps.

## 🎯 Overview

The GST File Organizer is a comprehensive solution designed to:
- **Organize** scattered GST Excel files into a structured folder hierarchy
- **Generate** ITC and Sales reports automatically using Excel templates
- **Refresh** Power Query connections in Excel reports
- **Extract** specific cell values from multiple reports into consolidated sheets
- **Support** batch processing for multiple clients
- **Prevent** Excel file corruption with multiple safety layers

## ✨ Features

### Core Features
- 📁 **Smart File Organization**: Automatically categorizes and organizes GST files by client and state
- 📊 **Automated Report Generation**: Creates ITC and Sales reports using Excel templates
- 🔄 **Power Query Support**: Refreshes data connections in Excel (Windows only)
- 📈 **Data Extraction**: Consolidates specific cell values from multiple reports
- 🎨 **Dark Mode**: Toggle between light and dark themes
- 💾 **Session Memory**: Remembers your last used folders and settings
- 🛡️ **Bulletproof Excel Handling**: Multiple safety layers to prevent file corruption
- 📋 **Comprehensive Summary**: Excel summary report with 5 detailed sheets

### File Type Support
- GSTR-2B Reconciliation files
- IMS Reconciliation files  
- GSTR-3B Export files (supports multiple monthly filings)
- Sales files
- Sales Reconciliation files
- Annual Reports

## 🔒 Excel Corruption Prevention

This version specifically addresses Excel corruption issues with:
- **Safe file operations**: All Excel operations use temporary files
- **Verification**: Files are verified before finalization  
- **Clean saves**: No VBA macros or external links preserved
- **Robust sanitization**: Filenames are thoroughly cleaned
- **Multiple safety layers**: Prevent file corruption at every step
- **Backup creation**: Original files are never modified

## 📋 System Requirements

- **Python**: 3.7 or higher (3.8+ recommended)
- **Operating System**: Windows 10/11 (best), macOS, Linux
- **Memory**: 4GB RAM minimum (8GB recommended)
- **Storage**: 500MB free space for application and temporary files
- **Excel**: Microsoft Excel 2016+ for Power Query features (Windows only)

## 🚀 Quick Start Setup

### Step 1: Install Python (if not already installed)

1. Download Python from [python.org](https://www.python.org/downloads/)
2. During installation, **CHECK** "Add Python to PATH"
3. Verify installation:
```bash
   python --version
```

### Step 2: Download and Extract Application

1. Download all files to a folder (e.g., `C:\GST_Organizer`)
2. Ensure this folder structure exists:
```
   GST_Organizer/
   ├── main.py
   ├── launch_extractor.py
   ├── requirements.txt
   ├── core/
   │   ├── __init__.py
   │   ├── excel_handler.py
   │   ├── file_organizer.py
   │   └── file_parser.py
   ├── gui/
   │   ├── handlers/
   │   ├── tabs/
   │   ├── utils/
   │   └── widgets/
   ├── utils/
   │   ├── __init__.py
   │   ├── constants.py
   │   └── helpers.py
   └── power_query_extractor/
       ├── core/
       ├── gui/
       └── config/
```

### Step 3: Install Dependencies

Open Command Prompt/Terminal in the application folder:
```bash
pip install -r requirements.txt
```

### Requirements.txt Contents
```txt
openpyxl>=3.0.0
pywin32>=305; sys_platform == 'win32'
```

### Step 4: Run the Application
```bash
python main.py
```

## 📖 Usage Guide

### 1. Prepare Your Files

**Excel Templates Requirements:**
- Create or obtain ITC Report template (.xlsx or .xltx)
- Create or obtain Sales Report template (.xlsx or .xltx)
- **IMPORTANT**: Templates must have a "Links" sheet with specific cell mappings

**GST Files Naming Convention (MUST match exactly):**

| File Type | Pattern | Example |
|-----------|---------|---------|
| GSTR-2B Reco | `GSTR-2B-Reco-{Client}-{State}-{Period}.xlsx` | GSTR-2B-Reco-ABC Ltd-Maharashtra-Apr24.xlsx |
| IMS Reco | `ImsReco-{Client}-{State}-{DDMMYYYY}.xlsx` | ImsReco-ABC Ltd-Maharashtra-30042024.xlsx |
| GSTR-3B | `GSTR3B-{Client}-{State}-{Month}.xlsx` | GSTR3B-ABC Ltd-Maharashtra-Apr24.xlsx |
| Sales | `Sales-{Client}-{State}-{Start}-{End}.xlsx` | Sales-ABC Ltd-Maharashtra-Apr-Jun.xlsx |
| Sales Reco | `SalesReco-{Client}-{State}-{Period}.xlsx` | SalesReco-ABC Ltd-Maharashtra-Q1.xlsx |
| Annual Report | `AnnualReport-{Client}-{State}-{Year}.xlsx` | AnnualReport-ABC Ltd-Maharashtra-2024.xlsx |

### 2. Using the Main Application

**Step 1: Setup Tab**
1. Click "Browse" to select **Source Folder** containing GST files
2. Select **ITC Template** file
3. Select **Sales Template** file  
4. Select **Target Folder** for organized files
5. Choose processing mode:
   - 🆕 **Fresh Run**: First time processing (recommended)
   - 🔄 **Re-Run**: Add new versions to existing structure
   - ▶️ **Resume**: Continue interrupted processing
6. Click **"Scan Files"**

**Step 2: Validation Tab**
1. Review scan summary and statistics
2. Check client list:
   - ✅ = All files present
   - ⚠️ = Missing some files
3. Select clients to process:
   - Click checkbox for individual selection
   - Use "Select All", "Clear All", or "Complete Only" buttons
4. Review variations tab for naming issues

**Step 3: Processing Tab**  
1. Click **"Dry Run"** to preview operations (optional)
2. Click **"Start Processing"** to begin
3. Monitor progress bar and detailed log
4. Wait for completion message

### 3. Using Power Query Extractor

After main processing, extract data from reports:
```bash
python launch_extractor.py
```

1. The extractor auto-loads target folder from cache
2. Select clients to process
3. Click **"Refresh & Extract"** to:
   - Refresh Power Query connections
   - Extract configured cell values
   - Generate consolidated Excel report

### 4. Output Structure & Files
```
Target Folder/
└── Annual Statement-DDMMYY HHMM/
    ├── GST_Processing_Summary_DDMMYYYY_HHMMSS.xlsx
    └── ClientName-StateCode/
        └── Version-DDMMYY HHMM/
            ├── GSTR-3B Exports (ClientName)/
            ├── Other ITC related files (ClientName)/
            ├── Sales related files (ClientName)/
            ├── ITC_Report_ClientName_State_DDMMYYYY_HHMMSS.xlsx
            ├── Sales_Report_ClientName_State_DDMMYYYY_HHMMSS.xlsx
            └── _Organization_Report_ClientName_State.txt
```

### Output Files Generated

1. **GST_Processing_Summary_DDMMYYYY_HHMMSS.xlsx** - Master summary with:
   - Summary Sheet: Overall processing statistics
   - Client Status: Status of each client
   - File Mapping: Source to destination mapping  
   - Errors: Detailed error information
   - Variations: Files that don't match naming convention

2. **ITC_Report_ClientName_State_DDMMYYYY_HHMMSS.xlsx** - ITC report from template
3. **Sales_Report_ClientName_State_DDMMYYYY_HHMMSS.xlsx** - Sales report from template
4. **_Organization_Report_ClientName_State.txt** - Processing log for each client

## ⚙️ Configuration

### Template Cell Mappings

Templates must have a "Links" sheet with these cells:

**ITC Template:**
- B2: GSTR-3B folder path
- B4: Annual report folder path
- B5: Annual report filename
- B7: GSTR-2B folder path
- B8: GSTR-2B filename
- B10: IMS folder path
- B11: IMS filename

**Sales Template:**
- B2: Sales folder path
- B3: Sales filename
- B5: Annual report folder path
- B6: Annual report filename
- B8: Sales Reco folder path
- B9: Sales Reco filename

### Cell Extraction Configuration

Edit `power_query_extractor/config/cell_mappings.py` to customize:
```python
EXTRACTION_CONFIG = {
    'GSTR-3B Difference': {
        'report_type': 'ITC',
        'mappings': [
            {
                'input_sheet': 'Diff GSTR-3B',
                'input_cell': 'J15',
                'output_column': 'Diff IGST',
                'calculation': None
            },
        ]
    },
}

EXTRACTION_OPTIONS = {
    'skip_refresh': False,  # Set to True to skip Power Query refresh
    'auto_refresh_if_missing': True,
    'show_timestamp': False,
    'missing_data_text': 'Missing',
}
```

## 🛠️ Troubleshooting

### Excel Files Won't Open / Show Corruption Error

**This version specifically addresses this issue with:**
1. Safe file operations using temporary files
2. File verification before finalization
3. Clean saves without VBA macros or external links
4. Robust filename sanitization

**If issues persist:**
- Check template files aren't corrupted
- Ensure sufficient disk space
- Run as administrator (Windows)
- Check antivirus isn't interfering

### Common Issues and Solutions

| Issue | Solution |
|-------|----------|
| "No module named 'win32com'" | `pip install pywin32` then run `python Scripts/pywin32_postinstall.py -install` as admin |
| "No Excel files found" | Check file extensions (.xlsx, .xls) and naming convention |
| "Permission denied" | Close all Excel files, run as administrator |
| "Template not found" | Ensure templates have "Links" sheet |
| Processing seems frozen | Large files take time, check log for activity |

### Debug Mode

Enable detailed logging:
```python
# In main.py or launch_extractor.py
logging.basicConfig(level=logging.DEBUG)
```

Logs are saved to:
- `gst_organizer.log` - Main application
- `pq_extractor.log` - Power Query extractor

## 🎯 Best Practices

1. **Always backup** your data before processing
2. **Test first** with a small subset of files
3. **Use templates** that you've verified work manually
4. **Check names** match the expected patterns exactly
5. **Review variations** and fix naming issues
6. **Keep templates simple** without complex formulas
7. **Close all Excel files** before processing

## 🔒 Security and Safety

- No data is sent outside your computer
- Original files are never modified  
- Backups created for existing files
- All operations are logged
- Extensive error checking
- Multiple verification steps before file finalization

## ⚠️ Known Limitations

- Excel files must be closed during processing
- Templates must have specific structure
- File names must match patterns exactly
- Windows path limit of 260 characters applies
- Power Query refresh requires Windows and Excel

## 📝 Recent Bug Fixes (Latest Commit)

- ✅ Fixed circular import error in utils/helpers.py (STATE_CODE_MAPPING)
- ✅ Fixed scanned_files attribute access in file_handler.py
- ✅ Resolved thread safety issues in processing_handler.py
- ✅ Corrected client name validation logic indentation
- ✅ Fixed progress callback handling
- ✅ Enhanced Excel corruption prevention mechanisms

## 📞 Support

### Getting Help

1. Check the log file: `gst_organizer.log`
2. Review error messages in the application
3. Ensure all prerequisites are met
4. Try with a small test set first

### Reporting Issues

When reporting issues, include:
- Python version (`python --version`)
- Operating system and version
- Error messages from log
- Steps to reproduce
- Sample file names (anonymized)

## 📄 License

This project is proprietary software for GST file management.

## 🏁 Conclusion

This application is designed to be robust and reliable for production use. The Excel corruption issues from previous versions have been specifically addressed with multiple safety mechanisms.

For enterprise deployments or custom requirements, the modular architecture allows easy extension and modification.

**Remember**: Always test with a small set of files first!

---

**Version:** 3.0 Production Ready  
**Status:** Stable Release  
**Last Updated:** 29th October 2025  
**Author:** Advanced AI Assistant