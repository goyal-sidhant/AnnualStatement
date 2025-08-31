# GST File Organizer v3.0 - Production Ready

A robust, enterprise-grade application for organizing GST files and generating Excel reports without corruption issues.

## 🌟 Key Features

- **Bulletproof Excel Handling**: Multiple safety layers to prevent file corruption
- **Smart File Recognition**: Automatically categorizes 6 types of GST files
- **Intelligent Organization**: Creates hierarchical folder structure
- **Safe Report Generation**: Creates ITC and Sales reports from templates
- **Comprehensive Summary**: Excel summary report with 5 detailed sheets
- **Modern GUI**: User-friendly interface with progress tracking
- **Error Recovery**: Robust error handling and recovery mechanisms

## 📋 System Requirements

- **Python**: 3.7 or higher
- **Operating System**: Windows 10/11, macOS, Linux
- **Memory**: 4GB RAM minimum (8GB recommended)
- **Storage**: 500MB free space for application and temporary files

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
2. Ensure this folder structure:
   ```
   GST_Organizer/
   ├── main.py
   ├── requirements.txt
   ├── README.md
   ├── core/
   │   ├── __init__.py
   │   ├── file_parser.py
   │   ├── file_organizer.py
   │   └── excel_handler.py
   ├── utils/
   │   ├── __init__.py
   │   ├── constants.py
   │   └── helpers.py
   └── gui/
       ├── __init__.py
       └── app.py
   ```

### Step 3: Install Dependencies

Open Command Prompt/Terminal in the application folder and run:

```bash
pip install -r requirements.txt
```

### Step 4: Run the Application

```bash
python main.py
```

## 📖 Usage Guide

### 1. Prepare Your Files

**Excel Templates:**
- Create or obtain ITC Report template (.xlsx or .xltx)
- Create or obtain Sales Report template (.xlsx or .xltx)
- Templates should have a "Links" sheet with cells for data insertion

**GST Files Naming Convention:**
- `GSTR-2B-Reco-ClientName-State-Period.xlsx`
- `ImsReco-ClientName-State-DDMMYYYY.xlsx`
- `GSTR3B-ClientName-State-Month.xlsx`
- `Sales-ClientName-State-StartMonth-EndMonth.xlsx`
- `SalesReco-ClientName-State-Period.xlsx`
- `AnnualReport-ClientName-State-Year.xlsx`

### 2. Using the Application

**Step 1: Setup Tab**
1. Select source folder containing GST files
2. Select ITC Report template file
3. Select Sales Report template file
4. **IMPORTANT**: Select target folder for organized files
5. Choose processing mode (use "Fresh Run" for first time)
6. Click "Scan Files"

**Step 2: Validation Tab**
1. Review scan summary
2. Check client list (✅ = complete, ⚠️ = missing files)
3. Select clients to process (click checkbox or use buttons)
4. View variations if files don't match naming convention
5. Click "Start Processing"

**Step 3: Processing Tab**
1. Monitor progress bar
2. View detailed log
3. Wait for completion message

### 3. Output Structure

```
Target Folder/
└── Annual Statement-DDMMYYYY HHMMSS/
    ├── GST_Processing_Summary_DDMMYYYY_HHMMSS.xlsx
    └── ClientName-State/
        └── Version-DDMMYYYY HHMMSS/
            ├── GSTR-3B Exports (ClientName)/
            ├── Other ITC related files (ClientName)/
            ├── Sales related files (ClientName)/
            ├── ITC_Report_ClientName_State_DDMMYYYY_HHMMSS.xlsx
            ├── Sales_Report_ClientName_State_DDMMYYYY_HHMMSS.xlsx
            └── _Organization_Report_ClientName_State.txt
```

## 🛠️ Troubleshooting

### Excel Files Won't Open / Show Corruption Error

**This version specifically addresses this issue with:**
1. **Safe file operations**: All Excel operations use temporary files
2. **Verification**: Files are verified before finalization
3. **Clean saves**: No VBA macros or external links preserved
4. **Robust sanitization**: Filenames are thoroughly cleaned

**If issues persist:**
1. Check template files aren't corrupted
2. Ensure sufficient disk space
3. Run as administrator (Windows)
4. Check antivirus isn't interfering

### Common Issues and Solutions

**"No Excel files found"**
- Check file extensions (.xlsx, .xls)
- Ensure files aren't open in Excel
- Verify folder permissions

**"Missing module" error**
- Run `pip install -r requirements.txt` again
- Ensure you're using Python 3.7+
- Check virtual environment if using one

**"Permission denied" error**
- Close all Excel files
- Run as administrator
- Check folder write permissions

**Processing seems frozen**
- Large files take time
- Check the log for activity
- Don't close Excel files during processing

## 🔧 Configuration

### Template Customization

Templates should have a "Links" sheet with these cells:

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

### Advanced Settings

Edit `utils/constants.py` to modify:
- File patterns
- Folder structure
- Excel safety settings
- GUI appearance

## 📊 Summary Report Contents

The Excel summary report includes:

1. **Summary Sheet**: Overall processing statistics
2. **Client Status**: Status of each client
3. **File Mapping**: Source to destination mapping
4. **Errors**: Detailed error information
5. **Variations**: Files that don't match naming convention

## 🔒 Security and Safety

- No data is sent outside your computer
- Original files are never modified
- Backups created for existing files
- All operations are logged
- Extensive error checking

## 📞 Support

### Getting Help

1. Check the log file: `gst_organizer.log`
2. Review error messages carefully
3. Ensure all prerequisites are met
4. Try with a small test set first

### Reporting Issues

When reporting issues, include:
- Python version (`python --version`)
- Operating system
- Error messages from log
- Steps to reproduce

## 🎯 Best Practices

1. **Always backup** your data before processing
2. **Test first** with a small subset of files
3. **Use templates** that you've verified work manually
4. **Check names** match the expected patterns exactly
5. **Review variations** and fix naming issues
6. **Keep templates simple** without complex formulas

## 📝 Version History

### v3.0 (Current)
- Complete rewrite for production stability
- Enhanced Excel corruption prevention
- Improved error handling and recovery
- Modern GUI with better feedback
- Comprehensive logging system

### Known Limitations
- Excel files must be closed during processing
- Templates must have specific structure
- File names must match patterns exactly

## 🏁 Conclusion

This application is designed to be robust and reliable for production use. The Excel corruption issues from previous versions have been specifically addressed with multiple safety mechanisms.

For enterprise deployments or custom requirements, the modular architecture allows easy extension and modification.

---

**Remember**: Always test with a small set of files first!