"""
Constants and configuration for GST File Organizer v3.0
Contains all application constants, patterns, and settings.
"""

import re
from typing import Dict, List, Tuple

# ============================================================================
# FILE PATTERNS - Comprehensive regex patterns for file matching
# ============================================================================

FILE_PATTERNS: Dict[str, Dict[str, any]] = {
    'GSTR-2B-Reco': {
        'pattern': re.compile(r'^GSTR-2B-Reco-([^-]+)-([^-]+)-(.+)\.xlsx?$', re.IGNORECASE),
        'type': 'GSTR-2B Reco',
        'category': 'ITC',
        'folder': 'itc',
        'description': 'GSTR-2B-Reco-ClientName-State-Period.xlsx',
        'groups': ['client', 'state', 'period']
    },
    'ImsReco': {
        'pattern': re.compile(r'^ImsReco-([^-]+)-([^-]+)-(\d{8})\.xlsx?$', re.IGNORECASE),
        'type': 'IMS Reco',
        'category': 'ITC',
        'folder': 'itc',
        'description': 'ImsReco-ClientName-State-DDMMYYYY.xlsx',
        'groups': ['client', 'state', 'date']
    },
    'GSTR3B': {
        'pattern': re.compile(r'^(?:\(\d+\)\s*)?GSTR3B-([^-]+)-([^-]+)-([^-]+)\.xlsx?$', re.IGNORECASE),
        'type': 'GSTR-3B Export',
        'category': 'GSTR3B',
        'folder': 'gstr3b',
        'description': 'GSTR3B-ClientName-State-Month.xlsx',
        'groups': ['client', 'state', 'month']
    },
    'Sales': {
        'pattern': re.compile(r'^Sales-([^-]+)-([^-]+)-([^-]+)-([^-]+)\.xlsx?$', re.IGNORECASE),
        'type': 'Sales',
        'category': 'Sales',
        'folder': 'sales',
        'description': 'Sales-ClientName-State-StartMonth-EndMonth.xlsx',
        'groups': ['client', 'state', 'start_month', 'end_month']
    },
    'SalesReco': {
        'pattern': re.compile(r'^SalesReco-([^-]+)-([^-]+)-(.+)\.xlsx?$', re.IGNORECASE),
        'type': 'Sales Reco',
        'category': 'Sales',
        'folder': 'sales',
        'description': 'SalesReco-ClientName-State-Period.xlsx',
        'groups': ['client', 'state', 'period']
    },
    'AnnualReport': {
        'pattern': re.compile(r'^AnnualReport-([^-]+)-([^-]+)-(.+)\.xlsx?$', re.IGNORECASE),
        'type': 'Annual Report',
        'category': 'Annual',
        'folder': 'root',
        'description': 'AnnualReport-ClientName-State-Year.xlsx',
        'groups': ['client', 'state', 'year']
    }
}

# ============================================================================
# EXPECTED FILE TYPES
# ============================================================================

EXPECTED_FILE_TYPES: List[str] = [
    'GSTR-2B Reco',
    'IMS Reco',
    'GSTR-3B Export',
    'Sales',
    'Sales Reco',
    'Annual Report'
]

# ============================================================================
# FOLDER STRUCTURE TEMPLATES
# ============================================================================

FOLDER_STRUCTURE = {
    'level1': 'Annual Statement-{timestamp}',
    'level2': '{client}-{state}',
    'version': 'Version-{timestamp}',
    'categories': {
        'gstr3b': 'GSTR-3B Exports ({client})',
        'itc': 'Other ITC related files ({client})',
        'sales': 'Sales related files ({client})'
    }
}

# ============================================================================
# EXCEL TEMPLATE MAPPINGS
# ============================================================================

EXCEL_TEMPLATE_MAPPING = {
    'ITC': {
        'sheet': 'Links',
        'cells': {
            'B2': 'gstr3b_folder',
            'B4': 'annual_folder',
            'B5': 'annual_filename',
            'B7': 'gstr2b_folder',
            'B8': 'gstr2b_filename',
            'B10': 'ims_folder',
            'B11': 'ims_filename'
        }
    },
    'Sales': {
        'sheet': 'Links',
        'cells': {
            'B2': 'sales_folder',
            'B3': 'sales_filename',
            'B5': 'annual_folder',
            'B6': 'annual_filename',
            'B8': 'sales_reco_folder',
            'B9': 'sales_reco_filename'
        }
    }
}

# ============================================================================
# GUI CONFIGURATION
# ============================================================================

GUI_CONFIG = {
    'window': {
        'title': 'GST File Organizer v3.0',
        'size': '1200x800',
        'min_size': (1000, 700)
    },
    'fonts': {
        'default': ('Segoe UI', 10),
        'title': ('Segoe UI', 16, 'bold'),
        'heading': ('Segoe UI', 12, 'bold'),
        'small': ('Segoe UI', 9),
        'code': ('Consolas', 9)
    },
    'colors': {
        'primary': '#0078D4',
        'success': '#107C10',
        'warning': '#FF8C00',
        'danger': '#D83B01',
        'info': '#0099BC',
        'light': '#F8F9FA',
        'dark': '#323130',
        'white': '#FFFFFF',
        'complete': '#E8F5E8',
        'incomplete': '#FFF4E5'
    }
}

# ============================================================================
# PROCESSING MODES
# ============================================================================

PROCESSING_MODES = {
    'fresh': {
        'name': '🆕 Fresh Run',
        'description': 'Create new folder structure (recommended for first time)',
        'icon': '🆕'
    },
    'rerun': {
        'name': '🔄 Re-Run',
        'description': 'Add new version folders to existing structure',
        'icon': '🔄'
    },
    'resume': {
        'name': '▶️ Resume',
        'description': 'Continue previously interrupted processing',
        'icon': '▶️'
    }
}

# ============================================================================
# FILE VALIDATION
# ============================================================================

EXCEL_EXTENSIONS = {'.xlsx', '.xls', '.xlsm'}
TEMPLATE_EXTENSIONS = {'.xlsx', '.xltx', '.xltm'}
MAX_FILE_SIZE_MB = 100
EXCEL_SIGNATURES = {
    b'PK\x03\x04': 'xlsx',  # ZIP-based Excel files
    b'\xd0\xcf\x11\xe0': 'xls'  # OLE-based Excel files
}

# ============================================================================
# MESSAGES
# ============================================================================

MESSAGES = {
    'errors': {
        'no_source': 'Please select a source folder containing GST files',
        'no_templates': 'Please select both ITC and Sales template files',
        'no_target': 'Please select a target folder for organized files',
        'no_selection': 'Please select at least one client to process',
        'invalid_excel': 'Invalid Excel file: {file}',
        'template_error': 'Error loading template: {error}',
        'processing_error': 'Error processing {client}: {error}'
    },
    'success': {
        'scan_complete': '✅ Found {count} clients with {files} files',
        'client_complete': '✅ Processed {client} successfully',
        'all_complete': '🎉 All processing completed! {count}/{total} successful'
    },
    'info': {
        'scanning': '🔍 Scanning folder for GST files...',
        'processing': '⚙️ Processing {client}...',
        'creating_report': '📊 Creating {report} report...'
    }
}

# ============================================================================
# EXCEL SAFETY CONFIGURATION
# ============================================================================

EXCEL_SAFETY = {
    'max_retries': 3,
    'retry_delay': 1,  # seconds
    'temp_prefix': 'temp_gst_',
    'backup_prefix': 'backup_',
    'verification_size_min': 10240,  # 10KB minimum
    'save_options': {
        'data_only': False,
        'keep_vba': False,
        'keep_links': False
    }
}

# ============================================================================
# STATE CODE MAPPING
# ============================================================================

STATE_CODE_MAPPING = {
    # Full state names to codes
    'andhra pradesh': 'AP',
    'arunachal pradesh': 'AR', 
    'assam': 'AS',
    'bihar': 'BR',
    'chhattisgarh': 'CG',
    'goa': 'GA',
    'gujarat': 'GJ',
    'haryana': 'HR',
    'himachal pradesh': 'HP',
    'jharkhand': 'JH',
    'karnataka': 'KA',
    'kerala': 'KL',
    'madhya pradesh': 'MP',
    'maharashtra': 'MH',
    'manipur': 'MN',
    'meghalaya': 'ML',
    'mizoram': 'MZ',
    'nagaland': 'NL',
    'odisha': 'OD',
    'orissa': 'OD',  # Alternative name
    'punjab': 'PB',
    'rajasthan': 'RJ',
    'sikkim': 'SK',
    'tamil nadu': 'TN',
    'telangana': 'TS',
    'tripura': 'TR',
    'uttar pradesh': 'UP',
    'uttarakhand': 'UK',
    'west bengal': 'WB',
    'andaman and nicobar islands': 'AN',
    'chandigarh': 'CH',
    'dadra and nagar haveli and daman and diu': 'DD',
    'delhi': 'DL',
    'national capital territory of delhi': 'DL',
    'jammu and kashmir': 'JK',
    'ladakh': 'LA',
    'lakshadweep': 'LD',
    'puducherry': 'PY',
    'pondicherry': 'PY'  # Alternative name
}