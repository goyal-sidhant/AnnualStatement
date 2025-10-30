"""
Cell Mapping Configuration - MASTER VERSION
This file controls what data is extracted and where it goes

HOW TO EDIT THIS FILE:
1. Each section in EXTRACTION_CONFIG creates a new sheet
2. Format for each entry:
   'Output Sheet Name': {
       'report_type': 'ITC' or 'Sales',
       'mappings': [
           {
               'input_sheet': 'Name of sheet in report',
               'input_cell': 'A1',  # Cell reference
               'output_column': 'Column name in output',
               'calculation': None  # or {'type': 'subtract', 'cell1': 'A1', 'cell2': 'A2'}
           },
           # Add more mappings here
       ]
   }

CALCULATION TYPES:
- None: Just copy the value
- 'subtract': cell1 minus cell2
- 'add': cell1 plus cell2
- 'multiply': cell1 times cell2
- 'divide': cell1 divided by cell2
"""

# ============================================================================
# MASTER CONFIGURATION - EDIT THIS SECTION
# ============================================================================

EXTRACTION_CONFIG = {
    # GSTR-3B Differences Sheet
    'GSTR-3B Difference': {
        'report_type': 'ITC',
        'mappings': [
            {
                'input_sheet': 'Diff GSTR-3B',
                'input_cell': 'J15',
                'output_column': 'Diff IGST',
                'calculation': None
            },
            {
                'input_sheet': 'Diff GSTR-3B',
                'input_cell': 'K15',
                'output_column': 'Diff CGST',
                'calculation': None
            },
            {
                'input_sheet': 'Diff GSTR-3B',
                'input_cell': 'L15',
                'output_column': 'Diff SGST',
                'calculation': None
            },
            # Add more cells here as needed
        ]
    },
    
    # Sales Differences Sheet
    'GSTR-1 Difference': {
        'report_type': 'Sales',
        'mappings': [
            {
                'input_sheet': 'Diff - GSTR-1',
                'input_cell': 'N1',
                'output_column': 'Diff Taxable Value',
                'calculation': None
            },
            {
                'input_sheet': 'Diff - GSTR-1',
                'input_cell': 'O1',
                'output_column': 'Diff IGST',
                'calculation': None
            },
            {
                'input_sheet': 'Diff - GSTR-1',
                'input_cell': 'P1',
                'output_column': 'Diff CGST',
                'calculation': None
            },
            {
                'input_sheet': 'Diff - GSTR-1',
                'input_cell': 'Q1',
                'output_column': 'Diff SGST',
                'calculation': None
            },
        ]
    },

    # Turnover Sheet
    'Turnover GSTR-1': {
        'report_type': 'Sales',
        'mappings': [
            {
                'input_sheet': 'Diff - GSTR-1',
                'input_cell': 'I1',
                'output_column': 'Taxable Value',
                'calculation': None
            },
        ]
    },    
    
    # Reco Diff Sheet
    'Reco Diff': {
        'report_type': 'ITC',
        'mappings': [
            {
                'input_sheet': 'Reco',
                'input_cell': 'C37',
                'output_column': 'IGST Diff',
                'calculation': None
            },
            {
                'input_sheet': 'Reco',
                'input_cell': 'D37',
                'output_column': 'CGST Diff',
                'calculation': None
            },
            {
                'input_sheet': 'Reco',
                'input_cell': 'E37',
                'output_column': 'SGST Diff',
                'calculation': None
            },
        ]
    },

    # Table 6J Match Sheet
    'Table 6J Match': {
        'report_type': 'ITC',
        'mappings': [
            {
                'input_sheet': 'T6 of R9',
                'input_cell': 'K29',
                'output_column': 'Table 6J Match',
                'calculation': None
            },
        ]
    },

    # Table 8D Limits Sheet
    'Table 8D Limits': {
        'report_type': 'ITC',
        'mappings': [
            {
                'input_sheet': 'T8 of R9',
                'input_cell': 'G6',
                'output_column': 'Table 8D Limits',
                'calculation': None
            },
        ]
    },

    # IMS Reco Sheet
    'IMS Reco': {
        'report_type': 'ITC',
        'mappings': [
            {
                'input_sheet': 'IMS_Reco',
                'input_cell': 'AA1',
                'output_column': 'IGST',
                'calculation': None
            },
            {
                'input_sheet': 'IMS_Reco',
                'input_cell': 'AD1',
                'output_column': 'CGST',
                'calculation': None
            },
            {
                'input_sheet': 'IMS_Reco',
                'input_cell': 'AG1',
                'output_column': 'SGST',
                'calculation': None
            },
        ]
    },

    # Extra In Purchase Sheet
    'Extra In Purchase': {
        'report_type': 'ITC',
        'mappings': [
            {
                'input_sheet': 'ExtraInPurchase',
                'input_cell': 'O1',
                'output_column': 'IGST',
                'calculation': None
            },
            {
                'input_sheet': 'ExtraInPurchase',
                'input_cell': 'P1',
                'output_column': 'CGST',
                'calculation': None
            },
            {
                'input_sheet': 'ExtraInPurchase',
                'input_cell': 'Q1',
                'output_column': 'SGST',
                'calculation': None
            },
        ]
    },

    # Extra In GSTR2B Sheet
    'Extra In GSTR2B': {
        'report_type': 'ITC',
        'mappings': [
            {
                'input_sheet': 'ExtraInGSTR2B',
                'input_cell': 'S1',
                'output_column': 'IGST',
                'calculation': None
            },
            {
                'input_sheet': 'ExtraInGSTR2B',
                'input_cell': 'T1',
                'output_column': 'CGST',
                'calculation': None
            },
            {
                'input_sheet': 'ExtraInGSTR2B',
                'input_cell': 'U1',
                'output_column': 'SGST',
                'calculation': None
            },
        ]
    },
    # ADD MORE SHEETS HERE AS NEEDED
    # Example:
    # 'My Custom Sheet': {
    #     'report_type': 'ITC',  # or 'Sales'
    #     'mappings': [
    #         {
    #             'input_sheet': 'Sheet Name in Report',
    #             'input_cell': 'A1',
    #             'output_column': 'My Column Name',
    #             'calculation': None
    #         },
    #     ]
    # },
}

# ============================================================================
# OPTIONS - Edit these settings
# ============================================================================

EXTRACTION_OPTIONS = {
    'skip_refresh': False,  # Set to True to skip Power Query refresh
    'auto_refresh_if_missing': True,  # Automatically refresh if no refreshed file found
    'show_timestamp': False,  # Add extraction timestamp to report
    'missing_data_text': 'Missing',  # What to show when data is not found
}

# ============================================================================
# DON'T EDIT BELOW THIS LINE
# ============================================================================

# Backward compatibility
CELL_MAPPINGS = {}
FIELD_DISPLAY_NAMES = {}

# Build old format for compatibility
for sheet_name, config in EXTRACTION_CONFIG.items():
    report_type = config['report_type']
    if report_type not in CELL_MAPPINGS:
        CELL_MAPPINGS[report_type] = {}
    
    for mapping in config['mappings']:
        if mapping['input_sheet'] and mapping['input_cell']:
            sheet = mapping['input_sheet']
            if sheet not in CELL_MAPPINGS[report_type]:
                CELL_MAPPINGS[report_type][sheet] = {}
            
            # Create unique field name
            field_name = f"{sheet_name}_{mapping['output_column']}".replace(' ', '_')
            CELL_MAPPINGS[report_type][sheet][mapping['input_cell']] = field_name
            FIELD_DISPLAY_NAMES[field_name] = mapping['output_column']