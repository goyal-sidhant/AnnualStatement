"""
Cell Mapping Configuration
EDIT THIS FILE to specify which cells to extract from reports
"""

# Define which cells to extract from each report type
CELL_MAPPINGS = {
    'ITC': {
        # Sheet name: {cell: field_name}
        'Diff GSTR-3B': {
            'J15': 'itc_diff_igst',
            'K15': 'itc_diff_cgst', 
            'L15': 'itc_diff_sgst',
            # Add more cells here as needed
            # 'E5': 'reversed_itc',
            # 'F5': 'net_itc',
        },
        # Add more sheets if needed
        # 'Another Sheet': {
        #     'A10': 'some_value',
        # }
    },
    'Sales': {
        'Diff - GSTR-1': {
            'N1': 'sales_diff_tv',
            'O1': 'sales_diff_igst',
            'P1': 'sales_diff_cgst',
            'Q1': 'sales_diff_sgst',
            # Add more cells here as needed
            # 'D10': 'export_sales',
            # 'E10': 'exempt_sales',
        },
        # Add more sheets if needed
    }
}

# Optional: Define display names for fields in the consolidated report
FIELD_DISPLAY_NAMES = {
    'sales_diff_tv': 'Diff Taxable Value (Sales)',
    'sales_diff_igst': 'Diff IGST (Sales)',
    'sales_diff_cgst': 'Diff CGST (Sales)',
    'sales_diff_sgst': 'Diff SGST (Sales)',
    'itc_diff_igst': 'Diff IGST (ITC)',
    'itc_diff_cgst': 'Diff CGST (ITC)',
    'itc_diff_sgst': 'Diff SGST (ITC)',
    # Add display names as you add fields
}

"""
HOW TO USE THIS FILE:
1. Add/remove entries in CELL_MAPPINGS dictionary
2. Format: 'Sheet Name': {'Cell': 'field_name'}
3. Optionally add display names in FIELD_DISPLAY_NAMES
4. Save the file - changes take effect immediately
5. No need to modify any other code!

Example:
To add a new cell from ITC report:
- Find the sheet name (e.g., 'Monthly Summary')
- Find the cell reference (e.g., 'E5')
- Add: 'E5': 'reversed_itc',
- Optionally add to FIELD_DISPLAY_NAMES
"""