"""
Utils package for GST File Organizer v3.0
Contains utility functions and constants.
"""

from .constants import (
    FILE_PATTERNS, EXPECTED_FILE_TYPES, FOLDER_STRUCTURE,
    EXCEL_TEMPLATE_MAPPING, GUI_CONFIG, PROCESSING_MODES,
    MESSAGES, EXCEL_EXTENSIONS, EXCEL_SAFETY
)

from .helpers import (
    get_timestamp, get_date_only, get_safe_timestamp,
    ensure_path_exists, safe_path_join, sanitize_filename,
    get_file_info, safe_copy_file, create_backup,
    validate_excel_file, find_excel_files, format_size,
    truncate_path, clean_windows_path, extract_filename_without_extension,
    validate_client_name, get_unique_path, ProgressTracker,
    format_duration, calculate_file_hash, get_relative_excel_path
)

__all__ = [
    # Constants
    'FILE_PATTERNS', 'EXPECTED_FILE_TYPES', 'FOLDER_STRUCTURE',
    'EXCEL_TEMPLATE_MAPPING', 'GUI_CONFIG', 'PROCESSING_MODES',
    'MESSAGES', 'EXCEL_EXTENSIONS', 'EXCEL_SAFETY',
    
    # Helpers
    'get_timestamp', 'get_date_only', 'get_safe_timestamp',
    'ensure_path_exists', 'safe_path_join', 'sanitize_filename',
    'get_file_info', 'safe_copy_file', 'create_backup',
    'validate_excel_file', 'find_excel_files', 'format_size',
    'truncate_path', 'clean_windows_path', 'extract_filename_without_extension',
    'validate_client_name', 'get_unique_path', 'ProgressTracker',
    'format_duration', 'calculate_file_hash', 'get_relative_excel_path'
]

__version__ = '3.0.0'