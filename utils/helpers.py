"""
Helper utilities for GST File Organizer v3.0
Robust utility functions with comprehensive error handling.
"""

import os
import re
import time
import shutil
import hashlib
import logging
from pathlib import Path
from datetime import datetime
from typing import Union, List, Optional, Dict, Any, Tuple
import platform

logger = logging.getLogger(__name__)

# ============================================================================
# UPDATED: Time and Date Functions
# ============================================================================

def get_timestamp() -> str:
    """Get formatted timestamp for folders (DDMMYY HHMM)"""
    return datetime.now().strftime("%d%m%y %H%M")

def get_date_only() -> str:
    """Get formatted date (DDMMYY)"""
    return datetime.now().strftime("%d%m%y")

def get_safe_timestamp() -> str:
    """Get timestamp safe for filenames (DDMMYY_HHMM)"""
    return datetime.now().strftime("%d%m%y_%H%M")

# ============================================================================
# Path and File Functions
# ============================================================================

def ensure_path_exists(path: Union[str, Path]) -> Path:
    """Ensure directory exists, create if needed"""
    path = Path(path)
    try:
        path.mkdir(parents=True, exist_ok=True)
        return path
    except Exception as e:
        logger.error(f"Failed to create directory {path}: {e}")
        raise

def safe_path_join(*parts: str) -> Path:
    """Safely join path components"""
    try:
        # Remove empty parts and convert to Path
        valid_parts = [p for p in parts if p]
        if not valid_parts:
            raise ValueError("No valid path components provided")
        
        result = Path(valid_parts[0])
        for part in valid_parts[1:]:
            result = result / part
        
        return result
    except Exception as e:
        logger.error(f"Error joining paths {parts}: {e}")
        raise

def sanitize_filename(filename: str, max_length: int = 200) -> str:
    """
    Sanitize filename with business abbreviations and length limit.
    """
    if not filename:
        return "unnamed"
    
    # Remove file extension if present
    base_name = Path(filename).stem
    extension = Path(filename).suffix
    
    # Apply business abbreviations FIRST
    base_name = base_name.replace('Private', 'Pvt')
    base_name = base_name.replace('Limited', 'Ltd')
    
    # Define invalid characters WITHOUT including hyphen
    invalid_chars = '<>:"/\\|?*[]{}+=!@#$%^,;\'"`~'
    
    # Replace control characters (ASCII 0-31 and 127-159)
    safe_name = ''.join(
        '_' if (ord(c) < 32 or (127 <= ord(c) <= 159)) else c 
        for c in base_name
    )
    
    # Replace other invalid characters
    for char in invalid_chars:
        safe_name = safe_name.replace(char, '_')
    
    # Replace multiple spaces/underscores with single underscore
    safe_name = re.sub(r'[\s_]+', '_', safe_name)
    
    # Remove leading/trailing underscores and dots
    safe_name = safe_name.strip('_. ')
    
    # Ensure not empty
    if not safe_name:
        safe_name = "unnamed"
    
    # Limit length (leave room for extension)
    if len(safe_name) > max_length - len(extension) - 10:
        safe_name = safe_name[:max_length - len(extension) - 10]
    
    # Add extension back if it was present
    if extension:
        safe_name += extension
    
    return safe_name


def get_file_info(file_path: Union[str, Path]) -> Dict[str, Any]:
    """Get comprehensive file information"""
    try:
        path = Path(file_path)
        if not path.exists():
            return {'error': 'File not found'}
        
        stat = path.stat()
        return {
            'name': path.name,
            'path': str(path.absolute()),
            'size': stat.st_size,
            'size_formatted': format_size(stat.st_size),
            'modified': datetime.fromtimestamp(stat.st_mtime),
            'created': datetime.fromtimestamp(stat.st_ctime),
            'extension': path.suffix.lower(),
            'is_file': path.is_file(),
            'is_dir': path.is_dir()
        }
    except Exception as e:
        logger.error(f"Error getting file info for {file_path}: {e}")
        return {'error': str(e)}

# ============================================================================
# File Operations
# ============================================================================

def safe_copy_file(source: Union[str, Path], 
                   destination: Union[str, Path],
                   verify: bool = True) -> bool:
    """
    Safely copy file with verification.
    
    Args:
        source: Source file path
        destination: Destination file path
        verify: Whether to verify copy by comparing sizes
        
    Returns:
        bool: True if successful
    """
    try:
        source_path = Path(source)
        dest_path = Path(destination)
        
        # Validate source
        if not source_path.exists() or not source_path.is_file():
            logger.error(f"Source file not found: {source_path}")
            return False
        
        # Ensure destination directory exists
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Copy file
        shutil.copy2(source_path, dest_path)
        
        # Verify if requested
        if verify:
            source_size = source_path.stat().st_size
            dest_size = dest_path.stat().st_size
            if source_size != dest_size:
                logger.error(f"File size mismatch after copy: {source_size} != {dest_size}")
                dest_path.unlink()  # Remove corrupted copy
                return False
        
        logger.info(f"Successfully copied {source_path.name} to {dest_path}")
        return True
        
    except Exception as e:
        logger.error(f"Error copying file {source} to {destination}: {e}")
        return False

def create_backup(file_path: Union[str, Path]) -> Optional[Path]:
    """Create backup of existing file"""
    try:
        path = Path(file_path)
        if not path.exists():
            return None
        
        timestamp = get_safe_timestamp()
        backup_name = f"{path.stem}_backup_{timestamp}{path.suffix}"
        backup_path = path.parent / backup_name
        
        shutil.copy2(path, backup_path)
        logger.info(f"Created backup: {backup_path}")
        return backup_path
        
    except Exception as e:
        logger.error(f"Error creating backup for {file_path}: {e}")
        return None

# ============================================================================
# Excel Validation
# ============================================================================

def validate_excel_file(file_path: Union[str, Path]) -> bool:
    """
    Validate if file is a genuine Excel file.
    
    Args:
        file_path: Path to file
        
    Returns:
        bool: True if valid Excel file
    """
    try:
        path = Path(file_path)
        
        # Check existence
        if not path.exists() or not path.is_file():
            return False
        
        # Check extension
        if path.suffix.lower() not in {'.xlsx', '.xls', '.xlsm', '.xltx', '.xltm'}:
            return False
        
        # Check file size (too small = probably corrupted)
        if path.stat().st_size < 1024:  # Less than 1KB
            return False
        
        # Check file signature
        with open(path, 'rb') as f:
            header = f.read(8)
        
        # Check for Excel signatures
        xlsx_sig = b'PK\x03\x04'  # ZIP-based
        xls_sig = b'\xd0\xcf\x11\xe0'  # OLE-based
        
        return header.startswith(xlsx_sig) or header.startswith(xls_sig)
        
    except Exception as e:
        logger.error(f"Error validating Excel file {file_path}: {e}")
        return False

def find_excel_files(folder: Union[str, Path]) -> List[Path]:
    """Find all Excel files in folder"""
    try:
        folder_path = Path(folder)
        excel_files = []
        
        patterns = ['*.xlsx', '*.xls', '*.xlsm']
        for pattern in patterns:
            excel_files.extend(folder_path.glob(pattern))
        
        # Filter valid Excel files only
        valid_files = [f for f in excel_files if validate_excel_file(f)]
        
        return sorted(valid_files)
        
    except Exception as e:
        logger.error(f"Error finding Excel files in {folder}: {e}")
        return []

# ============================================================================
# String and Display Functions
# ============================================================================

def format_size(size_bytes: int) -> str:
    """Format file size for display"""
    if size_bytes < 0:
        return "Unknown"
    
    units = ['B', 'KB', 'MB', 'GB', 'TB']
    size = float(size_bytes)
    unit_index = 0
    
    while size >= 1024.0 and unit_index < len(units) - 1:
        size /= 1024.0
        unit_index += 1
    
    return f"{size:.1f} {units[unit_index]}"

def truncate_path(path: Union[str, Path], max_length: int = 60) -> str:
    """Truncate long paths for display"""
    path_str = str(path)
    
    if len(path_str) <= max_length:
        return path_str
    
    path_obj = Path(path_str)
    parts = path_obj.parts
    
    if len(parts) <= 2:
        # Simple truncation
        return f"...{path_str[-(max_length-3):]}"
    
    # Try to keep first and last parts
    first = parts[0]
    last = parts[-1]
    middle_length = max_length - len(first) - len(last) - 8  # 8 for separators
    
    if middle_length > 0:
        return f"{first}{os.sep}...{os.sep}{last}"
    else:
        return f"...{os.sep}{last}"

def clean_windows_path(path: Union[str, Path]) -> str:
    """Convert path to Windows format for Excel"""
    if not path:
        return ""
    return str(Path(path)).replace('/', '\\')

# ============================================================================
# Validation Functions
# ============================================================================

def validate_client_name(name: str) -> Tuple[bool, str]:
    """
    Validate client name.
    
    Returns:
        Tuple[bool, str]: (is_valid, error_message)
    """
    if not name or not name.strip():
        return False, "Client name cannot be empty"
    
    if len(name) < 2:
        return False, "Client name too short"
    
    if len(name) > 100:
        return False, "Client name too long"
    
    # Check for invalid characters
    invalid_pattern = r'[<>:"/\\|?*]'
    if re.search(invalid_pattern, name):
        return False, "Client name contains invalid characters"
    
    return True, ""

def get_unique_path(base_path: Union[str, Path]) -> Path:
    """Get unique path by appending number if exists"""
    path = Path(base_path)
    
    if not path.exists():
        return path
    
    counter = 1
    while True:
        if path.is_file():
            new_path = path.parent / f"{path.stem}_{counter}{path.suffix}"
        else:
            new_path = path.parent / f"{path.name}_{counter}"
        
        if not new_path.exists():
            return new_path
        
        counter += 1
        if counter > 100:  # Safety limit
            raise ValueError(f"Cannot create unique path for {base_path}")

# ============================================================================
# Progress and Status Functions
# ============================================================================

class ProgressTracker:
    """Track and report progress"""
    
    def __init__(self, total: int, callback=None):
        self.total = total
        self.current = 0
        self.callback = callback
        self.start_time = time.time()
    
    def update(self, increment: int = 1, message: str = ""):
        """Update progress"""
        self.current += increment
        progress = (self.current / self.total * 100) if self.total > 0 else 0
        
        if self.callback:
            self.callback(self.current, self.total, progress, message)
    
    def get_elapsed_time(self) -> float:
        """Get elapsed time in seconds"""
        return time.time() - self.start_time
    
    def get_eta(self) -> str:
        """Get estimated time remaining"""
        if self.current == 0:
            return "Unknown"
        
        elapsed = self.get_elapsed_time()
        rate = self.current / elapsed
        remaining = self.total - self.current
        
        if rate > 0:
            eta_seconds = remaining / rate
            return format_duration(eta_seconds)
        
        return "Unknown"

def format_duration(seconds: float) -> str:
    """Format duration for display"""
    if seconds < 60:
        return f"{int(seconds)}s"
    elif seconds < 3600:
        minutes = int(seconds / 60)
        secs = int(seconds % 60)
        return f"{minutes}m {secs}s"
    else:
        hours = int(seconds / 3600)
        minutes = int((seconds % 3600) / 60)
        return f"{hours}h {minutes}m"

# ============================================================================
# File Hash Functions
# ============================================================================

def calculate_file_hash(file_path: Union[str, Path], 
                       algorithm: str = 'md5') -> Optional[str]:
    """Calculate file hash for verification"""
    try:
        path = Path(file_path)
        if not path.exists():
            return None
        
        hash_func = hashlib.new(algorithm)
        
        with open(path, 'rb') as f:
            for chunk in iter(lambda: f.read(8192), b''):
                hash_func.update(chunk)
        
        return hash_func.hexdigest()
        
    except Exception as e:
        logger.error(f"Error calculating hash for {file_path}: {e}")
        return None

# ============================================================================
# Excel Path Extraction
# ============================================================================

def extract_filename_without_extension(filename: str) -> str:
    """Extract filename without extension for Excel formulas"""
    return Path(filename).stem

def get_relative_excel_path(from_path: Union[str, Path], 
                           to_path: Union[str, Path]) -> str:
    """Get relative path for Excel formulas"""
    try:
        from_p = Path(from_path)
        to_p = Path(to_path)
        
        # Try to get relative path
        try:
            rel_path = to_p.relative_to(from_p.parent)
            return clean_windows_path(rel_path)
        except ValueError:
            # If not relative, return absolute path
            return clean_windows_path(to_p)
            
    except Exception as e:
        logger.error(f"Error getting relative path: {e}")
        return clean_windows_path(to_path)
    
# ============================================================================
# NEW: Short Path and State Code Functions
# ============================================================================

def get_short_path(file_path: Union[str, Path]) -> str:
    """Get 8.3 short path for long paths (Windows only)"""
    try:
        if platform.system() == 'Windows':
            import win32api
            return win32api.GetShortPathName(str(file_path))
        else:
            return str(file_path)
    except Exception as e:
        logger.debug(f"Could not get short path for {file_path}: {e}")
        return str(file_path)

def get_state_code(state_name: str) -> str:
    """Convert full state name to state code"""
    if not state_name:
        return state_name
    
    # Import here to avoid circular import
    from .constants import STATE_CODE_MAPPING
    
    # Clean and normalize
    clean_state = state_name.strip().lower()
    
    # Check mapping
    state_code = STATE_CODE_MAPPING.get(clean_state)
    if state_code:
        return state_code
    
    # Fallback: create abbreviation from first letters
    words = clean_state.split()
    if len(words) >= 2:
        fallback = ''.join(word[0].upper() for word in words[:2])
        logger.warning(f"Unknown state '{state_name}', using fallback: {fallback}")
        return fallback
    else:
        # Single word, take first 2-3 characters
        fallback = state_name[:3].upper()
        logger.warning(f"Unknown state '{state_name}', using fallback: {fallback}")
        return fallback

def create_client_state_key(client: str, state: str, max_length: int = 35) -> str:
    """Create Client-State key with length limit"""
    # Apply business abbreviations to client (case-insensitive)
    import re
    clean_client = re.sub(r'\bPrivate\b', 'Pvt', client, flags=re.IGNORECASE)
    clean_client = re.sub(r'\bLimited\b', 'Ltd', clean_client, flags=re.IGNORECASE)
    
    # Get state code
    state_code = get_state_code(state)
    
    # Create key
    key = f"{clean_client}-{state_code}"
    
    # Apply length limit
    if len(key) > max_length:
        # Calculate available space for client name
        available = max_length - len(state_code) - 1  # -1 for hyphen
        if available > 5:  # Minimum reasonable client name length
            clean_client = clean_client[:available]
            key = f"{clean_client}-{state_code}"
        else:
            # Truncate the whole thing
            key = key[:max_length]
    
    return key