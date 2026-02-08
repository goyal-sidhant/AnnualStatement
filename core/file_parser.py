"""
File Parser for GST File Organizer v3.0
Robust pattern matching and file categorization.
Fixed with proper Union import.
"""

import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Any, Union
from collections import defaultdict

from utils.constants import FILE_PATTERNS, EXPECTED_FILE_TYPES
from utils.helpers import (
    find_excel_files, validate_client_name, format_size, get_state_code
)

logger = logging.getLogger(__name__)


class FileParser:
    """
    Parse and categorize GST files based on naming patterns.
    """
    
    def __init__(self):
        self.patterns = FILE_PATTERNS
        self.expected_types = EXPECTED_FILE_TYPES
        self.progress_callback = None
        self.reset()
    
    def reset(self):
        """Reset parser state"""
        self.scanned_files = {}
        self.client_data = defaultdict(lambda: {
            'client': '',
            'state': '',
            'files': defaultdict(list),
            'missing_files': [],
            'extra_files': [],
            'status': 'Unknown',
            'total_size': 0,
            'file_count': 0
        })
        self.variations = []
        self.errors = []

    def _update_progress(self, current: int, total: int, message: str):
        """Safely update progress if callback exists"""
        if self.progress_callback:
            try:
                self.progress_callback(current, total, message)
            except Exception as e:
                logger.warning(f"Progress callback failed: {e}")
    
    def parse_filename(self, filename: str) -> Dict[str, Any]:
        """
        Parse filename according to GST patterns.
        
        Args:
            filename: Name of file to parse
            
        Returns:
            Dict containing parsed information
        """
        result = {
            'filename': filename,
            'parsed': False,
            'type': 'Unknown',
            'category': 'Unknown',
            'folder': 'Unknown',
            'client': '',
            'state': '',
            'metadata': {}
        }
        
        # Try each pattern
        for pattern_name, pattern_info in self.patterns.items():
            match = pattern_info['pattern'].match(filename)
            
            if match:
                groups = match.groups()
                
                # Update result with pattern info
                result.update({
                    'parsed': True,
                    'pattern': pattern_name,
                    'type': pattern_info['type'],
                    'category': pattern_info['category'],
                    'folder': pattern_info['folder']
                })
                
                # Extract group data based on pattern
                if pattern_info['groups']:
                    for i, group_name in enumerate(pattern_info['groups']):
                        if i < len(groups) and groups[i]:
                            if group_name == 'client':
                                result['client'] = groups[i].strip()
                            elif group_name == 'state':
                                result['state'] = groups[i].strip()
                            else:
                                result['metadata'][group_name] = groups[i].strip()
                
                # Validate extracted data
                if result['client'] and result['state']:
                    client_valid, client_error = validate_client_name(result['client'])
                    state_valid, state_error = validate_client_name(result['state'])
                    
                    if not client_valid:
                        result['warnings'] = result.get('warnings', [])
                        result['warnings'].append(f"Client name issue: {client_error}")
                    
                    if not state_valid:
                        result['warnings'] = result.get('warnings', [])
                        result['warnings'].append(f"State name issue: {state_error}")
                
                logger.debug(f"Parsed {filename} as {pattern_name}")
                break
        
        if not result['parsed']:
            logger.warning(f"Could not parse filename: {filename}")
            result['error'] = 'No matching pattern found'
        
        return result
    
    def scan_folder(self, folder_path: Union[str, Path], 
                   progress_callback=None) -> Tuple[Dict, Dict, List]:
        """
        Scan folder for GST files and categorize them.
        
        Args:
            folder_path: Path to folder to scan
            progress_callback: Optional callback for progress updates
            
        Returns:
            Tuple of (scanned_files, client_data, variations)
        """
        self.reset()
        self.progress_callback = progress_callback
        try:
            folder = Path(folder_path)
            if not folder.exists():
                raise FileNotFoundError(f"Folder not found: {folder}")
            if not folder.is_dir():
                raise ValueError(f"Not a directory: {folder}")
            # Find all Excel files (fast glob, no I/O per file)
            self._update_progress(0, 0, "Finding Excel files...")
            excel_files = find_excel_files(folder)
            if not excel_files:
                logger.warning(f"No Excel files found in {folder}")
                self._update_progress(0, 0, "No Excel files found")
                return self.scanned_files, dict(self.client_data), self.variations
            total_files = len(excel_files)
            logger.info(f"Found {total_files} Excel files to process")
            self._update_progress(0, total_files, f"Found {total_files} files, processing...")
            # Process each file
            for idx, file_path in enumerate(excel_files):
                self._update_progress(idx + 1, total_files, f"[{idx+1}/{total_files}] {file_path.name}")
                self._process_file(file_path, folder)
            # Analyze completeness
            self._analyze_client_completeness()
            # Log summary
            logger.info(f"Scan complete: {len(self.client_data)} clients, "
                       f"{len(self.variations)} variations")
            return self.scanned_files, dict(self.client_data), self.variations
        except Exception as e:
            logger.error(f"Error scanning folder: {e}")
            self.errors.append(str(e))
            raise
    
    def _process_file(self, file_path: Path, base_folder: Path):
        """Process a single file with minimal I/O (one stat + one open)"""
        try:
            # Single stat() call to get size and timestamps
            try:
                stat = file_path.stat()
            except OSError as e:
                logger.warning(f"Cannot stat {file_path.name}: {e}")
                return

            size = stat.st_size

            # Skip files too small to be valid Excel (<1KB)
            if size < 1024:
                self.variations.append({
                    'filename': file_path.name,
                    'path': str(file_path),
                    'issue': 'File too small to be valid Excel',
                    'size': size
                })
                return

            # Validate Excel signature (single open+read, no extra stat/exists)
            try:
                with open(file_path, 'rb') as f:
                    header = f.read(8)
                xlsx_sig = b'PK\x03\x04'
                xls_sig = b'\xd0\xcf\x11\xe0'
                if not (header.startswith(xlsx_sig) or header.startswith(xls_sig)):
                    self.variations.append({
                        'filename': file_path.name,
                        'path': str(file_path),
                        'issue': 'Invalid or corrupted Excel file',
                        'size': size
                    })
                    return
            except OSError as e:
                logger.warning(f"Cannot read {file_path.name}: {e}")
                self.variations.append({
                    'filename': file_path.name,
                    'path': str(file_path),
                    'issue': f'Cannot read file: {e}',
                    'size': size
                })
                return

            # Build file info from the stat we already have
            file_info = {
                'name': file_path.name,
                'path': str(file_path),
                'size': size,
                'size_formatted': format_size(size),
                'modified': datetime.fromtimestamp(stat.st_mtime),
                'created': datetime.fromtimestamp(stat.st_ctime),
                'extension': file_path.suffix.lower(),
            }

            # Parse filename
            parsed = self.parse_filename(file_path.name)

            # Combine information
            file_data = {
                **file_info,
                **parsed,
                'full_path': str(file_path),
                'relative_path': str(file_path.relative_to(base_folder))
            }

            # Store in scanned files
            self.scanned_files[file_path.name] = file_data

            # Process based on parsing result
            if parsed['parsed'] and parsed['client'] and parsed['state']:
                self._add_to_client_data(file_data)
            else:
                self.variations.append({
                    'filename': file_path.name,
                    'path': str(file_path),
                    'issue': parsed.get('error', 'Could not parse filename'),
                    'size': size,
                    'pattern_expected': self._suggest_pattern(file_path.name)
                })

        except Exception as e:
            logger.error(f"Error processing file {file_path}: {e}")
            self.errors.append(f"Error processing {file_path.name}: {str(e)}")
    
    def _add_to_client_data(self, file_data: Dict[str, Any]):
        """Add parsed file to client data structure using state codes"""
        # Convert state to code for consistent keying
        state_code = get_state_code(file_data['state'])
        client_key = f"{file_data['client']}-{state_code}"
        
        # Update client info
        client = self.client_data[client_key]
        client['client'] = file_data['client']
        client['state'] = file_data['state']  # Keep original state name
        client['state_code'] = state_code     # Add state code
        
        # Add file to appropriate type
        file_type = file_data['type']
        client['files'][file_type].append(file_data)
        
        # Update statistics
        client['total_size'] += file_data.get('size', 0)
        client['file_count'] += 1
        
        logger.debug(f"Added {file_data['name']} to {client_key} as {file_type}")
    
    def _analyze_client_completeness(self):
        """Analyze which files are missing for each client"""
        for client_key, client_info in self.client_data.items():
            found_types = set(client_info['files'].keys())
            expected_types = set(self.expected_types)
            
            # Find missing files
            missing = expected_types - found_types
            client_info['missing_files'] = sorted(list(missing))
            
            # Find duplicate/extra files
            extras = []
            for file_type, files in client_info['files'].items():
                if len(files) > 1:
                    # GSTR-3B can have multiple files (monthly filings - unlimited)
                    if file_type == 'GSTR-3B Export':
                        continue  # Always allow multiple GSTR-3B files
                    extras.append(f"Multiple {file_type} ({len(files)} files)")
            
            client_info['extra_files'] = extras
            
            # Calculate completeness
            completeness = (len(found_types) / len(expected_types) * 100) if expected_types else 0
            client_info['completeness'] = round(completeness, 1)
            
            # Determine status
            if not missing and not extras:
                client_info['status'] = 'Complete'
                client_info['status_icon'] = '✅'
            elif missing:
                client_info['status'] = f'Missing {len(missing)} files'
                client_info['status_icon'] = '⚠️'
            else:
                client_info['status'] = 'Has duplicates'
                client_info['status_icon'] = '⚠️'
    
    def _suggest_pattern(self, filename: str) -> str:
        """Suggest correct pattern for filename"""
        # Simple pattern matching to suggest correct format
        suggestions = []
        
        if 'gstr' in filename.lower() and '2b' in filename.lower():
            suggestions.append("GSTR-2B-Reco-ClientName-State-Period.xlsx")
        elif 'ims' in filename.lower():
            suggestions.append("ImsReco-ClientName-State-DDMMYYYY.xlsx")
        elif 'gstr' in filename.lower() and '3b' in filename.lower():
            suggestions.append("GSTR3B-ClientName-State-Month.xlsx")
        elif 'sales' in filename.lower() and 'reco' in filename.lower():
            suggestions.append("SalesReco-ClientName-State-Period.xlsx")
        elif 'sales' in filename.lower():
            suggestions.append("Sales-ClientName-State-StartMonth-EndMonth.xlsx")
        elif 'annual' in filename.lower():
            suggestions.append("AnnualReport-ClientName-State-Year.xlsx")
        
        return " or ".join(suggestions) if suggestions else "Check file naming convention"
    
    def get_statistics(self) -> Dict[str, Any]:
        """Get scanning statistics"""
        total_files = len(self.scanned_files)
        parsed_files = sum(1 for f in self.scanned_files.values() if f.get('parsed'))
        total_size = sum(f.get('size', 0) for f in self.scanned_files.values())
        
        complete_clients = sum(1 for c in self.client_data.values() 
                             if c['status'] == 'Complete')
        
        file_type_dist = defaultdict(int)
        for client in self.client_data.values():
            for file_type, files in client['files'].items():
                file_type_dist[file_type] += len(files)
        
        return {
            'total_files': total_files,
            'parsed_files': parsed_files,
            'unparsed_files': total_files - parsed_files,
            'total_clients': len(self.client_data),
            'complete_clients': complete_clients,
            'incomplete_clients': len(self.client_data) - complete_clients,
            'variations': len(self.variations),
            'total_size': total_size,
            'total_size_formatted': format_size(total_size),
            'file_type_distribution': dict(file_type_dist),
            'parsing_rate': (parsed_files / total_files * 100) if total_files > 0 else 0,
            'completion_rate': (complete_clients / len(self.client_data) * 100) 
                             if self.client_data else 0
        }
    
    def validate_results(self) -> List[Dict[str, str]]:
        """Validate parsing results"""
        issues = []
        
        for client_key, client_info in self.client_data.items():
            # Check for empty client/state names
            if not client_info['client'].strip():
                issues.append({
                    'type': 'Empty Client Name',
                    'client': client_key,
                    'severity': 'Error'
                })
            
            if not client_info['state'].strip():
                issues.append({
                    'type': 'Empty State Name',
                    'client': client_key,
                    'severity': 'Error'
                })
            
            # Check for critical missing files
            critical_types = ['Annual Report', 'GSTR-3B Export']
            missing_critical = [f for f in critical_types 
                              if f in client_info['missing_files']]
            
            if missing_critical:
                issues.append({
                    'type': 'Missing Critical Files',
                    'client': client_key,
                    'severity': 'Warning',
                    'details': ', '.join(missing_critical)
                })
        
        return issues