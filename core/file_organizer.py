"""
File Organizer for GST File Organizer v3.0
Handles folder creation and file organization.
"""

import logging
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple, Union
from datetime import datetime

from utils.constants import FOLDER_STRUCTURE, PROCESSING_MODES
from utils.helpers import (
    ensure_path_exists, safe_copy_file, create_backup,
    get_timestamp, sanitize_filename, safe_path_join,
    get_state_code, create_client_state_key
)

logger = logging.getLogger(__name__)


class FileOrganizer:
    """
    Organize GST files into structured folders.
    """
    
    def __init__(self, target_folder: Union[str, Path], 
                    processing_mode: str = 'fresh',
                    include_client_name: bool = False,
                    client_folder_settings: Dict[str, bool] = None,
                    client_name_max_length: int = 35):
        """
        Initialize file organizer.
        
        Args:
            target_folder: Base target folder
            processing_mode: One of 'fresh', 'rerun', 'resume'
        """
        self.target_folder = Path(target_folder)
        self.processing_mode = processing_mode
        self.include_client_name = include_client_name
        self.client_folder_settings = client_folder_settings or {}
        self.client_name_max_length = client_name_max_length if client_name_max_length > 0 else 35
        self.timestamp = get_timestamp()
        
        # Tracking
        self.created_folders = []
        self.copied_files = []
        self.errors = []
        
        # Validate inputs
        self._validate_inputs()
        
        logger.info(f"FileOrganizer initialized: target={self.target_folder}, "
                   f"mode={processing_mode}")
    
    def _validate_inputs(self):
        """Validate initialization inputs"""
        if self.processing_mode not in PROCESSING_MODES:
            raise ValueError(f"Invalid processing mode: {self.processing_mode}")
        
        # Ensure target folder exists
        try:
            ensure_path_exists(self.target_folder)
        except Exception as e:
            raise ValueError(f"Cannot create target folder: {e}")
    
    def create_client_structure(self, client_info: Dict[str, Any]) -> Dict[str, Path]:
        """
        Create folder structure for a client.
        
        Args:
            client_info: Client information dictionary
            
        Returns:
            Dict mapping folder names to paths
        """
        try:
            state_code = get_state_code(client_info['state'])
            client_state_key = create_client_state_key(
                client_info['client'], 
                client_info['state'], 
                max_length=self.client_name_max_length
            )

            # Extract sanitized parts
            client = sanitize_filename(client_info['client'])
            state = state_code  # Use state code instead of full name
            
            # Get or create Level 1 folder
            level1_folder = self._get_level1_folder()
            
            # Create Level 2 (client) folder
            level2_name = client_state_key  # Use the properly formatted client-state key
            level2_folder = safe_path_join(level1_folder, level2_name)
            ensure_path_exists(level2_folder)
            self.created_folders.append(level2_folder)
            
            # Create version folder
            version_name = FOLDER_STRUCTURE['version'].format(
                timestamp=self.timestamp
            )
            version_folder = safe_path_join(level2_folder, version_name)
            ensure_path_exists(version_folder)
            self.created_folders.append(version_folder)
            
            # Create category folders
            folders = {
                'level1': level1_folder,
                'level2': level2_folder,
                'version': version_folder
            }
            
            for folder_key, folder_template in FOLDER_STRUCTURE['categories'].items():
                # Format folder name
                folder_name = folder_template.format(client=client)
                
                # Check if we should include client name
                client_key = f"{client_info['client']}-{state_code}"
                
                # Global setting overrides individual settings
                if self.include_client_name:
                    # Global checkbox is ON - include for all
                    include_name = True
                else:
                    # Global checkbox is OFF - check individual setting
                    include_name = self.client_folder_settings.get(client_key, False)
                
                # Apply the decision
                if not include_name:
                    folder_name = folder_name.replace(f' ({client})', '')
                
                folder_path = safe_path_join(version_folder, folder_name)
                ensure_path_exists(folder_path)
                folders[folder_key] = folder_path
                self.created_folders.append(folder_path)
            
            logger.info(f"Created folder structure for {client}-{state}")
            return folders
            
        except Exception as e:
            error_msg = f"Error creating folders for {client_info['client']}: {e}"
            logger.error(error_msg)
            self.errors.append(error_msg)
            raise

    def _get_level1_folder(self) -> Path:
        """Get or create Level 1 folder based on processing mode"""
        if self.processing_mode == 'fresh':
            # Create new timestamped folder
            folder_name = FOLDER_STRUCTURE['level1'].format(
                timestamp=self.timestamp
            )
            level1_folder = safe_path_join(self.target_folder, folder_name)
            ensure_path_exists(level1_folder)
            self.created_folders.append(level1_folder)
            return level1_folder
        
        else:  # rerun or resume
            # Find latest Annual Statement folder
            annual_folders = [
                f for f in self.target_folder.iterdir()
                if f.is_dir() and f.name.startswith("Annual Statement-")
            ]
            
            if annual_folders:
                # Use the most recent one
                latest = max(annual_folders, key=lambda f: f.stat().st_mtime)
                logger.info(f"Using existing folder: {latest}")
                return latest
            else:
                # No existing folder, create new one
                logger.warning("No existing Annual Statement folder found, creating new")
                folder_name = FOLDER_STRUCTURE['level1'].format(
                    timestamp=self.timestamp
                )
                level1_folder = safe_path_join(self.target_folder, folder_name)
                ensure_path_exists(level1_folder)
                self.created_folders.append(level1_folder)
                return level1_folder
    
    def organize_files(self, client_info: Dict[str, Any], 
                      folders: Dict[str, Path],
                      progress_callback=None) -> List[Dict[str, Any]]:
        """
        Organize files for a client into appropriate folders.
        FIXED: Updates client_info with sanitized filenames after copying.
        
        Args:
            client_info: Client information with files
            folders: Dictionary of created folders
            progress_callback: Optional progress callback
            
        Returns:
            List of file operation results
        """
        results = []
        all_files = []
        
        # Collect all files
        for file_type, files in client_info['files'].items():
            for file_data in files:
                all_files.append((file_type, file_data))
        
        total_files = len(all_files)
        
        # Process each file and update client_info with new names
        for idx, (file_type, file_data) in enumerate(all_files):
            if progress_callback:
                progress_callback(idx + 1, total_files, 
                                f"Copying {file_data['name']}")
            
            result = self._organize_single_file(file_type, file_data, folders)
            results.append(result)
            self.copied_files.append(result)
            
            # IMPORTANT: Update the file data with the sanitized filename
            if result['status'] == 'Success':
                # Extract the sanitized filename from the destination path
                dest_path = Path(result['destination'])
                sanitized_name = dest_path.name
                
                # Update the file data in client_info with the sanitized name
                # This ensures Excel handler gets the correct filename
                for i, f in enumerate(client_info['files'][file_type]):
                    if f['full_path'] == file_data['full_path']:
                        client_info['files'][file_type][i]['name'] = sanitized_name
                        client_info['files'][file_type][i]['sanitized_name'] = sanitized_name
                        logger.debug(f"Updated filename from '{file_data['name']}' to '{sanitized_name}'")
                        break
        
        return results
    
    def _organize_single_file(self, file_type: str, 
                            file_data: Dict[str, Any],
                            folders: Dict[str, Path]) -> Dict[str, Any]:
        """Organize a single file"""
        result = {
            'filename': file_data['name'],
            'file_type': file_type,
            'source': file_data['full_path'],
            'destination': '',
            'status': 'Failed',
            'error': None
        }
        
        try:
            # Determine destination folder
            dest_folder = self._get_destination_folder(file_type, folders)
            if not dest_folder:
                result['error'] = f"No destination folder for type: {file_type}"
                return result
            
            # Prepare destination path with sanitized filename
            safe_filename = sanitize_filename(file_data['name'])
            dest_path = safe_path_join(dest_folder, safe_filename)
            result['destination'] = str(dest_path)
            result['sanitized_filename'] = safe_filename
            
            # Handle existing file
            if dest_path.exists():
                if self.processing_mode == 'fresh':
                    # Create backup
                    backup_path = create_backup(dest_path)
                    if backup_path:
                        result['backup'] = str(backup_path)
                elif self.processing_mode == 'resume':
                    # Skip if already exists
                    result['status'] = 'Skipped'
                    result['error'] = 'File already exists'
                    return result
            
            # Copy file
            if safe_copy_file(file_data['full_path'], dest_path, verify=True):
                result['status'] = 'Success'
                logger.info(f"Copied {file_data['name']} to {dest_folder.name}")
            else:
                result['error'] = 'Copy operation failed'
                
        except Exception as e:
            result['error'] = str(e)
            logger.error(f"Error organizing {file_data['name']}: {e}")
        
        return result
    
    def _get_destination_folder(self, file_type: str, 
                              folders: Dict[str, Path]) -> Optional[Path]:
        """Get destination folder for file type"""
        type_to_folder = {
            'GSTR-3B Export': 'gstr3b',
            'GSTR-2B Reco': 'itc',
            'IMS Reco': 'itc',
            'Sales': 'sales',
            'Sales Reco': 'sales',
            'Annual Report': 'version'  # Goes in version folder
        }
        
        folder_key = type_to_folder.get(file_type)
        return folders.get(folder_key) if folder_key else None
    
    def create_organization_report(self, folders: Dict[str, Path],
                                 client_info: Dict[str, Any]) -> Optional[Path]:
        """
        Create a text report of the organization process.
        
        Args:
            folders: Created folders
            client_info: Client information
            
        Returns:
            Path to report file or None
        """
        try:
            report_name = f"_Organization_Report_{client_info['client']}_{client_info['state']}.txt"
            report_path = safe_path_join(folders['version'], report_name)
            
            # Generate report content
            content = f"""GST FILE ORGANIZATION REPORT
{"=" * 50}
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Client: {client_info['client']}
State: {client_info['state']}
Processing Mode: {self.processing_mode}

FOLDER STRUCTURE:
{"=" * 50}
"""
            
            # Add folder paths
            for folder_type, folder_path in folders.items():
                if folder_type in ['level1', 'level2']:
                    continue
                rel_path = folder_path.relative_to(folders['level1'])
                content += f"{folder_type.upper()}: {rel_path}\n"
            
            # Add file summary
            content += f"\n\nFILES ORGANIZED:\n{'=' * 50}\n"
            
            for file_type, files in client_info['files'].items():
                content += f"\n{file_type} ({len(files)} files):\n"
                for file_data in files:
                    # Use sanitized name if available, otherwise original
                    filename = file_data.get('sanitized_name', file_data['name'])
                    content += f"  - {filename}\n"
            
            # Add statistics
            successful = sum(1 for f in self.copied_files 
                           if f.get('status') == 'Success')
            content += f"\n\nSTATISTICS:\n{'=' * 50}\n"
            content += f"Total Files: {client_info['file_count']}\n"
            content += f"Successfully Copied: {successful}\n"
            content += f"Folders Created: {len(self.created_folders)}\n"
            
            # Write report
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            logger.info(f"Created organization report: {report_path}")
            return report_path
            
        except Exception as e:
            logger.error(f"Error creating organization report: {e}")
            return None
    
    def cleanup_on_error(self, folders: Dict[str, Path]):
        """Clean up created folders on error"""
        if self.processing_mode != 'fresh':
            return  # Only cleanup for fresh runs
        
        try:
            # Remove version folder if empty
            if 'version' in folders:
                version_folder = folders['version']
                if self._is_folder_empty(version_folder):
                    import shutil
                    shutil.rmtree(version_folder, ignore_errors=True)
                    logger.info(f"Cleaned up empty folder: {version_folder}")
                    
        except Exception as e:
            logger.error(f"Error during cleanup: {e}")
    
    def _is_folder_empty(self, folder: Path) -> bool:
        """Check if folder is effectively empty"""
        try:
            # Check for any substantial files
            for item in folder.rglob('*'):
                if item.is_file() and not item.name.startswith('_'):
                    return False
            return True
        except Exception:
            return False
    
    def get_summary(self) -> Dict[str, Any]:
        """Get organization summary"""
        successful_copies = sum(1 for f in self.copied_files 
                              if f['status'] == 'Success')
        failed_copies = sum(1 for f in self.copied_files 
                          if f['status'] == 'Failed')
        
        return {
            'timestamp': self.timestamp,
            'processing_mode': self.processing_mode,
            'target_folder': str(self.target_folder),
            'folders_created': len(set(self.created_folders)),
            'files_processed': len(self.copied_files),
            'successful_copies': successful_copies,
            'failed_copies': failed_copies,
            'skipped_files': len(self.copied_files) - successful_copies - failed_copies,
            'errors': len(self.errors)
        }