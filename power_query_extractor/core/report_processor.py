"""
Report processing with Power Query refresh
"""

import os
import time
import shutil
import logging
from pathlib import Path
from datetime import datetime
import pythoncom
import win32com.client
import win32api
import win32con
import win32gui
from ..config.cell_mappings import CELL_MAPPINGS

logger = logging.getLogger(__name__)


class ReportProcessor:
    """Process Excel reports with Power Query refresh"""

    def __init__(self):
        self.excel_instance = None
        self.log_callback = None
        self.wait_time = 10  # Default wait time in seconds
        self.suffix_pattern = "_Refreshed_{timestamp}"  # Default suffix pattern
    
    def _log(self, message, level='info'):
        """Log to both file and GUI"""
        # Log to file
        logger.info(message)
        
        # Log to GUI if callback provided
        if self.log_callback:
            self.log_callback(f"  {message}", level)
    
    def process_client(self, client_data, log_callback=None):
        """Process selected reports for a client"""
        self.log_callback = log_callback  # Store the callback

        results = {
            'client': client_data['name'],
            'timestamp': datetime.now(),
            'itc': None,
            'sales': None
        }

        # Process ITC if selected
        if client_data.get('process_itc', True):  # Default to True for backward compatibility
            results['itc'] = self._process_report(client_data, 'ITC')
        else:
            results['itc'] = {
                'status': {'success': False, 'error': 'Not selected for processing'},
                'data': {},
                'extracted_values': {}
            }

        # Process Sales if selected
        if client_data.get('process_sales', True):  # Default to True for backward compatibility
            results['sales'] = self._process_report(client_data, 'Sales')
        else:
            results['sales'] = {
                'status': {'success': False, 'error': 'Not selected for processing'},
                'data': {},
                'extracted_values': {}
            }

        return results
    
    def _process_report(self, client_data, report_type):
        """Process a single report"""
        try:
            # Find report file
            pattern = f"{report_type}_Report_*.xlsx"
            report_files = list(client_data['latest_version'].glob(pattern))
            
            if not report_files:
                return {
                    'status': {
                        'success': False,
                        'error': f"No {report_type} report found"
                    },
                    'data': {},
                    'extracted_values': {}
                }
            
            original_file = report_files[0]

            # Create refreshed copy with custom suffix
            timestamp = datetime.now().strftime("%d%m%y_%H%M")
            suffix = self.suffix_pattern.replace("{timestamp}", timestamp)
            refreshed_file = original_file.parent / f"{original_file.stem}{suffix}.xlsx"
            
            # Check if we should skip refresh
            from ..config.cell_mappings import EXTRACTION_OPTIONS
            skip_refresh = EXTRACTION_OPTIONS.get('skip_refresh', False)
            
            if skip_refresh:
                # Look for existing refreshed file
                existing_refreshed = list(original_file.parent.glob(f"{original_file.stem}_Refreshed_*.xlsx"))
                if existing_refreshed:
                    # Use the most recent refreshed file
                    refreshed_file = max(existing_refreshed, key=lambda x: x.stat().st_mtime)
                    self._log(f"Skip refresh enabled - using existing: {refreshed_file.name}")
                    refresh_status = {
                        'success': True,
                        'error': None,
                        'validation_message': 'Skipped - Using existing refreshed file'
                    }
                else:
                    # No refreshed file exists, must refresh
                    self._log("No existing refreshed file found - will refresh")
                    skip_refresh = False
            
            # Only refresh if not skipping
            if not skip_refresh:
                logger.info(f"Creating copy: {refreshed_file.name}")
                shutil.copy2(original_file, refreshed_file)
                
                # Refresh Power Query
                try:
                    refresh_status = self._refresh_power_query_simple(refreshed_file)
                except Exception as e:
                    if "Open method of Workbooks class failed" in str(e) or "-2147352567" in str(e):
                        self._log(f"⚠️ Cannot open file for refresh: {str(e)[:100]}")
                        self._log("Proceeding with extraction without refresh...")
                        # Set refresh as failed but continue
                        refresh_status = {
                            'success': False,
                            'error': f'Could not open for refresh: {str(e)[:100]}',
                            'validation_message': 'Skipped - File could not be opened'
                        }
                    else:
                        # For other errors, re-raise
                        raise
            
            # Wait for file to be fully released
            self._log("Waiting for file to be released...")
            time.sleep(3)  # Give Excel time to fully release the file
            
            # Extract data regardless of refresh status
            extracted_data = self._extract_data(refreshed_file, report_type)
            
            return {
                'status': refresh_status,
                'data': extracted_data,
                'extracted_values': extracted_data,
                'original_file': str(original_file),
                'refreshed_file': str(refreshed_file)
            }
            
        except Exception as e:
            logger.error(f"Error processing {report_type} report: {e}", exc_info=True)
            return {
                'status': {
                    'success': False,
                    'error': str(e)
                },
                'data': {},
                'extracted_values': {}
            }
    
    def _refresh_power_query_simple(self, excel_file):
        """Simple refresh approach with proper error handling"""
        excel = None
        wb = None
        
        try:
            pythoncom.CoInitialize()
            
            self._log("Creating Excel instance...")
            excel = win32com.client.Dispatch("Excel.Application")
            
            # Set visibility with retry logic
            self._log("Setting Excel properties...")
            max_retries = 5
            for attempt in range(max_retries):
                try:
                    excel.Visible = True  # Keep it visible to avoid issues
                    excel.DisplayAlerts = False
                    excel.WindowState = -4137  # xlMaximized
                    break  # Success, exit loop
                except Exception as e:
                    logger.warning(f"Attempt {attempt + 1} failed: {e}")
                    if attempt < max_retries - 1:
                        time.sleep(2)  # Wait 2 seconds before retry
                    else:
                        logger.error("Could not set Excel properties, continuing anyway...")
            
            self._log(f"Opening Excel file: {excel_file.name}")
            
            # Open workbook with retry logic
            for attempt in range(3):
                try:
                    wb = self._safe_open_workbook(excel, excel_file)
                    self._log("✓ Excel file opened successfully")
                    break
                except Exception as e:
                    if "Call was rejected by callee" in str(e):
                        logger.warning(f"Excel busy, retrying... ({attempt + 1}/3)")
                        time.sleep(3)
                    else:
                        raise
            
            # Wait for workbook to fully load
            self._log("Waiting for workbook to fully load...")
            for i in range(5):  # Increased from 3 to 5
                self._log(f"  Loading... {i+1}/5 seconds")
                time.sleep(1)
            
            # Try to activate window with error handling
            try:
                self._log("Activating Excel window...")
                shell = win32com.client.Dispatch("WScript.Shell")
                
                # Try multiple times to activate
                activated = False
                for attempt in range(3):
                    try:
                        shell.AppActivate(excel.Caption)
                        activated = True
                        break
                    except Exception:
                        logger.warning(f"Could not activate window, attempt {attempt + 1}/3")
                        time.sleep(1)
                
                if not activated:
                    logger.warning("Could not activate Excel window, continuing anyway...")
                
                time.sleep(1)
                
                self._log("Starting refresh sequence...")
                
                # Send keys with delays
                self._log("  → Sending Alt key...")
                shell.SendKeys("%a")
                time.sleep(1)  # Increased from 0.5
                self._log("  ✓ Alt key sent")
                
                self._log("  → Sending R key...")
                shell.SendKeys("r")
                time.sleep(1)  # Increased from 0.5
                self._log("  ✓ R key sent")
                
                self._log("  → Sending A key...")
                shell.SendKeys("a")
                self._log("  ✓ A key sent - Refresh All command executed")
            
            # Rest of your code remains same...
                
                self._log(f"Waiting for Power Query refresh to complete... ({self.wait_time} seconds)")

                # Use configurable wait time
                for i in range(self.wait_time):
                    self._log(f"  Waiting... {i+1}/{self.wait_time} seconds")
                    time.sleep(1)
                
                # Check for any error windows
                self._log("Checking for error dialogs...")
                error_detected = False
                error_message = ""
                
                # Quick check for error dialogs
                for i in range(5):  # Check 5 times over 5 seconds
                    self._log(f"  Error check {i+1}/5...")
                    error_window = self._find_error_window()
                    if error_window:
                        error_detected = True
                        error_message = self._get_window_text(error_window)
                        logger.warning(f"  ⚠️ Error detected: {error_message}")
                        self._close_error_window(error_window)
                        break
                    time.sleep(1)
                
                if not error_detected:
                    self._log("  ✓ No errors detected")
                
                # Additional wait time for complex queries
                self._log("Giving additional time for complex queries...")
                extended_wait = max(10, self.wait_time)  # At least 10 seconds for complex queries
                for i in range(extended_wait):
                    if i % 2 == 0:
                        self._log(f"  Extended wait... {i}/{extended_wait} seconds")
                    time.sleep(1)
                
                # Final error check
                self._log("Final error check...")
                if not error_detected:
                    error_window = self._find_error_window()
                    if error_window:
                        error_detected = True
                        error_message = self._get_window_text(error_window)
                        logger.warning(f"  ⚠️ Late error detected: {error_message}")
                        self._close_error_window(error_window)
                    else:
                        self._log("  ✓ No errors found")

                # Validate refresh before saving
                validation_success, validation_message = self._validate_refresh(wb)
                
                # Save the workbook
                self._log("Saving workbook...")
                wb.Save()
                self._log("✓ Workbook saved")
                time.sleep(2)

                # Update status based on validation
                if not validation_success:
                    error_detected = True
                    error_message = f"Refresh validation failed: {validation_message}"
                    self._log(f"⚠️ {error_message}")
                
                # Close workbook
                self._log("Closing workbook...")
                wb.Close(SaveChanges=True)
                self._log("✓ Workbook closed")
                time.sleep(1)
                
                if error_detected:
                    self._log(f"⚠️ Refresh completed with errors: {error_message}")
                else:
                    self._log("✅ Refresh completed successfully")
                
                return {
                    'success': not error_detected,
                    'error': error_message if error_detected else None,
                    'validation_message': validation_message
                }
                
            except Exception as e:
                logger.error(f"❌ Error during refresh: {e}")
                return {
                    'success': False,
                    'error': str(e)
                }

        except Exception as e:
            logger.error(f"❌ Excel error: {e}", exc_info=True)
            return {
                'success': False,
                'error': str(e)
            }
        finally:
            # Cleanup
            self._log("Cleaning up Excel instance...")
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                    self._log("  Workbook closed")
                except Exception:
                    pass

            if excel:
                try:
                    excel.Quit()
                    self._log("  Excel quit")
                except Exception:
                    pass

            try:
                pythoncom.CoUninitialize()
                self._log("  COM uninitialized")
            except Exception:
                pass

    def _find_error_window(self):
        """Find Excel error dialog window"""
        def callback(hwnd, windows):
            try:
                if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                    window_text = win32gui.GetWindowText(hwnd).lower()
                    
                    # Common error window indicators
                    error_indicators = [
                        'microsoft excel',
                        'error',
                        'problem',
                        'cannot',
                        'failed',
                        'unable',
                        'could not',
                        'exception',
                        'invalid',
                        'missing'
                    ]
                    
                    if any(indicator in window_text for indicator in error_indicators):
                        # Also check if it's a dialog
                        class_name = win32gui.GetClassName(hwnd)
                        if class_name in ['#32770', 'NUIDialog', 'bosa_sdm_msword']:
                            windows.append(hwnd)
                            logger.debug(f"Found potential error window: {window_text}")
            except Exception:
                pass
            return True

        windows = []
        try:
            win32gui.EnumWindows(callback, windows)
            return windows[0] if windows else None
        except Exception:
            return None
    
    def _get_window_text(self, hwnd):
        """Get text from error window"""
        try:
            # Get main window text
            main_text = win32gui.GetWindowText(hwnd)
            all_text = [main_text] if main_text else []
            
            # Get text from child windows (static text, labels)
            def get_text_callback(child_hwnd, text_list):
                try:
                    class_name = win32gui.GetClassName(child_hwnd)
                    # Look for static text controls
                    if 'static' in class_name.lower() or 'text' in class_name.lower():
                        child_text = win32gui.GetWindowText(child_hwnd)
                        if child_text and len(child_text) > 5:
                            text_list.append(child_text)
                except Exception:
                    pass
                return True
            
            win32gui.EnumChildWindows(hwnd, get_text_callback, all_text)
            
            # Combine texts
            combined_text = ' '.join(all_text)
            
            # Clean up and limit length
            if combined_text:
                # Remove extra whitespace
                combined_text = ' '.join(combined_text.split())
                return combined_text[:200]
            else:
                return "Power Query error occurred"
                
        except Exception as e:
            logger.debug(f"Error getting window text: {e}")
            return "Power Query error occurred"
    
    def _close_error_window(self, hwnd):
        """Close error dialog window"""
        try:
            logger.info("Attempting to close error dialog...")
            
            # First try to find and click OK/Close button
            clicked = False
            
            def find_button(child_hwnd, unused):
                nonlocal clicked
                try:
                    class_name = win32gui.GetClassName(child_hwnd).lower()
                    if 'button' in class_name:
                        text = win32gui.GetWindowText(child_hwnd).lower()
                        if any(btn in text for btn in ['ok', 'close', 'yes', 'continue']):
                            logger.info(f"Clicking button: {text}")
                            win32gui.SendMessage(child_hwnd, win32con.BM_CLICK, 0, 0)
                            clicked = True
                            return False  # Stop enumeration
                except Exception:
                    pass
                return True
            
            win32gui.EnumChildWindows(hwnd, find_button, None)
            
            # If no button clicked, send close message
            if not clicked:
                logger.info("No button found, sending close message")
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            
            time.sleep(0.5)
            
        except Exception as e:
            logger.debug(f"Error closing window: {e}")
    
    def _extract_data(self, excel_file, report_type):
        """Extract data from Excel file with detailed logging"""
        excel = None
        wb = None
        
        try:
            pythoncom.CoInitialize()
            
            self._log("Creating new Excel instance for data extraction...")
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # Can be hidden for extraction
            excel.DisplayAlerts = False
            self._log("✓ Excel instance created (hidden mode)")
            
            self._log(f"Opening file in read-only mode: {excel_file.name}")
            wb = self._safe_open_workbook(excel, excel_file)
            self._log("✓ File opened in read-only mode")
            
            # Get cell mappings
            mappings = CELL_MAPPINGS.get(report_type, {})
            extracted_data = {}
            
            self._log(f"Starting data extraction for {report_type} report")
            self._log(f"Number of sheets to process: {len(mappings)}")
            
            if not mappings:
                logger.warning(f"⚠️ No mappings defined for {report_type}")
                return extracted_data
            
            for sheet_name, cells in mappings.items():
                self._log(f"\nProcessing sheet: '{sheet_name}'")
                self._log(f"  Cells to extract: {len(cells)}")
                
                try:
                    # Try to access sheet
                    self._log(f"  Looking for sheet '{sheet_name}'...")
                    ws = None
                    for sheet in wb.Sheets:
                        if sheet.Name == sheet_name:
                            ws = sheet
                            self._log(f"  ✓ Sheet found")
                            break
                    
                    if not ws:
                        logger.warning(f"  ⚠️ Sheet '{sheet_name}' not found")
                        # Add None values for all fields in this sheet
                        for cell_ref, field_name in cells.items():
                            extracted_data[field_name] = None
                            self._log(f"    {field_name} = None (sheet not found)")
                        continue
                    

                    # Extract cell values
                    self._log(f"  Reading cell values...")
                    for cell_ref, field_name in cells.items():
                        try:
                            self._log(f"    Reading {cell_ref}...")
                            value = ws.Range(cell_ref).Value
                            extracted_data[field_name] = value
                            
                            # Also store by cell reference for calculations
                            extracted_data[cell_ref] = value
                            
                            self._log(f"    ✓ {field_name} ({cell_ref}) = {value}")
                        except Exception as e:
                            logger.warning(f"    ⚠️ Could not read {cell_ref}: {e}")
                            extracted_data[field_name] = None
                            extracted_data[cell_ref] = None
                            
                except Exception as e:
                    logger.error(f"❌ Error accessing sheet '{sheet_name}': {e}")
                    # Add None values for all fields in this sheet
                    for cell_ref, field_name in cells.items():
                        extracted_data[field_name] = None
                        self._log(f"  {field_name} = None (sheet error)")
            
            self._log(f"\n✅ Extraction complete. Total values extracted: {len(extracted_data)}")
            
            # Log summary of extracted values
            self._log("Extraction summary:")
            for field, value in extracted_data.items():
                if value is not None:
                    self._log(f"  ✓ {field}: {value}")
                else:
                    self._log(f"  ✗ {field}: None")
            
            return extracted_data
            
        except Exception as e:
            logger.error(f"❌ Extraction error: {e}", exc_info=True)
            return {}
        finally:
            self._log("\nCleaning up extraction Excel instance...")
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                    self._log("  ✓ Workbook closed")
                except Exception:
                    pass
            if excel:
                try:
                    excel.Quit()
                    self._log("  ✓ Excel quit")
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
                self._log("  ✓ COM uninitialized")
            except Exception:
                pass

    def _validate_refresh(self, wb):
        """Validate that Power Query refresh was successful by checking Info sheet"""
        try:
            self._log(f"Validating refresh results for report...")
            
            # Check if Info sheet exists
            info_sheet = None
            for sheet in wb.Sheets:
                if sheet.Name.lower() == 'info':
                    info_sheet = sheet
                    break
            
            if not info_sheet:
                self._log("⚠️ Info sheet not found")
                return False, "Info sheet not found"
            
            # Check cell BB2 for client name
            self._log("Checking cell BB2 in Info sheet...")
            cell_value = info_sheet.Range('BB2').Value
            
            # Check if cell is empty, None, 0, or just whitespace
            if cell_value is None:
                self._log("❌ Cell BB2 is empty (None)")
                return False, "Cell BB2 is empty"
            
            # Convert to string and check
            cell_str = str(cell_value).strip()
            
            if not cell_str:
                self._log("❌ Cell BB2 is blank")
                return False, "Cell BB2 is blank"
            
            if cell_str == '0' or cell_str == '0.0':
                self._log("❌ Cell BB2 contains 0")
                return False, "Cell BB2 contains 0"
            
            # Additional checks for error values
            error_values = ['#REF!', '#VALUE!', '#NAME?', '#NULL!', '#DIV/0!', '#N/A', '#NUM!', '#ERROR!']
            if cell_str.upper() in error_values:
                self._log(f"❌ Cell BB2 contains error: {cell_str}")
                return False, f"Cell BB2 contains error: {cell_str}"
            
            self._log(f"✅ Cell BB2 populated successfully: {cell_str}")
            return True, f"Validated - Client Name: {cell_str}"
            
        except Exception as e:
            self._log(f"❌ Validation error: {e}")
            return False, f"Validation error: {str(e)}"
        
    def _safe_open_workbook(self, excel, file_path):
        """Safely open workbook with detailed diagnostics"""
        file_path = Path(file_path).absolute()
        
        # Check if file exists
        if not file_path.exists():
            self._log(f"❌ File does not exist: {file_path}")
            raise FileNotFoundError(f"File not found: {file_path}")
        
        # Check if file is accessible
        try:
            with open(file_path, 'rb') as f:
                f.read(1)
            self._log(f"✓ File is accessible: {file_path.name}")
        except Exception as e:
            self._log(f"❌ Cannot access file: {e}")
            raise
        
        # Check if already open in Excel
        for wb in excel.Workbooks:
            if Path(wb.FullName).absolute() == file_path:
                self._log(f"📌 File already open, using existing workbook")
                return wb
        
        file_str = str(file_path)

        # CHECK PATH LENGTH AND HANDLE LONG PATHS
        path_length = len(file_str)
        self._log(f"Path length: {path_length} characters")
        
        if path_length > 218:  # Excel's typical limit
            self._log("⚠️ Long path detected, applying fixes...")
            
            # Fix 1: Try normalized path with backslashes
            file_str = file_str.replace('/', '\\')
            
            # Fix 2: Try Windows extended path format
            if not file_str.startswith('\\\\?\\'):
                if file_str.startswith('\\\\'):  # Network path
                    file_str = '\\\\?\\UNC\\' + file_str[2:]
                else:  # Local path
                    file_str = '\\\\?\\' + file_str
                self._log(f"Using extended path format")
        
        # First attempt - standard open
        try:
            self._log(f"Opening workbook: {file_path.name}")
            return excel.Workbooks.Open(file_str)
        except Exception as e:
            error_code = getattr(e, 'args', [None])[0] if hasattr(e, 'args') else None
            self._log(f"First attempt failed - Error code: {error_code}")
            self._log(f"Error details: {str(e)[:200]}")
        
        # Second attempt - close any workbook with same name
        try:
            self._log("Checking for workbooks with same name...")
            for wb in excel.Workbooks:
                if wb.Name == file_path.name:
                    self._log(f"Closing existing workbook: {wb.Name}")
                    wb.Close(SaveChanges=False)
            time.sleep(3)
            return excel.Workbooks.Open(file_str)
        except Exception as e:
            self._log(f"Second attempt failed: {str(e)[:100]}")
        
        # Third attempt - for long paths specifically
        try:
            self._log("Trying short path workaround...")
            time.sleep(2)
            
            # Try to get 8.3 short path name
            try:
                import win32api
                short_path = win32api.GetShortPathName(str(file_path))
                if short_path != str(file_path):
                    self._log(f"Using short path: ...{short_path[-50:]}")
                    return excel.Workbooks.Open(short_path)
            except Exception:
                self._log("Could not get short path")
            
            # Last resort - remove extended path prefix if we added it
            if file_str.startswith('\\\\?\\'):
                normal_path = str(file_path)
                self._log("Trying without extended path prefix...")
                return excel.Workbooks.Open(normal_path)
            
            # Try with all parameters as before
            return excel.Workbooks.Open(
                Filename=str(file_path),  # Use original path
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
                Notify=False
            )
        except Exception as e:
            self._log(f"❌ All attempts failed. Final error: {str(e)}")
            
            # Enhanced guidance for long paths
            if "-2147352567" in str(e):
                self._log("💡 This usually means:")
                self._log("   - File is open in another program")
                self._log("   - File is corrupted")
                self._log(f"   - Path is too long ({path_length} chars)")
                self._log("   - OneDrive/SharePoint sync issue")
                
                if path_length > 218:
                    self._log("📋 Path length is definitely an issue!")
                    self._log("   Consider enabling long path support in Windows")
            
            raise

    def cleanup(self):
        """Cleanup method"""
        # Nothing to cleanup with this approach as we create/destroy Excel each time
        pass