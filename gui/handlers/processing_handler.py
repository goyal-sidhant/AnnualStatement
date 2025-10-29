# gui/handlers/processing_handler.py
"""Processing operations handler"""

import os
import platform
import logging
import threading
from datetime import datetime
from pathlib import Path
from tkinter import messagebox
from utils.constants import MESSAGES
from utils.helpers import get_timestamp, sanitize_filename, get_state_code, ProgressTracker

logger = logging.getLogger(__name__)


class ProcessingHandler:
    """Handles file processing operations"""
    
    def __init__(self, app_instance):
        self.app = app_instance
    
    def dry_run(self):
        """Perform dry run"""
        if not self.app.file_handler.validate_processing_inputs():
            return
        
        selected = self.app.client_handler.get_selected_clients()
        if not selected:
            messagebox.showwarning("Warning", MESSAGES['errors']['no_selection'])
            return
        
        # Switch to processing tab
        self.app.notebook.select(2)
        
        self.app.root.after(0, lambda: self.app.log_message("🧪 ════════ DRY RUN PREVIEW ════════", 'info'))
        self.app.root.after(0, lambda: self.app.log_message(f"🎯 Mode: {self.app.processing_mode.get()}", 'info'))
        self.app.root.after(0, lambda: self.app.log_message(f"👥 Selected: {len(selected)} clients", 'info'))
        
        for client_key in selected:
            client_info = self.app.client_data[client_key]
            self.app.root.after(0, lambda: self.app.log_message(f"\n🏢 {client_info['client']} - {client_info['state']}", 'info'))
            
            for file_type, files in client_info['files'].items():
                for file_info in files:
                    self.app.root.after(0, lambda: self.app.log_message(f"  📄 Would organize: {file_info['name']}", 'normal'))
                    
            self.app.root.after(0, lambda: self.app.log_message("  📊 Would create: ITC Report", 'success'))
            self.app.root.after(0, lambda: self.app.log_message("  💰 Would create: Sales Report", 'success'))
            
        self.app.root.after(0, lambda: self.app.log_message("\n🧪 ════════ DRY RUN COMPLETE ════════", 'success'))
        self.app.root.after(0, lambda: self.app.log_message("💡 No files were actually moved or created", 'warning'))
    
    def start_processing(self):
        """Start processing selected clients"""
        if not self.app.file_handler.validate_processing_inputs():
            return
        
        selected = self.app.client_handler.get_selected_clients()
        if not selected:
            messagebox.showwarning("Warning", MESSAGES['errors']['no_selection'])
            return

        # CHANGE: Only warn about long client names when the
        # "Include client name in folders" option is enabled.
        # Use the configurable `client_name_max_length` if available
        # (falls back to 10). This is safer and easier to review.
        if self.app.include_client_name_in_folders.get():
            # Determine max length from app setting, with safe fallback
            try:
                max_len = int(self.app.client_name_max_length.get())
                if max_len <= 0:
                    max_len = 10
            except Exception:
                max_len = 10

            long_names = []
            for client_key in selected:
                client_info = self.app.client_data.get(client_key, {})
                client_name = client_info.get('client', '')
                # Only consider non-empty names
                if client_name and len(client_name) > max_len:
                    long_names.append((client_name, len(client_name)))

            if long_names:
                msg = f"Long client names detected (exceeding {max_len} chars):\n\n"
                msg += "\n".join(f"• {name} ({length} chars)" for name, length in long_names[:5])
                msg += "\n\nIf you continue, these client names will be used in folder names."
                # Keep the existing modal confirmation behavior
                if not messagebox.askyesno("Warning", msg + "\n\nContinue anyway?"):
                    return
        
        # Switch to processing tab
        self.app.notebook.select(2)
        
        # Update buttons
        self.app.start_btn.config(state='disabled')
        self.app.stop_btn.config(state='normal')
        
        # Reset stop flag
        self.app.stop_requested = False
        self.app.is_processing = True
        
        self.app.root.after(0, lambda: self.app.log_message("🚀 ════════ PROCESSING STARTED ════════", 'success'))
        
        # Start processing thread
        self.app.processing_thread = threading.Thread(
            target=self.process_files_thread,
            args=(selected,),
            daemon=True
        )
        self.app.processing_thread.start()
    
    def stop_processing(self):
        """Stop processing"""
        self.app.stop_requested = True
        self.app.log_message("🛑 Stop requested by user", 'warning')
    
    def process_files_thread(self, selected_clients):
        """Process files in background thread with detailed logging"""
        try:
            start_time = datetime.now()
            timestamp = get_timestamp()
            
            # Initialize organizer
            self.app.file_organizer = self.app.FileOrganizer(
                self.app.target_folder.get(),
                self.app.processing_mode.get(),
                include_client_name=self.app.include_client_name_in_folders.get(),
                client_folder_settings=self.app.client_folder_settings,
                client_name_max_length=self.app.client_name_max_length.get()
            )
            
            total_clients = len(selected_clients)
            processed = 0
            total_files = 0
            total_reports = 0
            
            # Initialize progress tracker for time estimation
            progress_tracker = ProgressTracker(total_clients)
            
            # Track level 1 folder
            level1_folder = None
            
            # Collect data for summary report
            summary_data = {
                'timestamp': timestamp,
                'total_clients': total_clients,
                'successful_clients': 0,
                'failed_clients': 0,
                'total_files': 0,
                'reports_generated': 0,
                'itc_reports_created': 0,
                'sales_reports_created': 0,
                'report_errors': 0,
                'processing_mode': self.app.processing_mode.get(),
                'target_folder': self.app.target_folder.get(),
                'clients': [],
                'file_mappings': [],
                'errors': [],
                'variations': self.app.variations,
                'include_client_name': self.app.include_client_name_in_folders.get()
            }
            
            for idx, client_key in enumerate(selected_clients):
                if self.app.stop_requested:
                    break
                
                client_info = self.app.client_data[client_key]
                
                # Update progress tracker and get time estimate
                progress_tracker.update(0 if idx == 0 else 1, f"Processing {client_info['client']}")
                time_remaining = progress_tracker.get_eta() if idx > 0 else "Calculating..."
                
                # Update progress with client counter and time remaining
                progress = (idx / total_clients) * 100
                status_msg = f"Client {idx + 1} of {total_clients}: {client_info['client']} | ETA: {time_remaining}"
                self.app.root.after(0, lambda: self.app.update_progress(progress, status_msg))
                
                try:
                    self.app.root.after(0, lambda: self.app.log_message(f"\n🏢 Processing {client_info['client']} - {client_info['state']}", 'info'))

                    # Track report creation success for this client
                    itc_success = False
                    sales_success = False
                    
                    # Create folders
                    folder_msg = f"Client {idx + 1}/{total_clients}: Creating folders | ETA: {time_remaining}"
                    self.app.root.after(0, lambda: self.app.update_progress(progress, folder_msg))
                    folders = self.app.file_organizer.create_client_structure(client_info)
                    
                    if level1_folder is None:
                        level1_folder = folders['level1']
                    
                    self.app.root.after(0, lambda: self.app.log_message("  ✓ Created folder structure", 'success'))
                    
                    # Log folder structure
                    self.app.root.after(0, lambda: self.app.log_message(f"    Level 1: {folders['level1'].name}", 'normal'))
                    self.app.root.after(0, lambda: self.app.log_message(f"    Level 2: {folders['level2'].name}", 'normal'))
                    self.app.root.after(0, lambda: self.app.log_message(f"    Level 3: {folders['version'].name}", 'normal'))
                    
                    # Organize files
                    def file_progress(current, total, message):
                        sub_progress = progress + (current / total) * (50 / total_clients)
                        file_msg = f"Client {idx + 1}/{total_clients}: {message} | ETA: {time_remaining}"
                        self.app.root.after(0, lambda: self.app.update_progress(sub_progress, file_msg))
                    
                    file_results = self.app.file_organizer.organize_files(
                        client_info, folders, file_progress
                    )
                    
                    successful_files = sum(1 for r in file_results if r['status'] == 'Success')
                    total_files += successful_files
                    summary_data['total_files'] += successful_files
                    
                    self.app.root.after(0, lambda: self.app.log_message(f"  ✓ Organized {successful_files} files", 'success'))
                    
                    # Log file organization details
                    for result in file_results:
                        if result['status'] == 'Success':
                            self.app.root.after(0, lambda: self.app.log_message(f"    📁 {result['filename']} → {result['file_type']} folder", 'normal'))
                    
                    # Create reports
                    report_progress = progress + 50 / total_clients
                    report_msg = f"Client {idx + 1}/{total_clients}: Creating Excel reports | ETA: {time_remaining}"
                    self.app.root.after(0, lambda: self.app.update_progress(report_progress, report_msg))
                    
                    # Prepare safe filenames
                    safe_client = sanitize_filename(client_info['client'])
                    safe_state = client_info.get('state_code', get_state_code(client_info['state']))
                    safe_timestamp = timestamp.replace(' ', '_')
                    
                    # ITC Report
                    itc_name = f"ITC_Report_{safe_client}_{safe_state}_{safe_timestamp}"
                    itc_output = folders['version'] / f"{itc_name}.xlsx"
                    
                    self.app.root.after(0, lambda: self.app.log_message("  📊 Creating ITC report...", 'info'))
                    self.app.root.after(0, lambda: self.app.log_message(f"    Template: {Path(self.app.itc_template.get()).name}", 'normal'))
                    self.app.root.after(0, lambda: self.app.log_message(f"    Output: {itc_output.name}", 'normal'))
                    
                    try:
                        itc_data = self.app.excel_handler.prepare_template_data(
                            client_info, folders, 'ITC'
                        )
                        
                        if self.app.excel_handler.create_report_from_template(
                            self.app.itc_template.get(),
                            itc_output,
                            itc_data,
                            'ITC'
                        ):
                            total_reports += 1
                            summary_data['reports_generated'] += 1
                            summary_data['itc_reports_created'] += 1
                            itc_success = True
                            self.app.log_message(f"    ✅ Success: {itc_name}.xlsx", 'success')
                        else:
                            summary_data['report_errors'] += 1
                            self.app.log_message("    ❌ Failed to create ITC report", 'error')
                            self.app.log_message("    Check if template has 'Links' sheet", 'warning')
                    except Exception as e:
                        summary_data['report_errors'] += 1
                        self.app.log_message(f"    ❌ ITC Report Error: {str(e)}", 'error')
                    
                    # Sales Report
                    self.app.log_message("  💰 Starting Sales report creation...", 'info')
                    
                    # Debug checkpoint 1: Validate template path
                    sales_template_path = self.app.sales_template.get()
                    self.app.log_message(f"    [DEBUG] Sales template path: {sales_template_path}", 'normal')
                    
                    if not sales_template_path:
                        self.app.log_message("    ❌ Sales template path is empty!", 'error')
                        summary_data['report_errors'] += 1
                    else:
                        try:
                            # Debug checkpoint 2: Validate template file exists
                            template_exists = Path(sales_template_path).exists()
                            self.app.log_message(f"    [DEBUG] Template file exists: {template_exists}", 'normal')
                            
                            if not template_exists:
                                self.app.log_message(f"    ❌ Sales template file not found at: {sales_template_path}", 'error')
                                summary_data['report_errors'] += 1
                            else:
                                # Safe template name extraction
                                try:
                                    template_name = Path(sales_template_path).name
                                except Exception as e:
                                    template_name = "Unknown"
                                    self.app.log_message(f"    [DEBUG] Error getting template name: {e}", 'warning')
                                
                                sales_name = f"Sales_Report_{safe_client}_{safe_state}_{safe_timestamp}"
                                sales_output = folders['version'] / f"{sales_name}.xlsx"
                                
                                self.app.log_message(f"    Template: {template_name}", 'normal')
                                self.app.log_message(f"    Output: {sales_output.name}", 'normal')
                                
                                # Debug checkpoint 3: Check output path
                                self.app.log_message(f"    [DEBUG] Full output path: {sales_output}", 'normal')
                                self.app.log_message(f"    [DEBUG] Output directory exists: {sales_output.parent.exists()}", 'normal')
                                
                                try:
                                    # Debug checkpoint 4: Prepare template data
                                    self.app.log_message("    [DEBUG] Preparing template data...", 'normal')
                                    sales_data = self.app.excel_handler.prepare_template_data(
                                        client_info, folders, 'Sales'
                                    )
                                    self.app.log_message(f"    [DEBUG] Template data keys: {list(sales_data.keys())}", 'normal')
                                    
                                    # Debug checkpoint 5: Create report
                                    self.app.log_message("    [DEBUG] Calling create_report_from_template...", 'normal')
                                    
                                    if self.app.excel_handler.create_report_from_template(
                                        sales_template_path,
                                        sales_output,
                                        sales_data,
                                        'Sales'
                                    ):
                                        total_reports += 1
                                        summary_data['reports_generated'] += 1
                                        summary_data['sales_reports_created'] += 1
                                        sales_success = True
                                        self.app.log_message(f"    ✅ Success: {sales_name}.xlsx", 'success')
                                        
                                        # Debug checkpoint 6: Verify output file
                                        if sales_output.exists():
                                            file_size = sales_output.stat().st_size
                                            self.app.log_message(f"    [DEBUG] Output file created, size: {file_size} bytes", 'normal')
                                        else:
                                            self.app.log_message("    [DEBUG] Warning: Output file not found after creation", 'warning')
                                    else:
                                        summary_data['report_errors'] += 1
                                        self.app.log_message("    ❌ Failed to create Sales report", 'error')
                                        self.app.log_message("    Check if template has 'Links' sheet", 'warning')
                                        self.app.log_message("    [DEBUG] create_report_from_template returned False", 'normal')
                                        
                                except Exception as e:
                                    summary_data['report_errors'] += 1
                                    self.app.log_message(f"    ❌ Sales Report Error: {str(e)}", 'error')
                                    self.app.log_message(f"    [DEBUG] Exception type: {type(e).__name__}", 'error')
                                    self.app.log_message(f"    [DEBUG] Full traceback will be in log file", 'error')
                                    import traceback
                                    logger.error(f"Sales Report Full Traceback:\n{traceback.format_exc()}")
                                    
                        except Exception as e:
                            summary_data['report_errors'] += 1
                            self.app.log_message(f"    ❌ Sales Report Setup Error: {str(e)}", 'error')
                            self.app.log_message(f"    [DEBUG] Error occurred before report creation", 'error')
                    
                    # Organization report
                    self.app.file_organizer.create_organization_report(folders, client_info)
                    
                    processed += 1
                    summary_data['successful_clients'] += 1
                    
                    # Add to summary data
                    summary_data['clients'].append({
                        'client': client_info['client'],
                        'state': client_info['state'],
                        'status': 'Success',
                        'file_count': client_info['file_count'],
                        'missing_files': client_info['missing_files'],
                        'completeness': client_info.get('completeness', 0),
                        'itc_report_created': itc_success,
                        'sales_report_created': sales_success
                    })
                    
                    for result in file_results:
                        summary_data['file_mappings'].append({
                            'filename': result['filename'],
                            'client': client_key,
                            'type': result['file_type'],
                            'destination': result.get('destination', ''),
                            'status': result['status']
                        })
                    
                    self.app.log_message(f"  ✅ Completed {client_info['client']}", 'success')
                    
                except Exception as e:
                    summary_data['failed_clients'] += 1
                    error_msg = str(e)
                    
                    summary_data['errors'].append({
                        'time': datetime.now().strftime('%H:%M:%S'),
                        'client': client_key,
                        'type': 'Processing Error',
                        'message': error_msg
                    })
                    
                    self.app.log_message(f"  ❌ Failed {client_key}: {error_msg}", 'error')
            
            # Create summary report
            if processed > 0 and level1_folder:
                final_time_remaining = progress_tracker.get_eta()
                summary_msg = f"Creating summary report | ETA: {final_time_remaining}"
                self.app.update_progress(95, summary_msg)
                self.app.log_message("\n📊 Creating Level 1 summary report...", 'info')
                
                safe_timestamp = timestamp.replace(' ', '_')
                summary_name = f"GST_Processing_Summary_{safe_timestamp}"
                summary_path = level1_folder / f"{summary_name}.xlsx"
                
                self.app.log_message(f"  Output: {summary_path.name}", 'normal')
                
                try:
                    if self.app.excel_handler.create_summary_report(summary_path, summary_data):
                        self.app.log_message(f"  ✅ Success: {summary_name}.xlsx", 'success')
                        self.app.log_message("  Sheets: Summary, Client Status, File Mapping, Errors, Variations", 'normal')
                    else:
                        self.app.log_message("  ❌ Failed to create summary report", 'error')
                except Exception as e:
                    self.app.log_message(f"  ❌ Summary Report Error: {str(e)}", 'error')
            
            # Final update
            elapsed_time = progress_tracker.get_elapsed_time()
            minutes = int(elapsed_time // 60)
            seconds = int(elapsed_time % 60)
            completion_msg = f"Processing complete | Total time: {minutes}m {seconds}s"
            self.app.update_progress(100, completion_msg)
            
            self.app.log_message("\n🎉 ════════ PROCESSING COMPLETE ════════", 'success')
            self.app.log_message(f"✅ Processed: {processed}/{total_clients} clients", 'info')
            self.app.log_message(f"📁 Files organized: {total_files}", 'info')
            self.app.log_message(f"📊 Reports created: {total_reports}", 'info')

            if level1_folder:
                self.app.log_message(f"📂 Output location: {level1_folder}", 'info')

            end_time = datetime.now()
            duration = end_time - start_time
            minutes = int(duration.total_seconds() // 60)
            seconds = int(duration.total_seconds() % 60)
            self.app.log_message(f"\n⏱️ Total processing time: {minutes} minutes {seconds} seconds", 'info')

            messagebox.showinfo(
                "Processing Complete",
                f"Successfully processed {processed}/{total_clients} clients\n\n"
                f"Files organized in:\n{self.app.target_folder.get()}"
            )

            # Open the output folder
            if level1_folder:
                try:
                    if platform.system() == 'Windows':
                        os.startfile(level1_folder)
                    elif platform.system() == 'Darwin':  # macOS
                        os.system(f'open "{level1_folder}"')
                    else:  # Linux
                        os.system(f'xdg-open "{level1_folder}"')
                
                    self.app.log_message(f"📂 Opened output folder", 'success')
                except Exception as e:
                    self.app.log_message(f"Could not open folder: {e}", 'warning')
                    
        except Exception as e:
            self.app.log_message(f"💥 Fatal error: {str(e)}", 'error')
            logger.error(f"Processing error: {str(e)}", exc_info=True)
            messagebox.showerror("Processing Error", str(e))
            
        finally:
            self.app.is_processing = False
            self.app.start_btn.config(state='normal')
            self.app.stop_btn.config(state='disabled')