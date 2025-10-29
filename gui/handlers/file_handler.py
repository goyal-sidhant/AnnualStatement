# gui/handlers/file_handler.py
"""File handling operations for GST Organizer"""

import logging
from pathlib import Path
from tkinter import filedialog, messagebox
from utils.constants import MESSAGES
from utils.helpers import truncate_path

logger = logging.getLogger(__name__)


class FileHandler:
    """Handles file browsing, scanning, and validation"""
    
    def __init__(self, app_instance):
        self.app = app_instance
    
    def browse_source_folder(self):
        """Browse for source folder"""
        folder = filedialog.askdirectory(title="Select Source Folder")
        if folder:
            self.app.source_folder.set(folder)
            self.app.update_status(f"Source: {truncate_path(folder)}")
            self.app.save_cache()
    
    def browse_itc_template(self):
        """Browse for ITC template"""
        file = filedialog.askopenfilename(
            title="Select ITC Template",
            filetypes=[("Excel Template", "*.xltx *.xlsx")]
        )
        if file:
            self.app.itc_template.set(file)
            self.app.update_status(f"ITC template selected")
            self.app.save_cache()
    
    def browse_sales_template(self):
        """Browse for Sales template"""
        file = filedialog.askopenfilename(
            title="Select Sales Template",
            filetypes=[("Excel Template", "*.xltx *.xlsx")]
        )
        if file:
            self.app.sales_template.set(file)
            self.app.update_status(f"Sales template selected")
            self.app.save_cache()
    
    def browse_target_folder(self):
        """Browse for target folder"""
        folder = filedialog.askdirectory(title="Select Target Folder")
        if folder:
            self.app.target_folder.set(folder)
            self.app.update_status(f"Target: {truncate_path(folder)}")
            self.app.save_cache()
    
    def scan_files(self):
        """Scan files in source folder"""
        if not self.validate_scan_inputs():
            return
        
        try:
            self.app.update_status("Scanning files...")
            self.app.log_message("🔍 Starting file scan...", 'info')
            
            # Progress callback
            def progress_callback(current, total, message):
                self.app.update_status(f"Scanning: {message} ({current}/{total})")
            
            # Scan folder
            self.app.scanned_files, self.app.client_data, self.app.variations = \
                self.app.file_parser.scan_folder(self.app.source_folder.get(), progress_callback)
            
            # Show count while scanning
            file_count = len(self.app.scanned_files)
            self.app.update_status(f"🔍 Found {file_count} Excel files")
            
            # Update displays
            self.app.update_summary()
            self.app.update_client_tree()
            
            if self.app.client_data:
                msg = f"Found {len(self.app.client_data)} clients"
                self.app.update_status(msg)
                self.app.log_message(f"✅ {msg}", 'success')
                
                # Switch to validation tab
                self.app.notebook.select(1)
            else:
                self.app.update_status("No valid GST files found")
                self.app.log_message("⚠️ No valid GST files found", 'warning')
                
        except Exception as e:
            self.app.log_message(f"❌ Scan error: {str(e)}", 'error')
            messagebox.showerror("Scan Error", str(e))
    
    def rescan_files(self):
        """Re-scan files"""
        self.app.log_message("🔄 Re-scanning files...", 'info')
        
        # Clear data
        self.app.scanned_files.clear()
        self.app.client_data.clear()
        self.app.selected_clients.clear()
        self.app.variations.clear()
        
        # Clear displays
        if hasattr(self.app, 'summary_text'):
            self.app.summary_text.config(state='normal')
            self.app.summary_text.delete(1.0, 'end')
            self.app.summary_text.config(state='disabled')
        
        if hasattr(self.app, 'client_tree'):
            for item in self.app.client_tree.get_children():
                self.app.client_tree.delete(item)
        
        # Re-scan
        self.scan_files()
    
    def validate_scan_inputs(self):
        """Validate inputs for scanning"""
        if not self.app.source_folder.get():
            messagebox.showerror("Error", MESSAGES['errors']['no_source'])
            return False
        
        if not Path(self.app.source_folder.get()).exists():
            messagebox.showerror("Error", "Source folder does not exist")
            return False
        
        return True
    
    def validate_processing_inputs(self):
        """Validate processing inputs"""
        if not self.app.itc_template.get() or not self.app.sales_template.get():
            messagebox.showerror("Error", MESSAGES['errors']['no_templates'])
            return False
        
        if not self.app.target_folder.get():
            messagebox.showerror("Error", MESSAGES['errors']['no_target'])
            return False
        
        # Check templates exist
        if not Path(self.app.itc_template.get()).exists():
            messagebox.showerror("Error", "ITC template file not found")
            return False
        
        if not Path(self.app.sales_template.get()).exists():
            messagebox.showerror("Error", "Sales template file not found")
            return False
        
        return True