"""
Power Query Extractor GUI Window - Two Tab Design
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import logging
import threading
import json
from pathlib import Path
from datetime import datetime
from ..core.report_processor import ReportProcessor
from ..core.data_consolidator import DataConsolidator

logger = logging.getLogger(__name__)


class PowerQueryExtractorApp(tk.Tk):
    """Main window for PQ Extractor with two tabs"""
    
    def __init__(self, target_folder=None):
        super().__init__()
        
        self.target_folder = target_folder
        self.processor = ReportProcessor()
        self.consolidator = DataConsolidator()
        self.client_vars = {}
        self.cache_file = Path("gst_organizer_cache.json")
        self.is_processing = False
        
        self.setup_window()
        self.create_widgets()
        
        # Auto-load folder if available
        if target_folder:
            self.folder_path.set(target_folder)
            self.scan_folder()
        else:
            self.load_from_cache()
    
    def setup_window(self):
        """Configure window settings"""
        self.title("Power Query Extractor")
        self.geometry("1000x700")
        self.minsize(900, 600)
        
        # Variables
        self.folder_path = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.progress_var = tk.DoubleVar()
        
        # Style
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure notebook style
        style.configure('TNotebook', background='#f0f0f0')
        style.configure('TNotebook.Tab', padding=[20, 10])
    
    def create_widgets(self):
        """Create all GUI widgets"""
        # Title
        title_frame = tk.Frame(self, bg='#0078D4', height=60)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        
        tk.Label(
            title_frame,
            text="🔄 Power Query Report Extractor",
            font=('Segoe UI', 18, 'bold'),
            bg='#0078D4',
            fg='white'
        ).pack(pady=15)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Tab 1: Setup
        self.setup_tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.setup_tab, text='Setup')
        self.create_setup_tab()
        
        # Tab 2: Processing
        self.processing_tab = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.processing_tab, text='Processing')
        self.create_processing_tab()
        
        # Status bar
        self.create_status_bar()
        
        # Bind tab change event
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
    
    def create_setup_tab(self):
        """Create setup tab content"""
        # Main container
        main_frame = tk.Frame(self.setup_tab, bg='white')
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Folder selection section
        folder_frame = tk.LabelFrame(
            main_frame, 
            text="Target Folder", 
            font=('Segoe UI', 11, 'bold'),
            bg='white',
            padx=15, 
            pady=15
        )
        folder_frame.pack(fill='x', pady=(0, 20))
        
        # Folder path entry
        entry_frame = tk.Frame(folder_frame, bg='white')
        entry_frame.pack(fill='x')
        
        tk.Label(
            entry_frame,
            text="Folder Path:",
            bg='white',
            font=('Segoe UI', 10)
        ).pack(side='left', padx=(0, 10))
        
        tk.Entry(
            entry_frame,
            textvariable=self.folder_path,
            font=('Segoe UI', 10),
            width=50
        ).pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        tk.Button(
            entry_frame,
            text="📁 Browse",
            command=self.browse_folder,
            font=('Segoe UI', 10),
            bg='#6c757d',
            fg='white',
            padx=15
        ).pack(side='left', padx=(0, 5))
        
        tk.Button(
            entry_frame,
            text="🔍 Scan",
            command=self.scan_folder,
            font=('Segoe UI', 10, 'bold'),
            bg='#28a745',
            fg='white',
            padx=15
        ).pack(side='left')
        
        # Client selection section
        client_frame = tk.LabelFrame(
            main_frame,
            text="Select Clients",
            font=('Segoe UI', 11, 'bold'),
            bg='white',
            padx=15,
            pady=15
        )
        client_frame.pack(fill='both', expand=True, pady=(0, 20))
        
        # Buttons
        btn_frame = tk.Frame(client_frame, bg='white')
        btn_frame.pack(fill='x', pady=(0, 10))
        
        tk.Button(
            btn_frame,
            text="✓ Select All",
            command=self.select_all_clients,
            font=('Segoe UI', 9),
            bg='#17a2b8',
            fg='white',
            padx=10
        ).pack(side='left', padx=(0, 5))
        
        tk.Button(
            btn_frame,
            text="✗ Deselect All",
            command=self.deselect_all_clients,
            font=('Segoe UI', 9),
            bg='#6c757d',
            fg='white',
            padx=10
        ).pack(side='left')

        # Add some spacing
        tk.Label(btn_frame, text="    ", bg='white').pack(side='left')
        
        # Skip refresh checkbox
        self.skip_refresh_var = tk.BooleanVar(value=False)
        self.skip_refresh_check = tk.Checkbutton(
            btn_frame,
            text="Skip Refresh (Use existing refreshed files)",
            variable=self.skip_refresh_var,
            command=self.on_skip_refresh_toggle,
            font=('Segoe UI', 10),
            bg='white'
        )
        self.skip_refresh_check.pack(side='left', padx=20)
        
        # Client list with scrollbar
        list_frame = tk.Frame(client_frame, bg='white')
        list_frame.pack(fill='both', expand=True)
        
        canvas = tk.Canvas(list_frame, bg='white', highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg='white')
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Process button
        self.process_btn = tk.Button(
            main_frame,
            text="🚀 Start Processing",
            command=self.start_processing,
            bg='#0078D4',
            fg='white',
            font=('Segoe UI', 12, 'bold'),
            padx=40,
            pady=15,
            state='disabled',
            cursor='hand2'
        )
        self.process_btn.pack(pady=(10, 0))

    def on_skip_refresh_toggle(self):
        """Handle skip refresh checkbox toggle"""
        # Update the configuration
        from ..config.cell_mappings import EXTRACTION_OPTIONS
        EXTRACTION_OPTIONS['skip_refresh'] = self.skip_refresh_var.get()
        
        if self.skip_refresh_var.get():
            self.log_message("Skip refresh enabled - will use existing refreshed files", 'info')
        else:
            self.log_message("Skip refresh disabled - will refresh all files", 'info')
    
    def create_processing_tab(self):
        """Create processing tab content"""
        # Main container
        main_frame = tk.Frame(self.processing_tab, bg='white')
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Progress section
        progress_frame = tk.LabelFrame(
            main_frame,
            text="Processing Progress",
            font=('Segoe UI', 11, 'bold'),
            bg='white',
            padx=15,
            pady=15
        )
        progress_frame.pack(fill='x', pady=(0, 20))
        
        # Progress label
        self.progress_label = tk.Label(
            progress_frame,
            text="Ready to process",
            bg='white',
            font=('Segoe UI', 10)
        )
        self.progress_label.pack(anchor='w')
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            length=400,
            style='TProgressbar'
        )
        self.progress_bar.pack(fill='x', pady=(10, 0))
        
        # Log section
        log_frame = tk.LabelFrame(
            main_frame,
            text="Processing Log",
            font=('Segoe UI', 11, 'bold'),
            bg='white',
            padx=15,
            pady=15
        )
        log_frame.pack(fill='both', expand=True)
        
        # Log text with scrollbar
        log_container = tk.Frame(log_frame, bg='white')
        log_container.pack(fill='both', expand=True)
        
        self.log_text = tk.Text(
            log_container,
            height=20,
            wrap='word',
            font=('Consolas', 9),
            bg='#f8f9fa'
        )
        self.log_text.pack(side='left', fill='both', expand=True)
        
        log_scroll = ttk.Scrollbar(log_container, command=self.log_text.yview)
        log_scroll.pack(side='right', fill='y')
        self.log_text.config(yscrollcommand=log_scroll.set)
        
        # Configure text tags for colors
        self.log_text.tag_config('info', foreground='#000000')
        self.log_text.tag_config('success', foreground='#28a745')
        self.log_text.tag_config('warning', foreground='#ffc107')
        self.log_text.tag_config('error', foreground='#dc3545')
    
    def create_status_bar(self):
        """Create status bar"""
        status_frame = tk.Frame(self, bg='#e0e0e0', height=25)
        status_frame.pack(fill='x', side='bottom')
        status_frame.pack_propagate(False)
        
        tk.Label(
            status_frame,
            textvariable=self.status_var,
            bg='#e0e0e0',
            font=('Segoe UI', 9),
            anchor='w'
        ).pack(side='left', padx=10)
    
    def load_from_cache(self):
        """Load settings from cache file"""
        try:
            if self.cache_file.exists():
                with open(self.cache_file, 'r') as f:
                    cache_data = json.load(f)
                    folder = cache_data.get('target_folder', '')
                    if folder:
                        self.folder_path.set(folder)
                        self.log_message("Loaded folder from cache", 'info')
                        self.scan_folder()
        except Exception as e:
            logger.warning(f"Could not load cache: {e}")
    
    def browse_folder(self):
        """Browse for target folder"""
        folder = filedialog.askdirectory(
            title="Select Target Folder",
            initialdir=self.folder_path.get() or None
        )
        if folder:
            self.folder_path.set(folder)
            self.scan_folder()
    
    def scan_folder(self):
        """Scan folder for clients"""
        if not self.folder_path.get():
            messagebox.showwarning("No Folder", "Please select a folder first")
            return
        
        self.status_var.set("Scanning folder...")
        self.log_message("Scanning folder for clients...", 'info')
        
        try:
            # Find Annual Statement folder
            base_path = Path(self.folder_path.get())
            annual_folders = list(base_path.glob("Annual Statement-*"))
            
            if not annual_folders:
                self.log_message("No Annual Statement folders found", 'warning')
                self.status_var.set("No processed folders found")
                return
            
            # Use latest Annual Statement folder
            latest_folder = max(annual_folders, key=lambda x: x.stat().st_mtime)
            self.level1_folder = latest_folder  # Store for later use
            
            self.log_message(f"Using folder: {latest_folder.name}", 'info')
            
            # Find client folders
            clients = []
            for client_folder in latest_folder.iterdir():
                if client_folder.is_dir() and not client_folder.name.startswith('_'):
                    # Check for version folders
                    version_folders = list(client_folder.glob("Version-*"))
                    if version_folders:
                        latest_version = max(version_folders, key=lambda x: x.name)
                        
                        # Check for report files
                        itc_reports = list(latest_version.glob("ITC_Report_*.xlsx"))
                        sales_reports = list(latest_version.glob("Sales_Report_*.xlsx"))
                        
                        clients.append({
                            'name': client_folder.name,
                            'path': client_folder,
                            'latest_version': latest_version,
                            'has_itc': len(itc_reports) > 0,
                            'has_sales': len(sales_reports) > 0
                        })
            
            self.display_clients(clients)
            self.log_message(f"Found {len(clients)} clients", 'success')
            self.status_var.set(f"Found {len(clients)} clients")
            
        except Exception as e:
            self.log_message(f"Error scanning folder: {e}", 'error')
            self.status_var.set("Error scanning folder")
            logger.error(f"Scan error: {e}", exc_info=True)
    
    def display_clients(self, clients):
        """Display client checkboxes"""
        # Clear existing
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.client_vars.clear()
        
        if not clients:
            tk.Label(
                self.scrollable_frame,
                text="No clients found",
                bg='white',
                font=('Segoe UI', 10, 'italic'),
                fg='#666666'
            ).pack(pady=20)
            return
        
        # Create checkboxes
        for client in clients:
            var = tk.BooleanVar(value=True)
            self.client_vars[client['name']] = {
                'var': var,
                'data': client
            }
            
            # Frame for each client
            client_frame = tk.Frame(self.scrollable_frame, bg='white')
            client_frame.pack(fill='x', padx=5, pady=2)
            
            cb = tk.Checkbutton(
                client_frame,
                text=client['name'],
                variable=var,
                bg='white',
                font=('Segoe UI', 10),
                anchor='w'
            )
            cb.pack(side='left', fill='x', expand=True)
            
            # Status indicators
            if client['has_itc']:
                tk.Label(
                    client_frame,
                    text="ITC ✓",
                    bg='white',
                    fg='#28a745',
                    font=('Segoe UI', 9)
                ).pack(side='left', padx=5)
            
            if client['has_sales']:
                tk.Label(
                    client_frame,
                    text="Sales ✓",
                    bg='white',
                    fg='#28a745',
                    font=('Segoe UI', 9)
                ).pack(side='left', padx=5)
        
        # Enable process button
        if clients:
            self.process_btn.config(state='normal')
    
    def select_all_clients(self):
        """Select all clients"""
        for client_data in self.client_vars.values():
            client_data['var'].set(True)
    
    def deselect_all_clients(self):
        """Deselect all clients"""
        for client_data in self.client_vars.values():
            client_data['var'].set(False)
    
    def on_tab_changed(self, event):
        """Handle tab change"""
        selected_tab = event.widget.tab('current')['text']
        if selected_tab == 'Processing' and self.is_processing:
            # Refresh the display when switching to processing tab
            self.update_idletasks()
    
    def start_processing(self):
        """Start processing selected clients"""
        selected = [
            data['data'] for name, data in self.client_vars.items()
            if data['var'].get()
        ]
        
        if not selected:
            messagebox.showwarning("No Selection", "Please select at least one client")
            return
        
        # Switch to processing tab
        self.notebook.select(self.processing_tab)
        
        # Disable process button
        self.process_btn.config(state='disabled')
        self.is_processing = True
        
        # Clear log
        self.log_text.delete(1.0, tk.END)
        
        # Start processing thread
        thread = threading.Thread(
            target=self._process_clients_thread,
            args=(selected,)
        )
        thread.daemon = True
        thread.start()
    
    def _process_clients_thread(self, selected_clients):
        """Process clients (runs in thread)"""
        try:
            self.log_message("=" * 60, 'info')
            self.log_message("PROCESSING STARTED", 'info')
            self.log_message("=" * 60, 'info')
            
            results = []
            total = len(selected_clients)
            
            for i, client in enumerate(selected_clients):
                client_name = client['name']
                self.log_message(f"\n📋 Processing {client_name}...", 'info')
                
                # Update progress
                progress = (i / total) * 100
                self.update_progress(progress, f"Processing {client_name}")
                
                # Process reports
                result = self.processor.process_client(client, log_callback=self.log_message)
                results.append(result)
                
                # Log results
                itc_status = result['itc']['status']
                sales_status = result['sales']['status']
                
                if itc_status['success']:
                    self.log_message(f"  ✅ ITC Report - Refreshed successfully", 'success')
                else:
                    self.log_message(f"  ⚠️ ITC Report - {itc_status.get('error', 'Error occurred')}", 'warning')
                    if result['itc']['data']:
                        self.log_message(f"     (Data extracted despite error)", 'info')
                
                if sales_status['success']:
                    self.log_message(f"  ✅ Sales Report - Refreshed successfully", 'success')
                else:
                    self.log_message(f"  ⚠️ Sales Report - {sales_status.get('error', 'Error occurred')}", 'warning')
                    if result['sales']['data']:
                        self.log_message(f"     (Data extracted despite error)", 'info')
            
            # Create consolidated report in Level 1 folder
            self.log_message("\n📊 Creating consolidated report...", 'info')
            self.update_progress(95, "Creating consolidated report")
            
            # Use Level 1 folder (Annual Statement folder)
            report_path = self.consolidator.create_report(results, self.level1_folder)
            
            self.update_progress(100, "Processing complete")
            self.log_message(f"\n✅ Processing complete!", 'success')
            self.log_message(f"📄 Report saved: {report_path.name}", 'success')
            self.log_message(f"📂 Location: {report_path.parent}", 'info')
            
            self.status_var.set("Processing complete")
            
            # Ask to open report
            if messagebox.askyesno("Complete", "Processing complete!\n\nOpen the consolidated report?"):
                import os
                os.startfile(report_path)
            
        except Exception as e:
            self.log_message(f"\n❌ Error: {e}", 'error')
            logger.error(f"Processing error: {e}", exc_info=True)
            messagebox.showerror("Error", f"Processing failed: {e}")
        finally:
            self.is_processing = False
            self.process_btn.config(state='normal')
            self.update_progress(0, "Ready")

            # Cleanup Excel instance
            self.processor.cleanup()
    
    def update_progress(self, value, message):
        """Update progress bar and label"""
        self.progress_var.set(value)
        self.progress_label.config(text=message)
        self.status_var.set(message)
        self.update_idletasks()
    
    def log_message(self, message, level='info'):
        """Add message to log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Insert message with timestamp
        self.log_text.insert('end', f"[{timestamp}] {message}\n", level)
        self.log_text.see('end')
        self.update_idletasks()