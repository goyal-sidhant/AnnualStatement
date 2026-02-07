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
        """Create setup tab content with two-column layout"""
        # Main container with two columns
        main_frame = tk.Frame(self.setup_tab, bg='white')
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # LEFT COLUMN - Settings
        left_column = tk.Frame(main_frame, bg='white', width=400)
        left_column.pack(side='left', fill='both', padx=(0, 10))
        left_column.pack_propagate(False)  # Maintain fixed width

        # Folder selection section
        folder_frame = tk.LabelFrame(
            left_column,
            text="Target Folder",
            font=('Segoe UI', 11, 'bold'),
            bg='white',
            padx=15,
            pady=15
        )
        folder_frame.pack(fill='x', pady=(0, 10))

        # Folder path label
        tk.Label(
            folder_frame,
            text="Folder Path:",
            bg='white',
            font=('Segoe UI', 10)
        ).pack(anchor='w', pady=(0, 5))

        # Folder path entry
        tk.Entry(
            folder_frame,
            textvariable=self.folder_path,
            font=('Segoe UI', 9),
            width=35
        ).pack(fill='x', pady=(0, 10))

        # Browse and Scan buttons
        btn_frame = tk.Frame(folder_frame, bg='white')
        btn_frame.pack(fill='x')

        tk.Button(
            btn_frame,
            text="📁 Browse",
            command=self.browse_folder,
            font=('Segoe UI', 9),
            bg='#6c757d',
            fg='white',
            padx=15,
            pady=5
        ).pack(side='left', padx=(0, 5))

        tk.Button(
            btn_frame,
            text="🔍 Scan",
            command=self.scan_folder,
            font=('Segoe UI', 9, 'bold'),
            bg='#28a745',
            fg='white',
            padx=15,
            pady=5
        ).pack(side='left')

        # Processing Options Section
        options_frame = tk.LabelFrame(
            left_column,
            text="Processing Options",
            font=('Segoe UI', 11, 'bold'),
            bg='white',
            padx=15,
            pady=15
        )
        options_frame.pack(fill='x', pady=(0, 10))

        # Wait time configuration
        tk.Label(
            options_frame,
            text="Power Query Wait Time:",
            bg='white',
            font=('Segoe UI', 10)
        ).pack(anchor='w', pady=(0, 5))

        wait_frame = tk.Frame(options_frame, bg='white')
        wait_frame.pack(fill='x', pady=(0, 10))

        self.wait_time_var = tk.IntVar(value=10)
        wait_spinbox = tk.Spinbox(
            wait_frame,
            from_=5,
            to=60,
            textvariable=self.wait_time_var,
            width=8,
            font=('Segoe UI', 10)
        )
        wait_spinbox.pack(side='left')

        tk.Label(
            wait_frame,
            text="seconds (Min: 5, Max: 60)",
            bg='white',
            font=('Segoe UI', 9),
            fg='#666'
        ).pack(side='left', padx=(10, 0))

        # File suffix configuration
        tk.Label(
            options_frame,
            text="Refreshed File Suffix:",
            bg='white',
            font=('Segoe UI', 10)
        ).pack(anchor='w', pady=(0, 5))

        self.suffix_pattern_var = tk.StringVar(value="_Refreshed_{timestamp}")
        suffix_entry = tk.Entry(
            options_frame,
            textvariable=self.suffix_pattern_var,
            font=('Segoe UI', 9),
            width=35
        )
        suffix_entry.pack(fill='x', pady=(0, 5))

        tk.Label(
            options_frame,
            text="Use {timestamp} for datetime",
            bg='white',
            font=('Segoe UI', 9),
            fg='#666'
        ).pack(anchor='w', pady=(0, 10))

        # Skip refresh checkbox
        self.skip_refresh_var = tk.BooleanVar(value=False)
        self.skip_refresh_check = tk.Checkbutton(
            options_frame,
            text="Skip Refresh\n(Use existing refreshed files)",
            variable=self.skip_refresh_var,
            command=self.on_skip_refresh_toggle,
            font=('Segoe UI', 9),
            bg='white',
            justify='left'
        )
        self.skip_refresh_check.pack(anchor='w')

        # Process button (at bottom of left column)
        self.process_btn = tk.Button(
            left_column,
            text="🚀 Start Processing",
            command=self.start_processing,
            bg='#0078D4',
            fg='white',
            font=('Segoe UI', 11, 'bold'),
            padx=20,
            pady=12,
            state='disabled',
            cursor='hand2'
        )
        self.process_btn.pack(fill='x', pady=(10, 0))

        # RIGHT COLUMN - Client Selection
        right_column = tk.Frame(main_frame, bg='white')
        right_column.pack(side='left', fill='both', expand=True)

        # Client selection section
        client_frame = tk.LabelFrame(
            right_column,
            text="Select Clients & Reports",
            font=('Segoe UI', 11, 'bold'),
            bg='white',
            padx=15,
            pady=15
        )
        client_frame.pack(fill='both', expand=True)

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
            padx=10,
            pady=3
        ).pack(side='left', padx=(0, 5))

        tk.Button(
            btn_frame,
            text="✗ Deselect All",
            command=self.deselect_all_clients,
            font=('Segoe UI', 9),
            bg='#6c757d',
            fg='white',
            padx=10,
            pady=3
        ).pack(side='left', padx=(0, 5))

        tk.Button(
            btn_frame,
            text="ITC Only",
            command=self.select_itc_only,
            font=('Segoe UI', 9),
            bg='#0078D4',
            fg='white',
            padx=10,
            pady=3
        ).pack(side='left', padx=(0, 5))

        tk.Button(
            btn_frame,
            text="Sales Only",
            command=self.select_sales_only,
            font=('Segoe UI', 9),
            bg='#0078D4',
            fg='white',
            padx=10,
            pady=3
        ).pack(side='left')

        # Client list with scrollbar
        list_frame = tk.Frame(client_frame, bg='white')
        list_frame.pack(fill='both', expand=True)

        # Add header labels
        header_frame = tk.Frame(list_frame, bg='white')
        header_frame.pack(fill='x', pady=(5, 0))

        tk.Label(header_frame, text="Client Name", font=('Segoe UI', 10, 'bold'),
                 bg='white', width=25, anchor='w').pack(side='left', padx=(5, 0))
        tk.Label(header_frame, text="ITC", font=('Segoe UI', 10, 'bold'),
                 bg='white', width=8).pack(side='left')
        tk.Label(header_frame, text="Sales", font=('Segoe UI', 10, 'bold'),
                 bg='white', width=8).pack(side='left')
        tk.Label(header_frame, text="Last Refresh Status", font=('Segoe UI', 10, 'bold'),
                 bg='white', anchor='w').pack(side='left', fill='x', expand=True)

        canvas = tk.Canvas(list_frame, bg='white', highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas, bg='white')

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Mouse wheel scrolling function
        def _on_mousewheel(event):
            # Check if there's content to scroll
            if canvas.winfo_height() < self.scrollable_frame.winfo_reqheight():
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            return "break"

        # Windows mouse wheel binding
        canvas.bind("<MouseWheel>", _on_mousewheel)
        self.scrollable_frame.bind("<MouseWheel>", _on_mousewheel)

        # Focus-based binding to avoid conflicts
        def _bind_wheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_wheel(event):
            canvas.unbind_all("<MouseWheel>")

        canvas.bind("<Enter>", _bind_wheel)
        canvas.bind("<Leave>", _unbind_wheel)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

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
        
        # Mouse wheel scrolling for log text
        def _on_log_mousewheel(event):
            if self.log_text.compare("end-1c", "!=", "1.0"):  # Check if there's content
                self.log_text.yview_scroll(int(-1*(event.delta/120)), "units")
            return "break"

        # Windows mouse wheel binding for log text
        self.log_text.bind("<MouseWheel>", _on_log_mousewheel)
        
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
        """Display client checkboxes with ITC/Sales selection"""
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

        # Create checkboxes for each client
        for client in clients:
            # Get last refresh status
            itc_status, sales_status = self.get_refresh_status(client)

            # Create variables for ITC and Sales checkboxes
            itc_var = tk.BooleanVar(value=client['has_itc'])
            sales_var = tk.BooleanVar(value=client['has_sales'])

            self.client_vars[client['name']] = {
                'data': client,
                'itc_var': itc_var,
                'sales_var': sales_var,
                'itc_status': itc_status,
                'sales_status': sales_status
            }

            # Frame for each client
            client_frame = tk.Frame(self.scrollable_frame, bg='white')
            client_frame.pack(fill='x', padx=5, pady=2)

            # Client name
            name_label = tk.Label(
                client_frame,
                text=client['name'],
                bg='white',
                font=('Segoe UI', 10),
                width=30,
                anchor='w'
            )
            name_label.pack(side='left')

            # ITC checkbox
            itc_cb = tk.Checkbutton(
                client_frame,
                variable=itc_var,
                bg='white',
                state='normal' if client['has_itc'] else 'disabled',
                width=8
            )
            itc_cb.pack(side='left')

            # Sales checkbox
            sales_cb = tk.Checkbutton(
                client_frame,
                variable=sales_var,
                bg='white',
                state='normal' if client['has_sales'] else 'disabled',
                width=8
            )
            sales_cb.pack(side='left')

            # Status label
            status_text = self.format_refresh_status(itc_status, sales_status)
            status_label = tk.Label(
                client_frame,
                text=status_text,
                bg='white',
                font=('Segoe UI', 9),
                fg='#666',
                width=40,
                anchor='w'
            )
            status_label.pack(side='left')

        # Enable process button
        if clients:
            self.process_btn.config(state='normal')
    
    def get_refresh_status(self, client):
        """Get last refresh status from existing files"""
        itc_status = None
        sales_status = None

        try:
            # Check for existing refreshed ITC files
            itc_refreshed = list(client['latest_version'].glob("ITC_Report_*_Refreshed_*.xlsx"))
            if itc_refreshed:
                latest_itc = max(itc_refreshed, key=lambda x: x.stat().st_mtime)
                mtime = datetime.fromtimestamp(latest_itc.stat().st_mtime)
                itc_status = mtime.strftime("%Y-%m-%d %H:%M")

            # Check for existing refreshed Sales files
            sales_refreshed = list(client['latest_version'].glob("Sales_Report_*_Refreshed_*.xlsx"))
            if sales_refreshed:
                latest_sales = max(sales_refreshed, key=lambda x: x.stat().st_mtime)
                mtime = datetime.fromtimestamp(latest_sales.stat().st_mtime)
                sales_status = mtime.strftime("%Y-%m-%d %H:%M")
        except (OSError, ValueError) as e:
            logger.warning(f"Error checking refresh status: {e}")

        return itc_status, sales_status

    def format_refresh_status(self, itc_status, sales_status):
        """Format refresh status for display"""
        if itc_status and sales_status:
            return f"ITC: {itc_status} | Sales: {sales_status}"
        elif itc_status:
            return f"ITC: {itc_status} | Sales: Never"
        elif sales_status:
            return f"ITC: Never | Sales: {sales_status}"
        else:
            return "Never refreshed"

    def select_all_clients(self):
        """Select all clients and reports"""
        for client_data in self.client_vars.values():
            if client_data['data']['has_itc']:
                client_data['itc_var'].set(True)
            if client_data['data']['has_sales']:
                client_data['sales_var'].set(True)

    def deselect_all_clients(self):
        """Deselect all clients and reports"""
        for client_data in self.client_vars.values():
            client_data['itc_var'].set(False)
            client_data['sales_var'].set(False)

    def select_itc_only(self):
        """Select only ITC reports"""
        for client_data in self.client_vars.values():
            if client_data['data']['has_itc']:
                client_data['itc_var'].set(True)
            client_data['sales_var'].set(False)

    def select_sales_only(self):
        """Select only Sales reports"""
        for client_data in self.client_vars.values():
            client_data['itc_var'].set(False)
            if client_data['data']['has_sales']:
                client_data['sales_var'].set(True)
    
    def on_tab_changed(self, event):
        """Handle tab change"""
        selected_tab = event.widget.tab('current')['text']
        if selected_tab == 'Processing' and self.is_processing:
            # Refresh the display when switching to processing tab
            self.update_idletasks()
    
    def start_processing(self):
        """Start processing selected clients and reports"""
        selected = []

        for name, data in self.client_vars.items():
            process_itc = data['itc_var'].get()
            process_sales = data['sales_var'].get()

            if process_itc or process_sales:
                client_info = data['data'].copy()
                client_info['process_itc'] = process_itc
                client_info['process_sales'] = process_sales
                selected.append(client_info)

        if not selected:
            messagebox.showwarning("No Selection", "Please select at least one report to process")
            return

        # Pass wait time and suffix to processor
        self.processor.wait_time = self.wait_time_var.get()
        self.processor.suffix_pattern = self.suffix_pattern_var.get()

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
        """Process clients with selective report processing"""
        try:
            self.log_message("=" * 60, 'info')
            self.log_message("PROCESSING STARTED", 'info')
            self.log_message(f"Wait Time: {self.wait_time_var.get()} seconds", 'info')
            self.log_message(f"File Suffix: {self.suffix_pattern_var.get()}", 'info')
            self.log_message("=" * 60, 'info')

            results = []
            total = len(selected_clients)

            for i, client in enumerate(selected_clients):
                client_name = client['name']

                # Show what will be processed
                reports_to_process = []
                if client.get('process_itc'):
                    reports_to_process.append('ITC')
                if client.get('process_sales'):
                    reports_to_process.append('Sales')

                self.log_message(f"\n📋 Processing {client_name} - Reports: {', '.join(reports_to_process)}", 'info')

                # Update progress
                progress = (i / total) * 100
                self.update_progress(progress, f"Processing {client_name}")

                # Process reports
                result = self.processor.process_client(client, log_callback=self.log_message)
                results.append(result)

                # Log results
                if client.get('process_itc'):
                    itc_status = result['itc']['status']
                    if itc_status['success']:
                        self.log_message(f"  ✅ ITC Report - Refreshed successfully", 'success')
                    else:
                        self.log_message(f"  ⚠️ ITC Report - {itc_status.get('error', 'Error occurred')}", 'warning')
                        if result['itc']['data']:
                            self.log_message(f"     (Data extracted despite error)", 'info')
                else:
                    self.log_message(f"  ⏭️ ITC Report - Skipped", 'info')

                if client.get('process_sales'):
                    sales_status = result['sales']['status']
                    if sales_status['success']:
                        self.log_message(f"  ✅ Sales Report - Refreshed successfully", 'success')
                    else:
                        self.log_message(f"  ⚠️ Sales Report - {sales_status.get('error', 'Error occurred')}", 'warning')
                        if result['sales']['data']:
                            self.log_message(f"     (Data extracted despite error)", 'info')
                else:
                    self.log_message(f"  ⏭️ Sales Report - Skipped", 'info')
            
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