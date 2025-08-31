# gui/main_window.py
"""
Main GUI Application Window for GST File Organizer v3.0
Modularized version - ties all components together
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
import logging

# Import core modules (these imports remain the same)
from core.file_parser import FileParser
from core.file_organizer import FileOrganizer
from core.excel_handler import ExcelHandler

# Import utilities
from utils.constants import GUI_CONFIG, MESSAGES
from utils.helpers import get_timestamp

# Import GUI components
from .widgets.title_bar import TitleBar
from .widgets.collapsible_frame import CollapsibleFrame
from .tabs.setup_tab import SetupTab
from .tabs.validation_tab import ValidationTab
from .tabs.processing_tab import ProcessingTab
from .handlers.file_handler import FileHandler
from .handlers.client_handler import ClientHandler
from .handlers.processing_handler import ProcessingHandler
from .utils.dark_mode_manager import DarkModeManager
from .utils.cache_manager import CacheManager
from .utils.ui_helpers import UIHelpers
from .utils.status_bar import StatusBar

logger = logging.getLogger(__name__)


class GSTOrganizerApp:
    """
    Main application GUI - modularized version.
    """
    
    def __init__(self):
        """Initialize the application"""
        self.root = tk.Tk()
        self.setup_window()
        self.setup_variables()
        self.setup_components()
        self.setup_handlers()
        self.create_gui()
        self.bind_events()
        self.apply_cached_values()
        
    def setup_window(self):
        """Configure main window"""
        self.root.title(GUI_CONFIG['window']['title'])
        self.root.geometry(GUI_CONFIG['window']['size'])
        self.root.minsize(*GUI_CONFIG['window']['min_size'])
        
        # Set background color
        self.root.configure(bg=GUI_CONFIG['colors']['light'])
        
        # Center window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
        
        # Configure style
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam')
        except:
            pass
    
    def setup_variables(self):
        """Initialize application variables"""
        # GUI variables
        self.source_folder = tk.StringVar()
        self.itc_template = tk.StringVar()
        self.sales_template = tk.StringVar()
        self.client_folder_settings = {}
        self.target_folder = tk.StringVar()
        self.processing_mode = tk.StringVar(value='fresh')
        self.dark_mode = tk.BooleanVar(value=False)
        self.include_client_name_in_folders = tk.BooleanVar(value=False)
        self.client_name_max_length = tk.IntVar(value=35)
        
        # Data storage
        self.scanned_files = {}
        self.client_data = {}
        self.variations = []
        self.selected_clients = {}
        
        # Processing state
        self.is_processing = False
        self.stop_requested = False
        self.processing_thread = None
        
        # GUI references
        self.widgets = {}
        
        # Core component imports (needed by handlers)
        self.FileOrganizer = FileOrganizer
        
    def setup_components(self):
        """Initialize core components and managers"""
        # Core components
        self.file_parser = FileParser()
        self.excel_handler = ExcelHandler()
        self.file_organizer = None
        
        # Managers
        self.cache_manager = CacheManager()
        self.cache_data = self.cache_manager.load_cache()
        self.dark_mode_manager = DarkModeManager(self.root, self.style)
        self.dark_mode_manager.initialize()
    
    def setup_handlers(self):
        """Initialize handlers"""
        self.file_handler = FileHandler(self)
        self.client_handler = ClientHandler(self)
        self.processing_handler = ProcessingHandler(self)
    
    def create_gui(self):
        """Create the complete GUI interface"""
        # Create title bar
        self.title_bar = TitleBar(self.root, self.dark_mode, self.toggle_dark_mode)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=20, pady=(10, 20))
        
        # Create tabs
        self.setup_tab = SetupTab(self.notebook, self)
        self.validation_tab = ValidationTab(self.notebook, self)
        self.processing_tab = ProcessingTab(self.notebook, self)
        
        # Create status bar
        self.status_bar = StatusBar(self.root)
    
    def bind_events(self):
        """Bind application events"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
    
    # Dark mode methods
    def toggle_dark_mode(self):
        """Toggle dark mode on/off"""
        if self.dark_mode.get():
            self.dark_mode_manager.apply_dark_mode()
        else:
            self.dark_mode_manager.restore_original_colors()
        
        # Update special widgets
        if hasattr(self, 'title_bar'):
            # Restore title frame color
            if not self.dark_mode.get() and hasattr(self.title_bar, 'title_frame'):
                self.title_bar.title_frame.configure(bg=GUI_CONFIG['colors']['primary'])
                # Update all children of title frame
                for widget in self.title_bar.title_frame.winfo_children():
                    try:
                        widget.configure(bg=GUI_CONFIG['colors']['primary'])
                    except:
                        pass
        
        # Update tree tags
        if hasattr(self, 'client_tree'):
            self.dark_mode_manager.update_tree_tags(self.client_tree, self.dark_mode.get())
        
        self.save_cache()
        self.root.update_idletasks()
        
        # Force refresh of notebook tabs
        if hasattr(self, 'notebook'):
            self.notebook.update_idletasks()
            # Trigger style update
            self.style.configure('TNotebook.Tab', padding=[20, 10])

    # File operation methods (delegated to handlers)
    def browse_source_folder(self):
        self.file_handler.browse_source_folder()
    
    def browse_itc_template(self):
        self.file_handler.browse_itc_template()
    
    def browse_sales_template(self):
        self.file_handler.browse_sales_template()
    
    def browse_target_folder(self):
        self.file_handler.browse_target_folder()
    
    def scan_files(self):
        self.file_handler.scan_files()
    
    def rescan_files(self):
        self.file_handler.rescan_files()
    
    # Client operation methods (delegated to handlers)
    def update_client_tree(self):
        self.client_handler.update_client_tree()
    
    def on_client_click(self, event):
        self.client_handler.on_client_click(event)
    
    def on_space_key(self, event):
        return self.client_handler.on_space_key(event)
    
    def on_up_key(self, event):
        return self.client_handler.on_up_key(event)
    
    def on_down_key(self, event):
        return self.client_handler.on_down_key(event)
    
    def toggle_client_selection(self, item):
        self.client_handler.toggle_client_selection(item)
    
    def select_all_clients(self):
        self.client_handler.select_all_clients()
    
    def clear_all_clients(self):
        self.client_handler.clear_all_clients()
    
    def select_complete_clients(self):
        self.client_handler.select_complete_clients()
    
    def toggle_folder_name_setting(self, item):
        self.client_handler.toggle_folder_name_setting(item)
    
    def update_global_folder_setting(self):
        self.client_handler.update_global_folder_setting()
    
    def export_client_list(self):
        self.client_handler.export_client_list()
    
    # Processing methods (delegated to handlers)
    def dry_run(self):
        self.processing_handler.dry_run()
    
    def start_processing(self):
        self.processing_handler.start_processing()
    
    def stop_processing(self):
        self.processing_handler.stop_processing()
    
    def process_files_thread(self, selected_clients):
        self.processing_handler.process_files_thread(selected_clients)
    
    # UI update methods
    def update_status(self, message):
        """Update status bar"""
        self.status_bar.update_status(message)
    
    def log_message(self, message, level='normal'):
        """Add message to log"""
        UIHelpers.log_message(self.log_text, message, level)
    
    def update_progress(self, value, message):
        """Update progress display"""
        self.root.after(0, lambda: UIHelpers.update_progress(
            self.progress_var, self.progress_label, 
            self.current_operation, value, message
        ))
    
    def update_summary(self):
        """Update summary display"""
        if not hasattr(self, 'summary_text'):
            return
        
        stats = self.file_parser.get_statistics()
        
        summary = f"""📊 SCAN RESULTS SUMMARY:
════════════════════════════════════════════
📁 FILES ANALYSIS:
   Total Excel Files Found: {stats['total_files']}
   ✅ Successfully Parsed: {stats['parsed_files']}
   ❌ Unparsed/Variations: {stats['unparsed_files']}
   📊 Parsing Success Rate: {stats['parsing_rate']:.1f}%

👥 CLIENTS ANALYSIS:
   Total Clients Identified: {stats['total_clients']}
   ✅ Clients with All Files: {stats['complete_clients']}
   ⚠️ Clients with Missing Files: {stats['incomplete_clients']}
   📈 Completeness Rate: {stats['completion_rate']:.1f}%

📋 FILE TYPE DISTRIBUTION:"""
        
        for file_type, count in stats['file_type_distribution'].items():
            summary += f"\n   • {file_type}: {count} files"
        
        if self.variations:
            summary += f"\n\n🔍 FILE VARIATIONS: {len(self.variations)} files"
        
        summary += f"\n\n💾 TOTAL DATA SIZE: {stats['total_size_formatted']}"
        
        self.summary_text.config(state='normal')
        self.summary_text.delete(1.0, tk.END)
        self.summary_text.insert(1.0, summary)
        self.summary_text.config(state='disabled')
    
    # Event handlers
    def on_tab_changed(self, event):
        """Handle tab change"""
        current = self.notebook.index('current')
        tabs = ['Setup', 'Validation', 'Processing']
        self.update_status(f"Step {current + 1}: {tabs[current]}")
        
        # Force UI update when switching tabs
        if not self.is_processing:
            self.root.update_idletasks()
    
    def on_closing(self):
        """Handle window closing"""
        if self.is_processing:
            result = messagebox.askyesno(
                "Processing Active",
                "Processing is still running. Stop and exit?"
            )
            if result:
                self.stop_requested = True
                self.root.after(1000, self.root.destroy)
        else:
            self.root.destroy()
    
    # Cache methods
    def save_cache(self):
        """Save current settings"""
        cache_data = {
            'source_folder': self.source_folder.get(),
            'itc_template': self.itc_template.get(),
            'sales_template': self.sales_template.get(),
            'target_folder': self.target_folder.get(),
            'processing_mode': self.processing_mode.get(),
            'include_client_name': self.include_client_name_in_folders.get(),
            'dark_mode': self.dark_mode.get(),
            'client_name_max_length': self.client_name_max_length.get()
        }
        self.cache_manager.save_cache(cache_data)
    
    def apply_cached_values(self):
        """Apply cached values"""
        if self.cache_data:
            self.source_folder.set(self.cache_data.get('source_folder', ''))
            self.itc_template.set(self.cache_data.get('itc_template', ''))
            self.sales_template.set(self.cache_data.get('sales_template', ''))
            self.target_folder.set(self.cache_data.get('target_folder', ''))
            self.processing_mode.set(self.cache_data.get('processing_mode', 'fresh'))
            self.include_client_name_in_folders.set(self.cache_data.get('include_client_name', False))
            self.dark_mode.set(self.cache_data.get('dark_mode', False))
            # Set client name max length with validation (default 35 if 0 or invalid)
            cached_length = self.cache_data.get('client_name_max_length', 35)
            self.client_name_max_length.set(cached_length if cached_length > 0 else 35)
            if self.dark_mode.get():
                self.toggle_dark_mode()
    
    def run(self):
        """Run the application"""
        self.update_status("Welcome to GST File Organizer v3.0")
        self.log_message("🎉 Application started", 'info')
        self.log_message("💡 Ready to organize your GST files!", 'success')
        self.log_message("⚠️ IMPORTANT: Templates must have 'Links' sheet for reports to work!", 'warning')
        
        self.root.mainloop()