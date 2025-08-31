# gui/utils/ui_helpers.py
"""UI Helper Functions"""

import tkinter as tk
from datetime import datetime


class UIHelpers:
    """Helper methods for UI updates"""
    
    @staticmethod
    def update_status(status_label, message):
        """Update status bar message"""
        status_label.config(text=f"💡 {message}")
    
    @staticmethod
    def log_message(log_widget, message, level='normal'):
        """Add timestamped message to log widget"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        formatted_message = f"[{timestamp}] {message}\n"
        
        # Schedule the update
        log_widget.after(0, lambda: UIHelpers._add_log_message(
            log_widget, formatted_message, level
        ))
    
    @staticmethod
    def _add_log_message(log_widget, message, level):
        """Actually add the message to log widget"""
        log_widget.config(state='normal')
        log_widget.insert('end', message, level)
        log_widget.see('end')
        log_widget.config(state='disabled')
    
    @staticmethod
    def update_progress(progress_var, progress_label, current_op_label, value, message):
        """Update progress display"""
        progress_var.set(value)
        progress_label.config(text=f"📊 Progress: {value:.1f}%")
        current_op_label.config(text=message)
    
    @staticmethod
    def create_colored_section(parent, title, description, color):
        """Create a colored section frame"""
        section_frame = tk.Frame(parent, bg='white', relief='solid', borderwidth=2)
        section_frame.pack(fill='x', pady=10)
        
        # Header
        header_frame = tk.Frame(section_frame, bg=color, height=40)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text=title,
                font=('Arial', 12, 'bold'),
                bg=color, fg='white').pack(expand=True)
        
        # Description
        tk.Label(section_frame, text=description,
                font=('Arial', 10),
                bg='white', fg='#323130',
                wraplength=500).pack(pady=10)
        
        return section_frame