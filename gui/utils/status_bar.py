# gui/utils/status_bar.py
"""Status Bar Widget"""

import tkinter as tk
from utils.constants import GUI_CONFIG


class StatusBar:
    """Status bar for the application"""
    
    def __init__(self, parent):
        self.parent = parent
        self.create_status_bar()
    
    def create_status_bar(self):
        """Create status bar"""
        status_frame = tk.Frame(self.parent, bg=GUI_CONFIG['colors']['dark'], height=30)
        status_frame.pack(fill='x', side='bottom')
        status_frame.pack_propagate(False)
        
        self.status_label = tk.Label(status_frame,
                                   text="💡 Ready to organize your GST files!",
                                   font=('Arial', 9),
                                   bg=GUI_CONFIG['colors']['dark'],
                                   fg='white')
        self.status_label.pack(side='left', padx=10, pady=6)
        
        version_label = tk.Label(status_frame,
                               text="v3.0 | Production Ready",
                               font=('Arial', 9),
                               bg=GUI_CONFIG['colors']['dark'],
                               fg='#BDBDBD')
        version_label.pack(side='right', padx=10, pady=6)
    
    def update_status(self, message):
        """Update status message"""
        self.status_label.config(text=f"💡 {message}")