# gui/widgets/title_bar.py
"""Title Bar Widget with Dark Mode Toggle"""

import tkinter as tk
from utils.constants import GUI_CONFIG


class TitleBar:
    """Title bar with dark mode toggle"""
    
    def __init__(self, parent, dark_mode_var, toggle_callback):
        self.parent = parent
        self.dark_mode_var = dark_mode_var
        self.toggle_callback = toggle_callback
        self.create_title_bar()
    
    def create_title_bar(self):
        """Create colorful title bar"""
        colors = GUI_CONFIG['colors']
        
        self.title_frame = tk.Frame(self.parent, bg=colors['primary'], height=70)
        self.title_frame.pack(fill='x')
        self.title_frame.pack_propagate(False)
        
        # Dark mode toggle
        dark_mode_frame = tk.Frame(self.title_frame, bg=GUI_CONFIG['colors']['primary'])
        dark_mode_frame.place(x=self.parent.winfo_width()-150, y=15)

        self.dark_mode_check = tk.Checkbutton(dark_mode_frame,
                                      text="🌙 Dark Mode",
                                      variable=self.dark_mode_var,
                                      command=self.toggle_callback,
                                      font=('Arial', 10, 'bold'),
                                      bg=GUI_CONFIG['colors']['primary'],
                                      fg='white',
                                      selectcolor=GUI_CONFIG['colors']['primary'])
        self.dark_mode_check.pack()
        
        title_label = tk.Label(self.title_frame,
                              text="🏢 GST FILE ORGANIZER & REPORT GENERATOR",
                              font=('Arial', 18, 'bold'),
                              bg=colors['primary'],
                              fg='white')
        title_label.pack(expand=True)
        
        subtitle_label = tk.Label(self.title_frame,
                                 text="Organize files and generate Excel reports automatically",
                                 font=('Arial', 11),
                                 bg=colors['primary'],
                                 fg='#E3F2FD')
        subtitle_label.pack()