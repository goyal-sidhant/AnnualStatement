# gui/tabs/processing_tab.py
"""Processing Tab - Step 3 of GST File Organizer"""

import tkinter as tk
from tkinter import ttk, scrolledtext
from utils.constants import GUI_CONFIG


class ProcessingTab:
    """Processing tab for running file organization"""
    
    def __init__(self, notebook, app_instance):
        self.app = app_instance
        self.notebook = notebook
        self.create_tab()
    
    def create_tab(self):
        """Create processing tab with detailed logging"""
        self.tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_frame, text="🚀 Step 3: Processing")
        
        # Progress section
        progress_frame = tk.Frame(self.tab_frame, bg='white', relief='solid', borderwidth=2)
        progress_frame.pack(fill='x', padx=20, pady=20)
        
        header = tk.Frame(progress_frame, bg=GUI_CONFIG['colors']['primary'], height=40)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        tk.Label(header, text="📊 PROCESSING PROGRESS",
                font=('Arial', 12, 'bold'),
                bg=GUI_CONFIG['colors']['primary'],
                fg='white').pack(expand=True)
        
        content_frame = tk.Frame(progress_frame, bg='white')
        content_frame.pack(fill='x', padx=20, pady=20)
        
        # Progress bar
        self.app.progress_var = tk.DoubleVar()
        self.app.progress_bar = ttk.Progressbar(content_frame,
                                          variable=self.app.progress_var,
                                          mode='determinate')
        self.app.progress_bar.pack(fill='x', pady=(0, 10))
        
        # Progress label
        self.app.progress_label = tk.Label(content_frame,
                                     text="⏳ Ready to start...",
                                     font=('Arial', 11, 'bold'),
                                     bg='white',
                                     fg=GUI_CONFIG['colors']['dark'])
        self.app.progress_label.pack(anchor='w', pady=(0, 5))
        
        self.app.current_operation = tk.Label(content_frame,
                                        text="",
                                        font=('Arial', 10),
                                        bg='white',
                                        fg=GUI_CONFIG['colors']['primary'])
        self.app.current_operation.pack(anchor='w')
        
        # Control buttons
        control_frame = tk.Frame(content_frame, bg='white')
        control_frame.pack(fill='x', pady=(15, 0))
        
        self.app.start_btn = tk.Button(control_frame,
                                 text="🚀 START PROCESSING",
                                 font=('Arial', 10, 'bold'),
                                 bg=GUI_CONFIG['colors']['success'],
                                 fg='white',
                                 relief='flat',
                                 padx=20, 
                                 pady=8,
                                 cursor='hand2')
        self.app.start_btn.pack(side='left', padx=5)
        self.app.start_btn.bind('<Button-1>', lambda e: self.app.start_processing())
        
        self.app.stop_btn = tk.Button(control_frame,
                                text="⏹ STOP",
                                font=('Arial', 10, 'bold'),
                                bg=GUI_CONFIG['colors']['danger'],
                                fg='white',
                                relief='flat',
                                padx=20, 
                                pady=8,
                                state='disabled',
                                cursor='hand2')
        self.app.stop_btn.pack(side='left', padx=5)
        self.app.stop_btn.bind('<Button-1>', lambda e: self.app.stop_processing())
        
        # Log section
        log_frame = tk.Frame(self.tab_frame, bg='white', relief='solid', borderwidth=2)
        log_frame.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        header = tk.Frame(log_frame, bg=GUI_CONFIG['colors']['success'], height=40)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        tk.Label(header, text="📝 PROCESSING LOG",
                font=('Arial', 12, 'bold'),
                bg=GUI_CONFIG['colors']['success'],
                fg='white').pack(expand=True)
        
        # Log text with colors
        self.app.log_text = scrolledtext.ScrolledText(log_frame,
                                                 height=18,
                                                 font=('Consolas', 9),
                                                 bg='#1E1E1E',
                                                 fg='#FFFFFF',
                                                 state='disabled',
                                                 wrap=tk.WORD)
        self.app.log_text.pack(fill='both', expand=True, padx=15, pady=15)
        
        # Configure log tags
        self.app.log_text.tag_configure('success', foreground='#4CAF50')
        self.app.log_text.tag_configure('warning', foreground='#FF9800')
        self.app.log_text.tag_configure('error', foreground='#F44336')
        self.app.log_text.tag_configure('info', foreground='#2196F3')
        self.app.log_text.tag_configure('normal', foreground='#FFFFFF')