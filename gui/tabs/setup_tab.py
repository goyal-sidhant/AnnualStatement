# gui/tabs/setup_tab.py
"""Setup Tab - Step 1 of GST File Organizer"""

import tkinter as tk
from tkinter import ttk
from utils.constants import GUI_CONFIG, PROCESSING_MODES
from ..widgets.collapsible_frame import CollapsibleFrame
from ..utils.ui_helpers import UIHelpers


class SetupTab:
    """Setup tab for folder and template selection"""
    
    def __init__(self, notebook, app_instance):
        self.app = app_instance
        self.notebook = notebook
        self.create_tab()
    
    def create_tab(self):
        """Create setup tab with working mouse events"""
        self.tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_frame, text="📁 Step 1: Setup")
        
        # Create scrollable frame
        canvas = tk.Canvas(self.tab_frame, bg=GUI_CONFIG['colors']['light'])
        scrollbar = ttk.Scrollbar(self.tab_frame, orient='vertical', command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=GUI_CONFIG['colors']['light'])
        
        scrollable_frame.bind('<Configure>',
                             lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Create two-column layout
        main_container = tk.Frame(scrollable_frame, bg=GUI_CONFIG['colors']['light'])
        main_container.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Left column (60%)
        left_column = tk.Frame(main_container, bg=GUI_CONFIG['colors']['light'])
        left_column.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # Right column (40%)
        right_column = tk.Frame(main_container, bg=GUI_CONFIG['colors']['light'], width=400)
        right_column.pack(side='right', fill='y', padx=(10, 0))
        right_column.pack_propagate(False)
        
        # Create sections in left column
        self.create_welcome_section(left_column)
        self.create_source_folder_section(left_column)
        self.create_templates_section(left_column)
        self.create_target_folder_section(left_column)
        self.create_processing_mode_section(left_column)
        self.create_action_buttons_section(left_column)
        
        # Create help section in right column
        self.create_help_section(right_column)
        
        # Pack canvas and scrollbar
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Bind mouse wheel to canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def create_welcome_section(self, parent):
        """Create welcome section"""
        colors = GUI_CONFIG['colors']
        
        frame = tk.Frame(parent, bg='white', relief='solid', borderwidth=2)
        frame.pack(fill='x', pady=(0, 15))
        
        tk.Label(frame,
                text="🎉 Welcome to GST File Organizer!",
                font=('Arial', 14, 'bold'),
                bg='white',
                fg=colors['primary']).pack(pady=(15, 5))
        
        tk.Label(frame,
                text="This tool will help you organize GST files and create Excel reports automatically.\nLet's get started!",
                font=('Arial', 10),
                bg='white',
                fg=colors['dark'],
                justify='center').pack(pady=(0, 15))
    
    def create_source_folder_section(self, parent):
        """Create source folder selection with working browse button"""
        section = UIHelpers.create_colored_section(
            parent, "📂 SOURCE FOLDER", 
            "Select the folder containing your GST Excel files",
            GUI_CONFIG['colors']['success']
        )
        
        input_frame = tk.Frame(section, bg='white')
        input_frame.pack(fill='x', padx=15, pady=10)
        
        entry = tk.Entry(input_frame, textvariable=self.app.source_folder,
                        font=('Arial', 10), relief='solid', borderwidth=1)
        entry.pack(side='left', fill='x', expand=True, ipady=5)
        
        # Create button with proper event binding
        browse_btn = tk.Button(input_frame, text="📂 Browse",
                              font=('Arial', 10, 'bold'),
                              bg=GUI_CONFIG['colors']['success'], 
                              fg='white',
                              relief='flat', 
                              padx=15, 
                              pady=5,
                              cursor='hand2')
        browse_btn.pack(side='right', padx=(10, 0))
        
        # Bind click event
        browse_btn.bind('<Button-1>', lambda e: self.app.browse_source_folder())
    
    def create_templates_section(self, parent):
        """Create templates selection section with working browse buttons"""
        section = UIHelpers.create_colored_section(
            parent, "📋 EXCEL TEMPLATES",
            "Select your Excel template files (.xlsx or .xltx)",
            GUI_CONFIG['colors']['primary']
        )
        
        # ITC Template
        self.create_template_input(section, "ITC Report Template:",
                                 self.app.itc_template, self.app.browse_itc_template)
        
        # Sales Template
        self.create_template_input(section, "Sales Report Template:",
                                 self.app.sales_template, self.app.browse_sales_template)
    
    def create_template_input(self, parent, label, variable, command):
        """Create template input field with working browse button"""
        frame = tk.Frame(parent, bg='white')
        frame.pack(fill='x', padx=15, pady=5)
        
        tk.Label(frame, text=label,
                font=('Arial', 10, 'bold'),
                bg='white').pack(anchor='w')
        
        input_frame = tk.Frame(frame, bg='white')
        input_frame.pack(fill='x', pady=(5, 0))
        
        entry = tk.Entry(input_frame, textvariable=variable,
                        font=('Arial', 10), relief='solid', borderwidth=1)
        entry.pack(side='left', fill='x', expand=True, ipady=5)
        
        # Create button with proper event binding
        browse_btn = tk.Button(input_frame, text="📋 Browse",
                              font=('Arial', 9, 'bold'),
                              bg=GUI_CONFIG['colors']['primary'], 
                              fg='white',
                              relief='flat', 
                              padx=12, 
                              pady=5,
                              cursor='hand2')
        browse_btn.pack(side='right', padx=(10, 0))
        
        # Bind click event
        browse_btn.bind('<Button-1>', lambda e: command())
    
    def create_target_folder_section(self, parent):
        """Create target folder section with working browse button"""
        section = UIHelpers.create_colored_section(
            parent, "🎯 TARGET FOLDER ⭐ IMPORTANT!",
            "⚠️ Select where the organized files will be saved",
            GUI_CONFIG['colors']['danger']
        )
        
        # Add emphasis frame
        emphasis_frame = tk.Frame(section, bg='#FFEBEE', relief='solid', borderwidth=1)
        emphasis_frame.pack(fill='x', padx=15, pady=(0, 10))
        
        tk.Label(emphasis_frame,
                text="🚨 This is where your organized files will be created!",
                font=('Arial', 10, 'bold'),
                bg='#FFEBEE', fg='#C62828').pack(pady=5)
        
        input_frame = tk.Frame(section, bg='white')
        input_frame.pack(fill='x', padx=15, pady=10)
        
        entry = tk.Entry(input_frame, textvariable=self.app.target_folder,
                        font=('Arial', 10), relief='solid', borderwidth=2)
        entry.pack(side='left', fill='x', expand=True, ipady=8)
        
        # Create button with proper event binding
        browse_btn = tk.Button(input_frame, text="🎯 BROWSE",
                              font=('Arial', 10, 'bold'),
                              bg=GUI_CONFIG['colors']['danger'], 
                              fg='white',
                              relief='flat', 
                              padx=20, 
                              pady=8,
                              cursor='hand2')
        browse_btn.pack(side='right', padx=(10, 0))
        
        # Bind click event
        browse_btn.bind('<Button-1>', lambda e: self.app.browse_target_folder())
    
    def create_processing_mode_section(self, parent):
        """Create processing mode selection"""
        section = UIHelpers.create_colored_section(
            parent, "⚙️ PROCESSING MODE",
            "Choose how to organize files",
            GUI_CONFIG['colors']['info']
        )
        
        mode_frame = tk.Frame(section, bg='white')
        mode_frame.pack(fill='x', padx=15, pady=10)
        
        for mode_key, mode_info in PROCESSING_MODES.items():
            radio_frame = tk.Frame(mode_frame, bg='white')
            radio_frame.pack(fill='x', pady=3)
            
            tk.Radiobutton(radio_frame, text=mode_info['name'],
                          variable=self.app.processing_mode, value=mode_key,
                          font=('Arial', 10, 'bold'),
                          bg='white', fg=GUI_CONFIG['colors']['dark']).pack(anchor='w')
            
            tk.Label(radio_frame, text=f"   {mode_info['description']}",
                    font=('Arial', 9),
                    bg='white', fg='gray').pack(anchor='w')
        
        # Folder options
        options_frame = tk.Frame(mode_frame, bg='white')
        options_frame.pack(fill='x', pady=10)

        self.app.client_name_check = tk.Checkbutton(options_frame,
                                   text="Include client name in brackets for Level 4 folders",
                                   variable=self.app.include_client_name_in_folders,
                                   font=('Arial', 10),
                                   bg='white',
                                   command=self.app.update_global_folder_setting)
        self.app.client_name_check.pack(anchor='w')

        tk.Label(options_frame,
            text="⚠️ Names >10 chars will show warning\n"
                "ℹ️ This overrides individual client settings if checked",
            font=('Arial', 9, 'italic'),
            bg='white',
            fg='#FF8C00').pack(anchor='w', padx=20)
        
        # Client name length setting
        length_frame = tk.Frame(options_frame, bg='white')
        length_frame.pack(fill='x', pady=(10, 0))
        
        tk.Label(length_frame, text="Client folder name max length:",
                font=('Arial', 10),
                bg='white').pack(side='left')
        
        def validate_length(value):
            """Validate client name length input"""
            if value == "":
                return True
            try:
                num = int(value)
                return 10 <= num <= 100  # Reasonable range
            except ValueError:
                return False
        
        def on_length_change():
            """Handle length change and save to cache"""
            value = self.app.client_name_max_length.get()
            if value <= 0:
                self.app.client_name_max_length.set(35)  # Reset to default
            self.app.save_cache()
        
        vcmd = (self.app.root.register(validate_length), '%P')
        
        length_spinbox = tk.Spinbox(length_frame,
                                   from_=15, to=100, width=5,
                                   textvariable=self.app.client_name_max_length,
                                   validate='all', validatecommand=vcmd,
                                   command=on_length_change,
                                   font=('Arial', 10))
        length_spinbox.pack(side='left', padx=(10, 5))
        
        tk.Label(length_frame, text="(default: 35, range: 15-100)",
                font=('Arial', 9, 'italic'),
                bg='white', fg='gray').pack(side='left', padx=5)
    
    def create_action_buttons_section(self, parent):
        """Create action buttons with working events"""
        frame = tk.Frame(parent, bg='white', relief='solid', borderwidth=2)
        frame.pack(fill='x', pady=20)
        
        tk.Label(frame,
                text="🚀 READY TO SCAN?",
                font=('Arial', 14, 'bold'),
                bg='white', fg=GUI_CONFIG['colors']['success']).pack(pady=(15, 10))
        
        button_frame = tk.Frame(frame, bg='white')
        button_frame.pack(pady=(0, 15))
        
        # Scan button
        scan_btn = tk.Button(button_frame, text="🔍 SCAN FILES",
                            font=('Arial', 12, 'bold'),
                            bg=GUI_CONFIG['colors']['success'], 
                            fg='white',
                            relief='flat', 
                            padx=30, 
                            pady=10,
                            cursor='hand2')
        scan_btn.pack(side='left', padx=10)
        scan_btn.bind('<Button-1>', lambda e: self.app.scan_files())
        
        # Rescan button
        rescan_btn = tk.Button(button_frame, text="🔄 RE-SCAN",
                              font=('Arial', 12, 'bold'),
                              bg=GUI_CONFIG['colors']['warning'], 
                              fg='white',
                              relief='flat', 
                              padx=30, 
                              pady=10,
                              cursor='hand2')
        rescan_btn.pack(side='left', padx=10)
        rescan_btn.bind('<Button-1>', lambda e: self.app.rescan_files())
    
    def create_help_section(self, parent):
        """Create help section in right column"""
        colors = GUI_CONFIG['colors']
        
        # Instructions card
        instruction_card = tk.Frame(parent, bg='#E3F2FD', relief='solid', borderwidth=2)
        instruction_card.pack(fill='x', pady=(0, 20))
        
        tk.Label(instruction_card,
                text="📋 INSTRUCTIONS",
                font=('Arial', 12, 'bold'),
                bg='#E3F2FD',
                fg='#1565C0').pack(pady=(10, 5))
        
        instructions = [
            "1. 📂 Select folder with GST files",
            "2. 📋 Choose Excel templates",
            "3. 🎯 Select TARGET FOLDER",
            "4. ⚙️ Choose processing mode",
            "5. 🔍 Click 'Scan Files'"
        ]
        
        for instruction in instructions:
            tk.Label(instruction_card,
                    text=instruction,
                    font=('Arial', 10),
                    bg='#E3F2FD',
                    fg='#1565C0',
                    anchor='w').pack(fill='x', padx=10, pady=2)
        
        tk.Label(instruction_card, text="", bg='#E3F2FD').pack(pady=5)
        
        # File patterns card
        patterns_card = tk.Frame(parent, bg='#F3E5F5', relief='solid', borderwidth=2)
        patterns_card.pack(fill='x', pady=(0, 20))
        
        tk.Label(patterns_card,
                text="📝 EXPECTED FILE NAMES",
                font=('Arial', 12, 'bold'),
                bg='#F3E5F5',
                fg='#4A148C').pack(pady=(10, 5))
        
        patterns = [
            "GSTR-2B-Reco-Client-State-Period",
            "ImsReco-Client-State-DDMMYYYY",
            "GSTR3B-Client-State-Month",
            "Sales-Client-State-Start-End",
            "SalesReco-Client-State-Period",
            "AnnualReport-Client-State-Year"
        ]
        
        for pattern in patterns:
            tk.Label(patterns_card,
                    text=f"• {pattern}",
                    font=('Arial', 9),
                    bg='#F3E5F5',
                    fg='#4A148C',
                    anchor='w').pack(fill='x', padx=10, pady=1)