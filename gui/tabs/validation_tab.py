# gui/tabs/validation_tab.py
"""Validation Tab - Step 2 of GST File Organizer"""

import tkinter as tk
from tkinter import ttk, scrolledtext
from utils.constants import GUI_CONFIG
from ..widgets.collapsible_frame import CollapsibleFrame


class ValidationTab:
    """Validation tab for client review and selection"""
    
    def __init__(self, notebook, app_instance):
        self.app = app_instance
        self.notebook = notebook
        self.create_tab()
    
    def create_tab(self):
        """Create validation tab with improved layout"""
        self.tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_frame, text="✅ Step 2: Validation")
        
        # Create main container
        main_container = tk.Frame(self.tab_frame, bg=GUI_CONFIG['colors']['light'])
        main_container.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Create two-column layout: Left (40%) for info, Right (60%) for client list
        left_frame = tk.Frame(main_container, bg=GUI_CONFIG['colors']['light'])
        left_frame.pack(side='left', fill='both', expand=False, padx=(0, 10))
        left_frame.config(width=450)
        left_frame.pack_propagate(False)
        
        right_frame = tk.Frame(main_container, bg=GUI_CONFIG['colors']['light'])
        right_frame.pack(side='right', fill='both', expand=True, padx=(10, 0))
        
        # Left side: Instructions and Summary
        self.create_collapsible_instructions(left_frame)
        self.create_summary_section(left_frame)
        self.create_validation_actions(left_frame)
        
        # Right side: Client Selection
        self.create_client_selection_section(right_frame)
    
    def create_collapsible_instructions(self, parent):
        """Create collapsible instructions"""
        # Create collapsible frame
        collapse_frame = CollapsibleFrame(parent, "📋 VALIDATION INSTRUCTIONS", 
                                        bg='#FFF3E0', relief='solid', borderwidth=2)
        collapse_frame.pack(fill='x', pady=(0, 15))
        
        # Add instructions to the sub_frame
        instructions = [
            "• Review the scan summary below",
            "• Check client list - ✅ = complete, ⚠️ = missing files",
            "• Select clients using click or keyboard (↑↓ + SPACE)",
            "• Try 'Dry Run' to preview without processing",
            "• Click 'Start Processing' when ready"
        ]
        
        for instruction in instructions:
            tk.Label(collapse_frame.sub_frame,
                    text=instruction,
                    font=('Arial', 10),
                    bg='#FFF3E0',
                    fg='#E65100',
                    anchor='w').pack(fill='x', padx=15, pady=2)
    
    def create_summary_section(self, parent):
        """Create scan summary section"""
        frame = tk.Frame(parent, bg='white', relief='solid', borderwidth=2)
        frame.pack(fill='x', pady=(0, 15))
        
        header = tk.Frame(frame, bg=GUI_CONFIG['colors']['primary'], height=40)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        tk.Label(header, text="📊 SCAN SUMMARY",
                font=('Arial', 12, 'bold'),
                bg=GUI_CONFIG['colors']['primary'],
                fg='white').pack(expand=True)
        
        # Create scrollable text widget with fixed height
        text_frame = tk.Frame(frame, bg='white', height=200)
        text_frame.pack(fill='x', padx=15, pady=15)
        text_frame.pack_propagate(False)
        
        self.app.summary_text = scrolledtext.ScrolledText(text_frame, 
                                                     font=('Consolas', 9),
                                                     bg='#f8f9fa',
                                                     state='disabled',
                                                     wrap=tk.WORD)
        self.app.summary_text.pack(fill='both', expand=True)
    
    def create_validation_actions(self, parent):
        """Create validation action buttons"""
        frame = tk.Frame(parent, bg='white', relief='solid', borderwidth=2)
        frame.pack(fill='x')
        
        tk.Label(frame,
                text="🎯 READY TO PROCESS?",
                font=('Arial', 14, 'bold'),
                bg='white',
                fg=GUI_CONFIG['colors']['success']).pack(pady=(15, 10))
        
        button_frame = tk.Frame(frame, bg='white')
        button_frame.pack(pady=(0, 15))
        
        # Dry run button
        dry_btn = tk.Button(button_frame, text="🧪 DRY RUN",
                           font=('Arial', 10, 'bold'),
                           bg=GUI_CONFIG['colors']['warning'], 
                           fg='white',
                           relief='flat', 
                           padx=15, 
                           pady=5,
                           cursor='hand2')
        dry_btn.pack(side='left', padx=5)
        dry_btn.bind('<Button-1>', lambda e: self.app.dry_run())

        # Export button
        export_btn = tk.Button(button_frame, text="📋 EXPORT LIST",
                            font=('Arial', 10, 'bold'),
                            bg=GUI_CONFIG['colors']['info'], 
                            fg='white',
                            relief='flat', 
                            padx=15, 
                            pady=5,
                            cursor='hand2')
        export_btn.pack(side='left', padx=5)
        export_btn.bind('<Button-1>', lambda e: self.app.export_client_list())
        
        # Start processing button
        start_btn = tk.Button(button_frame, text="🚀 START",
                             font=('Arial', 10, 'bold'),
                             bg=GUI_CONFIG['colors']['success'], 
                             fg='white',
                             relief='flat', 
                             padx=15, 
                             pady=5,
                             cursor='hand2')
        start_btn.pack(side='left', padx=5)
        start_btn.bind('<Button-1>', lambda e: self.app.start_processing())
    
    def create_client_selection_section(self, parent):
        """Create client selection with keyboard support"""
        frame = tk.Frame(parent, bg='white', relief='solid', borderwidth=2)
        frame.pack(fill='both', expand=True)
        
        header = tk.Frame(frame, bg=GUI_CONFIG['colors']['success'], height=40)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        tk.Label(header, text="👥 CLIENT SELECTION",
                font=('Arial', 12, 'bold'),
                bg=GUI_CONFIG['colors']['success'],
                fg='white').pack(expand=True)
        
        # Selection buttons
        button_frame = tk.Frame(frame, bg='white')
        button_frame.pack(fill='x', padx=15, pady=15)
        
        tk.Label(button_frame,
                text="⌨️ Use ↑↓ arrows + SPACE to select, or click with mouse",
                font=('Arial', 9),
                bg='#E3F2FD',
                fg=GUI_CONFIG['colors']['primary'],
                relief='solid',
                borderwidth=1,
                pady=5).pack(fill='x', pady=(0, 10))
        
        action_frame = tk.Frame(button_frame, bg='white')
        action_frame.pack()
        
        buttons = [
            ("☑️ Select All", self.app.select_all_clients, GUI_CONFIG['colors']['success']),
            ("☐ Clear All", self.app.clear_all_clients, GUI_CONFIG['colors']['dark']),
            ("✅ Complete Only", self.app.select_complete_clients, GUI_CONFIG['colors']['primary'])
        ]
        
        for text, command, color in buttons:
            btn = tk.Button(action_frame, text=text,
                           font=('Arial', 10, 'bold'),
                           bg=color, 
                           fg='white',
                           relief='flat', 
                           padx=15, 
                           pady=5,
                           cursor='hand2')
            btn.pack(side='left', padx=5)
            btn.bind('<Button-1>', lambda e, cmd=command: cmd())
        
        # Client tree
        tree_frame = tk.Frame(frame, bg='white')
        tree_frame.pack(fill='both', expand=True, padx=15, pady=(0, 15))
        
        # Create treeview
        columns = ('Client', 'State', 'Status', 'Files', 'Missing', 'Extra','FolderName')
        self.app.client_tree = ttk.Treeview(tree_frame,
                                       columns=columns,
                                       show='tree headings',
                                       height=20)
        
        # Configure columns
        self.app.client_tree.column('#0', width=50)
        self.app.client_tree.column('Client', width=250)
        self.app.client_tree.column('State', width=100)
        self.app.client_tree.column('Status', width=150)
        self.app.client_tree.column('Files', width=60)
        self.app.client_tree.column('Missing', width=180)
        self.app.client_tree.column('Extra', width=80)
        self.app.client_tree.column('FolderName', width=100, minwidth=80)
        
        # Configure headings
        self.app.client_tree.heading('#0', text='✓')
        self.app.client_tree.heading('Client', text='🏢 Client Name')
        self.app.client_tree.heading('State', text='🌏 State')
        self.app.client_tree.heading('Status', text='📊 Status')
        self.app.client_tree.heading('Files', text='📁 Files')
        self.app.client_tree.heading('Missing', text='❌ Missing')
        self.app.client_tree.heading('Extra', text='➕ Extra')
        self.app.client_tree.heading('FolderName', text='📁 Name')
        
        # Scrollbar
        v_scroll = ttk.Scrollbar(tree_frame, orient='vertical',
                                command=self.app.client_tree.yview)
        self.app.client_tree.configure(yscrollcommand=v_scroll.set)

        # Horizontal scrollbar
        h_scroll = ttk.Scrollbar(tree_frame, orient='horizontal',
                                command=self.app.client_tree.xview)
        self.app.client_tree.configure(xscrollcommand=h_scroll.set)
        
        self.app.client_tree.grid(row=0, column=0, sticky='nsew')
        v_scroll.grid(row=0, column=1, sticky='ns')
        h_scroll.grid(row=1, column=0, sticky='ew')

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Bind events for keyboard and mouse
        self.app.client_tree.bind('<Button-1>', self.app.on_client_click)
        self.app.client_tree.bind('<space>', self.app.on_space_key)
        self.app.client_tree.bind('<Up>', self.app.on_up_key)
        self.app.client_tree.bind('<Down>', self.app.on_down_key)
        
        # Set focus
        self.app.client_tree.focus_set()