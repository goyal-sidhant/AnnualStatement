# gui/utils/dark_mode_manager.py
"""Dark Mode Manager - Handles theme switching with proper state restoration"""

import tkinter as tk
from tkinter import ttk
from utils.constants import GUI_CONFIG


class DarkModeManager:
    """Manages dark mode toggle with proper color restoration"""
    
    def __init__(self, root, style):
        self.root = root
        self.style = style
        self.original_colors = {}
        self.is_initialized = False
        self.is_dark_mode = False  # Track dark mode state
        
        # Dark mode colors
        self.dark_colors = {
            'bg': '#2b2b2b',
            'fg': '#ffffff',
            'widget_bg': '#3b3b3b',
            'button_bg': '#404040'
        }
    
    def initialize(self):
        """Store original colors on first run"""
        if not self.is_initialized:
            self._store_original_colors()
            self.is_initialized = True
    
    def _store_original_colors(self):
        """Store original widget colors before any changes"""
        # Store ttk style colors
        self.original_colors['ttk'] = {
            'TFrame': {
                'background': self.style.lookup('TFrame', 'background')
            },
            'TLabel': {
                'background': self.style.lookup('TLabel', 'background'),
                'foreground': self.style.lookup('TLabel', 'foreground')
            },
            'TNotebook': {
                'background': self.style.lookup('TNotebook', 'background')
            },
            'TNotebook.Tab': {
                'background': self.style.lookup('TNotebook.Tab', 'background'),
                'foreground': self.style.lookup('TNotebook.Tab', 'foreground')
            },
            'Treeview': {
                'background': self.style.lookup('Treeview', 'background'),
                'foreground': self.style.lookup('Treeview', 'foreground'),
                'fieldbackground': self.style.lookup('Treeview', 'fieldbackground')
            },
            'Treeview.Heading': {
                'background': self.style.lookup('Treeview.Heading', 'background'),
                'foreground': self.style.lookup('Treeview.Heading', 'foreground')
            }
        }
        
        # Store root background
        self.original_colors['root_bg'] = self.root.cget('bg')
    
    def apply_dark_mode(self):
        """Apply dark mode theme"""
        # Configure ttk styles for dark mode
        self.is_dark_mode = True  # Update state
        self.style.configure('TFrame', background=self.dark_colors['bg'])
        self.style.configure('TLabel', background=self.dark_colors['bg'], 
                           foreground=self.dark_colors['fg'])
        self.style.configure('TNotebook', background=self.dark_colors['bg'])
        self.style.configure('TNotebook.Tab', background=self.dark_colors['widget_bg'], 
                           foreground=self.dark_colors['fg'])
        self.style.map('TNotebook.Tab', 
                      background=[('selected', self.dark_colors['button_bg'])])
        
        # Configure Treeview for dark mode
        self.style.configure('Treeview', 
                           background=self.dark_colors['widget_bg'], 
                           foreground=self.dark_colors['fg'],
                           fieldbackground=self.dark_colors['widget_bg'])
        self.style.configure('Treeview.Heading', 
                           background=self.dark_colors['button_bg'], 
                           foreground=self.dark_colors['fg'])
        
        # Root window
        self.root.configure(bg=self.dark_colors['bg'])
        
        # Apply to all widgets
        self._apply_theme_to_widgets(self.root, is_dark=True)
    
    def restore_original_colors(self):
        """Restore original colors when dark mode is turned off"""
        # Restore ttk styles
        self.is_dark_mode = False  # Update state
        for widget_class, properties in self.original_colors.get('ttk', {}).items():
            for property_name, value in properties.items():
                if value is not None:
                    self.style.configure(widget_class, **{property_name: value})
        # Add map restoration for TNotebook.Tab
        self.style.map('TNotebook.Tab', 
                    background=[('selected', '#f0f0f0')],
                    foreground=[('selected', 'black')])

        # Force refresh of Treeview text colors
        self.style.map('Treeview',
                    foreground=[('selected', 'black')],
                    background=[('selected', '#0078D4')])
        
        # Restore root
        self.root.configure(bg=self.original_colors.get('root_bg', 'SystemButtonFace'))
        
        # Restore all widgets
        self._apply_theme_to_widgets(self.root, is_dark=False)
    
    def _apply_theme_to_widgets(self, widget, is_dark=True):
        """Recursively apply theme to all widgets"""
        try:
            widget_class = widget.winfo_class()
            
            if is_dark:
                self._apply_dark_to_widget(widget, widget_class)
            else:
                self._restore_widget_colors(widget, widget_class)
            
        except:
            pass
        
        # Process children
        for child in widget.winfo_children():
            self._apply_theme_to_widgets(child, is_dark)
    
    def _apply_dark_to_widget(self, widget, widget_class):
        """Apply dark mode to a single widget"""
        if widget_class == 'Frame' or widget_class == 'Toplevel':
            widget.configure(bg=self.dark_colors['bg'])
        elif widget_class == 'Label':
            # Store original bg if not stored
            if not hasattr(widget, '_original_bg'):
                widget._original_bg = widget.cget('bg')
            current_bg = str(widget.cget('bg'))
            if current_bg in ['white', '#ffffff', 'SystemButtonFace', GUI_CONFIG['colors']['light']]:
                widget.configure(bg=self.dark_colors['bg'], fg=self.dark_colors['fg'])
        elif widget_class == 'Button':
            # Store original color before changing
            if not hasattr(widget, '_original_bg'):
                widget._original_bg = widget.cget('bg')
            widget.configure(bg=self.dark_colors['button_bg'], 
                           fg=self.dark_colors['fg'], 
                           activebackground=self.dark_colors['widget_bg'])
        elif widget_class == 'Entry':
            widget.configure(bg=self.dark_colors['widget_bg'], 
                           fg=self.dark_colors['fg'], 
                           insertbackground=self.dark_colors['fg'])
        elif widget_class == 'Text':
            widget.configure(bg=self.dark_colors['widget_bg'], 
                           fg=self.dark_colors['fg'], 
                           insertbackground=self.dark_colors['fg'])
        elif widget_class in ['Checkbutton', 'Radiobutton']:
            widget.configure(bg=self.dark_colors['bg'], 
                           fg=self.dark_colors['fg'], 
                           activebackground=self.dark_colors['bg'], 
                           selectcolor=self.dark_colors['widget_bg'])
        elif widget_class == 'Canvas':
            widget.configure(bg=self.dark_colors['bg'])
    
    def _restore_widget_colors(self, widget, widget_class):
        """Restore original colors to a widget"""
        try:
            if widget_class in ['Frame', 'Canvas']:
                widget.configure(bg='white')
            elif widget_class == 'Label':
                current_bg = str(widget.cget('bg'))
                if current_bg in ['#2b2b2b', '#3b3b3b', '#404040']:
                    if hasattr(widget, '_original_bg'):
                        widget.configure(bg=widget._original_bg, fg='black')
                    else:
                        widget.configure(bg='white', fg='black')
            elif widget_class == 'Button':
                if hasattr(widget, '_original_bg'):
                    widget.configure(bg=widget._original_bg, fg='white')
                else:
                    widget.configure(fg='white')
            elif widget_class in ['Entry', 'Text']:
                widget.configure(bg='white', fg='black', insertbackground='black')
            elif widget_class in ['Checkbutton', 'Radiobutton']:
                widget.configure(bg='white', fg='black', selectcolor='white')
        except:
            pass
    
    def update_tree_tags(self, tree_widget, is_dark_mode=None):
        """Update tree widget tags for current theme"""
        # Use parameter if provided, otherwise use internal state
        if is_dark_mode is None:
            is_dark_mode = self.is_dark_mode
            
        if hasattr(tree_widget, 'tag_configure'):
            if is_dark_mode:
                tree_widget.tag_configure('complete', 
                                        background='#1a3d1a', 
                                        foreground=self.dark_colors['fg'])
                tree_widget.tag_configure('incomplete', 
                                        background='#3d2f1a', 
                                        foreground=self.dark_colors['fg'])
        else:
            tree_widget.tag_configure('complete', 
                                    background=GUI_CONFIG['colors']['complete'],
                                    foreground='black')  # Add this
            tree_widget.tag_configure('incomplete', 
                                    background=GUI_CONFIG['colors']['incomplete'],
                                    foreground='black')  # Add this