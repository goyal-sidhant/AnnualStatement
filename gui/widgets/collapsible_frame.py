# gui/widgets/collapsible_frame.py
"""Collapsible Frame Widget"""

import tkinter as tk


class CollapsibleFrame(tk.Frame):
    """A collapsible frame widget"""
    def __init__(self, parent, title="", **kwargs):
        tk.Frame.__init__(self, parent, **kwargs)
        
        self.show = tk.IntVar()
        self.show.set(1)
        
        self.title_frame = tk.Frame(self)
        self.title_frame.pack(fill="x", expand=1)
        
        self.toggle_button = tk.Button(self.title_frame, 
                                     text="▼", 
                                     width=2,
                                     command=self.toggle,
                                     relief="flat",
                                     font=('Arial', 10))
        self.toggle_button.pack(side="left")
        
        tk.Label(self.title_frame, text=title, 
                font=('Arial', 11, 'bold')).pack(side="left", padx=5)
        
        self.sub_frame = tk.Frame(self, relief="sunken", borderwidth=1)
        self.sub_frame.pack(fill="both", expand=1, padx=10, pady=5)
        
    def toggle(self):
        """Toggle the collapsible frame"""
        if self.show.get():
            self.sub_frame.pack_forget()
            self.toggle_button.config(text="▶")
            self.show.set(0)
        else:
            self.sub_frame.pack(fill="both", expand=1, padx=10, pady=5)
            self.toggle_button.config(text="▼")
            self.show.set(1)