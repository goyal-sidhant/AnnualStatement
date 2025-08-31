"""
Standalone launcher for Power Query Extractor
Run this file separately after GST Organizer processing
"""

import sys
import os

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def main():
    try:
        from power_query_extractor.extractor_main import PowerQueryExtractor
        
        # Launch extractor - it will auto-load from cache
        extractor = PowerQueryExtractor()
        extractor.run()
        
    except Exception as e:
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", f"Failed to launch: {str(e)}")
        raise

if __name__ == "__main__":
    main()