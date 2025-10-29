"""
GST File Organizer & Report Generator v3.0
Main entry point for the application.

Author: Advanced AI Assistant
Version: 3.0 (Production Ready)
"""

import sys
import io
import os
import logging
from pathlib import Path


# Set UTF-8 encoding for stdout
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('gst_organizer.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

def setup_environment():
    """Setup Python path and create necessary directories"""
    project_root = Path(__file__).parent
    sys.path.insert(0, str(project_root))
    
    # Create required directories
    for directory in ['core', 'utils', 'gui', 'logs', 'temp']:
        dir_path = project_root / directory
        dir_path.mkdir(exist_ok=True)
        
        # Create __init__.py for Python packages
        if directory in ['core', 'utils', 'gui']:
            init_file = dir_path / "__init__.py"
            if not init_file.exists():
                init_file.write_text('"""Package initialization"""')

def main():
    """Main entry point"""
    try:
        setup_environment()
        
        # Import after setup
        from gui.main_window import GSTOrganizerApp
        
        print("=" * 60)
        print("🚀 GST File Organizer v3.0 - Production Ready")
        print("=" * 60)
        
        app = GSTOrganizerApp()
        app.run()
        
    except ImportError as e:
        logging.error(f"Import Error: {e}")
        print("\n❌ Missing module! Please ensure all files are in place:")
        print("   📁 core/")
        print("      📄 file_parser.py")
        print("      📄 file_organizer.py")
        print("      📄 excel_handler.py")
        print("   📁 utils/")
        print("      📄 constants.py")
        print("      📄 helpers.py")
        print("   📁 gui/")
        print("      📄 app.py")
        input("\nPress Enter to exit...")
        sys.exit(1)
        
    except Exception as e:
        logging.error(f"Critical error: {e}", exc_info=True)
        print(f"\n💥 Critical error: {e}")
        input("\nPress Enter to exit...")
        sys.exit(1)

if __name__ == "__main__":
    main()