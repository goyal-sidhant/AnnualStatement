"""
GST File Organizer & Report Generator v4.0
Main entry point - PyQt5 unified application.
"""

import sys
import io
import logging
from pathlib import Path

# Set UTF-8 encoding for stdout
if hasattr(sys.stdout, 'buffer'):
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
    """Setup Python path"""
    project_root = Path(__file__).parent
    sys.path.insert(0, str(project_root))


def main():
    """Main entry point"""
    try:
        setup_environment()

        from PyQt5.QtWidgets import QApplication
        from gui.pyqt_main_window import GSTOrganizerWindow

        app = QApplication(sys.argv)
        app.setApplicationName("GST File Organizer")
        app.setApplicationVersion("4.0")

        window = GSTOrganizerWindow()
        window.show()

        sys.exit(app.exec_())

    except ImportError as e:
        logging.error(f"Import Error: {e}")
        print(f"\n Missing module: {e}")
        print("\nPlease install required packages:")
        print("  pip install PyQt5 openpyxl")
        sys.exit(1)

    except Exception as e:
        logging.error(f"Critical error: {e}", exc_info=True)
        print(f"\n Critical error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
