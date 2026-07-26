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

def _pause(message="\nPress Enter to exit..."):
    """Wait for the user, but never fail if there is no console input.

    Without this, a blocked start prints its real reason and then a spurious
    "Critical error: EOF when reading a line" on top of it.
    """
    try:
        input(message)
    except (EOFError, KeyboardInterrupt, OSError):
        pass


def main():
    """Main entry point"""
    try:
        setup_environment()

        print("=" * 60)
        print("🚀 GST File Organizer v3.0 - Production Ready")
        print("=" * 60)

        # Check dependencies BEFORE importing the GUI, so a missing package is
        # reported plainly instead of surfacing as an obscure ImportError.
        from utils.requirements_check import (
            check_requirements, format_report, missing_essential, missing_optional,
        )
        results = check_requirements()
        print(format_report(results))
        print("=" * 60)

        blocking = missing_essential(results)
        if blocking:
            for result in blocking:
                logging.error("Required package missing: %s",
                              result.requirement.package)
            _pause()
            sys.exit(1)

        for result in missing_optional(results):
            logging.warning("%s not installed - %s",
                            result.requirement.package, result.requirement.purpose)

        # Import after the check
        from gui.main_window import GSTOrganizerApp

        app = GSTOrganizerApp()
        app.run()

    except ImportError as e:
        logging.error(f"Import Error: {e}", exc_info=True)
        print(f"\n❌ Could not start: {e}")
        print("\nThis usually means a required package is missing for THIS Python,")
        print("or a source file has been moved.")
        try:
            from utils.requirements_check import interpreter_hint
            print(f"\n   Running: {interpreter_hint()}")
        except Exception:
            pass
        print("\n   Try:  pip install -r requirements.txt")
        print("   Note 'py' and 'python' can be different interpreters.")
        _pause()
        sys.exit(1)


    except Exception as e:
        logging.error(f"Critical error: {e}", exc_info=True)
        print(f"\n💥 Critical error: {e}")
        _pause()
        sys.exit(1)

if __name__ == "__main__":
    main()