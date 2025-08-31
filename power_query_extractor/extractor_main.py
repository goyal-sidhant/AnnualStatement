"""
Main entry point for Power Query Extractor
"""

import sys
import logging
import json
from pathlib import Path
from .gui.extractor_window import PowerQueryExtractorApp

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('pq_extractor.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)


class PowerQueryExtractor:
    """Main class to launch the PQ Extractor"""
    
    def __init__(self, target_folder=None):
        self.target_folder = target_folder
        
        # Try to load from cache if no folder provided
        if not self.target_folder:
            self.target_folder = self.load_cached_folder()
            
        logger.info(f"Initializing Power Query Extractor with folder: {self.target_folder}")
    
    def load_cached_folder(self):
        """Load target folder from main app's cache"""
        cache_files = [
            Path("gst_organizer_cache.json"),
            Path(".cache/app_state.json"),
            Path("config/last_session.json")
        ]
        
        for cache_file in cache_files:
            try:
                if cache_file.exists():
                    logger.info(f"Found cache file: {cache_file}")
                    with open(cache_file, 'r') as f:
                        cache_data = json.load(f)
                        
                        # Try different possible keys
                        for key in ['target_folder', 'target_path', 'output_folder', 'last_target']:
                            if key in cache_data and cache_data[key]:
                                logger.info(f"Found target folder in cache: {cache_data[key]}")
                                return cache_data[key]
            except Exception as e:
                logger.warning(f"Could not load cache from {cache_file}: {e}")
        
        logger.info("No cached folder found")
        return ''
    
    def run(self):
        """Launch the extractor GUI"""
        try:
            app = PowerQueryExtractorApp(self.target_folder)
            app.mainloop()
        except Exception as e:
            logger.error(f"Error launching extractor: {e}", exc_info=True)
            raise


def main():
    """Standalone entry point"""
    extractor = PowerQueryExtractor()
    extractor.run()


if __name__ == "__main__":
    main()