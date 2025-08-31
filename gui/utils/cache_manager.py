# gui/utils/cache_manager.py
"""Cache Manager for saving and loading application state"""

import json
import logging
from pathlib import Path

logger = logging.getLogger(__name__)


class CacheManager:
    """Manages application cache and settings"""
    
    def __init__(self, cache_file=None):
        self.cache_file = cache_file or (Path.home() / '.gst_organizer_cache.json')
        self.cache_data = {}
    
    def load_cache(self):
        """Load cached settings"""
        try:
            if self.cache_file.exists():
                with open(self.cache_file, 'r') as f:
                    self.cache_data = json.load(f)
                    return self.cache_data
        except Exception as e:
            logger.error(f"Error loading cache: {e}")
        return {}
    
    def save_cache(self, data):
        """Save current settings to cache"""
        try:
            # Preserve recent folders if they exist
            existing_cache = self.load_cache()
            recent_folders = existing_cache.get('recent_folders', [])
            
            # Add current folder if not empty
            current_folder = data.get('source_folder', '')
            if current_folder and current_folder not in recent_folders:
                recent_folders.insert(0, current_folder)
                recent_folders = recent_folders[:5]  # Keep only last 5
            
            # Update data with recent folders
            data['recent_folders'] = recent_folders
            
            with open(self.cache_file, 'w') as f:
                json.dump(data, f)
        except Exception as e:
            logger.error(f"Error saving cache: {e}")
    
    def get_cached_value(self, key, default=None):
        """Get a cached value"""
        return self.cache_data.get(key, default)
    
    def set_cached_value(self, key, value):
        """Set a cached value"""
        self.cache_data[key] = value