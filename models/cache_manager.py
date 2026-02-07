"""
Enhanced cache manager for persistent application state.
Replaces the old gui/utils/cache_manager.py with richer data storage.
"""

import json
import logging
from pathlib import Path
from typing import Any, Dict, Optional

from .app_state import AppSettings
from .run_history import RunHistory

logger = logging.getLogger(__name__)

# Default cache directory
_DEFAULT_CACHE_DIR = Path.home() / '.gst_organizer'


class CacheManager:
    """
    Manages persistent application state across sessions.

    Stores:
    - User settings (folders, templates, UI preferences)
    - Run history (per-client processing records, extraction results)
    - Session state (selected clients, active tab)
    - Recent source folders list
    """

    def __init__(self, cache_dir: Optional[Path] = None):
        self.cache_dir = cache_dir or _DEFAULT_CACHE_DIR
        self.cache_dir.mkdir(parents=True, exist_ok=True)

        self._settings_path = self.cache_dir / 'settings.json'
        self._history_path = self.cache_dir / 'run_history.json'
        self._session_path = self.cache_dir / 'session_state.json'

    # ── Settings ──────────────────────────────────────────────

    def save_settings(self, settings: AppSettings):
        """Save application settings"""
        data = settings.to_dict()

        # Also maintain recent folders list
        session = self._load_json(self._session_path)
        recent = session.get('recent_source_folders', [])
        if settings.source_folder and settings.source_folder not in recent:
            recent.insert(0, settings.source_folder)
            recent = recent[:5]  # Keep last 5
            session['recent_source_folders'] = recent
            self._save_json(self._session_path, session)

        self._save_json(self._settings_path, data)
        logger.debug("Settings saved")

    def load_settings(self) -> AppSettings:
        """Load application settings"""
        data = self._load_json(self._settings_path)
        if data:
            return AppSettings.from_dict(data)
        return AppSettings()

    # ── Run History ───────────────────────────────────────────

    def save_run_history(self, history: RunHistory):
        """Save run history"""
        history.save(self._history_path)

    def load_run_history(self) -> RunHistory:
        """Load run history"""
        return RunHistory.load(self._history_path)

    # ── Session State ─────────────────────────────────────────

    def save_session_state(self, state: Dict[str, Any]):
        """Save transient session state (selected clients, active tab, etc.)"""
        existing = self._load_json(self._session_path)
        existing.update(state)
        self._save_json(self._session_path, existing)

    def load_session_state(self) -> Dict[str, Any]:
        """Load session state"""
        return self._load_json(self._session_path)

    def get_recent_source_folders(self):
        """Get list of recently used source folders"""
        session = self._load_json(self._session_path)
        return session.get('recent_source_folders', [])

    # ── Migration ─────────────────────────────────────────────

    def migrate_from_old_cache(self):
        """
        Import data from the old cache file (~/.gst_organizer_cache.json)
        if it exists and we haven't migrated yet.
        """
        old_cache = Path.home() / '.gst_organizer_cache.json'
        migrated_marker = self.cache_dir / '.migrated'

        if not old_cache.exists() or migrated_marker.exists():
            return

        try:
            with open(old_cache, 'r', encoding='utf-8') as f:
                old_data = json.load(f)

            # Map old keys to new settings
            settings = AppSettings(
                source_folder=old_data.get('source_folder', ''),
                itc_template=old_data.get('itc_template', ''),
                sales_template=old_data.get('sales_template', ''),
                target_folder=old_data.get('target_folder', ''),
                processing_mode=old_data.get('processing_mode', 'fresh'),
                include_client_name_in_folders=old_data.get('include_client_name', False),
                client_name_max_length=old_data.get('client_name_max_length', 35),
                dark_mode=old_data.get('dark_mode', False),
            )
            self.save_settings(settings)

            # Save recent folders if available
            recent = old_data.get('recent_source_folders', [])
            if recent:
                self.save_session_state({'recent_source_folders': recent})

            # Mark as migrated
            migrated_marker.touch()
            logger.info("Migrated settings from old cache file")

        except Exception as e:
            logger.warning(f"Failed to migrate old cache: {e}")

    # ── Internal helpers ──────────────────────────────────────

    def _save_json(self, path: Path, data: Dict):
        """Save dict to JSON file"""
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, default=str)
        except Exception as e:
            logger.error(f"Failed to save {path.name}: {e}")

    def _load_json(self, path: Path) -> Dict:
        """Load dict from JSON file"""
        try:
            if path.exists():
                with open(path, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            logger.warning(f"Failed to load {path.name}: {e}")
        return {}
