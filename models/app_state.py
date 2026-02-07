"""
Framework-agnostic application state dataclasses.
These hold all state that was previously scattered across the Tkinter GSTOrganizerApp god object.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Any, Optional


@dataclass
class AppSettings:
    """User-configurable application settings"""
    source_folder: str = ''
    itc_template: str = ''
    sales_template: str = ''
    target_folder: str = ''
    processing_mode: str = 'fresh'
    include_client_name_in_folders: bool = False
    client_name_max_length: int = 35
    dark_mode: bool = False
    pq_wait_time: int = 10
    pq_suffix_pattern: str = '_Refreshed_{timestamp}'
    pq_skip_refresh: bool = False

    def to_dict(self) -> Dict[str, Any]:
        """Serialize settings to dict for caching"""
        return {
            'source_folder': self.source_folder,
            'itc_template': self.itc_template,
            'sales_template': self.sales_template,
            'target_folder': self.target_folder,
            'processing_mode': self.processing_mode,
            'include_client_name_in_folders': self.include_client_name_in_folders,
            'client_name_max_length': self.client_name_max_length,
            'dark_mode': self.dark_mode,
            'pq_wait_time': self.pq_wait_time,
            'pq_suffix_pattern': self.pq_suffix_pattern,
            'pq_skip_refresh': self.pq_skip_refresh,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'AppSettings':
        """Deserialize settings from dict"""
        known_fields = {f.name for f in cls.__dataclass_fields__.values()}
        filtered = {k: v for k, v in data.items() if k in known_fields}
        return cls(**filtered)


@dataclass
class ScanResult:
    """Results from scanning a source folder"""
    scanned_files: Dict[str, Any] = field(default_factory=dict)
    client_data: Dict[str, Any] = field(default_factory=dict)
    variations: List[Dict] = field(default_factory=list)

    @property
    def total_files(self) -> int:
        return len(self.scanned_files)

    @property
    def total_clients(self) -> int:
        return len(self.client_data)

    @property
    def complete_clients(self) -> int:
        return sum(1 for c in self.client_data.values()
                   if c.get('status') == 'Complete')

    def clear(self):
        """Reset scan results"""
        self.scanned_files.clear()
        self.client_data.clear()
        self.variations.clear()


@dataclass
class ProcessingState:
    """Mutable state during processing"""
    is_processing: bool = False
    stop_requested: bool = False
    selected_clients: Dict[str, bool] = field(default_factory=dict)
    client_folder_settings: Dict[str, bool] = field(default_factory=dict)

    def reset(self):
        """Reset processing state"""
        self.is_processing = False
        self.stop_requested = False
