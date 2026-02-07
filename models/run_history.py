"""
Run history tracking for accumulated data across processing runs.
Stores per-client processing history, PQ refresh status, and extraction results.
"""

import json
import logging
from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional

logger = logging.getLogger(__name__)


@dataclass
class ClientRunRecord:
    """Record of a single processing run for one client"""
    client_key: str = ''
    client_name: str = ''
    state: str = ''
    run_timestamp: str = ''
    processing_mode: str = ''
    files_organized: int = 0
    itc_report_created: bool = False
    sales_report_created: bool = False
    itc_report_path: str = ''
    sales_report_path: str = ''
    output_folder: str = ''
    pq_refreshed: bool = False
    pq_refresh_timestamp: str = ''
    pq_extraction_data: Dict = field(default_factory=dict)
    errors: List[str] = field(default_factory=list)


@dataclass
class RunHistory:
    """Accumulated run history across sessions"""
    runs: List[Dict] = field(default_factory=list)
    client_history: Dict[str, List[Dict]] = field(default_factory=dict)
    last_run_timestamp: str = ''
    total_runs: int = 0

    def add_run(self, run_summary: Dict, client_records: List[ClientRunRecord]):
        """Record a completed processing run"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        run_entry = {
            'timestamp': timestamp,
            'total_clients': run_summary.get('total_clients', 0),
            'successful_clients': run_summary.get('successful_clients', 0),
            'failed_clients': run_summary.get('failed_clients', 0),
            'total_files': run_summary.get('total_files', 0),
            'reports_generated': run_summary.get('reports_generated', 0),
            'processing_mode': run_summary.get('processing_mode', ''),
        }

        self.runs.append(run_entry)
        self.last_run_timestamp = timestamp
        self.total_runs += 1

        # Add per-client records
        for record in client_records:
            key = record.client_key
            if key not in self.client_history:
                self.client_history[key] = []
            self.client_history[key].append(asdict(record))

        # Keep only last 50 runs to prevent unbounded growth
        if len(self.runs) > 50:
            self.runs = self.runs[-50:]

    def add_extraction_result(self, client_key: str, extraction_data: Dict):
        """Record PQ extraction results for a client"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        if client_key not in self.client_history:
            self.client_history[client_key] = []

        # Find the most recent record for this client and update it
        if self.client_history[client_key]:
            latest = self.client_history[client_key][-1]
            latest['pq_refreshed'] = True
            latest['pq_refresh_timestamp'] = timestamp
            latest['pq_extraction_data'] = extraction_data
        else:
            # No prior record, create one
            record = ClientRunRecord(
                client_key=client_key,
                run_timestamp=timestamp,
                pq_refreshed=True,
                pq_refresh_timestamp=timestamp,
                pq_extraction_data=extraction_data
            )
            self.client_history[client_key].append(asdict(record))

    def get_last_run_for_client(self, client_key: str) -> Optional[Dict]:
        """Get the most recent run record for a client"""
        records = self.client_history.get(client_key, [])
        return records[-1] if records else None

    def get_last_run_timestamp_for_client(self, client_key: str) -> str:
        """Get formatted last run timestamp for display in client tree"""
        record = self.get_last_run_for_client(client_key)
        if record:
            return record.get('run_timestamp', '')
        return ''

    def to_dict(self) -> Dict[str, Any]:
        """Serialize for JSON storage"""
        return {
            'runs': self.runs,
            'client_history': self.client_history,
            'last_run_timestamp': self.last_run_timestamp,
            'total_runs': self.total_runs,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'RunHistory':
        """Deserialize from JSON"""
        history = cls()
        history.runs = data.get('runs', [])
        history.client_history = data.get('client_history', {})
        history.last_run_timestamp = data.get('last_run_timestamp', '')
        history.total_runs = data.get('total_runs', 0)
        return history

    def save(self, file_path: Path):
        """Save run history to JSON file"""
        try:
            file_path.parent.mkdir(parents=True, exist_ok=True)
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.to_dict(), f, indent=2, default=str)
            logger.debug(f"Run history saved to {file_path}")
        except Exception as e:
            logger.error(f"Failed to save run history: {e}")

    @classmethod
    def load(cls, file_path: Path) -> 'RunHistory':
        """Load run history from JSON file"""
        try:
            if file_path.exists():
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return cls.from_dict(data)
        except Exception as e:
            logger.warning(f"Failed to load run history: {e}")
        return cls()
