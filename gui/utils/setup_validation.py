"""Validation logic for the Setup tab.

Deliberately free of Tkinter so it can be unit-tested: the tab renders whatever
this returns. Path existence is injected (`exists`) so the GUI can supply a
cached/debounced checker - on a network share each existence check costs a real
round-trip (~45ms), which is too slow to run on every keystroke.
"""
import os
from pathlib import Path
from typing import Callable, Dict, List, NamedTuple, Optional

# Item states
OK = 'ok'            # set and present on disk
MISSING = 'missing'  # not chosen yet
INVALID = 'invalid'  # chosen but not found / wrong kind

# Which stage an item blocks if it is not OK
FOR_SCAN = 'scan'
FOR_PROCESSING = 'processing'

TEMPLATE_SUFFIXES = {'.xlsx', '.xltx', '.xltm', '.xlsm'}


class Check(NamedTuple):
    key: str
    label: str
    state: str
    message: str
    blocks: str

    @property
    def is_ok(self) -> bool:
        return self.state == OK


def _default_exists(path: str) -> bool:
    # Guard the blank case explicitly: Path('') normalises to Path('.'), the
    # current directory, which exists - so a blank value would look valid.
    if not path or not path.strip():
        return False
    try:
        return Path(path).exists()
    except (OSError, ValueError):
        return False


def _check_folder(key: str, label: str, value: str, blocks: str,
                  exists: Callable[[str], bool]) -> Check:
    if not value or not value.strip():
        return Check(key, label, MISSING, 'Not selected yet', blocks)
    if not exists(value):
        return Check(key, label, INVALID, 'Folder not found', blocks)
    return Check(key, label, OK, _short(value), blocks)


def _check_template(key: str, label: str, value: str, blocks: str,
                    exists: Callable[[str], bool]) -> Check:
    if not value or not value.strip():
        return Check(key, label, MISSING, 'Not selected yet', blocks)
    if not exists(value):
        return Check(key, label, INVALID, 'File not found', blocks)
    if Path(value).suffix.lower() not in TEMPLATE_SUFFIXES:
        return Check(key, label, INVALID, 'Not an Excel template', blocks)
    return Check(key, label, OK, Path(value).name, blocks)


def _short(value: str, width: int = 46) -> str:
    """Trim the middle of a long path so the tail (the useful part) stays visible."""
    if len(value) <= width:
        return value
    keep = width - 4
    return f"...{value[-keep:]}"


def evaluate_setup(source_folder: str, itc_template: str, sales_template: str,
                   target_folder: str,
                   exists: Optional[Callable[[str], bool]] = None) -> List[Check]:
    """Return one Check per required input, in display order."""
    ex = exists or _default_exists
    return [
        _check_folder('source', 'Source folder', source_folder, FOR_SCAN, ex),
        _check_template('itc', 'ITC template', itc_template, FOR_PROCESSING, ex),
        _check_template('sales', 'Sales template', sales_template, FOR_PROCESSING, ex),
        _check_folder('target', 'Target folder', target_folder, FOR_PROCESSING, ex),
    ]


def can_scan(checks: List[Check]) -> bool:
    """Scanning only needs a valid source folder (mirrors
    FileHandler.validate_scan_inputs)."""
    return all(c.is_ok for c in checks if c.blocks == FOR_SCAN)


def can_process(checks: List[Check]) -> bool:
    """Processing needs everything (mirrors validate_processing_inputs, plus the
    source folder that produced the scan)."""
    return all(c.is_ok for c in checks)


def blocking_labels(checks: List[Check], stage: str) -> List[str]:
    """Labels of the items still blocking the given stage."""
    if stage == FOR_SCAN:
        relevant = [c for c in checks if c.blocks == FOR_SCAN]
    else:
        relevant = list(checks)
    return [c.label for c in relevant if not c.is_ok]


def summary_line(checks: List[Check]) -> str:
    """One-line readiness summary for the header."""
    done = sum(1 for c in checks if c.is_ok)
    total = len(checks)
    if done == total:
        return f"Ready - all {total} items set"
    return f"{done} of {total} items set"
