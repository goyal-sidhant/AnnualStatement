"""Shared pytest fixtures + import path setup for the AnnualStatement test suite.

Adds the project root to sys.path so tests can `from utils.helpers import ...`
and `from core.file_parser import ...` without installing the package.
"""
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import pytest

# Real Excel file signatures (first bytes of the file on disk).
PK_SIG = b"PK\x03\x04"          # .xlsx / .xlsm / .xltx (ZIP-based)
OLE_SIG = b"\xd0\xcf\x11\xe0"   # .xls (OLE-based)


def _write_fake_excel(path, header=PK_SIG, size=2048):
    """Write a file with a valid Excel signature and enough padding to pass the
    >=1024-byte size check in validate_excel_file(). Not a real workbook - just
    enough for the signature/size validation the app performs."""
    path.write_bytes(header + b"\x00" * max(0, size - len(header)))
    return path


@pytest.fixture
def make_excel():
    """Fixture returning a helper to create fake-but-valid Excel files on disk."""
    return _write_fake_excel
