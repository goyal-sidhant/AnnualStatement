r"""Shared helpers for Excel COM workbook handling.

IMPORTANT: the two ``_safe_open_workbook`` methods (core.excel_handler and
power_query_extractor.core.report_processor) deliberately keep DIFFERENT
open-with-fallback strategies, and share only the piece below:

- ExcelHandler opens a freshly-copied LOCAL temp file (short path, no lock
  contention) -> a simple 8.3 short-path fallback is enough.
- ReportProcessor opens a NETWORK file that may still be locked from a
  just-finished Power Query refresh -> it needs extended ``\\?\`` paths plus
  close-by-name retries.

Do NOT merge those two strategies without validating BOTH flows against real
Microsoft Excel - they solve different problems.
"""
from pathlib import Path


def find_open_workbook(excel, abs_path):
    """Return an already-open COM workbook whose file is ``abs_path``, else None.

    ``abs_path`` must be an absolute ``pathlib.Path``. This mirrors, exactly, the
    lookup loop both ``_safe_open_workbook`` methods previously inlined, so a fix
    to the "already open" detection now applies to both call sites.
    """
    for wb in excel.Workbooks:
        if Path(wb.FullName).absolute() == abs_path:
            return wb
    return None
