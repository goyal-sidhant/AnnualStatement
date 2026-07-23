# Tests

Regression tests for the GST File Organizer's data-integrity logic — the parts
whose silent failure would misfile a client's files.

## Running

From the repository root:

```bash
pytest
```

(or `python -m pytest`). Requires `pytest` — `pip install -r requirements.txt`.

## What's covered

| File | Covers |
|------|--------|
| `test_file_patterns.py` | `parse_filename` / `FILE_PATTERNS` — every report type, case-insensitivity, the `.xls` extension, and the `(N)` duplicate markers (trailing suffix + GSTR-3B leading prefix). |
| `test_helpers.py` | Pure helpers in `utils.helpers`: `sanitize_filename`, `get_state_code`, `create_client_state_key`, `validate_client_name`, `format_size`, `format_duration`, path helpers. |
| `test_file_ops.py` | Filesystem helpers: `validate_excel_file`, `find_excel_files`, `get_file_info` (against real temp files). |
| `test_scan.py` | End-to-end `FileParser.scan_folder`: grouping by client, completeness/missing-file reporting, variations, invalid-file skipping. |

## Nature of these tests

These are **characterization tests**: they pin the *current* behavior so future
refactors can't change it silently. A handful of assertions are marked with a
`NOTE` comment where the current behavior is a known quirk (e.g. the
case-sensitive `Private`→`Pvt` substring replace mangling `Privateer`) — those
document reality; they are not endorsements of the behavior.

No GUI, Excel COM, or network access is exercised, so the suite runs fast
(<1s) and anywhere.
