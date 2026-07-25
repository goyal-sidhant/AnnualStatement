"""Tests for the Setup tab's validation logic (no Tkinter involved)."""
import pytest

from gui.utils import setup_validation as sv


def make_exists(present):
    """Fake existence checker: only the listed paths 'exist'."""
    return lambda p: p in set(present)


ALL_GOOD = dict(
    source_folder=r"C:\gst\source",
    itc_template=r"C:\gst\itc.xltx",
    sales_template=r"C:\gst\sales.xlsx",
    target_folder=r"C:\gst\out",
)


def evaluate(**overrides):
    kwargs = {**ALL_GOOD, **overrides}
    present = [v for v in kwargs.values() if v]
    return sv.evaluate_setup(exists=make_exists(present), **kwargs)


def by_key(checks):
    return {c.key: c for c in checks}


class TestEvaluateSetup:
    def test_all_present_is_all_ok(self):
        checks = evaluate()
        assert [c.key for c in checks] == ['source', 'itc', 'sales', 'target']
        assert all(c.is_ok for c in checks)

    def test_empty_value_is_missing(self):
        checks = by_key(evaluate(source_folder=""))
        assert checks['source'].state == sv.MISSING
        assert checks['source'].message == 'Not selected yet'

    def test_whitespace_only_is_missing(self):
        assert by_key(evaluate(target_folder="   "))['target'].state == sv.MISSING

    def test_absent_folder_is_invalid(self):
        checks = sv.evaluate_setup(**ALL_GOOD, exists=make_exists([]))
        assert by_key(checks)['source'].state == sv.INVALID
        assert by_key(checks)['source'].message == 'Folder not found'

    def test_absent_template_is_invalid(self):
        present = [ALL_GOOD['source_folder'], ALL_GOOD['target_folder']]
        checks = by_key(sv.evaluate_setup(**ALL_GOOD, exists=make_exists(present)))
        assert checks['itc'].state == sv.INVALID
        assert checks['itc'].message == 'File not found'

    def test_template_with_wrong_extension_is_invalid(self):
        checks = by_key(evaluate(itc_template=r"C:\gst\notes.txt"))
        assert checks['itc'].state == sv.INVALID
        assert checks['itc'].message == 'Not an Excel template'

    @pytest.mark.parametrize("suffix", ['.xlsx', '.xltx', '.xltm', '.xlsm'])
    def test_accepted_template_suffixes(self, suffix):
        checks = by_key(evaluate(itc_template=rf"C:\gst\tmpl{suffix}"))
        assert checks['itc'].is_ok

    def test_ok_template_shows_filename_only(self):
        assert by_key(evaluate())['itc'].message == 'itc.xltx'

    def test_long_folder_path_is_truncated_keeping_the_tail(self):
        long_path = "\\\\server\\share\\" + "\\".join(f"level{i}" for i in range(20))
        checks = by_key(evaluate(source_folder=long_path))
        msg = checks['source'].message
        assert msg.startswith('...')
        assert msg.endswith('level19')
        assert len(msg) <= 46


class TestGating:
    def test_scan_needs_only_the_source_folder(self):
        """Templates and target gate processing, NOT scanning - mirrors
        FileHandler.validate_scan_inputs."""
        checks = evaluate(itc_template="", sales_template="", target_folder="")
        assert sv.can_scan(checks) is True
        assert sv.can_process(checks) is False

    def test_scan_blocked_without_source(self):
        checks = evaluate(source_folder="")
        assert sv.can_scan(checks) is False
        assert sv.blocking_labels(checks, sv.FOR_SCAN) == ['Source folder']

    def test_process_needs_everything(self):
        assert sv.can_process(evaluate()) is True

    def test_blocking_labels_for_processing_lists_all_pending(self):
        checks = evaluate(sales_template="", target_folder="")
        assert sv.blocking_labels(checks, sv.FOR_PROCESSING) == \
            ['Sales template', 'Target folder']


class TestSummary:
    def test_summary_when_ready(self):
        assert sv.summary_line(evaluate()) == 'Ready - all 4 items set'

    def test_summary_counts_completed(self):
        assert sv.summary_line(evaluate(itc_template="", target_folder="")) == \
            '2 of 4 items set'


def test_default_exists_checker_handles_bad_input():
    """The real checker must never raise on junk (embedded NUL, etc.)."""
    assert sv._default_exists("") is False
    assert sv._default_exists("C:\\does\\not\\exist\\anywhere") is False
    assert sv._default_exists("bad\x00path") is False
