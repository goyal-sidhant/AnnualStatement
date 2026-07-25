"""Tests for recognising refreshed report files.

The "Refreshed File Suffix" is user-configurable, but the code that had to find
those files again hardcoded "_Refreshed_". With a customised suffix, reports
read as "Never refreshed" straight after a successful run, and skip-refresh
re-did work it should have reused.
"""
import pytest

from power_query_extractor.core.refresh_naming import (
    DEFAULT_SUFFIX_PATTERN, LEGACY_MARKER,
    refresh_markers, is_refreshed_name, is_refreshed_copy_of,
)


class TestRefreshMarkers:
    def test_default_pattern(self):
        assert refresh_markers(DEFAULT_SUFFIX_PATTERN) == ['_refreshed_']

    def test_custom_pattern(self):
        assert refresh_markers("_Updated_{timestamp}") == ['_updated_']

    def test_literal_on_both_sides_of_the_timestamp(self):
        assert refresh_markers("_v{timestamp}_final") == ['_v', '_final']

    def test_pattern_with_no_literal_falls_back_to_legacy(self):
        """'{timestamp}' alone cannot distinguish a refreshed file from the
        original, so matching everything would be wrong."""
        assert refresh_markers("{timestamp}") == [LEGACY_MARKER]

    @pytest.mark.parametrize("pattern", [None, ""])
    def test_missing_pattern_falls_back_to_legacy(self, pattern):
        assert refresh_markers(pattern) == [LEGACY_MARKER]


class TestIsRefreshedName:
    def test_default_suffix(self):
        assert is_refreshed_name("ITC_Report_ABC_Refreshed_250726_1130.xlsx",
                                 DEFAULT_SUFFIX_PATTERN)

    def test_custom_suffix_is_recognised(self):
        """The reported bug: a customised suffix was never matched."""
        assert is_refreshed_name("ITC_Report_ABC_Updated_250726_1130.xlsx",
                                 "_Updated_{timestamp}")

    def test_legacy_files_still_match_after_changing_the_setting(self):
        """Changing the setting must not make the existing history read as
        'Never refreshed'."""
        assert is_refreshed_name("ITC_Report_ABC_Refreshed_250726_1130.xlsx",
                                 "_Updated_{timestamp}")

    def test_original_report_is_not_a_refreshed_file(self):
        assert not is_refreshed_name("ITC_Report_ABC.xlsx", DEFAULT_SUFFIX_PATTERN)
        assert not is_refreshed_name("ITC_Report_ABC.xlsx", "_Updated_{timestamp}")

    def test_case_insensitive(self):
        assert is_refreshed_name("ITC_REPORT_ABC_UPDATED_250726.XLSX",
                                 "_Updated_{timestamp}")

    def test_wrong_extension_rejected(self):
        assert not is_refreshed_name("ITC_Report_ABC_Updated_250726.txt",
                                     "_Updated_{timestamp}")

    def test_both_literals_required_for_a_split_pattern(self):
        assert is_refreshed_name("ITC_Report_ABC_v250726_final.xlsx",
                                 "_v{timestamp}_final")
        # missing the trailing literal, and no legacy marker either
        assert not is_refreshed_name("ITC_Report_ABC_v250726.xlsx",
                                     "_v{timestamp}_final")


class TestIsRefreshedCopyOf:
    STEM = "ITC_Report_ABC_DL"

    def test_matches_a_refreshed_copy_of_that_report(self):
        assert is_refreshed_copy_of(f"{self.STEM}_Updated_250726.xlsx",
                                    self.STEM, "_Updated_{timestamp}")

    def test_rejects_a_different_report(self):
        assert not is_refreshed_copy_of("Sales_Report_ABC_DL_Updated_250726.xlsx",
                                        self.STEM, "_Updated_{timestamp}")

    def test_rejects_the_original(self):
        assert not is_refreshed_copy_of(f"{self.STEM}.xlsx", self.STEM,
                                        "_Updated_{timestamp}")


class TestStatusDetectionUsesTheConfiguredSuffix:
    """End-to-end through the function the client list actually calls."""

    def _tree(self, tmp_path, filename):
        version = tmp_path / "Version-250726 1000"
        version.mkdir()
        (version / "ITC_Report_ABC_DL.xlsx").write_bytes(b'PK\x03\x04' + bytes(2048))
        (version / filename).write_bytes(b'PK\x03\x04' + bytes(2048))
        return version

    def test_custom_suffix_reports_a_time_not_never(self, tmp_path):
        from power_query_extractor.gui.extractor_window import _latest_refresh_times
        version = self._tree(tmp_path, "ITC_Report_ABC_DL_Updated_250726_1130.xlsx")

        itc, sales = _latest_refresh_times(version, "_Updated_{timestamp}")
        assert itc is not None, "a custom suffix must be recognised"
        assert sales is None

    def test_custom_suffix_was_previously_missed(self, tmp_path):
        """Same tree, but detection told to expect the DEFAULT suffix: this is
        what used to happen, and it finds nothing."""
        from power_query_extractor.gui.extractor_window import _latest_refresh_times
        version = self._tree(tmp_path, "ITC_Report_ABC_DL_Updated_250726_1130.xlsx")

        itc, _ = _latest_refresh_times(version, DEFAULT_SUFFIX_PATTERN)
        assert itc is None

    def test_original_alone_reports_never(self, tmp_path):
        from power_query_extractor.gui.extractor_window import _latest_refresh_times
        version = tmp_path / "Version-250726 1000"
        version.mkdir()
        (version / "ITC_Report_ABC_DL.xlsx").write_bytes(b'PK\x03\x04' + bytes(2048))

        assert _latest_refresh_times(version, "_Updated_{timestamp}") == (None, None)
