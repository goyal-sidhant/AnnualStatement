"""Tests for the pure string/number helpers in utils.helpers.

These are characterization tests: they pin the CURRENT behaviour so future
refactors can't silently change how filenames, client keys, state codes and
display strings are produced. A few known quirks are pinned deliberately and
flagged with NOTE - they document real behaviour worth being aware of.
"""
import pytest

from utils.helpers import (
    sanitize_filename,
    get_state_code,
    create_client_state_key,
    validate_client_name,
    format_size,
    format_duration,
    clean_windows_path,
    extract_filename_without_extension,
)


# --- sanitize_filename -----------------------------------------------------

class TestSanitizeFilename:
    def test_abbreviates_private_limited_and_underscores_spaces(self):
        assert sanitize_filename(
            "GSTR-2B-Reco-ABC Private Limited-Delhi-Apr24.xlsx"
        ) == "GSTR-2B-Reco-ABC_Pvt_Ltd-Delhi-Apr24.xlsx"

    def test_removes_invalid_characters(self):
        assert sanitize_filename("Report<>:*.xlsx") == "Report.xlsx"

    def test_collapses_repeated_whitespace(self):
        assert sanitize_filename("  multiple   spaces  ") == "multiple_spaces"

    def test_empty_becomes_unnamed(self):
        assert sanitize_filename("") == "unnamed"

    def test_parentheses_are_preserved(self):
        # Duplicate markers like (1) survive sanitisation.
        assert sanitize_filename("(1) Sales-ABC-DL.xlsx") == "(1)_Sales-ABC-DL.xlsx"

    def test_abbreviation_uses_word_boundaries(self):
        # Words that merely CONTAIN 'Private'/'Limited' must NOT be abbreviated.
        assert sanitize_filename("Privateer Unlimited") == "Privateer_Unlimited"

    def test_abbreviation_is_case_insensitive(self):
        assert sanitize_filename("ABC PRIVATE LIMITED") == "ABC_Pvt_Ltd"

    def test_long_name_is_truncated_within_limit(self):
        result = sanitize_filename("x" * 250 + ".xlsx")
        assert result == "x" * 185 + ".xlsx"
        assert len(result) <= 200


# --- get_state_code --------------------------------------------------------

class TestGetStateCode:
    @pytest.mark.parametrize("name,code", [
        ("Delhi", "DL"),
        ("delhi", "DL"),
        ("  Maharashtra ", "MH"),
        ("Tamil Nadu", "TN"),
        ("West Bengal", "WB"),
    ])
    def test_known_states(self, name, code):
        assert get_state_code(name) == code

    def test_unknown_multiword_uses_initials(self):
        # NOTE: 'Unknown Place' -> 'UP', which collides with Uttar Pradesh.
        assert get_state_code("Unknown Place") == "UP"

    def test_unknown_singleword_uses_first_three(self):
        assert get_state_code("Xyz") == "XYZ"

    def test_empty_returns_empty(self):
        assert get_state_code("") == ""


# --- create_client_state_key ----------------------------------------------

class TestCreateClientStateKey:
    def test_abbreviates_with_word_boundaries_and_spaces(self):
        # Uses \bPrivate\b (case-insensitive), space-joined - distinct from
        # sanitize_filename's underscore/substring behaviour.
        assert create_client_state_key("ABC Private Limited", "Delhi") == "ABC Pvt Ltd-DL"

    def test_short_name_unchanged(self):
        assert create_client_state_key("XYZ", "Tamil Nadu") == "XYZ-TN"

    def test_long_key_capped_at_max_length(self):
        key = create_client_state_key(
            "A Very Long Client Name That Exceeds The Limit For Sure", "Maharashtra"
        )
        assert len(key) == 35
        assert key.endswith("-MH")


# --- validate_client_name --------------------------------------------------

class TestValidateClientName:
    @pytest.mark.parametrize("name,valid,msg", [
        ("", False, "Client name cannot be empty"),
        ("A", False, "Client name too short"),
        ("ABC", True, ""),
        ("Valid Name", True, ""),
        ("x" * 101, False, "Client name too long"),
        ("AB<C", False, "Client name contains invalid characters"),
    ])
    def test_validation(self, name, valid, msg):
        assert validate_client_name(name) == (valid, msg)


# --- format helpers --------------------------------------------------------

class TestFormatSize:
    @pytest.mark.parametrize("size,text", [
        (0, "0.0 B"),
        (1023, "1023.0 B"),
        (1024, "1.0 KB"),
        (1536, "1.5 KB"),
        (1048576, "1.0 MB"),
        (1073741824, "1.0 GB"),
        (-5, "Unknown"),
    ])
    def test_format_size(self, size, text):
        assert format_size(size) == text


class TestFormatDuration:
    @pytest.mark.parametrize("seconds,text", [
        (0, "0s"),
        (30, "30s"),
        (59, "59s"),
        (90, "1m 30s"),
        (3599, "59m 59s"),
        (3661, "1h 1m"),   # NOTE: hours format drops the seconds component
    ])
    def test_format_duration(self, seconds, text):
        assert format_duration(seconds) == text


# --- path helpers ----------------------------------------------------------

def test_clean_windows_path_uses_backslashes():
    assert clean_windows_path("a/b/c") == r"a\b\c"


@pytest.mark.parametrize("filename,stem", [
    ("report.xlsx", "report"),
    ("X.Y.xlsx", "X.Y"),          # only the final extension is stripped
    ("noext", "noext"),
])
def test_extract_filename_without_extension(filename, stem):
    assert extract_filename_without_extension(filename) == stem
