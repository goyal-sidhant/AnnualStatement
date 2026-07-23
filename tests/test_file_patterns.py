"""Tests for filename classification - FileParser.parse_filename + FILE_PATTERNS.

This is the most important logic in the app: it decides which client, state and
report type each file belongs to. A silent misclassification misfiles a client's
tax data, so every file type and its known edge cases (case-insensitivity, the
.xls extension, and the (N) duplicate markers) are pinned here.
"""
import pytest

from core.file_parser import FileParser


@pytest.fixture
def parser():
    return FileParser()


# (filename, expected_type, expected_client, expected_state)
HAPPY_CASES = [
    ("GSTR-2B-Reco-ABC-Delhi-Apr24.xlsx", "GSTR-2B Reco", "ABC", "Delhi"),
    ("ImsReco-ABC-Delhi-01012024.xlsx", "IMS Reco", "ABC", "Delhi"),
    ("GSTR3B-ABC-Delhi-Jan.xlsx", "GSTR-3B Export", "ABC", "Delhi"),
    ("Sales-ABC-Delhi-Apr-Jun.xlsx", "Sales", "ABC", "Delhi"),
    ("SalesReco-ABC-Delhi-Apr24.xlsx", "Sales Reco", "ABC", "Delhi"),
    ("AnnualReport-ABC-Delhi-2024.xlsx", "Annual Report", "ABC", "Delhi"),
    # .xls extension is accepted (pattern ends in \.xlsx?$)
    ("Sales-XYZ-Maharashtra-Apr-Jun.xls", "Sales", "XYZ", "Maharashtra"),
    # matching is case-insensitive (re.IGNORECASE on every pattern)
    ("gstr3b-abc-delhi-jan.xlsx", "GSTR-3B Export", "abc", "delhi"),
]


@pytest.mark.parametrize("filename,exp_type,exp_client,exp_state", HAPPY_CASES)
def test_parses_each_type(parser, filename, exp_type, exp_client, exp_state):
    r = parser.parse_filename(filename)
    assert r["parsed"] is True
    assert r["type"] == exp_type
    assert r["client"] == exp_client
    assert r["state"] == exp_state


# --- (N) duplicate markers -------------------------------------------------

# Trailing " (N)" already works for every type EXCEPT IMS Reco historically;
# IMS Reco was fixed to tolerate it (commit 5c8ce95). Pin all of them.
TRAILING_SUFFIX = [
    "GSTR-2B-Reco-ABC-Delhi-Apr24 (1).xlsx",
    "ImsReco-ABC-Delhi-01012024 (1).xlsx",
    "ImsReco-ABC-Delhi-01012024(2).xlsx",
    "ImsReco-ABC-Delhi-01012024 (10).xlsx",
    "Sales-ABC-Delhi-Apr-Jun (2).xlsx",
    "AnnualReport-ABC-Delhi-2024 (1).xlsx",
]


@pytest.mark.parametrize("filename", TRAILING_SUFFIX)
def test_trailing_duplicate_marker_still_parses(parser, filename):
    r = parser.parse_filename(filename)
    assert r["parsed"] is True
    assert r["client"] == "ABC"


def test_ims_suffix_keeps_clean_date(parser):
    """The (N) marker must NOT leak into the extracted 8-digit date."""
    r = parser.parse_filename("ImsReco-ABC-Delhi-01012024 (1).xlsx")
    assert r["type"] == "IMS Reco"
    assert r["metadata"].get("date") == "01012024"


# Leading "(N) " prefix is only handled for GSTR-3B (portal export convention).
def test_gstr3b_leading_prefix_parses(parser):
    r = parser.parse_filename("(1) GSTR3B-ABC-Delhi-Jan.xlsx")
    assert r["parsed"] is True
    assert r["type"] == "GSTR-3B Export"
    assert r["client"] == "ABC"


# --- negatives -------------------------------------------------------------

NON_MATCHING = [
    "RandomFile.xlsx",
    "Notes-2024.xlsx",
    "GSTR-2B-Reco-ABC.xlsx",            # too few dash groups
    "Sales-ABC-Delhi-Apr-Jun-Extra.xlsx",  # too many dash groups
    "report.pdf",                      # wrong extension
]


@pytest.mark.parametrize("filename", NON_MATCHING)
def test_unrecognized_names_do_not_parse(parser, filename):
    r = parser.parse_filename(filename)
    assert r["parsed"] is False
    assert r["type"] == "Unknown"
    assert r["client"] == ""
