"""Integration tests for FileParser.scan_folder + completeness analysis.

Exercises the full scan path (discover -> validate -> parse -> group -> analyze)
against real files, so a change to any step that alters which clients/types are
detected or how completeness is reported is caught.
"""
import pytest

from core.file_parser import FileParser
from utils.constants import EXPECTED_FILE_TYPES


@pytest.fixture
def parser():
    return FileParser()


def test_groups_files_by_client_and_reports_missing(parser, tmp_path, make_excel):
    make_excel(tmp_path / "GSTR3B-ABC-Delhi-Jan.xlsx")
    make_excel(tmp_path / "Sales-ABC-Delhi-Apr-Jun.xlsx")

    scanned, client_data, variations = parser.scan_folder(tmp_path)

    assert len(scanned) == 2
    assert set(client_data.keys()) == {"ABC-DL"}

    abc = client_data["ABC-DL"]
    assert abc["client"] == "ABC"
    assert abc["state"] == "Delhi"
    assert abc["state_code"] == "DL"
    assert set(abc["files"].keys()) == {"GSTR-3B Export", "Sales"}
    assert abc["file_count"] == 2

    # Two of six expected types present -> four missing.
    missing = set(abc["missing_files"])
    assert missing == set(EXPECTED_FILE_TYPES) - {"GSTR-3B Export", "Sales"}
    assert len(missing) == 4
    assert abc["status"] == "Missing 4 files"


def test_unparseable_valid_excel_goes_to_variations(parser, tmp_path, make_excel):
    make_excel(tmp_path / "GSTR3B-ABC-Delhi-Jan.xlsx")   # valid + parseable
    make_excel(tmp_path / "MysteryFile.xlsx")            # valid excel, no pattern match

    scanned, client_data, variations = parser.scan_folder(tmp_path)

    assert set(client_data.keys()) == {"ABC-DL"}
    variation_names = {v["filename"] for v in variations}
    assert "MysteryFile.xlsx" in variation_names


def test_invalid_excel_is_skipped_entirely(parser, tmp_path, make_excel):
    make_excel(tmp_path / "GSTR3B-ABC-Delhi-Jan.xlsx")
    make_excel(tmp_path / "corrupt.xlsx", header=b"NOPE")  # bad signature

    scanned, client_data, variations = parser.scan_folder(tmp_path)

    assert set(scanned.keys()) == {"GSTR3B-ABC-Delhi-Jan.xlsx"}


def test_empty_folder_yields_no_clients(parser, tmp_path):
    scanned, client_data, variations = parser.scan_folder(tmp_path)
    assert scanned == {}
    assert dict(client_data) == {}


def test_complete_client_marked_complete(parser, tmp_path, make_excel):
    # One file of every expected type for the same client -> status Complete.
    make_excel(tmp_path / "GSTR-2B-Reco-ABC-Delhi-Q1.xlsx")
    make_excel(tmp_path / "ImsReco-ABC-Delhi-01012024.xlsx")
    make_excel(tmp_path / "GSTR3B-ABC-Delhi-Jan.xlsx")
    make_excel(tmp_path / "Sales-ABC-Delhi-Apr-Jun.xlsx")
    make_excel(tmp_path / "SalesReco-ABC-Delhi-Q1.xlsx")
    make_excel(tmp_path / "AnnualReport-ABC-Delhi-2024.xlsx")

    _, client_data, _ = parser.scan_folder(tmp_path)

    abc = client_data["ABC-DL"]
    assert abc["missing_files"] == []
    assert abc["status"] == "Complete"
