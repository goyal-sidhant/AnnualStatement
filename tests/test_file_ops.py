"""Tests for the filesystem helpers in utils.helpers.

Covers validate_excel_file / find_excel_files / get_file_info against real
files on disk (pytest tmp_path). These are the functions optimised for network
paths, so pinning their behaviour guards against a refactor changing which
files are considered valid or discovered.
"""
import os

from utils.helpers import (
    validate_excel_file, find_excel_files, get_file_info, scan_excel_files,
)

OLE_SIG = b"\xd0\xcf\x11\xe0"   # .xls (OLE-based) signature


class TestValidateExcelFile:
    def test_valid_xlsx(self, tmp_path, make_excel):
        assert validate_excel_file(make_excel(tmp_path / "a.xlsx")) is True

    def test_valid_xls_ole(self, tmp_path, make_excel):
        assert validate_excel_file(make_excel(tmp_path / "a.xls", header=OLE_SIG)) is True

    def test_valid_xltx_template(self, tmp_path, make_excel):
        # validate_excel_file accepts template extensions too.
        assert validate_excel_file(make_excel(tmp_path / "t.xltx")) is True

    def test_too_small_is_rejected(self, tmp_path, make_excel):
        assert validate_excel_file(make_excel(tmp_path / "tiny.xlsx", size=500)) is False

    def test_wrong_signature_is_rejected(self, tmp_path, make_excel):
        assert validate_excel_file(make_excel(tmp_path / "fake.xlsx", header=b"NOPE")) is False

    def test_wrong_extension_is_rejected(self, tmp_path, make_excel):
        assert validate_excel_file(make_excel(tmp_path / "notes.txt")) is False

    def test_missing_file_is_rejected(self, tmp_path):
        assert validate_excel_file(tmp_path / "nope.xlsx") is False

    def test_directory_is_rejected(self, tmp_path):
        d = tmp_path / "adir.xlsx"
        d.mkdir()
        assert validate_excel_file(d) is False


class TestFindExcelFiles:
    def test_finds_only_valid_globbable_excels_sorted(self, tmp_path, make_excel):
        make_excel(tmp_path / "GSTR3B-ABC-DL-Jan.xlsx")
        make_excel(tmp_path / "Sales-ABC-DL-Apr-Jun.xls", header=OLE_SIG)
        make_excel(tmp_path / "ImsReco-ABC-DL-01012024.xlsm")
        make_excel(tmp_path / "UPPER.XLSX")               # case-insensitive match
        make_excel(tmp_path / "tiny.xlsx", size=500)      # invalid: too small
        make_excel(tmp_path / "fake.xlsx", header=b"NO")  # invalid: signature
        make_excel(tmp_path / "template.xltx")            # not globbed (only xlsx/xls/xlsm)
        (tmp_path / "notes.txt").write_text("hi")         # not excel
        (tmp_path / "looks.xlsx").mkdir()                 # a directory

        found = {p.name for p in find_excel_files(tmp_path)}
        assert found == {
            "GSTR3B-ABC-DL-Jan.xlsx",
            "Sales-ABC-DL-Apr-Jun.xls",
            "ImsReco-ABC-DL-01012024.xlsm",
            "UPPER.XLSX",
        }

    def test_returns_sorted(self, tmp_path, make_excel):
        for name in ["c.xlsx", "a.xlsx", "b.xlsx"]:
            make_excel(tmp_path / name)
        names = [p.name for p in find_excel_files(tmp_path)]
        assert names == sorted(names)

    def test_empty_folder_returns_empty_list(self, tmp_path):
        assert find_excel_files(tmp_path) == []


class TestScanExcelFiles:
    """scan_excel_files returns (path, stat) so callers can avoid re-stat'ing.

    Its results must match find_excel_files exactly - the only difference is
    that the cached stat comes along for the ride.
    """

    def test_matches_find_excel_files(self, tmp_path, make_excel):
        make_excel(tmp_path / "GSTR3B-ABC-DL-Jan.xlsx")
        make_excel(tmp_path / "Sales-ABC-DL-Apr-Jun.xls", header=OLE_SIG)
        make_excel(tmp_path / "tiny.xlsx", size=500)       # invalid
        make_excel(tmp_path / "fake.xlsx", header=b"NO")   # invalid
        (tmp_path / "looks.xlsx").mkdir()                  # directory

        pairs = scan_excel_files(tmp_path)
        assert [p for p, _ in pairs] == find_excel_files(tmp_path)

    def test_returns_usable_stat(self, tmp_path, make_excel):
        make_excel(tmp_path / "a.xlsx", size=4096)
        (path, st) = scan_excel_files(tmp_path)[0]
        assert isinstance(st, os.stat_result)
        assert st.st_size == 4096 == path.stat().st_size

    def test_rejects_directory_via_cached_stat(self, tmp_path):
        (tmp_path / "adir.xlsx").mkdir()
        assert scan_excel_files(tmp_path) == []

    def test_empty_folder(self, tmp_path):
        assert scan_excel_files(tmp_path) == []

    def test_missing_folder_returns_empty(self, tmp_path):
        assert scan_excel_files(tmp_path / "nope") == []


class TestCachedStatEquivalence:
    """Passing a pre-obtained stat must not change the outcome."""

    def test_validate_same_with_and_without_stat(self, tmp_path, make_excel):
        cases = [
            make_excel(tmp_path / "good.xlsx"),
            make_excel(tmp_path / "small.xlsx", size=500),
            make_excel(tmp_path / "bad.xlsx", header=b"NOPE"),
            make_excel(tmp_path / "notes.txt"),
        ]
        for p in cases:
            assert validate_excel_file(p) == validate_excel_file(p, p.stat()), p.name

    def test_get_file_info_same_with_and_without_stat(self, tmp_path, make_excel):
        p = make_excel(tmp_path / "a.xlsx", size=3072)
        assert get_file_info(p) == get_file_info(p, p.stat())


class TestGetFileInfo:
    def test_returns_metadata_for_real_file(self, tmp_path, make_excel):
        p = make_excel(tmp_path / "a.xlsx", size=2048)
        info = get_file_info(p)
        assert "error" not in info
        assert info["name"] == "a.xlsx"
        assert info["size"] == 2048
        assert info["is_file"] is True
        assert info["is_dir"] is False
        assert info["extension"] == ".xlsx"

    def test_reports_directory(self, tmp_path):
        info = get_file_info(tmp_path)
        assert info["is_dir"] is True
        assert info["is_file"] is False

    def test_missing_file_returns_error(self, tmp_path):
        assert get_file_info(tmp_path / "nope.xlsx") == {"error": "File not found"}
