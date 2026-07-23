"""Tests for the shared find_open_workbook helper (utils.excel_com).

Uses tiny fakes to stand in for the Excel COM Application/Workbook objects so
the "is this file already open?" detection can be tested without Excel.
"""
from utils.excel_com import find_open_workbook


class FakeWorkbook:
    def __init__(self, fullname):
        self.FullName = fullname


class FakeExcel:
    def __init__(self, workbooks):
        self.Workbooks = workbooks


def test_returns_matching_open_workbook(tmp_path):
    target = tmp_path / "ITC_Report.xlsx"
    match = FakeWorkbook(str(target))
    excel = FakeExcel([FakeWorkbook(str(tmp_path / "other.xlsx")), match])
    assert find_open_workbook(excel, target.absolute()) is match


def test_returns_none_when_not_open(tmp_path):
    excel = FakeExcel([FakeWorkbook(str(tmp_path / "a.xlsx"))])
    assert find_open_workbook(excel, (tmp_path / "b.xlsx").absolute()) is None


def test_returns_none_for_no_open_workbooks(tmp_path):
    assert find_open_workbook(FakeExcel([]), (tmp_path / "a.xlsx").absolute()) is None
