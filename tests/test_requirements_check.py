"""Tests for the startup dependency check.

Written after the app died with "No module named 'pythoncom'" behind a message
that blamed missing source files. The check has to name the real package, say
what breaks without it, and name the interpreter - because `py` and `python`
can be different Pythons with different packages installed.
"""
from utils.requirements_check import (
    REQUIREMENTS, Requirement, check_requirements, format_report,
    interpreter_hint, missing_essential, missing_optional,
)

PRESENT = Requirement('sys', 'sys-package', 'Always available', True)
ABSENT_ESSENTIAL = Requirement('no_such_module_xyz', 'ghost-pkg',
                               'Something vital', True)
ABSENT_OPTIONAL = Requirement('no_such_module_abc', 'optional-pkg',
                              'Something nice to have', False)


class TestDetection:
    def test_detects_an_installed_module(self):
        assert check_requirements([PRESENT])[0].installed is True

    def test_detects_a_missing_module(self):
        assert check_requirements([ABSENT_ESSENTIAL])[0].installed is False

    def test_does_not_import_the_module_to_check_it(self):
        """find_spec avoids executing the package just to test for it."""
        import sys
        assert 'no_such_module_xyz' not in sys.modules
        check_requirements([ABSENT_ESSENTIAL])
        assert 'no_such_module_xyz' not in sys.modules


class TestSeverity:
    def test_essential_and_optional_are_separated(self):
        results = check_requirements([PRESENT, ABSENT_ESSENTIAL, ABSENT_OPTIONAL])
        assert [r.requirement.package for r in missing_essential(results)] == ['ghost-pkg']
        assert [r.requirement.package for r in missing_optional(results)] == ['optional-pkg']

    def test_nothing_missing_when_all_present(self):
        results = check_requirements([PRESENT])
        assert missing_essential(results) == []
        assert missing_optional(results) == []


class TestRealRequirements:
    def test_openpyxl_is_essential(self):
        openpyxl = next(r for r in REQUIREMENTS if r.package == 'openpyxl')
        assert openpyxl.essential is True

    def test_pywin32_is_optional_not_fatal(self):
        """Steps 1-3 work without it, so a missing pywin32 must not block startup."""
        pywin32 = next(r for r in REQUIREMENTS if r.package == 'pywin32')
        assert pywin32.essential is False

    def test_pywin32_is_checked_by_its_import_name(self):
        """'pywin32' is not importable - 'pythoncom' is."""
        pywin32 = next(r for r in REQUIREMENTS if r.package == 'pywin32')
        assert pywin32.module == 'pythoncom'

    def test_openpyxl_is_actually_installed_here(self):
        results = check_requirements()
        assert missing_essential(results) == [], "openpyxl is needed to run the app"


class TestReport:
    def test_report_names_the_package_to_install(self):
        text = format_report(check_requirements([ABSENT_OPTIONAL]))
        assert 'pip install optional-pkg' in text

    def test_report_explains_what_breaks(self):
        text = format_report(check_requirements([ABSENT_OPTIONAL]))
        assert 'Something nice to have' in text

    def test_report_names_the_interpreter_when_something_is_missing(self):
        """The py-vs-python confusion is the whole reason this was hard."""
        text = format_report(check_requirements([ABSENT_OPTIONAL]))
        assert 'Running:' in text
        assert "'py' and 'python' can be different" in text

    def test_clean_report_has_no_scary_warnings(self):
        text = format_report(check_requirements([PRESENT]))
        assert '⚠️' not in text and '❌' not in text

    def test_interpreter_hint_includes_the_executable(self):
        import sys
        assert sys.executable in interpreter_hint()
