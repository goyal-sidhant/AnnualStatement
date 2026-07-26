"""Tests for the offer-to-install helper.

The dialog itself is a thin shell; what matters and is tested here is that pip
targets the RIGHT interpreter, that failures are reported rather than swallowed,
and that the helper meant to fix a broken install can never itself break the app.
"""
import subprocess
import sys

import pytest

from utils.dependency_installer import (
    INSTALL_TIMEOUT_SECONDS, build_pip_command, describe_missing, install,
)
from utils.requirements_check import Requirement, Result


def make_result(package='pywin32', essential=True, consequence='Reports break.'):
    req = Requirement(module='pythoncom', package=package,
                      purpose='Excel automation', essential=essential,
                      consequence=consequence)
    return Result(req, False)


class TestPipCommand:
    def test_targets_the_running_interpreter(self):
        """`py` and `python` can differ - installing into the wrong one is the
        exact confusion this feature exists to end."""
        command = build_pip_command(['pywin32'])
        assert command[0] == sys.executable
        assert command[1:4] == ['-m', 'pip', 'install']

    def test_includes_named_packages(self):
        assert build_pip_command(['pywin32', 'openpyxl'])[-2:] == ['pywin32', 'openpyxl']

    def test_uses_requirements_file_when_given(self):
        command = build_pip_command(requirements_path='reqs.txt')
        assert command[-2:] == ['-r', 'reqs.txt']

    def test_requirements_file_takes_both_when_supplied(self):
        command = build_pip_command(['extra'], requirements_path='reqs.txt')
        assert '-r' in command and 'extra' in command


class TestInstall:
    def test_reports_success(self, monkeypatch):
        monkeypatch.setattr(subprocess, 'run',
                            lambda *a, **k: subprocess.CompletedProcess(
                                a[0], 0, 'installed ok', ''))
        ok, output = install(['pywin32'])
        assert ok is True
        assert 'installed ok' in output

    def test_reports_failure_with_output(self, monkeypatch):
        monkeypatch.setattr(subprocess, 'run',
                            lambda *a, **k: subprocess.CompletedProcess(
                                a[0], 1, '', 'could not reach proxy'))
        ok, output = install(['pywin32'])
        assert ok is False
        assert 'proxy' in output

    def test_timeout_is_reported_not_raised(self, monkeypatch):
        def boom(*a, **k):
            raise subprocess.TimeoutExpired(cmd='pip', timeout=INSTALL_TIMEOUT_SECONDS)
        monkeypatch.setattr(subprocess, 'run', boom)
        ok, output = install(['pywin32'])
        assert ok is False
        assert 'did not finish' in output

    def test_missing_pip_is_reported_not_raised(self, monkeypatch):
        def boom(*a, **k):
            raise FileNotFoundError('no pip here')
        monkeypatch.setattr(subprocess, 'run', boom)
        ok, output = install(['pywin32'])
        assert ok is False
        assert 'Could not run pip' in output

    def test_progress_is_streamed_to_the_callback(self, monkeypatch):
        monkeypatch.setattr(subprocess, 'run',
                            lambda *a, **k: subprocess.CompletedProcess(a[0], 0, 'done', ''))
        seen = []
        install(['pywin32'], log=seen.append)
        assert any('pip' in line for line in seen)
        assert any('done' in line for line in seen)


class TestDescribeMissing:
    def test_names_package_and_purpose(self):
        text = describe_missing([make_result()])
        assert 'pywin32' in text and 'Excel automation' in text

    def test_includes_the_consequence(self):
        assert 'Reports break.' in describe_missing([make_result()])

    def test_handles_several(self):
        text = describe_missing([make_result('pywin32'), make_result('openpyxl')])
        assert 'pywin32' in text and 'openpyxl' in text


class TestNeverBreaksTheApp:
    def test_offer_returns_false_when_tk_is_unavailable(self, monkeypatch):
        """The helper meant to FIX a broken environment must not itself crash
        the app if Tk cannot load."""
        import builtins
        from utils import dependency_installer

        real_import = builtins.__import__

        def no_tk(name, *args, **kwargs):
            if name == 'tkinter' or name.startswith('tkinter.'):
                raise ImportError('no tkinter')
            return real_import(name, *args, **kwargs)

        monkeypatch.setattr(builtins, '__import__', no_tk)
        assert dependency_installer.offer_install([make_result()]) is False
