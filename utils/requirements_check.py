"""Check that the app's dependencies are installed, at startup.

Motivated by a real failure: the app died with "No module named 'pythoncom'"
and a message blaming missing source files. Two things made that hard to
diagnose - the message pointed at the wrong cause, and `py` and `python` can be
different interpreters with different packages installed, so "it works for me"
and "it fails for you" can both be true on one machine.

No Tkinter here, so it can be unit-tested and can also run before the GUI is
imported.
"""
import importlib.util
import sys
from typing import List, NamedTuple


ON_WINDOWS = sys.platform == 'win32'


class Requirement(NamedTuple):
    module: str          # what to import
    package: str         # what to install
    purpose: str
    essential: bool      # False = app still works, but something is degraded
    consequence: str = ''   # what goes wrong without it, for the message


class Result(NamedTuple):
    requirement: Requirement
    installed: bool

    @property
    def essential(self) -> bool:
        return self.requirement.essential


REQUIREMENTS = (
    Requirement(
        module='openpyxl',
        package='openpyxl',
        purpose='Reading and writing Excel files',
        essential=True,
    ),
    Requirement(
        module='pythoncom',
        package='pywin32',
        purpose=('Excel automation: preserves Power Query when creating reports, '
                 'and powers the refresh in Step 4'),
        # ESSENTIAL on Windows. Without Excel COM, report creation silently falls
        # back to openpyxl, which does not preserve Power Query - so the ITC and
        # Sales reports come out stripped of the queries that make them useful,
        # with nothing obviously wrong at the time. Producing damaged reports is
        # worse than refusing to start, so this blocks.
        #
        # Off Windows, COM does not exist at all and nothing can be done about
        # it, so there it stays a warning rather than a dead end.
        essential=ON_WINDOWS,
        consequence=('Reports would be created WITHOUT their Power Query '
                     'connections, and Step 4 could not refresh them.'),
    ),
)


def _is_installed(module_name: str) -> bool:
    if module_name in sys.modules:
        return True
    try:
        return importlib.util.find_spec(module_name) is not None
    except (ImportError, ValueError):
        return False


def check_requirements(requirements=REQUIREMENTS) -> List[Result]:
    return [Result(req, _is_installed(req.module)) for req in requirements]


def missing_essential(results: List[Result]) -> List[Result]:
    """Requirements the app genuinely cannot run without."""
    return [r for r in results if r.essential and not r.installed]


def missing_optional(results: List[Result]) -> List[Result]:
    """Requirements whose absence degrades the app but does not stop it."""
    return [r for r in results if not r.essential and not r.installed]


def interpreter_hint() -> str:
    """Which interpreter is running - the detail that made this confusing."""
    version = '.'.join(str(p) for p in sys.version_info[:3])
    return f"Python {version} at {sys.executable}"


def format_report(results: List[Result]) -> str:
    """A short human-readable summary for the console."""
    lines = []
    for result in results:
        mark = 'OK     ' if result.installed else ('MISSING' if result.essential
                                                   else 'not found')
        lines.append(f"   [{mark}] {result.requirement.package:10} - "
                     f"{result.requirement.purpose}")

    report = ["Checking requirements...", *lines]

    for result in missing_optional(results):
        report += [
            "",
            f"⚠️  {result.requirement.package} is not installed for this Python.",
            f"   Without it: {result.requirement.purpose}.",
        ]
        if result.requirement.consequence:
            report += [f"   {result.requirement.consequence}"]
        report += [f"   Install with:  pip install {result.requirement.package}"]

    for result in missing_essential(results):
        report += [
            "",
            f"❌ {result.requirement.package} is REQUIRED and is not installed.",
            f"   Needed for: {result.requirement.purpose}.",
        ]
        if result.requirement.consequence:
            report += [f"   {result.requirement.consequence}",
                       "   The app stops here rather than produce damaged output."]
        report += [f"   Install with:  pip install {result.requirement.package}"]

    if missing_optional(results) or missing_essential(results):
        report += [
            "",
            f"   Running: {interpreter_hint()}",
            "   Note 'py' and 'python' can be different interpreters - install",
            "   for the one you start the app with.",
        ]

    return "\n".join(report)
