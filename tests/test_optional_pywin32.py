"""The app must start without pywin32.

pywin32 is optional: organising files and creating reports (Steps 1-3) works
without it, and core/excel_handler.py already degrades gracefully. But
report_processor imported pythoncom/win32* unconditionally, and once the main
window hosted the extractor as Step 4 that import ran at startup - so the whole
app died on a Python without pywin32 instead of just losing Step 4.

Note 'py' and 'python' can resolve to different interpreters with different
packages installed, so this is easy to hit in practice.
"""
import json
import subprocess
import sys
import textwrap
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent

# Runs in a subprocess with the win32 modules blocked, simulating a Python that
# has no pywin32 regardless of what this machine actually has installed.
PROBE = textwrap.dedent(
    """
    import builtins, json, sys
    sys.path.insert(0, {project!r})

    BLOCKED = {{'pythoncom', 'win32com', 'win32com.client',
                'win32api', 'win32con', 'win32gui'}}
    _real_import = builtins.__import__

    def guarded(name, *args, **kwargs):
        if name in BLOCKED or name.split('.')[0] in BLOCKED:
            raise ImportError(f"No module named {{name!r}} (blocked for test)")
        return _real_import(name, *args, **kwargs)

    builtins.__import__ = guarded
    for mod in list(sys.modules):
        if mod.split('.')[0] in BLOCKED:
            del sys.modules[mod]

    out = {{}}
    try:
        from power_query_extractor.core import report_processor
        out['win32_available'] = report_processor.WIN32_AVAILABLE

        # Refusing clearly beats failing mysteriously
        proc = report_processor.ReportProcessor()
        result = proc.process_client({{'name': 'ABC-DL'}})
        out['refused_cleanly'] = result['itc']['status']['success'] is False
        out['mentions_pywin32'] = 'pywin32' in result['itc']['status']['error']

        from gui.tabs import extract_tab
        out['tab_module_imports'] = True

        import gui.main_window
        out['main_window_imports'] = True
        out['ok'] = True
    except Exception as exc:
        import traceback
        out['ok'] = False
        out['error'] = traceback.format_exc()[-500:]

    print('@@RESULT@@' + json.dumps(out))
    """
)


@pytest.fixture(scope="module")
def probe():
    code = PROBE.format(project=str(PROJECT_ROOT))
    result = subprocess.run([sys.executable, "-c", code],
                            capture_output=True, text=True, timeout=180,
                            cwd=str(PROJECT_ROOT))
    marker = '@@RESULT@@'
    line = next((l for l in result.stdout.splitlines() if l.startswith(marker)), None)
    if line is None:
        pytest.fail(f"probe produced no result (rc={result.returncode}): "
                    f"{result.stderr[-500:]}")
    return json.loads(line[len(marker):])


def test_probe_ran(probe):
    assert probe['ok'], probe.get('error')


def test_report_processor_imports_without_pywin32(probe):
    """Regression: unconditional `import pythoncom` took the whole app down."""
    assert probe['win32_available'] is False


def test_main_window_imports_without_pywin32(probe):
    """The whole point: Steps 1-3 must still be usable."""
    assert probe['tab_module_imports']
    assert probe['main_window_imports']


def test_processing_refuses_clearly_rather_than_crashing(probe):
    assert probe['refused_cleanly']
    assert probe['mentions_pywin32'], "the error should say how to fix it"
