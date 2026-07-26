"""Tests for the main window's tab set, including Step 4.

Runs in a subprocess: GSTOrganizerApp creates its own tk.Tk root, and a second
live root in a process makes Tk fail intermittently (see
tests/test_extractor_standalone.py for the same reasoning).
"""
import json
import subprocess
import sys
import textwrap
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
XL = b'PK\x03\x04' + bytes(2048)

PROBE = textwrap.dedent(
    """
    import json, sys, os
    sys.path.insert(0, {project!r})
    os.chdir({project!r})

    out = {{}}
    try:
        # Count scans so we can prove startup performs none: a scan at launch
        # would hit the network every time the app opens.
        import power_query_extractor.gui.extractor_window as ew
        scans = []
        _orig = ew.ExtractorPanel.scan_folder
        def counting(self, *a, **k):
            scans.append(self.folder_path.get())
            return _orig(self, *a, **k)
        ew.ExtractorPanel.scan_folder = counting

        from gui.main_window import GSTOrganizerApp
        app = GSTOrganizerApp()
        app.save_cache = lambda *a, **k: None
        app.target_folder.set({target!r})
        app.root.update()

        out['tabs'] = [app.notebook.tab(i)['text']
                       for i in range(app.notebook.index('end'))]
        out['scans_at_startup'] = list(scans)
        out['folder_prefilled'] = app.extract_tab.panel.folder_path.get() == {target!r}
        out['clients_empty_before_scan'] = app.extract_tab.panel.client_vars == {{}}
        out['panel_header_suppressed'] = app.extract_tab.panel.show_header is False
        out['panel_statusbar_suppressed'] = app.extract_tab.panel.show_status_bar is False

        # Selecting every tab must not raise. A hardcoded 3-name list previously
        # raised IndexError as soon as a fourth tab existed.
        #
        # Tk CATCHES exceptions raised inside event callbacks and routes them to
        # report_callback_exception instead of propagating, so select() itself
        # never raises - the error has to be intercepted here or the test would
        # pass while the app was visibly broken.
        errors = []
        app.root.report_callback_exception = (
            lambda exc, val, tb: errors.append(f'{{exc.__name__}}: {{val}}'))
        for i in range(app.notebook.index('end')):
            try:
                app.notebook.select(i)
                app.root.update()
            except Exception as exc:
                errors.append(f'tab {{i}}: {{type(exc).__name__}}: {{exc}}')
        out['tab_select_errors'] = errors

        # Scanning on demand still works
        app.extract_tab.panel.scan_folder()
        app.root.update()
        out['clients_after_scan'] = sorted(app.extract_tab.panel.client_vars)

        app.root.destroy()
        out['ok'] = True
    except Exception as exc:
        import traceback
        out['ok'] = False
        out['error'] = traceback.format_exc()[-600:]

    print('@@RESULT@@' + json.dumps(out))
    """
)


@pytest.fixture(scope="module")
def probe(tmp_path_factory):
    target = tmp_path_factory.mktemp("mainwin")
    version = target / "Annual Statement-250726 1000" / "ABC-DL" / "Version-250726 1000"
    version.mkdir(parents=True)
    (version / "ITC_Report_ABC_DL.xlsx").write_bytes(XL)
    (version / "Sales_Report_ABC_DL.xlsx").write_bytes(XL)

    code = PROBE.format(project=str(PROJECT_ROOT), target=str(target))
    result = subprocess.run([sys.executable, "-c", code],
                            capture_output=True, text=True, timeout=300,
                            cwd=str(PROJECT_ROOT))
    marker = '@@RESULT@@'
    line = next((l for l in result.stdout.splitlines() if l.startswith(marker)), None)
    if line is None:
        pytest.skip(f"main window could not start (rc={result.returncode}): "
                    f"{result.stderr[-400:]}")
    return json.loads(line[len(marker):])


def test_main_window_starts(probe):
    assert probe['ok'], probe.get('error')


def test_has_four_steps_ending_with_extract(probe):
    tabs = probe['tabs']
    assert len(tabs) == 4
    assert 'Setup' in tabs[0]
    assert 'Validation' in tabs[1]
    assert 'Processing' in tabs[2]
    assert 'Extract' in tabs[3]


def test_selecting_any_tab_does_not_raise(probe):
    """Regression: on_tab_changed indexed a hardcoded 3-name list, so selecting
    the fourth tab raised IndexError on every tab change."""
    assert probe['tab_select_errors'] == []


def test_startup_does_not_scan_the_network(probe):
    assert probe['scans_at_startup'] == []


def test_step4_prefills_the_target_folder(probe):
    assert probe['folder_prefilled']
    assert probe['clients_empty_before_scan']


def test_step4_suppresses_its_own_chrome(probe):
    """The main window already provides a banner and a status bar."""
    assert probe['panel_header_suppressed']
    assert probe['panel_statusbar_suppressed']


def test_scanning_on_demand_finds_clients(probe):
    assert probe['clients_after_scan'] == ['ABC-DL']
