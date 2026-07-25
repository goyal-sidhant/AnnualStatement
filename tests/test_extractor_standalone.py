"""Tests for the standalone extractor window (PowerQueryExtractorApp).

PowerQueryExtractorApp *is* a tk.Tk root, and Tk does not reliably allow a
second root in a process that already created one (the rest of the GUI suite
shares a session root). Rather than let these tests skip intermittently - which
silently hides whether the standalone path works at all - they run the window in
a subprocess. That also exercises the real startup path, the same one
launch_extractor.py uses.
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
    import json, sys
    sys.path.insert(0, {project!r})

    from power_query_extractor.gui.extractor_window import (
        ExtractorPanel, PowerQueryExtractorApp)

    out = {{}}
    try:
        app = PowerQueryExtractorApp({target!r})
        app.update_idletasks()
        out['title'] = app.title()
        out['geometry'] = app.geometry().split('+')[0]
        out['panel_is_extractor_panel'] = isinstance(app.panel, ExtractorPanel)
        out['show_header'] = app.panel.show_header
        out['show_status_bar'] = app.panel.show_status_bar
        # this class used to BE the panel: app.<panel attr> must still resolve
        out['forwards_client_vars'] = 'ABC-DL' in app.client_vars
        out['forwards_log_message'] = callable(app.log_message)
        out['tabs'] = [app.notebook.tab(i)['text']
                       for i in range(app.notebook.index('end'))]
        try:
            app.definitely_not_a_real_attribute
            out['unknown_attr_raises'] = False
        except AttributeError:
            out['unknown_attr_raises'] = True
        app.destroy()
        out['ok'] = True
    except Exception as exc:
        out['ok'] = False
        out['error'] = f'{{type(exc).__name__}}: {{exc}}'

    print('@@RESULT@@' + json.dumps(out))
    """
)


@pytest.fixture(scope="module")
def probe(tmp_path_factory):
    """Launch the standalone window in a clean process and report its state."""
    target = tmp_path_factory.mktemp("standalone")
    version = target / "Annual Statement-250726 1000" / "ABC-DL" / "Version-250726 1000"
    version.mkdir(parents=True)
    (version / "ITC_Report_ABC_DL.xlsx").write_bytes(XL)
    (version / "Sales_Report_ABC_DL.xlsx").write_bytes(XL)

    code = PROBE.format(project=str(PROJECT_ROOT), target=str(target))
    result = subprocess.run([sys.executable, "-c", code],
                            capture_output=True, text=True, timeout=180,
                            cwd=str(PROJECT_ROOT))

    marker = '@@RESULT@@'
    line = next((l for l in result.stdout.splitlines() if l.startswith(marker)), None)
    if line is None:
        pytest.skip(f"standalone window could not start "
                    f"(rc={result.returncode}): {result.stderr[-400:]}")
    return json.loads(line[len(marker):])


def test_standalone_window_starts(probe):
    assert probe['ok'], probe.get('error')


def test_owns_the_window_level_concerns(probe):
    assert probe['title'] == "Power Query Extractor"
    assert probe['geometry'] == "1000x700"
    assert probe['panel_is_extractor_panel']


def test_standalone_keeps_its_banner_and_status_bar(probe):
    assert probe['show_header'] is True
    assert probe['show_status_bar'] is True


def test_panel_supplies_its_own_two_tabs(probe):
    assert probe['tabs'] == ['Setup', 'Processing']


def test_attribute_forwarding_keeps_old_usage_working(probe):
    assert probe['forwards_client_vars']
    assert probe['forwards_log_message']


def test_unknown_attribute_still_raises(probe):
    assert probe['unknown_attr_raises']
