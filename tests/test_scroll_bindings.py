"""Guards against one widget's scroll binding destroying another's.

Reported: scroll worked on Step 1, but after visiting the other tabs and coming
back it was dead. Cause: the extractor's client list bound the wheel with
bind_all on <Enter> and called unbind_all on <Leave> - and unbind_all removes
EVERY application-level <MouseWheel> binding, not just its own. Merely hovering
that list and moving away killed the Setup tab's scrolling for the rest of the
session.
"""
import re
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
GUI_SOURCES = sorted(
    list((PROJECT_ROOT / 'gui').rglob('*.py'))
    + list((PROJECT_ROOT / 'power_query_extractor').rglob('*.py'))
)


def test_sources_were_found():
    assert GUI_SOURCES, "no GUI sources discovered - the guard below would be vacuous"


@pytest.mark.parametrize("path", GUI_SOURCES, ids=lambda p: p.name)
def test_no_unbind_all(path):
    """unbind_all is never safe here: it removes other widgets' bindings too."""
    source = path.read_text(encoding='utf-8')
    offenders = [line.strip() for line in source.splitlines()
                 if 'unbind_all(' in line and not line.strip().startswith('#')]
    assert not offenders, (
        f"{path.name} calls unbind_all, which removes EVERY application-level "
        f"binding including other tabs': {offenders}")


@pytest.mark.parametrize("path", GUI_SOURCES, ids=lambda p: p.name)
def test_bind_all_is_additive(path):
    """bind_all without add='+' REPLACES any existing binding for that event."""
    source = path.read_text(encoding='utf-8')
    bad = []
    for line in source.splitlines():
        stripped = line.strip()
        if stripped.startswith('#') or 'bind_all(' not in stripped:
            continue
        if "add='+'" not in stripped and 'add="+"' not in stripped:
            bad.append(stripped)
    assert not bad, (
        f"{path.name} uses bind_all without add='+', which replaces other "
        f"widgets' bindings for the same event: {bad}")


def test_wheel_handlers_check_the_pointer_is_over_them():
    """With a shared application-level binding, each handler must decide for
    itself whether the event belongs to it - otherwise every scrollable area
    reacts to every wheel event."""
    for name in ('gui/tabs/setup_tab.py',
                 'power_query_extractor/gui/extractor_window.py'):
        source = (PROJECT_ROOT / name).read_text(encoding='utf-8')
        if 'bind_all' not in source:
            continue
        assert re.search(r'def _?(is_)?in_scroll_area', source), (
            f"{name} binds the wheel application-wide but has no check that the "
            f"pointer is actually over its own scroll area")
