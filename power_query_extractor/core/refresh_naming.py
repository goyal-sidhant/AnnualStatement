"""Recognising refreshed report files.

The suffix appended when a report is refreshed is user-configurable ("Refreshed
File Suffix", default ``_Refreshed_{timestamp}``). Two places need to find those
files again afterwards: the client list, to show when each report was last
refreshed, and the "skip refresh" option, to reuse an existing refreshed file.

Both used to hardcode ``_Refreshed_``, so changing the setting silently broke
them - reports showed "Never refreshed" straight after a successful run, and
skip-refresh re-did work it should have reused. This module derives the match
from the configured pattern instead, so all three agree.
"""

DEFAULT_SUFFIX_PATTERN = "_Refreshed_{timestamp}"
TIMESTAMP_TOKEN = "{timestamp}"

# Files created under the default suffix. Always accepted in addition to the
# configured pattern, so changing the setting does not make the whole existing
# history read as "Never refreshed".
LEGACY_MARKER = "_refreshed_"


def refresh_markers(suffix_pattern=None):
    """Literal, lower-cased fragments a refreshed file's name must contain.

    The timestamp is a wildcard, so only the literal parts around it identify
    the file:

        "_Refreshed_{timestamp}"  -> ["_refreshed_"]
        "_Updated_{timestamp}"    -> ["_updated_"]
        "_v{timestamp}_final"     -> ["_v", "_final"]

    A pattern with no literal part at all (just "{timestamp}") cannot
    distinguish a refreshed file from the original, so the legacy marker is used
    rather than matching everything.
    """
    if not suffix_pattern:
        return [LEGACY_MARKER]
    parts = [p for p in suffix_pattern.lower().split(TIMESTAMP_TOKEN) if p]
    return parts or [LEGACY_MARKER]


def is_refreshed_name(filename, suffix_pattern=None):
    """True if `filename` looks like a refreshed report.

    Matches the configured pattern OR the legacy ``_Refreshed_`` marker, so
    files produced before the setting was changed are still recognised.
    """
    name = filename.lower()
    if not name.endswith('.xlsx'):
        return False
    markers = refresh_markers(suffix_pattern)
    if all(marker in name for marker in markers):
        return True
    return LEGACY_MARKER in name


def is_refreshed_copy_of(filename, original_stem, suffix_pattern=None):
    """True if `filename` is a refreshed copy of the report named `original_stem`.

    Replaces the hardcoded ``glob(f"{stem}_Refreshed_*.xlsx")`` lookup.
    """
    name = filename.lower()
    if not name.startswith(original_stem.lower()):
        return False
    return is_refreshed_name(filename, suffix_pattern)
