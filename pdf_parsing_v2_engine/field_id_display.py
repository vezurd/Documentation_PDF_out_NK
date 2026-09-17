"""Short display label derived from a stamp field id (editor canvas and v2 Excel report)."""

from __future__ import annotations

import re

_NUMERIC_FIELD_PREFIX_RE = re.compile(r"^(\d+(?:_\d+)*)")


def short_label_from_field_id(field_id: str) -> str:
    """Return a compact label for *field_id* used in UI and debug report headers.

    Uses the longest leading run of numeric segments separated by underscores; if
    none, a trailing ``_digits`` suffix; otherwise the first ``_``-separated segment.

    Args:
        field_id: Template/catalog field identifier (may be empty).

    Returns:
        Short label string, or empty string if *field_id* is empty.
    """
    if not field_id:
        return ""
    m = _NUMERIC_FIELD_PREFIX_RE.match(field_id)
    if m:
        return m.group(1)
    m_tail = re.search(r"_(\d+)$", field_id)
    if m_tail:
        return m_tail.group(1)
    return field_id.split("_")[0]
