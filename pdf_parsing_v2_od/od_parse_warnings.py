"""OD parse diagnostics as compact ``warnings`` lines (НК + вкладка ОД)."""

from __future__ import annotations

# Four-digit codes (same family as other normcontrol-adjacent diagnostics, e.g. 21xx extraction).
PAGE_FORMAT_WHITESPACE_WARNING_CODE = 3001
PAGE_FORMAT_UNKNOWN_KEY_CODE = 3002


def od_parse_warning_line(code: int, message: str) -> str:
    """One line for ``warnings_out`` / ``od_table['warnings']``: ``\"3001: …\"``."""
    return f"{int(code)}: {message}"


def append_od_parse_warning(
    warnings_out: list[str] | None,
    *,
    code: int,
    message: str,
) -> None:
    if warnings_out is None:
        return
    warnings_out.append(od_parse_warning_line(code, message))
