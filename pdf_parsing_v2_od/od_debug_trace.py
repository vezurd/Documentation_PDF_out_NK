"""Structured debug trace for OD parsing (optional; used by GUI standalone run)."""

from __future__ import annotations

MAX_OD_DEBUG_BODY_LEN = 16000


def append_od_debug_trace(
    trace: list[dict[str, str]] | None,
    *,
    function: str,
    stage: str,
    body: str,
) -> None:
    """Append one block to *trace* if the list is not ``None``.

    Args:
        trace: Target list (typically empty list from the caller), or ``None`` to skip.
        function: Dotted name of the logical unit (e.g. ``od_parsing.od_table_parsing_f``).
        stage: Short stage id / label (English id or Russian description).
        body: Multi-line detail; long bodies are truncated.
    """
    if trace is None:
        return
    b = body.strip()
    if len(b) > MAX_OD_DEBUG_BODY_LEN:
        b = b[:MAX_OD_DEBUG_BODY_LEN] + "\n... [truncated]"
    trace.append({"function": function, "stage": stage, "body": b})
