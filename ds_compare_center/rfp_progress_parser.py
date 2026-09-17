"""Parse structured RFP pipeline milestone messages from process output."""

from __future__ import annotations

import codecs
import json
from dataclasses import dataclass

_MARKER_PREFIX = "@@RFP_MILESTONE "

MILESTONE_IDS = (
    "prepare",
    "ds_id_check",
    "ds_mp_check",
    "parts_preflight",
    "ul_preflight",
    "rfp_load",
    "mto_google_load",
    "vo_load",
    "units_gate",
    "step4_match",
    "global_checks",
    "packing_lists",
    "save_excel",
    "bcc_accum_matrix",
    "complete",
)
MILESTONE_STATES = ("Waiting", "Running", "Done", "Skipped", "Error")

_VALID_IDS = frozenset(MILESTONE_IDS)
_STATE_BY_LOWER = {state.lower(): state for state in MILESTONE_STATES}


@dataclass(frozen=True, slots=True)
class MilestoneEvent:
    """A validated milestone update emitted by the RFP pipeline."""

    milestone_id: str
    state: str
    detail: str = ""


class RfpProgressParser:
    """Line-buffer parser for ``@@RFP_MILESTONE`` JSON messages."""

    def __init__(self) -> None:
        self._buffer = ""
        self._decoder = codecs.getincrementaldecoder("utf-8")(errors="replace")

    def feed(self, chunk: str | bytes | bytearray) -> list[MilestoneEvent]:
        """Consume an arbitrary output chunk and return complete events.

        Args:
            chunk: Text or UTF-8 process-output bytes. It may split a marker,
                JSON payload, newline, or a multibyte character.

        Returns:
            Valid milestone events found in complete lines.
        """
        if isinstance(chunk, str):
            text = chunk
        else:
            text = self._decoder.decode(bytes(chunk), final=False)
        self._buffer += text

        lines = self._buffer.splitlines(keepends=True)
        if lines and not _has_line_ending(lines[-1]):
            self._buffer = lines.pop()
        else:
            self._buffer = ""
        return _parse_lines(lines)

    def flush(self) -> list[MilestoneEvent]:
        """Parse the final unterminated line and clear buffered input."""
        self._buffer += self._decoder.decode(b"", final=True)
        line = self._buffer
        self._buffer = ""
        self._decoder = codecs.getincrementaldecoder("utf-8")(errors="replace")
        return _parse_lines([line]) if line else []


def _has_line_ending(text: str) -> bool:
    return text.endswith(("\n", "\r"))


def _parse_lines(lines: list[str]) -> list[MilestoneEvent]:
    events: list[MilestoneEvent] = []
    for line in lines:
        event = _parse_line(line.rstrip("\r\n"))
        if event is not None:
            events.append(event)
    return events


def _parse_line(line: str) -> MilestoneEvent | None:
    if not line.startswith(_MARKER_PREFIX):
        return None

    try:
        payload = json.loads(line[len(_MARKER_PREFIX) :])
    except (json.JSONDecodeError, TypeError):
        return None
    if not isinstance(payload, dict):
        return None

    milestone_id = payload.get("milestone_id")
    state_value = payload.get("state")
    if not isinstance(milestone_id, str) or milestone_id not in _VALID_IDS:
        return None
    if not isinstance(state_value, str):
        return None
    state = _STATE_BY_LOWER.get(state_value.strip().lower())
    if state is None:
        return None

    detail_value = payload.get("detail", "")
    detail = "" if detail_value is None else str(detail_value)
    return MilestoneEvent(milestone_id=milestone_id, state=state, detail=detail)
