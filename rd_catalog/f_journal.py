"""Column-F journal line format, date-ordered insert, and D/E proposals.

Qt-free. Does not write Google. Callers pass the current ``comment_raw`` and
apply the returned patch after confirmation.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import datetime
from difflib import SequenceMatcher
from html import escape as html_escape

from rd_catalog.kits import (
    F_LINE_MTO_ABSENT,
    KitEvent,
    format_revision,
    parse_history_comment,
    parse_history_line,
    parse_sheet_revision,
    revisions_equivalent,
    strip_history_line_suffix,
)

_STAGE_F_TEXT: dict[str, str] = {
    "code_a": "код А",
    "code_b": "код B",
    "code_c": "код C",
    "incoming_passed": "прошла вх контр",
    "sr_upload": "отпр. на загрузку в СР",
    "tdo_sent": "отпр на ТДО",
}
JOURNAL_STAGE_LABELS: dict[str, str] = _STAGE_F_TEXT

_STATUS_BY_STAGE: dict[str, str] = {
    "code_a": "РД Согласовано",
    "code_b": "РД_Корректировка по зам.",
    "code_c": "РД_Корректировка по зам.",
    "incoming_passed": "Прошла входной контроль",
    "sr_upload": "Отпр. на входной контроль",
    "tdo_sent": "Отпр. на входной контроль",
}
_DATE_PREFIX_RE = re.compile(r"^\d{1,2}\.\d{1,2}\.\d{4}\s*(.*)$")
_STUB_LEFTOVER_RE = re.compile(r"[\s.,;:()\[\]{}\-_–—/\\]+")


@dataclass(frozen=True, slots=True)
class JournalPatch:
    """One proposed edit of a KSB ИД row (F, optionally D/E)."""

    comment_before: str
    comment_after: str
    action: str
    f_line: str
    update_de: bool
    sheet_revision: str
    status_sheet: str
    added_start: int
    added_end: int


def journal_stage_key(text: str) -> str:
    """Map a stage key or Russian F-line label to the classifier key.

    Args:
        text: ``code_a`` or ``код А`` (any ``JOURNAL_STAGE_LABELS`` value).

    Returns:
        Classifier key, or empty string when ``text`` is unknown.
    """

    raw = (text or "").strip()
    if not raw:
        return ""
    if raw in _STAGE_F_TEXT:
        return raw
    folded = raw.casefold()
    for key, label in _STAGE_F_TEXT.items():
        if label.casefold() == folded:
            return key
    return ""


def catalog_f_line(
    *,
    date: str,
    stage: str,
    revision: str | None = None,
    transmittal: str | None = None,
    mto_revision: str | None = None,
    mto_absent: bool = False,
) -> str:
    """Build an F line written by the catalog (always ``auto``).

    Shared by «Письма о согласовании» and folder legalize. Callers that
    parse a letter still gather date/stage/rev/MTO; this wrapper only
    stamps ``auto``.

    Args:
        date: ``DD.MM.YYYY``.
        stage: Classifier key (``code_a``, ``incoming_passed``, …).
        revision: Filename/OD revision text such as ``04`` or ``01-AN02``.
        transmittal: TRM token, placeholder comment, or empty.
        mto_revision: Optional MTO filename revision.
        mto_absent: When True and ``mto_revision`` is empty, append
            ``MTO Нет``.

    Returns:
        Canonical F line from :func:`format_history_line`.

    Raises:
        ValueError: If ``date`` or ``stage`` cannot be formatted.
    """

    return format_history_line(
        date=date,
        stage=stage,
        revision=revision,
        transmittal=transmittal,
        mto_revision=mto_revision,
        mto_absent=mto_absent,
        from_robot_auto=True,
    )


def format_history_line(
    *,
    date: str,
    stage: str,
    revision: str | None = None,
    transmittal: str | None = None,
    mto_revision: str | None = None,
    mto_absent: bool = False,
    from_robot_auto: bool = False,
) -> str:
    """Build one canonical F line that ``parse_history_line`` understands.

    Args:
        date: ``DD.MM.YYYY``.
        stage: Classifier key (``code_a``, ``incoming_passed``, …).
        revision: Filename/OD revision text such as ``04`` or ``01-AN02``.
            Codes use ``на рев. X``; other stages (cover ``tdo_sent``, …)
            use ``рев. X`` so ``parse_history_line`` still sees the token.
        transmittal: TRM token, or empty.
        mto_revision: Optional MTO filename revision (``03``, ``01-AN02``).
            Appended as ``MTO <rev>`` after TRM — never ``MTO рев. X``.
        mto_absent: When True and ``mto_revision`` is empty, append
            ``MTO Нет`` (transfer listed no MTO file).
        from_robot_auto: When True, append ``auto`` (catalog writer).

    Returns:
        A single-line journal entry. Empty MTO, False ``mto_absent``, and
        False ``auto`` match the historical string.

    Raises:
        ValueError: If ``date`` or ``stage`` cannot be formatted.
    """

    if not _parse_date_sort(date):
        raise ValueError(f"F line date must be DD.MM.YYYY, got {date!r}")
    stage_text = _STAGE_F_TEXT.get(stage)
    if not stage_text:
        raise ValueError(f"Unsupported F stage {stage!r}")
    parts = [date.strip(), stage_text]
    rev_text = _revision_display(revision)
    if rev_text:
        if stage.startswith("code_"):
            parts.append(f"на рев. {rev_text}")
        else:
            parts.append(f"рев. {rev_text}")
    trm = (transmittal or "").strip()
    if trm:
        parts.append(trm)
    mto_text = _revision_display(mto_revision)
    if mto_text:
        parts.append(f"MTO {mto_text}")
    elif mto_absent:
        parts.append(f"MTO {F_LINE_MTO_ABSENT}")
    if from_robot_auto:
        parts.append("auto")
    return " ".join(parts)


def added_history_line(
    comment_before: str,
    comment_after: str,
    *,
    fallback: str = "",
) -> str:
    """Return the single new or changed F line in ``comment_after``.

    Used when the user may edit «F после» before a Google write. The write
    engine still applies one :class:`~rd_catalog.google_f_write.JournalWriteJob`
    line to the live cell.

    Args:
        comment_before: Current column F (F до).
        comment_after: Edited column F (F после).
        fallback: Proposed new line when after still contains it (no
            unique added line).

    Returns:
        One stripped F line.

    Raises:
        ValueError: No unique new/changed line.
    """

    before_set = {line.strip() for line in _split_comment_lines(comment_before)}
    after_lines = _split_comment_lines(comment_after)
    added = [line for line in after_lines if line.strip() not in before_set]
    fallback_stripped = (fallback or "").strip()
    if len(added) == 1:
        return added[0].strip()
    after_set = {line.strip() for line in after_lines}
    if fallback_stripped and fallback_stripped in after_set:
        return fallback_stripped
    raise ValueError(
        "В «F после» должна быть ровно одна новая или изменённая "
        "строка журнала."
    )


def apply_history_line(comment_raw: str, new_line: str) -> tuple[str, str]:
    """Insert or replace one F line without rewriting the rest of the cell.

    Identity is date + stage + OD revision + transmittals (MTO/auto
    suffix ignored). Equal full raw → ``unchanged``; same core with a
    different suffix (for example adding ``MTO 03``) → ``replaced``.
    A truncated same-date stub (date-only, or date plus TRM without a
    classified stage) is upgraded when revision/TRM do not conflict.
    Otherwise the new line is inserted before the first existing dated
    line with a later calendar date. Undated lines keep their positions.

    Args:
        comment_raw: Current column F text.
        new_line: Canonical line from :func:`format_history_line`.

    Returns:
        ``(new_comment, action)`` where action is ``inserted``, ``replaced``,
        or ``unchanged`` when the same event is already a full line in F.
    """

    new_event = parse_history_line(new_line)
    if not new_event.date or new_event.stage == "other":
        raise ValueError(f"New F line is not a classified dated event: {new_line!r}")
    lines = _split_comment_lines(comment_raw)
    for index, line in enumerate(lines):
        if not _same_journal_event(line, new_line):
            continue
        if line.strip() == new_line.strip():
            return comment_raw or "", "unchanged"
        lines[index] = new_line
        return _join_comment_lines(lines), "replaced"
    new_sort = _parse_date_sort(new_event.date)
    assert new_sort is not None
    for index, line in enumerate(lines):
        existing = parse_history_line(line)
        if existing.date == new_event.date and existing.stage == new_event.stage:
            lines[index] = new_line
            return _join_comment_lines(lines), "replaced"
        if _truncated_stub_upgrades_to(existing, new_event):
            lines[index] = new_line
            return _join_comment_lines(lines), "replaced"
    insert_at = len(lines)
    for index, line in enumerate(lines):
        existing = parse_history_line(line)
        existing_sort = _parse_date_sort(existing.date)
        if existing_sort is not None and existing_sort > new_sort:
            insert_at = index
            break
    lines.insert(insert_at, new_line)
    return _join_comment_lines(lines), "inserted"


def build_journal_patch(
    comment_raw: str,
    new_line: str,
    *,
    revision: str | None = None,
    stage: str | None = None,
) -> JournalPatch:
    """Build a previewable F patch and D/E proposal.

    D/E update only when the new line is the last dated event in cell order
    after the edit (same rule as ``last_event``).

    Args:
        comment_raw: Current column F.
        new_line: Canonical F line.
        revision: OD/kit revision for column D (``Рев. …``).
        stage: Stage key; inferred from ``new_line`` when omitted.

    Returns:
        A :class:`JournalPatch`.
    """

    after, action = apply_history_line(comment_raw, new_line)
    parsed = parse_history_line(new_line)
    resolved_stage = stage or parsed.stage
    start = after.find(new_line)
    end = start + len(new_line) if start >= 0 else -1
    return JournalPatch(
        comment_before=comment_raw or "",
        comment_after=after,
        action=action,
        f_line=new_line,
        update_de=_is_last_dated_event(after, new_line),
        sheet_revision=format_sheet_revision_cell(revision),
        status_sheet=status_sheet_for_stage(resolved_stage),
        added_start=start,
        added_end=end,
    )


def format_sheet_revision_cell(revision: str | None) -> str:
    """Format column D as ``Рев. 04`` / ``Рев. 01-AN02``.

    Args:
        revision: Raw revision token from the letter or filename.

    Returns:
        Sheet cell text, or empty when revision is missing.
    """

    parsed_rev, appendix = parse_sheet_revision(revision or "")
    text = format_revision(parsed_rev, appendix)
    if not text:
        return ""
    return f"Рев. {text}"


def journal_write_needed(
    patch: JournalPatch,
    *,
    live_d: str = "",
    live_e: str = "",
) -> tuple[bool, bool, bool]:
    """Decide which KSB ИД cells still need a write.

    F is skipped when the same event is already in the journal. D and E are
    written only when this line is the last dated event and the live cell
    does not already match the proposal.

    Args:
        patch: Previewable F/D/E patch.
        live_d: Current column D text.
        live_e: Current column E text.

    Returns:
        ``(write_f, write_d, write_e)``.
    """

    write_f = patch.action != "unchanged"
    if not patch.update_de:
        return write_f, False, False
    write_d = bool(patch.sheet_revision) and not _d_cell_matches(
        live_d, patch.sheet_revision
    )
    write_e = bool(patch.status_sheet) and not _e_cell_matches(
        live_e, patch.status_sheet
    )
    return write_f, write_d, write_e


def journal_highlight_spans(
    before: str, after: str
) -> tuple[tuple[tuple[int, int], ...], tuple[tuple[int, int], ...]]:
    """Return character ranges that differ in ``before`` and ``after``.

    Line-oriented journals (including chained/cumulative F cells) are
    compared as raw text. Equal cells yield empty spans. Whitespace-only
    opcodes are skipped so a marker lands on the changed tokens.

    Args:
        before: Current column F (F до).
        after: Proposed column F (F после).

    Returns:
        ``(before_spans, after_spans)`` of ``(start, end)`` half-open
        ranges into the corresponding string.
    """

    left = before or ""
    right = after or ""
    if left == right:
        return (), ()
    matcher = SequenceMatcher(a=left, b=right, autojunk=False)
    before_spans: list[tuple[int, int]] = []
    after_spans: list[tuple[int, int]] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag in {"replace", "delete"}:
            span = _nonempty_span(left, i1, i2)
            if span is not None:
                before_spans.append(span)
        if tag in {"replace", "insert"}:
            span = _nonempty_span(right, j1, j2)
            if span is not None:
                after_spans.append(span)
    return tuple(before_spans), tuple(after_spans)


def journal_diff_html(
    text: str,
    spans: Sequence[tuple[int, int]],
    *,
    mark_color: str = "#FFE082",
    fg: str | None = None,
) -> str:
    """Render F text with yellow-marker spans as HTML.

    Args:
        text: Cell text.
        spans: Half-open ranges from :func:`journal_highlight_spans`.
        mark_color: CSS background for differing characters.
        fg: Optional CSS color for the whole cell.

    Returns:
        HTML fragment. Newlines become ``<br/>``.
    """

    raw = text or ""
    marks = _clip_spans(raw, spans)
    parts: list[str] = []
    if fg:
        parts.append(f'<span style="color:{html_escape(fg, quote=True)}">')
    pos = 0
    for start, end in marks:
        if start > pos:
            parts.append(_html_plain(raw[pos:start]))
        parts.append(
            f'<span style="background-color:{html_escape(mark_color, quote=True)}">'
            f"{_html_plain(raw[start:end])}</span>"
        )
        pos = end
    if pos < len(raw):
        parts.append(_html_plain(raw[pos:]))
    if fg:
        parts.append("</span>")
    return "".join(parts)


def status_sheet_for_stage(stage: str) -> str:
    """Return column E text for a journal stage.

    Code A → ``РД Согласовано``. Codes B and C → ``РД_Корректировка по зам.``
    (not incoming-control). Incoming-passed stays ``Прошла входной контроль``.

    Args:
        stage: Classifier key.

    Returns:
        Status cell, or empty when the stage does not drive E.
    """

    return _STATUS_BY_STAGE.get(stage, "")


def _d_cell_matches(live: str, proposed: str) -> bool:
    live_rev, live_app = parse_sheet_revision(live)
    proposed_rev, proposed_app = parse_sheet_revision(proposed)
    if not proposed_rev:
        return True
    if live_rev:
        return revisions_equivalent(live_rev, live_app, proposed_rev, proposed_app)
    return live.strip() == proposed.strip()


def _e_cell_matches(live: str, proposed: str) -> bool:
    return live.strip().casefold() == proposed.strip().casefold()


def _truncated_stub_upgrades_to(existing: KitEvent, new_event: KitEvent) -> bool:
    """Return True when *existing* is a safe truncated form of *new_event*.

    Safe stubs: a date-only line, or date plus transmittal token(s) with no
    classified stage. A free-form note on that date, a conflicting TRM or
    revision, or a different classified stage must not be replaced.

    Args:
        existing: Parsed current F line.
        new_event: Parsed canonical line about to be applied.

    Returns:
        True when the stub can be replaced by the fuller classified line.
    """

    if new_event.stage == "other" or not new_event.date:
        return False
    if not existing.date or existing.stage != "other":
        return False
    existing_sort = _parse_date_sort(existing.date)
    new_sort = _parse_date_sort(new_event.date)
    if existing_sort is None or existing_sort != new_sort:
        return False
    if (existing.revision or "") and (existing.revision or "") != (
        new_event.revision or ""
    ):
        return False
    existing_trm = {token.casefold() for token in existing.transmittals}
    new_trm = {token.casefold() for token in new_event.transmittals}
    if existing_trm - new_trm:
        return False
    rest = _line_rest_after_date(existing.raw)
    rest, _mto_rev, _mto_app, _mto_absent, _auto = strip_history_line_suffix(
        rest
    )
    if not rest:
        return True
    if not existing.transmittals:
        return False
    leftover = rest
    for token in existing.transmittals:
        leftover = re.sub(re.escape(token), " ", leftover, flags=re.IGNORECASE)
    leftover = _STUB_LEFTOVER_RE.sub("", leftover)
    return leftover == ""


def _line_rest_after_date(raw: str) -> str:
    match = _DATE_PREFIX_RE.match((raw or "").strip())
    if match is None:
        return (raw or "").strip()
    return match.group(1).strip()


def _same_journal_event(existing_line: str, new_line: str) -> bool:
    """Return True when F already contains the same dated letter event.

    Identity is date + stage + OD revision + transmittals. MTO/auto suffix
    is ignored so a suffix upgrade can replace the existing line.
    """

    if existing_line.strip() == new_line.strip():
        return True
    left = parse_history_line(existing_line)
    right = parse_history_line(new_line)
    if not left.date or left.stage == "other":
        return False
    if left.date != right.date or left.stage != right.stage:
        return False
    if (left.revision or "") != (right.revision or ""):
        return False
    left_trm = tuple(token.casefold() for token in left.transmittals)
    right_trm = tuple(token.casefold() for token in right.transmittals)
    return left_trm == right_trm


def _is_last_dated_event(comment: str, new_line: str) -> bool:
    events = parse_history_comment(comment)
    dated = [event for event in events if event.date]
    if not dated:
        return True
    last = dated[-1]
    new_event = parse_history_line(new_line)
    return last.date == new_event.date and last.stage == new_event.stage


def _split_comment_lines(comment: str) -> list[str]:
    text = (comment or "").replace("\r\n", "\n").replace("\r", "\n")
    return [line for line in text.split("\n") if line.strip()]


def _join_comment_lines(lines: list[str]) -> str:
    return "\n".join(lines)


def _parse_date_sort(date_text: str | None) -> datetime | None:
    raw = (date_text or "").strip()
    try:
        return datetime.strptime(raw, "%d.%m.%Y")
    except ValueError:
        return None


def _revision_display(revision: str | None) -> str:
    parsed_rev, appendix = parse_sheet_revision(revision or "")
    return format_revision(parsed_rev, appendix)


def _nonempty_span(text: str, start: int, end: int) -> tuple[int, int] | None:
    if start >= end:
        return None
    chunk = text[start:end]
    if not chunk.strip():
        return None
    return start, end


def _clip_spans(
    text: str, spans: Sequence[tuple[int, int]]
) -> tuple[tuple[int, int], ...]:
    clipped: list[tuple[int, int]] = []
    length = len(text)
    for start, end in spans:
        lo = max(0, min(int(start), length))
        hi = max(lo, min(int(end), length))
        if hi > lo:
            clipped.append((lo, hi))
    clipped.sort()
    return tuple(clipped)


def _html_plain(chunk: str) -> str:
    return html_escape(chunk, quote=False).replace("\n", "<br/>")
