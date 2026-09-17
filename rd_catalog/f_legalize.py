"""Legalize an RD folder revision as code A in KSB ИД column F.

Qt-free. Reuses :func:`catalog_f_line` / :func:`build_journal_patch` from
the same engine as «Письма о согласовании». Does not write Google.
"""

from __future__ import annotations

from collections.abc import Sequence
from datetime import datetime

from rd_catalog.doc_bundle import file_revision_label
from rd_catalog.f_journal import (
    added_history_line,
    catalog_f_line,
)
from rd_catalog.google_f_write import JournalWriteJob
from rd_catalog.kits import format_revision, parse_history_line, parse_sheet_revision
from rd_catalog.models import FileKind, FileRecord
from rd_catalog.overlay import revision_rank

LEGALIZE_APPROVAL_ACTION = "Легализовать согласование РД (F)…"
LEGALIZE_APPROVAL_STAGE = "code_a"
LEGALIZE_APPROVAL_TOKEN = "Добавлен_для_легализации_ревизии"
LEGALIZE_APPROVAL_TITLE = "Легализовать согласование РД"


def f_line_date_from_mtime_ns(value: int | None) -> str:
    """Format a file mtime as an F-line ``DD.MM.YYYY``.

    Args:
        value: Nanoseconds since epoch, or ``None``.

    Returns:
        Local calendar date, or today's date when unknown.
    """

    stamp: datetime | None = None
    if value:
        try:
            stamp = datetime.fromtimestamp(int(value) / 1_000_000_000)
        except (TypeError, ValueError, OSError):
            stamp = None
    if stamp is None:
        stamp = datetime.now()
    return f"{stamp.day:02d}.{stamp.month:02d}.{stamp.year:04d}"


def mto_revision_from_records(records: Sequence[FileRecord]) -> str | None:
    """Return the highest MTO xlsx filename revision in a folder.

    Args:
        records: Catalog files of one issued package / tree node.

    Returns:
        Display revision such as ``01-AN02``, or ``None`` when the folder
        has no MTO workbook.
    """

    best_label = ""
    best_rank: tuple[int, int, str] | None = None
    for record in records:
        if str(record.data.get("file_kind") or "") != FileKind.MTO_XLSX.value:
            continue
        label = file_revision_label(record)
        if not label:
            continue
        revision, appendix = parse_sheet_revision(label)
        rank = revision_rank(revision, appendix)
        if best_rank is None or rank > best_rank:
            best_rank = rank
            best_label = label
    return best_label or None


def legalize_approval_f_line(
    *,
    date: str,
    revision: str,
    mto_revision: str | None = None,
    transmittal: str | None = None,
) -> str:
    """Build the prefilled code-A F line for an RD folder.

    TRM is not required for pipeline matching. When omitted, the line uses
    :data:`LEGALIZE_APPROVAL_TOKEN` so the cell still has a human marker.

    Args:
        date: ``DD.MM.YYYY`` (folder mtime, or today).
        revision: Folder filename revision (OD wins ties).
        mto_revision: MTO xlsx revision in that folder, or ``None``.
        transmittal: Optional TRM; empty/None → placeholder token.

    Returns:
        Canonical catalog F line (``auto``, ``MTO <rev|Нет>``).

    Raises:
        ValueError: If ``date`` cannot be formatted.
    """

    token = (transmittal if transmittal is not None else LEGALIZE_APPROVAL_TOKEN)
    token = token.strip() or LEGALIZE_APPROVAL_TOKEN
    return catalog_f_line(
        date=date,
        stage=LEGALIZE_APPROVAL_STAGE,
        revision=revision,
        transmittal=token,
        mto_revision=mto_revision,
        mto_absent=mto_revision is None,
    )


def job_from_edited_comment(
    *,
    title: str,
    mark: str,
    comment_before: str,
    comment_after: str,
    fallback_line: str,
    fallback_revision: str,
    fallback_stage: str = LEGALIZE_APPROVAL_STAGE,
) -> JournalWriteJob:
    """Build a Google write job from an edited F-после cell.

    Args:
        title: Four-digit title.
        mark: Latin AGCC mark.
        comment_before: F до.
        comment_after: Edited F после.
        fallback_line: Prefilled new line.
        fallback_revision: Folder OD revision for column D.
        fallback_stage: Stage when the parsed line has none.

    Returns:
        One :class:`JournalWriteJob` for :func:`execute_journal_writes`.

    Raises:
        ValueError: No unique new line, or the line is not a dated event.
    """

    line = added_history_line(
        comment_before, comment_after, fallback=fallback_line
    )
    parsed = parse_history_line(line)
    if not parsed.date or parsed.stage == "other":
        raise ValueError(
            "Строка F после должна быть датированным событием журнала."
        )
    revision = (
        format_revision(parsed.revision, parsed.appendix) or fallback_revision
    )
    return JournalWriteJob(
        title=title,
        mark=mark,
        f_line=line,
        revision=revision or None,
        stage=parsed.stage or fallback_stage,
    )
