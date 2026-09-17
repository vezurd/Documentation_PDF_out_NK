"""Propose KSB ИД D/E writes from the last column-F event.

Qt-free. Does not call Sheets. Callers pass Комплекты ``KitMatrixRow``
snapshots (SQLite Google F/D/E plus official RD). The write still goes
through :class:`~rd_catalog.google_f_write.JournalWriteJob` so Outlook
letters and this bulk fix share ``execute_journal_writes``.

A row is listed when the last dated classified F event already matches
the official RD filename revision, but live D and/or E would still
change. F is not rewritten (the existing last line is the job payload).
Revision text is always re-parsed from ``event.raw`` (Unicode dashes
normalized) so a line ``на рев. 02-AN02`` is not treated as bare ``02``.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass

from rd_catalog.f_journal import build_journal_patch, journal_write_needed
from rd_catalog.google_f_write import JournalWriteJob
from rd_catalog.kits import (
    KitEvent,
    KitMatrixRow,
    format_revision,
    parse_history_line,
    parse_sheet_revision,
    revisions_equivalent,
)
from rd_catalog.parse import normalize_unicode_dashes

_UNCHANGED = "как есть"


@dataclass(frozen=True, slots=True)
class SheetDeSyncRow:
    """One kit where D/E lag the last F event that already matches RD."""

    title: str
    mark: str
    rd_revision_text: str
    d_now: str
    d_next: str
    e_now: str
    e_next: str
    last_f_line: str
    last_f_revision: str
    last_f_stage: str
    write_d: bool
    write_e: bool
    job: JournalWriteJob
    d_cell: str = ""
    e_cell: str = ""

    def write_label(self) -> str:
        """Return a compact ``D``, ``E``, or ``D+E`` badge."""

        parts: list[str] = []
        if self.write_d:
            parts.append("D")
        if self.write_e:
            parts.append("E")
        return "+".join(parts) if parts else "—"


def list_sheet_de_sync_rows(
    rows: Sequence[KitMatrixRow],
) -> tuple[SheetDeSyncRow, ...]:
    """Return kits whose live D/E still differ from the last F event.

    Args:
        rows: Комплекты matrix rows (caller applies the ban filter).

    Returns:
        Sorted by title then mark. Empty when nothing is auto-writable.
    """

    out: list[SheetDeSyncRow] = []
    for row in rows:
        item = sheet_de_sync_row(row)
        if item is not None:
            out.append(item)
    out.sort(key=lambda item: (item.title, item.mark.casefold()))
    return tuple(out)


def sheet_de_sync_jobs(
    items: Sequence[SheetDeSyncRow],
) -> tuple[JournalWriteJob, ...]:
    """Collect write jobs in list order."""

    return tuple(item.job for item in items)


def sheet_de_sync_row(row: KitMatrixRow) -> SheetDeSyncRow | None:
    """Build one writable D/E sync row, or ``None`` when it is not safe.

    Args:
        row: One kit matrix row.

    Returns:
        A :class:`SheetDeSyncRow` when the last F event matches official
        RD and live D/E still need a write. Otherwise ``None``.
    """

    google = row.google
    if google is None:
        return None
    event = google.last_event
    if event is None or not event.date or event.stage == "other":
        return None
    rd_rev, rd_app = _official_rd_revision(row)
    if not rd_rev:
        return None
    parsed = _parse_last_f_line(event)
    if parsed.stage == "other" or not parsed.date:
        return None
    f_revision = format_revision(parsed.revision, parsed.appendix)
    f_rev, f_app = parsed.revision, parsed.appendix
    if not f_rev:
        f_revision = format_revision(event.revision, event.appendix)
        f_rev, f_app = parse_sheet_revision(f_revision)
    if not revisions_equivalent(f_rev, f_app, rd_rev, rd_app):
        return None
    f_line = (event.raw or "").strip()
    if not f_line:
        return None
    try:
        patch = build_journal_patch(
            google.comment_raw,
            f_line,
            revision=f_revision or None,
            stage=parsed.stage,
        )
    except ValueError:
        return None
    if patch.action != "unchanged":
        return None
    write_f, write_d, write_e = journal_write_needed(
        patch,
        live_d=google.sheet_revision_text,
        live_e=google.status_sheet,
    )
    if write_d and _d_display(google.sheet_revision_text) == _d_display(
        patch.sheet_revision
    ):
        write_d = False
    if write_f or not (write_d or write_e):
        return None
    if write_d:
        proposed_rev, proposed_app = parse_sheet_revision(patch.sheet_revision)
        if not revisions_equivalent(proposed_rev, proposed_app, rd_rev, rd_app):
            return None
        if not revisions_equivalent(proposed_rev, proposed_app, f_rev, f_app):
            return None
    job = JournalWriteJob(
        title=row.title,
        mark=row.mark,
        f_line=f_line,
        revision=f_revision or None,
        stage=parsed.stage,
    )
    d_now = _d_display(google.sheet_revision_text)
    e_now = (google.status_sheet or "").strip() or "—"
    return SheetDeSyncRow(
        title=row.title,
        mark=row.mark,
        rd_revision_text=format_revision(rd_rev, rd_app) or row.rd.revision_text,
        d_now=d_now or "—",
        d_next=_d_display(patch.sheet_revision) if write_d else _UNCHANGED,
        e_now=e_now,
        e_next=(patch.status_sheet.strip() or "—") if write_e else _UNCHANGED,
        last_f_line=f_line,
        last_f_revision=f_revision,
        last_f_stage=parsed.stage_label or parsed.stage,
        write_d=write_d,
        write_e=write_e,
        job=job,
        d_cell=patch.sheet_revision if write_d else "",
        e_cell=patch.status_sheet if write_e else "",
    )


def _parse_last_f_line(event: KitEvent) -> KitEvent:
    raw = normalize_unicode_dashes((event.raw or "").strip())
    if not raw:
        return event
    return parse_history_line(raw)


def _d_display(text: str) -> str:
    revision, appendix = parse_sheet_revision(text)
    return format_revision(revision, appendix)


def _official_rd_revision(row: KitMatrixRow) -> tuple[str | None, str | None]:
    if not row.rd.present:
        return None, None
    if row.rd.revision:
        return row.rd.revision, row.rd.appendix
    return parse_sheet_revision(row.rd.revision_text)
