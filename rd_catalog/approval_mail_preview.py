"""Build F-journal preview rows from parsed approval letters.

Qt-free. Sequential mails for the same kit chain: each uses the previous
``comment_after``. Does not write Google.
"""

from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass

from rd_catalog.approval_mail import ApprovalMail, KitLookup
from rd_catalog.f_journal import JournalPatch, build_journal_patch, journal_write_needed
from rd_catalog.google_f_write import JournalWriteJob
from rd_catalog.kits import GoogleKit, IssuanceKit, kit_identity_key


@dataclass(frozen=True, slots=True)
class KitJournalState:
    """Cached KSB ИД F/D/E for one title+mark."""

    comment_raw: str
    d_cell: str = ""
    e_cell: str = ""


CommentLookup = Callable[[str, str], KitJournalState | None]


@dataclass(frozen=True, slots=True)
class MailPreviewRow:
    """One letter as shown on «Письма о согласовании»."""

    mail: ApprovalMail
    comment_before: str
    patch: JournalPatch | None
    error: str
    letter_counts_text: str
    write_status: str = ""
    write_detail: str = ""
    write_f: bool = False
    write_d: bool = False
    write_e: bool = False

    @property
    def already_in_f(self) -> bool:
        return self.patch is not None and not self.write_f

    @property
    def already_complete(self) -> bool:
        return (
            self.patch is not None
            and not self.write_f
            and not self.write_d
            and not self.write_e
        )

    @property
    def already_recorded(self) -> bool:
        """True when the letter is already in F or was written this session."""

        return self.write_status == "written" or self.already_complete

    @property
    def writable(self) -> bool:
        return (
            not self.error
            and self.patch is not None
            and self.write_status != "written"
            and (self.write_f or self.write_d or self.write_e)
        )

    def to_job(self) -> JournalWriteJob | None:
        """Return a Google write job when the row is writable."""

        if not self.writable or self.patch is None:
            return None
        mail = self.mail
        return JournalWriteJob(
            title=mail.title,
            mark=mail.mark,
            f_line=mail.f_line,
            revision=mail.od_revision or None,
            stage=mail.stage,
        )


def mail_dedupe_key(mail: ApprovalMail) -> str:
    """Return a stable key for the same journal letter.

    Args:
        mail: Parsed letter.

    Returns:
        Empty string when the letter cannot be identified.
    """

    if mail.f_line and mail.title and mail.mark:
        return (
            f"f|{mail.title.casefold()}|{mail.mark.casefold()}|{mail.f_line}"
        )
    token = (mail.send_transmittal or "").strip()
    if token and mail.kind:
        return f"trm|{mail.kind}|{token.casefold()}|{mail.date}"
    return ""


def comment_lookup_from_kits(kits: Sequence[GoogleKit]) -> CommentLookup:
    """Map title+mark to cached F/D/E; missing kits return ``None``.

    Args:
        kits: KSB ИД snapshot (SQLite or last Google load).

    Returns:
        Lookup used by :func:`build_mail_previews`.
    """

    mapping = {
        kit_identity_key(kit.title, kit.mark): KitJournalState(
            comment_raw=kit.comment_raw,
            d_cell=kit.sheet_revision_text,
            e_cell=kit.status_sheet,
        )
        for kit in kits
    }

    def lookup(title: str, mark: str) -> KitJournalState | None:
        key = kit_identity_key(title, mark)
        if key not in mapping:
            return None
        return mapping[key]

    return lookup


def kit_lookup_from_issuance(sends: Sequence[IssuanceKit]) -> KitLookup:
    """Resolve a send TRM to ``(title, mark)``; last send wins.

    Args:
        sends: «Выдача РД ПД» rows.

    Returns:
        Lookup for TDO replies without filenames in the top of the thread.
    """

    mapping: dict[str, tuple[str, str]] = {}
    for send in sends:
        token = (send.send_transmittal or "").strip()
        if token:
            mapping[token.casefold()] = (send.title, send.mark)

    def lookup(trm: str) -> tuple[str, str] | None:
        return mapping.get((trm or "").strip().casefold())

    return lookup


def format_letter_counts(counts: tuple[tuple[str, int], ...]) -> str:
    """Format ``(('A', 15), ('B', 5))`` as ``A:15 · B:5``."""

    if not counts:
        return ""
    return " · ".join(f"{letter}:{n}" for letter, n in counts)


def build_mail_previews(
    mails: Sequence[ApprovalMail],
    *,
    comment_lookup: CommentLookup,
) -> tuple[MailPreviewRow, ...]:
    """Turn parsed letters into preview rows with chained F patches.

    Args:
        mails: Parsed letters in ingest order.
        comment_lookup: Current F/D/E by title+mark; ``None`` means the
            kit is absent from КСБ ИД.

    Returns:
        One row per letter. Error rows are not writable.
    """

    comments: dict[tuple[str, str], str] = {}
    de_cells: dict[tuple[str, str], tuple[str, str]] = {}
    rows: list[MailPreviewRow] = []
    for mail in mails:
        counts = format_letter_counts(mail.letter_counts)
        if mail.error or not mail.f_line or not mail.title or not mail.mark:
            rows.append(
                MailPreviewRow(
                    mail=mail,
                    comment_before="",
                    patch=None,
                    error=mail.error or "Письмо не разобрано.",
                    letter_counts_text=counts,
                )
            )
            continue
        key = kit_identity_key(mail.title, mail.mark)
        if key not in comments:
            looked = comment_lookup(mail.title, mail.mark)
            if looked is None:
                rows.append(
                    MailPreviewRow(
                        mail=mail,
                        comment_before="",
                        patch=None,
                        error=(
                            f"Нет комплекта {mail.title}-{mail.mark} в КСБ ИД. "
                            "Новые строки не создаются."
                        ),
                        letter_counts_text=counts,
                    )
                )
                continue
            comments[key] = looked.comment_raw
            de_cells[key] = (looked.d_cell, looked.e_cell)
        before = comments[key]
        live_d, live_e = de_cells[key]
        try:
            patch = build_journal_patch(
                before,
                mail.f_line,
                revision=mail.od_revision or None,
                stage=mail.stage,
            )
        except ValueError as exc:
            rows.append(
                MailPreviewRow(
                    mail=mail,
                    comment_before=before,
                    patch=None,
                    error=str(exc),
                    letter_counts_text=counts,
                )
            )
            continue
        write_f, write_d, write_e = journal_write_needed(
            patch, live_d=live_d, live_e=live_e
        )
        comments[key] = patch.comment_after
        if patch.update_de:
            de_cells[key] = (
                patch.sheet_revision or live_d,
                patch.status_sheet or live_e,
            )
        complete = not write_f and not write_d and not write_e
        rows.append(
            MailPreviewRow(
                mail=mail,
                comment_before=before,
                patch=patch,
                error="",
                letter_counts_text=counts,
                write_status="present" if complete else "",
                write_detail="уже в F" if complete else "",
                write_f=write_f,
                write_d=write_d,
                write_e=write_e,
            )
        )
    return tuple(rows)


def writable_jobs(rows: Sequence[MailPreviewRow]) -> tuple[JournalWriteJob, ...]:
    """Collect Google jobs from writable preview rows."""

    jobs: list[JournalWriteJob] = []
    for row in rows:
        job = row.to_job()
        if job is not None:
            jobs.append(job)
    return tuple(jobs)
