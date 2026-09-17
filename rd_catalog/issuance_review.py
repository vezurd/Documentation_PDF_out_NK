"""Qt-free issuance journal: fingerprints, rematch, and effective sends.

User decisions live in ``issuance_review`` and survive Google snapshot
replace. This module does not import Qt and does not write UNC paths.
"""

from __future__ import annotations

import hashlib
from collections.abc import Sequence
from dataclasses import dataclass

from rd_catalog.db import CatalogDatabase, IssuanceReviewRow
from rd_catalog.kits import (
    IssuanceKit,
    format_event_date_sortable,
    format_revision,
    kit_identity_key,
    latest_issuance_from_sends,
    parse_sheet_revision,
)
from rd_catalog.models import FileRecord, SourceKind
from rd_catalog.parse import normalize_unicode_dashes
from rd_catalog.perf_log import perf_span

_INCLUDE_DECISIONS = frozenset({"active", "legalized"})
_EXCLUDE_DECISIONS_NEED_COMMENT = frozenset(
    {"annulled", "erroneous", "duplicate"}
)


@dataclass(frozen=True, slots=True)
class JournalAutoMtoHit:
    """Auto-MTO presence used by the issuance journal (no customer_pi import)."""

    title: str
    mark: str
    revision_text: str
    relpath: str = ""


@dataclass(frozen=True, slots=True)
class IssuanceJournalRow:
    """One visible issuance-journal line (sheet, review, or orphan candidate)."""

    title: str
    mark: str
    kind: str
    source: str
    revision_text: str
    send_date: str
    send_date_sortable: str
    send_transmittal: str
    incoming_control_date: str
    incoming_control_date_sortable: str
    confirm_transmittal: str
    sheet_status: str
    note: str
    decision: str
    comment: str | None
    match_state: str
    identity_fingerprint: str
    evidence_fingerprint: str
    source_path: str
    path_key: str
    in_f: bool
    in_rd: bool
    in_robot: bool
    in_auto_mto: bool
    review_id: int | None
    issuance_send_id: int | None


def _canon(value: str | None) -> str:
    return normalize_unicode_dashes(value or "").strip().casefold()


def _canon_revision(text: str | None) -> str:
    raw = text or ""
    revision, appendix = parse_sheet_revision(raw)
    if revision is not None:
        return format_revision(revision, appendix)
    return _canon(raw)


def _sha256_payload(parts: Sequence[str]) -> str:
    payload = "|".join(parts)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _send_date_sortable(send: IssuanceKit) -> str:
    return send.send_date_sortable or format_event_date_sortable(send.send_date)


def _incoming_sortable(send: IssuanceKit) -> str:
    return send.incoming_control_date_sortable or format_event_date_sortable(
        send.incoming_control_date
    )


def _send_identity_parts(send: IssuanceKit) -> tuple[str, ...]:
    return (
        _canon(send.title),
        _canon(send.mark),
        _canon_revision(send.revision_text),
        _send_date_sortable(send),
        _canon(send.send_transmittal),
    )


def send_identity_fingerprint(send: IssuanceKit) -> str:
    """Return the identity fingerprint for one sheet send.

    Args:
        send: Parsed «Выдача РД ПД» row.

    Returns:
        SHA-256 hex of the canonical identity payload.
    """

    return _sha256_payload(_send_identity_parts(send))


def journal_row_matches_issuance(
    row: IssuanceJournalRow,
    issuance: IssuanceKit,
) -> bool:
    """Return whether ``row`` is the journal line behind ``issuance``.

    Комплекты «Выдача · рев.» is ``latest_effective_issuance_kits``. Sheet
    sends match by identity fingerprint; legalized orphans match the same
    canonical title/mark/rev/date/TRM parts (orphan fingerprints differ).

    Args:
        row: Visible issuance-journal line.
        issuance: Effective send painted on Комплекты.

    Returns:
        ``True`` when ``row`` is that effective send.
    """

    if kit_identity_key(row.title, row.mark) != kit_identity_key(
        issuance.title, issuance.mark
    ):
        return False
    fingerprint = send_identity_fingerprint(issuance)
    if row.identity_fingerprint and row.identity_fingerprint == fingerprint:
        return True
    row_date = row.send_date_sortable or format_event_date_sortable(row.send_date)
    return (
        _canon_revision(row.revision_text) == _canon_revision(issuance.revision_text)
        and row_date == _send_date_sortable(issuance)
        and _canon(row.send_transmittal) == _canon(issuance.send_transmittal)
    )


def pick_journal_row_for_issuance(
    rows: Sequence[IssuanceJournalRow],
    *,
    title: str,
    mark: str,
    issuance: IssuanceKit | None = None,
) -> IssuanceJournalRow | None:
    """Pick the journal line that paints Комплекты «Выдача · рев.».

    Args:
        rows: Snapshot from ``list_issuance_journal``.
        title: Four-digit title.
        mark: Latin AGCC mark.
        issuance: Effective send for the kit, if any.

    Returns:
        Matching journal row, else the first row of the kit, else ``None``.
    """

    key = kit_identity_key(title, mark)
    kit_rows = [
        row for row in rows if kit_identity_key(row.title, row.mark) == key
    ]
    if not kit_rows:
        return None
    if issuance is not None:
        matches = [
            row for row in kit_rows if journal_row_matches_issuance(row, issuance)
        ]
        if matches:
            return _prefer_sheet_journal_row(matches)
        rev = _canon_revision(issuance.revision_text)
        if rev:
            rev_hits = [
                row for row in kit_rows if _canon_revision(row.revision_text) == rev
            ]
            included = [
                row
                for row in rev_hits
                if row.decision not in _EXCLUDE_DECISIONS_NEED_COMMENT
                and (
                    row.kind == "send"
                    or row.source == "issuance"
                    or row.decision in _INCLUDE_DECISIONS
                )
            ]
            pool = included or rev_hits
            if pool:
                return _prefer_sheet_journal_row(pool)
    return kit_rows[0]


def _prefer_sheet_journal_row(
    rows: Sequence[IssuanceJournalRow],
) -> IssuanceJournalRow:
    return min(
        rows,
        key=lambda row: (
            0 if (row.kind == "send" or row.issuance_send_id is not None) else 1,
            -(row.issuance_send_id or 0),
        ),
    )


def send_evidence_fingerprint(send: IssuanceKit) -> str:
    """Return the evidence fingerprint for one sheet send.

    Args:
        send: Parsed «Выдача РД ПД» row.

    Returns:
        SHA-256 hex of identity plus sheet evidence fields.
    """

    return _sha256_payload(
        (
            *_send_identity_parts(send),
            _canon(send.status),
            _incoming_sortable(send),
            _canon(send.confirm_transmittal),
            _canon(send.note_raw),
        )
    )


def orphan_identity_fingerprint(
    source: str,
    title: str,
    mark: str,
    revision_text: str,
    path_key: str = "",
) -> str:
    """Return the identity fingerprint for an orphan / manual candidate.

    Args:
        source: Origin token (``google_f``, ``rd``, ``robot``, ``auto_mto``,
            ``manual``).
        title: Four-digit title.
        mark: Latin AGCC mark.
        revision_text: Revision display text.
        path_key: Optional path identity for robot / Auto MTO.

    Returns:
        SHA-256 hex of the canonical orphan payload.
    """

    return _sha256_payload(
        (
            _canon(source),
            _canon(title),
            _canon(mark),
            _canon_revision(revision_text),
            _canon(path_key),
        )
    )


def decision_includes(decision: str) -> bool:
    """Return whether ``decision`` keeps a send in the effective contour.

    Args:
        decision: Stored issuance decision, possibly empty.

    Returns:
        ``True`` for ``active`` and ``legalized``.
    """

    return decision in _INCLUDE_DECISIONS


def _issuance_kit_from_fields(
    *,
    title: str,
    mark: str,
    revision_text: str,
    send_date: str,
    send_transmittal: str,
    incoming_control_date: str,
    confirm_transmittal: str,
    sheet_status: str,
    note: str,
    send_date_sortable: str = "",
    incoming_control_date_sortable: str = "",
) -> IssuanceKit:
    """Build a fingerprint-ready send from journal / review fields."""

    revision, appendix = parse_sheet_revision(revision_text)
    return IssuanceKit(
        title=title,
        mark=mark,
        mark_raw=mark,
        title_system=f"{title}-{mark}",
        revision=revision,
        appendix=appendix,
        revision_text=revision_text,
        status=sheet_status,
        send_date=send_date,
        send_date_sortable=send_date_sortable or format_event_date_sortable(send_date),
        send_transmittal=send_transmittal,
        incoming_control_date=incoming_control_date,
        incoming_control_date_sortable=incoming_control_date_sortable
        or format_event_date_sortable(incoming_control_date),
        confirm_transmittal=confirm_transmittal,
        note_raw=note,
        row_index=0,
    )


def _is_sheet_journal_row(row: IssuanceJournalRow) -> bool:
    return row.kind == "send" or row.source == "issuance" or row.issuance_send_id is not None


def apply_journal_decision(
    database: CatalogDatabase,
    row: IssuanceJournalRow,
    *,
    decision: str,
    comment: str | None = None,
    revision_text: str | None = None,
    send_date: str | None = None,
    send_transmittal: str | None = None,
    incoming_control_date: str | None = None,
    confirm_transmittal: str | None = None,
    sheet_status: str | None = None,
    note: str | None = None,
) -> int:
    """Persist a journal decision, optionally overlaying field edits.

    Args:
        database: Catalog with ``issuance_review``.
        row: Visible journal line being edited.
        decision: Decision token to store (may be empty for a draft).
        comment: Reviewer comment; ``None`` keeps the current value.
        revision_text: Optional revision override.
        send_date: Optional send-date override.
        send_transmittal: Optional send TRM override.
        incoming_control_date: Optional incoming-control date override.
        confirm_transmittal: Optional confirmation TRM override.
        sheet_status: Optional sheet-status override.
        note: Optional note override.

    Returns:
        ``issuance_review.id``.

    Raises:
        ValueError: Invalid decision, legalize without revision, or
            exclude without a comment.
    """

    merged_revision = row.revision_text if revision_text is None else revision_text
    merged_comment = row.comment if comment is None else comment
    merged_send_date = row.send_date if send_date is None else send_date
    merged_send_trm = (
        row.send_transmittal if send_transmittal is None else send_transmittal
    )
    merged_incoming = (
        row.incoming_control_date
        if incoming_control_date is None
        else incoming_control_date
    )
    merged_confirm = (
        row.confirm_transmittal if confirm_transmittal is None else confirm_transmittal
    )
    merged_status = row.sheet_status if sheet_status is None else sheet_status
    merged_note = row.note if note is None else note
    send_date_sortable = (
        row.send_date_sortable
        if send_date is None
        else format_event_date_sortable(merged_send_date)
    )
    incoming_sortable = (
        row.incoming_control_date_sortable
        if incoming_control_date is None
        else format_event_date_sortable(merged_incoming)
    )
    validate_issuance_decision(
        decision=decision,
        revision_text=merged_revision,
        comment=merged_comment,
    )
    kind = row.kind
    source = row.source
    identity = row.identity_fingerprint
    evidence = row.evidence_fingerprint
    match_state = row.match_state
    has_review = row.review_id is not None
    is_sheet = _is_sheet_journal_row(row)
    kit = _issuance_kit_from_fields(
        title=row.title,
        mark=row.mark,
        revision_text=merged_revision,
        send_date=merged_send_date,
        send_transmittal=merged_send_trm,
        incoming_control_date=merged_incoming,
        confirm_transmittal=merged_confirm,
        sheet_status=merged_status,
        note=merged_note,
        send_date_sortable=send_date_sortable,
        incoming_control_date_sortable=incoming_sortable,
    )
    if not has_review and is_sheet:
        kind = "send"
        source = "issuance"
        identity = send_identity_fingerprint(kit)
        evidence = send_evidence_fingerprint(kit)
        match_state = "matched"
    elif not identity and (kind in {"orphan", "manual"} or source in {
        "google_f",
        "rd",
        "robot",
        "auto_mto",
        "manual",
    }):
        identity = orphan_identity_fingerprint(
            source or kind or "manual",
            row.title,
            row.mark,
            merged_revision,
            row.path_key,
        )
        if kind == "manual" and not has_review:
            match_state = ""
    elif is_sheet and (
        row.issuance_send_id is not None or row.match_state == "matched"
    ):
        match_state = "matched"
    if kind == "manual" and not has_review:
        match_state = ""
    return database.upsert_issuance_review(
        row.title,
        row.mark,
        kind,
        source,
        identity,
        decision=decision,
        revision_text=merged_revision,
        send_date=merged_send_date,
        send_date_sortable=send_date_sortable,
        send_transmittal=merged_send_trm,
        incoming_control_date=merged_incoming,
        incoming_control_date_sortable=incoming_sortable,
        confirm_transmittal=merged_confirm,
        sheet_status=merged_status,
        note=merged_note,
        comment=merged_comment,
        evidence_fingerprint=evidence,
        match_state=match_state,
        source_path=row.source_path,
        path_key=row.path_key,
        review_id=row.review_id,
    )


def add_manual_journal_row(
    database: CatalogDatabase,
    title: str,
    mark: str,
    *,
    revision_text: str = "",
    send_date: str = "",
    send_transmittal: str = "",
    incoming_control_date: str = "",
    confirm_transmittal: str = "",
    sheet_status: str = "",
    note: str = "",
    decision: str = "",
    comment: str | None = None,
) -> int:
    """Insert a manual journal review (draft or decided).

    Args:
        database: Catalog with ``issuance_review``.
        title: Four-digit title.
        mark: Latin AGCC mark.
        revision_text: Optional revision text.
        send_date: Optional send date.
        send_transmittal: Optional send TRM.
        incoming_control_date: Optional incoming-control date.
        confirm_transmittal: Optional confirmation TRM.
        sheet_status: Optional sheet status.
        note: Optional note.
        decision: Decision token; empty is a draft.
        comment: Reviewer comment; required for exclude decisions.

    Returns:
        ``issuance_review.id``.

    Raises:
        ValueError: Legalize without revision, or exclude without comment.
    """

    if decision in {"legalized"} | _EXCLUDE_DECISIONS_NEED_COMMENT:
        validate_issuance_decision(
            decision=decision,
            revision_text=revision_text,
            comment=comment,
        )
    identity = orphan_identity_fingerprint(
        "manual", title, mark, revision_text or "", ""
    )
    return database.upsert_issuance_review(
        title,
        mark,
        "manual",
        "manual",
        identity,
        decision=decision,
        revision_text=revision_text,
        send_date=send_date,
        send_date_sortable=format_event_date_sortable(send_date),
        send_transmittal=send_transmittal,
        incoming_control_date=incoming_control_date,
        incoming_control_date_sortable=format_event_date_sortable(
            incoming_control_date
        ),
        confirm_transmittal=confirm_transmittal,
        sheet_status=sheet_status,
        note=note,
        comment=comment,
        match_state="",
    )


def validate_issuance_decision(
    *,
    decision: str,
    revision_text: str,
    comment: str | None,
) -> None:
    """Validate a user issuance decision before persist.

    Args:
        decision: Chosen decision token.
        revision_text: Revision text supplied with the decision.
        comment: Reviewer comment, required for exclude decisions.

    Raises:
        ValueError: Legalize without revision, or exclude without comment.
    """

    if decision == "legalized" and not (revision_text or "").strip():
        raise ValueError("legalized issuance review requires revision_text")
    if decision in _EXCLUDE_DECISIONS_NEED_COMMENT and not (comment or "").strip():
        raise ValueError(f"{decision} issuance review requires a comment")


def issuance_kit_from_review(row: IssuanceReviewRow) -> IssuanceKit:
    """Build a synthetic send from a legalized orphan or manual review.

    Args:
        row: Persisted issuance review.

    Returns:
        ``IssuanceKit`` with ``row_index = -review.id`` (or ``0``).
    """

    revision, appendix = parse_sheet_revision(row.revision_text)
    row_index = -int(row.id) if row.id is not None else 0
    return IssuanceKit(
        title=row.title,
        mark=row.mark,
        mark_raw=row.mark,
        title_system=f"{row.title}-{row.mark}",
        revision=revision,
        appendix=appendix,
        revision_text=row.revision_text,
        status=row.sheet_status,
        send_date=row.send_date,
        send_date_sortable=row.send_date_sortable
        or format_event_date_sortable(row.send_date),
        send_transmittal=row.send_transmittal,
        incoming_control_date=row.incoming_control_date,
        incoming_control_date_sortable=row.incoming_control_date_sortable
        or format_event_date_sortable(row.incoming_control_date),
        confirm_transmittal=row.confirm_transmittal,
        note_raw=row.note,
        row_index=row_index,
    )


def _write_review(
    database: CatalogDatabase,
    row: IssuanceReviewRow,
    **overrides: object,
) -> int:
    payload = {
        "title": row.title,
        "mark": row.mark,
        "kind": row.kind,
        "source": row.source,
        "identity_fingerprint": row.identity_fingerprint,
        "decision": row.decision,
        "revision_text": row.revision_text,
        "send_date": row.send_date,
        "send_date_sortable": row.send_date_sortable,
        "send_transmittal": row.send_transmittal,
        "incoming_control_date": row.incoming_control_date,
        "incoming_control_date_sortable": row.incoming_control_date_sortable,
        "confirm_transmittal": row.confirm_transmittal,
        "sheet_status": row.sheet_status,
        "note": row.note,
        "comment": row.comment,
        "evidence_fingerprint": row.evidence_fingerprint,
        "match_state": row.match_state,
        "source_path": row.source_path,
        "path_key": row.path_key,
        "decided_at": row.decided_at,
        "review_id": row.id,
    }
    payload.update(overrides)
    return database.upsert_issuance_review(
        str(payload["title"]),
        str(payload["mark"]),
        str(payload["kind"]),
        str(payload["source"]),
        str(payload["identity_fingerprint"]),
        decision=str(payload["decision"]),
        revision_text=str(payload["revision_text"]),
        send_date=str(payload["send_date"]),
        send_date_sortable=str(payload["send_date_sortable"]),
        send_transmittal=str(payload["send_transmittal"]),
        incoming_control_date=str(payload["incoming_control_date"]),
        incoming_control_date_sortable=str(
            payload["incoming_control_date_sortable"]
        ),
        confirm_transmittal=str(payload["confirm_transmittal"]),
        sheet_status=str(payload["sheet_status"]),
        note=str(payload["note"]),
        comment=payload["comment"] if payload["comment"] is None else str(payload["comment"]),
        evidence_fingerprint=str(payload["evidence_fingerprint"]),
        match_state=str(payload["match_state"]),
        source_path=str(payload["source_path"]),
        path_key=str(payload["path_key"]),
        decided_at=str(payload["decided_at"] or ""),
        review_id=int(payload["review_id"]) if payload["review_id"] is not None else None,
    )


def _apply_send_match(database: CatalogDatabase, row: IssuanceReviewRow, send: IssuanceKit) -> None:
    _write_review(
        database,
        row,
        title=send.title,
        mark=send.mark,
        revision_text=send.revision_text,
        send_date=send.send_date,
        send_date_sortable=_send_date_sortable(send),
        send_transmittal=send.send_transmittal,
        incoming_control_date=send.incoming_control_date,
        incoming_control_date_sortable=_incoming_sortable(send),
        confirm_transmittal=send.confirm_transmittal,
        sheet_status=send.status,
        note=send.note_raw,
        identity_fingerprint=send_identity_fingerprint(send),
        evidence_fingerprint=send_evidence_fingerprint(send),
        match_state="matched",
    )


def rematch_issuance_reviews(database: CatalogDatabase) -> None:
    """Re-attach ``kind=send`` reviews to the current issuance snapshot.

    Exact identity wins. Otherwise a unique ``(revision, date)`` or
    ``(revision, TRM)`` hit on the same kit rematches. Zero candidates mark
    ``unmatched``; several candidates on one ladder step mark ``ambiguous``.
    Decision and comment are never cleared.

    Args:
        database: Catalog with a current Google snapshot already stored.
    """

    sends_with_ids = database.list_issuance_sends_with_ids()
    send_by_fp: dict[str, IssuanceKit] = {}
    sends_by_kit: dict[tuple[str, str], list[IssuanceKit]] = {}
    for _send_id, send in sends_with_ids:
        fingerprint = send_identity_fingerprint(send)
        send_by_fp[fingerprint] = send
        sends_by_kit.setdefault(kit_identity_key(send.title, send.mark), []).append(
            send
        )

    claimed: set[str] = set()
    reviews = [
        row for row in database.list_issuance_reviews() if row.kind == "send"
    ]
    pending: list[IssuanceReviewRow] = []
    for row in reviews:
        send = send_by_fp.get(row.identity_fingerprint)
        if send is not None:
            _apply_send_match(database, row, send)
            claimed.add(send_identity_fingerprint(send))
            continue
        pending.append(row)

    for row in pending:
        kit_sends = [
            send
            for send in sends_by_kit.get(kit_identity_key(row.title, row.mark), ())
            if send_identity_fingerprint(send) not in claimed
        ]
        rev = _canon_revision(row.revision_text)
        date_key = row.send_date_sortable or format_event_date_sortable(row.send_date)
        date_hits = [
            send
            for send in kit_sends
            if _canon_revision(send.revision_text) == rev
            and _send_date_sortable(send) == date_key
        ]
        if len(date_hits) == 1:
            send = date_hits[0]
            _apply_send_match(database, row, send)
            claimed.add(send_identity_fingerprint(send))
            continue
        if len(date_hits) > 1:
            _write_review(database, row, match_state="ambiguous")
            continue

        review_trm = _canon(row.send_transmittal)
        trm_hits: list[IssuanceKit] = []
        if review_trm:
            trm_hits = [
                send
                for send in kit_sends
                if _canon_revision(send.revision_text) == rev
                and _canon(send.send_transmittal) == review_trm
                and _canon(send.send_transmittal)
            ]
        if len(trm_hits) == 1:
            send = trm_hits[0]
            _apply_send_match(database, row, send)
            claimed.add(send_identity_fingerprint(send))
            continue
        if len(trm_hits) > 1:
            _write_review(database, row, match_state="ambiguous")
            continue
        _write_review(database, row, match_state="unmatched")


def effective_issuance_sends(
    database: CatalogDatabase,
    title: str | None = None,
    mark: str | None = None,
) -> list[tuple[int | None, IssuanceKit]]:
    """Return include-contour sends plus legalized orphan/manual kits.

    Sheet sends with no review are implicit ``active``. A ``kind=send``
    review whose current identity is present excludes that send when
    :func:`decision_includes` is false, including unmatched/ambiguous
    exclude decisions.

    Args:
        database: Initialized catalog database.
        title: Optional title filter (case-insensitive).
        mark: Optional mark filter (case-insensitive).

    Returns:
        ``(issuance_send.id or None, send)`` pairs sorted like the sheet
        snapshot (``send_date_sortable``, ``row_index``, then id).
    """

    reviews = database.list_issuance_reviews(title, mark)
    send_reviews = {
        row.identity_fingerprint: row
        for row in reviews
        if row.kind == "send"
    }
    result: list[tuple[int | None, IssuanceKit]] = []
    for send_id, send in database.list_issuance_sends_with_ids(title, mark):
        review = send_reviews.get(send_identity_fingerprint(send))
        if review is not None and not decision_includes(review.decision):
            continue
        result.append((send_id, send))
    for row in reviews:
        if row.kind not in {"orphan", "manual"}:
            continue
        if row.decision != "legalized":
            continue
        if not (row.revision_text or "").strip():
            continue
        result.append((None, issuance_kit_from_review(row)))
    result.sort(
        key=lambda item: (
            item[1].send_date_sortable,
            item[1].row_index,
            item[0] if item[0] is not None else 0,
        )
    )
    return result


def latest_effective_issuance_kits(
    database: CatalogDatabase,
    title: str | None = None,
    mark: str | None = None,
) -> tuple[IssuanceKit, ...]:
    """Return latest-wins effective sends per kit.

    Args:
        database: Initialized catalog database.
        title: Optional title filter (case-insensitive).
        mark: Optional mark filter (case-insensitive).

    Returns:
        One effective send per ``(title, mark)``, sorted by title and mark.
    """

    with perf_span("issuance.latest_effective_kits"):
        sends = [
            send
            for _send_id, send in effective_issuance_sends(database, title, mark)
        ]
        return latest_issuance_from_sends(sends)


def _revision_texts_equivalent(left: str, right: str) -> bool:
    from rd_catalog.pipeline import revision_texts_equivalent

    return revision_texts_equivalent(left, right)


def _record_revision_text(record: FileRecord) -> str:
    return format_revision(record.data.get("revision"), record.data.get("appendix"))


def _record_title_mark(record: FileRecord) -> tuple[str, str]:
    return str(record.data.get("title") or ""), str(record.data.get("mark") or "")


def _has_equivalent_revision(
    revision: str,
    items: Sequence[tuple[tuple[str, str], str]],
    kit: tuple[str, str],
) -> bool:
    return any(
        key == kit and _revision_texts_equivalent(revision, other)
        for key, other in items
    )


_ORPHAN_SOURCE_RANK = {
    "rd": 0,
    "google_f": 1,
    "robot": 2,
    "auto_mto": 3,
}


@dataclass(frozen=True, slots=True)
class _MergedOrphan:
    """One unseen revision collapsed across F / RD / robot / Auto MTO."""

    title: str
    mark: str
    revision_text: str
    sources: frozenset[str]
    path_key: str
    source_path: str
    path_source: str


def _orphan_source_rank(source: str) -> int:
    return _ORPHAN_SOURCE_RANK.get(source, 99)


def _preferred_orphan_source(sources: frozenset[str] | set[str]) -> str:
    return min(sources, key=_orphan_source_rank)


def _flags_for(
    kit: tuple[str, str],
    revision: str,
    f_revs: Sequence[tuple[tuple[str, str], str]],
    rd_revs: Sequence[tuple[tuple[str, str], str]],
    robot_revs: Sequence[tuple[tuple[str, str], str]],
    auto_revs: Sequence[tuple[tuple[str, str], str]],
) -> tuple[bool, bool, bool, bool]:
    return (
        _has_equivalent_revision(revision, f_revs, kit),
        _has_equivalent_revision(revision, rd_revs, kit),
        _has_equivalent_revision(revision, robot_revs, kit),
        _has_equivalent_revision(revision, auto_revs, kit),
    )


def _journal_row(
    *,
    title: str,
    mark: str,
    kind: str,
    source: str,
    revision_text: str,
    send_date: str = "",
    send_date_sortable: str = "",
    send_transmittal: str = "",
    incoming_control_date: str = "",
    incoming_control_date_sortable: str = "",
    confirm_transmittal: str = "",
    sheet_status: str = "",
    note: str = "",
    decision: str = "",
    comment: str | None = None,
    match_state: str = "",
    identity_fingerprint: str = "",
    evidence_fingerprint: str = "",
    source_path: str = "",
    path_key: str = "",
    in_f: bool = False,
    in_rd: bool = False,
    in_robot: bool = False,
    in_auto_mto: bool = False,
    review_id: int | None = None,
    issuance_send_id: int | None = None,
) -> IssuanceJournalRow:
    return IssuanceJournalRow(
        title=title,
        mark=mark,
        kind=kind,
        source=source,
        revision_text=revision_text,
        send_date=send_date,
        send_date_sortable=send_date_sortable,
        send_transmittal=send_transmittal,
        incoming_control_date=incoming_control_date,
        incoming_control_date_sortable=incoming_control_date_sortable,
        confirm_transmittal=confirm_transmittal,
        sheet_status=sheet_status,
        note=note,
        decision=decision,
        comment=comment,
        match_state=match_state,
        identity_fingerprint=identity_fingerprint,
        evidence_fingerprint=evidence_fingerprint,
        source_path=source_path,
        path_key=path_key,
        in_f=in_f,
        in_rd=in_rd,
        in_robot=in_robot,
        in_auto_mto=in_auto_mto,
        review_id=review_id,
        issuance_send_id=issuance_send_id,
    )


def _overlay_send_row(
    send_id: int,
    send: IssuanceKit,
    review: IssuanceReviewRow | None,
    flags: tuple[bool, bool, bool, bool],
) -> IssuanceJournalRow:
    if review is None:
        decision = "active"
        comment = None
        match_state = ""
        review_id = None
        source = "issuance"
        kind = "send"
        identity = send_identity_fingerprint(send)
        evidence = send_evidence_fingerprint(send)
        source_path = ""
        path_key = ""
    else:
        decision = review.decision
        comment = review.comment
        match_state = review.match_state
        review_id = review.id
        source = review.source
        kind = review.kind
        identity = review.identity_fingerprint
        evidence = review.evidence_fingerprint
        source_path = review.source_path
        path_key = review.path_key
    return _journal_row(
        title=send.title,
        mark=send.mark,
        kind=kind,
        source=source,
        revision_text=send.revision_text,
        send_date=send.send_date,
        send_date_sortable=_send_date_sortable(send),
        send_transmittal=send.send_transmittal,
        incoming_control_date=send.incoming_control_date,
        incoming_control_date_sortable=_incoming_sortable(send),
        confirm_transmittal=send.confirm_transmittal,
        sheet_status=send.status,
        note=send.note_raw,
        decision=decision,
        comment=comment,
        match_state=match_state,
        identity_fingerprint=identity,
        evidence_fingerprint=evidence,
        source_path=source_path,
        path_key=path_key,
        in_f=flags[0],
        in_rd=flags[1],
        in_robot=flags[2],
        in_auto_mto=flags[3],
        review_id=review_id,
        issuance_send_id=send_id,
    )


def _review_journal_row(
    row: IssuanceReviewRow,
    flags: tuple[bool, bool, bool, bool],
    *,
    issuance_send_id: int | None,
) -> IssuanceJournalRow:
    return _journal_row(
        title=row.title,
        mark=row.mark,
        kind=row.kind,
        source=row.source,
        revision_text=row.revision_text,
        send_date=row.send_date,
        send_date_sortable=row.send_date_sortable,
        send_transmittal=row.send_transmittal,
        incoming_control_date=row.incoming_control_date,
        incoming_control_date_sortable=row.incoming_control_date_sortable,
        confirm_transmittal=row.confirm_transmittal,
        sheet_status=row.sheet_status,
        note=row.note,
        decision=row.decision,
        comment=row.comment,
        match_state=row.match_state,
        identity_fingerprint=row.identity_fingerprint,
        evidence_fingerprint=row.evidence_fingerprint,
        source_path=row.source_path,
        path_key=row.path_key,
        in_f=flags[0],
        in_rd=flags[1],
        in_robot=flags[2],
        in_auto_mto=flags[3],
        review_id=row.id,
        issuance_send_id=issuance_send_id,
    )


def list_issuance_journal(
    database: CatalogDatabase,
    *,
    records: Sequence[FileRecord] = (),
    auto_mto_hits: Sequence[JournalAutoMtoHit] = (),
) -> tuple[IssuanceJournalRow, ...]:
    """Build the issuance journal: sheet rows, reviews, and orphan candidates.

    Does not persist orphan candidates and does not invent heatmap cells.

    Args:
        database: Initialized catalog database.
        records: Current scan records (RD / ROBOT presence).
        auto_mto_hits: Auto-MTO revisions supplied by the caller.

    Returns:
        Journal rows (sheet sends, detached send reviews, orphan/manual
        reviews, then one unseen orphan per kit+revision). Several PDFs
        or MTO files of the same revision collapse into one row; F / RD /
        robot / Auto MTO presence is shown as flags.
    """

    with perf_span("journal.list_issuance_journal"):
        sends_with_ids = database.list_issuance_sends_with_ids()
        reviews = database.list_issuance_reviews()
        send_by_fp = {
            send_identity_fingerprint(send): (send_id, send)
            for send_id, send in sends_with_ids
        }
        review_by_send_fp = {
            row.identity_fingerprint: row for row in reviews if row.kind == "send"
        }

        f_revs: list[tuple[tuple[str, str], str]] = []
        f_candidates: list[tuple[str, str, str, str, str]] = []
        for kit in database.list_google_kits():
            kit_key = kit_identity_key(kit.title, kit.mark)
            for event in kit.events:
                revision = format_revision(event.revision, event.appendix)
                if not revision:
                    continue
                f_revs.append((kit_key, revision))
                f_candidates.append((kit.title, kit.mark, revision, "google_f", ""))

        rd_revs: list[tuple[tuple[str, str], str]] = []
        robot_revs: list[tuple[tuple[str, str], str]] = []
        file_candidates: list[tuple[str, str, str, str, str, str]] = []
        for record in records:
            if not record.present:
                continue
            if record.source not in {SourceKind.RD, SourceKind.ROBOT}:
                continue
            title, mark = _record_title_mark(record)
            if not title or not mark:
                continue
            revision = _record_revision_text(record)
            if not revision:
                continue
            kit_key = kit_identity_key(title, mark)
            source = "rd" if record.source == SourceKind.RD else "robot"
            if source == "rd":
                rd_revs.append((kit_key, revision))
            else:
                robot_revs.append((kit_key, revision))
            file_candidates.append(
                (title, mark, revision, source, record.path_key, record.path)
            )

        auto_revs: list[tuple[tuple[str, str], str]] = []
        auto_candidates: list[tuple[str, str, str, str, str]] = []
        for hit in auto_mto_hits:
            if not hit.revision_text.strip():
                continue
            kit_key = kit_identity_key(hit.title, hit.mark)
            auto_revs.append((kit_key, hit.revision_text))
            auto_candidates.append(
                (hit.title, hit.mark, hit.revision_text, "auto_mto", hit.relpath)
            )

        represented: list[tuple[tuple[str, str], str]] = []
        for _send_id, send in sends_with_ids:
            represented.append(
                (kit_identity_key(send.title, send.mark), send.revision_text)
            )
        for row in reviews:
            if row.revision_text.strip():
                represented.append(
                    (kit_identity_key(row.title, row.mark), row.revision_text)
                )

        rows: list[IssuanceJournalRow] = []
        attached_review_ids: set[int] = set()
        for send_id, send in sends_with_ids:
            fingerprint = send_identity_fingerprint(send)
            review = review_by_send_fp.get(fingerprint)
            if review is not None and review.id is not None:
                attached_review_ids.add(review.id)
            flags = _flags_for(
                kit_identity_key(send.title, send.mark),
                send.revision_text,
                f_revs,
                rd_revs,
                robot_revs,
                auto_revs,
            )
            rows.append(_overlay_send_row(send_id, send, review, flags))

        for row in reviews:
            if row.kind != "send":
                continue
            if row.id in attached_review_ids:
                continue
            if row.match_state not in {"unmatched", "ambiguous"}:
                continue
            if row.identity_fingerprint in send_by_fp:
                continue
            flags = _flags_for(
                kit_identity_key(row.title, row.mark),
                row.revision_text,
                f_revs,
                rd_revs,
                robot_revs,
                auto_revs,
            )
            rows.append(_review_journal_row(row, flags, issuance_send_id=None))

        for row in reviews:
            if row.kind not in {"orphan", "manual"}:
                continue
            flags = _flags_for(
                kit_identity_key(row.title, row.mark),
                row.revision_text,
                f_revs,
                rd_revs,
                robot_revs,
                auto_revs,
            )
            rows.append(_review_journal_row(row, flags, issuance_send_id=None))

        merged_orphans: dict[tuple[tuple[str, str], str], _MergedOrphan] = {}
        orphan_specs: list[tuple[str, str, str, str, str, str]] = []
        for title, mark, revision, source, path_key in f_candidates:
            orphan_specs.append((title, mark, revision, source, path_key, ""))
        for title, mark, revision, source, path_key, source_path in file_candidates:
            orphan_specs.append((title, mark, revision, source, path_key, source_path))
        for title, mark, revision, source, path_key in auto_candidates:
            orphan_specs.append((title, mark, revision, source, path_key, path_key))

        for title, mark, revision, source, path_key, source_path in orphan_specs:
            kit = kit_identity_key(title, mark)
            if _has_equivalent_revision(revision, represented, kit):
                continue
            stamp = (kit, _canon_revision(revision))
            current = merged_orphans.get(stamp)
            if current is None:
                merged_orphans[stamp] = _MergedOrphan(
                    title=title,
                    mark=mark,
                    revision_text=revision,
                    sources=frozenset({source}),
                    path_key=path_key,
                    source_path=source_path,
                    path_source=source,
                )
                continue
            prefer_path = _orphan_source_rank(source) < _orphan_source_rank(
                current.path_source
            )
            merged_orphans[stamp] = _MergedOrphan(
                title=title if prefer_path else current.title,
                mark=mark if prefer_path else current.mark,
                revision_text=current.revision_text,
                sources=current.sources | {source},
                path_key=path_key if prefer_path else current.path_key,
                source_path=source_path if prefer_path else current.source_path,
                path_source=source if prefer_path else current.path_source,
            )

        for (kit, _rev_key), orphan in merged_orphans.items():
            source = _preferred_orphan_source(orphan.sources)
            flags = _flags_for(
                kit, orphan.revision_text, f_revs, rd_revs, robot_revs, auto_revs
            )
            rows.append(
                _journal_row(
                    title=orphan.title,
                    mark=orphan.mark,
                    kind="orphan",
                    source=source,
                    revision_text=orphan.revision_text,
                    decision="",
                    identity_fingerprint=orphan_identity_fingerprint(
                        source, orphan.title, orphan.mark, orphan.revision_text, ""
                    ),
                    source_path=orphan.source_path,
                    path_key=orphan.path_key,
                    in_f=flags[0],
                    in_rd=flags[1],
                    in_robot=flags[2],
                    in_auto_mto=flags[3],
                )
            )

        rows.sort(
            key=lambda row: (
                row.title,
                row.mark,
                row.send_date_sortable,
                row.revision_text,
                row.kind,
                row.source,
                row.path_key,
            )
        )
        return tuple(rows)
