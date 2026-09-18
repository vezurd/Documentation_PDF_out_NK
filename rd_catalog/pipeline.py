"""Match issuance sends to F events and packages; derive pipeline status.

Qt-free. Persists ``kit_package`` / ``kit_cycle`` / ``kit_pipeline`` via
:meth:`CatalogDatabase.replace_kit_derived`, then ``kit_revision_cell`` via
:func:`rebuild_revision_matrix`. Does not touch ``kit_liquidity_review``.
"""

from __future__ import annotations

import json
import re
from collections.abc import Collection, Mapping, Sequence
from dataclasses import dataclass, replace
from datetime import date, datetime
from enum import StrEnum
from pathlib import Path, PureWindowsPath

from rd_catalog.db import (
    CatalogDatabase,
    KitAnnulledFlagRow,
    KitCycleRow,
    KitPackageRow,
    KitPipelineRow,
    KitRevisionRow,
    KitWorkingFlagRow,
    LiquidityReviewRow,
)
from rd_catalog.kits import (
    GoogleKit,
    IssuanceKit,
    KitEvent,
    format_event_date_sortable,
    format_revision,
    is_rd_kit_mark,
    kit_identity_key,
    parse_sheet_revision,
    revisions_equivalent,
)
from rd_catalog.models import (
    CollisionKind,
    FileKind,
    FileRecord,
    OverlayResult,
    ParsedFile,
    ParseStatus,
    SourceKind,
    TransferMetadata,
)
from rd_catalog.overlay import (
    _newest_sort_key,
    build_rd_overlays_as_of,
    document_key_text,
    revision_rank,
)
from rd_catalog.parse import (
    issued_package_dir,
    normalize_unicode_dashes,
    path_is_as_build,
    record_has_canonical_layout,
    transfer_name_is_void,
)
from rd_catalog.path_actions import path_is_under
from rd_catalog.perf_log import perf_span


PIPELINE_STATUS_ALGORITHM_VERSION = 10
REVISION_MATRIX_ALGORITHM_VERSION = 1

PIPELINE_DISPLAY_V1 = 1
PIPELINE_DISPLAY_V2 = 2
PIPELINE_DISPLAY_V3 = 3
# Комплекты «Статус рассмотрения / согласования» (Qt + WEB). Не схема БД.
# 1 — текущее A/B/C всегда в «Статус согласования»; рассмотрение без буквы B/C.
# 2 — текущие B/C в рассмотрении (``Прошел ТДО · B``); согласование только
#     для письма не этого цикла (stale vs диск). Текущий A не дублируется.
# 3 — рассмотрение по циклу Google-таблиц (Выдача → D/E → F), не по папке РД.
#     Когда табличная рев. = РД, подпись как v2 (``РД {рев.}``). Переключить:
#     2 или 1, если v3 не зайдёт.
PIPELINE_DISPLAY_VERSION = PIPELINE_DISPLAY_V3

PIPELINE_FACE_RD = "rd"
PIPELINE_FACE_ISSUANCE = "issuance"
PIPELINE_FACE_DE = "de"
PIPELINE_FACE_F = "f"
_PIPELINE_FACE_TOKEN = {
    PIPELINE_FACE_RD: "РД",
    PIPELINE_FACE_ISSUANCE: "выдача",
    PIPELINE_FACE_DE: "D/E",
    PIPELINE_FACE_F: "F",
}


def pipeline_algorithm_needs_rebuild(
    pipelines: Sequence[KitPipelineRow],
) -> bool:
    """Return whether stored ``kit_pipeline`` rows predate the current algorithm.

    Empty tables are handled by the empty-derived hatch, not here.

    Args:
        pipelines: Rows from ``list_kit_pipelines``.

    Returns:
        True when at least one row has ``algorithm_version`` below
        ``PIPELINE_STATUS_ALGORITHM_VERSION``.
    """

    if not pipelines:
        return False
    return any(
        int(row.algorithm_version or 0) < PIPELINE_STATUS_ALGORITHM_VERSION
        for row in pipelines
    )

_TITLE_RE = re.compile(r"^\d{4}$")
_DATE_RE = re.compile(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$")
_NS = 1_000_000_000
_TDO_STAGE_PRIORITY = {
    "tdo_passed": 0,
    "incoming_passed": 1,
    "tdo_sent": 2,
    "incoming_sent": 3,
    "sr_upload": 4,
}
_CODE_STAGES = {
    "code_a": ("A", "code_a"),
    "code_b": ("B", "code_b"),
    "code_c": ("C", "code_c"),
}
_CODE_STAGE_KEYS = frozenset(_CODE_STAGES)
_PASSED_STAGES = frozenset({"tdo_passed", "incoming_passed"})
_SENT_STAGES = frozenset({"tdo_sent", "incoming_sent", "sr_upload"})
_SEND_DATE_STAGES = frozenset({"tdo_sent", "incoming_sent"})
_REVISION_STATUS_LETTER = {
    "code_a": "A",
    "code_b": "B",
    "code_c": "C",
    "tdo_review": "T",
    "working": "W",
    "annulled": "Ан",
    "no_mto": "–",
}
_SKIPPED_HEATMAP_STATUSES = frozenset({"working", "annulled"})
_AB_LETTER = "AB"
_LETTER_JOIN = "·"
_MATRIX_PROBLEM_KINDS = frozenset(
    {
        CollisionKind.TRANSFER_ORDER_CONFLICT.value,
        CollisionKind.TRANSFER_MTIME_CONFLICT.value,
        CollisionKind.DUP_SAME_REVISION.value,
        CollisionKind.FILE_MISSING.value,
    }
)
_MTO_CELL_PREDICATES = frozenset(
    {"code_a", "tdo_passed", "current", "current_ifc", "exclude_as_build"}
)
_TDO_PASSED_STATUSES = frozenset(
    {"tdo_review", "code_a", "code_b", "code_c"}
)
_APPROVED_CONTOUR_COLLISIONS = frozenset(
    {
        CollisionKind.DUP_SAME_REVISION.value,
        CollisionKind.TRANSFER_ORDER_CONFLICT.value,
    }
)
_CYCLE_PACKAGE_REASONS = frozenset({"trm", "revision", "date"})
_REVISION_MATCH_EXACT = 2
_REVISION_MATCH_BASE = 1
# Folder ``10_рев.AN01_…`` has no base digit after ``рев.``. Do not feed
# ``AN01`` to ``parse_sheet_revision`` — it would parse as revision ``01``.
_APPENDIX_ONLY_TOKEN_RE = re.compile(
    r"^(?:(?:рев(?:изия)?|rev(?:ision)?)\s*[._:\-]?\s*)?AN(?P<an>\d{1,2})$",
    re.IGNORECASE,
)
_FOLDER_APPENDIX_ONLY_RE = re.compile(
    r"(?:рев(?:изия)?|rev(?:ision)?)\s*[._:\-]?\s*AN(?P<an>\d{1,2})(?!\d)",
    re.IGNORECASE,
)
_AMBIGUITY_ORDER = (
    "dup_same_revision",
    "transfer_order_conflict",
    "no_mto_in_contour",
    "no_package",
    "journal_only",
)
_AMBIGUITY_WARNINGS = {
    "dup_same_revision": (
        "В согласованном контуре несколько файлов MTO с одной ревизией имени."
    ),
    "transfer_order_conflict": (
        "В согласованном контуре конфликт порядка передач."
    ),
    "no_mto_in_contour": "В согласованном контуре нет файла MTO.",
    "no_package": "Не найден пакет РД для согласованной ревизии.",
    "journal_only": "Комплект есть только в журнале выдачи, пакета РД нет.",
}


class KitPipelineStatus(StrEnum):
    """Current official-cycle review status for one kit."""

    NOT_UPLOADED = "not_uploaded"
    SENT_TDO = "sent_tdo"
    TDO_REVIEW = "tdo_review"
    AGREED = "agreed"


_STATUS_LABELS: dict[KitPipelineStatus, str] = {
    KitPipelineStatus.NOT_UPLOADED: "Не загружен в СР",
    KitPipelineStatus.SENT_TDO: "Отправлен на ТДО",
    KitPipelineStatus.TDO_REVIEW: "Прошел ТДО",
    KitPipelineStatus.AGREED: "Согласован",
}
_STATUS_SHORT_LABELS: dict[KitPipelineStatus, str] = {
    KitPipelineStatus.NOT_UPLOADED: "не в СР",
    KitPipelineStatus.SENT_TDO: "на ТДО",
    KitPipelineStatus.TDO_REVIEW: "ТДО",
    KitPipelineStatus.AGREED: "согласован",
}


@dataclass(frozen=True, slots=True)
class KitCard:
    """Read model for the kit card GUI (no F/path re-parsing)."""

    title: str
    mark: str
    pipeline: KitPipelineRow | None
    packages: tuple[KitPackageRow, ...]
    cycles: tuple[KitCycleRow, ...]
    google: GoogleKit | None
    working_revision_text: str = ""
    official_revision_text: str = ""
    liquidity_reviews: tuple[LiquidityReviewRow, ...] = ()
    issuance: IssuanceKit | None = None


@dataclass(frozen=True, slots=True)
class FolderTreeHint:
    """Google/pipeline extras for one RD transfer folder in «Все документы».

    ``review_status`` is empty when no issuance/F cycle matched this folder
    (TRM, revision, or date ±2 days). Working-revision is independent of that
    match. ``working_origin`` is ``manual`` for ``kit_working_flag`` and
    ``auto`` when the filename rev is strictly above the last send/F.
    ``is_annulled`` is independent of working; when both are somehow
    true, display treats the folder as annulled (``is_working=False``).

    MTO fields come from the ``kit_revision_cell`` snapshot for the revision
    of the MTO in that folder. ``mto_status`` is empty when no cell matched.
    It is the Google F / heatmap cell for that filename rev, not the
    working-folder mark (a working as-build of ``01-AN02`` can still be
    ``code_a``).
    """

    review_status: str = ""
    review_full: str = ""
    match_reason: str = ""
    f_label: str = ""
    is_working: bool = False
    is_annulled: bool = False
    working_origin: str = ""
    mto_status: str = ""
    mto_letters: str = ""
    mto_revision_text: str = ""
    has_mto_file: bool = False
    is_current_mto: bool = False
    is_current_ifc: bool = False
    problem_kinds: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class MtoWorklistRow:
    """One MTO worklist row: kit, revision, file, and presence flags.

    ``gap_kind`` is ``no_package``, ``no_mto_file``, or ``no_rd`` (a
    Google/issuance kit with no heatmap cells). Empty means the revision
    has an RD folder and an MTO file.
    """

    title: str
    mark: str
    revision_text: str
    status: str
    letters: str
    is_as_build: bool
    is_current: bool
    is_current_ifc: bool
    has_mto: bool
    mto_path: str = ""
    mto_mtime_ns: int | None = None
    package_path: str = ""
    package_label: str = ""
    package_count: int = 0
    problem_kinds: tuple[str, ...] = ()
    gap_kind: str = ""
    in_google: bool = False
    in_issuance: bool = False
    has_f_status: bool = False


@dataclass(frozen=True, slots=True)
class ApprovedContour:
    """Approved Google/issuance contour for one kit, with bounded overlay MTO.

    ``approved_revision_text`` is the journal revision (code, else official).
    Folder ``рев.*`` is never displayed here; it is only a fallback key when
    matching which package the journal is talking about.
    """

    title: str
    mark: str
    approved_revision_text: str
    package_id: int | None
    package_path: str
    package_sequence: int | None
    match_reason: str
    confidence: str
    mto_path: str
    mto_source: str
    ambiguity: tuple[str, ...]
    warnings: tuple[str, ...]


def pipeline_status_label(status: str) -> str:
    """Return the Russian GUI label for a pipeline status value.

    Args:
        status: Stored ``kit_pipeline.status`` string.

    Returns:
        Localized label, or the raw value when unknown.
    """

    try:
        return _STATUS_LABELS[KitPipelineStatus(status)]
    except ValueError:
        return status


def pipeline_status_short_label(status: str) -> str:
    """Return a compact Russian badge for a documents-tree node.

    Args:
        status: Stored ``kit_pipeline.status`` string.

    Returns:
        Short label such as ``согласован``, or the raw value when unknown.
    """

    try:
        return _STATUS_SHORT_LABELS[KitPipelineStatus(status)]
    except ValueError:
        return status


APPROVAL_REL_AHEAD = "новее диска"
APPROVAL_REL_PREVIOUS = "прошлый цикл"
APPROVAL_REL_NO_REV = "без рев. в F"
APPROVAL_REL_OTHER = "не этого цикла"


def pipeline_approval_relation(row: KitPipelineRow) -> str:
    """Return how a stale F letter relates to the official disk cycle.

    Args:
        row: Derived ``kit_pipeline`` row.

    Returns:
        Empty when there is no letter or the letter belongs to the
        official cycle. Otherwise one of ``новее диска``, ``прошлый цикл``,
        ``без рев. в F``, ``не этого цикла``.
    """

    if not row.code or not row.code_stale:
        return ""
    code_rev = (row.code_revision_text or "").strip()
    official = (row.official_revision_text or "").strip()
    if not code_rev:
        return APPROVAL_REL_NO_REV
    if not official:
        return APPROVAL_REL_OTHER
    if revision_texts_equivalent(code_rev, official):
        return APPROVAL_REL_PREVIOUS
    if _text_rank(code_rev) > _text_rank(official):
        return APPROVAL_REL_AHEAD
    return APPROVAL_REL_PREVIOUS


@dataclass(frozen=True, slots=True)
class PipelineReviewDisplay:
    """Display-only review cell facts (does not persist ``kit_pipeline``).

    ``uses_sheet_face`` is True on v3 when the Google-table revision
    differs from the RD-folder official revision.
    """

    status: str
    face_revision: str
    face_source: str
    uses_sheet_face: bool
    pass_date: str
    send_date: str
    cycle_letter: str
    agreed_date: str = ""


def _display_version(version: int | None) -> int:
    if version is None:
        return PIPELINE_DISPLAY_VERSION
    return version


def _resolve_pipeline_events(
    google: GoogleKit | None,
    events: Sequence[KitEvent],
) -> tuple[KitEvent, ...]:
    if events:
        return tuple(events)
    if google is not None:
        return tuple(google.events)
    return ()


def pipeline_sheet_face_revision(
    row: KitPipelineRow,
    *,
    google: GoogleKit | None = None,
    issuance: IssuanceKit | None = None,
    events: Sequence[KitEvent] = (),
) -> tuple[str, str]:
    """Return ``(revision_text, source)`` for the Google-facing review cycle.

    Order: last effective issuance, else D/E ``sheet_revision_text``, else
    last F event with a revision, else the RD-folder official revision.

    Args:
        row: Derived ``kit_pipeline`` row (disk fallback).
        google: КСБ ИД kit (D/E + F).
        issuance: Latest effective «Выдача РД ПД» send.
        events: Optional F events; empty uses ``google.events``.

    Returns:
        Face revision text and one of ``issuance`` / ``de`` / ``f`` / ``rd``.
    """

    if issuance is not None:
        text = (issuance.revision_text or "").strip()
        if text:
            return text, PIPELINE_FACE_ISSUANCE
    if google is not None:
        sheet = (google.sheet_revision_text or "").strip()
        if sheet:
            return sheet, PIPELINE_FACE_DE
    for event in reversed(_resolve_pipeline_events(google, events)):
        rev = format_revision(event.revision, event.appendix)
        if rev:
            return rev, PIPELINE_FACE_F
    official = (row.official_revision_text or "").strip()
    return official, PIPELINE_FACE_RD


def pipeline_display_uses_sheet_face(
    row: KitPipelineRow,
    face_revision: str,
) -> bool:
    """Return whether v3 should paint the Google cycle instead of disk.

    Args:
        row: Derived ``kit_pipeline`` row.
        face_revision: Google-facing revision from
            :func:`pipeline_sheet_face_revision`.

    Returns:
        True when the table cycle has a revision that is not the same as
        ``official_revision_text``.
    """

    face = (face_revision or "").strip()
    if not face:
        return False
    official = (row.official_revision_text or "").strip()
    if not official:
        return True
    return not revision_texts_equivalent(face, official)


def pipeline_letter_on_sheet_face(
    row: KitPipelineRow,
    face_revision: str,
) -> bool:
    """Return whether the stored F letter belongs to the Google face cycle.

    Same-rev stale vs disk (9000-KSB: later TDO of the same rev) is not
    the table face letter. A letter on a newer table rev (6550-SKUD)
    is on the face even when ``code_stale`` vs disk.

    Args:
        row: Derived ``kit_pipeline`` row.
        face_revision: Google-facing revision.

    Returns:
        True when the letter revision matches the face and is not a
        previous cycle of the same disk rev.
    """

    code_rev = (row.code_revision_text or "").strip()
    face = (face_revision or "").strip()
    if not code_rev or not face:
        return False
    if not revision_texts_equivalent(code_rev, face):
        return False
    official = (row.official_revision_text or "").strip()
    if (
        row.code_stale
        and official
        and revision_texts_equivalent(code_rev, official)
    ):
        return False
    return True


def _send_date_for_revision(
    revision_text: str,
    *,
    events: Sequence[KitEvent],
    issuance: IssuanceKit | None,
) -> str:
    official = (revision_text or "").strip()
    if issuance is not None:
        send_date = (issuance.send_date or "").strip()
        if send_date:
            send_rev = (issuance.revision_text or "").strip()
            if (
                not official
                or not send_rev
                or revision_texts_equivalent(send_rev, official)
            ):
                return _code_date_text(send_date)
    send_dates: list[str] = []
    for event in events:
        if not event.date:
            continue
        if event.stage not in _SEND_DATE_STAGES:
            continue
        event_rev = format_revision(event.revision, event.appendix)
        if official and event_rev:
            if not revision_texts_equivalent(event_rev, official):
                continue
        elif official and not event_rev:
            continue
        send_dates.append(event.date)
    if send_dates:
        return _code_date_text(send_dates[-1])
    return ""


def _disk_tdo_passed_date_text(
    row: KitPipelineRow,
    *,
    events: Sequence[KitEvent],
    issuance: IssuanceKit | None,
) -> str:
    if row.status != KitPipelineStatus.TDO_REVIEW.value:
        return ""
    stored = _code_date_text(row.tdo_date)
    if stored:
        return stored
    return _code_date_text(
        _tdo_passed_date_from_facts(
            events=events,
            issuance=issuance,
            official_revision=row.official_revision_text,
        )
    )


def _sheet_review_status(
    *,
    face_revision: str,
    events: Sequence[KitEvent],
    issuance: IssuanceKit | None,
) -> str:
    """Return a display status for the Google face rev (not persisted)."""

    numbered = list(enumerate(events, start=1))
    last_send: IssuanceKit | None = None
    if issuance is not None:
        send_rev = (issuance.revision_text or "").strip()
        face = (face_revision or "").strip()
        if not face or not send_rev or revision_texts_equivalent(send_rev, face):
            last_send = issuance
    status = _review_status(
        official_revision=face_revision,
        last_send=last_send,
        last_cycle=None,
        events=numbered,
        has_official_send=last_send is not None,
        allow_agreed=True,
    )
    if (
        status in {KitPipelineStatus.NOT_UPLOADED, KitPipelineStatus.SENT_TDO}
        and last_send is not None
        and _issuance_accepted(last_send)
    ):
        return KitPipelineStatus.TDO_REVIEW.value
    return status.value


def _cycle_letter_for_display(
    row: KitPipelineRow,
    *,
    version: int,
    uses_sheet_face: bool,
    face_revision: str,
) -> str:
    if version == PIPELINE_DISPLAY_V1:
        return ""
    letter = (row.code or "").strip().upper()
    if letter not in {"B", "C"}:
        return ""
    if uses_sheet_face:
        return letter if pipeline_letter_on_sheet_face(row, face_revision) else ""
    if row.code_stale:
        return ""
    return letter


def _agreed_date_for_display(
    row: KitPipelineRow,
    *,
    status: str,
    face_revision: str,
    uses_sheet_face: bool,
    events: Sequence[KitEvent],
) -> str:
    """Return DD.MM.YYYY of A / F «согласовано» on the displayed cycle."""

    if status != KitPipelineStatus.AGREED.value:
        return ""
    face = (face_revision or "").strip()
    event_dates: list[str] = []
    for event in events:
        if event.stage not in {"code_a", "agreed"}:
            continue
        if not event.date:
            continue
        event_rev = format_revision(event.revision, event.appendix)
        if face and event_rev:
            if not revision_texts_equivalent(event_rev, face):
                continue
        elif face and not event_rev:
            continue
        event_dates.append(event.date)
    if event_dates:
        return _code_date_text(event_dates[-1])
    letter = (row.code or "").strip().upper()
    if letter != "A":
        return ""
    if uses_sheet_face:
        if not pipeline_letter_on_sheet_face(row, face):
            return ""
    elif row.code_stale:
        return ""
    return _code_date_text(row.code_date)


def pipeline_review_display(
    row: KitPipelineRow,
    *,
    events: Sequence[KitEvent] = (),
    issuance: IssuanceKit | None = None,
    google: GoogleKit | None = None,
    version: int | None = None,
) -> PipelineReviewDisplay:
    """Return display-only review facts for Комплекты cells.

    v1/v2 and v3-when-tables-match-disk use stored ``kit_pipeline.status``
    and the RD-folder official revision. v3 with a different Google face
    recomputes the stage from F + issuance on that face. Does not write
    the database.

    Args:
        row: Derived ``kit_pipeline`` row.
        events: Optional F events; empty uses ``google.events``.
        issuance: Latest effective send.
        google: КСБ ИД kit (D/E revision + F).
        version: Display version; ``None`` uses ``PIPELINE_DISPLAY_VERSION``.

    Returns:
        Face revision, display status, pass/send dates, and B/C suffix.
    """

    ver = _display_version(version)
    resolved = _resolve_pipeline_events(google, events)
    face, source = pipeline_sheet_face_revision(
        row, google=google, issuance=issuance, events=resolved
    )
    uses_sheet = ver >= PIPELINE_DISPLAY_V3 and pipeline_display_uses_sheet_face(
        row, face
    )
    if uses_sheet:
        status = _sheet_review_status(
            face_revision=face, events=resolved, issuance=issuance
        )
        pass_date = ""
        if status == KitPipelineStatus.TDO_REVIEW.value:
            matched = issuance
            send_rev = (issuance.revision_text or "").strip() if issuance else ""
            if (
                issuance is not None
                and face
                and send_rev
                and not revision_texts_equivalent(send_rev, face)
            ):
                matched = None
            pass_date = _code_date_text(
                _tdo_passed_date_from_facts(
                    events=resolved,
                    issuance=matched,
                    official_revision=face,
                )
            )
        send_date = _send_date_for_revision(
            face, events=resolved, issuance=issuance
        )
        return PipelineReviewDisplay(
            status=status,
            face_revision=face,
            face_source=source,
            uses_sheet_face=True,
            pass_date=pass_date,
            send_date=send_date,
            cycle_letter=_cycle_letter_for_display(
                row,
                version=ver,
                uses_sheet_face=True,
                face_revision=face,
            ),
            agreed_date=_agreed_date_for_display(
                row,
                status=status,
                face_revision=face,
                uses_sheet_face=True,
                events=resolved,
            ),
        )
    official = (row.official_revision_text or "").strip()
    status = row.status
    return PipelineReviewDisplay(
        status=status,
        face_revision=official,
        face_source=PIPELINE_FACE_RD,
        uses_sheet_face=False,
        pass_date=_disk_tdo_passed_date_text(
            row, events=resolved, issuance=issuance
        ),
        send_date=_send_date_for_revision(
            official, events=resolved, issuance=issuance
        ),
        cycle_letter=_cycle_letter_for_display(
            row,
            version=ver,
            uses_sheet_face=False,
            face_revision=official,
        ),
        agreed_date=_agreed_date_for_display(
            row,
            status=status,
            face_revision=official,
            uses_sheet_face=False,
            events=resolved,
        ),
    )


def pipeline_display_review_status(
    row: KitPipelineRow,
    *,
    events: Sequence[KitEvent] = (),
    issuance: IssuanceKit | None = None,
    google: GoogleKit | None = None,
    version: int | None = None,
) -> str:
    """Return the review fill/filter status for the active display version.

    Args:
        row: Derived ``kit_pipeline`` row.
        events: Optional F events.
        issuance: Latest effective send.
        google: КСБ ИД kit.
        version: Display version; ``None`` uses ``PIPELINE_DISPLAY_VERSION``.

    Returns:
        ``not_uploaded`` / ``sent_tdo`` / ``tdo_review`` / ``agreed``.
    """

    return pipeline_review_display(
        row,
        events=events,
        issuance=issuance,
        google=google,
        version=version,
    ).status


def pipeline_display_code_a(
    row: KitPipelineRow,
    *,
    events: Sequence[KitEvent] = (),
    issuance: IssuanceKit | None = None,
    google: GoogleKit | None = None,
    version: int | None = None,
) -> bool:
    """Return whether filters should treat the kit as current letter A.

    Args:
        row: Derived ``kit_pipeline`` row.
        events: Optional F events.
        issuance: Latest effective send.
        google: КСБ ИД kit.
        version: Display version; ``None`` uses ``PIPELINE_DISPLAY_VERSION``.

    Returns:
        v1/v2 and v3-aligned: ``code == A`` and not ``code_stale``.
        v3 sheet face: ``code == A`` on the Google face cycle.
    """

    if (row.code or "").strip().upper() != "A":
        return False
    view = pipeline_review_display(
        row,
        events=events,
        issuance=issuance,
        google=google,
        version=version,
    )
    if view.uses_sheet_face:
        return pipeline_letter_on_sheet_face(row, view.face_revision)
    return not row.code_stale


def pipeline_review_cycle_letter(
    row: KitPipelineRow,
    *,
    version: int | None = None,
    events: Sequence[KitEvent] = (),
    issuance: IssuanceKit | None = None,
    google: GoogleKit | None = None,
) -> str:
    """Return a current-cycle B/C to show on the review label.

    Args:
        row: Derived ``kit_pipeline`` row.
        version: Display version; ``None`` uses ``PIPELINE_DISPLAY_VERSION``.
        events: Optional F events for the v3 face.
        issuance: Latest effective send for the v3 face.
        google: КСБ ИД kit for the v3 face.

    Returns:
        ``B`` / ``C`` when the letter belongs to the displayed cycle.
        Never ``A``. Empty on v1 and when the letter is missing or of
        another cycle.
    """

    return pipeline_review_display(
        row,
        events=events,
        issuance=issuance,
        google=google,
        version=version,
    ).cycle_letter


def pipeline_approval_shows_letter(
    row: KitPipelineRow,
    *,
    version: int | None = None,
    events: Sequence[KitEvent] = (),
    issuance: IssuanceKit | None = None,
    google: GoogleKit | None = None,
) -> bool:
    """Return whether the approval column should show the F letter.

    Args:
        row: Derived ``kit_pipeline`` row.
        version: Display version; ``None`` uses ``PIPELINE_DISPLAY_VERSION``.
        events: Optional F events for the v3 face.
        issuance: Latest effective send for the v3 face.
        google: КСБ ИД kit for the v3 face.

    Returns:
        v1: True when a letter exists. v2 and v3-aligned: True only when
        ``code_stale``. v3 sheet face: True when the letter is not of the
        Google face cycle.
    """

    if not row.code:
        return False
    ver = _display_version(version)
    if ver == PIPELINE_DISPLAY_V1:
        return True
    view = pipeline_review_display(
        row,
        events=events,
        issuance=issuance,
        google=google,
        version=ver,
    )
    if view.uses_sheet_face:
        if not pipeline_letter_on_sheet_face(row, view.face_revision):
            return True
        letter = (row.code or "").strip().upper()
        if letter in {"B", "C"}:
            return False
        if letter == "A" and view.status == KitPipelineStatus.AGREED.value:
            return False
        return True
    return bool(row.code_stale)


def _format_review_label(row: KitPipelineRow, view: PipelineReviewDisplay) -> str:
    label = pipeline_status_label(view.status)
    paren_date = view.pass_date or view.agreed_date
    if paren_date:
        label = f"{label} ({paren_date})"
    if view.cycle_letter:
        label = f"{label} · {view.cycle_letter}"
    if view.face_revision:
        token = (
            _PIPELINE_FACE_TOKEN.get(view.face_source, "F")
            if view.uses_sheet_face
            else "РД"
        )
        label = f"{label} · {token} {view.face_revision}"
    if view.send_date:
        label = f"{label} · отпр. {view.send_date}"
    if row.review_as_build:
        label = f"{label} (AB)"
    return label


def pipeline_review_label(
    row: KitPipelineRow,
    *,
    events: Sequence[KitEvent] = (),
    issuance: IssuanceKit | None = None,
    google: GoogleKit | None = None,
    version: int | None = None,
) -> str:
    """Return the review label for one pipeline row.

    Args:
        row: Derived ``kit_pipeline`` row.
        events: Optional F events for dates and the v3 face.
        issuance: Optional last send for dates and the v3 face.
        google: Optional КСБ ИД kit (D/E revision when events are empty).
        version: Display version; ``None`` uses ``PIPELINE_DISPLAY_VERSION``.

    Returns:
        ``{status}[(pass or agreed date)][ · B|C] · {РД|выдача|D/E|F} {rev}[ · отпр. {date}][ (AB)]``.
        On v1/v2 and when Google matches disk, ``РД`` is the official
        package (same as «РД · рев.»). v3 with a different table cycle
        uses ``выдача`` / ``D/E`` / ``F`` and the table revision.
        B/C never make the status «Согласован». For ``agreed``, the
        parentheses date is the F code A / «согласовано» of this cycle,
        not the send date.
    """

    view = pipeline_review_display(
        row,
        events=events,
        issuance=issuance,
        google=google,
        version=version,
    )
    return _format_review_label(row, view)


def pipeline_tdo_passed_date_text(
    row: KitPipelineRow,
    *,
    events: Sequence[KitEvent] = (),
    issuance: IssuanceKit | None = None,
    google: GoogleKit | None = None,
    version: int | None = None,
) -> str:
    """Return ``DD.MM.YYYY`` of the displayed-cycle TDO/incoming pass.

    Args:
        row: Derived ``kit_pipeline`` row.
        events: F events used when ``row.tdo_date`` is empty, or for v3 face.
        issuance: Last send used when stored/F dates are empty.
        google: КСБ ИД kit for the v3 face.
        version: Display version; ``None`` uses ``PIPELINE_DISPLAY_VERSION``.

    Returns:
        Normalized date text, or ``""``.
    """

    return pipeline_review_display(
        row,
        events=events,
        issuance=issuance,
        google=google,
        version=version,
    ).pass_date


def pipeline_send_date_text(
    row: KitPipelineRow,
    *,
    events: Sequence[KitEvent] = (),
    issuance: IssuanceKit | None = None,
    google: GoogleKit | None = None,
    version: int | None = None,
) -> str:
    """Return ``DD.MM.YYYY`` of the displayed-cycle send.

    Prefers ``issuance.send_date`` when the send revision matches the
    displayed face (v3) or the official revision (v1/v2). Otherwise uses
    the last F ``tdo_sent`` / ``incoming_sent`` on that rev.

    Args:
        row: Derived ``kit_pipeline`` row.
        events: F events used when issuance does not match the face.
        issuance: Last send of the kit.
        google: КСБ ИД kit for the v3 face.
        version: Display version; ``None`` uses ``PIPELINE_DISPLAY_VERSION``.

    Returns:
        Normalized date text, or ``""``.
    """

    return pipeline_review_display(
        row,
        events=events,
        issuance=issuance,
        google=google,
        version=version,
    ).send_date


def pipeline_approval_label(
    row: KitPipelineRow,
    *,
    version: int | None = None,
    events: Sequence[KitEvent] = (),
    issuance: IssuanceKit | None = None,
    google: GoogleKit | None = None,
) -> str:
    """Return the approval letter label for one pipeline row.

    Args:
        row: Derived ``kit_pipeline`` row.
        version: Display version; ``None`` uses ``PIPELINE_DISPLAY_VERSION``.
        events: Optional F events for the v3 face.
        issuance: Latest effective send for the v3 face.
        google: КСБ ИД kit for the v3 face.

    Returns:
        ``—`` when no code, and when the letter belongs to the displayed
        review cycle (current A is «Согласован»; current B/C sit on
        review). Otherwise ``A · {rev} · {DD.MM.YYYY}``. A letter not of
        this cycle appends ``новее диска`` / ``прошлый цикл`` /
        ``без рев. в F`` / ``не этого цикла``.
    """

    if not pipeline_approval_shows_letter(
        row,
        version=version,
        events=events,
        issuance=issuance,
        google=google,
    ):
        return "—"
    parts = [row.code]
    if row.code_revision_text:
        parts.append(row.code_revision_text)
    code_date = _code_date_text(row.code_date)
    if code_date:
        parts.append(code_date)
    relation = pipeline_approval_relation(row)
    if relation:
        parts.append(relation)
    return " · ".join(parts)


def pipeline_approval_color_key(
    row: KitPipelineRow,
    *,
    version: int | None = None,
    events: Sequence[KitEvent] = (),
    issuance: IssuanceKit | None = None,
    google: GoogleKit | None = None,
) -> str:
    """Return the GUI color key for the approval letter badge.

    Args:
        row: Derived ``kit_pipeline`` row.
        version: Display version; ``None`` uses ``PIPELINE_DISPLAY_VERSION``.
        events: Optional F events for the v3 face.
        issuance: Latest effective send for the v3 face.
        google: КСБ ИД kit for the v3 face.

    Returns:
        Empty string when the approval cell is ``—``;
        ``approval_stale`` when stale vs disk; otherwise ``code_a`` /
        ``code_b`` / ``code_c`` (v1 current letters).
    """

    if not pipeline_approval_shows_letter(
        row,
        version=version,
        events=events,
        issuance=issuance,
        google=google,
    ):
        return ""
    if row.code_stale:
        return "approval_stale"
    letter = row.code.upper()
    if letter == "A":
        return "code_a"
    if letter == "B":
        return "code_b"
    if letter == "C":
        return "code_c"
    return ""


def ingest_google_snapshot(
    database: CatalogDatabase,
    kits: Sequence[GoogleKit],
    sends: Sequence[IssuanceKit],
    *,
    loaded_at: str,
    source: str,
    warning: str | None = None,
) -> None:
    """Replace Google kit, event, and issuance-send rows.

    Does not rebuild derived pipeline tables; the GUI calls
    :func:`rebuild_pipeline` after a successful refresh. After the
    snapshot is stored, rematches ``issuance_review`` rows with
    ``kind=send`` (decisions are not cleared).

    Args:
        database: Initialized catalog database.
        kits: Parsed KSB ИД kits.
        sends: Every parsed «Выдача РД ПД» row.
        loaded_at: ISO timestamp of this snapshot load.
        source: Load origin (``network``, ``cache``, test label, …).
        warning: Optional load warning stored on ``google_load``.
    """

    with perf_span("pipeline.ingest_google", kits=len(kits), sends=len(sends)):
        from rd_catalog.issuance_review import rematch_issuance_reviews

        database.replace_google_snapshot(
            kits,
            sends,
            loaded_at=loaded_at,
            source=source,
            warning=warning,
        )
        rematch_issuance_reviews(database)


def list_kit_pipelines(database: CatalogDatabase) -> tuple[KitPipelineRow, ...]:
    """Return derived pipeline rows for the комплекты table.

    Args:
        database: Initialized catalog database.

    Returns:
        Pipeline rows in database order.
    """

    return tuple(database.list_kit_pipelines())


def _effective_issuance_sends(
    database: CatalogDatabase,
    title: str | None = None,
    mark: str | None = None,
) -> list[tuple[int | None, IssuanceKit]]:
    """Load include-contour sends; lazy-import avoids a cycle with rematch."""

    from rd_catalog.issuance_review import effective_issuance_sends

    return effective_issuance_sends(database, title, mark)


def get_kit_card(database: CatalogDatabase, title: str, mark: str) -> KitCard | None:
    """Load the persisted kit card for one identity.

    Args:
        database: Initialized catalog database.
        title: Four-digit title.
        mark: Latin AGCC mark.

    Returns:
        A :class:`KitCard`, or ``None`` when the identity is unknown.
    """

    with perf_span("pipeline.get_kit_card", title=title, mark=mark):
        key = kit_identity_key(title, mark)
        google = next(
            (
                kit
                for kit in database.list_google_kits()
                if kit_identity_key(kit.title, kit.mark) == key
            ),
            None,
        )
        pipelines = database.list_kit_pipelines(title, mark)
        pipeline = pipelines[0] if pipelines else None
        packages = tuple(database.list_kit_packages(title, mark))
        cycles = tuple(database.list_kit_cycles(title, mark))
        reviews = tuple(database.list_liquidity_reviews(title, mark))
        sends = _effective_issuance_sends(database, title, mark)
        issuance = sends[-1][1] if sends else None
        if (
            google is None
            and pipeline is None
            and not packages
            and not cycles
            and not reviews
            and issuance is None
        ):
            return None
        display_title = title
        display_mark = mark
        if google is not None:
            display_title, display_mark = google.title, google.mark
        elif issuance is not None:
            display_title, display_mark = issuance.title, issuance.mark
        elif packages:
            display_title, display_mark = packages[0].title, packages[0].mark
        return KitCard(
            title=display_title,
            mark=display_mark,
            pipeline=pipeline,
            packages=packages,
            cycles=cycles,
            google=google,
            working_revision_text=(
                pipeline.working_revision_text if pipeline is not None else ""
            ),
            official_revision_text=(
                pipeline.official_revision_text if pipeline is not None else ""
            ),
            liquidity_reviews=reviews,
            issuance=issuance,
        )


def list_folder_tree_hints(
    database: CatalogDatabase,
) -> dict[tuple[str, str, str], FolderTreeHint]:
    """Return per-folder Google/pipeline extras for the documents tree.

    Keys are ``(title.casefold(), mark.casefold(), transfer_or_folder.casefold())``.
    Grey and non-RD packages are omitted. Does not parse Google CSV.

    Args:
        database: Initialized catalog database.

    Returns:
        Lookup used by «Все документы» revision labels.
    """

    with perf_span("pipeline.list_folder_tree_hints"):
        packages = [
            pkg
            for pkg in database.list_kit_packages()
            if pkg.source == "rd" and not pkg.is_grey
        ]
        if not packages:
            return {}
        cycles = database.list_kit_cycles()
        pipelines = {
            kit_identity_key(row.title, row.mark): row
            for row in database.list_kit_pipelines()
        }
        sends_by_id: dict[int, IssuanceKit] = {}
        for send_id, send in database.list_issuance_sends_with_ids():
            sends_by_id[send_id] = send
        events_by_kit = database.list_google_events_by_kit()
        flags_by_key: dict[tuple[str, str], tuple[KitWorkingFlagRow, ...]] = {}
        for flag in database.list_working_flags():
            key = kit_identity_key(flag.title, flag.mark)
            flags_by_key[key] = flags_by_key.get(key, ()) + (flag,)
        cycles_by_kit: dict[tuple[str, str], list[KitCycleRow]] = {}
        for cycle in cycles:
            cycles_by_kit.setdefault(
                kit_identity_key(cycle.title, cycle.mark), []
            ).append(cycle)

        cells_by_kit_rev: dict[tuple[str, str, tuple], KitRevisionRow] = {}
        for cell in database.list_revision_cells():
            text = (cell.revision_text or "").strip()
            if not text:
                continue
            kit = kit_identity_key(cell.title, cell.mark)
            cells_by_kit_rev[(*kit, _revision_dedupe_key(text))] = cell

        hints: dict[tuple[str, str, str], FolderTreeHint] = {}
        for package in packages:
            key = kit_identity_key(package.title, package.mark)
            pipeline = pipelines.get(key)
            events = events_by_kit.get(key, ())
            kit_cycles = cycles_by_kit.get(key, ())
            mto_rev = _package_mto_revision(package)
            cell = (
                cells_by_kit_rev.get((*key, _revision_dedupe_key(mto_rev)))
                if mto_rev
                else None
            )
            hint = _hint_for_package(
                package,
                pipeline=pipeline,
                cycles=kit_cycles,
                events=events,
                sends_by_id=sends_by_id,
                cell=cell,
                flags=flags_by_key.get(key, ()),
            )
            names: list[str] = []
            if package.transfer_name:
                names.append(package.transfer_name)
            if package.package_path:
                folder = PureWindowsPath(package.package_path).name
                if folder:
                    names.append(folder)
            for name in names:
                folded = name.casefold()
                if folded:
                    hints[(*key, folded)] = hint
        return hints


def kit_keys_under_folders(
    records: Sequence[FileRecord],
    folders: Sequence[str],
) -> set[tuple[str, str]]:
    """Return kit identities whose files sit under any of ``folders``.

    Uses ``title`` / ``mark`` from scan data (including ``present=0``
    missing rows). Empty ``folders`` yields an empty set.

    Args:
        records: Catalog file rows (typically ``list_files()``).
        folders: UNC/local directory prefixes from a scoped rescan.

    Returns:
        ``kit_identity_key`` pairs for files under those prefixes.
    """

    if not folders:
        return set()
    keys: set[tuple[str, str]] = set()
    for record in records:
        if not any(path_is_under(record.path, folder) for folder in folders):
            continue
        title = str(record.data.get("title") or "").strip()
        mark = str(record.data.get("mark") or "").strip()
        if title and mark:
            keys.add(kit_identity_key(title, mark))
    return keys


def _normalize_kit_keys(
    kit_keys: Collection[tuple[str, str]] | None,
) -> set[tuple[str, str]] | None:
    if kit_keys is None:
        return None
    scoped = {kit_identity_key(title, mark) for title, mark in kit_keys}
    return scoped or None


def _records_in_contour(
    records: Sequence[FileRecord],
    rd_root: str | Path | None,
) -> Sequence[FileRecord]:
    root = str(rd_root or "").strip()
    if not root:
        return records
    return [
        record
        for record in records
        if record_has_canonical_layout(record, root)
    ]


def records_in_contour(
    records: Sequence[FileRecord],
    rd_root: str | Path | None,
    *,
    previous_records: Sequence[FileRecord] | None = None,
    previous_contour: Sequence[FileRecord] | None = None,
) -> Sequence[FileRecord]:
    """Keep SQ/ROBOT plus RD files on a canonical issued path.

    Args:
        records: Catalog file records from the last scan.
        rd_root: RD source root. Empty or ``None`` returns ``records``.
        previous_records: Prior ``list_files`` snapshot. When set with
            ``previous_contour``, unchanged ``id``/path/source rows reuse
            the last membership instead of walking every path.
        previous_contour: Contour computed from ``previous_records``.

    Returns:
        The filtered sequence (same objects, possibly a new list).
    """

    root = str(rd_root or "").strip()
    with perf_span("pipeline.records_in_contour", n=len(records)):
        if (
            root
            and previous_records
            and previous_contour is not None
        ):
            return _records_in_contour_reuse(
                records, root, previous_records, previous_contour
            )
        return _records_in_contour(records, rd_root)


def _records_in_contour_reuse(
    records: Sequence[FileRecord],
    rd_root: str,
    previous_records: Sequence[FileRecord],
    previous_contour: Sequence[FileRecord],
) -> list[FileRecord]:
    prev_stamp = {
        record.id: (record.path, record.source) for record in previous_records
    }
    members = {record.id for record in previous_contour}
    kept: list[FileRecord] = []
    for record in records:
        stamp = prev_stamp.get(record.id)
        if stamp == (record.path, record.source):
            if record.id in members:
                kept.append(record)
            continue
        if record_has_canonical_layout(record, rd_root):
            kept.append(record)
    return kept


def _data_kit_key(record: FileRecord) -> tuple[str, str] | None:
    title = str(record.data.get("title") or "").strip()
    mark = str(record.data.get("mark") or "").strip()
    if not title or not mark:
        return None
    return kit_identity_key(title, mark)


def rebuild_pipeline(
    database: CatalogDatabase,
    *,
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    kit_keys: Collection[tuple[str, str]] | None = None,
    rd_root: str | Path | None = None,
) -> None:
    """Rebuild ``kit_package`` / ``kit_cycle`` / ``kit_pipeline`` / heatmap.

    Does not delete ``kit_liquidity_review``, ``kit_working_flag``,
    ``kit_annulled_flag``, or ``file_mtime_override``. Google and issuance
    rows must already be in SQLite (typically via
    :func:`ingest_google_snapshot`).
    After :meth:`CatalogDatabase.replace_kit_derived`, calls
    :func:`rebuild_revision_matrix`.

    Args:
        database: Initialized catalog database.
        records: Present catalog files from the last scan.
        detected_current_ids: Overlay-current RD file ids.
        kit_keys: When set, derive and persist only these identities.
            Other kits' derived rows stay as they are. Empty is treated as
            a full rebuild. Overlay must already be rebuilt from all
            present RD files (store_scan).
        rd_root: RD source root. When set, non-canonical RD files are
            excluded via :func:`record_has_canonical_layout`. SQ and ROBOT
            rows are kept. Empty or ``None`` skips the filter.
    """

    scoped = _normalize_kit_keys(kit_keys)
    with perf_span("pipeline.rebuild_pipeline", kits=len(scoped) if scoped is not None else "all"):
        records = _records_in_contour(records, rd_root)
        google_kits = database.list_google_kits()
        google_by_key = {
            kit_identity_key(kit.title, kit.mark): kit for kit in google_kits
        }
        sends_with_ids = _effective_issuance_sends(database)
        work_records = records
        work_sends = sends_with_ids
        if scoped is not None:
            work_records = [
                record
                for record in records
                if (key := _data_kit_key(record)) is not None and key in scoped
            ]
            work_sends = [
                (send_id, send)
                for send_id, send in sends_with_ids
                if kit_identity_key(send.title, send.mark) in scoped
            ]
        database.apply_void_annulled_flags_from_records(work_records)
        sends_by_key: dict[tuple[str, str], list[tuple[int | None, IssuanceKit]]] = {}
        for send_id, send in work_sends:
            sends_by_key.setdefault(kit_identity_key(send.title, send.mark), []).append(
                (send_id, send)
            )

        packages = _build_source_packages(work_records, detected_current_ids)
        packages.extend(_build_grey_packages(packages, work_sends))
        packages = _assign_package_ids(packages)

        events_by_kit = database.list_google_events_by_kit()
        flags_by_key: dict[tuple[str, str], tuple[KitWorkingFlagRow, ...]] = {}
        for flag in database.list_working_flags():
            key = kit_identity_key(flag.title, flag.mark)
            flags_by_key[key] = flags_by_key.get(key, ()) + (flag,)
        annulled_by_key: dict[tuple[str, str], tuple[KitAnnulledFlagRow, ...]] = {}
        for flag in database.list_annulled_flags():
            key = kit_identity_key(flag.title, flag.mark)
            annulled_by_key[key] = annulled_by_key.get(key, ()) + (flag,)
        records_by_kit = index_records_by_kit(work_records)
        exclusions_by_key: dict[tuple[str, str], _WorkingExclusion] = {}
        preview_keys = (
            set(sends_by_key)
            | set(records_by_kit)
            | {
                kit_identity_key(pkg.title, pkg.mark)
                for pkg in packages
                if pkg.source == "rd"
            }
        )
        packages_by_preview: dict[tuple[str, str], list[KitPackageRow]] = {}
        for pkg in packages:
            packages_by_preview.setdefault(
                kit_identity_key(pkg.title, pkg.mark), []
            ).append(pkg)
        for key in preview_keys:
            exclusions_by_key[key] = _working_exclusion_for_kit(
                records=records_by_kit.get(key, ()),
                detected_current_ids=detected_current_ids,
                sends=sends_by_key.get(key, ()),
                events=events_by_kit.get(key, ()),
                packages=packages_by_preview.get(key, ()),
                flags=flags_by_key.get(key, ()),
                annulled_flags=annulled_by_key.get(key, ()),
                kit_key=key,
            )
        packages = _mark_current_rd_packages(packages, exclusions_by_key)

        cycles: list[KitCycleRow] = []
        for send_id, send in work_sends:
            key = kit_identity_key(send.title, send.mark)
            cycles.append(
                match_send_cycle(
                    send_id,
                    send,
                    events_by_kit.get(key, ()),
                    packages,
                )
            )

        reviews = database.list_liquidity_reviews()
        reviews_by_key: dict[tuple[str, str], list[LiquidityReviewRow]] = {}
        for review in reviews:
            reviews_by_key.setdefault(
                kit_identity_key(review.title, review.mark), []
            ).append(review)

        display_names = _display_names(
            google_by_key, sends_by_key, packages, work_records
        )
        if scoped is not None:
            identities = set(scoped)
        else:
            identities = (
                set(google_by_key)
                | set(sends_by_key)
                | {
                    kit_identity_key(pkg.title, pkg.mark)
                    for pkg in packages
                    if pkg.source == "rd"
                }
            )
        records_by_kit = index_records_by_kit(work_records)
        pipelines: list[KitPipelineRow] = []
        for key in sorted(identities):
            title, mark = display_names.get(key, (key[0], key[1]))
            pipelines.append(
                derive_kit_pipeline(
                    title,
                    mark,
                    records=records_by_kit.get(key, ()),
                    detected_current_ids=detected_current_ids,
                    sends=sends_by_key.get(key, ()),
                    events=events_by_kit.get(key, ()),
                    packages=[
                        pkg
                        for pkg in packages
                        if kit_identity_key(pkg.title, pkg.mark) == key
                    ],
                    cycles=[
                        cycle
                        for cycle in cycles
                        if kit_identity_key(cycle.title, cycle.mark) == key
                    ],
                    reviews=reviews_by_key.get(key, ()),
                    manual_flags=flags_by_key.get(key, ()),
                    annulled_flags=annulled_by_key.get(key, ()),
                    working_exclusion=exclusions_by_key.get(key),
                )
            )

        database.replace_kit_derived(
            packages, cycles, pipelines, identities=scoped
        )
        rebuild_revision_matrix(
            database,
            records=records,
            detected_current_ids=detected_current_ids,
            kit_keys=scoped,
            rd_root=rd_root,
        )


def match_event_reason(send: IssuanceKit, event: KitEvent) -> str | None:
    """Return how one F event matches an issuance send, if at all.

    Args:
        send: One «Выдача РД ПД» row.
        event: One column-F event of the same kit.

    Returns:
        ``trm``, ``revision``, ``date``, or ``None``.
    """

    send_trm = (send.send_transmittal or "").strip()
    if send_trm:
        folded = send_trm.casefold()
        if any((token or "").casefold() == folded for token in event.transmittals):
            return "trm"
    event_rev = format_revision(event.revision, event.appendix)
    if event_rev and revision_texts_equivalent(send.revision_text, event_rev):
        return "revision"
    if event.date and (
        dates_within_days(event.date, send.send_date)
        or dates_within_days(event.date, send.incoming_control_date)
    ):
        return "date"
    return None


def match_send_package(
    send: IssuanceKit,
    packages: Sequence[KitPackageRow],
) -> KitPackageRow | None:
    """Return the RD (or grey) package for one send.

    Args:
        send: One issuance send.
        packages: Packages of all kits (filtered by identity).

    Returns:
        Best RD package with the same revision, else a grey stub, else ``None``.
    """

    key = kit_identity_key(send.title, send.mark)
    kit_packages = [
        pkg for pkg in packages if kit_identity_key(pkg.title, pkg.mark) == key
    ]
    rd_matches = [
        pkg
        for pkg in kit_packages
        if pkg.source == "rd"
        and not pkg.is_grey
        and revision_texts_equivalent(pkg.revision_text, send.revision_text)
    ]
    if rd_matches:
        return max(
            rd_matches,
            key=lambda pkg: (
                pkg.sequence if pkg.sequence is not None else -1,
                pkg.id or 0,
            ),
        )
    for pkg in kit_packages:
        if pkg.is_grey and revision_texts_equivalent(
            pkg.revision_text, send.revision_text
        ):
            return pkg
    return None


def match_send_cycle(
    send_id: int | None,
    send: IssuanceKit,
    events: Sequence[tuple[int, KitEvent]],
    packages: Sequence[KitPackageRow],
) -> KitCycleRow:
    """Build one ``kit_cycle`` linking a send to F events and a package.

    Args:
        send_id: SQLite ``issuance_send.id``, or ``None`` for a legalized
            orphan/manual synthetic.
        send: Parsed or synthetic send.
        events: ``(google_event.id, event)`` pairs in sheet order.
        packages: All package rows with client ids assigned.

    Returns:
        Cycle row whose ``package_id`` is a client :attr:`KitPackageRow.id`.
        Legalized synthetics keep ``match_reason="legalized"`` even when a
        package or F event also matches.
    """

    matched: list[tuple[int, KitEvent, str]] = []
    for event_id, event in events:
        reason = match_event_reason(send, event)
        if reason is not None:
            matched.append((event_id, event, reason))
    match_reason = ""
    for preferred in ("trm", "revision", "date"):
        if any(reason == preferred for _eid, _event, reason in matched):
            match_reason = preferred
            break
    if not match_reason and match_send_package(send, packages) is not None:
        match_reason = "revision"
    if send_id is None:
        match_reason = "legalized"

    tdo_event_id = _best_tdo_event_id(matched, events, send)
    code_event_id = _last_code_event_id(
        matched, events, send, tdo_event_id=tdo_event_id
    )
    package = match_send_package(send, packages)
    return KitCycleRow(
        title=send.title,
        mark=send.mark,
        revision_text=send.revision_text,
        match_reason=match_reason,
        send_id=send_id,
        package_id=package.id if package is not None else None,
        tdo_event_id=tdo_event_id,
        code_event_id=code_event_id,
    )


@dataclass(frozen=True, slots=True)
class _WorkingExclusion:
    """Working packages for one kit: folder identity, not whole filename rev."""

    revision_text: str
    official_revision: str
    folder_keys: tuple[str, ...]
    sequences: tuple[int, ...]
    auto_revision: str
    annulled_folder_keys: tuple[str, ...] = ()
    annulled_sequences: tuple[int, ...] = ()


def _folder_key(name: str) -> str:
    return str(name or "").strip().casefold()


def _package_folder_keys(package: KitPackageRow) -> tuple[str, ...]:
    names: list[str] = []
    key = _folder_key(package.transfer_name or "")
    if key:
        names.append(key)
    path = str(package.package_path or "").strip()
    if path:
        path_key = _folder_key(Path(path).name)
        if path_key and path_key not in names:
            names.append(path_key)
    return tuple(names)


def _void_package_folder_keys(packages: Sequence[KitPackageRow]) -> tuple[str, ...]:
    """Return issued-folder names that contain a Void (annulled) token."""

    keys: list[str] = []
    seen: set[str] = set()
    for package in packages:
        if package.source != "rd" or package.is_grey:
            continue
        raw_names = (
            package.transfer_name or "",
            Path(str(package.package_path or "")).name,
        )
        for raw in raw_names:
            if not transfer_name_is_void(raw):
                continue
            key = _folder_key(raw)
            if key and key not in seen:
                seen.add(key)
                keys.append(raw)
    return tuple(keys)


def _shared_transfer_sequences(
    packages: Sequence[KitPackageRow],
) -> frozenset[int]:
    """Return NN values that belong to more than one RD package."""

    counts: dict[int, int] = {}
    for package in packages:
        if package.source != "rd" or package.is_grey or package.sequence is None:
            continue
        sequence = int(package.sequence)
        counts[sequence] = counts.get(sequence, 0) + 1
    return frozenset(sequence for sequence, count in counts.items() if count > 1)


def _package_matches_flags(
    package: KitPackageRow,
    flags: Sequence[KitWorkingFlagRow | KitAnnulledFlagRow],
    *,
    shared_sequences: frozenset[int] = frozenset(),
) -> bool:
    if package.source != "rd" or package.is_grey:
        return False
    keys = set(_package_folder_keys(package))
    for flag in flags:
        flag_key = _folder_key(flag.transfer_name)
        if flag_key and flag_key in keys:
            return True
        if (
            package.sequence is not None
            and flag.sequence is not None
            and int(package.sequence) == int(flag.sequence)
            and int(package.sequence) not in shared_sequences
        ):
            return True
        if (
            not flag_key
            and flag.sequence is None
            and flag.revision_text
            and package.revision_text
            and revision_texts_equivalent(package.revision_text, flag.revision_text)
        ):
            return True
    return False


def _folder_names_match(
    names: Sequence[str],
    keys: Sequence[str],
) -> bool:
    folded = {_folder_key(name) for name in names if str(name or "").strip()}
    wanted = {_folder_key(name) for name in keys if str(name or "").strip()}
    return bool(folded and folded & wanted)


def _folder_names_are_working(
    names: Sequence[str],
    exclusion: _WorkingExclusion,
) -> bool:
    return _folder_names_match(names, exclusion.folder_keys)


def _sequence_is_working(
    sequence: int | None,
    exclusion: _WorkingExclusion,
) -> bool:
    """Match NN only when folder names are unknown (legacy pipeline rows)."""

    if exclusion.folder_keys or sequence is None:
        return False
    return sequence in exclusion.sequences


def _package_is_working(
    package: KitPackageRow,
    exclusion: _WorkingExclusion | None,
) -> bool:
    if exclusion is None or package.source != "rd" or package.is_grey:
        return False
    if _package_is_annulled(package, exclusion):
        return False
    if _folder_names_are_working(_package_folder_keys(package), exclusion):
        return True
    if _sequence_is_working(package.sequence, exclusion):
        return True
    if (
        exclusion.auto_revision
        and package.revision_text
        and revision_texts_equivalent(package.revision_text, exclusion.auto_revision)
    ):
        return True
    return False


def _record_folder_keys(record: FileRecord) -> tuple[str, ...]:
    names: list[str] = []
    transfer = str(record.data.get("transfer_name") or "").strip()
    if transfer:
        names.append(_folder_key(transfer))
    package = issued_package_dir(record.path)
    if package:
        path_key = _folder_key(Path(package).name)
        if path_key and path_key not in names:
            names.append(path_key)
    return tuple(names)


def _record_matches_folder_keys(
    record: FileRecord,
    keys: Collection[str],
) -> bool:
    wanted = {_folder_key(name) for name in keys if str(name or "").strip()}
    return bool(wanted and set(_record_folder_keys(record)) & wanted)


def _folder_identity_from_flags(
    flags: Sequence[KitWorkingFlagRow | KitAnnulledFlagRow],
    packages: Sequence[KitPackageRow],
    *,
    shared_sequences: frozenset[int],
) -> tuple[list[str], list[int], list[KitPackageRow]]:
    matched = [
        pkg
        for pkg in packages
        if _package_matches_flags(pkg, flags, shared_sequences=shared_sequences)
    ]
    folder_keys: list[str] = []
    sequences: list[int] = []
    for flag in flags:
        flag_key = _folder_key(flag.transfer_name)
        if flag_key and flag_key not in folder_keys:
            folder_keys.append(flag_key)
        if (
            flag.sequence is not None
            and int(flag.sequence) not in sequences
            and int(flag.sequence) not in shared_sequences
        ):
            sequences.append(int(flag.sequence))
    for pkg in matched:
        for name in _package_folder_keys(pkg):
            if name not in folder_keys:
                folder_keys.append(name)
        if (
            pkg.sequence is not None
            and pkg.sequence not in sequences
            and int(pkg.sequence) not in shared_sequences
        ):
            sequences.append(int(pkg.sequence))
    return folder_keys, sequences, matched


def _sequence_is_annulled(
    sequence: int | None,
    exclusion: _WorkingExclusion,
) -> bool:
    if exclusion.annulled_folder_keys or sequence is None:
        return False
    return sequence in exclusion.annulled_sequences


def _package_is_annulled(
    package: KitPackageRow,
    exclusion: _WorkingExclusion | None,
) -> bool:
    if exclusion is None or package.source != "rd" or package.is_grey:
        return False
    if _folder_names_match(
        _package_folder_keys(package), exclusion.annulled_folder_keys
    ):
        return True
    return _sequence_is_annulled(package.sequence, exclusion)


def _package_is_skipped_from_official(
    package: KitPackageRow,
    exclusion: _WorkingExclusion | None,
) -> bool:
    return _package_is_working(package, exclusion) or _package_is_annulled(
        package, exclusion
    )


def _record_group_is_working(
    files: Sequence[FileRecord],
    exclusion: _WorkingExclusion | None,
    *,
    working_revision: str = "",
) -> bool:
    if not files:
        return False
    if exclusion is not None and (
        exclusion.folder_keys or exclusion.sequences
    ):
        names: list[str] = []
        sequences: list[int] = []
        for record in files:
            transfer = str(record.data.get("transfer_name") or "")
            if transfer:
                names.append(transfer)
            package = issued_package_dir(record.path)
            if package:
                names.append(Path(package).name)
            sequence = _int_or_none(record.data.get("transfer_sequence"))
            if sequence is not None:
                sequences.append(sequence)
        if _folder_names_are_working(names, exclusion):
            return True
        return any(_sequence_is_working(sequence, exclusion) for sequence in sequences)
    highest = _highest_revision_text(files)
    return bool(
        working_revision
        and highest
        and revision_texts_equivalent(highest, working_revision)
    )


def _record_group_is_annulled(
    files: Sequence[FileRecord],
    exclusion: _WorkingExclusion | None,
) -> bool:
    if exclusion is None or not files:
        return False
    if not (exclusion.annulled_folder_keys or exclusion.annulled_sequences):
        return False
    names: list[str] = []
    sequences: list[int] = []
    for record in files:
        names.extend(_record_folder_keys(record))
        sequence = _int_or_none(record.data.get("transfer_sequence"))
        if sequence is not None:
            sequences.append(sequence)
    if _folder_names_match(names, exclusion.annulled_folder_keys):
        return True
    return any(_sequence_is_annulled(sequence, exclusion) for sequence in sequences)


def _record_group_is_skipped_from_official(
    files: Sequence[FileRecord],
    exclusion: _WorkingExclusion | None,
    *,
    working_revision: str = "",
) -> bool:
    return _record_group_is_working(
        files, exclusion, working_revision=working_revision
    ) or _record_group_is_annulled(files, exclusion)


def _exclusion_from_pipeline(row: KitPipelineRow) -> _WorkingExclusion:
    return _WorkingExclusion(
        revision_text=row.working_revision_text,
        official_revision=row.official_revision_text,
        folder_keys=row.working_transfer_names,
        sequences=row.working_sequences,
        auto_revision="",
        annulled_folder_keys=row.annulled_transfer_names,
        annulled_sequences=row.annulled_sequences,
    )


def _revision_is_fully_working(
    revision_text: str,
    packages: Sequence[KitPackageRow],
    pipeline: KitPipelineRow | None,
) -> bool:
    if pipeline is None or not pipeline.working_revision_text:
        return False
    if not revision_texts_equivalent(pipeline.working_revision_text, revision_text):
        return False
    matching = [
        pkg
        for pkg in packages
        if pkg.source == "rd"
        and not pkg.is_grey
        and _package_matches_revision(pkg, revision_text)
    ]
    if not matching:
        return True
    exclusion = _exclusion_from_pipeline(pipeline)
    return all(_package_is_working(pkg, exclusion) for pkg in matching)


def _revision_is_fully_annulled(
    revision_text: str,
    packages: Sequence[KitPackageRow],
    pipeline: KitPipelineRow | None,
) -> bool:
    if pipeline is None or not (
        pipeline.annulled_transfer_names or pipeline.annulled_sequences
    ):
        return False
    matching = [
        pkg
        for pkg in packages
        if pkg.source == "rd"
        and not pkg.is_grey
        and _package_matches_revision(pkg, revision_text)
    ]
    if not matching:
        return False
    exclusion = _exclusion_from_pipeline(pipeline)
    return all(_package_is_annulled(pkg, exclusion) for pkg in matching)


def _last_issued_revision(
    sends: Sequence[tuple[int | None, IssuanceKit]],
    events: Sequence[tuple[int, KitEvent]],
) -> str:
    ordered = _sorted_sends(sends)
    latest_send = ordered[-1][1] if ordered else None
    last_issued = latest_send.revision_text if latest_send is not None else ""
    if last_issued:
        return last_issued
    for _event_id, event in reversed(list(events)):
        text = format_revision(event.revision, event.appendix)
        if text:
            return text
    return ""


def _working_exclusion_for_kit(
    *,
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    sends: Sequence[tuple[int | None, IssuanceKit]],
    events: Sequence[tuple[int, KitEvent]],
    packages: Sequence[KitPackageRow],
    flags: Sequence[KitWorkingFlagRow],
    kit_key: tuple[str, str] | None,
    annulled_flags: Sequence[KitAnnulledFlagRow] = (),
) -> _WorkingExclusion:
    key = kit_key
    if key is None and records:
        key = _record_kit_key(records[0])
    ordered = _sorted_sends(sends)
    last_issued = _last_issued_revision(sends, events)
    rd_packages = [
        pkg for pkg in packages if pkg.source == "rd" and not pkg.is_grey
    ]
    shared_sequences = _shared_transfer_sequences(rd_packages)
    annulled_folder_keys, annulled_sequences, _annulled_packages = (
        _folder_identity_from_flags(
            annulled_flags,
            rd_packages,
            shared_sequences=shared_sequences,
        )
    )
    void_keys = _void_package_folder_keys(rd_packages)
    if void_keys:
        seen = {_folder_key(name) for name in annulled_folder_keys}
        for name in void_keys:
            key = _folder_key(name)
            if key and key not in seen:
                seen.add(key)
                annulled_folder_keys.append(name)
    annulled_key_set = set(annulled_folder_keys)
    remaining_packages = [
        pkg
        for pkg in rd_packages
        if not _folder_names_match(_package_folder_keys(pkg), annulled_folder_keys)
        and not (
            not annulled_folder_keys
            and pkg.sequence is not None
            and int(pkg.sequence) in set(annulled_sequences)
        )
    ]
    remaining_records = [
        record
        for record in records
        if not _record_matches_folder_keys(record, annulled_folder_keys)
        and not (
            not annulled_folder_keys
            and (_seq := _int_or_none(record.data.get("transfer_sequence")))
            is not None
            and _seq in set(annulled_sequences)
        )
    ]
    overlay_rev = (
        _overlay_current_revision_text(
            remaining_records, detected_current_ids, key
        )
        if key is not None
        else ""
    )
    if not overlay_rev and remaining_packages:
        overlay_rev = max(
            remaining_packages, key=_current_package_sort_key
        ).revision_text
    auto_working = (
        derive_working_revision_text(
            remaining_records,
            key,
            last_issued,
            overlay_revision=overlay_rev,
        )
        if key is not None
        else ""
    )
    working_packages = [
        pkg
        for pkg in remaining_packages
        if _package_matches_flags(
            pkg, flags, shared_sequences=shared_sequences
        )
        or (
            auto_working
            and pkg.revision_text
            and revision_texts_equivalent(pkg.revision_text, auto_working)
        )
    ]
    folder_keys: list[str] = []
    sequences: list[int] = []
    for flag in flags:
        flag_key = _folder_key(flag.transfer_name)
        if flag_key and flag_key in annulled_key_set:
            continue
        if flag_key and flag_key not in folder_keys:
            folder_keys.append(flag_key)
        if (
            flag.sequence is not None
            and flag.sequence not in sequences
            and int(flag.sequence) not in shared_sequences
            and int(flag.sequence) not in set(annulled_sequences)
        ):
            sequences.append(int(flag.sequence))
    for pkg in working_packages:
        for name in _package_folder_keys(pkg):
            if name in annulled_key_set:
                continue
            if name not in folder_keys:
                folder_keys.append(name)
        if (
            pkg.sequence is not None
            and pkg.sequence not in sequences
            and int(pkg.sequence) not in shared_sequences
        ):
            sequences.append(int(pkg.sequence))
    working = ""
    if working_packages:
        working = max(
            working_packages,
            key=lambda pkg: _text_rank(pkg.revision_text),
        ).revision_text
        if not working and flags:
            working = flags[0].revision_text
        if not working:
            working = auto_working
    elif flags and not annulled_flags:
        working = max(
            (flag.revision_text for flag in flags if flag.revision_text),
            key=_text_rank,
            default="",
        )
    elif flags:
        remaining_flag_revs = [
            flag.revision_text
            for flag in flags
            if flag.revision_text
            and _folder_key(flag.transfer_name) not in annulled_key_set
        ]
        working = max(remaining_flag_revs, key=_text_rank, default="")
    elif auto_working:
        working = auto_working
    elif overlay_rev and last_issued and _text_rank(overlay_rev) > _text_rank(last_issued):
        working = overlay_rev
    exclusion = _WorkingExclusion(
        revision_text=working,
        official_revision="",
        folder_keys=tuple(folder_keys),
        sequences=tuple(sequences),
        auto_revision=auto_working,
        annulled_folder_keys=tuple(annulled_folder_keys),
        annulled_sequences=tuple(annulled_sequences),
    )
    non_skipped = [
        pkg
        for pkg in rd_packages
        if not _package_is_skipped_from_official(pkg, exclusion)
    ]
    if working:
        if last_issued and any(
            pkg.revision_text
            and revision_texts_equivalent(pkg.revision_text, last_issued)
            for pkg in non_skipped
        ):
            official = last_issued
        else:
            official = _last_issued_not_working(ordered, working, last_issued)
    elif annulled_folder_keys or annulled_sequences:
        if last_issued and any(
            pkg.revision_text
            and revision_texts_equivalent(pkg.revision_text, last_issued)
            for pkg in non_skipped
        ):
            official = last_issued
        elif overlay_rev and any(
            pkg.revision_text
            and revision_texts_equivalent(pkg.revision_text, overlay_rev)
            for pkg in non_skipped
        ):
            official = overlay_rev
        elif non_skipped:
            official = max(
                non_skipped, key=_current_package_sort_key
            ).revision_text
        else:
            official = ""
    else:
        official = overlay_rev or last_issued
    return _WorkingExclusion(
        revision_text=working,
        official_revision=official,
        folder_keys=tuple(folder_keys),
        sequences=tuple(sequences),
        auto_revision=auto_working,
        annulled_folder_keys=tuple(annulled_folder_keys),
        annulled_sequences=tuple(annulled_sequences),
    )


def compute_working_and_official_revisions(
    *,
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    sends: Sequence[tuple[int | None, IssuanceKit]],
    events: Sequence[tuple[int, KitEvent]],
    manual_revisions: Sequence[str] = (),
    manual_flags: Sequence[KitWorkingFlagRow] = (),
    packages: Sequence[KitPackageRow] = (),
    kit_key: tuple[str, str] | None = None,
    annulled_flags: Sequence[KitAnnulledFlagRow] = (),
) -> tuple[str, str]:
    """Return ``(working_revision_text, official_revision_text)``.

    Working is overlay-current (newest NN) strictly above the last
    effective send/F, or a manual flag on an issued folder. Official is the
    last issued revision that still exists on a non-working package.

    Args:
        records: Kit file records.
        detected_current_ids: Overlay-current RD file ids.
        sends: Ordered ``(id, send)`` pairs.
        events: Google F events.
        manual_revisions: Legacy whole-revision marks (no folder).
        manual_flags: User-marked working folders.
        packages: Kit packages used to keep a sibling of the same filename
            rev official when only one NN is flagged.
        kit_key: Identity used to read overlay-current files.
        annulled_flags: User-marked annulled folders (not auto-working).

    Returns:
        Working revision (empty if none) and official revision.
    """

    flags = tuple(manual_flags)
    if not flags and manual_revisions:
        flags = tuple(
            KitWorkingFlagRow(
                title="",
                mark="",
                revision_text=str(text).strip(),
            )
            for text in manual_revisions
            if str(text).strip()
        )
    exclusion = _working_exclusion_for_kit(
        records=records,
        detected_current_ids=detected_current_ids,
        sends=sends,
        events=events,
        packages=packages,
        flags=flags,
        kit_key=kit_key,
        annulled_flags=annulled_flags,
    )
    return exclusion.revision_text, exclusion.official_revision


def _last_issued_not_working(
    ordered_sends: Sequence[tuple[int | None, IssuanceKit]],
    working_revision: str,
    last_issued: str,
) -> str:
    for _send_id, send in reversed(list(ordered_sends)):
        text = send.revision_text or ""
        if text and not revision_texts_equivalent(text, working_revision):
            return text
    if last_issued and not revision_texts_equivalent(last_issued, working_revision):
        return last_issued
    return ""


def derive_kit_pipeline(
    title: str,
    mark: str,
    *,
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    sends: Sequence[tuple[int | None, IssuanceKit]],
    events: Sequence[tuple[int, KitEvent]],
    packages: Sequence[KitPackageRow],
    cycles: Sequence[KitCycleRow],
    reviews: Sequence[LiquidityReviewRow],
    manual_revisions: Sequence[str] = (),
    manual_flags: Sequence[KitWorkingFlagRow] = (),
    annulled_flags: Sequence[KitAnnulledFlagRow] = (),
    working_exclusion: _WorkingExclusion | None = None,
) -> KitPipelineRow:
    """Derive current pipeline status, code origin, working rev, liquidity.

    Official revision is the last issued cycle that still lives on a
    non-working package. Overlay-current files that are strictly ahead of
    that cycle, or a user-marked issued folder, are working and do not
    inherit TDO/A from the previous send.

    Args:
        title: Display title.
        mark: Display mark.
        records: Scan file records.
        detected_current_ids: Overlay-current RD file ids.
        sends: ``(issuance_send.id or None, send)`` pairs for this kit.
        events: ``(google_event.id, event)`` pairs for this kit.
        packages: Packages for this kit.
        cycles: Cycles for this kit.
        reviews: Liquidity decisions for this kit.
        manual_revisions: Legacy whole-revision marks.
        manual_flags: User-marked working folders.
        annulled_flags: User-marked annulled folders.
        working_exclusion: Precomputed working folders when rebuild already
            derived them for ``is_current``.

    Returns:
        One ``kit_pipeline`` row.
    """

    key = kit_identity_key(title, mark)
    ordered_sends = _sorted_sends(sends)
    flags = tuple(manual_flags)
    if not flags and manual_revisions:
        flags = tuple(
            KitWorkingFlagRow(
                title=title,
                mark=mark,
                revision_text=str(text).strip(),
            )
            for text in manual_revisions
            if str(text).strip()
        )
    exclusion = working_exclusion or _working_exclusion_for_kit(
        records=records,
        detected_current_ids=detected_current_ids,
        sends=ordered_sends,
        events=events,
        packages=packages,
        flags=flags,
        kit_key=key,
        annulled_flags=annulled_flags,
    )
    working_revision = exclusion.revision_text
    official_revision = exclusion.official_revision

    last_cycle, last_send = _last_official_cycle(
        official_revision, ordered_sends, cycles
    )
    send_f_revision = _official_send_f_revision(last_send, events)
    official_has_send = bool(
        official_revision
        and any(
            revision_texts_equivalent(send.revision_text, official_revision)
            for _send_id, send in ordered_sends
        )
    )

    status = _review_status(
        official_revision=official_revision,
        last_send=last_send,
        last_cycle=last_cycle,
        events=events,
        has_official_send=official_has_send,
    )
    (
        code,
        code_revision_text,
        code_date,
        code_stale,
        code_origin,
    ) = _derive_approval_fields(
        events=events,
        official_revision=official_revision,
        send_f_revision=send_f_revision,
        last_send=last_send,
        last_cycle=last_cycle,
        ordered_sends=ordered_sends,
        packages=packages,
    )
    review_as_build = _review_as_build_for_cycle(last_cycle, packages)
    working_as_build = _working_as_build(packages, exclusion, working_revision)
    tdo_date = ""
    if status is KitPipelineStatus.TDO_REVIEW:
        tdo_date = _tdo_passed_date_for_cycle(
            official_revision=official_revision,
            last_send=last_send,
            last_cycle=last_cycle,
            events=events,
        )

    suspicious = derive_liquidity_pending(
        official_revision=official_revision,
        send_f_revision=send_f_revision,
        last_send=last_send,
        events=events,
        packages=packages,
        reviews=reviews,
    )
    return KitPipelineRow(
        title=title,
        mark=mark,
        status=status.value,
        code=code,
        code_origin=code_origin,
        working_revision_text=working_revision,
        official_revision_text=official_revision,
        suspicious=suspicious,
        algorithm_version=PIPELINE_STATUS_ALGORITHM_VERSION,
        code_stale=code_stale,
        code_revision_text=code_revision_text,
        code_date=code_date,
        tdo_date=tdo_date,
        review_as_build=review_as_build,
        working_as_build=working_as_build,
        working_transfer_names=exclusion.folder_keys,
        working_sequences=exclusion.sequences,
        annulled_transfer_names=exclusion.annulled_folder_keys,
        annulled_sequences=exclusion.annulled_sequences,
    )


def derive_code_origin(
    code: str,
    code_event: KitEvent,
    code_revision_text: str,
    sends: Sequence[tuple[int | None, IssuanceKit]],
    packages: Sequence[KitPackageRow],
    *,
    last_send: IssuanceKit | None,
) -> str:
    """Decide whether a B/C code came from the customer or from PI.

    Args:
        code: ``A`` / ``B`` / ``C``.
        code_event: Last code event on the official revision.
        code_revision_text: Revision the code applies to.
        sends: Kit sends in chronological order.
        packages: Kit packages.
        last_send: Send of the last official cycle, if any.

    Returns:
        ``customer``, ``pi``, or ``unknown``.
    """

    if code.upper() == "A":
        return "customer"
    code_date = parse_ddmmyyyy(code_event.date)
    later_sends: list[IssuanceKit] = []
    for _send_id, send in _sorted_sends(sends):
        if _send_is_after_code(send, code_date, last_send):
            later_sends.append(send)

    code_rank = _text_rank(code_revision_text)
    if any(_text_rank(send.revision_text) > code_rank for send in later_sends):
        return "customer"
    if later_sends and revision_texts_equivalent(
        later_sends[0].revision_text, code_revision_text
    ):
        return "pi"
    if _later_higher_rd_package(packages, code_rank, code_date, code_event):
        return "customer"
    return "unknown"


def derive_working_revision_text(
    records: Sequence[FileRecord],
    key: tuple[str, str],
    send_f_revision: str,
    *,
    overlay_revision: str = "",
) -> str:
    """Return overlay-current rev when it is strictly ahead of the last send/F.

    Overlay-current (newest NN) is the head. A leftover file in an earlier
    folder whose filename ranks higher (``01`` vs later ``0-AN01``) is not
    working. When overlay text is empty, fall back to the disk-wide head.

    Args:
        records: Scan file records.
        key: Kit identity.
        send_f_revision: Last official send or F revision text.
        overlay_revision: Filename rev of overlay-current RD files.

    Returns:
        Working filename revision, or empty string.
    """

    disk_text = str(overlay_revision or "").strip()
    if not disk_text:
        disk_text = _highest_revision_text(
            [
                record
                for record in records
                if record.present
                and record.source is SourceKind.RD
                and _record_kit_key(record) == key
            ]
        )
    if not disk_text or not send_f_revision:
        return ""
    if _text_rank(disk_text) > _text_rank(send_f_revision):
        return disk_text
    return ""


def working_revision_for_display(working: str, official: str) -> str:
    """Return working rev only when it strictly outranks official RD.

    Args:
        working: Stored ``kit_pipeline.working_revision_text``.
        official: Official filename rev (``official_revision_text`` / disk).

    Returns:
        ``working``, or empty when missing, equal, or lower than official.
    """

    working_text = str(working or "").strip()
    if not working_text:
        return ""
    official_text = str(official or "").strip()
    if not official_text:
        return working_text
    if revision_texts_equivalent(working_text, official_text):
        return ""
    if _text_rank(working_text) <= _text_rank(official_text):
        return ""
    return working_text


def derive_liquidity_pending(
    *,
    official_revision: str,
    send_f_revision: str,
    last_send: IssuanceKit | None,
    events: Sequence[tuple[int, KitEvent]],
    packages: Sequence[KitPackageRow],
    reviews: Sequence[LiquidityReviewRow],
) -> bool:
    """Return whether any official-revision RD package is liquidity-pending.

    Grey packages and a disk revision strictly higher than the last send are
    not liquidity. ``confirmed_illiquid`` clears pending; ``confirmed_ok``
    clears it only while ``evidence_mtime_ns`` covers ``max_mtime_ns``.

    Args:
        official_revision: Overlay or issuance official revision text.
        send_f_revision: Last official send/F revision.
        last_send: Last cycle send for the official revision, if any.
        events: Kit F events with ids.
        packages: Kit packages.
        reviews: User liquidity decisions.

    Returns:
        ``True`` when ``kit_pipeline.suspicious`` should be set.
    """

    anchor_month = _liquidity_anchor_month(official_revision, last_send, events)
    if anchor_month is None:
        return False
    review_map = {
        _review_identity(review): review
        for review in reviews
    }
    pending = False
    for pkg in packages:
        if pkg.source != "rd" or pkg.is_grey:
            continue
        if not revision_texts_equivalent(pkg.revision_text, official_revision):
            continue
        if send_f_revision and _text_rank(pkg.revision_text) > _text_rank(
            send_f_revision
        ):
            continue
        if pkg.max_mtime_ns is None:
            continue
        package_month = year_month_from_mtime_ns(pkg.max_mtime_ns)
        if package_month is None or package_month <= anchor_month:
            continue
        review = review_map.get(
            (
                pkg.title.casefold(),
                pkg.mark.casefold(),
                (pkg.transfer_name or "").casefold(),
                (pkg.revision_text or "").casefold(),
            )
        )
        if review is None:
            pending = True
            continue
        if review.decision == "confirmed_illiquid":
            continue
        if review.decision == "confirmed_ok":
            evidence = review.evidence_mtime_ns
            if evidence is not None and pkg.max_mtime_ns > evidence:
                pending = True
            continue
        pending = True
    return pending


def revision_texts_equivalent(left: str, right: str) -> bool:
    """Compare two filename/sheet revision strings.

    Args:
        left: First revision text.
        right: Second revision text.

    Returns:
        ``True`` when both sides are non-empty and equivalent.
    """

    left_text = (left or "").strip()
    right_text = (right or "").strip()
    if not left_text or not right_text:
        return False
    if left_text.casefold() == right_text.casefold():
        return True
    left_rev, left_app = parse_sheet_revision(left_text)
    right_rev, right_app = parse_sheet_revision(right_text)
    return revisions_equivalent(left_rev, left_app, right_rev, right_app)


def _package_mto_revision(package: KitPackageRow) -> str:
    """Return the revision text that identifies the MTO of this package."""

    return (package.mto_revision_text or package.revision_text or "").strip()


def _record_mtime_ns(record: FileRecord) -> int | None:
    """Return the stored file mtime in nanoseconds, or ``None`` when absent."""

    return _int_or_none(record.data.get("mtime_ns"))


def _decode_problem_kinds(cell: KitRevisionRow) -> tuple[str, ...]:
    """Return stringified problem-kind tokens from a heatmap cell snapshot."""

    try:
        parsed = json.loads(cell.problem_kinds_json or "[]")
    except (TypeError, ValueError):
        return ()
    if not isinstance(parsed, list):
        return ()
    return tuple(str(item) for item in parsed)


def _decode_package_ids(cell: KitRevisionRow) -> tuple[int, ...]:
    """Return package row ids stored on a heatmap cell snapshot."""

    try:
        parsed = json.loads(cell.package_ids_json or "[]")
    except (TypeError, ValueError):
        return ()
    if not isinstance(parsed, list):
        return ()
    ids: list[int] = []
    for item in parsed:
        try:
            ids.append(int(item))
        except (TypeError, ValueError):
            continue
    return tuple(ids)


def _hint_for_package(
    package: KitPackageRow,
    *,
    pipeline: KitPipelineRow | None,
    cycles: Sequence[KitCycleRow],
    events: Sequence[tuple[int, KitEvent]],
    sends_by_id: dict[int, IssuanceKit],
    cell: KitRevisionRow | None = None,
    flags: Sequence[KitWorkingFlagRow] = (),
) -> FolderTreeHint:
    """Build tree extras for one RD package folder."""

    working_text = (
        pipeline.working_revision_text if pipeline is not None else ""
    )
    exclusion = (
        _exclusion_from_pipeline(pipeline) if pipeline is not None else None
    )
    is_working = False
    if pipeline is not None and (
        pipeline.working_transfer_names or pipeline.working_sequences
    ):
        is_working = _package_is_working(package, exclusion)
    elif (
        working_text
        and package.revision_text
        and revision_texts_equivalent(package.revision_text, working_text)
    ):
        is_working = True
    is_annulled = _package_is_annulled(package, exclusion)
    if is_annulled:
        is_working = False
    working_origin = ""
    if is_working:
        working_origin = (
            "manual" if _package_matches_flags(package, flags) else "auto"
        )
    mto_revision_text = _package_mto_revision(package)
    has_mto_file = bool((package.mto_revision_text or "").strip())
    mto_status = cell.pipeline_status if cell is not None else ""
    mto_letters = cell.letters if cell is not None else ""
    is_current_mto = bool(cell.is_current) if cell is not None else False
    is_current_ifc = bool(cell.is_current_ifc) if cell is not None else False
    problem_kinds = _decode_problem_kinds(cell) if cell is not None else ()
    pkg_cycles = _cycles_for_package(package, cycles)
    if not pkg_cycles:
        return FolderTreeHint(
            is_working=is_working,
            is_annulled=is_annulled,
            working_origin=working_origin,
            mto_status=mto_status,
            mto_letters=mto_letters,
            mto_revision_text=mto_revision_text,
            has_mto_file=has_mto_file,
            is_current_mto=is_current_mto,
            is_current_ifc=is_current_ifc,
            problem_kinds=problem_kinds,
        )
    last_cycle = max(
        pkg_cycles,
        key=lambda cycle: _cycle_sort_key(cycle, sends_by_id),
    )
    last_send = (
        sends_by_id.get(last_cycle.send_id)
        if last_cycle.send_id is not None
        else None
    )
    official_revision = package.revision_text or last_cycle.revision_text
    status = _review_status(
        official_revision=official_revision,
        last_send=last_send,
        last_cycle=last_cycle,
        events=events,
        has_official_send=last_send is not None,
    )
    events_by_id = {event_id: event for event_id, event in events}
    f_parts: list[str] = []
    for event_id in (last_cycle.tdo_event_id, last_cycle.code_event_id):
        event = events_by_id.get(event_id) if event_id else None
        if event is None:
            continue
        label = (event.stage_label or event.stage or "").strip()
        if label and label not in f_parts:
            f_parts.append(label)
    return FolderTreeHint(
        review_status=pipeline_status_short_label(status.value),
        review_full=pipeline_status_label(status.value),
        match_reason=last_cycle.match_reason or "",
        f_label=" · ".join(f_parts),
        is_working=is_working,
        is_annulled=is_annulled,
        working_origin=working_origin,
        mto_status=mto_status,
        mto_letters=mto_letters,
        mto_revision_text=mto_revision_text,
        has_mto_file=has_mto_file,
        is_current_mto=is_current_mto,
        is_current_ifc=is_current_ifc,
        problem_kinds=problem_kinds,
    )


def _cycles_for_package(
    package: KitPackageRow,
    cycles: Sequence[KitCycleRow],
) -> list[KitCycleRow]:
    matched = [
        cycle
        for cycle in cycles
        if package.id is not None and cycle.package_id == package.id
    ]
    if matched:
        return matched
    if not package.revision_text:
        return []
    return [
        cycle
        for cycle in cycles
        if cycle.package_id is None
        and revision_texts_equivalent(cycle.revision_text, package.revision_text)
    ]


def _cycle_sort_key(
    cycle: KitCycleRow,
    sends_by_id: dict[int, IssuanceKit],
) -> tuple[str, int, int]:
    send = sends_by_id.get(cycle.send_id) if cycle.send_id is not None else None
    if send is None:
        return ("", 0, cycle.id or 0)
    sortable, row_index = _send_sort_key(send)
    return (sortable, row_index, cycle.id or 0)


def parse_ddmmyyyy(text: str | None) -> date | None:
    """Parse a ``DD.MM.YYYY`` calendar date.

    Args:
        text: Sheet date text.

    Returns:
        A :class:`~datetime.date`, or ``None``.
    """

    raw = (text or "").strip()
    match = _DATE_RE.fullmatch(raw)
    if not match:
        return None
    day, month, year = (int(match.group(1)), int(match.group(2)), int(match.group(3)))
    try:
        return date(year, month, day)
    except ValueError:
        return None


def dates_within_days(left: str | None, right: str | None, days: int = 2) -> bool:
    """Return whether two ``DD.MM.YYYY`` dates are within ``days``.

    Args:
        left: First date.
        right: Second date.
        days: Inclusive calendar-day window.

    Returns:
        ``True`` when both parse and the absolute delta is at most ``days``.
    """

    first = parse_ddmmyyyy(left)
    second = parse_ddmmyyyy(right)
    if first is None or second is None:
        return False
    return abs((first - second).days) <= days


def year_month_from_mtime_ns(mtime_ns: int) -> tuple[int, int] | None:
    """Return local ``(year, month)`` for a nanosecond timestamp.

    Args:
        mtime_ns: File ``st_mtime_ns``.

    Returns:
        Calendar year and month, or ``None`` when the timestamp is invalid.
    """

    try:
        stamp = datetime.fromtimestamp(int(mtime_ns) / _NS)
    except (OSError, OverflowError, ValueError):
        return None
    return stamp.year, stamp.month


def _record_kit_key(record: FileRecord) -> tuple[str, str] | None:
    if str(record.data.get("parse_status") or "") != "parsed":
        return None
    title = str(record.data.get("title") or "").strip()
    mark = str(record.data.get("mark") or "").strip()
    if not _TITLE_RE.fullmatch(title) or not is_rd_kit_mark(mark):
        return None
    return kit_identity_key(title, mark)


def index_records_by_kit(
    records: Sequence[FileRecord],
) -> dict[tuple[str, str], list[FileRecord]]:
    """Group catalog files by ``kit_identity_key(title, mark)``.

    Uses :func:`_record_kit_key` (parsed 4-digit title + Latin AGCC mark).
    Records that do not yield a key are omitted. Order inside each group
    follows ``records``.

    Args:
        records: Catalog file records, typically from the last scan.

    Returns:
        Mapping of kit identity to the records that belong to that kit.
    """

    grouped: dict[tuple[str, str], list[FileRecord]] = {}
    for record in records:
        key = _record_kit_key(record)
        if key is None:
            continue
        grouped.setdefault(key, []).append(record)
    return grouped


def _record_title_mark(record: FileRecord) -> tuple[str, str] | None:
    key = _record_kit_key(record)
    if key is None:
        return None
    title = str(record.data.get("title") or "").strip()
    mark = str(record.data.get("mark") or "").strip()
    return title, mark


def _file_revision(record: FileRecord) -> tuple[str | None, str | None]:
    revision = record.data.get("revision")
    appendix = record.data.get("appendix")
    return (
        str(revision) if revision not in (None, "") else None,
        str(appendix) if appendix not in (None, "") else None,
    )


def _present_rd_ids_for_revision(
    records: Sequence[FileRecord],
    key: tuple[str, str],
    target: str,
    working: str = "",
    *,
    exclusion: _WorkingExclusion | None = None,
) -> set[int]:
    """Return present RD file ids in the issued package of ``target``.

    Groups files by ``issued_package_dir``. A working-head **folder** is
    skipped even when leftover files still carry the official filename
    revision. A sibling package with the same NN is not skipped. Prefer
    the package whose highest rev matches ``target``; if that issued
    folder is missing (send ``01-AN01``, disk still ``01``), fall back
    to the newest non-working package so «РД · рев.» and open-folder
    stay on the same NN. **All** present RD files in the chosen folder
    are returned.

    Args:
        records: Catalog file records.
        key: Kit identity.
        target: Official filename revision.
        working: Working filename revision to exclude by package.

    Returns:
        File ids of the official (or fallback non-working) package.
    """

    return {
        record.id
        for record in _official_package_files(
            records, key, target, working, exclusion=exclusion
        )
    }


def _pipeline_skips_official(row: KitPipelineRow) -> bool:
    return bool(
        row.working_revision_text
        or row.working_transfer_names
        or row.working_sequences
        or row.annulled_transfer_names
        or row.annulled_sequences
    )


def official_detected_current_ids(
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    pipelines: Sequence[KitPipelineRow],
) -> set[int]:
    """Return overlay-current ids with working and annulled kits replaced.

    A kit with ``working_revision_text`` or annulled folders drops **all**
    overlay-current files of that kit (not only those whose filename
    equals the working rev). Leftover overlay files in a skipped package
    — empty revision, delta MTO, editables, leftover copies still named
    as the official rev — must not keep ``paths[0]`` on that folder.
    Present RD files of the issued package whose highest filename rev is
    ``official_revision_text`` are promoted instead. If that package is
    absent, the newest non-working and non-annulled package is used
    (disk ``01`` when the send is ``01-AN01``). Working/annulled is a
    **folder**, not every package with the same NN. Kits with **only**
    annulled folders (no working head) still go through this promotion
    so overlay-current on an annulled newest NN is not official.

    Args:
        records: Catalog file records.
        detected_current_ids: Full overlay-current ids (includes working).
        pipelines: Derived kit pipelines with working/official revisions.

    Returns:
        File ids that kits / export / MTO compare treat as current.
    """

    with perf_span(
        "pipeline.official_detected_current_ids",
        n=len(detected_current_ids),
    ):
        skip_by_kit = {
            kit_identity_key(row.title, row.mark): row
            for row in pipelines
            if _pipeline_skips_official(row)
        }
        skip_kits = set(skip_by_kit)
        official_by_kit = {
            kit_identity_key(row.title, row.mark): row.official_revision_text
            for row in pipelines
            if row.official_revision_text
        }
        by_id = {record.id: record for record in records}
        by_kit = index_records_by_kit(records) if skip_kits else {}
        official: set[int] = set()
        for file_id in detected_current_ids:
            record = by_id.get(file_id)
            if record is None or not record.present:
                continue
            key = _record_kit_key(record)
            if key is None:
                official.add(file_id)
                continue
            if key in skip_kits:
                continue
            official.add(file_id)
        for key in skip_kits:
            pipeline = skip_by_kit[key]
            official |= _present_rd_ids_for_revision(
                by_kit.get(key, ()),
                key,
                official_by_kit.get(key, ""),
                pipeline.working_revision_text,
                exclusion=_exclusion_from_pipeline(pipeline),
            )
        return official


def patch_official_detected_current_ids(
    current: set[int],
    *,
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    pipelines: Sequence[KitPipelineRow],
    kit_keys: Collection[tuple[str, str]],
) -> set[int]:
    """Replace official overlay ids for ``kit_keys``; keep other kits.

    Args:
        current: Previously computed official ids.
        records: Contour catalog rows (same set as a full recompute).
        detected_current_ids: Full overlay-current ids (includes working).
        pipelines: Pipelines after the scoped rebuild (full in-memory map
            values are fine; only ``kit_keys`` are re-derived).
        kit_keys: Identities whose official files changed.

    Returns:
        Updated official id set.
    """

    wanted = {kit_identity_key(title, mark) for title, mark in kit_keys}
    if not wanted:
        return set(current)
    by_id = {record.id: record for record in records}
    kept = set()
    for file_id in current:
        record = by_id.get(file_id)
        if record is None:
            kept.add(file_id)
            continue
        key = _record_kit_key(record)
        if key is not None and key in wanted:
            continue
        kept.add(file_id)
    kit_records = [
        record
        for record in records
        if (key := _record_kit_key(record)) is not None and key in wanted
    ]
    kit_detected = {
        file_id
        for file_id in detected_current_ids
        if (record := by_id.get(file_id)) is not None
        and (key := _record_kit_key(record)) is not None
        and key in wanted
    }
    kit_pipelines = [
        row
        for row in pipelines
        if kit_identity_key(row.title, row.mark) in wanted
    ]
    return kept | official_detected_current_ids(
        kit_records, kit_detected, kit_pipelines
    )


def official_rd_mto_overlay(
    database: CatalogDatabase,
    *,
    records: Sequence[FileRecord] | None = None,
) -> OverlayResult:
    """Return the RD MTO overlay used for robot compare (skipped files out).

    Args:
        database: Initialized catalog database.
        records: Optional already-loaded catalog rows. When omitted,
            ``list_files`` is queried (child-process compare).

    Returns:
        Overlay whose ``current`` map has no working-kit or annulled-kit
        MTO files. Kits with a working or annulled head keep only the
        official issued MTO.
    """

    with perf_span("pipeline.official_rd_mto_overlay"):
        return _official_rd_mto_overlay(database, records=records)


def _official_rd_mto_overlay(
    database: CatalogDatabase,
    *,
    records: Sequence[FileRecord] | None,
) -> OverlayResult:
    overlay = database.reconstruct_current_rd_mto_overlay()
    pipelines = database.list_kit_pipelines()
    pipeline_by_kit = {
        kit_identity_key(row.title, row.mark): row
        for row in pipelines
        if _pipeline_skips_official(row)
    }
    official_by_kit = {
        kit_identity_key(row.title, row.mark): row.official_revision_text
        for row in pipelines
        if row.official_revision_text
    }
    kept: dict[str, ParsedFile] = {}
    working_kits = set(pipeline_by_kit)
    for doc_key, parsed in overlay.current.items():
        key = kit_identity_key(parsed.title or "", parsed.mark or "")
        if key in working_kits:
            continue
        kept[doc_key] = parsed
    if working_kits:
        source = records if records is not None else database.list_files()
        present_rd = [
            record
            for record in source
            if record.present and record.source is SourceKind.RD
        ]
        by_kit = index_records_by_kit(present_rd)
        parsed_by_key = {
            parsed.path_key: parsed
            for parsed in database.list_present_rd_mto_files()
        }
        for key in working_kits:
            pipeline = pipeline_by_kit[key]
            package_files = _official_package_files(
                by_kit.get(key, ()),
                key,
                official_by_kit.get(key, ""),
                pipeline.working_revision_text,
                exclusion=_exclusion_from_pipeline(pipeline),
            )
            mto = [
                parsed_by_key[record.path_key]
                for record in package_files
                if record.path_key in parsed_by_key
            ]
            if not mto:
                continue
            best = max(mto, key=_newest_sort_key)
            doc = document_key_text(best)
            if doc:
                kept[doc] = best
    result = OverlayResult(file_kind=FileKind.MTO_XLSX)
    result.current = kept
    return result


def _discipline_is_od(record: FileRecord) -> bool:
    block = str(record.data.get("discipline_block") or "").casefold()
    return block.startswith("od")


def _file_kind(record: FileRecord) -> str:
    return str(record.data.get("file_kind") or "").casefold()


def _highest_revision_text(records: Sequence[FileRecord]) -> str:
    best_text = ""
    best_key: tuple[tuple[int, int, str], int] | None = None
    for record in records:
        revision, appendix = _file_revision(record)
        if not revision:
            continue
        rank = (revision_rank(revision, appendix), int(_discipline_is_od(record)))
        if best_key is None or rank > best_key:
            best_key = rank
            best_text = format_revision(revision, appendix)
    return best_text


def _official_package_files(
    records: Sequence[FileRecord],
    key: tuple[str, str],
    official: str,
    working: str = "",
    *,
    exclusion: _WorkingExclusion | None = None,
) -> list[FileRecord]:
    """Return present RD files in the official issued package.

    A working-head or annulled package is excluded by **folder name**,
    not by leftover files that still carry ``official`` and not by
    sharing the same NN as a sibling package.
    Packages matching ``official`` win; otherwise the newest remaining
    (non-working and non-annulled) package is used so open-folder and
    «РД · рев.» share the same disk NN when F/send is ``01-AN01`` but
    files are still ``01``.

    Args:
        records: Catalog file records.
        key: Kit identity.
        official: Official filename revision from send/F.
        working: Working filename revision; used when ``exclusion`` has no
            folder identity (legacy rows).
        exclusion: Working and annulled folders/sequences for this kit.

    Returns:
        Files of the chosen official package, or empty.
    """

    has_skip = exclusion is not None and (
        exclusion.folder_keys
        or exclusion.sequences
        or exclusion.annulled_folder_keys
        or exclusion.annulled_sequences
    )
    if not official and not working and not has_skip:
        return []
    groups: dict[str, list[FileRecord]] = {}
    for record in records:
        if not record.present or record.source is not SourceKind.RD:
            continue
        if _record_kit_key(record) != key:
            continue
        package = issued_package_dir(record.path) or record.path
        groups.setdefault(package.casefold(), []).append(record)
    matched: list[list[FileRecord]] = []
    rest: list[list[FileRecord]] = []
    for files in groups.values():
        if _record_group_is_skipped_from_official(
            files, exclusion, working_revision=working
        ):
            continue
        highest = _highest_revision_text(files)
        rest.append(files)
        if official and highest and revision_texts_equivalent(highest, official):
            matched.append(files)
    pool = matched or rest
    if not pool:
        return []

    def _sequence(files: list[FileRecord]) -> int:
        values = [
            _int_or_none(item.data.get("transfer_sequence")) for item in files
        ]
        return max((value if value is not None else -1) for value in values)

    return max(pool, key=_sequence)


def pick_official_rd_package(
    packages: Sequence[KitPackageRow],
    pipeline: KitPipelineRow | None,
) -> KitPackageRow | None:
    """Return the issued RD folder that matches «РД · рев.».

    ``is_current`` (newest non-working NN) is only used when no package
    has ``revision_text`` equivalent to ``official_revision_text``. A
    later delta between the agreed folder and the working head must not
    win open-folder / export.

    Args:
        packages: ``kit_package`` rows for one kit (other kits ignored
            when ``pipeline`` is set).
        pipeline: Derived kit pipeline, or ``None`` when unknown.

    Returns:
        The official non-grey RD package, or ``None``.
    """

    key = (
        kit_identity_key(pipeline.title, pipeline.mark)
        if pipeline is not None
        else None
    )
    exclusion = (
        _exclusion_from_pipeline(pipeline) if pipeline is not None else None
    )
    eligible: list[KitPackageRow] = []
    for package in packages:
        if package.source != "rd" or package.is_grey:
            continue
        if not str(package.package_path or "").strip():
            continue
        if key is not None and kit_identity_key(package.title, package.mark) != key:
            continue
        if _package_is_skipped_from_official(package, exclusion):
            continue
        eligible.append(package)
    if not eligible:
        return None
    official = (
        (pipeline.official_revision_text or "").strip()
        if pipeline is not None
        else ""
    )
    matched = [
        package
        for package in eligible
        if official
        and revision_texts_equivalent(package.revision_text, official)
    ]
    pool = matched or eligible
    return max(
        pool,
        key=lambda package: (
            package.sequence if package.sequence is not None else -1,
            package.id or 0,
        ),
    )


def _highest_mto_revision_text(records: Sequence[FileRecord]) -> str:
    return _highest_revision_text(
        [
            record
            for record in records
            if _file_kind(record) == FileKind.MTO_XLSX.value
        ]
    )


def _as_bool(value: object) -> bool:
    return (
        bool(int(value))
        if isinstance(value, (int, str)) and str(value).isdigit()
        else bool(value)
    )


def _record_is_as_build(record: FileRecord) -> bool:
    return _as_bool(record.data.get("transfer_is_as_build")) or path_is_as_build(
        record.path
    )


def _int_or_none(value: object) -> int | None:
    if value in (None, ""):
        return None
    try:
        return int(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return None


def _build_source_packages(
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
) -> list[KitPackageRow]:
    grouped: dict[tuple[tuple[str, str], str, str], list[FileRecord]] = {}
    titles: dict[tuple[tuple[str, str], str, str], tuple[str, str]] = {}
    for record in records:
        if not record.present:
            continue
        if record.source not in (SourceKind.RD, SourceKind.SQ, SourceKind.ROBOT):
            continue
        identity = _record_title_mark(record)
        if identity is None:
            continue
        package_path = issued_package_dir(record.path)
        if not package_path:
            continue
        group_key = (
            kit_identity_key(*identity),
            record.source.value,
            package_path.casefold(),
        )
        grouped.setdefault(group_key, []).append(record)
        titles.setdefault(group_key, identity)

    packages: list[KitPackageRow] = []
    for group_key, files in grouped.items():
        _kit_key, source, _path_folded = group_key
        title, mark = titles[group_key]
        package_path = issued_package_dir(files[0].path)
        transfer_name = next(
            (
                str(item.data.get("transfer_name")).strip()
                for item in files
                if str(item.data.get("transfer_name") or "").strip()
            ),
            "",
        )
        if not transfer_name:
            transfer_name = PureWindowsPath(package_path).name
        sequence = None
        if source != SourceKind.ROBOT.value:
            for item in files:
                sequence = _int_or_none(item.data.get("transfer_sequence"))
                if sequence is not None:
                    break
        else:
            sequence = None
        mtimes = [
            _int_or_none(item.data.get("mtime_ns")) or 0
            for item in files
        ]
        packages.append(
            KitPackageRow(
                title=title,
                mark=mark,
                source=source,
                sequence=sequence,
                transfer_name=transfer_name,
                package_path=package_path,
                revision_text=_highest_revision_text(files),
                max_mtime_ns=max(mtimes) if mtimes else None,
                pdf_count=sum(
                    1 for item in files if _file_kind(item) == FileKind.PDF.value
                ),
                editable_count=sum(
                    1
                    for item in files
                    if _file_kind(item) == FileKind.SOURCE_EDITABLE.value
                ),
                mto_revision_text=_highest_mto_revision_text(files),
                overlay_current_count=sum(
                    1 for item in files if item.id in detected_current_ids
                ),
                is_grey=False,
                is_current=False,
                is_as_build=any(_record_is_as_build(item) for item in files),
            )
        )
    return packages


def _build_grey_packages(
    packages: Sequence[KitPackageRow],
    sends_with_ids: Sequence[tuple[int | None, IssuanceKit]],
) -> list[KitPackageRow]:
    existing: dict[tuple[str, str], list[KitPackageRow]] = {}
    for pkg in packages:
        if pkg.source == "rd" and not pkg.is_grey:
            existing.setdefault(kit_identity_key(pkg.title, pkg.mark), []).append(pkg)
    seen: set[tuple[tuple[str, str], tuple]] = set()
    grey: list[KitPackageRow] = []
    for _send_id, send in sends_with_ids:
        key = kit_identity_key(send.title, send.mark)
        rd_packages = existing.get(key, ())
        if any(
            revision_texts_equivalent(pkg.revision_text, send.revision_text)
            for pkg in rd_packages
        ):
            continue
        rev_key = _revision_dedupe_key(send.revision_text)
        stamp = (key, rev_key)
        if stamp in seen:
            continue
        seen.add(stamp)
        grey.append(
            KitPackageRow(
                title=send.title,
                mark=send.mark,
                source="issuance_grey",
                sequence=None,
                transfer_name="",
                package_path="",
                revision_text=send.revision_text,
                is_grey=True,
                is_current=False,
            )
        )
    return grey


def _revision_dedupe_key(text: str) -> tuple:
    revision, appendix = parse_sheet_revision(text)
    rank = revision_rank(revision, appendix)
    if rank[0] < 0:
        return ("raw", (text or "").casefold())
    return ("rank", rank[0], rank[1])


def _assign_package_ids(packages: Sequence[KitPackageRow]) -> list[KitPackageRow]:
    return [replace(pkg, id=index) for index, pkg in enumerate(packages, start=1)]


def _mark_current_rd_packages(
    packages: Sequence[KitPackageRow],
    working_by_kit: Mapping[tuple[str, str], str]
    | Mapping[tuple[str, str], _WorkingExclusion]
    | None = None,
) -> list[KitPackageRow]:
    current_ids: set[int] = set()
    working_by_kit = working_by_kit or {}
    by_kit: dict[tuple[str, str], list[KitPackageRow]] = {}
    for pkg in packages:
        if pkg.source == "rd" and not pkg.is_grey:
            by_kit.setdefault(kit_identity_key(pkg.title, pkg.mark), []).append(pkg)
    for key, group in by_kit.items():
        spec = working_by_kit.get(key)
        eligible: list[KitPackageRow]
        if isinstance(spec, _WorkingExclusion):
            eligible = [
                pkg
                for pkg in group
                if not _package_is_skipped_from_official(pkg, spec)
            ]
        else:
            working = spec or ""
            eligible = [
                pkg
                for pkg in group
                if not (
                    working
                    and pkg.revision_text
                    and revision_texts_equivalent(pkg.revision_text, working)
                )
            ]
        if not eligible:
            continue
        current = max(eligible, key=_current_package_sort_key)
        if current.id is not None:
            current_ids.add(current.id)
    return [
        replace(pkg, is_current=pkg.id in current_ids) if pkg.id in current_ids else pkg
        for pkg in packages
    ]


def _current_package_sort_key(
    pkg: KitPackageRow,
) -> tuple[int, tuple[int, int, str], int]:
    sequence = pkg.sequence if pkg.sequence is not None else -1
    return (sequence, _text_rank(pkg.revision_text), pkg.id or 0)


def _display_names(
    google_by_key: dict[tuple[str, str], GoogleKit],
    sends_by_key: dict[tuple[str, str], list[tuple[int | None, IssuanceKit]]],
    packages: Sequence[KitPackageRow],
    records: Sequence[FileRecord],
) -> dict[tuple[str, str], tuple[str, str]]:
    names: dict[tuple[str, str], tuple[str, str]] = {}
    for record in records:
        identity = _record_title_mark(record)
        if identity is None:
            continue
        names.setdefault(kit_identity_key(*identity), identity)
    for pkg in packages:
        names.setdefault(kit_identity_key(pkg.title, pkg.mark), (pkg.title, pkg.mark))
    for key, pairs in sends_by_key.items():
        send = pairs[-1][1]
        names[key] = (send.title, send.mark)
    for key, kit in google_by_key.items():
        names[key] = (kit.title, kit.mark)
    return names


def _text_rank(text: str) -> tuple[int, int, str]:
    revision, appendix = parse_sheet_revision(text)
    return revision_rank(revision, appendix)


def _overlay_current_revision_text(
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    key: tuple[str, str],
) -> str:
    return _highest_revision_text(
        [
            record
            for record in records
            if record.present
            and record.source is SourceKind.RD
            and record.id in detected_current_ids
            and _record_kit_key(record) == key
        ]
    )


def _sorted_sends(
    sends: Sequence[tuple[int | None, IssuanceKit]],
) -> list[tuple[int | None, IssuanceKit]]:
    return sorted(
        sends,
        key=lambda pair: (
            _send_sort_key(pair[1]),
            pair[0] if pair[0] is not None else 0,
        ),
    )


def _send_sort_key(send: IssuanceKit) -> tuple[str, int]:
    sortable = send.send_date_sortable or format_event_date_sortable(send.send_date)
    return (sortable, send.row_index)


def _last_official_cycle(
    official_revision: str,
    ordered_sends: Sequence[tuple[int | None, IssuanceKit]],
    cycles: Sequence[KitCycleRow],
) -> tuple[KitCycleRow | None, IssuanceKit | None]:
    cycle_by_send_id = {
        cycle.send_id: cycle for cycle in cycles if cycle.send_id is not None
    }
    matching = [
        (send_id, send)
        for send_id, send in ordered_sends
        if official_revision
        and revision_texts_equivalent(send.revision_text, official_revision)
    ]
    chosen = matching[-1] if matching else None
    if chosen is None:
        return None, None
    send_id, send = chosen
    if send_id is not None:
        return cycle_by_send_id.get(send_id), send
    for cycle in reversed(cycles):
        if cycle.send_id is None and revision_texts_equivalent(
            cycle.revision_text, send.revision_text
        ):
            return cycle, send
    return None, send


def _official_send_f_revision(
    last_send: IssuanceKit | None,
    events: Sequence[tuple[int, KitEvent]],
) -> str:
    if last_send is not None and last_send.revision_text:
        return last_send.revision_text
    for _event_id, event in reversed(list(events)):
        text = format_revision(event.revision, event.appendix)
        if text:
            return text
    return ""


def _events_on_revision(
    events: Sequence[tuple[int, KitEvent]],
    revision_text: str,
) -> list[tuple[int, KitEvent]]:
    matched: list[tuple[int, KitEvent]] = []
    for event_id, event in events:
        event_rev = format_revision(event.revision, event.appendix)
        if event_rev:
            if revision_text and revision_texts_equivalent(event_rev, revision_text):
                matched.append((event_id, event))
        elif not revision_text:
            matched.append((event_id, event))
    return matched


def _last_code_on_revision(
    events: Sequence[tuple[int, KitEvent]],
    official_revision: str,
    last_cycle: KitCycleRow | None,
) -> KitEvent | None:
    scoped = _events_on_revision(events, official_revision)
    codes = [
        event
        for _event_id, event in scoped
        if event.stage in _CODE_STAGES
    ]
    if codes:
        return codes[-1]
    if (
        last_cycle is not None
        and last_cycle.code_event_id is not None
        and (
            not official_revision
            or revision_texts_equivalent(last_cycle.revision_text, official_revision)
        )
    ):
        for event_id, event in events:
            if event_id == last_cycle.code_event_id:
                return event
    return None


def _issuance_accepted(send: IssuanceKit | None) -> bool:
    if send is None:
        return False
    if (send.incoming_control_date or "").strip():
        return True
    return "принят" in (send.status or "").casefold()


def _code_date_text(code_date: str) -> str:
    parsed = parse_ddmmyyyy(code_date)
    if parsed is not None:
        return f"{parsed.day:02d}.{parsed.month:02d}.{parsed.year:04d}"
    return code_date


def _tdo_passed_date_from_facts(
    *,
    events: Sequence[KitEvent],
    issuance: IssuanceKit | None,
    official_revision: str,
) -> str:
    tdo_dates: list[str] = []
    incoming_dates: list[str] = []
    for event in events:
        if not event.date:
            continue
        event_rev = format_revision(event.revision, event.appendix)
        if official_revision and event_rev:
            if not revision_texts_equivalent(event_rev, official_revision):
                continue
        elif official_revision and not event_rev:
            continue
        if event.stage == "tdo_passed":
            tdo_dates.append(event.date)
        elif event.stage == "incoming_passed":
            incoming_dates.append(event.date)
    if tdo_dates:
        return tdo_dates[-1]
    if incoming_dates:
        return incoming_dates[-1]
    if issuance is not None and _issuance_accepted(issuance):
        return (
            (issuance.incoming_control_date or "").strip()
            or (issuance.send_date or "").strip()
        )
    return ""


def _tdo_passed_date_for_cycle(
    *,
    official_revision: str,
    last_send: IssuanceKit | None,
    last_cycle: KitCycleRow | None,
    events: Sequence[tuple[int, KitEvent]],
) -> str:
    cycle_events = _cycle_scoped_events(
        events, official_revision, last_send, last_cycle
    )
    tdo_dates = [
        event.date
        for event in cycle_events
        if event.stage == "tdo_passed" and event.date
    ]
    if tdo_dates:
        return tdo_dates[-1]
    incoming_dates = [
        event.date
        for event in cycle_events
        if event.stage == "incoming_passed" and event.date
    ]
    if incoming_dates:
        return incoming_dates[-1]
    if last_send is not None and _issuance_accepted(last_send):
        return (
            (last_send.incoming_control_date or "").strip()
            or (last_send.send_date or "").strip()
        )
    return ""


def _events_for_send(
    events: Sequence[tuple[int, KitEvent]],
    send: IssuanceKit | None,
) -> list[KitEvent]:
    if send is None:
        return []
    return [
        event
        for _event_id, event in events
        if match_event_reason(send, event) is not None
    ]


def _cycle_scoped_events(
    events: Sequence[tuple[int, KitEvent]],
    official_revision: str,
    last_send: IssuanceKit | None,
    last_cycle: KitCycleRow | None,
) -> list[KitEvent]:
    scoped = _events_for_send(events, last_send)
    if last_cycle is not None and last_cycle.tdo_event_id is not None:
        events_by_id = {event_id: event for event_id, event in events}
        tdo_event = events_by_id.get(last_cycle.tdo_event_id)
        if tdo_event is not None and tdo_event not in scoped:
            scoped.append(tdo_event)
    if not scoped and official_revision:
        scoped = [event for _eid, event in _events_on_revision(events, official_revision)]
    return scoped


def _review_status(
    *,
    official_revision: str,
    last_send: IssuanceKit | None,
    last_cycle: KitCycleRow | None,
    events: Sequence[tuple[int, KitEvent]],
    has_official_send: bool,
    allow_agreed: bool = True,
) -> KitPipelineStatus:
    """Return review-only pipeline status for the last official cycle."""

    cycle_events = _cycle_scoped_events(
        events, official_revision, last_send, last_cycle
    )
    if allow_agreed:
        for event in cycle_events:
            if event.stage in {"code_a", "agreed"} and _code_belongs_to_current_cycle(
                event,
                official_revision=official_revision,
                last_send=last_send,
                last_cycle=last_cycle,
                events=events,
            ):
                return KitPipelineStatus.AGREED
    cycle_matches_official = last_cycle is not None and (
        not official_revision
        or revision_texts_equivalent(last_cycle.revision_text, official_revision)
    )
    accepted_send = last_send if cycle_matches_official else None
    if _issuance_accepted(accepted_send) or any(
        event.stage in _PASSED_STAGES for event in cycle_events
    ):
        return KitPipelineStatus.TDO_REVIEW
    if has_official_send or any(event.stage in _SENT_STAGES for event in cycle_events):
        return KitPipelineStatus.SENT_TDO
    return KitPipelineStatus.NOT_UPLOADED


def _cycle_anchor_date(
    send: IssuanceKit,
    tdo_event: KitEvent | None,
) -> date | None:
    if tdo_event is not None and tdo_event.date:
        parsed = parse_ddmmyyyy(tdo_event.date)
        if parsed is not None:
            return parsed
    for field in (send.incoming_control_date, send.send_date):
        parsed = parse_ddmmyyyy(field)
        if parsed is not None:
            return parsed
    return None


def _code_before_cycle(
    code_event: KitEvent,
    send: IssuanceKit,
    tdo_event: KitEvent | None,
) -> bool:
    code_date = parse_ddmmyyyy(code_event.date)
    if code_date is None:
        return False
    send_trm = (send.send_transmittal or "").strip()
    if not send_trm:
        return False
    folded = send_trm.casefold()
    if any((token or "").casefold() == folded for token in code_event.transmittals):
        return False
    anchor = _cycle_anchor_date(send, tdo_event)
    if anchor is None:
        return False
    return code_date < anchor


def _code_belongs_to_current_cycle(
    code_event: KitEvent,
    *,
    official_revision: str,
    last_send: IssuanceKit | None,
    last_cycle: KitCycleRow | None,
    events: Sequence[tuple[int, KitEvent]],
) -> bool:
    """Return whether an F letter belongs to the current official cycle.

    Same filename revision is current when there is no later send/TDO of
    that revision with a different TRM. A TRM-less legalized send (file
    mtime in the journal) is not a new customer cycle and must not stale
    the letter. A later sheet send with another TRM is 9000-KSB.
    """

    send_trm = (
        (last_send.send_transmittal or "").strip() if last_send is not None else ""
    )
    if send_trm:
        folded = send_trm.casefold()
        if any((token or "").casefold() == folded for token in code_event.transmittals):
            return True
    event_rev = format_revision(code_event.revision, code_event.appendix)
    if not event_rev or not revision_texts_equivalent(event_rev, official_revision):
        return False
    if last_send is None:
        return True
    if dates_within_days(code_event.date, last_send.send_date) or dates_within_days(
        code_event.date, last_send.incoming_control_date
    ):
        return True
    events_by_id = {event_id: event for event_id, event in events}
    tdo_event = (
        events_by_id.get(last_cycle.tdo_event_id)
        if last_cycle is not None and last_cycle.tdo_event_id is not None
        else None
    )
    if not send_trm:
        return True
    if _code_before_cycle(code_event, last_send, tdo_event):
        return False
    cycle_date = _cycle_anchor_date(last_send, tdo_event)
    code_dt = parse_ddmmyyyy(code_event.date)
    if code_dt is None or cycle_date is None:
        return False
    return cycle_date <= code_dt


def _derive_approval_fields(
    *,
    events: Sequence[tuple[int, KitEvent]],
    official_revision: str,
    send_f_revision: str,
    last_send: IssuanceKit | None,
    last_cycle: KitCycleRow | None,
    ordered_sends: Sequence[tuple[int | None, IssuanceKit]],
    packages: Sequence[KitPackageRow],
) -> tuple[str | None, str, str, bool, str | None]:
    cycle_events = _events_for_send(events, last_send)
    cycle_has_agreed = any(
        event.stage == "agreed"
        and _code_belongs_to_current_cycle(
            event,
            official_revision=official_revision,
            last_send=last_send,
            last_cycle=last_cycle,
            events=events,
        )
        for event in cycle_events
    )
    cycle_has_letter_a = any(
        event.stage == "code_a"
        and _code_belongs_to_current_cycle(
            event,
            official_revision=official_revision,
            last_send=last_send,
            last_cycle=last_cycle,
            events=events,
        )
        for event in cycle_events
    )

    letter_event: KitEvent | None = None
    for _event_id, event in events:
        if event.stage in _CODE_STAGES:
            letter_event = event

    if cycle_has_agreed and not cycle_has_letter_a:
        agreed_date = ""
        for event in reversed(cycle_events):
            if event.stage == "agreed":
                agreed_date = event.date or ""
                break
        return (
            "A",
            official_revision or send_f_revision,
            agreed_date,
            False,
            "customer",
        )

    if letter_event is None:
        return None, "", "", False, None

    code, _stage = _CODE_STAGES[letter_event.stage]
    code_revision_text = format_revision(letter_event.revision, letter_event.appendix)
    code_date = letter_event.date or ""
    code_stale = not _code_belongs_to_current_cycle(
        letter_event,
        official_revision=official_revision,
        last_send=last_send,
        last_cycle=last_cycle,
        events=events,
    )
    code_origin = derive_code_origin(
        code,
        letter_event,
        official_revision or send_f_revision,
        ordered_sends,
        packages,
        last_send=last_send,
    )
    return code, code_revision_text, code_date, code_stale, code_origin


def _review_as_build_for_cycle(
    last_cycle: KitCycleRow | None,
    packages: Sequence[KitPackageRow],
) -> bool:
    if last_cycle is None or last_cycle.package_id is None:
        return False
    for package in packages:
        if package.id == last_cycle.package_id:
            return package.is_as_build
    return False


def _working_as_build(
    packages: Sequence[KitPackageRow],
    exclusion: _WorkingExclusion,
    working_revision: str,
) -> bool:
    """Return whether a working folder of ``working_revision`` is as-build."""

    if not str(working_revision or "").strip():
        return False
    for package in packages:
        if package.source != "rd" or package.is_grey:
            continue
        if not _package_is_working(package, exclusion):
            continue
        if not revision_texts_equivalent(package.revision_text, working_revision):
            continue
        if package.is_as_build:
            return True
    return False


def _send_is_after_code(
    send: IssuanceKit,
    code_date: date | None,
    last_send: IssuanceKit | None,
) -> bool:
    send_date = parse_ddmmyyyy(send.send_date)
    if code_date is not None and send_date is not None:
        return send_date > code_date
    if last_send is not None:
        return _send_sort_key(send) > _send_sort_key(last_send)
    return False


def _later_higher_rd_package(
    packages: Sequence[KitPackageRow],
    code_rank: tuple[int, int, str],
    code_date: date | None,
    code_event: KitEvent,
) -> bool:
    code_dt = None
    if code_event.date:
        parsed = parse_ddmmyyyy(code_event.date)
        if parsed is not None:
            try:
                code_dt = datetime.combine(parsed, datetime.min.time())
            except ValueError:
                code_dt = None
    for pkg in packages:
        if pkg.source != "rd" or pkg.is_grey:
            continue
        if _text_rank(pkg.revision_text) <= code_rank:
            continue
        if code_date is None or pkg.max_mtime_ns is None:
            return True
        try:
            pkg_dt = datetime.fromtimestamp(int(pkg.max_mtime_ns) / _NS)
        except (OSError, OverflowError, ValueError):
            return True
        if code_dt is not None and pkg_dt > code_dt:
            return True
        if code_dt is None and pkg_dt.date() > code_date:
            return True
    return False


def _best_tdo_event_id(
    matched: Sequence[tuple[int, KitEvent, str]],
    events: Sequence[tuple[int, KitEvent]],
    send: IssuanceKit,
) -> int | None:
    candidates = [
        (event_id, event)
        for event_id, event, _reason in matched
        if event.stage in _TDO_STAGE_PRIORITY
    ]
    if not candidates:
        candidates = [
            (event_id, event)
            for event_id, event in events
            if event.stage in _TDO_STAGE_PRIORITY
            and (
                (
                    format_revision(event.revision, event.appendix)
                    and revision_texts_equivalent(
                        send.revision_text,
                        format_revision(event.revision, event.appendix),
                    )
                )
                or not format_revision(event.revision, event.appendix)
            )
        ]
    if not candidates:
        return None
    best_id, _best = min(
        candidates,
        key=lambda pair: (
            _TDO_STAGE_PRIORITY.get(pair[1].stage, 99),
            -_event_seq_hint(events, pair[0]),
        ),
    )
    return best_id


def _last_code_event_id(
    matched: Sequence[tuple[int, KitEvent, str]],
    events: Sequence[tuple[int, KitEvent]],
    send: IssuanceKit,
    *,
    tdo_event_id: int | None = None,
) -> int | None:
    events_by_id = {event_id: event for event_id, event in events}
    tdo_event = (
        events_by_id.get(tdo_event_id) if tdo_event_id is not None else None
    )
    for reason in ("trm", "date"):
        matched_codes = [
            (event_id, event)
            for event_id, event, match_reason in matched
            if event.stage in _CODE_STAGES and match_reason == reason
        ]
        if matched_codes:
            return matched_codes[-1][0]
    revision_codes = [
        (event_id, event)
        for event_id, event, match_reason in matched
        if event.stage in _CODE_STAGES
        and match_reason == "revision"
        and not _code_before_cycle(event, send, tdo_event)
    ]
    if revision_codes:
        return revision_codes[-1][0]
    fallback_codes = [
        (event_id, event)
        for event_id, event in events
        if event.stage in _CODE_STAGES
        and format_revision(event.revision, event.appendix)
        and revision_texts_equivalent(
            send.revision_text,
            format_revision(event.revision, event.appendix),
        )
        and not _code_before_cycle(event, send, tdo_event)
    ]
    if fallback_codes:
        return fallback_codes[-1][0]
    return None


def _event_seq_hint(
    events: Sequence[tuple[int, KitEvent]], event_id: int
) -> int:
    for index, (candidate_id, _event) in enumerate(events):
        if candidate_id == event_id:
            return index
    return -1


def _liquidity_anchor_month(
    official_revision: str,
    last_send: IssuanceKit | None,
    events: Sequence[tuple[int, KitEvent]],
) -> tuple[int, int] | None:
    tdo_dates = [
        event.date
        for _event_id, event in _events_on_revision(events, official_revision)
        if event.stage == "tdo_passed" and event.date
    ]
    anchor_text = ""
    if tdo_dates:
        anchor_text = tdo_dates[-1] or ""
    elif last_send is not None:
        anchor_text = (
            (last_send.incoming_control_date or "").strip()
            or (last_send.send_date or "").strip()
        )
    parsed = parse_ddmmyyyy(anchor_text)
    if parsed is None:
        return None
    return parsed.year, parsed.month


def _review_identity(review: LiquidityReviewRow) -> tuple[str, str, str, str]:
    return (
        review.title.casefold(),
        review.mark.casefold(),
        (review.transfer_name or "").casefold(),
        (review.revision_text or "").casefold(),
    )


def list_revision_matrix(database: CatalogDatabase) -> tuple[KitRevisionRow, ...]:
    """Return heatmap cells sorted for the «Ревизии MTO» tab.

    Args:
        database: Initialized catalog database.

    Returns:
        Cells grouped by title/mark, then filename-revision rank.
    """

    with perf_span("pipeline.list_revision_matrix"):
        cells = database.list_revision_cells()
        return tuple(
            sorted(
                cells,
                key=lambda cell: (
                    cell.title.casefold(),
                    cell.mark.casefold(),
                    _text_rank(cell.revision_text),
                    cell.id or 0,
                ),
            )
        )


def list_revision_columns(cells: Sequence[KitRevisionRow]) -> tuple[str, ...]:
    """Return unique filename revisions sorted by :func:`revision_rank`.

    Args:
        cells: Heatmap cells from :func:`list_revision_matrix`.

    Returns:
        Revision texts in rank order (one canonical string per equivalent rev).
    """

    unique: dict[tuple, str] = {}
    for cell in cells:
        text = (cell.revision_text or "").strip()
        if not text:
            continue
        unique[_revision_dedupe_key(text)] = text
    return tuple(sorted(unique.values(), key=_text_rank))


def _revision_has_f_status(
    revision_text: str,
    events: Sequence[tuple[int, KitEvent]],
    sends: Sequence[IssuanceKit],
) -> bool:
    """True when this revision has a column-F event or a send."""

    if not (revision_text or "").strip():
        return bool(events) or bool(sends)
    for _event_id, event in events:
        event_rev = format_revision(event.revision, event.appendix)
        if revision_texts_equivalent(revision_text, event_rev):
            return True
    for send in sends:
        send_rev = send.revision_text or format_revision(
            send.revision, send.appendix
        )
        if revision_texts_equivalent(revision_text, send_rev):
            return True
    return False


def _synthetic_revision_text(
    google: GoogleKit | None,
    sends: Sequence[IssuanceKit],
) -> str:
    """Google D/E revision text, else the latest send, else empty."""

    if google is not None:
        text = (google.sheet_revision_text or "").strip()
        if not text:
            text = format_revision(
                google.sheet_revision, google.sheet_appendix
            )
        if text:
            return text
    if sends:
        latest = sends[-1]
        return (
            latest.revision_text
            or format_revision(latest.revision, latest.appendix)
            or ""
        )
    return ""


def list_mto_worklist(
    database: CatalogDatabase,
    *,
    records: Sequence[FileRecord],
    rd_root: str | Path | None = None,
) -> tuple[MtoWorklistRow, ...]:
    """Return heatmap cells plus Google/issuance kits that have no cells.

    The universe matches «Комплекты»: RD heatmap cells union KSB ИД
    kits union «Выдача РД ПД» sends. A kit that already has cells never
    also gets a synthetic row. ``iter_mto_files_for_cells`` still reads
    ``list_revision_matrix`` and does not see synthetic rows.

    Args:
        database: Initialized catalog database.
        records: Catalog file records from the last scan.
        rd_root: RD source root. When set, non-canonical RD files are
            excluded via :func:`record_has_canonical_layout`.

    Returns:
        Rows sorted by title, mark, then filename-revision rank.
        ``gap_kind`` is ``no_package`` when no non-grey RD folder is
        linked, ``no_mto_file`` when a folder exists but no MTO xlsx
        matched, ``no_rd`` for a Google/issuance kit with no heatmap
        cells, otherwise empty. ``package_path`` is the official
        ``is_current`` issued folder, not the working NN.
    """

    with perf_span("pipeline.list_mto_worklist"):
        cells = list_revision_matrix(database)
        packages_by_id: dict[int, KitPackageRow] = {
            pkg.id: pkg
            for pkg in database.list_kit_packages()
            if pkg.id is not None
        }
        google_by_key = {
            kit_identity_key(kit.title, kit.mark): kit
            for kit in database.list_google_kits()
        }
        sends_by_key: dict[tuple[str, str], list[IssuanceKit]] = {}
        for _send_id, send in _effective_issuance_sends(database):
            key = kit_identity_key(send.title, send.mark)
            sends_by_key.setdefault(key, []).append(send)
        events_by_kit = database.list_google_events_by_kit()
        pipeline_by_kit = {
            kit_identity_key(row.title, row.mark): row
            for row in database.list_kit_pipelines()
        }
        index = _mto_records_by_kit_revision(database, records, rd_root=rd_root)
        ranked: list[tuple[MtoWorklistRow, int]] = []
        cell_keys: set[tuple[str, str]] = set()
        for cell in cells:
            key = kit_identity_key(cell.title, cell.mark)
            cell_keys.add(key)
            if cell.pipeline_status in _SKIPPED_HEATMAP_STATUSES:
                continue
            cell_keys.add(key)
            pipeline = pipeline_by_kit.get(key)
            official_text = (
                pipeline.official_revision_text if pipeline is not None else ""
            )
            working_text = (
                pipeline.working_revision_text if pipeline is not None else ""
            )
            rd_packages = [
                pkg
                for pkg_id in _decode_package_ids(cell)
                if (pkg := packages_by_id.get(pkg_id)) is not None
                and pkg.source == "rd"
                and not pkg.is_grey
            ]
            package_count = len(rd_packages)
            if rd_packages:
                chosen_packages = [pkg for pkg in rd_packages if pkg.is_current]
                if not chosen_packages and official_text:
                    chosen_packages = [
                        pkg
                        for pkg in rd_packages
                        if revision_texts_equivalent(pkg.revision_text, official_text)
                    ]
                newest = max(
                    chosen_packages or rd_packages,
                    key=lambda pkg: (
                        pkg.sequence if pkg.sequence is not None else -1,
                        pkg.id or 0,
                    ),
                )
                package_path = newest.package_path or ""
                package_label = (
                    newest.transfer_name
                    or PureWindowsPath(newest.package_path).name
                )
            else:
                package_path = ""
                package_label = ""
            candidates = index.get(
                (
                    *key,
                    _revision_dedupe_key(cell.revision_text),
                ),
                (),
            )
            chosen: FileRecord | None = None
            if candidates:
                prefix = package_path.casefold()
                chosen = max(
                    candidates,
                    key=lambda record: (
                        bool(prefix and record.path.casefold().startswith(prefix)),
                        _record_mtime_ns(record) or 0,
                        record.path_key,
                    ),
                )
            mto_path = chosen.path if chosen is not None else ""
            mto_mtime_ns = (
                _record_mtime_ns(chosen) if chosen is not None else None
            )
            if not rd_packages:
                gap_kind = "no_package"
            elif not mto_path:
                gap_kind = "no_mto_file"
            else:
                gap_kind = ""
            kit_sends = sends_by_key.get(key, ())
            is_current = bool(cell.is_current)
            if working_text and official_text:
                is_current = revision_texts_equivalent(
                    cell.revision_text, official_text
                )
            ranked.append(
                (
                    MtoWorklistRow(
                        title=cell.title,
                        mark=cell.mark,
                        revision_text=cell.revision_text,
                        status=cell.pipeline_status,
                        letters=cell.letters,
                        is_as_build=bool(cell.is_as_build),
                        is_current=is_current,
                        is_current_ifc=bool(cell.is_current_ifc),
                        has_mto=bool(cell.has_mto),
                        mto_path=mto_path,
                        mto_mtime_ns=mto_mtime_ns,
                        package_path=package_path,
                        package_label=package_label or "",
                        package_count=package_count,
                        problem_kinds=_decode_problem_kinds(cell),
                        gap_kind=gap_kind,
                        in_google=key in google_by_key,
                        in_issuance=bool(kit_sends),
                        has_f_status=_revision_has_f_status(
                            cell.revision_text,
                            events_by_kit.get(key, ()),
                            kit_sends,
                        ),
                    ),
                    cell.id or 0,
                )
            )
        extra_keys = (set(google_by_key) | set(sends_by_key)) - cell_keys
        for key in extra_keys:
            google = google_by_key.get(key)
            kit_sends = sends_by_key.get(key, ())
            if google is not None:
                title, mark = google.title, google.mark
            else:
                latest = kit_sends[-1]
                title, mark = latest.title, latest.mark
            revision_text = _synthetic_revision_text(google, kit_sends)
            ranked.append(
                (
                    MtoWorklistRow(
                        title=title,
                        mark=mark,
                        revision_text=revision_text,
                        status="",
                        letters="",
                        is_as_build=False,
                        is_current=False,
                        is_current_ifc=False,
                        has_mto=False,
                        gap_kind="no_rd",
                        in_google=google is not None,
                        in_issuance=bool(kit_sends),
                        has_f_status=_revision_has_f_status(
                            revision_text,
                            events_by_kit.get(key, ()),
                            kit_sends,
                        ),
                    ),
                    0,
                )
            )
        ranked.sort(
            key=lambda item: (
                item[0].title.casefold(),
                item[0].mark.casefold(),
                _text_rank(item[0].revision_text),
                item[1],
                item[0].revision_text.casefold(),
            )
        )
        return tuple(row for row, _tie in ranked)


def resolve_approved_contour(
    database: CatalogDatabase,
    *,
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    title: str,
    mark: str,
    rd_root: str | Path | None = None,
) -> ApprovedContour | None:
    """Return the approved contour for one kit, or ``None`` if unknown.

    Args:
        database: Initialized catalog database.
        records: Catalog file records from the last scan.
        detected_current_ids: Global overlay-current file ids (accepted for
            call-site symmetry; contour MTO is recomputed as-of the approved
            package and does not rank by these ids).
        title: Four-digit title.
        mark: Latin AGCC mark.
        rd_root: RD source root. When set, non-canonical RD files are
            excluded via :func:`record_has_canonical_layout`.

    Returns:
        Frozen contour, or ``None`` when the kit has no ``kit_pipeline`` row.
    """

    key = kit_identity_key(title, mark)
    for contour in resolve_all_approved_contours(
        database,
        records=records,
        detected_current_ids=detected_current_ids,
        rd_root=rd_root,
    ):
        if kit_identity_key(contour.title, contour.mark) == key:
            return contour
    return None


def resolve_all_approved_contours(
    database: CatalogDatabase,
    *,
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    rd_root: str | Path | None = None,
) -> tuple[ApprovedContour, ...]:
    """Resolve the approved contour for every kit that has a pipeline row.

    Loads packages, cycles, pipelines, overlay-MTO collisions, and files
    once. Does not write to SQLite.

    Args:
        database: Initialized catalog database.
        records: Catalog file records from the last scan.
        detected_current_ids: Global overlay-current file ids (unused for
            ranking; see :func:`resolve_approved_contour`).
        rd_root: RD source root. When set, non-canonical RD files are
            excluded via :func:`record_has_canonical_layout`.

    Returns:
        Contours in pipeline title/mark order.
    """

    return resolve_all_contours(
        database,
        records=records,
        detected_current_ids=detected_current_ids,
        rd_root=rd_root,
        anchor_stages=_CODE_STAGE_KEYS,
    )


def resolve_all_contours(
    database: CatalogDatabase,
    *,
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    rd_root: str | Path | None = None,
    anchor_stages: Collection[str] | None = None,
) -> tuple[ApprovedContour, ...]:
    """Resolve a journal-anchored contour for every kit with a pipeline row.

    Same matching as :func:`resolve_all_approved_contours`, but the cycle
    pool is filtered by ``anchor_stages`` instead of being hard-wired to
    code A/B/C. ``None`` (and :data:`_CODE_STAGE_KEYS`) keep the approved
    contour behaviour, including the fallback to any cycle. Passed-stage
    anchors (``tdo_passed`` / ``incoming_passed``) do not fall back to a
    non-matching cycle and do not pick the latest RD package when the
    journal revision is missing.

    Args:
        database: Initialized catalog database.
        records: Catalog file records from the last scan.
        detected_current_ids: Global overlay-current file ids (unused for
            ranking).
        rd_root: RD source root. When set, non-canonical RD files are
            excluded via :func:`record_has_canonical_layout`.
        anchor_stages: F-event stages that make a cycle eligible. Default
            is code A/B/C.

    Returns:
        Contours in pipeline title/mark order.
    """

    del detected_current_ids
    stages = frozenset(anchor_stages) if anchor_stages is not None else _CODE_STAGE_KEYS
    packages = database.list_kit_packages()
    cycles = database.list_kit_cycles()
    pipelines = database.list_kit_pipelines()
    collisions = [
        row
        for row in database.list_current_collisions()
        if str(row.get("scope") or "") == "overlay_mto"
    ]
    files_by_path = _files_by_path_key(database, records, rd_root=rd_root)
    events_by_kit = database.list_google_events_by_kit()
    events_by_id: dict[int, KitEvent] = {}
    for pairs in events_by_kit.values():
        for event_id, event in pairs:
            events_by_id[event_id] = event

    packages_by_kit: dict[tuple[str, str], list[KitPackageRow]] = {}
    packages_by_id: dict[int, KitPackageRow] = {}
    for package in packages:
        packages_by_kit.setdefault(
            kit_identity_key(package.title, package.mark), []
        ).append(package)
        if package.id is not None:
            packages_by_id[package.id] = package
    cycles_by_kit: dict[tuple[str, str], list[KitCycleRow]] = {}
    for cycle in cycles:
        cycles_by_kit.setdefault(
            kit_identity_key(cycle.title, cycle.mark), []
        ).append(cycle)
    files_by_kit: dict[tuple[str, str], list[FileRecord]] = {}
    for record in files_by_path.values():
        identity = _record_kit_key(record)
        if identity is None:
            continue
        files_by_kit.setdefault(identity, []).append(record)

    return tuple(
        _resolve_one_approved_contour(
            pipeline,
            packages=packages_by_kit.get(
                kit_identity_key(pipeline.title, pipeline.mark), ()
            ),
            packages_by_id=packages_by_id,
            cycles=cycles_by_kit.get(
                kit_identity_key(pipeline.title, pipeline.mark), ()
            ),
            kit_files=files_by_kit.get(
                kit_identity_key(pipeline.title, pipeline.mark), ()
            ),
            collisions=collisions,
            files_by_path=files_by_path,
            events=events_by_kit.get(
                kit_identity_key(pipeline.title, pipeline.mark), ()
            ),
            events_by_id=events_by_id,
            anchor_stages=stages,
        )
        for pipeline in pipelines
    )


def _resolve_one_approved_contour(
    pipeline: KitPipelineRow,
    *,
    packages: Sequence[KitPackageRow],
    packages_by_id: dict[int, KitPackageRow],
    cycles: Sequence[KitCycleRow],
    kit_files: Sequence[FileRecord],
    collisions: Sequence[dict],
    files_by_path: dict[str, FileRecord],
    events: Sequence[tuple[int, KitEvent]] = (),
    events_by_id: Mapping[int, KitEvent] | None = None,
    anchor_stages: Collection[str] = _CODE_STAGE_KEYS,
) -> ApprovedContour:
    stages = frozenset(anchor_stages)
    is_code_anchor = bool(stages & _CODE_STAGE_KEYS)
    event_map = events_by_id or {}
    cycle = _last_cycle_for_stages(
        cycles,
        stages=stages,
        events_by_id=event_map,
        fallback=is_code_anchor,
    )
    if is_code_anchor:
        approved_revision = (pipeline.code_revision_text or "").strip() or (
            pipeline.official_revision_text or ""
        ).strip()
    else:
        approved_revision = _anchor_revision_text(cycle, events, stages)
    # Approved keeps last-package always (journal vs filename divergence).
    # TDO uses the same chain only when a passed-stage anchor exists; a kit
    # that never passed ТДО must stay no_source, not "latest RD package".
    allow_last_package = is_code_anchor or cycle is not None or bool(
        approved_revision
    )
    kit_key = kit_identity_key(pipeline.title, pipeline.mark)
    present_rd = [
        record
        for record in kit_files
        if record.present and record.source is SourceKind.RD
    ]
    candidates = tuple(
        _iter_package_candidates(
            packages,
            packages_by_id=packages_by_id,
            cycle=cycle,
            kit_files=kit_files,
            approved_revision_text=approved_revision,
            allow_last_package=allow_last_package,
        )
    )
    package: KitPackageRow | None = None
    match_reason = "none"
    mto_path = ""
    mto_source = ""
    if is_code_anchor:
        if candidates:
            package, match_reason = candidates[0]
            mto_path, mto_source = _contour_mto(present_rd, package)
    else:
        for candidate, reason in candidates:
            path, source = _contour_mto(present_rd, candidate)
            if path:
                package, match_reason = candidate, reason
                mto_path, mto_source = path, source
                break
        if not mto_path and candidates:
            package, match_reason = candidates[0]
            mto_path, mto_source = _contour_mto(present_rd, package)

    in_contour_kinds = _in_contour_collision_kinds(
        collisions,
        files_by_path=files_by_path,
        kit_key=kit_key,
        max_sequence=package.sequence if package is not None else None,
        package_resolved=package is not None,
        annulled_folder_keys=pipeline.annulled_transfer_names,
    )
    ambiguity = _approved_ambiguity(
        package=package,
        mto_path=mto_path,
        in_contour_kinds=in_contour_kinds,
    )
    revision_mismatch = bool(
        package is not None
        and approved_revision
        and not revision_texts_equivalent(package.revision_text, approved_revision)
    )
    confidence = _approved_confidence(
        package=package,
        mto_path=mto_path,
        mto_source=mto_source,
        ambiguity=ambiguity,
        revision_mismatch=revision_mismatch,
    )
    warnings = _approved_warnings(
        match_reason=match_reason,
        mto_source=mto_source,
        ambiguity=ambiguity,
        approved_revision_text=approved_revision,
        package_revision_text=(
            package.revision_text if package is not None else ""
        ),
        revision_mismatch=revision_mismatch,
    )
    return ApprovedContour(
        title=pipeline.title,
        mark=pipeline.mark,
        approved_revision_text=approved_revision,
        package_id=package.id if package is not None else None,
        package_path=package.package_path if package is not None else "",
        package_sequence=package.sequence if package is not None else None,
        match_reason=match_reason,
        confidence=confidence,
        mto_path=mto_path,
        mto_source=mto_source,
        ambiguity=ambiguity,
        warnings=warnings,
    )


def _contour_mto(
    present_rd: Sequence[FileRecord],
    package: KitPackageRow | None,
) -> tuple[str, str]:
    """Return ``(mto_path, mto_source)`` for overlay as-of ``package``."""

    if package is None:
        return "", ""
    mto_file = _overlay_mto_as_of(present_rd, max_sequence=package.sequence)
    if mto_file is None:
        return "", ""
    if package.package_path and path_is_under(mto_file.path, package.package_path):
        return mto_file.path, "package"
    return mto_file.path, "inherited"


def _resolve_approved_package(
    packages: Sequence[KitPackageRow],
    *,
    packages_by_id: dict[int, KitPackageRow],
    cycle: KitCycleRow | None,
    kit_files: Sequence[FileRecord],
    approved_revision_text: str,
    allow_last_package: bool = True,
) -> tuple[KitPackageRow | None, str]:
    for package, reason in _iter_package_candidates(
        packages,
        packages_by_id=packages_by_id,
        cycle=cycle,
        kit_files=kit_files,
        approved_revision_text=approved_revision_text,
        allow_last_package=allow_last_package,
    ):
        return package, reason
    return None, "none"


def _iter_package_candidates(
    packages: Sequence[KitPackageRow],
    *,
    packages_by_id: dict[int, KitPackageRow],
    cycle: KitCycleRow | None,
    kit_files: Sequence[FileRecord],
    approved_revision_text: str,
    allow_last_package: bool,
) -> list[tuple[KitPackageRow, str]]:
    """Yield unique packages in cycle_* → folder_rev → file_rev → last_package order.

    Within ``cycle_*`` and ``folder_rev``, a package whose full revision
    (base plus appendix) matches ``approved_revision_text`` outranks a
    base-only match. Reason order itself is unchanged.
    """

    rd_packages = [
        pkg for pkg in packages if pkg.source == "rd" and not pkg.is_grey
    ]
    ordered: list[tuple[KitPackageRow, str]] = []
    seen: set[tuple] = set()

    def add(package: KitPackageRow | None, reason: str) -> None:
        if package is None:
            return
        key = (
            package.id
            if package.id is not None
            else (package.package_path.casefold(), package.sequence)
        )
        if key in seen:
            return
        seen.add(key)
        ordered.append((package, reason))

    if cycle is not None and cycle.package_id is not None:
        linked = packages_by_id.get(cycle.package_id)
        if (
            linked is not None
            and linked.source == "rd"
            and not linked.is_grey
        ):
            reason = (
                cycle.match_reason
                if cycle.match_reason in _CYCLE_PACKAGE_REASONS
                else (cycle.match_reason or "revision")
            )
            cycle_reason = f"cycle_{reason}"
            exact_matches = [
                pkg
                for pkg in rd_packages
                if _package_match_tier(
                    pkg, kit_files, approved_revision_text
                )
                >= _REVISION_MATCH_EXACT
            ]
            if exact_matches:
                add(
                    max(
                        exact_matches,
                        key=lambda pkg: (
                            int(
                                _revision_match_tier(
                                    pkg.revision_text, approved_revision_text
                                )
                                >= _REVISION_MATCH_EXACT
                            ),
                            _current_package_sort_key(pkg),
                        ),
                    ),
                    cycle_reason,
                )
            else:
                add(linked, cycle_reason)

    if not rd_packages:
        return ordered

    folder_matches = [
        pkg
        for pkg in rd_packages
        if _package_has_folder_revision(pkg, kit_files, approved_revision_text)
    ]
    if folder_matches:
        add(
            max(
                folder_matches,
                key=lambda pkg: (
                    _package_folder_match_tier(
                        pkg, kit_files, approved_revision_text
                    ),
                    int(
                        _revision_match_tier(
                            pkg.revision_text, approved_revision_text
                        )
                        >= _REVISION_MATCH_EXACT
                    ),
                    _current_package_sort_key(pkg),
                ),
            ),
            "folder_rev",
        )

    file_matches = [
        pkg
        for pkg in rd_packages
        if pkg.revision_text
        and revision_texts_equivalent(pkg.revision_text, approved_revision_text)
    ]
    if file_matches:
        add(max(file_matches, key=_current_package_sort_key), "file_rev")

    if allow_last_package:
        add(max(rd_packages, key=_current_package_sort_key), "last_package")
    return ordered


def _last_code_or_any_cycle(
    cycles: Sequence[KitCycleRow],
    *,
    events_by_id: Mapping[int, KitEvent] | None = None,
) -> KitCycleRow | None:
    """Return the last code-A/B/C cycle, else the last cycle.

    Kept as the approved-contour wrapper around
    :func:`_last_cycle_for_stages`.
    """

    return _last_cycle_for_stages(
        cycles,
        stages=_CODE_STAGE_KEYS,
        events_by_id=events_by_id or {},
        fallback=True,
    )


def _last_cycle_for_stages(
    cycles: Sequence[KitCycleRow],
    *,
    stages: Collection[str],
    events_by_id: Mapping[int, KitEvent],
    fallback: bool,
) -> KitCycleRow | None:
    """Return the latest cycle whose linked F event is in ``stages``.

    Args:
        cycles: Kit cycles, typically in insert order.
        stages: F-event stages that count as an anchor.
        events_by_id: ``google_event.id`` → parsed event. A missing code
            event still counts when ``stages`` includes code A/B/C, so
            the approved contour stays stable if the event map is empty.
        fallback: When True and no cycle matches, return the last cycle
            (approved-contour behaviour). When False, return ``None``.

    Returns:
        The highest-id matching cycle, the last cycle when ``fallback``
        is set, or ``None``.
    """

    if not cycles:
        return None
    stages_set = frozenset(stages)
    anchored = [
        cycle
        for cycle in cycles
        if _cycle_matches_anchor(cycle, stages_set, events_by_id)
    ]
    pool = anchored or (list(cycles) if fallback else [])
    if not pool:
        return None
    return max(pool, key=lambda cycle: cycle.id or 0)


def _cycle_matches_anchor(
    cycle: KitCycleRow,
    stages: Collection[str],
    events_by_id: Mapping[int, KitEvent],
) -> bool:
    stages_set = frozenset(stages)
    wants_code = bool(stages_set & _CODE_STAGE_KEYS)
    if wants_code and cycle.code_event_id is not None:
        event = events_by_id.get(cycle.code_event_id)
        if event is None or event.stage in stages_set:
            return True
    for event_id in (cycle.tdo_event_id, cycle.code_event_id):
        if event_id is None:
            continue
        event = events_by_id.get(event_id)
        if event is not None and event.stage in stages_set:
            return True
    return False


def _anchor_revision_text(
    cycle: KitCycleRow | None,
    events: Sequence[tuple[int, KitEvent]],
    stages: Collection[str],
) -> str:
    """Return the journal revision for a non-code contour anchor."""

    if cycle is not None and (cycle.revision_text or "").strip():
        return cycle.revision_text.strip()
    stages_set = frozenset(stages)
    last: KitEvent | None = None
    for _event_id, event in events:
        if event.stage in stages_set:
            last = event
    if last is None:
        return ""
    return format_revision(last.revision, last.appendix)


def _appendix_only_digits(text: str) -> str:
    """Return AN digits when ``text`` is appendix-only (``AN01`` / ``рев.AN01``).

    ``parse_sheet_revision('AN01')`` would treat ``01`` as the base revision.

    Args:
        text: Folder name, ``рев.AN01`` token, or ``AN01``.

    Returns:
        Appendix digits, or empty string.
    """

    raw = normalize_unicode_dashes(text or "").strip()
    if not raw:
        return ""
    token_match = _APPENDIX_ONLY_TOKEN_RE.fullmatch(raw)
    if token_match is not None:
        return token_match.group("an")
    folder_match = _FOLDER_APPENDIX_ONLY_RE.search(raw)
    if folder_match is not None:
        return folder_match.group("an")
    return ""


def _revision_match_tier(candidate_text: str, approved_text: str) -> int:
    """Rank how ``candidate_text`` matches the approved revision.

    Args:
        candidate_text: Package filename revision, folder ``рев.*`` text,
            or an appendix-only token such as ``AN01``.
        approved_text: Journal / official approved revision text.

    Returns:
        ``2`` exact (base+appendix, or folder ``AN01`` vs approved
        ``04-AN01``), ``1`` same base with a different appendix, ``0``
        no match.
    """

    approved = (approved_text or "").strip()
    candidate = (candidate_text or "").strip()
    if not approved or not candidate:
        return 0
    appendix_only = _appendix_only_digits(candidate)
    if appendix_only:
        approved_rev, approved_app = parse_sheet_revision(approved)
        if not approved_rev or not approved_app:
            return 0
        candidate_rank = revision_rank(approved_rev, appendix_only)
        approved_rank = revision_rank(approved_rev, approved_app)
        if (
            candidate_rank[0] == approved_rank[0]
            and candidate_rank[1] == approved_rank[1]
        ):
            return _REVISION_MATCH_EXACT
        return 0
    if revision_texts_equivalent(candidate, approved):
        return _REVISION_MATCH_EXACT
    candidate_rev, _candidate_app = parse_sheet_revision(candidate)
    approved_rev, _approved_app = parse_sheet_revision(approved)
    if not candidate_rev or not approved_rev:
        return 0
    if revision_rank(candidate_rev, None)[0] == revision_rank(approved_rev, None)[0]:
        return _REVISION_MATCH_BASE
    return 0


def _package_folder_revision_texts(
    package: KitPackageRow,
    kit_files: Sequence[FileRecord],
) -> tuple[str, ...]:
    """Return folder ``рев.*`` texts for matching, including appendix-only names.

    Args:
        package: Issued RD package.
        kit_files: Present files of the same kit.

    Returns:
        Unique folder revision tokens (``04``, ``04-AN01``, ``AN01``, …).
    """

    texts: list[str] = []
    seen: set[str] = set()

    def add(text: str) -> None:
        value = (text or "").strip()
        if not value:
            return
        key = value.casefold()
        if key in seen:
            return
        seen.add(key)
        texts.append(value)

    names: list[str] = []
    if package.transfer_name:
        names.append(package.transfer_name)
    if package.package_path:
        names.append(PureWindowsPath(package.package_path).name)
    for record in kit_files:
        if not record.present or record.source is not SourceKind.RD:
            continue
        if not package.package_path or not path_is_under(
            record.path, package.package_path
        ):
            continue
        add(
            format_revision(
                _optional_text(record.data.get("transfer_revision")),
                _optional_text(record.data.get("transfer_appendix")),
            )
        )
        transfer_name = str(record.data.get("transfer_name") or "")
        if transfer_name:
            names.append(transfer_name)
    for name in names:
        digits = _appendix_only_digits(name)
        if digits:
            add(f"AN{digits}")
    return tuple(texts)


def _package_folder_match_tier(
    package: KitPackageRow,
    kit_files: Sequence[FileRecord],
    approved_revision_text: str,
) -> int:
    """Return the best folder-revision match tier for ``package``.

    Args:
        package: Issued RD package.
        kit_files: Present files of the same kit.
        approved_revision_text: Journal / official approved revision.

    Returns:
        Match tier from :func:`_revision_match_tier`, or ``0``.
    """

    if not approved_revision_text or not package.package_path:
        return 0
    return max(
        (
            _revision_match_tier(text, approved_revision_text)
            for text in _package_folder_revision_texts(package, kit_files)
        ),
        default=0,
    )


def _package_match_tier(
    package: KitPackageRow,
    kit_files: Sequence[FileRecord],
    approved_revision_text: str,
) -> int:
    """Return the best of filename-revision and folder-revision match tiers.

    Args:
        package: Issued RD package.
        kit_files: Present files of the same kit.
        approved_revision_text: Journal / official approved revision.

    Returns:
        The higher of the filename and folder match tiers.
    """

    return max(
        _revision_match_tier(package.revision_text, approved_revision_text),
        _package_folder_match_tier(package, kit_files, approved_revision_text),
    )


def _package_has_folder_revision(
    package: KitPackageRow,
    kit_files: Sequence[FileRecord],
    approved_revision_text: str,
) -> bool:
    """Return whether this package's transfer-folder revision matches.

    Folder ``рев.*`` must never be used as *the* displayed revision. Here it
    is only a fallback matching key to find which package the journal is
    talking about. Appendix-only names (``рев.AN01``) match an approved
    ``04-AN01`` as an exact appendix issuance of that base.
    """

    return _package_folder_match_tier(
        package, kit_files, approved_revision_text
    ) >= _REVISION_MATCH_EXACT


def _overlay_mto_as_of(
    kit_files: Sequence[FileRecord],
    *,
    max_sequence: int | None,
    rd_root: str | Path | None = None,
) -> ParsedFile | None:
    parsed = [
        parsed_file
        for record in kit_files
        if (parsed_file := _parsed_file_from_record(record)) is not None
    ]
    if not parsed:
        return None
    _pdf_overlay, mto_overlay = build_rd_overlays_as_of(
        parsed, max_sequence=max_sequence, rd_root=rd_root
    )
    del _pdf_overlay
    current = list(mto_overlay.current.values())
    if not current:
        return None
    return max(current, key=_newest_sort_key)


def _parsed_file_from_record(record: FileRecord) -> ParsedFile | None:
    data = record.data
    kind_raw = str(data.get("file_kind") or "").strip()
    try:
        file_kind = FileKind(kind_raw)
    except ValueError:
        return None
    status_raw = str(data.get("parse_status") or ParseStatus.PARSED.value)
    try:
        parse_status = ParseStatus(status_raw)
    except ValueError:
        parse_status = ParseStatus.PARSED
    transfer_name = str(data.get("transfer_name") or "")
    transfer_is_as_build = _as_bool(data.get("transfer_is_as_build"))
    transfer = None
    if transfer_name or transfer_is_as_build:
        transfer = TransferMetadata(
            original_name=transfer_name,
            normalized_name=transfer_name,
            sequence=_int_or_none(data.get("transfer_sequence")),
            revision=_optional_text(data.get("transfer_revision")),
            appendix=_optional_text(data.get("transfer_appendix")),
            title=_optional_text(data.get("title")),
            mark=_optional_text(data.get("mark")),
            title_system=_optional_text(data.get("title_system")),
            is_as_build=transfer_is_as_build,
            parse_status=ParseStatus.PARSED,
        )
    return ParsedFile(
        path=record.path,
        path_key=record.path_key,
        name=str(data.get("name") or PureWindowsPath(record.path).name),
        source=record.source,
        file_kind=file_kind,
        size=int(data.get("size") or 0),
        mtime_ns=int(data.get("mtime_ns") or 0),
        parse_status=parse_status,
        contract=_optional_text(data.get("contract")),
        title_system=_optional_text(data.get("title_system")),
        title=_optional_text(data.get("title")),
        mark=_optional_text(data.get("mark")),
        discipline_block=_optional_text(data.get("discipline_block")),
        core_stem=_optional_text(data.get("core_stem")),
        revision=_optional_text(data.get("revision")),
        appendix=_optional_text(data.get("appendix")),
        language=_optional_text(data.get("language")),
        extension=_optional_text(data.get("extension")),
        transfer=transfer,
        parse_error=_optional_text(data.get("parse_error")),
    )


def _optional_text(value: object) -> str | None:
    if value in (None, ""):
        return None
    return str(value)


def _records_include_annulled_folder(
    records: Sequence[FileRecord],
    annulled_by_kit: Mapping[tuple[str, str], Collection[str]],
) -> bool:
    for record in records:
        key = _record_kit_key(record)
        if key is None:
            continue
        names = annulled_by_kit.get(key, ())
        if names and _record_matches_folder_keys(record, names):
            return True
    return False


def _in_contour_collision_kinds(
    collisions: Sequence[dict],
    *,
    files_by_path: dict[str, FileRecord],
    kit_key: tuple[str, str],
    max_sequence: int | None,
    package_resolved: bool,
    annulled_folder_keys: Collection[str] = (),
) -> frozenset[str]:
    if not package_resolved:
        return frozenset()
    found: set[str] = set()
    annulled_by_kit = {kit_key: annulled_folder_keys}
    for collision in collisions:
        kind = str(collision.get("kind") or "")
        if kind not in _APPROVED_CONTOUR_COLLISIONS:
            continue
        path_keys = collision.get("path_keys") or ()
        if not path_keys:
            continue
        mapped: list[FileRecord] = []
        missing = False
        for path_key in path_keys:
            record = files_by_path.get(str(path_key).casefold())
            if record is None:
                missing = True
                break
            mapped.append(record)
        if missing or not mapped:
            continue
        if _records_include_annulled_folder(mapped, annulled_by_kit):
            continue
        if any(_record_kit_key(record) != kit_key for record in mapped):
            continue
        if not _all_sequences_in_contour(mapped, max_sequence):
            continue
        found.add(kind)
    return frozenset(found)


def _all_sequences_in_contour(
    records: Sequence[FileRecord],
    max_sequence: int | None,
) -> bool:
    if max_sequence is None:
        return True
    for record in records:
        sequence = _int_or_none(record.data.get("transfer_sequence"))
        if sequence is not None and sequence > max_sequence:
            return False
    return True


def _approved_ambiguity(
    *,
    package: KitPackageRow | None,
    mto_path: str,
    in_contour_kinds: frozenset[str],
) -> tuple[str, ...]:
    tokens: list[str] = []
    if CollisionKind.DUP_SAME_REVISION.value in in_contour_kinds:
        tokens.append("dup_same_revision")
    if CollisionKind.TRANSFER_ORDER_CONFLICT.value in in_contour_kinds:
        tokens.append("transfer_order_conflict")
    if not mto_path:
        tokens.append("no_mto_in_contour")
    if package is None:
        tokens.append("no_package")
        tokens.append("journal_only")
    return tuple(token for token in _AMBIGUITY_ORDER if token in tokens)


def _approved_confidence(
    *,
    package: KitPackageRow | None,
    mto_path: str,
    mto_source: str,
    ambiguity: tuple[str, ...],
    revision_mismatch: bool = False,
) -> str:
    conflicted = any(
        token in {"dup_same_revision", "transfer_order_conflict"}
        for token in ambiguity
    )
    if (
        package is None
        or not mto_path
        or conflicted
        or "no_package" in ambiguity
        or "no_mto_in_contour" in ambiguity
    ):
        return "low"
    if mto_source == "package":
        level = "high"
    else:
        level = "medium"
    if revision_mismatch and level == "high":
        return "medium"
    return level


def _approved_warnings(
    *,
    match_reason: str,
    mto_source: str,
    ambiguity: tuple[str, ...],
    approved_revision_text: str = "",
    package_revision_text: str = "",
    revision_mismatch: bool = False,
) -> tuple[str, ...]:
    notes: list[str] = []
    for token in ambiguity:
        text = _AMBIGUITY_WARNINGS.get(token)
        if text:
            notes.append(text)
    if mto_source == "inherited":
        notes.append(
            "Файл MTO унаследован из более ранней передачи, в согласованном пакете его нет."
        )
    if match_reason == "last_package":
        notes.append(
            "Пакет выбран как последняя передача РД — ревизия журнала не совпала с папкой и файлами."
        )
    if match_reason == "folder_rev":
        notes.append(
            "Пакет сопоставлен по ревизии папки передачи (не по ревизии имени файла)."
        )
    if revision_mismatch:
        approved = (approved_revision_text or "").strip() or "—"
        chosen = (package_revision_text or "").strip() or "—"
        notes.append(
            f"Согласована ревизия {approved}, а выбрана папка с ревизией {chosen}."
        )
    return tuple(notes)


def _mto_records_by_kit_revision(
    database: CatalogDatabase,
    records: Sequence[FileRecord],
    *,
    rd_root: str | Path | None = None,
) -> dict[tuple[str, str, tuple], list[FileRecord]]:
    """Index present RD MTO files by kit identity and filename revision.

    Args:
        database: Initialized catalog database.
        records: Catalog file records from the last scan.
        rd_root: RD source root. When set, non-canonical RD files are
            excluded via :func:`record_has_canonical_layout`.

    Returns:
        Mapping of ``(title_cf, mark_cf, revision_dedupe_key)`` to the present
        RD ``mto_xlsx`` records at that revision, deduplicated by ``path_key``.
    """

    files_by_path = _files_by_path_key(database, records, rd_root=rd_root)
    index: dict[tuple[str, str, tuple], list[FileRecord]] = {}
    seen: set[str] = set()
    for record in files_by_path.values():
        if not record.present or record.source is not SourceKind.RD:
            continue
        if _file_kind(record) != FileKind.MTO_XLSX.value:
            continue
        identity = _record_title_mark(record)
        if identity is None:
            continue
        revision_text = format_revision(*_file_revision(record))
        if not revision_text:
            continue
        stamp = record.path_key.casefold()
        if stamp in seen:
            continue
        seen.add(stamp)
        kit = kit_identity_key(*identity)
        key = (*kit, _revision_dedupe_key(revision_text))
        index.setdefault(key, []).append(record)
    return index


def iter_mto_files_for_cells(
    database: CatalogDatabase,
    *,
    records: Sequence[FileRecord],
    predicate: str,
    rd_root: str | Path | None = None,
) -> tuple[FileRecord, ...]:
    """Return present RD MTO files matching heatmap cells for ``predicate``.

    Paths come from ``records`` / ``file_entry``, never from GUI text.
    Working and annulled cells (``pipeline_status`` ``working`` /
    ``annulled``) are excluded from export iteration. Predicate
    ``current`` follows the official issued revision when overlay-current
    is a working revision.

    Args:
        database: Initialized catalog database.
        records: Catalog file records (typically the last scan).
        predicate: ``code_a``, ``tdo_passed`` (``tdo_review`` or any
            ``code_*``), ``current``, ``current_ifc``, or ``exclude_as_build``.
        rd_root: RD source root. When set, non-canonical RD files are
            excluded via :func:`record_has_canonical_layout`.

    Returns:
        Matching present RD ``mto_xlsx`` records.

    Raises:
        ValueError: If ``predicate`` is not supported.
    """

    if predicate not in _MTO_CELL_PREDICATES:
        raise ValueError(f"Unsupported MTO cell predicate: {predicate!r}")
    pipeline_by_kit: dict[tuple[str, str], KitPipelineRow] = {}
    if predicate == "current":
        pipeline_by_kit = {
            kit_identity_key(row.title, row.mark): row
            for row in database.list_kit_pipelines()
        }
    selected = [
        cell
        for cell in list_revision_matrix(database)
        if _cell_matches_export_predicate(
            cell,
            predicate,
            pipeline_by_kit.get(kit_identity_key(cell.title, cell.mark)),
        )
    ]
    wanted = [
        (kit_identity_key(cell.title, cell.mark), cell.revision_text)
        for cell in selected
    ]
    index = _mto_records_by_kit_revision(database, records, rd_root=rd_root)
    matched: list[FileRecord] = []
    seen: set[str] = set()
    for bucket in index.values():
        for record in bucket:
            identity = _record_title_mark(record)
            if identity is None:
                continue
            key = kit_identity_key(*identity)
            revision_text = format_revision(*_file_revision(record))
            if not any(
                kit == key and revision_texts_equivalent(revision_text, cell_rev)
                for kit, cell_rev in wanted
            ):
                continue
            stamp = record.path_key.casefold()
            if stamp in seen:
                continue
            seen.add(stamp)
            matched.append(record)
    matched.sort(key=lambda item: (item.path_key, item.id))
    return tuple(matched)


def rebuild_revision_matrix(
    database: CatalogDatabase,
    *,
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    kit_keys: Collection[tuple[str, str]] | None = None,
    rd_root: str | Path | None = None,
) -> None:
    """Rebuild ``kit_revision_cell`` from packages, F, sends, and collisions.

    Called from :func:`rebuild_pipeline` after ``replace_kit_derived`` so
    ``package_ids_json`` uses persisted SQLite ids. Does not delete
    ``kit_liquidity_review``.

    Args:
        database: Initialized catalog database.
        records: Catalog file records from the last scan.
        detected_current_ids: Overlay-current RD file ids.
        kit_keys: When set, replace heatmap rows only for these identities.
        rd_root: RD source root. When set, non-canonical RD files are
            excluded via :func:`record_has_canonical_layout`.
    """

    scoped = _normalize_kit_keys(kit_keys)
    with perf_span("pipeline.rebuild_revision_matrix", kits=len(scoped) if scoped is not None else "all"):
        packages = database.list_kit_packages()
        cycles = database.list_kit_cycles()
        pipelines = database.list_kit_pipelines()
        events_by_kit = database.list_google_events_by_kit()
        sends_with_ids = _effective_issuance_sends(database)
        collisions = database.list_current_collisions()
        google_kits = database.list_google_kits()
        files_by_path = _files_by_path_key(database, records, rd_root=rd_root)

        packages_by_kit: dict[tuple[str, str], list[KitPackageRow]] = {}
        for package in packages:
            packages_by_kit.setdefault(
                kit_identity_key(package.title, package.mark), []
            ).append(package)
        cycles_by_kit: dict[tuple[str, str], list[KitCycleRow]] = {}
        for cycle in cycles:
            cycles_by_kit.setdefault(
                kit_identity_key(cycle.title, cycle.mark), []
            ).append(cycle)
        sends_by_kit: dict[tuple[str, str], list[tuple[int | None, IssuanceKit]]] = {}
        for send_id, send in sends_with_ids:
            sends_by_kit.setdefault(
                kit_identity_key(send.title, send.mark), []
            ).append((send_id, send))
        pipeline_by_kit = {
            kit_identity_key(row.title, row.mark): row for row in pipelines
        }
        google_by_kit = {
            kit_identity_key(kit.title, kit.mark): kit for kit in google_kits
        }

        present_rd = [
            record
            for record in files_by_path.values()
            if record.present and record.source is SourceKind.RD
        ]
        present_rd_by_kit = index_records_by_kit(present_rd)
        identities: set[tuple[str, str]] = set()
        if scoped is not None:
            identities = set(scoped)
        else:
            identities.update(google_by_kit)
            identities.update(sends_by_kit)
            for package in packages:
                if package.source == "rd":
                    identities.add(kit_identity_key(package.title, package.mark))
            identities.update(present_rd_by_kit)

        display_names = _display_names(
            google_by_kit, sends_by_kit, packages, present_rd
        )
        problems, collision_revs = _collision_problems(
            collisions,
            files_by_path,
            {
                kit_identity_key(row.title, row.mark): row.annulled_transfer_names
                for row in pipelines
                if row.annulled_transfer_names
            },
        )
        if scoped is None:
            for kit_key in collision_revs:
                identities.add(kit_key)

        cells: list[KitRevisionRow] = []
        for kit_key in sorted(identities):
            title, mark = display_names.get(kit_key, (kit_key[0], kit_key[1]))
            kit_packages = packages_by_kit.get(kit_key, [])
            kit_cycles = cycles_by_kit.get(kit_key, [])
            kit_sends = sends_by_kit.get(kit_key, [])
            kit_events = events_by_kit.get(kit_key, [])
            pipeline = pipeline_by_kit.get(kit_key)
            kit_files = present_rd_by_kit.get(kit_key, [])
            revision_texts = _kit_revision_texts(
                kit_files=kit_files,
                packages=kit_packages,
                cycles=kit_cycles,
                sends=kit_sends,
                events=kit_events,
                pipeline=pipeline,
                extra=collision_revs.get(kit_key, {}),
            )
            working_text = pipeline.working_revision_text if pipeline else ""
            official_text = pipeline.official_revision_text if pipeline else ""
            suspicious = bool(pipeline and pipeline.suspicious)
            annulled_names = (
                pipeline.annulled_transfer_names if pipeline is not None else ()
            )
            live_files = [
                record
                for record in kit_files
                if not _record_matches_folder_keys(record, annulled_names)
            ]
            ifc_text = _highest_revision_text(
                [
                    record
                    for record in live_files
                    if not _record_is_as_build(record)
                ]
            )
            current_mto_revs = [
                format_revision(*_file_revision(record))
                for record in live_files
                if record.id in detected_current_ids
                and _file_kind(record) == FileKind.MTO_XLSX.value
            ]
            current_any_revs = [
                format_revision(*_file_revision(record))
                for record in live_files
                if record.id in detected_current_ids
            ]
            has_overlay_current_mto = any(bool(text) for text in current_mto_revs)
            ordered_sends = _sorted_sends(kit_sends)
            for rev_key, revision_text in revision_texts.items():
                files_this_rev = [
                    record
                    for record in kit_files
                    if revision_texts_equivalent(
                        format_revision(*_file_revision(record)), revision_text
                    )
                ]
                has_mto = any(
                    _file_kind(record) == FileKind.MTO_XLSX.value
                    for record in files_this_rev
                )
                has_rd_pdf = any(
                    _file_kind(record) == FileKind.PDF.value for record in files_this_rev
                )
                has_rd_package = any(
                    package.source == "rd"
                    and _package_matches_revision(package, revision_text)
                    for package in kit_packages
                )
                file_as_build = any(
                    _record_is_as_build(record) for record in files_this_rev
                )
                us_build = any(
                    event.stage == "us_build"
                    for _event_id, event in _events_on_revision(
                        kit_events, revision_text
                    )
                )
                cell_as_build = (
                    file_as_build
                    or us_build
                    or any(
                        package.is_as_build
                        and _package_matches_revision(package, revision_text)
                        for package in kit_packages
                    )
                )
                if _revision_is_fully_annulled(
                    revision_text, kit_packages, pipeline
                ):
                    status = "annulled"
                else:
                    status = _revision_pipeline_status(
                        revision_text=revision_text,
                        events=kit_events,
                        sends=ordered_sends,
                        cycles=kit_cycles,
                        working_revision=(
                            working_text
                            if _revision_is_fully_working(
                                revision_text, kit_packages, pipeline
                            )
                            else ""
                        ),
                        present_on_disk=bool(files_this_rev),
                        has_mto=has_mto,
                        has_rd_pdf_or_package=has_rd_pdf or has_rd_package,
                    )
                if status is None:
                    continue
                kind_list = list(problems.get((kit_key, rev_key), ()))
                if (
                    suspicious
                    and official_text
                    and revision_texts_equivalent(official_text, revision_text)
                ):
                    if "liquidity" not in kind_list:
                        kind_list.append("liquidity")
                package_ids = _package_ids_for_revision(kit_packages, revision_text)
                cells.append(
                    KitRevisionRow(
                        title=title,
                        mark=mark,
                        revision_text=revision_text,
                        pipeline_status=status,
                        letters=_revision_letters(status, cell_as_build),
                        is_as_build=cell_as_build,
                        is_current=_rev_in(revision_text, current_mto_revs)
                        or (
                            not has_overlay_current_mto
                            and _rev_in(revision_text, current_any_revs)
                        ),
                        is_current_ifc=bool(
                            ifc_text
                            and revision_texts_equivalent(ifc_text, revision_text)
                        ),
                        has_mto=has_mto,
                        problem_kinds_json=json.dumps(kind_list, ensure_ascii=False),
                        package_ids_json=json.dumps(package_ids, ensure_ascii=False),
                        algorithm_version=REVISION_MATRIX_ALGORITHM_VERSION,
                    )
                )
        database.replace_revision_cells(cells, identities=scoped)


def _files_by_path_key(
    database: CatalogDatabase,
    records: Sequence[FileRecord],
    *,
    rd_root: str | Path | None = None,
) -> dict[str, FileRecord]:
    """Index files by casefolded ``path_key``.

    When ``records`` is non-empty it is the only source (callers already
    pass the full catalog). An empty ``records`` falls back to
    :meth:`CatalogDatabase.list_files`.
    """

    source = records if records else database.list_files()
    files = {record.path_key.casefold(): record for record in source}
    root = str(rd_root or "").strip()
    if not root:
        return files
    return {
        key: record
        for key, record in files.items()
        if record_has_canonical_layout(record, root)
    }


def _cell_matches_export_predicate(
    cell: KitRevisionRow,
    predicate: str,
    pipeline: KitPipelineRow | None,
) -> bool:
    """True when ``cell`` should contribute MTO files for export ``predicate``."""

    if cell.pipeline_status in _SKIPPED_HEATMAP_STATUSES:
        return False
    if predicate == "current":
        working_text = pipeline.working_revision_text if pipeline is not None else ""
        official_text = (
            pipeline.official_revision_text if pipeline is not None else ""
        )
        uses_official = bool(working_text) or (
            pipeline is not None
            and (
                pipeline.annulled_transfer_names or pipeline.annulled_sequences
            )
        )
        if uses_official:
            return bool(
                official_text
                and revision_texts_equivalent(cell.revision_text, official_text)
            )
        return cell.is_current
    return _cell_matches_predicate(cell, predicate)


def _cell_matches_predicate(cell: KitRevisionRow, predicate: str) -> bool:
    if predicate == "code_a":
        return cell.pipeline_status == "code_a"
    if predicate == "tdo_passed":
        return cell.pipeline_status in _TDO_PASSED_STATUSES
    if predicate == "current":
        return cell.is_current
    if predicate == "current_ifc":
        return cell.is_current_ifc
    if predicate == "exclude_as_build":
        return not cell.is_as_build
    return False


def _remember_revision_text(store: dict[tuple, str], text: str) -> None:
    raw = (text or "").strip()
    if not raw:
        return
    store.setdefault(_revision_dedupe_key(raw), raw)


def _kit_revision_texts(
    *,
    kit_files: Sequence[FileRecord],
    packages: Sequence[KitPackageRow],
    cycles: Sequence[KitCycleRow],
    sends: Sequence[tuple[int | None, IssuanceKit]],
    events: Sequence[tuple[int, KitEvent]],
    pipeline: KitPipelineRow | None,
    extra: dict[tuple, str],
) -> dict[tuple, str]:
    collected: dict[tuple, str] = {}
    for record in kit_files:
        _remember_revision_text(
            collected, format_revision(*_file_revision(record))
        )
    for package in packages:
        _remember_revision_text(collected, package.revision_text)
        _remember_revision_text(collected, package.mto_revision_text)
    for cycle in cycles:
        _remember_revision_text(collected, cycle.revision_text)
    for _send_id, send in sends:
        _remember_revision_text(collected, send.revision_text)
    for _event_id, event in events:
        _remember_revision_text(
            collected, format_revision(event.revision, event.appendix)
        )
    if pipeline is not None:
        _remember_revision_text(collected, pipeline.working_revision_text)
        _remember_revision_text(collected, pipeline.official_revision_text)
    for text in extra.values():
        _remember_revision_text(collected, text)
    return collected


def _collision_problems(
    collisions: Sequence[dict],
    files_by_path: dict[str, FileRecord],
    annulled_by_kit: Mapping[tuple[str, str], Collection[str]] | None = None,
) -> tuple[
    dict[tuple[tuple[str, str], tuple], list[str]],
    dict[tuple[str, str], dict[tuple, str]],
]:
    problems: dict[tuple[tuple[str, str], tuple], list[str]] = {}
    extra_revs: dict[tuple[str, str], dict[tuple, str]] = {}
    skipped = annulled_by_kit or {}
    for collision in collisions:
        kind = str(collision.get("kind") or "")
        if kind not in _MATRIX_PROBLEM_KINDS:
            continue
        mapped: list[FileRecord] = []
        for path_key in collision.get("path_keys") or ():
            record = files_by_path.get(str(path_key).casefold())
            if record is not None:
                mapped.append(record)
        if skipped and _records_include_annulled_folder(mapped, skipped):
            continue
        for record in mapped:
            identity = _record_title_mark(record)
            if identity is None:
                continue
            revision_text = format_revision(*_file_revision(record))
            if not revision_text:
                continue
            kit_key = kit_identity_key(*identity)
            rev_key = _revision_dedupe_key(revision_text)
            extra_revs.setdefault(kit_key, {})[rev_key] = revision_text
            bucket = problems.setdefault((kit_key, rev_key), [])
            if kind not in bucket:
                bucket.append(kind)
    return problems, extra_revs


def _package_matches_revision(package: KitPackageRow, revision_text: str) -> bool:
    if package.revision_text and revision_texts_equivalent(
        package.revision_text, revision_text
    ):
        return True
    return bool(
        package.mto_revision_text
        and revision_texts_equivalent(package.mto_revision_text, revision_text)
    )


def _package_ids_for_revision(
    packages: Sequence[KitPackageRow], revision_text: str
) -> list[int]:
    rd_ids: list[int] = []
    other_ids: list[int] = []
    for package in packages:
        if package.id is None or not _package_matches_revision(
            package, revision_text
        ):
            continue
        if package.source == "rd":
            rd_ids.append(package.id)
        else:
            other_ids.append(package.id)
    return rd_ids + other_ids


def _rev_in(text: str, candidates: Sequence[str]) -> bool:
    return any(
        candidate and revision_texts_equivalent(text, candidate)
        for candidate in candidates
    )


def _revision_letters(status: str, is_as_build: bool) -> str:
    parts: list[str] = []
    letter = _REVISION_STATUS_LETTER.get(status, "")
    if letter:
        parts.append(letter)
    if is_as_build:
        parts.append(_AB_LETTER)
    return _LETTER_JOIN.join(parts)


def _revision_pipeline_status(
    *,
    revision_text: str,
    events: Sequence[tuple[int, KitEvent]],
    sends: Sequence[tuple[int | None, IssuanceKit]],
    cycles: Sequence[KitCycleRow],
    working_revision: str,
    present_on_disk: bool,
    has_mto: bool,
    has_rd_pdf_or_package: bool,
) -> str | None:
    revision_sends = [
        (send_id, send)
        for send_id, send in sends
        if revision_text
        and revision_texts_equivalent(send.revision_text, revision_text)
    ]
    last_send = revision_sends[-1][1] if revision_sends else None
    cycle_by_send = {
        cycle.send_id: cycle for cycle in cycles if cycle.send_id is not None
    }
    last_cycle = (
        cycle_by_send.get(revision_sends[-1][0]) if revision_sends else None
    )
    if last_cycle is None:
        revision_cycles = [
            cycle
            for cycle in cycles
            if revision_text
            and revision_texts_equivalent(cycle.revision_text, revision_text)
        ]
        last_cycle = revision_cycles[-1] if revision_cycles else None
    code_event = _last_code_on_revision(events, revision_text, last_cycle)
    if code_event is not None and code_event.stage in _CODE_STAGES:
        _letter, status_value = _CODE_STAGES[code_event.stage]
        return status_value
    status = _review_status(
        official_revision=revision_text,
        last_send=last_send,
        last_cycle=last_cycle,
        events=events,
        has_official_send=bool(revision_sends),
        allow_agreed=False,
    )
    if status is not KitPipelineStatus.NOT_UPLOADED:
        return status.value
    if (
        working_revision
        and revision_texts_equivalent(working_revision, revision_text)
        and present_on_disk
    ):
        return "working"
    if has_mto:
        return "not_uploaded"
    if has_rd_pdf_or_package:
        return "no_mto"
    return None

