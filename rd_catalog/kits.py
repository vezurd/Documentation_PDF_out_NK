"""Expected RD kits from Google Sheets and presence matrix vs scan sources.

Contract
--------
- Key is ``(title, mark)`` / ``title_system`` (``####-MARK``), case-insensitive.
- Column B: strip ``_IFC`` / ``_IFR``; keep Latin AGCC marks only; denylist
  activity rows (interfaces, MDZ/TO/ID, comments).
- Column D: optional sheet revision; mismatch vs RD is highlight-only.
- Column F is the event source of truth; column E is stored as ``status_sheet``.
- Matrix columns: RD (detected_current), ROBOT (present MTO), Google, SQ
  (display-only answers). Missing SQ is not a kit status. SQ never joins
  the RD overlay.
- Kit ``revision_text`` (Комплекты «РД · рев.») is overlay-current **OD**
  when present, else MTO. Painted «MTO · рев.» is the MTO xlsx **in the
  official issued folder** (filename may lag OD). Empty folder → ``нет``;
  older NN are not searched. Snapshot ``mto_revision_text`` may still
  keep overlay-current MTO from an older package (robot origin / matrix).
  WIR/LAY and folder ``рев.*`` do not set either. Folders are sequenced
  only by the leading ``NN``. Snapshot ``paths`` are newest-first.
- Robot origin highlight (kits GUI): compare **MTO file mtimes** of robot vs
  RD and SQ, plus cached RD↔robot content equality. Close = identical mtime,
  same local calendar day, or ``|Δ| ≤ 3 days`` (transfer packaging can make
  the robot copy slightly newer or older). Origin green also needs the
  robot filename rev to match «MTO · рев.» of the official current folder
  (not «РД · рев.» / OD). Content equality without that MTO rev match is
  **bold** only, not origin green. Red = robot matches neither by date.
  This paint overrides revision-vs-MTO green.
- Transfer-review table roles: ``текущая`` = official NN of the agreed
  filename rev (highest non-working / non-annulled); ``другая`` /
  ``спорная`` = other **non-working and non-annulled** packages in that
  compare; ``рабочая`` = marked working or ahead of send; ``аннулирована``
  = marked annulled — shown, but not current, not «спорная», and not in
  MTO сверка. The package table and MTO сверка list **editable** MTO/OD
  (xlsx/xls/dwg/doc), not PDF copies. OD PDF stays only when a disputed
  kit has no editable MTO/OD at all.
- Never write to the Google sheet or UNC sources.
"""

from __future__ import annotations

import re
from collections.abc import Collection
from dataclasses import dataclass
from datetime import date, datetime
from enum import StrEnum
from pathlib import Path, PureWindowsPath
from typing import Any, Callable, Iterable, Mapping, Sequence

from rd_catalog.models import (
    FileKind,
    FileRecord,
    MtoContentStatus,
    ParseStatus,
    SourceKind,
    make_path_key,
)
from rd_catalog.overlay import revision_rank
from rd_catalog.parse import (
    constructed_kit_rd_title_folder,
    is_package_media_folder,
    is_transfer_gate_folder_name,
    issued_package_dir,
    issued_path_title_folder,
    kit_rd_mark_folder_from_path,
    nearest_issued_package_hint,
    normalize_unicode_dashes,
    path_is_as_build,
    path_relative_to_transfer_gate,
    record_has_canonical_layout,
)
from rd_catalog.path_actions import containing_folder, path_is_under
from rd_catalog.perf_log import perf_span

KIT_MATRIX_ALGORITHM_VERSION = 4
_NS = 1_000_000_000
ROBOT_MTIME_CLOSE_NS = 3 * 24 * 3600 * _NS
_MIXED_TITLE_NOTE_LIMIT = 6

_TITLE_RE = re.compile(r"^\d{4}$")
_MARK_SUFFIX_RE = re.compile(r"_IF[CR]$", re.IGNORECASE)
_MARK_KEEP_RE = re.compile(r"^[A-Za-z]{2,5}\d{0,2}(?:\.\d{1,2})?$")
_MARK_DENYLIST = frozenset(
    {
        "интерфейс",
        "расст. камер",
        "расст камер",
        "комментарии",
        "nanocad",
        "то",
        "мдз",
        "ид",
    }
)
_MARK_TRANSLIT = {
    "ксб": "KSB",
    "пд": "PD",
}

_REV_PREFIX_RE = re.compile(r"^(?:рев\.?|revision)\s*", re.IGNORECASE)
_REV_BODY_RE = re.compile(
    r"(?P<rev>\d{1,2}|[VS])(?:-AN(?P<an>\d{1,2}))?",
    re.IGNORECASE,
)
_LETTER_REV_RE = re.compile(r"^[A-Z]$", re.IGNORECASE)
_DATE_RE = re.compile(r"^(\d{1,2}\.\d{1,2}\.\d{4})\s*(.*)$")
_TRM_RE = re.compile(r"[A-Z0-9]+(?:[.-][A-Z0-9]+)*-TRM-\d+", re.IGNORECASE)
_REV_IN_TEXT_RE = re.compile(
    r"рев\.?\s*(\d{1,2}(?:-AN\d{1,2})?|[VS]|[A-Z])",
    re.IGNORECASE,
)

_STAGE_PATTERNS: tuple[tuple[str, str, re.Pattern[str]], ...] = (
    ("code_a", "код А", re.compile(r"код\s*[аa]", re.IGNORECASE)),
    # Latin C / Cyrillic Es before code_b so «код C/С» is not eaten by [бbв].
    ("code_c", "код C", re.compile(r"код\s*[cс](?![a-zа-яё])", re.IGNORECASE)),
    ("code_b", "код B", re.compile(r"код\s*[бbв]", re.IGNORECASE)),
    (
        "sr_upload",
        "загрузка в СР",
        re.compile(
            r"загрузк\w*.{0,20}ср|отпр\w*.{0,20}на\s+загр",
            re.IGNORECASE,
        ),
    ),
    ("tdo_passed", "прошла ТДО", re.compile(r"прош\w*\s+тдо", re.IGNORECASE)),
    ("tdo_sent", "отпр. на ТДО", re.compile(r"отпр\w*\s+на\s+тдо", re.IGNORECASE)),
    (
        "incoming_sent",
        "отпр. на входной контроль",
        re.compile(r"отпр\w*.{0,20}(вх|входн)", re.IGNORECASE),
    ),
    (
        "incoming_passed",
        "входной контроль",
        re.compile(
            r"прош\w*.{0,16}(вх|входн)|вх\s*контр|входн\w*\s+контр",
            re.IGNORECASE,
        ),
    ),
    ("dup", "ДУП", re.compile(r"\bдуп\b", re.IGNORECASE)),
    ("correction", "корректировка", re.compile(r"корректир", re.IGNORECASE)),
    ("agreed", "согласовано", re.compile(r"соглас", re.IGNORECASE)),
    (
        "us_build",
        "US-BUILD",
        re.compile(r"us[\s_.-]*build|as[\s_.-]*built?", re.IGNORECASE),
    ),
)


class KitSummary(StrEnum):
    """Primary row status for the комплекты matrix."""

    ALIGNED = "aligned"
    REV_MISMATCH = "rev_mismatch"
    TRANSFER_REVIEW = "transfer_review"
    MIXED_TITLES = "mixed_titles"
    GAP_RD = "gap_rd"
    GAP_ROBOT = "gap_robot"
    GAP_SQ = "gap_sq"
    EXTRA_RD = "extra_rd"
    EXTRA_ROBOT = "extra_robot"
    EXTRA_SQ = "extra_sq"
    GOOGLE_ONLY = "google_only"


class KitFlag(StrEnum):
    """Independent presence / revision flags on a matrix row."""

    ALIGNED = "aligned"
    REV_MISMATCH = "rev_mismatch"
    TRANSFER_REVIEW = "transfer_review"
    MIXED_TITLES = "mixed_titles"
    GAP_RD = "gap_rd"
    GAP_ROBOT = "gap_robot"
    GAP_SQ = "gap_sq"
    EXTRA_RD = "extra_rd"
    EXTRA_ROBOT = "extra_robot"
    EXTRA_SQ = "extra_sq"
    GOOGLE_ONLY = "google_only"


_SUMMARY_PRIORITY: tuple[KitSummary, ...] = (
    KitSummary.GOOGLE_ONLY,
    KitSummary.GAP_RD,
    KitSummary.MIXED_TITLES,
    KitSummary.TRANSFER_REVIEW,
    KitSummary.REV_MISMATCH,
    KitSummary.EXTRA_RD,
    KitSummary.GAP_ROBOT,
    KitSummary.EXTRA_ROBOT,
    KitSummary.EXTRA_SQ,
    KitSummary.ALIGNED,
)

_SUMMARY_LABELS: dict[KitSummary, str] = {
    KitSummary.ALIGNED: "Совпадает",
    KitSummary.REV_MISMATCH: "Ревизия расходится",
    KitSummary.TRANSFER_REVIEW: "Проверить передачи",
    KitSummary.MIXED_TITLES: "Смешанные титулы",
    KitSummary.GAP_RD: "Нет в РД",
    KitSummary.GAP_ROBOT: "Нет у робота",
    KitSummary.GAP_SQ: "Нет в SQ",
    KitSummary.EXTRA_RD: "Лишнее в РД",
    KitSummary.EXTRA_ROBOT: "Лишнее у робота",
    KitSummary.EXTRA_SQ: "Лишнее в SQ",
    KitSummary.GOOGLE_ONLY: "Только Google",
}

_SUMMARY_HINTS: dict[KitSummary, str] = {
    KitSummary.ALIGNED: (
        "Комплект есть в Google/Выдаче и в РД, ревизии Google↔РД совпадают, "
        "инверсии MTO/OD нет. Расхождение робота, Выдачи или Google F "
        "с РД только подсвечивает ячейку ревизии."
    ),
    KitSummary.REV_MISMATCH: (
        "Ревизия Google (если пусто — Выдачи) не совпадает с ревизией файлов РД. "
        "На готовность MTO это не влияет."
    ),
    KitSummary.TRANSFER_REVIEW: (
        "Поздняя папка NN стала текущей, хотя в более ранней папке тот же "
        "OD/MTO новее по ревизии файла или по дате; либо официальная NN "
        "не ближе к письму A / отправке (MTO в ней позже кода A). "
        "Текущий состав при этом не откатывается — проверьте, "
        "какая передача ушла заказчику. "
        "В таблице — редактируемые MTO/OD тех же пакетов (xlsx/xls/dwg), "
        "без PDF-копий; % близости к Выдаче/F/письму A "
        "(как «К согл. передаче» на АН) и сверка содержимого MTO "
        "спорных папок (четкое / ПоКоду и Кол-ву / не совпало). "
        "PDF MTO в сверку не берётся. OD в PDF — только если в спорных "
        "пакетах нет редактируемого файла. "
        "WIR/LAY на эту сводку не влияют. Папка открытия остаётся "
        "официальной NN. Рабочая папка — роль «рабочая»: не текущая, "
        "не «спорная» и не в сверке."
    ),
    KitSummary.MIXED_TITLES: (
        "У комплекта есть present-файлы РД, чей путь лежит в папке "
        "другого четырёхзначного титула (имя AGCC — этот комплект, "
        "сегмент пути — чужой титул). Файлы не отбрасываются. "
        "«Папка РД» может открыть чужое дерево, если такой пакет "
        "стал официальным NN."
    ),
    KitSummary.GAP_RD: (
        "Комплект есть в Google или Выдаче, официальной ревизии нет в РД "
        "(нет файлов в папке передачи или на диске другая ревизия)."
    ),
    KitSummary.GAP_ROBOT: (
        "Комплект есть в Google/Выдаче, у робота нет MTO. "
        "Пустой SQ на сводку не влияет."
    ),
    KitSummary.GAP_SQ: (
        "В SQ нет файлов этого комплекта. Обычно не показывается как сводка: "
        "пустой SQ — норма."
    ),
    KitSummary.EXTRA_RD: (
        "Файлы есть в РД, строки комплекта нет ни в Google, ни в Выдаче."
    ),
    KitSummary.EXTRA_ROBOT: (
        "MTO есть у робота, строки комплекта нет ни в Google, ни в Выдаче."
    ),
    KitSummary.EXTRA_SQ: (
        "Файлы есть в SQ, строки комплекта нет ни в Google, ни в Выдаче."
    ),
    KitSummary.GOOGLE_ONLY: (
        "Комплект есть в Google или Выдаче, файлов нет ни в РД, ни у робота, ни в SQ."
    ),
}


@dataclass(frozen=True, slots=True)
class KitEvent:
    """One parsed line from Google column F."""

    raw: str
    date: str | None
    stage: str
    stage_label: str
    revision: str | None
    appendix: str | None
    transmittals: tuple[str, ...]
    parsed: bool


@dataclass(frozen=True, slots=True)
class GoogleKit:
    """One expected kit row after mark filtering."""

    title: str
    mark: str
    mark_raw: str
    title_system: str
    sheet_revision: str | None
    sheet_appendix: str | None
    sheet_revision_text: str
    status_sheet: str
    comment_raw: str
    events: tuple[KitEvent, ...]
    last_event: KitEvent | None
    row_index: int


@dataclass(frozen=True, slots=True)
class IssuanceKit:
    """One send row from «Выдача РД ПД»."""

    title: str
    mark: str
    mark_raw: str
    title_system: str
    revision: str | None
    appendix: str | None
    revision_text: str
    status: str
    send_date: str
    send_date_sortable: str
    send_transmittal: str
    incoming_control_date: str
    incoming_control_date_sortable: str
    confirm_transmittal: str
    note_raw: str
    row_index: int


@dataclass(frozen=True, slots=True)
class SourceKitSnapshot:
    """Aggregated presence of one scan source for a title+mark kit.

    ``paths`` / ``file_ids`` are newest-first: highest transfer sequence,
    then filename revision, PDF before MTO xlsx, then mtime. Context-menu
    open/copy uses ``paths[0]``. Kits / export pass official-current ids
    (working package excluded, including leftover files named as the
    official rev), so ``paths[0]`` is the issued official folder.
    """

    present: bool = False
    revision: str | None = None
    appendix: str | None = None
    revision_text: str = ""
    mto_revision: str | None = None
    mto_appendix: str | None = None
    mto_revision_text: str = ""
    max_mtime_ns: int | None = None
    mto_mtime_ns: int | None = None
    file_count: int = 0
    as_build: bool = False
    transfer_name: str | None = None
    file_ids: tuple[int, ...] = ()
    paths: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class RobotOrigin:
    """Kits-table highlight plan for the robot revision cell.

    ``matched`` is ``True`` (origin green, counterpart RD/SQ), ``False``
    (red orphan), or ``None`` when there is no robot file **or** date/content
    is close but the filename revision does not match «MTO · рев.» of the
    official current folder / SQ.
    ``content_equal`` is independent: bold the robot cell even when origin
    green is withheld.
    """

    matched: bool | None
    rd: bool = False
    sq: bool = False
    reason: str = ""
    content_equal: bool = False


@dataclass(frozen=True, slots=True)
class TransferReviewMtoPair:
    """Two RD MTO workbooks from disputed transfer packages of one kit."""

    left_path: str
    right_path: str
    left_mtime_ns: int = 0
    right_mtime_ns: int = 0


@dataclass(frozen=True, slots=True)
class KitMatrixRow:
    """One union-matrix row across RD / robot / Google sheets / SQ."""

    title: str
    mark: str
    title_system: str
    rd: SourceKitSnapshot
    robot: SourceKitSnapshot
    sq: SourceKitSnapshot
    google: GoogleKit | None
    issuance: IssuanceKit | None
    flags: tuple[KitFlag, ...]
    summary: KitSummary
    transfer_review_notes: tuple[str, ...] = ()
    transfer_review_mto_pairs: tuple[TransferReviewMtoPair, ...] = ()
    mixed_title_notes: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class GoogleParseStats:
    """Counts of kept vs skipped sheet rows."""

    kept: int
    skipped: int
    skipped_reasons: tuple[str, ...]


def kit_identity_key(title: str, mark: str) -> tuple[str, str]:
    """Return the case-insensitive kit identity.

    Args:
        title: Four-digit title.
        mark: Latin AGCC mark.

    Returns:
        ``(title.casefold(), mark.casefold())``.
    """

    return title.casefold(), mark.casefold()


def google_kit_sheet_fingerprint(kit: GoogleKit) -> tuple[str, str, str]:
    """Return KSB ИД F/D/E text that can change a kit pipeline.

    Args:
        kit: One persisted or freshly fetched Google kit row.

    Returns:
        ``(comment_raw, sheet_revision_text, status_sheet)``.
    """

    return (kit.comment_raw, kit.sheet_revision_text, kit.status_sheet)


def changed_google_kit_keys(
    before: Sequence[GoogleKit],
    after: Sequence[GoogleKit],
) -> set[tuple[str, str]]:
    """Return kit identities whose KSB ИД F/D/E text appeared or changed.

    Used after a letter write refetch so ``rebuild_pipeline`` can stay
    scoped to kits that actually moved, including concurrent sheet edits.

    Args:
        before: In-memory kits from the previous Google snapshot.
        after: Kits from the post-write КСБ ИД fetch.

    Returns:
        ``kit_identity_key`` values that are new, removed, or whose
        fingerprint changed. Empty when both sides match.
    """

    previous = {
        kit_identity_key(kit.title, kit.mark): google_kit_sheet_fingerprint(kit)
        for kit in before
    }
    current = {
        kit_identity_key(kit.title, kit.mark): google_kit_sheet_fingerprint(kit)
        for kit in after
    }
    changed = {
        key for key, fingerprint in current.items() if previous.get(key) != fingerprint
    }
    changed.update(key for key in previous if key not in current)
    return changed


def outlook_od_search_query(title: str, mark: str) -> str:
    """Return a quoted Outlook Instant Search token such as ``"1600-SOT.OD"``.

    Args:
        title: Four-digit title.
        mark: Latin AGCC mark.

    Returns:
        Quoted ``title-mark.OD``, or empty when either part is blank.
    """

    title_text = (title or "").strip()
    mark_text = (mark or "").strip()
    if not title_text or not mark_text:
        return ""
    return f'"{title_text}-{mark_text}.OD"'


def format_revision(revision: str | None, appendix: str | None) -> str:
    """Format a revision/AN pair as ``01-AN02``.

    Args:
        revision: Revision token.
        appendix: Optional AN digits.

    Returns:
        Display text, or empty string when revision is missing.
    """

    if not revision:
        return ""
    return f"{revision}-AN{appendix}" if appendix else str(revision)


def summary_label(summary: KitSummary) -> str:
    """Return the Russian label for a matrix summary.

    Args:
        summary: Primary row status.

    Returns:
        Short GUI label.
    """

    return _SUMMARY_LABELS[summary]


def summary_tooltip(row: KitMatrixRow) -> str:
    """Return hover text for the комплекты summary cell.

    Args:
        row: One matrix row.

    Returns:
        Status explanation, plus transfer-review package notes when present.
    """

    lines = [_SUMMARY_HINTS.get(row.summary, summary_label(row.summary))]
    if row.mixed_title_notes:
        lines.append("")
        lines.extend(row.mixed_title_notes)
    if row.transfer_review_notes:
        lines.append("")
        lines.extend(row.transfer_review_notes)
    return "\n".join(lines)


def strip_mark_suffix(raw_mark: str) -> str:
    """Remove ``_IFC`` / ``_IFR`` from a sheet mark cell.

    Args:
        raw_mark: Column B text.

    Returns:
        Mark without the IFC/IFR suffix.
    """

    return _MARK_SUFFIX_RE.sub("", (raw_mark or "").strip())


def normalize_mark_token(raw_mark: str) -> str:
    """Strip IFC/IFR and transliterate known Cyrillic mark aliases.

    Args:
        raw_mark: Sheet mark cell.

    Returns:
        Normalized Latin mark token.
    """

    mark = strip_mark_suffix(raw_mark)
    mapped = _MARK_TRANSLIT.get(mark.casefold())
    return mapped if mapped else mark


def is_rd_kit_mark(mark: str) -> bool:
    """Return whether ``mark`` is a Latin AGCC kit mark.

    Args:
        mark: Mark after suffix stripping / transliteration.

    Returns:
        True when the mark should participate in the matrix.
    """

    cleaned = (mark or "").strip()
    if not cleaned or cleaned.casefold() in _MARK_DENYLIST:
        return False
    if "3d" in cleaned.casefold():
        return False
    return bool(_MARK_KEEP_RE.fullmatch(cleaned))


def normalize_revision_digits(revision: str | None) -> str | None:
    """Pad numeric revisions: ``1`` → ``01``; keep bare ``0``.

    Args:
        revision: Raw revision token.

    Returns:
        Normalized revision or ``None``.
    """

    if revision is None:
        return None
    text = str(revision).strip()
    if not text:
        return None
    if text.isdigit():
        value = int(text)
        if value == 0:
            return "0"
        return f"{value:02d}"
    return text


def parse_sheet_revision(text: str) -> tuple[str | None, str | None]:
    """Parse a revision cell into ``(revision, appendix)``.

    Args:
        text: Sheet cell such as ``Рев. 01-AN02``, ``1``, ``0-AN02``,
            or an old IFR letter ``A``.

    Returns:
        Normalized revision tokens, or ``(None, None)`` when unparseable.
    """

    raw = normalize_unicode_dashes(text or "").strip()
    if not raw:
        return None, None
    raw = raw.replace("АН", "AN").replace("ан", "AN")
    raw = _REV_PREFIX_RE.sub("", raw).strip()
    match = _REV_BODY_RE.search(raw)
    if match:
        revision = normalize_revision_digits(match.group("rev"))
        appendix = match.group("an")
        return revision, appendix
    if _LETTER_REV_RE.fullmatch(raw):
        return raw.upper(), None
    return None, None


def extract_confirm_transmittal(note: str) -> str:
    """Extract the first TRM token from an issuance note cell.

    Args:
        note: «Примечание» cell.

    Returns:
        Transmittal id or empty string.
    """

    match = _TRM_RE.search(normalize_unicode_dashes(note or ""))
    return match.group(0) if match else ""


def revisions_equivalent(
    left_revision: str | None,
    left_appendix: str | None,
    right_revision: str | None,
    right_appendix: str | None,
) -> bool:
    """Compare two revision/AN pairs ignoring zero-padding.

    Args:
        left_revision: First revision token.
        left_appendix: First AN token.
        right_revision: Second revision token.
        right_appendix: Second AN token.

    Returns:
        True when both sides are present and numerically equivalent.
    """

    if not left_revision or not right_revision:
        return False
    left = revision_rank(left_revision, left_appendix)
    right = revision_rank(right_revision, right_appendix)
    if left[0] < 0 or right[0] < 0:
        return (
            left_revision.casefold() == right_revision.casefold()
            and (left_appendix or "") == (right_appendix or "")
        )
    return left[0] == right[0] and left[1] == right[1]


def revision_matches_rd(
    rd_revision: str | None,
    rd_appendix: str | None,
    other_revision: str | None,
    other_appendix: str | None,
) -> bool | None:
    """Compare one source revision to the RD revision.

    Args:
        rd_revision: RD revision token.
        rd_appendix: RD AN token.
        other_revision: Other source revision token.
        other_appendix: Other source AN token.

    Returns:
        ``True`` when both sides have a revision and they match, ``False``
        when both have a revision and they differ, ``None`` when either
        side is missing.
    """

    if not rd_revision or not other_revision:
        return None
    return revisions_equivalent(
        rd_revision, rd_appendix, other_revision, other_appendix
    )


def kit_revision_match_flags(row: KitMatrixRow) -> dict[str, bool | None]:
    """Return match / mismatch / None for each kits revision source.

    Args:
        row: One комплекты matrix row.

    Returns:
        Keys ``google``, ``rd``, ``rd_mto``, ``sq``, ``robot``,
        ``google_f``, ``issuance``. Google / issuance / SQ stay vs
        «РД · рев.» (usually OD). MTO columns (``rd_mto``, ``robot``)
        compare to overlay-current MTO when it exists.
    """

    rd_rev = row.rd.revision if row.rd.present else None
    rd_app = row.rd.appendix if row.rd.present else None
    mto_rev = row.rd.mto_revision if row.rd.present else None
    mto_app = row.rd.mto_appendix if row.rd.present else None

    def vs_rd(revision: str | None, appendix: str | None) -> bool | None:
        return revision_matches_rd(rd_rev, rd_app, revision, appendix)

    def vs_mto(revision: str | None, appendix: str | None) -> bool | None:
        if mto_rev:
            return revision_matches_rd(mto_rev, mto_app, revision, appendix)
        return vs_rd(revision, appendix)

    google = row.google
    event = google.last_event if google is not None else None
    issuance = row.issuance
    return {
        "google": vs_rd(
            google.sheet_revision if google is not None else None,
            google.sheet_appendix if google is not None else None,
        ),
        "rd": vs_rd(rd_rev, rd_app),
        "rd_mto": vs_mto(mto_rev, mto_app),
        "sq": vs_rd(
            row.sq.revision if row.sq.present else None,
            row.sq.appendix if row.sq.present else None,
        ),
        "robot": vs_mto(
            row.robot.revision if row.robot.present else None,
            row.robot.appendix if row.robot.present else None,
        ),
        "google_f": vs_rd(
            event.revision if event is not None else None,
            event.appendix if event is not None else None,
        ),
        "issuance": vs_rd(
            issuance.revision if issuance is not None else None,
            issuance.appendix if issuance is not None else None,
        ),
    }


def _datetime_from_mtime_ns(mtime_ns: int) -> datetime | None:
    try:
        return datetime.fromtimestamp(int(mtime_ns) / _NS)
    except (OSError, OverflowError, ValueError):
        return None


def _format_mtime_delta(delta_ns: int) -> str:
    seconds = abs(int(delta_ns)) / _NS
    if seconds < 90:
        return "mtime совпадает"
    if seconds < 3600:
        return f"Δ {max(1, int(seconds // 60))} мин"
    if seconds < 86400:
        hours = seconds / 3600
        if hours < 10:
            return f"Δ {hours:.1f} ч"
        return f"Δ {int(hours)} ч"
    days = seconds / 86400
    if days < 1.5:
        return "Δ 1 сут"
    return f"Δ {int(round(days))} сут"


def _snapshot_mto_mtime_ns(snapshot: SourceKitSnapshot, *, allow_max: bool) -> int | None:
    if not snapshot.present:
        return None
    if snapshot.mto_mtime_ns:
        return int(snapshot.mto_mtime_ns)
    if allow_max and snapshot.max_mtime_ns:
        return int(snapshot.max_mtime_ns)
    return None


def _snapshot_mto_rev_pair(
    snapshot: SourceKitSnapshot,
) -> tuple[str | None, str | None, str]:
    """Return MTO filename tokens from the official/current source folder.

    Prefer ``mto_revision`` (Комплекты «MTO · рев.»). OD-first
    ``revision`` is only a fallback when the snapshot has no MTO file.
    """

    if snapshot.mto_revision:
        text = snapshot.mto_revision_text or format_revision(
            snapshot.mto_revision, snapshot.mto_appendix
        )
        return snapshot.mto_revision, snapshot.mto_appendix, text
    text = snapshot.revision_text or format_revision(
        snapshot.revision, snapshot.appendix
    )
    return snapshot.revision, snapshot.appendix, text


def _robot_rev_matches(
    robot: SourceKitSnapshot, other: SourceKitSnapshot
) -> bool:
    """Return whether robot MTO rev matches the counterpart folder's MTO rev."""

    if not robot.present or not other.present:
        return False
    other_rev, other_app, _text = _snapshot_mto_rev_pair(other)
    if not other_rev:
        return False
    robot_rev, robot_app, _robot_text = _snapshot_mto_rev_pair(robot)
    return revisions_equivalent(robot_rev, robot_app, other_rev, other_app)


def _robot_date_score(robot_ns: int | None, other_ns: int | None) -> int:
    """Rank date closeness: 0 none, 1 within 3 days, 2 same local day, 3 identical."""

    if not robot_ns or not other_ns:
        return 0
    delta = abs(int(robot_ns) - int(other_ns))
    if delta == 0:
        return 3
    robot_at = _datetime_from_mtime_ns(int(robot_ns))
    other_at = _datetime_from_mtime_ns(int(other_ns))
    if robot_at is not None and other_at is not None and robot_at.date() == other_at.date():
        return 2
    if delta <= ROBOT_MTIME_CLOSE_NS:
        return 1
    return 0


def mto_content_equal_by_kit(
    mto_rows: Iterable[Mapping[str, Any]],
) -> dict[tuple[str, str], bool]:
    """Map kit keys to cached RD↔robot MTO content equality.

    Args:
        mto_rows: GUI comparison dicts (``diff.content_status``).

    Returns:
        ``True`` when any pair for the kit is content-equal, ``False`` when
        the kit was compared and none are equal. Missing key = not compared
        (empty or ``not_compared`` content status is skipped).
    """

    result: dict[tuple[str, str], bool] = {}
    expected = MtoContentStatus.EQUAL.value
    for item in mto_rows:
        title = str(item.get("title") or "").strip()
        mark = str(item.get("mark") or "").strip()
        if not title or not mark:
            continue
        payload = item.get("diff")
        if not isinstance(payload, Mapping):
            payload = {}
        status = str(payload.get("content_status") or "").strip()
        if not status or status == MtoContentStatus.NOT_COMPARED.value:
            continue
        key = kit_identity_key(title, mark)
        if status == expected:
            result[key] = True
        else:
            result.setdefault(key, False)
    return result


def kit_robot_origin(
    row: KitMatrixRow,
    *,
    rd_content_equal: bool | None = None,
) -> RobotOrigin:
    """Decide robot-origin highlighting vs RD / SQ MTO dates and RD content.

    Date window: identical mtime, same local calendar day, or ``|Δ| ≤ 3 days``.
    RD/SQ dates are the MTO xlsx mtime (not PDF). Cached content equality with
    current RD MTO counts as an RD match even when dates drifted, **but only
    when the robot filename rev matches «MTO · рев.» of that official
    folder** (OD / «РД · рев.» is a different file and may differ). Content
    equality with a different MTO revision is ``content_equal`` (bold), not
    origin green. If both sources are close and both MTO revisions match,
    prefer content-equal RD, then the closer mtime, then RD.

    Args:
        row: One комплекты matrix row.
        rd_content_equal: ``True`` when any current RD MTO for this kit is
            semantically equal to the robot file, ``False`` when compared and
            not equal, ``None`` when there is no comparison.

    Returns:
        Paint plan and a Russian reason for the tooltip / detail pane.
    """

    if not row.robot.present:
        return RobotOrigin(matched=None)
    robot_ns = _snapshot_mto_mtime_ns(row.robot, allow_max=True)
    rd_ns = _snapshot_mto_mtime_ns(row.rd, allow_max=False)
    sq_ns = _snapshot_mto_mtime_ns(row.sq, allow_max=False)
    rd_score = _robot_date_score(robot_ns, rd_ns)
    sq_score = _robot_date_score(robot_ns, sq_ns)
    content_rd = bool(rd_content_equal) and row.rd.present
    rd_rev_ok = _robot_rev_matches(row.robot, row.rd)
    sq_rev_ok = _robot_rev_matches(row.robot, row.sq)
    rd_origin = rd_rev_ok and (bool(rd_score) or content_rd)
    sq_origin = sq_rev_ok and bool(sq_score)

    def _delta_ns(other_ns: int | None) -> int | None:
        if not robot_ns or not other_ns:
            return None
        return abs(int(robot_ns) - int(other_ns))

    pick_rd = False
    pick_sq = False
    if content_rd and rd_rev_ok:
        pick_rd = True
    elif rd_origin and sq_origin:
        if rd_score > sq_score:
            pick_rd = True
        elif sq_score > rd_score:
            pick_sq = True
        else:
            rd_delta = _delta_ns(rd_ns) or 0
            sq_delta = _delta_ns(sq_ns) or 0
            pick_rd = rd_delta <= sq_delta
            pick_sq = not pick_rd
    elif rd_origin:
        pick_rd = True
    elif sq_origin:
        pick_sq = True

    if pick_rd or pick_sq:
        parts: list[str] = []
        if content_rd:
            parts.append("содержимое MTO совпадает с РД")
        if pick_rd and rd_score:
            delta = _delta_ns(rd_ns)
            extra = f" ({_format_mtime_delta(delta)})" if delta is not None else ""
            parts.append(f"дата MTO близка к РД{extra}")
        if pick_sq and sq_score:
            delta = _delta_ns(sq_ns)
            extra = f" ({_format_mtime_delta(delta)})" if delta is not None else ""
            parts.append(f"дата MTO близка к SQ{extra}")
        if not parts:
            parts.append("робот сопоставлен с источником MTO")
        return RobotOrigin(
            matched=True,
            rd=pick_rd,
            sq=pick_sq,
            reason="Робот: " + "; ".join(parts) + ".",
            content_equal=content_rd,
        )

    close_without_origin = bool(rd_score) or bool(sq_score) or content_rd
    if close_without_origin:
        parts = []
        if content_rd:
            parts.append("содержимое MTO совпадает с РД")
        if rd_score:
            delta = _delta_ns(rd_ns)
            extra = f" ({_format_mtime_delta(delta)})" if delta is not None else ""
            parts.append(f"дата MTO близка к РД{extra}")
        if sq_score:
            delta = _delta_ns(sq_ns)
            extra = f" ({_format_mtime_delta(delta)})" if delta is not None else ""
            parts.append(f"дата MTO близка к SQ{extra}")
        robot_rev = (row.robot.revision_text or "").strip() or "—"
        if row.rd.present and not rd_rev_ok:
            _rev, _app, mto_rev = _snapshot_mto_rev_pair(row.rd)
            parts.append(
                f"ревизия {robot_rev} не совпадает с «MTO · рев.» "
                f"{mto_rev or '—'}"
            )
        elif row.sq.present and not sq_rev_ok:
            _rev, _app, sq_rev = _snapshot_mto_rev_pair(row.sq)
            parts.append(
                f"ревизия {robot_rev} не совпадает с SQ {sq_rev or '—'}"
            )
        return RobotOrigin(
            matched=None,
            reason="Робот: " + "; ".join(parts) + ".",
            content_equal=content_rd,
        )

    if robot_ns is None:
        reason = "У робота нет даты файла MTO."
    elif rd_ns is None and sq_ns is None:
        reason = "Нет даты MTO в РД и SQ, чтобы сравнить с роботом."
    else:
        reason = "Дата MTO робота не близка ни к РД, ни к SQ (окно 3 сут)"
        if rd_content_equal is False:
            reason += "; содержимое с РД не совпадает."
        else:
            reason += "."
    return RobotOrigin(matched=False, reason=reason, content_equal=False)


def _classify_stage(text: str) -> tuple[str, str]:
    for key, label, pattern in _STAGE_PATTERNS:
        if pattern.search(text):
            return key, label
    return "other", (text.strip() or "—")


def parse_history_line(line: str) -> KitEvent:
    """Parse one column-F line into date / stage / rev / transmittal.

    Args:
        line: Raw history line.

    Returns:
        Best-effort event; ``parsed`` is False when nothing was recognized.
    """

    raw = (line or "").strip()
    date: str | None = None
    rest = raw
    dated = _DATE_RE.match(raw)
    if dated:
        date = dated.group(1)
        rest = dated.group(2).strip()
    rest_norm = normalize_unicode_dashes(rest).replace("АН", "AN").replace("ан", "AN")
    stage, stage_label = _classify_stage(rest_norm)
    if stage == "other" and not rest_norm:
        stage_label = "—"
    rev_match = _REV_IN_TEXT_RE.search(rest_norm)
    revision, appendix = (None, None)
    if rev_match:
        revision, appendix = parse_sheet_revision(rev_match.group(1))
    transmittals = tuple(dict.fromkeys(_TRM_RE.findall(rest_norm)))
    parsed = bool(
        date or transmittals or revision or (stage != "other") or rest_norm
    )
    if stage == "other" and not rest_norm and not date and not transmittals:
        parsed = False
    return KitEvent(
        raw=raw,
        date=date,
        stage=stage,
        stage_label=stage_label,
        revision=revision,
        appendix=appendix,
        transmittals=transmittals,
        parsed=parsed,
    )


def parse_history_comment(comment: str) -> tuple[KitEvent, ...]:
    """Split column F into events, one per non-empty line.

    Args:
        comment: Multi-line sheet comment.

    Returns:
        Ordered events; malformed lines stay as ``raw`` with ``parsed=False``.
    """

    events: list[KitEvent] = []
    for line in (comment or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if not line.strip():
            continue
        events.append(parse_history_line(line))
    return tuple(events)


def _last_event(events: tuple[KitEvent, ...]) -> KitEvent | None:
    dated = [event for event in events if event.date]
    if dated:
        return dated[-1]
    return events[-1] if events else None


def parse_google_kit_row(
    row: list[str],
    *,
    row_index: int,
) -> GoogleKit | str:
    """Parse one sheet row or return a skip reason.

    Args:
        row: Cells A..F (extra cells ignored).
        row_index: 1-based sheet row number.

    Returns:
        A :class:`GoogleKit` or a human-readable skip reason.
    """

    padded = list(row) + [""] * max(0, 6 - len(row))
    title = str(padded[0] or "").strip()
    mark_raw = str(padded[1] or "").strip()
    if not title and not mark_raw:
        return "empty"
    if not _TITLE_RE.fullmatch(title):
        return f"row {row_index}: title {title!r}"
    mark = normalize_mark_token(mark_raw)
    if mark.casefold() in _MARK_DENYLIST or not is_rd_kit_mark(mark):
        return f"row {row_index}: mark {mark_raw!r}"
    revision, appendix = parse_sheet_revision(str(padded[3] or ""))
    events = parse_history_comment(str(padded[5] or ""))
    return GoogleKit(
        title=title,
        mark=mark,
        mark_raw=mark_raw,
        title_system=f"{title}-{mark}",
        sheet_revision=revision,
        sheet_appendix=appendix,
        sheet_revision_text=format_revision(revision, appendix),
        status_sheet=str(padded[4] or "").strip(),
        comment_raw=str(padded[5] or ""),
        events=events,
        last_event=_last_event(events),
        row_index=row_index,
    )


def parse_google_matrix(rows: list[list[str]]) -> tuple[tuple[GoogleKit, ...], GoogleParseStats]:
    """Parse an exported CSV matrix into filtered kits.

    Args:
        rows: Full sheet including header.

    Returns:
        Kept kits (last row wins on duplicate title+mark) and skip stats.
    """

    kits_by_key: dict[tuple[str, str], GoogleKit] = {}
    skipped_reasons: list[str] = []
    start = 0
    if rows:
        header = [str(cell or "").strip().casefold() for cell in rows[0]]
        if header and header[0] in {"титул", "title"}:
            start = 1
    for offset, row in enumerate(rows[start:], start=start + 1):
        parsed = parse_google_kit_row(row, row_index=offset)
        if isinstance(parsed, str):
            if parsed != "empty":
                skipped_reasons.append(parsed)
            continue
        kits_by_key[kit_identity_key(parsed.title, parsed.mark)] = parsed
    stats = GoogleParseStats(
        kept=len(kits_by_key),
        skipped=len(skipped_reasons),
        skipped_reasons=tuple(skipped_reasons[:40]),
    )
    ordered = tuple(
        sorted(kits_by_key.values(), key=lambda kit: (kit.title, kit.mark))
    )
    return ordered, stats


def parse_issuance_kit_row(
    row: list[str],
    *,
    row_index: int,
) -> IssuanceKit | str:
    """Parse one «Выдача РД ПД» row or return a skip reason.

    Args:
        row: Sheet cells (A=№, B=TRM, C=mark, D=rev, E=title, …).
        row_index: 1-based sheet row number.

    Returns:
        An :class:`IssuanceKit` or a skip reason.
    """

    padded = list(row) + [""] * max(0, 17 - len(row))
    mark_raw = str(padded[2] or "").strip()
    revision_raw = str(padded[3] or "").strip()
    title = str(padded[4] or "").strip()
    status = str(padded[5] or "").strip()
    send_date = str(padded[6] or "").strip()
    control_date = str(padded[15] or "").strip()
    note = str(padded[16] or "").strip()
    send_trm = str(padded[1] or "").strip()
    if not title and not mark_raw:
        return "empty"
    if "3d" in mark_raw.casefold() or "3d" in revision_raw.casefold():
        return f"row {row_index}: 3D {mark_raw!r}"
    if not _TITLE_RE.fullmatch(title):
        return f"row {row_index}: title {title!r}"
    mark = normalize_mark_token(mark_raw)
    if not is_rd_kit_mark(mark):
        return f"row {row_index}: mark {mark_raw!r}"
    revision, appendix = parse_sheet_revision(revision_raw)
    return IssuanceKit(
        title=title,
        mark=mark,
        mark_raw=mark_raw,
        title_system=f"{title}-{mark}",
        revision=revision,
        appendix=appendix,
        revision_text=format_revision(revision, appendix),
        status=status,
        send_date=send_date,
        send_date_sortable=format_event_date_sortable(send_date),
        send_transmittal=send_trm,
        incoming_control_date=control_date,
        incoming_control_date_sortable=format_event_date_sortable(control_date),
        confirm_transmittal=extract_confirm_transmittal(note),
        note_raw=note,
        row_index=row_index,
    )


def _issuance_data_start(rows: list[list[str]]) -> int:
    if not rows:
        return 0
    header0 = str(rows[0][0] or "").strip().casefold() if rows[0] else ""
    if header0.startswith("№") or "п.п" in header0 or header0 in {"no", "n"}:
        return 1
    return 0


def _parse_issuance_rows(
    rows: list[list[str]],
) -> tuple[tuple[IssuanceKit, ...], list[str]]:
    sends: list[IssuanceKit] = []
    skipped_reasons: list[str] = []
    start = _issuance_data_start(rows)
    for offset, row in enumerate(rows[start:], start=start + 1):
        parsed = parse_issuance_kit_row(row, row_index=offset)
        if isinstance(parsed, str):
            if parsed != "empty":
                skipped_reasons.append(parsed)
            continue
        sends.append(parsed)
    return tuple(sends), skipped_reasons


def parse_issuance_sends(rows: list[list[str]]) -> tuple[IssuanceKit, ...]:
    """Parse every valid «Выдача РД ПД» row (not latest-wins).

    Skip rules match :func:`parse_issuance_kit_row`. Sheet order is preserved.

    Args:
        rows: Full sheet including header.

    Returns:
        All parsed sends, including repeated title+mark rows.
    """

    sends, _skipped = _parse_issuance_rows(rows)
    return sends


def latest_issuance_from_sends(
    sends: Iterable[IssuanceKit],
) -> tuple[IssuanceKit, ...]:
    """Keep the latest send per title+mark (date, then sheet row).

    Args:
        sends: Every parsed «Выдача РД ПД» row.

    Returns:
        One kit per identity, sorted by title and mark.
    """

    best: dict[tuple[str, str], IssuanceKit] = {}
    for parsed in sends:
        key = kit_identity_key(parsed.title, parsed.mark)
        previous = best.get(key)
        if previous is None:
            best[key] = parsed
            continue
        prev_rank = (previous.send_date_sortable, previous.row_index)
        new_rank = (parsed.send_date_sortable, parsed.row_index)
        if new_rank >= prev_rank:
            best[key] = parsed
    return tuple(sorted(best.values(), key=lambda kit: (kit.title, kit.mark)))


def parse_issuance_matrix(
    rows: list[list[str]],
) -> tuple[tuple[IssuanceKit, ...], GoogleParseStats]:
    """Parse «Выдача РД ПД» CSV; keep the latest send per title+mark.

    Latest means highest sortable send date, then highest sheet row index.

    Args:
        rows: Full sheet including header.

    Returns:
        One kit per title+mark and skip stats.
    """

    sends, skipped_reasons = _parse_issuance_rows(rows)
    ordered = latest_issuance_from_sends(sends)
    stats = GoogleParseStats(
        kept=len(ordered),
        skipped=len(skipped_reasons),
        skipped_reasons=tuple(skipped_reasons[:40]),
    )
    return ordered, stats


def _record_mark(record: FileRecord) -> str | None:
    mark = str(record.data.get("mark") or "").strip()
    if not is_rd_kit_mark(mark):
        return None
    title = str(record.data.get("title") or "").strip()
    if not _TITLE_RE.fullmatch(title):
        return None
    if str(record.data.get("parse_status") or "") != "parsed":
        return None
    return mark


def _file_revision(record: FileRecord) -> tuple[str | None, str | None]:
    """Return revision/AN parsed from the filename, ignoring folder names."""

    revision = record.data.get("revision")
    appendix = record.data.get("appendix")
    return (
        str(revision) if revision not in (None, "") else None,
        str(appendix) if appendix not in (None, "") else None,
    )


def _file_kind_rank(record: FileRecord) -> int:
    kind = str(record.data.get("file_kind") or "").casefold()
    if kind == FileKind.PDF.value:
        return 2
    if kind == FileKind.MTO_XLSX.value:
        return 1
    return 0


def _discipline_is_od(record: FileRecord) -> bool:
    block = str(record.data.get("discipline_block") or "").casefold()
    return block.startswith("od")


def _discipline_is_mto(record: FileRecord) -> bool:
    if str(record.data.get("file_kind") or "").casefold() == FileKind.MTO_XLSX.value:
        return True
    return str(record.data.get("discipline_block") or "").casefold().startswith(
        "mto"
    )


def _is_kit_anchor_file(record: FileRecord) -> bool:
    """Return whether a file counts for kits transfer review (MTO / OD)."""

    return _discipline_is_od(record) or _discipline_is_mto(record)


def _record_file_kind(record: FileRecord) -> str:
    return str(record.data.get("file_kind") or "").casefold()


def _is_pdf_record(record: FileRecord) -> bool:
    return _record_file_kind(record) == FileKind.PDF.value


def _is_editable_mto(record: FileRecord) -> bool:
    return _record_file_kind(record) == FileKind.MTO_XLSX.value


def _path_suffix(path: str) -> str:
    return Path(str(path or "").replace("\\", "/")).suffix.casefold()


def _is_mto_workbook_path(path: str) -> bool:
    return _path_suffix(path) in {".xlsx", ".xls"}


def _is_review_mto_workbook(record: FileRecord) -> bool:
    """Return whether this MTO is an xlsx/xls workbook, not a PDF copy."""

    if _is_pdf_record(record) or not _discipline_is_mto(record):
        return False
    if _is_editable_mto(record):
        return True
    return _is_mto_workbook_path(str(record.path or ""))


def _is_editable_review_file(record: FileRecord) -> bool:
    """Return whether an MTO/OD file is editable (not a PDF copy)."""

    return _is_kit_anchor_file(record) and not _is_pdf_record(record)


_REVIEW_TABLE_HEADER = (
    "NN",
    "роль",
    "док.",
    "рев.",
    "дата",
    "к отпр.",
    "сверка",
    "путь",
)
_REVIEW_TABLE_WIDTHS = (4, 10, 12, 12, 18, 14, 16, 0)
_BEST_SEND_MIN_PERCENT = 35
_AGREED_AFTER_GRACE_DAYS = 7
_FOLDER_DATE_YMD_RE = re.compile(
    r"от[_ \-](\d{4})[.\-](\d{1,2})[.\-](\d{1,2})",
    re.IGNORECASE,
)
_FOLDER_DATE_DMY_RE = re.compile(
    r"от[_ \-](\d{1,2})[.\-](\d{1,2})[.\-](\d{2,4})",
    re.IGNORECASE,
)


def _anchor_doc_label(record: FileRecord) -> str:
    block = str(record.data.get("discipline_block") or "").strip()
    if block:
        return block
    if _discipline_is_mto(record):
        return "MTO"
    if _discipline_is_od(record):
        return "OD"
    return "док."


def _mtime_clock_dot(mtime_ns: int) -> str:
    stamp = _datetime_from_mtime_ns(mtime_ns)
    if stamp is None:
        return "—"
    return stamp.strftime("%Y.%m.%d %H:%M")


def working_folders_from_pipelines(
    pipelines: Iterable[Any],
) -> dict[tuple[str, str], frozenset[str]]:
    """Return casefolded issued-folder names treated as working, by kit.

    Annulled folders are listed separately via
    :func:`annulled_folders_from_pipelines`.

    Args:
        pipelines: ``KitPipelineRow`` objects (or duck-typed rows with
            ``title``, ``mark``, ``working_transfer_names``).

    Returns:
        Kit identity → folder keys. Empty names are omitted.
    """

    result: dict[tuple[str, str], frozenset[str]] = {}
    for row in pipelines:
        names = frozenset(
            str(name).strip().casefold()
            for name in getattr(row, "working_transfer_names", ())
            if str(name).strip()
        )
        if names:
            result[kit_identity_key(row.title, row.mark)] = names
    return result


def annulled_folders_from_pipelines(
    pipelines: Iterable[Any],
) -> dict[tuple[str, str], frozenset[str]]:
    """Return casefolded issued-folder names treated as annulled, by kit.

    Args:
        pipelines: ``KitPipelineRow`` objects (or duck-typed rows with
            ``title``, ``mark``, ``annulled_transfer_names``).

    Returns:
        Kit identity → folder keys. Empty names are omitted.
    """

    result: dict[tuple[str, str], frozenset[str]] = {}
    for row in pipelines:
        names = frozenset(
            str(name).strip().casefold()
            for name in getattr(row, "annulled_transfer_names", ())
            if str(name).strip()
        )
        if names:
            result[kit_identity_key(row.title, row.mark)] = names
    return result


def _transfer_folder_key(record: FileRecord) -> str:
    name = str(record.data.get("transfer_name") or "").strip()
    if name:
        return name.casefold()
    package = issued_package_dir(record.path)
    if package:
        return Path(package).name.casefold()
    return ""


def _working_folder_set(
    working_folders_by_kit: Mapping[tuple[str, str], Collection[str]] | None,
    key: tuple[str, str],
) -> frozenset[str]:
    if not working_folders_by_kit:
        return frozenset()
    return frozenset(
        str(name).strip().casefold()
        for name in working_folders_by_kit.get(key, ())
        if str(name).strip()
    )


def _record_is_working_folder(
    record: FileRecord, working_folders: Collection[str]
) -> bool:
    if not working_folders:
        return False
    key = _transfer_folder_key(record)
    return bool(key) and key in working_folders


def _working_folders_for_record(
    record: FileRecord,
    working_folders_by_kit: Mapping[tuple[str, str], Collection[str]] | None,
) -> frozenset[str]:
    mark = _record_mark(record)
    title = str(record.data.get("title") or "").strip()
    if mark is None or not title:
        return frozenset()
    return _working_folder_set(
        working_folders_by_kit, kit_identity_key(title, mark)
    )


def _skip_folders_for_record(
    record: FileRecord,
    working_folders_by_kit: Mapping[tuple[str, str], Collection[str]] | None,
    annulled_folders_by_kit: Mapping[tuple[str, str], Collection[str]] | None,
) -> frozenset[str]:
    return _working_folders_for_record(
        record, working_folders_by_kit
    ) | _working_folders_for_record(record, annulled_folders_by_kit)


def _transfer_review_role(
    record: FileRecord,
    *,
    current_seq: int,
    other_role: str,
    working_folders: Collection[str],
    annulled_folders: Collection[str] = (),
) -> str:
    if _record_is_working_folder(record, annulled_folders):
        return "аннулирована"
    if _record_is_working_folder(record, working_folders):
        return "рабочая"
    seq = record.data.get("transfer_sequence")
    if seq is not None and int(seq) == int(current_seq):
        return "текущая"
    return other_role


def _review_rel_path(record: FileRecord) -> str:
    path = str(record.path or "")
    relative = path_relative_to_transfer_gate(path)
    if relative and ("\\" in relative or "/" in relative):
        return relative
    folder = str(record.data.get("transfer_name") or "").strip()
    name = str(record.data.get("name") or "").strip()
    if not name and path:
        name = PureWindowsPath(path.replace("/", "\\")).name
    leaf = relative or name or "—"
    if not folder:
        return leaf
    parent = ""
    if path:
        parent = PureWindowsPath(path.replace("/", "\\")).parent.name
    if parent and is_package_media_folder(parent):
        return f"{folder}\\{parent}\\{leaf}"
    return f"{folder}\\{leaf}"


def _review_table_line(values: tuple[str, ...]) -> str:
    cells: list[str] = []
    for value, width in zip(values, _REVIEW_TABLE_WIDTHS, strict=True):
        text = value or "—"
        if width:
            text = f"{text:<{width}}"
        cells.append(text)
    return "\t".join(cells)


def _concern_phrase(rev_inverted: bool, mtime_inverted: bool) -> str:
    bits: list[str] = []
    if rev_inverted:
        bits.append("ревизия ниже")
    if mtime_inverted:
        bits.append("файл старше")
    return " и ".join(bits)


@dataclass(frozen=True, slots=True)
class _SendTarget:
    revision_text: str = ""
    date_text: str = ""
    source: str = ""


@dataclass(frozen=True, slots=True)
class _ReviewTableRow:
    seq: int
    role: str
    label: str
    rev_text: str
    date_text: str
    percent: int | None
    rel_path: str
    is_mto: bool
    abs_path: str = ""
    mto_id: str = ""
    mtime_ns: int = 0


def _kit_send_target(
    google: GoogleKit | None,
    issuance: IssuanceKit | None,
) -> _SendTarget:
    """Pick the send revision/date for transfer-review proximity.

    Выдача wins; last Google F fills a missing revision or date.
    """

    rev = ""
    date_text = ""
    source = ""
    if issuance is not None:
        rev = (issuance.revision_text or "").strip()
        date_text = (
            (issuance.send_date or "").strip()
            or (issuance.incoming_control_date or "").strip()
        )
        if rev:
            source = "Выдача"
    event = google.last_event if google is not None else None
    f_rev = format_revision(event.revision, event.appendix) if event else ""
    f_date = (event.date or "").strip() if event else ""
    if not rev and f_rev:
        rev = f_rev
        date_text = f_date or date_text
        source = "F"
    elif rev and not date_text and f_date:
        date_text = f_date
    if not rev and google is not None:
        sheet = (google.sheet_revision_text or "").strip()
        if sheet:
            rev = sheet
            source = "Google"
    return _SendTarget(revision_text=rev, date_text=date_text, source=source)


_CODE_LETTER_STAGES = frozenset({"code_a", "code_b", "code_c"})
_CODE_LETTER_BY_STAGE = {"code_a": "A", "code_b": "B", "code_c": "C"}
MTO_CATALOG_DATE_ACTION = "Заменить дату на дату последнего кода A/B/C"
MTO_CATALOG_DATE_MANUAL_ACTION = "Заменить дату вручную…"
MTO_CATALOG_DATE_FOLDER_ACTION = "Заменить дату на среднюю по папке"


def _revision_texts_eq(left: str, right: str) -> bool:
    if not left or not right:
        return False
    return revisions_equivalent(
        *parse_sheet_revision(left),
        *parse_sheet_revision(right),
    )


def _last_code_letter_event(
    google: GoogleKit | None,
    official_rev: str,
    *,
    stages: Collection[str] = _CODE_LETTER_STAGES,
) -> KitEvent | None:
    """Return the last F code-letter event for ``official_rev``.

    Args:
        google: KSB ИД kit with parsed F events.
        official_rev: Filename revision to match, or empty to allow
            unscoped letters.
        stages: Event stages to keep. Default is A/B/C.

    Returns:
        Last matching event in journal order, or ``None``.
    """

    if google is None:
        return None
    wanted = frozenset(stages)
    matched: list[KitEvent] = []
    unscoped: list[KitEvent] = []
    for event in google.events:
        if event.stage not in wanted:
            continue
        text = format_revision(event.revision, event.appendix)
        if text:
            if official_rev and _revision_texts_eq(text, official_rev):
                matched.append(event)
        else:
            unscoped.append(event)
    pool = matched or unscoped
    return pool[-1] if pool else None


def _last_code_a_event(
    google: GoogleKit | None,
    official_rev: str,
) -> KitEvent | None:
    """Return the last code-A event for the official revision."""

    return _last_code_letter_event(google, official_rev, stages=("code_a",))


def last_code_letter_for_revision(
    google: GoogleKit | None,
    revision_text: str,
) -> tuple[str, str]:
    """Return the last F code A/B/C date and stage for ``revision_text``.

    The last letter in journal order wins: a later B or C replaces A.

    Args:
        google: KSB ИД kit with parsed F events.
        revision_text: Filename revision of the MTO (``02``, ``01-AN02``).

    Returns:
        ``(DD.MM.YYYY, code_a|code_b|code_c)``. Both empty when there is
        no matching letter.
    """

    event = _last_code_letter_event(google, revision_text)
    if event is None:
        return "", ""
    date_text = str(event.date or "").strip()
    if not date_text:
        return "", ""
    stage = str(event.stage or "").strip()
    if stage not in _CODE_LETTER_BY_STAGE:
        return "", ""
    return date_text, stage


def code_letter_label(stage: str) -> str:
    """Return ``код A`` / ``код B`` / ``код C``, or empty."""

    letter = _CODE_LETTER_BY_STAGE.get(str(stage or "").strip(), "")
    return f"код {letter}" if letter else ""


def code_a_date_for_revision(
    google: GoogleKit | None,
    revision_text: str,
) -> str:
    """Return the last F code-A date for ``revision_text``.

    Args:
        google: KSB ИД kit with parsed F events.
        revision_text: Filename revision of the MTO (``02``, ``01-AN02``).

    Returns:
        ``DD.MM.YYYY``, or empty when there is no matching letter A.
    """

    event = _last_code_a_event(google, revision_text)
    if event is None:
        return ""
    return str(event.date or "").strip()


def _kit_agreed_target(
    google: GoogleKit | None,
    issuance: IssuanceKit | None,
) -> _SendTarget:
    """Pick the agreed-cycle revision/date for official-package scoring.

    Filename revision comes from Выдача / F (same as send). Date prefers
    the last code-A letter on that revision, else the send date.
    """

    send = _kit_send_target(google, issuance)
    code_event = _last_code_a_event(google, send.revision_text)
    code_date = (code_event.date or "").strip() if code_event is not None else ""
    if code_date:
        rev = send.revision_text
        if not rev and code_event is not None:
            rev = format_revision(code_event.revision, code_event.appendix)
        return _SendTarget(
            revision_text=rev,
            date_text=code_date,
            source="код A",
        )
    return send


def _parse_send_date(text: str) -> date | None:
    raw = (text or "").strip()
    if not raw:
        return None
    dmy = re.fullmatch(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", raw)
    if dmy:
        try:
            return date(int(dmy.group(3)), int(dmy.group(2)), int(dmy.group(1)))
        except ValueError:
            return None
    ymd = re.fullmatch(r"(\d{4})[.\-](\d{1,2})[.\-](\d{1,2})", raw)
    if ymd:
        try:
            return date(int(ymd.group(1)), int(ymd.group(2)), int(ymd.group(3)))
        except ValueError:
            return None
    return None


def _file_transfer_date(record: FileRecord) -> date | None:
    """Prefer folder ``от_YYYY.MM.DD``, else file mtime (same idea as АН)."""

    haystack = " ".join(
        (
            str(record.data.get("transfer_name") or ""),
            str(record.path or ""),
        )
    )
    text = normalize_unicode_dashes(haystack)
    ymd = _FOLDER_DATE_YMD_RE.search(text)
    if ymd:
        try:
            return date(int(ymd.group(1)), int(ymd.group(2)), int(ymd.group(3)))
        except ValueError:
            pass
    dmy = _FOLDER_DATE_DMY_RE.search(text)
    if dmy:
        year = int(dmy.group(3))
        if year < 100:
            year += 2000
        try:
            return date(year, int(dmy.group(2)), int(dmy.group(1)))
        except ValueError:
            pass
    stamp = _datetime_from_mtime_ns(int(record.data.get("mtime_ns") or 0))
    return stamp.date() if stamp is not None else None


def _file_mtime_date(record: FileRecord) -> date | None:
    stamp = _datetime_from_mtime_ns(int(record.data.get("mtime_ns") or 0))
    return stamp.date() if stamp is not None else None


def _send_date_proximity_points(
    actual: date | None, target: date | None
) -> tuple[int, str]:
    """Date bands copied from ``an_index._date_proximity_points``."""

    if target is None or actual is None:
        return 0, ""
    delta = abs((actual - target).days)
    stamp = actual.strftime("%d.%m.%Y")
    if delta <= 21:
        return 18, f"дата {stamp} близка к отправке"
    if delta <= 45:
        return 10, f"дата {stamp} в пределах полутора месяцев"
    if delta <= 90:
        return 3, f"дата {stamp} в том же квартале"
    if delta <= 150:
        return -12, f"дата {stamp} далеко от отправки"
    return -22, f"дата {stamp} не из цикла отправки"


def _agreed_anchor_label(source: str) -> tuple[str, str]:
    """Return (close-to, later-than) Russian phrases for the target source."""

    if source == "код A":
        return "письму A", "письма A"
    return "отправке", "отправки"


def _agreed_date_proximity_points(
    actual: date | None,
    target: date | None,
    *,
    source: str,
) -> tuple[int, str]:
    """Date bands like АН, with a penalty when the file is after the letter."""

    if target is None or actual is None:
        return 0, ""
    delta = (actual - target).days
    stamp = actual.strftime("%d.%m.%Y")
    _close_to, later_than = _agreed_anchor_label(source)
    if delta > _AGREED_AFTER_GRACE_DAYS:
        if delta <= 21:
            return -8, f"дата {stamp} позже {later_than}"
        if delta <= 45:
            return -16, f"дата {stamp} заметно позже {later_than}"
        return -24, f"дата {stamp} не из цикла {later_than}"
    abs_delta = abs(delta)
    if abs_delta <= 21:
        return 18, f"дата {stamp} близка к {_close_to}"
    if abs_delta <= 45:
        return 10, f"дата {stamp} в пределах полутора месяцев"
    if abs_delta <= 90:
        return 3, f"дата {stamp} в том же квартале"
    if abs_delta <= 150:
        return -12, f"дата {stamp} далеко от {later_than}"
    return -22, f"дата {stamp} не из цикла {later_than}"


def _score_rd_file_for_send(
    record: FileRecord,
    target: _SendTarget,
    *,
    package_mto_rev: str,
    mto_expected: bool,
    late_after_target: bool = False,
) -> int | None:
    """Return 0–100 closeness to the kit send, or ``None`` without a target.

    Weights follow АН «К согл. передаче»: exact filename rev, then date
    proximity. MTO matching the send gets an extra bonus; a package that
    should have MTO and does not is penalized.

    Args:
        record: RD OD or MTO file in a disputed package.
        target: Выдача / F revision and date for this kit.
        package_mto_rev: Filename revision of MTO in the same NN, if any.
        mto_expected: True when at least one disputed package has MTO.
        late_after_target: Penalize file mtime after the letter/send date.

    Returns:
        Clamped percent, or ``None`` when the kit has no send revision.
    """

    if not target.revision_text:
        return None
    points = 0
    file_rev, file_app = _file_revision(record)
    target_rev, target_app = parse_sheet_revision(target.revision_text)
    exact = bool(
        file_rev
        and target_rev
        and revisions_equivalent(file_rev, file_app, target_rev, target_app)
    )
    if exact:
        points += 40
        if _discipline_is_mto(record):
            points += 12
    elif (
        file_rev
        and target_rev
        and file_rev.isdigit()
        and target_rev.isdigit()
        and int(file_rev) == int(target_rev)
    ):
        points += 16
    if not _discipline_is_mto(record) and mto_expected:
        if package_mto_rev:
            mto_rev, mto_app = parse_sheet_revision(package_mto_rev)
            if (
                mto_rev
                and target_rev
                and revisions_equivalent(mto_rev, mto_app, target_rev, target_app)
            ):
                points += 8
        else:
            points -= 10
    if late_after_target:
        date_points, _reason = _agreed_date_proximity_points(
            _file_mtime_date(record),
            _parse_send_date(target.date_text),
            source=target.source,
        )
    else:
        date_points, _reason = _send_date_proximity_points(
            _file_transfer_date(record), _parse_send_date(target.date_text)
        )
    points += date_points
    return max(0, min(100, points))


def _format_send_percent(percent: int | None, *, is_best: bool) -> str:
    if percent is None:
        return "—"
    text = f"{percent}%"
    if is_best:
        return f"{text} · лучше"
    return text


def _mto_path_pair_key(left: str, right: str) -> tuple[str, str]:
    left_id = make_path_key(normalize_unicode_dashes(left))
    right_id = make_path_key(normalize_unicode_dashes(right))
    return (left_id, right_id) if left_id <= right_id else (right_id, left_id)


def _resolve_mto_compare_label(
    left: _ReviewTableRow,
    right: _ReviewTableRow,
    *,
    mto_compare: Callable[[str, str], str] | None,
    mto_compare_by_pair: Mapping[tuple[str, str], str] | None,
) -> str:
    if mto_compare is not None:
        return mto_compare(left.abs_path, right.abs_path)
    if mto_compare_by_pair:
        found = mto_compare_by_pair.get(
            _mto_path_pair_key(left.abs_path, right.abs_path)
        )
        if found:
            return found
    return "не сверялось"


def _mto_compare_column(
    rows: Sequence[_ReviewTableRow],
    *,
    mto_compare: Callable[[str, str], str] | None,
    mto_compare_by_pair: Mapping[tuple[str, str], str] | None,
) -> tuple[tuple[str, ...], tuple[TransferReviewMtoPair, ...]]:
    labels = ["—"] * len(rows)
    pairs: list[TransferReviewMtoPair] = []
    groups: dict[str, list[int]] = {}
    for index, row in enumerate(rows):
        if not row.is_mto:
            continue
        if row.role in {"рабочая", "аннулирована"}:
            continue
        if not row.abs_path:
            labels[index] = "нет пары"
            continue
        if not _is_mto_workbook_path(row.abs_path):
            continue
        groups.setdefault(row.mto_id or "mto", []).append(index)
    for indexes in groups.values():
        present = [index for index in indexes if rows[index].abs_path]
        if len(present) < 2:
            for index in present:
                labels[index] = "нет пары"
            continue
        present_rows = [rows[index] for index in present]
        if len(present) == 2:
            left, right = present_rows
            grade = _resolve_mto_compare_label(
                left,
                right,
                mto_compare=mto_compare,
                mto_compare_by_pair=mto_compare_by_pair,
            )
            for index in present:
                labels[index] = grade
            pairs.append(
                TransferReviewMtoPair(
                    left_path=left.abs_path,
                    right_path=right.abs_path,
                    left_mtime_ns=left.mtime_ns,
                    right_mtime_ns=right.mtime_ns,
                )
            )
            continue
        current = [index for index in present if rows[index].role == "текущая"]
        ref_index = current[0] if current else present[0]
        ref = rows[ref_index]
        labels[ref_index] = "эталон"
        for index in present:
            if index == ref_index:
                continue
            other = rows[index]
            labels[index] = _resolve_mto_compare_label(
                ref,
                other,
                mto_compare=mto_compare,
                mto_compare_by_pair=mto_compare_by_pair,
            )
            pairs.append(
                TransferReviewMtoPair(
                    left_path=ref.abs_path,
                    right_path=other.abs_path,
                    left_mtime_ns=ref.mtime_ns,
                    right_mtime_ns=other.mtime_ns,
                )
            )
    return tuple(labels), tuple(pairs)


def _package_mto_revision(files: Sequence[FileRecord]) -> str:
    mtos = [item for item in files if _is_review_mto_workbook(item)]
    if not mtos:
        return ""
    best = max(mtos, key=_kit_revision_sort_key)
    revision, appendix = _file_revision(best)
    return format_revision(revision, appendix)


def _package_highest_revision_text(files: Sequence[FileRecord]) -> str:
    anchors = [item for item in files if _is_kit_anchor_file(item)]
    if not anchors:
        return ""
    best = max(anchors, key=_kit_revision_sort_key)
    revision, appendix = _file_revision(best)
    return format_revision(revision, appendix)


def _pick_best_send_row(rows: Sequence[_ReviewTableRow]) -> int | None:
    scored = [
        (index, row)
        for index, row in enumerate(rows)
        if row.percent is not None and row.role not in {"рабочая", "аннулирована"}
    ]
    if not scored:
        return None
    best = max(row.percent or 0 for _index, row in scored)
    if best < _BEST_SEND_MIN_PERCENT:
        return None
    candidates = [index for index, row in scored if row.percent == best]
    mto = [index for index in candidates if rows[index].is_mto]
    pool = mto or candidates
    disputed = [index for index in pool if rows[index].role == "спорная"]
    return (disputed or pool)[0]


def _kit_revision_sort_key(
    record: FileRecord,
) -> tuple[tuple[int, int, str], int, int, int]:
    """Rank MTO/OD files for the kit revision cell (filename, not folder NN)."""

    revision, appendix = _file_revision(record)
    return (
        revision_rank(revision, appendix),
        int(_discipline_is_od(record)),
        _file_kind_rank(record),
        int(record.data.get("mtime_ns") or 0),
    )


def _file_sort_key(
    record: FileRecord,
) -> tuple[int, tuple[int, int, str], int, int, int]:
    """Rank overlay-current files for the open-folder path (newest-first).

    Transfer sequence is the only folder signal. Folder ``рев.*`` text is
    ignored. Kit revision is chosen separately from MTO/OD via
    :func:`_kit_revision_sort_key`.
    """

    sequence = record.data.get("transfer_sequence")
    seq_value = int(sequence) if sequence not in (None, "") else -1
    revision, appendix = _file_revision(record)
    mtime = int(record.data.get("mtime_ns") or 0)
    return (
        seq_value,
        revision_rank(revision, appendix),
        int(_discipline_is_od(record)),
        _file_kind_rank(record),
        mtime,
    )


def _as_bool(value: object) -> bool:
    return bool(int(value)) if isinstance(value, (int, str)) and str(value).isdigit() else bool(value)


def aggregate_source_kits(
    records: Iterable[FileRecord],
    *,
    source: SourceKind,
    detected_current_ids: set[int] | None = None,
) -> dict[tuple[str, str], SourceKitSnapshot]:
    """Group present parsed files of one source by title+mark.

    Args:
        records: Persisted catalog files.
        source: RD, SQ, or ROBOT.
        detected_current_ids: Required for RD (overlay current files only).

    Returns:
        Snapshots keyed by :func:`kit_identity_key`. ``revision_text`` is
        overlay-current **OD** (else MTO). ``mto_revision_text`` is
        overlay-current **MTO** only. ``paths`` stay newest-first.
    """

    grouped: dict[tuple[str, str], list[FileRecord]] = {}
    for record in records:
        if record.source is not source or not record.present:
            continue
        if source is SourceKind.RD:
            if detected_current_ids is None or record.id not in detected_current_ids:
                continue
        mark = _record_mark(record)
        if mark is None:
            continue
        title = str(record.data.get("title") or "").strip()
        grouped.setdefault(kit_identity_key(title, mark), []).append(record)

    snapshots: dict[tuple[str, str], SourceKitSnapshot] = {}
    for key, files in grouped.items():
        ordered = sorted(files, key=_file_sort_key, reverse=True)
        newest = ordered[0]
        od_files = [item for item in files if _discipline_is_od(item)]
        mto_files = [item for item in files if _discipline_is_mto(item)]
        if od_files:
            rev_source = max(od_files, key=_kit_revision_sort_key)
        elif mto_files:
            rev_source = max(mto_files, key=_kit_revision_sort_key)
        else:
            rev_source = newest
        revision, appendix = _file_revision(rev_source)
        mto_revision = mto_appendix = None
        mto_revision_text = ""
        if mto_files:
            mto_source = max(mto_files, key=_kit_revision_sort_key)
            mto_revision, mto_appendix = _file_revision(mto_source)
            mto_revision_text = format_revision(mto_revision, mto_appendix)
        mtimes = [int(item.data.get("mtime_ns") or 0) for item in files]
        mto_mtimes = [
            int(item.data.get("mtime_ns") or 0)
            for item in files
            if str(item.data.get("file_kind") or "").casefold()
            == FileKind.MTO_XLSX.value
        ]
        snapshots[key] = SourceKitSnapshot(
            present=True,
            revision=revision,
            appendix=appendix,
            revision_text=format_revision(revision, appendix),
            mto_revision=mto_revision,
            mto_appendix=mto_appendix,
            mto_revision_text=mto_revision_text,
            max_mtime_ns=max(mtimes) if mtimes else None,
            mto_mtime_ns=max(mto_mtimes) if any(mto_mtimes) else None,
            file_count=len(files),
            as_build=any(
                _as_bool(item.data.get("transfer_is_as_build"))
                or path_is_as_build(item.path)
                for item in files
            ),
            transfer_name=(
                str(newest.data.get("transfer_name"))
                if newest.data.get("transfer_name")
                else None
            ),
            file_ids=tuple(item.id for item in ordered),
            paths=tuple(item.path for item in ordered),
        )
    return snapshots


def _choose_display_name(
    key: tuple[str, str],
    google: GoogleKit | None,
    issuance: IssuanceKit | None,
    records_by_key: dict[tuple[str, str], tuple[str, str]],
) -> tuple[str, str]:
    if google is not None:
        return google.title, google.mark
    if issuance is not None:
        return issuance.title, issuance.mark
    return records_by_key.get(key, (key[0], key[1]))


def _rd_document_identity(record: FileRecord) -> tuple[str, ...] | None:
    if record.source is not SourceKind.RD or not record.present:
        return None
    if str(record.data.get("parse_status") or "") != ParseStatus.PARSED.value:
        return None
    kind = str(record.data.get("file_kind") or "")
    if kind == FileKind.PDF.value:
        stem = str(record.data.get("core_stem") or "").strip().casefold()
        if not stem:
            return None
        return (kind, stem)
    if kind == FileKind.MTO_XLSX.value:
        title_system = str(record.data.get("title_system") or "").strip().casefold()
        disc = str(record.data.get("discipline_block") or "").strip().casefold()
        if not title_system or not disc:
            return None
        return (kind, title_system, disc)
    return None


def _transfer_review_notes(
    records: Iterable[FileRecord],
    current_ids: set[int],
    *,
    google_by_key: Mapping[tuple[str, str], GoogleKit] | None = None,
    issuance_by_key: Mapping[tuple[str, str], IssuanceKit] | None = None,
    mto_compare: Callable[[str, str], str] | None = None,
    mto_compare_by_pair: Mapping[str, str] | None = None,
    working_folders_by_kit: Mapping[tuple[str, str], Collection[str]] | None = None,
    annulled_folders_by_kit: Mapping[tuple[str, str], Collection[str]] | None = None,
) -> tuple[
    dict[tuple[str, str], tuple[str, ...]],
    dict[tuple[str, str], tuple[TransferReviewMtoPair, ...]],
]:
    """Return per-kit human notes when later vs earlier transfers disagree.

    A later sequence stays current. Notes explain a higher filename revision
    or a newer mtime still sitting in an older package of the same **MTO or
    OD**, list companion MTO/OD from those packages, and score closeness
    to the kit send (Выдача / F) like АН «К согл. передаче». Working
    and annulled folders are not «спорная» counterparts and do not raise
    this flag.
    """

    google_map = google_by_key or {}
    issuance_map = issuance_by_key or {}
    grouped: dict[tuple[str, ...], list[FileRecord]] = {}
    package_files: dict[tuple[tuple[str, str], int], list[FileRecord]] = {}
    for record in records:
        if record.source is not SourceKind.RD or not record.present:
            continue
        if not _is_kit_anchor_file(record):
            continue
        identity = _rd_document_identity(record)
        if identity is None:
            continue
        grouped.setdefault(identity, []).append(record)
        mark = _record_mark(record)
        title = str(record.data.get("title") or "").strip()
        seq = record.data.get("transfer_sequence")
        if mark is None or not title or seq in (None, ""):
            continue
        package_files.setdefault(
            (kit_identity_key(title, mark), int(seq)), []
        ).append(record)

    hits_by_key: dict[tuple[str, str], list[tuple[str, int, int, str]]] = {}
    involved_seqs: dict[tuple[str, str], set[int]] = {}
    for copies in grouped.values():
        current = [
            item
            for item in copies
            if item.id in current_ids
            and not _record_is_working_folder(
                item,
                _skip_folders_for_record(
                    item, working_folders_by_kit, annulled_folders_by_kit
                ),
            )
        ]
        if not current:
            continue
        for newer in current:
            new_seq = newer.data.get("transfer_sequence")
            if new_seq in (None, ""):
                continue
            new_seq_int = int(new_seq)
            new_rev, new_app = _file_revision(newer)
            new_rank = revision_rank(new_rev, new_app)
            new_mtime = int(newer.data.get("mtime_ns") or 0)
            for older in copies:
                if older.id == newer.id:
                    continue
                if _record_is_working_folder(
                    older,
                    _skip_folders_for_record(
                        older, working_folders_by_kit, annulled_folders_by_kit
                    ),
                ):
                    continue
                old_seq = older.data.get("transfer_sequence")
                if old_seq in (None, ""):
                    continue
                old_seq_int = int(old_seq)
                if old_seq_int >= new_seq_int:
                    continue
                old_rev, old_app = _file_revision(older)
                old_rank = revision_rank(old_rev, old_app)
                old_mtime = int(older.data.get("mtime_ns") or 0)
                rev_inverted = (
                    new_rank[0] >= 0 and old_rank[0] >= 0 and new_rank < old_rank
                )
                mtime_inverted = new_mtime < old_mtime
                if not rev_inverted and not mtime_inverted:
                    continue
                mark = _record_mark(newer)
                title = str(newer.data.get("title") or "").strip()
                if mark is None or not title:
                    continue
                label = _anchor_doc_label(newer)
                concern = _concern_phrase(rev_inverted, mtime_inverted)
                key = kit_identity_key(title, mark)
                hits_by_key.setdefault(key, []).append(
                    (label, new_seq_int, old_seq_int, concern)
                )
                involved_seqs.setdefault(key, set()).update(
                    (new_seq_int, old_seq_int)
                )

    notes_by_key: dict[tuple[str, str], tuple[str, ...]] = {}
    pairs_by_key: dict[tuple[str, str], tuple[TransferReviewMtoPair, ...]] = {}
    for key, hits in hits_by_key.items():
        ordered_hits = list(dict.fromkeys(hits))
        ordered_hits.sort(
            key=lambda item: (item[0].casefold(), -int(item[1]), -int(item[2]))
        )
        sequences = sorted(involved_seqs.get(key, ()), reverse=True)
        if not sequences:
            continue
        current_seq = sequences[0]
        working = _working_folder_set(working_folders_by_kit, key)
        annulled = _working_folder_set(annulled_folders_by_kit, key)
        skipped = working | annulled
        target = _kit_send_target(google_map.get(key), issuance_map.get(key))
        table_rows, mto_expected = _review_rows_for_sequences(
            sequences=sequences,
            current_seq=current_seq,
            package_files=package_files,
            key=key,
            target=target,
            other_role="спорная",
            late_after_target=False,
            working_folders=working,
            annulled_folders=annulled,
        )
        if not table_rows:
            continue

        best_index = _pick_best_send_row(table_rows)
        compare_col, pairs = _mto_compare_column(
            table_rows,
            mto_compare=mto_compare,
            mto_compare_by_pair=mto_compare_by_pair,
        )
        lines = ["Что смущает:"]
        seen_concern: set[str] = set()
        for label, new_seq, old_seq, concern in ordered_hits:
            line = (
                f"  {label} — NN {new_seq:02d} текущая, но {concern}, "
                f"чем в NN {old_seq:02d}"
            )
            if line in seen_concern:
                continue
            seen_concern.add(line)
            lines.append(line)
        if mto_expected:
            later_mtos = [
                item
                for item in package_files.get((key, current_seq), ())
                if _is_review_mto_workbook(item)
                and not _record_is_working_folder(item, skipped)
            ]
            if not later_mtos:
                found_earlier = False
                for seq in sequences[1:]:
                    earlier_rev = _package_mto_revision(
                        [
                            item
                            for item in package_files.get((key, seq), [])
                            if not _record_is_working_folder(item, skipped)
                        ]
                    )
                    if earlier_rev:
                        lines.append(
                            f"  MTO — нет в NN {current_seq:02d}, есть в "
                            f"NN {seq:02d} (рев. {earlier_rev})"
                        )
                        found_earlier = True
                        break
                if not found_earlier:
                    lines.append(f"  MTO — нет в NN {current_seq:02d}")
        if best_index is not None:
            best = table_rows[best_index]
            current_scores = [
                row.percent
                for row in table_rows
                if row.role == "текущая" and row.percent is not None
            ]
            current_best = max(current_scores) if current_scores else None
            if best.role == "спорная" and (
                current_best is None or (best.percent or 0) > current_best
            ):
                extra = (
                    f" ({current_best}%)" if current_best is not None else ""
                )
                lines.append(
                    f"  к отправке ближе спорная NN {best.seq:02d} "
                    f"({best.percent}%), чем текущая NN {current_seq:02d}"
                    f"{extra}"
                )
        lines.append("")
        if target.revision_text:
            target_bits = [target.source, target.revision_text, target.date_text]
            lines.append(
                "Близость к отправке: "
                + " · ".join(bit for bit in target_bits if bit)
            )
        lines.append("Пути от папки передачи:")
        lines.append(_review_table_line(_REVIEW_TABLE_HEADER))
        for index, row in enumerate(table_rows):
            lines.append(
                _review_table_line(
                    (
                        f"{row.seq:02d}",
                        row.role,
                        row.label,
                        row.rev_text,
                        row.date_text,
                        _format_send_percent(
                            row.percent, is_best=index == best_index
                        ),
                        compare_col[index] if index < len(compare_col) else "—",
                        row.rel_path,
                    )
                )
            )
        lines.append(
            "Сверка MTO: только xlsx/xls, без PDF. "
            "Тот же алгоритм, что «Сверка Авто МТО» "
            "(четкое / ПоКоду и Кол-ву / не совпало). "
            "Нет второй MTO в спорных NN — нет пары."
        )
        notes_by_key[key] = tuple(lines)
        if pairs:
            pairs_by_key[key] = pairs
    return notes_by_key, pairs_by_key


def _late_official_anchor_line(
    files: Sequence[FileRecord],
    target: _SendTarget,
    current_seq: int,
) -> str:
    """Return a concern line when official MTO/OD mtime is after the letter."""

    mtos = [item for item in files if _is_review_mto_workbook(item)]
    if not mtos:
        mtos = [item for item in files if _discipline_is_mto(item)]
    anchors = mtos or [item for item in files if _discipline_is_od(item)]
    if not anchors:
        return ""
    record = max(anchors, key=_kit_revision_sort_key)
    actual = _file_mtime_date(record)
    target_date = _parse_send_date(target.date_text)
    if actual is None or target_date is None:
        return ""
    if (actual - target_date).days <= _AGREED_AFTER_GRACE_DAYS:
        return ""
    kind = "MTO" if mtos else "OD"
    later_than = _agreed_anchor_label(target.source)[1]
    return (
        f"  {kind} текущей NN {current_seq:02d} "
        f"({actual.strftime('%d.%m.%Y')}) позже {later_than} "
        f"({target_date.strftime('%d.%m.%Y')})"
    )


def _review_rows_for_sequences(
    *,
    sequences: Sequence[int],
    current_seq: int,
    package_files: Mapping[tuple[tuple[str, str], int], list[FileRecord]],
    key: tuple[str, str],
    target: _SendTarget,
    other_role: str,
    late_after_target: bool,
    working_folders: Collection[str] = (),
    annulled_folders: Collection[str] = (),
) -> tuple[list[_ReviewTableRow], bool]:
    skipped = frozenset(
        str(name).strip().casefold()
        for name in (*working_folders, *annulled_folders)
        if str(name).strip()
    )
    mto_expected = any(
        _is_review_mto_workbook(item)
        for seq in sequences
        for item in package_files.get((key, seq), ())
        if not _record_is_working_folder(item, skipped)
    )
    has_editable = any(
        _is_editable_review_file(item)
        for seq in sequences
        for item in package_files.get((key, seq), ())
        if not _record_is_working_folder(item, skipped)
    )
    table_rows: list[_ReviewTableRow] = []
    for seq in sequences:
        files = package_files.get((key, seq), [])
        mto_files = [
            item for item in files if _is_review_mto_workbook(item)
        ]
        od_files = [
            item
            for item in files
            if _discipline_is_od(item) and not _is_pdf_record(item)
        ]
        if not has_editable and not mto_files and not od_files:
            od_files = [item for item in files if _discipline_is_od(item)]
        seq_official = [
            item for item in files if not _record_is_working_folder(item, skipped)
        ]
        seq_annulled = bool(files) and all(
            _record_is_working_folder(item, annulled_folders) for item in files
        )
        seq_working = bool(files) and not seq_official and not seq_annulled
        placeholder_role = (
            "аннулирована"
            if seq_annulled
            else "рабочая"
            if seq_working
            else ("текущая" if seq == current_seq else other_role)
        )
        package_mto_rev = _package_mto_revision(files)
        if mto_expected and not mto_files:
            table_rows.append(
                _ReviewTableRow(
                    seq=seq,
                    role=placeholder_role,
                    label="MTO",
                    rev_text="—",
                    date_text="—",
                    percent=0 if target.revision_text else None,
                    rel_path="нет в пакете",
                    is_mto=True,
                )
            )
        for record in list(mto_files) + list(od_files):
            rev, appendix = _file_revision(record)
            mtime = int(record.data.get("mtime_ns") or 0)
            is_mto = _is_review_mto_workbook(record)
            table_rows.append(
                _ReviewTableRow(
                    seq=seq,
                    role=_transfer_review_role(
                        record,
                        current_seq=current_seq,
                        other_role=other_role,
                        working_folders=working_folders,
                        annulled_folders=annulled_folders,
                    ),
                    label=_anchor_doc_label(record),
                    rev_text=format_revision(rev, appendix) or "—",
                    date_text=_mtime_clock_dot(mtime),
                    percent=_score_rd_file_for_send(
                        record,
                        target,
                        package_mto_rev=package_mto_rev,
                        mto_expected=mto_expected,
                        late_after_target=late_after_target,
                    ),
                    rel_path=_review_rel_path(record),
                    is_mto=is_mto,
                    abs_path=str(record.path or ""),
                    mto_id=(
                        str(record.data.get("discipline_block") or "")
                        .strip()
                        .casefold()
                        if is_mto
                        else ""
                    ),
                    mtime_ns=mtime,
                )
            )
    return table_rows, mto_expected


def _agreed_cycle_package_notes(
    records: Iterable[FileRecord],
    *,
    google_by_key: Mapping[tuple[str, str], GoogleKit] | None = None,
    issuance_by_key: Mapping[tuple[str, str], IssuanceKit] | None = None,
    mto_compare: Callable[[str, str], str] | None = None,
    mto_compare_by_pair: Mapping[str, str] | None = None,
    working_folders_by_kit: Mapping[tuple[str, str], Collection[str]] | None = None,
    annulled_folders_by_kit: Mapping[tuple[str, str], Collection[str]] | None = None,
) -> tuple[
    dict[tuple[str, str], tuple[str, ...]],
    dict[tuple[str, str], tuple[TransferReviewMtoPair, ...]],
]:
    """Flag official packages that lose closeness to letter A / send.

    Official current is the highest **non-working and non-annulled** NN of
    the agreed filename revision. Working folders stay as ``рабочая`` and
    annulled folders as ``аннулирована``; neither wins ``текущая``,
    triggers closeness, or joins MTO сверка.
    """

    google_map = google_by_key or {}
    issuance_map = issuance_by_key or {}
    package_files: dict[tuple[tuple[str, str], int], list[FileRecord]] = {}
    for record in records:
        if record.source is not SourceKind.RD or not record.present:
            continue
        if not _is_kit_anchor_file(record):
            continue
        mark = _record_mark(record)
        title = str(record.data.get("title") or "").strip()
        seq = record.data.get("transfer_sequence")
        if mark is None or not title or seq in (None, ""):
            continue
        package_files.setdefault(
            (kit_identity_key(title, mark), int(seq)), []
        ).append(record)

    kit_seqs: dict[tuple[str, str], set[int]] = {}
    for (key, seq), _files in package_files.items():
        kit_seqs.setdefault(key, set()).add(seq)

    notes_by_key: dict[tuple[str, str], tuple[str, ...]] = {}
    pairs_by_key: dict[tuple[str, str], tuple[TransferReviewMtoPair, ...]] = {}
    keys = set(google_map) | set(issuance_map) | set(kit_seqs)
    for key in keys:
        target = _kit_agreed_target(google_map.get(key), issuance_map.get(key))
        if not target.revision_text or not target.date_text:
            continue
        working = _working_folder_set(working_folders_by_kit, key)
        annulled = _working_folder_set(annulled_folders_by_kit, key)
        skipped = working | annulled
        official_seqs: list[int] = []
        display_seqs: list[int] = []
        for seq in sorted(kit_seqs.get(key, ()), reverse=True):
            files = package_files.get((key, seq), [])
            official_files = [
                item
                for item in files
                if not _record_is_working_folder(item, skipped)
            ]
            highest_official = _package_highest_revision_text(official_files)
            highest_any = _package_highest_revision_text(files)
            matches_official = bool(
                highest_official
                and _revision_texts_eq(highest_official, target.revision_text)
            )
            matches_any = bool(
                highest_any and _revision_texts_eq(highest_any, target.revision_text)
            )
            if matches_official:
                official_seqs.append(seq)
            if matches_any:
                display_seqs.append(seq)
        if not official_seqs:
            continue
        current_seq = official_seqs[0]
        table_rows, _mto_expected = _review_rows_for_sequences(
            sequences=display_seqs or official_seqs,
            current_seq=current_seq,
            package_files=package_files,
            key=key,
            target=target,
            other_role="другая",
            late_after_target=True,
            working_folders=working,
            annulled_folders=annulled,
        )
        if not table_rows:
            continue
        best_index = _pick_best_send_row(table_rows)
        current_scores = [
            row.percent
            for row in table_rows
            if row.role == "текущая" and row.percent is not None
        ]
        current_best = max(current_scores) if current_scores else None
        closer = False
        if best_index is not None:
            best = table_rows[best_index]
            closer = best.role not in {"текущая", "рабочая", "аннулирована"} and (
                current_best is None or (best.percent or 0) > current_best
            )
        official_current_files = [
            item
            for item in package_files.get((key, current_seq), [])
            if not _record_is_working_folder(item, skipped)
        ]
        late_line = _late_official_anchor_line(
            official_current_files,
            target,
            current_seq,
        )
        if not closer and not late_line:
            continue
        close_to, _later_than = _agreed_anchor_label(target.source)
        compare_col, pairs = _mto_compare_column(
            table_rows,
            mto_compare=mto_compare,
            mto_compare_by_pair=mto_compare_by_pair,
        )
        lines = ["Согласованная передача:"]
        if late_line:
            lines.append(late_line)
        if closer and best_index is not None:
            best = table_rows[best_index]
            extra = f" ({current_best}%)" if current_best is not None else ""
            lines.append(
                f"  к {close_to} ближе NN {best.seq:02d} "
                f"({best.percent}%), чем текущая NN {current_seq:02d}"
                f"{extra}"
            )
        lines.append("")
        target_bits = [target.source, target.revision_text, target.date_text]
        lines.append(
            f"Близость к {close_to}: "
            + " · ".join(bit for bit in target_bits if bit)
        )
        lines.append("Пути от папки передачи:")
        lines.append(_review_table_line(_REVIEW_TABLE_HEADER))
        for index, row in enumerate(table_rows):
            lines.append(
                _review_table_line(
                    (
                        f"{row.seq:02d}",
                        row.role,
                        row.label,
                        row.rev_text,
                        row.date_text,
                        _format_send_percent(
                            row.percent, is_best=index == best_index
                        ),
                        compare_col[index] if index < len(compare_col) else "—",
                        row.rel_path,
                    )
                )
            )
        lines.append(
            "Сверка MTO: только xlsx/xls, без PDF. "
            "Тот же алгоритм, что «Сверка Авто МТО». "
            "Папка открытия остаётся официальной NN. "
            "Роль «рабочая» — помеченная / выше выдачи папка; "
            "«аннулирована» — снята с конкурса. Обе не в сверке и выборе NN."
        )
        notes_by_key[key] = tuple(lines)
        if pairs:
            pairs_by_key[key] = pairs
    return notes_by_key, pairs_by_key


def _merge_transfer_review_payloads(
    inversion_notes: Mapping[tuple[str, str], tuple[str, ...]],
    inversion_pairs: Mapping[tuple[str, str], tuple[TransferReviewMtoPair, ...]],
    agreed_notes: Mapping[tuple[str, str], tuple[str, ...]],
    agreed_pairs: Mapping[tuple[str, str], tuple[TransferReviewMtoPair, ...]],
) -> tuple[
    dict[tuple[str, str], tuple[str, ...]],
    dict[tuple[str, str], tuple[TransferReviewMtoPair, ...]],
]:
    notes: dict[tuple[str, str], tuple[str, ...]] = dict(inversion_notes)
    for key, lines in agreed_notes.items():
        previous = notes.get(key, ())
        notes[key] = previous + (("",) if previous else ()) + tuple(lines)
    pairs: dict[tuple[str, str], list[TransferReviewMtoPair]] = {
        key: list(value) for key, value in inversion_pairs.items()
    }
    for key, extra in agreed_pairs.items():
        existing = pairs.setdefault(key, [])
        seen = {
            _mto_path_pair_key(item.left_path, item.right_path)
            for item in existing
        }
        for pair in extra:
            ident = _mto_path_pair_key(pair.left_path, pair.right_path)
            if ident in seen:
                continue
            seen.add(ident)
            existing.append(pair)
    return notes, {key: tuple(value) for key, value in pairs.items()}


def _mixed_title_notes(
    records: Sequence[FileRecord],
) -> dict[tuple[str, str], tuple[str, ...]]:
    """Return tooltip lines when RD files sit under another title folder.

    Groups present parsed RD files whose issued-path title folder is not
    the filename title. Does not drop those files from the catalog.

    Args:
        records: Contour (or raw) catalog files.

    Returns:
        Kit identity → note lines, empty when every file matches its
        title folder.
    """

    grouped: dict[tuple[str, str], list[tuple[str, str, str]]] = {}
    for record in records:
        if record.source is not SourceKind.RD or not record.present:
            continue
        mark = _record_mark(record)
        if mark is None:
            continue
        parsed_title = str(record.data.get("title") or "").strip()
        folder_title = issued_path_title_folder(record.path)
        if not folder_title:
            continue
        if folder_title.casefold() == parsed_title.casefold():
            continue
        package = str(record.data.get("transfer_name") or "").strip()
        if not package:
            package = nearest_issued_package_hint(record.path)
        name = str(record.data.get("name") or "").strip()
        if not name:
            name = PureWindowsPath(record.path).name
        key = kit_identity_key(parsed_title, mark)
        grouped.setdefault(key, []).append((folder_title, package, name))

    notes: dict[tuple[str, str], tuple[str, ...]] = {}
    header = (
        "Смешанные титулы: файлы комплекта лежат в папке другого титула."
    )
    for key, items in grouped.items():
        lines = [header]
        seen: set[tuple[str, str, str]] = set()
        shown = 0
        for folder_title, package, name in items:
            stamp = (
                folder_title.casefold(),
                package.casefold(),
                name.casefold(),
            )
            if stamp in seen:
                continue
            seen.add(stamp)
            extra = f" · {package}" if package else ""
            lines.append(f"  папка {folder_title}{extra} · {name}")
            shown += 1
            if shown >= _MIXED_TITLE_NOTE_LIMIT:
                remaining = len(items) - shown
                if remaining > 0:
                    lines.append(f"  … ещё {remaining} файл(ов)")
                break
        notes[key] = tuple(lines)
    return notes


def _mixed_title_rd_records(
    records: Sequence[FileRecord],
    *,
    title: str,
    mark: str,
) -> list[FileRecord]:
    """Return present RD files of this kit that sit under another title."""

    want = kit_identity_key(title, mark)
    if not want[0] or not want[1]:
        return []
    found: list[FileRecord] = []
    for record in records:
        if record.source is not SourceKind.RD or not record.present:
            continue
        record_mark = _record_mark(record)
        if record_mark is None:
            continue
        parsed_title = str(record.data.get("title") or "").strip()
        if kit_identity_key(parsed_title, record_mark) != want:
            continue
        folder_title = issued_path_title_folder(record.path)
        if not folder_title:
            continue
        if folder_title.casefold() == parsed_title.casefold():
            continue
        found.append(record)
    return found


def mixed_title_open_folders(
    records: Sequence[FileRecord],
    *,
    title: str,
    mark: str,
    rd_root: str | Path,
) -> tuple[str, ...]:
    """Return directories that contain this kit's mixed-title files.

    One folder per unique parent of a stray file (``PDF`` / ``DWG`` /
    ``NN_…``). Explorer should open those, not the foreign mark folder.

    Args:
        records: Catalog file rows.
        title: Filename title of the kit.
        mark: Latin AGCC mark of that kit.
        rd_root: Configured RD source root.

    Returns:
        Unique containing folders in first-seen order.
    """

    root_text = str(rd_root or "").strip()
    if not root_text:
        return ()
    found: dict[str, str] = {}
    for record in _mixed_title_rd_records(
        records, title=title, mark=mark
    ):
        folder = containing_folder(record.path)
        if not folder:
            folder = issued_package_dir(record.path)
        if not folder:
            continue
        if not path_is_under(folder, root_text):
            continue
        if path_is_under(root_text, folder):
            continue
        key = make_path_key(folder)
        if key not in found:
            found[key] = folder
    return tuple(found.values())


def mixed_title_rescan_folders(
    records: Sequence[FileRecord],
    *,
    title: str,
    mark: str,
    rd_root: str | Path,
) -> tuple[str, ...]:
    """Return foreign title/mark folders where this kit's RD files sit.

    A point RD rescan walks these folders so missing detection can clear
    files that were moved or deleted under another title. Does not drop
    mixed files from the catalog and never returns ``rd_root``.

    Prefers the mark folder under the foreign title
    (``rd_root / 2612 / 05_KSB``). Falls back to the foreign title folder
    when the mark folder cannot be recovered from the path.

    Args:
        records: Catalog file rows (typically present contour records).
        title: Filename title of the kit being rescanned.
        mark: Latin AGCC mark of that kit.
        rd_root: Configured RD source root.

    Returns:
        Unique folders in first-seen order, empty when nothing is mixed.
    """

    root_text = str(rd_root or "").strip()
    if not root_text:
        return ()
    found: dict[str, str] = {}
    for record in _mixed_title_rd_records(records, title=title, mark=mark):
        folder_title = issued_path_title_folder(record.path)
        mark_folder = kit_rd_mark_folder_from_path(
            record.path, root_text, title=folder_title
        )
        title_folder = constructed_kit_rd_title_folder(root_text, folder_title)
        candidate = mark_folder or title_folder
        if not _usable_mixed_title_rescan_folder(candidate, root_text):
            if candidate != title_folder and _usable_mixed_title_rescan_folder(
                title_folder, root_text
            ):
                candidate = title_folder
            else:
                continue
        key = make_path_key(candidate)
        if key not in found:
            found[key] = candidate
    return tuple(found.values())


def _usable_mixed_title_rescan_folder(folder: str, rd_root: str) -> bool:
    """Return whether ``folder`` is a safe extra walk root under RD."""

    text = str(folder or "").strip()
    if not text:
        return False
    if not path_is_under(text, rd_root):
        return False
    if path_is_under(rd_root, text):
        return False
    if is_transfer_gate_folder_name(Path(text).name):
        return False
    return True


def _row_flags(
    google: GoogleKit | None,
    issuance: IssuanceKit | None,
    rd: SourceKitSnapshot,
    robot: SourceKitSnapshot,
    sq: SourceKitSnapshot,
    *,
    transfer_review: bool = False,
    mixed_titles: bool = False,
) -> tuple[KitFlag, ...]:
    has_google = google is not None or issuance is not None
    flags: list[KitFlag] = []
    if has_google and not rd.present and not robot.present and not sq.present:
        flags.append(KitFlag.GOOGLE_ONLY)
    if has_google and not rd.present:
        flags.append(KitFlag.GAP_RD)
    if has_google and not robot.present:
        flags.append(KitFlag.GAP_ROBOT)
    if rd.present and not has_google:
        flags.append(KitFlag.EXTRA_RD)
    if robot.present and not has_google:
        flags.append(KitFlag.EXTRA_ROBOT)
    if sq.present and not has_google:
        flags.append(KitFlag.EXTRA_SQ)
    compare_rev = None
    compare_app = None
    if google is not None and google.sheet_revision:
        compare_rev, compare_app = google.sheet_revision, google.sheet_appendix
    elif issuance is not None and issuance.revision:
        compare_rev, compare_app = issuance.revision, issuance.appendix
    if (
        compare_rev
        and rd.present
        and rd.revision
        and not revisions_equivalent(
            compare_rev,
            compare_app,
            rd.revision,
            rd.appendix,
        )
    ):
        flags.append(KitFlag.REV_MISMATCH)
    if transfer_review:
        flags.append(KitFlag.TRANSFER_REVIEW)
    if mixed_titles:
        flags.append(KitFlag.MIXED_TITLES)
    if has_google and rd.present and KitFlag.REV_MISMATCH not in flags:
        flags.append(KitFlag.ALIGNED)
    return tuple(flags)


def _primary_summary(flags: tuple[KitFlag, ...]) -> KitSummary:
    flag_values = {flag.value for flag in flags}
    for candidate in _SUMMARY_PRIORITY:
        if candidate.value in flag_values:
            return candidate
    return KitSummary.ALIGNED


def build_kit_matrix(
    google_kits: Iterable[GoogleKit],
    records: Iterable[FileRecord],
    detected_current_ids: set[int],
    issuance_kits: Iterable[IssuanceKit] = (),
    rd_root: str | Path | None = None,
    *,
    mto_compare: Callable[[str, str], str] | None = None,
    mto_compare_by_pair: Mapping[str, str] | None = None,
    working_folders_by_kit: Mapping[tuple[str, str], Collection[str]] | None = None,
    annulled_folders_by_kit: Mapping[tuple[str, str], Collection[str]] | None = None,
) -> tuple[KitMatrixRow, ...]:
    """Build the union presence matrix for комплекты.

    Args:
        google_kits: Filtered KSB ИД Google rows.
        records: Persisted scan files.
        detected_current_ids: Overlay current RD file ids.
        issuance_kits: Latest «Выдача РД ПД» rows per title+mark.
        rd_root: RD source root. When set, non-canonical RD files are
            excluded via :func:`record_has_canonical_layout`. SQ and ROBOT
            rows are kept.
        mto_compare: Optional test hook ``(left_path, right_path) → label``.
        mto_compare_by_pair: Sorted path-id pair → Russian compare label.
        working_folders_by_kit: Issued folder names treated as working.
            They are not ``текущая`` / ``спорная`` in «Проверить передачи».
        annulled_folders_by_kit: Issued folder names treated as annulled.
            Same skip as working; role is ``аннулирована``. ``None`` means
            no extra skip.

    Returns:
        Sorted matrix rows.
    """

    with perf_span("kits.build_kit_matrix"):
        google_by_key = {
            kit_identity_key(kit.title, kit.mark): kit for kit in google_kits
        }
        issuance_by_key = {
            kit_identity_key(kit.title, kit.mark): kit for kit in issuance_kits
        }
        materialized = list(records)
        root = str(rd_root or "").strip()
        if root:
            materialized = [
                record
                for record in materialized
                if record_has_canonical_layout(record, root)
            ]
        rd_map = aggregate_source_kits(
            materialized, source=SourceKind.RD, detected_current_ids=detected_current_ids
        )
        robot_map = aggregate_source_kits(materialized, source=SourceKind.ROBOT)
        sq_map = aggregate_source_kits(materialized, source=SourceKind.SQ)

        display_names: dict[tuple[str, str], tuple[str, str]] = {}
        for record in materialized:
            mark = _record_mark(record)
            if mark is None:
                continue
            title = str(record.data.get("title") or "").strip()
            display_names.setdefault(kit_identity_key(title, mark), (title, mark))
        review_notes, review_pairs = _merge_transfer_review_payloads(
            *_transfer_review_notes(
                materialized,
                detected_current_ids,
                google_by_key=google_by_key,
                issuance_by_key=issuance_by_key,
                mto_compare=mto_compare,
                mto_compare_by_pair=mto_compare_by_pair,
                working_folders_by_kit=working_folders_by_kit,
                annulled_folders_by_kit=annulled_folders_by_kit,
            ),
            *_agreed_cycle_package_notes(
                materialized,
                google_by_key=google_by_key,
                issuance_by_key=issuance_by_key,
                mto_compare=mto_compare,
                mto_compare_by_pair=mto_compare_by_pair,
                working_folders_by_kit=working_folders_by_kit,
                annulled_folders_by_kit=annulled_folders_by_kit,
            ),
        )

        mixed_notes = _mixed_title_notes(materialized)
        keys = (
            set(google_by_key)
            | set(issuance_by_key)
            | set(rd_map)
            | set(robot_map)
            | set(sq_map)
        )
        rows: list[KitMatrixRow] = []
        empty = SourceKitSnapshot()
        for key in sorted(keys):
            google = google_by_key.get(key)
            issuance = issuance_by_key.get(key)
            rd = rd_map.get(key, empty)
            robot = robot_map.get(key, empty)
            sq = sq_map.get(key, empty)
            title, mark = _choose_display_name(key, google, issuance, display_names)
            flags = _row_flags(
                google,
                issuance,
                rd,
                robot,
                sq,
                transfer_review=key in review_notes,
                mixed_titles=key in mixed_notes,
            )
            rows.append(
                KitMatrixRow(
                    title=title,
                    mark=mark,
                    title_system=f"{title}-{mark}",
                    rd=rd,
                    robot=robot,
                    sq=sq,
                    google=google,
                    issuance=issuance,
                    flags=flags,
                    summary=_primary_summary(flags),
                    transfer_review_notes=review_notes.get(key, ()),
                    transfer_review_mto_pairs=review_pairs.get(key, ()),
                    mixed_title_notes=mixed_notes.get(key, ()),
                )
            )
        return tuple(rows)


def format_event_date_sortable(date_text: str | None) -> str:
    """Convert ``DD.MM.YYYY`` to ``YYYY.MM.DD`` for display/sort.

    Args:
        date_text: Sheet date or ``None``.

    Returns:
        Sortable date, original text if unparseable, or empty string.
    """

    raw = (date_text or "").strip()
    if not raw:
        return ""
    match = re.fullmatch(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", raw)
    if not match:
        return raw
    day, month, year = (int(match.group(1)), int(match.group(2)), int(match.group(3)))
    return f"{year:04d}.{month:02d}.{day:02d}"


def last_event_parts(
    kit: GoogleKit | None,
) -> tuple[str, str, str, str]:
    """Split the last F event into date / stage / revision / TRM cells.

    Args:
        kit: Google kit or ``None``.

    Returns:
        ``(date_yyyy_mm_dd, stage_label, revision_text, transmittal)``.
        Missing parts are empty strings (GUI shows em-dash).
    """

    event = kit.last_event if kit else None
    if event is None:
        return "", "", "", ""
    revision = format_revision(event.revision, event.appendix)
    stage = event.stage_label if event.stage_label and event.stage_label != "—" else ""
    trm = event.transmittals[-1] if event.transmittals else ""
    return format_event_date_sortable(event.date), stage, revision, trm


def last_event_text(kit: GoogleKit | None) -> str:
    """Format the last dated F event for filters / detail text.

    Args:
        kit: Google kit or ``None``.

    Returns:
        Compact ``YYYY.MM.DD · stage · rev · TRM`` text, or em-dash.
    """

    date, stage, revision, trm = last_event_parts(kit)
    parts = [part for part in (date, stage, revision, trm) if part]
    if parts:
        return " · ".join(parts)
    event = kit.last_event if kit else None
    return event.raw if event and event.raw else "—"
