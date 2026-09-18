"""AN dump filename index (MTO xlsx + OD doc/docx) and kit matching (Qt-free)."""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from utils.file_name_converts import (
    AgccFilenamePatterns,
    parse_agcc_mto_xlsx_revision_for_chain,
)

from rd_catalog.customer_pi_auto_mto import revision_texts_match, spec_as_mto_stem
from rd_catalog.kits import format_revision, parse_sheet_revision
from rd_catalog.models import FileKind, make_path_key
from rd_catalog.overlay import revision_rank
from rd_catalog.parse import (
    _agcc_filename_parts,
    _file_kind_from_name,
    matches_agcc_filename,
    normalize_unicode_dashes,
    path_is_as_build,
)

_STEM_MISMATCH_NOTE = "ствол не совпал"
KIND_MTO = "MTO"
KIND_OD = "OD"
KIND_HEADER = "Вид"
_OD_SUFFIXES = frozenset({".doc", ".docx"})


@dataclass(frozen=True, slots=True)
class AnMtoFile:
    """One accepted AN dump file (MTO xlsx or OD Word), from the file name."""

    path: str
    path_key: str
    title: str
    mark: str
    revision_text: str
    core_stem: str
    discipline_block: str
    name: str
    parent_dir: str
    mtime_ns: int
    size: int


@dataclass(frozen=True, slots=True)
class KitAnTargets:
    """Filename-revision targets from other catalog sources for one kit.

    ``auto_mto_path`` / ``rd_mto_path`` are optional workbook paths used
    by the АН tab content compare. Matching still uses the revision
    strings only.
    """

    auto_mto: str = ""
    auto_mto_stem: str = ""
    auto_mto_path: str = ""
    rd_mto: str = ""
    rd_mto_path: str = ""
    rd_overlay: str = ""
    robot: str = ""
    sq: str = ""
    issuance: str = ""
    google_sheet: str = ""
    google_f: str = ""
    agreed: str = ""
    agreed_date: str = ""


@dataclass(frozen=True, slots=True)
class AnAgreedScore:
    """How likely one AN file is the dump of the last agreed transfer.

    ``percent`` is ``None`` when the kit has no agreed revision (no code A
    and no pipeline status ``agreed``). ``is_best`` prefers an exact
    filename-revision match closest to the letter-A date; folder-name
    bonuses must not beat a later transfer of the same rev. An OD in the
    folder that matches the agreed rev is an extra folder signal (kits
    may ship without MTO; OD carries the RD revision).
    """

    percent: int | None
    is_best: bool
    agreed_revision: str
    reasons: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class AnKitHit:
    """AN files of one kit ranked against :class:`KitAnTargets`."""

    files: tuple[AnMtoFile, ...]
    shown_revision: str
    shown_file: AnMtoFile | None
    closes_auto_mto: bool
    match_auto_mto: bool | None
    match_rd_mto: bool | None
    match_robot: bool | None
    match_issuance: bool | None
    match_google_f: bool | None
    match_sq: bool | None
    stem_matched: bool
    stem_note: str


def parse_an_mto_file(
    path: str | Path,
    *,
    size: int,
    mtime_ns: int,
) -> AnMtoFile | None:
    """Parse one AN MTO workbook from its file name and stats.

    The workbook body is not opened. Junk names return ``None``.

    Args:
        path: Local or UNC file path.
        size: File size from ``stat``.
        mtime_ns: Nanosecond modification time from ``stat``.

    Returns:
        A typed AN MTO file, or ``None`` when the name is rejected.
    """

    original_path = str(path)
    original_name = Path(original_path).name
    normalized_name = normalize_unicode_dashes(original_name)
    if original_name.startswith("~$") or normalized_name.startswith("~$"):
        return None
    if Path(normalized_name).suffix.casefold() != ".xlsx":
        return None
    if _file_kind_from_name(normalized_name) is not FileKind.MTO_XLSX:
        return None
    if not matches_agcc_filename(normalized_name):
        return None
    parts = _agcc_filename_parts(normalized_name)
    if parts is None or "-" not in parts.title_system:
        return None
    if not parts.discipline_block.casefold().startswith("mto"):
        return None

    title, mark = parts.title_system.split("-", 1)
    tail = AgccFilenamePatterns.split_revision_tail(parts.revision_tail)
    revision: str | None = None
    appendix: str | None = None
    if tail:
        revision = tail.rev_sheet.split("-", 1)[0]
        appendix = tail.an

    if parts.revision_tail:
        try:
            mto_revision, mto_appendix = parse_agcc_mto_xlsx_revision_for_chain(
                normalized_name
            )
        except ValueError:
            return None
        normalized_appendix = (
            mto_appendix.removeprefix("AN") if mto_appendix else None
        )
        if (revision, appendix) != (mto_revision, normalized_appendix):
            return None
        revision, appendix = mto_revision, normalized_appendix

    return AnMtoFile(
        path=original_path,
        path_key=make_path_key(original_path),
        title=title,
        mark=mark,
        revision_text=format_revision(revision, appendix),
        core_stem=parts.core_stem,
        discipline_block=parts.discipline_block,
        name=original_name,
        parent_dir=str(Path(original_path).parent),
        mtime_ns=int(mtime_ns),
        size=int(size),
    )


def parse_an_od_file(
    path: str | Path,
    *,
    size: int,
    mtime_ns: int,
) -> AnMtoFile | None:
    """Parse one AN OD Word file from its name and stats.

    Accepts AGCC ``.doc`` / ``.docx`` whose discipline starts with ``od``.
    The file body is not opened. Junk names return ``None``.

    Args:
        path: Local or UNC file path.
        size: File size from ``stat``.
        mtime_ns: Nanosecond modification time from ``stat``.

    Returns:
        A typed dump file, or ``None`` when the name is rejected.
    """

    original_path = str(path)
    original_name = Path(original_path).name
    normalized_name = normalize_unicode_dashes(original_name)
    if original_name.startswith("~$") or normalized_name.startswith("~$"):
        return None
    if Path(normalized_name).suffix.casefold() not in _OD_SUFFIXES:
        return None
    if not matches_agcc_filename(normalized_name):
        return None
    parts = _agcc_filename_parts(normalized_name)
    if parts is None or "-" not in parts.title_system:
        return None
    if not parts.discipline_block.casefold().startswith("od"):
        return None

    title, mark = parts.title_system.split("-", 1)
    tail = AgccFilenamePatterns.split_revision_tail(parts.revision_tail)
    revision: str | None = None
    appendix: str | None = None
    if tail:
        revision = tail.rev_sheet.split("-", 1)[0]
        appendix = tail.an

    return AnMtoFile(
        path=original_path,
        path_key=make_path_key(original_path),
        title=title,
        mark=mark,
        revision_text=format_revision(revision, appendix),
        core_stem=parts.core_stem,
        discipline_block=parts.discipline_block,
        name=original_name,
        parent_dir=str(Path(original_path).parent),
        mtime_ns=int(mtime_ns),
        size=int(size),
    )


def parse_an_dump_file(
    path: str | Path,
    *,
    size: int,
    mtime_ns: int,
) -> AnMtoFile | None:
    """Parse one AN dump candidate: MTO ``.xlsx`` or OD ``.doc`` / ``.docx``.

    Args:
        path: Local or UNC file path.
        size: File size from ``stat``.
        mtime_ns: Nanosecond modification time from ``stat``.

    Returns:
        A typed dump file, or ``None`` when the name is rejected.
    """

    suffix = Path(normalize_unicode_dashes(Path(path).name)).suffix.casefold()
    if suffix == ".xlsx":
        return parse_an_mto_file(path, size=size, mtime_ns=mtime_ns)
    if suffix in _OD_SUFFIXES:
        return parse_an_od_file(path, size=size, mtime_ns=mtime_ns)
    return None


def an_file_kind(file: AnMtoFile) -> str:
    """Return ``OD`` or ``MTO`` for a dump row.

    Args:
        file: Parsed dump record.

    Returns:
        ``OD`` when the discipline block starts with ``od``, else ``MTO``.
    """

    if (file.discipline_block or "").casefold().startswith("od"):
        return KIND_OD
    return KIND_MTO


def an_is_od(file: AnMtoFile) -> bool:
    """Return whether this dump row is an OD Word file.

    Args:
        file: Parsed dump record.

    Returns:
        True when :func:`an_file_kind` is ``OD``.
    """

    return an_file_kind(file) == KIND_OD


def _mto_files(files: Sequence[AnMtoFile]) -> tuple[AnMtoFile, ...]:
    return tuple(file for file in files if not an_is_od(file))


def an_cell_text(hit: AnKitHit) -> str:
    """Return the compact AN cell label for a kit hit.

    Args:
        hit: Matched AN files for one kit.

    Returns:
        ``—`` when there are no files, otherwise the shown revision,
        with `` · {N}`` when more than one file is present.
    """

    if not hit.files:
        return "—"
    text = hit.shown_revision
    if len(hit.files) > 1:
        return f"{text} · {len(hit.files)}"
    return text


def match_an_to_kit(
    files: Sequence[AnMtoFile],
    targets: KitAnTargets,
) -> AnKitHit:
    """Rank AN MTO files of one kit against other-source revision targets.

    Caller passes only the files that belong to the kit. Revisions are
    compared with :func:`revision_texts_match`. The shown file and
    ``closes_auto_mto`` prefer MTO workbooks; OD-only kits still get a
    shown OD (the filename rev of the ведомость).

    Args:
        files: AN dump files for one kit (MTO and/or OD), used as-is.
        targets: Filename revisions from Auto MTO, RD, robot, issuance, F, SQ.

    Returns:
        Shown file, close-dump flag, per-source revision matches, and stem note.
    """

    ordered = tuple(files)
    mto = _mto_files(ordered)
    shown = _pick_shown_file(mto or ordered, targets)
    shown_revision = shown.revision_text if shown is not None else ""
    stem_matched, stem_note = _stem_fields(shown, targets)
    return AnKitHit(
        files=ordered,
        shown_revision=shown_revision,
        shown_file=shown,
        closes_auto_mto=_closes_auto_mto(mto, targets),
        match_auto_mto=_match_revision(shown_revision, targets.auto_mto),
        match_rd_mto=_match_revision(shown_revision, targets.rd_mto),
        match_robot=_match_revision(shown_revision, targets.robot),
        match_issuance=_match_revision(shown_revision, targets.issuance),
        match_google_f=_match_revision(shown_revision, targets.google_f),
        match_sq=_match_revision(shown_revision, targets.sq),
        stem_matched=stem_matched,
        stem_note=stem_note,
    )


def _stem_matches(file_stem: str, target_stem: str) -> bool:
    if not target_stem:
        return False
    left = file_stem.casefold()
    right = target_stem.casefold()
    if left == right:
        return True
    converted = spec_as_mto_stem(target_stem).casefold()
    return bool(converted) and left == converted


def _file_revision_rank(file: AnMtoFile) -> tuple[int, int, str]:
    revision, appendix = parse_sheet_revision(file.revision_text)
    return revision_rank(revision, appendix)


def _pick_shown_file(
    files: Sequence[AnMtoFile],
    targets: KitAnTargets,
) -> AnMtoFile | None:
    if not files:
        return None
    if targets.auto_mto:
        matching = [
            file
            for file in files
            if revision_texts_match(file.revision_text, targets.auto_mto) is True
        ]
        if matching:
            return max(
                matching,
                key=lambda file: (
                    _stem_matches(file.core_stem, targets.auto_mto_stem),
                    _file_revision_rank(file),
                ),
            )
    return max(files, key=_file_revision_rank)


def _closes_auto_mto(files: Sequence[AnMtoFile], targets: KitAnTargets) -> bool:
    if not targets.auto_mto:
        return False
    if revision_texts_match(targets.auto_mto, targets.rd_mto) is True:
        return False
    return any(
        revision_texts_match(file.revision_text, targets.auto_mto) is True
        for file in files
    )


def _stem_fields(
    shown: AnMtoFile | None,
    targets: KitAnTargets,
) -> tuple[bool, str]:
    if not targets.auto_mto_stem:
        return False, ""
    if shown is None:
        return False, ""
    if _stem_matches(shown.core_stem, targets.auto_mto_stem):
        return True, ""
    if (
        targets.auto_mto
        and revision_texts_match(shown.revision_text, targets.auto_mto) is True
    ):
        return False, _STEM_MISMATCH_NOTE
    return False, ""


def _match_revision(shown_revision: str, target: str) -> bool | None:
    return revision_texts_match(shown_revision, target)


_FOLDER_REV_RE = re.compile(
    r"(?:rev\.?|рев\.?)\s*([0-9]{1,2}(?:-AN\d{1,2})?|AN\d{0,2}|[VS])",
    re.IGNORECASE,
)
_FOLDER_REV_ALT_RE = re.compile(
    r"(\d{1,2}|[VS])_рев_AN(\d{1,2})",
    re.IGNORECASE,
)
_FOLDER_DATE_RE = re.compile(
    r"(?:_от|[\s_]от[\s_])(\d{1,2})[.\-](\d{1,2})[.\-](\d{2,4})",
    re.IGNORECASE,
)
_SQ_ANSWER_RE = re.compile(
    r"ответы\s+на\s+sq|sq[\s_-]*запрос|\bответ\b",
    re.IGNORECASE,
)
_SEND_FOLDER_RE = re.compile(r"отправк|передач", re.IGNORECASE)
_PACKAGE_HINT_RE = re.compile(r"комплект", re.IGNORECASE)
_BEST_AGREED_MIN_PERCENT = 35


def agreed_revision_target(
    *,
    code: str = "",
    code_revision_text: str = "",
    code_date: str = "",
    status: str = "",
    official_revision_text: str = "",
) -> tuple[str, str]:
    """Return the last agreed filename revision and its letter date.

    Code A wins. Status ``agreed`` without a letter still uses the code
    revision, else the official RD revision. Other letters (B/C) and
    in-progress TDO are not «согласованная».

    Args:
        code: Pipeline letter (``A`` / ``B`` / ``C``), or empty.
        code_revision_text: Revision the letter applies to.
        code_date: Letter date as ``DD.MM.YYYY``.
        status: Kit pipeline status.
        official_revision_text: Official on-disk revision.

    Returns:
        ``(revision_text, date_text)``. Both empty when the kit is not agreed.
    """

    letter = (code or "").strip().upper()
    code_rev = (code_revision_text or "").strip()
    official = (official_revision_text or "").strip()
    when = (code_date or "").strip()
    if letter == "A" and code_rev:
        return code_rev, when
    if (status or "").strip() == "agreed":
        return (code_rev or official), when
    return "", ""


def score_an_files_for_agreed(
    files: Sequence[AnMtoFile],
    targets: KitAnTargets,
) -> dict[str, AnAgreedScore]:
    """Score AN files of one kit against the last agreed transfer.

    Args:
        files: AN dump files for one title–mark (MTO xlsx and/or OD).
        targets: Catalog revisions, including ``agreed`` / ``agreed_date``.

    Returns:
        ``path_key`` → score. Empty ``agreed`` yields ``percent=None``
        and no highlight.
    """

    agreed = (targets.agreed or "").strip()
    if not agreed:
        return {
            file.path_key: AnAgreedScore(
                percent=None,
                is_best=False,
                agreed_revision="",
                reasons=("нет согласованной ревизии (нет кода A / статуса «согласован»)",),
            )
            for file in files
        }
    od_folders = {
        make_path_key(file.parent_dir)
        for file in files
        if an_is_od(file)
        and revision_texts_match(file.revision_text, agreed) is True
    }
    raw: list[tuple[AnMtoFile, int, tuple[str, ...]]] = []
    for file in files:
        points, reasons = _score_an_file_for_agreed(
            file,
            targets,
            agreed,
            folder_has_agreed_od=make_path_key(file.parent_dir) in od_folders,
        )
        raw.append((file, max(0, min(100, points)), reasons))
    best_key = _pick_best_agreed_path(raw, targets, agreed)
    return {
        file.path_key: AnAgreedScore(
            percent=percent,
            is_best=file.path_key == best_key,
            agreed_revision=agreed,
            reasons=reasons,
        )
        for file, percent, reasons in raw
    }


def _pick_best_agreed_path(
    raw: Sequence[tuple[AnMtoFile, int, tuple[str, ...]]],
    targets: KitAnTargets,
    agreed: str,
) -> str:
    if not raw:
        return ""
    exact = [
        item
        for item in raw
        if revision_texts_match(item[0].revision_text, agreed) is True
    ]
    pool: list[tuple[AnMtoFile, int, tuple[str, ...]]] = (
        list(exact) if exact else list(raw)
    )
    clean = [item for item in pool if not _file_is_as_build(item[0])]
    if clean:
        pool = clean
    not_sq = [item for item in pool if not _file_is_sq_answer(item[0])]
    if not_sq:
        pool = not_sq
    viable = [
        item for item in pool if item[1] >= _BEST_AGREED_MIN_PERCENT
    ]
    if not viable:
        return ""

    def sort_key(
        item: tuple[AnMtoFile, int, tuple[str, ...]],
    ) -> tuple[int, int, int, int, str]:
        file, percent, _reasons = item
        delta = _date_delta_days(file, targets.agreed_date)
        date_key = delta if delta is not None else 10_000
        od_rank = 0 if an_is_od(file) else 1
        return (date_key, -percent, od_rank, -file.mtime_ns, file.path_key)

    chosen = min(viable, key=sort_key)
    return chosen[0].path_key


def _score_an_file_for_agreed(
    file: AnMtoFile,
    targets: KitAnTargets,
    agreed: str,
    *,
    folder_has_agreed_od: bool = False,
) -> tuple[int, tuple[str, ...]]:
    points = 0
    reasons: list[str] = [f"цель: согласованная рев. {agreed}"]
    haystack = " ".join((file.parent_dir, file.path, file.name))
    exact = revision_texts_match(file.revision_text, agreed) is True
    if exact:
        points += 40
        reasons.append("рев. файла совпала")
    elif _same_numeric_revision(file.revision_text, agreed):
        points += 16
        reasons.append("та же цифра рев., другой AN")
    if _path_has_agreed_folder_rev(haystack, agreed):
        points += 12
        reasons.append("папка с rev. согласованной")
    if _PACKAGE_HINT_RE.search(haystack):
        points += 6
        reasons.append("папка «Комплект»")
    if _FOLDER_DATE_RE.search(haystack) and (
        "рев" in haystack.casefold() or "rev." in haystack.casefold()
    ):
        points += 3
        reasons.append("пакет с датой «от»")
    if (
        targets.issuance
        and revision_texts_match(targets.issuance, agreed) is True
        and revision_texts_match(file.revision_text, targets.issuance) is True
    ):
        points += 10
        reasons.append("совпала с Выдачей")
    if (
        targets.google_f
        and revision_texts_match(targets.google_f, agreed) is True
        and revision_texts_match(file.revision_text, targets.google_f) is True
    ):
        points += 8
        reasons.append("совпала с F")
    date_points, date_reason = _date_proximity_points(file, targets.agreed_date)
    points += date_points
    if date_reason:
        reasons.append(date_reason)
    if path_is_as_build(file.path) or path_is_as_build(file.parent_dir):
        points -= 22
        reasons.append("путь as-build — позже согласованной")
    elif _SEND_FOLDER_RE.search(haystack):
        points += 4
        reasons.append("папка отправки / передачи")
    if _SQ_ANSWER_RE.search(haystack):
        points -= 14
        reasons.append("похоже на ответ по SQ, не на передачу")
    if folder_has_agreed_od:
        points += 12
        reasons.append("в папке OD согласованной рев. (рев. РД)")
    if _is_later_working_copy(file, targets, agreed):
        points -= 16
        reasons.append("совпала с рабочей/роботом новее согласованной")
    return points, tuple(reasons)


def _same_numeric_revision(left: str, right: str) -> bool:
    left_rev, _left_app = parse_sheet_revision(left)
    right_rev, _right_app = parse_sheet_revision(right)
    if not left_rev or not right_rev:
        return False
    if not left_rev.isdigit() or not right_rev.isdigit():
        return False
    return int(left_rev) == int(right_rev)


def _path_has_agreed_folder_rev(haystack: str, agreed: str) -> bool:
    tokens = _folder_revision_tokens(haystack)
    if not tokens:
        return False
    if any(revision_texts_match(token, agreed) is True for token in tokens):
        return True
    wanted = _norm_rev_token(agreed)
    return any(_norm_rev_token(token) == wanted for token in tokens)


def _folder_revision_tokens(haystack: str) -> tuple[str, ...]:
    text = normalize_unicode_dashes(haystack or "")
    found: list[str] = []
    for match in _FOLDER_REV_RE.finditer(text):
        found.append(match.group(1))
    for match in _FOLDER_REV_ALT_RE.finditer(text):
        found.append(f"{match.group(1)}-AN{match.group(2)}")
    return tuple(found)


def _norm_rev_token(value: str) -> str:
    return normalize_unicode_dashes(value or "").strip().casefold()


def _file_is_as_build(file: AnMtoFile) -> bool:
    return path_is_as_build(file.path) or path_is_as_build(file.parent_dir)


def _file_is_sq_answer(file: AnMtoFile) -> bool:
    haystack = " ".join((file.parent_dir, file.path))
    return bool(_SQ_ANSWER_RE.search(haystack))


def _file_date(file: AnMtoFile) -> date | None:
    return _folder_date(file.parent_dir) or _mtime_date(file.mtime_ns)


def _date_delta_days(file: AnMtoFile, agreed_date: str) -> int | None:
    target = _parse_agreed_date(agreed_date)
    actual = _file_date(file)
    if target is None or actual is None:
        return None
    return abs((actual - target).days)


def _date_proximity_points(file: AnMtoFile, agreed_date: str) -> tuple[int, str]:
    target = _parse_agreed_date(agreed_date)
    if target is None:
        return 0, ""
    actual = _file_date(file)
    if actual is None:
        return 0, ""
    delta = abs((actual - target).days)
    stamp = actual.strftime("%d.%m.%Y")
    if delta <= 21:
        return 18, f"дата {stamp} близка к письму A"
    if delta <= 45:
        return 10, f"дата {stamp} в пределах полутора месяцев"
    if delta <= 90:
        return 3, f"дата {stamp} в том же квартале"
    if delta <= 150:
        return -12, f"дата {stamp} далеко от письма A"
    return -22, f"дата {stamp} не из цикла согласованной"


def _parse_agreed_date(text: str) -> date | None:
    raw = (text or "").strip()
    match = re.fullmatch(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", raw)
    if not match:
        return None
    try:
        return date(int(match.group(3)), int(match.group(2)), int(match.group(1)))
    except ValueError:
        return None


def _folder_date(path: str) -> date | None:
    match = _FOLDER_DATE_RE.search(normalize_unicode_dashes(path or ""))
    if not match:
        return None
    day, month, year = (int(match.group(1)), int(match.group(2)), int(match.group(3)))
    if year < 100:
        year += 2000
    try:
        return date(year, month, day)
    except ValueError:
        return None


def _mtime_date(mtime_ns: int) -> date | None:
    if not mtime_ns:
        return None
    try:
        return datetime.fromtimestamp(mtime_ns / 1_000_000_000).date()
    except (OSError, OverflowError, ValueError):
        return None


def _is_later_working_copy(
    file: AnMtoFile,
    targets: KitAnTargets,
    agreed: str,
) -> bool:
    robot = (targets.robot or "").strip()
    if not robot:
        return False
    if revision_texts_match(file.revision_text, robot) is not True:
        return False
    file_rank = _text_revision_rank(file.revision_text)
    agreed_rank = _text_revision_rank(agreed)
    return file_rank > agreed_rank


def _text_revision_rank(text: str) -> tuple[int, int, str]:
    revision, appendix = parse_sheet_revision(text)
    return revision_rank(revision, appendix)
