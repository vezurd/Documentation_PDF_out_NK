"""AGCC file and transfer-folder parsing for the RD catalog."""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path, PureWindowsPath

from utils.file_name_converts import (
    AgccFilenamePatterns,
    normalize_unicode_dashes,
    parse_agcc_mto_xlsx_revision_for_chain,
)

from rd_catalog.models import (
    FileKind,
    FileRecord,
    ParsedFile,
    ParseStatus,
    SourceKind,
    TransferMetadata,
    make_path_key,
)

# Issued transfer under a gate: only the leading 1–2 digits matter (``10_…``).
_TRANSFER_SEQUENCE_RE = re.compile(r"^\s*(?P<sequence>\d{1,2})(?!\d)")
_TRANSFER_REVISION_RE = re.compile(
    r"(?:рев(?:изия)?|rev(?:ision)?)\s*[._:\-]?\s*"
    r"(?P<revision>\d{1,2}|[VS])"
    r"(?:\s*-\s*AN\s*(?P<appendix>\d{1,2}))?",
    re.IGNORECASE,
)
_AS_BUILD_RE = re.compile(
    r"(?:as[\s_.-]*built?|исполнит(?:ельн(?:ая|ой|ые))?)",
    re.IGNORECASE,
)
# Path segments like ``1513`` (title) or ``11_SPP`` / ``06_4130-KSB1`` (mark)
# start with digits but are not transfer folders.
_TITLE_ONLY_FOLDER_RE = re.compile(r"^\s*\d{4}\s*$")
# Gate between mark and issued ``NN_…`` packages. Name family, not a full match.
_TRANSFER_GATE_RE = re.compile(r"передач|отправк", re.IGNORECASE)


def is_transfer_gate_folder_name(folder_name: str) -> bool:
    """Return whether a folder is a transfer gate (not an issued package).

    Gates sit between mark and ``NN_…`` packages: ``Для передачи``,
    ``На_отправку``, ``03_Для передачи``, ``2025.12.08_Для передачи``, and
    other names containing ``передач`` / ``отправк``.

    Args:
        folder_name: One folder name, not a full path.

    Returns:
        ``True`` when the name matches the gate family.
    """

    normalized = normalize_unicode_dashes(folder_name).strip()
    return bool(normalized and _TRANSFER_GATE_RE.search(normalized))


def is_transfer_folder_name(folder_name: str, *, under_gate: bool = False) -> bool:
    """Return whether a path segment is an issued ``NN_…`` transfer folder.

    Under a gate, only the leading 1–2 digits matter: ``10_рев.AN01_…`` is a
    transfer. Outside a gate, mark folders (``11_SPP``, ``12_POS1``) stay
    rejected unless the name is a bare ``NN`` / ``NN_рев.<digit>``.

    Args:
        folder_name: One folder name, not a full path.
        under_gate: Whether a gate folder is an ancestor of this segment.

    Returns:
        ``True`` for issued transfer packages.
    """

    normalized = normalize_unicode_dashes(folder_name).strip()
    if not normalized or _TITLE_ONLY_FOLDER_RE.fullmatch(normalized):
        return False
    if is_transfer_gate_folder_name(normalized) and not under_gate:
        return False
    sequence_match = _TRANSFER_SEQUENCE_RE.match(normalized)
    if sequence_match is None:
        return False
    if under_gate:
        return True
    if _TRANSFER_REVISION_RE.search(normalized):
        return True
    remainder = normalized[sequence_match.end() :].strip(" ._-")
    remainder = _AS_BUILD_RE.sub("", remainder).strip(" ._-")
    return not remainder


def parse_transfer_folder(
    folder_name: str,
    *,
    under_gate: bool = False,
) -> TransferMetadata:
    """Parse an issued transfer folder name.

    The sequence is the leading 1–2 digits. Folder ``рев.*`` text is stored
    when present but is not a contract for overlay or kit revision.

    Args:
        folder_name: One folder name, not a full path.
        under_gate: Whether a gate folder is an ancestor of this segment.

    Returns:
        Parsed transfer metadata. Invalid names are represented by
        ``ParseStatus.UNPARSED_FOLDER`` rather than raising.
    """

    normalized = normalize_unicode_dashes(folder_name)
    sequence_match = _TRANSFER_SEQUENCE_RE.match(normalized.strip())
    revision_match = _TRANSFER_REVISION_RE.search(normalized)
    title_system = AgccFilenamePatterns.scan_title_system(normalized)
    is_as_build = bool(_AS_BUILD_RE.search(normalized))

    errors: list[str] = []
    if not is_transfer_folder_name(folder_name, under_gate=under_gate):
        if sequence_match is None:
            errors.append("transfer sequence is missing")
        else:
            errors.append(
                "folder is not an issued transfer (expected NN_… under a gate)"
            )

    title: str | None = None
    mark: str | None = None
    if title_system:
        title, mark = title_system.split("-", 1)

    return TransferMetadata(
        original_name=folder_name,
        normalized_name=normalized,
        sequence=(
            int(sequence_match.group("sequence")) if sequence_match is not None else None
        ),
        revision=revision_match.group("revision") if revision_match else None,
        appendix=revision_match.group("appendix") if revision_match else None,
        title=title,
        mark=mark,
        title_system=title_system,
        is_as_build=is_as_build,
        parse_status=(
            ParseStatus.UNPARSED_FOLDER if errors else ParseStatus.PARSED
        ),
        error="; ".join(errors) or None,
    )


def find_transfer_in_path(root: str | Path, source_root: str | Path) -> TransferMetadata:
    """Find the issued transfer folder for a file directory.

    Walks path segments from the leaf toward ``source_root``. A segment is an
    issued transfer when it starts with 1–2 digits and either sits under a
    gate folder or matches the legacy ``NN`` / ``NN_рев.<digit>`` shape.

    Args:
        root: Directory that contains the candidate file (may be ``PDF``/``DWG``).
        source_root: RD or SQ source root.

    Returns:
        Parsed transfer metadata, or an unparsed placeholder when none found.
    """

    try:
        relative = Path(root).relative_to(Path(source_root))
    except ValueError:
        relative = Path(root)
    parts = [part for part in relative.parts if part not in {"\\", "/", ""}]
    for index, part in enumerate(reversed(parts)):
        ancestors = parts[: len(parts) - index - 1]
        under_gate = any(is_transfer_gate_folder_name(name) for name in ancestors)
        if is_transfer_folder_name(part, under_gate=under_gate):
            return parse_transfer_folder(part, under_gate=under_gate)
    return TransferMetadata(
        original_name=Path(root).name,
        normalized_name=Path(root).name,
        parse_status=ParseStatus.UNPARSED_FOLDER,
        error="no issued NN_… transfer folder found",
    )


_PACKAGE_MEDIA_FOLDERS = frozenset({"pdf", "dwg", "пдф", "bbb"})


def is_package_media_folder(folder_name: str) -> bool:
    """Return whether a folder is an issued-package media directory.

    Args:
        folder_name: One path segment (``PDF``, ``DWG``, ``BBB``, ``ПДФ``).

    Returns:
        ``True`` when the name is a known media folder (case-insensitive).
    """

    return bool(folder_name) and folder_name.casefold() in _PACKAGE_MEDIA_FOLDERS


# Author-facing labels are instructions (what to do), not diagnoses.
_LAYOUT_REASON_LABELS: dict[str, str] = {
    "no_gate": "Создайте папку «Для передачи» между маркой и пакетом NN",
    "loose_in_mark": "Перенесите файл из папки марки в «Для передачи» / NN_рев.…",
    "extra_subfolder": (
        "Уберите лишнюю подпапку внутри пакета NN "
        "(оставьте PDF, DWG, BBB или корень пакета)"
    ),
    "not_a_package": "Переименуйте папку пакета в NN_рев.… (номер передачи в начале)",
    "too_shallow": "Соберите цепочку ТИТУЛ / МАРКА / «Для передачи» / NN_рев.…",
    "outside_gate": "Положите файлы в «Для передачи» / NN_рев.…",
    "package_nested": (
        "Поднимите пакет NN сразу под «Для передачи» (уберите промежуточную папку)"
    ),
    "duplicate_level": (
        "Уберите повтор папки титула или марки между маркой и «Для передачи»"
    ),
    "bad_title_folder": (
        "Положите комплект в папку четырёхзначного титула (не в служебную папку)"
    ),
    "gate_under_title": "Добавьте папку марки между титулом и «Для передачи»",
    "bad_mark_folder": (
        "Переименуйте папку марки (например 11_SKUD), "
        "не используйте МАРКАП, BBB или имя файла AGCC.…"
    ),
    "working_folder": (
        "Перенесите файлы из рабочей/резервной папки в «Для передачи» / NN_рев.…"
    ),
    "extra_before_gate": "Уберите лишний уровень между маркой и «Для передачи»",
    "bad_package_name": (
        "Переименуйте папку передачи в NN_рев.… и положите её в «Для передачи»"
    ),
    "other": "Приведите путь к ТИТУЛ / МАРКА / «Для передачи» / NN_рев.…",
}
# ``03_KSB4`` / ``14_POS``: an optional NN prefix in front of the mark token.
_MARK_LEVEL_RE = re.compile(r"^\s*(?:\d{1,2}[_.\-\s]+)?(?P<mark>.+?)\s*$")
# ``12_POS1`` / ``04-SOT``: required 1–2 digit prefix, remainder is the mark.
_MARK_FOLDER_RE = re.compile(
    r"^(?P<seq>\d{1,2})[_.\-\s]+(?P<mark>.+)$",
    re.IGNORECASE,
)
_MARK_CORE_RE = re.compile(r"^(?P<letters>[a-z]+)(?P<digits>\d*)$")
_FOLDER_SPLIT_RE = re.compile(r"[_.\-\s]+")
_TITLE_PREFIX_RE = re.compile(r"^\s*(?P<title>\d{4})\b")
_PLAUSIBLE_MARK_RE = re.compile(
    r"^(?:\d{1,2}[_.\-\s]+)?[A-Za-zА-Яа-яЁё][A-Za-zА-Яа-яЁё0-9]*"
    r"(?:[_\-][A-Za-zА-Яа-яЁё0-9]+)*$"
)
_WORKING_FOLDER_RE = re.compile(
    r"резерв|рабоч|в[_\s]?работе|отработ|замечан|тест|\btest\b|"
    r"копия|\bcopy\b|устарел|корректир|чернов|архив|"
    r"(?:^|[_\-\s])v[_\s]?\d",
    re.IGNORECASE,
)
_PACKAGE_DATE_RE = re.compile(r"20\d{2}[._-]\d{2}[._-]\d{2}")
_MISNAMED_MEDIA_REV_RE = re.compile(
    r"^(?:pdf|dwg|bbb|пдф)[_ ]*rev",
    re.IGNORECASE,
)
_NON_MARK_REMAINDERS = frozenset({"bbb", "pdf", "dwg", "пдф", "sq"})
_BAD_MARK_FOLDERS = frozenset({"маркап"})


def _windows_path_parts(path: str) -> tuple[str, ...]:
    normalized = path.replace("/", "\\")
    return PureWindowsPath(normalized).parts


def _join_windows_parts(parts: tuple[str, ...] | list[str]) -> str:
    if not parts:
        return ""
    return str(PureWindowsPath(*parts))


def _relative_windows_parts(
    path: str, source_root: str
) -> tuple[str, ...] | None:
    path_parts = _windows_path_parts(path)
    root_parts = _windows_path_parts(source_root)
    if not path_parts or not root_parts or len(path_parts) < len(root_parts):
        return None
    for path_part, root_part in zip(root_parts, path_parts):
        if path_part.casefold() != root_part.casefold():
            return None
    return path_parts[len(root_parts) :]


def path_is_as_build(path: str) -> bool:
    """Return whether any directory segment of ``path`` is as-build.

    Lexical and UNC-safe; does not touch the disk. The filename is ignored.
    Uses :data:`_AS_BUILD_RE` (``as-built`` / ``as-build`` / исполнительн…).

    Args:
        path: Local or UNC file path.

    Returns:
        ``True`` when a parent folder name matches the as-build pattern.
    """

    raw = (path or "").strip()
    if not raw:
        return False
    parts = _windows_path_parts(raw)
    if len(parts) < 2:
        return False
    for segment in parts[:-1]:
        if _AS_BUILD_RE.search(normalize_unicode_dashes(segment)):
            return True
    return False


def has_canonical_rd_issued_path(path: str, source_root: str | Path) -> bool:
    """Return whether a file sits in ``title/mark/gate/NN_`` under RD root.

    After dropping the filename and an optional trailing PDF/DWG/BBB/ПДФ
    folder, the path relative to ``source_root`` must be exactly four
    segments: a four-digit title folder, a mark folder, a transfer gate
    (``передач`` / ``отправк``, typically ``Для передачи``), and an issued
    ``NN_…`` package. A gate is mandatory.

    Args:
        path: Local or UNC file path.
        source_root: RD source root.

    Returns:
        ``True`` when the layout matches the issued-transfer contract.
    """

    raw = (path or "").strip()
    root = str(source_root or "").strip()
    if not raw or not root:
        return False
    relative = _relative_windows_parts(raw, root)
    if relative is None or len(relative) < 2:
        return False
    dirs = list(relative[:-1])
    if dirs and is_package_media_folder(dirs[-1]):
        dirs.pop()
    if len(dirs) != 4:
        return False
    title_folder, mark_folder, gate_folder, transfer_folder = dirs
    if not _TITLE_ONLY_FOLDER_RE.fullmatch(
        normalize_unicode_dashes(title_folder).strip()
    ):
        return False
    if is_transfer_gate_folder_name(mark_folder):
        return False
    if not is_transfer_gate_folder_name(gate_folder):
        return False
    return is_transfer_folder_name(transfer_folder, under_gate=True)


def kit_rd_mark_folder_from_path(
    path: str,
    source_root: str | Path,
    *,
    title: str,
) -> str:
    """Return the kit mark folder ``rd_root / title / mark`` from one RD file.

    Canonical issued files sit at
    ``<rd_root>/<4-digit title>/<mark>/<gate>/<NN_…>/[PDF|DWG]/file``.
    A kit-level RD rescan walks the **mark** directory (every gate and
    ``NN_…`` package), not a single transfer folder. A gate as the second
    segment (``title / Для передачи / …``) is not a mark folder.

    Lexical and UNC-safe; does not touch the disk.

    Args:
        path: Local or UNC file path.
        source_root: RD source root.
        title: Four-digit kit title; the first relative folder must match.

    Returns:
        Mark-folder path with Windows separators, or empty string when the
        file is not under ``source_root / title / <mark>``.
    """

    raw = (path or "").strip()
    root = str(source_root or "").strip()
    want_title = (title or "").strip()
    if not raw or not root or not want_title:
        return ""
    relative = _relative_windows_parts(raw, root)
    if relative is None or len(relative) < 2:
        return ""
    dirs = list(relative[:-1])
    if dirs and is_package_media_folder(dirs[-1]):
        dirs.pop()
    if len(dirs) < 2:
        return ""
    if dirs[0].strip().casefold() != want_title.casefold():
        return ""
    if is_transfer_gate_folder_name(dirs[1]):
        return ""
    root_parts = _windows_path_parts(root)
    if not root_parts:
        return ""
    return _join_windows_parts([*root_parts, dirs[0], dirs[1]])


def constructed_kit_rd_mark_folder(
    source_root: str | Path,
    title: str,
    mark: str,
) -> str:
    """Return ``rd_root / title / mark`` (Google mark as the folder name).

    Args:
        source_root: RD source root.
        title: Four-digit kit title.
        mark: Latin AGCC mark.

    Returns:
        Constructed mark-folder path, or empty string when any part is blank.
    """

    root = str(source_root or "").strip()
    title_text = (title or "").strip()
    mark_text = (mark or "").strip()
    if not root or not title_text or not mark_text:
        return ""
    root_parts = _windows_path_parts(root)
    if not root_parts:
        return ""
    return _join_windows_parts([*root_parts, title_text, mark_text])


def constructed_kit_rd_title_folder(
    source_root: str | Path,
    title: str,
) -> str:
    """Return ``rd_root / title`` (never a gate or the RD root itself).

    Args:
        source_root: RD source root.
        title: Four-digit kit title.

    Returns:
        Title-folder path, or empty string when either part is blank.
    """

    root = str(source_root or "").strip()
    title_text = (title or "").strip()
    if not root or not title_text:
        return ""
    root_parts = _windows_path_parts(root)
    if not root_parts:
        return ""
    return _join_windows_parts([*root_parts, title_text])


def unique_kit_rd_mark_folders(
    paths: Sequence[str],
    source_root: str | Path,
    *,
    title: str,
) -> tuple[str, ...]:
    """Return unique mark folders extracted from ``paths``.

    Args:
        paths: RD file paths for one kit.
        source_root: RD source root.
        title: Four-digit kit title.

    Returns:
        Distinct mark folders in first-seen order. Gate folders are omitted.
    """

    found: dict[str, str] = {}
    for path in paths:
        folder = kit_rd_mark_folder_from_path(path, source_root, title=title)
        if not folder:
            continue
        key = make_path_key(folder)
        if key not in found:
            found[key] = folder
    return tuple(found.values())


def record_has_canonical_layout(
    record: FileRecord | ParsedFile,
    rd_root: str | Path,
) -> bool:
    """Return whether a catalog file participates in the RD contour.

    Only :attr:`SourceKind.RD` files are checked against the issued-transfer
    layout (``title / mark / gate / NN_``). SQ and ROBOT records always
    participate because those sources use a different tree.

    Args:
        record: Persisted :class:`FileRecord` or a scan-time
            :class:`ParsedFile` (both expose ``source`` and ``path``).
        rd_root: RD source root from :class:`~rd_catalog.config.CatalogConfig`.

    Returns:
        ``True`` when the file should feed overlay and pipeline.
    """

    if record.source is not SourceKind.RD:
        return True
    return has_canonical_rd_issued_path(record.path, rd_root)


@dataclass(frozen=True, slots=True)
class LayoutViolation:
    """One present RD file that is not in the issued-transfer layout.

    Attributes:
        title: Four-digit title from the file record, if parsed.
        mark: Latin AGCC mark from the file record, if parsed.
        reason: Machine token. Existing tokens (``no_gate``,
            ``loose_in_mark``, ``extra_subfolder``, ``not_a_package``,
            ``too_shallow``, ``outside_gate``, ``package_nested``,
            ``duplicate_level``, ``other``) keep their meaning. Added
            tokens name the first break in the canonical chain:
            ``bad_title_folder``, ``gate_under_title``, ``bad_mark_folder``,
            ``working_folder``, ``extra_before_gate``, ``bad_package_name``.
            ``other`` is only a residual (path outside ``rd_root``, or an
            unforeseen shape). ``gate_nested`` was removed: after the
            title and mark-position checks it cannot fire.
        reason_label: Russian instruction for the author, not a diagnosis.
        path: File path as stored.
        package_hint: Nearest ``NN_`` folder name, or empty.
        is_mto: Whether the file is an MTO workbook. Those are the
            violations that silently change what export can offer.
    """

    title: str
    mark: str
    reason: str
    reason_label: str
    path: str
    package_hint: str
    is_mto: bool = False


def _directory_chain_after_media(
    path: str, source_root: str | Path
) -> tuple[str, ...] | None:
    raw = (path or "").strip()
    root = str(source_root or "").strip()
    if not raw or not root:
        return None
    relative = _relative_windows_parts(raw, root)
    if relative is None:
        return None
    dirs = list(relative[:-1])
    if dirs and is_package_media_folder(dirs[-1]):
        dirs.pop()
    return tuple(dirs)


def _mark_level_token(folder_name: str) -> str:
    """Return the mark part of a folder such as ``03_KSB4`` → ``KSB4``."""

    normalized = normalize_unicode_dashes(folder_name).strip()
    match = _MARK_LEVEL_RE.fullmatch(normalized)
    return (match.group("mark") if match else normalized).casefold()


def folder_matches_mark(folder_name: str, mark: str) -> bool:
    """Return whether a directory name is this kit's mark folder.

    Accepts ``POS1`` and ``12_POS1`` / ``04-SOT``. Does not match a different
    mark glued onto the same prefix.

    Args:
        folder_name: One path segment.
        mark: Latin AGCC mark.

    Returns:
        True when the folder belongs to ``mark``.
    """

    folded = normalize_unicode_dashes(folder_name).strip().casefold()
    want = mark.strip().casefold()
    if not folded or not want:
        return False
    if folded == want:
        return True
    match = _MARK_FOLDER_RE.match(folded)
    return bool(match and match.group("mark").casefold() == want)


def _latin_mark_core(text: str) -> tuple[str, str] | None:
    compact = _FOLDER_SPLIT_RE.sub(
        "", normalize_unicode_dashes(text).strip().casefold()
    )
    match = _MARK_CORE_RE.fullmatch(compact)
    if match is None:
        return None
    return match.group("letters"), match.group("digits")


def _mark_cores_compatible(
    folder_core: tuple[str, str], mark_core: tuple[str, str]
) -> bool:
    folder_letters, folder_digits = folder_core
    mark_letters, mark_digits = mark_core
    if folder_letters != mark_letters:
        return False
    if not folder_digits or not mark_digits:
        return True
    return folder_digits.startswith(mark_digits) or mark_digits.startswith(
        folder_digits
    )


def folder_hosts_filename_mark(folder_name: str, mark: str) -> bool:
    """Return whether an issued mark folder hosts this filename mark.

    ``POS1`` / ``12_POS1`` / ``04-SOT`` match that mark. ``06_KSB_21``
    hosts ``KSB`` and ``KSB1`` (same Latin stem; folder digits may extend
    the mark digits). ``SOS`` does not host ``SOT``; ``12_POS1`` does not
    host ``POS2``.

    Args:
        folder_name: Mark-level path segment.
        mark: Latin AGCC mark from the filename.

    Returns:
        True when the folder is this mark, not a stray sibling.
    """

    if folder_matches_mark(folder_name, mark):
        return True
    want = mark.strip()
    if not want:
        return False
    mark_core = _latin_mark_core(want)
    if mark_core is None:
        return False
    folded = normalize_unicode_dashes(folder_name).strip().casefold()
    if not folded:
        return False
    seen: set[str] = set()
    for token in (folded, *_FOLDER_SPLIT_RE.split(folded)):
        if not token or token in seen:
            continue
        seen.add(token)
        if token == want.casefold():
            return True
        core = _latin_mark_core(token)
        if core is not None and _mark_cores_compatible(core, mark_core):
            return True
    return False


def _repeats_title_or_mark(dirs: Sequence[str]) -> bool:
    """Return whether ``dirs[2]`` repeats the title or the mark level.

    Catches ``8630/KSB4/03_KSB4`` and ``1757/МАРКАП/1757_POS``, where an
    author nested a second title or mark folder inside the first one.
    """

    if len(dirs) < 3:
        return False
    title_match = _TITLE_PREFIX_RE.match(normalize_unicode_dashes(dirs[0]).strip())
    child = normalize_unicode_dashes(dirs[2]).strip()
    child_title = _TITLE_PREFIX_RE.match(child)
    if title_match and child_title:
        if child_title.group("title") == title_match.group("title"):
            return True
    return _mark_level_token(dirs[1]) == _mark_level_token(dirs[2])


def _folded_folder_name(folder_name: str) -> str:
    return normalize_unicode_dashes(folder_name).strip().casefold()


def _is_working_folder_name(folder_name: str) -> bool:
    """Return whether a folder is an obvious working/reserve dump."""

    folded = _folded_folder_name(folder_name)
    if not folded:
        return False
    if folded in {"раб", "rab"} or folded.startswith(("раб_", "rab_")):
        return True
    return bool(_WORKING_FOLDER_RE.search(folded))


def _is_plausible_mark_folder(folder_name: str) -> bool:
    """Return whether a folder looks like an AGCC mark directory."""

    normalized = normalize_unicode_dashes(folder_name).strip()
    folded = normalized.casefold()
    if not folded or folded in _BAD_MARK_FOLDERS:
        return False
    if is_transfer_gate_folder_name(normalized):
        return False
    if "agcc" in folded or " " in normalized:
        return False
    remainder = _mark_level_token(normalized)
    if remainder in _NON_MARK_REMAINDERS or is_package_media_folder(remainder):
        return False
    if not _PLAUSIBLE_MARK_RE.fullmatch(normalized):
        return False
    letters = [char for char in remainder if char.isalpha()]
    if (
        letters
        and all("а" <= char <= "я" or char == "ё" for char in letters)
        and len(letters) > 4
    ):
        return False
    return True


def _is_misnamed_package_folder(folder_name: str) -> bool:
    """Return whether a folder looks like a transfer that is not ``NN_…``."""

    folded = _folded_folder_name(folder_name)
    if not folded:
        return False
    if folded.startswith(("rev.", "rev_", "рев.", "рев ", "рев_")):
        return True
    if folded.startswith(("от ", "от_")):
        return True
    if folded.startswith("agcc."):
        return True
    if _PACKAGE_DATE_RE.search(folded):
        return True
    return bool(_MISNAMED_MEDIA_REV_RE.match(folded))


def layout_reason_label(reason: str) -> str:
    """Return the Russian author-facing instruction for a layout token.

    Args:
        reason: Machine token from :func:`classify_rd_layout_reason`.

    Returns:
        Instruction text, or the raw token when unknown. Empty ``reason``
        yields an empty string.
    """

    if not reason:
        return ""
    return _LAYOUT_REASON_LABELS.get(reason, reason)


def classify_rd_layout_reason(path: str, rd_root: str | Path) -> str:
    """Return a machine token for why an RD path is not canonical.

    Lexical; does not touch the disk. Empty string when the path matches
    :func:`has_canonical_rd_issued_path`. Classification names the *first*
    break in ``title / mark / gate / NN_`` and never changes *whether* a
    file is rejected.

    Existing tokens keep their meaning. New tokens split the former
    residual ``other`` bucket (and a few files whose earlier deviation
    was previously reported as ``loose_in_mark`` / ``extra_subfolder``).
    ``other`` is now only a residual: the path is outside ``rd_root``, or
    the shape matches no named rule.

    Args:
        path: Local or UNC file path.
        rd_root: RD source root.

    Returns:
        A layout token, or ``""`` when the path is canonical.
    """

    if has_canonical_rd_issued_path(path, rd_root):
        return ""
    dirs = _directory_chain_after_media(path, rd_root)
    if dirs is None:
        return "other"
    if len(dirs) <= 1:
        return "too_shallow"
    if not _TITLE_ONLY_FOLDER_RE.fullmatch(
        normalize_unicode_dashes(dirs[0]).strip()
    ):
        return "bad_title_folder"
    if is_transfer_gate_folder_name(dirs[1]):
        return "gate_under_title"
    if _is_working_folder_name(dirs[1]):
        return "working_folder"
    if not _is_plausible_mark_folder(dirs[1]):
        return "bad_mark_folder"
    if len(dirs) == 2:
        return "loose_in_mark"

    gate_idx = next(
        (
            index
            for index, segment in enumerate(dirs)
            if is_transfer_gate_folder_name(segment)
        ),
        None,
    )
    if gate_idx == 2:
        if len(dirs) == 3:
            return "too_shallow"
        if not is_transfer_folder_name(dirs[3], under_gate=True):
            return "not_a_package"
        if len(dirs) > 4:
            return "extra_subfolder"
        return "other"

    if _is_working_folder_name(dirs[2]):
        return "working_folder"
    if _repeats_title_or_mark(dirs):
        return "duplicate_level"
    if gate_idx is not None and gate_idx > 2:
        return "extra_before_gate"
    if is_transfer_folder_name(dirs[2], under_gate=True):
        return "no_gate"
    if any(is_transfer_folder_name(name, under_gate=True) for name in dirs[3:]):
        return "package_nested"
    if _is_misnamed_package_folder(dirs[2]):
        return "bad_package_name"
    if is_package_media_folder(dirs[2]):
        return "extra_before_gate"
    return "outside_gate"


def nearest_issued_package_hint(path: str) -> str:
    """Return the deepest ``NN_`` folder name on ``path``, if any.

    Args:
        path: Local or UNC file path.

    Returns:
        Folder name such as ``05_рев.0_…``, or empty string.
    """

    parts = list(_windows_path_parts(path))
    if parts:
        parts.pop()
    hint = ""
    for index, segment in enumerate(parts):
        ancestors = parts[:index]
        under_gate = any(is_transfer_gate_folder_name(name) for name in ancestors)
        if is_transfer_folder_name(segment, under_gate=under_gate):
            hint = segment
    return hint


def list_layout_violations(
    *,
    records: Sequence[FileRecord],
    rd_root: str | Path,
) -> tuple[LayoutViolation, ...]:
    """Return present RD files that are not in the issued-transfer layout.

    Qt-free. Classification is lexical from ``records`` (no disk IO).

    Args:
        records: Catalog file rows, typically ``list_files()``.
        rd_root: RD source root from :class:`~rd_catalog.config.CatalogConfig`.

    Returns:
        Violations sorted by title, mark, reason, then path.
    """

    rows: list[LayoutViolation] = []
    for record in records:
        if not record.present or record.source is not SourceKind.RD:
            continue
        if record_has_canonical_layout(record, rd_root):
            continue
        reason = classify_rd_layout_reason(record.path, rd_root)
        title = str(record.data.get("title") or "").strip()
        mark = str(record.data.get("mark") or "").strip()
        rows.append(
            LayoutViolation(
                title=title,
                mark=mark,
                reason=reason,
                reason_label=layout_reason_label(reason),
                path=record.path,
                package_hint=nearest_issued_package_hint(record.path),
                is_mto=_file_kind_from_name(Path(record.path).name)
                is FileKind.MTO_XLSX,
            )
        )
    rows.sort(
        key=lambda item: (
            item.title.casefold(),
            item.mark.casefold(),
            item.reason,
            item.path.casefold(),
        )
    )
    return tuple(rows)


def _issued_path_layout_dirs(path: str) -> tuple[str, str, str, str] | None:
    """Return ``title / mark / gate / NN_`` from a canonical issued path."""

    raw = (path or "").strip()
    if not raw:
        return None
    parts = list(_windows_path_parts(raw))
    if not parts:
        return None
    leaf = parts[-1]
    under_gate = any(is_transfer_gate_folder_name(name) for name in parts[:-1])
    if not is_transfer_folder_name(leaf, under_gate=under_gate):
        parts.pop()
        if parts and is_package_media_folder(parts[-1]):
            parts.pop()
    if len(parts) < 4:
        return None
    title_folder, mark_folder, gate_folder, transfer_folder = parts[-4:]
    if not _TITLE_ONLY_FOLDER_RE.fullmatch(
        normalize_unicode_dashes(title_folder).strip()
    ):
        return None
    if is_transfer_gate_folder_name(mark_folder):
        return None
    if not is_transfer_gate_folder_name(gate_folder):
        return None
    if not is_transfer_folder_name(transfer_folder, under_gate=True):
        return None
    return (
        title_folder.strip(),
        mark_folder.strip(),
        gate_folder.strip(),
        transfer_folder.strip(),
    )


def issued_path_title_folder(path: str) -> str:
    """Return the 4-digit title folder of a canonical issued RD path.

    Lexical; no disk IO. After dropping a filename and an optional
    PDF/DWG/BBB/ПДФ segment, the last four directories must be
    ``title / mark / gate / NN_``. A package folder (leaf is already
    ``NN_…``) is accepted the same way.

    Args:
        path: Local or UNC file path, or an issued package directory.

    Returns:
        Title folder name, or empty when the path is not that shape.
    """

    dirs = _issued_path_layout_dirs(path)
    return dirs[0] if dirs else ""


def issued_path_mark_folder(path: str) -> str:
    """Return the mark folder of a canonical issued RD path.

    Lexical; no disk IO. Same layout as :func:`issued_path_title_folder`.

    Args:
        path: Local or UNC file path, or an issued package directory.

    Returns:
        Mark folder name (``SOS``, ``06_KSB_21``, ``22_POS2``), or empty
        when the path is not that shape.
    """

    dirs = _issued_path_layout_dirs(path)
    return dirs[1] if dirs else ""


def issued_package_dir(path: str) -> str:
    """Return the issued package folder for a file path (lexical, no disk IO).

    Splits a local or UNC path (keeping ``\\\\server\\share``), drops the
    filename, then a trailing PDF/DWG/BBB/ПДФ segment. If the remaining leaf
    is an issued transfer (``is_transfer_folder_name`` with ``under_gate``
    from ancestors), that folder is returned; otherwise the parent after
    stripping the media folder (robot / SQ / mark folder).

    Args:
        path: Local or UNC file path.

    Returns:
        Package directory using Windows separators, or empty string when
        nothing remains.
    """

    raw = (path or "").strip()
    if not raw:
        return ""
    parts = list(_windows_path_parts(raw))
    if not parts:
        return ""
    parts.pop()
    if parts and is_package_media_folder(parts[-1]):
        parts.pop()
    if not parts:
        return ""
    leaf = parts[-1]
    under_gate = any(is_transfer_gate_folder_name(name) for name in parts[:-1])
    if is_transfer_folder_name(leaf, under_gate=under_gate):
        return _join_windows_parts(parts)
    return _join_windows_parts(parts)


def path_relative_to_transfer_gate(path: str) -> str:
    """Return the file path from the transfer-gate folder (lexical, no IO).

    The gate segment itself (``Для передачи`` / ``На_отправку`` / …) is
    omitted, so a canonical file becomes ``10_рев.…\\PDF\\file.pdf``.
    Without a gate, the path starts at the issued ``NN_…`` folder when
    that folder is recognized; otherwise only the filename.

    Args:
        path: Local or UNC file path.

    Returns:
        Relative Windows path, or empty string when ``path`` is empty.
    """

    raw = (path or "").strip()
    if not raw:
        return ""
    parts = list(_windows_path_parts(raw))
    if not parts:
        return ""
    filename = parts[-1]
    dirs = parts[:-1]
    gate_idx = -1
    for index, name in enumerate(dirs):
        if is_transfer_gate_folder_name(name):
            gate_idx = index
    if gate_idx >= 0:
        after = dirs[gate_idx + 1 :]
        if not after:
            return filename
        return _join_windows_parts([*after, filename])
    under_gate = False
    pkg_idx = -1
    for index, name in enumerate(dirs):
        if is_transfer_gate_folder_name(name):
            under_gate = True
        if is_transfer_folder_name(name, under_gate=under_gate):
            pkg_idx = index
    if pkg_idx >= 0:
        return _join_windows_parts([*dirs[pkg_idx:], filename])
    return filename


_SOURCE_EDITABLE_SUFFIXES = frozenset(
    {".xlsx", ".xls", ".dwg", ".docx", ".doc", ".zip", ".7z", ".7zip"}
)


def _file_kind_from_name(name: str) -> FileKind | None:
    folded = name.casefold()
    suffix = Path(name).suffix.casefold()
    if suffix == ".pdf":
        return FileKind.PDF
    if suffix == ".xlsx" and ("mto" in folded or "мто" in folded):
        return FileKind.MTO_XLSX
    if suffix in _SOURCE_EDITABLE_SUFFIXES:
        return FileKind.SOURCE_EDITABLE
    return None


def _agcc_filename_parts(file_name: str):
    normalized = normalize_unicode_dashes(Path(file_name).name)
    return AgccFilenamePatterns.parse_strict(
        normalized
    ) or AgccFilenamePatterns.parse_loose(normalized)


def matches_agcc_filename(name: str) -> bool:
    """Return whether a file name matches the AGCC working-document mask.

    Uses the same ``parse_strict`` then ``parse_loose`` sequence as
    :func:`parse_catalog_file`. Unicode dashes are normalized first.

    Args:
        name: File name or path; only the final segment is tested.

    Returns:
        ``True`` when :class:`~utils.file_name_converts.AgccFilenamePatterns`
        accepts the name.
    """

    return _agcc_filename_parts(name) is not None


def parse_catalog_file(
    path: str | Path,
    source: SourceKind,
    *,
    size: int,
    mtime_ns: int,
    transfer: TransferMetadata | None = None,
) -> ParsedFile:
    """Parse one already-statted catalog candidate.

    Args:
        path: Original local or UNC file path.
        source: Catalog source.
        size: File size from ``stat``.
        mtime_ns: Nanosecond modification time from ``stat``.
        transfer: Metadata of the containing transfer folder.

    Returns:
        A parsed or gracefully unparsed typed record.
    """

    original_path = str(path)
    original_name = Path(path).name
    normalized_name = normalize_unicode_dashes(original_name)
    file_kind = _file_kind_from_name(normalized_name)
    if file_kind is None:
        raise ValueError(f"Unsupported catalog file extension: {original_name}")

    parts = _agcc_filename_parts(normalized_name)
    if parts is None:
        return ParsedFile(
            path=original_path,
            path_key=make_path_key(original_path),
            name=original_name,
            source=source,
            file_kind=file_kind,
            size=size,
            mtime_ns=mtime_ns,
            parse_status=ParseStatus.UNPARSED_FILE,
            transfer=transfer,
            extension=Path(normalized_name).suffix.lstrip(".").casefold() or None,
            parse_error="name does not match AgccFilenamePatterns",
        )

    title, mark = parts.title_system.split("-", 1)
    tail = AgccFilenamePatterns.split_revision_tail(parts.revision_tail)
    revision: str | None = None
    appendix: str | None = None
    language: str | None = None
    extension = (
        tail.file_type.casefold()
        if tail
        else (parts.simple_ext or Path(normalized_name).suffix).lstrip(".").casefold()
    )
    parse_error: str | None = None

    if tail:
        revision = tail.rev_sheet.split("-", 1)[0]
        appendix = tail.an
        language = tail.lang

    if file_kind is FileKind.MTO_XLSX:
        if not parts.discipline_block.casefold().startswith("mto"):
            parse_error = "XLSX discipline is not MTO"
        elif parts.revision_tail:
            try:
                mto_revision, mto_appendix = (
                    parse_agcc_mto_xlsx_revision_for_chain(normalized_name)
                )
            except ValueError as exc:
                parse_error = str(exc)
            else:
                normalized_appendix = (
                    mto_appendix.removeprefix("AN") if mto_appendix else None
                )
                if (revision, appendix) != (mto_revision, normalized_appendix):
                    parse_error = "MTO revision parsers disagree"
                revision, appendix = mto_revision, normalized_appendix

    if file_kind is not FileKind.SOURCE_EDITABLE:
        if extension != file_kind.value.removeprefix("mto_"):
            expected = "xlsx" if file_kind is FileKind.MTO_XLSX else "pdf"
            if extension != expected:
                parse_error = (
                    f"parsed extension {extension!r} does not match {expected!r}"
                )

    return ParsedFile(
        path=original_path,
        path_key=make_path_key(original_path),
        name=original_name,
        source=source,
        file_kind=file_kind,
        size=size,
        mtime_ns=mtime_ns,
        parse_status=(
            ParseStatus.UNPARSED_FILE if parse_error else ParseStatus.PARSED
        ),
        contract=parts.contract,
        title_system=parts.title_system,
        title=title,
        mark=mark,
        discipline_block=parts.discipline_block,
        core_stem=parts.core_stem,
        revision=revision,
        appendix=appendix,
        language=language,
        extension=extension,
        transfer=transfer,
        parse_error=parse_error,
    )
