"""RD-tree dump index for MTO xlsx and OD doc/docx (Qt-free).

MTO names reuse ``parse_an_mto_file``. OD uses the same AGCC mask with
discipline ``od*`` and ``.doc`` / ``.docx``. Layout flags are lexical
against ``rd_root`` and are not stored in SQLite.
"""

from __future__ import annotations

from pathlib import Path

from utils.file_name_converts import AgccFilenamePatterns

from rd_catalog.an_compare import AN_AGREED_HEADER
from rd_catalog.an_index import AnMtoFile, parse_an_mto_file
from rd_catalog.kits import format_revision
from rd_catalog.models import make_path_key
from rd_catalog.parse import (
    _agcc_filename_parts,
    classify_rd_layout_reason,
    has_canonical_rd_issued_path,
    layout_reason_label,
    matches_agcc_filename,
    normalize_unicode_dashes,
)

KIND_HEADER = "Вид"
KIND_MTO = "MTO"
KIND_OD = "OD"
CANON_HEADER = "Канон."
LAYOUT_HEADER = "Раскладка"
RD_DUMP_CANON_YES_FILL = "#E2F2E1"
RD_DUMP_CANON_NO_FILL = "#F7E8BE"

_OD_SUFFIXES = frozenset({".doc", ".docx"})

_LAYOUT_SHORT: dict[str, str] = {
    "no_gate": "нет шлюза",
    "loose_in_mark": "в папке марки",
    "extra_subfolder": "лишняя папка",
    "not_a_package": "не пакет NN",
    "too_shallow": "короткий путь",
    "outside_gate": "вне шлюза",
    "package_nested": "пакет вложен",
    "duplicate_level": "повтор титула/марки",
    "bad_title_folder": "не титул",
    "gate_under_title": "шлюз без марки",
    "bad_mark_folder": "не марка",
    "working_folder": "рабочая папка",
    "extra_before_gate": "лишнее до шлюза",
    "bad_package_name": "имя пакета",
    "other": "прочее",
}

RD_DUMP_AGREED_ROW_HEADERS = (
    "Титул",
    "Марка",
    "Ревизия",
    KIND_HEADER,
    CANON_HEADER,
    LAYOUT_HEADER,
    AN_AGREED_HEADER,
    "Имя",
    "Дата",
    "Папка",
    "Путь",
)

RD_DUMP_HEADERS = (
    "Титул",
    "Марка",
    "Ревизия",
    KIND_HEADER,
    CANON_HEADER,
    LAYOUT_HEADER,
    AN_AGREED_HEADER,
    "vs Авто МТО",
    "vs MTO РД",
    "vs Робот",
    "vs Выдача",
    "vs F",
    "vs SQ",
    "Имя",
    "Дата",
    "Папка",
    "Путь",
)


def parse_rd_dump_file(
    path: str | Path,
    *,
    size: int,
    mtime_ns: int,
) -> AnMtoFile | None:
    """Parse one RD dump candidate from its name and stats.

    Accepts AGCC MTO ``.xlsx`` (same rules as АН) and AGCC OD ``.doc`` /
    ``.docx``. The file body is not opened. Junk names return ``None``.

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
    suffix = Path(normalized_name).suffix.casefold()
    if suffix == ".xlsx":
        return parse_an_mto_file(path, size=size, mtime_ns=mtime_ns)
    if suffix not in _OD_SUFFIXES:
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


parse_rd_dump_mto_file = parse_rd_dump_file


def rd_dump_kind(file: AnMtoFile) -> str:
    """Return ``OD`` or ``MTO`` for a dump row.

    Args:
        file: Parsed dump record.

    Returns:
        ``OD`` when the discipline block starts with ``od``, else ``MTO``.
    """

    if (file.discipline_block or "").casefold().startswith("od"):
        return KIND_OD
    return KIND_MTO


def rd_dump_is_od(file: AnMtoFile) -> bool:
    """Return whether this dump row is an OD Word file.

    Args:
        file: Parsed dump record.

    Returns:
        True when :func:`rd_dump_kind` is ``OD``.
    """

    return rd_dump_kind(file) == KIND_OD


def rd_dump_is_canonical(path: str, rd_root: str | Path) -> bool:
    """Return whether ``path`` is an issued ``title/mark/gate/NN_`` RD file.

    Args:
        path: Local or UNC file path.
        rd_root: Catalog ``rd_root``.

    Returns:
        ``True`` for a canonical issued path under ``rd_root``.
    """

    return has_canonical_rd_issued_path(path, rd_root)


def rd_dump_layout_token(path: str, rd_root: str | Path) -> str:
    """Return ``classify_rd_layout_reason`` for a dump file.

    Args:
        path: Local or UNC file path.
        rd_root: Catalog ``rd_root``.

    Returns:
        Empty string when the path is canonical, otherwise a layout token.
    """

    return classify_rd_layout_reason(path, rd_root)


def rd_dump_layout_short(token: str) -> str:
    """Return the compact table label for a layout token.

    Args:
        token: ``classify_rd_layout_reason`` value, or empty.

    Returns:
        ``канон`` when ``token`` is empty, else a short Russian tag.
    """

    if not token:
        return "канон"
    return _LAYOUT_SHORT.get(token, token)


def rd_dump_layout_tooltip(token: str) -> str:
    """Return the author-facing layout instruction, or a canonical note.

    Args:
        token: ``classify_rd_layout_reason`` value, or empty.

    Returns:
        Tooltip text for the «Раскладка» cell.
    """

    if not token:
        return "Канонический путь: титул / марка / шлюз / NN_ / [PDF|DWG]"
    return layout_reason_label(token)


__all__ = [
    "AnMtoFile",
    "CANON_HEADER",
    "KIND_HEADER",
    "KIND_MTO",
    "KIND_OD",
    "LAYOUT_HEADER",
    "RD_DUMP_AGREED_ROW_HEADERS",
    "RD_DUMP_CANON_NO_FILL",
    "RD_DUMP_CANON_YES_FILL",
    "RD_DUMP_HEADERS",
    "parse_rd_dump_file",
    "parse_rd_dump_mto_file",
    "rd_dump_is_canonical",
    "rd_dump_is_od",
    "rd_dump_kind",
    "rd_dump_layout_short",
    "rd_dump_layout_token",
    "rd_dump_layout_tooltip",
]
