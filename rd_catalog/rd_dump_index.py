"""RD-tree MTO dump index (Qt-free).

Same filename contract as АН (``parse_an_mto_file``). Layout flags are
lexical against ``rd_root`` and are not stored in SQLite.
"""

from __future__ import annotations

from pathlib import Path

from rd_catalog.an_compare import AN_AGREED_HEADER
from rd_catalog.an_index import AnMtoFile, parse_an_mto_file
from rd_catalog.parse import (
    classify_rd_layout_reason,
    has_canonical_rd_issued_path,
    layout_reason_label,
)

parse_rd_dump_mto_file = parse_an_mto_file

CANON_HEADER = "Канон."
LAYOUT_HEADER = "Раскладка"
RD_DUMP_CANON_YES_FILL = "#E2F2E1"
RD_DUMP_CANON_NO_FILL = "#F7E8BE"

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
    "LAYOUT_HEADER",
    "RD_DUMP_AGREED_ROW_HEADERS",
    "RD_DUMP_CANON_NO_FILL",
    "RD_DUMP_CANON_YES_FILL",
    "RD_DUMP_HEADERS",
    "parse_rd_dump_mto_file",
    "rd_dump_is_canonical",
    "rd_dump_layout_short",
    "rd_dump_layout_token",
    "rd_dump_layout_tooltip",
]
