"""UTF-8 dump of RFQ TPK rows not treated as ``position_row`` (like DS merge txt)."""

from __future__ import annotations

from pathlib import Path

from prettytable import PrettyTable

from base.base_classes import RowStd, RowType
from base.tables_columns import (
    CODE,
    DS_CODE_1C,
    DS_SPECIFICATION,
    DS_SYSTEM,
    DS_TITLE,
    NAME,
    NUMBERS,
    ROW_TYPE,
    TAGS,
    TYPE_MARK,
    UNITS,
    VALUES,
)
import utils.path

# UTF-8 рядом с RFQ xlsx / отчётом ДС_МТО: все строки, кроме position_row.
RFQ_TPK_NOT_POSITION_ROWS_FILENAME = "RFQ_TPK_not_position_rows.txt"

_RFQ_DUMP_KEYS: tuple[str, ...] = (
    NUMBERS,
    DS_TITLE,
    DS_SYSTEM,
    DS_SPECIFICATION,
    TAGS,
    DS_CODE_1C,
    CODE,
    NAME,
    TYPE_MARK,
    VALUES,
    UNITS,
    ROW_TYPE,
)

_RFQ_DUMP_HEADERS: tuple[str, ...] = (
    "NUMBERS",
    "DS_TITLE",
    "DS_SYSTEM",
    "DS_SPECIFICATION",
    "TAGS",
    "DS_CODE_1C",
    "CODE",
    "NAME",
    "TYPE_MARK",
    "VALUES",
    "UNITS",
    "ROW_TYPE",
)

_CELL_MAX_LEN = 40


def _dump_cell_value(row: RowStd, key: str):
    if key == ROW_TYPE:
        return row.row_type
    return row.get_value(key)


def _cell_str(value: object, max_len: int = _CELL_MAX_LEN) -> str:
    if value is None:
        s = ""
    else:
        s = str(value).replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    if len(s) <= max_len:
        return s
    if max_len <= 1:
        return s[:max_len]
    return s[: max_len - 1] + "…"


def write_rfq_non_position_rows_txt(
    rfq_rows: list[RowStd],
    out_dir: str,
) -> str | None:
    """Write UTF-8 PrettyTable: all RFQ rows except ``position_row``.

    Includes ``empty_row``, ``head_row``, ``other_row``, and any other non-position types.

    Args:
        rfq_rows: Rows after ``load_rfq_tpk_data`` / ``get_std_from_excel_file``.
        out_dir: Directory next to RFQ xlsx or grouped output (RFQ folder).

    Returns:
        Path to the written ``.txt`` file, or ``None`` if there were no rows to dump.
    """
    picked = [r for r in rfq_rows if r.row_type != RowType.position_row]

    if not picked:
        print("[RFQ TPK] Дамп строк (не position_row): нет строк — файл не создаётся.")
        return None

    table = PrettyTable()
    table.field_names = list(_RFQ_DUMP_HEADERS)
    for fn in table.field_names:
        table.align[fn] = "l"
    table.max_width = _CELL_MAX_LEN

    for row in picked:
        table.add_row(
            [_cell_str(_dump_cell_value(row, key)) for key in _RFQ_DUMP_KEYS]
        )

    position_count = sum(1 for r in rfq_rows if r.row_type == RowType.position_row)
    type_counts: dict[str, int] = {}
    for row in picked:
        type_counts[row.row_type] = type_counts.get(row.row_type, 0) + 1
    types_summary = ", ".join(f"{k}={v}" for k, v in sorted(type_counts.items()))
    header_lines = [
        "RFQ TPK: all rows except position_row (not used in DS vs MTO vs RFQ compare)",
        f"total_loaded_rows: {len(rfq_rows)}",
        f"position_row: {position_count}",
        f"dumped_rows: {len(picked)} ({types_summary})",
        "",
    ]
    body = table.get_string()
    text_out = "\n".join(header_lines) + body + "\n"

    print(text_out, end="")

    utils.path.make_dir(out_dir)
    out_path = str(Path(out_dir) / RFQ_TPK_NOT_POSITION_ROWS_FILENAME)
    Path(out_path).write_text(text_out, encoding="utf-8")
    print(
        f"[RFQ TPK] Список строк (не position_row): "
        f"{len(picked)} шт. -> {out_path}"
    )
    return out_path
