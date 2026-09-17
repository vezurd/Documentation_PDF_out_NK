"""Excel / openpyxl helpers."""

from __future__ import annotations

from typing import Any

import openpyxl


def get_text_from_excel(file_path):
    out_text = ""
    wb = openpyxl.load_workbook(file_path, True, rich_text=False)
    try:
        for sheet in wb:
            ws = wb.get_sheet_by_name(sheet)
            for row in ws.values:
                for value in row:
                    out_text += str(value) + "; "
        return out_text
    finally:
        wb.close()


def sanitize_cell_value_for_openpyxl(value: Any, *, enabled: bool) -> Any:
    """Strip characters that OOXML forbids in shared strings (openpyxl ``IllegalCharacterError``).

    Matches openpyxl's own check (control chars except tab/LF/CR). When ``enabled`` is False,
    returns ``value`` unchanged.

    Args:
        value: Cell value (typically ``str``, ``int``, ``float``, ``bool``, or ``None``).
        enabled: If False, no transformation.

    Returns:
        Sanitized value safe for ``openpyxl`` string cells, or the original non-string values.
    """
    if not enabled or value is None or not isinstance(value, str):
        return value
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

    return ILLEGAL_CHARACTERS_RE.sub("", value)
