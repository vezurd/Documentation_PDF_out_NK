"""Row classification for TSD packing-list Excel rows (TsdPacking table type)."""

from __future__ import annotations

import re

from base.base_classes import RowStd, RowType
from base.tables_columns import CODE, NAME, SPECIFICATION_NAME, TAGS, TYPE_MARK, UNITS, VALUES, VENDOR

# Signature / footer placeholders: "___", "____", … — not a lone "_".
_UNDERSCORE_PLACEHOLDER_RE = re.compile(r"^_{3,}$")
# Units cell filled with a year only, e.g. "2025г." / "2025 г".
_YEAR_ONLY_RE = re.compile(r"^\d{4}\s*г\.?$", re.IGNORECASE)


def _cell_text(row: RowStd, key: str) -> str:
    el = row.el.get(key)
    if el is None:
        return ""
    v = el.value
    if v is None:
        return ""
    if isinstance(v, list):
        return "" if len(v) == 0 else str(v)
    return str(v)


def _is_underscore_placeholder(text: str) -> bool:
    """True for ``___`` / longer underscore-only cells (single ``_`` is kept)."""
    return bool(_UNDERSCORE_PLACEHOLDER_RE.fullmatch(text.strip()))


def get_row_type(row: RowStd):
    """Classify TSD packing row: header/boilerplate, position, empty, or other.

    Position requires non-empty CODE, NAME, VALUES, UNITS; TAGS and VENDOR optional
    (manufacturer column is often blank while country is filled elsewhere).
    CODE format is not restricted (BCC, ECBL, …); only header/banner/signature
    rows are rejected.
    """
    code = _cell_text(row, CODE)
    name = _cell_text(row, NAME)
    values = _cell_text(row, VALUES)
    units = _cell_text(row, UNITS)
    vendor = _cell_text(row, VENDOR)
    spec = _cell_text(row, SPECIFICATION_NAME)

    # Template column-index banner (CODE=15, SPEC=16, NAME=21, …).
    if (
        code.strip().isdigit()
        and spec.strip().isdigit()
        and name.strip().isdigit()
        and len(code.strip()) <= 3
    ):
        return RowType.other_row

    code_u = code.strip().upper()
    name_u = name.strip().upper()
    values_u = values.strip().upper()
    units_u = units.strip().upper()
    vendor_u = vendor.strip().upper()

    # Header / banner / signature cells — avoid substring matches inside product
    # names (e.g. «…с количеством и сечением жил…» must stay a position_row).
    if (
        code_u in ("CODE", "КОД", "КОД РД", "BCC")
        or code_u.startswith("КОД РД")
        or "CODE RD" in code_u
        or code_u.startswith("СПЕЦИФИКАЦИИ")
        or "ВЕДУЩИЙ СПЕЦИАЛИСТ" in code_u
        # STD packing-list letterhead (колонка B / CODE).
        or code_u.startswith("ГРУЗООТПРАВИТЕЛЬ")
        or code_u.startswith("ГРУЗОПОЛУЧАТЕЛЬ")
        or code_u.startswith("КОНТАКТНЫЕ ДАННЫЕ")
        or code_u.startswith("ОСНОВАНИЕ ПОСТАВКИ")
        or code_u.startswith("УПАКОВОЧНЫЙ ЛИСТ")
        or _is_underscore_placeholder(code)
        or name_u in ("NAME", "NAME + TYPE_MARK", "NAME;TYPE_MARK")
        or name_u.startswith("НАИМЕНОВАНИЕ")
        or name_u.startswith("NAME ")
        or name_u.startswith("NAME;")
        or _is_underscore_placeholder(name)
        or values_u in ("VALUES", "QTY", "КОЛ-ВО", "КОЛИЧЕСТВО")
        or values_u.startswith("КОЛИЧЕСТВО")
        or values_u.startswith("QUANTITY")
        or values_u.startswith("ЕДИНИЦА ИЗМЕРЕНИЯ")
        # STD letterhead labels in VALUES column (H/I spill).
        or values_u.startswith("ПРОДАВЕЦ")
        or values_u.startswith("ПОКУПАТЕЛЬ")
        or values_u.startswith("ДАТА ДОКУМЕНТА")
        or values_u.startswith("МЕСТО НАЗНАЧЕНИЯ")
        or values_u.startswith("ЗИП /")
        or values_u.startswith("ЗИП/")
        or values_u.startswith("SELLER")
        or values_u.startswith("BUYER")
        or values_u.startswith("DESTINATION")
        or _is_underscore_placeholder(values)
        or units_u in ("UNITS", "ЕД. ИЗМ.", "ЕД.ИЗМ.", "UOM")
        or units_u.startswith("ЕДИНИЦА ИЗМЕРЕНИЯ")
        or units_u.startswith("UNIT OF MEASUREMENT")
        or units_u.startswith("МАССА")
        or _YEAR_ONLY_RE.fullmatch(units.strip()) is not None
        or _is_underscore_placeholder(units)
        or vendor_u in ("VENDOR", "ПОСТАВЩИК", "ИЗГОТОВИТЕЛЬ")
        or vendor_u.startswith("ЗАВОД ИЗГОТОВИТЕЛЬ")
        or vendor_u.startswith("MANUFACTURER")
        or _YEAR_ONLY_RE.fullmatch(vendor.strip()) is not None
    ):
        return RowType.other_row

    column_dict = {
        CODE: 1,
        NAME: 1,
        TAGS: 3,
        TYPE_MARK: 3,
        UNITS: 1,
        VALUES: 1,
        VENDOR: 3,
    }
    if RowType.check_row(row, column_dict):
        return RowType.position_row

    empty_dict = {
        CODE: 0,
        NAME: 0,
        TAGS: 0,
        TYPE_MARK: 0,
        UNITS: 0,
        VALUES: 0,
        VENDOR: 0,
    }
    if RowType.check_row(row, empty_dict):
        return RowType.empty_row

    return RowType.other_row
