"""Row classification for TSD SO - PL packing sheets (госфин / wide layout)."""

from __future__ import annotations

from base.base_classes import RowStd, RowType
from base.tables_columns import CODE, NAME, SPECIFICATION_NAME, TAGS, UNITS, VALUES, VENDOR


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


def get_row_type(row: RowStd):
    """Classify SO - PL row: header/boilerplate, position, empty, or other.

    Position requires non-empty CODE, NAME, VALUES, UNITS; TAGS and VENDOR optional.
    Rejects the wide-format header row (``Name of Zip`` / ``Name of document`` / ``PO item``).
    """
    code = _cell_text(row, CODE)
    name = _cell_text(row, NAME)
    values = _cell_text(row, VALUES)
    units = _cell_text(row, UNITS)
    vendor = _cell_text(row, VENDOR)
    spec = _cell_text(row, SPECIFICATION_NAME)

    code_u = code.strip().upper()
    name_u = name.strip().upper()
    spec_u = spec.strip().upper()
    values_u = values.strip().upper()
    units_u = units.strip().upper()
    vendor_u = vendor.strip().upper()

    if (
        code_u in ("PO ITEM", "CODE", "КОД", "BCC")
        or "NAME OF ZIP" in code_u
        or "NAME OF DOCUMENT" in spec_u
        or spec_u in ("NAME OF DOCUMENT", "NAME OF ZIP")
        or "RUSSIAN TRANSLATION" in name_u
        or "ITEM DESCRIPTION" in name_u
        or values_u in ("QUANTITY", "QTY", "VALUES", "КОЛИЧЕСТВО")
        or "QUANTITY" in values_u
        or units_u in ("UNITS", "UNIT", "UOM", "ЕД. ИЗМ.")
        or vendor_u in ("VENDOR", "ПОСТАВЩИК")
        or "TYPE OF FILE" in code_u
        # SO - PL letterhead / delivery block (cols G/N/O before table header).
        or code_u.startswith("FOR DELIVERY")
        or code_u.startswith("CONSIGNEE")
        or code_u.startswith("FINAL DESTINATION")
        or code_u.startswith("MARKS /")
        or code_u.startswith("MARKS/")
        or code_u.startswith("TYPE OF PACKAGE")
        or code_u.startswith("BUYER /")
        or code_u.startswith("BUYER/")
        or code_u.startswith("AMUR GAS")
        or code_u.startswith("AMUR GCC")
        or code_u.startswith("DDP")
        # Package / stackability / storage legend under the table.
        or code_u.startswith("001 =")
        or code_u.startswith("002 =")
        or code_u.startswith("003 =")
        or code_u.startswith("004 =")
        or code_u.startswith("005 =")
        or code_u.startswith("006 =")
        or name_u.startswith("STACKABILITY")
        or name_u.startswith("00 =")
        or name_u.startswith("01 =")
        or name_u.startswith("02 =")
        or values_u.startswith("PACKING LIST NO")
        or values_u.startswith("DATE /")
        or values_u.startswith("DATE/")
        or values_u.startswith("BASIS OF SALES")
        or values_u.startswith("CONTRACTOR PO")
        or values_u.startswith("PARTIAL SHIPMENT")
        or values_u.startswith("SHIPPING INVOICE")
        or values_u.startswith("SALES ORDER")
        or values_u.startswith("SELLER /")
        or values_u.startswith("SELLER/")
        or values_u.startswith("ADRESS")
        or values_u.startswith("ADDRESS")
        or values_u.startswith("АДРЕС")
    ):
        return RowType.other_row

    column_dict = {
        CODE: 1,
        NAME: 1,
        TAGS: 3,
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
        UNITS: 0,
        VALUES: 0,
        VENDOR: 0,
    }
    if RowType.check_row(row, empty_dict):
        return RowType.empty_row

    return RowType.other_row
