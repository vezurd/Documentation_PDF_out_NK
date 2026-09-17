

from __future__ import annotations

from base.base_classes import RowStd, RowType
from base.tables_columns import *

# (column_key, substring) — подвал / преамбула (не заголовки столбцов таблицы).
_DS_SPEC_BOILERPLATE_PAIRS: tuple[tuple[str, str], ...] = (
    (DS_TITLE, "Общая сумма"),
    (DS_TITLE, "Условия поставки"),
    (DS_TITLE, "Грузоотправитель"),
    (DS_TITLE, "Грузополучатель"),
    (DS_TITLE, "Изготовитель"),
    (DS_TITLE, "Стороны согласовали"),
    (DS_TITLE, "Поставщик подтверждает"),
    (DS_TITLE, "Во всем остальном"),
    (DS_TITLE, "В соответствии с"),
    (DS_TITLE, "платёж"),
    (DS_TITLE, "платеж"),
    (DS_TITLE, "со стороны Поставщика"),
    (DS_TITLE, "со стороны Покупателя"),
    (DS_UNIT_PRICE_EXCL_VAT, "к Договору"),
    (DS_UNIT_PRICE_EXCL_VAT, "(далее"),
    (DS_UNIT_PRICE_EXCL_VAT, "далее -"),
)


def _cell_text_raw(row: RowStd, key: str) -> str:
    """Cell text for substring checks without mutating VALUES via get_value float coercion."""
    el = row.el.get(key)
    if el is None:
        return ""
    v = el.value
    if v is None:
        return ""
    if isinstance(v, list):
        return "" if len(v) == 0 else str(v)
    return str(v)


def _ds_spec_table_head_row(row: RowStd) -> bool:
    """True for repeatable Excel column-title / party-label rows (not data positions)."""
    name = _cell_text_raw(row, NAME)
    if "Наименование Позиций Товара по РД" in name:
        return True
    if "Код 1С СОУ" in _cell_text_raw(row, DS_CODE_1C):
        return True
    if "Технические требования (ГОСТ/ ТУ и др.)" in _cell_text_raw(row, TYPE_MARK):
        return True
    units = _cell_text_raw(row, UNITS)
    if "Ед. изм." in units:
        return True
    u_st = units.strip()
    if u_st in ("ПОКУПАТЕЛЬ", "ПОСТАВЩИК", "Поставщик", "Покупатель"):
        return True
    title = _cell_text_raw(row, DS_TITLE).strip()
    code = _cell_text_raw(row, CODE)
    system = _cell_text_raw(row, DS_SYSTEM).strip()
    if title == "Титул" and ("Код РД" in code or str(code).strip() == "Код РД"):
        return True
    if title == "Титул" and system == "Раздел" and "Код РД" in code:
        return True
    pup = _cell_text_raw(row, DS_UNIT_PRICE_EXCL_VAT)
    if "Цена за Позицию Товара" in pup and "без НДС" in pup:
        return True
    if "Срок поставки" in _cell_text_raw(row, DS_DELIVERY_TIME) and "Прим." in _cell_text_raw(
        row, VENDOR
    ):
        return True
    # Мини-шапка «Кол-во» / «Код РД» только как подпись столбца (не числовая позиция).
    vals = _cell_text_raw(row, VALUES).strip()
    if vals == "Кол-во" and (title == "Титул" or "Код РД" in str(code)):
        return True
    return False


def _ds_spec_boilerplate_row(row: RowStd) -> bool:
    """True for known contract/footer phrases (not table column headers)."""
    for col, needle in _DS_SPEC_BOILERPLATE_PAIRS:
        if needle in _cell_text_raw(row, col):
            return True
    return False


def _row_all_keys_blank(row: RowStd, keys: tuple[str, ...] | list[str]) -> bool:
    """True if every listed column is blank (same emptiness test as mandatory-field loop)."""
    for key in keys:
        value = str(row.get_value(key)).strip()
        if value != "":
            return False
    return True


def get_row_type(row: RowStd) -> str:
    """Classify a DS specification row as empty, head, boilerplate, or position.

    Order: fully blank → empty_row; table column titles → head_row; contract/footer
    substrings → other_row; mandatory att==1 gap → other_row; check_row → position_row;
    else other_row.

    Args:
        row: Loaded DS specification row.

    Returns:
        One of ``RowType.empty_row``, ``RowType.head_row``, ``RowType.other_row``,
        ``RowType.position_row``.
    """
    column_dict = {
        NUMBERS : 3,  # 1 - Номер
        DS_TITLE : 1, # 6550 - Титул
        DS_SYSTEM : 1, # SOT - Система
        DS_SPECIFICATION : 3,  # Спецификация
        TAGS : 3, # 2265-SH-02-S-UZ-0711
        DS_CODE_1C : 3, # 002054119
        CODE : 1, # BCC0000664 — код РД (кол. «Код РД»)
        NAME : 1, # Наименование МТР
        TYPE_MARK : 3, # Elsys-SWPS-2И
        VALUES : 1, # 1
        UNITS : 1, # шт
    }
    all_keys = tuple(column_dict.keys())

    if _row_all_keys_blank(row, all_keys):
        return RowType.empty_row

    if _ds_spec_table_head_row(row):
        return RowType.head_row

    if _ds_spec_boilerplate_row(row):
        return RowType.other_row

    flag = 0
    for key, att in column_dict.items():
        value = str(row.get_value(key)).strip()
        if value == "" and att == 1:
            flag = 1
    if flag == 1:
        return RowType.other_row

    if RowType.check_row(row, column_dict):
        return RowType.position_row

    return RowType.other_row
