from base.base_classes import *
from base.tables_columns import (
    UNITS, CODE, TAGS, VALUES, VENDOR, MASS, NUMBERS, TYPE_MARK, NAME,
    DS_NUMBER, DS_TITLE, DS_SYSTEM, DS_SPECIFICATION, DS_RFQ, DS_CODE_1C, DS_NAME_BY_RFQ,
    DS_UNIT_PRICE_EXCL_VAT, DS_PACKAGING_PRICE_EXCL_VAT, DS_SHIPPING_PRICE_EXCL_VAT,
    DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT, DS_TOTAL_PRICE_EXCL_VAT, DS_VAT_AMOUNT,
    DS_TOTAL_PRICE_INCL_VAT, DS_DELIVERY_TIME,
    ANNOTATION_2, ANNOTATION_3, IN_CABINET,
    TAGS_2, NUMBERS_2, NAME_2, TYPE_MARK_2, CODE_2, VENDOR_2, UNITS_2, VALUES_2,
)


def get_row_type(row: RowStd):
    # Проверяем на пустые ячейки
    column_dict = {
        DS_NUMBER: 3,
        DS_TITLE: 1,
        DS_SYSTEM: 1,
        DS_SPECIFICATION: 3,
        DS_RFQ: 3,
        NAME: 3,
        DS_CODE_1C: 3,
        CODE: 3,
        DS_NAME_BY_RFQ: 3,
        TYPE_MARK: 3,
        UNITS: 3,
        VALUES: 3,
        DS_UNIT_PRICE_EXCL_VAT: 3,
        DS_PACKAGING_PRICE_EXCL_VAT: 3,
        DS_SHIPPING_PRICE_EXCL_VAT: 3,
        DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT: 3,
        DS_TOTAL_PRICE_EXCL_VAT: 3,
        DS_VAT_AMOUNT: 3,
        DS_TOTAL_PRICE_INCL_VAT: 3,
        DS_DELIVERY_TIME: 3,
        VENDOR: 1,
        ANNOTATION_2: 3,
        ANNOTATION_3: 3,
        IN_CABINET: 3,
        TAGS_2: 3,
        NUMBERS_2: 3,
        NAME_2: 3,
        TYPE_MARK_2: 3,
        CODE_2: 3,
        VENDOR_2: 3,
        UNITS_2: 3,
        VALUES_2: 3,
    }
    #
    if ("Наименование Позиций Товара по РД" in row.el[NAME].value
            or "Код 1С СОУ" in row.el[DS_CODE_1C].value
            or "Технические требования (ГОСТ/ ТУ и др.)" in row.el[TYPE_MARK].value
            or "Ед. изм." in row.el[UNITS].value):
        return RowType.other_row
    # POSITION DICT
    if RowType.check_row(row, column_dict):
        return RowType.position_row

    return RowType.other_row
