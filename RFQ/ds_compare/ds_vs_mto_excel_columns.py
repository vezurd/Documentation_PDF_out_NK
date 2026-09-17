"""
Column definitions for the fast DS vs MTO xlsx export (see ds_vs_mto_excel_xlsxwriter).

Same idea as step4_6_save_match_result_to_excel.OUTPUT_COLUMNS_CONFIG: toggle
columns (output), widths, and Excel outline (group_level / group_collapsed).

For the ``_xw`` file, header text, header background colour, and column width are
read from ``templates/ds_vs_mto_template.xlsx`` when present; values in this
module (``header_label``, ``width``, ``section``) are then fallbacks or used for
outline/section styling if the template is missing.

Order must match the export order; keep one entry per DsVsMto column.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Any, List, Literal

from base.tables_columns import (
    ANNOTATION_2,
    ANNOTATION_3,
    ANNOTATION_MTO,
    CODE,
    CODE_2,
    ColNames,
    DS_CODE_1C,
    DS_DELIVERY_TIME,
    DS_NAME,
    DS_NAME_BY_RFQ,
    DS_NUMBER,
    DS_PACKAGING_PRICE_EXCL_VAT,
    DS_RFQ,
    DS_SHIPPING_PRICE_EXCL_VAT,
    DS_SPECIFICATION,
    DS_SYSTEM,
    DS_TITLE,
    DS_TOTAL_PRICE_EXCL_VAT,
    DS_TOTAL_PRICE_INCL_VAT,
    DS_UNIT_PRICE_EXCL_VAT,
    DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT,
    DS_VAT_AMOUNT,
    IN_CABINET,
    MTO_DS_DIFF,
    MTO_DS_RFQ_DIFF,
    NAME,
    NAME_2,
    NUMBERS_2,
    DELIVERY_OVERALL_STATUS,
    RFQ_CODE,
    RFQ_COMPARE_STATUS,
    RFQ_NAME,
    RFQ_TYPE_MARK,
    RFQ_UNITS,
    RFQ_VALUES,
    TAGS_2,
    TYPE_MARK,
    TYPE_MARK_2,
    UL_CODE,
    UL_COMPARE_STATUS,
    UL_DATA_STATUS,
    UL_NAME,
    UL_ORDERED_VALUES,
    UL_REMAINING_VALUES,
    UL_SOURCE_FILES,
    UL_TAGS,
    UL_TYPE_MARK,
    UL_UNITS,
    UL_VALUES,
    UL_VENDOR,
    UNITS,
    UNITS_2,
    VALUES,
    VALUES_2,
    VENDOR,
    VENDOR_2,
)

# Logical blocks for header row fill color in xlsxwriter
Section = Literal["ds", "prices", "meta", "mto", "rfq", "ul"]

DEFAULT_COLUMN_WIDTH = 15


@dataclass
class ColumnDef:
    """One DS vs MTO output column (same role as in step4_6)."""

    col_name: str
    header_label: str
    section: Section
    output: bool = True
    width: int | None = None
    group_level: int | None = None
    group_collapsed: bool = False
    header_color: str | None = None


# Header labels: single-line (fallback if template missing); template overrides in _xw.
# output=True: column is written unless you set output=False.
DS_VS_MTO_OUTPUT_COLUMNS_CONFIG: List[ColumnDef] = [
    ColumnDef(DS_NAME, "Имя ДС", "ds",                                          output=True, width=7),
    ColumnDef(DS_NUMBER, "№ п/п", "ds",                                         output=True, width=7),  
    ColumnDef(DS_TITLE, "Титул", "ds",                                          output=True, width=9),
    ColumnDef(DS_SYSTEM, "Раздел", "ds",                                        output=True, width=9),
    ColumnDef(DS_SPECIFICATION, "Линия/ TAG-Номер/ Спецификация", "ds",             output=False, width=20),
    ColumnDef(DS_RFQ, "RFQ", "ds",                                                  output=False, width=10),
    ColumnDef(NAME, "Наименование Позиций Товара по РД", "ds",                  output=True, width=28),

    ColumnDef(DS_CODE_1C, "Код 1С СОУ", "ds",                                       output=False, width=12),

    ColumnDef(CODE, "Код РД", "ds",                                             output=True, width=14),

    ColumnDef(DS_NAME_BY_RFQ, "Наименование Позиций Товара Поставщика", "ds",       output=False, width=28),

    ColumnDef(TYPE_MARK, "Технические требования (ГОСТ/ ТУ и др.)", "ds",       output=True, width=22),
    ColumnDef(UNITS, "Ед. изм.", "ds",                                          output=True, width=8),
    ColumnDef(VALUES, "Кол-во", "ds",                                           output=True, width=7),
    ColumnDef(
        DS_UNIT_PRICE_EXCL_VAT,
        "Цена за Позицию Товара (RUB) без НДС (ЕР по Прил. 1.1. к Договору)",
        "prices",
        output=False,
        width=18,
    ),

    ColumnDef(DS_PACKAGING_PRICE_EXCL_VAT, "Цена (RUB) за упаковку без НДС", "prices", output=False, width=12),
    ColumnDef(DS_SHIPPING_PRICE_EXCL_VAT, "Цена (RUB) за транспортировку без НДС", "prices", output=False, width=12),
    ColumnDef(DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT, "Цена за Позицию Товара с учетом упаковки и транспортировки (RUB) без НДС", "prices", output=False, width=20),
    ColumnDef(DS_TOTAL_PRICE_EXCL_VAT, "Цена за Позиции Товара с учетом упаковки и транспортировки (RUB) без НДС", "prices", output=False, width=20),

    ColumnDef(DS_VAT_AMOUNT, "(RUB) НДС - 20%", "prices", output=False, width=12),
    ColumnDef(DS_TOTAL_PRICE_INCL_VAT, "Цена Позиций Товара (RUB) с НДС 20%", "prices", output=False, width=18),
    ColumnDef(DS_DELIVERY_TIME, "Срок поставки", "prices", output=False, width=12),
    ColumnDef(VENDOR, "Прим.", "meta", output=False, width=10),

    ColumnDef(ANNOTATION_2, "комментарий", "meta",                                      output=True, width=23),
    ColumnDef(ANNOTATION_3, "замена кода", "meta",                                      output=True, width=16),
    ColumnDef(IN_CABINET, "Оборудование относящиеся к шкафам (по тегу шкафа)", "meta",  output=True, width=30),

    ColumnDef(TAGS_2, "TAG №", "mto", output=True, width=20),
    ColumnDef(NUMBERS_2, "п.п.№ спецификации МТО", "mto", output=True, width=20),
    ColumnDef(ANNOTATION_MTO, "ANNOTATION (МТО)", "mto", output=True, width=24),
    ColumnDef(NAME_2, "Наименование и техническая характеристика", "mto", output=True, width=30),
    ColumnDef(TYPE_MARK_2, "Тип, марка, обозначение документа, опросного листа", "mto", output=True, width=18),
    ColumnDef(CODE_2, "Код продукции", "mto", output=True, width=12),
    ColumnDef(VENDOR_2, "Поставщик", "mto", output=True, width=16),
    ColumnDef(UNITS_2, "Ед. измерения", "mto", output=True, width=8),
    ColumnDef(VALUES_2, "Кол.", "mto", output=True, width=8),
    ColumnDef(MTO_DS_DIFF, "Спека − ДС", "rfq", output=True, width=10),
    ColumnDef(RFQ_VALUES, "Кол-во RFQ", "rfq", output=True, width=10),
    ColumnDef(RFQ_UNITS, "Ед. изм. RFQ", "rfq", output=True, width=10),
    ColumnDef(MTO_DS_RFQ_DIFF, "Спека − ДС − RFQ", "rfq", output=True, width=12),
    ColumnDef(RFQ_COMPARE_STATUS, "Статус RFQ", "rfq", output=True, width=16),
    ColumnDef(
        DELIVERY_OVERALL_STATUS,
        "Общий статус поставки",
        "rfq",
        output=True,
        width=18,
    ),
    ColumnDef(RFQ_NAME, "Наименование МТР", "rfq", output=True, width=28),
    ColumnDef(RFQ_CODE, "Код РД / BCC", "rfq", output=True, width=14),
    ColumnDef(RFQ_TYPE_MARK, "Техн. характеристики RFQ", "rfq", output=True, width=22),
    ColumnDef(UL_ORDERED_VALUES, "Заказано (ДС + RFQ)", "ul", output=True, width=13),
    ColumnDef(UL_VALUES, "Кол-во по УЛ", "ul", output=True, width=11),
    ColumnDef(UL_UNITS, "Ед. изм. УЛ", "ul", output=True, width=10),
    ColumnDef(UL_REMAINING_VALUES, "Остаток поставки", "ul", output=True, width=13),
    ColumnDef(UL_COMPARE_STATUS, "Статус УЛ", "ul", output=True, width=20),
    ColumnDef(UL_DATA_STATUS, "Качество данных УЛ", "ul", output=True, width=22),
    ColumnDef(UL_CODE, "Код УЛ", "ul", output=True, width=14),
    ColumnDef(
        UL_SOURCE_FILES,
        "Исходники УЛ (файл · вкладка · строка)",
        "ul",
        output=True,
        width=45,
    ),
    ColumnDef(UL_NAME, "Наименование УЛ", "ul", output=False, width=28),
    ColumnDef(UL_TYPE_MARK, "Техн. характеристики УЛ", "ul", output=False, width=22),
    ColumnDef(UL_VENDOR, "Поставщик УЛ", "ul", output=False, width=18),
    ColumnDef(UL_TAGS, "Теги УЛ", "ul", output=False, width=24),
]

HEADER_FILL_COLORS: dict[str, str] = {
    "ds": "#dbbcdb",
    "prices": "#c6e0b4",
    "meta": "#ffcc00",
    "mto": "#cd7f32",
    "rfq": "#b7dee8",
    "ul": "#c9daf8",
}

SECTION_LABELS: dict[str, str] = {
    "ds": "ДС",
    "prices": "Цены",
    "meta": "Мета",
    "mto": "МТО",
    "rfq": "RFQ",
    "ul": "УЛ",
}

_MIN_COLUMN_WIDTH = 5
_MAX_COLUMN_WIDTH = 120
MIN_COLUMN_WIDTH = _MIN_COLUMN_WIDTH
MAX_COLUMN_WIDTH = _MAX_COLUMN_WIDTH


def default_header_color(section: Section) -> str:
    """Default header fill for a logical section."""
    return HEADER_FILL_COLORS.get(section, HEADER_FILL_COLORS["ds"])


def normalize_header_color(value: object, *, fallback: str) -> str:
    """Normalize ``#RRGGBB`` header color; invalid values fall back."""
    text = str(value or "").strip()
    if not text:
        return fallback
    if not text.startswith("#"):
        text = f"#{text}"
    if len(text) != 7:
        return fallback
    try:
        int(text[1:], 16)
    except ValueError:
        return fallback
    return text.lower()


def column_def_to_settings_dict(col: ColumnDef) -> dict[str, Any]:
    """Serialize one ``ColumnDef`` for ``ds_compare_config.json``."""
    return {
        "col_name": col.col_name,
        "output": bool(col.output),
        "width": int(col.width or DEFAULT_COLUMN_WIDTH),
        "header_label": col.header_label,
        "header_color": col.header_color or default_header_color(col.section),
        "section": col.section,
    }


def default_columns_settings() -> list[dict[str, Any]]:
    """Default column settings list (same order as export)."""
    return [column_def_to_settings_dict(c) for c in DS_VS_MTO_OUTPUT_COLUMNS_CONFIG]


def merge_column_settings(saved: list[dict[str, Any]] | None) -> list[dict[str, Any]]:
    """Merge saved JSON overrides onto built-in ``ColumnDef`` defaults."""
    by_name: dict[str, dict[str, Any]] = {}
    if saved:
        for item in saved:
            if not isinstance(item, dict):
                continue
            name = str(item.get("col_name", "")).strip()
            if name:
                by_name[name] = item

    merged: list[dict[str, Any]] = []
    for base in DS_VS_MTO_OUTPUT_COLUMNS_CONFIG:
        row = column_def_to_settings_dict(base)
        override = by_name.get(base.col_name)
        if not override:
            merged.append(row)
            continue
        row["output"] = bool(override.get("output", row["output"]))
        try:
            width = int(override.get("width", row["width"]))
        except (TypeError, ValueError):
            width = row["width"]
        row["width"] = max(_MIN_COLUMN_WIDTH, min(_MAX_COLUMN_WIDTH, width))
        label = str(override.get("header_label", row["header_label"])).strip()
        row["header_label"] = label or row["header_label"]
        row["header_color"] = normalize_header_color(
            override.get("header_color"),
            fallback=row["header_color"],
        )
        merged.append(row)
    return merged


def column_settings_to_defs(settings: list[dict[str, Any]] | None) -> list[ColumnDef]:
    """Build export ``ColumnDef`` list from normalized settings."""
    merged = merge_column_settings(settings)
    base_by_name = {c.col_name: c for c in DS_VS_MTO_OUTPUT_COLUMNS_CONFIG}
    out: list[ColumnDef] = []
    for item in merged:
        base = base_by_name.get(item["col_name"])
        if base is None:
            continue
        out.append(
            replace(
                base,
                output=bool(item["output"]),
                width=int(item["width"]),
                header_label=str(item["header_label"]),
                header_color=normalize_header_color(
                    item.get("header_color"),
                    fallback=default_header_color(base.section),
                ),
            )
        )
    return out


def column_config_from_output_settings(output_cfg: dict[str, Any] | None) -> list[ColumnDef]:
    """Resolve export columns from ``ds_vs_mto_output`` config block."""
    columns = None
    if isinstance(output_cfg, dict):
        raw_columns = output_cfg.get("columns")
        if isinstance(raw_columns, list):
            columns = raw_columns
    return column_settings_to_defs(columns)


def _assert_config_covers_dsvsmto() -> None:
    expected = {v for v in ColNames.DsVsMto.column_dict.values()}
    have = {c.col_name for c in DS_VS_MTO_OUTPUT_COLUMNS_CONFIG}
    if expected != have:
        raise ValueError(
            f"ds_vs_mto_excel_columns: set mismatch {expected ^ have!r}. "
            "Add/remove ColumnDef to match ColNames.DsVsMto."
        )


_assert_config_covers_dsvsmto()


def visible_columns(
    config: List[ColumnDef] | None = None,
) -> List[ColumnDef]:
    """Returns only columns with output=True, in list order."""
    src = config if config is not None else DS_VS_MTO_OUTPUT_COLUMNS_CONFIG
    return [c for c in src if c.output]
