"""
Этап 4_6: Сохранение результата сопоставления в Excel (xlsxwriter).

Стили применяются при write() — без отдельного цикла стилизации.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import math
import os
import re
import shutil
import tempfile
import time

import xlsxwriter

from typing import List, Literal

from base.base_classes import RowStd, RowType
from base.tables_columns import (
    TAGS, CODE, NAME, TYPE_MARK, VENDOR, NUMBERS, VALUES, VALUES_2, UNITS,
    TAG_MTO, CODE_MTO, MTO_CODE_STRUCK, NAME_MTO, TYPE_MARK_MTO, NUMBERS_MTO, VALUES_MTO,
    UNITS_MTO, UNITS_CHECK_STATUS, UNITS_CONVERSION_TRACE, VALUES_MTO_RFP_DIFF,
    VALUES_MTO_UL_DIFF,
    TAG_VO, CODE_VO, NAME_VO, VALUES_VO, MATCH_STATUS_VO,
    MATCH_STATUS_TAGS,
    MATCH_STATUS, HAS_TAGS_RFP, POSITION_STATUS, RFP_SUPPLY_STATUS, DS_SPECIFICATION, DS_CODE_1C, NAME_2,
    DS_NAME, DS_ACTUAL, DS_LOT, DS_NUMBER, DS_MANAGER, DS_TITLE,
    PATH_RFP, PATH_MTO,
    MATCH_STATUS_CODE_REPLACEMENT, EQUIPMENT_TYPE_STATUS,
    MATCH_STATUS_VO_MTO_TAGS, MATCH_STATUS_RFP_MTO_TAGS, MATCH_STATUS_RFP_MTO_STRUCK_CODE,
    TAG_EFFECTIVE, MTO_CABINET_EQUIPMENT, IN_CABINET,
    UL_ORDERED_VALUES, UL_VALUES, UL_UNITS, UL_REMAINING_VALUES,
    UL_COMPARE_STATUS, UL_DATA_STATUS, UL_TAG_MATCH_STATUS, UL_CODE, UL_SOURCE_FILES,
    UL_NAME, UL_TYPE_MARK, UL_VENDOR, UL_TAGS,
)
from RFQ.tags_rfp_compare.rfp_tags_utils import ensure_result_dir_exists, append_timing_log
from RFQ.tags_rfp_compare.step4.step4_6_quality_sheets import write_quality_sheets
from RFQ.tags_rfp_compare.step4.step4_packing_compare import (
    FALLBACK_COMMENT_MARKER,
    RfpPackingAudit,
    STATUS_DATA_PROBLEM,
    STATUS_MTO_DELIVERED_SHORT,
    STATUS_SHORTFALL,
    STATUS_VO_DELIVERED_SHORT,
)
from utils.colors import Color


Section = Literal["pre_status", "rfp", "status", "mto", "vo", "ul", "paths"]


@dataclass
class ColumnDef:
    """Описание столбца для вывода в Excel."""
    col_name: str
    header_label: str
    section: Section
    output: bool = True
    width: int | None = None
    group_level: int | None = None  # уровень outline Excel (1–7); None = не в группе
    # При group_level>=1: True = столбцы группы скрыты при открытии.
    # При group_level=None: True = столбец скрыт при открытии без outline (разрывает смежные группы level=1).
    group_collapsed: bool = False


OUTPUT_COLUMNS_CONFIG: List[ColumnDef] = [
    ColumnDef(MATCH_STATUS_VO_MTO_TAGS,         "VO vs MTO, Теги",              "pre_status", width=22, output=False),
    ColumnDef(MATCH_STATUS_RFP_MTO_TAGS,        "RFP vs MTO, Теги",             "pre_status", width=22, output=False),
    # rfp — первая группа (скрыта при открытии), затем разрыв у «Титул/Марка», затем группа TAGS…VALUES
    ColumnDef(DS_NAME,                          "Порядковый ДС",                "rfp", width=12, output=True, group_level=1, group_collapsed=True),
    ColumnDef(DS_ACTUAL,                        "Фактический ДС",               "rfp", width=12, output=True, group_level=1, group_collapsed=True),
    ColumnDef(DS_LOT,                           "Лот",                          "rfp", width=12, output=False, group_level=1, group_collapsed=True),
    ColumnDef(DS_NUMBER,                        "№ позиции",                    "rfp", width=7,  output=True, group_level=1, group_collapsed=True),
    ColumnDef(DS_MANAGER,                       "Менеджер ДС",                  "rfp", width=16, output=True, group_level=None),
    # Основной столбец для фильтрации по титулу/марке (без group_level = разрыв outline)
    ColumnDef(DS_TITLE,                         "Титул/Марка",                  "rfp", width=15),
    # output=False не на листе — group_level/group_collapsed не задаём (игнорируются в outline).
    ColumnDef(DS_SPECIFICATION,                 "Спецификация",                 "rfp", output=False),
    # Одна группа Excel: смежные на листе столбцы с одинаковым group_level=1 сворачиваются одним «−».
    # group_collapsed=True — открыть файл со свёрнутой группой.
    ColumnDef(DS_CODE_1C,                       "Код 1С",                       "rfp", output=False),
    ColumnDef(CODE,                             "Код RFP",                      "rfp", width=15, group_level=1, group_collapsed=False),
    ColumnDef(NAME,                             "Наименование RFP",             "rfp", width=25, group_level=1, group_collapsed=False),
    ColumnDef(TYPE_MARK,                        "Тип марки RFP",                "rfp", width=20, group_level=1, group_collapsed=False),
    ColumnDef(VENDOR,                           "Поставщик RFP",                "rfp", width=18, group_level=1, group_collapsed=False),
    ColumnDef(VALUES,                           "Кол-во RFP",                   "rfp", width=7, group_level=0, group_collapsed=False),
    ColumnDef(UNITS,                            "Ед. изм. RFP",                 "rfp", width=10, group_level=0, group_collapsed=False),
    ColumnDef(TAGS,                             "Теги RFP",                     "rfp", width=25, group_level=0, group_collapsed=False),
    ColumnDef(NAME_2,                           "Наименование 2",               "rfp", output=False),
    ColumnDef(VALUES_2,                         "Закупка по Лоту, кол-во",      "rfp", output=False),
    # status — TAG_EFFECTIVE/POSITION: только скрыть при открытии (level=0), иначе слипнутся с RFP-группой TAGS…VALUES
    ColumnDef(HAS_TAGS_RFP,                     "RFP, строки с тегами",         "status", output=False, width=15, group_level=None),
    ColumnDef(TAG_EFFECTIVE,                    "Тег (VO/MTO/RFP)",             "status", output=True, width=23,  group_level=1, group_collapsed=True),
    ColumnDef(POSITION_STATUS,                  "MTO, Статус позиции",          "status", output=True, width=13, group_level=1, group_collapsed=True),
    ColumnDef(RFP_SUPPLY_STATUS,                "Статус поставки",              "status", output=True, width=22, group_level=1, group_collapsed=True),
    ColumnDef(MATCH_STATUS,                     "MTO, Сравнение Тегов",         "status", output=False, width=22, group_level=None),
    ColumnDef(MATCH_STATUS_VO,                  "VO, Сравнение Тегов",          "status", output=False, width=22, group_level=None),
    ColumnDef(MATCH_STATUS_TAGS,                "Итог, Сравнение Тегов",        "status", output=False, width=22, group_level=None),
    ColumnDef(MATCH_STATUS_CODE_REPLACEMENT,    "Замены Кодов",                 "status", output=False, width=22,  group_level=None),
    ColumnDef(EQUIPMENT_TYPE_STATUS,            "Тип оборудования",             "status", output=True, width=10, group_level=None),
    ColumnDef(IN_CABINET,                       "Тег Шкафа",                    "status", output=True, width=25, group_level=0),
    ColumnDef(MTO_CABINET_EQUIPMENT,            "MTO, Шкафное оборудование",    "status", output=False, width=22, group_level=None),
    ColumnDef(MATCH_STATUS_RFP_MTO_STRUCK_CODE, "RFP vs MTO, Код (зачеркнутый)","status", output=False, width=22, group_level=None),
    # mto 
    ColumnDef(TAG_MTO,                          "Тег MTO",                      "mto", width=25, group_level=None),
    ColumnDef(CODE_MTO,                         "Код MTO",                      "mto", width=15, group_level=None),
    ColumnDef(NAME_MTO,                         "Наименование MTO",             "mto", width=30, group_level=None),
    ColumnDef(TYPE_MARK_MTO,                    "Тип марки MTO",                "mto", width=20, group_level=None),
    ColumnDef(VALUES_MTO,                       "Количество MTO",               "mto", width=7,  group_level=None),
    ColumnDef(UNITS_MTO,                        "Ед. изм. MTO",                 "mto", width=10, group_level=None),
    ColumnDef(NUMBERS_MTO,                      "Номера MTO",                   "mto", width=8,  group_level=None),
    ColumnDef(UNITS_CHECK_STATUS,               "Статус ед. изм.",              "mto", width=18, output=False),
    ColumnDef(UNITS_CONVERSION_TRACE,           "Trace ед. изм.",               "mto", width=40, output=False),
    # vo 
    ColumnDef(TAG_VO,                           "Тег РКД",                      "vo", width=25, group_level=None),
    ColumnDef(CODE_VO,                          "Код РКД",                      "vo", width=15, group_level=None),
    ColumnDef(NAME_VO,                          "Наименование РКД",             "vo", width=30, group_level=None),
    ColumnDef(VALUES_VO,                        "Количество в РКД",             "vo", width=8,  group_level=None),
    # ul
    ColumnDef(UL_ORDERED_VALUES,                "Заказано по RFP",              "ul", width=13, output=False),
    ColumnDef(VALUES_MTO_RFP_DIFF,             "МТО − RFP",                    "mto", width=10, group_level=None),
    ColumnDef(VALUES_MTO_UL_DIFF,              "МТО − УЛ",                     "mto", width=10, group_level=None),
    ColumnDef(UL_VALUES,                        "Кол-во по УЛ",                 "ul", width=11),
    ColumnDef(UL_UNITS,                         "Ед. изм. УЛ",                  "ul", width=10),
    ColumnDef(UL_REMAINING_VALUES,              "RFP − УЛ",                     "ul", width=13),
    ColumnDef(UL_COMPARE_STATUS,                "Статус УЛ",                    "ul", width=32),
    ColumnDef(UL_DATA_STATUS,                   "Качество данных УЛ",            "ul", width=22, output=False),
    ColumnDef(UL_TAG_MATCH_STATUS,              "Статус тегов УЛ",              "ul", width=36),
    ColumnDef(UL_TAGS,                          "Тег УЛ",                       "ul", width=25),
    ColumnDef(UL_NAME,                          "Наименование УЛ",              "ul", width=28),
    ColumnDef(UL_CODE,                          "Код УЛ",                       "ul", width=14),
    ColumnDef(UL_SOURCE_FILES,                  "Источник УЛ (файл · вкладка · строка)", "ul", width=58),
    ColumnDef(UL_TYPE_MARK,                     "Тип марки УЛ",                 "ul", width=22, group_level=1, group_collapsed=True),
    ColumnDef(UL_VENDOR,                        "Поставщик УЛ",                 "ul", width=18, group_level=1, group_collapsed=True),
    ColumnDef(PATH_RFP,                         "Путь к RFP",                   "paths", width=55, group_level=None),
    ColumnDef(PATH_MTO,                         "Путь к МТО",                   "paths", width=55, group_level=None),
]

HEADER_FILL_COLORS: dict[str, str] = {
    "pre_status": "#ffcc00",
    "rfp": "#dbbcdb",
    "status": "#ffcc00",
    "mto": "#cd7f32",
    "vo": "#a7c7e7",
    "ul": "#C6EFCE",
    "paths": "#D9E2F3",
}

DEFAULT_COLUMN_WIDTH = 15
NUMERIC_QUANTITY_COLUMNS = frozenset((
    VALUES,
    VALUES_MTO,
    VALUES_MTO_RFP_DIFF,
    VALUES_MTO_UL_DIFF,
    UL_ORDERED_VALUES,
    UL_VALUES,
    UL_REMAINING_VALUES,
))

_UL_SOURCE_SEP = " · "
_UL_ROW_RE = re.compile(r"^строка\s+(\d+)$", re.IGNORECASE)
_HYPERLINK_FONT = "#0563C1"


def _is_absolute_file_path(path: str) -> bool:
    text = str(path or "").strip()
    if not text:
        return False
    if text.startswith("\\\\") or text.startswith("//"):
        return True
    return len(text) >= 3 and text[1] == ":" and text[0].isalpha()


def resolve_source_file_path(path: str, root: str = "") -> str:
    """Return an absolute/UNC path, joining *root* when *path* is relative."""
    text = str(path or "").strip()
    if not text:
        return ""
    if _is_absolute_file_path(text):
        return os.path.normpath(text)
    root_text = str(root or "").strip()
    if not root_text:
        return ""
    return os.path.normpath(os.path.join(root_text, text))


def parse_single_ul_source(text: str) -> tuple[str, str, int | None] | None:
    """Parse one ``file · sheet · строка N`` line; ``None`` if several sources."""
    raw = str(text or "").strip()
    if not raw or "\n" in raw:
        return None
    parts = [part.strip() for part in raw.split(_UL_SOURCE_SEP) if part.strip()]
    if not parts:
        return None
    file_part = parts[0]
    rest = parts[1:]
    excel_row: int | None = None
    if rest:
        row_match = _UL_ROW_RE.match(rest[-1])
        if row_match:
            excel_row = int(row_match.group(1))
            rest = rest[:-1]
    sheet = rest[0] if rest else ""
    return file_part, sheet, excel_row


def excel_external_file_url(
    path: str,
    *,
    sheet: str = "",
    excel_row: int | None = None,
) -> str:
    """Build an xlsxwriter ``external:`` URL, optionally to a sheet cell."""
    url = f"external:{os.path.normpath(path)}"
    if not sheet:
        return url
    cell = f"A{excel_row}" if isinstance(excel_row, int) and excel_row > 0 else "A1"
    escaped = str(sheet).replace("'", "''")
    return f"{url}#'{escaped}'!{cell}"


def excel_path_link(
    col_name: str,
    value: str,
    *,
    packing_root: str = "",
) -> tuple[str, str | None]:
    """Return ``(display, url_or_none)`` for the UL source cell only.

    Excel caps a workbook at ~65 530 hyperlinks. PATH_RFP / PATH_MTO stay
    plain full paths. UL keeps a link only when the cell has one source.
    No hyperlink tooltip.

    Args:
        col_name: Only ``UL_SOURCE_FILES`` produces a URL.
        value: Cell text already stored on the row.
        packing_root: TSD folder used to resolve relative UL files.

    Returns:
        Display string and optional ``external:`` URL.
    """
    text = str(value or "")
    if col_name != UL_SOURCE_FILES:
        return text, None
    parsed = parse_single_ul_source(text)
    if parsed is None:
        return text, None
    file_part, sheet, excel_row = parsed
    resolved = resolve_source_file_path(file_part, packing_root)
    if not resolved:
        return text, None
    return text, excel_external_file_url(
        resolved, sheet=sheet, excel_row=excel_row
    )


_UL_EXCEL_COMMENT_STATUSES = frozenset({
    STATUS_SHORTFALL,
    STATUS_DATA_PROBLEM,
    STATUS_MTO_DELIVERED_SHORT,
    STATUS_VO_DELIVERED_SHORT,
})


def should_write_excel_comment(col_name: str, row: RowStd, comment: object) -> bool:
    """Return True if *comment* should become an Excel cell comment.

    Allocation traces stay on ``RowStd`` for quality sheets. Excel clouds on
    ``UL_COMPARE_STATUS`` are limited to shortfall (including leftover MTO/VO
    short), packing data problems, and fallback matches. Full leftover
    delivery statuses do not get a cloud unless the comment has the fallback
    marker. Other columns (for example ``UNITS_MTO``) are unchanged.

    Args:
        col_name: Output column key.
        row: Row being written.
        comment: ``CheckElement.comment`` value.

    Returns:
        Whether xlsxwriter should call ``write_comment``.
    """
    text = str(comment or "").strip()
    if not text:
        return False
    if col_name != UL_COMPARE_STATUS:
        return True
    status = str(row.get_value(UL_COMPARE_STATUS) or "").strip()
    if status in _UL_EXCEL_COMMENT_STATUSES:
        return True
    return FALLBACK_COMMENT_MARKER in text


def _quantity_value_for_excel(value):
    """Return a finite number when a UL quantity is numeric."""
    if value is None or value == "" or isinstance(value, bool):
        return value
    try:
        numeric_value = float(str(value).strip().replace(",", "."))
    except (TypeError, ValueError):
        return value
    return numeric_value if math.isfinite(numeric_value) else value


def _columndef_for_grouping(col_name: str) -> ColumnDef | None:
    """ColumnDef для outline: только output=True.

    Строки с output=False не попадают на лист — их group_level/group_collapsed
    не участвуют в группировке (нет «фантомных» столбцов между видимыми).
    """
    for d in OUTPUT_COLUMNS_CONFIG:
        if d.col_name == col_name and d.output:
            return d
    return None


def _path_is_unc(path: str) -> bool:
    normalized = os.path.normpath(path)
    return normalized.startswith("\\\\") or normalized.startswith("//")


def save_match_result_to_excel(
    result_rows: List[RowStd],
    result_dir: str,
    debug_log: List[str] | None = None,
    packing_audit: RfpPackingAudit | None = None,
    column_config: list[ColumnDef] | None = None,
) -> str:
    """Save the RFP/MTO match result to an xlsx workbook.

    Args:
        result_rows: Matched rows to export.
        result_dir: Destination directory.
        debug_log: Optional list that receives timing lines.
        packing_audit: Optional packing-list audit for UL hyperlinks.
        column_config: Output columns; ``None`` uses the active Step4 template.

    Returns:
        Path of the published workbook.
    """
    tmp_path = None
    publish_ok = False
    try:
        ensure_result_dir_exists(result_dir)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"Шаг4_Сопоставление_RFP_MTO_{timestamp}.xlsx"
        file_path = os.path.join(result_dir, file_name)
        tmp_fd, tmp_path = tempfile.mkstemp(prefix="step4_", suffix=".xlsx")
        os.close(tmp_fd)

        wb = xlsxwriter.Workbook(
            tmp_path,
            {"constant_memory": False, "strings_to_urls": False},
        )
        ws = wb.add_worksheet("Сопоставление RFP и MTO")

        if column_config is None:
            from RFQ.tags_rfp_compare.step4.step4_excel_columns import (
                resolve_active_column_defs,
            )
            column_config = resolve_active_column_defs()
        from RFQ.tags_rfp_compare.step4.step4_excel_columns import (
            outline_options_for_visible,
        )
        visible_defs = [d for d in column_config if d.output]
        output_column_dict = {i: d.col_name for i, d in enumerate(visible_defs)}
        headers = [d.header_label for d in visible_defs]
        col_section = {i: d.section for i, d in enumerate(visible_defs)}
        col_width = {
            i: (d.width or DEFAULT_COLUMN_WIDTH) for i, d in enumerate(visible_defs)
        }
        num_cols = len(headers)

        # Pre-build formats
        header_formats: dict[str, xlsxwriter.workbook.Format] = {}
        for section, hex_color in HEADER_FILL_COLORS.items():
            header_formats[section] = wb.add_format({
                "bold": True,
                "font_color": "#000000",
                "bg_color": hex_color,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "text_wrap": True,
            })

        cell_base_fmt = wb.add_format({
            "align": "left",
            "valign": "vcenter",
            "border": 1,
        })
        cell_wrap_fmt = wb.add_format({
            "align": "left",
            "valign": "top",
            "border": 1,
            "text_wrap": True,
        })

        fill_format_cache: dict[tuple[str, bool, bool], xlsxwriter.workbook.Format] = {}

        def _get_cell_format(
            color_hex: str,
            *,
            wrap: bool = False,
            is_link: bool = False,
        ):
            cache_key = (color_hex, wrap, is_link)
            if cache_key not in fill_format_cache:
                props = {
                    "align": "left",
                    "valign": "top" if wrap else "vcenter",
                    "border": 1,
                }
                if color_hex:
                    props["bg_color"] = (
                        f"#{color_hex}" if not color_hex.startswith("#") else color_hex
                    )
                if wrap:
                    props["text_wrap"] = True
                if is_link:
                    props["font_color"] = _HYPERLINK_FONT
                    props["underline"] = 1
                fill_format_cache[cache_key] = wb.add_format(props)
            return fill_format_cache[cache_key]

        packing_root = ""
        if packing_audit is not None:
            meta = getattr(getattr(packing_audit, "dataset", None), "meta", None)
            packing_root = str(getattr(meta, "root", "") or "").strip()

        # Header row: one line per column; height 3× compact row (18)
        ws.set_row(0, 54)
        for ci, label in enumerate(headers):
            section = col_section.get(ci, "rfp")
            fmt = header_formats.get(section, header_formats["rfp"])
            ws.write(0, ci, label, fmt)

        # Write data rows with inline styling
        data_rows = [r for r in result_rows if r.row_type not in (RowType.empty_row, RowType.other_row)]
        total_rows = len(data_rows)
        print(
            f"Запись Excel: {total_rows} строк x {num_cols} столбцов -> {file_name}",
            flush=True,
        )

        t_write_start = time.perf_counter()
        sorted_col_keys = sorted(output_column_dict.keys())
        comment_write_errors: list[str] = []
        comment_cells = 0
        progress_every = 10000 if total_rows >= 20000 else 0

        for row_idx, row in enumerate(data_rows):
            excel_row = row_idx + 1
            if progress_every and excel_row % progress_every == 0:
                print(
                    f"  Excel: записано {excel_row}/{total_rows} строк "
                    f"({time.perf_counter() - t_write_start:.0f}s)",
                    flush=True,
                )
            source_line_count = 1
            for ci, col_key in enumerate(sorted_col_keys):
                col_name = output_column_dict[col_key]
                value = row.el[col_name].value
                if isinstance(value, list):
                    value = ', '.join(str(v) for v in value) if value else ""

                wrap = (
                    (col_name == UL_SOURCE_FILES and isinstance(value, str) and "\n" in value)
                    or col_name == UNITS_CONVERSION_TRACE
                )
                link_url = None
                if (
                    col_name == UL_SOURCE_FILES
                    and isinstance(value, str)
                    and value.strip()
                ):
                    value, link_url = excel_path_link(
                        col_name,
                        value,
                        packing_root=packing_root,
                    )
                if wrap and isinstance(value, str):
                    source_line_count = max(source_line_count, value.count("\n") + 1)
                el_color = row.el[col_name].color
                color_hex = "" if el_color == Color.no else el_color
                is_link = bool(link_url)
                if color_hex or is_link:
                    fmt = _get_cell_format(color_hex, wrap=wrap, is_link=is_link)
                else:
                    fmt = cell_wrap_fmt if wrap else cell_base_fmt

                if col_name in NUMERIC_QUANTITY_COLUMNS:
                    value = _quantity_value_for_excel(value)
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    ws.write_number(excel_row, ci, value, fmt)
                elif link_url:
                    url_status = ws.write_url(
                        excel_row,
                        ci,
                        link_url,
                        fmt,
                        value,
                    )
                    if url_status != 0:
                        ws.write(excel_row, ci, value, fmt)
                else:
                    ws.write(excel_row, ci, value, fmt)

                comment = row.el[col_name].comment
                if should_write_excel_comment(col_name, row, comment):
                    comment_cells += 1
                    try:
                        result = ws.write_comment(
                            excel_row,
                            ci,
                            str(comment),
                            {"width": 200, "height": 100},
                        )
                        if result != 0:
                            comment_write_errors.append(
                                f"row={excel_row + 1}, column={col_name}, return_code={result}"
                            )
                    except Exception as exc:
                        comment_write_errors.append(
                            f"row={excel_row + 1}, column={col_name}, error={exc!r}"
                        )

            if source_line_count > 1:
                ws.set_row(
                    excel_row,
                    min(15.0 * source_line_count + 2.0, 90.0),
                )

        t_write = time.perf_counter() - t_write_start
        append_timing_log(result_dir, f"save_match_result_to_excel::write+styles: {t_write:.3f}s")

        # Column widths + outline grouping (summary column is always level 0).
        col_group_opts = outline_options_for_visible(column_config)
        for ci in range(num_cols):
            w = col_width.get(ci, DEFAULT_COLUMN_WIDTH)
            opts = col_group_opts.get(ci) or {"level": 0}
            ws.set_column(ci, ci, w, None, opts)

        # Auto-filter and freeze
        if total_rows > 0:
            ws.autofilter(0, 0, total_rows, num_cols - 1)
        ws.freeze_panes(1, 0)

        write_quality_sheets(wb, result_rows, packing_audit)

        t_save_start = time.perf_counter()
        print(
            "Запись Excel: сериализация локально "
            f"(комментариев={comment_cells})...",
            flush=True,
        )
        wb.close()
        t_close = time.perf_counter() - t_save_start
        size_mb = os.path.getsize(tmp_path) / (1024 * 1024)
        append_timing_log(
            result_dir, f"save_match_result_to_excel::wb_close: {t_close:.3f}s"
        )
        print(
            f"Запись Excel: локальный xlsx {size_mb:.1f} MB за {t_close:.1f}s",
            flush=True,
        )

        t_copy_start = time.perf_counter()
        if _path_is_unc(file_path):
            print("Запись Excel: копирование на сетевой диск...", flush=True)
        shutil.copyfile(tmp_path, file_path)
        publish_ok = True
        t_copy = time.perf_counter() - t_copy_start
        append_timing_log(
            result_dir, f"save_match_result_to_excel::copy_to_dest: {t_copy:.3f}s"
        )
        if t_copy >= 1.0:
            print(f"Запись Excel: копия {t_copy:.1f}s", flush=True)

        if comment_write_errors:
            print(
                "Step4 Excel WARNING: "
                f"не записано комментариев={len(comment_write_errors)}; файл={file_name}"
            )
            for error in comment_write_errors:
                print(f"  {error}")

        if debug_log is not None:
            debug_log.append(
                f"  [save_match_result_to_excel] write+styles={t_write:.2f}s, "
                f"wb.close={t_close:.2f}s, copy={t_copy:.2f}s, "
                f"comments={comment_cells}, "
                f"comment_errors={len(comment_write_errors)}\n"
            )

        print(f"Запись Excel: {total_rows} строк -> {file_name}")
        return file_path

    except Exception as e:
        print(f"Ошибка при сохранении результата сопоставления в Excel: {e}")
        import traceback
        traceback.print_exc()
        if tmp_path and os.path.isfile(tmp_path) and not publish_ok:
            print(f"Локальная копия Excel оставлена: {tmp_path}")
            tmp_path = None
        raise
    finally:
        if tmp_path:
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def _insert_rfp_lot_column(
    filtered_rfp_dict: dict[int, str],
    rfp_column_dict: dict[int, str]
) -> dict[int, str]:
    ds_lot_def = next((d for d in OUTPUT_COLUMNS_CONFIG if d.col_name == DS_LOT), None)
    if ds_lot_def is not None and not ds_lot_def.output:
        return filtered_rfp_dict

    ds_name_idx = None
    ds_number_idx = None
    for idx, col in rfp_column_dict.items():
        if col == DS_NAME:
            ds_name_idx = idx
        elif col == DS_NUMBER:
            ds_number_idx = idx
    if ds_name_idx is None or ds_number_idx is None or ds_name_idx >= ds_number_idx:
        return filtered_rfp_dict

    updated_dict: dict[int, str] = {}
    for idx, col in filtered_rfp_dict.items():
        new_idx = idx + 1 if idx > ds_name_idx else idx
        updated_dict[new_idx] = col
    updated_dict[ds_name_idx + 1] = DS_LOT
    return updated_dict
