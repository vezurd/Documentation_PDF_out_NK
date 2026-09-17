"""Debug Excel workbook for the v2 stamp pipeline.

Each row is one page (:class:`V2PageResult`). Columns are file metadata, template
info, then cleaned field values (conditional formatting). If cleaned text differs
from raw extraction, raw text is attached as a cell comment.

Field column order and headers come from :class:`FieldCatalog` (see ``models``).

Example::

    from pdf_parsing_v2_engine.v2_report import save_v2_debug_report
    save_v2_debug_report(all_results, result_dir, catalog)

Output filename pattern: ``Отладка_v2_штамп_<timestamp>.xlsx`` under *result_dir*.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Sequence

import xlsxwriter

from pdf_parsing_v2_engine.field_id_display import short_label_from_field_id
from pdf_parsing_v2_engine.models import FieldCatalog, V2PageResult


# ---------------------------------------------------------------------------
# Цвета секций и статусов
# ---------------------------------------------------------------------------

_COLOR_HEADER_META = "#4472C4"       # meta section (file, page)
_COLOR_HEADER_TEMPLATE = "#70AD47"   # template / score
_COLOR_HEADER_FIELD = "#9DC3E6"      # cleaned fields
_COLOR_HEADER_WARN = "#FFD966"       # warnings

_COLOR_VALID_OK = "#E2EFDA"          # valid field
_COLOR_VALID_FAIL = "#FCE4D6"        # invalid / empty required
_COLOR_SCORE_LOW = "#FFEB9C"         # score < 0.7
_COLOR_SCORE_CRITICAL = "#FFC7CE"    # score < 0.3

_FONT_NAME = "Calibri"
_FONT_SIZE = 10


# ---------------------------------------------------------------------------
# ColumnDef — описание столбца (по образцу step4_6_save_match_result_to_excel)
# ---------------------------------------------------------------------------

Section = str  # "meta" | "template" | "field_cleaned" | "warnings"


@dataclass
class ColumnDef:
    """Definition of one output column."""
    col_id: str
    header_label: str  # May contain ``\\n`` for a two-line header.
    section: Section
    width: int = 18
    field_id: str = ""  # Set when the column maps to ``V2PageResult.fields``.


# ---------------------------------------------------------------------------
# Фиксированные мета-столбцы
# ---------------------------------------------------------------------------

_META_COLUMNS: list[ColumnDef] = [
    ColumnDef("file_name",      "Имя файла",          "meta",     width=30),
    ColumnDef("doc_type",       "Тип",                "meta",     width=7),
    ColumnDef("page_num",       "Стр.",               "meta",     width=5),
]

_TEMPLATE_COLUMNS: list[ColumnDef] = [
    ColumnDef("template_name",  "Шаблон",             "template", width=22),
    ColumnDef("template_score", "Score",              "template", width=8),
    ColumnDef("warnings",       "Предупреждения",     "warnings", width=40),
]


def _header_lines_for_field(field_id: str, catalog: FieldCatalog) -> tuple[str, str]:
    """First and second header lines for a field column."""
    short = short_label_from_field_id(field_id)
    entry = catalog.get_entry(field_id)
    if entry is None:
        return short, field_id
    note = (entry.debug_report_header_note or "").strip()
    line2 = note if note else (entry.label or "")
    return short, line2


def _build_field_columns(field_ids: Sequence[str], catalog: FieldCatalog) -> list[ColumnDef]:
    """Build one cleaned-value column per field id."""
    cols: list[ColumnDef] = []
    for fid in field_ids:
        h1, h2 = _header_lines_for_field(fid, catalog)
        cols.append(ColumnDef(
            col_id=f"field_{fid}",
            header_label=f"{h1}\n{h2}",
            section="field_cleaned",
            width=20,
            field_id=fid,
        ))
    return cols


def _collect_field_ids_ordered(
    all_results: list[tuple[str, str, list[V2PageResult]]],
    catalog: FieldCatalog,
) -> list[str]:
    """Unique field ids: catalog-ordered block first, then unknown ids sorted by id."""
    present: set[str] = set()
    for _, _, pages in all_results:
        for p in pages:
            present.update(p.fields.keys())

    catalog_block: list[tuple[int, str]] = []
    tail: list[str] = []
    for fid in present:
        entry = catalog.get_entry(fid)
        if entry is None:
            tail.append(fid)
            continue
        if not entry.include_in_debug_report:
            continue
        catalog_block.append((entry.debug_report_order, fid))

    catalog_block.sort(key=lambda t: (t[0], t[1]))
    tail.sort()
    return [fid for _, fid in catalog_block] + tail


# ---------------------------------------------------------------------------
# Main export
# ---------------------------------------------------------------------------

def save_v2_debug_report(
    all_results: list[tuple[str, str, list[V2PageResult]]],
    result_dir: str,
    catalog: FieldCatalog,
) -> str:
    """Write the debug Excel report for extracted pages.

    Args:
        all_results: List of ``(file_path, doc_type, list[V2PageResult])``.
        result_dir: Output directory (created if missing).
        catalog: Field catalog controlling column order, inclusion, and headers.

    Returns:
        Absolute path to the written ``.xlsx`` file.
    """
    os.makedirs(result_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y.%m.%d_%Hч%Mм")
    out_path = os.path.join(result_dir, f"Отладка_v2_штамп_{ts}.xlsx")

    field_ids = _collect_field_ids_ordered(all_results, catalog)
    field_cols = _build_field_columns(field_ids, catalog)
    all_cols: list[ColumnDef] = _META_COLUMNS + _TEMPLATE_COLUMNS + field_cols

    wb = xlsxwriter.Workbook(out_path, {"strings_to_urls": False})
    ws = wb.add_worksheet("Штамп v2")

    # Formats
    fmt_base = wb.add_format({
        "font_name": _FONT_NAME, "font_size": _FONT_SIZE,
        "border": 1, "valign": "top", "text_wrap": True,
    })
    fmt_by_section: dict[Section, Any] = {}
    header_colors = {
        "meta": _COLOR_HEADER_META,
        "template": _COLOR_HEADER_TEMPLATE,
        "field_cleaned": _COLOR_HEADER_FIELD,
        "warnings": _COLOR_HEADER_WARN,
    }
    for section, color in header_colors.items():
        fmt_by_section[section] = wb.add_format({
            "font_name": _FONT_NAME, "font_size": _FONT_SIZE,
            "bold": True, "border": 1,
            "bg_color": color, "font_color": "#FFFFFF" if section == "meta" else "#000000",
            "valign": "bottom", "text_wrap": True, "align": "center",
        })

    fmt_ok = wb.add_format({"font_name": _FONT_NAME, "font_size": _FONT_SIZE,
                             "border": 1, "bg_color": _COLOR_VALID_OK, "text_wrap": True})
    fmt_fail = wb.add_format({"font_name": _FONT_NAME, "font_size": _FONT_SIZE,
                               "border": 1, "bg_color": _COLOR_VALID_FAIL, "text_wrap": True})
    fmt_score_low = wb.add_format({"font_name": _FONT_NAME, "font_size": _FONT_SIZE,
                                    "border": 1, "bg_color": _COLOR_SCORE_LOW})
    fmt_score_crit = wb.add_format({"font_name": _FONT_NAME, "font_size": _FONT_SIZE,
                                     "border": 1, "bg_color": _COLOR_SCORE_CRITICAL})

    # Column widths
    for col_idx, col in enumerate(all_cols):
        ws.set_column(col_idx, col_idx, col.width)

    # Header row
    ws.set_row(0, 40)
    for col_idx, col in enumerate(all_cols):
        ws.write(0, col_idx, col.header_label, fmt_by_section.get(col.section, fmt_base))

    ws.freeze_panes(1, len(_META_COLUMNS))
    ws.autofilter(0, 0, 0, len(all_cols) - 1)

    # Data rows
    row_idx = 1
    for file_path, doc_type, pages in all_results:
        file_name = os.path.basename(file_path)
        for page_result in pages:
            col_idx = 0

            # META
            ws.write(row_idx, col_idx, file_name, fmt_base); col_idx += 1
            ws.write(row_idx, col_idx, doc_type, fmt_base); col_idx += 1
            ws.write(row_idx, col_idx, page_result.page_num, fmt_base); col_idx += 1

            # TEMPLATE
            ws.write(row_idx, col_idx, page_result.template_name, fmt_base); col_idx += 1
            score = page_result.template_score
            score_fmt = fmt_score_crit if score < 0.3 else (fmt_score_low if score < 0.7 else fmt_base)
            ws.write(row_idx, col_idx, round(score, 3), score_fmt); col_idx += 1
            ws.write(row_idx, col_idx, "; ".join(page_result.warnings) if page_result.warnings else "", fmt_base)
            col_idx += 1

            # FIELDS
            for fc in field_cols:
                fid = fc.field_id
                fr = page_result.fields.get(fid)
                val = (fr.cleaned_value or "") if fr else ""
                cell_fmt = fmt_base
                if fr:
                    if fr.is_valid is True:
                        cell_fmt = fmt_ok
                    elif fr.is_valid is False:
                        cell_fmt = fmt_fail
                ws.write(row_idx, col_idx, val, cell_fmt)
                comment_parts: list[str] = []
                if fr and (fr.raw_value or "") != val:
                    comment_parts.append(f"RAW:\n{fr.raw_value}")
                if fr and fr.parse_warnings:
                    warn_lines = [f"{w.code}: {w.message}" for w in fr.parse_warnings]
                    comment_parts.append("Field warnings:\n" + "\n".join(warn_lines))
                if comment_parts:
                    ws.write_comment(
                        row_idx, col_idx, "\n\n".join(comment_parts),
                        {"width": 360, "height": 160},
                    )
                col_idx += 1

            row_idx += 1

    wb.close()
    print(f"[v2_report] Отчёт сохранён: {out_path} ({row_idx - 1} строк)")
    return out_path


# ---------------------------------------------------------------------------
# Pipeline helper: build *all_results* from V2Document list
# ---------------------------------------------------------------------------

def collect_results_from_curr_proj(curr_proj: list) -> list[tuple[str, str, list]]:
    """Build ``all_results`` from a list of V2Document objects.

    Args:
        curr_proj: Documents processed by the v2 pipeline; each may define
            ``_v2_results`` with ``list[V2PageResult]``.

    Returns:
        ``(file_full_path, doc_type, v2_results)`` tuples for documents that
        have non-empty ``_v2_results``.
    """
    out = []
    for doc in curr_proj:
        v2_results = getattr(doc, "_v2_results", None)
        if v2_results:
            out.append((doc.file_full_path, doc.doc_Type, v2_results))
    return out
