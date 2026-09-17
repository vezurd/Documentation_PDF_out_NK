"""Low-level OD table extraction from PDF via pdfplumber (v2 port of od_raw_read_file)."""

from __future__ import annotations

from typing import Any, Sequence

import pdfplumber
from pdfplumber.table import TableSettings

from pdf_parsing_v2_od.od_debug_trace import append_od_debug_trace

MAX_PAGE_NUM = 4

# Merge pdfplumber columns narrower than this (pt) into the column to the left.
# 15 mm in PDF user space (1 pt = 1/72 in).
_MM_TO_PDF_PT = 72.0 / 25.4
OD_MERGE_NARROW_COLUMN_MM = 15.0
OD_MERGE_NARROW_COLUMN_PT = OD_MERGE_NARROW_COLUMN_MM * _MM_TO_PDF_PT

# When merging several column bboxes, collect chars and order LTR / TTB.
OD_MERGE_CHAR_LINE_Y_TOL_PT = 3.0
OD_MERGE_CHAR_SPACE_GAP_PT = 1.5


def _cell_str_or_none(val: str | None) -> str | None:
    if val is None:
        return None
    s = str(val).strip()
    if not s or s == "None":
        return None
    return s


def _union_bboxes(
    bboxes: list[tuple[float, float, float, float]],
) -> tuple[float, float, float, float]:
    return (
        min(b[0] for b in bboxes),
        min(b[1] for b in bboxes),
        max(b[2] for b in bboxes),
        max(b[3] for b in bboxes),
    )


def _char_midpoint_in_bbox(
    char: dict[str, Any],
    bbox: tuple[float, float, float, float],
) -> bool:
    """Same rule as pdfplumber ``Table.extract`` (midpoint inside cell)."""
    x0, top, x1, bottom = bbox
    h_mid = (float(char["x0"]) + float(char["x1"])) / 2
    v_mid = (float(char["top"]) + float(char["bottom"])) / 2
    return bool((h_mid >= x0) and (h_mid < x1) and (v_mid >= top) and (v_mid < bottom))


def _text_from_chars_ltr_ttb(
    chars: list[dict[str, Any]],
    *,
    y_tol: float = OD_MERGE_CHAR_LINE_Y_TOL_PT,
    space_gap_pt: float = OD_MERGE_CHAR_SPACE_GAP_PT,
) -> str:
    """Order glyphs left-to-right within each line, lines top-to-bottom."""
    if not chars:
        return ""
    by_y = sorted(chars, key=lambda c: (float(c["top"]), float(c["x0"])))
    lines: list[list[dict[str, Any]]] = []
    for c in by_y:
        if not lines:
            lines.append([c])
            continue
        ref = lines[-1][-1]
        if abs(float(c["top"]) - float(ref["top"])) <= y_tol:
            lines[-1].append(c)
        else:
            lines.append([c])
    out_lines: list[str] = []
    for line in lines:
        line_sorted = sorted(line, key=lambda c: float(c["x0"]))
        parts: list[str] = []
        prev_x1: float | None = None
        for c in line_sorted:
            x0 = float(c["x0"])
            x1 = float(c["x1"])
            t = c.get("text")
            if t is None:
                continue
            if prev_x1 is not None and x0 > prev_x1 + space_gap_pt:
                parts.append(" ")
            parts.append(str(t))
            prev_x1 = x1
        out_lines.append("".join(parts))
    return "\n".join(out_lines)


def _apply_column_merge_concat_fallback(
    text_rows: list[list[str | None]],
    groups: list[list[int]],
) -> list[list[str | None]]:
    """Join cell strings per group (used if geom/text row counts mismatch)."""
    out: list[list[str | None]] = []
    for row in text_rows:
        new_row: list[str | None] = []
        for grp in groups:
            parts: list[str] = []
            for j in grp:
                if j < len(row) and row[j] is not None:
                    s = str(row[j]).strip()
                    if s and s != "None":
                        parts.append(s)
            new_row.append("".join(parts) if parts else None)
        out.append(new_row)
    return out


def _apply_column_merge_with_union_bbox(
    table: Any,
    page: Any,
    text_rows: list[list[str | None]],
    groups: list[list[int]],
) -> list[list[str | None]]:
    """Merge columns: single-column groups keep ``extract()`` text; multi-column
    groups rebuild string from ``page.chars`` inside union bbox (LTR / TTB).

    If union extraction is empty, falls back to concatenating per-cell strings.
    """
    geom_rows = table.rows
    n_geom = len(geom_rows)
    n_text = len(text_rows)
    if n_geom != n_text:
        return _apply_column_merge_concat_fallback(text_rows, groups)

    n_rows = n_geom
    out: list[list[str | None]] = []
    for ri in range(n_rows):
        geom_row = geom_rows[ri]
        text_row = text_rows[ri]
        cells = geom_row.cells
        new_row: list[str | None] = []
        for grp in groups:
            if len(grp) == 1:
                j = grp[0]
                val = text_row[j] if j < len(text_row) else None
                new_row.append(_cell_str_or_none(val))
                continue

            bboxes: list[tuple[float, float, float, float]] = []
            for j in grp:
                if j < len(cells) and cells[j] is not None:
                    bboxes.append(
                        (
                            float(cells[j][0]),
                            float(cells[j][1]),
                            float(cells[j][2]),
                            float(cells[j][3]),
                        )
                    )
            fallback_parts: list[str] = []
            for j in grp:
                if j < len(text_row) and text_row[j] is not None:
                    s = str(text_row[j]).strip()
                    if s and s != "None":
                        fallback_parts.append(s)

            if not bboxes:
                joined = "".join(fallback_parts) if fallback_parts else ""
                new_row.append(_cell_str_or_none(joined or None))
                continue

            ubox = _union_bboxes(bboxes)
            ch_list = [
                c
                for c in page.chars
                if isinstance(c, dict) and _char_midpoint_in_bbox(c, ubox)
            ]
            merged = _text_from_chars_ltr_ttb(ch_list).strip()
            if not merged and fallback_parts:
                merged = "".join(fallback_parts)
            new_row.append(_cell_str_or_none(merged or None))
        out.append(new_row)
    return out


def _column_widths_pts_from_table(table: Any) -> list[float]:
    """Max cell width per column index from pdfplumber ``Table`` geometry."""
    rows = table.rows
    if not rows:
        return []
    n = len(rows[0].cells)
    widths = [0.0] * n
    for row in rows:
        for j, cell in enumerate(row.cells):
            if cell is not None and j < n:
                w = float(cell[2] - cell[0])
                if w > widths[j]:
                    widths[j] = w
    return widths


def _column_merge_groups(widths: list[float], min_w_pt: float) -> list[list[int]]:
    """Merge **narrow** columns into the column to the left.

    Only columns with ``0 < width < min_w_pt`` are merged. Width ``0`` means no
    cell bbox in that slot — do not merge, to avoid collapsing the whole grid.
    """
    n = len(widths)
    if n == 0:
        return []
    groups: list[list[int]] = [[0]]
    for j in range(1, n):
        w = widths[j]
        if 0 < w < min_w_pt:
            groups[-1].append(j)
        else:
            groups.append([j])
    return groups


def _extract_tables_with_narrow_column_merge(
    page: Any,
    settings: dict[str, Any],
    *,
    merge_debug: list[str] | None = None,
) -> list[list[list[str | None]]]:
    """Like ``Page.extract_tables`` but merge columns narrower than :data:`OD_MERGE_NARROW_COLUMN_PT`."""
    tset = TableSettings.resolve(settings)
    tables = page.find_tables(tset)
    text_kw = tset.text_settings or {}
    result: list[list[list[str | None]]] = []
    for ti, tbl in enumerate(tables):
        text_rows = tbl.extract(**text_kw)
        widths = _column_widths_pts_from_table(tbl)
        if not widths:
            result.append(text_rows)
            continue
        groups = _column_merge_groups(widths, OD_MERGE_NARROW_COLUMN_PT)
        n_before = len(widths)
        n_after = len(groups)
        if n_after == n_before:
            result.append(text_rows)
            if merge_debug is not None:
                merge_debug.append(
                    f"  table[{ti}]: column merge inactive "
                    f"(no column < {OD_MERGE_NARROW_COLUMN_MM} mm); "
                    f"widths_pt={[round(w, 1) for w in widths]}"
                )
        else:
            result.append(
                _apply_column_merge_with_union_bbox(tbl, page, text_rows, groups)
            )
            if merge_debug is not None:
                merge_debug.append(
                    f"  table[{ti}]: columns {n_before} -> {n_after} "
                    f"(merged into left if width < {OD_MERGE_NARROW_COLUMN_MM} mm); "
                    f"multi-col text = union bbox + chars LTR/TTB; "
                    f"widths_pt={[round(w, 1) for w in widths]}; groups={groups}"
                )
    return result


def _od_extract_table_settings() -> dict[str, Any]:
    """Settings for ``Page.extract_tables`` used by OD raw read.

    We only set ``edge_min_length``. A large ``snap_x_tolerance`` (e.g. tens of pt
    to mimic a “min column width”) merges *real* vertical grid lines on typical OD
    sheets and collapses the table to too few columns — not usable.
    """
    return {"edge_min_length": 60}


def _table_converter(table: list[list[str | None]]) -> str:
    """Convert a pdfplumber table to a structured pipe-delimited string."""
    rows: list[str] = []
    for row in table:
        cleaned = [
            item.replace("\n", " ")
            if item is not None and "\n" in item
            else "None"
            if item is None
            else item
            for item in row
        ]
        rows.append("|" + "|".join(cleaned) + "|")
    return "\n".join(rows)


def _crop_pdfplumber_page(
    page: Any,
    bbox: tuple[float, float, float, float],
    *,
    warnings_out: list[str] | None,
    page_number: int,
    file_name: str,
):
    """Crop page to bbox ``(x0, top, x1, bottom)`` in pdfplumber coordinates (origin top-left)."""
    pb = page.bbox
    x0, top, x1, bottom = bbox
    bb = (
        max(x0, pb[0]),
        max(top, pb[1]),
        min(x1, pb[2]),
        min(bottom, pb[3]),
    )
    if bb[0] >= bb[2] - 0.5 or bb[1] >= bb[3] - 0.5:
        msg = (
            f"Страница {page_number}: pdfplumber crop вырожден после пересечения "
            f"с bbox страницы ({file_name})."
        )
        if warnings_out is not None:
            warnings_out.append(msg)
        else:
            print(msg)
        return page
    try:
        return page.crop(bb, strict=False)
    except Exception as exc:
        msg = (
            f"Страница {page_number}: pdfplumber crop не применён ({exc}), "
            f"используется полная страница ({file_name})."
        )
        if warnings_out is not None:
            warnings_out.append(msg)
        else:
            print(msg)
        return page


def od_read_raw(
    file_name: str,
    *,
    warnings_out: list[str] | None = None,
    page_clip_rects: Sequence[tuple[float, float, float, float] | None] | None = None,
    debug_trace_out: list[dict[str, str]] | None = None,
) -> list[list[str]]:
    """Read raw table strings from an OD PDF (up to *MAX_PAGE_NUM* pages).

    Args:
        file_name: Path to the PDF.
        warnings_out: When set, short diagnostic strings are appended here instead
            of only printing to stdout (e.g. missing tables on a page).
        page_clip_rects: Optional per-page crop in displayed coordinates
            ``(x0, y0, x1, y1)`` (same convention as pdfplumber ``crop`` / MuPDF page).
            Index ``0`` = page 1. ``None`` at index *i* = use full page *i* + 1.
        debug_trace_out: When set, append structured blocks (pdfplumber stage).
    """
    tset = _od_extract_table_settings()
    append_od_debug_trace(
        debug_trace_out,
        function="od_raw_reader.od_read_raw",
        stage="entry",
        body=(
            f"file_name={file_name!r}\n"
            f"page_clip_rects={'set' if page_clip_rects is not None else 'None'} "
            f"(len={len(page_clip_rects) if page_clip_rects is not None else 0})\n"
            f"extract_tables: {tset!r} (defaults); then merge columns with width < "
            f"{OD_MERGE_NARROW_COLUMN_MM} mm (~{OD_MERGE_NARROW_COLUMN_PT:.2f} pt) into left; "
            f"multi-column cells: text from union bbox, chars LTR/TTB"
        ),
    )
    doc_content: list[list[str]] = []
    plumb_pdf = pdfplumber.open(file_name)
    try:
        for page in plumb_pdf.pages:
            if page.page_number > MAX_PAGE_NUM:
                continue
            pn = page.page_number
            idx = pn - 1
            clip_for_page: tuple[float, float, float, float] | None = None
            if page_clip_rects is not None and idx < len(page_clip_rects):
                clip_for_page = page_clip_rects[idx]
            if (
                page_clip_rects is not None
                and idx < len(page_clip_rects)
                and page_clip_rects[idx] is not None
            ):
                page = _crop_pdfplumber_page(
                    page,
                    page_clip_rects[idx],
                    warnings_out=warnings_out,
                    page_number=pn,
                    file_name=file_name,
                )
            merge_dbg: list[str] = []
            plumb_tables = _extract_tables_with_narrow_column_merge(
                page,
                _od_extract_table_settings(),
                merge_debug=merge_dbg if debug_trace_out is not None else None,
            )
            page_content: list[str] = []
            if not plumb_tables:
                msg = (
                    f"Страница {page.page_number}: pdfplumber не извлёк таблицы "
                    f"({file_name})."
                )
                if warnings_out is not None:
                    warnings_out.append(msg)
                else:
                    print(msg)
            for table in plumb_tables:
                if table is not None:
                    page_content.append(_table_converter(table))
                else:
                    msg = (
                        f"ERROR: OD TAB READER, table=NONE, {file_name}, "
                        f"стр.№ {page.page_number}"
                    )
                    if warnings_out is not None:
                        warnings_out.append(msg)
                    else:
                        print(msg)
                    page_content.append("")
            joined = "".join(page_content)
            tbl_lines: list[str] = []
            for ti, table in enumerate(plumb_tables or []):
                if table is None:
                    tbl_lines.append(f"  table[{ti}]: <None>")
                else:
                    conv = _table_converter(table)
                    tbl_lines.append(
                        f"  table[{ti}] pipe-delimited ({len(conv)} chars):\n{conv}"
                    )
            dbg_body = (
                f"page_number={pn}\n"
                f"crop_bbox={clip_for_page!r} (None = full page before crop)\n"
                f"tables_count={len(plumb_tables or [])}\n"
                f"joined_tables_len={len(joined)}\n"
            )
            if merge_dbg:
                dbg_body += "narrow-column merge:\n" + "\n".join(merge_dbg) + "\n"
            dbg_body += "\n".join(tbl_lines)
            append_od_debug_trace(
                debug_trace_out,
                function="od_raw_reader.od_read_raw",
                stage=f"page_{pn}_pdfplumber",
                body=dbg_body,
            )
            doc_content.append(page_content)
    finally:
        plumb_pdf.close()
    return doc_content
