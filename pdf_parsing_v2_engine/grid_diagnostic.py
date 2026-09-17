"""
Phase 0 diagnostic for stamp grid stability across document types.

What it does:
- scans PDF files by type (BBB/MTO/OD/CJ/DWG)
- detects frame (find_frame) and table cells (fitz.find_tables)
- filters cells to stamp area (template bbox + margin)
- clusters X/Y logical lines in mm
- computes per-file and cross-type statistics
- builds anchor candidates from large cells
- runs special OD split check for field 6.2
- exports Excel report + PNG overlays
"""

from __future__ import annotations

import argparse
import math
import os
from dataclasses import dataclass
from datetime import datetime
from statistics import mean, pstdev
from typing import Any

import fitz
import xlsxwriter

from pdf_parsing_v2_engine.coord_transform import SCALE, field_to_fitz_rect
from pdf_parsing_v2_engine.find_tables_settings import call_find_tables
from pdf_parsing_v2_engine.frame_detector import find_frame
from pdf_parsing_v2_engine.models import FieldDef, StampTemplate
from pdf_parsing_v2.v2_config import load_v2_config


DEFAULT_INPUT_DIR = os.path.join("pdf_parsing_v2", "templates", "test_pdf")
DEFAULT_OUTPUT_DIR = os.path.join("pdf_parsing_v2", "grid_diagnostic_output")

TYPE_ORDER = ["DWG", "MTO", "BBB", "OD", "CJ"]
MARGIN_MM = 20.0
CLUSTER_TOLERANCE_MM = 0.5
ANCHOR_MIN_AREA_FRACTION = 0.05


@dataclass
class Cell:
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def width(self) -> float:
        return max(0.0, self.x1 - self.x0)

    @property
    def height(self) -> float:
        return max(0.0, self.y1 - self.y0)

    @property
    def area(self) -> float:
        return self.width * self.height

    @property
    def center(self) -> tuple[float, float]:
        return ((self.x0 + self.x1) * 0.5, (self.y0 + self.y1) * 0.5)

    def to_rect(self) -> fitz.Rect:
        return fitz.Rect(self.x0, self.y0, self.x1, self.y1)


@dataclass
class FileDiag:
    file_path: str
    rel_path: str
    doc_group: str
    n_cells_total: int
    n_cells_stamp: int
    n_x_lines: int
    n_y_lines: int
    x_lines_mm: list[float]
    y_lines_mm: list[float]
    x_widths_mm: list[float]
    y_heights_mm: list[float]
    frame_bbox_pts: tuple[float, float, float, float]
    stamp_bbox_pts: tuple[float, float, float, float]
    od_split_detected: bool
    od_6_2_cells: int
    od_6_2_inner_x_mm: list[float]
    overlay_png: str


@dataclass
class AnchorCandidate:
    doc_group: str
    file_name: str
    cx_n: float
    cy_n: float
    w_n: float
    h_n: float
    area_frac: float


def cluster_lines(coords: list[float], tolerance: float) -> list[float]:
    if not coords:
        return []
    s = sorted(coords)
    groups: list[list[float]] = [[s[0]]]
    for val in s[1:]:
        if abs(val - groups[-1][-1]) <= tolerance:
            groups[-1].append(val)
        else:
            groups.append([val])
    return [sum(g) / len(g) for g in groups]


def _norm_doc_group(rel_path: str) -> str:
    parts = [p.upper() for p in rel_path.replace("\\", "/").split("/") if p]
    for token in parts:
        if token in {"DWG", "MTO", "BBB", "OD", "CJ"}:
            return token
    return "UNKNOWN"


def _load_templates_map(root: str) -> dict[str, StampTemplate]:
    tdir = os.path.join(root, "pdf_parsing_v2", "templates")
    mapping: dict[str, str] = {
        "DWG": "dwg_page1.json",
        "MTO": "bbb_mto_page1.json",
        "BBB": "bbb_mto_page1.json",
        "OD": "od_page1.json",
        "CJ": "cj_page1.json",
    }
    out: dict[str, StampTemplate] = {}
    for group, fn in mapping.items():
        path = os.path.join(tdir, fn)
        if os.path.isfile(path):
            out[group] = StampTemplate.from_json(path)
    return out


def _collect_pdf_files(input_dir: str) -> list[str]:
    out: list[str] = []
    for dirpath, _, filenames in os.walk(input_dir):
        for fn in filenames:
            if fn.lower().endswith(".pdf"):
                out.append(os.path.join(dirpath, fn))
    out.sort()
    return out


def _extract_cells(page: fitz.Page, cfg: dict[str, Any] | None = None) -> list[Cell]:
    result: list[Cell] = []
    try:
        tables_obj = call_find_tables(page, cfg=cfg)
    except Exception:
        return result

    tables = getattr(tables_obj, "tables", tables_obj)
    for table in tables:
        for cell in getattr(table, "cells", []):
            if not cell or len(cell) != 4:
                continue
            x0, y0, x1, y1 = cell
            if x1 <= x0 or y1 <= y0:
                continue
            result.append(Cell(float(x0), float(y0), float(x1), float(y1)))

    dedup: dict[tuple[int, int, int, int], Cell] = {}
    for c in result:
        key = (round(c.x0), round(c.y0), round(c.x1), round(c.y1))
        dedup[key] = c
    return list(dedup.values())


def _build_stamp_bbox(page: fitz.Page, template: StampTemplate, margin_mm: float) -> fitz.Rect | None:
    frame, _ = find_frame(page)
    rects: list[fitz.Rect] = []
    for fd in template.fields:
        r = field_to_fitz_rect(fd, frame, padding_mm=0.0)
        if not r.is_empty and not r.is_infinite:
            rects.append(r)
    if not rects:
        return None
    x0 = min(r.x0 for r in rects)
    y0 = min(r.y0 for r in rects)
    x1 = max(r.x1 for r in rects)
    y1 = max(r.y1 for r in rects)
    pad = margin_mm * SCALE
    stamp = fitz.Rect(x0 - pad, y0 - pad, x1 + pad, y1 + pad)
    page_rect = fitz.Rect(0, 0, page.rect.width, page.rect.height)
    return stamp & page_rect


def _filter_cells_by_stamp(cells: list[Cell], stamp_bbox: fitz.Rect) -> list[Cell]:
    out: list[Cell] = []
    for c in cells:
        inter = c.to_rect() & stamp_bbox
        if not inter.is_empty and inter.get_area() > 0:
            out.append(c)
    return out


def _x_line_mm_from_right(x_pts: float, frame_x1: float) -> float:
    return (frame_x1 - x_pts) / SCALE


def _y_line_mm_from_bottom(y_fitz_pts: float, page_h: float, frame_y0: float) -> float:
    y_pm = page_h - y_fitz_pts
    return (y_pm - frame_y0) / SCALE


def _compute_line_stats(page: fitz.Page, frame: Any, cells: list[Cell]) -> tuple[list[float], list[float], list[float], list[float]]:
    x_mm_raw: list[float] = []
    y_mm_raw: list[float] = []
    for c in cells:
        x_mm_raw.append(_x_line_mm_from_right(c.x0, frame.x1))
        x_mm_raw.append(_x_line_mm_from_right(c.x1, frame.x1))
        y_mm_raw.append(_y_line_mm_from_bottom(c.y0, page.rect.height, frame.y0))
        y_mm_raw.append(_y_line_mm_from_bottom(c.y1, page.rect.height, frame.y0))

    x_lines = cluster_lines(x_mm_raw, CLUSTER_TOLERANCE_MM)
    y_lines = cluster_lines(y_mm_raw, CLUSTER_TOLERANCE_MM)
    x_lines.sort()
    y_lines.sort()

    x_widths = [round(x_lines[i + 1] - x_lines[i], 3) for i in range(len(x_lines) - 1)]
    y_heights = [round(y_lines[i + 1] - y_lines[i], 3) for i in range(len(y_lines) - 1)]
    return x_lines, y_lines, x_widths, y_heights


def _find_field(template: StampTemplate, field_id: str) -> FieldDef | None:
    for fd in template.fields:
        if fd.id == field_id:
            return fd
    return None


def _od_split_check(
    page: fitz.Page,
    frame: Any,
    cells_stamp: list[Cell],
    template_od: StampTemplate | None,
    x_lines_mm: list[float],
) -> tuple[bool, int, list[float]]:
    if template_od is None:
        return (False, 0, [])
    fd = _find_field(template_od, "6_2_Quantity_of_sheets")
    if fd is None:
        return (False, 0, [])

    rect = field_to_fitz_rect(fd, frame, padding_mm=1.0)
    hits: list[Cell] = []
    for c in cells_stamp:
        inter = c.to_rect() & rect
        if inter.is_empty:
            continue
        if inter.get_area() / max(c.area, 1e-9) >= 0.15:
            hits.append(c)

    left_mm = _x_line_mm_from_right(rect.x0, frame.x1)
    right_mm = _x_line_mm_from_right(rect.x1, frame.x1)
    lo = min(left_mm, right_mm)
    hi = max(left_mm, right_mm)
    inner_x = [x for x in x_lines_mm if lo + 0.5 < x < hi - 0.5]
    return (len(hits) > 1, len(hits), inner_x)


def _collect_anchor_candidates(
    doc_group: str,
    file_name: str,
    cells_stamp: list[Cell],
    stamp_bbox: fitz.Rect,
) -> list[AnchorCandidate]:
    out: list[AnchorCandidate] = []
    sw = max(stamp_bbox.width, 1e-9)
    sh = max(stamp_bbox.height, 1e-9)
    s_area = max(stamp_bbox.get_area(), 1e-9)
    for c in cells_stamp:
        frac = c.area / s_area
        if frac < ANCHOR_MIN_AREA_FRACTION:
            continue
        cx, cy = c.center
        out.append(
            AnchorCandidate(
                doc_group=doc_group,
                file_name=file_name,
                cx_n=(cx - stamp_bbox.x0) / sw,
                cy_n=(cy - stamp_bbox.y0) / sh,
                w_n=c.width / sw,
                h_n=c.height / sh,
                area_frac=frac,
            )
        )
    return out


def _cluster_anchor_candidates(cands: list[AnchorCandidate], n_files: int) -> list[dict[str, Any]]:
    clusters: list[dict[str, Any]] = []
    for c in cands:
        best_idx = -1
        best_dist = 1e9
        for i, cl in enumerate(clusters):
            d = max(
                abs(c.cx_n - cl["cx"]),
                abs(c.cy_n - cl["cy"]),
                abs(c.w_n - cl["w"]),
                abs(c.h_n - cl["h"]),
            )
            if d < best_dist:
                best_dist = d
                best_idx = i
        if best_idx >= 0 and best_dist <= 0.06:
            cl = clusters[best_idx]
            cl["items"].append(c)
            n = len(cl["items"])
            cl["cx"] = (cl["cx"] * (n - 1) + c.cx_n) / n
            cl["cy"] = (cl["cy"] * (n - 1) + c.cy_n) / n
            cl["w"] = (cl["w"] * (n - 1) + c.w_n) / n
            cl["h"] = (cl["h"] * (n - 1) + c.h_n) / n
            cl["area"] = (cl["area"] * (n - 1) + c.area_frac) / n
            cl["files"].add(c.file_name)
        else:
            clusters.append(
                {
                    "items": [c],
                    "cx": c.cx_n,
                    "cy": c.cy_n,
                    "w": c.w_n,
                    "h": c.h_n,
                    "area": c.area_frac,
                    "files": {c.file_name},
                }
            )

    out: list[dict[str, Any]] = []
    stable_threshold = max(1, math.ceil(n_files * 0.7))
    for idx, cl in enumerate(clusters, start=1):
        file_count = len(cl["files"])
        out.append(
            {
                "cluster_id": idx,
                "count": len(cl["items"]),
                "file_count": file_count,
                "coverage": file_count / max(n_files, 1),
                "stable": file_count >= stable_threshold,
                "cx": cl["cx"],
                "cy": cl["cy"],
                "w": cl["w"],
                "h": cl["h"],
                "area": cl["area"],
                "example_files": ", ".join(sorted(cl["files"])[:5]),
            }
        )
    out.sort(key=lambda x: (x["stable"], x["coverage"], x["area"]), reverse=True)
    return out


def _draw_overlay(
    page: fitz.Page,
    cells_stamp: list[Cell],
    stamp_bbox: fitz.Rect,
    x_lines_mm: list[float],
    y_lines_mm: list[float],
    frame: Any,
    out_png: str,
) -> None:
    dpi = 140
    scale = dpi / 72.0
    pix = page.get_pixmap(dpi=dpi)
    doc = fitz.open()
    op = doc.new_page(width=pix.width, height=pix.height)
    op.insert_image(op.rect, pixmap=pix)

    def _srect(r: fitz.Rect) -> fitz.Rect:
        return fitz.Rect(r.x0 * scale, r.y0 * scale, r.x1 * scale, r.y1 * scale)

    # Stamp region (red dashed).
    shape = op.new_shape()
    shape.draw_rect(_srect(stamp_bbox))
    shape.finish(color=(1, 0, 0), width=2.0, dashes="[6 4] 0")
    shape.commit()

    # Detected grid lines (green).
    xs: list[float] = []
    ys: list[float] = []
    for c in cells_stamp:
        xs.extend([c.x0, c.x1])
        ys.extend([c.y0, c.y1])
    xs = sorted(set(round(x, 1) for x in xs))
    ys = sorted(set(round(y, 1) for y in ys))
    for x in xs:
        shape = op.new_shape()
        shape.draw_line(
            fitz.Point(x * scale, stamp_bbox.y0 * scale),
            fitz.Point(x * scale, stamp_bbox.y1 * scale),
        )
        shape.finish(color=(0, 0.7, 0), width=1.0)
        shape.commit()
    for y in ys:
        shape = op.new_shape()
        shape.draw_line(
            fitz.Point(stamp_bbox.x0 * scale, y * scale),
            fitz.Point(stamp_bbox.x1 * scale, y * scale),
        )
        shape.finish(color=(0, 0.7, 0), width=1.0)
        shape.commit()

    # Clustered logical lines as blue points on stamp borders.
    for x_mm in x_lines_mm:
        x = frame.x1 - x_mm * SCALE
        for y in (stamp_bbox.y0, stamp_bbox.y1):
            shape = op.new_shape()
            shape.draw_circle(fitz.Point(x * scale, y * scale), 2.8)
            shape.finish(color=(0.1, 0.3, 1.0), fill=(0.1, 0.3, 1.0))
            shape.commit()

    for y_mm in y_lines_mm:
        y_pm = frame.y0 + y_mm * SCALE
        y_fitz = page.rect.height - y_pm
        for x in (stamp_bbox.x0, stamp_bbox.x1):
            shape = op.new_shape()
            shape.draw_circle(fitz.Point(x * scale, y_fitz * scale), 2.8)
            shape.finish(color=(0.1, 0.3, 1.0), fill=(0.1, 0.3, 1.0))
            shape.commit()

    os.makedirs(os.path.dirname(os.path.abspath(out_png)) or ".", exist_ok=True)
    op.get_pixmap().save(out_png)
    doc.close()


def _fmt_list(nums: list[float], ndigits: int = 2) -> str:
    if not nums:
        return ""
    return ", ".join(f"{x:.{ndigits}f}" for x in nums)


def _stats(vals: list[float]) -> tuple[float, float, float, float]:
    if not vals:
        return (0.0, 0.0, 0.0, 0.0)
    if len(vals) == 1:
        v = float(vals[0])
        return (v, 0.0, v, v)
    return (float(mean(vals)), float(pstdev(vals)), float(min(vals)), float(max(vals)))


def _save_excel_report(
    out_xlsx: str,
    rows: list[FileDiag],
    anchor_by_type: dict[str, list[dict[str, Any]]],
) -> None:
    wb = xlsxwriter.Workbook(out_xlsx, {"strings_to_urls": False})
    fmt_h = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
    fmt_c = wb.add_format({"border": 1})

    # Per-file
    ws = wb.add_worksheet("Per-file")
    headers = [
        "file_name", "rel_path", "type",
        "N_cells_total", "N_cells_stamp", "N_x_lines", "N_y_lines",
        "x_widths_mm", "y_heights_mm",
        "stamp_bbox_pts", "frame_bbox_pts",
        "OD_split_detected", "OD_6_2_cells", "OD_6_2_inner_x_mm", "overlay_png",
    ]
    for col, h in enumerate(headers):
        ws.write(0, col, h, fmt_h)
    for r, row in enumerate(rows, start=1):
        ws.write(r, 0, os.path.basename(row.file_path), fmt_c)
        ws.write(r, 1, row.rel_path, fmt_c)
        ws.write(r, 2, row.doc_group, fmt_c)
        ws.write(r, 3, row.n_cells_total, fmt_c)
        ws.write(r, 4, row.n_cells_stamp, fmt_c)
        ws.write(r, 5, row.n_x_lines, fmt_c)
        ws.write(r, 6, row.n_y_lines, fmt_c)
        ws.write(r, 7, _fmt_list(row.x_widths_mm), fmt_c)
        ws.write(r, 8, _fmt_list(row.y_heights_mm), fmt_c)
        ws.write(r, 9, _fmt_list(list(row.stamp_bbox_pts)), fmt_c)
        ws.write(r, 10, _fmt_list(list(row.frame_bbox_pts)), fmt_c)
        ws.write(r, 11, "yes" if row.od_split_detected else "no", fmt_c)
        ws.write(r, 12, row.od_6_2_cells, fmt_c)
        ws.write(r, 13, _fmt_list(row.od_6_2_inner_x_mm), fmt_c)
        ws.write(r, 14, row.overlay_png, fmt_c)
    ws.autofilter(0, 0, max(len(rows), 1), len(headers) - 1)
    ws.freeze_panes(1, 0)
    ws.set_column(0, 1, 42)
    ws.set_column(2, 6, 14)
    ws.set_column(7, 14, 40)

    # X-lines
    wsx = wb.add_worksheet("X-lines")
    max_x = max((len(r.x_lines_mm) for r in rows), default=0)
    wsx.write(0, 0, "file_name", fmt_h)
    wsx.write(0, 1, "type", fmt_h)
    for i in range(max_x):
        wsx.write(0, i + 2, f"X{i + 1}_mm_from_right", fmt_h)
    for r, row in enumerate(rows, start=1):
        wsx.write(r, 0, os.path.basename(row.file_path), fmt_c)
        wsx.write(r, 1, row.doc_group, fmt_c)
        for i, x in enumerate(row.x_lines_mm):
            wsx.write(r, i + 2, round(x, 3), fmt_c)
    wsx.autofilter(0, 0, max(len(rows), 1), max_x + 1)
    wsx.freeze_panes(1, 2)

    # Y-lines
    wsy = wb.add_worksheet("Y-lines")
    max_y = max((len(r.y_lines_mm) for r in rows), default=0)
    wsy.write(0, 0, "file_name", fmt_h)
    wsy.write(0, 1, "type", fmt_h)
    for i in range(max_y):
        wsy.write(0, i + 2, f"Y{i + 1}_mm_from_bottom", fmt_h)
    for r, row in enumerate(rows, start=1):
        wsy.write(r, 0, os.path.basename(row.file_path), fmt_c)
        wsy.write(r, 1, row.doc_group, fmt_c)
        for i, y in enumerate(row.y_lines_mm):
            wsy.write(r, i + 2, round(y, 3), fmt_c)
    wsy.autofilter(0, 0, max(len(rows), 1), max_y + 1)
    wsy.freeze_panes(1, 2)

    # Cross-type summary
    wss = wb.add_worksheet("Cross-type summary")
    sh = [
        "type", "files",
        "N_x_mean", "N_x_std", "N_x_min", "N_x_max",
        "N_y_mean", "N_y_std", "N_y_min", "N_y_max",
        "N_cells_stamp_mean", "N_cells_stamp_std", "N_cells_stamp_min", "N_cells_stamp_max",
    ]
    for c, h in enumerate(sh):
        wss.write(0, c, h, fmt_h)
    rr = 1
    for t in TYPE_ORDER + ["UNKNOWN"]:
        subset = [x for x in rows if x.doc_group == t]
        if not subset:
            continue
        sx = _stats([x.n_x_lines for x in subset])
        sy = _stats([x.n_y_lines for x in subset])
        sc = _stats([x.n_cells_stamp for x in subset])
        vals = [t, len(subset), *sx, *sy, *sc]
        for c, v in enumerate(vals):
            wss.write(rr, c, v, fmt_c)
        rr += 1
    wss.autofilter(0, 0, max(rr, 1), len(sh) - 1)

    # Anchors
    wsa = wb.add_worksheet("Anchors")
    ah = [
        "type", "cluster_id", "stable", "coverage", "files_in_cluster", "items_total",
        "cx_norm", "cy_norm", "w_norm", "h_norm", "mean_area_frac", "example_files",
    ]
    for c, h in enumerate(ah):
        wsa.write(0, c, h, fmt_h)
    ra = 1
    for t in TYPE_ORDER + ["UNKNOWN"]:
        for cl in anchor_by_type.get(t, []):
            vals = [
                t, cl["cluster_id"], "yes" if cl["stable"] else "no",
                round(cl["coverage"], 3), cl["file_count"], cl["count"],
                round(cl["cx"], 4), round(cl["cy"], 4),
                round(cl["w"], 4), round(cl["h"], 4),
                round(cl["area"], 4), cl["example_files"],
            ]
            for c, v in enumerate(vals):
                wsa.write(ra, c, v, fmt_c)
            ra += 1
    wsa.autofilter(0, 0, max(ra, 1), len(ah) - 1)
    wsa.set_column(11, 11, 42)

    # OD splits
    wso = wb.add_worksheet("OD_splits")
    oh = ["file_name", "rel_path", "split_detected", "cells_in_6_2", "inner_x_mm"]
    for c, h in enumerate(oh):
        wso.write(0, c, h, fmt_h)
    ro = 1
    for row in rows:
        if row.doc_group != "OD":
            continue
        vals = [
            os.path.basename(row.file_path),
            row.rel_path,
            "yes" if row.od_split_detected else "no",
            row.od_6_2_cells,
            _fmt_list(row.od_6_2_inner_x_mm),
        ]
        for c, v in enumerate(vals):
            wso.write(ro, c, v, fmt_c)
        ro += 1
    wso.autofilter(0, 0, max(ro, 1), len(oh) - 1)
    wso.set_column(0, 1, 42)
    wso.set_column(4, 4, 30)

    wb.close()


def run_grid_diagnostic(input_dir: str, output_dir: str) -> int:
    pdf_files = _collect_pdf_files(input_dir)
    if not pdf_files:
        print(f"[grid_diagnostic] PDF files not found in: {input_dir}")
        return 1

    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    templates = _load_templates_map(root)
    if not templates:
        print("[grid_diagnostic] Template JSON files are missing.")
        return 2

    ts = datetime.now().strftime("%Y.%m.%d_%H%M")
    out_dir = os.path.join(output_dir, f"grid_diagnostic_{ts}")
    overlay_dir = os.path.join(out_dir, "overlay_png")
    os.makedirs(overlay_dir, exist_ok=True)

    rows: list[FileDiag] = []
    anchors_all: list[AnchorCandidate] = []
    v2_cfg = load_v2_config()

    print(f"[grid_diagnostic] files: {len(pdf_files)}")
    for idx, fp in enumerate(pdf_files, start=1):
        rel_path = os.path.relpath(fp, input_dir)
        doc_group = _norm_doc_group(rel_path)
        print(f"[{idx}/{len(pdf_files)}] {rel_path} ({doc_group})")

        try:
            doc = fitz.open(fp)
            if len(doc) == 0:
                continue
            page = doc[0]
            frame, _ = find_frame(page)
            cells_all = _extract_cells(page, cfg=v2_cfg)

            tmpl = templates.get(doc_group)
            if tmpl is None:
                continue
            stamp_bbox = _build_stamp_bbox(page, tmpl, MARGIN_MM)
            if stamp_bbox is None:
                continue

            cells_stamp = _filter_cells_by_stamp(cells_all, stamp_bbox)
            x_lines_mm, y_lines_mm, xw_mm, yh_mm = _compute_line_stats(page, frame, cells_stamp)

            od_split = False
            od_cnt = 0
            od_inner_x: list[float] = []
            if doc_group == "OD":
                od_split, od_cnt, od_inner_x = _od_split_check(
                    page, frame, cells_stamp, templates.get("OD"), x_lines_mm
                )

            basename = os.path.splitext(os.path.basename(fp))[0]
            overlay_png = os.path.join("overlay_png", f"{basename}_p1.png")
            _draw_overlay(
                page=page,
                cells_stamp=cells_stamp,
                stamp_bbox=stamp_bbox,
                x_lines_mm=x_lines_mm,
                y_lines_mm=y_lines_mm,
                frame=frame,
                out_png=os.path.join(out_dir, overlay_png),
            )

            rows.append(
                FileDiag(
                    file_path=fp,
                    rel_path=rel_path,
                    doc_group=doc_group,
                    n_cells_total=len(cells_all),
                    n_cells_stamp=len(cells_stamp),
                    n_x_lines=len(x_lines_mm),
                    n_y_lines=len(y_lines_mm),
                    x_lines_mm=x_lines_mm,
                    y_lines_mm=y_lines_mm,
                    x_widths_mm=xw_mm,
                    y_heights_mm=yh_mm,
                    frame_bbox_pts=(frame.x0, frame.y0, frame.x1, frame.y1),
                    stamp_bbox_pts=(stamp_bbox.x0, stamp_bbox.y0, stamp_bbox.x1, stamp_bbox.y1),
                    od_split_detected=od_split,
                    od_6_2_cells=od_cnt,
                    od_6_2_inner_x_mm=od_inner_x,
                    overlay_png=overlay_png,
                )
            )

            anchors_all.extend(
                _collect_anchor_candidates(doc_group, os.path.basename(fp), cells_stamp, stamp_bbox)
            )
        finally:
            try:
                doc.close()
            except Exception:
                pass

    if not rows:
        print("[grid_diagnostic] No diagnostics produced.")
        return 3

    anchor_by_type: dict[str, list[dict[str, Any]]] = {}
    for t in TYPE_ORDER + ["UNKNOWN"]:
        cands = [x for x in anchors_all if x.doc_group == t]
        n_files = len({r.file_path for r in rows if r.doc_group == t})
        if cands and n_files > 0:
            anchor_by_type[t] = _cluster_anchor_candidates(cands, n_files)

    out_xlsx = os.path.join(out_dir, "grid_diagnostic_report.xlsx")
    _save_excel_report(out_xlsx, rows, anchor_by_type)

    # Console summary by type.
    print("\n=== Grid Diagnostic Summary ===")
    for t in TYPE_ORDER + ["UNKNOWN"]:
        subset = [x for x in rows if x.doc_group == t]
        if not subset:
            continue
        mx, sx, _, _ = _stats([x.n_x_lines for x in subset])
        my, sy, _, _ = _stats([x.n_y_lines for x in subset])
        od_splits = sum(1 for x in subset if x.od_split_detected)
        extra = f", splits={od_splits}" if t == "OD" else ""
        print(f"{t:4s} ({len(subset):2d} files): X-lines {mx:.1f}±{sx:.1f}, Y-lines {my:.1f}±{sy:.1f}{extra}")

    print(f"\n[grid_diagnostic] report: {out_xlsx}")
    print(f"[grid_diagnostic] overlays: {os.path.join(out_dir, 'overlay_png')}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Phase 0 stamp grid diagnostic.")
    parser.add_argument("--input-dir", default=DEFAULT_INPUT_DIR, help="Folder with PDF files.")
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR, help="Folder for report and overlays.")
    args = parser.parse_args()
    return run_grid_diagnostic(
        input_dir=os.path.abspath(args.input_dir),
        output_dir=os.path.abspath(args.output_dir),
    )


if __name__ == "__main__":
    raise SystemExit(main())
