# -*- coding: utf-8 -*-
"""
MTO grid diagnostic -- analyses stamp-area grid lines only (filtered by template bbox).

Usage:
    python -m pdf_parsing_v2.mto_grid_diagnostic
    python -m pdf_parsing_v2.mto_grid_diagnostic "path/to/MTO/folder"
"""

from __future__ import annotations

import os
import sys

import fitz

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from pdf_parsing_v2_engine.coord_transform import field_to_fitz_rect, SCALE
from pdf_parsing_v2_engine.frame_detector import find_frame
from pdf_parsing_v2_engine.grid_matcher import cluster_lines, get_detected_stamp_cells
from pdf_parsing_v2_engine.models import StampTemplate
from pdf_parsing_v2.v2_config import load_v2_config

# Reference stamp size (for comparison)
REFERENCE_STAMP_W_MM = 185.0
REFERENCE_STAMP_H_MM = 110.0


def _pts_to_mm(pts: float) -> float:
    return pts / SCALE


def _get_stamp_bbox_fitz(template: StampTemplate, frame) -> fitz.Rect:
    """Compute stamp bbox in fitz displayed coordinates from all non-outside fields."""
    rects = []
    for f in template.fields:
        if not f.outside_stamp:
            r = field_to_fitz_rect(f, frame, padding_mm=0.0)
            if not r.is_empty and not r.is_infinite:
                rects.append(r)
    if not rects:
        return fitz.Rect()
    x0 = min(r.x0 for r in rects)
    y0 = min(r.y0 for r in rects)
    x1 = max(r.x1 for r in rects)
    y1 = max(r.y1 for r in rects)
    # Add 10mm margin
    m = 10 * SCALE
    return fitz.Rect(x0 - m, y0 - m, x1 + m, y1 + m)


def analyze_pdf(pdf_path: str, template: StampTemplate, cfg: dict) -> dict:
    """Analyze stamp grid lines for one PDF."""
    result = {
        "name": os.path.basename(pdf_path),
        "rotation": 0,
        "detected_cells_full": 0,
        "detected_cells_stamp": 0,
        "x_lines": [],
        "y_lines": [],
        "stamp_w_mm": 0.0,
        "stamp_h_mm": 0.0,
        "template_w_mm": 0.0,
        "template_h_mm": 0.0,
        "error": "",
    }
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]
        frame, _ = find_frame(page)
        result["rotation"] = frame.rotation

        # Template stamp bbox for this specific frame
        stamp_bbox = _get_stamp_bbox_fitz(template, frame)

        # Record template bbox size
        result["template_w_mm"] = _pts_to_mm(stamp_bbox.width - 20 * SCALE)  # minus margin
        result["template_h_mm"] = _pts_to_mm(stamp_bbox.height - 20 * SCALE)

        # All cells on page
        all_cells = get_detected_stamp_cells(
            fitz_page=page,
            stamp_bbox=page.rect,
            tolerance_mm=0.5,
            cfg=cfg,
            template=template,
        )
        result["detected_cells_full"] = len(all_cells)

        # Filtered to stamp area only
        stamp_cells = get_detected_stamp_cells(
            fitz_page=page,
            stamp_bbox=stamp_bbox,
            tolerance_mm=0.5,
            cfg=cfg,
            template=template,
        )
        doc.close()
        result["detected_cells_stamp"] = len(stamp_cells)

        if not stamp_cells:
            result["error"] = "no stamp cells"
            return result

        tol_pts = 0.5 * SCALE
        all_x = [p for c in stamp_cells for p in (c.x0, c.x1)]
        all_y = [p for c in stamp_cells for p in (c.y0, c.y1)]
        result["x_lines"] = cluster_lines(all_x, tol_pts)
        result["y_lines"] = cluster_lines(all_y, tol_pts)
        result["stamp_w_mm"] = _pts_to_mm(max(c.x1 for c in stamp_cells) - min(c.x0 for c in stamp_cells))
        result["stamp_h_mm"] = _pts_to_mm(max(c.y1 for c in stamp_cells) - min(c.y0 for c in stamp_cells))

    except Exception as exc:
        result["error"] = str(exc)
    return result


def run_diagnostic(mto_folder: str) -> None:
    cfg = load_v2_config()

    # Load the MTO/BBB page1 template
    templates_dir = cfg.get("templates_dir", os.path.join(_HERE, "templates"))
    template_path = os.path.join(templates_dir, "bbb_mto_page1.json")
    if not os.path.exists(template_path):
        print("Template not found: " + template_path)
        return
    template = StampTemplate.from_json(template_path)
    print("Template loaded: %s (%d fields)" % (os.path.basename(template_path), len(template.fields)))

    pdf_files = sorted(
        os.path.join(mto_folder, f)
        for f in os.listdir(mto_folder)
        if f.lower().endswith(".pdf")
    )
    if not pdf_files:
        print("No PDF files found in: " + mto_folder)
        return

    print("=" * 100)
    print("MTO Stamp Grid Diagnostic (filtered to stamp bbox)")
    print("Folder: " + mto_folder)
    print("Files: %d" % len(pdf_files))
    print("=" * 100)
    print()

    header = "%-44s %3s %6s %6s %7s %7s %7s %7s %6s %6s  %s" % (
        "File", "rot", "total", "stamp", "X-lns", "Y-lns", "W,mm", "H,mm", "dW%", "dH%", "err"
    )
    print(header)
    print("-" * 100)

    all_results = []
    for pdf_path in pdf_files:
        r = analyze_pdf(pdf_path, template, cfg)
        all_results.append(r)
        dw = (r["stamp_w_mm"] / REFERENCE_STAMP_W_MM - 1.0) * 100 if r["stamp_w_mm"] > 0 else 0.0
        dh = (r["stamp_h_mm"] / REFERENCE_STAMP_H_MM - 1.0) * 100 if r["stamp_h_mm"] > 0 else 0.0
        nx = len(r["x_lines"])
        ny = len(r["y_lines"])
        row = "%-44s %3d  %6d %6d %7d %7d %7.1f %7.1f %+6.1f%% %+6.1f%%  %s" % (
            r["name"][:44], r["rotation"], r["detected_cells_full"], r["detected_cells_stamp"],
            nx, ny, r["stamp_w_mm"], r["stamp_h_mm"], dw, dh, r["error"]
        )
        print(row)

    # Aggregate
    good = [r for r in all_results if not r["error"] and r["stamp_w_mm"] > 0]
    if not good:
        print("No valid results.")
        return

    print("-" * 100)
    ws = [r["stamp_w_mm"] for r in good]
    hs = [r["stamp_h_mm"] for r in good]
    dws = [(w / REFERENCE_STAMP_W_MM - 1.0) * 100 for w in ws]
    dhs = [(h / REFERENCE_STAMP_H_MM - 1.0) * 100 for h in hs]
    xs = [len(r["x_lines"]) for r in good]
    ys = [len(r["y_lines"]) for r in good]

    print()
    print("Aggregate (N=%d):" % len(good))
    print("  Stamp W (mm): min=%.1f  max=%.1f  mean=%.1f  SPREAD=%.1f mm (%.1f%%)" % (
        min(ws), max(ws), sum(ws)/len(ws), max(ws)-min(ws),
        (max(ws)-min(ws))/REFERENCE_STAMP_W_MM*100))
    print("  Stamp H (mm): min=%.1f  max=%.1f  mean=%.1f  SPREAD=%.1f mm (%.1f%%)" % (
        min(hs), max(hs), sum(hs)/len(hs), max(hs)-min(hs),
        (max(hs)-min(hs))/REFERENCE_STAMP_H_MM*100))
    print("  dW%%: min=%+.1f%%  max=%+.1f%%  spread=%.1f%%" % (min(dws), max(dws), max(dws)-min(dws)))
    print("  dH%%: min=%+.1f%%  max=%+.1f%%  spread=%.1f%%" % (min(dhs), max(dhs), max(dhs)-min(dhs)))
    print("  X-lines per file: min=%d  max=%d  mean=%.1f" % (min(xs), max(xs), sum(xs)/len(xs)))
    print("  Y-lines per file: min=%d  max=%d  mean=%.1f" % (min(ys), max(ys), sum(ys)/len(ys)))

    # Per-line spread analysis (for files with same count)
    print()
    x_counts = [len(r["x_lines"]) for r in good]
    y_counts = [len(r["y_lines"]) for r in good]

    common_x = min(x_counts)  # use min count for safe comparison
    common_y = min(y_counts)

    print("X-lines per file (mm, absolute fitz coords):")
    for r in good:
        xmm = ["%.1f" % _pts_to_mm(x) for x in r["x_lines"]]
        print("  %-40s [%s]" % (r["name"][:40], ", ".join(xmm)))

    print()
    print("Y-lines per file (mm, absolute fitz coords):")
    for r in good:
        ymm = ["%.1f" % _pts_to_mm(y) for y in r["y_lines"]]
        print("  %-40s [%s]" % (r["name"][:40], ", ".join(ymm)))

    print()
    print("--- Per-position spread (first %d X-lines, first %d Y-lines) ---" % (common_x, common_y))
    if common_x >= 2:
        x_spreads = []
        for i in range(common_x):
            vals = [_pts_to_mm(r["x_lines"][i]) for r in good]
            x_spreads.append(max(vals) - min(vals))
        print("  X-line spreads (mm): %s" % ", ".join("%.2f" % s for s in x_spreads))
        print("  X max_spread=%.2f mm  mean_spread=%.2f mm" % (max(x_spreads), sum(x_spreads)/len(x_spreads)))

    if common_y >= 2:
        y_spreads = []
        for i in range(common_y):
            vals = [_pts_to_mm(r["y_lines"][i]) for r in good]
            y_spreads.append(max(vals) - min(vals))
        print("  Y-line spreads (mm): %s" % ", ".join("%.2f" % s for s in y_spreads))
        print("  Y max_spread=%.2f mm  mean_spread=%.2f mm" % (max(y_spreads), sum(y_spreads)/len(y_spreads)))

    # Recommendation
    all_spreads = []
    if common_x >= 2:
        all_spreads += x_spreads
    if common_y >= 2:
        all_spreads += y_spreads

    if all_spreads:
        max_spread = max(all_spreads)
        # A template anchored at center needs to cover half the spread
        needed_snap = (max_spread / 2.0) * 1.5  # 1.5x safety margin
        print()
        print("=== SUMMARY ===")
        print("  Max per-line spread across files: %.2f mm" % max_spread)
        print("  -> With 1 anchor at center: needed max_snap_distance_mm >= %.1f mm" % needed_snap)
        print("  -> With 2 anchors at corners: needed max_snap_distance_mm >= %.1f mm" % (needed_snap * 0.5))
        print("  -> For 5%% distortion on %dmm: expected line shift = %.1f mm" % (
            REFERENCE_STAMP_W_MM, REFERENCE_STAMP_W_MM * 0.05 / 2))

    print()
    print("=" * 100)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        folder = sys.argv[1]
    else:
        folder = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "templates", "test_pdf", "MTO",
        )
    run_diagnostic(folder)
