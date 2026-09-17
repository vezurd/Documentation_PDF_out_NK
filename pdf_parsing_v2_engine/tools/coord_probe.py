"""
Standalone coordinate-pipeline diagnostic for MTO PDFs.

Renders overlay PNGs showing how frame detection + coordinate transforms
place template fields relative to actual page content.  Does NOT modify
any pipeline module — pure read-only analysis.

Run from repo root::

    set PYTHONPATH=.
    python -m pdf_parsing_v2_engine.tools.coord_probe
    python -m pdf_parsing_v2_engine.tools.coord_probe --dir "path\\to\\MTO"

Output: ``coord_probe_output/`` folder with one PNG per PDF (page 0) and
a summary table printed to stdout.
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

import fitz

from pdf_parsing_v2_engine.coord_transform import (
    field_to_fitz_rect,
    fitz_displayed_to_unrotated,
    pdfminer_to_fitz,
    resolve_field_bbox,
)
from pdf_parsing_v2_engine.find_tables_settings import call_find_tables
from pdf_parsing_v2_engine.frame_detector import find_frame
from pdf_parsing_v2_engine.models import FieldDef, FrameInfo, StampTemplate
from pdf_parsing_v2_engine.stamp_extractor import select_templates
from pdf_parsing_v2_engine.template_loader import load_all_templates
from pdf_parsing_v2_engine.models import StampTemplate as _ST_cls

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_DEFAULT_MTO = _PROJECT_ROOT / "pdf_parsing_v2_engine" / "templates" / "test_pdf" / "MTO"
_TEMPLATES_DIR = _PROJECT_ROOT / "pdf_parsing_v2_engine" / "templates"
_OUTPUT_DIR = _PROJECT_ROOT / "coord_probe_output"

_BAD = (
    "AGCC.287-4000-KSB.MTO-0001_0_RU",
    "AGCC.287-5631-KSB.MTO-0001_01-AN01_RU",
    "AGCC.287-7421-SOS.1.MTO-0001_01_RU",
    "AGCC.287-8525-SOS.MTO-0001_01_RU",
)
_GOOD = (
    "AGCC.287-7570-SOT.MTO-0001_01_RU",
    "AGCC.287-8441-SKUD.MTO-0001_04_RU",
    "AGCC.287-8630-KSB3.MTO-0001_02-AN01_RU",
)

_GREEN = (0, 0.8, 0)
_BLUE = (0, 0.3, 1)
_RED = (1, 0, 0)
_YELLOW = (1, 0.8, 0)
_CYAN = (0, 0.8, 0.8)
_MAGENTA = (0.8, 0, 0.8)
_ORANGE = (1, 0.5, 0)


def _label(name: str) -> str:
    stem = Path(name).stem
    for b in _BAD:
        if b in stem:
            return "BAD "
    for g in _GOOD:
        if g in stem:
            return "GOOD"
    return "??? "


def _frame_to_fitz_displayed(frame: FrameInfo) -> fitz.Rect:
    """Convert FrameInfo (pdfminer coords) → fitz displayed rect via pdfminer_to_fitz."""
    return pdfminer_to_fitz(
        (frame.x0, frame.y0, frame.x1, frame.y1),
        frame.page_height,
        frame.rotation,
    )


def _stamp_bbox(fields: list[FieldDef], frame: FrameInfo) -> fitz.Rect | None:
    """Union of non-outside_stamp field rects in fitz displayed coords."""
    rects = [
        field_to_fitz_rect(f, frame, padding_mm=0.0)
        for f in fields
        if not getattr(f, "outside_stamp", False)
    ]
    rects = [r for r in rects if not r.is_empty and not r.is_infinite]
    if not rects:
        return None
    x0 = min(r.x0 for r in rects)
    y0 = min(r.y0 for r in rects)
    x1 = max(r.x1 for r in rects)
    y1 = max(r.y1 for r in rects)
    return fitz.Rect(x0, y0, x1, y1)


def _draw(shape, rect: fitz.Rect, color, width=1.5, dashes: str | None = None):
    """Draw a rectangle outline on a fitz shape."""
    if rect.is_empty or rect.is_infinite:
        return
    try:
        shape.draw_rect(rect)
        kw = {"color": color, "width": width}
        if dashes:
            kw["dashes"] = dashes
        shape.finish(**kw)
    except Exception:
        pass


def _draw_label(page, text: str, pt: fitz.Point, color=(0, 0, 0), fontsize=7):
    """Insert a small text label at a point."""
    try:
        tw = fitz.TextWriter(page.rect)
        tw.append(pt, text, fontsize=fontsize)
        tw.write_text(page, color=color)
    except Exception:
        pass


def _get_find_tables_cells(fitz_page) -> list[fitz.Rect]:
    """Get all cells from find_tables as fitz.Rect list."""
    try:
        tables_obj = call_find_tables(fitz_page)
    except Exception:
        return []
    tables = getattr(tables_obj, "tables", tables_obj)
    cells: list[fitz.Rect] = []
    for t in tables:
        for c in getattr(t, "cells", []):
            if c and len(c) == 4:
                r = fitz.Rect(float(c[0]), float(c[1]), float(c[2]), float(c[3]))
                if r.width >= 2 and r.height >= 2:
                    cells.append(r)
    return cells


def _try_extract_text(fitz_page, fitz_rect: fitz.Rect, frame: FrameInfo) -> str:
    """Extract text from a fitz displayed rect (handles rotation)."""
    if fitz_rect.is_empty or fitz_rect.is_infinite:
        return ""
    query = fitz_displayed_to_unrotated(fitz_rect, fitz_page)
    try:
        return fitz_page.get_textbox(query).strip()[:60]
    except Exception:
        return "<err>"


def process_one_pdf(
    pdf_path: str,
    templates: list[StampTemplate],
    output_dir: str,
    doc_type: str = "MTO",
) -> dict:
    """Process one PDF: render overlay PNG, return summary dict."""
    name = os.path.basename(pdf_path)
    result = {"name": name, "label": _label(name)}

    doc = fitz.open(pdf_path)
    try:
        page = doc[0]
        rotation = page.rotation % 360
        page_w = page.rect.width
        page_h = page.rect.height
        result["rotation"] = rotation
        result["page_size"] = f"{page_w:.0f}x{page_h:.0f}"

        frame, dbg = find_frame(page, debug=True)
        result["frame_pm"] = f"({frame.x0:.1f},{frame.y0:.1f},{frame.x1:.1f},{frame.y1:.1f})"
        result["gost_fallback"] = dbg.get("gost_fallback_used", False) if dbg else "?"

        frame_fitz = _frame_to_fitz_displayed(frame)
        result["frame_fitz"] = (
            f"({frame_fitz.x0:.1f},{frame_fitz.y0:.1f},"
            f"{frame_fitz.x1:.1f},{frame_fitz.y1:.1f})"
        )

        candidates = select_templates(templates, doc_type, page_num=1)
        if not candidates:
            result["template"] = "NONE"
            result["error"] = "no template for MTO page 1"
            return result
        tmpl = candidates[0]
        result["template"] = tmpl.name[:40]

        sbbox = _stamp_bbox(tmpl.fields, frame)
        result["stamp_bbox_fitz"] = (
            f"({sbbox.x0:.1f},{sbbox.y0:.1f},{sbbox.x1:.1f},{sbbox.y1:.1f})"
            if sbbox else "None"
        )

        cells = _get_find_tables_cells(page)
        result["find_tables_cells"] = len(cells)

        stamp_overlap = 0
        if sbbox:
            for c in cells:
                if not (c & sbbox).is_empty:
                    stamp_overlap += 1
        result["cells_in_stamp"] = stamp_overlap

        sample_fields = [f for f in tmpl.fields if f.expected == "required"][:5]
        field_texts = {}
        for f in sample_fields:
            fr = field_to_fitz_rect(f, frame, padding_mm=0.5)
            txt = _try_extract_text(page, fr, frame)
            field_texts[f.id] = txt
        result["sample_fields"] = field_texts

        # --- Render overlay PNG ---
        dpi = 150
        pix = page.get_pixmap(dpi=dpi)
        tmp_doc = fitz.open()
        tmp_page = tmp_doc.new_page(width=pix.width, height=pix.height)
        tmp_page.insert_image(tmp_page.rect, pixmap=pix)

        scale = dpi / 72.0
        shape = tmp_page.new_shape()

        # Layer 1: frame (green, thick)
        fr_scaled = frame_fitz * scale
        _draw(shape, fr_scaled, _GREEN, width=3.0)
        _draw_label(tmp_page, "FRAME (from FrameInfo→pdfminer_to_fitz)", 
                    fitz.Point(fr_scaled.x0 + 2, fr_scaled.y0 + 10), _GREEN, 8)

        # Layer 2: stamp_bbox (blue, thick dashed)
        if sbbox:
            sb_scaled = sbbox * scale
            _draw(shape, sb_scaled, _BLUE, width=2.5)
            _draw_label(tmp_page, "STAMP_BBOX (union of fields)", 
                        fitz.Point(sb_scaled.x0 + 2, sb_scaled.y0 - 3), _BLUE, 8)

        # Layer 3: find_tables cells (red, thin)
        for c in cells[:200]:
            _draw(shape, c * scale, _RED, width=0.5)

        # Layer 4: sample template fields (yellow)
        for f in sample_fields:
            fr = field_to_fitz_rect(f, frame, padding_mm=0.5)
            fr_sc = fr * scale
            _draw(shape, fr_sc, _YELLOW, width=1.5)
            _draw_label(tmp_page, f.id[:25],
                        fitz.Point(fr_sc.x0 + 1, fr_sc.y0 + 8), _ORANGE, 6)

        # Layer 5: "correct" frame estimate — direct fitz displayed
        # For comparison: compute frame directly as a simple y-flip
        # (what frame_fitz SHOULD be if pdfminer coords were truly displayed)
        alt_frame = fitz.Rect(frame.x0, frame.page_height - frame.y1,
                              frame.x1, frame.page_height - frame.y0)
        if rotation in (90, 270):
            _draw(shape, alt_frame * scale, _CYAN, width=2.0)
            _draw_label(tmp_page, "ALT_FRAME (plain y-flip, no axis-unswap)", 
                        fitz.Point((alt_frame * scale).x0 + 2, 
                                   (alt_frame * scale).y1 + 10), _CYAN, 7)

        # Layer 6: "hypothesis fix" frame — undo axis swap for 90°
        if rotation in (90, 270):
            hyp_frame = fitz.Rect(frame.y0, frame.x0, frame.y1, frame.x1)
            _draw(shape, hyp_frame * scale, _MAGENTA, width=2.0)
            _draw_label(tmp_page, "HYP_FRAME (axis-unswap: pm_y→fitz_x, pm_x→fitz_y)", 
                        fitz.Point((hyp_frame * scale).x0 + 2, 
                                   (hyp_frame * scale).y1 + 22), _MAGENTA, 7)

        shape.commit()

        # Legend
        legend_y = 12
        for color, label_text in [
            (_GREEN, "GREEN = frame (current pdfminer_to_fitz)"),
            (_BLUE, "BLUE  = stamp_bbox (union of template fields)"),
            (_RED, "RED   = find_tables cells"),
            (_YELLOW, "YELLOW = sample required fields"),
            (_CYAN, "CYAN  = alt_frame (plain y-flip)  [only for rot!=0]"),
            (_MAGENTA, "MAGENTA = hyp_frame (axis-unswap) [only for rot!=0]"),
        ]:
            _draw_label(tmp_page, label_text, fitz.Point(5, legend_y), color, 9)
            legend_y += 12

        # Page info
        info_text = (
            f"rot={rotation}  page={page_w:.0f}x{page_h:.0f}  "
            f"frame_pm=({frame.x0:.0f},{frame.y0:.0f},{frame.x1:.0f},{frame.y1:.0f})  "
            f"cells={len(cells)}  overlap={stamp_overlap}  "
            f"label={result['label']}"
        )
        _draw_label(tmp_page, info_text, fitz.Point(5, pix.height - 10), (0, 0, 0), 8)

        out_name = Path(name).stem + "_coord_probe.png"
        out_path = os.path.join(output_dir, out_name)
        tmp_page.get_pixmap(dpi=72).save(out_path)
        tmp_doc.close()

        result["png"] = out_name

    finally:
        doc.close()

    return result


def _load_mto_templates(templates_dir: str) -> list[StampTemplate]:
    """Load MTO-applicable templates directly (bypass project loader)."""
    import json
    out: list[StampTemplate] = []
    for name in sorted(os.listdir(templates_dir)):
        if not name.lower().endswith(".json"):
            continue
        path = os.path.join(templates_dir, name)
        if not os.path.isfile(path):
            continue
        try:
            t = _ST_cls.from_json(path)
            if "MTO" in t.doc_types:
                out.append(t)
        except (OSError, ValueError, json.JSONDecodeError):
            pass
    return out


def main():
    parser = argparse.ArgumentParser(description="MTO coordinate pipeline diagnostic")
    parser.add_argument("--dir", default=str(_DEFAULT_MTO),
                        help="Directory with MTO test PDFs")
    parser.add_argument("--out", default=str(_OUTPUT_DIR),
                        help="Output directory for PNGs")
    args = parser.parse_args()

    mto_dir = Path(args.dir)
    if not mto_dir.exists():
        print(f"ERROR: directory not found: {mto_dir}", file=sys.stderr)
        sys.exit(1)

    pdfs = sorted(mto_dir.glob("*.pdf"))
    if not pdfs:
        print(f"ERROR: no PDFs in {mto_dir}", file=sys.stderr)
        sys.exit(1)

    os.makedirs(args.out, exist_ok=True)

    templates = _load_mto_templates(str(_TEMPLATES_DIR))
    print(f"Loaded {len(templates)} templates")
    print(f"Processing {len(pdfs)} PDFs from {mto_dir}")
    print(f"Output -> {args.out}\n")

    header = f"{'label':5s} {'rot':>4s} {'page_size':>10s} {'cells':>5s} {'ovlp':>5s} {'gost':>5s} {'name'}"
    print(header)
    print("-" * len(header) + "-" * 60)

    results = []
    for pdf in pdfs:
        try:
            r = process_one_pdf(str(pdf), templates, args.out)
            results.append(r)
            print(
                f"{r['label']:5s} "
                f"{r['rotation']:4d} "
                f"{r['page_size']:>10s} "
                f"{r.get('find_tables_cells', '?'):>5} "
                f"{r.get('cells_in_stamp', '?'):>5} "
                f"{str(r.get('gost_fallback', '?')):>5s} "
                f"{r['name']}"
            )
        except Exception as exc:
            print(f"ERROR {pdf.name}: {exc}", file=sys.stderr)

    # Detailed field extraction for each file
    print("\n\n=== Sample field texts ===")
    for r in results:
        sf = r.get("sample_fields", {})
        if sf:
            print(f"\n--- {r['label']} rot={r['rotation']} {r['name']} ---")
            print(f"    frame_pm   = {r.get('frame_pm', '?')}")
            print(f"    frame_fitz = {r.get('frame_fitz', '?')}")
            print(f"    stamp_bbox = {r.get('stamp_bbox_fitz', '?')}")
            for fid, txt in sf.items():
                status = "OK" if txt else "EMPTY"
                print(f"    [{status:5s}] {fid:30s} = {txt!r}")

    print(f"\n\nDone. {len(results)} overlay PNGs saved to: {args.out}")


if __name__ == "__main__":
    main()
