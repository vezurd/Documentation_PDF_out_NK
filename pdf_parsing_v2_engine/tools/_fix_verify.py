"""
Verify the proposed fix: corrected coordinate mapping for rotation=90 pages.

Findings from _coord_deep_diag.py:
1. get_drawings() returns UNROTATED fitz coordinates for 90deg pages
2. find_tables() returns DISPLAYED fitz coordinates for 90deg pages
3. get_textbox() works in UNROTATED fitz coordinates
4. fitz_displayed_to_unrotated is WRONG:
   - Current:  (ph-y1, x0, ph-y0, x1)  with ph=page_height(842)
   - Correct:  (y0, pw-x1, y1, pw-x0)  with pw=page_width(1191)
   - Verified via page.derotation_matrix = Matrix(0,-1,1,0,0,1191)
5. find_frame stores coordinates that are NOT true pdfminer displayed:
   - Stores:  pm_x=r.y, pm_y=r.x  (axis swap only, no offsets)
   - Correct: pm_x=pw-r.y, pm_y=ph-r.x  (axis swap + page-size offsets)

This script tests the fixes WITHOUT modifying any pipeline module.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import fitz

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_DEFAULT_MTO = _PROJECT_ROOT / "pdf_parsing_v2_engine" / "templates" / "test_pdf" / "MTO"
_TEMPLATES_DIR = _PROJECT_ROOT / "pdf_parsing_v2_engine" / "templates"
_OUTPUT_DIR = _PROJECT_ROOT / "coord_probe_output"

SCALE = 2.83444


def _fixed_find_frame(fitz_page):
    """find_frame with corrected 90deg coordinate mapping."""
    page_w = fitz_page.rect.width
    page_h = fitz_page.rect.height
    rotation = fitz_page.rotation % 360

    best_pm_x0 = best_pm_x1 = 0.0
    best_pm_y0 = best_pm_y1 = 0.0
    best_x_range = 0.0
    best_y_range = 0.0

    if rotation in (90, 270):
        for d in fitz_page.get_drawings():
            r = d["rect"]
            if r.y1 < (page_w - SCALE) and r.x1 < (page_h - SCALE):
                pm_x_range = r.y1 - r.y0
                pm_y_range = r.x1 - r.x0
                if pm_x_range > best_x_range:
                    # FIX: apply page_w offset to convert to displayed pdfminer
                    best_pm_x0 = page_w - r.y1
                    best_pm_x1 = page_w - r.y0
                    best_x_range = pm_x_range
                if pm_y_range > best_y_range:
                    # FIX: apply page_h offset (y-flip) to convert to displayed pdfminer
                    best_pm_y0 = page_h - r.x1
                    best_pm_y1 = page_h - r.x0
                    best_y_range = pm_y_range
    else:
        raw_y0 = raw_y1 = 0.0
        for d in fitz_page.get_drawings():
            r = d["rect"]
            if r.x1 < (page_w - SCALE) and r.y0 > SCALE:
                x_range = r.x1 - r.x0
                y_range = r.y1 - r.y0
                if x_range > best_x_range:
                    best_pm_x0 = r.x0
                    best_pm_x1 = r.x1
                    best_x_range = x_range
                if y_range > best_y_range:
                    raw_y0 = r.y0
                    raw_y1 = r.y1
                    best_y_range = y_range
        best_pm_y0 = page_h - raw_y1
        best_pm_y1 = page_h - raw_y0

    from pdf_parsing_v2_engine.models import FrameInfo
    return FrameInfo(
        x0=best_pm_x0, y0=best_pm_y0, x1=best_pm_x1, y1=best_pm_y1,
        page_width=page_w, page_height=page_h, rotation=rotation,
        border_left_mm=0, border_bottom_mm=0, border_top_mm=0,
    )


def _resolve_field_bbox(bbox_mm, frame, origin="frame_bottom_right"):
    """Simplified resolve_field_bbox for testing."""
    h1, v1, h2, v2 = bbox_mm
    if origin == "frame_bottom_right":
        ax0 = frame.x1 - h1 * SCALE
        ay0 = frame.y0 + v1 * SCALE
        ax1 = frame.x1 - h2 * SCALE
        ay1 = frame.y0 + v2 * SCALE
    else:
        raise NotImplementedError(origin)
    if ax0 > ax1:
        ax0, ax1 = ax1, ax0
    if ay0 > ay1:
        ay0, ay1 = ay1, ay0
    return (ax0, ay0, ax1, ay1)


def _pdfminer_to_fitz(rect_pm, page_height):
    pm_x0, pm_y0, pm_x1, pm_y1 = rect_pm
    return fitz.Rect(pm_x0, page_height - pm_y1, pm_x1, page_height - pm_y0)


def test_one_pdf(pdf_path: str, fields_to_test: list[dict]):
    name = os.path.basename(pdf_path)
    doc = fitz.open(pdf_path)
    try:
        page = doc[0]
        rot = page.rotation % 360
        pw, ph = page.rect.width, page.rect.height

        # Current (buggy) frame
        from pdf_parsing_v2_engine.frame_detector import find_frame
        from pdf_parsing_v2_engine.coord_transform import (
            field_to_fitz_rect, fitz_displayed_to_unrotated, pdfminer_to_fitz,
        )
        old_frame, _ = find_frame(page)

        # Fixed frame
        new_frame = _fixed_find_frame(page)

        print(f"\n{'='*70}")
        print(f"{name}  rot={rot}  page={pw:.0f}x{ph:.0f}")
        print(f"  OLD frame_pm: ({old_frame.x0:.1f}, {old_frame.y0:.1f}, {old_frame.x1:.1f}, {old_frame.y1:.1f})")
        print(f"  NEW frame_pm: ({new_frame.x0:.1f}, {new_frame.y0:.1f}, {new_frame.x1:.1f}, {new_frame.y1:.1f})")

        for fdef in fields_to_test:
            fid = fdef["id"]
            bbox_mm = fdef["bbox_mm"]
            print(f"\n  --- {fid} bbox_mm={[round(b,1) for b in bbox_mm]} ---")

            # OLD pipeline
            old_pm = _resolve_field_bbox(bbox_mm, old_frame)
            old_fitz = _pdfminer_to_fitz(old_pm, old_frame.page_height)
            old_unrot = fitz_displayed_to_unrotated(old_fitz, page)
            try:
                old_text = page.get_textbox(old_unrot).strip()[:60]
            except Exception:
                old_text = "<err>"

            # NEW pipeline (fixed)
            new_pm = _resolve_field_bbox(bbox_mm, new_frame)
            new_fitz = _pdfminer_to_fitz(new_pm, new_frame.page_height)
            new_unrot = fitz_displayed_to_unrotated(new_fitz, page)
            try:
                new_text = page.get_textbox(new_unrot).strip()[:60]
            except Exception:
                new_text = "<err>"

            print(f"    OLD pm_rect:  ({old_pm[0]:.1f},{old_pm[1]:.1f},{old_pm[2]:.1f},{old_pm[3]:.1f})")
            print(f"    OLD fitz:     ({old_fitz.x0:.1f},{old_fitz.y0:.1f},{old_fitz.x1:.1f},{old_fitz.y1:.1f})")
            print(f"    OLD unrot:    ({old_unrot.x0:.1f},{old_unrot.y0:.1f},{old_unrot.x1:.1f},{old_unrot.y1:.1f})")
            print(f"    OLD text:     {old_text!r}")
            print()
            print(f"    NEW pm_rect:  ({new_pm[0]:.1f},{new_pm[1]:.1f},{new_pm[2]:.1f},{new_pm[3]:.1f})")
            print(f"    NEW fitz:     ({new_fitz.x0:.1f},{new_fitz.y0:.1f},{new_fitz.x1:.1f},{new_fitz.y1:.1f})")
            print(f"    NEW unrot:    ({new_unrot.x0:.1f},{new_unrot.y0:.1f},{new_unrot.x1:.1f},{new_unrot.y1:.1f})")
            print(f"    NEW text:     {new_text!r}")

            if old_text != new_text:
                changed = "IMPROVED" if new_text and not old_text else "CHANGED"
                print(f"    >>> {changed}")

    finally:
        doc.close()


def main():
    mto_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else _DEFAULT_MTO

    # Load required field definitions from template
    tmpl_path = _TEMPLATES_DIR / "bbb_mto_page1.json"
    with open(tmpl_path, encoding="utf-8") as f:
        tmpl = json.load(f)

    fields_to_test = [
        fd for fd in tmpl["fields"]
        if fd.get("expected") == "required"
    ]
    print(f"Testing {len(fields_to_test)} required fields: "
          f"{[f['id'] for f in fields_to_test]}")

    pdfs = sorted(mto_dir.glob("*.pdf"))
    for pdf in pdfs:
        try:
            test_one_pdf(str(pdf), fields_to_test)
        except Exception as exc:
            print(f"ERROR {pdf.name}: {exc}")


if __name__ == "__main__":
    main()
