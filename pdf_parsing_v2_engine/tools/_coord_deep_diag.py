"""
Deep coordinate diagnostic: verify coordinate spaces empirically.

Checks:
1. Does get_drawings() return displayed or unrotated coordinates for 90deg pages?
2. Does get_textbox() work with unrotated coordinates on 90deg pages?
3. Where does text actually live on the page?
"""
from __future__ import annotations

import sys
from pathlib import Path

import fitz

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_DEFAULT_MTO = _PROJECT_ROOT / "pdf_parsing_v2_engine" / "templates" / "test_pdf" / "MTO"

# One BAD (90deg) and one GOOD (0deg) for comparison
_BAD_FILE = "AGCC.287-5631-KSB.MTO-0001_01-AN01_RU.pdf"
_GOOD_FILE = "AGCC.287-3340-KSB1.MTO-0001_01-AN01_RU.pdf"


def _check_coordinate_space(page, label: str):
    """Check which coordinate space get_drawings() uses."""
    rot = page.rotation % 360
    pw, ph = page.rect.width, page.rect.height
    print(f"\n{'='*70}")
    print(f"{label}: rotation={rot}  displayed_rect={page.rect}  (w={pw:.0f} h={ph:.0f})")

    if rot in (90, 270):
        print(f"  MediaBox (unrotated): width ~{ph:.0f} x height ~{pw:.0f}")

    drawings = list(page.get_drawings())
    print(f"  Total drawings: {len(drawings)}")

    all_x0 = [d["rect"].x0 for d in drawings]
    all_y0 = [d["rect"].y0 for d in drawings]
    all_x1 = [d["rect"].x1 for d in drawings]
    all_y1 = [d["rect"].y1 for d in drawings]

    print(f"  Drawing rect ranges:")
    print(f"    x: {min(all_x0):.1f} .. {max(all_x1):.1f}")
    print(f"    y: {min(all_y0):.1f} .. {max(all_y1):.1f}")

    if rot in (90, 270):
        print(f"  If DISPLAYED space: x should be 0..{pw:.0f}, y should be 0..{ph:.0f}")
        print(f"  If UNROTATED space: x should be 0..{ph:.0f}, y should be 0..{pw:.0f}")
        if max(all_x1) > ph + 1:
            print(f"  >>> x exceeds displayed height ({ph:.0f}) -> likely DISPLAYED coords")
        elif max(all_y1) > ph + 1:
            print(f"  >>> y exceeds displayed height ({ph:.0f}) -> likely UNROTATED coords")
        else:
            print(f"  >>> Ambiguous (both fit within smaller dimension)")


def _check_text_extraction(page, label: str):
    """Try extracting text from different rects to find where content lives."""
    rot = page.rotation % 360
    pw, ph = page.rect.width, page.rect.height

    print(f"\n--- Text extraction tests ({label}) ---")

    # Get all text blocks to see where text actually is
    blocks = page.get_text("blocks")
    print(f"  Total text blocks: {len(blocks)}")

    # Show first few blocks with their coordinates
    print(f"  Sample blocks (fitz coords from get_text):")
    for i, b in enumerate(blocks[:8]):
        x0, y0, x1, y1 = b[:4]
        text = b[4].strip()[:60] if len(b) > 4 else ""
        print(f"    [{i}] ({x0:.1f},{y0:.1f},{x1:.1f},{y1:.1f}) = {text!r}")

    # Try get_textbox at various positions
    print(f"\n  get_textbox at various rects (unrotated space):")
    test_rects = []
    if rot in (90, 270):
        # For 90deg: try rects in what would be the stamp area
        # In unrotated portrait (842x1191):
        # Bottom-right of landscape display = left-bottom of portrait
        # = x near 0, y near 1191
        unrot_h = pw  # 1191 for a landscape-displayed page
        unrot_w = ph  # 842

        test_rects = [
            ("bottom-right stamp area (unrot)", fitz.Rect(0, unrot_h - 300, 200, unrot_h)),
            ("center page (unrot)", fitz.Rect(unrot_w/3, unrot_h/3, 2*unrot_w/3, 2*unrot_h/3)),
            ("top-left (unrot)", fitz.Rect(0, 0, 200, 200)),
            ("full page (unrot)", fitz.Rect(0, 0, unrot_w, unrot_h)),
        ]
    else:
        test_rects = [
            ("bottom-right stamp area", fitz.Rect(pw - 400, ph - 300, pw, ph)),
            ("center page", fitz.Rect(pw/3, ph/3, 2*pw/3, 2*ph/3)),
            ("full page", fitz.Rect(0, 0, pw, ph)),
        ]

    for desc, r in test_rects:
        try:
            txt = page.get_textbox(r).strip()[:80]
        except Exception as e:
            txt = f"<error: {e}>"
        print(f"    {desc:40s} {r} -> {txt!r}")

    # For 90deg pages, also try the derotation matrix approach
    if rot in (90, 270):
        print(f"\n  page.derotation_matrix = {page.derotation_matrix}")
        print(f"  page.rotation_matrix = {page.rotation_matrix}")

        # Try displayed-space rect -> transform to unrotated -> get_textbox
        # The stamp area in displayed space is bottom-right
        disp_stamp = fitz.Rect(pw - 400, ph - 200, pw, ph)
        # Convert to unrotated using derotation matrix
        unrot_stamp = disp_stamp * page.derotation_matrix
        print(f"\n  Displayed stamp area: {disp_stamp}")
        print(f"  After derotation_matrix: {unrot_stamp}")
        try:
            txt = page.get_textbox(unrot_stamp).strip()[:80]
        except Exception as e:
            txt = f"<error: {e}>"
        print(f"  get_textbox result: {txt!r}")

        # Also try our fitz_displayed_to_unrotated (page matrices)
        from pdf_parsing_v2_engine.coord_transform import fitz_displayed_to_unrotated
        our_unrot = fitz_displayed_to_unrotated(disp_stamp, page)
        print(f"  Our fitz_displayed_to_unrotated: {our_unrot}")
        try:
            txt = page.get_textbox(our_unrot).strip()[:80]
        except Exception as e:
            txt = f"<error: {e}>"
        print(f"  get_textbox result: {txt!r}")

        # Compare the two unrotated rects
        print(f"\n  derotation vs our transform:")
        print(f"    derotation: ({unrot_stamp.x0:.1f},{unrot_stamp.y0:.1f},{unrot_stamp.x1:.1f},{unrot_stamp.y1:.1f})")
        print(f"    ours:       ({our_unrot.x0:.1f},{our_unrot.y0:.1f},{our_unrot.x1:.1f},{our_unrot.y1:.1f})")


def _check_find_tables_space(page, label: str):
    """Check what coordinate space find_tables returns cells in."""
    rot = page.rotation % 360
    pw, ph = page.rect.width, page.rect.height

    print(f"\n--- find_tables coordinate space ({label}) ---")

    try:
        tables = page.find_tables()
    except Exception as e:
        print(f"  find_tables error: {e}")
        return

    all_tables = getattr(tables, "tables", tables)
    all_cells = []
    for t in all_tables:
        for c in getattr(t, "cells", []):
            if c and len(c) == 4:
                all_cells.append(c)

    if not all_cells:
        print("  No cells found")
        return

    xs = [c[0] for c in all_cells] + [c[2] for c in all_cells]
    ys = [c[1] for c in all_cells] + [c[3] for c in all_cells]

    print(f"  Total cells: {len(all_cells)}")
    print(f"  Cell x range: {min(xs):.1f} .. {max(xs):.1f}")
    print(f"  Cell y range: {min(ys):.1f} .. {max(ys):.1f}")
    if rot in (90, 270):
        print(f"  If DISPLAYED: x in 0..{pw:.0f}, y in 0..{ph:.0f}")
        print(f"  If UNROTATED: x in 0..{ph:.0f}, y in 0..{pw:.0f}")
        if max(xs) > ph + 1:
            print(f"  >>> x exceeds {ph:.0f} -> DISPLAYED coords")
        if max(ys) > ph + 1:
            print(f"  >>> y exceeds {ph:.0f} -> UNROTATED coords")


def main():
    mto_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else _DEFAULT_MTO

    bad_path = mto_dir / _BAD_FILE
    good_path = mto_dir / _GOOD_FILE

    for path, label in [(good_path, f"GOOD (0deg) {_GOOD_FILE}"),
                        (bad_path, f"BAD (90deg) {_BAD_FILE}")]:
        if not path.exists():
            print(f"SKIP: {path} not found")
            continue
        doc = fitz.open(str(path))
        try:
            page = doc[0]
            _check_coordinate_space(page, label)
            _check_find_tables_space(page, label)
            _check_text_extraction(page, label)
        finally:
            doc.close()


if __name__ == "__main__":
    main()
