"""
Standalone MTO frame probe (does not touch pipeline).

Compares production find_frame(debug) with toy candidates on the same filtered
drawing rects as frame_detector.

Run from repo root::

  set PYTHONPATH=.
  python -m pdf_parsing_v2_engine.tools.frame_mto_probe
  python -m pdf_parsing_v2_engine.tools.frame_mto_probe --try-remove-rotation

Findings on the labeled AGCC MTO set (see run output):

- **Union AABB** of all filtered drawing bboxes in pdfminer matches **find_frame**
  within ~0-2 pt (column ``un``). So the stitched max-span heuristic agrees with
  the global envelope of filtered ``get_drawings()`` rects here; the issue is
  not "stitch vs union".
- **m20/m15** are usually empty: no single rect spans 15-20% of the page in
  both dimensions (frame is built from thin strokes).
- **set_rotation(0)** leaves composite/max ratio and same_src unchanged; metadata
  normalization alone does not repair the heuristic.
- Next place to improve real "bad" pages is likely **downstream** (grid /
  coords / stamp) or **filtering which drawings** enter the envelope, not a
  one-line swap to max-area rect.
"""

from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path
from typing import Callable

import fitz

from pdf_parsing_v2_engine.coord_transform import (
    fitz_rect_transform_by_matrix,
    unrotated_fitz_rect_to_pdfminer_bbox,
)
from pdf_parsing_v2_engine.frame_detector import SCALE, _MIN_FRAME_FRACTION, find_frame
from pdf_parsing_v2_engine.models import FrameInfo

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_DEFAULT_MTO = _PROJECT_ROOT / "pdf_parsing_v2_engine" / "templates" / "test_pdf" / "MTO"

_BAD = (
    "AGCC.287-4000-KSB.MTO-0001_0_RU.pdf",
    "AGCC.287-5631-KSB.MTO-0001_01-AN01_RU.pdf",
    "AGCC.287-7421-SOS.1.MTO-0001_01_RU.pdf",
    "AGCC.287-8525-SOS.MTO-0001_01_RU.pdf",
)
_GOOD = (
    "AGCC.287-7570-SOT.MTO-0001_01_RU.pdf",
    "AGCC.287-8441-SKUD.MTO-0001_04_RU.pdf",
    "AGCC.287-8630-KSB3.MTO-0001_02-AN01_RU.pdf",
)


def _bbox_area_pm(fr: FrameInfo) -> float:
    return max(0.0, fr.x1 - fr.x0) * max(0.0, fr.y1 - fr.y0)


def _iter_filtered_rects(page: fitz.Page) -> list[fitz.Rect]:
    pw, ph = page.rect.width, page.rect.height
    rot = page.rotation % 360
    out: list[fitz.Rect] = []
    for d in page.get_drawings():
        r = fitz.Rect(d["rect"])
        if rot in (90, 270):
            rd = fitz_rect_transform_by_matrix(r, page.rotation_matrix)
        else:
            rd = r
        if rd.x1 < (pw - SCALE) and rd.y0 > SCALE:
            out.append(fitz.Rect(r))
    return out


def _frame_from_pm(
    x0: float, y0: float, x1: float, y1: float, pw: float, ph: float, rot: int
) -> FrameInfo:
    return FrameInfo(
        x0=x0,
        y0=y0,
        x1=x1,
        y1=y1,
        page_width=pw,
        page_height=ph,
        rotation=rot,
        border_left_mm=(pw - x1) / SCALE,
        border_bottom_mm=y0 / SCALE,
        border_top_mm=(ph - y1) / SCALE,
    )


def _pm_boxes(
    page: fitz.Page,
) -> list[tuple[fitz.Rect, float, float, float, float, float, float, float]]:
    """Each tuple: fitz_rect, x0, y0, x1, y1, w, h, area (pdfminer)."""
    rows: list[tuple[fitz.Rect, float, float, float, float, float, float, float]] = []
    for fr in _iter_filtered_rects(page):
        x0, y0, x1, y1 = unrotated_fitz_rect_to_pdfminer_bbox(fr, page)
        w = max(0.0, x1 - x0)
        h = max(0.0, y1 - y0)
        a = w * h
        rows.append((fr, x0, y0, x1, y1, w, h, a))
    return rows


def best_max_area_if(
    page: fitz.Page,
    pred: Callable[[float, float, float, float], bool],
) -> tuple[FrameInfo | None, float]:
    """Max pdfminer area among filtered rects satisfying pred(w, h, pw, ph)."""
    pw, ph = page.rect.width, page.rect.height
    rot = page.rotation % 360
    best_a = -1.0
    best_box: tuple[float, float, float, float] | None = None
    for _fr, x0, y0, x1, y1, w, h, a in _pm_boxes(page):
        if not pred(w, h, pw, ph):
            continue
        if a > best_a:
            best_a, best_box = a, (x0, y0, x1, y1)
    if best_box is None:
        return None, 0.0
    x0, y0, x1, y1 = best_box
    return _frame_from_pm(x0, y0, x1, y1, pw, ph, rot), best_a


def union_bbox_filtered_rects(page: fitz.Page) -> FrameInfo | None:
    """Axis-aligned union in pdfminer space of every filtered drawing bbox."""
    pw, ph = page.rect.width, page.rect.height
    rot = page.rotation % 360
    boxes = _pm_boxes(page)
    if not boxes:
        return None
    x0 = min(t[1] for t in boxes)
    y0 = min(t[2] for t in boxes)
    x1 = max(t[3] for t in boxes)
    y1 = max(t[4] for t in boxes)
    return _frame_from_pm(x0, y0, x1, y1, pw, ph, rot)


def max_area_single_rect_frame(page: fitz.Page) -> tuple[FrameInfo | None, float]:
    pw, ph = page.rect.width, page.rect.height
    rot = page.rotation % 360
    rects = _iter_filtered_rects(page)
    if not rects:
        return None, 0.0
    best_a = -1.0
    best_r: fitz.Rect | None = None
    for fr in rects:
        x0, y0, x1, y1 = unrotated_fitz_rect_to_pdfminer_bbox(fr, page)
        a = max(0.0, x1 - x0) * max(0.0, y1 - y0)
        if a > best_a:
            best_a, best_r = a, fr
    if best_r is None:
        return None, 0.0
    x0, y0, x1, y1 = unrotated_fitz_rect_to_pdfminer_bbox(best_r, page)
    return (
        _frame_from_pm(x0, y0, x1, y1, pw, ph, rot),
        best_a,
    )


def _pred_min_fraction(frac: float) -> Callable[[float, float, float, float], bool]:
    def _inner(w: float, h: float, pw: float, ph: float) -> bool:
        return w >= frac * pw and h >= frac * ph

    return _inner


def _pred_non_thin(max_aspect: float, min_short_side_pts: float) -> Callable[..., bool]:
    def _inner(w: float, h: float, _pw: float, _ph: float) -> bool:
        if w <= 0 or h <= 0:
            return False
        short, long = (w, h) if w < h else (h, w)
        if short < min_short_side_pts:
            return False
        return (long / short) <= max_aspect

    return _inner


def _classify(name: str) -> str:
    n = name.lower()
    for b in _BAD:
        if b.lower() in n:
            return "bad"
    for g in _GOOD:
        if g.lower() in n:
            return "good"
    return "?"


def _collect_pdfs(base: Path, only: str) -> list[Path]:
    if not base.is_dir():
        return []
    pdfs = sorted(base.glob("*.pdf"))
    if only == "bad":
        pdfs = [p for p in pdfs if _classify(p.name) == "bad"]
    elif only == "good":
        pdfs = [p for p in pdfs if _classify(p.name) == "good"]
    elif only == "labeled":
        pdfs = [p for p in pdfs if _classify(p.name) in ("bad", "good")]
    return pdfs


def _corner_L2(a: FrameInfo, b: FrameInfo) -> float:
    return math.sqrt(
        (a.x0 - b.x0) ** 2
        + (a.y0 - b.y0) ** 2
        + (a.x1 - b.x1) ** 2
        + (a.y1 - b.y1) ** 2
    )


def _short_label(name: str, width: int = 18) -> str:
    stem = Path(name).stem
    return stem[:width] if len(stem) <= width else stem[: width - 1] + "~"


def run_probe(pdfs: list[Path], try_remove_rotation: bool) -> None:
    frac = _MIN_FRAME_FRACTION
    print(
        "frame_mto_probe - candidates: prod | raw_max_area | "
        f"m{int(frac*100)} | m15 | non_thin | "
        "union(all filtered bboxes in pdfminer)\n"
    )

    for path in pdfs:
        doc = fitz.open(path)
        try:
            page = doc[0]
            if try_remove_rotation:
                rm = getattr(page, "remove_rotation", None)
                if callable(rm):
                    try:
                        rm()
                    except Exception as exc:  # noqa: BLE001
                        print(f"[warn] remove_rotation: {exc}", file=sys.stderr)

            label = _classify(path.name)
            prod, dbg = find_frame(page, debug=True)
            rot = dbg["rotation_handled"] if dbg else page.rotation % 360
            mx = dbg.get("max_single_rect_area_pts2")
            comp = dbg.get("composite_area_pts2")
            ratio = (comp / mx) if mx and mx > 0 else None
            same = dbg.get("same_source_rect_for_both_axes")

            raw_max, _ = max_area_single_rect_frame(page)
            m20, _ = best_max_area_if(page, _pred_min_fraction(frac))
            m15, _ = best_max_area_if(page, _pred_min_fraction(0.15))
            nt, _ = best_max_area_if(
                page, _pred_non_thin(max_aspect=20.0, min_short_side_pts=10.0 * SCALE)
            )
            union = union_bbox_filtered_rects(page)

            def dist(x: FrameInfo | None) -> str:
                if x is None:
                    return "-"
                return f"{_corner_L2(prod, x):.0f}"

            print(
                f"{_short_label(path.name, 22):22} {label:4} rot={rot:3} "
                f"r={ratio or 0:6.0f} same={str(same)[:5]:5} | "
                f"L2 raw={dist(raw_max):>4} m20={dist(m20):>4} m15={dist(m15):>4} "
                f"nt={dist(nt):>4} un={dist(union):>4}"
            )
        finally:
            doc.close()

    print(
        "\nL2 = distance from production find_frame corners to candidate (pts).\n"
        "m20/m15 = max-area single rect with both spans >= fraction of page "
        "(no candidate => line-only art).\n"
        "union = AABB in pdfminer over all filtered drawing rects.\n"
    )
    _print_set_rotation_invariant_rows(pdfs)
    _print_conclusions()


def _print_set_rotation_invariant_rows(pdfs: list[Path]) -> None:
    print("--- set_rotation(0): ratio and same_src (mutating page) ---")
    for path in pdfs:
        doc = fitz.open(path)
        try:
            page = doc[0]
            _, d1 = find_frame(page, debug=True)
            mx1 = d1.get("max_single_rect_area_pts2") or 0.0
            r1 = (d1.get("composite_area_pts2") or 0.0) / mx1 if mx1 > 0 else 0.0
            s1 = d1.get("same_source_rect_for_both_axes")
            page.set_rotation(0)
            _, d2 = find_frame(page, debug=True)
            mx2 = d2.get("max_single_rect_area_pts2") or 0.0
            r2 = (d2.get("composite_area_pts2") or 0.0) / mx2 if mx2 > 0 else 0.0
            s2 = d2.get("same_source_rect_for_both_axes")
            inv = abs(r1 - r2) < 1e-3 and s1 == s2
            lab = _short_label(path.name, 20)
            print(f"  {lab:20} ratio {r1:.3f} -> {r2:.3f}  same {s1!s:5} -> {s2!s:5}  invariant={inv}")
        finally:
            doc.close()
    print()


def _print_conclusions() -> None:
    print(
        "Conclusions (harness only):\n"
        "  - union ~= find_frame: stitched frame matches global AABB of filtered drawings.\n"
        "  - Replacing frame with max-area or min-fraction single rect is not viable on line art.\n"
        "  - set_rotation(0) does not change ratio/same_src; need other levers for bad 90deg pages.\n"
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Probe alternative frame heuristics on MTO PDFs.")
    parser.add_argument("--dir", type=Path, default=_DEFAULT_MTO)
    parser.add_argument("--only", choices=("all", "labeled", "good", "bad"), default="labeled")
    parser.add_argument(
        "--try-remove-rotation",
        action="store_true",
        help="Call page.remove_rotation() before probing (mutates in-memory page; experimental).",
    )
    args = parser.parse_args(argv)

    base = args.dir.resolve()
    pdfs = _collect_pdfs(base, args.only)
    if not pdfs:
        print(f"No PDFs in {base}", file=sys.stderr)
        return 1

    run_probe(pdfs, try_remove_rotation=args.try_remove_rotation)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
