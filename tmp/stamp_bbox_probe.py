from __future__ import annotations

import ast
import csv
import json
import sys
from pathlib import Path

import fitz

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from pdf_parsing_v2_engine.coord_transform import pdfminer_to_fitz
from pdf_parsing_v2_engine.find_tables_settings import build_find_tables_clip
from pdf_parsing_v2_engine.frame_detector import find_frame
from pdf_parsing_v2_engine.grid_matcher import (
    _build_detected_grid_lines,
    _compute_border_merge_dist,
    _filter_outside_stamp,
    _merge_border_pairs,
    _template_stamp_bbox,
    get_detected_stamp_cells,
)
from pdf_parsing_v2_engine.models import StampTemplate
from pdf_parsing_v2.v2_config import load_v2_config


DATASET_DIR = REPO_ROOT / "pdf_parsing_v2_engine" / "templates" / "stamp_deform_bench"
GENERATED_DIR = DATASET_DIR / "output" / "generated"
MANIFEST_CSV = DATASET_DIR / "output" / "manifests" / "manifest.csv"
TEMPLATE_PATH = DATASET_DIR / "template" / "mto_page1.json"
OUTPUT_DIR = REPO_ROOT / "tmp" / "stamp_bbox_probe"

CASES = [
    "AGCC.287-2869-SOS.MTO-0001_01-AN02_RU__sx101__sy105.pdf",
    "AGCC.287-2869-SOS.MTO-0001_01-AN02_RU__sx109__sy111.pdf",
    "AGCC.287-2869-SOS.MTO-0001_01-AN02_RU__sx101__sy115.pdf",
]

_DPI = 130
_SCALE = _DPI / 72.0


def _parse_rect(value: str) -> fitz.Rect:
    coords = ast.literal_eval(value)
    return fitz.Rect(*[float(v) for v in coords])


def _draw_rect(page: fitz.Page, rect: fitz.Rect, color: tuple[float, float, float], width: float, dashes: str | None = None) -> None:
    r = fitz.Rect(rect.x0 * _SCALE, rect.y0 * _SCALE, rect.x1 * _SCALE, rect.y1 * _SCALE)
    if r.is_empty or r.is_infinite:
        return
    shape = page.new_shape()
    shape.draw_rect(r)
    shape.finish(color=color, width=width, dashes=dashes)
    shape.commit()


def _draw_line(page: fitz.Page, p1: fitz.Point, p2: fitz.Point, color: tuple[float, float, float], width: float, dashes: str | None = None) -> None:
    shape = page.new_shape()
    shape.draw_line(
        fitz.Point(p1.x * _SCALE, p1.y * _SCALE),
        fitz.Point(p2.x * _SCALE, p2.y * _SCALE),
    )
    shape.finish(color=color, width=width, dashes=dashes)
    shape.commit()


def _insert_text(page: fitz.Page, point: fitz.Point, text: str, color: tuple[float, float, float], fontsize: float = 7) -> None:
    page.insert_text(
        fitz.Point(point.x * _SCALE, point.y * _SCALE),
        text,
        fontsize=fontsize,
        color=color,
    )


def _load_manifest_map() -> dict[str, dict[str, str]]:
    with MANIFEST_CSV.open("r", encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.DictReader(fh))
    return {Path(row["output_pdf"]).name: row for row in rows}


def _line_key(line) -> tuple[float, float, float]:
    return (round(line.pos, 3), round(line.span_lo, 3), round(line.span_hi, 3))


def _render_case(
    *,
    pdf_path: Path,
    row: dict[str, str],
    template: StampTemplate,
    cfg: dict,
) -> dict:
    with fitz.open(pdf_path) as doc:
        fitz_page = doc[0]
        t_union = template if template.frame_mode == "drawing_union" else None
        frame, _ = find_frame(
            fitz_page,
            frame_mode=template.frame_mode,
            template=t_union,
            cfg=cfg,
        )
        stamp_bbox = _template_stamp_bbox(template, frame)
        if stamp_bbox is None:
            raise RuntimeError("Template stamp bbox is empty")
        clip = build_find_tables_clip(fitz_page, stamp_bbox)
        target_bbox = _parse_rect(row["stamp_bbox_target"])

        detected_cells = get_detected_stamp_cells(
            fitz_page=fitz_page,
            stamp_bbox=stamp_bbox,
            tolerance_mm=template.grid_tolerance_detected_mm,
            cfg=cfg,
            template=template,
        )
        h_raw, v_raw = _build_detected_grid_lines(detected_cells, template.grid_tolerance_detected_mm)
        merge_dist = _compute_border_merge_dist(detected_cells)
        h_merged = _merge_border_pairs(h_raw, merge_dist)
        v_merged = _merge_border_pairs(v_raw, merge_dist)
        margin = template.snap_max_distance_mm * 72.0 / 25.4
        h_filtered = _filter_outside_stamp(h_merged, stamp_bbox.y0, stamp_bbox.y1, margin)
        v_filtered = _filter_outside_stamp(v_merged, stamp_bbox.x0, stamp_bbox.x1, margin)

        h_filtered_keys = {_line_key(line) for line in h_filtered}
        v_filtered_keys = {_line_key(line) for line in v_filtered}
        h_dropped = [line for line in h_merged if _line_key(line) not in h_filtered_keys]
        v_dropped = [line for line in v_merged if _line_key(line) not in v_filtered_keys]

        pix = fitz_page.get_pixmap(dpi=_DPI)
        tmp_pdf = fitz.open()
        tmp_page = tmp_pdf.new_page(width=pix.width, height=pix.height)
        tmp_page.insert_image(tmp_page.rect, pixmap=pix)

        frame_rect = pdfminer_to_fitz((frame.x0, frame.y0, frame.x1, frame.y1), frame.page_height, frame.rotation)
        _draw_rect(tmp_page, frame_rect, (0.1, 0.7, 0.1), 2.0)
        _draw_rect(tmp_page, target_bbox, (1.0, 0.45, 0.0), 2.2)
        _draw_rect(tmp_page, stamp_bbox, (0.65, 0.2, 0.85), 2.0, dashes="[4 2] 0")
        if clip is not None:
            _draw_rect(tmp_page, clip, (0.0, 0.55, 1.0), 2.0, dashes="[8 3] 0")

        for line in h_merged:
            color = (0.9, 0.1, 0.1) if line in h_dropped else (0.05, 0.8, 0.9)
            width = 2.2 if line in h_dropped else 1.6
            _draw_line(
                tmp_page,
                fitz.Point(line.span_lo, line.pos),
                fitz.Point(line.span_hi, line.pos),
                color,
                width,
            )
        for line in v_merged:
            color = (0.9, 0.1, 0.1) if line in v_dropped else (0.2, 0.45, 1.0)
            width = 2.2 if line in v_dropped else 1.6
            _draw_line(
                tmp_page,
                fitz.Point(line.pos, line.span_lo),
                fitz.Point(line.pos, line.span_hi),
                color,
                width,
            )

        info_lines = [
            pdf_path.name,
            "orange = manifest stamp_bbox_target",
            "purple dashed = pipeline stamp_bbox/search_bbox",
            "blue dashed = find_tables clip x1.5",
            "cyan/blue = merged lines kept after stamp filter",
            "red = merged lines dropped by _filter_outside_stamp",
            f"stamp_bbox_y = {stamp_bbox.y0:.1f}..{stamp_bbox.y1:.1f}",
            f"target_bbox_y = {target_bbox.y0:.1f}..{target_bbox.y1:.1f}",
            f"h_raw={len(h_raw)} h_merged={len(h_merged)} h_filtered={len(h_filtered)} h_dropped={len(h_dropped)}",
            f"v_raw={len(v_raw)} v_merged={len(v_merged)} v_filtered={len(v_filtered)} v_dropped={len(v_dropped)}",
        ]
        y = 12.0
        for line in info_lines:
            _insert_text(tmp_page, fitz.Point(10.0 / _SCALE, y / _SCALE), line, (0, 0, 0), fontsize=8)
            y += 13.0

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        png_path = OUTPUT_DIR / f"{pdf_path.stem}__bbox_probe.png"
        final_pix = tmp_page.get_pixmap()
        final_pix.save(str(png_path))
        tmp_pdf.close()

        return {
            "pdf": str(pdf_path),
            "png": str(png_path),
            "stamp_bbox_pts": [round(stamp_bbox.x0, 3), round(stamp_bbox.y0, 3), round(stamp_bbox.x1, 3), round(stamp_bbox.y1, 3)],
            "target_bbox_pts": [round(target_bbox.x0, 3), round(target_bbox.y0, 3), round(target_bbox.x1, 3), round(target_bbox.y1, 3)],
            "clip_pts": (
                [round(clip.x0, 3), round(clip.y0, 3), round(clip.x1, 3), round(clip.y1, 3)]
                if clip is not None
                else None
            ),
            "h_counts": {
                "raw": len(h_raw),
                "merged": len(h_merged),
                "filtered": len(h_filtered),
                "dropped": len(h_dropped),
            },
            "v_counts": {
                "raw": len(v_raw),
                "merged": len(v_merged),
                "filtered": len(v_filtered),
                "dropped": len(v_dropped),
            },
            "h_dropped_positions": [round(line.pos, 3) for line in h_dropped],
            "v_dropped_positions": [round(line.pos, 3) for line in v_dropped],
        }


def main() -> None:
    template = StampTemplate.from_json(str(TEMPLATE_PATH))
    cfg = dict(load_v2_config())
    manifest_map = _load_manifest_map()

    results = []
    for name in CASES:
        pdf_path = GENERATED_DIR / name
        row = manifest_map[name]
        results.append(_render_case(pdf_path=pdf_path, row=row, template=template, cfg=cfg))

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    report_path = OUTPUT_DIR / "report.json"
    report_path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Saved report: {report_path}")
    for result in results:
        print(result["png"])


if __name__ == "__main__":
    main()
