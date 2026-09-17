from __future__ import annotations

import ast
import json
import os
import sys
from pathlib import Path

import fitz

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


RUN_LABEL = os.environ.get("STAMP_EXTENT_RENDER_RUN_LABEL", "extent_probe_virtual_subset")
RUN_DIR = (
    REPO_ROOT
    / "pdf_parsing_v2_engine"
    / "templates"
    / "stamp_deform_bench"
    / "output"
    / "benchmark_runs"
    / RUN_LABEL
)
CASES_DIR = RUN_DIR / "cases"
OUTPUT_DIR = REPO_ROOT / "tmp" / "extent_probe_images"

CASES = [
    "AGCC.287-2869-SOS.MTO-0001_01-AN02_RU__sx101__sy105.json",
    "AGCC.287-2869-SOS.MTO-0001_01-AN02_RU__sx109__sy111.json",
    "AGCC.287-2869-SOS.MTO-0001_01-AN02_RU__sx101__sy115.json",
]

_DPI = 140
_SCALE = _DPI / 72.0


def _parse_rect(raw: str | list[float]) -> fitz.Rect:
    coords = ast.literal_eval(raw) if isinstance(raw, str) else raw
    return fitz.Rect(*[float(v) for v in coords])


def _draw_rect(
    page: fitz.Page,
    rect: fitz.Rect,
    color: tuple[float, float, float],
    width: float = 1.8,
    dashes: str | None = None,
) -> None:
    if rect.is_empty or rect.is_infinite:
        return
    scaled = fitz.Rect(rect.x0 * _SCALE, rect.y0 * _SCALE, rect.x1 * _SCALE, rect.y1 * _SCALE)
    shape = page.new_shape()
    shape.draw_rect(scaled)
    shape.finish(color=color, width=width, dashes=dashes)
    shape.commit()


def _draw_line(
    page: fitz.Page,
    p1: fitz.Point,
    p2: fitz.Point,
    color: tuple[float, float, float],
    width: float = 1.4,
    dashes: str | None = None,
) -> None:
    shape = page.new_shape()
    shape.draw_line(
        fitz.Point(p1.x * _SCALE, p1.y * _SCALE),
        fitz.Point(p2.x * _SCALE, p2.y * _SCALE),
    )
    shape.finish(color=color, width=width, dashes=dashes)
    shape.commit()


def _insert_text(
    page: fitz.Page,
    x: float,
    y: float,
    text: str,
    color: tuple[float, float, float] = (0, 0, 0),
    fontsize: float = 7.5,
) -> None:
    page.insert_text(
        fitz.Point(x, y),
        text,
        fontsize=fontsize,
        color=color,
    )


def _detected_by_id(case_payload: dict) -> dict[str, dict]:
    return {line["id"]: line for line in case_payload.get("detected_grid_lines", [])}


def _render_case(case_path: Path) -> Path:
    payload = json.loads(case_path.read_text(encoding="utf-8"))
    summary = payload["summary_row"]
    manifest = payload["manifest"]
    pdf_path = Path(manifest["output_pdf"])
    if not pdf_path.is_file():
        raise FileNotFoundError(pdf_path)

    base_bbox = _parse_rect(summary["stamp_bbox_target"] if False else payload["stamp_bbox_base_pts"])
    target_bbox = _parse_rect(summary["stamp_bbox_target"])
    proposed_bbox = fitz.Rect(
        float(summary["extent_probe_proposed_bbox_x0_pts"]),
        float(summary["extent_probe_proposed_bbox_y0_pts"]),
        float(summary["extent_probe_proposed_bbox_x1_pts"]),
        float(summary["extent_probe_proposed_bbox_y1_pts"]),
    )
    virtual_top = summary.get("extent_probe_virtual_top_virtual_pos_pts", "")
    virtual_left = summary.get("extent_probe_virtual_left_virtual_pos_pts", "")

    detected = _detected_by_id(payload)
    top_id = summary.get("extent_probe_top_id", "")
    bottom_id = summary.get("extent_probe_bottom_id", "")
    left_id = summary.get("extent_probe_left_id", "")
    right_id = summary.get("extent_probe_right_id", "")

    with fitz.open(pdf_path) as doc:
        fitz_page = doc[int(manifest["page_num"]) - 1]
        pix = fitz_page.get_pixmap(dpi=_DPI)
        tmp_pdf = fitz.open()
        tmp_page = tmp_pdf.new_page(width=pix.width, height=pix.height)
        tmp_page.insert_image(tmp_page.rect, pixmap=pix)

        # All detected lines in soft colors.
        for line in payload.get("detected_grid_lines", []):
            if line["orientation"] == "h":
                _draw_line(
                    tmp_page,
                    fitz.Point(line["span_lo_pts"], line["pos_pts"]),
                    fitz.Point(line["span_hi_pts"], line["pos_pts"]),
                    color=(0.45, 0.8, 1.0),
                    width=1.1,
                )
            else:
                _draw_line(
                    tmp_page,
                    fitz.Point(line["pos_pts"], line["span_lo_pts"]),
                    fitz.Point(line["pos_pts"], line["span_hi_pts"]),
                    color=(0.35, 0.55, 1.0),
                    width=1.1,
                )

        # Selected candidates.
        candidate_colors = {
            top_id: (1.0, 0.0, 0.8),
            bottom_id: (0.0, 0.75, 0.25),
            left_id: (1.0, 0.2, 0.2),
            right_id: (1.0, 0.55, 0.0),
        }
        for line_id, color in candidate_colors.items():
            line = detected.get(line_id)
            if not line:
                continue
            if line["orientation"] == "h":
                _draw_line(
                    tmp_page,
                    fitz.Point(line["span_lo_pts"], line["pos_pts"]),
                    fitz.Point(line["span_hi_pts"], line["pos_pts"]),
                    color=color,
                    width=2.6,
                )
            else:
                _draw_line(
                    tmp_page,
                    fitz.Point(line["pos_pts"], line["span_lo_pts"]),
                    fitz.Point(line["pos_pts"], line["span_hi_pts"]),
                    color=color,
                    width=2.6,
                )

        # Rectangles.
        _draw_rect(tmp_page, target_bbox, color=(1.0, 0.45, 0.0), width=2.3)
        _draw_rect(tmp_page, base_bbox, color=(0.6, 0.15, 0.85), width=2.0, dashes="[4 2] 0")
        _draw_rect(tmp_page, proposed_bbox, color=(0.0, 0.75, 0.15), width=2.3, dashes="[8 3] 0")

        # Explicit inferred edges from proposed bbox for easier visual reading.
        _draw_line(
            tmp_page,
            fitz.Point(proposed_bbox.x0, proposed_bbox.y0),
            fitz.Point(proposed_bbox.x1, proposed_bbox.y0),
            color=(0.0, 0.75, 0.15),
            width=2.3,
        )
        _draw_line(
            tmp_page,
            fitz.Point(proposed_bbox.x0, proposed_bbox.y0),
            fitz.Point(proposed_bbox.x0, proposed_bbox.y1),
            color=(0.0, 0.75, 0.15),
            width=2.3,
        )
        if virtual_top != "":
            virtual_top_y = float(virtual_top)
            _draw_line(
                tmp_page,
                fitz.Point(base_bbox.x0 - 18.0, virtual_top_y),
                fitz.Point(base_bbox.x1 + 18.0, virtual_top_y),
                color=(0.9, 0.0, 0.0),
                width=2.7,
                dashes="[10 4] 0",
            )
        if virtual_left != "":
            virtual_left_x = float(virtual_left)
            _draw_line(
                tmp_page,
                fitz.Point(virtual_left_x, base_bbox.y0 - 18.0),
                fitz.Point(virtual_left_x, base_bbox.y1 + 18.0),
                color=(0.8, 0.0, 0.0),
                width=2.7,
                dashes="[10 4] 0",
            )

        # Legend / metrics.
        lines = [
            pdf_path.name,
            "orange = manifest target bbox",
            "purple dashed = current base bbox",
            "green dashed = extent probe proposed bbox",
            "red dashed = virtual outer top / left",
            "magenta = chosen top, green = chosen bottom",
            "red = chosen left candidate, orange = chosen right candidate",
            f"target:   x0={target_bbox.x0:.1f} y0={target_bbox.y0:.1f}",
            f"proposed: x0={proposed_bbox.x0:.1f} y0={proposed_bbox.y0:.1f}",
            f"base:     x0={base_bbox.x0:.1f} y0={base_bbox.y0:.1f}",
            f"virtual:  x0={virtual_left or '-'} y0={virtual_top or '-'}",
            f"verdict={summary['alignment_verdict']} skipped={summary['template_skipped']} avg={summary['avg_match_score']}",
            f"top={top_id or '-'} left={left_id or '-'}",
            f"bottom={bottom_id or '-'} right={right_id or '-'}",
            (
                "virtual_top="
                f"{summary.get('extent_probe_virtual_top_source_detected_id', '-')}"
                f" -> {summary.get('extent_probe_virtual_top_matched_template_id', '-')}"
                f" skip={summary.get('extent_probe_virtual_top_skip_count', '-')}"
            ),
            (
                "virtual_left="
                f"{summary.get('extent_probe_virtual_left_source_detected_id', '-')}"
                f" -> {summary.get('extent_probe_virtual_left_matched_template_id', '-')}"
                f" skip={summary.get('extent_probe_virtual_left_skip_count', '-')}"
            ),
        ]
        y = 14.0
        for line in lines:
            _insert_text(tmp_page, 12.0, y, line)
            y += 13.0

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        out_path = OUTPUT_DIR / f"{pdf_path.stem}__extent_probe.png"
        out_pix = tmp_page.get_pixmap()
        out_pix.save(str(out_path))
        tmp_pdf.close()
        return out_path


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    generated = []
    for name in CASES:
        generated.append(_render_case(CASES_DIR / name))
    print("Generated extent probe images:")
    for path in generated:
        print(path)


if __name__ == "__main__":
    main()
