"""Local benchmark runner for grid alignment on generated deform PDFs.

This script is meant to be edited directly and launched as a standalone file:

    python pdf_parsing_v2_engine/templates/stamp_deform_bench/run_alignment_benchmark.py

It benchmarks the current grid alignment algorithm on
``stamp_deform_bench/output/generated`` and writes a self-contained run folder
with CSV/JSON summaries plus per-file alignment logs.
"""

from __future__ import annotations

import csv
import json
import os
import sys
import traceback
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import fitz

DATASET_DIR = Path(__file__).resolve().parent
REPO_ROOT = DATASET_DIR.parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from pdf_parsing_v2.v2_config import load_v2_config
from pdf_parsing_v2_engine.coord_transform import SCALE, grid_line_to_fitz_pts
from pdf_parsing_v2_engine.grid_detected_prefilter import prefilter_detected_cells_for_stamp_table
from pdf_parsing_v2_engine.frame_detector import find_frame
from pdf_parsing_v2_engine.grid_matcher import (
    AlignmentLineDiag,
    AlignmentReport,
    AlignmentSummary,
    CellBbox,
    DetectedGridLineInfo,
    _DetectedGridLine,
    _align_grid_lines_ordered,
    _anchor_bracket_ids,
    _assign_detected_ids,
    _build_detected_grid_lines,
    _compute_border_merge_dist,
    _compute_global_transform,
    _EPS,
    _field_in_stamp_geometry_union,
    _filter_outside_stamp,
    _insert_anchor,
    _merge_border_pairs,
    _progressive_interp,
    _template_field_bbox,
    _template_stamp_bbox,
    _virtual_merge_template_lines,
    get_detected_stamp_cells,
)
from pdf_parsing_v2_engine.models import StampTemplate


def _pick_single_file(folder: Path, pattern: str) -> Path:
    """Return the single matching file from a folder."""
    matches = sorted(path for path in folder.glob(pattern) if path.is_file())
    if len(matches) != 1:
        pretty = ", ".join(path.name for path in matches[:10]) or "<none>"
        raise SystemExit(
            f"Expected exactly one file matching {pattern!r} in {folder}, got {len(matches)}: {pretty}"
        )
    return matches[0]


# ---------------------------------------------------------------------------
# Dataset paths
# ---------------------------------------------------------------------------
TEMPLATE_PATH = _pick_single_file(DATASET_DIR / "template", "*.json")
GENERATED_DIR = DATASET_DIR / "output" / "generated"
MANIFEST_CSV = DATASET_DIR / "output" / "manifests" / "manifest.csv"


# ---------------------------------------------------------------------------
# Benchmark run parameters
# ---------------------------------------------------------------------------
RUN_LABEL = os.environ.get("STAMP_ALIGN_RUN_LABEL", "baseline_current")
OUTPUT_ROOT = DATASET_DIR / "output" / "benchmark_runs"
EXPERIMENT_NAME = os.environ.get("STAMP_ALIGN_EXPERIMENT", "baseline")

_ONLY_OUTPUT_PDFS_ENV = os.environ.get("STAMP_ALIGN_ONLY", "").strip()
# Keep empty to process every manifest row.
ONLY_OUTPUT_PDFS: list[str] = [
    item.strip()
    for item in _ONLY_OUTPUT_PDFS_ENV.split(";")
    if item.strip()
]

# Use None to process all files.
LIMIT: int | None = None

SAVE_ALIGNMENT_LOGS = True
SAVE_CASE_JSON = True


# ---------------------------------------------------------------------------
# Optional engine overrides
# ---------------------------------------------------------------------------
CFG_OVERRIDES: dict[str, Any] = {
    # Example:
    # "find_frame_border_band_left_mm": 20.0,
}


_MAX_CELL_AREA_FRACTION = 0.5

_EXPERIMENT_PRESETS: dict[str, dict[str, Any]] = {
    "baseline": {
        "bbox_overflow_x_frac": 0.0,
        "bbox_overflow_y_frac": 0.0,
        "use_expanded_bbox_for_d_all": False,
        "scale_threshold": 0.05,
        "alignment_strategy": "baseline",
    },
    "bbox_overflow": {
        "bbox_overflow_x_frac": 0.06,
        "bbox_overflow_y_frac": 0.06,
        "use_expanded_bbox_for_d_all": False,
        "scale_threshold": 0.05,
        "alignment_strategy": "baseline",
    },
    "bbox_overflow_dall": {
        "bbox_overflow_x_frac": 0.06,
        "bbox_overflow_y_frac": 0.06,
        "use_expanded_bbox_for_d_all": True,
        "scale_threshold": 0.05,
        "alignment_strategy": "baseline",
    },
    "bbox_overflow_dall_scale3": {
        "bbox_overflow_x_frac": 0.06,
        "bbox_overflow_y_frac": 0.06,
        "use_expanded_bbox_for_d_all": True,
        "scale_threshold": 0.03,
        "alignment_strategy": "baseline",
    },
    "bbox_overflow_y20": {
        "bbox_overflow_x_frac": 0.08,
        "bbox_overflow_y_frac": 0.20,
        "use_expanded_bbox_for_d_all": False,
        "scale_threshold": 0.05,
        "alignment_strategy": "baseline",
    },
    "proposed_bbox_y20": {
        "bbox_overflow_x_frac": 0.08,
        "bbox_overflow_y_frac": 0.20,
        "use_expanded_bbox_for_d_all": False,
        "scale_threshold": 0.05,
        "alignment_strategy": "experimental",
        "use_proposed_bbox_for_d_all": True,
        "use_proposed_bbox_for_filter": True,
    },
    "proposed_bbox_y20_scale45": {
        "bbox_overflow_x_frac": 0.08,
        "bbox_overflow_y_frac": 0.20,
        "use_expanded_bbox_for_d_all": False,
        "scale_threshold": 0.045,
        "alignment_strategy": "experimental",
        "use_proposed_bbox_for_d_all": True,
        "use_proposed_bbox_for_filter": True,
    },
    "proposed_bbox_y20_scale50_ge": {
        "bbox_overflow_x_frac": 0.08,
        "bbox_overflow_y_frac": 0.20,
        "use_expanded_bbox_for_d_all": False,
        "scale_threshold": 0.05,
        "alignment_strategy": "experimental",
        "use_proposed_bbox_for_d_all": True,
        "use_proposed_bbox_for_filter": True,
        "inclusive_scale_threshold": True,
    },
    "ordered_gap_score": {
        "bbox_overflow_x_frac": 0.0,
        "bbox_overflow_y_frac": 0.0,
        "use_expanded_bbox_for_d_all": False,
        "scale_threshold": 0.05,
        "alignment_strategy": "experimental",
        "gap_score_weight_h": 0.30,
        "lookahead_skip_h": False,
        "lookahead_min_score_gain": 0.05,
    },
    "ordered_gap_lookahead": {
        "bbox_overflow_x_frac": 0.0,
        "bbox_overflow_y_frac": 0.0,
        "use_expanded_bbox_for_d_all": False,
        "scale_threshold": 0.05,
        "alignment_strategy": "experimental",
        "gap_score_weight_h": 0.30,
        "lookahead_skip_h": True,
        "lookahead_min_score_gain": 0.05,
    },
}


@dataclass
class _CandidateScore:
    """One scored detected-line candidate inside the search window."""

    dl: _DetectedGridLine
    index: int
    score: float
    dist: float
    pos_score: float
    span_ratio: float
    gap_score: float


def _score_alignment_candidate(
    *,
    dl: _DetectedGridLine,
    di: int,
    pos_interp: float,
    span_t: float,
    full_span_threshold: float,
    snap_max_pts: float,
    prev_matched_pos: float,
    gap_score_weight: float,
) -> _CandidateScore:
    """Score one candidate with optional local-spacing consistency."""
    dist = abs(dl.pos - pos_interp)
    pos_score = max(0.0, 1.0 - dist / max(snap_max_pts * 2, 1.0))

    if span_t < _EPS or dl.span_length < _EPS:
        span_ratio = 0.0
        base_score = 0.3 * pos_score
    else:
        span_ratio = min(span_t, dl.span_length) / max(span_t, dl.span_length)
        both_full = span_t > full_span_threshold and dl.span_length > full_span_threshold
        category_bonus = 0.2 if both_full else 0.0
        base_score = 0.4 * span_ratio + 0.4 * pos_score + category_bonus

    if prev_matched_pos <= -1e17:
        gap_score = 1.0
    else:
        expected_gap = max(pos_interp - prev_matched_pos, 0.5)
        actual_gap = dl.pos - prev_matched_pos
        gap_tol = max(snap_max_pts, expected_gap * 0.75)
        gap_score = max(0.0, 1.0 - abs(actual_gap - expected_gap) / max(gap_tol, 1.0))

    return _CandidateScore(
        dl=dl,
        index=di,
        score=base_score + gap_score_weight * gap_score,
        dist=dist,
        pos_score=pos_score,
        span_ratio=span_ratio,
        gap_score=gap_score,
    )


def _compute_search_window(
    *,
    gi: int,
    groups: list[list[tuple[str, str, float, float, float, str | None, float]]],
    anchors: list[tuple[float, float]],
    pos_interp: float,
    prev_matched_pos: float,
    snap_max_pts: float,
    recent_deltas: list[float],
) -> tuple[float, float]:
    """Mirror the current adaptive search-window logic for one group."""
    adaptive_margin = max(
        snap_max_pts,
        1.5 * max(recent_deltas[-5:], default=0.0),
    )
    if gi + 1 < len(groups):
        next_pos_t = groups[gi + 1][0][2]
        next_interp = _progressive_interp(next_pos_t, anchors)
        next_interp = max(next_interp, pos_interp + 0.5)
        window_hi = (pos_interp + next_interp) * 0.5 + adaptive_margin
    else:
        window_hi = pos_interp + adaptive_margin * 3
    window_lo = prev_matched_pos + 0.5 if prev_matched_pos > -1e17 else -1e18
    return window_lo, window_hi


def _collect_alignment_candidates(
    *,
    pool: list[_DetectedGridLine],
    start_index: int,
    window_lo: float,
    window_hi: float,
    pos_interp: float,
    span_t: float,
    full_span_threshold: float,
    snap_max_pts: float,
    prev_matched_pos: float,
    gap_score_weight: float,
) -> list[_CandidateScore]:
    """Return scored candidates for one group inside the current window."""
    candidates: list[_CandidateScore] = []
    for di in range(start_index, len(pool)):
        dl = pool[di]
        if dl.pos < window_lo:
            continue
        if dl.pos > window_hi:
            break
        candidates.append(
            _score_alignment_candidate(
                dl=dl,
                di=di,
                pos_interp=pos_interp,
                span_t=span_t,
                full_span_threshold=full_span_threshold,
                snap_max_pts=snap_max_pts,
                prev_matched_pos=prev_matched_pos,
                gap_score_weight=gap_score_weight,
            )
        )
    return candidates


def _should_force_interp_by_lookahead(
    *,
    orient: str,
    best: _CandidateScore | None,
    gi: int,
    groups: list[list[tuple[str, str, float, float, float, str | None, float]]],
    anchors: list[tuple[float, float]],
    pool: list[_DetectedGridLine],
    det_ptr: int,
    pos_interp: float,
    prev_matched_pos: float,
    recent_deltas: list[float],
    snap_max_pts: float,
    full_span_threshold: float,
    gap_score_weight: float,
    lookahead_min_score_gain: float,
) -> bool:
    """Decide whether the current group should be interpolated to keep index sync."""
    if orient != "h" or best is None or gi + 1 >= len(groups):
        return False
    if prev_matched_pos <= -1e17:
        return False
    if best.dl.pos <= pos_interp:
        return False

    expected_gap = max(pos_interp - prev_matched_pos, 0.5)
    positive_delta = best.dl.pos - pos_interp
    if positive_delta < max(snap_max_pts * 0.8, expected_gap * 0.6):
        return False
    if best.gap_score >= 0.55 and best.score >= 0.72:
        return False

    next_group = groups[gi + 1]
    next_pos_t = next_group[0][2]
    next_span_t = sum(t[6] for t in next_group) * SCALE
    next_interp_skip = max(_progressive_interp(next_pos_t, anchors), prev_matched_pos + 0.5)
    skip_window_lo, skip_window_hi = _compute_search_window(
        gi=gi + 1,
        groups=groups,
        anchors=anchors,
        pos_interp=next_interp_skip,
        prev_matched_pos=prev_matched_pos,
        snap_max_pts=snap_max_pts,
        recent_deltas=recent_deltas,
    )
    skip_candidates = _collect_alignment_candidates(
        pool=pool,
        start_index=det_ptr,
        window_lo=skip_window_lo,
        window_hi=skip_window_hi,
        pos_interp=next_interp_skip,
        span_t=next_span_t,
        full_span_threshold=full_span_threshold,
        snap_max_pts=snap_max_pts,
        prev_matched_pos=prev_matched_pos,
        gap_score_weight=gap_score_weight,
    )

    anchors_if_consume = list(anchors)
    _insert_anchor(anchors_if_consume, (groups[gi][0][2], best.dl.pos))
    consume_recent_deltas = [*recent_deltas, abs(best.dl.pos - pos_interp)]
    next_interp_consume = max(
        _progressive_interp(next_pos_t, anchors_if_consume),
        best.dl.pos + 0.5,
    )
    consume_window_lo, consume_window_hi = _compute_search_window(
        gi=gi + 1,
        groups=groups,
        anchors=anchors_if_consume,
        pos_interp=next_interp_consume,
        prev_matched_pos=best.dl.pos,
        snap_max_pts=snap_max_pts,
        recent_deltas=consume_recent_deltas,
    )
    consume_candidates = _collect_alignment_candidates(
        pool=pool,
        start_index=best.index + 1,
        window_lo=consume_window_lo,
        window_hi=consume_window_hi,
        pos_interp=next_interp_consume,
        span_t=next_span_t,
        full_span_threshold=full_span_threshold,
        snap_max_pts=snap_max_pts,
        prev_matched_pos=best.dl.pos,
        gap_score_weight=gap_score_weight,
    )

    best_skip = max(skip_candidates, key=lambda c: c.score, default=None)
    best_consume = max(consume_candidates, key=lambda c: c.score, default=None)
    if best_skip is None:
        return False
    if best_consume is None:
        return True
    return best_skip.score >= best_consume.score + lookahead_min_score_gain


def _read_manifest_rows() -> list[dict[str, str]]:
    """Load manifest rows in file order."""
    with MANIFEST_CSV.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        rows = list(reader)
    if ONLY_OUTPUT_PDFS:
        wanted = {name.lower() for name in ONLY_OUTPUT_PDFS}
        rows = [row for row in rows if Path(row["output_pdf"]).name.lower() in wanted]
    if LIMIT is not None:
        rows = rows[:LIMIT]
    return rows


def _normalize_manifest_value(value: str) -> Any:
    """Convert manifest scalar into int/float when possible."""
    if value == "":
        return value
    try:
        if "." in value:
            return float(value)
        return int(value)
    except ValueError:
        return value


def _load_base_cfg() -> dict[str, Any]:
    """Load project v2 config once, then apply local overrides per case."""
    cfg = dict(load_v2_config())
    cfg.update(CFG_OVERRIDES)
    return cfg


def _get_experiment_cfg() -> dict[str, Any]:
    """Return one predefined test-only experiment profile."""
    if EXPERIMENT_NAME not in _EXPERIMENT_PRESETS:
        pretty = ", ".join(sorted(_EXPERIMENT_PRESETS))
        raise SystemExit(
            f"Unknown STAMP_ALIGN_EXPERIMENT={EXPERIMENT_NAME!r}. Available: {pretty}"
        )
    return dict(_EXPERIMENT_PRESETS[EXPERIMENT_NAME])


def _prepare_case_cfg(base_cfg: dict[str, Any], pdf_path: Path, page_num: int) -> dict[str, Any]:
    """Build per-file config with debug context keys."""
    cfg = dict(base_cfg)
    cfg["_pdf_path"] = str(pdf_path)
    cfg["_page_num"] = page_num
    cfg["source_pdf_basename"] = pdf_path.name
    return cfg


def _expand_bbox_for_anchor(
    bbox: fitz.Rect,
    *,
    origin: str,
    x_frac: float,
    y_frac: float,
    page_rect: fitz.Rect,
) -> fitz.Rect:
    """Expand bbox asymmetrically opposite to the anchor corner."""
    if x_frac <= 0 and y_frac <= 0:
        return fitz.Rect(bbox)

    expanded = fitz.Rect(bbox)
    dx = bbox.width * max(0.0, x_frac)
    dy = bbox.height * max(0.0, y_frac)

    if origin.endswith("bottom_right"):
        expanded.x0 -= dx
        expanded.y0 -= dy
    elif origin.endswith("bottom_left"):
        expanded.x1 += dx
        expanded.y0 -= dy
    elif origin.endswith("top_right"):
        expanded.x0 -= dx
        expanded.y1 += dy
    else:
        expanded.x1 += dx
        expanded.y1 += dy

    expanded &= page_rect
    return expanded if not expanded.is_empty else fitz.Rect(bbox)


def _safe_name_from_pdf(pdf_path: Path) -> str:
    """Build a filesystem-safe stem for benchmark artifacts."""
    return pdf_path.stem.replace(" ", "_")


def _diag_to_dict(diag: AlignmentLineDiag) -> dict[str, Any]:
    """Serialize one alignment diagnostic line."""
    return {
        "line_id": diag.line_id,
        "orientation": diag.orientation,
        "pos_interp_pts": round(diag.pos_interp_pts, 3),
        "pos_snapped_pts": round(diag.pos_snapped_pts, 3),
        "snap_method": diag.snap_method,
        "matched_detected_pos": (
            round(diag.matched_detected_pos, 3)
            if diag.matched_detected_pos is not None
            else None
        ),
        "matched_detected_span": (
            round(diag.matched_detected_span, 3)
            if diag.matched_detected_span is not None
            else None
        ),
        "matched_detected_id": diag.matched_detected_id,
        "match_score": round(diag.match_score, 6),
        "span_ratio": round(diag.span_ratio, 6),
        "dp_action": diag.dp_action,
    }


def _detected_line_to_dict(line: DetectedGridLineInfo) -> dict[str, Any]:
    """Serialize one merged detected grid line."""
    return {
        "id": line.id,
        "orientation": line.orientation,
        "pos_pts": round(line.pos_pts, 3),
        "span_lo_pts": round(line.span_lo_pts, 3),
        "span_hi_pts": round(line.span_hi_pts, 3),
        "span_length_pts": round(line.span_hi_pts - line.span_lo_pts, 3),
        "cell_count": line.cell_count,
    }


def _prepare_alignment_inputs(
    *,
    template: StampTemplate,
    frame: Any,
    fitz_page: fitz.Page,
    cfg: dict[str, Any],
    experiment_cfg: dict[str, Any],
    clip_bbox_override: fitz.Rect | None = None,
) -> dict[str, Any]:
    """Mirror the alignment-related part of adapt_by_cell_assignment()."""
    stamp_bbox_base = _template_stamp_bbox(template, frame)
    if stamp_bbox_base is None:
        raise ValueError("Template produced no stamp bbox")

    page_rect = fitz.Rect(0.0, 0.0, fitz_page.rect.width, fitz_page.rect.height)
    stamp_bbox_search = _expand_bbox_for_anchor(
        stamp_bbox_base,
        origin=template.origin,
        x_frac=float(experiment_cfg["bbox_overflow_x_frac"]),
        y_frac=float(experiment_cfg["bbox_overflow_y_frac"]),
        page_rect=page_rect,
    )

    detected_raw = get_detected_stamp_cells(
        fitz_page=fitz_page,
        stamp_bbox=stamp_bbox_search,
        tolerance_mm=template.grid_tolerance_detected_mm,
        cfg=cfg,
        template=template,
    )
    stamp_area = max(stamp_bbox_search.get_area(), 1e-9)
    detected = [dc for dc in detected_raw if dc.area <= stamp_area * _MAX_CELL_AREA_FRACTION]

    t_bboxes = [
        _template_field_bbox(field, frame)
        for field in template.fields
        if _field_in_stamp_geometry_union(field)
    ]
    if not t_bboxes:
        raise ValueError("Template has no in-stamp geometry fields")

    t_all = CellBbox(
        min(b.x0 for b in t_bboxes),
        min(b.y0 for b in t_bboxes),
        max(b.x1 for b in t_bboxes),
        max(b.y1 for b in t_bboxes),
    )

    if detected:
        detected, _prefilter_diag = prefilter_detected_cells_for_stamp_table(
            detected,
            stamp_bbox_base=(
                stamp_bbox_base.x0,
                stamp_bbox_base.y0,
                stamp_bbox_base.x1,
                stamp_bbox_base.y1,
            ),
            t_all=t_all,
            tol_mm=float(template.grid_tolerance_detected_mm),
            scale_pts_per_mm=float(SCALE),
            cfg=cfg,
        )
        clip_bbox = clip_bbox_override or (
            stamp_bbox_search
            if bool(experiment_cfg["use_expanded_bbox_for_d_all"])
            else stamp_bbox_base
        )
        stamp_cb = CellBbox.from_rect(clip_bbox)
        clipped_for_bbox: list[CellBbox] = []
        for dc in detected:
            cx0 = max(dc.x0, stamp_cb.x0)
            cy0 = max(dc.y0, stamp_cb.y0)
            cx1 = min(dc.x1, stamp_cb.x1)
            cy1 = min(dc.y1, stamp_cb.y1)
            if cx1 > cx0 and cy1 > cy0:
                clipped_for_bbox.append(CellBbox(cx0, cy0, cx1, cy1))
        if clipped_for_bbox:
            d_all = CellBbox(
                min(c.x0 for c in clipped_for_bbox),
                min(c.y0 for c in clipped_for_bbox),
                max(c.x1 for c in clipped_for_bbox),
                max(c.y1 for c in clipped_for_bbox),
            )
        else:
            d_all = t_all
    else:
        clipped_for_bbox = []
        d_all = t_all

    sx_raw, sy_raw, _, _ = _compute_global_transform(t_all, d_all)
    scale_threshold = float(experiment_cfg["scale_threshold"])
    if bool(experiment_cfg.get("inclusive_scale_threshold")):
        use_scale = abs(sx_raw - 1.0) >= scale_threshold or abs(sy_raw - 1.0) >= scale_threshold
    else:
        use_scale = abs(sx_raw - 1.0) > scale_threshold or abs(sy_raw - 1.0) > scale_threshold
    if use_scale:
        sx = sx_raw
        sy = sy_raw
        dx = d_all.x0 - t_all.x0 * sx
        dy = d_all.y0 - t_all.y0 * sy
    else:
        sx = 1.0
        sy = 1.0
        dx = d_all.x0 - t_all.x0
        dy = d_all.y0 - t_all.y0

    return {
        "stamp_bbox": fitz.Rect(clip_bbox_override) if clip_bbox_override is not None else fitz.Rect(stamp_bbox_search),
        "stamp_bbox_base": stamp_bbox_base,
        "stamp_bbox_search": stamp_bbox_search,
        "clip_bbox": fitz.Rect(clip_bbox) if detected else fitz.Rect(clip_bbox_override or stamp_bbox_base),
        "detected_raw": detected_raw,
        "detected": detected,
        "filtered_oversized_count": len(detected_raw) - len(detected),
        "t_all": t_all,
        "d_all": d_all,
        "clipped_for_bbox_count": len(clipped_for_bbox),
        "sx_raw": sx_raw,
        "sy_raw": sy_raw,
        "sx": sx,
        "sy": sy,
        "dx": dx,
        "dy": dy,
        "use_scale": use_scale,
        "scale_threshold": scale_threshold,
    }


def _align_grid_lines_ordered_experiment(
    template: StampTemplate,
    frame: Any,
    sx: float,
    sy: float,
    dx: float,
    dy: float,
    detected: list[CellBbox],
    experiment_cfg: dict[str, Any],
    diagnostics: list[AlignmentLineDiag] | None = None,
    summary_out: list[AlignmentSummary] | None = None,
    detected_merged_out: list[DetectedGridLineInfo] | None = None,
    alignment_report_out: list[AlignmentReport] | None = None,
    template_path: str = "",
    pdf_path: str = "",
    page_num: int = 0,
    stamp_bbox_override: fitz.Rect | None = None,
) -> dict[str, float]:
    """Benchmark-only variant of ordered alignment with extra H-line heuristics."""
    _CO_LOCATE_PTS = 2.0 * SCALE
    MATCH_THRESHOLD = 0.35

    if not template.grid_lines:
        return {}

    h_det_raw, v_det_raw = _build_detected_grid_lines(
        detected, template.grid_tolerance_detected_mm
    ) if detected else ([], [])
    h_raw_count = len(h_det_raw)
    v_raw_count = len(v_det_raw)

    merged_spans_mm = _virtual_merge_template_lines(template.grid_lines)
    t_lines: list[tuple[str, str, float, float, float, str | None, float]] = []
    for gl in template.grid_lines:
        orientation, pos_fitz, start_fitz, end_fitz, bnd, line_id = grid_line_to_fitz_pts(
            gl, frame, template.origin
        )
        pos_t = pos_fitz * (sy if orientation == "h" else sx) + (dy if orientation == "h" else dx)
        msm = merged_spans_mm.get(line_id, abs(gl.end_mm - gl.start_mm))
        t_lines.append((line_id, orientation, pos_t, start_fitz, end_fitz, bnd, msm))

    stamp_bbox = fitz.Rect(stamp_bbox_override) if stamp_bbox_override is not None else _template_stamp_bbox(template, frame)
    if stamp_bbox is not None:
        stamp_w_pts = max(1.0, stamp_bbox.width)
        stamp_h_pts = max(1.0, stamp_bbox.height)
    else:
        stamp_w_pts = 1.0
        stamp_h_pts = 1.0

    snapped: dict[str, float] = {}
    snap_max_pts = template.snap_max_distance_mm * SCALE
    border_merge_pts = _compute_border_merge_dist(detected)
    h_det_merged = _merge_border_pairs(h_det_raw, border_merge_pts)
    v_det_merged = _merge_border_pairs(v_det_raw, border_merge_pts)
    h_merged_count = len(h_det_merged)
    v_merged_count = len(v_det_merged)

    if stamp_bbox is not None:
        margin = snap_max_pts
        h_det_merged = _filter_outside_stamp(h_det_merged, stamp_bbox.y0, stamp_bbox.y1, margin)
        v_det_merged = _filter_outside_stamp(v_det_merged, stamp_bbox.x0, stamp_bbox.x1, margin)

    h_filtered_count = len(h_det_merged)
    v_filtered_count = len(v_det_merged)
    h_id_map, v_id_map = _assign_detected_ids(h_det_merged, v_det_merged)

    if detected_merged_out is not None:
        for dl in h_det_merged:
            detected_merged_out.append(
                DetectedGridLineInfo(h_id_map[id(dl)], "h", dl.pos, dl.span_lo, dl.span_hi, dl.cell_count)
            )
        for dl in v_det_merged:
            detected_merged_out.append(
                DetectedGridLineInfo(v_id_map[id(dl)], "v", dl.pos, dl.span_lo, dl.span_hi, dl.cell_count)
            )

    boundary_diags: list[AlignmentLineDiag] = []
    summary = AlignmentSummary()
    report_problems: list[str] = []
    all_match_scores: list[float] = []
    line_method: dict[str, str] = {}
    line_matched_to: dict[str, str] = {}
    log_lines = [
        "=== GRID ALIGNMENT LOG ===",
        f"Template: {template_path or template.name} ({len(template.grid_lines)} grid lines)",
        f"PDF: {os.path.basename(pdf_path) if pdf_path else '(unknown)'} (page {page_num})",
        f"Frame: x0={frame.x0:.1f} y0={frame.y0:.1f} x1={frame.x1:.1f} y1={frame.y1:.1f}",
        f"Scale: sx={sx:.4f} sy={sy:.4f} dx={dx:.1f} dy={dy:.1f}",
        (
            f"Params: MATCH_THRESHOLD={MATCH_THRESHOLD} BORDER_MERGE={border_merge_pts:.1f}pts "
            f"CO_LOCATE={_CO_LOCATE_PTS:.1f}pts EXPERIMENT={EXPERIMENT_NAME}"
        ),
        "Log format: .cursor/rules/AI_v2_grid_log.mdc",
        "Engine rules: .cursor/rules/AI_v2_grid.mdc",
        "Code: benchmark experimental ordered aligner",
        "",
    ]

    for orient, pool, extent, id_map in [
        ("h", h_det_merged, stamp_w_pts, h_id_map),
        ("v", v_det_merged, stamp_h_pts, v_id_map),
    ]:
        boundaries = [t for t in t_lines if t[1] == orient and t[5] is not None]
        for lid, _o, pos_t, _sp_lo, _sp_hi, _bnd, _msm in boundaries:
            if not pool:
                snapped[lid] = pos_t
                line_method[lid] = "boundary-fallback"
                boundary_diags.append(AlignmentLineDiag(lid, orient, pos_t, pos_t, "fallback"))
                continue
            best_dl = None
            best_score = -1.0
            top_pref_dl = None
            top_pref_score = -1.0
            full_span_threshold = extent * 0.3
            for dl in pool:
                dist = abs(dl.pos - pos_t)
                if dist > snap_max_pts * 2.5:
                    continue
                dist_score = 1.0 - min(dist / max(snap_max_pts, 1.0), 1.0)
                span_score = min(dl.span_length / max(extent * 0.3, 1.0), 1.0)
                sc = 0.6 * dist_score + 0.4 * span_score
                if sc > best_score:
                    best_score = sc
                    best_dl = dl
                if (
                    orient == "h"
                    and lid == "h_top"
                    and dl.pos <= pos_t
                    and dist <= snap_max_pts
                    and dl.span_length >= full_span_threshold
                    and (top_pref_dl is None or dl.pos < top_pref_dl.pos)
                ):
                    top_pref_dl = dl
                    top_pref_score = sc
            method = "boundary"
            if (
                orient == "h"
                and lid == "h_top"
                and top_pref_dl is not None
                and best_dl is not None
                and best_dl is not top_pref_dl
                and best_dl.pos > pos_t
                and top_pref_score >= max(MATCH_THRESHOLD, best_score - 0.2)
            ):
                best_dl = top_pref_dl
                best_score = top_pref_score
                method = "boundary-top-pref"
            if best_dl is None:
                snapped[lid] = pos_t
                line_method[lid] = "boundary-fallback"
                boundary_diags.append(AlignmentLineDiag(lid, orient, pos_t, pos_t, "fallback"))
                continue
            snapped[lid] = best_dl.pos
            summary.boundary_matched += 1
            all_match_scores.append(best_score)
            det_id = id_map.get(id(best_dl), "?")
            line_method[lid] = method
            line_matched_to[lid] = det_id
            boundary_diags.append(
                AlignmentLineDiag(
                    lid,
                    orient,
                    pos_t,
                    best_dl.pos,
                    method,
                    matched_detected_pos=best_dl.pos,
                    matched_detected_span=best_dl.span_length,
                    matched_detected_id=det_id,
                    match_score=best_score,
                )
            )

    def _boundary_affine(orient: str) -> tuple[float, float]:
        pairs = sorted(
            [
                (t[2], snapped[t[0]])
                for t in t_lines
                if t[1] == orient and t[5] is not None and t[0] in snapped
            ],
            key=lambda p: p[0],
        )
        if len(pairs) < 2:
            return 1.0, 0.0
        t_lo, s_lo = pairs[0]
        t_hi, s_hi = pairs[-1]
        dt = t_hi - t_lo
        if abs(dt) < _EPS:
            return 1.0, 0.0
        local_s = (s_hi - s_lo) / dt
        return local_s, s_lo - t_lo * local_s

    h_skip_artifact = 0
    v_skip_artifact = 0
    for orient, pool_raw, extent, id_map in [
        ("h", h_det_merged, stamp_w_pts, h_id_map),
        ("v", v_det_merged, stamp_h_pts, v_id_map),
    ]:
        non_boundary = [t for t in t_lines if t[1] == orient and t[5] is None]
        if not non_boundary:
            continue

        bnd_s, bnd_d = _boundary_affine(orient)
        seed_interp = {lid: pos_t * bnd_s + bnd_d for lid, _o, pos_t, *_ in non_boundary}
        non_boundary.sort(key=lambda t: seed_interp[t[0]])

        groups: list[list[tuple[str, str, float, float, float, str | None, float]]] = []
        for t in non_boundary:
            if groups and abs(seed_interp[t[0]] - seed_interp[groups[-1][0][0]]) <= _CO_LOCATE_PTS:
                groups[-1].append(t)
            else:
                groups.append([t])

        anchors = sorted(
            [
                (t[2], snapped[t[0]])
                for t in t_lines
                if t[1] == orient and t[5] is not None and t[0] in snapped
            ],
            key=lambda p: p[0],
        )
        boundary_positions = {
            round(snapped[t[0]], 1)
            for t in t_lines
            if t[1] == orient and t[5] is not None and t[0] in snapped
        }
        pool = [
            dl
            for dl in pool_raw
            if not any(abs(dl.pos - bp) < 1.0 for bp in boundary_positions)
        ]
        pool.sort(key=lambda dl: dl.pos)

        full_span_threshold = extent * 0.3
        gap_score_weight = float(experiment_cfg.get("gap_score_weight_h", 0.0)) if orient == "h" else 0.0
        use_lookahead = orient == "h" and bool(experiment_cfg.get("lookahead_skip_h"))
        lookahead_min_score_gain = float(experiment_cfg.get("lookahead_min_score_gain", 0.05))

        walk_label = "H" if orient == "h" else "V"
        log_lines.append(f"WALK {walk_label} ({len(groups)} TPL groups, {len(pool)} det pool):")
        for bd in boundary_diags:
            if bd.orientation != orient:
                continue
            if bd.matched_detected_id:
                log_lines.append(
                    f"  TPL:{bd.line_id} -> {bd.matched_detected_id} "
                    f"({bd.pos_snapped_pts:.1f}) score={bd.match_score:.2f} {bd.snap_method} ✓"
                )

        det_ptr = 0
        prev_matched_pos = -1e18
        recent_deltas: list[float] = []

        for gi, group in enumerate(groups):
            pos_t = group[0][2]
            span_t = sum(t[6] for t in group) * SCALE
            pos_interp = max(_progressive_interp(pos_t, anchors), prev_matched_pos + 0.5)
            anchor_lo_id, anchor_hi_id = _anchor_bracket_ids(pos_t, anchors, t_lines, orient)
            window_lo, window_hi = _compute_search_window(
                gi=gi,
                groups=groups,
                anchors=anchors,
                pos_interp=pos_interp,
                prev_matched_pos=prev_matched_pos,
                snap_max_pts=snap_max_pts,
                recent_deltas=recent_deltas,
            )
            candidates = _collect_alignment_candidates(
                pool=pool,
                start_index=det_ptr,
                window_lo=window_lo,
                window_hi=window_hi,
                pos_interp=pos_interp,
                span_t=span_t,
                full_span_threshold=full_span_threshold,
                snap_max_pts=snap_max_pts,
                prev_matched_pos=prev_matched_pos,
                gap_score_weight=gap_score_weight,
            )
            best = max(candidates, key=lambda c: c.score, default=None)
            force_interp = (
                use_lookahead
                and _should_force_interp_by_lookahead(
                    orient=orient,
                    best=best,
                    gi=gi,
                    groups=groups,
                    anchors=anchors,
                    pool=pool,
                    det_ptr=det_ptr,
                    pos_interp=pos_interp,
                    prev_matched_pos=prev_matched_pos,
                    recent_deltas=recent_deltas,
                    snap_max_pts=snap_max_pts,
                    full_span_threshold=full_span_threshold,
                    gap_score_weight=gap_score_weight,
                    lookahead_min_score_gain=lookahead_min_score_gain,
                )
            )
            group_ids_str = ", ".join(f"TPL:{t[0]}" for t in group)

            if best is not None and best.score >= MATCH_THRESHOLD and not force_interp:
                for t in group:
                    snapped[t[0]] = best.dl.pos
                    line_method[t[0]] = "match-gap" if gap_score_weight > 0 else "match"
                summary.alignment_matched += 1
                all_match_scores.append(best.score)
                prev_matched_pos = best.dl.pos
                det_ptr = best.index + 1
                delta = best.dl.pos - pos_interp
                recent_deltas.append(abs(delta))
                _insert_anchor(anchors, (pos_t, best.dl.pos))
                det_id = id_map.get(id(best.dl), "?")
                for t in group:
                    line_matched_to[t[0]] = det_id
                log_lines.append(
                    f"  {group_ids_str} -> {det_id} ({best.dl.pos:.1f}) "
                    f"score={best.score:.2f} interp={pos_interp:.1f} delta={delta:+.1f} "
                    f"gap={best.gap_score:.2f} anchors=({anchor_lo_id},{anchor_hi_id}) ✓"
                )
                if diagnostics is not None:
                    for t in group:
                        diagnostics.append(
                            AlignmentLineDiag(
                                t[0],
                                orient,
                                pos_interp,
                                best.dl.pos,
                                "alignment",
                                matched_detected_pos=best.dl.pos,
                                matched_detected_span=best.dl.span_length,
                                matched_detected_id=det_id,
                                match_score=best.score,
                                span_ratio=best.span_ratio,
                                dp_action="match-gap" if gap_score_weight > 0 else "match",
                            )
                        )
                for candidate in sorted(candidates, key=lambda c: c.index):
                    if candidate.index == best.index:
                        continue
                    sk_id = id_map.get(id(candidate.dl), "?")
                    summary.detected_skipped += 1
                    if orient == "h":
                        h_skip_artifact += 1
                    else:
                        v_skip_artifact += 1
                    reason = f"split-artifact, span={candidate.dl.span_length:.0f}pts"
                    log_lines.append(f"    [skip {sk_id} ({candidate.dl.pos:.1f}) {reason}]")
                    report_problems.append(f"⊘ {sk_id} ({candidate.dl.pos:.1f}) skip: {reason}")
            else:
                summary.template_skipped += 1
                for t in group:
                    snapped[t[0]] = pos_interp
                    line_method[t[0]] = "interp-lookahead" if force_interp else "interp"
                nearest_info = ""
                if best is not None:
                    n_id = id_map.get(id(best.dl), "?")
                    nearest_info = (
                        f" (nearest {n_id} dist={best.dist:.1f}pts, "
                        f"score={best.score:.2f}{', lookahead-hold' if force_interp else ''})"
                    )
                log_lines.append(
                    f"  {group_ids_str} -> NO MATCH, interp={pos_interp:.1f}{nearest_info} "
                    f"anchors=({anchor_lo_id},{anchor_hi_id}) ⚠"
                )
                report_problems.append(
                    f"⚠ {group_ids_str} no_match -> interpolated {pos_interp:.1f}{nearest_info}"
                )
                if diagnostics is not None:
                    for t in group:
                        diagnostics.append(
                            AlignmentLineDiag(
                                t[0],
                                orient,
                                pos_interp,
                                pos_interp,
                                "fallback",
                                dp_action="lookahead-hold" if force_interp else "",
                            )
                        )
        log_lines.append("")

    if diagnostics is not None:
        diagnostics.extend(boundary_diags)
    if summary_out is not None:
        summary_out.append(summary)

    avg_score = sum(all_match_scores) / len(all_match_scores) if all_match_scores else 0.0
    summary.avg_match_score = avg_score
    warn_threshold = max(5, int(len(template.grid_lines) * 0.15))
    if summary.template_skipped == 0 and avg_score >= 0.5:
        verdict = "ok"
    elif summary.template_skipped <= warn_threshold or avg_score < 0.5:
        verdict = "warning"
    else:
        verdict = "fail"

    h_tpl_count = sum(1 for t in t_lines if t[1] == "h")
    v_tpl_count = sum(1 for t in t_lines if t[1] == "v")
    counts_lines = [
        f"VERDICT: {verdict.upper()} ({len(report_problems)} problems)" if report_problems else f"VERDICT: {verdict.upper()}",
        "",
        "COUNTS:",
        (
            f"  H: {h_tpl_count} TPL | {h_raw_count} det(raw) -> {h_merged_count} det(merged) "
            f"-> {h_filtered_count} det(in-stamp) | {h_skip_artifact} skip-artifact"
        ),
        (
            f"  V: {v_tpl_count} TPL | {v_raw_count} det(raw) -> {v_merged_count} det(merged) "
            f"-> {v_filtered_count} det(in-stamp) | {v_skip_artifact} skip-artifact"
        ),
        "",
        "Code: benchmark experimental ordered aligner",
        "",
    ]
    log_lines[8:8] = counts_lines
    if report_problems:
        log_lines.append("PROBLEMS:")
        log_lines.extend(f"  {problem}" for problem in report_problems)
        log_lines.append("")
    log_lines.append("SNAPPED POSITIONS (final):")
    log_lines.append("  Line ID               Pos pts Method             Matched to  ")
    log_lines.append("  -------------------- -------- ------------------ ------------")
    for lid, _o, pos_t, _sp_lo, _sp_hi, _bnd, _msm in sorted(
        t_lines, key=lambda t: (t[1], snapped.get(t[0], t[2]))
    ):
        log_lines.append(
            f"  {lid:<20} {snapped.get(lid, pos_t):8.1f} {line_method.get(lid, ''):<18} {line_matched_to.get(lid, ''):<12}"
        )

    if alignment_report_out is not None:
        alignment_report_out.append(
            AlignmentReport(
                verdict=verdict,
                message=f"verdict={verdict}; avg_score={avg_score:.2f}",
                avg_match_score=avg_score,
                problem_lines=report_problems,
                log="\n".join(log_lines) + "\n",
            )
        )
    return snapped


def _pick_h_top_candidate(
    *,
    detected_lines: list[DetectedGridLineInfo],
    stamp_bbox: fitz.Rect,
    h_top_diag: AlignmentLineDiag | None,
    snap_max_distance_mm: float,
) -> dict[str, Any]:
    """Collect top-boundary candidate information for the h_top hypothesis."""
    snap_max_pts = snap_max_distance_mm * 72.0 / 25.4
    full_span_threshold = stamp_bbox.width * 0.3

    h_lines = [line for line in detected_lines if line.orientation == "h"]
    full_span_lines = [
        line
        for line in h_lines
        if (line.span_hi_pts - line.span_lo_pts) >= full_span_threshold
    ]
    full_span_lines.sort(key=lambda line: line.pos_pts)

    result: dict[str, Any] = {
        "h_top_window_candidate_count": 0,
        "h_top_full_span_candidate_count": len(full_span_lines),
        "h_top_window_top_candidate_id": "",
        "h_top_window_top_candidate_pos_pts": "",
        "h_top_window_top_candidate_span_pts": "",
        "h_top_window_top_candidate_delta_pts": "",
        "h_top_matches_window_top_candidate": "",
        "h_top_over_top_candidate_delta_pts": "",
        "h_top_global_top_candidate_id": "",
        "h_top_global_top_candidate_pos_pts": "",
        "h_top_global_top_candidate_span_pts": "",
    }
    if not full_span_lines:
        return result

    global_top = full_span_lines[0]
    result["h_top_global_top_candidate_id"] = global_top.id
    result["h_top_global_top_candidate_pos_pts"] = round(global_top.pos_pts, 3)
    result["h_top_global_top_candidate_span_pts"] = round(
        global_top.span_hi_pts - global_top.span_lo_pts, 3
    )

    if h_top_diag is None:
        return result

    in_window = [
        line
        for line in full_span_lines
        if abs(line.pos_pts - h_top_diag.pos_interp_pts) <= snap_max_pts * 2.5
    ]
    result["h_top_window_candidate_count"] = len(in_window)
    if not in_window:
        return result

    top_candidate = min(in_window, key=lambda line: line.pos_pts)
    top_candidate_span = top_candidate.span_hi_pts - top_candidate.span_lo_pts
    result["h_top_window_top_candidate_id"] = top_candidate.id
    result["h_top_window_top_candidate_pos_pts"] = round(top_candidate.pos_pts, 3)
    result["h_top_window_top_candidate_span_pts"] = round(top_candidate_span, 3)
    result["h_top_window_top_candidate_delta_pts"] = round(
        top_candidate.pos_pts - h_top_diag.pos_interp_pts,
        3,
    )

    if h_top_diag.matched_detected_id:
        result["h_top_matches_window_top_candidate"] = (
            h_top_diag.matched_detected_id == top_candidate.id
        )
        result["h_top_over_top_candidate_delta_pts"] = round(
            h_top_diag.pos_snapped_pts - top_candidate.pos_pts,
            3,
        )

    return result


def _pick_boundary_like_line(
    *,
    detected_lines: list[DetectedGridLineInfo],
    orientation: str,
    expected_span_pts: float,
    edge: str,
    span_ratio_min: float = 0.8,
    fallback_ratio_min: float = 0.6,
) -> tuple[DetectedGridLineInfo | None, list[DetectedGridLineInfo], list[DetectedGridLineInfo]]:
    """Pick a boundary-like detected line by span compatibility and extremeness.

    The intent is test-only diagnostics for deformed stamp extent inference.
    This deliberately does not use template positional proximity.
    """
    orient_lines = [line for line in detected_lines if line.orientation == orientation]
    if not orient_lines:
        return None, [], []

    def _span_ratio(line: DetectedGridLineInfo) -> float:
        line_span = max(line.span_hi_pts - line.span_lo_pts, _EPS)
        return min(line_span, expected_span_pts) / max(line_span, expected_span_pts, _EPS)

    def _sort_key(line: DetectedGridLineInfo) -> tuple[float, float, str]:
        line_span = line.span_hi_pts - line.span_lo_pts
        edge_pos = line.pos_pts if edge in ("top", "left") else -line.pos_pts
        return (edge_pos, -line_span, line.id)

    strong = [line for line in orient_lines if _span_ratio(line) >= span_ratio_min]
    medium = [line for line in orient_lines if _span_ratio(line) >= fallback_ratio_min]
    chosen_pool = strong or medium
    chosen = sorted(chosen_pool, key=_sort_key)[0] if chosen_pool else None
    return chosen, strong, medium


def _build_merged_detected_line_infos(
    *,
    detected_cells: list[CellBbox],
    template: StampTemplate,
) -> list[DetectedGridLineInfo]:
    """Build merged detected grid lines without the final stamp-position filter."""
    h_det_raw, v_det_raw = _build_detected_grid_lines(
        detected_cells, template.grid_tolerance_detected_mm
    ) if detected_cells else ([], [])
    border_merge_pts = _compute_border_merge_dist(detected_cells)
    h_det_merged = _merge_border_pairs(h_det_raw, border_merge_pts)
    v_det_merged = _merge_border_pairs(v_det_raw, border_merge_pts)
    h_id_map, v_id_map = _assign_detected_ids(h_det_merged, v_det_merged)

    merged: list[DetectedGridLineInfo] = []
    for orient, merged_lines, id_map in (
        ("h", h_det_merged, h_id_map),
        ("v", v_det_merged, v_id_map),
    ):
        for dl in merged_lines:
            merged.append(
                DetectedGridLineInfo(
                    id=id_map[id(dl)],
                    orientation=orient,
                    pos_pts=dl.pos,
                    span_lo_pts=dl.span_lo,
                    span_hi_pts=dl.span_hi,
                    cell_count=dl.cell_count,
                )
            )
    return merged


def _collect_template_boundary_positions(
    *,
    template: StampTemplate,
    frame: fitz.Rect,
    orientation: str,
) -> list[tuple[str, float]]:
    """Return ordered template boundary positions in fitz coordinates."""
    positions: list[tuple[str, float]] = []
    for grid_line in template.grid_lines:
        line_orientation, pos_fitz, _start_fitz, _end_fitz, boundary, line_id = (
            grid_line_to_fitz_pts(grid_line, frame, template.origin)
        )
        if line_orientation != orientation or boundary is None:
            continue
        positions.append((line_id, pos_fitz))
    positions.sort(key=lambda item: item[1])
    return positions


def _collect_boundary_like_detected_positions(
    *,
    detected_lines: list[DetectedGridLineInfo],
    orientation: str,
    expected_span_pts: float,
    edge: str,
    span_ratio_min: float = 0.55,
) -> list[DetectedGridLineInfo]:
    """Return ordered detected lines that are plausible boundary carriers."""
    orient_lines = [line for line in detected_lines if line.orientation == orientation]
    if not orient_lines:
        return []

    def _span_ratio(line: DetectedGridLineInfo) -> float:
        span_pts = max(line.span_hi_pts - line.span_lo_pts, _EPS)
        return min(span_pts, expected_span_pts) / max(span_pts, expected_span_pts, _EPS)

    filtered = [line for line in orient_lines if _span_ratio(line) >= span_ratio_min]
    filtered.sort(key=lambda line: (line.pos_pts, -(line.span_hi_pts - line.span_lo_pts), line.id))
    if edge in ("right", "bottom"):
        filtered.reverse()
    return filtered


def _infer_virtual_outer_edge(
    *,
    template_positions: list[tuple[str, float]],
    detected_candidates: list[DetectedGridLineInfo],
    max_skip: int = 3,
) -> dict[str, Any]:
    """Infer a virtual outer edge from first detected lines and template offsets.

    This is a benchmark-only heuristic: the first detected line is treated as the
    first visible line from the PDF list, while the template side may skip several
    outer lines that disappeared in detection.
    """
    empty = {
        "virtual_pos_pts": "",
        "matched_template_id": "",
        "matched_template_pos_pts": "",
        "source_detected_id": "",
        "source_detected_pos_pts": "",
        "skip_count": "",
        "scale_estimate": "",
        "score": "",
        "candidate_count": 0,
        "gap_count_used": 0,
    }
    if not template_positions or not detected_candidates:
        return empty

    tpl_positions_only = [pos for _line_id, pos in template_positions]
    usable_detected = detected_candidates[:4]
    best: dict[str, Any] | None = None

    for skip_count, (matched_template_id, matched_template_pos) in enumerate(template_positions[: max_skip + 1]):
        compare_count = min(len(usable_detected), len(template_positions) - skip_count)
        if compare_count <= 0:
            continue
        tpl_slice = tpl_positions_only[skip_count : skip_count + compare_count]
        det_slice = [line.pos_pts for line in usable_detected[:compare_count]]
        gap_count = max(0, compare_count - 1)
        scale_estimate = 1.0
        gap_error = 0.0
        if gap_count:
            ratios: list[float] = []
            residuals: list[float] = []
            for index in range(gap_count):
                tpl_gap = tpl_slice[index + 1] - tpl_slice[index]
                det_gap = det_slice[index + 1] - det_slice[index]
                if abs(tpl_gap) <= _EPS:
                    continue
                ratios.append(det_gap / tpl_gap)
            if ratios:
                ratios_sorted = sorted(ratios)
                scale_estimate = ratios_sorted[len(ratios_sorted) // 2]
                for index in range(gap_count):
                    tpl_gap = tpl_slice[index + 1] - tpl_slice[index]
                    det_gap = det_slice[index + 1] - det_slice[index]
                    residuals.append(abs(det_gap - tpl_gap * scale_estimate))
                gap_error = sum(residuals) / len(residuals)

        virtual_pos = det_slice[0] - (matched_template_pos - tpl_positions_only[0]) * scale_estimate
        score = gap_error + skip_count * 0.75
        candidate = {
            "virtual_pos_pts": round(virtual_pos, 3),
            "matched_template_id": matched_template_id,
            "matched_template_pos_pts": round(matched_template_pos, 3),
            "source_detected_id": usable_detected[0].id,
            "source_detected_pos_pts": round(usable_detected[0].pos_pts, 3),
            "skip_count": skip_count,
            "scale_estimate": round(scale_estimate, 6),
            "score": round(score, 6),
            "candidate_count": len(usable_detected),
            "gap_count_used": gap_count,
        }
        if best is None or candidate["score"] < best["score"]:
            best = candidate

    return best or empty


def _extent_probe_proposed_bbox(extent_probe: dict[str, Any]) -> fitz.Rect | None:
    """Convert extent-probe scalar fields into a rect when available."""
    if not extent_probe.get("extent_probe_has_proposed_bbox"):
        return None
    keys = (
        "extent_probe_proposed_bbox_x0_pts",
        "extent_probe_proposed_bbox_y0_pts",
        "extent_probe_proposed_bbox_x1_pts",
        "extent_probe_proposed_bbox_y1_pts",
    )
    if any(extent_probe.get(key, "") == "" for key in keys):
        return None
    return fitz.Rect(*(float(extent_probe[key]) for key in keys))


def _probe_extent_candidates(
    *,
    detected_lines: list[DetectedGridLineInfo],
    stamp_bbox_base: fitz.Rect,
    origin: str,
    template_h_boundaries: list[tuple[str, float]],
    template_v_boundaries: list[tuple[str, float]],
) -> dict[str, Any]:
    """Infer a deformed stamp bbox from boundary-like detected lines.

    This is a benchmark-only probe. It does not affect alignment yet.
    """
    expected_h_span = max(stamp_bbox_base.width, _EPS)
    expected_v_span = max(stamp_bbox_base.height, _EPS)

    top_line, top_strong, top_medium = _pick_boundary_like_line(
        detected_lines=detected_lines,
        orientation="h",
        expected_span_pts=expected_h_span,
        edge="top",
    )
    left_line, left_strong, left_medium = _pick_boundary_like_line(
        detected_lines=detected_lines,
        orientation="v",
        expected_span_pts=expected_v_span,
        edge="left",
    )
    right_line, right_strong, right_medium = _pick_boundary_like_line(
        detected_lines=detected_lines,
        orientation="v",
        expected_span_pts=expected_v_span,
        edge="right",
    )
    bottom_line, bottom_strong, bottom_medium = _pick_boundary_like_line(
        detected_lines=detected_lines,
        orientation="h",
        expected_span_pts=expected_h_span,
        edge="bottom",
    )
    virtual_h_candidates = _collect_boundary_like_detected_positions(
        detected_lines=detected_lines,
        orientation="h",
        expected_span_pts=expected_h_span,
        edge="top",
    )
    virtual_v_candidates = _collect_boundary_like_detected_positions(
        detected_lines=detected_lines,
        orientation="v",
        expected_span_pts=expected_v_span,
        edge="left",
    )
    virtual_top = _infer_virtual_outer_edge(
        template_positions=template_h_boundaries,
        detected_candidates=virtual_h_candidates,
    )
    virtual_left = _infer_virtual_outer_edge(
        template_positions=template_v_boundaries,
        detected_candidates=virtual_v_candidates,
    )

    top_pos = top_line.pos_pts if top_line is not None else stamp_bbox_base.y0
    bottom_pos = bottom_line.pos_pts if bottom_line is not None else stamp_bbox_base.y1
    vertical_cover_tol = max(expected_v_span * 0.05, 8.0)

    def _pick_covering_vertical(
        preferred: DetectedGridLineInfo | None,
        strong_pool: list[DetectedGridLineInfo],
        medium_pool: list[DetectedGridLineInfo],
        edge_name: str,
    ) -> DetectedGridLineInfo | None:
        def _covers(line: DetectedGridLineInfo) -> bool:
            return (
                line.span_lo_pts <= top_pos + vertical_cover_tol
                and line.span_hi_pts >= bottom_pos - vertical_cover_tol
            )

        def _sort_key(line: DetectedGridLineInfo) -> tuple[float, float, str]:
            span_pts = line.span_hi_pts - line.span_lo_pts
            edge_pos = line.pos_pts if edge_name == "left" else -line.pos_pts
            return (edge_pos, -span_pts, line.id)

        if preferred is not None and _covers(preferred):
            return preferred
        for pool in (strong_pool, medium_pool):
            covering = [line for line in pool if _covers(line)]
            if covering:
                return sorted(covering, key=_sort_key)[0]
        return preferred

    left_line = _pick_covering_vertical(left_line, left_strong, left_medium, "left")
    right_line = _pick_covering_vertical(right_line, right_strong, right_medium, "right")

    x0_candidates: list[float] = []
    x1_candidates: list[float] = []
    y0_candidates: list[float] = []
    y1_candidates: list[float] = []
    if top_line is not None:
        x0_candidates.append(top_line.span_lo_pts)
        x1_candidates.append(top_line.span_hi_pts)
        y0_candidates.append(top_line.pos_pts)
    if bottom_line is not None:
        x0_candidates.append(bottom_line.span_lo_pts)
        x1_candidates.append(bottom_line.span_hi_pts)
        y1_candidates.append(bottom_line.pos_pts)
    if left_line is not None:
        x0_candidates.append(left_line.pos_pts)
        y0_candidates.append(left_line.span_lo_pts)
        y1_candidates.append(left_line.span_hi_pts)
    if right_line is not None:
        x1_candidates.append(right_line.pos_pts)
        y0_candidates.append(right_line.span_lo_pts)
        y1_candidates.append(right_line.span_hi_pts)

    proposed_bbox = fitz.Rect(stamp_bbox_base)
    inferred_edges: list[str] = []
    if origin.endswith("bottom_right"):
        if virtual_top["virtual_pos_pts"] != "":
            proposed_bbox.y0 = float(virtual_top["virtual_pos_pts"])
            inferred_edges.append("top_virtual")
        elif y0_candidates:
            proposed_bbox.y0 = min(y0_candidates)
            inferred_edges.append("top")
        if virtual_left["virtual_pos_pts"] != "":
            proposed_bbox.x0 = float(virtual_left["virtual_pos_pts"])
            inferred_edges.append("left_virtual")
        elif x0_candidates:
            proposed_bbox.x0 = min(x0_candidates)
            inferred_edges.append("left")
    elif origin.endswith("bottom_left"):
        if virtual_top["virtual_pos_pts"] != "":
            proposed_bbox.y0 = float(virtual_top["virtual_pos_pts"])
            inferred_edges.append("top_virtual")
        elif y0_candidates:
            proposed_bbox.y0 = min(y0_candidates)
            inferred_edges.append("top")
        if x1_candidates:
            proposed_bbox.x1 = max(x1_candidates)
            inferred_edges.append("right")
    elif origin.endswith("top_right"):
        if y1_candidates:
            proposed_bbox.y1 = max(y1_candidates)
            inferred_edges.append("bottom")
        if virtual_left["virtual_pos_pts"] != "":
            proposed_bbox.x0 = float(virtual_left["virtual_pos_pts"])
            inferred_edges.append("left_virtual")
        elif x0_candidates:
            proposed_bbox.x0 = min(x0_candidates)
            inferred_edges.append("left")
    else:
        if y1_candidates:
            proposed_bbox.y1 = max(y1_candidates)
            inferred_edges.append("bottom")
        if x1_candidates:
            proposed_bbox.x1 = max(x1_candidates)
            inferred_edges.append("right")

    def _line_payload(line: DetectedGridLineInfo | None, expected_span: float) -> dict[str, Any]:
        if line is None:
            return {
                "id": "",
                "pos_pts": "",
                "span_pts": "",
                "span_ratio_to_template": "",
            }
        span_pts = line.span_hi_pts - line.span_lo_pts
        span_ratio = min(span_pts, expected_span) / max(span_pts, expected_span, _EPS)
        return {
            "id": line.id,
            "pos_pts": round(line.pos_pts, 3),
            "span_pts": round(span_pts, 3),
            "span_ratio_to_template": round(span_ratio, 6),
        }

    result: dict[str, Any] = {
        "extent_probe_top_candidate_count_strong": len(top_strong),
        "extent_probe_top_candidate_count_medium": len(top_medium),
        "extent_probe_left_candidate_count_strong": len(left_strong),
        "extent_probe_left_candidate_count_medium": len(left_medium),
        "extent_probe_virtual_top_candidate_count": len(virtual_h_candidates),
        "extent_probe_virtual_left_candidate_count": len(virtual_v_candidates),
        "extent_probe_bottom_candidate_count_strong": len(bottom_strong),
        "extent_probe_bottom_candidate_count_medium": len(bottom_medium),
        "extent_probe_right_candidate_count_strong": len(right_strong),
        "extent_probe_right_candidate_count_medium": len(right_medium),
        "extent_probe_inferred_edges": ",".join(inferred_edges),
        "extent_probe_has_proposed_bbox": bool(inferred_edges),
        "extent_probe_line_source": "merged_prefilter",
        "extent_probe_proposed_bbox_x0_pts": round(proposed_bbox.x0, 3),
        "extent_probe_proposed_bbox_y0_pts": round(proposed_bbox.y0, 3),
        "extent_probe_proposed_bbox_x1_pts": round(proposed_bbox.x1, 3),
        "extent_probe_proposed_bbox_y1_pts": round(proposed_bbox.y1, 3),
        "extent_probe_base_bbox_x0_pts": round(stamp_bbox_base.x0, 3),
        "extent_probe_base_bbox_y0_pts": round(stamp_bbox_base.y0, 3),
        "extent_probe_base_bbox_x1_pts": round(stamp_bbox_base.x1, 3),
        "extent_probe_base_bbox_y1_pts": round(stamp_bbox_base.y1, 3),
        "extent_probe_proposed_dx0_pts": round(proposed_bbox.x0 - stamp_bbox_base.x0, 3),
        "extent_probe_proposed_dy0_pts": round(proposed_bbox.y0 - stamp_bbox_base.y0, 3),
        "extent_probe_proposed_dx1_pts": round(proposed_bbox.x1 - stamp_bbox_base.x1, 3),
        "extent_probe_proposed_dy1_pts": round(proposed_bbox.y1 - stamp_bbox_base.y1, 3),
    }

    top_payload = _line_payload(top_line, expected_h_span)
    left_payload = _line_payload(left_line, expected_v_span)
    bottom_payload = _line_payload(bottom_line, expected_h_span)
    right_payload = _line_payload(right_line, expected_v_span)
    for prefix, payload in (
        ("extent_probe_top", top_payload),
        ("extent_probe_left", left_payload),
        ("extent_probe_bottom", bottom_payload),
        ("extent_probe_right", right_payload),
    ):
        for key, value in payload.items():
            result[f"{prefix}_{key}"] = value
    for key, value in virtual_top.items():
        result[f"extent_probe_virtual_top_{key}"] = value
    for key, value in virtual_left.items():
        result[f"extent_probe_virtual_left_{key}"] = value

    return result


def _write_case_json(path: Path, payload: dict[str, Any]) -> None:
    """Write one per-file raw benchmark artifact."""
    with path.open("w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)


def _case_error_row(manifest_row: dict[str, str], exc: Exception) -> dict[str, Any]:
    """Build a summary row for a failed case."""
    row = {key: _normalize_manifest_value(value) for key, value in manifest_row.items()}
    row.update(
        {
            "status": "error",
            "alignment_verdict": "error",
            "boundary_matched": "",
            "alignment_matched": "",
            "template_skipped": "",
            "detected_skipped": "",
            "avg_match_score": "",
            "filtered_oversized_count": "",
            "use_scale": "",
            "transform_scale_x": "",
            "transform_scale_y": "",
            "transform_dx": "",
            "transform_dy": "",
            "h_top_method": "",
            "h_top_matched_to": "",
            "h_top_match_score": "",
            "h_top_pos_interp_pts": "",
            "h_top_pos_snapped_pts": "",
            "h_top_window_candidate_count": "",
            "h_top_full_span_candidate_count": "",
            "h_top_window_top_candidate_id": "",
            "h_top_window_top_candidate_pos_pts": "",
            "h_top_window_top_candidate_span_pts": "",
            "h_top_window_top_candidate_delta_pts": "",
            "h_top_matches_window_top_candidate": "",
            "h_top_over_top_candidate_delta_pts": "",
            "h_top_global_top_candidate_id": "",
            "h_top_global_top_candidate_pos_pts": "",
            "h_top_global_top_candidate_span_pts": "",
            "extent_probe_top_candidate_count_strong": "",
            "extent_probe_top_candidate_count_medium": "",
            "extent_probe_left_candidate_count_strong": "",
            "extent_probe_left_candidate_count_medium": "",
            "extent_probe_virtual_top_candidate_count": "",
            "extent_probe_virtual_left_candidate_count": "",
            "extent_probe_bottom_candidate_count_strong": "",
            "extent_probe_bottom_candidate_count_medium": "",
            "extent_probe_right_candidate_count_strong": "",
            "extent_probe_right_candidate_count_medium": "",
            "extent_probe_inferred_edges": "",
            "extent_probe_has_proposed_bbox": "",
            "extent_probe_line_source": "",
            "extent_probe_top_id": "",
            "extent_probe_top_pos_pts": "",
            "extent_probe_top_span_pts": "",
            "extent_probe_top_span_ratio_to_template": "",
            "extent_probe_left_id": "",
            "extent_probe_left_pos_pts": "",
            "extent_probe_left_span_pts": "",
            "extent_probe_left_span_ratio_to_template": "",
            "extent_probe_virtual_top_virtual_pos_pts": "",
            "extent_probe_virtual_top_matched_template_id": "",
            "extent_probe_virtual_top_matched_template_pos_pts": "",
            "extent_probe_virtual_top_source_detected_id": "",
            "extent_probe_virtual_top_source_detected_pos_pts": "",
            "extent_probe_virtual_top_skip_count": "",
            "extent_probe_virtual_top_scale_estimate": "",
            "extent_probe_virtual_top_score": "",
            "extent_probe_virtual_top_candidate_count": "",
            "extent_probe_virtual_top_gap_count_used": "",
            "extent_probe_virtual_left_virtual_pos_pts": "",
            "extent_probe_virtual_left_matched_template_id": "",
            "extent_probe_virtual_left_matched_template_pos_pts": "",
            "extent_probe_virtual_left_source_detected_id": "",
            "extent_probe_virtual_left_source_detected_pos_pts": "",
            "extent_probe_virtual_left_skip_count": "",
            "extent_probe_virtual_left_scale_estimate": "",
            "extent_probe_virtual_left_score": "",
            "extent_probe_virtual_left_candidate_count": "",
            "extent_probe_virtual_left_gap_count_used": "",
            "extent_probe_bottom_id": "",
            "extent_probe_bottom_pos_pts": "",
            "extent_probe_bottom_span_pts": "",
            "extent_probe_bottom_span_ratio_to_template": "",
            "extent_probe_right_id": "",
            "extent_probe_right_pos_pts": "",
            "extent_probe_right_span_pts": "",
            "extent_probe_right_span_ratio_to_template": "",
            "extent_probe_proposed_bbox_x0_pts": "",
            "extent_probe_proposed_bbox_y0_pts": "",
            "extent_probe_proposed_bbox_x1_pts": "",
            "extent_probe_proposed_bbox_y1_pts": "",
            "extent_probe_base_bbox_x0_pts": "",
            "extent_probe_base_bbox_y0_pts": "",
            "extent_probe_base_bbox_x1_pts": "",
            "extent_probe_base_bbox_y1_pts": "",
            "extent_probe_proposed_dx0_pts": "",
            "extent_probe_proposed_dy0_pts": "",
            "extent_probe_proposed_dx1_pts": "",
            "extent_probe_proposed_dy1_pts": "",
            "likely_top_boundary_issue": "",
            "error": f"{type(exc).__name__}: {exc}",
        }
    )
    return row


def _run_case(
    *,
    manifest_row: dict[str, str],
    template: StampTemplate,
    base_cfg: dict[str, Any],
    experiment_cfg: dict[str, Any],
    logs_dir: Path,
    cases_dir: Path,
) -> dict[str, Any]:
    """Run the benchmark for one generated PDF."""
    pdf_path = Path(manifest_row["output_pdf"])
    page_num = int(manifest_row["page_num"])
    cfg = _prepare_case_cfg(base_cfg, pdf_path, page_num)

    diagnostics: list[AlignmentLineDiag] = []
    summaries: list[AlignmentSummary] = []
    reports: list[AlignmentReport] = []
    detected_merged: list[DetectedGridLineInfo] = []

    with fitz.open(pdf_path) as doc:
        fitz_page = doc[page_num - 1]
        t_union = template if template.frame_mode == "drawing_union" else None
        frame, _ = find_frame(
            fitz_page,
            frame_mode=template.frame_mode,
            template=t_union,
            cfg=cfg,
        )

        template_h_boundaries = _collect_template_boundary_positions(
            template=template,
            frame=frame,
            orientation="h",
        )
        template_v_boundaries = _collect_template_boundary_positions(
            template=template,
            frame=frame,
            orientation="v",
        )
        initial_prep = _prepare_alignment_inputs(
            template=template,
            frame=frame,
            fitz_page=fitz_page,
            cfg=cfg,
            experiment_cfg=experiment_cfg,
        )
        extent_detected_lines = _build_merged_detected_line_infos(
            detected_cells=initial_prep["detected"],
            template=template,
        )
        extent_probe = _probe_extent_candidates(
            detected_lines=extent_detected_lines,
            stamp_bbox_base=initial_prep["stamp_bbox_base"],
            origin=template.origin,
            template_h_boundaries=template_h_boundaries,
            template_v_boundaries=template_v_boundaries,
        )
        proposed_bbox = _extent_probe_proposed_bbox(extent_probe)
        proposed_bbox_applied = bool(
            proposed_bbox is not None
            and (
                bool(experiment_cfg.get("use_proposed_bbox_for_d_all"))
                or bool(experiment_cfg.get("use_proposed_bbox_for_filter"))
            )
        )
        effective_bbox = fitz.Rect(proposed_bbox) if proposed_bbox_applied and proposed_bbox is not None else None
        prep = _prepare_alignment_inputs(
            template=template,
            frame=frame,
            fitz_page=fitz_page,
            cfg=cfg,
            experiment_cfg=experiment_cfg,
            clip_bbox_override=(
                effective_bbox
                if bool(experiment_cfg.get("use_proposed_bbox_for_d_all")) and effective_bbox is not None
                else None
            ),
        )
        align_fn = (
            _align_grid_lines_ordered_experiment
            if experiment_cfg.get("alignment_strategy") == "experimental"
            else _align_grid_lines_ordered
        )
        align_kwargs = {
            "diagnostics": diagnostics,
            "summary_out": summaries,
            "detected_merged_out": detected_merged,
            "alignment_report_out": reports,
            "template_path": template.source_path or template.name,
            "pdf_path": str(pdf_path),
            "page_num": page_num,
        }
        if align_fn is _align_grid_lines_ordered_experiment:
            align_fn(
                template,
                frame,
                prep["sx"],
                prep["sy"],
                prep["dx"],
                prep["dy"],
                prep["detected"],
                experiment_cfg=experiment_cfg,
                stamp_bbox_override=(
                    effective_bbox
                    if bool(experiment_cfg.get("use_proposed_bbox_for_filter")) and effective_bbox is not None
                    else None
                ),
                **align_kwargs,
            )
        else:
            align_fn(
                template,
                frame,
                prep["sx"],
                prep["sy"],
                prep["dx"],
                prep["dy"],
                prep["detected"],
                **align_kwargs,
            )

    summary = summaries[0] if summaries else AlignmentSummary()
    report = reports[0] if reports else AlignmentReport(verdict="error", message="No report generated")
    diag_by_id = {diag.line_id: diag for diag in diagnostics}
    h_top_diag = diag_by_id.get("h_top")
    h_top_probe = _pick_h_top_candidate(
        detected_lines=detected_merged,
        stamp_bbox=prep["stamp_bbox"],
        h_top_diag=h_top_diag,
        snap_max_distance_mm=template.snap_max_distance_mm,
    )
    effective_bbox_rect = effective_bbox if effective_bbox is not None else prep["clip_bbox"]
    effective_bbox_source = "extent_probe" if proposed_bbox_applied else "base/search"

    row = {key: _normalize_manifest_value(value) for key, value in manifest_row.items()}
    row.update(
        {
            "status": "ok",
            "alignment_verdict": report.verdict,
            "boundary_matched": summary.boundary_matched,
            "alignment_matched": summary.alignment_matched,
            "template_skipped": summary.template_skipped,
            "detected_skipped": summary.detected_skipped,
            "avg_match_score": round(summary.avg_match_score, 6),
            "filtered_oversized_count": prep["filtered_oversized_count"],
            "use_scale": prep["use_scale"],
            "transform_scale_x": round(prep["sx"], 6),
            "transform_scale_y": round(prep["sy"], 6),
            "transform_dx": round(prep["dx"], 3),
            "transform_dy": round(prep["dy"], 3),
            "scale_threshold": round(prep["scale_threshold"], 4),
            "proposed_bbox_applied": proposed_bbox_applied,
            "effective_bbox_source": effective_bbox_source,
            "effective_bbox_x0_pts": round(effective_bbox_rect.x0, 3),
            "effective_bbox_y0_pts": round(effective_bbox_rect.y0, 3),
            "effective_bbox_x1_pts": round(effective_bbox_rect.x1, 3),
            "effective_bbox_y1_pts": round(effective_bbox_rect.y1, 3),
            "h_top_method": h_top_diag.snap_method if h_top_diag else "",
            "h_top_matched_to": h_top_diag.matched_detected_id if h_top_diag else "",
            "h_top_match_score": round(h_top_diag.match_score, 6) if h_top_diag else "",
            "h_top_pos_interp_pts": round(h_top_diag.pos_interp_pts, 3) if h_top_diag else "",
            "h_top_pos_snapped_pts": round(h_top_diag.pos_snapped_pts, 3) if h_top_diag else "",
            "likely_top_boundary_issue": (
                bool(
                    h_top_probe["h_top_matches_window_top_candidate"] is False
                    and (
                        report.verdict != "ok"
                        or summary.template_skipped > 0
                    )
                )
                if h_top_probe["h_top_matches_window_top_candidate"] != ""
                else ""
            ),
            "error": "",
        }
    )
    row.update(h_top_probe)
    row.update(extent_probe)

    safe_name = _safe_name_from_pdf(pdf_path)
    if SAVE_ALIGNMENT_LOGS and report.log:
        log_path = logs_dir / f"{safe_name}.txt"
        log_path.write_text(report.log, encoding="utf-8")

    if SAVE_CASE_JSON:
        case_payload = {
            "manifest": manifest_row,
            "summary_row": row,
            "transform": {
                "use_scale": prep["use_scale"],
                "scale_threshold": round(prep["scale_threshold"], 4),
                "sx_raw": round(prep["sx_raw"], 6),
                "sy_raw": round(prep["sy_raw"], 6),
                "sx": round(prep["sx"], 6),
                "sy": round(prep["sy"], 6),
                "dx": round(prep["dx"], 3),
                "dy": round(prep["dy"], 3),
            },
            "stamp_bbox_pts": [
                round(prep["stamp_bbox"].x0, 3),
                round(prep["stamp_bbox"].y0, 3),
                round(prep["stamp_bbox"].x1, 3),
                round(prep["stamp_bbox"].y1, 3),
            ],
            "stamp_bbox_base_pts": [
                round(prep["stamp_bbox_base"].x0, 3),
                round(prep["stamp_bbox_base"].y0, 3),
                round(prep["stamp_bbox_base"].x1, 3),
                round(prep["stamp_bbox_base"].y1, 3),
            ],
            "stamp_bbox_search_pts": [
                round(prep["stamp_bbox_search"].x0, 3),
                round(prep["stamp_bbox_search"].y0, 3),
                round(prep["stamp_bbox_search"].x1, 3),
                round(prep["stamp_bbox_search"].y1, 3),
            ],
            "effective_bbox_pts": [
                round(effective_bbox_rect.x0, 3),
                round(effective_bbox_rect.y0, 3),
                round(effective_bbox_rect.x1, 3),
                round(effective_bbox_rect.y1, 3),
            ],
            "experiment": dict(experiment_cfg),
            "detected_counts": {
                "raw": len(prep["detected_raw"]),
                "filtered": len(prep["detected"]),
                "filtered_oversized": prep["filtered_oversized_count"],
                "clipped_for_bbox": prep["clipped_for_bbox_count"],
            },
            "proposed_bbox_applied": proposed_bbox_applied,
            "extent_probe": extent_probe,
            "grid_alignment_summary": {
                "boundary_matched": summary.boundary_matched,
                "alignment_matched": summary.alignment_matched,
                "template_skipped": summary.template_skipped,
                "detected_skipped": summary.detected_skipped,
                "avg_match_score": round(summary.avg_match_score, 6),
            },
            "alignment_report": {
                "verdict": report.verdict,
                "message": report.message,
                "problem_lines": report.problem_lines,
                "avg_match_score": round(report.avg_match_score, 6),
            },
            "diagnostics": [_diag_to_dict(diag) for diag in diagnostics],
            "detected_grid_lines": [_detected_line_to_dict(line) for line in detected_merged],
            "alignment_log": report.log,
        }
        _write_case_json(cases_dir / f"{safe_name}.json", case_payload)

    return row


def _write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    """Write summary rows to CSV."""
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    fieldnames: list[str] = []
    for row in rows:
        for key in row.keys():
            if key not in fieldnames:
                fieldnames.append(key)
    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def _write_jsonl(path: Path, rows: list[dict[str, Any]]) -> None:
    """Write summary rows to JSONL."""
    with path.open("w", encoding="utf-8") as fh:
        for row in rows:
            fh.write(json.dumps(row, ensure_ascii=False) + "\n")


def _build_decision_summary(rows: list[dict[str, Any]]) -> str:
    """Build a concise markdown summary for the run."""
    total = len(rows)
    ok_rows = [row for row in rows if row["status"] == "ok"]
    error_rows = [row for row in rows if row["status"] != "ok"]
    verdict_counts = Counter(row.get("alignment_verdict", "") for row in ok_rows)
    degraded_rows = [
        row
        for row in ok_rows
        if row["alignment_verdict"] != "ok" or int(row["template_skipped"]) > 0
    ]
    top_issue_rows = [row for row in ok_rows if row.get("likely_top_boundary_issue") is True]
    degraded_top_issue_rows = [
        row for row in degraded_rows if row.get("likely_top_boundary_issue") is True
    ]

    axis_buckets: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in ok_rows:
        sx = row.get("sx_percent")
        sy = row.get("sy_percent")
        if sx == 100 and sy == 100:
            axis = "identity"
        elif sx == 100:
            axis = "sy_only"
        elif sy == 100:
            axis = "sx_only"
        else:
            axis = "mixed"
        axis_buckets[axis].append(row)

    lines = [
        "# Alignment Benchmark Summary",
        "",
        f"- run_label: `{RUN_LABEL}`",
        f"- experiment: `{EXPERIMENT_NAME}`",
        f"- total_cases: `{total}`",
        f"- successful_cases: `{len(ok_rows)}`",
        f"- error_cases: `{len(error_rows)}`",
        f"- verdict_ok: `{verdict_counts.get('ok', 0)}`",
        f"- verdict_warning: `{verdict_counts.get('warning', 0)}`",
        f"- verdict_fail: `{verdict_counts.get('fail', 0)}`",
        f"- degraded_cases: `{len(degraded_rows)}`",
        f"- likely_top_boundary_issue_cases: `{len(top_issue_rows)}`",
        f"- degraded_cases_with_likely_top_issue: `{len(degraded_top_issue_rows)}`",
        "",
        "## Axis buckets",
    ]
    for axis_name in ("identity", "sx_only", "sy_only", "mixed"):
        bucket = axis_buckets.get(axis_name, [])
        if not bucket:
            continue
        degraded = sum(
            1
            for row in bucket
            if row["alignment_verdict"] != "ok" or int(row["template_skipped"]) > 0
        )
        top_issue = sum(1 for row in bucket if row.get("likely_top_boundary_issue") is True)
        lines.append(
            f"- `{axis_name}`: total={len(bucket)}, degraded={degraded}, likely_top_issue={top_issue}"
        )

    if degraded_top_issue_rows:
        lines.extend(
            [
                "",
                "## First degraded top-boundary cases",
            ]
        )
        sorted_rows = sorted(
            degraded_top_issue_rows,
            key=lambda row: (int(row["sy_percent"]), int(row["sx_percent"])),
        )
        for row in sorted_rows[:10]:
            lines.append(
                "- "
                f"`{Path(str(row['output_pdf'])).name}` "
                f"`sx={row['sx_percent']}` "
                f"`sy={row['sy_percent']}` "
                f"`verdict={row['alignment_verdict']}` "
                f"`h_top={row['h_top_matched_to'] or '-'} -> expected_top={row['h_top_window_top_candidate_id'] or '-'}`"
            )

    if error_rows:
        lines.extend(["", "## Errors"])
        for row in error_rows[:10]:
            lines.append(f"- `{Path(str(row['output_pdf'])).name}`: `{row['error']}`")

    return "\n".join(lines) + "\n"


def main() -> None:
    """Run the alignment benchmark on the generated deform dataset."""
    manifest_rows = _read_manifest_rows()
    if not manifest_rows:
        raise SystemExit("No manifest rows matched the current filters.")

    template = StampTemplate.from_json(str(TEMPLATE_PATH))
    base_cfg = _load_base_cfg()
    experiment_cfg = _get_experiment_cfg()

    run_dir = OUTPUT_ROOT / RUN_LABEL
    logs_dir = run_dir / "alignment_logs"
    cases_dir = run_dir / "cases"
    run_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)
    cases_dir.mkdir(parents=True, exist_ok=True)

    rows: list[dict[str, Any]] = []
    for index, manifest_row in enumerate(manifest_rows, start=1):
        pdf_name = Path(manifest_row["output_pdf"]).name
        print(f"[{index}/{len(manifest_rows)}] {pdf_name}")
        try:
            row = _run_case(
                manifest_row=manifest_row,
                template=template,
                base_cfg=base_cfg,
                experiment_cfg=experiment_cfg,
                logs_dir=logs_dir,
                cases_dir=cases_dir,
            )
        except Exception as exc:
            traceback.print_exc()
            row = _case_error_row(manifest_row, exc)
        rows.append(row)

    rows.sort(key=lambda row: (str(row.get("output_pdf", ""))))

    summary_csv = run_dir / "summary.csv"
    summary_jsonl = run_dir / "summary.jsonl"
    decision_md = run_dir / "decision_summary.md"

    _write_csv(summary_csv, rows)
    _write_jsonl(summary_jsonl, rows)
    decision_md.write_text(_build_decision_summary(rows), encoding="utf-8")

    print("Alignment benchmark finished")
    print(f"  template:           {TEMPLATE_PATH}")
    print(f"  experiment:         {EXPERIMENT_NAME}")
    print(f"  generated_dir:      {GENERATED_DIR}")
    print(f"  run_dir:            {run_dir}")
    print(f"  summary_csv:        {summary_csv}")
    print(f"  summary_jsonl:      {summary_jsonl}")
    print(f"  decision_summary:   {decision_md}")
    print(f"  processed_cases:    {len(rows)}")


if __name__ == "__main__":
    main()
