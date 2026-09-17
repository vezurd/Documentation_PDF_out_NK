"""Build drawable overlay items from ``stamp_prealign_debug`` bundle (JSON-safe dict)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from pdf_parsing_v2_engine.coord_transform import pdfminer_to_fitz

from .prealign_overlay_layers import LAYER_SPEC_BY_ID, PrealignLayerSpec


@dataclass
class PrealignOverlayItem:
    """One drawable primitive for the scene (PDF displayed pts → caller multiplies by dpi_scale)."""

    layer_id: str
    geometry: str  # rect | line | multiline
    # rect: one quad x0,y0,x1,y1; line: x0,y0,x1,y1; multiline: list of (x0,y0,x1,y1)
    segments: list[tuple[float, float, float, float]] = field(default_factory=list)
    label: str = ""
    color_rgb: tuple[int, int, int] = (0, 0, 0)
    pen_width: float = 1.5
    dashed: bool = False
    spec: PrealignLayerSpec | None = None


def _rect_to_displayed_quad(rect: Any, *, env: dict[str, Any] | None = None) -> list[float] | None:
    """Normalize bbox to [x0,y0,x1,y1] in **fitz displayed** (y-down), same as prealign rects."""
    if rect is None:
        return None
    if isinstance(rect, (list, tuple)) and len(rect) >= 4:
        try:
            return [float(rect[0]), float(rect[1]), float(rect[2]), float(rect[3])]
        except (TypeError, ValueError):
            return None
    if isinstance(rect, dict) and all(k in rect for k in ("x0", "y0", "x1", "y1")):
        try:
            x0, y0, x1, y1 = (
                float(rect["x0"]),
                float(rect["y0"]),
                float(rect["x1"]),
                float(rect["y1"]),
            )
        except (TypeError, ValueError):
            return None
        if "page_height" in rect:
            try:
                ph = float(rect["page_height"])
                rot = int((env or {}).get("rotation") or 0)
                r = pdfminer_to_fitz((x0, y0, x1, y1), ph, rot)
                return [float(r.x0), float(r.y0), float(r.x1), float(r.y1)]
            except (TypeError, ValueError, KeyError):
                return None
        return [x0, y0, x1, y1]
    return None


def _bbox_segments(rect: Any, *, env: dict[str, Any] | None = None) -> list[tuple[float, float, float, float]]:
    quad = _rect_to_displayed_quad(rect, env=env)
    if not quad:
        return []
    x0, y0, x1, y1 = quad
    if x1 <= x0 or y1 <= y0:
        return []
    return [(x0, y0, x1, y1)]


def _line_from_detected(line: dict[str, Any]) -> tuple[float, float, float, float] | None:
    try:
        if line.get("orientation") == "h":
            return (
                float(line["span_lo_pts"]),
                float(line["pos_pts"]),
                float(line["span_hi_pts"]),
                float(line["pos_pts"]),
            )
        return (
            float(line["pos_pts"]),
            float(line["span_lo_pts"]),
            float(line["pos_pts"]),
            float(line["span_hi_pts"]),
        )
    except (KeyError, TypeError, ValueError):
        return None


def _template_line_segment(
    line: dict[str, Any],
    snapped_pos: float | None,
) -> tuple[float, float, float, float] | None:
    try:
        pos = float(snapped_pos) if snapped_pos is not None else float(line["display_pos_pts"])
        if line.get("orientation") == "h":
            return (
                float(line["display_start_pts"]),
                pos,
                float(line["display_end_pts"]),
                pos,
            )
        return (
            pos,
            float(line["display_start_pts"]),
            pos,
            float(line["display_end_pts"]),
        )
    except (KeyError, TypeError, ValueError):
        return None


def build_overlay_items_from_bundle(bundle: dict[str, Any] | None) -> list[PrealignOverlayItem]:
    """Parse ``stamp_prealign_debug_bundle`` / metadata dict into overlay items."""
    if not bundle or not isinstance(bundle, dict):
        return []
    prealign = bundle.get("prealign") or {}
    extent = prealign.get("extent_probe") or {}
    selected = extent.get("selected_boundaries") or {}
    virtual = extent.get("virtual_edges") or {}
    search_bbox = prealign.get("search_bbox")
    items: list[PrealignOverlayItem] = []

    def add_rect(layer_id: str, rect: Any, *, dashed: bool = False) -> None:
        spec = LAYER_SPEC_BY_ID.get(layer_id)
        if spec is None or spec.geometry != "rect":
            return
        segs = _bbox_segments(rect, env=env)
        if not segs:
            return
        items.append(
            PrealignOverlayItem(
                layer_id=layer_id,
                geometry="rect",
                segments=segs,
                label=spec.label_full,
                color_rgb=spec.color_rgb,
                pen_width=2.0 if layer_id == "frame" else 1.7,
                dashed=dashed,
                spec=spec,
            )
        )

    def add_line(layer_id: str, seg: tuple[float, float, float, float] | None) -> None:
        spec = LAYER_SPEC_BY_ID.get(layer_id)
        if spec is None or seg is None:
            return
        items.append(
            PrealignOverlayItem(
                layer_id=layer_id,
                geometry="line",
                segments=[seg],
                label=spec.label_full,
                color_rgb=spec.color_rgb,
                pen_width=1.8,
                dashed=True,
                spec=spec,
            )
        )

    def add_multiline(layer_id: str, segs: list[tuple[float, float, float, float]]) -> None:
        spec = LAYER_SPEC_BY_ID.get(layer_id)
        if spec is None or not segs:
            return
        items.append(
            PrealignOverlayItem(
                layer_id=layer_id,
                geometry="multiline",
                segments=segs,
                label=spec.label_full,
                color_rgb=spec.color_rgb,
                pen_width=0.95 if layer_id == "extent_probe_merged_grid" else 2.0,
                dashed=False,
                spec=spec,
            )
        )

    env: dict[str, Any] = bundle.get("environment") or {}
    if not isinstance(env, dict):
        env = {}
    frame_rect = env.get("frame")
    add_rect("frame", frame_rect)

    add_rect("base_bbox", prealign.get("base_bbox"))
    add_rect("search_bbox", search_bbox, dashed=True)

    df = prealign.get("detected_prefilter") or {}
    if isinstance(df, dict):
        add_rect("prefilter_y_anchor_band", df.get("y_anchor_band_rect"))
        add_rect("prefilter_prior_rect", df.get("prior_rect"))
        add_rect("prefilter_kept_bbox", df.get("kept_component_bbox"))
        dropped = df.get("dropped_component_bboxes") or []
        drop_segs: list[tuple[float, float, float, float]] = []
        if isinstance(dropped, list):
            for bb in dropped:
                if isinstance(bb, (list, tuple)) and len(bb) >= 4:
                    try:
                        x0, y0, x1, y1 = (
                            float(bb[0]),
                            float(bb[1]),
                            float(bb[2]),
                            float(bb[3]),
                        )
                    except (TypeError, ValueError):
                        continue
                    if x1 <= x0 or y1 <= y0:
                        continue
                    drop_segs.extend(
                        [
                            (x0, y0, x1, y0),
                            (x1, y0, x1, y1),
                            (x1, y1, x0, y1),
                            (x0, y1, x0, y0),
                        ]
                    )
        add_multiline("prefilter_dropped_bboxes", drop_segs)

    add_rect("proposed_bbox", prealign.get("proposed_bbox"), dashed=True)
    add_rect("effective_bbox", prealign.get("effective_bbox"))
    add_rect("detected_union_bbox", prealign.get("detected_union_bbox"))

    pool_lines = prealign.get("prealign_detected_grid_lines") or []
    pool_segs: list[tuple[float, float, float, float]] = []
    detected_by_id: dict[str, dict[str, Any]] = {}
    for ln in pool_lines:
        if isinstance(ln, dict) and ln.get("id"):
            detected_by_id[str(ln["id"])] = ln
        seg = _line_from_detected(ln) if isinstance(ln, dict) else None
        if seg:
            pool_segs.append(seg)
    add_multiline("extent_probe_merged_grid", pool_segs)

    edge_colors = (
        ("top", "extent_boundary_top"),
        ("left", "extent_boundary_left"),
        ("bottom", "extent_boundary_bottom"),
        ("right", "extent_boundary_right"),
    )
    for edge, lid in edge_colors:
        payload = dict(selected.get(edge) or {})
        det_id = str(payload.get("id") or "")
        if not det_id:
            continue
        seg = _line_from_detected(detected_by_id.get(det_id, {}))
        add_line(lid, seg)

    if isinstance(search_bbox, list) and len(search_bbox) >= 4:
        sx0, sy0, sx1, sy1 = (float(search_bbox[0]), float(search_bbox[1]), float(search_bbox[2]), float(search_bbox[3]))
        top_v = virtual.get("top") or {}
        if top_v.get("virtual_pos_pts") not in ("", None):
            y = float(top_v["virtual_pos_pts"])
            add_line("virtual_top", (sx0, y, sx1, y))
        left_v = virtual.get("left") or {}
        if left_v.get("virtual_pos_pts") not in ("", None):
            x = float(left_v["virtual_pos_pts"])
            add_line("virtual_left", (x, sy0, x, sy1))

    diags = bundle.get("alignment_diagnostics") or []
    diag_by_line: dict[str, dict[str, Any]] = {}
    for d in diags:
        if isinstance(d, dict) and d.get("line_id"):
            diag_by_line[str(d["line_id"])] = d

    tpl_lines: list[dict[str, Any]] = [
        ln for ln in (bundle.get("grid_lines_template") or []) if isinstance(ln, dict)
    ]
    det_prod = bundle.get("detected_grid_lines_production") or []
    det_prod_by_id = {
        str(ln["id"]): ln
        for ln in det_prod
        if isinstance(ln, dict) and ln.get("id")
    }

    matched_segs: list[tuple[float, float, float, float]] = []
    matched_ids: set[str] = set()
    for ln in tpl_lines:
        if not ln.get("boundary"):
            continue
        lid = str(ln.get("id") or "")
        diag = diag_by_line.get(lid)
        if not diag:
            continue
        det_id = str(diag.get("matched_detected_id") or "")
        if det_id and det_id not in matched_ids:
            matched_ids.add(det_id)
            seg = _line_from_detected(det_prod_by_id.get(det_id, {}))
            if seg:
                matched_segs.append(seg)
    add_multiline("alignment_matched_detected_boundary", matched_segs)

    tpl_boundary_segs: list[tuple[float, float, float, float]] = []
    for ln in tpl_lines:
        if not ln.get("boundary"):
            continue
        lid = str(ln.get("id") or "")
        diag = diag_by_line.get(lid)
        snapped = diag.get("pos_snapped_pts") if diag else None
        if snapped is None:
            continue
        seg = _template_line_segment(ln, float(snapped))
        if seg:
            tpl_boundary_segs.append(seg)
    add_multiline("alignment_template_boundary", tpl_boundary_segs)

    items.sort(key=lambda it: (it.spec.pipeline_order if it.spec else 99))
    return items
