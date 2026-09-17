"""Registry of prealign / adapt debug overlay layers (editor + PNG naming)."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

GeometryKind = Literal["rect", "line", "multiline"]

# RGB 0-255 — aligned with adapt_debug PNG colors (fitz tuples scaled).
# Match adapt_debug._render_frame_and_search_bbox / proposed_vs_effective
_COLOR_FRAME = (0, 178, 0)
_COLOR_BASE = (25, 76, 229)
_COLOR_SEARCH = (255, 140, 0)
_COLOR_PROPOSED = (158, 31, 239)
_COLOR_EFFECTIVE = (209, 0, 153)
_COLOR_DETECTED_UNION = (120, 120, 120)
_COLOR_EXTENT_POOL = (140, 140, 140)
_COLOR_EXT_TOP = (229, 25, 25)
_COLOR_EXT_LEFT = (25, 76, 229)
_COLOR_EXT_BOTTOM = (0, 178, 0)
_COLOR_EXT_RIGHT = (158, 31, 239)
_COLOR_VIRTUAL_TOP = (209, 0, 153)
_COLOR_VIRTUAL_LEFT = (0, 158, 204)
_COLOR_ALIGN_DETECTED = (209, 0, 153)
_COLOR_ALIGN_TEMPLATE = (0, 178, 0)
_COLOR_LABEL = (30, 30, 30)
# Prefilter: editor-only overlay (png_group not rendered in adapt_debug PNG).
_COLOR_PREFILTER_BAND = (100, 149, 237)
_COLOR_PREFILTER_PRIOR = (200, 120, 60)
_COLOR_PREFILTER_DROPPED = (220, 90, 90)
_COLOR_PREFILTER_KEPT = (40, 170, 100)


@dataclass(frozen=True)
class PrealignLayerSpec:
    layer_id: str
    geometry: GeometryKind
    png_group: str
    label_full: str
    label_panel: str
    depends_on: tuple[str, ...]
    pipeline_order: int
    color_rgb: tuple[int, int, int]


PREALIGN_LAYER_SPECS: tuple[PrealignLayerSpec, ...] = (
    PrealignLayerSpec(
        "frame",
        "rect",
        "frame_and_search_bbox",
        "Рамка листа (find_frame): область формата страницы",
        "Рамка листа (find_frame)",
        (),
        1,
        _COLOR_FRAME,
    ),
    PrealignLayerSpec(
        "base_bbox",
        "rect",
        "frame_and_search_bbox",
        "Base bbox: объединение полей штампа в координатах PDF до расширения поиска",
        "Base bbox (каркас штампа)",
        ("frame",),
        2,
        _COLOR_BASE,
    ),
    PrealignLayerSpec(
        "search_bbox",
        "rect",
        "frame_and_search_bbox",
        "Search bbox: зона find_tables вокруг штампа (padding / overflow prealign)",
        "Search bbox (поиск таблицы)",
        ("base_bbox",),
        3,
        _COLOR_SEARCH,
    ),
    PrealignLayerSpec(
        "prefilter_y_anchor_band",
        "rect",
        "detected_prefilter_overlay",
        "Prefilter: якорная полоса по Y (расширенный stamp_bbox по высоте)",
        "Prefilter: полоса Y",
        ("search_bbox",),
        4,
        _COLOR_PREFILTER_BAND,
    ),
    PrealignLayerSpec(
        "prefilter_prior_rect",
        "rect",
        "detected_prefilter_overlay",
        "Prefilter: union полей штампа (t_all) для приоритета острова",
        "Prefilter: prior t_all",
        ("base_bbox",),
        5,
        _COLOR_PREFILTER_PRIOR,
    ),
    PrealignLayerSpec(
        "prefilter_dropped_bboxes",
        "multiline",
        "detected_prefilter_overlay",
        "Prefilter: отброшенные острова (контуры bbox компонент)",
        "Prefilter: отброшенные острова",
        ("search_bbox",),
        6,
        _COLOR_PREFILTER_DROPPED,
    ),
    PrealignLayerSpec(
        "prefilter_kept_bbox",
        "rect",
        "detected_prefilter_overlay",
        "Prefilter: выбранный остров (union bbox оставшихся ячеек)",
        "Prefilter: выбранный остров",
        ("prefilter_y_anchor_band",),
        7,
        _COLOR_PREFILTER_KEPT,
    ),
    PrealignLayerSpec(
        "extent_probe_merged_grid",
        "multiline",
        "extent_probe_boundaries",
        "Merged линии сетки внутри search (пул для extent_probe)",
        "Пул линий extent (merged)",
        ("search_bbox",),
        8,
        _COLOR_EXTENT_POOL,
    ),
    PrealignLayerSpec(
        "extent_boundary_top",
        "line",
        "extent_probe_boundaries",
        "Extent: выбранная граница top (detected horizontal boundary)",
        "Extent boundary top",
        ("extent_probe_merged_grid",),
        9,
        _COLOR_EXT_TOP,
    ),
    PrealignLayerSpec(
        "extent_boundary_left",
        "line",
        "extent_probe_boundaries",
        "Extent: выбранная граница left (detected vertical boundary)",
        "Extent boundary left",
        ("extent_probe_merged_grid",),
        10,
        _COLOR_EXT_LEFT,
    ),
    PrealignLayerSpec(
        "extent_boundary_bottom",
        "line",
        "extent_probe_boundaries",
        "Extent: выбранная граница bottom",
        "Extent boundary bottom",
        ("extent_probe_merged_grid",),
        11,
        _COLOR_EXT_BOTTOM,
    ),
    PrealignLayerSpec(
        "extent_boundary_right",
        "line",
        "extent_probe_boundaries",
        "Extent: выбранная граница right",
        "Extent boundary right",
        ("extent_probe_merged_grid",),
        12,
        _COLOR_EXT_RIGHT,
    ),
    PrealignLayerSpec(
        "virtual_top",
        "line",
        "extent_probe_boundaries",
        "Virtual top: виртуальная горизонталь (inferred edge) по ширине search_bbox",
        "Virtual top",
        ("extent_boundary_top", "search_bbox"),
        13,
        _COLOR_VIRTUAL_TOP,
    ),
    PrealignLayerSpec(
        "virtual_left",
        "line",
        "extent_probe_boundaries",
        "Virtual left: виртуальная вертикаль (inferred edge) по высоте search_bbox",
        "Virtual left",
        ("extent_boundary_left", "search_bbox"),
        14,
        _COLOR_VIRTUAL_LEFT,
    ),
    PrealignLayerSpec(
        "proposed_bbox",
        "rect",
        "proposed_vs_effective_bbox",
        "Proposed bbox: результат extent_probe (расширение/сужение до inferred границ)",
        "Proposed bbox",
        ("search_bbox",),
        15,
        _COLOR_PROPOSED,
    ),
    PrealignLayerSpec(
        "effective_bbox",
        "rect",
        "proposed_vs_effective_bbox",
        "Effective bbox: итоговый клип ячеек после proposed (фильтр matching)",
        "Effective bbox",
        ("proposed_bbox", "search_bbox"),
        16,
        _COLOR_EFFECTIVE,
    ),
    PrealignLayerSpec(
        "detected_union_bbox",
        "rect",
        "frame_and_search_bbox",
        "Detected union bbox: объединение detected-ячеек (probe/effective этап)",
        "Detected union bbox",
        ("search_bbox",),
        17,
        _COLOR_DETECTED_UNION,
    ),
    PrealignLayerSpec(
        "alignment_matched_detected_boundary",
        "multiline",
        "ordered_alignment_overlay",
        "Ordered alignment: matched PDF-линии для boundary шаблона (magenta в PNG)",
        "Alignment: matched detected (boundary)",
        ("effective_bbox",),
        18,
        _COLOR_ALIGN_DETECTED,
    ),
    PrealignLayerSpec(
        "alignment_template_boundary",
        "multiline",
        "ordered_alignment_overlay",
        "Ordered alignment: boundary линии шаблона в snapped позициях (green в PNG)",
        "Alignment: template boundary snapped",
        ("effective_bbox",),
        19,
        _COLOR_ALIGN_TEMPLATE,
    ),
)

LAYER_SPEC_BY_ID: dict[str, PrealignLayerSpec] = {s.layer_id: s for s in PREALIGN_LAYER_SPECS}
