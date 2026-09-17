"""
Движок извлечения полей штампа: один универсальный путь для всех типов документов.

v2 единый путь (заменяет 4 пути v1)::

    find_frame(fitz_page, …) → (FrameInfo, optional debug dict)
    → select_templates(doc_type, page_num) → отсортированные кандидаты
    → для каждого template:
        → для каждого field: bbox_mm + padding → fitz.Rect → get_textbox
        → clean_value → validate_regex
    → score_template → выбрать лучший
    → V2PageResult
"""

from __future__ import annotations

import os
import re
import time
from dataclasses import replace
from typing import TYPE_CHECKING, Any

import fitz

from pdf_parsing_v2_engine.char_text_extractor import PageCharIndex
from pdf_parsing_v2_engine.coord_transform import field_to_fitz_rect, fitz_displayed_to_unrotated
from pdf_parsing_v2_engine.frame_detector import find_frame
from pdf_parsing_v2_engine.grid_matcher import AdaptResult, adapt_by_cell_assignment
from pdf_parsing_v2_engine.models import (
    FieldDef,
    FieldResult,
    FrameInfo,
    StampTemplate,
    TemplateScore,
    V2PageResult,
)
from pdf_parsing_v2_engine.stamp_text.context import StampTextContext
from pdf_parsing_v2_engine.stamp_text.outcome import ParseWarning
from pdf_parsing_v2_engine.stamp_text.pipelines import (
    FIELD_CLEAN_PIPELINES,
    cleaner_exception_outcome,
    run_field_clean,
    unknown_cleaner_outcome,
)
from pdf_parsing_v2_engine.document_properties import (
    RESERVED_FIELD_IDS,
    build_document_properties_metadata,
)
from pdf_parsing_v2_engine.document_properties.registry import (
    KEY_FILE_NAME,
    KEY_PAGE_ANNOTATIONS,
    KEY_PAGE_HEIGHT_MM,
    KEY_PAGE_LAYERS,
    KEY_PAGE_REAL_FORMAT,
    KEY_PAGE_WIDTH_MM,
)

if TYPE_CHECKING:
    pass

MIN_SCORE_THRESHOLD: float = 0.3
_ABSENT_PENALTY: float = 0.3


def _timing_add(bucket: dict[str, float] | None, key: str, elapsed_sec: float) -> None:
    """Accumulate timing in a mutable dict when instrumentation is enabled."""
    if bucket is None:
        return
    bucket[key] = round(float(bucket.get(key, 0.0)) + float(elapsed_sec), 6)


def _source_basename_from_cfg(cfg: dict[str, Any] | None) -> str:
    if not cfg:
        return ""
    return str(cfg.get("source_pdf_basename", "") or "")


def _validate_template_reserved_field_ids(template: StampTemplate) -> None:
    """Reserved legacy ids must be document-property fields, not bbox stamp cells."""
    for f in template.fields:
        if f.id in RESERVED_FIELD_IDS and not f.document_property:
            raise ValueError(
                f"Template {template.name!r}: field id {f.id!r} is reserved for PDF document "
                "properties. Remove it or set \"document_property\" to the same id in JSON."
            )


def _stamp_grid_metadata_from_cell_info(
    ca_info: object,
    cfg: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Pickle/JSON-friendly stamp geometry from ``CellAssignmentInfo`` (displayed space).

    Keys (stable):
        ``stamp_effective_bbox``: ``tuple[float, float, float, float]`` — same coordinates
        as ``find_tables`` / field bboxes on the page.
        ``stamp_proposed_bbox_applied``: ``bool`` — whether proposed bbox replaced effective.
        ``stamp_search_bbox``: optional ``tuple[float, float, float, float]`` — expanded
        search region around the stamp (debug / coarse clip).
        ``stamp_grid_mismatch_score`` / ``stamp_grid_mismatch_breakdown_json`` — оценка разъезда сетки (монитор).
    """
    import json

    from pdf_parsing_v2_engine.grid_matcher import CellAssignmentInfo
    from pdf_parsing_v2_engine.stamp_grid_metrics import compute_stamp_grid_mismatch

    if not isinstance(ca_info, CellAssignmentInfo):
        return {}
    meta: dict[str, Any] = {}
    if ca_info.effective_bbox is not None:
        meta["stamp_effective_bbox"] = tuple(float(x) for x in ca_info.effective_bbox)
    meta["stamp_affine_scale_x"] = float(ca_info.transform_scale_x)
    meta["stamp_affine_scale_y"] = float(ca_info.transform_scale_y)
    score, breakdown = compute_stamp_grid_mismatch(ca_info, cfg)
    meta["stamp_grid_mismatch_score"] = int(score)
    meta["stamp_grid_mismatch_breakdown_json"] = json.dumps(breakdown, ensure_ascii=False)
    meta["stamp_proposed_bbox_applied"] = bool(ca_info.proposed_bbox_applied)
    if ca_info.search_bbox is not None:
        meta["stamp_search_bbox"] = tuple(float(x) for x in ca_info.search_bbox)
    return meta


def _attach_document_metadata(
    result: V2PageResult,
    fitz_page: fitz.Page,
    cfg: dict[str, Any] | None,
) -> V2PageResult:
    meta = build_document_properties_metadata(
        fitz_page,
        file_basename=_source_basename_from_cfg(cfg),
    )
    return replace(result, metadata={**result.metadata, **meta})


def _doc_prop_texts(property_key: str, raw_list: list[Any]) -> tuple[str, str]:
    """Build (raw_value, cleaned_value) strings for a document-property field."""
    if property_key in (KEY_PAGE_LAYERS, KEY_PAGE_ANNOTATIONS):
        v = int(raw_list[0]) if raw_list else -1
        s = str(v)
        return s, s
    if property_key in (KEY_PAGE_WIDTH_MM, KEY_PAGE_HEIGHT_MM):
        v = int(raw_list[0]) if raw_list else 0
        s = str(v)
        return s, s
    if property_key == KEY_PAGE_REAL_FORMAT:
        s = str(raw_list[0]) if raw_list else ""
        return s, s
    if property_key == KEY_FILE_NAME:
        s = str(raw_list[0]) if raw_list else ""
        return s, s
    return "", ""


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def extract_page(
    fitz_page: fitz.Page,
    doc_type: str,
    page_num: int,
    templates: list[StampTemplate],
    cfg: dict[str, Any] | None = None,
    timing: dict[str, float] | None = None,
) -> V2PageResult:
    """Extract stamp fields from *fitz_page* using the best matching template.

    Parameters
    ----------
    fitz_page : fitz.Page
        A page opened via ``fitz.open(path)[i]``.
    doc_type : str
        Document type code (``"WIR"``, ``"BOE"``, …).
    page_num : int
        1-based page number.
    templates : list[StampTemplate]
        All loaded templates; filtering by *doc_type* / *page_num* is done here.
    cfg : dict | None
        Optional v2 config dict.  Controls ``text_extraction_mode``:
        ``"char_center"`` (default) or ``"get_textbox"`` (legacy fallback).
        Forwarded to ``adapt_by_cell_assignment`` (e.g. ``find_tables`` tuning);
        optional keys ``_pdf_path``, ``_page_num`` for alignment diagnostics;
        ``source_pdf_basename`` for ``file_name`` in document-properties metadata.

    Returns
    -------
    V2PageResult
        Extracted fields, score, warnings.
    """
    extraction_mode: str = (cfg or {}).get("text_extraction_mode", "char_center")

    # Build char index once per page when using char_center mode.
    # The index stores chars in unrotated fitz coordinates (same space as
    # get_textbox expects), so query_rects can be passed to extract_text
    # without any additional conversion.
    char_index: PageCharIndex | None = None
    if extraction_mode == "char_center":
        t_char = time.perf_counter()
        try:
            tp = fitz_page.get_textpage(flags=fitz.TEXTFLAGS_TEXT | fitz.TEXT_MEDIABOX_CLIP)
            char_index = PageCharIndex.build(fitz_page, textpage=tp)
        except Exception:
            char_index = None  # fall back to get_textbox on error
        _timing_add(timing, "char_index_build", time.perf_counter() - t_char)

    t_select = time.perf_counter()
    candidates = select_templates(templates, doc_type, page_num)
    _timing_add(timing, "select_templates", time.perf_counter() - t_select)
    if not candidates:
        t_frame0 = time.perf_counter()
        frame0, _ = find_frame(fitz_page, timing=timing, cfg=cfg)
        _timing_add(timing, "find_frame", time.perf_counter() - t_frame0)
        return _attach_document_metadata(_empty_result(page_num, doc_type, frame0), fitz_page, cfg)

    scored: list[tuple[TemplateScore, dict[str, FieldResult], FrameInfo, dict[str, Any]]] = []
    for tmpl in candidates:
        _validate_template_reserved_field_ids(tmpl)
        t_union = tmpl if tmpl.frame_mode == "drawing_union" else None
        t_frame = time.perf_counter()
        frame, _ = find_frame(
            fitz_page,
            frame_mode=tmpl.frame_mode,
            template=t_union,
            timing=timing,
            cfg=cfg,
        )
        _timing_add(timing, "find_frame", time.perf_counter() - t_frame)
        fields, stamp_grid_meta = _extract_with_template(
            fitz_page,
            tmpl,
            frame,
            doc_type,
            page_num,
            char_index=char_index,
            cfg=cfg,
            timing=timing,
        )
        t_score = time.perf_counter()
        ts = _score_template(tmpl, fields)
        _timing_add(timing, "template_scoring", time.perf_counter() - t_score)
        scored.append((ts, fields, frame, stamp_grid_meta))

    scored.sort(key=lambda pair: pair[0].score, reverse=True)
    best_ts, best_fields, best_frame, best_stamp_meta = scored[0]

    warnings: list[str] = list(best_ts.details)
    if best_ts.score < MIN_SCORE_THRESHOLD:
        warnings.insert(
            0,
            f"Low template score {best_ts.score:.2f} (threshold {MIN_SCORE_THRESHOLD})",
        )

    t_build = time.perf_counter()
    result = V2PageResult(
        page_num=page_num,
        doc_type=doc_type,
        frame=best_frame,
        template_name=best_ts.template.name,
        template_score=best_ts.score,
        fields=best_fields,
        metadata=dict(best_stamp_meta),
        warnings=warnings,
    )
    _timing_add(timing, "build_page_result", time.perf_counter() - t_build)
    return _attach_document_metadata(result, fitz_page, cfg)


# ---------------------------------------------------------------------------
# Template selection
# ---------------------------------------------------------------------------

def select_templates(
    templates: list[StampTemplate],
    doc_type: str,
    page_num: int,
) -> list[StampTemplate]:
    """Return templates applicable to *(doc_type, page_num)*, sorted by priority desc."""
    out: list[StampTemplate] = []
    for t in templates:
        if doc_type not in t.doc_types:
            continue
        if t.page_selector == "first" and page_num != 1:
            continue
        if t.page_selector == "rest" and page_num == 1:
            continue
        out.append(t)
    out.sort(key=lambda t: t.priority, reverse=True)
    return out


# ---------------------------------------------------------------------------
# Per-template extraction
# ---------------------------------------------------------------------------

def _extract_with_template(
    fitz_page: fitz.Page,
    template: StampTemplate,
    frame: FrameInfo,
    doc_type: str,
    page_num: int,
    char_index: PageCharIndex | None = None,
    cfg: dict[str, Any] | None = None,
    timing: dict[str, float] | None = None,
) -> tuple[dict[str, FieldResult], dict[str, Any]]:
    _validate_template_reserved_field_ids(template)
    doc_props = build_document_properties_metadata(
        fitz_page,
        file_basename=_source_basename_from_cfg(cfg),
    )
    results: dict[str, FieldResult] = {}
    adapted_list: list[tuple[str, AdaptResult]] = []
    stamp_grid_meta: dict[str, Any] = {}
    if template.grid_adapt:
        t_adapt = time.perf_counter()
        try:
            adapted_list, ca_info = adapt_by_cell_assignment(
                template=template,
                frame=frame,
                fitz_page=fitz_page,
                cfg=cfg,
                timing=timing,
            )
            stamp_grid_meta = _stamp_grid_metadata_from_cell_info(ca_info, cfg)
            if cfg and cfg.get("editor_prealign_overlay"):
                try:
                    from pdf_template_editor.adapt_debug import stamp_prealign_debug_bundle

                    stamp_grid_meta["stamp_prealign_debug"] = stamp_prealign_debug_bundle(
                        frame=frame,
                        template=template,
                        ca_info=ca_info,
                        fitz_page=fitz_page,
                        cfg=cfg,
                    )
                except Exception:
                    pass
        except Exception:
            adapted_list = []
        _timing_add(timing, "grid_adapt", time.perf_counter() - t_adapt)

    for idx, field in enumerate(template.fields):
        if field.document_property:
            prop_key = field.document_property
            raw_list = doc_props.get(prop_key)
            if not isinstance(raw_list, list):
                raw_list = []
            raw_s, cleaned_s = _doc_prop_texts(prop_key, raw_list)
            results[field.id] = FieldResult(
                field_id=field.id,
                raw_value=raw_s,
                cleaned_value=cleaned_s,
                bbox_pts=(0.0, 0.0, 0.0, 0.0),
                is_valid=None,
                parse_warnings=[],
                clean_tier=None,
            )
            continue

        if template.grid_adapt and idx < len(adapted_list):
            _key, ar = adapted_list[idx]
            fitz_rect = fitz.Rect(ar.bbox.x0, ar.bbox.y0, ar.bbox.x1, ar.bbox.y1)
        else:
            effective_padding = (
                field.padding_mm if field.padding_mm is not None else template.padding_mm
            )
            fitz_rect = field_to_fitz_rect(field, frame, padding_mm=effective_padding)

        raw_text: str = ""
        if not fitz_rect.is_empty and not fitz_rect.is_infinite:
            t_text = time.perf_counter()
            # get_textbox() and PageCharIndex both operate in *unrotated* fitz
            # coordinates for rotation=90/270 pages, while fitz_rect is in
            # displayed space — convert once for both paths.
            query_rect = fitz_displayed_to_unrotated(fitz_rect, fitz_page)
            if char_index is not None:
                raw_text = char_index.extract_text(
                    query_rect, frame=frame, fitz_page=fitz_page
                )
            else:
                raw_text = fitz_page.get_textbox(query_rect)
            _timing_add(timing, "extract_fields_text", time.perf_counter() - t_text)

        ctx = StampTextContext(doc_type=doc_type, page_num=page_num, field_id=field.id)
        parse_warnings: list[ParseWarning] = []
        clean_tier: str | None = None
        if field.clean:
            t_clean = time.perf_counter()
            if field.clean not in FIELD_CLEAN_PIPELINES:
                outcome = unknown_cleaner_outcome(field.clean, raw_text)
            else:
                try:
                    outcome = run_field_clean(field.clean, raw_text, ctx)
                except Exception as exc:
                    outcome = cleaner_exception_outcome(raw_text, exc)
            cleaned = outcome.value
            parse_warnings = list(outcome.warnings)
            clean_tier = outcome.clean_tier
            _timing_add(timing, "clean_fields", time.perf_counter() - t_clean)
        else:
            cleaned = raw_text

        if isinstance(cleaned, list):
            cleaned = " ".join(str(x) for x in cleaned) if cleaned else ""
        elif cleaned is not None and not isinstance(cleaned, str):
            cleaned = str(cleaned)

        is_valid: bool | None = None
        if field.validate_regex and cleaned:
            is_valid = bool(re.search(field.validate_regex, cleaned))

        bbox_pts = (fitz_rect.x0, fitz_rect.y0, fitz_rect.x1, fitz_rect.y1)
        results[field.id] = FieldResult(
            field_id=field.id,
            raw_value=raw_text,
            cleaned_value=cleaned,
            bbox_pts=bbox_pts,
            is_valid=is_valid,
            parse_warnings=parse_warnings,
            clean_tier=clean_tier,
        )
    return results, stamp_grid_meta


# ---------------------------------------------------------------------------
# Scoring
# ---------------------------------------------------------------------------

def _score_template(
    template: StampTemplate,
    results: dict[str, FieldResult],
) -> TemplateScore:
    total_required = sum(1 for f in template.fields if f.expected == "required")
    required_found = 0
    absent_violations = 0
    details: list[str] = []
    score = 1.0

    for field in template.fields:
        if field.document_property:
            continue
        result = results.get(field.id)
        has_value = bool(
            result
            and result.cleaned_value
            and result.cleaned_value.strip()
        )

        if field.expected == "required":
            if has_value:
                required_found += 1
                if result and result.is_valid is False:
                    penalty = 0.5 / max(total_required, 1)
                    score -= penalty
                    details.append(f"{field.id}: required, present but invalid (−{penalty:.2f})")
            else:
                penalty = 1.0 / max(total_required, 1)
                score -= penalty
                details.append(f"{field.id}: required but empty (−{penalty:.2f})")

        elif field.expected == "absent":
            if has_value:
                absent_violations += 1
                score -= _ABSENT_PENALTY
                details.append(
                    f"{field.id}: expected absent but has data '{result.cleaned_value[:30]}…' "
                    f"(−{_ABSENT_PENALTY})"
                )

    return TemplateScore(
        template=template,
        score=max(score, 0.0),
        required_found=required_found,
        required_total=total_required,
        absent_violations=absent_violations,
        details=details,
    )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _empty_result(
    page_num: int,
    doc_type: str,
    frame: FrameInfo,
) -> V2PageResult:
    return V2PageResult(
        page_num=page_num,
        doc_type=doc_type,
        frame=frame,
        template_name="(no matching template)",
        template_score=0.0,
        fields={},
        metadata={},
        warnings=[f"No template found for doc_type={doc_type!r}, page={page_num}"],
    )


def extract_all_pages_for_file(
    file_path: str,
    doc_type: str,
    templates: list[StampTemplate],
    cfg: dict[str, Any],
    *,
    with_timing: bool = False,
) -> list[V2PageResult] | tuple[list[V2PageResult], dict[str, float]]:
    """Process one PDF file — pure function, safe for PPE.

    Opens and closes fitz document internally.
    Sets ``cfg["source_pdf_basename"]`` from *file_path*.
    No global state, no side effects beyond file I/O (read-only).

    All arguments and return values are pickle-safe
    (no fitz objects in V2PageResult).
    """
    results: list[V2PageResult] = []
    file_timing: dict[str, float] | None = {} if with_timing else None
    t_open = time.perf_counter()
    fitz_doc = fitz.open(file_path)
    _timing_add(file_timing, "open_pdf", time.perf_counter() - t_open)
    try:
        page_cfg = dict(cfg) if cfg else {}
        page_cfg["source_pdf_basename"] = os.path.basename(file_path)
        t_pages = time.perf_counter()
        for i in range(len(fitz_doc)):
            fitz_page = fitz_doc[i]
            page_num = i + 1
            page_timing: dict[str, float] | None = {} if with_timing else None
            v2_result = extract_page(
                fitz_page,
                doc_type,
                page_num,
                templates,
                cfg=page_cfg,
                timing=page_timing,
            )
            results.append(v2_result)
            if page_timing:
                for key, value in page_timing.items():
                    _timing_add(file_timing, key, value)
        _timing_add(file_timing, "page_loop_total", time.perf_counter() - t_pages)
    finally:
        fitz_doc.close()
    if with_timing:
        return results, file_timing or {}
    return results
