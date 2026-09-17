"""Shared OD manifest geometry: extract_page + frame/stamp clip for pdfplumber / diagnostics."""

from __future__ import annotations

import os
from dataclasses import dataclass
from os.path import isfile
from typing import Any, Literal

import fitz

from pdf_parsing_v2_engine.document import V2Document
from pdf_parsing_v2_engine.models import StampTemplate, V2PageResult
from pdf_parsing_v2_engine.stamp_extractor import extract_page
from pdf_parsing_v2_od.od_manifest_clip import build_manifest_clip_frame_minus_stamp

_MAX_PAGES = 4


def _norm_abs(path: str) -> str:
    try:
        return os.path.normcase(os.path.normpath(os.path.abspath(path)))
    except OSError:
        return os.path.normcase(os.path.normpath(path))


def _match_od_document(od_pdf_path: str, curr_proj: list[V2Document] | None) -> V2Document | None:
    if not curr_proj:
        return None
    target = _norm_abs(od_pdf_path)
    for doc in curr_proj:
        if _norm_abs(doc.file_full_path) == target:
            return doc
    return None


def _stamp_bbox_from_metadata(meta: dict[str, Any]) -> tuple[float, float, float, float] | None:
    seb = meta.get("stamp_effective_bbox")
    if not isinstance(seb, (list, tuple)) or len(seb) != 4:
        return None
    try:
        return (float(seb[0]), float(seb[1]), float(seb[2]), float(seb[3]))
    except (TypeError, ValueError):
        return None


@dataclass(frozen=True)
class OdManifestGeometryRow:
    """One page’s manifest geometry for OD (displayed / pdfplumber space)."""

    page_num: int
    geometry_source: Literal["extract_page", "cached_v2_result"]
    template_name: str
    clip: tuple[float, float, float, float] | None
    frame: tuple[float, float, float, float] | None
    stamp_effective_bbox: tuple[float, float, float, float] | None
    skip_reason: str | None


def extract_od_manifest_geometry_rows(
    od_pdf_path: str,
    templates: list[StampTemplate],
    cfg: dict[str, Any] | None,
    *,
    curr_proj: list[V2Document] | None = None,
    warnings_out: list[str] | None = None,
    max_pages: int = _MAX_PAGES,
) -> list[OdManifestGeometryRow]:
    """Run ``extract_page`` (or reuse cached ``V2PageResult``) and compute manifest clip per page.

    Args:
        od_pdf_path: Path to the OD PDF.
        templates: Stamp templates (same as extraction).
        cfg: v2 config dict.
        curr_proj: Optional documents from the same pipeline run to reuse cached geometry.
        warnings_out: Optional list for short diagnostics (missing bbox, empty clip).
        max_pages: Upper bound for pages (aligned with ``od_read_raw``).

    Returns:
        One :class:`OdManifestGeometryRow` per page ``1 .. min(page_count, max_pages)``.
        ``clip`` is ``None`` when the caller should use the **full page** for pdfplumber.
    """
    p = str(od_pdf_path or "").strip()
    out: list[OdManifestGeometryRow] = []
    if not p or not isfile(p) or not templates:
        return out

    cached_doc = _match_od_document(p, curr_proj)
    doc = fitz.open(p)
    try:
        n = min(doc.page_count, max_pages)
        for i in range(n):
            page = doc[i]
            page_num = i + 1
            geometry_source: Literal["extract_page", "cached_v2_result"] = "extract_page"
            result: V2PageResult | None = None
            try:
                if cached_doc is not None:
                    cached = cached_doc.get_v2_page_result(page_num)
                    if cached is not None and _stamp_bbox_from_metadata(cached.metadata or {}):
                        result = cached
                        geometry_source = "cached_v2_result"
                if result is None:
                    result = extract_page(page, "OD", page_num, templates, cfg)

                f = result.frame
                frame_rect = fitz.Rect(f.x0, f.y0, f.x1, f.y1)
                frame_t = (frame_rect.x0, frame_rect.y0, frame_rect.x1, frame_rect.y1)
                meta = result.metadata or {}
                seb = _stamp_bbox_from_metadata(meta)
                tmpl = next(
                    (t for t in templates if t.name == result.template_name),
                    None,
                )
                origin = tmpl.origin if tmpl is not None else "frame_bottom_right"
                template_name = result.template_name

                if seb is None:
                    msg = (
                        f"Страница {page_num}: нет stamp_effective_bbox в metadata — "
                        "pdfplumber без обрезки по штампу."
                    )
                    if warnings_out is not None:
                        warnings_out.append(msg)
                    out.append(
                        OdManifestGeometryRow(
                            page_num=page_num,
                            geometry_source=geometry_source,
                            template_name=template_name,
                            clip=None,
                            frame=frame_t,
                            stamp_effective_bbox=None,
                            skip_reason="no stamp_effective_bbox in metadata",
                        )
                    )
                    continue

                stamp_rect = fitz.Rect(seb[0], seb[1], seb[2], seb[3])
                page_rect = fitz.Rect(0.0, 0.0, page.rect.width, page.rect.height)
                clip_rect = build_manifest_clip_frame_minus_stamp(
                    frame_rect=frame_rect,
                    stamp_rect=stamp_rect,
                    page_rect=page_rect,
                    origin=origin,
                )
                stamp_t = (float(seb[0]), float(seb[1]), float(seb[2]), float(seb[3]))

                if clip_rect is None or clip_rect.is_empty:
                    msg = (
                        f"Страница {page_num}: пустой manifest clip — pdfplumber без обрезки."
                    )
                    if warnings_out is not None:
                        warnings_out.append(msg)
                    out.append(
                        OdManifestGeometryRow(
                            page_num=page_num,
                            geometry_source=geometry_source,
                            template_name=template_name,
                            clip=None,
                            frame=frame_t,
                            stamp_effective_bbox=stamp_t,
                            skip_reason="empty_manifest_clip",
                        )
                    )
                    continue

                ct = (clip_rect.x0, clip_rect.y0, clip_rect.x1, clip_rect.y1)
                out.append(
                    OdManifestGeometryRow(
                        page_num=page_num,
                        geometry_source=geometry_source,
                        template_name=template_name,
                        clip=ct,
                        frame=frame_t,
                        stamp_effective_bbox=stamp_t,
                        skip_reason=None,
                    )
                )
            except Exception as exc:
                msg = f"Страница {page_num}: ошибка геометрии ОД ({exc})."
                if warnings_out is not None:
                    warnings_out.append(msg)
                out.append(
                    OdManifestGeometryRow(
                        page_num=page_num,
                        geometry_source="extract_page",
                        template_name="",
                        clip=None,
                        frame=None,
                        stamp_effective_bbox=None,
                        skip_reason=str(exc),
                    )
                )
    finally:
        doc.close()
    return out
