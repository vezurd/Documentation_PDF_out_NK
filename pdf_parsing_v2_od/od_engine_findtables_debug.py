"""Engine-aligned manifest clip + find_tables probe for OD PDFs (diagnostic).

Uses the same per-page geometry as pdfplumber cropping when *geometry_rows* is
passed from :func:`~pdf_parsing_v2_od.od_parsing.od_table_parsing_f` (one
``extract_page`` pass per pipeline / monitor run).
"""

from __future__ import annotations

from os.path import isfile
from typing import Any

import fitz

from pdf_parsing_v2_engine.document import V2Document
from pdf_parsing_v2_engine.models import StampTemplate
from pdf_parsing_v2_od.od_manifest_clip import find_tables_od_manifest_region
from pdf_parsing_v2_od.od_manifest_geometry import (
    OdManifestGeometryRow,
    extract_od_manifest_geometry_rows,
)


_MAX_PAGES = 4


def collect_od_manifest_findtables_debug(
    od_pdf_path: str,
    templates: list[StampTemplate],
    cfg: dict[str, Any] | None,
    *,
    curr_proj: list[V2Document] | None = None,
    geometry_rows: list[OdManifestGeometryRow] | None = None,
) -> dict[str, Any]:
    """Run manifest clip + ``find_tables`` for first OD pages.

    When *geometry_rows* is provided (from ``od_table_parsing_f``), skips a second
    ``extract_page`` pass and uses those clips for MuPDF ``find_tables``.

    Args:
        od_pdf_path: Absolute path to the OD PDF.
        templates: Loaded stamp templates (same as extraction).
        cfg: Pipeline / v2 config dict.
        curr_proj: Optional documents from the same pipeline run (post-extraction).
        geometry_rows: Optional precomputed rows from
            :func:`~pdf_parsing_v2_od.od_manifest_geometry.extract_od_manifest_geometry_rows`.

    Returns:
        JSON-serializable dict with key ``pages``: list of per-page records
        (``page``, ``geometry_source`` ``cached_v2_result`` | ``extract_page``,
        ``clip``, ``n_tables``, ``n_cells``, ``template``, or ``skip`` / ``error``).
    """
    p = str(od_pdf_path or "").strip()
    if not p or not isfile(p) or not templates:
        return {}
    rows_src: list[OdManifestGeometryRow]
    if geometry_rows is not None and len(geometry_rows) > 0:
        rows_src = geometry_rows
    else:
        rows_src = extract_od_manifest_geometry_rows(
            p,
            templates,
            cfg,
            curr_proj=curr_proj,
            warnings_out=None,
            max_pages=_MAX_PAGES,
        )

    pages_out: list[dict[str, Any]] = []
    doc = fitz.open(p)
    try:
        for row in rows_src:
            page_num = row.page_num
            rec: dict[str, Any] = {"page": page_num}
            if page_num > doc.page_count or page_num < 1:
                rec["error"] = "page out of range"
                pages_out.append(rec)
                continue
            page = doc[page_num - 1]
            rec["geometry_source"] = row.geometry_source
            rec["template"] = row.template_name or None

            if row.frame is not None:
                rec["frame"] = [round(float(x), 3) for x in row.frame]
            if row.stamp_effective_bbox is not None:
                rec["stamp_effective_bbox"] = [
                    round(float(x), 3) for x in row.stamp_effective_bbox
                ]

            if row.skip_reason:
                rec["skip"] = row.skip_reason
                rec["clip"] = (
                    None
                    if row.clip is None
                    else [
                        round(float(row.clip[0]), 3),
                        round(float(row.clip[1]), 3),
                        round(float(row.clip[2]), 3),
                        round(float(row.clip[3]), 3),
                    ]
                )
                rec["n_tables"] = 0
                rec["n_cells"] = 0
                pages_out.append(rec)
                continue

            if row.clip is None:
                rec["skip"] = row.skip_reason or "no clip"
                rec["clip"] = None
                rec["n_tables"] = 0
                rec["n_cells"] = 0
                pages_out.append(rec)
                continue

            clip = fitz.Rect(row.clip[0], row.clip[1], row.clip[2], row.clip[3])
            rec["clip"] = [
                round(clip.x0, 3),
                round(clip.y0, 3),
                round(clip.x1, 3),
                round(clip.y1, 3),
            ]
            tmpl = next(
                (t for t in templates if t.name == row.template_name),
                None,
            )
            try:
                tbl = find_tables_od_manifest_region(page, clip, cfg, tmpl)
                tables = getattr(tbl, "tables", tbl) if tbl is not None else []
                if not isinstance(tables, list):
                    tables = list(tables) if tables else []
                n_cells = 0
                for t in tables:
                    cells = getattr(t, "cells", None) or []
                    n_cells += len(cells)
                rec["n_tables"] = len(tables)
                rec["n_cells"] = n_cells
            except Exception as exc:
                rec["error"] = str(exc)
            pages_out.append(rec)
    finally:
        doc.close()
    return {"pages": pages_out}
