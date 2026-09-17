"""
PDF v2 extraction engine (fitz + JSON templates).

Orchestration (folder pipeline, v1 post-steps) lives in ``pdf_parsing_v2``.
"""

from pdf_parsing_v2_engine.models import (
    FieldCatalog,
    FieldCatalogEntry,
    FieldDef,
    FieldResult,
    FrameInfo,
    StampTemplate,
    TemplateScore,
    TemplateSet,
    V2PageResult,
)
from pdf_parsing_v2_engine.stamp_extractor import (
    extract_all_pages_for_file,
    extract_page,
    select_templates,
)
from pdf_parsing_v2_engine.template_loader import (
    load_all_templates,
    load_catalog_for_project,
    load_project_templates,
    load_projects,
)

__all__ = [
    "FieldCatalog",
    "FieldCatalogEntry",
    "FieldDef",
    "FieldResult",
    "FrameInfo",
    "StampTemplate",
    "TemplateScore",
    "TemplateSet",
    "V2PageResult",
    "extract_all_pages_for_file",
    "extract_page",
    "select_templates",
    "load_all_templates",
    "load_catalog_for_project",
    "load_project_templates",
    "load_projects",
]
