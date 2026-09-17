"""
PDF parsing v2 — шаблонный движок извлечения штампа (fitz + JSON-шаблоны).

"""

from pdf_parsing_v2_engine.document import V2Document, V2PageData
from pdf_parsing_v2_engine.models import (
    FieldDef,
    FieldCatalog,
    FieldCatalogEntry,
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
from pdf_parsing_v2.v2_config import get_v2_config_path, load_v2_config
from pdf_parsing_v2.v2_pipeline import run_v2_pipeline
from pdf_parsing_v2.parallel import run_parallel, run_parallel_extraction
from pdf_parsing_v2.v2_timing import TimingCollector

__all__ = [
    "V2Document",
    "V2PageData",
    "FieldDef",
    "FieldCatalog",
    "FieldCatalogEntry",
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
    "get_v2_config_path",
    "load_v2_config",
    "run_v2_pipeline",
    "run_parallel",
    "run_parallel_extraction",
    "TimingCollector",
]
