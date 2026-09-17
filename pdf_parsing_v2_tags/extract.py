"""Tag extraction from PDF via pdfminer / fitz for the v2 pipeline.

Public API: ``parse_tags(curr_proj, tag_dict=None)``.
"""

from __future__ import annotations

import os
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from typing import Any

import fitz
from pdfminer.high_level import extract_pages
from pdfminer.layout import LAParams, LTTextContainer

from pdf_parsing_v2_engine.doc_types import LIST_OF_BBB, LIST_OF_TEXT_DOC_TYPES
from pdf_parsing_v2_engine.document import V2Document
from tags import tag_parser


def _timing_add(timing: dict[str, float], key: str, elapsed_sec: float) -> None:
    timing[key] = round(float(timing.get(key, 0.0)) + elapsed_sec, 6)


def _default_tag_workers() -> int:
    cpus = os.cpu_count() or 4
    return max(1, min(cpus, 8))


def _resolve_tag_workers(total: int, max_workers: int | None) -> int:
    if total <= 0:
        return 1
    if max_workers in (None, 0):
        return min(_default_tag_workers(), total)
    return max(1, min(int(max_workers), total))


def _parse_tags_for_document(
    file_path: str,
    doc_type: str,
    doc_od_style_file_name: str,
    backend: str = "fitz",
) -> tuple[dict, str, dict[str, float | int]]:
    """Parse tags for one PDF file and return a local tag dict."""
    local_tag_dict: dict = {}
    result = "None"
    detail: dict[str, float | int] = {
        "tag_backend": backend,
        "tag_pages": 0,
        "tag_page_objects": 0,
        "tag_text_containers": 0,
        "tag_text_chars": 0,
        "tag_containers_with_tags": 0,
        "tag_total_found": 0,
    }

    if backend == "fitz":
        t_extract = time.perf_counter()
        with fitz.open(file_path) as doc:
            pages_text = [
                page.get_text("text", sort=True)
                for page in doc
            ]
        _timing_add(detail, "tag_extract_pages", time.perf_counter() - t_extract)
        is_text_type = doc_type in LIST_OF_BBB or doc_type in LIST_OF_TEXT_DOC_TYPES
        for page_num, page_text in enumerate(pages_text):
            detail["tag_pages"] += 1
            detail["tag_text_chars"] += len(page_text)
            if is_text_type:
                t_add = time.perf_counter()
                before_count = sum(len(v) for v in local_tag_dict.values())
                result = tag_parser.add_context(
                    page_text,
                    doc_od_style_file_name,
                    page_num,
                    target_dict=local_tag_dict,
                    file_path=file_path,
                )
                _timing_add(detail, "tag_add_context", time.perf_counter() - t_add)
                after_count = sum(len(v) for v in local_tag_dict.values())
                detail["tag_total_found"] += max(0, after_count - before_count)
            else:
                t_add = time.perf_counter()
                temp_dict: dict = {}
                tag_parser.add_context(
                    page_text,
                    doc_od_style_file_name,
                    page_num,
                    target_dict=temp_dict,
                    file_path=file_path,
                )
                _timing_add(detail, "tag_add_context", time.perf_counter() - t_add)
                if temp_dict:
                    detail["tag_containers_with_tags"] += 1
                    tags_in_container = sum(len(v) for v in temp_dict.values())
                    detail["tag_total_found"] += tags_in_container
                    tag_parser.merge_dicts(local_tag_dict, temp_dict)
        return local_tag_dict, result, detail

    t_extract = time.perf_counter()
    pages = list(
        extract_pages(
            file_path,
            laparams=LAParams(detect_vertical=True),
        )
    )
    _timing_add(detail, "tag_extract_pages", time.perf_counter() - t_extract)
    is_text_type = doc_type in LIST_OF_BBB or doc_type in LIST_OF_TEXT_DOC_TYPES
    for page_num, page in enumerate(pages):
        detail["tag_pages"] += 1
        detail["tag_page_objects"] += len(page._objs)
        t_sort = time.perf_counter()
        page_elements = sorted(
            ((el.y1, el) for el in page._objs),
            key=lambda a: a[0],
            reverse=True,
        )
        _timing_add(detail, "tag_sort_page_elements", time.perf_counter() - t_sort)

        text_for_tags = ""
        t_walk = time.perf_counter()
        for _, element in page_elements:
            if not isinstance(element, LTTextContainer):
                continue
            detail["tag_text_containers"] += 1
            t_text = time.perf_counter()
            element_text = element.get_text()
            _timing_add(detail, "tag_get_text", time.perf_counter() - t_text)
            detail["tag_text_chars"] += len(element_text)
            if is_text_type:
                text_for_tags += element_text
            else:
                t_add = time.perf_counter()
                temp_dict: dict = {}
                tag_parser.add_context(
                    element_text,
                    doc_od_style_file_name,
                    page_num,
                    target_dict=temp_dict,
                    file_path=file_path,
                )
                _timing_add(detail, "tag_add_context", time.perf_counter() - t_add)
                if temp_dict:
                    detail["tag_containers_with_tags"] += 1
                    tags_in_container = sum(len(v) for v in temp_dict.values())
                    detail["tag_total_found"] += tags_in_container
                    tag_parser.merge_dicts(local_tag_dict, temp_dict)
        _timing_add(detail, "tag_walk_layout", time.perf_counter() - t_walk)

        if is_text_type:
            t_add = time.perf_counter()
            before_count = sum(len(v) for v in local_tag_dict.values())
            result = tag_parser.add_context(
                text_for_tags,
                doc_od_style_file_name,
                page_num,
                target_dict=local_tag_dict,
                file_path=file_path,
            )
            _timing_add(detail, "tag_add_context", time.perf_counter() - t_add)
            after_count = sum(len(v) for v in local_tag_dict.values())
            detail["tag_total_found"] += max(0, after_count - before_count)
    return local_tag_dict, result, detail


def parse_tags(
    curr_proj: list[V2Document],
    tag_dict: dict | None = None,
    callback: Any | None = None,
    timing: Any | None = None,
    *,
    max_workers: int | None = None,
    backend: str = "fitz",
) -> str:
    """Extract equipment tags from PDF documents via pdfminer.

    Args:
        curr_proj: list of document descriptors with ``file_full_path``,
            ``doc_Type``, ``doc_OD_style_file_name``.
        tag_dict: when ``None`` — resets and writes to the global
            ``tag_parser.source_dict``.  When a dict is passed — writes
            into it (compatible with per-file map-reduce for future PPE).
        callback: optional per-file progress callback with
            ``on_tag_file_start/done/error`` methods.
        timing: optional timing collector; records one stage row per file.
        max_workers: worker count for per-file parallel tag parsing;
            ``0``/``None`` means auto by CPU count.
        backend: tag text backend, ``pdfminer`` or ``fitz``.

    Returns:
        Result string from the last ``add_context`` call (legacy).
    """
    result = "None"
    if tag_dict is None:
        tag_parser.reset()
        target_dict = tag_parser.source_dict
    else:
        target_dict = tag_dict

    total = len(curr_proj)
    workers = _resolve_tag_workers(total, max_workers)
    if total < 2 or workers <= 1:
        for idx, document in enumerate(curr_proj):
            file_path = document.file_full_path
            if callback is not None:
                cb_start = getattr(callback, "on_tag_file_start", None)
                if callable(cb_start):
                    cb_start(file_path, idx, total)
            t_file = time.perf_counter()
            try:
                local_dict, result, detail = _parse_tags_for_document(
                    file_path,
                    document.doc_Type,
                    document.doc_OD_style_file_name,
                    backend=backend,
                )
                tag_parser.merge_dicts(target_dict, local_dict)
            except Exception as e:
                elapsed = time.perf_counter() - t_file
                if timing is not None:
                    timing.record_stage(
                        "tags",
                        "parse_tags_file",
                        elapsed,
                        status="error",
                        file_name=os.path.basename(file_path),
                        error=str(e),
                    )
                if callback is not None:
                    cb_error = getattr(callback, "on_tag_file_error", None)
                    if callable(cb_error):
                        cb_error(file_path, idx, e)
                raise
            else:
                elapsed = time.perf_counter() - t_file
                if timing is not None:
                    timing.record_stage(
                        "tags",
                        "parse_tags_file",
                        elapsed,
                        file_name=os.path.basename(file_path),
                        **detail,
                    )
                if callback is not None:
                    cb_done = getattr(callback, "on_tag_file_done", None)
                    if callable(cb_done):
                        cb_done(file_path, idx, elapsed)
        return result

    future_to_info: dict[Any, tuple[int, V2Document, str, float]] = {}
    with ProcessPoolExecutor(max_workers=workers) as executor:
        for idx, document in enumerate(curr_proj):
            file_path = document.file_full_path
            if callback is not None:
                cb_start = getattr(callback, "on_tag_file_start", None)
                if callable(cb_start):
                    cb_start(file_path, idx, total)
            t_file = time.perf_counter()
            future = executor.submit(
                _parse_tags_for_document,
                file_path,
                document.doc_Type,
                document.doc_OD_style_file_name,
                backend,
            )
            future_to_info[future] = (idx, document, file_path, t_file)

        for future in as_completed(future_to_info):
            idx, document, file_path, t_file = future_to_info[future]
            del document
            try:
                local_dict, local_result, detail = future.result()
                tag_parser.merge_dicts(target_dict, local_dict)
                result = local_result
            except Exception as e:
                elapsed = time.perf_counter() - t_file
                if timing is not None:
                    timing.record_stage(
                        "tags",
                        "parse_tags_file",
                        elapsed,
                        status="error",
                        file_name=os.path.basename(file_path),
                        error=str(e),
                    )
                if callback is not None:
                    cb_error = getattr(callback, "on_tag_file_error", None)
                    if callable(cb_error):
                        cb_error(file_path, idx, e)
                raise
            else:
                elapsed = time.perf_counter() - t_file
                if timing is not None:
                    timing.record_stage(
                        "tags",
                        "parse_tags_file",
                        elapsed,
                        file_name=os.path.basename(file_path),
                        **detail,
                    )
                if callback is not None:
                    cb_done = getattr(callback, "on_tag_file_done", None)
                    if callable(cb_done):
                        cb_done(file_path, idx, elapsed)

    return result
