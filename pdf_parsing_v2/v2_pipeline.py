"""
v2 pipeline: обработка папки PDF через шаблонный движок (fitz + JSON-шаблоны).

Запуск из командной строки::

    python pdf_parsing_v2/v2_pipeline.py "C:\\path\\to\\PDF"
    python pdf_parsing_v2/v2_pipeline.py "C:\\path\\to\\PDF" "path/to/pdf_v2_config.json"
    python pdf_parsing_v2/v2_pipeline.py "C:\\path\\to\\PDF" --mto-regression
    python pdf_parsing_v2/v2_pipeline.py "C:\\path\\to\\PDF" --mto-regression --mto-regression-strict

Опционально после прогона: ``--mto-regression`` (или ``mto_regression_after_run`` в JSON-конфиге) —
см. ``pdf_parsing_v2/tools/mto_page1_batch_probe.py``.

Последняя непустая строка stdout = путь к папке результатов (для GUI).
"""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
import os
import sys
import threading
import time
import traceback
from typing import Any, Protocol

from pdf_parsing_v2_engine.models import StampTemplate, V2PageResult
from pdf_parsing_v2_engine.stamp_extractor import extract_all_pages_for_file
from pdf_parsing_v2_engine.template_loader import (
    load_all_templates,
    load_catalog_for_project,
    load_project_templates,
)
from pdf_parsing_v2.v2_config import (
    excel_sanitize_illegal_chars_enabled,
    load_v2_config,
    resolve_templates_dir,
)

from pdf_parsing_v2_engine.doc_types import uncovered_doc_type_files
from pdf_parsing_v2_engine.document import V2Document, V2PageData

from pdf_parsing_v2.stamp_affine_summary import (
    stamp_affine_detail_for_ui,
    stamp_affine_uniformity_from_pages,
)


class ExtractionCallback(Protocol):
    """Optional progress reporting for extraction phase."""

    def on_file_start(self, file_path: str, index: int, total: int) -> None: ...

    def on_file_done(
        self,
        file_path: str,
        index: int,
        n_pages: int,
        elapsed_sec: float,
        file_detail: dict[str, Any] | None = None,
    ) -> None: ...

    def on_file_error(self, file_path: str, index: int, error: Exception) -> None: ...


class TagParseCallback(Protocol):
    """Optional progress reporting for per-file tag parsing."""

    def on_tag_file_start(self, file_path: str, index: int, total: int) -> None: ...

    def on_tag_file_done(
        self, file_path: str, index: int, elapsed_sec: float
    ) -> None: ...

    def on_tag_file_error(
        self, file_path: str, index: int, error: Exception
    ) -> None: ...


class PipelineCallback(ExtractionCallback, TagParseCallback, Protocol):
    """Optional progress reporting for the whole pipeline."""

    def on_batch_progress(self, completed: int, total: int) -> None: ...

    def on_phase_changed(self, phase: str) -> None: ...

    def on_stage_start(
        self,
        category: str,
        stage: str,
        file_name: str | None = None,
        detail: dict[str, Any] | None = None,
    ) -> None: ...

    def on_stage_done(
        self,
        category: str,
        stage: str,
        elapsed_sec: float,
        file_name: str | None = None,
        detail: dict[str, Any] | None = None,
    ) -> None: ...

    def on_stage_error(
        self,
        category: str,
        stage: str,
        error: Exception,
        file_name: str | None = None,
        detail: dict[str, Any] | None = None,
    ) -> None: ...

    def on_pipeline_finished(
        self,
        result_dir: str,
        total_elapsed: float,
        summary: dict[str, Any],
    ) -> None: ...


class _HeavyProgressProxy:
    """Forward callbacks and aggregate heavy-task progress.

    Heavy tasks = per-file extraction + per-file tag parsing.
    """

    def __init__(self, callback: Any, total_files: int) -> None:
        self._callback = callback
        self._total_tasks = max(total_files * 2, 1)
        self._completed_tasks = 0
        self._lock = threading.Lock()

    def _forward(self, name: str, *args: Any) -> None:
        method = _callback_method(self._callback, name)
        if method is not None:
            method(*args)

    def _mark_task_done(self) -> None:
        with self._lock:
            self._completed_tasks += 1
            completed = self._completed_tasks
        self._forward("on_batch_progress", completed, self._total_tasks)

    def on_file_start(self, file_path: str, index: int, total: int) -> None:
        self._forward("on_file_start", file_path, index, total)

    def on_file_done(
        self,
        file_path: str,
        index: int,
        n_pages: int,
        elapsed_sec: float,
        file_detail: dict[str, Any] | None = None,
    ) -> None:
        self._forward(
            "on_file_done",
            file_path,
            index,
            n_pages,
            elapsed_sec,
            file_detail,
        )
        self._mark_task_done()

    def on_file_error(self, file_path: str, index: int, error: Exception) -> None:
        self._forward("on_file_error", file_path, index, error)
        self._mark_task_done()

    def on_tag_file_start(self, file_path: str, index: int, total: int) -> None:
        self._forward("on_tag_file_start", file_path, index, total)

    def on_tag_file_done(self, file_path: str, index: int, elapsed_sec: float) -> None:
        self._forward("on_tag_file_done", file_path, index, elapsed_sec)
        self._mark_task_done()

    def on_tag_file_error(self, file_path: str, index: int, error: Exception) -> None:
        self._forward("on_tag_file_error", file_path, index, error)
        self._mark_task_done()

    def on_batch_progress(self, completed: int, total: int) -> None:
        """Swallow extraction-only progress; this proxy emits aggregate progress."""
        return None

    def on_phase_changed(self, phase: str) -> None:
        self._forward("on_phase_changed", phase)

    def on_stage_start(
        self,
        category: str,
        stage: str,
        file_name: str | None = None,
        detail: dict[str, Any] | None = None,
    ) -> None:
        self._forward("on_stage_start", category, stage, file_name, detail or {})

    def on_stage_done(
        self,
        category: str,
        stage: str,
        elapsed_sec: float,
        file_name: str | None = None,
        detail: dict[str, Any] | None = None,
    ) -> None:
        self._forward(
            "on_stage_done",
            category,
            stage,
            elapsed_sec,
            file_name,
            detail or {},
        )

    def on_stage_error(
        self,
        category: str,
        stage: str,
        error: Exception,
        file_name: str | None = None,
        detail: dict[str, Any] | None = None,
    ) -> None:
        self._forward(
            "on_stage_error",
            category,
            stage,
            error,
            file_name,
            detail or {},
        )

    def on_pipeline_finished(
        self,
        result_dir: str,
        total_elapsed: float,
        summary: dict[str, Any],
    ) -> None:
        self._forward("on_pipeline_finished", result_dir, total_elapsed, summary)


def _callback_method(callback: Any, name: str) -> Any | None:
    if callback is None:
        return None
    method = getattr(callback, name, None)
    return method if callable(method) else None


def _notify_phase(callback: Any, phase: str) -> None:
    method = _callback_method(callback, "on_phase_changed")
    if method is not None:
        method(phase)


def _notify_stage_start(
    callback: Any,
    category: str,
    stage: str,
    *,
    file_name: str | None = None,
    detail: dict[str, Any] | None = None,
) -> None:
    method = _callback_method(callback, "on_stage_start")
    if method is not None:
        method(category, stage, file_name, detail or {})


def _notify_stage_done(
    callback: Any,
    category: str,
    stage: str,
    elapsed_sec: float,
    *,
    file_name: str | None = None,
    detail: dict[str, Any] | None = None,
) -> None:
    method = _callback_method(callback, "on_stage_done")
    if method is not None:
        method(category, stage, elapsed_sec, file_name, detail or {})


def _uncovered_doc_type_payload(
    documents: list[V2Document],
    templates: list[StampTemplate],
) -> dict[str, Any]:
    """Build callback/summary payload for files whose type has no stamp template."""
    pairs = [
        (os.path.basename(str(doc.file_full_path or "")), str(doc.doc_Type or ""))
        for doc in documents
    ]
    uncovered = uncovered_doc_type_files(pairs, templates)
    types = sorted({doc_type for _, doc_type in uncovered})
    return {
        "uncovered_doc_types": types,
        "uncovered_n_files": len(uncovered),
        "uncovered_files": [f"{name} ({doc_type})" for name, doc_type in uncovered],
        "uncovered_pairs": [[name, doc_type] for name, doc_type in uncovered],
    }


def format_uncovered_doc_types_warning(payload: dict[str, Any], *, max_files: int = 20) -> str:
    """Human-readable warning for stdout and the control-center dialog."""
    types = [str(item) for item in (payload.get("uncovered_doc_types") or []) if str(item)]
    files = [str(item) for item in (payload.get("uncovered_files") or []) if str(item)]
    n_files = int(payload.get("uncovered_n_files") or len(files))
    type_text = ", ".join(types) if types else "—"
    shown = files[:max_files]
    lines = [
        "Нет шаблона штампа для типа(ов) документа: " + type_text + ".",
        f"Файлов без шаблона: {n_files}.",
        "Штамп для этих файлов не будет извлечён.",
    ]
    for item in shown:
        lines.append(f"  • {item}")
    remaining = n_files - len(shown)
    if remaining > 0:
        lines.append(f"  … и ещё {remaining}")
    return "\n".join(lines)


def _notify_stage_error(
    callback: Any,
    category: str,
    stage: str,
    error: Exception,
    *,
    file_name: str | None = None,
    detail: dict[str, Any] | None = None,
) -> None:
    method = _callback_method(callback, "on_stage_error")
    if method is not None:
        method(category, stage, error, file_name, detail or {})


def _notify_pipeline_finished(
    callback: Any,
    result_dir: str,
    total_elapsed: float,
    summary: dict[str, Any],
) -> None:
    method = _callback_method(callback, "on_pipeline_finished")
    if method is not None:
        method(result_dir, total_elapsed, summary)


def _ensure_project_root_on_path() -> str:
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if root not in sys.path:
        sys.path.insert(0, root)
    return root


def _log_v2_result(v2r: "V2PageResult", indent: str = "  ") -> None:
    """Print per-page template selection and field extraction summary to stdout."""
    score_str = f"{v2r.template_score:.2f}"
    print(f"{indent}Стр.{v2r.page_num}: «{v2r.template_name}» (score={score_str})")
    for field_id, fr in v2r.fields.items():
        val = fr.cleaned_value or ""
        has_val = bool(val.strip())
        if has_val:
            short_val = val[:80].replace("\n", " ")
            valid_mark = ""
            if fr.is_valid is False:
                valid_mark = " [regex fail]"
            print(f"{indent}  + {field_id} = {short_val}{valid_mark}")
        else:
            print(f"{indent}  - {field_id} = (пусто)")
    for w in v2r.warnings:
        print(f"{indent}  ! {w}")


def run_v2_extraction(
    curr_proj: list[V2Document],
    templates: list[StampTemplate],
    cfg: dict[str, Any],
    *,
    callback: ExtractionCallback | None = None,
    timing: Any | None = None,
) -> None:
    """Run per-file v2 stamp extraction into *curr_proj* (mutates documents)."""
    from pdf_parsing_v2.parallel import (
        _should_use_per_page_parallel,
        run_parallel_extraction_for_document,
    )

    total = len(curr_proj)
    verbose_log = cfg.get("debug_verbose_log", False)
    for idx, document in enumerate(curr_proj):
        file_path = document.file_full_path
        doc_type = document.doc_Type
        file_name = os.path.basename(file_path)
        t_file = time.perf_counter()
        if callback is not None:
            callback.on_file_start(file_path, idx, total)
        try:
            use_per_page, _per_page_cfg, _n_pages = _should_use_per_page_parallel(
                document, cfg
            )
            if use_per_page:
                file_result = run_parallel_extraction_for_document(
                    document,
                    templates,
                    cfg,
                    timing=None,
                )
                v2_results = file_result.pages
                timing_detail = dict(file_result.timing_detail)
                elapsed = file_result.elapsed_sec
            else:
                extracted = extract_all_pages_for_file(
                    file_path,
                    doc_type,
                    templates,
                    cfg,
                    with_timing=timing is not None,
                )
                if isinstance(extracted, tuple):
                    v2_results, timing_detail = extracted
                else:
                    v2_results, timing_detail = extracted, {}
                elapsed = time.perf_counter() - t_file
            for v2r in v2_results:
                page_data = V2PageData.from_v2_result(v2r, document)
                document.pages.append(page_data)
            document._v2_results = v2_results
            n_pages = len(v2_results)
            affine_timing = stamp_affine_uniformity_from_pages(v2_results)
            merged_detail = dict(timing_detail)
            merged_detail.update(affine_timing)
            affine_ui = stamp_affine_detail_for_ui(v2_results)
            if timing is not None:
                timing.record_file(
                    file_name,
                    elapsed,
                    n_pages=n_pages,
                    error="",
                    **merged_detail,
                )
            if callback is not None:
                callback.on_file_done(file_path, idx, n_pages, elapsed, affine_ui or None)
                batch_progress = _callback_method(callback, "on_batch_progress")
                if batch_progress is not None:
                    batch_progress(idx + 1, total)
            else:
                print(f"[{idx + 1}/{total}] {file_name}: {n_pages} стр. OK ({elapsed:.1f}s)")
            if verbose_log:
                for v2r in v2_results:
                    _log_v2_result(v2r)
        except Exception as e:
            elapsed = time.perf_counter() - t_file
            if timing is not None:
                timing.record_file(file_name, elapsed, n_pages=0, error=f"{type(e).__name__}: {e}")
            if callback is not None:
                callback.on_file_error(file_path, idx, e)
                batch_progress = _callback_method(callback, "on_batch_progress")
                if batch_progress is not None:
                    batch_progress(idx + 1, total)
            else:
                print(f"[{idx + 1}/{total}] {file_name}: ОШИБКА {e} ({elapsed:.1f}s)")
                traceback.print_exc()


def run_v2_tags(
    curr_proj: list[V2Document],
    *,
    tag_dict: dict | None = None,
    callback: PipelineCallback | None = None,
    timing: Any | None = None,
    max_workers: int | None = None,
    backend: str = "fitz",
) -> dict:
    """Collect tags via pdfminer; returns the dict used (global or *tag_dict*)."""
    t_tags = time.perf_counter()
    _notify_stage_start(callback, "tags", "parse_tags")
    try:
        from pdf_parsing_v2_tags import parse_tags as _parse_tags

        _parse_tags(
            curr_proj,
            tag_dict=tag_dict,
            callback=callback,
            timing=timing,
            max_workers=max_workers,
            backend=backend,
        )
    except Exception as e:
        print(f"[v2_pipeline] parse_tags: {e}")
        elapsed = time.perf_counter() - t_tags
        if timing is not None:
            timing.record_stage("tags", "parse_tags", elapsed, status="error", error=str(e))
        _notify_stage_error(callback, "tags", "parse_tags", e)
    else:
        elapsed = time.perf_counter() - t_tags
        if timing is not None:
            timing.record_stage(
                "tags",
                "parse_tags",
                elapsed,
                n_documents=len(curr_proj),
                max_workers=max_workers,
                backend=backend,
            )
        _notify_stage_done(
            callback,
            "tags",
            "parse_tags",
            elapsed,
            detail={
                "n_documents": len(curr_proj),
                "max_workers": max_workers,
                "backend": backend,
            },
        )
    from tags import tag_parser

    if tag_dict is None:
        return tag_parser.source_dict
    return tag_dict


def _run_extraction_stage(
    curr_proj: list[V2Document],
    templates: list[StampTemplate],
    cfg: dict[str, Any],
    *,
    use_parallel: bool,
    callback: ExtractionCallback | None = None,
    timing: Any | None = None,
) -> float:
    """Run extraction stage with standard timing and stage notifications."""
    _notify_stage_start(
        callback,
        "extraction",
        "run_extraction",
        detail={
            "use_parallel": use_parallel,
            "max_workers": cfg.get("max_workers"),
            "n_files": len(curr_proj),
        },
    )
    t_ext = time.perf_counter()
    if use_parallel:
        from pdf_parsing_v2.parallel import run_parallel_extraction

        run_parallel_extraction(
            curr_proj,
            templates,
            cfg,
            max_workers=cfg.get("max_workers"),
            callback=callback,
            timing=timing,
        )
    else:
        run_v2_extraction(
            curr_proj,
            templates,
            cfg,
            callback=callback,
            timing=timing,
        )
    elapsed_ext = time.perf_counter() - t_ext
    if timing is not None:
        timing.record(
            "extraction",
            elapsed_ext,
            used_parallel=use_parallel,
            max_workers=cfg.get("max_workers"),
        )
        timing.record_stage(
            "extraction",
            "run_extraction",
            elapsed_ext,
            used_parallel=use_parallel,
            max_workers=cfg.get("max_workers"),
        )
    _notify_stage_done(
        callback,
        "extraction",
        "run_extraction",
        elapsed_ext,
        detail={
            "use_parallel": use_parallel,
            "max_workers": cfg.get("max_workers"),
        },
    )
    return elapsed_ext


def run_v2_postprocess(
    curr_proj: list[V2Document],
    cfg: dict[str, Any],
    pdf_path: str,
    out_result_dir: str,
    *,
    tag_dict: dict | None = None,
    effective_project: str | None = None,
    templates: list[StampTemplate] | None = None,
    callback: PipelineCallback | None = None,
    timing: Any | None = None,
) -> dict[str, Any]:
    """OD → tag analyze → rules → optional debug Excel and debug PNGs."""
    artifacts: dict[str, Any] = {}
    normcontrol: dict[str, Any] = {}
    proj_od_list: list = []
    od_pdf_path = ""
    od_geom_out: list = []
    od_table_payload: dict[str, Any] = {}
    t_od = time.perf_counter()
    _notify_stage_start(callback, "od", "od_table_parsing")
    try:
        from pdf_parsing_v2_od.od_parsing import (
            build_od_table_payload,
            error_message_for_invalid_od_parse_result,
            find_od_pdf_path,
            od_table_parsing_f,
        )

        od_pdf_path = find_od_pdf_path(pdf_path) or ""
        if not od_pdf_path:
            proj_od_list = []
            od_table_payload = build_od_table_payload([], "", status="no_od_file")
        else:
            od_warnings: list[str] = []
            raw = od_table_parsing_f(
                pdf_full_path=od_pdf_path,
                debug_print_out_list=1,
                warnings_out=od_warnings,
                cfg=cfg,
                templates=templates or [],
                curr_proj=curr_proj,
                manifest_geometry_out=od_geom_out,
            )
            if not isinstance(raw, list):
                proj_od_list = []
                od_table_payload = build_od_table_payload(
                    [],
                    od_pdf_path,
                    status="error",
                    error=error_message_for_invalid_od_parse_result(od_pdf_path),
                    warnings=od_warnings,
                )
            elif len(raw) == 0:
                proj_od_list = []
                od_table_payload = build_od_table_payload(
                    [],
                    od_pdf_path,
                    status="no_rows",
                    warnings=od_warnings,
                )
            else:
                proj_od_list = raw
                od_table_payload = build_od_table_payload(
                    proj_od_list,
                    od_pdf_path,
                    status="ok",
                    warnings=od_warnings or None,
                )
    except Exception as e:
        print(f"[v2_pipeline] od_table_parsing: {e}")
        elapsed = time.perf_counter() - t_od
        od_table_payload = build_od_table_payload(
            [],
            od_pdf_path or "",
            status="error",
            error=str(e),
        )
        proj_od_list = []
        if timing is not None:
            timing.record_stage("od", "od_table_parsing", elapsed, status="error", error=str(e))
        _notify_stage_error(callback, "od", "od_table_parsing", e)
    else:
        elapsed = time.perf_counter() - t_od
        if timing is not None:
            timing.record_stage(
                "od",
                "od_table_parsing",
                elapsed,
                n_od_rows=len(proj_od_list),
            )
        _notify_stage_done(
            callback,
            "od",
            "od_table_parsing",
            elapsed,
            detail={"n_od_rows": len(proj_od_list)},
        )
    if od_pdf_path and templates:
        try:
            from pdf_parsing_v2_od.od_engine_findtables_debug import (
                collect_od_manifest_findtables_debug,
            )

            dbg = collect_od_manifest_findtables_debug(
                od_pdf_path,
                templates,
                cfg,
                curr_proj=curr_proj,
                geometry_rows=od_geom_out if od_geom_out else None,
            )
            if dbg:
                od_table_payload["engine_findtables_debug"] = dbg
        except Exception as exc:
            od_table_payload["engine_findtables_debug"] = {"error": str(exc), "pages": []}
    artifacts["od_table"] = od_table_payload

    t_tag_analyze = time.perf_counter()
    _notify_stage_start(callback, "tags", "tag_analyze")
    tag_analysis_payload: dict[str, Any] = {}
    try:
        from tags import tag_parser

        tag_analysis_payload = tag_parser.analyze(out_result_dir, source_dict_override=tag_dict)
        if not isinstance(tag_analysis_payload, dict):
            tag_analysis_payload = {
                "status": "error",
                "error": "tag_parser.analyze returned non-dict",
                "checks": [],
                "invalid_tags": [],
                "warnings": [],
                "sheet_to_pdf": {},
                "mto_wbs": "",
                "mto_sheet_name": "",
                "text_report_path": "",
                "high_frequency_threshold": 3,
            }
    except Exception as e:
        print(f"[v2_pipeline] tag_parser.analyze: {e}")
        tag_analysis_payload = {
            "status": "error",
            "error": str(e),
            "checks": [],
            "invalid_tags": [],
            "warnings": [],
            "sheet_to_pdf": {},
            "mto_wbs": "",
            "mto_sheet_name": "",
            "text_report_path": "",
            "high_frequency_threshold": 3,
        }
        elapsed = time.perf_counter() - t_tag_analyze
        if timing is not None:
            timing.record_stage("tags", "tag_analyze", elapsed, status="error", error=str(e))
        _notify_stage_error(callback, "tags", "tag_analyze", e)
    else:
        elapsed = time.perf_counter() - t_tag_analyze
        if timing is not None:
            timing.record_stage("tags", "tag_analyze", elapsed)
        _notify_stage_done(callback, "tags", "tag_analyze", elapsed)
    artifacts["tag_analysis"] = tag_analysis_payload

    t_rules = time.perf_counter()
    _notify_stage_start(callback, "rules", "rules_check")
    try:
        from pdf_parsing_v2_rules import rules_check_start_v2

        _od_tbl = artifacts.get("od_table") or {}
        normcontrol = rules_check_start_v2(
            curr_proj,
            proj_od_list,
            pdf_path,
            out_result_dir,
            effective_project=effective_project,
            sanitize_illegal_chars=excel_sanitize_illegal_chars_enabled(cfg),
            od_warnings=list(_od_tbl.get("warnings") or []),
            od_pdf_path=str(_od_tbl.get("od_pdf_path") or ""),
        )
    except Exception as e:
        print(f"[v2_pipeline] rules_check: {e}")
        traceback.print_exc()
        elapsed = time.perf_counter() - t_rules
        if timing is not None:
            timing.record_stage("rules", "rules_check", elapsed, status="error", error=str(e))
        _notify_stage_error(callback, "rules", "rules_check", e)
    else:
        elapsed = time.perf_counter() - t_rules
        if normcontrol.get("xlsx_path"):
            artifacts["normcontrol_path"] = normcontrol.get("xlsx_path", "")
        if timing is not None:
            timing.record_stage(
                "rules",
                "rules_check",
                elapsed,
                n_documents=len(curr_proj),
            )
        _notify_stage_done(
            callback,
            "rules",
            "rules_check",
            elapsed,
                detail={
                    "n_documents": len(curr_proj),
                    "n_errors": (normcontrol.get("stats", {}) or {}).get("total_errors", 0),
                },
        )

    if cfg.get("export_debug_excel", True):
        t_report = time.perf_counter()
        _notify_stage_start(callback, "report", "debug_report")
        try:
            if not effective_project or not str(effective_project).strip():
                raise ValueError(
                    "укажите project в pdf_v2_config.json или параметр project "
                    "у run_v2_pipeline — нужен каталог полей проекта.",
                )
            templates_dir = cfg.get("templates_dir", "pdf_parsing_v2_engine/templates")
            if not os.path.isabs(templates_dir):
                templates_dir = os.path.join(
                    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    templates_dir,
                )
            report_catalog = load_catalog_for_project(templates_dir, effective_project)
            if report_catalog is None:
                raise ValueError(
                    f"не найден catalog.json для проекта "
                    f"{effective_project!r} в {templates_dir}.",
                )
            from pdf_parsing_v2_engine.v2_report import collect_results_from_curr_proj, save_v2_debug_report

            all_results = collect_results_from_curr_proj(curr_proj)
            if all_results:
                artifacts["debug_report_path"] = save_v2_debug_report(
                    all_results,
                    out_result_dir,
                    report_catalog,
                )
        except Exception as e:
            print(f"[v2_pipeline] debug report пропущен: {e}")
            traceback.print_exc()
            elapsed = time.perf_counter() - t_report
            if timing is not None:
                timing.record_stage("report", "debug_report", elapsed, status="error", error=str(e))
            _notify_stage_error(callback, "report", "debug_report", e)
        else:
            elapsed = time.perf_counter() - t_report
            if timing is not None:
                timing.record_stage(
                    "report",
                    "debug_report",
                    elapsed,
                    output_path=artifacts.get("debug_report_path", ""),
                )
            _notify_stage_done(
                callback,
                "report",
                "debug_report",
                elapsed,
                detail={"output_path": artifacts.get("debug_report_path", "")},
            )

    if cfg.get("debug_visual", False) and templates:
        t_visual = time.perf_counter()
        _notify_stage_start(callback, "report", "debug_visual")
        try:
            from pdf_parsing_v2_engine.debug_visual import render_debug_for_file

            vis_dir = cfg.get("debug_visual_dir", "debug/v2_visual")
            if not os.path.isabs(vis_dir):
                vis_dir = os.path.join(out_result_dir, vis_dir)
            overlay = cfg.get("debug_visual_overlay", False)
            total_png = 0
            for document in curr_proj:
                v2r_list = getattr(document, "_v2_results", None)
                if not v2r_list:
                    continue
                pngs = render_debug_for_file(
                    document.file_full_path,
                    v2r_list,
                    templates,
                    vis_dir,
                    overlay=overlay,
                )
                total_png += len(pngs)
            if total_png:
                print(f"[v2_pipeline] debug_visual: {total_png} PNG в {vis_dir}")
            artifacts["debug_visual_dir"] = vis_dir
            artifacts["debug_visual_png_count"] = total_png
        except Exception as e:
            print(f"[v2_pipeline] debug_visual: {e}")
            traceback.print_exc()
            elapsed = time.perf_counter() - t_visual
            if timing is not None:
                timing.record_stage("report", "debug_visual", elapsed, status="error", error=str(e))
            _notify_stage_error(callback, "report", "debug_visual", e)
        else:
            elapsed = time.perf_counter() - t_visual
            if timing is not None:
                timing.record_stage(
                    "report",
                    "debug_visual",
                    elapsed,
                    output_dir=artifacts.get("debug_visual_dir", ""),
                    n_png=artifacts.get("debug_visual_png_count", 0),
                )
            _notify_stage_done(
                callback,
                "report",
                "debug_visual",
                elapsed,
                detail={
                    "output_dir": artifacts.get("debug_visual_dir", ""),
                    "n_png": artifacts.get("debug_visual_png_count", 0),
                },
            )
    artifacts["_normcontrol_payload"] = normcontrol
    return artifacts

def run_v2_pipeline(
    pdf_path: str,
    cfg: dict[str, Any],
    project: str | None = None,
    *,
    extraction_callback: ExtractionCallback | None = None,
) -> str:
    """Process *pdf_path* folder through the v2 template engine.

    Args:
        pdf_path: папка с PDF файлами.
        cfg: конфигурация (pdf_v2_config.json).
        project: фильтр по имени проекта (catalog.projects[] или имя папки).
            Если None — используются шаблоны всех проектов.
            Может также быть передан через cfg["project"].
        extraction_callback: optional progress callback for per-file extraction
            (used by GUI monitors). Compatible with both ``ExtractionCallback``
            and ``ParallelCallback`` protocols.

    Returns the result directory path (same convention as v1 pipeline).
    """
    _ensure_project_root_on_path()

    from utils.path import get_files_single, get_path_out_dir

    from pdf_parsing_v2.v2_timing import TimingCollector

    timing = TimingCollector()

    out_result_dir = get_path_out_dir(pdf_path)
    t_total_start = time.perf_counter()
    callback = extraction_callback

    _notify_phase(callback, "Find files")
    _notify_stage_start(callback, "pipeline", "find_files", detail={"pdf_path": pdf_path})
    t0 = time.perf_counter()
    curr_proj = get_files_single(pdf_path, factory=V2Document.from_file_path)
    if curr_proj == -1 or not curr_proj:
        print("[v2_pipeline] PDF файлы не найдены в", pdf_path)
        elapsed_total = time.perf_counter() - t_total_start
        timing.record("find_files", time.perf_counter() - t0)
        timing.record("total_pipeline", elapsed_total)
        summary = {
            "result_dir": out_result_dir,
            "effective_project": project or cfg.get("project") or None,
            "parallel_enabled": bool(cfg.get("parallel", False)),
            "used_parallel": False,
            "max_workers": cfg.get("max_workers"),
            "n_files": 0,
            "n_pages": 0,
            "artifacts": {},
            "od_table": {},
            "timing": timing.summary(),
        }
        _notify_stage_done(callback, "pipeline", "find_files", time.perf_counter() - t0, detail={"n_files": 0})
        _notify_pipeline_finished(callback, out_result_dir, elapsed_total, summary)
        print(out_result_dir)
        return out_result_dir
    elapsed_find = time.perf_counter() - t0
    timing.record("find_files", elapsed_find, n_files=len(curr_proj))
    timing.record_stage("pipeline", "find_files", elapsed_find, n_files=len(curr_proj))
    _notify_stage_done(callback, "pipeline", "find_files", elapsed_find, detail={"n_files": len(curr_proj)})
    print(f"[v2_pipeline] Найдено {len(curr_proj)} PDF ({elapsed_find:.1f}s)")

    templates_dir = resolve_templates_dir(cfg)

    effective_project: str | None = project or cfg.get("project") or None
    _notify_phase(callback, "Load templates")
    _notify_stage_start(
        callback,
        "pipeline",
        "load_templates",
        detail={"project": effective_project or "", "templates_dir": templates_dir},
    )
    t_tpl = time.perf_counter()
    if effective_project:
        templates = load_project_templates(templates_dir, effective_project)
        print(f"[v2_pipeline] Проект: {effective_project!r}, шаблонов: {len(templates)}")
    else:
        templates = load_all_templates(templates_dir)
        print(f"[v2_pipeline] Загружено {len(templates)} шаблонов (все проекты)")
    elapsed_tpl = time.perf_counter() - t_tpl
    timing.record("load_templates", elapsed_tpl, n_templates=len(templates))
    timing.record_stage(
        "pipeline",
        "load_templates",
        elapsed_tpl,
        n_templates=len(templates),
        project=effective_project or "",
    )
    uncovered_payload = _uncovered_doc_type_payload(curr_proj, templates)
    load_templates_detail: dict[str, Any] = {
        "n_templates": len(templates),
        "project": effective_project or "",
    }
    if uncovered_payload["uncovered_n_files"]:
        load_templates_detail.update(
            {
                "uncovered_doc_types": uncovered_payload["uncovered_doc_types"],
                "uncovered_n_files": uncovered_payload["uncovered_n_files"],
                "uncovered_files": uncovered_payload["uncovered_files"],
                "uncovered_pairs": uncovered_payload["uncovered_pairs"],
            }
        )
        warning_text = format_uncovered_doc_types_warning(uncovered_payload)
        print(f"[v2_pipeline] ПРЕДУПРЕЖДЕНИЕ:\n{warning_text}")
    _notify_stage_done(
        callback,
        "pipeline",
        "load_templates",
        elapsed_tpl,
        detail=load_templates_detail,
    )

    if not templates:
        print(f"[v2_pipeline] Нет шаблонов в {templates_dir}")
        elapsed_total = time.perf_counter() - t_total_start
        timing.record("total_pipeline", elapsed_total)
        summary = {
            "result_dir": out_result_dir,
            "effective_project": effective_project,
            "parallel_enabled": bool(cfg.get("parallel", False)),
            "used_parallel": False,
            "max_workers": cfg.get("max_workers"),
            "n_files": len(curr_proj),
            "n_pages": 0,
            "artifacts": {},
            "od_table": {},
            "timing": timing.summary(),
            "uncovered_doc_types": list(uncovered_payload.get("uncovered_doc_types") or []),
            "uncovered_n_files": int(uncovered_payload.get("uncovered_n_files") or 0),
        }
        _notify_pipeline_finished(callback, out_result_dir, elapsed_total, summary)
        print(out_result_dir)
        return out_result_dir

    use_parallel = cfg.get("parallel", False) and len(curr_proj) > 1
    tag_dict: dict = {}
    progress_callback: PipelineCallback | None = None
    tags_max_workers = cfg.get("tags_max_workers")
    tags_backend = str(cfg.get("tags_text_backend", "fitz") or "fitz")
    if callback is not None:
        progress_callback = _HeavyProgressProxy(callback, len(curr_proj))

    _notify_phase(progress_callback, "Extraction+Tags")
    t_tags = time.perf_counter()
    with ThreadPoolExecutor(max_workers=2) as executor:
        extraction_future = executor.submit(
            _run_extraction_stage,
            curr_proj,
            templates,
            cfg,
            use_parallel=use_parallel,
            callback=progress_callback,
            timing=timing,
        )
        tags_future = executor.submit(
            run_v2_tags,
            curr_proj,
            tag_dict=tag_dict,
            callback=progress_callback,
            timing=timing,
            max_workers=tags_max_workers,
            backend=tags_backend,
        )
        tags_future.result()
        timing.record("tags", time.perf_counter() - t_tags)
        elapsed_ext = extraction_future.result()

    _notify_phase(callback, "Postprocess")
    t_post = time.perf_counter()
    artifacts = run_v2_postprocess(
        curr_proj,
        cfg,
        pdf_path,
        out_result_dir,
        tag_dict=tag_dict,
        effective_project=effective_project,
        templates=templates,
        callback=callback,
        timing=timing,
    )
    timing.record("postprocess", time.perf_counter() - t_post)

    elapsed_total = time.perf_counter() - t_total_start
    timing.record("total_pipeline", elapsed_total)
    timing_path = os.path.join(out_result_dir, "v2_timing_log.xlsx")
    if cfg.get("timing_log", False):
        _notify_phase(callback, "Timing export")
        _notify_stage_start(callback, "timing", "timing_export", detail={"output_path": timing_path})
        t_export = time.perf_counter()
        try:
            os.makedirs(out_result_dir, exist_ok=True)
            timing.to_excel(
                timing_path,
                sanitize_illegal_chars=excel_sanitize_illegal_chars_enabled(cfg),
            )
            print(f"[v2_pipeline] Timing log: {timing_path}")
            artifacts["timing_path"] = timing_path
            elapsed_export = time.perf_counter() - t_export
            timing.record_stage(
                "timing",
                "timing_export",
                elapsed_export,
                output_path=timing_path,
            )
            _notify_stage_done(
                callback,
                "timing",
                "timing_export",
                elapsed_export,
                detail={"output_path": timing_path},
            )
        except Exception as e:
            print(f"[v2_pipeline] Timing log error: {e}")
            elapsed_export = time.perf_counter() - t_export
            timing.record_stage(
                "timing",
                "timing_export",
                elapsed_export,
                status="error",
                error=str(e),
                output_path=timing_path,
            )
            _notify_stage_error(callback, "timing", "timing_export", e, detail={"output_path": timing_path})

    m, s = divmod(int(elapsed_total), 60)
    total_pages = sum(len(getattr(doc, "_v2_results", []) or []) for doc in curr_proj)
    normcontrol = artifacts.pop("_normcontrol_payload", {})
    od_table = artifacts.pop("od_table", {})
    tag_analysis = artifacts.pop("tag_analysis", {})
    summary = {
        "result_dir": out_result_dir,
        "effective_project": effective_project,
        "parallel_enabled": bool(cfg.get("parallel", False)),
        "used_parallel": use_parallel,
        "max_workers": cfg.get("max_workers"),
        "tags_max_workers": tags_max_workers,
        "tags_text_backend": tags_backend,
        "n_files": len(curr_proj),
        "n_pages": total_pages,
        "artifacts": artifacts,
        "normcontrol": normcontrol,
        "od_table": od_table,
        "tag_analysis": tag_analysis,
        "timing": timing.summary(),
        "uncovered_doc_types": list(uncovered_payload.get("uncovered_doc_types") or []),
        "uncovered_n_files": int(uncovered_payload.get("uncovered_n_files") or 0),
    }
    _notify_pipeline_finished(callback, out_result_dir, elapsed_total, summary)
    print(f"[v2_pipeline] Готово за {m}м {s:02d}с")
    print(out_result_dir)
    return out_result_dir


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------


def _maybe_run_mto_regression(
    cfg: dict[str, Any],
    *,
    pdf_dir_override: str | None,
    template_override: str | None,
    strict: bool,
    max_nl_rot90: int,
) -> int:
    """Run MTO page-1 probe; return 0 OK, 1 strict fail, 2 setup error."""
    from pdf_parsing_v2_engine.tools.mto_page1_batch_probe import (
        check_mto_regression_strict,
        default_mto_regression_pdf_dir,
        default_mto_regression_template_path,
        run_mto_page1_batch_probe,
    )

    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    pdf_dir = (pdf_dir_override or "").strip() or (cfg.get("mto_regression_pdf_dir") or "").strip()
    if not pdf_dir:
        pdf_dir = default_mto_regression_pdf_dir()
    elif not os.path.isabs(pdf_dir):
        pdf_dir = os.path.normpath(os.path.join(root, pdf_dir))

    template_path = (template_override or "").strip() or (cfg.get("mto_regression_template") or "").strip()
    if not template_path:
        template_path = default_mto_regression_template_path()
    elif not os.path.isabs(template_path):
        template_path = os.path.normpath(os.path.join(root, template_path))

    report = run_mto_page1_batch_probe(
        pdf_dir,
        template_path,
        v2_cfg=cfg,
        print_report=True,
        json_out=None,
    )
    if report.get("error"):
        print(f"[v2_pipeline] mto_regression: {report['error']}", file=sys.stderr)
        return 2

    if strict:
        reasons = check_mto_regression_strict(
            report,
            max_newlines_rot90=int(max_nl_rot90),
        )
        if reasons:
            print("[v2_pipeline] mto_regression STRICT FAIL:", file=sys.stderr)
            for line in reasons:
                print(f"  {line}", file=sys.stderr)
            return 1
    return 0


def main() -> int:
    import argparse
    import json

    parser = argparse.ArgumentParser(description="v2 PDF pipeline")
    parser.add_argument("pdf_path", help="Папка с PDF файлами")
    parser.add_argument("config", nargs="?", default=None, help="Путь к JSON конфигу")
    parser.add_argument(
        "--project",
        default=None,
        help="Фильтр по имени проекта (папка или catalog.projects[])",
    )
    parser.add_argument(
        "--mto-regression",
        action="store_true",
        help="После прогона: smoke MTO стр.1 (char_center vs get_textbox), см. mto_page1_batch_probe",
    )
    parser.add_argument(
        "--mto-regression-dir",
        default=None,
        help="Каталог PDF для MTO-регрессии (по умолчанию templates/test_pdf/MTO или конфиг)",
    )
    parser.add_argument(
        "--mto-regression-template",
        default=None,
        help="Путь к mto_page1.json (по умолчанию agcc_287/mto_page1.json или конфиг)",
    )
    parser.add_argument(
        "--mto-regression-strict",
        action="store_true",
        help="Завершить с кодом 1 при провале порогов (90/270: лишние \\n в char_center)",
    )
    parser.add_argument(
        "--mto-regression-max-nl-rot90",
        type=int,
        default=None,
        help="Порог суммарных \\n в raw (char_center) для страниц 90/270 при --strict",
    )
    args = parser.parse_args()

    _ensure_project_root_on_path()

    pdf_path = os.path.abspath(args.pdf_path)
    if not os.path.isdir(pdf_path):
        print(f"Не директория: {pdf_path}", file=sys.stderr)
        return 1

    if args.config and os.path.isfile(args.config):
        with open(args.config, encoding="utf-8") as f:
            cfg = json.load(f)
    else:
        cfg = load_v2_config()

    project: str | None = args.project or cfg.get("project") or None

    print(f"[v2_pipeline] pdf_path={pdf_path}")
    print(f"[v2_pipeline] templates_dir={cfg.get('templates_dir', '(default)')}")
    if project:
        print(f"[v2_pipeline] project={project!r}")

    try:
        result_dir = run_v2_pipeline(pdf_path, cfg, project=project)
    except Exception:
        traceback.print_exc()
        return 1

    if not result_dir:
        return 0

    want_mto = bool(args.mto_regression or cfg.get("mto_regression_after_run"))
    if want_mto:
        strict = bool(args.mto_regression_strict or cfg.get("mto_regression_strict"))
        max_nl = args.mto_regression_max_nl_rot90
        if max_nl is None:
            max_nl = int(cfg.get("mto_regression_max_nl_rot90", 280))
        mto_code = _maybe_run_mto_regression(
            cfg,
            pdf_dir_override=args.mto_regression_dir,
            template_override=args.mto_regression_template,
            strict=strict,
            max_nl_rot90=max_nl,
        )
        if mto_code != 0:
            return mto_code

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
