"""Parallel PDF extraction via ProcessPoolExecutor.

Unit of parallelism: one PDF file (not one page).
Worker function: ``extract_all_pages_for_file`` from engine.

Usage::

    from pdf_parsing_v2.parallel import run_parallel_extraction
    run_parallel_extraction(curr_proj, templates, cfg, max_workers=4)
"""

from __future__ import annotations

import os
import pickle
import math
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import dataclass
from typing import Any, Protocol

import fitz

from pdf_parsing_v2_engine.models import StampTemplate, V2PageResult
from pdf_parsing_v2_engine.document import V2Document, V2PageData

from pdf_parsing_v2.stamp_affine_summary import (
    stamp_affine_detail_for_ui,
    stamp_affine_uniformity_from_pages,
)


# ---------------------------------------------------------------------------
# Result DTOs
# ---------------------------------------------------------------------------


@dataclass
class V2FileExtractResult:
    """Extraction result for one PDF file (pickle-safe, returned by worker)."""

    file_path: str
    doc_type: str
    index: int
    pages: list[V2PageResult]
    elapsed_sec: float
    timing_detail: dict[str, Any]
    error: str | None = None

    @property
    def ok(self) -> bool:
        return self.error is None


@dataclass
class V2ParallelResult:
    """Aggregate result from parallel extraction."""

    file_results: list[V2FileExtractResult]
    total_elapsed_sec: float
    n_workers: int

    @property
    def n_files(self) -> int:
        return len(self.file_results)

    @property
    def n_pages_total(self) -> int:
        return sum(len(r.pages) for r in self.file_results)

    @property
    def n_errors(self) -> int:
        return sum(1 for r in self.file_results if not r.ok)


@dataclass(frozen=True)
class _ParallelTaskSpec:
    """Internal task descriptor for mixed file/page scheduling."""

    kind: str
    file_path: str
    doc_type: str
    index: int
    page_indices: list[int] | None = None


# ---------------------------------------------------------------------------
# Callback protocol
# ---------------------------------------------------------------------------


class ParallelCallback(Protocol):
    """Progress reporting for parallel extraction.

    All methods called from the **main process** (not from workers).
    """

    def on_file_start(
        self, file_path: str, index: int, total: int
    ) -> None: ...

    def on_file_done(
        self,
        file_path: str,
        index: int,
        n_pages: int,
        elapsed_sec: float,
        file_detail: dict[str, Any] | None = None,
    ) -> None: ...

    def on_file_error(
        self, file_path: str, index: int, error: Exception
    ) -> None: ...

    def on_batch_progress(self, completed: int, total: int) -> None: ...


# ---------------------------------------------------------------------------
# Worker internals
# ---------------------------------------------------------------------------

_WORKER_TEMPLATES: list[StampTemplate] | None = None
_WORKER_CFG: dict | None = None


def _timing_add(bucket: dict[str, float] | None, key: str, elapsed_sec: float) -> None:
    """Accumulate timing in a mutable dict when instrumentation is enabled."""
    if bucket is None:
        return
    bucket[key] = round(float(bucket.get(key, 0.0)) + float(elapsed_sec), 6)


def _worker_init(templates_bytes: bytes, cfg_bytes: bytes) -> None:
    """Unpickle shared data once per worker process."""
    global _WORKER_TEMPLATES, _WORKER_CFG
    _WORKER_TEMPLATES = pickle.loads(templates_bytes)
    _WORKER_CFG = pickle.loads(cfg_bytes)


def _worker_extract(
    file_path: str, doc_type: str, index: int
) -> V2FileExtractResult:
    """Process one PDF in a worker process (never raises)."""
    from pdf_parsing_v2_engine.stamp_extractor import extract_all_pages_for_file

    t0 = time.perf_counter()
    try:
        pages = extract_all_pages_for_file(
            file_path,
            doc_type,
            _WORKER_TEMPLATES,
            _WORKER_CFG,
            with_timing=True,
        )
        if isinstance(pages, tuple):
            page_results, timing_detail = pages
        else:
            page_results, timing_detail = pages, {}
        return V2FileExtractResult(
            file_path=file_path,
            doc_type=doc_type,
            index=index,
            pages=page_results,
            elapsed_sec=time.perf_counter() - t0,
            timing_detail=timing_detail,
        )
    except Exception as e:
        return V2FileExtractResult(
            file_path=file_path,
            doc_type=doc_type,
            index=index,
            pages=[],
            elapsed_sec=time.perf_counter() - t0,
            timing_detail={},
            error=f"{type(e).__name__}: {e}",
        )


def _worker_extract_pages(
    file_path: str,
    doc_type: str,
    index: int,
    page_indices: list[int],
) -> V2FileExtractResult:
    """Process selected pages from one PDF in a worker process."""
    from pdf_parsing_v2_engine.stamp_extractor import extract_page

    t0 = time.perf_counter()
    timing_detail: dict[str, float] = {}
    try:
        fitz_doc = fitz.open(file_path)
        _timing_add(timing_detail, "open_pdf", time.perf_counter() - t0)
        page_results: list[V2PageResult] = []
        try:
            page_cfg = dict(_WORKER_CFG or {})
            page_cfg["source_pdf_basename"] = os.path.basename(file_path)
            t_pages = time.perf_counter()
            for page_index in page_indices:
                fitz_page = fitz_doc[page_index]
                page_timing: dict[str, float] = {}
                v2_result = extract_page(
                    fitz_page,
                    doc_type,
                    page_index + 1,
                    _WORKER_TEMPLATES or [],
                    cfg=page_cfg,
                    timing=page_timing,
                )
                page_results.append(v2_result)
                for key, value in page_timing.items():
                    _timing_add(timing_detail, key, value)
            _timing_add(timing_detail, "page_loop_total", time.perf_counter() - t_pages)
        finally:
            fitz_doc.close()
        page_results.sort(key=lambda item: item.page_num)
        return V2FileExtractResult(
            file_path=file_path,
            doc_type=doc_type,
            index=index,
            pages=page_results,
            elapsed_sec=time.perf_counter() - t0,
            timing_detail=timing_detail,
        )
    except Exception as e:
        return V2FileExtractResult(
            file_path=file_path,
            doc_type=doc_type,
            index=index,
            pages=[],
            elapsed_sec=time.perf_counter() - t0,
            timing_detail=timing_detail,
            error=f"{type(e).__name__}: {e}",
        )


def _default_max_workers() -> int:
    cpus = os.cpu_count() or 4
    return max(1, min(cpus - 1, 8))


def _default_page_workers(n_pages: int) -> int:
    cpus = os.cpu_count() or 4
    return max(1, min(cpus - 1, n_pages, 8))


def _get_per_page_cfg(cfg: dict[str, Any] | None) -> dict[str, Any]:
    """Return normalized nested config for per-page parallel extraction."""
    src = dict((cfg or {}).get("per_page") or {})
    doc_types_default = {
        "DW": True,
        "WIR": False,
        "LAY": False,
        "CAE": False,
        "GA": False,
        "PL": False,
        "NI": False,
        "MTO": False,
        "BOE": False,
        "BOM": False,
        "BOQ": False,
        "OD": False,
        "CJ": False,
        "VO": False,
    }
    raw_doc_types = src.get("doc_types")
    doc_types = dict(doc_types_default)
    if isinstance(raw_doc_types, dict):
        for key in doc_types:
            if key in raw_doc_types:
                doc_types[key] = bool(raw_doc_types[key])
    return {
        "enabled": bool(src.get("enabled", True)),
        "min_pages": max(1, int(src.get("min_pages", 4) or 1)),
        "max_workers": max(0, int(src.get("max_workers", 0) or 0)),
        "chunk_size_pages": max(0, int(src.get("chunk_size_pages", 0) or 0)),
        "doc_types": doc_types,
    }


def _resolve_chunk_size(n_pages: int, workers: int, cfg: dict[str, Any]) -> int:
    """Return pages per worker chunk."""
    chunk_size = int(cfg.get("chunk_size_pages", 0) or 0)
    if chunk_size > 0:
        return max(1, chunk_size)
    return max(1, math.ceil(n_pages / max(1, workers)))


def _split_page_indices(n_pages: int, chunk_size: int) -> list[list[int]]:
    """Split page indexes into contiguous chunks."""
    return [
        list(range(start, min(start + chunk_size, n_pages)))
        for start in range(0, n_pages, chunk_size)
    ]


def _should_use_per_page_parallel(
    document: V2Document,
    cfg: dict[str, Any],
) -> tuple[bool, dict[str, Any], int]:
    """Decide whether selective per-page PPE should be used for one document."""
    per_page_cfg = _get_per_page_cfg(cfg)
    if not per_page_cfg["enabled"]:
        return False, per_page_cfg, 0
    if not per_page_cfg["doc_types"].get(str(document.doc_Type or "").upper(), False):
        return False, per_page_cfg, 0
    try:
        with fitz.open(document.file_full_path) as fitz_doc:
            n_pages = len(fitz_doc)
    except Exception:
        return False, per_page_cfg, 0
    if n_pages < int(per_page_cfg["min_pages"]):
        return False, per_page_cfg, n_pages
    return True, per_page_cfg, n_pages


def run_parallel_extraction_for_document(
    document: V2Document,
    templates: list[StampTemplate],
    cfg: dict[str, Any],
    *,
    max_workers: int | None = None,
    timing: Any | None = None,
) -> V2FileExtractResult:
    """Run selective page-range PPE for one heavy document."""
    use_per_page, per_page_cfg, n_pages = _should_use_per_page_parallel(document, cfg)
    if not use_per_page:
        raise ValueError("Document is not eligible for per-page parallel extraction")

    workers = max_workers
    if workers in (None, 0):
        workers = int(per_page_cfg.get("max_workers", 0) or 0)
    if workers <= 0:
        workers = _default_page_workers(n_pages)
    workers = max(1, min(int(workers), n_pages))

    chunk_size = _resolve_chunk_size(n_pages, workers, per_page_cfg)
    tasks = _split_page_indices(n_pages, chunk_size)

    tpl_bytes = pickle.dumps(templates)
    cfg_bytes = pickle.dumps(cfg)
    results: list[V2FileExtractResult] = []
    t_total = time.perf_counter()

    with ProcessPoolExecutor(
        max_workers=min(workers, len(tasks)),
        initializer=_worker_init,
        initargs=(tpl_bytes, cfg_bytes),
    ) as executor:
        futures = [
            executor.submit(
                _worker_extract_pages,
                document.file_full_path,
                document.doc_Type,
                0,
                page_indices,
            )
            for page_indices in tasks
        ]
        for future in as_completed(futures):
            results.append(future.result())

    all_pages: list[V2PageResult] = []
    merged_timing: dict[str, Any] = {
        "per_page_parallel_used": True,
        "per_page_parallel_workers": workers,
        "per_page_parallel_chunk_size": chunk_size,
        "per_page_parallel_chunks": len(tasks),
    }
    error_messages: list[str] = []
    for result in results:
        if result.ok:
            all_pages.extend(result.pages)
            for key, value in result.timing_detail.items():
                if isinstance(value, (int, float)):
                    _timing_add(merged_timing, key, float(value))
                else:
                    merged_timing[key] = value
        else:
            error_messages.append(result.error or "unknown")
    all_pages.sort(key=lambda item: item.page_num)
    elapsed = time.perf_counter() - t_total
    merged_timing["per_page_parallel_total"] = round(elapsed, 6)
    merged_timing.update(stamp_affine_uniformity_from_pages(all_pages))
    error = "; ".join(error_messages) if error_messages else None
    file_result = V2FileExtractResult(
        file_path=document.file_full_path,
        doc_type=document.doc_Type,
        index=0,
        pages=all_pages,
        elapsed_sec=elapsed,
        timing_detail=merged_timing,
        error=error,
    )
    if timing is not None:
        timing.record_file(
            os.path.basename(document.file_full_path),
            elapsed,
            n_pages=len(all_pages),
            error=error,
            **merged_timing,
        )
    return file_result


def _resolve_per_page_workers(
    n_pages: int,
    per_page_cfg: dict[str, Any],
    max_workers: int | None = None,
) -> int:
    """Resolve worker count for page-range PPE."""
    workers = max_workers
    if workers in (None, 0):
        workers = int(per_page_cfg.get("max_workers", 0) or 0)
    if workers <= 0:
        workers = _default_page_workers(n_pages)
    return max(1, min(int(workers), n_pages))


def _build_mixed_task_specs(
    curr_proj: list[V2Document],
    cfg: dict[str, Any],
    *,
    max_workers: int | None = None,
) -> tuple[list[_ParallelTaskSpec], dict[int, dict[str, Any]]]:
    """Build mixed file/page task list and per-document metadata."""
    specs: list[_ParallelTaskSpec] = []
    doc_meta: dict[int, dict[str, Any]] = {}
    for idx, document in enumerate(curr_proj):
        use_per_page, per_page_cfg, n_pages = _should_use_per_page_parallel(document, cfg)
        if use_per_page:
            workers = _resolve_per_page_workers(n_pages, per_page_cfg, max_workers=max_workers)
            chunk_size = _resolve_chunk_size(n_pages, workers, per_page_cfg)
            chunks = _split_page_indices(n_pages, chunk_size)
            doc_meta[idx] = {
                "mode": "per_page",
                "n_pages": n_pages,
                "workers": workers,
                "chunk_size": chunk_size,
                "chunks": len(chunks),
            }
            for page_indices in chunks:
                specs.append(
                    _ParallelTaskSpec(
                        kind="pages",
                        file_path=document.file_full_path,
                        doc_type=document.doc_Type,
                        index=idx,
                        page_indices=page_indices,
                    )
                )
        else:
            doc_meta[idx] = {"mode": "file"}
            specs.append(
                _ParallelTaskSpec(
                    kind="file",
                    file_path=document.file_full_path,
                    doc_type=document.doc_Type,
                    index=idx,
                )
            )
    return specs, doc_meta


def _run_parallel_mixed_extraction(
    curr_proj: list[V2Document],
    templates: list[StampTemplate],
    cfg: dict[str, Any],
    *,
    max_workers: int | None = None,
    callback: ParallelCallback | None = None,
    timing: Any | None = None,
) -> V2ParallelResult:
    """Run file-level and selective page-level extraction in one executor."""
    specs, doc_meta = _build_mixed_task_specs(curr_proj, cfg, max_workers=max_workers)
    workers = max_workers or _default_max_workers()
    workers = min(workers, len(specs)) if specs else 1

    tpl_bytes = pickle.dumps(templates)
    cfg_bytes = pickle.dumps(cfg)
    t_total = time.perf_counter()
    total_docs = len(curr_proj)

    if callback is not None:
        for idx, document in enumerate(curr_proj):
            callback.on_file_start(document.file_full_path, idx, total_docs)

    aggregated: dict[int, dict[str, Any]] = {
        idx: {
            "file_path": doc.file_full_path,
            "doc_type": doc.doc_Type,
            "index": idx,
            "pages": [],
            "timing_detail": {},
            "error_messages": [],
            "pending": 0,
            "start_at": time.perf_counter(),
            "single_elapsed_sec": None,
        }
        for idx, doc in enumerate(curr_proj)
    }
    for spec in specs:
        aggregated[spec.index]["pending"] += 1

    completed_docs = 0
    with ProcessPoolExecutor(
        max_workers=workers,
        initializer=_worker_init,
        initargs=(tpl_bytes, cfg_bytes),
    ) as executor:
        future_to_spec: dict[Any, _ParallelTaskSpec] = {}
        for spec in specs:
            if spec.kind == "pages":
                future = executor.submit(
                    _worker_extract_pages,
                    spec.file_path,
                    spec.doc_type,
                    spec.index,
                    spec.page_indices or [],
                )
            else:
                future = executor.submit(
                    _worker_extract,
                    spec.file_path,
                    spec.doc_type,
                    spec.index,
                )
            future_to_spec[future] = spec

        for future in as_completed(future_to_spec):
            spec = future_to_spec[future]
            result = future.result()
            bucket = aggregated[spec.index]
            bucket["pending"] -= 1
            if result.ok:
                bucket["pages"].extend(result.pages)
                if spec.kind == "file":
                    bucket["single_elapsed_sec"] = result.elapsed_sec
                for key, value in result.timing_detail.items():
                    if isinstance(value, (int, float)):
                        _timing_add(bucket["timing_detail"], key, float(value))
                    else:
                        bucket["timing_detail"][key] = value
            else:
                bucket["error_messages"].append(result.error or "unknown")

            if bucket["pending"] == 0:
                meta = doc_meta.get(spec.index, {})
                pages = list(bucket["pages"])
                pages.sort(key=lambda item: item.page_num)
                if meta.get("mode") == "per_page":
                    elapsed_sec = time.perf_counter() - float(bucket["start_at"])
                else:
                    elapsed_sec = float(bucket.get("single_elapsed_sec") or 0.0)
                timing_detail = dict(bucket["timing_detail"])
                if meta.get("mode") == "per_page":
                    timing_detail.update(
                        {
                            "per_page_parallel_used": True,
                            "per_page_parallel_workers": meta.get("workers"),
                            "per_page_parallel_chunk_size": meta.get("chunk_size"),
                            "per_page_parallel_chunks": meta.get("chunks"),
                            "per_page_parallel_total": round(elapsed_sec, 6),
                        }
                    )
                error = "; ".join(bucket["error_messages"]) if bucket["error_messages"] else None
                file_result = V2FileExtractResult(
                    file_path=bucket["file_path"],
                    doc_type=bucket["doc_type"],
                    index=spec.index,
                    pages=pages,
                    elapsed_sec=elapsed_sec,
                    timing_detail=timing_detail,
                    error=error,
                )
                aggregated[spec.index]["result"] = file_result
                completed_docs += 1

                if callback is not None:
                    if file_result.ok:
                        affine_ui = stamp_affine_detail_for_ui(file_result.pages)
                        callback.on_file_done(
                            file_result.file_path,
                            file_result.index,
                            len(file_result.pages),
                            file_result.elapsed_sec,
                            affine_ui or None,
                        )
                    else:
                        callback.on_file_error(
                            file_result.file_path,
                            file_result.index,
                            RuntimeError(file_result.error or "unknown"),
                        )
                    callback.on_batch_progress(completed_docs, total_docs)

                if timing is not None:
                    affine = stamp_affine_uniformity_from_pages(file_result.pages)
                    timing.record_file(
                        os.path.basename(file_result.file_path),
                        file_result.elapsed_sec,
                        n_pages=len(file_result.pages),
                        error=file_result.error,
                        **file_result.timing_detail,
                        **affine,
                    )

    results: list[V2FileExtractResult] = [
        aggregated[idx]["result"]
        for idx in sorted(aggregated)
        if "result" in aggregated[idx]
    ]
    total_elapsed = time.perf_counter() - t_total
    if timing is not None:
        timing.record(
            "parallel_extraction_total",
            total_elapsed,
            n_workers=workers,
            n_files=total_docs,
        )
    return V2ParallelResult(
        file_results=results,
        total_elapsed_sec=total_elapsed,
        n_workers=workers,
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def run_parallel(
    tasks: list[tuple[str, str, int]],
    templates: list[StampTemplate],
    cfg: dict[str, Any],
    *,
    max_workers: int | None = None,
    callback: ParallelCallback | None = None,
    timing: Any | None = None,
) -> V2ParallelResult:
    """Run extraction for multiple PDF files in parallel.

    Args:
        tasks: list of ``(file_path, doc_type, original_index)``.
        templates: loaded stamp templates (pickle-safe).
        cfg: flat v2 config dict (shallow-copy-safe).
        max_workers: PPE workers; ``None`` -> auto (cpu_count - 1, max 8).
        callback: progress reporting (main process).
        timing: optional ``TimingCollector`` instance.

    Returns:
        V2ParallelResult with ``file_results`` sorted by original index.
    """
    workers = max_workers or _default_max_workers()
    workers = min(workers, len(tasks)) if tasks else 1

    tpl_bytes = pickle.dumps(templates)
    cfg_bytes = pickle.dumps(cfg)

    results: list[V2FileExtractResult] = []
    t_total = time.perf_counter()
    completed = 0
    total = len(tasks)

    with ProcessPoolExecutor(
        max_workers=workers,
        initializer=_worker_init,
        initargs=(tpl_bytes, cfg_bytes),
    ) as executor:
        future_to_idx: dict[Any, int] = {}
        for file_path, doc_type, idx in tasks:
            if callback is not None:
                callback.on_file_start(file_path, idx, total)
            fut = executor.submit(_worker_extract, file_path, doc_type, idx)
            future_to_idx[fut] = idx

        for future in as_completed(future_to_idx):
            result = future.result()
            results.append(result)
            completed += 1

            if callback is not None:
                if result.ok:
                    affine_ui = stamp_affine_detail_for_ui(result.pages)
                    callback.on_file_done(
                        result.file_path,
                        result.index,
                        len(result.pages),
                        result.elapsed_sec,
                        affine_ui or None,
                    )
                else:
                    callback.on_file_error(
                        result.file_path,
                        result.index,
                        RuntimeError(result.error or "unknown"),
                    )
                callback.on_batch_progress(completed, total)

            if timing is not None:
                affine = stamp_affine_uniformity_from_pages(result.pages)
                timing.record_file(
                    os.path.basename(result.file_path),
                    result.elapsed_sec,
                    n_pages=len(result.pages),
                    error=result.error,
                    **result.timing_detail,
                    **affine,
                )

    results.sort(key=lambda r: r.index)
    total_elapsed = time.perf_counter() - t_total

    if timing is not None:
        timing.record(
            "parallel_extraction_total",
            total_elapsed,
            n_workers=workers,
            n_files=total,
        )

    return V2ParallelResult(
        file_results=results,
        total_elapsed_sec=total_elapsed,
        n_workers=workers,
    )


def run_parallel_extraction(
    curr_proj: list[V2Document],
    templates: list[StampTemplate],
    cfg: dict[str, Any],
    *,
    max_workers: int | None = None,
    callback: ParallelCallback | None = None,
    timing: Any | None = None,
) -> None:
    """Parallel replacement for ``run_v2_extraction``.

    Mutates *curr_proj*: fills ``document.pages`` and ``document._v2_results``
    in the same way as the sequential version (deterministic order).

    Falls back to sequential for single-file inputs.
    """
    tasks = [
        (doc.file_full_path, doc.doc_Type, idx)
        for idx, doc in enumerate(curr_proj)
    ]

    if len(tasks) == 1:
        document = curr_proj[0]
        use_per_page, _per_page_cfg, n_pages = _should_use_per_page_parallel(document, cfg)
        if use_per_page:
            if callback is not None:
                callback.on_file_start(document.file_full_path, 0, 1)
            result = run_parallel_extraction_for_document(
                document,
                templates,
                cfg,
                max_workers=max_workers,
                timing=timing,
            )
            if result.ok:
                for v2r in result.pages:
                    page_data = V2PageData.from_v2_result(v2r, document)
                    document.pages.append(page_data)
                document._v2_results = result.pages
                if callback is not None:
                    affine_ui = stamp_affine_detail_for_ui(result.pages)
                    callback.on_file_done(
                        document.file_full_path,
                        0,
                        len(result.pages),
                        result.elapsed_sec,
                        affine_ui or None,
                    )
                    callback.on_batch_progress(1, 1)
                print(
                    "[parallel per-page] "
                    f"{os.path.basename(document.file_full_path)}: "
                    f"{n_pages} стр., {result.elapsed_sec:.1f}s"
                )
            else:
                if callback is not None:
                    callback.on_file_error(
                        document.file_full_path,
                        0,
                        RuntimeError(result.error or "unknown"),
                    )
                    callback.on_batch_progress(1, 1)
                print(
                    "[parallel per-page] "
                    f"{os.path.basename(document.file_full_path)}: ОШИБКА {result.error}"
                )
            return

    if len(tasks) < 2:
        from pdf_parsing_v2.v2_pipeline import run_v2_extraction

        run_v2_extraction(curr_proj, templates, cfg, callback=callback, timing=timing)
        return

    result = _run_parallel_mixed_extraction(
        curr_proj,
        templates,
        cfg,
        max_workers=max_workers,
        callback=callback,
        timing=timing,
    )

    for fr in result.file_results:
        doc = curr_proj[fr.index]
        if fr.ok:
            for v2r in fr.pages:
                page_data = V2PageData.from_v2_result(v2r, doc)
                doc.pages.append(page_data)
            doc._v2_results = fr.pages
        else:
            print(
                f"[parallel] {os.path.basename(fr.file_path)}: "
                f"ОШИБКА {fr.error}"
            )

    m, s = divmod(int(result.total_elapsed_sec), 60)
    summary = (
        f"[parallel] {result.n_files} файлов, "
        f"{result.n_pages_total} стр., "
        f"{result.n_workers} воркеров, {m}м {s:02d}с"
    )
    if result.n_errors:
        summary += f", ошибок: {result.n_errors}"
    print(summary)
