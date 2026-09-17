"""Benchmark sequential vs page-level parallel extraction for one PDF.

This tool does not modify the v2 pipeline. It reuses the same extraction
functions as production code and compares:

1. Sequential file loop via ``extract_all_pages_for_file()``.
2. Page-parallel tasks where each task opens the PDF for one page.
3. Chunked page-range workers where each worker opens the PDF once.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import pickle
import re
import statistics
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from typing import Any

import fitz

from pdf_parsing_v2.v2_config import load_v2_config, resolve_templates_dir
from pdf_parsing_v2_engine.models import StampTemplate, V2PageResult
from pdf_parsing_v2_engine.stamp_extractor import extract_all_pages_for_file, extract_page
from pdf_parsing_v2_engine.template_loader import load_all_templates

_WORKER_TEMPLATES: list[StampTemplate] | None = None
_WORKER_CFG: dict[str, Any] | None = None


def _timing_add(bucket: dict[str, float], key: str, value: float) -> None:
    bucket[key] = round(float(bucket.get(key, 0.0)) + float(value), 6)


def _infer_doc_type(path: str) -> str:
    match = re.search(r"\.([A-Z]+)-\d{4}", os.path.basename(path))
    if not match:
        raise ValueError(f"Cannot infer doc_type from {path!r}")
    return match.group(1)


def _page_fingerprint(page_result: V2PageResult) -> dict[str, Any]:
    return {
        "page_num": page_result.page_num,
        "doc_type": page_result.doc_type,
        "template_name": page_result.template_name,
        "template_score": round(float(page_result.template_score), 6),
        "warnings": list(page_result.warnings),
        "metadata": dict(sorted(page_result.metadata.items())),
        "fields": {
            field_id: {
                "raw_value": field_result.raw_value,
                "cleaned_value": field_result.cleaned_value,
                "bbox_pts": [round(float(v), 4) for v in field_result.bbox_pts],
                "is_valid": field_result.is_valid,
                "clean_tier": field_result.clean_tier,
                "parse_warnings": [
                    {
                        "code": warning.code,
                        "message": warning.message,
                        "details": dict(sorted(warning.details.items())),
                    }
                    for warning in field_result.parse_warnings
                ],
            }
            for field_id, field_result in sorted(page_result.fields.items())
        },
    }


def _results_digest(results: list[V2PageResult]) -> str:
    payload = [_page_fingerprint(result) for result in results]
    text = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _worker_init(templates_bytes: bytes, cfg_bytes: bytes) -> None:
    global _WORKER_TEMPLATES, _WORKER_CFG
    _WORKER_TEMPLATES = pickle.loads(templates_bytes)
    _WORKER_CFG = pickle.loads(cfg_bytes)


def _extract_page_indices(
    file_path: str,
    doc_type: str,
    page_indices: list[int],
) -> dict[str, Any]:
    if _WORKER_TEMPLATES is None or _WORKER_CFG is None:
        raise RuntimeError("Worker not initialized")

    results: list[V2PageResult] = []
    timing: dict[str, float] = {}
    t_open = time.perf_counter()
    doc = fitz.open(file_path)
    _timing_add(timing, "open_pdf", time.perf_counter() - t_open)
    try:
        page_cfg = dict(_WORKER_CFG)
        page_cfg["source_pdf_basename"] = os.path.basename(file_path)
        t_pages = time.perf_counter()
        for page_index in page_indices:
            page = doc[page_index]
            page_timing: dict[str, float] = {}
            result = extract_page(
                page,
                doc_type,
                page_index + 1,
                _WORKER_TEMPLATES,
                cfg=page_cfg,
                timing=page_timing,
            )
            results.append(result)
            for key, value in page_timing.items():
                _timing_add(timing, key, value)
        _timing_add(timing, "page_loop_total", time.perf_counter() - t_pages)
    finally:
        doc.close()

    return {
        "pages": results,
        "timing": timing,
        "page_indices": list(page_indices),
    }


def _split_page_indices(n_pages: int, n_workers: int) -> list[list[int]]:
    if n_pages <= 0:
        return []
    chunk_size = max(1, math.ceil(n_pages / max(1, n_workers)))
    chunks: list[list[int]] = []
    for start in range(0, n_pages, chunk_size):
        stop = min(start + chunk_size, n_pages)
        chunks.append(list(range(start, stop)))
    return chunks


def _run_parallel_mode(
    *,
    file_path: str,
    doc_type: str,
    n_pages: int,
    templates: list[StampTemplate],
    cfg: dict[str, Any],
    workers: int,
    mode: str,
) -> dict[str, Any]:
    if mode not in {"page_tasks", "chunked"}:
        raise ValueError(f"Unsupported mode: {mode}")

    if mode == "page_tasks":
        tasks = [[page_index] for page_index in range(n_pages)]
    else:
        tasks = _split_page_indices(n_pages, workers)

    tpl_bytes = pickle.dumps(templates)
    cfg_bytes = pickle.dumps(cfg)

    all_results: list[V2PageResult] = []
    timing_sum: dict[str, float] = {}
    wall_start = time.perf_counter()

    with ProcessPoolExecutor(
        max_workers=max(1, min(workers, len(tasks))),
        initializer=_worker_init,
        initargs=(tpl_bytes, cfg_bytes),
    ) as executor:
        futures = [
            executor.submit(_extract_page_indices, file_path, doc_type, page_indices)
            for page_indices in tasks
        ]
        for future in as_completed(futures):
            chunk = future.result()
            all_results.extend(chunk["pages"])
            for key, value in chunk["timing"].items():
                _timing_add(timing_sum, key, value)

    all_results.sort(key=lambda item: item.page_num)
    elapsed = time.perf_counter() - wall_start
    return {
        "mode": mode,
        "workers": workers,
        "elapsed_sec": round(elapsed, 6),
        "n_pages": len(all_results),
        "result_digest": _results_digest(all_results),
        "timing_sum": timing_sum,
    }


def _run_sequential(
    file_path: str,
    doc_type: str,
    templates: list[StampTemplate],
    cfg: dict[str, Any],
    *,
    with_timing: bool,
) -> dict[str, Any]:
    wall_start = time.perf_counter()
    extracted = extract_all_pages_for_file(
        file_path,
        doc_type,
        templates,
        cfg,
        with_timing=with_timing,
    )
    elapsed = time.perf_counter() - wall_start
    if with_timing:
        results, timing = extracted
    else:
        results, timing = extracted, {}
    return {
        "mode": "sequential",
        "workers": 1,
        "elapsed_sec": round(elapsed, 6),
        "n_pages": len(results),
        "result_digest": _results_digest(results),
        "timing_sum": timing,
    }


def _mean(values: list[float]) -> float:
    return round(statistics.mean(values), 6) if values else 0.0


def _stdev(values: list[float]) -> float:
    return round(statistics.pstdev(values), 6) if len(values) > 1 else 0.0


def benchmark_file(
    file_path: str,
    templates: list[StampTemplate],
    cfg: dict[str, Any],
    *,
    repeats: int,
    workers: list[int],
) -> dict[str, Any]:
    doc_type = _infer_doc_type(file_path)
    seq_detail = _run_sequential(file_path, doc_type, templates, cfg, with_timing=True)
    n_pages = int(seq_detail["n_pages"])

    runs: list[dict[str, Any]] = [dict(seq_detail, run_index=0, digest_match=True)]

    for run_index in range(repeats):
        seq_run = _run_sequential(file_path, doc_type, templates, cfg, with_timing=False)
        seq_run["run_index"] = run_index + 1
        seq_run["digest_match"] = seq_run["result_digest"] == seq_detail["result_digest"]
        runs.append(seq_run)

        for worker_count in workers:
            effective_workers = max(1, min(worker_count, n_pages))
            for mode in ("page_tasks", "chunked"):
                parallel_run = _run_parallel_mode(
                    file_path=file_path,
                    doc_type=doc_type,
                    n_pages=n_pages,
                    templates=templates,
                    cfg=cfg,
                    workers=effective_workers,
                    mode=mode,
                )
                parallel_run["run_index"] = run_index + 1
                parallel_run["digest_match"] = (
                    parallel_run["result_digest"] == seq_detail["result_digest"]
                )
                runs.append(parallel_run)

    aggregated: dict[str, dict[str, Any]] = {}
    for run in runs:
        key = f"{run['mode']}|{run['workers']}"
        bucket = aggregated.setdefault(
            key,
            {
                "mode": run["mode"],
                "workers": run["workers"],
                "elapsed_values": [],
                "digest_match_all": True,
            },
        )
        if run["run_index"] > 0:
            bucket["elapsed_values"].append(run["elapsed_sec"])
            bucket["digest_match_all"] = bucket["digest_match_all"] and bool(
                run["digest_match"]
            )

    summary_rows: list[dict[str, Any]] = []
    seq_mean = 0.0
    for bucket in aggregated.values():
        row = {
            "mode": bucket["mode"],
            "workers": bucket["workers"],
            "mean_elapsed_sec": _mean(bucket["elapsed_values"]),
            "stdev_elapsed_sec": _stdev(bucket["elapsed_values"]),
            "min_elapsed_sec": round(min(bucket["elapsed_values"]), 6)
            if bucket["elapsed_values"]
            else 0.0,
            "max_elapsed_sec": round(max(bucket["elapsed_values"]), 6)
            if bucket["elapsed_values"]
            else 0.0,
            "digest_match_all": bucket["digest_match_all"],
        }
        if row["mode"] == "sequential":
            seq_mean = row["mean_elapsed_sec"]
        summary_rows.append(row)

    for row in summary_rows:
        if seq_mean > 0:
            row["speedup_vs_sequential"] = round(seq_mean / row["mean_elapsed_sec"], 3)
        else:
            row["speedup_vs_sequential"] = 0.0

    summary_rows.sort(key=lambda item: (item["mode"], item["workers"]))

    return {
        "file_path": file_path,
        "file_name": os.path.basename(file_path),
        "doc_type": doc_type,
        "file_size_mb": round(Path(file_path).stat().st_size / (1024 * 1024), 3),
        "n_pages": n_pages,
        "sequential_detail": seq_detail,
        "runs": runs,
        "summary": summary_rows,
    }


def _parse_workers(raw: str) -> list[int]:
    values = []
    for item in raw.split(","):
        item = item.strip()
        if not item:
            continue
        values.append(max(1, int(item)))
    return values or [2, 4]


def main() -> int:
    parser = argparse.ArgumentParser(description="Benchmark page-level parallel extraction")
    parser.add_argument("pdf_files", nargs="+", help="PDF files to benchmark")
    parser.add_argument(
        "--repeats",
        type=int,
        default=2,
        help="Repeat count per mode after one detailed baseline run",
    )
    parser.add_argument(
        "--workers",
        default="2,4",
        help="Comma-separated worker counts for parallel modes",
    )
    parser.add_argument(
        "--json-out",
        default="",
        help="Optional path to save full JSON report",
    )
    args = parser.parse_args()

    cfg = load_v2_config()
    templates = load_all_templates(resolve_templates_dir(cfg))

    report = {
        "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "repeats": int(args.repeats),
        "workers": _parse_workers(args.workers),
        "files": [],
    }

    for pdf_file in args.pdf_files:
        result = benchmark_file(
            os.path.abspath(pdf_file),
            templates,
            cfg,
            repeats=max(1, int(args.repeats)),
            workers=_parse_workers(args.workers),
        )
        report["files"].append(result)

    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.json_out:
        out_path = os.path.abspath(args.json_out)
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        with open(out_path, "w", encoding="utf-8") as fh:
            fh.write(text)
        print(out_path)
    else:
        print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
