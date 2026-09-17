"""Child-process entry for catalog scans. Does not import Qt."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from rd_catalog.models import ScanProgress, SourceKind
from rd_catalog.perf_log import configure_perf_log
from rd_catalog.scan_job import config_from_job_dict, execute_scan


def _emit(payload: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def _arguments(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="RD catalog scan worker")
    parser.add_argument(
        "--job",
        type=Path,
        required=True,
        help="JSON file with resolved CatalogConfig paths and sources",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """Load a job file, run one scan, and stream JSON status lines to stdout."""

    args = _arguments(list(sys.argv[1:] if argv is None else argv))
    payload = json.loads(args.job.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        _emit({"t": "error", "m": "Job file must contain a JSON object."})
        return 2

    config = config_from_job_dict(payload)
    configure_perf_log(config.runtime_dir)
    sources = tuple(SourceKind(value) for value in payload.get("sources") or ())
    robot_subtree = payload.get("robot_subtree")
    rd_subtree = payload.get("rd_subtree")
    sq_subtree = payload.get("sq_subtree")
    cancel_path = Path(str(payload["cancel_path"])) if payload.get("cancel_path") else None

    def is_cancelled() -> bool:
        return bool(cancel_path is not None and cancel_path.exists())

    def log(message: str) -> None:
        _emit({"t": "log", "m": message})

    def progress(item: ScanProgress) -> None:
        _emit(
            {
                "t": "progress",
                "source": item.source.value,
                "path": item.path,
                "files_seen": item.files_seen,
                "message": item.message,
            }
        )

    outcome = execute_scan(
        config,
        sources,
        robot_subtree=robot_subtree,
        rd_subtree=rd_subtree,
        sq_subtree=sq_subtree,
        is_cancelled=is_cancelled,
        progress=progress,
        log=log,
    )
    if outcome.failure:
        _emit(
            {
                "t": "error",
                "m": f"{outcome.failure}\n{outcome.failure_traceback or ''}".rstrip(),
            }
        )
    _emit(
        {
            "t": "done",
            "run_id": outcome.run_id,
            "comparison_count": outcome.comparison_count,
            "failure": outcome.failure,
            "touched_mto_keys": [
                list(key) for key in sorted(outcome.touched_mto_keys)
            ],
            "mto_compare_skipped": outcome.mto_compare_skipped,
            "compare_enqueue_scope": outcome.compare_enqueue_scope,
        }
    )
    return 1 if outcome.failure else 0


if __name__ == "__main__":
    raise SystemExit(main())
