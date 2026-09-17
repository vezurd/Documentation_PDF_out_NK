"""Child-process entry for deferred MTO content compare. Does not import Qt."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from rd_catalog.mto_compare_job import (
    _keys_from_payload,
    config_from_job_dict,
    execute_mto_compare,
)


def _emit(payload: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def _arguments(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="RD catalog MTO compare worker")
    parser.add_argument(
        "--job",
        type=Path,
        required=True,
        help="JSON file with resolved CatalogConfig paths and compare scope",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """Load a job file, run one MTO compare, and stream JSON status lines."""

    args = _arguments(list(sys.argv[1:] if argv is None else argv))
    payload = json.loads(args.job.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        _emit({"t": "error", "m": "Job file must contain a JSON object."})
        return 2

    config = config_from_job_dict(payload)
    document_keys = _keys_from_payload(payload.get("document_keys"))
    priority_keys = _keys_from_payload(payload.get("priority_keys")) or []
    cancel_path = (
        Path(str(payload["cancel_path"])) if payload.get("cancel_path") else None
    )
    batch_size = int(payload.get("batch_size") or 1)

    def is_cancelled() -> bool:
        return bool(cancel_path is not None and cancel_path.exists())

    def log(message: str) -> None:
        _emit({"t": "log", "m": message})

    def progress(completed: int, total: int) -> None:
        _emit({"t": "progress", "completed": completed, "total": total})

    outcome = execute_mto_compare(
        config,
        document_keys=document_keys,
        priority_keys=priority_keys,
        batch_size=batch_size,
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
            "comparison_count": outcome.comparison_count,
            "planned": outcome.planned,
            "skipped_cache_hits": outcome.skipped_cache_hits,
            "cancelled": outcome.cancelled,
            "failure": outcome.failure,
            "remaining_keys": [list(key) for key in outcome.remaining_keys],
        }
    )
    return 1 if outcome.failure else 0


if __name__ == "__main__":
    raise SystemExit(main())
