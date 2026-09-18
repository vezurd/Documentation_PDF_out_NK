"""Child-process entry for AN dump scans. Does not import Qt."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from rd_catalog.an_scan import (
    AnScanProgress,
    config_from_an_job,
    scan_an_dump,
)
from rd_catalog.perf_log import configure_perf_log


def _emit(payload: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def _arguments(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="RD catalog AN dump scan worker (MTO xlsx + OD doc)"
    )
    parser.add_argument(
        "--job",
        type=Path,
        required=True,
        help="JSON file with CatalogConfig paths, an_root, and optional cancel_path",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """Load a job file, walk the AN dump, and stream JSON status lines.

    Args:
        argv: Optional argument list (defaults to ``sys.argv[1:]``).

    Returns:
        ``2`` if the job file is unusable, ``1`` if the walk reported
        ``failure``, otherwise ``0``.
    """

    args = _arguments(list(sys.argv[1:] if argv is None else argv))
    try:
        payload = json.loads(args.job.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        _emit({"t": "error", "m": f"Invalid job file: {exc}"})
        return 2
    if not isinstance(payload, dict):
        _emit({"t": "error", "m": "Job file must contain a JSON object."})
        return 2
    try:
        config = config_from_an_job(payload)
    except (KeyError, TypeError, ValueError) as exc:
        _emit({"t": "error", "m": f"Invalid job: {exc}"})
        return 2
    configure_perf_log(config.runtime_dir)

    cancel_path = (
        Path(str(payload["cancel_path"])) if payload.get("cancel_path") else None
    )

    def is_cancelled() -> bool:
        return bool(cancel_path is not None and cancel_path.exists())

    def log(message: str) -> None:
        _emit({"t": "log", "m": message})

    def progress(item: AnScanProgress) -> None:
        _emit(
            {
                "t": "progress",
                "path": item.path,
                "files_seen": item.files_seen,
                "message": item.message,
            }
        )

    outcome = scan_an_dump(
        config,
        persist=True,
        is_cancelled=is_cancelled,
        progress=progress,
        log=log,
    )
    if outcome.failure:
        _emit({"t": "error", "m": outcome.failure})
    _emit(
        {
            "t": "done",
            "accepted": outcome.accepted,
            "files_seen": outcome.files_seen,
            "failure": outcome.failure,
            "scanned_at": outcome.scanned_at,
            "cancelled": outcome.cancelled,
        }
    )
    return 1 if outcome.failure else 0


if __name__ == "__main__":
    raise SystemExit(main())
