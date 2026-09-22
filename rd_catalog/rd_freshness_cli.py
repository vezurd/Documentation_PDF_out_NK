"""Child-process RD file-stamp census. Does not import Qt."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from rd_catalog.rd_freshness import (
    RdFileStamp,
    collect_rd_file_stamps,
    dirty_mark_folders,
)


def _emit(payload: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def _arguments(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="RD file-stamp census")
    parser.add_argument("--job", type=Path, required=True)
    return parser.parse_args(argv)


def _stamps(raw: object) -> tuple[RdFileStamp, ...]:
    if not isinstance(raw, list):
        return ()
    stamps: list[RdFileStamp] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        path = str(item.get("path") or "").strip()
        path_key = str(item.get("path_key") or "").strip()
        if not path or not path_key:
            continue
        stamps.append(
            RdFileStamp(
                path=path,
                path_key=path_key,
                size=int(item.get("size") or 0),
                mtime_ns=int(item.get("mtime_ns") or 0),
            )
        )
    return tuple(stamps)


def main(argv: list[str] | None = None) -> int:
    """Compare on-disk RD size and mtime with the job's baseline snapshot."""

    args = _arguments(list(sys.argv[1:] if argv is None else argv))
    try:
        payload = json.loads(args.job.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        _emit({"t": "error", "m": f"Не прочитан job: {exc}"})
        return 2
    if not isinstance(payload, dict):
        _emit({"t": "error", "m": "Job file must contain a JSON object."})
        return 2
    rd_root = str(payload.get("rd_root") or "")
    sq_root = str(payload.get("sq_root") or "") or None
    skip_dirs = payload.get("skip_dirs") or ()
    if not isinstance(skip_dirs, list):
        skip_dirs = []
    cancel_path = Path(str(payload["cancel_path"])) if payload.get("cancel_path") else None
    baseline = _stamps(payload.get("baseline"))

    def is_cancelled() -> bool:
        return bool(cancel_path is not None and cancel_path.exists())

    census = collect_rd_file_stamps(
        rd_root,
        skip_dirs=skip_dirs,
        sq_root=sq_root,
        is_cancelled=is_cancelled,
    )
    for message in census.errors[:8]:
        _emit({"t": "log", "m": f"Автоперескан: {message}"})
    if not census.completed:
        _emit({"t": "done", "seen": len(census.stamps), "folders": [], "cancelled": True})
        return 1
    folders = dirty_mark_folders(baseline, census.stamps, rd_root)
    _emit(
        {
            "t": "done",
            "seen": len(census.stamps),
            "folders": list(folders),
            "cancelled": False,
        }
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
