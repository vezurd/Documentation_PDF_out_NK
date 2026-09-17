"""Disk cache for AN-vs-PI / AN-vs-RD pairwise content compares.

Qt-free. The JSON lives under the catalog runtime directory, never on UNC.
"""

from __future__ import annotations

import json
from collections.abc import Mapping
from pathlib import Path
from typing import Any, Literal

from rd_catalog.auto_mto_compare_cache import rd_mtime_ns
from rd_catalog.customer_pi_auto_mto import (
    AUTO_MTO_COMPARE_ALGORITHM_VERSION,
    MtoPairCompareResult,
    auto_mto_rd_path_key,
)

AN_CONTENT_COMPARE_CACHE_VERSION = 1
AN_CONTENT_COMPARE_CACHE_NAME = "an_content_compare_cache.json"
RD_DUMP_CONTENT_COMPARE_CACHE_NAME = "rd_dump_content_compare_cache.json"
_PERSISTABLE_KINDS = frozenset({"matched", "soft", "no_match", "not_compared"})


def cache_key(
    an_path: str,
    an_mtime_ns: int,
    auto_path: str,
    auto_mtime_ns: int,
    rd_path: str,
    rd_mtime_ns_value: int,
    *,
    algorithm_version: int = AUTO_MTO_COMPARE_ALGORITHM_VERSION,
) -> str:
    """Return the stable disk-cache identity for one AN file compare.

    Args:
        an_path: AN workbook path.
        an_mtime_ns: AN ``st_mtime_ns``, or ``0``.
        auto_path: AutoMTO counterpart path, or empty.
        auto_mtime_ns: AutoMTO ``st_mtime_ns``, or ``0``.
        rd_path: RD MTO counterpart path, or empty.
        rd_mtime_ns_value: RD ``st_mtime_ns``, or ``0``.
        algorithm_version: ``AUTO_MTO_COMPARE_ALGORITHM_VERSION``.

    Returns:
        JSON object string used as the cache map key.
    """

    payload = {
        "algorithm": int(algorithm_version),
        "an": auto_mto_rd_path_key(an_path),
        "an_mtime_ns": int(an_mtime_ns or 0),
        "auto": auto_mto_rd_path_key(auto_path),
        "auto_mtime_ns": int(auto_mtime_ns or 0),
        "rd": auto_mto_rd_path_key(rd_path),
        "rd_mtime_ns": int(rd_mtime_ns_value or 0),
    }
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True)


def result_to_entry(result: MtoPairCompareResult | None) -> dict[str, Any] | None:
    """Serialize one side of a pairwise compare, or ``None``."""

    if result is None:
        return None
    if result.kind not in _PERSISTABLE_KINDS:
        return None
    return {
        "kind": result.kind,
        "grade": result.grade,
        "left_rows": int(result.left_rows),
        "right_rows": int(result.right_rows),
        "error": str(result.error or ""),
    }


def result_from_entry(entry: Mapping[str, Any] | None) -> MtoPairCompareResult | None:
    """Rebuild a pairwise result from a cache side-object."""

    if not entry:
        return None
    kind = str(entry.get("kind") or "")
    if kind not in _PERSISTABLE_KINDS:
        return None
    grade_raw = str(entry.get("grade") or "")
    grade: Literal["exact", "soft", "none"]
    if kind == "matched":
        grade = "exact"
    elif kind == "soft":
        grade = "soft"
    else:
        grade = "none" if grade_raw != "soft" else "soft"
    return MtoPairCompareResult(
        kind=kind,  # type: ignore[arg-type]
        grade=grade,
        left_rows=int(entry.get("left_rows") or 0),
        right_rows=int(entry.get("right_rows") or 0),
        error=str(entry.get("error") or ""),
    )


def cache_path(
    runtime_dir: str | Path,
    *,
    name: str = AN_CONTENT_COMPARE_CACHE_NAME,
) -> Path:
    """Return ``runtime_dir / name`` for a pairwise content-compare JSON."""

    return Path(runtime_dir) / name


def load_an_content_compare_cache(
    runtime_dir: str | Path | None,
    *,
    name: str = AN_CONTENT_COMPARE_CACHE_NAME,
) -> dict[str, dict[str, Any]]:
    """Load persistable AN content-compare entries from the runtime JSON."""

    if runtime_dir is None:
        return {}
    path = cache_path(runtime_dir, name=name)
    if not path.is_file():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return {}
    if not isinstance(raw, dict):
        return {}
    if int(raw.get("version") or 0) != AN_CONTENT_COMPARE_CACHE_VERSION:
        return {}
    entries = raw.get("entries")
    if not isinstance(entries, dict):
        return {}
    cleaned: dict[str, dict[str, Any]] = {}
    for key, entry in entries.items():
        if isinstance(entry, dict):
            cleaned[str(key)] = dict(entry)
    return cleaned


def save_an_content_compare_cache(
    runtime_dir: str | Path | None,
    entries: Mapping[str, Mapping[str, Any]],
    *,
    name: str = AN_CONTENT_COMPARE_CACHE_NAME,
) -> None:
    """Write persistable AN content-compare entries to the runtime JSON."""

    if runtime_dir is None:
        return
    cleaned = {
        str(key): dict(entry)
        for key, entry in entries.items()
        if isinstance(entry, Mapping)
    }
    path = cache_path(runtime_dir, name=name)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": AN_CONTENT_COMPARE_CACHE_VERSION,
        "algorithm_version": AUTO_MTO_COMPARE_ALGORITHM_VERSION,
        "entries": cleaned,
    }
    text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    tmp_path = path.with_name(path.name + ".tmp")
    tmp_path.write_text(text, encoding="utf-8")
    tmp_path.replace(path)


def counterpart_mtime_ns(path: str) -> int:
    """Return ``st_mtime_ns`` for a counterpart path, or ``0``."""

    return rd_mtime_ns(path)
