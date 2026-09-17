"""Cache and labels for MTO content compare between disputed transfer packages.

Qt-free. Uses the AutoMTO pairwise grades (``четкое`` / ``ПоКоду и Кол-ву`` /
``не совпало``). The JSON lives under the catalog runtime directory, never
on UNC. Workbooks are opened only by the background thread.
"""

from __future__ import annotations

import json
from collections.abc import Mapping
from pathlib import Path
from typing import Any, Literal

from rd_catalog.customer_pi_auto_mto import (
    AUTO_MTO_COMPARE_ALGORITHM_VERSION,
    MtoPairCompareResult,
    compare_mto_pair,
)
from rd_catalog.models import make_path_key
from rd_catalog.parse import normalize_unicode_dashes

TRANSFER_REVIEW_COMPARE_CACHE_VERSION = 1
TRANSFER_REVIEW_COMPARE_CACHE_NAME = "transfer_review_mto_compare_cache.json"
_PERSISTABLE_KINDS = frozenset({"matched", "soft", "no_match", "not_compared"})


def pair_path_id(path: str) -> str:
    """Return a case-insensitive identity for one MTO path.

    Args:
        path: Local or UNC workbook path.

    Returns:
        Normalized path key, or ``""``.
    """

    text = str(path or "").strip()
    if not text:
        return ""
    return make_path_key(normalize_unicode_dashes(text))


def cache_entry_key(
    left_path: str,
    left_mtime_ns: int,
    right_path: str,
    right_mtime_ns: int,
    *,
    algorithm_version: int = AUTO_MTO_COMPARE_ALGORITHM_VERSION,
) -> str:
    """Return the disk-cache identity for one disputed-package MTO pair.

    Args:
        left_path: First RD MTO path.
        left_mtime_ns: First file ``mtime_ns`` from the scan.
        right_path: Second RD MTO path.
        right_mtime_ns: Second file ``mtime_ns``.
        algorithm_version: ``AUTO_MTO_COMPARE_ALGORITHM_VERSION``.

    Returns:
        Stable JSON object string used as the cache map key.
    """

    left_id = pair_path_id(left_path)
    right_id = pair_path_id(right_path)
    left_mtime = int(left_mtime_ns or 0)
    right_mtime = int(right_mtime_ns or 0)
    if (left_id, left_mtime) <= (right_id, right_mtime):
        first, first_mtime, second, second_mtime = (
            left_id,
            left_mtime,
            right_id,
            right_mtime,
        )
    else:
        first, first_mtime, second, second_mtime = (
            right_id,
            right_mtime,
            left_id,
            left_mtime,
        )
    payload = {
        "algorithm": int(algorithm_version),
        "left": first,
        "left_mtime_ns": first_mtime,
        "right": second,
        "right_mtime_ns": second_mtime,
    }
    return json.dumps(
        payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True
    )


def path_pair_key(left_path: str, right_path: str) -> tuple[str, str]:
    """Return a sorted identity pair for two MTO paths.

    Args:
        left_path: First workbook path.
        right_path: Second workbook path.

    Returns:
        Two ``pair_path_id`` values in sorted order.
    """

    left_id = pair_path_id(left_path)
    right_id = pair_path_id(right_path)
    return (left_id, right_id) if left_id <= right_id else (right_id, left_id)


def result_to_entry(result: MtoPairCompareResult | None) -> dict[str, Any] | None:
    """Serialize one pairwise compare, or ``None``."""

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
    """Rebuild a pairwise result from a cache object."""

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


def cache_path(runtime_dir: str | Path) -> Path:
    """Return ``runtime_dir / transfer_review_mto_compare_cache.json``."""

    return Path(runtime_dir) / TRANSFER_REVIEW_COMPARE_CACHE_NAME


def load_transfer_review_compare_cache(
    runtime_dir: str | Path | None,
) -> dict[str, dict[str, Any]]:
    """Load persistable pairwise entries from the runtime JSON."""

    if runtime_dir is None:
        return {}
    path = cache_path(runtime_dir)
    if not path.is_file():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, TypeError):
        return {}
    if not isinstance(payload, dict):
        return {}
    if int(payload.get("version") or 0) != TRANSFER_REVIEW_COMPARE_CACHE_VERSION:
        return {}
    entries = payload.get("entries")
    if not isinstance(entries, dict):
        return {}
    result: dict[str, dict[str, Any]] = {}
    for key, entry in entries.items():
        if isinstance(key, str) and isinstance(entry, dict):
            result[key] = entry
    return result


def save_transfer_review_compare_cache(
    runtime_dir: str | Path | None,
    entries: Mapping[str, Mapping[str, Any]],
) -> None:
    """Write persistable pairwise entries to the runtime JSON."""

    if runtime_dir is None:
        return
    path = cache_path(runtime_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": TRANSFER_REVIEW_COMPARE_CACHE_VERSION,
        "entries": {str(key): dict(entry) for key, entry in entries.items()},
    }
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def labels_from_cache(
    runtime_dir: str | Path | None,
) -> dict[str, str]:
    """Return ``cache_entry_key`` → Russian compare label.

    Args:
        runtime_dir: Catalog runtime directory, or ``None``.

    Returns:
        Labels for cache hits; misses are omitted.
    """

    labels: dict[str, str] = {}
    for key, entry in load_transfer_review_compare_cache(runtime_dir).items():
        result = result_from_entry(entry)
        if result is None:
            continue
        labels[key] = result.paren_label
    return labels


def path_pair_labels_from_cache(
    runtime_dir: str | Path | None,
) -> dict[tuple[str, str], str]:
    """Return sorted path-id pair → Russian compare label.

    Args:
        runtime_dir: Catalog runtime directory, or ``None``.

    Returns:
        Path-pair labels for cache hits.
    """

    labels: dict[tuple[str, str], str] = {}
    for key, entry in load_transfer_review_compare_cache(runtime_dir).items():
        result = result_from_entry(entry)
        if result is None:
            continue
        try:
            payload = json.loads(key)
            left = str(payload.get("left") or "")
            right = str(payload.get("right") or "")
        except (json.JSONDecodeError, TypeError):
            continue
        if not left or not right:
            continue
        pair = (left, right) if left <= right else (right, left)
        labels[pair] = result.paren_label
    return labels


def compare_transfer_mto_pair(
    left_path: str,
    right_path: str,
) -> MtoPairCompareResult:
    """Compare two RD MTO workbooks with the AutoMTO grades.

    Args:
        left_path: First workbook.
        right_path: Second workbook.

    Returns:
        Pairwise compare result (does not raise on load errors).
    """

    return compare_mto_pair(left_path, right_path)
