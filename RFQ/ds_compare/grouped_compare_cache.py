"""Pickle cache for grouped DS vs MTO comparison result (before RFQ enrichment)."""

from __future__ import annotations

import hashlib
import os
import pickle
from pathlib import Path
from typing import Any

from base.base_classes import RowStd
from RFQ.ds_compare.get_mto_list_from_ds import _NOT_FOUND_MARKER

_GROUPED_COMPARE_CACHE_DIR = Path(__file__).resolve().parent / "cache"
GROUPED_COMPARE_CACHE_VERSION = "v1"


def _cache_dir() -> Path:
    _GROUPED_COMPARE_CACHE_DIR.mkdir(exist_ok=True)
    return _GROUPED_COMPARE_CACHE_DIR


def _path_hash(path: str) -> str:
    normalized = os.path.normpath(os.path.abspath(path))
    return hashlib.md5(normalized.encode(errors="replace")).hexdigest()[:16]


def _cache_file_path(ds_path: str, extra_key: str) -> Path:
    extra_hash = hashlib.md5(extra_key.encode(errors="replace")).hexdigest()[:12]
    return _cache_dir() / f"grouped_cmp_{_path_hash(ds_path)}_{extra_hash}.cache"


def build_grouped_compare_extra_key(
    *,
    group_keys: list[str],
    flat_mto_structure: bool,
    group_mto_new_positions: bool,
    mto_path: str,
    replacement_table_file: str,
) -> str:
    """Stable config fingerprint for cache file name."""
    parts = (
        GROUPED_COMPARE_CACHE_VERSION,
        ",".join(group_keys),
        f"flat={int(flat_mto_structure)}",
        f"mto_new={int(group_mto_new_positions)}",
        mto_path,
        replacement_table_file,
    )
    return "|".join(parts)


def mto_spec_fingerprint(spec_dict: dict[str, Any]) -> tuple[tuple[str, str, float], ...]:
    """Collect spec key, MTO path and mtime for cache invalidation."""
    items: list[tuple[str, str, float]] = []
    for spec in sorted(spec_dict):
        path = spec_dict[spec]
        if path == _NOT_FOUND_MARKER or not path:
            items.append((spec, str(path), 0.0))
            continue
        path_str = str(path)
        if os.path.exists(path_str):
            items.append((spec, path_str, os.path.getmtime(path_str)))
        else:
            items.append((spec, path_str, -1.0))
    return tuple(items)


def _replacement_mtime(replacement_table_file: str) -> float:
    if replacement_table_file and os.path.exists(replacement_table_file):
        return os.path.getmtime(replacement_table_file)
    return 0.0


def _meta_matches(
    meta: dict[str, Any],
    *,
    ds_path: str,
    replacement_table_file: str,
    spec_dict: dict[str, Any],
) -> bool:
    if not isinstance(meta, dict):
        return False
    if meta.get("version") != GROUPED_COMPARE_CACHE_VERSION:
        return False
    if not os.path.exists(ds_path):
        return False
    ds_mtime = os.path.getmtime(ds_path)
    if abs(float(meta.get("ds_mtime", -1)) - ds_mtime) >= 1e-6:
        return False
    repl_mtime = _replacement_mtime(replacement_table_file)
    if abs(float(meta.get("replacement_mtime", -1)) - repl_mtime) >= 1e-6:
        return False
    cached_fp = meta.get("mto_fingerprint")
    current_fp = mto_spec_fingerprint(spec_dict)
    return cached_fp == current_fp


def grouped_compare_cache_file_path(ds_path: str, extra_key: str) -> Path:
    """Return path to grouped compare pickle for ``ds_path`` and config fingerprint."""
    return _cache_file_path(ds_path, extra_key)


def load_grouped_compare_cache(
    ds_path: str,
    extra_key: str,
    *,
    replacement_table_file: str,
    spec_dict: dict[str, Any],
) -> list[RowStd] | None:
    """Return cached compared rows or None if missing/invalid."""
    cache_path = _cache_file_path(ds_path, extra_key)
    if not cache_path.exists():
        return None
    try:
        with open(cache_path, "rb") as f:
            payload = pickle.load(f)
        if not isinstance(payload, dict) or "data" not in payload:
            return None
        meta = payload.get("meta")
        if not _meta_matches(
            meta,
            ds_path=ds_path,
            replacement_table_file=replacement_table_file,
            spec_dict=spec_dict,
        ):
            return None
        rows = payload["data"]
        if not isinstance(rows, list):
            return None
        print(f"Grouped DS vs MTO: результат загружен из кэша ({len(rows)} строк)")
        return rows
    except Exception as exc:
        print(f"Grouped DS vs MTO: ошибка чтения кэша — {exc}")
        return None


def save_grouped_compare_cache(
    ds_path: str,
    extra_key: str,
    rows: list[RowStd],
    *,
    replacement_table_file: str,
    spec_dict: dict[str, Any],
) -> None:
    """Persist compared rows with DS/MTO/replacement fingerprints."""
    cache_path = _cache_file_path(ds_path, extra_key)
    meta = {
        "version": GROUPED_COMPARE_CACHE_VERSION,
        "ds_mtime": os.path.getmtime(ds_path) if os.path.exists(ds_path) else 0.0,
        "replacement_mtime": _replacement_mtime(replacement_table_file),
        "mto_fingerprint": mto_spec_fingerprint(spec_dict),
    }
    payload = {"meta": meta, "data": rows}
    try:
        with open(cache_path, "wb") as f:
            pickle.dump(payload, f)
        print(f"Grouped DS vs MTO: результат сохранён в кэш ({cache_path.name})")
    except Exception as exc:
        print(f"Grouped DS vs MTO: ошибка записи кэша — {exc}")
