"""Disk cache for persistable AutoMTO content-compare results.

Qt-free. The JSON lives under the catalog runtime directory, never on UNC.
Identity includes kit, RD path, RD mtime, AutoMTO fingerprints, and the
compare algorithm version so a file or algorithm change misses the cache.
"""

from __future__ import annotations

import json
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any, Literal

from rd_catalog.customer_pi_auto_mto import (
    AUTO_MTO_COMPARE_ALGORITHM_VERSION,
    AutoMtoCompareResult,
    AutoMtoFile,
)
from rd_catalog.kits import kit_identity_key

AUTO_MTO_COMPARE_CACHE_VERSION = 1
AUTO_MTO_COMPARE_CACHE_NAME = "auto_mto_compare_cache.json"
_PERSISTABLE_KINDS = frozenset({"single", "composite", "no_match"})


def rd_mtime_ns(path: str) -> int:
    """Return ``st_mtime_ns`` for ``path``, or ``0`` if unknown.

    Args:
        path: RD MTO workbook path (local or UNC).

    Returns:
        Nanosecond mtime, or ``0`` when the file cannot be stat'd.
    """

    if not path:
        return 0
    try:
        return int(Path(path).stat().st_mtime_ns)
    except OSError:
        return 0


def _files_signature(
    files: Sequence[AutoMtoFile],
) -> tuple[tuple[str, str], ...]:
    return tuple(sorted((str(item.relpath), str(item.fingerprint)) for item in files))


def cache_key(
    title: str,
    mark: str,
    rd_path: str,
    rd_mtime_ns_value: int,
    files: Sequence[AutoMtoFile],
    *,
    algorithm_version: int = AUTO_MTO_COMPARE_ALGORITHM_VERSION,
) -> str:
    """Return the stable disk-cache identity for one kit comparison.

    Args:
        title: Kit title (case-insensitive).
        mark: Kit mark (case-insensitive).
        rd_path: RD MTO path; stored casefolded.
        rd_mtime_ns_value: RD file ``st_mtime_ns``, or ``0`` if unknown.
        files: AutoMTO files used for the compare (kit index).
        algorithm_version: ``AUTO_MTO_COMPARE_ALGORITHM_VERSION``.

    Returns:
        JSON object string used as the cache map key.
    """

    kit = kit_identity_key(title, mark)
    payload = {
        "algorithm": int(algorithm_version),
        "files": [list(item) for item in _files_signature(files)],
        "kit": [kit[0], kit[1]],
        "rd_mtime_ns": int(rd_mtime_ns_value or 0),
        "rd_path": str(rd_path).casefold(),
    }
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True)


def entry_from_result(
    result: AutoMtoCompareResult,
    *,
    title: str = "",
    mark: str = "",
    rd_path: str = "",
) -> dict[str, Any] | None:
    """Build a persistable cache entry, or ``None`` for ``not_compared``.

    Args:
        result: In-memory compare result.
        title: Optional kit title stored for hydrate-by-scan.
        mark: Optional kit mark stored for hydrate-by-scan.
        rd_path: Optional RD path stored for hydrate-by-scan.

    Returns:
        JSON-ready mapping, or ``None`` when the kind must not be cached.
    """

    if result.match_kind not in _PERSISTABLE_KINDS:
        return None
    kit = kit_identity_key(title, mark) if title and mark else ("", "")
    return {
        "match_kind": result.match_kind,
        "member_relpaths": [member.relpath for member in result.members],
        "rd_rows": int(result.rd_rows),
        "auto_rows": int(result.auto_rows),
        "combinations_checked": int(result.combinations_checked),
        "error": str(result.error or ""),
        "grade": result.content_grade,
        "title": kit[0],
        "mark": kit[1],
        "rd_path": str(rd_path).casefold(),
    }


def result_from_entry(
    entry: Mapping[str, Any] | None,
    files: Sequence[AutoMtoFile],
) -> AutoMtoCompareResult | None:
    """Rebuild a compare result from a cache entry.

    Args:
        entry: Persisted mapping, or ``None``.
        files: Current AutoMTO index for the kit.

    Returns:
        Result when ``match_kind`` is persistable and every member relpath
        is present in ``files``. ``None`` if a member is missing or the
        kind is not cacheable (including ``not_compared``).
    """

    if not entry:
        return None
    kind = str(entry.get("match_kind") or "")
    if kind not in _PERSISTABLE_KINDS:
        return None
    by_relpath = {item.relpath: item for item in files}
    members: list[AutoMtoFile] = []
    for relpath in entry.get("member_relpaths") or ():
        member = by_relpath.get(str(relpath))
        if member is None:
            return None
        members.append(member)
    grade_raw = str(entry.get("grade") or "")
    if kind in {"single", "composite"}:
        grade: Literal["exact", "soft", "none"] = (
            "soft" if grade_raw == "soft" else "exact"
        )
    else:
        grade = "none"
    return AutoMtoCompareResult(
        match_kind=kind,  # type: ignore[arg-type]
        members=tuple(members),
        rd_rows=int(entry.get("rd_rows") or 0),
        auto_rows=int(entry.get("auto_rows") or 0),
        combinations_checked=int(entry.get("combinations_checked") or 0),
        error=str(entry.get("error") or ""),
        grade=grade,
    )


def cache_path(runtime_dir: str | Path) -> Path:
    """Return ``runtime_dir / auto_mto_compare_cache.json``.

    Args:
        runtime_dir: Catalog runtime directory (never UNC source).

    Returns:
        JSON path under ``runtime_dir``.
    """

    return Path(runtime_dir) / AUTO_MTO_COMPARE_CACHE_NAME


def load_auto_mto_compare_cache(
    runtime_dir: str | Path | None,
) -> dict[str, dict[str, Any]]:
    """Load persistable AutoMTO compare entries from the runtime JSON.

    Args:
        runtime_dir: Catalog runtime directory. ``None`` or a missing /
            corrupt file yields an empty map.

    Returns:
        ``cache_key`` → entry. Entries with a non-persistable kind are
        dropped. A cache-format version mismatch returns ``{}``.
    """

    if runtime_dir is None:
        return {}
    path = cache_path(runtime_dir)
    if not path.is_file():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return {}
    if not isinstance(raw, dict):
        return {}
    if int(raw.get("version") or 0) != AUTO_MTO_COMPARE_CACHE_VERSION:
        return {}
    entries = raw.get("entries")
    if not isinstance(entries, dict):
        return {}
    cleaned: dict[str, dict[str, Any]] = {}
    for key, entry in entries.items():
        if not isinstance(entry, dict):
            continue
        if str(entry.get("match_kind") or "") not in _PERSISTABLE_KINDS:
            continue
        cleaned[str(key)] = dict(entry)
    return cleaned


def save_auto_mto_compare_cache(
    runtime_dir: str | Path | None,
    entries: Mapping[str, Mapping[str, Any]],
) -> None:
    """Write persistable AutoMTO compare entries to the runtime JSON.

    Args:
        runtime_dir: Catalog runtime directory. ``None`` is a no-op.
        entries: ``cache_key`` → entry. ``not_compared`` and other
            non-persistable kinds are not written.

    Raises:
        OSError: When the runtime directory or file cannot be written.
    """

    if runtime_dir is None:
        return
    cleaned: dict[str, dict[str, Any]] = {}
    for key, entry in entries.items():
        if not isinstance(entry, Mapping):
            continue
        if str(entry.get("match_kind") or "") not in _PERSISTABLE_KINDS:
            continue
        cleaned[str(key)] = dict(entry)
    path = cache_path(runtime_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": AUTO_MTO_COMPARE_CACHE_VERSION,
        "algorithm_version": AUTO_MTO_COMPARE_ALGORITHM_VERSION,
        "entries": cleaned,
    }
    text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    tmp_path = path.with_name(path.name + ".tmp")
    tmp_path.write_text(text, encoding="utf-8")
    tmp_path.replace(path)
