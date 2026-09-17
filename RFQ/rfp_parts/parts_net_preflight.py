"""Freshness check for ``rfp_parts_net.xlsx`` vs the single parts folder.

Launch uses this instead of the DS increase/decrease checklist. If the parts
folder has a new or newer workbook than the latest net, the same collection as
``python -m RFQ.rfp_parts`` is run again.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from RFQ.ds_compare.ds_units_normalize import normalize_units_text
from RFQ.units_convert import ALGORITHM_VERSION, GoogleUnitsIndex, UnitsConversionError
from RFQ.units_convert.matrix import (
    MatrixDocument,
    load_matrix,
    lookup_coefficient,
    matrix_row_for_code,
    resolved_conversion_target,
)
from RFQ.units_convert.models import normalize_code

SOURCES_JSON_NAME = "rfp_parts_sources.json"
BUILD_DEPS_JSON_NAME = "rfp_parts_build_deps.json"
_INVALID_COEF_SENTINEL = "?"


@dataclass(frozen=True)
class PartsFileStamp:
    """One workbook in ``RFP_Зиновьев``."""

    name: str
    mtime_ns: int
    size: int


@dataclass
class PartsNetFreshness:
    """Whether the latest parts net is still valid for Step1."""

    needs_rebuild: bool
    reason: str
    net_path: Path | None
    parts_dir: Path
    parts: list[Path] = field(default_factory=list)
    newer_than_net: list[Path] = field(default_factory=list)
    added_names: list[str] = field(default_factory=list)
    removed_names: list[str] = field(default_factory=list)


def write_parts_sources_snapshot(
    out_dir: Path,
    parts_dir: Path,
    files: list[Path],
) -> Path:
    """Write the parts-folder snapshot next to ``rfp_parts_net.xlsx``.

    Args:
        out_dir: Stamp folder of this collection run.
        parts_dir: Folder that was summed.
        files: Workbooks included in the sum.

    Returns:
        Path to the written JSON file.
    """
    payload: dict[str, Any] = {
        "parts_dir": str(parts_dir),
        "files": [
            {
                "name": path.name,
                "mtime_ns": path.stat().st_mtime_ns,
                "size": path.stat().st_size,
            }
            for path in files
        ],
    }
    out_path = Path(out_dir) / SOURCES_JSON_NAME
    out_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return out_path


def _load_sources_snapshot(run_dir: Path | None) -> list[str]:
    if run_dir is None:
        return []
    path = run_dir / SOURCES_JSON_NAME
    if not path.is_file():
        return []
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    files = raw.get("files") if isinstance(raw, dict) else None
    if not isinstance(files, list):
        return []
    names: list[str] = []
    for item in files:
        if isinstance(item, dict) and item.get("name"):
            names.append(str(item["name"]))
    return names


def _load_build_deps_snapshot(run_dir: Path | None) -> dict[str, Any] | None:
    if run_dir is None:
        return None
    path = run_dir / BUILD_DEPS_JSON_NAME
    if not path.is_file():
        return None
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    return raw if isinstance(raw, dict) else None


def _resolve_google_index(
    google_index: GoogleUnitsIndex | None,
) -> GoogleUnitsIndex:
    from RFQ.rfp_parts.analyze_rfp_parts import _load_google_rows_strict
    from RFQ.units_convert import build_google_units_index

    if google_index is not None:
        return google_index
    rows = _load_google_rows_strict()
    return build_google_units_index(rows)


def _resolve_matrix_path(units_matrix_path: Path | None) -> Path:
    from RFQ.rfp_parts.analyze_rfp_parts import _resolve_units_matrix_path

    return _resolve_units_matrix_path(units_matrix_path)


def _stored_dependency_items(stored: dict[str, Any]) -> list[dict[str, Any]]:
    raw = stored.get("dependencies")
    if not isinstance(raw, list):
        return []
    return [item for item in raw if isinstance(item, dict)]


def _fast_stored_deps_reason(
    stored: dict[str, Any] | None,
    *,
    matrix_path: Path,
) -> str | None:
    """Cheap checks on stored snapshot and configured matrix path."""
    if not stored:
        return "нет rfp_parts_build_deps.json — нужен сбор частей"
    if not _stored_dependency_items(stored):
        return "нет rfp_parts_build_deps.json — нужен сбор частей"

    if stored.get("ALGORITHM_VERSION") != ALGORITHM_VERSION:
        return (
            "изменилась версия алгоритма конвертации ед. изм.: "
            f"{stored.get('ALGORITHM_VERSION')!r} → {ALGORITHM_VERSION!r}"
        )

    from RFQ.rfp_parts.analyze_rfp_parts import AGGREGATION_KEY

    if stored.get("AGGREGATION_KEY") != AGGREGATION_KEY:
        return (
            "изменился ключ суммирования частей: "
            f"{stored.get('AGGREGATION_KEY')!r} → {AGGREGATION_KEY!r}"
        )

    current_matrix = str(matrix_path.resolve())
    stored_matrix = str(stored.get("matrix_path") or "")
    if stored_matrix != current_matrix:
        return "изменился путь матрицы конвертации ед. изм."

    if not matrix_path.is_file():
        return f"матрица конвертации недоступна: {matrix_path}"

    return None


def _current_coefficient_for_dependency(
    *,
    code_norm: str,
    source_norm: str,
    target_norm: str,
    matrix_doc: MatrixDocument,
) -> str:
    if source_norm == target_norm:
        return "1"
    pair = lookup_coefficient(
        matrix_doc,
        code_normalized=code_norm,
        source_normalized=source_norm,
    )
    if pair is None or pair.is_placeholder or pair.coefficient is None:
        return _INVALID_COEF_SENTINEL
    if pair.coefficient <= 0:
        return _INVALID_COEF_SENTINEL
    return str(pair.coefficient)


def _recompute_current_deps_from_stored(
    stored: dict[str, Any],
    *,
    google_index: GoogleUnitsIndex,
    matrix_doc: MatrixDocument,
    matrix_path: Path,
) -> dict[str, Any]:
    """Recompute dependency tuples for stored codes only (read-only matrix/Google)."""
    stored_items = _stored_dependency_items(stored)
    relevant_codes = sorted(
        {
            normalize_code(item.get("code"))
            for item in stored_items
            if normalize_code(item.get("code"))
        }
    )

    dependencies: list[dict[str, str]] = []
    for item in stored_items:
        code_norm = normalize_code(item.get("code"))
        source_norm = normalize_units_text(item.get("source_unit"))
        if not code_norm or not source_norm:
            continue
        google_norm, _display = resolved_conversion_target(
            code_norm,
            google_index=google_index,
            matrix_row=matrix_row_for_code(matrix_doc, code_norm),
        )
        target_norm = google_norm if google_norm else source_norm
        dependencies.append(
            {
                "code": code_norm,
                "source_unit": source_norm,
                "target_unit": target_norm,
                "coefficient": _current_coefficient_for_dependency(
                    code_norm=code_norm,
                    source_norm=source_norm,
                    target_norm=target_norm,
                    matrix_doc=matrix_doc,
                ),
            }
        )

    dependencies.sort(
        key=lambda row: (
            row["code"],
            row["source_unit"],
            row["target_unit"],
            row["coefficient"],
        )
    )
    return {
        "ALGORITHM_VERSION": ALGORITHM_VERSION,
        "matrix_path": str(matrix_path.resolve()),
        "dependencies": dependencies,
        "google_units_by_code": {
            code: google_index.google_unit(code) for code in relevant_codes
        },
    }


def _compare_stored_and_current_deps(
    stored: dict[str, Any],
    current: dict[str, Any],
) -> str | None:
    if stored.get("dependencies") != current.get("dependencies"):
        return "изменились зависимости конвертации ед. изм. (матрица/Google)"

    stored_google = stored.get("google_units_by_code")
    current_google = current.get("google_units_by_code")
    relevant_codes = {
        normalize_code(item.get("code"))
        for item in _stored_dependency_items(stored)
        if normalize_code(item.get("code"))
    }
    if isinstance(stored_google, dict) and isinstance(current_google, dict):
        for code in sorted(relevant_codes):
            if stored_google.get(code) != current_google.get(code):
                return (
                    "изменилось состояние Google UNITS для релевантного кода "
                    f"{code!r}"
                )
    elif stored_google != current_google:
        return "изменилось состояние Google UNITS для релевантных кодов"

    return None


def _assess_stored_build_deps(
    stored: dict[str, Any] | None,
    *,
    matrix_path: Path,
    google_index: GoogleUnitsIndex | None,
) -> str | None:
    fast_reason = _fast_stored_deps_reason(stored, matrix_path=matrix_path)
    if fast_reason:
        return fast_reason
    assert stored is not None

    try:
        matrix_doc = load_matrix(matrix_path)
        index = _resolve_google_index(google_index)
        current = _recompute_current_deps_from_stored(
            stored,
            google_index=index,
            matrix_doc=matrix_doc,
            matrix_path=matrix_path,
        )
    except (UnitsConversionError, OSError) as exc:
        return f"конвертация ед. изм. недоступна: {exc}"

    return _compare_stored_and_current_deps(stored, current)


def assess_parts_net_freshness(
    *,
    parts_dir: Path | None = None,
    reports_base: Path | None = None,
    units_matrix_path: Path | None = None,
    google_index: GoogleUnitsIndex | None = None,
    require_no_tags: bool = False,
) -> PartsNetFreshness:
    """Compare ``RFP_Зиновьев`` workbooks to the latest ``rfp_parts_net.xlsx``.

    Rebuild when the net is missing, a parts file is newer than the net, the
    set of file names changed vs the last snapshot, the aggregation key in
    stored deps no longer matches, or stored conversion dependencies no longer
    match the current read-only matrix/Google state.

    Args:
        parts_dir: Parts folder. Defaults to ``DEFAULT_PARTS_DIR``.
        reports_base: ``RFP сводный файл`` parent. Defaults to the UNC base.
        units_matrix_path: Optional conversion matrix override for dependency
            checks. Defaults to the main RFP profile path.
        google_index: Optional injectable Google units index for offline tests.
            When omitted, strict Google loading happens only after source files
            and stored snapshot fast-checks pass.
        require_no_tags: When True, a current tagged net still rebuilds if
            sibling ``rfp_parts_net_no_tags.xlsx`` is missing (``load_tags=false``).

    Returns:
        Freshness decision and the lists that explain it.

    Raises:
        FileNotFoundError: Parts folder is missing or contains no workbooks.
        NotADirectoryError: ``parts_dir`` exists but is not a directory.
    """
    from RFQ.rfp_parts.analyze_rfp_parts import (
        DEFAULT_PARTS_DIR,
        DEFAULT_REPORTS_BASE_DIR,
        NET_NO_TAGS_XLSX_NAME,
        _xlsx_files,
        find_latest_rfp_parts_run_dir,
        resolve_latest_rfp_parts_net_xlsx,
    )

    folder = Path(parts_dir or DEFAULT_PARTS_DIR)
    reports = Path(reports_base or DEFAULT_REPORTS_BASE_DIR)
    parts = _xlsx_files(folder)
    if not parts:
        raise FileNotFoundError(
            f"В папке частей нет xlsx: {folder}. "
            "Нечего суммировать в rfp_parts_net.xlsx."
        )

    net_path = resolve_latest_rfp_parts_net_xlsx(base=reports)
    run_dir = find_latest_rfp_parts_run_dir(base=reports)
    current_names = {path.name for path in parts}
    previous_names = set(_load_sources_snapshot(run_dir))
    added = sorted(current_names - previous_names) if previous_names else []
    removed = sorted(previous_names - current_names) if previous_names else []

    if net_path is None:
        return PartsNetFreshness(
            True,
            "нет rfp_parts_net.xlsx — нужен сбор частей",
            None,
            folder,
            parts,
            added_names=added or sorted(current_names),
            removed_names=removed,
        )

    net_mtime = net_path.stat().st_mtime
    newer = [path for path in parts if path.stat().st_mtime > net_mtime]
    if newer:
        sample = ", ".join(path.name for path in newer[:3])
        extra = f" (+{len(newer) - 3})" if len(newer) > 3 else ""
        return PartsNetFreshness(
            True,
            f"{len(newer)} файл(ов) в папке частей новее свода: {sample}{extra}",
            net_path,
            folder,
            parts,
            newer,
            added,
            removed,
        )
    if added:
        sample = ", ".join(added[:3])
        extra = f" (+{len(added) - 3})" if len(added) > 3 else ""
        return PartsNetFreshness(
            True,
            f"новые файлы в папке частей: {sample}{extra}",
            net_path,
            folder,
            parts,
            newer,
            added,
            removed,
        )
    if removed:
        sample = ", ".join(removed[:3])
        extra = f" (+{len(removed) - 3})" if len(removed) > 3 else ""
        return PartsNetFreshness(
            True,
            f"файлы исчезли из папки частей: {sample}{extra}",
            net_path,
            folder,
            parts,
            newer,
            added,
            removed,
        )

    stored = _load_build_deps_snapshot(run_dir)
    try:
        matrix_path = _resolve_matrix_path(units_matrix_path)
    except UnitsConversionError as exc:
        deps_reason = f"конвертация ед. изм. недоступна: {exc}"
    else:
        deps_reason = _assess_stored_build_deps(
            stored,
            matrix_path=matrix_path,
            google_index=google_index,
        )

    if deps_reason:
        return PartsNetFreshness(
            True,
            deps_reason,
            net_path,
            folder,
            parts,
            newer,
            added,
            removed,
        )

    if require_no_tags:
        no_tags_path = net_path.parent / NET_NO_TAGS_XLSX_NAME
        if not no_tags_path.is_file():
            return PartsNetFreshness(
                True,
                "нет rfp_parts_net_no_tags.xlsx — нужен сбор частей",
                net_path,
                folder,
                parts,
                newer,
                added,
                removed,
            )

    return PartsNetFreshness(
        False,
        "свод частей актуален",
        net_path,
        folder,
        parts,
        newer,
        added,
        removed,
    )


def _selected_net_path(tagged_net: Path, *, load_tags: bool) -> Path:
    """Tagged net, or sibling ``rfp_parts_net_no_tags.xlsx`` when tags are off."""
    from RFQ.rfp_parts.analyze_rfp_parts import NET_NO_TAGS_XLSX_NAME

    if load_tags:
        return tagged_net
    return tagged_net.parent / NET_NO_TAGS_XLSX_NAME


def ensure_rfp_parts_net_current(
    *,
    parts_dir: Path | None = None,
    reports_base: Path | None = None,
    progress: bool = False,
    units_matrix_path: Path | None = None,
    google_index: GoogleUnitsIndex | None = None,
    load_tags: bool = True,
) -> tuple[Path, PartsNetFreshness]:
    """Return a fresh parts net, rebuilding when the parts folder changed.

    Rebuild runs the same collection as the «RFP · Сбор частей» tab, without
    the DS checklist compare. Collection always writes both the tagged net and
    ``rfp_parts_net_no_tags.xlsx``. When ``load_tags`` is False the returned
    path is the no-tags sibling (missing file forces rebuild). The tagged net
    in an older stamp is not overwritten in place.

    Args:
        parts_dir: Parts folder. Defaults to ``DEFAULT_PARTS_DIR``.
        reports_base: Reports parent. Defaults to ``DEFAULT_REPORTS_BASE_DIR``.
        progress: When True, print short status lines.
        units_matrix_path: Optional conversion matrix override forwarded to
            ``run_rfp_parts_analyze`` on rebuild.
        google_index: Optional injectable Google units index for offline tests.
        load_tags: True returns ``rfp_parts_net.xlsx``. False requires and
            returns ``rfp_parts_net_no_tags.xlsx``.

    Returns:
        Path to the net workbook for this mode and the freshness decision used.

    Raises:
        FileNotFoundError: Parts folder empty/missing, or rebuild did not write
            the expected net file.
    """
    from RFQ.rfp_parts.analyze_rfp_parts import (
        DEFAULT_PARTS_DIR,
        DEFAULT_REPORTS_BASE_DIR,
        NET_NO_TAGS_XLSX_NAME,
        NET_XLSX_NAME,
        make_reports_out_dir,
        run_rfp_parts_analyze,
    )

    folder = Path(parts_dir or DEFAULT_PARTS_DIR)
    reports = Path(reports_base or DEFAULT_REPORTS_BASE_DIR)
    freshness = assess_parts_net_freshness(
        parts_dir=folder,
        reports_base=reports,
        units_matrix_path=units_matrix_path,
        google_index=google_index,
        require_no_tags=not load_tags,
    )
    if progress:
        print(f"[parts] {freshness.reason}", flush=True)
    if not freshness.needs_rebuild:
        assert freshness.net_path is not None
        selected = _selected_net_path(freshness.net_path, load_tags=load_tags)
        if not selected.is_file():
            raise FileNotFoundError(
                f"нет ожидаемого свода частей: {selected}"
            )
        return selected, freshness

    out_dir = make_reports_out_dir(base=reports)
    if progress:
        print(f"[parts] пересборка свода → {out_dir}", flush=True)
    net_path = run_rfp_parts_analyze(
        parts_dir=folder,
        out_dir=out_dir,
        no_checklist=True,
        units_matrix_path=units_matrix_path,
    )
    tagged = Path(net_path)
    if not tagged.is_file():
        raise FileNotFoundError(
            f"сбор частей не создал {NET_XLSX_NAME}: {out_dir}"
        )
    selected = _selected_net_path(tagged, load_tags=load_tags)
    if not selected.is_file():
        missing = NET_XLSX_NAME if load_tags else NET_NO_TAGS_XLSX_NAME
        raise FileNotFoundError(
            f"сбор частей не создал {missing}: {out_dir}"
        )
    return selected, freshness
