"""Preflight refresh and source manifest for grouped DS/MTO/packing compare."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import utils.path
from base.base_classes import RowType
from base.tables_columns import VALUES
from RFQ.ds_compare.ds_compare_config import (
    normalize_gui_paths,
    save_ds_compare_config,
)
from RFQ.ds_compare.ds_quantity_parse import try_parse_quantity
from RFQ.ds_compare.ds_merge_folder import merge_ds_folder
from RFQ.ds_compare.load_mto import refresh_stale_mto_caches
from RFQ.ds_compare.support_finctions import load_ds_data
from RFQ.ds_compare.tsd_packing_load import (
    collect_tsd_files,
    load_and_cache_tsd_packing,
)

_MTO_NOT_FOUND = "Файл МТО не найден"


@dataclass(frozen=True)
class SourceFileRecord:
    """One file listed on a source worksheet."""

    key: str
    role: str
    path: str
    modified_at: str
    status: str
    note: str = ""


@dataclass
class SourceManifest:
    """Files and cache metadata used to build the final workbook."""

    ds: list[SourceFileRecord] = field(default_factory=list)
    mto: list[SourceFileRecord] = field(default_factory=list)
    packing: list[SourceFileRecord] = field(default_factory=list)
    packing_root: str = ""
    packing_fingerprint: str = ""
    packing_cache_path: str = ""
    packing_summary_path: str = ""


@dataclass
class PreflightResult:
    """Preflight output consumed by the grouped orchestrator."""

    effective_ds_path: str
    ds_refreshed: bool
    packing_refreshed: bool
    manifest: SourceManifest


def _modified_at(path: str) -> str:
    """Return local ISO mtime or an empty string for a missing file."""
    try:
        return datetime.fromtimestamp(os.path.getmtime(path)).isoformat(
            sep=" ", timespec="seconds"
        )
    except OSError:
        return ""


def _collect_ds_files(root: str) -> list[Any]:
    if not root or not os.path.isdir(root):
        return []
    return utils.path.get_files_single(
        root,
        endswith=(".xlsx", ".XLSX"),
        forbidden_endswith=(),
        sub_folders=False,
    )


def _ds_manifest(
    source_docs: list[Any],
    summary_path: str,
) -> list[SourceFileRecord]:
    records = [
        SourceFileRecord(
            key=str(index),
            role="исходный",
            path=str(doc.file_full_path),
            modified_at=_modified_at(str(doc.file_full_path)),
            status="в merge",
        )
        for index, doc in enumerate(source_docs, 1)
    ]
    records.insert(
        0,
        SourceFileRecord(
            key="свод",
            role="свод",
            path=summary_path,
            modified_at=_modified_at(summary_path),
            status="использован" if os.path.isfile(summary_path) else "отсутствует",
        ),
    )
    return records


def _summary_has_invalid_position_quantities(summary_path: str) -> bool:
    """Return True when a merged DS contains nonnumeric position quantities."""
    if not os.path.isfile(summary_path):
        return False
    try:
        rows = load_ds_data(
            summary_path,
            summ_ds=True,
            force_update=False,
            use_cache=True,
        )
    except Exception as exc:
        print(f"Preflight ДС WARNING: не удалось проверить свод ({exc}).")
        return True
    invalid_count = 0
    for row in rows:
        if row.row_type != RowType.position_row:
            continue
        ok, _ = try_parse_quantity(row.get_value(VALUES))
        if not ok:
            invalid_count += 1
    if invalid_count:
        print(
            "Preflight ДС: в текущем своде найдены некорректные количества "
            f"position rows={invalid_count}; свод будет пересобран."
        )
    return invalid_count > 0


def _packing_manifest(root: str) -> list[SourceFileRecord]:
    if not root or not os.path.isdir(root):
        return []
    records: list[SourceFileRecord] = []
    for index, doc in enumerate(collect_tsd_files(root), 1):
        path = str(doc.file_full_path)
        try:
            relative = str(Path(path).resolve().relative_to(Path(root).resolve()))
        except (OSError, ValueError):
            relative = path
        records.append(
            SourceFileRecord(
                key=str(index),
                role="исходный",
                path=path,
                modified_at=_modified_at(path),
                status="использован",
                note=relative,
            )
        )
    return records


def refresh_sources_before_grouped(
    *,
    ds_path: str,
    rfq_only: bool,
    cfg: dict[str, Any],
) -> PreflightResult:
    """Refresh DS summary and packing cache before grouped comparison.

    Args:
        ds_path: Currently selected DS summary workbook.
        rfq_only: Whether this is an RFQ-only cached compare.
        cfg: Loaded DS compare configuration; updated paths are saved here.

    Returns:
        Effective DS path, refresh flags, and initial source manifest.

    Raises:
        RuntimeError: A stale DS folder cannot be merged or packing refresh fails.
    """
    gui_paths = normalize_gui_paths(cfg.get("gui_paths"))
    effective_ds_path = str(ds_path)
    ds_refreshed = False
    merge_root = str(gui_paths.get("last_merge_ds_folder", "")).strip()
    source_docs = _collect_ds_files(merge_root)

    if not rfq_only and source_docs:
        newest_source_mtime = max(
            os.path.getmtime(str(doc.file_full_path)) for doc in source_docs
        )
        try:
            summary_mtime = os.path.getmtime(effective_ds_path)
        except OSError:
            summary_mtime = -1.0
        invalid_summary = _summary_has_invalid_position_quantities(effective_ds_path)
        if newest_source_mtime > summary_mtime or invalid_summary:
            print(
                "Preflight ДС: требуется актуализация свода — выполняется объединение."
            )
            merged_path = merge_ds_folder(
                merge_root,
                include_subfolders=False,
                out_file_prefix="DS_summary",
                open_folder=False,
                dump_include_empty_row=False,
            )
            if not merged_path:
                raise RuntimeError(
                    "Preflight ДС: новый свод не создан. Проверьте сообщения "
                    "«ОШИБКА ФОРМАТА» и DS_merge_load_report.txt."
                )
            effective_ds_path = str(merged_path)
            gui_paths["last_ds_file"] = effective_ds_path
            ds_refreshed = True
    elif not rfq_only and merge_root and not os.path.isdir(merge_root):
        print(
            "Preflight ДС WARNING: папка исходных ДС недоступна; "
            "используется выбранный свод."
        )

    packing_root = str(gui_paths.get("last_tsd_packing_folder", "")).strip()
    packing_refreshed = False
    packing_summary_path = ""
    if packing_root and os.path.isdir(packing_root):
        print("Preflight УЛ: проверка fingerprint исходной папки.")
        try:
            packing_result = load_and_cache_tsd_packing(packing_root, force=False)
        except Exception as exc:
            raise RuntimeError(f"Preflight УЛ: не удалось обновить данные ({exc})") from exc
        packing_summary_path = str(packing_result.summary_path)
        gui_paths["last_tsd_summary_file"] = packing_summary_path
        packing_refreshed = True
    else:
        print(
            "Preflight УЛ WARNING: папка упаковочных листов недоступна; "
            "будет проверен существующий кэш."
        )

    cfg["gui_paths"] = gui_paths
    if (ds_refreshed or packing_summary_path) and not save_ds_compare_config(cfg):
        raise RuntimeError("Preflight: не удалось сохранить обновлённые пути в конфиге.")

    manifest = SourceManifest(
        ds=_ds_manifest(source_docs, effective_ds_path),
        packing=_packing_manifest(packing_root),
        packing_root=packing_root,
        packing_summary_path=packing_summary_path,
    )
    return PreflightResult(
        effective_ds_path=effective_ds_path,
        ds_refreshed=ds_refreshed,
        packing_refreshed=packing_refreshed,
        manifest=manifest,
    )


def refresh_mto_sources(
    spec_dict: dict[str, object],
    manifest: SourceManifest,
    *,
    rfq_only: bool,
) -> int:
    """Refresh stale relevant MTO pickles and populate the MTO manifest."""
    refreshed = 0 if rfq_only else refresh_stale_mto_caches(spec_dict)
    records: list[SourceFileRecord] = []
    for spec_key, raw_path in sorted(spec_dict.items(), key=lambda item: str(item[0])):
        path = str(raw_path or "")
        found = path != _MTO_NOT_FOUND and os.path.isfile(path)
        records.append(
            SourceFileRecord(
                key=str(spec_key),
                role="спецификация",
                path=path,
                modified_at=_modified_at(path) if found else "",
                status="найден" if found else _MTO_NOT_FOUND,
            )
        )
    manifest.mto = records
    return refreshed


def finalize_packing_manifest(
    manifest: SourceManifest,
    dataset: Any,
) -> None:
    """Attach the exact loaded packing cache metadata to the manifest."""
    meta = getattr(dataset, "meta", None)
    if meta is not None:
        manifest.packing_root = str(getattr(meta, "root", "") or manifest.packing_root)
        manifest.packing_fingerprint = str(getattr(meta, "fingerprint", "") or "")
        reports = getattr(meta, "report_paths", {}) or {}
        if not manifest.packing_summary_path:
            manifest.packing_summary_path = str(reports.get("summary", "") or "")
    manifest.packing_cache_path = str(getattr(dataset, "cache_path", "") or "")
    if not manifest.packing and manifest.packing_root:
        manifest.packing = _packing_manifest(manifest.packing_root)
