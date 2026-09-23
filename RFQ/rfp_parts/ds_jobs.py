"""FunctionJobRunner jobs for the RFP · Сбор частей DS/hybrid cockpit.

«Проверить реестр» keeps one working file in ``_RFP``. A legacy workbook is
renamed to ``Реестр_ДС_УЛ_old.xlsx`` and replaced. A file open in Excel stays
put; a dated copy is written instead, and only five dated copies are kept.
``backup_and_replace_registry`` is never imported or called from this module.
"""

from __future__ import annotations

import os
import re
import sys
import time
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Literal, Sequence

from RFQ.ds_compare.ds_compare_config import (
    load_ds_compare_config,
    normalize_gui_paths,
)
from RFQ.rfp_parts.analyze_rfp_parts import (
    DEFAULT_PARTS_DIR,
    DEFAULT_REPORTS_BASE_DIR,
    make_reports_out_dir,
)
from RFQ.rfp_parts.ds_baseline import (
    BASELINE_XLSX_NAME,
    ISSUE_EMPTY_CODE,
    ISSUE_EXACT_DUPLICATE,
    ISSUE_QTY_EMPTY,
    ISSUE_QTY_FORMULA,
    ISSUE_QTY_NEGATIVE,
    ISSUE_QTY_NON_FINITE,
    ISSUE_QTY_NON_NUMERIC,
    ISSUE_QTY_ZERO,
    ISSUE_TAG_DUPLICATE,
    ISSUE_TAG_MISMATCH,
    ISSUE_UNKNOWN_GOOGLE,
    DsBaselineResult,
    build_ds_baseline,
    collect_ds_workbooks,
    resolve_ds_source_id,
)
from RFQ.rfp_parts.ds_progress import DsFileProgressTracker, DsProgressSession
from RFQ.rfp_parts.ds_registry import (
    CANONICAL_REGISTRY_NAME,
    DEFAULT_REGISTRY_PATH,
    DEFAULT_RFP_BASE,
    FORMAT_VERSION,
    LEGEND_SHEET_NAME,
    MIGRATION_REPORT_PREFIX,
    OLD_REGISTRY_NAME,
    MODE_NO_UL,
    DsRegistryDocument,
    DsRegistryError,
    _ds_file_names,
    detect_registry_format,
    default_migration_report_path,
    ensure_registry_legend,
    RFP_RECONCILE_NO_FILE,
    RegistryLinks,
    index_rfp_files,
    scan_ul_catalog,
    install_working_registry,
    load_registry,
    migrate_legacy_rows,
    migrate_registry,
    read_legacy_registry,
    resolve_latest_registry,
)
from RFQ.rfp_parts.ds_rfp_hybrid import (
    HYBRID_XLSX_NAME,
    STATUS_MATCH,
    DsRfpHybridResult,
    build_ds_rfp_hybrid,
    collect_rfp_workbooks,
    rfp_identity_key,
)

Tone = Literal["error", "warn", "match", "ok"]
JobKind = Literal["registry", "baseline", "hybrid", "coverage"]

DS_REGISTRY_DIR_NAME = "_ds_registry"
ROBOT_REGISTRY_COPY_NAME = "Реестр_ДС_УЛ_migrated.xlsx"

_QTY_ISSUE_CODES = frozenset(
    {
        ISSUE_QTY_EMPTY,
        ISSUE_QTY_FORMULA,
        ISSUE_QTY_NEGATIVE,
        ISSUE_QTY_NON_FINITE,
        ISSUE_QTY_NON_NUMERIC,
        ISSUE_QTY_ZERO,
    }
)
_TAG_ISSUE_CODES = frozenset({ISSUE_TAG_DUPLICATE, ISSUE_TAG_MISMATCH})


@dataclass(frozen=True, slots=True)
class DsJobResult:
    """Duck-typed ``FunctionJobRunner`` result (no GUI import)."""

    success: bool
    message: str
    result_path: str | None = None


@dataclass(frozen=True, slots=True)
class CockpitRow:
    """One table row with a traffic-light tone for the GUI."""

    cells: tuple[str, ...]
    tone: Tone = "ok"


@dataclass
class DsCockpitSnapshot:
    """Last DS/hybrid job outcome for the RFP · Сбор частей tab."""

    kind: JobKind
    summary: str
    registry_path: Path
    registry_format: str = "unknown"
    format_version: int = FORMAT_VERSION
    output_dir: Path | None = None
    rfp_root: Path | None = None
    source_root: Path | None = None
    ul_root: Path | None = None
    baseline_path: Path | None = None
    hybrid_path: Path | None = None
    report_path: Path | None = None
    migrated_registry_path: Path | None = None
    next_step: str = ""
    active_count: int = 0
    group_count: int = 0
    error_count: int = 0
    warn_count: int = 0
    overlay_count: int = 0
    match_count: int = 0
    mismatch_count: int = 0
    ds_only_count: int = 0
    rfp_only_count: int = 0
    blocked_count: int = 0
    file_count: int = 0
    rfp_file_count: int = 0
    empty_code: int = 0
    qty_errors: int = 0
    unknown_google: int = 0
    tag_warnings: int = 0
    duplicates: int = 0
    registry_rows: list[CockpitRow] = field(default_factory=list)
    group_rows: list[CockpitRow] = field(default_factory=list)
    file_rows: list[CockpitRow] = field(default_factory=list)
    coverage_rows: list[CockpitRow] = field(default_factory=list)

    @property
    def banner_tone(self) -> Tone:
        if self.error_count or self.blocked_count:
            return "error"
        if self.warn_count or self.overlay_count or self.mismatch_count:
            return "warn"
        if self.match_count:
            return "match"
        return "ok"


_last_cockpit: DsCockpitSnapshot | None = None


def get_last_ds_cockpit() -> DsCockpitSnapshot | None:
    """Return the last cockpit snapshot from a DS job, if any."""
    return _last_cockpit


def _emit(message: str) -> None:
    text = message + "\n"
    stream = sys.stdout
    try:
        stream.write(text)
    except UnicodeEncodeError:
        encoding = getattr(stream, "encoding", None) or "utf-8"
        payload = text.encode(encoding, errors="backslashreplace")
        buf = getattr(stream, "buffer", None)
        if buf is not None:
            buf.write(payload)
        else:
            stream.write(payload.decode(encoding, errors="replace"))
    try:
        stream.flush()
    except Exception:
        pass


def _baseline_progress_hooks(
    label: str = "ДС",
) -> tuple[
    DsProgressSession,
    DsFileProgressTracker,
    dict[str, object],
]:
    """Build kwargs for ``build_ds_baseline`` with live stdout progress."""
    session = DsProgressSession(_emit)
    tracker = DsFileProgressTracker(_emit, label=label)

    def on_discovered(count: int, skipped: int) -> None:
        tracker.bind_total(count)
        session.message(f"{label}: {count} файлов, пропущено {skipped}")

    hooks: dict[str, object] = {
        "progress_callback": tracker.baseline_callback(),
        "phase_callback": session.message,
        "files_discovered_callback": on_discovered,
    }
    return session, tracker, hooks


def _rfp_progress_hooks(label: str = "RFP") -> tuple[DsFileProgressTracker, dict[str, object]]:
    session = DsProgressSession(_emit)
    tracker = DsFileProgressTracker(_emit, label=label)

    def on_phase(msg: str) -> None:
        session.message(msg)
        if msg.startswith("разбор RFP:"):
            match = re.search(r"(\d+) файлов", msg)
            if match:
                tracker.bind_total(int(match.group(1)))

    return tracker, {
        "progress_callback": tracker.rfp_callback(),
        "phase_callback": on_phase,
    }


def _is_dir(path: Path) -> bool:
    try:
        return path.is_dir()
    except OSError:
        return False


def _is_file(path: Path) -> bool:
    try:
        return path.is_file()
    except OSError:
        return False


def _same_file(left: Path, right: Path) -> bool:
    try:
        return left.resolve() == right.resolve()
    except OSError:
        a = str(left).replace("/", "\\").casefold()
        b = str(right).replace("/", "\\").casefold()
        return a == b


def _gui_paths() -> dict[str, str]:
    cfg = load_ds_compare_config()
    return dict(normalize_gui_paths(cfg.get("gui_paths")))


def _resolve_source_root(source_root: str | Path | None) -> Path | None:
    if source_root:
        text = str(source_root).strip()
        return Path(text) if text else None
    text = _gui_paths().get("last_ds_trusted_folder", "").strip()
    return Path(text) if text else None


def _resolve_registry_path(registry_path: str | Path | None) -> Path:
    if registry_path:
        text = str(registry_path).strip()
        if text:
            return Path(text)
    text = _gui_paths().get("last_ds_registry_file", "").strip()
    return Path(text) if text else DEFAULT_REGISTRY_PATH


def _resolve_ul_root(ul_root: str | Path | None) -> Path | None:
    text = str(ul_root).strip() if ul_root else ""
    if not text:
        text = _gui_paths().get("last_tsd_packing_folder", "").strip()
    if not text:
        return None
    path = Path(text)
    return path if _is_dir(path) else path


def _resolve_rfp_root(rfp_root: str | Path | None) -> Path:
    if rfp_root:
        text = str(rfp_root).strip()
        if text:
            return Path(text)
    return DEFAULT_PARTS_DIR


def ds_registry_robot_dir(reports_base: str | Path | None = None) -> Path:
    """Stable folder for the robot-facing registry copy. Not the UNC canon."""
    from RFQ.rfp_parts.analyze_rfp_parts import DEFAULT_REPORTS_BASE_DIR

    return Path(reports_base or DEFAULT_REPORTS_BASE_DIR) / DS_REGISTRY_DIR_NAME


def _resolve_output_dir(output_dir: str | Path | None) -> Path:
    if output_dir:
        return Path(output_dir)
    return make_reports_out_dir()


def _ensure_output_dir(output_dir: Path) -> str | None:
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return f"не удалось создать папку отчётов: {exc}"
    return None


def _optional_ul_for_validate(ul_root: Path | None) -> Path | None:
    if ul_root is None:
        return None
    return ul_root if _is_dir(ul_root) else None


def _catalog_if_dir(path: Path | None) -> Path | None:
    if path is None:
        return None
    return path if _is_dir(path) else None


def _load_new_registry(
    path: Path,
    *,
    ul_root: Path | None,
    ds_root: Path | None = None,
    rfp_root: Path | None = None,
    on_lap: Callable[[str, float], None] | None = None,
) -> tuple[DsRegistryDocument | None, str | None]:
    try:
        document = load_registry(
            path,
            ul_root=_optional_ul_for_validate(ul_root),
            ds_root=_catalog_if_dir(ds_root),
            rfp_root=_catalog_if_dir(rfp_root),
            on_lap=on_lap,
        )
    except DsRegistryError as exc:
        return None, str(exc)
    except OSError as exc:
        return None, f"не удалось прочитать реестр: {exc}"
    return document, None


def _tone_from_levels(levels: set[str], *, match: bool = False) -> Tone:
    if "ERROR" in levels:
        return "error"
    if "WARN" in levels or "OVERLAY" in levels:
        return "warn"
    if match:
        return "match"
    return "ok"


def _issue_note(messages: list[str], *, limit: int = 2) -> str:
    if not messages:
        return ""
    head = messages[:limit]
    extra = len(messages) - len(head)
    text = "; ".join(head)
    if extra > 0:
        text = f"{text} (+ ещё {extra})"
    return text


def _registry_tables(
    document: DsRegistryDocument,
    *,
    ds_files_by_id: dict[str, list[str]] | None = None,
    rfp_files_by_key: dict[str, list[str]] | None = None,
    ul_root: Path | None = None,
    group_status: dict[str, str] | None = None,
) -> tuple[list[CockpitRow], list[CockpitRow], list[CockpitRow]]:
    ds_files_by_id = ds_files_by_id or {}
    rfp_files_by_key = rfp_files_by_key or {}
    group_status = group_status or {}
    issues_by_source: dict[str, list[Any]] = defaultdict(list)
    issues_by_group: dict[str, list[Any]] = defaultdict(list)
    for item in document.validation.issues:
        if item.source_id:
            issues_by_source[item.source_id].append(item)
        if item.group_id:
            issues_by_group[item.group_id].append(item)

    registry_rows: list[CockpitRow] = []
    for row in document.rows:
        if not row.source_id and not row.relations:
            continue
        rels = [rel for rel in row.relations if not rel.is_empty()]
        groups = "; ".join(rel.group_id for rel in rels if rel.group_id) or "—"
        keys = "; ".join(rel.rfp_key for rel in rels if rel.rfp_key) or "—"
        folders = "; ".join(rel.ul_folder for rel in rels if rel.ul_folder) or "—"
        modes = "; ".join(rel.mode for rel in rels if rel.mode) or "—"
        items = issues_by_source.get(row.source_id, [])
        levels = {str(item.level) for item in items}
        note = _issue_note([item.message for item in items])
        registry_rows.append(
            CockpitRow(
                cells=(
                    row.source_id or "—",
                    row.status or "—",
                    groups,
                    keys,
                    folders,
                    modes,
                    note,
                ),
                tone=_tone_from_levels(levels),
            )
        )

    grouped: dict[str, dict[str, Any]] = {}
    for row in document.active_rows:
        for rel in row.relations:
            if rel.is_empty() or not rel.group_id:
                continue
            bucket = grouped.setdefault(
                rel.group_id,
                {
                    "sources": [],
                    "rfp_key": rel.rfp_key,
                    "ul_folder": rel.ul_folder,
                    "mode": rel.mode,
                },
            )
            if row.source_id and row.source_id not in bucket["sources"]:
                bucket["sources"].append(row.source_id)

    group_rows: list[CockpitRow] = []
    coverage_rows: list[CockpitRow] = []
    for group_id, bucket in sorted(grouped.items(), key=lambda item: item[0]):
        sources: list[str] = bucket["sources"]
        rfp_key = str(bucket["rfp_key"] or "")
        ul_folder = str(bucket["ul_folder"] or "")
        mode = str(bucket["mode"] or "")
        ds_names = [
            name for source_id in sources for name in ds_files_by_id.get(source_id, [])
        ]
        rfp_names = rfp_files_by_key.get(rfp_key, [])
        ul_ok = True
        if mode == MODE_NO_UL:
            ul_label = "Нет УЛ"
        elif not ul_folder:
            ul_label = "—"
        elif ul_root is None:
            ul_label = ul_folder
        else:
            ul_ok = _is_dir(ul_root / ul_folder)
            ul_label = ul_folder if ul_ok else f"нет: {ul_folder}"
        status = group_status.get(group_id, "")
        missing_ds = [sid for sid in sources if sid not in ds_files_by_id]
        missing_rfp = bool(rfp_key) and not rfp_names
        notes: list[str] = []
        levels: set[str] = set()
        for item in issues_by_group.get(group_id, []):
            levels.add(str(item.level))
            notes.append(item.message)
        if missing_ds:
            levels.add("WARN")
            notes.append("нет файла ДС: " + ", ".join(missing_ds))
        if missing_rfp:
            levels.add("WARN")
            notes.append("нет файла RFP")
        if not ul_ok:
            levels.add("WARN")
            notes.append("нет папки УЛ")
        match = status == STATUS_MATCH
        if status:
            notes.insert(0, status)
        tone = _tone_from_levels(levels, match=match)
        if match and tone == "ok":
            tone = "match"
        group_rows.append(
            CockpitRow(
                cells=(
                    group_id,
                    "; ".join(sources) or "—",
                    rfp_key or "—",
                    ul_label or "—",
                    "да" if any(item.blocks_overlay for item in issues_by_group.get(group_id, [])) else "нет",
                    status or "—",
                ),
                tone=tone,
            )
        )
        coverage_rows.append(
            CockpitRow(
                cells=(
                    group_id,
                    ", ".join(ds_names) or ("нет ДС" if missing_ds else "—"),
                    ", ".join(rfp_names) or ("нет RFP" if missing_rfp else "—"),
                    ul_label or "—",
                    "; ".join(notes[:3]) or "OK",
                ),
                tone=tone,
            )
        )
    return registry_rows, group_rows, coverage_rows


def _file_rows_from_collects(
    *,
    ds_files: list[Any],
    ds_skipped: list[Any],
    rfp_files: list[Any],
    rfp_skipped: list[Any],
    ds_files_by_id: dict[str, list[str]],
    rfp_files_by_key: dict[str, list[str]],
) -> list[CockpitRow]:
    id_by_name = {
        name: source_id
        for source_id, names in ds_files_by_id.items()
        for name in names
    }
    key_by_name = {
        name: key for key, names in rfp_files_by_key.items() for name in names
    }
    rows: list[CockpitRow] = []
    for item in ds_files:
        name = item.path.name
        rows.append(
            CockpitRow(
                cells=("ДС", item.relpath or name, id_by_name.get(name, "—"), "OK"),
                tone="ok",
            )
        )
    for item in ds_skipped:
        rows.append(
            CockpitRow(
                cells=("ДС", item.relpath or item.path.name, "—", item.reason),
                tone="warn",
            )
        )
    for item in rfp_files:
        name = item.path.name
        key = key_by_name.get(name, "")
        tone: Tone = "ok" if key else "warn"
        rows.append(
            CockpitRow(
                cells=("RFP", name, key or "не разобран", "OK" if key else "нет ключа"),
                tone=tone,
            )
        )
    for item in rfp_skipped:
        rows.append(
            CockpitRow(
                cells=("RFP", item.relpath or item.path.name, "—", item.reason),
                tone="warn",
            )
        )
    return rows


def _count_baseline_quality(baseline: DsBaselineResult) -> dict[str, int]:
    empty_code = sum(1 for item in baseline.issues if item.code == ISSUE_EMPTY_CODE)
    qty_errors = sum(1 for item in baseline.issues if item.code in _QTY_ISSUE_CODES)
    unknown_google = sum(
        1 for item in baseline.issues if item.code == ISSUE_UNKNOWN_GOOGLE
    )
    tag_warnings = sum(1 for item in baseline.issues if item.code in _TAG_ISSUE_CODES)
    duplicates = sum(1 for item in baseline.issues if item.code == ISSUE_EXACT_DUPLICATE)
    return {
        "empty_code": empty_code,
        "qty_errors": qty_errors,
        "unknown_google": unknown_google,
        "tag_warnings": tag_warnings,
        "duplicates": duplicates,
    }


def _map_ds_files(
    files: list[Any], active_ids: list[str]
) -> dict[str, list[str]]:
    mapped: dict[str, list[str]] = defaultdict(list)
    for item in files:
        resolved = resolve_ds_source_id(item.relpath, active_ids)
        if resolved.source_id:
            mapped[resolved.source_id].append(item.path.name)
    return dict(mapped)


def _map_rfp_files(files: list[Any]) -> dict[str, list[str]]:
    mapped: dict[str, list[str]] = defaultdict(list)
    for item in files:
        key = rfp_identity_key(parse_rfp_ds_identity(item.path.name)) or ""
        if key:
            mapped[key].append(item.path.name)
    return dict(mapped)


def _overlay_blocking_labels(document: DsRegistryDocument) -> list[str]:
    labels: list[str] = []
    seen: set[str] = set()
    for item in document.validation.issues:
        if not item.blocks_overlay:
            continue
        label = (item.source_id or item.group_id or "").strip()
        if not label or label in seen:
            continue
        seen.add(label)
        labels.append(label)
    return labels


def _registry_next_step(
    document: DsRegistryDocument,
    *,
    migrated_path: Path | None,
    registry_format: str,
    placement: str = "",
    archived_now: bool = False,
) -> str:
    """Build the Russian next-step sentence for a registry-kind snapshot."""
    labels = _overlay_blocking_labels(document)
    labels_text = ", ".join(labels)
    path = str(migrated_path or document.path)
    if placement == "dated":
        return (
            f"{CANONICAL_REGISTRY_NAME} открыт в Excel, заменить его нельзя. "
            f"Записана копия {path}. Программа берёт самый новый файл в папке _RFP "
            f"и хранит {5} таких копий, более старые удаляет. Закройте Excel и "
            "проверьте реестр ещё раз — тогда основной файл будет заменён."
        )
    if placement == "canonical":
        renamed = (
            f" Старый формат переименован в {OLD_REGISTRY_NAME}."
            if archived_now
            else ""
        )
        blocked = ""
        if labels:
            blocked = (
                f" Свод не пишется, пока не разобраны номера {labels_text}: "
                "один номер RFP или один файл RFP стоят на разных ДС."
            )
        return (
            f"Рабочий реестр один, в папке _RFP: {path}.{renamed}{blocked} "
            "Дальше: «Собрать свод только из ДС» или «Наложить RFP на группы»."
        )
    if registry_format == "new":
        if labels:
            return (
                "Реестр уже нового формата, канон не перезаписывался. "
                f"Связи без фильтров: {labels_text}. Их группы останутся из ДС, "
                "пока в реестре не заполнены фильтр титула и фильтр марки. "
                "Остальное можно собирать кнопкой «Собрать свод только из ДС»."
            )
        return (
            "Реестр нового формата, замечаний нет. Дальше: "
            "«Собрать свод только из ДС» или «Наложить RFP на группы»."
        )
    return ""


def _snapshot_from_registry(
    *,
    kind: JobKind,
    document: DsRegistryDocument,
    registry_format: str,
    output_dir: Path | None = None,
    source_root: Path | None = None,
    rfp_root: Path | None = None,
    ul_root: Path | None = None,
    baseline: DsBaselineResult | None = None,
    hybrid: DsRfpHybridResult | None = None,
    migrated_registry_path: Path | None = None,
    extra_warn: int = 0,
    placement: str = "",
    archived_now: bool = False,
) -> DsCockpitSnapshot:
    ul_for_cov = _optional_ul_for_validate(ul_root)
    active_ids = [row.source_id for row in document.active_rows if row.source_id]
    ds_files: list[Any] = []
    ds_skipped: list[Any] = []
    rfp_files: list[Any] = []
    rfp_skipped: list[Any] = []
    if source_root is not None and _is_dir(source_root):
        ds_files, ds_skipped = collect_ds_workbooks(
            source_root, registry_path=document.path
        )
    if rfp_root is not None and _is_dir(rfp_root):
        rfp_files, rfp_skipped = collect_rfp_workbooks(rfp_root)
    ds_files_by_id = _map_ds_files(ds_files, active_ids)
    rfp_files_by_key = _map_rfp_files(rfp_files)
    group_status = {}
    if hybrid is not None:
        group_status = {item.group_id: item.status for item in hybrid.groups}
    registry_rows, group_rows, coverage_rows = _registry_tables(
        document,
        ds_files_by_id=ds_files_by_id,
        rfp_files_by_key=rfp_files_by_key,
        ul_root=ul_for_cov,
        group_status=group_status,
    )
    file_rows = _file_rows_from_collects(
        ds_files=ds_files,
        ds_skipped=ds_skipped,
        rfp_files=rfp_files,
        rfp_skipped=rfp_skipped,
        ds_files_by_id=ds_files_by_id,
        rfp_files_by_key=rfp_files_by_key,
    )
    quality = (
        _count_baseline_quality(baseline)
        if baseline is not None
        else {
            "empty_code": 0,
            "qty_errors": 0,
            "unknown_google": 0,
            "tag_warnings": 0,
            "duplicates": 0,
        }
    )
    error_count = document.validation.error_count
    warn_count = sum(1 for item in document.validation.issues if item.level == "WARN")
    overlay_count = sum(
        1 for item in document.validation.issues if item.blocks_overlay
    )
    if baseline is not None:
        error_count = baseline.blocking_issue_count
        warn_count = baseline.warn_count
        overlay_count = baseline.overlay_count
    match_count = hybrid.match_count if hybrid is not None else 0
    mismatch_count = hybrid.mismatch_count if hybrid is not None else 0
    ds_only_count = hybrid.ds_only_count if hybrid is not None else 0
    rfp_only_count = hybrid.rfp_only_count if hybrid is not None else 0
    blocked_count = hybrid.blocked_count if hybrid is not None else 0
    warn_count += extra_warn
    if kind == "baseline" and baseline is not None:
        summary = baseline.summary_line()
    elif kind == "hybrid" and hybrid is not None:
        summary = hybrid.summary_line()
    elif kind == "coverage":
        missing_rfp = sum(
            1
            for row in coverage_rows
            if any("нет файла RFP" in cell or "нет RFP" in cell for cell in row.cells)
        )
        summary = (
            f"Покрытие: групп={len(group_rows)}, файлов ДС={len(ds_files)}, "
            f"RFP={len(rfp_files)}, без RFP={missing_rfp}, "
            f"ERROR={error_count}, WARN={warn_count}"
        )
    else:
        summary = document.validation.summary_line()
        if placement == "dated":
            format_label = "new, копия с датой"
        elif archived_now:
            format_label = "new, старый формат переименован"
        elif migrated_registry_path is not None:
            format_label = "new"
        else:
            format_label = registry_format
        summary = (
            f"{summary}; формат={format_label} v{document.format_version}; "
            f"активных={len(document.active_rows)}; групп={len(group_rows)}"
        )
    next_step = ""
    if kind == "registry":
        next_step = _registry_next_step(
            document,
            migrated_path=migrated_registry_path,
            registry_format=registry_format,
            placement=placement,
            archived_now=archived_now,
        )
    return DsCockpitSnapshot(
        kind=kind,
        summary=summary,
        registry_path=document.path,
        registry_format=registry_format,
        format_version=document.format_version,
        output_dir=output_dir,
        rfp_root=rfp_root,
        source_root=source_root,
        ul_root=ul_root,
        baseline_path=baseline.baseline_path if baseline is not None else None,
        hybrid_path=hybrid.hybrid_path if hybrid is not None else None,
        report_path=hybrid.report_path if hybrid is not None else None,
        migrated_registry_path=migrated_registry_path,
        next_step=next_step,
        active_count=len(document.active_rows),
        group_count=len(group_rows),
        error_count=error_count,
        warn_count=warn_count,
        overlay_count=overlay_count,
        match_count=match_count,
        mismatch_count=mismatch_count,
        ds_only_count=ds_only_count,
        rfp_only_count=rfp_only_count,
        blocked_count=blocked_count,
        file_count=len(ds_files),
        rfp_file_count=len(rfp_files),
        empty_code=quality["empty_code"],
        qty_errors=quality["qty_errors"],
        unknown_google=quality["unknown_google"],
        tag_warnings=quality["tag_warnings"],
        duplicates=quality["duplicates"],
        registry_rows=registry_rows,
        group_rows=group_rows,
        file_rows=file_rows,
        coverage_rows=coverage_rows,
    )


def _print_snapshot(snapshot: DsCockpitSnapshot) -> None:
    _emit(snapshot.summary)
    _emit(f"Реестр: {snapshot.registry_path} ({snapshot.registry_format})")
    if snapshot.source_root is not None:
        _emit(f"ДС: {snapshot.source_root}")
    if snapshot.rfp_root is not None:
        _emit(f"RFP: {snapshot.rfp_root}")
    if snapshot.ul_root is not None:
        _emit(f"УЛ: {snapshot.ul_root}")
    if snapshot.output_dir is not None and snapshot.kind != "registry":
        _emit(f"Отчёты: {snapshot.output_dir}")
    if snapshot.baseline_path is not None:
        _emit(f"{BASELINE_XLSX_NAME}: {snapshot.baseline_path}")
    if snapshot.hybrid_path is not None:
        _emit(f"{HYBRID_XLSX_NAME}: {snapshot.hybrid_path}")
    if snapshot.report_path is not None:
        _emit(f"{snapshot.report_path.name}: {snapshot.report_path}")
    if (
        snapshot.migrated_registry_path is not None
        and snapshot.migrated_registry_path != snapshot.registry_path
    ):
        _emit(f"Копия реестра: {snapshot.migrated_registry_path}")
    _emit(
        "Счётчики: "
        f"ERROR={snapshot.error_count}, WARN={snapshot.warn_count}, "
        f"OVERLAY={snapshot.overlay_count}, MATCH={snapshot.match_count}, "
        f"без кода={snapshot.empty_code}, qty={snapshot.qty_errors}, "
        f"вне Google={snapshot.unknown_google}, теги={snapshot.tag_warnings}, "
        f"дубли={snapshot.duplicates}"
    )
    if snapshot.next_step:
        _emit(snapshot.next_step)


def _fail(message: str, result_path: Path | None = None) -> DsJobResult:
    _emit(message)
    return DsJobResult(
        success=False,
        message=message,
        result_path=str(result_path) if result_path is not None else None,
    )


def _store(snapshot: DsCockpitSnapshot) -> DsCockpitSnapshot:
    global _last_cockpit
    _last_cockpit = snapshot
    _print_snapshot(snapshot)
    return snapshot


def _ensure_legend_on_copy(path: Path) -> None:
    """Add the how-to sheet to a non-canon copy. A locked Excel file is a warning."""

    try:
        added = ensure_registry_legend(path)
    except OSError as exc:
        _emit(
            f"Не удалось дописать лист «{LEGEND_SHEET_NAME}» "
            f"(закройте файл в Excel, если он открыт): {exc}"
        )
        return
    except Exception as exc:
        _emit(f"Лист «{LEGEND_SHEET_NAME}» не добавлен: {exc}")
        return
    if added:
        _emit(
            f"Добавлен лист «{LEGEND_SHEET_NAME}»: что обязательно и на что влияет. "
            "Канон UNC не менялся."
        )


def _is_canon_home(path: Path) -> bool:
    left = str(path).replace("/", "\\").casefold().rstrip("\\")
    right = str(DEFAULT_RFP_BASE).replace("/", "\\").casefold().rstrip("\\")
    return left == right


def _map_ds_paths(root: Path, source_ids: Sequence[str]) -> dict[str, tuple[Path, ...]]:
    """Pair DS workbooks to registry ids without reading workbook bytes."""

    found: dict[str, list[Path]] = defaultdict(list)
    ids = tuple(source_ids)
    try:
        candidates = list(root.rglob("*"))
    except OSError:
        return {}
    for path in candidates:
        if not path.is_file() or path.suffix.lower() not in {".xlsx", ".xlsm"}:
            continue
        if path.name.startswith("~$"):
            continue
        try:
            rel = path.relative_to(root).as_posix()
        except OSError:
            continue
        resolved = resolve_ds_source_id(rel, ids)
        if resolved.source_id:
            found[resolved.source_id].append(path)
    return {key: tuple(paths) for key, paths in found.items()}


def _reconcile_missing_rfp(
    rows: Sequence[Any],
    rfp_files: dict[str, tuple[Path, ...]],
    scanned_rfp: bool,
) -> dict[str, str]:
    """Write-only label. A named RFP with no file on disk is «нет файла».

    «совпало» and «расхождение» are the launch decision and are not guessed
    from file presence. The summary does not read this cell back.
    """

    if not scanned_rfp:
        return {}
    on_disk = {
        path.name.casefold()
        for paths in rfp_files.values()
        for path in paths
    }
    labels: dict[str, str] = {}
    for row in rows:
        source_id = str(getattr(row, "source_id", "") or "")
        if not source_id or not getattr(row, "is_active", False):
            continue
        named = False
        present = False
        for rel in getattr(row, "relations", ()):
            if rel.rfp_file:
                named = True
                if any(
                    name.casefold() in on_disk for name in _ds_file_names(rel.rfp_file)
                ):
                    present = True
            if rel.rfp_key and rfp_files.get(rel.rfp_key):
                named = True
                present = True
            elif rel.rfp_key:
                named = True
        if named and not present:
            labels[source_id] = RFP_RECONCILE_NO_FILE
    return labels


def _registry_links(
    rows: Sequence[Any],
    *,
    ds_root: Path | None,
    ul_root: Path | None,
    rfp_root: Path | None,
) -> RegistryLinks:
    source_ids = [row.source_id for row in rows if getattr(row, "source_id", "")]
    scanned_ds = ds_root is not None and _is_dir(ds_root)
    scanned_ul = ul_root is not None and _is_dir(ul_root)
    scanned_rfp = rfp_root is not None and _is_dir(rfp_root)
    rfp_files = index_rfp_files(rfp_root) if scanned_rfp else {}
    ul_by_actual, ul_names = scan_ul_catalog(ul_root) if scanned_ul else ({}, frozenset())
    return RegistryLinks(
        ds_files=_map_ds_paths(ds_root, source_ids) if scanned_ds and ds_root else {},
        rfp_files=rfp_files,
        ul_by_actual=ul_by_actual,
        scanned_ds=scanned_ds,
        scanned_rfp=scanned_rfp,
        scanned_ul=scanned_ul,
        ul_names=ul_names,
        reconcile_by_source=_reconcile_missing_rfp(rows, rfp_files, scanned_rfp),
    )


def _links_for_install(
    rows: Sequence[Any],
    ul_root: Path | None,
    attach_links: bool,
) -> RegistryLinks | None:
    if not attach_links:
        return None
    return _registry_links(
        rows,
        ds_root=_resolve_source_root(None),
        ul_root=ul_root,
        rfp_root=_resolve_rfp_root(None),
    )


def _install_rows(rows: list[Any], home: Path, links: RegistryLinks | None = None):
    """Install rows into the working registry folder.

    Returns:
        ``(RegistryInstall | None, error_message | None)``.
    """

    err = _ensure_output_dir(home)
    if err:
        return None, err
    try:
        installed = install_working_registry(rows, home, links=links)
    except DsRegistryError as exc:
        return None, f"реестр не записан: {exc}"
    _install_rows.last = installed  # type: ignore[attr-defined]
    return installed, None


def _emit_install(installed: Any) -> None:
    if installed.mode == "dated":
        _emit(
            f"{CANONICAL_REGISTRY_NAME} открыт в Excel. "
            f"Копия с датой: {installed.path}"
        )
        return
    if installed.archived_now and installed.archived_path is not None:
        _emit(
            f"Старый формат переименован в {installed.archived_path.name}. "
            f"Рабочий реестр: {installed.path}"
        )
        return
    _emit(f"Рабочий реестр: {installed.path}")


def _open_registry_for_job(
    registry_path: Path,
    *,
    ul_root: Path | None,
    output_dir: Path | None,
    migrate_legacy: bool,
    install_into_home: bool = False,
    attach_links: bool = False,
    ds_root: Path | None = None,
    rfp_root: Path | None = None,
    on_lap: Callable[[str, float], None] | None = None,
) -> tuple[DsRegistryDocument | None, str, Path | None, str | None, int]:
    """Load new-format registry, or migrate legacy into ``output_dir``.

    Returns:
        ``(document, format, migrated_path, error_message, extra_warn)``.
    """
    extra_warn = 0
    if not _is_file(registry_path):
        return None, "unknown", None, f"файл реестра не найден: {registry_path}", 0
    try:
        fmt = detect_registry_format(registry_path)
    except Exception as exc:
        cause = exc.__cause__
        detail = str(exc)
        if cause is not None:
            detail = f"{detail} ({type(cause).__name__}: {cause})"
        return None, "unknown", None, f"не удалось определить формат реестра: {detail}", 0
    if fmt == "new":
        document, err = _load_new_registry(
            registry_path,
            ul_root=ul_root,
            ds_root=ds_root,
            rfp_root=rfp_root,
            on_lap=on_lap,
        )
        if (
            not install_into_home
            or document is None
            or err is not None
            or output_dir is None
        ):
            return document, fmt, None, err, extra_warn
        started = time.perf_counter()
        links = _links_for_install(document.rows, ul_root, attach_links)
        if on_lap is not None:
            on_lap("ссылки ДС/RFP/УЛ", time.perf_counter() - started)
        started = time.perf_counter()
        installed, install_err = _install_rows(document.rows, output_dir, links)
        if on_lap is not None:
            on_lap("запись реестра", time.perf_counter() - started)
        if install_err or installed is None:
            return None, fmt, None, install_err or "реестр не записан", extra_warn
        _emit_install(installed)
        reloaded, err = _load_new_registry(
            installed.path,
            ul_root=ul_root,
            ds_root=ds_root,
            rfp_root=rfp_root,
            on_lap=on_lap,
        )
        return reloaded, "new", installed.path, err, extra_warn
    if fmt != "legacy":
        return (
            None,
            fmt,
            None,
            (
                "формат реестра не распознан "
                "(нужен лист «Реестр ДС» или legacy 4 колонки)"
            ),
            extra_warn,
        )
    if not migrate_legacy:
        return (
            None,
            fmt,
            None,
            (
                "реестр в старом формате; проверка без миграции невозможна. "
                "Запустите «Проверить реестр» — копия будет записана в папку отчётов."
            ),
            extra_warn,
        )
    if output_dir is None:
        return None, fmt, None, "не задана папка рабочего реестра", extra_warn
    if not install_into_home:
        migrated = output_dir / ROBOT_REGISTRY_COPY_NAME
        if _same_file(migrated, DEFAULT_REGISTRY_PATH) or _same_file(
            migrated, registry_path
        ):
            return (
                None,
                fmt,
                None,
                "миграция отказана: выход совпадает с исходным реестром",
                extra_warn,
            )
        try:
            result = migrate_registry(registry_path, migrated)
        except DsRegistryError as exc:
            return None, fmt, None, f"миграция реестра не выполнена: {exc}", extra_warn
        document, err = _load_new_registry(
            result.output_path,
            ul_root=ul_root,
            ds_root=ds_root,
            rfp_root=rfp_root,
            on_lap=on_lap,
        )
        return document, "new", result.output_path, err, extra_warn
    try:
        legacy_rows = read_legacy_registry(registry_path)
        new_rows, migrate_issues = migrate_legacy_rows(legacy_rows)
    except DsRegistryError as exc:
        return None, fmt, None, f"миграция реестра не выполнена: {exc}", extra_warn
    started = time.perf_counter()
    links = _links_for_install(new_rows, ul_root, attach_links)
    if on_lap is not None:
        on_lap("ссылки ДС/RFP/УЛ", time.perf_counter() - started)
    started = time.perf_counter()
    installed, install_err = _install_rows(new_rows, output_dir, links)
    if on_lap is not None:
        on_lap("запись реестра", time.perf_counter() - started)
    if install_err or installed is None:
        return None, fmt, None, install_err or "реестр не записан", extra_warn
    _emit_install(installed)
    document, err = _load_new_registry(
        installed.path,
        ul_root=ul_root,
        ds_root=ds_root,
        rfp_root=rfp_root,
        on_lap=on_lap,
    )
    if document is not None:
        report = default_migration_report_path(installed.path)
        try:
            from RFQ.rfp_parts.ds_registry import _write_migration_report

            _write_migration_report(
                report,
                source_path=registry_path,
                output_path=installed.path,
                legacy_rows=legacy_rows,
                new_rows=document.rows,
                issues=migrate_issues,
                validation=document.validation,
            )
        except OSError as exc:
            _emit(f"Отчёт миграции не записан: {exc}")
        else:
            _emit(f"{MIGRATION_REPORT_PREFIX}: {report}")
    return document, "new", installed.path, err, extra_warn


def run_ds_registry_check_job(
    registry_path: str | Path | None = None,
    output_dir: str | Path | None = None,
    ul_root: str | Path | None = None,
    ds_root: str | Path | None = None,
    rfp_root: str | Path | None = None,
) -> DsJobResult:
    """Validate the registry and keep one working file in the ``_RFP`` folder.

    A legacy file is renamed to ``Реестр_ДС_УЛ_old.xlsx`` and replaced. If Excel
    holds the file open, a dated copy is written and older dated copies beyond
    five are deleted. The program then reads whichever file is newest.

    Args:
        registry_path: Registry xlsx. Default ``gui_paths.last_ds_registry_file``.
        output_dir: Folder for the working registry. Default ``_RFP``.
            Tests pass a temp folder. Production passes ``DEFAULT_RFP_BASE``.
        ul_root: TSD root for UL-folder existence. Default
            ``gui_paths.last_tsd_packing_folder``.
        ds_root: Trusted DS folder. Folders and files not named on a row
            become yellow placeholder rows. Omitted means that catalog is
            not scanned.
        rfp_root: Root-only RFP folder. Same orphan rule. Omitted means the
            RFP catalog is not scanned.

    Returns:
        Russian summary. ``success`` is false when the file cannot be loaded.
    """
    os.environ.setdefault("PYTHONUTF8", "1")
    _install_rows.last = None  # type: ignore[attr-defined]
    home = Path(output_dir) if output_dir else DEFAULT_RFP_BASE
    requested = _resolve_registry_path(registry_path)
    latest = resolve_latest_registry(home)
    path = latest if _is_file(latest) else requested
    ul_path = _resolve_ul_root(ul_root)
    _emit("Проверка реестра ДС…")
    session = DsProgressSession(_emit)
    phase_start = time.perf_counter()
    session.phase_start("чтение и проверка реестра")

    def _on_lap(label: str, seconds: float) -> None:
        session.message(f"{label} — {seconds:.2f} с")

    document, fmt, migrated, err, extra_warn = _open_registry_for_job(
        path,
        ul_root=ul_path,
        output_dir=home,
        migrate_legacy=True,
        install_into_home=True,
        attach_links=_is_canon_home(home),
        ds_root=Path(ds_root) if ds_root else None,
        rfp_root=Path(rfp_root) if rfp_root else None,
        on_lap=_on_lap,
    )
    session.phase_done(
        "чтение и проверка реестра",
        time.perf_counter() - phase_start,
    )
    if document is None:
        return _fail(err or "реестр не прочитан")
    installed = getattr(_install_rows, "last", None)
    snapshot = _store(
        _snapshot_from_registry(
            kind="registry",
            document=document,
            registry_format=fmt,
            output_dir=home if migrated is not None else None,
            ul_root=ul_path,
            migrated_registry_path=migrated,
            extra_warn=extra_warn,
            placement=installed.mode if installed is not None else "",
            archived_now=bool(installed and installed.archived_now),
        )
    )
    success = document.validation.is_ok
    result_path = migrated or document.path
    return DsJobResult(success=success, message=snapshot.summary, result_path=str(result_path))


def run_ds_baseline_job(
    source_root: str | Path | None = None,
    registry_path: str | Path | None = None,
    output_dir: str | Path | None = None,
    ul_root: str | Path | None = None,
    **kwargs: Any,
) -> DsJobResult:
    """Audit trusted DS workbooks and write ``Свод ДС для запуска.xlsx``.

    Args:
        source_root: Recursive DS folder. Default ``last_ds_trusted_folder``.
        registry_path: Canonical registry.
        output_dir: Stamp folder for reports and the baseline.
        ul_root: Optional TSD root for registry UL checks.
        **kwargs: Forwarded to ``build_ds_baseline`` (converter, google_index).

    Returns:
        Russian summary. ``success`` follows ``not baseline.blocking``.
    """
    os.environ.setdefault("PYTHONUTF8", "1")
    source = _resolve_source_root(source_root)
    if source is None or not _is_dir(source):
        return _fail(f"папка доверенных ДС не найдена: {source or '—'}")
    path = _resolve_registry_path(registry_path)
    reports = _resolve_output_dir(output_dir)
    err = _ensure_output_dir(reports)
    if err:
        return _fail(err)
    ul_path = _resolve_ul_root(ul_root)
    _emit("Сбор входа Только ДС…")
    document, fmt, migrated, load_err, extra_warn = _open_registry_for_job(
        path,
        ul_root=ul_path,
        output_dir=reports,
        migrate_legacy=True,
        ds_root=source,
    )
    if document is None:
        return _fail(load_err or "реестр не прочитан")
    _emit("Аудит ДС и запись свода…")
    session, _tracker, progress_hooks = _baseline_progress_hooks("ДС")
    audit_start = time.perf_counter()
    session.phase_start("аудит ДС")
    merged_kwargs = {**kwargs, **progress_hooks}
    try:
        baseline = build_ds_baseline(source, document, reports, **merged_kwargs)
    except Exception as exc:
        return _fail(f"свод ДС не собран: {type(exc).__name__}: {exc}", reports)
    session.phase_done("аудит ДС", time.perf_counter() - audit_start)
    snapshot = _store(
        _snapshot_from_registry(
            kind="baseline",
            document=document,
            registry_format=fmt,
            output_dir=baseline.output_dir,
            source_root=source,
            ul_root=ul_path,
            baseline=baseline,
            migrated_registry_path=migrated,
            extra_warn=extra_warn,
        )
    )
    result_path = baseline.baseline_path or baseline.output_dir
    return DsJobResult(
        success=not baseline.blocking,
        message=snapshot.summary,
        result_path=str(result_path),
    )


def run_ds_hybrid_job(
    source_root: str | Path | None = None,
    registry_path: str | Path | None = None,
    output_dir: str | Path | None = None,
    ul_root: str | Path | None = None,
    rfp_root: str | Path | None = None,
    **kwargs: Any,
) -> DsJobResult:
    """Build DS baseline then overlay root RFP into the hybrid workbook.

    Args:
        source_root: Recursive DS folder.
        registry_path: Canonical registry.
        output_dir: Stamp folder for baseline, hybrid and the audit report.
        ul_root: Optional TSD root.
        rfp_root: Root-only ``RFP_Зиновьев``. Default ``DEFAULT_PARTS_DIR``.
        **kwargs: ``converter`` (DS), ``rfp_converter``, ``loader`` and other
            ``build_ds_baseline`` / ``build_ds_rfp_hybrid`` extras.

    Returns:
        Russian summary. Missing RFP root is a warning inside hybrid, not a
        crash. ``success`` is false when hybrid is blocked.
    """
    os.environ.setdefault("PYTHONUTF8", "1")
    source = _resolve_source_root(source_root)
    if source is None or not _is_dir(source):
        return _fail(f"папка доверенных ДС не найдена: {source or '—'}")
    path = _resolve_registry_path(registry_path)
    reports = _resolve_output_dir(output_dir)
    err = _ensure_output_dir(reports)
    if err:
        return _fail(err)
    ul_path = _resolve_ul_root(ul_root)
    rfp_path = _resolve_rfp_root(rfp_root)
    extra_warn = 0
    if not _is_dir(rfp_path):
        extra_warn += 1
        _emit(f"WARN: корень RFP не найден: {rfp_path}")
    _emit("Проверка RFP и наложение на ДС…")
    document, fmt, migrated, load_err, migrate_warn = _open_registry_for_job(
        path,
        ul_root=ul_path,
        output_dir=reports,
        migrate_legacy=True,
        ds_root=source,
        rfp_root=rfp_path,
    )
    extra_warn += migrate_warn
    if document is None:
        return _fail(load_err or "реестр не прочитан")
    ds_kwargs = {
        key: value
        for key, value in kwargs.items()
        if key in {"converter", "google_index", "matrix_path", "write_baseline", "stamp", "equipment_by_code"}
    }
    rfp_kwargs = {
        key: value
        for key, value in kwargs.items()
        if key in {"google_index", "matrix_path", "stamp", "equipment_by_code", "write_hybrid"}
    }
    if "rfp_converter" in kwargs:
        rfp_kwargs["converter"] = kwargs["rfp_converter"]
    if "loader" in kwargs:
        rfp_kwargs["loader"] = kwargs["loader"]
    session, _ds_tracker, ds_progress_hooks = _baseline_progress_hooks("ДС")
    ds_kwargs.update(ds_progress_hooks)
    _rfp_tracker, rfp_progress_hooks = _rfp_progress_hooks("RFP")
    rfp_kwargs.update(rfp_progress_hooks)
    _emit("Аудит ДС…")
    audit_start = time.perf_counter()
    session.phase_start("аудит ДС")
    try:
        baseline = build_ds_baseline(source, document, reports, **ds_kwargs)
    except Exception as exc:
        return _fail(f"свод ДС не собран: {type(exc).__name__}: {exc}", reports)
    session.phase_done("аудит ДС", time.perf_counter() - audit_start)
    _emit("Сверка корневого RFP и overlay…")
    hybrid_start = time.perf_counter()
    session.phase_start("overlay RFP")
    try:
        hybrid = build_ds_rfp_hybrid(baseline, document, rfp_path, reports, **rfp_kwargs)
    except Exception as exc:
        return _fail(f"гибрид ДС-RFP не собран: {type(exc).__name__}: {exc}", reports)
    session.phase_done("overlay RFP", time.perf_counter() - hybrid_start)
    snapshot = _store(
        _snapshot_from_registry(
            kind="hybrid",
            document=document,
            registry_format=fmt,
            output_dir=hybrid.output_dir,
            source_root=source,
            rfp_root=rfp_path,
            ul_root=ul_path,
            baseline=baseline,
            hybrid=hybrid,
            migrated_registry_path=migrated,
            extra_warn=extra_warn,
        )
    )
    result_path = hybrid.hybrid_path or hybrid.report_path or reports
    return DsJobResult(
        success=not hybrid.blocking,
        message=snapshot.summary,
        result_path=str(result_path),
    )


def run_ds_coverage_job(
    source_root: str | Path | None = None,
    registry_path: str | Path | None = None,
    output_dir: str | Path | None = None,
    ul_root: str | Path | None = None,
    rfp_root: str | Path | None = None,
) -> DsJobResult:
    """Name-only coverage: registry ↔ DS files ↔ root RFP ↔ UL folders.

    Args:
        source_root: Recursive DS folder.
        registry_path: Canonical registry.
        output_dir: Unused for writes; accepted for a uniform job signature.
        ul_root: TSD root.
        rfp_root: Root-only RFP folder.

    Returns:
        Russian summary. Missing RFP is a WARN and does not fail the job.
        ``success`` is false only when the registry cannot be loaded.
    """
    del output_dir
    os.environ.setdefault("PYTHONUTF8", "1")
    source = _resolve_source_root(source_root)
    path = _resolve_registry_path(registry_path)
    ul_path = _resolve_ul_root(ul_root)
    rfp_path = _resolve_rfp_root(rfp_root)
    extra_warn = 0
    _emit("Только покрытие реестр ↔ ДС ↔ RFP ↔ УЛ…")
    session = DsProgressSession(_emit)
    if source is None or not _is_dir(source):
        extra_warn += 1
        _emit(f"WARN: папка доверенных ДС не найдена: {source or '—'}")
    if not _is_dir(rfp_path):
        extra_warn += 1
        _emit(f"WARN: корень RFP не найден: {rfp_path}")
    load_start = time.perf_counter()
    session.phase_start("реестр и каталоги")
    document, fmt, migrated, load_err, migrate_warn = _open_registry_for_job(
        path,
        ul_root=ul_path,
        output_dir=None,
        migrate_legacy=False,
        ds_root=source,
        rfp_root=rfp_path,
    )
    session.phase_done("реестр и каталоги", time.perf_counter() - load_start)
    extra_warn += migrate_warn
    if document is None:
        return _fail(load_err or "реестр не прочитан")
    cover_start = time.perf_counter()
    session.phase_start("таблица покрытия")
    snapshot = _store(
        _snapshot_from_registry(
            kind="coverage",
            document=document,
            registry_format=fmt,
            source_root=source if source is not None and _is_dir(source) else None,
            rfp_root=rfp_path,
            ul_root=ul_path,
            migrated_registry_path=migrated,
            extra_warn=extra_warn,
        )
    )
    session.phase_done("таблица покрытия", time.perf_counter() - cover_start)
    return DsJobResult(
        success=True,
        message=snapshot.summary,
        result_path=str(document.path),
    )


__all__ = [
    "CockpitRow",
    "DEFAULT_PARTS_DIR",
    "DEFAULT_REGISTRY_PATH",
    "DEFAULT_REPORTS_BASE_DIR",
    "DsCockpitSnapshot",
    "DsJobResult",
    "ds_registry_robot_dir",
    "get_last_ds_cockpit",
    "run_ds_baseline_job",
    "run_ds_coverage_job",
    "run_ds_hybrid_job",
    "run_ds_registry_check_job",
]
