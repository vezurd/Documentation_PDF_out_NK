"""Freshness gate for DS baseline / hybrid workbooks used by RFP · Запуск.

Stable output lives under ``DEFAULT_REPORTS_BASE_DIR/_ds_baseline`` and
``_ds_hybrid``. Rebuilds when the workbook is missing or the fingerprint of
sources + registry changed. Hybrid never falls back to ``rfp_parts_net.xlsx``.
"""

from __future__ import annotations

import json
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from RFQ.rfp_parts.analyze_rfp_parts import (
    DEFAULT_PARTS_DIR,
    DEFAULT_REPORTS_BASE_DIR,
)
from RFQ.rfp_parts.ds_baseline import (
    ALGORITHM_VERSION as BASELINE_ALGORITHM_VERSION,
    BASELINE_XLSX_NAME,
    DUPLICATE_TAGS_REPORT_PREFIX,
    EMPTY_CODE_REPORT_PREFIX,
    QUALITY_REPORT_PREFIX,
    STRUCTURE_REPORT_PREFIX,
    TAG_MISMATCH_REPORT_PREFIX,
    DsBaselineResult,
    DsUnitsConverter,
    _tree_fingerprint,
    build_ds_baseline,
    collect_ds_workbooks,
    latest_audit_report_dir,
)
from RFQ.rfp_parts.ds_registry import (
    DEFAULT_REGISTRY_PATH,
    DsRegistryDocument,
    load_registry,
)
from RFQ.rfp_parts.ds_rfp_hybrid import (
    ALGORITHM_VERSION as HYBRID_ALGORITHM_VERSION,
    HYBRID_REPORT_PREFIX,
    HYBRID_XLSX_NAME,
    DsRfpHybridResult,
    _hybrid_fingerprint,
    build_ds_rfp_hybrid,
    collect_rfp_workbooks,
)
from RFQ.rfp_parts.ds_rfp_tag_placement import MIX_MIXED, MIX_SEPARATE
from RFQ.units_convert.models import GoogleUnitsIndex

INPUT_MODE_LEGACY_NET = "legacy_net"
INPUT_MODE_DS_ONLY = "ds_only"
INPUT_MODE_HYBRID = "hybrid"
ALLOWED_INPUT_MODES: frozenset[str] = frozenset(
    {INPUT_MODE_LEGACY_NET, INPUT_MODE_DS_ONLY, INPUT_MODE_HYBRID}
)

DS_BASELINE_DIR_NAME = "_ds_baseline"
DS_HYBRID_DIR_NAME = "_ds_hybrid"
STATE_JSON_NAME = "ds_preflight_state.json"

SIDECAR_NAME_PREFIXES: tuple[str, ...] = (
    STRUCTURE_REPORT_PREFIX,
    "Отчет по структуре файлов - ДС",
    QUALITY_REPORT_PREFIX,
    EMPTY_CODE_REPORT_PREFIX,
    DUPLICATE_TAGS_REPORT_PREFIX,
    TAG_MISMATCH_REPORT_PREFIX,
    "Отчет по сверке ДС-RFP",
    "Отчет по миграции реестра ДС",
    "дробные значения после конвертации ДС",
    "дробные значения после конвертации RFP",
)
_SIDECAR_PREFIXES_CF: tuple[str, ...] = tuple(
    item.casefold() for item in SIDECAR_NAME_PREFIXES
)
_NET_NAMES_CF: frozenset[str] = frozenset(
    {
        BASELINE_XLSX_NAME.casefold(),
        HYBRID_XLSX_NAME.casefold(),
        STATE_JSON_NAME.casefold(),
    }
)


class DsBaselineBlockedError(Exception):
    """DS audit has blocking issues; ``Свод ДС для запуска.xlsx`` was not written."""

    def __init__(
        self,
        message: str,
        *,
        output_dir: Path | None = None,
        result: DsBaselineResult | None = None,
    ) -> None:
        super().__init__(message)
        self.output_dir = output_dir
        self.result = result


class DsHybridBlockedError(Exception):
    """Hybrid overlay has global blockers; hybrid workbook was not written."""

    def __init__(
        self,
        message: str,
        *,
        output_dir: Path | None = None,
        result: DsRfpHybridResult | None = None,
    ) -> None:
        super().__init__(message)
        self.output_dir = output_dir
        self.result = result


@dataclass
class DsBaselineFreshness:
    """Whether the stable DS baseline workbook is still valid for Step1."""

    needs_rebuild: bool
    reason: str
    baseline_path: Path | None
    output_dir: Path
    fingerprint: str = ""
    result: DsBaselineResult | None = None


@dataclass
class DsHybridFreshness:
    """Whether the stable hybrid workbook is still valid for Step1."""

    needs_rebuild: bool
    reason: str
    hybrid_path: Path | None
    output_dir: Path
    fingerprint: str = ""
    summary: str = ""
    baseline_freshness: DsBaselineFreshness | None = None
    result: DsRfpHybridResult | None = None
    derived_registry_path: Path | None = None


def resolve_input_mode(config: dict[str, Any] | None = None) -> str:
    """Return ``legacy_net`` / ``ds_only`` / ``hybrid`` from ``rfp_parts``.

    Unknown, empty or missing values default to ``legacy_net``.

    Args:
        config: Full RFP profile (or None).

    Returns:
        One of ``ALLOWED_INPUT_MODES``.
    """

    raw = _parts_section(config).get("input_mode", INPUT_MODE_LEGACY_NET)
    text = str(raw or "").strip()
    if text in ALLOWED_INPUT_MODES:
        return text
    return INPUT_MODE_LEGACY_NET


def resolve_launch_mix_mode(config: dict[str, Any] | None = None) -> str:
    """Return launch cluster mix from ``rfp_parts.launch_mix_mode``.

    Does not read ``collect_mix_mode``. Unknown, empty or missing values
    default to ``separate``.

    Args:
        config: Full RFP profile (or None).

    Returns:
        ``MIX_SEPARATE`` or ``MIX_MIXED``.
    """

    raw = _parts_section(config).get("launch_mix_mode", MIX_SEPARATE)
    text = str(raw or "").strip()
    if text in (MIX_SEPARATE, MIX_MIXED):
        return text
    return MIX_SEPARATE


def ds_baseline_output_dir(reports_base: str | Path | None = None) -> Path:
    """Stable folder for ``Свод ДС для запуска.xlsx`` and DS audit reports."""

    return Path(reports_base or DEFAULT_REPORTS_BASE_DIR) / DS_BASELINE_DIR_NAME


def ds_hybrid_output_dir(reports_base: str | Path | None = None) -> Path:
    """Stable folder for ``Свод ДС-RFP для запуска.xlsx`` and overlay reports."""

    return Path(reports_base or DEFAULT_REPORTS_BASE_DIR) / DS_HYBRID_DIR_NAME


def resolve_ds_baseline_xlsx(*, reports_base: str | Path | None = None) -> Path | None:
    """Return the stable DS baseline workbook if it exists as a file."""

    path = ds_baseline_output_dir(reports_base) / BASELINE_XLSX_NAME
    try:
        return path if path.is_file() else None
    except OSError:
        return None


def resolve_ds_hybrid_xlsx(*, reports_base: str | Path | None = None) -> Path | None:
    """Return the stable hybrid workbook if it exists as a file."""

    path = ds_hybrid_output_dir(reports_base) / HYBRID_XLSX_NAME
    try:
        return path if path.is_file() else None
    except OSError:
        return None


def resolve_hybrid_derived_registry(
    *,
    reports_base: str | Path | None = None,
    registry_path: str | Path | None = None,
) -> Path | None:
    """Return the derived planting registry beside the hybrid summary.

    Looks up ``derived_registry_name`` in the hybrid state json, then the
    basename of ``registry_path``. Missing files yield ``None``.

    Args:
        reports_base: Parent of ``_ds_hybrid``.
        registry_path: Source registry whose basename was copied.

    Returns:
        Existing derived workbook path, or ``None``.
    """

    hybrid_dir = ds_hybrid_output_dir(reports_base)
    names: list[str] = []
    stored = _load_state(hybrid_dir)
    if stored:
        stored_name = str(stored.get("derived_registry_name") or "").strip()
        if stored_name:
            names.append(stored_name)
    if registry_path is not None and str(registry_path).strip():
        names.append(Path(str(registry_path).strip()).name)
    seen: set[str] = set()
    for name in names:
        if not name or name in seen:
            continue
        seen.add(name)
        path = hybrid_dir / name
        try:
            if path.is_file():
                return path
        except OSError:
            continue
    return None


def copy_ds_sidecars_to_result_dir(
    stamp_dir: Path | str | None,
    result_dir: Path | str | None,
) -> list[Path]:
    """Copy Russian-named DS/hybrid reports from the latest stamp folder.

    The net workbooks themselves are not copied. When ``stamp_dir`` contains
    ``YYYY.MM.DD_HH.MM`` children, only the newest of those is read. Missing
    sources are skipped. ``OSError`` is printed and the file is skipped.

    Args:
        stamp_dir: ``_ds_baseline``, ``_ds_hybrid``, or one stamp child.
        result_dir: Launch ``_результат_проверки_*`` folder.

    Returns:
        Destination paths that were copied.
    """

    if stamp_dir is None or result_dir is None:
        return []
    src_dir = latest_audit_report_dir(stamp_dir)
    if src_dir is None:
        return []
    dest_dir = Path(result_dir)
    try:
        if not src_dir.is_dir():
            return []
        names = list(src_dir.iterdir())
    except OSError as exc:
        print(f"Не удалось прочитать папку отчётов ДС {src_dir}: {exc}")
        return []

    copied: list[Path] = []
    for src in names:
        try:
            if not src.is_file() or not _is_sidecar_name(src.name):
                continue
            dest_dir.mkdir(parents=True, exist_ok=True)
            dest = dest_dir / src.name
            shutil.copy2(src, dest)
            copied.append(dest)
        except OSError as exc:
            print(f"Не удалось скопировать отчёт ДС {src.name}: {exc}")
    return copied


def ensure_ds_baseline_current(
    *,
    source_root: str | Path | None = None,
    registry_path: str | Path | None = None,
    reports_base: str | Path | None = None,
    config: dict[str, Any] | None = None,
    converter: DsUnitsConverter | None = None,
    google_index: GoogleUnitsIndex | None = None,
    matrix_path: str | Path | None = None,
    progress: bool = False,
) -> tuple[Path, DsBaselineFreshness]:
    """Return a fresh DS baseline, rebuilding on missing file or fingerprint change.

    Args:
        source_root: Recursive DS folder. Falls back to
            ``rfp_parts.ds_source_dir``, then GUI
            ``last_ds_trusted_folder``.
        registry_path: Canonical registry. Falls back to
            ``rfp_parts.ds_registry_path``, then GUI
            ``last_ds_registry_file``, then ``DEFAULT_REGISTRY_PATH``.
        reports_base: Parent of ``_ds_baseline``. Defaults to
            ``DEFAULT_REPORTS_BASE_DIR``.
        config: Optional RFP profile for path fallbacks.
        converter: Units conversion. Default uses RFQ plan/Google/matrix.
            Pass ``IdentityDsUnitsConverter`` for offline tests.
        google_index: Optional Google units index.
        matrix_path: Units matrix for the default converter.
        progress: When True, print short status lines.

    Returns:
        Path to ``Свод ДС для запуска.xlsx`` and the freshness decision.

    Raises:
        FileNotFoundError: Source folder is empty/missing, or rebuild did not
            write the baseline.
        DsBaselineBlockedError: Audit has blocking issues (baseline not written).
    """

    source = _resolve_source_root(source_root, config)
    registry_file = _resolve_registry_path(registry_path, config)
    out_dir = ds_baseline_output_dir(reports_base)
    out_dir.mkdir(parents=True, exist_ok=True)

    registry = load_registry(registry_file)
    fingerprint = _current_baseline_fingerprint(source, registry)
    stored = _load_state(out_dir)
    baseline_path = out_dir / BASELINE_XLSX_NAME
    stale_reason = _baseline_stale_reason(
        baseline_path=baseline_path,
        stored=stored,
        fingerprint=fingerprint,
        source_root=source,
        registry_path=registry.path,
    )
    if stale_reason is None:
        freshness = DsBaselineFreshness(
            False,
            "свод ДС актуален",
            baseline_path,
            out_dir,
            fingerprint=fingerprint,
        )
        if progress:
            print(f"[ds baseline] {freshness.reason}", flush=True)
        return baseline_path, freshness

    google, matrix = _resolve_units_inputs(
        converter=converter,
        google_index=google_index,
        matrix_path=matrix_path,
        config=config,
    )
    if progress:
        print(f"[ds baseline] пересборка: {stale_reason} → {out_dir}", flush=True)
    result = build_ds_baseline(
        source,
        registry,
        out_dir,
        write_baseline=True,
        converter=converter,
        google_index=google,
        matrix_path=matrix,
    )
    if result.blocking or result.baseline_path is None:
        raise DsBaselineBlockedError(
            result.summary_line(),
            output_dir=result.output_dir,
            result=result,
        )
    written = Path(result.baseline_path)
    if not written.is_file():
        raise FileNotFoundError(
            f"аудит ДС не создал {BASELINE_XLSX_NAME}: {out_dir}"
        )
    _write_state(
        out_dir,
        {
            "algorithm_version": BASELINE_ALGORITHM_VERSION,
            "fingerprint": result.fingerprint,
            "source_root": str(source),
            "registry_path": str(registry.path),
            "stamp": result.stamp,
        },
    )
    _prune_stale_sidecars(out_dir, keep=_baseline_keep_names(result))
    freshness = DsBaselineFreshness(
        True,
        stale_reason,
        written,
        out_dir,
        fingerprint=result.fingerprint,
        result=result,
    )
    return written, freshness


def ensure_ds_hybrid_current(
    *,
    source_root: str | Path | None = None,
    registry_path: str | Path | None = None,
    rfp_root: str | Path | None = None,
    reports_base: str | Path | None = None,
    config: dict[str, Any] | None = None,
    converter: DsUnitsConverter | None = None,
    rfp_converter: Any | None = None,
    google_index: GoogleUnitsIndex | None = None,
    matrix_path: str | Path | None = None,
    progress: bool = False,
    mix_mode: str = MIX_SEPARATE,
) -> tuple[Path, DsHybridFreshness]:
    """Return a fresh hybrid workbook; never falls back to parts net.

    Args:
        source_root: Recursive DS folder (see ``ensure_ds_baseline_current``).
        registry_path: Canonical registry path.
        rfp_root: Root-only RFP folder. Defaults to ``DEFAULT_PARTS_DIR``.
        reports_base: Parent of ``_ds_hybrid`` / ``_ds_baseline``.
        config: Optional RFP profile for path fallbacks.
        converter: DS units conversion (offline: Identity).
        rfp_converter: RFP units conversion (offline: Identity).
        google_index: Optional Google units index.
        matrix_path: Units matrix for default converters.
        progress: When True, print short status lines.
        mix_mode: Cluster placement passed to the hybrid rebuild. Defaults to
            ``separate`` so older callers keep working.

    Returns:
        Path to ``Свод ДС-RFP для запуска.xlsx`` and the freshness decision.
        The derived registry path is on ``DsHybridFreshness.derived_registry_path``.

    Raises:
        DsBaselineBlockedError: DS audit blocked; hybrid is not written.
        DsHybridBlockedError: Overlay has global blockers.
        FileNotFoundError: Source folder missing, or hybrid was not written.
    """

    source = _resolve_source_root(source_root, config)
    registry_file = _resolve_registry_path(registry_path, config)
    rfp_folder = Path(rfp_root) if rfp_root else DEFAULT_PARTS_DIR
    reports = Path(reports_base or DEFAULT_REPORTS_BASE_DIR)
    hybrid_dir = ds_hybrid_output_dir(reports)
    hybrid_dir.mkdir(parents=True, exist_ok=True)

    _baseline_path, baseline_fresh = ensure_ds_baseline_current(
        source_root=source,
        registry_path=registry_file,
        reports_base=reports,
        config=config,
        converter=converter,
        google_index=google_index,
        matrix_path=matrix_path,
        progress=progress,
    )
    registry = load_registry(registry_file)
    baseline_fp = baseline_fresh.fingerprint or _current_baseline_fingerprint(
        source, registry
    )
    dummy_baseline = DsBaselineResult(
        source_root=source,
        output_dir=ds_baseline_output_dir(reports),
        registry_path=registry.path,
        stamp="",
        fingerprint=baseline_fp,
    )
    rfp_files, _skipped = collect_rfp_workbooks(rfp_folder)
    fingerprint = _hybrid_fingerprint(
        baseline=dummy_baseline,
        registry=registry,
        files=rfp_files,
        mix_mode=mix_mode,
    )
    stored = _load_state(hybrid_dir)
    hybrid_path = hybrid_dir / HYBRID_XLSX_NAME
    source_mtime_ns = _file_mtime_ns(registry.path)
    stale_reason = _hybrid_stale_reason(
        hybrid_path=hybrid_path,
        stored=stored,
        fingerprint=fingerprint,
        source_root=source,
        registry_path=registry.path,
        rfp_root=rfp_folder,
        baseline_needs_rebuild=baseline_fresh.needs_rebuild,
        registry_mtime_ns=source_mtime_ns,
    )
    if stale_reason is None:
        derived = _derived_registry_beside(
            hybrid_dir, stored=stored, registry_path=registry.path
        )
        freshness = DsHybridFreshness(
            False,
            "свод ДС-RFP актуален",
            hybrid_path,
            hybrid_dir,
            fingerprint=fingerprint,
            summary="свод ДС-RFP актуален",
            baseline_freshness=baseline_fresh,
            derived_registry_path=derived,
        )
        if progress:
            print(f"[ds hybrid] {freshness.reason}", flush=True)
        return hybrid_path, freshness

    google, matrix = _resolve_units_inputs(
        converter=converter,
        google_index=google_index,
        matrix_path=matrix_path,
        config=config,
    )
    if progress:
        print(f"[ds hybrid] пересборка: {stale_reason} → {hybrid_dir}", flush=True)
    baseline_result = baseline_fresh.result
    if baseline_result is None:
        baseline_result = build_ds_baseline(
            source,
            registry,
            ds_baseline_output_dir(reports),
            write_baseline=True,
            converter=converter,
            google_index=google,
            matrix_path=matrix,
        )
        if baseline_result.blocking or baseline_result.baseline_path is None:
            raise DsBaselineBlockedError(
                baseline_result.summary_line(),
                output_dir=baseline_result.output_dir,
                result=baseline_result,
            )
    hybrid_result = build_ds_rfp_hybrid(
        baseline_result,
        registry,
        rfp_folder,
        hybrid_dir,
        write_hybrid=True,
        converter=rfp_converter,
        google_index=google,
        matrix_path=matrix,
        mix_mode=mix_mode,
    )
    if hybrid_result.blocking or hybrid_result.hybrid_path is None:
        raise DsHybridBlockedError(
            hybrid_result.summary_line(),
            output_dir=hybrid_result.output_dir,
            result=hybrid_result,
        )
    written = Path(hybrid_result.hybrid_path)
    if not written.is_file():
        raise FileNotFoundError(
            f"сверка ДС-RFP не создала {HYBRID_XLSX_NAME}: {hybrid_dir}"
        )
    derived_path = hybrid_result.derived_registry_path
    _write_state(
        hybrid_dir,
        {
            "algorithm_version": HYBRID_ALGORITHM_VERSION,
            "fingerprint": hybrid_result.fingerprint,
            "source_root": str(source),
            "registry_path": str(registry.path),
            "rfp_root": str(rfp_folder),
            "baseline_fingerprint": baseline_result.fingerprint,
            "stamp": hybrid_result.stamp,
            "registry_mtime_ns": _file_mtime_ns(registry.path),
            "derived_registry_name": registry.path.name,
        },
    )
    _prune_stale_sidecars(hybrid_dir, keep=_hybrid_keep_names(hybrid_result))
    summary = hybrid_result.summary_line()
    freshness = DsHybridFreshness(
        True,
        f"{stale_reason}; {summary}",
        written,
        hybrid_dir,
        fingerprint=hybrid_result.fingerprint,
        summary=summary,
        baseline_freshness=baseline_fresh,
        result=hybrid_result,
        derived_registry_path=derived_path,
    )
    return written, freshness


def _parts_section(config: dict[str, Any] | None) -> dict[str, Any]:
    if not isinstance(config, dict):
        return {}
    parts = config.get("rfp_parts")
    return parts if isinstance(parts, dict) else {}


def _paths_section(config: dict[str, Any] | None) -> dict[str, Any]:
    if not isinstance(config, dict):
        return {}
    paths = config.get("paths")
    return paths if isinstance(paths, dict) else {}


def _gui_ds_paths() -> dict[str, str]:
    """Last DS folder/registry from the center GUI config, if present."""

    try:
        from RFQ.ds_compare.ds_compare_config import (
            load_ds_compare_config,
            normalize_gui_paths,
        )
    except Exception:
        return {}
    try:
        cfg = load_ds_compare_config()
        return dict(normalize_gui_paths(cfg.get("gui_paths")))
    except Exception:
        return {}


def _resolve_source_root(
    source_root: str | Path | None,
    config: dict[str, Any] | None,
) -> Path:
    raw = source_root
    if raw is None or not str(raw).strip():
        raw = _parts_section(config).get("ds_source_dir") or ""
    if not str(raw).strip():
        raw = _gui_ds_paths().get("last_ds_trusted_folder") or ""
    text = str(raw).strip()
    if not text:
        raise FileNotFoundError(
            "Не задана папка закупочных ДС (rfp_parts.ds_source_dir "
            "или last_ds_trusted_folder на вкладке RFP · Сбор частей)."
        )
    path = Path(text)
    if not path.is_dir():
        raise FileNotFoundError(f"Папка закупочных ДС не найдена: {path}")
    return path


def _resolve_registry_path(
    registry_path: str | Path | None,
    config: dict[str, Any] | None,
) -> Path:
    if registry_path is not None and str(registry_path).strip():
        return Path(str(registry_path).strip())
    from RFQ.rfp_parts.ds_registry import DEFAULT_RFP_BASE, resolve_latest_registry

    try:
        latest = resolve_latest_registry(DEFAULT_RFP_BASE)
        if latest.is_file():
            return latest
    except OSError:
        pass
    raw = _parts_section(config).get("ds_registry_path") or ""
    if not str(raw).strip():
        raw = _gui_ds_paths().get("last_ds_registry_file") or ""
    text = str(raw).strip()
    return Path(text) if text else DEFAULT_REGISTRY_PATH


def _resolve_units_inputs(
    *,
    converter: DsUnitsConverter | None,
    google_index: GoogleUnitsIndex | None,
    matrix_path: str | Path | None,
    config: dict[str, Any] | None,
) -> tuple[GoogleUnitsIndex | None, Path | None]:
    matrix = Path(matrix_path) if matrix_path else None
    if matrix is None:
        raw = str(_paths_section(config).get("units_convert_matrix") or "").strip()
        matrix = Path(raw) if raw else None
    if converter is not None:
        return google_index, matrix
    if google_index is not None:
        return google_index, matrix
    from RFQ.rfp_parts.analyze_rfp_parts import _load_google_rows_strict
    from RFQ.units_convert.models import build_google_units_index

    return build_google_units_index(_load_google_rows_strict()), matrix


def _current_baseline_fingerprint(
    source_root: Path,
    registry: DsRegistryDocument,
) -> str:
    files, _skipped = collect_ds_workbooks(
        source_root, registry_path=registry.path
    )
    return _tree_fingerprint(files, registry=registry)


def _file_mtime_ns(path: Path) -> int | None:
    try:
        return path.stat().st_mtime_ns
    except OSError:
        return None


def _derived_registry_beside(
    hybrid_dir: Path,
    *,
    stored: dict[str, Any] | None,
    registry_path: Path,
) -> Path | None:
    names: list[str] = []
    if stored:
        stored_name = str(stored.get("derived_registry_name") or "").strip()
        if stored_name:
            names.append(stored_name)
    names.append(registry_path.name)
    seen: set[str] = set()
    for name in names:
        if not name or name in seen:
            continue
        seen.add(name)
        path = hybrid_dir / name
        try:
            if path.is_file():
                return path
        except OSError:
            continue
    return None


def _load_state(out_dir: Path) -> dict[str, Any] | None:
    path = out_dir / STATE_JSON_NAME
    try:
        if not path.is_file():
            return None
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    return raw if isinstance(raw, dict) else None


def _write_state(out_dir: Path, payload: dict[str, Any]) -> None:
    path = out_dir / STATE_JSON_NAME
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _baseline_stale_reason(
    *,
    baseline_path: Path,
    stored: dict[str, Any] | None,
    fingerprint: str,
    source_root: Path,
    registry_path: Path,
) -> str | None:
    try:
        exists = baseline_path.is_file()
    except OSError:
        exists = False
    if not exists:
        return f"нет {BASELINE_XLSX_NAME}"
    if not stored:
        return "нет снимка fingerprint свода ДС"
    if str(stored.get("algorithm_version") or "") != BASELINE_ALGORITHM_VERSION:
        return "сменилась версия алгоритма свода ДС"
    if str(stored.get("source_root") or "") != str(source_root):
        return "сменилась папка закупочных ДС"
    if str(stored.get("registry_path") or "") != str(registry_path):
        return "сменился путь реестра ДС"
    if str(stored.get("fingerprint") or "") != fingerprint:
        return "изменились файлы ДС или реестр"
    return None


def _hybrid_stale_reason(
    *,
    hybrid_path: Path,
    stored: dict[str, Any] | None,
    fingerprint: str,
    source_root: Path,
    registry_path: Path,
    rfp_root: Path,
    baseline_needs_rebuild: bool,
    registry_mtime_ns: int | None = None,
) -> str | None:
    if baseline_needs_rebuild:
        return "обновился свод ДС"
    try:
        exists = hybrid_path.is_file()
    except OSError:
        exists = False
    if not exists:
        return f"нет {HYBRID_XLSX_NAME}"
    if not stored:
        return "нет снимка fingerprint свода ДС-RFP"
    if str(stored.get("algorithm_version") or "") != HYBRID_ALGORITHM_VERSION:
        return "сменилась версия алгоритма свода ДС-RFP"
    if str(stored.get("source_root") or "") != str(source_root):
        return "сменилась папка закупочных ДС"
    if str(stored.get("registry_path") or "") != str(registry_path):
        return "сменился путь реестра ДС"
    stored_mtime = stored.get("registry_mtime_ns")
    if registry_mtime_ns is not None and stored_mtime not in (None, ""):
        try:
            stored_i = int(stored_mtime)
        except (TypeError, ValueError):
            stored_i = None
        if stored_i is not None and registry_mtime_ns > stored_i:
            return "реестр ДС/УЛ новее свода"
    if str(stored.get("rfp_root") or "") != str(rfp_root):
        return "сменился корень RFP_Зиновьев"
    if str(stored.get("fingerprint") or "") != fingerprint:
        return "изменились корневые RFP, ДС или реестр"
    return None


def _is_sidecar_name(name: str) -> bool:
    folded = name.casefold()
    if folded in _NET_NAMES_CF:
        return False
    if name.startswith("~$"):
        return False
    return any(folded.startswith(prefix) for prefix in _SIDECAR_PREFIXES_CF)


def _baseline_keep_names(result: DsBaselineResult) -> set[str]:
    names = {BASELINE_XLSX_NAME, STATE_JSON_NAME, HYBRID_REPORT_PREFIX}
    for path in (
        result.baseline_path,
        result.structure_report_path,
        result.quality_report_path,
        result.empty_code_report_path,
        result.duplicate_tags_report_path,
        result.tag_mismatch_report_path,
        *result.fractional_log_paths,
    ):
        if path is not None:
            names.add(Path(path).name)
    return names


def _hybrid_keep_names(result: DsRfpHybridResult) -> set[str]:
    names = {HYBRID_XLSX_NAME, STATE_JSON_NAME, HYBRID_REPORT_PREFIX}
    for path in (
        result.hybrid_path,
        result.report_path,
        result.derived_registry_path,
        *result.fractional_log_paths,
    ):
        if path is not None:
            names.add(Path(path).name)
    return names


def _prune_stale_sidecars(out_dir: Path, *, keep: set[str]) -> None:
    try:
        entries = list(out_dir.iterdir())
    except OSError:
        return
    keep_cf = {item.casefold() for item in keep}
    for path in entries:
        try:
            if not path.is_file():
                continue
            if path.name.casefold() in keep_cf:
                continue
            if not _is_sidecar_name(path.name):
                continue
            path.unlink()
        except OSError as exc:
            print(f"Не удалось удалить старый отчёт ДС {path.name}: {exc}")
