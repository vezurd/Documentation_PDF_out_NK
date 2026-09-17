"""Named MTO export rules, JSON targets/pins, and a copy plan.

Qt-free. Does not copy files. The only function that walks a destination
folder is :func:`scan_export_target`; JSON load/save touch ``runtime_dir``
only. Row universe is RD ∪ Google КСБ ИД ∪ «Выдача», including kits with
no files (``no_source``).
"""

from __future__ import annotations

import hashlib
import json
import os
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, replace
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from rd_catalog.config import CatalogConfig
from rd_catalog.db import CatalogDatabase, canonical_mto_pair_ids
from rd_catalog.kits import KitEvent, format_revision, kit_identity_key
from rd_catalog.models import FileKind, FileRecord, ParseStatus, SourceKind, make_path_key
from rd_catalog.overlay import _newest_sort_key, revision_rank
from rd_catalog.parse import (
    issued_package_dir,
    parse_catalog_file,
    record_has_canonical_layout,
)
from rd_catalog.path_actions import path_is_under
from rd_catalog.pipeline import (
    ApprovedContour,
    _PASSED_STAGES,
    _file_kind,
    _file_revision,
    _int_or_none,
    _parsed_file_from_record,
    _record_is_as_build,
    _record_kit_key,
    index_records_by_kit,
    resolve_all_contours,
)
from rd_catalog.robot_mto_sync import (
    _ref_from_record,
    archive_folder_name,
    robot_kit_directory,
)
from rd_catalog.scan import _ROBOT_EXCLUDED_PARTS

MTO_EXPORT_ALGORITHM_VERSION = 1
EXPORT_TARGETS_FILENAME = "mto_export_targets.json"
EXPORT_PINS_FILENAME = "mto_export_pins.json"
EXPORT_STORE_VERSION = 1
DEFAULT_ROBOT_TARGET_NAME = "Штатная папка робота"
DEFAULT_EXPORT_RULE = "latest_issued"
PIN_COLUMN_HEADER = "Ручной выбор MTO"
PIN_STALE_SUFFIX = " (устарел)"
EXPORT_RULES = frozenset(
    {"approved", "latest_tdo", "latest_issued", "latest_no_as_build"}
)
_ORIGIN_RULE = "rule"
_ORIGIN_PIN = "pin"
_ORIGIN_PIN_STALE = "pin_stale"
_STATE_SAME_DATA = "same_data"
_STATE_REPLACE = "replace"
_STATE_ADD = "add"
_STATE_NO_SOURCE = "no_source"
_STATE_PIN_STALE = "pin_stale"
_STATE_UNKNOWN = "unknown"
_SKIP_PLAN_STATES = frozenset(
    {
        _STATE_NO_SOURCE,
        _STATE_SAME_DATA,
        _STATE_UNKNOWN,
        _STATE_PIN_STALE,
    }
)


@dataclass(frozen=True, slots=True)
class ExportTarget:
    """One named destination folder for bulk MTO export."""

    name: str
    root: str
    flat_structure: bool
    is_default_robot: bool
    rule: str
    filter_text: str


@dataclass(frozen=True, slots=True)
class ExportPin:
    """User override: take this kit's MTO from a chosen package file."""

    title: str
    mark: str
    package_path: str
    file_path: str
    revision_text: str
    evidence: str
    created_at: str


@dataclass(frozen=True, slots=True)
class ExportPinCandidate:
    """One MTO file the user may pin, built from catalog records only."""

    package_path: str
    package_name: str
    sequence: int | None
    file_path: str
    file_name: str
    revision_text: str
    size: int
    is_rule_choice: bool


@dataclass(frozen=True, slots=True)
class ExportPinView:
    """Display payload for the «Ручной выбор MTO» column."""

    text: str
    tooltip: str
    is_stale: bool


@dataclass(frozen=True, slots=True)
class ExportSelection:
    """One kit row of an export preview.

    ``source_path`` is the file that would actually be copied (the pin,
    when a pin is honoured or stale). ``rule_path`` is what the named
    rule computes now; it differs from ``source_path`` when a pin
    overrides the rule (``origin`` is ``pin`` or ``pin_stale``).
    """

    title: str
    mark: str
    rule: str
    source_path: str
    source_revision_text: str
    package_path: str
    origin: str
    state: str
    destination_path: str
    existing_target_path: str
    confidence: str
    warnings: tuple[str, ...]
    rule_path: str = ""


@dataclass(frozen=True, slots=True)
class ExportPlanItem:
    """One copy/replace step of an export plan (no filesystem writes)."""

    selection: ExportSelection
    archive_dir: str


@dataclass(frozen=True, slots=True)
class ExportPlan:
    """Copy plan for one target.

    Skipped rows are ``no_source``, ``same_data``, ``unknown``, and
    ``pin_stale``. They never become copy items.
    """

    target: ExportTarget
    items: tuple[ExportPlanItem, ...]
    skipped: tuple[ExportSelection, ...]


def pin_evidence(
    events: Sequence[KitEvent],
    mto_files: Sequence[FileRecord],
) -> str:
    """Return a stable fingerprint of journal + candidate MTO facts.

    Hashes parsed F fields only (never ``raw``), then present canonical RD
    MTO file facts, then :data:`MTO_EXPORT_ALGORITHM_VERSION`.

    Args:
        events: Column-F events in sheet order.
        mto_files: Present canonical RD ``mto_xlsx`` records of the kit.

    Returns:
        Hex SHA-256 digest.
    """

    event_rows = [
        (
            str(event.date or ""),
            str(event.stage or ""),
            str(event.revision or ""),
            str(event.appendix or ""),
        )
        for event in events
    ]
    file_rows = sorted(
        _candidate_file_fact(record) for record in mto_files
    )
    payload = {
        "algorithm_version": MTO_EXPORT_ALGORITHM_VERSION,
        "events": event_rows,
        "files": file_rows,
    }
    encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def pin_is_current(pin: ExportPin, evidence: str) -> bool:
    """Return whether ``pin`` still matches ``evidence``.

    Args:
        pin: Stored pin.
        evidence: Fresh fingerprint from :func:`pin_evidence`.

    Returns:
        True when the stored digest equals ``evidence``.
    """

    return bool(pin.evidence) and pin.evidence == evidence


def reconcile_export_pin(
    pin: ExportPin,
    *,
    evidence: str,
    rule_path: str,
) -> tuple[str, ExportPin]:
    """Apply pin invalidation without deleting the pin.

    When evidence is stale but the rule still picks ``pin.file_path``,
    the pin is silently refreshed (new evidence, origin stays ``pin``).
    When the rule picks a different path, origin is ``pin_stale``.

    Args:
        pin: Stored pin.
        evidence: Fresh fingerprint.
        rule_path: Path the named rule resolved, or ``""``.

    Returns:
        ``(origin, pin)``. The returned pin has updated evidence on a
        silent refresh; otherwise it is ``pin`` unchanged.
    """

    if pin_is_current(pin, evidence):
        return _ORIGIN_PIN, pin
    if _paths_equal(rule_path, pin.file_path) and pin.file_path:
        return _ORIGIN_PIN, replace(pin, evidence=evidence)
    return _ORIGIN_PIN_STALE, pin


def export_pin_view(
    pin: ExportPin | None,
    *,
    origin: str = "",
    rule_path: str = "",
) -> ExportPinView:
    """Return table text, tooltip, and stale flag for one kit pin.

    Empty when ``pin`` is missing. Stale marker and the rule's current
    file appear only when ``origin`` is ``pin_stale``.

    Args:
        pin: Stored pin, or ``None``.
        origin: ``pin`` / ``pin_stale`` / ``rule`` / empty.
        rule_path: Path the named rule resolves now (stale tooltip).

    Returns:
        Display payload. ``text`` is empty when there is no pin.
    """

    if pin is None:
        return ExportPinView(text="", tooltip="", is_stale=False)
    name = Path(pin.file_path).name if pin.file_path else ""
    stale = origin == _ORIGIN_PIN_STALE
    text = f"{name}{PIN_STALE_SUFFIX}" if stale and name else name
    lines = [
        f"пакет: {pin.package_path or '—'}",
        f"файл: {pin.file_path or '—'}",
        f"создан: {pin.created_at or '—'}",
    ]
    if stale:
        lines.append(f"файл по правилу: {rule_path or '—'}")
    return ExportPinView(text=text, tooltip="\n".join(lines), is_stale=stale)


def list_export_pin_candidates(
    records: Sequence[FileRecord],
    *,
    title: str,
    mark: str,
    rd_root: str | Path,
    rule_path: str = "",
) -> tuple[ExportPinCandidate, ...]:
    """List every MTO file the user may pin for one kit.

    Uses catalog records only (no workbook IO). Present canonical RD
    ``mto_xlsx`` rows are included, plus present SQ MTO of the same kit
    (``latest_issued`` may fall back to SQ). Newest transfer sequence
    first. ``is_rule_choice`` marks the path the current rule picked.

    Args:
        records: Catalog file records from the last scan.
        title: Four-digit title.
        mark: Latin AGCC mark.
        rd_root: RD source root for the canonical-layout filter.
        rule_path: Path the named rule resolved, or ``""``.

    Returns:
        Candidates sorted newest transfer first.
    """

    key = kit_identity_key(title, mark)
    found: list[FileRecord] = []
    for record in records:
        if not record.present:
            continue
        if _file_kind(record) != FileKind.MTO_XLSX.value:
            continue
        if _record_kit_key(record) != key:
            continue
        if record.source is SourceKind.RD:
            if not record_has_canonical_layout(record, rd_root):
                continue
        elif record.source is not SourceKind.SQ:
            continue
        found.append(record)
    found.sort(key=_overlay_rank, reverse=True)
    rule_key = make_path_key(rule_path) if rule_path else ""
    candidates: list[ExportPinCandidate] = []
    for record in found:
        package_path = issued_package_dir(record.path) or str(
            Path(record.path).parent
        )
        revision, appendix = _file_revision(record)
        sequence = _int_or_none(record.data.get("transfer_sequence"))
        file_name = str(record.data.get("name") or Path(record.path).name)
        candidates.append(
            ExportPinCandidate(
                package_path=package_path,
                package_name=Path(package_path).name if package_path else "",
                sequence=sequence,
                file_path=record.path,
                file_name=file_name,
                revision_text=format_revision(revision, appendix),
                size=_int_or_none(record.data.get("size")) or 0,
                is_rule_choice=bool(
                    rule_key and make_path_key(record.path) == rule_key
                ),
            )
        )
    return tuple(candidates)


def export_pin_evidence_for_kit(
    database: CatalogDatabase,
    records: Sequence[FileRecord],
    *,
    title: str,
    mark: str,
    rd_root: str | Path,
) -> str:
    """Return :func:`pin_evidence` for one kit from journal + RD MTO facts.

    Args:
        database: Initialized catalog database.
        records: Catalog file records from the last scan.
        title: Four-digit title.
        mark: Latin AGCC mark.
        rd_root: RD source root for the canonical-layout filter.

    Returns:
        Hex SHA-256 digest.
    """

    key = kit_identity_key(title, mark)
    events = database.list_google_events(title, mark)
    mto_files = _canonical_rd_mto_records(records, key, rd_root)
    return pin_evidence(events, mto_files)


def default_robot_export_target(config: CatalogConfig) -> ExportTarget:
    """Return the synthesized default robot target from ``config``.

    Args:
        config: Catalog configuration.

    Returns:
        Target whose root and flat layout follow ``robot_root`` /
        ``robot_flat_structure``.
    """

    return ExportTarget(
        name=DEFAULT_ROBOT_TARGET_NAME,
        root=str(config.robot_root),
        flat_structure=bool(config.robot_flat_structure),
        is_default_robot=True,
        rule=DEFAULT_EXPORT_RULE,
        filter_text="",
    )


def load_export_targets(config: CatalogConfig) -> tuple[ExportTarget, ...]:
    """Load named targets from ``runtime_dir``, synthesizing the default.

    A missing or corrupt file yields only the default robot target. The
    default always uses ``config.robot_root`` / ``robot_flat_structure``.

    Args:
        config: Catalog configuration (``runtime_dir`` and robot folder).

    Returns:
        Targets with the default robot target first.
    """

    default = default_robot_export_target(config)
    path = Path(config.runtime_dir) / EXPORT_TARGETS_FILENAME
    raw = _read_json_object(path)
    if raw is None:
        return (default,)
    items = raw.get("targets")
    if not isinstance(items, list):
        return (default,)
    merged: list[ExportTarget] = []
    seen_default = False
    seen_names: set[str] = set()
    for item in items:
        target = _target_from_json(item)
        if target is None:
            continue
        if target.is_default_robot:
            if seen_default:
                continue
            seen_default = True
            name = target.name.strip() or default.name
            rule = target.rule if target.rule in EXPORT_RULES else default.rule
            target = replace(
                default,
                name=name,
                rule=rule,
                filter_text=target.filter_text,
            )
        name_key = target.name.casefold()
        if name_key in seen_names:
            continue
        seen_names.add(name_key)
        merged.append(target)
    if not seen_default:
        merged.insert(0, default)
    else:
        default_index = next(
            index
            for index, item in enumerate(merged)
            if item.is_default_robot
        )
        if default_index != 0:
            merged.insert(0, merged.pop(default_index))
    return tuple(merged)


def save_export_targets(
    config: CatalogConfig, targets: Sequence[ExportTarget]
) -> tuple[ExportTarget, ...]:
    """Write targets atomically. The default robot target cannot be removed.

    The default target's root and flat layout are always taken from
    ``config``. Custom targets may not reuse its name.

    Args:
        config: Catalog configuration.
        targets: Targets to persist (default may be omitted; it is
            re-injected).

    Returns:
        The stored tuple (default first).
    """

    default = default_robot_export_target(config)
    stored: list[ExportTarget] = []
    seen_names: set[str] = set()
    seen_default = False
    for target in targets:
        if target.is_default_robot:
            if seen_default:
                continue
            seen_default = True
            name = target.name.strip() or default.name
            rule = target.rule if target.rule in EXPORT_RULES else default.rule
            stored.append(
                replace(
                    default,
                    name=name,
                    rule=rule,
                    filter_text=str(target.filter_text or ""),
                )
            )
            seen_names.add(stored[-1].name.casefold())
            continue
        name = target.name.strip()
        if not name or name.casefold() in seen_names:
            continue
        if name.casefold() == default.name.casefold():
            continue
        root = str(target.root or "").strip()
        if not root:
            continue
        rule = target.rule if target.rule in EXPORT_RULES else DEFAULT_EXPORT_RULE
        seen_names.add(name.casefold())
        stored.append(
            ExportTarget(
                name=name,
                root=root,
                flat_structure=bool(target.flat_structure),
                is_default_robot=False,
                rule=rule,
                filter_text=str(target.filter_text or ""),
            )
        )
    if not seen_default:
        stored.insert(0, default)
    payload = {
        "version": EXPORT_STORE_VERSION,
        "targets": [_target_to_json(item) for item in stored],
    }
    _atomic_write_json(Path(config.runtime_dir) / EXPORT_TARGETS_FILENAME, payload)
    return tuple(stored)


def load_export_pins(config: CatalogConfig) -> tuple[ExportPin, ...]:
    """Load kit pins from ``runtime_dir``. Missing or corrupt → empty.

    Args:
        config: Catalog configuration.

    Returns:
        Pins in file order (duplicates by kit identity: last wins).
    """

    path = Path(config.runtime_dir) / EXPORT_PINS_FILENAME
    raw = _read_json_object(path)
    if raw is None:
        return ()
    items = raw.get("pins")
    if not isinstance(items, list):
        return ()
    by_key: dict[tuple[str, str], ExportPin] = {}
    for item in items:
        pin = _pin_from_json(item)
        if pin is None:
            continue
        by_key[kit_identity_key(pin.title, pin.mark)] = pin
    return tuple(by_key.values())


def save_export_pins(
    config: CatalogConfig, pins: Sequence[ExportPin]
) -> tuple[ExportPin, ...]:
    """Write pins atomically. Duplicate kit identities: last wins.

    Pins are never deleted here; the caller passes the list to keep.

    Args:
        config: Catalog configuration.
        pins: Pins to persist.

    Returns:
        The stored tuple.
    """

    by_key: dict[tuple[str, str], ExportPin] = {}
    for pin in pins:
        title = str(pin.title or "").strip()
        mark = str(pin.mark or "").strip()
        if not title or not mark or not str(pin.file_path or "").strip():
            continue
        stored = ExportPin(
            title=title,
            mark=mark,
            package_path=str(pin.package_path or ""),
            file_path=str(pin.file_path or "").strip(),
            revision_text=str(pin.revision_text or ""),
            evidence=str(pin.evidence or ""),
            created_at=str(pin.created_at or "").strip() or _utc_now(),
        )
        by_key[kit_identity_key(title, mark)] = stored
    stored_tuple = tuple(by_key.values())
    payload = {
        "version": EXPORT_STORE_VERSION,
        "pins": [_pin_to_json(pin) for pin in stored_tuple],
    }
    _atomic_write_json(Path(config.runtime_dir) / EXPORT_PINS_FILENAME, payload)
    return stored_tuple


def upsert_export_pin(config: CatalogConfig, pin: ExportPin) -> tuple[ExportPin, ...]:
    """Insert or replace one kit pin. Never deletes other pins.

    Args:
        config: Catalog configuration (``runtime_dir``).
        pin: Pin to store.

    Returns:
        The stored tuple after the write.
    """

    by_key = {
        kit_identity_key(item.title, item.mark): item
        for item in load_export_pins(config)
    }
    by_key[kit_identity_key(pin.title, pin.mark)] = pin
    return save_export_pins(config, tuple(by_key.values()))


def remove_export_pin(
    config: CatalogConfig, title: str, mark: str
) -> tuple[ExportPin, ...]:
    """Drop the pin for one kit. Other pins are kept.

    Args:
        config: Catalog configuration (``runtime_dir``).
        title: Four-digit title.
        mark: Latin AGCC mark.

    Returns:
        The stored tuple after the write.
    """

    key = kit_identity_key(title, mark)
    kept = [
        item
        for item in load_export_pins(config)
        if kit_identity_key(item.title, item.mark) != key
    ]
    return save_export_pins(config, kept)


def target_files_from_records(
    records: Sequence[FileRecord],
) -> dict[tuple[str, str], str]:
    """Map kit identity to the newest present robot MTO path.

    Lets the default-robot caller reuse an already-scanned robot snapshot
    instead of walking the folder again.

    Args:
        records: Catalog file records (robot rows are used).

    Returns:
        ``kit_identity_key`` → path. One file per kit (highest revision,
        then later mtime).
    """

    best: dict[tuple[str, str], tuple[tuple, str]] = {}
    for record in records:
        if record.source is not SourceKind.ROBOT:
            continue
        ref = _ref_from_record(record)
        if ref is None:
            continue
        key = kit_identity_key(ref.title, ref.mark)
        rank = _freshness_rank(record)
        previous = best.get(key)
        if previous is None or rank > previous[0]:
            best[key] = (rank, ref.path)
    return {key: path for key, (_rank, path) in best.items()}


def scan_export_target(target: ExportTarget) -> dict[tuple[str, str], str]:
    """Walk ``target.root`` once and map kit identity → existing MTO path.

    Missing or unreachable roots return ``{}`` and never raise. Archive
    folders (substring ``old``, ``archive``, …) are pruned like the robot
    scanner. This is the only function in the module that stats destination
    files.

    Args:
        target: Destination to scan.

    Returns:
        ``kit_identity_key`` → path. One file per kit.
    """

    root_text = str(target.root or "").strip()
    if not root_text:
        return {}
    root = Path(root_text)
    try:
        if not root.is_dir():
            return {}
    except OSError:
        return {}

    best: dict[tuple[str, str], tuple[tuple, str]] = {}

    def on_walk_error(_error: OSError) -> None:
        return None

    try:
        walker = os.walk(root_text, onerror=on_walk_error)
        for current_root, dirs, names in walker:
            dirs[:] = [
                directory
                for directory in dirs
                if not any(
                    excluded in directory.casefold()
                    for excluded in _ROBOT_EXCLUDED_PARTS
                )
            ]
            for name in names:
                folded = name.casefold()
                if not folded.endswith(".xlsx"):
                    continue
                if "mto" not in folded and "мто" not in folded:
                    continue
                path = os.path.join(current_root, name)
                try:
                    stat = os.stat(path, follow_symlinks=False)
                    parsed = _parse_target_xlsx(path, stat)
                except (OSError, ValueError):
                    continue
                if parsed is None:
                    continue
                key = kit_identity_key(parsed[0], parsed[1])
                rank = parsed[2]
                previous = best.get(key)
                if previous is None or rank > previous[0]:
                    best[key] = (rank, path)
    except OSError:
        return {}
    return {key: path for key, (_rank, path) in best.items()}


def resolve_export_selections(
    database: CatalogDatabase,
    *,
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    rule: str,
    target: ExportTarget,
    target_files: Mapping[tuple[str, str], str],
    pins: Sequence[ExportPin],
    rd_root: str | Path,
    verdicts: Mapping[tuple[int, int], str] | None = None,
    refreshed_pins: list[ExportPin] | None = None,
) -> tuple[ExportSelection, ...]:
    """Resolve one MTO file (or ``no_source``) for every known catalog kit.

    Does not read XLSX. When ``verdicts`` is ``None``, every selection that
    has both a source and an existing destination file is ``unknown``.
    Pins are never deleted. Silent evidence refreshes from
    :func:`reconcile_export_pin` are appended to ``refreshed_pins`` when
    that list is provided; the caller persists them.

    Args:
        database: Initialized catalog database.
        records: Catalog file records from the last scan.
        detected_current_ids: Overlay-current RD file ids.
        rule: One of :data:`EXPORT_RULES`.
        target: Destination (layout + root).
        target_files: Kit identity → existing file in the target. Keys may
            be raw ``(title, mark)`` or :func:`kit_identity_key`.
        pins: User overrides; never deleted here.
        rd_root: RD source root for the canonical-layout filter.
        verdicts: Optional ``(file_id, file_id)`` → ``content_equal`` /
            ``content_diff`` / ``not_compared`` from one
            ``list_mto_pair_comparisons`` query. Keys may be in either order.
        refreshed_pins: Optional sink for pins whose evidence was silently
            refreshed because the rule still picks the same file.

    Returns:
        One selection per kit, sorted by title then mark.

    Raises:
        ValueError: If ``rule`` is not a known export rule.
    """

    if rule not in EXPORT_RULES:
        raise ValueError(f"Unknown MTO export rule: {rule!r}")

    materialized = list(records)
    files_by_path = {
        make_path_key(record.path): record for record in materialized
    }
    target_map = {
        kit_identity_key(title, mark): path
        for (title, mark), path in target_files.items()
        if str(path or "").strip()
    }
    pin_map = {
        kit_identity_key(pin.title, pin.mark): pin for pin in pins
    }
    events_by_kit = database.list_google_events_by_kit()
    kits = _kit_universe(database, materialized, rd_root)
    rule_hits = _resolve_rule_hits(
        database,
        records=materialized,
        detected_current_ids=detected_current_ids,
        rule=rule,
        rd_root=rd_root,
        kits=kits,
    )
    mto_by_kit = _index_canonical_rd_mto(materialized, rd_root)
    selections: list[ExportSelection] = []
    for key in sorted(kits):
        title, mark = kits[key]
        hit = rule_hits.get(key)
        rule_path = hit.path if hit is not None else ""
        events = [event for _event_id, event in events_by_kit.get(key, ())]
        mto_files = mto_by_kit.get(key, [])
        evidence = pin_evidence(events, mto_files)
        pin = pin_map.get(key)
        origin = _ORIGIN_RULE
        source_path = rule_path
        warnings: list[str] = []
        confidence = hit.confidence if hit is not None else "low"
        package_path = hit.package_path if hit is not None else ""
        source_revision = hit.revision_text if hit is not None else ""
        if hit is not None:
            warnings.extend(hit.warnings)
        if pin is not None:
            stored = pin
            origin, pin = reconcile_export_pin(
                pin, evidence=evidence, rule_path=rule_path
            )
            if (
                refreshed_pins is not None
                and origin == _ORIGIN_PIN
                and pin is not stored
            ):
                refreshed_pins.append(pin)
            source_path = pin.file_path
            package_path = pin.package_path or package_path
            source_revision = _revision_for_path(
                pin.file_path, files_by_path, fallback=pin.revision_text
            )
            if origin == _ORIGIN_PIN:
                if not confidence or confidence == "low":
                    confidence = "high"
            else:
                warnings.append(
                    "Закреплённый файл больше не совпадает с правилом. "
                    f"Было: {pin.file_path}. "
                    f"Правило сейчас: {rule_path or 'нет файла'}."
                )
        existing = _lookup_target_file(target_map, key)
        destination_path = _destination_path(
            target,
            title=title,
            mark=mark,
            source_path=source_path,
            existing_target_path=existing,
        )
        state = _selection_state(
            origin=origin,
            source_path=source_path,
            existing_path=existing,
            files_by_path=files_by_path,
            verdict=_lookup_verdict(
                verdicts, source_path, existing, files_by_path
            ),
        )
        mismatch = _same_name_content_warning(
            source_path, existing, files_by_path
        )
        if mismatch:
            warnings.append(mismatch)
        selections.append(
            ExportSelection(
                title=title,
                mark=mark,
                rule=rule,
                source_path=source_path if source_path else "",
                source_revision_text=source_revision,
                package_path=package_path,
                origin=origin,
                state=state,
                destination_path=destination_path if source_path else "",
                existing_target_path=existing,
                confidence=confidence if source_path else "low",
                warnings=tuple(warnings),
                rule_path=rule_path,
            )
        )
    return tuple(selections)


def build_export_plan(
    selections: Sequence[ExportSelection],
    *,
    target: ExportTarget,
) -> ExportPlan:
    """Build a copy/replace plan. Does not write files.

    Destinations that do not sit under ``target.root`` raise — that is a
    bug, not a warning. ``no_source``, ``same_data``, ``unknown``, and
    ``pin_stale`` rows are skipped. ``replace`` and ``add`` become items.

    Args:
        selections: Rows from :func:`resolve_export_selections`.
        target: Destination whose root bounds every path.

    Returns:
        Plan with copy items and skipped rows.

    Raises:
        ValueError: If an item destination escapes ``target.root``.
    """

    items: list[ExportPlanItem] = []
    skipped: list[ExportSelection] = []
    for selection in selections:
        if selection.state in _SKIP_PLAN_STATES:
            skipped.append(selection)
            continue
        destination = selection.destination_path
        if not destination:
            skipped.append(selection)
            continue
        if not path_is_under(destination, target.root):
            raise ValueError(
                "Export destination escapes target root: "
                f"{destination!r} vs {target.root!r}"
            )
        archive_dir = ""
        if (
            selection.state == _STATE_REPLACE
            and selection.existing_target_path
        ):
            existing = Path(selection.existing_target_path)
            archive_dir = str(
                existing.parent
                / archive_folder_name(existing.stem, datetime.now())
            )
            if not path_is_under(archive_dir, target.root):
                raise ValueError(
                    "Export archive path escapes target root: "
                    f"{archive_dir!r} vs {target.root!r}"
                )
        items.append(ExportPlanItem(selection=selection, archive_dir=archive_dir))
    return ExportPlan(target=target, items=tuple(items), skipped=tuple(skipped))


@dataclass(frozen=True, slots=True)
class _RuleHit:
    path: str
    revision_text: str
    package_path: str
    confidence: str
    warnings: tuple[str, ...]


def _resolve_rule_hits(
    database: CatalogDatabase,
    *,
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    rule: str,
    rd_root: str | Path,
    kits: Mapping[tuple[str, str], tuple[str, str]],
) -> dict[tuple[str, str], _RuleHit]:
    if rule in {"approved", "latest_tdo"}:
        stages = _PASSED_STAGES if rule == "latest_tdo" else None
        contours = resolve_all_contours(
            database,
            records=records,
            detected_current_ids=detected_current_ids,
            rd_root=rd_root,
            anchor_stages=stages,
        )
        return {
            kit_identity_key(contour.title, contour.mark): _hit_from_contour(
                contour
            )
            for contour in contours
            if contour.mto_path
        }
    if rule == "latest_issued":
        return _latest_issued_hits(records, detected_current_ids, rd_root, kits)
    return _latest_no_as_build_hits(records, rd_root, kits)


def _hit_from_contour(contour: ApprovedContour) -> _RuleHit:
    return _RuleHit(
        path=contour.mto_path,
        revision_text=_revision_from_path_name(
            contour.mto_path, fallback=contour.approved_revision_text
        ),
        package_path=contour.package_path,
        confidence=contour.confidence or "low",
        warnings=contour.warnings,
    )


def _latest_issued_hits(
    records: Sequence[FileRecord],
    detected_current_ids: set[int],
    rd_root: str | Path,
    kits: Mapping[tuple[str, str], tuple[str, str]],
) -> dict[tuple[str, str], _RuleHit]:
    rd_best: dict[tuple[str, str], FileRecord] = {}
    sq_best: dict[tuple[str, str], FileRecord] = {}
    for record in records:
        if not record.present:
            continue
        if not record_has_canonical_layout(record, rd_root):
            continue
        key = _record_kit_key(record)
        if key is None or key not in kits:
            continue
        if _file_kind(record) != FileKind.MTO_XLSX.value:
            continue
        if record.source is SourceKind.RD:
            if record.id not in detected_current_ids:
                continue
            previous = rd_best.get(key)
            if previous is None or _overlay_rank(record) > _overlay_rank(previous):
                rd_best[key] = record
        elif record.source is SourceKind.SQ:
            previous = sq_best.get(key)
            if previous is None or _freshness_rank(record) > _freshness_rank(
                previous
            ):
                sq_best[key] = record
    hits: dict[tuple[str, str], _RuleHit] = {}
    for key in kits:
        record = rd_best.get(key) or sq_best.get(key)
        if record is None:
            continue
        hits[key] = _hit_from_record(record, confidence="high")
    return hits


def _latest_no_as_build_hits(
    records: Sequence[FileRecord],
    rd_root: str | Path,
    kits: Mapping[tuple[str, str], tuple[str, str]],
) -> dict[tuple[str, str], _RuleHit]:
    best: dict[tuple[str, str], FileRecord] = {}
    for record in records:
        if not record.present or record.source is not SourceKind.RD:
            continue
        if not record_has_canonical_layout(record, rd_root):
            continue
        if _file_kind(record) != FileKind.MTO_XLSX.value:
            continue
        if _record_is_as_build(record):
            continue
        key = _record_kit_key(record)
        if key is None or key not in kits:
            continue
        previous = best.get(key)
        if previous is None or _filename_rev_rank(record) > _filename_rev_rank(
            previous
        ):
            best[key] = record
    return {key: _hit_from_record(record, confidence="high") for key, record in best.items()}


def _hit_from_record(record: FileRecord, *, confidence: str) -> _RuleHit:
    revision, appendix = _file_revision(record)
    return _RuleHit(
        path=record.path,
        revision_text=format_revision(revision, appendix),
        package_path=issued_package_dir(record.path) or str(Path(record.path).parent),
        confidence=confidence,
        warnings=(),
    )


def _kit_universe(
    database: CatalogDatabase,
    records: Sequence[FileRecord],
    rd_root: str | Path,
) -> dict[tuple[str, str], tuple[str, str]]:
    names: dict[tuple[str, str], tuple[str, str]] = {}
    for record in records:
        if not record.present:
            continue
        if record.source is not SourceKind.RD:
            continue
        if not record_has_canonical_layout(record, rd_root):
            continue
        identity = _record_kit_key(record)
        if identity is None:
            continue
        title = str(record.data.get("title") or "").strip()
        mark = str(record.data.get("mark") or "").strip()
        names.setdefault(identity, (title, mark))
    for send in database.list_issuance_sends():
        key = kit_identity_key(send.title, send.mark)
        names[key] = (send.title, send.mark)
    for kit in database.list_google_kits():
        key = kit_identity_key(kit.title, kit.mark)
        names[key] = (kit.title, kit.mark)
    return names


def _index_canonical_rd_mto(
    records: Sequence[FileRecord],
    rd_root: str | Path,
) -> dict[tuple[str, str], list[FileRecord]]:
    matching = [
        record
        for record in records
        if record.present
        and record.source is SourceKind.RD
        and _file_kind(record) == FileKind.MTO_XLSX.value
        and record_has_canonical_layout(record, rd_root)
    ]
    return index_records_by_kit(matching)


def _canonical_rd_mto_records(
    records: Sequence[FileRecord],
    key: tuple[str, str],
    rd_root: str | Path,
) -> list[FileRecord]:
    return _index_canonical_rd_mto(records, rd_root).get(key, [])


def _candidate_file_fact(record: FileRecord) -> tuple[str, int, str, int, int]:
    package_path = issued_package_dir(record.path) or ""
    sequence = _int_or_none(record.data.get("transfer_sequence"))
    name = str(record.data.get("name") or Path(record.path).name)
    size = _int_or_none(record.data.get("size")) or 0
    mtime_ns = _int_or_none(record.data.get("mtime_ns")) or 0
    return (package_path, sequence if sequence is not None else -1, name, size, mtime_ns)


def _selection_state(
    *,
    origin: str,
    source_path: str,
    existing_path: str,
    files_by_path: Mapping[str, FileRecord],
    verdict: str | None = None,
) -> str:
    if origin == _ORIGIN_PIN_STALE:
        return _STATE_PIN_STALE
    if not source_path:
        return _STATE_NO_SOURCE
    if not existing_path:
        return _STATE_ADD
    if verdict == "content_equal":
        return _STATE_SAME_DATA
    if verdict == "content_diff":
        return _STATE_REPLACE
    return _STATE_UNKNOWN


def _lookup_verdict(
    verdicts: Mapping[tuple[int, int], str] | None,
    source_path: str,
    existing_path: str,
    files_by_path: Mapping[str, FileRecord],
) -> str | None:
    if verdicts is None or not source_path or not existing_path:
        return None
    source = files_by_path.get(make_path_key(source_path))
    existing = files_by_path.get(make_path_key(existing_path))
    if source is None or existing is None:
        return None
    key = canonical_mto_pair_ids(source.id, existing.id)
    if key in verdicts:
        return verdicts[key]
    return verdicts.get((source.id, existing.id)) or verdicts.get(
        (existing.id, source.id)
    )


def _same_target_file(
    source_path: str,
    existing_path: str,
    files_by_path: Mapping[str, FileRecord],
) -> bool:
    if Path(source_path).name.casefold() != Path(existing_path).name.casefold():
        return False
    source = files_by_path.get(make_path_key(source_path))
    existing = files_by_path.get(make_path_key(existing_path))
    if source is None or existing is None:
        return True
    return (
        (_int_or_none(source.data.get("size")) or 0)
        == (_int_or_none(existing.data.get("size")) or 0)
        and (_int_or_none(source.data.get("mtime_ns")) or 0)
        == (_int_or_none(existing.data.get("mtime_ns")) or 0)
    )


def _file_size(record: FileRecord | None) -> int:
    if record is None:
        return 0
    return _int_or_none(record.data.get("size")) or 0


def _same_name_content_warning(
    source_path: str,
    existing_path: str,
    files_by_path: Mapping[str, FileRecord],
) -> str | None:
    """Russian warning when dest has the same name but different stats."""

    if not source_path or not existing_path:
        return None
    if Path(source_path).name.casefold() != Path(existing_path).name.casefold():
        return None
    if _same_target_file(source_path, existing_path, files_by_path):
        return None
    source = files_by_path.get(make_path_key(source_path))
    existing = files_by_path.get(make_path_key(existing_path))
    source_size = _file_size(source)
    existing_size = _file_size(existing)
    return (
        "В папке назначения уже есть файл с тем же именем, но другим "
        "содержимым "
        f"(источник {source_size} байт, у робота {existing_size} байт)."
    )


def _destination_path(
    target: ExportTarget,
    *,
    title: str,
    mark: str,
    source_path: str,
    existing_target_path: str,
) -> str:
    if not source_path:
        return ""
    destination_dir = robot_kit_directory(
        target.root,
        title,
        mark,
        flat_structure=target.flat_structure,
        existing_robot_path=existing_target_path or None,
    )
    return str(destination_dir / Path(source_path).name)


def _lookup_target_file(
    target_map: Mapping[tuple[str, str], str],
    key: tuple[str, str],
) -> str:
    if key in target_map:
        return target_map[key]
    return ""


def _revision_for_path(
    path: str,
    files_by_path: Mapping[str, FileRecord],
    *,
    fallback: str,
) -> str:
    record = files_by_path.get(make_path_key(path))
    if record is None:
        return fallback
    return format_revision(*_file_revision(record)) or fallback


def _revision_from_path_name(path: str, *, fallback: str) -> str:
    try:
        parsed = parse_catalog_file(path, SourceKind.RD, size=0, mtime_ns=0)
    except ValueError:
        return fallback
    text = format_revision(parsed.revision, parsed.appendix)
    return text or fallback


def _parse_target_xlsx(
    path: str, stat: os.stat_result
) -> tuple[str, str, tuple] | None:
    parsed = parse_catalog_file(
        path,
        SourceKind.ROBOT,
        size=int(stat.st_size),
        mtime_ns=int(stat.st_mtime_ns),
    )
    if parsed.parse_status is not ParseStatus.PARSED:
        return None
    if parsed.file_kind is not FileKind.MTO_XLSX:
        return None
    if not parsed.title or not parsed.mark:
        return None
    rank = (
        revision_rank(parsed.revision, parsed.appendix),
        int(parsed.mtime_ns or 0),
        parsed.path_key,
    )
    return parsed.title, parsed.mark, rank


def _overlay_rank(record: FileRecord) -> tuple:
    parsed = _parsed_file_from_record(record)
    if parsed is None:
        return (-1, revision_rank(None, None), 0, record.path_key)
    return _newest_sort_key(parsed)


def _freshness_rank(record: FileRecord) -> tuple:
    revision, appendix = _file_revision(record)
    return (
        revision_rank(revision, appendix),
        _int_or_none(record.data.get("mtime_ns")) or 0,
        record.path_key,
    )


def _filename_rev_rank(record: FileRecord) -> tuple:
    revision, appendix = _file_revision(record)
    sequence = _int_or_none(record.data.get("transfer_sequence"))
    return (
        revision_rank(revision, appendix),
        sequence if sequence is not None else -1,
        _int_or_none(record.data.get("mtime_ns")) or 0,
        record.path_key,
    )


def _paths_equal(left: str, right: str) -> bool:
    if not left or not right:
        return False
    return make_path_key(left) == make_path_key(right)


def _target_from_json(item: Any) -> ExportTarget | None:
    if not isinstance(item, dict):
        return None
    name = str(item.get("name") or "").strip()
    root = str(item.get("root") or "").strip()
    if not name:
        return None
    rule = str(item.get("rule") or DEFAULT_EXPORT_RULE).strip()
    if rule not in EXPORT_RULES:
        rule = DEFAULT_EXPORT_RULE
    return ExportTarget(
        name=name,
        root=root,
        flat_structure=bool(item.get("flat_structure")),
        is_default_robot=bool(item.get("is_default_robot")),
        rule=rule,
        filter_text=str(item.get("filter_text") or ""),
    )


def _target_to_json(target: ExportTarget) -> dict[str, Any]:
    return {
        "name": target.name,
        "root": target.root,
        "flat_structure": target.flat_structure,
        "is_default_robot": target.is_default_robot,
        "rule": target.rule,
        "filter_text": target.filter_text,
    }


def _pin_from_json(item: Any) -> ExportPin | None:
    if not isinstance(item, dict):
        return None
    title = str(item.get("title") or "").strip()
    mark = str(item.get("mark") or "").strip()
    file_path = str(item.get("file_path") or "").strip()
    if not title or not mark or not file_path:
        return None
    return ExportPin(
        title=title,
        mark=mark,
        package_path=str(item.get("package_path") or ""),
        file_path=file_path,
        revision_text=str(item.get("revision_text") or ""),
        evidence=str(item.get("evidence") or ""),
        created_at=str(item.get("created_at") or ""),
    )


def _pin_to_json(pin: ExportPin) -> dict[str, Any]:
    return {
        "title": pin.title,
        "mark": pin.mark,
        "package_path": pin.package_path,
        "file_path": pin.file_path,
        "revision_text": pin.revision_text,
        "evidence": pin.evidence,
        "created_at": pin.created_at,
    }


def _read_json_object(path: Path) -> dict[str, Any] | None:
    if not path.is_file():
        return None
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return None
    if not isinstance(raw, dict):
        return None
    return raw


def _atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    tmp_path = path.with_name(path.name + ".tmp")
    tmp_path.write_text(text, encoding="utf-8")
    tmp_path.replace(path)


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()
