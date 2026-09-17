"""Qt-free helpers for the heatmap export cockpit and batch copy.

Does not read XLSX. Copy execution reuses :func:`execute_robot_mto_sync`
(archive convention ``_old_<stem>_YYYY.MM.DD_HH.MM`` and ``path_is_under``
against the target root).
"""

from __future__ import annotations

import os
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path

from rd_catalog.db import CatalogDatabase, canonical_mto_pair_ids, mto_file_stat_signature
from rd_catalog.kits import kit_identity_key
from rd_catalog.models import FileRecord, SourceKind, make_path_key
from rd_catalog.mto_export import (
    ExportPlan,
    ExportPlanItem,
    ExportSelection,
    ExportTarget,
    build_export_plan,
)
from rd_catalog.mto_pair_compare import MtoFilePair, PairPoolStatus
from rd_catalog.path_actions import path_is_under
from rd_catalog.robot_mto_sync import (
    MtoFileRef,
    RobotMtoSyncError,
    RobotMtoSyncPlan,
    execute_robot_mto_sync,
)

QSETTINGS_EXPORT_TARGET = "window/rev_matrix_export_target"

EXPORT_RULE_LABELS: tuple[tuple[str, str], ...] = (
    ("approved", "Согласованное (код А)"),
    ("latest_issued", "Последнее выданное"),
    ("latest_no_as_build", "Последнее без as-build"),
    ("latest_tdo", "Последнее, прошедшее ТДО"),
)

EXPORT_STATE_TEXT: dict[str, str] = {
    "add": "добавится",
    "replace": "заменится",
    "same_data": "данные совпадают",
    "no_source": "нет файла",
    "unknown": "не сравнено",
    "pin_stale": "ручной выбор устарел",
}

EXPORT_STATE_COLOR_KEY: dict[str, str] = {
    "add": "export_add",
    "replace": "export_replace",
    "same_data": "export_same",
    "no_source": "export_missing",
    "unknown": "export_unknown",
    "pin_stale": "export_pin_stale",
}

PREVIEW_GROUP_ORDER: tuple[tuple[str, str], ...] = (
    ("add", "добавится"),
    ("replace", "заменится"),
    ("same_data", "данные совпадают"),
    ("no_source", "нет файла"),
    ("unknown", "не сравнено"),
    ("pin_stale", "ручной выбор устарел"),
)

_MATERIAL_STATUSES = frozenset({"content_equal", "content_diff"})
_COPY_STATES = frozenset({"add", "replace"})

CancelCallback = Callable[[], bool]
ProgressCallback = Callable[[int, int], None]


@dataclass(frozen=True, slots=True)
class ExportPreviewGroup:
    """One labelled block of the combined export preview."""

    state: str
    title: str
    selections: tuple[ExportSelection, ...]

    @property
    def count(self) -> int:
        """Return the number of kits in this group."""

        return len(self.selections)


@dataclass(frozen=True, slots=True)
class ExportCopyItemResult:
    """Outcome of one copy/replace step."""

    selection: ExportSelection
    destination_path: str
    archived_paths: tuple[str, ...]
    error: str | None = None


@dataclass(frozen=True, slots=True)
class ExportCopyReport:
    """Filesystem outcome of :func:`execute_export_copy`."""

    copied: tuple[ExportCopyItemResult, ...]
    archived: tuple[ExportCopyItemResult, ...]
    skipped: tuple[ExportSelection, ...]
    failed: tuple[ExportCopyItemResult, ...]
    cancelled: bool = False


def group_export_preview(
    selections: Sequence[ExportSelection],
) -> tuple[ExportPreviewGroup, ...]:
    """Group visible-row selections for the combined preview dialog.

    Empty groups are omitted. ``same_data`` and ``no_source`` are included
    so the user sees them; they never become copy items.

    Args:
        selections: Currently visible heatmap rows.

    Returns:
        Groups in :data:`PREVIEW_GROUP_ORDER`.
    """

    by_state: dict[str, list[ExportSelection]] = {}
    for selection in selections:
        by_state.setdefault(selection.state, []).append(selection)
    groups: list[ExportPreviewGroup] = []
    for state, title in PREVIEW_GROUP_ORDER:
        rows = tuple(by_state.get(state) or ())
        if rows:
            groups.append(
                ExportPreviewGroup(state=state, title=title, selections=rows)
            )
    extra_states = [
        state for state in by_state if state not in dict(PREVIEW_GROUP_ORDER)
    ]
    for state in extra_states:
        rows = tuple(by_state[state])
        groups.append(
            ExportPreviewGroup(
                state=state,
                title=EXPORT_STATE_TEXT.get(state, state),
                selections=rows,
            )
        )
    return tuple(groups)


def copy_selections(
    selections: Sequence[ExportSelection],
) -> tuple[ExportSelection, ...]:
    """Return rows that the copy worker will actually write.

    Args:
        selections: Visible-row selections (preview universe).

    Returns:
        ``add`` and ``replace`` only.
    """

    return tuple(
        selection
        for selection in selections
        if selection.state in _COPY_STATES
    )


def plan_for_visible_rows(
    selections: Sequence[ExportSelection],
    *,
    target: ExportTarget,
) -> ExportPlan:
    """Build a copy plan from currently visible heatmap rows.

    Args:
        selections: Visible-row selections.
        target: Destination whose root bounds every write.

    Returns:
        Plan whose items never include ``same_data`` / ``no_source``.
    """

    return build_export_plan(selections, target=target)


def export_selection_tooltip(selection: ExportSelection) -> str:
    """Build the third-column tooltip from a stored selection.

    For ``pin_stale`` both the pinned file and ``rule_path`` are listed.

    Args:
        selection: Row stored on the table item.

    Returns:
        Multi-line Russian tooltip.
    """

    source_name = Path(selection.source_path).name if selection.source_path else "—"
    lines = [
        f"{selection.title}-{selection.mark}",
        f"состояние: {EXPORT_STATE_TEXT.get(selection.state, selection.state)}",
        f"файл: {source_name}",
        f"пакет: {selection.package_path or '—'}",
        f"назначение: {selection.destination_path or '—'}",
        f"уверенность: {selection.confidence or '—'}",
    ]
    if selection.state == "pin_stale" or selection.origin in {"pin", "pin_stale"}:
        lines.append(f"закреплённый файл: {selection.source_path or '—'}")
        lines.append(f"файл по правилу: {selection.rule_path or '—'}")
    elif selection.rule_path and not _paths_equal(
        selection.source_path, selection.rule_path
    ):
        lines.append(f"файл по правилу: {selection.rule_path}")
    if selection.warnings:
        lines.append("предупреждения:")
        lines.extend(f"  {warning}" for warning in selection.warnings)
    return "\n".join(lines)


def format_pool_label(status: PairPoolStatus) -> str:
    """Return the Russian remaining-pairs label for the heatmap cockpit.

    Args:
        status: Current compare-pool counters.

    Returns:
        One-line status text.
    """

    unpersisted_note = ""
    if status.unpersisted:
        unpersisted_note = (
            f" Пары вне каталога ({status.unpersisted}) пересчитываются "
            "каждый сеанс — папка назначения не в каталоге."
        )
    if not status.is_drained:
        return (
            f"осталось сравнить: {status.pending} из {status.total}"
            f"{unpersisted_note}"
        )
    text = "сравнение завершено"
    if status.failed:
        text += f", не удалось прочитать: {status.failed}"
    return text + unpersisted_note


def format_preview_body(
    groups: Sequence[ExportPreviewGroup],
    *,
    target: ExportTarget,
) -> str:
    """Build the combined preview dialog text.

    Args:
        groups: Output of :func:`group_export_preview`.
        target: Destination shown in the header.

    Returns:
        Multi-line Russian preview.
    """

    lines = [
        f"Папка: {target.name}",
        f"Путь: {target.root}",
        "",
    ]
    for group in groups:
        lines.append(f"{group.title} ({group.count})")
        for selection in group.selections:
            source_name = (
                Path(selection.source_path).name if selection.source_path else "—"
            )
            lines.append(f"  {selection.title}-{selection.mark}")
            lines.append(f"    файл: {source_name}")
            lines.append(f"    пакет: {selection.package_path or '—'}")
            lines.append(f"    назначение: {selection.destination_path or '—'}")
            if selection.warnings:
                for warning in selection.warnings:
                    lines.append(f"    ! {warning}")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def format_export_copy_report(report: ExportCopyReport) -> str:
    """Build the final copy-report text.

    Args:
        report: Outcome of :func:`execute_export_copy`.

    Returns:
        Multi-line Russian summary.
    """

    lines = []
    if report.cancelled:
        lines.append("Копирование прервано.")
        lines.append("")
    lines.append(f"Скопировано: {len(report.copied)}")
    for item in report.copied:
        kit = f"{item.selection.title}-{item.selection.mark}"
        lines.append(f"  {kit} → {item.destination_path}")
    lines.append(f"Отложено в архив: {len(report.archived)}")
    for item in report.archived:
        kit = f"{item.selection.title}-{item.selection.mark}"
        for archived in item.archived_paths:
            lines.append(f"  {kit}: {archived}")
    lines.append(f"Пропущено: {len(report.skipped)}")
    for selection in report.skipped:
        label = EXPORT_STATE_TEXT.get(selection.state, selection.state)
        lines.append(f"  {selection.title}-{selection.mark} ({label})")
    lines.append(f"Ошибки: {len(report.failed)}")
    for item in report.failed:
        kit = f"{item.selection.title}-{item.selection.mark}"
        lines.append(f"  {kit}: {item.error or 'неизвестно'}")
    return "\n".join(lines)


def session_status_map(
    session_verdicts: Mapping[tuple[str, str], object],
) -> dict[tuple[str, str], str]:
    """Flatten thread/session verdicts to ``content_*`` status strings.

    Args:
        session_verdicts: Path-key map of strings or objects with
            ``content_status``.

    Returns:
        Canonical path-key → status.
    """

    result: dict[tuple[str, str], str] = {}
    for key, value in session_verdicts.items():
        if isinstance(value, str):
            status = value
        else:
            status = str(getattr(value, "content_status", "") or "")
        if not status:
            continue
        left, right = key
        if left > right:
            left, right = right, left
        result[(left, right)] = status
    return result


def file_id_verdicts_from_db(
    database: CatalogDatabase,
    *,
    records: Sequence[FileRecord],
    pairs: Sequence[MtoFilePair],
) -> dict[tuple[int, int], str]:
    """Return signature-valid material verdicts keyed by canonical file ids.

    Args:
        database: Catalog SQLite.
        records: Catalog file rows for ids and signatures.
        pairs: Current export compare pool.

    Returns:
        ``(min_id, max_id)`` → ``content_equal`` / ``content_diff``.
    """

    index = {make_path_key(record.path): record for record in records}
    id_pairs: list[tuple[int, int]] = []
    resolved: list[tuple[tuple[int, int], FileRecord, FileRecord]] = []
    seen: set[tuple[int, int]] = set()
    for pair in pairs:
        left = index.get(make_path_key(pair.left_path))
        right = index.get(make_path_key(pair.right_path))
        if left is None or right is None:
            continue
        key = canonical_mto_pair_ids(left.id, right.id)
        if key in seen:
            continue
        seen.add(key)
        id_pairs.append(key)
        resolved.append((key, left, right))
    rows = database.list_mto_pair_comparisons(id_pairs)
    verdicts: dict[tuple[int, int], str] = {}
    for key, left, right in resolved:
        row = rows.get(key)
        if row is None:
            continue
        status = str(row.get("content_status") or "")
        if status not in _MATERIAL_STATUSES:
            continue
        if not _signatures_match(row, key[0], key[1], left, right):
            continue
        verdicts[key] = status
    return verdicts


def overlay_session_on_selections(
    selections: Sequence[ExportSelection],
    session_verdicts: Mapping[tuple[str, str], str],
) -> tuple[ExportSelection, ...]:
    """Apply path-keyed session verdicts to ``unknown`` rows.

    Used for custom-target files that are not in ``file_entry``, so
    :func:`resolve_export_selections` cannot look them up by file id.

    Args:
        selections: Rows from :func:`resolve_export_selections`.
        session_verdicts: Canonical path-key → status.

    Returns:
        Selections with ``same_data`` / ``replace`` patched in.
    """

    if not session_verdicts:
        return tuple(selections)
    updated: list[ExportSelection] = []
    for selection in selections:
        if selection.state in {"add", "no_source", "pin_stale"}:
            updated.append(selection)
            continue
        if not selection.source_path or not selection.existing_target_path:
            updated.append(selection)
            continue
        status = _lookup_session(
            session_verdicts, selection.source_path, selection.existing_target_path
        )
        if status == "content_equal":
            updated.append(replace(selection, state="same_data"))
        elif status == "content_diff":
            updated.append(replace(selection, state="replace"))
        else:
            updated.append(selection)
    return tuple(updated)


def patch_selections_with_verdicts(
    selections: Sequence[ExportSelection],
    *,
    records: Sequence[FileRecord],
    file_id_verdicts: Mapping[tuple[int, int], str],
    session_verdicts: Mapping[tuple[str, str], str],
) -> tuple[ExportSelection, ...]:
    """Update ``unknown``/material states from DB + session without re-resolve.

    Args:
        selections: Current heatmap export rows.
        records: Catalog files for path → id.
        file_id_verdicts: Signature-valid DB material verdicts.
        session_verdicts: This-session path-key statuses.

    Returns:
        Patched selections.
    """

    files_by_path = {make_path_key(record.path): record for record in records}
    updated: list[ExportSelection] = []
    for selection in selections:
        if selection.state in {"add", "no_source", "pin_stale"}:
            updated.append(selection)
            continue
        status = _lookup_file_id_verdict(
            selection, files_by_path, file_id_verdicts
        )
        if status is None:
            status = _lookup_session(
                session_verdicts,
                selection.source_path,
                selection.existing_target_path,
            )
        if status == "content_equal":
            updated.append(replace(selection, state="same_data"))
        elif status == "content_diff":
            updated.append(replace(selection, state="replace"))
        else:
            updated.append(selection)
    return tuple(updated)


def selections_by_kit(
    selections: Sequence[ExportSelection],
) -> dict[tuple[str, str], ExportSelection]:
    """Index selections by :func:`kit_identity_key`.

    Args:
        selections: Export preview rows.

    Returns:
        Identity → selection (last wins).
    """

    return {
        kit_identity_key(selection.title, selection.mark): selection
        for selection in selections
    }


def pair_identity_key(pairs: Sequence[MtoFilePair]) -> frozenset[tuple[str, str]]:
    """Return a stable identity of a compare pool.

    Args:
        pairs: Source/destination pairs.

    Returns:
        Frozen set of canonical path-key pairs.
    """

    return frozenset(
        _canonical_path_keys(pair.left_path, pair.right_path) for pair in pairs
    )


def sync_plan_from_item(
    item: ExportPlanItem,
    *,
    target: ExportTarget,
) -> RobotMtoSyncPlan:
    """Convert one export plan item into :class:`RobotMtoSyncPlan`.

    Args:
        item: Copy/replace step from :func:`build_export_plan`.
        target: Destination (root used as ``robot_root`` for the write guard).

    Returns:
        Plan that :func:`execute_robot_mto_sync` can run.

    Raises:
        RobotMtoSyncError: If the destination escapes ``target.root``.
    """

    selection = item.selection
    if not path_is_under(selection.destination_path, target.root):
        raise RobotMtoSyncError(
            "Путь назначения не находится в папке экспорта:\n"
            f"{selection.destination_path}"
        )
    dest = Path(selection.destination_path)
    source = _ref_for_copy(
        selection.source_path,
        title=selection.title,
        mark=selection.mark,
        source=SourceKind.RD,
    )
    robot_files: tuple[MtoFileRef, ...] = ()
    if selection.state == "replace" and selection.existing_target_path:
        robot_files = (
            _ref_for_copy(
                selection.existing_target_path,
                title=selection.title,
                mark=selection.mark,
                source=SourceKind.ROBOT,
            ),
        )
    return RobotMtoSyncPlan(
        title=selection.title,
        mark=selection.mark,
        source=source,
        robot_files=robot_files,
        destination_dir=str(dest.parent),
        destination_path=str(dest),
        robot_root=str(target.root),
    )


def execute_export_copy(
    plan: ExportPlan,
    *,
    now: datetime | None = None,
    cancel: CancelCallback | None = None,
    progress: ProgressCallback | None = None,
) -> ExportCopyReport:
    """Copy planned MTO files using :func:`execute_robot_mto_sync`.

    Every write is guarded by ``path_is_under(…, plan.target.root)``.
    Replaced files are archived with ``_old_<stem>_YYYY.MM.DD_HH.MM``.

    Args:
        plan: Output of :func:`build_export_plan`.
        now: Timestamp for archive folder names.
        cancel: Cooperative cancel, checked between items.
        progress: ``(completed, total)`` callback over copy items.

    Returns:
        Copied / archived / skipped / failed lists.
    """

    when = now or datetime.now()
    copied: list[ExportCopyItemResult] = []
    archived: list[ExportCopyItemResult] = []
    failed: list[ExportCopyItemResult] = []
    cancelled = False
    remaining: tuple[ExportSelection, ...] = ()
    total = len(plan.items)
    for index, item in enumerate(plan.items):
        if cancel is not None and cancel():
            cancelled = True
            remaining = tuple(entry.selection for entry in plan.items[index:])
            break
        selection = item.selection
        try:
            if not path_is_under(selection.destination_path, plan.target.root):
                raise RobotMtoSyncError(
                    "Путь назначения не находится в папке экспорта."
                )
            sync_plan = sync_plan_from_item(item, target=plan.target)
            result = execute_robot_mto_sync(sync_plan, now=when)
        except RobotMtoSyncError as exc:
            failed.append(
                ExportCopyItemResult(
                    selection=selection,
                    destination_path=selection.destination_path,
                    archived_paths=(),
                    error=str(exc),
                )
            )
        except OSError as exc:
            failed.append(
                ExportCopyItemResult(
                    selection=selection,
                    destination_path=selection.destination_path,
                    archived_paths=(),
                    error=f"{type(exc).__name__}: {exc}",
                )
            )
        else:
            archived_paths = tuple(entry.archived_path for entry in result.archived)
            row = ExportCopyItemResult(
                selection=selection,
                destination_path=result.destination_path,
                archived_paths=archived_paths,
            )
            copied.append(row)
            if archived_paths:
                archived.append(row)
        if progress is not None:
            progress(index + 1, total)
    skipped = tuple(plan.skipped) + remaining
    return ExportCopyReport(
        copied=tuple(copied),
        archived=tuple(archived),
        skipped=skipped,
        failed=tuple(failed),
        cancelled=cancelled,
    )


def preview_copy_count(selections: Sequence[ExportSelection]) -> int:
    """Return how many visible rows would actually be copied.

    Args:
        selections: Visible-row selections.

    Returns:
        Count of ``add`` + ``replace``.
    """

    return len(copy_selections(selections))


def _ref_for_copy(
    path: str,
    *,
    title: str,
    mark: str,
    source: SourceKind,
) -> MtoFileRef:
    name = Path(path).name
    mtime_ns = 0
    size = 0
    try:
        stat = os.stat(path, follow_symlinks=False)
        mtime_ns = int(stat.st_mtime_ns)
        size = int(stat.st_size)
    except OSError:
        pass
    return MtoFileRef(
        path=path,
        name=name,
        source=source,
        title=title,
        mark=mark,
        title_system=f"{title}-{mark}",
        discipline_block=None,
        revision=None,
        appendix=None,
        revision_text="",
        mtime_ns=mtime_ns,
        size=size,
    )


def _canonical_path_keys(left_path: str, right_path: str) -> tuple[str, str]:
    first = make_path_key(left_path)
    second = make_path_key(right_path)
    if first <= second:
        return first, second
    return second, first


def _paths_equal(left: str, right: str) -> bool:
    if not left or not right:
        return False
    return make_path_key(left) == make_path_key(right)


def _lookup_session(
    session_verdicts: Mapping[tuple[str, str], str],
    left_path: str,
    right_path: str,
) -> str | None:
    if not left_path or not right_path or not session_verdicts:
        return None
    key = _canonical_path_keys(left_path, right_path)
    status = session_verdicts.get(key)
    if status is not None:
        return status
    return session_verdicts.get((key[1], key[0]))


def _lookup_file_id_verdict(
    selection: ExportSelection,
    files_by_path: Mapping[str, FileRecord],
    verdicts: Mapping[tuple[int, int], str],
) -> str | None:
    if not selection.source_path or not selection.existing_target_path:
        return None
    source = files_by_path.get(make_path_key(selection.source_path))
    existing = files_by_path.get(make_path_key(selection.existing_target_path))
    if source is None or existing is None:
        return None
    key = canonical_mto_pair_ids(source.id, existing.id)
    return verdicts.get(key)


def _signatures_match(
    row: Mapping[str, object],
    left_id: int,
    right_id: int,
    left_file: FileRecord,
    right_file: FileRecord,
) -> bool:
    left_sig = mto_file_stat_signature(left_file)
    right_sig = mto_file_stat_signature(right_file)
    if int(left_id) > int(right_id):
        left_sig, right_sig = right_sig, left_sig
    return row.get("left_signature") == left_sig and row.get(
        "right_signature"
    ) == right_sig


__all__ = [
    "EXPORT_RULE_LABELS",
    "EXPORT_STATE_COLOR_KEY",
    "EXPORT_STATE_TEXT",
    "ExportCopyItemResult",
    "ExportCopyReport",
    "ExportPreviewGroup",
    "PREVIEW_GROUP_ORDER",
    "QSETTINGS_EXPORT_TARGET",
    "copy_selections",
    "execute_export_copy",
    "export_selection_tooltip",
    "file_id_verdicts_from_db",
    "format_export_copy_report",
    "format_pool_label",
    "format_preview_body",
    "group_export_preview",
    "overlay_session_on_selections",
    "pair_identity_key",
    "patch_selections_with_verdicts",
    "plan_for_visible_rows",
    "preview_copy_count",
    "selections_by_kit",
    "session_status_map",
    "sync_plan_from_item",
]
