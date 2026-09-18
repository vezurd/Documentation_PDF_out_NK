"""Official-kit dump for managers: preview rows and copy MTO + BOE/BOM/BOQ.

Qt-free. Does not write under ``rd_root`` / ``sq_root``. Skip copy when the
destination file already has the source ``(size, mtime_ns)`` fingerprint.
"""

from __future__ import annotations

import json
import shutil
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path, PureWindowsPath

from rd_catalog.db import KitPackageRow, KitPipelineRow, mto_file_stat_signature
from rd_catalog.kits import format_revision, kit_identity_key, parse_sheet_revision
from rd_catalog.models import FileKind, FileRecord, SourceKind
from rd_catalog.overlay import revision_rank
from rd_catalog.path_actions import path_is_under
from rd_catalog.pipeline import (
    KitPipelineStatus,
    index_records_by_kit,
    pick_official_rd_package,
    pipeline_status_label,
    revision_texts_equivalent,
)
from rd_catalog.table_xlsx import (
    ExportedCell,
    ExportedColumn,
    ExportedTable,
    write_exported_table_xlsx,
)

HANDOFF_DESTINATIONS_FILENAME = "handoff_export_destinations.json"
HANDOFF_DESTINATIONS_VERSION = 1
HANDOFF_DESTINATIONS_MAX = 20
HANDOFF_LAYOUT_TITLE_MARK = "title_mark"
HANDOFF_LAYOUT_FLAT = "flat"
HANDOFF_TDO_STATUSES = frozenset(
    {
        KitPipelineStatus.SENT_TDO.value,
        KitPipelineStatus.TDO_REVIEW.value,
    }
)
HANDOFF_HEADERS = (
    "Титул",
    "Марка",
    "РД · рев.",
    "Статус",
    "Код A",
    "MTO",
    "BOE",
    "BOM",
    "BOQ",
    "Примечание",
    "Пакет",
)
HANDOFF_MATCH_FILL = "#E2F2E1"
HANDOFF_PROBLEM_FILL = "#F7E8BE"
_DOC_KINDS = ("mto", "boe", "bom", "boq")
CancelCallback = Callable[[], bool]
ProgressCallback = Callable[[int, int], None]


@dataclass(frozen=True, slots=True)
class HandoffKitInput:
    """One Комплекты kit for the dump (painted ``code_a``, official rev)."""

    title: str
    mark: str
    official_revision_text: str
    status: str
    code_a: bool
    pipeline: KitPipelineRow | None = None


@dataclass(frozen=True, slots=True)
class HandoffDocCell:
    """One MTO/BOE/BOM/BOQ cell in the preview."""

    kind: str
    text: str
    present: bool
    matches_official: bool
    path: str = ""
    fill_hex: str | None = None


@dataclass(frozen=True, slots=True)
class HandoffExportRow:
    """One preview / перечень row for an official kit package."""

    title: str
    mark: str
    official_revision_text: str
    status: str
    status_label: str
    code_a: bool
    package_path: str
    package_label: str
    mto: HandoffDocCell
    boe: HandoffDocCell
    bom: HandoffDocCell
    boq: HandoffDocCell
    notes: str
    has_problem: bool
    copy_files: tuple[tuple[str, str, int, int, str], ...]
    # Each copy tuple: source_path, filename, size, mtime_ns, kind


@dataclass(frozen=True, slots=True)
class HandoffCopyItem:
    """One file to copy or skip."""

    source_path: str
    dest_path: str
    size: int
    mtime_ns: int
    kind: str


@dataclass(frozen=True, slots=True)
class HandoffCopyPlan:
    """Validated copy request plus the preview rows used for the xlsx."""

    dest_root: str
    layout: str
    items: tuple[HandoffCopyItem, ...]
    rows: tuple[HandoffExportRow, ...]


@dataclass(frozen=True, slots=True)
class HandoffCopyReport:
    """Filesystem outcome of :func:`execute_handoff_copy`."""

    copied: int = 0
    skipped: int = 0
    failed: tuple[str, ...] = ()
    cancelled: bool = False
    list_path: str = ""


@dataclass(frozen=True, slots=True)
class HandoffDestinations:
    """Remembered dump roots (runtime JSON, never UNC catalog sources)."""

    paths: tuple[str, ...] = ()
    last: str = ""


def kit_handoff_eligible(item: HandoffKitInput, *, include_tdo: bool) -> bool:
    """Return whether the kit belongs on the dump tab.

    Args:
        item: Painted Комплекты flags (``code_a`` is the kits filter).
        include_tdo: Also keep ``sent_tdo`` / ``tdo_review``.

    Returns:
        True when the kit should appear in the preview.
    """

    if item.code_a:
        return True
    if include_tdo and (item.status or "") in HANDOFF_TDO_STATUSES:
        return True
    return False


def _record_revision_text(record: FileRecord) -> str:
    revision = record.data.get("revision")
    appendix = record.data.get("appendix")
    return format_revision(
        str(revision) if revision not in (None, "") else None,
        str(appendix) if appendix not in (None, "") else None,
    )


def _discipline_prefix(record: FileRecord) -> str:
    return str(record.data.get("discipline_block") or "").strip().casefold()


def _record_is_mto(record: FileRecord) -> bool:
    if str(record.data.get("file_kind") or "") == FileKind.MTO_XLSX.value:
        return True
    return _discipline_prefix(record).startswith("mto")


def _record_matches_kind(record: FileRecord, kind: str) -> bool:
    if kind == "mto":
        return _record_is_mto(record)
    return _discipline_prefix(record).startswith(kind)


def _doc_sort_key(
    record: FileRecord, official: str
) -> tuple[int, tuple[int, int, str], int, int]:
    text = _record_revision_text(record)
    kind = str(record.data.get("file_kind") or "")
    kind_rank = 2 if kind == FileKind.MTO_XLSX.value else (
        1 if kind == FileKind.PDF.value else 0
    )
    revision, appendix = parse_sheet_revision(text)
    return (
        int(bool(official and revision_texts_equivalent(text, official))),
        revision_rank(revision, appendix),
        kind_rank,
        int(record.data.get("mtime_ns") or 0),
    )


def _files_in_package(
    records: Sequence[FileRecord], package_path: str
) -> list[FileRecord]:
    return [
        record
        for record in records
        if record.present
        and record.source is SourceKind.RD
        and path_is_under(record.path, package_path)
    ]


def _pick_doc_cell(
    files: Sequence[FileRecord], kind: str, official: str
) -> HandoffDocCell:
    candidates = [item for item in files if _record_matches_kind(item, kind)]
    if not candidates:
        return HandoffDocCell(
            kind=kind,
            text="нет",
            present=False,
            matches_official=False,
            fill_hex=HANDOFF_PROBLEM_FILL,
        )
    chosen = max(candidates, key=lambda item: _doc_sort_key(item, official))
    text = _record_revision_text(chosen)
    matches = bool(official and revision_texts_equivalent(text, official))
    return HandoffDocCell(
        kind=kind,
        text=text or "—",
        present=True,
        matches_official=matches,
        path=chosen.path,
        fill_hex=HANDOFF_MATCH_FILL if matches else HANDOFF_PROBLEM_FILL,
    )


def _source_stat(record: FileRecord) -> tuple[int, int]:
    signature = mto_file_stat_signature(record)
    return int(signature.get("size") or 0), int(signature.get("mtime_ns") or 0)


def _record_copy_kind(record: FileRecord) -> str | None:
    for kind in _DOC_KINDS:
        if _record_matches_kind(record, kind):
            return kind
    return None


def _copy_payload(
    files: Sequence[FileRecord],
) -> tuple[tuple[str, str, int, int, str], ...]:
    payload: list[tuple[str, str, int, int, str]] = []
    seen: set[str] = set()
    for record in files:
        kind = _record_copy_kind(record)
        if kind is None:
            continue
        key = (record.path or "").casefold()
        if not key or key in seen:
            continue
        seen.add(key)
        size, mtime_ns = _source_stat(record)
        payload.append(
            (record.path, Path(record.path).name, size, mtime_ns, kind)
        )
    return tuple(payload)


def _row_notes(
    *,
    package_path: str,
    official: str,
    cells: Sequence[HandoffDocCell],
) -> str:
    parts: list[str] = []
    if not package_path:
        parts.append("нет пакета")
    for cell in cells:
        label = cell.kind.upper()
        if not cell.present:
            parts.append(f"нет {label}")
        elif official and not cell.matches_official:
            parts.append(f"{label} {cell.text} ≠ {official}")
    return "; ".join(parts)


def build_handoff_rows(
    kits: Sequence[HandoffKitInput],
    packages: Sequence[KitPackageRow],
    records: Sequence[FileRecord],
    *,
    include_tdo: bool = False,
    is_banned: Callable[[str, str], bool] | None = None,
) -> tuple[HandoffExportRow, ...]:
    """Build preview rows for eligible official kits.

    Args:
        kits: Комплекты universe with painted ``code_a``.
        packages: All ``kit_package`` rows (filtered per kit).
        records: Catalog files (present RD used).
        include_tdo: Add sent/passed TDO kits without code A.
        is_banned: Optional title–mark hide predicate.

    Returns:
        Sorted preview rows (title, then mark).
    """

    banned = is_banned or (lambda _title, _mark: False)
    by_kit_packages: dict[tuple[str, str], list[KitPackageRow]] = {}
    for package in packages:
        by_kit_packages.setdefault(
            kit_identity_key(package.title, package.mark), []
        ).append(package)
    by_kit_records = index_records_by_kit(records)
    rows: list[HandoffExportRow] = []
    for item in kits:
        if banned(item.title, item.mark):
            continue
        if not kit_handoff_eligible(item, include_tdo=include_tdo):
            continue
        key = kit_identity_key(item.title, item.mark)
        package = pick_official_rd_package(
            by_kit_packages.get(key, ()), item.pipeline
        )
        package_path = package.package_path if package is not None else ""
        package_label = ""
        if package is not None:
            package_label = package.transfer_name or PureWindowsPath(
                package.package_path
            ).name
        files = _files_in_package(by_kit_records.get(key, ()), package_path)
        official = (item.official_revision_text or "").strip()
        if not official and item.pipeline is not None:
            official = (item.pipeline.official_revision_text or "").strip()
        mto = _pick_doc_cell(files, "mto", official)
        boe = _pick_doc_cell(files, "boe", official)
        bom = _pick_doc_cell(files, "bom", official)
        boq = _pick_doc_cell(files, "boq", official)
        notes = _row_notes(
            package_path=package_path,
            official=official,
            cells=(mto, boe, bom, boq),
        )
        status_label = pipeline_status_label(item.status) if item.status else "—"
        rows.append(
            HandoffExportRow(
                title=item.title,
                mark=item.mark,
                official_revision_text=official or "—",
                status=item.status,
                status_label=status_label or "—",
                code_a=item.code_a,
                package_path=package_path,
                package_label=package_label or "—",
                mto=mto,
                boe=boe,
                bom=bom,
                boq=boq,
                notes=notes or "—",
                has_problem=bool(notes),
                copy_files=_copy_payload(files) if package_path else (),
            )
        )
    rows.sort(key=lambda row: (row.title, row.mark.casefold()))
    return tuple(rows)


def handoff_dest_relpath(
    *,
    layout: str,
    title: str,
    mark: str,
    filename: str,
) -> str:
    """Return the relative destination path for one payload file.

    Args:
        layout: ``title_mark`` or ``flat``.
        title: Four-digit title.
        mark: Latin AGCC mark.
        filename: Source basename.

    Returns:
        Relative path using the destination root as parent.
    """

    name = Path(filename).name
    if layout == HANDOFF_LAYOUT_FLAT:
        return name
    return str(Path(title) / mark / name)


def dest_matches_source_fingerprint(
    dest: Path, *, size: int, mtime_ns: int
) -> bool:
    """Return whether ``dest`` already has the catalog size+mtime fingerprint."""

    try:
        stat = dest.stat()
    except OSError:
        return False
    if not dest.is_file():
        return False
    return int(stat.st_size) == int(size) and int(stat.st_mtime_ns) == int(
        mtime_ns
    )


def validate_handoff_dest(
    dest_root: str,
    *,
    rd_root: str = "",
    sq_root: str = "",
) -> str:
    """Normalize the dump root and reject catalog source trees.

    Args:
        dest_root: User-chosen folder.
        rd_root: Catalog RD UNC; dest must not sit under it.
        sq_root: Catalog SQ UNC; dest must not sit under it.

    Returns:
        Stripped destination path.

    Raises:
        ValueError: Russian message when the folder is unusable.
    """

    dest = str(dest_root or "").strip()
    if not dest:
        raise ValueError("Укажите папку выгрузки.")
    if rd_root and path_is_under(dest, rd_root):
        raise ValueError("Нельзя выгружать внутрь корня РД.")
    if sq_root and path_is_under(dest, sq_root):
        raise ValueError("Нельзя выгружать внутрь корня SQ.")
    return dest


def build_handoff_copy_plan(
    rows: Sequence[HandoffExportRow],
    *,
    dest_root: str,
    layout: str,
    rd_root: str = "",
    sq_root: str = "",
) -> HandoffCopyPlan:
    """Resolve destination paths for MTO and BOE/BOM/BOQ files.

    Args:
        rows: Visible preview rows.
        dest_root: Dump root.
        layout: ``title_mark`` or ``flat``.
        rd_root: Catalog RD root (guard).
        sq_root: Catalog SQ root (guard).

    Returns:
        Plan consumed by :func:`execute_handoff_copy`.
    """

    dest = validate_handoff_dest(dest_root, rd_root=rd_root, sq_root=sq_root)
    chosen = layout if layout == HANDOFF_LAYOUT_FLAT else HANDOFF_LAYOUT_TITLE_MARK
    items: list[HandoffCopyItem] = []
    seen_dest: set[str] = set()
    for row in rows:
        for source_path, filename, size, mtime_ns, kind in row.copy_files:
            relative = handoff_dest_relpath(
                layout=chosen,
                title=row.title,
                mark=row.mark,
                filename=filename,
            )
            dest_path = str(Path(dest) / relative)
            dest_key = dest_path.casefold()
            if dest_key in seen_dest:
                continue
            seen_dest.add(dest_key)
            items.append(
                HandoffCopyItem(
                    source_path=source_path,
                    dest_path=dest_path,
                    size=size,
                    mtime_ns=mtime_ns,
                    kind=kind,
                )
            )
    return HandoffCopyPlan(
        dest_root=dest,
        layout=chosen,
        items=tuple(items),
        rows=tuple(rows),
    )


def handoff_list_filename(now: datetime | None = None) -> str:
    """Return ``Перечень_YYYY.MM.DD.xlsx`` for the dump folder."""

    stamp = (now or datetime.now()).strftime("%Y.%m.%d")
    return f"Перечень_{stamp}.xlsx"


def build_handoff_exported_table(
    rows: Sequence[HandoffExportRow],
) -> ExportedTable:
    """Paint the preview as an :class:`ExportedTable`."""

    columns = tuple(
        ExportedColumn(header=header, width_px=100, logical_index=index)
        for index, header in enumerate(HANDOFF_HEADERS)
    )
    exported_rows: list[tuple[ExportedCell, ...]] = []
    for row in rows:
        values = (
            (row.title, None),
            (row.mark, None),
            (row.official_revision_text, None),
            (row.status_label, None),
            ("да" if row.code_a else "нет", None),
            (row.mto.text, row.mto.fill_hex),
            (row.boe.text, row.boe.fill_hex),
            (row.bom.text, row.bom.fill_hex),
            (row.boq.text, row.boq.fill_hex),
            (
                row.notes,
                HANDOFF_PROBLEM_FILL if row.has_problem else None,
            ),
            (row.package_label, None),
        )
        exported_rows.append(
            tuple(
                ExportedCell(text=text, fill_hex=fill) for text, fill in values
            )
        )
    return ExportedTable(
        columns=columns,
        rows=tuple(exported_rows),
        sheet_name="Выгрузка",
    )


def execute_handoff_copy(
    plan: HandoffCopyPlan,
    *,
    now: datetime | None = None,
    cancel: CancelCallback | None = None,
    progress: ProgressCallback | None = None,
) -> HandoffCopyReport:
    """Copy MTO/BOE/BOM/BOQ files and write the dated перечень xlsx.

    Args:
        plan: Destination and file list.
        now: Clock for the перечень filename.
        cancel: Cooperative cancel between files.
        progress: ``(done, total)`` callback.

    Returns:
        Copy counts, failures, and the перечень path.
    """

    copied = 0
    skipped = 0
    failed: list[str] = []
    Path(plan.dest_root).mkdir(parents=True, exist_ok=True)
    total = len(plan.items)
    for index, item in enumerate(plan.items, start=1):
        if cancel is not None and cancel():
            return HandoffCopyReport(
                copied=copied,
                skipped=skipped,
                failed=tuple(failed),
                cancelled=True,
            )
        dest = Path(item.dest_path)
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            if dest_matches_source_fingerprint(
                dest, size=item.size, mtime_ns=item.mtime_ns
            ):
                skipped += 1
            else:
                shutil.copy2(item.source_path, dest)
                copied += 1
        except OSError as exc:
            failed.append(f"{item.source_path} → {item.dest_path}: {exc}")
        if progress is not None:
            progress(index, total)
    list_name = handoff_list_filename(now)
    list_path = str(Path(plan.dest_root) / list_name)
    try:
        write_exported_table_xlsx(
            build_handoff_exported_table(plan.rows), list_path
        )
    except OSError as exc:
        failed.append(f"{list_path}: {exc}")
        list_path = ""
    return HandoffCopyReport(
        copied=copied,
        skipped=skipped,
        failed=tuple(failed),
        list_path=list_path,
    )


def _destinations_path(runtime_dir: str | Path) -> Path:
    return Path(runtime_dir) / HANDOFF_DESTINATIONS_FILENAME


def load_handoff_destinations(runtime_dir: str | Path) -> HandoffDestinations:
    """Read remembered dump roots from ``runtime_dir``."""

    path = _destinations_path(runtime_dir)
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, TypeError):
        return HandoffDestinations()
    if not isinstance(payload, dict):
        return HandoffDestinations()
    raw_paths = payload.get("paths")
    paths: list[str] = []
    if isinstance(raw_paths, list):
        for item in raw_paths:
            text = str(item or "").strip()
            if text and text not in paths:
                paths.append(text)
            if len(paths) >= HANDOFF_DESTINATIONS_MAX:
                break
    last = str(payload.get("last") or "").strip()
    if last and last not in paths:
        paths.insert(0, last)
        paths = paths[:HANDOFF_DESTINATIONS_MAX]
    if last and last not in paths:
        last = paths[0] if paths else ""
    return HandoffDestinations(paths=tuple(paths), last=last)


def remember_handoff_destination(
    runtime_dir: str | Path, dest_root: str
) -> HandoffDestinations:
    """Put ``dest_root`` first in the remembered list and persist JSON."""

    dest = str(dest_root or "").strip()
    current = load_handoff_destinations(runtime_dir)
    paths = [dest] if dest else []
    for item in current.paths:
        if item != dest:
            paths.append(item)
        if len(paths) >= HANDOFF_DESTINATIONS_MAX:
            break
    store = HandoffDestinations(paths=tuple(paths), last=dest)
    payload = {
        "version": HANDOFF_DESTINATIONS_VERSION,
        "paths": list(store.paths),
        "last": store.last,
    }
    path = _destinations_path(runtime_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return store
