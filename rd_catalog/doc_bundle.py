"""Pair RD PDF files with nearby source/editable companions for the catalog GUI.

This module is Qt-free. Overlay still covers only PDF and MTO XLSX;
``SOURCE_EDITABLE`` files are grouped here for display (the «Ред» column).
"""

from __future__ import annotations

import os
from collections import defaultdict
from collections.abc import Collection, Iterable
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path, PureWindowsPath

from rd_catalog.models import FileKind, FileRecord, SourceKind, make_path_key
from rd_catalog.overlay import revision_rank
from rd_catalog.parse import issued_package_dir, parse_transfer_folder, path_is_as_build

EDITABLE_KINDS = frozenset({FileKind.SOURCE_EDITABLE, FileKind.MTO_XLSX})
NO_REVISION_LABEL = "Без ревизии"
CURRENT_MARKER = "текущая"
MTO_MARKER = "MTO"
WORKING_MARKER = "рабочая"
ANNULLED_MARKER = "аннулирована"
AS_BUILD_MARKER = "AB"
WORKING_TOOLTIP_MANUAL = (
    "Помечена как рабочая (вручную). "
    "Не участвует в «РД · рев.», экспорте и сверке."
)
WORKING_TOOLTIP_AUTO = (
    "Рабочая ревизия: файлы строго выше последней официальной выдачи / F."
)
ANNULLED_TOOLTIP = (
    "Помечена как аннулированная. Не участвует в «РД · рев.», "
    "конкурсе рабочей ревизии, экспорте, сверке и «Проверить передачи»."
)

_BUNDLE_KINDS = frozenset({FileKind.PDF, FileKind.SOURCE_EDITABLE, FileKind.MTO_XLSX})


@dataclass(frozen=True, slots=True)
class DocumentTreeLabelOptions:
    """Which extras to append to a «Все документы» revision node.

    Filename revision is always the first token. Date stays on by default
    so the tree matches the historical ``01-AN02 (2024.08.16)`` shape.
    """

    show_date: bool = True
    show_folder: bool = False
    show_review: bool = False
    show_current: bool = False
    show_mto: bool = False
    show_working: bool = False
    show_as_build: bool = False
    show_mto_status: bool = False


@dataclass(frozen=True, slots=True)
class DocumentBundle:
    """One catalog tree row: a primary PDF and paired editable files.

    Attributes:
        title: Four-digit title number.
        mark: Mark/system component.
        revision_label: This document's filename revision (table column).
        folder_key: Transfer folder name or parent directory identity.
        core_stem: Normalized AGCC stem shared by the group.
        discipline: Discipline code such as ``OD`` or ``MTO``.
        pdf: Primary PDF record, if any.
        editables: MTO XLSX and source-editable files in the group.
        is_current: Whether ``pdf.id`` is in the overlay current set.
    """

    title: str
    mark: str
    revision_label: str
    folder_key: str
    core_stem: str
    discipline: str
    pdf: FileRecord | None
    editables: tuple[FileRecord, ...]
    is_current: bool


def format_rev_label(revision: str | None, appendix: str | None) -> str:
    """Format a revision/AN pair as ``01-AN02``.

    Args:
        revision: Revision token such as ``01``.
        appendix: Optional AN digits without the ``AN`` prefix.

    Returns:
        Display text, or empty string when revision is missing.
    """

    if not revision:
        return ""
    return f"{revision}-AN{appendix}" if appendix else str(revision)


def file_revision_label(record: FileRecord) -> str:
    """Return the filename revision/AN of one file, or empty string.

    Transfer-folder names are ignored: the document revision comes from
    the AGCC filename.

    Args:
        record: Persisted catalog file.

    Returns:
        Display text such as ``01-AN02``, or ``""`` when missing.
    """

    return format_rev_label(*_file_revision_tokens(record))


def record_revision_key(record: FileRecord) -> tuple[str, str]:
    """Return display revision label and string sort key for one file.

    Filename revision/AN is preferred, then the transfer folder, then
    :data:`NO_REVISION_LABEL`.

    Args:
        record: Persisted catalog file.

    Returns:
        ``(display_label, sort_key)``. Sort key is empty when there is no
        revision.
    """

    label = file_revision_label(record)
    if not label:
        revision, appendix = _transfer_revision_tokens(record)
        label = format_rev_label(revision, appendix)
    if not label:
        return (NO_REVISION_LABEL, "")
    return (label, label)


def record_folder_key(record: FileRecord) -> str:
    """Return the transfer-or-directory identity used to group a tree node.

    Args:
        record: Persisted catalog file.

    Returns:
        Case-folded transfer folder name, or a normalized parent path.
    """

    transfer_name = _data_str(record, "transfer_name")
    if transfer_name:
        return transfer_name.casefold()
    return make_path_key(str(Path(record.path).parent))


def folder_revision_label(records: Iterable[FileRecord]) -> str:
    """Return the highest filename revision among files in one folder.

    OD documents win ties, matching typical transfer packages.

    Args:
        records: Files that belong to the same transfer/directory.

    Returns:
        Display revision, or :data:`NO_REVISION_LABEL` when none have a
        filename revision.
    """

    materialized = [record for record in records]
    if not materialized:
        return NO_REVISION_LABEL

    def rank(record: FileRecord) -> tuple[tuple[int, int, str], int]:
        tokens = _file_revision_tokens(record)
        is_od = int(_discipline(record).casefold() == "od")
        return (revision_rank(*tokens), is_od)

    best = max(materialized, key=rank)
    return file_revision_label(best) or NO_REVISION_LABEL


def latest_save_mtime_ns(records: Iterable[FileRecord]) -> int:
    """Return the newest PDF / MTO / source-editable mtime in nanoseconds.

    Args:
        records: Catalog files (a folder, or one document bundle).

    Returns:
        Max ``mtime_ns``, or ``0`` when none is usable.
    """

    latest = 0
    for record in records:
        kind = _record_file_kind(record)
        if kind not in _BUNDLE_KINDS:
            continue
        mtime_ns = int(record.data.get("mtime_ns") or 0)
        if mtime_ns > latest:
            latest = mtime_ns
    return latest


def format_file_save_date(mtime_ns: int, *, with_time: bool = False) -> str:
    """Format a catalog mtime as ``YYYY.MM.DD`` (optionally with ``HH:MM``).

    Args:
        mtime_ns: Nanoseconds since epoch.
        with_time: When true, append local ``HH:MM``.

    Returns:
        Display text, or empty string when ``mtime_ns`` is missing.
    """

    if mtime_ns <= 0:
        return ""
    try:
        stamp = datetime.fromtimestamp(mtime_ns / 1_000_000_000)
    except (OSError, OverflowError, ValueError):
        return ""
    if with_time:
        return stamp.strftime("%Y.%m.%d %H:%M")
    return stamp.strftime("%Y.%m.%d")


def folder_latest_save_date(records: Iterable[FileRecord]) -> str:
    """Return ``YYYY.MM.DD`` of the newest PDF or source file mtime.

    Only catalog files that already match the AGCC mask are considered
    (PDF, MTO XLSX, source-editables).

    Args:
        records: Files that belong to the same transfer/directory.

    Returns:
        Local calendar date, or empty string when no usable mtime exists.
    """

    return format_file_save_date(latest_save_mtime_ns(records))


def file_save_date_tooltip(records: Iterable[FileRecord]) -> str:
    """Return one ``name — YYYY.MM.DD HH:MM`` line per catalog file.

    Args:
        records: Files shown on one documents-table row.

    Returns:
        Multiline tooltip text, or empty when the row has no files.
    """

    lines: list[str] = []
    for record in records:
        kind = _record_file_kind(record)
        if kind not in _BUNDLE_KINDS:
            continue
        stamp = format_file_save_date(
            int(record.data.get("mtime_ns") or 0),
            with_time=True,
        ) or "—"
        name = _record_name(record)
        extra = mtime_override_tooltip(record)
        if extra:
            lines.append(f"{name} — {stamp}\n{extra}")
        else:
            lines.append(f"{name} — {stamp}")
    return "\n".join(lines)


def record_disk_mtime_ns(record: FileRecord) -> int:
    """Return the scan/disk mtime, ignoring a catalog-date override.

    Args:
        record: Catalog file.

    Returns:
        Nanoseconds since epoch, or ``0``.
    """

    return int(record.data.get("disk_mtime_ns") or record.data.get("mtime_ns") or 0)


def override_date_text_from_mtime_ns(mtime_ns: int) -> str:
    """Return local calendar date as ``DD.MM.YYYY``.

    Args:
        mtime_ns: Nanoseconds since epoch.

    Returns:
        Override-date text, or empty when ``mtime_ns`` is missing.
    """

    if mtime_ns <= 0:
        return ""
    try:
        stamp = datetime.fromtimestamp(mtime_ns / 1_000_000_000)
    except (OSError, OverflowError, ValueError):
        return ""
    return f"{stamp.day:02d}.{stamp.month:02d}.{stamp.year:04d}"


def same_document_path_keys(
    current: FileRecord,
    folder_records: Iterable[FileRecord],
) -> frozenset[str]:
    """Return path keys of ``current`` and same-stem companions in the folder.

    PDF and xlsx of one MTO share ``core_stem`` and were often re-saved
    together, so both count as the outlier document.

    Args:
        current: File whose catalog date is being replaced.
        folder_records: Catalog files in the same issued package / tree node.

    Returns:
        Case-folded ``path_key`` values.
    """

    keys = {str(current.path_key or "").casefold()}
    stem = str(current.data.get("core_stem") or "").strip().casefold()
    if not stem:
        return frozenset(key for key in keys if key)
    for record in folder_records:
        other = str(record.data.get("core_stem") or "").strip().casefold()
        if other != stem:
            continue
        key = str(record.path_key or "").casefold()
        if key:
            keys.add(key)
    return frozenset(key for key in keys if key)


@dataclass(frozen=True, slots=True)
class FolderMeanOverride:
    """Mean disk date of package files after dropping the outlier document.

    Attributes:
        override_date: ``DD.MM.YYYY`` for ``file_mtime_override``.
        used_count: Files that entered the mean.
        excluded_count: Same-document files skipped as the outlier.
        mtime_ns: Mean timestamp (local-noon is applied at persist).
    """

    override_date: str
    used_count: int
    excluded_count: int
    mtime_ns: int


def folder_mean_override(
    folder_records: Iterable[FileRecord],
    *,
    exclude_path_keys: Collection[str],
) -> FolderMeanOverride | None:
    """Mean disk date of folder files, skipping the current outlier document.

    Uses scan/disk mtime so another file's catalog override does not shift
    the mean. No UNC walk.

    Args:
        folder_records: Catalog files in the issued package (tree node).
        exclude_path_keys: Current file and same-stem companions.

    Returns:
        Mean date, or ``None`` when no other dated catalog files remain.
    """

    excluded = {str(key or "").casefold() for key in exclude_path_keys if key}
    used: list[int] = []
    skipped = 0
    for record in folder_records:
        kind = _record_file_kind(record)
        if kind not in _BUNDLE_KINDS:
            continue
        key = str(record.path_key or "").casefold()
        if key in excluded:
            skipped += 1
            continue
        mtime_ns = record_disk_mtime_ns(record)
        if mtime_ns <= 0:
            continue
        used.append(mtime_ns)
    if not used:
        return None
    mean_ns = int(sum(used) / len(used))
    date_text = override_date_text_from_mtime_ns(mean_ns)
    if not date_text:
        return None
    return FolderMeanOverride(
        override_date=date_text,
        used_count=len(used),
        excluded_count=skipped,
        mtime_ns=mean_ns,
    )


def mtime_override_tooltip(record: FileRecord) -> str:
    """Return the disk-vs-catalog date note for an MTO override, or empty.

    Args:
        record: Catalog file; override flags live on ``data``.
    """

    date_text = str(record.data.get("mtime_override_date") or "").strip()
    disk = format_file_save_date(
        int(record.data.get("disk_mtime_ns") or 0),
        with_time=True,
    )
    if record.data.get("mtime_override_applied"):
        reason = str(record.data.get("mtime_override_reason") or "")
        origin = {
            "code_a": "код A",
            "code_b": "код B",
            "code_c": "код C",
            "manual": "вручную",
            "folder_mean": "средняя по папке",
        }.get(reason, reason or "вручную")
        disk_bit = f"; на диске {disk}" if disk else ""
        catalog = date_text or format_file_save_date(
            int(record.data.get("mtime_ns") or 0)
        )
        return f"Дата каталога {catalog} ({origin}){disk_bit}"
    if record.data.get("mtime_override_stale"):
        return (
            "Замена даты не применяется: файл изменился "
            f"(сохранена {date_text or '—'})"
        )
    return ""


def folder_display_name(records: Iterable[FileRecord]) -> str:
    """Return the issued NN folder name for files in one tree node.

    Prefers the scanned ``transfer_name``, then the last segment of
    :func:`issued_package_dir`.

    Args:
        records: Files that belong to the same transfer/directory.

    Returns:
        Folder name such as ``05_рев.0-AN02_AGCC.287-9110-KSB1``, or ``""``.
    """

    materialized = [record for record in records]
    for record in materialized:
        name = str(record.data.get("transfer_name") or "").strip()
        if name:
            return name
    for record in materialized:
        package = issued_package_dir(record.path)
        if package:
            name = PureWindowsPath(package).name.strip()
            if name:
                return name
    return ""


def folder_transfer_sequence(records: Iterable[FileRecord]) -> int | None:
    """Return the issued transfer ``NN`` for files in one tree node.

    Prefers scanned ``transfer_sequence``, then leading digits of
    :func:`folder_display_name`. Folder ``рев.*`` text is ignored.

    Args:
        records: Files that belong to the same transfer/directory.

    Returns:
        Sequence number, or ``None`` when the folder has no ``NN``.
    """

    materialized = [record for record in records]
    found: list[int] = []
    for record in materialized:
        parsed = _int_or_none(record.data.get("transfer_sequence"))
        if parsed is not None:
            found.append(parsed)
    if found:
        return max(found)
    name = folder_display_name(materialized)
    if not name:
        return None
    return parse_transfer_folder(name, under_gate=True).sequence


def folder_tree_sort_key(
    records: Iterable[FileRecord],
    *,
    folder_key: str = "",
) -> tuple[int, int, str, str]:
    """Return a sort key for a documents-tree revision node.

    Numbered transfers sort by ``NN`` ascending (issuance order). Folders
    without a sequence sort after them. Duplicate ``NN`` keeps a stable
    order by display name, then ``folder_key``.

    Args:
        records: Files that belong to the same transfer/directory.
        folder_key: Tree grouping identity (``bundle.folder_key``).

    Returns:
        Tuple suitable for ``sorted`` after title and mark.
    """

    materialized = [record for record in records]
    sequence = folder_transfer_sequence(materialized)
    name = folder_display_name(materialized).casefold()
    key = (folder_key or "").casefold()
    if sequence is None:
        return (1, 0, name, key)
    return (0, sequence, name, key)


def _int_or_none(value: object) -> int | None:
    if value in (None, ""):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def folder_has_as_build(records: Iterable[FileRecord]) -> bool:
    """Return whether any record in the folder is canonical as-build.

    Args:
        records: Files that belong to the same transfer/directory.

    Returns:
        ``True`` when any record has canonical as-build
        (``transfer_is_as_build`` or a directory segment of the path).
    """

    return any(_record_is_as_build(record) for record in records)


def _record_is_as_build(record: FileRecord) -> bool:
    value = record.data.get("transfer_is_as_build")
    if isinstance(value, bool) and value:
        return True
    if isinstance(value, (int, str)) and str(value).isdigit() and int(value):
        return True
    if value and not isinstance(value, (bool, int, str)):
        return True
    return path_is_as_build(record.path)


def folder_has_mto(records: Iterable[FileRecord]) -> bool:
    """Return whether the folder contains an MTO XLSX catalog file.

    Args:
        records: Files that belong to the same transfer/directory.

    Returns:
        ``True`` when at least one record is ``mto_xlsx``.
    """

    return any(_record_file_kind(record) is FileKind.MTO_XLSX for record in records)


def folder_tree_label(
    records: Iterable[FileRecord],
    *,
    options: DocumentTreeLabelOptions | None = None,
    folder_name: str = "",
    review_status: str = "",
    is_current: bool = False,
    has_mto: bool | None = None,
    is_working: bool = False,
    is_annulled: bool = False,
    has_as_build: bool | None = None,
    mto_label: str = "",
) -> str:
    """Return the tree node text: revision, optional date, and extras.

    Args:
        records: Files that belong to the same transfer/directory.
        options: Which extras to append. ``None`` uses date-only defaults.
        folder_name: Issued NN folder name; derived from ``records`` when
            empty and ``show_folder`` is on.
        review_status: Compact Google/pipeline badge, shown only when
            ``show_review`` is on and the text is non-empty.
        is_current: Whether any file in the folder is overlay-current.
        has_mto: Whether the folder has an MTO XLSX. ``None`` derives it.
        is_working: Whether this folder is working (manual flag or ahead of send).
            The word ``рабочая`` is always appended when true so a mark cannot
            vanish behind the optional checkbox.
        is_annulled: Whether this folder is user-annulled. The word
            ``аннулирована`` is always appended when true. Wins over
            ``is_working`` if both are set.
        has_as_build: Whether any file is as-build. ``None`` derives it.
        mto_label: Compact MTO status badge from the caller. Shown first
            among extras when ``show_mto_status`` is on.

    Returns:
        ``01-AN02 (2024.08.16)`` by default, plus `` · ``-joined extras
        when the corresponding options are on. Working and annulled are
        not optional.
    """

    materialized = [record for record in records]
    chosen = options or DocumentTreeLabelOptions()
    revision = folder_revision_label(materialized)
    if chosen.show_date:
        date = folder_latest_save_date(materialized)
        head = f"{revision} ({date})" if date else revision
    else:
        head = revision
    extras: list[str] = []
    if chosen.show_mto_status and mto_label:
        extras.append(f"MTO: {mto_label}")
    if chosen.show_folder:
        name = folder_name.strip() or folder_display_name(materialized)
        if name:
            extras.append(name)
    if chosen.show_review:
        status = review_status.strip()
        if status:
            extras.append(status)
    if chosen.show_current and is_current:
        extras.append(CURRENT_MARKER)
    if chosen.show_mto:
        mto_present = folder_has_mto(materialized) if has_mto is None else has_mto
        if mto_present:
            extras.append(MTO_MARKER)
    if is_annulled:
        extras.append(ANNULLED_MARKER)
    elif is_working:
        extras.append(WORKING_MARKER)
    if chosen.show_as_build:
        as_build = (
            folder_has_as_build(materialized)
            if has_as_build is None
            else has_as_build
        )
        if as_build:
            extras.append(AS_BUILD_MARKER)
    if not extras:
        return head
    return f"{head} · {' · '.join(extras)}"


def working_folder_tooltip(*, is_working: bool, origin: str = "") -> str:
    """Return the working-folder hover text, or empty when the folder is official.

    Args:
        is_working: Whether this issued folder is treated as working.
        origin: ``manual`` for ``kit_working_flag``, ``auto`` when the
            filename rev is strictly above the last send/F, else empty.

    Returns:
        One Russian sentence, or ``""``.
    """

    if not is_working:
        return ""
    if origin == "manual":
        return WORKING_TOOLTIP_MANUAL
    return WORKING_TOOLTIP_AUTO


def annulled_folder_tooltip(*, is_annulled: bool) -> str:
    """Return the annulled-folder hover text, or empty when not marked.

    Args:
        is_annulled: Whether this issued folder is user-annulled.

    Returns:
        One Russian sentence, or ``""``.
    """

    if not is_annulled:
        return ""
    return ANNULLED_TOOLTIP


def folder_tree_tooltip(
    records: Iterable[FileRecord],
    *,
    folder_name: str = "",
    review_full: str = "",
    match_reason: str = "",
    f_label: str = "",
    is_current: bool = False,
    has_mto: bool | None = None,
    is_working: bool = False,
    working_origin: str = "",
    is_annulled: bool = False,
    has_as_build: bool | None = None,
    mto_status_label: str = "",
    mto_revision_text: str = "",
    mto_problems: str = "",
) -> str:
    """Return a hover dump of folder extras, independent of checkbox state.

    Args:
        records: Files that belong to the same transfer/directory.
        folder_name: Issued NN folder name; derived from ``records`` when empty.
        review_full: Full pipeline status label when a Google cycle matched.
        match_reason: ``trm`` / ``revision`` / ``date`` from ``kit_cycle``.
        f_label: Last matched F-stage label from the sheet.
        is_current: Overlay-current files in this folder.
        has_mto: MTO XLSX present; ``None`` derives it from ``records``.
        is_working: Working folder (manual flag or ahead-of-issue).
        working_origin: ``manual`` / ``auto``; ignored when not working.
        is_annulled: User-annulled issued folder.
        has_as_build: As-build files present; ``None`` derives it from ``records``.
        mto_status_label: Long MTO pipeline label for this folder.
        mto_revision_text: Filename revision of the folder's MTO.
        mto_problems: Comma-joined ``problem_kinds`` from the matrix cell.

    Returns:
        Multiline tooltip, or empty string when nothing extra is known.
    """

    materialized = [record for record in records]
    lines: list[str] = []
    name = folder_name.strip() or folder_display_name(materialized)
    if name:
        lines.append(f"Папка: {name}")
    date = folder_latest_save_date(materialized)
    if date:
        lines.append(f"Сохранено: {date}")
    if review_full:
        lines.append(f"Статус: {review_full}")
        reason = _match_reason_label(match_reason)
        if reason:
            lines.append(reason)
        if f_label:
            lines.append(f"F: {f_label}")
    mto_present = folder_has_mto(materialized) if has_mto is None else has_mto
    if is_current:
        lines.append("Текущий состав")
    if mto_present:
        lines.append("В папке есть MTO")
    for record in materialized:
        extra = mtime_override_tooltip(record)
        if extra:
            lines.append(extra)
    working_line = working_folder_tooltip(
        is_working=is_working and not is_annulled, origin=working_origin
    )
    if working_line:
        lines.append(working_line)
    annulled_line = annulled_folder_tooltip(is_annulled=is_annulled)
    if annulled_line:
        lines.append(annulled_line)
    as_build = (
        folder_has_as_build(materialized) if has_as_build is None else has_as_build
    )
    if as_build:
        lines.append("As-build")
    if mto_status_label:
        lines.append(f"Статус MTO: {mto_status_label}")
    if mto_revision_text:
        lines.append(f"Рев. MTO: {mto_revision_text}")
    if mto_problems:
        lines.append(f"Проблемы: {mto_problems}")
    return "\n".join(lines)


def _match_reason_label(reason: str) -> str:
    mapping = {
        "trm": "Сопоставлено с Google по TRM",
        "revision": "Сопоставлено с Google по ревизии",
        "date": "Сопоставлено с Google по дате (±2 дня)",
    }
    return mapping.get((reason or "").strip().casefold(), "")


def folder_revision_rank(
    records: Iterable[FileRecord],
) -> tuple[int, int, str]:
    """Return the comparable rank of the highest filename revision.

    Args:
        records: Files that belong to the same transfer/directory.

    Returns:
        ``revision_rank`` tuple; missing revisions use the overlay missing
        sentinel (below ``V`` / ``S`` and every numeric token).
    """

    materialized = [record for record in records]
    if not materialized:
        return revision_rank(None, None)
    return max(
        revision_rank(*_file_revision_tokens(record)) for record in materialized
    )


def document_row_key(record: FileRecord) -> tuple[str, str, str, str]:
    """Return the grouping key used for catalog table rows.

    Args:
        record: Persisted catalog file.

    Returns:
        ``(title, mark, folder_key, stem)`` all casefolded.
    """

    title = _data_str(record, "title").casefold()
    mark = _data_str(record, "mark").casefold()
    stem = _data_str(record, "core_stem").casefold() or _record_name(record).casefold()
    return (title, mark, record_folder_key(record), stem)


def editable_pairs_with_pdf(pdf_dir: str, source_dir: str) -> bool:
    """Return whether an editable may be glued to a PDF by folder location.

    Pairing is allowed when both files share a directory, when the source
    folder is named ``DWG`` inside the PDF folder, or when ``DWG`` is a
    sibling of the PDF folder (one level above the PDF directory).

    Args:
        pdf_dir: Directory containing the PDF.
        source_dir: Directory containing the editable file.

    Returns:
        ``True`` when the editable belongs with that PDF.
    """

    pdf_norm = os.path.normcase(os.path.normpath(pdf_dir))
    source_norm = os.path.normcase(os.path.normpath(source_dir))
    if pdf_norm == source_norm:
        return True
    source_path = Path(source_dir)
    if source_path.name.casefold() != "dwg":
        return False
    parent_norm = os.path.normcase(os.path.normpath(str(source_path.parent)))
    if parent_norm == pdf_norm:
        return True
    pdf_parent_norm = os.path.normcase(os.path.normpath(str(Path(pdf_dir).parent)))
    return parent_norm == pdf_parent_norm


def is_rd_catalog_record(record: FileRecord) -> bool:
    """Return whether the record belongs to the RD source.

    Args:
        record: Persisted catalog file.

    Returns:
        ``True`` when ``record.source`` is RD.
    """

    return record.source is SourceKind.RD


def _has_title_mark(record: FileRecord) -> bool:
    return bool(_data_str(record, "title") and _data_str(record, "mark"))


def bundle_documents(
    records: Iterable[FileRecord],
    *,
    detected_current_ids: set[int] | frozenset[int] = frozenset(),
) -> list[DocumentBundle]:
    """Group RD PDF and editable files into catalog tree rows.

    PDFs and editables are grouped by title, mark, transfer/directory, and
    stem. Records without a parsed title and mark (names outside the AGCC
    mask) are omitted. An editable is attached to a PDF in the same group
    only when :func:`editable_pairs_with_pdf` succeeds; otherwise it forms
    a bundle with ``pdf=None``. Extra PDFs in a group are ignored after a
    primary is chosen (detected-current if any, else newest mtime).

    Args:
        records: Persisted catalog files from any source.
        detected_current_ids: Overlay-current file ids.

    Returns:
        Bundles sorted by title, mark, revision rank, and stem.
    """

    pdfs_by_key: dict[tuple[str, str, str, str], list[FileRecord]] = defaultdict(list)
    editables_by_key: dict[tuple[str, str, str, str], list[FileRecord]] = (
        defaultdict(list)
    )
    for record in records:
        if not is_rd_catalog_record(record):
            continue
        if not _has_title_mark(record):
            continue
        kind = _record_file_kind(record)
        if kind not in _BUNDLE_KINDS:
            continue
        key = document_row_key(record)
        if kind is FileKind.PDF:
            pdfs_by_key[key].append(record)
        else:
            editables_by_key[key].append(record)

    bundles: list[DocumentBundle] = []
    for key in set(pdfs_by_key) | set(editables_by_key):
        pdfs = pdfs_by_key.get(key, [])
        editables = editables_by_key.get(key, [])
        primary = _pick_primary_pdf(pdfs, detected_current_ids) if pdfs else None
        paired: list[FileRecord] = []
        unpaired: list[FileRecord] = []
        pdf_dir = str(Path(primary.path).parent) if primary is not None else ""
        for editable in editables:
            source_dir = str(Path(editable.path).parent)
            if primary is not None and editable_pairs_with_pdf(pdf_dir, source_dir):
                paired.append(editable)
            else:
                unpaired.append(editable)
        if primary is not None:
            bundles.append(
                _make_bundle(primary, paired, detected_current_ids)
            )
        if unpaired:
            bundles.append(
                _make_bundle(None, unpaired, detected_current_ids)
            )

    bundles.sort(
        key=lambda bundle: (
            bundle.title.casefold(),
            bundle.mark.casefold(),
            revision_rank(*_bundle_revision_tokens(bundle)),
            bundle.core_stem.casefold(),
        )
    )
    return bundles


def _data_str(record: FileRecord, key: str) -> str:
    value = record.data.get(key)
    return str(value) if value not in (None, "") else ""


def _record_name(record: FileRecord) -> str:
    return _data_str(record, "name") or Path(record.path).name


def _record_file_kind(record: FileRecord) -> FileKind | None:
    raw = record.data.get("file_kind")
    if raw is None:
        return None
    try:
        return FileKind(str(raw))
    except ValueError:
        return None


def _optional_token(value: object) -> str | None:
    if value in (None, ""):
        return None
    return str(value)


def _file_revision_tokens(record: FileRecord) -> tuple[str | None, str | None]:
    return (
        _optional_token(record.data.get("revision")),
        _optional_token(record.data.get("appendix")),
    )


def _transfer_revision_tokens(record: FileRecord) -> tuple[str | None, str | None]:
    return (
        _optional_token(record.data.get("transfer_revision")),
        _optional_token(record.data.get("transfer_appendix")),
    )


def _revision_tokens(record: FileRecord) -> tuple[str | None, str | None]:
    file_revision, file_appendix = _file_revision_tokens(record)
    if file_revision:
        return (file_revision, file_appendix)
    return _transfer_revision_tokens(record)


def _discipline(record: FileRecord) -> str:
    block = _data_str(record, "discipline_block")
    if "-" in block:
        return block.split("-", 1)[0]
    return block


def _pick_primary_pdf(
    pdfs: list[FileRecord],
    detected_current_ids: set[int] | frozenset[int],
) -> FileRecord:
    current = [record for record in pdfs if record.id in detected_current_ids]
    pool = current or pdfs
    return max(pool, key=lambda record: int(record.data.get("mtime_ns") or 0))


def _sorted_editables(editables: Iterable[FileRecord]) -> tuple[FileRecord, ...]:
    return tuple(sorted(editables, key=lambda record: (record.path_key, record.id)))


def _make_bundle(
    pdf: FileRecord | None,
    editables: list[FileRecord],
    detected_current_ids: set[int] | frozenset[int],
) -> DocumentBundle:
    representative = pdf or editables[0]
    title = _data_str(representative, "title")
    mark = _data_str(representative, "mark")
    core_stem = _data_str(representative, "core_stem") or _record_name(representative)
    return DocumentBundle(
        title=title,
        mark=mark,
        revision_label=file_revision_label(representative) or NO_REVISION_LABEL,
        folder_key=record_folder_key(representative),
        core_stem=core_stem,
        discipline=_discipline(representative),
        pdf=pdf,
        editables=_sorted_editables(editables),
        is_current=pdf is not None and pdf.id in detected_current_ids,
    )


def _bundle_revision_tokens(
    bundle: DocumentBundle,
) -> tuple[str | None, str | None]:
    representative = bundle.pdf or (bundle.editables[0] if bundle.editables else None)
    if representative is None:
        return (None, None)
    return _revision_tokens(representative)
