"""Generate AGCC MTO xlsx files from the customer PI pickle.

The catalog lives under ``rd_catalog/База заказчика/АвтоМто``. Files are
named like production MTO workbooks so ``load_mto_rows`` / robot compare
can read them. This module does not scan UNC and does not import Qt.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import tempfile
from collections import Counter, defaultdict
from collections.abc import Callable, Sequence
from dataclasses import dataclass, replace
from datetime import datetime, timezone
from itertools import combinations
from pathlib import Path
from typing import Any, Literal

import openpyxl

from rd_catalog.customer_pi import (
    DUMP_DIR,
    IDX_RD_REV,
    IDX_SPEC,
    CustomerPiRecord,
    CustomerPiStore,
    load_customer_pi,
    normalize_text,
)
from rd_catalog.kits import (
    format_revision,
    kit_identity_key,
    parse_sheet_revision,
    revision_matches_rd,
)
from rd_catalog.models import make_path_key
from rd_catalog.parse import normalize_unicode_dashes
from rd_catalog.overlay import revision_rank
from rd_catalog.mto_diff import (
    CanonicalMtoDocument,
    CanonicalMtoRow,
    RowLoader,
    canonicalize_mto_rows,
    load_canonical_mto,
    semantic_fingerprint,
)
from utils.file_name_converts import (
    AgccFilenamePatterns,
    parse_agcc_mto_xlsx_revision_for_chain,
)

AUTO_MTO_DIR_NAME = "АвтоМто"
AUTO_MTO_DIR = DUMP_DIR / AUTO_MTO_DIR_NAME
MANIFEST_NAME = "_manifest.json"
FORMAT_ID = "rd_catalog.auto_mto"
SCHEMA_VERSION = 1
AUTO_MTO_COMPARE_ALGORITHM_VERSION = 2
AUTO_MTO_EXACT_STATUS = "четкое"
AUTO_MTO_SOFT_STATUS = "ПоКоду и Кол-ву"
AUTO_MTO_COMPARE_STATUS_HEADER = "Сверка Авто МТО"
AUTO_MTO_COMPARE_STATUS_TOOLTIP = (
    "Сверка содержимого Авто МТО с последней MTO РД "
    "(если файл выбран вручную — с ним)."
)
LANGUAGE = "RU"
SHEET_NAME = "Спецификация"
MTO_DISCIPLINE = "MTO"
CONVERTIBLE_DISCIPLINES = frozenset({"MTO", "BOM", "BOE", "DS"})
_CONVERT_DISC_RE = re.compile(r"\.(MTO|BOM|BOE|DS)-", re.I)

STATUS_QUEUED = "queued"
STATUS_WRITING = "writing"
STATUS_WRITTEN = "written"
STATUS_SKIPPED = "skipped"
STATUS_NOT_MTO = "not_mto"
STATUS_EMPTY = "empty"
PAINT_OUTSIDE_KITS = "outside_kits"

STATUS_LABELS: dict[str, str] = {
    STATUS_QUEUED: "ожидает",
    STATUS_WRITING: "пишется",
    STATUS_WRITTEN: "записано",
    STATUS_SKIPPED: "пропуск",
    STATUS_NOT_MTO: "не MTO",
    STATUS_EMPTY: "нет имени",
}

STATUS_COLORS: dict[str, str] = {
    STATUS_QUEUED: "#e8eaed",
    STATUS_WRITING: "#f9ab00",
    STATUS_WRITTEN: "#137333",
    STATUS_SKIPPED: "#c5221f",
    STATUS_NOT_MTO: "#9aa0a6",
    STATUS_EMPTY: "#c5221f",
    # Faded green: PI title+mark is not a visible Комплекты row
    # (other task, banned, non-kit stem). Distinct from written.
    PAINT_OUTSIDE_KITS: "#87c99a",
}

HEADER_ROW: tuple[str, ...] = (
    "TAG №",
    "п.п.№",
    "Наименование и техническая характеристика",
    "Тип, марка, обозначение документа",
    "Код продукции",
    "Поставщик",
    "Ед. измерения",
    "Кол.",
    "Масса ед.",
    "Примечание",
)

ProgressCallback = Callable[[int, int, str], None]
CancelCallback = Callable[[], bool]


@dataclass(frozen=True, slots=True)
class AutoMtoFile:
    """One generated (or planned) MTO workbook for a PI spec."""

    title: str
    mark: str
    spec: str
    rd_revision: str
    relpath: str
    row_count: int
    fingerprint: str
    extras: tuple[str, ...] = ()
    source_specs: tuple[str, ...] = ()

    @property
    def kit_key(self) -> tuple[str, str]:
        return kit_identity_key(self.title, self.mark)


@dataclass(frozen=True, slots=True)
class AutoMtoCompareResult:
    """Graded AutoMTO-to-RD comparison result."""

    match_kind: Literal["single", "composite", "no_match", "not_compared"]
    members: tuple[AutoMtoFile, ...]
    rd_rows: int
    auto_rows: int
    combinations_checked: int
    error: str = ""
    grade: Literal["exact", "soft", "none"] = "none"

    @property
    def matched(self) -> bool:
        """Return whether a single or composite content match was found."""

        return self.match_kind in {"single", "composite"}

    @property
    def content_grade(self) -> Literal["exact", "soft", "none"]:
        """Return the content grade, defaulting single/composite to exact."""

        if self.match_kind not in {"single", "composite"}:
            return "none"
        return "soft" if self.grade == "soft" else "exact"

    def cell_text(self, fallback: str = "") -> str:
        """Return the revision label to show in an AutoMTO cell."""

        if self.match_kind == "composite" and self.members:
            return " + ".join(member.rd_revision for member in self.members)
        if self.members:
            return self.members[0].rd_revision
        return fallback

    def tooltip_lines(self) -> tuple[str, ...]:
        """Return Russian notes about the graded content match."""

        if self.match_kind == "single" and self.members:
            member = self.members[0]
            if self.content_grade == "soft":
                return (
                    f"Содержимое: {AUTO_MTO_SOFT_STATUS}, 1 файл "
                    f"({member.rd_revision}), строк РД {self.rd_rows} / "
                    f"Авто {self.auto_rows}.",
                    "Совпали код и количество. Теги или ед. изм. разошлись.",
                )
            return (
                "Содержимое: чёткое совпадение (код, количество, теги, "
                f"ед. изм.), 1 файл ({member.rd_revision}), строк {self.rd_rows}.",
                "NAME / VENDOR / TYPE_MARK не входят в сверку.",
            )
        if self.match_kind == "composite" and self.members:
            revisions = " + ".join(member.rd_revision for member in self.members)
            names = ", ".join(Path(member.relpath).name for member in self.members)
            if self.content_grade == "soft":
                return (
                    f"Содержимое: {AUTO_MTO_SOFT_STATUS}, сумма "
                    f"{len(self.members)} файлов ({revisions}), строк РД "
                    f"{self.rd_rows} / Авто {self.auto_rows}.",
                    f"Состав: {names}.",
                    "Совпали код и количество. Теги или ед. изм. разошлись.",
                )
            return (
                f"Содержимое: чёткое совпадение суммы {len(self.members)} файлов "
                f"({revisions}), строк {self.rd_rows}.",
                f"Состав: {names}.",
                "NAME / VENDOR / TYPE_MARK не входят в сверку.",
            )
        if self.match_kind == "no_match":
            extra = f" Непрочитанные файлы: {self.error}." if self.error else ""
            return (
                "Содержимое: не совпало по коду и количеству ни для одного "
                "файла, ни для допустимой суммы файлов."
                + extra,
                f"Проверено составных вариантов: {self.combinations_checked}.",
            )
        detail = f" — {self.error}" if self.error else "."
        return (f"Содержимое: сверка не завершена{detail}",)


@dataclass(frozen=True, slots=True)
class AutoMtoCompareStatus:
    """Scannable AutoMTO-vs-RD content-compare label for tables."""

    kind: str
    text: str
    tooltip: str = ""

    @property
    def paint_match(self) -> bool | None:
        """Return True/False for green/yellow fill, or None to leave unpainted."""

        if self.kind in {"matched", "composite", "soft", "rev_match"}:
            return True
        if self.kind in {"no_match", "stale"}:
            return False
        return None


def auto_mto_rd_path_key(path: str) -> str:
    """Return a case-insensitive path identity for AutoMTO cache matching.

    Normalizes Unicode dashes (including U+2010) and Windows slashes so the
    same UNC file is not treated as a different workbook.

    Args:
        path: RD MTO path from the scan, export selection, or cache.

    Returns:
        Empty string when ``path`` is empty; otherwise ``make_path_key``
        after dash normalization.
    """

    text = str(path or "").strip()
    if not text:
        return ""
    return make_path_key(normalize_unicode_dashes(text))


def auto_mto_rd_paths_match(left: str, right: str) -> bool:
    """Return whether two RD MTO paths name the same workbook.

    Args:
        left: First path.
        right: Second path.

    Returns:
        True when both sides are non-empty and share
        :func:`auto_mto_rd_path_key`.
    """

    if not left or not right:
        return False
    return auto_mto_rd_path_key(left) == auto_mto_rd_path_key(right)


def _compare_target_line(rd_path: str, *, pinned: bool) -> str:
    name = Path(rd_path).name if rd_path else ""
    if pinned:
        return f"с ручным выбором MTO: {name or '—'}"
    if name:
        return f"с последней MTO РД: {name}"
    return "с последней MTO РД: нет файла"


def auto_mto_compare_status(
    *,
    files: Sequence[AutoMtoFile] = (),
    rd_path: str = "",
    rd_revision: str = "",
    comparison: AutoMtoCompareResult | None = None,
    comparison_rd_path: str = "",
    pinned: bool = False,
) -> AutoMtoCompareStatus:
    """Return a short status for AutoMTO vs the selected RD MTO file.

    Args:
        files: Customer-PI AutoMTO files for the kit.
        rd_path: Latest overlay RD MTO, or the pinned file when set.
        rd_revision: Filename revision of ``rd_path``.
        comparison: Cached exact compare, if any.
        comparison_rd_path: RD path the cache was computed against.
        pinned: True when ``rd_path`` is a manual export pin.

    Returns:
        Table text, kind, and tooltip. Does not read workbooks.
    """

    target = _compare_target_line(rd_path, pinned=pinned)
    tips_prefix = [target]
    if rd_revision:
        tips_prefix.append(f"рев. MTO РД: {rd_revision}")
    if not files:
        return AutoMtoCompareStatus(
            kind="no_auto",
            text="нет в ПИ",
            tooltip="\n".join(
                [*tips_prefix, "Нет спецификации MTO в базе заказчика."]
            ),
        )
    if not rd_path:
        return AutoMtoCompareStatus(
            kind="no_rd",
            text="нет MTO РД",
            tooltip="\n".join([*tips_prefix, "Нет файла MTO РД для сверки."]),
        )
    cached_mismatch = bool(
        comparison is not None
        and comparison_rd_path
        and not auto_mto_rd_paths_match(comparison_rd_path, rd_path)
    )
    if cached_mismatch:
        extra = list(comparison.tooltip_lines()) if comparison else []
        return AutoMtoCompareStatus(
            kind="stale",
            text="другой файл",
            tooltip="\n".join(
                [
                    *tips_prefix,
                    "Кэш сверки относится к другому файлу: "
                    f"{Path(comparison_rd_path).name}.",
                    "Сверьте заново с текущим файлом РД.",
                    *extra,
                ]
            ),
        )
    if comparison is not None:
        tips = [*tips_prefix, *comparison.tooltip_lines()]
        if comparison.match_kind in {"single", "composite"}:
            kind, label = _comparison_status_text(comparison)
            return AutoMtoCompareStatus(
                kind=kind,
                text=label,
                tooltip="\n".join(tips),
            )
        if comparison.match_kind == "no_match":
            return AutoMtoCompareStatus(
                kind="no_match",
                text="не совпало",
                tooltip="\n".join(tips),
            )
        text = "ошибка" if comparison.error else "не сверялось"
        return AutoMtoCompareStatus(
            kind="not_compared",
            text=text,
            tooltip="\n".join(tips),
        )
    auto = pick_auto_mto_file(files, rd_revision)
    match = (
        revision_texts_match(auto.rd_revision, rd_revision)
        if auto is not None
        else None
    )
    if match is True:
        return AutoMtoCompareStatus(
            kind="rev_match",
            text="рев. совпала",
            tooltip="\n".join(
                [
                    *tips_prefix,
                    "Ревизия в имени совпала. Содержимое ещё не сверялось.",
                    "Сверка содержимого идёт в фоне.",
                ]
            ),
        )
    return AutoMtoCompareStatus(
        kind="not_compared",
        text="не сверялось",
        tooltip="\n".join(
            [
                *tips_prefix,
                    "Точное содержимое не сверялось.",
                    "Сверка содержимого идёт в фоне.",
            ]
        ),
    )


def needs_auto_mto_compare(status: AutoMtoCompareStatus) -> bool:
    """Return whether background AutoMTO compare should run for this status.

    Args:
        status: Kit-level «Сверка Авто МТО» label.

    Returns:
        True for ``not_compared``, ``rev_match``, or ``stale``.
    """

    return status.kind in {"not_compared", "rev_match", "stale"}


def _comparison_status_text(result: AutoMtoCompareResult) -> tuple[str, str]:
    """Return GUI kind and cell text for a single/composite compare."""

    revision = result.cell_text()
    if result.content_grade == "soft":
        prefix = AUTO_MTO_SOFT_STATUS
        kind = "soft"
    else:
        prefix = AUTO_MTO_EXACT_STATUS
        kind = "composite" if result.match_kind == "composite" else "matched"
    if result.match_kind == "composite":
        text = f"{prefix} · сумма {revision}" if revision else prefix
    else:
        text = f"{prefix} · {revision}" if revision else prefix
    return kind, text


@dataclass(frozen=True, slots=True)
class AutoMtoRebuildResult:
    """Counts after rewriting the АвтоМто folder."""

    dest: Path
    written: tuple[AutoMtoFile, ...]
    skipped: tuple[str, ...]
    removed: tuple[str, ...]
    manifest_path: Path


@dataclass(frozen=True, slots=True)
class SpecProgressRow:
    """One unique PI «Спецификация» value and how catalog build treats it."""

    spec: str
    rd_revision: str
    discipline: str
    title: str
    mark: str
    row_count: int
    status: str
    note: str = ""


def default_auto_mto_dir() -> Path:
    """Return the packaged АвтоМто catalog root."""

    return AUTO_MTO_DIR


def _is_primary_mto_spec(spec: str) -> bool:
    return spec_as_mto_stem(spec).upper().endswith("MTO-0001")


def spec_as_mto_stem(spec: str) -> str:
    """Rewrite BOM/BOE/DS in the AGCC stem to MTO.

    Args:
        spec: PI «Спецификация» stem (``….BOM-0001``, ``….DS-0003``, …).

    Returns:
        The same stem with the first discipline token replaced by ``MTO``.
        Stems without ``.MTO/.BOM/.BOE/.DS-`` are returned unchanged.
    """

    text = normalize_text(spec)
    if not text:
        return ""
    return _CONVERT_DISC_RE.sub(".MTO-", text, count=1)


def _is_convertible_spec(spec: str, discipline: str) -> bool:
    if discipline in CONVERTIBLE_DISCIPLINES:
        return True
    return bool(_CONVERT_DISC_RE.search(normalize_text(spec)))


def format_store_summary(store: CustomerPiStore) -> str:
    """Return a short Russian summary of a loaded PI dump."""

    counts = store.meta.get("counts") if isinstance(store.meta.get("counts"), dict) else {}
    discipline = counts.get("discipline") if isinstance(counts.get("discipline"), dict) else {}
    pickle_path = store.meta.get("pickle_path") or "—"
    source_path = store.meta.get("source_path") or "—"
    parsed_at = store.meta.get("parsed_at") or "—"
    layout = store.meta.get("layout") or "—"
    source_n_cols = store.meta.get("source_n_cols")
    layout_line = f"Шаблон xlsb: {layout}"
    if source_n_cols:
        layout_line += f" ({source_n_cols} столбцов в файле)"
    return (
        f"Pickle: {pickle_path}\n"
        f"Собран: {parsed_at}\n"
        f"Источник xlsb: {source_path}\n"
        f"{layout_line}\n"
        f"Строк BCC: {len(store)}\n"
        f"Спецификаций: {counts.get('specs', '—')}\n"
        f"MTO / BOM / DS: {discipline.get('MTO', 0)} / "
        f"{discipline.get('BOM', 0)} / {discipline.get('DS', 0)}"
    )


def auto_mto_filename(spec: str, rd_revision: str, *, language: str = LANGUAGE) -> str:
    """Build ``{spec}_{rd_revision}_RU.xlsx`` and validate the AGCC tail.

    ``spec`` must already be the MTO stem (call :func:`spec_as_mto_stem` for
    BOM/BOE/DS sources).

    Args:
        spec: AGCC stem without revision, with ``.MTO-``.
        rd_revision: PI «№ ревизии РД» token (``01``, ``01-AN01``).
        language: AGCC language token; production MTO uses ``RU``.

    Returns:
        File name including ``.xlsx``.

    Raises:
        ValueError: Empty parts, unsafe characters, or a name the AGCC
            parser rejects.
    """

    stem = normalize_text(spec)
    rev = normalize_text(rd_revision)
    lang = normalize_text(language) or LANGUAGE
    if not stem:
        raise ValueError("empty specification stem")
    if not rev:
        raise ValueError("empty RD revision")
    illegal = '<>:"/\\|?*'
    if any(ch in stem or ch in rev or ch in lang for ch in illegal):
        raise ValueError(f"unsafe characters in spec/rev: {stem!r} {rev!r}")
    name = f"{stem}_{rev}_{lang}.xlsx"
    parse_agcc_mto_xlsx_revision_for_chain(name)
    return name


def auto_mto_relpath(
    title: str,
    mark: str,
    spec: str,
    rd_revision: str,
) -> Path:
    """Return ``title/MARK/filename.xlsx`` relative to the catalog root."""

    title_text = normalize_text(title)
    mark_text = normalize_text(mark)
    name = auto_mto_filename(spec_as_mto_stem(spec), rd_revision)
    if title_text and mark_text:
        return Path(title_text) / mark_text / name
    if title_text:
        return Path(title_text) / name
    return Path(name)


def _pi_canonical_rows(records: Sequence[CustomerPiRecord]):
    """Canonical rows as the MTO xlsx loader would see them (``get_tag``)."""

    from tags.tag_parser import get_tag

    mapped: list[dict[str, object]] = []
    for record in records:
        fields = dict(record.as_mto_fields())
        fields["TAGS"] = get_tag(str(fields.get("TAGS") or ""))
        mapped.append(fields)
    return canonicalize_mto_rows(mapped)


def _data_cells(record: CustomerPiRecord, index: int) -> list[object]:
    fields = record.as_mto_fields()
    return [
        fields["TAGS"],
        index,
        fields["NAME"],
        fields["TYPE_MARK"],
        fields["CODE"],
        fields["VENDOR"],
        fields["UNITS"],
        fields["VALUES"],
        "",
        "",
    ]


def mto_sheet_row(record: CustomerPiRecord, index: int) -> tuple[object, ...]:
    """Return one MTO workbook row in ``HEADER_ROW`` order.

    Args:
        record: PI dump row.
        index: 1-based «п.п.№» as written to АвтоМто xlsx.

    Returns:
        Ten cells matching :data:`HEADER_ROW`.
    """

    return tuple(_data_cells(record, index))


def write_auto_mto_workbook(
    path: str | Path,
    records: Sequence[CustomerPiRecord],
) -> str:
    """Write one ``Спецификация`` workbook and return its fingerprint.

    Args:
        path: Destination ``.xlsx``.
        records: PI rows for this spec (order preserved).

    Returns:
        Semantic fingerprint of the canonical material rows.
    """

    dest = Path(path)
    dest.parent.mkdir(parents=True, exist_ok=True)
    workbook = openpyxl.Workbook()
    try:
        worksheet = workbook.active
        worksheet.title = SHEET_NAME
        worksheet.append(list(HEADER_ROW))
        for index, record in enumerate(records, start=1):
            worksheet.append(_data_cells(record, index))
        fd, tmp_name = tempfile.mkstemp(
            prefix=dest.stem + ".",
            suffix=".tmp.xlsx",
            dir=str(dest.parent),
        )
        os.close(fd)
        tmp_path = Path(tmp_name)
        try:
            workbook.save(tmp_path)
        except Exception:
            tmp_path.unlink(missing_ok=True)
            raise
    finally:
        workbook.close()
    os.replace(tmp_path, dest)
    rows = _pi_canonical_rows(records)
    return semantic_fingerprint(rows)


def _group_mto_records(
    store: CustomerPiStore,
) -> tuple[dict[tuple[str, str], list[int]], list[str]]:
    grouped: dict[tuple[str, str], list[int]] = defaultdict(list)
    skipped: list[str] = []
    for index in range(len(store)):
        record = store.record(index)
        spec = record.spec
        rev = record.rd_revision
        if not _is_convertible_spec(spec, record.discipline):
            continue
        if not spec or not rev:
            skipped.append(f"row {record.excel_row}: empty spec/rev")
            continue
        grouped[(spec, rev)].append(index)
    return dict(grouped), skipped


def _output_priority(orig_spec: str, discipline: str) -> tuple[int, int, str]:
    output = spec_as_mto_stem(orig_spec)
    native = 0 if normalize_text(orig_spec) == output else 1
    order = {"MTO": 0, "BOM": 1, "BOE": 2, "DS": 3}
    return (native, order.get(discipline, 9), orig_spec)


def _kit_file_sort_key(item: AutoMtoFile) -> tuple:
    rev, appendix = parse_sheet_revision(item.rd_revision)
    rank = revision_rank(rev, appendix)
    return (
        0 if _is_primary_mto_spec(item.spec) else 1,
        -rank[0],
        -rank[1],
        item.spec,
        item.rd_revision,
    )


def pick_auto_mto_file(
    files: Sequence[AutoMtoFile],
    rd_revision: str = "",
) -> AutoMtoFile | None:
    """Return the catalog file matching ``rd_revision``, else the primary.

    Primary is ``.MTO-0001`` with the highest filename revision. A later
    DS/BOM file with another revision is chosen when it matches ``rd_revision``.

    Args:
        files: Planned files for one title+mark kit.
        rd_revision: RD MTO filename revision to match, or empty.

    Returns:
        One file, or ``None`` when ``files`` is empty.
    """

    ordered = sorted(files, key=_kit_file_sort_key)
    if not ordered:
        return None
    wanted = normalize_text(rd_revision)
    if wanted:
        for item in ordered:
            if revision_texts_match(item.rd_revision, wanted):
                return item
    return ordered[0]


def format_auto_mto_cell_text(
    files: Sequence[AutoMtoFile],
    rd_revision: str = "",
) -> str:
    """Return the Комплекты / heatmap «Авто МТО» cell label.

    The shown revision is ``pick_auto_mto_file``: match ``rd_revision`` if
    present, else ``.MTO-0001`` with the highest rank. When the kit has more
    than one customer file, append `` (N)``.

    Args:
        files: Planned files for one title+mark kit.
        rd_revision: RD MTO filename revision to match, or empty.

    Returns:
        Revision text, optionally with a file count, or ``""``.
    """

    auto = pick_auto_mto_file(files, rd_revision)
    if auto is None or not auto.rd_revision:
        return ""
    if len(files) > 1:
        return f"{auto.rd_revision} ({len(files)})"
    return auto.rd_revision


def _kit_identity(record: CustomerPiRecord) -> tuple[str, str]:
    title = record.title or record.parsed_title
    mark = record.mark or record.parsed_mark
    return title, mark


def _classify_spec_status(
    spec: str,
    rd_revision: str,
    discipline: str,
    title: str,
    mark: str,
    dest: Path | None,
) -> tuple[str, str]:
    """Return ``(status, note)`` for one unique specification stem."""

    if not spec or not rd_revision:
        return STATUS_EMPTY, "пустая спецификация или ревизия"
    if not _is_convertible_spec(spec, discipline):
        return STATUS_NOT_MTO, ""
    output = spec_as_mto_stem(spec)
    try:
        name = auto_mto_filename(output, rd_revision)
    except ValueError as exc:
        return STATUS_SKIPPED, str(exc).splitlines()[0]
    note = f"как MTO → {name}" if output != spec else name
    if dest is not None:
        rel = auto_mto_relpath(title, mark, output, rd_revision)
        if (dest / rel).is_file():
            return STATUS_WRITTEN, note
    return STATUS_QUEUED, note


def list_spec_progress_rows(
    store: CustomerPiStore,
    dest: str | Path | None = None,
) -> tuple[SpecProgressRow, ...]:
    """Return unique «Спецификация» values with row counts and write status.

    Args:
        store: Loaded PI dump.
        dest: АвтоМто root; when set, MTO files already on disk are ``written``.

    Returns:
        One row per unique spec stem (empty stem included), sorted by name.
    """

    groups: dict[str, list[int]] = defaultdict(list)
    for index, raw in enumerate(store.rows):
        groups[normalize_text(raw[IDX_SPEC])].append(index)
    dest_root = Path(dest) if dest else None
    result: list[SpecProgressRow] = []
    for spec in sorted(groups, key=lambda text: text.casefold() or "\uffff"):
        indexes = groups[spec]
        first = store.record(indexes[0])
        revisions: list[str] = []
        seen: set[str] = set()
        for index in indexes:
            rev = normalize_text(store.rows[index][IDX_RD_REV])
            if rev not in seen:
                seen.add(rev)
                revisions.append(rev)
        rd_revision = ", ".join(revisions)
        title, mark = _kit_identity(first)
        status, note = _classify_spec_status(
            spec,
            revisions[0] if revisions else "",
            first.discipline,
            title,
            mark,
            dest_root,
        )
        result.append(
            SpecProgressRow(
                spec=spec,
                rd_revision=rd_revision,
                discipline=first.discipline or "other",
                title=title,
                mark=mark,
                row_count=len(indexes),
                status=status,
                note=note,
            )
        )
    return tuple(result)


def overlay_written_specs(
    rows: Sequence[SpecProgressRow],
    written: Sequence[AutoMtoFile] | Sequence[tuple[str, str]],
) -> tuple[SpecProgressRow, ...]:
    """Mark PI specs whose output stem+revision was written.

    Aliased BOM/DS that share the same ``{stem}_{rev}_RU.xlsx`` as a native
    MTO also become ``written``: the revision is in the folder, extra rows
    were not merged.
    """

    done: set[tuple[str, str]] = set()
    for item in written:
        if isinstance(item, AutoMtoFile):
            done.add((item.spec, normalize_text(item.rd_revision)))
        else:
            stem, rev = item
            done.add((spec_as_mto_stem(stem) or stem, normalize_text(rev)))
    result: list[SpecProgressRow] = []
    for row in rows:
        output = spec_as_mto_stem(row.spec) or row.spec
        rev = normalize_text(row.rd_revision.split(",")[0])
        if (output, rev) in done and row.status in {STATUS_QUEUED, STATUS_WRITING}:
            result.append(replace(row, status=STATUS_WRITTEN, note=row.note))
        else:
            result.append(row)
    return tuple(result)


def format_spec_progress_summary(rows: Sequence[SpecProgressRow]) -> str:
    """Return a one-line Russian tally of spec statuses and PI row counts."""

    counts = Counter(row.status for row in rows)
    written = counts.get(STATUS_WRITTEN, 0)
    queued = counts.get(STATUS_QUEUED, 0) + counts.get(STATUS_WRITING, 0)
    skipped = counts.get(STATUS_SKIPPED, 0)
    not_mto = counts.get(STATUS_NOT_MTO, 0)
    empty = counts.get(STATUS_EMPTY, 0)
    rows_all = sum(row.row_count for row in rows)
    rows_written = sum(
        row.row_count for row in rows if row.status == STATUS_WRITTEN
    )
    rows_skip = sum(row.row_count for row in rows if row.status == STATUS_SKIPPED)
    parts = [
        f"Итого: спек {len(rows)}",
        f"записано {written}",
        f"ожидает {queued}" if queued else "",
        f"пропуск {skipped}",
        f"не MTO {not_mto}",
        f"без имени {empty}" if empty else "",
        f"строк {rows_all} (MTO записано {rows_written}, пропуск {rows_skip})",
    ]
    return " · ".join(part for part in parts if part)


def spec_kit_key(row: SpecProgressRow) -> tuple[str, str] | None:
    """Return the Комплекты identity for a PI spec, or ``None`` if incomplete."""

    title = (row.title or "").strip()
    mark = (row.mark or "").strip()
    if not title or not mark:
        return None
    return kit_identity_key(title, mark)


def spec_is_in_kits(
    row: SpecProgressRow,
    kit_keys: set[tuple[str, str]] | frozenset[tuple[str, str]],
) -> bool:
    """Return whether this spec's title+mark is in the Комплекты universe."""

    key = spec_kit_key(row)
    return key is not None and key in kit_keys


def spec_paint_key(
    row: SpecProgressRow,
    kit_keys: set[tuple[str, str]] | frozenset[tuple[str, str]] | None,
    status: str | None = None,
) -> str:
    """Return the ``STATUS_COLORS`` key for one spec progress row.

    Live ``writing`` keeps amber even outside Комплекты so the build
    cursor stays visible. Other rows whose title+mark is not a visible
    Комплекты kit use :data:`PAINT_OUTSIDE_KITS`.
    """

    current = status if status is not None else row.status
    if current == STATUS_WRITING:
        return STATUS_WRITING
    if kit_keys is not None and not spec_is_in_kits(row, kit_keys):
        return PAINT_OUTSIDE_KITS
    return current


def format_kits_coverage_summary(
    rows: Sequence[SpecProgressRow],
    kit_keys: Sequence[tuple[str, str]] | set[tuple[str, str]] | frozenset[tuple[str, str]],
) -> str:
    """Return PI MTO coverage against the visible Комплекты list.

    Args:
        rows: Unique PI specifications.
        kit_keys: Non-banned Комплекты identities
            (``kit_identity_key`` pairs).

    Returns:
        One Russian line: kits total, kits with convertible PI, kits
        with a written АвтоМто file, kits missing from PI, and PI specs
        whose title+mark is not in Комплекты.
    """

    kit_set = {
        kit_identity_key(title, mark) for title, mark in kit_keys
    }
    received: set[tuple[str, str]] = set()
    written: set[tuple[str, str]] = set()
    outside = 0
    for row in rows:
        key = spec_kit_key(row)
        in_kits = key is not None and key in kit_set
        if not in_kits:
            outside += 1
        if key is None or row.status in {STATUS_NOT_MTO, STATUS_EMPTY}:
            continue
        if in_kits:
            received.add(key)
            if row.status == STATUS_WRITTEN:
                written.add(key)
    total = len(kit_set)
    missing = total - len(received)
    return (
        f"Комплекты: {total} · получено МТО: {len(received)} · "
        f"записано: {len(written)} · нет в ПИ: {missing} · "
        f"спеки вне комплектов: {outside}"
    )


def plan_auto_mto_files(store: CustomerPiStore) -> tuple[tuple[AutoMtoFile, ...], tuple[str, ...]]:
    """Return planned MTO files (no disk writes) and skip notes.

    BOM/BOE/DS stems are rewritten to ``.MTO-``. The output name is always
    ``{mto_stem}_{rd_revision}_RU.xlsx`` — two PI revisions never share a
    file. If two source docs would produce the same path (same stem+rev),
    native MTO wins and the other spec is aliased (listed in
    ``source_specs``, rows not merged).
    """

    grouped, skipped = _group_mto_records(store)
    slots: dict[str, tuple[tuple[int, int, str], AutoMtoFile, list[int], list[str]]] = {}
    for (orig_spec, rev), indexes in grouped.items():
        records = [store.record(i) for i in indexes]
        title, mark = _kit_identity(records[0])
        output = spec_as_mto_stem(orig_spec)
        try:
            rel = auto_mto_relpath(title, mark, output, rev)
        except ValueError as exc:
            skipped.append(f"{orig_spec} {rev}: {exc}")
            continue
        planned = AutoMtoFile(
            title=title,
            mark=mark,
            spec=output,
            rd_revision=rev,
            relpath=rel.as_posix(),
            row_count=len(records),
            fingerprint="",
            source_specs=(orig_spec,),
        )
        key = rel.as_posix().casefold()
        priority = _output_priority(orig_spec, records[0].discipline)
        existing = slots.get(key)
        if existing is None:
            slots[key] = (priority, planned, indexes, [orig_spec])
            continue
        old_priority, old_file, old_indexes, sources = existing
        if priority < old_priority:
            sources = [orig_spec] + [spec for spec in sources if spec != orig_spec]
            slots[key] = (priority, planned, indexes, sources)
        elif orig_spec not in sources:
            sources.append(orig_spec)
            slots[key] = (old_priority, old_file, old_indexes, sources)

    by_kit: dict[tuple[str, str], list[AutoMtoFile]] = defaultdict(list)
    for _priority, planned, indexes, sources in slots.values():
        by_kit[planned.kit_key].append(
            replace(
                planned,
                source_specs=tuple(sources),
                row_count=len(indexes),
            )
        )

    result: list[AutoMtoFile] = []
    for files in by_kit.values():
        files.sort(key=_kit_file_sort_key)
        extras = tuple(
            f"{item.spec} {item.rd_revision}" for item in files[1:]
        )
        primary = files[0]
        result.append(replace(primary, extras=extras))
        result.extend(files[1:])
    result.sort(key=lambda item: (item.title, item.mark, item.spec, item.rd_revision))
    return tuple(result), tuple(skipped)


def _load_manifest(dest: Path) -> dict[str, Any]:
    path = dest / MANIFEST_NAME
    if not path.is_file():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


def _write_manifest(dest: Path, payload: dict[str, Any]) -> Path:
    path = dest / MANIFEST_NAME
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return path


def rebuild_auto_mto_catalog(
    store: CustomerPiStore,
    dest: str | Path | None = None,
    *,
    progress: ProgressCallback | None = None,
    is_cancelled: CancelCallback | None = None,
) -> AutoMtoRebuildResult:
    """Rewrite АвтоМто xlsx files from ``store``.

    Previous files listed in ``_manifest.json`` and missing from this
    plan are deleted. Unrelated files in the folder are left untouched.

    Args:
        store: Loaded PI dump.
        dest: Catalog root; default ``rd_catalog/База заказчика/АвтоМто``.
        progress: Optional ``(index, total, spec)`` callback.
        is_cancelled: Cooperative cancel flag.

    Returns:
        Written files, skip notes, and removed relative paths.
    """

    root = Path(dest) if dest else default_auto_mto_dir()
    root.mkdir(parents=True, exist_ok=True)
    planned, skipped = plan_auto_mto_files(store)
    grouped, _ = _group_mto_records(store)
    previous = _load_manifest(root)
    old_relpaths = {
        str(item.get("relpath") or "")
        for item in previous.get("files", [])
        if isinstance(item, dict)
    }
    written: list[AutoMtoFile] = []
    total = len(planned)
    for index, item in enumerate(planned, start=1):
        if is_cancelled and is_cancelled():
            raise RuntimeError("АвтоМТО: сборка отменена")
        if progress is not None:
            progress(index, total, item.spec)
        winner = item.source_specs[0] if item.source_specs else item.spec
        records = [store.record(i) for i in grouped[(winner, item.rd_revision)]]
        path = root / Path(item.relpath)
        fingerprint = write_auto_mto_workbook(path, records)
        written.append(
            replace(item, fingerprint=fingerprint, row_count=len(records))
        )
    new_relpaths = {item.relpath for item in written}
    removed: list[str] = []
    for relpath in sorted(old_relpaths - new_relpaths):
        if not relpath or relpath == MANIFEST_NAME:
            continue
        stale = root / Path(relpath)
        if stale.is_file():
            stale.unlink()
            removed.append(relpath)
    payload = {
        "format": FORMAT_ID,
        "schema_version": SCHEMA_VERSION,
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "pickle_path": store.meta.get("pickle_path"),
        "source_path": store.meta.get("source_path"),
        "files": [
            {
                "title": item.title,
                "mark": item.mark,
                "spec": item.spec,
                "rd_revision": item.rd_revision,
                "relpath": item.relpath,
                "row_count": item.row_count,
                "fingerprint": item.fingerprint,
                "extras": list(item.extras),
                "source_specs": list(item.source_specs),
            }
            for item in written
        ],
        "skipped": list(skipped),
    }
    manifest_path = _write_manifest(root, payload)
    return AutoMtoRebuildResult(
        dest=root,
        written=tuple(written),
        skipped=skipped,
        removed=tuple(removed),
        manifest_path=manifest_path,
    )


def list_auto_mto_files_by_kit(
    store: CustomerPiStore | None = None,
    dest: str | Path | None = None,
) -> dict[tuple[str, str], tuple[AutoMtoFile, ...]]:
    """Map kit identity to all planned AutoMTO files, primary first.

    Files are sorted by ``.MTO-0001`` then highest filename revision.
    ``path`` is not stored; callers join ``dest / relpath``.
    """

    loaded = store if store is not None else load_customer_pi()
    planned, _skipped = plan_auto_mto_files(loaded)
    by_kit: dict[tuple[str, str], list[AutoMtoFile]] = defaultdict(list)
    for item in planned:
        by_kit[item.kit_key].append(item)
    _ = dest
    return {
        key: tuple(sorted(files, key=_kit_file_sort_key))
        for key, files in by_kit.items()
    }


def index_auto_mto_by_kit(
    store: CustomerPiStore | None = None,
    dest: str | Path | None = None,
) -> dict[tuple[str, str], AutoMtoFile]:
    """Map kit identity to the primary AutoMTO file (highest ``MTO-0001`` rev)."""

    return {
        key: files[0]
        for key, files in list_auto_mto_files_by_kit(store, dest=dest).items()
        if files
    }


def auto_mto_path(item: AutoMtoFile, dest: str | Path | None = None) -> Path:
    """Return the absolute path of a planned file under the catalog root."""

    root = Path(dest) if dest else default_auto_mto_dir()
    return root / Path(item.relpath)


def _mto_stem_from_path(path: str | Path) -> str:
    name = Path(path).name
    parts = AgccFilenamePatterns.parse_strict(name)
    if parts is None:
        parts = AgccFilenamePatterns.parse_loose(name)
    return parts.core_stem if parts is not None else ""


def _primary_source_origin(item: AutoMtoFile) -> str:
    return (item.source_specs[0] if item.source_specs else item.spec).casefold()


def _short_error(exc: Exception) -> str:
    text = str(exc).splitlines()[0].strip()
    return (text or type(exc).__name__)[:200]


def _exact_row_key(
    row: CanonicalMtoRow,
) -> tuple[str, str, tuple[str, ...], str]:
    return (row.code, row.values, row.tags, row.units)


def _soft_row_key(row: CanonicalMtoRow) -> tuple[str, str]:
    return (row.code, row.values)


def _match_grade(
    rows: Sequence[CanonicalMtoRow],
    rd_exact: Counter[tuple[str, str, tuple[str, ...], str]],
    rd_soft: Counter[tuple[str, str]],
) -> Literal["exact", "soft"] | None:
    """Return the strongest grade of ``rows`` against RD counters."""

    if Counter(_exact_row_key(row) for row in rows) == rd_exact:
        return "exact"
    if Counter(_soft_row_key(row) for row in rows) == rd_soft:
        return "soft"
    return None


@dataclass(frozen=True, slots=True)
class MtoPairCompareResult:
    """One-to-one workbook compare using the AutoMTO content grades."""

    kind: Literal["matched", "soft", "no_match", "not_compared"]
    grade: Literal["exact", "soft", "none"] = "none"
    left_rows: int = 0
    right_rows: int = 0
    error: str = ""

    @property
    def paren_label(self) -> str:
        """Return the short Russian label used in parentheses."""

        if self.kind == "matched":
            return AUTO_MTO_EXACT_STATUS
        if self.kind == "soft":
            return AUTO_MTO_SOFT_STATUS
        if self.kind == "no_match":
            return "не совпало"
        if self.error:
            return "ошибка"
        return "не сверялось"

    def tooltip_lines(self, *, other_label: str) -> tuple[str, ...]:
        """Return Russian notes for a pairwise content compare.

        Args:
            other_label: Human name of the counterpart (ПИ / MTO РД).

        Returns:
            One or two tooltip lines.
        """

        if self.kind == "matched":
            return (
                f"Содержимое: чёткое совпадение с {other_label} "
                f"(код, количество, теги, ед. изм.), строк {self.left_rows}.",
            )
        if self.kind == "soft":
            return (
                f"Содержимое: {AUTO_MTO_SOFT_STATUS} с {other_label}, "
                f"строк АН {self.left_rows} / {other_label} {self.right_rows}.",
                "Совпали код и количество. Теги или ед. изм. разошлись.",
            )
        if self.kind == "no_match":
            extra = f" {self.error}" if self.error else ""
            return (
                f"Содержимое: не совпало по коду и количеству с {other_label}."
                f"{extra}",
                f"Строк АН {self.left_rows} / {other_label} {self.right_rows}.",
            )
        detail = f" — {self.error}" if self.error else "."
        return (f"Содержимое: сверка с {other_label} не завершена{detail}",)


def compare_document_to_path(
    left: CanonicalMtoDocument,
    right_path: str | Path,
    *,
    loader: RowLoader | None = None,
) -> MtoPairCompareResult:
    """Grade ``left`` against the workbook at ``right_path``.

    Args:
        left: Already loaded canonical document (typically the AN file).
        right_path: Counterpart workbook (AutoMTO or RD MTO).
        loader: Optional side-effect-free row loader, primarily for tests.

    Returns:
        Exact / soft / no-match / not-compared result.
    """

    try:
        right = load_canonical_mto(right_path, loader=loader)
    except Exception as exc:
        return MtoPairCompareResult(
            kind="not_compared",
            left_rows=len(left.rows),
            error=_short_error(exc),
        )
    left_exact = Counter(_exact_row_key(row) for row in left.rows)
    left_soft = Counter(_soft_row_key(row) for row in left.rows)
    grade = _match_grade(right.rows, left_exact, left_soft)
    if grade == "exact":
        return MtoPairCompareResult(
            kind="matched",
            grade="exact",
            left_rows=len(left.rows),
            right_rows=len(right.rows),
        )
    if grade == "soft":
        return MtoPairCompareResult(
            kind="soft",
            grade="soft",
            left_rows=len(left.rows),
            right_rows=len(right.rows),
        )
    return MtoPairCompareResult(
        kind="no_match",
        left_rows=len(left.rows),
        right_rows=len(right.rows),
    )


def compare_mto_pair(
    left_path: str | Path,
    right_path: str | Path,
    *,
    loader: RowLoader | None = None,
) -> MtoPairCompareResult:
    """Compare two MTO workbooks with the AutoMTO exact/soft grades.

    Rows are not aggregated by code. ``NAME`` / ``VENDOR`` / ``TYPE_MARK``
    are ignored.

    Args:
        left_path: First workbook (typically AN).
        right_path: Second workbook (AutoMTO or RD).
        loader: Optional side-effect-free row loader, primarily for tests.

    Returns:
        Pairwise compare result.
    """

    try:
        left = load_canonical_mto(left_path, loader=loader)
    except Exception as exc:
        return MtoPairCompareResult(
            kind="not_compared",
            error=_short_error(exc),
        )
    return compare_document_to_path(left, right_path, loader=loader)


def compare_auto_mto_to_rd(
    files: Sequence[AutoMtoFile],
    rd_path: str | Path,
    *,
    dest: str | Path | None = None,
    loader: RowLoader | None = None,
    max_members: int = 4,
    max_combinations: int = 128,
) -> AutoMtoCompareResult:
    """Find a graded single-file or bounded composite match for an RD MTO.

    Grades, strongest first:

    * **exact** — same multiset of ``(CODE, VALUES, TAGS, UNITS)``.
    * **soft** — same multiset of ``(CODE, VALUES)``; tags or units differ.
    * **no_match** — code or quantity differs.

    ``NAME``, ``VENDOR``, and ``TYPE_MARK`` are ignored. Rows are not
    aggregated by code.

    Args:
        files: AutoMTO files planned for a kit.
        rd_path: Current RD MTO workbook.
        dest: AutoMTO catalog root.
        loader: Optional side-effect-free row loader, primarily for tests.
        max_members: Maximum files in one composite candidate.
        max_combinations: Maximum equal-length composite groups actually
            checked for an exact grade.

    Returns:
        Graded comparison result. Exact wins over soft; a single file wins
        over a composite of the same grade.
    """

    try:
        rd_document = load_canonical_mto(rd_path, loader=loader)
    except Exception as exc:
        return AutoMtoCompareResult(
            match_kind="not_compared",
            members=(),
            rd_rows=0,
            auto_rows=0,
            combinations_checked=0,
            error=f"RD MTO read failed: {_short_error(exc)}",
        )

    rd_stem = _mto_stem_from_path(rd_path)
    if not rd_stem:
        return AutoMtoCompareResult(
            match_kind="not_compared",
            members=(),
            rd_rows=len(rd_document.rows),
            auto_rows=0,
            combinations_checked=0,
            error="RD MTO filename stem is not recognized",
        )

    candidates = sorted(
        (
            item
            for item in files
            if item.spec.casefold() == rd_stem.casefold()
            and auto_mto_path(item, dest).is_file()
        ),
        key=_kit_file_sort_key,
    )
    if not candidates:
        return AutoMtoCompareResult(
            match_kind="not_compared",
            members=(),
            rd_rows=len(rd_document.rows),
            auto_rows=0,
            combinations_checked=0,
            error=f"no existing AutoMTO files for {rd_stem}",
        )

    loaded: list[tuple[AutoMtoFile, CanonicalMtoDocument]] = []
    read_errors: list[str] = []
    rd_exact = Counter(_exact_row_key(row) for row in rd_document.rows)
    rd_soft = Counter(_soft_row_key(row) for row in rd_document.rows)
    soft_single: tuple[AutoMtoFile, CanonicalMtoDocument] | None = None
    for item in candidates:
        path = auto_mto_path(item, dest)
        try:
            document = load_canonical_mto(path, loader=loader)
        except Exception as exc:
            read_errors.append(f"{path.name}: {_short_error(exc)}")
            continue
        loaded.append((item, document))
        grade = _match_grade(document.rows, rd_exact, rd_soft)
        if grade == "exact":
            return AutoMtoCompareResult(
                match_kind="single",
                members=(item,),
                rd_rows=len(rd_document.rows),
                auto_rows=len(document.rows),
                combinations_checked=0,
                grade="exact",
            )
        if grade == "soft" and soft_single is None:
            soft_single = (item, document)

    if not loaded:
        detail = read_errors[0] if read_errors else "unknown read error"
        return AutoMtoCompareResult(
            match_kind="not_compared",
            members=(),
            rd_rows=len(rd_document.rows),
            auto_rows=0,
            combinations_checked=0,
            error=f"no readable AutoMTO files: {detail}",
        )

    checked = 0
    max_size = min(max(1, max_members), len(loaded))
    limit = max(0, max_combinations)
    soft_composite_members: tuple[AutoMtoFile, ...] | None = None
    soft_composite_rows = 0
    for size in range(2, max_size + 1):
        for group in combinations(loaded, size):
            origins = {_primary_source_origin(item) for item, _document in group}
            if len(origins) != size:
                continue
            rows = tuple(
                row
                for _item, document in group
                for row in document.rows
            )
            members = tuple(item for item, _document in group)
            if len(rows) == len(rd_document.rows):
                if checked >= limit:
                    break
                checked += 1
                grade = _match_grade(rows, rd_exact, rd_soft)
                if grade == "exact":
                    return AutoMtoCompareResult(
                        match_kind="composite",
                        members=members,
                        rd_rows=len(rd_document.rows),
                        auto_rows=len(rows),
                        combinations_checked=checked,
                        grade="exact",
                    )
                if grade == "soft" and soft_composite_members is None:
                    soft_composite_members = members
                    soft_composite_rows = len(rows)
            elif (
                soft_composite_members is None
                and Counter(_soft_row_key(row) for row in rows) == rd_soft
            ):
                soft_composite_members = members
                soft_composite_rows = len(rows)
        if checked >= limit:
            break

    if soft_single is not None:
        item, document = soft_single
        return AutoMtoCompareResult(
            match_kind="single",
            members=(item,),
            rd_rows=len(rd_document.rows),
            auto_rows=len(document.rows),
            combinations_checked=checked,
            grade="soft",
        )
    if soft_composite_members is not None:
        return AutoMtoCompareResult(
            match_kind="composite",
            members=soft_composite_members,
            rd_rows=len(rd_document.rows),
            auto_rows=soft_composite_rows,
            combinations_checked=checked,
            grade="soft",
        )

    return AutoMtoCompareResult(
        match_kind="no_match",
        members=(),
        rd_rows=len(rd_document.rows),
        auto_rows=0,
        combinations_checked=checked,
        error="; ".join(read_errors),
    )


def revision_text_from_mto_filename(name: str) -> str:
    """Return ``04-AN01`` from an AGCC MTO file name, or empty if unparsed."""

    try:
        rev, appendix = parse_agcc_mto_xlsx_revision_for_chain(Path(name).name)
    except ValueError:
        return ""
    an = appendix[2:] if appendix.upper().startswith("AN") else appendix
    return format_revision(rev, an or None)


def revision_texts_match(left: str, right: str) -> bool | None:
    """Compare two filename-revision strings; ``None`` if either side is empty."""

    left_rev, left_app = parse_sheet_revision(left)
    right_rev, right_app = parse_sheet_revision(right)
    return revision_matches_rd(left_rev, left_app, right_rev, right_app)


def _cli(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Build АвтоМто xlsx catalog from the customer PI pickle."
    )
    parser.add_argument(
        "--pickle",
        type=Path,
        default=None,
        help="PI pickle (default: rd_catalog/База заказчика/customer_pi_bcc.pkl)",
    )
    parser.add_argument(
        "--dest",
        type=Path,
        default=None,
        help="Catalog folder (default: rd_catalog/База заказчика/АвтоМто)",
    )
    args = parser.parse_args(argv)
    store = load_customer_pi(args.pickle)
    result = rebuild_auto_mto_catalog(store, args.dest)
    print(
        f"wrote {len(result.written)} files → {result.dest} "
        f"(skipped {len(result.skipped)}, removed {len(result.removed)})"
    )
    return 0


if __name__ == "__main__":
    sys.exit(_cli())
