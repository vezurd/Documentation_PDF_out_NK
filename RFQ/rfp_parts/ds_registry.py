"""Canonical DS registry: several sheet rows may share one actual DS number.

One sheet row is one link (a UL folder, an RFP file, or both). Rows of the
same actual number collapse into one specification bag. The same UL folder
may be named by several actual numbers. One RFP number or RFP file cannot
sit on two actual numbers. The working workbook lives
in ``_RFP`` as ``Реестр_ДС_УЛ.xlsx``. «Проверить реестр» renames a legacy
file to ``Реестр_ДС_УЛ_old.xlsx`` and replaces it. If Excel holds the file
open, a dated copy is written; the program reads the newest of the main file
and at most five dated copies. An old wide sheet (blocks «Связь N») is still
read and collapsed into rows of the same number.
"""

from __future__ import annotations

import os
import re
import shutil
import time
from dataclasses import dataclass, field, replace
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Callable, Iterator, Literal, Sequence
from zipfile import BadZipFile, ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from RFQ.rfp_parts.ds_identity import (
    DsIdentity,
    parse_rfp_ds_identity,
    parse_ul_folder_ds_identity,
)

IssueLevel = Literal["ERROR", "WARN", "OVERLAY"]
RegistryFormat = Literal["new", "legacy", "unknown"]

# ---------------------------------------------------------------------------
# Public constants
# ---------------------------------------------------------------------------

FORMAT_VERSION = 1
REGISTRY_SHEET_NAME = "Реестр ДС"
LEGEND_SHEET_NAME = "Как заполнять"
REGISTRY_TABLE_NAME = "ReestrDS"
MIGRATION_REPORT_PREFIX = "Отчет по миграции реестра ДС"
BACKUP_SUFFIX_FORMAT = "%Y%m%d_%H%M%S"

DEFAULT_RFP_BASE = Path(
    r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP"
)
CANONICAL_REGISTRY_NAME = "Реестр_ДС_УЛ.xlsx"
OLD_REGISTRY_NAME = "Реестр_ДС_УЛ_old.xlsx"
DATED_REGISTRY_RE = re.compile(
    r"^Реестр_ДС_УЛ_(\d{8}_\d{6})\.xlsx$", re.IGNORECASE
)
MAX_DATED_REGISTRY_COPIES = 5
DEFAULT_REGISTRY_PATH = DEFAULT_RFP_BASE / CANONICAL_REGISTRY_NAME

STATUS_ACTIVE = "Активен"
STATUS_HISTORY = "История"
STATUS_DISABLED = "Отключен"
ALLOWED_STATUSES: frozenset[str] = frozenset(
    {STATUS_ACTIVE, STATUS_HISTORY, STATUS_DISABLED}
)

MODE_WHOLE = "Вся ДС"
MODE_FILTER = "По фильтру"
MODE_NO_UL = "Нет УЛ"
MODE_NEEDS_SPLIT = "Требует распределения"
ALLOWED_MODES: frozenset[str] = frozenset(
    {MODE_WHOLE, MODE_FILTER, MODE_NO_UL, MODE_NEEDS_SPLIT}
)

NO_UL_GROUP_PREFIX = "NO_UL:"
UL_SPLIT = ";"
HEADER_REQUIRED_MARK = "*"
HEADER_CONDITIONAL_MARK = "\u2020"  # †
DEFAULT_SPARE_ROWS = 15
MIN_RELATION_BLOCKS = 1

LEGACY_HEADER_CURRENT = "Актуальный ДС"
LEGACY_HEADER_OLD = "Старый ДС"
LEGACY_HEADER_UL = "УЛ"
LEGACY_HEADER_NOTE = "Примечание"

HDR_STATUS = f"Статус {HEADER_REQUIRED_MARK}"
HDR_SOURCE_ID = f"Актуальный номер ДС {HEADER_REQUIRED_MARK}"
HDR_SOURCE_ID_LEGACY = "ID ДС источника"
HDR_FILE_DS = "Файл ДС"
HDR_PREVIOUS = "Исторический номер ДС"
HDR_REVISION = "Ревизия"
HDR_NOTE = "Примечание"
HDR_ORIGINAL_ROW = "Исходная строка реестра"
HDR_UL_FOLDER = f"Папка УЛ {HEADER_CONDITIONAL_MARK}"
HDR_RFP_NUMBER = f"Номер RFP {HEADER_CONDITIONAL_MARK}"
HDR_RFP_FILE = "Файл RFP"
HDR_RFP_RECONCILE = "Сверка RFP"
SOURCE_INFO_SHEET_NAME = "Исходный ДС"
RFP_STATUS_MATCH = "совпало"
RFP_STATUS_MISS = "не совпало"
FLAT_HEADER_TITLES: tuple[str, ...] = (
    HDR_STATUS,
    HDR_SOURCE_ID,
    HDR_PREVIOUS,
    HDR_FILE_DS,
    HDR_UL_FOLDER,
    HDR_RFP_NUMBER,
    HDR_RFP_FILE,
    HDR_RFP_RECONCILE,
)
_FLAT_WIDTHS = (16, 13.14, 12, 93.57, 36, 16, 95.29, 18)
_LINE_HEIGHT = 15.0
_HEADER_MIN_HEIGHT = 45.0
_DATA_MIN_HEIGHT = 15.0

CORE_HEADER_TITLES: tuple[str, ...] = (
    HDR_STATUS,
    HDR_SOURCE_ID,
    HDR_PREVIOUS,
    HDR_FILE_DS,
)
_SOURCE_ID_HEADER_NAMES = ("Актуальный номер ДС", "Номер ДС", "ID ДС источника")
_HISTORICAL_HEADER_NAMES = ("Исторический номер ДС", "Предыдущий / старый ДС")

ISSUE_MISSING_SOURCE_ID = "missing_source_id"
ISSUE_DUPLICATE_SOURCE_ID = "duplicate_source_id"
ISSUE_MISSING_RELATION = "missing_relation"
ISSUE_PARTIAL_BLOCK = "partial_block"
ISSUE_GAPPED_BLOCK = "gapped_block"
ISSUE_GROUP_MISMATCH = "group_mismatch"
ISSUE_RFP_KEY_CONFLICT = "rfp_key_conflict"
ISSUE_UL_REQUIRED = "ul_required"
ISSUE_FILTERS_REQUIRED = "filters_required"
ISSUE_MULTI_LINK_NEEDS_SPLIT = "multi_link_needs_split"
ISSUE_INVALID_STATUS = "invalid_status"
ISSUE_INVALID_MODE = "invalid_mode"
ISSUE_UL_UNPARSED = "ul_unparsed"
ISSUE_UL_IDENTITY_MISMATCH = "ul_identity_mismatch"
ISSUE_UL_FOLDER_MISSING = "ul_folder_missing"
ISSUE_UL_UNEXPECTED = "ul_unexpected"
ISSUE_SHARED_UL = "shared_ul_folder"
ISSUE_SHARED_RFP = "shared_rfp"
ISSUE_DUPLICATE_LINK = "duplicate_link"
ISSUE_IDENTITY_SPLIT = "identity_split"
ISSUE_ORPHAN = "orphan_unassigned"
ISSUE_RFP_FILE_MISSING = "rfp_file_missing"

RFP_RECONCILE_MATCH = "совпало"
RFP_RECONCILE_DIFF = "расхождение"
RFP_RECONCILE_NO_FILE = "нет файла"

_FILL_REQUIRED = PatternFill(
    fill_type="solid", fgColor="1B4F72", start_color="1B4F72", end_color="1B4F72"
)
_FONT_REQUIRED = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
_FILL_CONDITIONAL = PatternFill(
    fill_type="solid", fgColor="FFC000", start_color="FFC000", end_color="FFC000"
)
_FONT_CONDITIONAL = Font(bold=True, color="000000", name="Calibri", size=11)
_FILL_AUX = PatternFill(
    fill_type="solid", fgColor="D9D9D9", start_color="D9D9D9", end_color="D9D9D9"
)
_FONT_AUX = Font(bold=True, color="000000", name="Calibri", size=11)
_FILL_BLANK = PatternFill(
    fill_type="solid", fgColor="FFFF00", start_color="FFFF00", end_color="FFFF00"
)
_FILL_UL_ABSENT = PatternFill(
    fill_type="solid", fgColor="BDD7EE", start_color="BDD7EE", end_color="BDD7EE"
)
_FILL_TODO = _FILL_BLANK
_FONT_DATA = Font(name="Calibri", size=11)
_ALIGN_HEADER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_ALIGN_DATA = Alignment(horizontal="left", vertical="center", wrap_text=True)
_THIN = Side(style="thin", color="B0B0B0")
_BLOCK_BORDER_COLORS = ("1B4F72", "548235", "C65911", "7030A0")
_COMMENT_AUTHOR = "реестр ДС"

_CORE_WIDTHS = (16, 22, 24, 42)
_BLOCK_WIDTHS = (22, 16, 42, 16, 32, 26, 18, 18)

_RELATION_FIELD_STEMS = (
    "ID группы поставки",
    "Фактический ДС / ключ RFP",
    "Папка УЛ",
    "Режим распределения",
    "Фильтр титула",
    "Фильтр марки",
)

_HEADER_RE_GROUP = re.compile(r"^ID группы поставки\s+(\d+)$", re.IGNORECASE)
_HEADER_RE_RFP = re.compile(
    r"^(?:Фактический ДС / ключ RFP|Номер RFP)\s+(\d+)$", re.IGNORECASE
)
_HEADER_RE_RFP_FILE = re.compile(r"^Файл RFP\s+(\d+)$", re.IGNORECASE)
_HEADER_RE_RFP_STATUS = re.compile(r"^Статус RFP\s+(\d+)$", re.IGNORECASE)
_DS_NUMBER_PREFIX_RE = re.compile(r"^ДС\s*", re.IGNORECASE)
_HEADER_RE_UL = re.compile(r"^Папка УЛ\s+(\d+)$", re.IGNORECASE)
_HEADER_RE_MODE = re.compile(r"^Режим распределения\s+(\d+)$", re.IGNORECASE)
_HEADER_RE_TITLE = re.compile(r"^Фильтр титула\s+(\d+)$", re.IGNORECASE)
_HEADER_RE_MARK = re.compile(r"^Фильтр марки\s+(\d+)$", re.IGNORECASE)
_WS_RE = re.compile(r"\s+")


# ---------------------------------------------------------------------------
# Exceptions
# ---------------------------------------------------------------------------


_FILTER_COLUMN_RE = re.compile(
    r"<filterColumn\b[^>]*/>|<filterColumn\b[^>]*>.*?</filterColumn>",
    re.DOTALL,
)


def _xlsx_without_filter_columns(data: bytes) -> bytes:
    """Drop Excel filter criteria so openpyxl can open the book.

    A table dropdown (``autoFilter ref``) stays. Criteria such as
    ``customFilter val=" "`` on «Исходный ДС» make openpyxl raise
    "could not read worksheets" and the check never starts. The criteria
    are not part of the registry contract.
    """

    try:
        source = ZipFile(BytesIO(data))
    except BadZipFile:
        return data
    changed = False
    pieces: list[tuple[object, bytes]] = []
    for info in source.infolist():
        blob = source.read(info.filename)
        if (
            info.filename.startswith("xl/worksheets/")
            and info.filename.endswith(".xml")
            and b"filterColumn" in blob
        ):
            text = blob.decode("utf-8")
            cleaned = _FILTER_COLUMN_RE.sub("", text)
            if cleaned != text:
                blob = cleaned.encode("utf-8")
                changed = True
        pieces.append((info, blob))
    if not changed:
        return data
    out = BytesIO()
    with ZipFile(out, "w") as target:
        for info, blob in pieces:
            target.writestr(info, blob)
    return out.getvalue()


def _open_workbook(path: Path, *, data_only: bool) -> Workbook:
    """Load xlsx from bytes so Windows does not keep the destination locked."""

    data = _xlsx_without_filter_columns(Path(path).read_bytes())
    return load_workbook(BytesIO(data), data_only=data_only)


class DsRegistryError(Exception):
    """Base error for registry load / migration / replace."""


class DsRegistryFormatError(DsRegistryError):
    """Workbook is not the canonical new registry layout."""


class DsRegistryReplaceError(DsRegistryError):
    """Canonical replace refused or the filesystem denied the swap."""


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class DsRegistryRelation:
    """One numbered supply-group link on a source-DS row."""

    group_id: str
    rfp_key: str
    ul_folder: str
    mode: str = ""
    title_filter: str = ""
    mark_filter: str = ""
    block_index: int = 1
    rfp_file: str = ""
    excel_row: int = 0

    def is_empty(self) -> bool:
        return not any(
            (
                self.group_id,
                self.rfp_key,
                self.ul_folder,
                self.mode,
                self.title_filter,
                self.mark_filter,
                self.rfp_file,
            )
        )

    def has_complete_filters(self) -> bool:
        return bool(self.title_filter.strip() and self.mark_filter.strip())


@dataclass(frozen=True, slots=True)
class DsRegistryRow:
    """One source-DS row of the canonical registry."""

    status: str
    source_id: str
    previous_ds: str = ""
    ds_file: str = ""
    revision: str = ""
    note: str = ""
    relations: tuple[DsRegistryRelation, ...] = ()
    excel_row: int = 0
    original_excel_row: int | None = None

    @property
    def is_active(self) -> bool:
        return self.status == STATUS_ACTIVE

    @property
    def is_history(self) -> bool:
        return self.status == STATUS_HISTORY


@dataclass(frozen=True, slots=True)
class DsRegistryIssue:
    """One validation finding."""

    code: str
    level: IssueLevel
    message: str
    excel_row: int | None = None
    source_id: str = ""
    group_id: str = ""
    block_index: int | None = None
    blocks_overlay: bool = False
    field: str = ""


@dataclass
class DsRegistryValidation:
    """Result of validating parsed registry rows."""

    issues: list[DsRegistryIssue] = field(default_factory=list)

    @property
    def error_count(self) -> int:
        return sum(1 for item in self.issues if item.level == "ERROR")

    @property
    def is_ok(self) -> bool:
        return self.error_count == 0

    @property
    def blocks_overlay(self) -> bool:
        return any(
            item.blocks_overlay or item.level == "OVERLAY" for item in self.issues
        )

    @property
    def overlay_blocked_group_ids(self) -> tuple[str, ...]:
        seen: list[str] = []
        for item in self.issues:
            if (item.blocks_overlay or item.level == "OVERLAY") and item.group_id:
                if item.group_id not in seen:
                    seen.append(item.group_id)
        return tuple(seen)

    def summary_line(self) -> str:
        if self.is_ok and not self.blocks_overlay:
            return f"OK: реестр ДС валиден ({len(self.issues)} замечаний нет)"
        parts = [f"ERROR={self.error_count}"]
        overlay = sum(1 for item in self.issues if item.blocks_overlay)
        warn = sum(1 for item in self.issues if item.level == "WARN")
        if overlay:
            parts.append(f"OVERLAY={overlay}")
        if warn:
            parts.append(f"WARN={warn}")
        return "Реестр ДС: " + ", ".join(parts)


@dataclass
class DsRegistryDocument:
    """Parsed canonical registry workbook."""

    path: Path
    rows: list[DsRegistryRow]
    validation: DsRegistryValidation
    max_relation_blocks: int
    format_version: int = FORMAT_VERSION

    @property
    def active_rows(self) -> list[DsRegistryRow]:
        return [row for row in self.rows if row.is_active]


@dataclass(frozen=True, slots=True)
class LegacyRegistryRow:
    """One row of the historical 4-column registry."""

    excel_row: int
    current_ds: str
    old_ds: str
    ul_text: str
    note: str


@dataclass
class DsRegistryMigrationResult:
    """Outcome of migrating a legacy workbook to an explicit new path."""

    source_path: Path
    output_path: Path
    report_path: Path
    rows: list[DsRegistryRow]
    validation: DsRegistryValidation
    legacy_rows: list[LegacyRegistryRow]
    issues: list[DsRegistryIssue] = field(default_factory=list)

    @property
    def active_count(self) -> int:
        return sum(1 for row in self.rows if row.is_active)

    @property
    def history_count(self) -> int:
        return sum(1 for row in self.rows if row.is_history)


@dataclass(frozen=True, slots=True)
class DsRegistryReplaceResult:
    """Outcome of a confirmed atomic canonical replace."""

    target_path: Path
    backup_path: Path
    replaced: bool


@dataclass(frozen=True, slots=True)
class _RelationBlockLayout:
    index: int
    group_id: int
    rfp_key: int
    ul_folder: int
    mode: int
    title_filter: int
    mark_filter: int
    rfp_file: int = 0
    rfp_status: int = 0

    def columns(self) -> tuple[int, ...]:
        ordered = (
            self.group_id,
            self.rfp_key,
            self.rfp_file,
            self.rfp_status,
            self.ul_folder,
            self.mode,
            self.title_filter,
            self.mark_filter,
        )
        return tuple(col for col in ordered if col)


@dataclass(frozen=True, slots=True)
class _HeaderLayout:
    headers: tuple[str, ...]
    status: int
    source_id: int
    previous: int
    revision: int
    note: int
    original_row: int
    blocks: tuple[_RelationBlockLayout, ...]
    file_ds: int = 0
    flat_ul: int = 0
    flat_rfp_number: int = 0
    flat_rfp_file: int = 0
    flat_reconcile: int = 0

    @property
    def is_flat(self) -> bool:
        return self.flat_ul > 0 or self.flat_rfp_number > 0

    @property
    def max_blocks(self) -> int:
        return len(self.blocks)

    @property
    def column_count(self) -> int:
        return len(self.headers)


# ---------------------------------------------------------------------------
# Header helpers
# ---------------------------------------------------------------------------


def relation_headers(block_index: int) -> tuple[str, ...]:
    """Return the canonical headers of relation block ``N``."""

    n = int(block_index)
    return (
        f"ID группы поставки {n} {HEADER_REQUIRED_MARK}",
        f"Номер RFP {n} {HEADER_REQUIRED_MARK}",
        f"Файл RFP {n}",
        f"Статус RFP {n}",
        f"Папка УЛ {n} {HEADER_CONDITIONAL_MARK}",
        f"Режим распределения {n} {HEADER_REQUIRED_MARK}",
        f"Фильтр титула {n} {HEADER_CONDITIONAL_MARK}",
        f"Фильтр марки {n} {HEADER_CONDITIONAL_MARK}",
    )


def build_header_row(max_blocks: int) -> list[str]:
    """Return the full header row for ``max_blocks`` relation groups."""

    count = max(MIN_RELATION_BLOCKS, int(max_blocks))
    headers = list(CORE_HEADER_TITLES)
    for index in range(1, count + 1):
        headers.extend(relation_headers(index))
    return headers


def normalize_header(text: object) -> str:
    """Strip markers, newlines and collapsed whitespace from a header cell."""

    raw = str(text or "").replace("\n", " ").replace("\r", " ")
    raw = raw.replace(HEADER_REQUIRED_MARK, " ").replace(HEADER_CONDITIONAL_MARK, " ")
    return _WS_RE.sub(" ", raw).strip()


def format_registry_ds_number(source_id: str) -> str:
    """Show a registry DS id as ``ДС11`` / ``ДС4905_1``."""

    text = cell_text(source_id)
    if not text:
        return ""
    bare = _DS_NUMBER_PREFIX_RE.sub("", text).strip()
    return f"ДС{bare}" if bare else text


def parse_registry_ds_number(value: object) -> str:
    """Read ``ДС11``, ``ДС 11`` or ``11`` back to the stored source id ``11``."""

    text = cell_text(value)
    if not text:
        return ""
    return _DS_NUMBER_PREFIX_RE.sub("", text).strip()


def cell_text(value: object) -> str:
    """Canonical string of an Excel cell used as an identifier or label."""

    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        text = format(value, "g").strip()
        return text
    return str(value).strip()


def split_ul_folders(text: str) -> list[str]:
    """Split a legacy UL cell on ``;`` only and strip each token."""

    parts: list[str] = []
    for token in str(text or "").split(UL_SPLIT):
        folder = token.strip()
        if folder:
            parts.append(folder)
    return parts


def no_ul_group_id(source_id: str) -> str:
    """Return the singleton group id for a source DS without a UL folder."""

    return f"{NO_UL_GROUP_PREFIX}{source_id}"


def canonical_supply_group_id(actual: int | str) -> str:
    """Return the human-readable supply-group id for a parsed UL actual.

    Args:
        actual: Parsed actual DS number (``13``, ``4905``).

    Returns:
        ``ДС13`` / ``ДС4905``. Does not apply to ``NO_UL:<id>`` groups.
        The RFP key stays the bare actual string (``13`` / ``4905``).
    """

    return f"ДС{cell_text(actual)}"


def _header_kind(header: str) -> Literal["required", "conditional", "auxiliary"]:
    if HEADER_CONDITIONAL_MARK in header:
        return "conditional"
    if HEADER_REQUIRED_MARK in header:
        return "required"
    return "auxiliary"


def _layout_from_max_blocks(max_blocks: int) -> _HeaderLayout:
    headers = tuple(build_header_row(max_blocks))
    blocks: list[_RelationBlockLayout] = []
    offset = len(CORE_HEADER_TITLES)
    count = max(MIN_RELATION_BLOCKS, int(max_blocks))
    width = len(relation_headers(1))
    for index in range(1, count + 1):
        base = offset + (index - 1) * width
        blocks.append(
            _RelationBlockLayout(
                index=index,
                group_id=base + 1,
                rfp_key=base + 2,
                rfp_file=base + 3,
                rfp_status=base + 4,
                ul_folder=base + 5,
                mode=base + 6,
                title_filter=base + 7,
                mark_filter=base + 8,
            )
        )
    return _HeaderLayout(
        headers=headers,
        status=1,
        source_id=2,
        previous=3,
        file_ds=4,
        revision=0,
        note=0,
        original_row=0,
        blocks=tuple(blocks),
    )


def _discover_layout(headers: Sequence[object]) -> _HeaderLayout:
    indexed: dict[str, int] = {}
    raw_by_col: dict[int, str] = {}
    for col, value in enumerate(headers, start=1):
        raw = str(value or "").strip()
        if not raw:
            continue
        raw_by_col[col] = raw
        key = normalize_header(raw)
        if key and key not in indexed:
            indexed[key] = col

    source_header = next(
        (name for name in _SOURCE_ID_HEADER_NAMES if name in indexed),
        "",
    )
    missing_core = [
        title
        for title, present in (
            ("Статус", "Статус" in indexed),
            ("Актуальный номер ДС", bool(source_header)),
        )
        if not present
    ]
    if missing_core:
        raise DsRegistryFormatError(
            "нет обязательных заголовков: " + ", ".join(missing_core)
        )

    found: dict[int, dict[str, int]] = {}
    known_norm: set[str] = {
        "Статус",
        "Актуальный номер ДС",
        "Номер ДС",
        "ID ДС источника",
        "Файл ДС",
        "Исторический номер ДС",
        "Предыдущий / старый ДС",
        "Ревизия",
        "Примечание",
        "Исходная строка реестра",
        "Папка УЛ",
        "Номер RFP",
        "Файл RFP",
        "Сверка RFP",
    }
    matchers = (
        ("group_id", _HEADER_RE_GROUP),
        ("rfp_key", _HEADER_RE_RFP),
        ("rfp_file", _HEADER_RE_RFP_FILE),
        ("rfp_status", _HEADER_RE_RFP_STATUS),
        ("ul_folder", _HEADER_RE_UL),
        ("mode", _HEADER_RE_MODE),
        ("title_filter", _HEADER_RE_TITLE),
        ("mark_filter", _HEADER_RE_MARK),
    )
    for col, raw in raw_by_col.items():
        key = normalize_header(raw)
        matched = False
        for field_name, pattern in matchers:
            hit = pattern.match(key)
            if not hit:
                continue
            block_n = int(hit.group(1))
            found.setdefault(block_n, {})[field_name] = col
            known_norm.add(key)
            matched = True
            break
        if not matched and key not in known_norm:
            raise DsRegistryFormatError(f"неожиданный заголовок: {raw}")

    if found:
        numbers = sorted(found)
        expected = list(range(1, numbers[-1] + 1))
        if numbers != expected:
            raise DsRegistryFormatError(
                "нумерованные блоки связей идут с пропусками: "
                + ",".join(str(n) for n in numbers)
            )
        blocks: list[_RelationBlockLayout] = []
        for index in numbers:
            fields = found[index]
            missing = [
                stem
                for stem, name in zip(
                    _RELATION_FIELD_STEMS,
                    (
                        "group_id",
                        "rfp_key",
                        "ul_folder",
                        "mode",
                        "title_filter",
                        "mark_filter",
                    ),
                )
                if name not in fields
            ]
            if missing:
                raise DsRegistryFormatError(
                    f"неполный блок связи {index}: нет " + ", ".join(missing)
                )
            blocks.append(
                _RelationBlockLayout(
                    index=index,
                    group_id=fields["group_id"],
                    rfp_key=fields["rfp_key"],
                    ul_folder=fields["ul_folder"],
                    mode=fields["mode"],
                    title_filter=fields["title_filter"],
                    mark_filter=fields["mark_filter"],
                    rfp_file=fields.get("rfp_file", 0),
                    rfp_status=fields.get("rfp_status", 0),
                )
            )
    else:
        blocks = []

    header_values = []
    max_col = max(
        [indexed[k] for k in indexed]
        + [col for block in blocks for col in block.columns()]
        + [len(headers)]
    )
    for col in range(1, max_col + 1):
        header_values.append(raw_by_col.get(col, ""))
    return _HeaderLayout(
        headers=tuple(header_values),
        status=indexed["Статус"],
        source_id=indexed[source_header],
        file_ds=indexed.get("Файл ДС", 0),
        previous=next(
            (indexed[name] for name in _HISTORICAL_HEADER_NAMES if name in indexed),
            0,
        ),
        revision=indexed.get("Ревизия", 0),
        note=indexed.get("Примечание", 0),
        original_row=indexed.get("Исходная строка реестра", 0),
        blocks=tuple(blocks),
        flat_ul=0 if blocks else indexed.get("Папка УЛ", 0),
        flat_rfp_number=0 if blocks else indexed.get("Номер RFP", 0),
        flat_rfp_file=0 if blocks else indexed.get("Файл RFP", 0),
        flat_reconcile=0 if blocks else indexed.get("Сверка RFP", 0),
    )


def _header_comment(header: str) -> str:
    comments = {
        HDR_STATUS: (
            "Обязательно (*). Допустимо: Активен, История, Отключен. "
            "Исторические строки не участвуют в запуске. "
            "Маркер обязательности — символ * и цвет шапки."
        ),
        HDR_SOURCE_ID: (
            "Обязательно для статуса Активен (*). Это номер файла ДС: ДС11, ДС8, ДС4905_1. "
            "В Шаге 4 ему соответствует ключ посадки, не ярлык «Порядковый ДС»."
        ),
        HDR_PREVIOUS: (
            "Справка. Старый номер не выбирает ключ посадки и не входит в мешок. "
            "Одна и та же папка УЛ на номерах 14 и 48 — ошибка реестра."
        ),
        HDR_UL_FOLDER: (
            "Одна папка на строку. Несколько папок одного номера — несколько строк. "
            "Одна папка не может стоять на разных актуальных номерах. "
            "Если папки ещё нет на диске, ячейка синяя: свод не останавливается."
        ),
        HDR_RFP_NUMBER: (
            "Номер файла в корне RFP_Зиновьев. Один номер принадлежит только "
            "одному актуальному ДС. Несколько файлов одного номера суммируются."
        ),
        HDR_RFP_FILE: (
            "Имена файлов RFP, каждое с новой строки. Каждое имя сверяется отдельно. "
            "Нет одного из имён — ячейка жёлтая, и в мешок RFP не входит только оно. "
            "Когда других файлов нет, свод остаётся из ДС."
        ),
        HDR_RFP_RECONCILE: (
            "Пишет робот: совпало, расхождение или нет файла. "
            "Одинаково на всех строках номера. При сборке свода не читается."
        ),
        HDR_FILE_DS: (
            "Имя файла ДС. Несколько файлов одного номера — по одному имени в строке. "
            "Проверка реестра пересобирает ячейку из папки ДС и стирает прежний текст."
        ),
        HDR_REVISION: "Вспомогательное поле. Заполняется человеком, робот не угадывает.",
        HDR_NOTE: "Вспомогательное поле. Примечание как есть.",
        HDR_ORIGINAL_ROW: (
            "Служебное provenance: номер строки в прежнем реестре. Не править вручную."
        ),
    }
    if header in comments:
        return comments[header]
    n_match = re.search(r"(\d+)", header)
    n = n_match.group(1) if n_match else "N"
    if header.startswith("ID группы поставки"):
        return (
            f"Обязательно (*) для заполненного блока {n}. Человекочитаемый ключ "
            "группы: ДС{actual} (ДС13, ДС4905). Для комплекта без УЛ — NO_UL:<id>. "
            "Повторённые одинаковые ID должны совпадать по ключу RFP и папке УЛ."
        )
    if header.startswith("Фактический ДС") or header.startswith("Номер RFP"):
        return (
            f"Обязательно (*) для блока {n}. Номер в имени файла в корне "
            "RFP_Зиновьев: 13, 4905. Один номер — одна группа."
        )
    if header.startswith("Файл RFP"):
        return (
            f"Имя файла RFP для блока {n}. Путь в ячейке не пишется. "
            "Робот подставляет файл из корня RFP_Зиновьев."
        )
    if header.startswith("Статус RFP"):
        return (
            f"Совпало — в корне RFP_Зиновьев есть файл с этим номером. "
            f"Не совпало — такого файла нет. Это не сверка количеств: "
            f"количества смотрит «Наложить RFP на группы»."
        )
    if header.startswith("Папка УЛ"):
        return (
            f"Условно обязательно (†) для блока {n}: точное имя каталога УЛ. "
            f"Не нужно только при режиме «{MODE_NO_UL}»."
        )
    if header.startswith("Режим распределения"):
        return (
            f"Обязательно (*) для блока {n}. Допустимо: {MODE_WHOLE}, {MODE_FILTER}, "
            f"{MODE_NO_UL}, {MODE_NEEDS_SPLIT}."
        )
    if header.startswith("Фильтр титула"):
        return (
            f"Условно обязательно (†) для блока {n} при режиме «{MODE_FILTER}». "
            "Оба фильтра должны быть заданы."
        )
    if header.startswith("Фильтр марки"):
        return (
            f"Условно обязательно (†) для блока {n} при режиме «{MODE_FILTER}». "
            "Оба фильтра должны быть заданы."
        )
    return "См. правила шапки реестра ДС."


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------


def _issue(
    code: str,
    level: IssueLevel,
    message: str,
    row: DsRegistryRow | None = None,
    *,
    group_id: str = "",
    block_index: int | None = None,
    blocks_overlay: bool = False,
    field: str = "",
) -> DsRegistryIssue:
    if level == "ERROR":
        blocks_overlay = True
    return DsRegistryIssue(
        code=code,
        level=level,
        message=message,
        excel_row=None if row is None else row.excel_row,
        source_id="" if row is None else row.source_id,
        group_id=group_id,
        block_index=block_index,
        blocks_overlay=blocks_overlay,
        field=field,
    )


def _row_nonempty_for_status(row: DsRegistryRow) -> bool:
    if (
        row.status
        or row.source_id
        or row.previous_ds
        or row.ds_file
        or row.revision
        or row.note
    ):
        return True
    return any(not rel.is_empty() for rel in row.relations)


def _append_ul_identity_issues(
    issues: list[DsRegistryIssue],
    row: DsRegistryRow,
    rel: DsRegistryRelation,
) -> None:
    """Require a parseable UL folder to match group id and RFP key.

    ``Нет УЛ`` is skipped (there is no folder). Unparsed folders are always
    an error, even when group_id is already filled.
    """

    if rel.mode == MODE_NO_UL or not rel.ul_folder:
        return
    ident = parse_ul_folder_ds_identity(rel.ul_folder)
    if ident.actual is None:
        issues.append(
            _issue(
                ISSUE_UL_UNPARSED,
                "ERROR",
                f"не разобрана папка УЛ {rel.ul_folder!r}",
                row,
                group_id=rel.group_id,
                block_index=rel.block_index,
            )
        )
        return
    expected_group = canonical_supply_group_id(ident.actual)
    expected_key = str(ident.actual)
    if rel.group_id != expected_group or rel.rfp_key != expected_key:
        issues.append(
            _issue(
                ISSUE_UL_IDENTITY_MISMATCH,
                "ERROR",
                (
                    f"блок {rel.block_index}: папка УЛ {rel.ul_folder!r} даёт "
                    f"группу {expected_group!r} и ключ RFP {expected_key!r}, "
                    f"в строке {rel.group_id!r} / {rel.rfp_key!r}"
                ),
                row,
                group_id=rel.group_id,
                block_index=rel.block_index,
            )
        )


def scan_ul_catalog(
    ul_root: str | Path | None,
) -> tuple[dict[str, tuple[str, ...]], frozenset[str]]:
    """First-level UL folders: actual-number map and every folder name.

    Names are casefolded and whitespace-collapsed, the same way a registry
    cell is compared. A missing root yields empty results.
    """

    if not ul_root:
        return {}, frozenset()
    root = Path(ul_root)
    found: dict[str, list[str]] = {}
    names: set[str] = set()
    try:
        children = [path for path in root.iterdir() if path.is_dir()]
    except OSError:
        return {}, frozenset()
    for folder in children:
        names.add(_norm_link(folder.name))
        identity = parse_ul_folder_ds_identity(folder.name)
        if identity.actual is None:
            continue
        found.setdefault(str(identity.actual), []).append(folder.name)
    return {key: tuple(items) for key, items in found.items()}, frozenset(names)


def index_ul_folders(ul_root: str | Path | None) -> dict[str, tuple[str, ...]]:
    """Map a UL actual number to first-level folder names under ``ul_root``."""

    mapped, _names = scan_ul_catalog(ul_root)
    return mapped


def index_rfp_files(rfp_root: str | Path | None) -> dict[str, tuple[Path, ...]]:
    """Map an RFP actual number to workbook paths in the folder root only."""

    if not rfp_root:
        return {}
    root = Path(rfp_root)
    found: dict[str, list[Path]] = {}
    try:
        children = [
            path
            for path in root.iterdir()
            if path.is_file()
            and path.suffix.lower() in {".xlsx", ".xlsm"}
            and not path.name.startswith("~$")
        ]
    except OSError:
        return {}
    for path in children:
        identity = parse_rfp_ds_identity(path.name)
        if identity.actual is None:
            continue
        found.setdefault(str(identity.actual), []).append(path)
    return {key: tuple(paths) for key, paths in found.items()}


def _norm_link(text: str) -> str:
    return _WS_RE.sub(" ", str(text or "").strip()).casefold()


def _claiming_links(row: DsRegistryRow) -> list[DsRegistryRelation]:
    return [
        rel
        for rel in row.relations
        if rel.ul_folder or rel.rfp_key or rel.rfp_file
    ]


def collapse_registry_rows(
    rows: Sequence[DsRegistryRow],
    *,
    compare_ds_file: bool = True,
) -> tuple[list[DsRegistryRow], list[DsRegistryIssue]]:
    """Merge sheet rows that share one actual number into one logical DS.

    A repeated actual number is the normal form. Status, historical number
    and DS file must agree across those rows. Each non-empty link stays.
    ``compare_ds_file`` is false when the check is about to rewrite «Файл ДС»
    from the folder, so a stale disagreement in that cell is not an error.
    """

    issues: list[DsRegistryIssue] = []
    buckets: dict[str, list[DsRegistryRow]] = {}
    for row in rows:
        if row.source_id:
            buckets.setdefault(row.source_id, []).append(row)

    merged_by_id: dict[str, DsRegistryRow] = {}
    for source_id, group in buckets.items():
        statuses = {item.status for item in group if item.status}
        previous = {item.previous_ds for item in group if item.previous_ds}
        files = (
            {_norm_link(item.ds_file) for item in group if item.ds_file}
            if compare_ds_file
            else set()
        )
        group_id = canonical_supply_group_id(source_id)
        if len(statuses) > 1 or len(previous) > 1 or len(files) > 1:
            issues.append(
                _issue(
                    ISSUE_IDENTITY_SPLIT,
                    "ERROR",
                    (
                        f"у номера {format_registry_ds_number(source_id)} на разных "
                        "строках разошлись статус, исторический номер или файл ДС"
                    ),
                    group[-1],
                    group_id=group_id,
                    field="status",
                )
            )
        first = group[0]
        relations: list[DsRegistryRelation] = []
        for item in group:
            for rel in item.relations:
                if not (rel.ul_folder or rel.rfp_key or rel.rfp_file or rel.mode):
                    continue
                relations.append(
                    replace(
                        rel,
                        group_id=group_id,
                        excel_row=rel.excel_row or item.excel_row,
                    )
                )
        merged_by_id[source_id] = DsRegistryRow(
            status=first.status,
            source_id=source_id,
            previous_ds=next(
                (item.previous_ds for item in group if item.previous_ds),
                first.previous_ds,
            ),
            ds_file=next(
                (item.ds_file for item in group if item.ds_file),
                first.ds_file,
            ),
            revision=first.revision,
            note=first.note,
            relations=tuple(relations),
            excel_row=first.excel_row,
            original_excel_row=first.original_excel_row,
        )

    ordered: list[DsRegistryRow] = []
    seen: set[str] = set()
    for row in rows:
        if not row.source_id:
            ordered.append(row)
            continue
        if row.source_id in seen:
            continue
        seen.add(row.source_id)
        ordered.append(merged_by_id[row.source_id])
    return ordered, issues


def _iter_named_workbooks(root: Path, *, recursive: bool) -> list[Path]:
    try:
        children = root.rglob("*") if recursive else root.iterdir()
        found = [
            path
            for path in children
            if path.is_file()
            and path.suffix.lower() in {".xlsx", ".xlsm"}
            and not path.name.startswith("~$")
        ]
    except OSError:
        return []
    return found


def _ds_file_names(text: str) -> list[str]:
    """Split a «Файл ДС» or «Файл RFP» cell into one filename per line."""

    names: list[str] = []
    for part in str(text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        name = part.strip()
        if name:
            names.append(name)
    return names


def _join_ds_file_names(paths: Sequence[Path]) -> str:
    ordered = sorted({path.name for path in paths}, key=str.casefold)
    return "\n".join(ordered)


def _ds_catalog_ready(ds_root: str | Path | None) -> bool:
    if not ds_root:
        return False
    try:
        return Path(ds_root).is_dir()
    except OSError:
        return False


def drop_robot_placeholder_rows(
    rows: Sequence[DsRegistryRow],
) -> list[DsRegistryRow]:
    """Drop yellow rows a previous check wrote: no number and a «сирота:» note.

    A row a person started stays, including one that already has a status
    but still has no actual number.
    """

    kept: list[DsRegistryRow] = []
    for row in rows:
        note = (row.note or "").lstrip()
        if (
            not row.source_id
            and not row.status
            and note.startswith("сирота:")
        ):
            continue
        kept.append(row)
    return kept


def _ds_workbook_name_skipped(name: str) -> bool:
    """True for locks, registry copies, manual summaries and own DS reports.

    Name only. The registry check must not open the workbook or hash its bytes;
    that walk belongs to the DS audit.
    """

    from RFQ.rfp_parts.ds_baseline import _SKIP_EXACT_NAMES, _SKIP_NAME_PREFIXES

    if name.startswith("~$"):
        return True
    folded = name.casefold()
    stem = Path(name).stem.casefold()
    if folded in _SKIP_EXACT_NAMES:
        return True
    return any(
        stem.startswith(prefix) or folded.startswith(prefix)
        for prefix in _SKIP_NAME_PREFIXES
    )


def _list_ds_workbook_names(ds_root: Path) -> list[tuple[Path, str]]:
    """List DS workbooks as ``(path, relative posix path)`` without reading them."""

    found: list[tuple[Path, str]] = []
    try:
        candidates = ds_root.rglob("*")
    except OSError:
        return found
    for path in candidates:
        try:
            if not path.is_file():
                continue
        except OSError:
            continue
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            continue
        if _ds_workbook_name_skipped(path.name):
            continue
        try:
            rel = path.relative_to(ds_root).as_posix()
        except ValueError:
            rel = path.name
        found.append((path, rel))
    return found


def _workbooks_by_source(
    rows: Sequence[DsRegistryRow],
    ds_root: str | Path,
) -> tuple[dict[str, list[Path]], list[Path]]:
    """Map each registry number to DS workbooks that resolve to it.

    Registry copies, manual summaries and Excel locks are skipped by name.
    The second list is workbooks that resolve to no number. File bytes are
    not read.
    """

    from RFQ.rfp_parts.ds_baseline import resolve_ds_source_id

    active_ids = [row.source_id for row in rows if row.source_id]
    found: dict[str, list[Path]] = {}
    unresolved: list[Path] = []
    for path, rel in _list_ds_workbook_names(Path(ds_root)):
        resolution = resolve_ds_source_id(rel, active_ids)
        if resolution.source_id:
            found.setdefault(resolution.source_id, []).append(path)
        else:
            unresolved.append(path)
    return found, unresolved


def sync_ds_files_from_folder(
    rows: Sequence[DsRegistryRow],
    ds_root: str | Path | None,
    *,
    catalog: tuple[dict[str, list[Path]], list[Path]] | None = None,
) -> list[DsRegistryRow]:
    """Rewrite «Файл ДС» from the folder. One filename per line.

    A number with no resolved workbook gets an empty cell, so a stale name
    does not stay. When the folder has no DS workbooks the cells are left
    as they were.
    """

    if catalog is None:
        if not _ds_catalog_ready(ds_root):
            return list(rows)
        assert ds_root is not None
        catalog = _workbooks_by_source(rows, ds_root)
    found, unresolved = catalog
    if not found and not unresolved:
        return list(rows)
    synced: list[DsRegistryRow] = []
    for row in rows:
        if not row.source_id:
            synced.append(row)
            continue
        synced.append(
            replace(
                row,
                ds_file=_join_ds_file_names(found.get(row.source_id, [])),
            )
        )
    return synced


def _lap(
    on_lap: Callable[[str, float], None] | None,
    label: str,
    started: float,
) -> None:
    if on_lap is None:
        return
    on_lap(label, time.perf_counter() - started)


def collect_orphan_rows(
    rows: Sequence[DsRegistryRow],
    *,
    ul_root: str | Path | None = None,
    rfp_root: str | Path | None = None,
    ds_root: str | Path | None = None,
    ds_catalog: tuple[dict[str, list[Path]], list[Path]] | None = None,
    on_lap: Callable[[str, float], None] | None = None,
) -> list[DsRegistryRow]:
    """Return yellow placeholder rows for folders and files absent from the sheet.

    A named-but-missing UL folder is not an orphan. A DS workbook whose name
    is already on its number (one filename per line) is not an orphan. A file
    that resolves to a number but is missing from that cell is a placeholder,
    as is a workbook that resolves to no number. Registry copies, manual
    summaries and Excel locks are not placeholders.
    """

    claimed_folders = {
        _norm_link(rel.ul_folder)
        for row in rows
        for rel in row.relations
        if rel.ul_folder
    }
    claimed_rfp_files = {
        _norm_link(name)
        for row in rows
        for rel in row.relations
        for name in _ds_file_names(rel.rfp_file)
    }
    claimed_rfp_keys = {
        rel.rfp_key
        for row in rows
        for rel in row.relations
        if rel.rfp_key
    }
    claimed_ds_files = {
        _norm_link(name)
        for row in rows
        for name in _ds_file_names(row.ds_file)
    }
    extras: list[DsRegistryRow] = []

    if ul_root:
        started = time.perf_counter()
        root = Path(ul_root)
        try:
            folders = [path for path in root.iterdir() if path.is_dir()]
        except OSError:
            folders = []
        for folder in folders:
            if _norm_link(folder.name) in claimed_folders:
                continue
            extras.append(
                DsRegistryRow(
                    status="",
                    source_id="",
                    note="сирота: папка УЛ не названа ни на одной строке",
                    relations=(
                        DsRegistryRelation(
                            group_id="",
                            rfp_key="",
                            ul_folder=folder.name,
                            mode="",
                        ),
                    ),
                )
            )
        _lap(on_lap, "сироты УЛ", started)

    if rfp_root:
        started = time.perf_counter()
        for path in _iter_named_workbooks(Path(rfp_root), recursive=False):
            if _norm_link(path.name) in claimed_rfp_files:
                continue
            identity = parse_rfp_ds_identity(path.name)
            key = str(identity.actual) if identity.actual is not None else ""
            if key and key in claimed_rfp_keys:
                continue
            extras.append(
                DsRegistryRow(
                    status="",
                    source_id="",
                    note="сирота: файл RFP не назван ни на одной строке",
                    relations=(
                        DsRegistryRelation(
                            group_id="",
                            rfp_key=key,
                            ul_folder="",
                            mode="",
                            rfp_file=path.name,
                        ),
                    ),
                )
            )
        _lap(on_lap, "сироты RFP", started)

    if ds_catalog is not None or _ds_catalog_ready(ds_root):
        started = time.perf_counter()
        if ds_catalog is not None:
            found, unresolved = ds_catalog
        else:
            assert ds_root is not None
            found, unresolved = _workbooks_by_source(rows, ds_root)
        for path in unresolved:
            if _norm_link(path.name) in claimed_ds_files:
                continue
            extras.append(
                DsRegistryRow(
                    status="",
                    source_id="",
                    ds_file=path.name,
                    note="сирота: файл ДС не назван ни на одной строке",
                    relations=(),
                )
            )
        for source_id, paths in found.items():
            owners = [row for row in rows if row.source_id == source_id]
            listed = {
                _norm_link(name)
                for row in owners
                for name in _ds_file_names(row.ds_file)
            }
            for path in paths:
                if (
                    _norm_link(path.name) in listed
                    or _norm_link(path.name) in claimed_ds_files
                ):
                    continue
                extras.append(
                    DsRegistryRow(
                        status="",
                        source_id="",
                        ds_file=path.name,
                        note=(
                            "сирота: файл ДС не записан в строке номера "
                            f"{format_registry_ds_number(source_id)}"
                        ),
                        relations=(),
                    )
                )
        _lap(on_lap, "сироты ДС", started)
    return extras


def validate_registry_rows(
    rows: Sequence[DsRegistryRow],
    *,
    ul_root: str | Path | None = None,
    rfp_root: str | Path | None = None,
    max_blocks: int | None = None,
) -> DsRegistryValidation:
    """Validate logical rows. A repeated actual number is not an error.

    The same RFP number or RFP file on two actual numbers is an error, as is
    a UL folder repeated inside one number. Two actual numbers may name the
    same UL folder. A named UL folder that is absent from disk is painted
    blue and does not stop the summary. A named RFP file that is absent is a
    warning: it stays out of the RFP bag.
    """

    del max_blocks  # wide-sheet width is not a rule anymore
    issues: list[DsRegistryIssue] = []
    folder_owners: dict[str, set[str]] = {}
    rfp_owner: dict[str, str] = {}
    file_owner: dict[str, str] = {}
    rfp_index = index_rfp_files(rfp_root) if rfp_root else {}
    rfp_names = {
        path.name.casefold()
        for paths in rfp_index.values()
        for path in paths
    }

    for row in rows:
        if not _row_nonempty_for_status(row):
            continue
        links = _claiming_links(row)
        if row.status and row.status not in ALLOWED_STATUSES:
            issues.append(
                _issue(
                    ISSUE_INVALID_STATUS,
                    "ERROR",
                    f"недопустимый статус {row.status!r}; "
                    f"ожидается {', '.join(sorted(ALLOWED_STATUSES))}",
                    row,
                    field="status",
                )
            )
        if not row.source_id:
            if links or row.ds_file:
                issues.append(
                    _issue(
                        ISSUE_ORPHAN,
                        "ERROR",
                        (
                            "строка без актуального номера: проставьте номер и статус "
                            f"«{STATUS_ACTIVE}» или «{STATUS_DISABLED}»"
                        ),
                        row,
                        field="row",
                    )
                )
            elif row.is_active:
                issues.append(
                    _issue(
                        ISSUE_MISSING_SOURCE_ID,
                        "ERROR",
                        "активная строка без актуального номера ДС",
                        row,
                        field="row",
                    )
                )
            continue
        if row.is_active and not links:
            issues.append(
                _issue(
                    ISSUE_MISSING_RELATION,
                    "ERROR",
                    "у активной строки нет папки УЛ и номера RFP",
                    row,
                    field="ul_folder",
                )
            )

        seen_folders: set[str] = set()
        seen_rfp: set[str] = set()
        seen_files: set[str] = set()
        group_id = canonical_supply_group_id(row.source_id) if row.source_id else ""
        for rel in links:
            excel_row = rel.excel_row or row.excel_row
            painted = replace(row, excel_row=excel_row) if excel_row else row
            if rel.ul_folder:
                key = _norm_link(rel.ul_folder)
                repeated_here = key in seen_folders
                if repeated_here:
                    issues.append(
                        _issue(
                            ISSUE_DUPLICATE_LINK,
                            "ERROR",
                            f"папка УЛ {rel.ul_folder!r} повторена внутри номера",
                            painted,
                            group_id=group_id,
                            block_index=rel.block_index,
                            field="ul_folder",
                        )
                    )
                seen_folders.add(key)
                owners = folder_owners.setdefault(key, set())
                if row.source_id in owners and not repeated_here:
                    issues.append(
                        _issue(
                            ISSUE_DUPLICATE_LINK,
                            "ERROR",
                            f"папка УЛ {rel.ul_folder!r} повторена внутри номера",
                            painted,
                            group_id=group_id,
                            block_index=rel.block_index,
                            field="ul_folder",
                        )
                    )
                owners.add(row.source_id)
            if rel.rfp_key:
                repeated_here = rel.rfp_key in seen_rfp
                if repeated_here:
                    issues.append(
                        _issue(
                            ISSUE_DUPLICATE_LINK,
                            "ERROR",
                            f"номер RFP {rel.rfp_key!r} повторен внутри номера ДС",
                            painted,
                            group_id=group_id,
                            block_index=rel.block_index,
                            field="rfp_key",
                        )
                    )
                seen_rfp.add(rel.rfp_key)
                owner = rfp_owner.get(rel.rfp_key)
                if owner and owner != row.source_id:
                    issues.append(
                        _issue(
                            ISSUE_SHARED_RFP,
                            "ERROR",
                            (
                                f"номер RFP {rel.rfp_key!r} указан у номеров "
                                f"{format_registry_ds_number(owner)} и "
                                f"{format_registry_ds_number(row.source_id)}"
                            ),
                            painted,
                            group_id=group_id,
                            block_index=rel.block_index,
                            field="rfp_key",
                        )
                    )
                elif owner == row.source_id and not repeated_here:
                    issues.append(
                        _issue(
                            ISSUE_DUPLICATE_LINK,
                            "ERROR",
                            f"номер RFP {rel.rfp_key!r} повторен внутри номера ДС",
                            painted,
                            group_id=group_id,
                            block_index=rel.block_index,
                            field="rfp_key",
                        )
                    )
                else:
                    rfp_owner[rel.rfp_key] = row.source_id
            file_names = _ds_file_names(rel.rfp_file)
            if file_names:
                for file_name in file_names:
                    file_key = _norm_link(file_name)
                    repeated_here = file_key in seen_files
                    if repeated_here:
                        issues.append(
                            _issue(
                                ISSUE_DUPLICATE_LINK,
                                "ERROR",
                                f"файл RFP {file_name!r} повторен внутри номера",
                                painted,
                                group_id=group_id,
                                field="rfp_file",
                            )
                        )
                    seen_files.add(file_key)
                    owner = file_owner.get(file_key)
                    if owner and owner != row.source_id:
                        issues.append(
                            _issue(
                                ISSUE_SHARED_RFP,
                                "ERROR",
                                (
                                    f"файл RFP {file_name!r} указан у номеров "
                                    f"{format_registry_ds_number(owner)} и "
                                    f"{format_registry_ds_number(row.source_id)}"
                                ),
                                painted,
                                group_id=group_id,
                                field="rfp_file",
                            )
                        )
                    elif owner == row.source_id and not repeated_here:
                        issues.append(
                            _issue(
                                ISSUE_DUPLICATE_LINK,
                                "ERROR",
                                f"файл RFP {file_name!r} повторен внутри номера",
                                painted,
                                group_id=group_id,
                                field="rfp_file",
                            )
                        )
                    else:
                        file_owner[file_key] = row.source_id
                    if rfp_root and file_key not in rfp_names:
                        issues.append(
                            _issue(
                                ISSUE_RFP_FILE_MISSING,
                                "WARN",
                                f"нет файла RFP {file_name!r}",
                                painted,
                                group_id=group_id,
                                field="rfp_file",
                            )
                        )
            elif rel.rfp_key and rfp_root and rel.rfp_key not in rfp_index:
                issues.append(
                    _issue(
                        ISSUE_RFP_FILE_MISSING,
                        "WARN",
                        f"в корне RFP нет файла с номером {rel.rfp_key}",
                        painted,
                        group_id=group_id,
                        field="rfp_key",
                    )
                )

    return DsRegistryValidation(issues=issues)




# ---------------------------------------------------------------------------
# Legacy reader / migrator
# ---------------------------------------------------------------------------


def _cell_at(values: Sequence[object], index: int) -> object:
    if index < 0 or index >= len(values):
        return None
    return values[index]


def _find_legacy_header(
    rows: Iterator[tuple[int, tuple[object, ...]]],
) -> tuple[int, dict[str, int]] | None:
    wanted = {
        normalize_header(LEGACY_HEADER_CURRENT): "current",
        normalize_header(LEGACY_HEADER_OLD): "old",
        normalize_header(LEGACY_HEADER_UL): "ul",
        normalize_header(LEGACY_HEADER_NOTE): "note",
    }
    for excel_row, values in rows:
        mapping: dict[str, int] = {}
        for col, value in enumerate(values):
            key = normalize_header(value)
            role = wanted.get(key)
            if role and role not in mapping:
                mapping[role] = col
        if len(mapping) == 4:
            return excel_row, mapping
    return None


def detect_registry_format(path: str | Path) -> RegistryFormat:
    """Return ``new``, ``legacy`` or ``unknown`` for an existing workbook."""

    workbook = _open_workbook(Path(path), data_only=True)
    try:
        if REGISTRY_SHEET_NAME in workbook.sheetnames:
            sheet = workbook[REGISTRY_SHEET_NAME]
            header = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
            if header:
                names = {normalize_header(v) for v in header if v}
                if "Статус" in names and (
                    "Актуальный номер ДС" in names
                    or "Номер ДС" in names
                    or "ID ДС источника" in names
                ):
                    return "new"
        sheet = workbook[workbook.sheetnames[0]]
        sample = []
        for excel_row, values in enumerate(
            sheet.iter_rows(max_row=20, values_only=True), start=1
        ):
            sample.append((excel_row, tuple(values or ())))
        if _find_legacy_header(iter(sample)):
            return "legacy"
        return "unknown"
    finally:
        workbook.close()


def read_legacy_registry(path: str | Path) -> list[LegacyRegistryRow]:
    """Read the 4-column ``Актуальный ДС / Старый ДС / УЛ / Примечание`` book."""

    source = Path(path)
    if not source.is_file():
        raise DsRegistryFormatError(f"файл реестра не найден: {source}")
    workbook = _open_workbook(source, data_only=True)
    try:
        sheet = workbook[workbook.sheetnames[0]]
        buffered: list[tuple[int, tuple[object, ...]]] = []
        for excel_row, values in enumerate(sheet.iter_rows(values_only=True), start=1):
            buffered.append((excel_row, tuple(values or ())))
    finally:
        workbook.close()

    found = _find_legacy_header(iter(buffered))
    if found is None:
        raise DsRegistryFormatError(
            "не найдены заголовки «Актуальный ДС», «Старый ДС», «УЛ», «Примечание»"
        )
    header_row, mapping = found
    rows: list[LegacyRegistryRow] = []
    for excel_row, values in buffered:
        if excel_row <= header_row:
            continue
        current = cell_text(_cell_at(values, mapping["current"]))
        old = cell_text(_cell_at(values, mapping["old"]))
        ul_text = cell_text(_cell_at(values, mapping["ul"]))
        note = cell_text(_cell_at(values, mapping["note"]))
        if not (current or old or ul_text or note):
            continue
        rows.append(
            LegacyRegistryRow(
                excel_row=excel_row,
                current_ds=current,
                old_ds=old,
                ul_text=ul_text,
                note=note,
            )
        )
    return rows


def _relation_from_ul_folder(
    folder: str,
    *,
    block_index: int,
    mode: str,
) -> tuple[DsRegistryRelation, DsRegistryIssue | None]:
    ident: DsIdentity = parse_ul_folder_ds_identity(folder)
    group_id = ""
    rfp_key = ""
    issue: DsRegistryIssue | None = None
    if ident.actual is not None:
        group_id = canonical_supply_group_id(ident.actual)
        rfp_key = str(ident.actual)
    else:
        issue = DsRegistryIssue(
            code=ISSUE_UL_UNPARSED,
            level="ERROR",
            message=f"не разобрана папка УЛ {folder!r}",
            block_index=block_index,
        )
    relation = DsRegistryRelation(
        group_id=group_id,
        rfp_key=rfp_key,
        ul_folder=folder,
        mode=mode,
        block_index=block_index,
    )
    return relation, issue


def migrate_legacy_rows(
    legacy_rows: Sequence[LegacyRegistryRow],
) -> tuple[list[DsRegistryRow], list[DsRegistryIssue]]:
    """Convert legacy rows to the canonical contract without guessing IDs."""

    rows: list[DsRegistryRow] = []
    issues: list[DsRegistryIssue] = []
    for item in legacy_rows:
        if not item.current_ds:
            rows.append(
                DsRegistryRow(
                    status=STATUS_HISTORY,
                    source_id="",
                    previous_ds=item.old_ds,
                    revision="",
                    note=item.note,
                    relations=(),
                    original_excel_row=item.excel_row,
                    excel_row=item.excel_row,
                )
            )
            continue

        folders = split_ul_folders(item.ul_text)
        relations: list[DsRegistryRelation] = []
        if not folders:
            relations.append(
                DsRegistryRelation(
                    group_id=no_ul_group_id(item.current_ds),
                    rfp_key=item.current_ds,
                    ul_folder="",
                    mode=MODE_NO_UL,
                    block_index=1,
                )
            )
        elif len(folders) == 1:
            relation, issue = _relation_from_ul_folder(
                folders[0], block_index=1, mode=MODE_WHOLE
            )
            if issue is not None:
                issues.append(
                    DsRegistryIssue(
                        code=issue.code,
                        level=issue.level,
                        message=issue.message,
                        excel_row=item.excel_row,
                        source_id=item.current_ds,
                        block_index=1,
                    )
                )
            relations.append(relation)
        else:
            for index, folder in enumerate(folders, start=1):
                relation, issue = _relation_from_ul_folder(
                    folder, block_index=index, mode=MODE_NEEDS_SPLIT
                )
                if issue is not None:
                    issues.append(
                        DsRegistryIssue(
                            code=issue.code,
                            level=issue.level,
                            message=issue.message,
                            excel_row=item.excel_row,
                            source_id=item.current_ds,
                            block_index=index,
                        )
                    )
                relations.append(relation)
            for relation in relations:
                issues.append(
                    DsRegistryIssue(
                        code=ISSUE_MULTI_LINK_NEEDS_SPLIT,
                        level="OVERLAY",
                        message=(
                            "несколько папок УЛ: overlay заблокирован до фильтров"
                        ),
                        excel_row=item.excel_row,
                        source_id=item.current_ds,
                        group_id=relation.group_id,
                        block_index=relation.block_index,
                        blocks_overlay=True,
                    )
                )

        rows.append(
            DsRegistryRow(
                status=STATUS_ACTIVE,
                source_id=item.current_ds,
                previous_ds=item.old_ds,
                revision="",
                note=item.note,
                relations=tuple(relations),
                original_excel_row=item.excel_row,
                excel_row=item.excel_row,
            )
        )
    return rows, issues


# ---------------------------------------------------------------------------
# Writer
# ---------------------------------------------------------------------------


def _make_comment(text: str) -> Comment:
    comment = Comment(text, _COMMENT_AUTHOR)
    comment.width = 280
    comment.height = 140
    return comment


def _block_side(block_index: int) -> Side:
    color = _BLOCK_BORDER_COLORS[(block_index - 1) % len(_BLOCK_BORDER_COLORS)]
    return Side(style="medium", color=color)


def _apply_header_cell(cell, header: str) -> None:
    kind = _header_kind(header)
    cell.value = header or None
    cell.alignment = _ALIGN_HEADER
    cell.comment = _make_comment(_header_comment(header)) if header else None
    if kind == "required":
        cell.fill = _FILL_REQUIRED
        cell.font = _FONT_REQUIRED
    elif kind == "conditional":
        cell.fill = _FILL_CONDITIONAL
        cell.font = _FONT_CONDITIONAL
    else:
        cell.fill = _FILL_AUX
        cell.font = _FONT_AUX
    cell.border = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _apply_block_borders(
    ws: Worksheet,
    layout: _HeaderLayout,
    last_row: int,
) -> None:
    for block in layout.blocks:
        side = _block_side(block.index)
        cols = block.columns()
        first, last = cols[0], cols[-1]
        for row in range(1, last_row + 1):
            for col in cols:
                cell = ws.cell(row=row, column=col)
                left = side if col == first else (cell.border.left or _THIN)
                right = side if col == last else (cell.border.right or _THIN)
                top = side if row == 1 else (cell.border.top or _THIN)
                bottom = side if row == last_row else (cell.border.bottom or _THIN)
                cell.border = Border(left=left, right=right, top=top, bottom=bottom)


def _write_file_links(
    ws: Worksheet,
    row: int,
    col: int,
    paths: Sequence[Path],
) -> None:
    """Write filenames only. One file becomes a hyperlink; the path stays hidden."""

    if col <= 0:
        return
    cell = ws.cell(row=row, column=col)
    cell.alignment = _ALIGN_DATA
    cell.number_format = "@"
    files = [path for path in paths if path.name]
    if len(files) == 1:
        cell.value = files[0].name
        cell.hyperlink = str(files[0])
        cell.font = Font(
            name="Calibri", size=11, color="0563C1", underline="single"
        )
        return
    cell.hyperlink = None
    cell.value = "\n".join(path.name for path in files) or None
    cell.font = _FONT_DATA


def _mark_todo(cell) -> None:
    cell.fill = _FILL_TODO


def _paint_human_todo(
    ws: Worksheet,
    layout: _HeaderLayout,
    rows: Sequence[DsRegistryRow],
    links: RegistryLinks,
) -> None:
    """Yellow cells a person still has to finish. Computed file cells stay plain."""

    ul_index = links.ul_by_actual if links.scanned_ul else {}
    for index, row in enumerate(rows):
        if not row.is_active:
            continue
        excel_row = index + 2
        if (
            links.scanned_ds
            and layout.file_ds
            and row.source_id
            and not links.ds_files.get(row.source_id)
        ):
            _mark_todo(ws.cell(row=excel_row, column=layout.file_ds))
        for block in layout.blocks:
            if block.index > len(row.relations):
                continue
            rel = row.relations[block.index - 1]
            if rel.is_empty():
                continue
            if not rel.group_id:
                _mark_todo(ws.cell(row=excel_row, column=block.group_id))
            if not rel.rfp_key:
                _mark_todo(ws.cell(row=excel_row, column=block.rfp_key))
            if not rel.mode:
                _mark_todo(ws.cell(row=excel_row, column=block.mode))
            if rel.mode == MODE_NEEDS_SPLIT:
                _mark_todo(ws.cell(row=excel_row, column=block.mode))
                if not rel.title_filter:
                    _mark_todo(ws.cell(row=excel_row, column=block.title_filter))
                if not rel.mark_filter:
                    _mark_todo(ws.cell(row=excel_row, column=block.mark_filter))
            if rel.mode == MODE_FILTER:
                if not rel.title_filter:
                    _mark_todo(ws.cell(row=excel_row, column=block.title_filter))
                if not rel.mark_filter:
                    _mark_todo(ws.cell(row=excel_row, column=block.mark_filter))
            if rel.mode and rel.mode != MODE_NO_UL and not rel.ul_folder:
                _mark_todo(ws.cell(row=excel_row, column=block.ul_folder))
            if (
                rel.mode == MODE_NO_UL
                and rel.rfp_key
                and ul_index.get(rel.rfp_key)
            ):
                _mark_todo(ws.cell(row=excel_row, column=block.mode))
            if (
                links.scanned_rfp
                and rel.rfp_key
                and block.rfp_status
                and not links.rfp_files.get(rel.rfp_key)
            ):
                _mark_todo(ws.cell(row=excel_row, column=block.rfp_status))


def _add_source_info_sheet(workbook: Workbook, rows: Sequence[DsRegistryRow]) -> None:
    """Keep old-DS notes off the working sheet."""

    if SOURCE_INFO_SHEET_NAME in workbook.sheetnames:
        del workbook[SOURCE_INFO_SHEET_NAME]
    sheet = workbook.create_sheet(SOURCE_INFO_SHEET_NAME)
    headers = (
        "Номер ДС",
        HDR_PREVIOUS,
        HDR_REVISION,
        HDR_NOTE,
        HDR_ORIGINAL_ROW,
    )
    for col, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=col, value=header)
        cell.font = _FONT_AUX
        cell.fill = _FILL_AUX
        cell.alignment = _ALIGN_HEADER
    for index, row in enumerate(rows, start=2):
        _write_text_cell(sheet, index, 1, format_registry_ds_number(row.source_id))
        _write_text_cell(sheet, index, 2, row.previous_ds)
        _write_text_cell(sheet, index, 3, row.revision)
        _write_text_cell(sheet, index, 4, row.note)
        origin = sheet.cell(
            row=index,
            column=5,
            value=row.original_excel_row if row.original_excel_row else None,
        )
        origin.alignment = _ALIGN_DATA
        origin.font = _FONT_DATA
    sheet.column_dimensions["A"].width = 18
    sheet.column_dimensions["B"].width = 28
    sheet.column_dimensions["C"].width = 14
    sheet.column_dimensions["D"].width = 42
    sheet.column_dimensions["E"].width = 24
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:E{max(1, len(rows) + 1)}"
    sheet.sheet_properties.tabColor = "D9D9D9"


def _visual_lines(text: object, width: float) -> int:
    """How many wrapped lines ``text`` needs in a column of ``width`` characters."""

    raw = "" if text is None else str(text)
    if raw == "":
        return 1
    capacity = max(1, int(width) - 1)
    total = 0
    for part in raw.split("\n"):
        if not part:
            total += 1
            continue
        total += (len(part) + capacity - 1) // capacity
    return max(1, total)


def _row_height(lines: int, *, minimum: float) -> float:
    return max(minimum, lines * _LINE_HEIGHT)


def _existing_flat_widths(path: Path, defaults: tuple[float, ...]) -> tuple[float, ...]:
    """Column widths already saved on the registry sheet, else ``defaults``.

    A dated copy that does not exist yet takes widths from the canonical file
    in the same folder. A filter Excel stored on the sheet is ignored.
    """

    source = path if path.is_file() else path.with_name(CANONICAL_REGISTRY_NAME)
    if not source.is_file():
        return defaults
    try:
        data = _xlsx_without_filter_columns(source.read_bytes())
        workbook = load_workbook(BytesIO(data), data_only=False)
    except (OSError, ValueError, BadZipFile):
        return defaults
    try:
        if REGISTRY_SHEET_NAME not in workbook.sheetnames:
            return defaults
        sheet = workbook[REGISTRY_SHEET_NAME]
        widths: list[float] = []
        for index, default in enumerate(defaults, start=1):
            saved = sheet.column_dimensions[get_column_letter(index)].width
            widths.append(float(saved) if saved else default)
        return tuple(widths)
    finally:
        workbook.close()


def _fit_row(ws: Worksheet, excel_row: int, widths: Sequence[float], *, minimum: float) -> None:
    lines = 1
    for index, width in enumerate(widths, start=1):
        lines = max(lines, _visual_lines(ws.cell(row=excel_row, column=index).value, width))
    ws.row_dimensions[excel_row].height = _row_height(lines, minimum=minimum)


def _write_text_cell(ws: Worksheet, row: int, col: int, value: str) -> None:
    if col <= 0:
        return
    cell = ws.cell(row=row, column=col, value=value if value else None)
    cell.alignment = _ALIGN_DATA
    cell.font = _FONT_DATA
    cell.number_format = "@"


def _add_blank_required_cf(ws: Worksheet, layout: _HeaderLayout, last_row: int) -> None:
    if last_row < 2:
        return
    status_col = get_column_letter(layout.status)
    source_col = get_column_letter(layout.source_id)
    last_letter = get_column_letter(layout.column_count)
    # Status is required when the row has any value.
    ws.conditional_formatting.add(
        f"{status_col}2:{status_col}{last_row}",
        FormulaRule(
            formula=[
                f'AND({status_col}2="",COUNTA($A2:${last_letter}2)>0)'
            ],
            fill=_FILL_BLANK,
        ),
    )
    ws.conditional_formatting.add(
        f"{source_col}2:{source_col}{last_row}",
        FormulaRule(
            formula=[
                f'AND({status_col}2="{STATUS_ACTIVE}",{source_col}2="")'
            ],
            fill=_FILL_BLANK,
        ),
    )
    if layout.is_flat and layout.flat_ul and layout.flat_rfp_number:
        ul = get_column_letter(layout.flat_ul)
        rfp = get_column_letter(layout.flat_rfp_number)
        active = f'${status_col}2="{STATUS_ACTIVE}"'
        for letter in (ul, rfp):
            ws.conditional_formatting.add(
                f"{letter}2:{letter}{last_row}",
                FormulaRule(
                    formula=[f'AND({active},{ul}2="",{rfp}2="")'],
                    fill=_FILL_BLANK,
                ),
            )
        return
    for block in layout.blocks:
        g = get_column_letter(block.group_id)
        r = get_column_letter(block.rfp_key)
        u = get_column_letter(block.ul_folder)
        m = get_column_letter(block.mode)
        t = get_column_letter(block.title_filter)
        k = get_column_letter(block.mark_filter)
        first = get_column_letter(block.columns()[0])
        last = get_column_letter(block.columns()[-1])
        active_and_used = (
            f'AND(${status_col}2="{STATUS_ACTIVE}",COUNTA({first}2:{last}2)>0)'
            if block.index > 1
            else f'AND(${status_col}2="{STATUS_ACTIVE}")'
        )
        for letter in (g, r, m):
            ws.conditional_formatting.add(
                f"{letter}2:{letter}{last_row}",
                FormulaRule(
                    formula=[f"AND({active_and_used},{letter}2=\"\")"],
                    fill=_FILL_BLANK,
                ),
            )
        ws.conditional_formatting.add(
            f"{u}2:{u}{last_row}",
            FormulaRule(
                formula=[
                    f'AND({m}2<>"",{m}2<>"{MODE_NO_UL}",{u}2="")'
                ],
                fill=_FILL_BLANK,
            ),
        )
        for letter in (t, k):
            ws.conditional_formatting.add(
                f"{letter}2:{letter}{last_row}",
                FormulaRule(
                    formula=[
                        f'AND({m}2="{MODE_FILTER}",{letter}2="")'
                    ],
                    fill=_FILL_BLANK,
                ),
            )


def _legend_rows() -> list[tuple[str, str, str, str]]:
    """How-to rows: field, obligation, when to fill, what the robot does."""

    star = "всегда, синий *"
    yellow = "иногда, жёлтый †"
    grey = "не обязательно, серый"
    return [
        (
            "Несколько строк",
            star,
            "Один актуальный номер может занимать несколько строк: каждая строка — одна папка УЛ и/или один файл RFP.",
            "Все строки номера — один мешок спецификации и один набор RFP. Повтор номера — нормальная форма.",
        ),
        (
            "Статус",
            star,
            "Активен, История или Отключен. На всех строках одного номера статус совпадает.",
            "В свод попадают только «Активен».",
        ),
        (
            "Актуальный номер ДС",
            star,
            "Номер файла ДС: ДС11, ДС8, ДС4905_1. Это ключ посадки Шага 4 (столбец «Фактический ДС»).",
            "Число в имени папки и первое число ярлыка ДС14_48 ключ не выбирают.",
        ),
        (
            "Исторический номер ДС",
            grey,
            "Справка. В ключ и в мешок не входит.",
            "Одинаковый на всех строках номера.",
        ),
        (
            "Файл ДС",
            grey,
            "Имя файла. Несколько файлов одного номера — каждое имя с новой строки. Одинаковое на всех строках номера.",
            "Проверка заново собирает ячейку из папки ДС. Файл, который уже записан на свой номер, жёлтым не становится.",
        ),
        (
            "Папка УЛ",
            yellow,
            "Одна папка на строку, либо пусто, если на строке есть номер RFP. Одна папка — только один актуальный номер.",
            "Папки ещё нет на диске — ячейка синяя, свод не останавливается. Папка в каталоге без строки — жёлтая строка, свод не пишется.",
        ),
        (
            "Номер RFP и файл RFP",
            yellow,
            "Несколько имён — каждое с новой строки. Одно и то же имя не повторяется внутри номера и не стоит на разных номерах.",
            "Каждое имя сверяется отдельно. Нет одного из имён — жёлтая ячейка, в мешок RFP не входит только оно. Точный дубль файла блокирует группу.",
        ),
        (
            "Сверка RFP",
            grey,
            "Не заполняйте. Робот пишет: совпало, расхождение или нет файла. Одинаково на всех строках номера.",
            "Свод эту ячейку не читает: решение каждый раз считается заново. Полное совпадение кодов, единиц и количеств берёт RFP, иначе ДС.",
        ),
        (
            "Жёлтая строка без номера",
            yellow,
            "Робот добавил папку или файл, которых не было ни на одной строке. Следующая проверка такую строку стирает и, если объект всё ещё не назван, добавляет заново.",
            "Пока не проставлены номер и статус «Активен» или «Отключен», свод не пишется. Строку со статусом проверка не стирает.",
        ),
    ]


def _add_legend_sheet(workbook: Workbook) -> None:
    """Add or replace the second sheet «Как заполнять»."""

    if LEGEND_SHEET_NAME in workbook.sheetnames:
        del workbook[LEGEND_SHEET_NAME]
    sheet = workbook.create_sheet(LEGEND_SHEET_NAME)
    names = workbook.sheetnames
    if names.index(LEGEND_SHEET_NAME) != 1 and REGISTRY_SHEET_NAME in names:
        sheet_index = names.index(LEGEND_SHEET_NAME)
        workbook.move_sheet(sheet, offset=1 - sheet_index)

    title_font = Font(bold=True, name="Calibri", size=16, color="1B4F72")
    intro_font = Font(name="Calibri", size=12)
    header_font = Font(bold=True, name="Calibri", size=11, color="FFFFFF")
    header_fill = PatternFill(fill_type="solid", fgColor="1B4F72")
    section_fill = {
        "всегда, синий *": _FILL_REQUIRED,
        "иногда, жёлтый †": _FILL_CONDITIONAL,
        "не обязательно, серый": _FILL_AUX,
    }
    section_font = {
        "всегда, синий *": _FONT_REQUIRED,
        "иногда, жёлтый †": _FONT_CONDITIONAL,
        "не обязательно, серый": _FONT_AUX,
    }

    sheet["A1"] = "Как заполнять реестр"
    sheet["A1"].font = title_font
    sheet.merge_cells("A1:D1")
    sheet.row_dimensions[1].height = 24

    sheet["A2"] = (
        "Лист «Реестр ДС» читает робот. Этот лист — памятка, на расчёт он не влияет. "
        "Синий заголовок со * заполняйте всегда. Жёлтый с † — только когда в столбце "
        "«Когда заполнять» сказано, что поле нужно. Серый можно не трогать."
    )
    sheet["A2"].font = intro_font
    sheet["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    sheet.merge_cells("A2:D2")
    sheet.row_dimensions[2].height = 48

    headers = ("Поле", "Обязательность", "Когда заполнять", "На что влияет")
    for col, text in enumerate(headers, start=1):
        cell = sheet.cell(row=3, column=col, value=text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = _ALIGN_HEADER
    sheet.row_dimensions[3].height = 22
    sheet.freeze_panes = "A4"
    sheet.auto_filter.ref = f"A3:D{3 + len(_legend_rows())}"

    for offset, (field_name, duty, when, effect) in enumerate(_legend_rows()):
        excel_row = 4 + offset
        values = (field_name, duty, when, effect)
        for col, text in enumerate(values, start=1):
            cell = sheet.cell(row=excel_row, column=col, value=text)
            cell.alignment = _ALIGN_DATA
            cell.font = _FONT_DATA
            cell.border = Border(
                left=_THIN, right=_THIN, top=_THIN, bottom=_THIN
            )
        duty_cell = sheet.cell(row=excel_row, column=2)
        duty_cell.fill = section_fill.get(duty, _FILL_AUX)
        duty_cell.font = section_font.get(duty, _FONT_AUX)
        sheet.row_dimensions[excel_row].height = 48

    sheet.column_dimensions["A"].width = 32
    sheet.column_dimensions["B"].width = 28
    sheet.column_dimensions["C"].width = 62
    sheet.column_dimensions["D"].width = 62
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.print_title_rows = "1:3"


@dataclass(frozen=True, slots=True)
class RegistryInstall:
    """Where the one working registry was written."""

    path: Path
    mode: Literal["canonical", "dated"]
    archived_path: Path | None = None
    archived_now: bool = False


def resolve_latest_registry(directory: str | Path | None = None) -> Path:
    """Newest working registry in ``_RFP``: the main file or a dated copy.

    ``Реестр_ДС_УЛ_old.xlsx`` is never chosen. If nothing exists yet, returns
    the canonical path (the file may be missing).
    """

    home = Path(directory) if directory else DEFAULT_RFP_BASE
    canonical = home / CANONICAL_REGISTRY_NAME
    found: list[Path] = []
    try:
        if canonical.is_file():
            found.append(canonical)
        for path in home.iterdir():
            if path.name.startswith("~$"):
                continue
            if DATED_REGISTRY_RE.match(path.name) and path.is_file():
                found.append(path)
    except OSError:
        return canonical
    if not found:
        return canonical
    return max(found, key=lambda item: item.stat().st_mtime)


def prune_dated_registries(
    directory: str | Path,
    *,
    keep: int = MAX_DATED_REGISTRY_COPIES,
) -> list[Path]:
    """Delete dated registry copies beyond ``keep`` newest stamps."""

    home = Path(directory)
    dated: list[tuple[str, Path]] = []
    try:
        entries = list(home.iterdir())
    except OSError:
        return []
    for path in entries:
        match = DATED_REGISTRY_RE.match(path.name)
        if match and path.is_file():
            dated.append((match.group(1), path))
    dated.sort(reverse=True)
    removed: list[Path] = []
    for _, path in dated[max(0, keep):]:
        try:
            path.unlink()
        except OSError:
            continue
        removed.append(path)
    return removed


def archive_legacy_canonical(home: str | Path) -> tuple[Path | None, bool]:
    """Rename a legacy ``Реестр_ДС_УЛ.xlsx`` to ``Реестр_ДС_УЛ_old.xlsx``.

    Returns:
        ``(archive_path, moved_now)``. A second run leaves an existing ``_old``
        in place. A locked file is not renamed.
    """

    folder = Path(home)
    canonical = folder / CANONICAL_REGISTRY_NAME
    old = folder / OLD_REGISTRY_NAME
    if not canonical.is_file():
        return (old if old.is_file() else None), False
    try:
        if detect_registry_format(canonical) != "legacy":
            return None, False
    except OSError:
        return None, False
    if old.exists():
        return old, False
    try:
        os.replace(canonical, old)
    except OSError:
        return None, False
    return old, True


def _next_dated_registry_path(folder: Path) -> Path:
    """Next free ``Реестр_ДС_УЛ_<YYYYMMDD_HHMMSS>.xlsx`` in ``folder``."""

    stamp_dt = datetime.now()
    for _ in range(8):
        candidate = folder / f"Реестр_ДС_УЛ_{stamp_dt.strftime('%Y%m%d_%H%M%S')}.xlsx"
        if not candidate.exists():
            return candidate
        stamp_dt += timedelta(seconds=1)
    return folder / f"Реестр_ДС_УЛ_{stamp_dt.strftime('%Y%m%d_%H%M%S')}.xlsx"


def install_working_registry(
    rows: Sequence[DsRegistryRow],
    home: str | Path,
    *,
    links: RegistryLinks | None = None,
) -> RegistryInstall:
    """Write the one working registry into ``home`` (the ``_RFP`` folder).

    The canonical name is replaced when it is free. If Excel holds it open, a
    ``Реестр_ДС_УЛ_<дата>.xlsx`` copy is written instead and older dated copies
    beyond five are deleted.

    Args:
        rows: New-format rows, including a migration result.
        home: Directory that contains the working registry. Not a reports stamp.

    Returns:
        The path the program should read next.

    Raises:
        DsRegistryError: Neither the canonical name nor a dated copy could be written.
    """

    folder = Path(home)
    folder.mkdir(parents=True, exist_ok=True)
    staging = folder / f".{Path(CANONICAL_REGISTRY_NAME).stem}.{os.getpid()}.staging.xlsx"
    written = write_registry_workbook(staging, rows, links=links)
    archived, moved_now = archive_legacy_canonical(folder)
    canonical = folder / CANONICAL_REGISTRY_NAME
    try:
        os.replace(written, canonical)
    except PermissionError:
        dated = _next_dated_registry_path(folder)
        try:
            os.replace(written, dated)
        except OSError:
            written.unlink(missing_ok=True)
            raise DsRegistryError(
                f"{CANONICAL_REGISTRY_NAME} открыт в Excel, заменить его нельзя, "
                "и копия с датой тоже не записалась. Закройте файл и повторите проверку."
            )
        prune_dated_registries(folder)
        return RegistryInstall(
            path=dated,
            mode="dated",
            archived_path=archived,
            archived_now=moved_now,
        )
    except OSError:
        written.unlink(missing_ok=True)
        if moved_now and archived is not None and archived.is_file() and not canonical.exists():
            try:
                os.replace(archived, canonical)
            except OSError:
                pass
        raise
    prune_dated_registries(folder)
    return RegistryInstall(
        path=canonical,
        mode="canonical",
        archived_path=archived,
        archived_now=moved_now,
    )


def _publish_workbook(tmp_path: Path, target: Path) -> Path:
    """Replace ``target`` with ``tmp_path``.

    If Excel holds ``target`` open, write ``<имя>_новый_<штамп>.xlsx`` beside it
    and leave the locked file untouched.

    Returns:
        The path that now contains the workbook.

    Raises:
        DsRegistryError: The locked name and the sibling both cannot be written.
    """

    try:
        os.replace(tmp_path, target)
        return target
    except PermissionError:
        pass
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    sibling = target.with_name(f"{target.stem}_новый_{stamp}{target.suffix}")
    try:
        os.replace(tmp_path, sibling)
    except OSError:
        tmp_path.unlink(missing_ok=True)
        raise DsRegistryError(
            f"{target.name} открыт в Excel, заменить его нельзя, "
            f"и соседний файл тоже не записался. Закройте {target.name} "
            "и повторите проверку."
        )
    return sibling


def ensure_registry_legend(path: str | Path) -> bool:
    """Add «Как заполнять» when the new-format workbook does not have it.

    Args:
        path: Registry xlsx to update in place.

    Returns:
        True when the sheet was added.

    Raises:
        OSError: The workbook is open or cannot be replaced.
        DsRegistryFormatError: The file has no sheet «Реестр ДС».
    """

    target = Path(path)
    workbook = load_workbook(target)
    try:
        if REGISTRY_SHEET_NAME not in workbook.sheetnames:
            raise DsRegistryFormatError(
                f"нет листа «{REGISTRY_SHEET_NAME}»: {target.name}"
            )
        if LEGEND_SHEET_NAME in workbook.sheetnames:
            return False
        _add_legend_sheet(workbook)
        buffer = BytesIO()
        workbook.save(buffer)
    finally:
        workbook.close()
    tmp_path = target.with_name(f".{target.stem}.{os.getpid()}.tmp.xlsx")
    try:
        tmp_path.write_bytes(buffer.getvalue())
        os.replace(tmp_path, target)
    except OSError:
        tmp_path.unlink(missing_ok=True)
        raise
    return True


@dataclass(frozen=True, slots=True)
class RegistryLinks:
    """Files and UL folders painted into the working registry sheet."""

    ds_files: dict[str, tuple[Path, ...]] = field(default_factory=dict)
    rfp_files: dict[str, tuple[Path, ...]] = field(default_factory=dict)
    ul_by_actual: dict[str, tuple[str, ...]] = field(default_factory=dict)
    scanned_ds: bool = False
    scanned_rfp: bool = False
    scanned_ul: bool = False
    ul_names: frozenset[str] = field(default_factory=frozenset)
    reconcile_by_source: dict[str, str] = field(default_factory=dict)


def _flat_layout() -> _HeaderLayout:
    return _HeaderLayout(
        headers=FLAT_HEADER_TITLES,
        status=1,
        source_id=2,
        previous=3,
        revision=0,
        note=0,
        original_row=0,
        blocks=(),
        file_ds=4,
        flat_ul=5,
        flat_rfp_number=6,
        flat_rfp_file=7,
        flat_reconcile=8,
    )


def registry_packing_maps(
    document: DsRegistryDocument,
) -> tuple[dict[str, int], dict[int, int]] | None:
    """Return folder and RFP-number maps for a clean registry.

    Conflicts make the document invalid; the caller then keeps the name parser.
    Only integer actual numbers are packing keys. A UL folder named by more
    than one active source is omitted so this one-int map cannot point a
    shared folder at a single DS.
    """

    if not document.validation.is_ok:
        return None
    folder_owners: dict[str, set[int]] = {}
    numbers: dict[int, int] = {}
    for row in document.active_rows:
        if not row.source_id.isdigit():
            continue
        actual = int(row.source_id)
        for rel in row.relations:
            if rel.ul_folder:
                folder_owners.setdefault(rel.ul_folder.casefold(), set()).add(
                    actual
                )
            if rel.rfp_key.isdigit():
                numbers[int(rel.rfp_key)] = actual
    folders = {
        key: next(iter(owners))
        for key, owners in folder_owners.items()
        if len(owners) == 1
    }
    return folders, numbers


def copy_registry_snapshot(source: Path, dest_dir: Path) -> Path | None:
    """Copy the registry workbook into the Step4 result folder.

    The destination name is ``source.name``, so a dated registry stays dated
    and the canonical ``Реестр_ДС_УЛ.xlsx`` keeps that name. A same-path
    request returns that path without copying. An ``OSError`` is printed in
    Russian and does not raise.

    Args:
        source: Registry file that was loaded.
        dest_dir: Step4 result directory.

    Returns:
        Destination path, or ``None`` when the copy could not be written.
    """

    try:
        source = Path(source)
        dest_dir = Path(dest_dir)
        dest = dest_dir / source.name
        if source.resolve() == dest.resolve():
            return dest
        shutil.copy2(source, dest)
    except OSError as exc:
        print(f"УЛ: копия реестра не записана в папку результата: {exc}")
        return None
    return dest


def expand_registry_rows(rows: Sequence[DsRegistryRow]) -> list[DsRegistryRow]:
    """One sheet row per link. Identity columns repeat on every row of the number."""

    expanded: list[DsRegistryRow] = []
    for row in rows:
        links = _claiming_links(row)
        if not links:
            expanded.append(row)
            continue
        for rel in links:
            expanded.append(
                DsRegistryRow(
                    status=row.status,
                    source_id=row.source_id,
                    previous_ds=row.previous_ds,
                    ds_file=row.ds_file,
                    revision=row.revision,
                    note=row.note,
                    relations=(rel,),
                    excel_row=rel.excel_row or row.excel_row,
                    original_excel_row=row.original_excel_row,
                )
            )
    return expanded


def _paint_ul_absent(cell) -> None:
    cell.fill = _FILL_UL_ABSENT
    if cell.comment is None:
        cell.comment = _make_comment(
            "Папки УЛ пока нет на диске. Свод из-за этого не останавливается."
        )


def _paint_cell(cell, comment: str) -> None:
    _mark_todo(cell)
    if comment and cell.comment is None:
        cell.comment = _make_comment(comment)


def _paint_flat_conflicts(
    ws: Worksheet,
    layout: _HeaderLayout,
    rows: Sequence[DsRegistryRow],
    links: RegistryLinks,
) -> None:
    """Yellow cells for conflicts, orphans and a missing named RFP file.

    A UL folder is yellow only when the same number repeats it. Several
    numbers naming one folder are not a conflict. A named UL folder that is
    not on disk is light blue, unless the same cell is already yellow.
    """

    folder_rows: dict[str, list[tuple[int, str]]] = {}
    rfp_rows: dict[str, list[tuple[int, str]]] = {}
    file_rows: dict[str, list[tuple[int, str]]] = {}
    identity: dict[str, list[tuple[int, str, str, str]]] = {}
    rfp_on_disk = {
        path.name.casefold()
        for paths in links.rfp_files.values()
        for path in paths
    }
    for index, row in enumerate(rows):
        excel_row = index + 2
        rel = row.relations[0] if row.relations else None
        folder = rel.ul_folder if rel is not None else ""
        rfp_key = rel.rfp_key if rel is not None else ""
        rfp_file = rel.rfp_file if rel is not None else ""
        if not row.source_id and (folder or rfp_key or rfp_file or row.ds_file):
            comment = row.note or (
                "Проставьте актуальный номер и статус «Активен» или «Отключен»."
            )
            for col in range(1, layout.column_count + 1):
                _paint_cell(ws.cell(row=excel_row, column=col), comment)
            continue
        if row.source_id:
            identity.setdefault(row.source_id, []).append(
                (excel_row, row.status, row.previous_ds, _norm_link(row.ds_file))
            )
        if folder:
            folder_rows.setdefault(_norm_link(folder), []).append(
                (excel_row, row.source_id)
            )
        if rfp_key:
            rfp_rows.setdefault(rfp_key, []).append((excel_row, row.source_id))
        if rfp_file:
            for name in _ds_file_names(rfp_file):
                file_rows.setdefault(_norm_link(name), []).append(
                    (excel_row, row.source_id)
                )
        elif (
            links.scanned_rfp
            and rfp_key
            and rfp_key not in links.rfp_files
        ):
            _paint_cell(
                ws.cell(row=excel_row, column=layout.flat_rfp_number),
                f"В корне RFP нет файла с номером {rfp_key}.",
            )

    def _shared(groups: dict[str, list[tuple[int, str]]], column: int, label: str) -> None:
        for items in groups.values():
            owners = {source for _row, source in items if source}
            if len(owners) > 1 or (len(items) > 1 and len(owners) <= 1):
                comment = (
                    f"{label} повторяется"
                    + (
                        " на разных номерах ДС"
                        if len(owners) > 1
                        else " внутри одного номера"
                    )
                )
                for excel_row, _source in items:
                    _paint_cell(ws.cell(row=excel_row, column=column), comment)

    for items in folder_rows.values():
        owners = {source for _row, source in items if source}
        if len(items) > 1 and len(owners) <= 1:
            comment = "Папка УЛ повторяется внутри одного номера"
            for excel_row, _source in items:
                _paint_cell(
                    ws.cell(row=excel_row, column=layout.flat_ul), comment
                )
    _shared(rfp_rows, layout.flat_rfp_number, "Номер RFP")
    _shared(file_rows, layout.flat_rfp_file, "Файл RFP")
    if links.scanned_rfp:
        for index, row in enumerate(rows):
            rel = row.relations[0] if row.relations else None
            rfp_file = rel.rfp_file if rel is not None else ""
            names = _ds_file_names(rfp_file)
            if not names or not row.source_id:
                continue
            missing = [name for name in names if name.casefold() not in rfp_on_disk]
            if not missing:
                continue
            cell = ws.cell(row=index + 2, column=layout.flat_rfp_file)
            rgb = str(getattr(getattr(cell.fill, "fgColor", None), "rgb", "") or "")
            if rgb.upper().endswith("FFFF00"):
                continue
            if len(missing) == 1:
                comment = (
                    f"Нет файла RFP {missing[0]}. В мешок RFP он не входит."
                )
            else:
                comment = (
                    "Нет файлов RFP:\n"
                    + "\n".join(missing)
                    + "\nВ мешок RFP они не входят."
                )
            _paint_cell(cell, comment)
    if links.scanned_ul:
        for index, row in enumerate(rows):
            rel = row.relations[0] if row.relations else None
            folder = rel.ul_folder if rel is not None else ""
            if not folder or not row.source_id:
                continue
            if _norm_link(folder) in links.ul_names:
                continue
            cell = ws.cell(row=index + 2, column=layout.flat_ul)
            rgb = str(getattr(getattr(cell.fill, "fgColor", None), "rgb", "") or "")
            if rgb.upper().endswith("FFFF00"):
                continue
            _paint_ul_absent(cell)
    for items in identity.values():
        statuses = {item[1] for item in items if item[1]}
        previous = {item[2] for item in items if item[2]}
        files = {item[3] for item in items if item[3]}
        if len(statuses) <= 1 and len(previous) <= 1 and len(files) <= 1:
            continue
        comment = "На строках одного номера разошлись статус, исторический номер или файл ДС."
        for excel_row, _status, _previous, _file in items:
            _paint_cell(ws.cell(row=excel_row, column=layout.status), comment)
            if layout.previous:
                _paint_cell(ws.cell(row=excel_row, column=layout.previous), comment)
            if layout.file_ds:
                _paint_cell(ws.cell(row=excel_row, column=layout.file_ds), comment)


def write_registry_workbook(
    path: str | Path,
    rows: Sequence[DsRegistryRow],
    *,
    spare_rows: int = DEFAULT_SPARE_ROWS,
    links: RegistryLinks | None = None,
) -> Path:
    """Write one sheet row per link, the how-to sheet and the source-info sheet.

    Args:
        path: Destination xlsx. Parent directories are created.
        rows: Logical rows. Several links of one number become several sheet rows.
        spare_rows: Extra empty table rows so a person can append data.
        links: Optional DS/RFP lookup and the reconcile text written into the sheet.
            The reconcile column is not read back into decisions.

    Column widths already stored on the registry sheet are kept. Header cells
    stay centered, wrapped and coloured by role. Data cells are left-aligned,
    vertically centered and wrapped. Row height follows the wrapped line count
    so a multi-line file name stays visible.

    Returns:
        The written path.

    Raises:
        OSError: If the destination cannot be written.
    """

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    layout = _flat_layout()
    sheet_rows = expand_registry_rows(rows)
    last_data = max(len(sheet_rows), 1) + max(0, int(spare_rows))
    last_row = last_data + 1
    link_book = links or RegistryLinks()
    widths = _existing_flat_widths(target, _FLAT_WIDTHS)

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = REGISTRY_SHEET_NAME
    ws.freeze_panes = "A2"

    for col, header in enumerate(layout.headers, start=1):
        _apply_header_cell(ws.cell(row=1, column=col), header)
    for col, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    _fit_row(ws, 1, widths, minimum=_HEADER_MIN_HEIGHT)

    for index, row in enumerate(sheet_rows):
        excel_row = index + 2
        rel = row.relations[0] if row.relations else None
        _write_text_cell(ws, excel_row, layout.status, row.status)
        _write_text_cell(
            ws, excel_row, layout.source_id, format_registry_ds_number(row.source_id)
        )
        _write_text_cell(ws, excel_row, layout.previous, row.previous_ds)
        ds_paths = link_book.ds_files.get(row.source_id, ())
        if row.ds_file:
            matched = tuple(
                path for path in ds_paths if path.name.casefold() == row.ds_file.casefold()
            )
            if matched:
                _write_file_links(ws, excel_row, layout.file_ds, matched[:1])
            else:
                _write_text_cell(ws, excel_row, layout.file_ds, row.ds_file)
        elif len(ds_paths) == 1:
            _write_file_links(ws, excel_row, layout.file_ds, ds_paths)
        rfp_key = rel.rfp_key if rel is not None else ""
        rfp_file = rel.rfp_file if rel is not None else ""
        _write_text_cell(ws, excel_row, layout.flat_ul, rel.ul_folder if rel else "")
        _write_text_cell(ws, excel_row, layout.flat_rfp_number, rfp_key)
        rfp_paths = link_book.rfp_files.get(rfp_key, ()) if rfp_key else ()
        if rfp_file:
            matched_rfp = tuple(
                path for path in rfp_paths if path.name.casefold() == rfp_file.casefold()
            )
            if matched_rfp:
                _write_file_links(ws, excel_row, layout.flat_rfp_file, matched_rfp[:1])
            else:
                _write_text_cell(ws, excel_row, layout.flat_rfp_file, rfp_file)
        elif len(rfp_paths) == 1:
            _write_file_links(ws, excel_row, layout.flat_rfp_file, rfp_paths)
        _write_text_cell(
            ws,
            excel_row,
            layout.flat_reconcile,
            link_book.reconcile_by_source.get(row.source_id, ""),
        )
        _fit_row(ws, excel_row, widths, minimum=_DATA_MIN_HEIGHT)

    for excel_row in range(len(sheet_rows) + 2, last_row + 1):
        for col in range(1, layout.column_count + 1):
            cell = ws.cell(row=excel_row, column=col, value=None)
            cell.alignment = _ALIGN_DATA
            cell.font = _FONT_DATA
            cell.number_format = "@"

    status_dv = DataValidation(
        type="list",
        formula1='"' + ",".join((STATUS_ACTIVE, STATUS_HISTORY, STATUS_DISABLED)) + '"',
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=True,
        errorTitle="Статус",
        error="Выберите Активен, История или Отключен.",
    )
    status_dv.add(
        f"{get_column_letter(layout.status)}2:{get_column_letter(layout.status)}{last_row}"
    )
    ws.add_data_validation(status_dv)
    _add_blank_required_cf(ws, layout, last_row)

    table_ref = f"A1:{get_column_letter(layout.column_count)}{last_row}"
    table = Table(displayName=REGISTRY_TABLE_NAME, ref=table_ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)
    for col, header in enumerate(layout.headers, start=1):
        _apply_header_cell(ws.cell(row=1, column=col), header)
    _paint_flat_conflicts(ws, layout, sheet_rows, link_book)
    _add_legend_sheet(wb)
    _add_source_info_sheet(wb, rows)

    buffer = BytesIO()
    try:
        wb.save(buffer)
    finally:
        wb.close()
    tmp_path = target.with_name(f".{target.stem}.{os.getpid()}.tmp.xlsx")
    try:
        tmp_path.write_bytes(buffer.getvalue())
        return _publish_workbook(tmp_path, target)
    except OSError:
        tmp_path.unlink(missing_ok=True)
        raise




def _parse_original_row(value: object) -> int | None:
    text = cell_text(value)
    if not text:
        return None
    try:
        return int(text)
    except ValueError:
        return None


def _relation_from_layout(
    values: Sequence[object],
    block: _RelationBlockLayout,
) -> DsRegistryRelation:
    def at(col: int) -> str:
        return cell_text(_cell_at(values, col - 1))

    return DsRegistryRelation(
        group_id=at(block.group_id),
        rfp_key=parse_registry_ds_number(at(block.rfp_key)),
        ul_folder=at(block.ul_folder),
        mode=at(block.mode),
        title_filter=at(block.title_filter),
        mark_filter=at(block.mark_filter),
        block_index=block.index,
        rfp_file=at(block.rfp_file) if block.rfp_file else "",
    )


def _read_source_info(
    workbook: Workbook,
) -> list[tuple[str, str, str, str, int | None]]:
    if SOURCE_INFO_SHEET_NAME not in workbook.sheetnames:
        return []
    sheet = workbook[SOURCE_INFO_SHEET_NAME]
    found: list[tuple[str, str, str, str, int | None]] = []
    for values in sheet.iter_rows(min_row=2, values_only=True):
        values = tuple(values or ())
        source_id = parse_registry_ds_number(_cell_at(values, 0))
        previous = cell_text(_cell_at(values, 1))
        revision = cell_text(_cell_at(values, 2))
        note = cell_text(_cell_at(values, 3))
        original = _parse_original_row(_cell_at(values, 4))
        if not any((source_id, previous, revision, note, original)):
            continue
        found.append((source_id, previous, revision, note, original))
    return found


def _merge_source_info(
    rows: list[DsRegistryRow],
    info: Sequence[tuple[str, str, str, str, int | None]],
) -> list[DsRegistryRow]:
    """Fill previous/revision/note/origin from «Исходный ДС» when the main sheet has none."""

    if not info:
        return rows
    used: set[int] = set()
    merged: list[DsRegistryRow] = []
    for row in rows:
        match_at = None
        for index, item in enumerate(info):
            if index in used:
                continue
            if item[0] == row.source_id:
                match_at = index
                break
        if match_at is None:
            merged.append(row)
            continue
        used.add(match_at)
        _source_id, previous, revision, note, original = info[match_at]
        merged.append(
            DsRegistryRow(
                status=row.status,
                source_id=row.source_id,
                previous_ds=row.previous_ds or previous,
                ds_file=row.ds_file,
                revision=row.revision or revision,
                note=row.note or note,
                relations=row.relations,
                excel_row=row.excel_row,
                original_excel_row=row.original_excel_row or original,
            )
        )
    return merged


def load_registry(
    path: str | Path,
    *,
    ul_root: str | Path | None = None,
    rfp_root: str | Path | None = None,
    ds_root: str | Path | None = None,
    on_lap: Callable[[str, float], None] | None = None,
) -> DsRegistryDocument:
    """Strict-read the canonical ``Реестр ДС`` workbook.

    Several sheet rows with the same actual number collapse into one logical
    DS. An old wide sheet is read the same way. The reconcile column is ignored.

    Args:
        path: Canonical (new-format) xlsx.
        ul_root: Optional TSD root. Folders not named on any row become orphans.
            A named folder that is missing on disk is not an error.
        rfp_root: Optional RFP root. Files not named on any row become orphans.
        ds_root: Optional DS folder. When it is a directory, previous yellow
            rows with no number are dropped and «Файл ДС» is rewritten from
            the workbooks that resolve to each number. Names only: the books
            are not opened and not hashed. A workbook still absent from every
            row becomes a new yellow row.
        on_lap: Optional ``(label, seconds)`` callback for step timings.

    Returns:
        Document with logical rows (including invalid ones) and a validation result.

    Raises:
        DsRegistryFormatError: File/sheet/headers are not the new contract.
    """

    source = Path(path)
    if not source.is_file():
        raise DsRegistryFormatError(f"файл реестра не найден: {source}")
    started = time.perf_counter()
    workbook = _open_workbook(source, data_only=False)
    try:
        if REGISTRY_SHEET_NAME not in workbook.sheetnames:
            raise DsRegistryFormatError(
                f"нет листа «{REGISTRY_SHEET_NAME}»: {source.name}"
            )
        sheet = workbook[REGISTRY_SHEET_NAME]
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header_row:
            raise DsRegistryFormatError("пустой лист реестра")
        layout = _discover_layout(header_row)
        rows: list[DsRegistryRow] = []
        for excel_row, values in enumerate(
            sheet.iter_rows(min_row=2, values_only=True), start=2
        ):
            values = tuple(values or ())
            status = cell_text(_cell_at(values, layout.status - 1))
            source_id = parse_registry_ds_number(
                _cell_at(values, layout.source_id - 1)
            )
            previous = (
                cell_text(_cell_at(values, layout.previous - 1))
                if layout.previous
                else ""
            )
            revision = (
                cell_text(_cell_at(values, layout.revision - 1))
                if layout.revision
                else ""
            )
            note = cell_text(_cell_at(values, layout.note - 1)) if layout.note else ""
            ds_file = (
                cell_text(_cell_at(values, layout.file_ds - 1))
                if layout.file_ds
                else ""
            )
            original = (
                _parse_original_row(_cell_at(values, layout.original_row - 1))
                if layout.original_row
                else None
            )
            if layout.is_flat:
                rel = DsRegistryRelation(
                    group_id=canonical_supply_group_id(source_id) if source_id else "",
                    rfp_key=parse_registry_ds_number(
                        _cell_at(values, layout.flat_rfp_number - 1)
                    )
                    if layout.flat_rfp_number
                    else "",
                    ul_folder=cell_text(_cell_at(values, layout.flat_ul - 1))
                    if layout.flat_ul
                    else "",
                    mode="",
                    rfp_file=cell_text(_cell_at(values, layout.flat_rfp_file - 1))
                    if layout.flat_rfp_file
                    else "",
                    block_index=1,
                    excel_row=excel_row,
                )
                relations = (rel,)
            else:
                relations = tuple(
                    replace(
                        _relation_from_layout(values, block),
                        excel_row=excel_row,
                    )
                    for block in layout.blocks
                )
            candidate = DsRegistryRow(
                status=status,
                source_id=source_id,
                previous_ds=previous,
                ds_file=ds_file,
                revision=revision,
                note=note,
                relations=relations,
                excel_row=excel_row,
                original_excel_row=original,
            )
            if not _row_nonempty_for_status(candidate):
                continue
            trimmed = list(relations)
            while trimmed and trimmed[-1].is_empty():
                trimmed.pop()
            rows.append(
                DsRegistryRow(
                    status=candidate.status,
                    source_id=candidate.source_id,
                    previous_ds=candidate.previous_ds,
                    ds_file=candidate.ds_file,
                    revision=candidate.revision,
                    note=candidate.note,
                    relations=tuple(trimmed),
                    excel_row=excel_row,
                    original_excel_row=original,
                )
            )
        rows = _merge_source_info(rows, _read_source_info(workbook))
    finally:
        workbook.close()
    _lap(on_lap, "чтение листа", started)

    catalog_ready = _ds_catalog_ready(ds_root)
    catalog = None
    if catalog_ready:
        assert ds_root is not None
        started = time.perf_counter()
        catalog = _workbooks_by_source(rows, ds_root)
        _lap(on_lap, "имена файлов ДС", started)
    rows, collapse_issues = collapse_registry_rows(
        rows, compare_ds_file=not catalog_ready
    )
    rows = drop_robot_placeholder_rows(rows)
    rows = sync_ds_files_from_folder(
        rows,
        ds_root if catalog_ready else None,
        catalog=catalog,
    )
    rows.extend(
        collect_orphan_rows(
            rows,
            ul_root=ul_root,
            rfp_root=rfp_root,
            ds_root=ds_root,
            ds_catalog=catalog,
            on_lap=on_lap,
        )
    )
    started = time.perf_counter()
    validation = validate_registry_rows(
        rows, ul_root=ul_root, rfp_root=rfp_root, max_blocks=layout.max_blocks
    )
    validation.issues = collapse_issues + validation.issues
    _lap(on_lap, "проверка строк", started)
    return DsRegistryDocument(
        path=source,
        rows=rows,
        validation=validation,
        max_relation_blocks=layout.max_blocks or MIN_RELATION_BLOCKS,
    )


# ---------------------------------------------------------------------------
# Migration report + migrate-to-path
# ---------------------------------------------------------------------------


def default_migration_report_path(output_path: str | Path) -> Path:
    """Russian-named report next to the migrated workbook."""

    stamp = datetime.now().strftime(BACKUP_SUFFIX_FORMAT)
    output = Path(output_path)
    return output.with_name(f"{MIGRATION_REPORT_PREFIX}_{stamp}.xlsx")


def _write_migration_report(
    report_path: Path,
    *,
    source_path: Path,
    output_path: Path,
    legacy_rows: Sequence[LegacyRegistryRow],
    new_rows: Sequence[DsRegistryRow],
    issues: Sequence[DsRegistryIssue],
    validation: DsRegistryValidation,
) -> Path:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    summary = wb.active
    assert summary is not None
    summary.title = "Сводка"
    summary["A1"] = "Поле"
    summary["B1"] = "Значение"
    summary["A1"].font = Font(bold=True)
    summary["B1"].font = Font(bold=True)
    lines = [
        ("Исходный файл", str(source_path)),
        ("Новый файл", str(output_path)),
        ("Строк legacy", len(legacy_rows)),
        ("Строк новых", len(new_rows)),
        ("Активных", sum(1 for row in new_rows if row.is_active)),
        ("История", sum(1 for row in new_rows if row.is_history)),
        ("ERROR", validation.error_count),
        ("OVERLAY", sum(1 for item in validation.issues if item.blocks_overlay)),
        ("Сводка", validation.summary_line()),
    ]
    for index, (label, value) in enumerate(lines, start=2):
        summary.cell(row=index, column=1, value=label)
        summary.cell(row=index, column=2, value=value)
    summary.column_dimensions["A"].width = 22
    summary.column_dimensions["B"].width = 80
    summary.freeze_panes = "A2"

    mapping = wb.create_sheet("Соответствие строк")
    map_headers = [
        "Старая строка",
        "Новая строка",
        "Статус",
        "ID ДС источника",
        "Актуальный ДС (старый)",
        "Старый ДС",
        "УЛ (старый)",
        "Примечание",
        "ID групп",
        "Ключи RFP",
        "Режимы",
        "Папки УЛ",
        "Коды проблем",
    ]
    for col, title in enumerate(map_headers, start=1):
        cell = mapping.cell(row=1, column=col, value=title)
        cell.font = Font(bold=True)
        cell.fill = _FILL_AUX
    new_by_original = {row.original_excel_row: row for row in new_rows}
    issues_by_row: dict[int, list[str]] = {}
    for item in issues:
        if item.excel_row:
            issues_by_row.setdefault(item.excel_row, []).append(item.code)
    for item in validation.issues:
        if item.excel_row:
            issues_by_row.setdefault(item.excel_row, []).append(item.code)

    for index, legacy in enumerate(legacy_rows, start=2):
        new_row = new_by_original.get(legacy.excel_row)
        mapping.cell(row=index, column=1, value=legacy.excel_row)
        mapping.cell(row=index, column=2, value=None if new_row is None else new_row.excel_row)
        mapping.cell(row=index, column=3, value="" if new_row is None else new_row.status)
        mapping.cell(
            row=index, column=4, value="" if new_row is None else new_row.source_id
        )
        mapping.cell(row=index, column=5, value=legacy.current_ds)
        mapping.cell(row=index, column=6, value=legacy.old_ds)
        mapping.cell(row=index, column=7, value=legacy.ul_text)
        mapping.cell(row=index, column=8, value=legacy.note)
        if new_row is not None:
            mapping.cell(
                row=index,
                column=9,
                value="; ".join(rel.group_id for rel in new_row.relations),
            )
            mapping.cell(
                row=index,
                column=10,
                value="; ".join(rel.rfp_key for rel in new_row.relations),
            )
            mapping.cell(
                row=index,
                column=11,
                value="; ".join(rel.mode for rel in new_row.relations),
            )
            mapping.cell(
                row=index,
                column=12,
                value="; ".join(rel.ul_folder for rel in new_row.relations if rel.ul_folder),
            )
        codes = issues_by_row.get(legacy.excel_row, [])
        if new_row is not None:
            codes.extend(issues_by_row.get(new_row.excel_row, []))
        mapping.cell(row=index, column=13, value="; ".join(dict.fromkeys(codes)))
    mapping.auto_filter.ref = f"A1:{get_column_letter(len(map_headers))}{max(1, len(legacy_rows) + 1)}"
    mapping.freeze_panes = "A2"
    for col in range(1, len(map_headers) + 1):
        mapping.column_dimensions[get_column_letter(col)].width = 22

    problems = wb.create_sheet("Проблемы")
    problem_headers = [
        "Уровень",
        "Код",
        "Строка Excel",
        "ID ДС",
        "Группа",
        "Блок",
        "Overlay",
        "Сообщение",
    ]
    for col, title in enumerate(problem_headers, start=1):
        cell = problems.cell(row=1, column=col, value=title)
        cell.font = Font(bold=True)
        cell.fill = _FILL_AUX
    all_issues = list(issues) + [
        item
        for item in validation.issues
        if not any(
            item.code == other.code
            and item.excel_row == other.excel_row
            and item.block_index == other.block_index
            for other in issues
        )
    ]
    for index, item in enumerate(all_issues, start=2):
        problems.cell(row=index, column=1, value=item.level)
        problems.cell(row=index, column=2, value=item.code)
        problems.cell(row=index, column=3, value=item.excel_row)
        problems.cell(row=index, column=4, value=item.source_id)
        problems.cell(row=index, column=5, value=item.group_id)
        problems.cell(row=index, column=6, value=item.block_index)
        problems.cell(row=index, column=7, value="да" if item.blocks_overlay else "")
        problems.cell(row=index, column=8, value=item.message)
    problems.auto_filter.ref = (
        f"A1:{get_column_letter(len(problem_headers))}{max(1, len(all_issues) + 1)}"
    )
    problems.freeze_panes = "A2"
    problems.column_dimensions["H"].width = 70

    buffer = BytesIO()
    try:
        wb.save(buffer)
    finally:
        wb.close()
    tmp_path = report_path.with_name(f".{report_path.stem}.{os.getpid()}.tmp.xlsx")
    tmp_path.write_bytes(buffer.getvalue())
    os.replace(tmp_path, report_path)
    return report_path


def migrate_registry(
    source_path: str | Path,
    output_path: str | Path,
    *,
    report_path: str | Path | None = None,
) -> DsRegistryMigrationResult:
    """Migrate a legacy 4-column registry to an explicit new-format path.

    Does not touch ``source_path``. Refuses if ``output_path`` is the same file.

    Args:
        source_path: Legacy ``Реестр_ДС_УЛ.xlsx`` (or a temp copy).
        output_path: New-format workbook to write.
        report_path: Optional mapping report. Default is a timestamped
            Russian-named xlsx next to ``output_path``.

    Returns:
        Paths, converted rows and validation.

    Raises:
        DsRegistryError: Same-path output or unreadable legacy file.
    """

    source = Path(source_path)
    output = Path(output_path)
    if source.resolve() == output.resolve():
        raise DsRegistryError(
            "миграция требует отдельный выходной путь, исходник не перезаписывается"
        )
    legacy_rows = read_legacy_registry(source)
    new_rows, migrate_issues = migrate_legacy_rows(legacy_rows)
    written = write_registry_workbook(output, new_rows)
    loaded = load_registry(written)
    # Re-number excel_row to the written file; keep original_excel_row.
    validation = loaded.validation
    report = Path(report_path) if report_path else default_migration_report_path(written)
    _write_migration_report(
        report,
        source_path=source,
        output_path=written,
        legacy_rows=legacy_rows,
        new_rows=loaded.rows,
        issues=migrate_issues,
        validation=validation,
    )
    merged_issues = list(migrate_issues) + [
        item
        for item in validation.issues
        if item.code not in {issue.code for issue in migrate_issues}
        or item.excel_row not in {issue.excel_row for issue in migrate_issues}
    ]
    return DsRegistryMigrationResult(
        source_path=source,
        output_path=written,
        report_path=report,
        rows=loaded.rows,
        validation=validation,
        legacy_rows=legacy_rows,
        issues=merged_issues,
    )


# ---------------------------------------------------------------------------
# Timestamped backup + atomic replace (confirm required)
# ---------------------------------------------------------------------------


def timestamped_backup_path(target: str | Path) -> Path:
    """Return a sibling backup path with ``_backup_YYYYMMDD_HHMMSS``."""

    path = Path(target)
    stamp = datetime.now().strftime(BACKUP_SUFFIX_FORMAT)
    candidate = path.with_name(f"{path.stem}_backup_{stamp}{path.suffix}")
    if not candidate.exists():
        return candidate
    stamp = datetime.now().strftime(BACKUP_SUFFIX_FORMAT + "_%f")
    return path.with_name(f"{path.stem}_backup_{stamp}{path.suffix}")


def backup_and_replace_registry(
    target_path: str | Path,
    new_path: str | Path,
    *,
    confirm: bool = False,
) -> DsRegistryReplaceResult:
    """Copy ``target`` aside, then atomically replace it with ``new_path``.

    Requires explicit ``confirm=True``. Does not invent a production UNC call:
    the caller must pass the path. If Excel has the file open, the original is
    left intact and a clear error is raised.

    Args:
        target_path: Canonical file to replace.
        new_path: Already written and preferably validated new workbook.
        confirm: Must be True; False always refuses without writing.

    Returns:
        Target and backup paths.

    Raises:
        DsRegistryReplaceError: ``confirm`` is false, paths missing, new file
            is not the canonical format, or replace/open is denied.
    """

    if not confirm:
        raise DsRegistryReplaceError(
            "замена канонического реестра требует явный confirm=True"
        )
    target = Path(target_path)
    incoming = Path(new_path)
    if not target.is_file():
        raise DsRegistryReplaceError(f"целевой файл не найден: {target}")
    if not incoming.is_file():
        raise DsRegistryReplaceError(f"новый файл не найден: {incoming}")
    if target.resolve() == incoming.resolve():
        raise DsRegistryReplaceError("новый файл совпадает с целевым")
    loaded = load_registry(incoming)
    if not loaded.validation.is_ok:
        raise DsRegistryReplaceError(
            "новый реестр не проходит валидацию: " + loaded.validation.summary_line()
        )
    backup = timestamped_backup_path(target)
    try:
        shutil.copy2(target, backup)
    except OSError as exc:
        raise DsRegistryReplaceError(
            f"не удалось создать резервную копию «{backup.name}»: {exc}"
        ) from exc

    tmp_path = target.with_name(f".{target.stem}.replace_{os.getpid()}.tmp{target.suffix}")
    try:
        shutil.copy2(incoming, tmp_path)
        os.replace(tmp_path, target)
    except OSError as exc:
        try:
            tmp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise DsRegistryReplaceError(
            f"не удалось заменить «{target.name}»: файл занят или запись запрещена. "
            f"Исходник не изменён. Резервная копия: {backup.name}. "
            "Закройте Excel и повторите."
        ) from exc
    return DsRegistryReplaceResult(
        target_path=target,
        backup_path=backup,
        replaced=True,
    )
