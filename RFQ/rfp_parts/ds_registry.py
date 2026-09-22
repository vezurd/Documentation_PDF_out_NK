"""Canonical DS registry: one sheet, one row per source DS, numbered links.

The one working workbook lives in ``_RFP`` as ``Реестр_ДС_УЛ.xlsx``.
«Проверить реестр» renames a legacy file to ``Реестр_ДС_УЛ_old.xlsx`` and
replaces it. If Excel holds the file open, a dated copy is written; the
program reads the newest of the main file and at most five dated copies.
"""

from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Iterator, Literal, Sequence

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
HDR_SOURCE_ID = f"Номер ДС {HEADER_REQUIRED_MARK}"
HDR_SOURCE_ID_LEGACY = "ID ДС источника"
HDR_FILE_DS = "Файл ДС"
HDR_PREVIOUS = "Предыдущий / старый ДС"
HDR_REVISION = "Ревизия"
HDR_NOTE = "Примечание"
HDR_ORIGINAL_ROW = "Исходная строка реестра"
SOURCE_INFO_SHEET_NAME = "Исходный ДС"
RFP_STATUS_MATCH = "совпало"
RFP_STATUS_MISS = "не совпало"

CORE_HEADER_TITLES: tuple[str, ...] = (
    HDR_STATUS,
    HDR_SOURCE_ID,
    HDR_FILE_DS,
)
_SOURCE_ID_HEADER_NAMES = ("Номер ДС", "ID ДС источника")

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
_FILL_TODO = _FILL_BLANK
_FONT_DATA = Font(name="Calibri", size=11)
_ALIGN_HEADER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_ALIGN_DATA = Alignment(horizontal="left", vertical="top", wrap_text=True)
_THIN = Side(style="thin", color="B0B0B0")
_BLOCK_BORDER_COLORS = ("1B4F72", "548235", "C65911", "7030A0")
_COMMENT_AUTHOR = "реестр ДС"

_CORE_WIDTHS = (16, 18, 42)
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


def _open_workbook(path: Path, *, data_only: bool) -> Workbook:
    """Load xlsx from bytes so Windows does not keep the destination locked."""

    return load_workbook(BytesIO(Path(path).read_bytes()), data_only=data_only)


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
    mode: str
    title_filter: str = ""
    mark_filter: str = ""
    block_index: int = 1

    def is_empty(self) -> bool:
        return not any(
            (
                self.group_id,
                self.rfp_key,
                self.ul_folder,
                self.mode,
                self.title_filter,
                self.mark_filter,
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
        file_ds=3,
        previous=0,
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
            ("Номер ДС", bool(source_header)),
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
        "Номер ДС",
        "ID ДС источника",
        "Файл ДС",
        "Предыдущий / старый ДС",
        "Ревизия",
        "Примечание",
        "Исходная строка реестра",
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
        previous=indexed.get("Предыдущий / старый ДС", 0),
        revision=indexed.get("Ревизия", 0),
        note=indexed.get("Примечание", 0),
        original_row=indexed.get("Исходная строка реестра", 0),
        blocks=tuple(blocks),
    )


def _header_comment(header: str) -> str:
    comments = {
        HDR_STATUS: (
            "Обязательно (*). Допустимо: Активен, История, Отключен. "
            "Исторические строки не участвуют в запуске. "
            "Маркер обязательности — символ * и цвет шапки."
        ),
        HDR_SOURCE_ID: (
            "Обязательно для статуса Активен (*). Пишите ДС11, ДС47, ДС4905_1. "
            "У активных строк номер не повторяется."
        ),
        HDR_FILE_DS: (
            "Имя файла ДС. Путь в ячейке не пишется: щелчок открывает файл. "
            "Робот подставляет его при проверке реестра."
        ),
        HDR_PREVIOUS: (
            "Вспомогательное поле. Текст «Старый ДС» как в исходнике, без разбора."
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
) -> DsRegistryIssue:
    return DsRegistryIssue(
        code=code,
        level=level,
        message=message,
        excel_row=None if row is None else row.excel_row,
        source_id="" if row is None else row.source_id,
        group_id=group_id,
        block_index=block_index,
        blocks_overlay=blocks_overlay,
    )


def _row_nonempty_for_status(row: DsRegistryRow) -> bool:
    if row.status or row.source_id or row.previous_ds or row.revision or row.note:
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


def index_ul_folders(ul_root: str | Path | None) -> dict[str, tuple[str, ...]]:
    """Map a UL actual number to first-level folder names under ``ul_root``."""

    if not ul_root:
        return {}
    root = Path(ul_root)
    found: dict[str, list[str]] = {}
    try:
        children = [path for path in root.iterdir() if path.is_dir()]
    except OSError:
        return {}
    for folder in children:
        identity = parse_ul_folder_ds_identity(folder.name)
        if identity.actual is None:
            continue
        found.setdefault(str(identity.actual), []).append(folder.name)
    return {key: tuple(names) for key, names in found.items()}


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


def validate_registry_rows(
    rows: Sequence[DsRegistryRow],
    *,
    ul_root: str | Path | None = None,
    max_blocks: int | None = None,
) -> DsRegistryValidation:
    """Validate parsed rows. History without ID/blocks is allowed and kept."""

    issues: list[DsRegistryIssue] = []
    active_ids: dict[str, DsRegistryRow] = {}
    group_anchor: dict[str, tuple[str, str, DsRegistryRow, int]] = {}
    rfp_anchor: dict[str, tuple[str, DsRegistryRow]] = {}
    ul_root_path = Path(ul_root) if ul_root else None
    ul_index = index_ul_folders(ul_root_path) if ul_root_path is not None else {}

    for row in rows:
        if not _row_nonempty_for_status(row):
            continue
        if row.status not in ALLOWED_STATUSES:
            issues.append(
                _issue(
                    ISSUE_INVALID_STATUS,
                    "ERROR",
                    f"недопустимый статус {row.status!r}; "
                    f"ожидается {', '.join(sorted(ALLOWED_STATUSES))}",
                    row,
                )
            )

        started = [rel for rel in row.relations if not rel.is_empty()]
        if row.relations:
            occupancy = [not rel.is_empty() for rel in row.relations]
            if any(occupancy):
                first_empty = None
                for rel in row.relations:
                    if rel.is_empty() and first_empty is None:
                        first_empty = rel.block_index
                    elif not rel.is_empty() and first_empty is not None:
                        issues.append(
                            _issue(
                                ISSUE_GAPPED_BLOCK,
                                "ERROR",
                                (
                                    f"блок связи {rel.block_index} заполнен, "
                                    f"а блок {first_empty} пуст"
                                ),
                                row,
                                block_index=rel.block_index,
                            )
                        )
                        break

        if row.is_history:
            for rel in started:
                if rel.mode and rel.mode not in ALLOWED_MODES:
                    issues.append(
                        _issue(
                            ISSUE_INVALID_MODE,
                            "WARN",
                            f"недопустимый режим {rel.mode!r}",
                            row,
                            group_id=rel.group_id,
                            block_index=rel.block_index,
                        )
                    )
                _append_ul_identity_issues(issues, row, rel)
            continue

        if row.is_active and not row.source_id:
            issues.append(
                _issue(
                    ISSUE_MISSING_SOURCE_ID,
                    "ERROR",
                    "у активной строки нет ID ДС источника",
                    row,
                )
            )
        if row.is_active and row.source_id:
            previous = active_ids.get(row.source_id)
            if previous is not None:
                issues.append(
                    _issue(
                        ISSUE_DUPLICATE_SOURCE_ID,
                        "ERROR",
                        (
                            f"активный ID {row.source_id!r} повторяется "
                            f"(строки {previous.excel_row} и {row.excel_row})"
                        ),
                        row,
                    )
                )
            else:
                active_ids[row.source_id] = row

        if row.is_active and not started:
            issues.append(
                _issue(
                    ISSUE_MISSING_RELATION,
                    "ERROR",
                    "у активной строки нет блока связи 1",
                    row,
                )
            )

        for rel in started:
            if rel.mode and rel.mode not in ALLOWED_MODES:
                issues.append(
                    _issue(
                        ISSUE_INVALID_MODE,
                        "ERROR",
                        f"недопустимый режим {rel.mode!r}",
                        row,
                        group_id=rel.group_id,
                        block_index=rel.block_index,
                    )
                )
            missing_core = not rel.group_id or not rel.rfp_key or not rel.mode
            if missing_core:
                issues.append(
                    _issue(
                        ISSUE_PARTIAL_BLOCK,
                        "ERROR",
                        (
                            f"блок связи {rel.block_index} заполнен частично "
                            "(нужны ID группы, ключ RFP и режим)"
                        ),
                        row,
                        group_id=rel.group_id,
                        block_index=rel.block_index,
                    )
                )
            if rel.mode and rel.mode != MODE_NO_UL and not rel.ul_folder:
                issues.append(
                    _issue(
                        ISSUE_UL_REQUIRED,
                        "ERROR",
                        (
                            f"блок {rel.block_index}: папка УЛ обязательна, "
                            f"кроме режима «{MODE_NO_UL}»"
                        ),
                        row,
                        group_id=rel.group_id,
                        block_index=rel.block_index,
                    )
                )
            if rel.mode == MODE_FILTER and not rel.has_complete_filters():
                issues.append(
                    _issue(
                        ISSUE_FILTERS_REQUIRED,
                        "ERROR",
                        (
                            f"блок {rel.block_index}: при режиме «{MODE_FILTER}» "
                            "нужны фильтр титула и фильтр марки"
                        ),
                        row,
                        group_id=rel.group_id,
                        block_index=rel.block_index,
                    )
                )
            _append_ul_identity_issues(issues, row, rel)
            if rel.mode == MODE_NO_UL and rel.rfp_key and ul_index.get(rel.rfp_key):
                found_names = ", ".join(ul_index[rel.rfp_key])
                issues.append(
                    _issue(
                        ISSUE_UL_UNEXPECTED,
                        "WARN",
                        (
                            f"блок {rel.block_index}: режим «{MODE_NO_UL}», "
                            f"но в папке УЛ уже есть: {found_names}"
                        ),
                        row,
                        group_id=rel.group_id,
                        block_index=rel.block_index,
                    )
                )
            if (
                ul_root_path is not None
                and rel.ul_folder
                and rel.mode != MODE_NO_UL
            ):
                folder_path = ul_root_path / rel.ul_folder
                if not folder_path.is_dir():
                    issues.append(
                        _issue(
                            ISSUE_UL_FOLDER_MISSING,
                            "ERROR",
                            f"нет папки УЛ {rel.ul_folder!r}",
                            row,
                            group_id=rel.group_id,
                            block_index=rel.block_index,
                        )
                    )
            if rel.group_id:
                anchor = group_anchor.get(rel.group_id)
                if anchor is None:
                    group_anchor[rel.group_id] = (
                        rel.rfp_key,
                        rel.ul_folder,
                        row,
                        rel.block_index,
                    )
                else:
                    a_key, a_ul, a_row, _a_block = anchor
                    if a_key != rel.rfp_key or a_ul != rel.ul_folder:
                        issues.append(
                            _issue(
                                ISSUE_GROUP_MISMATCH,
                                "ERROR",
                                (
                                    f"группа {rel.group_id!r} повторена с другими "
                                    f"ключом RFP/папкой УЛ (строка {a_row.excel_row} "
                                    f"vs {row.excel_row})"
                                ),
                                row,
                                group_id=rel.group_id,
                                block_index=rel.block_index,
                            )
                        )
            if rel.rfp_key:
                other = rfp_anchor.get(rel.rfp_key)
                if other is None:
                    rfp_anchor[rel.rfp_key] = (rel.group_id, row)
                elif other[0] and rel.group_id and other[0] != rel.group_id:
                    issues.append(
                        _issue(
                            ISSUE_RFP_KEY_CONFLICT,
                            "ERROR",
                            (
                                f"ключ RFP {rel.rfp_key!r} относится к разным "
                                f"группам {other[0]!r} и {rel.group_id!r}"
                            ),
                            row,
                            group_id=rel.group_id,
                            block_index=rel.block_index,
                        )
                    )

        if len(started) > 1:
            fully_filtered = all(
                rel.mode == MODE_FILTER and rel.has_complete_filters()
                for rel in started
            )
            if not fully_filtered:
                for rel in started:
                    issues.append(
                        _issue(
                            ISSUE_MULTI_LINK_NEEDS_SPLIT,
                            "OVERLAY",
                            (
                                "несколько связей без полных фильтров: overlay "
                                f"группы {rel.group_id or '—'} заблокирован, "
                                f"нужен режим «{MODE_NEEDS_SPLIT}» или «{MODE_FILTER}»"
                            ),
                            row,
                            group_id=rel.group_id,
                            block_index=rel.block_index,
                            blocks_overlay=True,
                        )
                    )
                    if rel.mode == MODE_WHOLE:
                        issues.append(
                            _issue(
                                ISSUE_INVALID_MODE,
                                "ERROR",
                                (
                                    f"блок {rel.block_index}: при нескольких УЛ "
                                    f"режим «{MODE_WHOLE}» недопустим"
                                ),
                                row,
                                group_id=rel.group_id,
                                block_index=rel.block_index,
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
                    "Номер ДС" in names or "ID ДС источника" in names
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
            "Одна строка",
            star,
            "Один ДС — одна строка на листе «Реестр ДС».",
            "Робот собирает свод по этой строке. Вторую поставку того же ДС "
            "пишите в Связь 2, не через точку с запятой.",
        ),
        (
            "Статус",
            star,
            "Активен, История или Отключен.",
            "В свод попадают только «Активен». История и Отключен робот пропускает.",
        ),
        (
            "Номер ДС",
            star,
            "Пишите ДС11, ДС47, ДС4905_1. У активных строк номер не повторяется.",
            "По нему робот находит файл ДС. Чужой номер привяжет не ту спецификацию.",
        ),
        (
            "Файл ДС",
            grey,
            "Не заполняйте вручную. Робот ставит имя файла при проверке реестра.",
            "В ячейке только имя. Щелчок открывает файл.",
        ),
        (
            "Номер RFP, Файл RFP, Статус RFP",
            star,
            "Номер — какой файл искать в корне RFP_Зиновьев. Файл и статус "
            "робот подставляет сам: «совпало» или «не совпало».",
            "Совпало значит, что файл с этим номером есть. "
            "Количества это не сравнивает.",
        ),
        (
            "Лист «Исходный ДС»",
            grey,
            "Там предыдущий ДС, ревизия, примечание и номер старой строки. "
            "На рабочем листе этих столбцов нет.",
            "На свод и на RFP не влияет. Робот этот лист только хранит.",
        ),
        (
            "Связь 1, Связь 2…",
            star,
            "Связь 1 заполните всегда. Связь 2 — только если у этого ДС вторая "
            "поставка (другая папка УЛ или другой RFP). Пустых дыр не оставляйте: "
            "нельзя заполнить Связь 2, оставив Связь 1 пустой.",
            "Каждая связь — отдельная поставка. Робот не смешивает их в одной ячейке.",
        ),
        (
            "ID группы поставки",
            star,
            "Обычно точное имя папки УЛ. Если два ДС делят одну поставку "
            "(например 13 и 47), у обоих одинаковые группа, ключ RFP и папка УЛ.",
            "Сверка с RFP идёт по группе целиком, не по одному ДС. "
            "Совпали код, единица и количество — в запуск попадает RFP. "
            "Нет — вся группа остаётся из ДС.",
        ),
        (
            "Фактический ДС / ключ RFP",
            star,
            "Номер, который робот ищет в имени файла в корне RFP_Зиновьев. Пример: 13. "
            "Один ключ принадлежит только одной группе.",
            "Подпапки RFP робот не читает. Чужой ключ наложит чужой файл или не найдёт RFP.",
        ),
        (
            "Папка УЛ",
            yellow,
            "Точное имя папки в ТСД. Пустой можно оставить только в режиме «Нет УЛ».",
            "Робот проверяет, что папка есть. Неверное имя — ошибка реестра.",
        ),
        (
            "Режим «Вся ДС»",
            star,
            "Все строки этого ДС входят в связь. Фильтры оставьте пустыми. Папку УЛ заполните.",
            "Обычный случай: один ДС — одна поставка.",
        ),
        (
            "Режим «По фильтру»",
            yellow,
            "Заполните и фильтр титула, и фильтр марки. Так один ДС делится на несколько папок УЛ.",
            "В связь попадут только строки с этим титулом и маркой. "
            "Пустой фильтр — ошибка, группу на RFP не заменят.",
        ),
        (
            "Режим «Нет УЛ»",
            star,
            "Папку УЛ оставьте пустой. Робот сам смотрит папку со всеми УЛ: "
            "если каталога с этим номером нет, ячейку трогать не нужно. "
            "Если каталог появился, ячейка режима станет жёлтой.",
            "У поставки нет папки УЛ. Группа получит имя NO_UL: и номер ДС.",
        ),
        (
            "Режим «Требует распределения»",
            yellow,
            "Так робот помечает несколько связей без полных фильтров. "
            "Исправьте: либо одна связь «Вся ДС», либо у каждой связи оба фильтра.",
            "Пока режим такой, эту группу нельзя заменить на RFP. В своде останется ДС.",
        ),
        (
            "Строки только из RFP",
            grey,
            "В реестре их не пишут.",
            "Если позиции нет в ДС, в свод для запуска она не попадёт.",
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


def write_registry_workbook(
    path: str | Path,
    rows: Sequence[DsRegistryRow],
    *,
    spare_rows: int = DEFAULT_SPARE_ROWS,
    links: RegistryLinks | None = None,
) -> Path:
    """Write the registry, the how-to sheet and the source-info sheet.

    Args:
        path: Destination xlsx. Parent directories are created.
        rows: Canonical rows in display order.
        spare_rows: Extra empty table rows so a person can append data.
        links: Optional DS/RFP/UL lookup filled by «Проверить реестр».

    Returns:
        The written path.

    Raises:
        OSError: If the destination cannot be written.
    """

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    max_blocks = max(
        (len(row.relations) for row in rows),
        default=MIN_RELATION_BLOCKS,
    )
    max_blocks = max(MIN_RELATION_BLOCKS, max_blocks)
    layout = _layout_from_max_blocks(max_blocks)
    last_data = max(len(rows), 1) + max(0, int(spare_rows))
    last_row = last_data + 1

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = REGISTRY_SHEET_NAME
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 48

    for col, header in enumerate(layout.headers, start=1):
        _apply_header_cell(ws.cell(row=1, column=col), header)

    for col, width in enumerate(_CORE_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    for block in layout.blocks:
        for offset, width in enumerate(_BLOCK_WIDTHS):
            ws.column_dimensions[get_column_letter(block.columns()[offset])].width = (
                width
            )

    link_book = links or RegistryLinks()
    for index, row in enumerate(rows):
        excel_row = index + 2
        _write_text_cell(ws, excel_row, layout.status, row.status)
        _write_text_cell(
            ws, excel_row, layout.source_id, format_registry_ds_number(row.source_id)
        )
        _write_file_links(
            ws,
            excel_row,
            layout.file_ds,
            link_book.ds_files.get(row.source_id, ()),
        )
        for block in layout.blocks:
            rel = None
            if block.index <= len(row.relations):
                rel = row.relations[block.index - 1]
            _write_text_cell(ws, excel_row, block.group_id, rel.group_id if rel else "")
            _write_text_cell(ws, excel_row, block.rfp_key, rel.rfp_key if rel else "")
            rfp_paths = (
                link_book.rfp_files.get(rel.rfp_key, ())
                if rel is not None and rel.rfp_key
                else ()
            )
            _write_file_links(ws, excel_row, block.rfp_file, rfp_paths)
            status_text = ""
            if rel is not None and rel.rfp_key and link_book.scanned_rfp:
                status_text = RFP_STATUS_MATCH if rfp_paths else RFP_STATUS_MISS
            _write_text_cell(ws, excel_row, block.rfp_status, status_text)
            _write_text_cell(ws, excel_row, block.ul_folder, rel.ul_folder if rel else "")
            _write_text_cell(ws, excel_row, block.mode, rel.mode if rel else "")
            _write_text_cell(
                ws, excel_row, block.title_filter, rel.title_filter if rel else ""
            )
            _write_text_cell(
                ws, excel_row, block.mark_filter, rel.mark_filter if rel else ""
            )
        ws.row_dimensions[excel_row].height = 18

    for excel_row in range(len(rows) + 2, last_row + 1):
        for col in range(1, layout.column_count + 1):
            cell = ws.cell(row=excel_row, column=col, value=None)
            cell.alignment = _ALIGN_DATA
            cell.font = _FONT_DATA
            if col in (
                layout.status,
                layout.source_id,
                *(c for block in layout.blocks for c in (block.group_id, block.rfp_key, block.mode)),
            ):
                cell.number_format = "@"

    _apply_block_borders(ws, layout, last_row)

    status_dv = DataValidation(
        type="list",
        formula1='"' + ",".join((STATUS_ACTIVE, STATUS_HISTORY, STATUS_DISABLED)) + '"',
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=True,
        errorTitle="Статус",
        error="Выберите Активен, История или Отключен.",
    )
    status_dv.add(f"{get_column_letter(layout.status)}2:{get_column_letter(layout.status)}{last_row}")
    ws.add_data_validation(status_dv)

    mode_dv = DataValidation(
        type="list",
        formula1='"' + ",".join((MODE_WHOLE, MODE_FILTER, MODE_NO_UL, MODE_NEEDS_SPLIT)) + '"',
        allow_blank=True,
        showDropDown=False,
        showErrorMessage=True,
        errorTitle="Режим распределения",
        error=(
            f"Выберите {MODE_WHOLE}, {MODE_FILTER}, {MODE_NO_UL} или {MODE_NEEDS_SPLIT}."
        ),
    )
    for block in layout.blocks:
        mode_dv.add(
            f"{get_column_letter(block.mode)}2:{get_column_letter(block.mode)}{last_row}"
        )
    ws.add_data_validation(mode_dv)

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
    # Re-apply header fills/comments after the table so colour is stored on cells
    # (Excel table style is not the only marker: headers keep * / †).
    for col, header in enumerate(layout.headers, start=1):
        _apply_header_cell(ws.cell(row=1, column=col), header)
    _apply_block_borders(ws, layout, last_row)
    _paint_human_todo(ws, layout, rows, link_book)
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


# ---------------------------------------------------------------------------
# Canonical reader
# ---------------------------------------------------------------------------


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
        if row.previous_ds or row.revision or row.note or row.original_excel_row:
            merged.append(row)
            continue
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
        source_id, previous, revision, note, original = info[match_at]
        merged.append(
            DsRegistryRow(
                status=row.status,
                source_id=row.source_id or source_id,
                previous_ds=previous,
                revision=revision,
                note=note,
                relations=row.relations,
                excel_row=row.excel_row,
                original_excel_row=original,
            )
        )
    return merged


def load_registry(
    path: str | Path,
    *,
    ul_root: str | Path | None = None,
) -> DsRegistryDocument:
    """Strict-read the canonical ``Реестр ДС`` workbook.

    Args:
        path: Canonical (new-format) xlsx.
        ul_root: Optional TSD root; when set, missing UL folders are errors.

    Returns:
        Document with rows (including invalid ones) and a validation result.

    Raises:
        DsRegistryFormatError: File/sheet/headers are not the new contract.
    """

    source = Path(path)
    if not source.is_file():
        raise DsRegistryFormatError(f"файл реестра не найден: {source}")
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
            original = (
                _parse_original_row(_cell_at(values, layout.original_row - 1))
                if layout.original_row
                else None
            )
            relations = tuple(
                _relation_from_layout(values, block) for block in layout.blocks
            )
            candidate = DsRegistryRow(
                status=status,
                source_id=source_id,
                previous_ds=previous,
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

    validation = validate_registry_rows(
        rows, ul_root=ul_root, max_blocks=layout.max_blocks
    )
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
