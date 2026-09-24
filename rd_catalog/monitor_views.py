"""Qt-free painted-cell join for the RD catalog GUI and WEB monitor.

Assembles ``MonitorCell`` text/fill/tooltip for Комплекты, heatmap,
worklist, АН, journal, MTO readiness, collisions, and the kit card.
Does not import Qt, does not rebuild the pipeline, and does not scan.
WEB may pass ``sheets=WEB_MONITOR_SHEETS`` to skip heatmap/journal DTO.
"""

from __future__ import annotations

import json
import re
from collections.abc import Collection, Iterable, Mapping, Sequence
from dataclasses import asdict, dataclass, field, is_dataclass, replace
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any

from rd_catalog.an_compare import (
    AN_AGREED_HEADER,
    AN_AGREED_HEADER_TIP,
    AN_AGREED_ROW_HEADERS,
    AN_HEADERS,
    an_agreed_cell_fill,
    an_agreed_cell_tooltip,
    an_agreed_row_fill,
    an_vs_cell_fill,
    an_vs_cell_tooltip,
    format_an_agreed_cell,
    format_an_vs_cell,
)
from rd_catalog.an_compare_cache import (
    cache_key as an_content_cache_key,
    counterpart_mtime_ns,
    load_an_content_compare_cache,
    result_from_entry as an_result_from_entry,
)
from rd_catalog.an_index import (
    AnAgreedScore,
    AnMtoFile,
    KIND_HEADER,
    KIND_OD,
    KitAnTargets,
    agreed_revision_target,
    an_cell_text,
    an_file_kind,
    an_is_od,
    match_an_to_kit,
    score_an_files_for_agreed,
)
from rd_catalog.auto_mto_compare_cache import (
    load_auto_mto_compare_cache,
    result_from_entry as auto_mto_result_from_entry,
)
from rd_catalog.ban_filter import BanFilterStore
from rd_catalog.config import CatalogConfig
from rd_catalog.customer_pi_auto_mto import (
    AUTO_MTO_COMPARE_STATUS_HEADER,
    AUTO_MTO_EXACT_STATUS,
    AUTO_MTO_SOFT_STATUS,
    AutoMtoCompareResult,
    AutoMtoCompareStatus,
    AutoMtoFile,
    auto_mto_compare_status,
    auto_mto_path,
    auto_mto_rd_paths_match,
    format_auto_mto_cell_text,
    list_auto_mto_files_by_kit,
    pick_auto_mto_file,
    revision_text_from_mto_filename,
    revision_texts_match,
)
from rd_catalog.db import (
    CatalogDatabase,
    KitPackageRow,
    KitPipelineRow,
    KitRevisionRow,
    RobotMtoAcceptRow,
)
from rd_catalog.google_kits import load_cached_google_kits
from rd_catalog.google_sheet_links import (
    GOOGLE_HREF_TIP,
    SheetLinkContext,
    journal_cell_href,
    kits_google_hrefs,
    sheet_link_context_from_config,
)
from rd_catalog.issuance_review import (
    IssuanceJournalRow,
    JournalAutoMtoHit,
    journal_row_matches_issuance,
    latest_effective_issuance_kits,
    list_issuance_journal,
)
from rd_catalog.kits import (
    F_LINE_MTO_ABSENT,
    GoogleKit,
    IssuanceKit,
    KitEvent,
    KitMatrixRow,
    KitSummary,
    aggregate_source_kits,
    build_kit_matrix,
    format_revision,
    kit_identity_key,
    kit_revision_match_flags,
    kit_robot_origin,
    last_event_parts,
    last_event_text,
    mto_content_equal_by_kit,
    parse_sheet_revision,
    summary_label,
    summary_tooltip,
    annulled_folders_from_pipelines,
    working_folders_from_pipelines,
)
from rd_catalog.models import FileKind, FileRecord, SourceKind, collision_kind_label
from rd_catalog.mto_export import (
    PIN_COLUMN_HEADER,
    ExportPin,
    ExportPinView,
    ExportSelection,
    export_pin_view,
    load_export_pins,
)
from rd_catalog.mto_export_copy import (
    EXPORT_STATE_COLOR_KEY,
    EXPORT_STATE_TEXT,
    export_selection_tooltip,
)
from rd_catalog.overlay import revision_rank
from rd_catalog.parse import record_has_canonical_layout
from rd_catalog.path_actions import path_is_under
from rd_catalog.perf_log import perf_span
from rd_catalog.robot_mto_accept import (
    ACCEPT_LIVE,
    ROBOT_MTO_ACCEPT_FOREGROUND,
    accepts_by_kit,
    mto_xlsx_record,
    robot_mto_accept_paints_blue,
    robot_mto_accept_paints_green,
    robot_mto_accept_state,
    robot_mto_accept_tooltip,
)
from rd_catalog.transfer_review_compare import path_pair_labels_from_cache
from rd_catalog.pipeline import (
    KitCard,
    MtoWorklistRow,
    get_kit_card,
    list_kit_pipelines,
    list_mto_worklist,
    list_revision_columns,
    list_revision_matrix,
    official_detected_current_ids,
    parse_ddmmyyyy,
    pipeline_approval_color_key,
    pipeline_approval_label,
    pipeline_approval_relation,
    pipeline_approval_shows_letter,
    pipeline_display_code_a,
    pipeline_display_review_status,
    pipeline_review_display,
    pipeline_review_label,
    revision_texts_equivalent,
    working_revision_for_display,
    APPROVAL_REL_AHEAD,
    APPROVAL_REL_NO_REV,
    APPROVAL_REL_OTHER,
    APPROVAL_REL_PREVIOUS,
    PIPELINE_DISPLAY_V2,
    PIPELINE_DISPLAY_V3,
    PIPELINE_DISPLAY_VERSION,
    PIPELINE_F_AUTO_SUFFIX,
    PIPELINE_FACE_DE,
    PIPELINE_FACE_F,
    PIPELINE_FACE_ISSUANCE,
    PIPELINE_FACE_RD,
)
from rd_catalog.skip_dirs import load_runtime_skip_dirs, path_has_skipped_dir
from rd_catalog.status_colors import (
    color_for,
    default_palette,
    load_status_colors,
    status_color_label,
    status_colors_path,
    status_short_label,
)

REV_MATCH_FILL = "#E2F2E1"
REV_DIFF_FILL = "#F7E8BE"
AUTO_MTO_AHEAD_FILL = "#F0D0EA"
ROBOT_ORIGIN_FILL = "#C5E8C4"
ROBOT_ORPHAN_FILL = "#F3C2C2"
TDO_STATUSES = frozenset({"tdo_review", "code_a", "code_b", "code_c"})
KITS_TDO_STATUSES = frozenset({"tdo_review", "agreed"})
_RD_TDO_PASSED_STAGES = frozenset({"tdo_passed", "incoming_passed"})
RD_AB_SUFFIX = " AB"
RD_MISSING_TRANSFER_NOTE = "Нет в РД"
RD_MISSING_TRANSFER_KEY = "rd_missing_transfer"
MIXED_TITLES_KEY = "mixed_titles"
EMPTY_PIPELINE_WARNING = "Пустой kit_pipeline — обновите каталог в Qt."
_NOT_COMPARED_ERROR = "Сверка RD ↔ robot ещё не выполнена"
_GREY_FOREGROUND = "#80868b"
_DEFAULT_FOREGROUND = "#202124"

KITS_OK_HEADER = "Ок"
KITS_OK_REV_KEYS: tuple[str, ...] = (
    "issuance",
    "google_f",
    "robot",
    "rd_mto",
)
KITS_OK_REV_REASON = {
    "issuance": "Выдача · рев. расходится с РД",
    "google_f": "Google · рев. F расходится с РД",
    "robot": "Робот МТО · рев. расходится с MTO РД",
    "rd_mto": "MTO · рев. расходится с OD",
}
KITS_OK_TOOLTIP = (
    "Комплексный признак «хорошего» комплекта: сводка «Совпадает», "
    "совпали Выдача / Google F / робот / MTO с официальной РД, "
    "диск не отстаёт от письма A, рассмотрение «Согласован», "
    "текущий код A.\n"
    "Жёлтые Авто МТО / сверка, пустые SQ и рабочая, "
    "проблемы старых ревизий MTO не мешают.\n"
    "Сортировка: «да» сверху."
)

KITS_HEADERS = (
    "Титул",
    "Марка",
    KITS_OK_HEADER,
    PIN_COLUMN_HEADER,
    "Выдача · рев.",
    "Google · рев.",
    "Робот МТО · рев.",
    "SQ · рев.",
    "РД · рев.",
    "Рабочая рев. РД",
    "MTO · рев.",
    "Авто МТО",
    AUTO_MTO_COMPARE_STATUS_HEADER,
    "АН МТО",
    "Сводка",
    "Статус рассмотрения",
    "Статус согласования",
    "Google · статус",
    "Google · дата F",
    "Google · этап F",
    "Google · рев. F",
    "Google · TRM F",
    "Выдача · статус",
    "Выдача · дата отпр.",
    "Выдача · TRM отпр.",
    "Выдача · дата вх.контр.",
    "Выдача · TRM подтв.",
    "РД · дата файла",
)

KITS_TOOLTIP_PRIORITY_HEADERS: tuple[str, ...] = (
    "Сводка",
    "РД · рев.",
)
KITS_F_SHEET_LABEL = "КСБ ИД (Контроль выдачи), столбец F"
KITS_DE_SHEET_LABEL = "КСБ ИД (Контроль выдачи), столбцы D/E"
ISSUANCE_SHEET_LABEL = "Выдача РД ПД"
KITS_TIPS_CTRL_CLICK_HINT = (
    "Ctrl+клик по ячейке таблицы комплектов прокручивает "
    "блок этого столбца к верху панели."
)
KITS_TIPS_WRAP_WIDTH = 100
# UNC anywhere; drive / ``//`` only at the start of a token so ``https://``
# is not treated as a filesystem path.
_FS_PATH_START_RE = re.compile(
    r"(?:\\\\|(?:(?<=\s)|^)(?:[A-Za-z]:[\\/]|//))"
)
_TIPS_TERMINATOR_RE = re.compile(r"[.!?]+[»\"”)\]]*")
_TIPS_ABBREVS = frozenset(
    {
        "рев",
        "согл",
        "др",
        "стр",
        "рис",
        "см",
        "гг",
        "т.д",
        "т.п",
        "т.е",
        "н.п",
    }
)
_REVIEW_PASS_STAGES = frozenset({"tdo_passed", "incoming_passed"})
_REVIEW_SEND_STAGES = frozenset({"tdo_sent", "incoming_sent"})
_REVIEW_AGREED_STAGES = frozenset({"code_a", "agreed"})
_APPROVAL_CODE_STAGES = frozenset({"code_a", "code_b", "code_c"})

KITS_REV_SOURCE_COLUMNS: dict[str, str] = {
    "Google · рев.": "google",
    "РД · рев.": "rd",
    "MTO · рев.": "rd_mto",
    "SQ · рев.": "sq",
    "Робот МТО · рев.": "robot",
    "Google · рев. F": "google_f",
    "Выдача · рев.": "issuance",
}

SUMMARY_FOREGROUND: dict[KitSummary, str] = {
    KitSummary.ALIGNED: "#137333",
    KitSummary.REV_MISMATCH: "#9a6700",
    KitSummary.TRANSFER_REVIEW: "#9a6700",
    KitSummary.MIXED_TITLES: "#b3261e",
    KitSummary.GAP_RD: "#b3261e",
    KitSummary.GOOGLE_ONLY: "#b3261e",
    KitSummary.GAP_ROBOT: "#9a6700",
    KitSummary.GAP_SQ: "#9a6700",
    KitSummary.EXTRA_RD: "#5f259f",
    KitSummary.EXTRA_ROBOT: "#5f259f",
    KitSummary.EXTRA_SQ: "#5f259f",
}

ISSUANCE_EXCLUDE_TOOLTIP = {
    "annulled": "аннулирована",
    "erroneous": "ошибочна",
    "duplicate": "дубликат",
}

OFFICIAL_FOLDER_MTO_MISSING = "нет"
F_MTO_HISTORICAL_MARK = "(и)"
KITS_MTO_REV_TOOLTIP = (
    "Ревизия по имени файла MTO xlsx в официальной папке передачи "
    "(даже если она отстаёт от OD).\n"
    "PDF копии MTO не смотрим.\n"
    "Если в этой папке нет MTO xlsx, ячейка пишет «нет»; "
    "файлы из предыдущих NN не подставляются.\n"
    "Эталон для Авто МТО / сверки / АН МТО."
)
KITS_WORKING_REV_TOOLTIP = (
    "Ревизия строго выше официальной «РД · рев.» "
    "(на диске выше последней выдачи или папка помечена вручную).\n"
    "Суффикс AB — рабочая папка as-build.\n"
    "Не участвует в комплектах, выгрузке спецификаций и сверке MTO."
)
_MTO_COMPARE_PENDING_TIP = (
    "Сверка MTO ещё не завершена; показано по дате файла"
)

WORKLIST_HEADERS = (
    "Титул",
    "Марка",
    "Ревизия (F/РД)",
    "Этап ревизии",
    "Метки",
    "Связь F → РД → MTO",
    "Google",
    "AB",
    PIN_COLUMN_HEADER,
    AUTO_MTO_COMPARE_STATUS_HEADER,
    "Файл MTO",
    "Дата MTO",
    "Пакет",
    "Чего не хватает",
    "Проблемы",
    "Путь MTO",
)

GAP_LABELS = {
    "no_package": "нет папки",
    "no_mto_file": "нет файла MTO",
    "no_rd": "нет в РД",
}

JOURNAL_HEADERS = (
    "Титул",
    "Марка",
    "Источник",
    "Рев.",
    "Дата отпр.",
    "TRM",
    "Статус листа",
    "Наш статус",
    "Комментарий",
    "F",
    "РД",
    "Робот",
    "Авто МТО",
    "Вх.контр.",
    "TRM подтв.",
    "Примечание",
    "Сопоставление",
)

JOURNAL_DECISION_LABELS = {
    "": "не задан",
    "active": "Активна",
    "legalized": "Легализована",
    "annulled": "Аннулирована",
    "erroneous": "Ошибочна",
    "duplicate": "Дубликат",
}
JOURNAL_SOURCE_LABELS = {
    "issuance": "Выдача",
    "google_f": "F",
    "rd": "РД",
    "robot": "Робот",
    "auto_mto": "Авто МТО",
    "manual": "вручную",
}
JOURNAL_MATCH_LABELS = {
    "": "",
    "matched": "сопоставлено",
    "unmatched": "не сопоставлено",
    "ambiguous": "неоднозначно",
}

MTO_READINESS_HEADERS = (
    "Статус",
    "Титул",
    "Марка",
    "Документ",
    "Рев. РД",
    "Рев. робота",
    "As-built",
    "Проверка",
    "Содержимое",
    "Дата РД",
    "Дата робота",
    "MTO РД",
    "MTO робота",
    "Diff / ошибка",
)

MTO_READINESS_FOREGROUND = {
    "ready": "#137333",
    "warning": "#9a6700",
    "blocked": "#b3261e",
}

COLLISION_HEADERS = (
    "Область",
    "Источник",
    "Тип",
    "Документ",
    "Сообщение",
    "Пути",
)

HEATMAP_FIXED_COLUMNS = (
    "Титул",
    "Марка",
    "В папку робота",
    "РД · рев.",
    "Авто МТО",
    AUTO_MTO_COMPARE_STATUS_HEADER,
)

CARD_PACKAGE_HEADERS = (
    "NN",
    "Пакет",
    "Рев.",
    "Дата",
    "Выдача",
    "Рассмотрение",
    "MTO",
    "AB",
    "Ликвидность",
)

_JSON_SKIP_TYPES = (
    CatalogDatabase,
    FileRecord,
    KitMatrixRow,
    KitCard,
    KitPackageRow,
    KitPipelineRow,
    KitRevisionRow,
    MtoWorklistRow,
    AnMtoFile,
    IssuanceJournalRow,
    GoogleKit,
    IssuanceKit,
    AutoMtoFile,
    AutoMtoCompareResult,
    ExportPin,
    ExportSelection,
)


@dataclass(frozen=True, slots=True)
class MonitorCell:
    """One painted table cell shared by Qt and WEB.

    ``href`` is a Google Sheets jump URL for TRM / Google-backed cells.
    """

    text: str
    tooltip: str = ""
    fill: str | None = None
    foreground: str | None = None
    bold: bool = False
    underline: bool = False
    sort_key: object | None = None
    palette_key: str | None = None
    href: str = ""


KITS_PAINT_LEGEND_BUTTON = "Показать легенду"
KITS_PAINT_LEGEND_TITLE = "Легенда таблицы Комплекты"
KITS_PAINT_LEGEND_INTRO = (
    "Жирный шрифт — не «важнее», а совпадение с эталоном.\n"
    "Он есть у «Авто МТО», «Сверка Авто МТО», «АН МТО» "
    "и у «Робот МТО · рев.» когда содержимое совпало с MTO РД "
    "(ревизия при этом может отличаться).\n"
    "У «MTO · рев.» смотрите цвет, не насыщенность букв."
)


@dataclass(frozen=True, slots=True)
class PaintLegendSample:
    """One example cell in the Комплекты paint legend."""

    column: str
    text: str
    meaning: str
    fill: str | None = None
    foreground: str | None = None
    bold: bool = False
    underline: bool = False

    def as_monitor_cell(self) -> MonitorCell:
        """Return the painted example as a table cell DTO."""

        foreground = self.foreground
        if foreground is None and self.fill:
            foreground = contrast_foreground(self.fill)
        return MonitorCell(
            text=self.text,
            fill=self.fill,
            foreground=foreground,
            bold=self.bold,
            underline=self.underline,
        )


@dataclass(frozen=True, slots=True)
class PaintLegendSection:
    """Named group of legend examples."""

    title: str
    intro: str
    samples: tuple[PaintLegendSample, ...]


@dataclass(frozen=True, slots=True)
class MtoKitFlags:
    """Per-kit rollup of the MTO worklist used by Комплекты filters."""

    has_code_a: bool = False
    has_tdo_passed: bool = False
    has_as_build: bool = False
    has_problem: bool = False


@dataclass(frozen=True, slots=True)
class KitsMonitorRow:
    """One Комплекты row with painted cells and filter flags."""

    title: str
    mark: str
    kit_key: tuple[str, str]
    cells: dict[str, MonitorCell]
    haystack: str
    rd_present: bool
    robot_present: bool
    has_google_or_issuance: bool
    summary_aligned: bool
    kit_ok: bool
    code_a: bool
    kit_tdo_passed: bool
    as_build: bool
    has_mto_problem: bool
    an_closes_auto_mto: bool
    pin_view: ExportPinView | None
    matrix_row: KitMatrixRow
    tooltips_text: str = ""


@dataclass(frozen=True, slots=True)
class KitsProgressStats:
    """Universe counts for the Комплекты header progress line."""

    total: int = 0
    ok: int = 0
    aligned: int = 0
    code_a: int = 0
    tdo: int = 0
    no_rd: int = 0
    mto_problems: int = 0
    titles: int = 0


KITS_PROGRESS_VISIBLE_SEP = " · видно: "


def kits_progress_stats(rows: Sequence[KitsMonitorRow]) -> KitsProgressStats:
    """Count kit-level progress flags for the Комплекты header.

    Args:
        rows: Non-banned painted kit rows (title–mark universe).

    Returns:
        Totals used by ``format_kits_progress_stats``. ``ok`` is the
        composite «Ок» flag. ``aligned`` is Сводка «Совпадает»
        (official Google↔РД). ``mto_problems`` is the worklist rollup
        of *any* revision (gap or collision) and can overlap aligned
        kits; it does not block ``ok``.
    """

    titles: set[str] = set()
    ok = aligned = code_a = tdo = no_rd = mto_problems = 0
    for row in rows:
        titles.add(str(row.title).casefold())
        if row.kit_ok:
            ok += 1
        if row.summary_aligned:
            aligned += 1
        if row.code_a:
            code_a += 1
        if row.kit_tdo_passed:
            tdo += 1
        if not row.rd_present:
            no_rd += 1
        if row.has_mto_problem:
            mto_problems += 1
    return KitsProgressStats(
        total=len(rows),
        ok=ok,
        aligned=aligned,
        code_a=code_a,
        tdo=tdo,
        no_rd=no_rd,
        mto_problems=mto_problems,
        titles=len(titles),
    )


def format_kits_progress_stats(
    stats: KitsProgressStats,
    *,
    visible: int | None = None,
) -> str:
    """Return one-line Комплекты progress text.

    Args:
        stats: Counts from ``kits_progress_stats``.
        visible: Currently shown table rows; appended when it differs
            from ``stats.total``.

    Returns:
        Compact Russian line for Qt and WEB. Clients must not rebuild
        the counts in JS or ``window.py``.
    """

    text = (
        f"Ок: {stats.ok} / {stats.total}"
        f" · Совпадает: {stats.aligned} / {stats.total}"
        f" · Код A: {stats.code_a}"
        f" · ТДО: {stats.tdo}"
        f" · Нет в РД: {stats.no_rd}"
        f" · Проблемы MTO: {stats.mto_problems}"
    )
    if visible is not None and visible != stats.total:
        text = f"{text}{KITS_PROGRESS_VISIBLE_SEP}{visible}"
    return text


def kit_ok_reasons(
    *,
    summary_aligned: bool,
    match_flags: Mapping[str, bool | None],
    lag_notes: Sequence[str],
    review_agreed: bool,
    code_a: bool,
) -> tuple[str, ...]:
    """Return Russian blockers for the Комплекты «Ок» flag.

    Empty SQ, empty working revision, yellow Авто МТО / сверка, Google
    D/E mismatch, robot-origin red, and old-revision MTO worklist
    problems are not blockers. ``match_flags`` ``None`` (nothing to
    compare) is not a blocker; only ``False`` is.

    Args:
        summary_aligned: Сводка is «Совпадает».
        match_flags: From ``kit_revision_match_flags``.
        lag_notes: Disk RD behind letter A / TDO-passed rev.
        review_agreed: Displayed review status is ``agreed``.
        code_a: Current letter A of the displayed cycle.

    Returns:
        Blocker lines, empty when the kit is «Ок».
    """

    reasons: list[str] = []
    if not summary_aligned:
        reasons.append("Сводка не «Совпадает»")
    for key in KITS_OK_REV_KEYS:
        if match_flags.get(key) is False:
            reasons.append(KITS_OK_REV_REASON[key])
    if lag_notes:
        reasons.append("РД · рев. отстаёт от письма A / ТДО")
    if not review_agreed:
        reasons.append("Рассмотрение не «Согласован»")
    if not code_a:
        reasons.append("Нет текущего кода A")
    return tuple(reasons)


def kits_ok_cell(ok: bool, reasons: Sequence[str] = ()) -> MonitorCell:
    """Return the painted «Ок» cell (not bold).

    Args:
        ok: Composite flag from ``kit_ok_reasons``.
        reasons: Blocker lines for the tooltip when not ok.

    Returns:
        Green «да» (sort 0) or yellow «нет» (sort 1).
    """

    if ok:
        return MonitorCell(
            text="да",
            tooltip=KITS_OK_TOOLTIP,
            fill=REV_MATCH_FILL,
            foreground=contrast_foreground(REV_MATCH_FILL),
            sort_key=0,
        )
    tip = "\n".join(reasons) if reasons else KITS_OK_TOOLTIP
    return MonitorCell(
        text="нет",
        tooltip=tip,
        fill=REV_DIFF_FILL,
        foreground=contrast_foreground(REV_DIFF_FILL),
        sort_key=1,
    )


@dataclass(frozen=True, slots=True)
class WorklistMonitorRow:
    """One MTO · Перечень row wrapping the domain worklist record."""

    cells: dict[str, MonitorCell]
    source: MtoWorklistRow
    kit_key: tuple[str, str]
    haystack: str = ""


@dataclass(frozen=True, slots=True)
class HeatmapMonitorRow:
    """One Ревизии MTO heatmap row."""

    title: str
    mark: str
    kit_key: tuple[str, str]
    export_cell: MonitorCell
    rd_rev: MonitorCell
    auto_mto: MonitorCell
    auto_mto_compare: MonitorCell
    cells: dict[str, MonitorCell]
    haystack: str = ""


@dataclass(frozen=True, slots=True)
class HeatmapMonitor:
    """Heatmap column set plus one row per kit in the documents universe."""

    revision_columns: tuple[str, ...]
    rows: tuple[HeatmapMonitorRow, ...]


@dataclass(frozen=True, slots=True)
class AnMonitorRow:
    """One АН dump-finder row."""

    cells: dict[str, MonitorCell]
    source: AnMtoFile
    kit_key: tuple[str, str]
    closes_auto_mto: bool
    best_agreed: bool = False
    haystack: str = ""


@dataclass(frozen=True, slots=True)
class JournalMonitorRow:
    """One Выдача · Журнал display row."""

    cells: dict[str, MonitorCell]
    source: IssuanceJournalRow
    kit_key: tuple[str, str]
    haystack: str = ""
    paints_kits_issuance: bool = False


@dataclass(frozen=True, slots=True)
class MtoReadinessMonitorRow:
    """One MTO · Готовность робота display row."""

    cells: dict[str, MonitorCell]
    source: dict[str, Any]
    kit_key: tuple[str, str]
    haystack: str = ""


@dataclass(frozen=True, slots=True)
class CollisionMonitorRow:
    """One current-collision display row (not ban-filtered)."""

    cells: dict[str, MonitorCell]
    source: dict[str, Any]
    haystack: str = ""


@dataclass(frozen=True, slots=True)
class KitCardPackageMonitor:
    """One kit-card package table row with painted cells."""

    nn: str
    name: str
    revision: str
    date: str
    issuance_text: str
    review_text: str
    mto_cell: MonitorCell
    ab_text: str
    liquidity: str
    is_grey: bool
    is_current: bool
    is_as_build: bool
    package_path: str
    review_status: str
    cells: dict[str, MonitorCell]


@dataclass(frozen=True, slots=True)
class KitCardMonitor:
    """Kit card header badges, MTO rollup, and package rows."""

    title: str
    mark: str
    review: MonitorCell
    approval: MonitorCell
    mto_rollup: str
    packages: tuple[KitCardPackageMonitor, ...]
    working_revision_text: str = ""
    official_revision_text: str = ""


@dataclass(frozen=True, slots=True)
class _CardContext:
    database: CatalogDatabase
    worklist_rows: tuple[MtoWorklistRow, ...]
    palette: dict[str, str]
    kit_rows: dict[tuple[str, str], KitMatrixRow]


@dataclass(frozen=True, slots=True)
class CatalogMonitor:
    """Read-only join payload for both catalog clients."""

    kits: tuple[KitsMonitorRow, ...]
    heatmap: HeatmapMonitor
    worklist: tuple[WorklistMonitorRow, ...]
    an_rows: tuple[AnMonitorRow, ...]
    journal: tuple[JournalMonitorRow, ...]
    mto_readiness: tuple[MtoReadinessMonitorRow, ...]
    collisions: tuple[CollisionMonitorRow, ...]
    palette: dict[str, str]
    last_scan: dict[str, Any] | None
    warning: str = ""
    _card_context: _CardContext | None = field(
        default=None, repr=False, compare=False
    )

    def kit_card(self, title: str, mark: str) -> KitCardMonitor | None:
        """Return the painted kit card, loading packages from SQLite.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.

        Returns:
            Card DTO, or ``None`` when the identity is unknown.
        """

        return build_kit_card_monitor(self, title, mark)

    def as_json(self) -> dict[str, Any]:
        """Return a JSON-ready mapping of this monitor."""

        return monitor_to_json(self)


ALL_MONITOR_SHEETS = frozenset(
    {
        "kits",
        "heatmap",
        "worklist",
        "an",
        "journal",
        "readiness",
        "collisions",
    }
)
WEB_MONITOR_SHEETS = frozenset({"kits", "an"})


def contrast_foreground(hex_color: str) -> str:
    """Return white or dark text for a hex fill (Qt HSL lightness).

    Args:
        hex_color: ``#RRGGBB`` background.

    Returns:
        ``#ffffff`` when lightness is below 140, else ``#202124``.
    """

    text = (hex_color or "").strip().lstrip("#")
    if len(text) != 6:
        return _DEFAULT_FOREGROUND
    try:
        red = int(text[0:2], 16)
        green = int(text[2:4], 16)
        blue = int(text[4:6], 16)
    except ValueError:
        return _DEFAULT_FOREGROUND
    lightness = (max(red, green, blue) + min(red, green, blue)) // 2
    return "#ffffff" if lightness < 140 else _DEFAULT_FOREGROUND


def kits_paint_legend(
    palette: Mapping[str, str] | None = None,
) -> tuple[PaintLegendSection, ...]:
    """Return Комплекты paint examples (fill + bold) for Qt and WEB.

    Args:
        palette: Optional user ``status_colors.json``. Factory defaults
            fill in missing keys (working / «Нет в РД»).

    Returns:
        Sections with sample cells. Does not read SQLite or UNC.
    """

    merged = dict(default_palette())
    if palette:
        merged.update(palette)
    working = color_for(merged, "working")
    annulled = color_for(merged, "annulled")
    missing = color_for(merged, RD_MISSING_TRANSFER_KEY)
    mixed = color_for(merged, MIXED_TITLES_KEY)
    no_mto = color_for(merged, "no_mto")
    problem = color_for(merged, "problem")
    code_a = color_for(merged, "code_a")
    exact = f"{AUTO_MTO_EXACT_STATUS} · 01-AN01"
    soft = f"{AUTO_MTO_SOFT_STATUS} · 01-AN01"
    composite = f"{AUTO_MTO_EXACT_STATUS} · сумма 02 + 0"
    missing_rd = f"01-AN01 {RD_MISSING_TRANSFER_NOTE}"
    as_build = f"01-AN02{RD_AB_SUFFIX}"

    def sample(
        column: str,
        text: str,
        meaning: str,
        *,
        fill: str | None = None,
        bold: bool = False,
        underline: bool = False,
        foreground: str | None = None,
    ) -> PaintLegendSample:
        if fill and foreground is None:
            foreground = contrast_foreground(fill)
        return PaintLegendSample(
            column=column,
            text=text,
            meaning=meaning,
            fill=fill,
            foreground=foreground,
            bold=bold,
            underline=underline,
        )

    return (
        PaintLegendSection(
            title="Столбец «Ок»",
            intro=(
                "Комплексный признак, не отдельная сверка. "
                "Жирного нет — только цвет. «да» можно сортировать вверх."
            ),
            samples=(
                sample(
                    KITS_OK_HEADER,
                    "да",
                    "Сводка «Совпадает», зелёные Выдача / F / робот / "
                    "MTO · рев., диск не отстаёт от письма A, "
                    "согласован, текущий код A. Жёлтые Авто МТО / сверка "
                    "и пустые SQ / рабочая не мешают.",
                    fill=REV_MATCH_FILL,
                ),
                sample(
                    KITS_OK_HEADER,
                    "нет",
                    "Не все условия. Подсказка ячейки — список причин.",
                    fill=REV_DIFF_FILL,
                ),
            ),
        ),
        PaintLegendSection(
            title="Жирный шрифт — столбцы MTO",
            intro=(
                "Жирный = совпало с эталоном. Обычный = не совпало или "
                "нечего сравнивать. «—» никогда не жирный. "
                "Эталон «Авто МТО» / «Сверка» — файл MTO РД; "
                "эталон «АН МТО» — столбец «MTO · рев.»."
            ),
            samples=(
                sample(
                    "Авто МТО",
                    "01-AN01",
                    "Ревизия в ПИ совпала с именем файла MTO РД "
                    "(или с ручным выбором). Эталон — MTO, не OD.",
                    fill=REV_MATCH_FILL,
                    bold=True,
                ),
                sample(
                    "Авто МТО",
                    "01-AN01",
                    "Ревизия в ПИ другая, чем у файла MTO РД, "
                    "и не выше официальной «РД · рев.». "
                    "Цвет жёлтый, шрифт обычный.",
                    fill=REV_DIFF_FILL,
                ),
                sample(
                    "Авто МТО",
                    "02-AN01",
                    "Показанная рев. ПИ строго выше официальной "
                    "«РД · рев.» (например 02 при Авто 02-AN01). "
                    "Светлая маджента важнее зелёного/жёлтого "
                    "сравнения с MTO РД.",
                    fill=AUTO_MTO_AHEAD_FILL,
                ),
                sample(
                    "Авто МТО",
                    "—",
                    "Нет спецификации MTO в базе заказчика.",
                ),
                sample(
                    "Авто МТО",
                    "01-AN01 (2)",
                    "Совпало. В скобках — сколько файлов заказчика "
                    "по этому комплекту.",
                    fill=REV_MATCH_FILL,
                    bold=True,
                ),
                sample(
                    AUTO_MTO_COMPARE_STATUS_HEADER,
                    exact,
                    "Содержимое совпало построчно с тем же файлом MTO РД.",
                    fill=REV_MATCH_FILL,
                    bold=True,
                ),
                sample(
                    AUTO_MTO_COMPARE_STATUS_HEADER,
                    soft,
                    "Совпали коды и количества, не все поля строки.",
                    fill=REV_MATCH_FILL,
                    bold=True,
                ),
                sample(
                    AUTO_MTO_COMPARE_STATUS_HEADER,
                    composite,
                    "Совпала сумма нескольких файлов ПИ.",
                    fill=REV_MATCH_FILL,
                    bold=True,
                ),
                sample(
                    AUTO_MTO_COMPARE_STATUS_HEADER,
                    "рев. совпала",
                    "Имена ревизий совпали, содержимое ещё не сверялось.",
                    fill=REV_MATCH_FILL,
                    bold=True,
                ),
                sample(
                    AUTO_MTO_COMPARE_STATUS_HEADER,
                    "не совпало",
                    "Содержимое не совпало.",
                    fill=REV_DIFF_FILL,
                ),
                sample(
                    AUTO_MTO_COMPARE_STATUS_HEADER,
                    "другой файл",
                    "Кэш сверки относится к другому xlsx MTO РД.",
                    fill=REV_DIFF_FILL,
                ),
                sample(
                    AUTO_MTO_COMPARE_STATUS_HEADER,
                    "не сверялось",
                    "Сверка содержимого ещё не сделана. Без заливки.",
                ),
                sample(
                    AUTO_MTO_COMPARE_STATUS_HEADER,
                    "нет в ПИ",
                    "Нет файла заказчика для сверки.",
                ),
                sample(
                    AUTO_MTO_COMPARE_STATUS_HEADER,
                    "нет MTO РД",
                    "Нечего сверять: в официальной папке передачи нет MTO.",
                ),
                sample(
                    "АН МТО",
                    "01-AN01 · 3",
                    "Ревизия АН совпала с официальным «MTO · рев.». "
                    "«· 3» — три файла в папке АН; сравнивается показанный. "
                    "Совпадение с Авто МТО — в подсказке, не жирным.",
                    fill=REV_MATCH_FILL,
                    bold=True,
                ),
                sample(
                    "АН МТО",
                    "01-AN02 · 3",
                    "Ревизия АН другая, чем у MTO РД. "
                    "Даже если она совпала с Авто МТО, ячейка остаётся "
                    "обычной и жёлтой.",
                    fill=REV_DIFF_FILL,
                ),
                sample(
                    "АН МТО",
                    "—",
                    "Нет файлов АН по комплекту.",
                ),
            ),
        ),
        PaintLegendSection(
            title="Робот, MTO РД, выдача",
            intro=(
                "Насыщенная зелёная заливка робота — копия РД/SQ "
                "и та же ревизия, что «MTO · рев.» актуальной папки "
                "(не «РД · рев.» / OD), **или** ручное подтверждение "
                "замены MTO РД с правками. "
                "Жирный — содержимое совпало с MTO РД, даже если "
                "ревизия другая. Красная заливка — не копия и нет "
                "подтверждения. "
                "Синий текст — подтверждённая замена с корректировками "
                "(не сверка байтов и не копия по дате). "
                "Это важнее бледной зелёной / жёлтой сверки ревизий."
            ),
            samples=(
                sample(
                    "Робот МТО · рев.",
                    "01-AN01",
                    "Файл робота — копия MTO РД или SQ "
                    "(дата близкая) и ревизия совпала с «MTO · рев.» "
                    "официальной папки / SQ.",
                    fill=ROBOT_ORIGIN_FILL,
                ),
                sample(
                    "Робот МТО · рев.",
                    "01-AN01",
                    "Копия РД/SQ, ревизия MTO совпала, содержимое тоже. "
                    "Насыщенный зелёный и жирный.",
                    fill=ROBOT_ORIGIN_FILL,
                    bold=True,
                ),
                sample(
                    "Робот МТО · рев.",
                    "03",
                    "Содержимое совпало с MTO РД (жирный), "
                    "но ревизия робота другая, чем «MTO · рев.» "
                    "актуальной папки. Origin-зелёного нет.",
                    fill=REV_DIFF_FILL,
                    bold=True,
                ),
                sample(
                    "Робот МТО · рев.",
                    "03",
                    "Дата близка к MTO РД, содержимое не сверено, "
                    "ревизия робота не совпала с «MTO · рев.» "
                    "актуальной папки (например 03 при MTO 04). "
                    "Origin-зелёного нет. OD / «РД · рев.» не сравнивается.",
                    fill=REV_DIFF_FILL,
                ),
                sample(
                    "Робот МТО · рев.",
                    "0-AN01",
                    "Файл робота не совпал ни с РД, ни с SQ.",
                    fill=ROBOT_ORPHAN_FILL,
                ),
                sample(
                    "Робот МТО · рев.",
                    "01-AN01",
                    "Подтверждена замена MTO РД с правками: зелёная "
                    "заливка вместо красного orphan, синий текст. "
                    "Не копия по дате. Не ставить синий, если "
                    "содержимое уже совпало (жирный).",
                    fill=ROBOT_ORIGIN_FILL,
                    foreground=ROBOT_MTO_ACCEPT_FOREGROUND,
                ),
                sample(
                    "Робот МТО · рев.",
                    "01-AN01",
                    "Ревизия имени совпала с MTO РД, но это не "
                    "«откуда скопировали» (нет origin-заливки).",
                    fill=REV_MATCH_FILL,
                ),
                sample(
                    "Робот МТО · рев.",
                    "01-AN02",
                    "Ревизия робота другая, чем у официального MTO РД.",
                    fill=REV_DIFF_FILL,
                ),
                sample(
                    "MTO · рев.",
                    "01-AN01",
                    "MTO лежит в официальной папке; ревизия совпала с OD "
                    "(столбец «РД · рев.»).",
                    fill=REV_MATCH_FILL,
                ),
                sample(
                    "MTO · рев.",
                    "03",
                    "MTO в официальной папке, имя отстаёт или опережает OD.",
                    fill=REV_DIFF_FILL,
                ),
                sample(
                    "MTO · рев.",
                    OFFICIAL_FOLDER_MTO_MISSING,
                    "В официальной папке нет MTO. Предыдущие передачи "
                    "не подставляются.",
                    fill=no_mto,
                ),
                sample(
                    "Выдача · рев.",
                    "01-AN02",
                    "Совпала с официальным «РД · рев.» (обычно OD). "
                    "То же для Google · рев. / Google · рев. F / SQ.",
                    fill=REV_MATCH_FILL,
                ),
                sample(
                    "Google · рев. F",
                    "01-AN01",
                    "Ревизия F другая, чем у «РД · рев.». Без жирного.",
                    fill=REV_DIFF_FILL,
                ),
                sample(
                    "Google · рев. F",
                    "04 · MTO 03",
                    "MTO в F другая, чем «MTO · рев.» на диске, "
                    "даже если ревизия F совпала с «РД · рев.».",
                    fill=REV_DIFF_FILL,
                ),
                sample(
                    "Google · рев. F",
                    "04 · MTO Нет",
                    "В последней строке F нет MTO, и в более ранних "
                    "передачах его тоже нет. Жёлтый, только если на "
                    "диске всё же лежит MTO xlsx.",
                    fill=REV_MATCH_FILL,
                ),
                sample(
                    "Google · рев. F",
                    f"01-AN02 · MTO 01-AN01 {F_MTO_HISTORICAL_MARK}",
                    "Последняя строка F: MTO Нет или без суффикса MTO; "
                    "ревизия MTO взята из предыдущей передачи. "
                    "Жёлтый, только если она не совпала с MTO xlsx "
                    "в официальной папке.",
                    fill=REV_MATCH_FILL,
                ),
            ),
        ),
        PaintLegendSection(
            title="РД · рев. и рабочая",
            intro=(
                "Суффиксы и серая заливка — про пакет на диске, не про MTO ПИ. "
                "Жёлтый — «Проверить передачи» или диск младше письма A / "
                "ревизии, прошедшей ТДО (важнее зелёного робота)."
            ),
            samples=(
                sample(
                    "РД · рев.",
                    as_build,
                    "Официальный пакет as-build. Подсказка называет "
                    "последнюю IFC без as-build.",
                    fill=REV_MATCH_FILL,
                ),
                sample(
                    "РД · рев.",
                    missing_rd,
                    "Выдача или F есть, пакета в «Для передачи» нет.",
                    fill=missing,
                ),
                sample(
                    "РД · рев.",
                    "01-AN01",
                    "Файлы комплекта лежат в папке другого титула или марки. "
                    "Сводка «Смешанные титулы». Файлы не отбрасываются.",
                    fill=mixed,
                ),
                sample(
                    "РД · рев.",
                    "0-AN01",
                    "Официальная NN не ближе к письму A / отправке, "
                    "или MTO в ней новее кода A. Сводка «Проверить передачи». "
                    "Рабочая папка не считается текущей и не идёт в сверку. "
                    "Папка открытия не меняется.",
                    fill=REV_DIFF_FILL,
                ),
                sample(
                    "РД · рев.",
                    "01-AN01",
                    "На диске младше письма A/B/C (в том числе «новее диска») "
                    "или младше последней F «прошла ТДО» / входной контроль. "
                    "Жёлтый важнее зелёного робота. Серый «Нет в РД» сильнее.",
                    fill=REV_DIFF_FILL,
                ),
                sample(
                    "Рабочая рев. РД",
                    "01-AN02",
                    "Строго выше официальной «РД · рев.» "
                    "(на диске выше последней выдачи или помечена вручную). "
                    "Не для выгрузки и сверки MTO.",
                    fill=working,
                ),
                sample(
                    "Рабочая рев. РД",
                    f"01-AN02{RD_AB_SUFFIX}",
                    "Рабочая папка as-build. Тот же суффикс AB, "
                    "что у официальной «РД · рев.».",
                    fill=working,
                ),
                sample(
                    "Ревизии MTO",
                    "аннулирована",
                    "Все пакеты этой ревизии имени помечены как "
                    "аннулированные. Не в официальном составе, "
                    "конкурсе рабочей, экспорте и сверке.",
                    fill=annulled,
                ),
                sample(
                    "Рабочая рев. РД",
                    "—",
                    "Рабочей ревизии нет.",
                ),
            ),
        ),
        PaintLegendSection(
            title="Другие вкладки — другой смысл жирного",
            intro=(
                "На «Ревизиях MTO» и «MTO · Перечень» жирный — проблемы, "
                "не совпадение ревизии."
            ),
            samples=(
                sample(
                    "Ревизии MTO",
                    "A",
                    "Жирный = есть проблемы по этой ревизии "
                    "(коллизия / разрыв). Заливка — этап F/ТДО.",
                    fill=code_a,
                    bold=True,
                    foreground=problem,
                ),
                sample(
                    "Ревизии MTO",
                    "A",
                    "Подчёркнутый = ревизия текущего состава. "
                    "Курсив (только там) = последняя IFC без as-build.",
                    fill=code_a,
                    underline=True,
                ),
                sample(
                    "MTO · Перечень",
                    "нет файла MTO",
                    "Жирная вся строка = проблемы или разрыв "
                    "(нет папки / файла / РД). Подчёркнутая = текущая.",
                    bold=True,
                    foreground=problem,
                ),
                sample(
                    "АН · vs Авто МТО",
                    "да (четкое)",
                    "На вкладке «АН» жирный «да» = ревизия совпала "
                    "с тем столбцом (Авто МТО или MTO РД). "
                    "В скобках — сверка содержимого, если файл есть.",
                    fill=REV_MATCH_FILL,
                    bold=True,
                ),
                sample(
                    "Карточка · пакет",
                    "05_рев.01-AN01",
                    "Жирная строка пакета = текущий официальный NN, "
                    "заливка строки не меняется.",
                    bold=True,
                ),
            ),
        ),
    )


def revision_or_dash(present: bool, revision_text: str) -> str:
    """Show active revision when present; otherwise an em-dash.

    Args:
        present: True when the source has files or a sheet row.
        revision_text: Filename or sheet revision.

    Returns:
        Revision text, ``есть`` when present without a revision, or ``—``.
    """

    if present and revision_text:
        return revision_text
    if present:
        return "есть"
    return "—"


def join_haystack(*parts: object) -> str:
    """Join filter tokens; ``None`` and empty values become nothing.

    Args:
        *parts: Strings, sequences, or ``None``.

    Returns:
        Space-joined haystack without empty tokens.
    """

    flat: list[str] = []
    for part in parts:
        if isinstance(part, (list, tuple)):
            flat.extend(str(item) for item in part if item)
            continue
        if part is None:
            continue
        text = str(part)
        if text:
            flat.append(text)
    return " ".join(flat)


def format_official_rd_rev_label(
    *,
    disk_present: bool,
    disk_revision: str,
    disk_as_build: bool,
    official_revision: str,
) -> tuple[str, bool]:
    """Return the «РД · рев.» label and whether send/F is missing on disk.

    The primary text is the official **disk** filename revision. When
    issuance/F ``official_revision`` is set and no non-working transfer
    file carries that name, the expected rev is appended as a status.

    Args:
        disk_present: True when the official (non-working) package has files.
        disk_revision: Filename revision of that package.
        disk_as_build: True when that package is as-build.
        official_revision: Last issued send/F revision text.

    Returns:
        Display text and True when the issued revision is not in
        «Для передачи».
    """

    official = (official_revision or "").strip()
    disk = (disk_revision or "").strip() if disk_present else ""
    suffix = RD_AB_SUFFIX if disk_present and disk_as_build else ""
    disk_label = ""
    if disk_present:
        disk_label = revision_or_dash(True, disk) + suffix
    missing = bool(official) and not revision_texts_equivalent(disk, official)
    if missing:
        if disk_label:
            return f"{disk_label} · {official} {RD_MISSING_TRANSFER_NOTE}", True
        return f"{official} {RD_MISSING_TRANSFER_NOTE}", True
    return disk_label or "—", False


def format_working_rd_rev_label(
    working: str,
    official: str,
    *,
    working_as_build: bool = False,
) -> str:
    """Return the Комплекты «Рабочая рев. РД» label, including `` AB``.

    The cell is empty when the working filename rev does not strictly
    outrank official «РД · рев.». As-build is a folder flag, not an
    AGCC filename token.

    Args:
        working: Stored ``kit_pipeline.working_revision_text``.
        official: Official filename rev (``official_revision_text``).
        working_as_build: True when a working folder of that rev is as-build.

    Returns:
        Filename revision with an as-build suffix, or empty string.
    """

    text = working_revision_for_display(working, official)
    if not text:
        return ""
    if working_as_build:
        return text + RD_AB_SUFFIX
    return text


def official_revision_missing_from_transfer(
    row: KitMatrixRow,
    pipeline: KitPipelineRow | None,
) -> bool:
    """Return whether send/F official rev is absent from the transfer folder.

    Args:
        row: Kit matrix row with the official (non-working) RD snapshot.
        pipeline: Derived pipeline, or ``None``.

    Returns:
        True when ``official_revision_text`` is set and not equivalent
        to the disk snapshot revision.
    """

    _text, missing = format_official_rd_rev_label(
        disk_present=row.rd.present,
        disk_revision=row.rd.revision_text,
        disk_as_build=row.rd.as_build,
        official_revision=(
            pipeline.official_revision_text if pipeline is not None else ""
        ),
    )
    return missing


def effective_kit_summary(
    row: KitMatrixRow,
    pipeline: KitPipelineRow | None,
) -> KitSummary:
    """Return the Комплекты / WEB summary after the missing-transfer overlay.

    Official send/F absent from «Для передачи» is ``GAP_RD`` («Нет в РД»),
    even when an older filename still sits in an issued package. That is
    what the «Нет в РД» filter uses. ``GOOGLE_ONLY`` stays when there are
    no RD / robot / SQ files at all. Present RD files under another title
    or mark folder become ``MIXED_TITLES`` («Смешанные титулы»).

    Args:
        row: Kit matrix row (flags stay the scan/Google facts).
        pipeline: Derived pipeline, or ``None``.

    Returns:
        Summary shown in the Сводка cell.
    """

    if row.summary is KitSummary.GOOGLE_ONLY:
        return KitSummary.GOOGLE_ONLY
    if official_revision_missing_from_transfer(row, pipeline):
        return KitSummary.GAP_RD
    if row.mixed_title_notes:
        return KitSummary.MIXED_TITLES
    return row.summary


def official_rd_rev_text(
    row: KitMatrixRow,
    pipeline: KitPipelineRow | None,
) -> str:
    """Return the Комплекты «РД · рев.» label, including `` AB``.

    The cell is the official **disk** filename revision (OD else MTO),
    never Google D/E as if it were on disk. When send/F official is not
    in «Для передачи», the expected rev is appended as ``Нет в РД``.

    Args:
        row: Kit matrix row with the official (non-working) RD snapshot.
        pipeline: Derived pipeline, or ``None``.

    Returns:
        Disk filename revision with an as-build suffix when the official
        snapshot package is as-build; missing-transfer status when the
        issued rev is absent; ``—`` when both are empty.
    """

    text, _missing = format_official_rd_rev_label(
        disk_present=row.rd.present,
        disk_revision=row.rd.revision_text,
        disk_as_build=row.rd.as_build,
        official_revision=(
            pipeline.official_revision_text if pipeline is not None else ""
        ),
    )
    return text


def _official_rd_rev_tooltip(
    *,
    disk_present: bool,
    disk_revision: str,
    disk_as_build: bool,
    official_revision: str,
    missing: bool,
    ifc: str = "",
) -> str:
    disk = (disk_revision or "").strip() if disk_present else ""
    official = (official_revision or "").strip()
    lines: list[str] = []
    if disk:
        lines.append(
            f"Официальная рев. РД: {disk} по имени файла "
            "(рабочая не участвует)."
        )
    else:
        lines.append("Официальная рев. РД: нет файлов в папке передачи.")
    if missing and official:
        lines.append(f"{official} {RD_MISSING_TRANSFER_NOTE}.")
    elif official and not missing:
        lines.append(f"Выдача/F: {official}.")
    if disk_present and disk_as_build:
        lines.append(f"As-build. Последняя IFC: {ifc or '—'}.")
    return "\n".join(lines)


def _revision_text_rank(text: str) -> tuple[int, int, str]:
    revision, appendix = parse_sheet_revision((text or "").strip())
    return revision_rank(revision, appendix)


def _revision_text_outranks(left: str, right: str) -> bool:
    left_text = (left or "").strip()
    right_text = (right or "").strip()
    if not left_text or not right_text:
        return False
    if revision_texts_equivalent(left_text, right_text):
        return False
    return _revision_text_rank(left_text) > _revision_text_rank(right_text)


def rd_rev_review_lag_notes(
    row: KitMatrixRow,
    pipeline: KitPipelineRow | None,
) -> tuple[str, ...]:
    """Return tooltip lines when disk RD lags letter A or a TDO-passed rev.

    Letter A/B/C uses ``pipeline.code_revision_text`` even when the code is
    stale. TDO uses the last Google F ``tdo_passed`` / ``incoming_passed``
    event whose revision outranks the official disk filename. Missing
    send/F on disk stays grey via ``official_revision_missing_from_transfer``.

    Args:
        row: Kit matrix row with the official disk snapshot and Google F.
        pipeline: Derived pipeline, or ``None``.

    Returns:
        Russian tooltip lines, empty when disk is not behind.
    """

    disk = (row.rd.revision_text or "").strip() if row.rd.present else ""
    if not disk:
        return ()
    notes: list[str] = []
    if (
        pipeline is not None
        and pipeline.code
        and pipeline.code_revision_text
    ):
        code_rev = pipeline.code_revision_text.strip()
        if _revision_text_outranks(code_rev, disk):
            relation = pipeline_approval_relation(pipeline)
            stale = f" ({relation})" if relation else ""
            notes.append(
                f"Письмо {pipeline.code} на {code_rev}{stale}, "
                f"в «РД · рев.» {disk}."
            )
    google = row.google
    if google is not None:
        tdo_event = None
        for event in google.events:
            if event.stage not in _RD_TDO_PASSED_STAGES:
                continue
            if format_revision(event.revision, event.appendix):
                tdo_event = event
        if tdo_event is not None:
            tdo_rev = format_revision(tdo_event.revision, tdo_event.appendix)
            if _revision_text_outranks(tdo_rev, disk):
                label = (tdo_event.stage_label or "прошла ТДО").strip()
                notes.append(
                    f"{label[:1].upper()}{label[1:]} на {tdo_rev}, "
                    f"в «РД · рев.» {disk}."
                )
    return tuple(notes)


def _paint_missing_transfer_cell(
    cell: MonitorCell,
    palette: Mapping[str, str],
) -> MonitorCell:
    fill = color_for(palette, RD_MISSING_TRANSFER_KEY)
    return replace(
        cell,
        fill=fill,
        foreground=contrast_foreground(fill),
        palette_key=RD_MISSING_TRANSFER_KEY,
    )


def _paint_mixed_titles_cell(
    cell: MonitorCell,
    palette: Mapping[str, str],
) -> MonitorCell:
    fill = color_for(palette, MIXED_TITLES_KEY)
    return replace(
        cell,
        fill=fill,
        foreground=contrast_foreground(fill),
        palette_key=MIXED_TITLES_KEY,
    )


def build_mto_kit_flags(
    worklist_rows: Sequence[MtoWorklistRow],
) -> dict[tuple[str, str], MtoKitFlags]:
    """Roll per-revision worklist rows into per-kit filter flags.

    Args:
        worklist_rows: Rows from ``list_mto_worklist``.

    Returns:
        Identity → flags. ``has_tdo_passed`` uses per-rev ``TDO_STATUSES``.
    """

    flags: dict[tuple[str, str], MtoKitFlags] = {}
    for row in worklist_rows:
        key = kit_identity_key(row.title, row.mark)
        current = flags.get(key, MtoKitFlags())
        flags[key] = MtoKitFlags(
            has_code_a=current.has_code_a or row.status == "code_a",
            has_tdo_passed=(
                current.has_tdo_passed or row.status in TDO_STATUSES
            ),
            has_as_build=current.has_as_build or row.is_as_build,
            has_problem=(
                current.has_problem
                or bool(row.gap_kind)
                or bool(row.problem_kinds)
            ),
        )
    return flags


def document_tree_kit_identities(
    records: Sequence[FileRecord],
    rd_root: str | Path,
    skip_dirs: Iterable[str],
    banned_keys: Iterable[tuple[str, str]],
) -> set[tuple[str, str]]:
    """Return contour ``(title, mark)`` pairs for the revision heatmap.

    Issued-transfer layout only. Skip-dirs hidden, parsed title+mark,
    banned pairs omitted.

    Args:
        records: Catalog file records.
        rd_root: RD source root.
        skip_dirs: Directory-name skip tokens.
        banned_keys: Banned kit identities.

    Returns:
        Case-insensitive kit keys present on «Все документы».
    """

    banned = {kit_identity_key(title, mark) for title, mark in banned_keys}
    skip = tuple(skip_dirs)
    identities: set[tuple[str, str]] = set()
    for record in records:
        if not record.present:
            continue
        if not record_has_canonical_layout(record, rd_root):
            continue
        if path_has_skipped_dir(record.path, skip):
            continue
        title = str(record.data.get("title") or "").strip()
        mark = str(record.data.get("mark") or "").strip()
        if not title or not mark:
            continue
        key = kit_identity_key(title, mark)
        if key in banned:
            continue
        identities.add(key)
    return identities


def worklist_google_cell(row: MtoWorklistRow) -> MonitorCell:
    """Return the Google-column cell for one worklist row.

    Args:
        row: Domain worklist row.

    Returns:
        Text, tooltip, and problem foreground flag via ``palette_key``.
    """

    if not row.in_google and not row.in_issuance:
        return MonitorCell(
            text="нет комплекта",
            tooltip="Комплекта нет ни в КСБ ИД, ни в «Выдача РД ПД»",
            palette_key="problem",
        )
    if not row.has_f_status:
        return MonitorCell(
            text="нет статуса",
            tooltip=(
                "Комплект есть в Google, но на этой ревизии нет "
                "события F и отправки"
            ),
            palette_key="problem",
        )
    return MonitorCell(
        text="есть",
        tooltip="Есть событие F или отправка на этой ревизии",
    )


def worklist_link_text(row: MtoWorklistRow) -> str:
    """Return compact Google-to-files presence indicators.

    Args:
        row: Domain worklist row.

    Returns:
        ``F✓/✗ · РД✓/✗ · MTO✓/✗``.
    """

    f_mark = "✓" if row.has_f_status else "✗"
    rd_mark = "✓" if row.package_path else "✗"
    mto_mark = "✓" if row.mto_path else "✗"
    return f"F{f_mark} · РД{rd_mark} · MTO{mto_mark}"


def current_ifc_revision_map(
    worklist_rows: Sequence[MtoWorklistRow],
) -> dict[tuple[str, str], str]:
    """Map kit identity to the current IFC revision from the worklist.

    Args:
        worklist_rows: Rows from ``list_mto_worklist``.

    Returns:
        Identity → revision text for ``is_current_ifc`` rows.
    """

    mapping: dict[tuple[str, str], str] = {}
    for row in worklist_rows:
        if not row.is_current_ifc:
            continue
        text = str(row.revision_text or "").strip()
        if not text:
            continue
        mapping[kit_identity_key(row.title, row.mark)] = text
    return mapping


def _official_rd_mto_record(
    row: KitMatrixRow,
    *,
    overlay_mto: Mapping[tuple[str, str], tuple[str, str]],
    records_by_path_key: Mapping[str, FileRecord] | None,
) -> FileRecord | None:
    if not records_by_path_key:
        return None
    key = kit_identity_key(row.title, row.mark)
    overlay = overlay_mto.get(key)
    paths: list[str] = []
    if overlay is not None:
        path = str(overlay[0] or "").strip()
        if path:
            paths.append(path)
    paths.extend(row.rd.paths)
    return mto_xlsx_record(paths, records_by_path_key)


def kits_official_folder_mto_text(
    *,
    rd_present: bool,
    revision_text: str,
) -> str:
    """Return Комплекты «MTO · рев.» text for the official issued folder.

    Args:
        rd_present: True when the kit has official RD files.
        revision_text: Filename revision of MTO in that folder.

    Returns:
        The filename revision, ``нет`` when RD exists but the official
        folder has no MTO, or ``—`` when there is no RD.
    """

    text = (revision_text or "").strip()
    if text:
        return text
    if rd_present:
        return OFFICIAL_FOLDER_MTO_MISSING
    return "—"


def _f_event_names_mto(event: KitEvent | None) -> bool:
    """Return True when the F line names a concrete MTO revision."""

    if event is None or event.mto_absent:
        return False
    return bool(format_revision(event.mto_revision, event.mto_appendix))


def historical_f_mto_event(
    events: Sequence[KitEvent],
    last_event: KitEvent | None,
) -> KitEvent | None:
    """Return the latest earlier F line that names an MTO revision.

    The last F event of a later transfer often says ``MTO Нет`` or has
    no MTO suffix, while an earlier agreed send still named the MTO
    (8260-SKUD 01-AN02; TDO lines without MTO). A last line that already
    names ``MTO <rev>`` is not replaced.

    Args:
        events: Full column-F history, oldest first.
        last_event: Current last dated F event.

    Returns:
        Prior event with a concrete MTO revision, or ``None``.
    """

    if last_event is None or _f_event_names_mto(last_event):
        return None
    remaining = list(events)
    for index in range(len(remaining) - 1, -1, -1):
        event = remaining[index]
        if event is last_event or (
            event.raw == last_event.raw and event.date == last_event.date
        ):
            del remaining[index]
            break
    for event in reversed(remaining):
        if event.mto_absent:
            continue
        if format_revision(event.mto_revision, event.mto_appendix):
            return event
    return None


def format_kits_google_f_rev(
    event: KitEvent | None,
    *,
    events: Sequence[KitEvent] = (),
) -> str:
    """Return Комплекты «Google · рев. F» text.

    OD revision stays the same as ``last_event_parts`` (still a 4-tuple).
    When the F line has an MTO suffix, append `` · MTO <rev>`` or
    `` · MTO Нет``. If the last line is ``MTO Нет`` or has no MTO
    suffix, but an earlier transfer named an MTO revision, show that
    revision with ``(и)``.
    The journal token ``auto`` is not shown here; review/approval cells
    append `` (auto)`` instead.

    Args:
        event: Last Google F event, or ``None``.
        events: Full column-F history (oldest first) for historical MTO.

    Returns:
        Display text such as ``04``, ``04 · MTO 03``, ``04 · MTO Нет``,
        ``01-AN02 · MTO 01-AN01 (и)``, or ``—``.
    """

    if event is None:
        return "—"
    od_text = format_revision(event.revision, event.appendix)
    hist = historical_f_mto_event(events, event)
    parts: list[str] = []
    if od_text:
        parts.append(od_text)
    if hist is not None:
        mto_text = format_revision(hist.mto_revision, hist.mto_appendix)
        parts.append(f"MTO {mto_text} {F_MTO_HISTORICAL_MARK}")
    elif event.mto_absent:
        parts.append(f"MTO {F_LINE_MTO_ABSENT}")
    else:
        mto_text = format_revision(event.mto_revision, event.mto_appendix)
        if mto_text:
            parts.append(f"MTO {mto_text}")
    return " · ".join(parts) or "—"


def format_kits_f_history_line(event: KitEvent) -> str:
    """Return one Google F history line for the Комплекты «Текст» pane.

    Keeps OD rev, TRM, ``MTO <rev|Нет>``, and ``auto`` so the dump does
    not drop the MTO suffix that ``parse_history_line`` split off.

    Args:
        event: One parsed column-F line.

    Returns:
        Indented ``date  stage  rev  TRM  MTO …  auto`` text.
    """

    extra = " ".join(event.transmittals)
    rev = format_revision(event.revision, event.appendix)
    mto = format_revision(event.mto_revision, event.mto_appendix)
    bits: list[str] = [event.date or "—", event.stage_label]
    if rev:
        bits.append(rev)
    if extra:
        bits.append(extra)
    if event.mto_absent:
        bits.append(f"MTO {F_LINE_MTO_ABSENT}")
    elif mto:
        bits.append(f"MTO {mto}")
    if event.from_robot_auto:
        bits.append("auto")
    return "  " + "  ".join(bits)


def _f_mto_historical_tips(
    hist: KitEvent,
    last_event: KitEvent | None = None,
) -> tuple[str, ...]:
    mto = format_revision(hist.mto_revision, hist.mto_appendix)
    od = format_revision(hist.revision, hist.appendix)
    date = hist.date or "—"
    origin = f"MTO {mto} {F_MTO_HISTORICAL_MARK} взята из передачи {date}"
    if od:
        origin += f" (рев. {od})"
    origin += "."
    if last_event is not None and last_event.mto_absent:
        last_note = f"В последней строке F указано MTO {F_LINE_MTO_ABSENT}."
    else:
        last_note = "В последней строке F ревизия MTO не указана."
    return (origin, last_note)


def _f_mto_lag_vs_od_tip(
    f_mto_text: str,
    disk_mto_text: str,
    od_rev: str,
) -> str:
    """Explain when F MTO matches a lagged official-folder filename."""

    od_text = (od_rev or "").strip()
    disk_shown = (disk_mto_text or "").strip()
    if not od_text or not disk_shown:
        return ""
    if revision_texts_match(f_mto_text, od_text) is True:
        return ""
    return f"Имя MTO xlsx {disk_shown} отстаёт от OD {od_text}."


def _official_folder_mto_shown(text: str) -> bool:
    stripped = (text or "").strip()
    return stripped not in {"", "—", OFFICIAL_FOLDER_MTO_MISSING}


def _kits_rev_yellow_wins(cell: MonitorCell) -> MonitorCell:
    """Force mismatch yellow over OD-match green; keep existing yellow."""

    if cell.fill == REV_DIFF_FILL:
        return cell
    fill = REV_DIFF_FILL
    return replace(
        cell,
        fill=fill,
        foreground=contrast_foreground(fill),
    )


def _paint_kits_google_f_rev_cell(
    cell: MonitorCell,
    *,
    event_rev: str,
    rd_rev: str,
    od_match: bool | None,
    f_mto_text: str,
    f_mto_absent: bool,
    disk_mto_text: str,
    f_mto_historical: bool = False,
    f_mto_hist_event: KitEvent | None = None,
    f_last_event: KitEvent | None = None,
) -> MonitorCell:
    """Fill and tooltip for «Google · рев. F» (OD plus optional F MTO)."""

    tips: list[str] = []
    od_text = (event_rev or "").strip()
    rd_text = (rd_rev or "").strip()
    stated = bool(f_mto_text) or f_mto_absent
    if stated:
        if od_text and rd_text:
            if od_match is True:
                tips.append(f"F · рев. {od_text} совпала с РД · рев.")
            elif od_match is False:
                tips.append(
                    f"F · рев. {od_text} расходится с РД · рев. {rd_text}"
                )
        if f_mto_historical and f_mto_hist_event is not None:
            tips.extend(
                _f_mto_historical_tips(
                    f_mto_hist_event, last_event=f_last_event
                )
            )
        disk_shown = _official_folder_mto_shown(disk_mto_text)
        if f_mto_absent:
            if disk_shown:
                cell = _kits_rev_yellow_wins(cell)
                shown_disk = (disk_mto_text or "").strip() or "—"
                tips.append(
                    f"В F указано MTO {F_LINE_MTO_ABSENT}, "
                    f"на диске MTO {shown_disk}"
                )
            else:
                tips.append(
                    f"В F указано MTO {F_LINE_MTO_ABSENT}; "
                    "в официальной папке нет файла MTO xlsx."
                )
        elif not disk_shown:
            cell = _kits_rev_yellow_wins(cell)
            tips.append(
                f"В F указано MTO {f_mto_text}, "
                "в официальной папке нет файла MTO xlsx."
            )
        elif revision_texts_match(f_mto_text, disk_mto_text) is True:
            tips.append(f"F MTO {f_mto_text} совпала с MTO · рев.")
            lag = _f_mto_lag_vs_od_tip(f_mto_text, disk_mto_text, od_text)
            if lag:
                tips.append(lag)
        else:
            cell = _kits_rev_yellow_wins(cell)
            shown_disk = (disk_mto_text or "").strip() or "—"
            tips.append(
                f"F MTO {f_mto_text} расходится с MTO · рев. {shown_disk}"
            )
    for tip in tips:
        cell = _append_tooltip(cell, tip)
    return cell


def _apply_f_mto_to_official_mto_cell(
    cell: MonitorCell,
    *,
    f_mto_text: str,
    f_mto_absent: bool,
    disk_mto_text: str,
    f_mto_historical: bool = False,
    f_mto_hist_event: KitEvent | None = None,
    f_last_event: KitEvent | None = None,
    od_rev: str = "",
) -> MonitorCell:
    """Yellow-win «MTO · рев.» when F names a different MTO revision."""

    disk_shown = _official_folder_mto_shown(disk_mto_text)
    hist_tips = (
        _f_mto_historical_tips(f_mto_hist_event, last_event=f_last_event)
        if f_mto_historical and f_mto_hist_event is not None
        else ()
    )

    def with_hist(cell: MonitorCell) -> MonitorCell:
        for tip in hist_tips:
            cell = _append_tooltip(cell, tip)
        return cell

    if f_mto_absent:
        if not disk_shown:
            cell = _append_tooltip(
                cell,
                f"В F указано MTO {F_LINE_MTO_ABSENT}; "
                "файла MTO xlsx в папке нет.",
            )
            return with_hist(cell)
        shown_disk = (disk_mto_text or "").strip() or "—"
        cell = _kits_rev_yellow_wins(cell)
        cell = _append_tooltip(
            cell,
            f"MTO xlsx на диске {shown_disk}, "
            f"в F указано MTO {F_LINE_MTO_ABSENT}",
        )
        return with_hist(cell)
    if not f_mto_text:
        return cell
    if not disk_shown:
        cell = _append_tooltip(
            cell,
            f"В F указано MTO {f_mto_text}, "
            "файла MTO xlsx в официальной папке нет.",
        )
        return with_hist(cell)
    if revision_texts_match(f_mto_text, disk_mto_text) is True:
        cell = _append_tooltip(
            cell,
            f"F MTO {f_mto_text} совпала с файлом в официальной папке.",
        )
        lag = _f_mto_lag_vs_od_tip(f_mto_text, disk_mto_text, od_rev)
        if lag:
            cell = _append_tooltip(cell, lag)
        return with_hist(cell)
    shown_disk = (disk_mto_text or "").strip() or "—"
    cell = _kits_rev_yellow_wins(cell)
    cell = _append_tooltip(
        cell,
        f"MTO xlsx на диске {shown_disk}, в F указано MTO {f_mto_text}",
    )
    return with_hist(cell)


def _official_package_path_by_kit(
    worklist_rows: Sequence[MtoWorklistRow],
    pipelines: Mapping[tuple[str, str], KitPipelineRow],
) -> dict[tuple[str, str], str]:
    """Return the official issued-folder path per kit from worklist rows."""

    official_pkg: dict[tuple[str, str], str] = {}
    fallback_pkg: dict[tuple[str, str], str] = {}
    for item in worklist_rows:
        if not item.package_path:
            continue
        key = kit_identity_key(item.title, item.mark)
        pipeline = pipelines.get(key)
        official = (
            pipeline.official_revision_text if pipeline is not None else ""
        )
        if official and revision_texts_equivalent(
            item.revision_text, official
        ):
            official_pkg[key] = item.package_path
        elif item.is_current:
            fallback_pkg[key] = item.package_path
    for key, path in fallback_pkg.items():
        official_pkg.setdefault(key, path)
    return official_pkg


def _mto_in_package_sort_key(revision_text: str, path: str) -> tuple:
    revision, appendix = parse_sheet_revision(revision_text)
    return (revision_rank(revision, appendix), path.casefold())


def rd_mto_overlay_by_kit(
    worklist_rows: Sequence[MtoWorklistRow],
    pipelines: Mapping[tuple[str, str], KitPipelineRow],
) -> dict[tuple[str, str], tuple[str, str]]:
    """Official MTO path and filename revision per kit for Auto MTO.

    The file must sit in the official issued folder. Filename revision
    may lag OD (copied ``03`` into the ``04`` package). Older NN folders
    are not searched.

    Args:
        worklist_rows: Worklist rows with ``mto_path``.
        pipelines: Kit pipelines keyed by identity.

    Returns:
        Identity → ``(path, revision_text)``. Empty when the official
        folder has no MTO xlsx.
    """

    official_pkg = _official_package_path_by_kit(worklist_rows, pipelines)
    result: dict[tuple[str, str], tuple[str, str]] = {}
    for item in worklist_rows:
        if not item.mto_path:
            continue
        key = kit_identity_key(item.title, item.mark)
        package = official_pkg.get(key, "")
        if not package or not path_is_under(item.mto_path, package):
            continue
        text = revision_text_from_mto_filename(Path(item.mto_path).name)
        payload = (item.mto_path, text)
        previous = result.get(key)
        if previous is None or _mto_in_package_sort_key(
            payload[1], payload[0]
        ) > _mto_in_package_sort_key(previous[1], previous[0]):
            result[key] = payload
    return result


def auto_mto_rd_target(
    title: str,
    mark: str,
    *,
    pins: Mapping[tuple[str, str], ExportPin],
    overlay_by_kit: Mapping[tuple[str, str], tuple[str, str]],
    selection: ExportSelection | None = None,
) -> tuple[str, str, bool]:
    """Return ``(path, revision, pinned)`` for AutoMTO content compare.

    Heatmap/export selection wins, then a manual pin, then MTO in the
    official issued folder (filename rev may lag OD).

    Args:
        title: Kit title.
        mark: Kit mark.
        pins: Loaded export pins keyed by identity.
        overlay_by_kit: Official MTO path/rev from the worklist.
        selection: Live heatmap export selection, if any.

    Returns:
        Path, filename revision, and whether the target is a pin.
    """

    if selection is not None and selection.source_path:
        pinned = selection.origin in {"pin", "pin_stale"}
        return (
            str(selection.source_path),
            selection.source_revision_text,
            pinned,
        )
    key = kit_identity_key(title, mark)
    pin = pins.get(key)
    if pin is not None and str(pin.file_path or "").strip():
        revision = pin.revision_text or revision_text_from_mto_filename(
            Path(pin.file_path).name
        )
        return str(pin.file_path), revision, True
    path, revision = overlay_by_kit.get(key, ("", ""))
    return path, revision, False


def kit_an_targets(
    row: KitMatrixRow,
    *,
    auto_files: Sequence[AutoMtoFile],
    rd_path: str,
    rd_mto_rev: str,
    pipeline: KitPipelineRow | None = None,
) -> KitAnTargets:
    """Build filename-revision targets for AN matching of one kit.

    Args:
        row: Комплекты matrix row.
        auto_files: Customer-PI Auto MTO files for the kit.
        rd_path: Official or pinned MTO path.
        rd_mto_rev: Filename revision of that MTO.
        pipeline: Derived kit pipeline; supplies the last agreed revision.

    Returns:
        Targets used by ``match_an_to_kit`` and ``score_an_files_for_agreed``.
    """

    auto = pick_auto_mto_file(auto_files, rd_mto_rev)
    auto_path = ""
    if auto is not None:
        candidate = auto_mto_path(auto)
        if candidate.is_file():
            auto_path = str(candidate)
    _event_date, _event_stage, event_rev, _event_trm = last_event_parts(
        row.google
    )
    agreed, agreed_date = agreed_revision_target(
        code=pipeline.code if pipeline is not None else "",
        code_revision_text=(
            pipeline.code_revision_text if pipeline is not None else ""
        ),
        code_date=pipeline.code_date if pipeline is not None else "",
        status=pipeline.status if pipeline is not None else "",
        official_revision_text=(
            pipeline.official_revision_text if pipeline is not None else ""
        ),
    )
    return KitAnTargets(
        auto_mto=auto.rd_revision if auto else "",
        auto_mto_stem=auto.spec if auto else "",
        auto_mto_path=auto_path,
        rd_mto=rd_mto_rev or "",
        rd_mto_path=rd_path or "",
        rd_overlay=row.rd.revision_text,
        robot=row.robot.revision_text if row.robot.present else "",
        sq=row.sq.revision_text if row.sq.present else "",
        issuance=row.issuance.revision_text if row.issuance else "",
        google_sheet=(row.google.sheet_revision_text if row.google else ""),
        google_f=event_rev or "",
        agreed=agreed,
        agreed_date=agreed_date,
    )


def package_liquidity_label(package: KitPackageRow, card: KitCard) -> str:
    """Return the card «Ликвидность» label for one package.

    Args:
        package: Card package row.
        card: Loaded kit card.

    Returns:
        ``OK``, ``неликвид``, ``pending``, or ``—``.
    """

    if package.is_grey or package.source != "rd":
        return "—"
    transfer = (package.transfer_name or "").casefold()
    revision = (package.revision_text or "").casefold()
    review = next(
        (
            item
            for item in card.liquidity_reviews
            if (item.transfer_name or "").casefold() == transfer
            and (item.revision_text or "").casefold() == revision
        ),
        None,
    )
    if review is not None:
        if review.decision == "confirmed_illiquid":
            return "неликвид"
        if review.decision == "confirmed_ok":
            evidence = review.evidence_mtime_ns
            if (
                evidence is not None
                and package.max_mtime_ns is not None
                and package.max_mtime_ns > evidence
            ):
                return "pending"
            return "OK"
    pipeline = card.pipeline
    official = (card.official_revision_text or "").casefold()
    if (
        pipeline is not None
        and pipeline.suspicious
        and official
        and revision == official
    ):
        return "pending"
    return "—"


def kits_tooltip_headers() -> tuple[str, ...]:
    """Return Комплекты headers with Сводка and «РД · рев.» first.

    Returns:
        Header names in the Подсказки pane order.
    """

    priority = set(KITS_TOOLTIP_PRIORITY_HEADERS)
    rest = tuple(header for header in KITS_HEADERS if header not in priority)
    return KITS_TOOLTIP_PRIORITY_HEADERS + rest


def format_kits_row_tooltips(
    title: str,
    mark: str,
    cells: Mapping[str, MonitorCell],
) -> str:
    """Build a readable dump of Комплекты cell tooltips.

    Empty tooltips and tips that only repeat the cell text are skipped.
    Qt and WEB both bind this string; they must not reorder columns.

    Args:
        title: Four-digit title.
        mark: Latin AGCC mark.
        cells: Painted cells keyed by ``KITS_HEADERS``.

    Returns:
        Monospace-friendly plain text for the Подсказки pane. Prose
        becomes one sentence (or ``; `` clause) per line; never
        ``textwrap`` mid-phrase. Tab tables and filesystem paths stay
        on one line.
    """

    blocks: list[str] = [f"{title}-{mark}"]
    for header in kits_tooltip_headers():
        cell = cells.get(header)
        if cell is None:
            continue
        tip = (cell.tooltip or "").strip()
        if not tip:
            continue
        text = (cell.text or "").strip()
        if tip == text:
            continue
        parts = [f"=== {header} ==="]
        if text and text not in tip:
            parts.append(text)
        parts.append(tip)
        blocks.append("\n".join(parts))
    if len(blocks) == 1:
        blocks.append("Нет подсказок для этой строки.")
    return wrap_kits_tips_text("\n\n".join(blocks) + "\n")


def tooltip_has_fs_path(text: str) -> bool:
    """Return True when ``text`` contains a UNC or drive filesystem path."""

    return bool(text) and _FS_PATH_START_RE.search(text) is not None


def _kits_tips_leading_indent(line: str) -> str:
    return line[: len(line) - len(line.lstrip(" "))]


def _kits_tips_path_split(line: str) -> tuple[str, str] | None:
    """Split a tips line into ``(prefix, path_tail)`` when a path is present.

    The path runs to the end of the line so folder names with spaces
    (``На отправку``, ``Ответы на SQ-запросы``) stay intact.
    """

    match = _FS_PATH_START_RE.search(line)
    if match is None:
        return None
    return line[: match.start()], line[match.start() :]


def _kits_tips_keep_raw(line: str) -> bool:
    """Return True when a Подсказки line must stay unwrapped."""

    stripped = line.strip()
    if not stripped:
        return True
    if stripped.startswith("===") and stripped.endswith("==="):
        return True
    if "\t" in line:
        return True
    if tooltip_has_fs_path(line):
        return True
    return False


def _kits_tips_token_before(prefix: str) -> str:
    match = re.search(r"([^\s«\"“(\[]+)$", prefix.rstrip())
    if match is None:
        return ""
    return re.sub(r"[)»\"”\]]+$", "", match.group(1)).casefold()


def _kits_tips_is_abbrev_stop(prefix: str) -> bool:
    token = _kits_tips_token_before(prefix)
    if token in _TIPS_ABBREVS:
        return True
    if token in {"д", "п", "е"} and prefix.rstrip().casefold().endswith(
        f"т.{token}"
    ):
        return True
    return False


def _split_kits_tips_sentences(text: str) -> list[str]:
    """Split prose at real sentence ends, keeping ``рев. РД`` intact."""

    text = text.strip()
    if not text:
        return []
    pieces: list[str] = []
    start = 0
    for match in _TIPS_TERMINATOR_RE.finditer(text):
        after = match.end()
        prefix = text[start:match.start()]
        if _kits_tips_is_abbrev_stop(prefix):
            continue
        rest = text[after:]
        if rest:
            spaces = len(rest) - len(rest.lstrip())
            if spaces == 0:
                continue
            nxt = rest[spaces : spaces + 1]
            if not (nxt.isupper() or nxt in "«\"“"):
                continue
            pieces.append(text[start:after].strip())
            start = after + spaces
            continue
        pieces.append(text[start:after].strip())
        start = after
    tail = text[start:].strip()
    if tail:
        pieces.append(tail)
    return pieces or [text]


def _split_kits_tips_clauses(sentence: str, *, width: int) -> list[str]:
    """Split a long sentence on ``; ``, keeping the semicolon."""

    text = sentence.strip()
    if len(text) <= width or "; " not in text:
        return [text]
    parts = text.split("; ")
    clauses: list[str] = []
    for index, part in enumerate(parts):
        piece = part.strip()
        if not piece:
            continue
        if index < len(parts) - 1 and not piece.endswith(";"):
            piece = f"{piece};"
        clauses.append(piece)
    return clauses or [text]


def _wrap_kits_tips_prose(line: str, width: int) -> str:
    """Put each sentence (and long ``; `` clause) on its own line.

    Never ``textwrap.fill``: mid-phrase breaks like ``или / по дате``
    are forbidden.
    """

    indent = _kits_tips_leading_indent(line)
    sentences = _split_kits_tips_sentences(line.strip())
    if not sentences:
        return line
    packed: list[str] = []
    for sentence in sentences:
        for clause in _split_kits_tips_clauses(sentence, width=width):
            packed.append(f"{indent}{clause}")
    return "\n".join(packed)


def wrap_kits_tips_text(
    text: str, width: int = KITS_TIPS_WRAP_WIDTH
) -> str:
    """Wrap Подсказки prose at sentence / clause ends; keep paths.

    Filesystem paths are peeled onto their own line and never wrapped,
    even when a folder name contains spaces. Several short sentences
    are not packed onto a width budget: one phrase, one line.

    Args:
        text: Pane text from :func:`format_kits_row_tooltips`.
        width: Soft length after which a sentence may also split on
            ``; ``. Not a ``textwrap`` fill width.

    Returns:
        The same text with sentence-aware line breaks.
    """

    if not text:
        return text
    trailing = text.endswith("\n")
    wrapped: list[str] = []
    for line in text.splitlines():
        if "\t" in line:
            wrapped.append(line)
            continue
        split = _kits_tips_path_split(line)
        if split is not None:
            prefix, path = split
            prefix_body = prefix.rstrip()
            if prefix_body:
                wrapped.append(_wrap_kits_tips_prose(prefix_body, width))
                indent = _kits_tips_leading_indent(line)
                wrapped.append(f"{indent}{path}")
            else:
                wrapped.append(line)
            continue
        if _kits_tips_keep_raw(line):
            wrapped.append(line)
            continue
        wrapped.append(_wrap_kits_tips_prose(line, width))
    result = "\n".join(wrapped)
    if trailing:
        result += "\n"
    return result


def format_kits_tips_pane(body: str, *, with_ctrl_click_hint: bool = False) -> str:
    """Return Подсказки pane text, optionally with the Qt Ctrl+click hint.

    Args:
        body: Output of :func:`format_kits_row_tooltips`.
        with_ctrl_click_hint: True in the Qt GUI only (WEB has no table Ctrl+click).

    Returns:
        Pane text with a trailing newline, or empty when ``body`` is empty.
    """

    text = (body or "").rstrip("\n")
    if not text:
        return ""
    if with_ctrl_click_hint:
        return f"{KITS_TIPS_CTRL_CLICK_HINT}\n\n{text}\n"
    return f"{text}\n"


def kits_tooltip_header_offset(text: str, header: str) -> int | None:
    """Return the character offset of ``=== header ===``, or None.

    Args:
        text: Pane text from :func:`format_kits_tips_pane` or
            :func:`format_kits_row_tooltips`.
        header: Комплекты column title (``KITS_HEADERS``).

    Returns:
        Start offset of the matching header line, or None when that column
        has no tooltip block.
    """

    needle = f"=== {(header or '').strip()} ==="
    if not text or needle == "===  ===":
        return None
    offset = 0
    for line in text.splitlines(keepends=True):
        if line.strip() == needle:
            return offset
        offset += len(line)
    return None


def kits_tooltip_section_start(text: str, pos: int) -> int:
    """Return the character offset of the ``=== header ===`` block at ``pos``.

    Positions in the title line before the first block return ``0``. The last
    block whose header starts at or before ``pos`` wins, so a position on the
    blank line after a block still belongs to that block.

    Args:
        text: Plain text from :func:`format_kits_row_tooltips`.
        pos: Character offset in ``text``.

    Returns:
        Non-negative start offset, clamped to the text length.
    """

    if not text:
        return 0
    pos = max(0, min(int(pos), len(text)))
    start = 0
    offset = 0
    for line in text.splitlines(keepends=True):
        stripped = line.strip()
        if stripped.startswith("=== ") and stripped.endswith(" ==="):
            line_start = offset
            if line_start <= pos:
                start = line_start
            else:
                break
        offset += len(line)
    return start


def format_kits_card_mto(
    worklist_rows: Sequence[MtoWorklistRow],
    row: KitMatrixRow,
) -> str:
    """One-line MTO rollup for the selected kit from the worklist.

    Args:
        worklist_rows: All worklist rows.
        row: Selected kit matrix row.

    Returns:
        Compact Russian rollup, or ``MTO: нет данных``.
    """

    kit_key = kit_identity_key(row.title, row.mark)
    items = [
        item
        for item in worklist_rows
        if kit_identity_key(item.title, item.mark) == kit_key
    ]
    if not items:
        return "MTO: нет данных"
    current = next((item for item in items if item.is_current), None)
    if current is None:
        current = max(items, key=_worklist_revision_rank)
    parts = [
        f"MTO: тек. рев. {current.revision_text or '—'}",
        status_short_label(current.status) or "—",
    ]
    missing = sum(1 for item in items if item.gap_kind)
    problems = sum(1 for item in items if item.problem_kinds)
    if missing:
        parts.append(f"нет файла: {missing}")
    if problems:
        parts.append(f"проблемы: {problems}")
    return " · ".join(parts)


def package_review_status(
    package: KitPackageRow,
    kit_row: KitMatrixRow,
    worklist_rows: Sequence[MtoWorklistRow],
) -> str:
    """Status key for the card «Рассмотрение» cell from the worklist.

    Args:
        package: Card package.
        kit_row: Parent kit matrix row.
        worklist_rows: Worklist rows.

    Returns:
        Pipeline status token, or empty.
    """

    kit_key = kit_identity_key(kit_row.title, kit_row.mark)
    path = (package.package_path or "").casefold()
    matched_path: MtoWorklistRow | None = None
    matched_rev: MtoWorklistRow | None = None
    for item in worklist_rows:
        if kit_identity_key(item.title, item.mark) != kit_key:
            continue
        if (
            path
            and matched_path is None
            and (item.package_path or "").casefold() == path
        ):
            matched_path = item
        elif matched_rev is None and revision_texts_equivalent(
            item.revision_text, package.revision_text
        ):
            matched_rev = item
        if matched_path is not None:
            break
    chosen = matched_path or matched_rev
    return (chosen.status if chosen is not None else "") or ""


def kits_auto_mto_cell(
    row: KitMatrixRow,
    *,
    auto_files: Sequence[AutoMtoFile],
    rd_path: str,
    rd_mto_rev: str,
    pinned: bool,
) -> MonitorCell:
    """Paint the Комплекты «Авто МТО» cell vs official/pinned MTO.

    Green/yellow still compare the shown PI revision with the MTO
    filename (or pin). Light magenta wins when that shown revision
    strictly outranks official disk «РД · рев.» (OD-first, not working).

    Args:
        row: Kit matrix row; official disk revision comes from
            ``row.rd``.
        auto_files: Customer-PI files for the kit.
        rd_path: Target RD MTO path.
        rd_mto_rev: Target filename revision.
        pinned: True when the target is a manual pin.

    Returns:
        Text, revision-match fill, bold on match, and file tooltip.
    """

    del rd_path
    text = format_auto_mto_cell_text(auto_files, rd_mto_rev) or "—"
    auto = pick_auto_mto_file(auto_files, rd_mto_rev)
    if auto is None or not auto.rd_revision:
        return MonitorCell(
            text=text if text != "—" else "—",
            tooltip="Нет спецификации MTO в базе заказчика.",
        )
    match = revision_texts_match(auto.rd_revision, rd_mto_rev)
    rd_rev = (row.rd.revision_text or "").strip() if row.rd.present else ""
    ahead = _revision_text_outranks(auto.rd_revision, rd_rev)
    target_note = (
        f"ручной выбор: {rd_mto_rev}"
        if pinned and rd_mto_rev
        else f"MTO РД: {rd_mto_rev}"
        if rd_mto_rev
        else "MTO РД: нет файла"
    )
    shown_note = " (совпала с РД)" if match is True else ""
    count_note = f"; файлов: {len(auto_files)}" if len(auto_files) > 1 else ""
    tips = [
        "Ревизия Авто МТО vs имя файла MTO РД, не OD текущего состава.",
        f"Показана рев. {auto.rd_revision}{shown_note}{count_note}",
        target_note,
        "Статус содержимого — столбец «Сверка Авто МТО».",
        "Файлы заказчика:",
    ]
    for file in auto_files:
        sources = ", ".join(file.source_specs) if file.source_specs else file.spec
        mark = " ← показано" if file.relpath == auto.relpath else ""
        tips.append(
            f"  {file.rd_revision}  {Path(file.relpath).name}  [{sources}]{mark}"
        )
    path = auto_mto_path(auto)
    if path.is_file():
        tips.append(str(path))
    else:
        tips.append(
            "xlsx ещё не собран (База заказчика → Собрать каталог АвтоМТО)"
        )
    fill = AUTO_MTO_AHEAD_FILL if ahead else _revision_fill(match)
    if ahead:
        tips.append(
            f"Авто МТО выше официальной «РД · рев.» ({rd_rev})."
        )
    return MonitorCell(
        text=text,
        tooltip="\n".join(tips),
        fill=fill,
        foreground=contrast_foreground(fill) if fill else None,
        bold=match is True,
        sort_key=auto.rd_revision,
    )


def kits_an_cell(hit: Any) -> MonitorCell:
    """Paint the Комплекты «АН МТО» cell vs official «MTO · рев.».

    Args:
        hit: ``AnKitHit`` from ``match_an_to_kit``.

    Returns:
        Compact AN label, match fill vs ``match_rd_mto``, bold on match,
        file tooltip (Auto MTO stays a separate note).
    """

    text = an_cell_text(hit)
    match = hit.match_rd_mto
    lines = [
        "Покраска: ревизия АН vs «MTO · рев.» (официальный MTO РД), "
        "не OD текущего состава.",
    ]
    if hit.closes_auto_mto:
        lines.append("АН закрывает Авто МТО.")
    if hit.match_auto_mto is True and hit.match_rd_mto is not True:
        lines.append("Ревизия совпала с Авто МТО, но не с MTO · рев.")
    elif hit.match_auto_mto is True:
        lines.append("Также совпала с Авто МТО.")
    if hit.stem_note:
        lines.append(hit.stem_note)
    if hit.files:
        lines.append("Файлы АН:")
        for file in hit.files:
            date = _mtime_ymd(file.mtime_ns)
            mark = (
                " ← показано"
                if hit.shown_file is not None
                and file.path_key == hit.shown_file.path_key
                else ""
            )
            identity = f"  {file.revision_text}  {file.name}  {date}"
            folder = (file.parent_dir or "").strip()
            if folder:
                lines.append(identity)
                lines.append(f"  {folder}{mark}")
            else:
                lines.append(f"{identity}{mark}")
    else:
        lines.append("Нет файлов АН.")
    return MonitorCell(
        text=text,
        tooltip="\n".join(lines),
        fill=_revision_fill(match),
        bold=match is True,
    )


def load_catalog_monitor(
    database: CatalogDatabase,
    config: CatalogConfig,
    *,
    selections: Mapping[tuple[str, str], ExportSelection] | None = None,
    comparisons: (
        Mapping[tuple[str, str], tuple[str, AutoMtoCompareResult]] | None
    ) = None,
    sheets: Collection[str] | None = None,
) -> CatalogMonitor:
    """Assemble painted monitor rows from SQLite and JSON caches.

    Does not call ``rebuild_pipeline``, ingest, or scan. Empty
    ``kit_pipeline`` yields a warning and review cells of ``—``; kits are
    still built from files/Google.

    Args:
        database: Initialized catalog database (schema 10).
        config: Catalog configuration (runtime dir, RD root).
        selections: Optional live heatmap export selections.
        comparisons: Optional in-memory Auto MTO compares. When ``None``,
            hydrate from ``auto_mto_compare_cache.json`` for the current
            RD path.
        sheets: Optional subset of ``ALL_MONITOR_SHEETS``. ``None`` builds
            every sheet (Qt). WEB uses ``WEB_MONITOR_SHEETS`` (kits + АН)
            and skips heatmap / journal / readiness / collisions DTO.

    Returns:
        Join DTO for Qt bind and WEB JSON.
    """

    wanted = ALL_MONITOR_SHEETS if sheets is None else (
        frozenset(sheets) & ALL_MONITOR_SHEETS
    )
    if not wanted:
        wanted = frozenset({"kits"})

    with perf_span("monitor.load_catalog_monitor"):
        records = database.list_files()
        records_by_id = {record.id: record for record in records}
        overlay_rows = database.current_overlay()
        detected = _overlay_current_ids(
            overlay_rows, records_by_id, str(config.rd_root)
        )
        google_kits, issuance_kits, google_loaded = _load_google_for_monitor(
            database, str(config.runtime_dir)
        )
        warning = ""
        pipelines: tuple[KitPipelineRow, ...] = ()
        try:
            pipelines = list_kit_pipelines(database)
        except Exception as exc:
            warning = f"Pipeline list: {type(exc).__name__}: {exc}"
        if not pipelines and not warning:
            if records or google_kits or issuance_kits or google_loaded:
                warning = EMPTY_PIPELINE_WARNING
        pipeline_by_kit = {
            kit_identity_key(item.title, item.mark): item for item in pipelines
        }
        official_ids = official_detected_current_ids(
            records, detected, pipelines
        )
        with perf_span("monitor.build_kit_matrix"):
            kit_rows = build_kit_matrix(
                google_kits,
                records,
                official_ids,
                issuance_kits=issuance_kits,
                rd_root=config.rd_root,
                mto_compare_by_pair=path_pair_labels_from_cache(config.runtime_dir),
                working_folders_by_kit=working_folders_from_pipelines(pipelines),
                annulled_folders_by_kit=annulled_folders_from_pipelines(pipelines),
            )
        kit_by_key = {
            kit_identity_key(row.title, row.mark): row for row in kit_rows
        }
        with perf_span("monitor.list_mto_worklist"):
            try:
                worklist_rows = list_mto_worklist(
                    database, records=records, rd_root=config.rd_root
                )
            except Exception:
                worklist_rows = ()
        revision_cells: tuple[Any, ...] = ()
        if "heatmap" in wanted:
            with perf_span("monitor.list_revision_matrix"):
                try:
                    revision_cells = tuple(list_revision_matrix(database))
                except Exception:
                    revision_cells = ()
        try:
            an_by_kit = database.list_an_files_by_kit()
        except Exception:
            an_by_kit = {}
        auto_mto_by_kit: dict[tuple[str, str], tuple[AutoMtoFile, ...]] = {}
        try:
            auto_mto_by_kit = list_auto_mto_files_by_kit()
        except FileNotFoundError:
            auto_mto_by_kit = {}
        except Exception:
            auto_mto_by_kit = {}
        mto_display = _build_mto_display_rows(
            database,
            records_by_id,
            overlay_rows,
            pending_error=_NOT_COMPARED_ERROR,
        )
        try:
            collision_rows = (
                database.list_current_collisions()
                if "collisions" in wanted
                else []
            )
        except Exception:
            collision_rows = []
        try:
            last_scan = database.last_scan_info(successful_only=True)
        except Exception:
            last_scan = None
        palette = load_status_colors(status_colors_path(config.runtime_dir))
        pins = {
            kit_identity_key(pin.title, pin.mark): pin
            for pin in load_export_pins(config)
        }
        overlay_mto = rd_mto_overlay_by_kit(worklist_rows, pipeline_by_kit)
        selection_map = dict(selections or {})
        if comparisons is None:
            comparison_map = _hydrate_auto_mto_comparisons(
                auto_mto_by_kit=auto_mto_by_kit,
                pins=pins,
                overlay_by_kit=overlay_mto,
                selections=selection_map,
                runtime_dir=config.runtime_dir,
            )
        else:
            comparison_map = dict(comparisons)
        ban_store = BanFilterStore.from_runtime_dir(config.runtime_dir)
        banned_keys = {pair.identity for pair in ban_store.pairs()}
        runtime_skip = load_runtime_skip_dirs(config.runtime_dir)
        skip_dirs = (
            tuple(runtime_skip)
            if runtime_skip is not None
            else tuple(config.skip_dirs)
        )
        mto_flags = build_mto_kit_flags(worklist_rows)
        mto_content = mto_content_equal_by_kit(mto_display)
        ifc_by_kit = current_ifc_revision_map(worklist_rows)
        excluded_sends = _excluded_issuance_sends(database)
        an_cache = load_an_content_compare_cache(config.runtime_dir)
        sheet_links = sheet_link_context_from_config(config)
        records_by_path_key = {record.path_key: record for record in records}
        accept_by_kit = accepts_by_kit(database.list_robot_mto_accepts())

        kits: list[KitsMonitorRow] = []
        for row in kit_rows:
            key = kit_identity_key(row.title, row.mark)
            if key in banned_keys:
                continue
            kits.append(
                build_kits_monitor_row(
                    row,
                    pipeline=pipeline_by_kit.get(key),
                    palette=palette,
                    pins=pins,
                    overlay_mto=overlay_mto,
                    selection=selection_map.get(key),
                    comparison=comparison_map.get(key),
                    auto_files=auto_mto_by_kit.get(key, ()),
                    an_files=an_by_kit.get(key, ()),
                    mto_flags=mto_flags.get(key, MtoKitFlags()),
                    mto_content_equal=mto_content.get(key),
                    ifc_by_kit=ifc_by_kit,
                    excluded_sends=excluded_sends.get(key, ()),
                    sheet_links=sheet_links,
                    robot_mto_accept=accept_by_kit.get(key),
                    records_by_path_key=records_by_path_key,
                )
            )

        heatmap = HeatmapMonitor(revision_columns=(), rows=())
        if "heatmap" in wanted:
            heatmap = _build_heatmap(
                revision_cells,
                records=records,
                official_ids=official_ids,
                rd_root=config.rd_root,
                skip_dirs=skip_dirs,
                banned_keys=banned_keys,
                palette=palette,
                auto_mto_by_kit=auto_mto_by_kit,
                pins=pins,
                overlay_mto=overlay_mto,
                selections=selection_map,
                comparisons=comparison_map,
                pipeline_by_kit=pipeline_by_kit,
            )

        worklist: list[WorklistMonitorRow] = []
        if "worklist" in wanted:
            for row in worklist_rows:
                key = kit_identity_key(row.title, row.mark)
                if key in banned_keys:
                    continue
                worklist.append(
                    _build_worklist_row(
                        row,
                        palette=palette,
                        pins=pins,
                        overlay_mto=overlay_mto,
                        selection=selection_map.get(key),
                        comparison=comparison_map.get(key),
                        auto_files=auto_mto_by_kit.get(key, ()),
                    )
                )

        an_rows: list[AnMonitorRow] = []
        if "an" in wanted:
            for files in an_by_kit.values():
                if not files:
                    continue
                key = kit_identity_key(files[0].title, files[0].mark)
                if key in banned_keys:
                    continue
                matrix = kit_by_key.get(key)
                path, rd_mto, _pinned = auto_mto_rd_target(
                    files[0].title,
                    files[0].mark,
                    pins=pins,
                    overlay_by_kit=overlay_mto,
                    selection=selection_map.get(key),
                )
                pipeline = pipeline_by_kit.get(key)
                if matrix is not None:
                    targets = kit_an_targets(
                        matrix,
                        auto_files=auto_mto_by_kit.get(key, ()),
                        rd_path=path,
                        rd_mto_rev=rd_mto,
                        pipeline=pipeline,
                    )
                else:
                    agreed, agreed_date = agreed_revision_target(
                        code=pipeline.code if pipeline is not None else "",
                        code_revision_text=(
                            pipeline.code_revision_text if pipeline is not None else ""
                        ),
                        code_date=pipeline.code_date if pipeline is not None else "",
                        status=pipeline.status if pipeline is not None else "",
                        official_revision_text=(
                            pipeline.official_revision_text if pipeline is not None else ""
                        ),
                    )
                    targets = KitAnTargets(
                        rd_mto=rd_mto,
                        rd_mto_path=path,
                        agreed=agreed,
                        agreed_date=agreed_date,
                    )
                scores = score_an_files_for_agreed(files, targets)
                for file in files:
                    an_rows.append(
                        _build_an_row(
                            file,
                            targets=targets,
                            cache=an_cache,
                            auto_files=auto_mto_by_kit.get(key, ()),
                            rd_mto_rev=rd_mto,
                            agreed_score=scores.get(file.path_key),
                        )
                    )

        journal: list[JournalMonitorRow] = []
        if "journal" in wanted:
            journal_hits = [
                JournalAutoMtoHit(file.title, file.mark, file.rd_revision, file.relpath)
                for files in auto_mto_by_kit.values()
                for file in files
            ]
            try:
                journal_source = list_issuance_journal(
                    database, records=records, auto_mto_hits=journal_hits
                )
            except Exception:
                journal_source = ()
            for row in journal_source:
                key = kit_identity_key(row.title, row.mark)
                if key in banned_keys:
                    continue
                matrix = kit_by_key.get(key)
                issuance = matrix.issuance if matrix is not None else None
                journal.append(
                    _build_journal_row(
                        row, issuance=issuance, sheet_links=sheet_links
                    )
                )

        readiness: list[MtoReadinessMonitorRow] = []
        if "readiness" in wanted:
            for row in mto_display:
                title = str(row.get("title") or "")
                mark = str(row.get("mark") or "")
                if title and mark and kit_identity_key(title, mark) in banned_keys:
                    continue
                readiness.append(_build_readiness_row(row))

        collisions = (
            tuple(_build_collision_row(row) for row in collision_rows)
            if "collisions" in wanted
            else ()
        )

        return CatalogMonitor(
            kits=tuple(kits),
            heatmap=heatmap,
            worklist=tuple(worklist),
            an_rows=tuple(an_rows),
            journal=tuple(journal),
            mto_readiness=tuple(readiness),
            collisions=collisions,
            palette=dict(palette),
            last_scan=last_scan,
            warning=warning,
            _card_context=_CardContext(
                database=database,
                worklist_rows=worklist_rows,
                palette=dict(palette),
                kit_rows=kit_by_key,
            ),
        )


def paint_kit_card(
    *,
    database: CatalogDatabase | None,
    title: str,
    mark: str,
    kit_row: KitMatrixRow | None,
    worklist_rows: Sequence[MtoWorklistRow] = (),
    palette: Mapping[str, str] | None = None,
) -> KitCardMonitor | None:
    """Paint one kit card without assembling the full catalog monitor.

    The Qt Комплекты pane must not call ``load_catalog_monitor`` just to
    fill the package table. WEB still goes through ``CatalogMonitor.kit_card``.

    Args:
        database: Catalog DB, or ``None`` when only the matrix row is known.
        title: Four-digit title.
        mark: Latin AGCC mark.
        kit_row: Already-joined Комплекты row, if the GUI holds it.
        worklist_rows: MTO worklist for the rollup and package paint.
        palette: Status colors. Empty mapping uses uncolored badges.

    Returns:
        Painted card, or ``None`` when the identity is unknown.
    """

    row = kit_row
    card: KitCard | None = None
    if database is not None:
        try:
            card = get_kit_card(database, title, mark)
        except Exception:
            card = None
    if card is None and row is None:
        return None
    display_title = card.title if card is not None else (row.title if row else title)
    display_mark = card.mark if card is not None else (row.mark if row else mark)
    pipeline = card.pipeline if card is not None else None
    colors = dict(palette or {})
    google = card.google if card is not None else None
    issuance = card.issuance if card is not None else None
    review_status = (
        pipeline_display_review_status(
            pipeline, google=google, issuance=issuance
        )
        if pipeline is not None
        else None
    )
    review = _pipeline_badge_cell(
        _pipeline_review_text(
            pipeline,
            google=google,
            issuance=issuance,
        ),
        color_for(colors, review_status) if review_status is not None else None,
        pipeline,
        kind="review",
        google=google,
        issuance=issuance,
    )
    approval_key = (
        pipeline_approval_color_key(
            pipeline, google=google, issuance=issuance
        )
        if pipeline is not None
        else None
    )
    approval = _pipeline_badge_cell(
        pipeline_approval_label(
            pipeline, google=google, issuance=issuance
        )
        if pipeline is not None
        else "—",
        color_for(colors, approval_key) if approval_key else None,
        pipeline,
        kind="approval",
        google=google,
        issuance=issuance,
    )
    mto_rollup = (
        format_kits_card_mto(worklist_rows, row)
        if row is not None
        else "MTO: нет данных"
    )
    packages: tuple[KitCardPackageMonitor, ...] = ()
    if card is not None and row is not None and database is not None:
        packages = _build_card_packages(
            card, row, worklist_rows, colors, database
        )
    official = (
        (card.official_revision_text if card is not None else "")
        or (pipeline.official_revision_text if pipeline is not None else "")
        or (row.rd.revision_text if row is not None else "")
    )
    working = format_working_rd_rev_label(
        (card.working_revision_text if card is not None else "")
        or (pipeline.working_revision_text if pipeline is not None else ""),
        official,
        working_as_build=bool(
            pipeline is not None and pipeline.working_as_build
        ),
    )
    return KitCardMonitor(
        title=display_title,
        mark=display_mark,
        review=review,
        approval=approval,
        mto_rollup=mto_rollup,
        packages=packages,
        working_revision_text=working,
        official_revision_text=official,
    )


def build_kit_card_monitor(
    monitor: CatalogMonitor,
    title: str,
    mark: str,
) -> KitCardMonitor | None:
    """Load and paint one kit card from the monitor's SQLite context.

    Args:
        monitor: Assembled catalog monitor.
        title: Four-digit title.
        mark: Latin AGCC mark.

    Returns:
        Painted card, or ``None`` when unknown.
    """

    ctx = monitor._card_context
    key = kit_identity_key(title, mark)
    row = ctx.kit_rows.get(key) if ctx is not None else None
    return paint_kit_card(
        database=ctx.database if ctx is not None else None,
        title=title,
        mark=mark,
        kit_row=row,
        worklist_rows=ctx.worklist_rows if ctx is not None else (),
        palette=ctx.palette if ctx is not None else monitor.palette,
    )


def monitor_to_json(monitor: CatalogMonitor) -> dict[str, Any]:
    """Convert a monitor DTO into JSON-serializable mappings.

    Skips ``KitMatrixRow``, ``FileRecord``, SQLite handles, and other
    non-JSON domain objects. Kit-key tuples become ``[title, mark]``.
    Enums become their ``value``.

    Args:
        monitor: Assembled catalog monitor.

    Returns:
        Nested dict/list structure safe for ``json.dumps``.
    """

    payload = asdict(monitor, dict_factory=_json_dict_factory)
    return payload


def _json_dict_factory(items: Sequence[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in items:
        if key.startswith("_"):
            continue
        if _is_non_json(value):
            continue
        result[str(key)] = _jsonify(value)
    return result


def _is_non_json(value: object) -> bool:
    if value is None:
        return False
    return isinstance(value, _JSON_SKIP_TYPES)


def jsonify_value(value: Any) -> Any:
    """Return a JSON-safe form of *value* (enums, paths, monitor DTOs).

    Args:
        value: Arbitrary monitor payload.

    Returns:
        Nested dict/list/scalars suitable for ``json.dumps``.
    """

    return _jsonify(value)


def _jsonify(value: Any) -> Any:
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, Enum):
        return value.value
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    if _is_non_json(value):
        return None
    if isinstance(value, tuple) and len(value) == 2 and all(
        isinstance(item, str) for item in value
    ):
        return [value[0], value[1]]
    if isinstance(value, dict):
        converted: dict[str, Any] = {}
        for key, item in value.items():
            if _is_non_json(item):
                continue
            if isinstance(key, tuple):
                mapped = "|".join(str(part) for part in key)
            else:
                mapped = str(key)
            converted[mapped] = _jsonify(item)
        return converted
    if isinstance(value, (list, tuple)):
        return [
            _jsonify(item) for item in value if not _is_non_json(item)
        ]
    if is_dataclass(value) and not isinstance(value, type):
        return asdict(value, dict_factory=_json_dict_factory)
    if isinstance(value, Mapping):
        return {str(key): _jsonify(item) for key, item in value.items()}
    return str(value)


def _revision_fill(matches: bool | None) -> str | None:
    if matches is True:
        return REV_MATCH_FILL
    if matches is False:
        return REV_DIFF_FILL
    return None


def _append_tooltip(cell: MonitorCell, extra: str) -> MonitorCell:
    if not extra:
        return cell
    tip = f"{cell.tooltip}\n{extra}" if cell.tooltip else extra
    return replace(cell, tooltip=tip)


def _dash(value: object) -> str:
    text = str(value or "").strip()
    return text if text else "—"


def _mtime_clock(value: Any) -> str:
    try:
        return datetime.fromtimestamp(int(value) / 1_000_000_000).strftime(
            "%Y-%m-%d %H:%M"
        )
    except (TypeError, ValueError, OSError):
        return "—"


def _mtime_ymd(value: Any) -> str:
    try:
        return datetime.fromtimestamp(int(value) / 1_000_000_000).strftime(
            "%Y-%m-%d"
        )
    except (TypeError, ValueError, OSError):
        return "—"


def _mtime_dot(value: int | None) -> str:
    if value is None:
        return ""
    try:
        return datetime.fromtimestamp(value / 1_000_000_000).strftime("%Y.%m.%d")
    except (OSError, OverflowError, ValueError):
        return ""


def _revision_pair(revision: Any, appendix: Any) -> str:
    value = str(revision or "—")
    return f"{value} AN{appendix}" if appendix else value


def _worklist_revision_rank(item: MtoWorklistRow) -> tuple[int, int, str]:
    revision, appendix = parse_sheet_revision(item.revision_text)
    return revision_rank(revision, appendix)


def _pipeline_fill_cell(
    text: str,
    palette: Mapping[str, str],
    status_key: str,
    *,
    tooltip: str = "",
) -> MonitorCell:
    fill = color_for(palette, status_key)
    return MonitorCell(
        text=text,
        tooltip=tooltip,
        fill=fill,
        foreground=contrast_foreground(fill),
        palette_key=status_key,
    )


def _pin_cell(view: ExportPinView, palette: Mapping[str, str]) -> MonitorCell:
    fill = color_for(palette, "export_pin_stale") if view.is_stale else None
    return MonitorCell(
        text=view.text,
        tooltip=view.tooltip,
        fill=fill,
        foreground=contrast_foreground(fill) if fill else None,
        palette_key="export_pin_stale" if view.is_stale else None,
    )


def _compare_status_cell(status: AutoMtoCompareStatus) -> MonitorCell:
    return MonitorCell(
        text=status.text or "—",
        tooltip=status.tooltip,
        fill=_revision_fill(status.paint_match),
        bold=status.paint_match is True,
        sort_key=status.kind,
        palette_key=status.kind or None,
    )


def _code_origin_label(origin: str | None) -> str:
    labels = {
        "customer": "Заказчик",
        "pi": "ПИ",
        "unknown": "неизвестно",
    }
    return labels.get(str(origin or "").casefold(), origin or "")


def _displayed_working_revision(pipeline: KitPipelineRow | None) -> str:
    if pipeline is None:
        return ""
    return format_working_rd_rev_label(
        pipeline.working_revision_text,
        pipeline.official_revision_text,
        working_as_build=pipeline.working_as_build,
    )


def _pipeline_review_text(
    pipeline: KitPipelineRow | None,
    *,
    google: GoogleKit | None = None,
    issuance: IssuanceKit | None = None,
) -> str:
    if pipeline is None:
        return "—"
    events = google.events if google is not None else ()
    return pipeline_review_label(
        pipeline, events=events, issuance=issuance, google=google
    )


def _same_calendar_date(left: str, right: str) -> bool:
    first = parse_ddmmyyyy(left)
    second = parse_ddmmyyyy(right)
    if first is not None and second is not None:
        return first == second
    stripped_left = (left or "").strip()
    stripped_right = (right or "").strip()
    return bool(stripped_left) and stripped_left == stripped_right


def _event_revision_text(event: KitEvent) -> str:
    return format_revision(event.revision, event.appendix)


def _event_rev_ok(event: KitEvent, official: str) -> bool:
    event_rev = _event_revision_text(event)
    if official and event_rev:
        return revision_texts_equivalent(event_rev, official)
    return True


def _format_f_event_compact(event: KitEvent) -> str:
    parts = [event.date or "—", event.stage_label or event.stage or "—"]
    rev = _event_revision_text(event)
    if rev:
        parts.append(rev)
    if event.transmittals:
        parts.append(", ".join(event.transmittals))
    return "  ".join(parts)


def _f_event_source_lines(event: KitEvent) -> list[str]:
    lines = [f"источник: {KITS_F_SHEET_LABEL}", _format_f_event_compact(event)]
    raw = (event.raw or "").strip()
    if raw and raw != lines[-1]:
        lines.append(raw)
    return lines


def _issuance_source_lines(issuance: IssuanceKit) -> list[str]:
    send = issuance.send_date or issuance.send_date_sortable or "—"
    incoming = (
        issuance.incoming_control_date
        or issuance.incoming_control_date_sortable
        or "—"
    )
    return [
        f"источник: {ISSUANCE_SHEET_LABEL}",
        f"рев. {issuance.revision_text or '—'} · {issuance.status or '—'}",
        f"отправка {send} · TRM {issuance.send_transmittal or '—'}",
        f"вх.контр. {incoming} · подтв. {issuance.confirm_transmittal or '—'}",
    ]


def _pick_event_on_date(
    events: Sequence[KitEvent],
    date_text: str,
    official: str,
    stages: frozenset[str],
) -> KitEvent | None:
    matched: list[KitEvent] = []
    for event in events:
        if event.stage not in stages:
            continue
        if not event.date or not _same_calendar_date(event.date, date_text):
            continue
        if not _event_rev_ok(event, official):
            continue
        matched.append(event)
    return matched[-1] if matched else None


def _issuance_matches_send_date(
    issuance: IssuanceKit | None,
    send_date: str,
    official: str,
) -> bool:
    if issuance is None or not send_date:
        return False
    if not _same_calendar_date(issuance.send_date, send_date):
        return False
    send_rev = (issuance.revision_text or "").strip()
    if official and send_rev:
        return revision_texts_equivalent(send_rev, official)
    return True


def _issuance_matches_pass_date(
    issuance: IssuanceKit | None,
    pass_date: str,
) -> bool:
    if issuance is None or not pass_date:
        return False
    incoming = (issuance.incoming_control_date or "").strip()
    if incoming and _same_calendar_date(incoming, pass_date):
        return True
    send = (issuance.send_date or "").strip()
    return bool(send) and _same_calendar_date(send, pass_date)


def _f_journal_lines(events: Sequence[KitEvent]) -> list[str]:
    if not events:
        return ["Журнал F: нет строк"]
    lines = ["Журнал F:"]
    for event in events:
        lines.append(f"  {_format_f_event_compact(event)}")
    return lines


def _google_de_agreed(google: GoogleKit | None) -> bool:
    """Return True when D/E ``status_sheet`` looks like «Согласовано»."""

    if google is None:
        return False
    text = (google.status_sheet or "").casefold()
    if "оглас" not in text:
        return False
    return "не соглас" not in text


def _pipeline_face_source_label(source: str) -> str:
    labels = {
        PIPELINE_FACE_ISSUANCE: ISSUANCE_SHEET_LABEL,
        PIPELINE_FACE_DE: KITS_DE_SHEET_LABEL,
        PIPELINE_FACE_F: KITS_F_SHEET_LABEL,
        PIPELINE_FACE_RD: "папка РД (колонка «РД · рев.»)",
    }
    return labels.get(source, source)


def pipeline_review_tooltip(
    pipeline: KitPipelineRow,
    *,
    google: GoogleKit | None = None,
    issuance: IssuanceKit | None = None,
) -> str:
    """Return a multiline review-status tooltip naming Google sheets.

    Args:
        pipeline: Derived ``kit_pipeline`` row.
        google: КСБ ИД kit with parsed F events.
        issuance: Latest effective «Выдача РД ПД» send.

    Returns:
        Plain-text tooltip for Qt, WEB, and the Подсказки pane.
    """

    events = google.events if google is not None else ()
    view = pipeline_review_display(
        pipeline, events=events, issuance=issuance, google=google
    )
    label = pipeline_review_label(
        pipeline, events=events, issuance=issuance, google=google
    )
    official = (pipeline.official_revision_text or "").strip()
    pass_date = view.pass_date
    agreed_date = view.agreed_date
    send_date = view.send_date
    anchor_rev = view.face_revision if view.uses_sheet_face else official
    version = PIPELINE_DISPLAY_VERSION
    if version >= PIPELINE_DISPLAY_V3:
        column_lines = [
            "Эта колонка — статус рассмотрения по Google-таблицам, не по папке РД.",
            (
                "Якорь рев.: последняя эффективная "
                f"{ISSUANCE_SHEET_LABEL}, иначе {KITS_DE_SHEET_LABEL}, "
                f"иначе последняя рев. в {KITS_F_SHEET_LABEL}, иначе папка РД."
            ),
            "Подпись D/E «Согласовано» сама по себе не делает статус «Согласован».",
            "Не рабочая папка.",
            (
                "B/C табличного цикла — здесь. Письмо не этого табличного "
                "цикла — соседняя колонка «Статус согласования»."
            ),
            "Коды B и C не делают статус «Согласован» (нужен A или F «согласовано» на этом цикле).",
        ]
        how_lines = [
            "Как читать подпись:",
            "  статус [(дата ТДО или согласования)][ · B|C] · {выдача|D/E|F|РД} {рев.}[ · отпр. {дата}][ (AB)][ (auto)]",
            "выдача / D/E / F — рев. из таблиц. «РД» в подписи — только когда таблицы совпадают с папкой РД.",
            "Дата в скобках у «Прошел ТДО» — прохождение ТДО / вх.контроля этого табличного цикла.",
            "Дата в скобках у «Согласован» — код A или F «согласовано», не дата отправки.",
            "отпр. — отправка этой же табличной рев.",
            (
                f"Табличный цикл: {_pipeline_face_source_label(view.face_source)} · "
                f"{view.face_revision or '—'}"
            ),
            f"РД на диске: {official or '—'} — колонка «РД · рев.», не этот статус.",
        ]
    elif version == PIPELINE_DISPLAY_V2:
        column_lines = [
            "Эта колонка — этап актуальной рев. в папке РД (колонка «РД · рев.»).",
            "Не столбцы D/E («Google · рев.», «РД Согласовано»).",
            "Не рабочая папка.",
            (
                "Текущие B/C этого цикла — здесь. Письмо не этого цикла — "
                "соседняя колонка «Статус согласования»."
            ),
            "Коды B и C не делают статус «Согласован» (нужен A или F «согласовано» на этом цикле).",
        ]
        how_lines = [
            "Как читать подпись:",
            "  статус [(дата ТДО или согласования)][ · B|C] · РД {рев.}[ · отпр. {дата}][ (AB)][ (auto)]",
            "РД — актуальный пакет в папке РД, то же что «РД · рев.».",
            "Дата в скобках у «Прошел ТДО» — прохождение ТДО / вх.контроля этой РД-рев.",
            "Дата в скобках у «Согласован» — код A или F «согласовано», не дата отправки.",
            "отпр. — отправка этой же РД-рев. (не письма на более новую выдачу).",
            f"РД рев.: {official or '—'}",
        ]
    else:
        column_lines = [
            "Эта колонка — этап актуальной рев. в папке РД (колонка «РД · рев.»).",
            "Не столбцы D/E («Google · рев.», «РД Согласовано»).",
            "Не рабочая папка.",
            "Не последнее письмо A/B/C — это соседняя колонка.",
            "Коды B и C не делают статус «Согласован» (нужен A или F «согласовано» на этом цикле).",
        ]
        how_lines = [
            "Как читать подпись:",
            "  статус [(дата ТДО или согласования)] · РД {рев.}[ · отпр. {дата}][ (AB)][ (auto)]",
            "РД — актуальный пакет в папке РД, то же что «РД · рев.».",
            "Дата в скобках у «Прошел ТДО» — прохождение ТДО / вх.контроля этой РД-рев.",
            "Дата в скобках у «Согласован» — код A или F «согласовано», не дата отправки.",
            "отпр. — отправка этой же РД-рев. (не письма на более новую выдачу).",
            f"РД рев.: {official or '—'}",
        ]
    lines = [label, "", *column_lines, "", *how_lines]
    if view.from_robot_auto:
        lines.append(
            "хвост (auto) — опорная строка F этой колонки заканчивается на auto "
            "(запись каталога)."
        )
    cycle_letter = view.cycle_letter
    if cycle_letter:
        lines.append(
            f"Буква {cycle_letter} — этого цикла, поэтому в этой колонке, "
            "не в «Статус согласования»."
        )
    working = _displayed_working_revision(pipeline)
    if working:
        lines.append(f"Рабочая рев. РД: {working} — не входит в этот статус")
    if (
        _google_de_agreed(google)
        and view.status != "agreed"
        and google is not None
    ):
        d_rev = google.sheet_revision_text or "—"
        d_status = google.status_sheet or "—"
        lines.append("")
        lines.append("Внимание: расхождение с D/E")
        lines.append(f"  источник: {KITS_DE_SHEET_LABEL}")
        lines.append(f"  на листе: {d_status} · рев. {d_rev}")
        lines.append(
            "  рассмотрение не берёт подпись D/E как «Согласован»; "
            f"показан цикл {view.face_revision or official or '—'}."
        )
    if (
        issuance is not None
        and official
        and _revision_text_outranks(issuance.revision_text, official)
    ):
        lines.append("")
        lines.append(
            f"Диск отстаёт от {ISSUANCE_SHEET_LABEL}: "
            f"выдача {issuance.revision_text or '—'}, "
            f"РД на диске {official}."
        )
    elif view.uses_sheet_face and official and view.face_revision:
        lines.append("")
        lines.append(
            "Диск отстаёт от табличного цикла: "
            f"{_pipeline_face_source_label(view.face_source)} "
            f"{view.face_revision}, РД на диске {official}."
        )
    if pass_date:
        lines.append("")
        lines.append(f"Прохождение {pass_date}:")
        pass_event = _pick_event_on_date(
            events, pass_date, anchor_rev, _REVIEW_PASS_STAGES
        )
        if pass_event is not None:
            lines.extend(f"  {item}" for item in _f_event_source_lines(pass_event))
        elif _issuance_matches_pass_date(issuance, pass_date) and issuance is not None:
            lines.extend(f"  {item}" for item in _issuance_source_lines(issuance))
        else:
            lines.append(f"  источник: {KITS_F_SHEET_LABEL} или {ISSUANCE_SHEET_LABEL}")
            lines.append("  строка в кэше не сопоставлена")
    if agreed_date:
        lines.append("")
        lines.append(f"Согласование {agreed_date}:")
        agreed_event = _pick_event_on_date(
            events, agreed_date, anchor_rev, _REVIEW_AGREED_STAGES
        )
        if agreed_event is not None:
            lines.extend(f"  {item}" for item in _f_event_source_lines(agreed_event))
        else:
            lines.append(f"  источник: {KITS_F_SHEET_LABEL}")
            lines.append("  строка кода A / «согласовано» в кэше не сопоставлена")
    if send_date:
        lines.append("")
        lines.append(f"Отправка {send_date}:")
        if _issuance_matches_send_date(issuance, send_date, anchor_rev) and issuance is not None:
            lines.extend(f"  {item}" for item in _issuance_source_lines(issuance))
        else:
            send_event = _pick_event_on_date(
                events, send_date, anchor_rev, _REVIEW_SEND_STAGES
            )
            if send_event is not None:
                lines.extend(
                    f"  {item}" for item in _f_event_source_lines(send_event)
                )
            else:
                lines.append(f"  источник: {KITS_F_SHEET_LABEL}")
                lines.append("  строка отправки не найдена")
    lines.append("")
    if google is None:
        lines.append(f"{KITS_F_SHEET_LABEL}: нет строки комплекта")
        lines.append(f"{KITS_DE_SHEET_LABEL}: нет строки комплекта")
    else:
        lines.extend(_f_journal_lines(events))
        d_rev = google.sheet_revision_text or "—"
        d_status = google.status_sheet or "—"
        lines.append(
            f"{KITS_DE_SHEET_LABEL} (справочно): {d_rev} · {d_status}"
        )
    if issuance is None:
        lines.append(f"{ISSUANCE_SHEET_LABEL}: нет строки")
    elif not send_date or not _issuance_matches_send_date(
        issuance, send_date, anchor_rev
    ):
        lines.append("")
        lines.append("Последняя эффективная выдача:")
        lines.extend(f"  {item}" for item in _issuance_source_lines(issuance))
    return "\n".join(lines)


def pipeline_approval_tooltip(
    pipeline: KitPipelineRow,
    *,
    google: GoogleKit | None = None,
    issuance: IssuanceKit | None = None,
) -> str:
    """Return a multiline approval-letter tooltip naming the F sheet.

    Args:
        pipeline: Derived ``kit_pipeline`` row.
        google: КСБ ИД kit with parsed F events.
        issuance: Latest effective «Выдача РД ПД» send (reference only).

    Returns:
        Plain-text tooltip for Qt, WEB, and the Подсказки pane.
    """

    events = google.events if google is not None else ()
    label = pipeline_approval_label(
        pipeline, google=google, issuance=issuance, events=events
    )
    relation = pipeline_approval_relation(pipeline)
    official = pipeline.official_revision_text or "—"
    shows_letter = pipeline_approval_shows_letter(
        pipeline, google=google, issuance=issuance, events=events
    )
    view = pipeline_review_display(
        pipeline, events=events, issuance=issuance, google=google
    )
    if PIPELINE_DISPLAY_VERSION >= PIPELINE_DISPLAY_V3:
        column_lines = [
            "Эта колонка — письмо A/B/C не этого табличного цикла "
            "(новее диска / прошлый цикл / без рев.).",
            "B/C табличного цикла — в «Статус рассмотрения». "
            "A табличного цикла уже есть в статусе «Согласован».",
        ]
    elif PIPELINE_DISPLAY_VERSION == PIPELINE_DISPLAY_V2:
        column_lines = [
            "Эта колонка — письмо A/B/C не этого цикла "
            "(новее диска / прошлый цикл / без рев.).",
            "Текущие B/C — в «Статус рассмотрения». "
            "Текущий A уже есть в статусе «Согласован».",
        ]
    else:
        column_lines = [
            "Эта колонка — последнее письмо A/B/C во всём журнале F.",
            "Не этап рассмотрения и не подпись D/E.",
        ]
    lines = [
        label,
        "",
        *column_lines,
        f"источник: {KITS_F_SHEET_LABEL}",
        f"{ISSUANCE_SHEET_LABEL} букву A/B/C не задаёт.",
        f"{KITS_DE_SHEET_LABEL} на букву не влияют.",
    ]
    if label.endswith(PIPELINE_F_AUTO_SUFFIX):
        lines.append(
            "хвост (auto) — письмо взято из строки F с окончанием auto "
            "(запись каталога)."
        )
    if not pipeline.code:
        lines.append("В столбце F нет кода A/B/C.")
        if google is not None and google.events:
            lines.append("")
            lines.extend(_f_journal_lines(google.events))
        return "\n".join(lines)
    if not shows_letter:
        letter = (pipeline.code or "").upper()
        if letter in {"B", "C"}:
            lines.append(
                f"Письмо {letter} этого цикла показано в «Статус рассмотрения»."
            )
        else:
            lines.append(
                "Код A этого цикла = «Согласован»; колонку не дублируем."
            )
            lines.append("Дата согласования — в скобках у «Согласован».")
    if relation:
        meaning = {
            APPROVAL_REL_AHEAD: (
                "письмо на старшую рев., чем РД на диске. "
                "Заказчик уже ответил на новую выдачу; диск отстаёт. "
                "Это не «протухшее» письмо."
            ),
            APPROVAL_REL_PREVIOUS: (
                "письмо на младшую рев. или на ту же рев. до новой "
                "отправки/ТДО этого цикла (как 9000-KSB)."
            ),
            APPROVAL_REL_NO_REV: "в строке F не разобрана ревизия.",
            APPROVAL_REL_OTHER: "письмо не привязано к официальному циклу.",
        }.get(relation, "")
        lines.append(f"отношение: {relation} — {meaning}".rstrip(" —"))
    elif view.uses_sheet_face:
        lines.append(
            "отношение: этот табличный цикл — письмо принадлежит "
            f"{view.face_revision or '—'}, не папке РД."
        )
    else:
        lines.append("отношение: этот цикл — письмо принадлежит актуальной РД на диске.")
    lines.append(
        f"РД на диске {official} · письмо "
        f"{pipeline.code_revision_text or '—'} · "
        f"{pipeline.code_date or '—'}"
    )
    if view.uses_sheet_face:
        lines.append(
            f"Табличный цикл: {_pipeline_face_source_label(view.face_source)} · "
            f"{view.face_revision or '—'}"
        )
    letter = None
    for event in events:
        if event.stage in _APPROVAL_CODE_STAGES:
            letter = event
        elif event.stage == "agreed" and letter is None:
            letter = event
    if letter is not None:
        lines.append("")
        lines.append("Последнее письмо:")
        lines.extend(f"  {item}" for item in _f_event_source_lines(letter))
    origin = _code_origin_label(pipeline.code_origin)
    if origin:
        lines.append(f"происхождение: {origin}")
    codes = [
        event
        for event in events
        if event.stage in _APPROVAL_CODE_STAGES or event.stage == "agreed"
    ]
    if codes:
        lines.append("")
        lines.append("Письма F:")
        for event in codes:
            lines.append(f"  {_format_f_event_compact(event)}")
    if google is not None:
        d_rev = google.sheet_revision_text or "—"
        d_status = google.status_sheet or "—"
        lines.append("")
        lines.append(f"{KITS_DE_SHEET_LABEL} (справочно): {d_rev} · {d_status}")
    if issuance is not None:
        lines.append("")
        lines.append(f"{ISSUANCE_SHEET_LABEL} (справочно, букву не задаёт):")
        lines.extend(f"  {item}" for item in _issuance_source_lines(issuance))
    return "\n".join(lines)


def _pipeline_badge_cell(
    text: str,
    hex_color: str | None,
    pipeline: KitPipelineRow | None,
    *,
    kind: str,
    google: GoogleKit | None = None,
    issuance: IssuanceKit | None = None,
) -> MonitorCell:
    tooltip = ""
    if pipeline is not None:
        if kind == "review":
            tooltip = pipeline_review_tooltip(
                pipeline, google=google, issuance=issuance
            )
        else:
            tooltip = pipeline_approval_tooltip(
                pipeline, google=google, issuance=issuance
            )
    if not hex_color:
        return MonitorCell(text=text, tooltip=tooltip)
    return MonitorCell(
        text=text,
        tooltip=tooltip,
        fill=hex_color,
        foreground=contrast_foreground(hex_color),
        bold=True,
    )


def _overlay_current_ids(
    overlay_rows: Sequence[Mapping[str, Any]],
    records_by_id: Mapping[int, FileRecord],
    rd_root: str,
) -> set[int]:
    detected: set[int] = set()
    for row in overlay_rows:
        file_id = int(row["file_id"])
        record = records_by_id.get(file_id)
        if record is None:
            continue
        if record_has_canonical_layout(record, rd_root):
            detected.add(file_id)
    return detected


def _load_google_for_monitor(
    database: CatalogDatabase,
    runtime_dir: str,
) -> tuple[tuple[GoogleKit, ...], tuple[IssuanceKit, ...], bool]:
    try:
        kits = tuple(database.list_google_kits())
        sends = tuple(database.list_issuance_sends())
    except Exception:
        kits = ()
        sends = ()
    if kits or sends:
        try:
            issuance = latest_effective_issuance_kits(database)
        except Exception:
            issuance = ()
        return kits, issuance, True
    cached = load_cached_google_kits(runtime_dir)
    if cached is None:
        return (), (), False
    return cached.kits, cached.issuance_kits, True


def _build_mto_display_rows(
    database: CatalogDatabase,
    records_by_id: Mapping[int, FileRecord],
    overlay_rows: Sequence[Mapping[str, Any]],
    *,
    pending_error: str,
) -> list[dict[str, Any]]:
    rows = database.list_mto_comparisons()
    rows.extend(database.list_robot_extra_mto())
    compared_ids = {
        int(row["rd_file_id"])
        for row in rows
        if row.get("rd_file_id") is not None
    }
    for overlay in overlay_rows:
        file_id = int(overlay["file_id"])
        if (
            overlay.get("file_kind") != FileKind.MTO_XLSX.value
            or file_id in compared_ids
        ):
            continue
        record = records_by_id.get(file_id)
        if record is None:
            continue
        data = record.data
        rows.append(
            {
                "rd_file_id": file_id,
                "robot_file_id": None,
                "status": "blocked",
                "title_system": data.get("title_system"),
                "title": data.get("title"),
                "mark": data.get("mark"),
                "discipline_block": data.get("discipline_block"),
                "rd_path": record.path,
                "rd_path_key": record.path_key,
                "rd_name": data.get("name"),
                "rd_revision": data.get("revision"),
                "rd_appendix": data.get("appendix"),
                "rd_transfer_revision": data.get("transfer_revision"),
                "rd_transfer_appendix": data.get("transfer_appendix"),
                "rd_is_as_build": data.get("transfer_is_as_build"),
                "rd_review_state": record.review_state.value,
                "rd_present": record.present,
                "rd_mtime_ns": data.get("mtime_ns"),
                "robot_path": None,
                "robot_present": False,
                "diff": {
                    "content_status": "not_compared",
                    "issues": ["not_compared"],
                    "error": pending_error,
                },
                "stats": {},
            }
        )
    rows.sort(
        key=lambda row: (
            str(row.get("title_system") or ""),
            str(row.get("discipline_block") or ""),
        )
    )
    return rows


def _excluded_issuance_sends(
    database: CatalogDatabase,
) -> dict[tuple[str, str], tuple[Any, ...]]:
    excluded: dict[tuple[str, str], list[Any]] = {}
    try:
        for review in database.list_issuance_reviews():
            if review.kind != "send":
                continue
            if review.decision not in ISSUANCE_EXCLUDE_TOOLTIP:
                continue
            key = kit_identity_key(review.title, review.mark)
            excluded.setdefault(key, []).append(review)
    except Exception:
        return {}
    return {key: tuple(items) for key, items in excluded.items()}


def excluded_issuance_sends_by_kit(
    database: CatalogDatabase,
) -> dict[tuple[str, str], tuple[Any, ...]]:
    """Map kit identity to excluded issuance-send reviews for tooltips.

    Args:
        database: Initialized catalog database.

    Returns:
        Excluded ``kind=send`` reviews grouped by ``kit_identity_key``.
    """

    return _excluded_issuance_sends(database)


def _hydrate_auto_mto_comparisons(
    *,
    auto_mto_by_kit: Mapping[tuple[str, str], tuple[AutoMtoFile, ...]],
    pins: Mapping[tuple[str, str], ExportPin],
    overlay_by_kit: Mapping[tuple[str, str], tuple[str, str]],
    selections: Mapping[tuple[str, str], ExportSelection],
    runtime_dir: str | Path,
) -> dict[tuple[str, str], tuple[str, AutoMtoCompareResult]]:
    disk = load_auto_mto_compare_cache(runtime_dir)
    result: dict[tuple[str, str], tuple[str, AutoMtoCompareResult]] = {}
    for kit, files in auto_mto_by_kit.items():
        if not files:
            continue
        title, mark = files[0].title, files[0].mark
        path, _rev, _pinned = auto_mto_rd_target(
            title,
            mark,
            pins=pins,
            overlay_by_kit=overlay_by_kit,
            selection=selections.get(kit),
        )
        if not path:
            continue
        rebuilt = _disk_compare_for_path(kit, path, files, disk)
        if rebuilt is not None:
            result[kit] = (path, rebuilt)
    return result


def _disk_compare_for_path(
    kit: tuple[str, str],
    rd_path: str,
    files: Sequence[AutoMtoFile],
    disk: Mapping[str, Mapping[str, Any]],
) -> AutoMtoCompareResult | None:
    for entry in disk.values():
        stored_path = str(entry.get("rd_path") or "")
        entry_kit = kit_identity_key(
            str(entry.get("title") or ""),
            str(entry.get("mark") or ""),
        )
        if entry_kit != kit:
            continue
        if not auto_mto_rd_paths_match(stored_path, rd_path):
            continue
        rebuilt = auto_mto_result_from_entry(entry, files)
        if rebuilt is not None:
            return rebuilt
    return None


def _auto_mto_status_for(
    title: str,
    mark: str,
    *,
    auto_files: Sequence[AutoMtoFile],
    pins: Mapping[tuple[str, str], ExportPin],
    overlay_mto: Mapping[tuple[str, str], tuple[str, str]],
    selection: ExportSelection | None,
    comparison: tuple[str, AutoMtoCompareResult] | None,
) -> tuple[str, str, bool, AutoMtoCompareStatus]:
    path, revision, pinned = auto_mto_rd_target(
        title,
        mark,
        pins=pins,
        overlay_by_kit=overlay_mto,
        selection=selection,
    )
    comparison_rd_path = ""
    comparison_result = None
    if comparison is not None:
        comparison_rd_path, comparison_result = comparison
    status = auto_mto_compare_status(
        files=auto_files,
        rd_path=path,
        rd_revision=revision,
        comparison=comparison_result,
        comparison_rd_path=comparison_rd_path,
        pinned=pinned,
    )
    return path, revision, pinned, status


def build_kits_monitor_row(
    row: KitMatrixRow,
    *,
    pipeline: KitPipelineRow | None,
    palette: Mapping[str, str],
    pins: Mapping[tuple[str, str], ExportPin],
    overlay_mto: Mapping[tuple[str, str], tuple[str, str]],
    selection: ExportSelection | None,
    comparison: tuple[str, AutoMtoCompareResult] | None,
    auto_files: Sequence[AutoMtoFile],
    an_files: Sequence[AnMtoFile],
    mto_flags: MtoKitFlags,
    mto_content_equal: bool | None,
    ifc_by_kit: Mapping[tuple[str, str], str],
    excluded_sends: Sequence[Any],
    sheet_links: SheetLinkContext | None = None,
    robot_mto_accept: RobotMtoAcceptRow | None = None,
    records_by_path_key: Mapping[str, FileRecord] | None = None,
) -> KitsMonitorRow:
    key = kit_identity_key(row.title, row.mark)
    google = row.google
    issuance = row.issuance
    event_date, event_stage, event_rev, event_trm = last_event_parts(google)
    event = google.last_event if google is not None else None
    events = google.events if google is not None else ()
    google_f_text = format_kits_google_f_rev(event, events=events)
    hist = historical_f_mto_event(events, event)
    if hist is not None:
        f_mto_absent = False
        f_mto_text = format_revision(hist.mto_revision, hist.mto_appendix)
        f_mto_historical = True
        f_mto_hist_event: KitEvent | None = hist
    else:
        f_mto_absent = bool(event is not None and event.mto_absent)
        f_mto_text = (
            format_revision(event.mto_revision, event.mto_appendix)
            if event is not None and not f_mto_absent
            else ""
        )
        f_mto_historical = False
        f_mto_hist_event = None
    review_view = (
        pipeline_review_display(
            pipeline, google=google, issuance=issuance
        )
        if pipeline is not None
        else None
    )
    review_text = _pipeline_review_text(
        pipeline, google=google, issuance=issuance
    )
    approval_text = (
        pipeline_approval_label(
            pipeline, google=google, issuance=issuance
        )
        if pipeline is not None
        else "—"
    )
    rd_rev_text = official_rd_rev_text(row, pipeline)
    display_summary = effective_kit_summary(row, pipeline)
    summary_row = (
        row
        if display_summary is row.summary
        else replace(row, summary=display_summary)
    )
    missing_official = official_revision_missing_from_transfer(row, pipeline)
    lag_notes = rd_rev_review_lag_notes(row, pipeline)
    displayed_working = _displayed_working_revision(pipeline)
    working_rev_text = displayed_working or "—"
    path, rd_mto_rev, pinned, compare_status = _auto_mto_status_for(
        row.title,
        row.mark,
        auto_files=auto_files,
        pins=pins,
        overlay_mto=overlay_mto,
        selection=selection,
        comparison=comparison,
    )
    auto_cell = kits_auto_mto_cell(
        row,
        auto_files=auto_files,
        rd_path=path,
        rd_mto_rev=rd_mto_rev,
        pinned=pinned,
    )
    targets = kit_an_targets(
        row,
        auto_files=auto_files,
        rd_path=path,
        rd_mto_rev=rd_mto_rev,
        pipeline=pipeline,
    )
    hit = match_an_to_kit(an_files, targets)
    an_cell = kits_an_cell(hit)
    pin_view = export_pin_view(
        pins.get(key),
        origin=selection.origin if selection is not None else "",
        rule_path=selection.rule_path if selection is not None else "",
    )
    _folder_path, folder_rev = overlay_mto.get(key, ("", ""))
    mto_rev_text = kits_official_folder_mto_text(
        rd_present=row.rd.present,
        revision_text=folder_rev,
    )
    values: dict[str, str] = {
        "Титул": row.title,
        "Марка": row.mark,
        PIN_COLUMN_HEADER: pin_view.text,
        "Выдача · рев.": revision_or_dash(
            issuance is not None,
            issuance.revision_text if issuance else "",
        ),
        "Google · рев.": revision_or_dash(
            google is not None,
            google.sheet_revision_text if google else "",
        ),
        "Робот МТО · рев.": revision_or_dash(
            row.robot.present, row.robot.revision_text
        ),
        "SQ · рев.": revision_or_dash(row.sq.present, row.sq.revision_text),
        "РД · рев.": rd_rev_text,
        "Рабочая рев. РД": working_rev_text,
        "MTO · рев.": mto_rev_text,
        "Авто МТО": auto_cell.text,
        AUTO_MTO_COMPARE_STATUS_HEADER: compare_status.text,
        "АН МТО": an_cell.text,
        "Сводка": summary_label(display_summary),
        "Статус рассмотрения": review_text,
        "Статус согласования": approval_text,
        "Google · статус": (
            google.status_sheet if google and google.status_sheet else "—"
        ),
        "Google · дата F": event_date or "—",
        "Google · этап F": event_stage or "—",
        "Google · рев. F": google_f_text,
        "Google · TRM F": event_trm or "—",
        "Выдача · статус": (
            issuance.status if issuance and issuance.status else "—"
        ),
        "Выдача · дата отпр.": (
            (issuance.send_date_sortable if issuance else "") or "—"
        ),
        "Выдача · TRM отпр.": (
            issuance.send_transmittal
            if issuance and issuance.send_transmittal
            else "—"
        ),
        "Выдача · дата вх.контр.": (
            (issuance.incoming_control_date_sortable if issuance else "")
            or "—"
        ),
        "Выдача · TRM подтв.": (
            issuance.confirm_transmittal
            if issuance and issuance.confirm_transmittal
            else "—"
        ),
        "РД · дата файла": _mtime_clock(row.rd.max_mtime_ns),
    }
    match_flags = kit_revision_match_flags(row)
    origin = kit_robot_origin(row, rd_content_equal=mto_content_equal)
    rd_mto_record = _official_rd_mto_record(
        row,
        overlay_mto=overlay_mto,
        records_by_path_key=records_by_path_key,
    )
    robot_mto_record = (
        mto_xlsx_record(row.robot.paths, records_by_path_key)
        if records_by_path_key
        else None
    )
    accept_state = robot_mto_accept_state(
        robot_mto_accept,
        rd=rd_mto_record,
        robot=robot_mto_record,
    )
    accept_paints_blue = robot_mto_accept_paints_blue(
        accept_state,
        content_equal=mto_content_equal is True or origin.content_equal,
    )
    accept_paints_green = robot_mto_accept_paints_green(accept_state)
    code_a = bool(
        pipeline is not None
        and pipeline_display_code_a(
            pipeline, google=google, issuance=issuance
        )
    )
    review_status = review_view.status if review_view is not None else ""
    ok_reasons = kit_ok_reasons(
        summary_aligned=display_summary is KitSummary.ALIGNED,
        match_flags=match_flags,
        lag_notes=lag_notes,
        review_agreed=review_status == "agreed",
        code_a=code_a,
    )
    kit_ok = not ok_reasons
    ok_cell = kits_ok_cell(kit_ok, ok_reasons)
    cells: dict[str, MonitorCell] = {}
    for header in KITS_HEADERS:
        text = values.get(header, "—")
        cell = MonitorCell(text=str(text or "—"))
        source = KITS_REV_SOURCE_COLUMNS.get(header)
        if source is not None:
            cell = replace(cell, fill=_revision_fill(match_flags.get(source)))
        if header == KITS_OK_HEADER:
            cell = ok_cell
        elif header == PIN_COLUMN_HEADER:
            cell = _pin_cell(pin_view, palette)
        elif header == "Сводка":
            cell = replace(
                cell,
                foreground=SUMMARY_FOREGROUND.get(
                    display_summary, _DEFAULT_FOREGROUND
                ),
                tooltip=summary_tooltip(summary_row),
            )
        elif header == "Статус рассмотрения" and pipeline is not None:
            cell = _pipeline_fill_cell(
                review_text,
                palette,
                review_status or pipeline.status,
                tooltip=pipeline_review_tooltip(
                    pipeline, google=google, issuance=issuance
                ),
            )
        elif header == "Статус согласования" and pipeline is not None:
            color_key = pipeline_approval_color_key(
                pipeline, google=google, issuance=issuance
            )
            tooltip = pipeline_approval_tooltip(
                pipeline, google=google, issuance=issuance
            )
            if color_key:
                cell = _pipeline_fill_cell(
                    approval_text, palette, color_key, tooltip=tooltip
                )
            else:
                cell = replace(cell, tooltip=tooltip)
        elif header == "Выдача · рев." and excluded_sends:
            shown = (
                issuance.revision_text
                if issuance is not None and issuance.revision_text
                else "—"
            )
            parts = [
                f"{item.revision_text or '—'} "
                f"{ISSUANCE_EXCLUDE_TOOLTIP.get(item.decision, item.decision)}"
                for item in excluded_sends
            ]
            cell = _append_tooltip(cell, f"{'; '.join(parts)}; показано {shown}")
        elif header == "Рабочая рев. РД":
            cell = _append_tooltip(cell, KITS_WORKING_REV_TOOLTIP)
            if displayed_working:
                if pipeline is not None and pipeline.working_as_build:
                    cell = _append_tooltip(cell, "As-build.")
                cell = replace(
                    cell,
                    fill=color_for(palette, "working"),
                    foreground=contrast_foreground(
                        color_for(palette, "working")
                    ),
                    palette_key="working",
                )
        elif header == "MTO · рев.":
            cell = _append_tooltip(cell, KITS_MTO_REV_TOOLTIP)
            if cell.text == OFFICIAL_FOLDER_MTO_MISSING:
                fill = color_for(palette, "no_mto")
                cell = replace(
                    cell,
                    fill=fill,
                    foreground=contrast_foreground(fill),
                    palette_key="no_mto",
                    tooltip="\n".join(
                        [
                            KITS_MTO_REV_TOOLTIP,
                            "В официальной папке передачи нет файла MTO.",
                        ]
                    ),
                )
            cell = _apply_f_mto_to_official_mto_cell(
                cell,
                f_mto_text=f_mto_text,
                f_mto_absent=f_mto_absent,
                disk_mto_text=mto_rev_text,
                f_mto_historical=f_mto_historical,
                f_mto_hist_event=f_mto_hist_event,
                f_last_event=event,
                od_rev=event_rev,
            )
        elif header == "Авто МТО":
            cell = auto_cell
        elif header == AUTO_MTO_COMPARE_STATUS_HEADER:
            cell = _compare_status_cell(compare_status)
        elif header == "АН МТО":
            cell = an_cell
        if header == "Google · рев. F":
            cell = _paint_kits_google_f_rev_cell(
                cell,
                event_rev=event_rev,
                rd_rev=row.rd.revision_text if row.rd.present else "",
                od_match=match_flags.get("google_f"),
                f_mto_text=f_mto_text,
                f_mto_absent=f_mto_absent,
                disk_mto_text=mto_rev_text,
                f_mto_historical=f_mto_historical,
                f_mto_hist_event=f_mto_hist_event,
                f_last_event=event,
            )
        if header == "Робот МТО · рев.":
            if origin.matched is True:
                cell = replace(
                    cell,
                    fill=ROBOT_ORIGIN_FILL,
                    tooltip=origin.reason or cell.tooltip,
                )
            elif origin.matched is False:
                cell = replace(
                    cell,
                    fill=ROBOT_ORPHAN_FILL,
                    tooltip=origin.reason or cell.tooltip,
                )
            elif origin.reason:
                cell = _append_tooltip(cell, origin.reason)
            if origin.content_equal:
                cell = replace(cell, bold=True)
            if row.robot.present and mto_content_equal is None:
                cell = _append_tooltip(cell, _MTO_COMPARE_PENDING_TIP)
            if accept_paints_green:
                cell = replace(cell, fill=ROBOT_ORIGIN_FILL)
            if accept_paints_blue and robot_mto_accept is not None:
                cell = replace(
                    cell,
                    foreground=ROBOT_MTO_ACCEPT_FOREGROUND,
                )
                cell = _append_tooltip(
                    cell,
                    robot_mto_accept_tooltip(
                        robot_mto_accept, rd=rd_mto_record
                    ),
                )
        if origin.matched is True:
            if (
                header == "РД · рев."
                and origin.rd
                and not official_revision_missing_from_transfer(row, pipeline)
                and not lag_notes
            ):
                cell = replace(
                    cell,
                    fill=ROBOT_ORIGIN_FILL,
                    tooltip=origin.reason or cell.tooltip,
                )
            elif header == "SQ · рев." and origin.sq:
                cell = replace(
                    cell,
                    fill=ROBOT_ORIGIN_FILL,
                    tooltip=origin.reason or cell.tooltip,
                )
        if header == "РД · рев.":
            missing = official_revision_missing_from_transfer(row, pipeline)
            official = (
                pipeline.official_revision_text
                if pipeline is not None
                else ""
            )
            cell = replace(
                cell,
                tooltip=_official_rd_rev_tooltip(
                    disk_present=row.rd.present,
                    disk_revision=row.rd.revision_text,
                    disk_as_build=row.rd.as_build,
                    official_revision=official,
                    missing=missing,
                    ifc=ifc_by_kit.get(key, ""),
                ),
            )
            if missing:
                cell = _paint_missing_transfer_cell(cell, palette)
            elif row.mixed_title_notes:
                cell = _paint_mixed_titles_cell(cell, palette)
                cell = _append_tooltip(
                    cell, "\n".join(row.mixed_title_notes)
                )
                if origin.matched is True and origin.rd and origin.reason:
                    cell = _append_tooltip(cell, origin.reason)
            elif row.transfer_review_notes or lag_notes:
                fill = REV_DIFF_FILL
                cell = replace(
                    cell,
                    fill=fill,
                    foreground=contrast_foreground(fill),
                )
                if row.transfer_review_notes:
                    cell = _append_tooltip(
                        cell, "\n".join(row.transfer_review_notes)
                    )
                if lag_notes:
                    cell = _append_tooltip(cell, "\n".join(lag_notes))
                if origin.matched is True and origin.rd and origin.reason:
                    cell = _append_tooltip(cell, origin.reason)
            elif origin.matched is True and origin.rd and origin.reason:
                cell = _append_tooltip(cell, origin.reason)
        cells[header] = cell
    if sheet_links is not None:
        for header, href in kits_google_hrefs(
            google=google, issuance=issuance, links=sheet_links
        ).items():
            cell = cells.get(header)
            if cell is None or not href:
                continue
            cells[header] = replace(
                _append_tooltip(cell, GOOGLE_HREF_TIP),
                href=href,
            )
    haystack = join_haystack(
        row.title,
        row.mark,
        row.title_system,
        row.summary.value,
        summary_label(row.summary),
        display_summary.value,
        summary_label(display_summary),
        (
            _pipeline_review_text(pipeline, google=google, issuance=issuance)
            if pipeline is not None
            else None
        ),
        pipeline_approval_label(
            pipeline, google=google, issuance=issuance
        )
        if pipeline
        else None,
        pipeline.status if pipeline else None,
        review_status or None,
        review_view.face_revision if review_view is not None else None,
        pipeline.code if pipeline else None,
        pipeline_approval_relation(pipeline) if pipeline else None,
        "stale устар" if pipeline and pipeline.code_stale else None,
        KITS_F_SHEET_LABEL if pipeline is not None else None,
        KITS_DE_SHEET_LABEL if pipeline is not None else None,
        ISSUANCE_SHEET_LABEL if pipeline is not None else None,
        google.status_sheet if google else None,
        event_date,
        event_stage,
        event_rev,
        google_f_text,
        event_trm,
        last_event_text(google),
        google.comment_raw if google else None,
        issuance.status if issuance else None,
        issuance.send_date_sortable if issuance else None,
        issuance.send_transmittal if issuance else None,
        issuance.incoming_control_date_sortable if issuance else None,
        issuance.confirm_transmittal if issuance else None,
        issuance.note_raw if issuance else None,
        issuance.revision_text if issuance else None,
        row.transfer_review_notes,
        lag_notes,
        row.mixed_title_notes,
        "смешанные титулы" if row.mixed_title_notes else None,
        "смешанные марки" if row.mixed_title_notes else None,
        auto_cell.text,
        compare_status.text,
        an_cell.text,
        "закрывает" if hit.closes_auto_mto else None,
        rd_rev_text,
        displayed_working,
        mto_rev_text,
        ok_cell.text,
        "хороший" if kit_ok else None,
        "правки робота" if accept_state else None,
        "актуально" if accept_state == ACCEPT_LIVE else None,
        "устарело" if accept_state == "stale" else None,
        accept_state or None,
    ).casefold()
    kit_tdo = review_status in KITS_TDO_STATUSES
    as_build = bool(
        (pipeline is not None and pipeline.review_as_build) or row.rd.as_build
    )
    return KitsMonitorRow(
        title=row.title,
        mark=row.mark,
        kit_key=key,
        cells=cells,
        haystack=haystack,
        rd_present=row.rd.present and not missing_official,
        robot_present=row.robot.present,
        has_google_or_issuance=google is not None or issuance is not None,
        summary_aligned=display_summary is KitSummary.ALIGNED,
        kit_ok=kit_ok,
        code_a=code_a,
        kit_tdo_passed=kit_tdo,
        as_build=as_build,
        has_mto_problem=mto_flags.has_problem,
        an_closes_auto_mto=hit.closes_auto_mto,
        pin_view=pin_view,
        matrix_row=row,
        tooltips_text=format_kits_row_tooltips(row.title, row.mark, cells),
    )


def _build_heatmap(
    revision_cells: Sequence[KitRevisionRow],
    *,
    records: Sequence[FileRecord],
    official_ids: set[int],
    rd_root: str | Path,
    skip_dirs: Sequence[str],
    banned_keys: set[tuple[str, str]],
    palette: Mapping[str, str],
    auto_mto_by_kit: Mapping[tuple[str, str], tuple[AutoMtoFile, ...]],
    pins: Mapping[tuple[str, str], ExportPin],
    overlay_mto: Mapping[tuple[str, str], tuple[str, str]],
    selections: Mapping[tuple[str, str], ExportSelection],
    comparisons: Mapping[tuple[str, str], tuple[str, AutoMtoCompareResult]],
    pipeline_by_kit: Mapping[tuple[str, str], KitPipelineRow] | None = None,
) -> HeatmapMonitor:
    allowed = document_tree_kit_identities(
        records, rd_root, skip_dirs, banned_keys
    )
    grouped: dict[tuple[str, str], dict[str, KitRevisionRow]] = {}
    order: list[tuple[str, str]] = []
    display: dict[tuple[str, str], tuple[str, str]] = {}
    for cell in revision_cells:
        key = kit_identity_key(cell.title, cell.mark)
        if key not in allowed:
            continue
        if key not in grouped:
            grouped[key] = {}
            order.append(key)
            display[key] = (cell.title, cell.mark)
        if cell.revision_text:
            grouped[key][cell.revision_text] = cell
    columns = list_revision_columns(
        tuple(cell for cells in grouped.values() for cell in cells.values())
    )
    canonical = [
        record
        for record in records
        if record_has_canonical_layout(record, rd_root)
    ]
    rd_by_kit = aggregate_source_kits(
        canonical,
        source=SourceKind.RD,
        detected_current_ids=official_ids,
    )
    ifc_by_kit = _heatmap_ifc_map(revision_cells)
    rows: list[HeatmapMonitorRow] = []
    for key in order:
        title, mark = display[key]
        cells_by_rev = grouped[key]
        rev_cells: dict[str, MonitorCell] = {}
        for rev in columns:
            rev_cells[rev] = _heatmap_revision_cell(
                cells_by_rev.get(rev), palette
            )
        snapshot = rd_by_kit.get(key)
        rd_cell = _heatmap_rd_rev_cell(
            snapshot,
            ifc_by_kit.get(key, ""),
            pipeline=(pipeline_by_kit or {}).get(key),
            palette=palette,
        )
        selection = selections.get(key)
        export_cell = _heatmap_export_cell(selection, palette)
        path, revision, pinned = auto_mto_rd_target(
            title,
            mark,
            pins=pins,
            overlay_by_kit=overlay_mto,
            selection=selection,
        )
        files = auto_mto_by_kit.get(key, ())
        export_rev = (
            selection.source_revision_text if selection is not None else revision
        )
        export_path = (
            str(selection.source_path)
            if selection is not None and selection.source_path
            else path
        )
        export_pinned = (
            bool(selection is not None and selection.origin in {"pin", "pin_stale"})
            if selection is not None
            else pinned
        )
        auto_cell = _heatmap_auto_mto_cell(
            files, export_rev, export_path, export_pinned, selection
        )
        cached = comparisons.get(key)
        status = auto_mto_compare_status(
            files=files,
            rd_path=export_path,
            rd_revision=export_rev,
            comparison=cached[1] if cached is not None else None,
            comparison_rd_path=cached[0] if cached is not None else "",
            pinned=export_pinned,
        )
        haystack = join_haystack(
            title,
            mark,
            rd_cell.text,
            auto_cell.text,
            status.text,
            *(cell.letters for cell in cells_by_rev.values()),
        ).casefold()
        rows.append(
            HeatmapMonitorRow(
                title=title,
                mark=mark,
                kit_key=key,
                export_cell=export_cell,
                rd_rev=rd_cell,
                auto_mto=auto_cell,
                auto_mto_compare=_compare_status_cell(status),
                cells=rev_cells,
                haystack=haystack,
            )
        )
    return HeatmapMonitor(revision_columns=columns, rows=tuple(rows))


def _heatmap_ifc_map(
    cells: Sequence[KitRevisionRow],
) -> dict[tuple[str, str], str]:
    mapping: dict[tuple[str, str], str] = {}
    for cell in cells:
        if not cell.is_current_ifc:
            continue
        revision = (cell.revision_text or "").strip()
        if not revision:
            continue
        key = kit_identity_key(cell.title, cell.mark)
        previous = mapping.get(key)
        if previous is None or _revision_column_key(revision) > _revision_column_key(
            previous
        ):
            mapping[key] = revision
    return mapping


def _revision_column_key(text: str) -> tuple[int, int, str]:
    revision, appendix = parse_sheet_revision(text)
    return revision_rank(revision, appendix)


def _cell_problems(cell: KitRevisionRow) -> tuple[str, ...]:
    try:
        parsed = json.loads(cell.problem_kinds_json or "[]")
    except json.JSONDecodeError:
        return ()
    return tuple(str(item) for item in parsed if item)


def _heatmap_revision_cell(
    cell: KitRevisionRow | None,
    palette: Mapping[str, str],
) -> MonitorCell:
    if cell is None:
        fill = color_for(palette, "empty")
        return MonitorCell(
            text="",
            tooltip="Нет ревизии",
            fill=fill,
            foreground=contrast_foreground(fill),
            palette_key="empty",
        )
    status = cell.pipeline_status or "empty"
    fill = color_for(palette, status)
    problems = _cell_problems(cell)
    tooltip_lines = [
        f"{cell.title}-{cell.mark}",
        f"рев. {cell.revision_text or '—'}",
        f"статус: {status_color_label(status)}",
    ]
    if cell.letters:
        tooltip_lines.append(f"буквы: {cell.letters}")
    if cell.is_as_build:
        tooltip_lines.append("as-build")
    if cell.is_current:
        tooltip_lines.append(status_color_label("current"))
    if cell.is_current_ifc:
        tooltip_lines.append("последняя IFC (без as-build)")
    tooltip_lines.append("есть MTO" if cell.has_mto else "нет MTO")
    if problems:
        tooltip_lines.append("проблемы: " + ", ".join(problems))
    return MonitorCell(
        text=cell.letters,
        tooltip="\n".join(tooltip_lines),
        fill=fill,
        foreground=(
            color_for(palette, "problem")
            if problems
            else contrast_foreground(fill)
        ),
        bold=bool(problems),
        underline=bool(cell.is_current),
        sort_key=(cell.letters, status),
        palette_key="problem" if problems else status,
    )


def _heatmap_rd_rev_cell(
    snapshot: Any,
    ifc: str,
    pipeline: KitPipelineRow | None = None,
    palette: Mapping[str, str] | None = None,
) -> MonitorCell:
    present = bool(snapshot is not None and snapshot.present)
    revision = ""
    as_build = False
    if snapshot is not None:
        revision = str(getattr(snapshot, "revision_text", "") or "")
        as_build = bool(getattr(snapshot, "as_build", False))
    official = (
        pipeline.official_revision_text if pipeline is not None else ""
    )
    text, missing = format_official_rd_rev_label(
        disk_present=present,
        disk_revision=revision,
        disk_as_build=as_build,
        official_revision=official,
    )
    tooltip = _official_rd_rev_tooltip(
        disk_present=present,
        disk_revision=revision,
        disk_as_build=as_build,
        official_revision=official,
        missing=missing,
        ifc=ifc,
    )
    sort_key = revision or official or text
    cell = MonitorCell(text=text, tooltip=tooltip, sort_key=sort_key)
    if missing and palette is not None:
        return _paint_missing_transfer_cell(cell, palette)
    if missing:
        return replace(cell, foreground=_GREY_FOREGROUND)
    return cell


def _heatmap_export_cell(
    selection: ExportSelection | None,
    palette: Mapping[str, str],
) -> MonitorCell:
    if selection is None:
        return MonitorCell(text="", tooltip="Нет данных экспорта")
    state = selection.state
    text = EXPORT_STATE_TEXT.get(state, state)
    color_key = EXPORT_STATE_COLOR_KEY.get(state, "export_unknown")
    fill = color_for(palette, color_key)
    return MonitorCell(
        text=text,
        tooltip=export_selection_tooltip(selection),
        fill=fill,
        foreground=contrast_foreground(fill),
        palette_key=color_key,
        sort_key=text,
    )


def _heatmap_auto_mto_cell(
    files: Sequence[AutoMtoFile],
    export_rev: str,
    export_path: str,
    pinned: bool,
    selection: ExportSelection | None,
) -> MonitorCell:
    del export_path, pinned
    auto = pick_auto_mto_file(files, export_rev)
    if auto is None or not auto.rd_revision:
        return MonitorCell(
            text="—",
            tooltip=(
                "Нет спецификации MTO в базе заказчика для этого комплекта."
            ),
        )
    text = format_auto_mto_cell_text(files, export_rev)
    match = revision_texts_match(auto.rd_revision, export_rev)
    shown_note = " (совпала с выгрузкой)" if match is True else ""
    count_note = f"; файлов: {len(files)}" if len(files) > 1 else ""
    tips = [
        "Ревизия Авто МТО vs файл, который уходит в папку робота, "
        "не OD текущего состава.",
        f"Показана рев. {auto.rd_revision}{shown_note}{count_note}",
    ]
    if export_rev:
        tips.append(f"к выгрузке: {export_rev}")
    elif selection is None:
        tips.append("к выгрузке: нет выбранного файла")
    tips.append("Файлы заказчика:")
    for file in files:
        sources = ", ".join(file.source_specs) if file.source_specs else file.spec
        mark = " ← показано" if file.relpath == auto.relpath else ""
        tips.append(
            f"  {file.rd_revision}  {Path(file.relpath).name}  "
            f"[{sources}]{mark}"
        )
    path = auto_mto_path(auto)
    if path.is_file():
        tips.append(str(path))
    else:
        tips.append("xlsx ещё не собран (Собрать каталог АвтоМТО)")
    return MonitorCell(
        text=text or "—",
        tooltip="\n".join(tips),
        fill=_revision_fill(match),
        bold=match is True,
        sort_key=auto.rd_revision,
    )


def _worklist_link_tooltip(row: MtoWorklistRow) -> str:
    revision = row.revision_text or "—"
    source = (
        "Источник этапа ревизии и меток — Google F / "
        "«Выдача РД ПД»; данные сопоставлены с ревизией имени "
        f"файла «{revision}». Они не подтверждают физическое "
        "наличие РД или MTO."
    )
    lines = [source]
    if row.has_f_status:
        lines.append(
            "F✓ — для этой ревизии имени файла найдено событие "
            "Google F или отправка в «Выдача РД ПД»."
        )
    else:
        lines.append(
            "F✗ — для этой ревизии имени файла нет события "
            "Google F и отправки в «Выдача РД ПД»."
        )
    if row.package_path:
        lines.append("РД✓ — физически найдена папка пакета:")
        lines.append(f"  {row.package_path}")
    else:
        lines.append(
            "РД✗ — package_path пуст: папка пакета РД физически "
            "не найдена."
        )
    if row.mto_path:
        lines.append("MTO✓ — физически найден MTO-файл:")
        lines.append(f"  {row.mto_path}")
    else:
        lines.append(
            "MTO✗ — mto_path пуст: MTO-файл физически не найден."
        )
    if row.gap_kind == "no_package":
        lines.append(
            "Важно: согласование в Google не доказывает наличие "
            "папки пакета РД."
        )
    return "\n".join(lines)


def _worklist_status_tooltip(row: MtoWorklistRow) -> str:
    status = status_short_label(row.status) if row.status else "не задан"
    gap = GAP_LABELS.get(row.gap_kind, row.gap_kind) or "нет"
    revision = row.revision_text or "—"
    lines = [
        (
            "Источник этапа ревизии и меток — Google F / "
            "«Выдача РД ПД»; данные сопоставлены с ревизией имени "
            f"файла «{revision}». Они не подтверждают физическое "
            "наличие РД или MTO."
        ),
        f"Этап ревизии: {status}.",
        f"Разрыв: {gap}.",
    ]
    if row.gap_kind == "no_package":
        lines.append(
            "Согласование в Google не доказывает наличие папки "
            "пакета РД."
        )
    return "\n".join(lines)


def _build_worklist_row(
    row: MtoWorklistRow,
    *,
    palette: Mapping[str, str],
    pins: Mapping[tuple[str, str], ExportPin],
    overlay_mto: Mapping[tuple[str, str], tuple[str, str]],
    selection: ExportSelection | None,
    comparison: tuple[str, AutoMtoCompareResult] | None,
    auto_files: Sequence[AutoMtoFile],
) -> WorklistMonitorRow:
    key = kit_identity_key(row.title, row.mark)
    package = row.package_label
    if row.package_count > 1:
        package = f"{package} (+{row.package_count - 1})"
    google_cell = worklist_google_cell(row)
    problem_text = ", ".join(
        collision_kind_label(kind) for kind in row.problem_kinds
    )
    pin_view = export_pin_view(
        pins.get(key),
        origin=selection.origin if selection is not None else "",
        rule_path=selection.rule_path if selection is not None else "",
    )
    _path, _rev, _pinned, compare_status = _auto_mto_status_for(
        row.title,
        row.mark,
        auto_files=auto_files,
        pins=pins,
        overlay_mto=overlay_mto,
        selection=selection,
        comparison=comparison,
    )
    problem_fg = color_for(palette, "problem")
    bold = bool(row.problem_kinds or row.gap_kind)
    underline = bool(row.is_current)
    status_fill = color_for(palette, row.status or "empty")
    values: dict[str, MonitorCell] = {
        "Титул": MonitorCell(text=row.title),
        "Марка": MonitorCell(text=row.mark),
        "Ревизия (F/РД)": MonitorCell(text=row.revision_text),
        "Этап ревизии": MonitorCell(
            text=status_short_label(row.status),
            tooltip=_worklist_status_tooltip(row),
            fill=status_fill,
            foreground=contrast_foreground(status_fill),
            palette_key=row.status or "empty",
        ),
        "Метки": MonitorCell(text=row.letters),
        "Связь F → РД → MTO": MonitorCell(
            text=worklist_link_text(row),
            tooltip=_worklist_link_tooltip(row),
            foreground=problem_fg if row.gap_kind else None,
        ),
        "Google": replace(
            google_cell,
            foreground=problem_fg if google_cell.palette_key == "problem" else None,
        ),
        "AB": MonitorCell(text="Да" if row.is_as_build else ""),
        PIN_COLUMN_HEADER: _pin_cell(pin_view, palette),
        AUTO_MTO_COMPARE_STATUS_HEADER: _compare_status_cell(compare_status),
        "Файл MTO": MonitorCell(
            text="Да" if row.mto_path else "Нет",
            foreground=(
                problem_fg if row.gap_kind and not row.mto_path else None
            ),
        ),
        "Дата MTO": MonitorCell(
            text=_mtime_dot(row.mto_mtime_ns),
            sort_key=row.mto_mtime_ns if row.mto_mtime_ns is not None else -1,
        ),
        "Пакет": MonitorCell(
            text=package,
            tooltip=row.package_path,
            foreground=(
                problem_fg if row.gap_kind in {"no_package", "no_rd"} else None
            ),
        ),
        "Чего не хватает": MonitorCell(
            text=GAP_LABELS.get(row.gap_kind, ""),
            foreground=problem_fg if row.gap_kind else None,
        ),
        "Проблемы": MonitorCell(
            text=problem_text,
            tooltip=problem_text if problem_text else "",
        ),
        "Путь MTO": MonitorCell(text=row.mto_path, tooltip=row.mto_path),
    }
    cells = {
        header: replace(cell, bold=bold or cell.bold, underline=underline or cell.underline)
        for header, cell in values.items()
    }
    haystack = join_haystack(
        row.title,
        row.mark,
        row.revision_text,
        row.letters,
        row.status,
        status_short_label(row.status),
        row.mto_path,
        row.package_path,
        row.package_label,
        worklist_link_text(row),
        GAP_LABELS.get(row.gap_kind, row.gap_kind),
        pin_view.text,
        compare_status.text,
        *row.problem_kinds,
    ).casefold()
    return WorklistMonitorRow(
        cells=cells,
        source=row,
        kit_key=key,
        haystack=haystack,
    )


def _yes_no_dash(match: bool | None) -> str:
    if match is True:
        return "да"
    if match is False:
        return "нет"
    return "—"


def _build_an_row(
    file: AnMtoFile,
    *,
    targets: KitAnTargets,
    cache: Mapping[str, Mapping[str, Any]],
    auto_files: Sequence[AutoMtoFile],
    rd_mto_rev: str,
    agreed_score: AnAgreedScore | None = None,
) -> AnMonitorRow:
    del auto_files, rd_mto_rev
    hit = match_an_to_kit((file,), targets)
    cache_key_value = an_content_cache_key(
        file.path,
        file.mtime_ns,
        targets.auto_mto_path,
        counterpart_mtime_ns(targets.auto_mto_path),
        targets.rd_mto_path,
        counterpart_mtime_ns(targets.rd_mto_path),
    )
    entry = cache.get(cache_key_value)
    vs_auto_content = (
        an_result_from_entry(entry.get("vs_auto")) if entry else None
    )
    vs_rd_content = an_result_from_entry(entry.get("vs_rd")) if entry else None
    auto_match = revision_texts_match(file.revision_text, targets.auto_mto)
    rd_match = revision_texts_match(file.revision_text, targets.rd_mto)
    compare_content = not an_is_od(file)
    vs_auto = format_an_vs_cell(
        auto_match,
        vs_auto_content,
        has_counterpart=compare_content and bool(targets.auto_mto_path),
    )
    vs_rd = format_an_vs_cell(
        rd_match,
        vs_rd_content,
        has_counterpart=compare_content and bool(targets.rd_mto_path),
    )
    agreed_text = format_an_agreed_cell(agreed_score)
    agreed_fill = an_agreed_cell_fill(agreed_score)
    row_fill = an_agreed_row_fill(agreed_score)
    kind_text = an_file_kind(file)
    cells = {
        "Титул": MonitorCell(text=file.title),
        "Марка": MonitorCell(text=file.mark),
        "Ревизия АН": MonitorCell(text=file.revision_text or "—"),
        KIND_HEADER: MonitorCell(text=kind_text, sort_key=0 if kind_text == KIND_OD else 1),
        AN_AGREED_HEADER: MonitorCell(
            text=agreed_text,
            tooltip=an_agreed_cell_tooltip(agreed_score) or AN_AGREED_HEADER_TIP,
            fill=agreed_fill,
            bold=bool(agreed_score and agreed_score.is_best),
            sort_key=(
                agreed_score.percent
                if agreed_score is not None and agreed_score.percent is not None
                else -1
            ),
        ),
        "vs Авто МТО": MonitorCell(
            text=vs_auto,
            tooltip=an_vs_cell_tooltip(
                rev_match=auto_match,
                content=vs_auto_content,
                an_path=file.path,
                other_path=targets.auto_mto_path,
                other_label="ПИ / Авто МТО",
            ),
            fill=an_vs_cell_fill(auto_match, vs_auto_content),
            bold=auto_match is True,
        ),
        "vs MTO РД": MonitorCell(
            text=vs_rd,
            tooltip=an_vs_cell_tooltip(
                rev_match=rd_match,
                content=vs_rd_content,
                an_path=file.path,
                other_path=targets.rd_mto_path,
                other_label="MTO РД",
            ),
            fill=an_vs_cell_fill(rd_match, vs_rd_content),
            bold=rd_match is True,
        ),
        "vs Робот": MonitorCell(
            text=_yes_no_dash(
                revision_texts_match(file.revision_text, targets.robot)
            )
        ),
        "vs Выдача": MonitorCell(
            text=_yes_no_dash(
                revision_texts_match(file.revision_text, targets.issuance)
            )
        ),
        "vs F": MonitorCell(
            text=_yes_no_dash(
                revision_texts_match(file.revision_text, targets.google_f)
            )
        ),
        "vs SQ": MonitorCell(
            text=_yes_no_dash(
                revision_texts_match(file.revision_text, targets.sq)
            )
        ),
        "Имя": MonitorCell(text=file.name),
        "Дата": MonitorCell(
            text=_mtime_dot(file.mtime_ns),
            sort_key=file.mtime_ns,
        ),
        "Папка": MonitorCell(text=file.parent_dir),
        "Путь": MonitorCell(text=file.path, tooltip=file.path),
    }
    if row_fill:
        for header in AN_AGREED_ROW_HEADERS:
            if header == AN_AGREED_HEADER:
                continue
            cell = cells.get(header)
            if cell is not None and not cell.fill:
                cells[header] = replace(cell, fill=row_fill)
    haystack = join_haystack(
        file.title,
        file.mark,
        file.revision_text,
        kind_text,
        agreed_text,
        vs_auto,
        vs_rd,
        file.name,
        file.parent_dir,
        file.path,
        "закрывает" if hit.closes_auto_mto else None,
        "лучше" if agreed_score is not None and agreed_score.is_best else None,
    ).casefold()
    return AnMonitorRow(
        cells=cells,
        source=file,
        kit_key=kit_identity_key(file.title, file.mark),
        closes_auto_mto=hit.closes_auto_mto,
        best_agreed=bool(agreed_score and agreed_score.is_best),
        haystack=haystack,
    )


def _build_journal_row(
    row: IssuanceJournalRow,
    *,
    issuance: IssuanceKit | None = None,
    sheet_links: SheetLinkContext | None = None,
) -> JournalMonitorRow:
    cells = {
        "Титул": MonitorCell(text=row.title),
        "Марка": MonitorCell(text=row.mark),
        "Источник": MonitorCell(
            text=JOURNAL_SOURCE_LABELS.get(row.source, row.source)
        ),
        "Рев.": MonitorCell(text=row.revision_text),
        "Дата отпр.": MonitorCell(text=row.send_date),
        "TRM": MonitorCell(text=row.send_transmittal),
        "Статус листа": MonitorCell(text=row.sheet_status),
        "Наш статус": MonitorCell(
            text=JOURNAL_DECISION_LABELS.get(row.decision, row.decision)
        ),
        "Комментарий": MonitorCell(text=row.comment or ""),
        "F": MonitorCell(text="да" if row.in_f else ""),
        "РД": MonitorCell(text="да" if row.in_rd else ""),
        "Робот": MonitorCell(text="да" if row.in_robot else ""),
        "Авто МТО": MonitorCell(text="да" if row.in_auto_mto else ""),
        "Вх.контр.": MonitorCell(text=row.incoming_control_date),
        "TRM подтв.": MonitorCell(text=row.confirm_transmittal),
        "Примечание": MonitorCell(text=row.note),
        "Сопоставление": MonitorCell(
            text=JOURNAL_MATCH_LABELS.get(row.match_state, row.match_state)
        ),
    }
    if sheet_links is not None:
        for header, cell in list(cells.items()):
            href = journal_cell_href(header, row.sheet_row_index, sheet_links)
            if not href:
                continue
            cells[header] = replace(
                _append_tooltip(cell, GOOGLE_HREF_TIP),
                href=href,
            )
    haystack = join_haystack(
        row.title,
        row.mark,
        f"{row.title}-{row.mark}",
        row.revision_text,
        row.send_date,
        row.send_transmittal,
        row.sheet_status,
        row.note,
        row.comment or "",
        row.source,
        JOURNAL_SOURCE_LABELS.get(row.source, ""),
        row.decision,
        JOURNAL_DECISION_LABELS.get(row.decision, ""),
        row.match_state,
        JOURNAL_MATCH_LABELS.get(row.match_state, ""),
    ).casefold()
    return JournalMonitorRow(
        cells=cells,
        source=row,
        kit_key=kit_identity_key(row.title, row.mark),
        haystack=haystack,
        paints_kits_issuance=bool(
            issuance is not None and journal_row_matches_issuance(row, issuance)
        ),
    )


def _compact_diff(row: dict[str, Any]) -> str:
    payload = row.get("diff") or {}
    error = str(payload.get("error") or "")
    if error:
        return error
    details = payload.get("diff") or {}
    stats = row.get("stats") or {}
    parts = [
        f"+{stats.get('added', len(details.get('added') or []))}",
        f"−{stats.get('removed', len(details.get('removed') or []))}",
        f"Δ{stats.get('changed', len(details.get('changed') or []))}",
    ]
    issues = ", ".join(payload.get("issues") or ())
    return " ".join(parts) + (f"; {issues}" if issues else "")


def _build_readiness_row(row: dict[str, Any]) -> MtoReadinessMonitorRow:
    status = str(row.get("status") or "")
    fg = MTO_READINESS_FOREGROUND.get(status.casefold())
    cells = {
        "Статус": MonitorCell(
            text=status.upper(),
            foreground=fg or _DEFAULT_FOREGROUND,
            palette_key=status.casefold() or None,
        ),
        "Титул": MonitorCell(text=_dash(row.get("title"))),
        "Марка": MonitorCell(text=_dash(row.get("mark"))),
        "Документ": MonitorCell(text=_dash(row.get("discipline_block"))),
        "Рев. РД": MonitorCell(
            text=_revision_pair(row.get("rd_revision"), row.get("rd_appendix"))
        ),
        "Рев. робота": MonitorCell(
            text=_revision_pair(
                row.get("robot_revision"), row.get("robot_appendix")
            )
        ),
        "As-built": MonitorCell(text="Да" if row.get("rd_is_as_build") else "Нет"),
        "Проверка": MonitorCell(text=_dash(row.get("rd_review_state"))),
        "Содержимое": MonitorCell(
            text=_dash((row.get("diff") or {}).get("content_status"))
        ),
        "Дата РД": MonitorCell(text=_mtime_clock(row.get("rd_mtime_ns"))),
        "Дата робота": MonitorCell(text=_mtime_clock(row.get("robot_mtime_ns"))),
        "MTO РД": MonitorCell(
            text=_dash(row.get("rd_path")),
            tooltip=str(row.get("rd_path") or "Файл не найден в паре"),
        ),
        "MTO робота": MonitorCell(
            text=_dash(row.get("robot_path")),
            tooltip=str(row.get("robot_path") or "Файл не найден в паре"),
        ),
        "Diff / ошибка": MonitorCell(text=_compact_diff(row)),
    }
    haystack = join_haystack(
        row.get("title_system"),
        row.get("discipline_block"),
        row.get("status"),
        row.get("rd_path"),
        row.get("robot_path"),
        _compact_diff(row),
    ).casefold()
    return MtoReadinessMonitorRow(
        cells=cells,
        source=row,
        kit_key=kit_identity_key(
            str(row.get("title") or ""), str(row.get("mark") or "")
        ),
        haystack=haystack,
    )


def _build_collision_row(row: dict[str, Any]) -> CollisionMonitorRow:
    paths = [str(path) for path in row.get("paths") or ()]
    kind = str(row.get("kind") or "")
    cells = {
        "Область": MonitorCell(text=str(row.get("scope") or "—")),
        "Источник": MonitorCell(text=str(row.get("source") or "—").upper()),
        "Тип": MonitorCell(text=collision_kind_label(kind)),
        "Документ": MonitorCell(text=str(row.get("document_key") or "—")),
        "Сообщение": MonitorCell(text=str(row.get("message") or "—")),
        "Пути": MonitorCell(
            text="\n".join(paths) or "—",
            tooltip="\n".join(paths) or "Путь не определён",
        ),
    }
    haystack = join_haystack(
        row.get("scope"),
        row.get("source"),
        kind,
        collision_kind_label(kind),
        row.get("document_key"),
        row.get("message"),
        *paths,
    ).casefold()
    return CollisionMonitorRow(cells=cells, source=row, haystack=haystack)


def _sorted_card_packages(
    packages: tuple[KitPackageRow, ...],
) -> list[KitPackageRow]:
    def _key(package: KitPackageRow) -> tuple[int, int, str]:
        if package.source in {"rd", "issuance_grey"} or package.is_grey:
            group = 0
        elif package.source == "robot":
            group = 1
        else:
            group = 2
        sequence = package.sequence if package.sequence is not None else -1
        return (group, -sequence, (package.transfer_name or "").casefold())

    return sorted(packages, key=_key)


def _cycles_for_package(package: KitPackageRow, cycles: Sequence[Any]) -> list:
    matched = [
        cycle
        for cycle in cycles
        if package.id is not None and cycle.package_id == package.id
    ]
    unmatched = [
        cycle
        for cycle in cycles
        if cycle.package_id is None
        and (cycle.revision_text or "").casefold()
        == (package.revision_text or "").casefold()
    ]
    seen: set[int | None] = set()
    result = []
    for cycle in (*matched, *unmatched):
        marker = cycle.id if cycle.id is not None else id(cycle)
        if marker in seen:
            continue
        seen.add(marker)
        result.append(cycle)
    return result


def _format_package_issuance(cycles: Sequence[Any], sends_by_id: Mapping) -> str:
    lines: list[str] = []
    for cycle in cycles:
        send = sends_by_id.get(cycle.send_id) if cycle.send_id else None
        if send is None:
            continue
        date = send.send_date_sortable or send.send_date or "—"
        status = send.status or "—"
        trm = send.send_transmittal or "—"
        lines.append(f"{date} · {status} · {trm}")
    return "\n".join(lines) or "—"


def _format_package_review(cycles: Sequence[Any], events_by_id: Mapping) -> str:
    lines: list[str] = []
    for cycle in cycles:
        for event_id in (cycle.tdo_event_id, cycle.code_event_id):
            event = events_by_id.get(event_id) if event_id else None
            if event is None:
                continue
            date = event.date or "—"
            label = event.stage_label or event.stage or "—"
            extra = " ".join(event.transmittals)
            line = f"{date} · {label}"
            if extra:
                line = f"{line} · {extra}"
            lines.append(line)
    return "\n".join(lines) or "—"


def _grey_cell(text: str, *, tooltip: str = "") -> MonitorCell:
    return MonitorCell(
        text=text or "—",
        tooltip=tooltip,
        foreground=_GREY_FOREGROUND,
    )


def _build_card_packages(
    card: KitCard,
    row: KitMatrixRow,
    worklist_rows: Sequence[MtoWorklistRow],
    palette: Mapping[str, str],
    database: CatalogDatabase,
) -> tuple[KitCardPackageMonitor, ...]:
    try:
        sends_with_ids = database.list_issuance_sends_with_ids(
            row.title, row.mark
        )
        events_with_ids = database.list_google_events_with_ids(
            row.title, row.mark
        )
    except Exception:
        sends_with_ids = []
        events_with_ids = []
    sends_by_id = {send_id: send for send_id, send in sends_with_ids}
    events_by_id = {event_id: event for event_id, event in events_with_ids}
    painted: list[KitCardPackageMonitor] = []
    for package in _sorted_card_packages(card.packages):
        cycles = _cycles_for_package(package, card.cycles)
        issuance_text = _format_package_issuance(cycles, sends_by_id)
        review_text = _format_package_review(cycles, events_by_id)
        liquidity = package_liquidity_label(package, card)
        nn = f"{package.sequence:02d}" if package.sequence is not None else "—"
        name = package.transfer_name or Path(package.package_path).name or "—"
        revision = package.revision_text or "—"
        date = _mtime_ymd(package.max_mtime_ns)
        ab_text = "AB" if package.is_as_build else ""
        status = package_review_status(package, row, worklist_rows)
        if package.is_grey:
            mto_cell = _grey_cell(package.mto_revision_text or "—")
            review_cell = _grey_cell(review_text)
            ab_cell = _grey_cell(ab_text)
            liquidity_cell = _grey_cell(liquidity)
            nn_cell = _grey_cell(nn)
            name_cell = _grey_cell(name)
            rev_cell = _grey_cell(revision)
            date_cell = _grey_cell(date)
            issuance_cell = _grey_cell(issuance_text)
        else:
            nn_cell = MonitorCell(text=nn)
            name_cell = MonitorCell(text=name)
            rev_cell = MonitorCell(text=revision)
            date_cell = MonitorCell(text=date)
            issuance_cell = MonitorCell(text=issuance_text)
            if status:
                fill = color_for(palette, status)
                review_cell = MonitorCell(
                    text=review_text,
                    fill=fill,
                    foreground=contrast_foreground(fill),
                    palette_key=status,
                )
            else:
                review_cell = MonitorCell(text=review_text)
            mto_text = (package.mto_revision_text or "").strip()
            if not mto_text:
                fill = color_for(palette, "no_mto")
                mto_cell = MonitorCell(
                    text="—",
                    fill=fill,
                    foreground=contrast_foreground(fill),
                    palette_key="no_mto",
                )
            elif revision_texts_equivalent(
                package.mto_revision_text, package.revision_text
            ):
                mto_cell = MonitorCell(
                    text=package.mto_revision_text or "—",
                    fill=REV_MATCH_FILL,
                )
            else:
                mto_cell = MonitorCell(
                    text=package.mto_revision_text or "—",
                    fill=REV_DIFF_FILL,
                    tooltip=(
                        f"MTO рев. {package.mto_revision_text} "
                        f"при рев. пакета {package.revision_text}"
                    ),
                )
            if package.is_as_build:
                fill = color_for(palette, "us_build")
                ab_cell = MonitorCell(
                    text=ab_text,
                    fill=fill,
                    foreground=contrast_foreground(fill),
                    palette_key="us_build",
                )
            else:
                ab_cell = MonitorCell(text=ab_text)
            if liquidity == "pending":
                fill = color_for(palette, "problem")
                liquidity_cell = MonitorCell(
                    text=liquidity,
                    fill=fill,
                    foreground=contrast_foreground(fill),
                    palette_key="problem",
                )
            else:
                liquidity_cell = MonitorCell(text=liquidity)
        if package.is_current:
            nn_cell = replace(nn_cell, bold=True)
            name_cell = replace(name_cell, bold=True)
            rev_cell = replace(rev_cell, bold=True)
            date_cell = replace(date_cell, bold=True)
            issuance_cell = replace(issuance_cell, bold=True)
            review_cell = replace(review_cell, bold=True)
            mto_cell = replace(mto_cell, bold=True)
            ab_cell = replace(ab_cell, bold=True)
            liquidity_cell = replace(liquidity_cell, bold=True)
        cells = {
            "NN": nn_cell,
            "Пакет": name_cell,
            "Рев.": rev_cell,
            "Дата": date_cell,
            "Выдача": issuance_cell,
            "Рассмотрение": review_cell,
            "MTO": mto_cell,
            "AB": ab_cell,
            "Ликвидность": liquidity_cell,
        }
        painted.append(
            KitCardPackageMonitor(
                nn=nn,
                name=name,
                revision=revision,
                date=date,
                issuance_text=issuance_text,
                review_text=review_text,
                mto_cell=mto_cell,
                ab_text=ab_text,
                liquidity=liquidity,
                is_grey=package.is_grey,
                is_current=package.is_current,
                is_as_build=package.is_as_build,
                package_path=package.package_path,
                review_status=status,
                cells=cells,
            )
        )
    return tuple(painted)
