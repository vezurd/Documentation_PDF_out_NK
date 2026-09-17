"""Russian Qt dashboard for RD versions and robot MTO readiness."""

from __future__ import annotations

import json
import webbrowser
from collections import Counter, defaultdict
from collections.abc import Callable, Iterator, Sequence
from dataclasses import dataclass, replace
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from PySide6.QtCore import (
    QByteArray,
    QDate,
    QSettings,
    Qt,
    QTimer,
    Signal,
    Slot,
)
from PySide6.QtGui import (
    QAction,
    QBrush,
    QCloseEvent,
    QColor,
    QFont,
    QIcon,
    QKeySequence,
    QPainter,
    QPixmap,
    QShortcut,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QStatusBar,
    QStyledItemDelegate,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QToolBar,
    QToolButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.an_index import (
    AnKitHit,
    AnMtoFile,
    KitAnTargets,
    match_an_to_kit,
)
from rd_catalog.an_scan import AnScanProgress, _an_root_disabled
from rd_catalog.an_scan_thread import AnScanThread
from rd_catalog.an_tab import AnTab
from rd_catalog.rd_dump_index import rd_dump_kind
from rd_catalog.rd_dump_scan import (
    RdDumpScanProgress,
    _rd_root_disabled as _rd_dump_root_disabled,
)
from rd_catalog.rd_dump_scan_thread import RdDumpScanThread
from rd_catalog.rd_dump_tab import RdDumpTab
from rd_catalog.customer_pi_auto_mto import (
    AUTO_MTO_COMPARE_STATUS_HEADER,
    AUTO_MTO_COMPARE_STATUS_TOOLTIP,
    AutoMtoCompareResult,
    AutoMtoCompareStatus,
    AutoMtoFile,
    auto_mto_compare_status,
    auto_mto_path,
    format_auto_mto_cell_text,
    list_auto_mto_files_by_kit,
    needs_auto_mto_compare,
    pick_auto_mto_file,
)
from rd_catalog.customer_pi_dialog import CustomerPiDialog
from rd_catalog.ban_filter import BanFilterStore, title_mark_from_row
from rd_catalog.config import CatalogConfig, load_config
from rd_catalog.db import (
    CatalogDatabase,
    KitPackageRow,
    KitPipelineRow,
    apply_mtime_override_to_data,
    parse_override_date,
)
from rd_catalog.doc_bundle import (
    DocumentBundle,
    DocumentTreeLabelOptions,
    ANNULLED_MARKER,
    ANNULLED_TOOLTIP,
    NO_REVISION_LABEL,
    WORKING_MARKER,
    FolderMeanOverride,
    bundle_documents,
    file_revision_label,
    file_save_date_tooltip,
    folder_display_name,
    folder_has_as_build,
    folder_has_mto,
    folder_latest_save_date,
    folder_revision_label,
    folder_revision_rank,
    folder_tree_label,
    folder_tree_sort_key,
    folder_tree_tooltip,
    folder_mean_override,
    format_file_save_date,
    latest_save_mtime_ns,
    override_date_text_from_mtime_ns,
    record_folder_key,
    same_document_path_keys,
    working_folder_tooltip,
)
from rd_catalog.issuance_journal_tab import IssuanceJournalTab
from rd_catalog.issuance_review import (
    JournalAutoMtoHit,
    latest_effective_issuance_kits,
    list_issuance_journal,
)
from rd_catalog.kit_thread import KitLoadThread
from rd_catalog.kits import (
    GoogleKit,
    IssuanceKit,
    KitMatrixRow,
    KitSummary,
    RobotOrigin,
    build_kit_matrix,
    MTO_CATALOG_DATE_ACTION,
    MTO_CATALOG_DATE_FOLDER_ACTION,
    MTO_CATALOG_DATE_MANUAL_ACTION,
    changed_google_kit_keys,
    code_letter_label,
    last_code_letter_for_revision,
    format_revision,
    kit_identity_key,
    mixed_title_open_folders,
    mixed_title_rescan_folders,
    outlook_od_search_query,
    kit_robot_origin,
    mto_content_equal_by_kit,
    summary_label,
    working_folders_from_pipelines,
    annulled_folders_from_pipelines,
)
from rd_catalog.kits_legend_dialog import KitsPaintLegendDialog
from rd_catalog.sheet_de_sync import list_sheet_de_sync_rows
from rd_catalog.sheet_de_sync_dialog import (
    SHEET_DE_SYNC_BUTTON,
    SHEET_DE_SYNC_TITLE,
    SheetDeSyncDialog,
)
from rd_catalog.kits_table_layout import (
    KitsLayoutSaveResult,
    KitsTableLayout,
    canonical_kits_header_name,
    load_default_kits_table_layout,
    merge_header_order,
    save_default_kits_table_layout,
)
from rd_catalog.models import (
    FileKind,
    FileRecord,
    ReviewState,
    ScanProgress,
    SourceKind,
    collision_kind_label,
    make_path_key,
)
from rd_catalog.mto_compare_thread import MtoCompareThread
from rd_catalog.layout_report import (
    LayoutReport,
    build_layout_report,
    format_layout_report_summary,
    layout_report_default_filename,
    write_layout_report_xlsx,
)
from rd_catalog.table_xlsx import (
    KITS_XLSX_BUTTON,
    dated_xlsx_filename,
    write_exported_table_xlsx,
)
from rd_catalog.table_xlsx_qt import snapshot_qtable
from rd_catalog.parse import (
    constructed_kit_rd_mark_folder,
    constructed_kit_rd_title_folder,
    is_transfer_gate_folder_name,
    issued_package_dir,
    parse_transfer_folder,
    record_has_canonical_layout,
    unique_kit_rd_mark_folders,
)
from rd_catalog.context_menu_qt import exec_tracked_menu
from rd_catalog.context_menu_usage import (
    MENU_COLLISION,
    MENU_DOC_TREE,
    MENU_HISTORY,
    MENU_KIT_PACKAGE,
    MENU_KITS,
    MENU_MTO_READINESS,
    configure_context_menu_usage,
    shutdown_context_menu_usage,
)
from rd_catalog.perf_log import (
    configure_perf_log,
    perf_note,
    perf_span,
    set_perf_sink,
    stamp_log_line,
)
from rd_catalog.path_actions import (
    common_parent_dir,
    containing_folder,
    open_containing_folder,
    open_path,
    path_is_under,
)
from rd_catalog.monitor_qt import ROLE_HREF, ROLE_SORT, apply_monitor_cell, _qt_tooltip
from rd_catalog.monitor_views import (
    CARD_PACKAGE_HEADERS,
    KITS_HEADERS,
    KITS_MTO_REV_TOOLTIP,
    KITS_OK_HEADER,
    KITS_OK_TOOLTIP,
    KITS_PAINT_LEGEND_BUTTON,
    KITS_REV_SOURCE_COLUMNS,
    KITS_WORKING_REV_TOOLTIP,
    CatalogMonitor,
    KitCardMonitor,
    KitsMonitorRow,
    MonitorCell,
    MtoKitFlags,
    auto_mto_rd_target,
    build_kits_monitor_row,
    build_mto_kit_flags,
    current_ifc_revision_map,
    effective_kit_summary,
    excluded_issuance_sends_by_kit,
    format_kits_card_mto,
    format_kits_progress_stats,
    format_kits_row_tooltips,
    format_kits_tips_pane,
    format_working_rd_rev_label,
    kits_progress_stats,
    kit_an_targets,
    kits_official_folder_mto_text,
    kits_tooltip_header_offset,
    kits_tooltip_section_start,
    load_catalog_monitor,
    package_liquidity_label,
    paint_kit_card,
    rd_mto_overlay_by_kit,
)
from rd_catalog.mto_worklist_tab import (
    MtoWorklistTab,
    apply_export_pin_view,
)
from rd_catalog.approval_mail_drop_filter import ApprovalMailDropFilter
from rd_catalog.outlook_ole import ensure_drop_hook
from rd_catalog.outlook_search import open_outlook_instant_search
from rd_catalog.approval_mail_preview import (
    comment_lookup_from_kits,
    kit_lookup_from_issuance,
)
from rd_catalog.approval_mail_tab import ApprovalMailTab
from rd_catalog.f_journal import build_journal_patch
from rd_catalog.f_legalize import (
    LEGALIZE_APPROVAL_ACTION,
    LEGALIZE_APPROVAL_STAGE,
    LEGALIZE_APPROVAL_TITLE,
    LEGALIZE_APPROVAL_TOKEN,
    f_line_date_from_mtime_ns,
    legalize_approval_f_line,
    mto_revision_from_records,
)
from rd_catalog.f_legalize_dialog import FLegalizeDialog
from rd_catalog.google_f_write_thread import GoogleFWriteThread
from rd_catalog.google_sheet_links import (
    SheetLinkContext,
    open_google_sheet_url,
    sheet_link_context_from_config,
)
from rd_catalog.google_sheet_links_qt import attach_google_href_clicks
from rd_catalog.mto_export import (
    PIN_COLUMN_HEADER,
    ExportPin,
    ExportPinView,
    ExportSelection,
    export_pin_view,
    load_export_pins,
)
from rd_catalog.pipeline import (
    FolderTreeHint,
    KitCard,
    MtoWorklistRow,
    get_kit_card,
    ingest_google_snapshot,
    kit_keys_under_folders,
    list_folder_tree_hints,
    list_kit_pipelines,
    list_mto_worklist,
    official_detected_current_ids,
    patch_official_detected_current_ids,
    pipeline_algorithm_needs_rebuild,
    pipeline_approval_color_key,
    pipeline_approval_label,
    pipeline_display_review_status,
    pipeline_review_label,
    rebuild_pipeline,
    records_in_contour,
    revision_texts_equivalent,
)
from rd_catalog.revision_matrix_tab import RevisionMatrixTab
from rd_catalog.robot_handoff import (
    format_label_options,
    format_robot_handoff,
    node_kind_label,
    yes_no,
)
from rd_catalog.robot_mto_dialog import RobotMtoSyncDialog
from rd_catalog.robot_mto_sync import (
    RobotMtoSyncError,
    RobotMtoSyncPlan,
    plan_robot_mto_sync,
)
from rd_catalog.robot_mto_thread import RobotMtoSyncThread
from rd_catalog.scan_thread import ScanThread
from rd_catalog.startup_hydrate import (
    StartupHydrateThread,
    StartupSnapshot,
    build_mto_display_rows,
    load_google_for_monitor,
)
from rd_catalog.skip_dirs import SkipDirsStore, path_has_skipped_dir
from rd_catalog.skip_dirs_dialog import SkipDirsDialog
from rd_catalog.sq_to_rd import (
    SqToRdError,
    SqToRdPlan,
    folder_matches_mark,
    plan_sq_to_rd_transfer,
)
from rd_catalog.transfer_review_compare import (
    cache_entry_key,
    labels_from_cache,
    load_transfer_review_compare_cache,
    path_pair_key,
    path_pair_labels_from_cache,
    result_to_entry,
    save_transfer_review_compare_cache,
)
from rd_catalog.transfer_review_compare_thread import (
    TransferReviewCompareOutcome,
    TransferReviewCompareThread,
)
from rd_catalog.sq_to_rd_dialog import SqToRdDialog
from rd_catalog.sq_to_rd_thread import SqToRdThread
from rd_catalog.status_colors import (
    color_for,
    load_status_colors,
    status_color_label,
    status_colors_path,
    status_short_label,
)
from rd_catalog.status_colors_dialog import StatusColorsDialog
from rd_catalog.web_server import WebServerProcess
from rd_catalog_web.urls import DEFAULT_PORT

_ROLE_ROW = Qt.ItemDataRole.UserRole
_ROLE_NODE_KIND = Qt.ItemDataRole.UserRole + 1
_ROLE_PACKAGE = Qt.ItemDataRole.UserRole + 2
_ROLE_HAYSTACK = Qt.ItemDataRole.UserRole + 3
_ROLE_MONITOR = Qt.ItemDataRole.UserRole + 4
_DOC_FILTER_DEBOUNCE_MS = 150
_NODE_TITLE = "title"
_NODE_MARK = "mark"
_NODE_REVISION = "revision"
_SETTINGS_ORGANIZATION = "Documentation_PDF_out_NK"
_SETTINGS_APPLICATION = "rd_catalog"
_WEB_AUTOSTART_KEY = "window/web_server_autostart"
_KITS_DETAIL_TAB_KEY = "window/kits_detail_tab_v2"
_KITS_DETAIL_PLACEMENT_KEY = "window/kits_detail_placement"
_KITS_DETAIL_PLACEMENT_BOTTOM = "bottom"
_KITS_DETAIL_PLACEMENT_RIGHT = "right"
_KITS_SPLITTER_BOTTOM_KEY = "window/kits_splitter"
_KITS_SPLITTER_RIGHT_KEY = "window/kits_splitter_h"
_KITS_TIPS_PLACEHOLDER = (
    "Выберите строку, чтобы прочитать подсказки ячеек "
    "(Сводка, РД · рев. и остальные). "
    "Ctrl+клик по ячейке таблицы прокручивает блок этого столбца "
    "к верху панели. Shift+клик по TRM / ячейке Google открывает лист."
)
_TREE_LABEL_SETTINGS: tuple[tuple[str, str, bool], ...] = (
    ("_doc_show_mto_status", "window/doc_tree_show_mto_status", True),
    ("_doc_show_date", "window/doc_tree_show_date", True),
    ("_doc_show_folder", "window/doc_tree_show_folder", False),
    ("_doc_show_review", "window/doc_tree_show_review", False),
    ("_doc_show_current", "window/doc_tree_show_current", False),
    ("_doc_show_mto", "window/doc_tree_show_mto", False),
    ("_doc_show_working", "window/doc_tree_show_working", False),
    ("_doc_show_as_build", "window/doc_tree_show_as_build", False),
)
_LAYOUT_REPORT_ACTION = "Отчёт о раскладке папок РД…"
_LAYOUT_REPORT_TITLE = "Отчёт о раскладке папок РД"
_LAYOUT_REPORT_NO_RECORDS = (
    "Сначала загрузите файлы РД (скан каталога). "
    "Отчёт строится по уже загруженным записям, без повторного скана."
)
_LAYOUT_REPORT_NO_ROOT = (
    "Корень РД не задан. Без него канонический фильтр выключен, "
    "отчёт о раскладке бессмысленен."
)
_LAYOUT_REPORT_READY_TOOLTIP = (
    "Файлы вне канонической раскладки "
    "(титул / марка / Для передачи / NN_). "
    "MTO в чужой папке молча пропадают из выгрузки."
)
_LAYOUT_REPORT_SAVE_CAPTION = "Сохранить отчёт о раскладке РД"
_LAYOUT_REPORT_SAVE_FILTER = "Excel (*.xlsx)"
_LAYOUT_REPORT_SAVE_QUESTION = "Сохранить книгу Excel?"
_KITS_HEADERS = KITS_HEADERS
_KITS_COL_TITLE = _KITS_HEADERS.index("Титул")
_KITS_COL_OK = _KITS_HEADERS.index(KITS_OK_HEADER)
_KITS_COL_PIN = _KITS_HEADERS.index(PIN_COLUMN_HEADER)
_KITS_COL_ISSUANCE_REV = _KITS_HEADERS.index("Выдача · рев.")
_KITS_COL_GOOGLE_REV = _KITS_HEADERS.index("Google · рев.")
_KITS_COL_ROBOT_REV = _KITS_HEADERS.index("Робот МТО · рев.")
_KITS_COL_SQ_REV = _KITS_HEADERS.index("SQ · рев.")
_KITS_COL_RD_REV = _KITS_HEADERS.index("РД · рев.")
_KITS_COL_WORKING = _KITS_HEADERS.index("Рабочая рев. РД")
_KITS_COL_MTO_REV = _KITS_HEADERS.index("MTO · рев.")
_KITS_COL_AUTO_MTO = _KITS_HEADERS.index("Авто МТО")
_KITS_COL_AUTO_MTO_COMPARE = _KITS_HEADERS.index(AUTO_MTO_COMPARE_STATUS_HEADER)
_KITS_COL_AN = _KITS_HEADERS.index("АН МТО")
_KITS_COL_SUMMARY = _KITS_HEADERS.index("Сводка")
_KITS_COL_PIPELINE = _KITS_HEADERS.index("Статус рассмотрения")
_KITS_COL_APPROVAL = _KITS_HEADERS.index("Статус согласования")
_KITS_COL_GOOGLE_F_REV = _KITS_HEADERS.index("Google · рев. F")
_KITS_REV_SOURCE_COLUMNS = {
    _KITS_HEADERS.index(name): source
    for name, source in KITS_REV_SOURCE_COLUMNS.items()
}
_KITS_MTO_REV_TOOLTIP = KITS_MTO_REV_TOOLTIP
_KITS_WORKING_REV_TOOLTIP = KITS_WORKING_REV_TOOLTIP
_KITS_PKG_STRETCH_COLS = frozenset({1, 4, 5})
_KITS_LAYOUT_BUTTON = "Сохранить шаблон колонок"
_KITS_LAYOUT_OK_FILL = "#C5E8C4"
_KITS_LAYOUT_PROBLEM_FILL = "#F7E8BE"
_KITS_LAYOUT_TOOLTIP = (
    "Запомнить текущий порядок и ширины колонок как шаблон по умолчанию "
    "(runtime и заводской JSON). Зелёная кнопка — оба файла записаны, "
    "жёлтая — смотрите Журнал. После сброса настроек таблица возьмёт этот шаблон."
)
_KITS_XLSX_BUTTON = KITS_XLSX_BUTTON
_KITS_XLSX_TITLE = "Выгрузка Комплекты"
_KITS_XLSX_TOOLTIP = (
    "Видимые строки текущей сортировки: порядок и ширины колонок как на экране, "
    "заливка как в таблице. Скрытые фильтром строки не попадают в файл."
)
_KITS_XLSX_SAVE_CAPTION = "Сохранить таблицу Комплекты"
_KITS_XLSX_SAVE_FILTER = "Excel (*.xlsx)"
_KITS_LEGEND_TOOLTIP = (
    "Цвета и жирный шрифт столбцов MTO и ревизий на Комплектах, "
    "с примерами как в таблице."
)
_KITS_DE_SYNC_TOOLTIP = (
    "Комплекты, где последнее событие F совпадает с ревизией РД, "
    "а столбцы D/E ещё нет. Запись тем же движком, что письма о согласовании."
)


class _KitsSortItem(QTableWidgetItem):
    """Prefer ``MonitorCell.sort_key`` so «Ок» sorts «да» before «нет»."""

    def __lt__(self, other: QTableWidgetItem) -> bool:
        left = self.data(ROLE_SORT)
        right = other.data(ROLE_SORT) if other is not None else None
        if left is not None and right is not None:
            try:
                return left < right
            except TypeError:
                pass
        return super().__lt__(other)


def capture_kits_header_layout(
    header: QHeaderView,
    names: Sequence[str],
) -> KitsTableLayout:
    """Return visual order and section widths for *header*.

    Args:
        header: Horizontal header of the Комплекты table.
        names: Logical header titles (same order as columns).

    Returns:
        Layout with every current column, left to right.

    Raises:
        ValueError: If *names* length does not match the header.
    """

    if header.count() != len(names):
        raise ValueError(
            f"kits header {header.count()} != names {len(names)}"
        )
    order: list[str] = []
    widths: dict[str, int] = {}
    for visual in range(header.count()):
        logical = header.logicalIndex(visual)
        name = str(names[logical])
        order.append(name)
        width = header.sectionSize(logical)
        if width > 0:
            widths[name] = width
    return KitsTableLayout(order=tuple(order), widths=widths)


def apply_kits_header_layout(
    header: QHeaderView,
    names: Sequence[str],
    layout: KitsTableLayout,
) -> None:
    """Move and resize header sections from a name-based template.

    Unknown template names are ignored. Current columns missing from the
    template stay after the known ones, in logical order.

    Args:
        header: Horizontal header of the Комплекты table.
        names: Logical header titles (same order as columns).
        layout: Saved visual order and optional default widths.
    """

    if header.count() != len(names):
        return
    name_to_logical = {str(name): index for index, name in enumerate(names)}
    merged = merge_header_order(layout.order, names)
    logicals = [name_to_logical[name] for name in merged]
    for target, logical in enumerate(logicals):
        current = header.visualIndex(logical)
        if current != target:
            header.moveSection(current, target)
    for name, width in layout.widths.items():
        logical = name_to_logical.get(canonical_kits_header_name(name))
        if logical is None or width <= 0:
            continue
        header.resizeSection(logical, width)


_KITS_CARD_LEGEND_KEYS = (
    "code_a",
    "tdo_review",
    "no_mto",
    "us_build",
    "problem",
)
_REV_MATCH_BG = QColor("#E2F2E1")
_REV_DIFF_BG = QColor("#F7E8BE")
_ROBOT_ORIGIN_BG = QColor("#C5E8C4")
_ROBOT_ORPHAN_BG = QColor("#F3C2C2")
_LAYOUT_MISMATCH_BG = _REV_DIFF_BG
_LAYOUT_MISMATCH_LEAF_TIP = (
    "Путь не в формате «титул / марка / Для передачи / NN_». "
    "Переложите в правильную структуру или удалите."
)
_LAYOUT_MISMATCH_PARENT_TIP = (
    "Есть вложенные файлы вне формата «титул / марка / Для передачи / NN_»."
)
_LATEST_FILE_DATE_BG = QColor("#DCEEFF")
_HISTORY_HEADERS = (
    "Документ",
    "Тип",
    "PDF",
    "Ред",
    "Текущий",
    "Наличие",
    "Имя",
    "Дата",
    "Рев.",
    "AB",
    "Рабочая",
    "Аннул.",
    PIN_COLUMN_HEADER,
    "Статус MTO",
    "Путь комплекта",
)
_HISTORY_COL_NAME = 6
_HISTORY_COL_DATE = 7
_HISTORY_COL_WORKING = 10
_HISTORY_COL_ANNULLED = 11
_HISTORY_COL_PIN = 12
_HISTORY_COL_MTO_STATUS = 13
_HISTORY_COL_PACKAGE = 14
_HISTORY_DATE_COLUMN_TIP = (
    "Дата сохранения файла (mtime скана). "
    "Самая свежая в выбранной папке — бледно-голубым."
)
_MTO_STATUS_COLUMN_TIP = (
    "Цикл Google F этой ревизии имени файла, не пометка папки. "
    "Рабочая папка с кодом A остаётся «код A»."
)
_DEFERRED_HEATMAP = "heatmap"
_DEFERRED_WORKLIST = "worklist"
_DEFERRED_MTO = "mto"
_DEFERRED_ISSUANCE_JOURNAL = "issuance_journal"
_DEFERRED_COLLISIONS = "collisions"
_DEFERRED_APPROVAL = "approval"
_DEFERRED_TREE = "tree"
_DEFERRED_AN = "an"
_DEFERRED_RD_DUMP = "rd_dump"
_DEFERRED_ALL = frozenset(
    {
        _DEFERRED_HEATMAP,
        _DEFERRED_WORKLIST,
        _DEFERRED_MTO,
        _DEFERRED_ISSUANCE_JOURNAL,
        _DEFERRED_COLLISIONS,
        _DEFERRED_AN,
        _DEFERRED_RD_DUMP,
        _DEFERRED_APPROVAL,
        _DEFERRED_TREE,
    }
)


def _revision(revision: Any, appendix: Any) -> str:
    value = str(revision or "—")
    return f"{value} AN{appendix}" if appendix else value


def _mtime(value: Any) -> str:
    try:
        return datetime.fromtimestamp(int(value) / 1_000_000_000).strftime(
            "%Y-%m-%d %H:%M"
        )
    except (TypeError, ValueError, OSError):
        return "—"


def _issuance_date_from_mtime_ns(value: int | None) -> str:
    """Format a file mtime as the issuance ``DD.MM.YYYY`` cell.

    Args:
        value: Nanoseconds since epoch, or ``None``.

    Returns:
        Local calendar date, or empty when unknown.
    """

    if not value:
        return ""
    try:
        stamp = datetime.fromtimestamp(int(value) / 1_000_000_000)
    except (TypeError, ValueError, OSError):
        return ""
    return f"{stamp.day:02d}.{stamp.month:02d}.{stamp.year:04d}"


def scroll_kits_tips_section_to_top(edit: QPlainTextEdit, char_pos: int) -> int:
    """Scroll the Подсказки pane so the ``===`` block at *char_pos* is on top.

    Args:
        edit: Read-only tips ``QPlainTextEdit``.
        char_pos: Character offset in ``edit.toPlainText()``.

    Returns:
        Character offset of the block header (or 0).
    """

    start = kits_tooltip_section_start(edit.toPlainText(), char_pos)
    cursor = edit.textCursor()
    cursor.setPosition(start)
    edit.setTextCursor(cursor)
    block = edit.document().findBlock(start)
    edit.verticalScrollBar().setValue(block.blockNumber())
    return start


def scroll_kits_tips_header_to_top(edit: QPlainTextEdit, header: str) -> int | None:
    """Scroll Подсказки so the ``=== header ===`` block of a column is on top.

    Args:
        edit: Read-only tips ``QPlainTextEdit``.
        header: Комплекты column title.

    Returns:
        Character offset of the block header, or None when that column has
        no tooltip block.
    """

    start = kits_tooltip_header_offset(edit.toPlainText(), header)
    if start is None:
        return None
    return scroll_kits_tips_section_to_top(edit, start)


class _ReadOnlyCopyDelegate(QStyledItemDelegate):
    """Read-only line editor so double-click can select and copy cell text."""

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setReadOnly(True)
        editor.setFrame(False)
        return editor

    def setEditorData(self, editor, index) -> None:
        if not isinstance(editor, QLineEdit):
            super().setEditorData(editor, index)
            return
        editor.setText(str(index.data(Qt.ItemDataRole.DisplayRole) or ""))
        editor.selectAll()

    def setModelData(self, editor, model, index) -> None:
        return


def _enable_readonly_cell_copy(table: QTableWidget) -> None:
    """Double-click selects cell text; the value itself stays read-only."""

    table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)
    table.setItemDelegate(_ReadOnlyCopyDelegate(table))


class _ClickCopyLabel(QLabel):
    """Label that copies its visible text on left click."""

    copied = Signal(str)

    def mouseReleaseEvent(self, event) -> None:
        if event is not None and event.button() == Qt.MouseButton.LeftButton:
            text = self.text().strip()
            if text:
                QApplication.clipboard().setText(text)
                self.copied.emit(text)
        super().mouseReleaseEvent(event)


def _paint_revision_cell(item: QTableWidgetItem, matches: bool | None) -> None:
    """Tint a revision cell: soft green match, warm yellow mismatch."""

    if matches is True:
        item.setBackground(QBrush(_REV_MATCH_BG))
    elif matches is False:
        item.setBackground(QBrush(_REV_DIFF_BG))


def _paint_robot_origin_cell(
    item: QTableWidgetItem,
    *,
    matched: bool,
    hint: str,
) -> None:
    """Override revision tint: richer green when robot origin matches, soft red if not."""

    item.setBackground(QBrush(_ROBOT_ORIGIN_BG if matched else _ROBOT_ORPHAN_BG))
    if hint:
        item.setToolTip(hint)


def _paint_pipeline_cell(item: QTableWidgetItem, hex_color: str) -> None:
    """Tint only the pipeline-status cell from the user palette."""

    color = QColor(hex_color)
    if not color.isValid():
        return
    item.setBackground(QBrush(color))
    item.setForeground(
        QBrush(QColor("#ffffff") if color.lightness() < 140 else QColor("#202124"))
    )


def _paint_tree_working(item: QTreeWidgetItem, hex_color: str) -> None:
    """Tint a documents-tree node that is a working issued folder."""

    color = QColor(hex_color)
    if not color.isValid():
        return
    item.setBackground(0, QBrush(color))


def _paint_layout_mismatch(item: QTreeWidgetItem, *, nested: bool) -> None:
    """Tint a documents-tree node whose files are outside title/mark/gate/NN_."""

    item.setBackground(0, QBrush(_LAYOUT_MISMATCH_BG))
    item.setToolTip(
        0, _LAYOUT_MISMATCH_PARENT_TIP if nested else _LAYOUT_MISMATCH_LEAF_TIP
    )


def _mto_node_label(hint: FolderTreeHint | None) -> str:
    """Return the compact MTO badge text for a documents-tree node."""

    if hint is None:
        return ""
    if not hint.mto_status:
        return "нет MTO" if not hint.has_mto_file else ""
    label = status_short_label(hint.mto_status)
    if hint.mto_status != "no_mto" and not hint.has_mto_file:
        label = f"{label} (нет файла)"
    return label


def _settings_bool(settings: QSettings, key: str, default: bool) -> bool:
    """Read a QSettings flag that may be stored as bool, int, or text."""

    value = settings.value(key, default)
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().casefold()
    if text in {"1", "true", "yes", "on"}:
        return True
    if text in {"0", "false", "no", "off", ""}:
        return False
    return default


@dataclass(frozen=True, slots=True)
class _ScanViewRestore:
    """Tab and kit/tree selection to restore after a scoped scan."""

    tab_index: int
    title: str = ""
    mark: str = ""
    transfer_name: str | None = None
    stay_on_documents: bool = False


@dataclass(frozen=True, slots=True)
class _MtimeOverrideMenu:
    """Context-menu actions that write ``file_mtime_override``."""

    code: QAction | None = None
    manual: QAction | None = None
    folder: QAction | None = None
    clear: QAction | None = None


def _style_kits_status_badge(label: QLabel, text: str, hex_color: str | None) -> None:
    """Apply bold badge styling when ``hex_color`` is set; clear otherwise."""

    label.setText(text)
    if hex_color:
        label.setStyleSheet(
            f"font-weight: bold; padding: 4px 10px; "
            f"background: {hex_color}; color: "
            f"{'#ffffff' if QColor(hex_color).lightness() < 140 else '#202124'};"
        )
    else:
        label.setStyleSheet("")


def _collision_kind_label(kind: str) -> str:
    """Return a short Russian label for a stored collision kind."""

    return collision_kind_label(kind)


def _record_name(record: FileRecord) -> str:
    value = record.data.get("name")
    if value not in (None, ""):
        return str(value)
    return Path(record.path).name


def _file_extension(record: FileRecord) -> str:
    raw = record.data.get("extension")
    if raw not in (None, ""):
        return str(raw).lstrip(".").casefold()
    return Path(record.path).suffix.lstrip(".").casefold()


def _bundle_records(bundle: DocumentBundle) -> tuple[FileRecord, ...]:
    files: list[FileRecord] = []
    if bundle.pdf is not None:
        files.append(bundle.pdf)
    files.extend(bundle.editables)
    return tuple(files)


def _bundle_mto_xlsx(bundle: DocumentBundle) -> FileRecord | None:
    """Return the MTO workbook on a history row, if any."""

    for record in _bundle_records(bundle):
        if str(record.data.get("file_kind") or "") == FileKind.MTO_XLSX.value:
            return record
    return None


def _bundle_haystack(bundle: DocumentBundle) -> str:
    parts = [
        bundle.title,
        bundle.mark,
        f"{bundle.title}-{bundle.mark}",
        bundle.revision_label,
        bundle.core_stem,
    ]
    for record in _bundle_records(bundle):
        parts.append(record.path)
        parts.append(_record_name(record))
    return " ".join(parts).casefold()


def _revision_node_haystack(
    bundles: list[DocumentBundle],
    revision_label: str,
    folder_name: str,
    review_status: str,
    *,
    is_working: bool = False,
    is_annulled: bool = False,
) -> str:
    """Return the casefolded search text stored on a revision tree node."""

    parts = [
        *(_bundle_haystack(bundle) for bundle in bundles),
        revision_label,
        folder_name,
        review_status,
    ]
    if is_working:
        parts.append(WORKING_MARKER)
    if is_annulled:
        parts.append(ANNULLED_MARKER)
    return " ".join(parts).casefold()


def _hint_folder_marks(hint) -> tuple[bool, bool]:
    """Return ``(is_working, is_annulled)``; annulled wins if both are set."""

    is_annulled = bool(hint is not None and getattr(hint, "is_annulled", False))
    is_working = bool(hint is not None and hint.is_working and not is_annulled)
    return is_working, is_annulled


def _tree_item_or_ancestor_hidden(item: QTreeWidgetItem | None) -> bool:
    """Return whether the item or any parent is hidden by the title filter."""

    cursor = item
    while cursor is not None:
        if cursor.isHidden():
            return True
        cursor = cursor.parent()
    return False


def _bundle_paths(bundle: DocumentBundle) -> list[str]:
    return [record.path for record in _bundle_records(bundle) if record.path]


def _bundle_package_path(bundle: DocumentBundle) -> str:
    """Return the issued package folder for a documents-table row.

    Args:
        bundle: PDF plus paired editables in one transfer folder.

    Returns:
        Package directory from ``issued_package_dir`` (no ``PDF``/``DWG``
        leaf and no filename), or an empty string when no path is known.
    """

    for record in _bundle_records(bundle):
        package = issued_package_dir(record.path)
        if package:
            return package
    return ""


def _ask_catalog_override_date(
    parent: QWidget,
    *,
    default_date: date,
    disk_text: str,
) -> str | None:
    """Ask for a free catalog date ``DD.MM.YYYY``.

    Args:
        parent: Qt parent.
        default_date: Initial calendar value.
        disk_text: Disk mtime shown in the hint.

    Returns:
        ``DD.MM.YYYY``, or ``None`` when cancelled.
    """

    dialog = QDialog(parent)
    dialog.setWindowTitle("Дата MTO")
    layout = QVBoxLayout(dialog)
    hint = QLabel(
        "Файл на диске не меняется. Дата каталога используется "
        "в дереве, таблице документов и «Проверить передачи».\n"
        f"На диске: {disk_text}",
        dialog,
    )
    hint.setWordWrap(True)
    layout.addWidget(hint)
    picker = QDateEdit(dialog)
    picker.setCalendarPopup(True)
    picker.setDisplayFormat("dd.MM.yyyy")
    picker.setDate(QDate(default_date.year, default_date.month, default_date.day))
    picker.setMinimumDate(QDate(1990, 1, 1))
    picker.setMaximumDate(QDate.currentDate().addYears(2))
    layout.addWidget(picker)
    buttons = QDialogButtonBox(
        QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
        dialog,
    )
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)
    if dialog.exec() != QDialog.DialogCode.Accepted:
        return None
    chosen = picker.date()
    return f"{chosen.day():02d}.{chosen.month():02d}.{chosen.year():04d}"


def _bundle_open_record(bundle: DocumentBundle) -> FileRecord | None:
    if bundle.pdf is not None:
        return bundle.pdf
    if bundle.editables:
        return bundle.editables[0]
    return None


def _editable_cell_text(bundle: DocumentBundle) -> str:
    if not bundle.editables:
        return "Нет"
    extensions: list[str] = []
    seen: set[str] = set()
    for editable in bundle.editables:
        ext = _file_extension(editable)
        if ext and ext not in seen:
            seen.add(ext)
            extensions.append(ext)
    if extensions:
        return f"Да ({', '.join(extensions)})"
    return "Да"


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


class CatalogWindow(QMainWindow):
    """Display persisted catalog state and orchestrate one background scan."""

    def __init__(
        self,
        config: CatalogConfig | None = None,
        database: CatalogDatabase | None = None,
        parent: QWidget | None = None,
    ) -> None:
        """Build the window without scanning or initializing production storage.

        Args:
            config: Optional injected configuration, useful for local smoke tests.
            database: Optional initialized database repository.
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        ensure_drop_hook()
        self.config = config or load_config()
        self.database = database or CatalogDatabase(self.config.db_path)
        self._ban_store = BanFilterStore.from_runtime_dir(self.config.runtime_dir)
        self._skip_store = SkipDirsStore.from_runtime_dir(self.config.runtime_dir)
        self.config = replace(self.config, skip_dirs=self._skip_store.tokens())
        configure_perf_log(self.config.runtime_dir)
        configure_context_menu_usage(self.config.runtime_dir)
        self._settings = QSettings(
            _SETTINGS_ORGANIZATION, _SETTINGS_APPLICATION
        )
        self._scan_thread: ScanThread | None = None
        self._an_scan_thread: AnScanThread | None = None
        self._rd_dump_scan_thread: RdDumpScanThread | None = None
        self._startup_thread: StartupHydrateThread | None = None
        self._google_thread: KitLoadThread | None = None
        self._google_write_thread: GoogleFWriteThread | None = None
        self._robot_sync_thread: RobotMtoSyncThread | None = None
        self._sq_to_rd_thread: SqToRdThread | None = None
        self._mto_compare_thread: MtoCompareThread | None = None
        self._mto_pending_keys: set[tuple[str, str]] = set()
        self._mto_queued: (
            tuple[frozenset[tuple[str, str]] | None, tuple[tuple[str, str], ...]] | None
        ) = None
        self._mto_progress: tuple[int, int] = (0, 0)
        self._mto_resume_on_idle = False
        self._scan_restore: _ScanViewRestore | None = None
        self._all_records: list[FileRecord] = []
        self._contour_records: tuple[FileRecord, ...] | None = None
        self._issuance_kits_fresh = False
        self._official_ids_token: tuple[int, frozenset[int], int] | None = None
        self._official_ids_value: set[int] = set()
        self._tree_kit_identities_cache: (
            tuple[tuple[object, ...], set[tuple[str, str]]] | None
        ) = None
        self._contour_ids_cache: tuple[int, set[int]] | None = None
        self._record_by_id: dict[int, FileRecord] = {}
        self._records_by_path_key: dict[str, FileRecord] = {}
        self._records_by_kit: dict[tuple[str, str], list[FileRecord]] = {}
        self._catalog_index_token: int | None = None
        self._official_rd_package_cache: (
            tuple[Any, ...] | None
        ) = None
        self._official_rd_package_by_kit: dict[
            tuple[str, str], KitPackageRow | None
        ] = {}
        self._kits_last_card: tuple[str, str, Any] | None = None
        self._mto_rows: list[dict[str, Any]] = []
        self._collision_rows: list[dict[str, Any]] = []
        self._detected_current_ids: set[int] = set()
        self._google_kits: tuple[GoogleKit, ...] = ()
        self._issuance_kits: tuple[IssuanceKit, ...] = ()
        self._issuance_sends: tuple[IssuanceKit, ...] = ()
        self._sheet_links: SheetLinkContext | None = None
        self._google_source: str = ""
        self._google_fetched_at: str = ""
        self._google_warning: str | None = None
        self._google_loaded = False
        self._kit_rows: tuple[KitMatrixRow, ...] = ()
        self._kit_pipelines: dict[tuple[str, str], KitPipelineRow] = {}
        self._mto_worklist_rows: tuple[MtoWorklistRow, ...] = ()
        self._mto_kit_flags: dict[tuple[str, str], MtoKitFlags] = {}
        self._status_colors_path = status_colors_path(self.config.runtime_dir)
        self._status_colors = load_status_colors(self._status_colors_path)
        self._mto_icon_cache: dict[str, QIcon] = {}
        self._mto_content_by_kit: dict[tuple[str, str], bool] = {}
        self._transfer_mto_results: dict[str, str] = labels_from_cache(
            self.config.runtime_dir
        )
        self._transfer_mto_by_pair = path_pair_labels_from_cache(
            self.config.runtime_dir
        )
        self._transfer_mto_entries = load_transfer_review_compare_cache(
            self.config.runtime_dir
        )
        self._transfer_mto_thread: TransferReviewCompareThread | None = None
        self._transfer_mto_paused = False
        self._doc_bundles: list[DocumentBundle] = []
        self._doc_bundles_token: tuple[Any, ...] | None = None
        self._folder_hints: dict[tuple[str, str, str], FolderTreeHint] = {}
        self._export_pins_by_key: dict[tuple[str, str], ExportPin] = {}
        self._auto_mto_by_kit: dict[tuple[str, str], tuple[AutoMtoFile, ...]] = {}
        self._an_files_by_kit: dict[tuple[str, str], tuple[AnMtoFile, ...]] = {}
        self._rd_dump_files: tuple[AnMtoFile, ...] = ()
        self._catalog_monitor: CatalogMonitor | None = None
        self._customer_pi_dialog: CustomerPiDialog | None = None
        self._startup_load_pending = False
        self._startup_load_running = False
        self._secondary_tabs_pending = False
        self._deferred_widgets: set[str] = set()
        self._startup_timer: QTimer | None = None
        self._secondary_timer: QTimer | None = None

        self.setWindowTitle("Каталог РД · AGCC")
        self.setMinimumSize(1050, 700)
        self.resize(1480, 900)
        self._build_toolbar()
        self._tabs = QTabWidget(self)
        self._kits_tab = self._build_kits_tab()
        self._tabs.addTab(self._kits_tab, "Комплекты")
        self._revision_matrix_tab = RevisionMatrixTab(self)
        self._revision_matrix_tab.bind_catalog(self.config, self.database)
        self._revision_matrix_tab.kit_activated.connect(self._on_revision_matrix_kit)
        self._revision_matrix_tab.export_copy_finished.connect(
            self._on_matrix_export_finished
        )
        self._revision_matrix_tab.export_pins_changed.connect(
            self._on_export_pins_changed
        )
        self._revision_matrix_tab.pair_compare_log.connect(self._append_log)
        self._auto_mto_compare_refresh_timer = QTimer(self)
        self._auto_mto_compare_refresh_timer.setSingleShot(True)
        self._auto_mto_compare_refresh_timer.setInterval(400)
        self._auto_mto_compare_refresh_timer.timeout.connect(
            self._on_auto_mto_compare_finished
        )
        self._revision_matrix_tab.auto_mto_compare_finished.connect(
            self._schedule_auto_mto_compare_refresh
        )
        self._revision_matrix_table = self._revision_matrix_tab.table()
        self._tabs.addTab(self._revision_matrix_tab, "Ревизии MTO")
        self._mto_worklist_tab = MtoWorklistTab(self)
        self._mto_worklist_tab.kit_activated.connect(self._on_revision_matrix_kit)
        self._mto_worklist_tab.kit_documents_requested.connect(
            self._jump_to_kit_documents
        )
        self._mto_worklist_tab.prepare_context_menu.connect(
            self._on_mto_worklist_prepare_menu
        )
        self._mto_worklist_table = self._mto_worklist_tab.table()
        self._tabs.addTab(self._mto_worklist_tab, "MTO · Перечень")
        self._an_tab = AnTab(self)
        self._an_tab.configure(self.config.runtime_dir)
        self._an_tab.kit_activated.connect(self._on_revision_matrix_kit)
        self._an_tab.scan_requested.connect(self._start_an_scan)
        self._an_tab.prepare_context_menu.connect(self._on_an_tab_prepare_menu)
        self._an_table = self._an_tab.table()
        self._tabs.addTab(self._an_tab, "АН")
        self._rd_dump_tab = RdDumpTab(self)
        self._rd_dump_tab.configure(self.config.runtime_dir)
        self._rd_dump_tab.kit_activated.connect(self._on_revision_matrix_kit)
        self._rd_dump_tab.scan_requested.connect(self._start_rd_dump_scan)
        self._rd_dump_tab.prepare_context_menu.connect(
            self._on_rd_dump_tab_prepare_menu
        )
        self._rd_dump_table = self._rd_dump_tab.table()
        self._tabs.addTab(self._rd_dump_tab, "РД")
        self._mto_readiness_tab = self._build_mto_tab()
        self._tabs.addTab(self._mto_readiness_tab, "MTO · Готовность робота")
        self._issuance_journal_tab = IssuanceJournalTab(
            self, database=self.database
        )
        self._issuance_journal_tab.set_sheet_links(self._sheet_link_context())
        self._issuance_journal_tab.kit_activated.connect(self._on_revision_matrix_kit)
        self._issuance_journal_tab.reviews_changed.connect(
            self._on_issuance_reviews_changed
        )
        self._issuance_journal_tab.prepare_context_menu.connect(
            self._on_issuance_journal_prepare_menu
        )
        self._issuance_journal_table = self._issuance_journal_tab.table()
        self._tabs.addTab(self._issuance_journal_tab, "Выдача · Журнал")
        self._approval_mail_tab = ApprovalMailTab(
            self, runtime_dir=self.config.runtime_dir
        )
        self._approval_mail_tab.write_requested.connect(self._on_approval_mail_write)
        self._approval_mail_tab.de_sync_requested.connect(
            self._open_sheet_de_sync_dialog
        )
        self._approval_mail_tab.kit_activated.connect(self._on_revision_matrix_kit)
        self._tabs.addTab(self._approval_mail_tab, "Письма о согласовании")
        self._documents_tab = self._build_documents_tab()
        self._tabs.addTab(self._documents_tab, "Все документы")
        self._collision_tab = self._build_collisions_tab()
        self._tabs.addTab(self._collision_tab, "Коллизии (0)")
        self._log_tab = self._build_log_tab()
        self._tabs.addTab(self._log_tab, "Журнал")
        set_perf_sink(self._append_log)
        perf_note("session", app="rd_catalog")
        self.setCentralWidget(self._tabs)
        self._tabs.setCurrentWidget(self._kits_tab)
        self.setAcceptDrops(True)
        self._mail_drop_filter = ApprovalMailDropFilter(
            self, self._on_approval_mail_drop
        )
        app = QApplication.instance()
        if app is not None:
            app.installEventFilter(self._mail_drop_filter)
        self.setStatusBar(QStatusBar(self))
        self._restore_layout()
        self._update_ban_action_label()
        if self._ban_store.load_error:
            self._append_log(self._ban_store.load_error)
        if self._skip_store.load_error:
            self._append_log(self._skip_store.load_error)
        self._tabs.currentChanged.connect(self._on_main_tab_changed)
        self._startup_timer = QTimer(self)
        self._startup_timer.setSingleShot(True)
        self._startup_timer.timeout.connect(self._startup_load)
        self._secondary_timer = QTimer(self)
        self._secondary_timer.setSingleShot(True)
        self._secondary_timer.timeout.connect(self._load_secondary_tabs)
        self._startup_load_pending = True
        self.statusBar().showMessage("Загрузка…")
        self._set_workers_enabled(False)
        self._update_action_states()
        self._web_server = WebServerProcess(self)
        self._web_open_when_ready = False
        self._web_server.ready.connect(self._on_web_server_ready)
        self._web_server.failed.connect(self._on_web_server_failed)
        self._web_server.stopped.connect(self._on_web_server_stopped)
        self._web_server.log_line.connect(self._on_web_server_log)
        self._restore_web_autostart()
        self._update_web_label()
        self._startup_timer.start(0)
        QTimer.singleShot(0, self._maybe_autostart_web)

    def _build_toolbar(self) -> None:
        toolbar = QToolBar("Сканирование", self)
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        self._scan_actions: list[QAction] = []
        for label, sources in (
            ("Сканировать всё", tuple(SourceKind)),
            ("Только РД", (SourceKind.RD,)),
            ("Только SQ", (SourceKind.SQ,)),
            ("Только MTO робота", (SourceKind.ROBOT,)),
        ):
            action = toolbar.addAction(label)
            action.triggered.connect(
                lambda _checked=False, selected=sources: self.start_scan(selected)
            )
            self._scan_actions.append(action)
        toolbar.addSeparator()
        self._google_action = toolbar.addAction("Загрузить комплекты Google")
        self._google_action.triggered.connect(self.start_google_kits_load)
        self._customer_pi_action = toolbar.addAction("База заказчика…")
        self._customer_pi_action.setToolTip(
            "Pickle ПИ и каталог АвтоМТО. Исходный xlsb не меняется."
        )
        self._customer_pi_action.triggered.connect(self._open_customer_pi_dialog)
        toolbar.addSeparator()
        self._ban_action = toolbar.addAction("Забаненные титулы")
        self._ban_action.triggered.connect(self._open_ban_dialog)
        self._status_colors_action = toolbar.addAction("Цвета статусов…")
        self._status_colors_action.triggered.connect(self._open_status_colors_dialog)
        self._layout_report_action = toolbar.addAction(_LAYOUT_REPORT_ACTION)
        self._layout_report_action.triggered.connect(self._on_layout_report)
        self._update_layout_report_action()
        toolbar.addSeparator()
        self._cancel_action = toolbar.addAction("Отмена")
        self._cancel_action.setEnabled(False)
        self._cancel_action.triggered.connect(self.cancel_scan)
        toolbar.addSeparator()
        self._progress = QProgressBar(toolbar)
        self._progress.setFixedWidth(180)
        self._progress.setRange(0, 1)
        toolbar.addWidget(self._progress)
        self._last_scan_label = QLabel("Успешных сканов нет", toolbar)
        self._last_scan_label.setContentsMargins(12, 0, 4, 0)
        toolbar.addWidget(self._last_scan_label)
        toolbar.addSeparator()
        self._web_autostart_cb = QCheckBox("Автозапуск WEB", toolbar)
        self._web_autostart_cb.setToolTip(
            "При открытии каталога запускать монитор на 0.0.0.0 "
            "(это bind, не адрес в браузере). Браузер сам не открывается. "
            "Снятие галки уже запущенный сервер не останавливает."
        )
        self._web_autostart_cb.toggled.connect(self._on_web_autostart_toggled)
        toolbar.addWidget(self._web_autostart_cb)
        self._web_open_btn = QPushButton("Открыть WEB", toolbar)
        self._web_open_btn.setToolTip(
            "Запустить сервер при необходимости и открыть "
            f"http://127.0.0.1:{DEFAULT_PORT} в браузере на этом ПК."
        )
        self._web_open_btn.clicked.connect(self._on_web_open_clicked)
        toolbar.addWidget(self._web_open_btn)
        self._web_copy_btn = QPushButton("Ссылка для коллег", toolbar)
        self._web_copy_btn.setToolTip(
            "Копирует http://<IP этого ПК>:"
            f"{DEFAULT_PORT} — откройте на другом компьютере в LAN. "
            "В браузер не пишите 0.0.0.0."
        )
        self._web_copy_btn.clicked.connect(self._copy_web_share_url)
        toolbar.addWidget(self._web_copy_btn)
        self._web_url_button = QPushButton("WEB выкл.", toolbar)
        self._web_url_button.setFlat(True)
        self._web_url_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self._web_url_button.setToolTip(
            "Текущий LAN-адрес (не 0.0.0.0). Клик тоже копирует. "
            "На этом ПК браузер открывает 127.0.0.1."
        )
        self._web_url_button.clicked.connect(self._copy_web_share_url)
        toolbar.addWidget(self._web_url_button)

    def _build_mto_tab(self) -> QWidget:
        tab = QWidget(self)
        layout = QVBoxLayout(tab)
        self._mto_sync_label = QLabel("Сверка: актуально", tab)
        layout.addWidget(self._mto_sync_label)
        self._cards: dict[str, QLabel] = {}
        cards = QHBoxLayout()
        for key, title, color in (
            ("ready", "READY", "#137333"),
            ("warning", "WARNING", "#9a6700"),
            ("blocked", "BLOCKED", "#b3261e"),
            ("pending", "Требует проверки", "#5f259f"),
        ):
            label = QLabel(f"{title}\n0", tab)
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            label.setMinimumHeight(66)
            label.setStyleSheet(
                f"font-size: 15px; font-weight: bold; color: {color}; "
                "border: 2px solid #707070; background: white; padding: 7px;"
            )
            self._cards[key] = label
            cards.addWidget(label)
        layout.addLayout(cards)

        filters = QHBoxLayout()
        filters.addWidget(QLabel("Фильтр:"))
        self._mto_filter = QLineEdit(tab)
        self._mto_filter.setPlaceholderText("Титул, марка, документ, статус…")
        self._mto_filter.textChanged.connect(self._apply_mto_filter)
        filters.addWidget(self._mto_filter, 1)
        self._mto_problems = QCheckBox("Только WARNING / BLOCKED", tab)
        self._mto_problems.toggled.connect(self._apply_mto_filter)
        filters.addWidget(self._mto_problems)
        self._mto_no_as_build = QCheckBox("Без as-build", tab)
        self._mto_no_as_build.setToolTip(
            "Скрыть пары, у которых MTO РД лежит в as-build передаче или пути."
        )
        self._mto_no_as_build.toggled.connect(self._apply_mto_filter)
        filters.addWidget(self._mto_no_as_build)
        layout.addLayout(filters)

        headers = [
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
        ]
        self._mto_table = self._make_table(headers)
        self._mto_table.itemSelectionChanged.connect(self._update_action_states)
        self._mto_table.cellDoubleClicked.connect(self._open_mto_cell)
        self._mto_table.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self._mto_table.customContextMenuRequested.connect(
            self._show_mto_context_menu
        )
        layout.addWidget(self._mto_table, 1)

        actions = QHBoxLayout()
        self._mto_buttons: dict[tuple[str, bool], QPushButton] = {}
        for side, file_label, folder_label in (
            ("rd", "Открыть MTO РД", "Папка MTO РД"),
            ("robot", "Открыть MTO робота", "Папка MTO робота"),
        ):
            for folder, label in ((False, file_label), (True, folder_label)):
                button = QPushButton(label, tab)
                button.clicked.connect(
                    lambda _checked=False, s=side, f=folder: self._open_mto_side(
                        s, folder=f
                    )
                )
                self._mto_buttons[(side, folder)] = button
                actions.addWidget(button)
        actions.addStretch(1)
        layout.addLayout(actions)
        return tab

    def _build_kits_tab(self) -> QWidget:
        tab = QWidget(self)
        layout = QVBoxLayout(tab)
        filters = QHBoxLayout()
        filters.addWidget(QLabel("Фильтр:"))
        self._kits_filter = QLineEdit(tab)
        self._kits_filter.setPlaceholderText("Титул, марка, статус, трансмиттал…")
        self._kits_filter.textChanged.connect(self._apply_kits_filter)
        filters.addWidget(self._kits_filter, 1)
        self._kits_mismatch = QCheckBox("Только расхождения", tab)
        self._kits_mismatch.toggled.connect(self._apply_kits_filter)
        filters.addWidget(self._kits_mismatch)
        self._kits_no_rd = QCheckBox("Нет в РД", tab)
        self._kits_no_robot = QCheckBox("Нет у робота", tab)
        self._kits_no_google = QCheckBox("Нет в Google", tab)
        for box in (
            self._kits_no_rd,
            self._kits_no_robot,
            self._kits_no_google,
        ):
            box.toggled.connect(self._apply_kits_filter)
            filters.addWidget(box)
        filter_block = QVBoxLayout()
        filter_block.addLayout(filters)
        mto_filters = QHBoxLayout()
        self._kits_code_a = QCheckBox("Код A", tab)
        self._kits_code_a.setToolTip(
            "Только комплекты с действующим кодом A."
        )
        self._kits_tdo = QCheckBox("Прошли ТДО", tab)
        self._kits_tdo.setToolTip(
            "Комплекты, прошедшие ТДО (включая полученные коды)."
        )
        self._kits_no_as_build = QCheckBox("Без as-build", tab)
        self._kits_no_as_build.setToolTip(
            "Скрыть комплекты, текущая передача которых as-build."
        )
        self._kits_only_as_build = QCheckBox("Только as-build", tab)
        self._kits_only_as_build.setToolTip(
            "Оставить только комплекты, текущая передача которых as-build."
        )
        self._kits_mto_problems = QCheckBox("Проблемы MTO", tab)
        self._kits_mto_problems.setToolTip(
            "Комплекты, где не хватает MTO или есть коллизии по ревизии."
        )
        self._kits_an_closes = QCheckBox("АН закрывает Авто МТО", tab)
        self._kits_an_closes.setToolTip(
            "Комплекты, где файлы АН закрывают расхождение Авто МТО с MTO РД."
        )
        for box in (
            self._kits_code_a,
            self._kits_tdo,
            self._kits_mto_problems,
            self._kits_an_closes,
        ):
            box.toggled.connect(self._apply_kits_filter)
            mto_filters.addWidget(box)
        for box in (self._kits_no_as_build, self._kits_only_as_build):
            box.toggled.connect(self._on_kits_as_build_toggled)
            mto_filters.addWidget(box)
        self._kits_progress = _ClickCopyLabel("", tab)
        empty_progress = kits_progress_stats(())
        self._kits_progress.setText(format_kits_progress_stats(empty_progress))
        progress_font = QFont(self._kits_progress.font())
        progress_font.setBold(True)
        self._kits_progress.setFont(progress_font)
        self._kits_progress.setCursor(Qt.CursorShape.PointingHandCursor)
        self._kits_progress.setWordWrap(False)
        self._kits_progress.setMinimumWidth(80)
        self._kits_progress.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Preferred,
        )
        self._kits_progress.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        self._kits_progress.copied.connect(self._on_kits_progress_copied)
        mto_filters.addWidget(self._kits_progress, 1)
        self._kits_legend_button = QPushButton(KITS_PAINT_LEGEND_BUTTON, tab)
        self._kits_legend_button.setToolTip(_KITS_LEGEND_TOOLTIP)
        self._kits_legend_button.clicked.connect(self._show_kits_paint_legend)
        mto_filters.addWidget(self._kits_legend_button)
        self._kits_de_sync_button = QPushButton(SHEET_DE_SYNC_BUTTON, tab)
        self._kits_de_sync_button.setToolTip(_KITS_DE_SYNC_TOOLTIP)
        self._kits_de_sync_button.clicked.connect(self._open_sheet_de_sync_dialog)
        mto_filters.addWidget(self._kits_de_sync_button)
        self._kits_layout_button = QPushButton(_KITS_LAYOUT_BUTTON, tab)
        self._kits_layout_button.setToolTip(_KITS_LAYOUT_TOOLTIP)
        self._kits_layout_button.clicked.connect(self._save_kits_column_template)
        mto_filters.addWidget(self._kits_layout_button)
        self._kits_xlsx_button = QPushButton(_KITS_XLSX_BUTTON, tab)
        self._kits_xlsx_button.setToolTip(_KITS_XLSX_TOOLTIP)
        self._kits_xlsx_button.clicked.connect(self._export_kits_xlsx)
        mto_filters.addWidget(self._kits_xlsx_button)
        filter_block.addLayout(mto_filters)
        layout.addLayout(filter_block)

        self._kits_load_banner = QFrame(tab)
        self._kits_load_banner.setObjectName("kitsLoadBanner")
        self._kits_load_banner.setFixedHeight(48)
        self._kits_load_banner.setAttribute(
            Qt.WidgetAttribute.WA_StyledBackground, True
        )
        self._kits_load_banner.setStyleSheet(
            "QFrame#kitsLoadBanner {"
            " background-color: #F7E8BE;"
            " border: 1px solid #D4C48A;"
            "}"
        )
        banner_row = QHBoxLayout(self._kits_load_banner)
        banner_row.setContentsMargins(12, 6, 12, 6)
        banner_label = QLabel("Загрузка комплектов…", self._kits_load_banner)
        banner_font = QFont(banner_label.font())
        banner_font.setBold(True)
        point_size = banner_font.pointSize()
        if point_size > 0:
            banner_font.setPointSize(max(point_size + 2, 11))
        banner_label.setFont(banner_font)
        banner_progress = QProgressBar(self._kits_load_banner)
        banner_progress.setRange(0, 0)
        banner_progress.setTextVisible(False)
        banner_row.addWidget(banner_label)
        banner_row.addWidget(banner_progress, 1)
        layout.addWidget(self._kits_load_banner)

        self._kits_splitter = QSplitter(Qt.Orientation.Vertical, tab)
        self._kits_table = self._make_table(list(_KITS_HEADERS))
        compare_header = self._kits_table.horizontalHeaderItem(
            _KITS_COL_AUTO_MTO_COMPARE
        )
        if compare_header is not None:
            compare_header.setToolTip(AUTO_MTO_COMPARE_STATUS_TOOLTIP)
        mto_rev_header = self._kits_table.horizontalHeaderItem(_KITS_COL_MTO_REV)
        if mto_rev_header is not None:
            mto_rev_header.setToolTip(_KITS_MTO_REV_TOOLTIP)
        working_rev_header = self._kits_table.horizontalHeaderItem(
            _KITS_COL_WORKING
        )
        if working_rev_header is not None:
            working_rev_header.setToolTip(_KITS_WORKING_REV_TOOLTIP)
        ok_header = self._kits_table.horizontalHeaderItem(_KITS_COL_OK)
        if ok_header is not None:
            ok_header.setToolTip(KITS_OK_TOOLTIP)
        self._kits_table.itemSelectionChanged.connect(self._on_kit_selected)
        _enable_readonly_cell_copy(self._kits_table)
        self._kits_table.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self._kits_table.customContextMenuRequested.connect(
            self._show_kits_context_menu
        )
        self._kits_detail_tabs = QTabWidget()
        card_tab = QWidget()
        card_layout = QVBoxLayout(card_tab)
        card_layout.setContentsMargins(0, 0, 0, 0)
        header_row = QHBoxLayout()
        self._kits_card_title = QLabel("Выберите комплект", card_tab)
        self._kits_card_title.setStyleSheet("font-weight: bold; font-size: 15px;")
        header_row.addWidget(self._kits_card_title, 1)
        self._kits_card_summary = QLabel("", card_tab)
        header_row.addWidget(self._kits_card_summary)
        self._kits_card_pipeline = QLabel("", card_tab)
        self._kits_card_pipeline.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._kits_card_pipeline.setMinimumWidth(180)
        header_row.addWidget(self._kits_card_pipeline)
        self._kits_card_approval = QLabel("", card_tab)
        self._kits_card_approval.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._kits_card_approval.setMinimumWidth(220)
        header_row.addWidget(self._kits_card_approval)
        card_layout.addLayout(header_row)
        self._kits_card_revs = QLabel("", card_tab)
        self._kits_card_revs.setWordWrap(True)
        card_layout.addWidget(self._kits_card_revs)
        self._kits_card_mto = QLabel("", card_tab)
        self._kits_card_mto.setWordWrap(True)
        card_layout.addWidget(self._kits_card_mto)
        self._kits_card_sources = QLabel("", card_tab)
        self._kits_card_sources.setWordWrap(True)
        self._kits_card_sources.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        card_layout.addWidget(self._kits_card_sources)
        self._kits_package_table = self._make_table(list(CARD_PACKAGE_HEADERS))
        pkg_header = self._kits_package_table.horizontalHeader()
        for index in range(pkg_header.count()):
            mode = (
                QHeaderView.ResizeMode.Stretch
                if index in _KITS_PKG_STRETCH_COLS
                else QHeaderView.ResizeMode.ResizeToContents
            )
            pkg_header.setSectionResizeMode(index, mode)
        self._kits_package_table.setWordWrap(True)
        _enable_readonly_cell_copy(self._kits_package_table)
        self._kits_package_table.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self._kits_package_table.customContextMenuRequested.connect(
            self._show_kit_package_context_menu
        )
        card_layout.addWidget(self._kits_package_table, 1)
        pkg_actions = QHBoxLayout()
        self._kits_pkg_open_button = QPushButton(
            "Открыть папку пакета", card_tab
        )
        self._kits_pkg_copy_button = QPushButton(
            "Копировать путь пакета", card_tab
        )
        self._kits_pkg_jump_button = QPushButton(
            "Показать в «Все документы»", card_tab
        )
        self._kits_pkg_handoff_button = QPushButton(
            "Передать роботу", card_tab
        )
        self._kits_pkg_open_button.clicked.connect(
            self._open_selected_kit_package
        )
        self._kits_pkg_copy_button.clicked.connect(
            self._copy_selected_kit_package_path
        )
        self._kits_pkg_jump_button.clicked.connect(
            self._jump_selected_kit_package_documents
        )
        self._kits_pkg_handoff_button.clicked.connect(
            lambda: self._copy_kit_package_handoff()
        )
        for button in (
            self._kits_pkg_open_button,
            self._kits_pkg_copy_button,
            self._kits_pkg_jump_button,
            self._kits_pkg_handoff_button,
        ):
            button.setEnabled(False)
            pkg_actions.addWidget(button)
        pkg_actions.addStretch(1)
        card_layout.addLayout(pkg_actions)
        self._kits_card_legend = QHBoxLayout()
        card_layout.addLayout(self._kits_card_legend)
        self._reload_kits_card_legend()
        self._kits_package_table.itemSelectionChanged.connect(
            self._update_action_states
        )
        self._kits_tooltips = self._make_kits_mono_edit()
        self._kits_tooltips.setPlaceholderText(_KITS_TIPS_PLACEHOLDER)
        self._kits_tips_click_filter = attach_google_href_clicks(
            self._kits_table,
            on_ctrl_click=self._on_kits_table_ctrl_click,
        )
        self._kits_detail = self._make_kits_mono_edit()
        self._kits_detail.setPlaceholderText(
            "Выберите строку, чтобы увидеть историю Google и файлы источников."
        )
        self._kits_detail_tabs.addTab(self._kits_tooltips, "Подсказки")
        self._kits_detail_tabs.addTab(card_tab, "Карточка")
        self._kits_detail_tabs.addTab(self._kits_detail, "Текст")
        self._kits_detail_tabs.setCurrentIndex(0)
        self._kits_detail_placement_button = QToolButton(self._kits_detail_tabs)
        self._kits_detail_placement_button.setAutoRaise(True)
        self._kits_detail_placement_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._kits_detail_placement_button.setCursor(
            Qt.CursorShape.PointingHandCursor
        )
        self._kits_detail_placement_button.setFixedSize(22, 20)
        self._kits_detail_placement_button.clicked.connect(
            self._toggle_kits_detail_placement
        )
        self._kits_detail_tabs.setCornerWidget(
            self._kits_detail_placement_button,
            Qt.Corner.TopRightCorner,
        )
        self._kits_splitter.addWidget(self._kits_table)
        self._kits_splitter.addWidget(self._kits_detail_tabs)
        self._kits_splitter.setChildrenCollapsible(False)
        self._kits_detail_placement_value = _KITS_DETAIL_PLACEMENT_BOTTOM
        self._apply_kits_detail_placement(
            _KITS_DETAIL_PLACEMENT_BOTTOM, restore_sizes=False
        )
        layout.addWidget(self._kits_splitter, 1)

        actions = QHBoxLayout()
        self._kit_buttons: dict[tuple[str, bool], QPushButton] = {}
        for source, file_label, folder_label in (
            ("rd", "Открыть файл РД", "Папка РД"),
            ("robot", "Открыть файл робота", "Папка робота"),
            ("sq", "Открыть файл SQ", "Папка SQ"),
        ):
            for folder, label in ((False, file_label), (True, folder_label)):
                button = QPushButton(label, tab)
                button.clicked.connect(
                    lambda _checked=False, s=source, f=folder: self._open_kit_source(
                        s, folder=f
                    )
                )
                self._kit_buttons[(source, folder)] = button
                actions.addWidget(button)
        self._kit_copy_button = QPushButton("Копировать пути", tab)
        self._kit_copy_button.clicked.connect(self._copy_kit_paths)
        actions.addWidget(self._kit_copy_button)
        actions.addStretch(1)
        layout.addLayout(actions)
        return tab

    def _build_collisions_tab(self) -> QWidget:
        tab = QWidget(self)
        layout = QVBoxLayout(tab)
        self._collision_count_label = QLabel("Текущие коллизии: 0", tab)
        self._collision_count_label.setStyleSheet(
            "font-size: 15px; font-weight: bold; color: #b3261e;"
        )
        layout.addWidget(self._collision_count_label)
        self._collision_table = self._make_table(
            ["Область", "Источник", "Тип", "Документ", "Сообщение", "Пути"]
        )
        self._collision_table.itemSelectionChanged.connect(
            self._update_action_states
        )
        self._collision_table.cellDoubleClicked.connect(
            lambda _row, _column: self._open_collision_folder()
        )
        self._collision_table.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self._collision_table.customContextMenuRequested.connect(
            self._show_collision_context_menu
        )
        layout.addWidget(self._collision_table, 1)

        actions = QHBoxLayout()
        self._collision_folder_button = QPushButton(
            "Открыть содержащую папку", tab
        )
        self._collision_copy_button = QPushButton("Копировать пути", tab)
        self._collision_folder_button.clicked.connect(
            self._open_collision_folder
        )
        self._collision_copy_button.clicked.connect(self._copy_collision_paths)
        actions.addWidget(self._collision_folder_button)
        actions.addWidget(self._collision_copy_button)
        actions.addStretch(1)
        layout.addLayout(actions)
        return tab

    def _build_documents_tab(self) -> QWidget:
        tab = QWidget(self)
        layout = QVBoxLayout(tab)
        filters = QHBoxLayout()
        self._doc_filter = QLineEdit(tab)
        self._doc_filter.setPlaceholderText("Фильтр по титулу, марке, документу, пути…")
        self._doc_filter_timer = QTimer(tab)
        self._doc_filter_timer.setSingleShot(True)
        self._doc_filter_timer.setInterval(_DOC_FILTER_DEBOUNCE_MS)
        self._doc_filter_timer.timeout.connect(self._apply_document_tree_filter)
        self._doc_filter.textChanged.connect(self._on_doc_filter_text_changed)
        filters.addWidget(self._doc_filter, 1)
        self._doc_pending = QCheckBox("Требует проверки", tab)
        self._doc_pending.toggled.connect(self._rebuild_document_tree)
        filters.addWidget(self._doc_pending)
        self._doc_problems = QCheckBox("Только проблемы", tab)
        self._doc_problems.setToolTip(
            "Нет файла, ошибка разбора имени или путь вне "
            "«титул / марка / Для передачи / NN_»."
        )
        self._doc_problems.toggled.connect(self._rebuild_document_tree)
        filters.addWidget(self._doc_problems)
        self._skip_edit_button = QPushButton("Skip-папки…", tab)
        self._skip_edit_button.clicked.connect(self._open_skip_dirs_dialog)
        filters.addWidget(self._skip_edit_button)
        self._skip_prune_button = QPushButton("Убрать skip из дерева", tab)
        self._skip_prune_button.setToolTip(
            "Пометить в базе файлы под skip-папки как отсутствующие, "
            "без скана сети. Дерево уже прячет их после сохранения списка."
        )
        self._skip_prune_button.clicked.connect(self._prune_skipped_from_db)
        filters.addWidget(self._skip_prune_button)
        layout.addLayout(filters)

        labels = QHBoxLayout()
        labels_caption = QLabel("В узле ревизии:", tab)
        labels_caption.setToolTip(
            "Дополнительные поля в дереве «титул → марка → ревизия». "
            "Сохраняются локально, как размеры окна. "
            "Статус Google показывается только если папка NN сопоставлена "
            "с Выдачей или столбцом F (TRM, ревизия или дата ±2 дня)."
        )
        labels.addWidget(labels_caption)
        self._doc_show_mto_status = QCheckBox("Статус MTO", tab)
        self._doc_show_mto_status.setChecked(True)
        self._doc_show_mto_status.setToolTip(
            "Статус MTO этой ревизии: код A / ТДО / нет MTO. "
            "Цвет квадрата — как на «Ревизии MTO»."
        )
        self._doc_show_date = QCheckBox("Дата", tab)
        self._doc_show_date.setChecked(True)
        self._doc_show_date.setToolTip(
            "Дата последнего сохранения PDF / MTO / DWG в этой папке."
        )
        self._doc_show_folder = QCheckBox("Папка NN", tab)
        self._doc_show_folder.setToolTip(
            "Имя папки передачи (05_рев.0-AN02_…), в которой лежит комплект."
        )
        self._doc_show_review = QCheckBox("Статус Google", tab)
        self._doc_show_review.setToolTip(
            "Статус согласования этой папки из Выдачи / столбца F, "
            "если цикл сопоставлен. Иначе поле не добавляется."
        )
        self._doc_show_current = QCheckBox("Текущая", tab)
        self._doc_show_current.setToolTip(
            "Метка «текущая», если в папке есть файлы текущего состава."
        )
        self._doc_show_mto = QCheckBox("MTO", tab)
        self._doc_show_mto.setToolTip("Метка «MTO», если в этой папке есть MTO XLSX.")
        self._doc_show_working = QCheckBox("Рабочая", tab)
        self._doc_show_working.setToolTip(
            "Слово «рабочая» в подписи узла всегда, когда папка рабочая "
            "(вручную или ревизия выше выдачи). Чекбокс больше не скрывает "
            "метку. Столбец «Рабочая» и цвет узла тоже всегда. "
            "«Статус MTO» — цикл Google F этой ревизии, не пометка папки."
        )
        self._doc_show_as_build = QCheckBox("As-build", tab)
        self._doc_show_as_build.setToolTip(
            "Метка «AB», если в папке есть as-build файлы "
            "(имя передачи или сегмент пути)."
        )
        for box in (
            self._doc_show_mto_status,
            self._doc_show_date,
            self._doc_show_folder,
            self._doc_show_review,
            self._doc_show_current,
            self._doc_show_mto,
            self._doc_show_working,
            self._doc_show_as_build,
        ):
            box.toggled.connect(self._on_tree_label_option_changed)
            labels.addWidget(box)
        labels.addStretch(1)
        layout.addLayout(labels)

        self._doc_splitter = QSplitter(Qt.Orientation.Horizontal, tab)
        self._doc_tree = QTreeWidget(self._doc_splitter)
        self._doc_tree.setHeaderLabel("Титул → марка → ревизия")
        self._doc_tree.itemSelectionChanged.connect(self._refresh_history)
        self._doc_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._doc_tree.customContextMenuRequested.connect(
            self._show_document_tree_context_menu
        )
        self._history = self._make_table(list(_HISTORY_HEADERS))
        date_header = self._history.horizontalHeaderItem(_HISTORY_COL_DATE)
        if date_header is not None:
            date_header.setToolTip(_HISTORY_DATE_COLUMN_TIP)
        self._history.itemSelectionChanged.connect(self._update_action_states)
        self._history.cellDoubleClicked.connect(
            lambda _row, _column: self._open_history(folder=False)
        )
        self._history.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._history.customContextMenuRequested.connect(
            self._show_history_context_menu
        )
        self._history_copy_shortcut = QShortcut(
            QKeySequence.StandardKey.Copy, self._history
        )
        self._history_copy_shortcut.setContext(Qt.ShortcutContext.WidgetShortcut)
        self._history_copy_shortcut.activated.connect(
            lambda: self._copy_history_column(self._history.currentColumn())
        )
        self._doc_splitter.addWidget(self._doc_tree)
        self._doc_splitter.addWidget(self._history)
        self._doc_splitter.setStretchFactor(0, 1)
        self._doc_splitter.setStretchFactor(1, 3)
        self._doc_splitter.setSizes([340, 1000])
        layout.addWidget(self._doc_splitter, 1)

        actions = QHBoxLayout()
        self._ack_button = QPushButton("Подтвердить", tab)
        self._comment_button = QPushButton("Комментарий", tab)
        self._ignore_button = QPushButton("Игнорировать", tab)
        self._history_open_button = QPushButton("Открыть файл", tab)
        self._history_folder_button = QPushButton("Открыть папку", tab)
        self._copy_button = QPushButton("Копировать полный путь", tab)
        self._ack_button.clicked.connect(
            lambda: self._review_selected(ReviewState.ACKNOWLEDGED)
        )
        self._comment_button.clicked.connect(lambda: self._review_selected("comment"))
        self._ignore_button.clicked.connect(
            lambda: self._review_selected(ReviewState.IGNORED)
        )
        self._history_open_button.clicked.connect(
            lambda: self._open_history(folder=False)
        )
        self._history_folder_button.clicked.connect(
            lambda: self._open_history(folder=True)
        )
        self._copy_button.clicked.connect(self._copy_history_path)
        for button in (
            self._ack_button,
            self._comment_button,
            self._ignore_button,
            self._history_open_button,
            self._history_folder_button,
            self._copy_button,
        ):
            actions.addWidget(button)
        actions.addStretch(1)
        layout.addLayout(actions)
        return tab

    def _build_log_tab(self) -> QWidget:
        tab = QWidget(self)
        layout = QVBoxLayout(tab)
        hint = QLabel(
            f"Профилирование: {self.config.runtime_dir / 'perf.log'}  "
            "(BEGIN/END; RD_CATALOG_PERF=0 выключает)",
            tab,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)
        self._log = QPlainTextEdit(tab)
        self._log.setReadOnly(True)
        layout.addWidget(self._log)
        return tab

    @staticmethod
    def _make_table(headers: list[str]) -> QTableWidget:
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setSortingEnabled(True)
        table.verticalHeader().setVisible(False)
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setSectionsMovable(True)
        header.setFirstSectionMovable(True)
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(40)
        return table

    @staticmethod
    def _make_kits_mono_edit() -> QPlainTextEdit:
        """Read-only monospace pane for tooltips and the legacy Текст dump."""

        edit = QPlainTextEdit()
        edit.setReadOnly(True)
        font = QFont("Consolas")
        font.setStyleHint(QFont.StyleHint.Monospace)
        edit.setFont(font)
        edit.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        space = edit.fontMetrics().horizontalAdvance(" ")
        edit.setTabStopDistance(max(space, 1) * 8)
        return edit

    def _header_settings_map(self) -> list[tuple[str, QTableWidget]]:
        """Return QSettings keys paired with tables whose column layout persists."""

        pairs: list[tuple[str, QTableWidget]] = []
        for key, attr in (
            ("window/mto_header", "_mto_table"),
            # v4: «Сверка Авто МТО» after pin; do not restore v3.
            ("window/mto_worklist_header_v4", "_mto_worklist_table"),
            ("window/issuance_journal_header_v1", "_issuance_journal_table"),
            ("window/an_tab_header_v2", "_an_table"),
            ("window/rd_dump_tab_header_v2", "_rd_dump_table"),
            # v10: «Ок» after «Марка»; do not restore v9.
            ("window/kits_header_v10", "_kits_table"),
            ("window/collision_header", "_collision_table"),
            # v8: column «Аннул.» after «Рабочая». Do not restore v7.
            ("window/history_header_v8", "_history"),
        ):
            table = getattr(self, attr, None)
            if isinstance(table, QTableWidget):
                pairs.append((key, table))
        return pairs

    def _restore_table_headers(self) -> None:
        kits_header_restored = False
        for key, table in self._header_settings_map():
            header = table.horizontalHeader()
            expected = table.columnCount()
            saved_count = self._settings.value(f"{key}_count")
            count_matches = True
            if saved_count is not None:
                try:
                    count_matches = int(saved_count) == expected
                except (TypeError, ValueError):
                    count_matches = False
            state = self._settings.value(key)
            if (
                count_matches
                and isinstance(state, QByteArray)
                and not state.isEmpty()
            ):
                header.restoreState(state)
                if key == "window/kits_header_v10":
                    kits_header_restored = True
            if header.count() != expected:
                header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
                header.setStretchLastSection(False)
            # Old saved states were written with immovable sections.
            header.setSectionsMovable(True)
            header.setFirstSectionMovable(True)
            header.setMinimumSectionSize(40)
        if not kits_header_restored:
            self._apply_default_kits_header_layout()
        self._apply_default_kits_title_sort()
        compare_header = self._kits_table.horizontalHeaderItem(
            _KITS_COL_AUTO_MTO_COMPARE
        )
        if compare_header is not None:
            compare_header.setToolTip(AUTO_MTO_COMPARE_STATUS_TOOLTIP)
        mto_header = self._kits_table.horizontalHeaderItem(_KITS_COL_MTO_REV)
        if mto_header is not None:
            mto_header.setToolTip(_KITS_MTO_REV_TOOLTIP)
        working_header = self._kits_table.horizontalHeaderItem(_KITS_COL_WORKING)
        if working_header is not None:
            working_header.setToolTip(_KITS_WORKING_REV_TOOLTIP)
        ok_header = self._kits_table.horizontalHeaderItem(_KITS_COL_OK)
        if ok_header is not None:
            ok_header.setToolTip(KITS_OK_TOOLTIP)
        date_header = self._history.horizontalHeaderItem(_HISTORY_COL_DATE)
        if date_header is not None:
            date_header.setToolTip(_HISTORY_DATE_COLUMN_TIP)

    def _apply_default_kits_title_sort(self) -> None:
        """Sort Комплекты by title from early numbers to late.

        ``QHeaderView`` defaults to section 0 + ``DescendingOrder``, so
        enabling sort showed 9110 before 1000. Saved ``kits_header_v10``
        may still carry that. A sort on another column is left as-is.
        """

        if not hasattr(self, "_kits_table"):
            return
        table = self._kits_table
        header = table.horizontalHeader()
        section = header.sortIndicatorSection()
        if header.isSortIndicatorShown() and section != _KITS_COL_TITLE:
            return
        table.sortByColumn(_KITS_COL_TITLE, Qt.SortOrder.AscendingOrder)
        header.setSortIndicatorShown(True)

    def _apply_default_kits_header_layout(self) -> None:
        """Apply the saved Комплекты template when QSettings header is stale.

        Runtime ``kits_table_layout.json`` wins over the packaged factory.
        Widths from the template become the default sizes. If the template
        has no widths, keep the v5 width copy as a one-time migration.
        """

        layout = load_default_kits_table_layout(self.config.runtime_dir)
        if layout is not None:
            apply_kits_header_layout(
                self._kits_table.horizontalHeader(),
                list(_KITS_HEADERS),
                layout,
            )
        if layout is None or not layout.widths:
            self._copy_kits_header_widths_from_v5()

    def _capture_kits_table_layout(self) -> KitsTableLayout:
        """Return the current Комплекты visual order and widths."""

        return capture_kits_header_layout(
            self._kits_table.horizontalHeader(),
            list(_KITS_HEADERS),
        )

    def _write_kits_column_template(
        self,
        *,
        write_packaged: bool = True,
    ) -> KitsLayoutSaveResult:
        """Persist the current Комплекты layout as the default template.

        Args:
            write_packaged: Also overwrite the factory JSON in the package.

        Returns:
            Runtime path and optional factory path/error.

        Raises:
            OSError: If the runtime file cannot be written.
            ValueError: If the table header does not match ``_KITS_HEADERS``.
        """

        return save_default_kits_table_layout(
            self._capture_kits_table_layout(),
            self.config.runtime_dir,
            write_packaged=write_packaged,
        )

    def _flash_kits_layout_button(self, *, ok: bool) -> None:
        """Paint the template button green on success or yellow on problems."""

        fill = _KITS_LAYOUT_OK_FILL if ok else _KITS_LAYOUT_PROBLEM_FILL
        button = self._kits_layout_button
        button.setStyleSheet(
            f"QPushButton {{ background-color: {fill}; color: #202124; }}"
        )
        button.style().unpolish(button)
        button.style().polish(button)
        button.update()

    def _show_kits_paint_legend(self) -> None:
        """Button: examples of Комплекты MTO/revision fill and bold."""

        dialog = KitsPaintLegendDialog(self._status_colors, self)
        dialog.exec()

    def _open_sheet_de_sync_dialog(self) -> None:
        """Button: list stale D/E vs last F and write selected jobs."""

        if self._busy():
            QMessageBox.information(
                self,
                SHEET_DE_SYNC_TITLE,
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return
        candidates = list_sheet_de_sync_rows(
            [
                row
                for row in self._kit_rows
                if not self._is_banned_pair(row.title, row.mark, row.title_system)
            ]
        )
        if not candidates:
            QMessageBox.information(
                self,
                SHEET_DE_SYNC_TITLE,
                "Нет комплектов, где последнее событие F совпадает с РД, "
                "а столбцы D/E ещё можно поправить.",
            )
            return
        dialog = SheetDeSyncDialog(candidates, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        jobs = dialog.selected_jobs()
        if not jobs:
            return
        self._start_google_journal_write(
            jobs, log_line="Запись D/E в КСБ ИД…"
        )

    def _export_kits_xlsx(self) -> None:
        """Save the visible Комплекты table as a styled xlsx workbook."""

        path, _filter = QFileDialog.getSaveFileName(
            self,
            _KITS_XLSX_SAVE_CAPTION,
            dated_xlsx_filename("Комплекты"),
            _KITS_XLSX_SAVE_FILTER,
        )
        if not path:
            return
        dest = Path(path)
        if dest.suffix.casefold() != ".xlsx":
            dest = dest.with_suffix(".xlsx")
        exported = snapshot_qtable(
            self._kits_table,
            visible_only=True,
            monitor_role=_ROLE_MONITOR,
            header_names=list(_KITS_HEADERS),
            sheet_name="Комплекты",
        )
        try:
            write_exported_table_xlsx(exported, dest)
        except (OSError, ValueError) as exc:
            QMessageBox.warning(
                self,
                _KITS_XLSX_TITLE,
                f"Не удалось сохранить таблицу:\n{type(exc).__name__}: {exc}",
            )
            return
        self._append_log(f"Таблица Комплекты сохранена: {dest}")
        open_reply = QMessageBox.question(
            self,
            _KITS_XLSX_TITLE,
            f"Таблица сохранена:\n{dest}\n\nОткрыть файл?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if open_reply == QMessageBox.StandardButton.Yes:
            self._open_result(str(dest), folder=False)

    def _save_kits_column_template(self) -> None:
        """Button: write the current Комплекты order and widths as default."""

        try:
            saved = self._write_kits_column_template()
        except (OSError, ValueError) as exc:
            self._append_log(
                f"Шаблон колонок: не записан ({type(exc).__name__}: {exc})"
            )
            self._flash_kits_layout_button(ok=False)
            return
        except Exception as exc:
            self._append_log(
                f"Шаблон колонок: ошибка ({type(exc).__name__}: {exc})"
            )
            self._flash_kits_layout_button(ok=False)
            return
        if saved.packaged_error:
            self._append_log(
                f"Шаблон колонок: runtime {saved.runtime_path}; "
                f"заводской не записан ({saved.packaged_error})"
            )
            self._flash_kits_layout_button(ok=False)
            return
        extra = (
            f"; заводской {saved.packaged_path}"
            if saved.packaged_path is not None
            else ""
        )
        self._append_log(f"Шаблон колонок сохранён: {saved.runtime_path}{extra}")
        self._flash_kits_layout_button(ok=True)

    def _copy_kits_header_widths_from_v5(self) -> None:
        """Keep v5 column widths after the v6 logical reorder."""

        state = self._settings.value("window/kits_header_v5")
        saved_count = self._settings.value("window/kits_header_v5_count")
        if not isinstance(state, QByteArray) or state.isEmpty():
            return
        try:
            if saved_count is not None and int(saved_count) != len(_KITS_HEADERS):
                return
        except (TypeError, ValueError):
            return
        old_headers = (
            "Титул",
            "Марка",
            PIN_COLUMN_HEADER,
            "Google · рев.",
            "РД · рев.",
            "SQ · рев.",
            "Робот · рев.",
            "Авто МТО",
            AUTO_MTO_COMPARE_STATUS_HEADER,
            "Сводка",
            "Статус рассмотрения",
            "Статус согласования",
            "Google · статус",
            "Google · дата F",
            "Google · этап F",
            "Google · рев. F",
            "Google · TRM F",
            "Выдача · рев.",
            "Выдача · статус",
            "Выдача · дата отпр.",
            "Выдача · TRM отпр.",
            "Выдача · дата вх.контр.",
            "Выдача · TRM подтв.",
            "РД · дата файла",
        )
        dummy = QTableWidget(0, len(old_headers))
        dummy.setHorizontalHeaderLabels(list(old_headers))
        dummy_header = dummy.horizontalHeader()
        if not dummy_header.restoreState(state):
            dummy.deleteLater()
            return
        new_by_name = {name: index for index, name in enumerate(_KITS_HEADERS)}
        target = self._kits_table.horizontalHeader()
        for old_index, name in enumerate(old_headers):
            new_index = new_by_name.get(name)
            if new_index is None:
                continue
            width = dummy_header.sectionSize(old_index)
            if width > 0:
                target.resizeSection(new_index, width)
        dummy.deleteLater()

    def _save_table_headers(self) -> None:
        for key, table in self._header_settings_map():
            self._settings.setValue(key, table.horizontalHeader().saveState())
            self._settings.setValue(f"{key}_count", table.columnCount())

    def _hydrate_google_from_cache(self) -> None:
        """Load KSB ИД / issuance from SQLite, else JSON cache, if needed."""

        if self._google_loaded:
            return
        if self._google_kits or self._issuance_kits or self._issuance_sends:
            self._google_loaded = True
            return
        state = load_google_for_monitor(
            self.database, str(self.config.runtime_dir)
        )
        if not state.loaded:
            return
        self._google_kits = state.kits
        self._issuance_kits = state.issuance_kits
        self._issuance_sends = state.issuance_sends
        self._google_source = state.source
        self._google_fetched_at = state.fetched_at
        self._google_warning = state.warning
        self._google_loaded = True
        self._issuance_kits_fresh = True

    def _assign_catalog_records(
        self,
        records: list[FileRecord],
        *,
        contour: Sequence[FileRecord] | None = None,
    ) -> None:
        """Replace the file snapshot and refresh the canonical contour.

        Args:
            records: Latest ``list_files`` rows.
            contour: Precomputed contour (startup worker). When omitted,
                reuse the previous membership for unchanged paths.
        """

        previous_records = self._all_records
        previous_contour = self._contour_records
        self._all_records = records
        self._record_by_id = {record.id: record for record in records}
        self._rebuild_catalog_record_indexes()
        self._official_rd_package_cache = None
        self._official_rd_package_by_kit = {}
        if contour is not None:
            self._contour_records = tuple(contour)
        else:
            self._contour_records = tuple(
                records_in_contour(
                    records,
                    self.config.rd_root,
                    previous_records=previous_records or None,
                    previous_contour=previous_contour,
                )
            )
        self._official_ids_token = None
        self._tree_kit_identities_cache = None
        self._contour_ids_cache = None

    def _rebuild_catalog_record_indexes(self) -> None:
        """Index catalog files by path_key and kit for GUI lookups."""

        by_path: dict[str, FileRecord] = {}
        by_kit: dict[tuple[str, str], list[FileRecord]] = {}
        for record in self._all_records:
            key = str(record.path_key or "").casefold()
            if key:
                by_path[key] = record
            title = str(record.data.get("title") or "").strip()
            mark = str(record.data.get("mark") or "").strip()
            if title and mark:
                by_kit.setdefault(kit_identity_key(title, mark), []).append(record)
        self._records_by_path_key = by_path
        self._records_by_kit = by_kit
        self._catalog_index_token = id(self._all_records)

    def _ensure_catalog_record_indexes(self) -> None:
        if getattr(self, "_catalog_index_token", None) == id(self._all_records):
            return
        self._rebuild_catalog_record_indexes()

    def _records_for_kit(self, title: str, mark: str) -> Sequence[FileRecord]:
        """Return in-memory files of one title–mark, or empty."""

        self._ensure_catalog_record_indexes()
        return self._records_by_kit.get(kit_identity_key(title, mark), ())

    def _pipeline_records(self) -> Sequence[FileRecord]:
        """Return contour-filtered records when refresh already computed them."""

        if self._contour_records is not None:
            return self._contour_records
        return self._all_records

    def _pipeline_rd_root(self) -> str | Path | None:
        """Skip a second canonical-path walk when contour records are ready."""

        if self._contour_records is not None:
            return None
        return self.config.rd_root

    def _should_rebuild_empty_derived(self) -> bool:
        """Return True when derived kit tables are missing or stale."""

        try:
            pipelines = list_kit_pipelines(self.database)
        except Exception:
            pipelines = ()
        if pipeline_algorithm_needs_rebuild(pipelines):
            return True
        if pipelines:
            return False
        return bool(
            self._all_records
            or self._google_kits
            or self._issuance_kits
            or self._issuance_sends
        )

    def _ingest_google_and_rebuild_pipeline(
        self,
        *,
        ingest_google: bool = True,
        kit_keys: set[tuple[str, str]] | None = None,
    ) -> None:
        """Write Google snapshots to SQLite and rebuild derived kit tables.

        Uses cached or in-memory kits plus every issuance send. Does not parse
        CSV. Skips while a scan/Google/robot/SQ worker is active. Database
        errors are logged and do not abort ``refresh``.

        Args:
            ingest_google: Rewrite Google/issuance snapshot tables. File
                scans must pass ``False`` — sheets did not change.
            kit_keys: Optional identities for a folder-scoped rebuild.
        """

        if self._catalog_workers_busy():
            return
        with perf_span(
            "gui.ingest_rebuild_pipeline",
            ingest_google=ingest_google,
            kits=len(kit_keys) if kit_keys else "all",
        ):
            try:
                if ingest_google:
                    self._hydrate_google_from_cache()
                    kits = self._google_kits
                    sends = self._issuance_sends
                    if kits or sends:
                        ingest_google_snapshot(
                            self.database,
                            kits,
                            sends,
                            loaded_at=self._google_fetched_at
                            or datetime.now(timezone.utc).isoformat(),
                            source=self._google_source or "cache",
                            warning=self._google_warning,
                        )
                        self._issuance_kits = latest_effective_issuance_kits(
                            self.database
                        )
                        self._issuance_kits_fresh = True
                rebuild_pipeline(
                    self.database,
                    records=self._pipeline_records(),
                    detected_current_ids=self._detected_current_ids,
                    kit_keys=kit_keys,
                    rd_root=self._pipeline_rd_root(),
                )
            except Exception as exc:
                self._append_log(f"Pipeline: {type(exc).__name__}: {exc}")

    def _load_folder_tree_hints(self) -> None:
        """Cache per-folder Google/pipeline extras for the documents tree."""

        with perf_span("gui.load_folder_tree_hints"):
            try:
                self._folder_hints = list_folder_tree_hints(self.database)
            except Exception as exc:
                self._folder_hints = {}
                self._append_log(f"Подписи дерева: {type(exc).__name__}: {exc}")

    def _tree_label_options(self) -> DocumentTreeLabelOptions:
        """Return the current documents-tree extra-field checkboxes."""

        if not hasattr(self, "_doc_show_date"):
            return DocumentTreeLabelOptions()
        return DocumentTreeLabelOptions(
            show_mto_status=self._doc_show_mto_status.isChecked(),
            show_date=self._doc_show_date.isChecked(),
            show_folder=self._doc_show_folder.isChecked(),
            show_review=self._doc_show_review.isChecked(),
            show_current=self._doc_show_current.isChecked(),
            show_mto=self._doc_show_mto.isChecked(),
            show_working=self._doc_show_working.isChecked(),
            show_as_build=self._doc_show_as_build.isChecked(),
        )

    def _current_ifc_revision_map(self) -> dict[tuple[str, str], str]:
        """Map kit identity to the current IFC revision from the worklist."""

        token = id(self._mto_worklist_rows)
        cached = getattr(self, "_ifc_revision_cache", None)
        if cached is not None and cached[0] == token:
            return cached[1]
        result = current_ifc_revision_map(self._mto_worklist_rows)
        self._ifc_revision_cache = (token, result)
        return result

    @Slot(bool)
    def _on_tree_label_option_changed(self, _checked: bool = False) -> None:
        """Persist label checkboxes and relabel existing tree nodes."""

        self._save_tree_label_options()
        self._relabel_document_tree()

    def _save_tree_label_options(self) -> None:
        for attr, key, _default in _TREE_LABEL_SETTINGS:
            box = getattr(self, attr, None)
            if isinstance(box, QCheckBox):
                self._settings.setValue(key, box.isChecked())

    def _restore_tree_label_options(self) -> None:
        for attr, key, default in _TREE_LABEL_SETTINGS:
            box = getattr(self, attr, None)
            if not isinstance(box, QCheckBox):
                continue
            box.blockSignals(True)
            box.setChecked(_settings_bool(self._settings, key, default))
            box.blockSignals(False)

    def _folder_hint_for(
        self,
        title: str,
        mark: str,
        files: list[FileRecord],
    ) -> FolderTreeHint | None:
        """Return the cached Google hint for one revision node, if any."""

        if not self._folder_hints or not files:
            return None
        identity = kit_identity_key(title, mark)
        names: list[str] = []
        display = folder_display_name(files)
        if display:
            names.append(display.casefold())
        names.append(record_folder_key(files[0]))
        package = issued_package_dir(files[0].path)
        if package:
            names.append(Path(package).name.casefold())
        seen: set[str] = set()
        for name in names:
            if not name or name in seen:
                continue
            seen.add(name)
            hint = self._folder_hints.get((*identity, name))
            if hint is not None:
                return hint
        return None

    def _contour_record_ids(self) -> set[int] | None:
        """Return cached contour file ids, or None when contour is unknown."""

        if self._contour_records is None:
            return None
        token = id(self._contour_records)
        cached = self._contour_ids_cache
        if cached is not None and cached[0] == token:
            return cached[1]
        ids = {record.id for record in self._contour_records}
        self._contour_ids_cache = (token, ids)
        return ids

    def _record_in_contour(self, record: FileRecord) -> bool:
        """Return whether ``record`` belongs to the canonical RD contour.

        Args:
            record: Catalog file row.

        Returns:
            Contour-id membership when ``_contour_records`` is set; otherwise
            a ``record_has_canonical_layout`` walk.
        """

        contour_ids = self._contour_record_ids()
        if contour_ids is not None:
            return record.id in contour_ids
        return record_has_canonical_layout(record, self.config.rd_root)

    def _overlay_id_in_contour(self, file_id: int) -> bool:
        """Return whether an overlay file id belongs to the catalog contour."""

        record = self._record_by_id.get(file_id)
        if record is None:
            return False
        return self._record_in_contour(record)

    def _stop_startup_timer(self) -> None:
        timer = getattr(self, "_startup_timer", None)
        if timer is not None:
            timer.stop()
        self._startup_load_pending = False

    def _stop_secondary_timer(self) -> None:
        timer = getattr(self, "_secondary_timer", None)
        if timer is not None:
            timer.stop()

    def _restore_deferred_load_timers(
        self, startup_pending: bool, secondary_pending: bool
    ) -> None:
        """Restart startup/secondary timers after a cancelled close."""

        if startup_pending:
            self._startup_load_pending = True
            timer = getattr(self, "_startup_timer", None)
            if timer is not None:
                timer.start(0)
            return
        if secondary_pending:
            self._secondary_tabs_pending = True
            self._deferred_widgets = set(_DEFERRED_ALL)

    def _startup_load(self) -> None:
        """Start the off-GUI snapshot thread after the first window show.

        Heatmap, MTO tables, collisions, approval lookups, and the
        documents tree stay empty until the user opens those tabs.
        """

        self._startup_load_running = True
        self._stop_startup_timer()
        if hasattr(self, "_kits_load_banner"):
            self._kits_load_banner.setVisible(True)
        self.statusBar().showMessage("Загрузка…")
        self._set_workers_enabled(False)
        self._update_action_states()
        thread = StartupHydrateThread(self.config, self)
        thread.log.connect(self._append_log)
        thread.finished.connect(self._on_startup_hydrate_finished)
        self._startup_thread = thread
        thread.start()

    @Slot()
    def _on_startup_hydrate_finished(self) -> None:
        """Apply the worker snapshot and paint Комплекты on the GUI thread."""

        thread = self._startup_thread
        self._startup_thread = None
        try:
            if thread is not None and thread.failure:
                self._append_log(f"Загрузка: {thread.failure}")
                self.refresh(
                    ingest_google=False,
                    rebuild_derived=False,
                    rebuild_document_tree=False,
                    defer_secondary=True,
                )
            elif thread is not None and thread.snapshot is not None:
                self._apply_startup_snapshot(thread.snapshot)
            else:
                self.refresh(
                    ingest_google=False,
                    rebuild_derived=False,
                    rebuild_document_tree=False,
                    defer_secondary=True,
                )
        finally:
            if thread is not None:
                thread.deleteLater()
            self._startup_load_running = False
            if hasattr(self, "_kits_load_banner"):
                self._kits_load_banner.setVisible(False)
            if not self._catalog_workers_busy():
                self._set_workers_enabled(True)
            self._update_action_states()
            if not self._catalog_workers_busy():
                self.statusBar().clearMessage()

    def _apply_startup_snapshot(self, snapshot: StartupSnapshot) -> None:
        """Install the worker snapshot and paint only the Комплекты table."""

        with perf_span("gui.apply_startup_snapshot"):
            google = snapshot.google
            self._assign_catalog_records(
                snapshot.records,
                contour=snapshot.contour_records or None,
            )
            self._detected_current_ids = snapshot.detected_current_ids
            self._issuance_kits_fresh = True
            self._mto_rows = snapshot.mto_rows
            self._collision_rows = snapshot.collision_rows
            self._google_kits = google.kits
            self._issuance_kits = google.issuance_kits
            self._issuance_sends = google.issuance_sends
            self._google_source = google.source
            self._google_fetched_at = google.fetched_at
            self._google_warning = google.warning
            self._google_loaded = google.loaded
            self._kit_rows = snapshot.kit_rows
            self._kit_pipelines = snapshot.kit_pipelines
            if snapshot.official_current_ids is not None:
                self._official_ids_value = set(snapshot.official_current_ids)
                self._official_ids_token = (
                    id(self._all_records),
                    frozenset(self._detected_current_ids),
                    id(self._kit_pipelines),
                )
            self._folder_hints = snapshot.folder_hints
            self._auto_mto_by_kit = snapshot.auto_mto_by_kit
            self._an_files_by_kit = snapshot.an_by_kit
            self._rd_dump_files = snapshot.rd_dump_files
            self._mto_content_by_kit = snapshot.mto_content_by_kit
            self._export_pins_by_key = snapshot.export_pins
            self._set_mto_worklist_rows(snapshot.worklist_rows)
            if hasattr(self, "_revision_matrix_tab"):
                self._revision_matrix_tab.seed_rd_mtimes(self._all_records)
                self._revision_matrix_tab.set_auto_mto_index(self._auto_mto_by_kit)
            self._apply_last_scan_info(snapshot.last_scan)
            self._deferred_widgets = set(_DEFERRED_ALL)
            self._secondary_tabs_pending = True
            self._refresh_kits_table(rebuild_rows=False)
            self._update_collision_tab_label()
            self._update_mto_sync_label()
            self._update_ban_action_label()
            self._update_action_states()
            if hasattr(self, "_tabs"):
                self._on_main_tab_changed(self._tabs.currentIndex())
            self._schedule_auto_mto_compares()
            self._schedule_transfer_mto_compares()

    def _apply_last_scan_info(self, scan: dict[str, Any] | None) -> None:
        if not hasattr(self, "_last_scan_label"):
            return
        if scan:
            value = str(scan.get("completed_at") or "").replace("T", " ")[:19]
            self._last_scan_label.setText(f"Последний успешный: {value}")
        else:
            self._last_scan_label.setText("Успешных сканов нет")

    def _ensure_deferred_widget(self, key: str) -> None:
        """Fill one hidden tab the first time the user opens it."""

        if key not in self._deferred_widgets:
            return
        if key == _DEFERRED_HEATMAP:
            self._refresh_revision_matrix(force=True)
        elif key == _DEFERRED_WORKLIST:
            self._refresh_mto_worklist(reload=False)
        elif key == _DEFERRED_MTO:
            self._refresh_mto_table()
        elif key == _DEFERRED_ISSUANCE_JOURNAL:
            self._refresh_issuance_journal()
        elif key == _DEFERRED_COLLISIONS:
            self._refresh_collision_table()
        elif key == _DEFERRED_APPROVAL:
            self._sync_approval_mail_lookups()
        elif key == _DEFERRED_TREE:
            self._rebuild_document_tree()
        elif key == _DEFERRED_AN:
            self._refresh_an_tab()
        elif key == _DEFERRED_RD_DUMP:
            self._refresh_rd_dump_tab()
        self._deferred_widgets.discard(key)
        if _DEFERRED_HEATMAP not in self._deferred_widgets and (
            _DEFERRED_TREE not in self._deferred_widgets
        ):
            self._secondary_tabs_pending = False
        if not self._deferred_widgets and not self._catalog_workers_busy():
            self.statusBar().clearMessage()

    def _load_secondary_tabs(self) -> None:
        """Fill heatmap and «Все документы» on first visit to those tabs."""

        self._stop_secondary_timer()
        self._ensure_deferred_widget(_DEFERRED_HEATMAP)
        self._ensure_deferred_widget(_DEFERRED_TREE)

    @Slot(int)
    def _on_main_tab_changed(self, index: int) -> None:
        """Load deferred widgets when the user opens those tabs."""

        if not hasattr(self, "_tabs"):
            return
        widget = self._tabs.widget(index)
        if widget is getattr(self, "_rd_dump_tab", None):
            if _DEFERRED_RD_DUMP in self._deferred_widgets:
                self._ensure_deferred_widget(_DEFERRED_RD_DUMP)
            elif self._rd_dump_tab.table().rowCount() == 0:
                self._refresh_rd_dump_tab()
            return
        if widget is getattr(self, "_an_tab", None):
            if _DEFERRED_AN in self._deferred_widgets:
                self._ensure_deferred_widget(_DEFERRED_AN)
            elif self._an_tab.table().rowCount() == 0:
                self._refresh_an_tab()
            return
        if not self._deferred_widgets:
            return
        if widget is getattr(self, "_revision_matrix_tab", None):
            self._ensure_deferred_widget(_DEFERRED_HEATMAP)
        elif widget is getattr(self, "_mto_worklist_tab", None):
            self._ensure_deferred_widget(_DEFERRED_WORKLIST)
        elif widget is getattr(self, "_mto_readiness_tab", None):
            self._ensure_deferred_widget(_DEFERRED_MTO)
        elif widget is getattr(self, "_issuance_journal_tab", None):
            self._ensure_deferred_widget(_DEFERRED_ISSUANCE_JOURNAL)
        elif widget is getattr(self, "_approval_mail_tab", None):
            self._ensure_deferred_widget(_DEFERRED_APPROVAL)
        elif widget is getattr(self, "_documents_tab", None):
            self._ensure_deferred_widget(_DEFERRED_TREE)
        elif widget is getattr(self, "_collision_tab", None):
            self._ensure_deferred_widget(_DEFERRED_COLLISIONS)

    def refresh(
        self,
        *,
        rebuild_document_tree: bool = True,
        ingest_google: bool = True,
        pipeline_subtrees: tuple[str, ...] | None = None,
        rebuild_derived: bool = True,
        defer_secondary: bool = False,
    ) -> None:
        """Reload both views from SQLite without touching source paths.

        Args:
            rebuild_document_tree: Rebuild «Все документы». Skip after a
                folder-scoped robot rescan — that tree is RD-only.
            ingest_google: Rewrite Google snapshot tables before pipeline.
                File scans pass ``False``.
            pipeline_subtrees: Folder prefixes from a scoped rescan. When
                set and kits are found under them, derived tables rebuild
                only for those identities, and Комплекты uses
                ``_refresh_kits_table(kit_keys=)`` (same as
                ``_refresh_after_scoped_pipeline``), not a full
                ``build_kit_matrix``.
            rebuild_derived: Persist Google snapshot and ``rebuild_pipeline``.
                Startup passes ``False``; scans keep the default ``True``.
            defer_secondary: Skip heatmap/export, MTO widgets, collisions,
                approval lookups, and (unless ``rebuild_document_tree``)
                the documents tree. A later tab visit fills them.
        """

        with perf_span(
            "gui.refresh",
            rebuild_document_tree=rebuild_document_tree,
            ingest_google=ingest_google,
            rebuild_derived=rebuild_derived,
            defer_secondary=defer_secondary,
            subtrees=len(pipeline_subtrees) if pipeline_subtrees else 0,
            kits="scoped" if pipeline_subtrees else "all",
        ):
            if not defer_secondary:
                self._stop_startup_timer()
                self._stop_secondary_timer()
                self._secondary_tabs_pending = False
                self._deferred_widgets.clear()
            self._assign_catalog_records(self.database.list_files())
            overlay_rows = self.database.current_overlay()
            self._detected_current_ids = {
                int(row["file_id"])
                for row in overlay_rows
                if self._overlay_id_in_contour(int(row["file_id"]))
            }
            self._load_mto_rows(overlay_rows)
            self._collision_rows = self.database.list_current_collisions()
            self._refresh_last_scan()
            kit_keys = None
            if pipeline_subtrees:
                kit_keys = kit_keys_under_folders(
                    self._all_records, pipeline_subtrees
                ) or None
            self._hydrate_google_from_cache()
            if not rebuild_derived and self._should_rebuild_empty_derived():
                rebuild_derived = True
                ingest_google = True
            if rebuild_derived:
                self._ingest_google_and_rebuild_pipeline(
                    ingest_google=ingest_google,
                    kit_keys=kit_keys,
                )
            self._load_folder_tree_hints()
            self._reload_mto_worklist_rows()
            if kit_keys is not None and self._kit_rows:
                self._refresh_kits_table(kit_keys=kit_keys)
            else:
                self._refresh_kits_table()
            if defer_secondary:
                self._secondary_tabs_pending = True
                self._deferred_widgets = set(_DEFERRED_ALL)
            else:
                self._refresh_revision_matrix()
                self._refresh_mto_table()
            if not defer_secondary:
                self._refresh_mto_worklist(reload=False)
                self._refresh_issuance_journal()
                self._refresh_collision_table()
                self._sync_approval_mail_lookups()
            else:
                self._update_collision_tab_label()
            if rebuild_document_tree:
                self._rebuild_document_tree()
                self._deferred_widgets.discard(_DEFERRED_TREE)
            self._update_ban_action_label()
            self._update_mto_sync_label()
            self._update_action_states()
            if not defer_secondary:
                # Full refresh paints heatmap/MTO/tree, but not dump finders.
                # Keep them lazy so a later tab click still loads SQLite.
                self._deferred_widgets.update({_DEFERRED_AN, _DEFERRED_RD_DUMP})
                if hasattr(self, "_tabs"):
                    self._on_main_tab_changed(self._tabs.currentIndex())

    def _load_mto_rows(self, overlay_rows: list[dict[str, Any]] | None = None) -> None:
        """Reload MTO comparison rows and synthetic not-compared placeholders.

        Args:
            overlay_rows: Already-fetched ``current_overlay`` rows. When
                omitted, the database is queried again.
        """

        with perf_span("gui.load_mto_rows"):
            if overlay_rows is None:
                overlay_rows = self.database.current_overlay()
            pending_error = (
                "Сверка выполняется…"
                if self._mto_compare_thread is not None
                else "Сверка RD ↔ robot ещё не выполнена"
            )
            self._mto_rows = build_mto_display_rows(
                self.database,
                self._record_by_id,
                overlay_rows,
                pending_error=pending_error,
            )

    def _refresh_mto_related(self) -> None:
        """Refresh MTO tab and kit robot-origin paint without a full reload."""

        with perf_span("gui.refresh_mto_related"):
            self._load_mto_rows()
            self._mto_content_by_kit = mto_content_equal_by_kit(self._mto_rows)
            mto_tab = getattr(self, "_mto_readiness_tab", None)
            visible = (
                hasattr(self, "_tabs")
                and mto_tab is not None
                and self._tabs.currentWidget() is mto_tab
            )
            if visible:
                self._deferred_widgets.discard(_DEFERRED_MTO)
                self._refresh_mto_table()
            else:
                self._deferred_widgets.add(_DEFERRED_MTO)
            self._refresh_kits_table(rebuild_rows=False)
            self._update_mto_sync_label()

    def _update_mto_sync_label(self) -> None:
        """Update the MTO tab sync badge from compare progress and row state."""

        if not hasattr(self, "_mto_sync_label"):
            return
        if self._mto_compare_thread is not None:
            completed, total = self._mto_progress
            suffix = f" ({completed}/{total})" if total else ""
            self._mto_sync_label.setText(f"Сверка: идёт{suffix}")
            return
        if self._mto_pending_keys or self._mto_queued is not None:
            self._mto_sync_label.setText("Сверка: в очереди")
        else:
            self._mto_sync_label.setText("Сверка: актуально")

    def _refresh_revision_matrix(
        self,
        *,
        kit_keys: set[tuple[str, str]] | None = None,
        force: bool = False,
    ) -> None:
        """Reload heatmap cells from SQLite; empty until schema 4 snapshot exists.

        Args:
            kit_keys: When set and the table is already painted, patch those
                kit rows via ``patch_cells_for_kits`` instead of rebuilding
                every column.
            force: Paint even when the heatmap tab is still deferred. First
                visit uses this; a one-kit write must not.
        """

        if not force and _DEFERRED_HEATMAP in self._deferred_widgets:
            return
        self._deferred_widgets.discard(_DEFERRED_HEATMAP)
        if not hasattr(self, "_revision_matrix_tab"):
            return
        with perf_span("gui.refresh_revision_matrix"):
            from rd_catalog import pipeline as pipeline_mod

            loader = getattr(pipeline_mod, "list_revision_matrix", None)
            cells: tuple[Any, ...] = ()
            if loader is not None:
                try:
                    cells = tuple(loader(self.database))
                except Exception as exc:
                    message = str(exc).casefold()
                    if "kit_revision_cell" not in message and "no such table" not in message:
                        self._append_log(f"Ревизии MTO: {type(exc).__name__}: {exc}")
            in_contour = self._contour_records is not None
            if kit_keys and self._revision_matrix_tab.patch_cells_for_kits(
                cells, kit_keys
            ):
                with perf_span("gui.heatmap_set_export_records"):
                    self._revision_matrix_tab.set_export_records(
                        self._pipeline_records(),
                        self._official_current_ids(),
                        pipelines=tuple(self._kit_pipelines.values()),
                        restart_compare=False,
                        in_contour=in_contour,
                        kit_keys=kit_keys,
                    )
                self._enqueue_auto_mto_compares_for_kits(kit_keys)
                return
            with perf_span("gui.document_tree_kit_identities"):
                allowed_kits = self._document_tree_kit_identities()
            with perf_span("gui.heatmap_set_cells"):
                self._revision_matrix_tab.set_cells(
                    cells,
                    palette=self._status_colors,
                    is_banned=lambda title, mark: self._is_banned_pair(title, mark),
                    allowed_kits=allowed_kits,
                )
            with perf_span("gui.heatmap_set_export_records"):
                self._revision_matrix_tab.set_export_records(
                    self._pipeline_records(),
                    self._official_current_ids(),
                    pipelines=tuple(self._kit_pipelines.values()),
                    in_contour=in_contour,
                )
            self._revision_matrix_tab.hydrate_auto_mto_compare_cache()
            self._schedule_auto_mto_compares()

    def _reload_mto_worklist_rows(self) -> None:
        """Reload MTO worklist rows and the per-kit rollup used by filters."""

        with perf_span("gui.reload_mto_worklist_rows"):
            self._set_mto_worklist_rows(
                list_mto_worklist(
                    self.database,
                    records=self._pipeline_records(),
                    rd_root=self._pipeline_rd_root(),
                )
            )

    def _set_mto_worklist_rows(self, rows: tuple[MtoWorklistRow, ...]) -> None:
        """Store worklist rows and rebuild the Комплекты filter rollup."""

        self._mto_worklist_rows = rows
        self._mto_kit_flags = build_mto_kit_flags(rows)

    def _refresh_mto_worklist(self, reload: bool = True) -> None:
        """Push cached MTO worklist rows into the Перечень tab.

        The worklist follows the Комплекты universe (RD ∪ Google ∪
        Выдача) and deliberately does not use «Все документы»
        membership.

        Args:
            reload: Re-query rows and kit flags. ``refresh`` passes False
                after ``_reload_mto_worklist_rows``.
        """

        if not hasattr(self, "_mto_worklist_tab"):
            return
        with perf_span("gui.refresh_mto_worklist"):
            self._deferred_widgets.discard(_DEFERRED_WORKLIST)
            if reload:
                self._reload_mto_worklist_rows()
            self._mto_worklist_tab.set_rows(
                self._mto_worklist_rows,
                palette=self._status_colors,
                is_banned=lambda title, mark: self._is_banned_pair(title, mark),
                allowed_kits=None,
                pin_view_for=self._pin_view_for,
                compare_status_for=self._auto_mto_compare_status_for,
            )
            self._schedule_auto_mto_compares()

    def _refresh_issuance_journal(self) -> None:
        """Reload «Выдача · Журнал» from SQLite and Auto MTO hits."""

        if not hasattr(self, "_issuance_journal_tab"):
            return
        with perf_span("gui.refresh_issuance_journal"):
            self._deferred_widgets.discard(_DEFERRED_ISSUANCE_JOURNAL)
            hits = [
                JournalAutoMtoHit(file.title, file.mark, file.rd_revision, file.relpath)
                for files in self._auto_mto_by_kit.values()
                for file in files
            ]
            try:
                rows = list_issuance_journal(
                    self.database,
                    records=self._all_records,
                    auto_mto_hits=hits,
                )
            except Exception as exc:
                self._append_log(f"Выдача · Журнал: {type(exc).__name__}: {exc}")
                rows = ()
            self._sheet_links = sheet_link_context_from_config(self.config)
            self._issuance_journal_tab.set_sheet_links(self._sheet_links)
            self._issuance_journal_tab.set_rows(
                rows,
                is_banned=lambda title, mark: self._is_banned_pair(title, mark),
                allowed_kits=None,
            )

    @Slot(object)
    def _on_issuance_reviews_changed(self, kit_keys: object) -> None:
        """Rebuild pipeline for kits whose issuance reviews just changed.

        Args:
            kit_keys: ``list[tuple[str, str]]`` of title/mark pairs.
        """

        if self._busy():
            QMessageBox.information(
                self,
                "Выдача · Журнал",
                "Статус записан. Таблицы обновятся после текущей "
                "загрузки или сканирования.",
            )
            return
        raw_keys = kit_keys if isinstance(kit_keys, (list, tuple, set)) else ()
        keys = {
            kit_identity_key(str(title), str(mark))
            for title, mark in raw_keys
            if title and mark
        }
        if not keys:
            return
        with perf_span("gui.issuance_reviews_changed"):
            try:
                rebuild_pipeline(
                    self.database,
                    records=self._pipeline_records(),
                    detected_current_ids=self._detected_current_ids,
                    kit_keys=keys,
                    rd_root=self._pipeline_rd_root(),
                )
                self._issuance_kits = latest_effective_issuance_kits(self.database)
                self._issuance_kits_fresh = True
            except Exception as exc:
                QMessageBox.warning(self, "Выдача · Журнал", str(exc))
                return
            selected = self._selected_kit_row()
            self._refresh_after_scoped_pipeline(keys)
            self._refresh_issuance_journal()
            if selected is not None:
                selected_key = kit_identity_key(selected.title, selected.mark)
                if selected_key in keys:
                    self._select_kit_row(selected.title, selected.mark)
                    self._update_kits_card()

    def _open_legalize_rd_dialog(self, row: KitMatrixRow) -> None:
        """Switch to the journal and pre-fill an RD legalize add-row dialog.

        Args:
            row: Selected комплекты matrix row.
        """

        self._ensure_deferred_widget(_DEFERRED_ISSUANCE_JOURNAL)
        self._tabs.setCurrentWidget(self._issuance_journal_tab)
        self._issuance_journal_tab.focus_kit(
            row.title, row.mark, issuance=row.issuance
        )
        revision = (row.rd.revision_text or "").strip()
        self._issuance_journal_tab.open_add_row_dialog(
            title=row.title,
            mark=row.mark,
            revision_text=revision,
            send_date=_issuance_date_from_mtime_ns(row.rd.max_mtime_ns),
            note="легализация РД",
            decision="legalized" if revision else "",
            lock_identity=True,
            window_title="Легализовать ревизию РД",
        )
        self._tabs.setCurrentWidget(self._kits_tab)
        self._select_kit_row(row.title, row.mark)

    def _on_issuance_journal_prepare_menu(self, menu: QMenu) -> None:
        """Inject ban into the issuance-journal context menu."""

        payload = self._issuance_journal_tab.selected_row()
        if payload is None:
            return
        ban = menu.addAction("Скрыть титул–марку (бан-фильтр)")
        ban.triggered.connect(
            lambda: self._confirm_ban_title_mark(payload.title, payload.mark)
        )

    def _on_revision_matrix_kit(self, title: str, mark: str) -> None:
        """Jump from the heatmap to the комплекты row for that kit."""

        self._tabs.setCurrentWidget(self._kits_tab)
        self._select_kit_row(title, mark)

    def _sync_approval_mail_lookups(self) -> None:
        """Point the mail tab at SQLite F text and issuance TRMs."""

        self._deferred_widgets.discard(_DEFERRED_APPROVAL)
        if not hasattr(self, "_approval_mail_tab"):
            return
        kits = self.database.list_google_kits()
        sends = self.database.list_issuance_sends()
        self._approval_mail_tab.set_lookups(
            comment_lookup=comment_lookup_from_kits(kits),
            kit_from_transmittal=kit_lookup_from_issuance(sends),
        )

    def _on_approval_mail_drop(self, mime: object) -> None:
        """Switch to «Письма о согласовании» and ingest the drop."""

        if not hasattr(self, "_approval_mail_tab"):
            return
        self._tabs.setCurrentWidget(self._approval_mail_tab)
        self._approval_mail_tab.ingest_mime(mime)

    @Slot(object)
    def _on_approval_mail_write(self, jobs: object) -> None:
        """Write letter F/D/E patches on a worker thread."""

        self._start_google_journal_write(
            tuple(jobs or ()), log_line="Запись писем в КСБ ИД…"
        )

    def _start_google_journal_write(
        self,
        jobs: object,
        *,
        log_line: str,
    ) -> None:
        """Run confirmed ``JournalWriteJob``s through ``GoogleFWriteThread``."""

        batch = tuple(jobs or ())
        if not batch:
            return
        if self._busy():
            QMessageBox.information(
                self,
                "КСБ ИД",
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return
        self._cancel_mto_compare(resume_later=True)
        self._cancel_export_pair_compare()
        thread = GoogleFWriteThread(self.config, batch, self)
        thread.log.connect(self._append_log)
        thread.error.connect(self._on_google_write_error)
        thread.finished.connect(self._on_google_write_finished)
        self._google_write_thread = thread
        self._set_workers_enabled(False)
        self._cancel_action.setEnabled(False)
        self._progress.setRange(0, 0)
        self._approval_mail_tab.set_catalog_busy(True)
        self._append_log(log_line)
        self.statusBar().showMessage("Запись в КСБ ИД…")
        thread.start()

    @Slot(str)
    def _on_google_write_error(self, message: str) -> None:
        self._append_log(message)
        self.statusBar().showMessage("Ошибка записи в КСБ ИД")

    @Slot()
    def _on_google_write_finished(self) -> None:
        thread = self._google_write_thread
        if thread is None:
            return
        with perf_span("gui.google_write_finished"):
            results = thread.results
            kits_result = thread.kits_result
            failure = thread.failure
            try:
                if hasattr(self, "_approval_mail_tab"):
                    self._approval_mail_tab.apply_write_results(results)
                self._apply_google_write_snapshot(results, kits_result)
            finally:
                self._google_write_thread = None
                thread.deleteLater()
                self._set_workers_enabled(True)
                self._progress.setRange(0, 1)
                self._progress.setValue(1)
            resumed = self._resume_pending_mto_compare()
            self._mto_resume_on_idle = (
                not resumed
                and bool(self._mto_pending_keys)
                and (self._mto_compare_thread is not None or self._busy())
            )
            ok = sum(1 for item in results if not item.error)
            bad = sum(1 for item in results if item.error)
            if not resumed:
                if failure and ok == 0:
                    self.statusBar().showMessage(f"КСБ ИД: {failure}", 10_000)
                else:
                    self.statusBar().showMessage(
                        f"КСБ ИД: записано {ok}, ошибок {bad}",
                        10_000,
                    )
            self._update_action_states()

    def _apply_google_write_snapshot(
        self,
        results: tuple[Any, ...],
        kits_result: Any,
    ) -> None:
        """Ingest the post-write КСБ ИД snapshot and refresh Комплекты only.

        Disk files did not change. Skips ``list_files``, the documents tree,
        heatmap paint, issuance journal, and MTO readiness. Rebuilds
        pipeline for kits whose F/D/E text moved (written jobs plus CSV
        diff, including concurrent sheet edits). Official overlay ids and
        the kit matrix still recompute because working vs official depends
        on the new F.

        Args:
            results: One write outcome per journal job.
            kits_result: ``fetch_google_kits(..., include_issuance=False)``
                payload, or ``None`` when the refetch was skipped.
        """

        written_keys = {
            kit_identity_key(item.title, item.mark)
            for item in results
            if not getattr(item, "error", "")
        }
        if kits_result is None or getattr(kits_result, "error", None):
            self._defer_secondary_after_google_write()
            return
        previous_kits = self._google_kits
        issuance_sends = kits_result.issuance_sends or self._issuance_sends
        self._google_kits = kits_result.kits
        self._issuance_sends = issuance_sends
        self._google_source = kits_result.source
        self._google_fetched_at = kits_result.fetched_at
        self._google_warning = kits_result.warning
        self._google_loaded = True
        try:
            ingest_google_snapshot(
                self.database,
                kits_result.kits,
                issuance_sends,
                loaded_at=kits_result.fetched_at
                or datetime.now(timezone.utc).isoformat(),
                source=kits_result.source,
                warning=kits_result.warning,
            )
            self._issuance_kits = latest_effective_issuance_kits(self.database)
            self._issuance_kits_fresh = True
        except Exception as exc:
            self._append_log(f"Google ingest: {type(exc).__name__}: {exc}")
            self._defer_secondary_after_google_write()
            return
        kit_keys = written_keys | changed_google_kit_keys(
            previous_kits, kits_result.kits
        )
        if not previous_kits:
            kit_keys = None
        try:
            rebuild_pipeline(
                self.database,
                records=self._pipeline_records(),
                detected_current_ids=self._detected_current_ids,
                kit_keys=kit_keys,
                rd_root=self._pipeline_rd_root(),
            )
        except Exception as exc:
            self._append_log(f"Pipeline: {type(exc).__name__}: {exc}")
        if hasattr(self, "_approval_mail_tab"):
            existing_lookup = self._approval_mail_tab._kit_lookup
            self._approval_mail_tab.set_lookups(
                comment_lookup=comment_lookup_from_kits(self._google_kits),
                kit_from_transmittal=existing_lookup,
            )
        self._reload_mto_worklist_rows()
        selected = self._selected_kit_row()
        self._refresh_kits_table()
        if selected is not None:
            self._select_kit_row(selected.title, selected.mark)
            self._update_kits_card()
        self._defer_secondary_after_google_write()

    def _defer_secondary_after_google_write(self) -> None:
        """Mark heatmap / journal / tree stale until the user opens them."""

        self._stop_startup_timer()
        self._stop_secondary_timer()
        self._secondary_tabs_pending = True
        self._deferred_widgets = set(_DEFERRED_ALL)
        self._deferred_widgets.discard(_DEFERRED_APPROVAL)

    def _refresh_last_scan(self) -> None:
        try:
            scan = self.database.last_scan_info(successful_only=True)
        except Exception:
            scan = None
        self._apply_last_scan_info(scan)

    def _refresh_mto_table(self) -> None:
        with perf_span("gui.refresh_mto_table"):
            self._deferred_widgets.discard(_DEFERRED_MTO)
            table = self._mto_table
            table.setSortingEnabled(False)
            table.setRowCount(len(self._mto_rows))
            visible_mto = [
                row for row in self._mto_rows if not self._is_banned_row(row)
            ]
            counts = Counter(
                str(row.get("status") or "").casefold() for row in visible_mto
            )
            pending = sum(
                str(row.get("rd_review_state")) == ReviewState.PENDING.value
                for row in visible_mto
            )
            for key in ("ready", "warning", "blocked"):
                title = key.upper()
                self._cards[key].setText(f"{title}\n{counts[key]}")
            self._cards["pending"].setText(f"Требует проверки\n{pending}")

            for table_row, row in enumerate(self._mto_rows):
                payload = row.get("diff") or {}
                values = [
                    str(row.get("status") or "").upper(),
                    row.get("title") or "—",
                    row.get("mark") or "—",
                    row.get("discipline_block") or "—",
                    _revision(row.get("rd_revision"), row.get("rd_appendix")),
                    _revision(row.get("robot_revision"), row.get("robot_appendix")),
                    "Да" if row.get("rd_is_as_build") else "Нет",
                    row.get("rd_review_state") or "—",
                    payload.get("content_status") or "—",
                    _mtime(row.get("rd_mtime_ns")),
                    _mtime(row.get("robot_mtime_ns")),
                    row.get("rd_path") or "—",
                    row.get("robot_path") or "—",
                    _compact_diff(row),
                ]
                for column, value in enumerate(values):
                    item = QTableWidgetItem(str(value))
                    item.setData(_ROLE_ROW, row)
                    if column in (11, 12):
                        path = row.get("rd_path" if column == 11 else "robot_path")
                        item.setToolTip(
                            _qt_tooltip(str(path or "Файл не найден в паре"))
                        )
                    if column == 0:
                        color = {
                            "ready": "#137333",
                            "warning": "#9a6700",
                            "blocked": "#b3261e",
                        }.get(str(row.get("status") or "").casefold(), "#202124")
                        item.setForeground(QBrush(QColor(color)))
                    table.setItem(table_row, column, item)
            table.setSortingEnabled(True)
            self._apply_mto_filter()

    def _official_current_ids(self) -> set[int]:
        """Return overlay-current file ids with working revisions removed."""

        token = (
            id(self._all_records),
            frozenset(self._detected_current_ids),
            id(self._kit_pipelines),
        )
        if self._official_ids_token == token:
            return self._official_ids_value
        with perf_span("gui.official_current_ids"):
            value = official_detected_current_ids(
                self._pipeline_records(),
                self._detected_current_ids,
                tuple(self._kit_pipelines.values()),
            )
        self._official_ids_token = token
        self._official_ids_value = value
        return value

    def _rebuild_kit_rows(self) -> None:
        with perf_span("gui.rebuild_kit_rows"):
            self._hydrate_google_from_cache()
            if self._issuance_kits_fresh:
                issuance_kits = self._issuance_kits
            else:
                try:
                    issuance_kits = latest_effective_issuance_kits(self.database)
                    self._issuance_kits_fresh = True
                except Exception:
                    issuance_kits = self._issuance_kits
            self._issuance_kits = issuance_kits
            self._kit_rows = build_kit_matrix(
                self._google_kits,
                self._pipeline_records(),
                self._official_current_ids(),
                issuance_kits=issuance_kits,
                rd_root=self._pipeline_rd_root(),
                mto_compare_by_pair=self._transfer_mto_by_pair,
                working_folders_by_kit=working_folders_from_pipelines(
                    self._kit_pipelines.values()
                ),
                annulled_folders_by_kit=annulled_folders_from_pipelines(
                    self._kit_pipelines.values()
                ),
            )
            self._mto_content_by_kit = mto_content_equal_by_kit(self._mto_rows)
            self._schedule_transfer_mto_compares()

    def _reload_export_pins(self) -> None:
        """Reload kit pins from ``runtime_dir`` JSON (never SQLite)."""

        self._export_pins_by_key = {
            kit_identity_key(pin.title, pin.mark): pin
            for pin in load_export_pins(self.config)
        }

    def _pin_view_for(self, title: str, mark: str) -> ExportPinView:
        """Return the display payload for the «Ручной выбор MTO» column."""

        key = kit_identity_key(title, mark)
        pin = self._export_pins_by_key.get(key)
        origin = ""
        rule_path = ""
        tab = getattr(self, "_revision_matrix_tab", None)
        if tab is not None:
            selection = tab.selection_for(title, mark)
            if selection is not None:
                origin = selection.origin
                rule_path = selection.rule_path
        return export_pin_view(pin, origin=origin, rule_path=rule_path)

    def _apply_pin_view(self, item: QTableWidgetItem, view: ExportPinView) -> None:
        """Write pin text, tooltip, and stale fill onto one table cell."""

        apply_export_pin_view(item, view, self._status_colors)

    def _on_export_pins_changed(self) -> None:
        """Repaint pin columns after a heatmap pin assign/clear."""

        selected = self._selected_kit_row()
        self._refresh_kits_table()
        if selected is not None:
            self._select_kit_row(selected.title, selected.mark)
        self._refresh_history()
        self._refresh_mto_worklist(reload=False)
        self._schedule_auto_mto_compares()

    def _live_heatmap_selections(
        self,
    ) -> dict[tuple[str, str], ExportSelection]:
        """Return heatmap export selections keyed by kit identity."""

        tab = getattr(self, "_revision_matrix_tab", None)
        if tab is None:
            return {}
        mapping: dict[tuple[str, str], ExportSelection] = {}
        for row in self._kit_rows:
            selection = tab.selection_for(row.title, row.mark)
            if selection is not None:
                mapping[kit_identity_key(row.title, row.mark)] = selection
        return mapping

    def _live_heatmap_comparisons(
        self,
    ) -> dict[tuple[str, str], tuple[str, AutoMtoCompareResult]] | None:
        """Return in-memory Auto MTO compares, or ``None`` to hydrate from disk."""

        tab = getattr(self, "_revision_matrix_tab", None)
        if tab is None:
            return None
        cached = getattr(tab, "_auto_mto_comparisons", None)
        if not cached:
            return None
        return dict(cached)

    def _reload_catalog_monitor(self) -> None:
        """Rebuild the Qt-free join payload used by Комплекты and the card."""

        with perf_span("gui.load_catalog_monitor"):
            try:
                self._catalog_monitor = load_catalog_monitor(
                    self.database,
                    self.config,
                    selections=self._live_heatmap_selections(),
                    comparisons=self._live_heatmap_comparisons(),
                )
            except Exception as exc:
                self._catalog_monitor = None
                self._append_log(f"Monitor join: {type(exc).__name__}: {exc}")

    def _kits_monitor_row_for(
        self,
        row: KitMatrixRow,
        *,
        overlay_mto: dict[tuple[str, str], tuple[str, str]] | None = None,
        ifc_by_kit: dict[tuple[str, str], str] | None = None,
        excluded_sends: tuple[Any, ...] = (),
    ) -> KitsMonitorRow:
        """Paint one Комплекты row from live window state."""

        key = kit_identity_key(row.title, row.mark)
        tab = getattr(self, "_revision_matrix_tab", None)
        selection = tab.selection_for(row.title, row.mark) if tab is not None else None
        comparison = (
            tab.comparison_record(row.title, row.mark) if tab is not None else None
        )
        return build_kits_monitor_row(
            row,
            pipeline=self._kit_pipelines.get(key),
            palette=self._status_colors,
            pins=self._export_pins_by_key,
            overlay_mto=overlay_mto
            if overlay_mto is not None
            else self._rd_mto_overlay_by_kit(),
            selection=selection,
            comparison=comparison,
            auto_files=self._auto_mto_by_kit.get(key, ()),
            an_files=self._an_files_by_kit.get(key, ()),
            mto_flags=self._mto_kit_flags.get(key, MtoKitFlags()),
            mto_content_equal=self._mto_content_by_kit.get(key),
            ifc_by_kit=ifc_by_kit
            if ifc_by_kit is not None
            else self._current_ifc_revision_map(),
            excluded_sends=excluded_sends,
            sheet_links=self._sheet_link_context(),
        )

    def _sheet_link_context(self) -> SheetLinkContext:
        """Return cached Google sheet ids/titles for cell jump URLs."""

        links = getattr(self, "_sheet_links", None)
        if links is None:
            self._sheet_links = sheet_link_context_from_config(self.config)
        return self._sheet_links

    def _refresh_kits_table(
        self,
        *,
        rebuild_rows: bool = True,
        kit_keys: set[tuple[str, str]] | None = None,
    ) -> None:
        """Repaint Комплекты from in-memory rows.

        Args:
            rebuild_rows: Reload pipelines and rebuild every ``KitMatrixRow``.
                Ignored when ``kit_keys`` is set.
            kit_keys: Rebuild and repaint only these identities. Does not
                call ``load_catalog_monitor``.
        """

        if not hasattr(self, "_kits_table"):
            return
        with perf_span(
            "gui.refresh_kits_table",
            rebuild_rows=rebuild_rows,
            kits=len(kit_keys) if kit_keys is not None else "all",
        ):
            if kit_keys is not None and self._kit_rows:
                self._rebuild_kit_rows_for_keys(kit_keys)
                self._paint_kits_table(kit_keys=kit_keys)
                return
            if rebuild_rows:
                self._reload_export_pins()
                try:
                    self._kit_pipelines = {
                        kit_identity_key(item.title, item.mark): item
                        for item in list_kit_pipelines(self.database)
                    }
                except Exception as exc:
                    self._kit_pipelines = {}
                    self._append_log(f"Pipeline list: {type(exc).__name__}: {exc}")
                self._rebuild_kit_rows()
            self._paint_kits_table()

    def _kits_paint_context(
        self,
    ) -> tuple[
        dict[tuple[str, str], tuple[str, str]],
        dict[tuple[str, str], str],
        dict[tuple[str, str], tuple[Any, ...]],
    ]:
        """Return overlay/IFC/excluded-send maps used to paint Комплекты."""
        overlay_mto = self._rd_mto_overlay_by_kit()
        ifc_by_kit = self._current_ifc_revision_map()
        try:
            excluded_by_kit = excluded_issuance_sends_by_kit(self.database)
        except Exception:
            excluded_by_kit = {}
        return overlay_mto, ifc_by_kit, excluded_by_kit

    def _write_kits_table_row(
        self,
        table_row: int,
        row: KitMatrixRow,
        *,
        overlay_mto: dict[tuple[str, str], tuple[str, str]],
        ifc_by_kit: dict[tuple[str, str], str],
        excluded_by_kit: dict[tuple[str, str], tuple[Any, ...]],
        reuse: bool = False,
    ) -> None:
        """Write one Комплекты matrix row onto the table.

        Args:
            table_row: QTableWidget row index.
            row: Kit matrix payload.
            overlay_mto: Official MTO path/rev by identity.
            ifc_by_kit: Latest IFC revision by identity.
            excluded_by_kit: Excluded issuance sends by identity.
            reuse: Update existing ``QTableWidgetItem``s when present.
        """

        table = self._kits_table
        key = kit_identity_key(row.title, row.mark)
        painted = self._kits_monitor_row_for(
            row,
            overlay_mto=overlay_mto,
            ifc_by_kit=ifc_by_kit,
            excluded_sends=excluded_by_kit.get(key, ()),
        )
        for column, header in enumerate(_KITS_HEADERS):
            cell = painted.cells.get(header, MonitorCell(text="—"))
            item = table.item(table_row, column) if reuse else None
            if item is None:
                item = (
                    _KitsSortItem()
                    if header == KITS_OK_HEADER
                    else QTableWidgetItem()
                )
                table.setItem(table_row, column, item)
            elif reuse:
                item.setBackground(QBrush())
                item.setForeground(QBrush())
                item.setToolTip("")
                font = item.font()
                font.setBold(False)
                font.setUnderline(False)
                item.setFont(font)
            apply_monitor_cell(item, cell)
            if header == PIN_COLUMN_HEADER:
                item.setText(cell.text)
            item.setData(_ROLE_ROW, row)
            item.setData(_ROLE_MONITOR, painted)

    def _paint_kits_table(
        self, *, kit_keys: set[tuple[str, str]] | None = None
    ) -> None:
        """Paint Комплекты rows; optionally only ``kit_keys``.

        Args:
            kit_keys: When set, patch matching table rows in place. Falls
                back to a full ``setRowCount`` paint if a key is missing
                from the table or disappeared from ``_kit_rows``.
        """

        overlay_mto, ifc_by_kit, excluded_by_kit = self._kits_paint_context()
        self._sheet_links = sheet_link_context_from_config(self.config)
        table = self._kits_table
        if kit_keys:
            wanted = {kit_identity_key(title, mark) for title, mark in kit_keys}
            rows_by_key = {
                kit_identity_key(row.title, row.mark): row for row in self._kit_rows
            }
            table_index: dict[tuple[str, str], int] = {}
            for table_row in range(table.rowCount()):
                item = table.item(table_row, 0)
                payload = item.data(_ROLE_ROW) if item else None
                if isinstance(payload, KitMatrixRow):
                    table_index[kit_identity_key(payload.title, payload.mark)] = (
                        table_row
                    )
            missing = [key for key in wanted if key not in table_index]
            stale = [
                key
                for key in table_index
                if key in wanted and key not in rows_by_key
            ]
            if not missing and not stale:
                with perf_span("gui.kits_table_paint", rows=len(wanted)):
                    table.setSortingEnabled(False)
                    try:
                        for key in wanted:
                            row = rows_by_key.get(key)
                            if row is None:
                                continue
                            self._write_kits_table_row(
                                table_index[key],
                                row,
                                overlay_mto=overlay_mto,
                                ifc_by_kit=ifc_by_kit,
                                excluded_by_kit=excluded_by_kit,
                                reuse=True,
                            )
                    finally:
                        table.setSortingEnabled(True)
                self._apply_kits_filter()
                self._update_kits_tab_label()
                if table.rowCount() == 0:
                    self._clear_kits_detail_panels()
                return
        with perf_span("gui.kits_table_paint", rows=len(self._kit_rows)):
            table.setSortingEnabled(False)
            table.setRowCount(len(self._kit_rows))
            for table_row, row in enumerate(self._kit_rows):
                self._write_kits_table_row(
                    table_row,
                    row,
                    overlay_mto=overlay_mto,
                    ifc_by_kit=ifc_by_kit,
                    excluded_by_kit=excluded_by_kit,
                )
            table.setSortingEnabled(True)
        self._apply_kits_filter()
        self._update_kits_tab_label()
        if table.rowCount() == 0:
            self._clear_kits_detail_panels()

    def _rebuild_kit_rows_for_keys(self, kit_keys: set[tuple[str, str]]) -> None:
        """Reload pipelines and ``KitMatrixRow``s for ``kit_keys`` only.

        Args:
            kit_keys: Identities whose derived pipeline just changed.
        """

        wanted = {kit_identity_key(title, mark) for title, mark in kit_keys}
        for title, mark in wanted:
            try:
                rows = self.database.list_kit_pipelines(title, mark)
            except Exception as exc:
                self._append_log(f"Pipeline list: {type(exc).__name__}: {exc}")
                continue
            if not rows:
                self._kit_pipelines.pop((title, mark), None)
                continue
            for item in rows:
                self._kit_pipelines[kit_identity_key(item.title, item.mark)] = item
        if isinstance(self._official_ids_value, set) and self._official_ids_value:
            self._official_ids_value = patch_official_detected_current_ids(
                self._official_ids_value,
                records=self._pipeline_records(),
                detected_current_ids=self._detected_current_ids,
                pipelines=tuple(self._kit_pipelines.values()),
                kit_keys=wanted,
            )
            self._official_ids_token = (
                id(self._all_records),
                frozenset(self._detected_current_ids),
                id(self._kit_pipelines),
            )
        else:
            self._official_ids_token = None
            self._official_current_ids()
        google = [
            kit
            for kit in self._google_kits
            if kit_identity_key(kit.title, kit.mark) in wanted
        ]
        if self._issuance_kits_fresh:
            issuance_kits = self._issuance_kits
        else:
            try:
                issuance_kits = latest_effective_issuance_kits(self.database)
                self._issuance_kits_fresh = True
            except Exception:
                issuance_kits = self._issuance_kits
            self._issuance_kits = issuance_kits
        issuance = [
            kit
            for kit in issuance_kits
            if kit_identity_key(kit.title, kit.mark) in wanted
        ]
        records = [
            record
            for record in self._pipeline_records()
            if kit_identity_key(
                str(record.data.get("title") or "").strip(),
                str(record.data.get("mark") or "").strip(),
            )
            in wanted
        ]
        scoped_rows = build_kit_matrix(
            google,
            records,
            self._official_current_ids(),
            issuance_kits=issuance,
            rd_root=self._pipeline_rd_root(),
            mto_compare_by_pair=self._transfer_mto_by_pair,
            working_folders_by_kit=working_folders_from_pipelines(
                self._kit_pipelines.values()
            ),
            annulled_folders_by_kit=annulled_folders_from_pipelines(
                self._kit_pipelines.values()
            ),
        )
        by_key = {
            kit_identity_key(row.title, row.mark): row for row in scoped_rows
        }
        new_rows: list[KitMatrixRow] = []
        seen: set[tuple[str, str]] = set()
        for row in self._kit_rows:
            key = kit_identity_key(row.title, row.mark)
            if key in wanted:
                replacement = by_key.get(key)
                new_rows.append(replacement if replacement is not None else row)
                seen.add(key)
            else:
                new_rows.append(row)
        for key, row in by_key.items():
            if key not in seen:
                new_rows.append(row)
        self._kit_rows = tuple(new_rows)

    @Slot(bool)
    def _on_kits_as_build_toggled(self, checked: bool) -> None:
        """Keep the two as-build filters mutually exclusive, then refilter."""

        other = (
            self._kits_only_as_build
            if self.sender() is self._kits_no_as_build
            else self._kits_no_as_build
        )
        other.setEnabled(not checked)
        if checked and other.isChecked():
            other.blockSignals(True)
            other.setChecked(False)
            other.blockSignals(False)
        self._apply_kits_filter()

    @Slot()
    def _apply_kits_filter(self) -> None:
        if not hasattr(self, "_kits_table"):
            return
        needle = self._kits_filter.text().strip().casefold()
        mismatch_only = self._kits_mismatch.isChecked()
        require_no_rd = self._kits_no_rd.isChecked()
        require_no_robot = self._kits_no_robot.isChecked()
        require_no_google = self._kits_no_google.isChecked()
        require_code_a = self._kits_code_a.isChecked()
        require_tdo = self._kits_tdo.isChecked()
        hide_as_build = self._kits_no_as_build.isChecked()
        require_as_build = self._kits_only_as_build.isChecked()
        require_mto_problems = self._kits_mto_problems.isChecked()
        require_an_closes = (
            hasattr(self, "_kits_an_closes") and self._kits_an_closes.isChecked()
        )
        progress_rows: list[KitsMonitorRow] = []
        visible_count = 0
        for row_index in range(self._kits_table.rowCount()):
            item = self._kits_table.item(row_index, 0)
            painted = item.data(_ROLE_MONITOR) if item else None
            row: KitMatrixRow | None = item.data(_ROLE_ROW) if item else None
            if not isinstance(painted, KitsMonitorRow):
                if isinstance(row, KitMatrixRow):
                    painted = self._kits_monitor_row_for(row)
                else:
                    self._kits_table.setRowHidden(row_index, True)
                    continue
            visible = not needle or needle in painted.haystack
            if mismatch_only and painted.summary_aligned:
                visible = False
            if require_no_rd and painted.rd_present:
                visible = False
            if require_no_robot and painted.robot_present:
                visible = False
            if require_no_google and not painted.has_google_or_issuance:
                visible = False
            if require_code_a and not painted.code_a:
                visible = False
            if require_tdo and not painted.kit_tdo_passed:
                visible = False
            if hide_as_build and painted.as_build:
                visible = False
            if require_as_build and not painted.as_build:
                visible = False
            if require_mto_problems and not painted.has_mto_problem:
                visible = False
            if require_an_closes and not painted.an_closes_auto_mto:
                visible = False
            banned = row is not None and self._is_banned_pair(
                row.title, row.mark, row.title_system
            )
            if banned:
                visible = False
            self._kits_table.setRowHidden(row_index, not visible)
            if not banned:
                progress_rows.append(painted)
                if visible:
                    visible_count += 1
        self._update_kits_progress_label(progress_rows, visible_count)

    def _update_kits_progress_label(
        self,
        rows: Sequence[KitsMonitorRow],
        visible: int,
    ) -> None:
        """Bind the Комплекты header progress line (no extra filter row)."""

        if not hasattr(self, "_kits_progress"):
            return
        stats = kits_progress_stats(rows)
        text = format_kits_progress_stats(stats, visible=visible)
        self._kits_progress.setText(text)

    @Slot(str)
    def _on_kits_progress_copied(self, text: str) -> None:
        """Confirm the Комплекты progress line was copied."""

        self.statusBar().showMessage(f"Скопировано: {text}", 3000)

    def _selected_kit_row(self) -> KitMatrixRow | None:
        if not hasattr(self, "_kits_table"):
            return None
        row = self._kits_table.currentRow()
        item = self._kits_table.item(row, 0) if row >= 0 else None
        payload = item.data(_ROLE_ROW) if item else None
        return payload if isinstance(payload, KitMatrixRow) else None

    def _selected_kits_monitor_row(self) -> KitsMonitorRow | None:
        """Return the painted Комплекты row stored on the selected table item."""

        if not hasattr(self, "_kits_table"):
            return None
        row = self._kits_table.currentRow()
        item = self._kits_table.item(row, 0) if row >= 0 else None
        payload = item.data(_ROLE_MONITOR) if item else None
        return payload if isinstance(payload, KitsMonitorRow) else None

    def _update_kits_tooltips(self) -> None:
        if not hasattr(self, "_kits_tooltips"):
            return
        painted = self._selected_kits_monitor_row()
        if painted is None:
            self._kits_tooltips.clear()
            return
        text = painted.tooltips_text or format_kits_row_tooltips(
            painted.title, painted.mark, painted.cells
        )
        self._kits_tooltips.setPlainText(
            format_kits_tips_pane(text, with_ctrl_click_hint=True)
        )

    def _on_kits_table_ctrl_click(self, row: int, column: int) -> None:
        """Select the kit row and scroll Подсказки to that column's block."""

        self._kits_table.selectRow(row)
        header, _ = self._table_click_cell(self._kits_table, row, column)
        if hasattr(self, "_kits_detail_tabs"):
            self._kits_detail_tabs.setCurrentIndex(0)
        if not hasattr(self, "_kits_tooltips"):
            return
        edit = self._kits_tooltips
        scroll_kits_tips_header_to_top(edit, header)
        QTimer.singleShot(0, lambda: scroll_kits_tips_header_to_top(edit, header))

    @Slot()
    def _on_kit_selected(self) -> None:
        self._update_kits_panels()
        self._update_action_states()

    def _update_kits_panels(self) -> None:
        self._update_kits_tooltips()
        self._update_kits_detail()
        self._update_kits_card()

    def _clear_kits_detail_panels(self) -> None:
        if hasattr(self, "_kits_tooltips"):
            self._kits_tooltips.clear()
        if hasattr(self, "_kits_detail"):
            self._kits_detail.clear()
        self._clear_kits_card()

    def _kit_display_summary_label(self, row: KitMatrixRow) -> str:
        """Return the Сводка text shown on Комплекты (missing-official → Нет в РД)."""

        pipeline = self._kit_pipelines.get(kit_identity_key(row.title, row.mark))
        return summary_label(effective_kit_summary(row, pipeline))

    def _update_kits_detail(self) -> None:
        if not hasattr(self, "_kits_detail"):
            return
        row = self._selected_kit_row()
        if row is None:
            self._kits_detail.clear()
            return
        lines = [
            f"{row.title_system} · {self._kit_display_summary_label(row)}",
            f"Флаги: {', '.join(flag.value for flag in row.flags) or '—'}",
            "",
        ]
        origin = self._robot_origin_for(row)
        if origin.reason:
            lines.append(origin.reason)
            lines.append("")
        if row.mixed_title_notes:
            lines.append("Смешанные титулы:")
            lines.append("")
            lines.extend(row.mixed_title_notes)
            lines.append("")
        if row.transfer_review_notes:
            lines.append("Проверить передачи:")
            lines.append("")
            lines.extend(row.transfer_review_notes)
            lines.append("")
        google = row.google
        if google is None:
            lines.append("Google: нет строки в таблице комплектов.")
        else:
            lines.append(
                f"Google: рев. {google.sheet_revision_text or '—'} · "
                f"{google.status_sheet or '—'}"
            )
            if google.events:
                lines.append("Google · история (столбец F):")
                for event in google.events:
                    extra = " ".join(event.transmittals)
                    rev = format_revision(event.revision, event.appendix)
                    lines.append(
                        f"  {event.date or '—'}  {event.stage_label}"
                        f"{'  ' + rev if rev else ''}"
                        f"{'  ' + extra if extra else ''}"
                    )
            else:
                lines.append("Google · история F: пусто")
        issuance = row.issuance
        lines.append("")
        if issuance is None:
            lines.append("Выдача РД ПД: нет строки.")
        else:
            lines.append(
                f"Выдача РД ПД: рев. {issuance.revision_text or '—'} · "
                f"{issuance.status or '—'}"
            )
            lines.append(
                f"  отправка {issuance.send_date_sortable or issuance.send_date or '—'} · "
                f"{issuance.send_transmittal or '—'}"
            )
            lines.append(
                f"  вх.контр. {issuance.incoming_control_date_sortable or issuance.incoming_control_date or '—'} · "
                f"подтв. {issuance.confirm_transmittal or '—'}"
            )
            if issuance.note_raw:
                lines.append(f"  примечание: {issuance.note_raw}")
        lines.append("")
        for label, snapshot in (
            ("РД", row.rd),
            ("Робот", row.robot),
            ("SQ", row.sq),
        ):
            if not snapshot.present:
                lines.append(f"{label}: нет файлов")
                continue
            as_built = " · as-built" if snapshot.as_build else ""
            transfer = (
                f" · {snapshot.transfer_name}" if snapshot.transfer_name else ""
            )
            extra_mto = ""
            if (
                label == "РД"
                and snapshot.mto_revision_text
                and snapshot.mto_revision_text != snapshot.revision_text
            ):
                extra_mto = f" · MTO {snapshot.mto_revision_text}"
            lines.append(
                f"{label}: {snapshot.file_count} файл(ов), "
                f"рев. {snapshot.revision_text or '—'}{extra_mto}{as_built}{transfer}"
            )
            for path in snapshot.paths:
                lines.append(f"  {path}")
        self._kits_detail.setPlainText("\n".join(lines))

    def _clear_kits_card(self) -> None:
        if not hasattr(self, "_kits_card_title"):
            return
        self._kits_card_title.setText("Выберите комплект")
        self._kits_card_summary.setText("")
        self._kits_card_pipeline.setText("")
        self._kits_card_pipeline.setStyleSheet("")
        self._kits_card_pipeline.setToolTip("")
        self._kits_card_approval.setText("")
        self._kits_card_approval.setStyleSheet("")
        self._kits_card_approval.setToolTip("")
        self._kits_card_revs.setText("")
        if hasattr(self, "_kits_card_mto"):
            self._kits_card_mto.setText("")
        self._kits_card_sources.setText("")
        if hasattr(self, "_kits_package_table"):
            self._kits_package_table.setRowCount(0)
        self._kits_last_card = None

    def _cached_kit_card(self, title: str, mark: str):
        """Return the last painted kit card when it is still the same kit."""

        cached = getattr(self, "_kits_last_card", None)
        if cached is None:
            return None
        cached_title, cached_mark, card = cached
        if kit_identity_key(cached_title, cached_mark) != kit_identity_key(
            title, mark
        ):
            return None
        return card

    def _update_kits_card(self) -> None:
        if not hasattr(self, "_kits_package_table"):
            return
        row = self._selected_kit_row()
        if row is None:
            self._clear_kits_card()
            return
        with perf_span("gui.update_kits_card"):
            try:
                card = get_kit_card(self.database, row.title, row.mark)
            except Exception as exc:
                self._append_log(f"Карточка комплекта: {type(exc).__name__}: {exc}")
                card = None
            painted = None
            if self._catalog_monitor is not None:
                painted = self._catalog_monitor.kit_card(row.title, row.mark)
            if painted is None:
                painted = paint_kit_card(
                    database=self.database,
                    title=row.title,
                    mark=row.mark,
                    kit_row=row,
                    worklist_rows=self._mto_worklist_rows,
                    palette=self._status_colors,
                )
            title = (
                painted.title
                if painted is not None
                else (card.title if card is not None else row.title)
            )
            mark = (
                painted.mark
                if painted is not None
                else (card.mark if card is not None else row.mark)
            )
            self._kits_card_title.setText(f"{title}-{mark}")
            self._kits_card_summary.setText(self._kit_display_summary_label(row))
            if painted is not None:
                _style_kits_status_badge(
                    self._kits_card_pipeline,
                    painted.review.text,
                    painted.review.fill,
                )
                self._kits_card_pipeline.setToolTip(painted.review.tooltip)
                _style_kits_status_badge(
                    self._kits_card_approval,
                    painted.approval.text,
                    painted.approval.fill,
                )
                self._kits_card_approval.setToolTip(painted.approval.tooltip)
                working = painted.working_revision_text
                official = painted.official_revision_text
                self._kits_card_mto.setText(painted.mto_rollup)
            else:
                pipeline = card.pipeline if card is not None else None
                if pipeline is None:
                    pipeline = self._kit_pipelines.get(
                        kit_identity_key(row.title, row.mark)
                    )
                if pipeline is not None:
                    review_label = pipeline_review_label(
                        pipeline,
                        events=(
                            row.google.events if row.google is not None else ()
                        ),
                        issuance=row.issuance,
                        google=row.google,
                    )
                    review_color = color_for(
                        self._status_colors,
                        pipeline_display_review_status(
                            pipeline,
                            google=row.google,
                            issuance=row.issuance,
                        ),
                    )
                    _style_kits_status_badge(
                        self._kits_card_pipeline, review_label, review_color
                    )
                    review_hints = [review_label]
                    displayed_working = format_working_rd_rev_label(
                        pipeline.working_revision_text,
                        pipeline.official_revision_text,
                        working_as_build=pipeline.working_as_build,
                    )
                    if displayed_working:
                        review_hints.append(f"раб. {displayed_working}")
                    self._kits_card_pipeline.setToolTip(" · ".join(review_hints))
                    approval_label = pipeline_approval_label(
                        pipeline, google=row.google, issuance=row.issuance
                    )
                    approval_key = pipeline_approval_color_key(
                        pipeline, google=row.google, issuance=row.issuance
                    )
                    approval_color = (
                        color_for(self._status_colors, approval_key)
                        if approval_key
                        else None
                    )
                    _style_kits_status_badge(
                        self._kits_card_approval, approval_label, approval_color
                    )
                    self._kits_card_approval.setToolTip("")
                else:
                    _style_kits_status_badge(self._kits_card_pipeline, "—", None)
                    self._kits_card_pipeline.setToolTip("")
                    _style_kits_status_badge(self._kits_card_approval, "—", None)
                    self._kits_card_approval.setToolTip("")
                official = (
                    (card.official_revision_text if card is not None else "")
                    or (pipeline.official_revision_text if pipeline is not None else "")
                    or row.rd.revision_text
                )
                working = format_working_rd_rev_label(
                    (card.working_revision_text if card is not None else "")
                    or (pipeline.working_revision_text if pipeline is not None else ""),
                    official,
                    working_as_build=bool(
                        pipeline is not None and pipeline.working_as_build
                    ),
                )
                self._kits_card_mto.setText(self._format_kits_card_mto(row))
            rev_parts: list[str] = []
            if working:
                rev_parts.append(f"Рабочая рев. {working}")
            if official:
                rev_parts.append(f"Офиц. рев. {official}")
            self._kits_card_revs.setText(" · ".join(rev_parts))
            google = card.google if card is not None else row.google
            issuance = card.issuance if card is not None else row.issuance
            google_bit = (
                f"Google D {google.sheet_revision_text or '—'} · "
                f"E {google.status_sheet or '—'}"
                if google is not None
                else "Google D/E: —"
            )
            issuance_bit = (
                f"Выдача {issuance.revision_text or '—'} · "
                f"{issuance.status or '—'} · "
                f"{issuance.send_date_sortable or issuance.send_date or '—'}"
                if issuance is not None
                else "Выдача: —"
            )
            packages = tuple(card.packages) if card is not None else ()
            source_bits = [
                google_bit,
                issuance_bit,
                f"РД: {self._format_source_package_paths(packages, row, 'rd')}",
                f"SQ: {self._format_source_package_paths(packages, row, 'sq')}",
                f"робот: {self._format_source_package_paths(packages, row, 'robot')}",
            ]
            self._kits_card_sources.setText(" · ".join(source_bits))
            self._fill_kits_package_table(row, card, painted)
            self._kits_last_card = (row.title, row.mark, card)

    def _format_source_package_paths(
        self,
        packages: tuple[KitPackageRow, ...],
        row: KitMatrixRow,
        source: str,
    ) -> str:
        paths: list[str] = []
        seen: set[str] = set()
        for package in packages:
            if package.source != source or package.is_grey:
                continue
            path = (package.package_path or "").strip()
            key = path.casefold()
            if path and key not in seen:
                seen.add(key)
                paths.append(path)
        if not paths:
            snapshot = self._kit_snapshot_for(row, source)
            for file_path in snapshot.paths if snapshot else ():
                package = issued_package_dir(str(file_path))
                key = package.casefold()
                if package and key not in seen:
                    seen.add(key)
                    paths.append(package)
        if not paths:
            return "—"
        if len(paths) == 1:
            return paths[0]
        return f"{paths[0]} (+{len(paths) - 1})"

    def _fill_kits_package_table(
        self,
        row: KitMatrixRow,
        card: KitCard | None,
        painted: KitCardMonitor | None = None,
    ) -> None:
        table = self._kits_package_table
        table.setSortingEnabled(False)
        table.setRowCount(0)
        if card is None or painted is None:
            table.setSortingEnabled(True)
            return
        package_by_path = {
            (package.package_path or "").casefold(): package
            for package in self._sorted_card_packages(card.packages)
        }
        sorted_packages = self._sorted_card_packages(card.packages)
        for index, painted_pkg in enumerate(painted.packages):
            package = package_by_path.get(
                (painted_pkg.package_path or "").casefold()
            )
            if package is None and index < len(sorted_packages):
                package = sorted_packages[index]
            table_row = table.rowCount()
            table.insertRow(table_row)
            for column, header in enumerate(CARD_PACKAGE_HEADERS):
                cell = painted_pkg.cells.get(header, MonitorCell(text="—"))
                item = QTableWidgetItem()
                apply_monitor_cell(item, cell)
                item.setData(_ROLE_PACKAGE, package)
                item.setData(_ROLE_ROW, row)
                item.setTextAlignment(
                    Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop
                )
                table.setItem(table_row, column, item)
        table.resizeRowsToContents()
        table.setSortingEnabled(True)
        for table_row in range(table.rowCount()):
            item = table.item(table_row, 0)
            package = item.data(_ROLE_PACKAGE) if item is not None else None
            if isinstance(package, KitPackageRow) and package.is_current:
                table.selectRow(table_row)
                break

    @staticmethod
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

    @staticmethod
    def _package_liquidity_label(package: KitPackageRow, card: KitCard) -> str:
        return package_liquidity_label(package, card)

    def _format_kits_card_mto(self, row: KitMatrixRow) -> str:
        """One-line MTO rollup for the selected kit from the worklist."""

        return format_kits_card_mto(self._mto_worklist_rows, row)

    def _reload_kits_card_legend(self) -> None:
        """Rebuild the card color legend from the current palette."""

        layout = getattr(self, "_kits_card_legend", None)
        if layout is None:
            return
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget() if item is not None else None
            if widget is not None:
                widget.deleteLater()
        parent = self._kits_card_title
        for key in _KITS_CARD_LEGEND_KEYS:
            sample = QLabel("  ", parent)
            hex_color = color_for(self._status_colors, key)
            sample.setStyleSheet(
                f"background:{hex_color}; border:1px solid #80868b; "
                f"min-width:16px;"
            )
            caption = QLabel(status_color_label(key), parent)
            layout.addWidget(sample)
            layout.addWidget(caption)
        layout.addStretch(1)

    def _selected_kit_package(self) -> KitPackageRow | None:
        if not hasattr(self, "_kits_package_table"):
            return None
        row = self._kits_package_table.currentRow()
        item = self._kits_package_table.item(row, 0) if row >= 0 else None
        payload = item.data(_ROLE_PACKAGE) if item else None
        return payload if isinstance(payload, KitPackageRow) else None

    def _open_selected_kit_package(self) -> None:
        package = self._selected_kit_package()
        if package is None or package.is_grey or not package.package_path:
            return
        self._open_result(package.package_path, folder=False)

    def _copy_selected_kit_package_path(self) -> None:
        package = self._selected_kit_package()
        if package is None or not package.package_path:
            return
        self._copy_text(package.package_path)

    def _jump_selected_kit_package_documents(self) -> None:
        package = self._selected_kit_package()
        kit = self._selected_kit_row()
        if package is None or kit is None:
            return
        self._jump_to_kit_documents(
            kit.title, kit.mark, transfer_name=package.transfer_name
        )

    def _copy_kit_package_handoff(self, column: int | None = None) -> None:
        table = self._kits_package_table
        row_index = table.currentRow()
        if row_index < 0:
            return
        if column is None:
            column = max(table.currentColumn(), 0)
        self._copy_robot_handoff(
            self._kit_package_handoff_text(row_index, column)
        )

    def _show_kit_package_context_menu(self, position) -> None:
        table = self._kits_package_table
        row_index = table.rowAt(position.y())
        if row_index < 0:
            return
        table.selectRow(row_index)
        package = self._selected_kit_package()
        kit = self._selected_kit_row()
        if package is None or kit is None:
            return
        liquidity = "—"
        card = self._cached_kit_card(kit.title, kit.mark)
        if card is None:
            try:
                card = get_kit_card(self.database, kit.title, kit.mark)
            except Exception:
                card = None
        if card is not None:
            liquidity = self._package_liquidity_label(package, card)
        menu = QMenu(self)
        open_action = menu.addAction("Открыть папку")
        open_action.setEnabled(bool(package.package_path) and not package.is_grey)
        jump_action = menu.addAction("Показать во Все документы (этот NN)")
        confirm_ok = None
        confirm_bad = None
        actionable = liquidity == "pending" or (
            card is not None
            and card.pipeline is not None
            and card.pipeline.suspicious
            and not package.is_grey
            and package.source == "rd"
        )
        if actionable:
            menu.addSeparator()
            confirm_ok = menu.addAction("Подтвердить корректность")
            confirm_bad = menu.addAction("Подтвердить неликвидность")
        handoff_action = self._add_robot_handoff_action(menu)
        chosen = exec_tracked_menu(
            menu, MENU_KIT_PACKAGE, table.viewport().mapToGlobal(position)
        )
        if chosen == open_action:
            self._open_selected_kit_package()
        elif chosen == jump_action:
            self._jump_selected_kit_package_documents()
        elif chosen == confirm_ok:
            self._confirm_package_liquidity(package, "confirmed_ok")
        elif chosen == confirm_bad:
            self._confirm_package_liquidity(package, "confirmed_illiquid")
        elif chosen == handoff_action:
            self._copy_kit_package_handoff(table.columnAt(position.x()))

    def _confirm_package_liquidity(
        self,
        package: KitPackageRow,
        decision: str,
    ) -> None:
        """Persist a liquidity decision and rebuild derived pipeline rows.

        Args:
            package: Selected card package.
            decision: ``confirmed_ok`` or ``confirmed_illiquid``.
        """

        kit = self._selected_kit_row()
        if kit is None:
            return
        if self._busy():
            QMessageBox.information(
                self,
                "Ликвидность",
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return
        try:
            self.database.upsert_liquidity_review(
                kit.title,
                kit.mark,
                package.transfer_name or "",
                package.revision_text or "",
                decision=decision,
                sequence=package.sequence,
                evidence_mtime_ns=package.max_mtime_ns,
            )
            rebuild_pipeline(
                self.database,
                records=self._pipeline_records(),
                detected_current_ids=self._detected_current_ids,
                kit_keys={kit_identity_key(kit.title, kit.mark)},
                rd_root=self._pipeline_rd_root(),
            )
        except Exception as exc:
            QMessageBox.warning(self, "Ликвидность", str(exc))
            return
        self._refresh_kits_table()
        self._select_kit_row(kit.title, kit.mark)

    def _robot_origin_for(self, row: KitMatrixRow) -> RobotOrigin:
        key = kit_identity_key(row.title, row.mark)
        return kit_robot_origin(
            row,
            rd_content_equal=self._mto_content_by_kit.get(key),
        )

    def _kit_snapshot_for(self, row: KitMatrixRow, source: str):
        return {"rd": row.rd, "robot": row.robot, "sq": row.sq}.get(source)

    def _open_kit_source(self, source: str, *, folder: bool) -> None:
        row = self._selected_kit_row()
        if row is None:
            return
        if source == "rd":
            package = self._official_rd_package(row.title, row.mark)
            if folder and package is not None and package.package_path:
                self._open_result(package.package_path, folder=False)
                return
        snapshot = self._kit_snapshot_for(row, source)
        path = snapshot.paths[0] if snapshot and snapshot.paths else None
        if source == "rd" and snapshot and snapshot.paths:
            package = self._official_rd_package(row.title, row.mark)
            if package is not None and package.package_path:
                prefix = package.package_path.casefold()
                for candidate in snapshot.paths:
                    folder_path = issued_package_dir(str(candidate)) or str(
                        candidate
                    )
                    if folder_path.casefold().startswith(prefix):
                        path = candidate
                        break
        if not snapshot or not path or not snapshot.present:
            self._append_log(f"Файл {source.upper()} отсутствует")
            return
        if folder:
            package = issued_package_dir(str(path))
            if package:
                self._open_result(package, folder=False)
                return
        self._open_result(str(path), folder=folder)

    def _auto_mto_files_for_kit(self, title: str, mark: str) -> tuple[AutoMtoFile, ...]:
        return self._auto_mto_by_kit.get(kit_identity_key(title, mark), ())

    def _auto_mto_for_kit(
        self,
        title: str,
        mark: str,
        rd_revision: str = "",
    ) -> AutoMtoFile | None:
        return pick_auto_mto_file(
            self._auto_mto_files_for_kit(title, mark),
            rd_revision,
        )

    def _auto_mto_cell_text(self, row: KitMatrixRow, rd_revision: str = "") -> str:
        _path, target_rev, _pinned = self._auto_mto_rd_target(row.title, row.mark)
        return format_auto_mto_cell_text(
            self._auto_mto_files_for_kit(row.title, row.mark),
            target_rev or rd_revision,
        )

    def _rd_mto_overlay_by_kit(self) -> dict[tuple[str, str], tuple[str, str]]:
        """Official MTO path and filename revision per kit for Auto MTO."""

        token = (id(self._mto_worklist_rows), id(self._kit_pipelines))
        cached = getattr(self, "_rd_mto_overlay_cache", None)
        if cached is not None and cached[0] == token:
            return cached[1]
        result = rd_mto_overlay_by_kit(self._mto_worklist_rows, self._kit_pipelines)
        self._rd_mto_overlay_cache = (token, result)
        return result

    def _auto_mto_rd_target(self, title: str, mark: str) -> tuple[str, str, bool]:
        """Return (path, revision, pinned) for AutoMTO content compare."""

        tab = getattr(self, "_revision_matrix_tab", None)
        selection = tab.selection_for(title, mark) if tab is not None else None
        return auto_mto_rd_target(
            title,
            mark,
            pins=self._export_pins_by_key,
            overlay_by_kit=self._rd_mto_overlay_by_kit(),
            selection=selection,
        )

    def _auto_mto_compare_status_for(
        self,
        title: str,
        mark: str,
    ) -> AutoMtoCompareStatus:
        """Return the kit-level AutoMTO content-compare label."""

        files = self._auto_mto_files_for_kit(title, mark)
        path, revision, pinned = self._auto_mto_rd_target(title, mark)
        comparison = None
        comparison_rd_path = ""
        tab = getattr(self, "_revision_matrix_tab", None)
        if tab is not None:
            record = tab.comparison_record(title, mark)
            if record is not None:
                comparison_rd_path, comparison = record
        return auto_mto_compare_status(
            files=files,
            rd_path=path,
            rd_revision=revision,
            comparison=comparison,
            comparison_rd_path=comparison_rd_path,
            pinned=pinned,
        )

    def _kit_an_targets(self, row: KitMatrixRow) -> KitAnTargets:
        path, rd_mto, _pinned = self._auto_mto_rd_target(row.title, row.mark)
        return kit_an_targets(
            row,
            auto_files=self._auto_mto_files_for_kit(row.title, row.mark),
            rd_path=path,
            rd_mto_rev=rd_mto,
            pipeline=self._kit_pipelines.get(
                kit_identity_key(row.title, row.mark)
            ),
        )

    def _an_hit_for(self, row: KitMatrixRow) -> AnKitHit:
        """Return the cached-files AN hit for a Комплекты row.

        Args:
            row: Комплекты matrix row.

        Returns:
            ``match_an_to_kit`` result for that identity.
        """

        key = kit_identity_key(row.title, row.mark)
        files = self._an_files_by_kit.get(key, ())
        return match_an_to_kit(files, self._kit_an_targets(row))

    def _an_targets_by_kit(self) -> dict[tuple[str, str], KitAnTargets]:
        """Return AN targets for every loaded Комплекты row."""

        return {
            kit_identity_key(row.title, row.mark): self._kit_an_targets(row)
            for row in self._kit_rows
        }

    def _reload_an_index(self) -> None:
        """Reload AN files from SQLite. Does not rebuild pipeline or the tree."""

        with perf_span("gui.reload_an_index"):
            try:
                self._an_files_by_kit = self.database.list_an_files_by_kit()
            except Exception as exc:
                self._an_files_by_kit = {}
                self._append_log(f"АН: {type(exc).__name__}: {exc}")

    def _refresh_an_tab(self) -> None:
        """Push the cached AN dump into the finder tab."""

        if not hasattr(self, "_an_tab"):
            return
        self._deferred_widgets.discard(_DEFERRED_AN)
        with perf_span("gui.refresh_an_tab"):
            self._an_tab.set_content_queue_paused(self._catalog_workers_busy())
            rows = tuple(
                file
                for files in self._an_files_by_kit.values()
                for file in files
            )
            known = {
                kit_identity_key(row.title, row.mark) for row in self._kit_rows
            }
            self._an_tab.set_rows(
                rows,
                is_banned=lambda title, mark: self._is_banned_pair(title, mark),
                kit_targets=self._an_targets_by_kit(),
                allowed_kits=known,
            )

    def _reload_rd_dump_index(self) -> None:
        """Reload RD dump files from SQLite. Does not rebuild pipeline."""

        with perf_span("gui.reload_rd_dump_index"):
            try:
                self._rd_dump_files = self.database.list_rd_dump_mto_files()
            except Exception as exc:
                self._rd_dump_files = ()
                self._append_log(f"РД: {type(exc).__name__}: {exc}")

    def _refresh_rd_dump_tab(self) -> None:
        """Push the cached RD dump into the finder tab."""

        if not hasattr(self, "_rd_dump_tab"):
            return
        self._deferred_widgets.discard(_DEFERRED_RD_DUMP)
        with perf_span("gui.refresh_rd_dump_tab"):
            self._rd_dump_tab.set_content_queue_paused(self._catalog_workers_busy())
            self._reload_rd_dump_index()
            known = {
                kit_identity_key(row.title, row.mark) for row in self._kit_rows
            }
            self._rd_dump_tab.set_rows(
                self._rd_dump_files,
                is_banned=lambda title, mark: self._is_banned_pair(title, mark),
                kit_targets=self._an_targets_by_kit(),
                rd_root=self.config.rd_root,
                allowed_kits=known,
            )

    def _open_auto_mto_cell(self, table_row: int) -> None:
        item = self._kits_table.item(table_row, _KITS_COL_AUTO_MTO)
        row = item.data(_ROLE_ROW) if item else None
        if not isinstance(row, KitMatrixRow):
            row = self._selected_kit_row()
        if row is None:
            return
        _path, revision, _pinned = self._auto_mto_rd_target(row.title, row.mark)
        auto = self._auto_mto_for_kit(row.title, row.mark, revision)
        if auto is None:
            self._append_log("Авто МТО: нет записи в базе заказчика")
            return
        path = auto_mto_path(auto)
        if not path.is_file():
            self._append_log(f"Авто МТО: файл ещё не собран ({path.name})")
            return
        self._open_result(str(path), folder=False)

    def _start_kit_auto_mto_compare(self, row: KitMatrixRow) -> None:
        """Compare customer AutoMTO files with the latest or pinned RD MTO."""

        path, _revision, _pinned = self._auto_mto_rd_target(row.title, row.mark)
        if not path:
            self._append_log("Авто МТО: нет текущего файла MTO РД")
            return
        started = self._revision_matrix_tab.start_auto_mto_compare(
            row.title,
            row.mark,
            path,
        )
        if not started:
            self._append_log("Авто МТО: сверка уже выполняется или нет файлов заказчика")

    @Slot()
    def _schedule_auto_mto_compare_refresh(self) -> None:
        self._auto_mto_compare_refresh_timer.start()

    @Slot()
    def _on_auto_mto_compare_finished(self) -> None:
        with perf_span("gui.auto_mto_compare_finished"):
            self._refresh_kits_table(rebuild_rows=False)
            self._refresh_mto_worklist(reload=False)

    def _schedule_auto_mto_compares(self) -> None:
        """Enqueue background AutoMTO compares for kits that still need them.

        Heatmap/export selections go through the tab (same path as the
        painted cell). Kits without a selection yet use the overlay or
        pin so «другой файл» on Комплекты is queued before the heatmap
        tab is opened.
        """

        if not hasattr(self, "_revision_matrix_tab"):
            return
        self._revision_matrix_tab.set_auto_mto_queue_paused(
            self._catalog_workers_busy()
        )
        self._revision_matrix_tab.enqueue_needed_auto_mto_compares()
        jobs: list[tuple[str, str, str]] = []
        for files in self._auto_mto_by_kit.values():
            if not files:
                continue
            title = files[0].title
            mark = files[0].mark
            selection = self._revision_matrix_tab.selection_for(title, mark)
            if selection is not None and selection.source_path:
                continue
            path, _revision, _pinned = self._auto_mto_rd_target(title, mark)
            if not path:
                continue
            status = self._auto_mto_compare_status_for(title, mark)
            if needs_auto_mto_compare(status):
                jobs.append((title, mark, path))
        self._revision_matrix_tab.enqueue_auto_mto_compares(jobs)

    def _enqueue_auto_mto_compares_for_kits(
        self, kit_keys: set[tuple[str, str]]
    ) -> None:
        """Queue Auto MTO content compare only for ``kit_keys``.

        Args:
            kit_keys: Identities whose official MTO target may have changed.
        """

        if not hasattr(self, "_revision_matrix_tab"):
            return
        self._revision_matrix_tab.set_auto_mto_queue_paused(
            self._catalog_workers_busy()
        )
        jobs: list[tuple[str, str, str]] = []
        for title, mark in kit_keys:
            path, _revision, _pinned = self._auto_mto_rd_target(title, mark)
            if not path:
                continue
            if needs_auto_mto_compare(self._auto_mto_compare_status_for(title, mark)):
                jobs.append((title, mark, path))
        self._revision_matrix_tab.enqueue_auto_mto_compares(jobs)

    def _set_transfer_mto_compare_paused(self, paused: bool) -> None:
        """Pause or resume disputed-package MTO compares (not ``_busy``)."""

        if not hasattr(self, "_transfer_mto_paused"):
            return
        self._transfer_mto_paused = bool(paused)
        if not self._transfer_mto_paused:
            self._schedule_transfer_mto_compares()

    def _schedule_transfer_mto_compares(self) -> None:
        """Queue AutoMTO-grade compares for disputed transfer MTO pairs."""

        if getattr(self, "_transfer_mto_paused", False):
            return
        thread = getattr(self, "_transfer_mto_thread", None)
        if thread is not None and thread.isRunning():
            return
        jobs = []
        seen: set[str] = set()
        for row in self._kit_rows:
            for pair in row.transfer_review_mto_pairs:
                key = cache_entry_key(
                    pair.left_path,
                    pair.left_mtime_ns,
                    pair.right_path,
                    pair.right_mtime_ns,
                )
                if key in self._transfer_mto_results or key in seen:
                    continue
                seen.add(key)
                jobs.append(pair)
        if not jobs:
            return
        worker = TransferReviewCompareThread(tuple(jobs), parent=self)
        worker.result_ready.connect(self._on_transfer_mto_compare_ready)
        worker.error.connect(self._on_transfer_mto_compare_error)
        worker.finished.connect(self._on_transfer_mto_compare_finished)
        self._transfer_mto_thread = worker
        worker.start()

    @Slot(object)
    def _on_transfer_mto_compare_ready(self, outcome: object) -> None:
        if not isinstance(outcome, TransferReviewCompareOutcome):
            return
        self._transfer_mto_results[outcome.cache_key] = (
            outcome.result.paren_label
        )
        self._transfer_mto_by_pair[
            path_pair_key(outcome.pair.left_path, outcome.pair.right_path)
        ] = outcome.result.paren_label
        entry = result_to_entry(outcome.result)
        if entry is not None:
            self._transfer_mto_entries[outcome.cache_key] = entry

    @Slot(str)
    def _on_transfer_mto_compare_error(self, message: str) -> None:
        self._append_log(f"Сверка MTO передач: {message[:160]}")

    @Slot()
    def _on_transfer_mto_compare_finished(self) -> None:
        thread = self._transfer_mto_thread
        self._transfer_mto_thread = None
        if thread is not None:
            thread.deleteLater()
        try:
            save_transfer_review_compare_cache(
                self.config.runtime_dir, self._transfer_mto_entries
            )
        except OSError as exc:
            self._append_log(f"Сверка MTO передач: кэш {exc}")
        self._rebuild_kit_rows()
        self._refresh_kits_table(rebuild_rows=False)

    def _reload_auto_mto_index(self) -> None:
        """Load PI Auto MTO plans from the local pickle (no UNC write)."""

        try:
            self._auto_mto_by_kit = list_auto_mto_files_by_kit()
        except FileNotFoundError:
            self._auto_mto_by_kit = {}
        except Exception as exc:
            self._auto_mto_by_kit = {}
            self._append_log(f"Авто МТО: {type(exc).__name__}: {exc}")
        if hasattr(self, "_revision_matrix_tab"):
            self._revision_matrix_tab.seed_rd_mtimes(self._all_records)
            self._revision_matrix_tab.set_auto_mto_index(self._auto_mto_by_kit)
            self._schedule_auto_mto_compares()

    @Slot()
    def _open_customer_pi_dialog(self) -> None:
        if self._busy():
            QMessageBox.information(
                self,
                "База заказчика",
                "Дождитесь окончания скана, Google или копирования MTO.",
            )
            return
        last = str(self._settings.value("window/customer_pi_xlsb") or "")
        kit_keys = {
            kit_identity_key(row.title, row.mark)
            for row in self._kit_rows
            if not self._is_banned_pair(row.title, row.mark, row.title_system)
        }
        dialog = CustomerPiDialog(
            self,
            xlsb_path=last or None,
            kit_keys=kit_keys,
        )
        self._customer_pi_dialog = dialog
        dialog.catalog_rebuilt.connect(self._on_auto_mto_catalog_rebuilt)
        dialog.pickle_updated.connect(self._on_auto_mto_pickle_updated)
        try:
            dialog.exec()
        finally:
            self._settings.setValue(
                "window/customer_pi_xlsb", str(dialog.xlsb_path())
            )
            self._customer_pi_dialog = None
            dialog.deleteLater()

    @Slot()
    def _on_auto_mto_pickle_updated(self) -> None:
        self._reload_auto_mto_index()
        self._refresh_kits_table()
        self._refresh_mto_worklist(reload=False)
        self._refresh_issuance_journal()

    @Slot()
    def _on_auto_mto_catalog_rebuilt(self) -> None:
        self._reload_auto_mto_index()
        self._refresh_kits_table()
        self._refresh_mto_worklist(reload=False)
        self._refresh_issuance_journal()
        self._append_log("Каталог АвтоМТО обновлён.")

    def _copy_kit_paths(self) -> None:
        row = self._selected_kit_row()
        if row is None:
            return
        paths: list[str] = []
        seen: set[str] = set()
        for snapshot in (row.rd, row.robot, row.sq):
            for file_path in snapshot.paths:
                package = issued_package_dir(str(file_path))
                key = package.casefold()
                if package and key not in seen:
                    seen.add(key)
                    paths.append(package)
        self._copy_text("\n".join(paths))

    def _is_banned_pair(
        self,
        title: Any = None,
        mark: Any = None,
        title_system: Any = None,
    ) -> bool:
        parsed_title, parsed_mark = title_mark_from_row(title, mark, title_system)
        return self._ban_store.contains(parsed_title, parsed_mark)

    def _document_tree_kit_identities(self) -> set[tuple[str, str]]:
        """Return contour ``(title, mark)`` pairs for the revision heatmap.

        Issued-transfer layout only. New RD scans omit non-canonical
        files; leftover present rows may still appear on «Все документы»
        until the next successful walk. Skip-dirs hidden, parsed
        title+mark, banned pairs omitted.
        """

        skip = tuple(self.config.skip_dirs)
        token = (
            id(self._all_records),
            skip,
            self._ban_store.revision,
            id(self._contour_records),
        )
        cached = self._tree_kit_identities_cache
        if cached is not None and cached[0] == token:
            return cached[1]
        identities: set[tuple[str, str]] = set()
        check_layout = self._pipeline_rd_root() is not None
        rd_root = self.config.rd_root
        for record in self._pipeline_records():
            if not record.present:
                continue
            if record.source is not SourceKind.RD:
                continue
            if check_layout and not record_has_canonical_layout(record, rd_root):
                continue
            if path_has_skipped_dir(record.path, skip):
                continue
            title = str(record.data.get("title") or "").strip()
            mark = str(record.data.get("mark") or "").strip()
            if not title or not mark:
                continue
            if self._is_banned_pair(title, mark):
                continue
            identities.add(kit_identity_key(title, mark))
        self._tree_kit_identities_cache = (token, identities)
        return identities

    def _is_banned_row(self, row: dict[str, Any] | None) -> bool:
        if not row:
            return False
        return self._is_banned_pair(
            row.get("title"), row.get("mark"), row.get("title_system")
        )

    def _update_kits_tab_label(self) -> None:
        """Show non-banned kit and unique-title counts on the Комплекты tab."""

        if not hasattr(self, "_kits_tab"):
            return
        visible = [
            row
            for row in self._kit_rows
            if not self._is_banned_pair(row.title, row.mark, row.title_system)
        ]
        kit_count = len(visible)
        title_count = len(
            {str(row.title).casefold() for row in visible if row.title}
        )
        tab_index = self._tabs.indexOf(self._kits_tab)
        self._tabs.setTabText(
            tab_index, f"Комплекты ({kit_count} компл., {title_count} тит.)"
        )

    def _update_ban_action_label(self) -> None:
        if not hasattr(self, "_ban_action"):
            return
        count = len(self._ban_store.pairs())
        self._ban_action.setText(
            f"Забаненные титулы ({count})" if count else "Забаненные титулы"
        )

    def _apply_ban_visibility(self) -> None:
        with perf_span("gui.apply_ban_visibility"):
            self._tree_kit_identities_cache = None
            self._apply_kits_filter()
            if _DEFERRED_MTO not in self._deferred_widgets:
                self._refresh_mto_table()
            if _DEFERRED_TREE not in self._deferred_widgets:
                self._rebuild_document_tree()
            if _DEFERRED_HEATMAP not in self._deferred_widgets:
                self._refresh_revision_matrix()
            if _DEFERRED_WORKLIST not in self._deferred_widgets:
                self._refresh_mto_worklist()
            if _DEFERRED_ISSUANCE_JOURNAL not in self._deferred_widgets:
                self._refresh_issuance_journal()
            if _DEFERRED_AN not in self._deferred_widgets:
                self._refresh_an_tab()
            if _DEFERRED_RD_DUMP not in self._deferred_widgets:
                self._refresh_rd_dump_tab()
            self._update_ban_action_label()
            self._update_kits_tab_label()

    @Slot()
    def _open_ban_dialog(self) -> None:
        dialog = BannedTitlesDialog(self._ban_store, self)
        dialog.changed.connect(self._apply_ban_visibility)
        dialog.exec()
        self._apply_ban_visibility()

    @Slot()
    def _open_status_colors_dialog(self) -> None:
        dialog = StatusColorsDialog(
            self._status_colors_path,
            self,
            palette=self._status_colors,
        )
        dialog.changed.connect(self._reload_status_colors)
        dialog.exec()
        self._reload_status_colors()

    def _layout_rd_root(self) -> str:
        """Return the configured RD root, or empty when the filter is off."""

        return str(self.config.rd_root or "").strip()

    def _layout_report_block_reason(self) -> str:
        """Return why the layout report cannot run, or empty when it can."""

        has_rd = any(
            record.present and record.source is SourceKind.RD
            for record in self._all_records
        )
        if not has_rd:
            return _LAYOUT_REPORT_NO_RECORDS
        if not self._layout_rd_root():
            return _LAYOUT_REPORT_NO_ROOT
        return ""

    def _update_layout_report_action(self) -> None:
        """Enable the layout-report action only with RD rows and an RD root."""

        if not hasattr(self, "_layout_report_action"):
            return
        reason = self._layout_report_block_reason()
        self._layout_report_action.setEnabled(not reason)
        self._layout_report_action.setToolTip(
            reason or _LAYOUT_REPORT_READY_TOOLTIP
        )

    @Slot()
    def _on_layout_report(self) -> None:
        """Build the folder-layout report from loaded records and offer save."""

        reason = self._layout_report_block_reason()
        if reason:
            QMessageBox.information(self, _LAYOUT_REPORT_TITLE, reason)
            return
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            report = build_layout_report(
                records=self._all_records,
                rd_root=self._layout_rd_root(),
            )
        except Exception as exc:
            QMessageBox.warning(
                self,
                _LAYOUT_REPORT_TITLE,
                f"Не удалось построить отчёт:\n{type(exc).__name__}: {exc}",
            )
            return
        finally:
            QApplication.restoreOverrideCursor()
        self._present_layout_report(report)

    def _present_layout_report(self, report: LayoutReport) -> None:
        """Show the scale summary, then save the workbook if the user agrees.

        Args:
            report: Grouped layout report already built from loaded records.
        """

        summary = format_layout_report_summary(report)
        reply = QMessageBox.question(
            self,
            _LAYOUT_REPORT_TITLE,
            f"{summary}\n{_LAYOUT_REPORT_SAVE_QUESTION}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        path, _filter = QFileDialog.getSaveFileName(
            self,
            _LAYOUT_REPORT_SAVE_CAPTION,
            layout_report_default_filename(),
            _LAYOUT_REPORT_SAVE_FILTER,
        )
        if not path:
            return
        dest = Path(path)
        if dest.suffix.casefold() != ".xlsx":
            dest = dest.with_suffix(".xlsx")
        try:
            write_layout_report_xlsx(report, dest)
        except OSError as exc:
            QMessageBox.warning(
                self,
                _LAYOUT_REPORT_TITLE,
                f"Не удалось сохранить отчёт:\n{exc}",
            )
            return
        self._append_log(f"Отчёт о раскладке сохранён: {dest}")
        open_reply = QMessageBox.question(
            self,
            _LAYOUT_REPORT_TITLE,
            f"Отчёт сохранён:\n{dest}\n\nОткрыть файл?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if open_reply == QMessageBox.StandardButton.Yes:
            self._open_result(str(dest), folder=False)

    def _reload_status_colors(self) -> None:
        selected = self._selected_kit_row()
        self._mto_icon_cache.clear()
        self._status_colors = load_status_colors(self._status_colors_path)
        self._reload_kits_card_legend()
        if _DEFERRED_HEATMAP not in self._deferred_widgets:
            self._refresh_revision_matrix(force=True)
        self._refresh_kits_table()
        if _DEFERRED_WORKLIST not in self._deferred_widgets:
            self._refresh_mto_worklist()
        if _DEFERRED_ISSUANCE_JOURNAL not in self._deferred_widgets:
            self._refresh_issuance_journal()
        if selected is not None:
            self._select_kit_row(selected.title, selected.mark)
        else:
            self._update_kits_card()
        if _DEFERRED_TREE not in self._deferred_widgets:
            self._relabel_document_tree()
            self._refresh_history()

    def _mto_status_icon(self, status: str) -> QIcon:
        """Return a cached 12px square icon for one MTO status color."""

        key = status or "empty"
        cached = self._mto_icon_cache.get(key)
        if cached is not None:
            return cached
        color = QColor(color_for(self._status_colors, key))
        pixmap = QPixmap(12, 12)
        pixmap.fill(color)
        painter = QPainter(pixmap)
        painter.setPen(color.darker(130))
        painter.drawRect(0, 0, 11, 11)
        painter.end()
        icon = QIcon(pixmap)
        self._mto_icon_cache[key] = icon
        return icon

    def _sync_skip_dirs_from_store(self) -> None:
        self.config = replace(self.config, skip_dirs=self._skip_store.tokens())
        self._tree_kit_identities_cache = None
        if _DEFERRED_TREE not in self._deferred_widgets:
            self._rebuild_document_tree()
        if _DEFERRED_HEATMAP not in self._deferred_widgets:
            self._refresh_revision_matrix(force=True)
        if _DEFERRED_WORKLIST not in self._deferred_widgets:
            self._refresh_mto_worklist()
        if _DEFERRED_ISSUANCE_JOURNAL not in self._deferred_widgets:
            self._refresh_issuance_journal()

    @Slot()
    def _open_skip_dirs_dialog(self) -> None:
        dialog = SkipDirsDialog(self._skip_store, self)
        dialog.changed.connect(self._sync_skip_dirs_from_store)
        dialog.exec()
        self._sync_skip_dirs_from_store()

    @Slot()
    def _prune_skipped_from_db(self) -> None:
        if self._busy():
            QMessageBox.information(
                self,
                "Skip-папки",
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return
        tokens = self.config.skip_dirs
        count = self.database.count_present_skipped(tokens)
        if count == 0:
            QMessageBox.information(
                self,
                "Skip-папки",
                "Нет файлов в базе под текущий skip.",
            )
            return
        reply = QMessageBox.question(
            self,
            "Убрать skip из дерева",
            f"Убрать {count} файлов из каталога без скана сети?\n"
            "Файлы на диске не трогаем. Вернуть можно только полным сканом РД/SQ.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        with perf_span("gui.prune_skip_dirs"):
            self._cancel_mto_compare(resume_later=True)
            self._cancel_export_pair_compare()
            removed = self.database.apply_skip_dirs(tokens)
            self._append_log(
                f"Skip-папки: убрано из базы {removed} файлов без скана UNC."
            )
            self.refresh(ingest_google=False)
            if self._resume_pending_mto_compare():
                self._mto_resume_on_idle = False

    def _ban_title_mark(
        self,
        title: str,
        mark: str,
        *,
        comment: str = "",
    ) -> None:
        try:
            pair, added = self._ban_store.add(title, mark, comment)
        except ValueError as exc:
            QMessageBox.warning(self, "Бан-фильтр", str(exc))
            return
        except OSError as exc:
            QMessageBox.warning(
                self, "Бан-фильтр", f"Не удалось сохранить список:\n{exc}"
            )
            return
        if not added:
            QMessageBox.information(
                self, "Бан-фильтр", f"{pair.label} уже скрыт."
            )
            return
        self._append_log(f"Бан-фильтр: скрыт {pair.label}")
        self._apply_ban_visibility()

    def _confirm_ban_title_mark(self, title: str, mark: str) -> None:
        parsed_title, parsed_mark = title_mark_from_row(title, mark)
        if not parsed_title or not parsed_mark:
            QMessageBox.warning(
                self, "Бан-фильтр", "У строки нет титула и марки."
            )
            return
        label = f"{parsed_title}-{parsed_mark}"
        if self._ban_store.contains(parsed_title, parsed_mark):
            QMessageBox.information(
                self, "Бан-фильтр", f"{label} уже скрыт."
            )
            return
        reply = QMessageBox.question(
            self,
            "Бан-фильтр",
            f"Скрыть {label} из комплектов, MTO и документов?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self._ban_title_mark(parsed_title, parsed_mark)

    def _show_kits_context_menu(self, position) -> None:
        row_index = self._kits_table.rowAt(position.y())
        if row_index < 0:
            return
        with perf_span("gui.show_kits_context_menu"):
            ctx = self._popup_kits_context_menu(row_index, position)
        if ctx is None:
            return
        row = ctx["row"]
        chosen = ctx["chosen"]
        jump_action = ctx["jump_action"]
        journal_action = ctx["journal_action"]
        legalize_rd_action = ctx["legalize_rd_action"]
        an_action = ctx["an_action"]
        rd_dump_action = ctx["rd_dump_action"]
        an_folder_action = ctx["an_folder_action"]
        shown_an = ctx["shown_an"]
        actions = ctx["actions"]
        auto_action = ctx["auto_action"]
        copy_all = ctx["copy_all"]
        outlook_query = ctx["outlook_query"]
        outlook_find_action = ctx["outlook_find_action"]
        outlook_copy_action = ctx["outlook_copy_action"]
        rescan_action = ctx["rescan_action"]
        mixed_open_action = ctx["mixed_open_action"]
        sync_action = ctx["sync_action"]
        compare_auto = ctx["compare_auto"]
        sq_to_rd_action = ctx["sq_to_rd_action"]
        ban_action = ctx["ban_action"]
        handoff_action = ctx["handoff_action"]
        if chosen == jump_action:
            self._jump_to_kit_documents(row.title, row.mark)
            return
        if chosen == journal_action:
            self._ensure_deferred_widget(_DEFERRED_ISSUANCE_JOURNAL)
            self._tabs.setCurrentWidget(self._issuance_journal_tab)
            self._issuance_journal_tab.focus_kit(
                row.title, row.mark, issuance=row.issuance
            )
            return
        if chosen == legalize_rd_action:
            self._open_legalize_rd_dialog(row)
            return
        if chosen == an_action:
            self._tabs.setCurrentWidget(self._an_tab)
            self._an_tab.set_kit_filter(row.title, row.mark)
            if not self._an_tab.focus_best_agreed():
                targets = self._kit_an_targets(row)
                hit = self._an_hit_for(row)
                self._an_tab.focus_revision(
                    targets.auto_mto or hit.shown_revision
                )
            return
        if chosen == rd_dump_action:
            self._tabs.setCurrentWidget(self._rd_dump_tab)
            self._rd_dump_tab.set_kit_filter(row.title, row.mark)
            if not self._rd_dump_tab.focus_best_agreed():
                targets = self._kit_an_targets(row)
                self._rd_dump_tab.focus_revision(targets.auto_mto or targets.rd_mto)
            return
        if chosen == an_folder_action:
            if shown_an is not None and shown_an.parent_dir:
                open_path(shown_an.parent_dir)
            return
        if chosen == auto_action:
            self._open_auto_mto_cell(row_index)
            return
        if chosen == copy_all:
            self._copy_kit_paths()
            return
        if chosen == outlook_find_action:
            self._open_outlook_od_search(outlook_query)
            return
        if chosen == outlook_copy_action:
            self._copy_text(
                outlook_query,
                message=f"Скопировано для Outlook: {outlook_query}",
            )
            return
        if chosen == ctx["google_action"]:
            href = str(ctx.get("google_href") or "")
            if href:
                open_google_sheet_url(href)
            return
        if chosen == rescan_action:
            self._rescan_kit_rd(row)
            return
        if chosen == mixed_open_action:
            self._open_mixed_title_folders(row)
            return
        if chosen == sync_action:
            self._confirm_robot_mto_sync(row)
            return
        if chosen == compare_auto:
            self._start_kit_auto_mto_compare(row)
            return
        if chosen == sq_to_rd_action:
            self._confirm_sq_to_rd(row)
            return
        if chosen == ban_action:
            self._confirm_ban_title_mark(row.title, row.mark)
            return
        if chosen == handoff_action:
            column = self._kits_table.columnAt(position.x())
            self._copy_robot_handoff(
                self._kits_row_handoff_text(row_index, column)
            )
            return
        command = actions.get(chosen)
        if not command:
            return
        if command[0] == "copy":
            self._copy_text(command[1])
        else:
            self._open_kit_source(command[0], folder=bool(command[1]))

    def _popup_kits_context_menu(
        self, row_index: int, position
    ) -> dict[str, Any] | None:
        """Build and exec the Комплекты menu without disk existence checks."""

        if self._kits_table.currentRow() != row_index:
            self._kits_table.selectRow(row_index)
        row = self._selected_kit_row()
        if row is None:
            return None
        menu = QMenu(self)
        jump_action = menu.addAction('Показать в «Все документы»')
        journal_action = menu.addAction("Показать в Выдача · Журнал")
        legalize_rd_action = menu.addAction("Легализовать ревизию РД…")
        legalize_rd_action.setEnabled(bool(row.title and row.mark and row.rd.present))
        legalize_rd_action.setToolTip(
            "Открыть журнал с предзаполненной строкой текущей ревизии РД."
        )
        an_action = menu.addAction('Показать в „АН“')
        rd_dump_action = menu.addAction('Показать в „РД“')
        an_folder_action = menu.addAction("Открыть папку · АН")
        shown_an = self._an_hit_for(row).shown_file
        an_folder_action.setEnabled(
            bool(shown_an is not None and shown_an.parent_dir)
        )
        menu.addSeparator()
        actions: dict[QAction, tuple[str, bool] | tuple[str, str]] = {}
        for source, label in (("rd", "РД"), ("robot", "робот"), ("sq", "SQ")):
            snapshot = self._kit_snapshot_for(row, source)
            path = snapshot.paths[0] if snapshot and snapshot.paths else ""
            if source == "rd":
                package = self._official_rd_package(row.title, row.mark)
                if package is not None and package.package_path:
                    path = package.package_path
                    if snapshot and snapshot.paths:
                        prefix = package.package_path.casefold()
                        for candidate in snapshot.paths:
                            folder_path = issued_package_dir(
                                str(candidate)
                            ) or str(candidate)
                            if folder_path.casefold().startswith(prefix):
                                path = candidate
                                break
            present = bool(
                (snapshot and snapshot.present and path)
                or (source == "rd" and path)
            )
            open_action = menu.addAction(f"Открыть файл · {label}")
            folder_action = menu.addAction(f"Открыть содержащую папку · {label}")
            copy_action = menu.addAction(f"Копировать путь · {label}")
            open_action.setEnabled(present)
            folder_action.setEnabled(present)
            copy_action.setEnabled(bool(path))
            actions[open_action] = (source, False)
            actions[folder_action] = (source, True)
            actions[copy_action] = ("copy", path)
            menu.addSeparator()
        auto_action = menu.addAction("Открыть файл · Авто МТО")
        rd_mto_path, auto_rev, _pinned = self._auto_mto_rd_target(
            row.title, row.mark
        )
        auto = self._auto_mto_for_kit(row.title, row.mark, auto_rev)
        auto_file = auto_mto_path(auto) if auto is not None else None
        auto_action.setEnabled(auto_file is not None)
        menu.addSeparator()
        copy_all = menu.addAction("Копировать все пути")
        outlook_query = outlook_od_search_query(row.title, row.mark)
        outlook_find_action = menu.addAction("Найти в Outlook")
        outlook_find_action.setEnabled(bool(outlook_query))
        outlook_find_action.setToolTip(
            f"Найти {outlook_query} во всех почтовых ящиках Outlook"
            if outlook_query
            else "Нужны титул и марка."
        )
        outlook_copy_action = menu.addAction("Копировать поиск Outlook")
        outlook_copy_action.setEnabled(bool(outlook_query))
        outlook_copy_action.setToolTip(
            f"Строка для поиска в Outlook: {outlook_query}"
            if outlook_query
            else "Нужны титул и марка."
        )
        column = self._kits_table.columnAt(position.x())
        google_href = ""
        href_item = (
            self._kits_table.item(row_index, column) if column >= 0 else None
        )
        if href_item is not None:
            google_href = str(href_item.data(ROLE_HREF) or "")
        google_action = menu.addAction("Открыть в Google")
        google_action.setEnabled(bool(google_href))
        google_action.setToolTip(
            "Открыть эту ячейку в Google Sheets (как Shift+клик)"
            if google_href
            else "Shift+клик по ячейке Google / TRM с номером строки снимка."
        )
        menu.addSeparator()
        rescan_action = menu.addAction("Пересканировать РД")
        rescan_action.setToolTip(
            "Частичный перескан папки комплекта в РД "
            r"(РД\титул\марка\Для передачи\NN_…)."
        )
        rescan_action.setEnabled(
            not self._busy() and bool(row.title and row.mark)
        )
        mixed_folders = self._mixed_title_open_folders_for_kit(row)
        mixed_open_action = menu.addAction("Открыть смешанные папки")
        mixed_open_action.setEnabled(bool(mixed_folders))
        if mixed_folders:
            mixed_open_action.setToolTip(
                "Открыть папки, где лежат файлы комплекта в чужом титуле:\n"
                + "\n".join(mixed_folders)
            )
        else:
            mixed_open_action.setToolTip(
                "Нет файлов комплекта в папке другого титула."
            )
        sync_action = menu.addAction("Обновить MTO у робота")
        compare_auto = menu.addAction("Сверить Авто МТО с MTO РД…")
        compare_auto.setEnabled(
            bool(rd_mto_path)
            and bool(self._auto_mto_files_for_kit(row.title, row.mark))
            and not self._busy()
        )
        sq_to_rd_action = menu.addAction("Перенести SQ в РД (новая передача)")
        sq_to_rd_action.setEnabled(bool(row.sq.present and row.sq.paths))
        menu.addSeparator()
        ban_action = menu.addAction("Скрыть титул–марку (бан-фильтр)")
        handoff_action = self._add_robot_handoff_action(menu)
        chosen = exec_tracked_menu(
            menu, MENU_KITS, self._kits_table.viewport().mapToGlobal(position)
        )
        return {
            "row": row,
            "chosen": chosen,
            "jump_action": jump_action,
            "journal_action": journal_action,
            "legalize_rd_action": legalize_rd_action,
            "an_action": an_action,
            "rd_dump_action": rd_dump_action,
            "an_folder_action": an_folder_action,
            "shown_an": shown_an,
            "actions": actions,
            "auto_action": auto_action,
            "copy_all": copy_all,
            "outlook_query": outlook_query,
            "outlook_find_action": outlook_find_action,
            "outlook_copy_action": outlook_copy_action,
            "google_action": google_action,
            "google_href": google_href,
            "rescan_action": rescan_action,
            "mixed_open_action": mixed_open_action,
            "sync_action": sync_action,
            "compare_auto": compare_auto,
            "sq_to_rd_action": sq_to_rd_action,
            "ban_action": ban_action,
            "handoff_action": handoff_action,
        }

    def _update_collision_tab_label(self) -> None:
        """Show the current collision count on the tab without filling the table."""

        if not hasattr(self, "_collision_tab"):
            return
        count = len(self._collision_rows)
        if hasattr(self, "_collision_count_label"):
            self._collision_count_label.setText(f"Текущие коллизии: {count}")
        tab_index = self._tabs.indexOf(self._collision_tab)
        self._tabs.setTabText(tab_index, f"Коллизии ({count})")

    def _refresh_collision_table(self) -> None:
        with perf_span("gui.refresh_collision_table"):
            self._deferred_widgets.discard(_DEFERRED_COLLISIONS)
            table = self._collision_table
            table.setSortingEnabled(False)
            table.setRowCount(len(self._collision_rows))
            self._update_collision_tab_label()
            for table_row, row in enumerate(self._collision_rows):
                paths = [str(path) for path in row.get("paths") or ()]
                kind = str(row.get("kind") or "")
                values = [
                    row.get("scope") or "—",
                    str(row.get("source") or "—").upper(),
                    _collision_kind_label(kind),
                    row.get("document_key") or "—",
                    row.get("message") or "—",
                    "\n".join(paths) or "—",
                ]
                for column, value in enumerate(values):
                    item = QTableWidgetItem(str(value))
                    item.setData(_ROLE_ROW, row)
                    if column == 5:
                        item.setToolTip(
                            _qt_tooltip("\n".join(paths) or "Путь не определён")
                        )
                    table.setItem(table_row, column, item)
            table.setSortingEnabled(True)

    @Slot()
    def _apply_mto_filter(self) -> None:
        needle = self._mto_filter.text().strip().casefold()
        problems_only = self._mto_problems.isChecked()
        for row_index in range(self._mto_table.rowCount()):
            item = self._mto_table.item(row_index, 0)
            row = item.data(_ROLE_ROW) if item else {}
            haystack = " ".join(
                str(value or "")
                for value in (
                    row.get("title_system"),
                    row.get("discipline_block"),
                    row.get("status"),
                    row.get("rd_path"),
                    row.get("robot_path"),
                    _compact_diff(row),
                )
            ).casefold()
            status = str(row.get("status") or "").casefold()
            as_build = bool(row.get("rd_is_as_build"))
            hide_as_build = (
                hasattr(self, "_mto_no_as_build") and self._mto_no_as_build.isChecked()
            )
            visible = (not needle or needle in haystack) and (
                not problems_only or status in {"warning", "blocked"}
            )
            if hide_as_build and as_build:
                visible = False
            if self._is_banned_row(row):
                visible = False
            self._mto_table.setRowHidden(row_index, not visible)

    @Slot()
    def _on_doc_filter_text_changed(self, *_args: object) -> None:
        """Restart the title-filter debounce timer on each keystroke."""

        if hasattr(self, "_doc_filter_timer"):
            self._doc_filter_timer.start()

    @Slot()
    def _apply_document_tree_filter(self) -> None:
        """Hide tree nodes that do not match the title filter.

        Does not regroup files or recreate widgets. If the current node
        is filtered out, selection and the history table are cleared;
        the first remaining leaf is not auto-selected.
        """

        if not hasattr(self, "_doc_tree"):
            return
        remembered = self._doc_tree.currentItem()
        self._doc_tree.blockSignals(True)
        self._doc_tree.setUpdatesEnabled(False)
        lost = False
        try:
            self._sync_document_tree_filter_hidden()
            lost = remembered is None or _tree_item_or_ancestor_hidden(remembered)
            if lost:
                self._doc_tree.clearSelection()
                self._doc_tree.setCurrentItem(None)
            else:
                self._doc_tree.setCurrentItem(remembered)
        finally:
            self._doc_tree.setUpdatesEnabled(True)
            self._doc_tree.blockSignals(False)
        if lost:
            self._history.setRowCount(0)
            self._update_action_states()

    def _sync_document_tree_filter_hidden(self) -> None:
        """Apply or clear ``setHidden`` from the current filter text."""

        needle = self._doc_filter.text().strip().casefold()
        for title_index in range(self._doc_tree.topLevelItemCount()):
            title_item = self._doc_tree.topLevelItem(title_index)
            if title_item is None:
                continue
            title_visible = False
            for mark_index in range(title_item.childCount()):
                mark_item = title_item.child(mark_index)
                if mark_item is None:
                    continue
                mark_visible = False
                for rev_index in range(mark_item.childCount()):
                    revision_item = mark_item.child(rev_index)
                    if revision_item is None:
                        continue
                    haystack = str(revision_item.data(0, _ROLE_HAYSTACK) or "")
                    visible = not needle or needle in haystack
                    revision_item.setHidden(not visible)
                    if visible:
                        mark_visible = True
                mark_item.setHidden(not mark_visible)
                if mark_visible:
                    title_visible = True
                    if needle:
                        mark_item.setExpanded(True)
            title_item.setHidden(not title_visible)
        if not needle:
            self._doc_tree.expandToDepth(1)

    def _paint_revision_tree_item(
        self,
        revision_item: QTreeWidgetItem,
        title_item: QTreeWidgetItem,
        mark_item: QTreeWidgetItem,
        bundles: Sequence[DocumentBundle],
        *,
        files: list[FileRecord],
        hint: FolderTreeHint | None,
    ) -> None:
        """Paint label, tooltip, and status fill for one revision node.

        Resets background, foreground, icon, and MTO font so unmarking
        working or hiding MTO status does not leave stale decoration.

        Args:
            revision_item: Transfer/revision ``QTreeWidgetItem``.
            title_item: Parent title node (layout-mismatch nested paint).
            mark_item: Parent mark node (layout-mismatch nested paint).
            bundles: Documents stored on ``_ROLE_ROW``.
            files: Flattened catalog files in this folder.
            hint: Cached Google/pipeline extras for the folder.
        """

        revision_item.setBackground(0, QBrush())
        revision_item.setForeground(0, QBrush())
        show_mto_status = (
            hasattr(self, "_doc_show_mto_status")
            and self._doc_show_mto_status.isChecked()
        )
        if not show_mto_status:
            revision_item.setIcon(0, QIcon())
        font = QFont(revision_item.font(0))
        font.setBold(False)
        font.setUnderline(False)
        revision_item.setFont(0, font)
        folder_name = folder_display_name(files)
        is_current = any(bundle.is_current for bundle in bundles)
        has_mto = folder_has_mto(files)
        is_working, is_annulled = _hint_folder_marks(hint)
        has_as_build = folder_has_as_build(files)
        review_status = hint.review_status if hint is not None else ""
        mto_label = _mto_node_label(hint)
        revision_label = folder_tree_label(
            files,
            options=self._tree_label_options(),
            folder_name=folder_name,
            review_status=review_status,
            is_current=is_current,
            has_mto=has_mto,
            is_working=is_working,
            is_annulled=is_annulled,
            has_as_build=has_as_build,
            mto_label=mto_label,
        )
        revision_item.setText(0, revision_label)
        revision_item.setData(
            0,
            _ROLE_HAYSTACK,
            _revision_node_haystack(
                list(bundles),
                revision_label,
                folder_name,
                review_status,
                is_working=is_working,
                is_annulled=is_annulled,
            ),
        )
        extra_tip = folder_tree_tooltip(
            files,
            folder_name=folder_name,
            review_full=hint.review_full if hint is not None else "",
            match_reason=hint.match_reason if hint is not None else "",
            f_label=hint.f_label if hint is not None else "",
            is_current=is_current,
            has_mto=has_mto,
            is_working=is_working,
            working_origin=hint.working_origin if hint is not None else "",
            is_annulled=is_annulled,
            has_as_build=has_as_build,
            mto_status_label=(
                status_color_label(hint.mto_status)
                if hint is not None and hint.mto_status
                else ""
            ),
            mto_revision_text=hint.mto_revision_text if hint is not None else "",
            mto_problems=", ".join(hint.problem_kinds) if hint is not None else "",
        )
        has_pending = any(
            record.review_state is ReviewState.PENDING for record in files
        )
        has_problem = any(
            not record.present or str(record.data.get("parse_status")) != "parsed"
            for record in files
        )
        has_layout_mismatch = any(
            not self._record_in_contour(record) for record in files
        )
        if has_pending:
            revision_item.setForeground(0, QBrush(QColor("#5f259f")))
        elif has_problem:
            revision_item.setForeground(0, QBrush(QColor("#b3261e")))
        if has_layout_mismatch:
            _paint_layout_mismatch(revision_item, nested=False)
            _paint_layout_mismatch(mark_item, nested=True)
            _paint_layout_mismatch(title_item, nested=True)
            if extra_tip:
                existing = revision_item.toolTip(0)
                revision_item.setToolTip(
                    0, f"{existing}\n{extra_tip}" if existing else extra_tip
                )
        elif extra_tip:
            revision_item.setToolTip(0, extra_tip)
        else:
            revision_item.setToolTip(0, "")
        if is_working and not has_layout_mismatch:
            _paint_tree_working(
                revision_item, color_for(self._status_colors, "working")
            )
        if is_annulled and not has_layout_mismatch:
            _paint_tree_working(
                revision_item, color_for(self._status_colors, "annulled")
            )
        if show_mto_status:
            status_key = hint.mto_status if hint is not None else ""
            revision_item.setIcon(0, self._mto_status_icon(status_key))
            if hint is not None and (hint.problem_kinds or hint.is_current_mto):
                mto_font = QFont(revision_item.font(0))
                if hint.problem_kinds:
                    mto_font.setBold(True)
                if hint.is_current_mto:
                    mto_font.setUnderline(True)
                revision_item.setFont(0, mto_font)

    def _sync_tree_layout_mismatch_ancestors(
        self,
        title_item: QTreeWidgetItem,
        mark_item: QTreeWidgetItem,
    ) -> None:
        """Set or clear nested yellow on mark/title after a revision relabel.

        Nested mismatch is the OR of child folders. Relabel resets only the
        revision item, so parents would otherwise keep a stale fill.

        Args:
            title_item: Title node that owns ``mark_item``.
            mark_item: Mark node whose revision children were just painted.
        """

        mark_mismatch = False
        for index in range(mark_item.childCount()):
            child = mark_item.child(index)
            if child is None or child.data(0, _ROLE_NODE_KIND) != _NODE_REVISION:
                continue
            files = self._tree_item_records(child)
            if any(not self._record_in_contour(record) for record in files):
                mark_mismatch = True
                break
        if mark_mismatch:
            _paint_layout_mismatch(mark_item, nested=True)
        else:
            mark_item.setBackground(0, QBrush())
            mark_item.setToolTip(0, "")
        title_mismatch = mark_mismatch
        if not title_mismatch:
            for mark_index in range(title_item.childCount()):
                sibling = title_item.child(mark_index)
                if sibling is None:
                    continue
                for rev_index in range(sibling.childCount()):
                    child = sibling.child(rev_index)
                    if (
                        child is None
                        or child.data(0, _ROLE_NODE_KIND) != _NODE_REVISION
                    ):
                        continue
                    files = self._tree_item_records(child)
                    if any(
                        not self._record_in_contour(record) for record in files
                    ):
                        title_mismatch = True
                        break
                if title_mismatch:
                    break
        if title_mismatch:
            _paint_layout_mismatch(title_item, nested=True)
        else:
            title_item.setBackground(0, QBrush())
            title_item.setToolTip(0, "")

    def _iter_revision_tree_items(
        self, kit_keys: set[tuple[str, str]] | None = None
    ) -> Iterator[tuple[QTreeWidgetItem, QTreeWidgetItem, QTreeWidgetItem]]:
        """Yield (revision_item, title_item, mark_item) for live tree nodes.

        Args:
            kit_keys: When set, only revision nodes of those identities.

        Yields:
            Live revision items with their title and mark parents.
        """

        if not hasattr(self, "_doc_tree"):
            return
        wanted = (
            {kit_identity_key(title, mark) for title, mark in kit_keys}
            if kit_keys is not None
            else None
        )
        for title_index in range(self._doc_tree.topLevelItemCount()):
            title_item = self._doc_tree.topLevelItem(title_index)
            if title_item is None:
                continue
            for mark_index in range(title_item.childCount()):
                mark_item = title_item.child(mark_index)
                if mark_item is None:
                    continue
                if wanted is not None:
                    key = kit_identity_key(title_item.text(0), mark_item.text(0))
                    if key not in wanted:
                        continue
                for rev_index in range(mark_item.childCount()):
                    revision_item = mark_item.child(rev_index)
                    if revision_item is None:
                        continue
                    if revision_item.data(0, _ROLE_NODE_KIND) != _NODE_REVISION:
                        continue
                    yield revision_item, title_item, mark_item

    def _relabel_document_tree(
        self,
        kit_keys: set[tuple[str, str]] | None = None,
    ) -> None:
        """Repaint existing revision nodes without ``QTreeWidget.clear``.

        Args:
            kit_keys: When set, only those identities. ``None`` relabels
                every live revision node (label-checkbox toggles).
        """

        if _DEFERRED_TREE in self._deferred_widgets:
            return
        items = list(self._iter_revision_tree_items(kit_keys))
        if not items:
            return
        with perf_span(
            "gui.relabel_document_tree",
            kits=len(kit_keys) if kit_keys is not None else "all",
        ):
            self._doc_tree.blockSignals(True)
            self._doc_tree.setUpdatesEnabled(False)
            try:
                for revision_item, title_item, mark_item in items:
                    bundles = [
                        bundle
                        for bundle in (revision_item.data(0, _ROLE_ROW) or ())
                        if isinstance(bundle, DocumentBundle)
                    ]
                    files = [
                        record
                        for bundle in bundles
                        for record in _bundle_records(bundle)
                    ]
                    hint = self._folder_hint_for(
                        title_item.text(0), mark_item.text(0), files
                    )
                    self._paint_revision_tree_item(
                        revision_item,
                        title_item,
                        mark_item,
                        bundles,
                        files=files,
                        hint=hint,
                    )
                seen_marks: set[int] = set()
                for _revision_item, title_item, mark_item in items:
                    mark_id = id(mark_item)
                    if mark_id in seen_marks:
                        continue
                    seen_marks.add(mark_id)
                    self._sync_tree_layout_mismatch_ancestors(
                        title_item, mark_item
                    )
                self._sync_document_tree_filter_hidden()
            finally:
                self._doc_tree.setUpdatesEnabled(True)
                self._doc_tree.blockSignals(False)
            current = self._doc_tree.currentItem()
            if (
                current is not None
                and current.data(0, _ROLE_NODE_KIND) == _NODE_REVISION
            ):
                if kit_keys is None:
                    self._refresh_history()
                else:
                    title, mark, _transfer = self._tree_item_identity(current)
                    if kit_identity_key(title, mark) in kit_keys:
                        self._refresh_history()

    @Slot()
    def _rebuild_document_tree(
        self,
        *_args: object,
        keep_selection: bool = False,
        select_first: bool = True,
    ) -> None:
        """Rebuild the documents tree from the in-memory file list.

        Extra Qt signal payloads (filter text, checkbox bool) are ignored.
        The title-filter text only hides nodes; it does not omit them from
        the tree. Revision folders under a mark are ordered by transfer
        ``NN`` ascending; folders without a sequence come last.

        Args:
            keep_selection: Reselect the previous title/mark/folder node
                after rebuild (label-checkbox toggles).
            select_first: Select the first visible revision when nothing
                else is restored. Filter typing passes ``False``.
        """

        self._deferred_widgets.discard(_DEFERRED_TREE)
        if not hasattr(self, "_doc_tree"):
            return
        with perf_span("gui.rebuild_document_tree"):
            if hasattr(self, "_doc_filter_timer"):
                self._doc_filter_timer.stop()
            previous: tuple[str, str, str | None] | None = None
            if keep_selection:
                current = self._doc_tree.currentItem()
                if current is not None:
                    previous = self._tree_item_identity(current)
            pending_only = self._doc_pending.isChecked()
            problems_only = self._doc_problems.isChecked()
            token = (
                id(self._all_records),
                frozenset(self._detected_current_ids),
                tuple(self.config.skip_dirs),
                pending_only,
                problems_only,
            )
            cache_hit = (
                self._doc_bundles_token == token
                and isinstance(self._doc_bundles, list)
            )
            with perf_span(
                "gui.bundle_documents",
                **({"cache": 1} if cache_hit else {}),
            ):
                if not cache_hit:
                    skip_dirs = self.config.skip_dirs
                    skip_by_parent: dict[str, bool] = {}
                    visible: list[FileRecord] = []
                    for record in self._all_records:
                        parent = str(Path(record.path).parent)
                        skipped = skip_by_parent.get(parent)
                        if skipped is None:
                            skipped = path_has_skipped_dir(record.path, skip_dirs)
                            skip_by_parent[parent] = skipped
                        if skipped:
                            continue
                        if record.present or pending_only or problems_only:
                            visible.append(record)
                    self._doc_bundles = bundle_documents(
                        visible,
                        detected_current_ids=self._detected_current_ids,
                    )
                    self._doc_bundles_token = token
            grouped: dict[tuple[str, str, str], list[DocumentBundle]] = defaultdict(
                list
            )
            for bundle in self._doc_bundles:
                title = bundle.title or "Без титула"
                mark = bundle.mark or "Без марки"
                if self._is_banned_pair(title, mark):
                    continue
                grouped[(title, mark, bundle.folder_key)].append(bundle)

            self._doc_tree.blockSignals(True)
            self._doc_tree.setUpdatesEnabled(False)
            try:
                with perf_span("gui.document_tree_qt"):
                    self._doc_tree.clear()
                    title_items: dict[str, QTreeWidgetItem] = {}
                    mark_items: dict[tuple[str, str], QTreeWidgetItem] = {}
                    ordered_nodes = sorted(
                        grouped.items(),
                        key=lambda pair: (
                            pair[0][0].casefold(),
                            pair[0][1].casefold(),
                            folder_tree_sort_key(
                                [
                                    record
                                    for bundle in pair[1]
                                    for record in _bundle_records(bundle)
                                ],
                                folder_key=pair[0][2],
                            ),
                        ),
                    )
                    for (title, mark, _folder_key), bundles in ordered_nodes:
                        files = [
                            record
                            for bundle in bundles
                            for record in _bundle_records(bundle)
                        ]
                        hint = self._folder_hint_for(title, mark, files)
                        has_pending = any(
                            record.review_state is ReviewState.PENDING for record in files
                        )
                        has_problem = any(
                            not record.present
                            or str(record.data.get("parse_status")) != "parsed"
                            for record in files
                        )
                        has_layout_mismatch = any(
                            not self._record_in_contour(record) for record in files
                        )
                        if pending_only and not has_pending:
                            continue
                        if problems_only and not has_problem and not has_layout_mismatch:
                            continue
                        title_item = title_items.get(title)
                        if title_item is None:
                            title_item = QTreeWidgetItem(self._doc_tree, [title])
                            title_item.setData(0, _ROLE_NODE_KIND, _NODE_TITLE)
                            title_items[title] = title_item
                        mark_item = mark_items.get((title, mark))
                        if mark_item is None:
                            mark_item = QTreeWidgetItem(title_item, [mark])
                            mark_item.setData(0, _ROLE_NODE_KIND, _NODE_MARK)
                            mark_items[(title, mark)] = mark_item
                        revision_item = QTreeWidgetItem(mark_item, [""])
                        revision_item.setData(0, _ROLE_NODE_KIND, _NODE_REVISION)
                        revision_item.setData(0, _ROLE_ROW, list(bundles))
                        self._paint_revision_tree_item(
                            revision_item,
                            title_item,
                            mark_item,
                            bundles,
                            files=files,
                            hint=hint,
                        )
                    self._doc_tree.expandToDepth(1)
                    self._sync_document_tree_filter_hidden()
            finally:
                self._doc_tree.setUpdatesEnabled(True)
                self._doc_tree.blockSignals(False)
            if previous is not None:
                self._select_document_tree_node(*previous)
                current = self._doc_tree.currentItem()
                if current is not None and not _tree_item_or_ancestor_hidden(current):
                    return
            if select_first:
                first = self._first_leaf()
                if first is not None:
                    self._doc_tree.setCurrentItem(first)
                    return
            self._history.setRowCount(0)
            self._update_action_states()

    def _first_leaf(self) -> QTreeWidgetItem | None:
        """Return the first visible revision node, skipping filtered titles."""

        for index in range(self._doc_tree.topLevelItemCount()):
            found = self._first_visible_revision(self._doc_tree.topLevelItem(index))
            if found is not None:
                return found
        return None

    def _first_visible_revision(
        self, item: QTreeWidgetItem | None
    ) -> QTreeWidgetItem | None:
        if item is None or item.isHidden():
            return None
        if item.data(0, _ROLE_NODE_KIND) == _NODE_REVISION:
            return item
        for index in range(item.childCount()):
            found = self._first_visible_revision(item.child(index))
            if found is not None:
                return found
        return None

    def _find_tree_child(
        self,
        parent: QTreeWidgetItem | QTreeWidget,
        text: str,
        *,
        include_hidden: bool = False,
    ) -> QTreeWidgetItem | None:
        needle = text.casefold()
        if isinstance(parent, QTreeWidget):
            count = parent.topLevelItemCount()
            getter = parent.topLevelItem
        else:
            count = parent.childCount()
            getter = parent.child
        for index in range(count):
            child = getter(index)
            if child is None:
                continue
            if not include_hidden and child.isHidden():
                continue
            if child.text(0).casefold() == needle:
                return child
        return None

    def _latest_revision_item(
        self, mark_item: QTreeWidgetItem
    ) -> QTreeWidgetItem | None:
        best: QTreeWidgetItem | None = None
        best_rank: tuple[int, int, str] | None = None
        for index in range(mark_item.childCount()):
            child = mark_item.child(index)
            if child is None or child.isHidden():
                continue
            bundles = child.data(0, _ROLE_ROW) or ()
            files = [
                record
                for bundle in bundles
                if isinstance(bundle, DocumentBundle)
                for record in _bundle_records(bundle)
            ]
            rank = folder_revision_rank(files)
            if best is None or rank > best_rank:
                best = child
                best_rank = rank
        return best

    def _official_rd_package(self, title: str, mark: str) -> KitPackageRow | None:
        """Return the issued RD package for the official (non-working) revision.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.

        Returns:
            ``kit_package`` with ``is_current``, else the newest package
            whose filename revision matches ``official_revision_text``.
        """

        token = (
            id(self._all_records),
            id(self._kit_pipelines),
            self._official_ids_token,
        )
        key = kit_identity_key(title, mark)
        if getattr(self, "_official_rd_package_cache", None) != token:
            self._official_rd_package_cache = token
            self._official_rd_package_by_kit = {}
        cached = self._official_rd_package_by_kit
        if key in cached:
            return cached[key]
        try:
            packages = self.database.list_kit_packages(title, mark)
        except Exception:
            packages = []
        current = [
            pkg
            for pkg in packages
            if pkg.source == "rd" and not pkg.is_grey and pkg.is_current
        ]
        result: KitPackageRow | None
        if current:
            result = max(
                current,
                key=lambda pkg: (
                    pkg.sequence if pkg.sequence is not None else -1,
                    pkg.id or 0,
                ),
            )
        else:
            pipeline = self._kit_pipelines.get(key)
            official = (
                pipeline.official_revision_text if pipeline is not None else ""
            )
            if not official:
                result = None
            else:
                matched = [
                    pkg
                    for pkg in packages
                    if pkg.source == "rd"
                    and not pkg.is_grey
                    and revision_texts_equivalent(pkg.revision_text, official)
                ]
                result = (
                    max(
                        matched,
                        key=lambda pkg: (
                            pkg.sequence if pkg.sequence is not None else -1,
                            pkg.id or 0,
                        ),
                    )
                    if matched
                    else None
                )
        cached[key] = result
        return result

    def _official_rd_transfer_name(self, title: str, mark: str) -> str:
        """Return the issued-folder name of the official (non-working) RD package.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.

        Returns:
            ``transfer_name`` of the official package, or empty.
        """

        package = self._official_rd_package(title, mark)
        if package is None:
            return ""
        return package.transfer_name or ""

    def _official_rd_folders(self, row: KitMatrixRow) -> tuple[str, ...]:
        """Return official RD package dirs for robot sibling-MTO search."""

        package = self._official_rd_package(row.title, row.mark)
        if package is not None and package.package_path:
            return (package.package_path,)
        folders: list[str] = []
        seen: set[str] = set()
        for path in row.rd.paths:
            folder = issued_package_dir(str(path)) or str(Path(path).parent)
            key = folder.casefold()
            if folder and key not in seen:
                seen.add(key)
                folders.append(folder)
        return tuple(folders)

    def _jump_to_kit_documents(
        self,
        title: str,
        mark: str,
        transfer_name: str | None = None,
    ) -> None:
        """Hide via the title-mark filter and select the official (or given) NN.

        Switches to Все документы. Rebuilds the tree only when that tab is
        still deferred (first visit / after keep_view defer) or the tree
        widget is empty.

        Args:
            title: Kit title number.
            mark: Kit mark/system component.
            transfer_name: Issued folder name; when set, select that NN
                node. When omitted, select the official (non-working)
                package, not the newest working NN.
        """

        if not hasattr(self, "_documents_tab"):
            return
        tree_deferred = _DEFERRED_TREE in self._deferred_widgets
        with perf_span(
            "gui.jump_to_kit_documents",
            title=title,
            mark=mark,
            deferred=tree_deferred,
        ):
            needle = f"{title}-{mark}"
            if hasattr(self, "_doc_filter_timer"):
                self._doc_filter_timer.stop()
            self._doc_filter.blockSignals(True)
            self._doc_filter.setText(needle)
            self._doc_filter.blockSignals(False)
            self._tabs.setCurrentWidget(self._documents_tab)
            if _DEFERRED_TREE in self._deferred_widgets:
                self._ensure_deferred_widget(_DEFERRED_TREE)
            elif self._doc_tree.topLevelItemCount() == 0:
                self._rebuild_document_tree(select_first=False)
            else:
                self._sync_document_tree_filter_hidden()
            title_item = self._find_tree_child(self._doc_tree, title)
            if title_item is None:
                return
            mark_item = self._find_tree_child(title_item, mark)
            if mark_item is None:
                return
            revision_item = None
            if transfer_name:
                revision_item = self._revision_item_for_transfer(
                    mark_item, transfer_name
                )
            if revision_item is None:
                official_transfer = self._official_rd_transfer_name(title, mark)
                if official_transfer:
                    revision_item = self._revision_item_for_transfer(
                        mark_item, official_transfer
                    )
            if revision_item is None:
                revision_item = self._latest_revision_item(mark_item)
            if revision_item is None:
                return
            title_item.setExpanded(True)
            mark_item.setExpanded(True)
            self._doc_tree.setCurrentItem(revision_item)
            self._doc_tree.scrollToItem(revision_item)

    def _revision_item_for_transfer(
        self,
        mark_item: QTreeWidgetItem,
        transfer_name: str,
    ) -> QTreeWidgetItem | None:
        needle = transfer_name.casefold()
        if not needle:
            return None
        for index in range(mark_item.childCount()):
            child = mark_item.child(index)
            if child is None or child.isHidden():
                continue
            bundles = child.data(0, _ROLE_ROW) or ()
            for bundle in bundles:
                if not isinstance(bundle, DocumentBundle):
                    continue
                if (bundle.folder_key or "").casefold() == needle:
                    return child
                for record in _bundle_records(bundle):
                    if record_folder_key(record) == needle:
                        return child
        return None

    def _select_document_tree_node(
        self,
        title: str,
        mark: str,
        transfer_name: str | None = None,
    ) -> None:
        """Reselect a documents-tree node without changing the filter.

        Args:
            title: Title folder / kit title.
            mark: Mark node text.
            transfer_name: Optional transfer/folder key for the revision leaf.
        """

        if not hasattr(self, "_doc_tree"):
            return
        title_item = self._find_tree_child(self._doc_tree, title)
        if title_item is None:
            return
        title_item.setExpanded(True)
        mark_item = self._find_tree_child(title_item, mark) if mark else None
        if mark_item is None:
            self._doc_tree.setCurrentItem(title_item)
            self._doc_tree.scrollToItem(title_item)
            return
        mark_item.setExpanded(True)
        revision_item = None
        if transfer_name:
            revision_item = self._revision_item_for_transfer(mark_item, transfer_name)
        target = revision_item or mark_item
        self._doc_tree.setCurrentItem(target)
        self._doc_tree.scrollToItem(target)

    def _tree_item_records(self, item: QTreeWidgetItem) -> list[FileRecord]:
        """Return catalog files stored under a documents-tree node."""

        kind = item.data(0, _ROLE_NODE_KIND)
        if kind == _NODE_REVISION:
            bundles = item.data(0, _ROLE_ROW) or ()
            return [
                record
                for bundle in bundles
                if isinstance(bundle, DocumentBundle)
                for record in _bundle_records(bundle)
            ]
        records: list[FileRecord] = []
        for index in range(item.childCount()):
            child = item.child(index)
            if child is not None:
                records.extend(self._tree_item_records(child))
        return records

    def _tree_item_identity(
        self, item: QTreeWidgetItem
    ) -> tuple[str, str, str | None]:
        """Return title, mark, and optional transfer key for a tree node."""

        kind = item.data(0, _ROLE_NODE_KIND)
        transfer: str | None = None
        cursor: QTreeWidgetItem | None = item
        if kind == _NODE_REVISION:
            bundles = item.data(0, _ROLE_ROW) or ()
            for bundle in bundles:
                if isinstance(bundle, DocumentBundle) and bundle.folder_key:
                    transfer = bundle.folder_key
                    break
            cursor = item.parent()
            kind = cursor.data(0, _ROLE_NODE_KIND) if cursor is not None else None
        mark = ""
        if kind == _NODE_MARK and cursor is not None:
            mark = cursor.text(0)
            cursor = cursor.parent()
        title = cursor.text(0) if cursor is not None else item.text(0)
        if item.data(0, _ROLE_NODE_KIND) == _NODE_TITLE:
            title = item.text(0)
            mark = ""
        elif item.data(0, _ROLE_NODE_KIND) == _NODE_MARK:
            title = (
                item.parent().text(0) if item.parent() is not None else title
            )
            mark = item.text(0)
        return (title, mark, transfer)

    def _tree_item_sequence(self, item: QTreeWidgetItem) -> int | None:
        """Return the issued-folder NN for a documents-tree revision node."""

        sequences: list[int] = []
        for record in self._tree_item_records(item):
            value = record.data.get("transfer_sequence")
            try:
                if value is not None and str(value) != "":
                    sequences.append(int(value))
            except (TypeError, ValueError):
                continue
        if sequences:
            return max(sequences)
        _title, _mark, transfer = self._tree_item_identity(item)
        if not transfer:
            return None
        return parse_transfer_folder(transfer, under_gate=True).sequence

    def _working_flag_matches_node(
        self,
        flag,
        item: QTreeWidgetItem,
        transfer_name: str,
    ) -> bool:
        folder = (transfer_name or "").casefold()
        flag_folder = (flag.transfer_name or "").casefold()
        if flag_folder:
            return bool(folder and folder == flag_folder)
        sequence = self._tree_item_sequence(item)
        if (
            sequence is not None
            and flag.sequence is not None
            and int(sequence) == int(flag.sequence)
        ):
            return True
        return False

    def _tree_folder_is_working(
        self,
        item: QTreeWidgetItem,
        transfer_name: str,
        pipeline,
    ) -> bool:
        if pipeline is None:
            return False
        folder = (transfer_name or "").casefold()
        names = {str(name).casefold() for name in pipeline.working_transfer_names}
        if folder and folder in names:
            return True
        if names:
            return False
        sequence = self._tree_item_sequence(item)
        if sequence is not None and sequence in pipeline.working_sequences:
            return True
        return False

    def _rd_folder_for_tree_item(self, item: QTreeWidgetItem) -> str:
        """Return the RD folder to open or rescan for a documents-tree node.

        Args:
            item: Title, mark, or revision node.

        Returns:
            Directory under ``rd_root``, or empty string when unknown.
        """

        rd_root = str(self.config.rd_root)
        records = self._tree_item_records(item)
        folders: list[str] = []
        for record in records:
            package = issued_package_dir(record.path)
            folders.append(package or containing_folder(record.path))
        common = common_parent_dir(folders, under_root=rd_root)
        if common:
            return common
        kind = item.data(0, _ROLE_NODE_KIND)
        title, mark, _transfer = self._tree_item_identity(item)
        if kind == _NODE_TITLE and title:
            candidate = str(Path(rd_root) / title)
            if path_is_under(candidate, rd_root):
                return candidate
        if kind == _NODE_MARK and title and mark:
            candidate = str(Path(rd_root) / title / mark)
            if path_is_under(candidate, rd_root):
                return candidate
        return ""

    def _rescan_tree_item(self, item: QTreeWidgetItem) -> None:
        """Start a scoped RD scan of the folder represented by ``item``."""

        folder = self._rd_folder_for_tree_item(item)
        if not folder:
            QMessageBox.information(
                self,
                "Перескан папки",
                "Не удалось определить папку РД для этого пункта дерева.",
            )
            return
        if not path_is_under(folder, self.config.rd_root):
            QMessageBox.warning(
                self,
                "Перескан папки",
                "Папка лежит вне корня РД.",
            )
            return
        title, mark, transfer = self._tree_item_identity(item)
        folders = self._with_mixed_title_rd_subtrees(
            folder, title=title, mark=mark
        )
        self.start_scan(
            (SourceKind.RD,),
            rd_subtree=folders if len(folders) > 1 else folder,
            keep_view=True,
            restore_tree=(title, mark, transfer),
        )

    def _kit_rd_paths(self, row: KitMatrixRow) -> tuple[str, ...]:
        """Return present RD file paths for a kits row."""

        return self._kit_source_paths(
            row.title, row.mark, SourceKind.RD
        ) or tuple(row.rd.paths or ())

    def _live_kit_rd_mark_folder(self, title: str, mark: str) -> str:
        """Return a unique on-disk mark folder under ``rd_root / title``.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.

        Returns:
            Existing mark folder, or empty string when missing or ambiguous.
        """

        with perf_span("gui.live_kit_rd_mark_folder"):
            title_dir = Path(self.config.rd_root) / title
            if not title_dir.is_dir():
                return ""
            try:
                children = list(title_dir.iterdir())
            except OSError:
                return ""
            matched = [
                child
                for child in children
                if child.is_dir() and folder_matches_mark(child.name, mark)
            ]
            if len(matched) != 1:
                return ""
            return str(matched[0])

    def _kit_rd_title_folder(self, title: str) -> str:
        """Return ``rd_root / title``, or empty when it would be the RD root."""

        rd_root = str(self.config.rd_root)
        folder = constructed_kit_rd_title_folder(rd_root, title)
        if not folder or not path_is_under(folder, rd_root):
            return ""
        if path_is_under(rd_root, folder):
            return ""
        return folder

    def _is_usable_kit_mark_folder(self, folder: str, title_folder: str) -> bool:
        """Return whether ``folder`` is a mark directory under the title."""

        if not folder or not title_folder:
            return False
        if not path_is_under(folder, title_folder):
            return False
        if path_is_under(title_folder, folder):
            return False
        name = Path(folder).name
        if is_transfer_gate_folder_name(name):
            return False
        return True

    def _rd_folder_for_kit(self, row: KitMatrixRow) -> str:
        """Return the preferred mark folder for a kits row (may not exist).

        Prefers the unique mark folder from RD file paths (gates skipped).
        Otherwise a unique live child of the title folder, else constructed
        ``rd_root / title / mark``. Never a gate and never the title itself.

        Args:
            row: Selected комплекты row.

        Returns:
            Mark-folder path, or empty string when unknown.
        """

        title_folder = self._kit_rd_title_folder(row.title)
        rd_root = str(self.config.rd_root)
        folders = [
            folder
            for folder in unique_kit_rd_mark_folders(
                self._kit_rd_paths(row),
                rd_root,
                title=row.title,
            )
            if self._is_usable_kit_mark_folder(folder, title_folder)
        ]
        if len(folders) == 1:
            return folders[0]
        live = self._live_kit_rd_mark_folder(row.title, row.mark)
        if self._is_usable_kit_mark_folder(live, title_folder):
            return live
        constructed = constructed_kit_rd_mark_folder(
            rd_root, row.title, row.mark
        )
        if self._is_usable_kit_mark_folder(constructed, title_folder):
            return constructed
        return ""

    def _rd_rescan_subtree_for_kit(self, row: KitMatrixRow) -> str:
        """Return the RD folder to walk for a kit rescan.

        Uses the mark folder when it exists on disk. Otherwise the title
        folder. Never a gate and never the RD root.

        Args:
            row: Selected комплекты row.

        Returns:
            Directory under ``rd_root / title``, possibly not yet created.
        """

        with perf_span("gui.rd_rescan_subtree", title=row.title, mark=row.mark):
            title_folder = self._kit_rd_title_folder(row.title)
            mark_folder = self._rd_folder_for_kit(row)
            if (
                self._is_usable_kit_mark_folder(mark_folder, title_folder)
                and Path(mark_folder).is_dir()
            ):
                return mark_folder
            return title_folder

    def _with_mixed_title_rd_subtrees(
        self,
        primary: str,
        *,
        title: str,
        mark: str,
    ) -> tuple[str, ...]:
        """Return the kit folder plus foreign mixed-title mark folders.

        Extra folders come from present RD files of this kit whose issued
        path sits under another title. Never the RD root, never a gate.

        Args:
            primary: Mark or title folder already chosen for this kit.
            title: Filename title of the selected kit.
            mark: Latin AGCC mark of the selected kit.

        Returns:
            Unique walk roots, ``primary`` first when it is usable.
        """

        rd_root = str(self.config.rd_root)
        folders: list[str] = []
        seen: set[str] = set()

        def add(folder: str, *, allow_gate: bool = False) -> None:
            text = str(folder or "").strip()
            if not text:
                return
            if not path_is_under(text, rd_root):
                return
            if path_is_under(rd_root, text):
                return
            if not allow_gate and is_transfer_gate_folder_name(Path(text).name):
                return
            key = make_path_key(text)
            if key in seen:
                return
            if primary and key != make_path_key(primary) and path_is_under(
                text, primary
            ):
                return
            seen.add(key)
            folders.append(text)

        add(primary, allow_gate=True)
        if title and mark:
            for extra in mixed_title_rescan_folders(
                self._all_records,
                title=title,
                mark=mark,
                rd_root=rd_root,
            ):
                add(extra)
        return tuple(folders)

    def _rd_rescan_subtrees_for_kit(self, row: KitMatrixRow) -> tuple[str, ...]:
        """Return RD folders to walk for a kit point rescan.

        Args:
            row: Selected комплекты row.

        Returns:
            Kit mark/title folder plus foreign mixed-title folders.
        """

        primary = self._rd_rescan_subtree_for_kit(row)
        return self._with_mixed_title_rd_subtrees(
            primary, title=row.title, mark=row.mark
        )

    def _mixed_title_folders_for_kit(self, row: KitMatrixRow) -> tuple[str, ...]:
        """Return foreign title/mark folders for mixed-title files of ``row``.

        Args:
            row: Selected комплекты row.

        Returns:
            Unique folders under ``rd_root``, never the RD root itself.
        """

        return mixed_title_rescan_folders(
            self._all_records,
            title=row.title,
            mark=row.mark,
            rd_root=self.config.rd_root,
        )

    def _mixed_title_open_folders_for_kit(
        self, row: KitMatrixRow
    ) -> tuple[str, ...]:
        """Return parent folders of mixed-title files for Explorer.

        Args:
            row: Selected комплекты row.

        Returns:
            Unique directories that contain the stray files.
        """

        return mixed_title_open_folders(
            self._records_for_kit(row.title, row.mark),
            title=row.title,
            mark=row.mark,
            rd_root=self.config.rd_root,
        )

    def _open_mixed_title_folders(self, row: KitMatrixRow) -> None:
        """Open Explorer on each folder that holds a mixed-title file.

        Opens the file's parent (``PDF`` / ``DWG`` / ``NN_…``), one window
        per unique folder. Does not probe the disk before ``startfile``.

        Args:
            row: Selected комплекты row.
        """

        folders = self._mixed_title_open_folders_for_kit(row)
        if not folders:
            QMessageBox.information(
                self,
                "Смешанные папки",
                "У комплекта нет файлов в папке другого титула.",
            )
            return
        opened = 0
        missing: list[str] = []
        for folder in folders:
            ok, message = open_path(folder)
            if ok:
                opened += 1
            else:
                missing.append(message)
                self._append_log(message)
        if opened:
            self.statusBar().showMessage(
                f"Открыто смешанных папок: {opened}",
                8_000,
            )
        if missing and not opened:
            QMessageBox.warning(
                self,
                "Смешанные папки",
                "Не удалось открыть папки:\n" + "\n".join(missing),
            )
        elif missing:
            self.statusBar().showMessage(
                f"Открыто {opened}, не открылось: {len(missing)}",
                10_000,
            )

    def _rescan_kit_rd(self, row: KitMatrixRow) -> None:
        """Start a scoped RD scan of this kit's mark folder, else title."""

        if self._busy():
            QMessageBox.information(
                self,
                "Перескан РД",
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return
        with perf_span("gui.rescan_kit_rd", title=row.title, mark=row.mark):
            folders = self._rd_rescan_subtrees_for_kit(row)
            rd_root = str(self.config.rd_root)
            folder = folders[0] if folders else ""
            if not folder or not path_is_under(folder, rd_root):
                QMessageBox.information(
                    self,
                    "Перескан РД",
                    "Не удалось определить папку РД для комплекта "
                    f"{row.title}-{row.mark}.",
                )
                return
            if path_is_under(rd_root, folder):
                QMessageBox.warning(
                    self,
                    "Перескан РД",
                    "Перескан не поднимается выше папки титула.",
                )
                return
            self.start_scan(
                (SourceKind.RD,),
                rd_subtree=folders if len(folders) > 1 else folder,
                keep_view=True,
                restore_kit=(row.title, row.mark),
            )

    def _add_robot_handoff_action(self, menu: QMenu) -> QAction:
        """Append the debug «Передать роботу» item after a separator."""

        menu.addSeparator()
        action = menu.addAction("Передать роботу")
        action.setToolTip(
            "Скопировать снимок этого пункта (лейбл, титул–марка, пути) "
            "в буфер, чтобы вставить в чат."
        )
        return action

    def _copy_robot_handoff(self, text: str) -> None:
        """Copy a chat handoff dump and confirm in the status bar."""

        self._copy_text(text, message="Скопировано для чата с роботом")

    @staticmethod
    def _tree_path_labels(item: QTreeWidgetItem) -> str:
        """Return ``титул / марка / лейбл`` from a tree node to the root."""

        parts: list[str] = []
        cursor: QTreeWidgetItem | None = item
        while cursor is not None:
            parts.append(cursor.text(0))
            cursor = cursor.parent()
        return " / ".join(reversed(parts))

    @staticmethod
    def _table_click_cell(
        table: QTableWidget, row: int, column: int
    ) -> tuple[str, str]:
        header_item = table.horizontalHeaderItem(column) if column >= 0 else None
        header = header_item.text() if header_item is not None else ""
        cell = table.item(row, column) if row >= 0 and column >= 0 else None
        return (header, cell.text() if cell is not None else "")

    def _documents_tree_handoff_text(self, item: QTreeWidgetItem) -> str:
        """Build a chat dump for a documents-tree node.

        Args:
            item: Title, mark, or revision node.

        Returns:
            Multiline handoff text.
        """

        kind = str(item.data(0, _ROLE_NODE_KIND) or "")
        title, mark, transfer = self._tree_item_identity(item)
        files = self._tree_item_records(item)
        folder_name = folder_display_name(files)
        hint = self._folder_hint_for(title, mark, files) if files else None
        options = self._tree_label_options()
        is_current = False
        if kind == _NODE_REVISION:
            stored = item.data(0, _ROLE_ROW) or ()
            is_current = any(
                isinstance(bundle, DocumentBundle) and bundle.is_current
                for bundle in stored
            )
        elif files:
            is_current = any(record.id in self._detected_current_ids for record in files)
        has_mto = folder_has_mto(files)
        is_working, is_annulled = _hint_folder_marks(hint)
        rd_root = str(self.config.rd_root)
        layout_mismatch = any(
            not record_has_canonical_layout(record, rd_root) for record in files
        )
        pending = any(
            record.review_state is ReviewState.PENDING for record in files
        )
        file_lines: list[str] = []
        limit = 24
        for record in files[:limit]:
            kind_name = str(record.data.get("file_kind") or "file")
            file_lines.append(
                f"{kind_name}\t{_record_name(record)}\t{record.path}"
            )
        if len(files) > limit:
            file_lines.append(f"… ещё {len(files) - limit}")
        fields: list[tuple[str, str]] = [
            ("Вкладка", "Все документы"),
            ("Уровень", node_kind_label(kind)),
            ("Путь в дереве", self._tree_path_labels(item)),
            ("Лейбл сейчас", item.text(0)),
            ("Подсказка", item.toolTip(0)),
            ("Титул", title),
            ("Марка", mark),
            ("folder_key", transfer or ""),
            ("Папка NN", folder_name),
            ("Ревизия файлов", folder_revision_label(files) if files else ""),
            ("Дата сохранения", folder_latest_save_date(files)),
            ("Чекбоксы узла", format_label_options(options)),
            (
                "Статус Google (кратко)",
                hint.review_status if hint is not None else "",
            ),
            (
                "Статус Google (полный)",
                hint.review_full if hint is not None else "",
            ),
            (
                "Сопоставление Google",
                hint.match_reason if hint is not None else "",
            ),
            ("F", hint.f_label if hint is not None else ""),
            ("Статус MTO", _mto_node_label(hint)),
            ("Текущий состав", yes_no(is_current)),
            ("MTO в папке", yes_no(has_mto)),
            ("Рабочая ревизия", yes_no(is_working)),
            ("Аннулирована", yes_no(is_annulled)),
            ("Путь вне title/mark/gate/NN", yes_no(layout_mismatch)),
            ("Требует проверки (pending)", yes_no(pending)),
            ("Папка на диске", self._rd_folder_for_tree_item(item)),
            (
                "Фильтр дерева",
                self._doc_filter.text().strip() if hasattr(self, "_doc_filter") else "",
            ),
            (
                "Фильтр «Требует проверки»",
                yes_no(self._doc_pending.isChecked())
                if hasattr(self, "_doc_pending")
                else "",
            ),
            (
                "Фильтр «Только проблемы»",
                yes_no(self._doc_problems.isChecked())
                if hasattr(self, "_doc_problems")
                else "",
            ),
            ("Файлы", "\n".join(file_lines)),
        ]
        return format_robot_handoff(
            "дерево «Все документы» (титул → марка → ревизия)",
            fields,
        )

    def _history_handoff_text(self, row: int, column: int) -> str:
        """Build a chat dump for a documents history table row."""

        bundle = self._selected_bundle()
        tree_item = self._doc_tree.currentItem() if hasattr(self, "_doc_tree") else None
        header, cell = self._table_click_cell(self._history, row, column)
        row_values: list[str] = []
        if row >= 0:
            for index in range(self._history.columnCount()):
                name_item = self._history.horizontalHeaderItem(index)
                name = name_item.text() if name_item is not None else str(index)
                value_item = self._history.item(row, index)
                value = value_item.text() if value_item is not None else ""
                row_values.append(f"{name}={value}")
        files = _bundle_records(bundle) if bundle is not None else ()
        hint = (
            self._folder_hint_for(bundle.title, bundle.mark, list(files))
            if bundle is not None
            else None
        )
        is_working, is_annulled = _hint_folder_marks(hint)
        file_lines = [
            f"{record.data.get('file_kind') or 'file'}\t{_record_name(record)}\t{record.path}"
            for record in files
        ]
        fields = [
            ("Вкладка", "Все документы"),
            ("Таблица", "документы выбранной ревизии"),
            ("Лейбл узла дерева", tree_item.text(0) if tree_item is not None else ""),
            (
                "Путь в дереве",
                self._tree_path_labels(tree_item) if tree_item is not None else "",
            ),
            ("Ячейка", f"{header} = {cell}" if header else cell),
            ("Строка таблицы", " | ".join(row_values)),
            ("Титул", bundle.title if bundle is not None else ""),
            ("Марка", bundle.mark if bundle is not None else ""),
            ("Документ", bundle.core_stem if bundle is not None else ""),
            ("Тип", bundle.discipline if bundle is not None else ""),
            (
                "Рев. файла",
                bundle.revision_label if bundle is not None else "",
            ),
            ("Статус MTO", _mto_node_label(hint)),
            (
                "Рабочая папка",
                yes_no(is_working),
            ),
            (
                "Аннулирована",
                yes_no(is_annulled),
            ),
            (
                "Путь комплекта",
                _bundle_package_path(bundle) if bundle is not None else "",
            ),
            ("Файлы", "\n".join(file_lines)),
        ]
        return format_robot_handoff("таблица документов выбранной ревизии", fields)

    def _kits_row_handoff_text(self, row: int, column: int) -> str:
        """Build a chat dump for a комплекты table row."""

        kit = self._selected_kit_row()
        header, cell = self._table_click_cell(self._kits_table, row, column)
        pipeline = None
        if kit is not None:
            pipeline = self._kit_pipelines.get(kit_identity_key(kit.title, kit.mark))
        fields = [
            ("Вкладка", "Комплекты"),
            ("Таблица", "матрица комплектов"),
            ("Ячейка", f"{header} = {cell}" if header else cell),
            ("Титул", kit.title if kit is not None else ""),
            ("Марка", kit.mark if kit is not None else ""),
            ("Сводка", self._kit_display_summary_label(kit) if kit is not None else ""),
            (
                "Статус рассмотрения",
                pipeline_review_label(
                    pipeline,
                    events=(
                        kit.google.events
                        if kit is not None and kit.google is not None
                        else ()
                    ),
                    issuance=kit.issuance if kit is not None else None,
                    google=kit.google if kit is not None else None,
                )
                if pipeline is not None
                else "",
            ),
            (
                "Статус согласования",
                pipeline_approval_label(
                    pipeline,
                    google=kit.google if kit is not None else None,
                    issuance=kit.issuance if kit is not None else None,
                )
                if pipeline is not None
                else "",
            ),
            (
                "Рабочая рев. РД",
                format_working_rd_rev_label(
                    pipeline.working_revision_text,
                    pipeline.official_revision_text,
                    working_as_build=pipeline.working_as_build,
                )
                if pipeline is not None
                else "",
            ),
            (
                "РД · рев.",
                (
                    pipeline.official_revision_text
                    if pipeline is not None and pipeline.working_revision_text
                    else kit.rd.revision_text if kit is not None else ""
                ),
            ),
            (
                "MTO · рев.",
                kits_official_folder_mto_text(
                    rd_present=bool(kit is not None and kit.rd.present),
                    revision_text=(
                        self._rd_mto_overlay_by_kit().get(
                            kit_identity_key(kit.title, kit.mark),
                            ("", ""),
                        )[1]
                        if kit is not None
                        else ""
                    ),
                ),
            ),
            (
                "Google · рев.",
                kit.google.sheet_revision_text if kit is not None and kit.google else "",
            ),
            (
                "SQ · рев.",
                kit.sq.revision_text if kit is not None else "",
            ),
            (
                "Робот МТО · рев.",
                kit.robot.revision_text if kit is not None else "",
            ),
            (
                "Авто МТО",
                self._auto_mto_cell_text(kit) if kit is not None else "",
            ),
            (
                AUTO_MTO_COMPARE_STATUS_HEADER,
                (
                    self._auto_mto_compare_status_for(kit.title, kit.mark).text
                    if kit is not None
                    else ""
                ),
            ),
            (
                "Выдача · рев.",
                kit.issuance.revision_text if kit is not None and kit.issuance else "",
            ),
        ]
        return format_robot_handoff("таблица комплектов", fields)

    def _kit_package_handoff_text(self, row: int, column: int) -> str:
        """Build a chat dump for a kit-card package row."""

        package = self._selected_kit_package()
        kit = self._selected_kit_row()
        header, cell = self._table_click_cell(self._kits_package_table, row, column)
        row_values: list[str] = []
        if row >= 0:
            for index in range(self._kits_package_table.columnCount()):
                name_item = self._kits_package_table.horizontalHeaderItem(index)
                name = name_item.text() if name_item is not None else str(index)
                value_item = self._kits_package_table.item(row, index)
                value = (value_item.text() if value_item is not None else "").replace(
                    "\n", " · "
                )
                row_values.append(f"{name}={value}")
        fields = [
            ("Вкладка", "Комплекты · карточка"),
            ("Таблица", "пакеты NN"),
            ("Ячейка", f"{header} = {cell}" if header else cell),
            ("Титул", kit.title if kit is not None else ""),
            ("Марка", kit.mark if kit is not None else ""),
            ("Строка таблицы", " | ".join(row_values)),
            ("source", package.source if package is not None else ""),
            (
                "NN",
                f"{package.sequence:02d}"
                if package is not None and package.sequence is not None
                else "",
            ),
            (
                "Папка NN",
                package.transfer_name or "" if package is not None else "",
            ),
            (
                "Рев. файлов",
                package.revision_text if package is not None else "",
            ),
            (
                "Путь пакета",
                package.package_path if package is not None else "",
            ),
            (
                "grey",
                yes_no(bool(package.is_grey)) if package is not None else "",
            ),
            (
                "текущий пакет",
                yes_no(bool(package.is_current)) if package is not None else "",
            ),
        ]
        return format_robot_handoff("таблица пакетов карточки комплекта", fields)

    def _mto_row_handoff_text(self, row: int, column: int) -> str:
        """Build a chat dump for an MTO readiness row."""

        payload = self._selected_mto_row()
        header, cell = self._table_click_cell(self._mto_table, row, column)
        data = payload if isinstance(payload, dict) else {}
        fields = [
            ("Вкладка", "MTO · Готовность робота"),
            ("Ячейка", f"{header} = {cell}" if header else cell),
            ("Титул", str(data.get("title") or "")),
            ("Марка", str(data.get("mark") or "")),
            ("title_system", str(data.get("title_system") or "")),
            ("discipline_block", str(data.get("discipline_block") or "")),
            ("status", str(data.get("status") or "")),
            ("rd_path", str(data.get("rd_path") or "")),
            ("robot_path", str(data.get("robot_path") or "")),
        ]
        return format_robot_handoff("таблица MTO", fields)

    def _on_mto_worklist_prepare_menu(self, menu: QMenu) -> None:
        action = self._add_robot_handoff_action(menu)
        action.triggered.connect(self._on_mto_worklist_handoff)

    def _on_an_tab_prepare_menu(self, menu: QMenu) -> None:
        action = self._add_robot_handoff_action(menu)
        action.triggered.connect(self._on_an_tab_handoff)

    def _on_an_tab_handoff(self) -> None:
        payload = self._an_tab.selected_file()
        if payload is None:
            return
        self._copy_robot_handoff(self._an_tab_handoff_text(payload))

    def _an_tab_handoff_text(self, payload: AnMtoFile) -> str:
        """Build a chat dump for an AN finder row.

        Args:
            payload: Selected ``AnMtoFile`` from the АН table.

        Returns:
            Robot-handoff text.
        """

        fields = [
            ("Вкладка", "АН"),
            ("Титул", payload.title),
            ("Марка", payload.mark),
            ("Ревизия АН", payload.revision_text or "—"),
            ("Имя", payload.name),
            ("Папка", payload.parent_dir),
            ("Путь", payload.path),
        ]
        return format_robot_handoff("таблица АН", fields)

    def _on_rd_dump_tab_prepare_menu(self, menu: QMenu) -> None:
        action = self._add_robot_handoff_action(menu)
        action.triggered.connect(self._on_rd_dump_tab_handoff)

    def _on_rd_dump_tab_handoff(self) -> None:
        payload = self._rd_dump_tab.selected_file()
        if payload is None:
            return
        self._copy_robot_handoff(self._rd_dump_tab_handoff_text(payload))

    def _rd_dump_tab_handoff_text(self, payload: AnMtoFile) -> str:
        """Build a chat dump for an RD finder row.

        Args:
            payload: Selected ``AnMtoFile`` from the РД table.

        Returns:
            Robot-handoff text.
        """

        fields = [
            ("Вкладка", "РД"),
            ("Титул", payload.title),
            ("Марка", payload.mark),
            ("Ревизия", payload.revision_text or "—"),
            ("Вид", rd_dump_kind(payload)),
            ("Имя", payload.name),
            ("Папка", payload.parent_dir),
            ("Путь", payload.path),
        ]
        return format_robot_handoff("таблица РД", fields)

    def _on_mto_worklist_handoff(self) -> None:
        payload = self._mto_worklist_tab.selected_row()
        if payload is None:
            return
        self._copy_robot_handoff(self._mto_worklist_handoff_text(payload))

    def _mto_worklist_handoff_text(self, payload: MtoWorklistRow) -> str:
        """Build a chat dump for an MTO worklist row.

        Args:
            payload: Selected ``MtoWorklistRow`` from the worklist table.

        Returns:
            Multiline handoff text.
        """

        table = self._mto_worklist_table
        header, cell = self._table_click_cell(
            table, table.currentRow(), table.currentColumn()
        )
        gap_label = {
            "no_package": "нет физической папки пакета РД",
            "no_mto_file": "папка РД есть, файла MTO нет",
            "no_rd": "комплект есть в Google/Выдаче, но отсутствует в РД",
        }.get(payload.gap_kind, "нет")
        link = (
            f"F={'есть' if payload.has_f_status else 'нет'}; "
            f"РД={'есть' if payload.package_path else 'нет'}; "
            f"MTO={'есть' if payload.mto_path else 'нет'}"
        )
        fields = [
            ("Вкладка", "MTO · Перечень"),
            ("Ячейка", f"{header} = {cell}" if header else cell),
            ("Титул", payload.title),
            ("Марка", payload.mark),
            ("Ревизия (ключ Google F / РД)", payload.revision_text),
            (
                "Этап ревизии (Google F / Выдача)",
                f"{status_short_label(payload.status) or '—'} "
                f"[{payload.status or 'нет статуса'}]",
            ),
            ("Метки этапа", payload.letters or "—"),
            ("Связь F → РД → MTO", link),
            ("Есть комплект в Google", yes_no(payload.in_google)),
            ("Есть строка в Выдаче", yes_no(payload.in_issuance)),
            ("as-build", yes_no(payload.is_as_build)),
            ("текущая", yes_no(payload.is_current)),
            ("Флаг MTO в матрице", yes_no(payload.has_mto)),
            ("mto_path", payload.mto_path),
            ("package_path", payload.package_path),
            (
                "Разрыв данных",
                f"{gap_label} [{payload.gap_kind}]" if payload.gap_kind else gap_label,
            ),
            ("проблемы", ", ".join(payload.problem_kinds)),
        ]
        return format_robot_handoff("таблица MTO перечень", fields)

    def _collision_row_handoff_text(self, row: int, column: int) -> str:
        """Build a chat dump for a collision row."""

        payload = self._selected_collision_row()
        header, cell = self._table_click_cell(self._collision_table, row, column)
        data = payload if isinstance(payload, dict) else {}
        paths = [str(path) for path in data.get("paths") or ()]
        fields = [
            ("Вкладка", "Коллизии"),
            ("Ячейка", f"{header} = {cell}" if header else cell),
            ("kind", str(data.get("kind") or "")),
            ("document_key", str(data.get("document_key") or "")),
            ("message", str(data.get("message") or "")),
            ("Пути", "\n".join(paths)),
        ]
        return format_robot_handoff("таблица коллизий", fields)

    def _show_document_tree_context_menu(self, position) -> None:
        item = self._doc_tree.itemAt(position)
        if item is None:
            return
        self._doc_tree.setCurrentItem(item)
        folder = self._rd_folder_for_tree_item(item)
        scanning = self._busy()
        menu = QMenu(self)
        open_action = menu.addAction("Открыть папку")
        copy_action = menu.addAction("Копировать путь папки")
        rescan_action = menu.addAction("Пересканировать эту папку")
        open_action.setEnabled(bool(folder) and not scanning)
        copy_action.setEnabled(bool(folder))
        rescan_action.setEnabled(bool(folder) and not scanning)
        handoff_action = self._add_robot_handoff_action(menu)
        menu.setToolTipsVisible(True)
        mark_action = None
        unmark_action = None
        mark_annulled_action = None
        unmark_annulled_action = None
        legalize_f_action = None
        if item.data(0, _ROLE_NODE_KIND) == _NODE_REVISION:
            revision_text = folder_revision_label(self._tree_item_records(item))
            if revision_text == NO_REVISION_LABEL:
                revision_text = ""
            title, mark, transfer = self._tree_item_identity(item)
            if revision_text and title and mark:
                menu.addSeparator()
                mark_action = menu.addAction("Пометить папку как рабочую")
                unmark_action = menu.addAction("Снять пометку «рабочая»")
                mark_annulled_action = menu.addAction(
                    "Пометить папку как аннулированную"
                )
                unmark_annulled_action = menu.addAction(
                    "Снять пометку «аннулирована»"
                )
                mark_action.setCheckable(True)
                mark_annulled_action.setCheckable(True)
                mark_action.setEnabled(not scanning)
                unmark_action.setEnabled(False)
                mark_annulled_action.setEnabled(not scanning)
                unmark_annulled_action.setEnabled(False)
                if scanning:
                    mark_action.setEnabled(False)
                    unmark_action.setEnabled(False)
                    mark_annulled_action.setEnabled(False)
                    unmark_annulled_action.setEnabled(False)
                else:
                    try:
                        flags = self.database.list_working_flags(title, mark)
                    except Exception:
                        flags = []
                    transfer_name = transfer or ""
                    has_manual = any(
                        self._working_flag_matches_node(
                            flag, item, transfer_name
                        )
                        for flag in flags
                    )
                    pipeline = self._kit_pipelines.get(
                        kit_identity_key(title, mark)
                    )
                    folder_is_working = self._tree_folder_is_working(
                        item, transfer_name, pipeline
                    )
                    is_auto = bool(folder_is_working and not has_manual)
                    mark_action.setChecked(has_manual)
                    if has_manual:
                        unmark_action.setEnabled(True)
                        if is_auto:
                            unmark_action.setToolTip(
                                "Снимается только ручная пометка. "
                                "Если папка всё ещё выше выдачи, "
                                "статус «рабочая» останется."
                            )
                    elif is_auto:
                        unmark_action.setEnabled(False)
                        unmark_action.setToolTip(
                            "Снять можно только ручную пометку. "
                            "На диске ревизия выше выдачи — статус "
                            "«рабочая» останется."
                        )
                    try:
                        annulled_flags = self.database.list_annulled_flags(
                            title, mark
                        )
                    except Exception:
                        annulled_flags = []
                    has_annulled = any(
                        self._working_flag_matches_node(
                            flag, item, transfer_name
                        )
                        for flag in annulled_flags
                    )
                    mark_annulled_action.setChecked(has_annulled)
                    unmark_annulled_action.setEnabled(has_annulled)
                menu.addSeparator()
                legalize_f_action = menu.addAction(LEGALIZE_APPROVAL_ACTION)
                has_rd = any(
                    record.source is SourceKind.RD and record.present
                    for record in self._tree_item_records(item)
                )
                legalize_f_action.setEnabled(
                    bool(revision_text and title and mark and has_rd)
                    and not scanning
                )
                legalize_f_action.setToolTip(
                    "Пишет код А в столбец F листа «Контроль выдачи» "
                    f"(вместо TRM: {LEGALIZE_APPROVAL_TOKEN})."
                )
        chosen = exec_tracked_menu(
            menu, MENU_DOC_TREE, self._doc_tree.viewport().mapToGlobal(position)
        )
        if chosen == open_action:
            self._open_result(folder, folder=False)
        elif chosen == copy_action:
            self._copy_text(folder, message="Путь папки скопирован")
        elif chosen == rescan_action:
            self._rescan_tree_item(item)
        elif chosen == handoff_action:
            self._copy_robot_handoff(self._documents_tree_handoff_text(item))
        elif mark_action is not None and chosen == mark_action:
            self._apply_tree_working_flag(item, remove=False)
        elif unmark_action is not None and chosen == unmark_action:
            self._apply_tree_working_flag(item, remove=True)
        elif (
            mark_annulled_action is not None and chosen == mark_annulled_action
        ):
            self._apply_tree_annulled_flag(item, remove=False)
        elif (
            unmark_annulled_action is not None
            and chosen == unmark_annulled_action
        ):
            self._apply_tree_annulled_flag(item, remove=True)
        elif legalize_f_action is not None and chosen == legalize_f_action:
            self._open_legalize_approval_f_dialog(item)

    def _open_legalize_approval_f_dialog(self, item: QTreeWidgetItem) -> None:
        """Preview F до/F после and write code A to КСБ ИД.

        Args:
            item: Documents-tree revision node.
        """

        if self._busy():
            QMessageBox.information(
                self,
                LEGALIZE_APPROVAL_TITLE,
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return
        records = self._tree_item_records(item)
        title = ""
        mark = ""
        bundles = item.data(0, _ROLE_ROW) or ()
        for bundle in bundles:
            if isinstance(bundle, DocumentBundle) and bundle.title and bundle.mark:
                title = bundle.title
                mark = bundle.mark
                break
        if not title or not mark:
            title, mark, _transfer = self._tree_item_identity(item)
        revision = folder_revision_label(records)
        if revision == NO_REVISION_LABEL:
            revision = ""
        if not title or not mark or not revision:
            QMessageBox.warning(
                self,
                LEGALIZE_APPROVAL_TITLE,
                "У папки нет разобранных титула, марки и ревизии файла.",
            )
            return
        if not any(
            record.source is SourceKind.RD and record.present
            for record in records
        ):
            QMessageBox.warning(
                self,
                LEGALIZE_APPROVAL_TITLE,
                "Легализация F доступна только для папки с файлами РД.",
            )
            return
        kits = self._google_kits or self.database.list_google_kits()
        looked = comment_lookup_from_kits(kits)(title, mark)
        if looked is None:
            QMessageBox.warning(
                self,
                LEGALIZE_APPROVAL_TITLE,
                f"Нет комплекта {title}-{mark} в КСБ ИД. "
                "Новые строки не создаются.",
            )
            return
        date_text = f_line_date_from_mtime_ns(latest_save_mtime_ns(records))
        mto_revision = mto_revision_from_records(records)
        try:
            f_line = legalize_approval_f_line(
                date=date_text,
                revision=revision,
                mto_revision=mto_revision,
            )
            patch = build_journal_patch(
                looked.comment_raw,
                f_line,
                revision=revision,
                stage=LEGALIZE_APPROVAL_STAGE,
            )
        except ValueError as exc:
            QMessageBox.warning(self, LEGALIZE_APPROVAL_TITLE, str(exc))
            return
        dialog = FLegalizeDialog(
            title=title,
            mark=mark,
            revision=revision,
            patch=patch,
            parent=self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        job = dialog.job()
        if job is None:
            return
        self._start_google_journal_write(
            (job,),
            log_line=f"Легализация F {title}-{mark} рев. {revision}…",
        )

    def _apply_tree_working_flag(
        self,
        item: QTreeWidgetItem,
        *,
        remove: bool,
    ) -> None:
        """Persist or clear a manual working-folder mark for a tree node.

        Args:
            item: Documents-tree revision node.
            remove: True deletes the manual flag; False upserts it.
        """

        if self._busy():
            QMessageBox.information(
                self,
                "Рабочая ревизия",
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return
        title, mark, transfer = self._tree_item_identity(item)
        revision_text = folder_revision_label(self._tree_item_records(item))
        if revision_text == NO_REVISION_LABEL:
            revision_text = ""
        if not title or not mark or not revision_text:
            return
        transfer_name = transfer or ""
        if not transfer_name:
            bundles = item.data(0, _ROLE_ROW) or ()
            for bundle in bundles:
                if isinstance(bundle, DocumentBundle) and bundle.folder_key:
                    transfer_name = bundle.folder_key
                    break
        with perf_span(
            "gui.apply_tree_working_flag",
            title=title,
            mark=mark,
            remove=remove,
        ):
            try:
                if remove:
                    self.database.delete_working_flag(
                        title, mark, transfer_name=transfer_name
                    )
                else:
                    self.database.upsert_working_flag(
                        title,
                        mark,
                        revision_text,
                        transfer_name=transfer_name,
                        sequence=self._tree_item_sequence(item),
                    )
                rebuild_pipeline(
                    self.database,
                    records=self._pipeline_records(),
                    detected_current_ids=self._detected_current_ids,
                    kit_keys={kit_identity_key(title, mark)},
                    rd_root=self._pipeline_rd_root(),
                )
            except Exception as exc:
                QMessageBox.warning(self, "Рабочая ревизия", str(exc))
                return
            self._refresh_after_scoped_pipeline({kit_identity_key(title, mark)})

    def _apply_tree_annulled_flag(
        self,
        item: QTreeWidgetItem,
        *,
        remove: bool,
    ) -> None:
        """Persist or clear a manual annulled-folder mark for a tree node.

        Args:
            item: Documents-tree revision node.
            remove: True deletes the manual flag; False upserts it.
        """

        if self._busy():
            QMessageBox.information(
                self,
                "Аннулированная папка",
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return
        title, mark, transfer = self._tree_item_identity(item)
        revision_text = folder_revision_label(self._tree_item_records(item))
        if revision_text == NO_REVISION_LABEL:
            revision_text = ""
        if not title or not mark or not revision_text:
            return
        transfer_name = transfer or ""
        if not transfer_name:
            bundles = item.data(0, _ROLE_ROW) or ()
            for bundle in bundles:
                if isinstance(bundle, DocumentBundle) and bundle.folder_key:
                    transfer_name = bundle.folder_key
                    break
        with perf_span(
            "gui.apply_tree_annulled_flag",
            title=title,
            mark=mark,
            remove=remove,
        ):
            try:
                if remove:
                    self.database.delete_annulled_flag(
                        title, mark, transfer_name=transfer_name
                    )
                else:
                    self.database.upsert_annulled_flag(
                        title,
                        mark,
                        revision_text,
                        transfer_name=transfer_name,
                        sequence=self._tree_item_sequence(item),
                    )
                rebuild_pipeline(
                    self.database,
                    records=self._pipeline_records(),
                    detected_current_ids=self._detected_current_ids,
                    kit_keys={kit_identity_key(title, mark)},
                    rd_root=self._pipeline_rd_root(),
                )
            except Exception as exc:
                QMessageBox.warning(self, "Аннулированная папка", str(exc))
                return
            self._refresh_after_scoped_pipeline({kit_identity_key(title, mark)})

    def _refresh_after_scoped_pipeline(
        self,
        kit_keys: set[tuple[str, str]],
        *,
        relabel_tree: bool = True,
    ) -> None:
        """Refresh derived GUI surfaces after a one-kit pipeline write.

        Does not rebuild the issuance journal, АН tab, or collisions.

        Args:
            kit_keys: Identities whose derived tables just changed.
            relabel_tree: Repaint matching documents-tree nodes when the
                tree is live.
        """

        with perf_span("gui.refresh_after_scoped_pipeline", kits=len(kit_keys)):
            self._load_folder_tree_hints()
            self._reload_mto_worklist_rows()
            self._refresh_kits_table(kit_keys=kit_keys)
            self._refresh_revision_matrix(kit_keys=kit_keys)
            if _DEFERRED_WORKLIST not in self._deferred_widgets:
                self._refresh_mto_worklist(reload=False)
            if relabel_tree:
                self._relabel_document_tree(kit_keys)

    def _google_kit_for(self, title: str, mark: str) -> GoogleKit | None:
        key = kit_identity_key(title, mark)
        for kit in self._google_kits:
            if kit_identity_key(kit.title, kit.mark) == key:
                return kit
        return None

    def _record_by_path_key(self, path: str) -> FileRecord | None:
        key = make_path_key(path)
        if not key:
            return None
        self._ensure_catalog_record_indexes()
        hit = self._records_by_path_key.get(key)
        if hit is not None:
            return hit
        return self._records_by_path_key.get(str(path).casefold())

    def _code_letter_for_record(self, record: FileRecord) -> tuple[str, str]:
        """Return last F code A/B/C date and stage for this MTO revision."""

        title = str(record.data.get("title") or "").strip()
        mark = str(record.data.get("mark") or "").strip()
        revision = file_revision_label(record)
        date_text, stage = last_code_letter_for_revision(
            self._google_kit_for(title, mark), revision
        )
        if date_text and stage:
            return date_text, stage
        pipeline = self._kit_pipelines.get(kit_identity_key(title, mark))
        if pipeline is None:
            return "", ""
        letter = (pipeline.code or "").upper()
        mapped = {"A": "code_a", "B": "code_b", "C": "code_c"}.get(letter, "")
        if not mapped:
            return "", ""
        stored = str(pipeline.code_date or "").strip()
        if not stored:
            return "", ""
        code_rev = str(pipeline.code_revision_text or "").strip()
        if (
            revision
            and code_rev
            and not revision_texts_equivalent(revision, code_rev)
        ):
            return "", ""
        return stored, mapped

    def _history_folder_records(self) -> list[FileRecord]:
        """Return catalog files in the selected documents-tree revision node."""

        if not hasattr(self, "_doc_tree"):
            return []
        selected = self._doc_tree.selectedItems()
        item = selected[0] if selected else None
        if item is None or item.data(0, _ROLE_NODE_KIND) != _NODE_REVISION:
            return []
        stored = item.data(0, _ROLE_ROW) or []
        files: list[FileRecord] = []
        for bundle in stored:
            if isinstance(bundle, DocumentBundle):
                files.extend(_bundle_records(bundle))
        return files

    def _package_folder_records(self, record: FileRecord) -> list[FileRecord]:
        """Return present catalog files in the same issued package as ``record``."""

        package = issued_package_dir(record.path)
        if not package:
            return [record]
        pkg_key = make_path_key(package)
        title = str(record.data.get("title") or "").strip()
        mark = str(record.data.get("mark") or "").strip()
        pool = (
            self._records_for_kit(title, mark)
            if title and mark
            else self._all_records
        )
        files: list[FileRecord] = []
        for candidate in pool:
            if not candidate.present:
                continue
            other = issued_package_dir(candidate.path)
            if other and make_path_key(other) == pkg_key:
                files.append(candidate)
        return files or [record]

    def _catalog_date_targets(
        self,
        record: FileRecord,
        folder_files: Sequence[FileRecord] | None = None,
    ) -> list[FileRecord]:
        """Return the outlier document files to stamp or clear together."""

        folder = list(folder_files) if folder_files is not None else self._package_folder_records(record)
        keys = same_document_path_keys(record, folder)
        by_key: dict[str, FileRecord] = {}
        for candidate in folder:
            key = str(candidate.path_key or "").casefold()
            if key in keys:
                by_key[key] = candidate
        current_key = str(record.path_key or "").casefold()
        if current_key and current_key not in by_key:
            by_key[current_key] = record
        return list(by_key.values()) or [record]

    def _folder_mean_for_record(
        self,
        record: FileRecord,
        folder_files: Sequence[FileRecord] | None = None,
    ) -> FolderMeanOverride | None:
        """Return the package mean date after dropping the current document."""

        folder = list(folder_files) if folder_files is not None else self._package_folder_records(record)
        return folder_mean_override(
            folder,
            exclude_path_keys=same_document_path_keys(record, folder),
        )

    def _record_has_mtime_override(self, record: FileRecord) -> bool:
        if record.data.get("mtime_override_applied") or record.data.get(
            "mtime_override_stale"
        ):
            return True
        try:
            return self.database.get_file_mtime_override(record.path_key) is not None
        except Exception:
            return False

    def _add_mtime_override_menu(
        self,
        menu: QMenu,
        mto_record: FileRecord | None,
        *,
        folder_files: Sequence[FileRecord] | None = None,
        hide_when_missing: bool = False,
    ) -> _MtimeOverrideMenu:
        """Append catalog-date actions on the documents table (not Комплекты)."""

        actions = _MtimeOverrideMenu()
        if mto_record is None and hide_when_missing:
            return actions
        menu.addSeparator()
        code_date, code_stage = (
            self._code_letter_for_record(mto_record)
            if mto_record is not None
            else ("", "")
        )
        code_label = code_letter_label(code_stage)
        mean = (
            self._folder_mean_for_record(mto_record, folder_files)
            if mto_record is not None
            else None
        )
        targets = (
            self._catalog_date_targets(mto_record, folder_files)
            if mto_record is not None
            else []
        )
        has_override = any(
            bool(
                item.data.get("mtime_override_applied")
                or item.data.get("mtime_override_stale")
            )
            for item in targets
        )
        if mto_record is not None and not has_override:
            has_override = self._record_has_mtime_override(mto_record)
        busy = self._busy()
        code_action = menu.addAction(MTO_CATALOG_DATE_ACTION)
        code_action.setEnabled(bool(mto_record is not None and code_date) and not busy)
        if mto_record is None:
            code_action.setToolTip("Нет файла MTO РД.")
        elif not code_date:
            code_action.setToolTip(
                "Нет даты кода A/B/C для этой ревизии в журнале F."
            )
        else:
            code_action.setToolTip(
                f"Дата каталога MTO → {code_date} ({code_label}). "
                "Дата на диске сохранится в подсказке."
            )
        manual_action = menu.addAction(MTO_CATALOG_DATE_MANUAL_ACTION)
        manual_action.setEnabled(mto_record is not None and not busy)
        if mto_record is None:
            manual_action.setToolTip("Нет файла MTO РД.")
        else:
            manual_action.setToolTip(
                "Задать дату каталога вручную. Файл на диске не меняется."
            )
        folder_action = menu.addAction(MTO_CATALOG_DATE_FOLDER_ACTION)
        folder_action.setEnabled(mto_record is not None and mean is not None and not busy)
        if mto_record is None:
            folder_action.setToolTip("Нет файла MTO РД.")
        elif mean is None:
            folder_action.setToolTip(
                "В папке нет других файлов для средней даты "
                "(текущий документ не учитывается)."
            )
        else:
            folder_action.setToolTip(
                f"Средняя дата {mean.used_count} файлов папки "
                f"(без текущего документа): {mean.override_date}."
            )
        clear_action = menu.addAction("Сбросить дату каталога")
        clear_action.setEnabled(bool(has_override) and not busy)
        return _MtimeOverrideMenu(
            code=code_action,
            manual=manual_action,
            folder=folder_action,
            clear=clear_action,
        )

    def _apply_catalog_date_overrides(
        self,
        records: Sequence[FileRecord],
        override_date: str,
        *,
        reason: str,
    ) -> bool:
        """Persist catalog dates and refresh the kit. Return False on hard fail."""

        if self._busy():
            QMessageBox.information(
                self,
                "Дата MTO",
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return False
        if not records:
            return False
        applied: list[FileRecord] = []
        errors: list[str] = []
        for record in records:
            try:
                override = self.database.upsert_file_mtime_override(
                    record.path_key, override_date, reason=reason
                )
            except Exception as exc:
                errors.append(f"{_record_name(record)}: {exc}")
                continue
            apply_mtime_override_to_data(record.path_key, record.data, override)
            applied.append(record)
        if not applied:
            QMessageBox.warning(
                self,
                "Дата MTO",
                "Не удалось заменить дату.\n" + "\n".join(errors),
            )
            return False
        self._refresh_after_mtime_override(applied[0])
        if errors:
            QMessageBox.warning(
                self,
                "Дата MTO",
                "Дата записана частично:\n" + "\n".join(errors),
            )
        return True

    def _confirm_mto_catalog_date(self, record: FileRecord) -> None:
        """Persist catalog mtime = last F code A/B/C date for one MTO xlsx."""

        if str(record.data.get("file_kind") or "") != FileKind.MTO_XLSX.value:
            QMessageBox.information(
                self, "Дата MTO", "Замена даты кода A/B/C только для файла MTO xlsx."
            )
            return
        code_date, code_stage = self._code_letter_for_record(record)
        letter = {"code_a": "A", "code_b": "B", "code_c": "C"}.get(code_stage, "A/B/C")
        if not code_date:
            QMessageBox.information(
                self,
                "Дата MTO",
                "Нет даты кода A/B/C для этой ревизии в журнале F.",
            )
            return
        name = _record_name(record)
        disk_ns = int(record.data.get("disk_mtime_ns") or record.data.get("mtime_ns") or 0)
        disk_text = format_file_save_date(disk_ns, with_time=True) or "—"
        reply = QMessageBox.question(
            self,
            "Дата MTO",
            (
                f"Заменить дату каталога на дату кода {letter} ({code_date})?\n\n"
                f"Файл: {name}\n"
                f"На диске: {disk_text}\n\n"
                "Файл на диске не меняется. Перескан того же файла "
                "(размер и mtime) сохранит замену. Старая дата остаётся "
                "в подсказке."
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        if self._apply_catalog_date_overrides(
            (record,), code_date, reason=code_stage
        ):
            self.statusBar().showMessage(
                f"Дата каталога MTO → {code_date} (на диске {disk_text})",
                5000,
            )

    def _confirm_mto_catalog_date_manual(
        self,
        record: FileRecord,
        *,
        folder_files: Sequence[FileRecord] | None = None,
    ) -> None:
        """Persist a typed catalog date for the current document."""

        targets = self._catalog_date_targets(record, folder_files)
        disk_ns = int(record.data.get("disk_mtime_ns") or record.data.get("mtime_ns") or 0)
        disk_text = format_file_save_date(disk_ns, with_time=True) or "—"
        default = date.today()
        applied = str(record.data.get("mtime_override_date") or "").strip()
        if record.data.get("mtime_override_applied") and applied:
            try:
                default = parse_override_date(applied)
            except ValueError:
                pass
        else:
            from_mtime = override_date_text_from_mtime_ns(
                int(record.data.get("mtime_ns") or disk_ns or 0)
            )
            if from_mtime:
                try:
                    default = parse_override_date(from_mtime)
                except ValueError:
                    pass
        chosen = _ask_catalog_override_date(
            self, default_date=default, disk_text=disk_text
        )
        if not chosen:
            return
        if self._apply_catalog_date_overrides(targets, chosen, reason="manual"):
            self.statusBar().showMessage(
                f"Дата каталога → {chosen} (вручную)",
                5000,
            )

    def _confirm_mto_catalog_date_folder_mean(
        self,
        record: FileRecord,
        *,
        folder_files: Sequence[FileRecord] | None = None,
    ) -> None:
        """Persist the mean disk date of other files in the current package."""

        mean = self._folder_mean_for_record(record, folder_files)
        if mean is None:
            QMessageBox.information(
                self,
                "Дата MTO",
                "В папке нет других файлов для средней даты.",
            )
            return
        targets = self._catalog_date_targets(record, folder_files)
        names = ", ".join(_record_name(item) for item in targets)
        reply = QMessageBox.question(
            self,
            "Дата MTO",
            (
                f"Заменить дату каталога на среднюю по папке "
                f"({mean.override_date})?\n\n"
                f"Учтено файлов: {mean.used_count}\n"
                f"Без текущего документа: {names or _record_name(record)}\n\n"
                "Файл на диске не меняется. Старая дата остаётся в подсказке."
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        if self._apply_catalog_date_overrides(
            targets, mean.override_date, reason="folder_mean"
        ):
            self.statusBar().showMessage(
                f"Дата каталога → {mean.override_date} "
                f"(средняя по {mean.used_count} файлам папки)",
                5000,
            )

    def _clear_mto_catalog_date(
        self,
        record: FileRecord,
        *,
        folder_files: Sequence[FileRecord] | None = None,
    ) -> None:
        """Drop catalog-date overrides for the current document."""

        if self._busy():
            QMessageBox.information(
                self,
                "Дата MTO",
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return
        targets = self._catalog_date_targets(record, folder_files)
        disk_ns = int(record.data.get("disk_mtime_ns") or record.data.get("mtime_ns") or 0)
        disk_text = format_file_save_date(disk_ns, with_time=True) or "—"
        reply = QMessageBox.question(
            self,
            "Дата MTO",
            (
                f"Сбросить дату каталога для {_record_name(record)}?\n"
                f"Вернётся дата с диска: {disk_text}."
            ),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        try:
            for item in targets:
                self.database.delete_file_mtime_override(item.path_key)
                apply_mtime_override_to_data(item.path_key, item.data, None)
        except Exception as exc:
            QMessageBox.warning(self, "Дата MTO", str(exc))
            return
        self._refresh_after_mtime_override(record)
        self.statusBar().showMessage("Дата каталога MTO сброшена на дату диска", 4000)

    def _refresh_after_mtime_override(self, record: FileRecord) -> None:
        title = str(record.data.get("title") or "").strip()
        mark = str(record.data.get("mark") or "").strip()
        with perf_span(
            "gui.apply_file_mtime_override",
            title=title,
            mark=mark,
        ):
            if title and mark:
                try:
                    rebuild_pipeline(
                        self.database,
                        records=self._pipeline_records(),
                        detected_current_ids=self._detected_current_ids,
                        kit_keys={kit_identity_key(title, mark)},
                        rd_root=self._pipeline_rd_root(),
                    )
                except Exception as exc:
                    self._append_log(
                        f"mtime override pipeline: {type(exc).__name__}: {exc}"
                    )
            self._official_ids_token = None
            self._load_folder_tree_hints()
            self._refresh_kits_table()
            if _DEFERRED_WORKLIST not in self._deferred_widgets:
                self._reload_mto_worklist_rows()
                self._refresh_mto_worklist(reload=False)
            if _DEFERRED_TREE not in self._deferred_widgets:
                self._relabel_tree_nodes_for_record(record)
                self._refresh_history()

    def _relabel_tree_nodes_for_record(self, record: FileRecord) -> None:
        if not hasattr(self, "_doc_tree"):
            return
        path_key = record.path_key.casefold()
        for item in self._iter_document_tree_items():
            if item.data(0, _ROLE_NODE_KIND) != _NODE_REVISION:
                continue
            files = self._tree_item_records(item)
            if not any(item_record.path_key.casefold() == path_key for item_record in files):
                continue
            self._relabel_revision_tree_item(item)

    def _iter_document_tree_items(self):
        def walk(node: QTreeWidgetItem | None):
            if node is None:
                return
            yield node
            for index in range(node.childCount()):
                yield from walk(node.child(index))

        root = self._doc_tree.invisibleRootItem()
        yield from walk(root)

    def _relabel_revision_tree_item(self, item: QTreeWidgetItem) -> None:
        mark_item = item.parent()
        title_item = mark_item.parent() if mark_item is not None else None
        if mark_item is None or title_item is None:
            return
        bundles = [
            bundle
            for bundle in (item.data(0, _ROLE_ROW) or ())
            if isinstance(bundle, DocumentBundle)
        ]
        files = [
            record
            for bundle in bundles
            for record in _bundle_records(bundle)
        ]
        if not files:
            return
        title, mark, _transfer = self._tree_item_identity(item)
        hint = self._folder_hint_for(title, mark, files)
        self._paint_revision_tree_item(
            item,
            title_item,
            mark_item,
            bundles,
            files=files,
            hint=hint,
        )
        self._sync_tree_layout_mismatch_ancestors(title_item, mark_item)

    @Slot()
    def _refresh_history(self) -> None:
        selected = self._doc_tree.selectedItems()
        item = selected[0] if selected else None
        bundles: list[DocumentBundle] = []
        if (
            item is not None
            and item.data(0, _ROLE_NODE_KIND) == _NODE_REVISION
        ):
            stored = item.data(0, _ROLE_ROW) or []
            bundles = [bundle for bundle in stored if isinstance(bundle, DocumentBundle)]
        hint: FolderTreeHint | None = None
        if bundles:
            folder_files = [
                record
                for bundle in bundles
                for record in _bundle_records(bundle)
            ]
            hint = self._folder_hint_for(
                bundles[0].title, bundles[0].mark, folder_files
            )
        mto_label = _mto_node_label(hint)
        self._reload_export_pins()
        self._history.setSortingEnabled(False)
        self._history.setRowCount(0)
        folder_latest_mtime = max(
            (latest_save_mtime_ns(_bundle_records(bundle)) for bundle in bundles),
            default=0,
        )
        for bundle in bundles:
            files = _bundle_records(bundle)
            package_path = _bundle_package_path(bundle)
            row_mtime = latest_save_mtime_ns(files)
            date_text = format_file_save_date(row_mtime) or "—"
            if bundle.pdf is None:
                current = "—"
            elif bundle.is_current:
                current = "Да"
            else:
                current = "Нет"
            display_record = _bundle_open_record(bundle)
            display_name = (
                _record_name(display_record) if display_record is not None else "—"
            )
            is_working, is_annulled = _hint_folder_marks(hint)
            working_origin = hint.working_origin if hint is not None else ""
            working_text = "Да" if is_working else "Нет"
            annulled_text = "Да" if is_annulled else "Нет"
            values = [
                bundle.core_stem or display_name,
                bundle.discipline or "—",
                "Да" if bundle.pdf is not None else "Нет",
                _editable_cell_text(bundle),
                current,
                "Есть" if any(record.present for record in files) else "Отсутствует",
                display_name,
                date_text,
                file_revision_label(display_record) or "—"
                    if display_record is not None
                    else "—",
                "Да" if folder_has_as_build(files) else "Нет",
                working_text,
                annulled_text,
                self._pin_view_for(bundle.title, bundle.mark).text,
                mto_label,
                package_path or "—",
            ]
            table_row = self._history.rowCount()
            self._history.insertRow(table_row)
            pin_view = self._pin_view_for(bundle.title, bundle.mark)
            working_tip = working_folder_tooltip(
                is_working=is_working, origin=working_origin
            )
            for column, value in enumerate(values):
                cell = QTableWidgetItem(str(value))
                cell.setData(_ROLE_ROW, bundle)
                if column == _HISTORY_COL_PIN:
                    self._apply_pin_view(cell, pin_view)
                if column == _HISTORY_COL_PACKAGE:
                    cell.setToolTip(
                        _qt_tooltip(package_path or "Папка комплекта неизвестна")
                    )
                if column == _HISTORY_COL_DATE:
                    tips = [
                        file_save_date_tooltip(files)
                        or "Дата сохранения неизвестна"
                    ]
                    if folder_latest_mtime > 0 and row_mtime == folder_latest_mtime:
                        cell.setBackground(QBrush(_LATEST_FILE_DATE_BG))
                        tips.append("Самая свежая дата в этой папке.")
                    cell.setToolTip("\n".join(tips))
                if column == _HISTORY_COL_WORKING:
                    if is_working:
                        _paint_pipeline_cell(
                            cell,
                            color_for(self._status_colors, "working"),
                        )
                    if working_tip:
                        cell.setToolTip(working_tip)
                if column == _HISTORY_COL_ANNULLED:
                    if is_annulled:
                        _paint_pipeline_cell(
                            cell,
                            color_for(self._status_colors, "annulled"),
                        )
                        cell.setToolTip(ANNULLED_TOOLTIP)
                if column == _HISTORY_COL_MTO_STATUS:
                    tips = [_MTO_STATUS_COLUMN_TIP]
                    if working_tip:
                        tips.append(working_tip)
                    cell.setToolTip("\n".join(tips))
                    if hint is not None and hint.mto_status:
                        _paint_pipeline_cell(
                            cell,
                            color_for(self._status_colors, hint.mto_status),
                        )
                self._history.setItem(table_row, column, cell)
        self._history.setSortingEnabled(True)
        if self._history.rowCount():
            self._history.selectRow(0)
        self._update_action_states()

    def _selected_mto_row(self) -> dict[str, Any] | None:
        row = self._mto_table.currentRow()
        item = self._mto_table.item(row, 0) if row >= 0 else None
        return item.data(_ROLE_ROW) if item else None

    def _selected_collision_row(self) -> dict[str, Any] | None:
        row = self._collision_table.currentRow()
        item = self._collision_table.item(row, 0) if row >= 0 else None
        return item.data(_ROLE_ROW) if item else None

    def _selected_bundle(self) -> DocumentBundle | None:
        """Return the DocumentBundle stored on the current history row."""

        if not hasattr(self, "_history"):
            return None
        row = self._history.currentRow()
        item = self._history.item(row, 0) if row >= 0 else None
        payload = item.data(_ROLE_ROW) if item else None
        return payload if isinstance(payload, DocumentBundle) else None

    def _selected_record(self) -> FileRecord | None:
        bundle = self._selected_bundle()
        if bundle is None:
            return None
        return _bundle_open_record(bundle)

    @Slot()
    def _update_action_states(self) -> None:
        scanning = self._busy()
        bundle = self._selected_bundle()
        record = _bundle_open_record(bundle) if bundle is not None else None
        selected = record is not None
        available = bool(record and record.present and record.path)
        for button in (
            getattr(self, "_ack_button", None),
            getattr(self, "_comment_button", None),
            getattr(self, "_ignore_button", None),
        ):
            if button is not None:
                button.setEnabled(selected and not scanning)
        for button in (
            getattr(self, "_history_open_button", None),
            getattr(self, "_history_folder_button", None),
        ):
            if button is not None:
                button.setEnabled(available)
        if hasattr(self, "_copy_button"):
            self._copy_button.setEnabled(bool(bundle and _bundle_paths(bundle)))

        pair = self._selected_mto_row() if hasattr(self, "_mto_table") else None
        for side in ("rd", "robot"):
            path = pair.get(f"{side}_path") if pair else None
            present = bool(pair and pair.get(f"{side}_present"))
            for folder in (False, True):
                button = self._mto_buttons.get((side, folder))
                if button:
                    button.setEnabled(bool(path and present))
        collision = (
            self._selected_collision_row()
            if hasattr(self, "_collision_table")
            else None
        )
        paths = list(collision.get("paths") or ()) if collision else []
        if hasattr(self, "_collision_folder_button"):
            self._collision_folder_button.setEnabled(len(paths) == 1)
            self._collision_copy_button.setEnabled(bool(paths))

        kit = self._selected_kit_row()
        for source in ("rd", "robot", "sq"):
            snapshot = self._kit_snapshot_for(kit, source) if kit else None
            enabled = bool(snapshot and snapshot.present and snapshot.paths)
            for folder in (False, True):
                button = self._kit_buttons.get((source, folder))
                if button:
                    button.setEnabled(enabled)
        if hasattr(self, "_kit_copy_button"):
            has_paths = bool(
                kit
                and (kit.rd.paths or kit.robot.paths or kit.sq.paths)
            )
            self._kit_copy_button.setEnabled(has_paths)
        package = (
            self._selected_kit_package()
            if hasattr(self, "_kits_package_table")
            else None
        )
        if hasattr(self, "_kits_pkg_open_button"):
            open_ok = False
            copy_ok = False
            if package is not None:
                open_ok = (
                    bool(package.package_path) and not package.is_grey
                )
                copy_ok = bool(package.package_path)
            self._kits_pkg_open_button.setEnabled(open_ok)
            self._kits_pkg_copy_button.setEnabled(copy_ok)
            self._kits_pkg_jump_button.setEnabled(package is not None)
            self._kits_pkg_handoff_button.setEnabled(package is not None)
        if hasattr(self, "_an_tab"):
            self._an_tab.set_scan_enabled(not scanning)
        if hasattr(self, "_rd_dump_tab"):
            self._rd_dump_tab.set_scan_enabled(not scanning)
        if (
            self._scan_thread is not None
            or self._an_scan_thread is not None
            or self._rd_dump_scan_thread is not None
        ):
            pass
        elif self._mto_compare_thread is not None:
            self._cancel_action.setEnabled(True)
        if hasattr(self, "_kits_de_sync_button"):
            self._kits_de_sync_button.setEnabled(not scanning)
        if hasattr(self, "_approval_mail_tab"):
            self._approval_mail_tab.set_catalog_busy(self._busy())
        if hasattr(self, "_revision_matrix_tab"):
            self._revision_matrix_tab.set_auto_mto_queue_paused(
                self._catalog_workers_busy()
            )
        if hasattr(self, "_an_tab"):
            self._an_tab.set_content_queue_paused(self._catalog_workers_busy())
        if hasattr(self, "_rd_dump_tab"):
            self._rd_dump_tab.set_content_queue_paused(self._catalog_workers_busy())
        self._set_transfer_mto_compare_paused(self._catalog_workers_busy())
        self._update_layout_report_action()

    def _catalog_workers_busy(self) -> bool:
        """Return True when a scan / АН / РД dump / Google / robot / SQ / export / PI worker is active.

        AutoMTO, AN content compares, and transfer-review MTO compares are
        not catalog workers: scan and Google stay available while those
        queues run. The WEB child process is also not a catalog worker.
        """

        export_copy = (
            hasattr(self, "_revision_matrix_tab")
            and self._revision_matrix_tab.is_copy_running()
        )
        customer_pi = (
            self._customer_pi_dialog is not None
            and self._customer_pi_dialog.is_busy()
        )
        return (
            self._scan_thread is not None
            or self._an_scan_thread is not None
            or self._rd_dump_scan_thread is not None
            or self._google_thread is not None
            or self._google_write_thread is not None
            or self._robot_sync_thread is not None
            or self._sq_to_rd_thread is not None
            or export_copy
            or customer_pi
        )

    def _startup_hydrate_busy(self) -> bool:
        """Return True while the first SQLite hydrate has not finished."""

        thread = getattr(self, "_startup_thread", None)
        return bool(
            getattr(self, "_startup_load_pending", False)
            or getattr(self, "_startup_load_running", False)
            or (thread is not None and thread.isRunning())
        )

    def _busy(self) -> bool:
        return self._catalog_workers_busy() or self._startup_hydrate_busy()

    def _set_workers_enabled(self, enabled: bool) -> None:
        for action in self._scan_actions:
            action.setEnabled(enabled)
        if hasattr(self, "_google_action"):
            self._google_action.setEnabled(enabled)
        if hasattr(self, "_customer_pi_action"):
            self._customer_pi_action.setEnabled(enabled)
        if hasattr(self, "_skip_prune_button"):
            self._skip_prune_button.setEnabled(enabled)
        if hasattr(self, "_an_tab"):
            self._an_tab.set_scan_enabled(enabled)
        if hasattr(self, "_rd_dump_tab"):
            self._rd_dump_tab.set_scan_enabled(enabled)
        if hasattr(self, "_kits_de_sync_button"):
            self._kits_de_sync_button.setEnabled(enabled)

    def _select_kit_row(self, title: str, mark: str) -> bool:
        """Select the комплекты row for ``title``/``mark`` after a refresh.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.

        Returns:
            True when a matching row was selected.
        """

        if not hasattr(self, "_kits_table"):
            return False
        needle = kit_identity_key(title, mark)
        table = self._kits_table
        for row_index in range(table.rowCount()):
            item = table.item(row_index, 0)
            row = item.data(_ROLE_ROW) if item else None
            if not isinstance(row, KitMatrixRow):
                continue
            if kit_identity_key(row.title, row.mark) != needle:
                continue
            table.selectRow(row_index)
            table.setCurrentCell(row_index, 0)
            table.scrollToItem(item)
            self._update_kits_panels()
            self._update_action_states()
            return True
        return False

    def _restore_scan_view(self) -> None:
        snapshot = self._scan_restore
        self._scan_restore = None
        if snapshot is None:
            return
        if 0 <= snapshot.tab_index < self._tabs.count():
            self._tabs.setCurrentIndex(snapshot.tab_index)
        if snapshot.stay_on_documents and snapshot.title:
            self._select_document_tree_node(
                snapshot.title, snapshot.mark, snapshot.transfer_name
            )
            return
        if snapshot.title and snapshot.mark:
            self._select_kit_row(snapshot.title, snapshot.mark)

    def start_scan(
        self,
        sources: tuple[SourceKind, ...],
        *,
        robot_subtree: str | Path | None = None,
        rd_subtree: str | Path | Sequence[str | Path] | None = None,
        sq_subtree: str | Path | None = None,
        keep_view: bool = False,
        restore_kit: tuple[str, str] | None = None,
        restore_tree: tuple[str, str, str | None] | None = None,
    ) -> None:
        """Start one source scan unless a worker is already active.

        Args:
            sources: Source subset to walk.
            robot_subtree: Optional robot folder for a scoped rescan.
            rd_subtree: Optional RD folder or folders for a scoped rescan.
            sq_subtree: Optional SQ folder for a scoped rescan.
            keep_view: Stay on the current tab instead of switching to Журнал.
            restore_kit: Title/mark to reselect on Комплекты.
            restore_tree: Title, mark, and transfer key to reselect in
                «Все документы».
        """

        if self._busy():
            QMessageBox.information(
                self,
                "Сканирование",
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return
        with perf_span(
            "gui.start_scan",
            sources=",".join(item.value for item in sources),
            rd=str(rd_subtree or ""),
            sq=str(sq_subtree or ""),
            robot=str(robot_subtree or ""),
        ):
            self._cancel_mto_compare(resume_later=True)
            self._cancel_export_pair_compare()
            if keep_view:
                kit = restore_kit
                if kit is None:
                    selected = self._selected_kit_row()
                    kit = (selected.title, selected.mark) if selected else ("", "")
                transfer_name: str | None = None
                stay_on_documents = False
                if restore_tree is not None:
                    kit = (restore_tree[0], restore_tree[1])
                    transfer_name = restore_tree[2]
                    stay_on_documents = True
                elif (
                    hasattr(self, "_documents_tab")
                    and self._tabs.currentWidget() is self._documents_tab
                ):
                    stay_on_documents = True
                self._scan_restore = _ScanViewRestore(
                    tab_index=self._tabs.currentIndex(),
                    title=kit[0],
                    mark=kit[1],
                    transfer_name=transfer_name,
                    stay_on_documents=stay_on_documents,
                )
            else:
                self._scan_restore = None
            thread = ScanThread(
                self.config,
                sources,
                self,
                robot_subtree=robot_subtree,
                rd_subtree=rd_subtree,
                sq_subtree=sq_subtree,
            )
            thread.progress.connect(self._on_scan_progress)
            thread.log.connect(self._append_log)
            thread.error.connect(self._on_scan_error)
            thread.finished.connect(self._on_scan_finished)
            self._scan_thread = thread
            self._set_workers_enabled(False)
            self._cancel_action.setEnabled(True)
            self._progress.setRange(0, 0)
            if not keep_view:
                self._tabs.setCurrentWidget(self._log_tab)
            elif robot_subtree:
                self.statusBar().showMessage("Обновляем папку MTO робота…")
            elif rd_subtree or sq_subtree:
                self.statusBar().showMessage("Обновляем папку комплекта…")
            self._update_action_states()
            thread.start()

    def _start_an_scan(self) -> None:
        """Start the AN dump child-process scan unless a catalog worker is busy."""

        if self._busy():
            return
        if _an_root_disabled(self.config.an_root):
            QMessageBox.information(
                self,
                "Скан АН",
                "Корень АН не задан (an_root пуст). Скан отключён.",
            )
            return
        self._cancel_mto_compare(resume_later=True)
        self._cancel_export_pair_compare()
        thread = AnScanThread(self.config, self)
        thread.progress.connect(self._on_an_scan_progress)
        thread.log.connect(self._append_log)
        thread.error.connect(self._on_an_scan_error)
        thread.finished.connect(self._on_an_scan_finished)
        self._an_scan_thread = thread
        self._set_workers_enabled(False)
        self._cancel_action.setEnabled(True)
        self._progress.setRange(0, 0)
        self._append_log("Скан АН…")
        self.statusBar().showMessage("Сканирование АН…")
        self._update_action_states()
        thread.start()

    def _start_rd_dump_scan(self) -> None:
        """Start the RD dump child-process scan unless a catalog worker is busy."""

        if self._busy():
            return
        if _rd_dump_root_disabled(self.config.rd_root):
            QMessageBox.information(
                self,
                "Скан РД",
                "Корень РД не задан (rd_root пуст). Скан отключён.",
            )
            return
        self._cancel_mto_compare(resume_later=True)
        self._cancel_export_pair_compare()
        thread = RdDumpScanThread(self.config, self)
        thread.progress.connect(self._on_rd_dump_scan_progress)
        thread.log.connect(self._append_log)
        thread.error.connect(self._on_rd_dump_scan_error)
        thread.finished.connect(self._on_rd_dump_scan_finished)
        self._rd_dump_scan_thread = thread
        self._set_workers_enabled(False)
        self._cancel_action.setEnabled(True)
        self._progress.setRange(0, 0)
        self._append_log("Скан РД (xlsx/doc)…")
        self.statusBar().showMessage("Сканирование РД (xlsx/doc)…")
        self._update_action_states()
        thread.start()

    def start_google_kits_load(self) -> None:
        """Fetch the Google kits sheet in a background thread."""

        if self._busy():
            QMessageBox.information(
                self,
                "Комплекты Google",
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return
        self._cancel_mto_compare(resume_later=True)
        self._cancel_export_pair_compare()
        thread = KitLoadThread(self.config, self)
        thread.log.connect(self._append_log)
        thread.error.connect(self._on_google_error)
        thread.finished.connect(self._on_google_finished)
        self._google_thread = thread
        self._set_workers_enabled(False)
        self._cancel_action.setEnabled(False)
        self._progress.setRange(0, 0)
        self._append_log("Загрузка комплектов Google…")
        self.statusBar().showMessage("Загрузка комплектов Google…")
        thread.start()

    @Slot()
    def cancel_scan(self) -> None:
        """Request cancellation of the active scan or MTO compare."""

        if self._scan_thread is not None:
            self._cancel_action.setEnabled(False)
            self._scan_thread.request_cancel()
        elif self._an_scan_thread is not None:
            self._cancel_action.setEnabled(False)
            self._an_scan_thread.request_cancel()
        elif self._rd_dump_scan_thread is not None:
            self._cancel_action.setEnabled(False)
            self._rd_dump_scan_thread.request_cancel()
        elif self._mto_compare_thread is not None:
            self._cancel_mto_compare(resume_later=False)

    def _cancel_export_pair_compare(self) -> None:
        """Cancel the heatmap export pair-compare child, if any."""

        if hasattr(self, "_revision_matrix_tab"):
            self._revision_matrix_tab.cancel_pair_compare()

    def _on_matrix_export_finished(self, folder: str, is_default_robot: bool) -> None:
        """Rescan the robot folder after a bulk copy into the default target."""

        if not is_default_robot or not folder:
            return
        if self._busy():
            return
        self.start_scan(
            (SourceKind.ROBOT,),
            robot_subtree=folder,
            keep_view=True,
        )

    def _cancel_mto_compare(self, *, resume_later: bool = False) -> None:
        """Request cooperative cancellation of a running MTO compare thread.

        Args:
            resume_later: When True, leftover keys are restarted after the
                interrupting worker finishes. User «Отмена» passes False.
        """

        if resume_later:
            self._mto_resume_on_idle = True
        else:
            self._mto_resume_on_idle = False
        thread = self._mto_compare_thread
        if thread is None:
            return
        thread.request_cancel()
        self._cancel_action.setEnabled(False)

    def _launch_mto_compare(
        self,
        document_keys: frozenset[tuple[str, str]] | None,
        priority_keys: tuple[tuple[str, str], ...],
        *,
        planned: int,
    ) -> None:
        """Start ``MtoCompareThread`` and wire status-bar progress only."""

        thread = MtoCompareThread(
            self.config,
            document_keys=document_keys,
            priority_keys=priority_keys,
            batch_size=1,
            parent=self,
        )
        thread.log.connect(self._append_log)
        thread.error.connect(self._on_mto_compare_error)
        thread.progress.connect(self._on_mto_compare_progress)
        thread.finished.connect(self._on_mto_compare_finished)
        self._mto_compare_thread = thread
        self._mto_progress = (0, planned)
        if self._scan_thread is None:
            self._cancel_action.setEnabled(True)
        self.statusBar().showMessage(f"MTO сверка: 0/{planned}")
        self._update_mto_sync_label()
        self._update_action_states()
        thread.start()

    def _start_mto_compare_after_scan(
        self,
        scope: str,
        touched_keys: tuple[tuple[str, str], ...] | list[tuple[str, str]],
    ) -> None:
        """Enqueue or start deferred MTO compare after a scan refresh.

        Args:
            scope: ``none``, ``subtree``, or ``full`` from the scan worker.
            touched_keys: MTO document keys touched during the walk.
        """

        if scope == "none":
            self._resume_pending_mto_compare()
            return
        priority = tuple(touched_keys)
        self._mto_pending_keys |= set(touched_keys)
        if scope == "subtree":
            document_keys: frozenset[tuple[str, str]] | None = frozenset(
                self._mto_pending_keys
            )
        else:
            document_keys = None
        self._enqueue_mto_compare(document_keys, priority)

    def _resume_pending_mto_compare(self) -> bool:
        """Restart leftover MTO work after Google or a cancelled job.

        Returns:
            True when a compare thread was started.
        """

        if self._busy() or self._mto_compare_thread is not None:
            return False
        if not self._mto_pending_keys:
            return False
        keys = frozenset(self._mto_pending_keys)
        return self._enqueue_mto_compare(keys, tuple(keys))

    def _enqueue_mto_compare(
        self,
        document_keys: frozenset[tuple[str, str]] | None,
        priority_keys: tuple[tuple[str, str], ...],
    ) -> bool:
        """Start or queue the compare thread; the child job builds the plan.

        Args:
            document_keys: Subtree/pending keys, or ``None`` for full incremental.
            priority_keys: Keys to compare first.

        Returns:
            True when a thread is running or queued.
        """

        if document_keys is not None and not document_keys and not priority_keys:
            self._update_mto_sync_label()
            return False
        if self._mto_compare_thread is not None:
            self._mto_queued = (document_keys, priority_keys)
            return True
        planned = (
            max(len(priority_keys), 1)
            if document_keys is None or document_keys or priority_keys
            else 0
        )
        self._launch_mto_compare(
            document_keys,
            priority_keys,
            planned=planned,
        )
        return True

    @Slot(int, int)
    def _on_mto_compare_progress(self, completed: int, total: int) -> None:
        self._mto_progress = (completed, total)
        self.statusBar().showMessage(f"MTO сверка: {completed}/{total}")
        self._update_mto_sync_label()

    @Slot(str)
    def _on_mto_compare_error(self, message: str) -> None:
        self._append_log(message)

    @Slot()
    def _on_mto_compare_finished(self) -> None:
        thread = self._mto_compare_thread
        if thread is None:
            return
        self._mto_compare_thread = None
        thread.deleteLater()
        scope_keys = thread._document_keys
        if thread.cancelled or thread.remaining_keys:
            self._mto_pending_keys |= set(thread.remaining_keys)
        elif scope_keys is None:
            self._mto_pending_keys.clear()
        else:
            self._mto_pending_keys -= set(scope_keys)
        if self._busy():
            self._update_mto_sync_label()
            self._update_action_states()
            return
        queued = self._mto_queued
        if queued is not None:
            self._mto_queued = None
            document_keys, priority_keys = queued
            if self._enqueue_mto_compare(document_keys, priority_keys):
                return
        elif self._mto_resume_on_idle:
            self._mto_resume_on_idle = False
            if self._resume_pending_mto_compare():
                return
        self._refresh_mto_related()
        if thread.failure:
            self.statusBar().showMessage(f"MTO сверка: {thread.failure}", 10_000)
        else:
            self.statusBar().showMessage("MTO сверка завершена", 10_000)
        self._update_mto_sync_label()
        self._update_action_states()

    @Slot(object)
    def _on_scan_progress(self, progress: ScanProgress) -> None:
        message = progress.message or progress.path
        self.statusBar().showMessage(
            f"{progress.source.value.upper()}: {progress.files_seen} · {message}"
        )

    @Slot(object)
    def _on_an_scan_progress(self, progress: AnScanProgress) -> None:
        message = progress.message or progress.path
        self.statusBar().showMessage(f"АН: {progress.files_seen} · {message}")

    @Slot(str)
    def _on_an_scan_error(self, message: str) -> None:
        self._append_log(message)
        self.statusBar().showMessage("Ошибка скана АН")

    @Slot()
    def _on_an_scan_finished(self) -> None:
        thread = self._an_scan_thread
        if thread is None:
            return
        with perf_span("gui.an_scan_finished"):
            failure = thread.failure
            accepted = thread.accepted
            files_seen = thread.files_seen
            cancelled = thread.cancelled
            self._set_workers_enabled(True)
            self._cancel_action.setEnabled(False)
            self._progress.setRange(0, 1)
            self._progress.setValue(1)
            self._an_scan_thread = None
            thread.deleteLater()
            if failure:
                self._append_log(f"АН: {failure}")
            elif cancelled:
                self._append_log("АН: скан отменён")
            else:
                self._append_log(f"АН: принято {accepted} из {files_seen}")
            self._reload_an_index()
            self._refresh_kits_table(rebuild_rows=False)
            self._refresh_an_tab()
            resumed = self._resume_pending_mto_compare()
            self._mto_resume_on_idle = (
                not resumed
                and bool(self._mto_pending_keys)
                and (self._mto_compare_thread is not None or self._busy())
            )
            if not resumed:
                if failure:
                    status = f"АН: {failure}"
                elif cancelled:
                    status = "Скан АН отменён"
                else:
                    status = f"АН: принято {accepted}"
                self.statusBar().showMessage(status, 10_000)
            self._update_action_states()

    @Slot(object)
    def _on_rd_dump_scan_progress(self, progress: RdDumpScanProgress) -> None:
        message = progress.message or progress.path
        self.statusBar().showMessage(f"РД: {progress.files_seen} · {message}")

    @Slot(str)
    def _on_rd_dump_scan_error(self, message: str) -> None:
        self._append_log(message)
        self.statusBar().showMessage("Ошибка скана РД")

    @Slot()
    def _on_rd_dump_scan_finished(self) -> None:
        thread = self._rd_dump_scan_thread
        if thread is None:
            return
        with perf_span("gui.rd_dump_scan_finished"):
            failure = thread.failure
            accepted = thread.accepted
            files_seen = thread.files_seen
            cancelled = thread.cancelled
            self._set_workers_enabled(True)
            self._cancel_action.setEnabled(False)
            self._progress.setRange(0, 1)
            self._progress.setValue(1)
            self._rd_dump_scan_thread = None
            thread.deleteLater()
            if failure:
                self._append_log(f"РД: {failure}")
            elif cancelled:
                self._append_log("РД: скан отменён")
            else:
                self._append_log(f"РД: принято {accepted} из {files_seen}")
            self._reload_rd_dump_index()
            self._refresh_rd_dump_tab()
            resumed = self._resume_pending_mto_compare()
            self._mto_resume_on_idle = (
                not resumed
                and bool(self._mto_pending_keys)
                and (self._mto_compare_thread is not None or self._busy())
            )
            if not resumed:
                if failure:
                    status = f"РД: {failure}"
                elif cancelled:
                    status = "Скан РД отменён"
                else:
                    status = f"РД: принято {accepted}"
                self.statusBar().showMessage(status, 10_000)
            self._update_action_states()

    @Slot(str)
    def _append_log(self, message: str) -> None:
        text = stamp_log_line(message)
        if text:
            self._log.appendPlainText(text)

    @Slot(str)
    def _on_scan_error(self, message: str) -> None:
        self._append_log(message)
        self.statusBar().showMessage("Ошибка сканирования")

    @Slot()
    def _on_scan_finished(self) -> None:
        thread = self._scan_thread
        if thread is None:
            return
        with perf_span(
            "gui.scan_finished",
            rd=getattr(thread, "_rd_subtree", "") or "",
            sq=getattr(thread, "_sq_subtree", "") or "",
            robot=getattr(thread, "_robot_subtree", "") or "",
        ):
            failure = thread.failure
            compare_scope = thread.compare_enqueue_scope
            touched_mto_keys = tuple(thread.touched_mto_keys)
            robot_only_subtree = bool(
                thread._robot_subtree and not thread._rd_subtree and not thread._sq_subtree
            )
            kit_subtree = bool(thread._rd_subtree or thread._sq_subtree)
            rd_paths = tuple(getattr(thread, "_rd_subtrees", ()) or ())
            if not rd_paths and thread._rd_subtree:
                rd_paths = (thread._rd_subtree,)
            pipeline_subtrees = tuple(
                path
                for path in (
                    *rd_paths,
                    thread._sq_subtree,
                    thread._robot_subtree,
                )
                if path
            )
            self._set_workers_enabled(True)
            self._cancel_action.setEnabled(False)
            self._progress.setRange(0, 1)
            self._progress.setValue(1)
            self._scan_thread = None
            thread.deleteLater()
            restore = self._scan_restore
            stay_on_documents = bool(restore is not None and restore.stay_on_documents)
            defer_secondary = restore is not None
            rebuild_document_tree = restore is None or stay_on_documents
            self.statusBar().showMessage("Обновляем комплекты и таблицы…")
            QApplication.processEvents()
            self.refresh(
                rebuild_document_tree=rebuild_document_tree,
                ingest_google=False,
                pipeline_subtrees=pipeline_subtrees or None,
                defer_secondary=defer_secondary,
            )
            if defer_secondary:
                self._schedule_auto_mto_compares()
            if restore is not None:
                self._restore_scan_view()
                if robot_only_subtree:
                    prefix = "Папка робота обновлена"
                elif kit_subtree:
                    prefix = "Папка комплекта обновлена"
                else:
                    prefix = "Сканирование завершено"
                self.statusBar().showMessage(
                    f"{prefix}{': ' + failure if failure else ''}",
                    10_000,
                )
            else:
                self.statusBar().showMessage(
                    f"Сканирование завершено{': ' + failure if failure else ''}",
                    10_000,
                )
            self._start_mto_compare_after_scan(compare_scope, touched_mto_keys)
            self._update_action_states()

    @Slot(str)
    def _on_google_error(self, message: str) -> None:
        self._append_log(message)
        self.statusBar().showMessage("Ошибка загрузки комплектов Google")

    @Slot()
    def _on_google_finished(self) -> None:
        thread = self._google_thread
        if thread is None:
            return
        with perf_span("gui.google_finished"):
            result = thread.result
            failure = thread.failure
            self._set_workers_enabled(True)
            self._progress.setRange(0, 1)
            self._progress.setValue(1)
            self._google_thread = None
            thread.deleteLater()
            if result is not None and not result.error:
                self._google_kits = result.kits
                self._issuance_kits = result.issuance_kits
                self._issuance_sends = result.issuance_sends
                self._google_source = result.source
                self._google_fetched_at = result.fetched_at
                self._google_warning = result.warning
                self._google_loaded = True
                try:
                    ingest_google_snapshot(
                        self.database,
                        result.kits,
                        result.issuance_sends,
                        loaded_at=result.fetched_at
                        or datetime.now(timezone.utc).isoformat(),
                        source=result.source,
                        warning=result.warning,
                    )
                    self._issuance_kits = latest_effective_issuance_kits(
                        self.database
                    )
                    self._issuance_kits_fresh = True
                except Exception as exc:
                    self._append_log(f"Google ingest: {type(exc).__name__}: {exc}")
                if result.stats.skipped_reasons:
                    preview = "; ".join(result.stats.skipped_reasons[:8])
                    self._append_log(f"Google · пропущены: {preview}")
                if result.issuance_stats.skipped_reasons:
                    preview = "; ".join(result.issuance_stats.skipped_reasons[:8])
                    self._append_log(f"Выдача РД ПД · пропущены: {preview}")
                self._append_log(
                    f"Google: {result.stats.kept}; Выдача РД ПД: {result.issuance_stats.kept} "
                    f"({result.source})"
                )
            self.refresh(ingest_google=False)
            resumed = self._resume_pending_mto_compare()
            self._mto_resume_on_idle = (
                not resumed
                and bool(self._mto_pending_keys)
                and (self._mto_compare_thread is not None or self._busy())
            )
            if not resumed:
                self.statusBar().showMessage(
                    f"Комплекты Google: {failure}"
                    if failure
                    else "Комплекты Google загружены",
                    10_000,
                )
            self._update_action_states()

    def _confirm_robot_mto_sync(self, row: KitMatrixRow) -> None:
        """Confirm copying the newest RD/SQ MTO into the robot folder."""

        if self._busy():
            QMessageBox.information(
                self,
                "MTO робота",
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return
        with perf_span("gui.plan_robot_mto", title=row.title, mark=row.mark):
            try:
                plan = plan_robot_mto_sync(
                    title=row.title,
                    mark=row.mark,
                    records=self._all_records,
                    detected_current_ids=self._official_current_ids(),
                    robot_root=self.config.robot_root,
                    robot_flat_structure=self.config.robot_flat_structure,
                    rd_folders=self._official_rd_folders(row),
                    sq_folders=tuple(str(Path(path).parent) for path in row.sq.paths),
                )
            except RobotMtoSyncError as exc:
                QMessageBox.warning(self, "MTO робота", str(exc))
                return
        dialog = RobotMtoSyncDialog(
            plan,
            self,
            runtime_dir=self.config.runtime_dir,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        self._start_robot_mto_sync(plan)

    def _start_robot_mto_sync(self, plan: RobotMtoSyncPlan) -> None:
        self._cancel_mto_compare(resume_later=True)
        self._cancel_export_pair_compare()
        thread = RobotMtoSyncThread(plan, self, now=datetime.now())
        thread.log.connect(self._append_log)
        thread.error.connect(self._on_robot_sync_error)
        thread.finished.connect(self._on_robot_sync_finished)
        self._robot_sync_thread = thread
        self._set_workers_enabled(False)
        self._cancel_action.setEnabled(False)
        self._progress.setRange(0, 0)
        self.statusBar().showMessage("Копируем MTO в папку робота…")
        self._update_action_states()
        thread.start()

    @Slot(str)
    def _on_robot_sync_error(self, message: str) -> None:
        self._append_log(message)
        self.statusBar().showMessage("Ошибка обновления MTO робота")

    @Slot()
    def _on_robot_sync_finished(self) -> None:
        thread = self._robot_sync_thread
        if thread is None:
            return
        with perf_span("gui.robot_sync_finished"):
            result = thread.result
            failure = thread.failure
            plan = thread.plan
            self._set_workers_enabled(True)
            self._progress.setRange(0, 1)
            self._progress.setValue(1)
            self._robot_sync_thread = None
            thread.deleteLater()
            self._update_action_states()
            if failure or result is None:
                self.statusBar().showMessage("Ошибка обновления MTO робота", 10_000)
                QMessageBox.warning(
                    self,
                    "MTO робота",
                    failure or "Не удалось обновить файл робота.",
                )
                return
            action = "добавлен" if result.added else "заменён"
            folder = str(Path(result.destination_path).parent)
            self.statusBar().showMessage(f"MTO робота {action}. Обновляем папку…", 10_000)
            self.start_scan(
                (SourceKind.ROBOT,),
                robot_subtree=folder,
                keep_view=True,
                restore_kit=(plan.title, plan.mark),
            )

    def _kit_source_paths(self, title: str, mark: str, source: SourceKind) -> tuple[str, ...]:
        """Return present catalog paths for one title+mark source.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
            source: RD, SQ, or ROBOT.

        Returns:
            File paths recorded as present for this kit.
        """

        key = kit_identity_key(title, mark)
        paths: list[str] = []
        for record in self._all_records:
            if record.source is not source or not record.present:
                continue
            rec_title = str(record.data.get("title") or "").strip()
            rec_mark = str(record.data.get("mark") or "").strip()
            if kit_identity_key(rec_title, rec_mark) != key:
                continue
            paths.append(record.path)
        return tuple(paths)

    def _confirm_sq_to_rd(self, row: KitMatrixRow) -> None:
        """Confirm moving the SQ kit folder into a new RD transfer."""

        if self._busy():
            QMessageBox.information(
                self,
                "Перенос SQ в РД",
                "Дождитесь завершения текущей загрузки или сканирования.",
            )
            return
        if not row.sq.present or not row.sq.paths:
            QMessageBox.information(
                self,
                "Перенос SQ в РД",
                "У этой строки нет файлов SQ.",
            )
            return
        try:
            plan = plan_sq_to_rd_transfer(
                title=row.title,
                mark=row.mark,
                records=self._all_records,
                sq_paths=self._kit_source_paths(
                    row.title, row.mark, SourceKind.SQ
                )
                or row.sq.paths,
                rd_paths=self._kit_source_paths(
                    row.title, row.mark, SourceKind.RD
                )
                or row.rd.paths,
                sq_revision_text=row.sq.revision_text,
                sq_max_mtime_ns=row.sq.max_mtime_ns,
                rd_root=self.config.rd_root,
                sq_root=self.config.sq_root,
            )
        except SqToRdError as exc:
            QMessageBox.warning(self, "Перенос SQ в РД", str(exc))
            return
        dialog = SqToRdDialog(plan, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        self._start_sq_to_rd(plan)

    def _start_sq_to_rd(self, plan: SqToRdPlan) -> None:
        self._cancel_mto_compare(resume_later=True)
        self._cancel_export_pair_compare()
        thread = SqToRdThread(plan, self)
        thread.log.connect(self._append_log)
        thread.error.connect(self._on_sq_to_rd_error)
        thread.finished.connect(self._on_sq_to_rd_finished)
        self._sq_to_rd_thread = thread
        self._set_workers_enabled(False)
        self._cancel_action.setEnabled(False)
        self._progress.setRange(0, 0)
        self.statusBar().showMessage("Переносим папку SQ в РД…")
        self._update_action_states()
        thread.start()

    @Slot(str)
    def _on_sq_to_rd_error(self, message: str) -> None:
        self._append_log(message)
        self.statusBar().showMessage("Ошибка переноса SQ в РД")

    @Slot()
    def _on_sq_to_rd_finished(self) -> None:
        thread = self._sq_to_rd_thread
        if thread is None:
            return
        with perf_span("gui.sq_to_rd_finished"):
            result = thread.result
            failure = thread.failure
            plan = thread.plan
            self._set_workers_enabled(True)
            self._progress.setRange(0, 1)
            self._progress.setValue(1)
            self._sq_to_rd_thread = None
            thread.deleteLater()
            self._update_action_states()
            if failure or result is None:
                self.statusBar().showMessage("Ошибка переноса SQ в РД", 10_000)
                QMessageBox.warning(
                    self,
                    "Перенос SQ в РД",
                    failure or "Не удалось переместить папку SQ.",
                )
                return
            self.statusBar().showMessage(
                "SQ перенесена в РД. Обновляем папку комплекта…", 10_000
            )
            self.start_scan(
                (SourceKind.RD, SourceKind.SQ),
                rd_subtree=plan.gate_folder,
                sq_subtree=result.source_folder,
                keep_view=True,
                restore_kit=(plan.title, plan.mark),
            )

    def _open_result(self, path: str | None, *, folder: bool) -> None:
        ok, message = (
            open_containing_folder(path) if folder else open_path(path)
        )
        if not ok:
            self._append_log(message)
            self.statusBar().showMessage(message, 10_000)

    def _open_mto_side(self, side: str, *, folder: bool) -> None:
        row = self._selected_mto_row()
        if not row:
            return
        path = row.get(f"{side}_path")
        if not path or not row.get(f"{side}_present"):
            self._append_log(f"Файл {side.upper()} отсутствует: {path or 'путь не указан'}")
            return
        self._open_result(str(path), folder=folder)

    @Slot(int, int)
    def _open_mto_cell(self, _row: int, column: int) -> None:
        robot_columns = {5, 10, 12}
        side = "robot" if column in robot_columns else "rd"
        self._open_mto_side(side, folder=False)

    def _open_collision_folder(self) -> None:
        row = self._selected_collision_row()
        paths = list(row.get("paths") or ()) if row else []
        if len(paths) != 1:
            return
        self._open_result(str(paths[0]), folder=True)

    def _copy_collision_paths(self) -> None:
        row = self._selected_collision_row()
        paths = [str(path) for path in (row.get("paths") or ())] if row else []
        self._copy_text("\n".join(paths))

    def _open_history(self, *, folder: bool) -> None:
        record = self._selected_record()
        if not record:
            return
        if not record.present:
            self._append_log(f"Файл отсутствует по последнему скану: {record.path}")
            return
        self._open_result(record.path, folder=folder)

    def _copy_history_path(self) -> None:
        bundle = self._selected_bundle()
        if bundle is None:
            return
        self._copy_text("\n".join(_bundle_paths(bundle)), message="Полный путь скопирован")

    def _copy_history_name(self) -> None:
        bundle = self._selected_bundle()
        if bundle is None:
            return
        record = _bundle_open_record(bundle)
        name = _record_name(record) if record is not None else ""
        self._copy_text(name, message="Имя скопировано")

    def _copy_history_package_path(self) -> None:
        bundle = self._selected_bundle()
        if bundle is None:
            return
        package = _bundle_package_path(bundle)
        self._copy_text(package, message="Путь комплекта скопирован")

    def _copy_history_column(self, column: int) -> None:
        """Copy the visible text of one history-table column for the current row."""

        if column == _HISTORY_COL_NAME:
            self._copy_history_name()
        elif column == _HISTORY_COL_DATE:
            item = self._history.item(self._history.currentRow(), _HISTORY_COL_DATE)
            self._copy_text(
                item.text() if item is not None else "",
                message="Дата скопирована",
            )
        elif column == _HISTORY_COL_PACKAGE:
            self._copy_history_package_path()
        else:
            self._copy_history_path()

    def _open_outlook_od_search(self, query: str) -> None:
        """Open Outlook Instant Search for the kit OD token, all mailboxes."""

        try:
            open_outlook_instant_search(query)
        except RuntimeError as exc:
            self._copy_text(
                query,
                message=f"Скопировано для Outlook: {query}",
            )
            QMessageBox.warning(self, "Найти в Outlook", str(exc))
            return
        self.statusBar().showMessage(f"Outlook: поиск {query}", 4000)

    def _copy_text(self, text: str | None, *, message: str = "Скопировано") -> None:
        if text:
            QApplication.clipboard().setText(text)
            self.statusBar().showMessage(message, 3000)

    def _show_mto_context_menu(self, position) -> None:
        row_index = self._mto_table.rowAt(position.y())
        if row_index < 0:
            return
        self._mto_table.selectRow(row_index)
        row = self._selected_mto_row()
        if not row:
            return
        menu = QMenu(self)
        actions: dict[QAction, tuple[str, bool] | tuple[str, str]] = {}
        for side, label in (("rd", "MTO РД"), ("robot", "MTO робота")):
            path = str(row.get(f"{side}_path") or "")
            present = bool(row.get(f"{side}_present"))
            open_action = menu.addAction(f"Открыть файл · {label}")
            folder_action = menu.addAction(f"Открыть содержащую папку · {label}")
            copy_action = menu.addAction(f"Копировать полный путь · {label}")
            open_action.setEnabled(bool(path and present))
            folder_action.setEnabled(bool(path and present))
            copy_action.setEnabled(bool(path))
            actions[open_action] = (side, False)
            actions[folder_action] = (side, True)
            actions[copy_action] = ("copy", path)
            if side == "rd":
                menu.addSeparator()
        menu.addSeparator()
        ban_action = menu.addAction("Скрыть титул–марку (бан-фильтр)")
        handoff_action = self._add_robot_handoff_action(menu)
        chosen = exec_tracked_menu(
            menu,
            MENU_MTO_READINESS,
            self._mto_table.viewport().mapToGlobal(position),
        )
        if chosen == ban_action:
            self._confirm_ban_title_mark(
                str(row.get("title") or ""),
                str(row.get("mark") or ""),
            )
            return
        if chosen == handoff_action:
            column = self._mto_table.columnAt(position.x())
            self._copy_robot_handoff(self._mto_row_handoff_text(row_index, column))
            return
        command = actions.get(chosen)
        if not command:
            return
        if command[0] == "copy":
            self._copy_text(command[1])
        else:
            self._open_mto_side(command[0], folder=bool(command[1]))

    def _show_history_context_menu(self, position) -> None:
        row_index = self._history.rowAt(position.y())
        if row_index < 0:
            return
        self._history.selectRow(row_index)
        bundle = self._selected_bundle()
        if bundle is None:
            return
        display_record = _bundle_open_record(bundle)
        display_name = (
            _record_name(display_record) if display_record is not None else ""
        )
        menu = QMenu(self)
        pdf = bundle.pdf
        pdf_available = bool(pdf is not None and pdf.present and pdf.path)
        open_pdf = menu.addAction("Открыть PDF")
        folder_pdf = menu.addAction("Открыть папку PDF")
        open_pdf.setEnabled(pdf_available)
        folder_pdf.setEnabled(pdf_available)
        editable_actions: dict[QAction, tuple[FileRecord, bool]] = {}
        for editable in bundle.editables:
            ext = _file_extension(editable) or _record_name(editable)
            open_editable = menu.addAction(f"Открыть ред · {ext}")
            folder_editable = menu.addAction(f"Папка ред · {ext}")
            available = bool(editable.present and editable.path)
            open_editable.setEnabled(available)
            folder_editable.setEnabled(available)
            editable_actions[open_editable] = (editable, False)
            editable_actions[folder_editable] = (editable, True)
        copy_name = menu.addAction("Копировать имя")
        copy_package = menu.addAction("Копировать путь комплекта")
        copy_action = menu.addAction("Копировать полные пути файлов")
        copy_name.setEnabled(bool(display_name))
        copy_package.setEnabled(bool(_bundle_package_path(bundle)))
        copy_action.setEnabled(bool(_bundle_paths(bundle)))
        handoff_action = self._add_robot_handoff_action(menu)
        mto_record = _bundle_mto_xlsx(bundle)
        folder_files = self._history_folder_records()
        date_menu = self._add_mtime_override_menu(
            menu,
            mto_record,
            folder_files=folder_files,
            hide_when_missing=True,
        )
        chosen = exec_tracked_menu(
            menu, MENU_HISTORY, self._history.viewport().mapToGlobal(position)
        )
        if chosen == open_pdf and pdf is not None:
            self._open_result(pdf.path, folder=False)
        elif chosen == folder_pdf and pdf is not None:
            self._open_result(pdf.path, folder=True)
        elif chosen == copy_name:
            self._copy_history_name()
        elif chosen == copy_package:
            self._copy_history_package_path()
        elif chosen == copy_action:
            self._copy_history_path()
        elif chosen == handoff_action:
            column = self._history.columnAt(position.x())
            self._copy_robot_handoff(self._history_handoff_text(row_index, column))
        elif date_menu.code is not None and chosen == date_menu.code:
            if mto_record is not None:
                self._confirm_mto_catalog_date(mto_record)
        elif date_menu.manual is not None and chosen == date_menu.manual:
            if mto_record is not None:
                self._confirm_mto_catalog_date_manual(
                    mto_record, folder_files=folder_files
                )
        elif date_menu.folder is not None and chosen == date_menu.folder:
            if mto_record is not None:
                self._confirm_mto_catalog_date_folder_mean(
                    mto_record, folder_files=folder_files
                )
        elif date_menu.clear is not None and chosen == date_menu.clear:
            if mto_record is not None:
                self._clear_mto_catalog_date(mto_record, folder_files=folder_files)
        else:
            command = editable_actions.get(chosen) if chosen is not None else None
            if command:
                record, folder = command
                self._open_result(record.path, folder=folder)

    def _show_collision_context_menu(self, position) -> None:
        row_index = self._collision_table.rowAt(position.y())
        if row_index < 0:
            return
        self._collision_table.selectRow(row_index)
        row = self._selected_collision_row()
        if not row:
            return
        paths = [str(path) for path in row.get("paths") or ()]
        menu = QMenu(self)
        folder_action = menu.addAction("Открыть содержащую папку")
        copy_action = menu.addAction("Копировать пути")
        folder_action.setEnabled(len(paths) == 1)
        copy_action.setEnabled(bool(paths))
        handoff_action = self._add_robot_handoff_action(menu)
        chosen = exec_tracked_menu(
            menu,
            MENU_COLLISION,
            self._collision_table.viewport().mapToGlobal(position),
        )
        if chosen == folder_action:
            self._open_collision_folder()
        elif chosen == copy_action:
            self._copy_collision_paths()
        elif chosen == handoff_action:
            column = self._collision_table.columnAt(position.x())
            self._copy_robot_handoff(
                self._collision_row_handoff_text(row_index, column)
            )

    def _review_selected(self, action: ReviewState | str) -> None:
        record = self._selected_record()
        if record is None or self._busy():
            return
        action_value = action.value if isinstance(action, ReviewState) else action
        comment: str | None = None
        if action_value in ("comment", ReviewState.IGNORED.value):
            title = (
                "Игнорировать файл"
                if action_value == ReviewState.IGNORED.value
                else "Комментарий"
            )
            comment, accepted = QInputDialog.getMultiLineText(
                self, title, "Комментарий:", ""
            )
            if not accepted:
                return
            comment = comment.strip()
            if action_value == ReviewState.IGNORED.value and not comment:
                QMessageBox.warning(
                    self, title, "Для игнорирования обязателен комментарий."
                )
                return
        try:
            self.database.record_review(record.id, action, comment=comment)
        except (ValueError, KeyError) as exc:
            QMessageBox.warning(self, "Проверка", str(exc))
            return
        self._append_log(
            f"Review: file_id={record.id}, action={action_value}, "
            f"comment={json.dumps(comment, ensure_ascii=False)}"
        )
        self.refresh(ingest_google=False)

    @staticmethod
    def _normalize_kits_detail_placement(value: object) -> str:
        text = str(value or "").strip().casefold()
        if text in {_KITS_DETAIL_PLACEMENT_RIGHT, "side"}:
            return _KITS_DETAIL_PLACEMENT_RIGHT
        return _KITS_DETAIL_PLACEMENT_BOTTOM

    @staticmethod
    def _kits_splitter_key_for(placement: str) -> str:
        if placement == _KITS_DETAIL_PLACEMENT_RIGHT:
            return _KITS_SPLITTER_RIGHT_KEY
        return _KITS_SPLITTER_BOTTOM_KEY

    def _update_kits_detail_placement_button(self) -> None:
        if not hasattr(self, "_kits_detail_placement_button"):
            return
        bottom = (
            getattr(
                self,
                "_kits_detail_placement_value",
                _KITS_DETAIL_PLACEMENT_BOTTOM,
            )
            == _KITS_DETAIL_PLACEMENT_BOTTOM
        )
        if bottom:
            self._kits_detail_placement_button.setText("→")
            self._kits_detail_placement_button.setToolTip("Панель справа")
        else:
            self._kits_detail_placement_button.setText("↓")
            self._kits_detail_placement_button.setToolTip("Панель снизу")

    def _apply_kits_detail_placement(
        self, placement: str, *, restore_sizes: bool = True
    ) -> None:
        """Put the kits detail pane under the table or to the right of it.

        Args:
            placement: ``bottom`` or ``right``.
            restore_sizes: Restore the matching QSettings splitter state when True.
        """

        if not hasattr(self, "_kits_splitter"):
            return
        placement = self._normalize_kits_detail_placement(placement)
        self._kits_detail_placement_value = placement
        bottom = placement == _KITS_DETAIL_PLACEMENT_BOTTOM
        orientation = (
            Qt.Orientation.Vertical if bottom else Qt.Orientation.Horizontal
        )
        self._kits_splitter.setOrientation(orientation)
        if bottom:
            self._kits_detail_tabs.setMinimumWidth(0)
            self._kits_detail_tabs.setMinimumHeight(80)
            self._kits_detail_tabs.setSizePolicy(
                QSizePolicy.Policy.Expanding,
                QSizePolicy.Policy.Preferred,
            )
            stretch_table, stretch_detail = 3, 1
            default_sizes = [720, 220]
        else:
            self._kits_detail_tabs.setMinimumHeight(0)
            self._kits_detail_tabs.setMinimumWidth(140)
            self._kits_detail_tabs.setSizePolicy(
                QSizePolicy.Policy.Ignored,
                QSizePolicy.Policy.Expanding,
            )
            stretch_table, stretch_detail = 1, 0
            default_sizes = [1200, 180]
        restored = False
        if restore_sizes:
            state = self._settings.value(self._kits_splitter_key_for(placement))
            if isinstance(state, QByteArray) and not state.isEmpty():
                restored = bool(self._kits_splitter.restoreState(state))
        self._kits_splitter.setOrientation(orientation)
        self._kits_splitter.setChildrenCollapsible(False)
        self._kits_splitter.setStretchFactor(0, stretch_table)
        self._kits_splitter.setStretchFactor(1, stretch_detail)
        if not restored:
            self._kits_splitter.setSizes(default_sizes)
        self._update_kits_detail_placement_button()

    def _persist_kits_detail_layout(self) -> None:
        if not hasattr(self, "_kits_splitter"):
            return
        placement = self._normalize_kits_detail_placement(
            getattr(
                self,
                "_kits_detail_placement_value",
                _KITS_DETAIL_PLACEMENT_BOTTOM,
            )
        )
        self._settings.setValue(_KITS_DETAIL_PLACEMENT_KEY, placement)
        self._settings.setValue(
            self._kits_splitter_key_for(placement),
            self._kits_splitter.saveState(),
        )

    def _toggle_kits_detail_placement(self) -> None:
        current = self._normalize_kits_detail_placement(
            getattr(
                self,
                "_kits_detail_placement_value",
                _KITS_DETAIL_PLACEMENT_BOTTOM,
            )
        )
        self._persist_kits_detail_layout()
        nxt = (
            _KITS_DETAIL_PLACEMENT_RIGHT
            if current == _KITS_DETAIL_PLACEMENT_BOTTOM
            else _KITS_DETAIL_PLACEMENT_BOTTOM
        )
        self._apply_kits_detail_placement(nxt)
        self._persist_kits_detail_layout()
        self._settings.sync()

    def _restore_layout(self) -> None:
        geometry = self._settings.value("window/geometry")
        if isinstance(geometry, QByteArray) and not geometry.isEmpty():
            self.restoreGeometry(geometry)
        splitter = self._settings.value("window/doc_splitter")
        if isinstance(splitter, QByteArray) and not splitter.isEmpty():
            self._doc_splitter.restoreState(splitter)
        if hasattr(self, "_kits_splitter"):
            placement = self._normalize_kits_detail_placement(
                self._settings.value(
                    _KITS_DETAIL_PLACEMENT_KEY,
                    _KITS_DETAIL_PLACEMENT_BOTTOM,
                )
            )
            self._apply_kits_detail_placement(placement)
        tab_index = self._settings.value(_KITS_DETAIL_TAB_KEY)
        if hasattr(self, "_kits_detail_tabs") and tab_index is not None:
            try:
                index = int(tab_index)
            except (TypeError, ValueError):
                index = 0
            if 0 <= index < self._kits_detail_tabs.count():
                self._kits_detail_tabs.setCurrentIndex(index)
        self._restore_table_headers()
        self._restore_tree_label_options()
        if hasattr(self, "_revision_matrix_tab"):
            self._revision_matrix_tab.restore_filters(self._settings)
        if hasattr(self, "_mto_worklist_tab"):
            self._mto_worklist_tab.restore_filters(self._settings)
        if hasattr(self, "_issuance_journal_tab"):
            self._issuance_journal_tab.restore_filters(self._settings)
        if hasattr(self, "_an_tab"):
            self._an_tab.restore_filters(self._settings)
        if hasattr(self, "_rd_dump_tab"):
            self._rd_dump_tab.restore_filters(self._settings)
        if hasattr(self, "_approval_mail_tab"):
            self._approval_mail_tab.restore_settings(self._settings)
        if hasattr(self, "_kits_filter"):
            self._restore_kits_filters()

    def _restore_web_autostart(self) -> None:
        """Load the WEB autostart checkbox; default is off."""

        if not hasattr(self, "_web_autostart_cb"):
            return
        self._web_autostart_cb.blockSignals(True)
        self._web_autostart_cb.setChecked(
            _settings_bool(self._settings, _WEB_AUTOSTART_KEY, False)
        )
        self._web_autostart_cb.blockSignals(False)

    def _maybe_autostart_web(self) -> None:
        """Start the monitor when the autostart checkbox is on (no browser)."""

        if not getattr(self, "_web_autostart_cb", None):
            return
        if not self._web_autostart_cb.isChecked():
            return
        if self._web_server.running:
            return
        self._web_open_when_ready = False
        self._web_server.start()
        self._update_web_label()

    def _on_web_autostart_toggled(self, checked: bool) -> None:
        self._settings.setValue(_WEB_AUTOSTART_KEY, checked)
        self._settings.sync()
        if checked and not self._web_server.running:
            self._web_open_when_ready = False
            self._web_server.start()
            self._update_web_label()

    def _on_web_open_clicked(self) -> None:
        if self._web_server.is_ready:
            self._open_web_browser()
            return
        self._web_open_when_ready = True
        self._web_open_btn.setEnabled(False)
        self._web_server.start()
        self._update_web_label()

    def _open_web_browser(self) -> None:
        webbrowser.open(self._web_server.local_url())

    def _copy_web_share_url(self) -> None:
        url = self._web_server.share_url()
        QApplication.clipboard().setText(url)
        self.statusBar().showMessage(f"Скопировано: {url}", 4000)

    def _update_web_label(self) -> None:
        if not hasattr(self, "_web_url_button"):
            return
        urls = "\n".join(self._web_server.display_urls())
        self._web_url_button.setToolTip(
            "0.0.0.0 в браузер не пишут.\n"
            f"{urls}\n"
            "Первая строка — этот ПК. Остальные — коллеги в LAN. "
            "Клик копирует адрес для коллег."
        )
        if self._web_server.is_ready:
            self._web_url_button.setText(self._web_server.share_url())
        elif self._web_server.running:
            self._web_url_button.setText("WEB запускается…")
        else:
            self._web_url_button.setText("WEB выкл.")

    @Slot()
    def _on_web_server_ready(self) -> None:
        if hasattr(self, "_web_open_btn"):
            self._web_open_btn.setEnabled(True)
        self._update_web_label()
        if self._web_open_when_ready:
            self._web_open_when_ready = False
            self._open_web_browser()
        self.statusBar().showMessage(f"WEB: {self._web_server.share_url()}", 8000)

    @Slot(str)
    def _on_web_server_failed(self, message: str) -> None:
        self._web_open_when_ready = False
        if hasattr(self, "_web_open_btn"):
            self._web_open_btn.setEnabled(True)
        self._update_web_label()
        self._append_log(f"WEB: {message}")
        QMessageBox.warning(
            self,
            "WEB монитор",
            f"Не удалось запустить монитор.\n{message}",
        )

    @Slot()
    def _on_web_server_stopped(self) -> None:
        self._web_open_when_ready = False
        if hasattr(self, "_web_open_btn"):
            self._web_open_btn.setEnabled(True)
        self._update_web_label()

    @Slot(str)
    def _on_web_server_log(self, line: str) -> None:
        self._append_log(f"WEB: {line}")

    def _restore_kits_filters(self) -> None:
        """Load Комплекты filter widgets from QSettings."""

        boxes = (
            (self._kits_mismatch, "window/kits_filter_mismatch"),
            (self._kits_no_rd, "window/kits_filter_no_rd"),
            (self._kits_no_robot, "window/kits_filter_no_robot"),
            (self._kits_no_google, "window/kits_filter_no_google"),
            (self._kits_code_a, "window/kits_filter_code_a"),
            (self._kits_tdo, "window/kits_filter_tdo"),
            (self._kits_no_as_build, "window/kits_filter_no_as_build"),
            (
                self._kits_only_as_build,
                "window/kits_filter_only_as_build",
            ),
            (self._kits_mto_problems, "window/kits_filter_mto_problems"),
            (self._kits_an_closes, "window/kits_filter_an_closes"),
        )
        widgets = (self._kits_filter, *(box for box, _key in boxes))
        for widget in widgets:
            widget.blockSignals(True)
        try:
            self._kits_filter.setText(
                str(self._settings.value("window/kits_filter_text") or "")
            )
            for box, key in boxes:
                box.setChecked(_settings_bool(self._settings, key, False))
        finally:
            for widget in widgets:
                widget.blockSignals(False)
        if self._kits_no_as_build.isChecked():
            self._kits_only_as_build.setEnabled(False)
        elif self._kits_only_as_build.isChecked():
            self._kits_no_as_build.setEnabled(False)
        self._apply_kits_filter()

    def _save_kits_filters(self) -> None:
        """Persist Комплекты filter widgets to QSettings."""

        self._settings.setValue(
            "window/kits_filter_text", self._kits_filter.text()
        )
        self._settings.setValue(
            "window/kits_filter_mismatch", self._kits_mismatch.isChecked()
        )
        self._settings.setValue(
            "window/kits_filter_no_rd", self._kits_no_rd.isChecked()
        )
        self._settings.setValue(
            "window/kits_filter_no_robot", self._kits_no_robot.isChecked()
        )
        self._settings.setValue(
            "window/kits_filter_no_google", self._kits_no_google.isChecked()
        )
        self._settings.setValue(
            "window/kits_filter_code_a", self._kits_code_a.isChecked()
        )
        self._settings.setValue(
            "window/kits_filter_tdo", self._kits_tdo.isChecked()
        )
        self._settings.setValue(
            "window/kits_filter_no_as_build",
            self._kits_no_as_build.isChecked(),
        )
        self._settings.setValue(
            "window/kits_filter_only_as_build",
            self._kits_only_as_build.isChecked(),
        )
        self._settings.setValue(
            "window/kits_filter_mto_problems",
            self._kits_mto_problems.isChecked(),
        )
        self._settings.setValue(
            "window/kits_filter_an_closes",
            self._kits_an_closes.isChecked(),
        )

    def closeEvent(self, event: QCloseEvent) -> None:
        """Persist layout and avoid destroying a running worker."""

        set_perf_sink(None)
        startup_pending = self._startup_load_pending
        secondary_pending = self._secondary_tabs_pending
        self._stop_startup_timer()
        self._stop_secondary_timer()
        if hasattr(self, "_doc_filter_timer"):
            self._doc_filter_timer.stop()
        if hasattr(self, "_auto_mto_compare_refresh_timer"):
            self._auto_mto_compare_refresh_timer.stop()
        startup_thread = getattr(self, "_startup_thread", None)
        if startup_thread is not None:
            try:
                startup_thread.finished.disconnect(
                    self._on_startup_hydrate_finished
                )
            except (RuntimeError, TypeError):
                pass
            if startup_thread.isRunning() and not startup_thread.wait(15000):
                QMessageBox.information(
                    self,
                    "Каталог РД",
                    "Первая загрузка ещё идёт. Повторите закрытие через несколько секунд.",
                )
                self._restore_deferred_load_timers(startup_pending, secondary_pending)
                event.ignore()
                return
            self._startup_thread = None
            startup_thread.deleteLater()
            self._startup_load_running = False
        if hasattr(self, "_revision_matrix_tab"):
            if not self._revision_matrix_tab.prepare_close():
                QMessageBox.information(
                    self,
                    "Ревизии MTO",
                    "Сравнение или копирование MTO ещё завершается. "
                    "Повторите закрытие через несколько секунд.",
                )
                self._restore_deferred_load_timers(startup_pending, secondary_pending)
                event.ignore()
                return
        if hasattr(self, "_an_tab"):
            if not self._an_tab.prepare_close():
                QMessageBox.information(
                    self,
                    "АН",
                    "Сверка АН ещё завершается. Повторите закрытие через несколько секунд.",
                )
                self._restore_deferred_load_timers(startup_pending, secondary_pending)
                event.ignore()
                return
        if hasattr(self, "_rd_dump_tab"):
            if not self._rd_dump_tab.prepare_close():
                QMessageBox.information(
                    self,
                    "РД",
                    "Сверка РД ещё завершается. Повторите закрытие через несколько секунд.",
                )
                self._restore_deferred_load_timers(startup_pending, secondary_pending)
                event.ignore()
                return
        transfer_thread = getattr(self, "_transfer_mto_thread", None)
        if transfer_thread is not None and transfer_thread.isRunning():
            transfer_thread.requestInterruption()
            if not transfer_thread.wait(3000):
                QMessageBox.information(
                    self,
                    "Комплекты",
                    "Сверка MTO спорных передач ещё идёт. "
                    "Повторите закрытие через несколько секунд.",
                )
                self._restore_deferred_load_timers(startup_pending, secondary_pending)
                event.ignore()
                return
        if self._scan_thread is not None:
            self._scan_thread.request_cancel()
            if not self._scan_thread.wait(3000):
                QMessageBox.information(
                    self,
                    "Сканирование",
                    "Сканирование ещё завершается. Повторите закрытие через несколько секунд.",
                )
                self._restore_deferred_load_timers(startup_pending, secondary_pending)
                event.ignore()
                return
        if self._an_scan_thread is not None:
            self._an_scan_thread.request_cancel()
            if not self._an_scan_thread.wait(3000):
                QMessageBox.information(
                    self,
                    "Скан АН",
                    "Скан АН ещё завершается. Повторите закрытие через несколько секунд.",
                )
                self._restore_deferred_load_timers(startup_pending, secondary_pending)
                event.ignore()
                return
        if self._rd_dump_scan_thread is not None:
            self._rd_dump_scan_thread.request_cancel()
            if not self._rd_dump_scan_thread.wait(3000):
                QMessageBox.information(
                    self,
                    "Скан РД",
                    "Скан РД ещё завершается. Повторите закрытие через несколько секунд.",
                )
                self._restore_deferred_load_timers(startup_pending, secondary_pending)
                event.ignore()
                return
        if self._google_thread is not None:
            if not self._google_thread.wait(3000):
                QMessageBox.information(
                    self,
                    "Комплекты Google",
                    "Загрузка Google ещё идёт. Повторите закрытие через несколько секунд.",
                )
                self._restore_deferred_load_timers(startup_pending, secondary_pending)
                event.ignore()
                return
        if self._google_write_thread is not None:
            if not self._google_write_thread.wait(3000):
                QMessageBox.information(
                    self,
                    "Письма о согласовании",
                    "Запись в КСБ ИД ещё идёт. Повторите закрытие через несколько секунд.",
                )
                self._restore_deferred_load_timers(startup_pending, secondary_pending)
                event.ignore()
                return
        if self._robot_sync_thread is not None:
            if not self._robot_sync_thread.wait(3000):
                QMessageBox.information(
                    self,
                    "MTO робота",
                    "Копирование MTO ещё идёт. Повторите закрытие через несколько секунд.",
                )
                self._restore_deferred_load_timers(startup_pending, secondary_pending)
                event.ignore()
                return
        if self._sq_to_rd_thread is not None:
            if not self._sq_to_rd_thread.wait(3000):
                QMessageBox.information(
                    self,
                    "Перенос SQ в РД",
                    "Перенос папки SQ ещё идёт. Повторите закрытие через несколько секунд.",
                )
                self._restore_deferred_load_timers(startup_pending, secondary_pending)
                event.ignore()
                return
        if self._mto_compare_thread is not None:
            self._mto_compare_thread.request_cancel()
            if not self._mto_compare_thread.wait(3000):
                QMessageBox.information(
                    self,
                    "Сверка MTO",
                    "Сверка MTO ещё завершается. Повторите закрытие через несколько секунд.",
                )
                self._restore_deferred_load_timers(startup_pending, secondary_pending)
                event.ignore()
                return
        if hasattr(self, "_mail_drop_filter"):
            app = QApplication.instance()
            if app is not None:
                app.removeEventFilter(self._mail_drop_filter)
        self._settings.setValue("window/geometry", self.saveGeometry())
        self._settings.setValue(
            "window/doc_splitter", self._doc_splitter.saveState()
        )
        if hasattr(self, "_kits_splitter"):
            self._persist_kits_detail_layout()
        if hasattr(self, "_kits_detail_tabs"):
            self._settings.setValue(
                _KITS_DETAIL_TAB_KEY, self._kits_detail_tabs.currentIndex()
            )
        self._save_table_headers()
        self._save_tree_label_options()
        if hasattr(self, "_revision_matrix_tab"):
            self._revision_matrix_tab.save_filters(self._settings)
        if hasattr(self, "_mto_worklist_tab"):
            self._mto_worklist_tab.save_filters(self._settings)
        if hasattr(self, "_issuance_journal_tab"):
            self._issuance_journal_tab.save_filters(self._settings)
        if hasattr(self, "_an_tab"):
            self._an_tab.save_filters(self._settings)
        if hasattr(self, "_rd_dump_tab"):
            self._rd_dump_tab.save_filters(self._settings)
        if hasattr(self, "_approval_mail_tab"):
            self._approval_mail_tab.save_settings(self._settings)
        if hasattr(self, "_kits_filter"):
            self._save_kits_filters()
        if hasattr(self, "_web_autostart_cb"):
            self._settings.setValue(
                _WEB_AUTOSTART_KEY, self._web_autostart_cb.isChecked()
            )
        if hasattr(self, "_web_server"):
            for signal, slot in (
                (self._web_server.ready, self._on_web_server_ready),
                (self._web_server.failed, self._on_web_server_failed),
                (self._web_server.stopped, self._on_web_server_stopped),
                (self._web_server.log_line, self._on_web_server_log),
            ):
                try:
                    signal.disconnect(slot)
                except (RuntimeError, TypeError):
                    pass
            self._web_server.stop()
        self._settings.sync()
        shutdown_context_menu_usage()
        super().closeEvent(event)
