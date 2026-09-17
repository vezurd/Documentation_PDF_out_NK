"""Heatmap tab: title–mark rows by filename-revision columns.

Qt monitor only. Cell fill and letters come from ``list_revision_matrix``.
The third column («В папку робота») is the export cockpit: one
:class:`~rd_catalog.mto_export.ExportSelection` per kit. «РД · рев.» shows
the official RD filename revision on disk (missing send/F as a grey
status). «Авто МТО» is the customer-PI
revision vs the export file. «Сверка Авто МТО» is the content-compare
status against that same file (latest RD, or the manual pin).
"""

from __future__ import annotations

import json
from collections.abc import Callable, Collection, Mapping, Sequence
from dataclasses import replace
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from PySide6.QtCore import QEvent, QObject, QSettings, Qt, Signal, Slot
from PySide6.QtGui import QAction, QBrush, QColor, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.auto_mto_compare_cache import (
    cache_key,
    entry_from_result,
    load_auto_mto_compare_cache,
    rd_mtime_ns,
    result_from_entry,
    save_auto_mto_compare_cache,
)
from rd_catalog.customer_pi_auto_mto import (
    AUTO_MTO_COMPARE_STATUS_HEADER,
    AUTO_MTO_COMPARE_STATUS_TOOLTIP,
    AutoMtoCompareResult,
    AutoMtoCompareStatus,
    AutoMtoFile,
    auto_mto_compare_status,
    auto_mto_rd_path_key,
    auto_mto_rd_paths_match,
    needs_auto_mto_compare,
)
from rd_catalog.config import CatalogConfig
from rd_catalog.context_menu_qt import exec_tracked_menu
from rd_catalog.context_menu_usage import MENU_HEATMAP
from rd_catalog.customer_pi_compare_thread import AutoMtoCompareThread
from rd_catalog.db import CatalogDatabase
from rd_catalog.kits import (
    SourceKitSnapshot,
    aggregate_source_kits,
    kit_identity_key,
    parse_sheet_revision,
)
from rd_catalog.models import FileRecord, SourceKind
from rd_catalog.perf_log import perf_span
from rd_catalog.mto_export import (
    DEFAULT_EXPORT_RULE,
    EXPORT_RULES,
    ExportPin,
    ExportPinCandidate,
    ExportSelection,
    ExportTarget,
    default_robot_export_target,
    export_pin_evidence_for_kit,
    list_export_pin_candidates,
    load_export_pins,
    load_export_targets,
    remove_export_pin,
    resolve_export_selections,
    save_export_pins,
    save_export_targets,
    scan_export_target,
    target_files_from_records,
    upsert_export_pin,
)
from rd_catalog.mto_export_copy import (
    EXPORT_RULE_LABELS,
    EXPORT_STATE_COLOR_KEY,
    EXPORT_STATE_TEXT,
    QSETTINGS_EXPORT_TARGET,
    export_selection_tooltip,
    file_id_verdicts_from_db,
    format_export_copy_report,
    format_pool_label,
    overlay_session_on_selections,
    pair_identity_key,
    patch_selections_with_verdicts,
    plan_for_visible_rows,
    preview_copy_count,
    selections_by_kit,
    session_status_map,
)
from rd_catalog.mto_export_dialog import MtoExportPreviewDialog, MtoExportProgressDialog
from rd_catalog.mto_export_thread import MtoExportCopyThread
from rd_catalog.mto_pair_compare import (
    MtoFilePair,
    PairPoolStatus,
    pair_pool_status,
    pairs_from_export_selections,
)
from rd_catalog.mto_pair_compare_thread import MtoPairCompareThread
from rd_catalog.overlay import revision_rank
from rd_catalog.parse import record_has_canonical_layout
from rd_catalog.robot_mto_sync import format_size
from rd_catalog.monitor_qt import apply_monitor_cell
from rd_catalog.monitor_views import (
    RD_AB_SUFFIX,
    REV_DIFF_FILL,
    REV_MATCH_FILL,
    TDO_STATUSES as _TDO_STATUSES,
    _compare_status_cell,
    _heatmap_auto_mto_cell,
    _heatmap_rd_rev_cell,
)
from rd_catalog.status_colors import color_for, status_color_label

_ROLE_CELL = Qt.ItemDataRole.UserRole
_ROLE_KIT = Qt.ItemDataRole.UserRole + 1
_ROLE_EXPORT = Qt.ItemDataRole.UserRole + 2
_ROLE_EXPORT_COLOR = Qt.ItemDataRole.UserRole + 3
_ROLE_SORT = Qt.ItemDataRole.UserRole + 4

_COL_TITLE = 0
_COL_MARK = 1
_COL_EXPORT = 2
_COL_RD_REV = 3
_COL_AUTO_MTO = 4
_COL_AUTO_MTO_COMPARE = 5
_COL_REV_FIRST = 6

_FIXED_COLUMNS = (
    "Титул",
    "Марка",
    "В папку робота",
    "РД · рев.",
    "Авто МТО",
    AUTO_MTO_COMPARE_STATUS_HEADER,
)
_REV_MATCH_BG = QColor(REV_MATCH_FILL)
_REV_DIFF_BG = QColor(REV_DIFF_FILL)
_LEGEND_KEYS = (
    "empty",
    "no_mto",
    "not_uploaded",
    "sent_tdo",
    "tdo_review",
    "code_a",
    "code_b",
    "code_c",
    "working",
    "problem",
    "current",
    "us_build",
)
_EXPORT_LEGEND_KEYS = (
    "export_add",
    "export_replace",
    "export_same",
    "export_missing",
    "export_unknown",
    "export_pin_stale",
)
_EXPORT_GATE_TOOLTIP = (
    "Сравнение ещё не завершено — выгрузка по неполным данным недоступна."
)
_EXPORT_READY_TOOLTIP = (
    "Скопировать MTO всех видимых строк в выбранную папку."
)
_PIN_ACTION_LABEL = "Выбрать файл MTO вручную…"
_UNPIN_ACTION_LABEL = "Снять ручной выбор"
_AUTO_MTO_COMPARE_ACTION_LABEL = "Сверить Авто МТО с выбранным MTO РД…"
_PIN_DIALOG_TITLE = "Выбор файла MTO"
_PIN_CONFIRM_TITLE = "Ручной выбор MTO"
_PIN_CONFIRM_TEXT = "Заменить текущий ручной выбор другим файлом?"
_PIN_EMPTY_TEXT = "Нет файлов MTO для этого комплекта."
_PIN_DIALOG_HEADERS = ("Пакет", "NN", "Файл", "Рев.", "Размер", "Правило")
_PIN_RULE_MARK = "да"
_AUTO_MTO_COMPARE_KIND_ORDER = {
    "matched": 0,
    "composite": 1,
    "soft": 2,
    "rev_match": 3,
    "not_compared": 4,
    "no_match": 5,
    "stale": 6,
    "no_rd": 7,
    "no_auto": 8,
}
_PERSISTABLE_COMPARE_KINDS = frozenset({"single", "composite", "no_match"})


class _SortItem(QTableWidgetItem):
    """Table item that sorts by ``_ROLE_SORT`` when both sides have it."""

    def __lt__(self, other: QTableWidgetItem) -> bool:
        left = self.data(_ROLE_SORT)
        right = other.data(_ROLE_SORT) if other is not None else None
        if left is not None and right is not None:
            return bool(left < right)
        return super().__lt__(other)


def _revision_column_key(text: str) -> tuple[int, int, str]:
    revision, appendix = parse_sheet_revision(text)
    return revision_rank(revision, appendix)


def _title_sort_key(text: str) -> tuple[int, int | str]:
    stripped = str(text or "").strip()
    if stripped.isdigit():
        return (0, int(stripped))
    return (1, stripped.casefold())


def _shown_revision_sort_key(
    text: str,
    *,
    strip_ab: bool = False,
) -> tuple[int, tuple[int, int, str]]:
    token = str(text or "").strip()
    if " · " in token:
        token = token.split(" · ", 1)[0]
    if strip_ab and token.endswith(RD_AB_SUFFIX):
        token = token[: -len(RD_AB_SUFFIX)].rstrip()
    if token.endswith(")") and "(" in token:
        token = token[: token.rfind("(")].rstrip()
    if not token or token == "—":
        return (1, (0, 0, ""))
    return (0, _revision_column_key(token))


def _compare_kind_sort_key(kind: str) -> tuple[int, str]:
    return (_AUTO_MTO_COMPARE_KIND_ORDER.get(kind, 99), kind)


def _heatmap_cell_sort_key(
    cell: Any | None,
    letters: str,
) -> tuple[int, tuple[str, str]]:
    if cell is None and not letters:
        return (1, ("", ""))
    status = _cell_status(cell) if cell is not None else ""
    return (0, (letters or "", status or ""))


def _auto_mto_files_sig(
    files: Sequence[AutoMtoFile],
) -> tuple[tuple[str, str], ...]:
    return tuple(sorted((item.relpath, item.fingerprint) for item in files))


def _cell_status(cell: Any) -> str:
    return str(getattr(cell, "pipeline_status", "") or "")


def _cell_letters(cell: Any) -> str:
    return str(getattr(cell, "letters", "") or "")


def _cell_problems(cell: Any) -> tuple[str, ...]:
    raw = getattr(cell, "problem_kinds_json", None)
    if isinstance(raw, str):
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            parsed = []
        return tuple(str(item) for item in parsed if item)
    kinds = getattr(cell, "problem_kinds", None)
    if kinds:
        return tuple(str(item) for item in kinds)
    return ()


def _bool_attr(cell: Any, name: str) -> bool:
    value = getattr(cell, name, False)
    if isinstance(value, (int, str)) and str(value).isdigit():
        return bool(int(value))
    return bool(value)


class MtoPinPickDialog(QDialog):
    """List every MTO candidate for one kit and return the chosen file."""

    def __init__(
        self,
        candidates: Sequence[ExportPinCandidate],
        *,
        current_path: str = "",
        parent: QWidget | None = None,
    ) -> None:
        """Build a table of package / file / revision / size rows.

        Args:
            candidates: Newest-transfer-first list from
                :func:`~rd_catalog.mto_export.list_export_pin_candidates`.
            current_path: Already-pinned file, if any (preselected).
            parent: Optional Qt parent.
        """

        super().__init__(parent)
        self._candidates = tuple(candidates)
        self.setWindowTitle(_PIN_DIALOG_TITLE)
        self.resize(780, 420)
        layout = QVBoxLayout(self)
        hint = QLabel(
            "Каждая строка — файл MTO в одной передаче. "
            "«да» в колонке «Правило» — то, что правило выбрало бы сейчас.",
            self,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)
        table = QTableWidget(len(self._candidates), len(_PIN_DIALOG_HEADERS), self)
        table.setHorizontalHeaderLabels(list(_PIN_DIALOG_HEADERS))
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.verticalHeader().setVisible(False)
        table.setAlternatingRowColors(True)
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        header.setStretchLastSection(True)
        current_key = current_path.casefold()
        select_row = 0
        for row, candidate in enumerate(self._candidates):
            values = (
                candidate.package_name or candidate.package_path,
                "" if candidate.sequence is None else f"{int(candidate.sequence):02d}",
                candidate.file_name,
                candidate.revision_text or "—",
                format_size(candidate.size),
                _PIN_RULE_MARK if candidate.is_rule_choice else "",
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, candidate)
                    item.setToolTip(candidate.package_path)
                elif column == 2:
                    item.setToolTip(candidate.file_path)
                table.setItem(row, column, item)
            if candidate.is_rule_choice:
                select_row = row
            if current_key and candidate.file_path.casefold() == current_key:
                select_row = row
        table.selectRow(select_row)
        table.cellDoubleClicked.connect(lambda *_args: self.accept())
        layout.addWidget(table, 1)
        self._table = table
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel,
            self,
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def selected_candidate(self) -> ExportPinCandidate | None:
        """Return the highlighted candidate, or ``None``."""

        row = self._table.currentRow()
        if row < 0 or row >= len(self._candidates):
            return None
        return self._candidates[row]


class RevisionMatrixTab(QWidget):
    """Title–mark × filename-revision heatmap for MTO pipeline status."""

    kit_activated = Signal(str, str)
    export_copy_finished = Signal(str, bool)
    export_pins_changed = Signal()
    pair_compare_log = Signal(str)
    auto_mto_compare_finished = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._palette: dict[str, str] = {}
        self._cells: tuple[Any, ...] = ()
        self._is_banned: Callable[[str, str], bool] = lambda _t, _m: False
        self._allowed_kits: set[tuple[str, str]] | None = None
        self._config: CatalogConfig | None = None
        self._database: CatalogDatabase | None = None
        self._records: tuple[FileRecord, ...] = ()
        self._records_in_contour: bool = False
        self._detected_current_ids: set[int] = set()
        self._rd_by_kit: dict[tuple[str, str], SourceKitSnapshot] = {}
        self._targets: list[ExportTarget] = []
        self._target_files: dict[tuple[str, str], str] = {}
        self._selections: tuple[ExportSelection, ...] = ()
        self._selections_by_key: dict[tuple[str, str], ExportSelection] = {}
        self._session_verdicts: dict[tuple[str, str], str] = {}
        self._pool_status = PairPoolStatus(total=0, compared=0, pending=0, failed=0)
        self._current_pairs: tuple[MtoFilePair, ...] = ()
        self._pair_thread: MtoPairCompareThread | None = None
        self._pair_thread_key: frozenset[tuple[str, str]] = frozenset()
        self._pair_queued_pairs: tuple[MtoFilePair, ...] | None = None
        self._copy_thread: MtoExportCopyThread | None = None
        self._copy_dialog: MtoExportProgressDialog | None = None
        self._updating_controls = False
        self._last_target_index = 0
        self._auto_mto_by_kit: dict[tuple[str, str], tuple[AutoMtoFile, ...]] = {}
        self._auto_mto_comparisons: dict[
            tuple[str, str], tuple[str, AutoMtoCompareResult]
        ] = {}
        self._auto_mto_comparison_sigs: dict[
            tuple[str, str], tuple[tuple[str, str], ...]
        ] = {}
        self._auto_mto_compare_thread: AutoMtoCompareThread | None = None
        self._auto_mto_compare_context: tuple[tuple[str, str], str] | None = None
        self._auto_mto_queue: list[tuple[tuple[str, str], str]] = []
        self._auto_mto_queue_paused = False
        self._auto_mto_queue_done = 0
        self._auto_mto_queue_total = 0
        self._mtime_by_path: dict[str, int] = {}
        self._pipelines_by_kit: dict[tuple[str, str], Any] = {}
        self._build()

    def table(self) -> QTableWidget:
        return self._table

    def is_copy_running(self) -> bool:
        """Return True while a bulk export copy worker is alive."""

        return self._copy_thread is not None and self._copy_thread.isRunning()

    def current_rule(self) -> str:
        """Return the selected export rule id."""

        rule = str(self._rule_combo.currentData() or "")
        return rule if rule in EXPORT_RULES else DEFAULT_EXPORT_RULE

    def filter_text(self) -> str:
        """Return the current heatmap text filter."""

        return self._filter.text()

    def current_target(self) -> ExportTarget | None:
        """Return the selected named destination, if any."""

        index = self._target_combo.currentIndex()
        if 0 <= index < len(self._targets):
            return self._targets[index]
        return None

    def visible_export_selections(self) -> tuple[ExportSelection, ...]:
        """Return export selections for currently visible heatmap rows.

        The text filter and checkboxes define the batch.

        Returns:
            One selection per visible kit that has a stored
            :class:`ExportSelection`.
        """

        table = self._table
        result: list[ExportSelection] = []
        for row in range(table.rowCount()):
            if table.isRowHidden(row):
                continue
            item = table.item(row, _COL_EXPORT)
            selection = item.data(_ROLE_EXPORT) if item is not None else None
            if isinstance(selection, ExportSelection):
                result.append(selection)
                continue
            title_item = table.item(row, _COL_TITLE)
            kit = title_item.data(_ROLE_KIT) if title_item else None
            if isinstance(kit, tuple) and len(kit) == 2:
                stored = self._selections_by_key.get(kit_identity_key(*kit))
                if stored is not None:
                    result.append(stored)
        return tuple(result)

    def apply_export_selections(
        self, selections: Sequence[ExportSelection]
    ) -> None:
        """Replace cached selections and repaint the third column only.

        Args:
            selections: Rows from :func:`resolve_export_selections` or tests.
        """

        self._selections = tuple(selections)
        self._selections_by_key = selections_by_kit(self._selections)
        self._current_pairs = pairs_from_export_selections(self._selections)
        self._fill_export_column()
        self._fill_auto_mto_column()
        self.enqueue_needed_auto_mto_compares()
        self._update_export_button()

    def selection_for(self, title: str, mark: str) -> ExportSelection | None:
        """Return the cached export selection for one kit, if any.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.

        Returns:
            Stored :class:`ExportSelection`, or ``None``.
        """

        return self._selections_by_key.get(kit_identity_key(title, mark))

    def pin_for(self, title: str, mark: str) -> ExportPin | None:
        """Return the persisted pin for one kit, if any.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.

        Returns:
            Stored :class:`ExportPin`, or ``None``.
        """

        if self._config is None:
            return None
        key = kit_identity_key(title, mark)
        for pin in load_export_pins(self._config):
            if kit_identity_key(pin.title, pin.mark) == key:
                return pin
        return None

    def assign_export_pin(
        self, title: str, mark: str, candidate: ExportPinCandidate
    ) -> bool:
        """Write a pin for ``title``/``mark`` and refresh the third column.

        Does not ask for confirmation. Does not delete other pins.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
            candidate: File chosen in the pick dialog.

        Returns:
            True when the pin was stored.
        """

        if self._config is None or self._database is None:
            return False
        if not str(candidate.file_path or "").strip():
            return False
        evidence = export_pin_evidence_for_kit(
            self._database,
            self._records,
            title=title,
            mark=mark,
            rd_root=self._config.rd_root,
        )
        pin = ExportPin(
            title=title,
            mark=mark,
            package_path=candidate.package_path,
            file_path=candidate.file_path,
            revision_text=candidate.revision_text,
            evidence=evidence,
            created_at=datetime.now(timezone.utc).isoformat(),
        )
        upsert_export_pin(self._config, pin)
        self._refresh_export(restart_compare=True)
        self.export_pins_changed.emit()
        return True

    def clear_export_pin(self, title: str, mark: str) -> bool:
        """Remove the pin for one kit and restore the rule's answer.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.

        Returns:
            True when pins were rewritten.
        """

        if self._config is None:
            return False
        remove_export_pin(self._config, title, mark)
        self._refresh_export(restart_compare=True)
        self.export_pins_changed.emit()
        return True

    def set_pool_status(self, status: PairPoolStatus) -> None:
        """Replace the pool counter (tests and live updates).

        Args:
            status: Exact remaining / compared / failed counts.
        """

        self._pool_status = status
        self._pool_label.setText(format_pool_label(status))
        self._update_export_button()

    def bind_catalog(
        self,
        config: CatalogConfig,
        database: CatalogDatabase,
    ) -> None:
        """Attach runtime config and SQLite so the cockpit can resolve rows.

        Args:
            config: Catalog configuration (``runtime_dir``, robot root).
            database: Initialized catalog database.
        """

        self._config = config
        self._database = database
        self._rebuild_rd_index()
        self._reload_targets(keep_name=None)

    def seed_rd_mtimes(self, records: Sequence[FileRecord]) -> None:
        """Fill path→mtime from the catalog snapshot (no UNC stat).

        ``hydrate_auto_mto_compare_cache`` uses this map so startup does
        not ``stat`` every cached RD workbook on the GUI thread.

        Args:
            records: Catalog file records from the last scan.
        """

        self._mtime_by_path = {
            record.path.casefold(): int(record.data.get("mtime_ns") or 0)
            for record in records
            if record.path
        }

    def set_export_records(
        self,
        records: Sequence[FileRecord],
        detected_current_ids: set[int],
        pipelines: Sequence[Any] | None = None,
        *,
        restart_compare: bool = True,
        in_contour: bool = False,
        kit_keys: Collection[tuple[str, str]] | None = None,
    ) -> None:
        """Refresh file snapshot used to resolve export selections.

        Args:
            records: Catalog file records from the last scan. Prefer the
                canonical contour, not the full ``list_files`` snapshot.
            detected_current_ids: Overlay-current RD file ids.
            pipelines: Kit pipelines for «РД · рев.» missing-transfer status.
            restart_compare: Restart the export pair-compare worker after
                resolving selections. File-unchanged Google F updates pass
                ``False`` when the heatmap is painted in place.
            in_contour: True when ``records`` are already
                ``records_in_contour`` / ``_pipeline_records``. Skips a
                second ``record_has_canonical_layout`` walk on the GUI thread.
            kit_keys: When set, repaint «РД · рев.» only for these identities.
                The RD index is still rebuilt from ``records``.
        """

        self._records = tuple(records)
        self._records_in_contour = bool(in_contour)
        self._detected_current_ids = set(detected_current_ids)
        self._pipelines_by_kit = {
            kit_identity_key(item.title, item.mark): item
            for item in (pipelines or ())
        }
        self._mtime_by_path = {
            record.path.casefold(): int(record.data.get("mtime_ns") or 0)
            for record in self._records
            if record.path
        }
        self._rebuild_rd_index()
        self._fill_rd_revision_column(kit_keys=kit_keys)
        self._refresh_export(restart_compare=restart_compare)

    def cancel_pair_compare(self) -> None:
        """Request cooperative cancel of the pair-compare child process."""

        self._pair_queued_pairs = None
        thread = self._pair_thread
        if thread is None:
            return
        thread.request_cancel()

    def prepare_close(self, *, wait_ms: int = 3000) -> bool:
        """Cancel workers and wait. Return False when the user should retry.

        Args:
            wait_ms: Milliseconds to wait for each worker.

        Returns:
            True when both workers are idle.
        """

        self._auto_mto_queue.clear()
        self._auto_mto_queue_paused = True
        self._reset_auto_mto_queue_counts_if_idle()
        self.cancel_pair_compare()
        if self._copy_thread is not None and self._copy_thread.isRunning():
            self._copy_thread.request_cancel()
            if not self._copy_thread.wait(wait_ms):
                return False
        if self._pair_thread is not None and self._pair_thread.isRunning():
            if not self._pair_thread.wait(wait_ms):
                return False
        if (
            self._auto_mto_compare_thread is not None
            and self._auto_mto_compare_thread.isRunning()
            and not self._auto_mto_compare_thread.wait(wait_ms)
        ):
            return False
        return True

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        status_row = QHBoxLayout()
        self._pool_label = QLabel("сравнение завершено", self)
        self._auto_mto_queue_label = QLabel("Авто МТО сверка: актуально", self)
        status_row.addWidget(self._pool_label, 1)
        status_row.addWidget(self._auto_mto_queue_label)
        layout.addLayout(status_row)

        cockpit = QHBoxLayout()
        cockpit.addWidget(QLabel("Правило:"))
        self._rule_combo = QComboBox(self)
        for rule, label in EXPORT_RULE_LABELS:
            self._rule_combo.addItem(label, rule)
        self._rule_combo.currentIndexChanged.connect(self._on_rule_changed)
        cockpit.addWidget(self._rule_combo, 1)
        cockpit.addWidget(QLabel("Папка:"))
        self._target_combo = QComboBox(self)
        self._target_combo.currentIndexChanged.connect(self._on_target_changed)
        cockpit.addWidget(self._target_combo, 1)
        self._add_target_button = QPushButton("Добавить…", self)
        self._add_target_button.clicked.connect(self._on_add_target)
        cockpit.addWidget(self._add_target_button)
        self._rename_target_button = QPushButton("Переименовать", self)
        self._rename_target_button.clicked.connect(self._on_rename_target)
        cockpit.addWidget(self._rename_target_button)
        self._remove_target_button = QPushButton("Удалить", self)
        self._remove_target_button.clicked.connect(self._on_remove_target)
        cockpit.addWidget(self._remove_target_button)
        self._folder_button = QPushButton("Папка…", self)
        self._folder_button.clicked.connect(self._on_choose_folder)
        cockpit.addWidget(self._folder_button)
        self._flat_box = QCheckBox("Плоская структура", self)
        self._flat_box.toggled.connect(self._on_flat_toggled)
        cockpit.addWidget(self._flat_box)
        layout.addLayout(cockpit)

        filters = QHBoxLayout()
        filters.addWidget(QLabel("Фильтр:"))
        self._filter = QLineEdit(self)
        self._filter.setPlaceholderText("Титул, марка, буквы ячейки…")
        self._filter.textChanged.connect(self._on_filter_text_changed)
        filters.addWidget(self._filter, 1)
        self._no_as_build = QCheckBox("Без as-build", self)
        self._no_as_build.setToolTip(
            "Скрыть as-build ячейки. «Только последние» тогда берёт IFC."
        )
        self._latest_only = QCheckBox("Только последние", self)
        self._problems_only = QCheckBox("Только проблемы", self)
        self._code_a_only = QCheckBox("Код A", self)
        self._tdo_only = QCheckBox("Прошли ТДО", self)
        for box in (
            self._no_as_build,
            self._latest_only,
            self._problems_only,
            self._code_a_only,
            self._tdo_only,
        ):
            box.toggled.connect(self._rebuild_table)
            filters.addWidget(box)
        self._export_button = QPushButton("Копировать видимые (0)", self)
        self._export_button.clicked.connect(self._on_export_clicked)
        filters.addWidget(self._export_button)
        layout.addLayout(filters)

        self._legend = QHBoxLayout()
        layout.addLayout(self._legend)
        self._export_legend = QHBoxLayout()
        layout.addLayout(self._export_legend)

        self._table = QTableWidget(0, len(_FIXED_COLUMNS), self)
        self._table.setHorizontalHeaderLabels(list(_FIXED_COLUMNS))
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._table.setAlternatingRowColors(True)
        self._table.verticalHeader().setVisible(False)
        self._table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._table.setWordWrap(False)
        header = self._table.horizontalHeader()
        header.setSectionsMovable(False)
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(16)
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        header.setTextElideMode(Qt.TextElideMode.ElideRight)
        header.setSortIndicatorShown(True)
        header.sectionClicked.connect(self._on_heatmap_header_clicked)
        self._table.setColumnWidth(_COL_TITLE, 70)
        self._table.setColumnWidth(_COL_MARK, 80)
        self._table.setColumnWidth(_COL_EXPORT, 150)
        self._table.setColumnWidth(_COL_RD_REV, 90)
        self._table.setColumnWidth(_COL_AUTO_MTO, 110)
        self._table.setColumnWidth(_COL_AUTO_MTO_COMPARE, 130)
        self._table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._table.customContextMenuRequested.connect(self._show_pin_context_menu)
        self._table.viewport().installEventFilter(self)
        layout.addWidget(self._table, 1)
        self._update_export_button()

    def reload_legend(self, palette: Mapping[str, str]) -> None:
        """Rebuild the color legend from the current palette."""

        self._palette = dict(palette)
        self._fill_legend_row(self._legend, _LEGEND_KEYS)
        self._fill_legend_row(self._export_legend, _EXPORT_LEGEND_KEYS)

    def _fill_legend_row(self, row: QHBoxLayout, keys: Sequence[str]) -> None:
        while row.count():
            item = row.takeAt(0)
            widget = item.widget() if item is not None else None
            if widget is not None:
                widget.deleteLater()
        for key in keys:
            sample = QLabel("  ", self)
            hex_color = color_for(self._palette, key)
            sample.setStyleSheet(
                f"background:{hex_color}; border:1px solid #80868b; min-width:16px;"
            )
            caption = QLabel(status_color_label(key), self)
            row.addWidget(sample)
            row.addWidget(caption)
        row.addStretch(1)

    def set_cells(
        self,
        cells: Sequence[Any],
        *,
        palette: Mapping[str, str],
        is_banned: Callable[[str, str], bool],
        allowed_kits: set[tuple[str, str]] | None = None,
    ) -> None:
        """Replace the heatmap snapshot and rebuild the table.

        Args:
            cells: Persisted ``kit_revision_cell`` rows.
            palette: Status color map.
            is_banned: Hide banned ``(title, mark)`` pairs.
            allowed_kits: If set, only these identities (same as «Все
                документы») are shown. ``None`` keeps every non-banned kit.
        """

        self._cells = tuple(cells)
        self._is_banned = is_banned
        self._allowed_kits = allowed_kits
        self.reload_legend(palette)
        self._rebuild_table()

    def patch_cells_for_kits(
        self,
        cells: Sequence[Any],
        kit_keys: Collection[tuple[str, str]],
    ) -> bool:
        """Replace heatmap cells for ``kit_keys`` without rebuilding columns.

        When the union of filename-revision columns changes (legalize adding
        a new rev), falls back to ``_rebuild_table``. Returns False when the
        table has never been painted so the caller can ``set_cells``.

        Args:
            cells: Full or partial ``kit_revision_cell`` snapshot. Rows whose
                identity is in ``kit_keys`` replace the previous cells for
                those kits; other identities in ``cells`` are ignored.
            kit_keys: Kit identities to patch.

        Returns:
            True when this tab applied the update; False when the caller
            must run a full ``set_cells``.
        """

        wanted = {kit_identity_key(title, mark) for title, mark in kit_keys}
        if not wanted or self._table.rowCount() == 0:
            return False
        old_columns = self._revision_columns(self._grouped_rows())
        kept = [
            cell
            for cell in self._cells
            if kit_identity_key(
                str(getattr(cell, "title", "") or ""),
                str(getattr(cell, "mark", "") or ""),
            )
            not in wanted
        ]
        incoming = [
            cell
            for cell in cells
            if kit_identity_key(
                str(getattr(cell, "title", "") or ""),
                str(getattr(cell, "mark", "") or ""),
            )
            in wanted
        ]
        self._cells = tuple(kept) + tuple(incoming)
        new_columns = self._revision_columns(self._grouped_rows())
        if new_columns != old_columns:
            self._rebuild_table()
            return True
        grouped = {
            kit_identity_key(title, mark): cell_map
            for title, mark, cell_map in self._grouped_rows()
        }
        table = self._table
        was_sorting = self._pause_table_sorting()
        try:
            for row in range(table.rowCount()):
                kit = self._kit_at_row(row)
                if kit is None:
                    continue
                key = kit_identity_key(*kit)
                if key not in wanted:
                    continue
                cell_map = grouped.get(key, {})
                for offset, rev in enumerate(new_columns):
                    cell = cell_map.get(rev)
                    shown = cell if self._visible_cell(cell) else None
                    item = self._make_cell_item(shown)
                    item.setData(_ROLE_KIT, kit)
                    table.setItem(row, _COL_REV_FIRST + offset, item)
            self._apply_row_visibility()
        finally:
            self._resume_table_sorting(was_sorting)
        return True

    def restore_filters(self, settings: QSettings) -> None:
        """Load filter widgets from QSettings without emitting rebuild twice."""

        stored_filter = ""
        self._filter.blockSignals(True)
        self._no_as_build.blockSignals(True)
        self._latest_only.blockSignals(True)
        self._problems_only.blockSignals(True)
        self._code_a_only.blockSignals(True)
        self._tdo_only.blockSignals(True)
        try:
            stored_filter = str(settings.value("window/rev_matrix_filter") or "")
            self._filter.setText(stored_filter)
            self._no_as_build.setChecked(
                _settings_bool(settings, "window/rev_matrix_no_as_build", False)
            )
            self._latest_only.setChecked(
                _settings_bool(settings, "window/rev_matrix_latest", False)
            )
            self._problems_only.setChecked(
                _settings_bool(settings, "window/rev_matrix_problems", False)
            )
            self._code_a_only.setChecked(
                _settings_bool(settings, "window/rev_matrix_code_a", False)
            )
            self._tdo_only.setChecked(
                _settings_bool(settings, "window/rev_matrix_tdo", False)
            )
        finally:
            self._filter.blockSignals(False)
            self._no_as_build.blockSignals(False)
            self._latest_only.blockSignals(False)
            self._problems_only.blockSignals(False)
            self._code_a_only.blockSignals(False)
            self._tdo_only.blockSignals(False)
        target_name = str(settings.value(QSETTINGS_EXPORT_TARGET) or "")
        if target_name:
            self._select_target_name(target_name)
        else:
            self._apply_target_memory()
        target = self.current_target()
        if target is not None and not target.filter_text and stored_filter:
            self._filter.blockSignals(True)
            self._filter.setText(stored_filter)
            self._filter.blockSignals(False)

    def save_filters(self, settings: QSettings) -> None:
        settings.setValue("window/rev_matrix_filter", self._filter.text())
        settings.setValue("window/rev_matrix_no_as_build", self._no_as_build.isChecked())
        settings.setValue("window/rev_matrix_latest", self._latest_only.isChecked())
        settings.setValue("window/rev_matrix_problems", self._problems_only.isChecked())
        settings.setValue("window/rev_matrix_code_a", self._code_a_only.isChecked())
        settings.setValue("window/rev_matrix_tdo", self._tdo_only.isChecked())
        target = self.current_target()
        if target is not None:
            settings.setValue(QSETTINGS_EXPORT_TARGET, target.name)
        self._persist_current_target_fields()

    def select_target_by_name(self, name: str) -> None:
        """Select a named destination (tests and restore).

        Args:
            name: Target display name.
        """

        self._select_target_name(name)

    def _grouped_rows(self) -> list[tuple[str, str, dict[str, Any]]]:
        grouped: dict[tuple[str, str], dict[str, Any]] = {}
        order: list[tuple[str, str]] = []
        display: dict[tuple[str, str], tuple[str, str]] = {}
        for cell in self._cells:
            title = str(getattr(cell, "title", "") or "")
            mark = str(getattr(cell, "mark", "") or "")
            if not title or not mark:
                continue
            if self._kit_is_hidden(title, mark):
                continue
            key = kit_identity_key(title, mark)
            if key not in grouped:
                grouped[key] = {}
                order.append(key)
                display[key] = (title, mark)
            rev = str(getattr(cell, "revision_text", "") or "")
            if rev:
                grouped[key][rev] = cell
        rows: list[tuple[str, str, dict[str, Any]]] = []
        for key in order:
            title, mark = display[key]
            rows.append((title, mark, grouped[key]))
        rows.sort(key=lambda item: (item[0], item[1].casefold()))
        return rows

    def _kit_is_hidden(self, title: str, mark: str) -> bool:
        if self._is_banned(title, mark):
            return True
        if self._allowed_kits is None:
            return False
        return kit_identity_key(title, mark) not in self._allowed_kits

    def _revision_columns(
        self, rows: Sequence[tuple[str, str, dict[str, Any]]]
    ) -> list[str]:
        seen: set[str] = set()
        columns: list[str] = []
        for _title, _mark, cells in rows:
            for rev in cells:
                if rev not in seen:
                    seen.add(rev)
                    columns.append(rev)
        columns.sort(key=_revision_column_key)
        return columns

    def _pause_table_sorting(self) -> bool:
        table = self._table
        enabled = table.isSortingEnabled()
        table.setSortingEnabled(False)
        return enabled

    def _resume_table_sorting(
        self,
        enabled: bool,
        column: int | None = None,
        order: Qt.SortOrder | None = None,
    ) -> None:
        table = self._table
        header = table.horizontalHeader()
        header.setSortIndicatorShown(True)
        if not enabled:
            return
        table.setSortingEnabled(True)
        restore_column = header.sortIndicatorSection() if column is None else column
        restore_order = header.sortIndicatorOrder() if order is None else order
        if 0 <= restore_column < table.columnCount():
            table.sortByColumn(restore_column, restore_order)

    @Slot(int)
    def _on_heatmap_header_clicked(self, section: int) -> None:
        table = self._table
        if table.isSortingEnabled():
            return
        header = table.horizontalHeader()
        header.setSortIndicatorShown(True)
        table.setSortingEnabled(True)
        if 0 <= section < table.columnCount():
            table.sortByColumn(section, Qt.SortOrder.AscendingOrder)

    def _rebuild_table(self) -> None:
        rows = self._grouped_rows()
        columns = self._revision_columns(rows)
        table = self._table
        header = table.horizontalHeader()
        was_sorting = table.isSortingEnabled()
        sort_column = header.sortIndicatorSection()
        sort_order = header.sortIndicatorOrder()
        table.setSortingEnabled(False)
        table.clear()
        headers = list(_FIXED_COLUMNS) + columns
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setRowCount(0)
        for title, mark, cells in rows:
            row_index = table.rowCount()
            table.insertRow(row_index)
            title_item = _SortItem(title)
            mark_item = _SortItem(mark)
            title_item.setData(_ROLE_KIT, (title, mark))
            mark_item.setData(_ROLE_KIT, (title, mark))
            title_item.setData(_ROLE_SORT, _title_sort_key(title))
            mark_item.setData(_ROLE_SORT, mark.casefold())
            table.setItem(row_index, _COL_TITLE, title_item)
            table.setItem(row_index, _COL_MARK, mark_item)
            export_item = _SortItem("")
            export_item.setData(_ROLE_KIT, (title, mark))
            export_item.setData(_ROLE_SORT, "")
            table.setItem(row_index, _COL_EXPORT, export_item)
            rd_item = _SortItem("—")
            rd_item.setData(_ROLE_KIT, (title, mark))
            rd_item.setData(_ROLE_SORT, _shown_revision_sort_key("—"))
            table.setItem(row_index, _COL_RD_REV, rd_item)
            auto_item = _SortItem("—")
            auto_item.setData(_ROLE_KIT, (title, mark))
            auto_item.setData(_ROLE_SORT, _shown_revision_sort_key("—"))
            table.setItem(row_index, _COL_AUTO_MTO, auto_item)
            compare_item = _SortItem("—")
            compare_item.setData(_ROLE_KIT, (title, mark))
            compare_item.setData(_ROLE_SORT, _compare_kind_sort_key(""))
            table.setItem(row_index, _COL_AUTO_MTO_COMPARE, compare_item)
            for offset, rev in enumerate(columns):
                cell = cells.get(rev)
                shown = cell if self._visible_cell(cell) else None
                item = self._make_cell_item(shown)
                item.setData(_ROLE_KIT, (title, mark))
                table.setItem(row_index, _COL_REV_FIRST + offset, item)
        for offset, rev in enumerate(columns):
            header_item = table.horizontalHeaderItem(_COL_REV_FIRST + offset)
            if header_item is not None:
                header_item.setToolTip(rev)
                header_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self._fill_export_column()
        self._fill_rd_revision_column()
        self._fill_auto_mto_column()
        self._apply_row_visibility()
        if table.columnCount() >= 1:
            table.setColumnWidth(_COL_TITLE, 70)
        if table.columnCount() >= 2:
            table.setColumnWidth(_COL_MARK, 80)
        if table.columnCount() >= 3:
            table.setColumnWidth(_COL_EXPORT, 150)
        if table.columnCount() >= 4:
            table.setColumnWidth(_COL_RD_REV, 90)
        if table.columnCount() >= 5:
            table.setColumnWidth(_COL_AUTO_MTO, 110)
        if table.columnCount() >= 6:
            table.setColumnWidth(_COL_AUTO_MTO_COMPARE, 130)
            compare_header = table.horizontalHeaderItem(_COL_AUTO_MTO_COMPARE)
            if compare_header is not None:
                compare_header.setToolTip(AUTO_MTO_COMPARE_STATUS_TOOLTIP)
        self._apply_column_layout()
        self._update_export_button()
        self._resume_table_sorting(was_sorting, sort_column, sort_order)

    def _make_cell_item(self, cell: Any | None) -> QTableWidgetItem:
        if cell is None:
            item = _SortItem("")
            item.setData(_ROLE_SORT, _heatmap_cell_sort_key(None, ""))
            _paint_fill(item, color_for(self._palette, "empty"))
            item.setToolTip("Нет ревизии")
            return item
        letters = _cell_letters(cell)
        item = _SortItem(letters)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        item.setData(_ROLE_CELL, cell)
        item.setData(_ROLE_SORT, _heatmap_cell_sort_key(cell, letters))
        status = _cell_status(cell) or "empty"
        _paint_fill(item, color_for(self._palette, status))
        problems = _cell_problems(cell)
        if problems:
            item.setForeground(QBrush(QColor(color_for(self._palette, "problem"))))
            font = QFont(item.font())
            font.setBold(True)
            item.setFont(font)
        if _bool_attr(cell, "is_current"):
            font = QFont(item.font())
            font.setUnderline(True)
            item.setFont(font)
        if _bool_attr(cell, "is_current_ifc") and not _bool_attr(cell, "is_as_build"):
            font = QFont(item.font())
            font.setItalic(True)
            item.setFont(font)
        item.setToolTip(self._cell_tooltip(cell))
        return item

    def _make_export_item(
        self,
        selection: ExportSelection | None,
        *,
        kit: tuple[str, str] | None,
    ) -> QTableWidgetItem:
        if selection is None:
            item = _SortItem("")
            item.setData(_ROLE_SORT, "")
            if kit is not None:
                item.setData(_ROLE_KIT, kit)
            item.setToolTip("Нет данных экспорта")
            return item
        state = selection.state
        text = EXPORT_STATE_TEXT.get(state, state)
        color_key = EXPORT_STATE_COLOR_KEY.get(state, "export_unknown")
        item = _SortItem(text)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        item.setData(_ROLE_EXPORT, selection)
        item.setData(_ROLE_EXPORT_COLOR, color_key)
        item.setData(_ROLE_KIT, (selection.title, selection.mark))
        item.setData(_ROLE_SORT, text)
        _paint_fill(item, color_for(self._palette, color_key))
        item.setToolTip(export_selection_tooltip(selection))
        return item

    def _fill_export_column(self) -> None:
        table = self._table
        if table.columnCount() < _COL_REV_FIRST:
            return
        was_sorting = self._pause_table_sorting()
        try:
            self._fill_export_column_rows()
        finally:
            if was_sorting:
                self._resume_table_sorting(True)

    def _fill_export_column_rows(self) -> None:
        table = self._table
        for row in range(table.rowCount()):
            title_item = table.item(row, _COL_TITLE)
            mark_item = table.item(row, _COL_MARK)
            kit = title_item.data(_ROLE_KIT) if title_item else None
            key = (
                kit_identity_key(kit[0], kit[1])
                if isinstance(kit, tuple) and len(kit) == 2
                else None
            )
            selection = self._selections_by_key.get(key) if key is not None else None
            kit_tuple = kit if isinstance(kit, tuple) and len(kit) == 2 else None
            table.setItem(
                row,
                _COL_EXPORT,
                self._make_export_item(selection, kit=kit_tuple),
            )
            loud = False
            if selection is not None and selection.state == "no_source":
                loud = True
            if title_item is not None and not title_item.text().strip():
                loud = True
            if mark_item is not None and not mark_item.text().strip():
                loud = True
            missing = color_for(self._palette, "export_missing")
            for identity_item in (title_item, mark_item):
                if identity_item is None:
                    continue
                if loud:
                    _paint_fill(identity_item, missing)
                else:
                    identity_item.setBackground(QBrush())
                    identity_item.setForeground(QBrush())

    def _rebuild_rd_index(self) -> None:
        """Aggregate the current RD overlay with kits-tab semantics."""

        records = self._records
        if not self._records_in_contour and self._config is not None:
            records = tuple(
                record
                for record in records
                if record_has_canonical_layout(record, self._config.rd_root)
            )
        self._rd_by_kit = aggregate_source_kits(
            records,
            source=SourceKind.RD,
            detected_current_ids=self._detected_current_ids,
        )

    def _current_ifc_by_kit(self) -> dict[tuple[str, str], str]:
        """Return latest IFC revisions identified by grouped heatmap cells."""

        mapping: dict[tuple[str, str], str] = {}
        for cell in self._cells:
            if not _bool_attr(cell, "is_current_ifc"):
                continue
            title = str(getattr(cell, "title", "") or "")
            mark = str(getattr(cell, "mark", "") or "")
            revision = str(getattr(cell, "revision_text", "") or "").strip()
            if not title or not mark or not revision:
                continue
            key = kit_identity_key(title, mark)
            previous = mapping.get(key)
            if previous is None or _revision_column_key(revision) > _revision_column_key(
                previous
            ):
                mapping[key] = revision
        return mapping

    def _fill_rd_revision_column(
        self, kit_keys: Collection[tuple[str, str]] | None = None
    ) -> None:
        """Paint the current RD overlay filename revision.

        Args:
            kit_keys: When set, only those heatmap rows.
        """

        table = self._table
        if table.columnCount() <= _COL_RD_REV:
            return
        was_sorting = self._pause_table_sorting()
        try:
            self._fill_rd_revision_column_rows(kit_keys=kit_keys)
        finally:
            if was_sorting:
                self._resume_table_sorting(True)

    def _fill_rd_revision_column_rows(
        self, kit_keys: Collection[tuple[str, str]] | None = None
    ) -> None:
        table = self._table
        wanted = (
            {kit_identity_key(title, mark) for title, mark in kit_keys}
            if kit_keys is not None
            else None
        )
        ifc_by_kit = self._current_ifc_by_kit()
        for row in range(table.rowCount()):
            kit = self._kit_at_row(row)
            key = kit_identity_key(*kit) if kit is not None else None
            if wanted is not None and (key is None or key not in wanted):
                continue
            snapshot = self._rd_by_kit.get(key) if key is not None else None
            pipeline = self._pipelines_by_kit.get(key) if key is not None else None
            cell = _heatmap_rd_rev_cell(
                snapshot,
                ifc_by_kit.get(key, "") if key is not None else "",
                pipeline=pipeline,
                palette=self._palette,
            )
            item = _SortItem()
            apply_monitor_cell(item, cell, sort_role=_ROLE_SORT)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setData(
                _ROLE_SORT,
                _shown_revision_sort_key(
                    str(cell.sort_key or cell.text), strip_ab=True
                ),
            )
            if kit is not None:
                item.setData(_ROLE_KIT, kit)
            table.setItem(row, _COL_RD_REV, item)

    def set_auto_mto_index(
        self,
        index: Mapping[tuple[str, str], AutoMtoFile | tuple[AutoMtoFile, ...]],
    ) -> None:
        """Store the customer-PI Auto MTO files and repaint its column."""

        normalized: dict[tuple[str, str], tuple[AutoMtoFile, ...]] = {}
        for key, value in index.items():
            if isinstance(value, AutoMtoFile):
                normalized[key] = (value,)
            else:
                normalized[key] = tuple(value)
        self._auto_mto_by_kit = normalized
        self.hydrate_auto_mto_compare_cache()

    def comparison_for(
        self,
        title: str,
        mark: str,
        rd_path: str,
    ) -> AutoMtoCompareResult | None:
        """Return a cached exact compare for this kit and RD workbook."""

        if not rd_path:
            return None
        cached = self._auto_mto_comparisons.get(kit_identity_key(title, mark))
        if cached is None:
            return None
        stored_path, result = cached
        if not auto_mto_rd_paths_match(stored_path, rd_path):
            return None
        return result

    def comparison_record(
        self,
        title: str,
        mark: str,
    ) -> tuple[str, AutoMtoCompareResult] | None:
        """Return the cached compare path and result for one kit, if any."""

        cached = self._auto_mto_comparisons.get(kit_identity_key(title, mark))
        if cached is None:
            return None
        stored_path, result = cached
        return str(stored_path), result

    def is_auto_mto_compare_running(self) -> bool:
        """Return True while a kit AutoMTO content compare is alive."""

        thread = self._auto_mto_compare_thread
        return thread is not None and thread.isRunning()

    def set_auto_mto_queue_paused(self, paused: bool) -> None:
        """Pause or resume starting the next queued AutoMTO compare.

        The kit already running may finish. Call this for scan / Google /
        robot / SQ→RD / export copy / customer PI, not for AutoMTO itself.

        Args:
            paused: True to stop starting further queued kits.
        """

        was_paused = self._auto_mto_queue_paused
        self._auto_mto_queue_paused = bool(paused)
        if was_paused and not self._auto_mto_queue_paused:
            self._kick_auto_mto_queue()

    def enqueue_auto_mto_compares(
        self,
        jobs: Sequence[tuple[str, str, str]],
    ) -> None:
        """Queue AutoMTO content compares for kits that still need them.

        Args:
            jobs: ``(title, mark, rd_path)`` triples. Empty paths,
                already-queued kits, and kits whose in-memory cache
                already matches that RD path are skipped.
        """

        added = 0
        for title, mark, rd_path in jobs:
            title_text = str(title or "").strip()
            mark_text = str(mark or "").strip()
            path = str(rd_path or "").strip()
            if not title_text or not mark_text or not path:
                continue
            key = kit_identity_key(title_text, mark_text)
            self._drop_other_auto_mto_jobs(key, path)
            if self._has_matching_auto_mto_comparison(key, path):
                continue
            if self._auto_mto_job_queued(key, path):
                continue
            self._auto_mto_queue.append((key, path))
            added += 1
        if added:
            self._auto_mto_queue_total += added
        self._kick_auto_mto_queue()

    def enqueue_needed_auto_mto_compares(self) -> None:
        """Queue kits whose painted «Сверка Авто МТО» still needs a run.

        Uses the same export ``source_path`` and cache record as the
        heatmap cell, so ``другой файл`` / ``не сверялось`` /
        ``рев. совпала`` enter the queue when the file of record changes
        (export resolve, pin, hydrate), not only on a heatmap refresh.
        """

        jobs: list[tuple[str, str, str]] = []
        for key, files in self._auto_mto_by_kit.items():
            if not files:
                continue
            rd_path = self._selection_rd_path(key)
            if not rd_path:
                continue
            selection = self._selections_by_key.get(key)
            cached = self._auto_mto_comparisons.get(key)
            status = auto_mto_compare_status(
                files=files,
                rd_path=rd_path,
                rd_revision=(
                    selection.source_revision_text if selection is not None else ""
                ),
                comparison=cached[1] if cached is not None else None,
                comparison_rd_path=cached[0] if cached is not None else "",
                pinned=bool(
                    selection is not None
                    and selection.origin in {"pin", "pin_stale"}
                ),
            )
            if needs_auto_mto_compare(status):
                jobs.append((files[0].title, files[0].mark, rd_path))
        self.enqueue_auto_mto_compares(jobs)

    def hydrate_auto_mto_compare_cache(self) -> None:
        """Keep valid in-memory compares and fill gaps from the runtime JSON.

        Does not wipe persistable results whose RD path and AutoMTO
        fingerprints still match the current index.
        """

        with perf_span("gui.hydrate_auto_mto_cache"):
            self._prune_and_hydrate_auto_mto_comparisons()
            self._fill_auto_mto_column()
            self.enqueue_needed_auto_mto_compares()

    def apply_auto_mto_compare_result(
        self,
        title: str,
        mark: str,
        rd_path: str,
        result: AutoMtoCompareResult,
    ) -> None:
        """Cache one exact AutoMTO comparison and repaint its kit cell."""

        key = kit_identity_key(title, mark)
        stored_path = str(rd_path)
        self._auto_mto_comparisons[key] = (stored_path, result)
        self._auto_mto_comparison_sigs[key] = _auto_mto_files_sig(
            self._auto_mto_by_kit.get(key, ())
        )
        self._persist_auto_mto_compare_result(title, mark, stored_path, result)
        self._fill_auto_mto_column()
        current = self._selection_rd_path(key)
        if current and not auto_mto_rd_paths_match(current, stored_path):
            self.enqueue_auto_mto_compares([(title, mark, current)])
            return
        if not self._auto_mto_queue and not self.is_auto_mto_compare_running():
            self.auto_mto_compare_finished.emit()

    def _fill_auto_mto_column(self) -> None:
        """Paint PI revision and content-compare status vs the export file."""

        table = self._table
        if table.columnCount() <= _COL_AUTO_MTO_COMPARE:
            return
        was_sorting = self._pause_table_sorting()
        try:
            self._fill_auto_mto_column_rows()
        finally:
            if was_sorting:
                self._resume_table_sorting(True)

    def _fill_auto_mto_column_rows(self) -> None:
        table = self._table
        for row in range(table.rowCount()):
            title_item = table.item(row, _COL_TITLE)
            kit = title_item.data(_ROLE_KIT) if title_item else None
            key = (
                kit_identity_key(kit[0], kit[1])
                if isinstance(kit, tuple) and len(kit) == 2
                else None
            )
            files = self._auto_mto_by_kit.get(key) if key is not None else None
            selection = (
                self._selections_by_key.get(key) if key is not None else None
            )
            auto_item = _SortItem("—")
            auto_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            auto_item.setData(_ROLE_SORT, _shown_revision_sort_key("—"))
            compare_item = _SortItem("—")
            compare_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            compare_item.setData(_ROLE_SORT, _compare_kind_sort_key(""))
            if isinstance(kit, tuple) and len(kit) == 2:
                auto_item.setData(_ROLE_KIT, kit)
                compare_item.setData(_ROLE_KIT, kit)
            export_rev = (
                selection.source_revision_text if selection is not None else ""
            )
            rd_path = (
                str(selection.source_path)
                if selection is not None and selection.source_path
                else ""
            )
            pinned = bool(
                selection is not None
                and selection.origin in {"pin", "pin_stale"}
            )
            cached = self._auto_mto_comparisons.get(key) if key is not None else None
            comparison = cached[1] if cached is not None else None
            comparison_rd_path = cached[0] if cached is not None else ""
            status = auto_mto_compare_status(
                files=files or (),
                rd_path=rd_path,
                rd_revision=export_rev,
                comparison=comparison,
                comparison_rd_path=comparison_rd_path,
                pinned=pinned,
            )
            auto_cell = _heatmap_auto_mto_cell(
                files or (), export_rev, rd_path, pinned, selection
            )
            apply_monitor_cell(auto_item, auto_cell, sort_role=_ROLE_SORT)
            auto_item.setData(
                _ROLE_SORT,
                _shown_revision_sort_key(str(auto_cell.sort_key or "—")),
            )
            self._paint_auto_mto_compare_item(compare_item, status)
            table.setItem(row, _COL_AUTO_MTO, auto_item)
            table.setItem(row, _COL_AUTO_MTO_COMPARE, compare_item)

    def _paint_auto_mto_compare_item(
        self,
        item: QTableWidgetItem,
        status: AutoMtoCompareStatus,
    ) -> None:
        cell = _compare_status_cell(status)
        apply_monitor_cell(item, cell, sort_role=_ROLE_SORT)
        item.setData(_ROLE_SORT, _compare_kind_sort_key(status.kind))

    def _cell_tooltip(self, cell: Any) -> str:
        lines = [
            f"{getattr(cell, 'title', '')}-{getattr(cell, 'mark', '')}",
            f"рев. {getattr(cell, 'revision_text', '') or '—'}",
            f"статус: {status_color_label(_cell_status(cell) or 'empty')}",
        ]
        letters = _cell_letters(cell)
        if letters:
            lines.append(f"буквы: {letters}")
        if _bool_attr(cell, "is_as_build"):
            lines.append("as-build")
        if _bool_attr(cell, "is_current"):
            lines.append(status_color_label("current"))
        if _bool_attr(cell, "is_current_ifc"):
            lines.append("последняя IFC (без as-build)")
        if _bool_attr(cell, "has_mto"):
            lines.append("есть MTO")
        else:
            lines.append("нет MTO")
        problems = _cell_problems(cell)
        if problems:
            lines.append("проблемы: " + ", ".join(problems))
        return "\n".join(lines)

    def _visible_cell(self, cell: Any | None) -> bool:
        if cell is None:
            return not (
                self._latest_only.isChecked()
                or self._code_a_only.isChecked()
                or self._tdo_only.isChecked()
                or self._problems_only.isChecked()
            )
        if self._no_as_build.isChecked() and _bool_attr(cell, "is_as_build"):
            return False
        if self._latest_only.isChecked():
            current_flag = (
                "is_current_ifc"
                if self._no_as_build.isChecked()
                else "is_current"
            )
            if not _bool_attr(cell, current_flag):
                return False
        if self._problems_only.isChecked() and not _cell_problems(cell):
            return False
        if self._code_a_only.isChecked() and _cell_status(cell) != "code_a":
            return False
        if self._tdo_only.isChecked() and _cell_status(cell) not in _TDO_STATUSES:
            return False
        return True

    def _row_matches_text(self, row: int, needle: str) -> bool:
        if not needle:
            return True
        table = self._table
        chunks: list[str] = []
        for column in range(table.columnCount()):
            item = table.item(row, column)
            if item is None:
                continue
            chunks.append(item.text())
            cell = item.data(_ROLE_CELL)
            if cell is not None:
                chunks.append(str(getattr(cell, "revision_text", "") or ""))
                chunks.extend(_cell_problems(cell))
        return needle in " ".join(chunks).casefold()

    @Slot()
    def _on_filter_text_changed(self) -> None:
        self._apply_row_visibility()
        self._update_export_button()

    @Slot()
    def _apply_row_visibility(self) -> None:
        needle = self._filter.text().strip().casefold()
        table = self._table
        has_status_filter = (
            self._latest_only.isChecked()
            or self._problems_only.isChecked()
            or self._code_a_only.isChecked()
            or self._tdo_only.isChecked()
            or self._no_as_build.isChecked()
        )
        for row in range(table.rowCount()):
            title_item = table.item(row, _COL_TITLE)
            kit = title_item.data(_ROLE_KIT) if title_item else None
            if isinstance(kit, tuple) and len(kit) == 2 and self._kit_is_hidden(*kit):
                table.setRowHidden(row, True)
                continue
            any_cell = False
            for column in range(_COL_REV_FIRST, table.columnCount()):
                item = table.item(row, column)
                if item is not None and item.data(_ROLE_CELL) is not None:
                    any_cell = True
                    break
            visible = self._row_matches_text(row, needle)
            if has_status_filter and not any_cell:
                visible = False
            table.setRowHidden(row, not visible)
        self._update_export_button()

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        if watched is self._table.viewport() and event.type() == QEvent.Type.Resize:
            self._apply_column_layout()
        return super().eventFilter(watched, event)

    def resizeEvent(self, event) -> None:  # noqa: ANN001
        super().resizeEvent(event)
        self._apply_column_layout()

    def showEvent(self, event) -> None:  # noqa: ANN001
        super().showEvent(event)
        self._apply_column_layout()

    def _apply_column_layout(self) -> None:
        """Fit identity columns plus equally stretched revision columns.

        Revision columns share leftover viewport width so the heatmap does not
        need a horizontal scrollbar. Six fixed columns stay
        user-resizable, then shrink if they would overflow the viewport.
        """

        if getattr(self, "_layout_guard", False):
            return
        table = self._table
        header = table.horizontalHeader()
        count = header.count()
        if count <= 0:
            return
        self._layout_guard = True
        try:
            table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            header.setMinimumSectionSize(16)
            header.setStretchLastSection(False)
            if count >= 1:
                header.setSectionResizeMode(
                    _COL_TITLE, QHeaderView.ResizeMode.Interactive
                )
            if count >= 2:
                header.setSectionResizeMode(
                    _COL_MARK, QHeaderView.ResizeMode.Interactive
                )
            if count >= 3:
                header.setSectionResizeMode(
                    _COL_EXPORT, QHeaderView.ResizeMode.Interactive
                )
            if count >= 4:
                header.setSectionResizeMode(
                    _COL_RD_REV, QHeaderView.ResizeMode.Interactive
                )
            if count >= 5:
                header.setSectionResizeMode(
                    _COL_AUTO_MTO, QHeaderView.ResizeMode.Interactive
                )
            if count >= 6:
                header.setSectionResizeMode(
                    _COL_AUTO_MTO_COMPARE, QHeaderView.ResizeMode.Interactive
                )
            for column in range(_COL_REV_FIRST, count):
                header.setSectionResizeMode(column, QHeaderView.ResizeMode.Stretch)
            n_rev = count - _COL_REV_FIRST
            if n_rev <= 0:
                return
            viewport = table.viewport().width()
            if viewport <= 0:
                return
            title_w = header.sectionSize(_COL_TITLE) if count >= 1 else 0
            mark_w = header.sectionSize(_COL_MARK) if count >= 2 else 0
            export_w = header.sectionSize(_COL_EXPORT) if count >= 3 else 0
            rd_w = header.sectionSize(_COL_RD_REV) if count >= 4 else 0
            auto_w = header.sectionSize(_COL_AUTO_MTO) if count >= 5 else 0
            compare_w = (
                header.sectionSize(_COL_AUTO_MTO_COMPARE) if count >= 6 else 0
            )
            max_fixed = max(viewport - 16 * n_rev, 80)
            fixed = title_w + mark_w + export_w + rd_w + auto_w + compare_w
            if fixed <= max_fixed or fixed <= 0:
                return
            scale = max_fixed / fixed
            if count >= 1:
                header.resizeSection(_COL_TITLE, max(48, int(title_w * scale)))
            if count >= 2:
                header.resizeSection(_COL_MARK, max(48, int(mark_w * scale)))
            if count >= 3:
                header.resizeSection(_COL_EXPORT, max(72, int(export_w * scale)))
            if count >= 4:
                header.resizeSection(_COL_RD_REV, max(56, int(rd_w * scale)))
            if count >= 5:
                header.resizeSection(_COL_AUTO_MTO, max(64, int(auto_w * scale)))
            if count >= 6:
                header.resizeSection(
                    _COL_AUTO_MTO_COMPARE, max(72, int(compare_w * scale))
                )
        finally:
            self._layout_guard = False

    def _kit_at_row(self, row: int) -> tuple[str, str] | None:
        item = self._table.item(row, _COL_TITLE)
        kit = item.data(_ROLE_KIT) if item else None
        if isinstance(kit, tuple) and len(kit) == 2:
            return str(kit[0]), str(kit[1])
        return None

    def _add_show_kit_action(
        self, menu: QMenu, title: str, mark: str
    ) -> QAction:
        """Add and wire the context action that opens «Комплекты»."""

        action = menu.addAction('Показать в „Комплекты“')
        action.triggered.connect(
            lambda _checked=False: self._emit_kit_activated(title, mark)
        )
        return action

    def _emit_kit_activated(self, title: str, mark: str) -> None:
        """Request opening one title–mark on the kits tab."""

        self.kit_activated.emit(title, mark)

    def _show_pin_context_menu(self, position) -> None:
        """Offer the kits jump, AutoMTO compare, and pin actions."""

        row = self._table.rowAt(position.y())
        kit = self._kit_at_row(row) if row >= 0 else None
        if kit is None:
            return
        title, mark = kit
        key = kit_identity_key(title, mark)
        selection = self.selection_for(title, mark)
        files = self._auto_mto_by_kit.get(key, ())
        existing = self.pin_for(title, mark)
        menu = QMenu(self)
        self._add_show_kit_action(menu, title, mark)
        menu.addSeparator()
        compare_action = menu.addAction(_AUTO_MTO_COMPARE_ACTION_LABEL)
        compare_action.setEnabled(
            bool(selection and selection.source_path and files)
        )
        menu.addSeparator()
        pin_action = menu.addAction(_PIN_ACTION_LABEL)
        unpin_action = menu.addAction(_UNPIN_ACTION_LABEL)
        unpin_action.setEnabled(existing is not None)
        chosen = exec_tracked_menu(
            menu, MENU_HEATMAP, self._table.viewport().mapToGlobal(position)
        )
        if chosen == compare_action:
            self.start_auto_mto_compare(title, mark)
        elif chosen == pin_action:
            self._choose_export_pin(title, mark, existing=existing)
        elif chosen == unpin_action:
            self.clear_export_pin(title, mark)

    def start_auto_mto_compare(
        self,
        title: str,
        mark: str,
        rd_path: str = "",
    ) -> bool:
        """Start or queue exact matching for one RD MTO workbook.

        If a compare is already running, the job is prepended to the
        queue. If the queue is paused, the job is queued and not started.

        Args:
            title: Four-digit title.
            mark: Latin AGCC mark.
            rd_path: RD MTO to compare. Empty uses the export source file.

        Returns:
            True when the worker started or the job was queued.
        """

        files = self._auto_mto_by_kit.get(kit_identity_key(title, mark), ())
        path = str(rd_path or "").strip()
        if not path:
            selection = self.selection_for(title, mark)
            path = str(selection.source_path) if selection is not None else ""
        if not path or not files:
            return False
        key = kit_identity_key(title, mark)
        if self.is_auto_mto_compare_running() or self._auto_mto_queue_paused:
            self._enqueue_auto_mto_job(key, path, prepend=True)
            self._update_auto_mto_queue_label()
            return True
        if self._start_auto_mto_job(key, path):
            self._auto_mto_queue_total += 1
            self._update_auto_mto_queue_label()
            return True
        return False

    @Slot(object)
    def _on_auto_mto_compare_ready(self, result: object) -> None:
        context = self._auto_mto_compare_context
        if context is None or not isinstance(result, AutoMtoCompareResult):
            return
        key, rd_path = context
        self.apply_auto_mto_compare_result(key[0], key[1], rd_path, result)
        if result.match_kind == "single":
            grade = (
                "ПоКоду и Кол-ву"
                if result.content_grade == "soft"
                else "четкое"
            )
            summary = f"{grade}, 1 файл, строк {result.rd_rows}"
        elif result.match_kind == "composite":
            grade = (
                "ПоКоду и Кол-ву"
                if result.content_grade == "soft"
                else "четкое"
            )
            summary = (
                f"{grade}, сумма {result.cell_text()}, строк {result.rd_rows}"
            )
        elif result.match_kind == "no_match":
            summary = "не совпало по коду и количеству"
        else:
            summary = f"не выполнена: {result.error or 'нет данных'}"
        self.pair_compare_log.emit(f"Авто МТО: {key[0]}-{key[1]} — {summary}.")

    @Slot(str)
    def _on_auto_mto_compare_error(self, message: str) -> None:
        self.pair_compare_log.emit(f"Авто МТО: ошибка сверки: {message}")

    @Slot()
    def _on_auto_mto_compare_finished(self) -> None:
        thread = self._auto_mto_compare_thread
        self._auto_mto_compare_thread = None
        self._auto_mto_compare_context = None
        if thread is not None:
            thread.deleteLater()
        self._auto_mto_queue_done += 1
        self._kick_auto_mto_queue()
        if not self._auto_mto_queue and not self.is_auto_mto_compare_running():
            self.auto_mto_compare_finished.emit()

    def _has_matching_auto_mto_comparison(
        self,
        key: tuple[str, str],
        rd_path: str,
    ) -> bool:
        cached = self._auto_mto_comparisons.get(key)
        if cached is None:
            return False
        stored_path, result = cached
        if not auto_mto_rd_paths_match(stored_path, rd_path):
            return False
        return result.match_kind in _PERSISTABLE_COMPARE_KINDS

    def _auto_mto_job_queued(self, key: tuple[str, str], rd_path: str) -> bool:
        return any(
            queued_key == key and auto_mto_rd_paths_match(queued_path, rd_path)
            for queued_key, queued_path in self._auto_mto_queue
        )

    def _drop_other_auto_mto_jobs(
        self,
        key: tuple[str, str],
        keep_path: str,
    ) -> None:
        """Remove queued jobs for ``key`` that target a different RD file."""

        self._auto_mto_queue = [
            item
            for item in self._auto_mto_queue
            if item[0] != key or auto_mto_rd_paths_match(item[1], keep_path)
        ]

    def _enqueue_auto_mto_job(
        self,
        key: tuple[str, str],
        rd_path: str,
        *,
        prepend: bool,
    ) -> None:
        path = str(rd_path)
        self._drop_other_auto_mto_jobs(key, path)
        if self._auto_mto_job_queued(key, path):
            if prepend:
                self._auto_mto_queue = [
                    item
                    for item in self._auto_mto_queue
                    if not (
                        item[0] == key
                        and auto_mto_rd_paths_match(item[1], path)
                    )
                ]
                self._auto_mto_queue.insert(0, (key, path))
            return
        if prepend:
            self._auto_mto_queue.insert(0, (key, path))
        else:
            self._auto_mto_queue.append((key, path))
        self._auto_mto_queue_total += 1

    def _start_auto_mto_job(self, key: tuple[str, str], rd_path: str) -> bool:
        files = self._auto_mto_by_kit.get(key, ())
        path = str(rd_path or "").strip()
        if not path or not files:
            return False
        thread = AutoMtoCompareThread(files, path, parent=self)
        self._auto_mto_compare_thread = thread
        self._auto_mto_compare_context = (key, path)
        thread.result_ready.connect(self._on_auto_mto_compare_ready)
        thread.error.connect(self._on_auto_mto_compare_error)
        thread.finished.connect(self._on_auto_mto_compare_finished)
        title, mark = files[0].title, files[0].mark
        self.pair_compare_log.emit(
            f"Авто МТО: сверка {title}-{mark} с {Path(path).name}…"
        )
        thread.start()
        return True

    def _kick_auto_mto_queue(self) -> None:
        if self._auto_mto_queue_paused:
            self._update_auto_mto_queue_label()
            return
        if self.is_auto_mto_compare_running():
            self._update_auto_mto_queue_label()
            return
        while self._auto_mto_queue:
            key, rd_path = self._auto_mto_queue.pop(0)
            if self._has_matching_auto_mto_comparison(key, rd_path):
                continue
            if self._start_auto_mto_job(key, rd_path):
                break
        self._reset_auto_mto_queue_counts_if_idle()
        self._update_auto_mto_queue_label()

    def _reset_auto_mto_queue_counts_if_idle(self) -> None:
        if self._auto_mto_queue or self.is_auto_mto_compare_running():
            return
        self._auto_mto_queue_done = 0
        self._auto_mto_queue_total = 0

    def _update_auto_mto_queue_label(self) -> None:
        if not hasattr(self, "_auto_mto_queue_label"):
            return
        running = 1 if self.is_auto_mto_compare_running() else 0
        queued = len(self._auto_mto_queue)
        if running == 0 and queued == 0:
            self._auto_mto_queue_label.setText("Авто МТО сверка: актуально")
            return
        done = self._auto_mto_queue_done
        total = max(self._auto_mto_queue_total, done + running + queued)
        self._auto_mto_queue_label.setText(f"Авто МТО сверка: {done}/{total}")

    def _selection_rd_path(self, key: tuple[str, str]) -> str:
        selection = self._selections_by_key.get(key)
        if selection is None or not selection.source_path:
            return ""
        return str(selection.source_path)

    def _rd_mtime_ns(self, rd_path: str) -> int:
        """Return RD mtime from the last scan snapshot, else a path stat."""

        if not rd_path:
            return 0
        cached = self._mtime_by_path.get(rd_path.casefold())
        if cached is not None:
            return int(cached)
        needle = auto_mto_rd_path_key(rd_path)
        if needle:
            for stored, mtime in self._mtime_by_path.items():
                if auto_mto_rd_path_key(stored) == needle:
                    return int(mtime)
        return rd_mtime_ns(rd_path)

    def _in_memory_comparison_current(
        self,
        key: tuple[str, str],
        result: AutoMtoCompareResult,
        files: Sequence[AutoMtoFile],
    ) -> bool:
        if result.match_kind not in _PERSISTABLE_COMPARE_KINDS:
            return False
        stored_sig = self._auto_mto_comparison_sigs.get(key)
        if stored_sig is not None and stored_sig != _auto_mto_files_sig(files):
            return False
        rebuilt = result_from_entry(
            {
                "match_kind": result.match_kind,
                "member_relpaths": [member.relpath for member in result.members],
                "rd_rows": result.rd_rows,
                "auto_rows": result.auto_rows,
                "combinations_checked": result.combinations_checked,
                "error": result.error,
                "grade": result.content_grade,
            },
            files,
        )
        return rebuilt is not None

    def _disk_result_for_rd_path(
        self,
        kit: tuple[str, str],
        rd_path: str,
        files: Sequence[AutoMtoFile],
        disk: Mapping[str, Mapping[str, Any]],
    ) -> AutoMtoCompareResult | None:
        """Return a persistable disk result for ``rd_path``, if still valid."""

        if not rd_path or not files:
            return None
        title, mark = kit
        mtime = self._rd_mtime_ns(rd_path)
        exact = disk.get(cache_key(title, mark, rd_path, mtime, files))
        rebuilt = result_from_entry(exact, files) if exact else None
        if rebuilt is not None:
            return rebuilt
        for stored_key, entry in disk.items():
            stored_path = str(entry.get("rd_path") or "")
            entry_kit = kit_identity_key(
                str(entry.get("title") or ""),
                str(entry.get("mark") or ""),
            )
            if entry_kit != kit or not auto_mto_rd_paths_match(stored_path, rd_path):
                continue
            stored_mtime = self._rd_mtime_ns(stored_path)
            for candidate_path, candidate_mtime in (
                (rd_path, mtime),
                (stored_path, stored_mtime),
                (stored_path, mtime),
                (rd_path, stored_mtime),
            ):
                if stored_key == cache_key(
                    title, mark, candidate_path, candidate_mtime, files
                ):
                    return result_from_entry(entry, files)
        return None

    def _prune_and_hydrate_auto_mto_comparisons(self) -> None:
        files_by_kit = self._auto_mto_by_kit
        kept: dict[tuple[str, str], tuple[str, AutoMtoCompareResult]] = {}
        kept_sigs: dict[tuple[str, str], tuple[tuple[str, str], ...]] = {}
        disk: dict[str, dict[str, Any]] = {}
        if self._config is not None:
            disk = load_auto_mto_compare_cache(self._config.runtime_dir)

        def _remember(
            kit: tuple[str, str],
            rd_path: str,
            result: AutoMtoCompareResult,
            *,
            sig: tuple[tuple[str, str], ...] | None = None,
        ) -> None:
            kept[kit] = (rd_path, result)
            files = files_by_kit.get(kit, ())
            kept_sigs[kit] = sig or _auto_mto_files_sig(files)

        for key, (rd_path, result) in self._auto_mto_comparisons.items():
            files = files_by_kit.get(key)
            if not files:
                continue
            current = self._selection_rd_path(key)
            if current and not auto_mto_rd_paths_match(rd_path, current):
                continue
            if not self._in_memory_comparison_current(key, result, files):
                continue
            _remember(
                key,
                current or rd_path,
                result,
                sig=self._auto_mto_comparison_sigs.get(key),
            )
        for kit, files in files_by_kit.items():
            if kit in kept or not files:
                continue
            current = self._selection_rd_path(kit)
            if not current:
                continue
            rebuilt = self._disk_result_for_rd_path(kit, current, files, disk)
            if rebuilt is not None:
                _remember(kit, current, rebuilt)
        for key, (rd_path, result) in self._auto_mto_comparisons.items():
            if key in kept:
                continue
            files = files_by_kit.get(key)
            if not files:
                continue
            if not self._in_memory_comparison_current(key, result, files):
                continue
            _remember(
                key,
                rd_path,
                result,
                sig=self._auto_mto_comparison_sigs.get(key),
            )
        for stored_key, entry in disk.items():
            title = str(entry.get("title") or "")
            mark = str(entry.get("mark") or "")
            rd_path = str(entry.get("rd_path") or "")
            if not title or not mark or not rd_path:
                continue
            kit = kit_identity_key(title, mark)
            if kit in kept:
                continue
            files = files_by_kit.get(kit)
            if not files:
                continue
            current_key = cache_key(
                title, mark, rd_path, self._rd_mtime_ns(rd_path), files
            )
            if stored_key != current_key:
                continue
            rebuilt = result_from_entry(entry, files)
            if rebuilt is None:
                continue
            _remember(kit, rd_path, rebuilt)
        self._auto_mto_comparisons = kept
        self._auto_mto_comparison_sigs = kept_sigs

    def _persist_auto_mto_compare_result(
        self,
        title: str,
        mark: str,
        rd_path: str,
        result: AutoMtoCompareResult,
    ) -> None:
        entry = entry_from_result(
            result, title=title, mark=mark, rd_path=rd_path
        )
        if entry is None or self._config is None:
            return
        key = kit_identity_key(title, mark)
        files = self._auto_mto_by_kit.get(key, ())
        entries = load_auto_mto_compare_cache(self._config.runtime_dir)
        entries[
            cache_key(title, mark, rd_path, self._rd_mtime_ns(rd_path), files)
        ] = entry
        try:
            save_auto_mto_compare_cache(self._config.runtime_dir, entries)
        except OSError as exc:
            self.pair_compare_log.emit(
                f"Авто МТО: не записан кэш сверки: {type(exc).__name__}: {exc}"
            )

    def _choose_export_pin(
        self,
        title: str,
        mark: str,
        *,
        existing: ExportPin | None,
    ) -> None:
        """Open the candidate dialog and optionally confirm an override."""

        if self._config is None:
            return
        selection = self.selection_for(title, mark)
        rule_path = selection.rule_path if selection is not None else ""
        candidates = list_export_pin_candidates(
            self._records,
            title=title,
            mark=mark,
            rd_root=self._config.rd_root,
            rule_path=rule_path,
        )
        if not candidates:
            QMessageBox.information(self, _PIN_DIALOG_TITLE, _PIN_EMPTY_TEXT)
            return
        dialog = MtoPinPickDialog(
            candidates,
            current_path=existing.file_path if existing is not None else "",
            parent=self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        candidate = dialog.selected_candidate()
        if candidate is None:
            return
        if existing is not None:
            reply = QMessageBox.question(
                self,
                _PIN_CONFIRM_TITLE,
                _PIN_CONFIRM_TEXT,
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return
        self.assign_export_pin(title, mark, candidate)

    def _visible_row_count(self) -> int:
        table = self._table
        return sum(
            1 for row in range(table.rowCount()) if not table.isRowHidden(row)
        )

    def _update_export_button(self) -> None:
        if not hasattr(self, "_export_button"):
            return
        visible = self._visible_row_count()
        self._export_button.setText(f"Копировать видимые ({visible})")
        pending = self._pool_status.pending
        blocked = pending > 0 or self.is_copy_running()
        self._export_button.setEnabled(not blocked)
        if pending > 0:
            self._export_button.setToolTip(_EXPORT_GATE_TOOLTIP)
        elif self.is_copy_running():
            self._export_button.setToolTip("Копирование уже выполняется.")
        else:
            self._export_button.setToolTip(_EXPORT_READY_TOOLTIP)

    def _reload_targets(self, keep_name: str | None) -> None:
        if self._config is None:
            return
        self._targets = list(load_export_targets(self._config))
        self._updating_controls = True
        try:
            self._target_combo.clear()
            for target in self._targets:
                self._target_combo.addItem(target.name, target.name)
            name = keep_name or (self._targets[0].name if self._targets else "")
            self._select_target_name(name)
        finally:
            self._updating_controls = False
        self._apply_target_memory()
        self._update_target_controls()

    def _select_target_name(self, name: str) -> None:
        needle = name.casefold()
        self._updating_controls = True
        try:
            for index, target in enumerate(self._targets):
                if target.name.casefold() == needle:
                    self._target_combo.setCurrentIndex(index)
                    self._last_target_index = index
                    break
        finally:
            self._updating_controls = False
        self._apply_target_memory()
        self._update_target_controls()

    def _apply_target_memory(self) -> None:
        target = self.current_target()
        if target is None:
            return
        self._updating_controls = True
        try:
            rule_index = self._rule_combo.findData(target.rule)
            if rule_index < 0:
                rule_index = self._rule_combo.findData(DEFAULT_EXPORT_RULE)
            if rule_index >= 0:
                self._rule_combo.setCurrentIndex(rule_index)
            self._filter.blockSignals(True)
            self._filter.setText(target.filter_text)
            self._filter.blockSignals(False)
            self._flat_box.blockSignals(True)
            self._flat_box.setChecked(bool(target.flat_structure))
            self._flat_box.blockSignals(False)
        finally:
            self._updating_controls = False
        self._apply_row_visibility()

    def _update_target_controls(self) -> None:
        target = self.current_target()
        is_default = bool(target is not None and target.is_default_robot)
        self._remove_target_button.setEnabled(target is not None and not is_default)
        self._folder_button.setEnabled(target is not None and not is_default)
        self._flat_box.setEnabled(target is not None and not is_default)
        self._rename_target_button.setEnabled(target is not None)

    def _persist_current_target_fields(self) -> None:
        if self._config is None:
            return
        target = self.current_target()
        if target is None:
            return
        index = self._target_combo.currentIndex()
        updated = replace(
            target,
            rule=self.current_rule(),
            filter_text=self._filter.text(),
            flat_structure=(
                target.flat_structure
                if target.is_default_robot
                else self._flat_box.isChecked()
            ),
        )
        if 0 <= index < len(self._targets):
            self._targets[index] = updated
        self._targets = list(save_export_targets(self._config, self._targets))

    def _refresh_target_files(self) -> None:
        target = self.current_target()
        if target is None:
            self._target_files = {}
            return
        if target.is_default_robot:
            self._target_files = target_files_from_records(self._records)
            return
        self._target_files = scan_export_target(target)

    def _merged_session(self) -> dict[tuple[str, str], str]:
        merged = dict(self._session_verdicts)
        thread = self._pair_thread
        if thread is not None and thread.session_verdicts:
            merged.update(session_status_map(thread.session_verdicts))
        return merged

    def _refresh_export(self, *, restart_compare: bool) -> None:
        if self._config is None or self._database is None:
            return
        target = self.current_target()
        if target is None:
            return
        self._refresh_target_files()
        pins = list(load_export_pins(self._config))
        refreshed: list[ExportPin] = []
        try:
            selections = resolve_export_selections(
                self._database,
                records=self._records,
                detected_current_ids=self._detected_current_ids,
                rule=self.current_rule(),
                target=target,
                target_files=self._target_files,
                pins=pins,
                rd_root=self._config.rd_root,
                refreshed_pins=refreshed,
            )
        except ValueError:
            return
        if refreshed:
            by_key = {
                kit_identity_key(item.title, item.mark): item for item in pins
            }
            for pin in refreshed:
                by_key[kit_identity_key(pin.title, pin.mark)] = pin
            pins = list(save_export_pins(self._config, tuple(by_key.values())))
        pairs = pairs_from_export_selections(selections)
        file_verdicts = file_id_verdicts_from_db(
            self._database, records=self._records, pairs=pairs
        )
        if file_verdicts:
            selections = resolve_export_selections(
                self._database,
                records=self._records,
                detected_current_ids=self._detected_current_ids,
                rule=self.current_rule(),
                target=target,
                target_files=self._target_files,
                pins=pins,
                rd_root=self._config.rd_root,
                verdicts=file_verdicts,
            )
        session = self._merged_session()
        selections = overlay_session_on_selections(selections, session)
        self.apply_export_selections(selections)
        self._recompute_pool_status()
        if restart_compare:
            self._ensure_pair_compare()

    def _recompute_pool_status(self) -> None:
        if self._database is None:
            return
        status = pair_pool_status(
            self._database,
            pairs=self._current_pairs,
            records=self._records,
            session_verdicts=self._merged_session(),
        )
        self.set_pool_status(status)

    def _patch_export_from_store(self) -> None:
        if self._database is None:
            return
        file_verdicts = file_id_verdicts_from_db(
            self._database, records=self._records, pairs=self._current_pairs
        )
        session = self._merged_session()
        patched = patch_selections_with_verdicts(
            self._selections,
            records=self._records,
            file_id_verdicts=file_verdicts,
            session_verdicts=session,
        )
        self.apply_export_selections(patched)
        self._recompute_pool_status()

    def _ensure_pair_compare(self) -> None:
        if self._config is None or self._database is None:
            return
        if self.is_copy_running():
            return
        if self._pool_status.pending <= 0:
            self._pair_queued_pairs = None
            return
        pairs = self._current_pairs
        wanted = pair_identity_key(pairs)
        thread = self._pair_thread
        if thread is not None and thread.isRunning():
            if wanted == self._pair_thread_key:
                return
            self._pair_queued_pairs = pairs
            thread.request_cancel()
            return
        self._start_pair_compare(pairs)

    def _start_pair_compare(self, pairs: Sequence[MtoFilePair]) -> None:
        if self._config is None or not pairs:
            return
        if self._pool_status.pending <= 0:
            return
        thread = MtoPairCompareThread(
            self._config, pairs=pairs, batch_size=1, parent=self
        )
        thread.progress.connect(self._on_pair_progress)
        thread.log.connect(self.pair_compare_log)
        thread.error.connect(self.pair_compare_log)
        thread.finished.connect(self._on_pair_finished)
        self._pair_thread = thread
        self._pair_thread_key = pair_identity_key(pairs)
        self._pair_queued_pairs = None
        thread.start()

    @Slot(int, int)
    def _on_pair_progress(self, _completed: int, _total: int) -> None:
        self._patch_export_from_store()

    @Slot()
    def _on_pair_finished(self) -> None:
        thread = self._pair_thread
        self._pair_thread = None
        queued = self._pair_queued_pairs
        self._pair_queued_pairs = None
        if thread is not None:
            if thread.session_verdicts:
                self._session_verdicts.update(
                    session_status_map(thread.session_verdicts)
                )
            thread.deleteLater()
        self._patch_export_from_store()
        if queued:
            self._current_pairs = tuple(queued)
            self._recompute_pool_status()
            self._start_pair_compare(queued)
            return
        self._ensure_pair_compare()

    @Slot()
    def _on_rule_changed(self) -> None:
        if self._updating_controls:
            return
        self._session_verdicts.clear()
        self._persist_current_target_fields()
        self._refresh_export(restart_compare=True)

    @Slot()
    def _on_target_changed(self) -> None:
        if self._updating_controls:
            return
        previous_index = self._last_target_index
        if 0 <= previous_index < len(self._targets) and self._config is not None:
            previous = self._targets[previous_index]
            self._targets[previous_index] = replace(
                previous,
                rule=self.current_rule(),
                filter_text=self._filter.text(),
            )
            save_export_targets(self._config, self._targets)
        self._last_target_index = self._target_combo.currentIndex()
        self._session_verdicts.clear()
        self._apply_target_memory()
        self._update_target_controls()
        self._refresh_export(restart_compare=True)

    @Slot()
    def _on_flat_toggled(self) -> None:
        if self._updating_controls:
            return
        self._persist_current_target_fields()
        self._session_verdicts.clear()
        self._refresh_export(restart_compare=True)

    @Slot()
    def _on_add_target(self) -> None:
        if self._config is None:
            return
        name, ok = QInputDialog.getText(self, "Новая папка экспорта", "Имя:")
        if not ok:
            return
        name = name.strip()
        if not name:
            return
        folder = QFileDialog.getExistingDirectory(self, "Папка назначения")
        if not folder:
            return
        new_target = ExportTarget(
            name=name,
            root=folder,
            flat_structure=self._flat_box.isChecked(),
            is_default_robot=False,
            rule=self.current_rule(),
            filter_text=self._filter.text(),
        )
        self._persist_current_target_fields()
        self._targets.append(new_target)
        saved = save_export_targets(self._config, self._targets)
        self._targets = list(saved)
        self._reload_targets(keep_name=name)
        self._session_verdicts.clear()
        self._refresh_export(restart_compare=True)

    @Slot()
    def _on_rename_target(self) -> None:
        if self._config is None:
            return
        target = self.current_target()
        if target is None:
            return
        name, ok = QInputDialog.getText(
            self, "Переименовать папку", "Имя:", text=target.name
        )
        if not ok:
            return
        name = name.strip()
        if not name:
            return
        index = self._target_combo.currentIndex()
        self._targets[index] = replace(target, name=name)
        self._targets = list(save_export_targets(self._config, self._targets))
        self._reload_targets(keep_name=name)

    @Slot()
    def _on_remove_target(self) -> None:
        if self._config is None:
            return
        target = self.current_target()
        if target is None or target.is_default_robot:
            return
        reply = QMessageBox.question(
            self,
            "Удалить папку",
            f"Удалить «{target.name}» из списка? Файлы на диске не трогаем.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self._targets = [item for item in self._targets if item.name != target.name]
        saved = save_export_targets(self._config, self._targets)
        self._targets = list(saved)
        default = default_robot_export_target(self._config)
        self._reload_targets(keep_name=default.name)
        self._session_verdicts.clear()
        self._refresh_export(restart_compare=True)

    @Slot()
    def _on_choose_folder(self) -> None:
        if self._config is None:
            return
        target = self.current_target()
        if target is None or target.is_default_robot:
            return
        folder = QFileDialog.getExistingDirectory(
            self, "Папка назначения", target.root or ""
        )
        if not folder:
            return
        index = self._target_combo.currentIndex()
        updated = replace(target, root=folder)
        self._targets[index] = updated
        self._targets = list(save_export_targets(self._config, self._targets))
        self._reload_targets(keep_name=updated.name)
        self._session_verdicts.clear()
        self._refresh_export(restart_compare=True)

    @Slot()
    def _on_export_clicked(self) -> None:
        if self._pool_status.pending > 0:
            QMessageBox.information(self, "Копирование MTO", _EXPORT_GATE_TOOLTIP)
            return
        if self.is_copy_running():
            return
        target = self.current_target()
        if target is None:
            QMessageBox.warning(self, "Копирование MTO", "Не выбрана папка назначения.")
            return
        if not str(target.root or "").strip():
            QMessageBox.warning(self, "Копирование MTO", "У папки назначения нет пути.")
            return
        visible = self.visible_export_selections()
        preview = MtoExportPreviewDialog(visible, target, self)
        if preview.exec() != QDialog.DialogCode.Accepted:
            return
        if preview_copy_count(visible) == 0:
            return
        try:
            plan = plan_for_visible_rows(visible, target=target)
        except ValueError as exc:
            QMessageBox.warning(self, "Копирование MTO", str(exc))
            return
        if not plan.items:
            QMessageBox.information(
                self,
                "Копирование MTO",
                "Нет файлов для копирования (все строки пропущены).",
            )
            return
        self.cancel_pair_compare()
        thread = MtoExportCopyThread(plan, self, now=datetime.now())
        thread.log.connect(self.pair_compare_log)
        dialog = MtoExportProgressDialog(thread, self)
        self._copy_thread = thread
        self._copy_dialog = dialog
        self._update_export_button()
        dialog.start_copy()
        dialog.exec()
        self._on_copy_finished()

    def _on_copy_finished(self) -> None:
        thread = self._copy_thread
        dialog = self._copy_dialog
        self._copy_thread = None
        self._copy_dialog = None
        report = dialog.report if dialog is not None else None
        failure = thread.failure if thread is not None else None
        target = thread.plan.target if thread is not None else None
        if thread is not None:
            thread.deleteLater()
        self._update_export_button()
        if failure:
            QMessageBox.warning(self, "Копирование MTO", failure)
            return
        if report is not None:
            QMessageBox.information(
                self,
                "Копирование MTO",
                format_export_copy_report(report),
            )
        if target is not None:
            self._session_verdicts.clear()
            self._refresh_export(restart_compare=True)
            self.export_copy_finished.emit(
                str(target.root), bool(target.is_default_robot)
            )


def _paint_fill(item: QTableWidgetItem, hex_color: str) -> None:
    color = QColor(hex_color)
    if not color.isValid():
        return
    item.setBackground(QBrush(color))
    item.setForeground(
        QBrush(QColor("#ffffff") if color.lightness() < 140 else QColor("#202124"))
    )


def _settings_bool(settings: QSettings, key: str, default: bool) -> bool:
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
