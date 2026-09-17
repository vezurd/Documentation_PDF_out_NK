"""Worklist tab: one MTO row per heatmap cell with filters.

Qt monitor only. Rows come from ``list_mto_worklist``.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from dataclasses import replace
from datetime import datetime

from PySide6.QtCore import QPoint, QSettings, Qt, Signal, Slot
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.context_menu_qt import exec_tracked_menu
from rd_catalog.context_menu_usage import MENU_MTO_WORKLIST
from rd_catalog.customer_pi_auto_mto import (
    AUTO_MTO_COMPARE_STATUS_HEADER,
    AUTO_MTO_COMPARE_STATUS_TOOLTIP,
    AutoMtoCompareStatus,
)
from rd_catalog.kits import kit_identity_key
from rd_catalog.models import collision_kind_label
from rd_catalog.mto_export import PIN_COLUMN_HEADER, ExportPinView, export_pin_view
from rd_catalog.monitor_qt import apply_monitor_cell
from rd_catalog.monitor_views import (
    GAP_LABELS,
    TDO_STATUSES as _TDO_STATUSES,
    worklist_google_cell,
    worklist_link_text,
)
from rd_catalog.path_actions import open_path
from rd_catalog.pipeline import MtoWorklistRow
from rd_catalog.status_colors import color_for, status_color_label, status_short_label

_ROLE_ROW = Qt.ItemDataRole.UserRole
_ROLE_SORT = Qt.ItemDataRole.UserRole + 1
_HEADERS = (
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
_COL_STATUS = 3
_COL_LINK = 5
_COL_GOOGLE = 6
_COL_PIN = 8
_COL_COMPARE = 9
_COL_FILE = 10
_COL_DATE = 11
_COL_PACKAGE = 12
_COL_GAP = 13
_COL_PROBLEMS = 14
_COL_PATH = 15
_EMPTY_COMPARE_STATUS = AutoMtoCompareStatus(kind="empty", text="—", tooltip="")
_REV_MATCH_BG = QColor("#E2F2E1")
_REV_DIFF_BG = QColor("#F7E8BE")
_LEGEND_KEYS = (
    "code_a",
    "code_b",
    "code_c",
    "tdo_review",
    "sent_tdo",
    "not_uploaded",
    "no_mto",
    "working",
    "problem",
)
_COUNTS_SCOPE_EMPTY = (
    "Титулы: 0 / 0 · Комплекты: 0 / 0 · Строки: 0 / 0"
)
_COUNTS_BREAKDOWN_EMPTY = (
    "Код A: 0 · ТДО: 0 · Нет MTO: 0 · Нет в РД: 0 · "
    "Нет в Google: 0 · Без статуса F: 0 · Проблемы: 0"
)


class _SortItem(QTableWidgetItem):
    """Table item that sorts by ``_ROLE_SORT`` when both sides have it."""

    def __lt__(self, other: QTableWidgetItem) -> bool:
        left = self.data(_ROLE_SORT)
        right = other.data(_ROLE_SORT) if other is not None else None
        if left is not None and right is not None:
            return bool(left < right)
        return super().__lt__(other)


class MtoWorklistTab(QWidget):
    """Filterable MTO file list: one row per heatmap revision cell."""

    kit_activated = Signal(str, str)
    kit_documents_requested = Signal(str, str)
    prepare_context_menu = Signal(QMenu)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._palette: dict[str, str] = {}
        self._rows: tuple[MtoWorklistRow, ...] = ()
        self._is_banned: Callable[[str, str], bool] = (
            lambda _t, _m: False
        )
        self._allowed_kits: set[tuple[str, str]] | None = None
        self._pin_view_for: Callable[[str, str], ExportPinView] = (
            lambda _title, _mark: export_pin_view(None)
        )
        self._compare_status_for: Callable[[str, str], AutoMtoCompareStatus] = (
            lambda _title, _mark: _EMPTY_COMPARE_STATUS
        )
        self._total_titles = 0
        self._total_kits = 0
        self._total_rows = 0
        self._build()

    def table(self) -> QTableWidget:
        return self._table

    def selected_row(self) -> MtoWorklistRow | None:
        """Return the worklist payload for the current table selection.

        Returns:
            The ``MtoWorklistRow`` stored on the selected row, or
            ``None`` when nothing is selected.
        """

        return self._row_at(self._table.currentRow())

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        self._counts_scope = QLabel(_COUNTS_SCOPE_EMPTY, self)
        scope_font = QFont(self._counts_scope.font())
        scope_font.setBold(True)
        self._counts_scope.setFont(scope_font)
        self._counts_breakdown = QLabel(_COUNTS_BREAKDOWN_EMPTY, self)
        layout.addWidget(self._counts_scope)
        layout.addWidget(self._counts_breakdown)

        status_filters = QHBoxLayout()
        status_filters.addWidget(QLabel("Фильтр:"))
        self._filter = QLineEdit(self)
        self._filter.setPlaceholderText(
            "Титул, марка, ревизия, путь…"
        )
        self._filter.textChanged.connect(self._apply_row_visibility)
        status_filters.addWidget(self._filter, 1)
        self._code_a = QCheckBox("Код A", self)
        self._tdo = QCheckBox("Прошли ТДО", self)
        self._current = QCheckBox("Только текущие", self)
        self._current_ifc = QCheckBox("Только последние IFC", self)
        for box in (
            self._code_a,
            self._tdo,
            self._current,
            self._current_ifc,
        ):
            box.toggled.connect(self._apply_row_visibility)
            status_filters.addWidget(box)
        layout.addLayout(status_filters)

        gap_filters = QHBoxLayout()
        self._missing = QCheckBox("Нет MTO", self)
        self._no_rd = QCheckBox("Нет в РД", self)
        self._no_google = QCheckBox("Нет в Google", self)
        self._no_f_status = QCheckBox("Без статуса F", self)
        self._problems = QCheckBox("Только проблемы", self)
        self._no_as_build = QCheckBox("Без as-build", self)
        self._only_as_build = QCheckBox("Только as-build", self)
        for box in (
            self._missing,
            self._no_rd,
            self._no_google,
            self._no_f_status,
            self._problems,
        ):
            box.toggled.connect(self._apply_row_visibility)
            gap_filters.addWidget(box)
        self._no_as_build.toggled.connect(self._on_as_build_toggled)
        self._only_as_build.toggled.connect(self._on_as_build_toggled)
        gap_filters.addWidget(self._no_as_build)
        gap_filters.addWidget(self._only_as_build)
        gap_filters.addStretch(1)
        layout.addLayout(gap_filters)

        self._legend = QHBoxLayout()
        layout.addLayout(self._legend)

        self._table = QTableWidget(0, len(_HEADERS), self)
        self._table.setHorizontalHeaderLabels(list(_HEADERS))
        compare_header = self._table.horizontalHeaderItem(_COL_COMPARE)
        if compare_header is not None:
            compare_header.setToolTip(AUTO_MTO_COMPARE_STATUS_TOOLTIP)
        self._table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self._table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self._table.setEditTriggers(
            QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self._table.setAlternatingRowColors(True)
        self._table.setSortingEnabled(True)
        self._table.verticalHeader().setVisible(False)
        self._table.setWordWrap(False)
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setSectionsMovable(True)
        header.setFirstSectionMovable(True)
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(40)
        self._table.cellDoubleClicked.connect(self._on_cell_activated)
        self._table.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self._table.customContextMenuRequested.connect(
            self._show_context_menu
        )
        layout.addWidget(self._table, 1)

    def reload_legend(self, palette: Mapping[str, str]) -> None:
        """Rebuild the color legend from the current palette."""

        self._palette = dict(palette)
        while self._legend.count():
            item = self._legend.takeAt(0)
            widget = item.widget() if item is not None else None
            if widget is not None:
                widget.deleteLater()
        for key in _LEGEND_KEYS:
            sample = QLabel("  ", self)
            hex_color = color_for(self._palette, key)
            sample.setStyleSheet(
                f"background:{hex_color}; border:1px solid #80868b;"
                f" min-width:16px;"
            )
            caption = QLabel(status_color_label(key), self)
            self._legend.addWidget(sample)
            self._legend.addWidget(caption)
        self._legend.addStretch(1)

    def set_rows(
        self,
        rows: Sequence[MtoWorklistRow],
        *,
        palette: Mapping[str, str],
        is_banned: Callable[[str, str], bool],
        allowed_kits: set[tuple[str, str]] | None = None,
        pin_view_for: Callable[[str, str], ExportPinView] | None = None,
        compare_status_for: Callable[[str, str], AutoMtoCompareStatus]
        | None = None,
    ) -> None:
        """Replace the worklist snapshot and rebuild the table.

        Args:
            rows: ``MtoWorklistRow`` values from ``list_mto_worklist``.
            palette: Status color map.
            is_banned: Hide banned ``(title, mark)`` pairs.
            allowed_kits: If set, only these identities are shown.
                ``None`` keeps every non-banned kit.
            pin_view_for: Kit-level pin display lookup. ``None`` leaves
                the pin column empty.
            compare_status_for: Kit-level AutoMTO content-compare status.
                ``None`` leaves the column as «—».
        """

        self._rows = tuple(rows)
        self._is_banned = is_banned
        self._allowed_kits = allowed_kits
        self._pin_view_for = pin_view_for or (
            lambda _title, _mark: export_pin_view(None)
        )
        self._compare_status_for = compare_status_for or (
            lambda _title, _mark: _EMPTY_COMPARE_STATUS
        )
        self.reload_legend(palette)
        self._rebuild_table()

    def restore_filters(self, settings: QSettings) -> None:
        """Load filter widgets from QSettings without applying twice."""

        boxes = self._filter_checkboxes()
        self._filter.blockSignals(True)
        for box in boxes:
            box.blockSignals(True)
        try:
            self._filter.setText(
                str(settings.value("window/mto_worklist_filter") or "")
            )
            self._code_a.setChecked(
                _settings_bool(
                    settings, "window/mto_worklist_code_a", False
                )
            )
            self._tdo.setChecked(
                _settings_bool(settings, "window/mto_worklist_tdo", False)
            )
            self._current.setChecked(
                _settings_bool(
                    settings, "window/mto_worklist_current", False
                )
            )
            self._current_ifc.setChecked(
                _settings_bool(
                    settings, "window/mto_worklist_current_ifc", False
                )
            )
            self._no_as_build.setChecked(
                _settings_bool(
                    settings, "window/mto_worklist_no_as_build", False
                )
            )
            self._only_as_build.setChecked(
                _settings_bool(
                    settings, "window/mto_worklist_only_as_build", False
                )
            )
            self._missing.setChecked(
                _settings_bool(
                    settings, "window/mto_worklist_missing", False
                )
            )
            self._no_rd.setChecked(
                _settings_bool(
                    settings, "window/mto_worklist_no_rd", False
                )
            )
            self._no_google.setChecked(
                _settings_bool(
                    settings, "window/mto_worklist_no_google", False
                )
            )
            self._no_f_status.setChecked(
                _settings_bool(
                    settings, "window/mto_worklist_no_f_status", False
                )
            )
            self._problems.setChecked(
                _settings_bool(
                    settings, "window/mto_worklist_problems", False
                )
            )
        finally:
            self._filter.blockSignals(False)
            for box in boxes:
                box.blockSignals(False)

    def save_filters(self, settings: QSettings) -> None:
        """Persist filter widgets into QSettings."""

        settings.setValue(
            "window/mto_worklist_filter", self._filter.text()
        )
        settings.setValue(
            "window/mto_worklist_code_a", self._code_a.isChecked()
        )
        settings.setValue(
            "window/mto_worklist_tdo", self._tdo.isChecked()
        )
        settings.setValue(
            "window/mto_worklist_current", self._current.isChecked()
        )
        settings.setValue(
            "window/mto_worklist_current_ifc",
            self._current_ifc.isChecked(),
        )
        settings.setValue(
            "window/mto_worklist_no_as_build",
            self._no_as_build.isChecked(),
        )
        settings.setValue(
            "window/mto_worklist_only_as_build",
            self._only_as_build.isChecked(),
        )
        settings.setValue(
            "window/mto_worklist_missing", self._missing.isChecked()
        )
        settings.setValue(
            "window/mto_worklist_no_rd", self._no_rd.isChecked()
        )
        settings.setValue(
            "window/mto_worklist_no_google",
            self._no_google.isChecked(),
        )
        settings.setValue(
            "window/mto_worklist_no_f_status",
            self._no_f_status.isChecked(),
        )
        settings.setValue(
            "window/mto_worklist_problems", self._problems.isChecked()
        )

    def _filter_checkboxes(self) -> tuple[QCheckBox, ...]:
        return (
            self._code_a,
            self._tdo,
            self._current,
            self._current_ifc,
            self._no_as_build,
            self._only_as_build,
            self._missing,
            self._no_rd,
            self._no_google,
            self._no_f_status,
            self._problems,
        )

    def _kit_is_hidden(self, title: str, mark: str) -> bool:
        if self._is_banned(title, mark):
            return True
        if self._allowed_kits is None:
            return False
        return kit_identity_key(title, mark) not in self._allowed_kits

    def _rebuild_table(self) -> None:
        table = self._table
        table.setSortingEnabled(False)
        visible_rows = [
            row
            for row in self._rows
            if not self._kit_is_hidden(row.title, row.mark)
        ]
        table.setRowCount(len(visible_rows))
        titles: set[str] = set()
        kits: set[tuple[str, str]] = set()
        for row_index, row in enumerate(visible_rows):
            self._fill_row(row_index, row)
            titles.add(row.title)
            kits.add(kit_identity_key(row.title, row.mark))
        self._total_titles = len(titles)
        self._total_kits = len(kits)
        self._total_rows = len(visible_rows)
        table.setSortingEnabled(True)
        self._apply_row_visibility()

    def _fill_row(self, row_index: int, row: MtoWorklistRow) -> None:
        table = self._table
        package = row.package_label
        if row.package_count > 1:
            package = f"{package} (+{row.package_count - 1})"
        google_cell = worklist_google_cell(row)
        if google_cell.palette_key == "problem":
            google_cell = replace(
                google_cell,
                foreground=color_for(self._palette, "problem"),
            )
        problem_text = ", ".join(
            collision_kind_label(k) for k in row.problem_kinds
        )
        pin_view = self._pin_view_for(row.title, row.mark)
        compare_status = self._compare_status_for(row.title, row.mark)
        values = (
            row.title,
            row.mark,
            row.revision_text,
            status_short_label(row.status),
            row.letters,
            worklist_link_text(row),
            google_cell.text,
            "Да" if row.is_as_build else "",
            pin_view.text,
            compare_status.text,
            "Да" if row.mto_path else "Нет",
            _format_mtime(row.mto_mtime_ns),
            package,
            GAP_LABELS.get(row.gap_kind, ""),
            problem_text,
            row.mto_path,
        )
        for column, text in enumerate(values):
            if column == _COL_DATE:
                item: QTableWidgetItem = _SortItem(text)
                mtime = row.mto_mtime_ns
                item.setData(_ROLE_SORT, mtime if mtime is not None else -1)
            else:
                item = QTableWidgetItem(text)
            item.setData(_ROLE_ROW, row)
            if column == _COL_PIN:
                apply_export_pin_view(item, pin_view, self._palette)
            if column == _COL_COMPARE:
                apply_auto_mto_compare_status(item, compare_status)
            if column == _COL_STATUS:
                _paint_fill(
                    item,
                    color_for(self._palette, row.status or "empty"),
                )
                item.setToolTip(_status_tooltip(row))
            if column == _COL_LINK:
                item.setToolTip(_revision_link_tooltip(row))
                if row.gap_kind:
                    item.setForeground(
                        QBrush(
                            QColor(color_for(self._palette, "problem"))
                        )
                    )
            if column == _COL_GOOGLE:
                apply_monitor_cell(item, google_cell)
            if (
                column == _COL_FILE
                and row.gap_kind
                and not row.mto_path
            ):
                item.setForeground(
                    QBrush(QColor(color_for(self._palette, "problem")))
                )
            if column == _COL_PACKAGE:
                item.setToolTip(row.package_path)
                if row.gap_kind in {"no_package", "no_rd"}:
                    item.setForeground(
                        QBrush(
                            QColor(color_for(self._palette, "problem"))
                        )
                    )
            if column == _COL_PATH:
                item.setToolTip(row.mto_path)
            if column == _COL_GAP and row.gap_kind:
                item.setForeground(
                    QBrush(QColor(color_for(self._palette, "problem")))
                )
            if column == _COL_PROBLEMS and problem_text:
                item.setToolTip(problem_text)
            table.setItem(row_index, column, item)
        if row.problem_kinds or row.gap_kind or row.is_current:
            for column in range(table.columnCount()):
                item = table.item(row_index, column)
                if item is None:
                    continue
                font = QFont(item.font())
                if row.problem_kinds or row.gap_kind:
                    font.setBold(True)
                if row.is_current:
                    font.setUnderline(True)
                item.setFont(font)

    def _row_at(self, row_index: int) -> MtoWorklistRow | None:
        if row_index < 0:
            return None
        item = self._table.item(row_index, 0)
        payload = item.data(_ROLE_ROW) if item else None
        return payload if isinstance(payload, MtoWorklistRow) else None

    def _row_matches_text(self, row: MtoWorklistRow, needle: str) -> bool:
        if not needle:
            return True
        chunks = [
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
            self._pin_view_for(row.title, row.mark).text,
            self._compare_status_for(row.title, row.mark).text,
            *row.problem_kinds,
            *(collision_kind_label(k) for k in row.problem_kinds),
        ]
        return needle in " ".join(chunks).casefold()

    def _row_matches_filters(self, row: MtoWorklistRow) -> bool:
        if self._code_a.isChecked() and row.status != "code_a":
            return False
        if self._tdo.isChecked() and row.status not in _TDO_STATUSES:
            return False
        if self._current.isChecked() and not row.is_current:
            return False
        if self._current_ifc.isChecked() and not row.is_current_ifc:
            return False
        if self._no_as_build.isChecked() and row.is_as_build:
            return False
        if self._only_as_build.isChecked() and not row.is_as_build:
            return False
        if self._missing.isChecked() and not (
            row.gap_kind or not row.mto_path
        ):
            return False
        if self._no_rd.isChecked() and row.gap_kind != "no_rd":
            return False
        if self._no_google.isChecked() and (
            row.in_google or row.in_issuance
        ):
            return False
        if self._no_f_status.isChecked() and row.has_f_status:
            return False
        if self._problems.isChecked() and not row.problem_kinds:
            return False
        return True

    @Slot(bool)
    def _on_as_build_toggled(self, checked: bool) -> None:
        sender = self.sender()
        other = (
            self._only_as_build
            if sender is self._no_as_build
            else self._no_as_build
        )
        other.setEnabled(not checked)
        if checked and other.isChecked():
            other.blockSignals(True)
            other.setChecked(False)
            other.blockSignals(False)
        self._apply_row_visibility()

    @Slot()
    def _apply_row_visibility(self) -> None:
        needle = self._filter.text().strip().casefold()
        table = self._table
        for index in range(table.rowCount()):
            payload = self._row_at(index)
            if payload is None or self._kit_is_hidden(
                payload.title, payload.mark
            ):
                table.setRowHidden(index, True)
                continue
            visible = self._row_matches_filters(
                payload
            ) and self._row_matches_text(payload, needle)
            table.setRowHidden(index, not visible)
        self._update_counts()

    def _update_counts(self) -> None:
        titles: set[str] = set()
        kits: set[tuple[str, str]] = set()
        shown = code_a = tdo = missing = 0
        no_rd = no_google = no_f = problems = 0
        table = self._table
        for index in range(table.rowCount()):
            if table.isRowHidden(index):
                continue
            payload = self._row_at(index)
            if payload is None:
                continue
            shown += 1
            titles.add(payload.title)
            kits.add(kit_identity_key(payload.title, payload.mark))
            if payload.status == "code_a":
                code_a += 1
            if payload.status in _TDO_STATUSES:
                tdo += 1
            if payload.gap_kind or not payload.mto_path:
                missing += 1
            if payload.gap_kind == "no_rd":
                no_rd += 1
            if not (payload.in_google or payload.in_issuance):
                no_google += 1
            if not payload.has_f_status:
                no_f += 1
            if payload.problem_kinds:
                problems += 1
        self._counts_scope.setText(
            f"Титулы: {len(titles)} / {self._total_titles} · "
            f"Комплекты: {len(kits)} / {self._total_kits} · "
            f"Строки: {shown} / {self._total_rows}"
        )
        self._counts_breakdown.setText(
            f"Код A: {code_a} · ТДО: {tdo} · Нет MTO: {missing} · "
            f"Нет в РД: {no_rd} · Нет в Google: {no_google} · "
            f"Без статуса F: {no_f} · Проблемы: {problems}"
        )

    @Slot(int, int)
    def _on_cell_activated(self, row: int, _column: int) -> None:
        payload = self._row_at(row)
        if payload is None:
            return
        path = payload.mto_path or payload.package_path
        if path:
            open_path(path)

    def _show_context_menu(self, position: QPoint) -> None:
        table = self._table
        row_index = table.rowAt(position.y())
        if row_index < 0:
            return
        table.selectRow(row_index)
        column = table.columnAt(position.x())
        if column >= 0:
            table.setCurrentCell(row_index, column)
        payload = self._row_at(row_index)
        if payload is None:
            return
        menu = QMenu(self)
        open_mto = menu.addAction("Открыть файл MTO")
        open_pkg = menu.addAction("Открыть папку пакета")
        copy_mto = menu.addAction("Копировать путь MTO")
        copy_pkg = menu.addAction("Копировать путь пакета")
        copy_list = menu.addAction("Копировать список (видимые строки)")
        open_mto.setEnabled(bool(payload.mto_path))
        open_pkg.setEnabled(bool(payload.package_path))
        copy_mto.setEnabled(bool(payload.mto_path))
        copy_pkg.setEnabled(bool(payload.package_path))
        menu.addSeparator()
        show_kits = menu.addAction('Показать в «Комплекты»')
        show_docs = menu.addAction('Показать в «Все документы»')
        self.prepare_context_menu.emit(menu)
        chosen = exec_tracked_menu(
            menu, MENU_MTO_WORKLIST, table.viewport().mapToGlobal(position)
        )
        if chosen == open_mto:
            open_path(payload.mto_path)
        elif chosen == open_pkg:
            open_path(payload.package_path)
        elif chosen == copy_mto:
            _copy_text(payload.mto_path)
        elif chosen == copy_pkg:
            _copy_text(payload.package_path)
        elif chosen == copy_list:
            self._copy_visible_tsv()
        elif chosen == show_kits:
            self.kit_activated.emit(payload.title, payload.mark)
        elif chosen == show_docs:
            self.kit_documents_requested.emit(
                payload.title, payload.mark
            )

    def _copy_visible_tsv(self) -> None:
        lines = ["\t".join(_HEADERS)]
        table = self._table
        for index in range(table.rowCount()):
            if table.isRowHidden(index):
                continue
            values: list[str] = []
            for column in range(table.columnCount()):
                item = table.item(index, column)
                text = item.text() if item is not None else ""
                values.append(
                    text.replace("\t", " ").replace("\n", " ")
                )
            lines.append("\t".join(values))
        _copy_text("\n".join(lines) + "\n")


def _revision_source_text(row: MtoWorklistRow) -> str:
    revision = row.revision_text or "—"
    return (
        "Источник этапа ревизии и меток — Google F / "
        "«Выдача РД ПД»; данные сопоставлены с ревизией имени "
        f"файла «{revision}». Они не подтверждают физическое "
        "наличие РД или MTO."
    )


def _revision_link_tooltip(row: MtoWorklistRow) -> str:
    """Explain journal status separately from physical file presence."""

    lines = [_revision_source_text(row)]
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
        lines.append(
            "РД✓ — физически найдена папка пакета package_path: "
            f"{row.package_path}"
        )
    else:
        lines.append(
            "РД✗ — package_path пуст: папка пакета РД физически "
            "не найдена."
        )
    if row.mto_path:
        lines.append(
            "MTO✓ — физически найден MTO-файл mto_path: "
            f"{row.mto_path}"
        )
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


def _status_tooltip(row: MtoWorklistRow) -> str:
    """Explain the revision status source and its physical gap."""

    status = status_short_label(row.status) if row.status else "не задан"
    gap = GAP_LABELS.get(row.gap_kind, row.gap_kind) or "нет"
    lines = [
        _revision_source_text(row),
        f"Этап ревизии: {status}.",
        f"Разрыв: {gap}.",
    ]
    if row.gap_kind == "no_package":
        lines.append(
            "Согласование в Google не доказывает наличие папки "
            "пакета РД."
        )
    return "\n".join(lines)


def _format_mtime(value: int | None) -> str:
    if value is None:
        return ""
    try:
        return datetime.fromtimestamp(value / 1_000_000_000).strftime(
            "%Y.%m.%d"
        )
    except (OSError, OverflowError, ValueError):
        return ""


def apply_export_pin_view(
    item: QTableWidgetItem,
    view: ExportPinView,
    palette: Mapping[str, str],
) -> None:
    """Write pin text, tooltip, and stale fill onto one table cell.

    Shared by «MTO · Перечень», Комплекты, and the documents list.

    Args:
        item: Target table cell.
        view: Display payload from ``export_pin_view``.
        palette: User status-color map; stale uses ``export_pin_stale``.
    """

    item.setText(view.text)
    item.setToolTip(view.tooltip)
    if view.is_stale:
        _paint_fill(item, color_for(palette, "export_pin_stale"))


def apply_auto_mto_compare_status(
    item: QTableWidgetItem,
    status: AutoMtoCompareStatus,
) -> None:
    """Write AutoMTO content-compare text, tooltip, and match fill.

    Shared by «MTO · Перечень» and Комплекты.

    Args:
        item: Target table cell.
        status: Display payload from ``auto_mto_compare_status``.
    """

    item.setText(status.text or "—")
    item.setToolTip(status.tooltip)
    if status.paint_match is True:
        item.setBackground(QBrush(_REV_MATCH_BG))
        font = QFont(item.font())
        font.setBold(True)
        item.setFont(font)
    elif status.paint_match is False:
        item.setBackground(QBrush(_REV_DIFF_BG))


def _copy_text(text: str) -> None:
    if text:
        QApplication.clipboard().setText(text)


def _paint_fill(item: QTableWidgetItem, hex_color: str) -> None:
    color = QColor(hex_color)
    if not color.isValid():
        return
    item.setBackground(QBrush(color))
    item.setForeground(
        QBrush(
            QColor("#ffffff")
            if color.lightness() < 140
            else QColor("#202124")
        )
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
