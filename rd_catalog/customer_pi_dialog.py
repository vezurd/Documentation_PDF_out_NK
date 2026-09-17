"""Dialog to reload the customer PI pickle and rebuild the АвтоМто catalog."""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from dataclasses import replace
from pathlib import Path

from PySide6.QtCore import QEvent, Qt, Signal, Slot
from PySide6.QtGui import QBrush, QCloseEvent, QColor, QResizeEvent, QShowEvent
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.customer_pi import (
    DEFAULT_XLSB,
    CustomerPiRecord,
    CustomerPiStore,
    default_pickle_path,
    load_customer_pi,
)
from rd_catalog.customer_pi_auto_mto import (
    HEADER_ROW,
    PAINT_OUTSIDE_KITS,
    STATUS_COLORS,
    STATUS_LABELS,
    STATUS_SKIPPED,
    STATUS_WRITING,
    SpecProgressRow,
    default_auto_mto_dir,
    format_kits_coverage_summary,
    format_spec_progress_summary,
    format_store_summary,
    list_spec_progress_rows,
    mto_sheet_row,
    spec_as_mto_stem,
    spec_paint_key,
)
from rd_catalog.customer_pi_thread import (
    JOB_BOTH,
    JOB_BUILD_CATALOG,
    JOB_RELOAD_XLSB,
    CustomerPiThread,
)
from rd_catalog.kits import kit_identity_key
from rd_catalog.path_actions import open_path

_SPEC_COLUMNS = (
    "Спецификация",
    "Титул",
    "Марка",
    "Дисц.",
    "Рев.",
    "Строк",
    "Статус",
)
_COL_SPEC = 0
_COL_TITLE = 1
_COL_MARK = 2
_COL_DISC = 3
_COL_REV = 4
_COL_ROWS = 5
_COL_STATUS = 6
# Relative widths: spec takes about half the table, the rest share the remainder.
_SPEC_STRETCH = (46, 10, 10, 8, 10, 8, 8)
_ROLE_SPEC = Qt.ItemDataRole.UserRole
_ROLE_ROW = Qt.ItemDataRole.UserRole + 1
_KITS_SUMMARY_STYLE = (
    f"background-color: {STATUS_COLORS[PAINT_OUTSIDE_KITS]}; "
    "color: #202124; padding: 4px 8px; border-radius: 2px;"
)


def _paint_row_status(table: QTableWidget, row: int, paint_key: str) -> None:
    hex_color = STATUS_COLORS.get(paint_key, "#e8eaed")
    color = QColor(hex_color)
    background = QBrush(color)
    foreground = QBrush(
        QColor("#ffffff") if color.lightness() < 140 else QColor("#202124")
    )
    for column in range(table.columnCount()):
        item = table.item(row, column)
        if item is None:
            continue
        item.setBackground(background)
        item.setForeground(foreground)


def _spec_filter_haystack(row: SpecProgressRow) -> str:
    title = (row.title or "").strip()
    mark = (row.mark or "").strip()
    status = STATUS_LABELS.get(row.status, row.status)
    return " ".join(
        (
            row.spec,
            title,
            mark,
            f"{title}-{mark}",
            f"{title}_{mark}",
            row.discipline,
            row.rd_revision,
            status,
        )
    ).casefold()


def _stretch_table_columns(
    table: QTableWidget,
    weights: Sequence[int],
) -> None:
    """Resize Interactive sections so they share the viewport by ``weights``."""

    header = table.horizontalHeader()
    available = table.viewport().width()
    count = len(weights)
    if available < 40 or header.count() != count:
        return
    total = sum(weights)
    if total <= 0:
        return
    minimum = max(header.minimumSectionSize(), 32)
    widths = [max(minimum, available * weight // total) for weight in weights]
    others = sum(widths[1:]) if count > 1 else 0
    widths[0] = max(minimum, available - others)
    header.blockSignals(True)
    try:
        for column, width in enumerate(widths):
            header.resizeSection(column, width)
    finally:
        header.blockSignals(False)


class MtoPreviewDialog(QDialog):
    """Show PI rows of one spec in АвтоМто / ``ColNames.MTO`` column order."""

    def __init__(
        self,
        parent: QWidget | None,
        *,
        spec: str,
        records: Sequence[CustomerPiRecord],
    ) -> None:
        super().__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)
        self.setWindowFlag(Qt.WindowType.Window, True)
        count = len(records)
        self.setWindowTitle(f"MTO · {spec or '—'} ({count})")
        self.setMinimumSize(960, 420)
        self.resize(1100, 560)
        layout = QVBoxLayout(self)
        caption = QLabel(
            f"{spec or '—'} · строк {count} · столбцы как лист «Спецификация»",
            self,
        )
        caption.setWordWrap(True)
        layout.addWidget(caption)
        table = QTableWidget(count, len(HEADER_ROW), self)
        table.setHorizontalHeaderLabels(list(HEADER_ROW))
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSortingEnabled(False)
        table.verticalHeader().setVisible(False)
        table.setWordWrap(False)
        flags = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        for row_index, record in enumerate(records):
            cells = mto_sheet_row(record, row_index + 1)
            for column, value in enumerate(cells):
                item = QTableWidgetItem()
                item.setFlags(flags)
                if isinstance(value, int):
                    item.setData(Qt.ItemDataRole.DisplayRole, value)
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    )
                else:
                    item.setText("" if value is None else str(value))
                table.setItem(row_index, column, item)
        table.setSortingEnabled(True)
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(48)
        header.setSortIndicatorShown(True)
        layout.addWidget(table, 1)
        self._table = table
        self._column_weights = tuple(10 for _ in HEADER_ROW)
        close_row = QHBoxLayout()
        close_row.addStretch(1)
        close_button = QPushButton("Закрыть", self)
        close_button.clicked.connect(self.close)
        close_row.addWidget(close_button)
        layout.addLayout(close_row)
        table.installEventFilter(self)
        _stretch_table_columns(table, self._column_weights)

    def eventFilter(self, watched: QWidget, event: QEvent) -> bool:
        if watched is self._table and event.type() == QEvent.Type.Resize:
            _stretch_table_columns(self._table, self._column_weights)
        return super().eventFilter(watched, event)

    def showEvent(self, event: QShowEvent) -> None:
        super().showEvent(event)
        _stretch_table_columns(self._table, self._column_weights)


class CustomerPiDialog(QDialog):
    """Show pickle meta and start xlsb/catalog workers.

    The source ``.xlsb`` is never overwritten. Window glue should treat this
    dialog as a busy worker while ``is_busy()`` is true.
    """

    pickle_updated = Signal()
    catalog_rebuilt = Signal()

    def __init__(
        self,
        parent: QWidget | None = None,
        *,
        pickle_path: str | Path | None = None,
        xlsb_path: str | Path | None = None,
        dest_dir: str | Path | None = None,
        store: CustomerPiStore | None = None,
        kit_keys: Iterable[tuple[str, str]] | None = None,
    ) -> None:
        """Build the dialog.

        Args:
            parent: Optional Qt parent.
            pickle_path: PI pickle to read/write; default packaged dump.
            xlsb_path: Last known customer ``.xlsb``; pickle ``source_path``
                wins when the line is empty.
            dest_dir: АвтоМто catalog root.
            store: Optional already loaded dump (skips pickle IO on open).
            kit_keys: Visible Комплекты identities. Specs whose title+mark
                is missing from this set are painted faded green. ``None``
                skips that overlay.
        """

        super().__init__(parent)
        self._pickle_path = Path(pickle_path) if pickle_path else default_pickle_path()
        self._dest_dir = Path(dest_dir) if dest_dir else default_auto_mto_dir()
        self._store = store
        self._kit_keys: frozenset[tuple[str, str]] | None = (
            None
            if kit_keys is None
            else frozenset(kit_identity_key(title, mark) for title, mark in kit_keys)
        )
        self._thread: CustomerPiThread | None = None
        self._spec_rows: tuple[SpecProgressRow, ...] = ()
        self._spec_row_index: dict[str, list[int]] = {}
        self._preview_windows: list[MtoPreviewDialog] = []
        self._syncing_columns = False
        self.setWindowTitle("База заказчика (Авто МТО)")
        self.setMinimumSize(960, 520)
        self.resize(1180, 640)
        self._build()
        initial_xlsb = str(xlsb_path) if xlsb_path else ""
        if not initial_xlsb:
            initial_xlsb = self._source_path_from_store() or str(DEFAULT_XLSB)
        self._xlsb_edit.setText(initial_xlsb)
        self._refresh_meta()

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        hint = QLabel(
            "Pickle и каталог АвтоМТО лежат локально в "
            "«rd_catalog/База заказчика». Исходный xlsb заказчика не меняется. "
            "Читаются полный отчёт (49 столбцов) и урезанный BCC (23 столбца). "
            "Каталог — MTO, плюс BOM/BOE/DS (в имени токен меняется на MTO), "
            "лист «Спецификация», имя «{спека}_{ревизия РД}_RU.xlsx». Справа — "
            "уникальные спеки pickle: статус и цвет меняются по ходу сборки. "
            "Двойной щелчок открывает строки в порядке столбцов MTO. "
            "Блекло-зелёные строки — титул–марка нет в Комплектах.",
            self,
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        splitter = QSplitter(Qt.Orientation.Horizontal, self)
        splitter.setChildrenCollapsible(False)

        left = QWidget(splitter)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        self._meta_view = QPlainTextEdit(left)
        self._meta_view.setReadOnly(True)
        left_layout.addWidget(self._meta_view, 1)

        xlsb_row = QHBoxLayout()
        xlsb_row.addWidget(QLabel("xlsb:", left))
        self._xlsb_edit = QLineEdit(left)
        xlsb_row.addWidget(self._xlsb_edit, 1)
        browse = QPushButton("Обзор…", left)
        browse.clicked.connect(self._on_browse)
        xlsb_row.addWidget(browse)
        left_layout.addLayout(xlsb_row)

        self._log_view = QPlainTextEdit(left)
        self._log_view.setReadOnly(True)
        self._log_view.setPlaceholderText("Журнал фоновой задачи…")
        left_layout.addWidget(self._log_view, 1)
        splitter.addWidget(left)

        right = QWidget(splitter)
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.addWidget(QLabel("Спецификации pickle", right))
        self._spec_summary_label = QLabel("Нет данных", right)
        self._spec_summary_label.setWordWrap(True)
        right_layout.addWidget(self._spec_summary_label)
        self._kits_summary_label = QLabel("", right)
        self._kits_summary_label.setWordWrap(True)
        self._kits_summary_label.setStyleSheet(_KITS_SUMMARY_STYLE)
        self._kits_summary_label.setVisible(self._kit_keys is not None)
        right_layout.addWidget(self._kits_summary_label)
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Фильтр:", right))
        self._spec_filter = QLineEdit(right)
        self._spec_filter.setPlaceholderText("Титул, марка…")
        self._spec_filter.setClearButtonEnabled(True)
        self._spec_filter.textChanged.connect(self._apply_spec_filter)
        filter_row.addWidget(self._spec_filter, 1)
        right_layout.addLayout(filter_row)
        self._spec_table = QTableWidget(0, len(_SPEC_COLUMNS), right)
        self._spec_table.setHorizontalHeaderLabels(list(_SPEC_COLUMNS))
        self._spec_table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self._spec_table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self._spec_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._spec_table.verticalHeader().setVisible(False)
        self._spec_table.setWordWrap(False)
        self._spec_table.setSortingEnabled(True)
        header = self._spec_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(48)
        header.setSortIndicatorShown(True)
        self._spec_table.cellDoubleClicked.connect(self._on_spec_double_clicked)
        self._spec_table.installEventFilter(self)
        right_layout.addWidget(self._spec_table, 1)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([640, 480])
        layout.addWidget(splitter, 1)

        buttons = QHBoxLayout()
        self._reload_button = QPushButton("Обновить pickle из xlsb", self)
        self._catalog_button = QPushButton("Собрать каталог АвтоМТО", self)
        self._both_button = QPushButton("Обновить и собрать", self)
        self._reload_button.clicked.connect(self._on_reload)
        self._catalog_button.clicked.connect(self._on_catalog)
        self._both_button.clicked.connect(self._on_both)
        buttons.addWidget(self._reload_button)
        buttons.addWidget(self._catalog_button)
        buttons.addWidget(self._both_button)
        layout.addLayout(buttons)

        footer = QHBoxLayout()
        self._open_catalog_button = QPushButton("Открыть каталог АвтоМТО", self)
        self._open_catalog_button.setToolTip(str(self._dest_dir))
        self._open_catalog_button.clicked.connect(self._on_open_catalog)
        self._close_button = QPushButton("Закрыть", self)
        self._close_button.clicked.connect(self.close)
        footer.addWidget(self._open_catalog_button)
        footer.addStretch(1)
        footer.addWidget(self._close_button)
        layout.addLayout(footer)

    def xlsb_path(self) -> Path:
        """Return the path currently shown in the xlsb line."""

        return Path(self._xlsb_edit.text().strip())

    def pickle_path(self) -> Path:
        """Return the pickle this dialog reads and writes."""

        return self._pickle_path

    def dest_dir(self) -> Path:
        """Return the АвтоМто catalog root."""

        return self._dest_dir

    def is_busy(self) -> bool:
        """Return True while a worker thread is running."""

        return self._thread is not None and self._thread.isRunning()

    def eventFilter(self, watched: QWidget, event: QEvent) -> bool:
        if watched is self._spec_table and event.type() == QEvent.Type.Resize:
            self._sync_column_widths()
        return super().eventFilter(watched, event)

    def showEvent(self, event: QShowEvent) -> None:
        super().showEvent(event)
        self._sync_column_widths()

    def resizeEvent(self, event: QResizeEvent) -> None:
        super().resizeEvent(event)
        self._sync_column_widths()

    def _sync_column_widths(self) -> None:
        if self._syncing_columns:
            return
        self._syncing_columns = True
        try:
            _stretch_table_columns(self._spec_table, _SPEC_STRETCH)
        finally:
            self._syncing_columns = False

    def _source_path_from_store(self) -> str:
        if self._store is None:
            return ""
        return str(self._store.meta.get("source_path") or "")

    def _refresh_meta(self) -> None:
        if self._store is None and self._pickle_path.is_file():
            try:
                self._store = load_customer_pi(self._pickle_path)
            except Exception as exc:
                self._meta_view.setPlainText(
                    f"Не удалось прочитать pickle:\n{self._pickle_path}\n{exc}"
                )
                self._fill_spec_table(())
                return
        if self._store is None:
            self._meta_view.setPlainText(
                f"Pickle ещё нет:\n{self._pickle_path}\n"
                "Сначала обновите его из xlsb."
            )
            self._fill_spec_table(())
            return
        extra = (
            f"\nКаталог АвтоМТО: {self._dest_dir}"
            f"\n({ 'есть' if self._dest_dir.is_dir() else 'папки ещё нет' })"
        )
        self._meta_view.setPlainText(format_store_summary(self._store) + extra)
        self._fill_spec_table(
            list_spec_progress_rows(self._store, dest=self._dest_dir)
        )

    def _fill_spec_table(self, rows: Sequence[SpecProgressRow]) -> None:
        self._spec_rows = tuple(rows)
        index_map: dict[str, list[int]] = {}
        for index, row in enumerate(self._spec_rows):
            index_map.setdefault(row.spec, []).append(index)
            output = spec_as_mto_stem(row.spec)
            if output and output != row.spec:
                index_map.setdefault(output, []).append(index)
        self._spec_row_index = index_map
        table = self._spec_table
        table.setSortingEnabled(False)
        table.setRowCount(len(self._spec_rows))
        flags = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        for index, row in enumerate(self._spec_rows):
            spec_item = QTableWidgetItem(row.spec or "—")
            spec_item.setFlags(flags)
            spec_item.setData(_ROLE_SPEC, row.spec)
            spec_item.setData(_ROLE_ROW, row)
            tooltip = row.note or ""
            kit_label = f"{row.title}-{row.mark}".strip("-")
            if kit_label:
                tooltip = f"{kit_label}\n{tooltip}".strip()
            if self._kit_keys is not None and spec_paint_key(
                row, self._kit_keys
            ) == PAINT_OUTSIDE_KITS:
                extra = "нет в Комплектах"
                tooltip = f"{tooltip}\n{extra}".strip() if tooltip else extra
            if tooltip:
                spec_item.setToolTip(tooltip)
            table.setItem(index, _COL_SPEC, spec_item)
            title_item = QTableWidgetItem(row.title or "—")
            title_item.setFlags(flags)
            table.setItem(index, _COL_TITLE, title_item)
            mark_item = QTableWidgetItem(row.mark or "—")
            mark_item.setFlags(flags)
            table.setItem(index, _COL_MARK, mark_item)
            disc_item = QTableWidgetItem(row.discipline)
            disc_item.setFlags(flags)
            table.setItem(index, _COL_DISC, disc_item)
            rev_item = QTableWidgetItem(row.rd_revision or "—")
            rev_item.setFlags(flags)
            table.setItem(index, _COL_REV, rev_item)
            count_item = QTableWidgetItem()
            count_item.setFlags(flags)
            count_item.setData(Qt.ItemDataRole.DisplayRole, row.row_count)
            count_item.setTextAlignment(
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
            )
            table.setItem(index, _COL_ROWS, count_item)
            status_item = QTableWidgetItem(
                STATUS_LABELS.get(row.status, row.status)
            )
            status_item.setFlags(flags)
            if tooltip:
                status_item.setToolTip(tooltip)
            table.setItem(index, _COL_STATUS, status_item)
            _paint_row_status(
                table, index, spec_paint_key(row, self._kit_keys)
            )
        table.setSortingEnabled(True)
        self._apply_spec_filter()
        self._sync_column_widths()

    def _progress_row_at(self, visual_row: int) -> SpecProgressRow | None:
        item = self._spec_table.item(visual_row, _COL_SPEC)
        if item is None:
            return None
        payload = item.data(_ROLE_ROW)
        return payload if isinstance(payload, SpecProgressRow) else None

    def _visual_rows_for_spec(self, spec: str) -> list[int]:
        wanted = set(self._spec_row_index.get(spec) or ())
        if not wanted:
            return []
        specs = {self._spec_rows[index].spec for index in wanted}
        found: list[int] = []
        table = self._spec_table
        for visual in range(table.rowCount()):
            item = table.item(visual, _COL_SPEC)
            if item is None:
                continue
            if item.data(_ROLE_SPEC) in specs:
                found.append(visual)
        return found

    def _apply_spec_filter(self) -> None:
        needle = self._spec_filter.text().strip().casefold()
        table = self._spec_table
        for visual in range(table.rowCount()):
            row = self._progress_row_at(visual)
            visible = True if not needle or row is None else needle in _spec_filter_haystack(row)
            table.setRowHidden(visual, not visible)
        self._refresh_spec_summary_label()

    def _refresh_spec_summary_label(self) -> None:
        if not self._spec_rows:
            self._spec_summary_label.setText("Нет данных")
            self._kits_summary_label.setText("")
            return
        text = format_spec_progress_summary(self._spec_rows)
        visible = sum(
            1
            for index in range(self._spec_table.rowCount())
            if not self._spec_table.isRowHidden(index)
        )
        if visible != len(self._spec_rows):
            text = f"{text} · показано {visible}"
        self._spec_summary_label.setText(text)
        if self._kit_keys is None:
            self._kits_summary_label.setVisible(False)
            return
        self._kits_summary_label.setVisible(True)
        self._kits_summary_label.setText(
            format_kits_coverage_summary(self._spec_rows, self._kit_keys)
        )

    def _set_spec_status(self, spec: str, status: str) -> None:
        indices = self._spec_row_index.get(spec) or ()
        if not indices:
            return
        scrolled = False
        rows = list(self._spec_rows)
        changed = False
        for row in indices:
            current = rows[row]
            if current.status == status:
                continue
            rows[row] = replace(current, status=status)
            changed = True
        if changed:
            self._spec_rows = tuple(rows)
        visual_rows = self._visual_rows_for_spec(spec)
        for visual in visual_rows:
            current = self._progress_row_at(visual)
            if current is None:
                continue
            updated = replace(current, status=status) if current.status != status else current
            spec_item = self._spec_table.item(visual, _COL_SPEC)
            status_item = self._spec_table.item(visual, _COL_STATUS)
            if spec_item is not None:
                spec_item.setData(_ROLE_ROW, updated)
            if status_item is not None:
                status_item.setText(STATUS_LABELS.get(status, status))
            _paint_row_status(
                self._spec_table,
                visual,
                spec_paint_key(updated, self._kit_keys, status),
            )
            if status == STATUS_WRITING and not scrolled and spec_item is not None:
                self._spec_table.scrollToItem(
                    spec_item,
                    QAbstractItemView.ScrollHint.PositionAtCenter,
                )
                self._spec_table.selectRow(visual)
                scrolled = True
        if changed:
            self._refresh_spec_summary_label()

    def _focus_first_problem(self) -> None:
        for visual in range(self._spec_table.rowCount()):
            if self._spec_table.isRowHidden(visual):
                continue
            row = self._progress_row_at(visual)
            if row is None or row.status != STATUS_SKIPPED:
                continue
            item = self._spec_table.item(visual, _COL_SPEC)
            if item is not None:
                self._spec_table.scrollToItem(
                    item,
                    QAbstractItemView.ScrollHint.PositionAtCenter,
                )
                self._spec_table.selectRow(visual)
            return

    @Slot(int, int)
    def _on_spec_double_clicked(self, row: int, _column: int) -> None:
        progress = self._progress_row_at(row)
        if progress is None:
            return
        if self._store is None:
            QMessageBox.information(
                self,
                "База заказчика",
                "Нет pickle со строками этой спецификации.",
            )
            return
        records = self._store.records_for_spec(progress.spec)
        preview = MtoPreviewDialog(self, spec=progress.spec, records=records)
        self._preview_windows.append(preview)
        preview.destroyed.connect(
            lambda _obj=None, dialog=preview: self._forget_preview(dialog)
        )
        preview.show()

    def _forget_preview(self, dialog: MtoPreviewDialog) -> None:
        self._preview_windows = [
            item for item in self._preview_windows if item is not dialog
        ]

    @Slot()
    def _on_browse(self) -> None:
        chosen, _filter = QFileDialog.getOpenFileName(
            self,
            "Отчёт систем ПИ (xlsb)",
            str(self.xlsb_path().parent if str(self.xlsb_path()) else ""),
            "Excel binary (*.xlsb);;Все файлы (*.*)",
        )
        if chosen:
            self._xlsb_edit.setText(chosen)

    @Slot()
    def _on_reload(self) -> None:
        self._start_job(JOB_RELOAD_XLSB)

    @Slot()
    def _on_catalog(self) -> None:
        self._start_job(JOB_BUILD_CATALOG)

    @Slot()
    def _on_both(self) -> None:
        self._start_job(JOB_BOTH)

    @Slot()
    def _on_open_catalog(self) -> None:
        dest = self._dest_dir
        if not dest.is_dir():
            QMessageBox.warning(
                self,
                "База заказчика",
                f"Папка каталога ещё не создана:\n{dest}\n"
                "Сначала соберите каталог АвтоМТО.",
            )
            return
        ok, message = open_path(dest)
        if not ok:
            QMessageBox.warning(self, "База заказчика", message)

    def _start_job(self, job: str) -> None:
        if self.is_busy():
            return
        xlsb = self.xlsb_path()
        if job in {JOB_RELOAD_XLSB, JOB_BOTH} and not xlsb.is_file():
            QMessageBox.warning(
                self,
                "База заказчика",
                f"Файл xlsb не найден:\n{xlsb}",
            )
            return
        if job == JOB_BUILD_CATALOG and not self._pickle_path.is_file():
            QMessageBox.warning(
                self,
                "База заказчика",
                f"Нет pickle для сборки каталога:\n{self._pickle_path}",
            )
            return
        self._log_view.clear()
        thread = CustomerPiThread(
            job=job,
            xlsb_path=xlsb if job in {JOB_RELOAD_XLSB, JOB_BOTH} else None,
            pickle_path=self._pickle_path,
            dest_dir=self._dest_dir,
            parent=self,
        )
        thread.log.connect(self._append_log)
        thread.error.connect(self._append_log)
        thread.specs_ready.connect(self._on_specs_ready)
        thread.spec_status.connect(self._set_spec_status)
        thread.finished.connect(self._on_finished)
        self._thread = thread
        self._set_busy(True)
        thread.start()

    def _set_busy(self, busy: bool) -> None:
        self._reload_button.setEnabled(not busy)
        self._catalog_button.setEnabled(not busy)
        self._both_button.setEnabled(not busy)
        self._xlsb_edit.setEnabled(not busy)

    @Slot(str)
    def _append_log(self, message: str) -> None:
        self._log_view.appendPlainText(message)

    @Slot(object)
    def _on_specs_ready(self, rows: object) -> None:
        if not isinstance(rows, (tuple, list)):
            return
        typed = tuple(item for item in rows if isinstance(item, SpecProgressRow))
        self._fill_spec_table(typed)

    @Slot()
    def _on_finished(self) -> None:
        thread = self._thread
        self._thread = None
        self._set_busy(False)
        if thread is None:
            return
        thread.deleteLater()
        if thread.failure:
            QMessageBox.warning(self, "База заказчика", thread.failure)
            return
        if thread.store is not None:
            self._store = thread.store
            self._refresh_meta()
            if thread.job in {JOB_RELOAD_XLSB, JOB_BOTH}:
                self.pickle_updated.emit()
        if thread.rebuild is not None:
            self.catalog_rebuilt.emit()
            self._focus_first_problem()

    def closeEvent(self, event: QCloseEvent) -> None:
        """Do not destroy a running xlsb/catalog worker."""

        if self.is_busy() and self._thread is not None:
            self._thread.requestInterruption()
            if not self._thread.wait(4000):
                QMessageBox.information(
                    self,
                    "База заказчика",
                    "Задача ещё идёт. Повторите закрытие через несколько секунд.",
                )
                event.ignore()
                return
        for preview in list(self._preview_windows):
            preview.close()
        super().closeEvent(event)
