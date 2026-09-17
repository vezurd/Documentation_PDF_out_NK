"""RD-tree dump finder tab: MTO xlsx and OD doc/docx under ``rd_root``.

Qt monitor only. Matching stays in ``an_index``; this widget does not walk UNC.
Canonical / layout flags are lexical against the configured RD root.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QPoint, QSettings, Qt, QTimer, Signal, Slot
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
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from rd_catalog.an_compare import (
    AN_AGREED_HEADER,
    AN_AGREED_HEADER_TIP,
    an_agreed_cell_fill,
    an_agreed_cell_tooltip,
    an_agreed_row_fill,
    an_vs_cell_fill,
    an_vs_cell_tooltip,
    format_an_agreed_cell,
    format_an_vs_cell,
)
from rd_catalog.an_compare_cache import (
    RD_DUMP_CONTENT_COMPARE_CACHE_NAME,
    cache_key as an_content_cache_key,
    counterpart_mtime_ns,
    load_an_content_compare_cache,
    result_from_entry,
    result_to_entry,
    save_an_content_compare_cache,
)
from rd_catalog.an_compare_thread import AnContentCompareOutcome, AnContentCompareThread
from rd_catalog.an_index import (
    AnAgreedScore,
    AnKitHit,
    AnMtoFile,
    KitAnTargets,
    match_an_to_kit,
    score_an_files_for_agreed,
)
from rd_catalog.context_menu_qt import exec_tracked_menu
from rd_catalog.context_menu_usage import MENU_RD_DUMP
from rd_catalog.customer_pi_auto_mto import (
    MtoPairCompareResult,
    revision_texts_match,
)
from rd_catalog.kits import kit_identity_key
from rd_catalog.path_actions import open_path
from rd_catalog.rd_dump_index import (
    CANON_HEADER,
    KIND_HEADER,
    KIND_OD,
    LAYOUT_HEADER,
    RD_DUMP_AGREED_ROW_HEADERS,
    RD_DUMP_CANON_NO_FILL,
    RD_DUMP_CANON_YES_FILL,
    RD_DUMP_HEADERS,
    rd_dump_is_canonical,
    rd_dump_is_od,
    rd_dump_kind,
    rd_dump_layout_short,
    rd_dump_layout_token,
    rd_dump_layout_tooltip,
)

_ROLE_ROW = Qt.ItemDataRole.UserRole
_ROLE_SORT = Qt.ItemDataRole.UserRole + 1
_HEADERS = RD_DUMP_HEADERS
_COL_VS_AUTO = RD_DUMP_HEADERS.index("vs Авто МТО")
_COL_VS_RD = RD_DUMP_HEADERS.index("vs MTO РД")
_COL_DATE = RD_DUMP_HEADERS.index("Дата")
_COL_AGREED = RD_DUMP_HEADERS.index(AN_AGREED_HEADER)
_COL_KIND = RD_DUMP_HEADERS.index(KIND_HEADER)
_COL_CANON = RD_DUMP_HEADERS.index(CANON_HEADER)
_COL_LAYOUT = RD_DUMP_HEADERS.index(LAYOUT_HEADER)
_VS_AUTO_HEADER_TIP = (
    "Ревизия в имени (да/нет) и сверка содержимого с файлом Авто МТО из ПИ."
)
_VS_RD_HEADER_TIP = (
    "Ревизия в имени (да/нет) и сверка содержимого с текущим MTO РД каталога."
)
_CANON_HEADER_TIP = (
    "Канонический путь РД: титул / марка / шлюз / NN_ / [PDF|DWG]. "
    "Неканоничные файлы каталог не хранит — эта вкладка их показывает."
)
_LAYOUT_HEADER_TIP = "Почему путь не канонический (или «канон»)."
_KIND_HEADER_TIP = "MTO — xlsx; OD — ведомость документов (.doc / .docx)."


@dataclass(frozen=True, slots=True)
class _RdDumpTabRow:
    """One table row: the file plus kit-level match and layout flags."""

    file: AnMtoFile
    hit: AnKitHit
    targets: KitAnTargets
    is_canonical: bool
    layout_token: str
    agreed_score: AnAgreedScore | None = None


@dataclass(frozen=True, slots=True)
class _RdDumpContentPair:
    """Cached pairwise content results for one dump workbook."""

    vs_auto: MtoPairCompareResult | None = None
    vs_rd: MtoPairCompareResult | None = None


@dataclass(frozen=True, slots=True)
class _RdDumpContentJob:
    """One queued dump file plus the kit counterpart paths."""

    dump_path: str
    auto_path: str
    rd_path: str


class _SortItem(QTableWidgetItem):
    """Table item that sorts by ``_ROLE_SORT`` when both sides have it."""

    def __lt__(self, other: QTableWidgetItem) -> bool:
        left = self.data(_ROLE_SORT)
        right = other.data(_ROLE_SORT) if other is not None else None
        if left is not None and right is not None:
            return bool(left < right)
        return super().__lt__(other)


class RdDumpTab(QWidget):
    """Filterable RD MTO dump list with kit jump and a scan button."""

    kit_activated = Signal(str, str)
    prepare_context_menu = Signal(QMenu)
    scan_requested = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._rows: tuple[_RdDumpTabRow, ...] = ()
        self._is_banned: Callable[[str, str], bool] = lambda _t, _m: False
        self._kit_targets: Mapping[tuple[str, str], KitAnTargets] = {}
        self._allowed_kits: set[tuple[str, str]] | None = None
        self._rd_root: str = ""
        self._runtime_dir: Path | None = None
        self._content_by_key: dict[str, _RdDumpContentPair] = {}
        self._content_queue: list[_RdDumpContentJob] = []
        self._content_queue_paused = False
        self._content_thread: AnContentCompareThread | None = None
        self._content_running_path = ""
        self._content_done = 0
        self._content_total = 0
        self._build()

    def configure(self, runtime_dir: str | Path | None) -> None:
        """Set the runtime directory used for the content-compare cache.

        Args:
            runtime_dir: Catalog runtime directory, or ``None`` to skip persist.
        """

        self._runtime_dir = Path(runtime_dir) if runtime_dir else None

    def prepare_close(self, *, wait_ms: int = 3000) -> bool:
        """Cancel the content-compare queue and wait for the worker.

        Args:
            wait_ms: Milliseconds to wait for the running job.

        Returns:
            True when the worker is idle.
        """

        self._content_queue.clear()
        self._content_queue_paused = True
        thread = self._content_thread
        if thread is None or not thread.isRunning():
            return True
        return bool(thread.wait(wait_ms))

    def set_content_queue_paused(self, paused: bool) -> None:
        """Pause or resume background dump content compares.

        Args:
            paused: True while a catalog worker (scan / Google / …) is busy.
        """

        was_paused = self._content_queue_paused
        self._content_queue_paused = bool(paused)
        if was_paused and not self._content_queue_paused:
            self._kick_content_queue()

    def table(self) -> QTableWidget:
        """Return the dump files table."""

        return self._table

    def selected_file(self) -> AnMtoFile | None:
        """Return the dump workbook on the current table selection."""

        payload = self._row_at(self._table.currentRow())
        return payload.file if payload is not None else None

    def set_scan_enabled(self, enabled: bool) -> None:
        """Enable or disable the «Сканировать РД» button."""

        self._scan_button.setEnabled(enabled)

    def set_rows(
        self,
        rows: Sequence[AnMtoFile],
        *,
        is_banned: Callable[[str, str], bool],
        kit_targets: Mapping[tuple[str, str], KitAnTargets],
        rd_root: str | Path,
        allowed_kits: set[tuple[str, str]] | None = None,
    ) -> None:
        """Replace the RD dump snapshot and rebuild the table.

        Args:
            rows: Flat dump (all kits).
            is_banned: Hide banned ``(title, mark)`` pairs.
            kit_targets: Filename-revision targets keyed by
                ``kit_identity_key``.
            rd_root: Catalog RD root used for canonical / layout flags.
            allowed_kits: Known Комплекты identities. Used only by the
                «нет в Комплектах» filter. ``None`` means no kit is known.
        """

        self._is_banned = is_banned
        self._kit_targets = kit_targets
        self._allowed_kits = allowed_kits
        self._rd_root = str(rd_root or "")
        grouped: dict[tuple[str, str], list[AnMtoFile]] = {}
        for file in rows:
            grouped.setdefault(
                kit_identity_key(file.title, file.mark), []
            ).append(file)
        hits: dict[tuple[str, str], AnKitHit] = {}
        scores: dict[tuple[str, str], dict[str, AnAgreedScore]] = {}
        for key, files in grouped.items():
            targets = kit_targets.get(key, KitAnTargets())
            mto_files = tuple(item for item in files if not rd_dump_is_od(item))
            hits[key] = match_an_to_kit(mto_files, targets)
            scores[key] = score_an_files_for_agreed(tuple(files), targets)
        built: list[_RdDumpTabRow] = []
        for file in rows:
            key = kit_identity_key(file.title, file.mark)
            token = rd_dump_layout_token(file.path, self._rd_root)
            built.append(
                _RdDumpTabRow(
                    file=file,
                    hit=hits[key],
                    targets=kit_targets.get(key, KitAnTargets()),
                    is_canonical=rd_dump_is_canonical(file.path, self._rd_root),
                    layout_token=token,
                    agreed_score=scores.get(key, {}).get(file.path_key),
                )
            )
        self._rows = tuple(built)
        self._rebuild_table()
        QTimer.singleShot(0, self._after_dump_rows_painted)

    def _after_dump_rows_painted(self) -> None:
        """Hydrate content-compare cache after the table is already visible."""

        self._hydrate_content_cache()
        table = self._table
        for index in range(table.rowCount()):
            payload = self._row_at(index)
            if payload is not None:
                self._fill_vs_items(index, payload)
        self.enqueue_needed_content_compares()

    def set_kit_filter(self, title: str, mark: str) -> None:
        """Show only files of this title–mark via the text filter.

        Clears «нет в Комплектах» so a jump from Комплекты is visible.
        Always sorts by **Дата** (file mtime, oldest first).
        """

        if self._not_in_kits.isChecked():
            self._not_in_kits.blockSignals(True)
            self._not_in_kits.setChecked(False)
            self._not_in_kits.blockSignals(False)
        self._filter.blockSignals(True)
        self._filter.setText(f"{title}-{mark}")
        self._filter.blockSignals(False)
        self._table.sortByColumn(_COL_DATE, Qt.SortOrder.AscendingOrder)
        self._apply_row_visibility()

    def focus_revision(self, rev: str = "") -> None:
        """Select the first matching revision row, preferring visible rows."""

        table = self._table
        target = (rev or "").strip().casefold()
        first_visible: int | None = None
        first_any: int | None = None
        for index in range(table.rowCount()):
            payload = self._row_at(index)
            if payload is None:
                continue
            if first_any is None:
                first_any = index
            hidden = table.isRowHidden(index)
            if first_visible is None and not hidden:
                first_visible = index
            if target and payload.file.revision_text.casefold() != target:
                continue
            if hidden and first_visible is not None:
                continue
            table.selectRow(index)
            item = table.item(index, 0)
            if item is not None:
                table.scrollToItem(item)
            return
        chosen = first_visible if first_visible is not None else first_any
        if chosen is None:
            return
        table.selectRow(chosen)
        item = table.item(chosen, 0)
        if item is not None:
            table.scrollToItem(item)

    def focus_best_agreed(self) -> bool:
        """Select the highlighted agreed-transfer row if one is visible."""

        table = self._table
        first_hidden: int | None = None
        for index in range(table.rowCount()):
            payload = self._row_at(index)
            if payload is None or payload.agreed_score is None:
                continue
            if not payload.agreed_score.is_best:
                continue
            if table.isRowHidden(index):
                if first_hidden is None:
                    first_hidden = index
                continue
            table.selectRow(index)
            item = table.item(index, 0)
            if item is not None:
                table.scrollToItem(item)
            return True
        if first_hidden is None:
            return False
        table.selectRow(first_hidden)
        item = table.item(first_hidden, 0)
        if item is not None:
            table.scrollToItem(item)
        return True

    def restore_filters(self, settings: QSettings) -> None:
        """Load filter widgets from QSettings without applying twice."""

        boxes = self._filter_checkboxes()
        self._filter.blockSignals(True)
        for box in boxes:
            box.blockSignals(True)
        try:
            self._filter.setText(
                str(settings.value("window/rd_dump_tab_filter") or "")
            )
            self._closes.setChecked(
                _settings_bool(
                    settings, "window/rd_dump_tab_closes_auto_mto", False
                )
            )
            self._no_rd_match.setChecked(
                _settings_bool(
                    settings, "window/rd_dump_tab_no_rd_match", False
                )
            )
            self._not_in_kits.setChecked(
                _settings_bool(
                    settings, "window/rd_dump_tab_not_in_kits", False
                )
            )
            self._noncanonical.setChecked(
                _settings_bool(
                    settings, "window/rd_dump_tab_noncanonical", False
                )
            )
            self._only_od.setChecked(
                _settings_bool(settings, "window/rd_dump_tab_only_od", False)
            )
        finally:
            self._filter.blockSignals(False)
            for box in boxes:
                box.blockSignals(False)

    def save_filters(self, settings: QSettings) -> None:
        """Persist filter widgets into QSettings."""

        settings.setValue("window/rd_dump_tab_filter", self._filter.text())
        settings.setValue(
            "window/rd_dump_tab_closes_auto_mto", self._closes.isChecked()
        )
        settings.setValue(
            "window/rd_dump_tab_no_rd_match", self._no_rd_match.isChecked()
        )
        settings.setValue(
            "window/rd_dump_tab_not_in_kits", self._not_in_kits.isChecked()
        )
        settings.setValue(
            "window/rd_dump_tab_noncanonical", self._noncanonical.isChecked()
        )
        settings.setValue("window/rd_dump_tab_only_od", self._only_od.isChecked())

    def showEvent(self, event) -> None:
        """Start pending content compares when the tab becomes visible."""

        super().showEvent(event)
        self.enqueue_needed_content_compares()

    def enqueue_needed_content_compares(self) -> None:
        """Queue dump files that still need a PI / RD content compare."""

        if not self.isVisible():
            return
        queued = {job.dump_path.casefold() for job in self._content_queue}
        if self._content_running_path:
            queued.add(self._content_running_path.casefold())
        added = 0
        for row in self._rows:
            if self._kit_is_hidden(row.file.title, row.file.mark):
                continue
            if rd_dump_is_od(row.file):
                continue
            auto_path = str(row.targets.auto_mto_path or "").strip()
            rd_path = str(row.targets.rd_mto_path or "").strip()
            if not auto_path and not rd_path:
                continue
            if self._content_pair_is_current(row):
                continue
            path = row.file.path
            if path.casefold() in queued:
                continue
            self._content_queue.append(
                _RdDumpContentJob(
                    dump_path=path, auto_path=auto_path, rd_path=rd_path
                )
            )
            queued.add(path.casefold())
            added += 1
        if added:
            self._content_total += added
        self._kick_content_queue()

    def _content_pair_is_current(self, row: _RdDumpTabRow) -> bool:
        pair = self._content_by_key.get(row.file.path_key)
        if pair is None:
            return False
        auto_path = str(row.targets.auto_mto_path or "").strip()
        rd_path = str(row.targets.rd_mto_path or "").strip()
        if auto_path and pair.vs_auto is None:
            return False
        if rd_path and pair.vs_rd is None:
            return False
        return True

    def _hydrate_content_cache(self) -> None:
        disk = load_an_content_compare_cache(
            self._runtime_dir, name=RD_DUMP_CONTENT_COMPARE_CACHE_NAME
        )
        kept: dict[str, _RdDumpContentPair] = {}
        mtimes: dict[str, int] = {}

        def counterpart_mtime(path: str) -> int:
            text = str(path or "").strip()
            if not text:
                return 0
            key = text.casefold()
            if key not in mtimes:
                mtimes[key] = counterpart_mtime_ns(text)
            return mtimes[key]

        for row in self._rows:
            key = an_content_cache_key(
                row.file.path,
                row.file.mtime_ns,
                row.targets.auto_mto_path,
                counterpart_mtime(row.targets.auto_mto_path),
                row.targets.rd_mto_path,
                counterpart_mtime(row.targets.rd_mto_path),
            )
            entry = disk.get(key)
            if not entry:
                continue
            kept[row.file.path_key] = _RdDumpContentPair(
                vs_auto=result_from_entry(entry.get("vs_auto")),
                vs_rd=result_from_entry(entry.get("vs_rd")),
            )
        self._content_by_key = kept

    def _persist_content_pair(
        self, row: _RdDumpTabRow, pair: _RdDumpContentPair
    ) -> None:
        if self._runtime_dir is None:
            return
        entries = load_an_content_compare_cache(
            self._runtime_dir, name=RD_DUMP_CONTENT_COMPARE_CACHE_NAME
        )
        key = an_content_cache_key(
            row.file.path,
            row.file.mtime_ns,
            row.targets.auto_mto_path,
            counterpart_mtime_ns(row.targets.auto_mto_path),
            row.targets.rd_mto_path,
            counterpart_mtime_ns(row.targets.rd_mto_path),
        )
        entries[key] = {
            "vs_auto": result_to_entry(pair.vs_auto),
            "vs_rd": result_to_entry(pair.vs_rd),
        }
        save_an_content_compare_cache(
            self._runtime_dir,
            entries,
            name=RD_DUMP_CONTENT_COMPARE_CACHE_NAME,
        )

    def _kick_content_queue(self) -> None:
        if self._content_queue_paused:
            self._update_content_queue_label()
            return
        if self._content_thread is not None and self._content_thread.isRunning():
            self._update_content_queue_label()
            return
        if not self._content_queue:
            self._content_done = 0
            self._content_total = 0
            self._update_content_queue_label()
            return
        job = self._content_queue.pop(0)
        thread = AnContentCompareThread(
            job.dump_path,
            auto_path=job.auto_path,
            rd_path=job.rd_path,
            parent=self,
        )
        thread.result_ready.connect(self._on_content_ready)
        thread.error.connect(self._on_content_error)
        thread.finished.connect(self._on_content_finished)
        self._content_thread = thread
        self._content_running_path = job.dump_path
        self._update_content_queue_label()
        thread.start()

    def _update_content_queue_label(self) -> None:
        label = getattr(self, "_content_queue_label", None)
        if label is None:
            return
        if self._content_thread is not None and self._content_thread.isRunning():
            label.setText(
                f"РД сверка: {self._content_done}/{max(self._content_total, 1)}"
            )
            return
        if self._content_queue:
            label.setText(
                f"РД сверка: в очереди {len(self._content_queue)}"
            )
            return
        label.setText("РД сверка: актуально")

    @Slot(object)
    def _on_content_ready(self, outcome: object) -> None:
        if not isinstance(outcome, AnContentCompareOutcome):
            return
        self._content_done += 1
        row = next(
            (
                item
                for item in self._rows
                if item.file.path.casefold() == outcome.an_path.casefold()
            ),
            None,
        )
        if row is None:
            return
        pair = _RdDumpContentPair(vs_auto=outcome.vs_auto, vs_rd=outcome.vs_rd)
        self._content_by_key[row.file.path_key] = pair
        self._persist_content_pair(row, pair)
        self._repaint_content_cells(row)

    @Slot(str)
    def _on_content_error(self, message: str) -> None:
        label = getattr(self, "_content_queue_label", None)
        if label is not None:
            label.setText(f"РД сверка: {message[:80]}")

    @Slot()
    def _on_content_finished(self) -> None:
        thread = self._content_thread
        self._content_thread = None
        self._content_running_path = ""
        if thread is not None:
            thread.deleteLater()
        self._kick_content_queue()

    def _repaint_content_cells(self, row: _RdDumpTabRow) -> None:
        table = self._table
        for index in range(table.rowCount()):
            payload = self._row_at(index)
            if payload is None or payload.file.path_key != row.file.path_key:
                continue
            self._fill_vs_items(index, payload)
            return

    def _build(self) -> None:
        layout = QVBoxLayout(self)
        counts_row = QHBoxLayout()
        self._counts = QLabel("Строки: 0 / 0", self)
        self._content_queue_label = QLabel("РД сверка: актуально", self)
        counts_row.addWidget(self._counts, 1)
        counts_row.addWidget(self._content_queue_label)
        layout.addLayout(counts_row)

        filters = QHBoxLayout()
        self._scan_button = QPushButton("Сканировать РД", self)
        self._scan_button.setToolTip(
            "Обойти всё дерево РД и собрать AGCC MTO xlsx и OD doc/docx, "
            "включая неканоничные пути. Skip-папки как у каталога. "
            "В file_entry / «Все документы» не пишет."
        )
        self._scan_button.clicked.connect(self.scan_requested.emit)
        filters.addWidget(self._scan_button)
        filters.addWidget(QLabel("Фильтр:"))
        self._filter = QLineEdit(self)
        self._filter.setPlaceholderText(
            "Титул, марка, ревизия, путь, раскладка…"
        )
        self._filter.textChanged.connect(self._apply_row_visibility)
        filters.addWidget(self._filter, 1)
        self._noncanonical = QCheckBox("только неканоничные", self)
        self._noncanonical.setToolTip(
            "Скрыть файлы в титул/марка/шлюз/NN_. "
            "Оставить копии вне выданного контура."
        )
        self._only_od = QCheckBox("только OD", self)
        self._only_od.setToolTip("Только ведомость документов (.doc / .docx).")
        self._closes = QCheckBox("закрывают Авто МТО", self)
        self._closes.setToolTip(
            "Комплекты, где файлы дампа закрывают расхождение Авто МТО с MTO РД."
        )
        self._no_rd_match = QCheckBox(
            "есть в дампе, нет совпадения с MTO РД", self
        )
        self._not_in_kits = QCheckBox("нет в Комплектах", self)
        for box in (
            self._noncanonical,
            self._only_od,
            self._closes,
            self._no_rd_match,
            self._not_in_kits,
        ):
            box.toggled.connect(self._apply_row_visibility)
            filters.addWidget(box)
        layout.addLayout(filters)

        self._table = QTableWidget(0, len(_HEADERS), self)
        self._table.setHorizontalHeaderLabels(list(_HEADERS))
        auto_header = self._table.horizontalHeaderItem(_COL_VS_AUTO)
        if auto_header is not None:
            auto_header.setToolTip(_VS_AUTO_HEADER_TIP)
        rd_header = self._table.horizontalHeaderItem(_COL_VS_RD)
        if rd_header is not None:
            rd_header.setToolTip(_VS_RD_HEADER_TIP)
        agreed_header = self._table.horizontalHeaderItem(_COL_AGREED)
        if agreed_header is not None:
            agreed_header.setToolTip(AN_AGREED_HEADER_TIP)
        canon_header = self._table.horizontalHeaderItem(_COL_CANON)
        if canon_header is not None:
            canon_header.setToolTip(_CANON_HEADER_TIP)
        kind_header = self._table.horizontalHeaderItem(_COL_KIND)
        if kind_header is not None:
            kind_header.setToolTip(_KIND_HEADER_TIP)
        layout_header = self._table.horizontalHeaderItem(_COL_LAYOUT)
        if layout_header is not None:
            layout_header.setToolTip(_LAYOUT_HEADER_TIP)
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

    def _filter_checkboxes(self) -> tuple[QCheckBox, ...]:
        return (
            self._noncanonical,
            self._only_od,
            self._closes,
            self._no_rd_match,
            self._not_in_kits,
        )

    def _kit_is_hidden(self, title: str, mark: str) -> bool:
        return self._is_banned(title, mark)

    def _rebuild_table(self) -> None:
        table = self._table
        table.setSortingEnabled(False)
        visible_rows = [
            row
            for row in self._rows
            if not self._kit_is_hidden(row.file.title, row.file.mark)
        ]
        table.setRowCount(len(visible_rows))
        for row_index, row in enumerate(visible_rows):
            self._fill_row(row_index, row)
        table.setSortingEnabled(True)
        self._apply_row_visibility()

    def _fill_row(self, row_index: int, row: _RdDumpTabRow) -> None:
        file = row.file
        targets = row.targets
        vs_auto, vs_rd = self._vs_cell_texts(row)
        agreed_text = format_an_agreed_cell(row.agreed_score)
        kind_text = rd_dump_kind(file)
        canon_text = "да" if row.is_canonical else "нет"
        layout_text = rd_dump_layout_short(row.layout_token)
        values = {
            "Титул": file.title,
            "Марка": file.mark,
            "Ревизия": file.revision_text or "—",
            KIND_HEADER: kind_text,
            CANON_HEADER: canon_text,
            LAYOUT_HEADER: layout_text,
            AN_AGREED_HEADER: agreed_text,
            "vs Авто МТО": vs_auto,
            "vs MTO РД": vs_rd,
            "vs Робот": _yes_no_dash(
                revision_texts_match(file.revision_text, targets.robot)
            ),
            "vs Выдача": _yes_no_dash(
                revision_texts_match(file.revision_text, targets.issuance)
            ),
            "vs F": _yes_no_dash(
                revision_texts_match(file.revision_text, targets.google_f)
            ),
            "vs SQ": _yes_no_dash(
                revision_texts_match(file.revision_text, targets.sq)
            ),
            "Имя": file.name,
            "Дата": _format_mtime(file.mtime_ns),
            "Папка": file.parent_dir,
            "Путь": file.path,
        }
        table = self._table
        for column, header in enumerate(_HEADERS):
            text = values[header]
            if column == _COL_DATE:
                item: QTableWidgetItem = _SortItem(text)
                item.setData(_ROLE_SORT, file.mtime_ns)
            elif column == _COL_AGREED:
                item = _SortItem(text)
                percent = (
                    row.agreed_score.percent
                    if row.agreed_score is not None
                    else None
                )
                item.setData(
                    _ROLE_SORT, percent if percent is not None else -1
                )
            elif column == _COL_CANON:
                item = _SortItem(text)
                item.setData(_ROLE_SORT, 0 if row.is_canonical else 1)
            elif column == _COL_KIND:
                item = _SortItem(text)
                item.setData(_ROLE_SORT, 0 if kind_text == KIND_OD else 1)
            else:
                item = QTableWidgetItem(text)
            item.setData(_ROLE_ROW, row)
            table.setItem(row_index, column, item)
        self._fill_vs_items(row_index, row)
        self._paint_layout_cells(row_index, row)
        self._paint_agreed_row(row_index, row)

    def _paint_layout_cells(self, row_index: int, row: _RdDumpTabRow) -> None:
        fill = (
            RD_DUMP_CANON_YES_FILL
            if row.is_canonical
            else RD_DUMP_CANON_NO_FILL
        )
        canon_item = self._table.item(row_index, _COL_CANON)
        if canon_item is not None:
            canon_item.setBackground(QBrush(QColor(fill)))
            canon_item.setToolTip(_CANON_HEADER_TIP)
        layout_item = self._table.item(row_index, _COL_LAYOUT)
        if layout_item is not None:
            layout_item.setBackground(QBrush(QColor(fill)))
            layout_item.setToolTip(rd_dump_layout_tooltip(row.layout_token))

    def _vs_cell_texts(self, row: _RdDumpTabRow) -> tuple[str, str]:
        file = row.file
        targets = row.targets
        pair = self._content_by_key.get(file.path_key)
        compare_content = not rd_dump_is_od(file)
        vs_auto = format_an_vs_cell(
            revision_texts_match(file.revision_text, targets.auto_mto),
            pair.vs_auto if pair is not None else None,
            has_counterpart=compare_content and bool(targets.auto_mto_path),
        )
        vs_rd = format_an_vs_cell(
            revision_texts_match(file.revision_text, targets.rd_mto),
            pair.vs_rd if pair is not None else None,
            has_counterpart=compare_content and bool(targets.rd_mto_path),
        )
        return vs_auto, vs_rd

    def _fill_vs_items(self, row_index: int, row: _RdDumpTabRow) -> None:
        file = row.file
        targets = row.targets
        pair = self._content_by_key.get(file.path_key)
        vs_auto, vs_rd = self._vs_cell_texts(row)
        auto_match = revision_texts_match(file.revision_text, targets.auto_mto)
        rd_match = revision_texts_match(file.revision_text, targets.rd_mto)
        auto_item = self._table.item(row_index, _COL_VS_AUTO)
        rd_item = self._table.item(row_index, _COL_VS_RD)
        if auto_item is not None:
            auto_item.setText(vs_auto)
            auto_item.setData(_ROLE_ROW, row)
            _paint_vs_match_cell(
                auto_item,
                auto_match,
                pair.vs_auto if pair is not None else None,
            )
            auto_item.setToolTip(
                an_vs_cell_tooltip(
                    rev_match=auto_match,
                    content=pair.vs_auto if pair is not None else None,
                    an_path=file.path,
                    other_path=targets.auto_mto_path,
                    other_label="ПИ / Авто МТО",
                )
            )
        if rd_item is not None:
            rd_item.setText(vs_rd)
            rd_item.setData(_ROLE_ROW, row)
            _paint_vs_match_cell(
                rd_item,
                rd_match,
                pair.vs_rd if pair is not None else None,
            )
            rd_item.setToolTip(
                an_vs_cell_tooltip(
                    rev_match=rd_match,
                    content=pair.vs_rd if pair is not None else None,
                    an_path=file.path,
                    other_path=targets.rd_mto_path,
                    other_label="MTO РД",
                )
            )

    def _paint_agreed_row(self, row_index: int, row: _RdDumpTabRow) -> None:
        score = row.agreed_score
        agreed_item = self._table.item(row_index, _COL_AGREED)
        if agreed_item is not None:
            agreed_item.setText(format_an_agreed_cell(score))
            agreed_item.setToolTip(an_agreed_cell_tooltip(score))
            agreed_item.setData(_ROLE_ROW, row)
            fill = an_agreed_cell_fill(score)
            if fill:
                agreed_item.setBackground(QBrush(QColor(fill)))
            if score is not None and score.is_best:
                font = QFont(agreed_item.font())
                font.setBold(True)
                agreed_item.setFont(font)
        row_fill = an_agreed_row_fill(score)
        if not row_fill:
            return
        for column, header in enumerate(_HEADERS):
            if header not in RD_DUMP_AGREED_ROW_HEADERS or header == AN_AGREED_HEADER:
                continue
            if header in (CANON_HEADER, LAYOUT_HEADER):
                continue
            item = self._table.item(row_index, column)
            if item is None:
                continue
            item.setBackground(QBrush(QColor(row_fill)))

    def _row_at(self, row_index: int) -> _RdDumpTabRow | None:
        if row_index < 0:
            return None
        item = self._table.item(row_index, 0)
        payload = item.data(_ROLE_ROW) if item else None
        return payload if isinstance(payload, _RdDumpTabRow) else None

    def _row_matches_text(self, row: _RdDumpTabRow, needle: str) -> bool:
        if not needle:
            return True
        file = row.file
        vs_auto, vs_rd = self._vs_cell_texts(row)
        chunks = [
            file.title,
            file.mark,
            file.revision_text,
            rd_dump_kind(file),
            "да" if row.is_canonical else "нет",
            "канон" if row.is_canonical else "неканон",
            rd_dump_layout_short(row.layout_token),
            rd_dump_layout_tooltip(row.layout_token),
            format_an_agreed_cell(row.agreed_score),
            vs_auto,
            vs_rd,
            file.name,
            file.parent_dir,
            file.path,
            "закрывает" if row.hit.closes_auto_mto else "",
            "лучше" if row.agreed_score is not None and row.agreed_score.is_best else "",
        ]
        return needle in " ".join(chunks).casefold()

    def _row_matches_filters(self, row: _RdDumpTabRow) -> bool:
        if self._noncanonical.isChecked() and row.is_canonical:
            return False
        if self._only_od.isChecked() and not rd_dump_is_od(row.file):
            return False
        if self._closes.isChecked() and (
            rd_dump_is_od(row.file) or not row.hit.closes_auto_mto
        ):
            return False
        if self._no_rd_match.isChecked() and (
            rd_dump_is_od(row.file)
            or not (row.hit.files and row.hit.match_rd_mto is not True)
        ):
            return False
        if self._not_in_kits.isChecked():
            key = kit_identity_key(row.file.title, row.file.mark)
            if self._allowed_kits is not None and key in self._allowed_kits:
                return False
        return True

    @Slot()
    def _apply_row_visibility(self) -> None:
        needle = self._filter.text().strip().casefold()
        table = self._table
        shown = 0
        for index in range(table.rowCount()):
            payload = self._row_at(index)
            if payload is None or self._kit_is_hidden(
                payload.file.title, payload.file.mark
            ):
                table.setRowHidden(index, True)
                continue
            visible = self._row_matches_filters(
                payload
            ) and self._row_matches_text(payload, needle)
            table.setRowHidden(index, not visible)
            if visible:
                shown += 1
        self._counts.setText(f"Строки: {shown} / {table.rowCount()}")

    @Slot(int, int)
    def _on_cell_activated(self, row: int, _column: int) -> None:
        payload = self._row_at(row)
        if payload is None:
            return
        if payload.file.parent_dir:
            open_path(payload.file.parent_dir)

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
        file = payload.file
        menu = QMenu(self)
        open_folder = menu.addAction("Открыть папку")
        copy_path = menu.addAction("Копировать путь")
        copy_folder = menu.addAction("Копировать папку")
        open_folder.setEnabled(bool(file.parent_dir))
        copy_path.setEnabled(bool(file.path))
        copy_folder.setEnabled(bool(file.parent_dir))
        menu.addSeparator()
        show_kits = menu.addAction('Показать в «Комплекты»')
        self.prepare_context_menu.emit(menu)
        chosen = exec_tracked_menu(
            menu, MENU_RD_DUMP, table.viewport().mapToGlobal(position)
        )
        if chosen == open_folder:
            open_path(file.parent_dir)
        elif chosen == copy_path:
            _copy_text(file.path)
        elif chosen == copy_folder:
            _copy_text(file.parent_dir)
        elif chosen == show_kits:
            self.kit_activated.emit(file.title, file.mark)


def _yes_no_dash(match: bool | None) -> str:
    if match is True:
        return "да"
    if match is False:
        return "нет"
    return "—"


def _paint_vs_match_cell(
    item: QTableWidgetItem,
    match: bool | None,
    content: MtoPairCompareResult | None = None,
) -> None:
    fill = an_vs_cell_fill(match, content)
    if fill:
        item.setBackground(QBrush(QColor(fill)))
    if match is True:
        font = QFont(item.font())
        font.setBold(True)
        item.setFont(font)


def _format_mtime(value: int | None) -> str:
    if value is None:
        return ""
    try:
        return datetime.fromtimestamp(value / 1_000_000_000).strftime(
            "%Y.%m.%d"
        )
    except (OSError, OverflowError, ValueError):
        return ""


def _copy_text(text: str) -> None:
    if text:
        QApplication.clipboard().setText(text)


def _settings_bool(settings: QSettings, key: str, default: bool) -> bool:
    value = settings.value(key, default)
    if isinstance(value, bool):
        return bool(value)
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().casefold()
    if text in {"1", "true", "yes", "on"}:
        return True
    if text in {"0", "false", "no", "off", ""}:
        return False
    return default
