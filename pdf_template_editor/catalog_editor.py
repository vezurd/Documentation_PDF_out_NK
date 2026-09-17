"""Редактор каталога полей проекта.

Два класса:
  - CatalogEditorPanel(QWidget)  -- встроенная нижняя панель (T7.2+)
  - CatalogEditorDialog(QDialog) -- legacy модальный диалог (совместимость)

Колонки таблицы:
  COL_NAV  [0] -- кнопка навигации к полю
  COL_ID   [1] -- id поля
  COL_LABEL[2] -- label (источник для шаблона рядом с catalog.json)
  Колонки expected / expected_text — только для правки в UI (синхронизация с шаблоном),
  в JSON каталога не сохраняются (как clean — источник истины в шаблоне).
"""

from __future__ import annotations

import os
from PySide6.QtCore import QPoint, QSize, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListView,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from pdf_parsing_v2_engine.document_properties.registry import MANDATORY_PROPERTY_KEYS
from pdf_parsing_v2_engine.models import (
    FieldCatalog,
    FieldCatalogEntry,
    FieldDef,
    _FIELD_TYPE_VALUES,
    _EXPECTED_VALUES,
    normalize_field_type,
    normalize_expected,
)

_GROUP_OPTIONS = ["", "main", "signatures", "revision_table", "labels", "external"]
_FIELD_TYPE_OPTIONS = sorted(_FIELD_TYPE_VALUES)
_EXPECTED_OPTIONS = sorted(_EXPECTED_VALUES)


def _scene_field_lookup(
    scene_fields: list[FieldDef] | None, fid: str,
) -> FieldDef | None:
    if not scene_fields or not fid:
        return None
    for fd in scene_fields:
        if fd.id == fid:
            return fd
    return None


def _read_expected_from_table_row(
    table: QTableWidget, col_exp: int, col_et: int, row: int,
) -> tuple[str, str]:
    w = table.cellWidget(row, col_exp)
    exp = w.currentText().strip() if isinstance(w, QComboBox) else ""
    exp = exp or "optional"
    et_item = table.item(row, col_et)
    et = et_item.text().strip() if et_item else ""
    return exp, et


# Column indices (with new nav button column at 0)
_COL_NAV = 0
_COL_ID = 1
_COL_LABEL = 2
_COL_GROUP = 3
_COL_SHORT_NUM = 4
_COL_FIELD_TYPE = 5
_COL_EXPECTED = 6
_COL_EXPECTED_TEXT = 7
_COL_DESCRIPTION = 8
_COL_REPORT_HEADER_NOTE = 9
_NCOLS = 10

_HEADERS = [
    "", "id", "label", "group", "#", "field_type",
    "expected", "expected_text", "description", "заголовок отчёта",
]

# Новое поле с каталога: компактный прямоугольник у внутреннего угла рамки (origin=frame_bottom_right).
_DEFAULT_CATALOG_NEW_FIELD_BBOX_MM = (2.0, 2.0, 24.0, 8.0)

# id cell: canvas field id this row syncs to while cell text may differ (rename pending).
# Survives row reorder/sort (unlike dict[int] keyed by row index).
USER_ROLE_SYNC_LINK = Qt.ItemDataRole.UserRole + 77
# Last committed unique id for the row (strip); used to revert duplicate id edits.
USER_ROLE_LAST_VALID_ID = Qt.ItemDataRole.UserRole + 78
# Порядок / включение колонок отладочного Excel v2_report (не в JSON ячейки id).
USER_ROLE_DEBUG_REPORT_ORDER = Qt.ItemDataRole.UserRole + 80
USER_ROLE_INCLUDE_DEBUG_REPORT = Qt.ItemDataRole.UserRole + 81
# Скрытое поле свойств PDF (61…file_name) — в JSON каталога как is_document_property.
USER_ROLE_IS_DOCUMENT_PROPERTY = Qt.ItemDataRole.UserRole + 82


def _init_catalog_id_last_valid(item: QTableWidgetItem | None) -> None:
    if item is None:
        return
    item.setData(USER_ROLE_LAST_VALID_ID, item.text().strip())


def _validate_catalog_table_unique_ids(table: QTableWidget, col_id: int) -> None:
    """Raise ValueError if two non-empty id cells share the same text (after strip)."""
    seen: dict[str, int] = {}
    for row in range(table.rowCount()):
        it = table.item(row, col_id)
        if not it:
            continue
        key = it.text().strip()
        if not key:
            continue
        if key in seen:
            raise ValueError(
                f"В каталоге дублируется id «{key}» (строки {seen[key] + 1} и {row + 1}). "
                "Исправьте таблицу перед сохранением."
            )
        seen[key] = row


def _enforce_unique_catalog_id_row(
    table: QTableWidget,
    col_id: int,
    row: int,
    parent: QWidget,
) -> bool:
    """If id duplicates another row, revert cell text and warn. Returns True if reverted."""
    it = table.item(row, col_id)
    if it is None:
        return False
    new_key = it.text().strip()
    if not new_key:
        it.setData(USER_ROLE_LAST_VALID_ID, "")
        return False
    for r in range(table.rowCount()):
        if r == row:
            continue
        other = table.item(r, col_id)
        if other and other.text().strip() == new_key:
            prev = it.data(USER_ROLE_LAST_VALID_ID)
            prev_s = str(prev) if prev is not None else ""
            table.blockSignals(True)
            try:
                it.setText(prev_s)
            finally:
                table.blockSignals(False)
            QMessageBox.warning(
                parent,
                "Каталог",
                f"Идентификатор «{new_key}» уже используется в другой строке каталога.\n"
                "Значение откатано — у каждого поля должен быть уникальный id.",
            )
            return True
    it.setData(USER_ROLE_LAST_VALID_ID, new_key)
    return False


class NumericSortTableWidgetItem(QTableWidgetItem):
    """Сортировка столбца «#» по числу (1, 2, … 10), не лексикографически."""

    def __lt__(self, other: QTableWidgetItem) -> bool:
        tw = self.tableWidget()
        col = tw.sortColumn() if tw is not None else -1
        if col == _COL_SHORT_NUM and self.column() == _COL_SHORT_NUM:
            try:
                return int(self.text() or "0") < int(other.text() or "0")
            except ValueError:
                return self.text() < other.text()
        return self.text() < other.text()


# Colors for coverage
_COLOR_ON_CANVAS = QColor(210, 240, 210)   # light green
_COLOR_NOT_ON_CANVAS = QColor(240, 210, 210)  # light red/pink
_COLOR_NEUTRAL = None  # no background (new row or no template loaded)

# Wrapped multi-column layout for DebugReportOrderDialog (column-major: top-to-bottom, then next column).
_DEBUG_REPORT_ORDER_GRID_W = 260
_DEBUG_REPORT_ORDER_GRID_H = 26


class DebugReportOrderDialog(QDialog):
    """Dialog to edit field column order and ``include_in_debug_report`` for v2 Excel."""

    def __init__(self, entries: list[FieldCatalogEntry], parent=None):
        super().__init__(parent)
        self._reset_all_orders = False
        self.setWindowTitle("Порядок в отчёте Excel")
        self.resize(900, 520)
        lay = QVBoxLayout(self)
        lay.addWidget(
            QLabel(
                "Перетащите строки (DnD) или «Вверх»/«Вниз» для текущей строки.\n"
                "Список заполняется сверху вниз; если по высоте не помещается — следующая колонка "
                "справа. Порядок в отчёте такой же: сверху вниз, затем соседняя колонка слева направо.\n"
                "Галочка — колонка в отладочном xlsx. «Сбросить порядок» выставит "
                "всем полям debug_report_order = 999 (порядок в файле — по id)."
            ),
        )
        self._list = QListWidget()
        self._list.setFlow(QListView.Flow.TopToBottom)
        self._list.setWrapping(True)
        self._list.setResizeMode(QListView.ResizeMode.Adjust)
        self._list.setUniformItemSizes(True)
        self._list.setGridSize(
            QSize(_DEBUG_REPORT_ORDER_GRID_W, _DEBUG_REPORT_ORDER_GRID_H),
        )
        self._list.setSpacing(3)
        self._list.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self._list.setDefaultDropAction(Qt.DropAction.MoveAction)
        self._list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        sorted_e = sorted(entries, key=lambda e: (e.debug_report_order, e.id))
        for e in sorted_e:
            it = QListWidgetItem(f"{e.id}  —  {e.label or '—'}")
            it.setData(Qt.ItemDataRole.UserRole, e.id)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            it.setCheckState(
                Qt.CheckState.Checked if e.include_in_debug_report else Qt.CheckState.Unchecked,
            )
            self._list.addItem(it)
        lay.addWidget(self._list)
        try:
            self._list.model().rowsMoved.connect(lambda *a: setattr(self, "_reset_all_orders", False))
        except Exception:
            pass

        btn_row = QHBoxLayout()
        btn_up = QPushButton("Вверх")
        btn_up.setToolTip(
            "Сдвинуть на одну позицию раньше в порядке отчёта "
            "(визуально — обычно строка выше в той же колонке или низ левой колонки).",
        )
        btn_down = QPushButton("Вниз")
        btn_down.setToolTip(
            "Сдвинуть на одну позицию позже в порядке отчёта "
            "(визуально — обычно строка ниже в той же колонке или верх правой колонки).",
        )
        btn_reset = QPushButton("Сбросить порядок")
        btn_up.clicked.connect(self._move_up)
        btn_down.clicked.connect(self._move_down)
        btn_reset.clicked.connect(self._reset_order)
        btn_row.addWidget(btn_up)
        btn_row.addWidget(btn_down)
        btn_row.addStretch()
        btn_row.addWidget(btn_reset)
        lay.addLayout(btn_row)

        bbox = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
        )
        bbox.accepted.connect(self.accept)
        bbox.rejected.connect(self.reject)
        lay.addWidget(bbox)

    def _move_up(self) -> None:
        self._reset_all_orders = False
        r = self._list.currentRow()
        if r <= 0:
            return
        it = self._list.takeItem(r)
        self._list.insertItem(r - 1, it)
        self._list.setCurrentRow(r - 1)

    def _move_down(self) -> None:
        self._reset_all_orders = False
        r = self._list.currentRow()
        if r < 0 or r >= self._list.count() - 1:
            return
        it = self._list.takeItem(r)
        self._list.insertItem(r + 1, it)
        self._list.setCurrentRow(r + 1)

    def _reset_order(self) -> None:
        self._reset_all_orders = True
        items: list[QListWidgetItem] = []
        while self._list.count():
            items.append(self._list.takeItem(0))
        items.sort(key=lambda it: str(it.data(Qt.ItemDataRole.UserRole) or ""))
        for it in items:
            self._list.addItem(it)

    def get_result(self) -> tuple[list[tuple[str, bool]], bool]:
        """Return ordered rows and whether to reset all ``debug_report_order`` to 999.

        Returns:
            A pair ``(pairs, reset_all_orders)`` where *pairs* is
            ``(field_id, include_in_debug_report)`` in list order, and
            *reset_all_orders* is True if the user applied *reset order* (all
            ``debug_report_order`` values become 999).
        """
        pairs: list[tuple[str, bool]] = []
        for i in range(self._list.count()):
            it = self._list.item(i)
            fid = str(it.data(Qt.ItemDataRole.UserRole) or "")
            inc = it.checkState() == Qt.CheckState.Checked
            pairs.append((fid, inc))
        return pairs, self._reset_all_orders


# ---------------------------------------------------------------------------
# Свойства каталога (name / description / projects) — JSON FieldCatalog
# ---------------------------------------------------------------------------


class CatalogProjectPropertiesDialog(QDialog):
    """Редактирование метаданных каталога: имя, описание, список идентификаторов проектов."""

    def __init__(
        self,
        name: str,
        description: str,
        projects: list[str],
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Свойства каталога")
        self.resize(500, 440)

        layout = QVBoxLayout(self)

        form = QFormLayout()
        self._name = QLineEdit(name)
        form.addRow("Имя каталога:", self._name)

        self._desc = QPlainTextEdit()
        self._desc.setPlainText(description)
        self._desc.setPlaceholderText("Краткое описание каталога…")
        self._desc.setMinimumHeight(72)
        form.addRow("Описание:", self._desc)
        layout.addLayout(form)

        hint = QLabel(
            "Идентификаторы проектов используются при выборе шаблонов "
            "(сопоставление с catalog.projects и именем папки проекта). "
            "Один идентификатор — одна строка в списке."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #555; font-size: 11px;")
        layout.addWidget(hint)

        proj_row = QHBoxLayout()
        self._lst = QListWidget()
        for p in projects:
            if str(p).strip():
                self._lst.addItem(str(p).strip())
        self._lst.setMinimumHeight(120)
        proj_row.addWidget(self._lst, stretch=1)

        vbtn = QVBoxLayout()
        btn_add = QPushButton("Добавить")
        btn_add.setToolTip("Добавить идентификатор проекта")
        btn_add.clicked.connect(self._add_project)
        btn_del = QPushButton("Удалить")
        btn_del.setToolTip("Удалить выделенную строку")
        btn_del.clicked.connect(self._del_project)
        vbtn.addWidget(btn_add)
        vbtn.addWidget(btn_del)
        vbtn.addStretch()
        proj_row.addLayout(vbtn)
        layout.addLayout(proj_row)

        bbox = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
        )
        bbox.accepted.connect(self.accept)
        bbox.rejected.connect(self.reject)
        layout.addWidget(bbox)

    def _add_project(self) -> None:
        text, ok = QInputDialog.getText(self, "Проект", "Идентификатор проекта:")
        if ok and text.strip():
            self._lst.addItem(text.strip())

    def _del_project(self) -> None:
        row = self._lst.currentRow()
        if row >= 0:
            self._lst.takeItem(row)

    def catalog_name(self) -> str:
        t = self._name.text().strip()
        return t or "catalog"

    def catalog_description(self) -> str:
        return self._desc.toPlainText().strip()

    def catalog_projects(self) -> list[str]:
        out: list[str] = []
        for i in range(self._lst.count()):
            s = self._lst.item(i).text().strip()
            if s:
                out.append(s)
        return out


# ---------------------------------------------------------------------------
# CatalogEditorPanel — встроенная нижняя панель (T7.2+)
# ---------------------------------------------------------------------------

class CatalogEditorPanel(QWidget):
    """Embedded bottom panel for editing the active project field catalog.

    Signals:
        catalog_saved(FieldCatalog)    -- emitted after each successful save
        navigate_to_field(str)         -- request canvas navigation to field_id
        sync_requested()               -- ask main window to push table → canvas (fresh snapshot)
        template_save_requested()      -- after catalog save in sync mode: persist template JSON
        catalog_dirty_changed(bool)    -- unsaved edits in table (like template modified)
        catalog_field_link_apply_requested(str, object)
            -- link row→template field: scene_field_id before merge, merged FieldDef (see CATALOG_TEMPLATE_ROLES.md)
        catalog_add_field_to_workspace_requested(object)
            -- новое поле на canvas из строки каталога (FieldDef с типовым bbox_mm)
    """

    catalog_saved = Signal(object)          # FieldCatalog
    navigate_to_field = Signal(str)         # field_id
    sync_requested = Signal()
    template_save_requested = Signal()
    catalog_dirty_changed = Signal(bool)
    catalog_field_link_apply_requested = Signal(str, object)  # scene_field_id, merged FieldDef
    catalog_add_field_to_workspace_requested = Signal(object)  # FieldDef

    def __init__(
        self,
        catalog: FieldCatalog | None = None,
        templates_dir: str = "",
        current_path: str = "",
        parent=None,
    ):
        super().__init__(parent)
        self._catalog: FieldCatalog | None = catalog
        self._templates_dir: str = templates_dir
        self._current_path: str = current_path
        self._scene_field_ids: set[str] = set()
        self._scene_fields: list[FieldDef] = []  # updated from main_window
        # Параллельно _scene_fields: номер поля как в панели «Список полей» (FieldRectItem.cell_index)
        self._scene_field_cell_indices: list[int] = []
        self._sync_enabled: bool = False
        self._sync_push_timer = QTimer(self)
        self._sync_push_timer.setSingleShot(True)
        self._sync_push_timer.setInterval(50)
        self._sync_push_timer.timeout.connect(self._push_sync_to_scene)
        self._catalog_dirty: bool = False
        self._catalog_dirty_suppress: int = 0

        self._build_ui()
        self._connect_shortcuts()

        if catalog:
            self._fill_from_catalog(catalog)

    # ------------------------------------------------------------------ public API

    def load_catalog(self, catalog: FieldCatalog, path: str = "") -> None:
        """Load catalog from outside (e.g. browser activate). Does not open dialog."""
        self._catalog = catalog
        self._current_path = path
        self._fill_from_catalog(catalog)
        if self._scene_field_ids:
            self._apply_coverage()

    def update_scene_fields(
        self,
        scene_fields: list[FieldDef],
        cell_indices: list[int] | None = None,
    ) -> None:
        """Called by main_window whenever fields on canvas change.
        Refreshes coverage highlight and nav button states.

        cell_indices: тот же порядок, что scene_fields; значения = # из панели «Список полей»
        (FieldRectItem.cell_index). Если None — подставляется 1..n (только fallback).
        """
        prev_exp = {
            fd.id: (fd.expected, fd.expected_text or "") for fd in self._scene_fields
        }
        self._scene_fields = list(scene_fields)
        n = len(scene_fields)
        if cell_indices is not None and len(cell_indices) == n:
            self._scene_field_cell_indices = list(cell_indices)
        else:
            self._scene_field_cell_indices = list(range(1, n + 1))
        self._scene_field_ids = {fd.id for fd in scene_fields}
        new_exp = {
            fd.id: (fd.expected, fd.expected_text or "") for fd in scene_fields
        }
        if prev_exp != new_exp:
            self._refresh_expected_columns_from_scene()
        self._sync_refresh_scene_links()
        self._apply_coverage()
        self._update_sync_button_state()
        self._update_status_bar()

    def _scene_field_for_id(self, fid: str) -> FieldDef | None:
        return _scene_field_lookup(self._scene_fields, fid)

    def _refresh_expected_columns_from_scene(self) -> None:
        """Fill expected / expected_text columns from the loaded template (not persisted in catalog)."""
        self._table.blockSignals(True)
        try:
            by_id = {fd.id: fd for fd in self._scene_fields}
            for row in range(self._table.rowCount()):
                fid = self._cell_text(row, _COL_ID)
                fd = by_id.get(fid)
                w_exp = self._table.cellWidget(row, _COL_EXPECTED)
                if isinstance(w_exp, QComboBox):
                    w_exp.setCurrentText(fd.expected if fd else "optional")
                item_et = self._table.item(row, _COL_EXPECTED_TEXT)
                if item_et:
                    item_et.setText((fd.expected_text or "") if fd else "")
        finally:
            self._table.blockSignals(False)

    def _sync_link_scene_id_for_row(self, row: int) -> str | None:
        """Canvas field id stored on id cell (pending rename: cell text may differ)."""
        it = self._table.item(row, _COL_ID)
        if it is None:
            return None
        v = it.data(USER_ROLE_SYNC_LINK)
        if v is None or v == "":
            return None
        return str(v).strip()

    def _set_sync_link_scene_id(self, row: int, scene_id: str | None) -> None:
        it = self._table.item(row, _COL_ID)
        if it is None:
            return
        it.setData(USER_ROLE_SYNC_LINK, (scene_id or "").strip() or None)

    def _sync_refresh_scene_links(self) -> None:
        """Align USER_ROLE_SYNC_LINK with cell id and scene; stable across sort (not row-index based)."""
        if not self._scene_field_ids:
            for r in range(self._table.rowCount()):
                it = self._table.item(r, _COL_ID)
                if it is not None:
                    it.setData(USER_ROLE_SYNC_LINK, None)
            return
        for r in range(self._table.rowCount()):
            tid = self._cell_text(r, _COL_ID)
            it = self._table.item(r, _COL_ID)
            if it is None:
                continue
            prev_link = it.data(USER_ROLE_SYNC_LINK)
            prev_s = str(prev_link).strip() if prev_link else ""
            if tid and tid in self._scene_field_ids:
                it.setData(USER_ROLE_SYNC_LINK, tid)
            elif tid and tid not in self._scene_field_ids:
                if prev_s and prev_s in self._scene_field_ids:
                    it.setData(USER_ROLE_SYNC_LINK, prev_s)
                else:
                    it.setData(USER_ROLE_SYNC_LINK, None)
            else:
                it.setData(USER_ROLE_SYNC_LINK, None)

    def _sync_link_after_id_edit(self, row: int) -> None:
        """After user edits id cell: keep link to old canvas id while new id not yet on scene."""
        tid = self._cell_text(row, _COL_ID)
        it = self._table.item(row, _COL_ID)
        if it is None:
            return
        prev_link = it.data(USER_ROLE_SYNC_LINK)
        prev_s = str(prev_link).strip() if prev_link else ""
        if tid and tid in self._scene_field_ids:
            it.setData(USER_ROLE_SYNC_LINK, tid)
        elif tid and tid not in self._scene_field_ids:
            if prev_s and prev_s in self._scene_field_ids:
                it.setData(USER_ROLE_SYNC_LINK, prev_s)
            else:
                it.setData(USER_ROLE_SYNC_LINK, None)
        else:
            it.setData(USER_ROLE_SYNC_LINK, None)

    def get_catalog(self) -> FieldCatalog:
        """Build FieldCatalog from current table content."""
        _validate_catalog_table_unique_ids(self._table, _COL_ID)
        entries: list[FieldCatalogEntry] = []
        for row in range(self._table.rowCount()):
            if not self._cell_text(row, _COL_ID):
                continue
            entries.append(self._entry_from_table_row(row))
        projects = self._catalog.projects if self._catalog else []
        description = self._catalog.description if self._catalog else ""
        return FieldCatalog(
            name=self._catalog.name if self._catalog else "catalog",
            projects=projects,
            description=description,
            entries=entries,
        )

    @property
    def current_path(self) -> str:
        return self._current_path

    def _find_row_for_scene_fd(self, fd: FieldDef) -> int | None:
        """Row whose catalog row applies to this scene field (by id or USER_ROLE sync link)."""
        fid = fd.id
        for r in range(self._table.rowCount()):
            if self._cell_text(r, _COL_ID) == fid:
                return r
        for r in range(self._table.rowCount()):
            if self._sync_link_scene_id_for_row(r) == fid:
                return r
        return None

    def refresh_sync_links_after_scene_apply(
        self,
        old_fds: list[FieldDef],
        new_fds: list[FieldDef],
    ) -> None:
        """After catalog→canvas sync: set sync link UserRole to new canvas ids (handles rename)."""
        for old_fd, new_fd in zip(old_fds, new_fds):
            r = self._find_row_for_scene_fd(old_fd)
            if r is None:
                continue
            nid = (new_fd.id or "").strip()
            if nid:
                self._set_sync_link_scene_id(r, nid)

    def _entry_from_table_row(self, row: int) -> FieldCatalogEntry:
        fid = self._cell_text(row, _COL_ID)
        label = self._cell_text(row, _COL_LABEL)
        try:
            sn = int(self._cell_text(row, _COL_SHORT_NUM) or "0")
        except ValueError:
            sn = 0
        group = self._combo_text(row, _COL_GROUP)
        ft = self._combo_text(row, _COL_FIELD_TYPE) or "data"
        desc = self._cell_text(row, _COL_DESCRIPTION)
        id_it = self._table.item(row, _COL_ID)
        dro_raw = id_it.data(USER_ROLE_DEBUG_REPORT_ORDER) if id_it else None
        try:
            debug_report_order = int(dro_raw) if dro_raw is not None else 999
        except (TypeError, ValueError):
            debug_report_order = 999
        inc_raw = id_it.data(USER_ROLE_INCLUDE_DEBUG_REPORT) if id_it else None
        include_in_debug_report = True if inc_raw is None else bool(inc_raw)
        note_w = self._table.cellWidget(row, _COL_REPORT_HEADER_NOTE)
        hdr_note = note_w.text().strip() if isinstance(note_w, QLineEdit) else ""
        is_doc = bool(id_it and id_it.data(USER_ROLE_IS_DOCUMENT_PROPERTY))
        return FieldCatalogEntry(
            id=fid,
            label=label,
            short_num=sn,
            description=desc,
            group=group,
            default_field_type=ft,
            debug_report_order=debug_report_order,
            include_in_debug_report=include_in_debug_report,
            debug_report_header_note=hdr_note,
            is_document_property=is_doc,
        )

    def _fielddef_from_catalog_row_new_on_canvas(self, row: int) -> FieldDef | None:
        """Семантика строки каталога + типовой bbox для нового поля (без копирования с существующего)."""
        from pdf_parsing_v2_engine.models import FieldDef as _FD

        table_id = self._cell_text(row, _COL_ID).strip()
        if not table_id:
            return None
        entry = self._entry_from_table_row(row)
        exp, et = _read_expected_from_table_row(
            self._table, _COL_EXPECTED, _COL_EXPECTED_TEXT, row,
        )
        exp_n = normalize_expected(exp)
        et_s = (et or "").strip()
        ft_n = normalize_field_type(entry.default_field_type or "data")
        lbl = (entry.label or "").strip()
        if entry.is_document_property:
            if table_id not in MANDATORY_PROPERTY_KEYS:
                QMessageBox.warning(
                    self,
                    "Каталог",
                    f"id «{table_id}» не из обязательного набора свойств PDF:\n"
                    f"{', '.join(MANDATORY_PROPERTY_KEYS)}",
                )
                return None
            from dataclasses import replace

            from pdf_parsing_v2_engine.document_properties.template_defaults import (
                default_document_property_field_def,
            )

            base = default_document_property_field_def(table_id)
            return replace(
                base,
                label=lbl or base.label,
                expected=exp_n,
                expected_text=et_s or None,
                field_type=ft_n,
            )
        return _FD(
            id=table_id,
            label=lbl,
            bbox_mm=_DEFAULT_CATALOG_NEW_FIELD_BBOX_MM,
            clean=None,
            validate_regex=None,
            expected=exp_n,
            padding_mm=None,
            field_type=ft_n,
            expected_text=et_s or None,
            is_anchor=False,
            outside_stamp=False,
            origin="frame_bottom_right",
            stretch_to_page=(),
        )

    def _merge_catalog_row_into_fielddef(self, row: int, fd: FieldDef) -> FieldDef | None:
        """Catalog semantics (id, label, types, expected) + unchanged template field geometry/settings."""
        from pdf_parsing_v2_engine.models import FieldDef as _FD

        table_id = self._cell_text(row, _COL_ID).strip()
        if not table_id:
            return None
        entry = self._entry_from_table_row(row)
        exp, et = _read_expected_from_table_row(
            self._table, _COL_EXPECTED, _COL_EXPECTED_TEXT, row,
        )
        exp_n = normalize_expected(exp)
        et_s = (et or "").strip()
        ft_n = normalize_field_type(entry.default_field_type or "data")
        lbl = (entry.label or "").strip()
        if entry.is_document_property:
            doc_key = table_id if table_id in MANDATORY_PROPERTY_KEYS else (fd.document_property or "")
            if not doc_key or doc_key not in MANDATORY_PROPERTY_KEYS:
                return None
            return _FD(
                id=table_id,
                label=lbl,
                bbox_mm=fd.bbox_mm,
                clean=fd.clean,
                validate_regex=fd.validate_regex,
                expected=exp_n,
                padding_mm=fd.padding_mm,
                field_type=ft_n,
                expected_text=et_s or None,
                is_anchor=False,
                outside_stamp=False,
                origin=fd.origin,
                stretch_to_page=fd.stretch_to_page,
                bound_top=None,
                bound_bottom=None,
                bound_left=None,
                bound_right=None,
                document_property=doc_key,
            )
        return _FD(
            id=table_id,
            label=lbl,
            bbox_mm=fd.bbox_mm,
            clean=fd.clean,
            validate_regex=fd.validate_regex,
            expected=exp_n,
            padding_mm=fd.padding_mm,
            field_type=ft_n,
            expected_text=et_s or None,
            is_anchor=fd.is_anchor,
            outside_stamp=fd.outside_stamp,
            origin=fd.origin,
            stretch_to_page=fd.stretch_to_page,
            bound_top=fd.bound_top,
            bound_bottom=fd.bound_bottom,
            bound_left=fd.bound_left,
            bound_right=fd.bound_right,
            document_property=fd.document_property,
        )

    def compute_sync_to_template(
        self, scene_fields: list[FieldDef],
    ) -> list[FieldDef]:
        """Return updated copies of *scene_fields* with semantics from catalog."""
        try:
            self.get_catalog()
        except ValueError:
            return scene_fields

        updated: list[FieldDef] = []
        for fd in scene_fields:
            row = self._find_row_for_scene_fd(fd)
            if row is None:
                updated.append(fd)
                continue
            merged = self._merge_catalog_row_into_fielddef(row, fd)
            if merged is None:
                updated.append(fd)
                continue
            fd_ft = normalize_field_type(fd.field_type)
            new_ft = normalize_field_type(merged.field_type)
            fd_exp = normalize_expected(fd.expected)
            new_exp = normalize_expected(merged.expected)
            fd_et = (fd.expected_text or "").strip()
            new_et = (merged.expected_text or "").strip()
            lbl_fd = (fd.label or "").strip()
            lbl_new = (merged.label or "").strip()
            if (
                fd.id == merged.id
                and lbl_fd == lbl_new
                and fd_ft == new_ft
                and fd_exp == new_exp
                and fd_et == new_et
            ):
                updated.append(fd)
            else:
                updated.append(merged)
        return updated

    def sync_selection_from_scene(self, field_id: str | None) -> None:
        """In sync mode: select catalog row for the field selected on canvas; scroll into view."""
        if not self._sync_enabled:
            return
        if not field_id or not str(field_id).strip():
            self._table.clearSelection()
            return
        fid = str(field_id).strip()
        for row in range(self._table.rowCount()):
            cell_id = self._cell_text(row, _COL_ID)
            row_sid = self._sync_link_scene_id_for_row(row)
            if cell_id != fid and row_sid != fid:
                continue
            self._table.selectRow(row)
            item = self._table.item(row, _COL_ID)
            if item is not None:
                self._table.scrollToItem(
                    item, QAbstractItemView.ScrollHint.PositionAtCenter,
                )
            self._table.setFocus(Qt.FocusReason.OtherFocusReason)
            return
        self._table.clearSelection()

    def focus_row_by_field_id(self, field_id: str) -> bool:
        """Select and focus row by catalog id (or sync-link to scene id)."""
        fid = (field_id or "").strip()
        if not fid:
            return False
        for row in range(self._table.rowCount()):
            cell_id = self._cell_text(row, _COL_ID)
            row_sid = self._sync_link_scene_id_for_row(row)
            if cell_id != fid and row_sid != fid:
                continue
            self._table.selectRow(row)
            item = self._table.item(row, _COL_ID)
            if item is not None:
                self._table.scrollToItem(
                    item, QAbstractItemView.ScrollHint.PositionAtCenter,
                )
            self._table.setFocus(Qt.FocusReason.OtherFocusReason)
            return True
        return False

    def on_scene_field_changed(self, action: str, field_def: FieldDef | None) -> None:
        """Called by main_window in sync mode: action = 'add'|'rename'.

        Note: 'remove' is intentionally not handled here — catalog is a superset
        of template fields; deleting a field from canvas does not remove it from
        the catalog. Coverage highlight is updated via _update_catalog_panel_coverage
        (debounced timer T7.1) which fires on every scene mutation.
        """
        if not self._sync_enabled:
            return
        if action in ("add", "rename") and field_def and field_def.id:
            # If not in catalog — add a new row silently
            fid = field_def.id
            for row in range(self._table.rowCount()):
                if self._cell_text(row, _COL_ID) == fid:
                    return  # already exists
            entry = FieldCatalogEntry(
                id=field_def.id,
                label=field_def.label,
                short_num=0,
                default_field_type=field_def.field_type,
                is_document_property=bool(field_def.document_property),
            )
            self._insert_entry_row(entry)

    def _open_debug_report_order_dialog(self) -> None:
        try:
            cat = self.get_catalog()
        except ValueError as exc:
            QMessageBox.warning(self, "Каталог", str(exc))
            return
        dlg = DebugReportOrderDialog(cat.entries, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        pairs, reset_all = dlg.get_result()
        by_id_row: dict[str, int] = {}
        for r in range(self._table.rowCount()):
            fid = self._cell_text(r, _COL_ID)
            if fid:
                by_id_row[fid] = r
        for idx, (fid, inc) in enumerate(pairs):
            row = by_id_row.get(fid)
            if row is None:
                continue
            id_it = self._table.item(row, _COL_ID)
            if id_it is None:
                continue
            ord_val = 999 if reset_all else 10 * idx
            id_it.setData(USER_ROLE_DEBUG_REPORT_ORDER, ord_val)
            id_it.setData(USER_ROLE_INCLUDE_DEBUG_REPORT, inc)
        self._mark_catalog_dirty()

    # ------------------------------------------------------------------ UI build

    def _build_ui(self) -> None:
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(2, 2, 2, 2)
        main_layout.setSpacing(2)

        # Title bar with close/detach buttons
        title_bar = self._build_title_bar()
        main_layout.addLayout(title_bar)

        # Table
        self._table = QTableWidget(0, _NCOLS)
        self._table.setHorizontalHeaderLabels(_HEADERS)
        hh = self._table.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hh.setStretchLastSection(True)
        hh.resizeSection(_COL_NAV, 28)
        hh.resizeSection(_COL_ID, 170)
        hh.resizeSection(_COL_LABEL, 200)
        hh.resizeSection(_COL_GROUP, 120)
        hh.resizeSection(_COL_SHORT_NUM, 35)
        hh.resizeSection(_COL_FIELD_TYPE, 90)
        hh.resizeSection(_COL_EXPECTED, 120)
        hh.resizeSection(_COL_EXPECTED_TEXT, 200)
        hh.resizeSection(_COL_REPORT_HEADER_NOTE, 140)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.verticalHeader().setDefaultSectionSize(22)
        self._table.itemChanged.connect(self._on_table_item_changed)
        self._table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._table.customContextMenuRequested.connect(self._on_catalog_context_menu)
        self._table.setSortingEnabled(True)
        hh.sortIndicatorChanged.connect(
            lambda _logical, _order: self._on_catalog_sort_indicator_changed(),
        )
        main_layout.addWidget(self._table, stretch=1)

        # Button row
        btn_row = self._build_button_row()
        main_layout.addLayout(btn_row)

        # Status bar
        self._lbl_status = QLabel("")
        self._lbl_status.setStyleSheet("color: #555; font-size: 11px;")
        main_layout.addWidget(self._lbl_status)

    def _build_title_bar(self) -> QHBoxLayout:
        hb = QHBoxLayout()
        hb.setContentsMargins(0, 0, 0, 0)

        self._lbl_title = QLabel("Каталог полей")
        f = self._lbl_title.font()
        f.setBold(True)
        self._lbl_title.setFont(f)
        hb.addWidget(self._lbl_title)
        hb.addStretch()

        # Sync toggle
        self._btn_sync = QToolButton()
        self._btn_sync.setCheckable(True)
        self._btn_sync.setText("Синхр.")
        self._btn_sync.setToolTip(
            "Режим синхронизации: изменения в шаблоне обновляют каталог и наоборот.\n"
            "Доступно при загруженном каталоге и шаблоне.\n"
            "Сохранение каталога (Ctrl+S) также сохраняет шаблон, чтобы JSON не разъехались."
        )
        self._btn_sync.setEnabled(False)
        self._btn_sync.toggled.connect(self._on_sync_toggled)
        hb.addWidget(self._btn_sync)

        # Filter: only missing
        self._chk_only_missing = QCheckBox("Только отсутств.")
        self._chk_only_missing.setToolTip("Показать только поля каталога, отсутствующие на canvas")
        self._chk_only_missing.stateChanged.connect(self._apply_coverage)
        hb.addWidget(self._chk_only_missing)

        # Detach button
        self._btn_detach = QToolButton()
        self._btn_detach.setText("⬆")
        self._btn_detach.setToolTip("Открыть в отдельном окне / вернуть обратно")
        self._btn_detach.setFixedWidth(24)
        self._btn_detach.clicked.connect(self._on_detach_clicked)
        hb.addWidget(self._btn_detach)

        # Close button
        btn_close = QToolButton()
        btn_close.setText("✕")
        btn_close.setToolTip("Скрыть панель каталога")
        btn_close.setFixedWidth(24)
        btn_close.clicked.connect(self.hide)
        hb.addWidget(btn_close)

        return hb

    def _build_button_row(self) -> QHBoxLayout:
        btn_row = QHBoxLayout()
        btn_row.setSpacing(4)

        btn_add = QPushButton("Добавить")
        btn_add.setToolTip("Добавить новую строку в каталог")
        btn_add.clicked.connect(self._add_row)
        btn_row.addWidget(btn_add)

        btn_del = QPushButton("Удалить")
        btn_del.setToolTip("Удалить выделенные строки")
        btn_del.clicked.connect(self._delete_selected)
        btn_row.addWidget(btn_del)

        btn_import = QPushButton("Импорт из шаблона")
        btn_import.setToolTip("Добавить поля текущего шаблона, отсутствующие в каталоге")
        btn_import.clicked.connect(self._import_from_template)
        btn_row.addWidget(btn_import)

        btn_report_order = QPushButton("Порядок в отчёте…")
        btn_report_order.setToolTip(
            "Порядок колонок и включение полей в отладочный Excel (v2 pipeline, save_v2_debug_report)",
        )
        btn_report_order.clicked.connect(self._open_debug_report_order_dialog)
        btn_row.addWidget(btn_report_order)

        btn_row.addStretch()

        btn_props = QPushButton("Свойства каталога…")
        btn_props.setToolTip(
            "Имя каталога, описание и идентификаторы проектов (JSON: name, description, projects)"
        )
        btn_props.clicked.connect(self._open_catalog_properties)
        btn_row.addWidget(btn_props)

        btn_load = QPushButton("Загрузить…")
        btn_load.clicked.connect(self._load)
        btn_row.addWidget(btn_load)

        btn_save = QPushButton("Сохранить (Ctrl+S)")
        btn_save.setToolTip(
            "Сохранить каталог JSON.\n"
            "Если включена «Синхр.», после сохранения каталога сохраняется и текущий шаблон."
        )
        btn_save.clicked.connect(self._save_quick)
        btn_row.addWidget(btn_save)

        btn_save_as = QPushButton("Сохранить как…")
        btn_save_as.clicked.connect(self._save_as)
        btn_row.addWidget(btn_save_as)

        return btn_row

    def _connect_shortcuts(self) -> None:
        sc = QShortcut(QKeySequence.StandardKey.Save, self)
        sc.activated.connect(self._save_quick)

    def _ensure_catalog_meta(self) -> None:
        if self._catalog is None:
            self._catalog = FieldCatalog(name="catalog", entries=[])

    def _open_catalog_properties(self) -> None:
        self._ensure_catalog_meta()
        assert self._catalog is not None
        dlg = CatalogProjectPropertiesDialog(
            self._catalog.name,
            self._catalog.description,
            list(self._catalog.projects),
            self,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        self._catalog.name = dlg.catalog_name()
        self._catalog.description = dlg.catalog_description()
        self._catalog.projects = dlg.catalog_projects()
        self._mark_catalog_dirty()

    # ------------------------------------------------------------------ coverage

    def _apply_coverage(self, _=None) -> None:
        """Colorize rows and update nav buttons based on _scene_field_ids."""
        only_missing = self._chk_only_missing.isChecked()
        has_template = bool(self._scene_field_ids)

        for row in range(self._table.rowCount()):
            fid = self._cell_text(row, _COL_ID)
            on_canvas = fid in self._scene_field_ids if has_template and fid else None

            # Background color
            if on_canvas is True:
                bg = _COLOR_ON_CANVAS
            elif on_canvas is False:
                bg = _COLOR_NOT_ON_CANVAS
            else:
                bg = None

            for col in range(_NCOLS):
                if col in (_COL_NAV, _COL_REPORT_HEADER_NOTE):
                    continue
                item = self._table.item(row, col)
                if item:
                    if bg:
                        item.setBackground(bg)
                    else:
                        item.setData(Qt.ItemDataRole.BackgroundRole, None)

            # Nav button
            nav_btn = self._table.cellWidget(row, _COL_NAV)
            if isinstance(nav_btn, QPushButton):
                nav_btn.setEnabled(bool(on_canvas))
                nav_btn.setToolTip(
                    "Перейти к полю на canvas" if on_canvas
                    else ("Поле отсутствует на canvas" if on_canvas is False else "Шаблон не загружен")
                )

            # Visibility for "only missing" filter
            hide_row = only_missing and has_template and on_canvas is True
            self._table.setRowHidden(row, hide_row)

        self._update_status_bar()

    def _update_status_bar(self) -> None:
        total_cat = sum(
            1 for row in range(self._table.rowCount())
            if self._cell_text(row, _COL_ID)
        )
        if not self._scene_field_ids:
            self._lbl_status.setText(f"Полей в каталоге: {total_cat}  |  шаблон не загружен")
            return
        on_canvas = sum(
            1 for row in range(self._table.rowCount())
            if self._cell_text(row, _COL_ID) in self._scene_field_ids
        )
        extra = len(self._scene_field_ids - {
            self._cell_text(row, _COL_ID)
            for row in range(self._table.rowCount())
            if self._cell_text(row, _COL_ID)
        })
        self._lbl_status.setText(
            f"На canvas: {on_canvas}/{total_cat} полей каталога"
            + (f"  |  {extra} поля canvas не в каталоге" if extra else "")
        )

    # ------------------------------------------------------------------ fill / rows

    def _fill_from_catalog(self, cat: FieldCatalog) -> None:
        self._catalog_dirty_suppress += 1
        try:
            so = self._table.isSortingEnabled()
            self._table.setSortingEnabled(False)
            self._table.blockSignals(True)
            try:
                self._table.setRowCount(0)
                for e in cat.entries:
                    self._insert_entry_row(e)
                self._refresh_expected_columns_from_scene()
            finally:
                self._table.blockSignals(False)
            self._sync_refresh_scene_links()
            self._apply_coverage()
            if so:
                self._table.setSortingEnabled(True)
        finally:
            self._catalog_dirty_suppress -= 1
        self._clear_catalog_dirty()

    def _on_catalog_sort_indicator_changed(self) -> None:
        """После сортировки: связь id↔canvas хранится в UserRole ячейки (не в индексе строки)."""
        self._sync_refresh_scene_links()
        self._apply_coverage()

    def _insert_entry_row(
        self, e: FieldCatalogEntry | None = None, at: int | None = None,
    ) -> int:
        n_before = self._table.rowCount()
        r = at if at is not None else n_before
        so = self._table.isSortingEnabled()
        if so:
            self._table.setSortingEnabled(False)
        self._table.insertRow(r)

        # Nav button (T7.4)
        nav_btn = QPushButton("→")
        nav_btn.setFixedWidth(26)
        nav_btn.setEnabled(False)
        nav_btn.setToolTip("Шаблон не загружен")
        fid_for_btn = e.id if e else ""
        # capture by value
        nav_btn.clicked.connect(lambda _checked, fid=fid_for_btn: self._on_nav_clicked(fid))
        self._table.setCellWidget(r, _COL_NAV, nav_btn)

        id_it = QTableWidgetItem(e.id if e else "")
        _init_catalog_id_last_valid(id_it)
        if e:
            id_it.setData(USER_ROLE_DEBUG_REPORT_ORDER, e.debug_report_order)
            id_it.setData(USER_ROLE_INCLUDE_DEBUG_REPORT, e.include_in_debug_report)
            id_it.setData(USER_ROLE_IS_DOCUMENT_PROPERTY, bool(e.is_document_property))
        else:
            id_it.setData(USER_ROLE_DEBUG_REPORT_ORDER, 999)
            id_it.setData(USER_ROLE_INCLUDE_DEBUG_REPORT, True)
            id_it.setData(USER_ROLE_IS_DOCUMENT_PROPERTY, False)
        self._table.setItem(r, _COL_ID, id_it)
        self._table.setItem(r, _COL_LABEL, QTableWidgetItem(e.label if e else ""))

        cmb_grp = _make_combo(_GROUP_OPTIONS, e.group if e else "")
        self._table.setCellWidget(r, _COL_GROUP, cmb_grp)

        self._table.setItem(
            r, _COL_SHORT_NUM, NumericSortTableWidgetItem(str(e.short_num) if e else "0"),
        )

        cmb_ft = _make_combo(_FIELD_TYPE_OPTIONS, e.default_field_type if e else "data")
        self._table.setCellWidget(r, _COL_FIELD_TYPE, cmb_ft)

        fid_init = e.id if e else ""
        fd_scene = self._scene_field_for_id(fid_init)
        exp_init = fd_scene.expected if fd_scene else "optional"
        et_init = (fd_scene.expected_text or "") if fd_scene else ""
        cmb_exp = _make_combo(_EXPECTED_OPTIONS, exp_init)
        self._table.setCellWidget(r, _COL_EXPECTED, cmb_exp)

        self._table.setItem(
            r, _COL_EXPECTED_TEXT, QTableWidgetItem(et_init),
        )
        self._table.setItem(
            r, _COL_DESCRIPTION, QTableWidgetItem(e.description if e else ""),
        )
        note_w = QLineEdit(e.debug_report_header_note if e else "")
        note_w.setPlaceholderText("None")
        note_w.textChanged.connect(lambda _t: self._mark_catalog_dirty())
        self._table.setCellWidget(r, _COL_REPORT_HEADER_NOTE, note_w)
        self._wire_row_sync_combos(r)
        self._sync_refresh_scene_links()
        if so:
            self._table.setSortingEnabled(True)
        if self._catalog_dirty_suppress == 0:
            self._mark_catalog_dirty()
        return r

    # ------------------------------------------------------------------ navigation (T7.4)

    def _bind_nav_button(self, row: int) -> None:
        """Reconnect → button so it navigates using the current id cell."""
        nav_btn = self._table.cellWidget(row, _COL_NAV)
        if not isinstance(nav_btn, QPushButton):
            return
        fid = self._cell_text(row, _COL_ID)
        try:
            nav_btn.clicked.disconnect()
        except (TypeError, RuntimeError):
            pass
        nav_btn.clicked.connect(lambda _checked, f=fid: self._on_nav_clicked(f))

    def _on_nav_clicked(self, field_id: str) -> None:
        """Nav button clicked — emit signal to main_window."""
        if field_id and field_id in self._scene_field_ids:
            self.navigate_to_field.emit(field_id)

    # ------------------------------------------------------------------ sync mode (T7.5)

    def _on_sync_toggled(self, checked: bool) -> None:
        self._sync_enabled = checked
        style = "background: #cfe8cf; font-weight: bold;" if checked else ""
        self._btn_sync.setStyleSheet(style)
        if checked:
            self._schedule_push_sync()

    def _schedule_push_sync(self) -> None:
        if not self._sync_enabled:
            return
        self._sync_push_timer.start(50)

    def _push_sync_to_scene(self) -> None:
        if not self._sync_enabled:
            return
        self.sync_requested.emit()

    def _on_table_item_changed(self, item) -> None:
        reverted_dup = False
        if item.column() == _COL_ID:
            reverted_dup = _enforce_unique_catalog_id_row(
                self._table, _COL_ID, item.row(), self,
            )
            self._sync_link_after_id_edit(item.row())
            self._bind_nav_button(item.row())
        if not reverted_dup:
            self._mark_catalog_dirty()
        if self._sync_enabled and not reverted_dup:
            self._schedule_push_sync()

    def _wire_row_sync_combos(self, row: int) -> None:
        def _emit() -> None:
            self._mark_catalog_dirty()
            if self._sync_enabled:
                self._schedule_push_sync()

        for col in (_COL_GROUP, _COL_FIELD_TYPE, _COL_EXPECTED):
            w = self._table.cellWidget(row, col)
            if isinstance(w, QComboBox):
                w.currentTextChanged.connect(lambda _t, _emit=_emit: _emit())

    def _update_sync_button_state(self) -> None:
        can_sync = bool(self._current_path and self._scene_field_ids)
        self._btn_sync.setEnabled(can_sync)
        if not can_sync and self._sync_enabled:
            self._btn_sync.setChecked(False)
            self._sync_enabled = False

    # ------------------------------------------------------------------ add / delete

    def _add_row(self) -> None:
        self._insert_entry_row()

    def _delete_selected(self) -> None:
        rows = sorted(
            {idx.row() for idx in self._table.selectedIndexes()}, reverse=True,
        )
        if not rows:
            return
        for r in rows:
            self._table.removeRow(r)
        self._sync_refresh_scene_links()
        self._apply_coverage()
        self._mark_catalog_dirty()
        if self._sync_enabled:
            self._schedule_push_sync()

    # ------------------------------------------------------------------ load / save

    def _catalogs_dir(self) -> str:
        if self._current_path:
            return os.path.dirname(self._current_path)
        if self._templates_dir:
            return self._templates_dir
        return ""

    def _load(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Загрузить каталог", self._catalogs_dir(), "JSON (*.json)",
        )
        if not path:
            return
        try:
            cat = FieldCatalog.from_json(path)
            self._catalog = cat
            self._current_path = path
            self._fill_from_catalog(cat)
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка", str(exc))

    def _save_quick(self) -> None:
        if self._current_path:
            self._write_catalog(self._current_path)
        else:
            self._save_as()

    def _save_as(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить каталог", self._catalogs_dir(), "JSON (*.json)",
        )
        if not path:
            return
        self._write_catalog(path)

    def _write_catalog(self, path: str) -> None:
        try:
            cat = self.get_catalog()
        except ValueError as exc:
            QMessageBox.warning(self, "Ошибка валидации", str(exc))
            return
        try:
            cat.to_json(path)
            self._current_path = path
            self._catalog = cat
            self._clear_catalog_dirty()
            self._update_sync_button_state()
            self.catalog_saved.emit(cat)
            if self._sync_enabled:
                self.sync_requested.emit()
                self.template_save_requested.emit()
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка", str(exc))

    def _mark_catalog_dirty(self) -> None:
        if self._catalog_dirty_suppress > 0:
            return
        prev = self._catalog_dirty
        self._catalog_dirty = True
        self._update_title()
        if not prev:
            self.catalog_dirty_changed.emit(True)

    def _clear_catalog_dirty(self) -> None:
        prev = self._catalog_dirty
        self._catalog_dirty = False
        self._update_title()
        if prev:
            self.catalog_dirty_changed.emit(False)

    def _update_title(self) -> None:
        star = " *" if self._catalog_dirty else ""
        if self._current_path:
            base = os.path.basename(self._current_path)
            self._lbl_title.setText(f"Каталог полей: {base}{star}")
        elif self._catalog and self._catalog.name and self._catalog.name != "catalog":
            self._lbl_title.setText(
                f"Каталог полей: {self._catalog.name} (не сохранён){star}",
            )
        else:
            self._lbl_title.setText(f"Каталог полей{star}")

    def _on_catalog_context_menu(self, pos: QPoint) -> None:
        idx = self._table.indexAt(pos)
        row = idx.row()
        if row < 0:
            return
        menu = QMenu(self)
        row_id = self._cell_text(row, _COL_ID).strip()
        id_it = self._table.item(row, _COL_ID)
        is_doc_row = bool(id_it and id_it.data(USER_ROLE_IS_DOCUMENT_PROPERTY))
        on_canvas = row_id in self._scene_field_ids if row_id else False
        act_add_canvas = menu.addAction(
            "Добавить свойство PDF в шаблон…" if is_doc_row else "Добавить поле на canvas…",
        )
        act_add_canvas.setEnabled(bool(row_id) and not on_canvas)
        act_add_canvas.setToolTip(
            "Добавить скрытое поле свойств PDF в JSON шаблона (без размещения на холсте)."
            if is_doc_row
            else (
                "Создать на рабочей области новый прямоугольник с id и семантикой из этой строки "
                "(типовой размер у угла рамки). Недоступно, если такой id уже есть в шаблоне."
            ),
        )
        act_link = menu.addAction("Привязать к полю на canvas…")
        act_link.setEnabled(bool(self._scene_fields) and not is_doc_row)
        act_link.setToolTip(
            "Указать прямоугольник поля на сцене: в него запишутся данные из строки каталога; "
            "геометрия и clean/regex останутся у поля (см. CATALOG_TEMPLATE_ROLES.md).",
        )
        chosen = menu.exec(self._table.viewport().mapToGlobal(pos))
        if chosen is act_add_canvas:
            self._add_catalog_row_to_workspace(row)
        elif chosen is act_link:
            self._link_row_to_canvas(row)

    def _add_catalog_row_to_workspace(self, row: int) -> None:
        row_id = self._cell_text(row, _COL_ID).strip()
        if not row_id:
            QMessageBox.warning(
                self, "Каталог",
                "В строке должен быть указан id — без него поле не создать.",
            )
            return
        if row_id in self._scene_field_ids:
            QMessageBox.information(
                self, "Каталог",
                f"Поле «{row_id}» уже есть на рабочей области (шаблоне).",
            )
            return
        try:
            self.get_catalog()
        except ValueError as exc:
            QMessageBox.warning(self, "Каталог", str(exc))
            return
        fd = self._fielddef_from_catalog_row_new_on_canvas(row)
        if fd is None:
            return
        self.catalog_add_field_to_workspace_requested.emit(fd)

    def _link_row_to_canvas(self, row: int) -> None:
        if not self._scene_fields:
            QMessageBox.information(self, "Каталог", "Нет полей на canvas (загрузите шаблон).")
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("Привязать строку к полю на canvas")
        dlg.setMinimumWidth(480)
        lay = QVBoxLayout(dlg)
        lay.addWidget(
            QLabel(
                "Выберите прямоугольник поля на сцене, к которому относится эта строка каталога.\n"
                "В шаблон запишутся данные из строки каталога (id, label, тип, expected…); "
                "координаты, clean, regex, padding, anchor, origin и т.п. останутся с поля на сцене.\n"
                "Номер в списке совпадает с колонкой «#» в панели «Список полей» справа; "
                "показан id поля (как в каталоге), без label.",
            ),
        )
        combo = QComboBox()
        combo.setMinimumWidth(440)
        for i, fd in enumerate(self._scene_fields):
            if fd.document_property:
                continue
            num = (
                self._scene_field_cell_indices[i]
                if i < len(self._scene_field_cell_indices)
                else i + 1
            )
            fid = (fd.id or "").strip() or "—"
            text = f"{num}. {fid}"
            combo.addItem(text, fd.id)
        lay.addWidget(combo)
        if combo.count() == 0:
            QMessageBox.information(
                self,
                "Каталог",
                "Нет полей на canvas для привязки (в шаблоне только скрытые свойства PDF).",
            )
            return
        bbox = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
        )
        bbox.accepted.connect(dlg.accept)
        bbox.rejected.connect(dlg.reject)
        lay.addWidget(bbox)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        scene_field_id = combo.currentData()
        if not scene_field_id:
            return
        base_fd = _scene_field_lookup(self._scene_fields, scene_field_id)
        if base_fd is None:
            return
        try:
            self.get_catalog()
        except ValueError as exc:
            QMessageBox.warning(self, "Каталог", str(exc))
            return
        merged = self._merge_catalog_row_into_fielddef(row, base_fd)
        if merged is None:
            QMessageBox.warning(
                self, "Каталог",
                "В строке должен быть указан id — из него формируется поле шаблона.",
            )
            return
        for r in range(self._table.rowCount()):
            if r != row and self._cell_text(r, _COL_ID) == merged.id:
                rep = QMessageBox.question(
                    self,
                    "Дубликат id",
                    f"Идентификатор «{merged.id}» уже указан в строке {r + 1}.\n"
                    "Очистить id в той строке и продолжить?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if rep != QMessageBox.StandardButton.Yes:
                    return
                o = self._table.item(r, _COL_ID)
                if o:
                    o.setText("")
                self._bind_nav_button(r)
                break
        self.catalog_field_link_apply_requested.emit(scene_field_id, merged)

    # ------------------------------------------------------------------ import

    def _import_from_template(self) -> None:
        if not self._scene_fields:
            QMessageBox.information(self, "Импорт", "Нет полей на сцене.")
            return

        try:
            current_cat = self.get_catalog()
        except ValueError:
            current_cat = FieldCatalog(name="catalog", entries=[])

        # Snapshot current expected/expected_text from catalog table for diff display
        catalog_expected: dict[str, tuple] = {}
        for row in range(self._table.rowCount()):
            fid = self._cell_text(row, _COL_ID)
            if not fid:
                continue
            exp, et = _read_expected_from_table_row(
                self._table, _COL_EXPECTED, _COL_EXPECTED_TEXT, row,
            )
            catalog_expected[fid] = (exp, et)

        dlg = _ImportFromTemplateDialog(
            scene_fields=self._scene_fields,
            catalog=current_cat,
            catalog_expected=catalog_expected,
            parent=self,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        existing_updates, bind_updates, new_entries = dlg.get_result()
        if not existing_updates and not bind_updates and not new_entries:
            QMessageBox.information(self, "Импорт", "Нет изменений для применения.")
            return

        self._catalog_dirty_suppress += 1
        so = self._table.isSortingEnabled()
        if so:
            self._table.setSortingEnabled(False)
        added = updated = bound = 0
        try:
            # 1. Update expected/expected_text for existing catalog entries
            for fid, new_exp, new_et in existing_updates:
                for row in range(self._table.rowCount()):
                    if self._cell_text(row, _COL_ID) == fid:
                        w_exp = self._table.cellWidget(row, _COL_EXPECTED)
                        if isinstance(w_exp, QComboBox):
                            w_exp.setCurrentText(new_exp)
                        it_et = self._table.item(row, _COL_EXPECTED_TEXT)
                        if it_et:
                            it_et.setText(new_et)
                        updated += 1
                        break

            # 2. Insert new catalog entries
            # (_insert_entry_row auto-fills expected/expected_text from _scene_fields)
            for entry in new_entries:
                self._insert_entry_row(entry)
                added += 1

            # 3. Bind: update catalog row expected/expected_text from template
            #    field, then re-assign the canvas field to the catalog entry's id
            for scene_fid, cat_entry_id in bind_updates:
                fd = next(
                    (f for f in self._scene_fields if f.id == scene_fid), None,
                )
                if fd is None:
                    continue
                cat_row = None
                for row in range(self._table.rowCount()):
                    if self._cell_text(row, _COL_ID) == cat_entry_id:
                        cat_row = row
                        break
                if cat_row is None:
                    continue
                w_exp = self._table.cellWidget(cat_row, _COL_EXPECTED)
                if isinstance(w_exp, QComboBox):
                    w_exp.setCurrentText(normalize_expected(fd.expected))
                it_et = self._table.item(cat_row, _COL_EXPECTED_TEXT)
                if it_et:
                    it_et.setText(fd.expected_text or "")
                # Build merged FieldDef (catalog semantics + canvas geometry)
                merged = self._merge_catalog_row_into_fielddef(cat_row, fd)
                if merged is not None:
                    self.catalog_field_link_apply_requested.emit(scene_fid, merged)
                    bound += 1

            self._sync_refresh_scene_links()
            self._apply_coverage()
            if self._sync_enabled:
                self._schedule_push_sync()
        finally:
            if so:
                self._table.setSortingEnabled(True)
            self._catalog_dirty_suppress -= 1

        if added or updated or bound:
            self._mark_catalog_dirty()

        parts: list[str] = []
        if updated:
            parts.append(f"обновлено expected/expected_text: {updated}")
        if added:
            parts.append(f"добавлено новых записей: {added}")
        if bound:
            parts.append(f"привязано к записям каталога: {bound}")
        QMessageBox.information(self, "Импорт", "\n".join(parts) if parts else "Нет изменений.")

    # ------------------------------------------------------------------ detach (T7.detach)

    def _on_detach_clicked(self) -> None:
        """Signal to parent to detach/dock this panel."""
        # The main_window connects to this via a custom approach:
        # If already detached (top-level window), dock back; else detach.
        parent = self.parent()
        if hasattr(parent, "_detach_catalog_panel"):
            parent._detach_catalog_panel()
        elif hasattr(parent, "_dock_catalog_panel"):
            parent._dock_catalog_panel()

    # ------------------------------------------------------------------ cell helpers

    def _cell_text(self, row: int, col: int) -> str:
        item = self._table.item(row, col)
        return item.text().strip() if item else ""

    def _combo_text(self, row: int, col: int) -> str:
        w = self._table.cellWidget(row, col)
        if isinstance(w, QComboBox):
            return w.currentText().strip()
        return self._cell_text(row, col)


# ---------------------------------------------------------------------------
# CatalogEditorDialog — legacy modal dialog (backward compat)
# ---------------------------------------------------------------------------

# Column offsets for the old dialog (no nav column)
_D_COL_ID = 0
_D_COL_LABEL = 1
_D_COL_GROUP = 2
_D_COL_SHORT_NUM = 3
_D_COL_FIELD_TYPE = 4
_D_COL_EXPECTED = 5
_D_COL_EXPECTED_TEXT = 6
_D_COL_DESCRIPTION = 7
_D_COL_REPORT_HEADER_NOTE = 8
_D_NCOLS = 9
_D_HEADERS = [
    "id", "label", "group", "#", "field_type",
    "expected", "expected_text", "description", "заголовок отчёта",
]


class CatalogEditorDialog(QDialog):
    """Legacy modal dialog. Use CatalogEditorPanel for embedded/modeless UX."""

    def __init__(
        self,
        catalog: FieldCatalog | None = None,
        templates_dir: str = "",
        current_path: str = "",
        scene_fields: list[FieldDef] | None = None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Каталог полей")
        self.resize(1000, 560)
        self._catalog = catalog
        self._templates_dir = templates_dir
        self._current_path: str = current_path
        self._scene_fields: list[FieldDef] | None = scene_fields
        self._sync_summary: str = ""

        layout = QVBoxLayout(self)

        self._table = QTableWidget(0, _D_NCOLS)
        self._table.setHorizontalHeaderLabels(_D_HEADERS)
        hh = self._table.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hh.setStretchLastSection(True)
        hh.resizeSection(_D_COL_ID, 160)
        hh.resizeSection(_D_COL_LABEL, 180)
        hh.resizeSection(_D_COL_GROUP, 100)
        hh.resizeSection(_D_COL_SHORT_NUM, 40)
        hh.resizeSection(_D_COL_FIELD_TYPE, 80)
        hh.resizeSection(_D_COL_EXPECTED, 80)
        hh.resizeSection(_D_COL_EXPECTED_TEXT, 120)
        hh.resizeSection(_D_COL_REPORT_HEADER_NOTE, 130)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.itemChanged.connect(self._on_dialog_table_item_changed)
        layout.addWidget(self._table)

        btn_row = QHBoxLayout()
        btn_add = QPushButton("Добавить поле")
        btn_add.clicked.connect(self._add_row)
        btn_row.addWidget(btn_add)
        btn_del = QPushButton("Удалить выделенные")
        btn_del.clicked.connect(self._delete_selected)
        btn_row.addWidget(btn_del)
        btn_row.addStretch()
        if self._scene_fields is not None:
            btn_import = QPushButton("Импорт из шаблона")
            btn_import.clicked.connect(self._import_from_template)
            btn_row.addWidget(btn_import)
        btn_props = QPushButton("Свойства каталога…")
        btn_props.setToolTip(
            "Имя каталога, описание и идентификаторы проектов (JSON: name, description, projects)"
        )
        btn_props.clicked.connect(self._open_catalog_properties)
        btn_row.addWidget(btn_props)
        btn_load = QPushButton("Загрузить…")
        btn_load.clicked.connect(self._load)
        btn_row.addWidget(btn_load)
        btn_save = QPushButton("Сохранить (Ctrl+S)")
        btn_save.clicked.connect(self._save_quick)
        btn_row.addWidget(btn_save)
        btn_save_as = QPushButton("Сохранить как…")
        btn_save_as.clicked.connect(self._save_as)
        btn_row.addWidget(btn_save_as)
        layout.addLayout(btn_row)

        bottom = QHBoxLayout()
        bottom.addStretch()
        btn_ok = QPushButton("OK")
        btn_ok.clicked.connect(self.accept)
        bottom.addWidget(btn_ok)
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(self.reject)
        bottom.addWidget(btn_cancel)
        layout.addLayout(bottom)

        sc = QShortcut(QKeySequence.StandardKey.Save, self)
        sc.activated.connect(self._save_quick)

        if catalog:
            self._fill_from_catalog(catalog)

    @property
    def current_path(self) -> str:
        return self._current_path

    @property
    def sync_summary(self) -> str:
        return self._sync_summary

    def get_catalog(self) -> FieldCatalog:
        _validate_catalog_table_unique_ids(self._table, _D_COL_ID)
        entries: list[FieldCatalogEntry] = []
        for row in range(self._table.rowCount()):
            fid = self._cell_text(row, _D_COL_ID)
            if not fid:
                continue
            label = self._cell_text(row, _D_COL_LABEL)
            try:
                sn = int(self._cell_text(row, _D_COL_SHORT_NUM) or "0")
            except ValueError:
                sn = 0
            group = self._combo_text(row, _D_COL_GROUP)
            ft = self._combo_text(row, _D_COL_FIELD_TYPE) or "data"
            desc = self._cell_text(row, _D_COL_DESCRIPTION)
            id_it = self._table.item(row, _D_COL_ID)
            dro_raw = id_it.data(USER_ROLE_DEBUG_REPORT_ORDER) if id_it else None
            try:
                debug_report_order = int(dro_raw) if dro_raw is not None else 999
            except (TypeError, ValueError):
                debug_report_order = 999
            inc_raw = id_it.data(USER_ROLE_INCLUDE_DEBUG_REPORT) if id_it else None
            include_in_debug_report = True if inc_raw is None else bool(inc_raw)
            note_w = self._table.cellWidget(row, _D_COL_REPORT_HEADER_NOTE)
            hdr_note = note_w.text().strip() if isinstance(note_w, QLineEdit) else ""
            is_doc = bool(id_it and id_it.data(USER_ROLE_IS_DOCUMENT_PROPERTY))
            entries.append(
                FieldCatalogEntry(
                    id=fid,
                    label=label,
                    short_num=sn,
                    description=desc,
                    group=group,
                    default_field_type=ft,
                    debug_report_order=debug_report_order,
                    include_in_debug_report=include_in_debug_report,
                    debug_report_header_note=hdr_note,
                    is_document_property=is_doc,
                )
            )
        projects = self._catalog.projects if self._catalog else []
        description = self._catalog.description if self._catalog else ""
        return FieldCatalog(
            name=self._catalog.name if self._catalog else "catalog",
            projects=projects,
            description=description,
            entries=entries,
        )

    def _fill_from_catalog(self, cat: FieldCatalog) -> None:
        self._table.blockSignals(True)
        try:
            self._table.setRowCount(0)
            for e in cat.entries:
                self._insert_entry_row(e)
            self._refresh_expected_columns_from_scene()
        finally:
            self._table.blockSignals(False)

    def _ensure_catalog_meta(self) -> None:
        if self._catalog is None:
            self._catalog = FieldCatalog(name="catalog", entries=[])

    def _open_catalog_properties(self) -> None:
        self._ensure_catalog_meta()
        assert self._catalog is not None
        dlg = CatalogProjectPropertiesDialog(
            self._catalog.name,
            self._catalog.description,
            list(self._catalog.projects),
            self,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        self._catalog.name = dlg.catalog_name()
        self._catalog.description = dlg.catalog_description()
        self._catalog.projects = dlg.catalog_projects()
        if self._current_path:
            self.setWindowTitle(f"Каталог полей — {os.path.basename(self._current_path)}")
        else:
            self.setWindowTitle(f"Каталог полей — {self._catalog.name}")

    def _on_dialog_table_item_changed(self, item) -> None:
        if item.column() == _D_COL_ID:
            _enforce_unique_catalog_id_row(self._table, _D_COL_ID, item.row(), self)

    def _refresh_expected_columns_from_scene(self) -> None:
        by_id = {fd.id: fd for fd in (self._scene_fields or [])}
        for row in range(self._table.rowCount()):
            fid = self._cell_text(row, _D_COL_ID)
            fd = by_id.get(fid)
            w_exp = self._table.cellWidget(row, _D_COL_EXPECTED)
            if isinstance(w_exp, QComboBox):
                w_exp.setCurrentText(fd.expected if fd else "optional")
            item_et = self._table.item(row, _D_COL_EXPECTED_TEXT)
            if item_et:
                item_et.setText((fd.expected_text or "") if fd else "")

    def _insert_entry_row(
        self, e: FieldCatalogEntry | None = None, at: int | None = None,
    ) -> int:
        r = at if at is not None else self._table.rowCount()
        self._table.insertRow(r)
        id_it = QTableWidgetItem(e.id if e else "")
        _init_catalog_id_last_valid(id_it)
        if e:
            id_it.setData(USER_ROLE_DEBUG_REPORT_ORDER, e.debug_report_order)
            id_it.setData(USER_ROLE_INCLUDE_DEBUG_REPORT, e.include_in_debug_report)
            id_it.setData(USER_ROLE_IS_DOCUMENT_PROPERTY, bool(e.is_document_property))
        else:
            id_it.setData(USER_ROLE_DEBUG_REPORT_ORDER, 999)
            id_it.setData(USER_ROLE_INCLUDE_DEBUG_REPORT, True)
            id_it.setData(USER_ROLE_IS_DOCUMENT_PROPERTY, False)
        self._table.setItem(r, _D_COL_ID, id_it)
        self._table.setItem(r, _D_COL_LABEL, QTableWidgetItem(e.label if e else ""))
        cmb_grp = _make_combo(_GROUP_OPTIONS, e.group if e else "")
        self._table.setCellWidget(r, _D_COL_GROUP, cmb_grp)
        self._table.setItem(r, _D_COL_SHORT_NUM, QTableWidgetItem(str(e.short_num) if e else "0"))
        cmb_ft = _make_combo(_FIELD_TYPE_OPTIONS, e.default_field_type if e else "data")
        self._table.setCellWidget(r, _D_COL_FIELD_TYPE, cmb_ft)
        fid_init = e.id if e else ""
        fd_scene = _scene_field_lookup(self._scene_fields, fid_init)
        exp_init = fd_scene.expected if fd_scene else "optional"
        et_init = (fd_scene.expected_text or "") if fd_scene else ""
        cmb_exp = _make_combo(_EXPECTED_OPTIONS, exp_init)
        self._table.setCellWidget(r, _D_COL_EXPECTED, cmb_exp)
        self._table.setItem(r, _D_COL_EXPECTED_TEXT, QTableWidgetItem(et_init))
        self._table.setItem(r, _D_COL_DESCRIPTION, QTableWidgetItem(e.description if e else ""))
        note_w = QLineEdit(e.debug_report_header_note if e else "")
        note_w.setPlaceholderText("None")
        self._table.setCellWidget(r, _D_COL_REPORT_HEADER_NOTE, note_w)
        return r

    def _cell_text(self, row: int, col: int) -> str:
        item = self._table.item(row, col)
        return item.text().strip() if item else ""

    def _combo_text(self, row: int, col: int) -> str:
        w = self._table.cellWidget(row, col)
        if isinstance(w, QComboBox):
            return w.currentText().strip()
        return self._cell_text(row, col)

    def _add_row(self) -> None:
        self._insert_entry_row()

    def _delete_selected(self) -> None:
        rows = sorted({idx.row() for idx in self._table.selectedIndexes()}, reverse=True)
        for r in rows:
            self._table.removeRow(r)

    def _catalogs_dir(self) -> str:
        if self._current_path:
            return os.path.dirname(self._current_path)
        if self._templates_dir:
            d = os.path.join(self._templates_dir, "catalogs")
            if os.path.isdir(d):
                return d
            return self._templates_dir
        return ""

    def _load(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Загрузить каталог", self._catalogs_dir(), "JSON (*.json)")
        if not path:
            return
        try:
            cat = FieldCatalog.from_json(path)
            self._catalog = cat
            self._current_path = path
            self._fill_from_catalog(cat)
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка", str(exc))

    def _save_quick(self) -> None:
        if self._current_path:
            self._write_catalog(self._current_path)
        else:
            self._save_as()

    def _save_as(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить каталог", self._catalogs_dir(), "JSON (*.json)")
        if not path:
            return
        self._write_catalog(path)

    def _write_catalog(self, path: str) -> None:
        try:
            cat = self.get_catalog()
        except ValueError as exc:
            QMessageBox.warning(self, "Ошибка валидации", str(exc))
            return
        try:
            cat.to_json(path)
            self._current_path = path
            self._catalog = cat
            self.setWindowTitle(f"Каталог полей — {os.path.basename(path)}")
        except Exception as exc:
            QMessageBox.critical(self, "Ошибка", str(exc))

    def _import_from_template(self) -> None:
        if not self._scene_fields:
            QMessageBox.information(self, "Импорт", "Нет полей на сцене.")
            return
        try:
            current_cat = self.get_catalog()
        except ValueError:
            current_cat = FieldCatalog(name="catalog", entries=[])
        existing_ids = {e.id for e in current_cat.entries}
        new_fields = [f for f in self._scene_fields if f.id not in existing_ids]
        update_fields = [f for f in self._scene_fields if f.id in existing_ids]
        if not new_fields and not update_fields:
            QMessageBox.information(self, "Импорт", "Нет новых или изменённых полей.")
            return
        parts: list[str] = []
        if new_fields:
            parts.append(f"Новые ({len(new_fields)}): " + ", ".join(f.id for f in new_fields[:15]))
        if update_fields:
            parts.append(f"Обновить ({len(update_fields)}): " + ", ".join(f.id for f in update_fields[:15]))
        r = QMessageBox.question(
            self, "Импорт из шаблона",
            "\n".join(parts) + "\n\nДобавить/обновить?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            return
        added = 0
        for fd in new_fields:
            entry = FieldCatalogEntry(
                id=fd.id,
                label=fd.label,
                short_num=0,
                default_field_type=fd.field_type,
                is_document_property=bool(fd.document_property),
            )
            self._insert_entry_row(entry)
            added += 1
        # Only update expected/expected_text (template-owned); label/field_type
        # belong to the catalog and are not overwritten from the template.
        updated = 0
        for fd in update_fields:
            for row in range(self._table.rowCount()):
                if self._cell_text(row, _D_COL_ID) == fd.id:
                    w_exp = self._table.cellWidget(row, _D_COL_EXPECTED)
                    if isinstance(w_exp, QComboBox):
                        w_exp.setCurrentText(normalize_expected(fd.expected))
                    it_et = self._table.item(row, _D_COL_EXPECTED_TEXT)
                    if it_et:
                        it_et.setText(fd.expected_text or "")
                    updated += 1
                    break
        QMessageBox.information(self, "Импорт", f"Добавлено: {added}, обновлено expected: {updated}")

    def _find_row_for_scene_fd_dialog(self, fd: FieldDef) -> int | None:
        for r in range(self._table.rowCount()):
            if self._cell_text(r, _D_COL_ID) == fd.id:
                return r
        return None

    def _entry_from_table_row_dialog(self, row: int) -> FieldCatalogEntry:
        fid = self._cell_text(row, _D_COL_ID)
        label = self._cell_text(row, _D_COL_LABEL)
        try:
            sn = int(self._cell_text(row, _D_COL_SHORT_NUM) or "0")
        except ValueError:
            sn = 0
        group = self._combo_text(row, _D_COL_GROUP)
        ft = self._combo_text(row, _D_COL_FIELD_TYPE) or "data"
        desc = self._cell_text(row, _D_COL_DESCRIPTION)
        id_it = self._table.item(row, _D_COL_ID)
        dro_raw = id_it.data(USER_ROLE_DEBUG_REPORT_ORDER) if id_it else None
        try:
            debug_report_order = int(dro_raw) if dro_raw is not None else 999
        except (TypeError, ValueError):
            debug_report_order = 999
        inc_raw = id_it.data(USER_ROLE_INCLUDE_DEBUG_REPORT) if id_it else None
        include_in_debug_report = True if inc_raw is None else bool(inc_raw)
        note_w = self._table.cellWidget(row, _D_COL_REPORT_HEADER_NOTE)
        hdr_note = note_w.text().strip() if isinstance(note_w, QLineEdit) else ""
        is_doc = bool(id_it and id_it.data(USER_ROLE_IS_DOCUMENT_PROPERTY))
        return FieldCatalogEntry(
            id=fid,
            label=label,
            short_num=sn,
            description=desc,
            group=group,
            default_field_type=ft,
            debug_report_order=debug_report_order,
            include_in_debug_report=include_in_debug_report,
            debug_report_header_note=hdr_note,
            is_document_property=is_doc,
        )

    def compute_sync_to_template(self, scene_fields: list[FieldDef]) -> list[FieldDef]:
        try:
            self.get_catalog()
        except ValueError:
            self._sync_summary = "Ошибка валидации каталога, синхронизация пропущена"
            return scene_fields
        from pdf_parsing_v2_engine.models import FieldDef as _FD

        updated: list[FieldDef] = []
        count = 0
        for fd in scene_fields:
            row = self._find_row_for_scene_fd_dialog(fd)
            if row is None:
                updated.append(fd)
                continue
            table_id = self._cell_text(row, _D_COL_ID)
            if not table_id:
                updated.append(fd)
                continue
            entry = self._entry_from_table_row_dialog(row)
            exp, et = _read_expected_from_table_row(
                self._table, _D_COL_EXPECTED, _D_COL_EXPECTED_TEXT, row,
            )
            exp_n = normalize_expected(exp)
            et_s = (et or "").strip()
            ft_n = normalize_field_type(entry.default_field_type or "data")
            fd_ft = normalize_field_type(fd.field_type)
            fd_exp = normalize_expected(fd.expected)
            fd_et = (fd.expected_text or "").strip()
            lbl = (entry.label or "").strip()
            lbl_fd = (fd.label or "").strip()
            if entry.is_document_property:
                doc_key = (
                    table_id if table_id in MANDATORY_PROPERTY_KEYS else (fd.document_property or "")
                )
                if not doc_key or doc_key not in MANDATORY_PROPERTY_KEYS:
                    updated.append(fd)
                    continue
                new_fd = _FD(
                    id=table_id,
                    label=lbl,
                    bbox_mm=fd.bbox_mm,
                    clean=fd.clean,
                    validate_regex=fd.validate_regex,
                    expected=exp_n,
                    padding_mm=fd.padding_mm,
                    field_type=ft_n,
                    expected_text=et_s or None,
                    is_anchor=False,
                    outside_stamp=False,
                    origin=fd.origin,
                    stretch_to_page=fd.stretch_to_page,
                    bound_top=None,
                    bound_bottom=None,
                    bound_left=None,
                    bound_right=None,
                    document_property=doc_key,
                )
                changed = (
                    fd.id != new_fd.id
                    or lbl_fd != lbl
                    or fd_ft != ft_n
                    or fd_exp != exp_n
                    or fd_et != et_s
                    or fd.document_property != new_fd.document_property
                )
            else:
                changed = (
                    fd.id != table_id
                    or lbl_fd != lbl
                    or fd_ft != ft_n
                    or fd_exp != exp_n
                    or fd_et != et_s
                )
                new_fd = _FD(
                    id=table_id,
                    label=lbl,
                    bbox_mm=fd.bbox_mm,
                    clean=fd.clean,
                    validate_regex=fd.validate_regex,
                    expected=exp_n,
                    padding_mm=fd.padding_mm,
                    field_type=ft_n,
                    expected_text=et_s or None,
                    is_anchor=fd.is_anchor,
                    outside_stamp=fd.outside_stamp,
                    origin=fd.origin,
                    stretch_to_page=fd.stretch_to_page,
                    bound_top=fd.bound_top,
                    bound_bottom=fd.bound_bottom,
                    bound_left=fd.bound_left,
                    bound_right=fd.bound_right,
                    document_property=fd.document_property,
                )
            if changed:
                updated.append(new_fd)
                count += 1
            else:
                updated.append(fd)
        self._sync_summary = f"Обновлено {count} полей шаблона из каталога" if count else ""
        return updated


# ---------------------------------------------------------------------------
# _ImportFromTemplateDialog  (T7.2: interactive import)
# ---------------------------------------------------------------------------

_BIND_NEW_SENTINEL = ""  # empty string = "create new catalog entry"


class _ImportFromTemplateDialog(QDialog):
    """Interactive import dialog: template fields -> catalog (T7.2).

    Section 1 (existing): template fields whose id already exists in catalog.
      Only expected / expected_text can be updated.
      label / field_type are catalog-owned and are NOT overwritten.

    Section 2 (new): template fields whose id is NOT in catalog.
      User can bind to an existing catalog entry (canvas field gets catalog
      semantics), or create a new catalog entry.
    """

    def __init__(
        self,
        scene_fields: list,
        catalog: object,
        catalog_expected: dict,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Импорт полей из шаблона в каталог")
        self.resize(960, 560)

        catalog_ids = {e.id for e in catalog.entries}
        self._existing = [fd for fd in scene_fields if fd.id in catalog_ids]
        self._new_fields = [fd for fd in scene_fields if fd.id not in catalog_ids]
        self._catalog = catalog
        self._catalog_expected = catalog_expected

        self._existing_rows: list[dict] = []
        self._new_rows: list[dict] = []

        self._build_ui()

    # ---- UI -----------------------------------------------------------------

    def _build_ui(self) -> None:
        lay = QVBoxLayout(self)
        lay.setSpacing(6)
        # Created before _populate_existing / _refresh_existing_visibility (they call _update_summary).
        self._lbl_summary = QLabel("")
        self._lbl_summary.setStyleSheet("font-style: italic; color: #555;")
        nothing = True

        if self._existing:
            nothing = False
            grp1 = QGroupBox(
                f"Поля, уже есть в каталоге ({len(self._existing)}) "
                f"— обновляются только expected / expected_text"
            )
            v1 = QVBoxLayout(grp1)
            v1.setContentsMargins(4, 4, 4, 4)

            tb1 = QHBoxLayout()
            btn_all1 = QPushButton("Выбрать все")
            btn_none1 = QPushButton("Снять все")
            self._chk_show_nodiff = QCheckBox("Показать без изменений")
            self._chk_show_nodiff.setChecked(False)
            lbl_note1 = QLabel("label / field_type — источник каталог, не изменяются")
            lbl_note1.setStyleSheet("color: #666; font-size: 10px;")
            tb1.addWidget(btn_all1)
            tb1.addWidget(btn_none1)
            tb1.addSpacing(12)
            tb1.addWidget(self._chk_show_nodiff)
            tb1.addSpacing(12)
            tb1.addWidget(lbl_note1)
            tb1.addStretch()
            v1.addLayout(tb1)

            self._tbl1 = QTableWidget(0, 7)
            self._tbl1.setHorizontalHeaderLabels([
                "", "id",
                "expected (каталог)", "expected (шаблон)",
                "expected_text (каталог)", "expected_text (шаблон)",
                "label / field_type (инфо)",
            ])
            hh1 = self._tbl1.horizontalHeader()
            hh1.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
            hh1.setStretchLastSection(True)
            hh1.resizeSection(0, 28)
            hh1.resizeSection(1, 160)
            hh1.resizeSection(2, 90)
            hh1.resizeSection(3, 90)
            hh1.resizeSection(4, 130)
            hh1.resizeSection(5, 130)
            self._tbl1.setEditTriggers(
                QAbstractItemView.EditTrigger.NoEditTriggers
            )
            self._tbl1.setSelectionMode(
                QAbstractItemView.SelectionMode.NoSelection
            )
            self._tbl1.verticalHeader().setVisible(False)
            self._tbl1.setMaximumHeight(200)
            v1.addWidget(self._tbl1)

            self._populate_existing()
            self._refresh_existing_visibility()

            btn_all1.clicked.connect(lambda: self._sel_existing(True))
            btn_none1.clicked.connect(lambda: self._sel_existing(False))
            self._chk_show_nodiff.stateChanged.connect(self._refresh_existing_visibility)
            lay.addWidget(grp1)

        if self._new_fields:
            nothing = False
            grp2 = QGroupBox(
                f"Новые поля (нет в каталоге) ({len(self._new_fields)})"
            )
            v2 = QVBoxLayout(grp2)
            v2.setContentsMargins(4, 4, 4, 4)

            tb2 = QHBoxLayout()
            btn_all2 = QPushButton("Выбрать все")
            btn_none2 = QPushButton("Снять все")
            lbl_note2 = QLabel(
                "«— новая запись —»: создать запись в каталоге  |  "
                "выбрать запись каталога: поле на canvas получит id / label / field_type из каталога"
            )
            lbl_note2.setStyleSheet("color: #666; font-size: 10px;")
            tb2.addWidget(btn_all2)
            tb2.addWidget(btn_none2)
            tb2.addSpacing(12)
            tb2.addWidget(lbl_note2)
            tb2.addStretch()
            v2.addLayout(tb2)

            self._tbl2 = QTableWidget(0, 7)
            self._tbl2.setHorizontalHeaderLabels([
                "", "id (шаблон)", "label", "field_type",
                "expected", "expected_text",
                "Привязать к записи каталога",
            ])
            hh2 = self._tbl2.horizontalHeader()
            hh2.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
            hh2.setStretchLastSection(True)
            hh2.resizeSection(0, 28)
            hh2.resizeSection(1, 150)
            hh2.resizeSection(2, 130)
            hh2.resizeSection(3, 75)
            hh2.resizeSection(4, 75)
            hh2.resizeSection(5, 110)
            self._tbl2.setEditTriggers(
                QAbstractItemView.EditTrigger.NoEditTriggers
            )
            self._tbl2.setSelectionMode(
                QAbstractItemView.SelectionMode.NoSelection
            )
            self._tbl2.verticalHeader().setVisible(False)
            self._tbl2.setMinimumHeight(140)
            v2.addWidget(self._tbl2)

            self._populate_new()

            btn_all2.clicked.connect(lambda: self._sel_new(True))
            btn_none2.clicked.connect(lambda: self._sel_new(False))
            lay.addWidget(grp2)

        if nothing:
            lay.addWidget(QLabel("Нет полей для импорта."))

        lay.addWidget(self._lbl_summary)
        self._update_summary()

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    # ---- populate -----------------------------------------------------------

    @staticmethod
    def _make_chk_cell(checked: bool, enabled: bool = True):
        """Return (wrapper QWidget, QCheckBox) for inserting into table cell."""
        chk = QCheckBox()
        chk.setChecked(checked)
        chk.setEnabled(enabled)
        w = QWidget()
        lay = QHBoxLayout(w)
        lay.setContentsMargins(4, 0, 0, 0)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(chk)
        return w, chk

    def _populate_existing(self) -> None:
        self._tbl1.setRowCount(0)
        self._existing_rows.clear()

        diff_bg = QColor(255, 240, 150)
        same_bg = QColor(235, 235, 235)
        gray_fg = QColor(130, 130, 130)
        warn_fg = QColor(160, 90, 0)

        for fd in self._existing:
            fid = fd.id
            cat_exp, cat_et = self._catalog_expected.get(fid, ("optional", ""))
            tmpl_exp = normalize_expected(fd.expected)
            tmpl_et = (fd.expected_text or "").strip()
            has_diff = (cat_exp != tmpl_exp) or (cat_et != tmpl_et)

            rinfo = {
                "id": fid,
                "expected": tmpl_exp,
                "expected_text": tmpl_et,
                "has_diff": has_diff,
            }
            self._existing_rows.append(rinfo)

            r = self._tbl1.rowCount()
            self._tbl1.insertRow(r)
            rinfo["row_idx"] = r

            chk_w, chk = self._make_chk_cell(has_diff, enabled=has_diff)
            self._tbl1.setCellWidget(r, 0, chk_w)
            rinfo["chk"] = chk
            chk.stateChanged.connect(self._update_summary)

            self._tbl1.setItem(r, 1, QTableWidgetItem(fid))

            def _it(text: str, bg=None, fg=None) -> QTableWidgetItem:
                item = QTableWidgetItem(text)
                if bg is not None:
                    item.setBackground(bg)
                if fg is not None:
                    item.setForeground(fg)
                return item

            self._tbl1.setItem(r, 2, _it(cat_exp, fg=gray_fg))
            self._tbl1.setItem(
                r, 3, _it(tmpl_exp, bg=diff_bg if cat_exp != tmpl_exp else same_bg)
            )
            self._tbl1.setItem(r, 4, _it(cat_et, fg=gray_fg))
            self._tbl1.setItem(
                r, 5, _it(tmpl_et, bg=diff_bg if cat_et != tmpl_et else same_bg)
            )

            # label / field_type info (catalog-owned, not changed)
            cat_entry = next((e for e in self._catalog.entries if e.id == fid), None)
            cat_lbl = cat_entry.label if cat_entry else ""
            cat_ft = cat_entry.default_field_type if cat_entry else ""
            tmpl_lbl = (fd.label or "").strip()
            tmpl_ft = normalize_field_type(fd.field_type)
            diffs: list[str] = []
            if cat_lbl != tmpl_lbl:
                diffs.append(f"label: кат={cat_lbl!r} шабл={tmpl_lbl!r}")
            if cat_ft != tmpl_ft:
                diffs.append(f"type: кат={cat_ft!r} шабл={tmpl_ft!r}")
            info_txt = "; ".join(diffs) if diffs else "совпадает"
            it_info = _it(info_txt, fg=warn_fg if diffs else gray_fg)
            it_info.setToolTip(
                "label и field_type принадлежат каталогу и не изменяются при импорте.\n"
                "Используйте «Синхр.», чтобы применить значения каталога к полям шаблона."
            )
            self._tbl1.setItem(r, 6, it_info)

    def _refresh_existing_visibility(self, _=None) -> None:
        show_all = (
            self._chk_show_nodiff.isChecked()
            if hasattr(self, "_chk_show_nodiff")
            else False
        )
        for rinfo in self._existing_rows:
            r = rinfo.get("row_idx", -1)
            if r < 0:
                continue
            self._tbl1.setRowHidden(r, not rinfo["has_diff"] and not show_all)
        self._update_summary()

    def _sel_existing(self, checked: bool) -> None:
        for rinfo in self._existing_rows:
            if rinfo["has_diff"]:
                chk = rinfo.get("chk")
                if chk:
                    chk.setChecked(checked)

    def _populate_new(self) -> None:
        self._tbl2.setRowCount(0)
        self._new_rows.clear()

        bind_display = ["— новая запись —"] + [
            f"{e.id}  ({e.label})" if e.label else e.id
            for e in self._catalog.entries
        ]
        bind_ids = [_BIND_NEW_SENTINEL] + [e.id for e in self._catalog.entries]

        for fd in self._new_fields:
            rinfo: dict = {"fd": fd, "bind_id": None}
            self._new_rows.append(rinfo)

            r = self._tbl2.rowCount()
            self._tbl2.insertRow(r)
            rinfo["row_idx"] = r

            chk_w, chk = self._make_chk_cell(True)
            self._tbl2.setCellWidget(r, 0, chk_w)
            rinfo["chk"] = chk
            chk.stateChanged.connect(self._update_summary)

            self._tbl2.setItem(r, 1, QTableWidgetItem(fd.id or ""))
            self._tbl2.setItem(r, 2, QTableWidgetItem(fd.label or ""))
            self._tbl2.setItem(r, 3, QTableWidgetItem(normalize_field_type(fd.field_type)))
            self._tbl2.setItem(r, 4, QTableWidgetItem(normalize_expected(fd.expected)))
            self._tbl2.setItem(r, 5, QTableWidgetItem(fd.expected_text or ""))

            cmb = QComboBox()
            if fd.document_property:
                cmb.addItem("— только новая запись (свойство PDF) —")
                cmb.setItemData(0, _BIND_NEW_SENTINEL)
                cmb.setEnabled(False)
                rinfo["bind_id"] = None
            else:
                for i, (txt, bid) in enumerate(zip(bind_display, bind_ids)):
                    cmb.addItem(txt)
                    cmb.setItemData(i, bid)
                cmb.setCurrentIndex(0)
                cmb.currentIndexChanged.connect(
                    lambda _idx, ri=rinfo, c=cmb: self._on_bind_changed(ri, c)
                )
            self._tbl2.setCellWidget(r, 6, cmb)

    def _on_bind_changed(self, rinfo: dict, cmb: QComboBox) -> None:
        bid = cmb.currentData()
        rinfo["bind_id"] = bid if bid else None
        self._update_summary()

    def _sel_new(self, checked: bool) -> None:
        for rinfo in self._new_rows:
            chk = rinfo.get("chk")
            if chk:
                chk.setChecked(checked)

    def _update_summary(self, _=None) -> None:
        n_update = sum(
            1 for ri in self._existing_rows
            if ri["has_diff"] and ri.get("chk") and ri["chk"].isChecked()
        )
        n_new = sum(
            1 for ri in self._new_rows
            if ri.get("chk") and ri["chk"].isChecked() and not ri["bind_id"]
        )
        n_bind = sum(
            1 for ri in self._new_rows
            if ri.get("chk") and ri["chk"].isChecked() and ri["bind_id"]
        )
        parts: list[str] = []
        if n_update:
            parts.append(f"обновить expected для {n_update} полей")
        if n_new:
            parts.append(f"создать {n_new} новых записей")
        if n_bind:
            parts.append(f"привязать {n_bind} полей к записям каталога")
        text = "Итого: " + (", ".join(parts) if parts else "ничего не изменится") + "."
        self._lbl_summary.setText(text)

    # ---- result -------------------------------------------------------------

    def get_result(self) -> tuple:
        """Return (existing_updates, bind_updates, new_entries).

        existing_updates: list of (id, expected, expected_text)
        bind_updates:     list of (scene_field_id, catalog_entry_id)
        new_entries:      list of FieldCatalogEntry
        """
        existing_updates = []
        for ri in self._existing_rows:
            if ri["has_diff"] and ri.get("chk") and ri["chk"].isChecked():
                existing_updates.append((ri["id"], ri["expected"], ri["expected_text"]))

        bind_updates = []
        new_entries = []
        for ri in self._new_rows:
            if not (ri.get("chk") and ri["chk"].isChecked()):
                continue
            fd = ri["fd"]
            bid = ri["bind_id"]
            if bid:
                bind_updates.append((fd.id, bid))
            else:
                new_entries.append(
                    FieldCatalogEntry(
                        id=fd.id,
                        label=fd.label or "",
                        short_num=0,
                        default_field_type=normalize_field_type(fd.field_type),
                        is_document_property=bool(fd.document_property),
                    )
                )
        return existing_updates, bind_updates, new_entries


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _make_combo(options: list[str], current: str) -> QComboBox:
    cmb = QComboBox()
    cmb.setEditable(True)
    cmb.addItems(options)
    idx = cmb.findText(current)
    if idx >= 0:
        cmb.setCurrentIndex(idx)
    else:
        cmb.setCurrentText(current)
    return cmb
