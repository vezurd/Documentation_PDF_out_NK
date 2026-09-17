"""QMainWindow: меню, layout, координация редактора шаблонов."""

from __future__ import annotations

import copy
import os
import subprocess
import tempfile
import traceback
from collections import Counter
from dataclasses import dataclass, replace
from typing import Any

import fitz
from PySide6.QtCore import Qt, QPointF, QRectF, QTimer, QSettings
from PySide6.QtGui import QAction, QKeySequence, QColor, QPen, QBrush, QFont, QIcon, QShortcut
from PySide6.QtWidgets import (
    QMainWindow,
    QSplitter,
    QToolBar,
    QToolButton,
    QMenu,
    QDoubleSpinBox,
    QLabel,
    QFileDialog,
    QMessageBox,
    QComboBox,
    QWidget,
    QVBoxLayout,
    QSizePolicy,
    QGraphicsItem,
    QGraphicsLineItem,
    QGraphicsRectItem,
    QGraphicsSimpleTextItem,
    QPushButton,
    QCheckBox,
    QAbstractSpinBox,
    QLineEdit,
    QTextEdit,
)

from pdf_parsing_v2_engine.frame_detector import find_frame, SCALE
from pdf_parsing_v2_engine.models import (
    FieldCatalog,
    FieldDef,
    FrameInfo,
    StampTemplate,
    normalize_expected,
    normalize_field_type,
)
from .graphics_view import StampGraphicsView
from pdf_parsing_v2_engine.grid_matcher import (
    cluster_lines,
    adapt_by_cell_assignment,
    _make_unique_keys,
)


@dataclass
class _ConsolidationState:
    x_lines: list[float]
    y_lines: list[float]
    tolerance_px: float
    bad_item_ids: set[int]
    stamp_scene_rect: QRectF


def _log(msg: str, level: str = "info") -> None:
    tag = {"info": "INFO", "warn": "WARN", "error": "ERROR"}.get(level, "INFO")
    print(f"[TemplateEditor] {tag}: {msg}", flush=True)


_RECENT_PDF_MAX = 5
_SETTINGS_ORG = "Documentation_PDF_out_NK"
_SETTINGS_APP = "TemplateEditorV2"
_SETTINGS_RECENT_KEY = "recent/pdf_paths"
_SETTINGS_AUTOLOAD_MATCHING_TEMPLATE_KEY = "ui/auto_load_matching_template"
# True: при включении «Границы» показывать отдельное окно; False — нижняя зона сплиттера (после «Вернуть в панель»)
_SETTINGS_PREALIGN_OPEN_FLOATING = "prealign_borders_open_floating"


def load_template_editor_app_icon() -> QIcon | None:
    """Return packaged ``ico/editor_favicon.ico`` if present."""
    ico = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ico", "editor_favicon.ico")
    if os.path.isfile(ico):
        return QIcon(ico)
    return None


class TemplateEditorWindow(QMainWindow):
    def __init__(self, cfg: dict[str, Any] | None = None, parent=None):
        super().__init__(parent)

        _icon = load_template_editor_app_icon()
        if _icon is not None:
            self.setWindowIcon(_icon)

        self._cfg: dict[str, Any] = cfg or {}
        self._templates_dir: str = self._cfg.get("templates_dir", "")

        # State
        self._fitz_doc: fitz.Document | None = None
        self._pdf_path: str = ""
        self._catalog: FieldCatalog | None = None
        self._catalog_path: str = ""
        self._frame_info: FrameInfo | None = None
        self._active_template_path: str = ""
        self._is_modified: bool = False
        self._catalog_is_dirty: bool = False
        self._geo_tracked_item = None  # FieldRectItem | None

        # Catalog panel state (T7.2)
        self._catalog_panel = None           # CatalogEditorPanel | None
        self._catalog_floating_win = None    # QWidget | None (detached window)
        self._catalog_is_floating: bool = False
        self._prealign_floating_win = None   # QWidget | None (detached prealign panel)
        self._prealign_is_floating: bool = False
        self._last_test_png: str = ""  # path to last generated test debug PNG
        self._meta_panel = None  # TemplateMetaPanel — created in _wire_panels
        self._consolidation_state: _ConsolidationState | None = None
        self._consolidation_overlay: list[object] = []
        self._adapt_overlay: list[object] = []
        self._prealign_overlay_items: list[object] = []
        self._last_prealign_debug_bundle: dict[str, Any] | None = None
        self._prealign_overlay_panel = None  # PrealignOverlayPanel | None
        self._detected_grid_overlay: list[object] = []
        self._origin_marker_items: list[object] = []
        self._origin_line_item: object = None
        self._scene_field_count: int = 0
        self._pending_catalog_field = None  # FieldDef | None — waiting for place_confirmed
        self._grid_lines: list = []  # list[TemplateGridLine] — wireframe lines for the current template
        self._wireframe_items: list = []  # QGraphicsItem overlay for wireframe mode
        self._snapped_line_positions: dict[str, float] = {}  # after Тест-адаптация
        self._adapt_transform: tuple[float, float, float, float] = (1.0, 1.0, 0.0, 0.0)  # sx, sy, dx, dy
        self._last_ca_info = None  # CellAssignmentInfo from last adapt
        self._pick_line_which: str = ""  # "start" or "end" — active pick-line target
        self._pick_line_hidden_items: list = []  # FieldRectItems hidden during pick mode
        self._recent_pdf_paths: list[str] = []
        # PDF document-property FieldDefs (61/62/…/file_name), not shown on canvas; persisted on save.
        self._document_property_field_defs: list[FieldDef] = []
        # Templates matching current PDF page + toolbar doc_type (``select_templates`` order).
        self._matching_templates: list[StampTemplate] = []

        # Debounced catalog coverage timer (T7.1): deduplicates rapid-fire scene
        # mutations (e.g. deleting many fields at once) into a single coverage refresh.
        self._catalog_coverage_timer = QTimer(self)
        self._catalog_coverage_timer.setSingleShot(True)
        self._catalog_coverage_timer.setInterval(100)
        self._catalog_coverage_timer.timeout.connect(self._update_catalog_panel_coverage)

        self._set_active_template(None)
        self.resize(1600, 900)

        self._init_actions()
        self._recent_pdf_paths = self._load_recent_pdf_paths()
        self._init_toolbar()
        self._init_central()
        self._init_statusbar()
        self._init_menus()

        self._refresh_matching_templates_ui(run_autoload=False)

        _log(
            f"Редактор запущен. templates_dir={self._templates_dir!r}",
        )

    # ================================================================== log
    def _log(self, msg: str, level: str = "info") -> None:
        _log(msg, level)

    # ================================================================== modified / title
    def _set_modified(self, modified: bool) -> None:
        self._is_modified = modified
        self._update_window_modified_state()
        if self._meta_panel is not None:
            self._meta_panel.set_modified(modified)

    def _meta_find_tables_snap_kwargs(self) -> dict[str, float | None]:
        """Optional find_tables snap overrides from the template meta panel."""
        meta = self._meta_panel
        if meta is None:
            return {
                "find_tables_snap_x_tolerance": None,
                "find_tables_snap_y_tolerance": None,
            }
        return {
            "find_tables_snap_x_tolerance": meta.get_find_tables_snap_x_tolerance(),
            "find_tables_snap_y_tolerance": meta.get_find_tables_snap_y_tolerance(),
        }

    def _frame_detection_template_stub(self) -> StampTemplate | None:
        """StampTemplate slice for ``drawing_union`` ROI (inner fields vs GOST)."""
        if self._meta_panel is None or self._meta_panel.get_frame_mode() != "drawing_union":
            return None
        from .cell_items import FieldRectItem

        items = [
            it for it in self._gfx_view.scene().items()
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def is not None
        ]
        fields: list[FieldDef] = self._collect_field_defs_from_scene(items) if items else []
        return StampTemplate(
            schema_version=2,
            name="(editor-frame)",
            doc_types=[],
            page_selector="all",
            fields=fields,
            frame_mode="drawing_union",
        )

    def _recompute_and_apply_frame(self, fitz_page: fitz.Page) -> None:
        stub = self._frame_detection_template_stub()
        if stub is not None:
            frame, _ = find_frame(fitz_page, frame_mode="drawing_union", template=stub)
        else:
            frame, _ = find_frame(fitz_page)
        self._frame_info = frame
        self._gfx_view.load_page(fitz_page, frame)
        self._draw_origin_markers(frame)

    def _update_window_modified_state(self) -> None:
        """Asterisk in title: template and/or catalog has unsaved edits."""
        self.setWindowModified(self._is_modified or self._catalog_is_dirty)

    def _on_catalog_dirty_changed(self, dirty: bool) -> None:
        self._catalog_is_dirty = dirty
        self._update_window_modified_state()

    @staticmethod
    def _item_scene_rect(it) -> tuple[float, float, float, float]:
        """Return (x, y, w, h) of item in scene coords WITHOUT pen inflation.

        ``sceneBoundingRect()`` includes half-pen-width on each side, causing
        bbox drift on every save.  Use ``scenePos() + rect()`` instead.
        """
        pos = it.scenePos()
        r = it.rect()
        return (pos.x() + r.x(), pos.y() + r.y(), r.width(), r.height())

    def _collect_field_defs_from_scene(self, items: list) -> list:
        """Пересчитать FieldDef из текущих rect сцены (как при сохранении в JSON).

        Drag/resize меняют только QGraphicsItem; без этого bbox_mm в field_def устаревает.
        Учитывает per-field origin при обратном пересчёте координат.
        """
        from pdf_parsing_v2_engine.coord_transform import absolute_to_field_bbox_mm

        frame = self._frame_info
        dpi_scale = self._gfx_view._dpi / 72.0
        out: list[FieldDef] = []
        for it in items:
            fd = it.field_def
            if fd is None:
                continue
            if frame is not None:
                sx, sy, sw, sh = self._item_scene_rect(it)
                fitz_x0 = sx / dpi_scale
                fitz_y0 = sy / dpi_scale
                fitz_x1 = (sx + sw) / dpi_scale
                fitz_y1 = (sy + sh) / dpi_scale
                pm_x0 = fitz_x0
                pm_y0 = frame.page_height - fitz_y1
                pm_x1 = fitz_x1
                pm_y1 = frame.page_height - fitz_y0
                bbox_mm = absolute_to_field_bbox_mm(
                    fd.origin, frame, pm_x0, pm_y0, pm_x1, pm_y1,
                )
                out.append(
                    FieldDef(
                        id=fd.id,
                        label=fd.label,
                        bbox_mm=bbox_mm,
                        clean=fd.clean,
                        validate_regex=fd.validate_regex,
                        expected=fd.expected,
                        padding_mm=fd.padding_mm,
                        field_type=fd.field_type,
                        expected_text=fd.expected_text,
                        is_anchor=fd.is_anchor,
                        outside_stamp=fd.outside_stamp,
                        origin=fd.origin,
                        stretch_to_page=fd.stretch_to_page,
                        bound_top=fd.bound_top,
                        bound_bottom=fd.bound_bottom,
                        bound_left=fd.bound_left,
                        bound_right=fd.bound_right,
                    )
                )
            else:
                out.append(fd)
        return out

    def _next_unique_field_id_for_duplicate(self, base_id: str) -> str:
        from .cell_items import FieldRectItem

        used = {
            it.field_def.id
            for it in self._scene_field_items()
            if isinstance(it, FieldRectItem) and it.field_def
        }
        cat = self._catalog
        if cat:
            ent = cat.get_entry(base_id)
            if ent is not None:
                grp = getattr(ent, "group", "") or ""
                peers = [e for e in cat.entries if (getattr(e, "group", "") or "") == grp]
                peers.sort(key=lambda e: (e.short_num, e.id))
                for e in peers:
                    if e.id not in used:
                        return e.id
        n = 2
        while True:
            cand = f"{base_id}_{n}"
            if cand not in used:
                return cand
            n += 1

    def _clear_duplicate_id_highlights(self) -> None:
        from .cell_items import FieldRectItem

        for it in self._scene_field_items():
            if isinstance(it, FieldRectItem):
                it.set_id_duplicate_highlight(False)

    def _highlight_duplicate_field_items(self, items, fields) -> None:
        from .cell_items import FieldRectItem

        ids = [fd.id for fd in fields]
        cnt = Counter(ids)
        dup = {i for i, n in cnt.items() if n > 1}
        for it, fd in zip(items, fields):
            if isinstance(it, FieldRectItem) and fd.id in dup:
                it.set_id_duplicate_highlight(True)
        if self._cells_panel:
            self._rebuild_cells_list()

    def _auto_fix_duplicate_field_ids_on_scene(self) -> None:
        from .cell_items import FieldRectItem

        items = [
            it for it in self._scene_field_items()
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def
        ]
        by_id: dict[str, list] = {}
        for it in items:
            fid = it.field_def.id
            by_id.setdefault(fid, []).append(it)
        for fid, group in by_id.items():
            if len(group) <= 1:
                continue
            for it in group[1:]:
                new_id = self._next_unique_field_id_for_duplicate(fid)
                fd = copy.deepcopy(it.field_def)
                fd.id = new_id
                it.assign_field(fd)
        self._set_modified(True)
        if self._cells_panel:
            self._rebuild_cells_list()

    def _validate_template_ids_before_save(self, items, fields: list[FieldDef]) -> bool:
        from .cell_items import FieldRectItem

        ids = [fd.id for fd in fields]
        cnt = Counter(ids)
        dup_ids = [i for i, n in cnt.items() if n > 1]
        while dup_ids:
            box = QMessageBox(self)
            box.setWindowTitle("Дубликаты id полей")
            box.setIcon(QMessageBox.Icon.Warning)
            box.setText(
                "Повторяются id полей:\n"
                + ", ".join(dup_ids[:25])
                + ("…" if len(dup_ids) > 25 else "")
            )
            btn_fix = box.addButton("Автоисправить", QMessageBox.ButtonRole.AcceptRole)
            btn_show = box.addButton("Показать на сцене", QMessageBox.ButtonRole.ActionRole)
            box.addButton("Отмена", QMessageBox.ButtonRole.RejectRole)
            box.exec()
            clicked = box.clickedButton()
            if clicked == btn_fix:
                self._auto_fix_duplicate_field_ids_on_scene()
                items = [
                    it for it in self._gfx_view.scene().items()
                    if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def is not None
                ]
                fields = self._collect_field_defs_from_scene(items)
                ids = [fd.id for fd in fields]
                cnt = Counter(ids)
                dup_ids = [i for i, n in cnt.items() if n > 1]
                if not dup_ids:
                    self._clear_duplicate_id_highlights()
                    break
                QMessageBox.warning(
                    self,
                    "Сохранение",
                    "Не удалось устранить дубликаты id автоматически.",
                )
                return False
            if clicked == btn_show:
                self._highlight_duplicate_field_items(items, fields)
                return False
            return False

        self._clear_duplicate_id_highlights()

        cat = self._catalog
        if cat and cat.entries:
            catalog_ids = set(cat.all_ids())
            unknown = [i for i in ids if i not in catalog_ids]
            if unknown:
                r = QMessageBox.question(
                    self,
                    "Каталог полей",
                    f"Есть поля, не перечисленные в каталоге:\n"
                    f"{', '.join(unknown[:20])}"
                    f"{'…' if len(unknown) > 20 else ''}\n\n"
                    "Сохранить всё равно?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if r != QMessageBox.StandardButton.Yes:
                    return False
        return True

    def _set_active_template(self, path: str | None) -> None:
        self._active_template_path = path or ""
        name = os.path.basename(path) if path else "— без имени —"
        self.setWindowTitle(f"Редактор шаблонов v2 — {name}[*]")
        if hasattr(self, "_lbl_active_template"):
            self._lbl_active_template.setText(f" Активный: {name}")
            self._lbl_active_template.setToolTip(
                os.path.abspath(path) if path else "Файл шаблона не выбран",
            )
        if hasattr(self, "_templates_browser") and self._templates_browser:
            self._templates_browser.set_active_template(path or "")

    # ================================================================== actions
    def _init_actions(self):
        self.act_open_pdf = QAction("Открыть PDF…", self)
        self.act_open_pdf.setShortcut(QKeySequence("Ctrl+P"))
        self.act_open_pdf.triggered.connect(self._on_open_pdf)

        self.act_load_template = QAction("Загрузить шаблон…", self)
        self.act_load_template.setShortcut(QKeySequence.StandardKey.Open)
        self.act_load_template.triggered.connect(self._on_load_template)

        self.act_save_template = QAction("Сохранить шаблон", self)
        self.act_save_template.setShortcut(QKeySequence.StandardKey.Save)
        self.act_save_template.triggered.connect(self._on_save_template)

        self.act_save_template_as = QAction("Сохранить шаблон как…", self)
        self.act_save_template_as.setShortcut(QKeySequence("Ctrl+Shift+S"))
        self.act_save_template_as.triggered.connect(self._on_save_template_as)

        self.act_new = QAction("Новый", self)
        self.act_new.setShortcut(QKeySequence.StandardKey.New)
        self.act_new.triggered.connect(self._on_new)

        self.act_auto_detect = QAction("Авто-разметка", self)
        self.act_auto_detect.setShortcut(QKeySequence("Ctrl+D"))
        self.act_auto_detect.triggered.connect(self._on_auto_detect)

        self.act_add_field_draw = QAction("Добавить поле", self)
        self.act_add_field_draw.setCheckable(True)
        self.act_add_field_draw.setToolTip("Нарисовать новое поле на PDF")
        self.act_add_field_draw.toggled.connect(self._on_toggle_add_field_draw)

        self.act_test = QAction("Тестировать (F5)", self)
        self.act_test.setShortcut(QKeySequence("F5"))
        self.act_test.triggered.connect(self._on_test)

        self.act_consolidate_grid = QAction("Консолидировать сетку", self)
        self.act_consolidate_grid.setToolTip("Показать логические линии сетки и проблемные ячейки")
        self.act_consolidate_grid.triggered.connect(self._on_consolidate_grid)

        self.act_apply_consolidation = QAction("Применить", self)
        self.act_apply_consolidation.setToolTip("Подтянуть границы ячеек к логическим линиям")
        self.act_apply_consolidation.setEnabled(False)
        self.act_apply_consolidation.triggered.connect(self._on_apply_consolidation)

        self.act_adapt_to_grid = QAction("Тест-адаптация", self)
        self.act_adapt_to_grid.setToolTip("Адаптировать поля и каркас шаблона к штампу текущей страницы PDF")
        self.act_adapt_to_grid.setEnabled(False)
        self.act_adapt_to_grid.triggered.connect(self._on_adapt_to_grid)

        self.act_fill_stamp_gaps = QAction("Заполнить зазоры до соседей", self)
        self.act_fill_stamp_gaps.setToolTip(
            "Пост-проход: расширить поля внутри габарита штампа до ближайших "
            "соседей по строке/столбцу (если PDF разбил одну ячейку на несколько)."
        )
        self.act_fill_stamp_gaps.triggered.connect(self._on_fill_stamp_gaps)

        self.act_clear_unassigned = QAction("Очистить неназначенные", self)
        self.act_clear_unassigned.triggered.connect(self._on_clear_unassigned)

        self.act_delete_selected = QAction("Удалить выделенные", self)
        self.act_delete_selected.setShortcut(QKeySequence.StandardKey.Delete)
        self.act_delete_selected.triggered.connect(self._on_delete_selected)

        self.act_catalog_editor = QAction("Каталог полей (Ctrl+K)", self)
        self.act_catalog_editor.setShortcut(QKeySequence("Ctrl+K"))
        self.act_catalog_editor.triggered.connect(self._on_catalog_editor)

        self.act_set_editor = QAction("Набор шаблонов…", self)
        self.act_set_editor.triggered.connect(self._on_set_editor)

        self.act_adapt_debug_dump = QAction("Дамп отладки адаптации", self)
        self.act_adapt_debug_dump.setToolTip(
            "Сохранить полный снимок: координаты полей, detected cells, "
            "grid lines, результат адаптации — для анализа"
        )
        self.act_adapt_debug_dump.triggered.connect(self._on_adapt_debug_dump)

        self.act_open_test_png = QAction("Открыть PNG теста", self)
        self.act_open_test_png.setToolTip("Открыть последний debug PNG (со спанами текста PDF)")
        self.act_open_test_png.setEnabled(False)
        self.act_open_test_png.triggered.connect(self._on_open_test_png)

        self.act_wireframe_mode = QAction("Wireframe", self)
        self.act_wireframe_mode.setCheckable(True)
        self.act_wireframe_mode.setToolTip(
            "Режим wireframe: управление линиями сетки и привязками полей к линиям"
        )
        self.act_wireframe_mode.toggled.connect(self._on_wireframe_mode_toggled)

        self.act_prealign_borders = QAction("Границы", self)
        self.act_prealign_borders.setCheckable(True)
        self.act_prealign_borders.setChecked(False)
        self.act_prealign_borders.setToolTip(
            "Слои prealign / adapt на PDF: по умолчанию отдельное окно; кнопка ⬇/⬆ в панели — вернуть вниз слева"
        )
        self.act_prealign_borders.toggled.connect(self._on_prealign_panel_toggle)

    # ================================================================== toolbar
    def _init_toolbar(self):
        # Две строки панелей: меньше горизонтальная перегрузка и меньше «сворачивания» в меню >>.
        tb_main = QToolBar("Редактирование", self)
        tb_main.setMovable(False)
        self.addToolBar(tb_main)

        self._menu_recent_pdfs = QMenu(self)
        self._menu_recent_pdfs.setToolTipsVisible(True)
        self._menu_recent_pdfs.aboutToShow.connect(self._populate_recent_pdf_menu)
        self._menu_recent_pdfs.triggered.connect(self._on_recent_pdf_menu_triggered)
        self._btn_open_pdf = QToolButton(self)
        self._btn_open_pdf.setDefaultAction(self.act_open_pdf)
        self._btn_open_pdf.setMenu(self._menu_recent_pdfs)
        self._btn_open_pdf.setPopupMode(
            QToolButton.ToolButtonPopupMode.MenuButtonPopup,
        )
        tb_main.addWidget(self._btn_open_pdf)
        self._populate_recent_pdf_menu()

        tb_main.addAction(self.act_auto_detect)
        tb_main.addAction(self.act_add_field_draw)
        tb_main.addSeparator()
        tb_main.addAction(self.act_test)
        tb_main.addSeparator()
        tb_main.addAction(self.act_consolidate_grid)
        tb_main.addAction(self.act_apply_consolidation)
        tb_main.addAction(self.act_adapt_to_grid)
        tb_main.addAction(self.act_fill_stamp_gaps)
        tb_main.addSeparator()
        tb_main.addAction(self.act_wireframe_mode)
        tb_main.addAction(self.act_prealign_borders)

        self.addToolBarBreak(Qt.ToolBarArea.TopToolBarArea)

        tb_extra = QToolBar("Параметры", self)
        tb_extra.setMovable(False)
        self.addToolBar(tb_extra)

        tb_extra.addWidget(QLabel(" Tol мм: "))
        self._grid_tol_spin = QDoubleSpinBox()
        self._grid_tol_spin.setRange(0.1, 20.0)
        self._grid_tol_spin.setDecimals(1)
        self._grid_tol_spin.setValue(2.0)
        self._grid_tol_spin.setSingleStep(0.1)
        self._grid_tol_spin.valueChanged.connect(self._on_grid_tolerance_changed)
        tb_extra.addWidget(self._grid_tol_spin)
        tb_extra.addSeparator()
        tb_extra.addAction(self.act_adapt_debug_dump)
        tb_extra.addSeparator()
        tb_extra.addAction(self.act_clear_unassigned)
        tb_extra.addAction(self.act_delete_selected)
        tb_extra.addSeparator()
        tb_extra.addAction(self.act_open_test_png)

        tb_extra.addSeparator()
        tb_extra.addWidget(QLabel(" Страница: "))
        self._page_combo = QComboBox()
        self._page_combo.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToContents,
        )
        self._page_combo.setMinimumContentsLength(3)
        self._page_combo.setEnabled(False)
        self._page_combo.currentIndexChanged.connect(self._on_page_changed)
        tb_extra.addWidget(self._page_combo)

        tb_extra.addSeparator()
        tb_extra.addWidget(QLabel(" Тип: "))
        self._doc_type_combo = QComboBox()
        self._doc_type_combo.addItems(
            [
                "WIR",
                "LAY",
                "CAE",
                "GA",
                "PL",
                "DW",
                "NI",
                "BOE",
                "BOM",
                "BOQ",
                "MTO",
                "OD",
                "CJ",
            ],
        )
        self._doc_type_combo.setCurrentText("WIR")
        self._doc_type_combo.setEditable(True)
        self._doc_type_combo.currentTextChanged.connect(
            self._on_doc_type_for_matching_changed,
        )
        tb_extra.addWidget(self._doc_type_combo)

        tb_extra.addSeparator()
        tb_extra.addWidget(QLabel(" Шаблон: "))
        self._matching_templates_combo = QComboBox()
        self._matching_templates_combo.setMinimumWidth(200)
        self._matching_templates_combo.setMaximumWidth(420)
        self._matching_templates_combo.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon,
        )
        self._matching_templates_combo.setMinimumContentsLength(24)
        self._matching_templates_combo.setToolTip(
            "Шаблоны, подходящие для типа документа и страницы "
            "(порядок по убыванию приоритета, как в pipeline v2)",
        )
        tb_extra.addWidget(self._matching_templates_combo)

        self._btn_load_matching_template = QPushButton("Загрузить шаблон")
        self._btn_load_matching_template.setToolTip(
            "Загрузить первый шаблон из списка (наивысший приоритет)",
        )
        self._btn_load_matching_template.clicked.connect(
            self._on_load_matching_template_clicked,
        )
        tb_extra.addWidget(self._btn_load_matching_template)

        self._chk_autoload_matching_template = QCheckBox("Автозагрузка шаблона")
        self._chk_autoload_matching_template.setToolTip(
            "После открытия PDF или смены страницы загружать первый подходящий шаблон; "
            "при несохранённых изменениях будет запрос подтверждения",
        )
        s_autoload = self._pdf_recent_settings()
        raw_auto = s_autoload.value(_SETTINGS_AUTOLOAD_MATCHING_TEMPLATE_KEY, False)
        if isinstance(raw_auto, str):
            autoload_on = raw_auto.strip().lower() in ("1", "true", "yes", "on")
        else:
            autoload_on = bool(raw_auto)
        self._chk_autoload_matching_template.setChecked(autoload_on)
        self._chk_autoload_matching_template.toggled.connect(
            self._on_autoload_matching_template_toggled,
        )
        tb_extra.addWidget(self._chk_autoload_matching_template)

        spring = QWidget()
        spring.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        spring.setMinimumWidth(0)
        tb_extra.addWidget(spring)

        self._lbl_active_template = QLabel(" Активный: — без имени —")
        self._lbl_active_template.setToolTip("Текущий загруженный файл шаблона")
        self._lbl_active_template.setMinimumWidth(120)
        self._lbl_active_template.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Preferred,
        )
        tb_extra.addWidget(self._lbl_active_template)

    # ================================================================== central
    def _init_central(self):
        from .catalog_editor import CatalogEditorPanel

        # Main horizontal splitter: [left_area | right_panel]
        self._splitter = QSplitter(Qt.Orientation.Horizontal, self)
        self.setCentralWidget(self._splitter)

        # Left area: vertical splitter [browser+canvas | prealign | catalog]
        self._left_v_splitter = QSplitter(Qt.Orientation.Vertical, self)
        self._splitter.addWidget(self._left_v_splitter)

        # Top horizontal splitter: [browser+column | canvas]
        self._top_h_splitter = QSplitter(Qt.Orientation.Horizontal, self)
        from .templates_browser import TemplatesBrowserPanel
        from .prealign_overlay_panel import PrealignOverlayPanel

        self._templates_browser = TemplatesBrowserPanel(self._templates_dir, self)
        self._templates_browser.setMinimumWidth(140)
        self._templates_browser.setMaximumWidth(320)

        self._prealign_overlay_panel = PrealignOverlayPanel(self._pdf_recent_settings(), self)
        self._prealign_overlay_panel.overlay_refresh_requested.connect(
            self._rebuild_prealign_overlay_scene,
        )
        self._prealign_overlay_panel.detach_toggle_requested.connect(
            self._on_prealign_detach_toggle,
        )
        self._prealign_overlay_panel.canvas_layer_visibility_changed.connect(
            self._on_canvas_layer_visibility_changed,
        )

        self._top_h_splitter.addWidget(self._templates_browser)

        # Center: canvas
        self._gfx_view = StampGraphicsView(self)
        self._top_h_splitter.addWidget(self._gfx_view)

        self._top_h_splitter.setStretchFactor(0, 0)  # browser fixed-ish
        self._top_h_splitter.setStretchFactor(1, 6)  # canvas
        self._top_h_splitter.setSizes([200, 900])

        self._left_v_splitter.addWidget(self._top_h_splitter)

        # Prealign overlay — нижняя зона слева (как каталог): по умолчанию скрыта (тулбар «Границы»)
        self._left_v_splitter.addWidget(self._prealign_overlay_panel)
        self._prealign_overlay_panel.setMinimumHeight(0)
        self._prealign_overlay_panel.hide()

        # Catalog panel — bottom of left area, hidden by default
        self._catalog_panel = CatalogEditorPanel(
            catalog=self._catalog,
            templates_dir=self._templates_dir,
            current_path=self._catalog_path,
            parent=self._left_v_splitter,
        )
        self._catalog_panel.setMinimumHeight(80)
        self._catalog_panel.catalog_saved.connect(self._on_catalog_panel_saved)
        self._catalog_panel.navigate_to_field.connect(self._on_catalog_navigate)
        self._catalog_panel.sync_requested.connect(self._on_catalog_sync_requested)
        self._catalog_panel.template_save_requested.connect(
            self._on_template_save_after_catalog_sync,
        )
        self._catalog_panel.catalog_dirty_changed.connect(self._on_catalog_dirty_changed)
        self._catalog_panel.catalog_field_link_apply_requested.connect(
            self._on_catalog_field_link_apply,
        )
        self._catalog_panel.catalog_add_field_to_workspace_requested.connect(
            self._on_catalog_add_field_to_workspace,
        )
        # Connect detach button directly to main_window methods (T7.detach)
        self._catalog_panel._btn_detach.clicked.disconnect()
        self._catalog_panel._btn_detach.clicked.connect(self._on_catalog_detach_toggle)
        self._left_v_splitter.addWidget(self._catalog_panel)
        self._catalog_panel.hide()

        self._left_v_splitter.setStretchFactor(0, 5)   # top: browser + canvas
        self._left_v_splitter.setStretchFactor(1, 2)   # prealign
        self._left_v_splitter.setStretchFactor(2, 2)   # catalog
        self._left_v_splitter.setSizes([720, 0, 1])

        # Right: cells + properties (full height)
        self._right_panel = QWidget(self)
        self._right_layout = QVBoxLayout(self._right_panel)
        self._right_layout.setContentsMargins(2, 2, 2, 2)
        self._cells_panel = None
        self._props_panel = None
        self._splitter.addWidget(self._right_panel)

        self._splitter.setStretchFactor(0, 6)  # left area (browser+canvas+catalog)
        self._splitter.setStretchFactor(1, 3)  # right panel
        self._splitter.setSizes([1100, 400])

        self._gfx_view.mouse_scene_pos_changed.connect(self._on_mouse_moved)
        self._gfx_view.zoom_changed.connect(self._on_zoom_changed)
        self._gfx_view.rect_drawn.connect(self._on_rect_drawn_add_field)
        self._gfx_view.place_confirmed.connect(self._on_place_confirmed)
        self._gfx_view.wire_line_place_finished.connect(self._on_wire_line_place_finished)

        # Browser signals
        self._templates_browser.template_load_requested.connect(self._on_template_browser_load)
        self._templates_browser.catalog_activate_requested.connect(self._on_catalog_browser_activate)
        self._templates_browser.catalog_edit_requested.connect(self._on_catalog_editor)
        self._templates_browser.message.connect(self._on_browser_message)

    def _wire_panels(self):
        """Wire cells_panel, properties_panel and template meta panel."""
        from .cells_panel import CellsPanel
        from .properties_panel import PropertiesPanel
        from .template_meta_panel import TemplateMetaPanel

        self._meta_panel = TemplateMetaPanel(self)
        self._right_layout.addWidget(self._meta_panel, stretch=0)
        self._meta_panel.meta_changed.connect(self._on_meta_changed)
        self._meta_panel.save_requested.connect(self._on_save_template)
        self._grid_tol_spin.setValue(self._meta_panel.get_grid_tolerance_template_mm())
        # Apply global v2 config defaults for find_tables/snap tuning.
        self._meta_panel.set_grid_tolerance_detected_mm(
            float(self._cfg.get("grid_tolerance_detected_mm", 0.5))
        )
        self._meta_panel.set_snap_max_distance_mm(
            float(self._cfg.get("snap_max_distance_mm", 5.0))
        )
        self._meta_panel.set_max_shape_change_ratio(
            float(self._cfg.get("max_shape_change_ratio", 2.5))
        )
        self._meta_panel.set_cascade_score_threshold(
            float(self._cfg.get("cascade_score_threshold", 0.4))
        )

        self._cells_panel = CellsPanel(self._gfx_view, self._catalog, self)
        self._props_panel = PropertiesPanel(self._catalog, self._gfx_view, self)
        # Keep field properties docked at the bottom; give extra height to the list.
        self._right_layout.addWidget(self._cells_panel, stretch=1)
        self._right_layout.addWidget(self._props_panel, stretch=0)

        from .wireframe_panel import GridLinesPanel, FieldBindingsPanel
        self._grid_lines_panel = GridLinesPanel(self)
        self._field_bindings_panel = FieldBindingsPanel(self)
        self._right_layout.insertWidget(1, self._grid_lines_panel, stretch=1)
        self._right_layout.insertWidget(2, self._field_bindings_panel, stretch=0)
        self._grid_lines_panel.hide()
        self._field_bindings_panel.hide()

        self._grid_lines_panel.interactive_add_line_requested.connect(
            self._on_interactive_add_grid_line,
        )
        self._shortcut_delete_grid_line = QShortcut(QKeySequence.StandardKey.Delete, self)
        self._shortcut_delete_grid_line.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self._shortcut_delete_grid_line.activated.connect(self._on_wireframe_line_delete_key)

        self._grid_lines_panel.lines_changed.connect(self._on_grid_lines_changed)
        self._grid_lines_panel.line_selected.connect(self._on_grid_line_selected)
        self._grid_lines_panel.hide_fields_toggled.connect(self._on_wireframe_hide_fields)
        self._grid_lines_panel.show_template_grid_toggled.connect(self._on_toggle_template_grid)
        self._grid_lines_panel.show_detected_grid_toggled.connect(self._on_toggle_detected_grid)
        self._grid_lines_panel._btn_generate.clicked.connect(self._on_generate_wireframe)
        self._grid_lines_panel._btn_autobind.clicked.connect(self._on_autobind_all)
        self._grid_lines_panel.pick_line_requested.connect(self._on_pick_line_requested)
        self._grid_lines_panel._btn_auto_refs.clicked.connect(self._on_auto_line_refs)
        self._field_bindings_panel.bindings_changed.connect(self._on_field_bindings_changed)
        self._field_bindings_panel._btn_autobind_one.clicked.connect(self._on_autobind_selected)

        self._cells_panel.cell_selected.connect(self._on_cell_selected_in_panel)
        self._cells_panel.cell_double_clicked.connect(self._on_cell_double_clicked)
        self._props_panel.field_changed.connect(self._on_field_changed)
        self._props_panel.document_field_changed.connect(self._on_document_property_field_changed)
        self._props_panel.edit_field_type_requested.connect(
            self._on_props_edit_field_type_requested,
        )

        from .cell_items import FieldRectItem

        FieldRectItem.set_id_dup_resolver(self._next_unique_field_id_for_duplicate)

        self._gfx_view.scene().selectionChanged.connect(self._on_scene_selection_changed)
        self._gfx_view.scene().changed.connect(self._on_scene_changed)

        self._apply_all_canvas_layers_from_panel()

    # ================================================================== status bar
    def _init_statusbar(self):
        self._lbl_coords = QLabel("—")
        self._lbl_zoom = QLabel("100%")
        bar = self.statusBar()
        bar.addWidget(self._lbl_coords, 1)
        bar.addPermanentWidget(self._lbl_zoom)

    def _on_mouse_moved(self, scene_pos: QPointF):
        mm = self._gfx_view.scene_to_mm_from_frame(scene_pos)
        if mm is not None:
            self._lbl_coords.setText(f"от рамки: {mm[0]:.1f} мм вправо, {mm[1]:.1f} мм вверх")
        else:
            pts = self._gfx_view.scene_to_pts(scene_pos)
            self._lbl_coords.setText(f"pts: ({pts.x():.1f}, {pts.y():.1f})")

    def _on_zoom_changed(self, zoom: float):
        self._lbl_zoom.setText(f"{zoom * 100:.0f}%")

    def _on_toggle_add_field_draw(self, enabled: bool) -> None:
        if enabled and self._gfx_view.current_fitz_page is None:
            QMessageBox.information(self, "Добавить поле", "Сначала откройте PDF.")
            self.act_add_field_draw.blockSignals(True)
            self.act_add_field_draw.setChecked(False)
            self.act_add_field_draw.blockSignals(False)
            return
        self._gfx_view.set_draw_mode(enabled)
        if enabled:
            self.statusBar().showMessage("Режим рисования: протяните прямоугольник на PDF", 4000)

    def _on_rect_drawn_add_field(self, rect: QRectF) -> None:
        from .cell_items import FieldRectItem

        self._clear_preview_overlays()
        existing = self._scene_field_items()
        next_idx = max((it.cell_index for it in existing), default=0) + 1
        item = FieldRectItem(rect.x(), rect.y(), rect.width(), rect.height(), cell_index=next_idx)
        self._gfx_view.scene().addItem(item)
        item.setSelected(True)

        # Auto-exit draw mode after creating one field.
        self.act_add_field_draw.blockSignals(True)
        self.act_add_field_draw.setChecked(False)
        self.act_add_field_draw.blockSignals(False)
        self._gfx_view.set_draw_mode(False)

        if self._cells_panel:
            self._rebuild_cells_list()
            self._cells_panel.highlight_items([item])
        if self._props_panel:
            self._props_panel.load_from_item(item)
            self._props_panel.setFocus()
        self._gfx_view.centerOn(item)
        self._set_modified(True)
        self._log(f"Добавлено новое поле: index={next_idx}, rect=({rect.x():.1f},{rect.y():.1f},{rect.width():.1f},{rect.height():.1f})")
        self.statusBar().showMessage("Поле добавлено. Можно назначить id в свойствах.", 4000)

    def _on_meta_changed(self) -> None:
        self._set_modified(True)
        if self._meta_panel is None:
            return
        v = self._meta_panel.get_grid_tolerance_template_mm()
        if abs(self._grid_tol_spin.value() - v) > 1e-9:
            self._grid_tol_spin.blockSignals(True)
            self._grid_tol_spin.setValue(v)
            self._grid_tol_spin.blockSignals(False)
        self._update_adapt_action_state()
        fp = self._gfx_view.current_fitz_page if self._gfx_view else None
        if fp is not None:
            self._recompute_and_apply_frame(fp)

    # ================================================================== menus
    def _init_menus(self):
        file_menu = self.menuBar().addMenu("Файл")
        file_menu.addAction(self.act_open_pdf)
        file_menu.addAction(self.act_load_template)
        file_menu.addAction(self.act_save_template)
        file_menu.addAction(self.act_save_template_as)
        file_menu.addAction(self.act_new)
        file_menu.addSeparator()
        file_menu.addAction("Выход", self.close)

        edit_menu = self.menuBar().addMenu("Правка")
        edit_menu.addAction(self.act_auto_detect)
        edit_menu.addAction(self.act_add_field_draw)
        edit_menu.addAction(self.act_clear_unassigned)
        edit_menu.addAction(self.act_delete_selected)
        edit_menu.addAction(self.act_test)
        edit_menu.addAction(self.act_consolidate_grid)
        edit_menu.addAction(self.act_apply_consolidation)
        edit_menu.addAction(self.act_adapt_to_grid)
        edit_menu.addAction(self.act_fill_stamp_gaps)
        edit_menu.addSeparator()
        edit_menu.addAction(self.act_adapt_debug_dump)

        view_menu = self.menuBar().addMenu("Вид")
        view_menu.addAction(self.act_wireframe_mode)
        view_menu.addAction(self.act_prealign_borders)

        tools_menu = self.menuBar().addMenu("Инструменты")
        tools_menu.addAction(self.act_catalog_editor)
        tools_menu.addAction(self.act_set_editor)

    # ================================================================== handlers

    # ---- PDF ----
    def _pdf_recent_settings(self) -> QSettings:
        return QSettings(_SETTINGS_ORG, _SETTINGS_APP)

    def _normalize_pdf_path(self, path: str) -> str:
        return os.path.normpath(os.path.abspath(os.path.expanduser(path)))

    def _apply_doc_type_from_pdf_path(self, path: str) -> None:
        """Set toolbar doc_type from PDF basename — same rule as v2 pipeline (``getDocTypeFromFile``)."""
        try:
            from utils.string_parsing import getDocTypeFromFile
        except ImportError:
            return
        base = os.path.basename(path)
        try:
            inferred = str(getDocTypeFromFile(base)).strip()
        except Exception:
            inferred = ""
        if not inferred:
            return
        self._doc_type_combo.blockSignals(True)
        self._doc_type_combo.setCurrentText(inferred)
        self._doc_type_combo.blockSignals(False)
        self._log(
            f"Тип документа из имени файла (как в pipeline): {inferred!r}"
        )

    def _load_recent_pdf_paths(self) -> list[str]:
        s = self._pdf_recent_settings()
        raw = s.value(_SETTINGS_RECENT_KEY, [])
        if isinstance(raw, str):
            return [raw] if raw else []
        if not isinstance(raw, list):
            return []
        out: list[str] = []
        for item in raw:
            if not item:
                continue
            p = str(item).strip()
            if p:
                out.append(self._normalize_pdf_path(p))
        return out[:_RECENT_PDF_MAX]

    def _save_recent_pdf_paths(self) -> None:
        s = self._pdf_recent_settings()
        s.setValue(_SETTINGS_RECENT_KEY, self._recent_pdf_paths[:_RECENT_PDF_MAX])

    def _remember_recent_pdf(self, path: str) -> None:
        norm = self._normalize_pdf_path(path)
        rest = [
            p
            for p in self._recent_pdf_paths
            if self._normalize_pdf_path(p) != norm
        ]
        self._recent_pdf_paths = [norm] + rest[: _RECENT_PDF_MAX - 1]
        self._save_recent_pdf_paths()

    def _remove_recent_pdf_path(self, path: str) -> None:
        norm = self._normalize_pdf_path(path)
        self._recent_pdf_paths = [
            p
            for p in self._recent_pdf_paths
            if self._normalize_pdf_path(p) != norm
        ]
        self._save_recent_pdf_paths()

    def _populate_recent_pdf_menu(self) -> None:
        self._menu_recent_pdfs.clear()
        filtered = [
            p for p in self._recent_pdf_paths if os.path.isfile(p)
        ][: _RECENT_PDF_MAX]
        if filtered != self._recent_pdf_paths:
            self._recent_pdf_paths = filtered
            self._save_recent_pdf_paths()

        if not self._recent_pdf_paths:
            ph = self._menu_recent_pdfs.addAction("Нет недавних файлов")
            ph.setEnabled(False)
            return

        for p in self._recent_pdf_paths[:_RECENT_PDF_MAX]:
            label = os.path.basename(p) or p
            act = self._menu_recent_pdfs.addAction(label)
            act.setData(p)
            act.setToolTip(p)

    def _on_recent_pdf_menu_triggered(self, action: QAction) -> None:
        path = action.data()
        if path:
            self._open_pdf(str(path))

    def _on_open_pdf(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Открыть PDF",
            os.path.dirname(self._pdf_path) if self._pdf_path else "",
            "PDF файлы (*.pdf);;Все файлы (*)"
        )
        if not path:
            return
        self._open_pdf(path)

    def _open_pdf(self, path: str):
        norm = self._normalize_pdf_path(path)
        if not os.path.isfile(norm):
            QMessageBox.warning(
                self,
                "Открыть PDF",
                f"Файл не найден:\n{norm}",
            )
            self._remove_recent_pdf_path(norm)
            self._populate_recent_pdf_menu()
            return

        import time
        t0 = time.perf_counter()
        if self._fitz_doc:
            self._fitz_doc.close()
        try:
            self._fitz_doc = fitz.open(norm)
        except Exception as e:
            QMessageBox.critical(
                self,
                "Открыть PDF",
                f"Не удалось открыть файл:\n{norm}\n\n{e!s}",
            )
            self._remove_recent_pdf_path(norm)
            self._populate_recent_pdf_menu()
            return

        self._pdf_path = norm
        self._remember_recent_pdf(norm)
        self._populate_recent_pdf_menu()
        self._apply_doc_type_from_pdf_path(norm)
        self._rebuild_page_combo(len(self._fitz_doc), select_1based=1)
        self._load_current_page()
        elapsed = time.perf_counter() - t0
        fitz_page = self._gfx_view.current_fitz_page
        rotation = fitz_page.rotation if fitz_page else "?"
        self._log(
            f"Открыт PDF: {os.path.basename(norm)}, "
            f"страниц={len(self._fitz_doc)}, rotation={rotation}, "
            f"время={elapsed:.2f}с"
        )

    def _load_current_page(self):
        if not self._fitz_doc:
            return
        self._reset_adaptation_session_state()
        page_idx = self._page_combo.currentIndex()
        if page_idx < 0 or page_idx >= len(self._fitz_doc):
            return
        fitz_page = self._fitz_doc[page_idx]
        self._recompute_and_apply_frame(fitz_page)
        self._update_adapt_action_state()
        self._refresh_matching_templates_ui(run_autoload=True)

    def _on_page_changed(self, index: int) -> None:
        if index < 0:
            return
        self._load_current_page()

    def _on_autoload_matching_template_toggled(self, checked: bool) -> None:
        s = self._pdf_recent_settings()
        s.setValue(_SETTINGS_AUTOLOAD_MATCHING_TEMPLATE_KEY, bool(checked))

    def _on_doc_type_for_matching_changed(self, _text: str) -> None:
        self._refresh_matching_templates_ui(run_autoload=False)

    def _confirm_discard_for_matching_template_load(self) -> bool:
        if not self._is_modified and not self._catalog_is_dirty:
            return True
        r = QMessageBox.question(
            self,
            "Несохранённые изменения",
            "Есть несохранённые изменения (шаблон и/или каталог).\n"
            "Загрузить подходящий шаблон без сохранения?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return r == QMessageBox.StandardButton.Yes

    def _first_matching_template_path(self) -> str:
        if not self._matching_templates:
            return ""
        p = (self._matching_templates[0].source_path or "").strip()
        return p if p and os.path.isfile(p) else ""

    def _refresh_matching_templates_ui(self, *, run_autoload: bool) -> None:
        """Fill toolbar combo via ``select_templates`` (priority desc). Optionally autoload first."""
        from pdf_parsing_v2_engine.stamp_extractor import select_templates
        from pdf_parsing_v2_engine.template_loader import load_all_templates

        self._matching_templates = []
        self._matching_templates_combo.blockSignals(True)
        self._matching_templates_combo.clear()

        if not self._fitz_doc or not (self._templates_dir or "").strip():
            self._matching_templates_combo.addItem("— нет PDF или папки шаблонов —")
            self._matching_templates_combo.setEnabled(False)
            self._btn_load_matching_template.setEnabled(False)
            self._matching_templates_combo.blockSignals(False)
            return

        page_num = self._current_pdf_page_1based()
        doc_type = (self._doc_type_combo.currentText() or "").strip()
        if not doc_type:
            self._matching_templates_combo.addItem("— укажите тип документа —")
            self._matching_templates_combo.setEnabled(False)
            self._btn_load_matching_template.setEnabled(False)
            self._matching_templates_combo.blockSignals(False)
            return

        try:
            all_t = load_all_templates(self._templates_dir)
        except Exception as e:
            self._log(f"Список шаблонов: {e}", "error")
            self._matching_templates_combo.addItem(f"— ошибка загрузки списка —")
            self._matching_templates_combo.setEnabled(False)
            self._btn_load_matching_template.setEnabled(False)
            self._matching_templates_combo.blockSignals(False)
            return

        matched = select_templates(all_t, doc_type, page_num)
        self._matching_templates = matched

        if not matched:
            self._matching_templates_combo.addItem("— нет подходящих шаблонов —")
            self._matching_templates_combo.setEnabled(True)
            self._btn_load_matching_template.setEnabled(False)
            self._matching_templates_combo.blockSignals(False)
            return

        for t in matched:
            path = (t.source_path or "").strip()
            base = os.path.basename(path) if path else "(нет пути)"
            label = f"{t.name}  [pr {t.priority}]  {t.page_selector}  ·  {base}"
            self._matching_templates_combo.addItem(label, path)

        self._matching_templates_combo.setCurrentIndex(0)
        self._matching_templates_combo.setEnabled(True)
        ok0 = bool(self._first_matching_template_path())
        self._btn_load_matching_template.setEnabled(ok0)
        self._matching_templates_combo.blockSignals(False)

        if run_autoload and self._chk_autoload_matching_template.isChecked():
            self._run_autoload_matching_template()

    def _run_autoload_matching_template(self) -> None:
        path = self._first_matching_template_path()
        if not path:
            return
        if not self._confirm_discard_for_matching_template_load():
            return
        self._load_template_from_file(path, clear_first=True)

    def _on_load_matching_template_clicked(self) -> None:
        path = self._first_matching_template_path()
        if not path:
            QMessageBox.warning(
                self,
                "Загрузить шаблон",
                "Нет подходящего шаблона с файлом на диске.",
            )
            return
        if not self._confirm_discard_for_matching_template_load():
            return
        self._load_template_from_file(path, clear_first=True)

    def _on_grid_tolerance_changed(self, value: float) -> None:
        if self._meta_panel is None:
            return
        # Keep toolbar and meta panel in sync (single source in template metadata).
        if abs(self._meta_panel.get_grid_tolerance_template_mm() - value) > 1e-9:
            self._meta_panel.set_grid_tolerance_template_mm(value)
        if self._consolidation_state is not None:
            self._on_consolidate_grid()

    def _update_adapt_action_state(self) -> None:
        enabled = bool(
            self._meta_panel is not None
            and self._meta_panel.get_grid_adapt()
            and self._gfx_view.current_fitz_page is not None
            and self._frame_info is not None
        )
        self.act_adapt_to_grid.setEnabled(enabled)

    def _scene_field_items(self) -> list:
        from .cell_items import FieldRectItem

        return [
            it for it in self._gfx_view.scene().items()
            if isinstance(it, FieldRectItem)
        ]

    def _rebuild_cells_list(self) -> None:
        """Refresh the field list panel including hidden PDF document-property rows."""
        if self._cells_panel is None:
            return
        doc = self._document_property_field_defs or None
        self._cells_panel.rebuild_list(doc)

    def _ordered_scene_field_items_and_defs(self) -> tuple[list, list[FieldDef]]:
        """FieldRectItems and FieldDefs in the same order (for catalog ↔ canvas sync)."""
        from .cell_items import FieldRectItem

        items = [
            it for it in self._scene_field_items()
            if isinstance(it, FieldRectItem) and it.field_def is not None
        ]
        fds = [it.field_def for it in items]
        return items, fds

    def _compute_stamp_search_bbox(self, template: StampTemplate, frame: FrameInfo) -> fitz.Rect | None:
        """Compute the exact stamp bbox used for find_tables filtering in adaptation."""
        from pdf_parsing_v2_engine.coord_transform import field_to_fitz_rect

        rects = [
            field_to_fitz_rect(f, frame, padding_mm=0.0)
            for f in template.fields
            if not getattr(f, "outside_stamp", False)
            and not getattr(f, "document_property", None)
        ]
        rects = [r for r in rects if not r.is_empty and not r.is_infinite]
        if not rects:
            return None
        x0 = min(r.x0 for r in rects)
        y0 = min(r.y0 for r in rects)
        x1 = max(r.x1 for r in rects)
        y1 = max(r.y1 for r in rects)
        return fitz.Rect(x0, y0, x1, y1)

    def _on_scene_changed(self, _rects) -> None:
        if self._cells_panel is None:
            return
        count = len(self._scene_field_items())
        if count != self._scene_field_count:
            self._scene_field_count = count
            self._rebuild_cells_list()
            self._schedule_catalog_coverage_update()

    def _clear_consolidation_overlay(self, reset_warnings: bool = True) -> None:
        if self._consolidation_overlay:
            scene = self._gfx_view.scene()
            for item in self._consolidation_overlay:
                if item.scene() is scene:
                    scene.removeItem(item)
            self._consolidation_overlay = []
        if self._consolidation_state is not None:
            self._consolidation_state = None
        self.act_apply_consolidation.setEnabled(False)
        if reset_warnings and self._cells_panel is not None:
            self._cells_panel.set_warning_items(set())

    def _clear_adapt_overlay(self) -> None:
        if not self._adapt_overlay:
            return
        scene = self._gfx_view.scene()
        for item in self._adapt_overlay:
            if item.scene() is scene:
                scene.removeItem(item)
        self._adapt_overlay = []

    def _clear_prealign_overlay(self) -> None:
        if not self._prealign_overlay_items:
            return
        scene = self._gfx_view.scene()
        for item in self._prealign_overlay_items:
            if item.scene() is scene:
                scene.removeItem(item)
        self._prealign_overlay_items.clear()

    def _rebuild_prealign_overlay_scene(self) -> None:
        """Redraw prealign bbox/line overlay from ``_last_prealign_debug_bundle`` and panel visibility."""
        self._clear_prealign_overlay()
        bundle = self._last_prealign_debug_bundle
        if not bundle or self._prealign_overlay_panel is None:
            return
        from .prealign_overlay_model import build_overlay_items_from_bundle

        dpi = self._gfx_view._dpi / 72.0
        scene = self._gfx_view.scene()
        z_shapes = 893
        z_labels = 897
        items = build_overlay_items_from_bundle(bundle)
        label_offset = {
            "frame": (4, 16),
            "base_bbox": (4, 30),
            "search_bbox": (4, 44),
            "proposed_bbox": (4, 58),
            "effective_bbox": (4, 72),
            "detected_union_bbox": (4, 86),
            "extent_probe_merged_grid": (4, 100),
            "extent_boundary_top": (8, -10),
            "extent_boundary_left": (8, 8),
            "extent_boundary_bottom": (8, -10),
            "extent_boundary_right": (8, 8),
            "virtual_top": (8, -14),
            "virtual_left": (8, 8),
            "alignment_matched_detected_boundary": (4, 114),
            "alignment_template_boundary": (4, 128),
        }
        for ov in items:
            if not self._prealign_overlay_panel.is_layer_visible(ov.layer_id):
                continue
            color = QColor(ov.color_rgb[0], ov.color_rgb[1], ov.color_rgb[2], 210)
            pen = QPen(color, max(0.8, ov.pen_width * 0.9), Qt.PenStyle.SolidLine)
            if ov.dashed:
                pen.setStyle(Qt.PenStyle.DashLine)
            ox, oy = label_offset.get(ov.layer_id, (6, 6))
            for seg_i, seg in enumerate(ov.segments):
                x0, y0, x1, y1 = seg[0] * dpi, seg[1] * dpi, seg[2] * dpi, seg[3] * dpi
                if ov.geometry == "rect":
                    r = QGraphicsRectItem(
                        QRectF(
                            min(x0, x1),
                            min(y0, y1),
                            max(abs(x1 - x0), 1.0),
                            max(abs(y1 - y0), 1.0),
                        )
                    )
                    r.setPen(pen)
                    r.setBrush(Qt.BrushStyle.NoBrush)
                    r.setZValue(z_shapes)
                    scene.addItem(r)
                    self._prealign_overlay_items.append(r)
                    if seg_i == 0:
                        lbl = QGraphicsSimpleTextItem(ov.label)
                        lbl.setBrush(QColor(40, 40, 40, 230))
                        lbl.setFont(QFont("Segoe UI", 8))
                        lbl.setPos(min(x0, x1) + ox, min(y0, y1) + oy)
                        lbl.setZValue(z_labels)
                        scene.addItem(lbl)
                        self._prealign_overlay_items.append(lbl)
                else:
                    ln = QGraphicsLineItem(x0, y0, x1, y1)
                    ln.setPen(pen)
                    ln.setZValue(z_shapes)
                    scene.addItem(ln)
                    self._prealign_overlay_items.append(ln)
                    if seg_i == 0:
                        lbl = QGraphicsSimpleTextItem(ov.label)
                        lbl.setBrush(QColor(40, 40, 40, 230))
                        lbl.setFont(QFont("Segoe UI", 7))
                        lbl.setPos(min(x0, x1) + ox, min(y0, y1) + oy)
                        lbl.setZValue(z_labels)
                        scene.addItem(lbl)
                        self._prealign_overlay_items.append(lbl)

    def _clear_detected_grid_overlay(self) -> None:
        if self._detected_grid_overlay:
            scene = self._gfx_view.scene()
            for item in self._detected_grid_overlay:
                if item.scene() is scene:
                    scene.removeItem(item)
            self._detected_grid_overlay = []
        if self._grid_lines_panel:
            self._grid_lines_panel.set_detected_lines([])
            self._grid_lines_panel._chk_show_detected_grid.setEnabled(False)

    def _on_toggle_detected_grid(self, visible: bool) -> None:
        for item in self._detected_grid_overlay:
            item.setVisible(visible)
        if self._prealign_overlay_panel:
            self._prealign_overlay_panel.sync_canvas_checkbox("pdf_grid", visible)

    def _on_toggle_template_grid(self, visible: bool) -> None:
        for item in self._wireframe_items:
            item.setVisible(visible)
        if self._prealign_overlay_panel:
            self._prealign_overlay_panel.sync_canvas_checkbox("template_grid", visible)

    def _clear_preview_overlays(self, reset_warnings: bool = True) -> None:
        self._clear_consolidation_overlay(reset_warnings=reset_warnings)
        self._clear_detected_grid_overlay()
        self._clear_adapt_overlay()
        self._clear_prealign_overlay()

    def _reset_adaptation_session_state(self, reset_warnings: bool = True) -> None:
        """Clear adapt/detected/consolidation previews, wireframe overlay, and cached snap state.

        Call after PDF/page change or template reload so stale graphics and _last_ca_info
        do not mix with the new context.
        """
        self._clear_preview_overlays(reset_warnings=reset_warnings)
        self._clear_wireframe_canvas()
        self._snapped_line_positions = {}
        self._adapt_transform = (1.0, 1.0, 0.0, 0.0)
        self._last_ca_info = None
        self._last_prealign_debug_bundle = None

    def _on_consolidate_grid(self) -> None:
        items = self._scene_field_items()
        if not items:
            QMessageBox.information(self, "Консолидация", "Нет ячеек для консолидации.")
            return

        self._clear_preview_overlays(reset_warnings=False)

        tolerance_mm = self._grid_tol_spin.value()
        tol_px = tolerance_mm * SCALE * (self._gfx_view._dpi / 72.0)
        x_coords: list[float] = []
        y_coords: list[float] = []
        min_x = float("inf")
        min_y = float("inf")
        max_x = float("-inf")
        max_y = float("-inf")

        for it in items:
            sx, sy, sw, sh = self._item_scene_rect(it)
            x_coords.extend([sx, sx + sw])
            y_coords.extend([sy, sy + sh])
            min_x = min(min_x, sx)
            min_y = min(min_y, sy)
            max_x = max(max_x, sx + sw)
            max_y = max(max_y, sy + sh)

        x_lines = cluster_lines(x_coords, tol_px)
        y_lines = cluster_lines(y_coords, tol_px)
        if not x_lines or not y_lines:
            QMessageBox.warning(self, "Консолидация", "Не удалось построить логические линии.")
            return

        stamp_scene_rect = QRectF(min_x, min_y, max_x - min_x, max_y - min_y)
        bad_item_ids: set[int] = set()

        def nearest_dist(val: float, lines: list[float]) -> float:
            return min(abs(val - x) for x in lines)

        for it in items:
            sx, sy, sw, sh = self._item_scene_rect(it)
            d = max(
                nearest_dist(sx, x_lines),
                nearest_dist(sx + sw, x_lines),
                nearest_dist(sy, y_lines),
                nearest_dist(sy + sh, y_lines),
            )
            if d > tol_px:
                bad_item_ids.add(id(it))

        scene = self._gfx_view.scene()
        x_pen = QPen(QColor(70, 130, 210, 170), 1.5, Qt.PenStyle.SolidLine)
        y_pen = QPen(QColor(230, 145, 40, 170), 1.5, Qt.PenStyle.SolidLine)
        bad_pen = QPen(QColor(220, 50, 50, 200), 2.0, Qt.PenStyle.DashLine)

        overlay: list[object] = []
        for x in x_lines:
            line = QGraphicsLineItem(x, stamp_scene_rect.top(), x, stamp_scene_rect.bottom())
            line.setPen(x_pen)
            line.setZValue(900)
            scene.addItem(line)
            overlay.append(line)
        for y in y_lines:
            line = QGraphicsLineItem(stamp_scene_rect.left(), y, stamp_scene_rect.right(), y)
            line.setPen(y_pen)
            line.setZValue(900)
            scene.addItem(line)
            overlay.append(line)
        for it in items:
            if id(it) not in bad_item_ids:
                continue
            sx, sy, sw, sh = self._item_scene_rect(it)
            rr = QGraphicsRectItem(QRectF(sx, sy, sw, sh))
            rr.setPen(bad_pen)
            rr.setBrush(Qt.BrushStyle.NoBrush)
            rr.setZValue(910)
            scene.addItem(rr)
            overlay.append(rr)

        self._consolidation_overlay = overlay
        self._apply_consolidation_overlay_visibility()
        self._consolidation_state = _ConsolidationState(
            x_lines=x_lines,
            y_lines=y_lines,
            tolerance_px=tol_px,
            bad_item_ids=bad_item_ids,
            stamp_scene_rect=stamp_scene_rect,
        )
        self.act_apply_consolidation.setEnabled(True)
        if self._cells_panel is not None:
            self._cells_panel.set_warning_items(bad_item_ids)
        self._log(
            f"Консолидация: {len(x_lines)} вертикальных линий, "
            f"{len(y_lines)} горизонтальных, "
            f"{len(bad_item_ids)} ячеек с отклонением > tolerance"
        )
        self.statusBar().showMessage(
            f"Консолидация: X={len(x_lines)}, Y={len(y_lines)}, проблемных={len(bad_item_ids)}",
            5000,
        )
        if self._grid_lines_panel and self._grid_lines_panel._chk_hide_fields.isChecked():
            self._on_wireframe_hide_fields(True)

    def _on_apply_consolidation(self) -> None:
        state = self._consolidation_state
        if state is None:
            self._on_consolidate_grid()
            state = self._consolidation_state
            if state is None:
                return

        items = self._scene_field_items()
        moved = 0
        for it in items:
            ix, iy, iw, ih = self._item_scene_rect(it)
            sx0 = min(state.x_lines, key=lambda x: abs(x - ix))
            sx1 = min(state.x_lines, key=lambda x: abs(x - (ix + iw)))
            sy0 = min(state.y_lines, key=lambda y: abs(y - iy))
            sy1 = min(state.y_lines, key=lambda y: abs(y - (iy + ih)))
            if sx1 - sx0 < 2.0:
                sx1 = sx0 + 2.0
            if sy1 - sy0 < 2.0:
                sy1 = sy0 + 2.0
            old = (ix, iy, ix + iw, iy + ih)
            new = (sx0, sy0, sx1, sy1)
            if any(abs(a - b) > 0.01 for a, b in zip(old, new)):
                it.setPos(QPointF(sx0, sy0))
                it.setRect(QRectF(0.0, 0.0, sx1 - sx0, sy1 - sy0))
                moved += 1

        self._set_modified(True)
        self._clear_preview_overlays()
        if self._cells_panel is not None:
            self._rebuild_cells_list()
        self._log(f"Применение консолидации: обновлено {moved} ячеек")
        self.statusBar().showMessage(f"Применено: обновлено {moved} ячеек", 4000)

    def _on_adapt_to_grid(self) -> None:
        if not self.act_adapt_to_grid.isEnabled():
            QMessageBox.information(
                self,
                "Тест-адаптация",
                "Нужен открытый PDF и включённый флаг 'Адаптировать под сетку PDF'.",
            )
            return
        fitz_page = self._gfx_view.current_fitz_page
        frame = self._frame_info
        if fitz_page is None or frame is None:
            return

        from .cell_items import FieldRectItem

        items = [
            it for it in self._gfx_view.scene().items()
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def is not None
        ]
        if not items:
            QMessageBox.information(self, "Тест-адаптация", "Нет назначенных ячеек.")
            return

        self._clear_preview_overlays()
        fields = self._collect_field_defs_from_scene(items)
        meta = self._meta_panel
        template = StampTemplate(
            schema_version=2,
            name="(adapt-preview)",
            doc_types=[self._doc_type_combo.currentText()],
            page_selector="all",
            priority=meta.get_priority() if meta else 10,
            padding_mm=meta.get_padding_mm() if meta else 0.5,
            grid_adapt=meta.get_grid_adapt() if meta else True,
            grid_tolerance_template_mm=meta.get_grid_tolerance_template_mm() if meta else 2.0,
            grid_tolerance_detected_mm=(meta.get_grid_tolerance_detected_mm() if meta else 0.5),
            snap_max_distance_mm=(meta.get_snap_max_distance_mm() if meta else 5.0),
            max_shape_change_ratio=(meta.get_max_shape_change_ratio() if meta else 2.5),
            cascade_score_threshold=(meta.get_cascade_score_threshold() if meta else 0.4),
            frame_mode=meta.get_frame_mode() if meta else "gost",
            fields=fields,
            grid_lines=list(self._grid_lines),
            **self._meta_find_tables_snap_kwargs(),
        )

        adapted_list, ca_info = adapt_by_cell_assignment(
            template=template, frame=frame, fitz_page=fitz_page, cfg=self._cfg,
        )
        self._last_ca_info = ca_info
        if not adapted_list:
            QMessageBox.warning(self, "Тест-адаптация", "Не удалось получить результат адаптации.")
            return

        keys = _make_unique_keys(fields)

        item_by_key: dict[str, object] = {}
        for key, it in zip(keys, items):
            item_by_key[key] = it

        dpi_scale = self._gfx_view._dpi / 72.0
        scene = self._gfx_view.scene()
        overlay: list[object] = []
        detected_pen = QPen(QColor(80, 150, 255, 140), 1.3, Qt.PenStyle.SolidLine)
        derived_pen = QPen(QColor(180, 140, 220, 200), 1.8, Qt.PenStyle.DashLine)
        miss_pen = QPen(QColor(220, 50, 50, 220), 2.0, Qt.PenStyle.DashLine)
        excluded_pen = QPen(QColor(199, 93, 0, 220), 2.0, Qt.PenStyle.DashLine)
        extended_pen = QPen(QColor(40, 200, 80, 220), 2.0, Qt.PenStyle.DashLine)
        grid_bound_pen = QPen(QColor(0, 190, 210, 220), 2.0, Qt.PenStyle.DashDotLine)

        statuses: dict[str, int] = {}
        max_snap_mm = 0.0
        seen_detected: set[tuple[int, int, int, int]] = set()
        moved = 0

        search_bbox = self._compute_stamp_search_bbox(template, frame)
        if search_bbox is not None:
            sb = QGraphicsRectItem(
                QRectF(
                    search_bbox.x0 * dpi_scale,
                    search_bbox.y0 * dpi_scale,
                    max(1.0, (search_bbox.x1 - search_bbox.x0) * dpi_scale),
                    max(1.0, (search_bbox.y1 - search_bbox.y0) * dpi_scale),
                )
            )
            sb.setPen(QPen(QColor(170, 70, 220, 220), 2.0, Qt.PenStyle.DashLine))
            sb.setBrush(Qt.BrushStyle.NoBrush)
            sb.setZValue(902)
            scene.addItem(sb)
            overlay.append(sb)

        # Detected grid lines (after border-pair merge) — magenta, selectable.
        from .grid_line_items import DetectedGridLineItem
        self._clear_detected_grid_overlay()

        # Build matched_to lookup (from alignment report diags in ca_info)
        matched_to_lookup: dict[str, str] = {}
        if ca_info.alignment_report and ca_info.alignment_report.log:
            pass  # matched_to info is in DetectedGridLineInfo objects below

        for dgl in ca_info.detected_grid_lines:
            det_item = DetectedGridLineItem(
                line_id=dgl.id,
                orientation=dgl.orientation,
                pos_pts=dgl.pos_pts,
                span_lo_pts=dgl.span_lo_pts,
                span_hi_pts=dgl.span_hi_pts,
                cell_count=dgl.cell_count,
                dpi_scale=dpi_scale,
            )
            scene.addItem(det_item)
            self._detected_grid_overlay.append(det_item)
        if self._detected_grid_overlay and self._grid_lines_panel:
            self._grid_lines_panel._chk_show_detected_grid.setEnabled(True)
            self._apply_detected_grid_overlay_visibility()
            self._grid_lines_panel.set_detected_lines(ca_info.detected_grid_lines)

        for field_key, ar in adapted_list:
            it = item_by_key.get(field_key)
            if it is None:
                continue
            statuses[ar.status] = statuses.get(ar.status, 0) + 1
            max_snap_mm = max(max_snap_mm, ar.snap_distance_mm)

            sx0 = ar.bbox.x0 * dpi_scale
            sy0 = ar.bbox.y0 * dpi_scale
            sw = max(2.0, (ar.bbox.x1 - ar.bbox.x0) * dpi_scale)
            sh = max(2.0, (ar.bbox.y1 - ar.bbox.y0) * dpi_scale)
            ox, oy, ow, oh = self._item_scene_rect(it)
            can_move = ar.status in ("matched", "anchor", "matched_extended", "grid_bound")
            if can_move:
                if (
                    abs(ox - sx0) > 0.01
                    or abs(oy - sy0) > 0.01
                    or abs(ow - sw) > 0.01
                    or abs(oh - sh) > 0.01
                ):
                    it.setPos(QPointF(sx0, sy0))
                    it.setRect(QRectF(0.0, 0.0, sw, sh))
                    moved += 1

            if ar.status in ("derived", "no_match", "excluded", "matched_extended", "grid_bound"):
                rr = QGraphicsRectItem(QRectF(sx0, sy0, sw, sh))
                if ar.status == "derived":
                    rr.setPen(derived_pen)
                elif ar.status == "excluded":
                    rr.setPen(excluded_pen)
                elif ar.status == "matched_extended":
                    rr.setPen(extended_pen)
                elif ar.status == "grid_bound":
                    rr.setPen(grid_bound_pen)
                else:
                    rr.setPen(miss_pen)
                rr.setBrush(Qt.BrushStyle.NoBrush)
                rr.setZValue(908)
                scene.addItem(rr)
                overlay.append(rr)

            # Score/iter label (only at sufficient zoom).
            if self._gfx_view.transform().m11() > 0.3:
                from PySide6.QtWidgets import QGraphicsSimpleTextItem
                from PySide6.QtGui import QFont
                lbl_text = ""
                if ar.status == "anchor":
                    lbl_text = "A"
                elif ar.status == "derived":
                    lbl_text = "D"
                elif ar.status == "matched_extended":
                    lbl_text = "E"
                elif ar.status == "grid_bound":
                    lbl_text = "G"
                elif ar.confidence > 0:
                    lbl_text = f"{ar.confidence:.2f}"
                if lbl_text:
                    lbl = QGraphicsSimpleTextItem(lbl_text)
                    font = QFont("Consolas", 7)
                    lbl.setFont(font)
                    lbl.setBrush(QColor(40, 40, 40, 200))
                    lbl.setPos(sx0 + 1, sy0 + 1)
                    lbl.setZValue(910)
                    scene.addItem(lbl)
                    overlay.append(lbl)

            for dc in ar.matched_detected:
                key = (round(dc.x0), round(dc.y0), round(dc.x1), round(dc.y1))
                if key in seen_detected:
                    continue
                seen_detected.add(key)
                dr = QGraphicsRectItem(
                    QRectF(
                        dc.x0 * dpi_scale,
                        dc.y0 * dpi_scale,
                        max(1.0, (dc.x1 - dc.x0) * dpi_scale),
                        max(1.0, (dc.y1 - dc.y0) * dpi_scale),
                    )
                )
                dr.setPen(detected_pen)
                dr.setBrush(Qt.BrushStyle.NoBrush)
                dr.setZValue(904)
                scene.addItem(dr)
                overlay.append(dr)

        self._adapt_overlay = overlay
        self._apply_adapt_overlay_visibility()
        try:
            from .adapt_debug import stamp_prealign_debug_bundle

            self._last_prealign_debug_bundle = stamp_prealign_debug_bundle(
                frame=frame,
                template=template,
                ca_info=ca_info,
                fitz_page=fitz_page,
                cfg=self._cfg,
            )
        except Exception:
            self._last_prealign_debug_bundle = None
        self._rebuild_prealign_overlay_scene()
        if moved > 0:
            self._set_modified(True)
        if self._cells_panel is not None:
            self._rebuild_cells_list()
            results_by_item_id: dict[int, object] = {}
            for key, ar in adapted_list:
                it = item_by_key.get(key)
                if it is not None:
                    results_by_item_id[id(it)] = ar
            self._cells_panel.update_adapt_results(results_by_item_id)
        warn_str = "; ".join(ca_info.warnings) if ca_info.warnings else "none"
        n_extended = statuses.get('matched_extended', 0)
        n_grid_bound = statuses.get('grid_bound', 0)
        self._log(
            f"Cell-assignment: "
            f"matched={statuses.get('matched', 0)}, "
            f"extended={n_extended}, "
            f"grid_bound={n_grid_bound}, "
            f"derived={statuses.get('derived', 0)}, "
            f"excluded={statuses.get('excluded', 0)}, "
            f"max_snap_mm={max_snap_mm:.2f}, moved={moved}, "
            f"detected={ca_info.n_detected}, rect={ca_info.rectangularity:.0%}, "
            f"scale=({ca_info.transform_scale_x:.4f},{ca_info.transform_scale_y:.4f}), "
            f"warnings=[{warn_str}]"
        )
        ext_msg = f", extended={n_extended}" if n_extended else ""
        grid_msg = f", grid_bound={n_grid_bound}" if n_grid_bound else ""
        self.statusBar().showMessage(
            f"Подогнано (cell-assign): "
            f"matched={statuses.get('matched', 0)}/{ca_info.n_fields}{ext_msg}{grid_msg}, "
            f"derived={statuses.get('derived', 0)}, "
            f"rect={ca_info.rectangularity:.0%}",
            6000,
        )
        # Store snapped positions + transform and show wireframe at adapted positions
        if ca_info.snapped_line_positions:
            self._snapped_line_positions = dict(ca_info.snapped_line_positions)
            self._adapt_transform = (
                ca_info.transform_scale_x,
                ca_info.transform_scale_y,
                ca_info.transform_dx,
                ca_info.transform_dy,
            )
            self._draw_wireframe_on_canvas()

        # Log alignment report summary
        if ca_info.alignment_report:
            ar = ca_info.alignment_report
            self._log(
                f"Alignment: {ar.verdict.upper()} — "
                f"matched={ar.matched_count}/{ar.total_tpl_lines}, "
                f"no_match={ar.no_match_count}, "
                f"skipped_pdf={ar.skipped_detected_count}, "
                f"avg_score={ar.avg_match_score:.2f}"
            )
            if ar.problem_lines:
                for p in ar.problem_lines[:10]:
                    self._log(f"  {p}")

        if self._grid_lines_panel and self._grid_lines_panel._chk_hide_fields.isChecked():
            self._on_wireframe_hide_fields(True)

    def _stamp_scene_rect_for_fill(self, items: list) -> QRectF | None:
        """Габарит штампа в координатах сцены (union полей шаблона, не outside_stamp)."""
        frame = self._frame_info
        if frame is None or not items:
            return None
        fields = self._collect_field_defs_from_scene(items)
        meta = self._meta_panel
        template = StampTemplate(
            schema_version=2,
            name="(fill-gaps)",
            doc_types=[self._doc_type_combo.currentText()],
            page_selector="all",
            priority=meta.get_priority() if meta else 10,
            padding_mm=meta.get_padding_mm() if meta else 0.5,
            grid_adapt=meta.get_grid_adapt() if meta else True,
            grid_tolerance_template_mm=meta.get_grid_tolerance_template_mm() if meta else 2.0,
            grid_tolerance_detected_mm=(meta.get_grid_tolerance_detected_mm() if meta else 0.5),
            snap_max_distance_mm=(meta.get_snap_max_distance_mm() if meta else 5.0),
            max_shape_change_ratio=(meta.get_max_shape_change_ratio() if meta else 2.5),
            cascade_score_threshold=(meta.get_cascade_score_threshold() if meta else 0.4),
            frame_mode=meta.get_frame_mode() if meta else "gost",
            fields=fields,
            grid_lines=list(self._grid_lines),
            **self._meta_find_tables_snap_kwargs(),
        )
        fb = self._compute_stamp_search_bbox(template, frame)
        if fb is None:
            return None
        dpi_scale = self._gfx_view._dpi / 72.0
        return QRectF(
            fb.x0 * dpi_scale,
            fb.y0 * dpi_scale,
            (fb.x1 - fb.x0) * dpi_scale,
            (fb.y1 - fb.y0) * dpi_scale,
        )

    # ---- fill gaps to neighbors (editor-only post-pass) ----
    def _on_fill_stamp_gaps(self) -> None:
        """Expand assigned field rects to meet left/right/top/bottom neighbors inside stamp."""
        from pdf_parsing_v2_engine.stamp_fill_gaps import fill_stamp_gaps_for_rects
        from .cell_items import FieldRectItem

        items = [
            it for it in self._gfx_view.scene().items()
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def is not None
            and not getattr(it.field_def, "outside_stamp", False)
        ]
        if len(items) < 2:
            QMessageBox.information(
                self, "Заполнить зазоры",
                "Нужно минимум 2 назначенных поля внутри штампа.",
            )
            return

        stamp_r = self._stamp_scene_rect_for_fill(items)
        if stamp_r is None or stamp_r.width() < 4 or stamp_r.height() < 4:
            QMessageBox.warning(
                self, "Заполнить зазоры",
                "Не удалось вычислить габарит штампа.",
            )
            return

        from pdf_parsing_v2_engine.coord_transform import SCALE as _SCALE

        tol = self._grid_tol_spin.value() if hasattr(self, "_grid_tol_spin") else 2.0
        tol_scene = tol * _SCALE * (self._gfx_view._dpi / 72.0)

        stamp_t = (
            stamp_r.left(), stamp_r.top(),
            stamp_r.right(), stamp_r.bottom(),
        )
        pairs: list[tuple[int, tuple[float, float, float, float]]] = []
        for it in items:
            sx, sy, sw, sh = self._item_scene_rect(it)
            pairs.append((
                id(it),
                (sx, sy, sx + sw, sy + sh),
            ))

        changed = fill_stamp_gaps_for_rects(
            pairs,
            stamp_t,
            overlap_frac=0.22,
            eps_scene=max(1.0, tol_scene),
            expand_horizontal=True,
            expand_vertical=True,
        )

        moved = 0
        for it in items:
            new_t = changed.get(id(it))
            if new_t is None:
                continue
            nl, nt, nr, nb = new_t
            w = max(2.0, nr - nl)
            h = max(2.0, nb - nt)
            it.setPos(QPointF(nl, nt))
            it.setRect(QRectF(0.0, 0.0, w, h))
            moved += 1

        if moved > 0:
            self._set_modified(True)
        if self._cells_panel is not None:
            self._rebuild_cells_list()
        self._log(
            f"Заполнить зазоры: расширено {moved} из {len(items)} полей "
            f"(габарит штампа {stamp_r.width():.0f}×{stamp_r.height():.0f} px сцены)"
        )
        self.statusBar().showMessage(
            f"Зазоры: обновлено полей: {moved}/{len(items)}", 5000,
        )

    # ---- debug dump ----
    def _on_adapt_debug_dump(self) -> None:
        """Collect and save a comprehensive debug snapshot for adaptation analysis."""
        fitz_page = self._gfx_view.current_fitz_page
        frame = self._frame_info
        if fitz_page is None or frame is None:
            QMessageBox.information(
                self, "Дамп отладки", "Сначала откройте PDF."
            )
            return

        from .cell_items import FieldRectItem
        from .adapt_debug import (
            collect_adapt_debug_snapshot,
            save_adapt_debug_snapshot,
        )

        items = [
            it for it in self._gfx_view.scene().items()
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def is not None
        ]
        if not items:
            QMessageBox.information(self, "Дамп отладки", "Нет назначенных ячеек.")
            return

        fields = self._collect_field_defs_from_scene(items)
        meta = self._meta_panel
        page_num = self._current_pdf_page_1based()
        template_source_path = os.path.abspath(self._active_template_path) if self._active_template_path else ""
        template_name = (
            os.path.splitext(os.path.basename(template_source_path))[0]
            if template_source_path
            else "(debug-dump)"
        )
        dump_cfg = dict(self._cfg or {})
        if self._pdf_path:
            dump_cfg["_pdf_path"] = self._pdf_path
        dump_cfg["_page_num"] = page_num
        template = StampTemplate(
            schema_version=2,
            name=template_name,
            doc_types=[self._doc_type_combo.currentText()],
            page_selector="all",
            priority=meta.get_priority() if meta else 10,
            padding_mm=meta.get_padding_mm() if meta else 0.5,
            grid_adapt=meta.get_grid_adapt() if meta else True,
            grid_tolerance_template_mm=meta.get_grid_tolerance_template_mm() if meta else 2.0,
            grid_tolerance_detected_mm=(meta.get_grid_tolerance_detected_mm() if meta else 0.5),
            snap_max_distance_mm=(meta.get_snap_max_distance_mm() if meta else 5.0),
            max_shape_change_ratio=(meta.get_max_shape_change_ratio() if meta else 2.5),
            cascade_score_threshold=(meta.get_cascade_score_threshold() if meta else 0.4),
            frame_mode=meta.get_frame_mode() if meta else "gost",
            fields=fields,
            grid_lines=list(self._grid_lines),
            source_path=template_source_path,
            **self._meta_find_tables_snap_kwargs(),
        )

        dpi_scale = self._gfx_view._dpi / 72.0
        try:
            snapshot = collect_adapt_debug_snapshot(
                fitz_page=fitz_page,
                frame=frame,
                template=template,
                scene_items=items,
                fields_from_scene=fields,
                pdf_path=self._pdf_path,
                page_num=page_num,
                dpi_scale=dpi_scale,
                cfg=dump_cfg,
            )
        except Exception as exc:
            tb = traceback.format_exc()
            self._log(f"Дамп отладки: ошибка сбора данных:\n{tb}", "error")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при сборе данных:\n{exc}")
            return

        output_dir = os.path.join(
            os.path.dirname(self._pdf_path) if self._pdf_path else ".",
            "adapt_debug",
        )
        alignment_log = ""
        ca_info_snap = getattr(self, '_last_ca_info', None)
        if ca_info_snap and ca_info_snap.alignment_report:
            alignment_log = ca_info_snap.alignment_report.log

        try:
            path = save_adapt_debug_snapshot(
                snapshot, output_dir=output_dir, alignment_log=alignment_log,
            )
        except Exception as exc:
            self._log(f"Дамп отладки: ошибка записи файла: {exc}", "error")
            QMessageBox.critical(self, "Ошибка", f"Не удалось записать файл:\n{exc}")
            return

        self._log(f"Дамп отладки сохранён: {path}")
        self.statusBar().showMessage(f"Debug dump: {os.path.basename(path)}", 8000)
        try:
            os.startfile(os.path.dirname(path))
        except Exception:
            pass

    # ---- auto detect ----
    def _on_auto_detect(self):
        self._clear_preview_overlays()
        fitz_page = self._gfx_view.current_fitz_page
        if fitz_page is None:
            QMessageBox.warning(self, "Авто-разметка", "Сначала откройте PDF.")
            return
        from .auto_detect import detect_cells
        from .cell_items import FieldRectItem

        snap_tmpl = StampTemplate(
            schema_version=2,
            name="(auto-detect)",
            doc_types=[self._doc_type_combo.currentText()],
            page_selector="all",
            **self._meta_find_tables_snap_kwargs(),
        )
        bboxes = detect_cells(fitz_page, cfg=self._cfg, template=snap_tmpl)
        dpi_scale = self._gfx_view._dpi / 72.0
        for idx, (x0, y0, x1, y1) in enumerate(bboxes):
            sx0, sy0, sx1, sy1 = x0 * dpi_scale, y0 * dpi_scale, x1 * dpi_scale, y1 * dpi_scale
            item = FieldRectItem(sx0, sy0, sx1 - sx0, sy1 - sy0, cell_index=idx + 1)
            self._gfx_view.scene().addItem(item)

        if self._cells_panel:
            self._rebuild_cells_list()
        self._set_modified(True)
        msg = f"Авто-разметка: {len(bboxes)} ячеек"
        self.statusBar().showMessage(msg, 3000)
        self._log(msg)

    # ---- test / preview (F5) ----
    def _on_test(self):
        fitz_page = self._gfx_view.current_fitz_page
        if fitz_page is None:
            QMessageBox.warning(self, "Тест", "Сначала откройте PDF.")
            return
        from .cell_items import FieldRectItem
        from pdf_parsing_v2_engine.stamp_extractor import extract_page

        items = [
            it for it in self._gfx_view.scene().items()
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def is not None
        ]
        if not items:
            QMessageBox.information(self, "Тест", "Нет назначенных ячеек.")
            self._log("Тест F5: нет назначенных ячеек", "warn")
            return

        self._clear_adapt_overlay()
        self._clear_prealign_overlay()

        doc_type = self._doc_type_combo.currentText()
        page_num = self._current_pdf_page_1based()
        self._log(
            f"Тест F5: doc_type={doc_type!r}, страница={page_num}, "
            f"назначенных ячеек={len(items)}"
        )

        # bbox из текущей геометрии сцены (не устаревший field_def с диска)
        fields: list[FieldDef] = self._collect_field_defs_from_scene(items)
        from pdf_parsing_v2_engine.coord_transform import snap_outside_stamp_vertical_to_frame

        fi = self._frame_info
        if fi is not None:
            for i, fd in enumerate(fields):
                if fd.outside_stamp:
                    new_mm = snap_outside_stamp_vertical_to_frame(fd, fi)
                    if new_mm is not None:
                        fields[i] = replace(fd, bbox_mm=new_mm)
        for it, fd in zip(items, fields):
            it.assign_field(fd)

        from pdf_parsing_v2_engine.document_properties import ensure_document_property_field_defs

        merge_ids = {f.id for f in fields}
        extra_doc = [f for f in self._document_property_field_defs if f.id not in merge_ids]
        combined_fields = fields + extra_doc
        meta = self._meta_panel
        template = StampTemplate(
            schema_version=2,
            name="(тестовый)",
            doc_types=[doc_type],
            page_selector="all",
            priority=meta.get_priority() if meta else 10,
            padding_mm=meta.get_padding_mm() if meta else 0.5,
            grid_adapt=meta.get_grid_adapt() if meta else False,
            grid_tolerance_template_mm=meta.get_grid_tolerance_template_mm() if meta else 2.0,
            grid_tolerance_detected_mm=(meta.get_grid_tolerance_detected_mm() if meta else 0.5),
            snap_max_distance_mm=(meta.get_snap_max_distance_mm() if meta else 5.0),
            max_shape_change_ratio=(meta.get_max_shape_change_ratio() if meta else 2.5),
            cascade_score_threshold=(meta.get_cascade_score_threshold() if meta else 0.4),
            frame_mode=meta.get_frame_mode() if meta else "gost",
            fields=combined_fields,
            grid_lines=list(self._grid_lines),
            **self._meta_find_tables_snap_kwargs(),
        )
        template = ensure_document_property_field_defs(template)
        self._log(
            "Тест F5: bbox полей с холста + grid_lines в памяти (как при Тест-адаптация)"
        )

        test_cfg = {
            **(self._cfg or {}),
            "_pdf_path": self._pdf_path,
            "_page_num": page_num,
            "source_pdf_basename": os.path.basename(self._pdf_path) if self._pdf_path else "",
            "editor_prealign_overlay": True,
        }
        try:
            result = extract_page(
                fitz_page, doc_type, page_num, [template], cfg=test_cfg
            )
        except Exception as exc:
            tb = traceback.format_exc()
            self._log(f"Тест F5: исключение из extract_page:\n{tb}", "error")
            QMessageBox.critical(self, "Ошибка теста", f"extract_page завершился с ошибкой:\n{exc}")
            return

        dbg = (result.metadata or {}).get("stamp_prealign_debug")
        self._last_prealign_debug_bundle = dbg if isinstance(dbg, dict) else None
        self._rebuild_prealign_overlay_scene()

        fields_with_data = 0
        fields_empty = 0
        for it in items:
            fd = it.field_def
            fr = result.fields.get(fd.id)
            if fr is None:
                it.set_test_status("error")
                fields_empty += 1
                continue
            has_val = bool(fr.cleaned_value and fr.cleaned_value.strip())
            if has_val:
                fields_with_data += 1
            else:
                fields_empty += 1
            if fr.raw_value == "" or fr.raw_value is None:
                self._log(
                    f"  поле {fd.id!r}: пустой raw_value",
                    "warn",
                )
            if fd.expected == "required" and not has_val:
                it.set_test_status("error")
            elif fd.expected == "absent" and has_val:
                it.set_test_status("error")
            elif not has_val:
                it.set_test_status("warning")
            elif fr.is_valid is False:
                it.set_test_status("warning")
            else:
                it.set_test_status("ok")

        if self._cells_panel:
            self._cells_panel.update_test_values(result)

        # Generate debug PNG with text span overlay
        png_path = self._generate_test_debug_png(result, items, page_num)

        msg = (
            f"Тест завершён: score={result.template_score:.2f}, "
            f"полей с данными: {fields_with_data}/{len(items)}"
        )
        if png_path:
            msg += " | PNG сохранён"
        self.statusBar().showMessage(msg, 8000)
        self._log(
            f"Тест F5: score={result.template_score:.2f}, "
            f"с данными={fields_with_data}, пустых={fields_empty}, "
            f"предупреждений={len(result.warnings)}"
        )

    # ---- clear / delete ----
    def _on_clear_unassigned(self):
        self._clear_preview_overlays()
        from .cell_items import FieldRectItem
        to_remove = [
            it for it in self._gfx_view.scene().items()
            if isinstance(it, FieldRectItem) and not it.is_assigned
        ]
        for it in to_remove:
            self._gfx_view.scene().removeItem(it)
        if self._cells_panel:
            self._rebuild_cells_list()
        if to_remove:
            self._set_modified(True)
        self.statusBar().showMessage(f"Удалено {len(to_remove)} неназначенных ячеек", 3000)

    def _on_delete_selected(self):
        self._clear_preview_overlays()
        from .cell_items import FieldRectItem
        selected = [
            it for it in self._gfx_view.scene().selectedItems()
            if isinstance(it, FieldRectItem)
        ]
        if len(selected) > 10:
            r = QMessageBox.question(
                self, "Удаление",
                f"Удалить {len(selected)} ячеек?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                return
        for it in selected:
            self._gfx_view.scene().removeItem(it)
        if self._cells_panel:
            self._rebuild_cells_list()
        if selected:
            self._set_modified(True)

    # ---- save / load ----
    def _on_load_template(self):
        if self._is_modified or self._catalog_is_dirty:
            r = QMessageBox.question(
                self,
                "Несохранённые изменения",
                "Есть несохранённые изменения (шаблон и/или каталог).\n"
                "Загрузить другой шаблон без сохранения?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                return
        path, _ = QFileDialog.getOpenFileName(
            self, "Загрузить шаблон",
            self._templates_dir or "",
            "JSON шаблоны (*.json)"
        )
        if not path:
            return
        self._load_template_from_file(path, clear_first=True)

    def _load_template_from_file(self, path: str, clear_first: bool = False):
        from pdf_parsing_v2_engine.models import StampTemplate
        from .cell_items import FieldRectItem

        try:
            tmpl = StampTemplate.from_json(path)
            from pdf_parsing_v2_engine.models import FieldCatalog, merge_stamp_template_with_catalog
            from pdf_parsing_v2_engine.document_properties import ensure_document_property_field_defs

            cat_path = os.path.join(os.path.dirname(os.path.abspath(path)), "catalog.json")
            if os.path.isfile(cat_path):
                tmpl = merge_stamp_template_with_catalog(
                    tmpl, FieldCatalog.from_json(cat_path),
                )
            elif self._catalog:
                tmpl = merge_stamp_template_with_catalog(tmpl, self._catalog)
            tmpl = ensure_document_property_field_defs(tmpl)
            self._document_property_field_defs = [f for f in tmpl.fields if f.document_property]
        except Exception as e:
            tb = traceback.format_exc()
            self._log(f"Ошибка загрузки шаблона {path!r}:\n{tb}", "error")
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить шаблон:\n{e}")
            return

        fitz_page = self._gfx_view.current_fitz_page
        if fitz_page is None:
            QMessageBox.warning(self, "Загрузка", "Сначала откройте PDF для корректного размещения полей.")
            return

        if tmpl.frame_mode == "drawing_union":
            frame, _ = find_frame(fitz_page, frame_mode="drawing_union", template=tmpl)
        else:
            frame, _ = find_frame(fitz_page)
        self._frame_info = frame
        self._gfx_view.load_page(fitz_page, frame)
        self._draw_origin_markers(frame)

        from pdf_parsing_v2_engine.coord_transform import snap_outside_stamp_vertical_to_frame

        for i, fd in enumerate(tmpl.fields):
            if fd.outside_stamp:
                new_mm = snap_outside_stamp_vertical_to_frame(fd, frame)
                if new_mm is not None:
                    tmpl.fields[i] = replace(fd, bbox_mm=new_mm)
                    self._log(f"Auto-snap outside_stamp: {fd.id}")

        self._reset_adaptation_session_state()

        if clear_first:
            for it in list(self._gfx_view.scene().items()):
                if isinstance(it, FieldRectItem):
                    self._gfx_view.scene().removeItem(it)

        from pdf_parsing_v2_engine.coord_transform import field_to_fitz_rect
        dpi_scale = self._gfx_view._dpi / 72.0

        vis_idx = 0
        for i, fd in enumerate(tmpl.fields):
            if fd.document_property:
                continue
            vis_idx += 1
            fitz_rect = field_to_fitz_rect(fd, frame, padding_mm=0)
            sx0 = fitz_rect.x0 * dpi_scale
            sy0 = fitz_rect.y0 * dpi_scale
            sw = fitz_rect.width * dpi_scale
            sh = fitz_rect.height * dpi_scale
            item = FieldRectItem(sx0, sy0, sw, sh, cell_index=vis_idx)
            item.assign_field(fd)
            self._gfx_view.scene().addItem(item)
            self._snap_item_to_field_def(item, fd)

        if self._cells_panel:
            self._rebuild_cells_list()

        dup_fixed = False
        id_counts = Counter(fd.id for fd in tmpl.fields)
        if any(c > 1 for c in id_counts.values()):
            dup = [i for i, c in id_counts.items() if c > 1]
            r = QMessageBox.question(
                self,
                "Шаблон",
                f"В файле повторяются id полей: {', '.join(dup[:15])}"
                f"{'…' if len(dup) > 15 else ''}\n"
                "Исправить дубликаты автоматически?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if r == QMessageBox.StandardButton.Yes:
                self._auto_fix_duplicate_field_ids_on_scene()
                dup_fixed = True

        if self._meta_panel is not None:
            self._meta_panel.load_from_template(tmpl)
            self._grid_tol_spin.setValue(self._meta_panel.get_grid_tolerance_template_mm())
        self._update_adapt_action_state()

        self._grid_lines = list(tmpl.grid_lines)

        # Check for doubled (border-pair) grid lines
        self._check_and_offer_collapse_doubled_lines()

        if hasattr(self, '_grid_lines_panel') and self._grid_lines_panel:
            self._grid_lines_panel.set_lines(self._grid_lines)
            self._field_bindings_panel.set_lines(self._grid_lines)
            self._refresh_wireframe_canvas()

        self._apply_all_canvas_layers_from_panel()
        self._set_active_template(path)
        self._set_modified(dup_fixed)
        self._log(f"Загружен шаблон: {tmpl.name!r} из {os.path.basename(path)!r}, полей={len(tmpl.fields)}, grid_lines={len(self._grid_lines)}")
        self.statusBar().showMessage(f"Загружен шаблон: {tmpl.name} ({len(tmpl.fields)} полей)", 3000)
        # Update catalog panel coverage (T7.3)
        self._update_catalog_panel_coverage()

    def _on_save_template(self):
        """Ctrl+S — save to current file; if no current file, behave like Save As."""
        if self._active_template_path:
            self._do_save(self._active_template_path)
        else:
            self._on_save_template_as()

    def _on_save_template_as(self):
        """Ctrl+Shift+S — always ask for a file path."""
        default_dir = (
            os.path.dirname(self._active_template_path)
            if self._active_template_path
            else self._templates_dir or ""
        )
        path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить шаблон как", default_dir, "JSON шаблоны (*.json)"
        )
        if not path:
            return
        self._do_save(path)

    def _do_save(self, path: str) -> None:
        """Core save logic: collect assigned cells, build StampTemplate, write JSON."""
        from .cell_items import FieldRectItem

        items = [
            it for it in self._gfx_view.scene().items()
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def is not None
        ]
        unassigned = sum(
            1 for it in self._gfx_view.scene().items()
            if isinstance(it, FieldRectItem) and not it.is_assigned
        )
        if not items:
            QMessageBox.information(self, "Сохранение", "Нет назначенных ячеек для сохранения.")
            return
        if unassigned > 0:
            self._log(f"Сохранение: {unassigned} неназначенных ячеек пропущено", "warn")

        fields = self._collect_field_defs_from_scene(items)
        fields_full = fields + list(self._document_property_field_defs)

        if not self._validate_template_ids_before_save(items, fields_full):
            return

        fields_full = self._collect_field_defs_from_scene(items) + list(
            self._document_property_field_defs
        )

        # Read metadata from the persistent panel (fallback to defaults if panel not ready)
        meta = self._meta_panel
        tmpl = StampTemplate(
            schema_version=2,
            name=meta.get_name() if meta else "template",
            doc_types=meta.get_doc_types() if meta else [],
            page_selector=meta.get_page_selector() if meta else "first",
            priority=meta.get_priority() if meta else 10,
            padding_mm=meta.get_padding_mm() if meta else 0.5,
            grid_adapt=meta.get_grid_adapt() if meta else False,
            grid_tolerance_template_mm=meta.get_grid_tolerance_template_mm() if meta else 2.0,
            grid_tolerance_detected_mm=(meta.get_grid_tolerance_detected_mm() if meta else 0.5),
            snap_max_distance_mm=(meta.get_snap_max_distance_mm() if meta else 5.0),
            max_shape_change_ratio=(meta.get_max_shape_change_ratio() if meta else 2.5),
            cascade_score_threshold=(meta.get_cascade_score_threshold() if meta else 0.4),
            frame_mode=meta.get_frame_mode() if meta else "gost",
            fields=fields_full,
            grid_lines=list(self._grid_lines),
            **self._meta_find_tables_snap_kwargs(),
        )
        try:
            catalog_for_save = None
            cat_path = os.path.join(os.path.dirname(os.path.abspath(path)), "catalog.json")
            if os.path.isfile(cat_path):
                try:
                    catalog_for_save = FieldCatalog.from_json(cat_path)
                except (OSError, ValueError):
                    catalog_for_save = None
            tmpl.to_json(path, catalog=catalog_for_save)
            self._set_active_template(path)
            self._set_modified(False)
            if hasattr(self, "_templates_browser") and self._templates_browser:
                self._templates_browser.refresh()
            self.statusBar().showMessage(f"Шаблон сохранён: {os.path.basename(path)}", 5000)
            self._log(f"Сохранён шаблон: {path!r}, полей={len(fields_full)}")
        except Exception as e:
            self._log(f"Ошибка сохранения шаблона: {e}", "error")
            QMessageBox.critical(self, "Ошибка", str(e))

    def _on_new(self):
        self._reset_adaptation_session_state()
        from .cell_items import FieldRectItem
        self._document_property_field_defs = []
        to_remove = [
            it for it in self._gfx_view.scene().items()
            if isinstance(it, FieldRectItem)
        ]
        for it in to_remove:
            self._gfx_view.scene().removeItem(it)
        self._grid_lines = []
        if self._cells_panel:
            self._rebuild_cells_list()
        self._set_active_template(None)
        self._set_modified(False)
        self._update_adapt_action_state()
        self._log("Новый шаблон (canvas очищен)")

    def _current_pdf_page_1based(self) -> int:
        """Номер страницы PDF (1…N) по выпадающему списку."""
        idx = self._page_combo.currentIndex()
        if idx < 0:
            return 1
        return idx + 1

    def _rebuild_page_combo(self, num_pages: int, select_1based: int = 1) -> None:
        """Заполнить список страниц «1»…«N»; при num_pages==0 — пусто и disabled."""
        self._page_combo.blockSignals(True)
        self._page_combo.clear()
        if num_pages <= 0:
            self._page_combo.setEnabled(False)
            self._page_combo.blockSignals(False)
            return
        for p in range(1, num_pages + 1):
            self._page_combo.addItem(str(p))
        pick = max(1, min(select_1based, num_pages))
        self._page_combo.setCurrentIndex(pick - 1)
        self._page_combo.setEnabled(True)
        self._page_combo.blockSignals(False)

    # ---- test debug PNG ----
    def _generate_test_debug_png(
        self,
        result,
        items: list,
        page_num: int,
    ) -> str:
        """Render debug PNG with PDF text spans + field bboxes. Returns path or ''."""
        fitz_page = self._gfx_view.current_fitz_page
        if fitz_page is None:
            return ""
        try:
            from pdf_parsing_v2_engine.debug_visual import render_debug_page

            pdf_base = os.path.splitext(os.path.basename(self._pdf_path))[0] if self._pdf_path else "nofile"
            png_name = f"te_debug_{pdf_base}_p{page_num}.png"
            png_path = os.path.join(tempfile.gettempdir(), png_name)

            field_expected_map = {
                it.field_def.id: it.field_def.expected
                for it in items
                if it.field_def is not None
            }

            render_debug_page(
                fitz_page,
                result,
                png_path,
                field_expected_map=field_expected_map,
                show_text_spans=True,
            )
            self._last_test_png = png_path
            self.act_open_test_png.setEnabled(True)
            self._log(f"Debug PNG: {png_path}")
            return png_path
        except Exception as exc:
            self._log(f"Ошибка генерации debug PNG: {exc}", "warn")
            return ""

    def _on_open_test_png(self) -> None:
        if not self._last_test_png or not os.path.isfile(self._last_test_png):
            QMessageBox.warning(self, "PNG теста", "Файл не найден. Сначала запустите тест (F5).")
            return
        try:
            os.startfile(self._last_test_png)  # Windows — открывает в просмотрщике по умолчанию
        except Exception:
            subprocess.Popen(["explorer", self._last_test_png])

    # ---- catalog / set editors ----
    def _on_catalog_editor(self, catalog_path: str = ""):
        """Open catalog editor panel (T7.2+). Falls back to modal dialog until panel is ready."""
        # If catalog_path provided (from browser context menu), activate it first
        if catalog_path and os.path.isfile(catalog_path):
            self._on_catalog_browser_activate(catalog_path)

        # T7.2: show embedded panel if available; fall back to modal dialog
        if hasattr(self, "_catalog_panel") and self._catalog_panel is not None:
            self._show_catalog_panel()
            return

        # Legacy modal dialog (until T7.2 is implemented)
        from .catalog_editor import CatalogEditorDialog
        from .cell_items import FieldRectItem

        scene_fds = [
            it.field_def for it in self._scene_field_items()
            if isinstance(it, FieldRectItem) and it.field_def
        ]
        dlg = CatalogEditorDialog(
            self._catalog,
            self._templates_dir,
            current_path=self._catalog_path,
            scene_fields=scene_fds or None,
            parent=self,
        )
        if dlg.exec() == dlg.DialogCode.Accepted:
            self._catalog = dlg.get_catalog()
            if dlg.current_path:
                self._catalog_path = dlg.current_path
            if self._props_panel:
                self._props_panel.set_catalog(self._catalog)
            if self._cells_panel:
                self._cells_panel.set_catalog(self._catalog)
            self._log(f"Каталог обновлён: {len(self._catalog.entries)} записей")

            if scene_fds:
                updated = dlg.compute_sync_to_template(scene_fds)
                self._apply_catalog_sync_to_scene(updated)
                if dlg.sync_summary:
                    self._log(dlg.sync_summary)
                    self.statusBar().showMessage(dlg.sync_summary, 4000)

    def _on_props_edit_field_type_requested(self, field_id: str) -> None:
        """Properties panel: open catalog editor and focus the requested field row."""
        fid = (field_id or "").strip()
        if not fid:
            return
        self._on_catalog_editor()
        if self._catalog_panel is None:
            QMessageBox.information(
                self,
                "Каталог",
                "Редактор каталога недоступен в текущем режиме.",
            )
            return
        found = self._catalog_panel.focus_row_by_field_id(fid)
        if found:
            self.statusBar().showMessage(
                f"Каталог: выбрана строка для поля «{fid}».",
                3000,
            )
            return
        QMessageBox.information(
            self,
            "Каталог",
            f"Строка для поля «{fid}» не найдена в активном каталоге.",
        )

    # ------------------------------------------------------------------ catalog panel (T7.2+)

    def _show_catalog_panel(self) -> None:
        """Show the embedded catalog panel (restores from hidden state)."""
        if self._catalog_is_floating and self._catalog_floating_win:
            self._catalog_floating_win.show()
            self._catalog_floating_win.raise_()
            return
        if self._catalog_panel:
            self._catalog_panel.show()
            if self._catalog and self._catalog_panel._catalog is None:
                self._catalog_panel.load_catalog(self._catalog, self._catalog_path)
            # Update coverage
            self._update_catalog_panel_coverage()

    def _hide_catalog_panel(self) -> None:
        if self._catalog_panel:
            self._catalog_panel.hide()

    def _schedule_catalog_coverage_update(self) -> None:
        """Debounced trigger for _update_catalog_panel_coverage (T7.1).

        Safe to call from any scene-mutation handler; rapid successive calls
        collapse into a single update 100 ms after the last call.
        """
        self._catalog_coverage_timer.start()

    def _update_catalog_panel_coverage(self) -> None:
        """Sync canvas fields to catalog panel (coverage highlight, nav buttons).

        Uses isVisible() (checks full parent chain) instead of isHidden() so
        that a floating/detached catalog window that has been closed or minimised
        is correctly skipped, while one that is merely behind another window is
        still updated.
        """
        if not self._catalog_panel or not self._catalog_panel.isVisible():
            return
        from .cell_items import FieldRectItem

        items = [
            it for it in self._scene_field_items()
            if isinstance(it, FieldRectItem) and it.field_def is not None
        ]
        items_sorted = sorted(items, key=lambda it: it.cell_index)
        scene_fds = [it.field_def for it in items_sorted]
        cell_indices = [it.cell_index for it in items_sorted]
        doc_fds = list(self._document_property_field_defs)
        if doc_fds:
            base_idx = max(cell_indices, default=0)
            for j, dfd in enumerate(doc_fds):
                scene_fds.append(dfd)
                cell_indices.append(base_idx + j + 1)
        self._catalog_panel.update_scene_fields(scene_fds, cell_indices=cell_indices)

    def _on_catalog_detach_toggle(self) -> None:
        """Toggle between docked and floating catalog panel."""
        if self._catalog_is_floating:
            self._dock_catalog_panel()
        else:
            self._detach_catalog_panel()

    def _on_prealign_panel_toggle(self, visible: bool) -> None:
        """Показать/скрыть панель prealign (док или откреплённое окно)."""
        if not self._prealign_overlay_panel:
            return
        if self._prealign_is_floating and self._prealign_floating_win is not None:
            self._prealign_floating_win.setVisible(visible)
            if visible:
                self._prealign_floating_win.raise_()
        else:
            s_cfg = self._pdf_recent_settings()
            prefer_floating = s_cfg.value(_SETTINGS_PREALIGN_OPEN_FLOATING, True)
            prefer_floating = prefer_floating in (True, "true", 1, "1")
            self._prealign_overlay_panel.setMinimumHeight(96 if visible else 0)
            self._prealign_overlay_panel.setVisible(visible)
            if visible and prefer_floating:
                self._detach_prealign_panel()
            elif visible and self._left_v_splitter is not None:
                sp = self._left_v_splitter.sizes()
                if len(sp) >= 3 and sp[1] < 72:
                    total = max(sum(sp), 800)
                    cat_h = sp[2]
                    mid = min(260, max(120, total // 4))
                    top = max(360, total - mid - cat_h)
                    self._left_v_splitter.setSizes([top, mid, cat_h])
        if visible:
            self._apply_all_canvas_layers_from_panel()

    def _on_prealign_detach_toggle(self) -> None:
        """Toggle between docked and floating prealign overlay panel."""
        if self._prealign_is_floating:
            self._dock_prealign_panel()
        else:
            self._detach_prealign_panel()

    def _detach_prealign_panel(self) -> None:
        """Move prealign panel to a floating modeless window (широкие колонки дерева)."""
        if not self._prealign_overlay_panel or self._prealign_is_floating:
            return
        if not self._prealign_overlay_panel.isVisible():
            self._prealign_overlay_panel.setVisible(True)
        self._pdf_recent_settings().setValue(_SETTINGS_PREALIGN_OPEN_FLOATING, True)
        self._prealign_overlay_panel.setParent(None)
        win = QWidget(None, Qt.WindowType.Window)
        win.setWindowTitle("Границы prealign / adapt (оверлей)")
        win.resize(1520, 440)
        win.setMinimumSize(580, 280)
        layout = QVBoxLayout(win)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._prealign_overlay_panel)
        win.show()
        self._prealign_floating_win = win
        self._prealign_is_floating = True
        self._prealign_overlay_panel._btn_detach.setText("⬇")
        self._prealign_overlay_panel._btn_detach.setToolTip("Вернуть в нижнюю панель")

    def _dock_prealign_panel(self, *, persist_floating_pref: bool = True) -> None:
        """Dock prealign panel back under canvas (index 1 in left_v_splitter)."""
        if not self._prealign_overlay_panel or not self._prealign_is_floating:
            return
        if self._prealign_floating_win:
            self._prealign_overlay_panel.setParent(self._left_v_splitter)
            self._left_v_splitter.insertWidget(1, self._prealign_overlay_panel)
            self._prealign_floating_win.close()
            self._prealign_floating_win = None
        self._prealign_is_floating = False
        self._prealign_overlay_panel._btn_detach.setText("⬆")
        self._prealign_overlay_panel._btn_detach.setToolTip("Открыть в отдельном окне")
        if persist_floating_pref:
            self._pdf_recent_settings().setValue(_SETTINGS_PREALIGN_OPEN_FLOATING, False)
        vis = self.act_prealign_borders.isChecked()
        self._prealign_overlay_panel.setMinimumHeight(96 if vis else 0)
        self._prealign_overlay_panel.setVisible(vis)

    def _detach_catalog_panel(self) -> None:
        """Move catalog panel to a floating modeless window."""
        if not self._catalog_panel or self._catalog_is_floating:
            return
        # Remove from splitter
        self._catalog_panel.setParent(None)
        win = QWidget(None, Qt.WindowType.Window)
        win.setWindowTitle("Каталог полей")
        win.resize(1100, 280)
        layout = QVBoxLayout(win)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._catalog_panel)
        win.show()
        self._catalog_floating_win = win
        self._catalog_is_floating = True
        self._catalog_panel._btn_detach.setText("⬇")
        self._catalog_panel._btn_detach.setToolTip("Вернуть в нижнюю панель")

    def _dock_catalog_panel(self) -> None:
        """Dock catalog panel back into the left_v_splitter."""
        if not self._catalog_panel or not self._catalog_is_floating:
            return
        if self._catalog_floating_win:
            self._catalog_panel.setParent(self._left_v_splitter)
            self._left_v_splitter.addWidget(self._catalog_panel)
            self._catalog_floating_win.close()
            self._catalog_floating_win = None
        self._catalog_is_floating = False
        self._catalog_panel._btn_detach.setText("⬆")
        self._catalog_panel._btn_detach.setToolTip("Открыть в отдельном окне")
        self._catalog_panel.show()

    # ---- catalog panel slots ----

    def _on_catalog_panel_saved(self, catalog) -> None:
        self._catalog = catalog
        self._catalog_path = self._catalog_panel.current_path if self._catalog_panel else ""
        if self._props_panel:
            self._props_panel.set_catalog(catalog)
        if self._cells_panel:
            self._cells_panel.set_catalog(catalog)
        self._log(f"Каталог сохранён: {len(catalog.entries)} записей")
        self.statusBar().showMessage(f"Каталог сохранён: {len(catalog.entries)} полей", 3000)

    def _on_catalog_navigate(self, field_id: str) -> None:
        """Navigate canvas to field with given id (T7.4)."""
        from .cell_items import FieldRectItem
        for it in self._scene_field_items():
            if not isinstance(it, FieldRectItem):
                continue
            fd = it.field_def
            if fd and fd.id == field_id:
                self._gfx_view.centerOn(it)
                scene = self._gfx_view.scene()
                if scene:
                    scene.clearSelection()
                it.setSelected(True)
                it.highlight(True)
                if self._cells_panel:
                    self._cells_panel.highlight_items([it])
                if self._props_panel:
                    self._props_panel.load_from_item(it)
                return
        for dfd in self._document_property_field_defs:
            if dfd.id == field_id:
                scene = self._gfx_view.scene()
                if scene:
                    scene.clearSelection()
                if self._cells_panel:
                    self._cells_panel.select_field_id(field_id)
                if self._props_panel:
                    self._props_panel.load_from_document_property_field(dfd)
                return

    def _on_catalog_add_field_to_workspace(self, fd) -> None:
        """Контекстное меню каталога: запуск place_mode для размещения поля (T7.3).

        Вместо размещения на фиксированном месте — переходим в place_mode:
        курсор становится перекрестием, пунктирный прямоугольник следует за мышью,
        клик размещает поле. Escape отменяет операцию.
        """
        from pdf_parsing_v2_engine.models import FieldDef
        from .cell_items import FieldRectItem

        if not isinstance(fd, FieldDef):
            return
        fid = (fd.id or "").strip()
        if not fid:
            return
        if fd.document_property:
            for existing in self._document_property_field_defs:
                if existing.id == fid:
                    QMessageBox.information(
                        self,
                        "Каталог",
                        f"Свойство PDF «{fid}» уже есть в шаблоне.",
                    )
                    return
            self._document_property_field_defs.append(fd)
            self._set_modified(True)
            self._rebuild_cells_list()
            self._update_catalog_panel_coverage()
            self._log(f"Каталог: добавлено свойство PDF {fid!r}")
            self.statusBar().showMessage(
                f"Свойство PDF «{fid}» добавлено в шаблон (без холста).",
                5000,
            )
            return
        if self._gfx_view.current_fitz_page is None:
            QMessageBox.warning(
                self, "Каталог",
                "Сначала откройте PDF, чтобы разместить поле.",
            )
            return
        if self._frame_info is None:
            QMessageBox.warning(
                self, "Каталог",
                "Рамка штампа не найдена. Откройте страницу PDF, где определяется рамка.",
            )
            return
        for it in self._scene_field_items():
            if isinstance(it, FieldRectItem) and it.field_def and it.field_def.id == fid:
                QMessageBox.information(
                    self, "Каталог",
                    f"Поле «{fid}» уже есть на рабочей области.",
                )
                return

        # Compute preview size in scene pixels from the default bbox_mm
        from pdf_parsing_v2_engine.frame_detector import SCALE as _SCALE
        from .catalog_editor import (
            _DEFAULT_CATALOG_NEW_FIELD_BBOX_MM as _DEF_BBOX,
        )
        dpi_scale = self._gfx_view._dpi / 72.0
        h1, v1, h2, v2 = _DEF_BBOX
        w_px = (h2 - h1) * _SCALE * dpi_scale
        h_px = (v2 - v1) * _SCALE * dpi_scale

        self._pending_catalog_field = fd
        self._gfx_view.set_place_mode(True, preview_w_px=w_px, preview_h_px=h_px)
        self.statusBar().showMessage(
            f"Разместите поле «{fid}»: кликните на canvas. Escape — отмена.",
            0,  # persistent until replaced
        )

    def _on_place_confirmed(self, placed_rect: QRectF) -> None:
        """place_confirmed signal: create FieldRectItem at clicked position (T7.3)."""
        from .cell_items import FieldRectItem

        fd = self._pending_catalog_field
        self._pending_catalog_field = None

        if fd is None:
            return

        fid = (fd.id or "").strip()
        existing = self._scene_field_items()
        next_idx = max((it.cell_index for it in existing), default=0) + 1

        item = FieldRectItem(
            placed_rect.x(), placed_rect.y(),
            placed_rect.width(), placed_rect.height(),
            cell_index=next_idx,
        )
        item.assign_field(fd)
        self._gfx_view.scene().addItem(item)
        item.setSelected(True)
        if self._cells_panel:
            self._rebuild_cells_list()
            self._cells_panel.highlight_items([item])
        if self._props_panel:
            self._props_panel.load_from_item(item)
        self._gfx_view.centerOn(item)
        self._set_modified(True)
        self._update_catalog_panel_coverage()
        self._clear_duplicate_id_highlights()
        self._log(f"Каталог: поле {fid!r} размещено по клику")
        self.statusBar().showMessage(
            f"Поле «{fid}» добавлено. Перетащите и подгоните размер при необходимости.",
            5000,
        )

    def _on_catalog_field_link_apply(self, scene_field_id: str, merged) -> None:
        """Context menu «Привязать к полю»: write catalog semantics onto chosen canvas field."""
        from pdf_parsing_v2_engine.models import FieldDef
        from .cell_items import FieldRectItem

        if not isinstance(merged, FieldDef):
            return
        it_found = None
        old_fd = None
        for it in self._scene_field_items():
            if isinstance(it, FieldRectItem) and it.field_def and it.field_def.id == scene_field_id:
                it_found = it
                old_fd = it.field_def
                break
        if it_found is None or old_fd is None:
            self._log(
                f"Привязка каталога: поле «{scene_field_id}» не найдено на сцене",
                "warn",
            )
            QMessageBox.warning(
                self,
                "Привязка",
                f"Поле «{scene_field_id}» не найдено на сцене.\n"
                "Возможно, список полей уже изменился — обновите привязку.",
            )
            return
        it_found.assign_field(merged)
        self._set_modified(True)
        if self._cells_panel:
            self._rebuild_cells_list()
        self._update_catalog_panel_coverage()
        if self._catalog_panel is not None:
            self._catalog_panel.refresh_sync_links_after_scene_apply([old_fd], [merged])
            self._catalog_panel._sync_refresh_scene_links()
            self._catalog_panel._apply_coverage()
        if self._props_panel:
            self._props_panel.refresh_from_current_item()
        self._log(
            f"Каталог→шаблон: поле «{scene_field_id}» обновлено из строки каталога → id={merged.id!r}",
        )
        self.statusBar().showMessage(
            f"Поле обновлено из каталога: {merged.id}",
            4000,
        )

    def _on_catalog_sync_requested(self) -> None:
        """Catalog sync mode: push table edits to canvas using a fresh snapshot of the scene."""
        if not self._catalog_panel or not self._catalog_panel._sync_enabled:
            return
        items, fds_canvas = self._ordered_scene_field_items_and_defs()
        fds_doc = list(self._document_property_field_defs)
        fds_all = fds_canvas + fds_doc
        if not fds_all:
            return
        updated_all = self._catalog_panel.compute_sync_to_template(fds_all)
        n_canvas = len(fds_canvas)
        self._apply_catalog_sync_to_scene(updated_all[:n_canvas], items=items)
        self._apply_catalog_sync_to_document_properties(fds_doc, updated_all[n_canvas:])

    def _on_template_save_after_catalog_sync(self) -> None:
        """After catalog JSON save with «Синхр.» on: persist template so field defs match catalog."""
        if not self._catalog_panel or not self._catalog_panel._sync_enabled:
            return
        self._log("Режим «Синхр.»: сохранение шаблона после каталога…")
        self._on_save_template()

    def _apply_catalog_sync_to_scene(
        self,
        synced_fields: list[FieldDef],
        *,
        items: list | None = None,
    ) -> None:
        """Apply updated FieldDefs (from catalog sync) back to scene items."""
        from pdf_parsing_v2_engine.models import FieldDef as _FD
        from .cell_items import FieldRectItem

        if items is None:
            items = [
                it for it in self._scene_field_items()
                if isinstance(it, FieldRectItem) and it.field_def is not None
            ]
        if len(items) != len(synced_fields):
            self._log(
                f"sync: несовпадение длины списков (canvas={len(items)}, sync={len(synced_fields)})",
                "warn",
            )
            return
        old_fds = [it.field_def for it in items]
        changed = 0
        for it, new_fd in zip(items, synced_fields):
            old = it.field_def
            old_lbl = (old.label or "").strip()
            new_lbl = (new_fd.label or "").strip()
            old_ft = normalize_field_type(old.field_type)
            new_ft = normalize_field_type(new_fd.field_type)
            old_exp = normalize_expected(old.expected)
            new_exp = normalize_expected(new_fd.expected)
            old_et = (old.expected_text or "").strip()
            new_et = (new_fd.expected_text or "").strip()
            if (
                old.id != new_fd.id
                or old_lbl != new_lbl
                or old_exp != new_exp
                or old_ft != new_ft
                or old_et != new_et
            ):
                merged = _FD(
                    id=new_fd.id,
                    label=new_fd.label,
                    bbox_mm=old.bbox_mm,
                    clean=old.clean,
                    validate_regex=old.validate_regex,
                    expected=new_fd.expected,
                    padding_mm=old.padding_mm,
                    field_type=new_fd.field_type,
                    expected_text=new_fd.expected_text,
                    is_anchor=old.is_anchor,
                    outside_stamp=old.outside_stamp,
                    origin=old.origin,
                    stretch_to_page=old.stretch_to_page,
                    bound_top=old.bound_top,
                    bound_bottom=old.bound_bottom,
                    bound_left=old.bound_left,
                    bound_right=old.bound_right,
                    document_property=old.document_property,
                )
                it.assign_field(merged)
                changed += 1
        if self._catalog_panel is not None:
            self._catalog_panel.refresh_sync_links_after_scene_apply(old_fds, synced_fields)
        if changed:
            self._set_modified(True)
            if self._catalog_panel is not None:
                try:
                    self._catalog = self._catalog_panel.get_catalog()
                except ValueError:
                    pass
                else:
                    if self._props_panel:
                        self._props_panel.set_catalog(self._catalog)
                    if self._cells_panel:
                        self._cells_panel.set_catalog(self._catalog)
            if self._cells_panel:
                self._rebuild_cells_list()
            self._update_catalog_panel_coverage()
            if self._props_panel:
                self._props_panel.refresh_from_current_item()
        elif self._catalog_panel is not None:
            self._update_catalog_panel_coverage()

    def _apply_catalog_sync_to_document_properties(
        self,
        old_docs: list[FieldDef],
        synced_docs: list[FieldDef],
    ) -> None:
        """Apply catalog row edits to hidden PDF document-property fields (not on canvas)."""
        if not old_docs and not synced_docs:
            return
        if len(old_docs) != len(synced_docs):
            self._log(
                "sync: несовпадение длины списков (document_property)",
                "warn",
            )
            return
        new_list = list(self._document_property_field_defs)
        changed = 0
        for i, (old, new_fd) in enumerate(zip(old_docs, synced_docs)):
            old_lbl = (old.label or "").strip()
            new_lbl = (new_fd.label or "").strip()
            old_ft = normalize_field_type(old.field_type)
            new_ft = normalize_field_type(new_fd.field_type)
            old_exp = normalize_expected(old.expected)
            new_exp = normalize_expected(new_fd.expected)
            old_et = (old.expected_text or "").strip()
            new_et = (new_fd.expected_text or "").strip()
            if (
                old.id != new_fd.id
                or old_lbl != new_lbl
                or old_exp != new_exp
                or old_ft != new_ft
                or old_et != new_et
            ):
                new_list[i] = new_fd
                changed += 1
        if not changed:
            return
        self._document_property_field_defs = new_list
        self._set_modified(True)
        if self._catalog_panel is not None:
            try:
                self._catalog = self._catalog_panel.get_catalog()
            except ValueError:
                pass
            else:
                if self._props_panel:
                    self._props_panel.set_catalog(self._catalog)
                if self._cells_panel:
                    self._cells_panel.set_catalog(self._catalog)
        self._rebuild_cells_list()
        self._update_catalog_panel_coverage()
        if self._props_panel:
            self._props_panel.refresh_from_current_item()
        self._log(
            f"Каталог→шаблон: обновлено свойств PDF: {changed}",
        )

    def _on_set_editor(self):
        from .set_editor import SetEditorDialog
        dlg = SetEditorDialog(self, templates_dir=self._templates_dir)
        dlg.exec()
        self._templates_browser.refresh()

    def _on_browser_message(self, level: str, text: str) -> None:
        lv = level if level in ("info", "warn", "error") else "info"
        self._log(text, lv)
        if lv == "error":
            self.statusBar().showMessage(text, 5000)

    # ---- template browser handlers ----
    def _on_template_browser_load(self, path: str) -> None:
        """Double-click on template in browser — ask if unsaved changes."""
        if self._is_modified or self._catalog_is_dirty:
            r = QMessageBox.question(
                self,
                "Несохранённые изменения",
                "Есть несохранённые изменения (шаблон и/или каталог).\n"
                "Загрузить другой шаблон без сохранения?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                return
        self._load_template_from_file(path, clear_first=True)

    def _on_catalog_browser_activate(self, path: str) -> None:
        """Double-click on catalog.json in browser — activate and show catalog panel."""
        if self._catalog_is_dirty:
            r = QMessageBox.question(
                self,
                "Несохранённые изменения",
                "Каталог полей изменён и не сохранён.\n"
                "Открыть другой файл каталога без сохранения?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                return
        try:
            cat = FieldCatalog.from_json(path)
        except Exception as e:
            self._log(f"Ошибка загрузки каталога {path!r}: {e}", "error")
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить каталог:\n{e}")
            return
        self._catalog = cat
        self._catalog_path = path
        if self._props_panel:
            self._props_panel.set_catalog(cat)
        if self._cells_panel:
            self._cells_panel.set_catalog(cat)
        self._templates_browser.set_active_catalog(path)
        # Mark the project folder as active
        self._templates_browser.set_active_project(os.path.dirname(path))
        self._log(f"Активирован каталог: {os.path.basename(path)!r}, записей={len(cat.entries)}")
        self.statusBar().showMessage(
            f"Каталог: {os.path.basename(path)} ({len(cat.entries)} полей)", 3000
        )
        # T7.1: load into catalog panel and show it
        if self._catalog_panel is not None:
            self._catalog_panel.load_catalog(cat, path)
            self._show_catalog_panel()

    # ---- geometry changed (drag/resize) ----
    def _on_item_geometry_changed(self, item) -> None:
        self._clear_preview_overlays()
        self._set_modified(True)
        if self._props_panel:
            self._props_panel.refresh_bbox_from_item(item)

    # ---- panel wiring ----
    def _on_cell_selected_in_panel(self, item):
        from .cell_items import FieldRectItem
        if isinstance(item, FieldDef) and item.document_property:
            self._gfx_view.scene().clearSelection()
            if self._props_panel:
                self._props_panel.load_from_document_property_field(item)
            return
        if isinstance(item, FieldRectItem):
            self._gfx_view.centerOn(item)
            item.highlight(True)
            if self._props_panel:
                self._props_panel.load_from_item(item)

    def _on_cell_double_clicked(self, item):
        if isinstance(item, FieldDef) and item.document_property:
            if self._props_panel:
                self._props_panel.load_from_document_property_field(item)
            if self._cells_panel is None:
                return
            rows = self._cells_panel.get_field_def_details_rows(item)
            if not rows:
                return
            title = f"Детали: {item.id}"
            from .cell_details_dialog import CellDetailsDialog
            CellDetailsDialog(title, rows, self).exec()
            return
        if self._props_panel:
            self._props_panel.load_from_item(item)
        if self._cells_panel is None:
            return
        rows = self._cells_panel.get_item_details_rows(item)
        if not rows:
            return
        title = "Детали строки поля"
        if getattr(item, "field_def", None) and item.field_def:
            fid = (item.field_def.id or "").strip()
            if fid:
                title = f"Детали строки поля: {fid}"
        from .cell_details_dialog import CellDetailsDialog
        dlg = CellDetailsDialog(title, rows, self)
        dlg.exec()

    def _on_document_property_field_changed(self, fd: FieldDef) -> None:
        for i, old in enumerate(self._document_property_field_defs):
            if old.id == fd.id:
                self._document_property_field_defs[i] = fd
                break
        self._set_modified(True)
        self._rebuild_cells_list()

    def _on_field_changed(self, item):
        self._clear_preview_overlays()
        self._clear_origin_line()
        if self._cells_panel:
            self._cells_panel.update_item_row(item)
        self._set_modified(True)
        from .cell_items import FieldRectItem
        if (
            isinstance(item, FieldRectItem)
            and item.field_def
            and self._frame_info is not None
        ):
            self._snap_item_to_field_def(item, item.field_def)
        if isinstance(item, FieldRectItem) and item.isSelected():
            self._draw_origin_line(item)
        if self.act_wireframe_mode.isChecked():
            self._sync_wireframe_field_selection()
        # Schedule catalog coverage update (T7.1: debounced)
        self._schedule_catalog_coverage_update()
        # Sync mode: notify catalog panel of field change
        if self._catalog_panel and self._catalog_panel._sync_enabled:
            fd = item.field_def if isinstance(item, FieldRectItem) else None
            self._catalog_panel.on_scene_field_changed("rename", fd)

    def _on_scene_selection_changed(self):
        from .cell_items import FieldRectItem
        self._clear_preview_overlays()
        self._clear_origin_line()

        # Clear geometry callback from previously tracked item
        if self._geo_tracked_item is not None:
            self._geo_tracked_item._geo_cb = None
            self._geo_tracked_item = None

        selected = [
            it for it in self._gfx_view.scene().selectedItems()
            if isinstance(it, FieldRectItem)
        ]
        if len(selected) == 1:
            selected[0]._geo_cb = self._on_item_geometry_changed
            self._geo_tracked_item = selected[0]
            if self._props_panel:
                self._props_panel.load_from_item(selected[0])
            self._draw_origin_line(selected[0])
            # Sync mode: highlight matching row in Field Catalog panel
            if self._catalog_panel and self._catalog_panel._sync_enabled:
                fd = selected[0].field_def
                fid = fd.id if fd else None
                self._catalog_panel.sync_selection_from_scene(fid)
        elif self._catalog_panel and self._catalog_panel._sync_enabled:
            self._catalog_panel.sync_selection_from_scene(None)
        if self._cells_panel:
            self._cells_panel.highlight_items(selected)
        self._sync_wireframe_field_selection()

    # ---- origin markers & line ----
    def _draw_origin_markers(self, frame: FrameInfo | None) -> None:
        """Draw 4 small «+» markers at frame corners (origin points)."""
        self._remove_origin_markers()
        if frame is None:
            return
        from pdf_parsing_v2_engine.coord_transform import get_origin_points

        dpi_scale = self._gfx_view._dpi / 72.0
        origins = get_origin_points(frame)
        scene = self._gfx_view.scene()
        marker_color = QColor(220, 40, 40, 200)
        pen = QPen(marker_color, 2.0)
        sz = 8.0
        for _name, (pm_x, pm_y) in origins.items():
            fitz_x = pm_x
            fitz_y = frame.page_height - pm_y
            sx = fitz_x * dpi_scale
            sy = fitz_y * dpi_scale
            h_line = QGraphicsLineItem(sx - sz, sy, sx + sz, sy)
            h_line.setPen(pen)
            h_line.setZValue(950)
            scene.addItem(h_line)
            v_line = QGraphicsLineItem(sx, sy - sz, sx, sy + sz)
            v_line.setPen(pen)
            v_line.setZValue(950)
            scene.addItem(v_line)
            self._origin_marker_items.extend([h_line, v_line])
        vis_om = (
            self._prealign_overlay_panel.is_canvas_layer_visible("origin_markers")
            if self._prealign_overlay_panel
            else True
        )
        for it in self._origin_marker_items:
            it.setVisible(vis_om)

    def _remove_origin_markers(self) -> None:
        scene = self._gfx_view.scene()
        for it in self._origin_marker_items:
            if it.scene() is scene:
                scene.removeItem(it)
        self._origin_marker_items.clear()

    def _draw_origin_line(self, item) -> None:
        """Thin line from field center to its origin marker."""
        self._clear_origin_line()
        if self._frame_info is None or item.field_def is None:
            return
        from pdf_parsing_v2_engine.coord_transform import get_origin_points

        dpi_scale = self._gfx_view._dpi / 72.0
        origins = get_origin_points(self._frame_info)
        origin_name = item.field_def.origin
        pm_pt = origins.get(origin_name)
        if pm_pt is None:
            return
        fitz_x = pm_pt[0]
        fitz_y = self._frame_info.page_height - pm_pt[1]
        ox = fitz_x * dpi_scale
        oy = fitz_y * dpi_scale
        sx, sy, sw, sh = self._item_scene_rect(item)
        cx = sx + sw * 0.5
        cy = sy + sh * 0.5
        line = QGraphicsLineItem(cx, cy, ox, oy)
        line.setPen(QPen(QColor(220, 40, 40, 120), 1.0, Qt.PenStyle.DashDotLine))
        line.setZValue(940)
        self._gfx_view.scene().addItem(line)
        self._origin_line_item = line
        if self._prealign_overlay_panel and not self._prealign_overlay_panel.is_canvas_layer_visible(
            "origin_markers",
        ):
            line.setVisible(False)

    def _clear_origin_line(self) -> None:
        if self._origin_line_item is not None:
            scene = self._gfx_view.scene()
            if self._origin_line_item.scene() is scene:
                scene.removeItem(self._origin_line_item)
            self._origin_line_item = None

    # ---- catalog auto-load ----
    def _try_auto_load_catalog(self) -> None:
        """Auto-load first non-empty catalog.json found in project folders."""
        if self._catalog is not None:
            return
        if not self._templates_dir:
            return
        from pdf_parsing_v2_engine.template_loader import load_projects
        projects = load_projects(self._templates_dir)
        for proj in projects:
            try:
                from pdf_parsing_v2_engine.models import FieldCatalog as _FC
                cat = _FC.from_json(proj.catalog_path)
                if not cat.entries:
                    continue
                self._catalog = cat
                self._catalog_path = proj.catalog_path
                if self._props_panel:
                    self._props_panel.set_catalog(cat)
                if self._cells_panel:
                    self._cells_panel.set_catalog(cat)
                if hasattr(self, "_templates_browser") and self._templates_browser:
                    self._templates_browser.set_active_catalog(proj.catalog_path)
                    self._templates_browser.set_active_project(proj.folder)
                self._log(f"Авто-загрузка каталога: {proj.name!r}, записей={len(cat.entries)}")
                self.statusBar().showMessage(
                    f"Каталог загружен: {proj.name} ({len(cat.entries)} полей)", 4000
                )
                return
            except Exception as e:
                self._log(f"Авто-загрузка каталога {proj.catalog_path!r}: ошибка: {e}", "warn")

        # Legacy fallback: templates_dir/catalogs/ (old structure)
        cat_dir = os.path.join(self._templates_dir, "catalogs")
        if os.path.isdir(cat_dir):
            for fname in sorted(os.listdir(cat_dir)):
                if fname.lower().endswith(".json"):
                    full = os.path.join(cat_dir, fname)
                    try:
                        from pdf_parsing_v2_engine.models import FieldCatalog as _FC
                        cat = _FC.from_json(full)
                        if not cat.entries:
                            continue
                        self._catalog = cat
                        self._catalog_path = full
                        if self._props_panel:
                            self._props_panel.set_catalog(cat)
                        if self._cells_panel:
                            self._cells_panel.set_catalog(cat)
                        if hasattr(self, "_templates_browser") and self._templates_browser:
                            self._templates_browser.set_active_catalog(full)
                        self._log(f"Авто-загрузка каталога (legacy): {fname!r}, записей={len(cat.entries)}")
                        self.statusBar().showMessage(
                            f"Каталог загружен: {fname} ({len(cat.entries)} полей)", 4000
                        )
                        return
                    except Exception as e:
                        self._log(f"Авто-загрузка каталога {fname!r}: ошибка: {e}", "warn")

    # ================================================================== wireframe mode

    def _on_wireframe_mode_toggled(self, enabled: bool) -> None:
        """Toggle between normal edit mode and wireframe mode."""
        from .cell_items import FieldRectItem

        if self._cells_panel is None:
            return
        if enabled:
            self._cells_panel.hide()
            if self._meta_panel:
                self._meta_panel.hide()
            if self._props_panel:
                self._props_panel.hide()
            self._grid_lines_panel.set_lines(self._grid_lines)
            self._grid_lines_panel.show()
            self._field_bindings_panel.set_lines(self._grid_lines)
            self._field_bindings_panel.show()
            self._refresh_wireframe_canvas()
            self._sync_wireframe_field_selection()
        else:
            self._gfx_view.set_wire_line_place_mode(None)
            self._grid_lines_panel.hide()
            self._field_bindings_panel.hide()
            self._cells_panel.show()
            if self._meta_panel:
                self._meta_panel.show()
            if self._props_panel:
                self._props_panel.show()
            if self._grid_lines_panel._chk_hide_fields.isChecked():
                self._grid_lines_panel._chk_hide_fields.setChecked(False)
            self._on_wireframe_hide_fields(False)
            # Keep wireframe visible if snapped (after adaptation)
            if not self._snapped_line_positions:
                self._clear_wireframe_canvas()
        # Update binding-state fill colors on all field items.
        for it in self._gfx_view.scene().items():
            if isinstance(it, FieldRectItem):
                it.set_wireframe_mode(enabled)

    def _on_wireframe_hide_fields(self, hide: bool) -> None:
        """Скрыть только прямоугольники полей (как подпись чекбокса). Оверлеи adapt — отдельно в панели «Границы»."""
        from .cell_items import FieldRectItem

        vis = not hide
        for it in self._gfx_view.scene().items():
            if isinstance(it, FieldRectItem):
                it.setVisible(vis)
        if self._prealign_overlay_panel:
            self._prealign_overlay_panel.sync_canvas_checkbox("stamp_fields", vis)

    def _set_stamp_field_rects_visible(self, visible: bool) -> None:
        from .cell_items import FieldRectItem

        for it in self._gfx_view.scene().items():
            if isinstance(it, FieldRectItem):
                it.setVisible(visible)

    def _sync_wireframe_visibility_checks_from_panel(self) -> None:
        if not self._grid_lines_panel or not self._prealign_overlay_panel:
            return
        p = self._prealign_overlay_panel
        self._grid_lines_panel.apply_external_visibility_checks(
            stamp_fields_visible=p.is_canvas_layer_visible("stamp_fields"),
            template_grid_visible=p.is_canvas_layer_visible("template_grid"),
            pdf_grid_visible=p.is_canvas_layer_visible("pdf_grid"),
        )

    def _on_canvas_layer_visibility_changed(self, layer_id: str, visible: bool) -> None:
        if layer_id == "stamp_fields":
            self._set_stamp_field_rects_visible(visible)
        elif layer_id == "template_grid":
            for item in self._wireframe_items:
                item.setVisible(visible)
        elif layer_id == "pdf_grid":
            for item in self._detected_grid_overlay:
                item.setVisible(visible)
        elif layer_id == "adapt_overlay":
            for item in self._adapt_overlay:
                if isinstance(item, QGraphicsItem):
                    item.setVisible(visible)
        elif layer_id == "consolidation_overlay":
            for item in self._consolidation_overlay:
                if isinstance(item, QGraphicsItem):
                    item.setVisible(visible)
        elif layer_id == "origin_markers":
            for it in self._origin_marker_items:
                it.setVisible(visible)
            if self._origin_line_item is not None:
                self._origin_line_item.setVisible(visible)
        if layer_id in ("stamp_fields", "template_grid", "pdf_grid"):
            self._sync_wireframe_visibility_checks_from_panel()

    def _apply_adapt_overlay_visibility(self) -> None:
        vis = (
            self._prealign_overlay_panel.is_canvas_layer_visible("adapt_overlay")
            if self._prealign_overlay_panel
            else True
        )
        for item in self._adapt_overlay:
            if isinstance(item, QGraphicsItem):
                item.setVisible(vis)

    def _apply_consolidation_overlay_visibility(self) -> None:
        vis = (
            self._prealign_overlay_panel.is_canvas_layer_visible("consolidation_overlay")
            if self._prealign_overlay_panel
            else True
        )
        for item in self._consolidation_overlay:
            if isinstance(item, QGraphicsItem):
                item.setVisible(vis)

    def _apply_detected_grid_overlay_visibility(self) -> None:
        vis = (
            self._prealign_overlay_panel.is_canvas_layer_visible("pdf_grid")
            if self._prealign_overlay_panel
            else True
        )
        for item in self._detected_grid_overlay:
            item.setVisible(vis)

    def _apply_all_canvas_layers_from_panel(self) -> None:
        """Применить все флаги видимости canvas (после показа панели или загрузки сцены)."""
        if not self._prealign_overlay_panel:
            return
        p = self._prealign_overlay_panel
        self._set_stamp_field_rects_visible(p.is_canvas_layer_visible("stamp_fields"))
        for item in self._wireframe_items:
            item.setVisible(p.is_canvas_layer_visible("template_grid"))
        for item in self._detected_grid_overlay:
            item.setVisible(p.is_canvas_layer_visible("pdf_grid"))
        self._apply_adapt_overlay_visibility()
        self._apply_consolidation_overlay_visibility()
        om = p.is_canvas_layer_visible("origin_markers")
        for it in self._origin_marker_items:
            it.setVisible(om)
        if self._origin_line_item is not None:
            self._origin_line_item.setVisible(om)
        self._sync_wireframe_visibility_checks_from_panel()

    def _canvas_stamp_fields_wanted_visible(self) -> bool:
        if self._prealign_overlay_panel:
            return self._prealign_overlay_panel.is_canvas_layer_visible("stamp_fields")
        return True

    def _on_grid_lines_changed(self) -> None:
        """Grid lines edited in the panel (add/delete/edit cell)."""
        from pdf_parsing_v2_engine.grid_lines_utils import (
            sanitize_field_line_bindings,
            sanitize_grid_line_refs,
            snap_endpoints_to_refs,
        )
        from .cell_items import FieldRectItem

        raw = sanitize_grid_line_refs(self._grid_lines_panel.get_lines())
        if self._grid_lines_panel.line_intersection_snap_enabled():
            self._grid_lines = snap_endpoints_to_refs(raw)
        else:
            self._grid_lines = raw
        valid = {gl.id for gl in self._grid_lines}
        for it in self._scene_field_items():
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def:
                new_fd = sanitize_field_line_bindings(it.field_def, valid)
                if new_fd is not it.field_def:
                    it.assign_field(new_fd)
                    self._snap_item_to_field_def(it, new_fd)

        self._grid_lines_panel.set_lines(self._grid_lines)
        self._field_bindings_panel.set_lines(self._grid_lines)
        self._set_modified(True)
        self._refresh_wireframe_canvas()
        self._snap_bound_fields_to_grid()
        self._sync_wireframe_field_selection()

    def _on_grid_line_selected(self, line_id: str) -> None:
        """A line was selected in the GridLinesPanel table."""
        self._highlight_grid_line(line_id)

    def _on_interactive_add_grid_line(self, orientation: str) -> None:
        """Start two-click placement for +H / +V (canvas defines span)."""
        if not self._frame_info:
            QMessageBox.information(
                self,
                "Wireframe",
                "Откройте PDF и дождитесь определения рамки штампа, затем укажите линию на полотне.",
            )
            return
        self._gfx_view.set_place_mode(False)
        self._exit_pick_line_mode()
        self._gfx_view.set_wire_line_place_mode(orientation)
        self._log("Wireframe: два клика — начало и конец линии (Esc — отмена)")

    def _on_wire_line_place_finished(self, orient: str, a: QPointF, b: QPointF) -> None:
        fi = self._frame_info
        if not fi:
            return
        from pdf_parsing_v2_engine.coord_transform import fitz_pts_to_grid_line_mm

        fax = self._gfx_view.scene_to_pts(a).x()
        fay = self._gfx_view.scene_to_pts(a).y()
        fbx = self._gfx_view.scene_to_pts(b).x()
        fby = self._gfx_view.scene_to_pts(b).y()
        if orient == "h":
            pos_fitz = (fay + fby) / 2.0
            start_fitz, end_fitz = min(fax, fbx), max(fax, fbx)
        else:
            pos_fitz = (fax + fbx) / 2.0
            start_fitz, end_fitz = min(fay, fby), max(fay, fby)
        pos_mm, start_mm, end_mm = fitz_pts_to_grid_line_mm(
            orient, pos_fitz, start_fitz, end_fitz, fi, "frame_bottom_right",
        )
        self._grid_lines_panel.append_line_mm(orient, pos_mm, start_mm, end_mm)
        self._log(
            f"Wireframe: линия ({orient}) pos={pos_mm:.2f} мм, span {start_mm:.2f}…{end_mm:.2f} мм",
        )

    def _on_wireframe_line_delete_key(self) -> None:
        fw = self.focusWidget()
        if isinstance(fw, (QLineEdit, QAbstractSpinBox, QTextEdit)):
            return
        if not self.act_wireframe_mode.isChecked():
            return
        if not self._grid_lines_panel.isVisible():
            return
        if self._grid_lines_panel.selected_line_id() is None:
            return
        self._grid_lines_panel.delete_selected_line()

    def _on_generate_wireframe(self) -> None:
        """Generate wireframe lines from current field bounding boxes."""
        from .cell_items import FieldRectItem
        from pdf_parsing_v2_engine.grid_lines_utils import generate_grid_lines_from_fields

        items = [
            it for it in self._scene_field_items()
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def
            and not it.field_def.outside_stamp
        ]
        if not items:
            QMessageBox.information(self, "Wireframe", "Нет назначенных полей внутри штампа.")
            return

        fields = self._collect_field_defs_from_scene(items)
        new_lines = generate_grid_lines_from_fields(fields)
        if not new_lines:
            QMessageBox.information(self, "Wireframe", "Не удалось сгенерировать линии (нет данных).")
            return

        if self._grid_lines:
            r = QMessageBox.question(
                self, "Wireframe",
                f"Текущих линий: {len(self._grid_lines)}.\n"
                f"Заменить на {len(new_lines)} автоматически сгенерированных?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                return

        self._grid_lines = new_lines
        self._log(f"Wireframe: сгенерировано {len(new_lines)} линий из {len(items)} полей")

        self._check_and_offer_collapse_doubled_lines()

        self._grid_lines_panel.set_lines(self._grid_lines)
        self._field_bindings_panel.set_lines(self._grid_lines)
        self._set_modified(True)
        self._refresh_wireframe_canvas()
        self.statusBar().showMessage(
            f"Сгенерировано {len(self._grid_lines)} линий сетки", 4000,
        )

    def _check_and_offer_collapse_doubled_lines(self) -> None:
        """Detect border-pair doubled grid lines and offer to collapse them."""
        from pdf_parsing_v2_engine.grid_lines_utils import (
            find_doubled_lines, collapse_doubled_lines, find_split_segments,
        )
        from .cell_items import FieldRectItem

        if not self._grid_lines:
            return

        items = [
            it for it in self._scene_field_items()
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def
            and not it.field_def.outside_stamp
        ]
        if not items:
            return

        fields = self._collect_field_defs_from_scene(items)

        splits = find_split_segments(self._grid_lines, fields)
        if splits:
            split_desc = "\n".join(
                f"  {a.id} ({a.pos_mm:.2f} mm, span {a.start_mm:.0f}→{a.end_mm:.0f}) ↔ "
                f"{b.id} ({b.pos_mm:.2f} mm, span {b.start_mm:.0f}→{b.end_mm:.0f})"
                for a, b in splits[:10]
            )
            self._log(
                f"Split segments (близкие позиции, разные spans — это нормально): "
                f"{len(splits)} пар\n{split_desc}"
            )

        pairs = find_doubled_lines(self._grid_lines, fields)
        if not pairs:
            return

        pair_desc = "\n".join(
            f"  {a.id} ({a.pos_mm:.2f} mm) ↔ {b.id} ({b.pos_mm:.2f} mm), "
            f"Δ = {abs(b.pos_mm - a.pos_mm):.2f} mm"
            for a, b in pairs[:15]
        )
        more = f"\n  … и ещё {len(pairs) - 15}" if len(pairs) > 15 else ""
        r = QMessageBox.question(
            self, "Задвоение линий сетки",
            f"Обнаружено {len(pairs)} пар(а) задвоенных линий "
            f"(перекрывающиеся spans, верхний/нижний край одной линии):\n\n{pair_desc}{more}\n\n"
            "Схлопнуть задвоенные линии (медиана → центр линии)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if r != QMessageBox.StandardButton.Yes:
            return

        merged_lines, updated_fields, messages = collapse_doubled_lines(
            self._grid_lines, fields,
        )
        self._grid_lines = merged_lines

        for fd_new in updated_fields:
            for it in items:
                if isinstance(it, FieldRectItem) and it.field_def and it.field_def.id == fd_new.id:
                    it.assign_field(fd_new)
                    break

        if hasattr(self, '_grid_lines_panel') and self._grid_lines_panel:
            self._grid_lines_panel.set_lines(self._grid_lines)
            self._field_bindings_panel.set_lines(self._grid_lines)
            self._refresh_wireframe_canvas()

        self._set_modified(True)
        self._log(
            f"Wireframe: схлопнуто {len(messages)} пар задвоенных линий: "
            + "; ".join(messages[:10])
        )
        self.statusBar().showMessage(
            f"Схлопнуто {len(messages)} пар задвоенных линий", 4000,
        )

    def _snap_item_to_field_def(self, item, fd) -> None:
        """Update FieldRectItem visual rect from fd.bbox_mm (after bind/snap)."""
        from pdf_parsing_v2_engine.coord_transform import field_to_fitz_rect

        frame = self._frame_info
        if frame is None:
            return
        dpi_scale = self._gfx_view._dpi / 72.0
        fitz_rect = field_to_fitz_rect(fd, frame, padding_mm=0)
        sx0 = fitz_rect.x0 * dpi_scale
        sy0 = fitz_rect.y0 * dpi_scale
        sw = max(2.0, fitz_rect.width * dpi_scale)
        sh = max(2.0, fitz_rect.height * dpi_scale)
        item.setPos(QPointF(sx0, sy0))
        item.setRect(QRectF(0.0, 0.0, sw, sh))

    def _snap_bound_fields_to_grid(self) -> None:
        """Snap bbox_mm of all bound fields to current grid line positions.

        Called automatically whenever grid lines change (drag, panel edit,
        auto-refs) so that bound fields track the wireframe without requiring
        an explicit «Привязать все поля» call.
        """
        from .cell_items import FieldRectItem
        from pdf_parsing_v2_engine.grid_lines_utils import snap_fields_to_bindings

        items = [
            it for it in self._scene_field_items()
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def
            and not it.field_def.outside_stamp
            and (it.field_def.bound_top or it.field_def.bound_bottom
                 or it.field_def.bound_left or it.field_def.bound_right)
        ]
        if not items:
            return
        fields = [it.field_def for it in items]
        snapped = snap_fields_to_bindings(fields, self._grid_lines)
        for it, new_fd in zip(items, snapped):
            if new_fd is not it.field_def:
                it.assign_field(new_fd)
                self._snap_item_to_field_def(it, new_fd)

    def _on_autobind_all(self) -> None:
        """Auto-bind all fields to nearest grid lines and snap geometry."""
        from .cell_items import FieldRectItem
        from pdf_parsing_v2_engine.grid_lines_utils import auto_bind_fields_to_lines, snap_fields_to_bindings

        if not self._grid_lines:
            QMessageBox.information(self, "Wireframe", "Сначала сгенерируйте линии сетки.")
            return

        items = [
            it for it in self._scene_field_items()
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def
            and not it.field_def.outside_stamp
        ]
        if not items:
            return

        fields = self._collect_field_defs_from_scene(items)
        bound_fields = auto_bind_fields_to_lines(fields, self._grid_lines)
        snapped_fields = snap_fields_to_bindings(bound_fields, self._grid_lines)

        count_bound = 0
        for it, new_fd in zip(items, snapped_fields):
            if new_fd.bound_top or new_fd.bound_bottom or new_fd.bound_left or new_fd.bound_right:
                count_bound += 1
            it.assign_field(new_fd)
            self._snap_item_to_field_def(it, new_fd)

        self._set_modified(True)
        if self._cells_panel:
            self._rebuild_cells_list()
        self._sync_wireframe_field_selection()
        self._refresh_wireframe_canvas()
        self.statusBar().showMessage(
            f"Привязано и подогнано {count_bound} из {len(items)} полей", 4000,
        )
        self._log(f"Wireframe: auto-bind+snap {count_bound}/{len(items)} полей")

    def _on_autobind_selected(self) -> None:
        """Auto-bind only the currently selected field and snap geometry."""
        from .cell_items import FieldRectItem
        from pdf_parsing_v2_engine.grid_lines_utils import auto_bind_fields_to_lines, snap_fields_to_bindings

        if not self._grid_lines:
            QMessageBox.information(self, "Wireframe", "Сначала сгенерируйте линии сетки.")
            return

        sel = [
            it for it in self._gfx_view.scene().selectedItems()
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def
        ]
        if not sel:
            return

        item = sel[0]
        if item.field_def.outside_stamp:
            self.statusBar().showMessage(
                "Поля вне штампа не привязываются к линиям каркаса", 3000,
            )
            return
        fields = self._collect_field_defs_from_scene([item])
        bound_fields = auto_bind_fields_to_lines(fields, self._grid_lines)
        snapped_fields = snap_fields_to_bindings(bound_fields, self._grid_lines)
        new_fd = snapped_fields[0]
        item.assign_field(new_fd)
        self._snap_item_to_field_def(item, new_fd)

        self._set_modified(True)
        self._sync_wireframe_field_selection()
        self._refresh_wireframe_canvas()
        self.statusBar().showMessage(f"Привязано и подогнано: {item.field_def.id}", 3000)

    def _on_field_bindings_changed(self) -> None:
        """User manually changed a combo in FieldBindingsPanel — snap geometry."""
        from .cell_items import FieldRectItem
        from pdf_parsing_v2_engine.grid_lines_utils import snap_fields_to_bindings
        from dataclasses import replace as dc_replace

        sel = [
            it for it in self._gfx_view.scene().selectedItems()
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def
        ]
        if not sel:
            return
        item = sel[0]
        if item.field_def.outside_stamp:
            return
        bt, bb, bl, br = self._field_bindings_panel.get_bindings()
        fd_with_bindings = dc_replace(
            item.field_def,
            bound_top=bt, bound_bottom=bb,
            bound_left=bl, bound_right=br,
        )
        snapped = snap_fields_to_bindings([fd_with_bindings], self._grid_lines)
        new_fd = snapped[0]
        item.assign_field(new_fd)
        self._snap_item_to_field_def(item, new_fd)
        self._set_modified(True)
        self._refresh_wireframe_canvas()

    def _sync_wireframe_field_selection(self) -> None:
        """Update panels when scene selection changes in wireframe mode."""
        if not self.act_wireframe_mode.isChecked():
            return
        if self._pick_line_which:
            self._on_pick_line_scene_click()
            return

        from .cell_items import FieldRectItem
        from .grid_line_items import GridLineGraphicsItem

        all_sel = self._gfx_view.scene().selectedItems()

        grid_sel = [it for it in all_sel if isinstance(it, GridLineGraphicsItem)]
        if grid_sel:
            lid = grid_sel[0].line_id
            self._grid_lines_panel.select_line_by_id(lid)

        field_sel = [
            it for it in all_sel
            if isinstance(it, FieldRectItem) and it.is_assigned and it.field_def
        ]
        if field_sel:
            self._field_bindings_panel.set_field(field_sel[0].field_def)
        else:
            self._field_bindings_panel.set_field(None)

    # ---- pick-line mode ----

    def _on_pick_line_requested(self, which: str) -> None:
        """Enter pick-line mode: user will click a grid line on canvas to select it."""
        from .cell_items import FieldRectItem
        self._pick_line_which = which
        for it in self._gfx_view.scene().items():
            if isinstance(it, FieldRectItem):
                if it.isVisible():
                    self._pick_line_hidden_items.append(it)
                    it.setVisible(False)
        self._gfx_view.viewport().setCursor(Qt.CursorShape.CrossCursor)
        self.statusBar().showMessage(
            f"Кликните по линии для выбора {'начала' if which == 'start' else 'конца'} пересечения. Esc — отмена.",
            0,
        )

    def _on_pick_line_scene_click(self) -> None:
        """Handle scene selection change while in pick-line mode."""
        from .grid_line_items import GridLineGraphicsItem
        if not self._pick_line_which:
            return
        sel = [
            it for it in self._gfx_view.scene().selectedItems()
            if isinstance(it, GridLineGraphicsItem)
        ]
        if not sel:
            return
        picked_id = sel[0].line_id
        which = self._pick_line_which
        self._exit_pick_line_mode()
        self._grid_lines_panel.set_picked_line(which, picked_id)
        self.statusBar().showMessage(
            f"{'Начало' if which == 'start' else 'Конец'} → {picked_id}", 3000,
        )

    def _exit_pick_line_mode(self) -> None:
        """Leave pick-line mode, restore hidden field items."""
        self._pick_line_which = ""
        vis_fields = self._canvas_stamp_fields_wanted_visible()
        for it in self._pick_line_hidden_items:
            it.setVisible(vis_fields)
        self._pick_line_hidden_items.clear()
        self._gfx_view.viewport().unsetCursor()
        self.statusBar().clearMessage()

    def _on_auto_line_refs(self) -> None:
        """Auto-detect start_line_id/end_line_id and snap endpoints to refs.

        Assigns start_line_id/end_line_id to the nearest covering perpendicular
        line (no distance limit — cover check is the semantic guard).
        Then snap_endpoints_to_refs extends start_mm/end_mm to match the
        pos_mm of the referenced lines ("дотягивание").

        pos_mm is never touched — the line's perpendicular position stays fixed.
        """
        from pdf_parsing_v2_engine.grid_lines_utils import (
            auto_assign_line_refs,
            snap_endpoints_to_refs,
        )

        if not self._grid_lines:
            return

        self._grid_lines = snap_endpoints_to_refs(auto_assign_line_refs(self._grid_lines))
        self._grid_lines_panel.set_lines(self._grid_lines)
        self._field_bindings_panel.set_lines(self._grid_lines)
        self._set_modified(True)
        self._refresh_wireframe_canvas()
        self._snap_bound_fields_to_grid()
        count = sum(1 for gl in self._grid_lines if gl.start_line_id or gl.end_line_id)
        self.statusBar().showMessage(f"Пересечения: {count} линий с привязками", 4000)

    # ---- wireframe canvas helpers (grid line overlay) ----

    def _apply_template_grid_wireframe_visibility(self) -> None:
        vis = (
            self._prealign_overlay_panel.is_canvas_layer_visible("template_grid")
            if self._prealign_overlay_panel
            else True
        )
        for item in self._wireframe_items:
            item.setVisible(vis)

    def _refresh_wireframe_canvas(self) -> None:
        """Rebuild all GridLineGraphicsItem objects on the canvas."""
        self._clear_wireframe_canvas()
        if not self.act_wireframe_mode.isChecked():
            return
        if not self._frame_info or not self._grid_lines:
            return

        from pdf_parsing_v2_engine.coord_transform import grid_line_to_fitz_pts
        from .grid_line_items import GridLineGraphicsItem

        frame = self._frame_info
        origin = "frame_bottom_right"
        dpi_scale = self._gfx_view._dpi / 72.0

        sel_field_bindings: set[str] = set()
        from .cell_items import FieldRectItem
        for it in self._gfx_view.scene().selectedItems():
            if isinstance(it, FieldRectItem) and it.field_def:
                fd = it.field_def
                for b in (fd.bound_top, fd.bound_bottom, fd.bound_left, fd.bound_right):
                    if b:
                        sel_field_bindings.add(b)

        selected_line_id = None
        sel_line_refs: set[str] = set()
        if hasattr(self, '_grid_lines_panel') and self._grid_lines_panel:
            selected_line_id = self._grid_lines_panel.selected_line_id()
        if selected_line_id:
            for gl in self._grid_lines:
                if gl.id == selected_line_id:
                    if gl.start_line_id:
                        sel_line_refs.add(gl.start_line_id)
                    if gl.end_line_id:
                        sel_line_refs.add(gl.end_line_id)
                    break

        highlight_ids = sel_field_bindings | sel_line_refs

        snapped = getattr(self, '_snapped_line_positions', None) or {}
        sx, sy, dxf, dyf = getattr(self, '_adapt_transform', (1.0, 1.0, 0.0, 0.0))

        from pdf_parsing_v2_engine.grid_lines_utils import resolve_line_endpoints_from_snapped
        resolved_spans: dict[str, tuple[float | None, float | None]] = {}
        if snapped:
            resolved_spans = resolve_line_endpoints_from_snapped(
                self._grid_lines, snapped,
            )

        for gl in self._grid_lines:
            try:
                orient, pos_pts, start_pts, end_pts, boundary, lid = grid_line_to_fitz_pts(
                    gl, frame, origin,
                )
            except Exception:
                continue

            snapped_pos = snapped.get(lid)

            if snapped:
                ref_start, ref_end = resolved_spans.get(lid, (None, None))
                if orient == "h":
                    sp = ref_start if ref_start is not None else (start_pts * sx + dxf)
                    ep = ref_end if ref_end is not None else (end_pts * sx + dxf)
                    x1s = min(sp, ep) * dpi_scale
                    x2s = max(sp, ep) * dpi_scale
                    ys = (snapped_pos if snapped_pos is not None else pos_pts) * dpi_scale
                    line_item = GridLineGraphicsItem(
                        lid, orient, boundary, x1s, ys, x2s, ys,
                    )
                else:
                    sp = ref_start if ref_start is not None else (start_pts * sy + dyf)
                    ep = ref_end if ref_end is not None else (end_pts * sy + dyf)
                    y1s = min(sp, ep) * dpi_scale
                    y2s = max(sp, ep) * dpi_scale
                    xs = (snapped_pos if snapped_pos is not None else pos_pts) * dpi_scale
                    line_item = GridLineGraphicsItem(
                        lid, orient, boundary, xs, y1s, xs, y2s,
                    )
            else:
                if orient == "h":
                    x1s = start_pts * dpi_scale
                    x2s = end_pts * dpi_scale
                    ys = pos_pts * dpi_scale
                    line_item = GridLineGraphicsItem(
                        lid, orient, boundary, x1s, ys, x2s, ys,
                    )
                else:
                    y1s = start_pts * dpi_scale
                    y2s = end_pts * dpi_scale
                    xs = pos_pts * dpi_scale
                    line_item = GridLineGraphicsItem(
                        lid, orient, boundary, xs, y1s, xs, y2s,
                    )

            if lid in highlight_ids:
                line_item.set_highlighted(True)
            if lid == selected_line_id:
                line_item.setSelected(True)

            line_item.geo_changed_cb = self._on_grid_line_geo_changed
            self._gfx_view.scene().addItem(line_item)
            self._wireframe_items.append(line_item)
        self._apply_template_grid_wireframe_visibility()

    def _draw_wireframe_on_canvas(self) -> None:
        """Draw wireframe lines unconditionally (ignoring wireframe-mode toggle).

        Used after Тест-адаптация to always show the snapped wireframe so
        the user can see where the template grid landed on the PDF.
        Span coordinates (start/end) are scaled to match the PDF stamp size
        using intersection refs or the global adaptation transform.
        """
        self._clear_wireframe_canvas()
        if not self._frame_info or not self._grid_lines:
            return

        from pdf_parsing_v2_engine.coord_transform import grid_line_to_fitz_pts
        from pdf_parsing_v2_engine.grid_lines_utils import resolve_line_endpoints_from_snapped
        from .grid_line_items import GridLineGraphicsItem

        frame = self._frame_info
        origin = "frame_bottom_right"
        dpi_scale = self._gfx_view._dpi / 72.0
        snapped = self._snapped_line_positions or {}
        sx, sy, dx, dy = self._adapt_transform

        resolved_spans: dict[str, tuple[float | None, float | None]] = {}
        if snapped:
            resolved_spans = resolve_line_endpoints_from_snapped(
                self._grid_lines, snapped,
            )

        for gl in self._grid_lines:
            try:
                orient, pos_pts, start_pts, end_pts, boundary, lid = grid_line_to_fitz_pts(
                    gl, frame, origin,
                )
            except Exception:
                continue

            snapped_pos = snapped.get(lid)

            if snapped:
                ref_start, ref_end = resolved_spans.get(lid, (None, None))
                if orient == "h":
                    sp = ref_start if ref_start is not None else (start_pts * sx + dx)
                    ep = ref_end if ref_end is not None else (end_pts * sx + dx)
                    x1s = min(sp, ep) * dpi_scale
                    x2s = max(sp, ep) * dpi_scale
                    ys = (snapped_pos if snapped_pos is not None else pos_pts) * dpi_scale
                    line_item = GridLineGraphicsItem(
                        lid, orient, boundary, x1s, ys, x2s, ys,
                    )
                else:
                    sp = ref_start if ref_start is not None else (start_pts * sy + dy)
                    ep = ref_end if ref_end is not None else (end_pts * sy + dy)
                    y1s = min(sp, ep) * dpi_scale
                    y2s = max(sp, ep) * dpi_scale
                    xs = (snapped_pos if snapped_pos is not None else pos_pts) * dpi_scale
                    line_item = GridLineGraphicsItem(
                        lid, orient, boundary, xs, y1s, xs, y2s,
                    )
            else:
                if orient == "h":
                    x1s = start_pts * dpi_scale
                    x2s = end_pts * dpi_scale
                    ys = pos_pts * dpi_scale
                    line_item = GridLineGraphicsItem(
                        lid, orient, boundary, x1s, ys, x2s, ys,
                    )
                else:
                    y1s = start_pts * dpi_scale
                    y2s = end_pts * dpi_scale
                    xs = pos_pts * dpi_scale
                    line_item = GridLineGraphicsItem(
                        lid, orient, boundary, xs, y1s, xs, y2s,
                    )

            line_item.geo_changed_cb = self._on_grid_line_geo_changed
            self._gfx_view.scene().addItem(line_item)
            self._wireframe_items.append(line_item)
        self._apply_template_grid_wireframe_visibility()

    def _clear_wireframe_canvas(self) -> None:
        for item in self._wireframe_items:
            self._gfx_view.scene().removeItem(item)
        self._wireframe_items.clear()

    def _highlight_grid_line(self, line_id: str) -> None:
        """Select the matching GridLineGraphicsItem and highlight its ref lines."""
        from .grid_line_items import GridLineGraphicsItem

        ref_ids: set[str] = set()
        for gl in self._grid_lines:
            if gl.id == line_id:
                if gl.start_line_id:
                    ref_ids.add(gl.start_line_id)
                if gl.end_line_id:
                    ref_ids.add(gl.end_line_id)
                break

        for item in self._wireframe_items:
            if isinstance(item, GridLineGraphicsItem):
                item.setSelected(item.line_id == line_id)
                item.set_highlighted(item.line_id in ref_ids)

    def _on_grid_line_geo_changed(
        self, item, commit: bool = True, *, reassign_refs: bool = False,
    ) -> None:
        """A GridLineGraphicsItem was moved/resized on the canvas → update model.

        *commit* False: during line-body drag — update only this line's mm + bound
        fields (no full canvas rebuild, so drag stays smooth).
        *commit* True: refresh dependent spans and canvas. If *reassign_refs*
        (endpoint handle released), run ``auto_assign_line_refs`` so the moved
        end picks the nearest covering perpendicular line, then ``snap_endpoints_to_refs``.
        """
        from pdf_parsing_v2_engine.coord_transform import fitz_pts_to_grid_line_mm
        from .grid_line_items import GridLineGraphicsItem
        from dataclasses import replace as dc_replace

        if not isinstance(item, GridLineGraphicsItem):
            return
        if not self._frame_info:
            return

        dpi_scale = self._gfx_view._dpi / 72.0
        sx1, sy1, sx2, sy2 = item.scene_endpoints()
        fx1, fy1 = sx1 / dpi_scale, sy1 / dpi_scale
        fx2, fy2 = sx2 / dpi_scale, sy2 / dpi_scale

        if item.orientation == "h":
            pos_fitz = fy1
            start_fitz = fx1
            end_fitz = fx2
        else:
            pos_fitz = fx1
            start_fitz = fy1
            end_fitz = fy2

        pos_mm, start_mm, end_mm = fitz_pts_to_grid_line_mm(
            item.orientation, pos_fitz, start_fitz, end_fitz,
            self._frame_info, "frame_bottom_right",
        )

        for i, gl in enumerate(self._grid_lines):
            if gl.id == item.line_id:
                self._grid_lines[i] = dc_replace(
                    gl, pos_mm=pos_mm, start_mm=start_mm, end_mm=end_mm,
                )
                break

        self._set_modified(True)
        if commit:
            from pdf_parsing_v2_engine.grid_lines_utils import (
                auto_assign_line_refs,
                snap_endpoints_to_refs,
            )

            if self._grid_lines_panel.line_intersection_snap_enabled():
                if reassign_refs:
                    self._grid_lines = snap_endpoints_to_refs(
                        auto_assign_line_refs(self._grid_lines),
                    )
                else:
                    self._grid_lines = snap_endpoints_to_refs(self._grid_lines)
            self._grid_lines_panel.set_lines(self._grid_lines)
            self._field_bindings_panel.set_lines(self._grid_lines)
            self._refresh_wireframe_canvas()
        self._snap_bound_fields_to_grid()

    def keyPressEvent(self, event) -> None:
        if event.key() == Qt.Key.Key_Escape and self._pick_line_which:
            self._exit_pick_line_mode()
            event.accept()
            return
        super().keyPressEvent(event)

    # ---- startup ----
    def showEvent(self, event):
        super().showEvent(event)
        if self._cells_panel is None:
            self._wire_panels()
            self._try_auto_load_catalog()
            self._update_adapt_action_state()

    def closeEvent(self, event):
        if self._is_modified or self._catalog_is_dirty:
            r = QMessageBox.question(
                self,
                "Несохранённые изменения",
                "Есть несохранённые изменения (шаблон и/или каталог полей).\n"
                "Выйти без сохранения?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
        if self._prealign_is_floating and self._prealign_overlay_panel:
            self._dock_prealign_panel(persist_floating_pref=False)
        if self._catalog_is_floating and self._catalog_panel:
            self._dock_catalog_panel()
        if self._fitz_doc:
            self._fitz_doc.close()
            self._fitz_doc = None
        super().closeEvent(event)
