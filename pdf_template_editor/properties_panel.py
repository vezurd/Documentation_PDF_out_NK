"""Панель свойств выбранной ячейки: id/clean/regex/expected/тип/origin и др."""

from __future__ import annotations

from dataclasses import replace
from typing import TYPE_CHECKING

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QWidget,
    QFormLayout,
    QComboBox,
    QLineEdit,
    QDoubleSpinBox,
    QLabel,
    QGroupBox,
    QVBoxLayout,
    QCheckBox,
    QPushButton,
    QHBoxLayout,
    QSizePolicy,
)

from pdf_parsing_v2_engine.models import FieldDef

if TYPE_CHECKING:
    from pdf_parsing_v2_engine.models import FieldCatalog
    from .cell_items import FieldRectItem
    from .graphics_view import StampGraphicsView

_NONE_LABEL = "— не назначена —"
_NONE_CLEAN = "(без очистки)"


class PropertiesPanel(QWidget):
    """Properties editor for the currently selected FieldRectItem."""

    field_changed = Signal(object)  # emits the FieldRectItem
    document_field_changed = Signal(object)  # emits updated FieldDef (document property)
    edit_field_type_requested = Signal(str)  # emits field_id

    def __init__(
        self,
        catalog: FieldCatalog | None = None,
        gfx_view: StampGraphicsView | None = None,
        parent=None,
    ):
        super().__init__(parent)
        self._catalog = catalog
        self._gfx_view = gfx_view
        self._current_item: FieldRectItem | None = None
        self._current_doc_fd: FieldDef | None = None
        self._updating: bool = False

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        grp = QGroupBox("Свойства поля")
        grp.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        outer.addWidget(grp)
        form = QFormLayout(grp)

        self._lbl_doc_mode = QLabel(
            "Служебное поле свойств PDF — не отображается на холсте; значение из метаданных."
        )
        self._lbl_doc_mode.setWordWrap(True)
        self._lbl_doc_mode.setStyleSheet("color: #6a4a8a; font-weight: bold;")
        self._lbl_doc_mode.hide()
        form.addRow(self._lbl_doc_mode)

        self._cmb_id = QComboBox()
        self._cmb_id.setEditable(True)
        self._cmb_id.currentTextChanged.connect(self._on_id_changed)
        form.addRow("Поле:", self._cmb_id)

        self._cmb_clean = QComboBox()
        self._cmb_clean.setEditable(True)
        self._cmb_clean.currentTextChanged.connect(self._refresh_clean_combo_appearance)
        self._cmb_clean.currentTextChanged.connect(self._apply)
        form.addRow("Clean:", self._cmb_clean)

        self._txt_regex = QLineEdit()
        self._txt_regex.setPlaceholderText(r"Напр.: ^\d{2}\.\d{2}\.\d{4}$")
        self._txt_regex.setToolTip(
            "Проверка значения после очистки (clean) на этапе extract/test.\n"
            "Используется как re.search(validate_regex, cleaned_value)."
        )
        self._txt_regex.editingFinished.connect(self._apply)
        form.addRow("Regex:", self._txt_regex)

        self._cmb_expected = QComboBox()
        self._cmb_expected.addItems(["required", "optional", "absent"])
        self._cmb_expected.currentTextChanged.connect(self._apply)
        form.addRow("Expected:", self._cmb_expected)

        self._lbl_field_type = QLabel("data")
        self._btn_edit_field_type = QPushButton("Изменить в каталоге…")
        self._btn_edit_field_type.clicked.connect(self._on_edit_field_type_clicked)
        field_type_row = QWidget()
        field_type_layout = QHBoxLayout(field_type_row)
        field_type_layout.setContentsMargins(0, 0, 0, 0)
        field_type_layout.addWidget(self._lbl_field_type, 1)
        field_type_layout.addWidget(self._btn_edit_field_type)
        form.addRow("Тип поля:", field_type_row)

        self._txt_expected_text = QLineEdit()
        self._txt_expected_text.setPlaceholderText("Ожидаемый текст (для label)")
        self._txt_expected_text.editingFinished.connect(self._apply)
        form.addRow("Expected text:", self._txt_expected_text)

        self._chk_is_anchor = QCheckBox("Якорь для grid adaptation")
        self._chk_is_anchor.stateChanged.connect(self._apply)
        form.addRow("Anchor:", self._chk_is_anchor)

        self._chk_outside_stamp = QCheckBox("Поле вне штампа (исключить из подгонки)")
        self._chk_outside_stamp.setToolTip(
            "Исключает поле из привязки к линиям каркаса штампа.\n"
            "Вертикаль: внутренняя горизонталь к рамке чертежа подстраивается автоматически "
            "при загрузке шаблона и в pipeline (тип колонтитула — по Origin: верх/низ). "
            "Внешний край — Stretch к краю листа; горизонталь — bbox_mm и stretch left/right."
        )
        self._chk_outside_stamp.stateChanged.connect(self._apply)
        form.addRow("Вне штампа:", self._chk_outside_stamp)

        self._cmb_origin = QComboBox()
        self._cmb_origin.addItems([
            "frame_bottom_right", "frame_top_right",
            "frame_bottom_left", "frame_top_left",
        ])
        self._cmb_origin.setToolTip(
            "Базовый угол рамки чертежа, от которого задаются bbox_mm (h1,v1,h2,v2).\n"
            "Для полей вне штампа: верх/низ в названии origin задаёт тип колонтитула; "
            "внутренняя горизонталь к рамке выравнивается автоматически; внешний край листа — Stretch.\n"
            "Смена origin перепривязывает те же мм к другому углу (после drag учитывается размер)."
        )
        # currentIndexChanged: one emission per user pick (avoids duplicate _apply vs text).
        self._cmb_origin.currentIndexChanged.connect(self._apply)
        form.addRow("Origin:", self._cmb_origin)

        stretch_row = QWidget()
        stretch_layout = QHBoxLayout(stretch_row)
        stretch_layout.setContentsMargins(0, 0, 0, 0)
        self._chk_stretch_bottom = QCheckBox("bottom")
        self._chk_stretch_top = QCheckBox("top")
        self._chk_stretch_left = QCheckBox("left")
        self._chk_stretch_right = QCheckBox("right")
        for chk in (self._chk_stretch_bottom, self._chk_stretch_top,
                     self._chk_stretch_left, self._chk_stretch_right):
            chk.stateChanged.connect(self._apply)
            stretch_layout.addWidget(chk)
        self._lbl_stretch = QLabel("Stretch:")
        self._lbl_stretch.setToolTip(
            "Прижатие внешнего края к краю листа (pdfminer/fitz). "
            "Для полей вне штампа: вертикально — нижний/верхний край к низу/верху листа; "
            "горизонтально — left/right к соответствующему краю. Рамка чертежа — через Origin "
            "(внутренняя горизонталь к линии рамки автоматически при загрузке)."
        )
        form.addRow(self._lbl_stretch, stretch_row)

        self._spn_padding = QDoubleSpinBox()
        self._spn_padding.setRange(-1.0, 10.0)
        self._spn_padding.setDecimals(1)
        self._spn_padding.setSpecialValueText("(шаблон)")
        self._spn_padding.setValue(-1.0)
        self._spn_padding.valueChanged.connect(self._apply)
        form.addRow("Padding мм:", self._spn_padding)

        self._lbl_bbox = QLabel("—")
        form.addRow("bbox_mm:", self._lbl_bbox)

        # Rows toggled off for document-property fields (see _set_doc_property_form_visible).
        self._geom_form_rows: list[tuple[str, QWidget]] = [
            ("Clean:", self._cmb_clean),
            ("Regex:", self._txt_regex),
            ("Тип поля:", field_type_row),
            ("Anchor:", self._chk_is_anchor),
            ("Вне штампа:", self._chk_outside_stamp),
            ("Origin:", self._cmb_origin),
            ("Stretch:", stretch_row),
            ("Padding мм:", self._spn_padding),
            ("bbox_mm:", self._lbl_bbox),
        ]

        self._populate_combos()
        outer.addStretch(1)

    def set_catalog(self, catalog: FieldCatalog | None) -> None:
        self._catalog = catalog
        self._populate_combos()
        # После перестройки комбо восстановить отображение текущего поля,
        # иначе выпадающий список «Поле» сбрасывается в «не назначен».
        self.refresh_from_current_item()

    def set_gfx_view(self, gfx_view: StampGraphicsView | None) -> None:
        self._gfx_view = gfx_view

    def refresh_from_current_item(self) -> None:
        """Перезагрузить поля из текущего/выбранного поля (после синхронизации с каталогом и т.п.)."""
        from .cell_items import FieldRectItem

        if self._current_doc_fd is not None:
            self.load_from_document_property_field(self._current_doc_fd)
            return
        if self._current_item is not None:
            self.load_from_item(self._current_item)
            return
        if self._gfx_view is None:
            return
        sel = [
            s for s in self._gfx_view.scene().selectedItems()
            if isinstance(s, FieldRectItem)
        ]
        if len(sel) == 1:
            self.load_from_item(sel[0])

    def load_from_item(self, item: FieldRectItem) -> None:
        self._updating = True
        self._current_doc_fd = None
        self._current_item = item
        self._set_doc_property_form_visible(False)
        fd = item.field_def
        if fd:
            idx = self._find_combo_index_by_id(fd.id)
            if idx >= 0:
                self._cmb_id.setCurrentIndex(idx)
            else:
                self._cmb_id.setCurrentText(fd.id)
            cat_entry = self._catalog.get_entry(fd.id) if self._catalog else None
            if cat_entry:
                self._lbl_field_type.setText(cat_entry.default_field_type or "data")
            else:
                ft = fd.field_type
                if ft in ("static", "static_empty"):
                    from pdf_parsing_v2_engine.models import normalize_field_type

                    ft = normalize_field_type(ft)
                self._lbl_field_type.setText(ft or "data")
            clean_text = fd.clean or _NONE_CLEAN
            idx_c = self._cmb_clean.findText(clean_text)
            if idx_c >= 0:
                self._cmb_clean.setCurrentIndex(idx_c)
            else:
                self._cmb_clean.setCurrentText(clean_text)
            self._txt_regex.setText(fd.validate_regex or "")
            self._cmb_expected.setCurrentText(fd.expected)
            self._txt_expected_text.setText(fd.expected_text or "")
            self._chk_is_anchor.setChecked(bool(fd.is_anchor))
            self._chk_outside_stamp.setChecked(bool(getattr(fd, "outside_stamp", False)))
            self._cmb_origin.setCurrentText(fd.origin)
            stretch = set(fd.stretch_to_page)
            self._chk_stretch_bottom.setChecked("bottom" in stretch)
            self._chk_stretch_top.setChecked("top" in stretch)
            self._chk_stretch_left.setChecked("left" in stretch)
            self._chk_stretch_right.setChecked("right" in stretch)
            self._spn_padding.setValue(fd.padding_mm if fd.padding_mm is not None else -1.0)
        else:
            self._cmb_id.setCurrentText(_NONE_LABEL)
            self._lbl_field_type.setText("data")
            self._cmb_clean.setCurrentText(_NONE_CLEAN)
            self._txt_regex.clear()
            self._cmb_expected.setCurrentText("optional")
            self._txt_expected_text.clear()
            self._chk_is_anchor.setChecked(False)
            self._chk_outside_stamp.setChecked(False)
            self._cmb_origin.setCurrentText("frame_bottom_right")
            self._chk_stretch_bottom.setChecked(False)
            self._chk_stretch_top.setChecked(False)
            self._chk_stretch_left.setChecked(False)
            self._chk_stretch_right.setChecked(False)
            self._spn_padding.setValue(-1.0)
        self._updating = False
        self._cmb_id.setEnabled(True)
        self._toggle_expected_text_enabled()
        self._refresh_clean_combo_appearance()
        # Update bbox display from actual scene position
        self._refresh_bbox_display(item)

    def load_from_document_property_field(self, fd: FieldDef) -> None:
        """Edit a hidden PDF document-property field (no canvas item)."""
        from pdf_parsing_v2_engine.document_properties.registry import MANDATORY_PROPERTY_KEYS

        self._updating = True
        self._current_item = None
        self._current_doc_fd = fd
        self._set_doc_property_form_visible(True)
        self._cmb_id.setEnabled(False)
        idx = self._find_combo_index_by_id(fd.id)
        if idx >= 0:
            self._cmb_id.setCurrentIndex(idx)
        else:
            self._cmb_id.setCurrentText(fd.id)
        prop_key = fd.document_property or fd.id
        if prop_key not in MANDATORY_PROPERTY_KEYS:
            self._lbl_doc_mode.setText(
                "Служебное поле свойств PDF — неизвестный ключ document_property "
                f"({prop_key!r}). Сохранение может вызвать ошибки в pipeline."
            )
        else:
            self._lbl_doc_mode.setText(
                "Служебное поле свойств PDF — не отображается на холсте; значение из метаданных."
            )
        self._lbl_field_type.setText("data")
        self._cmb_clean.setCurrentText(_NONE_CLEAN)
        self._txt_regex.clear()
        self._cmb_expected.setCurrentText(fd.expected)
        self._txt_expected_text.setText(fd.expected_text or "")
        self._chk_is_anchor.setChecked(False)
        self._chk_outside_stamp.setChecked(False)
        self._cmb_origin.setCurrentText(fd.origin)
        for chk in (
            self._chk_stretch_bottom, self._chk_stretch_top,
            self._chk_stretch_left, self._chk_stretch_right,
        ):
            chk.setChecked(False)
        self._spn_padding.setValue(-1.0)
        self._lbl_bbox.setText("— (не на холсте)")
        self._updating = False
        self._toggle_expected_text_enabled()
        self._refresh_clean_combo_appearance()

    def refresh_bbox_from_item(self, item: FieldRectItem) -> None:
        """Called when item is dragged/resized to update the bbox_mm label."""
        if item is not self._current_item:
            return
        self._refresh_bbox_display(item)

    # ---- internals ----

    def _refresh_clean_combo_appearance(self, *_args) -> None:
        """Highlight unknown ``clean`` keys; tooltips from ``CLEANER_DESCRIPTIONS``."""
        le = self._cmb_clean.lineEdit()
        if le is None:
            return
        try:
            from pdf_parsing_v2_engine.field_cleaners import CLEANER_DESCRIPTIONS, CLEANERS
        except ImportError:
            le.setStyleSheet("")
            le.setToolTip("")
            return
        raw = self._cmb_clean.currentText().strip()
        if raw == _NONE_CLEAN or not raw:
            le.setStyleSheet("")
            le.setToolTip("")
            return
        if raw not in CLEANERS:
            le.setStyleSheet("background-color: #fff9c4;")
            le.setToolTip(
                "This clean key is not in the v2 registry — extraction will emit "
                "UNKNOWN_CLEANER and pass through raw text."
            )
            return
        le.setStyleSheet("")
        le.setToolTip(CLEANER_DESCRIPTIONS.get(raw, ""))

    def _refresh_bbox_display(self, item: FieldRectItem) -> None:
        bbox = self._compute_bbox_mm(item)
        if bbox is not None:
            self._lbl_bbox.setText(
                f"[{bbox[0]:.1f}, {bbox[1]:.1f}, {bbox[2]:.1f}, {bbox[3]:.1f}]"
            )
        else:
            fd = item.field_def
            if fd:
                b = fd.bbox_mm
                self._lbl_bbox.setText(
                    f"[{b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}]"
                )
            else:
                self._lbl_bbox.setText("—")

    def _pdfminer_rect_from_item(
        self, item: FieldRectItem,
    ) -> tuple[float, float, float, float] | None:
        """Scene rect → pdfminer (y-up) absolute pts; same convention as save/drag."""
        if self._gfx_view is None or self._gfx_view.frame_info is None:
            return None
        frame = self._gfx_view.frame_info
        pos = item.scenePos()
        r = item.rect()
        sx, sy, sw, sh = pos.x() + r.x(), pos.y() + r.y(), r.width(), r.height()
        dpi_scale = self._gfx_view._dpi / 72.0
        fitz_x0 = sx / dpi_scale
        fitz_y0 = sy / dpi_scale
        fitz_x1 = (sx + sw) / dpi_scale
        fitz_y1 = (sy + sh) / dpi_scale
        pm_x0 = fitz_x0
        pm_y0 = frame.page_height - fitz_y1
        pm_x1 = fitz_x1
        pm_y1 = frame.page_height - fitz_y0
        return (pm_x0, pm_y0, pm_x1, pm_y1)

    def _compute_bbox_mm(
        self, item: FieldRectItem, origin: str | None = None,
    ) -> tuple[float, float, float, float] | None:
        """Compute bbox_mm from item scene rect, relative to *origin* (ignores stretch)."""
        if self._gfx_view is None or self._gfx_view.frame_info is None:
            return None
        from pdf_parsing_v2_engine.coord_transform import absolute_to_field_bbox_mm

        frame = self._gfx_view.frame_info
        if origin is None:
            origin = (item.field_def.origin if item.field_def else "frame_bottom_right")
        pm = self._pdfminer_rect_from_item(item)
        if pm is None:
            return None
        pm_x0, pm_y0, pm_x1, pm_y1 = pm
        return absolute_to_field_bbox_mm(origin, frame, pm_x0, pm_y0, pm_x1, pm_y1)

    def _populate_combos(self) -> None:
        self._updating = True
        self._cmb_id.clear()
        self._cmb_id.addItem(_NONE_LABEL)
        if self._catalog:
            self._populate_id_combo_grouped(self._catalog)

        self._cmb_clean.clear()
        self._cmb_clean.addItem(_NONE_CLEAN)
        try:
            from pdf_parsing_v2_engine.field_cleaners import CLEANER_DESCRIPTIONS, CLEANERS

            for k in sorted(CLEANERS.keys()):
                self._cmb_clean.addItem(k)
                idx = self._cmb_clean.count() - 1
                desc = CLEANER_DESCRIPTIONS.get(k, "")
                if desc:
                    self._cmb_clean.setItemData(idx, desc, Qt.ItemDataRole.ToolTipRole)
        except ImportError:
            pass
        self._updating = False
        self._refresh_clean_combo_appearance()

    def _populate_id_combo_grouped(self, catalog: FieldCatalog) -> None:
        from collections import OrderedDict

        groups: OrderedDict[str, list] = OrderedDict()
        for e in catalog.entries:
            g = e.group or ""
            groups.setdefault(g, []).append(e)

        first_group = True
        for grp_name, entries in groups.items():
            if not first_group:
                self._cmb_id.insertSeparator(self._cmb_id.count())
            first_group = False
            if grp_name:
                idx = self._cmb_id.count()
                self._cmb_id.addItem(f"── {grp_name} ──")
                model = self._cmb_id.model()
                if model is not None:
                    item = model.item(idx)
                    if item is not None:
                        item.setEnabled(False)
            for e in entries:
                display = f"{e.id}  —  {e.label}" if e.label else e.id
                self._cmb_id.addItem(display, userData=e.id)

    def _resolve_id_from_combo(self) -> str:
        """Extract actual field id from current combo selection (strip label part)."""
        data = self._cmb_id.currentData()
        if data:
            return str(data)
        text = self._cmb_id.currentText().strip()
        if "  —  " in text:
            return text.split("  —  ", 1)[0].strip()
        return text

    def _find_combo_index_by_id(self, field_id: str) -> int:
        """Find combo item index by userData (field id) set during grouped populate."""
        for i in range(self._cmb_id.count()):
            if self._cmb_id.itemData(i) == field_id:
                return i
        return self._cmb_id.findText(field_id)

    def _set_doc_property_form_visible(self, doc_mode: bool) -> None:
        self._lbl_doc_mode.setVisible(doc_mode)
        for _label, w in self._geom_form_rows:
            w.setVisible(not doc_mode)

    def _on_id_changed(self, _text: str) -> None:
        if self._updating or self._current_doc_fd is not None:
            return
        if not self._current_item:
            return
        field_id = self._resolve_id_from_combo()
        if not field_id or field_id == _NONE_LABEL:
            self._current_item.unassign()
            self.field_changed.emit(self._current_item)
            return
        if self._catalog:
            entry = self._catalog.get_entry(field_id)
            if entry:
                self._updating = True
                # expected / expected_text — только из шаблона (как clean), не из каталога
                self._lbl_field_type.setText(entry.default_field_type or "data")
                self._updating = False
                self._toggle_expected_text_enabled()
            else:
                current_fd = self._current_item.field_def
                self._lbl_field_type.setText((current_fd.field_type if current_fd else "data") or "data")
        self._apply()

    def _toggle_expected_text_enabled(self) -> None:
        if self._current_doc_fd is not None:
            is_label = self._current_field_type() == "label"
            self._txt_expected_text.setEnabled(is_label)
            return
        is_label = self._current_field_type() == "label"
        self._txt_expected_text.setEnabled(is_label)
        if not is_label and not self._updating:
            self._txt_expected_text.clear()

    def _current_field_type(self) -> str:
        ft = self._lbl_field_type.text().strip()
        return ft or "data"

    def _on_edit_field_type_clicked(self) -> None:
        if self._current_doc_fd is not None:
            return
        field_id = self._resolve_id_from_combo()
        if not field_id or field_id == _NONE_LABEL:
            return
        self.edit_field_type_requested.emit(field_id)

    def _current_stretch_to_page(self) -> tuple[str, ...]:
        edges: list[str] = []
        if self._chk_stretch_bottom.isChecked():
            edges.append("bottom")
        if self._chk_stretch_top.isChecked():
            edges.append("top")
        if self._chk_stretch_left.isChecked():
            edges.append("left")
        if self._chk_stretch_right.isChecked():
            edges.append("right")
        return tuple(edges)

    def _apply(self, *_args) -> None:
        if self._updating:
            return
        if self._current_doc_fd is not None:
            self._apply_document_property_field()
            return
        if not self._current_item:
            return
        field_id = self._resolve_id_from_combo()
        if not field_id or field_id == _NONE_LABEL:
            self._current_item.unassign()
            self.field_changed.emit(self._current_item)
            return

        clean = self._cmb_clean.currentText()
        if clean == _NONE_CLEAN:
            clean = None

        padding = self._spn_padding.value()
        padding_mm = None if padding < 0 else padding

        regex = self._txt_regex.text().strip() or None

        origin = self._cmb_origin.currentText() or "frame_bottom_right"
        stretch = self._current_stretch_to_page()

        prev = self._current_item.field_def
        frame = self._gfx_view.frame_info if self._gfx_view else None

        # Changing origin: re-encode the *current* scene rect in the previous origin's
        # mm, then store those four numbers with the new origin (re-anchor to another
        # frame corner). With stretch_to_page, the scene rect is post-stretch — undo
        # clamp using resolve_field_bbox_core(prev) before absolute_to_field_bbox_mm.
        if prev and prev.origin != origin and frame is not None:
            if prev.stretch_to_page:
                from pdf_parsing_v2_engine.coord_transform import (
                    absolute_to_field_bbox_mm,
                    pdfminer_rect_restore_before_stretch,
                    resolve_field_bbox_core,
                )
                pm_scene = self._pdfminer_rect_from_item(self._current_item)
                if pm_scene is not None:
                    core = resolve_field_bbox_core(prev, frame)
                    pm_un = pdfminer_rect_restore_before_stretch(
                        pm_scene, prev.stretch_to_page, core,
                    )
                    carried = absolute_to_field_bbox_mm(
                        prev.origin, frame, *pm_un,
                    )
                else:
                    carried = None
                bbox_mm = carried if carried is not None else prev.bbox_mm
            else:
                carried = self._compute_bbox_mm(self._current_item, origin=prev.origin)
                bbox_mm = carried if carried is not None else prev.bbox_mm
        else:
            bbox_mm = (
                self._compute_bbox_mm(self._current_item, origin=origin)
                or (prev.bbox_mm if prev else (0.0, 0.0, 0.0, 0.0))
            )
        self._lbl_bbox.setText(
            f"[{bbox_mm[0]:.1f}, {bbox_mm[1]:.1f}, {bbox_mm[2]:.1f}, {bbox_mm[3]:.1f}]"
        )

        entry = self._catalog.get_entry(field_id) if self._catalog else None
        if entry:
            label = entry.label
            field_type = entry.default_field_type or "data"
        else:
            current_fd = self._current_item.field_def
            label = (current_fd.label or "").strip() if current_fd else ""
            if not label:
                label = field_id
            field_type = self._current_field_type()
        if self._chk_outside_stamp.isChecked():
            bt = bb = bl = br = None
        elif prev:
            bt, bb, bl, br = prev.bound_top, prev.bound_bottom, prev.bound_left, prev.bound_right
        else:
            bt = bb = bl = br = None

        from pdf_parsing_v2_engine.models import FieldDef
        fd = FieldDef(
            id=field_id,
            label=label,
            bbox_mm=bbox_mm,
            clean=clean,
            validate_regex=regex,
            expected=self._cmb_expected.currentText(),
            padding_mm=padding_mm,
            field_type=field_type,
            expected_text=(self._txt_expected_text.text().strip() or None),
            is_anchor=self._chk_is_anchor.isChecked(),
            outside_stamp=self._chk_outside_stamp.isChecked(),
            origin=origin,
            stretch_to_page=stretch,
            bound_top=bt,
            bound_bottom=bb,
            bound_left=bl,
            bound_right=br,
        )
        self._current_item.assign_field(fd)
        self.field_changed.emit(self._current_item)

    def _apply_document_property_field(self) -> None:
        fd0 = self._current_doc_fd
        if fd0 is None:
            return
        key = fd0.document_property or fd0.id
        new_fd = replace(
            fd0,
            expected=self._cmb_expected.currentText(),
            expected_text=(self._txt_expected_text.text().strip() or None),
            document_property=key,
        )
        self._current_doc_fd = new_fd
        self.document_field_changed.emit(new_fd)
