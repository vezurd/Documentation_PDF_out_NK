"""Панель метаданных шаблона (постоянная, в правой панели).

Показывает и позволяет редактировать: имя шаблона, doc_types, page_selector,
priority, padding_mm. Заполняется при загрузке шаблона; значения используются
при сохранении — диалог SaveTemplateDialog больше не нужен.

Содержит кнопку «Сохранить» и индикатор несохранённых изменений «● не сохранён».
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

if TYPE_CHECKING:
    from pdf_parsing_v2_engine.models import StampTemplate

_ALL_DOC_TYPES = [
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
]
_DEFAULT_DOC_TYPES = {"WIR", "LAY", "CAE", "GA", "PL", "DW", "NI"}


class TemplateMetaPanel(QWidget):
    """Persistent widget for template-level metadata editing."""

    meta_changed = Signal()
    save_requested = Signal()   # emitted when the Save button is clicked

    def __init__(self, parent=None):
        super().__init__(parent)
        self._updating = False

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(2)

        grp = QGroupBox("Параметры шаблона")
        outer.addWidget(grp)
        grp_layout = QVBoxLayout(grp)
        grp_layout.setSpacing(4)

        form = QFormLayout()
        form.setVerticalSpacing(4)
        grp_layout.addLayout(form)

        self._txt_name = QLineEdit("template")
        self._txt_name.setPlaceholderText("Имя шаблона")
        self._txt_name.editingFinished.connect(self._on_changed)
        form.addRow("Имя:", self._txt_name)

        # doc_types — two compact rows of checkboxes
        dt_widget = QWidget()
        dt_layout_outer = QVBoxLayout(dt_widget)
        dt_layout_outer.setContentsMargins(0, 0, 0, 0)
        dt_layout_outer.setSpacing(2)
        self._chk_doc_types: dict[str, QCheckBox] = {}
        row1 = QHBoxLayout()
        row2 = QHBoxLayout()
        for i, dt in enumerate(_ALL_DOC_TYPES):
            chk = QCheckBox(dt)
            chk.setChecked(dt in _DEFAULT_DOC_TYPES)
            chk.stateChanged.connect(self._on_changed)
            self._chk_doc_types[dt] = chk
            (row1 if i < 6 else row2).addWidget(chk)
        row1.addStretch()
        row2.addStretch()
        dt_layout_outer.addLayout(row1)
        dt_layout_outer.addLayout(row2)
        form.addRow("Типы:", dt_widget)

        self._cmb_page = QComboBox()
        self._cmb_page.addItems(["first", "rest", "all"])
        self._cmb_page.currentTextChanged.connect(self._on_changed)
        form.addRow("Страница:", self._cmb_page)

        self._spn_priority = QSpinBox()
        self._spn_priority.setRange(1, 100)
        self._spn_priority.setValue(10)
        self._spn_priority.setToolTip("Приоритет шаблона (выше = пробуется первым)")
        self._spn_priority.valueChanged.connect(self._on_changed)
        form.addRow("Priority:", self._spn_priority)

        self._chk_grid_adapt = QCheckBox("Адаптировать под сетку PDF")
        self._chk_grid_adapt.stateChanged.connect(self._on_changed)
        form.addRow("Grid adapt:", self._chk_grid_adapt)

        self._cmb_frame_mode = QComboBox()
        self._cmb_frame_mode.addItem("ГОСТ (get_drawings)", "gost")
        self._cmb_frame_mode.addItem("Свободная (union в ROI)", "drawing_union")
        self._cmb_frame_mode.setToolTip(
            "ГОСТ: как в чертежах (max span + порог 20% или поля ГОСТ). "
            "Свободная: объединение линий get_drawings внутри печатного поля и ROI полей штампа."
        )
        self._cmb_frame_mode.currentIndexChanged.connect(self._on_changed)
        form.addRow("Рамка:", self._cmb_frame_mode)

        self._btn_adv = QToolButton()
        self._btn_adv.setText("Расширенные (pipeline)")
        self._btn_adv.setCheckable(True)
        self._btn_adv.setChecked(False)
        self._btn_adv.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self._btn_adv.setArrowType(Qt.ArrowType.RightArrow)
        self._btn_adv.setAutoRaise(True)
        self._btn_adv.toggled.connect(self._on_adv_toggled)
        form.addRow("", self._btn_adv)

        self._adv_widget = QWidget()
        adv_form = QFormLayout(self._adv_widget)
        adv_form.setVerticalSpacing(4)

        self._spn_padding = QDoubleSpinBox()
        self._spn_padding.setRange(0.0, 10.0)
        self._spn_padding.setDecimals(1)
        self._spn_padding.setValue(0.5)
        self._spn_padding.setToolTip("Допуск в мм при извлечении текста (расширяет bbox)")
        self._spn_padding.valueChanged.connect(self._on_changed)
        adv_form.addRow("Padding мм:", self._spn_padding)

        self._spn_detected_tol = QDoubleSpinBox()
        self._spn_detected_tol.setRange(0.1, 20.0)
        self._spn_detected_tol.setDecimals(1)
        self._spn_detected_tol.setValue(0.5)
        self._spn_detected_tol.setToolTip("Tolerance кластеризации обнаруженных ячеек (мм)")
        self._spn_detected_tol.valueChanged.connect(self._on_changed)
        adv_form.addRow("Detected tol, мм:", self._spn_detected_tol)

        self._spn_grid_tol = QDoubleSpinBox()
        self._spn_grid_tol.setRange(0.1, 20.0)
        self._spn_grid_tol.setDecimals(1)
        self._spn_grid_tol.setValue(2.0)
        self._spn_grid_tol.setToolTip("Tolerance для консолидации сетки шаблона (мм)")
        self._spn_grid_tol.valueChanged.connect(self._on_changed)
        adv_form.addRow("Grid tol, мм:", self._spn_grid_tol)

        self._spn_snap_max_dist = QDoubleSpinBox()
        self._spn_snap_max_dist.setRange(0.1, 50.0)
        self._spn_snap_max_dist.setDecimals(1)
        self._spn_snap_max_dist.setValue(5.0)
        self._spn_snap_max_dist.setToolTip("Максимальная дистанция snap границы к grid-линии (мм)")
        self._spn_snap_max_dist.valueChanged.connect(self._on_changed)
        adv_form.addRow("Snap max dist, мм:", self._spn_snap_max_dist)

        self._spn_max_shape = QDoubleSpinBox()
        self._spn_max_shape.setRange(1.0, 5.0)
        self._spn_max_shape.setDecimals(1)
        self._spn_max_shape.setSingleStep(0.5)
        self._spn_max_shape.setValue(2.5)
        self._spn_max_shape.setToolTip("Макс. допустимое изменение формы (ratio w/h)")
        self._spn_max_shape.valueChanged.connect(self._on_changed)
        adv_form.addRow("Max shape change:", self._spn_max_shape)

        self._spn_min_score = QDoubleSpinBox()
        self._spn_min_score.setRange(0.1, 1.0)
        self._spn_min_score.setDecimals(2)
        self._spn_min_score.setSingleStep(0.05)
        self._spn_min_score.setValue(0.4)
        self._spn_min_score.setToolTip("Мин. score для принятия snap-привязки")
        self._spn_min_score.valueChanged.connect(self._on_changed)
        adv_form.addRow("Min score:", self._spn_min_score)

        self._chk_find_tables_snap = QCheckBox(
            "Свои snap для find_tables (вместо pdf_v2_config)"
        )
        self._chk_find_tables_snap.setToolTip(
            "Если включено, PyMuPDF find_tables использует указанные snap_x/y "
            "вместо find_tables_snap_* из общего конфига."
        )
        self._chk_find_tables_snap.stateChanged.connect(self._on_find_tables_snap_toggled)
        self._chk_find_tables_snap.stateChanged.connect(self._on_changed)
        adv_form.addRow("Find tables:", self._chk_find_tables_snap)

        self._spn_find_tables_snap_x = QDoubleSpinBox()
        self._spn_find_tables_snap_x.setRange(0.1, 50.0)
        self._spn_find_tables_snap_x.setDecimals(2)
        self._spn_find_tables_snap_x.setSingleStep(0.1)
        self._spn_find_tables_snap_x.setValue(2.2)
        self._spn_find_tables_snap_x.setToolTip("snap_x_tolerance для find_tables() (pts)")
        self._spn_find_tables_snap_x.valueChanged.connect(self._on_changed)
        adv_form.addRow("  snap X:", self._spn_find_tables_snap_x)

        self._spn_find_tables_snap_y = QDoubleSpinBox()
        self._spn_find_tables_snap_y.setRange(0.1, 50.0)
        self._spn_find_tables_snap_y.setDecimals(2)
        self._spn_find_tables_snap_y.setSingleStep(0.1)
        self._spn_find_tables_snap_y.setValue(2.0)
        self._spn_find_tables_snap_y.setToolTip("snap_y_tolerance для find_tables() (pts)")
        self._spn_find_tables_snap_y.valueChanged.connect(self._on_changed)
        adv_form.addRow("  snap Y:", self._spn_find_tables_snap_y)

        self._on_find_tables_snap_toggled()

        self._adv_widget.setVisible(False)
        grp_layout.addWidget(self._adv_widget)

        # ---- bottom row: unsaved indicator + save button ----
        bottom = QHBoxLayout()
        bottom.setContentsMargins(0, 2, 0, 0)

        self._lbl_modified = QLabel("● не сохранён")
        self._lbl_modified.setToolTip("Есть несохранённые изменения")
        palette = self._lbl_modified.palette()
        palette.setColor(QPalette.ColorRole.WindowText, QColor(210, 100, 0))
        self._lbl_modified.setPalette(palette)
        self._lbl_modified.setVisible(False)
        bottom.addWidget(self._lbl_modified)

        bottom.addStretch()

        self._btn_save = QPushButton("💾 Сохранить")
        self._btn_save.setToolTip("Сохранить шаблон (Ctrl+S)")
        self._btn_save.setFixedHeight(26)
        self._btn_save.clicked.connect(self.save_requested)
        bottom.addWidget(self._btn_save)

        grp_layout.addLayout(bottom)

    # ------------------------------------------------------------------ public

    def load_from_template(self, tmpl: StampTemplate) -> None:
        """Populate all fields from a StampTemplate instance."""
        self._updating = True
        self._txt_name.setText(tmpl.name or "")
        for dt, chk in self._chk_doc_types.items():
            chk.setChecked(dt in (tmpl.doc_types or []))
        idx = self._cmb_page.findText(tmpl.page_selector or "first")
        self._cmb_page.setCurrentIndex(idx if idx >= 0 else 0)
        self._spn_priority.setValue(tmpl.priority if tmpl.priority is not None else 10)
        self._spn_padding.setValue(tmpl.padding_mm if tmpl.padding_mm is not None else 0.5)
        self._chk_grid_adapt.setChecked(bool(getattr(tmpl, "grid_adapt", False)))
        fm = getattr(tmpl, "frame_mode", "gost") or "gost"
        idx_fm = self._cmb_frame_mode.findData(fm)
        self._cmb_frame_mode.setCurrentIndex(idx_fm if idx_fm >= 0 else 0)
        self._spn_grid_tol.setValue(float(getattr(tmpl, "grid_tolerance_template_mm", 2.0)))
        self._spn_detected_tol.setValue(float(getattr(tmpl, "grid_tolerance_detected_mm", 0.5)))
        self._spn_snap_max_dist.setValue(float(getattr(tmpl, "snap_max_distance_mm", 5.0)))
        self._spn_max_shape.setValue(float(getattr(tmpl, "max_shape_change_ratio", 2.5)))
        self._spn_min_score.setValue(float(getattr(tmpl, "cascade_score_threshold", 0.4)))
        x_snap = getattr(tmpl, "find_tables_snap_x_tolerance", None)
        y_snap = getattr(tmpl, "find_tables_snap_y_tolerance", None)
        if x_snap is not None or y_snap is not None:
            self._chk_find_tables_snap.setChecked(True)
            self._spn_find_tables_snap_x.setValue(
                float(x_snap) if x_snap is not None else 2.2
            )
            self._spn_find_tables_snap_y.setValue(
                float(y_snap) if y_snap is not None else 2.0
            )
        else:
            self._chk_find_tables_snap.setChecked(False)
            self._spn_find_tables_snap_x.setValue(2.2)
            self._spn_find_tables_snap_y.setValue(2.0)
        self._on_find_tables_snap_toggled()
        self._updating = False

    def set_modified(self, modified: bool) -> None:
        """Show/hide the unsaved indicator."""
        self._lbl_modified.setVisible(modified)

    def get_name(self) -> str:
        return self._txt_name.text().strip() or "template"

    def get_doc_types(self) -> list[str]:
        return [dt for dt, chk in self._chk_doc_types.items() if chk.isChecked()]

    def get_page_selector(self) -> str:
        return self._cmb_page.currentText()

    def get_priority(self) -> int:
        return self._spn_priority.value()

    def get_padding_mm(self) -> float:
        return self._spn_padding.value()

    def get_grid_adapt(self) -> bool:
        return self._chk_grid_adapt.isChecked()

    def get_frame_mode(self) -> str:
        return str(self._cmb_frame_mode.currentData() or "gost")

    def get_grid_tolerance_template_mm(self) -> float:
        return self._spn_grid_tol.value()

    def set_grid_tolerance_template_mm(self, value: float) -> None:
        self._spn_grid_tol.setValue(value)

    def set_grid_tolerance_detected_mm(self, value: float) -> None:
        self._spn_detected_tol.setValue(value)

    def set_snap_max_distance_mm(self, value: float) -> None:
        self._spn_snap_max_dist.setValue(value)

    def get_grid_tolerance_detected_mm(self) -> float:
        return self._spn_detected_tol.value()

    def get_snap_max_distance_mm(self) -> float:
        return self._spn_snap_max_dist.value()

    def get_max_shape_change_ratio(self) -> float:
        return self._spn_max_shape.value()

    def set_max_shape_change_ratio(self, value: float) -> None:
        self._spn_max_shape.setValue(value)

    def get_cascade_score_threshold(self) -> float:
        return self._spn_min_score.value()

    def set_cascade_score_threshold(self, value: float) -> None:
        self._spn_min_score.setValue(value)

    def get_find_tables_snap_x_tolerance(self) -> float | None:
        if not self._chk_find_tables_snap.isChecked():
            return None
        return float(self._spn_find_tables_snap_x.value())

    def get_find_tables_snap_y_tolerance(self) -> float | None:
        if not self._chk_find_tables_snap.isChecked():
            return None
        return float(self._spn_find_tables_snap_y.value())

    # ------------------------------------------------------------------ private

    def _on_find_tables_snap_toggled(self, *_args) -> None:
        en = self._chk_find_tables_snap.isChecked()
        self._spn_find_tables_snap_x.setEnabled(en)
        self._spn_find_tables_snap_y.setEnabled(en)

    def _on_adv_toggled(self, expanded: bool) -> None:
        self._adv_widget.setVisible(expanded)
        self._btn_adv.setArrowType(
            Qt.ArrowType.DownArrow if expanded else Qt.ArrowType.RightArrow
        )

    def _on_changed(self, *_args) -> None:
        if not self._updating:
            self.meta_changed.emit()
