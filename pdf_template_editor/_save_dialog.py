"""Dialog for template metadata before saving (name, doc_types, page_selector, etc.)."""

from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QFormLayout,
    QHBoxLayout,
    QPushButton,
    QLineEdit,
    QComboBox,
    QSpinBox,
    QDoubleSpinBox,
    QLabel,
    QCheckBox,
    QGroupBox,
)


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


class SaveTemplateDialog(QDialog):
    def __init__(self, n_fields: int, n_unassigned: int, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Сохранить шаблон")
        self.resize(420, 400)

        layout = QVBoxLayout(self)

        if n_unassigned > 0:
            layout.addWidget(QLabel(
                f"Внимание: {n_unassigned} неназначенных ячеек будут пропущены.\n"
                f"Сохраняется {n_fields} назначенных полей."
            ))

        form = QFormLayout()
        self._txt_name = QLineEdit("DWG стр.1 (стандартный)")
        form.addRow("Имя шаблона:", self._txt_name)

        self._cmb_page = QComboBox()
        self._cmb_page.addItems(["first", "rest", "all"])
        form.addRow("page_selector:", self._cmb_page)

        self._spn_priority = QSpinBox()
        self._spn_priority.setRange(1, 100)
        self._spn_priority.setValue(10)
        form.addRow("priority:", self._spn_priority)

        self._spn_padding = QDoubleSpinBox()
        self._spn_padding.setRange(0.0, 10.0)
        self._spn_padding.setDecimals(1)
        self._spn_padding.setValue(0.5)
        form.addRow("padding_mm:", self._spn_padding)

        layout.addLayout(form)

        grp = QGroupBox("doc_types")
        grp_layout = QVBoxLayout(grp)
        self._chk_doc_types: dict[str, QCheckBox] = {}
        row = QHBoxLayout()
        for i, dt in enumerate(_ALL_DOC_TYPES):
            chk = QCheckBox(dt)
            chk.setChecked(dt in ("WIR", "LAY", "CAE", "GA", "PL", "DW", "NI"))
            self._chk_doc_types[dt] = chk
            row.addWidget(chk)
            if (i + 1) % 6 == 0:
                grp_layout.addLayout(row)
                row = QHBoxLayout()
        if row.count() > 0:
            grp_layout.addLayout(row)
        layout.addWidget(grp)

        bottom = QHBoxLayout()
        bottom.addStretch()
        btn_ok = QPushButton("Сохранить")
        btn_ok.clicked.connect(self.accept)
        bottom.addWidget(btn_ok)
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(self.reject)
        bottom.addWidget(btn_cancel)
        layout.addLayout(bottom)

    def template_name(self) -> str:
        return self._txt_name.text().strip() or "template"

    def doc_types(self) -> list[str]:
        return [dt for dt, chk in self._chk_doc_types.items() if chk.isChecked()]

    def page_selector(self) -> str:
        return self._cmb_page.currentText()

    def priority(self) -> int:
        return self._spn_priority.value()

    def padding_mm(self) -> float:
        return self._spn_padding.value()
