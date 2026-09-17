"""QDialog для редактирования TemplateSet (привязка шаблонов к проектам)."""

from __future__ import annotations

import os

from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QPushButton,
    QLineEdit,
    QListWidget,
    QFileDialog,
    QMessageBox,
)

from pdf_parsing_v2_engine.models import TemplateSet


class SetEditorDialog(QDialog):
    def __init__(
        self,
        parent=None,
        template_set: TemplateSet | None = None,
        templates_dir: str = "",
    ):
        super().__init__(parent)
        self.setWindowTitle("Набор шаблонов")
        self.resize(500, 450)
        self._tset = template_set
        self._templates_dir = templates_dir

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self._txt_name = QLineEdit()
        form.addRow("Имя набора:", self._txt_name)

        self._txt_catalog = QLineEdit()
        btn_catalog = QPushButton("…")
        btn_catalog.setFixedWidth(30)
        btn_catalog.clicked.connect(self._pick_catalog)
        row_cat = QHBoxLayout()
        row_cat.addWidget(self._txt_catalog)
        row_cat.addWidget(btn_catalog)
        form.addRow("Каталог полей:", row_cat)

        layout.addLayout(form)

        # projects
        layout.addWidget(_section_label("Проекты"))
        self._lst_projects = QListWidget()
        layout.addWidget(self._lst_projects)
        prow = QHBoxLayout()
        self._txt_proj = QLineEdit()
        self._txt_proj.setPlaceholderText("AGCC.287")
        prow.addWidget(self._txt_proj)
        btn_add_proj = QPushButton("+")
        btn_add_proj.setFixedWidth(30)
        btn_add_proj.clicked.connect(self._add_project)
        prow.addWidget(btn_add_proj)
        btn_del_proj = QPushButton("−")
        btn_del_proj.setFixedWidth(30)
        btn_del_proj.clicked.connect(self._del_project)
        prow.addWidget(btn_del_proj)
        layout.addLayout(prow)

        # template files
        layout.addWidget(_section_label("Шаблоны"))
        self._lst_templates = QListWidget()
        layout.addWidget(self._lst_templates)
        trow = QHBoxLayout()
        btn_add_tmpl = QPushButton("Добавить…")
        btn_add_tmpl.clicked.connect(self._add_template)
        trow.addWidget(btn_add_tmpl)
        btn_del_tmpl = QPushButton("Удалить")
        btn_del_tmpl.clicked.connect(self._del_template)
        trow.addWidget(btn_del_tmpl)
        trow.addStretch()
        layout.addLayout(trow)

        # bottom buttons
        bottom = QHBoxLayout()
        btn_load = QPushButton("Загрузить…")
        btn_load.clicked.connect(self._load)
        bottom.addWidget(btn_load)
        btn_save = QPushButton("Сохранить…")
        btn_save.clicked.connect(self._save)
        bottom.addWidget(btn_save)
        bottom.addStretch()
        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.accept)
        bottom.addWidget(btn_close)
        layout.addLayout(bottom)

        if template_set:
            self._fill(template_set)

    def _fill(self, ts: TemplateSet) -> None:
        self._txt_name.setText(ts.name)
        self._txt_catalog.setText(ts.field_catalog)
        self._lst_projects.clear()
        self._lst_projects.addItems(ts.projects)
        self._lst_templates.clear()
        self._lst_templates.addItems(ts.template_files)

    def _collect(self) -> TemplateSet:
        projects = [self._lst_projects.item(i).text() for i in range(self._lst_projects.count())]
        templates = [self._lst_templates.item(i).text() for i in range(self._lst_templates.count())]
        return TemplateSet(
            name=self._txt_name.text().strip() or "набор",
            projects=projects,
            field_catalog=self._txt_catalog.text().strip(),
            template_files=templates,
        )

    def _catalogs_dir(self) -> str:
        if self._templates_dir:
            d = os.path.join(self._templates_dir, "catalogs")
            if os.path.isdir(d):
                return d
            return self._templates_dir
        return ""

    def _sets_dir(self) -> str:
        if self._templates_dir:
            d = os.path.join(self._templates_dir, "sets")
            if os.path.isdir(d):
                return d
            return self._templates_dir
        return ""

    def _pick_catalog(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Каталог полей", self._catalogs_dir(), "JSON (*.json)"
        )
        if path:
            self._txt_catalog.setText(path)

    def _add_project(self) -> None:
        t = self._txt_proj.text().strip()
        if t:
            self._lst_projects.addItem(t)
            self._txt_proj.clear()

    def _del_project(self) -> None:
        for it in self._lst_projects.selectedItems():
            self._lst_projects.takeItem(self._lst_projects.row(it))

    def _add_template(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Добавить шаблон", self._templates_dir or "", "JSON (*.json)"
        )
        if path:
            self._lst_templates.addItem(path)

    def _del_template(self) -> None:
        for it in self._lst_templates.selectedItems():
            self._lst_templates.takeItem(self._lst_templates.row(it))

    def _load(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "Загрузить набор", self._sets_dir(), "JSON (*.json)"
        )
        if not path:
            return
        try:
            ts = TemplateSet.from_json(path)
            self._fill(ts)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))

    def _save(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить набор", self._sets_dir(), "JSON (*.json)"
        )
        if not path:
            return
        ts = self._collect()
        try:
            ts.to_json(path)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))


def _section_label(text: str):
    from PySide6.QtWidgets import QLabel
    from PySide6.QtGui import QFont
    lbl = QLabel(text)
    f = lbl.font()
    f.setBold(True)
    lbl.setFont(f)
    return lbl
