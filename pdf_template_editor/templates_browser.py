"""Браузер шаблонов (IDE-style левая панель): проекты > шаблоны.

Структура дерева:
    project_a              <- project folder (bold)
      catalog.json         <- FieldCatalog (особая иконка)
      dwg_page1.json       <- StampTemplate
      dwg_page2.json
    project_b
      catalog.json
      ...

Двойной клик:
  - catalog.json   → catalog_activate_requested(path)
  - шаблон         → template_load_requested(path)
  - папка проекта  → развернуть/свернуть (или активировать проект)

Контекстное меню (правая кнопка):
  - на шаблоне: «Загрузить», «Показать в проводнике»
  - на catalog.json: «Редактировать каталог», «Активировать»
  - на проекте: «Обновить»
"""

from __future__ import annotations

import os

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QMenu,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from pdf_parsing_v2_engine.template_loader import load_projects


class TemplatesBrowserPanel(QWidget):
    """Left sidebar: project tree from templates_dir."""

    template_load_requested = Signal(str)   # path to .json template
    catalog_activate_requested = Signal(str)  # path to catalog.json
    catalog_edit_requested = Signal(str)    # path to catalog.json (open editor)
    message = Signal(str, str)              # level: info|warn|error, text

    # kept for compat — no longer emitted
    set_edit_requested = Signal(str)

    ROLE_PATH = Qt.ItemDataRole.UserRole
    ROLE_KIND = Qt.ItemDataRole.UserRole + 1

    KIND_PROJECT = 0     # project folder row
    KIND_CATALOG = 1     # catalog.json row
    KIND_TMPL = 2        # StampTemplate .json row

    def __init__(self, templates_dir: str = "", parent=None):
        super().__init__(parent)
        self._templates_dir: str = templates_dir
        self._active_template_path: str = ""
        self._active_catalog_path: str = ""
        self._active_project_folder: str = ""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(4)

        hdr = QHBoxLayout()
        lbl = QLabel("Проекты")
        f = lbl.font()
        f.setBold(True)
        lbl.setFont(f)
        hdr.addWidget(lbl)
        hdr.addStretch()
        btn_refresh = QPushButton("↻")
        btn_refresh.setFixedWidth(28)
        btn_refresh.setToolTip("Обновить список файлов")
        btn_refresh.clicked.connect(self.refresh)
        hdr.addWidget(btn_refresh)
        layout.addLayout(hdr)

        self._tree = QTreeWidget()
        self._tree.setHeaderHidden(True)
        self._tree.setRootIsDecorated(True)
        self._tree.setIndentation(14)
        self._tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._tree.itemDoubleClicked.connect(self._on_double_clicked)
        self._tree.customContextMenuRequested.connect(self._on_context_menu)
        layout.addWidget(self._tree)

        if templates_dir:
            self.refresh()

    # ------------------------------------------------------------------ public

    def set_templates_dir(self, templates_dir: str) -> None:
        self._templates_dir = templates_dir
        self.refresh()

    def set_active_template(self, path: str) -> None:
        self._active_template_path = os.path.normpath(path) if path else ""
        self._update_active_marks()

    def set_active_catalog(self, path: str) -> None:
        self._active_catalog_path = os.path.normpath(path) if path else ""
        self._update_active_marks()

    def set_active_project(self, folder: str) -> None:
        self._active_project_folder = os.path.normpath(folder) if folder else ""
        self._update_active_marks()

    def refresh(self) -> None:
        """Rescan templates_dir and rebuild project tree."""
        self._tree.clear()

        td = self._templates_dir
        if not td or not os.path.isdir(td):
            self._update_active_marks()
            return

        projects = load_projects(td)
        if not projects:
            # Ничего не найдено — показать подсказку
            hint = QTreeWidgetItem(["(нет проектов)"])
            hint.setForeground(0, QBrush(QColor(150, 150, 150)))
            hint.setFlags(Qt.ItemFlag.ItemIsEnabled)
            self._tree.addTopLevelItem(hint)
            return

        for proj in projects:
            proj_item = self._make_project_item(proj.name, proj.folder)

            # catalog.json — первым
            cat_item = self._make_catalog_item(proj.catalog_path)
            proj_item.addChild(cat_item)

            # шаблоны
            for tmpl in proj.templates:
                fname = os.path.basename(tmpl.source_path) if hasattr(tmpl, "source_path") else ""
                if not fname:
                    # имя из пути (source_path может отсутствовать у old templates)
                    fname = "???"
                tmpl_item = self._make_template_item(fname, _tmpl_path(tmpl, proj.folder))
                proj_item.addChild(tmpl_item)

            self._tree.addTopLevelItem(proj_item)
            proj_item.setExpanded(True)

        self._update_active_marks()

    # ------------------------------------------------------------------ private helpers

    def _make_project_item(self, name: str, folder: str) -> QTreeWidgetItem:
        item = QTreeWidgetItem([name])
        item.setData(0, self.ROLE_PATH, folder)
        item.setData(0, self.ROLE_KIND, self.KIND_PROJECT)
        item.setToolTip(0, folder)
        f = item.font(0)
        f.setBold(True)
        item.setFont(0, f)
        item.setForeground(0, QBrush(QColor(40, 80, 160)))
        item.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
        )
        return item

    def _make_catalog_item(self, catalog_path: str) -> QTreeWidgetItem:
        item = QTreeWidgetItem(["catalog.json"])
        item.setData(0, self.ROLE_PATH, catalog_path)
        item.setData(0, self.ROLE_KIND, self.KIND_CATALOG)
        item.setToolTip(0, f"{catalog_path}\n(каталог полей проекта)")
        item.setForeground(0, QBrush(QColor(0, 140, 80)))
        item.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
        )
        return item

    def _make_template_item(self, fname: str, full_path: str) -> QTreeWidgetItem:
        item = QTreeWidgetItem([fname])
        item.setData(0, self.ROLE_PATH, full_path)
        item.setData(0, self.ROLE_KIND, self.KIND_TMPL)
        item.setToolTip(0, full_path)
        item.setFlags(
            Qt.ItemFlag.ItemIsEnabled
            | Qt.ItemFlag.ItemIsSelectable
        )
        return item

    def _update_active_marks(self) -> None:
        active_tmpl = self._active_template_path
        active_cat = self._active_catalog_path
        active_proj = self._active_project_folder

        for i in range(self._tree.topLevelItemCount()):
            proj_it = self._tree.topLevelItem(i)
            if proj_it is None:
                continue
            proj_folder = proj_it.data(0, self.ROLE_PATH) or ""
            proj_active = bool(active_proj and os.path.normpath(proj_folder) == active_proj)

            pf = proj_it.font(0)
            pf.setBold(True)
            proj_it.setFont(0, pf)
            proj_it.setForeground(
                0,
                QBrush(QColor(20, 60, 200) if proj_active else QColor(40, 80, 160)),
            )

            for j in range(proj_it.childCount()):
                child = proj_it.child(j)
                if child is None:
                    continue
                kind = child.data(0, self.ROLE_KIND)
                path = child.data(0, self.ROLE_PATH) or ""
                norm = os.path.normpath(path) if path else ""

                if kind == self.KIND_CATALOG:
                    is_active = bool(active_cat and norm == active_cat)
                    cf = child.font(0)
                    cf.setBold(is_active)
                    child.setFont(0, cf)
                    child.setForeground(
                        0,
                        QBrush(QColor(0, 160, 100) if is_active else QColor(0, 140, 80)),
                    )
                elif kind == self.KIND_TMPL:
                    is_active = bool(active_tmpl and norm == active_tmpl)
                    tf = child.font(0)
                    tf.setBold(is_active)
                    child.setFont(0, tf)
                    child.setForeground(
                        0,
                        QBrush(QColor(60, 120, 220) if is_active else QColor(30, 30, 30)),
                    )

    def _on_double_clicked(self, item: QTreeWidgetItem, _col: int) -> None:
        kind = item.data(0, self.ROLE_KIND)
        path = item.data(0, self.ROLE_PATH)
        if not path:
            return

        if kind == self.KIND_TMPL:
            self.template_load_requested.emit(str(path))
        elif kind == self.KIND_CATALOG:
            # Двойной клик: активировать каталог (и открыть редактор — см. T7.1)
            self.catalog_activate_requested.emit(str(path))
        elif kind == self.KIND_PROJECT:
            # Развернуть/свернуть проект
            item.setExpanded(not item.isExpanded())

    def _on_context_menu(self, pos) -> None:
        item = self._tree.itemAt(pos)
        if item is None:
            return
        kind = item.data(0, self.ROLE_KIND)
        path = str(item.data(0, self.ROLE_PATH) or "")

        menu = QMenu(self)

        if kind == self.KIND_TMPL:
            act_load = menu.addAction("Загрузить шаблон")
            act_load.triggered.connect(lambda: self.template_load_requested.emit(path))
            menu.addSeparator()
            act_show = menu.addAction("Показать в проводнике")
            act_show.triggered.connect(lambda: _open_in_explorer(path))

        elif kind == self.KIND_CATALOG:
            act_edit = menu.addAction("Редактировать каталог")
            act_edit.triggered.connect(lambda: self.catalog_edit_requested.emit(path))
            act_activate = menu.addAction("Активировать")
            act_activate.triggered.connect(lambda: self.catalog_activate_requested.emit(path))
            menu.addSeparator()
            act_show = menu.addAction("Показать в проводнике")
            act_show.triggered.connect(lambda: _open_in_explorer(path))

        elif kind == self.KIND_PROJECT:
            act_refresh = menu.addAction("Обновить")
            act_refresh.triggered.connect(self.refresh)
            menu.addSeparator()
            act_show = menu.addAction("Открыть папку")
            act_show.triggered.connect(lambda: _open_in_explorer(path, is_dir=True))

        if menu.isEmpty():
            return
        menu.exec(self._tree.viewport().mapToGlobal(pos))


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _tmpl_path(tmpl, project_folder: str) -> str:
    """Extract absolute path from StampTemplate."""
    sp = getattr(tmpl, "source_path", None)
    if sp and os.path.isfile(sp):
        return sp
    # fallback: search by template name in project folder
    p = os.path.join(project_folder, f"{tmpl.name}.json")
    if os.path.isfile(p):
        return p
    return p


def _open_in_explorer(path: str, is_dir: bool = False) -> None:
    import subprocess
    try:
        if is_dir:
            subprocess.Popen(["explorer", os.path.normpath(path)])
        else:
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
    except Exception:
        pass
