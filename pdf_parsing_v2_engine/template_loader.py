"""
Загрузка JSON-шаблонов и проектов.

Структура папки шаблонов (новый формат):
    templates_dir/
      project_a/
        catalog.json       <- FieldCatalog (маркер проекта)
        dwg_page1.json     <- StampTemplate
        dwg_page2.json
      project_b/
        catalog.json
        mto_page1.json

Каждая подпапка с catalog.json считается проектом (ProjectTemplates).
Плоские *.json в корне templates_dir по-прежнему загружаются для
обратной совместимости через load_all_templates().
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field

from pdf_parsing_v2_engine.models import FieldCatalog, StampTemplate, merge_stamp_template_with_catalog


@dataclass
class ProjectTemplates:
    """Шаблоны одного проекта, загруженные из project folder."""
    catalog: FieldCatalog
    catalog_path: str
    folder: str                        # абсолютный путь к папке проекта
    templates: list[StampTemplate] = field(default_factory=list)

    @property
    def name(self) -> str:
        return self.catalog.name or os.path.basename(self.folder)


def list_project_combo_entries(templates_dir: str) -> list[tuple[str, str]]:
    """Лёгкий список проектов для GUI (выпадающий список Project).

    Для каждой подпапки с ``catalog.json`` возвращает пары
    ``(отображаемое_имя, basename_папки)`` — только чтение каталога,
    без загрузки файлов шаблонов (в отличие от :func:`load_projects`).
    """
    if not os.path.isdir(templates_dir):
        return []

    out: list[tuple[str, str]] = []
    for entry in sorted(os.scandir(templates_dir), key=lambda e: e.name):
        if not entry.is_dir():
            continue
        catalog_path = os.path.join(entry.path, "catalog.json")
        if not os.path.isfile(catalog_path):
            continue
        try:
            catalog = FieldCatalog.from_json(catalog_path)
        except (OSError, ValueError, json.JSONDecodeError):
            continue
        folder_name = os.path.basename(entry.path)
        display = catalog.name or folder_name
        out.append((display, folder_name))
    return out


def load_projects(templates_dir: str) -> list[ProjectTemplates]:
    """Сканирует подпапки templates_dir; каждая папка с catalog.json = проект.

    Возвращает список ProjectTemplates, отсортированный по имени папки.
    Подпапки без catalog.json игнорируются (например test_pdf/).
    """
    if not os.path.isdir(templates_dir):
        return []

    projects: list[ProjectTemplates] = []
    for entry in sorted(os.scandir(templates_dir), key=lambda e: e.name):
        if not entry.is_dir():
            continue
        catalog_path = os.path.join(entry.path, "catalog.json")
        if not os.path.isfile(catalog_path):
            continue
        try:
            catalog = FieldCatalog.from_json(catalog_path)
        except (OSError, ValueError, json.JSONDecodeError):
            continue

        templates: list[StampTemplate] = []
        for fname in sorted(os.listdir(entry.path)):
            if not fname.lower().endswith(".json"):
                continue
            if fname.lower() == "catalog.json":
                continue
            fpath = os.path.join(entry.path, fname)
            if not os.path.isfile(fpath):
                continue
            try:
                t = StampTemplate.from_json(fpath)
                templates.append(merge_stamp_template_with_catalog(t, catalog))
            except (OSError, ValueError, json.JSONDecodeError):
                pass

        projects.append(ProjectTemplates(
            catalog=catalog,
            catalog_path=catalog_path,
            folder=entry.path,
            templates=templates,
        ))

    return projects


def load_all_templates(templates_dir: str) -> list[StampTemplate]:
    """Загружает все шаблоны из templates_dir.

    Сначала пробует проектный формат (подпапки с catalog.json).
    Если проектов нет — падает на legacy flat-scan корня templates_dir.
    Обратная совместимость: сигнатура и поведение неизменны для вызывающего кода.
    """
    projects = load_projects(templates_dir)
    if projects:
        out: list[StampTemplate] = []
        for p in projects:
            out.extend(p.templates)
        return out

    # Legacy: flat *.json в корне (без проектной структуры)
    return _load_flat(templates_dir)


def load_project_templates(
    templates_dir: str,
    project_filter: str | None,
) -> list[StampTemplate]:
    """Возвращает шаблоны только из проектов, чьи projects[] содержат project_filter.

    Если project_filter is None — возвращает все шаблоны (как load_all_templates).
    Сравнение регистронезависимое.
    """
    if project_filter is None:
        return load_all_templates(templates_dir)

    pf_lower = project_filter.lower()
    out: list[StampTemplate] = []
    for p in load_projects(templates_dir):
        project_ids = [pid.lower() for pid in p.catalog.projects]
        folder_name = os.path.basename(p.folder).lower()
        if pf_lower in project_ids or pf_lower == folder_name:
            out.extend(p.templates)
    return out


def load_catalog_for_project(
    templates_dir: str,
    project_filter: str | None,
) -> FieldCatalog | None:
    """Load ``catalog.json`` for the first project matching *project_filter*.

    Matching follows :func:`load_project_templates` (``catalog.projects[]`` and
    project folder name; case-insensitive). If several folders match, returns the
    catalog from the **first** match in :func:`load_projects` order (sorted folder
    names). If *project_filter* is empty or whitespace-only, returns ``None``.

    Args:
        templates_dir: Root directory containing per-project subfolders with
            ``catalog.json``.
        project_filter: Project id to match, or ``None`` / blank for no load.

    Returns:
        :class:`FieldCatalog` for the matched project, or ``None`` if no match.
    """
    if not project_filter or not str(project_filter).strip():
        return None
    pf_lower = str(project_filter).strip().lower()
    for p in load_projects(templates_dir):
        project_ids = [pid.lower() for pid in p.catalog.projects]
        folder_name = os.path.basename(p.folder).lower()
        if pf_lower in project_ids or pf_lower == folder_name:
            return p.catalog
    return None


def _load_flat(templates_dir: str) -> list[StampTemplate]:
    """Legacy: загружает *.json напрямую из корня templates_dir (без подпапок)."""
    if not os.path.isdir(templates_dir):
        return []
    out: list[StampTemplate] = []
    for name in sorted(os.listdir(templates_dir)):
        if not name.lower().endswith(".json"):
            continue
        path = os.path.join(templates_dir, name)
        if os.path.isfile(path):
            try:
                out.append(StampTemplate.from_json(path))
            except (OSError, ValueError, json.JSONDecodeError):
                pass
    return out
