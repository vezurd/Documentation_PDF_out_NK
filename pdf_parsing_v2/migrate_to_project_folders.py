"""Скрипт миграции: flat templates/ + catalogs/ + sets/ → project folder structure.

Для каждого набора (sets/*.json):
  1. Создать подпапку templates/<project_folder_name>/
  2. Скопировать шаблоны из корня templates/ в подпапку
  3. Прочитать каталог из catalogs/, добавить поля projects/description
  4. Сохранить как <project_folder_name>/catalog.json
  5. (опционально) удалить исходные файлы

Запуск:
    python pdf_parsing_v2/migrate_to_project_folders.py [--dry-run] [--delete-originals]

    --dry-run          : только показать что будет сделано, не изменять файлы
    --delete-originals : удалить исходные flat шаблоны, sets/ и catalogs/ после успешной миграции
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys


def _slug(name: str) -> str:
    """Convert set name to a valid folder name."""
    safe = "".join(c if c.isalnum() or c in "-_." else "_" for c in name)
    return safe.strip("_")


def migrate(templates_dir: str, dry_run: bool = False, delete_originals: bool = False) -> None:
    sets_dir = os.path.join(templates_dir, "sets")
    catalogs_dir = os.path.join(templates_dir, "catalogs")

    if not os.path.isdir(sets_dir):
        print(f"[migrate] Нет папки sets/ в {templates_dir!r} — нечего мигрировать.")
        return

    set_files = sorted(
        f for f in os.listdir(sets_dir) if f.lower().endswith(".json")
    )
    if not set_files:
        print("[migrate] Папка sets/ пуста — нечего мигрировать.")
        return

    migrated_template_files: set[str] = set()
    migrated_ok: list[str] = []

    for set_fname in set_files:
        set_path = os.path.join(sets_dir, set_fname)
        with open(set_path, encoding="utf-8") as f:
            set_data = json.load(f)

        set_name: str = set_data.get("name", os.path.splitext(set_fname)[0])
        projects: list[str] = set_data.get("projects", [])
        field_catalog_ref: str = set_data.get("field_catalog", "")
        template_files: list[str] = set_data.get("template_files", [])

        # Folder name = slug from set filename (e.g. agcc_287.json → agcc_287)
        folder_name = os.path.splitext(set_fname)[0]
        project_folder = os.path.join(templates_dir, folder_name)

        print(f"\n=== Набор: {set_name!r} -> папка {folder_name!r} ===")
        print(f"  Проекты: {projects}")
        print(f"  Шаблонов: {len(template_files)}")

        if not dry_run:
            os.makedirs(project_folder, exist_ok=True)

        # --- Copy / migrate templates ---
        copied_templates: list[str] = []
        for tf in template_files:
            # tf is basename or relative path
            basename = os.path.basename(tf)
            src = os.path.join(templates_dir, basename)
            if not os.path.isfile(src):
                print(f"  [WARN] Шаблон не найден: {src!r}")
                continue
            dst = os.path.join(project_folder, basename)
            print(f"  Копировать: {basename}")
            if not dry_run:
                shutil.copy2(src, dst)
            copied_templates.append(src)
            migrated_template_files.add(os.path.normpath(src))

        # --- Load and upgrade catalog ---
        catalog_data: dict = {"name": set_name, "entries": []}
        catalog_src_path: str = ""

        if field_catalog_ref:
            # Resolve path: may be relative to project root or absolute
            if os.path.isabs(field_catalog_ref):
                cat_path = field_catalog_ref
            else:
                # Try relative to project root (parent of templates_dir)
                project_root = os.path.dirname(templates_dir)
                cat_path = os.path.normpath(os.path.join(project_root, field_catalog_ref))
                if not os.path.isfile(cat_path):
                    # Try relative to templates_dir
                    cat_path = os.path.normpath(os.path.join(templates_dir, field_catalog_ref))
                if not os.path.isfile(cat_path):
                    # Try just basename in catalogs/
                    cat_path = os.path.join(catalogs_dir, os.path.basename(field_catalog_ref))

            if os.path.isfile(cat_path):
                with open(cat_path, encoding="utf-8") as f:
                    catalog_data = json.load(f)
                catalog_src_path = cat_path
                print(f"  Каталог: {os.path.basename(cat_path)!r}")
            else:
                print(f"  [WARN] Каталог не найден: {field_catalog_ref!r}")

        # Inject project metadata
        catalog_data["projects"] = projects
        if not catalog_data.get("description"):
            catalog_data["description"] = set_name

        dst_catalog = os.path.join(project_folder, "catalog.json")
        print(f"  Сохранить каталог: catalog.json")
        if not dry_run:
            with open(dst_catalog, "w", encoding="utf-8") as f:
                json.dump(catalog_data, f, ensure_ascii=False, indent=2)

        migrated_ok.append(folder_name)

    # --- Optional: also copy full_grid templates not in any set ---
    extra_templates = []
    for fname in os.listdir(templates_dir):
        if not fname.lower().endswith(".json"):
            continue
        fpath = os.path.normpath(os.path.join(templates_dir, fname))
        if fpath not in migrated_template_files:
            if os.path.isfile(fpath):
                extra_templates.append(fname)

    if extra_templates:
        print(f"\n  Шаблоны не вошедшие ни в один набор: {extra_templates}")
        print("  (не мигрируются автоматически — добавьте вручную в нужный проект)")

    # --- Delete originals ---
    if delete_originals and not dry_run and migrated_ok:
        print("\n=== Удаление исходных файлов ===")
        for src in sorted(migrated_template_files):
            if os.path.isfile(src):
                os.remove(src)
                print(f"  Удалён: {os.path.basename(src)}")
        if os.path.isdir(sets_dir):
            shutil.rmtree(sets_dir)
            print(f"  Удалена папка: sets/")
        if os.path.isdir(catalogs_dir):
            shutil.rmtree(catalogs_dir)
            print(f"  Удалена папка: catalogs/")

    if dry_run:
        print("\n[dry-run] Изменений не внесено.")
    else:
        print(f"\n[migrate] Готово. Мигрировано проектов: {len(migrated_ok)}")
        for name in migrated_ok:
            print(f"  OK {name}/")


def main() -> None:
    parser = argparse.ArgumentParser(description="Migrate templates to project folder structure")
    parser.add_argument(
        "templates_dir",
        nargs="?",
        default=None,
        help="Путь к папке templates (по умолчанию: pdf_parsing_v2_engine/templates от корня проекта)",
    )
    parser.add_argument("--dry-run", action="store_true", help="Только показать план")
    parser.add_argument(
        "--delete-originals",
        action="store_true",
        help="Удалить исходные flat шаблоны и папки sets/, catalogs/ после миграции",
    )
    args = parser.parse_args()

    if args.templates_dir:
        td = os.path.abspath(args.templates_dir)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        root = os.path.dirname(script_dir)
        td = os.path.join(root, "pdf_parsing_v2_engine", "templates")

    if not os.path.isdir(td):
        print(f"[migrate] Папка не найдена: {td!r}", file=sys.stderr)
        sys.exit(1)

    print(f"[migrate] templates_dir = {td!r}")
    migrate(td, dry_run=args.dry_run, delete_originals=args.delete_originals)


if __name__ == "__main__":
    main()
