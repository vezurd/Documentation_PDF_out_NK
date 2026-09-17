"""Сборка PROJECT_CONTEXT_FOR_CHAT.txt. Запуск: python _gen_project_context.py"""
from __future__ import annotations

import ast
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SKIP_DIRS = frozenset({".git", "__pycache__", ".venv", "venv", "node_modules", ".pytest_cache", ".cursor"})
TREE_NO_DESCEND = frozenset({"cache", ".idea", "_release"})
TREE_NO_DESCEND_NAMES = frozenset({"adapt_debug"})

# Не перечислять файлы в дереве — одна строка с описанием и ~числом файлов.
def _count_files_under(d: Path) -> int:
    try:
        return sum(1 for _ in d.rglob("*") if _.is_file())
    except OSError:
        return -1


def tree_summarize_branch(dir_path: Path, rel_to_root: Path) -> str | None:
    """Если не None — не спускаться внутрь; дописать суффикс к строке каталога."""
    rel_pos = rel_to_root.as_posix()
    parts = rel_to_root.parts
    name = rel_to_root.name

    if rel_pos == "templates":
        n = _count_files_under(dir_path)
        suf = f"~{n} files" if n >= 0 else "many files"
        return (
            f"  [шаблоны выгрузок Excel, JSON штампов, HTML, примеры имён документов; {suf}; перечисление опущено]"
        )

    if rel_pos == "GUI/test":
        n = _count_files_under(dir_path)
        suf = f"~{n} files" if n >= 0 else ""
        return f"  [демо CustomTkinter (виджеты/темы); {suf.strip()}; перечисление опущено]"

    if len(parts) >= 2 and parts[-2] == "pdf_parsing_v2" and name == "grid_diagnostic_output":
        n = _count_files_under(dir_path)
        suf = f"~{n} files" if n >= 0 else ""
        return f"  [выход grid diagnostic: PNG-оверлеи и вложенные отчёты; {suf.strip()}; дерево опущено]"

    if rel_pos == "pdf_parsing_v2/templates/test_pdf":
        n = _count_files_under(dir_path)
        suf = f"~{n} files" if n >= 0 else ""
        return (
            f"  [тестовые PDF v2 (BBB/CJ/DWG/MTO/OD); {suf.strip()}; перечисление опущено]"
        )

    if rel_pos == "pdf_parsing_v2/templates/agcc_287":
        n = _count_files_under(dir_path)
        suf = f"~{n} files" if n >= 0 else ""
        return f"  [JSON шаблоны проекта agcc_287; {suf.strip()}; перечисление опущено]"

    if rel_pos == "pdf_parsing_v2/templates/catalogs":
        n = _count_files_under(dir_path)
        suf = f"~{n} files" if n >= 0 else ""
        return f"  [каталоги полей fields_*.json; {suf.strip()}; перечисление опущено]"

    if rel_pos == "pdf_parsing_v2/templates/sets":
        n = _count_files_under(dir_path)
        suf = f"~{n} files" if n >= 0 else ""
        return f"  [legacy TemplateSet *.json; {suf.strip()}; перечисление опущено]"

    if rel_pos == "RFQ/ds_compare/templates":
        n = _count_files_under(dir_path)
        suf = f"~{n} files" if n >= 0 else ""
        return f"  [xlsx-шаблоны сравнения ДС↔MTO; {suf.strip()}; перечисление опущено]"

    return None


def should_skip_symbol_scan(rel_posix: str) -> bool:
    """Не дублировать API демо-папок в секции классов/функций."""
    if rel_posix.startswith("GUI/test/"):
        return True
    return False

PROJECT_TOP_PACKAGES = frozenset(
    {
        "RFQ",
        "base",
        "utils",
        "tags",
        "pdf_parsing_v2",
        "cable_mapping",
        "GUI",
        "nano_cad",
    }
)

# Назначение файла (если нет/скудный модульный docstring). Пути — posix от корня репозитория.
FILE_PURPOSE: dict[str, str] = {
    "main.py": "Главное GUI (CustomTkinter): PDF v2 (монитор), сравнение тегов RFP↔MTO↔VO, BBB, DS/MTO, кабельные карты, nano_cad, zip; subprocess + busy-кнопки.",
    "gui_subprocess_job.py": "Паттерн фоновой работы: threading, статус на кнопке, результат subprocess.",
    "help_descriptions.py": "Тексты подсказок для кнопок GUI.",
    "RFQ/tags_rfp_compare/agregate_tags.py": "Оркестратор tags_rfp_compare: step1–4, кэш, load_base параллельно step2/3, timing/memory логи.",
    "RFQ/tags_rfp_compare/step4_analyze_and_match.py": "Весь этап 4: цикл по title_mark, optional ProcessPoolExecutor, merge, постобработка, Excel.",
    "RFQ/tags_rfp_compare/step4/step4_worker.py": "Один title_mark: split, check MTO/VO, match, статусы, unmatched, collapse, postmerge per-row.",
    "RFQ/tags_rfp_compare/step1_load_rfp.py": "Чтение RFP Excel → RowStd; опционально split по тегам/VALUES.",
    "RFQ/tags_rfp_compare/step2_load_mto.py": "Загрузка MTO (кэш), RowStd, опциональный экспорт шага 2.",
    "RFQ/tags_rfp_compare/step3_load_vo.py": "Загрузка VO (кэш), RowStd, связь с тегами.",
    "RFQ/tags_rfp_compare/rfp_tags_utils.py": "JSON-конфиг сравнения тегов, пути result_dir, буферизованные timing/memory логи.",
    "RFQ/tags_rfp_compare/rfp_tags_settings_gui.py": "Окно настроек tags_rfp_compare.",
    "RFQ/tags_rfp_compare/column_optimization.py": "Сужение набора колонок итоговой выгрузки по хэшу конфига.",
    "RFQ/tags_rfp_compare/compare_step4_files.py": "Сравнение двух Excel step4 как мультимножества строк.",
    "RFQ/tags_rfp_compare/mto_file_filter.py": "Отбор MTO-файлов по правилам путей/имён.",
    "RFQ/tags_rfp_compare/step4/step4_1_check_mto_data.py": "Проверка и нормализация MTO/VO, split «без тегов», отчёты по тегам.",
    "RFQ/tags_rfp_compare/step4/step4_2_match_rfp_with_mto.py": "Сопоставление RFP↔MTO (коды, теги, оптимизации pass_match_code).",
    "RFQ/tags_rfp_compare/step4/step4_3_assign_position_status.py": "Назначение POSITION_STATUS и производных статусов строки.",
    "RFQ/tags_rfp_compare/step4/step4_4_add_unmatched_mto_rows.py": "Несопоставленные MTO/VO, дозаполнение, часть с optional MP.",
    "RFQ/tags_rfp_compare/step4/step4_5_check_values_sum.py": "Сверка сумм VALUES между RFP/MTO/VO после матчинга.",
    "RFQ/tags_rfp_compare/step4/step4_6_save_match_result_to_excel.py": "Запись итогового Excel step4 (стили, автофильтр, быстрый writer).",
    "RFQ/tags_rfp_compare/step4/step4_6_cell_colors.py": "Раскраска ячеек по состоянию CheckElement.",
    "RFQ/tags_rfp_compare/step4/step4_postmerge_ops.py": "Пост-merge: обогащение из Google base, TAG_EFFECTIVE, лоты, цвета, служебные поля.",
    "base/base_classes.py": "RowStd, CheckElement, RowType; копии строк для split (_COPY_COLS, batch_copy_light).",
    "base/tables_columns.py": "Имена колонок Excel и статусов матчинга; общие константы таблиц.",
    "base/base_xlsx_load.py": "Загрузка xlsx в RowStd, кэш формул, снятие зачёркивания, подготовка к 1С.",
    "base/base_google.py": "Code base с Google Drive / локальный кэш (meta + json).",
    "base/base_excel_out.py": "Выгрузка таблиц RowStd в Excel (общие утилиты).",
    "base/bbb_analysis.py": "Конвейер анализа BBB-спецификаций.",
    "base/bbb_load.py": "Поиск файлов BBB по конфигу.",
    "base/bbb_config.py": "Загрузка JSON конфигурации BBB.",
    "base/bbb_excel_out.py": "Формирование отчётных Excel по BBB.",
    "base/bbb_checks.py": "Правила проверок строк/секций BBB.",
    "utils/cache_utils.py": "Кэш загрузки MTO/VO Excel на диске, инвалидация по mtime/size.",
    "utils/release_zip.py": "Сборка zip-дистрибутива для коллег (.release_zipignore).",
    "utils/path.py": "Пути к папкам PDF/результатов, имена файлов, open_dir.",
    "tags/tag_parser.py": "Парсинг тегов из текста PDF, глобальный source_dict, analyze/merge для parallel.",
    "tags/tag_classes.py": "Классы представления тегов.",
    "pdf_parsing_v2/stamp_extractor.py": "Движок v2: рамка → кандидаты шаблонов → поля (fitz/char index), grid_adapt.",
    "pdf_parsing_v2/grid_matcher.py": "Привязка линий шаблона к сетке PDF: find_tables, walk, progressive interp, отчёт.",
    "pdf_parsing_v2/grid_lines_utils.py": "Генерация grid lines из полей, refs, snap bindings, схлопывание дублей.",
    "pdf_parsing_v2/v2_pipeline.py": "Пакетный прогон папки PDF через движок v2 и шаблоны.",
    "pdf_parsing_v2/models.py": "Датаклассы шаблона, поля, grid line, результатов страницы.",
    "pdf_parsing_v2/coord_transform.py": "Мм ↔ fitz pts, bbox полей, линии сетки, учёт rotation.",
    "pdf_parsing_v2/frame_detector.py": "Поиск границ штампа (рамки) на странице.",
    "pdf_parsing_v2/template_loader.py": "Обнаружение проектов шаблонов (catalog.json) и загрузка JSON.",
    "pdf_parsing_v2/char_text_extractor.py": "Индекс символов страницы для точного извлечения текста в bbox.",
    "pdf_parsing_v2/v2_config.py": "Путь и чтение/запись pdf_v2_config.json.",
    "pdf_parsing_v2/pdf_v2_settings_gui.py": "Окно настроек PDF v2.",
    "pdf_template_editor/main_window.py": "Редактор шаблонов PySide6: PDF-сцена, поля, wireframe, тест-адаптация.",
    "pdf_template_editor/wireframe_panel.py": "Панель линий сетки и привязок полей к линиям.",
    "pdf_template_editor/adapt_debug.py": "Снимок отладки адаптации: JSON + alignment_log.txt.",
    "cable_mapping/cab_mapping.py": "Старт конвейера кабельных карт / сводок.",
    "cable_mapping/mapping/load_map_from_google.py": "Загрузка mapping с Google + локальный кэш.",
    "nano_cad/nc_start.py": "Точка входа подсистемы экспорта полей в NanoCAD.",
}

# Точечные подписи для часто читаемых классов без своего docstring.
CLASS_NAME_NOTES: dict[str, str] = {
    "RowStd": "Строка таблицы: dict колонка → CheckElement, тип строки, комментарии.",
    "CheckElement": "Ячейка: value, struck_value, color.",
    "RowType": "Категория строки (секция, позиция, пустая, заголовок, …).",
    "TableComments": "Метаданные файла/листа для веток MTO/VO/RFP в парсере.",
}


def should_skip_dir(rel: Path) -> bool:
    return any(part in SKIP_DIRS for part in rel.parts)


def tree_omit_reason(dir_name: str) -> str | None:
    if dir_name in TREE_NO_DESCEND:
        return "содержимое опущено (кэш/IDE/релизы)"
    if dir_name in TREE_NO_DESCEND_NAMES:
        return "содержимое опущено (отладочные артефакты)"
    return None


def tree_lines(base: Path, prefix: str = "") -> list[str]:
    lines: list[str] = []
    try:
        entries = sorted(
            [x for x in base.iterdir() if not should_skip_dir(x.relative_to(ROOT))],
            key=lambda x: (not x.is_dir(), x.name.lower()),
        )
    except PermissionError:
        return lines
    for i, e in enumerate(entries):
        is_last = i == len(entries) - 1
        branch = "└── " if is_last else "├── "
        if e.is_dir():
            reason = tree_omit_reason(e.name)
            if reason:
                try:
                    n = sum(1 for _ in e.rglob("*") if _.is_file())
                except OSError:
                    n = -1
                suf = f"  [{reason}; ~{n} files]" if n >= 0 else f"  [{reason}]"
                lines.append(prefix + branch + e.name + "/" + suf)
                continue
            rel_e = e.relative_to(ROOT)
            summ = tree_summarize_branch(e, rel_e)
            if summ:
                lines.append(prefix + branch + e.name + "/" + summ)
                continue
        lines.append(prefix + branch + e.name + ("/" if e.is_dir() else ""))
        if e.is_dir():
            ext = "    " if is_last else "│   "
            lines.extend(tree_lines(e, prefix + ext))
    return lines


def first_paragraph(text: str, max_len: int = 320) -> str:
    if not text or not text.strip():
        return ""
    block = text.strip().split("\n\n", 1)[0]
    one_line = re.sub(r"\s+", " ", block.replace("\r\n", "\n")).strip()
    if len(one_line) > max_len:
        return one_line[: max_len - 1] + "…"
    return one_line


def module_docstring(tree: ast.Module) -> str:
    if not tree.body:
        return ""
    first = tree.body[0]
    if isinstance(first, ast.Expr) and isinstance(first.value, ast.Constant):
        val = first.value.value
        if isinstance(val, str):
            return first_paragraph(val)
    return ""


def brief_doc_paragraph(node: ast.AST) -> str:
    if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
        return ""
    if not node.body:
        return ""
    first = node.body[0]
    if isinstance(first, ast.Expr) and isinstance(first.value, ast.Constant):
        val = first.value.value
        if isinstance(val, str):
            return first_paragraph(val)
    return ""


def base_name(node: ast.expr) -> str:
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        return f"{base_name(node.value)}.{node.attr}"
    return ast.unparse(node)


def _body_without_docstring(
    node: ast.FunctionDef | ast.AsyncFunctionDef,
) -> list[ast.stmt]:
    body = list(node.body)
    if body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant):
        if isinstance(body[0].value.value, str):
            body = body[1:]
    return body


def function_stmt_count(node: ast.FunctionDef | ast.AsyncFunctionDef) -> int:
    body = _body_without_docstring(node)
    return sum(1 for st in body if not isinstance(st, ast.Pass))


def infer_symbol_note(kind: str, name: str, bases: list[str], stmt_count: int, has_doc: bool) -> str:
    """Короткая подпись только если нет docstring; мелкие функции пропускаем."""
    if has_doc:
        return ""
    if kind == "def":
        if name.startswith("wire_"):
            return "Связывание кнопки/виджета с фоновой задачей или командой."
        if stmt_count <= 2:
            return ""
        if name == "main":
            return "Точка входа при запуске файла как скрипта."
        if name.startswith("load_"):
            return "Загрузка данных или конфигурации."
        if name.startswith("save_") or name.startswith("write_"):
            return "Сохранение/запись результата."
        if name.startswith("get_"):
            return "Получение значения, пути или конфига."
        if name.startswith("set_") or name.startswith("enable_") or name.startswith("disable_"):
            return "Установка параметра или флага."
        if name.startswith("run_") or name.startswith("start"):
            return "Запуск сценария или подсистемы."
        if name.startswith("build_") or name.startswith("create_"):
            return "Сборка структуры или объекта."
        if name.startswith("parse_") or name.startswith("read_"):
            return "Разбор или чтение входных данных."
        if name.startswith("export_"):
            return "Экспорт в файл."
        if name.startswith("check_") or name.startswith("validate_"):
            return "Проверка данных или условий."
        if name.startswith("process_") or name.startswith("analyze_"):
            return "Обработка или анализ набора данных."
        if name.startswith("compare_"):
            return "Сравнение сущностей или результатов."
        if name.startswith("merge_"):
            return "Слияние данных или папок."
        if name.startswith("extract_"):
            return "Извлечение признаков или текста."
        if name.startswith("find_") or name.startswith("search_"):
            return "Поиск совпадений или файлов."
        if name.startswith("ensure_"):
            return "Гарантия инварианта (путь, директория и т.д.)."
        if name.startswith("append_") or name.startswith("finalize_"):
            return "Накопление лога / финализация записи."
        if name.startswith("show_") or name.startswith("open_"):
            return "Показ UI или открытие диалога/папки."
        if name.startswith("excel_"):
            return "Формирование или запись Excel."
        return ""
    # class
    if name in CLASS_NAME_NOTES:
        return CLASS_NAME_NOTES[name]
    bb = ",".join(bases)
    if any(x in bb for x in ("QWidget", "QMainWindow", "QDialog", "QGraphics", "QFrame")):
        return "Компонент GUI (Qt)."
    if "Enum" in bb:
        return "Перечисление."
    if "Exception" in bb or "Error" in name:
        return "Тип исключения."
    if not bases and stmt_count <= 3:
        return ""
    return "Контейнер данных / логики модуля."


def base_class_stmt_hint(class_node: ast.ClassDef) -> int:
    """Грубая оценка «размера» класса: число методов верхнего уровня (без вложенных классов)."""
    n = 0
    for item in class_node.body:
        if isinstance(item, (ast.FunctionDef, ast.AsyncFunctionDef)):
            n += 1
    return n


def scan_py(path: Path) -> tuple[list[tuple[str, str, str, list[str], int, str]], str | None]:
    """Список (kind, name, doc, bases, stmt_count, infer_note); второе — ошибка парсера."""
    try:
        src = path.read_text(encoding="utf-8", errors="replace")
        tree = ast.parse(src)
    except Exception as e:
        return [], str(e)

    out: list[tuple[str, str, str, list[str], int, str]] = []
    for node in tree.body:
        if isinstance(node, ast.ClassDef):
            doc = brief_doc_paragraph(node)
            bases = [base_name(b) for b in node.bases]
            stmts = base_class_stmt_hint(node)
            infer = infer_symbol_note("class", node.name, bases, stmts, bool(doc))
            out.append(("class", node.name, doc, bases, stmts, infer))
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            if node.name.startswith("_") and not node.name.startswith("__"):
                continue
            doc = brief_doc_paragraph(node)
            sc = function_stmt_count(node)
            infer = infer_symbol_note("def", node.name, [], sc, bool(doc))
            out.append(("def", node.name, doc, [], sc, infer))
    return out, None


def top_level_imports(path: Path) -> tuple[list[str], list[str]]:
    try:
        src = path.read_text(encoding="utf-8", errors="replace")
        tree = ast.parse(src)
    except Exception:
        return [], []
    mods: list[str] = []
    for node in tree.body:
        if isinstance(node, ast.Import):
            for alias in node.names:
                mods.append(alias.name.split(".")[0])
        elif isinstance(node, ast.ImportFrom):
            if node.module:
                mods.append(node.module.split(".")[0])
    uniq = sorted(set(mods))
    proj = [m for m in uniq if m in PROJECT_TOP_PACKAGES]
    return uniq, proj


def package_of_relpath(rel: str) -> str:
    parts = rel.replace("\\", "/").split("/")
    if len(parts) == 1:
        return "(root)"
    top = parts[0]
    return top if top in PROJECT_TOP_PACKAGES else f"({top})"


def collect_package_edges(py_paths: list[Path]) -> dict[tuple[str, str], int]:
    edges: dict[tuple[str, str], int] = {}
    for p in py_paths:
        rel = p.relative_to(ROOT).as_posix()
        if should_skip_dir(Path(rel)):
            continue
        src_pkg = package_of_relpath(rel)
        try:
            src = p.read_text(encoding="utf-8", errors="replace")
            tree = ast.parse(src)
        except Exception:
            continue
        targets: set[str] = set()
        for node in tree.body:
            if isinstance(node, ast.Import):
                for alias in node.names:
                    r = alias.name.split(".")[0]
                    if r in PROJECT_TOP_PACKAGES:
                        targets.add(r)
            elif isinstance(node, ast.ImportFrom):
                if node.module:
                    r = node.module.split(".")[0]
                    if r in PROJECT_TOP_PACKAGES and r != src_pkg:
                        targets.add(r)
        for t in targets:
            key = (src_pkg, t)
            edges[key] = edges.get(key, 0) + 1
    return edges


def reverse_package_edges(edges: dict[tuple[str, str], int]) -> dict[str, list[tuple[str, int]]]:
    by_to: dict[str, list[tuple[str, int]]] = {}
    for (a, b), c in edges.items():
        by_to.setdefault(b, []).append((a, c))
    for b in by_to:
        by_to[b].sort(key=lambda x: (-x[1], x[0]))
    return by_to


def build_module_links_lines(
    imports_by_file: dict[str, list[str]],
    edges: dict[tuple[str, str], int],
) -> list[str]:
    out: list[str] = [
        str(ROOT),
        "Связи модулей по статическому анализу import/from (только верхний уровень файла, ast).",
        "Это не граф вызовов функций: отражены зависимости пакетов друг от друга.",
        f"Учитываются корни: {', '.join(sorted(PROJECT_TOP_PACKAGES))}.",
        "Обновление: python _gen_project_context.py",
        "",
        "=== Матрица: файлов в пакете A с импортом пакета B ===",
    ]
    by_from: dict[str, list[tuple[str, int]]] = {}
    for (a, b), c in edges.items():
        by_from.setdefault(a, []).append((b, c))
    for a in sorted(by_from.keys(), key=lambda x: (x.startswith("("), x.lower())):
        parts = sorted(by_from[a], key=lambda x: (-x[1], x[0]))
        out.append(f"  {a}:")
        for b, c in parts:
            out.append(f"    → {b}: {c} file(s)")

    out.append("")
    out.append("=== Обратная проекция: кто импортирует пакет B (число файлов-источников) ===")
    by_to = reverse_package_edges(edges)
    for b in sorted(by_to.keys(), key=lambda x: (x.startswith("("), x.lower())):
        out.append(f"  {b} ←")
        for a, c in by_to[b]:
            out.append(f"    {a}: {c} file(s)")

    out.append("")
    out.append("=== Файлы → импортируемые пакеты проекта ===")
    for rel, im in sorted(imports_by_file.items()):
        out.append(f"  {rel} -> {', '.join(im)}")
    return out


def main() -> None:
    lines: list[str] = []
    lines.append(str(ROOT))
    lines.append("Цель: сжатый контекст для другого чата (структура, API, автономность).")
    lines.append("Обновление: python _gen_project_context.py")
    lines.append(
        "Связи пакетов (импорты между base, RFQ, pdf_parsing_v2, …): см. PROJECT_MODULE_LINKS.txt"
    )
    lines.append("")
    lines.append(
        "Подписи к символам: docstring (первый абзац); иначе — эвристика по имени для «толстых» "
        "функций (>2 stmt); у классов Qt/Enum — краткая метка. Мелкие хелперы без docstring не помечаются."
    )
    lines.append("")
    lines.append("=== ДЕРЕВО (пропуск: .git __pycache__ .venv venv node_modules .pytest_cache .cursor) ===")
    lines.extend(tree_lines(ROOT))

    py_files: list[Path] = []
    for p in ROOT.rglob("*.py"):
        rel = p.relative_to(ROOT)
        if should_skip_dir(rel):
            continue
        py_files.append(p)

    imports_by_file: dict[str, list[str]] = {}
    for p in sorted(py_files, key=lambda x: str(x).lower()):
        rel = p.relative_to(ROOT).as_posix()
        _, proj = top_level_imports(p)
        if proj:
            imports_by_file[rel] = proj

    edges = collect_package_edges(py_files)
    links_path = ROOT / "PROJECT_MODULE_LINKS.txt"
    links_path.write_text(
        "\n".join(build_module_links_lines(imports_by_file, edges)),
        encoding="utf-8",
    )

    lines.append("")
    lines.append("=== PYTHON: классы и публичные функции (верхний уровень; _private пропущены) ===")
    for p in sorted(py_files, key=lambda x: str(x).lower()):
        rel = p.relative_to(ROOT).as_posix()
        if rel == "_gen_project_context.py":
            continue
        if should_skip_symbol_scan(rel):
            continue
        try:
            src = p.read_text(encoding="utf-8", errors="replace")
            tree = ast.parse(src)
        except Exception as e:
            lines.append(f"--- {rel} ---")
            lines.append(f"  [parse error] {e}")
            lines.append("")
            continue

        items, err = scan_py(p)
        if err:
            lines.append(f"--- {rel} ---")
            lines.append(f"  [parse error] {err}")
            lines.append("")
            continue
        if not items:
            continue

        lines.append(f"--- {rel} ---")
        purpose = FILE_PURPOSE.get(rel)
        mod_doc = module_docstring(tree)
        if purpose:
            lines.append(f"  [файл] {purpose}")
        elif mod_doc:
            lines.append(f"  [модуль] {mod_doc}")

        for kind, name, doc, bases, _stmt, infer in items:
            if kind == "class":
                b = (" -> " + ", ".join(bases)) if bases else ""
                lines.append(f"  class {name}{b}")
                if doc:
                    lines.append(f"    # {doc}")
                elif infer:
                    lines.append(f"    # {infer}")
            else:
                lines.append(f"  def {name}()")
                if doc:
                    lines.append(f"    # {doc}")
                elif infer:
                    lines.append(f"    # {infer}")
        lines.append("")

    lines.append("")
    lines.append("=== АНАЛИТИКА: автономные куски и связи (ручная сводка для рефакторинга) ===")
    lines.append("""
Пакеты (логические подсистемы):
- main.py — точка входа GUI/оркестрации (теги RFP, PDF v2 через монитор, zip, 1C и т.д.).
- RFQ/ — сравнение спецификаций: tags_rfp_compare (ядро), ds_compare, ZIP, pre_Address и др.
- cable_mapping/ — отдельная подсистема кабельных карт/сводок (импортирует base/utils/tags).
- base/ — общие классы строк Excel, колонки, Google base, BBB Excel, загрузка xlsx.
- utils/ — кэш, файлы, zip релиза, вспомогательные функции.
- tags/ — парсинг тегов.
- pdf_parsing_v2/ — оркестратор v2; pdf_parsing_v2_engine/ — движок; pdf_template_editor/ — Qt-редактор шаблонов; grid matcher; шаблоны JSON.

Относительно автономные модули (мало или нет импортов из «своего» дерева):
- utils/* часто зависит только от stdlib / paths — кандидат в общую библиотеку.
- tags/tag_parser.py — обычно изолирован от RFQ step4.
- pdf_parsing_v2* — единственная ветка PDF в репозитории; метаданные из имени файла — `V2Document.from_file_path` (см. utils/path.py).

Сильные связи (упростить = ввести интерфейсы / пакетные границы):
- RFQ/tags_rfp_compare/* <-> base/* (RowStd, columns, xlsx) — ядро домена.
- RFQ/tags_rfp_compare/* <-> utils/cache_utils — кэш MTO/VO.
- pdf_template_editor/* <-> pdf_parsing_v2/{models,coord_transform,grid_matcher,grid_lines_utils,...} — толстый UI-слой.
- main.py импортирует почти всё — «божественный объект»; вынос подкоманд в cli/подпакеты снизит связность.

Рекомендации для независимости (для следующего чата):
1) Выделить пакет rfp_domain (base RowStd + tables_columns + rfp_tags_utils) с явным публичным API.
2) step1–step4 оставить как pipeline, зависящий только от rfp_domain + cache interface.
3) pdf_parsing_v2: слой core (без PyQt) vs editor (только Qt + вызовы core).
4) Документировать контракты JSON шаблонов (models) отдельно от UI.

Файл сгенерирован _gen_project_context.py — при больших изменениях перезапустите: python _gen_project_context.py
""".strip())

    out_path = ROOT / "PROJECT_CONTEXT_FOR_CHAT.txt"
    out_path.write_text("\n".join(lines), encoding="utf-8")
    print(out_path, out_path.stat().st_size, "bytes")
    print(links_path, links_path.stat().st_size, "bytes")


if __name__ == "__main__":
    main()
