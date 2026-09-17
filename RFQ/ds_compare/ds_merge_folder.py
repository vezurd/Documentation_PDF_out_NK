from __future__ import annotations

import os
import re
from pathlib import Path
from typing import List

from prettytable import PrettyTable

from base.base_classes import RowStd, RowType
from base.base_class_std_table import STDTable
from base.tables_columns import (
    CODE,
    DS_DELIVERY_TIME,
    DS_NAME,
    DS_PACKAGING_PRICE_EXCL_VAT,
    DS_SHIPPING_PRICE_EXCL_VAT,
    DS_TITLE,
    DS_TOTAL_PRICE_EXCL_VAT,
    DS_TOTAL_PRICE_INCL_VAT,
    DS_UNIT_PRICE_EXCL_VAT,
    DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT,
    DS_VAT_AMOUNT,
    DS_SYSTEM,
    NAME,
    ROW_TYPE,
    TYPE_MARK,
    UNITS,
    VALUES,
    VENDOR,
)
import utils.path
from RFQ.ds_compare.ds_quantity_parse import try_parse_quantity
from RFQ.ds_compare.support_finctions import load_ds_data

# UTF-8 рядом со сводным xlsx: по умолчанию только other_row (см. write_merge_ds_empty_other_rows_txt).
MERGE_DS_EMPTY_OTHER_ROWS_FILENAME = "DS_merge_empty_and_other_rows.txt"

# UTF-8: все найденные xlsx и результат load_ds_data (OK / Empty / Error).
MERGE_DS_LOAD_REPORT_FILENAME = "DS_merge_load_report.txt"

# Первое «ДС» + цифры в имени файла (кириллица) — ключ сортировки отчёта загрузки (4 < 47).
_DS_MERGE_LOAD_REPORT_NUM_RE = re.compile(r"ДС(\d+)")

# Опционально: дамп всех строк одного Имя ДС с перечитыванием xlsx (см. debug_merge_ds_folder_by_ds_name).
MERGE_DS_FOLDER_DEBUG_DS_NAME: str | None = None


def _normalize_ds_merge_label(value: str | None) -> str:
    """Match labels like 'ДС66', 'ДС 66', 'дс66' to the same key."""
    if value is None:
        return ""
    s = str(value).strip().lower()
    return re.sub(r"\s+", "", s)


# Порядок колонок в DS_merge_empty_and_other_rows.txt (значения ячеек через get_value; ROW_TYPE → row.row_type).
_MERGE_DS_EMPTY_OTHER_DUMP_KEYS: tuple[str, ...] = (
    DS_NAME,
    DS_TITLE,
    DS_SYSTEM,
    CODE,
    NAME,
    TYPE_MARK,
    UNITS,
    VALUES,
    DS_UNIT_PRICE_EXCL_VAT,
    DS_PACKAGING_PRICE_EXCL_VAT,
    DS_SHIPPING_PRICE_EXCL_VAT,
    DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT,
    DS_TOTAL_PRICE_EXCL_VAT,
    DS_VAT_AMOUNT,
    DS_TOTAL_PRICE_INCL_VAT,
    DS_DELIVERY_TIME,
    VENDOR,
    ROW_TYPE,
)

_MERGE_DS_EMPTY_OTHER_DUMP_HEADERS: tuple[str, ...] = (
    "DS_NAME",
    "DS_TITLE",
    "DS_SYSTEM",
    "CODE",
    "NAME",
    "TYPE_MARK",
    "UNITS",
    "VALUES",
    "DS_UNIT_PRICE_EXCL_VAT",
    "DS_PACKAGING_PRICE_EXCL_VAT",
    "DS_SHIPPING_PRICE_EXCL_VAT",
    "DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT",
    "DS_TOTAL_PRICE_EXCL_VAT",
    "DS_VAT_AMOUNT",
    "DS_TOTAL_PRICE_INCL_VAT",
    "DS_DELIVERY_TIME",
    "VENDOR",
    "ROW_TYPE",
)


# Максимальная длина текста в ячейке для DS_merge_empty_and_other_rows.txt (PrettyTable).
_MERGE_DS_EMPTY_OTHER_CELL_MAX_LEN = 40


def _merge_ds_dump_cell_value(row: RowStd, key: str):
    """Value for merge txt column *key* (ROW_TYPE uses ``row.row_type``)."""
    if key == ROW_TYPE:
        return row.row_type
    return row.get_value(key)


def _merge_ds_empty_other_cell_str(value, max_len: int = _MERGE_DS_EMPTY_OTHER_CELL_MAX_LEN) -> str:
    """Single-line string for txt cell, truncated to *max_len* characters."""
    if value is None:
        s = ""
    else:
        s = str(value).replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    if len(s) <= max_len:
        return s
    if max_len <= 1:
        return s[:max_len]
    return s[: max_len - 1] + "…"


def write_merge_ds_empty_other_rows_txt(
    merged: List[RowStd],
    out_dir: str,
    *,
    include_empty_row: bool = False,
) -> str | None:
    """
    Writes UTF-8 text with PrettyTable: rows omitted from the summary xlsx.

    By default only ``other_row`` is listed. Pass ``include_empty_row=True`` to also
    include ``empty_row`` (previous behavior).

    If nothing matches, no file is created and ``None`` is returned.

    Columns match ``_MERGE_DS_EMPTY_OTHER_DUMP_KEYS`` / ``_MERGE_DS_EMPTY_OTHER_DUMP_HEADERS``;
    each cell is truncated to ``_MERGE_DS_EMPTY_OTHER_CELL_MAX_LEN``. The same text is printed to stdout.

    Args:
        merged: Rows after ``load_ds_data`` / merge scan.
        out_dir: Same directory as ``DS_summary*.xlsx``.
        include_empty_row: If False, dump only ``other_row``. If True, dump both
            ``empty_row`` and ``other_row``.

    Returns:
        Path to the written ``.txt`` file, or ``None`` if there were no rows to dump.
    """
    want: frozenset[str] = (
        frozenset({RowType.empty_row, RowType.other_row})
        if include_empty_row
        else frozenset({RowType.other_row})
    )
    picked = [r for r in merged if r.row_type in want]

    if not picked:
        mode = "other_row only" if not include_empty_row else "empty_row and other_row"
        print(f"[DS merge] Дамп строк ({mode}): нет строк — файл не создаётся.")
        return None

    table = PrettyTable()
    table.field_names = list(_MERGE_DS_EMPTY_OTHER_DUMP_HEADERS)
    for fn in table.field_names:
        table.align[fn] = "l"
    table.max_width = _MERGE_DS_EMPTY_OTHER_CELL_MAX_LEN

    for row in picked:
        table.add_row(
            [
                _merge_ds_empty_other_cell_str(_merge_ds_dump_cell_value(row, key))
                for key in _MERGE_DS_EMPTY_OTHER_DUMP_KEYS
            ]
        )

    dump_label = "empty_row and other_row" if include_empty_row else "other_row only"
    header_lines = [
        f"DS merge: rows with row_type ({dump_label}), not written to DS_summary xlsx",
        f"total_merged_rows: {len(merged)}",
        f"dumped_rows: {len(picked)}",
        "",
    ]
    body = table.get_string()
    text_out = "\n".join(header_lines) + body + "\n"

    print(text_out, end="")

    utils.path.make_dir(out_dir)
    out_path = str(Path(out_dir) / MERGE_DS_EMPTY_OTHER_ROWS_FILENAME)
    Path(out_path).write_text(text_out, encoding="utf-8")
    print(
        f"[DS merge] Список строк ({dump_label}, не в сводном xlsx): "
        f"{len(picked)} шт. -> {out_path}"
    )
    return out_path


def _ds_merge_load_report_sort_key(file_name: str) -> tuple[int, str]:
    """Sort key: numeric value after the first Cyrillic ``ДС`` in the basename, then name."""
    m = _DS_MERGE_LOAD_REPORT_NUM_RE.search(file_name)
    if m:
        return (int(m.group(1)), file_name)
    return (1 << 30, file_name)


def _merge_ds_summary_output_middle(merged: List[RowStd], source_xlsx_count: int) -> str:
    """Middle part of merged summary xlsx: ``ДС{{min}}_ДС{{max}}_{{count}}`` from ``DS_NAME``."""
    nums: list[int] = []
    for row in merged:
        label = row.get_value(DS_NAME) or ""
        m = _DS_MERGE_LOAD_REPORT_NUM_RE.search(str(label))
        if m:
            nums.append(int(m.group(1)))
    if not nums:
        return f"merged_{source_xlsx_count}"
    lo, hi = min(nums), max(nums)
    return f"ДС{lo}_ДС{hi}_{source_xlsx_count}"


def write_merge_ds_load_report_txt(
    source_dir: str,
    out_dir: str,
    file_results: list[tuple[str, str, str]],
) -> str:
    """Write UTF-8 report (PrettyTable): each scanned xlsx and ``load_ds_data`` outcome.

    Rows are ordered by the integer after the first ``ДС`` in the file name (``ДС4`` before
    ``ДС47…``), then by file name.

    Args:
        source_dir: Folder that was scanned for ``*.xlsx``.
        out_dir: Result folder (same as ``DS_summary*.xlsx``).
        file_results: Rows ``(file_name, status, detail)`` where *status* is
            ``ОК``, ``Пусто``, or ``Ошибка``.

    Returns:
        Path to the written ``.txt`` file.
    """
    ordered = sorted(file_results, key=lambda row: _ds_merge_load_report_sort_key(row[0]))

    table = PrettyTable()
    table.field_names = ["имя_файла", "статус", "комментарий"]
    for fn in table.field_names:
        table.align[fn] = "l"
    table.max_width["имя_файла"] = 96
    table.max_width["комментарий"] = 72

    for file_name, status, detail in ordered:
        detail_clean = (
            str(detail).replace("\t", " ").replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
        )
        table.add_row([file_name, status, detail_clean])

    header_lines = [
        "DS merge: отчёт по загрузке xlsx (load_ds_data)",
        f"Папка источников: {source_dir}",
        f"Найдено файлов в списке: {len(file_results)}",
        "Порядок строк: по числу после первого «ДС» в имени файла (4 < 47), затем по имени.",
        "",
    ]
    body = table.get_string()
    text_out = "\n".join(header_lines) + body + "\n"
    utils.path.make_dir(out_dir)
    out_path = str(Path(out_dir) / MERGE_DS_LOAD_REPORT_FILENAME)
    Path(out_path).write_text(text_out, encoding="utf-8")
    print(f"[DS merge] Отчёт по загрузке файлов: {out_path}")
    return out_path


def debug_merge_ds_folder_by_ds_name(
    dir_path: str,
    ds_name: str,
    *,
    include_subfolders: bool = False,
    out_dir: str | None = None,
    open_folder: bool = False,
) -> str | None:
    """
    For merge-folder debugging: reload each source xlsx like ``merge_ds_folder``,
    keep rows whose ``DS_NAME`` matches *ds_name* (e.g. ``ДС66``), and print/save
    columns as read from Excel: ДС, Титул, Марка (TYPE_MARK), CODE, VALUES.

    Output: PrettyTable to console and UTF-8 text file. Pass *out_dir* from
    ``merge_ds_folder`` so the txt lands next to ``DS_summary*.xlsx`` (one timestamp
    folder). If *out_dir* is omitted, a new ``get_path_out_dir`` folder is used (CLI).

    Only ``DS_NAME`` is used for filtering. Row types are **not** filtered here:
    ``get_row_type`` for DS marks many real lines as ``empty_row`` / ``other_row`` /
    ``head_row`` if optional columns are blank, while ``merge_ds_folder`` still loads them into
    ``merged``. Export to xlsx hides ``empty_row`` / ``other_row`` / ``head_row`` — debug shows raw
    rows so dumps match what is in memory after ``load_ds_data``.

    Args:
        dir_path: Folder with per-DS ``.xlsx`` files.
        ds_name: Target ``Имя ДС`` label (must match value set from filename in
            ``load_ds_data``, e.g. ``ДС66``).
        include_subfolders: Passed to ``get_files_single``.
        out_dir: Optional explicit output directory (created if missing).
        open_folder: If True, open the output directory in the file explorer.

    Returns:
        Path to the written ``.txt`` file, or ``None`` if nothing to dump.
    """
    if not dir_path or not os.path.isdir(dir_path):
        print(f"Указана неверная папка: {dir_path}")
        return None

    print("[DS merge][debug] Режим дампа по Имя ДС (исходные xlsx перечитываются).")
    target = _normalize_ds_merge_label(ds_name)
    if not target:
        print("Пустое имя ДС для отладки")
        return None

    print("Собираем список xlsx файлов (режим отладки по Имя ДС)...")
    docs = utils.path.get_files_single(
        dir_path,
        endswith=(".xlsx", ".XLSX"),
        forbidden_endswith=(),
        sub_folders=include_subfolders,
    )
    if not docs:
        print("Файлы .xlsx не найдены")
        return None

    collected: list[tuple[str, RowStd]] = []
    for d in docs:
        try:
            ds_rows = load_ds_data(d.file_full_path, force_update=True, use_cache=False)
        except Exception as e:
            print(f"Ошибка загрузки '{d.file_name}': {e}")
            continue
        if not ds_rows:
            continue
        for row in ds_rows:
            if _normalize_ds_merge_label(row.get_value(DS_NAME)) != target:
                continue
            collected.append((d.file_name, row))

    if not collected:
        print(
            f"[DS merge][debug] Нет строк с Имя ДС={ds_name!r} после load_ds_data "
            f"(нормализовано: {target!r}). Проверьте имя файла: в имени должен быть "
            f"фрагмент «ДС» + две цифры, как задаёт find_ds_pattern в support_finctions."
        )
        return None

    table = PrettyTable()
    table.field_names = ["ДС", "Титул", "Марка", "CODE", "VALUES", "row_type"]
    for fn in table.field_names:
        table.align[fn] = "l"
    table.max_width["Марка"] = 48
    table.max_width["CODE"] = 24

    for _file_name, row in collected:
        table.add_row(
            [
                row.get_value(DS_NAME),
                row.get_value(DS_TITLE),
                row.get_value(TYPE_MARK),
                row.get_value(CODE),
                row.get_value(VALUES),
                row.row_type or "",
            ]
        )

    source_files = sorted({fn for fn, _ in collected})
    header_lines = [
        "debug_merge_ds_folder_by_ds_name",
        f"dir_path: {dir_path}",
        f"ds_name filter (normalized): {target!r} (input: {ds_name!r})",
        f"rows: {len(collected)}",
        f"source_xlsx: {', '.join(source_files)}",
        "",
    ]
    body = table.get_string()
    text_out = "\n".join(header_lines) + body + "\n"

    print("\n" + "=" * 80)
    print(text_out, end="")
    print("=" * 80 + "\n")

    out_base = out_dir if out_dir else utils.path.get_path_out_dir(dir_path)
    utils.path.make_dir(out_base)
    safe = re.sub(r'[^\w\-]+', "_", ds_name, flags=re.UNICODE).strip("_") or "ds"
    out_name = f"debug_merge_ds_{safe}.txt"
    out_path = str(Path(out_base) / out_name)
    Path(out_path).write_text(text_out, encoding="utf-8")
    print(f"[DS merge][debug] Сохранён текстовый дамп (рядом со сводным xlsx):")
    print(f"  {out_path}")

    if open_folder:
        utils.path.open_dir(out_path)

    return out_path


def merge_ds_folder(
    dir_path: str,
    include_subfolders: bool = True,
    out_file_prefix: str = "DS_summary",
    open_folder: bool = True,
    *,
    dump_include_empty_row: bool = False,
) -> str | None:
    """
    Ищет *.xlsx файлы в папке, загружает их через load_ds_data, склеивает в один список и
    сохраняет результат в сводный xlsx через STDTable.to_excel_ds_vs_mto_spec.

    :param dir_path: Путь к папке с DS файлами (.xlsx)
    :param include_subfolders: Искать ли файлы во вложенных папках
    :param out_file_prefix: Префикс имени выходного файла
    :param open_folder: Открыть папку с результатом по завершению
    :param dump_include_empty_row: Если True, текстовый дамп включает и ``empty_row``;
        по умолчанию False — в дампе только ``other_row``.
    :return: Путь к созданному xlsx файлу или None (при ошибке до записи сводки;
        отчёт ``DS_merge_load_report.txt`` в папке результата пишется, если найдены xlsx).
        Имя сводного файла: ``{{prefix}}_ДС{{min}}_ДС{{max}}_{{N}}.xlsx``, где min/max —
        по колонке ``DS_NAME`` в объединённых строках, *N* — число исходных xlsx в скане.
    """
    if not dir_path or not os.path.isdir(dir_path):
        print(f"Указана неверная папка: {dir_path}")
        print("[DS merge] ОСТАНОВ: неверный путь.\n")
        return None

    print("")
    print("=" * 72)
    print("[DS merge] НАЧАЛО: объединение ДС из папки")
    print(f"  Вход (исходные xlsx): {dir_path}")
    print("  Выход: подпапка вида __результат_проверки_<дата> (НЕ в каталог источников).")
    print(
        f"  UTF-8 txt `{MERGE_DS_EMPTY_OTHER_ROWS_FILENAME}`: по умолчанию только other_row "
        f"(+ колонка «Тип строки»); файл не создаётся, если other_row нет. "
        f"Пустые empty_row в дамп: merge_ds_folder(..., dump_include_empty_row=True)."
    )
    print(
        f"  UTF-8 txt `{MERGE_DS_LOAD_REPORT_FILENAME}`: PrettyTable — все xlsx и статус "
        f"load_ds_data; порядок строк по числу после «ДС» в имени (4 раньше 47)."
    )
    print("=" * 72)

    print("Собираем список xlsx файлов...")
    docs = utils.path.get_files_single(
        dir_path,
        endswith=(".xlsx", ".XLSX"),
        forbidden_endswith=(),
        sub_folders=include_subfolders,
    )

    if not docs:
        print("Файлы .xlsx не найдены")
        print("[DS merge] ОСТАНОВ: нет файлов.")
        print("=" * 72 + "\n")
        return None

    merged: List[RowStd] = []
    load_report_rows: list[tuple[str, str, str]] = []
    format_error_files: list[str] = []
    for d in docs:
        try:
            ds_rows = load_ds_data(d.file_full_path, force_update=True, use_cache=False)
            if ds_rows:
                invalid_quantities: list[tuple[int, object]] = []
                for row_index, row in enumerate(ds_rows, start=1):
                    if row.row_type != RowType.position_row:
                        continue
                    raw_quantity = row.get_value(VALUES)
                    ok, _ = try_parse_quantity(raw_quantity)
                    if not ok:
                        invalid_quantities.append((row_index, raw_quantity))
                if invalid_quantities:
                    samples = ", ".join(
                        f"строка {row_index}: {value!r}"
                        for row_index, value in invalid_quantities[:5]
                    )
                    detail = (
                        "Неправильный формат ДС: в колонке количества VALUES "
                        f"нечисловые значения ({len(invalid_quantities)}); {samples}. "
                        "Проверьте порядок столбцов исходного файла."
                    )
                    print(f"ОШИБКА ФОРМАТА '{d.file_name}': {detail}")
                    load_report_rows.append((d.file_name, "Ошибка формата", detail))
                    format_error_files.append(d.file_name)
                    continue
                n = len(ds_rows)
                merged.extend(ds_rows)
                print(f"Добавлено строк: {n} из {d.file_name}")
                load_report_rows.append((d.file_name, "ОК", f"{n} строк в сводку"))
            else:
                print(f"Пустой файл или нет данных: {d.file_name}")
                load_report_rows.append((d.file_name, "Пусто", "0 строк из load_ds_data"))
        except Exception as e:
            print(f"Ошибка загрузки '{d.file_name}': {e}")
            load_report_rows.append((d.file_name, "Ошибка", str(e)))

    out_dir = utils.path.get_path_out_dir(dir_path)
    utils.path.make_dir(out_dir)
    load_report_path = write_merge_ds_load_report_txt(dir_path, out_dir, load_report_rows)
    print(f"[DS merge] Папка результата (сюда пишутся сводка и debug-txt):")
    print(f"  {os.path.normpath(out_dir)}")

    if format_error_files:
        print("\n" + "=" * 72)
        print("[DS merge] ОСТАНОВ: обнаружен неправильный формат исходных ДС.")
        for file_name in format_error_files:
            print(f"  - {file_name}")
        print("Сводный Excel не создан, чтобы не смешивать смещённые колонки.")
        print(f"Подробности: {load_report_path}")
        print("=" * 72 + "\n")
        return None

    if not merged:
        print("Нет данных для вывода")
        print("[DS merge] ОСТАНОВ: нет строк для сводки.")
        print(f"  Отчёт по файлам: {load_report_path}")
        print("=" * 72 + "\n")
        return None

    label = (MERGE_DS_FOLDER_DEBUG_DS_NAME or "").strip()
    debug_txt_path: str | None = None
    if label:
        print(f"[DS merge] Опция: дамп всех строк одного Имя ДС={label!r} (перечитывание xlsx)")
        debug_txt_path = debug_merge_ds_folder_by_ds_name(
            dir_path,
            label,
            include_subfolders=include_subfolders,
            out_dir=out_dir,
            open_folder=False,
        )
        if not debug_txt_path:
            print(
                "[DS merge] Опция: файл debug_merge_ds_*.txt не создан "
                "(см. сообщения [DS merge][debug] выше)."
            )

    empty_other_txt_path = write_merge_ds_empty_other_rows_txt(
        merged,
        out_dir,
        include_empty_row=dump_include_empty_row,
    )

    summary_middle = _merge_ds_summary_output_middle(merged, len(load_report_rows))
    out_path = STDTable.to_excel_ds_vs_mto_spec(
        merged,
        out_dir,
        out_file_prefix,
        show_row_type=False,
        output_middle_name=summary_middle,
    )

    if open_folder:
        utils.path.open_dir(out_path)

    print("")
    print("=" * 72)
    print("[DS merge] ЗАВЕРШЕНО: объединение ДС из папки")
    print(f"  Сводный Excel: {out_path}")
    print(f"  Отчёт по загрузке файлов (txt): {load_report_path}")
    if empty_other_txt_path:
        print(f"  Дамп не вошедших в сводку строк (txt): {empty_other_txt_path}")
    else:
        print("  Дамп не вошедших в сводку строк (txt): не создан (нет other_row при текущих флагах).")
    if debug_txt_path:
        print(f"  Опция Имя ДС (txt): {debug_txt_path}")
    print("  Дальнейшие сообщения в консоли могут быть от других частей программы.")
    print("=" * 72 + "\n")

    return out_path


if __name__ == "__main__":
    import argparse

    default_dir = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\Спецификации к ДС в Excel"
    parser = argparse.ArgumentParser(
        description="Merge DS xlsx in a folder, or dump loaded rows for one DS label.",
    )
    parser.add_argument(
        "directory",
        nargs="?",
        default=default_dir,
        help="Folder with per-DS .xlsx files (default: project UNC path in script).",
    )
    parser.add_argument(
        "--debug-ds",
        metavar="LABEL",
        help='Debug: PrettyTable + UTF-8 txt for one DS_NAME (summary col 1; from xlsx filename, e.g. ...DC66...).',
    )
    parser.add_argument(
        "--subfolders",
        action="store_true",
        help="Include subfolders when scanning xlsx (merge default is True; debug default False).",
    )
    parser.add_argument(
        "--no-open-folder",
        action="store_true",
        help="Do not open Explorer on the output path.",
    )
    parser.add_argument(
        "--include-empty-rows",
        action="store_true",
        help="Dump txt includes empty_row as well as other_row (default: other_row only).",
    )
    args = parser.parse_args()
    target_dir = args.directory
    if args.debug_ds:
        debug_merge_ds_folder_by_ds_name(
            target_dir,
            args.debug_ds,
            include_subfolders=args.subfolders,
            open_folder=not args.no_open_folder,
        )
    else:
        merge_ds_folder(
            target_dir,
            include_subfolders=args.subfolders,
            out_file_prefix="DS_summary",
            open_folder=not args.no_open_folder,
            dump_include_empty_row=args.include_empty_rows,
        )


