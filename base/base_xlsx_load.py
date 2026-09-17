import gc
import os
import shutil
import tempfile
from datetime import datetime
from typing import Any

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.utils import get_column_letter

from base.base_classes import *
from utils.colors import Color
from utils.error_log import ErrorLog


class FormulaCacheMissingError(RuntimeError):
    """В xlsx у ячейки с формулой нет кэшированного значения (data_only → None)."""


def load_rows_from_worksheet(
    ws: Any,
    column_dict: dict[int, str],
    *,
    dbg: int = 0,
    max_empty_rows: int | None = 300,
    skip_empty_rows: bool = False,
    track_source_rows: bool = False,
) -> list[RowStd]:
    """Read mapped columns from an already-open worksheet into raw ``RowStd`` rows.

    Same cell/strike logic as ``load_from_xlsx_file``, without opening the workbook.
    Suitable for ``read_only`` worksheets (single pass).

    Args:
        ws: Open worksheet.
        column_dict: Zero-based worksheet-column mapping.
        dbg: Print an early-stop message when enabled.
        max_empty_rows: Stop after this many consecutive empty mapped rows;
            ``None`` scans the full worksheet.
        skip_empty_rows: Do not materialize rows with no data in mapped columns.
        track_source_rows: Preserve the one-based worksheet row number on each
            returned row as ``_xlsx_source_row``.

    Returns:
        Raw rows mapped to ``RowStd``.
    """
    table_raw_obj: list[RowStd] = []
    empty_rows_count = 0
    for excel_row, row in enumerate(ws.iter_rows(values_only=False), start=1):
        mto_row = RowStd()
        cell_index = 0
        row_has_data = False
        for cell in row:
            if cell_index in column_dict:
                key = column_dict[cell_index]
                v, v_strike = get_strike_unstruck(cell)
                mto_row.el[key] = CheckElement(v, Color.no, v_strike)
                if v or v_strike:
                    row_has_data = True
            cell_index += 1

        if row_has_data:
            empty_rows_count = 0
        else:
            empty_rows_count += 1
            if max_empty_rows is not None and empty_rows_count > max_empty_rows:
                if dbg:
                    print(f"\tПрервано: более {max_empty_rows} пустых строк подряд")
                break

        if skip_empty_rows and not row_has_data:
            continue
        if track_source_rows:
            mto_row._xlsx_source_row = excel_row
        table_raw_obj.append(mto_row)
    return table_raw_obj


def load_from_xlsx_file(t_com: TableComments, dbg=0):
    # Открываем файл...
    wb = None
    table_raw_obj = []
    try:
        if dbg:
            print(f"\tОткрываем |{t_com.file_full_path}| <load_from_xlsx_file>")
        wb = openpyxl.load_workbook(
            t_com.file_full_path.strip(),
            read_only=True,
            data_only=True,
            rich_text=True,
        )
        # grab the <sheet_name> worksheet or any active
        if t_com.sheet_name == -1:
            ws = wb.active
        else:
            try:
                ws = wb[t_com.sheet_name]
            except KeyError as e:
                ErrorLog.add_error(
                    f"Ошибка при открытии листа <{t_com.sheet_name}> в файле <{t_com.file_full_path}>"
                )
                raise e
        table_raw_obj = load_rows_from_worksheet(ws, t_com.column_dict, dbg=dbg)
    finally:
        if wb is not None:
            wb.close()
        del wb
        gc.collect()
    if dbg:
        for i in table_raw_obj:
            for k in i.el:
                if k == "tags":
                    if i.el[k].value != "":
                        print(i.el[k].value)
    return table_raw_obj


def list_raw_to_std_old(list_raw_obj, t_com: TableComments) -> list[RowStd]:
    list_std_obj = []
    for row in list_raw_obj:  # row -> obj RowStd()

        # row_type = RowType.get_row_type(row, column_dict)
        # column_dict[ROW_TYPE] = ROW_TYPE
        # row.el[ROW_TYPE] = CTElement(row_type)

        dic = {}
        for k, v in t_com.column_dict.items():
            try:
                dic[v] = row.el[v].value
            except AttributeError:
                print(f"att = {v}")
                dic[v] = ""
        row = RowStd.get_std_check_row(dic, t_com)
        list_std_obj.append(row)

        # print(row)
    return list_std_obj


def list_raw_to_std(list_raw_obj: list[RowStd], t_com: TableComments) -> list[RowStd]:
    list_std_obj = []
    for raw_row in list_raw_obj:  # row -> obj RowStd()
        dic = {}
        for k, v in raw_row.el.items():
            dic[k] = v
        row = RowStd.get_std_check_row(dic, t_com)
        source_row = getattr(raw_row, "_xlsx_source_row", None)
        if source_row is not None:
            row._xlsx_source_row = source_row
        list_std_obj.append(row)

        # print(row)
    return list_std_obj


def get_strike_unstruck(cell) -> tuple[str, str]:
    dbg_flag = False
    t_normal = ""
    t_struck = ""

    if isinstance(cell.value, CellRichText):
        if dbg_flag:
            print(f"\n1_get_strike_unstruck\n"
                  f"0. <{cell.value}>")

        for rich_text in cell.value:
            if dbg_flag:
                print(f"1.1 <{rich_text}>   (t_normal=\"{t_normal}\"; t_struck=\"{t_struck}\"")
            # Если это текстовый блок текст и шрифт не зачеркнутый
            if isinstance(rich_text, TextBlock):
                rich_text.text = str_remove_n_x000D_(rich_text.text)  #Удаляем переносы строк

                if rich_text.font.strike:
                    t_struck += str(rich_text.text)
                    if dbg_flag:
                        print(f"3.1 <{t_struck}>   (t_normal=\"{t_normal}\"; t_struck=\"{t_struck}\"")
                else:
                    t_normal += str(rich_text.text)
                    if dbg_flag:
                        print(f"3.2 <{t_normal}>   (t_normal=\"{t_normal}\"; t_struck=\"{t_struck}\"")
            # Если это просто текст и шрифт НЕ ЗАЧЕРКНУТЫЙ
            elif isinstance(rich_text, str) and not cell.font.strikethrough:
                t_normal += str_remove_n_x000D_(rich_text)
            # Если это просто текст и шрифт ЗАЧЕРКНУТЫЙ
            elif cell.font.strikethrough:
                rich_text = str_remove_n_x000D_(rich_text)
                t_struck += str(rich_text)
                if dbg_flag:
                    print(f"elif cell.font.strikethrough <{t_normal}> | <{t_struck}>")
            else:
                t_normal += str(cell.value)
                if dbg_flag:
                    print(f"else: <{t_normal}> | <{t_struck}>")

        if dbg_flag:
            print(f"finally - <{t_normal}> | t_struck - <{t_struck}>")

    elif isinstance(cell.value, NoneType):
        pass
    else:
        val = cell.value
        if isinstance(val, float):
            val = round(val, 10)
        if cell.font.strikethrough:
            t_struck += str(val)
        else:
            t_normal += str(val)

    return t_normal, t_struck


def backup_xlsx_to_old_subfolder(file_path: str) -> str:
    """
    Копия xlsx в подпапку old_backup_<timestamp> рядом с файлом (как у снятия зачёркивания).
    """
    file_path = file_path.strip()
    file_dir = os.path.dirname(file_path)
    file_name = os.path.basename(file_path)
    stamp = datetime.now().strftime("%Y.%m.%d_%H.%M.%S")
    backup_dir = os.path.join(file_dir, f"old_backup_{stamp}")
    os.makedirs(backup_dir, exist_ok=True)
    backup_path = os.path.join(backup_dir, file_name)
    shutil.copy2(file_path, backup_path)
    print(f"Резервная копия: {backup_path}")
    return backup_path


def remove_strikethrough_from_xlsx(file_path: str,
                                   sheet_name=None,
                                   backup_suffix="_backup",
                                   create_backup: bool = True) -> str:
    """
    Создаёт резервную копию Excel-файла и удаляет зачёркнутый текст из ячеек,
    сохраняя остальное форматирование (шрифт, размер, цвет, границы и т.д.).

    Args:
        file_path:      полный путь к xlsx-файлу
        sheet_name:     имя листа для обработки (None — все листы)
        backup_suffix:  суффикс для имени резервной копии
        create_backup:  если False — копию не создавать (например, уже сделана до другого шага)

    Returns:
        Путь к созданной резервной копии или пустая строка, если create_backup=False.
    """
    file_path = file_path.strip()
    file_dir = os.path.dirname(file_path)
    backup_path = ""
    if create_backup:
        backup_path = backup_xlsx_to_old_subfolder(file_path)

    wb = openpyxl.load_workbook(file_path, rich_text=True)

    if sheet_name is not None:
        sheets = [wb[sheet_name]]
    else:
        sheets = wb.worksheets

    cells_modified = 0
    for ws in sheets:
        for row in ws.iter_rows():
            for cell in row:
                if _remove_strike_from_cell(cell):
                    cells_modified += 1

    # Сохраняем во временный файл, затем заменяем оригинал — иначе на Windows
    # файл остаётся заблокированным до закрытия программы (openpyxl не освобождает
    # handle при save в тот же путь).
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=file_dir)
    os.close(fd)
    try:
        wb.save(temp_path)
        wb.close()
        del wb
        gc.collect()
        shutil.move(temp_path, file_path)
    except Exception:
        if os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except OSError:
                pass
        raise

    print(f"Удалён зачёркнутый текст из {cells_modified} ячеек -> {file_path}")
    return backup_path


def _remove_strike_from_cell(cell) -> bool:
    """
    Удаляет зачёркнутые фрагменты из одной ячейки.
    После удаления убирает переносы строк и пробелы на краях оставшегося текста.
    Возвращает True если ячейка была изменена.
    """
    if cell.value is None:
        return False

    if isinstance(cell.value, CellRichText):
        kept = []
        has_strike = False
        for part in cell.value:
            if isinstance(part, TextBlock):
                if part.font.strike:
                    has_strike = True
                else:
                    kept.append(part)
            elif isinstance(part, str):
                if cell.font and cell.font.strikethrough:
                    has_strike = True
                else:
                    kept.append(part)

        if not has_strike:
            return False

        kept = _strip_rich_parts(kept)

        if not kept:
            cell.value = None
        elif len(kept) == 1 and isinstance(kept[0], str):
            cell.value = kept[0]
        else:
            cell.value = CellRichText(*kept)
        return True

    else:
        if cell.font and cell.font.strikethrough:
            cell.value = None
            return True
        return False


def _strip_rich_parts(parts: list) -> list:
    """
    Убирает переносы строк (_x000D_, \\n) и пробелы
    с начала первого и конца последнего фрагмента.
    Пустые фрагменты после strip удаляются.
    """
    if not parts:
        return parts

    _STRIP_CHARS = " \n\r\t"
    _X000D = "_x000D_"

    def _clean(text: str, *, left: bool, right: bool) -> str:
        t = text.replace(_X000D, "")
        if left:
            t = t.lstrip(_STRIP_CHARS)
        if right:
            t = t.rstrip(_STRIP_CHARS)
        return t

    def _clean_part(part, *, left: bool, right: bool):
        if isinstance(part, str):
            cleaned = _clean(part, left=left, right=right)
            return cleaned if cleaned else None
        if isinstance(part, TextBlock):
            cleaned = _clean(part.text, left=left, right=right)
            if not cleaned:
                return None
            part.text = cleaned
            return part
        return part

    first_only = len(parts) == 1
    parts[0] = _clean_part(parts[0], left=True, right=first_only)
    if not first_only:
        parts[-1] = _clean_part(parts[-1], left=False, right=True)

    return [p for p in parts if p is not None]


def _atomic_save_workbook_replace(wb, file_path: str) -> None:
    """Сохранить книгу во временный xlsx и атомарно заменить file_path (Windows)."""
    file_path = file_path.strip()
    file_dir = os.path.dirname(file_path)
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=file_dir)
    os.close(fd)
    try:
        wb.save(temp_path)
        wb.close()
        del wb
        gc.collect()
        shutil.move(temp_path, file_path)
    except Exception:
        if os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except OSError:
                pass
        raise


def _merge_slave_coordinates(ws) -> set[str]:
    """Координаты ячеек внутри объединённых диапазонов, кроме левой верхней (master)."""
    slaves: set[str] = set()
    for m_range in ws.merged_cells.ranges:
        top_left = f"{get_column_letter(m_range.min_col)}{m_range.min_row}"
        for row in range(m_range.min_row, m_range.max_row + 1):
            for col in range(m_range.min_col, m_range.max_col + 1):
                addr = f"{get_column_letter(col)}{row}"
                if addr != top_left:
                    slaves.add(addr)
    return slaves


def _cell_has_formula(cell) -> bool:
    if getattr(cell, "data_type", None) == "f":
        return True
    v = cell.value
    if isinstance(v, str) and v.startswith("="):
        return True
    # Объекты формул openpyxl (массив и т.п.)
    if v is not None:
        cls_name = type(v).__name__
        if cls_name in ("ArrayFormula", "DataTableFormula"):
            return True
    return False


def replace_formulas_with_cached_values(
    file_path: str, sheet_name: str | None = None
) -> None:
    """
    Заменяет в книге формулы на кэшированные значения из файла (режим data_only).

    Значения берутся из кэша последнего пересчёта Excel. Если у формулы нет кэша,
    выбрасывает FormulaCacheMissingError (файл на диске не перезаписывается).

    sheet_name: если задан (например "BOE", "BOM", "BOQ"), обрабатывается только
    этот лист — скрытые справочные листы ("Лист1" и т.п.) не трогаются.

    Важно: вызывать до любого openpyxl.save по этому файлу — после сохранения
    openpyxl кэш результатов в xlsx часто теряется, и data_only даёт None.
    """
    file_path = file_path.strip()
    wb_values = None
    wb = None
    replaced = 0
    try:
        wb_values = openpyxl.load_workbook(file_path, data_only=True, rich_text=True)
        wb = openpyxl.load_workbook(file_path, data_only=False, rich_text=True)
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                raise ValueError(
                    f"[1C] Лист {sheet_name!r} не найден в книге: {file_path}"
                )
            worksheets = [wb[sheet_name]]
        else:
            worksheets = list(wb.worksheets)
        for ws in worksheets:
            if ws.title not in wb_values.sheetnames:
                print(f"  [1C] Пропуск листа (нет в data_only-копии): {ws.title}")
                continue
            ws_v = wb_values[ws.title]
            merge_slaves = _merge_slave_coordinates(ws)
            for row in ws.iter_rows():
                for cell in row:
                    if cell.coordinate in merge_slaves:
                        continue
                    if not _cell_has_formula(cell):
                        continue
                    v_src = ws_v[cell.coordinate].value
                    if v_src is None:
                        raise FormulaCacheMissingError(
                            f"[1C] Нет кэша значения у формулы (откройте файл в Excel, "
                            f"полный пересчёт и сохранение; пайплайн «Для 1C» уже вызывает Excel): "
                            f"{file_path} :: {ws.title}!{cell.coordinate}"
                        )
                    cell.value = v_src
                    replaced += 1
        _atomic_save_workbook_replace(wb, file_path)
    except Exception:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
            del wb
            gc.collect()
        raise
    finally:
        if wb_values is not None:
            wb_values.close()
        gc.collect()
    print(f"  [1C] Формулы заменены на значения: {replaced} ячеек -> {file_path}")


def remove_hidden_sheets_for_1c(file_path: str, doc_type: str) -> None:
    """
    Удаляет все скрытые и очень скрытые листы (экспорт «Для 1C»).
    Должен вызываться после замены формул на значения на основном листе.
    """
    file_path = file_path.strip()
    wb = None
    try:
        wb = openpyxl.load_workbook(file_path, rich_text=True)
        to_remove = [
            ws
            for ws in wb.worksheets
            if ws.sheet_state in ("hidden", "veryHidden")
        ]
        if not to_remove:
            wb.close()
            gc.collect()
            print(f"  [1C] Скрытые листы для удаления не найдены: {file_path}")
            return

        survivors = [ws for ws in wb.worksheets if ws not in to_remove]
        if not survivors:
            wb.close()
            gc.collect()
            raise ValueError(
                f"Нельзя удалить все листы: {file_path} (doc_type={doc_type})"
            )

        if wb.active in to_remove:
            wb.active = survivors[0]

        removed_titles = ", ".join(ws.title for ws in to_remove)
        for ws in to_remove:
            wb.remove(ws)

        _atomic_save_workbook_replace(wb, file_path)
    except Exception:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        gc.collect()
        raise
    print(f"  [1C] Удалены скрытые листы ({len(to_remove)}): {removed_titles} -> {file_path}")
