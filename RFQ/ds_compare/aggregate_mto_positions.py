import os
from typing import List

from base.base_classes import RowStd, TableComments
import base.t_comm_initial_classes as t_com_init_cls
from base.base_mto import get_mto_std_from_file
from base.base_cheks import check_position_row
from base.base_excel_out import check_color_out
from base.tables_columns import ColNames, ROW_TYPE, CODE, VALUES
import utils.path


def _load_single_mto(file_full_path: str) -> List[RowStd]:
    t_com = TableComments(file_full_path=file_full_path,
                          dir_path="-1",
                          tabel_class=t_com_init_cls.MTO)
    base = get_mto_std_from_file(t_com, dbg=0)
    return base or []


def _iter_mto_files(mto_dir: str) -> List[str]:
    docs = utils.path.get_files_single(mto_dir, endswith=(".xlsx", ".XLSX"))
    if docs == -1:
        return []
    return [d.file_full_path for d in docs]


def _merge_by_code(rows: List[RowStd]) -> List[RowStd]:
    """
    Объединяет строки по коду, суммируя количества.
    Берет первую строку как основу и суммирует VALUES из остальных.
    """
    from collections import defaultdict
    
    # Группируем по коду
    code_groups = defaultdict(list)
    for row in rows:
        code = str(row.get_value(CODE)).strip()
        if code:
            code_groups[code].append(row)
    
    merged = []
    for code, group in code_groups.items():
        if len(group) == 1:
            # Если только одна строка с таким кодом - оставляем как есть
            merged.append(group[0])
        else:
            # Берем первую строку как основу
            base_row = group[0]
            
            # Суммируем количества
            total_value = 0.0
            for row in group:
                try:
                    value = float(row.get_value(VALUES) or 0)
                    total_value += value
                except (ValueError, TypeError):
                    pass
            
            # Устанавливаем суммарное количество
            base_row.el[VALUES].value = total_value
            
            # Добавляем информацию о количестве объединенных строк в комментарий
            base_row.el[VALUES].comment += f"Объединено {len(group)} строк по коду {code}\n"
            
            merged.append(base_row)
    
    return merged


def aggregate_mto_positions(mto_path: str, open_folder: bool = True, merge_by_code=True) -> list[RowStd] | None:
    print("Сбор всех МТО в один Excel (только позиции)")
    print(f"Источник: {mto_path}")

    # 1) Сканируем директорию на файлы МТО
    files = _iter_mto_files(mto_path)
    if not files:
        print(f"По пути <{mto_path}> файлы .xlsx не найдены.")
        return None

    # 2) Загружаем и собираем только позиции
    aggregated: List[RowStd] = []
    for f in files:
        print(f"    Open MTO - {os.path.basename(f)}")
        rows = _load_single_mto(f)
        # фильтруем только строки-позиции
        only_positions = [r for r in rows if check_position_row(r)]
        aggregated.extend(only_positions)

    if not aggregated:
        print("Нет позиционных строк для экспорта")
        return None

    # 3) Объединяем по CODE - суммируем количества
    if merge_by_code:
        print(f"Объединение по кодам: {len(aggregated)} строк -> ", end="")
        aggregated = _merge_by_code(aggregated)
        print(f"{len(aggregated)} строк")
    else:
        print("Объединение по кодам не производится")

    # 4) Готовим выходную директорию
    out_dir = utils.path.get_path_out_dir(mto_path, dir_result_prefix="/__результат_агрегации_МТО_")
    utils.path.make_dir(out_dir)

    # 5) Экспорт в Excel (используем существующий форматтер)
    #    Покажем только стандартные MTO столбцы, скрывая служебные
    col_dict = dict(ColNames.MTO.column_dict)
    file_path = check_color_out(
        aggregated,
        col_for_out_dict=col_dict,
        out_dir=out_dir,
        file_prefix="ALL_MTO_POSITIONS",
        row_not_print_list=(),
        show_row_type=False,
        open_folder=open_folder,
    )

    return aggregated, file_path


if __name__ in {"__main__"}:
    # Пример запуска
    mto_path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\Таблички_графики\14_сравнение с ДС\МТО для закупки"
    mto_aggregated, out_path = aggregate_mto_positions(mto_path, merge_by_code=False)



