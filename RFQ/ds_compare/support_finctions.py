import re
import sys
from typing import Dict, List, Tuple

import RFQ.ds_compare.get_mto_list_from_ds
import RFQ.ds_compare.load_mto
import base.base_mto
import base.t_comm_initial_classes
import utils.path
from base.base_excel_out import *
import pandas as pd
from prettytable import PrettyTable

from pdf_parsing_v2_engine.document import V2Document


def _try_float_value(val) -> Tuple[bool, float]:
    """Пытается преобразовать значение в float. Возвращает (успех, значение)."""
    from RFQ.ds_compare.ds_quantity_parse import try_parse_quantity

    return try_parse_quantity(val)


def validate_values_float(ds_data: List[RowStd], mto_dict: Dict[str, List[RowStd]]) -> List[Tuple[str, str, str, int]]:
    """
    Проверяет, что все значения в колонке VALUES могут быть преобразованы в float.
    Возвращает список ошибок: [(spec_name, invalid_value, code, row_index), ...]
    """
    errors = []

    from RFQ.ds_compare.ds_quantity_parse import read_raw_quantity

    # Проверка MTO
    for spec_name, mto_rows in mto_dict.items():
        for row_index, row in enumerate(mto_rows, start=1):
            if not isinstance(row, RowStd) or row.row_type != RowType.position_row:
                continue
            val = read_raw_quantity(row, VALUES)
            ok, _ = _try_float_value(val)
            if not ok:
                code = row.get_value(CODE) or ""
                errors.append((spec_name, str(val), str(code), row_index))

    # Проверка DS
    for row_index, row in enumerate(ds_data, start=1):
        if not isinstance(row, RowStd) or row.row_type != RowType.position_row:
            continue
        val = read_raw_quantity(row, VALUES)
        ok, _ = _try_float_value(val)
        if not ok:
            try:
                spec_name = V2Document.from_file_path(row.get_value(DS_SPECIFICATION)).doc_Short_Title
            except (AttributeError, TypeError):
                spec_name = "?"
            code = row.get_value(CODE) or ""
            errors.append((spec_name, str(val), str(code), row_index))

    return errors


def validate_and_exit_if_errors(ds_data: List[RowStd], mto_dict: Dict[str, List[RowStd]]) -> None:
    """
    Проверяет данные на недопустимые значения в колонке количества (VALUES).
    При обнаружении ошибок выводит таблицу и завершает программу.
    """
    errors = validate_values_float(ds_data, mto_dict)
    if not errors:
        return

    # Уникальные спецификации для таблицы
    table = PrettyTable()
    table.field_names = ["№", "Строка", "Имя спецификации", "Недопустимое значение", "Код"]
    for col in table.field_names:
        table.align[col] = "l"

    for n, (spec_name, invalid_val, code, row_index) in enumerate(errors, start=1):
        table.add_row([n, row_index, spec_name, invalid_val, code])

    print("\n" + "=" * 80)
    print("ОШИБКА: Обнаружены недопустимые значения в колонке количества (VALUES)")
    print("Значение должно быть числом. Примеры ошибок: формулы Excel (=H51), текст и т.п.")
    print("=" * 80)
    print(table)
    print("=" * 80)
    print("\nНеобходимо откорректировать данные MTO/DS и повторить запуск.")
    print("=" * 80 + "\n")
    sys.exit(1)


def load_ds_data(
    ds_path: str,
    summ_ds=False,
    force_update: bool = False,
    use_cache: bool = True,
) -> List[RowStd]:
    """
    Загружает данные из DS файла с поддержкой кэширования
    
    Args:
        ds_path: Путь к DS файлу
        summ_ds: Флаг для определения типа DS файла
        force_update: Принудительное обновление данных из файла (игнорирует кэш)
        use_cache: Если False — не читать и не писать кэш (полностью отключено)
    
    Returns:
        Список объектов RowStd с данными DS
    """
    from utils.cache_utils import cache_manager
    from RFQ.ds_compare.ds_units_normalize import normalize_rows_units
    
    # Создаем уникальный ключ кэша на основе параметров
    cache_key = f"{ds_path}_{summ_ds}"
    
    from RFQ.ds_compare.ds_units_normalize import normalize_rows_units

    # Проверяем кэш, если не требуется принудительное обновление
    if use_cache and not force_update and cache_manager.is_cache_valid(ds_path):
        cached_data = cache_manager.load_from_cache(ds_path)
        if cached_data is not None:
            normalize_rows_units(cached_data, (UNITS,))
            print(f"    Данные загружены из кэша - {len(cached_data)} строк")
            return cached_data
    
    # Загружаем данные из файла
    print(f"    Загрузка данных из файла - {ds_path}")
    if not summ_ds:
        ds_t_com = TableComments(
            file_full_path=ds_path,
            dir_path="-1",
            tabel_class=base.t_comm_initial_classes.DsSpecification
        )
        print(f"    Open DS (not summ_ds) - {ds_path}")
    else:
        ds_t_com = TableComments(
            file_full_path=ds_path,
            dir_path="-1",
            tabel_class=base.t_comm_initial_classes.ListDsSpecification,
        )
        print(f"    Open DS (summ_ds) - {ds_path}")

    ds_base_full = base.base_mto.get_std_from_excel_file(ds_t_com)
    print(f"Загруженны {len(ds_base_full)} строк из {ds_path}")

    # Проверка на пустые столбцы
    flag = False
    for row in ds_base_full:
        v = row.get_value(DS_NUMBER)
        if v != "" and v is not None:
            flag = True
    if not flag:
        print(f"Ошибка формата - первый столбец пустой в {ds_path}")
        exit(0)

    def find_ds_pattern(text):
        """
        Ищет в строке шаблон: кириллические «ДС» и номер из 1–2 цифр (без обязательного пробела).

        Args:
            text (str): входная строка для поиска (обычно имя файла).

        Returns:
            str: первое совпадение, нормализованное (без пробелов внутри метки).

        Raises:
            IndexError: если совпадений нет (вызывающий код должен обрабатывать или править шаблон).
        """
        pattern = r'ДС\d{1,2}'
        matches = re.findall(pattern, text)
        out_str = matches[0].strip().replace(' ', '')
        return out_str

    if not summ_ds:
        ds_name = find_ds_pattern(ds_t_com.file_name)
        for row in ds_base_full:
            row.el[DS_NAME].value = ds_name

    del_list = [DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT, DS_TOTAL_PRICE_EXCL_VAT, DS_VAT_AMOUNT, DS_TOTAL_PRICE_INCL_VAT]
    for row in ds_base_full:
        for attr in del_list:
            row.el[attr].value = ""

    # RowStd.set_color_by_row_type_in_list(ds_base_full)

    normalize_rows_units(ds_base_full, (UNITS,))

    if use_cache:
        cache_manager.save_to_cache(ds_path, ds_base_full)

    return ds_base_full


def load_rfq_tpk_data(
    rfq_path: str,
    *,
    dump_non_position_dir: str | None = None,
) -> list[RowStd]:
    """Load RFQ TPK workbook into ``RowStd`` rows via ``get_std_from_excel_file``.

    Args:
        rfq_path: Path to RFQ TPK xlsx file.
        dump_non_position_dir: If set, write ``RFQ_TPK_not_position_rows.txt`` there
            (all rows except ``position_row``: empty_row, head_row, other_row, …).

    Returns:
        Parsed rows with ``row_type`` assigned by ``rfq_tpk.get_row_type``.
    """
    rfq_path = str(rfq_path or "").strip()
    if not rfq_path:
        return []
    rfq_t_com = TableComments(
        file_full_path=rfq_path,
        dir_path="-1",
        tabel_class=base.t_comm_initial_classes.RFQTPK,
    )
    print(f"    Open RFQ TPK - {rfq_path}")
    rfq_rows = base.base_mto.get_std_from_excel_file(rfq_t_com)
    position_count = sum(1 for row in rfq_rows if row.row_type == RowType.position_row)
    from RFQ.ds_compare.ds_units_normalize import normalize_rows_units

    normalize_rows_units(rfq_rows, (UNITS,))
    from RFQ.ds_compare.ds_quantity_parse import validate_rows_quantities_or_exit

    validate_rows_quantities_or_exit(
        rfq_rows,
        headline="ОШИБКА: недопустимое количество (VALUES) в RFQ TPK",
        file_hint=rfq_path,
        columns=(VALUES,),
        only_position_row=True,
        normalize=True,
    )
    print(
        f"RFQ TPK: загружено {len(rfq_rows)} строк "
        f"({position_count} position_row) из {rfq_path}"
    )
    if dump_non_position_dir:
        from RFQ.ds_compare.ds_rfq_non_position_dump import write_rfq_non_position_rows_txt

        write_rfq_non_position_rows_txt(rfq_rows, dump_non_position_dir)
    return rfq_rows


def load_mto_data(
    ds_data: List[RowStd],
    mto_path: str,
    flat_structure: bool = False,
    mto_use_cache: bool = True,
    mto_force_update: bool = False,
    ds_source_path: str | None = None,
) -> tuple[dict, dict, dict]:
    """Загружает MTO данные.

    Returns:
        mto_dict, spec_dict, load_audit (spec key → path / cache meta).
    """
    print("Загружает MTO данные")
    err_dir: str | None = None
    if ds_source_path:
        err_dir = utils.path.get_path_from_file_path(ds_source_path)
    spec_dict = RFQ.ds_compare.get_mto_list_from_ds.get_mto_list_from_ds(
        ds_data,
        mto_path,
        flat_structure=flat_structure,
        error_report_out_dir=err_dir,
        ds_source_path=ds_source_path,
    )
    mto_dict, load_audit = RFQ.ds_compare.load_mto.load_mto_by_dict(
        spec_dict,
        print_mto_to_console=False,
        use_mto_cache=mto_use_cache,
        mto_force_update=mto_force_update,
    )
    return mto_dict, spec_dict, load_audit


def load_replacement_table(file_path: str,
                           print_replacement_table_info=False) -> Dict[str, List[Tuple[str, str]]]:
    """
    Загружает таблицу замен кодов из XLSX файла
    """
    try:
        df = pd.read_excel(file_path, sheet_name="Коды_замен")

        # Проверяем наличие необходимых столбцов
        required_columns = ['СТАРЫЙ_КОД', 'НОВЫЙ_КОД', 'СТАТУС']
        for col in required_columns:
            if col not in df.columns:
                print(f"Ошибка: Отсутствует столбец {col} во вкладке 'Коды_замен'")
                return {}

        replacement_dict = {}

        for _, row in df.iterrows():
            old_code = str(row['СТАРЫЙ_КОД']).strip()
            new_code = str(row['НОВЫЙ_КОД']).strip()
            status = str(row['СТАТУС']).strip().upper() if pd.notna(row['СТАТУС']) else 'НЕ_ПРОВЕРЕН'

            if not old_code or not new_code:
                continue

            if old_code not in replacement_dict:
                replacement_dict[old_code] = []

            replacement_dict[old_code].append((new_code, status))

        # Вывод информации в консоль
        if print_replacement_table_info:
            _print_replacement_table_info(replacement_dict)

        return replacement_dict

    except Exception as e:
        print(f"Ошибка загрузки таблицы замен: {e}")
        raise e


def _print_replacement_table_info(replacement_dict: Dict[str, List[Tuple[str, str]]]):
    """
    Выводит информацию о таблице замен в консоль
    """
    table = PrettyTable()
    table.field_names = ["Старый код", "Новый код", "Статус"]
    table.align = "l"

    for old_code, replacements in replacement_dict.items():
        for new_code, status in replacements:
            table.add_row([old_code, new_code, status])

    print("\nТаблица замен кодов:")
    print(table)

def ds_codr_vs_base_google (ds_base: list[RowStd], google_base: list[RowStd]):
    google_base_set = set("")

    for row in google_base:
        google_base_set.add(str(row.el[CODE].value).strip())
    for row in ds_base:
        code = str(row.el[CODE_2].value).strip()
        if not code:
            continue
        value_2_el = row.el[CODE_2]
        if code not in google_base_set:
            Color.set_el_color(value_2_el, Color.red)
            comment_text = f"Нет в ГуглБазе\n"
            value_2_el.comment = value_2_el.comment + comment_text
            row.el[ANNOTATION_3].value = comment_text
            Color.set_el_color(row.el[ANNOTATION_3], Color.red)


def check_comparison_results(
    ds_data: List[RowStd],
    spec_dict,
    mto_use_cache: bool = True,
    mto_force_update: bool = False,
):
    """
    Проверяет результаты сравнения DS и MTO по спецификациям
    Сравнивает суммы количеств из VALUES_2 (распределенные из MTO) с исходными количествами в MTO
    
    Особенности обработки агрегированных строк:
    - Если в комментарии VALUES_2 найден паттерн [AGG_CODES: код1:количество1,код2:количество2,...],
      то количество из VALUES_2 распределяется по исходным кодам согласно указанным количествам
    - Это позволяет избежать ложных расхождений при агрегации нескольких MTO строк в одну DS строку
    - Парсинг устойчив к наличию другой информации в комментарии
    """
    print("\n" + "=" * 60)
    print("ПРОВЕРКА РЕЗУЛЬТАТОВ СРАВНЕНИЯ")
    print("=" * 60)
    mto_dict, _load_audit = RFQ.ds_compare.load_mto.load_mto_by_dict(
        spec_dict,
        print_mto_to_console=False,
        use_mto_cache=mto_use_cache,
        mto_force_update=mto_force_update,
    )

    # Собираем данные по спецификациям
    spec_results = {}
    code_quantity_results = {}  # Для проверки по код->количество

    # 1. Суммируем распределенные количества из DS (VALUES_2)
    for ds_row in ds_data:
        if isinstance(ds_row, RowStd) and ds_row.row_type == RowType.position_row:
            spec_name = f"{ds_row.get_value(DS_TITLE)}-{ds_row.get_value(DS_SYSTEM)}"
            if not spec_name:
                continue

            if spec_name not in spec_results:
                spec_results[spec_name] = {
                    'ds_distributed': 0.0,  # Сумма распределенных количеств из MTO
                    'mto_original': 0.0,  # Сумма исходных количеств в MTO
                    'ds_rows_count': 0,  # Количество строк DS для этой спецификации
                    'mto_rows_count': 0  # Количество строк MTO для этой спецификации
                }

            # Суммируем распределенное количество
            distributed_amount = float(ds_row.get_value(VALUES_2) or 0)
            spec_results[spec_name]['ds_distributed'] += distributed_amount
            spec_results[spec_name]['ds_rows_count'] += 1
            
            # Собираем данные по кодам для детальной проверки
            code = ds_row.get_value(CODE_2)
            if code:
                if spec_name not in code_quantity_results:
                    code_quantity_results[spec_name] = {}
                
                # Проверяем комментарий VALUES_2 на наличие информации об агрегированных кодах
                values_2_comment = ds_row.el[VALUES_2].comment or ""
                aggregated_codes_info = _parse_aggregated_codes_from_comment(values_2_comment)
                
                if aggregated_codes_info:
                    # Для агрегированной строки распределяем количество по исходным кодам согласно комментарию
                    for source_code, amount in aggregated_codes_info.items():
                        if source_code not in code_quantity_results[spec_name]:
                            code_quantity_results[spec_name][source_code] = {
                                'ds_distributed': 0.0,
                                'mto_original': 0.0,
                                'ds_rows': [],
                                'mto_rows': []
                            }
                        code_quantity_results[spec_name][source_code]['ds_distributed'] += amount
                        code_quantity_results[spec_name][source_code]['ds_rows'].append(ds_row)
                else:
                    # Обычная строка - обрабатываем как раньше
                    if code not in code_quantity_results[spec_name]:
                        code_quantity_results[spec_name][code] = {
                            'ds_distributed': 0.0,
                            'mto_original': 0.0,
                            'ds_rows': [],
                            'mto_rows': []
                        }
                    code_quantity_results[spec_name][code]['ds_distributed'] += distributed_amount
                    code_quantity_results[spec_name][code]['ds_rows'].append(ds_row)

    # 2. Суммируем исходные количества из MTO
    for spec_name, mto_rows in mto_dict.items():
        if spec_name not in spec_results:
            spec_results[spec_name] = {
                'ds_distributed': 0.0,
                'mto_original': 0.0,
                'ds_rows_count': 0,
                'mto_rows_count': len(mto_rows)
            }

        total_mto_amount = sum(float(row.get_value(VALUES) or 0) for row in mto_rows)
        spec_results[spec_name]['mto_original'] = total_mto_amount
        spec_results[spec_name]['mto_rows_count'] = len(mto_rows)
        
        # Собираем данные по кодам из MTO
        if spec_name not in code_quantity_results:
            code_quantity_results[spec_name] = {}
            
        for mto_row in mto_rows:
            code = mto_row.get_value(CODE)
            if code:
                if code not in code_quantity_results[spec_name]:
                    code_quantity_results[spec_name][code] = {
                        'ds_distributed': 0.0,
                        'mto_original': 0.0,
                        'ds_rows': [],
                        'mto_rows': []
                    }
                mto_amount = float(mto_row.get_value(VALUES) or 0)
                code_quantity_results[spec_name][code]['mto_original'] += mto_amount
                code_quantity_results[spec_name][code]['mto_rows'].append(mto_row)

    # 3. Анализируем результаты
    total_discrepancy = 0.0
    specs_with_issues = []

    from prettytable import PrettyTable
    from colorama import Fore, Style  # Для цветового оформления (опционально)

    print("\nРезультаты по спецификациям:")

    # Создаем таблицу
    table = PrettyTable()
    table.field_names = ["Спецификация", "MTO исходное", "DS распределено", "Разница", "Статус"]

    # Настраиваем выравнивание
    table.align["Спецификация"] = "l"
    table.align["MTO исходное"] = "r"
    table.align["DS распределено"] = "r"
    table.align["Разница"] = "r"
    table.align["Статус"] = "l"

    for spec_name, data in spec_results.items():
        original = data['mto_original']
        distributed = data['ds_distributed']
        difference = abs(original - distributed)

        # Определяем статус
        if original == 0 and distributed == 0:
            status = "Нет данных"
            status_color = ""
        elif abs(difference) < 0.001:
            status = "✓ СОВПАДАЕТ"
            status_color = Fore.GREEN
        else:
            status = "✗ РАСХОЖДЕНИЕ"
            total_discrepancy += difference
            specs_with_issues.append(spec_name)
            status_color = Fore.RED

        # Добавляем строку в таблицу
        table.add_row([
            spec_name,
            f"{original:.2f}",
            f"{distributed:.2f}",
            f"{difference:.2f}",
            f"{status_color}{status}{Style.RESET_ALL}" if status_color else status
        ])

    # Выводим таблицу
    print(table)

    # 4. Детальная проверка по код->количество
    print("\n" + "=" * 60)
    print("ДЕТАЛЬНАЯ ПРОВЕРКА ПО КОДАМ")
    print("=" * 60)
    
    code_issues = []
    total_code_discrepancy = 0.0
    
    for spec_name, codes_data in code_quantity_results.items():
        spec_code_issues = []
        for code, data in codes_data.items():
            mto_original = data['mto_original']
            ds_distributed = data['ds_distributed']
            difference = abs(mto_original - ds_distributed)
            
            if mto_original > 0 or ds_distributed > 0:  # Только если есть данные
                if abs(difference) >= 0.001:  # Есть расхождение
                    spec_code_issues.append({
                        'code': code,
                        'mto_original': mto_original,
                        'ds_distributed': ds_distributed,
                        'difference': difference
                    })
                    total_code_discrepancy += difference
        
        if spec_code_issues:
            code_issues.append({
                'spec_name': spec_name,
                'issues': spec_code_issues
            })
    
    if code_issues:
        print(f"Найдены расхождения по кодам в {len(code_issues)} спецификациях:")
        for spec_issue in code_issues:
            print(f"\nСпецификация: {spec_issue['spec_name']}")
            for issue in spec_issue['issues']:
                print(f"  Код {issue['code']}: MTO={issue['mto_original']:.2f} -> DS={issue['ds_distributed']:.2f} "
                      f"(разница: {issue['difference']:.2f})")
    else:
        print("✓ Все коды совпадают по количеству!")

    # 5. Сводная статистика
    print("\n" + "=" * 60)
    print("СВОДНАЯ СТАТИСТИКА")
    print("=" * 60)

    total_specs = len(spec_results)
    problem_specs = len(specs_with_issues)
    problem_code_specs = len(code_issues)

    print(f"Всего спецификаций: {total_specs}")
    print(f"Спецификаций с расхождениями (общие): {problem_specs}")
    print(f"Спецификаций с расхождениями по кодам: {problem_code_specs}")
    print(f"Общая сумма расхождений (общие): {total_discrepancy:.2f}")
    print(f"Общая сумма расхождений (по кодам): {total_code_discrepancy:.2f}")

    if problem_specs > 0:
        print(f"\nСпецификации с общими расхождениями:")
        for spec in specs_with_issues:
            data = spec_results[spec]
            print(f"  - {spec}: MTO={data['mto_original']:.2f} -> DS={data['ds_distributed']:.2f} "
                  f"(разница: {abs(data['mto_original'] - data['ds_distributed']):.2f})")



    # 6. Итоговый вывод
    print("\n" + "=" * 60)
    print("ИТОГОВЫЙ ВЫВОД")
    print("=" * 60)

    if total_discrepancy == 0:
        print("✅ ВСЕ ПРОВЕРОК ПРОЙДЕНЫ УСПЕШНО!")
        print("Распределение количеств из MTO в DS выполнено корректно.")
    else:
        print("⚠️  ОБНАРУЖЕНЫ РАСХОЖДЕНИЯ!")

    return spec_results

def load_ds_data_list(ds_path: str, force_update: bool = False) -> List[RowStd]:
    """
    Загружает данные из DS файла с поддержкой кэширования
    
    Args:
        ds_path: Путь к DS файлу
        force_update: Принудительное обновление данных из файла (игнорирует кэш)
    
    Returns:
        Список объектов RowStd с данными DS
    """
    from utils.cache_utils import cache_manager
    
    # Проверяем кэш, если не требуется принудительное обновление
    if not force_update and cache_manager.is_cache_valid(ds_path):
        cached_data = cache_manager.load_from_cache(ds_path)
        if cached_data is not None:
            print(f"    Данные загружены из кэша - {len(cached_data)} строк")
            return cached_data
    
    # Загружаем данные из файла
    print(f"    Загрузка данных из файла - {ds_path}")
    ds_t_com = TableComments(
        file_full_path=ds_path,
        dir_path="-1",
        tabel_class=base.t_comm_initial_classes.ListDsSpecificationVsMto,
    )
    print(f"    Open DS List - {ds_path}")

    ds_base_full = base.base_mto.get_std_from_excel_file(ds_t_com)
    print(f"\tФайл ДС список загружен - {len(ds_base_full)} строк")

    # Сохраняем в кэш
    cache_manager.save_to_cache(ds_path, ds_base_full)

    return ds_base_full


def _parse_aggregated_codes_from_comment(comment: str) -> dict:
    """
    Парсит комментарий для извлечения информации об агрегированных кодах и количествах
    
    Ищет паттерн [AGG_CODES: код1:количество1,код2:количество2,...]
    Возвращает словарь {код: суммарное_количество}
    
    Особенности:
    - Если один код встречается несколько раз, количества суммируются
    - Пример: [AGG_CODES: BCC0000550:3.00,BCC0000550:5712.00] -> {'BCC0000550': 5715.00}
    
    Args:
        comment: Комментарий для парсинга
        
    Returns:
        Словарь с кодами и суммарными количествами, или пустой словарь если информация не найдена
    """
    import re
    
    if not comment:
        return {}
    
    # Ищем паттерн [AGG_CODES: ...]
    pattern = r'\[AGG_CODES:\s*([^\]]+)\]'
    match = re.search(pattern, comment)
    
    if not match:
        return {}
    
    codes_data_str = match.group(1)
    result = {}
    
    try:
        # Разбираем строку вида "код1:количество1,код2:количество2,..."
        for code_amount_pair in codes_data_str.split(','):
            code_amount_pair = code_amount_pair.strip()
            if ':' in code_amount_pair:
                code, amount_str = code_amount_pair.split(':', 1)
                code = code.strip()
                amount_str = amount_str.strip()
                
                try:
                    amount = float(amount_str)
                    # Суммируем количества для одинаковых кодов
                    if code in result:
                        result[code] += amount
                    else:
                        result[code] = amount
                except ValueError:
                    # Пропускаем некорректные значения
                    continue
    except Exception:
        # В случае любой ошибки возвращаем пустой словарь
        return {}
    
    return result