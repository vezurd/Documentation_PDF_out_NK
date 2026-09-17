"""
Этап 4_3: Определение статуса позиции для строк RFP
Проверка наличия позиций в MTO по CODE и тегам
"""

from typing import List, Dict
from collections import defaultdict
from datetime import datetime
import os
import builtins

from base.base_classes import RowStd, RowType
from base.tables_columns import (
    CODE, DS_TITLE, POSITION_STATUS, MATCH_STATUS, VALUES,
    TAG_MTO, CODE_MTO, NAME_MTO, TYPE_MARK_MTO, NUMBERS_MTO, VALUES_MTO, NAME, TYPE_MARK, NUMBERS
)
from RFQ.ds_compare.ds_units_normalize import copy_mto_units_fields
from RFQ.tags_rfp_compare.rfp_supply_status import is_excluded_from_supply
from RFQ.tags_rfp_compare.step4.step4_2_match_rfp_with_mto import (
    add_mto_data_to_rfp_row,
    _safe_float,
    remaining_mto_output_qty,
)


def assign_position_status(
    rfp_rows: List[RowStd],
    mto_data: Dict[str, List[RowStd]],
    result_dir: str = None,
    debug: bool = False,
    debug_tag: List[str] = None,
    debug_code: List[str] = None,
    debug_title_system: List[str] = None,
    debug_log: List[str] = None,
    verbose_progress_messages: bool = True):
    """
    Определяет статус позиции для всех строк RFP
    
    Статусы:
    - "Исключен" - тега в RFP нет, т.к. в MTO не хватило позиций с таким CODE
    - "Не протегирован в МТО" - тега в RFP нет, т.к. в МТО для позиции с таким CODE не указан ТЕГ,
      но позиция есть и количество хватает. При этом вставляются данные из МТО с поправкой на количество.
    - "Не протегирован в RFP" - в строке RFP нет тегов
    
    Args:
        rfp_rows: Список строк RFP для обработки
        mto_data: Словарь title_mark -> список строк MTO
        result_dir: Папка для сохранения debug-лога (txt)
        debug: Флаг отладки
        debug_tag: Список тегов для отладки (AND с другими, OR внутри списка)
        debug_code: Список кодов для отладки (AND с другими, OR внутри списка)
        debug_title_system: Список title_system для отладки (AND с другими, OR внутри списка)
    """
    debug_lines: List[str] = []
    _builtin_print = builtins.print

    def _debug_print(*args, should_log: bool = True, **kwargs):
        """
        Выводит отладочную информацию в консоль и опционально в debug_log.
        
        Args:
            *args: Аргументы для print
            should_log: Если True, записывает в debug_log. Если False, только в консоль.
            **kwargs: Kwargs для print (sep, end)
        """
        _builtin_print(*args, **kwargs)
        sep = kwargs.get("sep", " ")
        end = kwargs.get("end", "\n")
        if debug:
            debug_lines.append(sep.join(str(arg) for arg in args) + end)
        if debug_log is not None and should_log:
            debug_log.append(sep.join(str(arg) for arg in args) + end)

    has_debug_filters = bool(debug_tag or debug_code or debug_title_system)

    def _matches_debug_filters(title_mark: str, code: str, tags: List[str]) -> bool:
        """
        Проверяет совпадение с фильтрами отладки.
        AND между фильтрами, OR внутри каждого списка.
        """
        if not has_debug_filters:
            return False
        # debug_title_system: проверяем, что title_mark входит в список
        if debug_title_system and title_mark not in debug_title_system:
            return False
        # debug_code: проверяем, что code входит в список
        if debug_code and code not in debug_code:
            return False
        # debug_tag: проверяем, что хотя бы один тег из tags входит в debug_tag
        if debug_tag and (not tags or not any(t in debug_tag for t in tags)):
            return False
        return True

    def _matches_debug_group(title_mark: str, code: str, group_rows: List[RowStd]) -> bool:
        """
        Проверяет совпадение группы строк с фильтрами отладки.
        AND между фильтрами, OR внутри каждого списка.
        """
        if not has_debug_filters:
            return False
        # debug_title_system: проверяем, что title_mark входит в список
        if debug_title_system and title_mark not in debug_title_system:
            return False
        # debug_code: проверяем, что code входит в список
        if debug_code and code not in debug_code:
            return False
        # debug_tag: проверяем, что хотя бы один тег из любой строки группы входит в debug_tag
        if debug_tag:
            for group_row in group_rows:
                row_tags = group_row.get_tags_list()
                if row_tags and any(t in debug_tag for t in row_tags):
                    return True
            return False
        return True

    if verbose_progress_messages:
        _builtin_print("\n  Определение статуса позиции...")
    
    # Создаем индекс MTO строк по title_mark, CODE и наличию тегов
    # Структура: title_mark -> CODE -> (с_тегами: список строк, без_тегов: список строк)
    mto_by_code: Dict[str, Dict[str, Dict[str, List[RowStd]]]] = defaultdict(
        lambda: defaultdict(lambda: {"with_tags": [], "without_tags": []})
    )
    
    for title_mark, mto_rows in mto_data.items():
        for mto_row in mto_rows:
            if mto_row.row_type != RowType.position_row:
                continue
            
            code = mto_row.get_value(CODE) or ""
            if not code:
                continue
            
            mto_tags = mto_row.get_tags_list()
            if mto_tags:
                mto_by_code[title_mark][code]["with_tags"].append(mto_row)
            else:
                mto_by_code[title_mark][code]["without_tags"].append(mto_row)
    
    # Отладочная информация: вывод индекса MTO для debug_code
    if debug_code:
        # Определяем, нужно ли логировать в debug_log (все non-None фильтры должны совпадать)
        # Для заголовка: должен существовать хотя бы один title_mark, который пройдет фильтр
        has_matching_title = any(
            (not debug_title_system or tm in debug_title_system) and any(dc in codes for dc in debug_code)
            for tm, codes in mto_by_code.items()
        )
        should_log_header = has_matching_title and (not debug_tag)  # Если debug_tag задан, нужна более детальная проверка
        _debug_print(f"\n  [DEBUG_CODE={debug_code}] Индекс MTO по CODE:", should_log=should_log_header)
        for title_mark, codes_dict in mto_by_code.items():
            if debug_title_system and title_mark not in debug_title_system:
                continue
            # Проверяем каждый код из списка debug_code
            for dc in debug_code:
                if dc not in codes_dict:
                    continue
                code_data = codes_dict[dc]
                # Проверяем, есть ли совпадение по тегу (если debug_tag задан)
                matching_tags_in_with = any(
                    any(t in debug_tag for t in mto_row.get_tags_list())
                    for mto_row in code_data['with_tags']
                ) if debug_tag else True
                should_log_section = matching_tags_in_with or (not debug_tag)
                _debug_print(f"    title_mark={title_mark}, CODE={dc}:", should_log=should_log_section)
                _debug_print(f"      С тегами: {len(code_data['with_tags'])} позиций", should_log=should_log_section)
                for idx, mto_row in enumerate(code_data['with_tags'], 1):
                    mto_tags = mto_row.get_tags_list()
                    mto_qty = mto_row.get_value(VALUES) or 0
                    # Логируем только если тег совпадает (или debug_tag не задан)
                    should_log_row = (not debug_tag) or any(t in debug_tag for t in mto_tags)
                    _debug_print(f"        [{idx}] Теги={mto_tags}, VALUES={mto_qty}", should_log=should_log_row)
                _debug_print(f"      Без тегов: {len(code_data['without_tags'])} позиций", should_log=should_log_section and not debug_tag)
                for idx, mto_row in enumerate(code_data['without_tags'], 1):
                    mto_qty = mto_row.get_value(VALUES) or 0
                    # Без тегов - логируем только если debug_tag не задан
                    _debug_print(f"        [{idx}] VALUES={mto_qty}", should_log=not debug_tag)
    
    # Группируем строки RFP по title_mark и CODE для подсчета требуемого количества
    # Структура: (title_mark, code) -> список строк RFP с тегами, которые не были сопоставлены
    rfp_by_code: Dict[tuple[str, str], List[RowStd]] = defaultdict(list)
    
    for rfp_row in rfp_rows:
        if rfp_row.row_type != RowType.position_row:
            continue
        
        rfp_tags = rfp_row.get_tags_list()
        title_mark = rfp_row.get_value(DS_TITLE)
        code = rfp_row.get_value(CODE) or ""
        
        # Проверка на совпадение для отладки
        is_debug_row = _matches_debug_filters(title_mark, code, rfp_tags)
        
        if is_debug_row:
            _debug_print(f"\n  [DEBUG][assign_position_status][find_row] Найдена строка для отладки:")
            _debug_print(f"    CODE={code}, title_mark={title_mark}, теги={rfp_tags}")
            _debug_print(f"    VALUES={rfp_row.get_value(VALUES)}, MATCH_STATUS={rfp_row.el[MATCH_STATUS].value}")
            if debug_title_system:
                _debug_print(f"    debug_title_system={debug_title_system}")

        if is_excluded_from_supply(rfp_row):
            if is_debug_row:
                _debug_print(f"    [DEBUG][assign_position_status][skip] Пропуск: исключен из поставки")
            continue
        
        # Определяем статус позиции
        if not rfp_tags:
            # В строке RFP нет тегов
            rfp_row.el[POSITION_STATUS].value = ""
            if is_debug_row:
                _debug_print(f"    [DEBUG][assign_position_status][skip] Пропуск: нет тегов в RFP")
            continue
        
        if not title_mark or not code:
            if is_debug_row:
                _debug_print(f"    [DEBUG][assign_position_status][skip] Пропуск: нет title_mark или code")
            continue
        
        # Проверяем, была ли строка сопоставлена
        match_status = rfp_row.el[MATCH_STATUS].value
        if match_status in {"Тег сопоставлен", "Тег в МТО заменен", "Добавлен из МТО"}:
            # Строка уже сопоставлена, статус позиции не устанавливаем
            if is_debug_row:
                _debug_print(f"    [DEBUG][assign_position_status][skip] Пропуск: строка уже сопоставлена (MATCH_STATUS='{match_status}')")
            continue
        
        # Строка не была сопоставлена - добавляем в группу для проверки
        rfp_by_code[(title_mark, code)].append(rfp_row)
        if is_debug_row:
            _debug_print(f"    [DEBUG][assign_position_status][add_to_group] Добавлена в группу для проверки: (title_mark={title_mark}, code={code})")
    
    # Обрабатываем каждую группу строк RFP с одинаковым CODE
    assigned_count = 0
    for (title_mark, code), rfp_group in rfp_by_code.items():
        # Проверка на совпадение для отладки
        is_debug_group = _matches_debug_group(title_mark, code, rfp_group)
        
        if is_debug_group:
            _debug_print(f"\n  [DEBUG][assign_position_status][process_group] Обработка группы: CODE={code}, title_mark={title_mark}")
            if debug_title_system:
                _debug_print(f"    debug_title_system={debug_title_system}")
            _debug_print(f"    Количество строк RFP в группе: {len(rfp_group)}")
        
        # Подсчитываем требуемое количество для всех строк RFP в группе
        total_rfp_quantity = 0.0
        for rfp_row in rfp_group:
            rfp_quantity = rfp_row.get_value(VALUES) or 0.0
            try:
                total_rfp_quantity += float(rfp_quantity)
            except (ValueError, TypeError):
                pass
        
        if is_debug_group:
            _debug_print(f"    [DEBUG][assign_position_status][step_1] Подсчет требуемого количества RFP")
            _debug_print(f"      Общее требуемое количество: {total_rfp_quantity}")
            for idx, rfp_row in enumerate(rfp_group, 1):
                rfp_tags = rfp_row.get_tags_list()
                rfp_qty = rfp_row.get_value(VALUES) or 0.0
                _debug_print(f"        Строка [{idx}]: теги={rfp_tags}, VALUES={rfp_qty}")
        
        # Получаем доступные позиции MTO для этого CODE
        title_mto_by_code = mto_by_code.get(title_mark, {})
        code_mto_data = title_mto_by_code.get(code, {"with_tags": [], "without_tags": []})
        
        # Собираем все теги из строк RFP в группе, которые нужно найти в MTO
        required_rfp_tags = set()
        for rfp_row in rfp_group:
            rfp_tags = rfp_row.get_tags_list()
            if rfp_tags:
                required_rfp_tags.update(rfp_tags)
        
        if is_debug_group:
            _debug_print(f"    [DEBUG][assign_position_status][step_2] Получение доступных позиций MTO")
            _debug_print(f"      Найдено позиций с тегами: {len(code_mto_data['with_tags'])}")
            _debug_print(f"      Найдено позиций без тегов: {len(code_mto_data['without_tags'])}")
            _debug_print(f"      Требуемые теги из RFP: {required_rfp_tags}")
        
        # Подсчитываем доступное количество позиций с тегами, которые МОГУТ быть использованы
        # (т.е. позиции с тегами, которые совпадают с требуемыми тегами из RFP)
        available_with_tags = 0.0
        usable_mto_with_tags = []
        for mto_row in code_mto_data["with_tags"]:
            mto_tags = mto_row.get_tags_list()
            # Проверяем, есть ли совпадение тегов
            if mto_tags and any(tag in required_rfp_tags for tag in mto_tags):
                # Эта позиция может быть использована, т.к. тег совпадает
                mto_quantity = mto_row.get_value(VALUES) or 0.0
                try:
                    available_with_tags += float(mto_quantity)
                    usable_mto_with_tags.append(mto_row)
                except (ValueError, TypeError):
                    pass
        
        if is_debug_group:
            _debug_print(
                f"      Позиций с тегами, которые МОГУТ быть использованы (теги совпадают): {len(usable_mto_with_tags)}"
            )
            for idx, mto_row in enumerate(usable_mto_with_tags, 1):
                mto_tags = mto_row.get_tags_list()
                mto_qty = mto_row.get_value(VALUES) or 0
                _debug_print(f"        [{idx}] Теги={mto_tags}, VALUES={mto_qty}")
            unused_with_tags = [mto for mto in code_mto_data["with_tags"] if mto not in usable_mto_with_tags]
            if unused_with_tags:
                _debug_print(
                    f"      Позиций с тегами, которые НЕ МОГУТ быть использованы (теги НЕ совпадают): {len(unused_with_tags)}"
                )
                for idx, mto_row in enumerate(unused_with_tags, 1):
                    mto_tags = mto_row.get_tags_list()
                    mto_qty = mto_row.get_value(VALUES) or 0
                    _debug_print(f"        [{idx}] Теги={mto_tags}, VALUES={mto_qty}")
        
        # Подсчитываем доступное количество позиций без тегов (остаток после step4_2 / VO)
        available_without_tags = 0.0
        for mto_row in code_mto_data["without_tags"]:
            rem = remaining_mto_output_qty(mto_row)
            if rem > 0:
                available_without_tags += rem
        
        total_available = available_with_tags + available_without_tags
        
        if is_debug_group:
            _debug_print(f"    [DEBUG][assign_position_status][step_3] Подсчет доступного количества MTO")
            _debug_print(f"      Доступно с тегами: {available_with_tags}")
            _debug_print(f"      Доступно без тегов: {available_without_tags}")
            _debug_print(f"      Всего доступно: {total_available}")
            _debug_print(f"      Требуется: {total_rfp_quantity}")
            _debug_print(f"      Разница: {total_available - total_rfp_quantity}")
        
        # Определяем статус и распределяем данные
        # 
        # ЛОГИКА ОПРЕДЕЛЕНИЯ СТАТУСА:
        # 
        # 1. Статус "Исключен" устанавливается когда:
        #    total_available < total_rfp_quantity
        #    Это означает, что общее количество позиций в MTO (с тегами + без тегов) 
        #    меньше, чем требуется в RFP. В этом случае невозможно покрыть все требования,
        #    поэтому позиция исключается из обработки.
        #
        # 2. Статус "Не протегирован в МТО" устанавливается когда:
        #    available_with_tags < total_rfp_quantity AND available_without_tags > 0 AND total_available >= total_rfp_quantity
        #    Это означает, что не хватает позиций с тегами, но есть позиции без тегов,
        #    и общее количество хватает. В этом случае используем позиции без тегов.
        #
        # 3. Статус не устанавливается (остается пустым) когда:
        #    available_with_tags >= total_rfp_quantity (достаточно позиций с тегами)
        #    ИЛИ available_without_tags == 0 (нет позиций без тегов для использования)
        #    В этих случаях строки либо уже обработаны на этапе сопоставления,
        #    либо не требуют дополнительной обработки.
        #
        if total_available < total_rfp_quantity:
            # Не хватает позиций в MTO - статус "Исключен"
            # Условие: total_available < total_rfp_quantity
            # Это означает, что сумма всех позиций MTO (с тегами + без тегов) 
            # меньше требуемого количества из RFP
            if is_debug_group:
                _debug_print(f"    [DEBUG][assign_position_status][step_4_not_enough] Определение статуса")
                _debug_print(
                    f"      Результат: НЕ ХВАТАЕТ позиций (доступно {total_available} < требуется {total_rfp_quantity})"
                )
                _debug_print(f"      Статус: 'Исключен'")
                _debug_print(
                    "      Причина: Общее количество позиций в MTO недостаточно для покрытия требований RFP"
                )
            
            for rfp_row in rfp_group:
                rfp_row.el[POSITION_STATUS].value = "Исключен"
                assigned_count += 1
            if is_debug_group:
                _debug_print(
                    f"    Исключено: CODE={code}, title_mark={title_mark}, "
                    f"требуется={total_rfp_quantity}, доступно={total_available}"
                )
        elif available_with_tags < total_rfp_quantity and available_without_tags > 0:
            # Не хватает позиций с тегами, но есть позиции без тегов и общее количество хватает
            # Статус "Не протегирован в МТО" и вставляем данные из позиций без тегов
            
            if is_debug_group:
                _debug_print(f"    [DEBUG][assign_position_status][step_4_use_no_tag] Определение статуса")
                _debug_print(
                    f"      Результат: НЕ ХВАТАЕТ позиций с тегами ({available_with_tags} < {total_rfp_quantity})"
                )
                _debug_print(
                    f"      Но есть позиции без тегов ({available_without_tags}) и общее количество хватает"
                )
                _debug_print(f"      Статус: 'Не протегирован в МТО'")
                _debug_print(f"    [DEBUG][assign_position_status][step_5] Распределение позиций без тегов")
            
            # Распределяем позиции без тегов между строками RFP
            # Создаем список доступных позиций без тегов с их количествами
            available_mto_rows = []
            for mto_row in code_mto_data["without_tags"]:
                rem = remaining_mto_output_qty(mto_row)
                if rem > 0:
                    available_mto_rows.append((mto_row, rem))
            
            if is_debug_group:
                _debug_print(f"      Доступные позиции MTO без тегов: {len(available_mto_rows)}")
                for idx, (mto_row, qty) in enumerate(available_mto_rows, 1):
                    _debug_print(f"        [{idx}] CODE={mto_row.get_value(CODE)}, VALUES={qty}")
            
            # Распределяем количество между строками RFP
            mto_index = 0
            remaining_in_current_mto = available_mto_rows[0][1] if available_mto_rows else 0.0
            
            for rfp_idx, rfp_row in enumerate(rfp_group, 1):
                rfp_tags = rfp_row.get_tags_list()
                rfp_quantity = rfp_row.get_value(VALUES) or 0.0
                try:
                    rfp_quantity_float = float(rfp_quantity)
                except (ValueError, TypeError):
                    rfp_quantity_float = 0.0
                
                if is_debug_group:
                    _debug_print(
                        f"      Обработка строки RFP [{rfp_idx}]: теги={rfp_tags}, VALUES={rfp_quantity_float}"
                    )
                
                if rfp_quantity_float <= 0:
                    if is_debug_group:
                        _debug_print(f"        Пропуск: VALUES <= 0")
                    continue
                
                # Распределяем количество из доступных позиций MTO без тегов
                remaining_needed = rfp_quantity_float
                mto_row_to_use = None
                quantity_to_use = 0.0
                
                if is_debug_group:
                    _debug_print(
                        f"        Начальное состояние: remaining_needed={remaining_needed}, "
                        f"mto_index={mto_index}, remaining_in_current_mto={remaining_in_current_mto}"
                    )
                
                while remaining_needed > 0 and mto_index < len(available_mto_rows):
                    mto_row, mto_available = available_mto_rows[mto_index]
                    
                    # Берем сколько нужно или сколько доступно
                    take_quantity = min(remaining_needed, remaining_in_current_mto)
                    
                    if is_debug_group:
                        _debug_print(
                            f"        Использование позиции MTO [{mto_index + 1}]: "
                            f"CODE={mto_row.get_value(CODE)}, take_quantity={take_quantity}"
                        )
                    
                    if mto_row_to_use is None:
                        mto_row_to_use = mto_row
                        quantity_to_use = 0.0
                    
                    quantity_to_use += take_quantity
                    remaining_needed -= take_quantity
                    remaining_in_current_mto -= take_quantity
                    if take_quantity > 0:
                        mto_row._used_qty_no_tag = (
                            _safe_float(getattr(mto_row, "_used_qty_no_tag", 0.0)) + take_quantity
                        )
                    
                    if is_debug_group:
                        _debug_print(
                            f"          После взятия: quantity_to_use={quantity_to_use}, "
                            f"remaining_needed={remaining_needed}, remaining_in_current_mto={remaining_in_current_mto}"
                        )
                    
                    # Если текущая позиция MTO исчерпана, переходим к следующей
                    if remaining_in_current_mto <= 0:
                        mto_index += 1
                        if mto_index < len(available_mto_rows):
                            remaining_in_current_mto = available_mto_rows[mto_index][1]
                            if is_debug_group:
                                _debug_print(
                                    f"          Переход к следующей позиции MTO [{mto_index + 1}], "
                                    f"remaining_in_current_mto={remaining_in_current_mto}"
                                )
                        else:
                            if is_debug_group:
                                _debug_print(f"          Позиции MTO закончились")
                            break
                
                # Вставляем данные из MTO с поправкой на количество
                if mto_row_to_use and quantity_to_use > 0:
                    if is_debug_group:
                        _debug_print(
                            f"        Вставка данных MTO: CODE={mto_row_to_use.get_value(CODE)}, "
                            f"quantity={quantity_to_use}"
                        )
                    _add_mto_data_with_quantity(rfp_row, mto_row_to_use, quantity_to_use)
                
                rfp_row.el[POSITION_STATUS].value = "Не протегирован в МТО"
                assigned_count += 1
                
                if is_debug_group:
                    _debug_print(
                        f"    Не протегирован в МТО: CODE={code}, title_mark={title_mark}, "
                        f"количество RFP={rfp_quantity_float}, количество MTO={quantity_to_use}"
                    )
        else:
            # Этот блок достигается когда:
            # - total_available >= total_rfp_quantity (общее количество достаточно)
            # - И НЕ (available_with_tags < total_rfp_quantity AND available_without_tags > 0)
            #
            # Это означает один из двух случаев:
            # 1. available_with_tags >= total_rfp_quantity - достаточно позиций с тегами
            #    В этом случае строки уже были сопоставлены на этапе step4_2 (match_rfp_with_mto),
            #    поэтому статус позиции не устанавливается (остается пустым)
            #
            # 2. available_with_tags < total_rfp_quantity AND available_without_tags == 0
            #    НО total_available >= total_rfp_quantity
            #    Это математически невозможно, т.к. total_available = available_with_tags + available_without_tags
            #    Если available_without_tags == 0, то total_available = available_with_tags < total_rfp_quantity
            #    И такой случай попадет в первый блок if (статус "Исключен")
            #
            # ДОПОЛНИТЕЛЬНАЯ ПРОВЕРКА: На всякий случай проверяем, не пропустили ли мы случай,
            # когда нужно установить "Исключен". Если не хватает позиций с тегами и нет позиций без тегов,
            # но по какой-то причине мы попали в этот блок, устанавливаем "Исключен"
            if available_with_tags < total_rfp_quantity and available_without_tags == 0:
                # Это не должно произойти, т.к. должно попасть в первый блок if,
                # но на всякий случай проверяем и устанавливаем статус "Исключен"
                if is_debug_group:
                    _debug_print(f"    [DEBUG][assign_position_status][step_4_excluded] Определение статуса")
                    _debug_print(f"      ВНИМАНИЕ: Попали в блок else, но не хватает позиций!")
                    _debug_print(
                        f"      available_with_tags={available_with_tags} < total_rfp_quantity={total_rfp_quantity}"
                    )
                    _debug_print(
                        f"      available_without_tags={available_without_tags} (нет позиций без тегов)"
                    )
                    _debug_print(f"      total_available={total_available}")
                    _debug_print(f"      Статус: 'Исключен' (дополнительная проверка)")
                
                for rfp_row in rfp_group:
                    rfp_row.el[POSITION_STATUS].value = "Исключен"
                    assigned_count += 1
                if is_debug_group:
                    _debug_print(
                        f"    Исключено (доп. проверка): CODE={code}, title_mark={title_mark}, "
                        f"требуется={total_rfp_quantity}, доступно={total_available}"
                    )
            else:
                # Статус остается пустым, т.к. эти строки либо уже обработаны на этапе сопоставления,
                # либо не требуют дополнительной обработки
                if is_debug_group:
                    _debug_print(f"    [DEBUG][assign_position_status][step_4_ok] Определение статуса")
                    _debug_print(
                        f"      Результат: Достаточно позиций с тегами ({available_with_tags} >= {total_rfp_quantity})"
                    )
                    _debug_print(f"      ИЛИ нет позиций без тегов ({available_without_tags} = 0)")
                    _debug_print(f"      Статус: не устанавливается (остается пустым)")
    
    if debug:
        # Общая статистика - логируем в debug_log только если нет специфичных фильтров
        should_log_summary = not has_debug_filters
        _debug_print(f"    Назначено статусов позиции: {assigned_count}", should_log=should_log_summary)

    # Отдельный файл не создаем — используем единый debug-лог на уровне step4


def _add_mto_data_with_quantity(rfp_row: RowStd, mto_row: RowStd, quantity: float):
    """
    Добавляет данные из MTO в строку RFP с указанным количеством
    
    Args:
        rfp_row: Строка RFP для модификации
        mto_row: Строка MTO с данными
        quantity: Количество для установки в VALUES_MTO
    """
    # Добавляем данные из MTO (тег будет пустым, т.к. позиция без тегов)
    rfp_row.el[TAG_MTO].value = ""
    rfp_row.el[CODE_MTO].value = mto_row.get_value(CODE) or ""
    rfp_row.el[NAME_MTO].value = mto_row.get_value(NAME) or ""
    rfp_row.el[TYPE_MARK_MTO].value = mto_row.get_value(TYPE_MARK) or ""
    rfp_row.el[VALUES_MTO].value = quantity
    rfp_row.el[NUMBERS_MTO].value = mto_row.get_value(NUMBERS) or ""
    copy_mto_units_fields(rfp_row, mto_row)
