"""
Этап 4_2: Сопоставление строк RFP с MTO по тегам
"""

import time
from typing import List, Dict, Set, Tuple, Optional
from collections import defaultdict

from RFQ.tags_rfp_compare.rfp_tags_utils import append_timing_log

from base.base_classes import RowStd, RowType
from base.tables_columns import (
    CODE, DS_TITLE, MATCH_STATUS, HAS_TAGS_RFP, NAME, TYPE_MARK, NUMBERS, VALUES,
    TAG_MTO, CODE_MTO, MTO_CODE_STRUCK, NAME_MTO, TYPE_MARK_MTO, NUMBERS_MTO, VALUES_MTO,
    TAG_VO, CODE_VO, NAME_VO, VALUES_VO, MATCH_STATUS_VO, MATCH_STATUS_CODE_REPLACEMENT,
    POSITION_STATUS, IN_CABINET
)
from RFQ.ds_compare.ds_units_normalize import copy_mto_units_fields
from RFQ.tags_rfp_compare.rfp_supply_status import is_excluded_from_supply


def _should_debug(
    debug: bool,
    title_system: str,
    code: str,
    tags: List[str],
    debug_tag: List[str],
    debug_code: List[str],
    debug_title_system: List[str]) -> bool:
    """
    Проверяет, нужно ли выводить отладочную информацию.
    Все указанные (не None/пустые) фильтры должны совпадать (AND между фильтрами).
    Внутри каждого фильтра-списка работает логика OR (совпадение с любым элементом).
    """
    if not debug:
        return False
    if not (debug_tag or debug_code or debug_title_system):
        return True
    # debug_title_system: проверяем, что title_system входит в список
    if debug_title_system and title_system not in debug_title_system:
        return False
    # debug_code: проверяем, что code входит в список
    if debug_code and code not in debug_code:
        return False
    # debug_tag: проверяем, что хотя бы один тег из tags входит в debug_tag
    if debug_tag and (not tags or not any(t in debug_tag for t in tags)):
        return False
    return True


def _should_debug_cross_title(
    debug: bool,
    title_system: str,
    code: str,
    tags: List[str],
    debug_tag: List[str],
    debug_code: List[str],
    debug_title_system: List[str]) -> bool:
    """
    Проверяет, нужно ли выводить отладочную информацию.
    Все указанные (не None/пустые) фильтры должны совпадать (AND между фильтрами).
    Внутри каждого фильтра-списка работает логика OR (совпадение с любым элементом).
    """
    if not debug:
        return False
    if not (debug_tag or debug_code or debug_title_system):
        return True
    # debug_title_system: проверяем, что title_system входит в список
    if debug_title_system and title_system not in debug_title_system:
        return False
    # debug_code: проверяем, что code входит в список
    if debug_code and code not in debug_code:
        return False
    # debug_tag: проверяем, что хотя бы один тег из tags входит в debug_tag
    if debug_tag and (not tags or not any(t in debug_tag for t in tags)):
        return False
    return True


def _debug_print(debug_log: List[str], *args, should_log: bool = True, **kwargs) -> None:
    """
    Выводит отладочную информацию в консоль и опционально в debug_log.
    
    Args:
        debug_log: Список строк для накопления debug-лога
        *args: Аргументы для print
        should_log: Если True, записывает в debug_log. Если False, только в консоль.
        **kwargs: Kwargs для print (sep, end)
    """
    print(*args, **kwargs)
    if debug_log is None or not should_log:
        return
    sep = kwargs.get("sep", " ")
    end = kwargs.get("end", "\n")
    debug_log.append(sep.join(str(arg) for arg in args) + end)


def match_rfp_with_mto(
    rfp_rows: List[RowStd],
    mto_data: Dict[str, List[RowStd]],
    replacement_table: Optional[Dict[str, List[Tuple[str, str]]]] = None,
    result_dir: Optional[str] = None,
    debug: bool = False,
    debug_tag: List[str] = None,
    debug_code: List[str] = None,
    debug_title_system: List[str] = None,
    debug_log: List[str] = None) -> tuple[Set[str], List[RowStd]]:
    """
    Сопоставляет строки RFP с MTO по тегам в рамках одного title_mark
    
    Args:
        rfp_rows: Список строк RFP для сопоставления
        mto_data: Словарь title_mark -> список строк MTO
        replacement_table: Таблица замен кодов (старый -> список новых кодов)
        debug: Флаг отладки
        debug_tag: Тег для точечной отладки (AND с другими фильтрами)
        debug_code: Код для точечной отладки (AND с другими фильтрами)
        debug_title_system: title_system для точечной отладки (AND с другими фильтрами)
        
    Returns:
        Кортеж: (множество сопоставленных тегов, список несопоставленных строк MTO)
    """
    matched_tags: Set[str] = set()
    matched_tag_sources: Dict[str, Set[str]] = defaultdict(set)
    unmatched_mto_rows: List[RowStd] = []
    used_mto_rows: Set[RowStd] = set()
    replacement_table = replacement_table or {}
    mto_uin_counter = 0

    def _ensure_mto_uin(mto_row: RowStd) -> int:
        nonlocal mto_uin_counter
        current = getattr(mto_row, "_uin_mto", None)
        if current is None:
            mto_uin_counter += 1
            mto_row._uin_mto = mto_uin_counter
            return mto_uin_counter
        return current
    
    # Создаем индекс MTO строк по title_mark и тегам для быстрого поиска
    step_start = time.perf_counter()
    mto_index: Dict[str, Dict[str, List[RowStd]]] = defaultdict(lambda: defaultdict(list))
    # Индекс MTO строк по title_mark и коду
    mto_code_index: Dict[str, Dict[str, List[RowStd]]] = defaultdict(lambda: defaultdict(list))
    for title_mark, mto_rows in mto_data.items():
        for mto_row in mto_rows:
            if mto_row.row_type != RowType.position_row:
                continue
            mto_code = mto_row.get_value(CODE)
            if mto_code:
                mto_code_index[title_mark][str(mto_code)].append(mto_row)
            mto_tags = mto_row.get_tags_list()
            mto_uin = _ensure_mto_uin(mto_row)
            if _should_debug(
                debug,
                title_mark,
                mto_code or "",
                mto_tags,
                debug_tag,
                debug_code,
                debug_title_system
            ):
                _debug_print(
                    debug_log,
                    f"  [DEBUG][match_rfp_with_mto][index_build] MTO строка в индексе: title_mark={title_mark}, "
                    f"CODE={mto_code}, TAGS={mto_tags}, VALUES={mto_row.get_value(VALUES)}, "
                    f"IN_CABINET='{mto_row.get_value(IN_CABINET) or ''}', UIN_MTO={mto_uin}"
                )
            if mto_tags:
                # В MTO каждая строка содержит не более одного тега
                tag = mto_tags[0]
                mto_index[title_mark][tag].append(mto_row)
    if result_dir:
        append_timing_log(result_dir, f"match_rfp_with_mto::mto_index_build: {time.perf_counter() - step_start:.3f}s")
    
    # Предварительный список position_row и кэш полей для ускорения проходов
    step_start = time.perf_counter()
    position_row_list = [r for r in rfp_rows if r.row_type == RowType.position_row]
    rfp_precomputed = [
        (r, r.get_value(DS_TITLE) or "", r.get_tags_list(), r.get_value(CODE) or "", _normalize_code(r.get_value(CODE) or ""))
        for r in position_row_list
        if not is_excluded_from_supply(r)
    ]
    if result_dir:
        append_timing_log(result_dir, f"match_rfp_with_mto::rfp_precomputed_build: {time.perf_counter() - step_start:.3f}s")
    
    # Сопоставляем строки RFP с MTO
    # Сначала устанавливаем статус наличия тегов для всех строк RFP
    step_start = time.perf_counter()
    for rfp_row in position_row_list:
        if rfp_row.row_type != RowType.position_row:
            continue
        
        rfp_tags = rfp_row.get_tags_list()
        # Устанавливаем статус наличия тегов в RFP
        rfp_row.el[HAS_TAGS_RFP].value = "Да" if rfp_tags else "Нет"
        if not rfp_row.el[MATCH_STATUS_CODE_REPLACEMENT].value:
            rfp_row.el[MATCH_STATUS_CODE_REPLACEMENT].value = ""
        if is_excluded_from_supply(rfp_row):
            continue

        title_mark = rfp_row.get_value(DS_TITLE)
        if not title_mark:
            # Строки без title_mark не участвуют в сопоставлении
            continue

        # Инициализируем статус как "Не найден в МТО" для ВСЕХ строк с title_mark
        # (включая строки без тегов - они могут быть сопоставлены по коду)
        rfp_row.el[MATCH_STATUS].value = "Не найден в МТО"
    if result_dir:
        append_timing_log(result_dir, f"match_rfp_with_mto::pass_init_status: {time.perf_counter() - step_start:.3f}s")

    # PASS 1: приоритетное сопоставление по тегу + коду
    step_start = time.perf_counter()
    for rfp_row, title_mark, rfp_tags, rfp_code_raw, rfp_code in rfp_precomputed:
        if rfp_row.el[MATCH_STATUS].value != "Не найден в МТО":
            continue
        if not rfp_tags:
            continue
        if not title_mark:
            continue
        if not rfp_code:
            continue

        title_mto_index = mto_index.get(title_mark, {})
        for rfp_tag in rfp_tags:
            candidates = title_mto_index.get(rfp_tag, [])
            if not candidates:
                continue
            mto_row = None
            for candidate in candidates:
                if candidate in used_mto_rows:
                    continue
                mto_code = _normalize_code(candidate.get_value(CODE))
                if not mto_code or mto_code != rfp_code:
                    continue
                mto_row = candidate
                break
            if mto_row is None:
                continue

            matched_tags.add(rfp_tag)
            matched_tag_sources[rfp_tag].add(title_mark)
            used_mto_rows.add(mto_row)
            if _should_debug_cross_title(
                debug,
                title_mark,
                mto_row.get_value(CODE) or "",
                mto_row.get_tags_list(),
                debug_tag,
                debug_code,
                debug_title_system
            ):
                _debug_print(
                    debug_log,
                    f"  [DEBUG][match_rfp_with_mto][match_tag_code] MTO использована (тег+код): title_mark={title_mark}, "
                    f"CODE={mto_row.get_value(CODE)}, TAG={mto_row.get_tags_list()}, "
                    f"RFP_TAG={rfp_tag}, RFP_CODE={rfp_code_raw}"
                )
            add_mto_data_to_rfp_row(rfp_row, mto_row)
            rfp_row.el[MATCH_STATUS].value = "Тег сопоставлен"
            if _should_debug(
                debug,
                title_mark,
                rfp_code_raw,
                rfp_tags,
                debug_tag,
                debug_code,
                debug_title_system
            ):
                _debug_print(
                    debug_log,
                    f"  Сопоставлено по тегу+коду: RFP тег '{rfp_tag}' (title_mark: {title_mark}) "
                    f"с MTO код '{mto_row.get_value(CODE)}'"
                )
            break
    if result_dir:
        append_timing_log(result_dir, f"match_rfp_with_mto::pass_1_tag_code: {time.perf_counter() - step_start:.3f}s")

    # PASS 1.5: сопоставление по тегу + код через таблицу замен (только для RFP с тегами)
    step_start = time.perf_counter()
    for rfp_row, title_mark, rfp_tags, rfp_code_raw, rfp_code in rfp_precomputed:
        if rfp_row.el[MATCH_STATUS].value != "Не найден в МТО":
            continue
        if not rfp_tags:
            continue
        if not title_mark:
            continue
        if not rfp_code:
            continue

        # Получаем коды замены для RFP кода
        replacement_codes = _get_replacement_codes(rfp_code, replacement_table)
        if not replacement_codes:
            continue

        title_mto_index = mto_index.get(title_mark, {})
        matched = False
        
        for rfp_tag in rfp_tags:
            candidates = title_mto_index.get(rfp_tag, [])
            if not candidates:
                continue
            
            for replacement_code in replacement_codes:
                normalized_replacement = _normalize_code(replacement_code)
                if not normalized_replacement:
                    continue
                
                mto_row = None
                for candidate in candidates:
                    if candidate in used_mto_rows:
                        continue
                    mto_code = _normalize_code(candidate.get_value(CODE))
                    if not mto_code or mto_code != normalized_replacement:
                        continue
                    mto_row = candidate
                    break
                
                if mto_row is None:
                    continue

                matched_tags.add(rfp_tag)
                matched_tag_sources[rfp_tag].add(title_mark)
                used_mto_rows.add(mto_row)
                
                add_mto_data_to_rfp_row(rfp_row, mto_row)
                rfp_row.el[MATCH_STATUS].value = "Тег сопоставлен"
                
                _append_replacement_status(
                    rfp_row,
                    f"MTO(тег+код): {rfp_code} -> {replacement_code}"
                )
                
                if _should_debug_cross_title(
                    debug,
                    title_mark,
                    mto_row.get_value(CODE) or "",
                    mto_row.get_tags_list(),
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rfp_with_mto][match_tag_code_replacement] MTO использована (тег+код замена): "
                        f"title_mark={title_mark}, CODE={mto_row.get_value(CODE)}, TAG={mto_row.get_tags_list()}, "
                        f"RFP_TAG={rfp_tag}, RFP_CODE={rfp_code_raw}, REPL={replacement_code}"
                    )
                
                if _should_debug(
                    debug,
                    title_mark,
                    rfp_code_raw,
                    rfp_tags,
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  Сопоставлено по тегу+коду (замена): RFP тег '{rfp_tag}' (title_mark: {title_mark}) "
                        f"с MTO код '{mto_row.get_value(CODE)}' (RFP код '{rfp_code_raw}' -> '{replacement_code}')"
                    )
                matched = True
                break
            
            if matched:
                break
    if result_dir:
        append_timing_log(result_dir, f"match_rfp_with_mto::pass_1_5_replacement: {time.perf_counter() - step_start:.3f}s")

    # PASS 1.6: сопоставление по тегу + код через ОБРАТНУЮ таблицу замен (только для RFP с тегами)
    step_start = time.perf_counter()
    # RFP код A - ищем в таблице где A является заменой (X → A) - ищем MTO с кодом X
    for rfp_row, title_mark, rfp_tags, rfp_code_raw, rfp_code in rfp_precomputed:
        if rfp_row.el[MATCH_STATUS].value != "Не найден в МТО":
            continue
        if not rfp_tags:
            continue
        if not title_mark:
            continue
        if not rfp_code:
            continue

        # Получаем оригинальные коды (обратный поиск)
        original_codes = _get_reverse_replacement_codes(rfp_code, replacement_table)
        if not original_codes:
            continue

        title_mto_index = mto_index.get(title_mark, {})
        matched = False

        for rfp_tag in rfp_tags:
            candidates = title_mto_index.get(rfp_tag, [])
            if not candidates:
                continue

            for original_code in original_codes:
                normalized_original = _normalize_code(original_code)
                if not normalized_original:
                    continue

                mto_row = None
                for candidate in candidates:
                    if candidate in used_mto_rows:
                        continue
                    mto_code = _normalize_code(candidate.get_value(CODE))
                    if not mto_code or mto_code != normalized_original:
                        continue
                    mto_row = candidate
                    break

                if mto_row is None:
                    continue

                matched_tags.add(rfp_tag)
                matched_tag_sources[rfp_tag].add(title_mark)
                used_mto_rows.add(mto_row)

                add_mto_data_to_rfp_row(rfp_row, mto_row)
                rfp_row.el[MATCH_STATUS].value = "Тег сопоставлен"

                _append_replacement_status(
                    rfp_row,
                    f"MTO(тег+код обр.): {original_code} <- {rfp_code}"
                )

                if _should_debug_cross_title(
                    debug,
                    title_mark,
                    mto_row.get_value(CODE) or "",
                    mto_row.get_tags_list(),
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rfp_with_mto][match_tag_code_reverse] MTO использована (тег+код обратная замена): "
                        f"title_mark={title_mark}, CODE={mto_row.get_value(CODE)}, TAG={mto_row.get_tags_list()}, "
                        f"RFP_TAG={rfp_tag}, RFP_CODE={rfp_code_raw}, ORIG={original_code}"
                    )

                if _should_debug(
                    debug,
                    title_mark,
                    rfp_code_raw,
                    rfp_tags,
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  Сопоставлено по тегу+коду (обратная замена): RFP тег '{rfp_tag}' (title_mark: {title_mark}) "
                        f"с MTO код '{mto_row.get_value(CODE)}' (MTO код '{original_code}' <- RFP код '{rfp_code_raw}')"
                    )
                matched = True
                break

            if matched:
                break
    if result_dir:
        append_timing_log(result_dir, f"match_rfp_with_mto::pass_1_6_reverse: {time.perf_counter() - step_start:.3f}s")
    
    # Дополнительный проход: сопоставление по коду для оставшихся строк RFP
    # Pre-filter + pointer: O(N) суммарно вместо O(N^2)
    step_start = time.perf_counter()
    mto_code_avail: Dict[tuple, list] = {}
    for tm, code_dict in mto_code_index.items():
        for code, rows in code_dict.items():
            avail = [r for r in rows if r not in used_mto_rows]
            if avail:
                mto_code_avail[(tm, code)] = avail

    code_ptr: Dict[tuple, int] = {}
    for rfp_row, title_mark, rfp_tags, rfp_code_raw, rfp_code in rfp_precomputed:
        if rfp_row.el[MATCH_STATUS].value != "Не найден в МТО":
            continue
        if not rfp_code_raw:
            continue
        if not title_mark:
            continue
        
        key = (title_mark, str(rfp_code_raw))
        avail = mto_code_avail.get(key)
        if not avail:
            continue
        start = code_ptr.get(key, 0)
        if start >= len(avail):
            continue
        
        rfp_value = rfp_row.get_value(VALUES)
        for i in range(start, len(avail)):
            mto_row = avail[i]
            mto_value = mto_row.get_value(VALUES)
            if _is_mto_exceeds_rfp(rfp_value, mto_value):
                continue
            
            add_mto_data_to_rfp_row(rfp_row, mto_row)
            rfp_row.el[MATCH_STATUS].value = "Тег в МТО заменен"
            used_mto_rows.add(mto_row)
            code_ptr[key] = i + 1
            if _should_debug_cross_title(
                debug,
                title_mark,
                mto_row.get_value(CODE) or "",
                mto_row.get_tags_list(),
                debug_tag,
                debug_code,
                debug_title_system
            ):
                _debug_print(
                    debug_log,
                    f"  [DEBUG][match_rfp_with_mto][match_code] MTO использована (код): title_mark={title_mark}, "
                    f"CODE={mto_row.get_value(CODE)}, TAG={mto_row.get_tags_list()}, RFP_CODE={rfp_code_raw}"
                )
            mto_tags = mto_row.get_tags_list()
            if mto_tags:
                matched_tags.add(mto_tags[0])
                matched_tag_sources[mto_tags[0]].add(title_mark)
            else:
                if rfp_row.el[POSITION_STATUS].value in (None, ""):
                    rfp_row.el[POSITION_STATUS].value = "Не протегирован в МТО"
            
            if _should_debug(
                debug,
                title_mark,
                rfp_code_raw or "",
                rfp_tags,
                debug_tag,
                debug_code,
                debug_title_system
            ):
                _debug_print(
                    debug_log,
                    f"  Сопоставлено по коду: RFP CODE '{rfp_code_raw}' (title_mark: {title_mark}) "
                    f"с MTO тег '{mto_tags[0] if mto_tags else ''}'"
                )
            break
    if result_dir:
        append_timing_log(result_dir, f"match_rfp_with_mto::pass_match_code: {time.perf_counter() - step_start:.3f}s")

    # Дополнительный проход: сопоставление по коду через таблицу замен
    step_start = time.perf_counter()
    for rfp_row, title_mark, rfp_tags, rfp_code_raw, rfp_code in rfp_precomputed:
        if rfp_row.el[MATCH_STATUS].value != "Не найден в МТО":
            continue
        if not rfp_code_raw:
            continue
        if not title_mark:
            continue

        replacement_codes = _get_replacement_codes(str(rfp_code_raw), replacement_table)
        if not replacement_codes:
            continue

        rfp_value = rfp_row.get_value(VALUES)
        for replacement_code in replacement_codes:
            candidates = mto_code_index.get(title_mark, {}).get(str(replacement_code), [])
            if not candidates:
                continue

            for mto_row in candidates:
                if mto_row in used_mto_rows:
                    continue
                mto_value = mto_row.get_value(VALUES)
                if _is_mto_exceeds_rfp(rfp_value, mto_value):
                    continue

                add_mto_data_to_rfp_row(rfp_row, mto_row)
                rfp_row.el[MATCH_STATUS].value = "Тег в МТО заменен"
                used_mto_rows.add(mto_row)
                if _should_debug_cross_title(
                    debug,
                    title_mark,
                    mto_row.get_value(CODE) or "",
                    mto_row.get_tags_list(),
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rfp_with_mto][match_replacement] MTO использована (замена): title_mark={title_mark}, "
                        f"CODE={mto_row.get_value(CODE)}, TAG={mto_row.get_tags_list()}, "
                        f"RFP_CODE={rfp_code_raw}, REPL={replacement_code}"
                    )
                _append_replacement_status(
                    rfp_row,
                    f"MTO: {rfp_code_raw} -> {replacement_code}"
                )
                mto_tags = mto_row.get_tags_list()
                if mto_tags:
                    matched_tags.add(mto_tags[0])
                    matched_tag_sources[mto_tags[0]].add(title_mark)
                else:
                    if rfp_row.el[POSITION_STATUS].value in (None, ""):
                        rfp_row.el[POSITION_STATUS].value = "Не протегирован в МТО"

                if _should_debug(
                    debug,
                    title_mark,
                    rfp_code_raw or "",
                    rfp_tags,
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  Сопоставлено по таблице замен: RFP CODE '{rfp_code_raw}' -> '{replacement_code}' "
                        f"(title_mark: {title_mark}) с MTO тег '{mto_tags[0] if mto_tags else ''}'"
                    )
                break

            if rfp_row.el[MATCH_STATUS].value != "Не найден в МТО":
                break
    if result_dir:
        append_timing_log(result_dir, f"match_rfp_with_mto::pass_match_code_replacement: {time.perf_counter() - step_start:.3f}s")
    
    # Собираем несопоставленные строки MTO (с тегами и без тегов — ни одна position_row не теряется)
    step_start = time.perf_counter()
    for title_mark, mto_rows in mto_data.items():
        for mto_row in mto_rows:
            if mto_row.row_type != RowType.position_row:
                continue
            mto_tags = mto_row.get_tags_list()
            code_val = mto_row.get_value(CODE) or ""
            is_debug_row = _should_debug(
                debug,
                title_mark,
                code_val,
                mto_tags,
                debug_tag,
                debug_code,
                debug_title_system,
            )
            if mto_tags:
                tag = mto_tags[0]
                if mto_row not in used_mto_rows:
                    unmatched_mto_rows.append(mto_row)
                    if is_debug_row:
                        note = ""
                        if tag in matched_tags:
                            sources = sorted(matched_tag_sources.get(tag, set()))
                            note = f", NOTE=tag_matched:{','.join(sources) if sources else 'unknown'}"
                        _debug_print(
                            debug_log,
                            f"  [DEBUG][match_rfp_with_mto][unmatched] MTO несопоставлена: title_mark={title_mark}, "
                            f"CODE={code_val}, TAG={mto_tags}, "
                            f"VALUES={mto_row.get_value(VALUES)}{note}"
                        )
                elif is_debug_row:
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rfp_with_mto][unmatched_skip] MTO НЕ добавлена в несопоставленные: title_mark={title_mark}, "
                        f"CODE={code_val}, TAG={mto_tags}, REASON=row_used"
                    )
                continue

            # Без тегов: остаток после match по коду / частичного _used_qty_no_tag
            if mto_row in used_mto_rows:
                if is_debug_row:
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rfp_with_mto][unmatched_skip] MTO НЕ добавлена в несопоставленные: title_mark={title_mark}, "
                        f"CODE={code_val}, TAG=[], REASON=row_used"
                    )
                continue
            if getattr(mto_row, "_consumed_in_output", False):
                if is_debug_row:
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rfp_with_mto][unmatched_skip] MTO НЕ добавлена в несопоставленные: title_mark={title_mark}, "
                        f"CODE={code_val}, TAG=[], REASON=consumed_in_output"
                    )
                continue
            remaining_qty = max(
                0.0,
                _safe_float(mto_row.get_value(VALUES))
                - _safe_float(getattr(mto_row, "_used_qty_no_tag", 0.0)),
            )
            if remaining_qty <= 0:
                if is_debug_row:
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rfp_with_mto][unmatched_skip] MTO НЕ добавлена в несопоставленные: title_mark={title_mark}, "
                        f"CODE={code_val}, TAG=[], REASON=no_remaining_qty"
                    )
                continue
            unmatched_mto_rows.append(mto_row)
            if is_debug_row:
                _debug_print(
                    debug_log,
                    f"  [DEBUG][match_rfp_with_mto][unmatched_no_tag] MTO несопоставлена (без тегов): title_mark={title_mark}, "
                    f"CODE={code_val}, VALUES={mto_row.get_value(VALUES)}, remaining={remaining_qty}"
                )
    if result_dir:
        append_timing_log(result_dir, f"match_rfp_with_mto::collect_unmatched: {time.perf_counter() - step_start:.3f}s")
    
    return matched_tags, unmatched_mto_rows


def _get_replacement_codes(code: str, replacement_table: Dict[str, List[Tuple[str, str]]]) -> List[str]:
    if not code:
        return []
    replacements = replacement_table.get(code, [])
    return [new_code for new_code, _ in replacements if new_code]


def _get_reverse_replacement_codes(
    code: str,
    replacement_table: Dict[str, List[Tuple[str, str]]]) -> List[str]:
    """
    Обратный поиск в таблице замен: по коду замены находит оригинальные коды.
    
    Args:
        code: Код (возможно, замененный код)
        replacement_table: Таблица замен (оригинальный код -> список (новый код, описание))
        
    Returns:
        Список оригинальных кодов, которые заменяются на code
    """
    if not code or not replacement_table:
        return []
    
    code_normalized = str(code).strip().upper()
    original_codes = []
    
    for original_code, replacements in replacement_table.items():
        for new_code, _ in replacements:
            if new_code and new_code.strip().upper() == code_normalized:
                original_codes.append(original_code)
                break
    
    return original_codes


def _append_replacement_status(rfp_row: RowStd, message: str) -> None:
    if not message:
        return
    current = rfp_row.el[MATCH_STATUS_CODE_REPLACEMENT].value or ""
    if current:
        rfp_row.el[MATCH_STATUS_CODE_REPLACEMENT].value = f"{current}; {message}"
    else:
        rfp_row.el[MATCH_STATUS_CODE_REPLACEMENT].value = message


def _normalize_code(code_value) -> str:
    if code_value is None:
        return ""
    return str(code_value).strip().upper()


def _normalize_title_system(title_system) -> str:
    if title_system is None:
        return ""
    return str(title_system).strip()


def _safe_float(value) -> float:
    try:
        return float(value) if value not in (None, "") else 0.0
    except (ValueError, TypeError):
        return 0.0


def remaining_mto_output_qty(mto_row: RowStd) -> float:
    """Остаток MTO для вывода в result после частичных привязок (_used_qty_no_tag)."""
    if getattr(mto_row, "_consumed_in_output", False):
        return 0.0
    return max(
        0.0,
        _safe_float(mto_row.get_value(VALUES))
        - _safe_float(getattr(mto_row, "_used_qty_no_tag", 0.0)),
    )


def consume_mto_for_export(mto_row: RowStd, qty: float) -> None:
    """Списать количество, уже выведенное в result_rows (отдельная строка или VALUES_MTO)."""
    if qty <= 0:
        return
    mto_row._used_qty_no_tag = _safe_float(getattr(mto_row, "_used_qty_no_tag", 0.0)) + qty
    if remaining_mto_output_qty(mto_row) <= 0:
        mto_row._consumed_in_output = True


def _ref_cabinet_forbids_empty_mto_cabinet(ref_in_cabinet, mto_in_cabinet) -> bool:
    """Если у строки VO/RFP задан шкаф, MTO с пустым IN_CABINET не подходит."""
    ref = str(ref_in_cabinet or "").strip()
    mto = str(mto_in_cabinet or "").strip()
    return bool(ref) and not bool(mto)


def _expand_no_tag_candidate_codes(
    seed: str,
    replacement_table: Optional[Dict[str, List[Tuple[str, str]]]],
) -> List[str]:
    """Прямой код + коды из таблицы замен (прямой и обратный поиск), без дубликатов."""
    seed_n = _normalize_code(seed)
    if not seed_n:
        return []
    out: List[str] = [seed_n]
    rt = replacement_table or {}
    if not rt:
        return out
    for rc in _get_replacement_codes(seed_n, rt):
        rc_n = _normalize_code(rc)
        if rc_n and rc_n not in out:
            out.append(rc_n)
    for orig in _get_reverse_replacement_codes(seed_n, rt):
        orig_n = _normalize_code(orig)
        if orig_n and orig_n not in out:
            out.append(orig_n)
    return out


def _collect_mto_no_tag_rows_for_codes(
    title_system: str,
    candidate_codes: List[str],
    mto_no_tag_by_title: Dict[str, Dict[str, List[RowStd]]],
) -> List[RowStd]:
    by_title = mto_no_tag_by_title.get(title_system)
    if not by_title:
        return []
    seen: Set[int] = set()
    merged: List[RowStd] = []
    for code in candidate_codes:
        for row in by_title.get(code, []):
            rid = id(row)
            if rid in seen:
                continue
            seen.add(rid)
            merged.append(row)
    return merged


def _try_match_mto_no_tag_for_vo(
    rfp_row: RowStd,
    title_system: str,
    mto_no_tag_by_title: Dict[str, Dict[str, List[RowStd]]],
    mto_remaining_qty: Dict[RowStd, float],
    replacement_table: Optional[Dict[str, List[Tuple[str, str]]]] = None,
    debug: bool = False,
    debug_tag: List[str] = None,
    debug_code: List[str] = None,
    debug_title_system: List[str] = None,
    debug_log: List[str] = None):
    if not mto_no_tag_by_title:
        return
    if rfp_row.el.get(CODE_MTO) and rfp_row.el[CODE_MTO].value:
        return

    seeds: List[str] = []
    rfp_code = rfp_row.get_value(CODE)
    if rfp_code:
        n = _normalize_code(rfp_code)
        if n:
            seeds.append(n)
    vo_el = rfp_row.el.get(CODE_VO)
    vo_code_val = vo_el.value if vo_el else ""
    if vo_code_val:
        n = _normalize_code(vo_code_val)
        if n and n not in seeds:
            seeds.append(n)

    if not seeds:
        return

    candidate_codes: List[str] = []
    for s in seeds:
        for c in _expand_no_tag_candidate_codes(s, replacement_table):
            if c not in candidate_codes:
                candidate_codes.append(c)

    candidates = _collect_mto_no_tag_rows_for_codes(
        title_system, candidate_codes, mto_no_tag_by_title)
    if not candidates:
        return

    ref_ic = str(rfp_row.get_value(IN_CABINET) or "").strip()
    candidates = [
        r for r in candidates
        if not _ref_cabinet_forbids_empty_mto_cabinet(ref_ic, r.get_value(IN_CABINET))
    ]
    if not candidates:
        return

    if ref_ic and len(candidates) > 1:
        candidates = sorted(
            candidates,
            key=lambda r: (
                0 if _in_cabinet_matches(ref_ic, str(r.get_value(IN_CABINET) or "").strip())
                else (1 if not str(r.get_value(IN_CABINET) or "").strip() else 2)
            ),
        )
    else:
        candidates = _sort_candidates_by_in_cabinet(candidates, rfp_row)

    for mto_row in candidates:
        remaining = mto_remaining_qty.get(mto_row, 0.0)
        if remaining < 1.0:
            continue
        mto_ic = mto_row.get_value(IN_CABINET) or ""
        if _in_cabinet_conflict(ref_ic, mto_ic):
            continue

        mto_remaining_qty[mto_row] = remaining - 1.0

        add_mto_data_to_rfp_row(rfp_row, mto_row, qty_override=1.0)
        if rfp_row.el[MATCH_STATUS].value in (None, ""):
            rfp_row.el[MATCH_STATUS].value = "Тег в МТО заменен"
        if rfp_row.el[POSITION_STATUS].value in (None, ""):
            rfp_row.el[POSITION_STATUS].value = "Не протегирован в МТО"

        mto_code_dbg = _normalize_code(mto_row.get_value(CODE))
        if _should_debug(
            debug,
            rfp_row.get_value(DS_TITLE),
            rfp_row.get_value(CODE) or "",
            rfp_row.get_tags_list(),
            debug_tag,
            debug_code,
            debug_title_system
        ):
            rfp_ic = rfp_row.get_value(IN_CABINET) or ""
            _debug_print(
                debug_log,
                f"  [DEBUG][_try_match_mto_no_tag_for_vo][match_no_tag] VO->MTO (без тегов) по коду: "
                f"title_system={title_system}, MTO_CODE={mto_code_dbg}, "
                f"RFP_IN_CABINET='{rfp_ic}', MTO_IN_CABINET='{mto_ic}', остаток={mto_remaining_qty[mto_row]}"
            )
        return


def _merge_in_cabinet(current_value: str, new_value: str) -> str:
    """
    Объединяет значения IN_CABINET: если совпадают - одно значение, если различаются - через /
    """
    new_value = str(new_value).strip() if new_value else ""
    if not new_value:
        return current_value or ""
    current_value = str(current_value).strip() if current_value else ""
    if not current_value:
        return new_value
    if current_value == new_value:
        return current_value
    parts = [p.strip() for p in current_value.split("/") if p.strip()]
    if new_value in parts:
        return current_value
    return f"{current_value}/{new_value}"


def _in_cabinet_matches(rfp_in_cabinet: str, vo_in_cabinet: str) -> bool:
    """Проверяет совпадение IN_CABINET: VO входит в список шкафов RFP."""
    rfp_val = str(rfp_in_cabinet or "").strip()
    vo_val = str(vo_in_cabinet or "").strip()
    if not rfp_val or not vo_val:
        return False
    rfp_parts = [p.strip() for p in rfp_val.split("/") if p.strip()]
    return vo_val in rfp_parts or vo_val == rfp_val


def _in_cabinet_conflict(ref_in_cabinet: str, vo_in_cabinet: str) -> bool:
    """
    Возвращает True если VO кандидат должен быть пропущен: оба IN_CABINET не пусты и отличаются.
    Цель: строки остаются в своём шкафу, без перечислений в IN_CABINET.
    """
    ref_val = str(ref_in_cabinet or "").strip()
    vo_val = str(vo_in_cabinet or "").strip()
    if not ref_val or not vo_val:
        return False
    return not _in_cabinet_matches(ref_val, vo_val)


def _sort_candidates_by_in_cabinet(candidates: List[RowStd], rfp_row: RowStd) -> List[RowStd]:
    """Сортирует кандидатов: сначала с совпадающим IN_CABINET, затем остальные."""
    if len(candidates) <= 1:
        return candidates
    rfp_ic = str(rfp_row.get_value(IN_CABINET) or "").strip()
    if not rfp_ic:
        return candidates
    return sorted(
        candidates,
        key=lambda c: (0 if _in_cabinet_matches(rfp_ic, c.get_value(IN_CABINET)) else 1)
    )


def add_mto_data_to_rfp_row(
    rfp_row: RowStd,
    mto_row: RowStd,
    qty_override: float = None):
    """
    Добавляет данные из MTO в строку RFP
    
    Args:
        rfp_row: Строка RFP для модификации
        mto_row: Строка MTO с данными
    """
    # Получаем тег из MTO
    mto_tags = mto_row.get_tags_list()
    mto_tag = mto_tags[0] if mto_tags else ""
    
    # Добавляем данные в соответствующие поля
    rfp_row.el[TAG_MTO].value = mto_tag
    rfp_row.el[CODE_MTO].value = mto_row.get_value(CODE) or ""
    rfp_row.el[CODE_MTO].struck_value = mto_row.el[CODE].struck_value
    rfp_row.el[MTO_CODE_STRUCK].value = mto_row.el[CODE].struck_value or ""
    rfp_row.el[NAME_MTO].value = mto_row.get_value(NAME) or ""
    rfp_row.el[TYPE_MARK_MTO].value = mto_row.get_value(TYPE_MARK) or ""
    if qty_override is None:
        rfp_row.el[VALUES_MTO].value = mto_row.get_value(VALUES) or ""
        # Помечаем MTO строку как полностью потребленную (полное совпадение по коду)
        mto_row._consumed_in_output = True
    else:
        rfp_row.el[VALUES_MTO].value = qty_override
        consume_mto_for_export(mto_row, _safe_float(qty_override))
    rfp_row.el[NUMBERS_MTO].value = mto_row.get_value(NUMBERS) or ""
    copy_mto_units_fields(rfp_row, mto_row)
    
    # Переносим IN_CABINET из MTO строки (объединяем если уже есть значение)
    mto_in_cabinet = str(mto_row.get_value(IN_CABINET) or "").strip()
    if mto_in_cabinet:
        current_cabinet = str(rfp_row.get_value(IN_CABINET) or "").strip()
        rfp_row.el[IN_CABINET].value = _merge_in_cabinet(current_cabinet, mto_in_cabinet)


def _is_mto_exceeds_rfp(rfp_value, mto_value) -> bool:
    try:
        rfp_val = float(rfp_value) if rfp_value not in (None, "") else None
    except (ValueError, TypeError):
        rfp_val = None
    try:
        mto_val = float(mto_value) if mto_value not in (None, "") else None
    except (ValueError, TypeError):
        mto_val = None
    
    if rfp_val is None or rfp_val == 0:
        return False
    if mto_val is None:
        return False
    return mto_val > rfp_val


def match_rfp_with_vo(
    rfp_rows: List[RowStd],
    vo_data: Dict[str, List[RowStd]],
    replacement_table: Optional[Dict[str, List[Tuple[str, str]]]] = None,
    mto_data: Optional[Dict[str, List[RowStd]]] = None,
    unmatched_mto_rows: Optional[List[RowStd]] = None,
    debug: bool = False,
    debug_tag: List[str] = None,
    debug_code: List[str] = None,
    debug_title_system: List[str] = None,
    debug_log: List[str] = None) -> tuple[Set[str], List[RowStd]]:
    """
    Сопоставляет строки RFP с VO по тегам в рамках одного title_system
    
    Args:
        rfp_rows: Список строк RFP для сопоставления
        vo_data: Словарь title_system -> список строк VO
        replacement_table: Таблица замен кодов (старый -> список новых кодов)
        mto_data: Словарь title_system -> список строк MTO (для поиска строк без тегов)
        unmatched_mto_rows: Список несопоставленных строк MTO (для поиска совпадений с VO по CODE+TAG)
        debug: Флаг отладки
        debug_tag: Тег для точечной отладки (AND с другими фильтрами)
        debug_code: Код для точечной отладки (AND с другими фильтрами)
        debug_title_system: title_system для точечной отладки (AND с другими фильтрами)
        
    Returns:
        Кортеж: (множество сопоставленных тегов, список несопоставленных строк VO)
    """
    matched_tags: Set[str] = set()
    matched_tag_sources: Dict[str, Set[str]] = defaultdict(set)
    unmatched_vo_rows: List[RowStd] = []
    used_vo_rows: Set[RowStd] = set()
    replacement_table = replacement_table or {}
    
    # Создаем индекс VO строк по title_system и тегам для быстрого поиска
    vo_index: Dict[str, Dict[str, List[RowStd]]] = defaultdict(lambda: defaultdict(list))
    # Индекс VO строк по title_system и коду
    vo_code_index: Dict[str, Dict[str, List[RowStd]]] = defaultdict(lambda: defaultdict(list))
    for title_system, vo_rows in vo_data.items():
        normalized_title_system = _normalize_title_system(title_system)
        for vo_row in vo_rows:
            if vo_row.row_type != RowType.position_row:
                continue
            vo_code = vo_row.get_value(CODE)
            vo_tags = vo_row.get_tags_list()
            if _should_debug(
                debug,
                normalized_title_system,
                _normalize_code(vo_code) if vo_code else "",
                vo_tags,
                debug_tag,
                debug_code,
                debug_title_system
            ):
                _debug_print(
                    debug_log,
                    f"  [DEBUG][match_rfp_with_vo][index_build_vo] VO строка в индексе: title_system={normalized_title_system}, "
                    f"CODE={vo_code}, TAGS={vo_tags}, VALUES={vo_row.get_value(VALUES)}, IN_CABINET={vo_row.get_value(IN_CABINET) or ''}"
                )
            if vo_code:
                normalized_code = _normalize_code(vo_code)
                if normalized_code:
                    vo_code_index[normalized_title_system][normalized_code].append(vo_row)
            if vo_tags:
                tag = vo_tags[0]
                vo_index[normalized_title_system][tag].append(vo_row)

    # Индекс несопоставленных MTO строк по (title_system, tag, code) для поиска совпадений с VO
    # Создаем обратный индекс: строка MTO -> title_system (из ключей mto_data)
    mto_row_to_title: Dict[RowStd, str] = {}
    if mto_data:
        for title_system, mto_rows_list in mto_data.items():
            normalized_ts = _normalize_title_system(title_system)
            for mto_row in mto_rows_list:
                mto_row_to_title[mto_row] = normalized_ts

    unmatched_mto_index: Dict[str, Dict[str, Dict[str, List[RowStd]]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(list))
    )
    used_unmatched_mto_rows: Set[RowStd] = set()
    if unmatched_mto_rows:
        for mto_row in unmatched_mto_rows:
            if mto_row.row_type != RowType.position_row:
                continue
            # Получаем title_system из обратного индекса (ключ словаря mto_data)
            mto_title_system = mto_row_to_title.get(mto_row, "")
            if not mto_title_system:
                continue
            mto_tags = mto_row.get_tags_list()
            mto_code = mto_row.get_value(CODE)
            mto_code_normalized = _normalize_code(mto_code) if mto_code else ""
            if mto_tags and mto_code_normalized:
                mto_tag = mto_tags[0]
                unmatched_mto_index[mto_title_system][mto_tag][mto_code_normalized].append(mto_row)
                if _should_debug(
                    debug,
                    mto_title_system,
                    mto_code_normalized,
                    mto_tags,
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rfp_with_vo][index_build_unmatched_mto] Unmatched MTO в индексе: title_system={mto_title_system}, "
                        f"CODE={mto_code}, TAG={mto_tag}, VALUES={mto_row.get_value(VALUES)}"
                    )

    mto_no_tag_by_title: Dict[str, Dict[str, List[RowStd]]] = defaultdict(lambda: defaultdict(list))
    mto_remaining_qty: Dict[RowStd, float] = {}
    if mto_data:
        for title_system, mto_rows in mto_data.items():
            normalized_title_system = _normalize_title_system(title_system)
            for mto_row in mto_rows:
                if mto_row.row_type != RowType.position_row:
                    continue
                if mto_row.get_tags_list():
                    continue
                mto_code = mto_row.get_value(CODE)
                if not mto_code:
                    continue
                normalized_code = _normalize_code(mto_code)
                if not normalized_code:
                    continue
                mto_no_tag_by_title[normalized_title_system][normalized_code].append(mto_row)
                if mto_row not in mto_remaining_qty:
                    used_qty = getattr(mto_row, "_used_qty_no_tag", 0.0)
                    remaining = _safe_float(mto_row.get_value(VALUES)) - _safe_float(used_qty)
                    mto_remaining_qty[mto_row] = max(0.0, remaining)
    
    # Инициализируем статус для всех строк
    for rfp_row in rfp_rows:
        if rfp_row.row_type != RowType.position_row:
            continue
        if is_excluded_from_supply(rfp_row):
            continue
        rfp_row.el[MATCH_STATUS_VO].value = "Не найден в VO"
    
    # ПРИОРИТЕТНЫЙ ПРОХОД: Резервируем VO строки, совпадающие с unmatched_mto_rows по CODE+TAG
    # Эти VO строки не будут использованы для сопоставления с RFP, т.к. соответствующие MTO
    # не нашли пару в RFP (теги в VO уникальны, поэтому если VO совпадает с MTO по тегу+коду,
    # она должна быть зарезервирована для этой MTO строки)
    if unmatched_mto_rows:
        for title_system, vo_rows in vo_data.items():
            normalized_title_system = _normalize_title_system(title_system)
            title_unmatched_mto = unmatched_mto_index.get(normalized_title_system, {})
            if not title_unmatched_mto:
                continue
            
            for vo_row in vo_rows:
                if vo_row.row_type != RowType.position_row:
                    continue
                if vo_row in used_vo_rows:
                    continue
                
                vo_tags = vo_row.get_tags_list()
                vo_code = vo_row.get_value(CODE)
                vo_code_normalized = _normalize_code(vo_code) if vo_code else ""
                
                if not vo_tags or not vo_code_normalized:
                    continue
                
                vo_tag = vo_tags[0]
                tag_mto_index = title_unmatched_mto.get(vo_tag, {})
                
                # Собираем кандидатов: прямое совпадение + через таблицу замен
                # (VO имеет новый код, MTO может иметь как новый, так и старый)
                mto_candidates = tag_mto_index.get(vo_code_normalized, [])
                replacement_match_code = None
                
                if not any(r not in used_unmatched_mto_rows for r in mto_candidates):
                    # Прямых кандидатов нет или все использованы — ищем через таблицу замен
                    # VO имеет новый код → находим старые коды → ищем MTO со старым кодом
                    old_codes = _get_reverse_replacement_codes(vo_code_normalized, replacement_table)
                    for old_code in old_codes:
                        old_code_normalized = _normalize_code(old_code)
                        if not old_code_normalized:
                            continue
                        repl_candidates = tag_mto_index.get(old_code_normalized, [])
                        available = [r for r in repl_candidates if r not in used_unmatched_mto_rows]
                        if available:
                            mto_candidates = repl_candidates
                            replacement_match_code = old_code_normalized
                            break
                
                for mto_row in mto_candidates:
                    if mto_row in used_unmatched_mto_rows:
                        continue
                    
                    # Нашли совпадение VO с unmatched MTO по CODE+TAG (или через таблицу замен)
                    # Резервируем VO строку - она не будет использована для сопоставления с RFP
                    used_vo_rows.add(vo_row)
                    used_unmatched_mto_rows.add(mto_row)
                    matched_tags.add(vo_tag)
                    matched_tag_sources[vo_tag].add(normalized_title_system)
                    
                    # Помечаем MTO строку, что VO найден (для последующей обработки в add_unmatched_mto_rows)
                    mto_row._matched_vo_row = vo_row
                    
                    if _should_debug(
                        debug,
                        normalized_title_system,
                        vo_code_normalized,
                        vo_tags,
                        debug_tag,
                        debug_code,
                        debug_title_system
                    ):
                        repl_info = (f", ЗАМЕНА: MTO_OLD_CODE={replacement_match_code} -> VO_NEW_CODE={vo_code_normalized}"
                                     if replacement_match_code else "")
                        mto_ic = mto_row.get_value(IN_CABINET) or ""
                        vo_ic = vo_row.get_value(IN_CABINET) or ""
                        _debug_print(
                            debug_log,
                            f"  [DEBUG][match_rfp_with_vo][reserve_vo_for_mto] VO зарезервирована для unmatched MTO: title_system={normalized_title_system}, "
                            f"VO_CODE={vo_code}, VO_TAG={vo_tag}, MTO_CODE={mto_row.get_value(CODE)}, "
                            f"MTO_TAG={mto_row.get_tags_list()}{repl_info}, "
                            f"MTO_IN_CABINET='{mto_ic}', VO_IN_CABINET='{vo_ic}'"
                        )
                    break

    # PASS 0: Сопоставляем VO с уже сопоставленной MTO частью (TAG_MTO + CODE_MTO)
    for rfp_row in rfp_rows:
        if rfp_row.row_type != RowType.position_row:
            continue
        if is_excluded_from_supply(rfp_row):
            continue
        if rfp_row.el[MATCH_STATUS_VO].value != "Не найден в VO":
            continue
        
        # Проверяем, есть ли сопоставленная MTO часть
        mto_tag_value = rfp_row.el.get(TAG_MTO)
        mto_code_value = rfp_row.el.get(CODE_MTO)
        
        mto_tag = mto_tag_value.value if mto_tag_value else ""
        mto_code = _normalize_code(mto_code_value.value if mto_code_value else "")
        
        if not mto_tag or not mto_code:
            continue
        
        title_system = _normalize_title_system(rfp_row.get_value(DS_TITLE))
        if not title_system:
            continue
        
        title_vo_index = vo_index.get(title_system, {})
        candidates = title_vo_index.get(mto_tag, [])
        if not candidates:
            continue
        candidates = _sort_candidates_by_in_cabinet(candidates, rfp_row)
        
        # Коды замены для MTO кода (MTO может иметь старый код, VO — новый)
        mto_replacement_codes_normalized = set()
        if replacement_table:
            for rc in _get_replacement_codes(mto_code, replacement_table):
                rc_normalized = _normalize_code(rc)
                if rc_normalized:
                    mto_replacement_codes_normalized.add(rc_normalized)
        
        vo_row = None
        replacement_match_info = ""
        for candidate in candidates:
            if candidate in used_vo_rows:
                continue
            vo_code = _normalize_code(candidate.get_value(CODE))
            if not vo_code:
                continue
            if vo_code == mto_code:
                vo_row = candidate
                break
            # Проверяем через таблицу замен: MTO имеет старый код, VO — новый
            if vo_code in mto_replacement_codes_normalized:
                vo_row = candidate
                replacement_match_info = f"MTO_OLD_CODE={mto_code} -> VO_NEW_CODE={vo_code}"
                break
        
        if vo_row is None:
            continue
        
        matched_tags.add(mto_tag)
        matched_tag_sources[mto_tag].add(title_system)
        used_vo_rows.add(vo_row)
        
        if _should_debug_cross_title(
            debug,
            title_system,
            vo_row.get_value(CODE) or "",
            vo_row.get_tags_list(),
            debug_tag,
            debug_code,
            debug_title_system
        ):
            repl_info = f", ЗАМЕНА: {replacement_match_info}" if replacement_match_info else ""
            rfp_ic = rfp_row.get_value(IN_CABINET) or ""
            vo_ic = vo_row.get_value(IN_CABINET) or ""
            _debug_print(
                debug_log,
                f"  [DEBUG][match_rfp_with_vo][match_mto_tag_code] VO использована (MTO тег+код): title_system={title_system}, "
                f"CODE={vo_row.get_value(CODE)}, TAG={vo_row.get_tags_list()}, "
                f"MTO_TAG={mto_tag}, MTO_CODE={mto_code}{repl_info}, "
                f"RFP_IN_CABINET='{rfp_ic}', VO_IN_CABINET='{vo_ic}'"
            )
        
        add_vo_data_to_rfp_row(rfp_row, vo_row)
        rfp_row.el[MATCH_STATUS_VO].value = "Тег сопоставлен"
        
        if replacement_match_info:
            _append_replacement_status(rfp_row, f"VO(MTO тег+код): {replacement_match_info}")
        
        if _should_debug(
            debug,
            rfp_row.get_value(DS_TITLE),
            rfp_row.get_value(CODE) or "",
            rfp_row.get_tags_list(),
            debug_tag,
            debug_code,
            debug_title_system
        ):
            repl_suffix = f" (замена: {replacement_match_info})" if replacement_match_info else ""
            _debug_print(
                debug_log,
                f"  Сопоставлено по MTO тегу+коду: MTO тег '{mto_tag}' (title_system: {title_system}) "
                f"с VO код '{vo_row.get_value(CODE)}'{repl_suffix}"
            )

    # PASS 0.5: Сопоставляем VO с RFP по MTO код + шкаф (без требования совпадения тегов)
    # Когда VO тег != MTO тег, но код и шкаф совпадают (напр. VO тег 1606, MTO тег 1604, код BCC0002122)
    for rfp_row in rfp_rows:
        if rfp_row.row_type != RowType.position_row:
            continue
        if is_excluded_from_supply(rfp_row):
            continue
        if rfp_row.el[MATCH_STATUS_VO].value != "Не найден в VO":
            continue

        mto_code_value = rfp_row.el.get(CODE_MTO)
        mto_code = _normalize_code(mto_code_value.value if mto_code_value else "")
        if not mto_code:
            continue

        title_system = _normalize_title_system(rfp_row.get_value(DS_TITLE))
        if not title_system:
            continue

        rfp_in_cabinet = str(rfp_row.get_value(IN_CABINET) or "").strip()
        if not rfp_in_cabinet:
            continue

        title_vo_code_index = vo_code_index.get(title_system, {})
        candidate_codes = [mto_code]
        if replacement_table:
            for rc in _get_replacement_codes(mto_code, replacement_table):
                rc_n = _normalize_code(rc)
                if rc_n and rc_n not in candidate_codes:
                    candidate_codes.append(rc_n)
            for orig in _get_reverse_replacement_codes(mto_code, replacement_table):
                orig_n = _normalize_code(orig)
                if orig_n and orig_n not in candidate_codes:
                    candidate_codes.append(orig_n)

        vo_row = None
        replacement_match_info = ""
        rfp_value = rfp_row.get_value(VALUES)
        for candidate_code in candidate_codes:
            candidates = title_vo_code_index.get(candidate_code, [])
            if not candidates:
                continue
            candidates = _sort_candidates_by_in_cabinet(candidates, rfp_row)
            for candidate in candidates:
                if candidate in used_vo_rows:
                    continue
                if not _in_cabinet_matches(rfp_in_cabinet, candidate.get_value(IN_CABINET)):
                    continue
                if _is_mto_exceeds_rfp(rfp_value, candidate.get_value(VALUES)):
                    continue
                vo_row = candidate
                if candidate_code != mto_code:
                    replacement_match_info = f"MTO_CODE={mto_code} <-> VO_CODE={candidate_code}"
                break
            if vo_row is not None:
                break

        if vo_row is None:
            continue

        mto_tag_value = rfp_row.el.get(TAG_MTO)
        mto_tag = mto_tag_value.value if mto_tag_value else ""
        if mto_tag:
            matched_tags.add(mto_tag)
            matched_tag_sources[mto_tag].add(title_system)
        vo_tags = vo_row.get_tags_list()
        if vo_tags:
            matched_tags.add(vo_tags[0])
            matched_tag_sources[vo_tags[0]].add(title_system)

        used_vo_rows.add(vo_row)
        add_vo_data_to_rfp_row(rfp_row, vo_row)
        rfp_row.el[MATCH_STATUS_VO].value = "Тег в VO заменен"

        if replacement_match_info:
            _append_replacement_status(rfp_row, f"VO(MTO код+шкаф): {replacement_match_info}")

        _try_match_mto_no_tag_for_vo(
            rfp_row,
            title_system,
            mto_no_tag_by_title,
            mto_remaining_qty,
            replacement_table=replacement_table,
            debug=debug,
            debug_tag=debug_tag,
            debug_code=debug_code,
            debug_title_system=debug_title_system,
            debug_log=debug_log
        )

        if _should_debug_cross_title(
            debug,
            title_system,
            vo_row.get_value(CODE) or "",
            vo_row.get_tags_list(),
            debug_tag,
            debug_code,
            debug_title_system
        ):
            repl_info = f", ЗАМЕНА: {replacement_match_info}" if replacement_match_info else ""
            rfp_ic = rfp_row.get_value(IN_CABINET) or ""
            vo_ic = vo_row.get_value(IN_CABINET) or ""
            _debug_print(
                debug_log,
                f"  [DEBUG][match_rfp_with_vo][match_mto_code_cabinet] VO использована (MTO код+шкаф): title_system={title_system}, "
                f"CODE={vo_row.get_value(CODE)}, TAG={vo_row.get_tags_list()}, MTO_CODE={mto_code}{repl_info}, "
                f"RFP_IN_CABINET='{rfp_ic}', VO_IN_CABINET='{vo_ic}'"
            )

        if _should_debug(
            debug,
            rfp_row.get_value(DS_TITLE),
            rfp_row.get_value(CODE) or "",
            rfp_row.get_tags_list(),
            debug_tag,
            debug_code,
            debug_title_system
        ):
            repl_suffix = f" (замена: {replacement_match_info})" if replacement_match_info else ""
            _debug_print(
                debug_log,
                f"  Сопоставлено по MTO код+шкаф: MTO код '{mto_code}' (title_system: {title_system}) "
                f"с VO тег '{vo_tags[0] if vo_tags else ''}'{repl_suffix}"
            )

    # PASS 0.6: Сопоставляем VO с RFP по CODE_MTO + replacement_code (без требования шкафа)
    # Фокус на позициях, существующих в MTO — расширяем PASS 0.5 за счёт таблицы замен
    for rfp_row in rfp_rows:
        if rfp_row.row_type != RowType.position_row:
            continue
        if is_excluded_from_supply(rfp_row):
            continue
        if rfp_row.el[MATCH_STATUS_VO].value != "Не найден в VO":
            continue

        mto_code_value = rfp_row.el.get(CODE_MTO)
        mto_code = _normalize_code(mto_code_value.value if mto_code_value else "")
        if not mto_code:
            continue

        title_system = _normalize_title_system(rfp_row.get_value(DS_TITLE))
        if not title_system:
            continue

        candidate_codes = [mto_code]
        if replacement_table:
            for rc in _get_replacement_codes(mto_code, replacement_table):
                rc_n = _normalize_code(rc)
                if rc_n and rc_n not in candidate_codes:
                    candidate_codes.append(rc_n)
            for orig in _get_reverse_replacement_codes(mto_code, replacement_table):
                orig_n = _normalize_code(orig)
                if orig_n and orig_n not in candidate_codes:
                    candidate_codes.append(orig_n)

        title_vo_code_index = vo_code_index.get(title_system, {})
        vo_row = None
        replacement_match_info = ""
        rfp_value = rfp_row.get_value(VALUES)

        for candidate_code in candidate_codes:
            candidates = title_vo_code_index.get(candidate_code, [])
            if not candidates:
                continue
            candidates = _sort_candidates_by_in_cabinet(candidates, rfp_row)
            for candidate in candidates:
                if candidate in used_vo_rows:
                    continue
                if _in_cabinet_conflict(rfp_row.get_value(IN_CABINET), candidate.get_value(IN_CABINET)):
                    continue
                if _is_mto_exceeds_rfp(rfp_value, candidate.get_value(VALUES)):
                    continue
                vo_row = candidate
                if candidate_code != mto_code:
                    replacement_match_info = f"MTO_CODE={mto_code} <-> VO_CODE={candidate_code}"
                break
            if vo_row is not None:
                break

        if vo_row is None:
            continue

        mto_tag_value = rfp_row.el.get(TAG_MTO)
        mto_tag = mto_tag_value.value if mto_tag_value else ""
        if mto_tag:
            matched_tags.add(mto_tag)
            matched_tag_sources[mto_tag].add(title_system)
        vo_tags = vo_row.get_tags_list()
        if vo_tags:
            matched_tags.add(vo_tags[0])
            matched_tag_sources[vo_tags[0]].add(title_system)

        used_vo_rows.add(vo_row)
        add_vo_data_to_rfp_row(rfp_row, vo_row)
        rfp_row.el[MATCH_STATUS_VO].value = "Тег в VO заменен"

        if replacement_match_info:
            _append_replacement_status(rfp_row, f"VO(MTO код+замена): {replacement_match_info}")

        _try_match_mto_no_tag_for_vo(
            rfp_row,
            title_system,
            mto_no_tag_by_title,
            mto_remaining_qty,
            replacement_table=replacement_table,
            debug=debug,
            debug_tag=debug_tag,
            debug_code=debug_code,
            debug_title_system=debug_title_system,
            debug_log=debug_log
        )

        if _should_debug_cross_title(
            debug,
            title_system,
            vo_row.get_value(CODE) or "",
            vo_row.get_tags_list(),
            debug_tag,
            debug_code,
            debug_title_system
        ):
            repl_info = f", ЗАМЕНА: {replacement_match_info}" if replacement_match_info else ""
            rfp_ic = rfp_row.get_value(IN_CABINET) or ""
            vo_ic = vo_row.get_value(IN_CABINET) or ""
            _debug_print(
                debug_log,
                f"  [DEBUG][match_rfp_with_vo][match_mto_code_replacement] VO использована (CODE_MTO+замена): "
                f"title_system={title_system}, CODE={vo_row.get_value(CODE)}, TAG={vo_row.get_tags_list()}, "
                f"MTO_CODE={mto_code}{repl_info}, RFP_IN_CABINET='{rfp_ic}', VO_IN_CABINET='{vo_ic}'"
            )

        if _should_debug(
            debug,
            rfp_row.get_value(DS_TITLE),
            rfp_row.get_value(CODE) or "",
            rfp_row.get_tags_list(),
            debug_tag,
            debug_code,
            debug_title_system
        ):
            repl_suffix = f" (замена: {replacement_match_info})" if replacement_match_info else ""
            _debug_print(
                debug_log,
                f"  Сопоставлено по CODE_MTO+замена: MTO код '{mto_code}' (title_system: {title_system}) "
                f"с VO тег '{vo_tags[0] if vo_tags else ''}'{repl_suffix}"
            )

    # PASS 0.7: Резервируем VO для unmatched MTO по CODE + replacement_code (теги могут отличаться)
    # Расширяем резервный проход 0: сопоставление по коду через таблицу замен
    if unmatched_mto_rows and replacement_table:
        for title_system, vo_rows in vo_data.items():
            normalized_title_system = _normalize_title_system(title_system)
            title_unmatched_mto = unmatched_mto_index.get(normalized_title_system, {})
            if not title_unmatched_mto:
                continue

            for vo_row in vo_rows:
                if vo_row.row_type != RowType.position_row:
                    continue
                if vo_row in used_vo_rows:
                    continue

                vo_code = vo_row.get_value(CODE)
                vo_code_normalized = _normalize_code(vo_code) if vo_code else ""
                if not vo_code_normalized:
                    continue

                candidate_codes = [vo_code_normalized]
                for rc in _get_replacement_codes(vo_code_normalized, replacement_table):
                    rc_n = _normalize_code(rc)
                    if rc_n and rc_n not in candidate_codes:
                        candidate_codes.append(rc_n)
                for orig in _get_reverse_replacement_codes(vo_code_normalized, replacement_table):
                    orig_n = _normalize_code(orig)
                    if orig_n and orig_n not in candidate_codes:
                        candidate_codes.append(orig_n)

                for mto_tag, code_index in title_unmatched_mto.items():
                    for candidate_code in candidate_codes:
                        mto_candidates = code_index.get(candidate_code, [])
                        available = [r for r in mto_candidates if r not in used_unmatched_mto_rows]
                        if not available:
                            continue

                        mto_row = None
                        for mto_cand in available:
                            if not _in_cabinet_conflict(mto_cand.get_value(IN_CABINET), vo_row.get_value(IN_CABINET)):
                                mto_row = mto_cand
                                break
                        if mto_row is None:
                            continue

                        used_vo_rows.add(vo_row)
                        used_unmatched_mto_rows.add(mto_row)
                        vo_tags = vo_row.get_tags_list()
                        vo_tag = vo_tags[0] if vo_tags else ""
                        if vo_tag:
                            matched_tags.add(vo_tag)
                            matched_tag_sources[vo_tag].add(normalized_title_system)

                        mto_row._matched_vo_row = vo_row

                        if _should_debug(
                            debug,
                            normalized_title_system,
                            vo_code_normalized,
                            vo_tags or [],
                            debug_tag,
                            debug_code,
                            debug_title_system
                        ):
                            repl_note = f" (CODE+замена: MTO={mto_row.get_value(CODE)}, VO={vo_code})"
                            mto_ic = mto_row.get_value(IN_CABINET) or ""
                            vo_ic = vo_row.get_value(IN_CABINET) or ""
                            _debug_print(
                                debug_log,
                                f"  [DEBUG][match_rfp_with_vo][reserve_vo_mto_code_replacement] VO зарезервирована для unmatched MTO: "
                                f"title_system={normalized_title_system}, VO_CODE={vo_code}, VO_TAG={vo_tag}, "
                                f"MTO_TAG={mto_tag}, MTO_CODE={mto_row.get_value(CODE)}{repl_note}, "
                                f"MTO_IN_CABINET='{mto_ic}', VO_IN_CABINET='{vo_ic}'"
                            )
                        break
                    else:
                        continue
                    break

    # PASS 1: Сопоставляем строки RFP с VO по тегу+коду (приоритет)
    for rfp_row in rfp_rows:
        if rfp_row.row_type != RowType.position_row:
            continue
        if is_excluded_from_supply(rfp_row):
            continue
        if rfp_row.el[MATCH_STATUS_VO].value != "Не найден в VO":
            continue

        rfp_tags = rfp_row.get_tags_list()
        if not rfp_tags:
            continue
        title_system = _normalize_title_system(rfp_row.get_value(DS_TITLE))
        if not title_system:
            continue

        rfp_code = _normalize_code(rfp_row.get_value(CODE))
        if not rfp_code:
            continue

        title_vo_index = vo_index.get(title_system, {})
        for rfp_tag in rfp_tags:
            candidates = title_vo_index.get(rfp_tag, [])
            if not candidates:
                continue
            candidates = _sort_candidates_by_in_cabinet(candidates, rfp_row)
            vo_row = None
            for candidate in candidates:
                if candidate in used_vo_rows:
                    continue
                if _in_cabinet_conflict(rfp_row.get_value(IN_CABINET), candidate.get_value(IN_CABINET)):
                    continue
                vo_code = _normalize_code(candidate.get_value(CODE))
                if not vo_code or vo_code != rfp_code:
                    continue
                vo_row = candidate
                break
            if vo_row is None:
                continue

            matched_tags.add(rfp_tag)
            matched_tag_sources[rfp_tag].add(title_system)
            used_vo_rows.add(vo_row)
            if _should_debug_cross_title(
                debug,
                title_system,
                vo_row.get_value(CODE) or "",
                vo_row.get_tags_list(),
                debug_tag,
                debug_code,
                debug_title_system
            ):
                rfp_ic = rfp_row.get_value(IN_CABINET) or ""
                vo_ic = vo_row.get_value(IN_CABINET) or ""
                _debug_print(
                    debug_log,
                    f"  [DEBUG][match_rfp_with_vo][match_tag_code] VO использована (тег+код): title_system={title_system}, "
                    f"CODE={vo_row.get_value(CODE)}, TAG={vo_row.get_tags_list()}, "
                    f"RFP_TAG={rfp_tag}, RFP_CODE={rfp_code}, RFP_IN_CABINET='{rfp_ic}', VO_IN_CABINET='{vo_ic}'"
                )

            add_vo_data_to_rfp_row(rfp_row, vo_row)
            rfp_row.el[MATCH_STATUS_VO].value = "Тег сопоставлен"
            _try_match_mto_no_tag_for_vo(
                rfp_row,
                title_system,
                mto_no_tag_by_title,
                mto_remaining_qty,
                replacement_table=replacement_table,
                debug=debug,
                debug_tag=debug_tag,
                debug_code=debug_code,
                debug_title_system=debug_title_system,
                debug_log=debug_log
            )

            if _should_debug(
                debug,
                rfp_row.get_value(DS_TITLE),
                rfp_row.get_value(CODE) or "",
                rfp_tags,
                debug_tag,
                debug_code,
                debug_title_system
            ):
                _debug_print(
                    debug_log,
                    f"  Сопоставлено по тегу+коду: RFP тег '{rfp_tag}' (title_system: {title_system}) "
                    f"с VO код '{vo_row.get_value(CODE)}'"
                )
            break

    # PASS 1.5: Сопоставляем VO по тегу + код через таблицу замен (только для RFP с тегами)
    for rfp_row in rfp_rows:
        if rfp_row.row_type != RowType.position_row:
            continue
        if is_excluded_from_supply(rfp_row):
            continue
        if rfp_row.el[MATCH_STATUS_VO].value != "Не найден в VO":
            continue

        rfp_tags = rfp_row.get_tags_list()
        if not rfp_tags:
            continue

        title_system = _normalize_title_system(rfp_row.get_value(DS_TITLE))
        if not title_system:
            continue

        rfp_code_raw = rfp_row.get_value(CODE) or ""
        rfp_code = _normalize_code(rfp_code_raw)
        if not rfp_code:
            continue

        # Получаем коды замены для RFP кода
        replacement_codes = _get_replacement_codes(rfp_code, replacement_table)
        if not replacement_codes:
            continue

        title_vo_index = vo_index.get(title_system, {})
        matched = False

        for rfp_tag in rfp_tags:
            candidates = title_vo_index.get(rfp_tag, [])
            if not candidates:
                continue
            candidates = _sort_candidates_by_in_cabinet(candidates, rfp_row)

            for replacement_code in replacement_codes:
                normalized_replacement = _normalize_code(replacement_code)
                if not normalized_replacement:
                    continue

                vo_row = None
                for candidate in candidates:
                    if candidate in used_vo_rows:
                        continue
                    if _in_cabinet_conflict(rfp_row.get_value(IN_CABINET), candidate.get_value(IN_CABINET)):
                        continue
                    vo_code = _normalize_code(candidate.get_value(CODE))
                    if not vo_code or vo_code != normalized_replacement:
                        continue
                    vo_row = candidate
                    break

                if vo_row is None:
                    continue

                matched_tags.add(rfp_tag)
                matched_tag_sources[rfp_tag].add(title_system)
                used_vo_rows.add(vo_row)

                add_vo_data_to_rfp_row(rfp_row, vo_row)
                rfp_row.el[MATCH_STATUS_VO].value = "Тег сопоставлен"

                _append_replacement_status(
                    rfp_row,
                    f"VO(тег+код): {rfp_code} -> {replacement_code}"
                )

                _try_match_mto_no_tag_for_vo(
                    rfp_row,
                    title_system,
                    mto_no_tag_by_title,
                    mto_remaining_qty,
                    replacement_table=replacement_table,
                    debug=debug,
                    debug_tag=debug_tag,
                    debug_code=debug_code,
                    debug_title_system=debug_title_system,
                    debug_log=debug_log
                )

                if _should_debug_cross_title(
                    debug,
                    title_system,
                    vo_row.get_value(CODE) or "",
                    vo_row.get_tags_list(),
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    rfp_ic = rfp_row.get_value(IN_CABINET) or ""
                    vo_ic = vo_row.get_value(IN_CABINET) or ""
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rfp_with_vo][match_tag_code_replacement] VO использована (тег+код замена): "
                        f"title_system={title_system}, CODE={vo_row.get_value(CODE)}, TAG={vo_row.get_tags_list()}, "
                        f"RFP_TAG={rfp_tag}, RFP_CODE={rfp_code_raw}, REPL={replacement_code}, "
                        f"RFP_IN_CABINET='{rfp_ic}', VO_IN_CABINET='{vo_ic}'"
                    )

                if _should_debug(
                    debug,
                    rfp_row.get_value(DS_TITLE),
                    rfp_code_raw,
                    rfp_tags,
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  Сопоставлено по тегу+коду (замена): RFP тег '{rfp_tag}' (title_system: {title_system}) "
                        f"с VO код '{vo_row.get_value(CODE)}' (RFP код '{rfp_code_raw}' -> '{replacement_code}')"
                    )
                matched = True
                break

            if matched:
                break

    # PASS 1.6: Сопоставляем VO по тегу + код через ОБРАТНУЮ таблицу замен (только для RFP с тегами)
    # RFP код A - ищем в таблице где A является заменой (X → A) - ищем VO с кодом X
    for rfp_row in rfp_rows:
        if rfp_row.row_type != RowType.position_row:
            continue
        if is_excluded_from_supply(rfp_row):
            continue
        if rfp_row.el[MATCH_STATUS_VO].value != "Не найден в VO":
            continue

        rfp_tags = rfp_row.get_tags_list()
        if not rfp_tags:
            continue

        title_system = _normalize_title_system(rfp_row.get_value(DS_TITLE))
        if not title_system:
            continue

        rfp_code_raw = rfp_row.get_value(CODE) or ""
        rfp_code = _normalize_code(rfp_code_raw)
        if not rfp_code:
            continue

        # Получаем оригинальные коды (обратный поиск)
        original_codes = _get_reverse_replacement_codes(rfp_code, replacement_table)
        if not original_codes:
            continue

        title_vo_index = vo_index.get(title_system, {})
        matched = False

        for rfp_tag in rfp_tags:
            candidates = title_vo_index.get(rfp_tag, [])
            if not candidates:
                continue
            candidates = _sort_candidates_by_in_cabinet(candidates, rfp_row)

            for original_code in original_codes:
                normalized_original = _normalize_code(original_code)
                if not normalized_original:
                    continue

                vo_row = None
                for candidate in candidates:
                    if candidate in used_vo_rows:
                        continue
                    if _in_cabinet_conflict(rfp_row.get_value(IN_CABINET), candidate.get_value(IN_CABINET)):
                        continue
                    vo_code = _normalize_code(candidate.get_value(CODE))
                    if not vo_code or vo_code != normalized_original:
                        continue
                    vo_row = candidate
                    break

                if vo_row is None:
                    continue

                matched_tags.add(rfp_tag)
                matched_tag_sources[rfp_tag].add(title_system)
                used_vo_rows.add(vo_row)

                add_vo_data_to_rfp_row(rfp_row, vo_row)
                rfp_row.el[MATCH_STATUS_VO].value = "Тег сопоставлен"

                _append_replacement_status(
                    rfp_row,
                    f"VO(тег+код обр.): {original_code} <- {rfp_code}"
                )

                _try_match_mto_no_tag_for_vo(
                    rfp_row,
                    title_system,
                    mto_no_tag_by_title,
                    mto_remaining_qty,
                    replacement_table=replacement_table,
                    debug=debug,
                    debug_tag=debug_tag,
                    debug_code=debug_code,
                    debug_title_system=debug_title_system,
                    debug_log=debug_log
                )

                if _should_debug_cross_title(
                    debug,
                    title_system,
                    vo_row.get_value(CODE) or "",
                    vo_row.get_tags_list(),
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    rfp_ic = rfp_row.get_value(IN_CABINET) or ""
                    vo_ic = vo_row.get_value(IN_CABINET) or ""
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rfp_with_vo][match_tag_code_reverse] VO использована (тег+код обратная замена): "
                        f"title_system={title_system}, CODE={vo_row.get_value(CODE)}, TAG={vo_row.get_tags_list()}, "
                        f"RFP_TAG={rfp_tag}, RFP_CODE={rfp_code_raw}, ORIG={original_code}, "
                        f"RFP_IN_CABINET='{rfp_ic}', VO_IN_CABINET='{vo_ic}'"
                    )

                if _should_debug(
                    debug,
                    rfp_row.get_value(DS_TITLE),
                    rfp_code_raw,
                    rfp_tags,
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  Сопоставлено по тегу+коду (обратная замена): RFP тег '{rfp_tag}' (title_system: {title_system}) "
                        f"с VO код '{vo_row.get_value(CODE)}' (VO код '{original_code}' <- RFP код '{rfp_code_raw}')"
                    )
                matched = True
                break

            if matched:
                break
    
    # Дополнительный проход: сопоставление по коду (RFP CODE или MTO CODE)
    for rfp_row in rfp_rows:
        if rfp_row.row_type != RowType.position_row:
            continue
        if is_excluded_from_supply(rfp_row):
            continue
        if rfp_row.el[MATCH_STATUS_VO].value != "Не найден в VO":
            continue
        title_system = _normalize_title_system(rfp_row.get_value(DS_TITLE))
        if not title_system:
            continue
        
        candidate_codes = []
        rfp_code = rfp_row.get_value(CODE)
        if rfp_code:
            normalized_code = _normalize_code(rfp_code)
            if normalized_code:
                candidate_codes.append(normalized_code)
        mto_code = rfp_row.el[CODE_MTO].value if rfp_row.el.get(CODE_MTO) else ""
        if mto_code:
            normalized_code = _normalize_code(mto_code)
            if normalized_code:
                candidate_codes.append(normalized_code)
        
        if not candidate_codes:
            continue
        
        rfp_value = rfp_row.get_value(VALUES)
        for candidate_code in candidate_codes:
            candidates = vo_code_index.get(title_system, {}).get(candidate_code, [])
            if not candidates:
                continue
            candidates = _sort_candidates_by_in_cabinet(candidates, rfp_row)

            matched_in_code = False
            for vo_row in candidates:
                if vo_row in used_vo_rows:
                    continue
                if _in_cabinet_conflict(rfp_row.get_value(IN_CABINET), vo_row.get_value(IN_CABINET)):
                    continue
                vo_value = vo_row.get_value(VALUES)
                if _is_mto_exceeds_rfp(rfp_value, vo_value):
                    continue

                add_vo_data_to_rfp_row(rfp_row, vo_row)
                rfp_row.el[MATCH_STATUS_VO].value = "Тег в VO заменен"
                used_vo_rows.add(vo_row)
                if _should_debug_cross_title(
                    debug,
                    title_system,
                    vo_row.get_value(CODE) or "",
                    vo_row.get_tags_list(),
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    rfp_ic = rfp_row.get_value(IN_CABINET) or ""
                    vo_ic = vo_row.get_value(IN_CABINET) or ""
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rfp_with_vo][match_code] VO использована (код): title_system={title_system}, "
                        f"CODE={vo_row.get_value(CODE)}, TAG={vo_row.get_tags_list()}, CAND_CODE={candidate_code}, "
                        f"RFP_IN_CABINET='{rfp_ic}', VO_IN_CABINET='{vo_ic}'"
                    )
                vo_tags = vo_row.get_tags_list()
                if vo_tags:
                    matched_tags.add(vo_tags[0])
                    matched_tag_sources[vo_tags[0]].add(title_system)

                if _should_debug(
                    debug,
                    rfp_row.get_value(DS_TITLE),
                    rfp_row.get_value(CODE) or "",
                    rfp_row.get_tags_list(),
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  Сопоставлено по коду: CODE '{candidate_code}' (title_system: {title_system}) "
                        f"с VO тег '{vo_tags[0] if vo_tags else ''}'"
                    )
                matched_in_code = True
                break

            if not matched_in_code:
                for vo_row in candidates:
                    if vo_row in used_vo_rows:
                        continue
                    if _in_cabinet_conflict(rfp_row.get_value(IN_CABINET), vo_row.get_value(IN_CABINET)):
                        continue

                    add_vo_data_to_rfp_row(rfp_row, vo_row)
                    rfp_row.el[MATCH_STATUS_VO].value = "Тег в VO заменен"
                    used_vo_rows.add(vo_row)
                    if _should_debug_cross_title(
                        debug,
                        title_system,
                        vo_row.get_value(CODE) or "",
                        vo_row.get_tags_list(),
                        debug_tag,
                        debug_code,
                        debug_title_system
                    ):
                        rfp_ic = rfp_row.get_value(IN_CABINET) or ""
                        vo_ic = vo_row.get_value(IN_CABINET) or ""
                        _debug_print(
                            debug_log,
                            f"  [DEBUG][match_rfp_with_vo][match_code_no_check] VO использована (код без проверки): title_system={title_system}, "
                            f"CODE={vo_row.get_value(CODE)}, TAG={vo_row.get_tags_list()}, CAND_CODE={candidate_code}, "
                            f"RFP_IN_CABINET='{rfp_ic}', VO_IN_CABINET='{vo_ic}'"
                        )
                    vo_tags = vo_row.get_tags_list()
                    if vo_tags:
                        matched_tags.add(vo_tags[0])
                        matched_tag_sources[vo_tags[0]].add(title_system)

                    if _should_debug(
                        debug,
                        rfp_row.get_value(DS_TITLE),
                        rfp_row.get_value(CODE) or "",
                        rfp_row.get_tags_list(),
                        debug_tag,
                        debug_code,
                        debug_title_system
                    ):
                        _debug_print(
                            debug_log,
                            f"  Сопоставлено по коду (без проверки количества): CODE '{candidate_code}' "
                            f"(title_system: {title_system}) с VO тег '{vo_tags[0] if vo_tags else ''}'"
                        )
                    matched_in_code = True
                    break

            if matched_in_code:
                break

    # Дополнительный проход: сопоставление по коду через таблицу замен
    for rfp_row in rfp_rows:
        if rfp_row.row_type != RowType.position_row:
            continue
        if is_excluded_from_supply(rfp_row):
            continue
        if rfp_row.el[MATCH_STATUS_VO].value != "Не найден в VO":
            continue
        title_system = _normalize_title_system(rfp_row.get_value(DS_TITLE))
        if not title_system:
            continue

        candidate_codes = []
        rfp_code = rfp_row.get_value(CODE)
        if rfp_code:
            normalized_code = _normalize_code(rfp_code)
            if normalized_code:
                candidate_codes.append(normalized_code)
        mto_code = rfp_row.el[CODE_MTO].value if rfp_row.el.get(CODE_MTO) else ""
        if mto_code:
            normalized_code = _normalize_code(mto_code)
            if normalized_code:
                candidate_codes.append(normalized_code)

        replacement_candidates = []
        for code in candidate_codes:
            for replacement_code in _get_replacement_codes(code, replacement_table):
                replacement_candidates.append((code, replacement_code))

        if not replacement_candidates:
            continue

        rfp_value = rfp_row.get_value(VALUES)
        for source_code, replacement_code in replacement_candidates:
            candidates = vo_code_index.get(title_system, {}).get(_normalize_code(replacement_code), [])
            if not candidates:
                continue
            candidates = _sort_candidates_by_in_cabinet(candidates, rfp_row)

            matched_in_code = False
            for vo_row in candidates:
                if vo_row in used_vo_rows:
                    continue
                if _in_cabinet_conflict(rfp_row.get_value(IN_CABINET), vo_row.get_value(IN_CABINET)):
                    continue
                vo_value = vo_row.get_value(VALUES)
                if _is_mto_exceeds_rfp(rfp_value, vo_value):
                    continue

                add_vo_data_to_rfp_row(rfp_row, vo_row)
                rfp_row.el[MATCH_STATUS_VO].value = "Тег в VO заменен"
                used_vo_rows.add(vo_row)
                if _should_debug_cross_title(
                    debug,
                    title_system,
                    vo_row.get_value(CODE) or "",
                    vo_row.get_tags_list(),
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    rfp_ic = rfp_row.get_value(IN_CABINET) or ""
                    vo_ic = vo_row.get_value(IN_CABINET) or ""
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rfp_with_vo][match_replacement] VO использована (замена): title_system={title_system}, "
                        f"CODE={vo_row.get_value(CODE)}, TAG={vo_row.get_tags_list()}, "
                        f"SRC={source_code}, REPL={replacement_code}, RFP_IN_CABINET='{rfp_ic}', VO_IN_CABINET='{vo_ic}'"
                    )
                _append_replacement_status(
                    rfp_row,
                    f"VO: {source_code} -> {replacement_code}"
                )
                vo_tags = vo_row.get_tags_list()
                if vo_tags:
                    matched_tags.add(vo_tags[0])
                    matched_tag_sources[vo_tags[0]].add(title_system)

                if _should_debug(
                    debug,
                    rfp_row.get_value(DS_TITLE),
                    rfp_row.get_value(CODE) or "",
                    rfp_row.get_tags_list(),
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  Сопоставлено по таблице замен: CODE '{source_code}' -> '{replacement_code}' "
                        f"(title_system: {title_system}) с VO тег '{vo_tags[0] if vo_tags else ''}'"
                    )
                matched_in_code = True
                break

            if not matched_in_code:
                for vo_row in candidates:
                    if vo_row in used_vo_rows:
                        continue
                    if _in_cabinet_conflict(rfp_row.get_value(IN_CABINET), vo_row.get_value(IN_CABINET)):
                        continue

                    add_vo_data_to_rfp_row(rfp_row, vo_row)
                    rfp_row.el[MATCH_STATUS_VO].value = "Тег в VO заменен"
                    used_vo_rows.add(vo_row)
                    if _should_debug_cross_title(
                        debug,
                        title_system,
                        vo_row.get_value(CODE) or "",
                        vo_row.get_tags_list(),
                        debug_tag,
                        debug_code,
                        debug_title_system
                    ):
                        rfp_ic = rfp_row.get_value(IN_CABINET) or ""
                        vo_ic = vo_row.get_value(IN_CABINET) or ""
                        _debug_print(
                            debug_log,
                            f"  [DEBUG][match_rfp_with_vo][match_replacement_no_check] VO использована (замена без проверки): title_system={title_system}, "
                            f"CODE={vo_row.get_value(CODE)}, TAG={vo_row.get_tags_list()}, "
                            f"SRC={source_code}, REPL={replacement_code}, RFP_IN_CABINET='{rfp_ic}', VO_IN_CABINET='{vo_ic}'"
                        )
                    _append_replacement_status(
                        rfp_row,
                        f"VO: {source_code} -> {replacement_code}"
                    )
                    vo_tags = vo_row.get_tags_list()
                    if vo_tags:
                        matched_tags.add(vo_tags[0])
                        matched_tag_sources[vo_tags[0]].add(title_system)

                    if _should_debug(
                        debug,
                        rfp_row.get_value(DS_TITLE),
                        rfp_row.get_value(CODE) or "",
                        rfp_row.get_tags_list(),
                        debug_tag,
                        debug_code,
                        debug_title_system
                    ):
                        _debug_print(
                            debug_log,
                            f"  Сопоставлено по таблице замен (без проверки количества): CODE '{source_code}' -> "
                            f"'{replacement_code}' (title_system: {title_system}) с VO тег '{vo_tags[0] if vo_tags else ''}'"
                        )
                    matched_in_code = True
                    break

            if matched_in_code:
                break
    
    # Собираем несопоставленные строки VO
    for title_system, vo_rows in vo_data.items():
        for vo_row in vo_rows:
            if vo_row.row_type != RowType.position_row:
                continue
            vo_tags = vo_row.get_tags_list()
            tag = vo_tags[0] if vo_tags else ""
            is_debug_row = _should_debug(
                debug,
                _normalize_title_system(title_system),
                _normalize_code(vo_row.get_value(CODE)) if vo_row.get_value(CODE) else "",
                vo_tags,
                debug_tag,
                debug_code,
                debug_title_system
            )
            if vo_row not in used_vo_rows:
                unmatched_vo_rows.append(vo_row)
                if is_debug_row:
                    note = ""
                    if tag and tag in matched_tags:
                        sources = sorted(matched_tag_sources.get(tag, set()))
                        note = f", NOTE=tag_matched:{','.join(sources) if sources else 'unknown'}"
                    vo_ic = vo_row.get_value(IN_CABINET) or ""
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rfp_with_vo][unmatched] VO несопоставлена: title_system={title_system}, "
                        f"CODE={vo_row.get_value(CODE)}, TAG={vo_tags}, "
                        f"VALUES={vo_row.get_value(VALUES)}, VO_IN_CABINET='{vo_ic}'{note}"
                    )
            elif is_debug_row:
                _debug_print(
                    debug_log,
                    f"  [DEBUG][match_rfp_with_vo][unmatched_skip] VO НЕ добавлена в несопоставленные: title_system={title_system}, "
                    f"CODE={vo_row.get_value(CODE)}, TAG={vo_tags}, REASON=row_used"
                )
    
    return matched_tags, unmatched_vo_rows


def add_vo_data_to_rfp_row(rfp_row: RowStd, vo_row: RowStd):
    """
    Добавляет данные из VO в строку RFP
    
    Args:
        rfp_row: Строка RFP для модификации
        vo_row: Строка VO с данными
    """
    vo_tags = vo_row.get_tags_list()
    vo_tag = vo_tags[0] if vo_tags else ""
    
    rfp_row.el[TAG_VO].value = vo_tag
    rfp_row.el[CODE_VO].value = vo_row.get_value(CODE) or ""
    rfp_row.el[NAME_VO].value = vo_row.get_value(NAME) or ""
    rfp_row.el[VALUES_VO].value = vo_row.get_value(VALUES) or ""
    
    # Переносим IN_CABINET из VO строки (объединяем если уже есть значение)
    vo_in_cabinet = str(vo_row.get_value(IN_CABINET) or "").strip()
    if vo_in_cabinet:
        current_cabinet = str(rfp_row.get_value(IN_CABINET) or "").strip()
        rfp_row.el[IN_CABINET].value = _merge_in_cabinet(current_cabinet, vo_in_cabinet)
