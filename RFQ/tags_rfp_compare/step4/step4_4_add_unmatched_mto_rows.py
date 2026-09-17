"""
Этап 4_4: Добавление несопоставленных строк MTO как новых строк
"""

from typing import List, Dict, Tuple, Optional
from collections import defaultdict
import os

_MULTIPROCESSING_MIN_TITLES = 4


def _mp_fill_vo(chunk):
    """Top-level wrapper for ProcessPoolExecutor (picklable)."""
    repl = chunk[4] if len(chunk) > 4 else None
    return _worker_fill_vo_rows_for_title(chunk[0], chunk[1], chunk[2], chunk[3], repl)


def _mp_fill_rfp_vo(chunk):
    """Top-level wrapper for ProcessPoolExecutor (picklable)."""
    return _worker_fill_rfp_vo_rows_for_title(chunk[0], chunk[1], chunk[2], chunk[3])

from base.base_classes import CheckElement, RowStd, RowType
from base.tables_columns import (
    DS_TITLE, MATCH_STATUS, HAS_TAGS_RFP, POSITION_STATUS,
    TAG_MTO, CODE_MTO, MTO_CODE_STRUCK, NAME_MTO, TYPE_MARK_MTO, NUMBERS_MTO, VALUES_MTO,
    UNITS_MTO, UNITS_CHECK_STATUS, UNITS_CONVERSION_TRACE,
    CODE, NAME, TYPE_MARK, NUMBERS, VALUES,
    TAG_VO, CODE_VO, NAME_VO, VALUES_VO, MATCH_STATUS_VO, MATCH_STATUS_CODE_REPLACEMENT,
    IN_CABINET
)
from RFQ.ds_compare.ds_units_normalize import copy_mto_units_fields
from RFQ.tags_rfp_compare.rfp_supply_status import is_excluded_from_supply
from RFQ.tags_rfp_compare.step4.step4_2_match_rfp_with_mto import (
    add_mto_data_to_rfp_row,
    _sort_candidates_by_in_cabinet,
    _in_cabinet_conflict,
    _in_cabinet_matches,
    _expand_no_tag_candidate_codes,
    _collect_mto_no_tag_rows_for_codes,
    _ref_cabinet_forbids_empty_mto_cabinet,
    remaining_mto_output_qty,
    consume_mto_for_export,
)


def _should_debug(
    debug: bool,
    title_system: str,
    code: str,
    tags: List[str],
    debug_tag: List[str],
    debug_code: List[str],
    debug_title_system: List[str]
) -> bool:
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


def add_unmatched_mto_rows(
    rfp_rows: List[RowStd],
    unmatched_mto_rows: List[RowStd],
    mto_data: Dict[str, List[RowStd]],
    debug: bool = False,
    debug_tag: List[str] = None,
    debug_code: List[str] = None,
    debug_title_system: List[str] = None,
    debug_log: List[str] = None,
    verbose_progress_messages: bool = True,
    include_mto_vo_without_rfp_anchor: bool = False,
):
    """
    Добавляет несопоставленные строки MTO как новые строки после последней строки RFP
    с тем же title_mark
    
    Args:
        rfp_rows: Список строк RFP (будет модифицирован)
        unmatched_mto_rows: Список несопоставленных строк MTO
        mto_data: Словарь title_mark -> список строк MTO (для определения title_mark)
        debug: Флаг отладки
        debug_tag: Тег для точечной отладки (AND с другими фильтрами)
        debug_code: Код для точечной отладки (AND с другими фильтрами)
        debug_title_system: title_system для точечной отладки (AND с другими фильтрами)
        include_mto_vo_without_rfp_anchor: Если True, при отсутствии RFP position-строки с тем же
            DS_TITLE добавлять несопоставленные MTO в конец списка (без якоря в RFP).
    """
    if not unmatched_mto_rows:
        return
    
    # Создаем обратный индекс: строка MTO -> title_mark
    mto_to_title: Dict[RowStd, str] = {}
    for title_mark, mto_rows in mto_data.items():
        for mto_row in mto_rows:
            mto_to_title[mto_row] = title_mark
    
    # Группируем несопоставленные строки MTO по title_mark
    unmatched_by_title: Dict[str, List[RowStd]] = defaultdict(list)
    for mto_row in unmatched_mto_rows:
        title_mark = mto_to_title.get(mto_row)
        if title_mark:
            unmatched_by_title[title_mark].append(mto_row)
    
    # Находим последние индексы строк RFP для каждого title_mark
    last_indices_by_title: Dict[str, int] = {}
    for idx, rfp_row in enumerate(rfp_rows):
        if rfp_row.row_type == RowType.position_row:
            title_mark = rfp_row.get_value(DS_TITLE)
            if title_mark:
                last_indices_by_title[title_mark] = idx
    
    rows_to_insert: List[tuple[int, RowStd]] = []
    appended_without_anchor = 0

    for title_mark, unmatched_rows in unmatched_by_title.items():
        last_idx = last_indices_by_title.get(title_mark)

        if last_idx is None:
            if not include_mto_vo_without_rfp_anchor:
                if _should_debug(
                    debug,
                    title_mark,
                    "",
                    [],
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  Предупреждение: не найдено строк с title_mark '{title_mark}' для вставки"
                    )
                continue
            base_rfp_row = None
            for mto_row in unmatched_rows:
                rem = remaining_mto_output_qty(mto_row)
                if rem <= 0:
                    continue
                matched_vo_row = getattr(mto_row, "_matched_vo_row", None)
                if _should_debug(
                    debug,
                    title_mark,
                    mto_row.get_value(CODE) or "",
                    mto_row.get_tags_list(),
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    vo_info = ""
                    if matched_vo_row:
                        vo_info = (
                            f", MATCHED_VO: CODE={matched_vo_row.get_value(CODE)}, "
                            f"TAG={matched_vo_row.get_tags_list()}"
                        )
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][add_unmatched_mto_rows][append_no_anchor] MTO без якоря RFP: title_mark={title_mark}, "
                        f"CODE={mto_row.get_value(CODE)}, TAGS={mto_row.get_tags_list()}, "
                        f"VALUES={mto_row.get_value(VALUES)}, remaining={rem}{vo_info}"
                    )
                new_row = create_rfp_row_from_mto(mto_row, base_rfp_row, title_mark)
                consume_mto_for_export(mto_row, rem)
                mto_row._added_via_unmatched_mto = True
                rfp_rows.append(new_row)
                appended_without_anchor += 1
            continue

        insert_idx = last_idx + 1
        while insert_idx < len(rfp_rows):
            next_row = rfp_rows[insert_idx]
            if next_row.row_type == RowType.position_row:
                next_title_mark = next_row.get_value(DS_TITLE)
                if next_title_mark and next_title_mark != title_mark:
                    break
            insert_idx += 1
        
        base_rfp_row = rfp_rows[last_idx] if 0 <= last_idx < len(rfp_rows) else None
        
        for mto_row in unmatched_rows:
            rem = remaining_mto_output_qty(mto_row)
            if rem <= 0:
                continue
            matched_vo_row = getattr(mto_row, "_matched_vo_row", None)
            if _should_debug(
                debug,
                title_mark,
                mto_row.get_value(CODE) or "",
                mto_row.get_tags_list(),
                debug_tag,
                debug_code,
                debug_title_system
            ):
                vo_info = ""
                if matched_vo_row:
                    vo_info = (
                        f", MATCHED_VO: CODE={matched_vo_row.get_value(CODE)}, "
                        f"TAG={matched_vo_row.get_tags_list()}"
                    )
                _debug_print(
                    debug_log,
                    f"  [DEBUG][add_unmatched_mto_rows][add_row] Добавление несопоставленной MTO строки: title_mark={title_mark}, "
                    f"CODE={mto_row.get_value(CODE)}, TAGS={mto_row.get_tags_list()}, "
                    f"VALUES={mto_row.get_value(VALUES)}, remaining={rem}{vo_info}"
                )
            new_row = create_rfp_row_from_mto(mto_row, base_rfp_row, title_mark)
            consume_mto_for_export(mto_row, rem)
            mto_row._added_via_unmatched_mto = True
            rows_to_insert.append((insert_idx, new_row))
            
            if _should_debug(
                debug,
                title_mark,
                mto_row.get_value(CODE) or "",
                mto_row.get_tags_list(),
                debug_tag,
                debug_code,
                debug_title_system
            ) and len(rows_to_insert) % 100 == 0:
                _debug_print(
                    debug_log,
                    f"  Подготовлено строк для вставки: {len(rows_to_insert)}..."
                )
    
    insert_groups: Dict[int, List[RowStd]] = defaultdict(list)
    for insert_idx, new_row in rows_to_insert:
        insert_groups[insert_idx].append(new_row)

    sorted_indices = sorted(insert_groups.keys(), reverse=True)
    if verbose_progress_messages and (rows_to_insert or appended_without_anchor):
        n_ins = len(rows_to_insert)
        if n_ins and appended_without_anchor:
            print(f"  Вставка {n_ins} несопоставленных строк (+ {appended_without_anchor} без якоря RFP в конец)...")
        elif n_ins:
            print(f"  Вставка {n_ins} несопоставленных строк...")
        else:
            print(f"  Добавлено в конец (без якоря RFP): {appended_without_anchor} несопоставленных строк MTO...")
    added_count = 0
    try:
        for insert_idx in sorted_indices:
            rows_for_insert = insert_groups[insert_idx]
            for new_row in reversed(rows_for_insert):
                rfp_rows.insert(insert_idx, new_row)
                added_count += 1
                if verbose_progress_messages and added_count % 500 == 0:
                    print(f"    Вставлено строк: {added_count}/{len(rows_to_insert)}...")

        if verbose_progress_messages:
            total = added_count + appended_without_anchor
            print(f"\n  Всего добавлено несопоставленных строк: {total}")
    except Exception as e:
        print(f"\n  ОШИБКА при вставке строк: {e}")
        print(f"  Успешно вставлено: {added_count} из {len(rows_to_insert)}")
        import traceback
        traceback.print_exc()
        raise


def add_unmatched_vo_rows(
    rfp_rows: List[RowStd],
    unmatched_vo_rows: List[RowStd],
    vo_data: Dict[str, List[RowStd]],
    mto_data: Dict[str, List[RowStd]],
    replacement_table: Optional[Dict[str, List[Tuple[str, str]]]] = None,
    debug: bool = False,
    debug_tag: List[str] = None,
    debug_code: List[str] = None,
    debug_title_system: List[str] = None,
    debug_log: List[str] = None,
    verbose_progress_messages: bool = True,
    include_mto_vo_without_rfp_anchor: bool = False,
):
    """
    Добавляет несопоставленные строки VO как новые строки после последней строки RFP
    с тем же title_mark
    
    Args:
        rfp_rows: Список строк RFP (будет модифицирован)
        unmatched_vo_rows: Список несопоставленных строк VO
        vo_data: Словарь title_mark -> список строк VO (для определения title_mark)
        debug: Флаг отладки
        debug_tag: Тег для точечной отладки (AND с другими фильтрами)
        debug_code: Код для точечной отладки (AND с другими фильтрами)
        debug_title_system: title_system для точечной отладки (AND с другими фильтрами)
        include_mto_vo_without_rfp_anchor: Если True, при отсутствии RFP position-строки с тем же
            title_system добавлять несопоставленные VO в конец списка.
    """
    if not unmatched_vo_rows:
        return
    
    # Сначала пытаемся сопоставить VO с уже добавленными строками из MTO
    _match_vo_with_added_mto_rows(
        rfp_rows,
        unmatched_vo_rows,
        vo_data,
        debug=debug,
        debug_tag=debug_tag,
        debug_code=debug_code,
        debug_title_system=debug_title_system,
        debug_log=debug_log
    )
    if not unmatched_vo_rows:
        return
    
    # Создаем обратный индекс: строка VO -> title_mark
    vo_to_title: Dict[RowStd, str] = {}
    for title_mark, vo_rows in vo_data.items():
        normalized_title = _normalize_title_system(title_mark)
        for vo_row in vo_rows:
            vo_to_title[vo_row] = normalized_title

    # Индекс MTO без тегов по title_mark и коду + остатки количества
    # Пропускаем строки с _consumed_in_output — они уже полностью использованы в match_rfp_with_mto
    mto_no_tag_by_title: Dict[str, Dict[str, List[RowStd]]] = defaultdict(lambda: defaultdict(list))
    mto_remaining_qty: Dict[RowStd, float] = {}
    for title_mark, mto_rows in mto_data.items():
        for mto_row in mto_rows:
            if mto_row.row_type != RowType.position_row:
                continue
            if mto_row.get_tags_list():
                continue
            if getattr(mto_row, "_consumed_in_output", False):
                continue
            mto_code = mto_row.get_value(CODE)
            if mto_code:
                normalized_code = _normalize_code(mto_code)
                if normalized_code:
                    mto_no_tag_by_title[title_mark][normalized_code].append(mto_row)
                    if mto_row not in mto_remaining_qty:
                        used_qty = _safe_float(getattr(mto_row, "_used_qty_no_tag", 0.0))
                        remaining = _safe_float(mto_row.get_value(VALUES)) - used_qty
                        mto_remaining_qty[mto_row] = max(0.0, remaining)
    
    # Группируем несопоставленные строки VO по title_mark
    unmatched_by_title: Dict[str, List[RowStd]] = defaultdict(list)
    for vo_row in unmatched_vo_rows:
        title_mark = vo_to_title.get(vo_row)
        if title_mark:
            unmatched_by_title[title_mark].append(vo_row)
    
    # Находим последние индексы строк RFP для каждого title_mark
    last_indices_by_title: Dict[str, int] = {}
    for idx, rfp_row in enumerate(rfp_rows):
        if rfp_row.row_type == RowType.position_row:
            title_mark = _normalize_title_system(rfp_row.get_value(DS_TITLE))
            if title_mark:
                last_indices_by_title[title_mark] = idx
    
    rows_to_insert: List[tuple[int, RowStd]] = []
    appended_without_anchor = 0

    for title_mark, unmatched_rows in unmatched_by_title.items():
        last_idx = last_indices_by_title.get(title_mark)

        if last_idx is None:
            if not include_mto_vo_without_rfp_anchor:
                if _should_debug(
                    debug,
                    title_mark,
                    "",
                    [],
                    debug_tag,
                    debug_code,
                    debug_title_system,
                ):
                    _debug_print(
                        debug_log,
                        f"  Предупреждение: не найдено строк с title_mark '{title_mark}' для вставки VO",
                    )
                continue
            base_rfp_row = None
            for vo_row in unmatched_rows:
                if _should_debug(
                    debug,
                    title_mark,
                    vo_row.get_value(CODE) or "",
                    vo_row.get_tags_list(),
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][add_unmatched_vo_rows][append_no_anchor] VO без якоря RFP: title_mark={title_mark}, "
                        f"CODE={vo_row.get_value(CODE)}, TAGS={vo_row.get_tags_list()}, "
                        f"VALUES={vo_row.get_value(VALUES)}, VO_IN_CABINET='{vo_row.get_value(IN_CABINET) or ''}'"
                    )
                new_row = create_rfp_row_from_vo(vo_row, base_rfp_row, title_mark)
                matched_mto_row = _find_mto_no_tag_match(
                    mto_no_tag_by_title,
                    mto_remaining_qty,
                    title_mark,
                    vo_row,
                    replacement_table=replacement_table,
                    debug=debug,
                    debug_tag=debug_tag,
                    debug_code=debug_code,
                    debug_title_system=debug_title_system,
                    debug_log=debug_log
                )
                if matched_mto_row:
                    add_mto_data_to_rfp_row(new_row, matched_mto_row, qty_override=1.0)
                    new_row.el[MATCH_STATUS].value = "Добавлен из МТО"
                rfp_rows.append(new_row)
                appended_without_anchor += 1
            continue

        insert_idx = last_idx + 1
        while insert_idx < len(rfp_rows):
            next_row = rfp_rows[insert_idx]
            if next_row.row_type == RowType.position_row:
                next_title_mark = next_row.get_value(DS_TITLE)
                if next_title_mark and next_title_mark != title_mark:
                    break
            insert_idx += 1
        base_rfp_row = rfp_rows[last_idx] if 0 <= last_idx < len(rfp_rows) else None
        
        for vo_row in unmatched_rows:
            if _should_debug(
                debug,
                title_mark,
                vo_row.get_value(CODE) or "",
                vo_row.get_tags_list(),
                debug_tag,
                debug_code,
                debug_title_system
            ):
                _debug_print(
                    debug_log,
                    f"  [DEBUG][add_unmatched_vo_rows][add_row] Добавление несопоставленной VO строки: title_mark={title_mark}, "
                    f"CODE={vo_row.get_value(CODE)}, TAGS={vo_row.get_tags_list()}, "
                    f"VALUES={vo_row.get_value(VALUES)}, VO_IN_CABINET='{vo_row.get_value(IN_CABINET) or ''}'"
                )
            new_row = create_rfp_row_from_vo(vo_row, base_rfp_row, title_mark)

            matched_mto_row = _find_mto_no_tag_match(
                mto_no_tag_by_title,
                mto_remaining_qty,
                title_mark,
                vo_row,
                replacement_table=replacement_table,
                debug=debug,
                debug_tag=debug_tag,
                debug_code=debug_code,
                debug_title_system=debug_title_system,
                debug_log=debug_log
            )
            if matched_mto_row:
                add_mto_data_to_rfp_row(new_row, matched_mto_row, qty_override=1.0)
                new_row.el[MATCH_STATUS].value = "Добавлен из МТО"
                # new_row.el[POSITION_STATUS].value = "Не протегирован в МТО"

            rows_to_insert.append((insert_idx, new_row))
    
    insert_groups: Dict[int, List[RowStd]] = defaultdict(list)
    for insert_idx, new_row in rows_to_insert:
        insert_groups[insert_idx].append(new_row)

    sorted_indices = sorted(insert_groups.keys(), reverse=True)
    if verbose_progress_messages and (rows_to_insert or appended_without_anchor):
        n_ins = len(rows_to_insert)
        if n_ins and appended_without_anchor:
            print(f"  Вставка {n_ins} несопоставленных строк VO (+ {appended_without_anchor} без якоря RFP в конец)...")
        elif n_ins:
            print(f"  Вставка {n_ins} несопоставленных строк VO...")
        else:
            print(f"  Добавлено в конец (без якоря RFP): {appended_without_anchor} несопоставленных строк VO...")
    added_count = 0
    try:
        for insert_idx in sorted_indices:
            rows_for_insert = insert_groups[insert_idx]
            for new_row in reversed(rows_for_insert):
                rfp_rows.insert(insert_idx, new_row)
                added_count += 1
                if verbose_progress_messages and added_count % 500 == 0:
                    print(f"    Вставлено строк VO: {added_count}/{len(rows_to_insert)}...")

        if verbose_progress_messages:
            total = added_count + appended_without_anchor
            print(f"\n  Всего добавлено несопоставленных строк VO: {total}")
    except Exception as e:
        print(f"\n  ОШИБКА при вставке строк VO: {e}")
        print(f"  Успешно вставлено: {added_count} из {len(rows_to_insert)}")
        import traceback
        traceback.print_exc()
        raise


def create_rfp_row_from_mto(
    mto_row: RowStd, 
    base_rfp_row: RowStd = None,
    title_mark: str = None
) -> RowStd:
    """
    Создает новую строку RFP на основе строки MTO
    Правая часть (RFP столбцы) должна быть пустой, кроме DS_TITLE (title_mark)
    
    Если на mto_row есть атрибут _matched_vo_row, также добавляет данные из VO.
    
    Args:
        mto_row: Строка MTO
        base_rfp_row: Базовая строка RFP для копирования структуры (опционально, не используется)
        title_mark: Значение title_mark для установки в новую строку
        
    Returns:
        Новая строка RowStd
    """
    # Создаем новую пустую строку
    new_row = RowStd(mto_row.t_com)
    new_row.row_type = RowType.position_row
    
    # Устанавливаем только title_mark из RFP столбцов
    if title_mark:
        new_row.el[DS_TITLE].value = title_mark
    
    # Все остальные RFP столбцы остаются пустыми
    
    # Копируем данные из MTO в соответствующие поля (правая часть)
    mto_tags = mto_row.get_tags_list()
    mto_tag = mto_tags[0] if mto_tags else ""
    
    # Заполняем данные MTO
    new_row.el[TAG_MTO].value = mto_tag
    new_row.el[CODE_MTO].value = mto_row.get_value(CODE) or ""
    new_row.el[NAME_MTO].value = mto_row.get_value(NAME) or ""
    new_row.el[TYPE_MARK_MTO].value = mto_row.get_value(TYPE_MARK) or ""
    remaining_mto_qty = remaining_mto_output_qty(mto_row)
    new_row.el[VALUES_MTO].value = (
        remaining_mto_qty if remaining_mto_qty > 0 else (mto_row.get_value(VALUES) or "")
    )
    new_row.el[NUMBERS_MTO].value = mto_row.get_value(NUMBERS) or ""
    copy_mto_units_fields(new_row, mto_row)
    
    # Копируем IN_CABINET из MTO строки
    mto_in_cabinet = str(mto_row.get_value(IN_CABINET) or "").strip()
    if mto_in_cabinet:
        new_row.el[IN_CABINET].value = mto_in_cabinet
    
    # Устанавливаем статус "Добавлен из МТО"
    new_row.el[MATCH_STATUS].value = "Добавлен из МТО"
    # Для строк, добавленных из MTO, отмечаем как новые
    new_row.el[HAS_TAGS_RFP].value = "Нов."
    # Статус позиции для строк, добавленных из MTO - они не были в RFP
    new_row.el[POSITION_STATUS].value = ""
    
    # Если есть сопоставленная VO строка (найдена в match_rfp_with_vo по CODE+TAG)
    matched_vo_row = getattr(mto_row, "_matched_vo_row", None)
    if matched_vo_row:
        vo_tags = matched_vo_row.get_tags_list()
        vo_tag = vo_tags[0] if vo_tags else ""
        new_row.el[TAG_VO].value = vo_tag
        new_row.el[CODE_VO].value = matched_vo_row.get_value(CODE) or ""
        new_row.el[NAME_VO].value = matched_vo_row.get_value(NAME) or ""
        new_row.el[VALUES_VO].value = matched_vo_row.get_value(VALUES) or ""
        new_row.el[MATCH_STATUS_VO].value = "Тег сопоставлен"
        # Объединяем IN_CABINET из VO строки
        vo_in_cabinet = str(matched_vo_row.get_value(IN_CABINET) or "").strip()
        if vo_in_cabinet:
            current_cabinet = str(new_row.el[IN_CABINET].value or "").strip()
            new_row.el[IN_CABINET].value = _merge_in_cabinet(current_cabinet, vo_in_cabinet)
    
    return new_row


def create_rfp_row_from_vo(
    vo_row: RowStd,
    base_rfp_row: RowStd = None,
    title_mark: str = None) -> RowStd:
    """
    Создает новую строку RFP на основе строки VO
    Правая часть (RFP столбцы) должна быть пустой, кроме DS_TITLE (title_mark)
    
    Args:
        vo_row: Строка VO
        base_rfp_row: Базовая строка RFP для копирования структуры (опционально, не используется)
        title_mark: Значение title_mark для установки в новую строку
        
    Returns:
        Новая строка RowStd
    """
    new_row = RowStd(vo_row.t_com)
    new_row.row_type = RowType.position_row
    
    if title_mark:
        new_row.el[DS_TITLE].value = title_mark
    
    vo_tags = vo_row.get_tags_list()
    vo_tag = vo_tags[0] if vo_tags else ""
    
    new_row.el[TAG_VO].value = vo_tag
    new_row.el[CODE_VO].value = vo_row.get_value(CODE) or ""
    new_row.el[NAME_VO].value = vo_row.get_value(NAME) or ""
    new_row.el[VALUES_VO].value = vo_row.get_value(VALUES) or ""
    
    # Копируем IN_CABINET из VO строки
    vo_in_cabinet = str(vo_row.get_value(IN_CABINET) or "").strip()
    if vo_in_cabinet:
        new_row.el[IN_CABINET].value = vo_in_cabinet
    
    new_row.el[MATCH_STATUS_VO].value = "Добавлен из VO"
    new_row.el[HAS_TAGS_RFP].value = "Нов."
    new_row.el[POSITION_STATUS].value = ""
    
    return new_row


def _match_vo_with_added_mto_rows(
    rfp_rows: List[RowStd],
    unmatched_vo_rows: List[RowStd],
    vo_data: Dict[str, List[RowStd]],
    debug: bool = False,
    debug_tag: str = None,
    debug_code: str = None,
    debug_title_system: str = None,
    debug_log: List[str] = None):
    """
    Пытается сопоставить строки VO с уже добавленными строками из MTO
    (MATCH_STATUS = "Добавлен из МТО") сначала по тегу, затем по коду.
    """
    # Индексы VO по title_system, тегу и коду
    vo_by_tag: Dict[str, Dict[str, List[RowStd]]] = defaultdict(lambda: defaultdict(list))
    vo_by_code: Dict[str, Dict[str, List[RowStd]]] = defaultdict(lambda: defaultdict(list))
    
    vo_to_title: Dict[RowStd, str] = {}
    for title_mark, vo_rows in vo_data.items():
        for vo_row in vo_rows:
            vo_to_title[vo_row] = title_mark
    
    for vo_row in unmatched_vo_rows:
        title_mark = vo_to_title.get(vo_row)
        if not title_mark:
            continue
        vo_tags = vo_row.get_tags_list()
        if vo_tags:
            vo_by_tag[title_mark][vo_tags[0]].append(vo_row)
        vo_code = vo_row.get_value(CODE)
        if vo_code:
            vo_by_code[title_mark][str(vo_code)].append(vo_row)
    
    used_vo_rows: set[RowStd] = set()
    
    for rfp_row in rfp_rows:
        if rfp_row.row_type != RowType.position_row:
            continue
        if rfp_row.el[MATCH_STATUS].value != "Добавлен из МТО":
            continue
        
        # Пропускаем строки, у которых уже заполнены VO данные
        # (например, из _matched_vo_row в create_rfp_row_from_mto)
        existing_vo_tag = rfp_row.el[TAG_VO].value if rfp_row.el.get(TAG_VO) else ""
        existing_vo_code = rfp_row.el[CODE_VO].value if rfp_row.el.get(CODE_VO) else ""
        if existing_vo_tag or existing_vo_code:
            continue
        
        title_mark = rfp_row.get_value(DS_TITLE)
        if not title_mark:
            continue
        
        # Сначала пробуем сопоставить по тегу
        mto_tag = rfp_row.el[TAG_MTO].value if rfp_row.el.get(TAG_MTO) else ""
        matched_vo_row = None
        rfp_in_cabinet = rfp_row.get_value(IN_CABINET)
        if mto_tag and mto_tag in vo_by_tag.get(title_mark, {}):
            for candidate in vo_by_tag[title_mark][mto_tag]:
                if candidate not in used_vo_rows:
                    if _in_cabinet_conflict(rfp_in_cabinet, candidate.get_value(IN_CABINET)):
                        continue
                    matched_vo_row = candidate
                    break
        
        # Если по тегу не нашли - пробуем по коду
        if matched_vo_row is None:
            mto_code = rfp_row.el[CODE_MTO].value if rfp_row.el.get(CODE_MTO) else ""
            if mto_code and str(mto_code) in vo_by_code.get(title_mark, {}):
                for candidate in vo_by_code[title_mark][str(mto_code)]:
                    if candidate not in used_vo_rows:
                        if _in_cabinet_conflict(rfp_in_cabinet, candidate.get_value(IN_CABINET)):
                            continue
                        matched_vo_row = candidate
                        break
        
        if matched_vo_row is None:
            continue
        
        # Заполняем VO данные в строке RFP
        vo_tags = matched_vo_row.get_tags_list()
        vo_tag = vo_tags[0] if vo_tags else ""
        rfp_row.el[TAG_VO].value = vo_tag
        rfp_row.el[CODE_VO].value = matched_vo_row.get_value(CODE) or ""
        rfp_row.el[NAME_VO].value = matched_vo_row.get_value(NAME) or ""
        rfp_row.el[VALUES_VO].value = matched_vo_row.get_value(VALUES) or ""
        rfp_row.el[MATCH_STATUS_VO].value = "Тег сопоставлен"
        # Переносим IN_CABINET из VO строки
        vo_in_cabinet = str(matched_vo_row.get_value(IN_CABINET) or "").strip()
        if vo_in_cabinet:
            current_cabinet = str(rfp_row.get_value(IN_CABINET) or "").strip()
            rfp_row.el[IN_CABINET].value = _merge_in_cabinet(current_cabinet, vo_in_cabinet)
        
        used_vo_rows.add(matched_vo_row)
        if _should_debug(
            debug,
            title_mark,
            rfp_row.el[CODE_MTO].value if rfp_row.el.get(CODE_MTO) else "",
            [mto_tag] if mto_tag else [],
            debug_tag,
            debug_code,
            debug_title_system
        ):
            _debug_print(
                debug_log,
                f"  [DEBUG][match_added_vo_rows_with_mto_by_code][match_added_mto] Сопоставлено VO с добавленной MTO строкой: "
                f"TAG_MTO='{mto_tag}', CODE_MTO='{rfp_row.el[CODE_MTO].value}'"
            )
    
    if used_vo_rows:
        unmatched_vo_rows[:] = [row for row in unmatched_vo_rows if row not in used_vo_rows]


def _normalize_code(code_value) -> str:
    if code_value is None:
        return ""
    return str(code_value).strip().upper()


def _normalize_code_cached(code_value, cache: Dict) -> str:
    """Кэшированная нормализация кода для ускорения повторных вызовов."""
    key = "__none__" if code_value is None else str(code_value)
    if key not in cache:
        cache[key] = _normalize_code(code_value)
    return cache[key]


def _normalize_title_system(title_value) -> str:
    if title_value is None:
        return ""
    return str(title_value).strip()


def _safe_float(value) -> float:
    try:
        return float(value) if value not in (None, "") else 0.0
    except (ValueError, TypeError):
        return 0.0


def _find_mto_no_tag_match(
    mto_no_tag_by_title: Dict[str, Dict[str, List[RowStd]]],
    mto_remaining_qty: Dict[RowStd, float],
    title_mark: str,
    vo_row: RowStd,
    replacement_table: Optional[Dict[str, List[Tuple[str, str]]]] = None,
    debug: bool = False,
    debug_tag: str = None,
    debug_code: str = None,
    debug_title_system: str = None,
    debug_log: List[str] = None) -> RowStd:
    if not title_mark:
        return None
    vo_code = vo_row.get_value(CODE)
    if not vo_code:
        return None
    normalized_code = _normalize_code(vo_code)
    if not normalized_code:
        return None

    candidate_codes = _expand_no_tag_candidate_codes(normalized_code, replacement_table)
    candidates = _collect_mto_no_tag_rows_for_codes(
        title_mark, candidate_codes, mto_no_tag_by_title)
    if not candidates:
        return None

    _is_debug = _should_debug(
        debug,
        title_mark,
        normalized_code,
        vo_row.get_tags_list(),
        debug_tag,
        debug_code,
        debug_title_system,
    )
    vo_ic = str(vo_row.get_value(IN_CABINET) or "").strip()
    candidates = [
        r for r in candidates
        if not _ref_cabinet_forbids_empty_mto_cabinet(vo_ic, r.get_value(IN_CABINET))
    ]
    if not candidates:
        return None

    if vo_ic and len(candidates) > 1:
        candidates = sorted(
            candidates,
            key=lambda r: (
                0 if _in_cabinet_matches(vo_ic, str(r.get_value(IN_CABINET) or "").strip())
                else (1 if not str(r.get_value(IN_CABINET) or "").strip() else 2)
            ),
        )
    else:
        candidates = _sort_candidates_by_in_cabinet(candidates, vo_row)
    for mto_row in candidates:
        remaining = mto_remaining_qty.get(mto_row, 0.0)
        mto_ic = mto_row.get_value(IN_CABINET) or ""
        if remaining < 1.0:
            continue
        if _in_cabinet_conflict(vo_ic, mto_ic):
            if _is_debug:
                _debug_print(
                    debug_log,
                    f"  [DEBUG][_try_match_mto_no_tag_for_vo][skip_cabinet] Пропуск MTO (конфликт шкафа): title_mark={title_mark}, "
                    f"VO_CODE={normalized_code}, MTO_CODE={mto_row.get_value(CODE)}, VO_IN_CABINET='{vo_ic}', MTO_IN_CABINET='{mto_ic}', остаток={remaining}"
                )
            continue
        mto_remaining_qty[mto_row] = remaining - 1.0
        if _is_debug:
            mto_c = _normalize_code(mto_row.get_value(CODE))
            _debug_print(
                debug_log,
                f"  [DEBUG][_try_match_mto_no_tag_for_vo][match_no_tag] VO->MTO (без тегов) по коду: title_mark={title_mark}, "
                f"VO_CODE={normalized_code}, MTO_CODE={mto_c}, VO_IN_CABINET='{vo_ic}', MTO_IN_CABINET='{mto_ic}', остаток={mto_remaining_qty[mto_row]}"
            )
        return mto_row

    return None


def _build_mto_index_for_code_match(
    mto_data: Dict[str, List[RowStd]],
    code_cache: Dict
) -> Tuple[Dict[str, Dict[str, List[RowStd]]], Dict[RowStd, float]]:
    """Строит индекс MTO по title_mark и коду + остатки количества. Общий для обеих функций."""
    mto_by_title: Dict[str, Dict[str, List[RowStd]]] = defaultdict(lambda: defaultdict(list))
    mto_remaining_qty: Dict[RowStd, float] = {}
    for title_mark, mto_rows in mto_data.items():
        for mto_row in mto_rows:
            if mto_row.row_type != RowType.position_row:
                continue
            if getattr(mto_row, "_consumed_in_output", False):
                continue
            mto_code = mto_row.get_value(CODE)
            if not mto_code:
                continue
            normalized_code = _normalize_code_cached(mto_code, code_cache)
            if not normalized_code:
                continue
            mto_by_title[title_mark][normalized_code].append(mto_row)
            if mto_row not in mto_remaining_qty:
                used_qty = _safe_float(getattr(mto_row, "_used_qty_no_tag", 0.0))
                remaining = _safe_float(mto_row.get_value(VALUES)) - used_qty
                mto_remaining_qty[mto_row] = max(0.0, remaining)
    return mto_by_title, mto_remaining_qty


def _copy_mto_el_to_row(src_row: RowStd, dest_row: RowStd) -> None:
    """Копирует MTO-поля из src_row в dest_row (для слияния результатов multiprocessing)."""
    for att in (
        TAG_MTO, CODE_MTO, MTO_CODE_STRUCK, NAME_MTO, TYPE_MARK_MTO,
        NUMBERS_MTO, VALUES_MTO, UNITS_MTO, UNITS_CHECK_STATUS, UNITS_CONVERSION_TRACE,
    ):
        if not src_row.el.get(att):
            continue
        if att not in dest_row.el:
            dest_row.el[att] = CheckElement(None)
        dest_row.el[att].value = src_row.el[att].value
        if att == CODE_MTO and hasattr(src_row.el[att], "struck_value"):
            dest_row.el[att].struck_value = src_row.el[att].struck_value
    for att in (MATCH_STATUS, POSITION_STATUS, IN_CABINET):
        if dest_row.el.get(att) and src_row.el.get(att):
            dest_row.el[att].value = src_row.el[att].value


def _worker_fill_vo_rows_for_title(
    title_mark: str,
    indexed_rows: List[Tuple[int, RowStd]],
    mto_rows: List[RowStd],
    mto_qtys: List[float],
    replacement_table: Optional[Dict[str, List[Tuple[str, str]]]] = None,
) -> Tuple[List[Tuple[int, RowStd]], List[Tuple[int, float]]]:
    """
    Worker для multiprocessing: заполняет MTO для VO-строк одного title_mark.
    Возвращает (modified_rows, consumed) где consumed = [(mto_idx, qty), ...].
    """
    from RFQ.tags_rfp_compare.step4.step4_2_match_rfp_with_mto import (
        add_mto_data_to_rfp_row,
        _sort_candidates_by_in_cabinet,
        _in_cabinet_conflict,
        _expand_no_tag_candidate_codes,
        _ref_cabinet_forbids_empty_mto_cabinet,
    )
    mto_idx_map = {id(m): i for i, m in enumerate(mto_rows)}
    mto_remaining_qty = {mto_row: qty for mto_row, qty in zip(mto_rows, mto_qtys)}
    mto_by_code: Dict[str, List[Tuple[RowStd, int]]] = {}
    for i, mto_row in enumerate(mto_rows):
        code = mto_row.get_value(CODE)
        if not code:
            continue
        norm = str(code).strip().upper() if code else ""
        if norm:
            mto_by_code.setdefault(norm, []).append((mto_row, i))
    consumed: List[Tuple[int, float]] = []
    result = []
    for idx, row in indexed_rows:
        code_vo = row.el[CODE_VO].value if row.el.get(CODE_VO) else ""
        if not code_vo:
            result.append((idx, row))
            continue
        norm = str(code_vo).strip().upper()
        if not norm:
            result.append((idx, row))
            continue
        code_keys = _expand_no_tag_candidate_codes(norm, replacement_table)
        seen_m = set()
        candidates_with_idx: List[Tuple[RowStd, int]] = []
        for ck in code_keys:
            for pair in mto_by_code.get(ck, []):
                mto_row, midx = pair
                rid = id(mto_row)
                if rid in seen_m:
                    continue
                seen_m.add(rid)
                candidates_with_idx.append(pair)
        candidates = [c[0] for c in candidates_with_idx]
        row_ic = row.get_value(IN_CABINET)
        candidates = [
            r for r in candidates
            if not _ref_cabinet_forbids_empty_mto_cabinet(row_ic, r.get_value(IN_CABINET))
        ]
        candidates = _sort_candidates_by_in_cabinet(candidates, row)
        for mto_row in candidates:
            remaining = mto_remaining_qty.get(mto_row, 0.0)
            if remaining < 1.0:
                continue
            if _in_cabinet_conflict(row_ic, mto_row.get_value(IN_CABINET)):
                continue
            mto_idx = mto_idx_map.get(id(mto_row))
            mto_remaining_qty[mto_row] = remaining - 1.0
            add_mto_data_to_rfp_row(row, mto_row, qty_override=1.0)
            if not row.el[MATCH_STATUS].value:
                row.el[MATCH_STATUS].value = "Добавлен из МТО"
            # row.el[POSITION_STATUS].value = "Не протегирован в МТО"
            if mto_idx is not None:
                consumed.append((mto_idx, 1.0))
            break
        result.append((idx, row))
    return result, consumed


def _worker_fill_rfp_vo_rows_for_title(
    title_mark: str,
    indexed_rows: List[Tuple[int, RowStd]],
    mto_rows: List[RowStd],
    mto_qtys: List[float]
) -> Tuple[List[Tuple[int, RowStd]], List[Tuple[int, float]]]:
    """
    Worker для multiprocessing: заполняет MTO для RFP+VO строк одного title_mark.
    Возвращает (modified_rows, consumed) где consumed = [(mto_idx, qty), ...].
    """
    from RFQ.tags_rfp_compare.step4.step4_2_match_rfp_with_mto import (
        add_mto_data_to_rfp_row,
        _sort_candidates_by_in_cabinet,
        _in_cabinet_conflict,
    )
    mto_idx_map = {id(m): i for i, m in enumerate(mto_rows)}
    mto_remaining_qty = {mto_row: qty for mto_row, qty in zip(mto_rows, mto_qtys)}
    mto_by_code: Dict[str, List[RowStd]] = {}
    for mto_row in mto_rows:
        code = mto_row.get_value(CODE)
        if not code:
            continue
        norm = str(code).strip().upper() if code else ""
        if norm:
            mto_by_code.setdefault(norm, []).append(mto_row)
    consumed: List[Tuple[int, float]] = []
    result = []
    for idx, row in indexed_rows:
        rfp_code = row.get_value(CODE)
        vo_code = row.el[CODE_VO].value if row.el.get(CODE_VO) else ""
        candidate_codes = []
        for cv in (rfp_code, vo_code):
            n = str(cv).strip().upper() if cv else ""
            if n and n not in candidate_codes:
                candidate_codes.append(n)
        row_ic = row.get_value(IN_CABINET)
        matched = False
        for candidate_code in candidate_codes:
            candidates = mto_by_code.get(candidate_code, [])
            candidates = _sort_candidates_by_in_cabinet(candidates, row)
            for mto_row in candidates:
                remaining = mto_remaining_qty.get(mto_row, 0.0)
                if remaining < 1.0:
                    continue
                if _in_cabinet_conflict(row_ic, mto_row.get_value(IN_CABINET)):
                    continue
                mto_idx = mto_idx_map.get(id(mto_row))
                mto_remaining_qty[mto_row] = remaining - 1.0
                add_mto_data_to_rfp_row(row, mto_row, qty_override=1.0)
                if row.el[MATCH_STATUS].value in (None, "", "Не найден в МТО"):
                    row.el[MATCH_STATUS].value = "Тег в МТО заменен"
                # if row.el[POSITION_STATUS].value in (None, ""):
                #     row.el[POSITION_STATUS].value = "Не протегирован в МТО"
                if mto_idx is not None:
                    consumed.append((mto_idx, 1.0))
                matched = True
                break
            if matched:
                break
        result.append((idx, row))
    return result, consumed


def _consume_mto_qty_from_rfp_rows(
    rfp_rows: List[RowStd],
    mto_by_title: Dict[str, Dict[str, List[RowStd]]],
    mto_remaining_qty: Dict[RowStd, float],
    code_cache: Dict
) -> None:
    """Списывает уже использованные количества по MTO в результатах."""
    consume_rows = [
        row for row in rfp_rows
        if row.row_type == RowType.position_row
        and row.get_value(DS_TITLE)
        and (row.el[CODE_MTO].value if row.el.get(CODE_MTO) else "")
    ]
    for row in consume_rows:
        title_mark = row.get_value(DS_TITLE)
        code_mto = row.el[CODE_MTO].value if row.el.get(CODE_MTO) else ""
        normalized_code = _normalize_code_cached(code_mto, code_cache)
        if not normalized_code:
            continue
        used_qty = _safe_float(row.el[VALUES_MTO].value if row.el.get(VALUES_MTO) else 0)
        if used_qty <= 0:
            used_qty = 1.0
        candidates = mto_by_title.get(title_mark, {}).get(normalized_code, [])
        for mto_row in candidates:
            remaining = mto_remaining_qty.get(mto_row, 0.0)
            if remaining <= 0:
                continue
            take = min(remaining, used_qty)
            mto_remaining_qty[mto_row] = remaining - take
            used_qty -= take
            if used_qty <= 0:
                break


def match_added_vo_rows_with_mto_by_code(
    rfp_rows: List[RowStd],
    mto_data: Dict[str, List[RowStd]],
    mto_by_title: Optional[Dict[str, Dict[str, List[RowStd]]]] = None,
    mto_remaining_qty: Optional[Dict[RowStd, float]] = None,
    replacement_table: Optional[Dict[str, List[Tuple[str, str]]]] = None,
    debug: bool = False,
    debug_tag: List[str] = None,
    debug_code: List[str] = None,
    debug_title_system: List[str] = None,
    debug_log: List[str] = None,
    use_multiprocessing: bool = False
) -> Tuple[Optional[Dict[str, Dict[str, List[RowStd]]]], Optional[Dict[RowStd, float]]]:
    """
    Доп. шаг: для строк, добавленных из VO без тегов, пытается найти MTO по коду.
    Использует остатки количества по MTO строкам (VALUES).
    Если переданы mto_by_title и mto_remaining_qty — использует их (для объединённого вызова).
    Возвращает (mto_by_title, mto_remaining_qty) для передачи в match_rows_with_vo_and_rfp_code_missing_mto.
    """
    if not mto_data or not rfp_rows:
        return None, None

    code_cache: Dict = {}
    if mto_by_title is None or mto_remaining_qty is None:
        mto_by_title, mto_remaining_qty = _build_mto_index_for_code_match(mto_data, code_cache)
        _consume_mto_qty_from_rfp_rows(rfp_rows, mto_by_title, mto_remaining_qty, code_cache)

    vo_rows_to_fill = [
        row for row in rfp_rows
        if row.row_type == RowType.position_row
        and (row.el[MATCH_STATUS_VO].value if row.el.get(MATCH_STATUS_VO) else "") == "Добавлен из VO"
        and not (row.el[TAG_VO].value if row.el.get(TAG_VO) else "")
        and (row.el[CODE_VO].value if row.el.get(CODE_VO) else "")
        and not (row.el[CODE_MTO].value if row.el.get(CODE_MTO) else "")
    ]
    if use_multiprocessing and len(mto_by_title) >= _MULTIPROCESSING_MIN_TITLES:
        try:
            from concurrent.futures import ProcessPoolExecutor
            title_to_rows: Dict[str, List[Tuple[int, RowStd]]] = defaultdict(list)
            for idx, row in enumerate(rfp_rows):
                if row not in vo_rows_to_fill:
                    continue
                tm = row.get_value(DS_TITLE)
                if tm:
                    title_to_rows[tm].append((idx, row))
            chunks = []
            for tm, rows_tuples in title_to_rows.items():
                mto_rows_tm = []
                for code, lst in mto_by_title.get(tm, {}).items():
                    mto_rows_tm.extend(lst)
                qtys = [mto_remaining_qty.get(m, 0.0) for m in mto_rows_tm]
                if mto_rows_tm and any(q > 0 for q in qtys):
                    chunks.append((tm, rows_tuples, mto_rows_tm, qtys, replacement_table))
            if chunks:
                max_workers = min(len(chunks), (os.cpu_count() or 4) - 1 or 1)
                with ProcessPoolExecutor(max_workers=max_workers) as ex:
                    for chunk, res in zip(chunks, ex.map(_mp_fill_vo, chunks)):
                        mod_rows, consumed = res
                        for idx, mod_row in mod_rows:
                            _copy_mto_el_to_row(mod_row, rfp_rows[idx])
                        mto_rows_tm = chunk[2]
                        for mto_idx, qty in consumed:
                            mto_row = mto_rows_tm[mto_idx]
                            mto_remaining_qty[mto_row] = mto_remaining_qty.get(mto_row, 0.0) - qty
                            consume_mto_for_export(mto_row, qty)
                return mto_by_title, mto_remaining_qty
        except Exception as e:
            print(f"  [match_added_vo] multiprocessing fallback to sequential: {e}")
    for row in vo_rows_to_fill:
        title_mark = row.get_value(DS_TITLE)
        if not title_mark:
            continue
        code_vo = row.el[CODE_VO].value if row.el.get(CODE_VO) else ""
        normalized_code = _normalize_code_cached(code_vo, code_cache)
        if not normalized_code:
            continue

        code_keys = _expand_no_tag_candidate_codes(normalized_code, replacement_table)
        title_bucket = mto_by_title.get(title_mark, {})
        seen_id = set()
        candidates = []
        for ck in code_keys:
            for mto_row in title_bucket.get(ck, []):
                rid = id(mto_row)
                if rid in seen_id:
                    continue
                seen_id.add(rid)
                candidates.append(mto_row)
        row_in_cabinet = row.get_value(IN_CABINET)
        candidates = [
            r for r in candidates
            if not _ref_cabinet_forbids_empty_mto_cabinet(row_in_cabinet, r.get_value(IN_CABINET))
        ]
        if not candidates:
            if _should_debug(
                debug,
                title_mark,
                normalized_code,
                row.get_tags_list(),
                debug_tag,
                debug_code,
                debug_title_system
            ):
                _debug_print(
                    debug_log,
                    f"  [DEBUG][match_added_vo_rows_with_mto_by_code][not_found] VO->MTO не найдено по коду (доп. шаг): title_mark={title_mark}, CODE={normalized_code}"
                )
            continue
        candidates = _sort_candidates_by_in_cabinet(candidates, row)
        matched = False
        for mto_row in candidates:
            remaining = mto_remaining_qty.get(mto_row, 0.0)
            if remaining < 1.0:
                continue
            if _in_cabinet_conflict(row_in_cabinet, mto_row.get_value(IN_CABINET)):
                continue

            mto_remaining_qty[mto_row] = remaining - 1.0
            add_mto_data_to_rfp_row(row, mto_row, qty_override=1.0)
            if not row.el[MATCH_STATUS].value:
                row.el[MATCH_STATUS].value = "Добавлен из МТО"
            # row.el[POSITION_STATUS].value = "Не протегирован в МТО"

            if _should_debug(
                debug,
                title_mark,
                normalized_code,
                row.get_tags_list(),
                debug_tag,
                debug_code,
                debug_title_system
            ):
                _debug_print(
                    debug_log,
                    f"  [DEBUG][match_added_vo_rows_with_mto_by_code][match_by_code] VO->MTO по коду (доп. шаг): title_mark={title_mark}, CODE={normalized_code}"
                )
            matched = True
            break

        if not matched and _should_debug(
            debug,
            title_mark,
            normalized_code,
            row.get_tags_list(),
            debug_tag,
            debug_code,
            debug_title_system
        ):
            _debug_print(
                debug_log,
                f"  [DEBUG][match_added_vo_rows_with_mto_by_code][not_found] VO->MTO не найдено по коду (доп. шаг): title_mark={title_mark}, CODE={normalized_code}"
            )
    return mto_by_title, mto_remaining_qty


def match_rows_with_vo_and_rfp_code_missing_mto(
    rfp_rows: List[RowStd],
    mto_data: Dict[str, List[RowStd]],
    mto_by_title: Optional[Dict[str, Dict[str, List[RowStd]]]] = None,
    mto_remaining_qty: Optional[Dict[RowStd, float]] = None,
    debug: bool = False,
    debug_tag: List[str] = None,
    debug_code: List[str] = None,
    debug_title_system: List[str] = None,
    debug_log: List[str] = None,
    use_multiprocessing: bool = False
):
    """
    Доп. шаг: для строк, где есть RFP CODE и VO CODE, но нет MTO CODE,
    пытается найти MTO по коду и заполнить.
    Если переданы mto_by_title и mto_remaining_qty — использует их (индекс уже построен и consume выполнен).
    """
    if not mto_data or not rfp_rows:
        return

    code_cache: Dict = {}
    if mto_by_title is None or mto_remaining_qty is None:
        mto_by_title, mto_remaining_qty = _build_mto_index_for_code_match(mto_data, code_cache)
        _consume_mto_qty_from_rfp_rows(rfp_rows, mto_by_title, mto_remaining_qty, code_cache)

    rfp_vo_rows_missing_mto = [
        row for row in rfp_rows
        if row.row_type == RowType.position_row
        and not is_excluded_from_supply(row)
        and row.get_value(CODE)
        and (row.el[CODE_VO].value if row.el.get(CODE_VO) else "")
        and not (row.el[CODE_MTO].value if row.el.get(CODE_MTO) else "")
    ]
    for row in rfp_vo_rows_missing_mto:
        title_mark = row.get_value(DS_TITLE)
        if not title_mark:
            continue

        rfp_code = row.get_value(CODE)
        vo_code = row.el[CODE_VO].value if row.el.get(CODE_VO) else ""
        candidate_codes: List[str] = []
        for code_value in (rfp_code, vo_code):
            normalized = _normalize_code_cached(code_value, code_cache)
            if normalized and normalized not in candidate_codes:
                candidate_codes.append(normalized)

        if not candidate_codes:
            continue

        matched = False
        row_in_cabinet = row.get_value(IN_CABINET)
        for candidate_code in candidate_codes:
            candidates = mto_by_title.get(title_mark, {}).get(candidate_code, [])
            candidates = _sort_candidates_by_in_cabinet(candidates, row)
            for mto_row in candidates:
                remaining = mto_remaining_qty.get(mto_row, 0.0)
                if remaining < 1.0:
                    continue
                if _in_cabinet_conflict(row_in_cabinet, mto_row.get_value(IN_CABINET)):
                    continue

                mto_remaining_qty[mto_row] = remaining - 1.0
                add_mto_data_to_rfp_row(row, mto_row, qty_override=1.0)

                if row.el[MATCH_STATUS].value in (None, "", "Не найден в МТО"):
                    row.el[MATCH_STATUS].value = "Тег в МТО заменен"
                if row.el[POSITION_STATUS].value in (None, ""):
                    row.el[POSITION_STATUS].value = "Не протегирован в МТО"

                if _should_debug(
                    debug,
                    title_mark,
                    candidate_code,
                    row.get_tags_list(),
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_rows_with_vo_and_rfp_code_missing_mto][match_by_code] RFP+VO->MTO по коду: title_mark={title_mark}, CODE={candidate_code}"
                    )
                matched = True
                break
            if matched:
                break

        code_for_debug = ""
        if debug_code and debug_code in candidate_codes:
            code_for_debug = debug_code
        elif not debug_code and candidate_codes:
            code_for_debug = candidate_codes[0]
        if not matched and _should_debug(
            debug,
            title_mark,
            code_for_debug,
            row.get_tags_list(),
            debug_tag,
            debug_code,
            debug_title_system
        ):
            _debug_print(
                debug_log,
                f"  [DEBUG][match_rows_with_vo_and_rfp_code_missing_mto][not_found] RFP+VO->MTO не найдено по коду: title_mark={title_mark}, "
                f"CODES={candidate_codes}"
            )


def _get_reverse_replacement_codes(
    vo_code: str,
    replacement_table: Dict[str, List[Tuple[str, str]]]
) -> List[str]:
    """
    Обратный поиск в таблице замен: по коду замены находит оригинальные коды.
    
    Args:
        vo_code: Код из VO (возможно, замененный код)
        replacement_table: Таблица замен (оригинальный код -> список (новый код, описание))
        
    Returns:
        Список оригинальных кодов, которые заменяются на vo_code
    """
    if not vo_code or not replacement_table:
        return []
    
    vo_code_normalized = vo_code.strip().upper()
    original_codes = []
    
    for original_code, replacements in replacement_table.items():
        for new_code, _ in replacements:
            if new_code and new_code.strip().upper() == vo_code_normalized:
                original_codes.append(original_code)
                break
    
    return original_codes


def match_unmatched_vo_with_mto_by_reverse_replacement(
    rfp_rows: List[RowStd],
    mto_data: Dict[str, List[RowStd]],
    replacement_table: Optional[Dict[str, List[Tuple[str, str]]]] = None,
    debug: bool = False,
    debug_tag: List[str] = None,
    debug_code: List[str] = None,
    debug_title_system: List[str] = None,
    debug_log: List[str] = None,
    verbose_progress_messages: bool = True,
):
    """
    Дополнительный шаг: для строк, добавленных из VO без MTO,
    ищет MTO по обратной таблице замен.
    
    Если VO код является заменой для какого-то оригинального кода,
    ищем свободные MTO строки с этим оригинальным кодом.
    
    Пример: VO код BCC0003171, в таблице замен BCC0000764 -> BCC0003171
    Ищем MTO строки с кодом BCC0000764
    
    Args:
        rfp_rows: Список строк результата (будет модифицирован)
        mto_data: Словарь title_mark -> список строк MTO
        replacement_table: Таблица замен кодов
        debug: Флаг отладки
        debug_tag: Список тегов для отладки
        debug_code: Список кодов для отладки
        debug_title_system: Список title_system для отладки
        debug_log: Список строк debug-лога
    """
    if not mto_data or not rfp_rows or not replacement_table:
        return
    
    # Индекс MTO по title_mark и коду + остатки количества
    mto_by_title: Dict[str, Dict[str, List[RowStd]]] = defaultdict(lambda: defaultdict(list))
    mto_remaining_qty: Dict[RowStd, float] = {}
    
    for title_mark, mto_rows in mto_data.items():
        for mto_row in mto_rows:
            if mto_row.row_type != RowType.position_row:
                continue
            if getattr(mto_row, "_consumed_in_output", False):
                continue
            mto_code = mto_row.get_value(CODE)
            if not mto_code:
                continue
            normalized_code = _normalize_code(mto_code)
            if not normalized_code:
                continue
            mto_by_title[title_mark][normalized_code].append(mto_row)
            if mto_row not in mto_remaining_qty:
                # _used_qty_no_tag уже содержит информацию об использовании из предыдущих функций
                # (match_added_vo_rows_with_mto_by_code, match_rows_with_vo_and_rfp_code_missing_mto)
                # Дополнительный цикл по rfp_rows не нужен - он приводит к двойному учёту
                used_qty = _safe_float(getattr(mto_row, "_used_qty_no_tag", 0.0))
                remaining = _safe_float(mto_row.get_value(VALUES)) - used_qty
                mto_remaining_qty[mto_row] = max(0.0, remaining)
    
    matched_count = 0
    
    # Ищем VO строки без MTO и пытаемся найти MTO через обратную таблицу замен
    for row in rfp_rows:
        if row.row_type != RowType.position_row:
            continue
        
        # Только строки, добавленные из VO
        if row.el[MATCH_STATUS_VO].value != "Добавлен из VO":
            continue
        
        # Пропускаем, если уже есть MTO
        if row.el.get(CODE_MTO) and row.el[CODE_MTO].value:
            continue
        
        title_mark = row.get_value(DS_TITLE)
        if not title_mark:
            continue
        
        code_vo = row.el[CODE_VO].value if row.el.get(CODE_VO) else ""
        if not code_vo:
            continue
        
        normalized_vo_code = _normalize_code(code_vo)
        if not normalized_vo_code:
            continue
        
        # Обратный поиск: находим оригинальные коды для этого VO кода
        original_codes = _get_reverse_replacement_codes(normalized_vo_code, replacement_table)
        
        if not original_codes:
            continue
        
        if _should_debug(
            debug,
            title_mark,
            normalized_vo_code,
            row.get_tags_list(),
            debug_tag,
            debug_code,
            debug_title_system
        ):
            _debug_print(
                debug_log,
                f"  [DEBUG][match_unmatched_vo_with_mto_by_reverse_replacement][find_codes] Обратный поиск замены: VO_CODE={normalized_vo_code} -> "
                f"оригинальные коды: {original_codes}"
            )
        
        matched = False
        used_original_code = None
        
        for original_code in original_codes:
            normalized_original = _normalize_code(original_code)
            if not normalized_original:
                continue
            
            candidates = mto_by_title.get(title_mark, {}).get(normalized_original, [])
            
            for mto_row in candidates:
                remaining = mto_remaining_qty.get(mto_row, 0.0)
                if remaining < 1.0:
                    continue
                if _in_cabinet_conflict(row.get_value(IN_CABINET), mto_row.get_value(IN_CABINET)):
                    continue
                
                # Нашли свободную MTO строку с оригинальным кодом
                mto_remaining_qty[mto_row] = remaining - 1.0
                add_mto_data_to_rfp_row(row, mto_row, qty_override=1.0)
                
                if not row.el[MATCH_STATUS].value:
                    row.el[MATCH_STATUS].value = "Тег в МТО заменен"
                
                # Устанавливаем статус позиции для корректного учета сумм MTO без тегов
                # if row.el[POSITION_STATUS].value in (None, ""):
                #     row.el[POSITION_STATUS].value = "Не протегирован в МТО"
                
                # Добавляем информацию о замене
                _append_replacement_status(
                    row,
                    f"MTO(обр.): {original_code} <- {normalized_vo_code}"
                )
                
                if _should_debug(
                    debug,
                    title_mark,
                    normalized_vo_code,
                    row.get_tags_list(),
                    debug_tag,
                    debug_code,
                    debug_title_system
                ):
                    _debug_print(
                        debug_log,
                        f"  [DEBUG][match_unmatched_vo_with_mto_by_reverse_replacement][match] VO->MTO (обратная замена): title_mark={title_mark}, "
                        f"VO_CODE={normalized_vo_code}, MTO_CODE={normalized_original}"
                    )
                
                matched = True
                used_original_code = normalized_original
                matched_count += 1
                break
            
            if matched:
                break
        
        if not matched and _should_debug(
            debug,
            title_mark,
            normalized_vo_code,
            row.get_tags_list(),
            debug_tag,
            debug_code,
            debug_title_system
        ):
            _debug_print(
                debug_log,
                f"  [DEBUG][match_unmatched_vo_with_mto_by_reverse_replacement][not_found] VO->MTO (обратная замена) не найдено: title_mark={title_mark}, "
                f"VO_CODE={normalized_vo_code}, искали MTO с кодами: {original_codes}"
            )
    
    if verbose_progress_messages and matched_count > 0:
        print(f"    Сопоставлено через обратную таблицу замен: {matched_count}")


def _append_replacement_status(rfp_row: RowStd, message: str) -> None:
    """Добавляет сообщение о замене кода в статус."""
    if not message:
        return
    current = rfp_row.el[MATCH_STATUS_CODE_REPLACEMENT].value or ""
    if current:
        rfp_row.el[MATCH_STATUS_CODE_REPLACEMENT].value = f"{current}; {message}"
    else:
        rfp_row.el[MATCH_STATUS_CODE_REPLACEMENT].value = message


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


def add_unmatched_mto_cabinet_rows(
    result_rows: List[RowStd],
    mto_data: Dict[str, List[RowStd]],
    debug: bool = False,
    debug_tag: List[str] = None,
    debug_code: List[str] = None,
    debug_title_system: List[str] = None,
    debug_log: List[str] = None,
    verbose_progress_messages: bool = True,
    include_mto_vo_without_rfp_anchor: bool = False,
):
    """
    Добавляет несопоставленные строки MTO с непустым IN_CABINET в финальный вывод.
    Обрабатывает только строки БЕЗ тегов (строки с тегами уже обрабатываются через add_unmatched_mto_rows).
    
    Строка считается несопоставленной если:
    - _consumed_in_output не установлен (не была полностью сопоставлена по коду)
    - Оставшееся количество (VALUES - _used_qty_no_tag) > 0
    
    Args:
        result_rows: Список строк результата (будет модифицирован)
        mto_data: Словарь title_system -> список строк MTO
        debug: Флаг отладки
        debug_tag: Список тегов для отладки
        debug_code: Список кодов для отладки
        debug_title_system: Список title_system для отладки
        debug_log: Список строк debug-лога
        include_mto_vo_without_rfp_anchor: Если True, при отсутствии RFP якоря добавлять строки в конец.
    """
    if not mto_data:
        return
    
    # Собираем строки MTO без тегов с непустым IN_CABINET, которые не были полностью потреблены
    rows_to_add: List[tuple] = []  # (title_system, mto_row, remaining_qty)
    
    for title_system, mto_rows in mto_data.items():
        for mto_row in mto_rows:
            if mto_row.row_type != RowType.position_row:
                continue
            # Уже добавлены через add_unmatched_mto_rows (в т.ч. без тегов)
            if getattr(mto_row, "_added_via_unmatched_mto", False):
                continue
            # Пропускаем строки с тегами — они обрабатываются через add_unmatched_mto_rows
            if mto_row.get_tags_list():
                continue
            # Проверяем наличие IN_CABINET
            in_cabinet = str(mto_row.get_value(IN_CABINET) or "").strip()
            if not in_cabinet:
                continue
            # Пропускаем полностью потребленные по коду
            if getattr(mto_row, "_consumed_in_output", False):
                continue
            # Проверяем оставшееся количество
            total_qty = _safe_float(mto_row.get_value(VALUES))
            used_qty = _safe_float(getattr(mto_row, "_used_qty_no_tag", 0.0))
            remaining = total_qty - used_qty
            if remaining <= 0:
                continue
            
            rows_to_add.append((title_system, mto_row, remaining))
    
    if not rows_to_add:
        return
    
    if verbose_progress_messages:
        print(f"\n  Добавление {len(rows_to_add)} несопоставленных MTO строк с IN_CABINET (без тегов)...")
    
    # Находим последние индексы строк для каждого title_system
    last_indices_by_title: Dict[str, int] = {}
    for idx, rfp_row in enumerate(result_rows):
        if rfp_row.row_type == RowType.position_row:
            title = rfp_row.get_value(DS_TITLE)
            if title:
                last_indices_by_title[str(title).strip()] = idx
    
    rows_to_insert: List[tuple] = []  # (insert_idx, new_row)
    appended_without_anchor = 0

    for title_system, mto_row, remaining_qty in rows_to_add:
        title_key = str(title_system).strip() if title_system else ""
        last_idx = last_indices_by_title.get(title_key)

        if last_idx is None:
            if not include_mto_vo_without_rfp_anchor:
                if _should_debug(debug, title_key, mto_row.get_value(CODE) or "", [],
                                 debug_tag, debug_code, debug_title_system):
                    _debug_print(
                        debug_log,
                        f"  [add_unmatched_mto_cabinet] Не найдено строк с title_system '{title_key}' для вставки"
                    )
                continue
            new_row = create_rfp_row_from_mto(mto_row, None, title_key)
            new_row.el[VALUES_MTO].value = remaining_qty
            consume_mto_for_export(mto_row, remaining_qty)
            if _should_debug(debug, title_key, mto_row.get_value(CODE) or "", [],
                             debug_tag, debug_code, debug_title_system):
                _debug_print(
                    debug_log,
                    f"  [add_unmatched_mto_cabinet][append_no_anchor] title={title_key}, "
                    f"CODE={mto_row.get_value(CODE)}, IN_CABINET={mto_row.get_value(IN_CABINET)}, "
                    f"VALUES_MTO={remaining_qty}"
                )
            result_rows.append(new_row)
            appended_without_anchor += 1
            continue

        # Определяем точку вставки (после последней строки с тем же title_system)
        insert_idx = last_idx + 1
        while insert_idx < len(result_rows):
            next_row = result_rows[insert_idx]
            if next_row.row_type == RowType.position_row:
                next_title = next_row.get_value(DS_TITLE)
                if next_title and str(next_title).strip() != title_key:
                    break
            insert_idx += 1
        
        # Создаем новую строку из MTO
        new_row = create_rfp_row_from_mto(mto_row, None, title_key)
        # Устанавливаем оставшееся количество (может быть меньше исходного если часть была потреблена)
        new_row.el[VALUES_MTO].value = remaining_qty
        consume_mto_for_export(mto_row, remaining_qty)
        
        if _should_debug(debug, title_key, mto_row.get_value(CODE) or "", [],
                         debug_tag, debug_code, debug_title_system):
            _debug_print(
                debug_log,
                f"  [add_unmatched_mto_cabinet] Добавление MTO с IN_CABINET: title={title_key}, "
                f"CODE={mto_row.get_value(CODE)}, IN_CABINET={mto_row.get_value(IN_CABINET)}, "
                f"VALUES_MTO={remaining_qty}"
            )
        
        rows_to_insert.append((insert_idx, new_row))

    if not rows_to_insert and not appended_without_anchor:
        return

    # Вставляем строки (в обратном порядке чтобы индексы не сдвигались)
    insert_groups: Dict[int, List[RowStd]] = defaultdict(list)
    for insert_idx, new_row in rows_to_insert:
        insert_groups[insert_idx].append(new_row)

    added_count = 0
    for insert_idx in sorted(insert_groups.keys(), reverse=True):
        for new_row in reversed(insert_groups[insert_idx]):
            result_rows.insert(insert_idx, new_row)
            added_count += 1

    if verbose_progress_messages:
        total = added_count + appended_without_anchor
        if appended_without_anchor and added_count:
            print(f"  Добавлено MTO строк с IN_CABINET: {total} (в т.ч. {appended_without_anchor} без якоря RFP в конец)")
        else:
            print(f"  Добавлено MTO строк с IN_CABINET: {total}")
