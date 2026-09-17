"""
Система раскраски ячеек данных для итогового Excel-файла (Шаг 4).

Присваивает цвета в row.el[column].color на основе статусов сопоставления
MTO и VO.  При записи в Excel цвет считывается из свойства ячейки.

Структура:
  Группы столбцов   — множества имён столбцов (RFP / MTO / VO / статусы).
  assign_row_colors  — точка входа: проставляет row.el[col].color для всех строк.
  _assign_colors_for_row — определяет цвет каждого столбца одной строки.
  _assign_mto_old_replacement_code_colors — CODE_MTO фиолетовый для новых строк
      со «старым» кодом из таблицы замен (СТАРЫЙ_КОД).

Для добавления новых правил раскраски:
  1. Добавьте цвет в Color (utils/colors.py) при необходимости.
  2. Добавьте логику в _assign_colors_for_row (новый блок «Правило N»)
     или отдельный проход в assign_row_colors (как _assign_mto_old_replacement_code_colors).
"""

from typing import Dict, List, Set, Tuple

from utils.colors import Color
from base.base_classes import RowStd, RowType
from base.tables_columns import *


# ── Группы столбцов (по имени, не по позиции в Excel) ────────────────────────

_RFP_COLUMNS: Set[str] = {
    DS_NAME, DS_ACTUAL, DS_LOT, DS_NUMBER, DS_TITLE, TAGS, CODE, NAME, TYPE_MARK, VALUES,
}

_MTO_COLUMNS: Set[str] = {
    TAG_MTO, CODE_MTO, NAME_MTO, TYPE_MARK_MTO, VALUES_MTO, NUMBERS_MTO, MTO_CODE_STRUCK,
}

_VO_COLUMNS: Set[str] = {
    TAG_VO, CODE_VO, NAME_VO, VALUES_VO,
}

# Правило 1: VO найден + MTO не найден — разнополярная окраска.
# Столбцы, получающие цвет «не найден» (MTO-блок + статус MTO):
_RULE1_NOT_FOUND: Set[str] = _MTO_COLUMNS | {MATCH_STATUS}
# Столбцы, получающие цвет VO (RFP + VO + отдельные статусы):
_RULE1_VO_FILL: Set[str] = _RFP_COLUMNS | _VO_COLUMNS | {MATCH_STATUS_VO, MATCH_STATUS_TAGS}


# ── Основная функция ─────────────────────────────────────────────────────────


def assign_row_colors(
    result_rows: List[RowStd],
    replacement_table: Dict[str, List[Tuple[str, str]]] = None,
    assign_rfp_mto_code_compare_colors: bool = True,
) -> None:
    """
    Присваивает цвет в row.el[col].color для всех position_row.

    Вызывается ПЕРЕД save_match_result_to_excel.
    При записи в Excel цвет считывается из свойства ячейки.

    Args:
        result_rows: Список строк результата.
        replacement_table: Таблица замен кодов {старый_код -> [(новый_код, статус), ...]}.
        assign_rfp_mto_code_compare_colors: Включить раскраску CODE RFP vs CODE MTO.
    """
    repl = replacement_table or {}
    old_replacement_codes = _build_old_replacement_codes(repl)
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        _assign_colors_for_row_old(row)
        _assign_colors_for_row(row, repl)
        _assign_mto_vo_compare_colors(row, repl)
        if assign_rfp_mto_code_compare_colors:
            _assign_rfp_mto_code_compare_colors(row)
        _assign_mto_old_replacement_code_colors(row, old_replacement_codes)


# ── Логика раскраски одной строки ─────────────────────────────────────────────
def _assign_colors_for_row(row: RowStd, replacement_table: Dict[str, List[Tuple[str, str]]],) -> None:
    """
    Определяет и присваивает цвет каждому столбцу строки.
    Правила проверяются сверху вниз; первое подходящее завершает обработку.    """
   

    rfp_code = row.get_value(CODE)
    mto_code = row.get_value(CODE_MTO)    

    if rfp_code and mto_code and rfp_code==mto_code:
        _set_all_columns_color(row, Color.match_matched)  

def _assign_colors_for_row_old(row: RowStd) -> None:
    """
    Определяет и присваивает цвет каждому столбцу строки.

    Правила проверяются сверху вниз; первое подходящее завершает обработку.
    """
    status_mto = row.el[MATCH_STATUS].value or ""
    status_vo = row.el[MATCH_STATUS_VO].value or ""

    is_mto_matched = status_mto in ("Тег сопоставлен", "Добавлен из МТО")
    is_mto_replaced = status_mto == "Тег в МТО заменен"
    is_mto_not_found = status_mto == "Не найден в МТО"
    is_vo_matched = status_vo in ("Тег сопоставлен", "Тег в VO заменен")
    is_vo_replaced = status_vo == "Тег в VO заменен"
    is_vo_added = status_vo == "Добавлен из VO"

    # ── Правило 1: VO найден, MTO не найден — разнополярная окраска ──────
    #   RFP + VO столбцы → цвет VO,  MTO столбцы → «не найден».
    if is_vo_matched and is_mto_not_found:
        vo_color = Color.match_replaced_vo if is_vo_replaced else Color.match_matched
        for col_name in ColNames.column_list:
            if col_name in _RULE1_NOT_FOUND:
                row.el[col_name].color = Color.match_not_found
            elif col_name in _RULE1_VO_FILL:
                row.el[col_name].color = vo_color
        return

    # ── Правило 2: VO заменен, MTO не сопоставлен и не заменен ───────────
    if is_vo_replaced and not is_mto_matched and not is_mto_replaced:
        _set_all_columns_color(row, Color.match_replaced_vo)
        return

    # ── Правило 3: единый цвет строки по статусу MTO / VO ───────────────
    row_color = _get_row_color_by_status(status_mto, is_vo_added)
    if row_color != Color.no:
        _set_all_columns_color(row, row_color)


# ── Сравнение MTO и VO частей строки ─────────────────────────────────────────

# Столбцы MTO-части (без MTO_CODE_STRUCK — он не участвует в сравнении)
_MTO_PART: Set[str] = {TAG_MTO, CODE_MTO, NAME_MTO, TYPE_MARK_MTO, VALUES_MTO, NUMBERS_MTO}
_VO_PART: Set[str] = {TAG_VO, CODE_VO, NAME_VO, VALUES_VO}


def _assign_mto_vo_compare_colors(
    row: RowStd,
    replacement_table: Dict[str, List[Tuple[str, str]]],) -> None:
    """
    Раскрашивает MTO-часть и VO-часть строки по результату сравнения
    кодов и тегов MTO vs VO.

    Вызывается ПОСЛЕ _assign_colors_for_row — может перекрасить
    отдельные столбцы поверх базовой раскраски.

    Args:
        row: Строка данных.
        replacement_table: Таблица замен {старый_код -> [(новый_код, статус), ...]}.
    """
    tag_mto = str(row.el[TAG_MTO].value or "").strip()
    tag_vo = str(row.el[TAG_VO].value or "").strip()
    code_mto = str(row.el[CODE_MTO].value or "").strip()
    code_vo = str(row.el[CODE_VO].value or "").strip()

    # ── 1. Коды и теги совпадают — MTO и VO равны ────────────────────────
    if code_mto and code_vo and code_mto == code_vo and tag_mto == tag_vo:
        _set_columns_color(row, _MTO_PART, Color.match_matched)
        _set_columns_color(row, _VO_PART, Color.match_matched)
        return

    # ── 2. Теги совпадают, коды различаются (оба не пустые) ──────────────
    if tag_mto == tag_vo and code_mto and code_vo and code_mto != code_vo:
        _set_columns_color(row, _MTO_PART, Color.match_matched)
        _set_columns_color(row, _VO_PART, Color.match_matched)
        # Прямая замена по таблице: старый код в MTO → новый код в VO
        if _is_replacement(code_mto, code_vo, replacement_table):
            row.el[CODE_MTO].color = Color.is_replacement_mto_vo  # (светло зеленый)
            row.el[CODE_VO].color = Color.is_replacement_mto_vo
        else:
            row.el[CODE_MTO].color = Color.match_replaced_mto
            row.el[CODE_VO].color = Color.match_replaced_mto
        return

    # ── 3. Коды совпадают, теги различаются (оба не пустые) ──────────────
    if code_mto and code_vo and code_mto == code_vo and tag_mto != tag_vo:
        _set_columns_color(row, _MTO_PART, Color.match_matched)
        _set_columns_color(row, _VO_PART, Color.match_matched)
        row.el[TAG_MTO].color = Color.match_replaced_mto
        row.el[TAG_VO].color = Color.match_replaced_mto
        return

    # ── 4. Код VO есть, код MTO пустой — VO добавлен ────────────────────
    if code_vo and not code_mto:
        _set_columns_color(row, _VO_PART, Color.match_added_vo)
        return

    # ── 5. Код MTO есть, код VO пустой, шкафное оборудование — нет в VO ──
    #   IN_CABINET не пуст → оборудование относится к шкафу и должно быть в VO.
    if code_mto and not code_vo:
        in_cabinet = str(row.el[IN_CABINET].value or "").strip()
        if in_cabinet:
            _set_columns_color(row, _MTO_PART, Color.match_added_mto)
            _set_columns_color(row, _VO_PART, Color.match_added_mto)
            return


def _assign_rfp_mto_code_compare_colors(row: RowStd) -> None:
    """
    Сравнивает CODE (RFP) и CODE_MTO. Раскрашивает столбцы по результату.

    Вызывается ПОСЛЕ _assign_mto_vo_compare_colors — перекрашивает CODE и CODE_MTO.

    Правила:
      - Коды равны (оба не пустые): CODE → match_matched, CODE_MTO не меняем.
      - Коды различаются (оба не пустые): CODE и CODE_MTO → unmatch_rfp_mto.
    """
    code_rfp = str(row.el[CODE].value or "").strip()
    code_mto = str(row.el[CODE_MTO].value or "").strip()

    if not code_rfp or not code_mto:
        return

    if code_rfp == code_mto:
        row.el[CODE].color = Color.match_matched
    else:
        row.el[CODE].color = Color.unmatch_rfp_mto
        row.el[CODE_MTO].color = Color.unmatch_rfp_mto


def _assign_mto_old_replacement_code_colors(
    row: RowStd,
    old_replacement_codes: Set[str],
) -> None:
    """
    Подсвечивает CODE_MTO фиолетовым для строк, новых относительно ДС,
    если код MTO входит в «старые» коды таблицы замен (СТАРЫЙ_КОД).

    Вызывается ПОСЛЕ всех остальных проходов — перекрашивает только CODE_MTO.
    """
    if not old_replacement_codes or not _is_new_vs_ds(row):
        return
    code_mto = _normalize_code(row.el[CODE_MTO].value)
    if code_mto and code_mto in old_replacement_codes:
        row.el[CODE_MTO].color = Color.violet


def _set_columns_color(row: RowStd, columns: Set[str], color: str) -> None:
    """Устанавливает цвет для заданного набора столбцов."""
    for col_name in columns:
        row.el[col_name].color = color


def _is_replacement(
    code_mto: str,
    code_vo: str,
    replacement_table: Dict[str, List[Tuple[str, str]]],
) -> bool:
    """
    Проверяет, является ли пара (code_mto → code_vo) прямой заменой
    по таблице замен.  Старый код в MTO, новый код в VO.
    """
    if not replacement_table:
        return False
    replacements = replacement_table.get(code_mto)
    if not replacements:
        return False
    return any(new_code == code_vo for new_code, _status in replacements)


# ── Вспомогательные функции ──────────────────────────────────────────────────


def _normalize_code(code_value) -> str:
    if code_value is None:
        return ""
    return str(code_value).strip().upper()


def _build_old_replacement_codes(
    replacement_table: Dict[str, List[Tuple[str, str]]],
) -> Set[str]:
    """Множество нормализованных ключей replacement_table (СТАРЫЙ_КОД)."""
    if not replacement_table:
        return set()
    return {_normalize_code(old_code) for old_code in replacement_table if old_code}


def _is_new_vs_ds(row: RowStd) -> bool:
    """Строка новая относительно ДС: нет № позиции или добавлена из MTO."""
    ds_number = str(row.get_value(DS_NUMBER) or "").strip()
    status_mto = str(row.el[MATCH_STATUS].value or "").strip()
    return not ds_number or status_mto == "Добавлен из МТО"


def _set_all_columns_color(row: RowStd, color: str) -> None:
    """Устанавливает один цвет для всех столбцов строки."""
    for col_name in ColNames.column_list:
        row.el[col_name].color = color


def _get_row_color_by_status(status_mto: str, is_vo_added: bool) -> str:
    """Определяет единый цвет строки по статусу MTO (приоритет) или VO."""
    if status_mto == "Тег сопоставлен":
        return Color.match_matched
    if status_mto == "Не найден в МТО":
        return Color.match_not_found
    if status_mto == "Добавлен из МТО":
        return Color.match_added_mto
    if status_mto == "Тег в МТО заменен":
        return Color.match_replaced_mto
    if is_vo_added:
        return Color.match_added_vo
    return Color.no
