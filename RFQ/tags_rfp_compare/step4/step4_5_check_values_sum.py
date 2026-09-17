"""
Проверка сумм VALUES для контроля целостности данных
Сравнение сумм VALUES из исходных RFP и MTO с результатом сопоставления
"""

from typing import List, Dict
from collections import defaultdict
from datetime import datetime
import os

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from base.base_classes import RowStd, RowType
from base.tables_columns import (
    VALUES, DS_TITLE, DS_NUMBER, MATCH_STATUS, MATCH_STATUS_VO, POSITION_STATUS,
    CODE_MTO, VALUES_MTO, VALUES_VO, TAG_VO, TAG_MTO, CODE, CODE_VO, TAGS
)
from RFQ.tags_rfp_compare.rfp_tags_utils import ensure_result_dir_exists


def check_values_sum(
    rfp_sums_by_title: Dict[str, float],
    rfp_counts_by_title: Dict[str, int],
    mto_tagged_sums_by_title: Dict[str, float],
    mto_tagged_values_by_code: Dict[tuple, float],
    vo_tagged_sums_by_title: Dict[str, float],
    result_rows: List[RowStd],
    result_dir: str,
):
    """
    Проверяет суммы VALUES из исходных данных и результата сопоставления.

    Args:
        rfp_sums_by_title: Суммы VALUES из RFP по title_mark (снапшот до модификации)
        rfp_counts_by_title: Количество строк RFP по title_mark
        mto_tagged_sums_by_title: Pre-computed суммы VALUES tagged MTO строк per title_mark
        mto_tagged_values_by_code: Pre-computed (title_mark, CODE) -> sum(VALUES) tagged MTO
        vo_tagged_sums_by_title: Pre-computed суммы VALUES tagged VO строк per title_mark
        result_rows: Результат сопоставления (список строк)
        result_dir: Путь к папке результатов
    """
    mto_used_without_tags = _calculate_used_mto_without_tags_from_summary(
        result_rows, mto_tagged_values_by_code,
    )
    mto_sums_by_title = _merge_sums(mto_tagged_sums_by_title, mto_used_without_tags)

    vo_used_without_tags = _calculate_used_vo_without_tags(result_rows)
    vo_sums_by_title = _merge_sums(vo_tagged_sums_by_title, vo_used_without_tags)

    result_rfp_sums_by_title = _calculate_result_rfp_sums(result_rows)
    result_mto_sums_by_title = _calculate_result_mto_sums(result_rows)
    mto_empty_values_by_title = _count_mto_empty_values(result_rows)
    result_vo_sums_by_title = _calculate_result_vo_sums(result_rows)

    mismatches = _compare_sums(
        rfp_sums_by_title,
        mto_sums_by_title,
        result_rfp_sums_by_title,
        result_mto_sums_by_title,
        vo_sums_by_title,
        result_vo_sums_by_title,
        rfp_counts_by_title
    )

    _print_summary_table(
        rfp_sums_by_title,
        mto_sums_by_title,
        result_rfp_sums_by_title,
        result_mto_sums_by_title,
        mto_empty_values_by_title,
        vo_sums_by_title,
        result_vo_sums_by_title,
        mismatches,
        rfp_counts_by_title
    )

    if mismatches:
        excel_path = _save_values_sum_check_to_excel(
            rfp_sums_by_title,
            mto_sums_by_title,
            result_rfp_sums_by_title,
            result_mto_sums_by_title,
            mto_empty_values_by_title,
            vo_sums_by_title,
            result_vo_sums_by_title,
            mismatches,
            result_dir,
            rfp_counts_by_title
        )
        if excel_path:
            print(f"Проверка сумм VALUES -> {os.path.basename(excel_path)}")
    # else:
    #     print("\nOK: Несовпадений сумм не найдено")

    return mismatches


def _save_vo_debug_rows(
    original_vo_data: Dict[str, List[RowStd]],
    result_rows: List[RowStd],
    result_dir: str,
    debug_title_mark: str
) -> None:
    normalized_title = _normalize_title_mark(debug_title_mark)
    if not normalized_title:
        return

    vo_source_rows = _collect_vo_source_rows(original_vo_data, normalized_title)
    vo_result_rows = _collect_vo_result_rows(result_rows, normalized_title)

    if not vo_source_rows and not vo_result_rows:
        return

    ensure_result_dir_exists(result_dir)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    file_name = f"Шаг4_VO_debug_rows_{normalized_title}_{timestamp}.xlsx"
    file_path = os.path.join(result_dir, file_name)

    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "VO Debug Rows"

        headers = ["Источник", "TAG", "CODE", "VALUE", "ANNOTATION"]
        ws.append(headers)

        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal="center", vertical="center")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = border

        for row in vo_source_rows:
            ws.append(["VO исходная", row["tag"], row["code"], row["value"], row["annotation"]])
        for row in vo_result_rows:
            ws.append(["VO результат", row["tag"], row["code"], row["value"], row["annotation"]])

        for col_num in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col_num)].width = 28

        last_col_letter = get_column_letter(len(headers))
        ws.auto_filter.ref = f'A1:{last_col_letter}{ws.max_row}'
        ws.freeze_panes = 'A2'

        wb.save(file_path)
        print(f"VO debug rows -> {os.path.basename(file_path)}")
    except Exception as exc:
        print(f"Ошибка сохранения VO debug rows: {exc}")


def _collect_vo_source_rows(
    original_vo_data: Dict[str, List[RowStd]],
    normalized_title: str
) -> list[dict]:
    rows: list[dict] = []
    for title_mark, vo_rows in original_vo_data.items():
        if _normalize_title_mark(title_mark) != normalized_title:
            continue
        for row in vo_rows:
            if row.row_type != RowType.position_row:
                continue
            tag_value = row.get_tags_list()
            tag = tag_value[0] if tag_value else ""
            code = row.get_value(CODE) or ""
            value = row.get_value(VALUES) or ""
            annotation = "tagged" if tag else "no_tag"
            rows.append({
                "tag": tag,
                "code": code,
                "value": value,
                "annotation": annotation
            })
    return rows


def _collect_vo_result_rows(
    result_rows: List[RowStd],
    normalized_title: str
) -> list[dict]:
    rows: list[dict] = []
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        if _normalize_title_mark(row.get_value(DS_TITLE)) != normalized_title:
            continue
        tag_vo = row.el[TAG_VO].value if row.el.get(TAG_VO) else ""
        code_vo = row.el[CODE_VO].value if row.el.get(CODE_VO) else ""
        value_vo = row.el[VALUES_VO].value if row.el.get(VALUES_VO) else ""
        status_vo = row.el[MATCH_STATUS_VO].value if row.el.get(MATCH_STATUS_VO) else ""
        if not (tag_vo or code_vo or value_vo not in (None, "")):
            continue
        parts = []
        if status_vo:
            parts.append(f"status={status_vo}")
        if not tag_vo:
            parts.append("tag_empty")
        if value_vo in (None, ""):
            parts.append("value_empty")
        annotation = "; ".join(parts)
        rows.append({
            "tag": tag_vo,
            "code": code_vo,
            "value": value_vo,
            "annotation": annotation
        })
    return rows


def _save_mto_values_debug_log(
    original_mto_data: Dict[str, List[RowStd]],
    result_rows: List[RowStd],
    result_dir: str,
    debug_title_mark: str
) -> None:
    """
    Компактный диагностический лог: откуда набегает сумма MTO VALUES.
    Выводит сводку + топ строк с наибольшим VALUES_MTO для быстрого анализа.
    """
    normalized_title = _normalize_title_mark(debug_title_mark)
    if not normalized_title:
        return

    ensure_result_dir_exists(result_dir)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    safe_title = normalized_title.replace(" ", "_").replace("/", "_")
    file_path = os.path.join(result_dir, f"step4_MTO_debug_values_{safe_title}_{timestamp}.txt")

    lines: list[str] = []
    def w(text: str = "") -> None:
        lines.append(text + "\n")

    # --- Сбор данных из MTO источника ---
    sum_tagged = 0.0
    sum_untagged = 0.0
    count_tagged = 0
    count_untagged = 0
    # CODE -> {tagged_values, untagged_values, tagged_count, untagged_count}
    mto_source_by_code: dict[str, dict] = defaultdict(lambda: {
        "tagged_val": 0.0, "untagged_val": 0.0, "tagged_cnt": 0, "untagged_cnt": 0
    })

    for title_mark_key, mto_rows in original_mto_data.items():
        if _normalize_title_mark(title_mark_key) != normalized_title:
            continue
        for row in mto_rows:
            if row.row_type != RowType.position_row:
                continue
            tags_list = row.get_tags_list()
            code_val = row.get_value(CODE) or "<пусто>"
            values_val = row.get_value(VALUES)
            try:
                vf = float(values_val) if values_val else 0.0
            except (ValueError, TypeError):
                vf = 0.0
            entry = mto_source_by_code[code_val]
            if tags_list:
                sum_tagged += vf
                count_tagged += 1
                entry["tagged_val"] += vf
                entry["tagged_cnt"] += 1
            else:
                sum_untagged += vf
                count_untagged += 1
                entry["untagged_val"] += vf
                entry["untagged_cnt"] += 1

    # --- Сбор данных из результата ---
    sum_result_total = 0.0
    sums_by_match_status: dict[str, float] = defaultdict(float)
    counts_by_match_status: dict[str, int] = defaultdict(int)
    sums_by_position_status: dict[str, float] = defaultdict(float)
    counts_by_position_status: dict[str, int] = defaultdict(int)
    # CODE_MTO -> список {values_mto, match_status, position_status, tag_mto, ds_number}
    result_by_code_mto: dict[str, list[dict]] = defaultdict(list)

    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        if _normalize_title_mark(row.get_value(DS_TITLE)) != normalized_title:
            continue

        code_mto_el = row.el.get(CODE_MTO)
        values_mto_el = row.el.get(VALUES_MTO)
        has_code_mto = bool(code_mto_el and code_mto_el.value)
        has_values_mto = bool(values_mto_el and values_mto_el.value)
        if not has_code_mto and not has_values_mto:
            continue

        code_mto = (code_mto_el.value if code_mto_el else "") or "<пусто>"
        values_mto_raw = (values_mto_el.value if values_mto_el else "")
        try:
            vmf = float(values_mto_raw) if values_mto_raw else 0.0
        except (ValueError, TypeError):
            vmf = 0.0
        match_status = row.el[MATCH_STATUS].value or ""
        position_status = row.el[POSITION_STATUS].value or ""
        tag_mto_el = row.el.get(TAG_MTO)
        tag_mto = (tag_mto_el.value if tag_mto_el else "") or ""
        ds_number = row.get_value(DS_NUMBER) or ""

        sum_result_total += vmf
        sums_by_match_status[match_status] += vmf
        counts_by_match_status[match_status] += 1
        sums_by_position_status[position_status] += vmf
        counts_by_position_status[position_status] += 1
        result_by_code_mto[code_mto].append({
            "values_mto": vmf,
            "match_status": match_status,
            "position_status": position_status,
            "tag_mto": tag_mto,
            "ds_number": ds_number,
        })

    # Считаем untagged-вклад по новой формуле: excess = result_sum_per_code - tagged_sum_per_code
    sum_used_without_tags = 0.0
    count_codes_with_untagged_excess = 0
    for code_mto, entries in result_by_code_mto.items():
        result_sum_for_code = sum(e["values_mto"] for e in entries)
        tagged_sum_for_code = mto_source_by_code.get(code_mto, {}).get("tagged_val", 0.0)
        excess = result_sum_for_code - tagged_sum_for_code
        if excess > 0.01:
            sum_used_without_tags += excess
            count_codes_with_untagged_excess += 1
    mto_original_total = sum_tagged + sum_used_without_tags
    delta = sum_result_total - mto_original_total

    # ═══════════════ ВЫВОД ═══════════════
    w(f"MTO VALUES DEBUG | title_mark='{normalized_title}' | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    w(f"{'=' * 80}")

    # --- 1. Итоговая формула ---
    w()
    w("1. ФОРМУЛА РАСХОЖДЕНИЯ")
    w(f"   MTO исходная  = {mto_original_total:>10.2f}  (с тегами: {sum_tagged:.2f} + untagged excess: {sum_used_without_tags:.2f} из {count_codes_with_untagged_excess} кодов)")
    w(f"   MTO результат = {sum_result_total:>10.2f}")
    w(f"   ДЕЛЬТА        = {delta:>+10.2f}")

    # --- 2. MTO источник: сводка по CODE ---
    w()
    w("2. MTO ИСТОЧНИК (после split) — сводка по CODE")
    w(f"   Строк с тегами: {count_tagged}, SUM={sum_tagged:.2f}  |  Без тегов: {count_untagged}, SUM={sum_untagged:.2f}")
    # Топ кодов без тегов (они — потенциальный источник набегания)
    untagged_codes = [
        (code, d["untagged_val"], d["untagged_cnt"])
        for code, d in mto_source_by_code.items()
        if d["untagged_cnt"] > 0
    ]
    untagged_codes.sort(key=lambda x: x[1], reverse=True)
    if untagged_codes:
        w(f"   Коды БЕЗ тегов (не попадают в 'MTO исходная', топ-15):")
        for code, val, cnt in untagged_codes[:15]:
            w(f"     CODE={code:<18} строк={cnt:<5} SUM(VALUES)={val:.2f}")
        if len(untagged_codes) > 15:
            w(f"     ... и ещё {len(untagged_codes) - 15} кодов")

    # --- 3. Результат: разбивка по MATCH_STATUS ---
    w()
    w("3. РЕЗУЛЬТАТ — разбивка по MATCH_STATUS")
    for status in sorted(sums_by_match_status.keys()):
        w(f"   {status:<40} строк={counts_by_match_status[status]:<6} SUM(VALUES_MTO)={sums_by_match_status[status]:.2f}")
    w(f"   {'ИТОГО':<40} строк={sum(counts_by_match_status.values()):<6} SUM={sum_result_total:.2f}")

    # --- 4. Результат: разбивка по POSITION_STATUS ---
    w()
    w("4. РЕЗУЛЬТАТ — разбивка по POSITION_STATUS")
    for status in sorted(sums_by_position_status.keys()):
        w(f"   {status:<40} строк={counts_by_position_status[status]:<6} SUM(VALUES_MTO)={sums_by_position_status[status]:.2f}")
    w(f"   {'ИТОГО':<40} строк={sum(counts_by_position_status.values()):<6} SUM={sum_result_total:.2f}")

    # --- 5. Топ CODE_MTO по набеганию VALUES_MTO ---
    w()
    w("5. ТОП CODE_MTO по суммарному VALUES_MTO (топ-20)")
    code_mto_totals = []
    for code_mto, entries in result_by_code_mto.items():
        total = sum(e["values_mto"] for e in entries)
        statuses = defaultdict(float)
        for e in entries:
            statuses[e["match_status"]] += e["values_mto"]
        code_mto_totals.append((code_mto, total, len(entries), dict(statuses)))
    code_mto_totals.sort(key=lambda x: x[1], reverse=True)
    for code_mto, total, cnt, statuses in code_mto_totals[:20]:
        status_parts = ", ".join(f"{s}={v:.0f}" for s, v in statuses.items())
        # Проверяем — был ли этот код в MTO источнике с тегами
        src = mto_source_by_code.get(code_mto)
        src_info = ""
        if src:
            if src["tagged_cnt"] > 0 and src["untagged_cnt"] > 0:
                src_info = f"  [исх: tagged={src['tagged_val']:.0f}, untagged={src['untagged_val']:.0f}]"
            elif src["tagged_cnt"] > 0:
                src_info = f"  [исх: tagged={src['tagged_val']:.0f}]"
            elif src["untagged_cnt"] > 0:
                src_info = f"  [исх: ТОЛЬКО untagged={src['untagged_val']:.0f} -> не в MTO исх!]"
        else:
            src_info = "  [исх: НЕТ в MTO источнике!]"
        w(f"   {code_mto:<18} SUM_MTO={total:>8.2f} ({cnt} строк) | {status_parts}{src_info}")
    if len(code_mto_totals) > 20:
        w(f"   ... и ещё {len(code_mto_totals) - 20} кодов")

    # --- Запись ---
    try:
        with open(file_path, "w", encoding="utf-8") as f:
            f.writelines(lines)
        print(f"\nMTO VALUES debug log сохранён в: {file_path}")
    except Exception as exc:
        print(f"Ошибка сохранения MTO VALUES debug log: {exc}")


def _normalize_title_mark(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _calculate_rfp_sums(rfp_data: List[RowStd]) -> Dict[str, float]:
    """
    Вычисляет суммы VALUES из исходных данных RFP по title_mark
    
    Args:
        rfp_data: Исходные данные RFP
        
    Returns:
        Словарь: title_mark -> сумма VALUES
    """
    sums_by_title = defaultdict(float)
    
    for row in rfp_data:
        if row.row_type != RowType.position_row:
            continue
        
        title_mark = row.get_value(DS_TITLE)
        if not title_mark:
            continue
        
        values = row.get_value(VALUES)
        try:
            values_float = float(values) if values else 0.0
        except (ValueError, TypeError):
            values_float = 0.0
        
        sums_by_title[title_mark] += values_float
    
    return dict(sums_by_title)


def _calculate_rfp_row_counts(rfp_data: List[RowStd]) -> Dict[str, int]:
    counts_by_title = defaultdict(int)
    for row in rfp_data:
        if row.row_type != RowType.position_row:
            continue
        title_mark = row.get_value(DS_TITLE)
        if not title_mark:
            continue
        counts_by_title[title_mark] += 1
    return dict(counts_by_title)


def _calculate_mto_tagged_sums(mto_data: Dict[str, List[RowStd]]) -> Dict[str, float]:
    """
    Вычисляет суммы VALUES из исходных данных MTO по title_mark
    Учитываются строки с тегами и строки без тегов, которые были использованы
    для строк со статусом "Не протегирован в МТО"
    
    Args:
        mto_data: Исходные данные MTO (словарь: title_mark -> список строк)
        
    Returns:
        Словарь: title_mark -> сумма VALUES (строки с тегами + использованные строки без тегов)
    """
    sums_by_title = defaultdict(float)
    
    for title_mark, mto_rows in mto_data.items():
        for row in mto_rows:
            if row.row_type != RowType.position_row:
                continue
            if not row.get_tags_list():
                continue
            values = row.get_value(VALUES)
            try:
                values_float = float(values) if values else 0.0
            except (ValueError, TypeError):
                values_float = 0.0
            
            sums_by_title[title_mark] += values_float
    
    return dict(sums_by_title)


def _calculate_used_mto_without_tags(
    result_rows: List[RowStd],
    original_mto_data: Dict[str, List[RowStd]]
) -> Dict[str, float]:
    """
    Считает VALUES_MTO из результата, пришедшие от untagged MTO-строк.

    Логика:
    1. Для каждого (title_mark, CODE) считаем SUM(VALUES) tagged строк из MTO-источника.
    2. Для каждого (title_mark, CODE_MTO) считаем SUM(VALUES_MTO) из результата.
    3. Если SUM(результат) > SUM(tagged источник) — разница пришла от untagged строк
       и не учтена в _calculate_mto_tagged_sums.
    """
    # (title_mark, CODE) -> суммарный VALUES из строк С тегами в MTO-источнике
    tagged_values_by_code: dict[tuple[str, str], float] = defaultdict(float)
    for title_mark, mto_rows in original_mto_data.items():
        for row in mto_rows:
            if row.row_type != RowType.position_row:
                continue
            if not row.get_tags_list():
                continue
            code = row.get_value(CODE) or ""
            if code:
                values_val = row.get_value(VALUES)
                try:
                    vf = float(values_val) if values_val else 0.0
                except (ValueError, TypeError):
                    vf = 0.0
                tagged_values_by_code[(title_mark, code)] += vf

    # (title_mark, CODE_MTO) -> суммарный VALUES_MTO из результата
    result_values_by_code: dict[tuple[str, str], float] = defaultdict(float)
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        title_mark = row.get_value(DS_TITLE)
        if not title_mark:
            continue

        code_mto_el = row.el.get(CODE_MTO)
        values_mto_el = row.el.get(VALUES_MTO)
        code_mto = (code_mto_el.value if code_mto_el else "") or ""
        has_mto_data = bool(code_mto) or bool(values_mto_el and values_mto_el.value)
        if not has_mto_data:
            continue

        if values_mto_el and values_mto_el.value:
            try:
                vmf = float(values_mto_el.value)
            except (ValueError, TypeError):
                vmf = 0.0
        else:
            vmf = 0.0

        result_values_by_code[(title_mark, code_mto)] += vmf

    # Для каждого кода: если результат > tagged источник → разница от untagged
    sums_by_title = defaultdict(float)
    for key, result_sum in result_values_by_code.items():
        title_mark = key[0]
        tagged_sum = tagged_values_by_code.get(key, 0.0)
        excess = result_sum - tagged_sum
        if excess > 0.01:
            sums_by_title[title_mark] += excess
    return dict(sums_by_title)


def _calculate_used_mto_without_tags_from_summary(
    result_rows: List[RowStd],
    tagged_values_by_code: Dict[tuple, float],
) -> Dict[str, float]:
    """Аналог _calculate_used_mto_without_tags, но использует pre-computed summary.

    tagged_values_by_code: {(title_mark, CODE) -> sum(VALUES)} для tagged MTO строк,
    вычислен в worker'е и агрегирован в main-процессе.
    """
    result_values_by_code: dict[tuple[str, str], float] = defaultdict(float)
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        title_mark = row.get_value(DS_TITLE)
        if not title_mark:
            continue
        code_mto_el = row.el.get(CODE_MTO)
        values_mto_el = row.el.get(VALUES_MTO)
        code_mto = (code_mto_el.value if code_mto_el else "") or ""
        has_mto_data = bool(code_mto) or bool(values_mto_el and values_mto_el.value)
        if not has_mto_data:
            continue
        if values_mto_el and values_mto_el.value:
            try:
                vmf = float(values_mto_el.value)
            except (ValueError, TypeError):
                vmf = 0.0
        else:
            vmf = 0.0
        result_values_by_code[(title_mark, code_mto)] += vmf

    sums_by_title = defaultdict(float)
    for key, result_sum in result_values_by_code.items():
        title_mark = key[0]
        tagged_sum = tagged_values_by_code.get(key, 0.0)
        excess = result_sum - tagged_sum
        if excess > 0.01:
            sums_by_title[title_mark] += excess
    return dict(sums_by_title)


def _merge_sums(base_sums: Dict[str, float], add_sums: Dict[str, float]) -> Dict[str, float]:
    merged = defaultdict(float)
    for title_mark, value in base_sums.items():
        merged[title_mark] += value
    for title_mark, value in add_sums.items():
        merged[title_mark] += value
    return dict(merged)


def _calculate_result_rfp_sums(result_rows: List[RowStd]) -> Dict[str, float]:
    """
    Вычисляет суммы VALUES из левой части (RFP) результата сопоставления по title_mark
    
    Args:
        result_rows: Результат сопоставления
        
    Returns:
        Словарь: title_mark -> сумма VALUES
    """
    sums_by_title = defaultdict(float)
    
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        
        title_mark = row.get_value(DS_TITLE)
        if not title_mark:
            continue
        
        # Учитываем только строки, которые были в исходном RFP (не добавленные из MTO)
        status = row.el[MATCH_STATUS].value
        status_vo = row.el[MATCH_STATUS_VO].value
        if status == "Добавлен из МТО" or status_vo == "Добавлен из VO":
            continue  # Пропускаем строки, добавленные из MTO
        
        values = row.get_value(VALUES)
        try:
            values_float = float(values) if values else 0.0
        except (ValueError, TypeError):
            values_float = 0.0
        
        sums_by_title[title_mark] += values_float
    
    return dict(sums_by_title)


def _calculate_result_mto_sums(result_rows: List[RowStd]) -> Dict[str, float]:
    """
    Вычисляет суммы VALUES из правой части (MTO) результата сопоставления по title_mark
    Учитывает строки с тегами и строки без тегов (со статусом "Не протегирован в МТО")
    
    Args:
        result_rows: Результат сопоставления
        
    Returns:
        Словарь: title_mark -> сумма VALUES_MTO (из MTO данных, включая позиции без тегов)
    """
    sums_by_title = defaultdict(float)
    
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        
        title_mark = row.get_value(DS_TITLE)
        if not title_mark:
            continue
        
        # Учитываем строки, где есть данные MTO
        # Это включает:
        # 1. Строки с сопоставленными тегами (MATCH_STATUS = "Тег сопоставлен")
        # 2. Строки, добавленные из MTO (MATCH_STATUS = "Добавлен из МТО")
        # 3. Строки со статусом "Не протегирован в МТО" (использованы позиции без тегов)
        status = row.el[MATCH_STATUS].value
        position_status = row.el[POSITION_STATUS].value
        
        # Проверяем наличие данных MTO через CODE_MTO или VALUES_MTO
        code_mto_value = row.el.get(CODE_MTO)
        values_mto_value = row.el.get(VALUES_MTO)
        
        has_mto_data = False
        if code_mto_value and code_mto_value.value:
            has_mto_data = True
        elif values_mto_value and values_mto_value.value:
            has_mto_data = True
        
        if has_mto_data:
            # Используем VALUES_MTO, если оно есть, иначе считаем как 0.0
            if values_mto_value and values_mto_value.value:
                try:
                    values_float = float(values_mto_value.value)
                except (ValueError, TypeError):
                    values_float = 0.0
            else:
                values_float = 0.0
            
            sums_by_title[title_mark] += values_float
    
    return dict(sums_by_title)


def _count_mto_empty_values(result_rows: List[RowStd]) -> Dict[str, int]:
    counts_by_title = defaultdict(int)
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue

        title_mark = row.get_value(DS_TITLE)
        if not title_mark:
            continue

        code_mto_value = row.el.get(CODE_MTO)
        values_mto_value = row.el.get(VALUES_MTO)

        has_mto_data = False
        if code_mto_value and code_mto_value.value:
            has_mto_data = True
        elif values_mto_value and values_mto_value.value:
            has_mto_data = True

        if not has_mto_data:
            continue

        if not values_mto_value or not values_mto_value.value:
            counts_by_title[title_mark] += 1

    return dict(counts_by_title)


def _calculate_vo_tagged_sums(vo_data: Dict[str, List[RowStd]]) -> Dict[str, float]:
    sums_by_title = defaultdict(float)
    for title_mark, vo_rows in vo_data.items():
        for row in vo_rows:
            if row.row_type != RowType.position_row:
                continue
            if not row.get_tags_list():
                continue
            values = row.get_value(VALUES)
            try:
                values_float = float(values) if values else 0.0
            except (ValueError, TypeError):
                values_float = 0.0
            sums_by_title[title_mark] += values_float
    return dict(sums_by_title)


def _calculate_used_vo_without_tags(result_rows: List[RowStd]) -> Dict[str, float]:
    sums_by_title = defaultdict(float)
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        title_mark = row.get_value(DS_TITLE)
        if not title_mark:
            continue
        status_vo = row.el[MATCH_STATUS_VO].value
        if status_vo not in {"Тег сопоставлен", "Тег в VO заменен", "Добавлен из VO"}:
            continue

        tag_vo_value = row.el.get(TAG_VO)
        tag_vo = tag_vo_value.value if tag_vo_value else ""
        if tag_vo:
            continue

        values_vo_value = row.el.get(VALUES_VO)
        if values_vo_value and values_vo_value.value:
            try:
                values_float = float(values_vo_value.value)
            except (ValueError, TypeError):
                values_float = 0.0
        else:
            values_float = 0.0

        sums_by_title[title_mark] += values_float
    return dict(sums_by_title)


def _calculate_result_vo_sums(result_rows: List[RowStd]) -> Dict[str, float]:
    sums_by_title = defaultdict(float)
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        title_mark = row.get_value(DS_TITLE)
        if not title_mark:
            continue
        values_vo_value = row.el.get(VALUES_VO)
        if values_vo_value and values_vo_value.value:
            try:
                values_float = float(values_vo_value.value)
            except (ValueError, TypeError):
                values_float = 0.0
            sums_by_title[title_mark] += values_float
    return dict(sums_by_title)


def _compare_sums(
    rfp_sums: Dict[str, float],
    mto_sums: Dict[str, float],
    result_rfp_sums: Dict[str, float],
    result_mto_sums: Dict[str, float],
    vo_sums: Dict[str, float],
    result_vo_sums: Dict[str, float],
    rfp_row_counts: Dict[str, int]
) -> List[dict]:
    """
    Сравнивает суммы и находит несовпадения
    
    Args:
        rfp_sums: Суммы VALUES из исходного RFP
        mto_sums: Суммы VALUES из исходного MTO
        result_rfp_sums: Суммы VALUES из результата (левая часть)
        result_mto_sums: Суммы VALUES из результата (правая часть)
        
    Returns:
        Список словарей с информацией о несовпадениях
    """
    mismatches = []
    all_titles = (
        set(rfp_sums.keys())
        | set(mto_sums.keys())
        | set(result_rfp_sums.keys())
        | set(result_mto_sums.keys())
        | set(vo_sums.keys())
        | set(result_vo_sums.keys())
    )
    
    for title_mark in sorted(all_titles, key=lambda x: (x is None, x or "")):
        if rfp_row_counts.get(title_mark, 0) == 0:
            continue
        rfp_sum = rfp_sums.get(title_mark, 0.0)
        mto_sum = mto_sums.get(title_mark, 0.0)
        result_rfp_sum = result_rfp_sums.get(title_mark, 0.0)
        result_mto_sum = result_mto_sums.get(title_mark, 0.0)
        vo_sum = vo_sums.get(title_mark, 0.0)
        result_vo_sum = result_vo_sums.get(title_mark, 0.0)
        
        has_mismatch = False
        mismatch_reasons = []
        
        # Проверка 1: Сумма RFP в результате должна совпадать с исходной
        if abs(rfp_sum - result_rfp_sum) > 0.01:  # Учитываем погрешность округления
            has_mismatch = True
            mismatch_reasons.append(f"RFP: исходная={rfp_sum}, результат={result_rfp_sum}")
        
        # Проверка 2: Сумма MTO в результате должна совпадать с исходной
        if abs(mto_sum - result_mto_sum) > 0.01:
            has_mismatch = True
            mismatch_reasons.append(f"MTO: исходная={mto_sum}, результат={result_mto_sum}")
        
        # Проверка 3: Сумма VO в результате должна совпадать с исходной
        if abs(vo_sum - result_vo_sum) > 0.01:
            has_mismatch = True
            mismatch_reasons.append(f"VO: исходная={vo_sum}, результат={result_vo_sum}")
        
        if has_mismatch:
            mismatches.append({
                'title_mark': title_mark,
                'rfp_sum_original': rfp_sum,
                'rfp_sum_result': result_rfp_sum,
                'mto_sum_original': mto_sum,
                'mto_sum_result': result_mto_sum,
                'vo_sum_original': vo_sum,
                'vo_sum_result': result_vo_sum,
                'mismatch_reasons': '; '.join(mismatch_reasons)
            })
    
    return mismatches


def _print_summary_table(
    rfp_sums: Dict[str, float],
    mto_sums: Dict[str, float],
    result_rfp_sums: Dict[str, float],
    result_mto_sums: Dict[str, float],
    mto_empty_values: Dict[str, int],
    vo_sums: Dict[str, float],
    result_vo_sums: Dict[str, float],
    mismatches: List[dict],
    rfp_row_counts: Dict[str, int]
):
    """
    Выводит таблицу с суммами в консоль
    """
    from prettytable import PrettyTable
    
    table = PrettyTable()
    table.field_names = [
        "title_mark",
        "RFP исходная",
        "RFP результат",
        "MTO исходная",
        "MTO результат",
        "MTO пустые VALUES",
        "VO исходная",
        "VO результат",
        "Статус"
    ]
    table.border = True
    table.align = "l"
    table.align["RFP исходная"] = "r"
    table.align["RFP результат"] = "r"
    table.align["MTO исходная"] = "r"
    table.align["MTO результат"] = "r"
    
    all_titles = (
        set(rfp_sums.keys())
        | set(mto_sums.keys())
        | set(result_rfp_sums.keys())
        | set(result_mto_sums.keys())
        | set(vo_sums.keys())
        | set(result_vo_sums.keys())
    )
    
    for title_mark in sorted(all_titles, key=lambda x: (x is None, x or "")):
        if rfp_row_counts.get(title_mark, 0) == 0:
            continue
        rfp_sum = rfp_sums.get(title_mark, 0.0)
        mto_sum = mto_sums.get(title_mark, 0.0)
        result_rfp_sum = result_rfp_sums.get(title_mark, 0.0)
        result_mto_sum = result_mto_sums.get(title_mark, 0.0)
        mto_empty_count = mto_empty_values.get(title_mark, 0)
        vo_sum = vo_sums.get(title_mark, 0.0)
        result_vo_sum = result_vo_sums.get(title_mark, 0.0)
        
        # Проверяем совпадение
        rfp_match = abs(rfp_sum - result_rfp_sum) <= 0.01
        mto_match = abs(mto_sum - result_mto_sum) <= 0.01
        
        vo_match = abs(vo_sum - result_vo_sum) <= 0.01
        if rfp_match and mto_match and vo_match:
            status = "OK"
        else:
            status = "ERROR"

        if status == "ERROR":
            table.add_row([
                title_mark,
                f"{rfp_sum:.2f}",
                f"{result_rfp_sum:.2f}",
                f"{mto_sum:.2f}",
                f"{result_mto_sum:.2f}",
                f"{mto_empty_count}" if mto_empty_count else "",
                f"{vo_sum:.2f}",
                f"{result_vo_sum:.2f}",
                status
            ])
    
    if table.rowcount > 0:
        print(table)
    
    if mismatches:
        print(f"Проверка сумм VALUES: WARN ({len(mismatches)} несовпадений)")
    else:
        print("Проверка сумм VALUES: OK")


def _save_values_sum_check_to_excel(
    rfp_sums: Dict[str, float],
    mto_sums: Dict[str, float],
    result_rfp_sums: Dict[str, float],
    result_mto_sums: Dict[str, float],
    mto_empty_values: Dict[str, int],
    vo_sums: Dict[str, float],
    result_vo_sums: Dict[str, float],
    mismatches: List[dict],
    result_dir: str,
    rfp_row_counts: Dict[str, int]
) -> str:
    """
    Сохраняет результаты проверки сумм в Excel файл
    
    Args:
        rfp_sums: Суммы VALUES из исходного RFP
        mto_sums: Суммы VALUES из исходного MTO
        result_rfp_sums: Суммы VALUES из результата (левая часть)
        result_mto_sums: Суммы VALUES из результата (правая часть)
        mismatches: Список несовпадений
        result_dir: Путь к папке результатов
        
    Returns:
        Путь к сохраненному Excel файлу или None в случае ошибки
    """
    try:
        ensure_result_dir_exists(result_dir)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Проверка сумм VALUES"
        
        # Заголовки
        headers = [
            "№",
            "title_mark",
            "RFP исходная",
            "RFP результат",
            "Разница RFP",
            "MTO исходная",
            "MTO результат",
            "Разница MTO",
            "MTO пустые VALUES",
            "VO исходная",
            "VO результат",
            "Разница VO",
            "Статус",
            "Причина несовпадения"
        ]
        ws.append(headers)
        
        # Стили для заголовков
        header_fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal="center", vertical="center")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = border
        
        # Заполняем данные
        all_titles = (
            set(rfp_sums.keys())
            | set(mto_sums.keys())
            | set(result_rfp_sums.keys())
            | set(result_mto_sums.keys())
            | set(vo_sums.keys())
            | set(result_vo_sums.keys())
        )
        
        row_idx = 0
        for title_mark in sorted(all_titles, key=lambda x: (x is None, x or "")):
            if rfp_row_counts.get(title_mark, 0) == 0:
                continue
            row_idx += 1
            rfp_sum = rfp_sums.get(title_mark, 0.0)
            mto_sum = mto_sums.get(title_mark, 0.0)
            result_rfp_sum = result_rfp_sums.get(title_mark, 0.0)
            result_mto_sum = result_mto_sums.get(title_mark, 0.0)
            mto_empty_count = mto_empty_values.get(title_mark, 0)
            vo_sum = vo_sums.get(title_mark, 0.0)
            result_vo_sum = result_vo_sums.get(title_mark, 0.0)
            
            rfp_diff = result_rfp_sum - rfp_sum
            mto_diff = result_mto_sum - mto_sum
            vo_diff = result_vo_sum - vo_sum
            
            # Проверяем совпадение
            rfp_match = abs(rfp_diff) <= 0.01
            mto_match = abs(mto_diff) <= 0.01
            
            vo_match = abs(vo_diff) <= 0.01
            if rfp_match and mto_match and vo_match:
                status = "OK"
                fill_color = "C6EFCE"  # Зеленый
            else:
                status = "ОШИБКА"
                fill_color = "FFC7CE"  # Красный
            
            # Находим причину несовпадения
            mismatch_reason = ""
            for mismatch in mismatches:
                if mismatch['title_mark'] == title_mark:
                    mismatch_reason = mismatch['mismatch_reasons']
                    break
            
            row_data = [
                row_idx,
                title_mark,
                rfp_sum,
                result_rfp_sum,
                rfp_diff,
                mto_sum,
                result_mto_sum,
                mto_diff,
                mto_empty_count if mto_empty_count else "",
                vo_sum,
                result_vo_sum,
                vo_diff,
                status,
                mismatch_reason
            ]
            
            ws.append(row_data)
            
            # Применяем стили к строке
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=ws.max_row, column=col_num)
                cell.border = border
                
                if col_num == 13:  # Столбец "Статус"
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                
                if col_num in [3, 4, 5, 6, 7, 8, 9, 10, 11, 12]:  # Числовые столбцы
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
        
        # Настраиваем ширину колонок
        ws.column_dimensions['A'].width = 8   # №
        ws.column_dimensions['B'].width = 20  # title_mark
        ws.column_dimensions['C'].width = 15  # RFP исходная
        ws.column_dimensions['D'].width = 15  # RFP результат
        ws.column_dimensions['E'].width = 15  # Разница RFP
        ws.column_dimensions['F'].width = 15  # MTO исходная
        ws.column_dimensions['G'].width = 15  # MTO результат
        ws.column_dimensions['H'].width = 15  # Разница MTO
        ws.column_dimensions['I'].width = 15  # MTO пустые VALUES
        ws.column_dimensions['J'].width = 15  # VO исходная
        ws.column_dimensions['K'].width = 15  # VO результат
        ws.column_dimensions['L'].width = 15  # Разница VO
        ws.column_dimensions['M'].width = 12  # Статус
        ws.column_dimensions['N'].width = 50  # Причина несовпадения
        
        # Добавляем автофильтр
        last_col_letter = get_column_letter(len(headers))
        ws.auto_filter.ref = f'A1:{last_col_letter}{ws.max_row}'
        
        # Замораживаем первую строку
        ws.freeze_panes = 'A2'
        
        # Сохраняем файл
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_name = f"Шаг4_Проверка_сумм_VALUES_{timestamp}.xlsx"
        file_path = os.path.join(result_dir, file_name)
        
        wb.save(file_path)
        return file_path
        
    except Exception as e:
        print(f"Ошибка при сохранении проверки сумм в Excel: {e}")
        import traceback
        traceback.print_exc()
        return None
