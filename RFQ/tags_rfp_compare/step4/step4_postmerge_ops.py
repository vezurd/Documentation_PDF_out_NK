"""
Per-row post-merge operations for step4.

These operations read/write within a single RowStd — no cross-TM dependencies.
Safe to run per-TM inside worker processes.
"""

from __future__ import annotations

import math
import re
from typing import Dict, List, Tuple

from base.base_classes import CheckElement, RowStd, RowType
from base.tables_columns import (
    CODE, CODE_MTO, CODE_VO, DS_LOT, DS_NAME,
    EQUIPMENT_TYPE_STATUS, MATCH_STATUS, MATCH_STATUS_RFP_MTO_STRUCK_CODE,
    MATCH_STATUS_RFP_MTO_TAGS, MATCH_STATUS_TAGS, MATCH_STATUS_VO,
    MATCH_STATUS_VO_MTO_TAGS, MTO_CABINET_EQUIPMENT, MTO_CODE_STRUCK,
    NAME, POSITION_STATUS, TAG_EFFECTIVE, TAG_MTO, TAG_VO,
    TYPE_MARK, UL_UNITS, UL_VALUES, UNITS, UNITS_CHECK_STATUS,
    UNITS_CONVERSION_TRACE, UNITS_MTO, VALUES, VALUES_MTO,
    VALUES_MTO_RFP_DIFF, VALUES_MTO_UL_DIFF, VENDOR,
)
from RFQ.ds_compare.ds_units_normalize import normalize_units_text
from RFQ.tags_rfp_compare.rfp_supply_status import is_excluded_from_supply
from RFQ.tags_rfp_compare.step4.step4_6_cell_colors import assign_row_colors
from RFQ.units_convert.models import STATUS_CONVERTED
from utils.colors import Color
from tags.tag_classes import TagClass

_CONVERSION_TRACE_RE = re.compile(
    r"code=(?P<code>[^;]*);\s*src=(?P<src>[^;]*);\s*tgt=(?P<tgt>[^;]*);\s*"
    r"qty=(?P<qty>[^;]*);\s*coef=(?P<coef>[^;]*);\s*result=(?P<result>[^;]*);\s*"
    r"status=(?P<status>[^;]+)"
)


def apply_all_postmerge_per_row_ops(
    result_rows: List[RowStd],
    replacement_table: Dict[str, List[Tuple[str, str]]],
    lot_map: Dict[str, str] | None = None,
    code_base_by_code: Dict[str, RowStd] | None = None,
    assign_rfp_mto_code_compare_colors: bool = True,
) -> None:
    """Applies all per-row post-merge operations in correct order.

    Must be called AFTER collapse and BEFORE save_match_result_to_excel.
    Phase 2 of TAG_EFFECTIVE (global duplicate detection) stays in the
    main process.

    Before operations, creates full copies of each row via RowStd.get_row_copy,
    since some cells may share CheckElement objects (from get_row_copy_light split).
    Without copying, assign_row_colors would change color for multiple cells at once.
    """
    result_rows[:] = [RowStd.get_row_copy(row, row.t_com) for row in result_rows]

    set_equipment_type_status(result_rows)
    set_tags_match_summary(result_rows)
    set_tags_pair_statuses(result_rows)
    set_rfp_mto_struck_code_status(result_rows)
    if lot_map:
        apply_lot_map_to_rows(result_rows, lot_map)
    set_effective_tags_per_row(result_rows)
    # 
    set_mto_cabinet_equipment(result_rows)    
    # Раскрашиваем строки 
    assign_row_colors(
        result_rows,
        replacement_table=replacement_table,
        assign_rfp_mto_code_compare_colors=assign_rfp_mto_code_compare_colors,
    )
    apply_mto_rfp_quantity_diff(result_rows)
    # Загружаем Google-базу в workers и подставляем значения в RFP строки после всех операций
    enrich_rfp_from_google_base(result_rows, code_base_by_code=code_base_by_code)


def apply_mto_rfp_quantity_diff(result_rows: List[RowStd]) -> None:
    """Compute VALUES_MTO_RFP_DIFF and highlight compatible/incompatible units.

    After units_gate both RFP and MTO quantities are in Google units. If
    ``UNITS_MTO`` was not copied onto a matched row, fill it from RFP
    ``UNITS`` so the Google unit is visible and the numeric diff is not
    skipped. Excel comments on ``UNITS_MTO`` are added only for matrix
    conversions (not identity / spelling-only).
    """
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        if VALUES_MTO_RFP_DIFF not in row.el:
            row.el[VALUES_MTO_RFP_DIFF] = CheckElement(None)
        diff_el = row.el[VALUES_MTO_RFP_DIFF]
        diff_el.value = ""
        diff_el.color = Color.no
        diff_el.comment = ""

        _fill_missing_mto_units_from_rfp(row)
        _apply_converted_units_comment(row)

        rfp_units = normalize_units_text(row.get_value(UNITS))
        mto_units_el = row.el.get(UNITS_MTO)
        mto_units = normalize_units_text(
            mto_units_el.value if mto_units_el is not None else ""
        )
        if not rfp_units or not mto_units:
            continue
        if rfp_units != mto_units:
            for column in (UNITS, UNITS_MTO):
                if column not in row.el:
                    row.el[column] = CheckElement(None)
                row.el[column].color = Color.yellow
                row.el[column].comment = (
                    f"Несовместимые ед. изм.: RFP={rfp_units!r}, MTO={mto_units!r}"
                )
            continue

        try:
            rfp_qty = float(row.get_value(VALUES))
            mto_qty = float(row.get_value(VALUES_MTO))
        except (TypeError, ValueError):
            continue
        if not (math.isfinite(rfp_qty) and math.isfinite(mto_qty)):
            continue

        diff_value = mto_qty - rfp_qty
        diff_el.value = diff_value
        diff_el.color = Color.match_matched if diff_value == 0 else Color.yellow


def apply_mto_ul_quantity_diff(result_rows: List[RowStd]) -> None:
    """Compute VALUES_MTO_UL_DIFF after packing (VALUES_MTO − UL_VALUES).

    Same colouring as MTO−RFP: 0 → green, else yellow. Missing UL qty stays
    blank (not 0). Incompatible ``UNITS_MTO`` / ``UL_UNITS`` leave the diff
    empty and mark ``UL_UNITS`` yellow without overwriting an existing comment.
    """
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        if VALUES_MTO_UL_DIFF not in row.el:
            row.el[VALUES_MTO_UL_DIFF] = CheckElement(None)
        diff_el = row.el[VALUES_MTO_UL_DIFF]
        diff_el.value = ""
        diff_el.color = Color.no
        diff_el.comment = ""
        if is_excluded_from_supply(row):
            continue

        if not _has_numeric_quantity(row, VALUES_MTO):
            continue
        if not _has_numeric_quantity(row, UL_VALUES):
            continue

        mto_units = normalize_units_text(row.get_value(UNITS_MTO))
        ul_units = normalize_units_text(row.get_value(UL_UNITS))
        if not mto_units or not ul_units:
            continue
        if mto_units != ul_units:
            if UL_UNITS not in row.el:
                row.el[UL_UNITS] = CheckElement(None)
            ul_el = row.el[UL_UNITS]
            ul_el.color = Color.yellow
            if not str(ul_el.comment or "").strip():
                ul_el.comment = (
                    f"Несовместимые ед. изм.: MTO={mto_units!r}, УЛ={ul_units!r}"
                )
            continue

        try:
            mto_qty = float(row.get_value(VALUES_MTO))
            ul_qty = float(row.get_value(UL_VALUES))
        except (TypeError, ValueError):
            continue
        if not (math.isfinite(mto_qty) and math.isfinite(ul_qty)):
            continue

        diff_value = mto_qty - ul_qty
        diff_el.value = diff_value
        diff_el.color = Color.match_matched if diff_value == 0 else Color.yellow


def _has_numeric_quantity(row: RowStd, column: str) -> bool:
    try:
        quantity = float(row.get_value(column))
    except (TypeError, ValueError):
        return False
    return math.isfinite(quantity)


def _fill_missing_mto_units_from_rfp(row: RowStd) -> None:
    """Fill empty UNITS_MTO from RFP UNITS after gate (already Google unit)."""
    if UNITS_MTO not in row.el:
        row.el[UNITS_MTO] = CheckElement(None)
    if normalize_units_text(row.el[UNITS_MTO].value):
        return
    if not _has_numeric_quantity(row, VALUES_MTO):
        return
    rfp_units = normalize_units_text(row.get_value(UNITS))
    if rfp_units:
        row.el[UNITS_MTO].value = rfp_units


def _coef_is_unity(coef_text: str) -> bool:
    text = str(coef_text or "").strip().replace(",", ".")
    if not text:
        return False
    if "/" in text:
        try:
            numerator_text, denominator_text = text.split("/", 1)
            return float(numerator_text) == float(denominator_text)
        except (TypeError, ValueError):
            return False
    try:
        return abs(float(text) - 1.0) <= 1e-12
    except (TypeError, ValueError):
        return False


def converted_units_comment_text(status: object, trace: object) -> str:
    """Return Excel comment text for matrix-converted units, else empty.

    Identity, no_google, and spelling-only changes (``шт.`` → ``шт``) are
    skipped. Only traces with ``status=converted`` and a real unit/qty change
    are included.

    Args:
        status: Merged ``UNITS_CHECK_STATUS`` value.
        trace: Merged ``UNITS_CONVERSION_TRACE`` value.

    Returns:
        Multi-line comment or ``""``.
    """
    if STATUS_CONVERTED not in f"{status or ''} {trace or ''}":
        return ""
    lines: list[str] = []
    for match in _CONVERSION_TRACE_RE.finditer(str(trace or "")):
        if match.group("status").strip() != STATUS_CONVERTED:
            continue
        source_unit = match.group("src").strip()
        target_unit = match.group("tgt").strip()
        coef_text = match.group("coef").strip()
        if (
            normalize_units_text(source_unit) == normalize_units_text(target_unit)
            and _coef_is_unity(coef_text)
        ):
            continue
        lines.append(
            f"{match.group('qty').strip()} {source_unit} × {coef_text} → "
            f"{match.group('result').strip()} {target_unit}"
        )
    if not lines:
        return ""
    return "Конвертация по матрице:\n" + "\n".join(lines)


def _apply_converted_units_comment(row: RowStd) -> None:
    if UNITS_MTO not in row.el:
        row.el[UNITS_MTO] = CheckElement(None)
    comment = converted_units_comment_text(
        row.get_value(UNITS_CHECK_STATUS),
        row.get_value(UNITS_CONVERSION_TRACE),
    )
    if comment:
        row.el[UNITS_MTO].comment = comment


def set_equipment_type_status(result_rows: List[RowStd]) -> None:
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        equipment_values: List[str] = []
        rfp_tags = row.get_tags_list()
        for tag in rfp_tags:
            equipment = _safe_get_equipment(tag)
            if equipment and equipment not in equipment_values:
                equipment_values.append(equipment)
        tag_mto = row.el[TAG_MTO].value if row.el.get(TAG_MTO) else ""
        if tag_mto:
            equipment = _safe_get_equipment(tag_mto)
            if equipment and equipment not in equipment_values:
                equipment_values.append(equipment)
        tag_vo = row.el[TAG_VO].value if row.el.get(TAG_VO) else ""
        if tag_vo:
            equipment = _safe_get_equipment(tag_vo)
            if equipment and equipment not in equipment_values:
                equipment_values.append(equipment)
        row.el[EQUIPMENT_TYPE_STATUS].value = "/".join(equipment_values)


def set_tags_match_summary(result_rows: List[RowStd]) -> None:
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        rfp_tags = row.get_tags_list()
        tag_mto = row.el[TAG_MTO].value if row.el.get(TAG_MTO) else ""
        tag_vo = row.el[TAG_VO].value if row.el.get(TAG_VO) else ""
        status_mto = row.el[MATCH_STATUS].value or ""
        status_vo = row.el[MATCH_STATUS_VO].value or ""
        if (
            tag_mto
            or tag_vo
            or status_mto in {"Тег сопоставлен", "Добавлен из МТО", "Тег в МТО заменен"}
            or status_vo in {"Тег сопоставлен", "Тег в VO заменен"}
        ):
            row.el[MATCH_STATUS_TAGS].value = "Тег найден"
        elif not rfp_tags:
            row.el[MATCH_STATUS_TAGS].value = "Нет тегов"
        else:
            row.el[MATCH_STATUS_TAGS].value = "Тег не найден"


def set_tags_pair_statuses(result_rows: List[RowStd]) -> None:
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        rfp_tags = row.get_tags_list()
        tag_mto = row.el[TAG_MTO].value if row.el.get(TAG_MTO) else ""
        tag_vo = row.el[TAG_VO].value if row.el.get(TAG_VO) else ""

        if tag_vo and tag_mto:
            row.el[MATCH_STATUS_VO_MTO_TAGS].value = (
                "Совпадают" if tag_vo == tag_mto else "Различаются"
            )
        elif tag_vo:
            row.el[MATCH_STATUS_VO_MTO_TAGS].value = "Только VO"
        elif tag_mto:
            row.el[MATCH_STATUS_VO_MTO_TAGS].value = "Только MTO"
        else:
            row.el[MATCH_STATUS_VO_MTO_TAGS].value = "Нет тегов"

        if rfp_tags and tag_mto:
            row.el[MATCH_STATUS_RFP_MTO_TAGS].value = (
                "Совпадают" if tag_mto in rfp_tags else "Различаются"
            )
        elif rfp_tags:
            row.el[MATCH_STATUS_RFP_MTO_TAGS].value = "Только RFP"
        elif tag_mto:
            row.el[MATCH_STATUS_RFP_MTO_TAGS].value = "Только MTO"
        else:
            row.el[MATCH_STATUS_RFP_MTO_TAGS].value = "Нет тегов"


def set_rfp_mto_struck_code_status(result_rows: List[RowStd]) -> None:
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        rfp_code = row.get_value(CODE)
        mto_struck_code = row.el[MTO_CODE_STRUCK].value if row.el.get(MTO_CODE_STRUCK) else ""
        if not rfp_code and not mto_struck_code:
            row.el[MATCH_STATUS_RFP_MTO_STRUCK_CODE].value = ""
            continue
        if not rfp_code:
            continue
        if not mto_struck_code:
            continue
        rfp_norm = str(rfp_code).strip().upper()
        mto_norm = str(mto_struck_code).strip().upper()
        row.el[MATCH_STATUS_RFP_MTO_STRUCK_CODE].value = (
            "Совпадает c RFP" if rfp_norm == mto_norm else "Отличается от RFP"
        )


def set_effective_tags_per_row(result_rows: List[RowStd]) -> None:
    """Phase 1: per-row TAG_EFFECTIVE assignment.

    Phase 2 (global duplicate detection) stays in main process.
    """
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        vo_tag = row.el[TAG_VO].value if row.el.get(TAG_VO) else ""
        mto_tag = row.el[TAG_MTO].value if row.el.get(TAG_MTO) else ""
        rfp_tags = row.get_tags_list()
        position_status = row.el[POSITION_STATUS].value if row.el.get(POSITION_STATUS) else ""
        excluded = is_excluded_from_supply(row)

        effective_tag = ""
        if vo_tag:
            effective_tag = vo_tag
        elif mto_tag:
            effective_tag = mto_tag
        elif rfp_tags and position_status != "Исключен":
            code_mto = row.get_value(CODE_MTO)
            code_vo = row.get_value(CODE_VO)
            if not code_mto and not code_vo:
                if not excluded:
                    row.el[POSITION_STATUS].value = "Исключен"
                    effective_tag = ""
                else:
                    effective_tag = rfp_tags[0]
            else:
                effective_tag = rfp_tags[0]
        elif position_status != "Исключен":
            code_mto = row.get_value(CODE_MTO)
            code_vo = row.get_value(CODE_VO)
            if not code_mto and not code_vo and not excluded:
                row.el[POSITION_STATUS].value = "Исключен"

        row.el[TAG_EFFECTIVE].value = effective_tag


def set_mto_cabinet_equipment(result_rows: List[RowStd]) -> None:
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        current_value = ""
        mto_tag = row.el[TAG_MTO].value if row.el.get(TAG_MTO) else ""
        current_value = _merge_shield_number(current_value, _get_shield_number_from_tag(mto_tag))
        vo_tag = row.el[TAG_VO].value if row.el.get(TAG_VO) else ""
        current_value = _merge_shield_number(current_value, _get_shield_number_from_tag(vo_tag))
        rfp_tags = row.get_tags_list()
        rfp_tag_value = rfp_tags[0] if rfp_tags else ""
        current_value = _merge_shield_number(current_value, _get_shield_number_from_tag(rfp_tag_value))
        row.el[MTO_CABINET_EQUIPMENT].value = current_value


def enrich_rfp_from_google_base(
    result_rows: List[RowStd],
    code_base_by_code: Dict[str, tuple] | None = None,
) -> None:
    """Подставляет NAME, TYPE_MARK, VENDOR из Google-базы по коду RFP.
    Если код найден — записывает значения и подсвечивает ячейки Color.is_replacement_mto_vo.
    code_base_by_code — словарь {code: (name, type_mark, vendor)}, строится один раз
    в main-процессе и передаётся в workers. Tuple вместо RowStd минимизирует pickle-объём.
    """
    if not code_base_by_code:
        return
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        code = row.el[CODE].value
        if not code:
            continue
        entry = code_base_by_code.get(code)  # (name_val, type_mark_val, vendor_val)
        if entry is None:
            continue
        for col, base_val in zip((NAME, TYPE_MARK, VENDOR), entry):
            if base_val is not None and str(base_val).strip():
                row.el[col].value = base_val
            Color.set_el_color(row.el[col], Color.match_matched)


def apply_lot_map_to_rows(
    result_rows: List[RowStd],
    lot_map: Dict[str, str],
) -> None:
    for row in result_rows:
        if row.row_type != RowType.position_row:
            continue
        ds_name = _normalize_ds_name(row.get_value(DS_NAME))
        row.el[DS_LOT].value = lot_map.get(ds_name, "")


def _normalize_ds_name(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _safe_get_equipment(tag: str) -> str:
    if not tag:
        return ""
    try:
        tag_obj = TagClass(str(tag))
        return tag_obj.equipment or ""
    except Exception:
        return ""


def _get_shield_number_from_tag(tag_value) -> str:
    if not tag_value:
        return ""
    tag_value = tag_value[0] if isinstance(tag_value, list) and tag_value else tag_value
    try:
        tag_obj = TagClass(str(tag_value))
        return tag_obj.get_shield_number() or ""
    except Exception:
        return ""


def _merge_shield_number(current_value: str, new_value: str) -> str:
    if not new_value:
        return current_value or ""
    if not current_value:
        return new_value
    if current_value == new_value:
        return current_value
    parts = [p.strip() for p in str(current_value).split("/") if p.strip()]
    if new_value in parts:
        return current_value
    return f"{current_value}/{new_value}"
