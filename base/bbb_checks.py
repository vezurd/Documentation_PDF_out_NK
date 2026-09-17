import re
from collections import Counter, defaultdict
from typing import Any, Dict

from base.base_classes import RowStd, RowType
from base.base_cheks import (
    check_by_code, check_tags_value, check_value,
    check_duplicate_tags, check_prohibition, check_equipment_codes,
    check_mass,
)
from base.tables_columns import (
    NAME, TYPE_MARK, UNITS, VENDOR, CODE, TAGS, VALUES, NUMBERS, ROW_TYPE,
    BBB_WORK_CODE, BBB_MTR_GROUP, BBB_WORK_CODE_LIST, TITLE, BBB_MARKA, BBB_REVISION, ANNOTATION,
    SECTION_TYPE,
)
from pdf_parsing_v2_engine.document import V2Document
from tags.tag_parser import get_bbb_work_code
from utils.colors import Color


_COLUMNS_FOR_CHECK = [NAME, TYPE_MARK, UNITS, VENDOR, CODE]


def _normalize_simple_text(value: Any) -> str:
    return str(value or "").strip()


def _build_code_base_work_indexes(code_base_data_std):
    work_codes_by_code = {}
    mtr_group_by_code = {}

    for row in code_base_data_std or []:
        code_val = _normalize_simple_text(row.el[CODE].value)
        if not code_val:
            continue

        work_code_list_raw = _normalize_simple_text(row.el[BBB_WORK_CODE_LIST].value)
        work_codes_by_code[code_val] = set(get_bbb_work_code(work_code_list_raw))
        mtr_group_by_code[code_val] = _normalize_simple_text(row.el[BBB_MTR_GROUP].value)

    return work_codes_by_code, mtr_group_by_code


def _validate_bbb_work_fields_by_google(bbb_data, code_base_data_std):
    work_codes_by_code, mtr_group_by_code = _build_code_base_work_indexes(code_base_data_std)

    for row in bbb_data:
        if row.row_type != RowType.position_row:
            continue
        if _normalize_simple_text(row.t_com.sheet_name) not in ("BOE", "BOM"):
            continue

        code_val = _normalize_simple_text(row.el[CODE].value)
        if not code_val:
            continue

        # BBB_WORK_CODE vs BBB_WORK_CODE_LIST (Google)
        expected_work_codes = work_codes_by_code.get(code_val, set())
        bbb_work_code_raw = _normalize_simple_text(row.el[BBB_WORK_CODE].value)
        actual_work_codes = set(get_bbb_work_code(bbb_work_code_raw))
        if not actual_work_codes and bbb_work_code_raw:
            actual_work_codes = {bbb_work_code_raw}

        if expected_work_codes:
            if actual_work_codes and (actual_work_codes & expected_work_codes):
                Color.set_el_color(row.el[BBB_WORK_CODE], Color.green)
            else:
                row.el[BBB_WORK_CODE].comment += (
                    "BBB_WORK_CODE не совпал с GoogleBase.\n"
                    f"Допустимые BBB_WORK_CODE_LIST: {', '.join(sorted(expected_work_codes))}\n"
                )
                Color.set_el_color(row.el[BBB_WORK_CODE], Color.red)
        else:
            row.el[BBB_WORK_CODE].comment += (
                "В GoogleBase не заполнен BBB_WORK_CODE_LIST для этого CODE. "
                "Добавьте список кодов работ в Google-таблице.\n"
            )
            Color.set_el_color(row.el[BBB_WORK_CODE], Color.yellow)

        # BBB_MTR_GROUP vs Google BBB_MTR_GROUP
        expected_mtr_group = mtr_group_by_code.get(code_val, "")
        bbb_mtr_group = _normalize_simple_text(row.el[BBB_MTR_GROUP].value)
        if expected_mtr_group:
            if bbb_mtr_group == expected_mtr_group:
                Color.set_el_color(row.el[BBB_MTR_GROUP], Color.green)
            else:
                row.el[BBB_MTR_GROUP].comment += (
                    "BBB_MTR_GROUP не совпал с GoogleBase.\n"
                    f"Ожидаемое значение: {expected_mtr_group}\n"
                )
                Color.set_el_color(row.el[BBB_MTR_GROUP], Color.red)
        else:
            row.el[BBB_MTR_GROUP].comment += (
                "В GoogleBase не заполнено поле 'Группа МТР' для этого CODE. "
                "Добавьте значение в Google-таблице.\n"
            )
            Color.set_el_color(row.el[BBB_MTR_GROUP], Color.yellow)


def bbb_vs_code_base(bbb_data, code_base_data_std, cfg: Dict[str, Any] | None = None):
    """
    Этап 1: проверка BBB данных по code_base (аналог base_vs_base_std).
    cfg — секция bbb_vs_code_base из bbb_config.
    """
    if cfg is None:
        cfg = {}

    if cfg.get("check_by_code", True):
        check_by_code(
            bbb_data,
            code_base_data_std,
            _COLUMNS_FOR_CHECK,
            mode=("correction", "comment"),
            enable_comment=False,
            enable_soft_compare_color=False)

    RowStd.set_color_by_row_type_in_list(bbb_data)

    if cfg.get("check_equipment_codes", True):
        check_equipment_codes(bbb_data, code_base_data_std)
    if cfg.get("check_tags_value", True):
        check_tags_value(bbb_data)
    if cfg.get("check_value", True):
        check_value(bbb_data)
    if cfg.get("check_duplicate_tags", True):
        check_duplicate_tags(bbb_data)
    if cfg.get("check_mass", True):
        check_mass(bbb_data, code_base_data_std)
    if cfg.get("check_prohibition", False):
        check_prohibition(bbb_data, code_base_data_std)
    if cfg.get("check_work_code_and_mtr_group", True):
        _validate_bbb_work_fields_by_google(bbb_data, code_base_data_std)

    return bbb_data


def _new_doc_value_map():
    return {"BOE": 0.0, "BOM": 0.0}


def _new_doc_rows_map():
    return {"BOE": [], "BOM": []}


def _new_doc_template_map():
    return {"BOE": None, "BOM": None}


def _build_mto_index(mto_data):
    """
    Строит индекс MTO строк по CODE.
    Возвращает dict: {code_str: [RowStd, ...]}
    """
    index = defaultdict(list)
    for row in mto_data:
        if row.row_type != RowType.position_row:
            continue
        code_val = str(row.el[CODE].value or "").strip()
        if code_val:
            index[code_val].append(row)
    return index


def _build_mto_code_tag_index(mto_data):
    """
    Строит индекс MTO строк по CODE и TAG.
    Возвращает dict: {code_str: {tag_str: RowStd}}
    """
    index = defaultdict(dict)
    for row in mto_data:
        if row.row_type != RowType.position_row:
            continue
        code_val = str(row.el[CODE].value or "").strip()
        if not code_val:
            continue
        tags = _get_tags_set(row)
        for tag in tags:
            if tag not in index[code_val]:
                index[code_val][tag] = row
    return index


def _set_light_blue_diff_color(row: RowStd):
    """Подсветка добавленной diff-строки (MTO-only tag)."""
    diff_color = Color.match_added_mto
    for col in (CODE, TAGS, NAME, TYPE_MARK, VENDOR, UNITS, VALUES):
        Color.set_el_color(row.el[col], diff_color)


def _build_mto_only_tag_row(code_val: str, tag: str, template_row: RowStd | None, mto_row: RowStd | None):
    """
    Создаёт строку diff для тега, который есть в MTO, но отсутствует в BBB.
    В качестве основы берёт BBB-шаблон по коду; если его нет, берёт MTO-строку.
    """
    base_row = template_row or mto_row
    if base_row is None:
        return None

    out_row = RowStd.get_row_copy(base_row, t_com=base_row.t_com)
    out_row.row_type = RowType.position_row
    out_row.el[ROW_TYPE].value = RowType.position_row
    out_row.el[NUMBERS].value = ""
    out_row.el[CODE].value = code_val
    out_row.el[TAGS].value = [tag]
    out_row.el[TAGS].comment += f"Тег есть в MTO, отсутствует в BBB: {tag}\n"

    if mto_row is not None:
        for col in (NAME, TYPE_MARK, VENDOR, UNITS, VALUES):
            out_row.el[col].value = mto_row.el[col].value

    _set_light_blue_diff_color(out_row)
    return out_row


def _get_tags_set(row):
    """Извлекает множество тегов из строки."""
    tags_val = row.el[TAGS].value
    if not tags_val:
        return set()
    if isinstance(tags_val, list):
        return {t.strip() for t in tags_val if isinstance(t, str) and t.strip()}
    tags_str = str(tags_val).strip()
    if not tags_str:
        return set()
    tags = set()
    for sep in ("\n", ",", ";"):
        if sep in tags_str:
            for t in tags_str.split(sep):
                t = t.strip()
                if t:
                    tags.add(t)
            return tags
    tags.add(tags_str)
    return tags


def _extract_bbb_tags(row: RowStd, bbb_type: str):
    """Единая точка получения тегов BBB после унификации BOM -> TAGS."""
    del bbb_type
    return _get_tags_set(row)


def _parse_values(val, row: RowStd):
    """Парсит значение VALUES в float; при ошибке поднимает ValueError с контекстом."""
    doc_name = str(getattr(getattr(row, "t_com", None), "file_name", "") or "unknown_document")
    code_val = str(row.el[CODE].value or "").strip() or "unknown_code"
    row_no = str(row.el[NUMBERS].value or "").strip() or "unknown_row"

    if val is None:
        raise ValueError(
            "Не удалось преобразовать VALUES в число: "
            f"документ='{doc_name}', строка='{row_no}', код='{code_val}', value='{val}'"
        )
    try:
        return float(val)
    except (ValueError, TypeError):
        s = str(val).strip().replace(",", ".").replace(" ", "")
        try:
            return float(s)
        except (ValueError, TypeError):
            raise ValueError(
                "Не удалось преобразовать VALUES в число: "
                f"документ='{doc_name}', строка='{row_no}', код='{code_val}', value='{val}'"
            )


def _get_bbb_vs_mto_flags(cfg: Dict[str, Any]):
    return (
        cfg.get("compare_fields", True),
        cfg.get("compare_values", True),
        cfg.get("compare_tags", True),
        cfg.get("compare_title_marka_revision", True),
        cfg.get("compare_bom_annotation_with_mto", True),
        cfg.get("check_missing_mto_codes", True),
    )


def _extract_mto_doc_values(mto_data) -> dict[str, str]:
    """
    Достаёт TITLE/MARKA/REVISION из имени MTO файла через V2Document.from_file_path.
    """
    if not mto_data:
        return {}

    t_com = getattr(mto_data[0], "t_com", None)
    if t_com is None:
        return {}

    source_name = str(getattr(t_com, "file_full_path", "") or getattr(t_com, "file_name", "")).strip()
    if not source_name:
        return {}

    try:
        mto_doc = V2Document.from_file_path(source_name)
    except Exception:
        return {}

    return {
        TITLE: _normalize_simple_text(mto_doc.doc_Title_4d),
        BBB_MARKA: _normalize_simple_text(mto_doc.doc_Marka),
        BBB_REVISION: _normalize_simple_text(mto_doc.doc_Revision),
    }


def _compare_title_marka_revision_with_mto(row: RowStd, mto_doc_values: dict[str, str]):
    """
    Сравнивает TITLE/BBB_MARKA/BBB_REVISION строки BBB со значениями из имени MTO файла.
    """
    if not mto_doc_values:
        return

    for col, label in (
            (TITLE, "TITLE"),
            (BBB_MARKA, "BBB_MARKA"),
            (BBB_REVISION, "BBB_REVISION"),
    ):
        if col not in row.el:
            continue
        expected_val = _normalize_simple_text(mto_doc_values.get(col, ""))
        actual_val = _normalize_simple_text(row.el[col].value)
        if actual_val == expected_val:
            Color.set_el_color(row.el[col], Color.green)
            continue

        row.el[col].comment += (
            f"{label} не совпадает с MTO.\n"
            f"  В MTO: {expected_val or '<пусто>'}\n"
        )
        Color.set_el_color(row.el[col], Color.red)


def _aggregate_mto_by_code(mto_data, do_values: bool, do_tags: bool):
    mto_values_by_code = defaultdict(float)
    mto_tags_by_code = defaultdict(set)
    for row in mto_data:
        if row.row_type != RowType.position_row:
            continue
        code_val = str(row.el[CODE].value or "").strip()
        if not code_val:
            continue
        if do_values:
            mto_values_by_code[code_val] += _parse_values(row.el[VALUES].value, row=row)
        if do_tags:
            mto_tags_by_code[code_val] |= _get_tags_set(row)
    return mto_values_by_code, mto_tags_by_code


def _aggregate_bbb_combined(boe_data, bom_data, do_values: bool, do_tags: bool):
    values_total_by_code = defaultdict(float)
    tags_union_by_code = defaultdict(set)
    values_breakdown_by_code = defaultdict(_new_doc_value_map)
    row_refs_by_code = defaultdict(_new_doc_rows_map)
    template_row_by_code = defaultdict(_new_doc_template_map)

    for bbb_type, data in (("BOE", boe_data), ("BOM", bom_data)):
        for row in data:
            if row.row_type != RowType.position_row:
                continue
            code_val = str(row.el[CODE].value or "").strip()
            if not code_val:
                continue

            row_refs_by_code[code_val][bbb_type].append(row)
            if template_row_by_code[code_val][bbb_type] is None:
                template_row_by_code[code_val][bbb_type] = row

            if do_values:
                val = _parse_values(row.el[VALUES].value, row=row)
                values_total_by_code[code_val] += val
                values_breakdown_by_code[code_val][bbb_type] += val

            if do_tags:
                row_tags = _extract_bbb_tags(row, bbb_type)
                tags_union_by_code[code_val] |= row_tags

    return (
        values_total_by_code,
        tags_union_by_code,
        values_breakdown_by_code,
        row_refs_by_code,
        template_row_by_code,
    )


def _add_combined_values_comment(row: RowStd, code_val: str, values_breakdown_by_code, total_sum: float, mto_sum: float):
    breakdown = values_breakdown_by_code.get(code_val, _new_doc_value_map())
    row.el[VALUES].comment += (
        f"Расхождение VALUES по CODE {code_val}:\n"
        f"  BOE: {breakdown['BOE']}\n"
        f"  BOM: {breakdown['BOM']}\n"
        f"  BOE+BOM: {total_sum}\n"
        f"  MTO: {mto_sum}\n"
    )
    Color.set_el_color(row.el[VALUES], Color.red)


def _extract_annotation_terms(annotation_text: str) -> list[int]:
    """
    Извлекает слагаемые из ANNOTATION.
    Берёт только ведущие целые числа в частях, разделённых '+'.
    """
    if not annotation_text:
        return []

    # Если строка уже дополнялась блоком "Сумма из МТО:", не учитываем добавленную часть в сравнении.
    source = annotation_text.split("Сумма из МТО:")[0]
    terms: list[int] = []
    for part in source.split("+"):
        m = re.match(r"^\s*(\d+)", part)
        if m:
            terms.append(int(m.group(1)))
    return terms


def _values_match_as_multiset(left: list[int], right: list[int]) -> bool:
    return Counter(left) == Counter(right)


def _mto_row_to_annotation_term(mto_row: RowStd) -> tuple[int, str]:
    val_num = _parse_values(mto_row.el[VALUES].value, row=mto_row)
    val_int = int(round(val_num))
    pos = str(mto_row.el[NUMBERS].value or "").strip()
    if pos:
        return val_int, f"{val_int}(поз. {pos})"
    return val_int, str(val_int)


def _set_annotation_mto_sum_block(row: RowStd, right_sum_text: str):
    marker = "Сумма из МТО:"
    current = str(row.el[ANNOTATION].value or "")
    base = current
    if marker in base:
        base = base.split(marker)[0].rstrip()

    # Визуально отделяем корректную сумму переносом строки.
    if base:
        row.el[ANNOTATION].value = f"{base}\n{marker}\n{right_sum_text}"
    else:
        row.el[ANNOTATION].value = f"{marker}\n{right_sum_text}"


def _find_subset_indices_for_sum(values: list[int], target_sum: int) -> list[int] | None:
    if target_sum < 0:
        return None
    if target_sum == 0:
        return []

    memo = {}

    def _dfs(i: int, remain: int):
        if remain == 0:
            return ()
        if i >= len(values) or remain < 0:
            return None
        key = (i, remain)
        if key in memo:
            return memo[key]

        take = _dfs(i + 1, remain - values[i])
        if take is not None:
            memo[key] = (i, *take)
            return memo[key]

        skip = _dfs(i + 1, remain)
        memo[key] = skip
        return skip

    found = _dfs(0, target_sum)
    if found is None:
        return None
    return list(found)


def _take_terms_for_target(terms_pool: list[tuple[int, str]], target_sum: int):
    """
    Берёт из residual-пула сумму target_sum.
    Допускает частичное использование одного MTO-терма (например 30 -> 6 + остаток 24),
    что нужно для BOM, где один MTO-терм может быть размазан на несколько строк.
    """
    if target_sum < 0:
        return None, terms_pool
    if target_sum == 0:
        return [], terms_pool

    remain = target_sum
    selected: list[tuple[int, str]] = []
    new_pool: list[tuple[int, str]] = []

    for val, label in terms_pool:
        if remain <= 0:
            new_pool.append((val, label))
            continue

        if val <= remain:
            selected.append((val, label))
            remain -= val
            continue

        # Частично расходуем текущий терм.
        take = remain
        selected.append((take, label))
        new_pool.append((val - take, label))
        remain = 0

    if remain > 0:
        return None, terms_pool

    return selected, new_pool


def _subtract_sum_from_mto_terms(terms: list[tuple[int, str]], subtract_sum: int) -> list[tuple[int, str]]:
    """
    Вычитает сумму BOE из MTO-слагаемых по принципу sum-only.
    Возвращает остаточный пул MTO-термов для BOM.
    """
    remain_to_subtract = max(0, subtract_sum)
    residual: list[tuple[int, str]] = []

    for val, label in terms:
        if remain_to_subtract <= 0:
            residual.append((val, label))
            continue

        if val <= remain_to_subtract:
            remain_to_subtract -= val
            continue

        new_val = val - remain_to_subtract
        remain_to_subtract = 0
        m = re.search(r"\(поз\.\s*([^)]+)\)", label)
        if m:
            residual.append((new_val, f"{new_val}(поз. {m.group(1)})"))
        else:
            residual.append((new_val, str(new_val)))

    return residual


def _terms_to_texts(terms: list[tuple[int, str]]) -> tuple[str, str]:
    nums = [x[0] for x in terms]
    with_pos = "+".join(x[1] for x in terms) if terms else "0"
    plain = "+".join(str(x) for x in nums) if nums else "0"
    return with_pos, plain


def _validate_bom_annotation_row(row: RowStd, expected_terms: list[tuple[int, str]], force_mismatch_comment: str | None = None):
    ann_text = str(row.el[ANNOTATION].value or "").strip()
    ann_numbers = _extract_annotation_terms(ann_text)
    expected_numbers = [x[0] for x in expected_terms]
    right_sum_text_with_pos, right_sum_text_plain = _terms_to_texts(expected_terms)

    if force_mismatch_comment is not None:
        row.el[ANNOTATION].comment += (
            f"{force_mismatch_comment}\n"
            f"Правильная последовательность: {right_sum_text_with_pos}\n"
        )
        _set_annotation_mto_sum_block(row, right_sum_text_plain)
        Color.set_el_color(row.el[ANNOTATION], Color.red)
        return

    if len(expected_numbers) == 1:
        if not ann_text:
            return
        if _values_match_as_multiset(ann_numbers, expected_numbers):
            row.el[ANNOTATION].comment += "Не требуется, позиция встречается единично.\n"
            Color.set_el_color(row.el[ANNOTATION], Color.yellow)
            return
        row.el[ANNOTATION].comment += (
            "ANNOTATION не совпадает с суммой из MTO с учетом BOE.\n"
            f"Правильная последовательность: {right_sum_text_with_pos}\n"
        )
        _set_annotation_mto_sum_block(row, right_sum_text_plain)
        Color.set_el_color(row.el[ANNOTATION], Color.red)
        return

    # Если ожидается несколько слагаемых, ANNOTATION обязателен и должен совпасть по мультимножеству.
    if _values_match_as_multiset(ann_numbers, expected_numbers):
        Color.set_el_color(row.el[ANNOTATION], Color.green)
        return

    row.el[ANNOTATION].comment += (
        "ANNOTATION не совпадает с суммой из MTO с учетом BOE.\n"
        f"Правильная последовательность: {right_sum_text_with_pos}\n"
    )
    _set_annotation_mto_sum_block(row, right_sum_text_plain)
    Color.set_el_color(row.el[ANNOTATION], Color.red)


def _group_position_rows_by_code(rows: list[RowStd]) -> dict[str, list[RowStd]]:
    out: dict[str, list[RowStd]] = defaultdict(list)
    for row in rows:
        if row.row_type != RowType.position_row:
            continue
        code_val = str(row.el[CODE].value or "").strip()
        if code_val:
            out[code_val].append(row)
    return out


def _validate_bom_annotations_with_mto_residual(boe_data, bom_data, mto_index):
    """
    Проверяет BOM.ANNOTATION по остаточному пулу MTO после вычитания BOE по sum-only.
    Учитывает все строки BOM для одного CODE и распределяет residual термы построчно.
    """
    boe_rows_by_code = _group_position_rows_by_code(boe_data)
    bom_rows_by_code = _group_position_rows_by_code(bom_data)

    for code_val, bom_rows in bom_rows_by_code.items():
        mto_rows = mto_index.get(code_val, [])
        if not mto_rows:
            continue

        mto_terms = [_mto_row_to_annotation_term(r) for r in mto_rows]
        boe_total = int(round(sum(_parse_values(r.el[VALUES].value, row=r) for r in boe_rows_by_code.get(code_val, []))))
        residual_terms = _subtract_sum_from_mto_terms(mto_terms, boe_total)

        # Если после учета BOE остатка нет, для BOM ANNOTATION сумма из MTO не требуется.
        if not residual_terms:
            for row in bom_rows:
                ann_text = str(row.el[ANNOTATION].value or "").strip()
                if ann_text:
                    _validate_bom_annotation_row(
                        row,
                        expected_terms=[],
                        force_mismatch_comment="После вычитания BOE сумма для BOM по этому CODE не требуется.",
                    )
            continue

        pool = residual_terms[:]
        # Сначала строки с непустым ANNOTATION, затем пустые: это уменьшает ложные распределения.
        ordered_rows = sorted(bom_rows, key=lambda r: (0 if str(r.el[ANNOTATION].value or "").strip() else 1))

        for row in ordered_rows:
            target_sum = int(round(_parse_values(row.el[VALUES].value, row=row)))
            selected, new_pool = _take_terms_for_target(pool, target_sum)
            if selected is None:
                _validate_bom_annotation_row(
                    row,
                    expected_terms=pool[:],
                    force_mismatch_comment=(
                        "Не удалось распределить остаточные суммы MTO по строкам BOM "
                        f"(CODE {code_val}, target={target_sum})."
                    ),
                )
                continue

            pool = new_pool
            _validate_bom_annotation_row(row, expected_terms=selected)


def _compare_bbb_row_with_mto_combined(
        row: RowStd,
        bbb_type: str,
        do_fields: bool,
        do_values: bool,
        do_tags: bool,
        do_title_marka_revision: bool,
        do_bom_annotation_with_mto: bool,
        mto_doc_values: dict[str, str],
        mto_index,
        values_total_by_code,
        values_breakdown_by_code,
        mto_values_by_code,
        mto_tags_by_code,):
    code_val = str(row.el[CODE].value or "").strip()
    if not code_val:
        return

    mto_rows = mto_index.get(code_val)
    if not mto_rows:
        row.el[CODE].comment += "Код не найден в MTO\n"
        Color.set_el_color(row.el[CODE], Color.red)
        return

    mto_first = mto_rows[0]
    if do_fields:
        _compare_field(row, mto_first, NAME, "MTO")
        _compare_field(row, mto_first, TYPE_MARK, "MTO")
        _compare_field(row, mto_first, VENDOR, "MTO")
        _compare_field(row, mto_first, UNITS, "MTO")
    if do_title_marka_revision:
        _compare_title_marka_revision_with_mto(row, mto_doc_values)
    del do_bom_annotation_with_mto

    if do_values:
        total_sum = values_total_by_code.get(code_val, 0.0)
        mto_sum = mto_values_by_code.get(code_val, 0.0)
        if abs(total_sum - mto_sum) > 0.001:
            _add_combined_values_comment(row, code_val, values_breakdown_by_code, total_sum, mto_sum)
        else:
            Color.set_el_color(row.el[VALUES], Color.green)

    if do_tags:
        mto_tags = mto_tags_by_code.get(code_val, set())
        row_tags = _extract_bbb_tags(row, bbb_type)
        missing_in_mto = sorted(tag for tag in row_tags if tag not in mto_tags)
        if missing_in_mto:
            row.el[TAGS].comment += (
                f"Не найдено в MTO по CODE {code_val}: {', '.join(missing_in_mto)}\n"
            )
            Color.set_el_color(row.el[TAGS], Color.red)
        elif row_tags:
            Color.set_el_color(row.el[TAGS], Color.green)


def _select_diff_target_doc(code_val: str, row_refs_by_code):
    doc_rows = row_refs_by_code.get(code_val, _new_doc_rows_map())
    has_boe = bool(doc_rows["BOE"])
    has_bom = bool(doc_rows["BOM"])

    if has_boe and has_bom:
        return "BOE"
    if has_boe:
        return "BOE"
    if has_bom:
        return "BOM"
    return "BOE"


def _build_mto_only_diff_rows_combined(
        tags_union_by_code,
        row_refs_by_code,
        template_row_by_code,
        mto_tags_by_code,
        mto_index,
        mto_code_tag_index,
):
    diff_rows_by_doc = {"BOE": [], "BOM": []}

    for code_val, mto_tags in mto_tags_by_code.items():
        bbb_tags = tags_union_by_code.get(code_val, set())
        only_in_mto = sorted(mto_tags - bbb_tags)
        if not only_in_mto:
            continue

        target_doc = _select_diff_target_doc(code_val, row_refs_by_code)
        template_row = template_row_by_code.get(code_val, _new_doc_template_map()).get(target_doc)
        mto_rows = mto_index.get(code_val, [])

        for tag in only_in_mto:
            mto_row = mto_code_tag_index.get(code_val, {}).get(tag)
            if mto_row is None and mto_rows:
                mto_row = mto_rows[0]
            diff_row = _build_mto_only_tag_row(code_val, tag, template_row, mto_row)
            if diff_row is not None:
                diff_rows_by_doc[target_doc].append(diff_row)

    return diff_rows_by_doc


def _build_missing_code_diff_rows(
        mto_index: dict,
        boe_codes: set,
        bom_codes: set,
        section_types_cfg: dict,
) -> dict:
    """
    Строит diff-строки для кодов из MTO, которые полностью отсутствуют в BOE и BOM.
    Целевой документ (BOE/BOM) определяется по SECTION_TYPE первой MTO-строки кода
    через конфиг section_types.
    """
    from base.bbb_config import get_section_target

    diff_rows_by_doc = {"BOE": [], "BOM": []}

    for code_val, mto_rows in mto_index.items():
        if code_val in boe_codes or code_val in bom_codes:
            continue  # Код уже представлен в BBB

        if not mto_rows:
            continue

        mto_first = mto_rows[0]
        section_name = str(mto_first.el[SECTION_TYPE].value or "").strip()
        target_doc = get_section_target(section_types_cfg, section_name)

        # Создаём одну diff-строку на код (без разбивки по тегам — код полностью отсутствует)
        out_row = RowStd.get_row_copy(mto_first, t_com=mto_first.t_com)
        out_row.row_type = RowType.position_row
        out_row.el[ROW_TYPE].value = RowType.position_row
        out_row.el[NUMBERS].value = ""
        out_row.el[CODE].value = code_val

        section_info = f" (секция: {section_name})" if section_name else ""
        out_row.el[CODE].comment += (
            f"Код из MTO не найден ни в BOE, ни в BOM{section_info}.\n"
            f"Добавлено в {target_doc} автоматически.\n"
        )

        for col in (CODE, TAGS, NAME, TYPE_MARK, VENDOR, UNITS, VALUES):
            Color.set_el_color(out_row.el[col], Color.match_added_mto)

        diff_rows_by_doc[target_doc].append(out_row)

    return diff_rows_by_doc


def bbb_vs_mto_combined(boe_data, bom_data, mto_data, cfg: Dict[str, Any] | None = None):
    """
    Этап 2: объединённое сравнение BOE+BOM с MTO по CODE.
    VALUES и TAGS проверяются по суммарной картине двух BBB документов,
    при этом построчные field-checks сохраняются отдельно для каждой строки.
    """
    if cfg is None:
        cfg = {}

    boe_data = boe_data or []
    bom_data = bom_data or []

    do_fields, do_values, do_tags, do_title_marka_revision, do_bom_annotation_with_mto, do_missing_codes = _get_bbb_vs_mto_flags(cfg)
    mto_doc_values = _extract_mto_doc_values(mto_data)
    mto_index = _build_mto_index(mto_data)
    mto_code_tag_index = _build_mto_code_tag_index(mto_data) if do_tags else {}
    mto_values_by_code, mto_tags_by_code = _aggregate_mto_by_code(mto_data, do_values=do_values, do_tags=do_tags)
    (
        values_total_by_code,
        tags_union_by_code,
        values_breakdown_by_code,
        row_refs_by_code,
        template_row_by_code,
    ) = _aggregate_bbb_combined(boe_data, bom_data, do_values=do_values, do_tags=do_tags)

    for bbb_type, data in (("BOE", boe_data), ("BOM", bom_data)):
        for row in data:
            if row.row_type != RowType.position_row:
                continue
            _compare_bbb_row_with_mto_combined(
                row=row,
                bbb_type=bbb_type,
                do_fields=do_fields,
                do_values=do_values,
                do_tags=do_tags,
                do_title_marka_revision=do_title_marka_revision,
                do_bom_annotation_with_mto=do_bom_annotation_with_mto,
                mto_doc_values=mto_doc_values,
                mto_index=mto_index,
                values_total_by_code=values_total_by_code,
                values_breakdown_by_code=values_breakdown_by_code,
                mto_values_by_code=mto_values_by_code,
                mto_tags_by_code=mto_tags_by_code,
            )

    if do_bom_annotation_with_mto:
        _validate_bom_annotations_with_mto_residual(
            boe_data=boe_data,
            bom_data=bom_data,
            mto_index=mto_index,
        )

    if do_tags:
        diff_rows_by_doc = _build_mto_only_diff_rows_combined(
            tags_union_by_code=tags_union_by_code,
            row_refs_by_code=row_refs_by_code,
            template_row_by_code=template_row_by_code,
            mto_tags_by_code=mto_tags_by_code,
            mto_index=mto_index,
            mto_code_tag_index=mto_code_tag_index,
        )
        if diff_rows_by_doc["BOE"]:
            boe_data = [*boe_data, *diff_rows_by_doc["BOE"]]
        if diff_rows_by_doc["BOM"]:
            bom_data = [*bom_data, *diff_rows_by_doc["BOM"]]

    if do_missing_codes:
        boe_codes = {
            str(row.el[CODE].value or "").strip()
            for row in boe_data
            if row.row_type == RowType.position_row
        }
        bom_codes = {
            str(row.el[CODE].value or "").strip()
            for row in bom_data
            if row.row_type == RowType.position_row
        }
        # Убираем пустые строки из наборов
        boe_codes.discard("")
        bom_codes.discard("")

        section_types_cfg = cfg.get("section_types", {})
        missing_diff = _build_missing_code_diff_rows(
            mto_index=mto_index,
            boe_codes=boe_codes,
            bom_codes=bom_codes,
            section_types_cfg=section_types_cfg,
        )
        if missing_diff["BOE"]:
            boe_data = [*boe_data, *missing_diff["BOE"]]
        if missing_diff["BOM"]:
            bom_data = [*bom_data, *missing_diff["BOM"]]

    return boe_data, bom_data


def _compare_field(bbb_row, mto_row, field, source_label):
    """Сравнивает поле BBB и MTO строк, подсвечивает расхождения."""
    bbb_val = str(bbb_row.el[field].value or "").strip()
    mto_val = str(mto_row.el[field].value or "").strip()

    if not bbb_val or not mto_val:
        return

    if bbb_val == mto_val:
        return

    bbb_row.el[field].comment += (
        f"Не совпадает с {source_label}:\n"
        f"  По {source_label}: {mto_val}\n"
    )
    Color.set_el_color(bbb_row.el[field], Color.yellow)
