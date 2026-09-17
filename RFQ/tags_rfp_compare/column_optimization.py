"""
Оптимизация столбцов: обрезка строк и обнуление (load: false).
Применяется после загрузки данных в рамках tags_rfp_compare.
"""

from __future__ import annotations

from typing import Any

from base.base_classes import RowStd, RowType
from base.tables_columns import TAGS


def apply_column_optimization(rows: list[RowStd], col_cfg: dict[str, Any]) -> None:
    """Модифицирует rows in-place: обрезка строк и обнуление столбцов по конфигу.

    Args:
        rows: Список строк RowStd для модификации
        col_cfg: Конфиг по столбцам, например {"name": {"load": true, "truncate": 200}}
            - load: false — обнулить value и struck_value
            - truncate: N — обрезать строки до N символов
    """
    if not col_cfg or not rows:
        return

    for row in rows:
        for key in list(row.el.keys()):
            cfg = col_cfg.get(key, {})
            if not cfg:
                continue
            el = row.el[key]
            if el is None:
                continue

            if cfg.get("load") is False:
                el.value = ""
                el.struck_value = ""
                continue

            truncate_n = cfg.get("truncate")
            if truncate_n is not None and isinstance(truncate_n, (int, float)):
                n = int(truncate_n)
                if n >= 0:
                    if isinstance(el.value, str) and len(el.value) > n:
                        el.value = el.value[:n]
                    if isinstance(el.struck_value, str) and len(el.struck_value) > n:
                        el.struck_value = el.struck_value[:n]


def _tag_cell_is_populated(value: object) -> bool:
    if value is None:
        return False
    if isinstance(value, list):
        return any(str(item).strip() for item in value)
    if isinstance(value, str):
        return bool(value.strip())
    return bool(str(value).strip())


def strip_loaded_tags(rows: list[RowStd] | None) -> int:
    """Blank TAGS on position rows in memory. Does not rewrite Excel or cache.

    Args:
        rows: Loaded RFP / MTO / VO / packing rows, or None.

    Returns:
        Number of position rows that had a non-empty TAGS value or struck value.
    """
    if not rows:
        return 0
    cleared = 0
    for row in rows:
        if getattr(row, "row_type", None) != RowType.position_row:
            continue
        el = row.el.get(TAGS)
        if el is None:
            continue
        had_tags = _tag_cell_is_populated(el.value) or _tag_cell_is_populated(
            el.struck_value
        )
        el.value = ""
        el.struck_value = ""
        if had_tags:
            cleared += 1
    return cleared


def strip_loaded_tags_grouped(
    grouped: dict[str, list[RowStd]] | None,
) -> int:
    """Blank TAGS on every list in a title_system → rows mapping.

    Args:
        grouped: MTO or VO dict from step2/step3, or None.

    Returns:
        Total position rows that had tags across all groups.
    """
    if not grouped:
        return 0
    return sum(strip_loaded_tags(rows) for rows in grouped.values())


def apply_load_tags_mode(
    *,
    load_tags: bool,
    rfp_rows: list[RowStd] | None,
    mto_data: dict[str, list[RowStd]] | None,
    vo_data: dict[str, list[RowStd]] | None,
    packing_rows: list[RowStd] | None,
) -> dict[str, int]:
    """Honor ``load_tags``: when false, blank TAGS before gate/split/match.

    Cache pickle and TSD packing Excel are not rewritten. The RFP parts net
    for ``load_tags=false`` is a sibling file ``rfp_parts_net_no_tags.xlsx``,
    not this in-memory strip. Default ``load_tags=true`` is a no-op.

    Args:
        load_tags: True keeps tags as loaded. False blanks RFP/MTO/VO/UL TAGS.
        rfp_rows: Raw RFP rows after ``step1_load_rfp_raw``.
        mto_data: MTO rows by title_system.
        vo_data: VO rows by title_system.
        packing_rows: In-memory packing dataset rows, or None.

    Returns:
        Counts of cleared rows per contour (``rfp`` / ``mto`` / ``vo`` / ``ul``).
        All zeros when ``load_tags`` is true.
    """
    if load_tags:
        return {"rfp": 0, "mto": 0, "vo": 0, "ul": 0}
    counts = {
        "rfp": strip_loaded_tags(rfp_rows),
        "mto": strip_loaded_tags_grouped(mto_data),
        "vo": strip_loaded_tags_grouped(vo_data),
        "ul": strip_loaded_tags(packing_rows),
    }
    print(
        "Теги отключены (load_tags=false): очищены "
        f"RFP={counts['rfp']}, MTO={counts['mto']}, VO={counts['vo']}, "
        f"УЛ={counts['ul']}. Сопоставление по коду, без раскладки "
        "VALUES=1 на несколько тегов."
    )
    return counts
