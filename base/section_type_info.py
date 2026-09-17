"""
Заполнение колонки SECTION_TYPE для position_row строк.

Принцип: каждая position_row наследует название ближайшего предшествующего section_row.
Новые типы секций, не известные общему Excel файлу, добавляются автоматически
(target="BOE", verified=False) с уведомлением в консоль.
"""

from datetime import datetime
from typing import Dict, Any

from base.base_classes import RowStd, RowType
from base.tables_columns import NAME, SECTION_TYPE


def add_section_type_info(table_std: list[RowStd], source_file: str = "") -> list[RowStd]:
    """
    Проходит по таблице и записывает в SECTION_TYPE каждой position_row
    название ближайшей предшествующей section_row.

    Если встречается section_row с названием, которого нет в общем Excel файле,
    оно добавляется автоматически (target="BOE", verified=False) с уведомлением.

    Args:
        table_std:   список строк RowStd
        source_file: имя файла-источника (для записи в Excel конфиг)
    """
    from base.bbb_section_types_excel import (
        load_section_types, add_new_sections_if_missing, get_excel_path,
    )

    section_types: Dict[str, Any] = load_section_types()
    new_sections: Dict[str, Dict[str, Any]] = {}
    current_section_name = ""

    for row in table_std:
        row_type = row.row_type

        if row_type == RowType.head_row:
            current_section_name = ""
            continue

        if row_type == RowType.section_row:
            section_name = str(row.el[NAME].value or "").strip()
            if not section_name:
                continue
            current_section_name = section_name

            if section_name not in section_types and section_name not in new_sections:
                added_at = datetime.now().strftime("%Y-%m-%d %H:%M")
                new_sections[section_name] = {
                    "target": "BOE",
                    "verified": False,
                    "source_file": source_file,
                    "added_at": added_at,
                }
                print(
                    f'[INFO] Новый тип секции: "{section_name}" → BOE\n'
                    f'       Источник: {source_file or "неизвестен"}  |  {added_at}\n'
                    f'       Для подтверждения измените "Нет" → "Да" в столбце "Проверено":\n'
                    f'       {get_excel_path()}'
                )
            continue

        if row_type == RowType.position_row and current_section_name:
            row.el[SECTION_TYPE].value = current_section_name

    if new_sections:
        add_new_sections_if_missing(new_sections)

    return table_std
