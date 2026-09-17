from base.base_classes import *
from base.tables_columns import NUMBERS, ROW_TYPE, TAGS, NAME, ANNOTATION, IN_CABINET
from pdf_parsing_v2_engine.document import V2Document

def add_cabinet_composition_info(table_std: list[RowStd]) -> list[RowStd]:
    """
    Добавляет информацию о составе шкафов в столбец IN_CABINET.
    Определяет какие позиции относятся к каждому шкафу на основе иерархической нумерации.
    """
    cabinet_counter = 1  # Счетчик шкафов для генерации уникального тега

    for i, row in enumerate(table_std):
        # Пропускаем пустые строки
        if row.el[ROW_TYPE].value == RowType.empty_row:
            continue

        # Если нашли заголовок шкафа
        if row.el[ROW_TYPE].value == RowType.cabinet_title_row:
            cabinet_number = row.el[NUMBERS].value
            cabinet_tag = get_cabinet_tag_simple(table_std, i, cabinet_number, cabinet_counter)
            cabinet_counter += 1

            # Помечаем все позиции этого шкафа
            mark_cabinet_positions(table_std, i, cabinet_number, cabinet_tag)

    return table_std


def get_cabinet_tag_simple(table_std: list[RowStd],
                           cabinet_index: int,
                           cabinet_number: str,
                           cabinet_counter: int) -> str:
    """
    Упрощенная логика получения тега шкафа:
    1. Особый случай: если в NAME есть "Оборудование в шкафу" — извлекаем тег через get_tag
    2. Ищем в следующих строках position_row (пропуская empty_row)
    3. Берем из заголовка шкафа
    4. Генерируем дефолтный тег если не нашли
    """
    cabinet_row = table_std[cabinet_index]

    # 1. Особый случай: если в NAME есть "Оборудование в шкафу" или "Оборудование для шкафа" — извлекаем тег через get_tag
    name_value = cabinet_row.el.get(NAME)
    if name_value and name_value.value and ("Оборудование в шкафу" in str(name_value.value) or "Оборудование для шкафа" in str(name_value.value)):
        tags_from_name = get_tag(str(name_value.value))
        if tags_from_name:
            return ", ".join(tags_from_name)

    # 2. Пытаемся взять тег из следующих строк position_row (пропуская empty_row)
    for i in range(cabinet_index + 1, min(cabinet_index + 3, len(table_std))):  # Проверяем до 3 строк вперед
        next_row = table_std[i]

        # Пропускаем пустые строки
        if next_row.el[ROW_TYPE].value == RowType.empty_row:
            continue

        # Если нашли не пустую строку, но это не позиция - прекращаем поиск
        if next_row.el[ROW_TYPE].value != RowType.position_row:
            break

        # Если нашли позицию с тегами
        if (next_row.el[TAGS].value and
                isinstance(next_row.el[TAGS].value, list) and
                next_row.el[TAGS].value):
            return ", ".join(next_row.el[TAGS].value)
        else:
            #Если не нашли тега в перовой попавшейся строке с позициями то выходим из цикла перебора строк
            break

    # 3. Пытаемся взять тег из заголовка шкафа
    if cabinet_row.el[TAGS].value and len(cabinet_row.el[TAGS].value) > 0:
        return ", ".join(cabinet_row.el[TAGS].value)

    # 4. Генерируем дефолтный тег
    t_com = table_std[cabinet_index].t_com
    mto_obj = V2Document.from_file_path(t_com.file_name)
    return f"ТЕГ шкафа не найден {mto_obj.doc_Title_4d}.{mto_obj.doc_Marka} шкаф_{cabinet_counter}"


def is_position_part_of_cabinet(row: RowStd, cabinet_number: str, table_std: list[RowStd], current_index: int) -> bool:
    """
    Проверяет, относится ли позиция к шкафу с учетом возможных ошибок нумерации.
    """
    # Если есть номер - проверяем иерархию (основной случай)
    if row.el[NUMBERS].value:
        return is_hierarchically_related(row.el[NUMBERS].value, cabinet_number)

    # Если номера нет - проверяем контекст
    return is_contextually_related(table_std, current_index, cabinet_number)


def is_hierarchically_related(item_number: str, cabinet_number: str) -> bool:
    """
    Проверяет иерархическую связь с учетом возможных ошибок нумерации.
    """
    if not item_number or not cabinet_number:
        return False

    # Нормальный случай: 1.2.1 относится к 1.2
    if item_number.startswith(cabinet_number + '.'):
        return True

    # Случай ошибки нумерации: проверяем совпадение уровня иерархии
    item_parts = item_number.split('.')
    cabinet_parts = cabinet_number.split('.')


    # Если разное количество уровней - не относится
    if len(item_parts) == len(cabinet_parts) + 1:
        return True
    else:
        return False





def is_contextually_related(table_std: list[RowStd], current_index: int, cabinet_number: str) -> bool:
    """
    Проверяет принадлежность к шкафу по контексту для строк без номера.
    """
    # Проверяем предыдущие строки (до 5 строк назад)
    for i in range(max(0, current_index - 5), current_index):
        prev_row = table_std[i]
        if (prev_row.el[NUMBERS].value and
                is_hierarchically_related(prev_row.el[NUMBERS].value, cabinet_number) and
                prev_row.el[ROW_TYPE].value == RowType.position_row):
            return True

    # Проверяем следующие строки (до 5 строк вперед)
    for i in range(current_index + 1, min(current_index + 6, len(table_std))):
        next_row = table_std[i]
        if (next_row.el[NUMBERS].value and
                is_hierarchically_related(next_row.el[NUMBERS].value, cabinet_number) and
                next_row.el[ROW_TYPE].value == RowType.position_row):
            return True

    return False


def mark_cabinet_positions(table_std: list[RowStd], cabinet_index: int, cabinet_number: str, cabinet_tag: str):
    """
    Помечает все позиции, относящиеся к данному шкафу с учетом ошибок нумерации.
    """
    # Идем от заголовка шкафа вперед
    for i in range(cabinet_index + 1, len(table_std)):
        row = table_std[i]

        # Если нашли следующий шкаф, раздел или систему - прекращаем
        if row.el[ROW_TYPE].value in [RowType.cabinet_title_row, RowType.section_row, RowType.system_row]:
            break

        # Если это позиция и относится к шкафу
        if (row.el[ROW_TYPE].value == RowType.position_row and
                is_position_part_of_cabinet(row, cabinet_number, table_std, i)):
            # Убедимся что поле IN_CABINET инициализировано
            if IN_CABINET not in row.el:
                row.el[IN_CABINET] = CheckElement("")
            row.el[IN_CABINET].value = cabinet_tag

            # Дополнительно: исправляем ошибки нумерации в памяти (если нужно)
            if (row.el[NUMBERS].value and
                    not row.el[NUMBERS].value.startswith(cabinet_number + '.')):
                # Сохраняем оригинальный номер в комментарии
                original_number = row.el[NUMBERS].value
                if row.el[ANNOTATION].value:
                    row.el[ANNOTATION].value += f"; Оригинальный номер: {original_number}"
                else:
                    row.el[ANNOTATION].value = f"Оригинальный номер: {original_number}"