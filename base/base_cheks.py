from base.base_classes import *
from base.base_utils import str_remove_junk
from base.tables_columns import *
from nano_cad.fields_in_tables import TableFields
from tags.tag_classes import check_tag_94s_4_2, check_tag_94s_4_4_additional_code, TagClass, GoogleSystemCode
from utils.colors import Color
import base.t_comm_initial_classes as t_com_init_cls


def _cfg_enabled(cfg: dict | None, key: str, default: bool = True) -> bool:
    if not isinstance(cfg, dict):
        return default
    return cfg.get(key, default)


def base_vs_base_std(input_base_std, code_base_std, columns_for_check, cfg=None):
    if not _cfg_enabled(cfg, "enabled", True):
        RowStd.set_color_by_row_type_in_list(input_base_std)
        return input_base_std

    # Проверка по столбцу с кодом продукции [CODE] остальных столбцов - Наименование, Парт-Номер и т.п.
    if _cfg_enabled(cfg, "check_by_code", True):
        check_by_code(input_base_std, code_base_std, columns_for_check)

    # Покраска рядов в цвет RowType
    RowStd.set_color_by_row_type_in_list(input_base_std)

    # Проверка Кодов оборудования - ARK, BZ и т.п. (НЕ) вносит изменения в данные
    if _cfg_enabled(cfg, "check_equipment_codes", True):
        check_equipment_codes(input_base_std, code_base_std)

    # Проверка на совпадение кол-ва тегов и кол-ва оборудования (НЕ) вносит изменения в данные
    if _cfg_enabled(cfg, "check_tags_value", True):
        check_tags_value(input_base_std)

    # Проверка на дробные значения для единиц измерения типа шт. и т.п.
    if _cfg_enabled(cfg, "check_value", True):
        check_value(input_base_std)

    # Проверка на дублирование тегов
    if _cfg_enabled(cfg, "check_duplicate_tags", True):
        check_duplicate_tags(input_base_std)

    # Проверка массы
    if _cfg_enabled(cfg, "check_mass", True):
        check_mass(input_base_std, code_base_std)

    # Проверка на запрещенные к применению коды оборудования
    if _cfg_enabled(cfg, "check_prohibition", True):
        check_prohibition(input_base_std, code_base_std)

    # Проверка на 94S 4.2 4.4 (SkT, EX и т.п.)
    if _cfg_enabled(cfg, "check_tags_4_2_4_4", True):
        check_tags_4_2_4_4(input_base_std)

    # Разбивка позиций с большим кол-вом тегов
    # check_long_tag_list(input_base_std)

    # Проверка нумерации позиций
    if _cfg_enabled(cfg, "check_position_numeration", True):
        check_position_numeration(input_base_std)

    return input_base_std


def base_vs_base_correction(input_base_std, code_base_std, columns_for_check, cfg=None):
    if not _cfg_enabled(cfg, "enabled", True):
        RowStd.set_color_by_row_type_in_list(input_base_std)
        return input_base_std

    # Проверка по столбцу с кодом продукции [CODE] остальных столбцов - Наименование, Парт-Номер и т.п.
    if _cfg_enabled(cfg, "check_by_code", True):
        check_by_code(
            input_base_std,
            code_base_std,
            columns_for_check,
            mode=("correction", "comment"),
            enable_comment=False,
            enable_soft_compare_color=False)

    # Проверка на дробные значения для единиц измерения типа шт. и т.п.
    if _cfg_enabled(cfg, "check_value", True):
        check_value(input_base_std)

    # Проверка на совпадение кол-ва тегов и кол-ва оборудования (НЕ) вносит изменения в данные
    if _cfg_enabled(cfg, "check_tags_value", True):
        check_tags_value(input_base_std)

    # Проверка на дублирование тегов
    if _cfg_enabled(cfg, "check_duplicate_tags", True):
        check_duplicate_tags(input_base_std)

    # Проверка на 94S 4.2 4.4 (SkT, EX и т.п.)
    if _cfg_enabled(cfg, "check_tags_4_2_4_4", True):
        check_tags_4_2_4_4(input_base_std)

    # Проверка нумерации позиций
    if _cfg_enabled(cfg, "check_position_numeration", True):
        check_position_numeration(input_base_std)

    # Покраска рядов в цвет RowType
    RowStd.set_color_by_row_type_in_list(input_base_std)

    return input_base_std


def check_by_code(input_base_std, code_base_std, columns_for_check, mode=(0, "comment"), enable_comment=True, enable_soft_compare_color=True):
    """
    mode:
        "comment" - Только комментирование
        "correction" - Корректирование значений по code_base_std
    enable_comment: если True (по умолчанию) — добавлять комментарии к ячейкам
    enable_soft_compare_color: если True (по умолчанию) — красить в жёлтый при soft_compare,
        если False — красить в бледно-зелёный
    """
    mode_comment = "comment"
    mode_correction = "correction"

    """
    msg_index:
        0 - Индекс комментариев для ГуглБазы
        1 - Индекс комментариев для МТО
    
    """
    msg_index = code_base_std[0].t_com.tabel_type_index
    msg_soft_compare_1 = ["Немного не совпадает c ГуглБазой (не критично)",
                          "Немного не совпадает c МТО (не критично)"]
    msg_soft_compare_2 = ["По ГУГЛ БАЗЕ",
                          "По MTO"]
    msg_compare_1 = ["Не совпадает c ГуглБазой (критично);",
                     "Не совпадает c MTO (критично);"]
    msg_compare_2 = ["По ГУГЛ БАЗЕ",
                     "По MTO"]

    for row_input_base in input_base_std:
        code = row_input_base.el[CODE].value
        row_code_base = RowStd.get_row_by_code(code, code_base_std)
        if isinstance(row_code_base, NoneType):
            # print(f"\t pass : {code}  (base_vs_base_std)")
            continue
        for att in columns_for_check:
            if att == TAGS:
                continue
            att_value_mto = row_input_base.el[att].value
            att_value_base = row_code_base.el[att].value

            # if code == "BCC0000545" and att == NAME:
            #     print(code)
            #     print(f"{att_value_mto} [att_value_mto]")
            #     print(f"{att_value_base} [att_value_base]")
            #     print("Soft_mode:")
            #     print(f"{soft_compare(att_value_mto)} [att_value_mto]")
            #     print(f"{soft_compare(att_value_base)} [att_value_base]")

            el = row_input_base.el[att]
            # Если совпали - красим в зеленый
            if att_value_mto == att_value_base:
                Color.set_el_color(el, Color.green)
            elif soft_compare(att_value_mto) == soft_compare(att_value_base):
                if enable_soft_compare_color:
                    Color.set_el_color(el, Color.yellow)
                else:
                    # Color.set_el_color(el, Color.match_matched)
                    Color.set_el_color(el, Color.green)
                if mode_comment in mode:
                    if enable_comment:
                        el.comment += (f"{msg_soft_compare_1[msg_index]}:\n"
                                       f"  {msg_soft_compare_2[msg_index]}:\n"
                                       f"    {att_value_base}\n")
                if mode_correction in mode:
                    row_input_base.el[att].value = row_code_base.el[att].value
                    if enable_soft_compare_color:
                        Color.set_el_color(el, Color.correction)
                        # Для корректировки БД Нанокада (14/05/2025)
                        # Проверяем есть ли данный атрибут в таблице БД
                        db_column = TableFields.get_db_column_by_table_std_name(row_input_base.el[BD_TABLE_NAME].value, att)
                        if db_column:
                            Color.set_el_color(el, Color.correction_db)
            else:
                # Если НЕ совпали - красим в красный
                if mode_comment in mode:
                    Color.set_el_color(el, Color.red)
                    if enable_comment:
                        el.comment += (f"{msg_compare_1[msg_index]}\n" +
                                       f"{el.value}\n{msg_compare_2[msg_index]}:\n{att_value_base}")
                        el.value = f"{el.value}\n{msg_compare_2[msg_index]}:\n{att_value_base}"
                if mode_correction in mode:
                    row_input_base.el[att].value = row_code_base.el[att].value
                    Color.set_el_color(el, Color.correction)
                    # Для корректировки БД Нанокада (14/05/2025)
                    # Проверяем есть ли данный атрибут в таблице БД
                    db_column = TableFields.get_db_column_by_table_std_name(row_input_base.el[BD_TABLE_NAME].value, att)
                    if db_column:
                        Color.set_el_color(el, Color.correction_db)

def check_long_tag_list(base):
    """
    Проверяем кол-во тегов в строчке
    Если больше tag_max_count разбиваем на несколько строк добавляя к нумерации дополнительный индекс
    :param base: list of RowCheck()
    """
    tag_max_count = 15

    row_index = -1
    for row in base:
        row_index += 1
        """
        Проверка относится ли тип строки к позиции
        """
        if not check_position_row(row):
            continue
        # Проверяем равно ли количество тегов значению количества позиций
        flag = tag_vs_value_sub_check(row)
        if flag is None:
            continue
        try:
            tag_curr_count = int(row.el[VALUES].value)
        except ValueError:
            continue

        if tag_curr_count > tag_max_count:
            if flag:
                curr_tag_list = row.el[TAGS].value
                chunk_section = tag_curr_count // tag_max_count
                if chunk_section == 1:
                    chunk_section += 1
                chunk_size = (tag_curr_count // chunk_section) + 1
                chunks = [curr_tag_list[i:i + chunk_size] for i in range(0, len(curr_tag_list), chunk_size)]

                print(f"chunk_size={chunk_size}; chunk_section={chunk_section};\n"
                      f"    tag_curr_count={tag_curr_count}\n"
                      f"    chunks")
                # for i in range(len(chunks)):
                #     print(chunks[i])
                #     print()

                # dic = {}
                # for k, v in column_dict.items():
                #     try:
                #         dic[v] = row.el[v].value
                #     except AttributeError:
                #         print(f"att = {v}")
                #         dic[v] = ""
                # row = ColNames.get_std_check_row(dic)
                # list_std_obj.append(row)


def tag_vs_value_sub_check(row):
    tags_count = -1
    value_count = -1
    if isinstance(row.el[TAGS].value, list) and len(row.el[TAGS].value) > 0:
        tags_count = len(row.el[TAGS].value)
    if not isinstance(row.el[VALUES].value, NoneType):
        value_count = row.el[VALUES].value
    if tags_count != -1:
        try:
            if int(tags_count) == int(value_count):
                return True
            else:
                return False
        except ValueError:
            return False
    else:
        return None


def check_equipment_codes(base, google_base_std):
    """
    Проверка Кодов оборудования. Для позиций для которых требуется тегирование ARK, BZ и т.п.
    проверяем что есть тег и он корректный
    :param google_base_std:
    :param base:
    """
    prohibition_eq = "N_N"

    for row in base:
        """
        Проверка относится ли тип строки к позиции
        """
        if not check_position_row(row):
            continue

        tags_list = []
        color = Color.green  # Базовый цвет для тестирования, показывает что ячейка проверена и все ок
        # Получаем код оборудования
        code = row.el[CODE].value
        # Ищем в гугл базе строку по коду
        google_base = RowStd.get_row_by_code(code, google_base_std)
        if google_base is None:
            print(f"\"{code}\" ERROR check_equipment_codes: google_base row code -  is NONE")
            continue

        #  Читаем значение из гугл базы столбца "Код оборудования"
        google_eq = google_base.el[EQUIPMENT_CODE].value

        # Проверяем кол-во тегов в МТО
        if isinstance(row.el[TAGS].value, list) and len(row.el[TAGS].value) > 0:
            tags_count = len(row.el[TAGS].value)
            tags_list = row.el[TAGS].value
        else:
            tags_count = 0

        # Логика
        if google_eq is not None and google_eq != "":
            if tags_count > 0:
                if google_eq == prohibition_eq:
                    row.el[TAGS].comment += f"Позиция не тегируется!\n"
                    color = Color.red
                    Color.set_el_color(row.el[TAGS], color)
                    continue
                for tag in tags_list:
                    equipment_flag = False
                    additional_code_flag = False
                    """
                    c_tag - Создаем экземпляр класса TagClass
                    Из него нам надо сравнить код оборудования и дополнительный код (Ex, IS)
                    """
                    c_tag = TagClass(tag)
                    """
                    g_tag - код из гугл базы, формат <код_оборудования(дополнительный_код>, AEX(IS)                                        
                    """
                    g_tag = GoogleSystemCode(google_eq)

                    # Сравниваем значения тега и гугл базы
                    if c_tag.equipment == g_tag.equipment:
                        equipment_flag = True
                    if c_tag.additional_code == g_tag.additional_code:
                        additional_code_flag = True

                    #
                    if not equipment_flag:
                        row.el[TAGS].comment += (f"Код оборудования не совпадает с ГуглБазой:\n"
                                                 f"    {c_tag.equipment} | в теге {c_tag.tag}\n"
                                                 f"    {g_tag.equipment} | в ГуглБазе {g_tag.tag}\n")
                        color = Color.red

                    if not additional_code_flag:
                        row.el[TAGS].comment += (f"Дополнительный код не совпадает с ГуглБазой:\n"
                                                 f"    {c_tag.additional_code} | в теге {c_tag.tag}\n"
                                                 f"    {g_tag.additional_code} | в ГуглБазе {g_tag.tag}\n")
                        color = Color.red
                    #
                    Color.set_el_color(row.el[TAGS], color)
            else:
                if google_eq != prohibition_eq:
                    equipment_flag = True
                    """
                    g_tag - код из гугл базы, формат <код_оборудования(дополнительный_код>, AEX(IS)                                        
                    """
                    g_tag = GoogleSystemCode(google_eq)

                    # Сравниваем значения тега и гугл базы
                    if g_tag.equipment is not None:
                        equipment_flag = False

                    #
                    if not equipment_flag:
                        row.el[TAGS].comment += (f"Необходимо оттегировать оборудование:\n"
                                                 f"    в ГуглБазе \"{g_tag.tag}\"\n")
                        color = Color.red
                        Color.set_el_color(row.el[TAGS], color)
        elif google_eq is not None and google_eq == "" and tags_count > 0:
            row.el[TAGS].comment += f"Добавить тег в гугл базу;\n"
            color = Color.violet
            Color.set_el_color(row.el[TAGS], color)
            #


def check_tags_value(base: list[RowStd]):
    """
    Сравнение количества тегов у позиции с количеством материала
    :param base: list of RowCheck()
    """
    for row in base:

        # Проверяем равно ли количество тегов значению количества позиций
        flag = tag_vs_value_sub_check(row)
        if flag is None:
            continue

        if flag:
            color = Color.green
        else:
            color = Color.red

        row.el[VALUES].set_color_safe(color)
        row.el[TAGS].set_color_safe(color)
        if color == Color.red:
            row.el[VALUES].comment += f"Не совпадает кол-во тегов и кол-во оборудования;\n"
            row.el[TAGS].comment += f"Не совпадает кол-во тегов и кол-во оборудования;\n"


def check_value(base):
    """
        Проверка на дробные значения в поле value - если шт. и дробное значение - ошибка
        :param base: list of RowCheck()
        """
    units_list = ["шт", "упак", "компл"]
    color = Color.red
    for row in base:
        """
        Проверка относится ли тип строки к позиции
        """
        if not check_position_row(row):
            continue

        el = row.el[VALUES]
        if not isinstance(el.value, NoneType):
            value = el.value
            unit = row.el[UNITS].value
            if unit in units_list:
                flag = False
                try:
                    if not int(value) == float(value):  # Проверяем на дробную часть
                        flag = True
                except ValueError:
                    flag = True
                if flag:
                    Color.set_el_color(el, color)
                    el.comment += f"Запрещена дробная часть для ед.изм.: {unit}; \n"
                else:
                    Color.set_el_color(el, Color.green)


def check_duplicate_tags(base):
    """
    Проверка на дублирование тегов в списке тегов по всем строкам base
    :param base: list of RowCheck()
    """
    color = Color.red
    # Словарь для хранения информации о тегах: tag -> список строк, где он встречается
    tag_to_rows = {}
    
    # Первый проход: собираем все теги из всех строк
    for row in base:
        """
        Проверка относится ли тип строки к позиции
        """
        if not check_position_row(row):
            continue

        # Проверяем что теги хранятся как список
        if isinstance(row.el[TAGS].value, list) and len(row.el[TAGS].value) > 0:
            tags_list = row.el[TAGS].value
            for tag in tags_list:
                if tag not in tag_to_rows:
                    tag_to_rows[tag] = []
                tag_to_rows[tag].append(row)
    
    # Находим теги, которые встречаются более одного раза (дубликаты)
    duplicate_tags = {tag: rows for tag, rows in tag_to_rows.items() if len(rows) > 1}
    
    # Второй проход: подсвечиваем все строки с дублирующимися тегами
    for duplicate_tag, rows_with_tag in duplicate_tags.items():
        for row in rows_with_tag:
            Color.set_el_color(row.el[TAGS], color)
            # Формируем список кодов позиций, где встречается этот тег
            codes_with_tag = [r.el[CODE].value for r in rows_with_tag]
            codes_str = ", ".join(str(code) for code in codes_with_tag)
            row.el[TAGS].comment += f"Обнаружено дублирование тега {duplicate_tag} в позициях: {codes_str};\n"


def check_prohibition(base, google_base_std):
    """
    Проверка на запрещенные к применению коды оборудования (столбец K в ГуглТаблице - Группа)
    :param google_base_std:
    :param base:
    """
    g_base_t_com = google_base_std[0].t_com  # -> TableComments

    color = Color.red
    for row in base:
        """
        Проверка относится ли тип строки к позиции
        """
        if not check_position_row(row):
            continue
        if g_base_t_com.tabel_type != t_com_init_cls.GoogleBase.tabel_type:
            continue
        code = row.el[CODE].value
        google_base = RowStd.get_row_by_code(code, google_base_std)
        if isinstance(google_base, NoneType):  # Если код не найден - присваиваем текст "Не найден в гугл базе..."
            prohibition = PROHIBITION_LIST[1]
        else:
            prohibition = google_base.el[PROHIBITION].value  # Если код есть - значение поля К
        if prohibition in PROHIBITION_LIST:
            row.el[CODE].color = color
            row.el[CODE].comment += f"{prohibition}\n"
            row.el[ANNOTATION].color = color
            row.el[ANNOTATION].value = prohibition

            if not isinstance(google_base, NoneType) and prohibition != PROHIBITION_LIST[1]:
                if (not isinstance(google_base.el[REPLACEMENT].value, NoneType)
                        and google_base.el[REPLACEMENT].value != ""):
                    # Ищем в столбце Q гугл базы ("Заменен на:")
                    code = str(google_base.el[REPLACEMENT].value).strip()
                    # Находим строку с кодом для замены
                    google_base = RowStd.get_row_by_code(code, google_base_std)
                    if google_base is None:
                        print(f"\"{code}\" ERROR check_prohibition: google_base row code - is NONE")

                    # Указываем описание строки для замены
                    if row.el[NAME].value != google_base.el[NAME].value:
                        row.el[NAME].value = (f"{row.el[NAME].value}\n"
                                              f"Замена по ГУГЛ БАЗЕ:\n"
                                              f"{google_base.el[NAME].value}")
                        row.el[NAME].color = color
                    if row.el[TYPE_MARK].value != google_base.el[TYPE_MARK].value:
                        row.el[TYPE_MARK].value = (f"{row.el[TYPE_MARK].value}\n"
                                                   f"Замена по ГУГЛ БАЗЕ:\n"
                                                   f"{google_base.el[TYPE_MARK].value}")
                        row.el[TYPE_MARK].color = color
                    if row.el[VENDOR].value != google_base.el[VENDOR].value:
                        row.el[VENDOR].value = (f"{row.el[VENDOR].value}\n"
                                                f"Замена по ГУГЛ БАЗЕ:"
                                                f"\n{google_base.el[VENDOR].value}")
                        row.el[VENDOR].color = color
                    if row.el[UNITS].value != google_base.el[UNITS].value:
                        row.el[UNITS].value = (f"{row.el[UNITS].value}\n"
                                               f"Замена по ГУГЛ БАЗЕ:"
                                               f"\n{google_base.el[UNITS].value}")
                        row.el[UNITS].color = color
                    if row.el[MASS].value != google_base.el[MASS].value:
                        row.el[MASS].value = (f"{row.el[MASS].value}\n"
                                              f"Зам.:"
                                              f"\n{google_base.el[MASS].value}")
                        row.el[MASS].color = color
                    row.el[CODE].value = (f"{row.el[CODE].value}\n"
                                          f"Замена по ГУГЛ БАЗЕ:"
                                          f"\n{google_base.el[CODE].value}")


def check_mass(base, google_base_std):
    """
    Проверка совпадения массы с ГуглТаблицей
    :param google_base_std:
    :param base:
    """
    color = Color.red
    for row in base:
        """
        Проверка относится ли тип строки к позиции
        """
        if not check_position_row(row):
            continue

        code = row.el[CODE].value
        element_mass = row.el[MASS].value
        google_base = RowStd.get_row_by_code(code, google_base_std)
        google_mass = ""
        if not isinstance(google_base, NoneType):
            google_mass = google_base.el[MASS].value

        google_mass = str(google_mass.replace(".", ","))
        element_mass = str(element_mass.replace(".", ","))

        if str(google_mass) == str(0):
            google_mass = ""
        if str(element_mass) == str(0):
            element_mass = ""

        if not (str(google_mass) == str(element_mass)):
            row.el[MASS].value = f"{row.el[MASS].value}\n({google_mass})"
            row.el[MASS].color = color
            row.el[MASS].comment += f"Масса позиции <{element_mass}> не совпадает с ГуглБазой массой <{google_mass}>;\n"


def check_tags_4_2_4_4(base):
    """
    Проверка тегов на приложения 4.2 и 4.4
    :param base:
    """
    color = Color.red

    for row in base:
        if isinstance(row.el[TAGS].value, list) and len(row.el[TAGS].value) > 0:
            tags_list = row.el[TAGS].value
            for tag in tags_list:
                if check_tag_94s_4_2(tag) and check_tag_94s_4_4_additional_code(tag):
                    pass
                else:
                    row.el[TAGS].color = color
                    row.el[TAGS].comment += f"Некорректный тег {tag} по процедуре 94S; \n"


def soft_compare(text: str):
    text = str_remove_junk(text)
    # check_text = check_text.lower()
    text = text.replace("х", "x")
    text = text.replace("×", "x")
    # check_text = check_text.replace("°", "")
    # check_text = check_text.replace("\'", "")
    # check_text = check_text.replace("\"", "")
    # check_text = check_text.replace("\n", "")
    # check_text = check_text.replace(" ", "")  # неразрывный пробел?
    # check_text = check_text.strip(",0")
    # check_text = check_text.strip(".")
    return text.strip()


def check_position_row(row):
    if row.el[ROW_TYPE].value != RowType.position_row:
        return False
    return True


def check_position_numeration(base: list):
    """
    Проверка нумерации во втором столбце
    """
    color = Color.red
    checked_row_list = (RowType.section_row, RowType.cabinet_title_row, RowType.position_row)

    prev_number = 0

    def subtraction(a, b):
        a = str(a).split(".")
        b = str(b).split(".")

        dif = len(a) - len(b)
        #
        # a == 1.3.2 ; b= 1.3.1
        try:
            if dif == 0:
                return int(a[-1]) - int(b[-1])

            # a == 1.3.1 ; b == 1.3
            if dif > 0:
                return int(a[len(a) - dif - 1]) - int(b[-1])
            if dif < 0:
                return int(a[-1]) - int(b[len(b) - abs(dif) - 1])
        except ValueError:
            return 200

        return 100

    for row in base:
        """
        Проверка относится ли тип к перечню нумеруемых строк
        """
        if row.el[ROW_TYPE].value in checked_row_list:

            curr_number = row.el[NUMBERS].value
            res = subtraction(curr_number, prev_number)
            if abs(res) > 1:
                # print(f"a={curr_number} ; b={prev_number}; res={res}")
                row.el[NUMBERS].color = color
                row.el[NUMBERS].comment += (f"Пропущена нумерация:\n"
                                            f"    <{curr_number}> | ткущее значение\n"
                                            f"    <{prev_number}> | предыдущее значение ;\n")

            prev_number = curr_number
