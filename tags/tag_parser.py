import re

import utils.path
from utils import string_parsing
from prettytable import PrettyTable
from tags.tag_classes import *
from utils.error_log import ErrorLog

source_dict = {}
file_out_list = []  # Список для вывода в файл результатов
index = 0
file_out_name = ""

dbg_print_source_print_full_list = 0  # Вывести все теги на всех листах

dbg_dont_ptf = 0  # 1 - Не выводить результат сравнения в файл

dbg_add_context = 0  # Выводить весь текст в процессе парсинга тегов

global stdout11


#############################################################################################
# 1. Функция внесения данных на этапе открытия PDF файлов подсистемой "files_att_collector" #
#############################################################################################
def reset():
    global index, source_dict
    source_dict = {}
    index = 0
    file_out_list.clear()


def merge_dicts(base_dict, local_dict):
    """Мерджит *local_dict* в *base_dict* (append значений по ключам)."""
    for key, values in local_dict.items():
        if key in base_dict:
            base_dict[key].extend(values)
        else:
            base_dict[key] = list(values)


def add_context(text, doc_name, page_num, target_dict=None, *, file_path=None):
    """Parse tags from *text* into *target_dict* (or global ``source_dict``).

    Dict key is ``(path_or_doc_id, page_num, doc_name)``: *doc_name* groups sheets;
    pass *file_path* (absolute PDF) from v2 so tag analysis can open the file.
    """
    result = True
    global index, source_dict
    dest = target_dict if target_dict is not None else source_dict
    if dbg_add_context:
        print(f"(1.add_context) <{text}>\n")
    text = string_cleaner(text)
    if dbg_add_context:
        print(f"(2.add_context) <{text}>\n")
    value_list = get_tag(text)
    if len(value_list) > 0:
        index = index + 1
        loc = file_path if (file_path and str(file_path).strip()) else doc_name
        key = (loc, page_num, doc_name)
        for i in range(len(value_list)):
            value = value_list[i]
            if key in dest:
                dest[key].append(value)
            else:
                dest[key] = [value]
    return result


#############################################################################################
# 2. Анализ собранной информации по тегам                                                   #
#############################################################################################
def analyze(out_dir,
            main_doc_title="МТО",
            main_doc_signature="MTO-0001",
            file_name_mode="short",
            *,
            source_dict_override=None):
    """Run tag analysis; write legacy ``*_tag_analyze.txt`` and return structured payload."""
    from tags.tag_analysis_payload import (
        apply_text_report_path,
        build_tag_analysis_payload,
        render_tag_analysis_text_from_payload,
    )

    print("<TAG PARSER>>>>>>>>>>>>>>>>>>>>")
    file_out_list.clear()
    _src = source_dict_override if source_dict_override is not None else source_dict

    payload = build_tag_analysis_payload(
        _src,
        main_doc_title=main_doc_title,
        main_doc_signature=main_doc_signature,
        high_frequency_threshold=3,
    )
    text_body = render_tag_analysis_text_from_payload(
        payload,
        main_doc_title=main_doc_title,
        main_doc_signature=main_doc_signature,
    )
    print(text_body)

    if dbg_print_source_print_full_list:
        out_dict_dbg: dict[str, list[str]] = {}
        mdc_dbg: dict[str, list[str]] = {}
        cc_dbg: dict[str, list[str]] = {}
        for i, v in _src.items():
            short_name = i[2]
            content = v
            for x in range(len(content)):
                add_to_dict(out_dict_dbg, short_name, content[x])
                add_to_dict_doubles(mdc_dbg, short_name, content[x])
                add_to_dict_doubles(cc_dbg, f"{short_name};{i[1]}", content[x])
        print_ex(f"<---Вывод в консоль всех тегов всех листов--->")
        for i in out_dict_dbg:
            content_list = out_dict_dbg[i]
            short_title = i
            print_ex(f" кол-во тегов:{len(content_list)}, short_title {short_title}")
            print_ex(print_width_list_str(content_list))
        print_ex(f"<---Конец вывода в консоль всех тегов всех листов--->")
        print_ex()

    print("<<<<<<<<<<<<<<<<<<<<<<<<<<TAG PARSER>")

    file_out_name = str(payload.get("file_out_name") or "NO_TAGS_IN_PROJECT_MTO")
    if file_name_mode == "short":
        if file_out_name != "NO_TAGS_IN_PROJECT_MTO":
            f_short_name = string_parsing.getDocTitle(file_out_name) + "-" + string_parsing.getMarkaFromFileName(
                file_out_name)
            file_name = out_dir + f_short_name + "_tag_analyze.txt"
        else:
            file_name = out_dir + "NO_TAGS_IN_PROJECT_MTO_tag_analyze.txt"
    else:
        file_name = out_dir + "NO_TAGS_IN_PROJECT_MTO_tag_analyze.txt"

    file_out_list.clear()
    file_out_list.extend(text_body.splitlines())

    save_tags_to_file("\n".join(file_out_list), file_name, out_dir)
    apply_text_report_path(payload, file_name)
    return payload


def save_tags_to_file(text, file_name, out_dir):
    try:
        utils.path.make_dir(out_dir)
        with open(file_name, "w", encoding="utf-8") as f:
            try:
                f.write(text)
            except UnicodeEncodeError:
                text = str(text.encode(encoding="utf-8", errors="strict"))
                f.write(text)
                print(f"ERROR: UnicodeEncodeError")
    except PermissionError as err:
        print(err)
        print(f"ERROR: Не удалось сохранить теги в файл - {file_name}")


def diff_lists_MTO_B(MTO_list, B_list):
    """Tags on sheet *B_list* not matched to *MTO_list*.

    Uses ``TagClass(..., strict=False)`` so malformed tags are skipped (not listed here)
    instead of aborting the pipeline. Hidden equipment (``hide_tag``, e.g. XT) is
    still listed: tag analysis reports every valid tag.
    """
    mto_set = set(MTO_list)
    C_list = []
    for b_tag in B_list:
        b_tc = TagClass(b_tag, strict=False)
        if not b_tc.is_valid:
            continue
        if b_tag not in mto_set:
            C_list.append(b_tag)
    return C_list


def print_width_list_str(input_list, max_count=8):
    out_str = ""
    start_margin = "\t\t"
    if len(input_list) == 0:
        return -1
    count = 0
    out_str = out_str + start_margin
    for x in input_list:
        out_str = out_str + x + " ; "
        count = count + 1
        if count > max_count:
            count = 0
            out_str = out_str + "\n" + start_margin
    out_str = out_str + "\n"

    return out_str


def add_to_dict(dictionary, key, value):
    if key in dictionary:
        flag = 0
        for elem in dictionary[key]:
            if elem == value:
                flag = 1
        if flag == 0:
            dictionary[key].append(value)
    else:
        dictionary[key] = [value]


def add_to_dict_doubles(dictionary, key, value):
    if key in dictionary:
        dictionary[key].append(value)
    else:
        dictionary[key] = [value]


def get_tag(text):

    # reg_univers = r'\d{4}(?:-[A-Z]{2,3}-\d{2}){0,1}-[A-ZА-Я]-[a-zA-Zа-яА-Я]{1,5}-\d{4}(?:\((?:[EЕ][xXхХ]|IS)\)){0,1}'
    reg_univers = r'\d{4}(?:-[A-ZА-Я]{2,3}-\d{2}){0,1}-[A-ZА-Я]-[a-zA-Zа-яА-Я]{1,5}-\d{4}(?:\((?:[EЕ][xXхХ]|IS)\)){0,1}(?:-\d{2}-[a-zA-Zа-яА-Я]{2,5}-\d{3}){0,1}'
    # reg_short = r'\d{4}-[A-ZА-Я]-[a-zA-Zа-яА-Я]{2,5}-\d{4}(?:\((?:[EЕ][xXхХ]|IS)\)){0,1}'
    # reg_long = r'\d{4}-[A-Z]{2,3}-\d{2}-[A-ZА-Я]-[a-zA-Zа-яА-Я]{2,5}-\d{4}(?:\((?:[EЕ][xXхХ]|IS)\)){0,1}'

    match = re.findall(reg_univers, text)
    # match_long = re.findall(reg_long, text)
    # match = (match_short + match_long)

    # Электрические коробки (8950-NDB-007-JB-007)
    reg_em = r"\d{4}-[A-Z]{3}-\d{3}-[a-zA-Zа-яА-Я]{2,5}-\d{3}"
    
    match += re.findall(reg_em, text)

    return match


# Очищаем входные данные - убираем переносы строк, неразрывные пробелы и т.п.
def string_cleaner(input_text):
    res = input_text.replace(u'\xa0', u' ')
    res = res.replace(u'°', u' ')
    res = res.replace("\n", "")
    return res

def get_bbb_work_code(text):
    reg_univers = r'[A-ZА-Я0-9]{4}-\d{2}-\d{2}-\d{1,2}'
    match = re.findall(reg_univers, text)
    return match

def print_ex(text=""):
    print(text)
    file_out_list.append(text)

if __name__ == "__main__":
     test_string_1 = "В чем разница с BCC0002694?"
     test_string_2 = "TP00-03-01-1 EL00-01-02-6"
     t3 = "8445-GA-01-S-SX-1002-03-JB-001,8445-GA-01-S-SX-1002-03-JB-002,8445-GA-01-S-SX-1002-03-JB-003,8445-GA-01-S-SX-1002-04-JB-001,8445-GA-01-S-SX-1002-04-JB-002,8445-GA-01-S-SX-1002-04-JB-003,8950-NDB-007-JB-001, 8950-NDB-007-JB-002, 8950-NDB-007-JB-003, 8950-NDB-007-JB-004,8950-NDB-007-JB-005, 8950-NDB-007-JB-006, 8950-NDB-007-JB-007"
     print(get_bbb_work_code(test_string_1))
     print(get_bbb_work_code(test_string_2))
     print(get_bbb_work_code(t3))