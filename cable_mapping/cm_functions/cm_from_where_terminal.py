import re
replace_dict = {
        "[TAG]": [r'\d{4}(?:-[A-Z]{2,3}-\d{2}){0,1}-[A-ZА-Я]-[a-zA-Zа-яА-Я]{2,5}-\d{4}(?:\((?:[EЕ][xXхХ]|IS)\)){0,1}', ""],
        "[NUM1]": [r'\d{1}', "#"],
        "[NUM2]": [r'\d{2}', "##"],
        "[NUM1-2]": [r'\d{1,2}', "##"],
    }
"""
replace_regular_from_where
    Заменяем квадратные скобки на куски регулярного выражения
    Пример:
        Порт [NUM2] -> Порт \d{2}
"""


def replace_regular_from_where(reg_q: str):

    for key, value in replace_dict.items():
        while key in reg_q:
            reg_q = reg_q.replace(key, value[0])
            # print(reg_q)
    return reg_q


"""
regular_compare
"""


def regular_compare(check_text: str, reg_q: str):
    reg_q = replace_regular_from_where(reg_q)
    try:
        match = re.findall(reg_q, check_text)
    except re.error:
        print("re.error")
        print("\n", check_text, reg_q)
        exit(0)
    return match


def get_clean_terminal_from_where_no_reg(reg_exp: str):
    for key, value in replace_dict.items():
        while key in reg_exp:
            reg_exp = reg_exp.replace(key, value[1])
            reg_exp = reg_exp.strip(",")
            reg_exp = reg_exp.strip()
    reg_exp = reg_exp.replace("\\","")
    return reg_exp
