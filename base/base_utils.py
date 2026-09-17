import codecs
import re


def check_code(code):
    reg = r'\w{1,6}\d{5,7}'
    if len(re.findall(reg, str(code))) > 0:
        return 1
    return -1


"""
Функция для google_sheets
"""


def str_remove_junk(input_string):
    ch = codecs.decode(b"\xe2\x80\x90", 'UTF-8')
    input_string = input_string.replace(' ', '')
    input_string = input_string.replace('\n', '')
    input_string = input_string.replace(u'\xa0', u' ')
    input_string = input_string.replace(ch, "-")
    return str(input_string)


def str_remove_n_x000D_(s: str):
    if s is not None:
        s = s.replace("_x000D_", "")
        s = s.replace("\n", "")
    return s


def tags_to_str(tag_list: list) -> str:
    tag_to_str = ""
    for el in tag_list:
        tag_to_str += el + ", \n"
    return tag_to_str.strip(", \n")

# print(check_code(r"Лоток неперфорированный ONL 200х60, длина 3000 мм, толщина металла 2,5 мм, "
#            r"горячеоцинкованная сталь методом погружения, толщина цинкового покрытия не менее 40 мкм, ON0201130025Z"))