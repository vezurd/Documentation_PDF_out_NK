import re
import os
import codecs
from pathlib import Path

from prettytable import PrettyTable

from utils.file_name_converts import parse_agcc_mto_xlsx_revision_for_chain

list_of_BBB = ["BOM", "BOE", "BOQ", "MTO"]
list_of_text_doc_types = ["OD", "CJ", "VO"]


def check_mto_bbb_page(doc_type):
    if doc_type in list_of_BBB:
        return True
    else:
        return False


def getOdStyleFileName(fileName, mode="Normal"):
    #Для примера AGCC.287-7417-SOS.MTO-0001_01_RU.pdf - ищем чать "AGCC.287-7417-SOS.MTO-0001"
    #Находим положение в строке второго знака "-"
    try:
        if type(fileName) == list:
            fileName = fileName[0]
        stat_index = 0
        end_index = fileName.find("_")

        output_text = fileName[stat_index: end_index]

        #  AGCC.287-7560-SOT.WIR-0005.10 -> AGCC.287-7560-SOT.WIR-0005
        if mode == "Clear":
            reg = r'\d{4}\.\d{1,2}'  # 1.2; 11.11
            match_long = re.findall(reg, output_text)
            if match_long:
                output_text = output_text[0:len(output_text) - len(match_long[0]) + 4]
                # print(output_text, match_long)

        return output_text
    except:
        return "<No Found>"


def getRevisionFromFileName(fileName):
    #Для примера AGCC.287-7417-SOS.MTO-0001_01_RU.pdf - ищем чать "01"
    #Находим положение в строке второго знака "-"
    try:
        if type(fileName) == list:
            fileName = fileName[0]
        stat_index = fileName.find("_") + 1
        end_index = fileName.find("_", stat_index)

        output_text = fileName[stat_index: end_index]
        return output_text
    except:
        return "<No Found>"


def parse_mto_xlsx_revision_for_chain(file_name: str) -> tuple[str, str]:
    """MTO chain mode: revision + appendix via project rules in ``file_name_converts``.

    Делегирует :func:`utils.file_name_converts.parse_agcc_mto_xlsx_revision_for_chain`
    (разбор ``_01_RU.xlsx`` / ``_01-AN01_RU.xlsx`` как у ``AgccFilenamePatterns``).

    Does not alter ``getRevisionFromFileName``.
    """
    return parse_agcc_mto_xlsx_revision_for_chain(file_name)


#Извлекаем МАРКУ (SOS, SKUD, SOT) титула из штампа или имени файла
def getMarkaFromFileName(fileName):
    #Для примера AGCC.287-7417-SOS.MTO-0001_01_RU.pdf - ищем часть "SOS"
    #Находим положение в строке второго знака "-"
    try:
        if type(fileName) == list:
            fileName = fileName[0]
        stat_index = fileName.find("-")
        stat_index = fileName.find("-", stat_index + 1) + 1

        # Находим положение в строке второго знака "."
        end_index = fileName.find(".")
        end_index = fileName.find(".", end_index + 1)

        output_text = fileName[stat_index: end_index]
        return output_text
    except:
        return "<No Found>"


def getDocTypeFromFile(fileName):
    #Для примера AGCC.287-7417-SOS.MTO-0001_01_RU.pdf - ищем чать "MTO"
    #Находим положение в строке второго знака "-"
    stat_index = fileName.find(".")
    stat_index = fileName.find(".", stat_index + 1) + 1

    line_to_check = fileName[stat_index:stat_index + 1]
    is_digit_present = any(character.isdigit() for character in line_to_check)
    if is_digit_present:
        stat_index = fileName.find(".", stat_index + 1) + 1

    # Находим положение в строке второго знака "."
    end_index = fileName.find("-")
    end_index = fileName.find("-", end_index + 1)
    end_index = fileName.find("-", end_index + 1)

    output_text = fileName[stat_index: end_index]
    return output_text


def getDocNumber(fileName):
    # Для примера AGCC.287-7417-SOS.MTO-0001_01_RU.pdf - ищем часть "0001"
    # Находим положение в строке второго знака "-"
    stat_index = fileName.find("-")
    stat_index = fileName.find("-", stat_index + 1)
    stat_index = fileName.find("-", stat_index + 1) + 1

    end_index = stat_index + 4

    output_text = fileName[stat_index: end_index]
    return output_text


def getDocTitle(fileName):
    #Для примера AGCC.287-7417-SOS.MTO-0001_01_RU.pdf - ищем чать "7417"
    #Находим положение в строке второго знака "-"
    stat_index = fileName.find("-")
    end_index = fileName.find("-", stat_index + 1)

    output_text = fileName[stat_index + 1: end_index]
    return output_text


def string_Remove_New_Lines(input_text):
    if isinstance(input_text, list):
        input_text = input_text[0]
    output_text = input_text
    output_text = re.sub(r"\n", '', output_text)
    output_text = output_text.strip()
    return output_text


def string_Remove_New_Lines_With_Space(input_text):
    output_text = ""

    if "\n" in input_text:
        t = input_text.split("\n")
        for i in range(len(t)):
            v = t[i].strip()
            output_text = output_text + v + " "
    else:
        output_text = input_text

    output_text = output_text.strip()
    return output_text


def string_Junk_Cleaner(input_text, flag=0):
    ch = codecs.decode(b"\xe2\x80\x90", 'UTF-8')
    if type(input_text) != list:
        l = input_text.split("\n")
        for i in range(len(l)):
            l[i] = l[i].replace(u'\xa0', u' ')
            l[i] = l[i].replace('°', u' ')
            l[i] = l[i].replace(ch, "-")
            l[i] = l[i].replace("‐", "-")
    else:
        l = input_text.replace(u'\xa0', u' ')
        l = l.replace(u'°', u' ')
        l = l.replace(ch, "-")
        l = l.replace("‐", "-")

    names = ['Дата', 'Date', "Рев.", "Rev.", "Назначение выпуска", "Ревизия",
             "Purpose of Issue", "ООО \"Би.Си.Си.\"", "Примечание", "Стадия", "АО \"НИПИГАЗ\"",
             "В связи со значительными изменениями весь документ переделан"]

    if flag == 1:
        names.append("Лист")
        names.append("Листов")

    indexes = []
    for i in l:
        for n in names:
            if n in i or i == "" or i == " ":
                indexes.append(l.index(i))

    res = []
    for i in range(0, len(l)):
        if i not in indexes:
            res.append(l[i].strip())

    return res


def list_6_1_extrction(input_text, docType="LAY", pageNum=-1):
    # Обработка и фильтрация для первого листа
    if docType not in list_of_BBB or pageNum < 2:
        res = get_6_1_att_reg_exp(input_text)
    else:
        # Если это второй лист и далее
        l = string_Junk_Cleaner(input_text)
        names = ["AGCC"]

        indexes = []
        for i in l:
            for n in names:
                if n in i:
                    indexes.append(l.index(i))

        res = []
        for i in range(0, len(l)):
            if i not in indexes:
                res.append(l[i])

        sting_f = str(res)
        start_index = sting_f.find("л.")
        end_index = sting_f.find("из", start_index + 1)
        res = sting_f[start_index + 2:end_index].strip()
        res = res.replace(" ", "")

    return res


def list_6_2_extrction(input_text, docType="LAY", pageNum=-1):
    if docType not in list_of_BBB or pageNum < 2:
        l = string_Junk_Cleaner(input_text)
        names = ["на", "На"]

        indexes = []
        for i in l:
            for n in names:
                if n in i:
                    indexes.append(l.index(i))

        res = []
        for i in range(0, len(l)):
            if i in indexes:
                res.append(l[i])
        sting_f = str(res).lower()
        start_index = sting_f.find("на")
        end_index = sting_f.find("л.", start_index + 1)

        res = sting_f[start_index + 2:end_index].strip()
        #res = res.replace(u'\xa0', u' ')
    else:
        # if pageNum>1: print(f"\t\t\t\t 6_2: {input_text}")
        l = string_Junk_Cleaner(input_text)
        # if pageNum>1: print(f"\t\t\t\t 6_2 l: {l}")
        names = ["AGCC"]

        indexes = []
        for i in l:
            for n in names:
                if n in i:
                    indexes.append(l.index(i))

        res = []
        for i in range(0, len(l)):
            if i not in indexes:
                res.append(l[i])

        sting_f = str(res)
        start_index = sting_f.find("из")
        end_index = sting_f.find("листов", start_index + 1)
        res = sting_f[start_index + 2:end_index].strip()

        #res = res.replace(u'\xa0', u' ')
    res = res.replace("\'", "")
    return res


def list_18_1_extrction(input_text):
    if "Кабели одного типа" in input_text:
        return ""
    l = string_Junk_Cleaner(input_text)
    #print(f"\nClean Junk: {l}")

    res = []
    for i in l:
        res.append(i.split())
    l = list_flatter(res)

    #print(f"\nClean Junk 2: {l}")

    names = ["IFC", "IFR", "IFU"]

    indexes = []
    for i in l:
        for n in names:
            if n in i or len(i.split('.')) == 3:
                indexes.append(l.index(i))

    res = []
    for i in range(0, len(l)):
        if i not in indexes:
            res.append(l[i])
    if not res:
        return res

    return res[0]


def list_18_2_extrction(input_text, flag=0):
    if "Кабели одного типа" in input_text:
        return ""
    try:
        l = string_Junk_Cleaner(input_text)
        a = str(l)
        a = a.replace(" ", "")
        if flag == 10:
            # Для 10 поля
            res = re.sub(r"[/\\.]", ".", re.search(r"(\d\d.\d\d.\d\d)", a).group(1))
        elif flag == 18:
            # Для 18 поля
            res = re.sub(r"[/\\.]", ".", re.search(r"(\d\d.\d\d.\d\d\d\d)", a).group(1))
    except:
        return ""
    return res


def list_18_3_extrction(input_text):
    l = string_Junk_Cleaner(input_text)
    a = str(l)
    a = a.strip()
    a = re.sub(r"[\][\']", "", a)

    if a is None:
        return a

    data_junk = re.sub(r"[/\\.]", ".", a)
    data_junk = re.search(r"(\d+.*?\d+.*?\d+)", data_junk)
    if data_junk:
        data_junk = data_junk.group(1)

    if data_junk is not None and data_junk in a:
        stat_index = a.find(data_junk)
        result = a[stat_index + len(data_junk):len(a)].strip()
    else:
        result = a

    items = []
    items = result.split(",")
    result = items[0].strip()

    return result


def remove_dates(input_text):
    return re.sub(r"(\d\d.\d\d.\d{2,4})", "", input_text).strip()


def list_flatter(input_list):
    out_list = []
    for x in input_list:
        out_list.extend(x if isinstance(x, list) else [x])
    return out_list


def string_remove_spaces(input_string):
    return input_string.replace(' ', '')


def get_Format_from_51(input_format, mode="std"):
    output_format = input_format

    output_format = input_format.replace("Формат", "")
    output_format = output_format.replace(",", "")
    output_format = output_format.replace("А", "A")
    output_format = output_format.replace("х", "x")
    output_format = output_format.replace("×", "x")
    output_format = output_format.replace("X", "x")
    if mode == "extra":
        output_format = output_format.replace("Size", "")
        output_format = output_format.replace("/", "")

    return output_format.strip()


def get_Format_Count(input_format):
    input_format = get_Format_from_51(input_format)
    # print(input_format)
    if input_format[0] == "A":
        return 1
    else:
        return int(input_format.split("A")[0])


def get_Format_Single(input_format):
    input_format = get_Format_from_51(str(input_format))
    if "A" not in input_format:
        return f"ERR<{input_format}>"
    if input_format[0] == "A":
        return input_format
    else:
        return "A" + input_format.split("A")[1]


def get_page_num_prefix(value):
    if "AGCC" in value:
        value = int(getDocNumber(value))
    elif "." in value:
        value = int(value.split(".")[0])
    else:
        value = int(value)
    return value


def get_page_num_suffix(value):
    if "." in value:
        value = int(value.split(".")[1])
    else:
        value = int(value)
    return value


def check_data_format_long(dd_mm_yyyy):
    import datetime
    try:
        datetime.datetime.strptime(dd_mm_yyyy, '%d.%m.%Y')
        return 'Ok'
    except Exception:
        return (f'неверный формат даты: {dd_mm_yyyy} вместо dd.mm.yyyy')


def check_data_format_short(dd_mm_yy):
    import datetime
    try:
        datetime.datetime.strptime(dd_mm_yy, '%d.%m.%y')
        return 'Ok'
    except Exception:
        return (f'неверный формат даты: {dd_mm_yy} вместо dd.mm.yy')


def compare_data(dd_mm_yy, dd_mm_yyyy):
    import datetime
    try:
        if datetime.datetime.strptime(dd_mm_yy, '%d.%m.%y') == datetime.datetime.strptime(dd_mm_yyyy, '%d.%m.%Y'):
            return True
        else:
            return False
    except Exception:
        return False


def check_pdf_suffix(file_mame):
    suff_list = (".doc", ".docx", ".dwg", ".xls", ".xlsx", ".xlsm", "xls")
    from pathlib import Path
    x = Path(file_mame).suffix.lower()
    # print(file_mame, x)
    # В случае если нет расширения - вернется честь ".CAE-0003"
    if x in suff_list:
        return "True"
    else:
        return x


def check_1_cirillic_letters(text, alphabet=set('абвгдеёжзийклмнопрстуфхцчшщъыьэюя')):
    # test - True
    # тест - False
    return alphabet.isdisjoint(text.lower())


def check_2_cirillic_letters(text):
    kirill = ('абвгдеёжзийклмнопрстуфхцчшщъыьэюя')
    find_kirill = [x.upper() for x in kirill if x in text.lower()]
    out_text = ""
    for c in text:
        if c in find_kirill:
            out_text = out_text + c
        else:
            out_text = out_text + c.lower()

    return out_text + "  (" + str(find_kirill) + ")"


def get_6_1_att_reg_exp(text):
    reg_long = r'\d{1,2}.\d{1,2}'  # 1.2; 11.11
    reg_short = r'\d{1,2}'  # 1; 09
    match_long = re.findall(reg_long, text)
    if match_long:
        return match_long
    else:
        match_short = re.findall(reg_short, text)
        if match_short:
            return match_short
        else:
            return []


def print_att_list_table(raw_list: list[list],
                         max_len=15,
                         title_row=None,
                         title="",
                         titles_print_flag=0):
    if raw_list:
        out_string = ""
        if title_row is None:
            title_row = []
        size_x = len(raw_list)
        size_y = len(raw_list[0])
        if titles_print_flag:
            out_string += f"\n  {title}\n"
            out_string += f"   print_att_list_table ({size_x}:{size_y})\n"
        # print(f"\n  {title}")
        # print(f"   print_att_list_table ({size_x}:{size_y})")
        table = PrettyTable()
        t_f_n = ["n/n"]
        if title_row:
            for v in title_row:
                t_f_n.append(v)
        else:
            for y in range(size_y):
                t_f_n.append(y)
        table.field_names = t_f_n
        table.border = 1
        table.align = "l"

        for x in range(len(raw_list)):
            s = [x]
            for y in range(len(raw_list[x])):
                s.append(str(raw_list[x][y]).replace("\n", " ")[0:max_len])
            temp_len_s = len(s)
            if temp_len_s < len(table.field_names):
                for _ in range((len(table.field_names) - temp_len_s)):
                    s.append([])
            table.add_row(s)
        out_string += str(table)
        print(out_string)

        return out_string
    else:
        return f"Список для вывода пуст [{title}] "


def cable_x_fixer(cable: str):
    cable = cable.lower()
    cable = cable.replace("х", "x")
    cable = cable.replace("×", "x")
    return cable.strip()
