
import OD_tab_reader
import re
from os import listdir
from os.path import isfile, join

import shtamp_extract_classes
import string_parsing



title_Vedomost_Osnovnogo_Komplecta = "ВЕДОМОСТЬ ДОКУМЕНТОВ ОСНОВНОГО КОМПЛЕКТА РАБОЧИХ ЧЕРТЕЖЕЙ"
titel_Vedomost_Prilagaemih_Documentov = "ВЕДОМОСТЬ ПРИЛАГАЕМЫХ ДОКУМЕНТОВ"
tab_separator = "None"

doc_Osnovnoi = "Основной комлект"
doc_Prilagaemie = "Приоагаемые документы"

debug_print_input_list = 0
debug_print_out_list = 1

class doc_OD_ATTRIBUTES:
    def __init__ (self, Doc_Title, Page_Format, Document_name, Document_Revision, doc_range):
        self.Doc_Title = Doc_Title
        self.Page_Format = Page_Format
        self.Document_name = Document_name
        self.Document_Revision = Document_Revision
        self.doc_page_count = -1
        self.doc_range = doc_range

def string_to_ATT_list(input_string):

    input_string = re.sub(r"[|]", "\n", input_string)
    input_string = input_string.replace(tab_separator, "\n")
    input_string = input_string.split("\n")

    out_list = list(filter(None, input_string))
    return out_list


def OD_table_parsing_f(pdf_path, save_to_file=1):

    # Читаем все файлы из папки (без директорий)
    onlyfiles = [f for f in listdir(pdf_path) if isfile(join(pdf_path, f))]

    # Ищем файл ОД.
    pdf_full_path = "no found"
    for next_file in onlyfiles:
        if "OD" in next_file:
            pdf_full_path = pdf_path + "\\" + next_file

    if pdf_full_path == "no found":
        print("ОД не найден")
        return 0
    pdf_path = pdf_full_path
    input_list = OD_tab_reader.OD_table_reader_f(pdf_path)
    proj_OD_list = []

    #Печатаем входные данные из ОД
    if debug_print_input_list:
        for index in range(len(input_list)):
                 print(index)
                 print(''.join(input_list[index]))

    for page_index in range(len(input_list)):
        page_text = ''.join(input_list[page_index])
        #Проверяем наличие заголовка "Ведомость документов основного комплекта рабочих чертежей"
        if title_Vedomost_Osnovnogo_Komplecta in page_text:
            #Разделяем тест на list
            page_arr= page_text.split('\n')

            #Перебираем массив строк страницы
            for index in range(len(page_arr)):
                string_arr = string_to_ATT_list(page_arr[index])
                if len(string_arr) ==0:
                    break
                if "AGCC.287" in string_arr[0]:
                    proj_OD_list.append(doc_OD_ATTRIBUTES(string_parsing.string_remove_spaces(string_arr[0]), string_arr[1], string_arr[2], string_arr[3], doc_Osnovnoi))

        # Проверяем наличие заголовка "ВЕДОМОСТЬ ПРИЛАГАЕМЫХ ДОКУМЕНТОВ"
        if titel_Vedomost_Prilagaemih_Documentov in page_text:
            # Разделяем тест на list
            page_arr = page_text.split('\n')

            # Перебираем массив строк страницы
            for index in range(len(page_arr)):
                string_arr = string_to_ATT_list(page_arr[index])
                if len(string_arr) == 0:
                    break
                if "AGCC.287" in string_arr[0]:
                    proj_OD_list.append(doc_OD_ATTRIBUTES(string_parsing.string_remove_spaces(string_arr[0]), -1, string_arr[1], string_arr[2], doc_Prilagaemie))

    #Функция для подсчета количества страниц в формате


    #Подсчитываем общее количество страниц в ОД и страниц в каждом документе
    total_page_count = 0
    for el in proj_OD_list:
        el.doc_page_count = get_paege_count(el.Page_Format)
        if el.doc_range == doc_Osnovnoi:
            total_page_count = total_page_count + el.doc_page_count

    if debug_print_out_list:
        print(pdf_path)
        from prettytable import PrettyTable

        table = PrettyTable()
        table.field_names = ["Doc_Title", "Page_Format", "Page_Count:"+str(total_page_count), "Document_name", "Document_Revision", "Комплект"]
        table.align["Doc_Title"] = "l"
        table.align["Document_name"] = "l"

        for el in proj_OD_list:
            item = []
            item.append(el.Doc_Title)
            item.append(el.Page_Format)
            item.append(el.doc_page_count)
            item.append(el.Document_name)
            item.append(el.Document_Revision)
            item.append(el.doc_range)
            table.add_row(item)
        print(table)
    #Конец основной функции OD_table_parsing_f
        if save_to_file:
            from openpyxl import Workbook
            excel_out_file_name = "__result_OD_output_"

            wb = Workbook()
            # grab the active worksheet
            ws = wb.active
            # Пишем остальные строчки с данными
            for el in proj_OD_list:
                item = []
                item.append(el.Doc_Title)
                item.append(el.Page_Format)
                item.append(el.doc_page_count)
                item.append(el.Document_name)
                item.append(el.Document_Revision)
                item.append(el.doc_range)
                ws.append(item)

            # Save the file
            wb.save(excel_out_file_name +
                    string_parsing.getDocTitle(proj_OD_list[0].Doc_Title) + "-" +
                    string_parsing.getMarkaFromFileName(proj_OD_list[0].Doc_Title) + ".xlsx")
            print("FINISH  OD EXCEL OUT")
    return proj_OD_list

def OD_table_from_file(excel_file_name):
    debug_print_console_final_output = 1

    proj_OD_list = []

    import openpyxl
    from openpyxl import Workbook

    wb = Workbook()

    wb = openpyxl.load_workbook(excel_file_name)

    # grab the active worksheet
    ws = wb.active

    #Читаем строки из файла
    for row in ws.values:
        if "AGCC" in row[0]:
            proj_OD_list.append(doc_OD_ATTRIBUTES(row[0], row[1], row[3], row[4], row[5]))

    # Подсчитываем общее количество страниц в ОД и страниц в каждом документе
    total_page_count = 0
    for el in proj_OD_list:
        el.doc_page_count = get_paege_count(el.Page_Format)
        if el.doc_range == doc_Osnovnoi:
            total_page_count = total_page_count + el.doc_page_count

    if debug_print_console_final_output:
        console_print_OD_list(proj_OD_list, excel_file_name, total_page_count)

    return proj_OD_list

def console_print_OD_list(proj_OD_list, pdf_path, total_page_count):
    print(pdf_path)
    from prettytable import PrettyTable

    table = PrettyTable()
    table.field_names = ["Doc_Title", "Page_Format", "Page_Count:" + str(total_page_count), "Document_name", "Document_Revision", "Комплект"]
    table.align["Doc_Title"] = "l"
    table.align["Document_name"] = "l"

    for el in proj_OD_list:
        item = []
        item.append(el.Doc_Title)
        item.append(el.Page_Format)
        item.append(el.doc_page_count)
        item.append(el.Document_name)
        item.append(el.Document_Revision)
        item.append(el.doc_range)
        table.add_row(item)
    print(table)

def get_paege_count(page_format):
    if page_format == -1:
        return -1
    f_letters = ["A", "А"]
    s = page_format.replace(" ", "")
    s = s.split(",")
    count = 0
    for i in range(len(s)):
        if s[i][0] in f_letters:
            count += 1
        else:
            end_index = 0
            for a in range(len(f_letters)):
                index = s[i].find(f_letters[a])
                if index > 0:
                    end_index = int(index)
            count = count + int(s[i][0:end_index])
    return count