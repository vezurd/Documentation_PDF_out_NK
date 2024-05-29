

#Общий рубильник дебага

debug_f1 = 1

dubug_only_fitz_opening = 1

debug_print_console_final_output = 1
debug_print_only_title_pages = 0
debug_print_excel_final_output = 1

debug_print_files_list = 1

#Не выводить результаты в консоль
debug_no_output = 0
debug_draw_borders = 1

# !! Рисует все элементы страницы
debug_draw_all_elements = 0

#Рисует рамку (для оперделения колонтитулов)
debug_draw_biggest_element = 1

#Рисует область штампа
debug_draw_shtamp = 1

#Рисует рамки элементов относящихся к штампу
debug_draw_shtamp_elements = 1

#Рисует рамки элементов из element_coordinates через функцию "debug_f1_function"
debug_draw_atributes = 1

#Вывод информации в консоль для функции "debug_f1_function"
debug_flag_print_console = 0



# Для считывания PDF
import PyPDF2
import elements_coordinates
import openpyxl
import shtamp_extract_classes
# Для анализа структуры PDF и извлечения текста
from pdfminer.high_level import extract_pages, extract_text
from pdfminer.layout import LTTextContainer, LTChar, LTRect, LTFigure
# Для извлечения текста из таблиц в PDF
import pdfplumber
# Для извлечения изображений из PDF
from PIL import Image
from pdf2image import convert_from_path
# Для выполнения OCR, чтобы извлекать тексты из изображений 
import pytesseract 
# Для удаления дополнительно созданных файлов
import os
from shtamp_extract_classes import *
from elements_coordinates import *

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
import fitz

from shtamp_extract_classes import Current_Project, doc_ATTRIBUTES, page_SHTAMP_ATTRIBUTES

from PDF_check_proc import extract_table, table_converter, text_extraction, check_element_position, pdf_get_text_from_file, crop_image

from os import listdir
from os.path import isfile, join

from elements_coordinates import Element_Shtamp, File_Name_Shtamp_50
import string_parsing

#Осносной цикл поиска атрибутов в заданной области
# temp_element_cooridinates - класс из elements_coordinates.py
# temp_shtamp_extract_class - название класса из shtamp_extract_classes.py

def check_MTO_BBB_page():
    if document.doc_Type == "MTO" or document.doc_Type == "BOM" or document.doc_Type == "BOE" or document.doc_Type == "BOQ":
        return True
    else:
        return False
def find_attribute_func(temp_element_cooridinates,temp_shtamp_extract_class):

    if check_element_position(temp_element_cooridinates(page_tables, document.doc_Type, cur_page.page_num),element):
        # Это текст?
        if isinstance(element, LTTextContainer):
            if debug_f1 and debug_draw_shtamp_elements:
                page_fitz.draw_rect([element.x0, page_tables.height - element.y1, element.x1, page_tables.height - element.y0], color=(0, 1, 0), width=2)
            check_text = element.get_text()
            if temp_element_cooridinates.Check_Text(check_text):
                check_text = temp_element_cooridinates.Clean_Text(check_text,document.doc_Type, cur_page.page_num)
                cur_page.dict_attributes[temp_shtamp_extract_class].append(check_text)
                cur_page.dict_attributes[temp_shtamp_extract_class] = string_parsing.list_flatter(cur_page.dict_attributes[temp_shtamp_extract_class])
#
def debug_f1_function (element_1, print_text, debug_flag, draw_color=(0, 1, 0)):
        if debug_f1 and debug_flag:
            if print_text == 0:
                print_text=str(element_1.__class__.__name__)
            if debug_flag_print_console:
                print("<START DEBUG:\t" + print_text + ">")
                print(element_1.x0, element_1.y0, element_1.x1, element_1.y1, "\t<page_widh:" + str(page_tables.width) + ", page_height:" + str(page_tables.height) + ">")
                print("<END DEBUG:\t" + print_text + ">")
            page_fitz.draw_rect([element_1.x0 - 1, page_tables.height - element_1.y1 - 1, element_1.x1 + 1, page_tables.height - element_1.y0 - 1], color=draw_color, width=2)



from elements_coordinates import scale_number

from string_parsing import list_of_BBB


class biggest_element:
    x0 = 0
    y0 = 0
    x1 = 0
    y1 = 0


#Счетчик страниц в документе
page_counter = 0

#Массив для хранения всех Документов (файлов) текщуего проекта
Curr_Proj = []

import time
start_time = time.time() ## точка отсчета времени

# Находим путь к PDF
pdf_path = r'test_Project'

#Читаем все файлы из папки (без директорий)
onlyfiles = [f for f in listdir(pdf_path) if isfile(join(pdf_path, f))]

#Заносим все пути файлов и создаем массив документов - уникальное свойство - File_Full_Name (без пути).
for next_file in onlyfiles:
    Curr_Proj.append(doc_ATTRIBUTES(pdf_path+"\\"+next_file))


############################
if debug_print_files_list:
    for document in Curr_Proj:
        print(document.File_Full_Path, "\tFile Full Name:", document.File_Full_Name)
    print("_________________________________________________________________________")

#Цикл перебора всех документов (PDF файлов)
for document in Curr_Proj:

    if debug_f1 or dubug_only_fitz_opening:
        fitz_doc = fitz.open(document.File_Full_Path)


    # создаём объект файла PDF
    pdfFileObj = open(document.File_Full_Path, 'rb')

    # создаём объект считывателя PDF
    pdfReaded = PyPDF2.PdfReader(pdfFileObj)

    # Создаём словарь для извлечения текста из каждого изображения
    text_per_page = {}
    page_counter =0

    # Извлекаем страницы из PDF
    for pagenum, page in enumerate(extract_pages(document.File_Full_Path)):
        #Добавляем страницу в документ
        page_counter = page_counter+1

        cur_page = page_SHTAMP_ATTRIBUTES(page_counter)


        # Инициализируем переменные, необходимые для извлечения текста со страницы
        pageObj = pdfReaded.pages[pagenum]
        page_text = []
        line_format = []
        text_from_images = []
        text_from_tables = []
        page_content = []

        # Инициализируем количество исследованных таблиц
        table_num = 0
        first_element = True
        table_extraction_flag = False

        # Открываем файл pdf
        pdf = pdfplumber.open(document.File_Full_Path)
        # Находим исследуемую страницу
        page_tables = pdf.pages[pagenum]
        # Находим количество таблиц на странице
        #tables = page_tables.find_tables()

        # Находим все элементы
        page_elements = [(element.y1, element) for element in page._objs]

        # Сортируем все элементы по порядку нахождения на странице
        page_elements.sort(key=lambda a: a[0], reverse=True)

        # Open the pdf
        if debug_f1 or dubug_only_fitz_opening:
            page_fitz = fitz_doc[pagenum]
            page_fitz.clean_contents()


        # Находим границы рамки текущей страницы
        biggest_element.x0 = 0
        biggest_element.y0 = 0
        biggest_element.x1 = 0
        biggest_element.y1 = 0

        for i, component in enumerate(page_elements):
            element = component[1]

            if (element.x1 < (page_tables.width-elements_coordinates.scale_number) and element.y1 < (page_tables.height-elements_coordinates.scale_number)):

                biggest_element_width = biggest_element.x1 - biggest_element.x0
                biggest_element_height = biggest_element.y1 - biggest_element.y0

                curr_element_width = element.x1 - element.x0
                curr_element_height = element.y1 - element.y0

                if curr_element_width > biggest_element_width:
                    biggest_element.x0 = element.x0
                    biggest_element.x1 = element.x1

                if curr_element_height > biggest_element_height:
                    biggest_element.y0 = element.y0
                    biggest_element.y1 = element.y1

        elements_coordinates.border_bottom = int((biggest_element.y0/scale_number)+1)
        elements_coordinates.border_left = int(((page_tables.width-biggest_element.x1)/scale_number)+1)
        elements_coordinates.border_top = int((page_tables.height-biggest_element.y1)/scale_number)-1

        if debug_draw_borders :
            print(document.File_Full_Name, ", стр.", cur_page.page_num,
                  "Нижний колонтиул:",elements_coordinates.border_bottom,
                  "\tПравый колонтитул:",elements_coordinates.border_left,
                  "\tВерхний колонтитул:",elements_coordinates.border_top)

        element_shtamp = elements_coordinates.Element_Shtamp(page,document.doc_Type)

        # Рисуем рамку листа
        debug_f1_function(biggest_element,"biggest_element", debug_draw_biggest_element, (0,1,1))

        #Рисуем рамку штампа
        if not check_MTO_BBB_page or cur_page.page_num < 2:
            debug_f1_function(elements_coordinates.Element_Shtamp(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_shtamp, (0, 0, 1))

        # Рисуем рамки атрибутов
        debug_f1_function(elements_coordinates.DOC_TITLE_1(page_tables, document.doc_Type, cur_page.page_num), 0,                   debug_draw_atributes, (1, 0, 0))
        if not check_MTO_BBB_page or cur_page.page_num < 2:
            debug_f1_function(elements_coordinates.Current_Document_Revision_26(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))
        debug_f1_function(elements_coordinates.File_Name_Shtamp_50(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))
        if not check_MTO_BBB_page or cur_page.page_num < 2:
            debug_f1_function(elements_coordinates.Page_Format_51(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))
        debug_f1_function(elements_coordinates.Sheet_number_Current_sheet_number_6_1(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))
        debug_f1_function(elements_coordinates.Sheet_number_Current_sheet_number_6_2(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))

        if cur_page.page_num < 2:
            debug_f1_function(elements_coordinates.Construction_facility_name_2(page_tables, document.doc_Type, cur_page.page_num), 0,  debug_draw_atributes, (1, 0, 0))
            debug_f1_function(elements_coordinates.Unit_title_name_3(page_tables, document.doc_Type, cur_page.page_num), 0,             debug_draw_atributes, (1, 0, 0))
            debug_f1_function(elements_coordinates.Document_name_4(page_tables, document.doc_Type, cur_page.page_num), 0,               debug_draw_atributes, (1, 0, 0))
            debug_f1_function(elements_coordinates.Designation_of_documentation_type_5(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))

            debug_f1_function(elements_coordinates.Total_number_of_sheets_7(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))
            debug_f1_function(elements_coordinates.Signatures_Date_10(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))
            debug_f1_function(elements_coordinates.Current_Revision_18_1(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))
            debug_f1_function(elements_coordinates.Revision_Date_18_2(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))
            debug_f1_function(elements_coordinates.Purpose_of_issue_18_3(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))


        # Находим элементы, составляющие страницу
        for i, component in enumerate(page_elements):
            # Извлекаем положение верхнего края элемента в PDF
            pos = component[0]
            # Извлекаем элемент структуры страницы
            element = component[1]

            if debug_f1 and debug_draw_all_elements:
                page_fitz.draw_rect([element.x0, page_tables.height - element.y1, element.x1, page_tables.height - element.y0], color=(0, 0, 1), width=2)

                #Проверяем "шум" элементов (бывают артефакты - тонкие по ширине)
            if ((element.x1-element.x0)<=element_shtamp.min_x_x_dif):
                continue

            if debug_f1 and debug_draw_shtamp_elements:
                if check_element_position(elements_coordinates.Element_Shtamp(page_tables, document.doc_Type, cur_page.page_num), element):
                    if isinstance(element, LTTextContainer):
                        page_fitz.draw_rect([element.x0, page_tables.height - element.y1, element.x1, page_tables.height - element.y0], color=(0, 1, 1), width=2)



            #Проверяем - этот объет в зоне FILE_NAME_SHTAMP_50?

            #Проверяем - этот объет в зоне DOC_TITLE_1?
            find_attribute_func(elements_coordinates.DOC_TITLE_1,shtamp_extract_classes._1_DOC_TITLE)
            find_attribute_func(elements_coordinates.Sheet_number_Current_sheet_number_6_1, shtamp_extract_classes._6_1_Sheet_number_Current_sheet_number)
            find_attribute_func(elements_coordinates.Sheet_number_Current_sheet_number_6_2, shtamp_extract_classes._6_2_Sheet_number_Total_quantity_of_sheets)
            if not check_MTO_BBB_page or cur_page.page_num < 2:
                find_attribute_func(elements_coordinates.Current_Document_Revision_26, shtamp_extract_classes._26_Current_Document_Revision)
            find_attribute_func(elements_coordinates.File_Name_Shtamp_50, shtamp_extract_classes._50_File_Name_Shtamp)
            find_attribute_func(elements_coordinates.Page_Format_51, shtamp_extract_classes._51_Page_Format)

            if cur_page.page_num < 2:
                find_attribute_func(elements_coordinates.Construction_facility_name_2, shtamp_extract_classes._2_Construction_facility_name)
                find_attribute_func(elements_coordinates.Unit_title_name_3, shtamp_extract_classes._3_Unit_title_name)
                find_attribute_func(elements_coordinates.Document_name_4, shtamp_extract_classes._4_Document_name)
                find_attribute_func(elements_coordinates.Designation_of_documentation_type_5, shtamp_extract_classes._5_Designation_of_documentation_type)

                find_attribute_func(elements_coordinates.Total_number_of_sheets_7, shtamp_extract_classes._7_Total_number_of_sheets)
                find_attribute_func(elements_coordinates.Signatures_Date_10, shtamp_extract_classes._10_Signatures_Date)
                find_attribute_func(elements_coordinates.Current_Revision_18_1, shtamp_extract_classes._18_1_Current_Revision)
                find_attribute_func(elements_coordinates.Revision_Date_18_2, shtamp_extract_classes._18_2_Revision_Date)
                find_attribute_func(elements_coordinates.Purpose_of_issue_18_3, shtamp_extract_classes._18_3_Purpose_of_issue)



            # # Проверяем, является ли элемент текстовым
            # if isinstance(element, LTTextContainer):
            #
            #     # Проверяем, находится ли текст в таблице
            #     if table_extraction_flag == False:
            #         # Используем функцию извлечения текста и формата для каждого текстового элемента
            #         (line_text, format_per_line) = text_extraction(element)
            #         # Добавляем текст каждой строки к тексту страницы
            #         page_text.append(line_text)
            #
            #         if ((element.x1-element.x0)<=element_shtamp.min_x_x_dif):
            #             page_fitz.draw_rect(
            #                 [element.x0, page_tables.height - element.y1, element.x1, page_tables.height - element.y0],
            #                 color=(1, 0, 0), width=2)
            #         else:
            #             page_fitz.draw_rect([element.x0, page_tables.height - element.y1, element.x1, page_tables.height - element.y0],color=(0, 1, 0), width=2)
            #         # Добавляем формат каждой строки, содержащей текстзз
            #         line_format.append(format_per_line)
            #         page_content.append(line_text)
            #     else:
            #         # Пропускаем текст, находящийся в таблице
            #         pass

            # # Проверяем элементы на наличие таблиц
            # if isinstance(element, LTRect):
            #     # Если первый прямоугольный элемент
            #     if first_element == True and (table_num + 1) <= len(tables):
            #         # Находим ограничивающий прямоугольник таблицы
            #         lower_side = page.bbox[3] - tables[table_num].bbox[3]
            #         upper_side = element.y1
            #         # Извлекаем информацию из таблицы
            #         table = extract_table(pdf_path, pagenum, table_num)
            #         # Преобразуем информацию таблицы в формат структурированной строки
            #         table_string = table_converter(table)
            #         # Добавляем строку таблицы в список
            #         text_from_tables.append(table_string)
            #         page_content.append(table_string)
            #         # Устанавливаем флаг True, чтобы избежать повторения содержимого
            #         table_extraction_flag = True
            #         # Преобразуем в другой элемент
            #         first_element = False
            #         # Добавляем условное обозначение в списки текста и формата
            #         page_text.append('table')
            #         line_format.append('table')

                    # # Проверяем, извлекли ли мы уже таблицы из этой страницы
                    # if element.y0 >= lower_side and element.y1 <= upper_side:
                    #     pass
                    # elif not isinstance(page_elements[i + 1][1], LTRect):
                    #     table_extraction_flag = False
                    #     first_element = True
                    #     table_num += 1

        # Создаём ключ для словаря
        #dctkey = 'Page_' + str(pagenum)
        # Добавляем список списков как значение ключа страницы
        #text_per_page[dctkey] = [page_text, line_format, text_from_images, text_from_tables, page_content]

        cur_page.page_marka = string_parsing.getMarkaFromFileName(cur_page.dict_attributes[shtamp_extract_classes._50_File_Name_Shtamp])

        #Добавляем cur_page к массиву pages
        document.pages.append(cur_page)
    # КОНЕЦ ЦИКЛА СТРАНИЦЫ

    # Закрываем объект файла pdf
    pdfFileObj.close()
    # Save pdf fitz
    if debug_f1:
        fitz_save_file_name = "fitz_"+document.File_Full_Name+".pdf"
        fitz_doc.save(fitz_save_file_name)
#КОНЕЦ ЦИКЛА ДОКУМЕНТА

if debug_print_console_final_output:
    print()
    for out_doc in Curr_Proj:
        print("DOC name:", out_doc.File_Full_Name, " DOC Type:", "<" + out_doc.doc_Type + ">")
        for out_page in out_doc.pages:
            if out_page.page_num > 1:
                if debug_print_only_title_pages:
                    continue
            print("\t","стр.", out_page.page_num, "Марка страницы:", out_page.page_marka)
            if out_page.page_num >1:
                out_atr = shtamp_extract_classes._1_DOC_TITLE
                print("\t\t", out_atr, ":", out_page.dict_attributes[out_atr])

                out_atr = shtamp_extract_classes._6_1_Sheet_number_Current_sheet_number
                print("\t\t", out_atr, ":", out_page.dict_attributes[out_atr])

                out_atr = shtamp_extract_classes._6_2_Sheet_number_Total_quantity_of_sheets
                print("\t\t", out_atr, ":", out_page.dict_attributes[out_atr])
                if out_doc.doc_Type not in list_of_BBB:
                    out_atr = shtamp_extract_classes._26_Current_Document_Revision
                    print("\t\t", out_atr, ":", out_page.dict_attributes[out_atr])

                out_atr = shtamp_extract_classes._50_File_Name_Shtamp
                print("\t\t", out_atr, ":", out_page.dict_attributes[out_atr])
                if out_doc.doc_Type not in list_of_BBB:
                    out_atr = shtamp_extract_classes._51_Page_Format
                    print("\t\t", out_atr, ":", out_page.dict_attributes[out_atr])
            else:
                for out_atr in out_page.dict_attributes:
                    print("\t\t", out_atr, ":", out_page.dict_attributes[out_atr])

        print()
    print("FINISH CONSOLE OUT")


if debug_print_excel_final_output:
    from openpyxl import Workbook
    debug_excel_out_file_from_template = 1
    wb = Workbook()
    if debug_excel_out_file_from_template:
        wb = openpyxl.load_workbook("out_template.xlsx")
    #wb = openpyxl.open("out_template.xlsx")

    # grab the active worksheet
    ws = wb.active

    #Пишем первую строчку заголовков

    if not debug_print_excel_final_output:
        p = shtamp_extract_classes.page_SHTAMP_ATTRIBUTES()
        excel_string = []
        excel_string.append("File_Full_Name")
        excel_string.append("doc_Type")
        excel_string.append("page_num")
        excel_string.append("page_marka")
        for key in p.dict_attributes:
            excel_string.append(key)
        ws.append(excel_string)

    #Пишем остальные строчки с данными
    for out_doc in Curr_Proj:

        for out_page in out_doc.pages:
            excel_string = []
            excel_string.append(out_doc.File_Full_Name)
            excel_string.append(out_doc.doc_Type)
            excel_string.append(out_page.page_num)
            excel_string.append(out_page.page_marka)
            for out_atr in out_page.dict_attributes:
                excel_string.append(str(out_page.dict_attributes[out_atr]))
            ws.append(excel_string)

    # Save the file
    wb.save("sample.xlsx")
    print("FINISH EXCEL OUTT")

end_time = time.time() - start_time ##время работы программы
print("Время работы программы: ", end_time)

