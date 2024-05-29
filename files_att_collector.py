

#Общий рубильник дебага

debug_f1 = 1
dubug_fitz = 1

debug_print_console_final_output = 1
debug_print_only_title_pages = 0

debug_print_files_list = 1

#Не выводить результаты в консоль
debug_no_output = 0
debug_draw_borders = 0

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

import shtamp_extract_classes
# Для анализа структуры PDF и извлечения текста
from pdfminer.high_level import extract_pages, extract_text
from pdfminer.layout import LTTextContainer, LTChar, LTRect, LTFigure
# Для извлечения текста из таблиц в PDF
import pdfplumber
# Для извлечения изображений из PDF
from PIL import Image
from pdf2image import convert_from_path

# Для удаления дополнительно созданных файлов
import os
from shtamp_extract_classes import *
from elements_coordinates import *

import fitz

from shtamp_extract_classes import Current_Project, doc_ATTRIBUTES, page_SHTAMP_ATTRIBUTES

from PDF_check_proc import extract_table, table_converter, text_extraction, check_element_position, pdf_get_text_from_file, crop_image

from os import listdir
from os.path import isfile, join

from elements_coordinates import Element_Shtamp, File_Name_Shtamp_50
import string_parsing

from elements_coordinates import scale_number

from string_parsing import list_of_BBB


def check_MTO_BBB_page(doc_type):
    if doc_type in list_of_BBB:
        return True
    else:
        return False

def files_attribute_collector_f (pdf_folder):
    # Находим путь к PDF
    pdf_path = pdf_folder

    def debug_f1_function(element_1, print_text, debug_flag, draw_color=(0, 1, 0)):
        if debug_f1 and debug_flag:
            if print_text == 0:
                print_text = str(element_1.__class__.__name__)
            if debug_flag_print_console:
                print("<START DEBUG:\t" + print_text + ">")
                print(element_1.x0, element_1.y0, element_1.x1, element_1.y1, "\t<page_widh:" + str(page_tables.width) + ", page_height:" + str(page_tables.height) + ">")
                print("<END DEBUG:\t" + print_text + ">")
            page_fitz.draw_rect([element_1.x0 - 1, page_tables.height - element_1.y1 - 1, element_1.x1 + 1, page_tables.height - element_1.y0 - 1], color=draw_color, width=2)

    # Осносной цикл поиска атрибутов в заданной области
    # temp_element_cooridinates - класс из elements_coordinates.py
    # temp_shtamp_extract_class - название класса из shtamp_extract_classes.py
    def find_attribute_func(temp_element_cooridinates, temp_shtamp_extract_class):
        # Если границы елементов пересекаются или element полностью находится внутри искомой области атрибута
        if check_element_position(temp_element_cooridinates(page_tables, document.doc_Type, cur_page.page_num), element):
            # Это текст?
            if isinstance(element, LTTextContainer):
                if debug_f1 and debug_draw_shtamp_elements:
                    page_fitz.draw_rect([element.x0, page_tables.height - element.y1, element.x1, page_tables.height - element.y0], color=(0, 1, 0), width=2)
                check_text = element.get_text()
                if temp_element_cooridinates.Check_Text(check_text):
                    check_text = temp_element_cooridinates.Clean_Text(check_text, document.doc_Type, cur_page.page_num)
                    cur_page.dict_attributes[temp_shtamp_extract_class].append(check_text)
                    cur_page.dict_attributes[temp_shtamp_extract_class] = string_parsing.list_flatter(cur_page.dict_attributes[temp_shtamp_extract_class])

    #

    class biggest_element:
        x0 = 0
        y0 = 0
        x1 = 0
        y1 = 0


    #Массив для хранения всех Документов (файлов) текщуего проекта
    Curr_Proj = []
    #Читаем все файлы из папки (без директорий)
    onlyfiles = [f for f in listdir(pdf_path) if isfile(join(pdf_path, f))]
    #Заносим все пути файлов и создаем массив документов - уникальное свойство - File_Full_Name (без пути).

    for next_file in onlyfiles:
        if not next_file.endswith(".pdf"):
            continue
        Curr_Proj.append(doc_ATTRIBUTES(pdf_path+"\\"+next_file))

    Curr_Proj =  list(sorted(Curr_Proj, key=lambda x: x.doc_Number_for_sort))

    ############################
    if debug_print_files_list:
        from prettytable import PrettyTable
        table = PrettyTable()
        table.field_names = ["Полный путь файла", "Имя файла"]
        table.border = 0
        table.align = "l"
        for document in Curr_Proj:
            table.add_row([document.File_Full_Path, document.File_Full_Name])
        print(table)

    #Цикл перебора всех документов (PDF файлов)
    for document in Curr_Proj:

        fitz_doc = fitz.open(document.File_Full_Path)
        # Проверяем наличие мульти-слоёв в документе
        doc_multi_layers = fitz_doc.layer_ui_configs()
        # doc_multi_layers = [-1] - нет многослойности
        # doc_multi_layers = [1] - есть многослойность
        if doc_multi_layers == []:
            doc_multi_layers = [-1]
        else:
            doc_multi_layers = [1]


        # создаём объект файла PDF
        pdfFileObj = open(document.File_Full_Path, 'rb')
        # создаём объект считывателя PDF
        pdfReaded = PyPDF2.PdfReader(pdfFileObj)

        # Счетчик страниц в документе
        page_counter = 0

        # Извлекаем страницы из PDF
        for pagenum, page in enumerate(extract_pages(document.File_Full_Path)):
            #Добавляем страницу в документ
            page_counter += 1
            #Создаем объект страницы
            cur_page = page_SHTAMP_ATTRIBUTES(page_counter)

            #Проверяем наличие аннотаций-комментариев в PDF
            # annotations_flag = [-1] - аннотаций нет
            # annotations_flag = [1] - аннотаций есть
            fitz_page = fitz_doc[page_counter-1]
            annotations_flag = [-1]
            for annot in fitz_page.annots(types=[fitz.PDF_ANNOT_TEXT]):
                if len(annot.info)>0:
                    annotations_flag = [1]

            # Открываем файл pdf
            pdf = pdfplumber.open(document.File_Full_Path)

            # Находим исследуемую страницу
            page_tables = pdf.pages[pagenum]


            # Находим все элементы
            page_elements = [(element.y1, element) for element in page._objs]

            # Сортируем все элементы по порядку нахождения на странице
            page_elements.sort(key=lambda a: a[0], reverse=True)

            # Open the pdf
            if debug_f1:
                page_fitz = fitz_doc[pagenum]
                page_fitz.clean_contents()

            # Находим границы рамки текущей страницы
            biggest_element.x0 = biggest_element.y0 = biggest_element.x1 = biggest_element.y1 = 0
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

            #element_shtamp = elements_coordinates.Element_Shtamp(page,document.doc_Type)

            # Рисуем рамку листа
            debug_f1_function(biggest_element,"biggest_element", debug_draw_biggest_element, (0,1,1))

            #Рисуем рамку штампа
            if not check_MTO_BBB_page(document.doc_Type) or cur_page.page_num < 2:
                debug_f1_function(elements_coordinates.Element_Shtamp(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_shtamp, (0, 0, 1))

            # Рисуем рамки атрибутов
            debug_f1_function(elements_coordinates.DOC_TITLE_1(page_tables, document.doc_Type, cur_page.page_num), 0,                   debug_draw_atributes, (1, 0, 0))

            if not check_MTO_BBB_page(document.doc_Type) or cur_page.page_num < 2:
                debug_f1_function(elements_coordinates.Current_Document_Revision_26(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))

            debug_f1_function(elements_coordinates.File_Name_Shtamp_50(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))

            if not check_MTO_BBB_page(document.doc_Type) or cur_page.page_num < 2:
                debug_f1_function(elements_coordinates.Page_Format_51(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))

            debug_f1_function(elements_coordinates.Sheet_number_Current_sheet_number_6_1(page_tables, document.doc_Type, cur_page.page_num), 0, debug_draw_atributes, (1, 0, 0))

            if check_MTO_BBB_page(document.doc_Type) or cur_page.page_num < 2:
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
                if ((element.x1-element.x0)<=elements_coordinates.min_x_x_dif):
                    continue

                if debug_f1 and debug_draw_shtamp_elements:
                    if check_element_position(elements_coordinates.Element_Shtamp(page_tables, document.doc_Type, cur_page.page_num), element):
                        if isinstance(element, LTTextContainer):
                            page_fitz.draw_rect([element.x0, page_tables.height - element.y1, element.x1, page_tables.height - element.y0], color=(0, 1, 1), width=2)

                #Проверяем - этот объет в зоне DOC_TITLE_1?
                find_attribute_func(elements_coordinates.DOC_TITLE_1,shtamp_extract_classes._1_DOC_TITLE)
                find_attribute_func(elements_coordinates.Sheet_number_Current_sheet_number_6_1, shtamp_extract_classes._6_1_Sheet_number_Current_sheet_number)
                if check_MTO_BBB_page(document.doc_Type) or cur_page.page_num < 2:
                    find_attribute_func(elements_coordinates.Sheet_number_Current_sheet_number_6_2, shtamp_extract_classes._6_2_Sheet_number_Total_quantity_of_sheets)
                if not check_MTO_BBB_page(document.doc_Type) or cur_page.page_num < 2:
                    find_attribute_func(elements_coordinates.Current_Document_Revision_26, shtamp_extract_classes._26_Current_Document_Revision)
                find_attribute_func(elements_coordinates.File_Name_Shtamp_50, shtamp_extract_classes._50_File_Name_Shtamp)
                if not check_MTO_BBB_page(document.doc_Type) or cur_page.page_num < 2:
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

            #Присвоение вторичных атрибутов - которых нет в штампе в чистом виде
            cur_page.page_marka = string_parsing.getMarkaFromFileName(cur_page.dict_attributes[shtamp_extract_classes._50_File_Name_Shtamp])
            cur_page.dict_attributes[shtamp_extract_classes._61_Page_Layers] = doc_multi_layers
            cur_page.dict_attributes[shtamp_extract_classes._62_Page_Bookmarks] = annotations_flag
            cur_page.dict_attributes[shtamp_extract_classes._63_Page_Width] = p_r_width = [int(page_tables.width / scale_number)]
            cur_page.dict_attributes[shtamp_extract_classes._64_Page_Height] = p_r_height = [int(page_tables.height / scale_number)]
            from functions.parsing_page_real_format import get_Real_Page_Format
            cur_page.dict_attributes[shtamp_extract_classes._65_Page_Real_Format] = [get_Real_Page_Format(p_r_width[0],p_r_height[0])]
            #Добавляем cur_page к массиву pages
            document.pages.append(cur_page)
        # КОНЕЦ ЦИКЛА СТРАНИЦЫ

        # Закрываем объект файла pdf
        pdfFileObj.close()
        # Save pdf fitz
        if debug_f1:
            fitz_save_file_name = "debug\\"+"dbg_"+document.File_Full_Name
            fitz_doc.save(fitz_save_file_name)
    #КОНЕЦ ЦИКЛА ДОКУМЕНТА

    for doc in Curr_Proj:
        for page in doc.pages:
            for attr in page.dict_attributes:
                page.dict_attributes[attr] = list(filter(None, page.dict_attributes[attr]))
                if len(page.dict_attributes[attr]) != 0:
                    item = []
                    if attr == shtamp_extract_classes._2_Construction_facility_name:
                        output_text = ""
                        for i in range(len(page.dict_attributes[attr])):
                            output_text = output_text + page.dict_attributes[attr][i].strip() + " "
                        page.dict_attributes[attr][0] = output_text.strip()
                    item.append(page.dict_attributes[attr][0])
                    page.dict_attributes[attr] = item

    if debug_print_console_final_output:
        from functions.console_out_Cur_Proj import console_output
        console_output(Curr_Proj, debug_print_only_title_pages)

    return Curr_Proj

