import shtamp_extract_classes
from shtamp_extract_classes import doc_ATTRIBUTES, page_SHTAMP_ATTRIBUTES
import string_parsing
import re


def read_Curr_Proj_from_excel(excel_file_name,  debug_print_console_final_output = 1):

    Curr_Proj = []

    import openpyxl
    from openpyxl import Workbook

    wb = Workbook()

    wb = openpyxl.load_workbook(excel_file_name)

    # grab the active worksheet
    ws = wb.active
    prev_doc_name = "none"
    for row in ws.values:

        if "AGCC" in row[0]:
            if prev_doc_name != row[0]:
                # Добавляем документ к проекту
                # doc_Type, doc_Title_4d, doc_Marka, doc_Number - Берутся из имени файла
                prev_doc_name = row[0] # - запоминаем имя документа для следующего цикла
                Curr_Proj.append(doc_ATTRIBUTES(prev_doc_name))

            document = Curr_Proj[-1]

            excel_page_number = row[2]

            cur_page = page_SHTAMP_ATTRIBUTES(excel_page_number)


            i = 4 #Номер начального столбца в экселе
            for out_atr in cur_page.dict_attributes:
                x = row[i]
                x = x.replace("[", "")
                x = x.replace("]", "")
                x = x.replace("'", "")
                #print("<"+x+">")
                x = x.strip()
                cur_page.dict_attributes[out_atr] = [x]
                #print("\t<" + x + ">")
                i += 1

            cur_page.page_marka = string_parsing.getMarkaFromFileName(cur_page.dict_attributes[shtamp_extract_classes._50_File_Name_Shtamp])
            document.pages.append(cur_page)



    if debug_print_console_final_output:
        from functions.console_out_Cur_Proj import console_output
        console_output(Curr_Proj)


    return Curr_Proj

