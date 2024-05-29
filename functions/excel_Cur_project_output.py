
def excel_Curr_Proj_output_f (Curr_Proj, excel_template_file_name):
    print()
    print("START EXCEL OUT")
    import openpyxl
    from openpyxl import Workbook
    excel_out_file_name = "__result_output_"

    #Загружаем шаблон экселя
    wb = openpyxl.load_workbook(excel_template_file_name)

    # grab the active worksheet
    ws = wb.active

    # Добавляем к шаблону строчки с данными
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
    wb.save(excel_out_file_name + Curr_Proj[0].doc_Title_4d + "-" + Curr_Proj[0].doc_Marka + ".xlsx")
    print("\tFINISH EXCEL OUT")