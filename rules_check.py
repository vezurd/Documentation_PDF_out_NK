import OD_tab_parsing
import elements_coordinates
import shtamp_extract_classes as e
import string_parsing
import inspect

def rules_check(Curr_Proj, proj_OD_list):
    check_list = []
    c_code_list = []

    #Проверка на кириллицу

    c_code = [1001,"Проверка на кириллицу"]
    if 1 == 1:
        c_code_list.append(c_code)
        attr_list_to_check = [e._1_DOC_TITLE,
                              e._18_1_Current_Revision,
                              e._26_Current_Document_Revision,
                              e._50_File_Name_Shtamp]
        for doc in Curr_Proj:
            for page in doc.pages:
                for attr in attr_list_to_check:
                    #Проверяем на пустой список
                    if not page.dict_attributes[attr]:
                        result = True
                        text = "Кириллица в атрибуте " + attr + " не найдена"
                    else:
                        #Если не пустой то берем первый елемент и проверяем его на кириллицу
                        value = page.dict_attributes[attr][0]

                        result = check_1_cirillic_letters(value)
                        if result == False:
                            text = "Кириллица в атрибуте " + attr + ": "+value
                        else:
                            text = "Кириллица в атрибуте " + attr + " не найдена"

                    check_list.append(check_row(result, c_code,doc.doc_OD_style_file_name,page,text))

    # Проверка на многослойность
    c_code = [1002,"Проверка на многослойность"]
    if 1 == 1:
        c_code_list.append(c_code)
        attr_list_to_check = [e._61_Page_Layers]
        for doc in Curr_Proj:
            for page in doc.pages:
                for attr in attr_list_to_check:
                    # [-1] - нет многослойности
                    # [1] - есть многослойность
                    value = page.dict_attributes[attr]
                    if value == [-1]:
                        result = True
                        text = "Нет многослойности " + attr
                    elif value == [1]:
                        result = False
                        text = "Есть многослойность " + attr
                    check_list.append(check_row(result, c_code, doc.doc_OD_style_file_name, page, text))

    # Проверка на заметки, аннотации, комментарии
    c_code = [1003,"Проверка на заметки, аннотации, комментарии"]
    if 1 == 1:
        c_code_list.append(c_code)
        attr_list_to_check = [e._62_Page_Bookmarks]
        for doc in Curr_Proj:
            for page in doc.pages:
                for attr in attr_list_to_check:
                    # [-1] - нет
                    # [1] - есть
                    value = page.dict_attributes[attr]
                    if value == [-1]:
                        result = True
                        text = "Нет Аннотаций " + attr
                    elif value == [1]:
                        result = False
                        text = "Есть Аннотации " + attr
                    check_list.append(check_row(result, c_code, doc.doc_OD_style_file_name, page, text))

    # Проверка на форматы листов в штампе и реального
    c_code = [1004,"Проверка на форматы листов в штампе и реального"]
    if 1 == 1:
        c_code_list.append(c_code)
        attr_list_to_check = [e._51_Page_Format]
        for doc in Curr_Proj:
            for page in doc.pages:
                #Проверяем что документ не относится к МТО или BBB и страница больше первой
                if not (doc.doc_Type in elements_coordinates.list_of_BBB and page.page_num >1):
                    #print (page.page_type, page.page_num )
                    for attr in attr_list_to_check:
                        # [-1] - не совпали
                        # [1] - совпали
                        try:
                            value = page.dict_attributes[attr][0]
                        except:
                            value = "Формат не найден"

                        page_format = string_parsing.get_Format_from_51(value)
                        real_format = page.dict_attributes[e._65_Page_Real_Format][0]
                        if page_format == real_format:
                            result = True
                            text = "Формат правльный " + attr + "   P <" + page_format+ ">   R <" + real_format + ">"

                        else:
                            result = False
                            text = "Формат не совпадет " + attr + "  P <" + page_format + ">   R <" + real_format + ">"

                        check_list.append(check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
                        #print(check_list[-1])

    # Проверка на форматы листов в штампе и из ОД
    c_code = [1005, "Проверка на Ревизии, внутри Документа"]
    if 1 == 1:
        c_code_list.append(c_code)
        attr_list_to_check = [e._18_1_Current_Revision,
                              e._26_Current_Document_Revision,
                              e._50_File_Name_Shtamp]
        list_OD_CJ = ["OD","CJ"]
        for doc in Curr_Proj:
            for page in doc.pages:
                # Проверяем что документ не относится к МТО или BBB и страница больше первой
                if not (doc.doc_Type in elements_coordinates.list_of_BBB and page.page_num > 1):
                    i = 0
                    value = []
                    for attr in attr_list_to_check:
                        if attr == e._18_1_Current_Revision and page.page_num>1:
                            continue
                        else:
                            try:
                                #Добавляем значение в лист для сравнения из документов
                                value.append(page.dict_attributes[attr][0])
                            except:
                                t = "Нет ревизии:"+str(doc.doc_Type)+str(page.page_num)+str(attr)
                                value.append(t)
                        if attr == e._50_File_Name_Shtamp:
                            value[-1] = string_parsing.getRevisionFromFileName(value[-1])
                    #print(value)
                    result = all(x == value[0] for x in value)
                    #print("result",result)
                    if result:
                        result = True
                        text = "Ревизии совпадают "+str(value)
                    else:
                        result = False
                        text = "Ревизии НЕ совпадают "+str(value)

                    check_list.append(check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
                    #print(check_list[-1])

    # Проверка на форматы листов в штампе и из ОД
    c_code = [1006, "Проверка на Ревизии, в Документе и ОД"]
    if 1 == 1:
        c_code_list.append(c_code)
        dict_OD_rev = {}
        for doc in proj_OD_list:
            dict_OD_rev[doc.Doc_Title] = doc.Document_Revision

        for doc in Curr_Proj:

            doc_OD_Style_Fomat_File_Name = doc.doc_OD_style_file_name # AGCC.287-8525-SKUD.BOM-0001
            doc_rev = doc.doc_Revision # A, B... 0, 01...
            try:
                OD_rev = dict_OD_rev[doc_OD_Style_Fomat_File_Name] # Рев. 0
                OD_rev = OD_rev.replace("Рев.", "").strip()
            except:
                OD_rev = "в ОД нет документа "+doc_OD_Style_Fomat_File_Name
            result = False
            if doc_rev == OD_rev:
                result = True

            if result:
                result = True
                text = "Ревизии совпадают. DOC:" +str(doc_rev)+" OD:"+ str(OD_rev)
            else:
                result = False
                text = "Ревизии НЕ совпадают. DOC:" +str(doc_rev)+" OD:"+ str(OD_rev)

            check_list.append(check_row(result, c_code, doc.doc_OD_style_file_name, "нет", text))

    # Проверка на форматы листов в штампе и из ОД
    c_code = [1007, "Проверка на форматы листов ОД и документов"]
    if 1 == 1:
        c_code_list.append(c_code)
        #Создаем словарь со списком всех форматов
        dict_doc_formats = {}
        for doc in Curr_Proj:
            curr_doc_name = doc.doc_OD_style_file_name
            dict_doc_formats[curr_doc_name] = []
            for page in doc.pages:
                try:
                    value = page.dict_attributes[e._51_Page_Format][0]
                except:
                    value = "Формат не найден"
                page_format = string_parsing.get_Format_from_51(value)
                dict_doc_formats[curr_doc_name].append(page_format)
        for doc in proj_OD_list:
            if doc.doc_range ==  OD_tab_parsing.doc_Prilagaemie:
                try:
                    dict_doc_formats.pop(doc.Doc_Title)
                except:
                    print("ERROR",doc.Doc_Title,"удаление прилагаемых докуменов" )
                    continue
        dict_OD_formats = {}
        for doc in proj_OD_list:
            if doc.doc_range != OD_tab_parsing.doc_Prilagaemie:
                title = doc.Doc_Title
                dict_OD_formats[title] = []
                fomat_list = doc.Page_Format.split()
                for x in fomat_list:
                    x = string_parsing.get_Format_from_51(x)
                    count = string_parsing.get_Format_Count(x)
                    single_format = string_parsing.get_Format_Single(x)
                    for i in range(count):
                        dict_OD_formats[title].append(single_format)

        for x in dict_doc_formats:
            dict_doc_formats[x].sort()
        for x in dict_OD_formats:
            dict_OD_formats[x].sort()
        for x in dict_doc_formats:
            try:
                if dict_doc_formats[x]!= dict_OD_formats[x]:
                    result = False
                    text = "Форматы НЕ совпадают. В документе: <"+formats_convolution(dict_doc_formats[x])+"> , в ОД: <"+formats_convolution(dict_OD_formats[x])+">"
                else:
                    result = True
                    text = "Форматы совпадают. В документе: <"+formats_convolution(dict_doc_formats[x])+"> , в ОД: <"+formats_convolution(dict_OD_formats[x])+">"
            except:
                result = False
                text = "Ошибка, <"+x+"> нет в од "

            check_list.append(check_row(result, c_code, x, "нет", text))

    c_code = [1008, "Проверка на кол-во листов в штампе и реальное кол-во"]
    if 1 == 1:
        c_code_list.append(c_code)
        for doc in Curr_Proj:
            for page in doc.pages:
                # Проверяем тлько ольшой штам на первой странице
                if page.page_num < 2:
                    count_6_2 = page.dict_attributes[e._6_2_Sheet_number_Total_quantity_of_sheets][0]

                    if int(count_6_2) == len(doc.pages):
                        result = True
                        text = f"Кол-во листов в штампе ({count_6_2}) и реальное кол-во ({len(doc.pages)}) - совпадают"
                    else:
                        result = False
                        text = f"Кол-во листов в штампе ({count_6_2})  и реальное кол-во ({len(doc.pages)}) - НЕ совпадают"

                    check_list.append(check_row(result, c_code, doc.doc_OD_style_file_name, page, text))

    c_code = [1009, "Проверка номеров документа (1==50==6.1_prefix)"]
    if 1 == 1:
        c_code_list.append(c_code)
        attr_list_to_check = [e._1_DOC_TITLE,
                              e._50_File_Name_Shtamp,
                              e._6_1_Sheet_number_Current_sheet_number]
        for doc in Curr_Proj:
            for page in doc.pages:
                page_num_list_6_1 = []
                if not (doc.doc_Type in elements_coordinates.list_of_BBB and page.page_num > 1):
                    for attr in attr_list_to_check:
                        page_num_list_6_1.append(string_parsing.get_page_num_prefix(page.dict_attributes[attr][0]))
                    result = all(x == page_num_list_6_1[0] for x in page_num_list_6_1)
                    if result:
                        result = True
                        text = f"Номера документов совпадают: {page_num_list_6_1}"
                    else:
                        result = False
                        text = f"Номера документов НЕ совпадают: {page_num_list_6_1}"
                    check_list.append(check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
                    #print(result, page_num_list_6_1, doc.doc_OD_style_file_name)

    c_code = [1010, "Проверка номеров страниц (№ стр.==6.1_suffix)"]
    if 1 == 1:
        c_code_list.append(c_code)
        for doc in Curr_Proj:
            for page in doc.pages:
                    try:
                        count = page.dict_attributes[e._6_2_Sheet_number_Total_quantity_of_sheets][0]
                    except:
                        count = ""
                    if count != "" and int(count) == 1:
                        value = 1
                    else:
                        value = string_parsing.get_page_num_suffix(page.dict_attributes[e._6_1_Sheet_number_Current_sheet_number][0])

                    if value == page.page_num:
                        result = True
                        text = f"Номера страниц совпадают: {value} и {page.page_num}"
                    else:
                        result = False
                        text = f"Номера страниц НЕ совпадают: {value} вместо {page.page_num}"
                    check_list.append(check_row(result, c_code, doc.doc_OD_style_file_name, page, text))
                    print(check_list[-1])

    for c in c_code_list:
        print(c)

    return check_list




def check_row (result,c_code, doc, page, text):
    if page == "нет":
        page_num = "нет"
    else:
        page_num = page.page_num

    return_list = [result,c_code ,doc, page_num,text]
    return return_list

def print_check_list(check_list, print_check_mode=0, c_code=0):
    # print_check_mode == 0 - только ошибки, все проверки
    # print_check_mode == 1 - и ошибки и не ошибки, все проверки
    # print_check_mode == 2 - только ошибки, проверики по c_code
    # print_check_mode == 3 - и ошибки и не ошибки, проверики по c_code
    from prettytable import PrettyTable
    table = PrettyTable()
    table.field_names = ["Флаг", "Код проверки", "Имя документа", "№ стр.", "Вывод проверки"]
    table.border = 1
    table.align = "l"
    print("Выводим список ошибок...")
    for i in range(len(check_list)):
        if print_check_mode==0:
            if not check_list[i][0]:
                table.add_row(check_list[i])
        elif print_check_mode==1:
            table.add_row(check_list[i])
        elif print_check_mode==2:
            if c_code == check_list[i][1][0]:
                if not check_list[i][0]:
                    table.add_row(check_list[i])
        elif print_check_mode==3:
            if c_code == check_list[i][1][0]:
                table.add_row(check_list[i])

    print(table)
    print("Конец вывода ошибок.")
def check_1_cirillic_letters(text, alphabet=set('абвгдеёжзийклмнопрстуфхцчшщъыьэюя')):
    # test - True
    # тест - False
    return alphabet.isdisjoint(text.lower())

def formats_convolution (input_list):
    dict_list = {}
    for i in input_list:
        if i in dict_list:
            dict_list[i] = dict_list[i]+1
        else:
            dict_list[i] = 1
    output_str = ""
    for k in dict_list:
        if dict_list[k] > 1:
            output_str = output_str+str(dict_list[k])+k+", "
        else:
            output_str = output_str+k+", "
    output_str = output_str.strip().strip(",")
    return output_str