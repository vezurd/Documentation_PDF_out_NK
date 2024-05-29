import re

list_of_BBB = ["BOM", "BOE", "BOQ", "MTO"]

def getOdStyleFileName (fileName):
    #Для примера AGCC.287-7417-SOS.MTO-0001_01_RU.pdf - ищем чать "AGCC.287-7417-SOS.MTO-0001"
    #Находим положение в строке второго знака "-"
    try:
        if type(fileName) == list:
            fileName = fileName[0]
        stat_index = 0
        end_index = fileName.find("_")

        output_text = fileName[stat_index: end_index]
        return output_text
    except:
        return "<No Found>"
def getRevisionFromFileName (fileName):
    #Для примера AGCC.287-7417-SOS.MTO-0001_01_RU.pdf - ищем чать "01"
    #Находим положение в строке второго знака "-"
    try:
        if type(fileName) == list:
            fileName = fileName[0]
        stat_index = fileName.find("_")+1
        end_index = fileName.find("_",stat_index)

        output_text = fileName[stat_index: end_index]
        return output_text
    except:
        return "<No Found>"
#Извлекаем МАРКУ (SOS, SKUD, SOT) титула из штампа или имени файла
def getMarkaFromFileName (fileName):
    #Для примера AGCC.287-7417-SOS.MTO-0001_01_RU.pdf - ищем чать "SOS"
    #Находим положение в строке второго знака "-"
    try:
        if type(fileName) == list:
            fileName = fileName[0]
        stat_index = fileName.find("-")
        stat_index = fileName.find("-",stat_index+1)+1

        # Находим положение в строке второго знака "."
        end_index = fileName.find(".")
        end_index = fileName.find(".", end_index + 1)

        output_text = fileName[stat_index: end_index]
        return output_text
    except:
        return "<No Found>"

def getDocTypeFromFile (fileName):
    #Для примера AGCC.287-7417-SOS.MTO-0001_01_RU.pdf - ищем чать "MTO"
    #Находим положение в строке второго знака "-"
    stat_index = fileName.find(".")
    stat_index = fileName.find(".",stat_index+1)+1

    # Находим положение в строке второго знака "."
    end_index = fileName.find("-")
    end_index = fileName.find("-", end_index + 1)
    end_index = fileName.find("-", end_index + 1)

    output_text = fileName[stat_index: end_index]
    return output_text

def getDocNumber(fileName):
    # Для примера AGCC.287-7417-SOS.MTO-0001_01_RU.pdf - ищем чать "0001"
    # Находим положение в строке второго знака "-"
    stat_index = fileName.find("-")
    stat_index = fileName.find("-", stat_index + 1)
    stat_index = fileName.find("-", stat_index + 1)+1

    end_index = stat_index + 4

    output_text = fileName[stat_index: end_index]
    return output_text
def getDocTitle (fileName):
    #Для примера AGCC.287-7417-SOS.MTO-0001_01_RU.pdf - ищем чать "7417"
    #Находим положение в строке второго знака "-"
    stat_index = fileName.find("-")
    end_index = fileName.find("-",stat_index+1)

    output_text = fileName[stat_index+1: end_index]
    return output_text
def string_Remove_New_Lines(input_text):
    output_text = input_text
    output_text = re.sub(r"\n", '', output_text)
    #output_text = re.sub(r"\r", '', output_text)
    output_text = output_text.strip()
    return output_text

def string_Junk_Cleaner(input_text, flag=0):
    if type(input_text) != list:
        l = input_text.split("\n")
    else:
        l = input_text


    names = ['Дата', 'Date', "Рев.", "Rev.", "Назначение выпуска","Ревизия",
             "Purpose of Issue", "ООО \"Би.Си.Си.\"", "Примечание", "Стадия"]

    if flag == 1:

        names.append("Лист")
        names.append("Листов")

    indexes = []
    for i in l:
        for n in names:
            if n in i or i =="" or i ==" ":
                indexes.append(l.index(i))

    res = []
    for i in range(0, len(l)):
        if i not in indexes:
            res.append(l[i].strip())

    return res

def list_6_1_extrction(input_text, docType="LAY", pageNum=-1):
    if docType not in list_of_BBB or pageNum<2:
        l = string_Junk_Cleaner(input_text)
        names = ["Лист", "Sheet", "на", "AGCC" ]

        indexes = []
        for i in l:
            for n in names:
                if n in i:
                    indexes.append(l.index(i))

        res = []
        for i in range(0, len(l)):
            if i not in indexes:
                res.append(l[i])
    else:

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
        end_index = sting_f.find("из", start_index+1)
        res = sting_f[start_index+2:end_index].strip()

    return res

def list_6_2_extrction(input_text, docType="LAY", pageNum=-1):

    if docType not in list_of_BBB or pageNum<2:
        l = string_Junk_Cleaner(input_text)
        names = ["на"]

        indexes = []
        for i in l:
            for n in names:
                if n in i:
                    indexes.append(l.index(i))

        res = []
        for i in range(0, len(l)):
            if i in indexes:
                res.append(l[i])
        sting_f = str(res)
        start_index = sting_f.find("на")
        end_index = sting_f.find("л.", start_index + 1)
        res = sting_f[start_index + 2:end_index].strip()
    else:
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
        start_index = sting_f.find("из")
        end_index = sting_f.find("листов", start_index + 1)
        res = sting_f[start_index + 2:end_index].strip()

    return res



def list_18_1_extrction(input_text):

    if "Кабели одного типа" in input_text:
        return ""
    l = string_Junk_Cleaner(input_text)
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



def list_18_2_extrction(input_text):

    try:
        l = string_Junk_Cleaner(input_text)
        a = str(l)
        res = re.sub(r"[/\\.]", ".", re.search(r"(\d+.*?\d+.*?\d+)", a).group(1))
    except:
        return ""
    return res

def list_18_3_extrction(input_text):

    l = string_Junk_Cleaner(input_text)
    a = str(l)
    a = a.strip()
    a =  re.sub(r"[\][\']", "", a)

    if a is None:
        return a

    data_junk = re.sub(r"[/\\.]", ".",a)
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

def list_flatter(input_list):
    out_list = []
    for x in input_list:
        out_list.extend(x if isinstance(x, list) else [x])
    return out_list

def string_remove_spaces(input_string):
    return input_string.replace(' ', '')

def get_Format_from_51 (input_format):
    output_format = input_format

    output_format = input_format.replace("Формат", "")
    output_format = output_format.replace(",", "")
    output_format = output_format.replace("А","A")
    output_format = output_format.replace("х", "x")
    output_format = output_format.replace("×", "x")

    return output_format.strip()

def get_Format_Count (input_format):
    if input_format[0] == "A":
        return 1
    else:
        return int(input_format.split("A")[0])
def get_Format_Single (input_format):
    if input_format[0] == "A":
        return input_format
    else:
        return "A"+input_format.split("A")[1]

def get_page_num_prefix (value):
    if "AGCC" in value:
        value = int(getDocNumber(value))
    elif "." in value:
        value = int(value.split(".")[0])
    else:
        value = int(value)
    return value

def get_page_num_suffix (value):
    if "." in value:
        value = int(value.split(".")[1])
    else:
        value = int(value)
    return value