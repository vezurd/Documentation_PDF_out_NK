import re

import string_parsing
from string_parsing import list_of_BBB

# Коэффициент для перевода мм в размеры PDF длинн
scale_number = 2.83444
border_left = 0
border_bottom = 0
border_top = 1
border_top_max = 12
borders_margin = -1
min_x_x_dif = 1 * scale_number

class Element_Shtamp:
    # Для отсева шумов элементов (слишком маленьких)


    def __init__(self, page_tables, docType="LAY", pageNum=-1):


        left_bottom_x = 185
        left_bottom_y = -10
        right_top_x = 0
        right_top_y = 110

        if pageNum>1:
            right_top_y = 15

        self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y, page_tables)


class File_Name_Shtamp_50:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 185
        left_bottom_y = -border_bottom
        right_top_x = 80
        right_top_y = -2

        self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y,page_tables)

        if (docType=="MTO" and pageNum==1):
           self.x0 = 0
        if (docType=="BOM" or docType=="BOE" or docType=="BOQ"):
           self.x0 = 0

    def Check_Text(input_text):
        if "AGCC" in input_text:
            return 1
        else:
            return 0

    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        index = input_text.find("AGCC")
        if input_text.find("AGCC",index+1)!= -1:
            return "Задвоение надписи"
        output_text = input_text[input_text.find("AGCC"): len(input_text)]
        output_text = string_parsing.string_Remove_New_Lines(output_text)
        return output_text

class Page_Format_51:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 80
        left_bottom_y = -border_bottom
        right_top_x = 0
        right_top_y = -2

        self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y,page_tables)


    def Check_Text(input_text):
        # #if "Формат" in input_text:
        #     return 1
        # else:
        #     return 0
        return 1

    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        out_text = string_parsing.string_Junk_Cleaner(input_text)
        return out_text

class DOC_TITLE_1:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 120
        left_bottom_y = 45
        right_top_x = 0
        right_top_y = 55

        if pageNum >1:
            left_bottom_x = 110
            left_bottom_y = 0
            right_top_x = 25
            right_top_y = 15

        if pageNum > 1 and docType in list_of_BBB:
            left_bottom_x = 40
            right_top_x = 0

            self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y, page_tables)

            if (border_top > border_top_max):
                self.y0 = page_tables.height - (border_top_max*scale_number)
            else:
                self.y0 = page_tables.height - (border_top*scale_number)
            self.y1 = page_tables.height

        else:
            self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y, page_tables)

    def Check_Text(input_text):
        if "AGCC" in input_text:
            return 1
        else:
            return 0

    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        start_index = input_text.find("AGCC")
        end_index = input_text.find("\n",start_index + 1)
        output_text = input_text[start_index:end_index]
        return output_text

class Construction_facility_name_2:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 90
        left_bottom_y = 33
        right_top_x = 0
        right_top_y = 45

        self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y,page_tables)

    def Check_Text(input_text):
        return 1
    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        output_text = input_text
        if "AGCC" in input_text:
            start_index = input_text.find("AGCC")
            end_index = input_text.find("\n", start_index + 1)
            output_text = input_text[end_index+1:len(input_text)]

        s_list = string_parsing.string_Junk_Cleaner(output_text,1)
        output_text = ""
        for i in range(len(s_list)):
            output_text = output_text + s_list[i] +" "
        return output_text.strip()

class Unit_title_name_3:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 100
        left_bottom_y = 15
        right_top_x = 55
        right_top_y = 30

        self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y,page_tables)

    def Check_Text(input_text):
        return 1
    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        return string_parsing.string_Remove_New_Lines(input_text)

class Document_name_4:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 100
        left_bottom_y = 0
        right_top_x = 55
        right_top_y = 15

        self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y,page_tables)

    def Check_Text(input_text):
        return 1
    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        return string_parsing.string_Remove_New_Lines(input_text)

class Designation_of_documentation_type_5:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 45
        left_bottom_y = 15
        right_top_x = 35
        right_top_y = 25

        self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y,page_tables)

    def Check_Text(input_text):
        return 1
    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        out = string_parsing.string_Remove_New_Lines(input_text)
        return string_parsing.string_Junk_Cleaner(out,0)

class Sheet_number_Current_sheet_number_6_1:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 35
        left_bottom_y = 20
        right_top_x = 20
        right_top_y = 25

        if pageNum >1:
            left_bottom_x = 10
            left_bottom_y = 0
            right_top_x = 0
            right_top_y = 8

        if pageNum > 1 and docType in list_of_BBB:
            left_bottom_x = 40
            right_top_x = 0

            self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y, page_tables)

            if (border_top > border_top_max):
                self.y0 = page_tables.height - (border_top_max * scale_number)
            else:
                self.y0 = page_tables.height - (border_top * scale_number)
            self.y1 = page_tables.height

        else:
            self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y, page_tables)

    def Check_Text(input_text):
        return 1
    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        out_text = string_parsing.list_6_1_extrction(input_text,docType,pageNum)
        return out_text

class Sheet_number_Current_sheet_number_6_2:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 35
        left_bottom_y = 15
        right_top_x = 20
        right_top_y = 20

        if (pageNum > 1) and (docType in list_of_BBB):
            left_bottom_x = 40
            right_top_x = 0

            self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y, page_tables)

            if (border_top > border_top_max):
                self.y0 = page_tables.height - (border_top_max * scale_number)
            else:
                self.y0 = page_tables.height - (border_top * scale_number)
            self.y1 = page_tables.height

        else:
            self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y, page_tables)

    def Check_Text(input_text):
        return 1
    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        out_text = string_parsing.list_6_2_extrction(input_text,docType,pageNum)
        return out_text

class Total_number_of_sheets_7:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 17
        left_bottom_y = 15
        right_top_x = 0
        right_top_y = 25

        self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y,page_tables)

    def Check_Text(input_text):
        return 1
    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        out = string_parsing.string_Remove_New_Lines(input_text)
        return string_parsing.string_Junk_Cleaner(out, 1)

class Signatures_Date_10:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 130
        left_bottom_y = 25
        right_top_x = 120
        right_top_y = 30

        self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y,page_tables)

    def Check_Text(input_text):
        return 1
    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        out_text = string_parsing.list_18_2_extrction(input_text)
        return out_text

class Current_Revision_18_1:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 185 #180
        left_bottom_y = 90 #95
        right_top_x = 170
        right_top_y = 100

        if docType in list_of_BBB:
            #left_bottom_x=left_bottom_x+10
            right_top_x = 156
            left_bottom_y  = 85

        self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y,page_tables)

    def Check_Text(input_text):
        return 1
    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        out_text = string_parsing.list_18_1_extrction(input_text)
        return out_text

class Revision_Date_18_2:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 166
        left_bottom_y = 90 #95
        right_top_x = 155
        right_top_y = 100 #100

        if docType in list_of_BBB:
            #left_bottom_x=left_bottom_x+10
            right_top_x = 144
            left_bottom_y  = 85



        self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y,page_tables)

    def Check_Text(input_text):
        return 1
    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        out_text = string_parsing.list_18_2_extrction(input_text)
        return out_text

class Purpose_of_issue_18_3:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 145
        left_bottom_y = 90 #95
        right_top_x = 60
        right_top_y = 100 #100

        if docType in list_of_BBB:
            left_bottom_y  = 85

        self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y,page_tables)

    def Check_Text(input_text):
        return 1
    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        out_text = string_parsing.list_18_3_extrction(input_text)
        return out_text

class Current_Document_Revision_26:
    def __init__(self, page_tables, docType="LAY", pageNum=-1):
        left_bottom_x = 17
        left_bottom_y = 55
        right_top_x = 0
        right_top_y = 70

        if pageNum >1:
            left_bottom_x = 20
            left_bottom_y = 0
            right_top_x = 10
            right_top_y = 8

        self.x0, self.x1, self.y0, self.y1 = baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y,page_tables)

    def Check_Text(input_text):
        return 1
    def Clean_Text(input_text,docType="LAY",pageNum=-1):
        out_text = string_parsing.string_Junk_Cleaner(input_text)
        return out_text
def baese_point_diff_calc(left_bottom_x, left_bottom_y, right_top_x, right_top_y, page_tables):

    x0 = int((page_tables.width - (border_left + borders_margin + left_bottom_x) * scale_number))
    x1 = int((page_tables.width - (border_left - borders_margin + right_top_x) * scale_number))

    y0 = int(((border_bottom - borders_margin + left_bottom_y) * scale_number))
    y1 = int(((border_bottom + borders_margin + right_top_y) * scale_number))

    return x0, x1, y0, y1

