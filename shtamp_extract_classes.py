#Current_Project
#	Docs
#		Pages
import os

import string_parsing

_1_DOC_TITLE = "1_DOC_TITLE"
_2_Construction_facility_name = "2_Construction_facility_name"
_3_Unit_title_name = "3_Unit_title_name"
_4_Document_name = "4_Document_name"
_5_Designation_of_documentation_type = "5_Designation_of_documentation_type"
_6_1_Sheet_number_Current_sheet_number = "6_1_Sheet_number_Current_sheet_number"
_6_2_Sheet_number_Total_quantity_of_sheets = "6_2_Sheet_number_Total_quantity_of_sheets"
_7_Total_number_of_sheets = "7_Total_number_of_sheets"
_10_Signatures_Date = "10_Signatures_Date"
_18_1_Current_Revision = "18_1_Current_Revision"
_18_2_Revision_Date = "18_2_Revision_Date"
_18_3_Purpose_of_issue = "18_3_Purpose_of_issue"
_26_Current_Document_Revision = "26_Current_Document_Revision"
_50_File_Name_Shtamp = "50_File_Name_Shtamp"
_51_Page_Format = '51_Page_Format'
_61_Page_Layers = '61_Page_Layers'
_62_Page_Bookmarks = '62_Page_Bookmarks'
_63_Page_Width = '63_Page_Width'
_64_Page_Height = '64_Page_Height'
_65_Page_Real_Format = '65_Page_Real_Format'

class Current_Project:
    Project_Path = ""
class doc_ATTRIBUTES:

    File_Revision = ""
    File_Project_Number = ""
    File_Project_DOC_TITLE = ""

    def __init__ (self, file_path):
        self.File_Full_Path = file_path #
        self.File_Full_Name = os.path.basename(file_path) # AGCC.287-8525-SKUD.BOM-0001_A_RU.pdf
        self.pages = []
        self.doc_Type = string_parsing.getDocTypeFromFile(self.File_Full_Name) # MTO
        self.doc_Title_4d = string_parsing.getDocTitle(self.File_Full_Name) # 7417
        self.doc_Marka = string_parsing.getMarkaFromFileName(self.File_Full_Name) # SKUD
        self.doc_Number = string_parsing.getDocNumber(self.File_Full_Name) # 0001
        self.doc_Revision = string_parsing.getRevisionFromFileName(self.File_Full_Name) # A, B... 0, 01...

        self.doc_Short_Title = self.doc_Title_4d+"-"+self.doc_Marka #7417-SKUD
        self.doc_Short_File_Name = self.doc_Title_4d + "-" + self.doc_Marka+"."+ self.doc_Type # 7417-SKUD.MTO
        self.doc_OD_style_file_name = string_parsing.getOdStyleFileName(self.File_Full_Name)  # AGCC.287-8525-SKUD.BOM-0001

        sort_index = self.doc_Number
        if self.doc_Type == "MTO":
            sort_index = "5001"
        if self.doc_Type == "BOE":
            sort_index = "5002"
        if self.doc_Type == "BOM":
            sort_index = "5003"
        if self.doc_Type == "BOQ":
            sort_index = "5004"
        if self.doc_Type == "VO":
            sort_index = "5005"
        self.doc_Number_for_sort = sort_index

class page_SHTAMP_ATTRIBUTES:

    def __init__(self, pageNumber=-1):
        self.page_num = pageNumber
        self.page_marka = "" # SKUD
        self.page_type = "" # MTO
        self.page_title = "" # 7417
        self.page_revision = "" # A 0

        self.dict_attributes = {
            _1_DOC_TITLE :                                      [],
            _2_Construction_facility_name:                      [],
            _3_Unit_title_name:                                 [],
            _4_Document_name:                                   [],
            _5_Designation_of_documentation_type:               [],
            _6_1_Sheet_number_Current_sheet_number:             [],
            _6_2_Sheet_number_Total_quantity_of_sheets:         [],
            _7_Total_number_of_sheets:                          [],
            _10_Signatures_Date:                                [],
            _18_1_Current_Revision:                             [],
            _18_2_Revision_Date:                                [],
            _18_3_Purpose_of_issue:                             [],
            _26_Current_Document_Revision:                      [],
            _50_File_Name_Shtamp:                               [],
            _51_Page_Format:                                    [],
            _61_Page_Layers:                                    [],
            _62_Page_Bookmarks:                                 [],
            _63_Page_Width:                                     [],
            _64_Page_Height:                                    [],
            _65_Page_Real_Format:                               []
        }


