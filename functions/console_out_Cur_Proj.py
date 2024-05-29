import shtamp_extract_classes
from string_parsing import list_of_BBB
def console_output(Curr_Proj, debug_print_only_title_pages = 0):
    print()
    print("START CONSOLE OUT for File Atribute Collecor")


    from prettytable import PrettyTable
    for out_doc in Curr_Proj:

        print("DOC name:", out_doc.File_Full_Name,
              " DOC Type:", "<" + out_doc.doc_Type + ">",
              " DOC Title:", "<" + out_doc.doc_Title_4d + ">",
              " DOC Marka:", "<" + out_doc.doc_Marka + ">",
              " DOC Number:", "<" + out_doc.doc_Number + ">")

        for out_page in out_doc.pages:
            if out_page.page_num > 1:
                if debug_print_only_title_pages:
                    continue
            print("\t", "стр.", out_page.page_num, "Марка страницы:", out_page.page_marka)

            table = PrettyTable()
            table.field_names = ["    ", "     ", "    Attribute", "    Value"]
            table.border = 0
            table.align = "l"

            if out_page.page_num > 1:

                out_atr = shtamp_extract_classes._1_DOC_TITLE
                table.add_row(['    ', '    ', out_atr,out_page.dict_attributes[out_atr] ])

                out_atr = shtamp_extract_classes._6_1_Sheet_number_Current_sheet_number
                table.add_row(['    ', '    ', out_atr,out_page.dict_attributes[out_atr] ])

                if out_doc.doc_Type in list_of_BBB:
                    out_atr = shtamp_extract_classes._6_2_Sheet_number_Total_quantity_of_sheets
                    table.add_row(['    ', '    ', out_atr,out_page.dict_attributes[out_atr] ])

                if out_doc.doc_Type not in list_of_BBB:
                    out_atr = shtamp_extract_classes._26_Current_Document_Revision
                    table.add_row(['    ', '    ', out_atr,out_page.dict_attributes[out_atr] ])

                out_atr = shtamp_extract_classes._50_File_Name_Shtamp
                table.add_row(['    ', '    ', out_atr,out_page.dict_attributes[out_atr] ])

                if out_doc.doc_Type not in list_of_BBB:
                    out_atr = shtamp_extract_classes._51_Page_Format
                    table.add_row(['    ', '    ', out_atr,out_page.dict_attributes[out_atr] ])

                out_atr = shtamp_extract_classes._61_Page_Layers
                table.add_row(['    ', '    ', out_atr, out_page.dict_attributes[out_atr]])
                out_atr = shtamp_extract_classes._62_Page_Bookmarks
                table.add_row(['    ', '    ', out_atr, out_page.dict_attributes[out_atr]])
                out_atr = shtamp_extract_classes._63_Page_Width
                table.add_row(['    ', '    ', out_atr, out_page.dict_attributes[out_atr]])
                out_atr = shtamp_extract_classes._64_Page_Height
                table.add_row(['    ', '    ', out_atr, out_page.dict_attributes[out_atr]])
                out_atr = shtamp_extract_classes._65_Page_Real_Format
                table.add_row(['    ', '    ', out_atr, out_page.dict_attributes[out_atr]])

                print(table)
            else:
                table = PrettyTable()
                table.field_names = ["    ", "     ", "    Attribute", "    Value"]
                table.border = 0
                table.align = "l"

                for out_atr in out_page.dict_attributes:
                    item = []
                    item.append('    ')
                    item.append('     ')
                    item.append(out_atr)
                    item.append(out_page.dict_attributes[out_atr])
                    table.add_row(item)
                print(table)

    print("FINISH CONSOLE OUT for File Atribute Collecor")