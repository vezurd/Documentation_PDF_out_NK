import RFQ.Value_Compare
import utils
import utils.path
from base.base_cheks import base_vs_base_std
from base.base_classes import *
from base.base_excel_out import check_color_out
from base.base_google import load_base
from base.base_mto import get_mto_std_from_file, get_std_from_excel_file
from tags import tag_parser
from utils.path import open_dir

debug_print_files_list = 1
debug_print_progress = 1


def rfq_start(pdf_path):
    print("RFQ сравнение тегов")

    # Массив для хранения всех Документов (файлов) текущего проекта
    curr_proj = utils.path.get_files_single(pdf_path, endswith=(".xlsx", ".XLSX"))
    if curr_proj == -1:
        print(f"По пути <{pdf_path}> файлы не найдены.")
        exit(0)

    ############################
    if debug_print_files_list:
        from prettytable import PrettyTable
        table = PrettyTable()
        table.field_names = ["Полный путь файла", "Имя файла"]
        table.border = 0
        table.align = "l"
        for document in curr_proj:
            table.add_row([document.file_full_path, document.file_name])
        print(table)
    mto_base = []
    rfq_base = []
    all_bases = []
    for document in curr_proj:
        if "MTO" in document.file_name:
            print(f"    Open MTO - {document.file_name}")
            mto_base = get_mto_std_from_file(t_com=TableComments(file_full_path=document.file_full_path,
                                                                 dir_path="-1",
                                                                 tabel_class=TableComments.MTO,
                                                                 sheet_name=ColNames.MTO.sheet_name)
                                             )
            all_bases.append(mto_base)
        if "RFQ" in document.file_name:
            print(f"    Open RFQ - {document.file_name}")
            t_com = TableComments(file_full_path=document.file_full_path,
                                  tabel_class=TableComments.RFQ,
                                  sheet_name=ColNames.RFQ.sheet_name)
            rfq_base = get_std_from_excel_file(t_com)
            all_bases.append(rfq_base)

    # 1 Сравнение RFQ с MTO
    input_base = base_vs_base_std(rfq_base, mto_base, [NAME, TYPE_MARK, UNITS, CODE, TAGS, MASS])

    column_dict = ColNames.RFQ.column_dict
    column_dict[99] = ROW_TYPE
    path_out_dir = utils.path.get_path_out_dir(pdf_path)

    return_file_path = check_color_out(input_base,
                                       column_dict,
                                       path_out_dir,
                                       "RFQ_vs_MTO",
                                       excel_template_check_color=r"templates\out_template_check_rfq_color.xlsx")

    # 2 Tag Analyze
    # Сброс для tag_analyzer
    tag_parser.reset()
    for base in all_bases:
        doc_name = base[0].t_com.doc_od_style_file_name
        text_for_tags = ""
        for row in base:
            tag_list = row.el[TAGS].value
            tag_parser.add_context(str(tag_list), doc_name, 1)
    tag_parser.analyze(path_out_dir)

    open_dir(return_file_path)
    print(f"Успешно завершена проверка: <{return_file_path}>")

    #  3 Сравнение колва по уникальным кодам
    diff_base = RFQ.Value_Compare.start(all_bases)

    RFQ.Value_Compare.excel_base_out(diff_base,path_out_dir )


def rfq_file_start(file_path, dir_path):
    print("RFQ - проверка позиций")
    print("\t", file_path)

    print(f"    Open RFQ - {file_path}")
    t_com = TableComments(file_full_path=file_path,
                          tabel_class=TableComments.RFQ,
                          sheet_name=ColNames.RFQ.sheet_name)
    rfq_base = get_std_from_excel_file(t_com)

    mto_base = load_base()

    # 1 Сравнение RFQ с MTO
    input_base = base_vs_base_std(rfq_base, mto_base, [NAME, TYPE_MARK, UNITS, CODE, TAGS, MASS])

    column_dict = ColNames.RFQ.column_dict
    column_dict[99] = ROW_TYPE
    path_out_dir = utils.path.get_path_out_dir(dir_path)

    return_file_path = check_color_out(input_base,
                                       column_dict,
                                       path_out_dir,
                                       "RFQ_vs_MTO",
                                       excel_template_check_color=r"templates\out_template_check_rfq_color.xlsx")

    open_dir(return_file_path)
    print(f"Успешно завершена проверка: <{return_file_path}>")

