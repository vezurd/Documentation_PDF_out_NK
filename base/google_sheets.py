import utils.path
from base.AVEVA import aveva_get_std_from_file, equipment_get_std_from_file
from base.base_mto import get_mto_std_from_file
from base.base_google import load_base, google_base_vs_base
from base.base_excel_out import check_color_out
from base.base_cheks import base_vs_base_std, base_vs_base_correction
from base.base_classes import *
from base.boot import base_co_get_std_from_file
from base.output import output_get_std_from_file
from GUI.gui_constants import GuiConst
from base.tables_columns import ColNames, ROW_TYPE, IN_CABINET, SECTION_TYPE
import base.t_comm_initial_classes as t_com_init_cls

def start(dir_path=-1, base_mode="MTO", file_path=-1, out_dir=None, cfg=None, correction_cfg=None):
    """
    01. Открытие МТО
        Ищем МТО в папке mto_path
        Заносим путь к найденному МТО в mto_file_path
    """
    return_file_path = None
    # 00 Получаем путь для сохранения результатов работы
    path_out_dir = out_dir or utils.path.get_path_out_dir(dir_path)
    # Папка создается перед выгрузкой файла

    """   
    01. Работа с базой Google Sheets 
    """
    t_com = TableComments(dir_path=dir_path,
                          file_full_path=file_path,
                          )
    code_base_data_std = load_base()
    if base_mode in GuiConst.dict:
        input_base = None
        if base_mode == GuiConst.MTO_DIR:
            t_com.set_table_class(t_com_init_cls.MTO)
            input_base = get_mto_std_from_file(t_com)
        elif base_mode == GuiConst.MTO_FILE:
            t_com.set_table_class(t_com_init_cls.MTO)
            t_com.dir_path = "-1"
            input_base = get_mto_std_from_file(t_com)

        # elif base_mode == GuiConst.BOOT_DIR:
        #     input_base = base_co_get_std_from_file(dir_path)
        elif base_mode == GuiConst.BOOT_FILE:
            # "Открыть файл CO.xlsx"
            t_com.set_table_class(t_com_init_cls.BOOT_CO)
            t_com.sheet_name = -1
            input_base = base_co_get_std_from_file(t_com)

        # elif base_mode == GuiConst.OUTPUT_DIR:
        #     input_base = output_get_std_from_file(dir_path)
        elif base_mode == GuiConst.OUTPUT_FILE:
            t_com.set_table_class(t_com_init_cls.MTO)
            t_com.sheet_name = "СС"
            input_base = output_get_std_from_file(t_com)

        elif base_mode == GuiConst.AVEVA_FILE:
            t_com.set_table_class(t_com_init_cls.AVEVA)
            t_com.sheet_name = -1
            input_base = aveva_get_std_from_file(t_com)

        elif base_mode == GuiConst.EQUIPMENT_FILE:
            t_com.set_table_class(t_com_init_cls.EQUIPMENT)
            input_base = equipment_get_std_from_file(t_com)
        # elif base_mode == GuiConst.NANOCAD_DB:
        #     input_base = nanocad_db_get_std_from_file(file_path=file_path)

        # Finally
        if input_base:
            columns_for_check = GuiConst.dict[base_mode][3]

            input_base_copy = [RowStd.get_row_copy(row, row.t_com) for row in input_base]
            input_base = base_vs_base_std(input_base, code_base_data_std, columns_for_check, cfg=cfg)
            column_dict = ColNames.MTO.column_dict
            column_dict[99] = ROW_TYPE
            column_dict[100] = IN_CABINET
            column_dict[101] = SECTION_TYPE
            return_file_path = check_color_out(input_base,
                                               ColNames.MTO.column_dict,
                                               path_out_dir,
                                               GuiConst.dict[base_mode][2],
                                               date_stamp=True,
                                               row_index=1,
                                               row_not_print_list=(RowType.empty_row,RowType.head_row)
                                               )
            corr_cfg = correction_cfg if isinstance(correction_cfg, dict) else cfg
            input_base_copy = base_vs_base_correction(input_base_copy, code_base_data_std, columns_for_check, cfg=corr_cfg)
            check_color_out(input_base_copy,
                            ColNames.MTO.column_dict,
                            path_out_dir,
                            "correction_" + GuiConst.dict[base_mode][2],
                            date_stamp=True,
                            row_index=1,
                            row_not_print_list=(RowType.empty_row,RowType.head_row)
                            )
            if base_mode == GuiConst.EQUIPMENT_FILE:
                google_base_vs_base(input_base)
        else:
            print("Ошибка: база input_base не загружена <google_sheets>")
            return None
    else:
        print(f"Не найден режим работы программы: base_mode - <{base_mode}>")
        return None
    return return_file_path
