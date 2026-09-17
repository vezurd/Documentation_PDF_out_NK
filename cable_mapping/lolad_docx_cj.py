from cable_mapping.cm_clss.cm_ColumnsCableTable import ColumnsCableTable
from cable_mapping.cm_clss.cm_CableTable import CableTable
from cable_mapping.cm_clss.cm_TableCommentsCabMap import DocumentCabMap
from cable_mapping.cm_clss.cm_CableTableRow import CableTableRow
from docx import Document
from cable_mapping.constants import *
from cable_mapping.mapping.summary_cables import SummaryRow

from utils.string_parsing import print_att_list_table

docx_cj_dict = {
    0: CJ_CABLE_LABEL,
    1: CJ_CABLE_FULL_TAG,
    2: CJ_FROM,
    3: CJ_FROM_TERMINAL,
    4: CJ_WHERE,
    5: CJ_WHERE_TERMINAL,
    6: CJ_CABLE_MARK,
    7: CABLE_LAYING_IN_CABIN,
    8: CABLE_LAYING_IN_TRAY,
    9: CABLE_LAYING_IN_CABLE_CHANEL,
    10: CABLE_LAYING_IN_OPEN,
    11: CABLE_LAYING_IN_FLUTED,
    12: CABLE_LAYING_IN_METAL_HOUSE,
    13: CABLE_LAYING_IN_TUBE,
    14: CABLE_LAYING_IN_TRENCH,
    15: CJ_TOTAL_LENGTH,
    16: CJ_ANNOTAION
}


def load_cj_from_docx(doc: DocumentCabMap, cj_exist_cable_cls: CableTable, dbg=0, start_index=2, max_len=25):
    print(f"\tОткрываем |{doc.file_full_path}| <load_cj_from_docx>")

    doc_docx = Document(doc.file_full_path)
    # последовательность всех таблиц документа
    all_tables = doc_docx.tables
    print('Всего таблиц в документе:', len(all_tables))

    # создаем пустой словарь под данные таблиц
    data_tables = {i: None for i in range(len(all_tables))}
    # проходимся по таблицам
    for i, table in enumerate(all_tables):
        print('\nДанные таблицы №', i)
        # создаем список строк для таблицы `i` (пока пустые)
        data_tables[i] = [[] for _ in range(len(table.rows))]
        # проходимся по строкам таблицы `i`
        for j, row in enumerate(table.rows):
            # проходимся по ячейкам таблицы `i` и строки `j`
            for cell in row.cells:
                # добавляем значение ячейки в соответствующий
                # список, созданного словаря под данные таблиц
                data_tables[i][j].append(cell.text)

    for k, v in data_tables.items():
        if k == 0:
            summ_list_raw = v
            summ_column_dict = {}
            for ind in range(len(summ_list_raw[0])):
                title = summ_list_raw[0][ind]
                if title in summ_column_dict:
                    pass
                else:
                    summ_column_dict[title] = ind

            summ_list = []
            for x in range(len(summ_list_raw)):
                sum_row = []
                for y in range(len(summ_list_raw[x])):
                    for name, col_ind in summ_column_dict.items():
                        if y == col_ind:
                            sum_row.append(summ_list_raw[x][y])
                summ_list.append(sum_row)
            print_att_list_table(summ_list, max_len=40, title="load_cj_from_docx.summ_list")

            # На базе сырой таблицы анализируем строки и добавляем их в cj_exist_cable_cls
            r_system = doc.t_com.system
            for row in summ_list:
                check_result = check_row_type(row)
                if check_result == 0:
                    pass
                elif check_result == 1:
                    sum_row = SummaryRow(row[0] + " " + row[1],
                                         row[0],
                                         row[1],
                                         value=row[2],
                                         r_system=r_system)
                    cj_exist_cable_cls.sum_add_summary_row(sum_row)
                else:
                    r_system = check_result

        if k == 1:
            summ_list_raw = v
            summ_column_dict = {}
            for ind in range(len(summ_list_raw[2])):
                title = summ_list_raw[2][ind]
                if title in summ_column_dict:
                    pass
                else:
                    summ_column_dict[title] = ind

            summ_list = []
            for x in range(len(summ_list_raw)):
                sum_row = []
                for y in range(len(summ_list_raw[x])):
                    for name, col_ind in summ_column_dict.items():
                        if y == col_ind:
                            sum_row.append(summ_list_raw[x][y])
                summ_list.append(sum_row)
            # На базе сырой таблицы анализируем строки и добавляем их в cj_exist_cable_cls
            r_system = doc.t_com.system
            cable_dict = {
                CJ_CABLE_LABEL: 0,
                CJ_CABLE_FULL_TAG: 1,
                CJ_FROM: 2,
                CJ_FROM_TERMINAL: 3,
                CJ_WHERE: 4,
                CJ_WHERE_TERMINAL: 5,
                CJ_CABLE_MARK: 6,
                CABLE_LAYING_IN_CABIN: 7,
                CABLE_LAYING_IN_TRAY: 8,
                CABLE_LAYING_IN_CABLE_CHANEL: 9,
                CABLE_LAYING_IN_OPEN: 10,
                CABLE_LAYING_IN_FLUTED: 11,
                CABLE_LAYING_IN_METAL_HOUSE: 12,
                CABLE_LAYING_IN_TUBE: 13,
                CABLE_LAYING_IN_TRENCH: 14,
                CJ_TOTAL_LENGTH: 15,
                CJ_ANNOTAION: 16
            }
            for row in summ_list:
                check_result = check_row_type(row)
                if check_result == 0:
                    pass
                elif check_result == 1:
                    cable_row = CableTableRow(row_type="data",
                                              t_com=doc.t_com,
                                              system=r_system)
                    for att in ColumnsCableTable.main_list:
                        if att in cable_dict:
                            cable_row.el[att].value = row[cable_dict[att]]

                    cj_exist_cable_cls.add_std_row(cable_row)
                else:
                    r_system = check_result
            print_att_list_table(summ_list, max_len=40, title="load_cj_from_docx.table")
    # print('Данные всех таблиц документа:')
    # print(data_tables)


def check_row_type(row):
    r_system = 0
    pass_list = ["Обозначение кабеля, провода", "Марка", "1"]
    if row[0] in pass_list:
        pass
    elif str(row[0]).strip() == "":
        pass
    elif "Система контроля и управления доступом" in row[0] or "СКУД" in row[0]:
        r_system = "SKUD"
    elif "Система охранной" in row[0] or "СОС" in row[0]:
        r_system = "SOS"
    elif "Система охранного телевидения" in row[0] or "СОТ" in row[0]:
        r_system = "SOT"
    else:
        r_system = 1
    return r_system
