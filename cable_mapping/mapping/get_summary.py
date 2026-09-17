from cable_mapping.cm_clss.cm_CableTable import CableTable
from cable_mapping.mapping.map_calsses import *
from cable_mapping.mapping.summary_cables import NO_OUT_CABLES, SummaryRow
from utils.error_log import ErrorLog


def get_summary_from_cj_table(map_table: MapTable, cable_table_cls: CableTable):
    cable_table_cls.summary_table = {}  # Обнуляем таблицу
    stop_list_cables = []
    for ind, row in cable_table_cls.table.items():
        if isinstance(row, CableTableRow):
            sum_row = map_table.SummaryCables.get_summary_row_by_cable_cj_mark(row.el[CJ_CABLE_MARK].value)
            if isinstance(sum_row, SummaryRow):
                sum_row.add_length_value(row.el[CJ_TOTAL_LENGTH].value)
                cable_table_cls.sum_concatenate_summary_row(sum_row)
            elif sum_row == NO_OUT_CABLES:
                pass
            else:
                stop_list_cables.append(row.el[CJ_CABLE_MARK].value)
    if stop_list_cables:
        ErrorLog.add_error(f"ОШБКА: GET_SUMMARY_FROM_CJ_TABLE. Не найдены сочетания кабелей:")
        for _ in stop_list_cables:
            ErrorLog.add_error(f"   {_}")
        ErrorLog.exit()
