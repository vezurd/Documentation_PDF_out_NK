from base.base_classes import RowStd, RowType
from base.base_excel_out import *
from base.tables_columns import ColNames
from utils.string_parsing import print_att_list_table


class STDTable:
    @staticmethod
    def to_list(std_base: list[RowStd], column_dict=None, row_not_print_list=(RowType.empty_row), count_rows=0):
        # Если список выводимых столбцов не передан то выводим все столбцы из стандартного ряда
        if column_dict is None and std_base:
            column_dict = []
            row = std_base[0]
            column_dict = row.el.items()
        #
        out_list = []
        for row in std_base:
            if row.row_type not in row_not_print_list:
                out_row = []
                for k, v in column_dict.items():
                    out_row.append(row.el[v].value)
                out_list.append(out_row)
                # Если count_rows > 0, ограничиваем количество выводимых строк
                if count_rows > 0 and len(out_list) >= count_rows:
                    break
        return out_list

    @staticmethod
    def print_rows(std_base: list[RowStd], column_dict=None, row_not_print_list=(RowType.empty_row), count_rows=0):
        # Подсчитываем общее количество строк (после фильтрации)
        total_rows = sum(1 for row in std_base if row.row_type not in row_not_print_list)
        # Получаем список для вывода
        out_list = STDTable.to_list(std_base, column_dict=column_dict, row_not_print_list=row_not_print_list, count_rows=count_rows)
        # Количество выведенных строк
        printed_rows = len(out_list)
        # Формируем заголовок с информацией о количестве выведенных строк
        title = f"Выведено строк: {printed_rows} из {total_rows}"
        print_att_list_table(out_list, title=title, titles_print_flag=1)

    @staticmethod
    def get_title_list(std_base: list[RowStd]):
        out_row = []
        row = std_base[0]
        for key, v in row.el.items():
            out_row.append(key)
        return out_row

    @staticmethod
    def to_excel(base,
                 col_for_out_dict: dict,
                 out_dir,
                 file_prefix,
                 row_not_print_list=(
                         RowType.empty_row,
                 ),
                 excel_template_check_color=excel_template_check_color,
                 row_index=0,
                 comment_width=400,
                 comment_height=150,
                 show_row_type=False,
                 open_folder=False,
                 output_middle_name=None):
        return check_color_out(base,
                        col_for_out_dict,
                        out_dir,
                        file_prefix,
                        row_not_print_list,
                        excel_template_check_color,
                        row_index,
                        comment_width,
                        comment_height,
                        show_row_type=show_row_type,
                        open_folder=open_folder,
                        output_middle_name=output_middle_name)

    @staticmethod
    def to_excel_nanocad_db(base, out_dir, file_prefix):
        return STDTable.to_excel(base,
                          ColNames.NanocadDB.column_dict,
                          out_dir,
                          file_prefix,
                          excel_template_check_color=excel_template_nanocad_db)

    @staticmethod
    def to_excel_mto(base, out_dir, file_prefix):
        return STDTable.to_excel(base,
                          ColNames.MTO.column_dict,
                          out_dir,
                          file_prefix,
                          excel_template_check_color=excel_template_check_color)

    @staticmethod
    def to_excel_ds_spec(base, out_dir, file_prefix, show_row_type=False):
        return STDTable.to_excel(base,
                          ColNames.DsSpecification.column_dict,
                          out_dir,
                          file_prefix,
                          excel_template_check_color=excel_template_ds_spec,
                          show_row_type=show_row_type,
                          row_index=1,
                          row_not_print_list=(
                              RowType.empty_row,
                              RowType.other_row,
                              RowType.head_row,
                          )
                          )

    @staticmethod
    def to_excel_rfq_tpk(
        base,
        out_dir,
        file_prefix,
        show_row_type=False,
        output_middle_name=None,
    ):
        """Export RFQ TPK rows (position_row only) for load verification."""
        col_for_out = dict(ColNames.RFQTSpecification.column_dict)
        if show_row_type:
            col_for_out[99] = ROW_TYPE
        return STDTable.to_excel(
            base,
            col_for_out,
            out_dir,
            file_prefix,
            excel_template_check_color=excel_template_ds_spec,
            show_row_type=False,
            row_index=1,
            row_not_print_list=(
                RowType.empty_row,
                RowType.other_row,
                RowType.head_row,
            ),
            output_middle_name=output_middle_name,
        )

    @staticmethod
    def to_excel_ds_vs_mto_spec(
        base,
        out_dir,
        file_prefix,
        show_row_type=False,
        open_folder=False,
        output_middle_name=None,
    ):
        return STDTable.to_excel(base,
                          ColNames.DsVsMto.column_dict,
                          out_dir,
                          file_prefix,
                          excel_template_check_color=excel_template_ds_vs_mto,
                          show_row_type=show_row_type,
                          row_index=1,
                          row_not_print_list=(
                              RowType.empty_row,
                              RowType.other_row,
                              RowType.head_row,
                          ),
                          open_folder=open_folder,
                          output_middle_name=output_middle_name)

    @staticmethod
    def print_to_console(std_base: list[RowStd], column_dict=None):
        std_list = STDTable.to_list(std_base,
                                    column_dict=column_dict,
                                    row_not_print_list=(RowType.empty_row,
                                                        RowType.other_row,
                                                        RowType.head_row,)
                                    )
        file_name = std_base[0].t_com.file_name
        print_att_list_table(std_list, titles_print_flag=True, title=file_name)