from types import NoneType

import base.t_comm_initial_classes as t_com_init_cls
import utils.path
from base.base_utils import str_remove_n_x000D_
from base.tables_columns import (
    ColNames, TYPE_MARK, CODE, NAME,
    TAGS, TAGS_2, VALUES, VALUES_2, VALUES_MTO, VALUES_VO,
    TAG_MTO, CODE_MTO, MTO_CODE_STRUCK, NAME_MTO, TYPE_MARK_MTO, NUMBERS_MTO, MTO_NAME,
    TAG_VO, CODE_VO, NAME_VO,
    MATCH_STATUS, MATCH_STATUS_VO, POSITION_STATUS, IN_CABINET, HAS_TAGS_RFP,
    ROW_TYPE,
    TAG_EFFECTIVE, MATCH_STATUS_TAGS, MATCH_STATUS_CODE_REPLACEMENT,
    MATCH_STATUS_VO_MTO_TAGS, MATCH_STATUS_RFP_MTO_TAGS,
    MATCH_STATUS_RFP_MTO_STRUCK_CODE, EQUIPMENT_TYPE_STATUS,
    MTO_CABINET_EQUIPMENT, DS_LOT, DS_ACTUAL,
    UL_ORDERED_VALUES, UL_VALUES, UL_UNITS, UL_REMAINING_VALUES,
    UL_COMPARE_STATUS, UL_DATA_STATUS, UL_TAG_MATCH_STATUS, UL_CODE, UL_SOURCE_FILES,
    UL_NAME, UL_TYPE_MARK, UL_VENDOR, UL_TAGS,
    RFP_SUPPLY_STATUS,
)
# Реэкспорт всех constants из tables_columns для модулей с "from base.base_classes import *"
from base.tables_columns import *  # noqa: F401, F403
from tags.tag_parser import get_tag
from utils.colors import Color
from utils.string_parsing import getOdStyleFileName

PROHIBITION_LIST = ("Не_используем", "Код не найден в Google базе")


class TableComments:
    def __init__(self,
                 dir_path="",
                 file_full_path="",
                 file_full_name="",
                 tabel_class=t_com_init_cls.MTO
                 ):
        self.tabel_class = None
        self.sheet_name = None
        self.tabel_type = None
        self.tabel_type_index = None  # Индекс для выбора комментариев при проверке позиций
        #
        self.column_dict = None
        self.dir_path = dir_path
        self.file_full_path = file_full_path
        self.file_name = file_full_name
        if file_full_name == "" and len(str(self.file_full_path)) > 5:
            self.file_name = utils.path.get_file_name_from_full_file_path(self.file_full_path)
        self.doc_od_style_file_name = getOdStyleFileName(self.file_name)
        #
        self.set_table_class(tabel_class)

    def set_table_class(self, tabel_class):
        self.tabel_class = tabel_class
        self.column_dict = tabel_class.column_dict
        self.sheet_name = tabel_class.sheet_name
        self.tabel_type = tabel_class.tabel_type
        self.tabel_type_index = tabel_class.tabel_type_index


class RowStd:
    def __init__(self, t_com=TableComments()):
        self.el = {}
        for att in ColNames.column_list:
            self.el[att] = CheckElement(None)
        self.row_type = None
        self.t_com = t_com

    def get_text_for_debug(self, columns=None):
        if columns is None:
            columns = [MTO_NAME, CODE, NAME, VALUES]
        out_text = ""
        for att in columns:
            out_text += f"|{self.el[att].value}|\t"
        return out_text

    def get_value(self, el_att):
        if el_att in (VALUES, VALUES_2, VALUES_MTO, VALUES_VO):
            value_temp = self.el[el_att].value
            try:
                self.el[el_att].value = float(self.el[el_att].value)
                return self.el[el_att].value
            except (ValueError, TypeError):
                try:
                    self.el[el_att].value = str(self.el[el_att].value).replace(",", ".")
                    self.el[el_att].value = float(self.el[el_att].value)
                    return self.el[el_att].value
                except (ValueError, TypeError):
                    pass
            self.el[el_att].value = value_temp
        return self.el[el_att].value

    @staticmethod
    def set_color_by_row_type(row):
        row_colors = {
            RowType.section_row: Color.section_row,
            RowType.position_row: Color.position_row,
            RowType.empty_row: Color.empty_row,
            RowType.other_row: Color.other_row,
            RowType.cabinet_title_row: Color.cabinet_title_row,
            RowType.system_row: Color.system_row
        }
        row_type = row.el[ROW_TYPE].value
        if row_type not in row_colors:
            print(f"ALERT ROW COLOR: {row.el[CODE].value}")
        for att in ColNames.column_list:
            if not isinstance(row.el[att], NoneType):
                if row_colors[row_type] != Color.no:
                    Color.set_el_color(row.el[att], row_colors[row_type])

    @staticmethod
    def set_color_by_row_type_in_list(in_list):
        for row in in_list:
            row_colors = {
                RowType.section_row: Color.section_row,
                RowType.position_row: Color.position_row,
                RowType.empty_row: Color.empty_row,
                RowType.other_row: Color.other_row,
                RowType.cabinet_title_row: Color.cabinet_title_row,
                RowType.system_row: Color.system_row,
                RowType.head_row: Color.head_row
            }
            row_type = row.el[ROW_TYPE].value
            if row_type not in row_colors:
                print(f"ALERT ROW COLOR: {row.el[CODE].value}")
            for att in ColNames.column_list:
                if not isinstance(row.el[att], NoneType):
                    if row_colors[row_type] != Color.no:
                        Color.set_el_color(row.el[att], row_colors[row_type])

    @staticmethod
    def get_std_check_row(dic, t_com: TableComments):
        """
        Запрос на создание стандартной строки RowStd[lbl.att] -> CTElement (value, color ...)
        :param t_com:
        :param row_type:
        :type dic: dict
        """
        row = RowStd()
        row.t_com = t_com

        #Список теговой которые должны инициализироваться как списки
        list_atts = [TAGS, TAGS_2, MTO_NAME]
        # Список теговой которые должны инициализироваться как списки
        list_tags = [TAGS, TAGS_2]

        for k, v in dic.items():
            if isinstance(v, CheckElement):
                if k in row.el:
                    if k in list_tags:
                        v.value= get_tag(str(v.value))
                    elif k == TYPE_MARK:
                        v.value = str_remove_n_x000D_(v.value)
                        v.struck_value = str_remove_n_x000D_(v.struck_value)
                        if isinstance(v.value, str):
                            v.value = v.value.strip()
                    elif k == NAME:
                        v.value = str_remove_n_x000D_(v.value)
                        v.struck_value = str_remove_n_x000D_(v.struck_value)
                        if isinstance(v.value, str):
                            v.value = v.value.strip()
                    elif k == CODE:
                        v.value = str_remove_n_x000D_(v.value)
                        v.struck_value = str_remove_n_x000D_(v.struck_value)
                        if isinstance(v.value, str):
                            v.value = v.value.strip()
                    if isinstance(v.value, str) and k not in [CODE, TYPE_MARK, NAME]:
                        v.value = v.value.strip()
                    v.color = Color.no
                    row.el[k] = v
            else:
                if k in row.el:
                    if k in list_tags:
                        v = get_tag(str(v))
                    elif k == TYPE_MARK:
                        v = str_remove_n_x000D_(v)
                    elif k == NAME:
                        v = str_remove_n_x000D_(v)
                    elif k == CODE:
                        v = str_remove_n_x000D_(v)
                    if isinstance(v, str):
                        v = v.strip()
                    row.el[k] = CheckElement(v, Color.no)
        # Заполняем остальные поля стандартного ряда пустыми значениями
        for att in ColNames.column_list:
            if isinstance(row.el[att].value, NoneType):
                v = ""
                if att in list_atts:
                    v = []
                row.el[att] = CheckElement(v, Color.no)

        row.row_type = RowType.get_row_type(row)
        row.el[ROW_TYPE].value = row.row_type
        return row

    @staticmethod
    def get_row_by_code(code, base) -> "RowStd" or None:
        for row in base:
            if row.el[CODE].value == code:
                return row
        # print(code, "no found (get_row_by_code)")
        return None

    @staticmethod
    def get_row_list_by_code(code, base) -> list['RowStd']:
        """Возвращает список объектов RowStd с указанным кодом"""
        return [row for row in base if row.el[CODE].value == code]

    @staticmethod
    def get_row_copy(in_row, t_com=TableComments()):
        out_row = RowStd()
        out_row.t_com = t_com
        out_row.row_type = in_row.row_type
        for k, v in in_row.el.items():
            out_row.el[k] = CheckElement(in_row.el[k].value, struck_value=in_row.el[k].struck_value)
        return out_row

    @staticmethod
    def get_row_copy_light(in_row, t_com=None, tags_override=None, values_override=None):
        """
        Облегчённая копия: разделяет el с исходной строкой, переопределяет только TAGS и VALUES.
        Экономит память при раскрытии строк (split по тегам/VALUES).

        Все столбцы из _COPY_COLS получают собственный CheckElement.
        Остальные столбцы разделяют ссылку с оригиналом (экономия памяти).

        _COPY_COLS включает ВСЕ столбцы, которые записываются при match или post-match:
          - VALUES*, TAG_MTO/VO, CODE_MTO/VO, NAME_MTO/VO и пр. — перезаписываются
            в add_mto_data_to_rfp_row / add_vo_data_to_rfp_row
          - TAG_EFFECTIVE, MATCH_STATUS_TAGS, EQUIPMENT_TYPE_STATUS и пр. — перезаписываются
            в post-match функциях (_set_effective_tags_before_export и т.д.)
          - MATCH_STATUS_CODE_REPLACEMENT — записывается во время matching
          - DS_LOT — записывается в _apply_lot_registry_to_rows

        Без копирования этих столбцов все light-копии из одного split разделяют
        один CheckElement и последняя запись перезаписывает значение для всех →
        массовые ложные дубликаты TAG_EFFECTIVE.
        """
        # Столбцы, которые перезаписываются при match или post-match обработке —
        # нужна своя копия CheckElement, иначе light-копии из одного split
        # будут разделять один объект и последняя запись перезапишет всех.
        _COPY_COLS = (
            VALUES, VALUES_2, VALUES_MTO, VALUES_VO,
            UNITS, UNITS_MTO, UNITS_CHECK_STATUS, UNITS_CONVERSION_TRACE,
            TAG_MTO, CODE_MTO, MTO_CODE_STRUCK, NAME_MTO, TYPE_MARK_MTO, NUMBERS_MTO,
            TAG_VO, CODE_VO, NAME_VO,
            MATCH_STATUS, MATCH_STATUS_VO, POSITION_STATUS, IN_CABINET, HAS_TAGS_RFP,
            TAG_EFFECTIVE, MATCH_STATUS_TAGS, MATCH_STATUS_CODE_REPLACEMENT,
            MATCH_STATUS_VO_MTO_TAGS, MATCH_STATUS_RFP_MTO_TAGS,
            MATCH_STATUS_RFP_MTO_STRUCK_CODE, EQUIPMENT_TYPE_STATUS,
            MTO_CABINET_EQUIPMENT, DS_LOT, DS_ACTUAL, RFP_SUPPLY_STATUS,
            UL_ORDERED_VALUES, UL_VALUES, UL_UNITS, UL_REMAINING_VALUES,
            UL_COMPARE_STATUS, UL_DATA_STATUS, UL_TAG_MATCH_STATUS, UL_CODE, UL_SOURCE_FILES,
            UL_NAME, UL_TYPE_MARK, UL_VENDOR, UL_TAGS,
        )
        t_com = t_com if t_com is not None else in_row.t_com
        out_row = object.__new__(RowStd)
        out_row.el = {}
        for k in in_row.el:
            if k in _COPY_COLS:
                src = in_row.el.get(k)
                if src is not None:
                    val = values_override if (k == VALUES and values_override is not None) else src.value
                    struck = getattr(src, "struck_value", "") if src else ""
                    out_row.el[k] = CheckElement(val, struck_value=struck)
                else:
                    out_row.el[k] = in_row.el[k]
            else:
                out_row.el[k] = in_row.el[k]
        out_row.t_com = t_com
        out_row.row_type = in_row.row_type
        if tags_override is not None:
            out_row.el[TAGS] = CheckElement(tags_override)
        if values_override is not None:
            src_vals = in_row.el.get(VALUES)
            struck = getattr(src_vals, "struck_value", "") if src_vals else ""
            out_row.el[VALUES] = CheckElement(values_override, struck_value=struck)
        return out_row

    @staticmethod
    def batch_copy_light(in_row, count, values_override=1):
        """
        Batch-создание count копий одной строки с values_override.
        Быстрее чем count вызовов get_row_copy_light за счёт
        предварительного разделения shared/copy столбцов.

        Для _COPY_COLS с пустым значением (value is None, struck_value пуст)
        создаётся один shared CheckElement на столбец вместо N отдельных —
        безопасно, т.к. batch_copy_light используется только для split MTO/VO
        строк, чьи CheckElement'ы не перезаписываются в pipeline.
        """
        _COPY_COLS = frozenset((
            VALUES, VALUES_2, VALUES_MTO, VALUES_VO,
            UNITS, UNITS_MTO, UNITS_CHECK_STATUS, UNITS_CONVERSION_TRACE,
            TAG_MTO, CODE_MTO, MTO_CODE_STRUCK, NAME_MTO, TYPE_MARK_MTO, NUMBERS_MTO,
            TAG_VO, CODE_VO, NAME_VO,
            MATCH_STATUS, MATCH_STATUS_VO, POSITION_STATUS, IN_CABINET, HAS_TAGS_RFP,
            TAG_EFFECTIVE, MATCH_STATUS_TAGS, MATCH_STATUS_CODE_REPLACEMENT,
            MATCH_STATUS_VO_MTO_TAGS, MATCH_STATUS_RFP_MTO_TAGS,
            MATCH_STATUS_RFP_MTO_STRUCK_CODE, EQUIPMENT_TYPE_STATUS,
            MTO_CABINET_EQUIPMENT, DS_LOT, DS_ACTUAL, RFP_SUPPLY_STATUS,
            UL_ORDERED_VALUES, UL_VALUES, UL_UNITS, UL_REMAINING_VALUES,
            UL_COMPARE_STATUS, UL_DATA_STATUS, UL_TAG_MATCH_STATUS, UL_CODE, UL_SOURCE_FILES,
            UL_NAME, UL_TYPE_MARK, UL_VENDOR, UL_TAGS,
        ))
        _ALWAYS_COPY_COLS = frozenset((
            UL_ORDERED_VALUES, UL_VALUES, UL_UNITS, UL_REMAINING_VALUES,
            UL_COMPARE_STATUS, UL_DATA_STATUS, UL_TAG_MATCH_STATUS, UL_CODE, UL_SOURCE_FILES,
            UL_NAME, UL_TYPE_MARK, UL_VENDOR, UL_TAGS,
        ))
        _CE = CheckElement
        shared = {}
        copy_sources = {}
        for k, v in in_row.el.items():
            if k in _COPY_COLS:
                if v is not None:
                    val = values_override if k == VALUES else v.value
                    struck = v.struck_value
                    if val is None and not struck and k not in _ALWAYS_COPY_COLS:
                        shared[k] = _CE(None)
                    else:
                        copy_sources[k] = (val, struck)
                else:
                    shared[k] = v
            else:
                shared[k] = v
        t_com = in_row.t_com
        row_type = in_row.row_type
        _new = object.__new__
        _cls = RowStd
        results = []
        _append = results.append
        for _ in range(count):
            out = _new(_cls)
            el = dict(shared)
            for k, (val, struck) in copy_sources.items():
                el[k] = _CE(val, struck_value=struck)
            out.el = el
            out.t_com = t_com
            out.row_type = row_type
            _append(out)
        return results

    def get_tags_list(self) -> list[str]:
        """Получает TAGS как список"""
        tags_value = self.get_value(TAGS) or ""
        if isinstance(tags_value, list):
            return tags_value
        elif isinstance(tags_value, str) and tags_value:
            # Разделяем по запятым или другим разделителям
            return [tag.strip() for tag in tags_value.split(',') if tag.strip()]
        return []

    def set_tags_list(self, tags_list: list[str]):
        """Устанавливает TAGS как список"""
        self.el[TAGS].value = tags_list

    def pop_tag(self, amount: int = 1) -> list[str]:
        """Извлекает указанное количество тегов из списка"""
        tags_list = self.get_tags_list()
        popped_tags = tags_list[:amount]
        remaining_tags = tags_list[amount:]
        self.set_tags_list(remaining_tags)
        return popped_tags

class CheckElement:
    def __init__(self, value, color=Color.no, struck_value=""):
        self.value = value  # Основное значение
        self.struck_value = struck_value  # Для чтения зачеркнутого текста из МТО
        self.color = color  # Цвет для вывода в результат
        self.comment = ""
        self.db_column = ""

    def set_color_safe(self, color):
        Color.set_el_color(self, color)


class RowType:
    section_row = "section_row"  # Заголовки разделов "Приборы приемно-контрольные", "Кабеленесущие конструкции" и т.п.
    cabinet_title_row = "cabinet_title_row"  # Заголовки шкафов и сборных позиций
    position_row = "position_row"  # Обычная строка с позицией
    empty_row = "empty_row"  # Пустые строки
    other_row = "other_row"  # Все что не попало в другие типы
    system_row = "system_row"  # Название подсистемы
    head_row = "head_row" # Заголовок табилцы

    row_types_list = {
        section_row,
        cabinet_title_row,
        empty_row,
        other_row,
        system_row,
        head_row
    }

    @staticmethod
    def get_row_type(row: RowStd):
        table_type = row.t_com.tabel_type
        #
        if table_type == t_com_init_cls.MTO.tabel_type:
            import base.get_row_type_variants.MTO
            return base.get_row_type_variants.MTO.get_row_type(row)
        #
        if table_type == t_com_init_cls.RFQ.tabel_type:
            import base.get_row_type_variants.RFQ
            return base.get_row_type_variants.RFQ.get_row_type(row)
        #
        if table_type == t_com_init_cls.RFP.tabel_type:
            import base.get_row_type_variants.RFP
            return base.get_row_type_variants.RFP.get_row_type(row)
        #
        if table_type == t_com_init_cls.GoogleBase.tabel_type:
            import base.get_row_type_variants.GoogleBase
            return base.get_row_type_variants.GoogleBase.get_row_type(row)
        #
        if table_type == t_com_init_cls.BOOT_CO.tabel_type:
            import base.get_row_type_variants.MTO
            return base.get_row_type_variants.MTO.get_row_type(row)
        #
        if table_type == t_com_init_cls.OUTPUT.tabel_type:
            import base.get_row_type_variants.MTO
            return base.get_row_type_variants.MTO.get_row_type(row)
        #
        if table_type == t_com_init_cls.ZipCompare.tabel_type:
            import base.get_row_type_variants.MTO
            return base.get_row_type_variants.MTO.get_row_type(row)
        #
        if table_type == t_com_init_cls.DsSpecification.tabel_type:
            import base.get_row_type_variants.ds_specification
            return base.get_row_type_variants.ds_specification.get_row_type(row)
        #
        if table_type == t_com_init_cls.RFQTPK.tabel_type:
            import base.get_row_type_variants.rfq_tpk
            return base.get_row_type_variants.rfq_tpk.get_row_type(row)
        #
        if table_type == t_com_init_cls.ListDsSpecificationVsMto.tabel_type:
            import base.get_row_type_variants.DsVsMto
            return base.get_row_type_variants.DsVsMto.get_row_type(row)
        #
        if table_type == t_com_init_cls.RFP_AGGREGATED.tabel_type:
            import base.get_row_type_variants.RFP_AGGREGATED
            return base.get_row_type_variants.RFP_AGGREGATED.get_row_type(row)
        #
        if table_type == t_com_init_cls.VO_MTO.tabel_type:
            import base.get_row_type_variants.VO_MTO
            return base.get_row_type_variants.VO_MTO.get_row_type(row)
        #
        if table_type == t_com_init_cls.EQUIPMENT.tabel_type:
            return RowType.position_row
        #
        if table_type == t_com_init_cls.BOE.tabel_type:
            import base.get_row_type_variants.BOE
            return base.get_row_type_variants.BOE.get_row_type(row)
        #
        if table_type == t_com_init_cls.BOM.tabel_type:
            import base.get_row_type_variants.BOM
            return base.get_row_type_variants.BOM.get_row_type(row)
        #
        if table_type == t_com_init_cls.BOQ.tabel_type:
            import base.get_row_type_variants.BOQ
            return base.get_row_type_variants.BOQ.get_row_type(row)
        #
        if table_type == t_com_init_cls.TsdPacking.tabel_type:
            import base.get_row_type_variants.TsdPacking
            return base.get_row_type_variants.TsdPacking.get_row_type(row)

        if table_type == t_com_init_cls.TsdPackingSoPl.tabel_type:
            import base.get_row_type_variants.TsdPackingSoPl
            return base.get_row_type_variants.TsdPackingSoPl.get_row_type(row)
        print(f"ROW TYPE ERROR: {row.el[CODE]}")
        raise


    @staticmethod
    def check_row(row, column_dict):
        """
        0 - должно быть пустым
        1 - должно быть НЕ пустым
        3 - не важно значение
        """
        for k, v in column_dict.items():
            val = row.el[k].value
            if isinstance(val, list):
                if len(val) > 0:
                    val = "1"
                else:
                    val = ""
            if str(val).strip() != "":
                c = 1
            else:
                c = 0
            if v != 3 and c != v:
                return False
        return True
