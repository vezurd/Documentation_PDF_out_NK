from cable_mapping.constants import *

UNITS = "units"
CODE = "code"
TAGS = "tags"
VALUES = "values"
VENDOR = "vendor"
MASS = "mass"
ANNOTATION = "annotation"
NUMBERS = "numbers"
NOMENCLATURE = "nomenclature"
TYPE_MARK = "type_mark"
NAME = "name"
PROHIBITION = "prohibition"
REPLACEMENT = "replacement"
EQUIPMENT_CODE = "equipment_code"
ROW_TYPE = "row_type"
TITLE = "TITLE"
SPECIFICATION_NAME = "SPECIFICATION_NAME"
#
UNITS_2 = "units_2"
TAGS_2 = "tags_2"
VALUES_2 = "values_2"
VENDOR_2 = "vendor_2"
MASS_2 = "mass_2"
ANNOTATION_2 = "annotation_2"
ANNOTATION_3 = "annotation_3"
NUMBERS_2 = "numbers_2"
ANNOTATION_MTO = "annotation_mto"  # MTO spec notes (ДС vs MTO export), distinct from ANNOTATION_2/3
NOMENCLATURE_2 = "nomenclature_2"
TYPE_MARK_2 = "type_mark_2"
NAME_2 = "name_2"
RFP_SUPPLY_STATUS = "rfp_supply_status"  # Parts header «MTO, Статус позиции»; not POSITION_STATUS
PROHIBITION_2 = "prohibition_2"
CODE_2 = "code_2"
#
DIFF_FLAG_TAG = "diff_flag_TAG"
DIFF_FLAG_VALUE = "diff_flag_Value"
DIFF_FLAG_NAME = "diff_flag_name"
DIFF_FLAG_TYPE = "diff_flag_Type"
DIFF_FLAG_SUMMARY = "diff_flag_summary"
DIFF_FLAG_MATCH = "diff_flag_Match"
MTO_NAME = "mto_name"
ZIP_CODE = "zip_code"
ZIP_VALUE = "zip_value"
ZIP_BUILDING = "ZIP_BUILDING"
BD_NAME = "BD_NAME"
BD_TABLE_NAME = "BD_TABLE_NAME"
BD_COLUMN_NAME = "BD_COLUMN_NAME"
BD_ID = "BD_ID"
# Для ДС
DS_NAME = "DS_NAME"  # Само наименование ДС , ДС27
DS_ACTUAL = "DS_ACTUAL"  # Фактический номер ДС (92 в ДС92_24Б / папка согл УЛ ДС92)
DS_MANAGER = "DS_MANAGER"  # Фамилия МП, ведущего ДС
DS_LOT = "DS_LOT"  # Номер лота из реестра
PATH_RFP = "PATH_RFP"  # UNC path to source RFP/DS workbook
PATH_MTO = "PATH_MTO"  # UNC path to source MTO workbook
DS_NUMBER = "DS_NUMBER"
DS_TITLE = "DS_TITLE"
DS_SYSTEM = "DS_SYSTEM"
DS_SPECIFICATION = "DS_SPECIFICATION"
DS_RFQ = "DS_RFQ"
DS_CODE_1C = "DS_CODE_1C"
DS_NAME_BY_RFQ = "DS_NAME_BY_RFQ"
# Новые константы для столбцов с ценами
DS_UNIT_PRICE_EXCL_VAT = "DS_UNIT_PRICE_EXCL_VAT"  # Цена за Позицию Товара (RUB) без НДС
DS_PACKAGING_PRICE_EXCL_VAT = "DS_PACKAGING_PRICE_EXCL_VAT"  # Цена (RUB) за упаковку без НДС
DS_SHIPPING_PRICE_EXCL_VAT = "DS_SHIPPING_PRICE_EXCL_VAT"  # Цена (RUB) за транспортировку без НДС
DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT = "DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT"  # Цена за Позицию Товара с учетом упаковки и транспортировки (RUB) без НДС
DS_TOTAL_PRICE_EXCL_VAT = "DS_TOTAL_PRICE_EXCL_VAT"  # Цена за Позиции Товара с учетом упаковки и транспортировки (RUB) без НДС
DS_VAT_AMOUNT = "DS_VAT_AMOUNT"  # (RUB) НДС - 20%
DS_TOTAL_PRICE_INCL_VAT = "DS_TOTAL_PRICE_INCL_VAT"  # Цена Позиций Товара (RUB) с НДС 20%
DS_DELIVERY_TIME = "DS_DELIVERY_TIME"  # Срок поставки
# Grouped DS vs MTO vs RFQ (этап 2)
MTO_DS_DIFF = "mto_ds_diff"  # VALUES_2 − VALUES (Спека − ДС)
RFQ_VALUES = "rfq_values"  # Кол-во из RFQ TPK после группировки
RFQ_UNITS = "rfq_units"  # ЕИ из RFQ TPK
MTO_DS_RFQ_DIFF = "mto_ds_rfq_diff"  # MTO_DS_DIFF − RFQ_VALUES
RFQ_COMPARE_STATUS = "rfq_compare_status"  # Статус сравнения с RFQ
DELIVERY_OVERALL_STATUS = "delivery_overall_status"  # Общий статус поставки (все строки)
RFQ_NAME = "rfq_name"  # Наименование МТР из RFQ
RFQ_CODE = "rfq_code"  # Код РД / BCC из RFQ
RFQ_TYPE_MARK = "rfq_type_mark"  # Технические характеристики RFQ
# Grouped delivery comparison with packing lists (УЛ)
UL_ORDERED_VALUES = "ul_ordered_values"  # VALUES + RFQ_VALUES
UL_VALUES = "ul_values"  # Accumulated delivered quantity from packing lists
UL_UNITS = "ul_units"
UL_REMAINING_VALUES = "ul_remaining_values"  # ordered − delivered
UL_COMPARE_STATUS = "ul_compare_status"
UL_DATA_STATUS = "ul_data_status"
UL_TAG_MATCH_STATUS = "ul_tag_match_status"
UL_CODE = "ul_code"
UL_SOURCE_FILES = "ul_source_files"  # file · sheet · Excel row
UL_NAME = "ul_name"
UL_TYPE_MARK = "ul_type_mark"
UL_VENDOR = "ul_vendor"
UL_TAGS = "ul_tags"
#
IN_CABINET = "IN_CABINET"  # Поле для внесения к какому шкафу относится строка
SECTION_TYPE = "section_type"  # Тип секции MTO (из section_row, для маппинга → BOE/BOM)
# Столбцы для данных из MTO при сопоставлении с RFP
TAG_MTO = "tag_mto"  # Тег из MTO
CODE_MTO = "code_mto"  # Код из MTO
MTO_CODE_STRUCK = "mto_code_struck"  # Зачеркнутый код из MTO
NAME_MTO = "name_mto"  # Наименование из MTO
TYPE_MARK_MTO = "type_mark_mto"  # Тип марки из MTO
NUMBERS_MTO = "numbers_mto"  # Номера из MTO
VALUES_MTO = "values_mto"  # Количество из MTO
UNITS_MTO = "units_mto"  # Ед. изм. из MTO
UNITS_CHECK_STATUS = "units_check_status"  # Статус проверки/конвертации ед. изм.
UNITS_CONVERSION_TRACE = "units_conversion_trace"  # Trace конвертации ед. изм.
VALUES_MTO_RFP_DIFF = "values_mto_rfp_diff"  # VALUES_MTO − VALUES при совместимых ЕИ
VALUES_MTO_UL_DIFF = "values_mto_ul_diff"  # VALUES_MTO − UL_VALUES при совместимых ЕИ
MATCH_STATUS = "match_status"  # Статус сопоставления: "Тег сопоставлен", "Не найден в МТО", "Добавлен из МТО"
HAS_TAGS_RFP = "has_tags_rfp"  # Наличие тегов в RFP до сопоставления: "Да", "Нет"
POSITION_STATUS = "position_status"  # Статус позиции: "Исключен", "Не протегирован в МТО", "Не протегирован в RFP"
# Статус итогового сравнения по тегам (MTO/VO)
MATCH_STATUS_TAGS = "match_status_tags"  # Итог: "Тег найден", "Тег не найден", "Нет тегов"
MATCH_STATUS_CODE_REPLACEMENT = "match_status_code_replacement"  # Статус замен по кодам
MATCH_STATUS_VO_MTO_TAGS = "match_status_vo_mto_tags"  # Сравнение тегов VO и MTO
MATCH_STATUS_RFP_MTO_TAGS = "match_status_rfp_mto_tags"  # Сравнение тегов RFP и MTO
MATCH_STATUS_RFP_MTO_STRUCK_CODE = "match_status_rfp_mto_struck_code"  # Сравнение CODE и зачеркнутого MTO
EQUIPMENT_TYPE_STATUS = "equipment_type_status"  # Тип оборудования по тегам
TAG_EFFECTIVE = "tag_effective"  # Тег для вывода (VO/MTO/RFP)
MTO_CABINET_EQUIPMENT = "mto_cabinet_equipment"  # MTO, Шкафное оборудовани?
# Столбцы для данных из VO при сопоставлении с RFP
TAG_VO = "tag_vo"  # Тег из VO
CODE_VO = "code_vo"  # Код из VO
NAME_VO = "name_vo"  # Наименование из VO
VALUES_VO = "values_vo"  # Количество из VO
MATCH_STATUS_VO = "match_status_vo"  # Статус сопоставления с VO: "Тег сопоставлен", "Не найден в VO"
#
# BBB (BOM/BOE/BOQ) specific columns
BBB_MARKA = "bbb_marka"
BBB_SPEC_NUMBER = "bbb_spec_number"
BBB_REVISION = "bbb_revision"
BBB_OBJECT_NAME = "bbb_object_name"
BBB_MTR_TYPE = "bbb_mtr_type"
BBB_MTR_GROUP = "bbb_mtr_group"
BBB_WORK_CODE_LIST = "bbb_work_code_list"
BBB_SUPPLY_ZONE = "bbb_supply_zone"
BBB_WORK_CODE = "bbb_work_code"
BBB_POSITION = "bbb_position"
BBB_CONSUMPTION_RATE = "bbb_consumption_rate"
BBB_UNIT_PRICE = "bbb_unit_price"
BBB_TOTAL_PRICE = "bbb_total_price"
BBB_ASSEMBLY = "bbb_assembly"
# BOQ specific
BBB_WORK_NAME = "bbb_work_name"
BBB_LABOR_UNIT = "bbb_labor_unit"
BBB_LABOR_TOTAL = "bbb_labor_total"
BBB_EQUIPMENT_UNIT = "bbb_equipment_unit"
BBB_EQUIPMENT_TOTAL = "bbb_equipment_total"
BBB_LABOR_RATE = "bbb_labor_rate"
BBB_EQUIPMENT_RATE = "bbb_equipment_rate"
BBB_DIRECT_COST_RATE = "bbb_direct_cost_rate"
BBB_DIRECT_COST_TOTAL = "bbb_direct_cost_total"

class ColNames:
    """
    Это основной список атрибутов для класса RowStd (стандартный ряд таблиц)
    Все столбцы используемые в работе должны быть занесены сюда
    """
    column_list = {
        UNITS,
        CODE,
        TAGS,
        VALUES,
        VENDOR,
        MASS,
        ANNOTATION,
        NUMBERS,
        NOMENCLATURE,
        TYPE_MARK,
        NAME,
        PROHIBITION,
        REPLACEMENT,
        EQUIPMENT_CODE,
        ROW_TYPE,
        TITLE,
        SPECIFICATION_NAME,
        G_BASE_CABLE_LAYING_FLAG,
        G_BASE_CABLE_LAYING_TYPE,
        G_BASE_CABLE_LAYING_VARIANTS,
        G_BASE_CABLE_SUMMARY_VARIANTS,
        UNITS_2,
        TAGS_2,
        VALUES_2,
        VENDOR_2,
        MASS_2,
        ANNOTATION_2,
        ANNOTATION_3,
        NUMBERS_2,
        ANNOTATION_MTO,
        NOMENCLATURE_2,
        TYPE_MARK_2,
        NAME_2,
        RFP_SUPPLY_STATUS,
        PROHIBITION_2,
        CODE_2,
        DIFF_FLAG_TAG,
        DIFF_FLAG_VALUE,
        DIFF_FLAG_NAME,
        DIFF_FLAG_TYPE,
        DIFF_FLAG_SUMMARY,
        DIFF_FLAG_MATCH,
        ZIP_CODE,
        ZIP_VALUE,
        ZIP_BUILDING,
        MTO_NAME,
        BD_NAME,
        BD_TABLE_NAME,
        BD_COLUMN_NAME,
        BD_ID,
        DS_NUMBER,
        DS_TITLE,
        DS_SYSTEM,
        DS_SPECIFICATION,
        DS_RFQ,
        DS_CODE_1C,
        DS_NAME_BY_RFQ,
        DS_UNIT_PRICE_EXCL_VAT,
        DS_PACKAGING_PRICE_EXCL_VAT,
        DS_SHIPPING_PRICE_EXCL_VAT,
        DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT,
        DS_TOTAL_PRICE_EXCL_VAT,
        DS_VAT_AMOUNT,
        DS_TOTAL_PRICE_INCL_VAT,
        DS_DELIVERY_TIME,
        DS_NAME,
        DS_ACTUAL,
        DS_MANAGER,
        DS_LOT,
        PATH_RFP,
        PATH_MTO,
        MTO_DS_DIFF,
        RFQ_VALUES,
        RFQ_UNITS,
        MTO_DS_RFQ_DIFF,
        RFQ_COMPARE_STATUS,
        DELIVERY_OVERALL_STATUS,
        RFQ_NAME,
        RFQ_CODE,
        RFQ_TYPE_MARK,
        UL_ORDERED_VALUES,
        UL_VALUES,
        UL_UNITS,
        UL_REMAINING_VALUES,
        UL_COMPARE_STATUS,
        UL_DATA_STATUS,
        UL_TAG_MATCH_STATUS,
        UL_CODE,
        UL_SOURCE_FILES,
        UL_NAME,
        UL_TYPE_MARK,
        UL_VENDOR,
        UL_TAGS,
        #
        IN_CABINET,
        SECTION_TYPE,
        # Столбцы для данных из MTO при сопоставлении с RFP
        TAG_MTO,
        CODE_MTO,
        MTO_CODE_STRUCK,
        NAME_MTO,
        TYPE_MARK_MTO,
        NUMBERS_MTO,
        VALUES_MTO,
        UNITS_MTO,
        UNITS_CHECK_STATUS,
        UNITS_CONVERSION_TRACE,
        VALUES_MTO_RFP_DIFF,
        VALUES_MTO_UL_DIFF,
        MATCH_STATUS,
        HAS_TAGS_RFP,
        POSITION_STATUS,
        MATCH_STATUS_TAGS,
        MATCH_STATUS_CODE_REPLACEMENT,
        MATCH_STATUS_VO_MTO_TAGS,
        MATCH_STATUS_RFP_MTO_TAGS,
        MATCH_STATUS_RFP_MTO_STRUCK_CODE,
        EQUIPMENT_TYPE_STATUS,
        TAG_EFFECTIVE,
        MTO_CABINET_EQUIPMENT,
        TAG_VO,
        CODE_VO,
        NAME_VO,
        VALUES_VO,
        MATCH_STATUS_VO,
        # BBB (BOM/BOE/BOQ)
        BBB_MARKA,
        BBB_SPEC_NUMBER,
        BBB_REVISION,
        BBB_OBJECT_NAME,
        BBB_MTR_TYPE,
        BBB_MTR_GROUP,
        BBB_WORK_CODE_LIST,
        BBB_SUPPLY_ZONE,
        BBB_WORK_CODE,
        BBB_POSITION,
        BBB_CONSUMPTION_RATE,
        BBB_UNIT_PRICE,
        BBB_TOTAL_PRICE,
        BBB_ASSEMBLY,
        BBB_WORK_NAME,
        BBB_LABOR_UNIT,
        BBB_LABOR_TOTAL,
        BBB_EQUIPMENT_UNIT,
        BBB_EQUIPMENT_TOTAL,
        BBB_LABOR_RATE,
        BBB_EQUIPMENT_RATE,
        BBB_DIRECT_COST_RATE,
        BBB_DIRECT_COST_TOTAL,
    }

    class GoogleBase:
        column_dict = {
            0: NAME,
            1: TYPE_MARK,
            2: NOMENCLATURE,
            3: UNITS,
            4: CODE,
            9: ANNOTATION,
            10: PROHIBITION,
            11: VENDOR,
            14: MASS,
            16: REPLACEMENT,
            17: EQUIPMENT_CODE,
            18: G_BASE_CABLE_LAYING_FLAG,
            19: G_BASE_CABLE_LAYING_TYPE,
            20: G_BASE_CABLE_LAYING_VARIANTS,
            21: G_BASE_CABLE_SUMMARY_VARIANTS,
            25: ZIP_CODE,
            26: BBB_WORK_CODE_LIST,
            27: BBB_MTR_GROUP,
        }
        sheet_name = None

    class MTO:
        column_dict = {
            0: TAGS,
            1: NUMBERS,
            2: NAME,
            3: TYPE_MARK,
            4: CODE,
            5: VENDOR,
            6: UNITS,
            7: VALUES,
            8: MASS,
            9: ANNOTATION
        }
        sheet_name = "Спецификация"

    class DsZin:
        column_dict = {
            0: DS_NUMBER,
            1: NUMBERS,
            2: DS_TITLE,
            3: DS_SYSTEM,
            4: DS_SPECIFICATION,
            5: DS_RFQ,
            6: NAME,
            7: DS_CODE_1C,
            8: CODE,
            9: DS_NAME_BY_RFQ,
            10: TYPE_MARK,
            11: UNITS,
            12: VALUES,
            21: VENDOR,
        }
        sheet_name = "Сводная по ДС с НДС"

    class DsSpecification:
        column_dict = {
            0: DS_NUMBER,
            1: DS_TITLE,
            2: DS_SYSTEM,
            3: DS_SPECIFICATION,
            4: DS_RFQ,
            5: NAME,
            6: DS_CODE_1C,
            7: CODE,
            8: DS_NAME_BY_RFQ,
            9: TYPE_MARK,
            10: UNITS,
            11: VALUES,
            12: DS_UNIT_PRICE_EXCL_VAT,
            13: DS_PACKAGING_PRICE_EXCL_VAT,
            14: DS_SHIPPING_PRICE_EXCL_VAT,
            15: DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT,
            16: DS_TOTAL_PRICE_EXCL_VAT,
            17: DS_VAT_AMOUNT,
            18: DS_TOTAL_PRICE_INCL_VAT,
            19: DS_DELIVERY_TIME,
            20: VENDOR,
        }
        sheet_name = -1

    class ListDsSpecification:
        column_dict = {
            0 : DS_NAME,
            1: DS_NUMBER,
            2: DS_TITLE,
            3: DS_SYSTEM,
            4: DS_SPECIFICATION,
            5: DS_RFQ,
            6: NAME,
            7: DS_CODE_1C,
            8: CODE,
            9: DS_NAME_BY_RFQ,
            10: TYPE_MARK,
            11: UNITS,
            12: VALUES,
            13: DS_UNIT_PRICE_EXCL_VAT,
            14: DS_PACKAGING_PRICE_EXCL_VAT,
            15: DS_SHIPPING_PRICE_EXCL_VAT,
            16: DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT,
            17: DS_TOTAL_PRICE_EXCL_VAT,
            18: DS_VAT_AMOUNT,
            19: DS_TOTAL_PRICE_INCL_VAT,
            20: DS_DELIVERY_TIME,
            21: VENDOR,
        }
        sheet_name = -1

    # RFQTPK 27/05/2026 — индексы = 0-based номер столбца в xlsx (см. строку шапки ~row 23).
    class RFQTSpecification:
        column_dict = {
            # 0: служебный № строки источника (не используем)
            1: NUMBERS,  # № п/п
            2: DS_TITLE,  # Титул
            3: DS_SYSTEM,  # Раздел / марка
            4: DS_SPECIFICATION,  # Спецификация
            5: TAGS,  # Линия, Tag-номер
            6: DS_CODE_1C,  # Код 1С СОУ
            7: CODE,  # Код РД / BCC
            8: NAME,  # Наименование МТР
            9: TYPE_MARK,  # Технические характеристики
            10: VALUES,  # Кол-во
            11: UNITS,  # Ед. изм.
        }
        sheet_name = -1

    class DsVsMto:
        """
        Словарь для вывода в формате сравнения ДС и МТО
        """
        column_dict = {
            0: DS_NAME,
            1: DS_NUMBER,
            2: DS_TITLE,
            3: DS_SYSTEM,
            4: DS_SPECIFICATION,
            5: DS_RFQ,
            6: NAME,
            7: DS_CODE_1C,
            8: CODE,
            9: DS_NAME_BY_RFQ,
            10: TYPE_MARK,
            11: UNITS,
            12: VALUES,
            13: DS_UNIT_PRICE_EXCL_VAT,
            14: DS_PACKAGING_PRICE_EXCL_VAT,
            15: DS_SHIPPING_PRICE_EXCL_VAT,
            16: DS_UNIT_PRICE_WITH_EXTRAS_EXCL_VAT,
            17: DS_TOTAL_PRICE_EXCL_VAT,
            18: DS_VAT_AMOUNT,
            19: DS_TOTAL_PRICE_INCL_VAT,
            20: DS_DELIVERY_TIME,
            21: VENDOR,
            22: ANNOTATION_2,
            23: ANNOTATION_3,
            24: IN_CABINET,
            25: TAGS_2,
            26: NUMBERS_2,
            27: ANNOTATION_MTO,
            28: NAME_2,
            29: TYPE_MARK_2,
            30: CODE_2,
            31: VENDOR_2,
            32: UNITS_2,
            33: VALUES_2,
            34: MTO_DS_DIFF,
            35: RFQ_VALUES,
            36: RFQ_UNITS,
            37: MTO_DS_RFQ_DIFF,
            38: RFQ_COMPARE_STATUS,
            39: DELIVERY_OVERALL_STATUS,
            40: RFQ_NAME,
            41: RFQ_CODE,
            42: RFQ_TYPE_MARK,
            43: UL_ORDERED_VALUES,
            44: UL_VALUES,
            45: UL_UNITS,
            46: UL_REMAINING_VALUES,
            47: UL_COMPARE_STATUS,
            48: UL_DATA_STATUS,
            49: UL_CODE,
            50: UL_SOURCE_FILES,
            51: UL_NAME,
            52: UL_TYPE_MARK,
            53: UL_VENDOR,
            54: UL_TAGS,
        }
        sheet_name = -1

    class Output:
        column_dict = {
            0: TAGS,
            1: NUMBERS,
            2: NAME,
            3: TYPE_MARK,
            4: CODE,
            5: VENDOR,
            6: UNITS,
            7: VALUES,
            8: MASS,
            9: ANNOTATION
        }
        sheet_name = "СС"

    class RFQ:
        column_dict = {
            2: NUMBERS,
            3: TITLE,
            5: SPECIFICATION_NAME,
            6: TAGS,
            8: CODE,
            9: NAME,
            10: TYPE_MARK,
            11: UNITS,
            12: VALUES,
        }
        sheet_name = "Перечень материалов"

    class RFP:
        column_dict = {
            0: NUMBERS,
            1: TITLE,
            2: SPECIFICATION_NAME,
            3: TAGS,
            5: CODE,
            6: NAME,
            7: TYPE_MARK,
            9: UNITS,
            8: VALUES,
        }
        sheet_name = "Перечень материалов"
    
    class RFP_AGGREAGATED:
        column_dict = {
            0 : DS_NAME, # Имя ДС, например ДС15, ДС16 или просто 15 , 16
            1: DS_NUMBER, # Порядковый нормер позиции в ДС 1, 2, 3 и т.д
            2: DS_TITLE, # 2265-KSB в данном случае             
            3: DS_SPECIFICATION, #AGCC.287-2265-KSB.MTO-0001
            4: TAGS, # 2265-SH-02-S-UZ-0711
            5: DS_CODE_1C, #002054119
            6: CODE, # BCC0001016
            7: NAME, # Плата источника питания для установки в корпус Elsys-MB
            8: TYPE_MARK, # Elsys-SWPS-2И
            # 8.5 is not read by load_rows_from_worksheet (cell_index is int).
            # Do not insert a physical Excel column here: VALUES must stay at 9.
            # rfp_parts_net.xlsx follows integer keys 0–10 via step1_net_core_columns().
            8.5: VENDOR, # unused by the xlsx loader
            9: VALUES, # 1
            10: UNITS, # шт
            14: NAME_2,
            15: RFP_SUPPLY_STATUS,  # «Статус поставки»; hole between NAME_2 and VALUES_2
            17: VALUES_2,
            18: UNITS_CHECK_STATUS,
            19: UNITS_CONVERSION_TRACE,
        }
        sheet_name = "Перечень материалов"    

    class VO_MTO:
        column_dict = {
            2 : NUMBERS, # Порядковый нормер позиции в ДС 1, 2, 3 и т.д
            3: TAGS, # 2265-SH-02-S-UZ-0711
            4: NAME, # Плата источника питания для установки в корпус Elsys-MB             
            5: VALUES, # 1            
            6: ANNOTATION, # из AGCC.287-2879-SOT.MTO-0001
            7: CODE, # BCC0001016
                      
        }
        sheet_name = -1   

    class BootCO:
        column_dict = {
            8: TAGS,
            0: NUMBERS,
            1: NAME,
            2: TYPE_MARK,
            3: CODE,
            4: VENDOR,
            5: UNITS,
            6: VALUES,
            7: MASS
        }
        sheet_name = "Спецификация"

    class AVEVA:
        column_dict = {
            2: NAME,
            1: TYPE_MARK,
            3: CODE,
            5: UNITS,
        }
        sheet_name = "Спецификация"

    class Equipment:
        column_dict = {
            0: NAME,
            1: TYPE_MARK,
            7: CODE,
            3: VENDOR,
            4: UNITS,
            6: MASS,
        }
        sheet_name = "Спецификация"

    class NanocadDB:
        column_dict = {
            0: NAME,
            1: TYPE_MARK,
            2: CODE,
            3: VENDOR,
            4: UNITS,
            5: MASS,
            6: BD_TABLE_NAME,
            7: BD_ID
        }
        sheet_name = None

    class ZIP:
        column_dict = {
            0: TAGS,
            # 1: NUMBERS,
            2: NAME,
            3: TYPE_MARK,
            4: CODE,
            # 5: VENDOR,
            6: UNITS,
            7: VALUES,
            # 8: MASS,
            # 9: ANNOTATION
            10: MTO_NAME,
            11: ZIP_CODE,
            12: ZIP_VALUE,
        }
        column_dict_adress = {
            0: ZIP_BUILDING,
            1: TAGS,
            2: NAME,
            3: TYPE_MARK,
            4: CODE,
            # 5: VENDOR,
            6: UNITS,
            7: VALUES,
            # 8: MASS,
            # 9: ANNOTATION
            10: MTO_NAME,
            11: ZIP_CODE,
            12: ZIP_VALUE,
        }
        sheet_name = "Спецификация"

    class ZipCompare:
        column_dict = {
            0: TAGS,
            1: NAME,
            2: TYPE_MARK,
            3: CODE,
            4: UNITS,
            13: VALUES,
        }
        sheet_name = "Спецификация"

    class BOE:
        column_dict = {
            0:  NUMBERS,
            1:  TITLE,
            2:  BBB_MARKA,
            3:  BBB_SPEC_NUMBER,
            4:  BBB_REVISION,
            5:  BBB_OBJECT_NAME,
            6:  BBB_MTR_TYPE,
            7:  BBB_MTR_GROUP,
            8:  BBB_SUPPLY_ZONE,
            9:  BBB_WORK_CODE,
            10: TAGS,
            11: NAME,
            12: TYPE_MARK,
            13: CODE,
            14: VENDOR,
            15: UNITS,
            16: VALUES,
            17: MASS,
            18: ANNOTATION,
            19: BBB_UNIT_PRICE,
            20: BBB_TOTAL_PRICE,
            21: BBB_ASSEMBLY,
        }
        sheet_name = "BOE"

    class BOM:
        column_dict = {
            0:  NUMBERS,
            1:  TITLE,
            2:  BBB_MARKA,
            3:  BBB_SPEC_NUMBER,
            4:  BBB_REVISION,
            5:  BBB_OBJECT_NAME,
            6:  BBB_MTR_TYPE,
            7:  BBB_MTR_GROUP,
            8:  BBB_SUPPLY_ZONE,
            9:  BBB_WORK_CODE,
            10: TAGS,
            11: NAME,
            12: TYPE_MARK,
            13: CODE,
            14: VENDOR,
            15: UNITS,
            16: VALUES,
            17: BBB_CONSUMPTION_RATE,
            18: MASS,
            19: ANNOTATION,
            20: BBB_UNIT_PRICE,
            21: BBB_TOTAL_PRICE,
            22: BBB_ASSEMBLY,
        }
        sheet_name = "BOM"

    class BOQ:
        column_dict = {
            0:  NUMBERS,
            1:  TITLE,
            2:  BBB_MARKA,
            3:  BBB_REVISION,
            4:  BBB_OBJECT_NAME,
            5:  BBB_WORK_CODE,
            6:  BBB_WORK_NAME,
            7:  UNITS,
            8:  VALUES,
            9:  BBB_LABOR_UNIT,
            10: BBB_LABOR_TOTAL,
            11: BBB_EQUIPMENT_UNIT,
            12: BBB_EQUIPMENT_TOTAL,
            13: BBB_LABOR_RATE,
            14: BBB_EQUIPMENT_RATE,
            15: BBB_DIRECT_COST_RATE,
            16: BBB_DIRECT_COST_TOTAL,
            17: ANNOTATION,
        }
        sheet_name = "BOQ"

    # ТСД / упаковочные листы (лист Single 1; оркестратор может грузить все Single*).
    # B=CODE, C=имя файла спецификации, D=TAGS, H=NAME;TYPE_MARK, I=VALUES, J=UNITS, M=VENDOR.
    class TsdPacking:
        column_dict = {
            1: CODE,
            2: SPECIFICATION_NAME,
            3: TAGS,
            7: NAME,
            8: VALUES,
            9: UNITS,
            12: VENDOR,
        }
        sheet_name = "Single 1"

    # ТСД формат SO - PL (госфин AGCC.323 и часть PL-АГХК-*): широкий лист.
    # G=PO item→CODE, C=Name of document→SPEC, J=TAG, M=Russian translation→NAME,
    # N=Quantity, O=Units, F=Vendor, BE=Титул→DS_TITLE, BF=Марка→DS_SYSTEM; TYPE_MARK нет.
    class TsdPackingSoPl:
        column_dict = {
            2: SPECIFICATION_NAME,
            5: VENDOR,
            6: CODE,
            9: TAGS,
            12: NAME,
            13: VALUES,
            14: UNITS,
            56: DS_TITLE,
            57: DS_SYSTEM,
        }
        sheet_name = "SO - PL"

    # Сводный Excel после обхода ТСД (порядок колонок вывода).
    class TsdPackingSummary:
        column_dict = {
            0: ANNOTATION,
            1: CODE,
            2: SPECIFICATION_NAME,
            3: DS_TITLE,
            4: DS_SYSTEM,
            5: TAGS,
            6: NAME,
            7: TYPE_MARK,
            8: VALUES,
            9: UNITS,
            10: VENDOR,
            11: TITLE,
        }
        sheet_name = "ТСД свод"
