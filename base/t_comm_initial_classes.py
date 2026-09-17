from base.tables_columns import ColNames

class DSZin:
    column_dict = ColNames.DsZin.column_dict
    tabel_type = "MTO"
    #
    """
    0 - Индекс комментариев для ГуглБазы
    1 - Индекс комментариев для МТО    
    """
    tabel_type_index = 1
    #
    sheet_name = ColNames.DsZin.sheet_name

class DsSpecification:
    column_dict = ColNames.DsSpecification.column_dict
    tabel_type = "DsSpecification"
    tabel_type_index = 1
    #
    sheet_name = ColNames.DsSpecification.sheet_name

class ListDsSpecification:
    column_dict = ColNames.ListDsSpecification.column_dict
    tabel_type = "DsSpecification"
    tabel_type_index = 1
    #
    sheet_name = ColNames.DsSpecification.sheet_name

class RFQTPK:
    column_dict = ColNames.RFQTSpecification.column_dict
    tabel_type = "RFQTSpecification"
    tabel_type_index = 1  
    #
    sheet_name = ColNames.RFQTSpecification.sheet_name

class ListDsSpecificationVsMto:
    column_dict = ColNames.DsVsMto.column_dict
    tabel_type = "DsVsMto"
    tabel_type_index = 1
    #
    sheet_name = ColNames.DsSpecification.sheet_name

class MTO:
    column_dict = ColNames.MTO.column_dict
    tabel_type = "MTO"
    #
    """
    0 - Индекс комментариев для ГуглБазы
    1 - Индекс комментариев для МТО    
    """
    tabel_type_index = 1
    #
    sheet_name = ColNames.MTO.sheet_name


class OUTPUT:
    column_dict = ColNames.Output.column_dict
    tabel_type = "OUTPUT"
    tabel_type_index = 1
    sheet_name = ColNames.Output.sheet_name


class ZIP:
    column_dict = ColNames.ZIP.column_dict
    tabel_type = "ZIP"
    tabel_type_index = 1
    sheet_name = ColNames.ZIP.sheet_name


class ZipCompare:
    column_dict = ColNames.ZipCompare.column_dict
    tabel_type = "ZipCompare"
    tabel_type_index = 1
    sheet_name = ColNames.ZipCompare.sheet_name


class BOOT_CO:
    column_dict = ColNames.BootCO.column_dict
    tabel_type = "BOOT_CO"
    tabel_type_index = 1
    sheet_name = ColNames.BootCO.sheet_name


class AVEVA:
    column_dict = ColNames.AVEVA.column_dict
    tabel_type = "AVEVA"
    tabel_type_index = 0
    sheet_name = ColNames.AVEVA.sheet_name


class RFQ:
    column_dict = ColNames.BootCO.column_dict
    tabel_type = "RFQ"
    tabel_type_index = 1
    sheet_name = ColNames.BootCO.sheet_name


class RFP:
    column_dict = ColNames.RFP.column_dict
    tabel_type = "RFP"
    tabel_type_index = 1
    sheet_name = ColNames.RFP.sheet_name

class RFP_AGGREGATED:
    column_dict = ColNames.RFP_AGGREAGATED.column_dict
    tabel_type = "RFP_AGGREAGATED"
    tabel_type_index = 1
    sheet_name = ColNames.RFP_AGGREAGATED.sheet_name

class VO_MTO:
    column_dict = ColNames.VO_MTO.column_dict
    tabel_type = "VO_MTO"
    tabel_type_index = 1
    sheet_name = ColNames.VO_MTO.sheet_name

class GoogleBase:
    column_dict = ColNames.MTO.column_dict
    # column_dict = ColNames.GoogleBase.column_dict
    tabel_type = "GoogleBase"
    tabel_type_index = 0
    sheet_name = ColNames.GoogleBase.sheet_name


class EQUIPMENT:
    column_dict = ColNames.Equipment.column_dict
    tabel_type = "EQUIPMENT"
    tabel_type_index = 1
    sheet_name = ColNames.Equipment.sheet_name


class BOE:
    column_dict = ColNames.BOE.column_dict
    tabel_type = "BOE"
    tabel_type_index = 1
    sheet_name = ColNames.BOE.sheet_name


class BOM:
    column_dict = ColNames.BOM.column_dict
    tabel_type = "BOM"
    tabel_type_index = 1
    sheet_name = ColNames.BOM.sheet_name


class BOQ:
    column_dict = ColNames.BOQ.column_dict
    tabel_type = "BOQ"
    tabel_type_index = 1
    sheet_name = ColNames.BOQ.sheet_name


class TsdPacking:
    column_dict = ColNames.TsdPacking.column_dict
    tabel_type = "TsdPacking"
    tabel_type_index = 1
    sheet_name = ColNames.TsdPacking.sheet_name


class TsdPackingSoPl:
    column_dict = ColNames.TsdPackingSoPl.column_dict
    tabel_type = "TsdPackingSoPl"
    tabel_type_index = 1
    sheet_name = ColNames.TsdPackingSoPl.sheet_name
