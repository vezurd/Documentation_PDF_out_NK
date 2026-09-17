TEMPLATE_CJ = r"templates\AGCC.287-XXXX-YYYY.CJ-00ZZ_0_RU.docx"

G_BASE_CABLE_LAYING_FLAG = "G_BASE_CABLE_LAYING_FLAG"
G_BASE_CABLE_LAYING_FLAG_VALUE = "КНС_ДЛЯ_КЖ"

G_BASE_CABLE_SUMMARY_FLAG = "G_BASE_CABLE_SUMMARY_FLAG"
G_BASE_CABLE_SUMMARY_FLAG_VALUE = "КАБЕЛЬ_КЖ"
G_BASE_CABLE_SUMMARY_VARIANTS = "G_BASE_CABLE_SUMMARY_VARIANTS"

G_BASE_CABLE_LAYING_TYPE = "G_BASE_CABLE_LAYING_TYPE"
G_BASE_CABLE_LAYING_VARIANTS = "G_BASE_CABLE_LAYING_VARIANTS"

#   ТПК.xlsx
FILE_NAME_TPK = ["ТПК", "TPK"]
FILE_NAME_CJ = ["CJ", "CJ"]
FILE_DICT = {
    FILE_NAME_TPK[1]: [-1],  # [Имя_листа]
    FILE_NAME_CJ[1]: [-1],  # [Имя_листа]
}

TPK_R0_CABLE = "Кабель, жгут"
# TPK_R1_FROM                     = "Откуда"
TPK_R2_FROM_FLOOR = "Откуда: Этаж"
TPK_R3_FROM_ROOM = "Откуда: Помещение"
# TPK_R4_WHERE                    = "Куда"
TPK_R5_WHERE_FLOOR = "Куда: Этаж"
TPK_R6_WHERE_ROOM = "Куда: Помещение"
# TPK_R7_CABLE_MARK               = "Кабель Марка"
TPK_R8_TOTAL_LENGTH = "Кабель, провод: Проект-ая длина кабеля, м"
TPK_R9_CABLE_LAYING_TYPE = "Способ прокладки: Обозначение"
TPK_R10_CABLE_LAYING_LENGTH = "Способ прокладки: Проект-ая длина, м"

#   CJ
CJ_CABLE_CODE = "CJ_CABLE_CODE"  # A, I, C etc
CJ_CABLE_NUMBER = "CJ_CABLE_NUMBER"  # 0001 0002
CJ_ADDITIONAL_CODE = "CJ_ADDITIONAL_CODE"  # Ex IS
CJ_CABLE_LABEL = "CJ_CABLE_LABEL"  # A-0001
CJ_CABLE_FULL_TAG = "CJ_CABLE_FULL_TAG"
CJ_FROM = "CJ_FROM"
CJ_FROM_TERMINAL = "CJ_FROM_TERMINAL"
CJ_WHERE = "CJ_WHERE"
CJ_WHERE_TERMINAL = "CJ_WHERE_TERMINAL"
CJ_CABLE_MARK = "CJ_CABLE_MARK"
CJ_TOTAL_LENGTH = "CJ_TOTAL_LENGTH"
CJ_ANNOTAION = "CJ_ANNOTAION"

CJ_SYSTEM = "CJ_SYSTEM"
CJ_CABLE_CODE_FLAG = "CJ_CABLE_CODE_FLAG"  # откуда взят номер кабеля (сгенерирован или из старого КЖ)

TEXT_CABLE_CODE_FLAG_1 = "Номер заменен из предыдущего КЖ"
TEXT_CABLE_CODE_FLAG_2 = "Номер заменен Автоматически"

########################################
CABLE_LAYING_IN_CABIN = "В шкафу"
CABLE_LAYING_IN_TRAY = "В лотке"
CABLE_LAYING_IN_CABLE_CHANEL = "В кабель-канале"
CABLE_LAYING_IN_OPEN = "Открыто по конструкциям"
CABLE_LAYING_IN_FLUTED = "В гофротрубе"
CABLE_LAYING_IN_METAL_HOUSE = "В металлорукаве"
CABLE_LAYING_IN_TUBE = "В трубе"
CABLE_LAYING_IN_TRENCH = "В траншее"

CABLE_LAYING_TYPES_LIST = [
    CABLE_LAYING_IN_CABIN,
    CABLE_LAYING_IN_TRAY,
    CABLE_LAYING_IN_CABLE_CHANEL,
    CABLE_LAYING_IN_OPEN,
    CABLE_LAYING_IN_FLUTED,
    CABLE_LAYING_IN_METAL_HOUSE,
    CABLE_LAYING_IN_TUBE,
    CABLE_LAYING_IN_TRENCH,
]

# SummaryTableRow
SUM_MARK = "SUM_MARK"
SUM_CORES = "SUM_CORES"
SUM_LENGTH = "SUM_LENGTH"

# Системы
SYS_SOS = "SOS"
SYS_SKUD = "SKUD"
SYS_SOT = "SOT"
SYS_SPP = "SPP"
SYS_POS = "POS"
SYS_SOO = "SOO"
SYS_KBI = "KBI"
SYS_KSB = "KSB"
SYS_LIST = (SYS_SOS,
            SYS_SKUD,
            SYS_SOT,
            SYS_SPP,
            SYS_POS,
            SYS_SOO,
            SYS_KBI,
            SYS_KSB)
