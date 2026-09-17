
from base.tables_columns import UNITS, CODE, VENDOR, MASS, TYPE_MARK, NAME
import base.t_comm_initial_classes as t_com_init_cls

class GuiConst:
    MTO_DIR = "MTO_DIR"
    MTO_FILE = "MTO_FILE"
    OUTPUT_FILE = "OUTPUT_FILE"
    OUTPUT_DIR = "OUTPUT_DIR"
    BOOT_DIR = "BOOT_DIR"
    BOOT_FILE = "BOOT_FILE"
    AVEVA_FILE = "AVEVA_FILE"
    EQUIPMENT_FILE = "EQUIPMENT_FILE"
    NANOCAD_DB = "NANOCAD_DB"
    RFQ_DIR = "RFQ_DIR"
    RFP = "RFP"
    dict = {
        MTO_FILE: ["Запуск сравнения файла MTO с ГуглБазой",
                   "Открыть файл MTO",
                   "check_mto_vs_base",
                   [NAME, TYPE_MARK, UNITS, VENDOR, CODE, ],
                   t_com_init_cls.MTO],
        MTO_DIR: ["Запуск поиска файла MTO и сравнение MTO с ГуглБазой",
                  "Открыть папку DWG с MTO",
                  "check_mto_vs_base",
                  [NAME, TYPE_MARK, UNITS, VENDOR, CODE, ],
                  t_com_init_cls.MTO],
        AVEVA_FILE: ["Запуск сравнения базы AVEVA с ГуглБазой",
                     "Открыть файл базы AVEVA",
                     "aveva_vs_base",
                     [NAME, TYPE_MARK, UNITS, CODE],
                     t_com_init_cls.AVEVA],
        OUTPUT_FILE: ["Запуск сравнения OUTPUT с ГуглБазой",
                      "Открыть файл output",
                      "check_output_CO",
                      [NAME, TYPE_MARK, UNITS, VENDOR, CODE, ],
                      t_com_init_cls.MTO],
        OUTPUT_DIR: ["Запуск поиска файла OUTPUT и сравнение MTO с ГуглБазой",
                     "Открыть папку с OUTPUT",
                     "check_output_CO",
                     [NAME, TYPE_MARK, UNITS, VENDOR, CODE, ],
                     t_com_init_cls.MTO],
        BOOT_DIR: ["Запуск поиска файла CO.xlsx и сравнение с ГуглБазой",
                   "Открыть папку BOOT",
                   "check_Boot_CO",
                   [NAME, TYPE_MARK, UNITS, VENDOR, CODE, ],
                   t_com_init_cls.MTO],
        BOOT_FILE: ["Запуск сравнения CO.xlsx с ГуглБазой",
                    "Открыть файл CO.xlsx",
                    "check_Boot_CO",
                    [NAME, TYPE_MARK, UNITS, VENDOR, CODE, ],
                    t_com_init_cls.MTO],
        EQUIPMENT_FILE: ["Запуск сравнения базы Оборудования и материалов с ГуглБазой",
                         "Открыть файл Оборудование",
                         "equipment_vs_base",
                         [NAME, TYPE_MARK, UNITS, CODE, VENDOR, MASS],
                         t_com_init_cls.MTO],
        NANOCAD_DB: ["Запуск сравнения базы NanoCAD и материалов с ГуглБазой",
                         "Открыть файл *.db (NanoCAD)",
                         "NanoCAD_DB_vs_base",
                         [NAME, TYPE_MARK, UNITS, CODE, VENDOR, MASS],
                     t_com_init_cls.MTO],
        RFQ_DIR: ["Запуск сравнения МТО и RFQ пакета",
                     "Открыть папку с RFQ",
                     "MTO_vs_RFQ",
                     [NAME, TYPE_MARK, UNITS, CODE, VENDOR, MASS],
                     t_com_init_cls.RFQ],
        RFP: ["Запуск сравнения МТО и RFQ пакета",
                  "Открыть папку с RFQ",
                  "MTO_vs_RFQ",
                  [NAME, TYPE_MARK, UNITS, CODE, VENDOR, MASS],
                  t_com_init_cls.RFP],
    }
