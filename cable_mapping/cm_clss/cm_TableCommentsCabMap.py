import os

import utils
from cable_mapping.constants import SYS_LIST, FILE_NAME_TPK, FILE_NAME_CJ, FILE_DICT


class TableCommentsCabMap:

    def __init__(self,
                 dir_path="",
                 file_full_path="",
                 file_full_name="",
                 tabel_type="",
                 system="NO_SYS",
                 sheet_name=""):
        self.sheet_name = None
        self.tabel_type = None
        self.tabel_type_index = None  # Индекс для выбора комментариев при проверке позиций
        self.column_dict = None
        self.out_dir = None
        self.system = system

        self.dir_path = dir_path
        self.file_full_path = file_full_path
        self.file_full_name = file_full_name
        if file_full_name == "" and len(str(self.file_full_path)) > 5:
            self.file_full_name = utils.path.get_file_name_from_full_file_path(self.file_full_path)
        if self.dir_path != "":
            self.out_dir = utils.path.get_path_out_dir(dir_path)

        for SYS in SYS_LIST:
            if SYS in self.file_full_name:
                self.system = SYS
                continue


class DocumentCabMap:
    def __init__(self, file_path, dir_path=""):
        self.file_full_path = file_path  #
        self.file_name = os.path.basename(file_path)  # AGCC.287-8525-SKUD.BOM-0001_A_RU.pdf
        self.doc_type = DocumentCabMap.getDocType(self.file_name)
        self.sheet_name = DocumentCabMap.getSheetName(self.doc_type)
        self.dir_path = dir_path
        self.t_com = TableCommentsCabMap(dir_path=self.dir_path,
                                         file_full_path=self.file_full_path,
                                         file_full_name=self.file_name)

    @staticmethod
    def is_tpk_normalized_output(file_name: str) -> bool:
        """True for xlsx that is TPK normalization output, not raw nanoCAD export."""
        name_lower = file_name.casefold()
        if name_lower.startswith("tpk_normalized_"):
            return True
        if name_lower.startswith("normalized_"):
            if FILE_NAME_TPK[0].casefold() in name_lower:
                return True
            if "tpk" in name_lower:
                return True
        return False

    @staticmethod
    def getDocType(file_name):
        if DocumentCabMap.is_tpk_normalized_output(file_name):
            return None
        if FILE_NAME_TPK[0] in file_name:
            return FILE_NAME_TPK[1]
        if FILE_NAME_CJ[0] in file_name:
            return FILE_NAME_CJ[1]
        else:
            return None

    @staticmethod
    def getSheetName(doc_type):
        if doc_type in FILE_DICT:
            return FILE_DICT[doc_type][0]
        else:
            return None
