import utils.path
from base.base_classes import TableComments
from base.base_mto import get_std_from_excel_file
import base.t_comm_initial_classes as t_com_init_cls


_DOC_TYPE_TO_CLASS = {
    "BOE": t_com_init_cls.BOE,
    "BOM": t_com_init_cls.BOM,
    "BOQ": t_com_init_cls.BOQ,
}


def find_bbb_files(dir_path: str) -> dict:
    """
    Ищет в dir_path файлы .xlsx и классифицирует их по doc_Type.
    Возвращает dict вида {"BOE": V2Document, "BOM": V2Document, ...}
    """
    all_files = utils.path.get_files_single(dir_path, [".xlsx"])
    if not all_files:
        print(f"По пути <{dir_path}> файлы .xlsx не найдены.")
        return {}

    result = {}
    unrecognized = []
    for doc in all_files:
        if doc.doc_Type in ("BOE", "BOM", "BOQ", "MTO"):
            if doc.doc_Type not in result:
                result[doc.doc_Type] = doc
            else:
                print(f"  Внимание: найден дубль {doc.doc_Type}: {doc.file_full_path}")
        else:
            unrecognized.append(f"    '{doc.file_name}' -> тип: '{doc.doc_Type}'")

    if unrecognized:
        print(f"  Файлы .xlsx найдены, но тип не распознан (ожидается BOE/BOM/BOQ/MTO в имени файла):")
        for line in unrecognized:
            print(line)
    return result


def load_bbb_file(doc, doc_type: str):
    """
    Загружает BBB файл (BOE/BOM/BOQ) в list[RowStd].
    """
    tabel_class = _DOC_TYPE_TO_CLASS.get(doc_type)
    if tabel_class is None:
        print(f"Неизвестный тип BBB: {doc_type}")
        return None

    t_com = TableComments(file_full_path=doc.file_full_path)
    t_com.set_table_class(tabel_class)

    print(f"  Загрузка {doc_type}: {doc.file_full_path}")
    data = get_std_from_excel_file(t_com)
    if data:
        print(f"    -> загружено {len(data)} строк")
    return data


def load_mto_file(doc):
    """
    Загружает MTO файл из найденного V2Document (метаданные из имени файла).
    """
    t_com = TableComments(file_full_path=doc.file_full_path)
    t_com.set_table_class(t_com_init_cls.MTO)

    print(f"  Загрузка MTO: {doc.file_full_path}")
    data = get_std_from_excel_file(t_com)
    if data:
        print(f"    -> загружено {len(data)} строк")
    return data


def load_all_from_dir(dir_path: str) -> dict:
    """
    Ищет и загружает все BBB и MTO файлы из dir_path.
    Возвращает dict: {"BOE": list[RowStd], "BOM": list[RowStd], "BOQ": list[RowStd], "MTO": list[RowStd]}
    """
    found = find_bbb_files(dir_path)
    if not found:
        return {}

    print(f"Найдены файлы: {', '.join(found.keys())}")
    loaded = {}

    for doc_type in ("BOE", "BOM", "BOQ"):
        if doc_type in found:
            data = load_bbb_file(found[doc_type], doc_type)
            if data:
                loaded[doc_type] = data

    if "MTO" in found:
        data = load_mto_file(found["MTO"])
        if data:
            loaded["MTO"] = data

    return loaded
