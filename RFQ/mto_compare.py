import os
from collections import Counter, defaultdict
from typing import Any, List, Optional, Tuple, Union

from prettytable import PrettyTable

import RFQ.Value_Compare
import utils
import utils.path
import base.t_comm_initial_classes as t_com_init_cls
from base.base_class_std_table import STDTable
from base.base_classes import *
from base.base_mto import get_mto_std_from_file, get_std_from_excel_file
from base.tables_columns import ColNames
from RFQ.func_get_unique_row_list import get_unique_row_list
from RFQ.Value_Compare import file_name_remove_junk, two_mto_compare
from utils.prints_to_console import print_file_list_to_console
from utils.string_parsing import getDocTitle, getMarkaFromFileName, parse_mto_xlsx_revision_for_chain

debug_print_files_list = 1
debug_print_progress = 1


def mto_multi_compare_start(pdf_path):
    print("MTO vs MTO multi сравнение")

    path_out_dir = utils.path.get_path_out_dir(pdf_path)

    # Массив для хранения всех Документов (файлов) текущего проекта
    curr_proj = utils.path.get_files_single(pdf_path, endswith=(".xlsx", ".XLSX"))
    if curr_proj == -1:
        print(f"По пути <{pdf_path}> файлы не найдены.")
        exit(0)

    ############################
    if debug_print_files_list:
        print_file_list_to_console(curr_proj)
    ############################
    all_bases = []
    for document in curr_proj:
        if "MTO" in document.file_name:
            print(f"    Open MTO - {document.file_name}")
            mto_base = get_mto_std_from_file(t_com=TableComments(file_full_path=document.file_full_path,
                                                                 dir_path="-1",
                                                                 tabel_class=t_com_init_cls.MTO,
                                                                 )
                                             , dbg=0)
            all_bases.append(mto_base)
            STDTable.to_excel_mto(mto_base, path_out_dir, "_")
        elif "RFQ" in document.file_name:
            print(f"    Open RFQ - {document.file_name}")
            t_com = TableComments(file_full_path=document.file_full_path,
                                  tabel_class=t_com_init_cls.RFQ,
                                  )
            rfq_base = get_std_from_excel_file(t_com)
            all_bases.append(rfq_base)
        elif "output" in document.file_name:
            print(f"    Open Output - {document.file_name}")
            t_com = TableComments(file_full_path=document.file_full_path,
                                  tabel_class=t_com_init_cls.OUTPUT)
            rfq_base = get_std_from_excel_file(t_com)
            all_bases.append(rfq_base)
        elif "ZIP" in document.file_name:
            print(f"    Open ZIP - {document.file_name}")
            t_com = TableComments(file_full_path=document.file_full_path,
                                  tabel_class=t_com_init_cls.ZipCompare)
            rfq_base = get_std_from_excel_file(t_com)
            STDTable.to_excel_mto(rfq_base, path_out_dir, "01_ZipCompare")
            all_bases.append(rfq_base)

    print("3 Сравнение кол-ва по уникальным кодам")
    diff_base, all_base_unique = RFQ.Value_Compare.start(all_bases, check_position_row=0)
    # excel_base_out(diff_base, path_out_dir)
    # 4
    print("4 Формирование сравнительной таблицы")
    RFQ.Value_Compare.multi_mto_compare(all_base_unique, path_out_dir)

    print("MTO_COMPARE FINISH")


def mto_compare_start(pdf_path):
    print("MTO vs MTO сравнение")

    # Массив для хранения всех Документов (файлов) текущего проекта
    curr_proj = utils.path.get_files_single(pdf_path, endswith=(".xlsx", ".XLSX"))
    if curr_proj == -1:
        print(f"По пути <{pdf_path}> файлы не найдены.")
        exit(0)

    ############################
    if debug_print_files_list:
        print_file_list_to_console(curr_proj)
    ############################
    all_bases = []
    for document in curr_proj:
        if "MTO" in document.file_name:
            print(f"    Open MTO - {document.file_name}")
            mto_base = get_mto_std_from_file(t_com=TableComments(file_full_path=document.file_full_path,
                                                                 dir_path="-1",
                                                                 tabel_class=t_com_init_cls.MTO,
                                                                 )
                                             , dbg=0)
            # exit(0)
            all_bases.append(mto_base)
        elif "RFQ" in document.file_name:
            print(f"    Open RFQ - {document.file_name}")
            t_com = TableComments(file_full_path=document.file_full_path,
                                  tabel_class=t_com_init_cls.RFQ,
                                  )
            rfq_base = get_std_from_excel_file(t_com)
            all_bases.append(rfq_base)
        elif "output" in document.file_name:
            print(f"    Open Output - {document.file_name}")
            t_com = TableComments(file_full_path=document.file_full_path,
                                  tabel_class=t_com_init_cls.OUTPUT)
            rfq_base = get_std_from_excel_file(t_com)
            all_bases.append(rfq_base)

    path_out_dir = utils.path.get_path_out_dir(pdf_path)

    #  3 Сравнение колва по уникальным кодам
    diff_base, all_base_unique = RFQ.Value_Compare.start(all_bases)
    # excel_base_out(diff_base, path_out_dir)
    # 4
    RFQ.Value_Compare.two_mto_compare(all_base_unique, path_out_dir)

    print("MTO_COMPARE FINISH")


def _mto_chain_group_key(file_name: str) -> str:
    """Build title-mark key; raises ValueError if title or marka cannot be parsed."""
    title = getDocTitle(file_name)
    marka = getMarkaFromFileName(file_name)
    if title == "<No Found>" or marka == "<No Found>":
        raise ValueError(
            "Не удалось определить титул или марку по имени файла (как для PDF-штампа).\n"
            f"Файл: {file_name!r}"
        )
    return f"{title}-{marka}"


def _sanitize_path_segment(segment: str, max_len: int = 120) -> str:
    """Make a single path segment safe on Windows."""
    bad = '<>:"/\\|?*'
    s = segment.strip()
    for c in bad:
        s = s.replace(c, "_")
    s = s.strip(". ") or "_"
    if len(s) > max_len:
        s = s[:max_len].rstrip(". ")
    return s or "_"


def _revision_sort_key(rev: str) -> Tuple[int, Union[int, str], str]:
    """Sort numeric revisions as integers; preserve display string as tie-breaker."""
    if rev.isdigit():
        return (0, int(rev), rev)
    return (1, rev, rev)


def _chain_item_sort_key(rev: str, appendix: str) -> Tuple[Tuple[int, Union[int, str], str], str]:
    """Order files in a group: by revision, then by appendix (e.g. AN01 before AN02)."""
    return (_revision_sort_key(rev), appendix)


def _chain_rev_ap_display(rev: str, appendix: str) -> str:
    if appendix:
        return f"{rev}-{appendix}"
    return rev


def _print_mto_chain_plan(
    pair_rows: List[Tuple[str, str, str, str, str]],
    singles: List[Tuple[str, str]],
) -> None:
    """Console: PrettyTable of planned A–B pairs; then groups with a lone file."""
    if pair_rows:
        t = PrettyTable()
        t.field_names = ["Группа (титул–марка)", "Файл A", "Файл B", "Рев. A", "Рев. B"]
        t.align["Файл A"] = "l"
        t.align["Файл B"] = "l"
        for gkey, fa, fb, la, lb in pair_rows:
            t.add_row([gkey, fa, fb, la, lb])
        print("\nMTO цепочка — пары к сравнению (соседи после сортировки):")
        print(t)
    else:
        print("\nMTO цепочка — пар для сравнения нет (или все группы по одному файлу).")

    if singles:
        u = PrettyTable()
        u.field_names = ["Группа (титул–марка)", "Файл без пары в группе"]
        u.align["Файл без пары в группе"] = "l"
        for gkey, fn in singles:
            u.add_row([gkey, fn])
        print(
            "\nФайлы в группах без цепочки (в группе один MTO — не с чем сравнить внутри неё):"
        )
        print(u)


def mto_chain_compare_start(pdf_path: str) -> Optional[str]:
    """Compare adjacent MTO revisions per title-mark group (3+ files supported).

    Файлы группируются по паре (титул, марка) из имени; внутри группы сортировка по
    ревизии и суффиксу вроде ``-AN01``. Сравниваются соседи в этом порядке (не
    склейка «разные титулы в одну цепочку»).

    Args:
        pdf_path: Folder containing MTO ``.xlsx`` files.

    Returns:
        None on success; a Russian error message string if the GUI should show ``showerror``.
    """
    print("MTO vs MTO — цепочка ревизий")

    curr_proj = utils.path.get_files_single(pdf_path, endswith=(".xlsx", ".XLSX"))
    if not curr_proj:
        return "В выбранной папке не найдены файлы .xlsx."

    mto_docs = [d for d in curr_proj if "MTO" in d.file_name]
    if not mto_docs:
        return "В папке нет файлов MTO (.xlsx с «MTO» в имени)."

    parsed: List[Tuple[Any, str, str, str]] = []
    try:
        for document in mto_docs:
            fn = document.file_name
            gkey = _mto_chain_group_key(fn)
            rev, appendix = parse_mto_xlsx_revision_for_chain(fn)
            parsed.append((document, gkey, rev, appendix))
    except ValueError as e:
        return str(e)

    by_group: dict[str, List[Tuple[Any, str, str, str]]] = defaultdict(list)
    for item in parsed:
        by_group[item[1]].append(item)

    pair_rows: List[Tuple[str, str, str, str, str]] = []
    singles: List[Tuple[str, str]] = []
    pair_specs: List[Tuple[Any, Any, str, str, str, str, str, str]] = []

    base_dir = os.path.normpath(pdf_path)

    for gkey in sorted(by_group.keys()):
        items = by_group[gkey]
        if len(items) < 2:
            for t in items:
                singles.append((gkey, t[0].file_name))
            continue

        items.sort(key=lambda t: _chain_item_sort_key(t[2], t[3]))
        rev_ap_keys = [(t[2], t[3]) for t in items]
        dup = [k for k, c in Counter(rev_ap_keys).items() if c > 1]
        if dup:
            bad_rev, bad_apx = dup[0]
            names = [
                t[0].file_name
                for t in items
                if t[2] == bad_rev and t[3] == bad_apx
            ]
            dup_label = _chain_rev_ap_display(bad_rev, bad_apx)
            return (
                f"В группе {gkey!r} дублируется позиция в цепочке {dup_label!r} "
                "(одинаковые ревизия и суффикс листа).\n" + "\n".join(names)
            )

        for i in range(len(items) - 1):
            doc_a, _, rev_a, apx_a = items[i]
            doc_b, _, rev_b, apx_b = items[i + 1]
            la = _chain_rev_ap_display(rev_a, apx_a)
            lb = _chain_rev_ap_display(rev_b, apx_b)
            pair_rows.append((gkey, doc_a.file_name, doc_b.file_name, la, lb))
            pair_specs.append((doc_a, doc_b, gkey, rev_a, apx_a, rev_b, apx_b))

    _print_mto_chain_plan(pair_rows, singles)

    any_pair_written = False
    for doc_a, doc_b, gkey, rev_a, apx_a, rev_b, apx_b in pair_specs:
        mto1_name = file_name_remove_junk(doc_a.file_name)
        mto2_name = file_name_remove_junk(doc_b.file_name)
        pair_label = f"{mto1_name}_vs_{mto2_name}"
        sub = _sanitize_path_segment(pair_label, max_len=150)
        pair_out = os.path.join(base_dir, sub)
        os.makedirs(pair_out, exist_ok=True)
        pair_out_sep = pair_out + os.sep

        la = _chain_rev_ap_display(rev_a, apx_a)
        lb = _chain_rev_ap_display(rev_b, apx_b)
        print(f"    Chain pair {gkey}: {la} -> {lb}")
        print(f"        Open MTO - {doc_a.file_name}")
        base_a = get_mto_std_from_file(
            t_com=TableComments(
                file_full_path=doc_a.file_full_path,
                dir_path="-1",
                tabel_class=t_com_init_cls.MTO,
            ),
            dbg=0,
        )
        print(f"        Open MTO - {doc_b.file_name}")
        base_b = get_mto_std_from_file(
            t_com=TableComments(
                file_full_path=doc_b.file_full_path,
                dir_path="-1",
                tabel_class=t_com_init_cls.MTO,
            ),
            dbg=0,
        )
        u_a = get_unique_row_list(base_a, check_position_flag=0)
        u_b = get_unique_row_list(base_b, check_position_flag=0)
        two_mto_compare([u_a, u_b], pair_out_sep, auto_open_dir=False)
        any_pair_written = True

    if any_pair_written:
        utils.path.open_dir(base_dir)
    else:
        print(
            "    Ни одной пары для сравнения (в каждой группе титул–марка только по одному файлу)."
        )

    print("MTO_CHAIN_COMPARE FINISH")
    return None


if __name__ in {"__main__"}:
    mto_path = r"C:\YandexDisk\темп\MTO MTO"
    # mto_path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\2225\KSB\Для передачи\05_рев.01_AGCC.287‐2225‐KSB\_ИД для корректировки\01 output (MTO CJ)\проверка"
    mto_compare_start(mto_path)
