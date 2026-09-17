from __future__ import annotations

import os
from pathlib import Path
from typing import List, TypedDict

import base.base_mto
import base.t_comm_initial_classes as t_com_init_cls
from base.base_cheks import check_position_row
from base.base_class_std_table import STDTable
from base.base_classes import RowStd, TableComments
from base.tables_columns import ColNames
from RFQ.ds_compare.ds_units_normalize import normalize_mto_positions_units
from pdf_parsing_v2_engine.document import V2Document
from utils.cache_utils import CacheManager

# Per-file MTO cache (pickle + mtime), same mechanism as step2_load_mto / cache_manager.
# Separate directory so DS compare does not share files with the project root "cache" folder
# unless the same absolute paths are read from both code paths.
_DS_COMPARE_MTO_CACHE = Path(__file__).resolve().parent / "cache"
_mto_file_cache = CacheManager(str(_DS_COMPARE_MTO_CACHE))
# Bump if get_mto_std / Excel mapping semantics change for this pipeline.
MTO_PICKLE_EXTRA_KEY = "ds_mto_get_mto_std_v2"
_MTO_PICKLE_EXTRA_KEY = MTO_PICKLE_EXTRA_KEY


class MtoLoadAudit(TypedDict):
    """How one MTO file was loaded (for IN_CABINET debug trace)."""

    path: str
    from_cache: bool
    cache_file: str
    position_row_count: int


def _load_mto_file_raw(
    mto_path: str,
    *,
    use_mto_cache: bool,
    mto_force_update: bool,
) -> tuple[List[RowStd] | None, MtoLoadAudit]:
    """Load full MTO table (before ``check_position_row`` filter)."""
    audit: MtoLoadAudit = {
        "path": mto_path,
        "from_cache": False,
        "cache_file": "",
        "position_row_count": 0,
    }
    mto_base: List[RowStd] | None = None
    cache_path = _mto_file_cache._get_cache_file_path(mto_path, extra_key=_MTO_PICKLE_EXTRA_KEY)

    if (
        use_mto_cache
        and not mto_force_update
        and _mto_file_cache.is_cache_valid(mto_path, extra_key=_MTO_PICKLE_EXTRA_KEY)
    ):
        mto_base = _mto_file_cache.load_from_cache(
            mto_path, verbose=False, extra_key=_MTO_PICKLE_EXTRA_KEY
        )
        if mto_base is not None:
            audit["from_cache"] = True
            audit["cache_file"] = str(cache_path)

    if mto_base is None:
        mto_obj = V2Document.from_file_path(mto_path)
        print(f"    Open MTO - {mto_obj.file_name}")
        mto_base = base.base_mto.get_mto_std_from_file(
            t_com=TableComments(
                file_full_path=mto_obj.file_full_path,
                dir_path="-1",
                tabel_class=t_com_init_cls.MTO,
            ),
            dbg=0,
        )
        if use_mto_cache and mto_base is not None:
            _mto_file_cache.save_to_cache(
                mto_path, mto_base, verbose=False, extra_key=_MTO_PICKLE_EXTRA_KEY
            )

    if mto_base is None:
        return None, audit

    positions = [row for row in mto_base if check_position_row(row)]
    normalize_mto_positions_units(positions)
    audit["position_row_count"] = len(positions)
    return positions, audit


# Что делать со строчками которые не проходят по проверке кода закупочного
# 0 - добавляем все в сравнение, в т.ч. пустые строки
# 1 - добавляем только прошедшие проверку на коды, только валидные строки
# 2 - убираем пустые строки, не выалидные но не пустые - не убираем
check_position_flag = 1


def load_mto_by_dict(
    spec_dict: dict,
    info_flag: bool = False,
    print_mto_to_console: bool = False,
    use_mto_cache: bool = True,
    mto_force_update: bool = False,
) -> tuple[dict, dict[str, MtoLoadAudit]]:
    mto_dict = {}
    load_audit: dict[str, MtoLoadAudit] = {}
    for mto_name, mto_path in spec_dict.items():
        if mto_path == "Файл МТО не найден":
            print(f"{mto_name}: {mto_path}")
            continue

        if (
            use_mto_cache
            and not mto_force_update
            and _mto_file_cache.is_cache_valid(
                mto_path, extra_key=_MTO_PICKLE_EXTRA_KEY
            )
        ):
            print(f"    MTO из кэша — {os.path.basename(mto_path)}")

        mto_base_positions, audit = _load_mto_file_raw(
            mto_path,
            use_mto_cache=use_mto_cache,
            mto_force_update=mto_force_update,
        )
        if info_flag and not audit["from_cache"]:
            mto_obj = V2Document.from_file_path(mto_path)
            mto_obj.print_debug()

        if mto_base_positions is None:
            continue

        if print_mto_to_console:
            STDTable.print_to_console(mto_base_positions, ColNames.MTO.column_dict)
        mto_dict[mto_name] = mto_base_positions
        load_audit[mto_name] = audit

    return mto_dict, load_audit


def refresh_stale_mto_caches(spec_dict: dict[str, object]) -> int:
    """Refresh only relevant MTO pickles whose source mtime has changed.

    Args:
        spec_dict: Mapping from specification key to MTO workbook path.

    Returns:
        Number of MTO workbooks re-read from Excel.
    """
    refreshed = 0
    for raw_path in spec_dict.values():
        mto_path = str(raw_path or "")
        if not mto_path or mto_path == "Файл МТО не найден" or not os.path.isfile(mto_path):
            continue
        if _mto_file_cache.is_cache_valid(
            mto_path, extra_key=_MTO_PICKLE_EXTRA_KEY
        ):
            continue
        _load_mto_file_raw(
            mto_path,
            use_mto_cache=True,
            mto_force_update=False,
        )
        refreshed += 1
    if refreshed:
        print(f"Preflight MTO: обновлено файлов кэша={refreshed}.")
    else:
        print("Preflight MTO: все релевантные файлы актуальны.")
    return refreshed


def reload_mto_positions_fresh(mto_path: str) -> List[RowStd] | None:
    """Reload one MTO from Excel bypassing pickle (same pipeline as ``google_sheets.start``)."""
    rows, _audit = _load_mto_file_raw(
        mto_path, use_mto_cache=False, mto_force_update=True
    )
    return rows
