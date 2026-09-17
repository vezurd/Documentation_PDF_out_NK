import json
import os

import pygsheets

from base.base_google import google_get_data, load_json, serv_file
from utils.string_parsing import print_att_list_table

json_file_name_code_base = 'base_check/mapping_base_data.json'
json_file_no_out_cables = "base_check/no_out_cables.json"
json_open_laying = "base_check/json_open_laying.json"
_CACHE_META_FILE_MAP = 'base_check/mapping_base_meta.json'

_MAP_SPREADSHEET_ID = "1VyHiGjcbc9yow3D5ratq3tSq-A2-S-mmSrwRfA4oNo8"
_MAP_SPREADSHEET_URL = (
    f"https://docs.google.com/spreadsheets/d/{_MAP_SPREADSHEET_ID}/edit?usp=sharing"
)


def _auto_sync_map():
    """Проверяет modifiedTime mapping-таблицы и обновляет все три листа при изменении."""
    all_cached = (os.path.exists(json_file_name_code_base)
                  and os.path.exists(json_file_no_out_cables)
                  and os.path.exists(json_open_laying))
    meta = _load_meta()

    if all_cached and meta:
        try:
            gc = pygsheets.authorize(service_file=serv_file)
            current_mt = gc.drive.get_update_time(_MAP_SPREADSHEET_ID)
            if current_mt == meta.get('modified_time'):
                print("Mapping-таблица не изменилась — используем кэш")
                return
            print("Mapping-таблица изменилась — загружаем обновление...")
            _fetch_all_sheets(gc, current_mt)
        except Exception as e:
            print(f"Не удалось проверить Mapping-таблицу: {e}")
            print("Используем локальный кэш")
    else:
        try:
            gc = pygsheets.authorize(service_file=serv_file)
            current_mt = gc.drive.get_update_time(_MAP_SPREADSHEET_ID)
            _fetch_all_sheets(gc, current_mt)
        except Exception as e:
            if all_cached:
                print(f"Ошибка загрузки Mapping-таблицы: {e}, используем устаревший кэш")
            else:
                raise RuntimeError(
                    f"Нет локального кэша и не удалось загрузить Mapping-таблицу: {e}"
                ) from e


def _fetch_all_sheets(gc, modified_time):
    sh = gc.open_by_url(_MAP_SPREADSHEET_URL)
    for ws_name, cache_file in [("mapping", json_file_name_code_base),
                                ("no_out_cables", json_file_no_out_cables),
                                ("open_laying", json_open_laying)]:
        wks = sh.worksheet_by_title(ws_name)
        data = wks.get_values(None, None, returnas='matrix')
        with open(cache_file, 'w') as f:
            json.dump(data, f)
    _save_meta(modified_time)
    print(f"Mapping-таблица загружена (modifiedTime: {modified_time})")


def _load_meta():
    if not os.path.exists(_CACHE_META_FILE_MAP):
        return None
    try:
        with open(_CACHE_META_FILE_MAP) as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return None


def _save_meta(modified_time):
    with open(_CACHE_META_FILE_MAP, 'w') as f:
        json.dump({'modified_time': modified_time}, f)


def load_map_google_base(dbg=0, max_len=25):
    _auto_sync_map()
    code_base_data_raw = load_json(json_file_name_code_base)
    if dbg:
        print_att_list_table(code_base_data_raw, max_len=25, title="load_map_google_base")
    return code_base_data_raw


def load_no_out_cables_google_base(dbg=0, max_len=25):
    code_base_data_raw = load_json(json_file_no_out_cables)
    if dbg:
        print_att_list_table(code_base_data_raw, max_len=25, title="load_no_out_cables_google_base")
    return code_base_data_raw


def open_laying_google_base(dbg=0, max_len=25):
    base_data_raw = load_json(json_open_laying)
    if dbg:
        print_att_list_table(base_data_raw, max_len=25, title="open_laying_google_base")
    return base_data_raw
