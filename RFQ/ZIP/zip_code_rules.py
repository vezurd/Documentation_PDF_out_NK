import math

from base.base_classes import *
from base.tables_columns import VALUES, MTO_NAME, ZIP_CODE, ZIP_VALUE
from utils.colors import Color


def zip_code_rule(zip_code, row: RowStd):
    if zip_code == "1":
        proc_no_less_count(15, 1, row)
    elif zip_code == "2":
        proc_no_less_count(5, 2, row)
    elif zip_code == "3":
        proc_no_less_count(10, 1, row)
    elif zip_code == "4":
        proc_no_less_count(10, 1, row)
    elif zip_code == "5":
        proc_no_less_count(20, 2, row)
    elif zip_code == "6":
        one_position(row)
    elif zip_code == "7":
        proc_no_less_count(15, 1, row)
    elif zip_code == "8":
        proc_no_less_count(20, 2, row)
    elif zip_code == "9":
        proc_no_less_count(15, 1, row)
    elif zip_code == "10":
        proc_no_less_count(10, 1, row)
    elif zip_code == "11":
        proc_no_less_count(15, 1, row)
    elif zip_code == "12":
        proc_no_less_count(10, 1, row)
    elif zip_code == "13":
        proc_no_less_count(15, 1, row)
    elif zip_code == "14":
        proc_no_less_count(20, 2, row)
    elif zip_code == "15":
        proc_no_less_count(50, 2, row, in_every_title=True)
    elif zip_code == "16":
        one_position(row)
    elif zip_code == "17":
        proc_no_less_count(10, 1, row, in_area=True)
    elif zip_code == "18":
        one_position_in_area(row)
    elif zip_code == "19":
        proc_no_less_count(10, 1, row)
    elif zip_code == "20":
        one_position(row)
    elif zip_code == "21":
        one_position(row)
    elif zip_code == "22":
        proc_no_less_count(20, 1, row)
    elif zip_code == "23":
        proc_no_less_count(10, 1, row)
    elif zip_code == "24":
        proc_no_less_count(10, 1, row)
    elif zip_code == "25":
        proc_no_less_count(10, 1, row)
    elif zip_code == "26":
        thousand_fo_object(row)
    elif zip_code == "27":
        thousand_fo_object(row)
    elif zip_code == "28":
        proc_no_less_count(30, 2, row)
    elif zip_code == "29":
        proc_no_less_count(15, 0, row)
    elif zip_code == "30":
        proc_no_less_count(5, 1, row)
    elif zip_code == "31":
        proc_no_less_count(5, 1, row)

def proc_no_less_count(proc: int, count: int, row: RowStd, in_every_title=False, in_area=False):
    # Подготовка
    row_value = math.ceil(float(str(row.el[VALUES].value).replace(",", ".")))
    in_every_title_text = ""
    in_area_text = ""
    zip_value_color = Color.green
    if in_every_title:
        # Берем кол-во мто в которых упоминаются материалы
        title_count = len(row.el[MTO_NAME].value)
        count = count * title_count
        in_every_title_text = "в каждом титуле"
    if in_area:
        in_area_text = (f"на площадке\n"
                        f"ПРОВЕРИТЬ ПРИНАДЛЕЖНОСТЬ К ПЛОЩАДКАМ\n")
        zip_value_color = Color.yellow
    # Считаем значение
    zip_value = math.ceil((row_value / 100) * proc)
    if zip_value < count:
        zip_value = count
    row.el[ZIP_VALUE].value = zip_value
    row.el[ZIP_VALUE].comment = (f"{proc}% по каждой поз., но  не менее {count} шт. "
                                 f"{in_every_title_text}{in_area_text}"
                                 f"\n")
    row.el[ZIP_CODE].color = Color.green
    row.el[ZIP_VALUE].color = zip_value_color


def one_position(row: RowStd):
    row.el[ZIP_VALUE].value = 1
    row.el[ZIP_VALUE].comment = f"1 шт. по каждой позиции\n"
    row.el[ZIP_CODE].color = Color.green
    row.el[ZIP_VALUE].color = Color.green


def one_position_in_area(row: RowStd):
    row.el[ZIP_VALUE].value = 1
    row.el[ZIP_VALUE].comment = (f"1 шт. по каждой поз., но не менее 1 шт. на площадке\n"
                                 f"ПРОВЕРИТЬ ПРИНАДЛЕЖНОСТЬ К ПЛОЩАДКАМ\n")
    row.el[ZIP_CODE].color = Color.green
    row.el[ZIP_VALUE].color = Color.yellow
def thousand_fo_object(row: RowStd):
    row.el[ZIP_VALUE].value = 1000
    row.el[ZIP_VALUE].comment = (f"1000 шт. на весь объект\n")
    row.el[ZIP_CODE].color = Color.green
    row.el[ZIP_VALUE].color = Color.yellow
