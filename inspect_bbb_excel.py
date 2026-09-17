# -*- coding: utf-8 -*-
"""Inspect BBB output Excel: row types and fill colors for rows 9, 23 and a few others."""
import openpyxl
import sys

path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\as-build\MTO as-build\2245\test\__результат_MTO_BBB_2026.03.19_13H-15M\check_BOE_vs_base_AGCC.287-2245-KSB.BOE-0001_01-AN01_RU_2026_03_19_13_15.xlsx"

try:
    wb = openpyxl.load_workbook(path, data_only=False)  # data_only=False to see formulas/fills
except FileNotFoundError as e:
    print("FileNotFoundError:", e)
    sys.exit(1)

ws = wb.active
print("Sheet:", ws.title)
print("Headers (row 1):", [ws.cell(1, c).value for c in range(1, 24)])
print()

# BOE columns: A=1 NUMBERS, B=2 TITLE, C=3 BBB_MARKA, D=4 BBB_REVISION, ...
# ROW_TYPE is typically near end - check bbb_output_config: BOE has ROW_TYPE at col 21 (index 21)
# Section 0-indexed in config: 0=NUMBERS, 1=TITLE, 2=BBB_MARKA, 3=BBB_REVISION, ...
# So Excel: col 1 = A = NUMBERS, col 2 = B = TITLE, col 3 = C = BBB_MARKA, col 4 = D = BBB_REVISION
# ROW_TYPE - need to find. From bbb_output_config BOE_COLUMNS_CONFIG, output=True cols:
# 0 NUMBERS, 1 TITLE, 2 BBB_MARKA, 3 BBB_REVISION, 4 BBB_OBJECT_NAME, 5 BBB_MTR_GROUP, 6 BBB_WORK_CODE,
# 7 TAGS, 8 NAME, 9 TYPE_MARK, 10 CODE, 11 VENDOR, 12 UNITS, 13 VALUES, 14 MASS, 15 ANNOTATION,
# 16 ROW_TYPE, 17 SECTION_TYPE
# So ROW_TYPE is col 17 (1-based), SECTION_TYPE col 18

target_rows = [9, 23]
# Also sample a few other rows that should be colored (position_row)
all_data_rows = list(range(2, min(30, ws.max_row + 1)))

for r in all_data_rows:
    row_vals = []
    fills_bcd = []
    for c in range(1, 20):
        val = ws.cell(r, c).value
        cell = ws.cell(r, c)
        fill = cell.fill
        fg = getattr(fill, "fgColor", None)
        rgb = getattr(fg, "rgb", None) if fg else None
        if c in (2, 3, 4):  # B, C, D
            fills_bcd.append((c, rgb))
        row_vals.append(repr(val)[:25] if val is not None else "")
    row_type_val = ws.cell(r, 17).value  # ROW_TYPE column
    print(f"Row {r}: ROW_TYPE={row_type_val!r}  B/C/D fills={fills_bcd}")
    if r in target_rows:
        print(f"  Full row sample: A={row_vals[0]}, B={row_vals[1]}, C={row_vals[2]}, D={row_vals[3]}, CODE(col11)={row_vals[9]}")

wb.close()
print("\nDone.")
