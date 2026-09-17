# -*- coding: utf-8 -*-
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")

from RFQ.ds_compare.tsd_zinoviev_compare import (
    compare_robot_vs_zinoviev,
    normalize_packing_list_name,
    ul_compare_key,
)

assert normalize_packing_list_name(
    r"согл УЛ ДС13\Packing list PL_2076961.1_2424.02 ред. 28.04.xlsx"
) == normalize_packing_list_name("PL_2076961.1_2424.02")
assert normalize_packing_list_name("№5") == normalize_packing_list_name("№5")
k1 = ul_compare_key("8350", "KSB", "BCC1", "PL_2076961.1_2424.01")
k2 = ul_compare_key(
    "8350",
    "KSB",
    "BCC1",
    r"folder\Packing list PL_2076961.1_2424.01.xlsx",
)
assert k1 == k2
print("OK normalize", k1)

r = compare_robot_vs_zinoviev()
print(r.message)
