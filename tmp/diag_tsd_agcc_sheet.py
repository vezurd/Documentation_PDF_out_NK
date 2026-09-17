"""Diagnose AGCC sheet with 0 positions; write UTF-8 dump."""
from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import openpyxl

from base.base_classes import RowType
from base.tables_columns import (
    CODE,
    NAME,
    SPECIFICATION_NAME,
    TAGS,
    TITLE,
    UNITS,
    VALUES,
    VENDOR,
)
from RFQ.ds_compare.tsd_packing_load import (
    DEFAULT_TSD_PACKING_ROOT,
    collect_tsd_files,
    load_tsd_file_rows,
)

OUT = ROOT / "tmp" / "tsd_agcc_sheet_diag.txt"
TARGET_SUFFIX = "PL-АГХК-ПЭППиЛАО-2076961.15.2.xlsx"
SHEET = "AGCC-2076961.15-2"


def main() -> int:
    docs = collect_tsd_files(DEFAULT_TSD_PACKING_ROOT)
    cand = None
    for d in docs:
        if d.file_name.endswith(TARGET_SUFFIX) or TARGET_SUFFIX in d.file_name:
            cand = d
            break
    lines: list[str] = []
    if cand is None:
        lines.append(f"FILE NOT FOUND ending with {TARGET_SUFFIX!r}")
        OUT.write_text("\n".join(lines), encoding="utf-8")
        print(lines[0])
        return 1

    fp = cand.file_full_path
    lines.append(f"file: {fp}")
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    lines.append(f"sheets: {wb.sheetnames}")
    ws = wb[SHEET]
    lines.append(f"--- raw first 35 non-empty rows on {SHEET!r} (col_index, value) ---")
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=40, max_col=16, values_only=True)):
        cells = [(j, v) for j, v in enumerate(row) if v not in (None, "")]
        if cells:
            lines.append(f"{i + 1}: {cells}")
    wb.close()

    pos, plan, all_rows, hits = load_tsd_file_rows(fp, root=DEFAULT_TSD_PACKING_ROOT)
    lines.append(f"plan data_sheets={plan.data_sheets!r}")
    lines.append(f"hits={[(h.sheet_name, h.position_rows, h.status) for h in hits]}")
    sheet_rows = [r for r in all_rows if r.el[TITLE].value == SHEET]
    lines.append(f"types={dict(Counter(r.row_type for r in sheet_rows))} n={len(sheet_rows)}")
    lines.append("--- classified non-empty rows (first 20) ---")
    shown = 0
    missing_vendor = 0
    missing_code = 0
    for r in sheet_rows:
        if r.row_type == RowType.empty_row:
            continue
        vals = {
            k: r.el[k].value
            for k in (CODE, SPECIFICATION_NAME, TAGS, NAME, VALUES, UNITS, VENDOR)
        }
        nonempty = False
        for v in vals.values():
            if isinstance(v, list):
                if v:
                    nonempty = True
            elif str(v).strip():
                nonempty = True
        if not nonempty:
            continue
        code_ok = bool(str(vals[CODE] or "").strip())
        name_ok = bool(str(vals[NAME] or "").strip())
        values_ok = bool(str(vals[VALUES] or "").strip())
        units_ok = bool(str(vals[UNITS] or "").strip())
        vendor_ok = bool(str(vals[VENDOR] or "").strip())
        if not vendor_ok and code_ok and name_ok and values_ok and units_ok:
            missing_vendor += 1
        if not code_ok and name_ok and values_ok and units_ok:
            missing_code += 1
        if shown < 20:
            lines.append(
                f"{r.row_type} code_ok={code_ok} name_ok={name_ok} "
                f"values_ok={values_ok} units_ok={units_ok} vendor_ok={vendor_ok} "
                f"{ {k: (repr(v)[:70] if not isinstance(v, list) else v) for k, v in vals.items()} }"
            )
            shown += 1
    lines.append(f"other_rows looking like pos but missing VENDOR only: {missing_vendor}")
    lines.append(f"other_rows looking like pos but missing CODE: {missing_code}")
    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"wrote {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
