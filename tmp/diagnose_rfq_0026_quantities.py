"""One-off: RFQ_0026 quantity audit (run from repo root on machine with UNC access)."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

RFQ_XLSX = Path(
    r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFQ\RFQ_0026\ТКП № 7768 от 08 05 2026.xlsx"
)


def main() -> None:
    if not RFQ_XLSX.is_file():
        print("File not found:", RFQ_XLSX)
        sys.exit(1)

    from RFQ.ds_compare.ds_rfq_quantity_audit import (
        compute_rfq_quantity_audit,
        sum_excel_column_k_numeric,
        sum_non_position_values,
        sum_position_values,
    )
    from RFQ.ds_compare.support_finctions import load_rfq_tpk_data

    path = str(RFQ_XLSX)
    rows = load_rfq_tpk_data(path, dump_non_position_dir=str(RFQ_XLSX.parent))
    ex, ec = sum_excel_column_k_numeric(path)
    ps, pc = sum_position_values(rows)
    oth_s, oth_c = sum_non_position_values(rows)
    print(f"Excel K sum: {ex:g} ({ec} cells)")
    print(f"position_row: {ps:g} ({pc} rows)")
    print(f"non-position with VALUES: {oth_s:g} ({oth_c} rows)")
    print("See RFQ_TPK_not_position_rows.txt for ИТОГО / head rows (e.g. 50925)")


if __name__ == "__main__":
    main()
