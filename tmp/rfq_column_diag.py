"""One-off RFQ column layout diagnostic."""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl

from RFQ.ds_compare.support_finctions import load_rfq_tpk_data

RFQ_PATH = Path(
    r"//bcc/eng/PrDoc/377_НИПИГАЗ/АГХК/КСБ/RFP_MTO_VO/_RFQ/"
    r"AGCC.287-0000-12.4.1-RFQ-0025_Форма ТКП 2076961 43_ДИКТ.xlsx"
)
OUT = Path(__file__).resolve().parent / "rfq_column_diag.txt"


def main() -> None:
    lines: list[str] = []
    rows = load_rfq_tpk_data(str(RFQ_PATH))
    pos = [r for r in rows if r.row_type == "position_row"]
    lines.append(f"position rows: {len(pos)}")

    bcc = re.compile(r"^BCC\d+", re.I)
    fields = [
        "code",
        "name",
        "DS_CODE_1C",
        "type_mark",
        "DS_TITLE",
        "DS_SYSTEM",
        "values",
        "DS_SPECIFICATION",
        "tags",
    ]
    for field in fields:
        samples: list[str] = []
        bcc_count = 0
        for row in pos:
            value = str(row.get_value(field) or "").strip()
            if bcc.match(value):
                bcc_count += 1
            if value and len(samples) < 5:
                samples.append(value[:100])
        lines.append(f"{field}: bcc_count={bcc_count}, samples={samples!r}")

    hits = [
        r
        for r in pos
        if str(r.get_value("DS_TITLE") or "").strip() == "8525"
        and str(r.get_value("DS_SYSTEM") or "").strip() == "SKUD"
    ]
    lines.append(f"8525/SKUD rows: {len(hits)}")
    if hits:
        row = hits[0]
        for field in fields:
            value = str(row.get_value(field) or "")[:120]
            lines.append(f"  8525 sample {field}={value!r}")

    wb = openpyxl.load_workbook(str(RFQ_PATH), read_only=True, data_only=True)
    ws = wb.active
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=400, values_only=True), start=1):
        vals = [row[j] if j < len(row) else None for j in range(13)]
        text = " ".join(str(v) for v in vals if v is not None)
        if "BCC0000545" in text or i <= 25:
            cols = [repr(str(v)[:50] if v is not None else "") for v in vals]
            lines.append(f"excel row {i}: " + " | ".join(cols))
        if "BCC0000545" in text:
            break
    wb.close()

    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"written {OUT}")


if __name__ == "__main__":
    main()
