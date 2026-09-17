"""Parse GF-related RFP names vs UL folder identities."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.rfp_parts.ds_identity import parse_rfp_ds_identity, parse_ul_folder_ds_identity

RFP = Path(r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\RFP_Зиновьев")
OUT = Path(__file__).with_name("_list_gf_rfp.txt")

lines: list[str] = ["=== RFP ==="]
for p in sorted(RFP.iterdir(), key=lambda x: x.name.lower()):
    if p.suffix.lower() not in {".xlsx", ".xlsm", ".xls"} or p.name.startswith("~$"):
        continue
    ident = parse_rfp_ds_identity(p.name)
    if ident.actual in {1, 2, 7} or "ГФ" in p.name.upper() or "ГОС" in p.name.upper():
        lines.append(
            f"{p.name} | seq={ident.sequential} act={ident.actual} kind={ident.kind}"
        )

lines.append("=== UL folders ===")
for name in ("согл УЛ ДС1 ГФ 5титулов", "согл УЛ ДС7 ГФ 2 тит"):
    ident = parse_ul_folder_ds_identity(name)
    lines.append(
        f"{name} | kind={ident.kind} act={ident.actual} seq={ident.sequential}"
    )

OUT.write_text("\n".join(lines), encoding="utf-8")
print(f"wrote {OUT} lines={len(lines)}")
