"""List GF UL folders and RFP files for DS1/2/7 (read-only)."""
from __future__ import annotations

import json
from pathlib import Path

UL = Path(r"\\bcc\root\CurProjects\di_manegers\Пятницкий_П\Амурский ГХК\Поставки\ТСД по всем ДС")
RFP = Path(r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\RFP_Зиновьев")
OUT = Path(__file__).with_name("_list_gf_ul.json")


def main() -> None:
    folders = []
    for child in sorted(UL.iterdir(), key=lambda p: p.name.lower()):
        if not child.is_dir():
            continue
        name = child.name
        if "ГФ" not in name.upper() and "ГОСФИН" not in name.upper() and "ГОС.ФИН" not in name.upper():
            continue
        files = sorted(p.name for p in child.iterdir() if p.is_file() and not p.name.startswith("~$"))
        folders.append({"folder": name, "files": files, "n": len(files)})

    rfp = []
    if RFP.exists():
        for p in sorted(RFP.iterdir(), key=lambda x: x.name.lower()):
            n = p.name.upper()
            if p.suffix.lower() in {".xlsx", ".xlsm", ".xls"} and (
                n.startswith("ДС1") or n.startswith("ДС2") or n.startswith("ДС7")
                or "ГФ" in n or "ГОСФИН" in n
            ):
                rfp.append(p.name)

    payload = {
        "ul_exists": UL.exists(),
        "rfp_exists": RFP.exists(),
        "folders": folders,
        "rfp_candidates": rfp,
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"wrote {OUT} folders={len(folders)} rfp={len(rfp)}")


if __name__ == "__main__":
    main()
