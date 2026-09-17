"""Refine UL↔RFP coverage from already scanned tmp/ul_rfp_ds_coverage.json."""
from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "tmp" / "ul_rfp_ds_coverage.json"
OUT = ROOT / "tmp" / "ul_rfp_ds_join.json"

# Real corrections look like «ДС92_24Б. AGCC» / «ДС48_14. » — dot+space or space.
# Reject «ДС4905_1.xlsx» (copy suffix before extension).
_DS_COMPOUND_FILE = re.compile(
    r"ДС\s*(\d+)\s*[_\-]\s*(\d+)([А-ЯA-Zа-яa-z]*)(?=\.\s|\s)",
    re.IGNORECASE,
)
_DS_SIMPLE = re.compile(r"ДС\s*(\d+)", re.IGNORECASE)
_DS_FOLDER_COMPOUND = re.compile(
    r"ДС\s*(\d+)\s*[_\-]\s*(\d+)([А-ЯA-Zа-яa-z]*)",
    re.IGNORECASE,
)
_BARE_UL_NUM = re.compile(r"УЛ\s+(\d+)\s*$", re.IGNORECASE)
_CYR = str.maketrans({"A": "А", "B": "В", "E": "Е", "K": "К", "M": "М", "H": "Н", "O": "О", "P": "Р", "C": "С", "T": "Т", "X": "Х"})


def _letter(s: str) -> str:
    return (s or "").upper().translate(_CYR)


def parse_rfp_name(name: str) -> dict:
    m = _DS_COMPOUND_FILE.search(name)
    if m:
        seq, orig, let = int(m.group(1)), int(m.group(2)), _letter(m.group(3))
        return {
            "sequential": seq,
            "original": orig,
            "letter": let,
            "compound": True,
            "match_key": orig,
            "label": f"ДС{seq}_{orig}{let}",
        }
    m = _DS_SIMPLE.search(name)
    if m:
        n = int(m.group(1))
        return {
            "sequential": n,
            "original": n,
            "letter": "",
            "compound": False,
            "match_key": n,
            "label": f"ДС{n}",
        }
    return {
        "sequential": None,
        "original": None,
        "letter": "",
        "compound": False,
        "match_key": None,
        "label": "",
    }


def parse_ul_folder(name: str) -> dict:
    upper = name.upper()
    if "ГФ" in upper or "ГОСФИН" in upper or "ГОС.ФИН" in upper:
        return {
            "kind": "gf",
            "sequential": None,
            "original": None,
            "letter": "",
            "compound": False,
            "match_keys": [],
            "label": "ГФ",
        }
    m = _DS_FOLDER_COMPOUND.search(name)
    if m:
        seq, orig, let = int(m.group(1)), int(m.group(2)), _letter(m.group(3))
        return {
            "kind": "compound",
            "sequential": seq,
            "original": orig,
            "letter": let,
            "compound": True,
            "match_keys": [seq, orig],
            "label": f"ДС{seq}_{orig}{let}",
        }
    m = _DS_SIMPLE.search(name)
    if m:
        n = int(m.group(1))
        return {
            "kind": "simple",
            "sequential": n,
            "original": n,
            "letter": "",
            "compound": False,
            "match_keys": [n],
            "label": f"ДС{n}",
        }
    m = _BARE_UL_NUM.search(name)
    if m:
        n = int(m.group(1))
        return {
            "kind": "bare",
            "sequential": n,
            "original": n,
            "letter": "",
            "compound": False,
            "match_keys": [n],
            "label": f"ДС{n}",
        }
    return {
        "kind": "unparsed",
        "sequential": None,
        "original": None,
        "letter": "",
        "compound": False,
        "match_keys": [],
        "label": "",
    }


def _short(names: list[str], n: int = 2) -> str:
    if not names:
        return "—"
    shown = names[:n]
    extra = f" +{len(names) - n}" if len(names) > n else ""
    return "; ".join(shown) + extra


def main() -> int:
    data = json.loads(SRC.read_text(encoding="utf-8"))
    ul_src = data["ul_folders"]
    rfp_src = data["rfp_files"]

    ul = []
    for rec in ul_src:
        p = parse_ul_folder(rec["name"])
        ul.append({**p, "folder": rec["name"], "xlsx": rec["xlsx_count"]})

    rfp = []
    for rec in rfp_src:
        p = parse_rfp_name(rec["name"])
        rfp.append({**p, "file": rec["name"]})

    rfp_by_orig: dict[int, list] = defaultdict(list)
    rfp_by_seq: dict[int, list] = defaultdict(list)
    for rec in rfp:
        if rec["original"] is not None:
            rfp_by_orig[rec["original"]].append(rec)
        if rec["sequential"] is not None:
            rfp_by_seq[rec["sequential"]].append(rec)

    ul_by_key: dict[int, list] = defaultdict(list)
    for rec in ul:
        for k in rec["match_keys"]:
            ul_by_key[k].append(rec)

    rows = []
    used_rfp: set[str] = set()
    used_ul: set[str] = set()

    # Primary join: UL folder actual DS (first number for simple/bare;
    # both numbers for compound folders) vs RFP original, fallback RFP sequential.
    for rec in sorted(ul, key=lambda x: (x["sequential"] is None, x["sequential"] or 10**9, x["folder"])):
        if rec["kind"] == "gf":
            gf_rfp = [
                x
                for x in rfp
                if x["sequential"] in {1, 2, 7}
                or "ГФ" in x["file"].upper()
                or "ГОСФИН" in x["file"].upper()
            ]
            for item in gf_rfp:
                used_rfp.add(item["file"])
            rows.append(
                {
                    "ds": "ГФ",
                    "status": "gf_special",
                    "status_ru": "Особая группа ГФ",
                    "ul_folder": rec["folder"],
                    "ul_xlsx": rec["xlsx"],
                    "ul_parse": rec["label"],
                    "rfp_files": [x["file"] for x in gf_rfp],
                    "rfp_labels": [x["label"] for x in gf_rfp],
                    "rfp_seq": sorted(
                        {x["sequential"] for x in gf_rfp if x["sequential"]}
                    ),
                    "note": (
                        "Папка госфинансирования, не номер ДС. "
                        "В RFP: ДС1_Госфин, ДС2, ДС7_ГФ — 1:1 по номеру нет."
                    ),
                }
            )
            used_ul.add(rec["folder"])
            continue
        if not rec["match_keys"]:
            rows.append(
                {
                    "ds": "?",
                    "status": "ul_unparsed",
                    "status_ru": "УЛ без номера",
                    "ul_folder": rec["folder"],
                    "ul_xlsx": rec["xlsx"],
                    "ul_parse": rec["label"] or "—",
                    "rfp_files": [],
                    "rfp_labels": [],
                    "rfp_seq": [],
                    "note": "Номер ДС из имени папки не извлечён.",
                }
            )
            used_ul.add(rec["folder"])
            continue

        hits: list = []
        how = ""
        # Prefer original-number match (UL folder DS24 ↔ RFP ДС92_24Б)
        for k in rec["match_keys"]:
            for item in rfp_by_orig.get(k, []):
                if item["file"] not in {h["file"] for h in hits}:
                    hits.append(item)
            if hits and not how:
                how = "по фактическому"

        # Compound UL ДС4_11: also allow sequential 4 → RFP ДС4
        if rec["compound"]:
            seq_hits = rfp_by_seq.get(rec["sequential"], [])
            for item in seq_hits:
                if item["file"] not in {h["file"] for h in hits}:
                    hits.append(item)
                    how = (how + " + порядковый").strip(" +") if how else "по порядковому UL"

        if hits:
            status = "both"
            status_ru = "Есть в УЛ и RFP"
            note = f"Связка {how}."
            if rec["compound"]:
                note += f" Папка составная {rec['label']}: порядковый {rec['sequential']}, исходный {rec['original']}{rec['letter']}."
            if any(h["compound"] for h in hits):
                corr = [h["label"] for h in hits if h["compound"]]
                note += f" RFP-корректировка: {', '.join(corr)}."
        else:
            status = "ul_only"
            status_ru = "Только УЛ"
            note = "Папка УЛ есть, файла RFP с этим фактическим номером нет."

        for h in hits:
            used_rfp.add(h["file"])
        used_ul.add(rec["folder"])
        rows.append(
            {
                "ds": rec["label"],
                "status": status,
                "status_ru": status_ru,
                "ul_folder": rec["folder"],
                "ul_xlsx": rec["xlsx"],
                "ul_parse": rec["label"],
                "rfp_files": [h["file"] for h in hits],
                "rfp_labels": [h["label"] for h in hits],
                "rfp_seq": sorted({h["sequential"] for h in hits}),
                "note": note,
            }
        )

    # RFP files not used
    leftover_groups: dict[int, list] = defaultdict(list)
    for rec in rfp:
        if rec["file"] in used_rfp:
            continue
        leftover_groups[rec["match_key"] if rec["match_key"] is not None else -1].append(rec)

    for key in sorted(leftover_groups, key=lambda x: (x < 0, x)):
        items = leftover_groups[key]
        labels = sorted({x["label"] for x in items})
        note = "Файл RFP есть, папки УЛ с этим фактическим номером нет."
        if any("ПРИЛОЖЕНИЕ" in x["file"].upper() for x in items):
            note += " Несколько приложений одной ДС47."
        rows.append(
            {
                "ds": labels[0] if labels else "?",
                "status": "rfp_only",
                "status_ru": "Только RFP",
                "ul_folder": "",
                "ul_xlsx": 0,
                "ul_parse": "—",
                "rfp_files": [x["file"] for x in items],
                "rfp_labels": [x["label"] for x in items],
                "rfp_seq": sorted({x["sequential"] for x in items if x["sequential"] is not None}),
                "note": note,
            }
        )

    both = sum(1 for r in rows if r["status"] == "both")
    ul_only = sum(1 for r in rows if r["status"] == "ul_only")
    rfp_only = sum(1 for r in rows if r["status"] == "rfp_only")
    gf = sum(1 for r in rows if r["status"] == "gf_special")
    compounds_rfp = [x for x in rfp if x["compound"]]
    compounds_ul = [x for x in ul if x["compound"]]

    payload = {
        "ul_root": data["ul_root"],
        "rfp_root": data["rfp_root"],
        "ul_folders": len(ul),
        "ul_xlsx": data["summary"]["ul_xlsx_files"],
        "rfp_files": len(rfp),
        "rfp_compound": len(compounds_rfp),
        "ul_compound": len(compounds_ul),
        "both": both,
        "ul_only": ul_only,
        "rfp_only": rfp_only,
        "gf": gf,
        "need_parser": True,
        "join_rule": (
            "Ключ сверки = фактический номер ДС. "
            "У папки УЛ это номер в имени (ДС24, 4905). "
            "У файла RFP при ДС92_24Б это 24, иначе сам порядковый. "
            "Порядковый номер RFP (92) с папками УЛ не совпадает."
        ),
        "compound_rfp": [
            {
                "label": x["label"],
                "sequential": x["sequential"],
                "original": x["original"],
                "letter": x["letter"],
                "file": x["file"],
            }
            for x in sorted(compounds_rfp, key=lambda z: z["sequential"])
        ],
        "compound_ul": [
            {
                "label": x["label"],
                "sequential": x["sequential"],
                "original": x["original"],
                "folder": x["folder"],
                "xlsx": x["xlsx"],
            }
            for x in compounds_ul
        ],
        "rows": rows,
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"rows={len(rows)} both={both} ul_only={ul_only} rfp_only={rfp_only} gf={gf}")
    print(f"wrote {OUT}")
    for r in rows:
        if r["status"] != "both":
            print(f"  {r['status']:12} {r['ds']:12} ul={r['ul_folder'][:40]:40} rfp={_short(r['rfp_labels'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
