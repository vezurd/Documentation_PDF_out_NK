"""Extra: comment columns, code->name injectivity, sample comment size."""
from __future__ import annotations

import collections
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _safe(s: str) -> None:
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    sys.stdout.buffer.write((s + "\n").encode(enc, errors="backslashreplace"))


def col_letters(ref: str) -> str:
    return "".join(ch for ch in ref if ch.isalpha())


def load_shared(zf: zipfile.ZipFile) -> list[str]:
    root = ET.parse(zf.open("xl/sharedStrings.xml")).getroot()
    out = []
    for si in root.findall("m:si", NS):
        out.append("".join(t.text or "" for t in si.findall(".//m:t", NS)))
    return out


def main() -> int:
    src = Path(
        r"c:\Users\ydruzev\PycharmProjects\Documentation_PDF_out_NK\tmp\step4_excel_perf_sample.xlsx"
    )
    with zipfile.ZipFile(src) as zf:
        shared = load_shared(zf)
        # comments
        croot = ET.parse(zf.open("xl/comments1.xml")).getroot()
        by_col: collections.Counter[str] = collections.Counter()
        lengths: list[int] = []
        samples: dict[str, str] = {}
        for c in croot.findall("m:commentList/m:comment", NS):
            ref = c.get("ref", "")
            letters = col_letters(ref)
            by_col[letters] += 1
            texts = [t.text or "" for t in c.findall(".//m:t", NS)]
            text = "".join(texts)
            lengths.append(len(text))
            if letters not in samples:
                samples[letters] = text[:400].replace("\n", " | ")
        _safe("=== comments by column ===")
        for col, n in by_col.most_common():
            _safe(f"  {col}: {n:,}  sample={samples.get(col, '')[:180]!r}")
        _safe(
            f"comment text len: n={len(lengths)}, avg={sum(lengths)/len(lengths):.0f}, "
            f"max={max(lengths)}, p50={sorted(lengths)[len(lengths)//2]}"
        )

        # sheet cells for injectivity: headers + CODE/NAME pairs
        headers: dict[str, str] = {}
        pairs: dict[str, dict[str, set[str]]] = {
            "rfp": collections.defaultdict(set),  # code -> names
            "mto": collections.defaultdict(set),
            "ul": collections.defaultdict(set),
        }
        # need column letters from first row
        # From previous diag: E=Код RFP F=Наименование RFP Q=Код MTO R=Наименование MTO AH=Код УЛ AG=Наименование УЛ
        map_code_name = {
            "rfp": ("E", "F"),
            "mto": ("Q", "R"),
            "ul": ("AH", "AG"),
        }
        context = ET.iterparse(zf.open("xl/worksheets/sheet1.xml"), events=("end",))
        ns = NS["m"]
        row_vals: dict[str, str] = {}
        current_row = 0
        for _, elem in context:
            if elem.tag != f"{{{ns}}}c":
                continue
            ref = elem.get("r", "")
            letters = col_letters(ref)
            row_n = int("".join(ch for ch in ref if ch.isdigit()) or 0)
            if row_n != current_row and current_row > 1:
                for kind, (cc, nc) in map_code_name.items():
                    code = (row_vals.get(cc) or "").strip()
                    name = (row_vals.get(nc) or "").strip()
                    if code:
                        pairs[kind][code].add(name)
                row_vals = {}
            current_row = row_n
            cell_type = elem.get("t")
            v = elem.find(f"{{{ns}}}v")
            value = ""
            if cell_type == "s" and v is not None and v.text is not None:
                idx = int(v.text)
                value = shared[idx] if 0 <= idx < len(shared) else ""
            elif v is not None and v.text:
                value = v.text
            if row_n == 1:
                headers[letters] = value
            else:
                row_vals[letters] = value
            elem.clear()
        # last row
        if current_row > 1:
            for kind, (cc, nc) in map_code_name.items():
                code = (row_vals.get(cc) or "").strip()
                name = (row_vals.get(nc) or "").strip()
                if code:
                    pairs[kind][code].add(name)

        _safe("=== headers (confirm) ===")
        for k in ("E", "F", "Q", "R", "AG", "AH", "AE", "AM", "AN", "AI"):
            _safe(f"  {k}={headers.get(k)}")

        _safe("=== code -> name injectivity ===")
        for kind, d in pairs.items():
            multi = {c: ns_ for c, ns_ in d.items() if len(ns_) > 1}
            empty_name = sum(1 for ns_ in d.values() if "" in ns_ and len(ns_) == 1)
            _safe(
                f"  {kind}: unique codes={len(d):,}, "
                f"codes with >1 distinct name={len(multi):,}, "
                f"codes with only empty name={empty_name:,}"
            )
            if multi:
                example = next(iter(multi.items()))
                names = list(example[1])[:3]
                _safe(f"    e.g. {example[0]!r} -> {names!r} (+{len(example[1])-len(names)} more)")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
