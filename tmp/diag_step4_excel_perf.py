"""Read-only diagnostics of a Step4 xlsx: zip parts, styles, strings, filters.

Does not modify the workbook. Writes a UTF-8 markdown report next to the file.
"""
from __future__ import annotations

import collections
import io
import os
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _safe(s: str) -> None:
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    sys.stdout.buffer.write((s + "\n").encode(enc, errors="backslashreplace"))


def _col_letters(cell_ref: str) -> str:
    return "".join(ch for ch in cell_ref if ch.isalpha())


def analyze_zip(path: Path) -> dict:
    info: dict = {"parts": [], "total_comp": 0, "total_uncomp": 0}
    with zipfile.ZipFile(path) as zf:
        for zi in sorted(zf.infolist(), key=lambda x: -x.file_size):
            rec = {
                "name": zi.filename,
                "comp": zi.compress_size,
                "uncomp": zi.file_size,
                "ratio": (zi.compress_size / zi.file_size) if zi.file_size else 0,
            }
            info["parts"].append(rec)
            info["total_comp"] += zi.compress_size
            info["total_uncomp"] += zi.file_size
    return info


def parse_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    name = "xl/sharedStrings.xml"
    if name not in zf.namelist():
        return []
    root = ET.parse(zf.open(name)).getroot()
    strings: list[str] = []
    for si in root.findall("m:si", NS):
        texts = [t.text or "" for t in si.findall(".//m:t", NS)]
        strings.append("".join(texts))
    return strings


def parse_styles(zf: zipfile.ZipFile) -> dict:
    name = "xl/styles.xml"
    if name not in zf.namelist():
        return {}
    root = ET.parse(zf.open(name)).getroot()
    fills = root.find("m:fills", NS)
    fonts = root.find("m:fonts", NS)
    borders = root.find("m:borders", NS)
    xfs = root.find("m:cellXfs", NS)
    dxfs = root.find("m:dxfs", NS)
    return {
        "fills": int(fills.get("count", 0)) if fills is not None else 0,
        "fonts": int(fonts.get("count", 0)) if fonts is not None else 0,
        "borders": int(borders.get("count", 0)) if borders is not None else 0,
        "cellXfs": int(xfs.get("count", 0)) if xfs is not None else 0,
        "dxfs": int(dxfs.get("count", 0)) if dxfs is not None else 0,
    }


def parse_workbook_sheets(zf: zipfile.ZipFile) -> list[tuple[str, str]]:
    root = ET.parse(zf.open("xl/workbook.xml")).getroot()
    rels = ET.parse(zf.open("xl/_rels/workbook.xml.rels")).getroot()
    rel_map = {
        rel.get("Id"): rel.get("Target")
        for rel in rels
    }
    sheets = []
    for sh in root.findall("m:sheets/m:sheet", NS):
        rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rel_map.get(rid, "")
        if target and not target.startswith("xl/"):
            target = "xl/" + target.lstrip("/")
        sheets.append((sh.get("name", ""), target))
    return sheets


def count_hyperlinks_comments(zf: zipfile.ZipFile) -> dict:
    counts = {
        "hyperlinks_xml": 0,
        "comments_xml": 0,
        "vml_drawings": 0,
        "rels_hyperlinks": 0,
    }
    for name in zf.namelist():
        if name.endswith(".xml") and "comments" in name.lower():
            root = ET.parse(zf.open(name)).getroot()
            counts["comments_xml"] += len(root.findall("m:commentList/m:comment", NS))
        if name.endswith(".vml") or "vmlDrawing" in name:
            counts["vml_drawings"] += 1
        if "/worksheets/" in name and name.endswith(".xml") and "_rels" not in name:
            root = ET.parse(zf.open(name)).getroot()
            hl = root.find("m:hyperlinks", NS)
            if hl is not None:
                counts["hyperlinks_xml"] += len(list(hl))
        if name.endswith(".rels") and "/worksheets/_rels/" in name:
            root = ET.parse(zf.open(name)).getroot()
            for rel in root:
                rtype = (rel.get("Type") or "").lower()
                if "hyperlink" in rtype:
                    counts["rels_hyperlinks"] += 1
    return counts


def analyze_sheet(
    zf: zipfile.ZipFile,
    sheet_path: str,
    shared: list[str],
    header_row: int = 1,
    sample_unique: bool = True,
) -> dict:
    """Stream-parse worksheet XML: dim, cells, unique values per column."""
    ns_main = NS["m"]
    dim = ""
    n_cells = 0
    n_styled = 0
    n_inline_str = 0
    n_shared = 0
    n_number = 0
    n_formula = 0
    n_blank = 0
    style_ids: collections.Counter[str] = collections.Counter()
    col_uniques: dict[str, set] = collections.defaultdict(set)
    col_counts: collections.Counter[str] = collections.Counter()
    col_bytes: collections.Counter[str] = collections.Counter()
    headers: dict[str, str] = {}
    max_row = 0
    has_autofilter = False
    has_merge = 0
    has_outline = False

    # Iterparse to keep memory reasonable
    context = ET.iterparse(zf.open(sheet_path), events=("end",))
    for _, elem in context:
        tag = elem.tag
        if tag == f"{{{ns_main}}}dimension":
            dim = elem.get("ref", "")
        elif tag == f"{{{ns_main}}}autoFilter":
            has_autofilter = True
        elif tag == f"{{{ns_main}}}mergeCell":
            has_merge += 1
        elif tag == f"{{{ns_main}}}col":
            if elem.get("outlineLevel"):
                has_outline = True
        elif tag == f"{{{ns_main}}}c":
            n_cells += 1
            ref = elem.get("r", "")
            letters = _col_letters(ref)
            row_n = int("".join(ch for ch in ref if ch.isdigit()) or 0)
            max_row = max(max_row, row_n)
            sid = elem.get("s")
            if sid:
                n_styled += 1
                style_ids[sid] += 1
            cell_type = elem.get("t")
            v = elem.find(f"{{{ns_main}}}v")
            is_el = elem.find(f"{{{ns_main}}}is")
            f_el = elem.find(f"{{{ns_main}}}f")
            value = ""
            if f_el is not None:
                n_formula += 1
                value = (f_el.text or "")[:80]
            elif cell_type == "s" and v is not None and v.text is not None:
                n_shared += 1
                idx = int(v.text)
                value = shared[idx] if 0 <= idx < len(shared) else ""
            elif cell_type == "inlineStr" or is_el is not None:
                n_inline_str += 1
                texts = [t.text or "" for t in is_el.findall(f".//{{{ns_main}}}t")] if is_el is not None else []
                value = "".join(texts)
            elif v is not None and v.text is not None:
                n_number += 1
                value = v.text
            else:
                n_blank += 1
            if letters:
                col_counts[letters] += 1
                col_bytes[letters] += len(value)
                if row_n == header_row:
                    headers[letters] = value
                elif sample_unique and row_n > header_row:
                    # cap unique set memory
                    s = col_uniques[letters]
                    if len(s) < 50000:
                        s.add(value)
            elem.clear()

    n_cols = len(col_counts)
    col_stats = []
    for letters, cnt in sorted(col_counts.items(), key=lambda kv: (len(kv[0]), kv[0])):
        uniques = col_uniques.get(letters, set())
        nuniq = len(uniques)
        capped = nuniq >= 50000
        avg_len = col_bytes[letters] / cnt if cnt else 0
        col_stats.append({
            "col": letters,
            "header": headers.get(letters, ""),
            "cells": cnt,
            "unique": nuniq,
            "unique_capped": capped,
            "avg_chars": round(avg_len, 1),
            "total_chars": col_bytes[letters],
        })

    return {
        "path": sheet_path,
        "dimension": dim,
        "max_row": max_row,
        "n_cols": n_cols,
        "n_cells": n_cells,
        "n_styled": n_styled,
        "n_shared": n_shared,
        "n_inline_str": n_inline_str,
        "n_number": n_number,
        "n_formula": n_formula,
        "n_blank": n_blank,
        "style_ids_used": len(style_ids),
        "top_styles": style_ids.most_common(8),
        "autofilter": has_autofilter,
        "merges": has_merge,
        "outline": has_outline,
        "col_stats": col_stats,
    }


def fmt_mb(n: int) -> str:
    return f"{n / (1024 * 1024):.2f} MB"


def main() -> int:
    src = Path(
        r"c:\Users\ydruzev\PycharmProjects\Documentation_PDF_out_NK\tmp\step4_excel_perf_sample.xlsx"
    )
    if len(sys.argv) > 1:
        src = Path(sys.argv[1])
    if not src.exists():
        _safe(f"missing: {src}")
        return 1

    zip_info = analyze_zip(src)
    lines: list[str] = []
    lines.append("# Step4 Excel — диагностика объёма и фильтров")
    lines.append("")
    lines.append(f"- Файл: `{src.name}`")
    lines.append(f"- На диске (zip): **{fmt_mb(src.stat().st_size)}** ({src.stat().st_size:,} байт)")
    lines.append(
        f"- Распакованный XML: **{fmt_mb(zip_info['total_uncomp'])}** "
        f"(сжатие {zip_info['total_comp'] / zip_info['total_uncomp']:.2f})"
    )
    lines.append("")
    lines.append("## Крупнейшие части zip")
    lines.append("")
    lines.append("| Часть | Uncompressed | Compressed | Ratio |")
    lines.append("|---|---:|---:|---:|")
    for rec in zip_info["parts"][:25]:
        if rec["uncomp"] < 50_000 and rec["name"] not in (
            "xl/sharedStrings.xml",
            "xl/styles.xml",
            "xl/workbook.xml",
        ):
            continue
        lines.append(
            f"| `{rec['name']}` | {fmt_mb(rec['uncomp'])} | {fmt_mb(rec['comp'])} | {rec['ratio']:.2f} |"
        )

    with zipfile.ZipFile(src) as zf:
        shared = parse_shared_strings(zf)
        styles = parse_styles(zf)
        sheets = parse_workbook_sheets(zf)
        extra = count_hyperlinks_comments(zf)

        # shared string stats
        ss_len = [len(s) for s in shared]
        ss_counter = collections.Counter(shared)
        lines.append("")
        lines.append("## Shared strings и стили")
        lines.append("")
        lines.append(f"- Уникальных строк в `sharedStrings.xml`: **{len(shared):,}**")
        if ss_len:
            lines.append(
                f"- Длина: avg={sum(ss_len)/len(ss_len):.1f}, "
                f"median={sorted(ss_len)[len(ss_len)//2]}, max={max(ss_len)}"
            )
            long_n = sum(1 for n in ss_len if n >= 80)
            lines.append(f"- Строк ≥80 символов: {long_n:,}")
        dup_in_ss = sum(1 for _, c in ss_counter.items() if c > 1)
        lines.append(f"- Дубликаты внутри sharedStrings (не должно быть): {dup_in_ss}")
        lines.append(
            f"- Стили: fonts={styles.get('fonts')}, fills={styles.get('fills')}, "
            f"borders={styles.get('borders')}, **cellXfs={styles.get('cellXfs')}**, "
            f"dxfs={styles.get('dxfs')}"
        )
        lines.append(
            f"- Гиперссылки (worksheet XML): **{extra['hyperlinks_xml']:,}**; "
            f"rels: {extra['rels_hyperlinks']:,}"
        )
        lines.append(
            f"- Комментарии: **{extra['comments_xml']:,}**; VML drawings: {extra['vml_drawings']}"
        )
        lines.append("")
        lines.append("## Листы")
        lines.append("")
        sheet_reports = []
        for name, target in sheets:
            _safe(f"parsing sheet {name!r} -> {target}")
            report = analyze_sheet(zf, target, shared)
            report["name"] = name
            sheet_reports.append(report)
            lines.append(
                f"### {name}"
            )
            lines.append("")
            lines.append(
                f"- dim=`{report['dimension']}`, max_row={report['max_row']:,}, "
                f"cols={report['n_cols']}, cells={report['n_cells']:,}"
            )
            lines.append(
                f"- shared={report['n_shared']:,}, numbers={report['n_number']:,}, "
                f"inlineStr={report['n_inline_str']:,}, formulas={report['n_formula']:,}, "
                f"blank={report['n_blank']:,}, styled={report['n_styled']:,}"
            )
            lines.append(
                f"- autofilter={report['autofilter']}, outline={report['outline']}, "
                f"merges={report['merges']}, unique style ids used={report['style_ids_used']}"
            )
            data_rows = max(report["max_row"] - 1, 0)
            if data_rows and report["n_cols"]:
                lines.append(
                    f"- Плотность: {report['n_cells'] / (data_rows * report['n_cols']):.2%} "
                    f"ячеек от сетки {data_rows:,}×{report['n_cols']}"
                )
            lines.append("")
            if report["col_stats"]:
                lines.append(
                    "| Col | Заголовок | Unique | Unique/rows | Avg chars | Total chars |"
                )
                lines.append("|---|---|---:|---:|---:|---:|")
                for st in report["col_stats"]:
                    ratio = st["unique"] / data_rows if data_rows else 0
                    cap = "+" if st["unique_capped"] else ""
                    hdr = (st["header"] or "").replace("|", "/")
                    lines.append(
                        f"| {st['col']} | {hdr} | {st['unique']:,}{cap} | {ratio:.1%} | "
                        f"{st['avg_chars']} | {st['total_chars']:,} |"
                    )
                lines.append("")

    out = src.with_suffix(".md").with_name(src.stem + "_diag.md")
    # write next to sample
    out = src.parent / "diag_step4_excel_perf.md"
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    _safe(f"wrote {out}")
    for line in lines[:40]:
        _safe(line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
