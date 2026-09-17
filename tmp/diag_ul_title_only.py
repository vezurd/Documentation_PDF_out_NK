"""Diagnose packing-only Step4 rows whose Титул/Марка looks title-only."""

from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tmp" / "ul_title_diag_20260818_1620" / "diag.txt"

STEP4 = Path(
    r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO"
    r"\_РЕЗУЛЬТАТА_ПРОВЕРКИ\_результат_проверки_2026.08.18.16.20"
    r"\Шаг4_Сопоставление_RFP_MTO_20260818_162221.xlsx"
)
SUMMARY = Path(
    r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO"
    r"\_RFP\УЛ сводный файл\tsd_packing_summary_2026.08.12_10.30.20.xlsx"
)

COMPOSITE_RE = re.compile(
    r"^[0-9]+-[A-Za-zА-Яа-я0-9]+$",
)
TITLE_ONLY_RE = re.compile(r"^[0-9]+$")
KEY_RE = re.compile(r"Ключ УЛ:\s*(\([^)]+\))")


def _write(lines: list[str]) -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _colmap(ws) -> dict[str, int]:
    headers = {}
    for cell in ws[1]:
        if cell.value:
            headers[str(cell.value).strip()] = cell.column
    return headers


def main() -> int:
    lines: list[str] = []
    lines.append(f"step4 exists={STEP4.exists()} size={STEP4.stat().st_size if STEP4.exists() else 0}")
    lines.append(f"summary exists={SUMMARY.exists()}")

    wb = openpyxl.load_workbook(STEP4, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    headers = _colmap(ws)
    lines.append("headers: " + ", ".join(f"{k}={v}" for k, v in headers.items()))
    title_col = headers.get("Титул/Марка")
    status_col = headers.get("Статус УЛ")
    code_ul_col = headers.get("Код УЛ")
    source_col = headers.get("Источник УЛ (файл · вкладка · строка)")
    name_col = headers.get("Наименование УЛ")
    if not title_col or not status_col:
        lines.append("MISSING TITLE/STATUS COLUMNS")
        wb.close()
        _write(lines)
        return 1

    status_counts: Counter[str] = Counter()
    packing_only_title_kinds: Counter[str] = Counter()
    packing_only_titles: Counter[str] = Counter()
    samples_title_only: list[str] = []
    samples_composite: list[str] = []
    samples_other: list[str] = []
    title_only_8527 = 0
    packing_only = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        status = str(row[status_col - 1] or "").strip()
        status_counts[status or "(empty)"] += 1
        if status != "Только в УЛ":
            continue
        packing_only += 1
        title = str(row[title_col - 1] or "").strip()
        if COMPOSITE_RE.fullmatch(title):
            kind = "composite"
            if len(samples_composite) < 8:
                samples_composite.append(
                    f"{title!r} code={row[code_ul_col - 1]!r} src={row[source_col - 1]!r}"
                )
        elif TITLE_ONLY_RE.fullmatch(title):
            kind = "title_only"
            packing_only_titles[title] += 1
            if title == "8527":
                title_only_8527 += 1
            if len(samples_title_only) < 12:
                samples_title_only.append(
                    f"{title!r} code={row[code_ul_col - 1]!r} "
                    f"name={row[name_col - 1]!r} src={row[source_col - 1]!r}"
                )
        elif not title:
            kind = "empty"
        else:
            kind = "other"
            packing_only_titles[title] += 1
            if len(samples_other) < 8:
                samples_other.append(
                    f"{title!r} code={row[code_ul_col - 1]!r} src={row[source_col - 1]!r}"
                )
        packing_only_title_kinds[kind] += 1

    wb.close()
    lines.append("")
    lines.append("STATUS COUNTS:")
    for key, value in status_counts.most_common():
        lines.append(f"  {key}: {value}")
    lines.append(f"packing_only={packing_only}")
    lines.append("packing-only Титул/Марка kinds:")
    for key, value in packing_only_title_kinds.most_common():
        lines.append(f"  {kind_label(key)}: {value}")
    lines.append(f"title-only 8527 packing-only rows: {title_only_8527}")
    lines.append("top title-only / other values:")
    for key, value in packing_only_titles.most_common(20):
        lines.append(f"  {key!r}: {value}")
    lines.append("samples title_only:")
    lines.extend(f"  {s}" for s in samples_title_only)
    lines.append("samples composite:")
    lines.extend(f"  {s}" for s in samples_composite)
    lines.append("samples other:")
    lines.extend(f"  {s}" for s in samples_other)

    # Comments need a non-read_only pass; sample first 200 packing-only with comments.
    lines.append("")
    lines.append("COMMENT SAMPLE (first packing-only with comment):")
    wb = openpyxl.load_workbook(STEP4, read_only=False, data_only=True)
    ws = wb[wb.sheetnames[0]]
    comment_keys: Counter[str] = Counter()
    comment_samples = 0
    empty_system_in_key = 0
    nonempty_system_in_key = 0
    for row_idx in range(2, ws.max_row + 1):
        status = ws.cell(row_idx, status_col).value
        if str(status or "").strip() != "Только в УЛ":
            continue
        cell = ws.cell(row_idx, status_col)
        text = cell.comment.text if cell.comment else ""
        match = KEY_RE.search(text.replace("\n", " "))
        if not match:
            continue
        key = match.group(1)
        comment_keys[key[:80]] += 1
        # ('title', 'system', 'code')
        parts = [p.strip().strip("'\"") for p in key.strip("()").split(",")]
        system = parts[1] if len(parts) >= 2 else ""
        if system:
            nonempty_system_in_key += 1
        else:
            empty_system_in_key += 1
        if comment_samples < 10:
            title = ws.cell(row_idx, title_col).value
            lines.append(f"  excel_title={title!r} key={key} comment_head={text.splitlines()[:3]}")
            comment_samples += 1
        if nonempty_system_in_key + empty_system_in_key >= 400:
            break
    wb.close()
    lines.append(f"sampled keys with system: {nonempty_system_in_key}")
    lines.append(f"sampled keys with EMPTY system: {empty_system_in_key}")
    lines.append("top sampled keys:")
    for key, value in comment_keys.most_common(12):
        lines.append(f"  {key}: {value}")

    if SUMMARY.exists():
        lines.append("")
        lines.append("PACKING SUMMARY for title 8527:")
        swb = openpyxl.load_workbook(SUMMARY, read_only=True, data_only=True)
        sws = swb[swb.sheetnames[0]]
        sheaders = _colmap(sws)
        lines.append("summary headers: " + ", ".join(sheaders))
        tcol = sheaders.get("Титул")
        scol = sheaders.get("Марка")
        spec_col = sheaders.get("Спецификация") or sheaders.get("Имя спецификации")
        src_col = sheaders.get("Файл") or sheaders.get("Источник")
        # fallback by position from _SUMMARY_COLUMNS: 0 source, 1 sheet, 5 title, 6 system
        n8527 = 0
        systems: Counter[str] = Counter()
        specs: Counter[str] = Counter()
        empty_sys = 0
        samples = []
        for row in sws.iter_rows(min_row=2, values_only=True):
            title = str(row[(tcol or 6) - 1] or "").strip()
            if title != "8527":
                continue
            n8527 += 1
            system = str(row[(scol or 7) - 1] or "").strip()
            systems[system or "(empty)"] += 1
            if not system:
                empty_sys += 1
                if len(samples) < 8:
                    samples.append(repr(row[:8]))
            spec = str(row[(spec_col or 5) - 1] or "").strip() if spec_col else ""
            if spec:
                specs[spec] += 1
        swb.close()
        lines.append(f"summary rows title=8527: {n8527}, empty system: {empty_sys}")
        lines.append("systems: " + ", ".join(f"{k}={v}" for k, v in systems.most_common(15)))
        lines.append("specs: " + ", ".join(f"{k}={v}" for k, v in specs.most_common(8)))
        lines.extend(f"  empty-sys sample: {s}" for s in samples)

    _write(lines)
    print(f"wrote {OUT}", flush=True)
    return 0


def kind_label(kind: str) -> str:
    return kind


if __name__ == "__main__":
    raise SystemExit(main())
