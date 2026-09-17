"""Compare registry xlsx (col B names) vs actual TSD packing-list files.

Registry names may be a substring of the UL filename.
Read-only. Writes UTF-8 report under tmp/.
"""
from __future__ import annotations

import os
import sys
from collections import defaultdict
from pathlib import Path

os.environ.setdefault("PYTHONUTF8", "1")

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

REGISTRY = Path(
    r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP"
    r"\реестра УЛ с самими УЛ.xlsx"
)
TSD_ROOT = Path(
    r"\\bcc\root\CurProjects\di_manegers\Пятницкий_П"
    r"\Амурский ГХК\Поставки\ТСД по всем ДС"
)
OUT_MD = ROOT / "tmp" / "ul_registry_vs_files.md"
OUT_TXT = ROOT / "tmp" / "ul_registry_vs_files.txt"


def _safe(s: object) -> str:
    return str(s) if s is not None else ""


def peek_registry() -> tuple[list[str], list[tuple[int, str, list[str]]]]:
    """Return sheet names and first rows of each sheet (row idx 1-based)."""
    import openpyxl

    wb = openpyxl.load_workbook(REGISTRY, read_only=True, data_only=True)
    sheets = list(wb.sheetnames)
    samples: list[tuple[int, str, list[str]]] = []
    for name in sheets:
        ws = wb[name]
        for i, row in enumerate(ws.iter_rows(max_row=8, values_only=True), start=1):
            cells = [_safe(c) for c in (row or ())]
            samples.append((i, name, cells))
    wb.close()
    return sheets, samples


def load_registry_col_b() -> tuple[str, list[tuple[int, str]]]:
    """Unique-preserving list of (excel_row, col_B) from first sheet."""
    import openpyxl

    wb = openpyxl.load_workbook(REGISTRY, read_only=True, data_only=True)
    sheet = wb.sheetnames[0]
    ws = wb[sheet]
    items: list[tuple[int, str]] = []
    for i, row in enumerate(ws.iter_rows(min_col=2, max_col=2, values_only=True), start=1):
        val = row[0] if row else None
        if val is None:
            continue
        text = str(val).strip()
        if not text:
            continue
        items.append((i, text))
    wb.close()
    return sheet, items


def collect_ul_files() -> list[Path]:
    from RFQ.ds_compare.tsd_packing_load import collect_tsd_files

    docs = collect_tsd_files(str(TSD_ROOT))
    return [Path(d.file_full_path) for d in docs]


def match_name(reg_name: str, filename: str) -> bool:
    """Registry name is a substring of the filename (case-insensitive)."""
    a = reg_name.casefold().strip()
    b = filename.casefold()
    if not a:
        return False
    return a in b


def main() -> int:
    lines: list[str] = []

    def log(msg: str) -> None:
        lines.append(msg)
        enc = getattr(sys.stdout, "encoding", None) or "utf-8"
        sys.stdout.buffer.write((msg + "\n").encode(enc, errors="backslashreplace"))

    log("=== UL registry vs actual files ===")
    log(f"registry exists: {REGISTRY.exists()}  {ascii(str(REGISTRY))}")
    log(f"tsd root exists: {TSD_ROOT.exists()}  {ascii(str(TSD_ROOT))}")

    if not REGISTRY.exists():
        log("STOP: registry xlsx not reachable")
        OUT_TXT.write_text("\n".join(lines), encoding="utf-8")
        return 1

    sheets, samples = peek_registry()
    log(f"sheets: {sheets}")
    for i, name, cells in samples:
        preview = " | ".join(cells[:8])
        log(f"  [{name}] r{i}: {preview}")

    sheet, items = load_registry_col_b()
    # drop header-like first row if col B looks like a header
    headerish = {"имя", "ул", "наименование", "packing", "файл", "name", "ul"}
    if items:
        first = items[0][1].casefold()
        if any(h in first for h in headerish) and len(first) < 40:
            log(f"skip header row {items[0][0]}: {items[0][1]!r}")
            items = items[1:]

    unique_names: list[str] = []
    seen: set[str] = set()
    for _, name in items:
        key = name.casefold()
        if key not in seen:
            seen.add(key)
            unique_names.append(name)

    log(f"sheet={sheet!r} col B rows={len(items)} unique={len(unique_names)}")

    if not TSD_ROOT.exists():
        log("STOP: TSD root not reachable")
        OUT_TXT.write_text("\n".join(lines), encoding="utf-8")
        return 1

    files = collect_ul_files()
    log(f"actual xlsx under TSD: {len(files)}")

    file_by_name = [(p.name, p) for p in files]

    matched: dict[str, list[Path]] = {}
    unmatched_reg: list[str] = []
    used_files: set[str] = set()

    for name in unique_names:
        hits = [p for fname, p in file_by_name if match_name(name, fname)]
        if hits:
            matched[name] = hits
            for p in hits:
                used_files.add(str(p).casefold())
        else:
            unmatched_reg.append(name)

    extra_files = [p for p in files if str(p).casefold() not in used_files]

    # ambiguous: one registry name hits many files, or many names hit one file
    multi = {k: v for k, v in matched.items() if len(v) > 1}
    file_to_regs: dict[str, list[str]] = defaultdict(list)
    for name, paths in matched.items():
        for p in paths:
            file_to_regs[p.name].append(name)
    many_regs = {fn: regs for fn, regs in file_to_regs.items() if len(regs) > 1}

    log("")
    log(f"MATCHED unique registry names: {len(matched)}")
    log(f"ONLY IN REGISTRY (no file contains the name): {len(unmatched_reg)}")
    log(f"ONLY IN FILES (no registry name is a substring of filename): {len(extra_files)}")
    log(f"AMBIGUOUS registry→many files: {len(multi)}")
    log(f"AMBIGUOUS file←many registry names: {len(many_regs)}")

    md: list[str] = []
    md.append("# Сверка реестра УЛ (колонка B) с фактическими файлами")
    md.append("")
    md.append(f"- Реестр: `{REGISTRY}`")
    md.append(f"- Папка ТСД (GUI «ДС · Упаковочные листы»): `{TSD_ROOT}`")
    md.append(f"- Лист: `{sheet}`")
    md.append(f"- Строк колонки B (без пустых/шапки): **{len(items)}**, уникальных имён: **{len(unique_names)}**")
    md.append(f"- Фактических xlsx УЛ: **{len(files)}**")
    md.append("- Правило: имя из реестра — **подстрока** имени файла (без учёта регистра).")
    md.append("")
    md.append("| Статус | Кол-во |")
    md.append("|---|---|")
    md.append(f"| Совпадение | {len(matched)} |")
    md.append(f"| Только в реестре | {len(unmatched_reg)} |")
    md.append(f"| Только среди файлов | {len(extra_files)} |")
    md.append(f"| Реестр → несколько файлов | {len(multi)} |")
    md.append(f"| Файл ← несколько имён реестра | {len(many_regs)} |")
    md.append("")

    if unmatched_reg:
        md.append("## Только в реестре")
        md.append("")
        for name in unmatched_reg:
            md.append(f"- `{name}`")
        md.append("")

    if extra_files:
        md.append("## Только среди файлов (имя реестра не найдено в имени файла)")
        md.append("")
        for p in extra_files:
            try:
                rel = p.relative_to(TSD_ROOT)
            except Exception:
                rel = p
            md.append(f"- `{rel}`")
        md.append("")

    if multi:
        md.append("## Реестр → несколько файлов")
        md.append("")
        for name, paths in multi.items():
            md.append(f"- `{name}` → {len(paths)} файлов")
            for p in paths[:12]:
                md.append(f"  - `{p.name}`")
            if len(paths) > 12:
                md.append(f"  - … ещё {len(paths) - 12}")
        md.append("")

    if many_regs:
        md.append("## Файл ← несколько имён реестра")
        md.append("")
        for fn, regs in sorted(many_regs.items()):
            md.append(f"- `{fn}` ← {regs}")
        md.append("")

    md.append("## Совпадения (имя реестра → файл)")
    md.append("")
    for name in unique_names:
        paths = matched.get(name)
        if not paths:
            continue
        files_s = ", ".join(f"`{p.name}`" for p in paths[:3])
        extra = f" (+{len(paths) - 3})" if len(paths) > 3 else ""
        md.append(f"- `{name}` → {files_s}{extra}")

    OUT_TXT.write_text("\n".join(lines), encoding="utf-8")
    OUT_MD.write_text("\n".join(md), encoding="utf-8")
    log(f"wrote {OUT_MD}")
    log(f"wrote {OUT_TXT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
