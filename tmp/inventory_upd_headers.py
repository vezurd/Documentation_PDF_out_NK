"""Inventory UPD upload xlsx headers under the network folder. UTF-8 report only."""

from __future__ import annotations

import os
import sys
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "tmp" / "upd_headers_inventory.txt"

DEFAULT_UPD_ROOT = (
    r"\\bcc\root\CurProjects\di_manegers\Пятницкий_П"
    r"\Амурский ГХК\Поставки\Файл закачки УПД по всем ДС"
)

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".xlsb"}


def _cell_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.replace("\r\n", " ").replace("\n", " ").strip()
    return str(value).strip()


def _norm_header(cells: list[str]) -> tuple[str, ...]:
    # Drop trailing empty cells so padding differences do not split groups.
    while cells and not cells[-1]:
        cells = cells[:-1]
    return tuple(c.casefold() for c in cells)


def _looks_like_header(cells: list[str]) -> bool:
    filled = [c for c in cells if c]
    if len(filled) < 3:
        return False
    text = " ".join(filled).casefold()
    hints = (
        "наимен",
        "код",
        "кол",
        "номенкл",
        "уpd",
        "упд",
        "документ",
        "артикул",
        "ед.",
        "единиц",
        "сумма",
        "ндс",
        "bcc",
        "специф",
        "титул",
    )
    return any(h in text for h in hints) or len(filled) >= 6


def inspect_file(path: Path, root: Path) -> dict:
    rel = str(path.relative_to(root))
    suffix = path.suffix.lower()
    rec: dict = {
        "rel": rel,
        "suffix": suffix,
        "size": path.stat().st_size,
        "ok": False,
        "error": "",
        "sheets": [],
        "header_rows": [],
    }
    if suffix in {".xls", ".xlsb"}:
        rec["error"] = f"unsupported suffix {suffix}"
        return rec
    try:
        import openpyxl
    except Exception as exc:  # pragma: no cover
        rec["error"] = f"openpyxl import: {exc}"
        return rec
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        rec["error"] = f"open: {type(exc).__name__}: {exc}"
        return rec
    try:
        for ws in wb.worksheets:
            sheet_info = {
                "name": ws.title,
                "max_row": ws.max_row,
                "max_col": ws.max_column,
                "header_row": None,
                "headers": (),
                "preview": [],
            }
            preview_rows: list[list[str]] = []
            header_idx = None
            headers: tuple[str, ...] = ()
            for i, row in enumerate(ws.iter_rows(min_row=1, max_row=12, values_only=True), start=1):
                cells = [_cell_text(v) for v in row]
                preview_rows.append(cells)
                if header_idx is None and _looks_like_header(cells):
                    header_idx = i
                    headers = _norm_header(cells)
            sheet_info["preview"] = preview_rows
            sheet_info["header_row"] = header_idx
            sheet_info["headers"] = headers
            rec["sheets"].append(sheet_info)
            if headers:
                rec["header_rows"].append((ws.title, header_idx, headers))
        rec["ok"] = True
    except Exception as exc:
        rec["error"] = f"read: {type(exc).__name__}: {exc}"
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return rec


def main() -> int:
    root = Path(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_UPD_ROOT)
    lines: list[str] = []

    def out(msg: str = "") -> None:
        lines.append(msg)

    out(f"root: {root}")
    out(f"exists: {root.exists()}")
    out(f"is_dir: {root.is_dir()}")
    if not root.is_dir():
        OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
        print(f"wrote {OUT} (folder missing)", file=sys.stderr)
        return 1

    files: list[Path] = []
    other: list[Path] = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        for name in filenames:
            if name.startswith("~$"):
                continue
            p = Path(dirpath) / name
            if p.suffix.lower() in EXCEL_SUFFIXES:
                files.append(p)
            else:
                other.append(p)
    files.sort(key=lambda p: str(p).casefold())
    out(f"excel_files: {len(files)}")
    out(f"other_files: {len(other)}")
    suf = Counter(p.suffix.lower() for p in files)
    out(f"excel_suffixes: {dict(suf)}")
    folders = Counter(str(p.parent.relative_to(root)) or "." for p in files)
    out("folders:")
    for folder, n in sorted(folders.items(), key=lambda kv: (-kv[1], kv[0].casefold())):
        out(f"  {n:4d}  {folder}")
    if other:
        out("other_files_sample:")
        for p in other[:30]:
            out(f"  {p.relative_to(root)}  {p.suffix}")

    workers = 8 if len(files) >= 2 else 1
    results: list[dict] = []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futs = {pool.submit(inspect_file, p, root): p for p in files}
        done = 0
        for fut in as_completed(futs):
            results.append(fut.result())
            done += 1
            if done % 20 == 0 or done == len(files):
                print(f"inspected {done}/{len(files)}", file=sys.stderr)

    results.sort(key=lambda r: r["rel"].casefold())
    failed = [r for r in results if not r["ok"]]
    out("")
    out(f"read_ok: {len(results) - len(failed)}")
    out(f"read_fail: {len(failed)}")
    for r in failed:
        out(f"  FAIL {r['rel']}: {r['error']}")

    sheet_names = Counter()
    header_groups: dict[tuple[str, ...], list[str]] = defaultdict(list)
    files_multi_sheet: list[str] = []
    files_no_header: list[str] = []
    for r in results:
        if not r["ok"]:
            continue
        if len(r["sheets"]) != 1:
            files_multi_sheet.append(f"{r['rel']} sheets={[s['name'] for s in r['sheets']]}")
        for s in r["sheets"]:
            sheet_names[s["name"]] += 1
            if s["headers"]:
                header_groups[s["headers"]].append(f"{r['rel']} :: {s['name']} :: row {s['header_row']}")
            else:
                files_no_header.append(f"{r['rel']} :: {s['name']}")

    out("")
    out("sheet_names:")
    for name, n in sheet_names.most_common():
        out(f"  {n:4d}  {name}")

    out("")
    out(f"header_signatures: {len(header_groups)}")
    for i, (hdr, members) in enumerate(
        sorted(header_groups.items(), key=lambda kv: (-len(kv[1]), kv[0])),
        start=1,
    ):
        out("")
        out(f"=== signature {i}  files={len(members)}  cols={len(hdr)} ===")
        out("headers:")
        for idx, title in enumerate(hdr, start=1):
            out(f"  {idx:3d}  {title}")
        out("members:")
        for m in members[:40]:
            out(f"  {m}")
        if len(members) > 40:
            out(f"  ... +{len(members) - 40} more")

    out("")
    out(f"multi_sheet_files: {len(files_multi_sheet)}")
    for m in files_multi_sheet[:40]:
        out(f"  {m}")
    out("")
    out(f"no_header_sheets: {len(files_no_header)}")
    for m in files_no_header[:40]:
        out(f"  {m}")

    # Dump preview of first 3 files of the largest signature.
    if header_groups:
        top_hdr, top_members = max(header_groups.items(), key=lambda kv: len(kv[1]))
        sample_rel = top_members[0].split(" :: ", 1)[0]
        sample = next(r for r in results if r["rel"] == sample_rel)
        out("")
        out(f"sample_preview: {sample_rel}")
        for s in sample["sheets"]:
            out(f"  sheet={s['name']} max_row={s['max_row']} max_col={s['max_col']}")
            for i, row in enumerate(s["preview"], start=1):
                shown = " | ".join(c for c in row[:20] if True)
                out(f"    r{i}: {shown}")

    OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"wrote {OUT}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
