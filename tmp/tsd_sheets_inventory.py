"""Inventory Excel sheet names under TSD packing-list root (read-only).

Walks all ``*.xlsx`` recursively, records sheet names per file, and checks that
each workbook has exactly ``Single 1`` and ``Master 1``.

Run from repo root::

    python tmp/tsd_sheets_inventory.py
    python tmp/tsd_sheets_inventory.py "D:\\local\\copy\\ТСД по всем ДС"

Exit code: 0 if every file matches the expected pair; 1 on anomalies or errors.
"""
from __future__ import annotations

import argparse
import os
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import openpyxl

import utils.path

DEFAULT_ROOT = (
    r"\\bcc\root\CurProjects\di_manegers\Пятницкий_П"
    r"\Амурский ГХК\Поставки\ТСД по всем ДС"
)
EXPECTED_SHEETS = frozenset({"Single 1", "Master 1"})
REPORT_PATH = ROOT / "tmp" / "tsd_sheets_inventory_report.txt"


def _safe_print(s: str) -> None:
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    sys.stdout.buffer.write((s + "\n").encode(enc, errors="backslashreplace"))


def _rel_path(root: Path, file_path: str) -> str:
    try:
        return str(Path(file_path).resolve().relative_to(root.resolve()))
    except Exception:
        return file_path


def _sheet_names(file_path: str) -> list[str]:
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=False)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def _is_ok(names: list[str]) -> bool:
    return len(names) == 2 and frozenset(names) == EXPECTED_SHEETS


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="TSD xlsx sheet inventory")
    parser.add_argument(
        "root",
        nargs="?",
        default=DEFAULT_ROOT,
        help="Root folder with packing-list xlsx (recursive)",
    )
    args = parser.parse_args(argv)
    root = Path(args.root)
    os.chdir(ROOT)

    lines: list[str] = []
    lines.append(f"root: {root}")
    lines.append(f"expected sheets: {sorted(EXPECTED_SHEETS)}")
    lines.append("")

    if not root.exists():
        msg = f"ERROR: root does not exist or is inaccessible: {root}"
        _safe_print(msg)
        lines.append(msg)
        REPORT_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")
        return 1

    docs = utils.path.get_files_single(
        str(root),
        endswith=(".xlsx", ".XLSX"),
        sub_folders=True,
        factory=lambda p: type(
            "Doc",
            (),
            {
                "file_full_path": p,
                "file_name": os.path.basename(p),
                "doc_Number_for_sort": os.path.basename(p),
            },
        )(),
    )

    ok_count = 0
    anomaly_rows: list[tuple[str, list[str], str]] = []
    set_counter: Counter[str] = Counter()
    error_rows: list[tuple[str, str]] = []

    for doc in docs:
        fp = doc.file_full_path
        rel = _rel_path(root, fp)
        try:
            names = _sheet_names(fp)
        except Exception as exc:
            error_rows.append((rel, repr(exc)))
            set_counter["<read_error>"] += 1
            continue
        key = " | ".join(names) if names else "<empty>"
        set_counter[key] += 1
        if _is_ok(names):
            ok_count += 1
        else:
            anomaly_rows.append((rel, names, key))

    total = len(docs)
    anomaly_count = len(anomaly_rows)
    error_count = len(error_rows)

    summary = [
        f"total_xlsx: {total}",
        f"ok (exactly Single 1 + Master 1): {ok_count}",
        f"anomaly: {anomaly_count}",
        f"read_errors: {error_count}",
        "",
        "sheet-set frequencies:",
    ]
    for key, cnt in set_counter.most_common():
        summary.append(f"  {cnt:5d}  {key}")

    lines.extend(summary)
    lines.append("")
    lines.append("ANOMALIES:")
    if not anomaly_rows:
        lines.append("  (none)")
    else:
        for rel, names, _key in anomaly_rows:
            lines.append(f"  {rel}")
            lines.append(f"    sheets ({len(names)}): {names}")

    lines.append("")
    lines.append("READ ERRORS:")
    if not error_rows:
        lines.append("  (none)")
    else:
        for rel, err in error_rows:
            lines.append(f"  {rel}")
            lines.append(f"    {err}")

    report_text = "\n".join(lines) + "\n"
    REPORT_PATH.write_text(report_text, encoding="utf-8")

    for line in summary:
        _safe_print(line)
    _safe_print(f"report: {REPORT_PATH}")
    if anomaly_count or error_count:
        _safe_print(
            f"STOP: {anomaly_count} anomal(y/ies), {error_count} read error(s). "
            "Do not assume Single 1 only until decided."
        )
        return 1
    _safe_print("OK: all workbooks have exactly Single 1 + Master 1.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
