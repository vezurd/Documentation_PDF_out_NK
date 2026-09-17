"""Read-only occupancy scan of RFP_Зиновьев columns (analysis, no writes)."""
from __future__ import annotations

import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.rfp_parts.analyze_rfp_parts import (  # noqa: E402
    EXPECTED_PART_HEADER,
    _choose_part_sheet,
    _find_header_on_sheet,
    split_part_sheets,
)

PARTS_DIR = Path("//bcc/eng/PrDoc/377_НИПИГАЗ/АГХК/КСБ/RFP_MTO_VO/_RFP/RFP_Зиновьев")
OUT = Path(__file__).with_name("scan_rfp_unused_columns.txt")
MAX_HEADER_ROWS = 8
READ_COLS = 30  # A..AD


def _cell_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return text.replace("\n", " | ")


def _is_empty(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def main() -> None:
    files = sorted(
        p
        for p in PARTS_DIR.iterdir()
        if p.suffix.lower() in {".xlsx", ".xlsm"} and not p.name.startswith("~$")
    )
    lines: list[str] = []
    w = lines.append
    w(f"dir={PARTS_DIR}")
    w(f"files={len(files)}")
    w("")

    header_by_col: dict[int, Counter[str]] = defaultdict(Counter)
    group_by_col: dict[int, Counter[str]] = defaultdict(Counter)
    occupied_files: dict[int, int] = Counter()
    occupied_rows: dict[int, int] = Counter()
    sample_values: dict[int, list[str]] = defaultdict(list)
    file_header_row: list[tuple[str, int, str, dict[str, int]]] = []
    optional_hits: Counter[str] = Counter()
    errors: list[str] = []

    for path in files:
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
        except Exception as exc:
            errors.append(f"{path.name}: open failed {exc}")
            continue
        try:
            warnings: list[tuple[str, str, str]] = []
            _skipped, candidates = split_part_sheets(wb.sheetnames)
            if not candidates:
                errors.append(f"{path.name}: no working sheets sheets={wb.sheetnames}")
                continue
            probes = [
                (
                    name,
                    _find_header_on_sheet(
                        wb[name], warnings, path.name, prefer_lot_qty=True
                    ),
                )
                for name in candidates
            ]
            _chosen, header = _choose_part_sheet(candidates, probes)
            if header is None:
                errors.append(f"{path.name}: header not found sheets={wb.sheetnames}")
                continue
            ws = wb[header.sheet]
            file_header_row.append(
                (path.name, header.row_number, header.sheet, dict(header.columns))
            )
            for field in ("NAME_2", "VENDOR", "VALUES_2"):
                if field in header.columns:
                    optional_hits[f"{field}@{header.columns[field]}"] += 1

            # Group header = row above detected header, if any.
            group_row = None
            header_row_vals = None
            for i, row in enumerate(
                ws.iter_rows(min_row=1, max_row=header.row_number, values_only=True),
                start=1,
            ):
                cells = list(row)
                if i == header.row_number:
                    header_row_vals = cells
                elif i == header.row_number - 1:
                    group_row = cells
            if header_row_vals:
                for idx, val in enumerate(header_row_vals[:READ_COLS]):
                    text = _cell_text(val)
                    if text:
                        header_by_col[idx][text] += 1
            if group_row:
                for idx, val in enumerate(group_row[:READ_COLS]):
                    text = _cell_text(val)
                    if text:
                        group_by_col[idx][text] += 1

            file_occupied = set()
            for row in ws.iter_rows(
                min_row=header.row_number + 1, values_only=True
            ):
                cells = list(row)
                watched = cells[:8] + (
                    cells[16:18] if len(cells) > 17 else []
                )
                if not any(not _is_empty(v) for v in watched):
                    continue
                for idx in range(min(READ_COLS, len(cells))):
                    if _is_empty(cells[idx]):
                        continue
                    occupied_rows[idx] += 1
                    file_occupied.add(idx)
                    if len(sample_values[idx]) < 8:
                        sample = _cell_text(cells[idx])
                        if sample and sample not in sample_values[idx]:
                            sample_values[idx].append(sample[:120])
            for idx in file_occupied:
                occupied_files[idx] += 1
        finally:
            wb.close()

    w("=== header finder (parts-mode, lot qty rightmost) ===")
    sigs = Counter(
        ", ".join(f"{k}:{v}" for k, v in sorted(cols.items()))
        for _, _, _, cols in file_header_row
    )
    for sig, n in sigs.most_common():
        w(f"  {n:3d}  {sig}")
    w(f"optional field hits: {dict(optional_hits)}")
    w("")

    w("=== expected part header ===")
    for field, idx in EXPECTED_PART_HEADER.items():
        w(f"  {get_column_letter(idx + 1):>3} idx={idx:2d}  {field}")
    w("")

    w("=== group row (header-1) labels ===")
    for idx in range(READ_COLS):
        if idx not in group_by_col:
            continue
        labels = ", ".join(f"{lab!r}×{n}" for lab, n in group_by_col[idx].most_common(4))
        w(f"  {get_column_letter(idx + 1):>3} ({idx:2d})  {labels}")
    w("")

    w("=== header-row labels ===")
    for idx in range(READ_COLS):
        if idx not in header_by_col:
            w(f"  {get_column_letter(idx + 1):>3} ({idx:2d})  <no header text in any file>")
            continue
        labels = ", ".join(f"{lab!r}×{n}" for lab, n in header_by_col[idx].most_common(4))
        w(f"  {get_column_letter(idx + 1):>3} ({idx:2d})  {labels}")
    w("")

    n_files = len(file_header_row) or 1
    w("=== occupancy on data rows (non-empty cells) ===")
    w(
        f"{'col':>4} {'idx':>3} {'files':>6} {'%files':>7} {'rows':>8}  header  samples"
    )
    for idx in range(READ_COLS):
        files_n = occupied_files.get(idx, 0)
        rows_n = occupied_rows.get(idx, 0)
        hdr = ""
        if idx in header_by_col:
            hdr = header_by_col[idx].most_common(1)[0][0][:40]
        elif idx in group_by_col:
            hdr = "[group] " + group_by_col[idx].most_common(1)[0][0][:32]
        samples = " | ".join(sample_values.get(idx, [])[:3])
        w(
            f"{get_column_letter(idx + 1):>4} {idx:3d} {files_n:6d} "
            f"{100 * files_n / n_files:6.1f}% {rows_n:8d}  {hdr!r}  {samples}"
        )
    w("")

    w("=== columns with ZERO data in all scanned files ===")
    empty_cols = [idx for idx in range(READ_COLS) if occupied_files.get(idx, 0) == 0]
    if empty_cols:
        for idx in empty_cols:
            hdr = header_by_col[idx].most_common(1)[0][0] if idx in header_by_col else ""
            grp = group_by_col[idx].most_common(1)[0][0] if idx in group_by_col else ""
            w(f"  {get_column_letter(idx + 1)} idx={idx} header={hdr!r} group={grp!r}")
    else:
        w("  none in A.. range")
    w("")

    w("=== columns unused by parts extractor (not in EXPECTED_PART_HEADER) ===")
    used_idx = set(EXPECTED_PART_HEADER.values())
    for idx in range(READ_COLS):
        if idx in used_idx:
            continue
        files_n = occupied_files.get(idx, 0)
        hdr = header_by_col[idx].most_common(1)[0][0] if idx in header_by_col else ""
        grp = group_by_col[idx].most_common(1)[0][0] if idx in group_by_col else ""
        w(
            f"  {get_column_letter(idx + 1)} idx={idx} files={files_n}/{n_files} "
            f"rows={occupied_rows.get(idx, 0)} header={hdr!r} group={grp!r}"
        )
    w("")

    if errors:
        w("=== errors ===")
        for item in errors:
            w(f"  {item}")

    text = "\n".join(lines) + "\n"
    OUT.write_text(text, encoding="utf-8")
    sys.stdout.buffer.write(text.encode("utf-8", errors="backslashreplace"))
    print(f"\nwrote {OUT}", file=sys.stderr)


if __name__ == "__main__":
    main()
