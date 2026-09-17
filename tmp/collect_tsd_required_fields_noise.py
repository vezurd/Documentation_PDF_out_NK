"""Collect unique cell values from ``source_required_fields`` critical remarks.

Purpose
-------
Separate packing-list **document banner / footer / legend** noise from real
incomplete position rows — without guessing filters that could drop data.

Strategy (safe-by-default)
--------------------------
1. Parse ``tsd_packing_critical.txt`` (or rebuild issues via loader).
2. Re-open each cited xlsx once; dump mapped CODE/NAME/VALUES/UNITS/… cells.
3. Build frequency tables of unique cell texts for flagged rows.
4. Cross-check against **confirmed** ``position_row`` values from the packing
   cache: anything that also appears on real positions is marked KEEP
   (must not become a filter).
5. Emit candidate header/footer labels = frequent in flagged rows AND absent
   from position cells, plus pattern hints (bilingual ``/``, ``No.``, address
   prose, package-type legend, etc.).

Run (from repo root, needs UNC access)::

    python tmp/collect_tsd_required_fields_noise.py
    python tmp/collect_tsd_required_fields_noise.py --max-files 50   # smoke
    python tmp/collect_tsd_required_fields_noise.py --reload-issues  # slower

Outputs (UTF-8) under ``tmp/tsd_required_fields_noise/``:
- ``summary.txt`` — counts, zone split, top missing-field combos
- ``unique_by_field.tsv`` — value \\t field \\t flagged_count \\t in_positions \\t keep
- ``candidates_filter.txt`` — suggested exact/startswith labels (review!)
- ``samples_by_zone.txt`` — example rows for banner / footer / mid-table
- ``issues_index.tsv`` — every cited location with cell snapshot
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import openpyxl

from RFQ.ds_compare.ds_compare_config import (
    load_ds_compare_config,
    normalize_gui_paths,
)
from RFQ.ds_compare.tsd_packing_load import (
    TSD_PACKING_CRITICAL_REPORT_NAME,
    is_sopl_sheet,
    load_tsd_file_rows,
    tsd_packing_cache_dir,
)
from RFQ.packing_list_provider import load_packing_dataset
from base.tables_columns import (
    CODE,
    NAME,
    SPECIFICATION_NAME,
    TAGS,
    UNITS,
    VALUES,
    VENDOR,
)

OUT_DIR = ROOT / "tmp" / "tsd_required_fields_noise"

_ISSUE_RE = re.compile(
    r"ERROR source_required_fields \[(?P<path>.+?) · (?P<sheet>.+?) · "
    r"строка (?P<row>\d+)\] field=(?P<field>[^:]+):",
    re.IGNORECASE,
)

# SO - PL mapped columns (0-based), STD Single* mapped columns.
_SOPL_COLS = {
    SPECIFICATION_NAME: 2,
    VENDOR: 5,
    CODE: 6,
    TAGS: 9,
    NAME: 12,
    VALUES: 13,
    UNITS: 14,
}
_STD_COLS = {
    CODE: 1,
    SPECIFICATION_NAME: 2,
    TAGS: 3,
    NAME: 7,
    VALUES: 8,
    UNITS: 9,
    VENDOR: 12,
}

_FIELDS = (CODE, NAME, VALUES, UNITS, SPECIFICATION_NAME, VENDOR, TAGS)

# Soft hints only for reporting — not applied as filters by this script.
_LABEL_HINT_RE = re.compile(
    r"(?i)"
    r"(packing list|date\s*/|consignee|consignor|seller|buyer|marks\s*/|"
    r"for delivery|final destination|partial shipment|shipping invoice|"
    r"sales order|contractor po|basis of sales|type of package|stackability|"
    r"adress:|address:|адрес:|поставщик|получатель|продавец|покупатель|"
    r"упаковочн|дата\s*:|№\s*:|no\.\s*/|name of (zip|document)|po item|"
    r"russian translation|quantity|vendor\b|units\b)"
)


@dataclass(frozen=True)
class IssueLoc:
    source_rel: str
    sheet: str
    excel_row: int
    field: str


def _norm_cell(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        value = int(value)
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", " / ", text)
    return text.strip()


def _parse_critical(path: Path) -> tuple[str, list[IssueLoc]]:
    text = path.read_text(encoding="utf-8")
    root = ""
    for line in text.splitlines()[:5]:
        if line.lower().startswith("root:"):
            root = line.split(":", 1)[1].strip()
            break
    locs: list[IssueLoc] = []
    for match in _ISSUE_RE.finditer(text):
        locs.append(
            IssueLoc(
                source_rel=match.group("path").strip(),
                sheet=match.group("sheet").strip(),
                excel_row=int(match.group("row")),
                field=match.group("field").strip(),
            )
        )
    return root, locs


def _resolve_root(cli_root: str | None, critical_root: str) -> Path:
    if cli_root:
        return Path(cli_root)
    if critical_root:
        return Path(critical_root)
    gui = normalize_gui_paths(load_ds_compare_config().get("gui_paths"))
    folder = gui.get("last_tsd_packing_folder") or ""
    if not folder:
        raise SystemExit("No TSD root: pass --root or set gui_paths.last_tsd_packing_folder")
    return Path(folder)


def _position_value_sets() -> dict[str, set[str]]:
    """Unique normalized cell texts from confirmed packing positions (cache)."""
    dataset = load_packing_dataset()
    buckets: dict[str, set[str]] = {key: set() for key in _FIELDS}
    if not dataset.available:
        print(f"WARN packing cache unavailable: {[i.code for i in dataset.issues]}")
        return buckets
    for row in dataset.rows:
        for key in _FIELDS:
            el = row.el.get(key)
            text = _norm_cell(None if el is None else el.value)
            if text:
                buckets[key].add(text)
                buckets[key].add(text.upper())
    print(f"position cache rows={len(dataset.rows)} unique CODE={len(buckets[CODE])}")
    return buckets


def _zone(excel_row: int, first_pos: int | None, last_pos: int | None) -> str:
    if first_pos is None or last_pos is None:
        if excel_row <= 16:
            return "banner_or_unknown"
        return "footer_or_unknown"
    if excel_row < first_pos:
        return "banner_before_table"
    if excel_row > last_pos:
        return "footer_after_table"
    return "mid_table"


def _read_row_snapshot(
    ws: object,
    excel_row: int,
    colmap: dict[str, int],
) -> dict[str, str]:
    # openpyxl 1-based columns
    max_col = max(colmap.values()) + 1
    cells = next(
        ws.iter_rows(
            min_row=excel_row,
            max_row=excel_row,
            max_col=max_col,
            values_only=True,
        ),
        (),
    )
    snap: dict[str, str] = {}
    for key, idx0 in colmap.items():
        value = cells[idx0] if idx0 < len(cells) else None
        snap[key] = _norm_cell(value)
    return snap


def _collect_from_files(
    root: Path,
    locs: list[IssueLoc],
    *,
    max_files: int | None,
) -> tuple[
    list[dict[str, object]],
    Counter[str],
    dict[str, Counter[str]],
    Counter[str],
]:
    by_file: dict[str, list[IssueLoc]] = defaultdict(list)
    for loc in locs:
        by_file[loc.source_rel].append(loc)

    file_list = sorted(by_file.keys(), key=str.lower)
    if max_files is not None:
        file_list = file_list[:max_files]

    rows_out: list[dict[str, object]] = []
    zone_counter: Counter[str] = Counter()
    field_value_counter: dict[str, Counter[str]] = {key: Counter() for key in _FIELDS}
    missing_field_counter: Counter[str] = Counter()

    for i, rel in enumerate(file_list, start=1):
        path = root / rel
        file_locs = by_file[rel]
        print(f"[{i}/{len(file_list)}] {rel} ({len(file_locs)} issues)")
        if not path.is_file():
            for loc in file_locs:
                rows_out.append(
                    {
                        "source_rel": rel,
                        "sheet": loc.sheet,
                        "excel_row": loc.excel_row,
                        "missing_field": loc.field,
                        "zone": "file_missing",
                        "error": "file not found",
                    }
                )
            continue

        # Position band from loader (same classification as production).
        try:
            positions, _plan, _all_rows, _hits, _issues = load_tsd_file_rows(
                str(path), root=str(root)
            )
            pos_rows = [
                int(getattr(r, "_packing_source_row", 0) or 0) for r in positions
            ]
            first_pos = min(pos_rows) if pos_rows else None
            last_pos = max(pos_rows) if pos_rows else None
        except Exception as exc:
            first_pos = last_pos = None
            print(f"  WARN load_tsd_file_rows: {type(exc).__name__}: {exc}")

        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            sheet_groups: dict[str, list[IssueLoc]] = defaultdict(list)
            for loc in file_locs:
                sheet_groups[loc.sheet].append(loc)
            for sheet_name, sheet_locs in sheet_groups.items():
                if sheet_name not in wb.sheetnames:
                    for loc in sheet_locs:
                        rows_out.append(
                            {
                                "source_rel": rel,
                                "sheet": sheet_name,
                                "excel_row": loc.excel_row,
                                "missing_field": loc.field,
                                "zone": "sheet_missing",
                                "error": "sheet not found",
                            }
                        )
                    continue
                ws = wb[sheet_name]
                colmap = _SOPL_COLS if is_sopl_sheet(sheet_name) else _STD_COLS
                # Unique excel rows (several issues shouldn't re-read).
                seen_rows: set[int] = set()
                for loc in sheet_locs:
                    missing_field_counter[loc.field] += 1
                    if loc.excel_row in seen_rows:
                        continue
                    seen_rows.add(loc.excel_row)
                    snap = _read_row_snapshot(ws, loc.excel_row, colmap)
                    zone = _zone(loc.excel_row, first_pos, last_pos)
                    zone_counter[zone] += 1
                    label_hit = any(
                        _LABEL_HINT_RE.search(snap.get(k, "") or "") for k in _FIELDS
                    )
                    rec: dict[str, object] = {
                        "source_rel": rel,
                        "sheet": sheet_name,
                        "excel_row": loc.excel_row,
                        "missing_field": loc.field,
                        "zone": zone,
                        "first_pos": first_pos,
                        "last_pos": last_pos,
                        "label_hint": label_hit,
                        "sopl": is_sopl_sheet(sheet_name),
                    }
                    for key in _FIELDS:
                        text = snap.get(key, "")
                        rec[key] = text
                        if text:
                            field_value_counter[key][text] += 1
                    rows_out.append(rec)
        finally:
            wb.close()

    return rows_out, zone_counter, field_value_counter, missing_field_counter


def _write_outputs(
    *,
    root: Path,
    locs: list[IssueLoc],
    rows: list[dict[str, object]],
    zone_counter: Counter[str],
    field_value_counter: dict[str, Counter[str]],
    missing_field_counter: Counter[str],
    position_sets: dict[str, set[str]],
) -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    # --- summary ---
    sopl_n = sum(1 for r in rows if r.get("sopl"))
    label_n = sum(1 for r in rows if r.get("label_hint"))
    summary_lines = [
        f"root: {root}",
        f"critical issues (source_required_fields): {len(locs)}",
        f"unique files: {len({loc.source_rel for loc in locs})}",
        f"snapshotted unique rows: {len(rows)}",
        f"sopl rows: {sopl_n}",
        f"label_hint rows: {label_n}",
        "",
        "=== zones ===",
    ]
    for zone, count in zone_counter.most_common():
        summary_lines.append(f"{zone}\t{count}")
    summary_lines.append("")
    summary_lines.append("=== missing field combos ===")
    for field, count in missing_field_counter.most_common():
        summary_lines.append(f"{field}\t{count}")
    (OUT_DIR / "summary.txt").write_text(
        "\n".join(summary_lines) + "\n", encoding="utf-8"
    )

    # --- unique_by_field.tsv ---
    tsv_lines = [
        "field\tflagged_count\tin_positions\tkeep\tlabel_hint\tvalue"
    ]
    candidates: list[tuple[str, str, int]] = []
    for field, counter in field_value_counter.items():
        pos_set = position_sets.get(field, set())
        for value, count in counter.most_common():
            in_pos = value in pos_set or value.upper() in pos_set
            keep = "KEEP" if in_pos else "candidate"
            hint = "1" if _LABEL_HINT_RE.search(value) else "0"
            # Escape tabs/newlines already normalized.
            tsv_lines.append(
                f"{field}\t{count}\t{int(in_pos)}\t{keep}\t{hint}\t{value}"
            )
            if not in_pos and (count >= 3 or hint == "1"):
                candidates.append((field, value, count))
    (OUT_DIR / "unique_by_field.tsv").write_text(
        "\n".join(tsv_lines) + "\n", encoding="utf-8"
    )

    # --- candidates_filter.txt (human review) ---
    cand_lines = [
        "REVIEW ONLY — do not paste blindly into get_row_type.",
        "Rule of thumb: prefer exact match / startswith on CODE/VALUES/UNITS",
        "labels; never filter NAME/CODE values marked KEEP in unique_by_field.tsv.",
        "Outside table band: also consider requiring product-like CODE (BCC*/digits)",
        "instead of expanding label lists forever.",
        "",
    ]
    by_field: dict[str, list[tuple[str, int]]] = defaultdict(list)
    for field, value, count in candidates:
        by_field[field].append((value, count))
    for field in _FIELDS:
        items = sorted(by_field.get(field, []), key=lambda x: (-x[1], x[0].lower()))
        if not items:
            continue
        cand_lines.append(f"=== {field} ({len(items)} candidates) ===")
        for value, count in items[:200]:
            cand_lines.append(f"{count:5d}  {value}")
        cand_lines.append("")
    (OUT_DIR / "candidates_filter.txt").write_text(
        "\n".join(cand_lines) + "\n", encoding="utf-8"
    )

    # --- samples_by_zone.txt ---
    sample_lines: list[str] = []
    by_zone: dict[str, list[dict[str, object]]] = defaultdict(list)
    for rec in rows:
        by_zone[str(rec.get("zone"))].append(rec)
    for zone, items in sorted(by_zone.items()):
        sample_lines.append(f"=== {zone} (n={len(items)}) ===")
        for rec in items[:25]:
            sample_lines.append(
                f"{rec.get('source_rel')} | {rec.get('sheet')} | "
                f"row={rec.get('excel_row')} missing={rec.get('missing_field')} "
                f"hint={rec.get('label_hint')}"
            )
            sample_lines.append(
                f"  CODE={rec.get(CODE)!r}\n"
                f"  NAME={rec.get(NAME)!r}\n"
                f"  VALUES={rec.get(VALUES)!r}\n"
                f"  UNITS={rec.get(UNITS)!r}"
            )
        sample_lines.append("")
    (OUT_DIR / "samples_by_zone.txt").write_text(
        "\n".join(sample_lines) + "\n", encoding="utf-8"
    )

    # --- issues_index.tsv ---
    idx_lines = [
        "source_rel\tsheet\texcel_row\tzone\tmissing_field\tlabel_hint\t"
        "CODE\tNAME\tVALUES\tUNITS\tSPEC\tVENDOR"
    ]
    for rec in rows:
        idx_lines.append(
            "\t".join(
                [
                    str(rec.get("source_rel", "")),
                    str(rec.get("sheet", "")),
                    str(rec.get("excel_row", "")),
                    str(rec.get("zone", "")),
                    str(rec.get("missing_field", "")),
                    str(int(bool(rec.get("label_hint")))),
                    str(rec.get(CODE, "")).replace("\t", " "),
                    str(rec.get(NAME, "")).replace("\t", " "),
                    str(rec.get(VALUES, "")).replace("\t", " "),
                    str(rec.get(UNITS, "")).replace("\t", " "),
                    str(rec.get(SPECIFICATION_NAME, "")).replace("\t", " "),
                    str(rec.get(VENDOR, "")).replace("\t", " "),
                ]
            )
        )
    (OUT_DIR / "issues_index.tsv").write_text(
        "\n".join(idx_lines) + "\n", encoding="utf-8"
    )

    print(f"wrote {OUT_DIR}")


def _rebuild_locs_via_loader(root: Path, *, max_files: int | None) -> list[IssueLoc]:
    """Slow path: rescan TSD tree and collect source_required_fields issues."""
    from RFQ.ds_compare.tsd_packing_load import collect_tsd_files

    docs = collect_tsd_files(str(root))
    if max_files is not None:
        docs = docs[:max_files]
    locs: list[IssueLoc] = []
    for i, doc in enumerate(docs, start=1):
        path = doc.file_full_path
        print(f"reload [{i}/{len(docs)}] {path}")
        try:
            _pos, _plan, _rows, _hits, issues = load_tsd_file_rows(path, root=str(root))
        except Exception as exc:
            print(f"  skip: {exc}")
            continue
        for issue in issues:
            if issue.code != "source_required_fields":
                continue
            if issue.excel_row is None:
                continue
            locs.append(
                IssueLoc(
                    source_rel=issue.source_file,
                    sheet=issue.sheet,
                    excel_row=issue.excel_row,
                    field=issue.field or "",
                )
            )
    return locs


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", default=None, help="TSD source root")
    parser.add_argument(
        "--critical",
        default=None,
        help="Path to tsd_packing_critical.txt (default: cache dir)",
    )
    parser.add_argument(
        "--max-files",
        type=int,
        default=None,
        help="Limit unique files (smoke / partial run)",
    )
    parser.add_argument(
        "--reload-issues",
        action="store_true",
        help="Ignore critical.txt and rescan loader issues (slow)",
    )
    args = parser.parse_args()

    critical_path = (
        Path(args.critical)
        if args.critical
        else tsd_packing_cache_dir() / TSD_PACKING_CRITICAL_REPORT_NAME
    )
    critical_root = ""
    if args.reload_issues:
        root = _resolve_root(args.root, "")
        locs = _rebuild_locs_via_loader(root, max_files=args.max_files)
    else:
        if not critical_path.is_file():
            raise SystemExit(f"critical report not found: {critical_path}")
        critical_root, locs = _parse_critical(critical_path)
        root = _resolve_root(args.root, critical_root)
        print(f"parsed {len(locs)} source_required_fields from {critical_path}")

    if not locs:
        raise SystemExit("No source_required_fields locations found")

    position_sets = _position_value_sets()
    rows, zone_counter, field_value_counter, missing_field_counter = _collect_from_files(
        root, locs, max_files=args.max_files
    )
    _write_outputs(
        root=root,
        locs=locs,
        rows=rows,
        zone_counter=zone_counter,
        field_value_counter=field_value_counter,
        missing_field_counter=missing_field_counter,
        position_sets=position_sets,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
