"""Read-only layout-report measurement on ``tmp/rd_catalog_copy.sqlite3``.

Does not touch the live catalog database or UNC. Writes the xlsx into ``tmp/``.
"""

from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog.config import load_config
from rd_catalog.db import CatalogDatabase
from rd_catalog.kits import kit_identity_key
from rd_catalog.layout_report import (
    build_layout_report,
    format_layout_report_text,
    write_layout_report_xlsx,
)
from rd_catalog.models import SourceKind
from rd_catalog.parse import layout_reason_label

COPY_DB = ROOT / "tmp" / "rd_catalog_copy.sqlite3"
XLSX_OUT = ROOT / "tmp" / "rd_catalog_layout_report.xlsx"


def main() -> None:
    """Print taxonomy, grouped rows, fully-lost kits, and top title-marks."""

    config = load_config()
    database = CatalogDatabase(COPY_DB)
    records = database.list_files(source=SourceKind.RD, present_only=True)
    report = build_layout_report(records=records, rd_root=config.rd_root)
    write_layout_report_xlsx(report, XLSX_OUT)

    print(f"rd_root={config.rd_root}")
    print(f"present RD files={len(records)}")
    print(f"total_files={report.total_files}")
    print(f"total_mto_files={report.total_mto_files}")
    print(f"grouped_rows={len(report.rows)}")
    print(f"total_kits={report.total_kits}")
    print(f"fully_lost_kits={len(report.fully_lost_kits)}")
    other_count = next(
        (count for token, count, _mto in report.by_reason if token == "other"),
        0,
    )
    other_share = (
        100.0 * other_count / report.total_files if report.total_files else 0.0
    )
    print(
        f"residual other={other_count} "
        f"({other_share:.1f}% of {report.total_files})"
    )
    print()
    print("taxonomy (files / MTO):")
    for token, count, mto_count in report.by_reason:
        share = 100.0 * count / report.total_files if report.total_files else 0.0
        print(
            f"  {count:5d}  MTO={mto_count:4d}  {share:5.1f}%  {token}  "
            f"{layout_reason_label(token)}"
        )
    print()
    print("fully_lost_kits:")
    if not report.fully_lost_kits:
        print("  (none)")
    else:
        lost_counts: Counter[tuple[str, str]] = Counter()
        lost_mto: Counter[tuple[str, str]] = Counter()
        for row in report.rows:
            key = kit_identity_key(row.title, row.mark)
            lost_counts[key] += row.file_count
            lost_mto[key] += row.mto_file_count
        for title, mark in report.fully_lost_kits:
            key = kit_identity_key(title, mark)
            print(
                f"  {title}/{mark}  files={lost_counts[key]}  "
                f"MTO={lost_mto[key]}"
            )
    print()
    print("top 10 title-marks by files to fix:")
    kit_totals: Counter[tuple[str, str]] = Counter()
    display: dict[tuple[str, str], tuple[str, str]] = {}
    for row in report.rows:
        key = kit_identity_key(row.title, row.mark)
        kit_totals[key] += row.file_count
        display.setdefault(key, (row.title, row.mark))
    for key, count in kit_totals.most_common(10):
        title, mark = display[key]
        print(f"  {count:5d}  {title}/{mark}")
    print()
    print(f"xlsx={XLSX_OUT}")
    print()
    print("--- text report (head) ---")
    text = format_layout_report_text(report)
    head = "\n".join(text.splitlines()[:80])
    print(head)


if __name__ == "__main__":
    main()
