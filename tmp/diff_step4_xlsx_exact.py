"""Read-only: positional cell-value and comment-text diff of two Step4 xlsx files."""

from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

DIR = Path(__file__).resolve().parent / "step4_compare"
OLD = DIR / "Шаг4_Сопоставление_RFP_MTO_20260909_131931.xlsx"
NEW = DIR / "Шаг4_Сопоставление_RFP_MTO_20260909_134240.xlsx"
MAX_SHOWN = 15


def collect(path: Path) -> dict[str, tuple[list[tuple], dict[str, str]]]:
    wb = openpyxl.load_workbook(path, data_only=False)
    try:
        out: dict[str, tuple[list[tuple], dict[str, str]]] = {}
        for sheet in wb.worksheets:
            values: list[tuple] = []
            comments: dict[str, str] = {}
            for row in sheet.iter_rows():
                values.append(tuple(cell.value for cell in row))
                for cell in row:
                    if cell.comment is not None:
                        comments[cell.coordinate] = cell.comment.text
            out[sheet.title] = (values, comments)
        return out
    finally:
        wb.close()


def main() -> None:
    old = collect(OLD)
    new = collect(NEW)
    total_value_diffs = 0
    total_comment_diffs = 0
    for title, (old_values, old_comments) in old.items():
        new_values, new_comments = new.get(title, ([], {}))
        print(f"\n=== лист «{title}» ===")
        if len(old_values) != len(new_values):
            print(f"  РАЗНОЕ ЧИСЛО СТРОК: old={len(old_values)} new={len(new_values)}")
            total_value_diffs += 1
            continue
        shown = 0
        diffs = 0
        for index, (old_row, new_row) in enumerate(zip(old_values, new_values), start=1):
            if old_row == new_row:
                continue
            for column, (a, b) in enumerate(zip(old_row, new_row), start=1):
                if a != b:
                    diffs += 1
                    if shown < MAX_SHOWN:
                        shown += 1
                        print(f"  значение r{index}c{column}: old={a!r} new={b!r}")
        print(f"  значения: расхождений {diffs}")
        total_value_diffs += diffs

        comment_diffs = 0
        shown = 0
        for coord in sorted(set(old_comments) | set(new_comments)):
            a = old_comments.get(coord)
            b = new_comments.get(coord)
            if a != b:
                comment_diffs += 1
                if shown < MAX_SHOWN:
                    shown += 1
                    print(f"  комментарий {coord}:\n    old={a!r}\n    new={b!r}")
        print(f"  комментарии: {len(old_comments)} шт., расхождений {comment_diffs}")
        total_comment_diffs += comment_diffs

    print(
        f"\nИТОГО: расхождений значений {total_value_diffs}, "
        f"расхождений комментариев {total_comment_diffs}"
    )


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    main()
