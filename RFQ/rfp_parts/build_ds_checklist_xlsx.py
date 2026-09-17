"""Build DS increase/decrease checklist xlsx from the summary table.

Reads column 3 of the Bi.Si.Si. sheet, classifies DS numbers, checks UNC RFP
folders, and writes ``ДС_увеличение_уменьшение.xlsx`` with status colours.
"""

from __future__ import annotations

import argparse
import os
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill

from RFQ.rfp_parts.ds_checklist import (
    ChecklistCompareResult,
    DsChecklist,
    DEFAULT_DECREASE_DIR,
    DEFAULT_INCREASE_DIR,
    collect_folder_ds,
    compare_checklist_to_folders,
    load_ds_checklist,
    present_in,
)

_PACKAGE_DIR = Path(__file__).resolve().parent
DEFAULT_OUT = _PACKAGE_DIR / "ДС_увеличение_уменьшение.xlsx"
_SUMMARY_GLOB = "Сводная таблица ДС по вед.договорам*.xlsx"

_MAIN_DS_RE = re.compile(r"^\d+(?:/[^\s]+)?$")
_DEC_WORD_RE = re.compile(r"уменьш", re.IGNORECASE)
_HEADER_SKIP = {"№ ДС", "ДС/договор"}
_DS_PREFIX_RE = re.compile(r"^ДС\s*", re.IGNORECASE)

# Status columns: B, E, G, H (user column order; C is spacer).
_STATUS_COLS = (2, 5, 7, 8)
_FILL_YES = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
_FILL_NO = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


@dataclass(frozen=True)
class ClassifiedDsLists:
    """DS labels classified from the source summary workbook."""

    increase: list[str]
    decrease: list[str]
    other: list[str]


@dataclass(frozen=True)
class ChecklistUpdateResult:
    """Outcome of synchronizing the maintained checklist with its source."""

    status: Literal["unchanged", "updated", "error"]
    changes_count: int
    summary_path: Path
    checklist_path: Path
    increase_dir: Path
    decrease_dir: Path
    compare_result: ChecklistCompareResult | None
    error: str | None
    summary_line: str


def _sort_key(s: str) -> tuple:
    s_norm = _DS_PREFIX_RE.sub("", str(s).upper().replace(" ", ""))
    left = s_norm.split("/", 1)[0].split("_", 1)[0]
    m = re.match(r"^(\d+)(.*)$", left)
    if m:
        return (0, int(m.group(1)), m.group(2), s)
    return (1, 10**9, s, s)


def _canonical_main_key(ds: str) -> str | None:
    """Normalize ``ДС14`` / ``14`` / ``14А`` to a main-series identity key."""
    s = str(ds).strip().upper().replace(" ", "")
    s = _DS_PREFIX_RE.sub("", s)
    if "/" in s:
        return s
    m = re.fullmatch(r"(\d+[А-ЯA-Z]*)", s)
    return m.group(1) if m else None


def _yn(flag: bool) -> str:
    return "ЕСТЬ" if flag else "НЕТ"

def _find_summary_xlsx(directory: Path) -> Path:
    matches = sorted(Path(directory).glob(_SUMMARY_GLOB))
    if not matches:
        raise FileNotFoundError(
            f"No summary xlsx matching {_SUMMARY_GLOB!r} in {directory}"
        )
    return matches[-1]


def resolve_summary_xlsx(summary_path: Path) -> Path:
    """Return the configured summary file, or a glob match in the same folder.

    The exact ``summary_path`` wins when it exists. Otherwise the newest file
    matching ``Сводная таблица ДС по вед.договорам*.xlsx`` in that directory is
    used so a renamed source still feeds the checklist updater.
    """
    path = Path(summary_path)
    if path.is_file():
        return path
    parent = path.parent
    if parent.is_dir():
        matches = sorted(parent.glob(_SUMMARY_GLOB))
        if matches:
            return matches[-1]
    return path


def normalize_ds_identity(ds: object) -> str:
    """Return a case/space-insensitive DS identity.

    A leading ``ДС`` is removed and slash/underscore forms are equivalent.
    """
    value = re.sub(r"\s+", "", str(ds).upper())
    value = _DS_PREFIX_RE.sub("", value)
    return value.replace("/", "_")


def classify_ds_summary(summary_path: Path) -> ClassifiedDsLists:
    """Read and classify DS labels from a source summary workbook.

    Args:
        summary_path: Source workbook containing sheet ``Би.Си.Си.``.

    Returns:
        Classified increase, decrease, and other labels.

    Raises:
        OSError: If the workbook cannot be read.
        KeyError: If the required sheet is absent.
    """
    workbook = load_workbook(summary_path, read_only=True, data_only=True)
    try:
        sheet = workbook["Би.Си.Си."]
        decrease: set[str] = set()
        increase: set[str] = set()
        other_raw: set[str] = set()

        for row in sheet.iter_rows(min_row=1, max_col=4, values_only=True):
            raw = row[2] if len(row) >= 3 else None
            if raw is None:
                continue
            ds = str(raw).strip()
            if not ds or ds in _HEADER_SKIP:
                continue
            desc = str(row[3]).strip() if len(row) >= 4 and row[3] is not None else ""
            if _MAIN_DS_RE.match(ds):
                if "/" in ds or _DEC_WORD_RE.search(desc):
                    decrease.add(ds)
                else:
                    increase.add(ds)
            else:
                other_raw.add(ds)
    finally:
        workbook.close()

    increase -= decrease

    main_keys: set[str] = set()
    for ds in increase | decrease:
        canonical = _canonical_main_key(ds)
        if canonical:
            main_keys.add(canonical)
            main_keys.add(canonical.split("/")[0])

    other: set[str] = set()
    for ds in other_raw:
        canonical = _canonical_main_key(ds)
        if canonical is not None and canonical in main_keys:
            continue
        other.add(ds)

    return ClassifiedDsLists(
        increase=sorted(increase, key=_sort_key),
        decrease=sorted(decrease, key=_sort_key),
        other=sorted(other, key=_sort_key),
    )


def _classify_from_summary(summary_path: Path) -> tuple[list[str], list[str], list[str]]:
    """Backward-compatible tuple form of :func:`classify_ds_summary`."""
    classified = classify_ds_summary(summary_path)
    return classified.increase, classified.decrease, classified.other


def _add_status_conditional_formatting(worksheet, max_row: int) -> None:
    """ЕСТЬ → green, НЕТ → yellow on status columns (same idea as Excel CF)."""
    if max_row < 2:
        return
    for col in _STATUS_COLS:
        letter = chr(64 + col)
        cell_range = f"{letter}2:{letter}{max_row}"
        worksheet.conditional_formatting.add(
            cell_range,
            CellIsRule(operator="equal", formula=['"ЕСТЬ"'], fill=_FILL_YES),
        )
        worksheet.conditional_formatting.add(
            cell_range,
            CellIsRule(operator="equal", formula=['"НЕТ"'], fill=_FILL_NO),
        )


def create_ds_checklist_workbook(
    classified: ClassifiedDsLists,
    *,
    increase_dir: Path,
    decrease_dir: Path,
    not_required_by_identity: dict[str, bool] | None = None,
) -> Workbook:
    """Create a checklist workbook from already classified DS lists.

    Args:
        classified: Source summary classification.
        increase_dir: Folder used to calculate increase presence statuses.
        decrease_dir: Folder used to calculate decrease presence statuses.
        not_required_by_identity: Manual ``Не требуется`` flags keyed only by
            normalized DS identity, independent of category.

    Returns:
        An unsaved openpyxl workbook. The caller owns and must close it.
    """
    preserved = not_required_by_identity or {}
    inc_folder = collect_folder_ds(increase_dir)
    dec_folder = collect_folder_ds(decrease_dir)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "ДС"

    headers = {
        1: "ДС на увеличение",
        2: "В папке увеличение",
        4: "ДС на уменьшение",
        5: "В папке уменьшение",
        6: "Прочее (не увелич/уменьш)",
        7: "Прочее в папке увеличение",
        8: "Прочее в папке уменьшение",
    }
    for col, title in headers.items():
        cell = worksheet.cell(1, col, title)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    row_count = max(
        len(classified.increase),
        len(classified.decrease),
        len(classified.other),
        1,
    )
    for index in range(row_count):
        row = index + 2
        if index < len(classified.increase):
            ds = classified.increase[index]
            status = (
                "Не требуется"
                if preserved.get(normalize_ds_identity(ds), False)
                else _yn(present_in(inc_folder, ds))
            )
            worksheet.cell(row, 1, ds)
            worksheet.cell(row, 2, status)
        if index < len(classified.decrease):
            ds = classified.decrease[index]
            status = (
                "Не требуется"
                if preserved.get(normalize_ds_identity(ds), False)
                else _yn(present_in(dec_folder, ds))
            )
            worksheet.cell(row, 4, ds)
            worksheet.cell(row, 5, status)
        if index < len(classified.other):
            ds = classified.other[index]
            worksheet.cell(row, 6, ds)
            worksheet.cell(row, 7, _yn(present_in(inc_folder, ds)))
            worksheet.cell(row, 8, _yn(present_in(dec_folder, ds)))

    for col, width in {
        1: 20,
        2: 18,
        3: 3,
        4: 20,
        5: 18,
        6: 28,
        7: 22,
        8: 22,
    }.items():
        worksheet.column_dimensions[chr(64 + col)].width = width
    worksheet.row_dimensions[1].height = 30
    _add_status_conditional_formatting(worksheet, row_count + 1)
    return workbook


def write_ds_checklist_workbook(
    out_path: Path,
    classified: ClassifiedDsLists,
    *,
    increase_dir: Path,
    decrease_dir: Path,
    not_required_by_identity: dict[str, bool] | None = None,
) -> Path:
    """Build and save a checklist workbook, closing it in all cases."""
    workbook = create_ds_checklist_workbook(
        classified,
        increase_dir=increase_dir,
        decrease_dir=decrease_dir,
        not_required_by_identity=not_required_by_identity,
    )
    try:
        workbook.save(out_path)
    finally:
        workbook.close()
    return out_path


def _semantic_composition(classified: ClassifiedDsLists) -> set[tuple[str, str]]:
    return {
        *((("increase", normalize_ds_identity(ds)) for ds in classified.increase)),
        *((("decrease", normalize_ds_identity(ds)) for ds in classified.decrease)),
        *((("other", normalize_ds_identity(ds)) for ds in classified.other)),
    }


def _checklist_composition(checklist: DsChecklist) -> set[tuple[str, str]]:
    return {
        *((("increase", normalize_ds_identity(item.ds)) for item in checklist.increase)),
        *((("decrease", normalize_ds_identity(item.ds)) for item in checklist.decrease)),
        *((("other", normalize_ds_identity(ds)) for ds in checklist.other)),
    }


def _not_required_map(checklist: DsChecklist) -> dict[str, bool]:
    return {
        normalize_ds_identity(item.ds): True
        for item in checklist.increase + checklist.decrease
        if item.not_required
    }


def _compare_safely(
    checklist_path: Path,
    increase_dir: Path,
    decrease_dir: Path,
    *,
    progress: bool,
) -> ChecklistCompareResult:
    try:
        checklist = load_ds_checklist(checklist_path)
        return compare_checklist_to_folders(
            checklist,
            increase_dir=increase_dir,
            decrease_dir=decrease_dir,
            progress=progress,
        )
    except Exception as exc:
        return ChecklistCompareResult(checklist_path=checklist_path, load_error=str(exc))


def ensure_ds_checklist_current(
    summary_path: Path,
    checklist_path: Path,
    increase_dir: Path,
    decrease_dir: Path,
    progress: bool = False,
) -> ChecklistUpdateResult:
    """Synchronize a checklist with its source and compare it to folders.

    Semantic changes compare category plus normalized DS identity. Therefore a
    category move contributes two items to ``changes_count`` (symmetric
    difference). Manual ``Не требуется`` flags follow identity only.

    If the source workbook cannot be read but ``checklist_path`` already exists
    and is readable, the checklist is left unchanged and folder comparison still
    runs. That is not a source ``error``: the pipeline is gated by folder
    compare. A missing source is still ``error`` when the checklist cannot be
    created or read.

    Args:
        summary_path: Source summary workbook.
        checklist_path: Maintained checklist workbook.
        increase_dir: Increase RFP folder.
        decrease_dir: Decrease RFP folder.
        progress: Whether folder comparison prints progress.

    Returns:
        ``ChecklistUpdateResult`` with status ``unchanged``, ``updated``, or
        ``error`` and the folder comparison result when available.
    """
    summary = resolve_summary_xlsx(summary_path)
    checklist_target = Path(checklist_path)
    inc_dir = Path(increase_dir)
    dec_dir = Path(decrease_dir)
    changes_count = 0
    temp_path: Path | None = None

    try:
        classified = classify_ds_summary(summary)
    except Exception as exc:
        compare_result = _compare_safely(
            checklist_target, inc_dir, dec_dir, progress=progress
        )
        can_keep_checklist = (
            checklist_target.is_file() and not compare_result.load_error
        )
        if can_keep_checklist:
            message = (
                "WARN: источник списка ДС недоступен "
                f"({exc}); текущий список оставлен без изменений; "
                f"{compare_result.summary_line()}"
            )
            return ChecklistUpdateResult(
                "unchanged",
                0,
                summary,
                checklist_target,
                inc_dir,
                dec_dir,
                compare_result,
                None,
                message,
            )
        message = f"ERROR: не удалось прочитать источник списка ДС: {exc}"
        return ChecklistUpdateResult(
            "error",
            0,
            summary,
            checklist_target,
            inc_dir,
            dec_dir,
            compare_result,
            str(exc),
            message,
        )

    source_composition = _semantic_composition(classified)
    preserved: dict[str, bool] = {}
    if checklist_target.exists():
        try:
            current = load_ds_checklist(checklist_target)
        except Exception as exc:
            compare_result = _compare_safely(
                checklist_target, inc_dir, dec_dir, progress=progress
            )
            message = f"ERROR: не удалось прочитать текущий список ДС: {exc}"
            return ChecklistUpdateResult(
                "error",
                0,
                summary,
                checklist_target,
                inc_dir,
                dec_dir,
                compare_result,
                str(exc),
                message,
            )
        current_composition = _checklist_composition(current)
        changes_count = len(source_composition.symmetric_difference(current_composition))
        preserved = _not_required_map(current)
    else:
        changes_count = len(source_composition)

    status: Literal["unchanged", "updated", "error"] = "unchanged"
    error: str | None = None
    if not checklist_target.exists() or changes_count:
        try:
            checklist_target.parent.mkdir(parents=True, exist_ok=True)
            handle = tempfile.NamedTemporaryFile(
                mode="wb",
                suffix=".xlsx",
                prefix=f".{checklist_target.stem}.",
                dir=checklist_target.parent,
                delete=False,
            )
            temp_path = Path(handle.name)
            handle.close()
            write_ds_checklist_workbook(
                temp_path,
                classified,
                increase_dir=inc_dir,
                decrease_dir=dec_dir,
                not_required_by_identity=preserved,
            )
            os.replace(temp_path, checklist_target)
            temp_path = None
            status = "updated"
        except Exception as exc:
            status = "error"
            error = str(exc)
        finally:
            if temp_path is not None:
                try:
                    temp_path.unlink(missing_ok=True)
                except OSError:
                    pass

    compare_result = _compare_safely(
        checklist_target, inc_dir, dec_dir, progress=progress
    )
    if status == "error":
        summary_line = f"ERROR: не удалось обновить список ДС: {error}"
    elif status == "updated":
        summary_line = (
            f"UPDATED: список ДС обновлён (изменений состава: {changes_count}); "
            f"{compare_result.summary_line()}"
        )
    else:
        summary_line = f"UNCHANGED: список ДС актуален; {compare_result.summary_line()}"
    return ChecklistUpdateResult(
        status,
        changes_count,
        summary,
        checklist_target,
        inc_dir,
        dec_dir,
        compare_result,
        error,
        summary_line,
    )


def build_ds_checklist_xlsx(
    summary_path: Path | None = None,
    out_path: Path | None = None,
    increase_dir: Path | None = None,
    decrease_dir: Path | None = None,
) -> Path:
    """Build the checklist workbook and return the output path.

    Args:
        summary_path: Source summary xlsx. Defaults to newest match in package dir.
        out_path: Output path. Defaults to ``ДС_увеличение_уменьшение.xlsx``.
        increase_dir: RFP increase folder. Defaults to package UNC constant.
        decrease_dir: RFP decrease folder. Defaults to package UNC constant.

    Returns:
        Path to the written xlsx.
    """
    summary = summary_path or _find_summary_xlsx(_PACKAGE_DIR)
    out = out_path or DEFAULT_OUT
    inc_dir = Path(increase_dir or DEFAULT_INCREASE_DIR)
    dec_dir = Path(decrease_dir or DEFAULT_DECREASE_DIR)

    classified = classify_ds_summary(summary)
    return write_ds_checklist_workbook(
        out,
        classified,
        increase_dir=inc_dir,
        decrease_dir=dec_dir,
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--summary", type=Path, default=None, help="Source summary xlsx")
    parser.add_argument("--out", type=Path, default=None, help="Output checklist xlsx")
    args = parser.parse_args(argv)
    path = build_ds_checklist_xlsx(summary_path=args.summary, out_path=args.out)
    print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
