"""One-workbook checks that reuse the full RFP / DS / UL collection tools.

Reports go under ``RFP сводный файл/_проверка_одного_файла/<контур>/``.
That folder is not a stamp of the full run, so Launch keeps the last full
``rfp_parts_net.xlsx``, ``Свод ДС для запуска.xlsx`` and the UL cache.
"""

from __future__ import annotations

import os
from pathlib import Path

from RFQ.ds_compare.tsd_packing_load import inspect_one_tsd_file
from RFQ.rfp_parts.analyze_rfp_parts import (
    DEFAULT_REPORTS_BASE_DIR,
    make_reports_out_dir,
    run_rfp_parts_analyze,
)
from RFQ.rfp_parts.ds_jobs import DsJobResult, run_ds_baseline_job

ONE_FILE_CHECK_DIR_NAME = "_проверка_одного_файла"


def one_file_kind_dir(kind: str, *, base: Path | None = None) -> Path:
    """Stable parent for one-file reports of ``kind`` (``RFP`` / ``ДС`` / ``УЛ``).

    Args:
        kind: Contour folder name.
        base: Reports base. Default is the RFP parts reports base.

    Returns:
        ``<base>/_проверка_одного_файла/<kind>``.
    """

    return Path(base or DEFAULT_REPORTS_BASE_DIR) / ONE_FILE_CHECK_DIR_NAME / kind


def one_file_stamp_dir(kind: str, *, base: Path | None = None) -> Path:
    """Unique ``YYYY.MM.DD_HH.MM`` folder under :func:`one_file_kind_dir`.

    The directory is not created here.

    Args:
        kind: Contour folder name.
        base: Reports base.

    Returns:
        Stamp path for this check.
    """

    return make_reports_out_dir(base=one_file_kind_dir(kind, base=base))


def _reject_workbook(path: Path) -> DsJobResult | None:
    if path.name.startswith("~$"):
        return DsJobResult(False, "это временный файл Excel (~$)", None)
    if not path.is_file():
        return DsJobResult(False, f"файл не найден: {path}", None)
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return DsJobResult(False, f"нужен xlsx или xlsm: {path.name}", None)
    return None


def run_rfp_one_file_job(file_path: str | Path) -> DsJobResult:
    """Run ``run_rfp_parts_analyze`` on one RFP workbook.

    Args:
        file_path: One part workbook.

    Returns:
        Russian summary. ``result_path`` is the stamp folder of this check.
        The full-run stamp under ``RFP сводный файл`` is left as it was.
    """

    os.environ.setdefault("PYTHONUTF8", "1")
    path = Path(file_path)
    rejected = _reject_workbook(path)
    if rejected is not None:
        return rejected
    out = one_file_stamp_dir("RFP")
    print(f"Проверка одного файла RFP: {path}", flush=True)
    print(f"Отчёт: {out}", flush=True)
    try:
        net = run_rfp_parts_analyze(
            parts_dir=path.parent,
            out_dir=out,
            only_files=[path],
            no_checklist=True,
        )
    except Exception as exc:
        return DsJobResult(
            False,
            f"проверка RFP не выполнена: {type(exc).__name__}: {exc}",
            str(out),
        )
    return DsJobResult(
        True,
        f"Один файл {path.name}. Свод проверки: {net}",
        str(out),
    )


def run_ds_one_file_job(
    file_path: str | Path,
    source_root: str | Path | None = None,
    registry_path: str | Path | None = None,
    ul_root: str | Path | None = None,
) -> DsJobResult:
    """Run the DS baseline audit on one workbook.

    Same registry load and parsers as ``run_ds_baseline_job``. Reports go to
    the one-file folder. ``Свод ДС для запуска.xlsx`` is not written.

    Args:
        file_path: One DS workbook.
        source_root: Trusted DS folder (identity and skip rules).
        registry_path: Registry the full audit reads.
        ul_root: Optional TSD root for registry UL checks.

    Returns:
        Russian summary. ``result_path`` is the audit stamp folder.
    """

    os.environ.setdefault("PYTHONUTF8", "1")
    path = Path(file_path)
    rejected = _reject_workbook(path)
    if rejected is not None:
        return rejected
    parent = one_file_kind_dir("ДС")
    print(f"Проверка одного файла ДС: {path}", flush=True)
    print(
        "Отчёты этой проверки пишутся в папку одного файла. "
        "«Свод ДС для запуска.xlsx» остаётся от полного прогона.",
        flush=True,
    )
    result = run_ds_baseline_job(
        source_root,
        registry_path,
        parent,
        ul_root,
        write_baseline=False,
        only_paths=(path,),
    )
    return DsJobResult(
        success=result.success,
        message=f"Один файл {path.name}. {result.message}",
        result_path=result.result_path,
    )


def run_ul_one_file_job(file_path: str | Path) -> DsJobResult:
    """Run the TSD packing reader on one workbook.

    Args:
        file_path: One packing-list xlsx.

    Returns:
        Russian summary. ``result_path`` is the stamp folder. The UL cache
        and the full summary stay untouched.
    """

    os.environ.setdefault("PYTHONUTF8", "1")
    path = Path(file_path)
    if path.name.startswith("~$"):
        return DsJobResult(False, "это временный файл Excel (~$)", None)
    if not path.is_file():
        return DsJobResult(False, f"файл не найден: {path}", None)
    if path.suffix.lower() != ".xlsx":
        return DsJobResult(
            False,
            f"упаковочный лист читается как xlsx: {path.name}",
            None,
        )
    out = one_file_stamp_dir("УЛ")
    print(f"Проверка одного файла УЛ: {path}", flush=True)
    print(f"Отчёт: {out}", flush=True)
    inspected = inspect_one_tsd_file(str(path), out)
    return DsJobResult(
        success=inspected.success,
        message=inspected.message,
        result_path=inspected.result_path,
    )
