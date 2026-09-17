"""Headless wrappers for legacy workflows previously wired through ``main.py`` GUI.

Callers supply paths explicitly (no folder/file pickers). Exceptions are raised as
``ValueError`` / ``RuntimeError`` with user-facing Russian messages where the legacy
GUI used ``tkinter.messagebox``.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import utils.path
from GUI.gui_constants import GuiConst
from RFQ.RFQ_compare import rfq_file_start
from RFQ.ds_compare.ds_merge_folder import merge_ds_folder
from RFQ.ds_compare.ds_compare_config import resolve_mto_path_from_config
from RFQ.ds_compare.ds_grouped_compare import (
    analyze_grouped_ds_specification,
    get_last_packing_compare_audit,
    get_last_rfq_quantity_audit,
)
from RFQ.ds_compare.ds_start_init import analyze_ds_specification
from RFQ.ds_compare.mto_ds_compare import mto_ds_list_compare
from RFQ.ds_compare.tsd_packing_load import load_and_cache_tsd_packing
from RFQ.ds_compare.upd_load import load_and_merge_upd
from RFQ.ds_compare.tsd_zinoviev_compare import (
    DEFAULT_ZINOVIEV_XLSX,
    compare_robot_vs_zinoviev,
)
from RFQ.mto_compare import (
    mto_chain_compare_start,
    mto_compare_start,
    mto_multi_compare_start,
)
from base import google_sheets
from base.base_xlsx_load import (
    FormulaCacheMissingError,
    backup_xlsx_to_old_subfolder,
    remove_strikethrough_from_xlsx,
    replace_formulas_with_cached_values,
    remove_hidden_sheets_for_1c,
)
from base.bbb_analysis import start_bbb_analysis
from base.bbb_config import load_config as load_bbb_config
from base.bbb_load import find_bbb_files
from base.excel_recalc_xlwings import recalculate_workbook_save
from cable_mapping import cab_mapping
from cable_mapping.utils_cm import get_files_list
from nano_cad import nc_start
from utils.path import get_path_from_file_path, get_files_single
from utils.release_zip import build_release_zip


@dataclass(frozen=True)
class ActionResult:
    """Outcome of a legacy action for a Qt (or other) host UI.

    Attributes:
        success: Whether the action completed without user-visible failure.
        message: Short human-readable summary (Russian allowed for parity with legacy UI).
        result_path: Primary output file path when applicable.
    """

    success: bool
    message: str
    result_path: str | None = None


def _as_str(path: str | Path) -> str:
    return str(path).strip()


def _require_dir(dir_path: str | Path) -> str:
    p = Path(_as_str(dir_path))
    if not p.is_dir():
        raise ValueError(f"Каталог не найден или это не папка: {p}")
    return str(p.resolve())


def _require_file(file_path: str | Path) -> str:
    p = Path(_as_str(file_path))
    if not p.is_file():
        raise ValueError(f"Файл не найден: {p}")
    return str(p.resolve())


def _get_mto_run_config() -> tuple[bool, dict[str, Any], dict[str, Any]]:
    cfg = load_bbb_config()
    mto_cfg = cfg.get("mto_vs_code_base", {})
    mto_corr_cfg = cfg.get("mto_vs_code_base_correction", {})
    run_mto_check = mto_cfg.get("enabled", True)
    return run_mto_check, mto_cfg, mto_corr_cfg


def run_mto_dwg_with_optional_bbb(dir_path: str | Path) -> ActionResult:
    """Run MTO and/or BBB checks for a DWG folder (same flags and rules as ``main.py``).

    Args:
        dir_path: Project DWG directory (name must be ``DWG`` when the setting requires it).

    Returns:
        ``ActionResult`` with the timestamped result directory path.

    Raises:
        ValueError: Invalid path or DWG folder naming rule violated.
        RuntimeError: Both checks disabled, or pipeline failure.
    """
    d = _require_dir(dir_path)
    cfg = load_bbb_config()
    folder_rules_cfg = cfg.get("folder_rules", {})
    search_only_in_dwg = folder_rules_cfg.get("search_only_in_dwg", True)
    if search_only_in_dwg:
        dir_name = os.path.basename(os.path.normpath(d))
        if dir_name.upper() != "DWG":
            raise ValueError(
                "Ограничение включено: выберите папку с именем DWG "
                "(или отключите флаг «Искать только в папке DWG» в настройках)."
            )

    mto_cfg = cfg.get("mto_vs_code_base", {})
    mto_corr_cfg = cfg.get("mto_vs_code_base_correction", {})
    bbb_cb_cfg = cfg.get("bbb_vs_code_base", {})
    bbb_mto_cfg = cfg.get("bbb_vs_mto", {})
    run_mto_check = mto_cfg.get("enabled", True)
    run_bbb_check = bbb_cb_cfg.get("enabled", True) or bbb_mto_cfg.get("enabled", True)

    if not run_mto_check and not run_bbb_check:
        raise RuntimeError("Обе проверки отключены в настройках (MTO и BBB).")

    result_dir = utils.path.get_path_out_dir(d, dir_result_prefix="/__результат_MTO_BBB_")
    mto_ran = False
    mto_skipped = False
    try:
        docs_in_dir = get_files_single(d, [".xlsx"])
        has_mto_in_dir = False
        if docs_in_dir != -1:
            has_mto_in_dir = any(getattr(doc, "doc_Type", "") == "MTO" for doc in docs_in_dir)

        if run_mto_check:
            if has_mto_in_dir:
                google_sheets.start(
                    d,
                    GuiConst.MTO_DIR,
                    out_dir=result_dir,
                    cfg=mto_cfg,
                    correction_cfg=mto_corr_cfg,
                )
                mto_ran = True
            else:
                mto_skipped = True
        if run_bbb_check:
            start_bbb_analysis(d, out_dir=result_dir)
    except Exception as e:
        raise RuntimeError(f"Ошибка объединённой проверки MTO/BBB: {e}") from e

    parts = [f"Результаты: {os.path.normpath(result_dir)}"]
    if mto_skipped:
        parts.append("MTO: пропущено (нет MTO .xlsx в папке).")
    elif mto_ran:
        parts.append("MTO: выполнено.")
    if run_bbb_check:
        parts.append("BBB: выполнено.")
    return ActionResult(True, " ".join(parts), result_path=os.path.normpath(result_dir))


def run_google_file_att(file_path: str | Path, att: Any) -> ActionResult:
    """Run ``google_sheets.start`` for a selected file (``att`` is a ``GuiConst`` mode key).

    Args:
        file_path: Path to the workbook to analyse.
        att: Mode constant (e.g. ``GuiConst.MTO_FILE``).

    Returns:
        Result with the output workbook path from ``google_sheets.start``.

    Raises:
        RuntimeError: MTO disabled in settings, or Google pipeline returned no path.
    """
    fp = _require_file(file_path)
    dir_path = get_path_from_file_path(fp)
    if att == GuiConst.MTO_FILE:
        run_mto_check, mto_cfg, mto_corr_cfg = _get_mto_run_config()
        if not run_mto_check:
            raise RuntimeError("Проверка MTO отключена в настройках.")
        result = google_sheets.start(
            dir_path,
            att,
            fp,
            cfg=mto_cfg,
            correction_cfg=mto_corr_cfg,
        )
    else:
        result = google_sheets.start(dir_path, att, fp)
    if not result:
        raise RuntimeError("Проверка не вернула путь к файлу результата (см. консольный лог).")
    return ActionResult(True, f"Готово. Результат: {result}", result_path=str(result))


def run_rfq_file(file_path: str | Path) -> ActionResult:
    """Run RFQ workflow for one workbook.

    Args:
        file_path: RFQ ``.xlsx`` file.

    Returns:
        Outcome with the same input path as ``result_path`` for convenience.

    Raises:
        RuntimeError: If ``rfq_file_start`` fails.
    """
    fp = _require_file(file_path)
    dir_path = get_path_from_file_path(fp)
    try:
        rfq_file_start(fp, dir_path)
    except Exception as e:
        raise RuntimeError(f"Ошибка RFQ: {e}") from e
    return ActionResult(True, "RFQ: обработка запущена/завершена (см. лог).", result_path=fp)


def remove_strikethrough_mto(file_path: str | Path) -> ActionResult:
    """Remove strikethrough text from sheet «Спецификация» in an MTO xlsx.

    Args:
        file_path: Path to the MTO workbook.

    Returns:
        Outcome pointing at the modified file.

    Raises:
        RuntimeError: openpyxl / IO failure.
    """
    fp = _require_file(file_path)
    try:
        remove_strikethrough_from_xlsx(fp, sheet_name="Спецификация")
    except Exception as e:
        raise RuntimeError(f"Ошибка снятия зачёркивания: {e}") from e
    return ActionResult(True, "Зачёркивание снято (лист «Спецификация»).", result_path=fp)


def remove_strikethrough_bbb_dir(dir_path: str | Path) -> ActionResult:
    """Remove strikethrough from BOE/BOM/BOQ sheets in the BBB folder.

    Args:
        dir_path: Directory scanned by ``find_bbb_files``.

    Returns:
        Outcome with the input directory as ``result_path``.

    Raises:
        ValueError: No BOE/BOM/BOQ files found.
        RuntimeError: Per-file processing errors.
    """
    d = _require_dir(dir_path)
    found = find_bbb_files(d)
    if not found:
        raise ValueError("BOE/BOM/BOQ файлы не найдены в выбранной папке.")
    errors: list[str] = []
    for doc_type in ("BOE", "BOM", "BOQ"):
        if doc_type not in found:
            continue
        doc = found[doc_type]
        path = doc.file_full_path
        try:
            remove_strikethrough_from_xlsx(path, sheet_name=doc_type)
        except Exception as e:
            errors.append(f"{doc_type}: {e}")
    if errors:
        raise RuntimeError("Ошибки при обработке:\n" + "\n".join(errors))
    return ActionResult(True, "Зачёркивание снято для найденных BBB.", result_path=d)


def prepare_bbb_for_1c(dir_path: str | Path) -> ActionResult:
    """Prepare BOE/BOM/BOQ for 1C export (same pipeline as ``main.py``).

    Args:
        dir_path: BBB directory.

    Returns:
        Outcome with the input directory as ``result_path``.

    Raises:
        ValueError: No BBB files found.
        RuntimeError: Excel/xlwings or formula-cache errors.
    """
    d = _require_dir(dir_path)
    found = find_bbb_files(d)
    if not found:
        raise ValueError("BOE/BOM/BOQ файлы не найдены в выбранной папке.")
    for doc_type in ("BOE", "BOM", "BOQ"):
        if doc_type not in found:
            continue
        path = found[doc_type].file_full_path
        try:
            backup_xlsx_to_old_subfolder(path)
            remove_strikethrough_from_xlsx(path, sheet_name=doc_type, create_backup=False)
            recalculate_workbook_save(path, label="до замены формул")
            replace_formulas_with_cached_values(path, sheet_name=doc_type)
            remove_hidden_sheets_for_1c(path, doc_type)
            recalculate_workbook_save(path, label="после скрытых листов")
        except FormulaCacheMissingError as e:
            raise RuntimeError(str(e)) from e
        except Exception as e:
            raise RuntimeError(f"Ошибка при подготовке {doc_type} для 1C: {e}") from e
    return ActionResult(True, "Подготовка BBB для 1C завершена.", result_path=d)


def run_cj_dir(dir_path: str | Path) -> ActionResult:
    """Cable journal / TPK folder check (``cab_mapping.cab_mapping_start``).

    Args:
        dir_path: Folder with TPK / CJ inputs.

    Returns:
        Success result; ``result_path`` is left ``None`` because ``cab_mapping`` opens
        Explorer internally and does not return the output path.

    Raises:
        ValueError: Empty or unreadable directory for this workflow.
        RuntimeError: Wrapped failure from ``cab_mapping_start``.
    """
    d = _require_dir(dir_path)
    listing = get_files_list(d, endswith=(".xlsx", ".XLSX", ".docx", ".DOCX"))
    if listing == -1:
        raise ValueError(f"Каталог недоступен: {d}")
    if not listing:
        raise ValueError("В папке нет файлов .xlsx / .docx для проверки ТПК.")
    try:
        cab_mapping.cab_mapping_start(d)
    except Exception as e:
        raise RuntimeError(f"Ошибка проверки ТПК: {e}") from e
    return ActionResult(
        True,
        "Проверка ТПК завершена (результат мог быть открыт в проводнике из legacy-кода).",
        result_path=None,
    )


def run_mto_compare_dir(dir_path: str | Path) -> ActionResult:
    """MTO vs MTO (pair-style) for all workbooks in a folder.

    Args:
        dir_path: Folder with ``.xlsx`` inputs.

    Returns:
        Outcome with the source folder as ``result_path`` (legacy code writes timestamped
        subfolders under it).

    Raises:
        RuntimeError: On pipeline failure (including ``exit``-free errors).
    """
    d = _require_dir(dir_path)
    try:
        mto_compare_start(pdf_path=d)
    except Exception as e:
        raise RuntimeError(f"MTO vs MTO: {e}") from e
    return ActionResult(True, "MTO vs MTO (2+) завершено (см. лог и папку результатов).", result_path=d)


def run_mto_multi_compare_dir(dir_path: str | Path) -> ActionResult:
    """Multi-document MTO/RFQ/output comparison for a folder.

    Args:
        dir_path: Folder with ``.xlsx`` inputs.

    Returns:
        Outcome with the source folder as ``result_path``.

    Raises:
        RuntimeError: On pipeline failure.
    """
    d = _require_dir(dir_path)
    try:
        mto_multi_compare_start(pdf_path=d)
    except Exception as e:
        raise RuntimeError(f"MTO multi: {e}") from e
    return ActionResult(True, "MTO multi-сравнение завершено (см. лог).", result_path=d)


def run_mto_chain_compare_dir(dir_path: str | Path) -> ActionResult:
    """Adjacent revision chain compare per title–mark group.

    Args:
        dir_path: Folder with MTO ``.xlsx`` files.

    Returns:
        Outcome with the source folder as ``result_path``.

    Raises:
        RuntimeError: Validation message from ``mto_chain_compare_start`` or unexpected error.
    """
    d = _require_dir(dir_path)
    try:
        err = mto_chain_compare_start(pdf_path=d)
    except Exception as e:
        raise RuntimeError(f"MTO — цепочка ревизий: неожиданная ошибка:\n{e!s}") from e
    if err:
        raise RuntimeError(err)
    return ActionResult(True, "MTO — цепочка ревизий: готово.", result_path=d)


def run_merge_ds_dir(dir_path: str | Path) -> ActionResult:
    """Merge DS xlsx files from a folder into one summary workbook.

    Args:
        dir_path: Folder with DS ``.xlsx`` files (non-recursive, same as ``main.py`` call).

    Returns:
        Outcome with path to the merged workbook.

    Raises:
        RuntimeError: Merge produced no output file.
    """
    d = _require_dir(dir_path)
    out_path = merge_ds_folder(
        d,
        include_subfolders=False,
        out_file_prefix="DS_summary",
        open_folder=False,
        dump_include_empty_row=False,
    )
    if not out_path:
        raise RuntimeError("Объединение ДС: нет результата (нет данных или неверная папка).")
    return ActionResult(True, f"Сводный файл: {out_path}", result_path=out_path)


def run_tsd_packing_load(dir_path: str | Path) -> ActionResult:
    """Walk TSD packing-list folder, load Single* sheets, write cache + summary.

    Args:
        dir_path: Root folder with packing-list ``.xlsx`` (recursive).

    Returns:
        Outcome with path to timestamped ``tsd_packing_summary_*.xlsx``.
        ``success`` is False when some files had unknown sheets (partial cache
        still written).
    """
    d = _require_dir(dir_path)
    try:
        result = load_and_cache_tsd_packing(d, force=True)
    except FileNotFoundError as e:
        raise ValueError(str(e)) from e
    except Exception as e:
        raise RuntimeError(f"Упаковочные листы (ТСД): {e}") from e
    stats = result.stats
    critical = result.critical
    verdict = (
        critical.format_short()
        if critical is not None
        else "критичные замечания не собраны"
    )
    msg = (
        f"{verdict} | файлов OK={stats.files_ok}, ошибок={stats.files_failed}, "
        f"позиций={stats.position_rows}. Свод: {result.summary_path}"
    )
    return ActionResult(result.success, msg, result_path=result.summary_path)


def run_upd_load(dir_path: str | Path) -> ActionResult:
    """Walk 1C UPD upload folder, keep position_row, write merged summary.

    Args:
        dir_path: Root folder with UPD ``.xlsx`` (recursive).

    Returns:
        Outcome with path to timestamped ``upd_summary_*.xlsx``.
        ``success`` is True when at least one ``position_row`` was written;
        unsupported ``.xls`` and foreign formats stay in the stats/report.
    """
    d = _require_dir(dir_path)
    try:
        result = load_and_merge_upd(d)
    except FileNotFoundError as e:
        raise ValueError(str(e)) from e
    except Exception as e:
        raise RuntimeError(f"УПД (файлы закачки): {e}") from e
    stats = result.stats
    msg = (
        f"{stats.format_short()}. Свод: {result.summary_path}"
    )
    return ActionResult(result.success, msg, result_path=result.summary_path)


def run_tsd_zinoviev_compare(
    zinoviev_path: str | Path | None = None,
) -> ActionResult:
    """Compare robot packing cache with Zinoviev manual TSD summary workbook.

    Args:
        zinoviev_path: Optional path to ``ТСД по всем ДС_общий.xlsx``;
            default is the network file under «Сводный от Зиновьева».

    Returns:
        Outcome with path to ``_результат_сравнения_*`` folder.
    """
    path = Path(zinoviev_path) if zinoviev_path else DEFAULT_ZINOVIEV_XLSX
    if not path.is_file():
        raise ValueError(f"Файл Зиновьева не найден:\n{path}")
    try:
        result = compare_robot_vs_zinoviev(path)
    except FileNotFoundError as e:
        raise ValueError(str(e)) from e
    except Exception as e:
        raise RuntimeError(f"Сравнение УЛ робот vs Зиновьев: {e}") from e
    return ActionResult(
        result.success,
        result.message,
        result_path=result.result_dir,
    )


def run_ds_mto_file(ds_file: str | Path) -> ActionResult:
    """DS specification file vs MTO folder from ``ds_compare_config``.

    Args:
        ds_file: DS workbook path.

    Returns:
        Outcome with the DS file path; detailed outputs follow legacy behaviour.

    Raises:
        ValueError: MTO folder not configured.
        RuntimeError: Analysis failure.
    """
    fp = _require_file(ds_file)
    mto_path = resolve_mto_path_from_config()
    if not mto_path:
        raise ValueError(
            "Не задана папка МТО. Выберите пресет в настройках ДС vs MTO (ds_compare_config)."
        )
    try:
        analyze_ds_specification(fp, mto_path=mto_path, summ_ds=True)
    except Exception as e:
        raise RuntimeError(f"ДС vs MTO: {e}") from e
    return ActionResult(True, "ДС vs MTO: анализ выполнен (см. лог / всплывающие окна legacy).", result_path=fp)


def run_grouped_ds_mto_file(
    ds_file: str | Path,
    rfq_file: str | Path | None = None,
    *,
    rfq_only: bool = False,
) -> ActionResult:
    """Grouped DS vs MTO comparison MVP (optional RFQ merge)."""
    fp = _require_file(ds_file)
    rfq_fp: str | None = None
    if rfq_file is not None:
        rfq_fp = str(_require_file(rfq_file))
    elif rfq_only:
        raise ValueError("RFQ-only: укажите файл RFQ (TPK xlsx).")
    mto_path = resolve_mto_path_from_config()
    if not mto_path:
        raise ValueError(
            "Не задана папка МТО. Выберите пресет в настройках ДС vs MTO (ds_compare_config)."
        )
    try:
        out_path = analyze_grouped_ds_specification(
            fp,
            mto_path=mto_path,
            rfq_path=rfq_fp,
            summ_ds=True,
            rfq_only=rfq_only,
        )
    except Exception as e:
        label = "Grouped ДС vs MTO (только RFQ)" if rfq_only else "Grouped ДС vs MTO"
        raise RuntimeError(f"{label}: {e}") from e
    if not out_path:
        raise RuntimeError(
            "Grouped ДС vs MTO: итоговый Excel не создан. "
            "Проверьте выбранные столбцы и доступ на запись в папку результата."
        )
    msg = (
        "Grouped ДС vs MTO (только RFQ): готово."
        if rfq_only
        else "Grouped ДС vs MTO: анализ выполнен."
    )
    audit = get_last_rfq_quantity_audit()
    if audit is not None:
        msg = f"{msg} {audit.format_short()}"
    packing_audit = get_last_packing_compare_audit()
    if packing_audit is not None:
        msg = f"{msg} {packing_audit.format_short()}"
    return ActionResult(
        True,
        msg,
        result_path=out_path,
    )


def run_grouped_ds_mto_rfq_only_file(
    ds_file: str | Path,
    rfq_file: str | Path,
) -> ActionResult:
    """Re-run RFQ merge and Excel using cached grouped DS vs MTO result."""
    return run_grouped_ds_mto_file(ds_file, rfq_file, rfq_only=True)


def run_mto_ds_file(ds_file: str | Path) -> ActionResult:
    """MTO vs DS list file comparison.

    Args:
        ds_file: Input workbook for ``mto_ds_list_compare``.

    Returns:
        Outcome with the input path; legacy code may show Tk popups.

    Raises:
        RuntimeError: On failure.
    """
    fp = _require_file(ds_file)
    try:
        mto_ds_list_compare(fp)
    except Exception as e:
        raise RuntimeError(f"MTO vs список ДС: {e}") from e
    return ActionResult(True, "MTO vs список ДС: готово (см. лог).", result_path=fp)


def run_nanocad_db(file_path: str | Path) -> ActionResult:
    """Load a NanoCAD database file through ``nc_start.load_db``.

    Args:
        file_path: Path to the database file.

    Returns:
        Outcome with the same file path.

    Raises:
        RuntimeError: If ``load_db`` fails.
    """
    fp = _require_file(file_path)
    try:
        nc_start.load_db(fp)
    except Exception as e:
        raise RuntimeError(f"NanoCAD.db: {e}") from e
    return ActionResult(True, "NanoCAD.db: загрузка выполнена.", result_path=fp)


def run_release_zip(project_root: str | Path) -> ActionResult:
    """Build a colleague-facing release zip from the project tree.

    Args:
        project_root: Repository root directory.

    Returns:
        Outcome with path to the created ``.zip``.

    Raises:
        RuntimeError: If ``build_release_zip`` fails.
    """
    root = _require_dir(project_root)
    try:
        zip_path, included_files, skipped_files = build_release_zip(root)
    except Exception as e:
        raise RuntimeError(f"Ошибка сборки ZIP: {e}") from e
    msg = f"Архив создан. Включено файлов: {included_files}; пропущено: {skipped_files}."
    return ActionResult(True, msg, result_path=str(zip_path))


def run_rfp_checklist_compare() -> ActionResult:
    """Compare maintained DS checklist xlsx to the RFP parts folder.

    Prints progress and result tables to stdout (mirrored into Job monitor).

    Returns:
        Outcome; ``success`` is False on load failure or missing expected DS
        (WARN-only mismatches still count as success).
    """
    from RFQ.rfp_parts.ds_checklist import run_checklist_compare

    result = run_checklist_compare(progress=True)
    has_errors = bool(result.load_error or result.error_count)
    return ActionResult(
        not has_errors,
        result.summary_line(),
        result_path=str(result.checklist_path),
    )


__all__ = [
    "ActionResult",
    "prepare_bbb_for_1c",
    "remove_strikethrough_bbb_dir",
    "remove_strikethrough_mto",
    "run_cj_dir",
    "run_ds_mto_file",
    "run_google_file_att",
    "run_merge_ds_dir",
    "run_mto_chain_compare_dir",
    "run_mto_compare_dir",
    "run_mto_ds_file",
    "run_mto_dwg_with_optional_bbb",
    "run_mto_multi_compare_dir",
    "run_nanocad_db",
    "run_release_zip",
    "run_rfp_checklist_compare",
    "run_rfq_file",
    "run_tsd_packing_load",
    "run_tsd_zinoviev_compare",
    "run_upd_load",
]
