"""
Утилиты для работы с RFP тегами
Функции для работы с путями и директориями результатов
"""

from __future__ import annotations

import copy
import json
import os
import re
import sys
import hashlib
import tracemalloc
from datetime import datetime
from dataclasses import dataclass
from typing import List, Tuple, Dict, Any, Optional

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from RFQ.rfp_parts.ds_checklist import (
    DEFAULT_BASE,
    DEFAULT_CHECKLIST_FILE,
    DEFAULT_DECREASE_DIR,
    DEFAULT_INCREASE_DIR,
)
from RFQ.rfp_parts.ds_registry import DEFAULT_REGISTRY_PATH
from utils.path import make_dir

_tracemalloc_started: bool = False
_memory_log_enabled: bool = True
_TIMING_LOG_FILENAME = "timing_log.xlsx"
_TIMING_SHEET_NAME = "timings"
_TIMING_TOP_SHEET_NAME = "top_longest"

_timing_buffer: List[Tuple[str, str, Optional[float], str]] = []
_timing_buffer_result_dir: Optional[str] = None

# Имя файла конфигурации (в корневой папке rfp_tags_utils.py)
_CONFIG_FILENAME = "rfp_tags_compare_config.json"
_COLUMN_OPTIMIZATION_DEFAULT_FILENAME = "column_optimization_default.json"

# Единицы измерения: для них не выполняется раскладка position_row в «по 1» (RFP/MTO).
DEFAULT_UNITS_SPLIT_BAN = ["м", "м2", "бухта", "кг", "т"]

# Режим участия RFP в пайплайне (step1 + step4 include_mto_vo_without_rfp_anchor).
RFP_PIPELINE_MODE_STANDARD = "standard"
RFP_PIPELINE_MODE_WITH_ORPHAN_MTO_VO = "with_orphan_mto_vo"
RFP_PIPELINE_MODE_MTO_VO_ONLY = "mto_vo_only"
RFP_PIPELINE_MODES_ALL = (
    RFP_PIPELINE_MODE_STANDARD,
    RFP_PIPELINE_MODE_WITH_ORPHAN_MTO_VO,
    RFP_PIPELINE_MODE_MTO_VO_ONLY,
)


def resolve_rfp_pipeline_mode(step1_cfg: Optional[Dict[str, Any]] = None) -> str:
    """Возвращает нормализованный rfp_pipeline_mode из step1 конфига.

    Поддерживает устаревший ключ ``skip_rfp_load: true`` как ``mto_vo_only``.

    Args:
        step1_cfg: Секция ``config[\"step1\"]`` или None.

    Returns:
        Одна из строк ``RFP_PIPELINE_MODES_ALL``.
    """
    if not isinstance(step1_cfg, dict):
        return RFP_PIPELINE_MODE_STANDARD
    if step1_cfg.get("skip_rfp_load") is True:
        return RFP_PIPELINE_MODE_MTO_VO_ONLY
    raw = step1_cfg.get("rfp_pipeline_mode", RFP_PIPELINE_MODE_STANDARD)
    if not isinstance(raw, str):
        raw = str(raw)
    norm = raw.strip()
    if norm in RFP_PIPELINE_MODES_ALL:
        return norm
    low = norm.lower().replace("-", "_")
    if low in ("only_mto_vo", "mto_vo_only", "no_rfp"):
        return RFP_PIPELINE_MODE_MTO_VO_ONLY
    if low in ("with_orphan", "orphan_mto_vo", "with_orphan_mto_vo"):
        return RFP_PIPELINE_MODE_WITH_ORPHAN_MTO_VO
    return RFP_PIPELINE_MODE_STANDARD


def resolve_units_split_ban_config(config: Optional[Dict[str, Any]] = None) -> List[str]:
    """
    Список единиц (нижний регистр), для которых не делается раскладка в единичные строки.
    Если в конфиге units_split_ban.enabled=false — возвращает [] (раскладка для всех допустимых случаев).
    """
    if config is None:
        return list(DEFAULT_UNITS_SPLIT_BAN)
    block = config.get("units_split_ban")
    if not isinstance(block, dict):
        return list(DEFAULT_UNITS_SPLIT_BAN)
    if not block.get("enabled", True):
        return []
    units = block.get("units")
    if isinstance(units, list):
        cleaned = [str(u).strip().lower() for u in units if str(u).strip()]
        if cleaned:
            return cleaned
    return list(DEFAULT_UNITS_SPLIT_BAN)


def _get_config_path() -> str:
    """Путь к файлу конфигурации в папке rfp_tags_utils.py"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, _CONFIG_FILENAME)


def _load_column_optimization_default() -> Dict[str, Any]:
    """Загружает column_optimization из отдельного файла по умолчанию."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(script_dir, _COLUMN_OPTIMIZATION_DEFAULT_FILENAME)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"Ошибка загрузки {path}: {e}")
        return {}


def get_default_config() -> Dict[str, Any]:
    """Возвращает конфигурацию по умолчанию"""
    base = {
        "memory_log": True,
        "load_tags": True,
        "paths": {
            "rfp_path": r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\Сводная RFP.xlsx",
            "rfp_registr_lot_path": r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\registr_lot.xlsx",
            "mto_path": r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_ГОТОВЫЕ_ДЛЯ_РОБОТА",
            "vo_path": r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РКД\Материалы шкафов из РКД для робота",
            "code_ban_file": r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\code_ban.txt",
            "units_convert_matrix": r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\матрица_ед_изм.xlsx",
            "ds_manager_matrix": r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\Список ДС - Фамилии МП.xlsx",
            "gem_supply_codes": r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\Коды_Поставок_ГЭМ_8950.xlsx",
            "replacement_table_file": r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\Таблички_графики\14_сравнение с ДС\_замены кодов\Code_Replacement_Table.xlsx",
            "result_dir_base": r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_РЕЗУЛЬТАТА_ПРОВЕРКИ",
        },
        "rfp_parts": {
            "auto_update_checklist": True,
            "use_latest_net": True,
            "input_mode": "legacy_net",
            "ds_source_dir": "",
            "ds_registry_path": str(DEFAULT_REGISTRY_PATH),
            "summary_path": str(
                DEFAULT_BASE / "Сводная таблица ДС по вед.договорам.xlsx"
            ),
            "checklist_path": str(DEFAULT_CHECKLIST_FILE),
            "increase_dir": str(DEFAULT_INCREASE_DIR),
            "decrease_dir": str(DEFAULT_DECREASE_DIR),
        },
        "step1": {
            "debug": False,
            "skip_split": False,
            "use_rfp_code_ban": True,
            "rfp_pipeline_mode": RFP_PIPELINE_MODE_STANDARD,
        },
        "units_split_ban": {
            "enabled": True,
            "units": list(DEFAULT_UNITS_SPLIT_BAN),
        },
        "step2": {
            "debug": False,
            "export_load_results_excel": True,
            "export_positions_database_excel": False,
            "flat_mto_structure": False,
        },
        "step3": {"debug": False, "export_to_excel": False},
        "step4": {
            "debug": True,
            "debug_step4_1": True,
            "debug_step4_2": True,
            "debug_step4_3": True,
            "debug_step4_4": True,
            "collapse_debug": False,
            "unified_debug": True,
            "timing_log": True,
            "use_multiprocessing": False,
            "parallel_by_title_mark": False,
            "max_workers": 4,
            "verbose_progress_messages": True,
            "memory_top_stats": False,
            "memory_top_n": 15,
            "debug_tag": "",
            "debug_code": "BCC0003424",
            "debug_title_system": "8445-SOT1",
            "print_title_systems_table": False,
            "export_title_systems_comparison_excel": True,
            "assign_rfp_mto_code_compare_colors": True,
            "filter_to_mto_titles": False,
            "include_mto_vo_without_rfp_anchor": False,
            "include_packing_lists": True,
            "export_bcc_accum_matrix": True,
            "ul_match_use_mto_tags": True,
        },
        "rfp_tags_utils": {
            "finalize_timing_top_n": 5,
            "save_input_fingerprints": True,
        },
        "step4_excel_columns": {
            "active_id": "default",
            "templates": [],
        },
    }
    col_opt = _load_column_optimization_default()
    if col_opt:
        base["column_optimization"] = col_opt
    return base


def load_config() -> Dict[str, Any]:
    """Загружает конфигурацию из файла. При отсутствии или ошибке — возвращает значения по умолчанию."""
    config_path = _get_config_path()
    default = get_default_config()
    if not os.path.exists(config_path):
        return default
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        # Мержим с дефолтами, чтобы новые поля подтягивались
        return _deep_merge(default, loaded)
    except Exception as e:
        print(f"Ошибка загрузки конфига {config_path}: {e}")
        return default


def _deep_merge(base: Dict, override: Dict) -> Dict:
    """Рекурсивно объединяет override в base"""
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


def get_column_optimization_hash(col_opt: Dict[str, Any]) -> str:
    """Хэш column_optimization для ключа кэша. Пустой/None -> пустая строка."""
    if not col_opt:
        return ""
    import hashlib
    s = json.dumps(col_opt, sort_keys=True, ensure_ascii=False)
    return hashlib.md5(s.encode("utf-8")).hexdigest()[:12]


def save_config(config: Dict[str, Any]) -> bool:
    """Сохраняет конфигурацию в файл. Возвращает True при успехе."""
    config_path = _get_config_path()
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Ошибка сохранения конфига {config_path}: {e}")
        return False


# ============================================================================
# AS-BUILD КОНФИГ (отдельный файл, отдельные пути MTO и результатов)
# ============================================================================

_ASBUILD_CONFIG_FILENAME = "rfp_tags_compare_config_asbuild.json"


def get_asbuild_config_path() -> str:
    """Путь к as-build конфигу (рядом с основным конфигом)."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, _ASBUILD_CONFIG_FILENAME)


def get_default_asbuild_config() -> Dict[str, Any]:
    """Дефолтная конфигурация для as-build: те же параметры, но пути MTO и результатов — as-build."""
    base = get_default_config()
    base["paths"]["rfp_path"] = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\Сводная RFP_asbuild.xlsx"
    base["paths"]["mto_path"] = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\as-build\MTO as-build"
    base["paths"]["result_dir_base"] = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\as-build\MTO as-build_результат проверки роботом"
    base["paths"]["code_ban_file"] = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\code_ban_asbuild.txt"
    # Файлы as-build лежат прямо в корне папки (без подпапок 4-символьных кодов)
    base["step2"]["flat_mto_structure"] = True
    # Для as-build: раскраска CODE RFP vs CODE MTO не нужна
    base["step4"]["assign_rfp_mto_code_compare_colors"] = False
    # Для as-build: оставлять только title_system из MTO папки
    base["step4"]["filter_to_mto_titles"] = True
    # Упаковочные листы относятся к основному RFP, не к as-build.
    base["step4"]["include_packing_lists"] = False
    base["rfp_parts"]["auto_update_checklist"] = False
    base["rfp_parts"]["use_latest_net"] = False
    return base


def _apply_asbuild_rfp_parts_fixed(config: Dict[str, Any]) -> Dict[str, Any]:
    """Force as-build RFP-parts flags that must not follow the main profile."""
    if not isinstance(config.get("rfp_parts"), dict):
        config["rfp_parts"] = {}
    config["rfp_parts"]["auto_update_checklist"] = False
    config["rfp_parts"]["use_latest_net"] = False
    return config


@dataclass(frozen=True)
class EffectiveRfpPath:
    """Resolved RFP workbook for Step1."""

    path: str
    used_latest_net: bool
    detail: str


def resolve_effective_rfp_path(config: Dict[str, Any] | None) -> EffectiveRfpPath:
    """Choose parts net, DS/hybrid stamp, or ``paths.rfp_path`` from config.

    Args:
        config: Loaded RFP profile. ``rfp_parts.use_latest_net`` defaults to
            True (main profile). As-build must force the flag off before call;
            then this function always returns ``paths.rfp_path`` and ignores
            ``input_mode``. ``load_tags=false`` selects
            ``rfp_parts_net_no_tags.xlsx`` only in ``legacy_net``.

    Returns:
        Resolved filesystem path and a short source label.
        ``used_latest_net`` is True for stamp modes (parts / DS / hybrid).

    Raises:
        FileNotFoundError: Latest stamp workbook is requested but missing, or
            the fallback ``paths.rfp_path`` is empty. Hybrid never falls back
            to parts net.
    """
    cfg = config if isinstance(config, dict) else {}
    parts = cfg.get("rfp_parts") if isinstance(cfg.get("rfp_parts"), dict) else {}
    use_latest = bool(parts.get("use_latest_net", True))
    load_tags = bool(cfg.get("load_tags", True))
    fallback = str((cfg.get("paths") or {}).get("rfp_path") or "").strip()
    if not use_latest:
        if not fallback:
            raise FileNotFoundError(
                "paths.rfp_path пуст, а галка «Брать последний свод частей» выключена."
            )
        return EffectiveRfpPath(
            path=fallback,
            used_latest_net=False,
            detail="paths.rfp_path",
        )
    from RFQ.rfp_parts.ds_hybrid_preflight import (
        INPUT_MODE_DS_ONLY,
        INPUT_MODE_HYBRID,
        resolve_ds_baseline_xlsx,
        resolve_ds_hybrid_xlsx,
        resolve_input_mode,
    )

    input_mode = resolve_input_mode(cfg)
    if input_mode == INPUT_MODE_DS_ONLY:
        latest = resolve_ds_baseline_xlsx()
        if latest is None:
            raise FileNotFoundError(
                "Не найден «Свод ДС для запуска.xlsx» в папке "
                "«RFP сводный файл\\_ds_baseline». На вкладке "
                "RFP · Сбор частей выполните «Собрать вход Только ДС» "
                "или запустите RFP · Запуск в режиме «Свод ДС» "
                "(preflight пересоберёт свод). Либо снимите галку "
                "и укажите «Файл RFP»."
            )
        return EffectiveRfpPath(
            path=str(latest),
            used_latest_net=True,
            detail="latest Свод ДС для запуска.xlsx",
        )
    if input_mode == INPUT_MODE_HYBRID:
        latest = resolve_ds_hybrid_xlsx()
        if latest is None:
            raise FileNotFoundError(
                "Не найден «Свод ДС-RFP для запуска.xlsx» в папке "
                "«RFP сводный файл\\_ds_hybrid». Нет подстановки свода "
                "частей rfp_parts_net.xlsx. На вкладке RFP · Сбор частей "
                "выполните «Проверить RFP и наложить» или запустите "
                "RFP · Запуск в режиме «Свод ДС-RFP» (preflight "
                "пересоберёт гибрид). Либо снимите галку и укажите "
                "«Файл RFP»."
            )
        return EffectiveRfpPath(
            path=str(latest),
            used_latest_net=True,
            detail="latest Свод ДС-RFP для запуска.xlsx",
        )

    from RFQ.rfp_parts.analyze_rfp_parts import (
        resolve_latest_rfp_parts_net_no_tags_xlsx,
        resolve_latest_rfp_parts_net_xlsx,
    )

    if load_tags:
        latest = resolve_latest_rfp_parts_net_xlsx()
        label = "latest rfp_parts_net.xlsx"
        missing = (
            "Не найден последний rfp_parts_net.xlsx в папке "
            "«RFP сводный файл». Соберите свод на вкладке "
            "«RFP · Сбор частей» или снимите галку и укажите «Файл RFP»."
        )
    else:
        latest = resolve_latest_rfp_parts_net_no_tags_xlsx()
        label = "latest rfp_parts_net_no_tags.xlsx"
        missing = (
            "Не найден последний rfp_parts_net_no_tags.xlsx в папке "
            "«RFP сводный файл». Запустите RFP со снятой галкой тегов "
            "(preflight пересоберёт свод) или соберите части на вкладке "
            "«RFP · Сбор частей»."
        )
    if latest is None:
        raise FileNotFoundError(missing)
    return EffectiveRfpPath(
        path=str(latest),
        used_latest_net=True,
        detail=label,
    )


def load_asbuild_config() -> Dict[str, Any]:
    """Загружает as-build конфигурацию. При отсутствии — возвращает дефолт с as-build путями."""
    config_path = get_asbuild_config_path()
    default = get_default_asbuild_config()
    if not os.path.exists(config_path):
        return default
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        merged = _deep_merge(default, loaded)
        return _apply_asbuild_rfp_parts_fixed(merged)
    except Exception as e:
        print(f"Ошибка загрузки as-build конфига {config_path}: {e}")
        return default


def save_asbuild_config(config: Dict[str, Any]) -> bool:
    """Сохраняет as-build конфигурацию. Возвращает True при успехе."""
    config_path = get_asbuild_config_path()
    try:
        normalized = copy.deepcopy(config)
        _apply_asbuild_rfp_parts_fixed(normalized)
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(normalized, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Ошибка сохранения as-build конфига {config_path}: {e}")
        return False


_STEP4_FILE_PREFIX = "Шаг4_Сопоставление_RFP_MTO_"
_STEP4_FILE_SUFFIX = ".xlsx"


def find_latest_step4_result_file() -> Optional[str]:
    """
    Ищет последний по дате/времени файл Шаг4_Сопоставление_RFP_MTO_*.xlsx
    во всех папках результатов (основной и as-build конфиги).
    Возвращает полный путь к файлу или None, если не найден.
    """
    candidates: List[Tuple[float, str]] = []

    for config_fn in (load_config, load_asbuild_config):
        try:
            cfg = config_fn()
            base = cfg.get("paths", {}).get("result_dir_base", "")
            if not base or not os.path.isdir(base):
                continue
            for entry in os.scandir(base):
                if not entry.is_dir():
                    continue
                folder = entry.path
                for f in os.listdir(folder):
                    if f.startswith(_STEP4_FILE_PREFIX) and f.endswith(_STEP4_FILE_SUFFIX):
                        fp = os.path.join(folder, f)
                        try:
                            mtime = os.path.getmtime(fp)
                            candidates.append((mtime, fp))
                        except OSError:
                            pass
        except Exception:
            pass

    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


_BCC_MATRIX_FILE_PREFIX = "Шаг4_Матрица_BCC_"


def find_latest_bcc_matrix_file() -> Optional[str]:
    """
    Ищет последний по дате/времени файл Шаг4_Матрица_BCC_*.xlsx
    во всех папках результатов (основной и as-build конфиги).
    Возвращает полный путь к файлу или None, если не найден.
    """
    candidates: List[Tuple[float, str]] = []

    for config_fn in (load_config, load_asbuild_config):
        try:
            cfg = config_fn()
            base = cfg.get("paths", {}).get("result_dir_base", "")
            if not base or not os.path.isdir(base):
                continue
            for entry in os.scandir(base):
                if not entry.is_dir():
                    continue
                folder = entry.path
                for f in os.listdir(folder):
                    if f.startswith(_BCC_MATRIX_FILE_PREFIX) and f.endswith(
                        _STEP4_FILE_SUFFIX
                    ):
                        fp = os.path.join(folder, f)
                        try:
                            mtime = os.path.getmtime(fp)
                            candidates.append((mtime, fp))
                        except OSError:
                            pass
        except Exception:
            pass

    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def get_result_dir_path(rfp_path: str) -> str:
    """
    Создает путь для папки результатов на основе пути к RFP файлу
    
    Args:
        rfp_path: Путь к RFP файлу
        
    Returns:
        Путь к папке результатов в формате: <директория_файла>_результат_проверки_ГГГГ.ММ.ДД.ЧЧ.ММ
    """
    # Получаем директорию файла
    if os.path.isfile(rfp_path):
        base_dir = os.path.dirname(rfp_path)
    else:
        base_dir = rfp_path
    
    # Формируем дату в формате ГГГГ.ММ.ДД.ЧЧ.ММ
    now = datetime.now()
    date_str = now.strftime("%Y.%m.%d.%H.%M")
    
    # Формируем имя папки
    result_dir_name = f"_результат_проверки_{date_str}"
    result_path = os.path.join(base_dir, result_dir_name)
    
    return result_path


def ensure_result_dir_exists(result_dir: str) -> str:
    """
    Проверяет наличие папки результатов и создает её при необходимости
    
    Args:
        result_dir: Путь к папке результатов
        
    Returns:
        Путь к папке результатов (нормализованный)
    """
    result_dir = os.path.normpath(result_dir)
    if not os.path.exists(result_dir):
        make_dir(result_dir)
        print(f"Создана папка результатов: {result_dir}")
    return result_dir


_timing_log_enabled: bool = True


def set_timing_log_enabled(enabled: bool) -> None:
    """Включает/выключает запись в timing_log.xlsx (глобально для всего процесса)."""
    global _timing_log_enabled
    _timing_log_enabled = enabled


def set_memory_log_enabled(enabled: bool) -> None:
    """Включает/выключает запись в memory_log.txt (глобально для всего процесса)."""
    global _memory_log_enabled
    _memory_log_enabled = enabled


def append_memory_log(result_dir: str, label: str) -> None:
    """
    Добавляет строку в лог использования памяти (memory_log.txt).
    Использует только psutil (без tracemalloc — без накладных расходов на аллокации).
    
    Args:
        result_dir: Путь к папке результатов
        label: Метка точки замера (например "after_step1_load_rfp")
    """
    if not _memory_log_enabled or not result_dir:
        return
    try:
        rss_mb = ""
        system_mem = ""
        try:
            import psutil
            proc = psutil.Process()
            rss_mb = f" rss={proc.memory_info().rss / 1024 / 1024:.1f}MB"
            vm = psutil.virtual_memory()
            total_mb = vm.total / 1024 / 1024
            available_mb = vm.available / 1024 / 1024
            used_pct = vm.percent
            system_mem = f" total={total_mb:.0f}MB available={available_mb:.0f}MB used_pct={used_pct:.1f}%"
        except ImportError:
            pass
        ensure_result_dir_exists(result_dir)
        log_path = os.path.join(result_dir, "memory_log.txt")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(f"[{timestamp}] {label}:{rss_mb}{system_mem}\n")
    except Exception as e:
        print(f"Ошибка записи memory_log: {e}")


def append_memory_top_stats(result_dir: str, label: str, top_n: int = 15) -> None:
    """
    Записывает топ-N аллокаций памяти (tracemalloc) в memory_log.txt.
    
    Args:
        result_dir: Путь к папке результатов
        label: Метка точки замера (например "step4_after_save_match_result_to_excel")
        top_n: Количество строк в отчёте
    """
    if not _memory_log_enabled or not result_dir:
        return
    global _tracemalloc_started
    try:
        if not _tracemalloc_started:
            tracemalloc.start()
            _tracemalloc_started = True
        snapshot = tracemalloc.take_snapshot()
        tracemalloc.stop()
        _tracemalloc_started = False
        top_stats = snapshot.statistics('lineno')
        ensure_result_dir_exists(result_dir)
        log_path = os.path.join(result_dir, "memory_log.txt")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(f"[{timestamp}] {label} Top {top_n} allocations:\n")
            for i, stat in enumerate(top_stats[:top_n], 1):
                size_kib = stat.size / 1024
                frame = stat.traceback[0] if stat.traceback else None
                loc = f"{frame.filename}:{frame.lineno}" if frame else "?"
                log_file.write(f"  {i}. {loc}: {size_kib:.1f} KiB, {stat.count} blocks\n")
            log_file.write("\n")
        del snapshot, top_stats
    except Exception as e:
        print(f"Ошибка записи memory_top_stats: {e}")


def _estimate_size_by_sampling(obj: Any, use_pympler: bool, asizeof_module: Any) -> Tuple[int, int, Any]:
    """
    Оценка размера через выборку (до 50 элементов) — быстрее полного asizeof.
    Возвращает (size_bytes, rows, extra) или (0, 0, False) при ошибке.
    """
    try:
        if isinstance(obj, (list, tuple)):
            rows = len(obj)
            if rows == 0:
                return 0, 0, False
            sample_size = min(50, rows)
            step = max(1, rows // sample_size)
            sample = [obj[i] for i in range(0, rows, step)][:sample_size]
            if use_pympler and asizeof_module:
                sample_bytes = asizeof_module.asizeof(sample)
                size_bytes = int(sample_bytes * rows / len(sample))
            else:
                size_bytes = sys.getsizeof(obj)
            return size_bytes, rows, sample_size < rows
        elif isinstance(obj, dict):
            values = [v for v in obj.values() if isinstance(v, (list, tuple))]
            total_rows = sum(len(v) for v in values)
            if total_rows == 0:
                sz = asizeof_module.asizeof(obj) if (use_pympler and asizeof_module) else sys.getsizeof(obj)
                return sz, 0, False
            collected = []
            for v in values:
                collected.extend(v)
                if len(collected) >= 50:
                    break
            sample = collected[:50]
            if use_pympler and asizeof_module and sample:
                sample_bytes = asizeof_module.asizeof(sample)
                size_bytes = int(sample_bytes * total_rows / len(sample)) + sys.getsizeof(obj)
            else:
                size_bytes = sys.getsizeof(obj)
            return size_bytes, total_rows, len(values)
        else:
            sz = asizeof_module.asizeof(obj) if (use_pympler and asizeof_module) else sys.getsizeof(obj)
            return sz, 0, False
    except Exception:
        return 0, 0, False


def log_data_structures_memory(result_dir: str, label: str, **structures: Any) -> None:
    """
    Записывает в memory_log.txt размер переданных структур данных и системную память.
    Использует выборку (до 50 элементов) вместо полного обхода — для ускорения.
    
    Args:
        result_dir: Путь к папке результатов
        label: Метка точки замера (например "after_step1_load_rfp")
        **structures: Именованные структуры: rfp_data=..., mto_data=..., vo_data=..., result_rows=...
    """
    if not _memory_log_enabled or not result_dir:
        return
    try:
        ensure_result_dir_exists(result_dir)
        log_path = os.path.join(result_dir, "memory_log.txt")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lines = [f"[{timestamp}] {label}:"]
        try:
            import psutil
            vm = psutil.virtual_memory()
            total_mb = vm.total / 1024 / 1024
            available_mb = vm.available / 1024 / 1024
            used_pct = vm.percent
            lines.append(f"  System: total={total_mb:.0f} MB, available={available_mb:.0f} MB, used={used_pct:.1f}%")
        except ImportError:
            pass
        use_pympler = True
        asizeof_module = None
        try:
            from pympler import asizeof as asizeof_module
        except ImportError:
            use_pympler = False
        for name, obj in structures.items():
            if obj is None:
                continue
            try:
                size_bytes, rows, extra = _estimate_size_by_sampling(obj, use_pympler, asizeof_module)
                size_mb = size_bytes / 1024 / 1024
                if isinstance(obj, (list, tuple)):
                    lines.append(f"  {name}: {rows} rows, {size_mb:.1f} MB")
                elif isinstance(obj, dict):
                    keys = len(obj)
                    lines.append(f"  {name}: {keys} title_systems, {rows} rows, {size_mb:.1f} MB")
                else:
                    lines.append(f"  {name}: {size_mb:.1f} MB")
            except Exception as e:
                lines.append(f"  {name}: (error: {e})")
        if not use_pympler:
            lines.append("  (pympler not installed, sizes are approximate)")
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write("\n".join(lines) + "\n\n")
    except Exception as e:
        print(f"Ошибка записи log_data_structures_memory: {e}")


def log_mto_tag_distribution(result_dir: str, mto_data: Dict[str, List[Any]], label: str = "MTO tag distribution") -> None:
    """
    Логирует распределение тегов по строкам MTO (до раскрытия).
    Сколько строк с 0, 1, 2, 3-9, 10+ тегами; строк без тегов с VALUES>1.
    
    Args:
        result_dir: Путь к папке результатов
        mto_data: Словарь title_system -> список RowStd
        label: Метка для лога
    """
    if not _memory_log_enabled or not result_dir or not mto_data:
        return
    try:
        from base.base_classes import RowType
        dist = {0: 0, 1: 0, 2: 0, "3-9": 0, "10+": 0}
        no_tags_values_gt1 = 0
        for rows in mto_data.values():
            if not isinstance(rows, (list, tuple)):
                continue
            for row in rows:
                if getattr(row, "row_type", None) != RowType.position_row:
                    continue
                tags_list = row.get_tags_list() if hasattr(row, "get_tags_list") else []
                n = len(tags_list) if tags_list else 0
                if n == 0:
                    dist[0] += 1
                    try:
                        from base.tables_columns import VALUES
                        v = row.get_value(VALUES) if hasattr(row, "get_value") else 0
                        if v is not None and int(float(v)) > 1:
                            no_tags_values_gt1 += 1
                    except (ValueError, TypeError, KeyError, ImportError):
                        pass
                elif n == 1:
                    dist[1] += 1
                elif n == 2:
                    dist[2] += 1
                elif n < 10:
                    dist["3-9"] += 1
                else:
                    dist["10+"] += 1
        ensure_result_dir_exists(result_dir)
        log_path = os.path.join(result_dir, "memory_log.txt")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        total = sum(dist.values())
        lines = [
            f"[{timestamp}] {label}:",
            f"  Rows with 0 tags: {dist[0]} (of these, VALUES>1 will split: {no_tags_values_gt1})",
            f"  Rows with 1 tag: {dist[1]}",
            f"  Rows with 2 tags: {dist[2]}",
            f"  Rows with 3-9 tags: {dist['3-9']}",
            f"  Rows with 10+ tags: {dist['10+']}",
            f"  Total position rows: {total}",
        ]
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write("\n".join(lines) + "\n\n")
    except Exception as e:
        print(f"Ошибка записи log_mto_tag_distribution: {e}")


def append_timing_log(result_dir: str, message: str) -> None:
    """Буферизует запись timing_log в памяти. Запись на диск -- в finalize_timing_log()."""
    global _timing_buffer_result_dir
    if not _timing_log_enabled or not result_dir:
        return
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    step_name, duration_val = _parse_timing_message(message)
    _timing_buffer.append((timestamp, step_name, duration_val, message))
    if _timing_buffer_result_dir is None:
        _timing_buffer_result_dir = result_dir


def save_input_fingerprint(
    result_dir: str,
    source_name: str,
    base_path: str,
    file_paths: List[str],
) -> None:
    """
    Сохраняет легковесный fingerprint входных файлов (mtime/size/path) в JSON.
    Нужен для подтверждения, что разные прогоны сделаны на одном и том же наборе данных.
    """
    if not result_dir or not source_name:
        return
    try:
        ensure_result_dir_exists(result_dir)
        normalized_base = os.path.normpath(base_path or "")
        file_records = []
        hash_chunks = []
        existing_count = 0
        missing_count = 0
        total_size = 0
        latest_mtime = 0.0

        for raw_path in sorted(set(file_paths or [])):
            path = os.path.normpath(raw_path)
            if not os.path.exists(path):
                missing_count += 1
                continue
            existing_count += 1
            size = os.path.getsize(path)
            mtime = os.path.getmtime(path)
            total_size += size
            if mtime > latest_mtime:
                latest_mtime = mtime

            try:
                rel_path = os.path.relpath(path, normalized_base) if normalized_base else path
            except Exception:
                rel_path = path
            rel_path = rel_path.replace("\\", "/")

            mtime_token = f"{mtime:.6f}"
            hash_chunks.append(f"{rel_path}|{size}|{mtime_token}")
            file_records.append({"path": rel_path, "size": size, "mtime": mtime_token})

        fingerprint_source = "\n".join(hash_chunks)
        fingerprint = hashlib.md5(fingerprint_source.encode("utf-8", errors="replace")).hexdigest() if hash_chunks else ""

        payload = {
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "base_path": normalized_base,
            "files_total": len(set(file_paths or [])),
            "files_existing": existing_count,
            "files_missing": missing_count,
            "total_size_bytes": total_size,
            "latest_mtime": f"{latest_mtime:.6f}" if latest_mtime else "",
            "fingerprint_md5": fingerprint,
            "files": file_records,
        }

        out_path = os.path.join(result_dir, "input_fingerprints.json")
        full_payload: Dict[str, Any] = {}
        if os.path.exists(out_path):
            try:
                with open(out_path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    if isinstance(loaded, dict):
                        full_payload = loaded
            except Exception:
                full_payload = {}
        full_payload[source_name] = payload
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(full_payload, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Ошибка записи input_fingerprint ({source_name}): {e}")


def finalize_timing_log(result_dir: str, top_n: int = 5) -> None:
    """Flush buffered timing entries to Excel and build top_longest summary sheet."""
    global _timing_buffer_result_dir
    if not _timing_log_enabled:
        _timing_buffer.clear()
        _timing_buffer_result_dir = None
        return
    effective_dir = result_dir or _timing_buffer_result_dir
    if not effective_dir or not _timing_buffer:
        _timing_buffer.clear()
        _timing_buffer_result_dir = None
        return
    try:
        ensure_result_dir_exists(effective_dir)
        log_path = os.path.join(effective_dir, _TIMING_LOG_FILENAME)

        wb = Workbook()
        ws = wb.active
        ws.title = _TIMING_SHEET_NAME
        ws.append(["timestamp", "step", "duration_sec", "message"])

        durations: List[Tuple[str, float]] = []
        for ts, step_name, duration_val, message in _timing_buffer:
            ws.append([ts, step_name, duration_val, message])
            if duration_val is not None:
                ws.cell(row=ws.max_row, column=3).number_format = "0.000"
                durations.append((step_name, duration_val))
        ws.auto_filter.ref = f"A1:D{ws.max_row}"
        _autofit_worksheet_columns(ws, max_col=4)

        if durations:
            durations.sort(key=lambda x: x[1], reverse=True)
            top_items = durations[:max(1, top_n)]
            ws_top = wb.create_sheet(_TIMING_TOP_SHEET_NAME)
            ws_top.append(["generated_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "", ""])
            ws_top.append(["rank", "step", "duration_sec", ""])
            for idx, (name, duration) in enumerate(top_items, start=1):
                ws_top.append([idx, name, duration, ""])
                ws_top.cell(row=ws_top.max_row, column=3).number_format = "0.000"
            ws_top.auto_filter.ref = f"A2:C{ws_top.max_row}"
            _autofit_worksheet_columns(ws_top, max_col=3)

        wb.save(log_path)
    except Exception as e:
        print(f"Ошибка записи timing_log: {e}")
    finally:
        _timing_buffer.clear()
        _timing_buffer_result_dir = None


def _parse_timing_message(message: str) -> Tuple[str, Optional[float]]:
    """
    Парсит строку формата "<step>: <seconds>s" и возвращает (step, duration_sec).
    Если формат не распознан, возвращает исходное сообщение и None.
    """
    # Берем последнее двоеточие перед "<число>s", чтобы корректно работать со step4::...
    # Примеры:
    #   "step4::check_mto_data: 58.711s"
    #   "save_match_result_to_excel::fill: 4.762s"
    m = re.match(r"^\s*(.+?)\s*:\s*([0-9]+(?:\.[0-9]+)?)\s*s\b.*$", message)
    if m:
        step_name = m.group(1).strip()
        try:
            return step_name, float(m.group(2))
        except Exception:
            return step_name, None
    return message, None


def _autofit_worksheet_columns(ws, max_col: int, min_width: int = 10, max_width: int = 80) -> None:
    """
    Подбирает ширину колонок по содержимому в разумных пределах.
    """
    for col_idx in range(1, max_col + 1):
        max_len = 0
        for row in ws.iter_rows(min_row=1, min_col=col_idx, max_col=col_idx, values_only=True):
            value = row[0]
            if value is None:
                continue
            value_len = len(str(value))
            if value_len > max_len:
                max_len = value_len
        adjusted_width = min(max(min_width, max_len + 2), max_width)
        ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width


if __name__ in {"__main__"}:
    # Тестирование утилит
    test_path = r"C:\test\file.xlsx"
    result_path = get_result_dir_path(test_path)
    print(f"Тестовый путь результата: {result_path}")
    
    # Тест создания директории
    test_dir = ensure_result_dir_exists(result_path)
    print(f"Директория создана/проверена: {test_dir}")

