"""
Конфигурация для BBB Analysis.
Загрузка/сохранение JSON, значения по умолчанию.

Секция section_types вынесена в отдельный Excel файл на сетевом диске:
  base/bbb_section_types_excel.py
"""

import json
import os
from typing import Any, Dict

_CONFIG_FILENAME = "bbb_analysis_config.json"


def _get_config_path() -> str:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, _CONFIG_FILENAME)


def get_default_config() -> Dict[str, Any]:
    return {
        "folder_rules": {
            "search_only_in_dwg": True,
        },
        "mto_vs_code_base": {
            "enabled": True,
            "check_by_code": True,
            "check_equipment_codes": True,
            "check_tags_value": True,
            "check_value": True,
            "check_duplicate_tags": True,
            "check_mass": True,
            "check_prohibition": False,
            "check_tags_4_2_4_4": True,
            "check_position_numeration": True,
        },
        "mto_vs_code_base_correction": {
            "enabled": True,
            "check_by_code": True,
            "check_tags_value": True,
            "check_value": True,
            "check_duplicate_tags": True,
            "check_tags_4_2_4_4": True,
            "check_position_numeration": True,
        },
        "bbb_vs_code_base": {
            "enabled": True,
            "check_by_code": True,
            "check_equipment_codes": True,
            "check_tags_value": True,
            "check_value": True,
            "check_duplicate_tags": True,
            "check_mass": True,
            "check_prohibition": False,
            "check_work_code_and_mtr_group": True,
        },
        "bbb_vs_mto": {
            "enabled": True,
            "compare_fields": True,
            "compare_values": True,
            "compare_tags": True,
            "compare_title_marka_revision": True,
            "compare_bom_annotation_with_mto": True,
            "check_missing_mto_codes": True,
        },
    }


def _deep_merge(base: dict, override: dict) -> dict:
    result = base.copy()
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


def get_section_target(section_types_cfg: dict, section_name: str) -> str:
    """Возвращает целевой документ (BOE/BOM) для секции MTO."""
    entry = section_types_cfg.get(section_name, {})
    return entry.get("target", "BOE") if isinstance(entry, dict) else "BOE"


def load_config() -> Dict[str, Any]:
    config_path = _get_config_path()
    default = get_default_config()
    if not os.path.exists(config_path):
        return default
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        return _deep_merge(default, loaded)
    except Exception as e:
        print(f"Ошибка загрузки конфига BBB {config_path}: {e}")
        return default


def save_config(config: Dict[str, Any]) -> bool:
    config_path = _get_config_path()
    # Убираем section_types из JSON — они хранятся в отдельном Excel файле
    config_to_save = {k: v for k, v in config.items() if k != "section_types"}
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config_to_save, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Ошибка сохранения конфига BBB {config_path}: {e}")
        return False
