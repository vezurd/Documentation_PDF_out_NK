"""Persisted palette for pipeline and F-stage status colors.

JSON lives under the injected catalog runtime directory (never UNC).
Unknown keys resolve to :data:`UNKNOWN_STATUS_COLOR`.
"""

from __future__ import annotations

import json
import re
from collections.abc import Mapping
from pathlib import Path
from typing import Any

STATUS_COLORS_FILENAME = "status_colors.json"
STATUS_COLORS_VERSION = 1
UNKNOWN_STATUS_COLOR = "#5f6368"

_HEX_RE = re.compile(r"^#?[0-9A-Fa-f]{6}$")

# Pipeline statuses first, then optional F-stage keys not already listed.
_PIPELINE_DEFAULTS: dict[str, str] = {
    "not_uploaded": "#80868b",
    "sent_tdo": "#f9ab00",
    "tdo_review": "#1a73e8",
    "code_a": "#137333",
    "code_b": "#e37400",
    "code_c": "#c5221f",
    "approval_stale": "#80868b",
    "rd_missing_transfer": "#c4c7c5",
    "mixed_titles": "#f3c2c2",
}

_F_STAGE_DEFAULTS: dict[str, str] = {
    "sr_upload": "#f9ab00",
    "tdo_sent": "#f9ab00",
    "incoming_sent": "#f9ab00",
    "tdo_passed": "#1a73e8",
    "incoming_passed": "#1a73e8",
    "dup": "#5f6368",
    "correction": "#5f6368",
    "agreed": "#137333",
    "other": "#5f6368",
}

_HEATMAP_DEFAULTS: dict[str, str] = {
    "empty": "#e8eaed",
    "no_mto": "#9aa0a6",
    "working": "#8ab4f8",
    "annulled": "#b39ddb",
    "problem": "#c5221f",
    "current": "#137333",
    "us_build": "#5f259f",
}

# Loud red for «нет файла»: that row must not be missed in the heatmap.
_EXPORT_DEFAULTS: dict[str, str] = {
    "export_add": "#1a73e8",
    "export_replace": "#e37400",
    "export_same": "#137333",
    "export_missing": "#ff1744",
    "export_unknown": "#9aa0a6",
    "export_pin_stale": "#a142f4",
}

STATUS_COLOR_LABELS: dict[str, str] = {
    "not_uploaded": "Не загружен в СР",
    "sent_tdo": "Отправлен на ТДО",
    "tdo_review": "Прошел ТДО",
    "code_a": "Получен код A",
    "code_b": "Получен код B",
    "code_c": "Получен код C",
    "approval_stale": "письмо не этого цикла",
    "rd_missing_transfer": "Нет в РД",
    "mixed_titles": "Смешанные титулы",
    "sr_upload": "загрузка в СР",
    "tdo_sent": "отпр. на ТДО",
    "incoming_sent": "отпр. на входной контроль",
    "tdo_passed": "прошла ТДО",
    "incoming_passed": "входной контроль",
    "dup": "ДУП",
    "correction": "корректировка",
    "agreed": "согласовано",
    "other": "прочее",
    "empty": "нет ревизии",
    "no_mto": "нет MTO",
    "working": "рабочая",
    "annulled": "аннулирована",
    "problem": "проблема",
    "current": "текущий состав",
    "us_build": "US-BUILD / as-build",
    "export_add": "добавится в папку робота",
    "export_replace": "заменится в папке робота",
    "export_same": "данные совпадают",
    "export_missing": "нет файла MTO",
    "export_unknown": "не сравнено",
    "export_pin_stale": "ручной выбор устарел",
}

STATUS_SHORT_LABELS: dict[str, str] = {
    "not_uploaded": "не в СР",
    "sent_tdo": "на ТДО",
    "tdo_review": "ТДО",
    "code_a": "код A",
    "code_b": "код B",
    "code_c": "код C",
    "no_mto": "нет MTO",
    "working": "рабочая",
    "annulled": "аннул.",
    "empty": "нет данных",
    "rd_missing_transfer": "Нет в РД",
}


def default_palette() -> dict[str, str]:
    """Return a copy of the factory pipeline + F-stage + heatmap palette.

    Returns:
        Key to ``#rrggbb`` map in display order.
    """

    palette = dict(_PIPELINE_DEFAULTS)
    for key, color in _F_STAGE_DEFAULTS.items():
        palette.setdefault(key, color)
    for key, color in _HEATMAP_DEFAULTS.items():
        palette.setdefault(key, color)
    for key, color in _EXPORT_DEFAULTS.items():
        palette.setdefault(key, color)
    return palette


def normalize_hex_color(value: Any) -> str | None:
    """Return a ``#rrggbb`` color, or ``None`` when the text is not a hex color.

    Args:
        value: User or JSON color text.

    Returns:
        Lower-case ``#rrggbb``, or ``None``.
    """

    text = str(value or "").strip()
    if not _HEX_RE.fullmatch(text):
        return None
    if not text.startswith("#"):
        text = f"#{text}"
    return text.casefold()


def color_for(palette: Mapping[str, str], key: str) -> str:
    """Return the palette color for ``key``, or the unknown fallback.

    Args:
        palette: Loaded or factory palette.
        key: Pipeline status or F-stage id.

    Returns:
        ``#rrggbb`` string.
    """

    normalized = normalize_hex_color(palette.get(key, ""))
    if normalized:
        return normalized
    return UNKNOWN_STATUS_COLOR


def status_color_label(key: str) -> str:
    """Return the Russian dialog label for a palette key.

    Args:
        key: Pipeline status or F-stage id.

    Returns:
        Localized label, or the raw key when unknown.
    """

    return STATUS_COLOR_LABELS.get(key, key)


def status_short_label(key: str) -> str:
    """Return a compact Russian badge for a palette key.

    Args:
        key: Pipeline status or heatmap fill key.

    Returns:
        Short label, or :func:`status_color_label` when no short form exists.
    """

    return STATUS_SHORT_LABELS.get(key) or status_color_label(key)


def load_status_colors(path: str | Path) -> dict[str, str]:
    """Load a palette from JSON, overlaying factory defaults.

    Missing or corrupt files yield the factory palette. Extra keys with valid
    hex colors are kept (optional F stages).

    Args:
        path: Target JSON file under the catalog runtime directory.

    Returns:
        Merged key to ``#rrggbb`` map.
    """

    palette = default_palette()
    target = Path(path)
    if not target.is_file():
        return palette
    try:
        raw = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return palette
    items: Any
    if isinstance(raw, dict) and isinstance(raw.get("colors"), dict):
        items = raw["colors"]
    elif isinstance(raw, dict):
        items = {
            key: value
            for key, value in raw.items()
            if key not in {"version", "colors"}
        }
    else:
        return palette
    if not isinstance(items, dict):
        return palette
    for key, value in items.items():
        name = str(key or "").strip()
        color = normalize_hex_color(value)
        if name and color:
            palette[name] = color
    return palette


def save_status_colors(path: str | Path, palette: Mapping[str, str]) -> None:
    """Write a palette atomically as UTF-8 JSON.

    Args:
        path: Target JSON file under the catalog runtime directory.
        palette: Key to hex-color map.

    Raises:
        OSError: If the file cannot be written.
    """

    colors: dict[str, str] = {}
    for key, value in palette.items():
        name = str(key or "").strip()
        color = normalize_hex_color(value)
        if name and color:
            colors[name] = color
    payload = {
        "version": STATUS_COLORS_VERSION,
        "colors": colors,
    }
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    tmp_path = target.with_name(target.name + ".tmp")
    tmp_path.write_text(text, encoding="utf-8")
    tmp_path.replace(target)


def status_colors_path(runtime_dir: str | Path) -> Path:
    """Return ``status_colors.json`` under the catalog runtime directory.

    Args:
        runtime_dir: Injected catalog runtime directory.

    Returns:
        JSON path (not created).
    """

    return Path(runtime_dir) / STATUS_COLORS_FILENAME
