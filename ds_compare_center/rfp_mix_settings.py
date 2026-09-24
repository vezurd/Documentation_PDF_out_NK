"""Qt-free helpers for RFP collect/launch mix-mode JSON keys.

``rfp_parts.collect_mix_mode`` belongs to the parts tab.
``rfp_parts.launch_mix_mode`` belongs to the launch tab.
The two keys are independent: writing one must not change the other.
``rfp_parts.input_mode`` is the launch file picker and is not used here.
"""

from __future__ import annotations

from typing import Any, Callable

from RFQ.rfp_parts.ds_hybrid_preflight import (
    INPUT_MODE_DS_ONLY,
    INPUT_MODE_HYBRID,
    INPUT_MODE_LEGACY_NET,
    resolve_launch_mix_mode,
)
from RFQ.rfp_parts.ds_rfp_tag_placement import MIX_MIXED, MIX_SEPARATE

KEY_COLLECT_MIX_MODE = "collect_mix_mode"
KEY_LAUNCH_MIX_MODE = "launch_mix_mode"
KEY_COLLECT_JOB = "collect_job"

COLLECT_JOB_PARTS = "parts"
COLLECT_JOB_DS_ONLY = "ds_only"
COLLECT_JOB_HYBRID = "hybrid"
ALLOWED_COLLECT_JOBS: frozenset[str] = frozenset(
    {COLLECT_JOB_PARTS, COLLECT_JOB_DS_ONLY, COLLECT_JOB_HYBRID}
)

COLLECT_JOB_CHOICES: tuple[tuple[str, str, str], ...] = (
    (
        COLLECT_JOB_HYBRID,
        "Свод из ДС и наложить RFP",
        "Сначала аудит папки ДС, затем корень RFP_Зиновьев без подпапок. "
        "Пишет «Свод ДС-RFP для запуска.xlsx» в _ds_hybrid. Группа целиком "
        "из RFP только если совпали код, единица и количество. Иначе группа "
        "остаётся из ДС.",
    ),
    (
        COLLECT_JOB_DS_ONLY,
        "Свод только из ДС",
        "Аудит папки ДС без RFP. Пишет «Свод ДС для запуска.xlsx» в "
        "_ds_baseline. При ошибке раскладки свод не создаётся.",
    ),
    (
        COLLECT_JOB_PARTS,
        "Свод частей RFP",
        "Прежний сбор rfp_parts_net.xlsx. Контур ДС и реестр не трогает.",
    ),
)
MIX_CHOICES: tuple[tuple[str, str, str], ...] = (
    (
        MIX_SEPARATE,
        "Без смешения",
        "Теги RFP кластера садятся на один ДС; ярлык остаётся своим номером.",
    ),
    (
        MIX_MIXED,
        "Со смешением",
        "Все строки кластера получают общий ярлык вида ДС15/61.",
    ),
)

_JOB_BUTTON_SHORT = {
    COLLECT_JOB_PARTS: "части RFP",
    COLLECT_JOB_DS_ONLY: "только ДС",
    COLLECT_JOB_HYBRID: "ДС+RFP",
}
_JOB_TITLE_SHORT = {
    COLLECT_JOB_PARTS: "части RFP",
    COLLECT_JOB_DS_ONLY: "только ДС",
    COLLECT_JOB_HYBRID: "ДС+RFP",
}
_MIX_TITLE = {mode: label for mode, label, _hint in MIX_CHOICES}
_MIX_BUTTON_SHORT = {
    MIX_SEPARATE: "без смешения",
    MIX_MIXED: "со смешением",
}
_LAUNCH_SOURCE_SHORT = {
    INPUT_MODE_LEGACY_NET: "части RFP",
    INPUT_MODE_DS_ONLY: "только ДС",
    INPUT_MODE_HYBRID: "ДС+RFP",
}


def resolve_mix_mode(raw: object | None) -> str:
    """Return ``separate`` or ``mixed``; unknown/missing values are separate."""
    text = str(raw or "").strip()
    if text in (MIX_SEPARATE, MIX_MIXED):
        return text
    return MIX_SEPARATE


def resolve_collect_job(raw: object | None) -> str:
    """Return a collect job id; unknown/missing values are hybrid."""
    text = str(raw or "").strip()
    if text in ALLOWED_COLLECT_JOBS:
        return text
    return COLLECT_JOB_HYBRID


def rfp_parts_section(config: dict[str, Any] | None) -> dict[str, Any]:
    """Return the ``rfp_parts`` dict from ``config``, or an empty dict."""
    if not isinstance(config, dict):
        return {}
    parts = config.get("rfp_parts")
    return parts if isinstance(parts, dict) else {}


def read_collect_mix_mode(config: dict[str, Any] | None) -> str:
    """Read ``rfp_parts.collect_mix_mode``; default ``separate``."""
    return resolve_mix_mode(rfp_parts_section(config).get(KEY_COLLECT_MIX_MODE))


def read_launch_mix_mode(config: dict[str, Any] | None) -> str:
    """Read ``rfp_parts.launch_mix_mode``; default ``separate``.

    Does not look at ``collect_mix_mode``.
    """
    return resolve_launch_mix_mode(config)


def read_collect_job(config: dict[str, Any] | None) -> str:
    """Read ``rfp_parts.collect_job``; default ``hybrid``."""
    return resolve_collect_job(rfp_parts_section(config).get(KEY_COLLECT_JOB))


def set_rfp_parts_key(
    config: dict[str, Any],
    key: str,
    value: str,
) -> dict[str, Any]:
    """Set one ``rfp_parts`` key in ``config`` without touching sibling keys.

    Args:
        config: Full RFP profile dict (mutated in place).
        key: Name under ``rfp_parts`` (for example ``collect_mix_mode``).
        value: Value to store.

    Returns:
        The same ``config`` object.
    """
    parts = config.get("rfp_parts")
    if not isinstance(parts, dict):
        parts = {}
        config["rfp_parts"] = parts
    parts[key] = value
    return config


def persist_rfp_parts_key(
    key: str,
    value: str,
    *,
    load_config_fn: Callable[[], dict[str, Any]] | None = None,
    save_config_fn: Callable[[dict[str, Any]], bool] | None = None,
) -> bool:
    """Load the RFP JSON, write one ``rfp_parts`` key, save.

    Args:
        key: Name under ``rfp_parts``.
        value: Value to store.
        load_config_fn: Optional loader (tests). Default Main RFP JSON.
        save_config_fn: Optional saver (tests).

    Returns:
        True when the saver reports success.
    """
    if load_config_fn is None or save_config_fn is None:
        from RFQ.tags_rfp_compare.rfp_tags_utils import load_config, save_config

        loader = load_config_fn or load_config
        saver = save_config_fn or save_config
    else:
        loader = load_config_fn
        saver = save_config_fn
    config = loader()
    if not isinstance(config, dict):
        config = {}
    set_rfp_parts_key(config, key, value)
    return bool(saver(config))


def mix_suffix_ru(mix: str) -> str:
    """Short Russian mix label used in parentheses, without the parens."""
    return _MIX_BUTTON_SHORT[resolve_mix_mode(mix)]


def format_collect_run_button(job: str, mix: str) -> str:
    """Label for the single parts-tab run button.

    Mix is appended only when the job is hybrid.

    Args:
        job: ``parts`` / ``ds_only`` / ``hybrid``.
        mix: ``separate`` / ``mixed``.

    Returns:
        For example ``Запуск (ДС+RFP · без смешения)``.
    """
    resolved_job = resolve_collect_job(job)
    short = _JOB_BUTTON_SHORT[resolved_job]
    if resolved_job != COLLECT_JOB_HYBRID:
        return f"Запуск ({short})"
    return f"Запуск ({short} · {mix_suffix_ru(mix)})"


def format_collect_block_title(job: str, mix: str) -> str:
    """Collapsed header showing the current collect job and mix choice."""
    job_label = _JOB_TITLE_SHORT[resolve_collect_job(job)]
    mix_label = _MIX_TITLE[resolve_mix_mode(mix)]
    return f"Сбор: {job_label} · Посадка RFP: {mix_label}"


def format_mix_block_title(mix: str) -> str:
    """Collapsed header for a mix-only block."""
    return f"Посадка RFP: {_MIX_TITLE[resolve_mix_mode(mix)]}"


def format_launch_rfp_button(
    *,
    include_packing: bool,
    input_mode: str,
    mix: str,
) -> str:
    """Launch-tab run button: ``Запуск`` plus the selected source and mix.

    Mix is appended only when the source is hybrid. ``include_packing`` does
    not change the text: packing stays a Step4 setting, not this label.

    Args:
        include_packing: Kept for callers. Ignored.
        input_mode: Launch file picker (``legacy_net`` / ``ds_only`` / ``hybrid``).
        mix: ``launch_mix_mode`` (ignored unless ``input_mode`` is hybrid).

    Returns:
        For example ``Запуск (ДС+RFP · без смешения)`` or ``Запуск (части RFP)``.
    """
    del include_packing
    mode = str(input_mode or "").strip()
    short = _LAUNCH_SOURCE_SHORT.get(mode, _LAUNCH_SOURCE_SHORT[INPUT_MODE_LEGACY_NET])
    if mode != INPUT_MODE_HYBRID:
        return f"Запуск ({short})"
    return f"Запуск ({short} · {mix_suffix_ru(mix)})"
