"""Clipboard dump for pasting a catalog UI item into an AI chat.

Qt-free. The GUI collects labeled fields; this module only formats them.
"""

from __future__ import annotations

from collections.abc import Sequence

from rd_catalog.doc_bundle import DocumentTreeLabelOptions

HANDOFF_HEADER = "RD Catalog · Передать роботу"

_NODE_KIND_LABELS = {
    "title": "титул",
    "mark": "марка",
    "revision": "ревизия (папка NN)",
}


def format_robot_handoff(
    surface: str,
    fields: Sequence[tuple[str, str]],
) -> str:
    """Format a paste block for an AI chat.

    Args:
        surface: Where the user right-clicked (tree, table, …).
        fields: ``(label, value)`` pairs; empty values are omitted.

    Returns:
        Multiline UTF-8 text ending with a newline.
    """

    lines = [HANDOFF_HEADER, f"Где: {surface}", ""]
    for key, value in fields:
        text = str(value).strip() if value is not None else ""
        if not text:
            continue
        if "\n" in text:
            lines.append(f"{key}:")
            lines.extend(f"  {part}" for part in text.splitlines())
        else:
            lines.append(f"{key}: {text}")
    return "\n".join(lines).rstrip() + "\n"


def node_kind_label(kind: str | None) -> str:
    """Return a Russian name for a documents-tree node kind.

    Args:
        kind: Stored ``title`` / ``mark`` / ``revision``, or empty.

    Returns:
        Localized level name, or the raw kind.
    """

    raw = (kind or "").strip()
    return _NODE_KIND_LABELS.get(raw, raw or "—")


def format_label_options(options: DocumentTreeLabelOptions) -> str:
    """Return a compact dump of documents-tree label checkboxes.

    Args:
        options: Current extra-field flags.

    Returns:
        Comma-separated ``ключ=да|нет`` list.
    """

    bits = (
        ("статус MTO", options.show_mto_status),
        ("дата", options.show_date),
        ("папка NN", options.show_folder),
        ("статус Google", options.show_review),
        ("текущая", options.show_current),
        ("MTO", options.show_mto),
        ("рабочая", options.show_working),
        ("as-build", options.show_as_build),
    )
    return ", ".join(f"{name}={_yes_no(flag)}" for name, flag in bits)


def yes_no(flag: bool) -> str:
    """Return ``да`` or ``нет`` for a boolean flag.

    Args:
        flag: Value to format.

    Returns:
        Russian yes/no token.
    """

    return _yes_no(flag)


def _yes_no(flag: bool) -> str:
    return "да" if flag else "нет"
