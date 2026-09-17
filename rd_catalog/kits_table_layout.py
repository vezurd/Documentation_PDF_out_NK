"""Default Комплекты column order and widths.

QSettings ``window/kits_header_vN`` still stores the last session. When that
key is bumped or the saved column count no longer matches, the GUI applies
this name-based template instead of the logical header order.

The working copy is ``kits_table_layout.json`` under the catalog runtime
directory. The packaged file next to this module is the factory default
shipped with the program (heal template when runtime is missing).

Qt applies this template when QSettings ``window/kits_header_vN`` is stale.
The WEB monitor always uses the same JSON (runtime, else factory) for
Комплекты column order and default widths. The live QSettings session is
Qt-only until the user clicks «Сохранить шаблон колонок».
"""

from __future__ import annotations

import json
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

KITS_TABLE_LAYOUT_FILENAME = "kits_table_layout.json"
KITS_TABLE_LAYOUT_VERSION = 1

PACKAGED_LAYOUT_PATH = Path(__file__).with_name(KITS_TABLE_LAYOUT_FILENAME)
KITS_HEADER_ALIASES = {
    "Робот · рев.": "Робот МТО · рев.",
    "АН": "АН МТО",
}
# New columns that should sit next to an existing header when an old
# template does not list them yet (otherwise leftovers go to the end).
KITS_HEADER_INSERT_AFTER = {
    "Ок": "Марка",
}


@dataclass
class KitsTableLayout:
    """Visual header order and optional default section widths.

    Attributes:
        order: Header titles from left to right. Unknown names are ignored
            on apply; current columns missing from the list stay at the end
            in logical order.
        widths: Header title → pixel width. Missing or non-positive values
            leave the section at its Qt default size.
    """

    order: tuple[str, ...]
    widths: dict[str, int] = field(default_factory=dict)


def runtime_layout_path(runtime_dir: str | Path) -> Path:
    """Return the runtime JSON path under *runtime_dir*."""

    return Path(runtime_dir) / KITS_TABLE_LAYOUT_FILENAME


def canonical_kits_header_name(name: str) -> str:
    """Map a stored header title to the current Комплекты name."""

    return KITS_HEADER_ALIASES.get(name, name)


def merge_header_order(
    template_order: Sequence[str],
    known_names: Sequence[str],
) -> tuple[str, ...]:
    """Return visual names: template first, then leftovers in logical order.

    Leftover ``Ок`` is inserted after ``Марка`` when that column is
    already in the template, so an old saved layout does not push the
    composite flag to the far right.

    Args:
        template_order: Saved left-to-right titles.
        known_names: Current logical header titles.

    Returns:
        Names that exist in *known_names*, each once.
    """

    known = tuple(str(name) for name in known_names)
    known_set = set(known)
    seen: set[str] = set()
    merged: list[str] = []
    for raw in template_order:
        name = canonical_kits_header_name(str(raw))
        if name not in known_set or name in seen:
            continue
        merged.append(name)
        seen.add(name)
    for name in known:
        if name in seen:
            continue
        anchor = KITS_HEADER_INSERT_AFTER.get(name)
        if anchor is not None and anchor in merged:
            merged.insert(merged.index(anchor) + 1, name)
        else:
            merged.append(name)
        seen.add(name)
    return tuple(merged)


def parse_kits_table_layout(raw: Any) -> KitsTableLayout | None:
    """Parse a JSON object into a layout, or return None if unusable.

    Args:
        raw: Decoded JSON value.

    Returns:
        Layout, or ``None`` when required fields are missing or invalid.
    """

    if not isinstance(raw, dict):
        return None
    order_raw = raw.get("order")
    if not isinstance(order_raw, list):
        return None
    order: list[str] = []
    seen: set[str] = set()
    for item in order_raw:
        if not isinstance(item, str) or not item.strip():
            return None
        name = item.strip()
        if name in seen:
            continue
        order.append(name)
        seen.add(name)
    if not order:
        return None
    widths: dict[str, int] = {}
    widths_raw = raw.get("widths")
    if widths_raw is not None:
        if not isinstance(widths_raw, dict):
            return None
        for key, value in widths_raw.items():
            if not isinstance(key, str) or not key.strip():
                continue
            try:
                width = int(value)
            except (TypeError, ValueError):
                continue
            if width > 0:
                widths[key.strip()] = width
    return KitsTableLayout(order=tuple(order), widths=widths)


def load_kits_table_layout(path: str | Path) -> KitsTableLayout | None:
    """Load a layout JSON file.

    Args:
        path: File to read.

    Returns:
        Layout, or ``None`` when the file is missing or corrupt.
    """

    target = Path(path)
    if not target.is_file():
        return None
    try:
        raw = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError):
        return None
    return parse_kits_table_layout(raw)


def load_default_kits_table_layout(
    runtime_dir: str | Path,
    *,
    packaged_path: Path | None = None,
) -> KitsTableLayout | None:
    """Load runtime layout, else the packaged factory.

    Args:
        runtime_dir: Catalog runtime directory.
        packaged_path: Factory file. ``None`` uses :data:`PACKAGED_LAYOUT_PATH`.

    Returns:
        Layout, or ``None`` when both sources are missing or corrupt.
    """

    runtime = load_kits_table_layout(runtime_layout_path(runtime_dir))
    if runtime is not None:
        return runtime
    factory = PACKAGED_LAYOUT_PATH if packaged_path is None else packaged_path
    return load_kits_table_layout(factory)


def resolve_kits_table_layout(
    runtime_dir: str | Path,
    known_names: Sequence[str],
    *,
    packaged_path: Path | None = None,
) -> KitsTableLayout:
    """Return visual Комплекты order and widths for the current headers.

    Runtime JSON wins over the packaged factory. Unknown template names are
    dropped; columns missing from the file stay at the end in logical order.

    Args:
        runtime_dir: Catalog runtime directory.
        known_names: Current logical header titles (``KITS_HEADERS``).
        packaged_path: Factory file. ``None`` uses :data:`PACKAGED_LAYOUT_PATH`.

    Returns:
        Layout whose ``order`` contains every known name once.
    """

    known = tuple(str(name) for name in known_names)
    loaded = load_default_kits_table_layout(
        runtime_dir, packaged_path=packaged_path
    )
    if loaded is None:
        return KitsTableLayout(order=known, widths={})
    order = merge_header_order(loaded.order, known)
    known_set = set(known)
    widths: dict[str, int] = {}
    for raw_name, width in loaded.widths.items():
        name = canonical_kits_header_name(raw_name)
        if name in known_set and width > 0:
            widths[name] = width
    return KitsTableLayout(order=order, widths=widths)


def write_kits_table_layout(path: str | Path, layout: KitsTableLayout) -> Path:
    """Atomically write a layout JSON file.

    Args:
        path: Target JSON path.
        layout: Order and optional widths to persist.

    Returns:
        The written path.

    Raises:
        OSError: If the file cannot be written.
    """

    payload: dict[str, Any] = {
        "version": KITS_TABLE_LAYOUT_VERSION,
        "order": list(layout.order),
    }
    if layout.widths:
        payload["widths"] = {
            name: int(width)
            for name, width in layout.widths.items()
            if int(width) > 0
        }
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(payload, ensure_ascii=False, indent=2) + "\n"
    tmp_path = target.with_name(target.name + ".tmp")
    tmp_path.write_text(text, encoding="utf-8")
    tmp_path.replace(target)
    return target


def save_default_kits_table_layout(
    layout: KitsTableLayout,
    runtime_dir: str | Path,
    *,
    packaged_path: Path | None = None,
    write_packaged: bool = True,
) -> tuple[Path, Path | None]:
    """Write the runtime template and, when possible, the packaged factory.

    Args:
        layout: Order and widths to store as the new default.
        runtime_dir: Catalog runtime directory.
        packaged_path: Factory file. ``None`` uses :data:`PACKAGED_LAYOUT_PATH`.
        write_packaged: When False, skip the factory file (tests).

    Returns:
        ``(runtime_path, packaged_path_or_none)``. Packaged is ``None`` when
        skipped or the write failed.

    Raises:
        OSError: If the runtime file cannot be written.
    """

    runtime = write_kits_table_layout(runtime_layout_path(runtime_dir), layout)
    if not write_packaged:
        return runtime, None
    factory = PACKAGED_LAYOUT_PATH if packaged_path is None else packaged_path
    try:
        return runtime, write_kits_table_layout(factory, layout)
    except OSError:
        return runtime, None
