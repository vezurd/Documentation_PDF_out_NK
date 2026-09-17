"""Break down the «прочее» bucket of RD layout violations.

Read-only against ``tmp/rd_catalog_copy.sqlite3``. The point is to find out
whether «прочее» hides nameable patterns that deserve their own reason before
a report UI is built on top of the classifier.
"""

from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog.config import load_config  # noqa: E402
from rd_catalog.db import CatalogDatabase  # noqa: E402
from rd_catalog.parse import (  # noqa: E402
    is_transfer_gate_folder_name,
    list_layout_violations,
)

config = load_config()
database = CatalogDatabase(ROOT / "tmp" / "rd_catalog_copy.sqlite3")
database.initialize()
violations = list_layout_violations(records=database.list_files(), rd_root=config.rd_root)

rd_root = Path(config.rd_root)


def relative_parts(path: str) -> tuple[str, ...]:
    try:
        return Path(path).relative_to(rd_root).parts
    except ValueError:
        return Path(path).parts


others = [v for v in violations if v.reason == "other"]
print(f"всего нарушений {len(violations)}, из них «прочее» {len(others)}")
print()

depth = Counter(len(relative_parts(v.path)) - 1 for v in others)
print("глубина папок после корня РД:", dict(sorted(depth.items())))
print()

has_gate = Counter(
    any(is_transfer_gate_folder_name(part) for part in relative_parts(v.path)[:-1])
    for v in others
)
print("есть ли на пути шлюз «Для передачи»:", {str(k): v for k, v in has_gate.items()})
print()

third = Counter(
    relative_parts(v.path)[2] if len(relative_parts(v.path)) > 3 else "<файл сразу в марке>"
    for v in others
)
print("папка сразу под маркой (топ-15):")
for name, count in third.most_common(15):
    print(f"   {count:5d}  {name}")
print()

kinds = Counter(Path(v.path).suffix.lower() for v in others)
print("расширения:", dict(kinds.most_common(8)))
print()

print("сколько из «прочее» — это MTO (влияет на выгрузку):")
mto = [v for v in others if ".mto-" in Path(v.path).name.casefold()]
print(f"   {len(mto)} файлов, из них xlsx: "
      f"{sum(1 for v in mto if Path(v.path).suffix.lower() == '.xlsx')}")
print()

print("примеры (по одному на характерную папку):")
seen: set[str] = set()
for v in others:
    parts = relative_parts(v.path)
    key = parts[2] if len(parts) > 3 else "<файл сразу в марке>"
    if key in seen:
        continue
    seen.add(key)
    print("   ", "/".join(parts))
    if len(seen) >= 12:
        break
