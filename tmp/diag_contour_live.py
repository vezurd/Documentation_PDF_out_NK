"""Acceptance check: run the new approved-contour resolver against the live DB.

Read-only. Verifies the three hand-verified kits and prints the confidence
distribution over the whole catalog.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog.config import load_config  # noqa: E402
from rd_catalog.db import CatalogDatabase  # noqa: E402
from rd_catalog.pipeline import (  # noqa: E402
    resolve_all_approved_contours,
    resolve_approved_contour,
)

config = load_config()
database = CatalogDatabase(config.db_path)
records = database.list_files()
overlay_ids = {
    int(row["file_id"])
    for row in database.current_overlay()
    if row.get("detected_current")
}

EXPECTED = {
    ("1600", "SOT"): "MTO-0001_03_RU.xlsx",
    ("6400", "SOT"): "MTO-0001_05-AN01_RU.xlsx",
    ("2225", "KSB"): "MTO-0001_01-AN02_RU.xlsx",
}

print("=== hand-verified kits ===")
failures = 0
for (title, mark), expected_name in EXPECTED.items():
    contour = resolve_approved_contour(
        database,
        records=records,
        detected_current_ids=overlay_ids,
        title=title,
        mark=mark,
    )
    if contour is None:
        print(f"{title}/{mark}: FAIL — resolver returned None")
        failures += 1
        continue
    got = Path(contour.mto_path).name if contour.mto_path else "<нет>"
    matched = got.casefold().endswith(expected_name.casefold())
    verdict = "OK  " if matched else "FAIL"
    failures += not matched
    print(f"{verdict} {title}/{mark}")
    print(f"       согл.рев = {contour.approved_revision_text!r}")
    print(f"       пакет    = {Path(contour.package_path).name!r} (NN={contour.package_sequence})")
    print(f"       как нашли= {contour.match_reason}  доверие={contour.confidence}")
    print(f"       MTO      = {got}  ({contour.mto_source})")
    print(f"       ожидалось= {expected_name}")
    if contour.ambiguity:
        print(f"       неоднозначность: {', '.join(contour.ambiguity)}")
    if contour.warnings:
        for warning in contour.warnings:
            print(f"       ! {warning}")
    print()

banned: set[tuple[str, str]] = set()
banned_path = Path(config.runtime_dir) / "banned_title_marks.json"
if banned_path.exists():
    raw = json.loads(banned_path.read_text(encoding="utf-8"))
    items = raw.get("pairs", raw) if isinstance(raw, dict) else raw
    for item in items:
        if isinstance(item, dict):
            banned.add((str(item.get("title", "")).casefold(), str(item.get("mark", "")).casefold()))
        elif isinstance(item, (list, tuple)) and len(item) == 2:
            banned.add((str(item[0]).casefold(), str(item[1]).casefold()))

contours = [
    contour
    for contour in resolve_all_approved_contours(
        database, records=records, detected_current_ids=overlay_ids
    )
    if (contour.title.casefold(), contour.mark.casefold()) not in banned
]

print("=== whole catalog ===")
print(f"комплектов: {len(contours)}")
for level in ("high", "medium", "low"):
    items = [c for c in contours if c.confidence == level]
    share = 100 * len(items) / max(len(contours), 1)
    print(f"  {level:6s}: {len(items):3d}  ({share:.0f}%)")

print()
print("как определился согласованный пакет:")
reasons: dict[str, int] = {}
for contour in contours:
    reasons[contour.match_reason] = reasons.get(contour.match_reason, 0) + 1
for reason, count in sorted(reasons.items(), key=lambda item: -item[1]):
    print(f"  {count:4d}  {reason}")

print()
print("неоднозначности:")
kinds: dict[str, int] = {}
for contour in contours:
    for token in contour.ambiguity:
        kinds[token] = kinds.get(token, 0) + 1
for token, count in sorted(kinds.items(), key=lambda item: -item[1]):
    print(f"  {count:4d}  {token}")

no_mto = [c for c in contours if not c.mto_path]
print()
print(f"комплектов без MTO в согласованном контуре: {len(no_mto)}")
for contour in no_mto[:15]:
    print(f"  {contour.title}/{contour.mark} рев={contour.approved_revision_text!r} {contour.match_reason}")

print()
print(f"ИТОГ по эталонным комплектам: {'все совпали' if not failures else f'{failures} расхождений'}")
