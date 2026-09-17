"""Audit the confidence downgrades and the kits whose exported MTO changed.

Read-only against ``tmp/rd_catalog_copy.sqlite3``. Looks for downgrades where
the two revision texts differ only cosmetically (zero padding, dashes, case),
which would be a false alarm rather than a real mismatch.
"""

from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog.config import load_config  # noqa: E402
from rd_catalog.db import CatalogDatabase  # noqa: E402
from rd_catalog.pipeline import resolve_all_approved_contours  # noqa: E402

WATCH = {("8445", "SOS"), ("8630", "KSB3"), ("8950", "SOO2"), ("6100", "SOS")}


def cosmetic(text: str) -> str:
    """Collapse padding, dashes and case so only real differences remain."""

    value = re.sub(r"[\s\u2010\u2011\u2012\u2013\u2014_-]+", "", (text or "").casefold())
    return re.sub(r"(?<!\d)0+(\d)", r"\1", value)


config = load_config()
database = CatalogDatabase(ROOT / "tmp" / "rd_catalog_copy.sqlite3")
database.initialize()
records = database.list_files()
overlay_ids = {
    int(row["file_id"]) for row in database.current_overlay() if row.get("detected_current")
}

contours = resolve_all_approved_contours(
    database,
    records=records,
    detected_current_ids=overlay_ids,
    rd_root=config.rd_root,
)

packages: dict[tuple[str, str], list[tuple[int | None, str, str]]] = {}
with database._connection() as connection:
    for row in connection.execute(
        "SELECT title, mark, sequence, package_path, revision_text FROM kit_package"
    ):
        packages.setdefault((row["title"], row["mark"]), []).append(
            (row["sequence"], Path(row["package_path"]).name, row["revision_text"] or "")
        )

confidence = Counter(c.confidence for c in contours)
print("доверие:", dict(confidence.most_common()))
print()

suspicious: list[str] = []
real = 0
for contour in contours:
    if contour.confidence == "high" or not contour.package_path:
        continue
    approved = contour.approved_revision_text or ""
    chosen = ""
    for sequence, name, revision_text in packages.get((contour.title, contour.mark), []):
        if name == Path(contour.package_path).name:
            chosen = revision_text
            break
    if not approved or not chosen:
        continue
    if cosmetic(approved) == cosmetic(chosen):
        suspicious.append(
            f"{contour.title}/{contour.mark}: согласовано «{approved}», выбрано «{chosen}»"
        )
    else:
        real += 1

print(f"понижений с реальным различием ревизий: {real}")
print(f"понижений, где различие только косметическое: {len(suspicious)}")
for line in suspicious[:20]:
    print("   ", line)
print()

print("=== контрольные комплекты ===")
for contour in contours:
    if (contour.title, contour.mark) not in WATCH:
        continue
    print(f"  {contour.title}/{contour.mark}")
    print(f"      согласовано : {contour.approved_revision_text}")
    print(f"      папка       : {Path(contour.package_path).name if contour.package_path else '—'}"
          f"  NN={contour.package_sequence}")
    print(f"      файл        : {Path(contour.mto_path).name if contour.mto_path else '—'}")
    print(f"      причина     : {contour.match_reason}   доверие: {contour.confidence}")
    if contour.warnings:
        for warning in contour.warnings:
            print(f"      ! {warning}")
    print("      все папки комплекта:")
    for sequence, name, revision_text in sorted(
        packages.get((contour.title, contour.mark), []), key=lambda item: (item[0] is None, item[0])
    ):
        print(f"         NN={str(sequence or '?'):<4} rev={revision_text:<10} {name}")
