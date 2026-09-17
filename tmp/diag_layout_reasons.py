"""Report RD layout violations by reason, and prove the gate did not drift.

Read-only against ``tmp/rd_catalog_copy.sqlite3``. The canonical/non-canonical
split must match the numbers measured before the classifier was refined:
25741 present RD files, 2213 of them non-canonical.
"""

from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog.config import load_config  # noqa: E402
from rd_catalog.db import CatalogDatabase  # noqa: E402
from rd_catalog.models import SourceKind  # noqa: E402
from rd_catalog.parse import (  # noqa: E402
    list_layout_violations,
    record_has_canonical_layout,
)

EXPECTED_PRESENT = 25741
EXPECTED_VIOLATIONS = 2213

config = load_config()
database = CatalogDatabase(ROOT / "tmp" / "rd_catalog_copy.sqlite3")
database.initialize()
records = database.list_files()

rd_present = [r for r in records if r.present and r.source is SourceKind.RD]
canonical = sum(1 for r in rd_present if record_has_canonical_layout(r, config.rd_root))
non_canonical = len(rd_present) - canonical

print(f"present RD файлов        : {len(rd_present)}  (ожидалось {EXPECTED_PRESENT})")
print(f"каноничных               : {canonical}  ({canonical / max(len(rd_present), 1):.1%})")
print(f"неканоничных             : {non_canonical}  (ожидалось {EXPECTED_VIOLATIONS})")
# The present count moves as the user works; only the verdict must hold.
drifted = non_canonical != EXPECTED_VIOLATIONS
print(f"отбор файлов             : {'ИЗМЕНИЛСЯ — разбирать' if drifted else 'не изменился'}")
if len(rd_present) != EXPECTED_PRESENT:
    print(
        f"   (всего файлов стало {len(rd_present)} вместо {EXPECTED_PRESENT} — "
        "более свежий скан, не сдвиг отбора)"
    )
print()

violations = list_layout_violations(records=records, rd_root=config.rd_root)
print(f"строк в отчёте: {len(violations)}")
print()

by_reason = Counter((v.reason, v.reason_label) for v in violations)
print("по причинам:")
for (reason, label), count in by_reason.most_common():
    share = count / max(len(violations), 1)
    print(f"   {count:6d}  {share:5.1%}  {reason:<16} {label}")
print()

mto = [v for v in violations if v.is_mto]
print(f"из них файлов MTO: {len(mto)}")
mto_by_reason = Counter(v.reason for v in mto)
for reason, count in mto_by_reason.most_common():
    print(f"   {count:6d}  {reason}")
print()

kits = Counter((v.title, v.mark) for v in violations)
print("худшие 10 титул-марок:")
for (title, mark), count in kits.most_common(10):
    print(f"   {count:6d}  {title or '—'}/{mark or '—'}")
print()

print("по два примера на причину:")
shown: Counter[str] = Counter()
rd_root = Path(config.rd_root)
for violation in violations:
    if shown[violation.reason] >= 2:
        continue
    shown[violation.reason] += 1
    try:
        shortened = Path(violation.path).relative_to(rd_root)
    except ValueError:
        shortened = Path(violation.path)
    print(f"   [{violation.reason}] {shortened}")
