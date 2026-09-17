"""Read-only: impact of ranking overlay by filename revision before sequence.

Current rule : (sequence, revision_rank, mtime_ns, path_key)
Candidate rule: (revision_rank, sequence, mtime_ns, path_key)

Nothing is written; the live database is only read.
"""

from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog import overlay as rd_overlay  # noqa: E402
from rd_catalog.config import load_config  # noqa: E402
from rd_catalog.db import CatalogDatabase  # noqa: E402
from rd_catalog.models import FileKind  # noqa: E402

config = load_config()
database = CatalogDatabase(config.db_path)


def candidate_key(file):
    """Rank by filename revision first, then transfer sequence.

    Element 0 must stay an ``int``: ``build_overlay`` negates it when sorting
    entries, so the revision rank is folded into one integer here.
    """

    transfer = file.transfer
    sequence = transfer.sequence if transfer and transfer.sequence is not None else -1
    rev_int, app_int, _raw = rd_overlay._revision_rank(file.revision, file.appendix)
    return (
        rev_int * 1000 + (app_int + 1),
        sequence,
        file.mtime_ns,
        file.path_key,
    )


with database._connection() as conn:  # noqa: SLF001 - read-only diagnostic
    parsed = list(CatalogDatabase._present_rd_parsed_files(conn))
    ignored = list(CatalogDatabase._ignored_path_keys(conn))

print(f"файлов РД на входе: {len(parsed)}")

results: dict[str, dict[str, dict]] = {}
for label in ("current", "candidate"):
    if label == "candidate":
        rd_overlay._newest_sort_key = candidate_key
    pdf, mto = rd_overlay.build_rd_overlays(
        parsed, ignored_path_keys=ignored, rd_root=config.rd_root
    )
    results[label] = {"pdf": pdf, "mto": mto}

for kind in ("pdf", "mto"):
    now = results["current"][kind]
    new = results["candidate"][kind]
    changed = []
    for key, current_file in now.current.items():
        other = new.current.get(key)
        if other is None or other.path_key != current_file.path_key:
            changed.append((key, current_file, other))

    print()
    print(f"=== {kind.upper()} ===")
    print(f"  документов в overlay: {len(now.current)} -> {len(new.current)}")
    print(f"  меняется текущий файл: {len(changed)}")

    coll_now = Counter(c.kind.value for c in now.collisions)
    coll_new = Counter(c.kind.value for c in new.collisions)
    print("  коллизии сейчас  :", dict(coll_now))
    print("  коллизии после   :", dict(coll_new))

    for key, before, after in changed[:25]:
        b_rev = rd_overlay._file_rev_text(before)
        a_rev = rd_overlay._file_rev_text(after) if after else "—"
        b_nn = before.transfer.sequence if before.transfer else None
        a_nn = after.transfer.sequence if after and after.transfer else None
        print(f"    {key}")
        print(f"        сейчас: NN={b_nn} рев={b_rev}  {Path(before.path).name}")
        print(f"        после : NN={a_nn} рев={a_rev}  "
              f"{Path(after.path).name if after else '<нет>'}")
    if len(changed) > 25:
        print(f"    … и ещё {len(changed) - 25}")
