"""Read-only: impact of the narrow "dominated later transfer" overlay rule.

Rule under test: the winner picked by today's key loses to an older-sequence
candidate that is *both* higher by filename revision *and* newer by mtime.
That is exactly the 1600/SOT shape (NN21 holds rev 02 from 26.11, NN20 holds
rev 03 from 27.11) and nothing else.

Nothing is written; the live database is only read.
"""

from __future__ import annotations

import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog import overlay as rd_overlay  # noqa: E402
from rd_catalog.config import load_config  # noqa: E402
from rd_catalog.db import CatalogDatabase  # noqa: E402
from rd_catalog.models import FileKind, ParseStatus, SourceKind  # noqa: E402
from rd_catalog.parse import record_has_canonical_layout  # noqa: E402

config = load_config()
database = CatalogDatabase(config.db_path)

with database._connection() as conn:  # noqa: SLF001 - read-only diagnostic
    parsed = list(CatalogDatabase._present_rd_parsed_files(conn))
    ignored = {k.casefold() for k in CatalogDatabase._ignored_path_keys(conn)}

parsed = [f for f in parsed if record_has_canonical_layout(f, config.rd_root)]
pdf_overlay, mto_overlay = rd_overlay.build_rd_overlays(
    parsed, ignored_path_keys=ignored, rd_root=config.rd_root
)


def sequence_of(file) -> int:
    transfer = file.transfer
    return transfer.sequence if transfer and transfer.sequence is not None else -1


def rank_of(file) -> tuple[int, int, str]:
    return rd_overlay._revision_rank(file.revision, file.appendix)


def dominates(challenger, winner) -> bool:
    """Older transfer, strictly higher real revision, strictly newer file."""

    if sequence_of(challenger) >= sequence_of(winner):
        return False
    challenger_rank = rank_of(challenger)
    winner_rank = rank_of(winner)
    if challenger_rank[0] < 0 or winner_rank[0] < 0:
        return False  # never let V / S / unparsed take over
    if challenger_rank <= winner_rank:
        return False
    return int(challenger.mtime_ns) > int(winner.mtime_ns)


grouped: dict[tuple[str, str], list] = defaultdict(list)
for file in parsed:
    if file.source is not SourceKind.RD:
        continue
    if file.parse_status is not ParseStatus.PARSED:
        continue
    if not file.transfer or file.transfer.parse_status is not ParseStatus.PARSED:
        continue
    if file.path_key.casefold() in ignored:
        continue
    key = rd_overlay.document_key_text(file)
    if key is None:
        continue
    grouped[(file.file_kind.value, key)].append(file)

for kind, overlay_result in (("pdf", pdf_overlay), ("mto_xlsx", mto_overlay)):
    changed = []
    for key, winner in overlay_result.current.items():
        candidates = grouped.get((kind, key), ())
        challengers = [c for c in candidates if dominates(c, winner)]
        if not challengers:
            continue
        best = max(challengers, key=rd_overlay._newest_sort_key)
        changed.append((key, winner, best))

    print()
    print(f"=== {kind.upper()} ===")
    print(f"  документов в overlay: {len(overlay_result.current)}")
    print(f"  меняется текущий файл: {len(changed)}")
    for key, before, after in changed:
        print(f"    {key}")
        print(f"        сейчас: NN={sequence_of(before):<3} "
              f"рев={rd_overlay._file_rev_text(before):<8} {Path(before.path).name}")
        print(f"        после : NN={sequence_of(after):<3} "
              f"рев={rd_overlay._file_rev_text(after):<8} {Path(after.path).name}")
