"""Read-only: side-by-side explanation of the export choice for a few kits."""

from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog.config import load_config  # noqa: E402
from rd_catalog.db import CatalogDatabase  # noqa: E402
from rd_catalog.pipeline import resolve_all_approved_contours  # noqa: E402
from rd_catalog.robot_mto_sync import (  # noqa: E402
    _freshness_key,
    collect_kit_mto_candidates,
    collect_robot_mto_files,
)

KITS = [
    ("3240", "KSB1"),
    ("6100", "SOS"),
    ("6550", "SOS"),
    ("1600", "SOT"),
]

config = load_config()
database = CatalogDatabase(config.db_path)
records = database.list_files()
overlay_ids = {
    int(row["file_id"]) for row in database.current_overlay() if row.get("detected_current")
}
contours = {
    (c.title, c.mark): c
    for c in resolve_all_approved_contours(
        database, records=records, detected_current_ids=overlay_ids
    )
}


def dt(ns: int | None) -> str:
    if not ns:
        return "?"
    return datetime.fromtimestamp(int(ns) / 1e9).strftime("%d.%m.%Y")


with database._connection() as conn:  # noqa: SLF001 - read-only diagnostic
    for title, mark in KITS:
        print("=" * 78)
        print(f"КОМПЛЕКТ {title}/{mark}")
        print("=" * 78)

        print("\n  Папки передач на диске (NN, ревизия папки, MTO внутри):")
        pkgs = conn.execute(
            """
            SELECT id, sequence, package_path, revision_text, max_mtime_ns, is_as_build
            FROM kit_package WHERE title=? AND mark=? ORDER BY sequence
            """,
            (title, mark),
        ).fetchall()
        for p in pkgs:
            folder = Path(p["package_path"]).name
            prefix = str(p["package_path"]).casefold() + "\\"
            mtos = [
                r for r in conn.execute(
                    "SELECT path, mtime_ns FROM file_entry"
                    " WHERE file_kind='mto_xlsx' AND present=1"
                ).fetchall()
                if r["path"].casefold().startswith(prefix)
                and f"-{title}-{mark}.".casefold() in r["path"].casefold()
            ]
            names = ", ".join(f"{Path(m['path']).name} ({dt(m['mtime_ns'])})" for m in mtos) or "— нет MTO —"
            flag = "  [as-build]" if p["is_as_build"] else ""
            print(f"    NN={str(p['sequence'] or '?'):<3} {folder}{flag}")
            print(f"        файлы папки до {dt(p['max_mtime_ns'])};  MTO: {names}")

        print("\n  Журнал Google, столбец F (согласование):")
        evs = conn.execute(
            """
            SELECT e.event_date, e.stage_label, e.revision, e.appendix, e.raw
            FROM google_event e JOIN google_kit g ON g.id=e.kit_id
            WHERE g.title=? AND g.mark=? ORDER BY e.seq
            """,
            (title, mark),
        ).fetchall()
        for e in evs[-6:]:
            print(f"    «{(e['raw'] or '').strip()[:80]}»   -> этап={e['stage_label']}")

        cand = collect_kit_mto_candidates(
            title=title, mark=mark, records=records,
            detected_current_ids=overlay_ids, discover_siblings=False,
        )
        today = max(cand, key=_freshness_key).path if cand else ""
        robot = collect_robot_mto_files(title=title, mark=mark, records=records)
        contour = contours.get((title, mark))

        print("\n  ИТОГ:")
        print(f"    у робота сейчас лежит : {Path(robot[0].path).name if robot else '<нет>'}")
        print(f"    вариант «последний»   : {Path(today).parent.parent.name}\\...\\{Path(today).name}" if today else "    вариант «последний»   : <нет>")
        if contour and contour.mto_path:
            print(f"    вариант «согласован.» : {contour.package_path.split(chr(92))[-1]}\\...\\{Path(contour.mto_path).name}")
            print(f"    код А стоит на рев.   : {contour.approved_revision_text}   (нашли: {contour.match_reason}, доверие {contour.confidence})")
        print()
