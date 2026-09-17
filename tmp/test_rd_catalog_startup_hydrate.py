"""Startup snapshot prefers SQLite and skips derived rebuild when pipelines exist."""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.config import load_config
from rd_catalog.db import CatalogDatabase
from rd_catalog.kits import IssuanceKit, latest_issuance_from_sends
from rd_catalog.startup_hydrate import load_google_for_monitor, load_startup_snapshot


def _issuance(title: str, mark: str, *, date: str, row: int) -> IssuanceKit:
    return IssuanceKit(
        title=title,
        mark=mark,
        mark_raw=mark,
        title_system=f"{title}-{mark}",
        revision="01",
        appendix=None,
        revision_text="01",
        status="",
        send_date=date,
        send_date_sortable=date,
        send_transmittal="",
        incoming_control_date="",
        incoming_control_date_sortable="",
        confirm_transmittal="",
        note_raw="",
        row_index=row,
    )


class LatestIssuanceTests(unittest.TestCase):
    def test_keeps_later_date_then_higher_row(self) -> None:
        older = _issuance("9110", "KSB", date="2024-01-01", row=9)
        newer = _issuance("9110", "KSB", date="2024-02-01", row=1)
        other = _issuance("8445", "SOT", date="2024-01-15", row=2)
        latest = latest_issuance_from_sends((older, other, newer))
        by_mark = {kit.mark: kit for kit in latest}
        self.assertEqual(by_mark["KSB"].send_date_sortable, "2024-02-01")
        self.assertEqual(by_mark["SOT"].title, "8445")


class StartupSnapshotTests(unittest.TestCase):
    def test_empty_db_has_no_hatch_and_empty_rows(self) -> None:
        with tempfile.TemporaryDirectory(prefix="rd_startup_") as temp:
            root = Path(temp)
            runtime = root / "runtime"
            runtime.mkdir()
            override = root / "config.json"
            override.write_text(
                json.dumps(
                    {
                        "rd_root": str(root / "rd"),
                        "sq_root": str(root / "sq"),
                        "robot_root": str(root / "robot"),
                        "runtime_dir": str(runtime),
                        "db_path": str(runtime / "catalog.sqlite"),
                        "robot_flat_structure": True,
                        "skip_dirs": ["old"],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            config = load_config(override)
            database = CatalogDatabase(config.db_path)
            database.initialize()
            google = load_google_for_monitor(database, str(config.runtime_dir))
            self.assertFalse(google.loaded)
            snapshot = load_startup_snapshot(database, config)
            self.assertEqual(snapshot.records, [])
            self.assertEqual(snapshot.kit_rows, ())
            self.assertEqual(snapshot.worklist_rows, ())
            self.assertEqual(snapshot.an_by_kit, {})
            self.assertFalse(snapshot.google.loaded)
            self.assertEqual(snapshot.log_lines, ())


if __name__ == "__main__":
    unittest.main()
