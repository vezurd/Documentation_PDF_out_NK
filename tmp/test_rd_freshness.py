"""RD file-stamp census and the missing-file Auto MTO stop."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from rd_catalog.customer_pi_auto_mto import (
    AutoMtoCompareResult,
    AutoMtoFile,
    auto_mto_compare_status,
    needs_auto_mto_compare,
)
from rd_catalog.kits import GoogleKit
from rd_catalog.models import make_path_key
from rd_catalog.rd_freshness import (
    AUTO_REFRESH_MAX_FOLDERS,
    RdFileStamp,
    collect_rd_file_stamps,
    dirty_mark_folders,
    google_snapshot_fingerprint,
    rd_read_error_is_missing,
    rescan_folder_for_file,
    take_rescan_batch,
)
from rd_catalog.rd_freshness_cli import main as census_main


_RD = Path(r"C:\rd")
_FILE = (
    _RD
    / "9103"
    / "05_SOS"
    / "Для передачи"
    / "04_рев.0-AN02_AGCC.287-9103-SOS"
    / "DWG"
    / "AGCC.287-9103-SOS.MTO-0001_0-AN02_RU.xlsx"
)


def _stamp(path: Path, size: int, mtime_ns: int) -> RdFileStamp:
    text = str(path)
    return RdFileStamp(text, make_path_key(text), size, mtime_ns)


def _kit(comment: str) -> GoogleKit:
    return GoogleKit(
        title="9103",
        mark="SOS",
        mark_raw="SOS",
        title_system="9103-SOS",
        sheet_revision="0",
        sheet_appendix="AN02",
        sheet_revision_text="0-AN02",
        status_sheet="",
        comment_raw=comment,
        events=(),
        last_event=None,
        row_index=2,
    )


class RdFreshnessTests(unittest.TestCase):
    def test_missing_read_error_and_compare_is_not_requeued(self) -> None:
        self.assertTrue(
            rd_read_error_is_missing(
                "RD MTO read failed: [Errno 2] No such file or directory: 'x'"
            )
        )
        self.assertFalse(rd_read_error_is_missing("RD MTO filename stem is not recognized"))
        file = AutoMtoFile(
            title="9103",
            mark="SOS",
            spec="AGCC.287-9103-SOS.MTO-0001",
            rd_revision="0-AN02",
            relpath="9103/SOS/a.xlsx",
            row_count=1,
            fingerprint="fp",
        )
        status = auto_mto_compare_status(
            files=(file,),
            rd_path=str(_FILE),
            rd_revision="0-AN02",
            comparison=AutoMtoCompareResult(
                match_kind="not_compared",
                members=(),
                rd_rows=0,
                auto_rows=0,
                combinations_checked=0,
                error="RD MTO read failed: [Errno 2] No such file or directory",
            ),
            comparison_rd_path=str(_FILE),
        )
        self.assertEqual(status.kind, "error")
        self.assertEqual(status.text, "ошибка")
        self.assertFalse(needs_auto_mto_compare(status))

    def test_dirty_folder_is_the_mark_not_the_root(self) -> None:
        base = _stamp(_FILE, 10, 100)
        self.assertEqual(
            rescan_folder_for_file(str(_FILE), str(_RD)),
            str(_RD / "9103" / "05_SOS"),
        )
        self.assertEqual(dirty_mark_folders((base,), (base,), str(_RD)), ())
        changed = _stamp(_FILE, 11, 100)
        folders = dirty_mark_folders((base,), (changed,), str(_RD))
        self.assertEqual(folders, (str(_RD / "9103" / "05_SOS"),))
        missing = dirty_mark_folders((base,), (), str(_RD))
        self.assertEqual(missing, (str(_RD / "9103" / "05_SOS"),))
        self.assertEqual(rescan_folder_for_file(str(_RD), str(_RD)), "")

    def test_batch_keeps_four(self) -> None:
        folders = tuple(rf"C:\rd\{index}\SOS" for index in range(6))
        batch, left = take_rescan_batch(folders, limit=AUTO_REFRESH_MAX_FOLDERS)
        self.assertEqual(len(batch), 4)
        self.assertEqual(left, 2)

    def test_google_fingerprint_ignores_order(self) -> None:
        left = google_snapshot_fingerprint((_kit("a"), _kit("b")), ())
        right = google_snapshot_fingerprint((_kit("b"), _kit("a")), ())
        self.assertEqual(left, right)
        self.assertNotEqual(left, google_snapshot_fingerprint((_kit("c"),), ()))

    def test_census_lists_size_and_mtime_and_skips_old(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            kept = (
                root
                / "9103"
                / "05_SOS"
                / "Для передачи"
                / "04_рев.0-AN02_AGCC.287-9103-SOS"
                / "DWG"
                / "AGCC.287-9103-SOS.MTO-0001_0-AN02_RU.xlsx"
            )
            skipped = (
                root
                / "9103"
                / "05_SOS"
                / "old"
                / "Для передачи"
                / "01_рев.0_AGCC.287-9103-SOS"
                / "DWG"
                / "AGCC.287-9103-SOS.MTO-0001_0_RU.xlsx"
            )
            noise = kept.parent / "readme.txt"
            kept.parent.mkdir(parents=True)
            skipped.parent.mkdir(parents=True)
            kept.write_bytes(b"abc")
            skipped.write_bytes(b"nope")
            noise.write_text("x", encoding="utf-8")
            census = collect_rd_file_stamps(root, skip_dirs=("old",))
            self.assertTrue(census.completed)
            self.assertEqual(len(census.stamps), 1)
            self.assertEqual(census.stamps[0].size, 3)
            self.assertGreater(census.stamps[0].mtime_ns, 0)
            kept.write_bytes(b"abcd")
            again = collect_rd_file_stamps(root, skip_dirs=("old",))
            folders = dirty_mark_folders(census.stamps, again.stamps, str(root))
            self.assertEqual(folders, (str(root / "9103" / "05_SOS"),))
            job = root / "job.json"
            job.write_text(
                __import__("json").dumps(
                    {
                        "rd_root": str(root),
                        "sq_root": "",
                        "skip_dirs": ["old"],
                        "baseline": [
                            {
                                "path": item.path,
                                "path_key": item.path_key,
                                "size": item.size,
                                "mtime_ns": item.mtime_ns,
                            }
                            for item in census.stamps
                        ],
                    }
                ),
                encoding="utf-8",
            )
            self.assertEqual(census_main(["--job", str(job)]), 0)


if __name__ == "__main__":
    unittest.main()
