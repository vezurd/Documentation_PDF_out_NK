"""Smoke: one-file RFP / DS / UL checks stay off the full-run outputs."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.ds_compare.tsd_packing_load import (
    TSD_PACKING_CRITICAL_REPORT_NAME,
    inspect_one_tsd_file,
)
from RFQ.rfp_parts.analyze_rfp_parts import (
    RUN_DIR_STAMP_RE,
    resolve_parts_workbooks,
)
from RFQ.rfp_parts.ds_baseline import collect_ds_workbooks
from RFQ.rfp_parts.ds_jobs import DsJobResult
from RFQ.rfp_parts.one_file_check import (
    ONE_FILE_CHECK_DIR_NAME,
    one_file_kind_dir,
    run_ds_one_file_job,
    run_rfp_one_file_job,
    run_ul_one_file_job,
)


def _touch_xlsx(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.active.title = "Лист1"
    wb.active["A1"] = "x"
    wb.save(path)
    wb.close()


class OneFileCheckSmoke(unittest.TestCase):
    def test_side_folder_is_not_a_full_run_stamp(self) -> None:
        self.assertFalse(RUN_DIR_STAMP_RE.fullmatch(ONE_FILE_CHECK_DIR_NAME))
        kind = one_file_kind_dir("RFP", base=Path("C:/reports"))
        self.assertEqual(kind.name, "RFP")
        self.assertEqual(kind.parent.name, ONE_FILE_CHECK_DIR_NAME)

    def test_parts_only_files_ignores_siblings(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            keep = folder / "keep.xlsx"
            other = folder / "other.xlsx"
            _touch_xlsx(keep)
            _touch_xlsx(other)
            picked = resolve_parts_workbooks(folder, [keep])
            self.assertEqual(picked, [keep])
            self.assertEqual(len(resolve_parts_workbooks(folder)), 2)

    def test_ds_only_paths_ignores_siblings(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            keep = folder / "keep.xlsx"
            other = folder / "other.xlsx"
            _touch_xlsx(keep)
            _touch_xlsx(other)
            files, skipped = collect_ds_workbooks(folder, only_paths=[keep])
            self.assertEqual(skipped, [])
            self.assertEqual([item.path.name for item in files], ["keep.xlsx"])
            empty, _skipped_empty = collect_ds_workbooks(folder, only_paths=[])
            self.assertEqual(empty, [])

    def test_rfp_job_limits_analyze_and_uses_side_dir(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "part.xlsx"
            stamp = Path(tmp) / ONE_FILE_CHECK_DIR_NAME / "RFP" / "2026.09.23_11.00"
            _touch_xlsx(path)
            with (
                patch(
                    "RFQ.rfp_parts.one_file_check.one_file_stamp_dir",
                    return_value=stamp,
                ),
                patch(
                    "RFQ.rfp_parts.one_file_check.run_rfp_parts_analyze",
                    return_value=stamp / "rfp_parts_net.xlsx",
                ) as analyze,
            ):
                result = run_rfp_one_file_job(path)
            self.assertTrue(result.success)
            self.assertEqual(analyze.call_args.kwargs["only_files"], [path])
            self.assertEqual(Path(analyze.call_args.kwargs["out_dir"]), stamp)

    def test_ds_job_does_not_write_launch_baseline(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "ds.xlsx"
            _touch_xlsx(path)
            with patch(
                "RFQ.rfp_parts.one_file_check.run_ds_baseline_job",
                return_value=DsJobResult(True, "аудит", str(Path(tmp) / "stamp")),
            ) as baseline:
                result = run_ds_one_file_job(path, tmp, str(Path(tmp) / "reg.xlsx"))
            self.assertTrue(result.success)
            self.assertIn("ds.xlsx", result.message)
            self.assertFalse(baseline.call_args.kwargs["write_baseline"])
            self.assertEqual(baseline.call_args.kwargs["only_paths"], (path,))
            parent = Path(baseline.call_args.args[2])
            self.assertEqual(parent.name, "ДС")
            self.assertEqual(parent.parent.name, ONE_FILE_CHECK_DIR_NAME)

    def test_missing_file_does_not_call_the_full_tool(self) -> None:
        with patch(
            "RFQ.rfp_parts.one_file_check.run_rfp_parts_analyze"
        ) as analyze:
            result = run_rfp_one_file_job(Path("no-such-rfp.xlsx"))
        self.assertFalse(result.success)
        analyze.assert_not_called()

    def test_ul_one_file_writes_reports_without_cache(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            src = folder / "pl.xlsx"
            out = folder / "out"
            _touch_xlsx(src)

            def _boom() -> Path:
                raise AssertionError("production UL cache must not be touched")

            with (
                patch(
                    "RFQ.ds_compare.tsd_packing_load.tsd_packing_cache_dir",
                    side_effect=_boom,
                ),
                patch(
                    "RFQ.rfp_parts.one_file_check.one_file_stamp_dir",
                    return_value=out,
                ),
            ):
                inspected = inspect_one_tsd_file(str(src), out)
                job = run_ul_one_file_job(src)
            self.assertFalse(inspected.success)
            self.assertTrue((out / TSD_PACKING_CRITICAL_REPORT_NAME).is_file())
            self.assertFalse(any(out.rglob("*.cache")))
            self.assertFalse(job.success)
            self.assertEqual(Path(job.result_path or ""), out)


if __name__ == "__main__":
    unittest.main()
