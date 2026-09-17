"""Headless smoke for the UPD merge panel."""

from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PySide6.QtWidgets import QApplication

from RFQ.ds_compare.upd_load import UpdLoadResult, UpdLoadStats
from ds_compare_center.upd_panel import UpdPanel


class UpdPanelSmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_stats_and_buttons(self) -> None:
        panel = UpdPanel()
        self.app.processEvents()
        self.assertTrue(panel._btn_run.text())
        self.assertTrue(panel._edit_root.text())
        result = UpdLoadResult(
            summary_path=r"C:\tmp\upd_summary.xlsx",
            report_path=r"C:\tmp\upd_load_report.txt",
            stats=UpdLoadStats(files_ok=2, files_xlsx=3, position_rows=10),
            success=False,
        )
        panel._set_stats_from_result(result)
        self.app.processEvents()
        text = panel._stats.text()
        self.assertIn("position_row=10", text)
        panel.close()
        panel.deleteLater()
        self.app.processEvents()


if __name__ == "__main__":
    unittest.main(verbosity=2)
