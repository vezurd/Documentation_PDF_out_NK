"""Headless smoke for the RFP ↔ manager-surname panel."""

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

from RFQ.tags_rfp_compare.ds_manager_roster import (
    DsRosterCompareResult,
    DsRosterRow,
)
from ds_compare_center.rfp_ds_mp_panel import RfpDsMpPanel


def _fake_result() -> DsRosterCompareResult:
    return DsRosterCompareResult(
        matrix_path=Path("book.xlsx"),
        parts_dir=Path("parts"),
        rows=[
            DsRosterRow(
                actual_label="ДС24",
                sequential_label="ДС92",
                copy_name="ДС24_92Б",
                manager="Михеев Д.",
                status="фамилия с Лист2, имя на Лист2 перевёрнуто",
                was_on_sheet1="ДС92_24Б",
            ),
            DsRosterRow(
                actual_label="ДС99",
                sequential_label="ДС99",
                copy_name="ДС99",
                manager="",
                status="нет на Лист2 — заполните",
                was_on_sheet1="",
            ),
        ],
        load_error=None,
    )


class RfpDsMpPanelSmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_fill_from_result_rows_and_banner(self) -> None:
        panel = RfpDsMpPanel()
        result = _fake_result()
        panel.fill_from_result(result)
        self.app.processEvents()
        self.assertEqual(panel._table.rowCount(), 2)
        self.assertEqual(panel._table.item(0, 0).text(), "ДС24_92Б")
        self.assertEqual(panel._table.item(0, 3).text(), "Михеев Д.")
        banner = panel._banner.text()
        self.assertIn("дозаполн", banner.lower())
        self.assertTrue(result.needs_fill)
        self.assertFalse(result.is_ok)
        panel.close()
        panel.deleteLater()
        self.app.processEvents()


if __name__ == "__main__":
    unittest.main(verbosity=2)
