"""Headless smoke for the RFP run-tab progress table."""

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

from ds_compare_center.rfp_progress_parser import MILESTONE_IDS, MilestoneEvent
from ds_compare_center.rfp_run_panel import RfpRunPanel

_LONG_DETAIL = (
    r"result_dir=\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP"
    r"\RFP сводный файл\2026.09.03_12.00"
)


class RfpRunProgressTableSmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_table_splits_status_and_keeps_full_detail(self) -> None:
        panel = RfpRunPanel()
        panel.resize(960, 720)
        panel.show()
        self.app.processEvents()

        table = panel._table
        self.assertEqual(table.rowCount(), len(MILESTONE_IDS))
        self.assertEqual(table.columnCount(), 3)
        self.assertEqual(
            [table.horizontalHeaderItem(i).text() for i in range(3)],
            ["Шаг", "Статус", "Детали"],
        )

        prepare_row = panel._milestone_row_by_id["prepare"]
        self.assertEqual(table.item(prepare_row, 1).text(), "Ожидает")
        self.assertEqual(table.item(prepare_row, 2).text(), "")
        mp_row = panel._milestone_row_by_id["ds_mp_check"]
        self.assertEqual(
            table.item(mp_row, 0).text(),
            "Соответствие ДС: RFP ↔ фамилии МП",
        )
        ul_row = panel._milestone_row_by_id["ul_preflight"]
        self.assertEqual(
            table.item(ul_row, 0).text(),
            "Свежесть свода УЛ (при необходимости пересборка)",
        )
        self.assertEqual(
            ul_row,
            panel._milestone_row_by_id["parts_preflight"] + 1,
        )

        panel.apply_milestone("prepare", "Done", _LONG_DETAIL)
        panel.apply_milestone(
            MilestoneEvent("units_gate", "Running", "RFP: проверено 2, преобразовано 1")
        )
        self.app.processEvents()

        self.assertEqual(table.item(prepare_row, 1).text(), "Готово")
        self.assertEqual(table.item(prepare_row, 2).text(), _LONG_DETAIL)
        self.assertEqual(table.item(prepare_row, 2).toolTip(), _LONG_DETAIL)
        self.assertNotIn("Готово —", table.item(prepare_row, 2).text())

        units_row = panel._milestone_row_by_id["units_gate"]
        self.assertEqual(table.item(units_row, 1).text(), "Выполняется")
        self.assertEqual(panel.active_milestone_id, "units_gate")
        self.assertGreater(table.rowHeight(prepare_row), 22)

        panel.close()
        panel.deleteLater()
        self.app.processEvents()


if __name__ == "__main__":
    unittest.main(verbosity=2)
