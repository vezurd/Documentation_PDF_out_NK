"""Headless smoke for the RFP run-tab progress table."""

from __future__ import annotations

import copy
import os
import sys
import unittest
from pathlib import Path
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PySide6.QtWidgets import QApplication, QLabel

from RFQ.tags_rfp_compare.rfp_tags_utils import get_default_config
from ds_compare_center.rfp_progress_parser import MILESTONE_IDS, MilestoneEvent
import ds_compare_center.rfp_run_panel as run_module
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

    def test_input_mode_block_shows_mode_and_saves_choice(self) -> None:
        config = get_default_config()
        saved: dict[str, dict] = {}

        def load() -> dict:
            return copy.deepcopy(config)

        def save(cfg: dict) -> bool:
            saved["config"] = copy.deepcopy(cfg)
            config.clear()
            config.update(copy.deepcopy(cfg))
            return True

        with (
            mock.patch.object(run_module, "load_config", side_effect=load),
            mock.patch.object(run_module, "save_config", side_effect=save),
            mock.patch.object(
                run_module,
                "resolve_effective_rfp_path",
                side_effect=FileNotFoundError("нет файла"),
            ),
        ):
            panel = RfpRunPanel()
            panel.resize(520, 720)
            panel.show()
            self.app.processEvents()

            block = panel._source_mode
            self.assertFalse(block.is_expanded())
            self.assertFalse(block._body.isVisible())
            self.assertIn("Свод частей RFP (rfp_parts_net)", block._title.text())

            block._header.clicked.emit()
            self.app.processEvents()
            self.assertTrue(block.is_expanded())
            self.assertTrue(block._body.isVisible())
            descriptions = "\n".join(
                label.text() for label in block._body.findChildren(QLabel)
            )
            self.assertIn("rfp_parts_net.xlsx", descriptions)
            self.assertIn("_ds_baseline", descriptions)
            self.assertIn("_ds_hybrid", descriptions)

            block._radios["hybrid"].click()
            self.app.processEvents()
            self.assertEqual(saved["config"]["rfp_parts"]["input_mode"], "hybrid")
            self.assertIn("Свод ДС-RFP для запуска", block._title.text())
            self.assertTrue(block.is_expanded())

            block._header.clicked.emit()
            self.app.processEvents()
            self.assertFalse(block.is_expanded())
            self.assertIn("Свод ДС-RFP для запуска", block._title.text())

            panel.close()
            panel.deleteLater()
            self.app.processEvents()


if __name__ == "__main__":
    unittest.main(verbosity=2)
