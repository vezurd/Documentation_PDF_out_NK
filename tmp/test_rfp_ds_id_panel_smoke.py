"""Headless smoke for the RFP ↔ UL DS-id coverage panel."""

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

from RFQ.rfp_parts.ds_id_coverage import DsIdCoverRow, DsIdCoverageResult
from ds_compare_center.rfp_ds_id_panel import RfpDsIdPanel


def _fake_result() -> DsIdCoverageResult:
    return DsIdCoverageResult(
        ul_root=Path("."),
        rfp_root=Path("."),
        rows=[
            DsIdCoverRow(
                actual_label="ДС24",
                status="both",
                status_ru="Есть в УЛ и RFP",
                ul_folder="согл УЛ ДС24",
                ul_xlsx=1,
                rfp_labels=("ДС92_24Б",),
                rfp_files=("ДС92_24Б.xlsx",),
                rfp_sequential=(24,),
                note="Связка по фактическому номеру ДС.",
            ),
            DsIdCoverRow(
                actual_label="ДС95",
                status="rfp_only",
                status_ru="Только RFP",
                ul_folder="",
                ul_xlsx=0,
                rfp_labels=("ДС95",),
                rfp_files=("ДС95.xlsx",),
                rfp_sequential=(95,),
                note="Файл RFP есть, папки УЛ нет.",
            ),
        ],
        both=1,
        rfp_only=1,
    )


class RfpDsIdPanelSmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_fill_from_result_rows_and_banner(self) -> None:
        panel = RfpDsIdPanel()
        result = _fake_result()
        panel.fill_from_result(result)
        self.app.processEvents()
        self.assertEqual(panel._table.rowCount(), 2)
        banner = panel._banner.text()
        self.assertTrue(
            "только RFP" in banner or "ОК" in banner,
            f"unexpected banner: {banner!r}",
        )
        panel.close()
        panel.deleteLater()
        self.app.processEvents()


if __name__ == "__main__":
    unittest.main(verbosity=2)
