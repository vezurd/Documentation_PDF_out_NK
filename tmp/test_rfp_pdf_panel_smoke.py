"""Headless smoke for the RFP · PDF extract panel."""

from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PySide6.QtWidgets import QApplication, QCheckBox, QPushButton

from RFQ.rfp_parts.pdf_rfp_extract import PdfRfpIssue, PdfRfpResult
from ds_compare_center.center_window import _TAB_BY_NAME, _TAB_INDEX_RFP_PDF
from ds_compare_center.rfp_pdf_panel import RfpPdfPanel


def _fake_result(*, errors: int = 0, outside: int = 0) -> PdfRfpResult:
    return PdfRfpResult(
        xlsx_path=Path("out.xlsx"),
        report_path=Path("report.txt"),
        out_dir=Path("."),
        ds_label="ДС98_35Б",
        row_count=12,
        contract_errors=errors,
        outside_words=outside,
        issues=(
            PdfRfpIssue("WARN", "пример замечания"),
            PdfRfpIssue("ERROR", "ошибка контракта"),
        ),
    )


class RfpPdfPanelSmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_recognize_button_and_show_result(self) -> None:
        panel = RfpPdfPanel(on_run=lambda *a: None)
        texts = [btn.text() for btn in panel.findChildren(QPushButton)]
        boxes = [box.text() for box in panel.findChildren(QCheckBox)]
        self.assertTrue(any("Диадок" in text for text in boxes))
        self.assertTrue(panel._chk_strip.isChecked())
        self.assertIn("Распознать", texts)
        self.assertIn("Выбрать PDF", texts)
        self.assertIn("Открыть xlsx", texts)
        self.assertIn("Открыть отчёт", texts)
        self.assertFalse(panel._btn_open_xlsx.isEnabled())

        panel.show_result(_fake_result())
        self.app.processEvents()
        self.assertEqual(panel._table.rowCount(), 2)
        self.assertEqual(panel._table.item(0, 0).text(), "WARN")
        banner = panel._banner.text()
        self.assertIn("12", banner)
        self.assertIn("0", banner)
        self.assertTrue(panel._btn_open_xlsx.isEnabled())
        self.assertTrue(panel._btn_open_report.isEnabled())
        panel.close()
        panel.deleteLater()
        self.app.processEvents()

    def test_show_result_caps_issues_at_30(self) -> None:
        panel = RfpPdfPanel(on_run=lambda *a: None)
        issues = tuple(
            PdfRfpIssue("WARN", f"row {i}") for i in range(45)
        )
        result = PdfRfpResult(
            xlsx_path=Path("out.xlsx"),
            report_path=Path("report.txt"),
            out_dir=Path("."),
            ds_label="ДС1",
            row_count=45,
            contract_errors=1,
            outside_words=2,
            issues=issues,
        )
        panel.show_result(result)
        self.app.processEvents()
        self.assertEqual(panel._table.rowCount(), 30)
        panel.close()
        panel.deleteLater()
        self.app.processEvents()

    def test_tab_index_rfp_pdf_is_11(self) -> None:
        self.assertEqual(_TAB_INDEX_RFP_PDF, 11)
        self.assertEqual(_TAB_BY_NAME["rfp_pdf"], 11)
        self.assertEqual(_TAB_BY_NAME["pdf"], 11)


if __name__ == "__main__":
    unittest.main(verbosity=2)
