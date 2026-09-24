"""Smoke: RFP root census separates equipment tags from a column-shifted DS."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from openpyxl import Workbook

from RFQ.rfp_parts.rfp_tag_census import scan_rfp_tag_census

_HEADER = (
    "№ п/п",
    "Титул",
    "Спецификация",
    "Линия, Tag-номер",
    "Код 1С",
    "Код РД",
    "Наименование МТР",
    "Технические характеристики",
    "Кол-во",
    "Ед. изм.",
)


def _write(path: Path, rows: list[tuple[object, ...]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Перечень материалов"
    ws.append(_HEADER)
    for row in rows:
        ws.append(row)
    wb.save(path)
    wb.close()


class RfpTagCensusSmoke(unittest.TestCase):
    def test_zero_tags_with_positions_sorts_first(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            _write(
                folder / "b_real.xlsx",
                [("1", "8529", "spec", "8529-K-SB-1234", "", "BCC1", "насос", "", 1, "шт")],
            )
            _write(
                folder / "a_shifted_ds.xlsx",
                [("1", "8529", "spec", "просто текст ДС", "", "BCC2", "труба", "", 2, "м")],
            )
            wb = Workbook()
            wb.active.title = "Титульный лист"
            wb.save(folder / "c_no_sheet.xlsx")
            wb.close()

            census = scan_rfp_tag_census(folder)

        by_name = {row.file_name: row for row in census.rows}
        self.assertEqual(by_name["b_real.xlsx"].tag_count, 1)
        self.assertEqual(by_name["b_real.xlsx"].position_count, 1)
        self.assertTrue(by_name["a_shifted_ds.xlsx"].suspect)
        self.assertEqual(by_name["a_shifted_ds.xlsx"].tag_count, 0)
        self.assertEqual(by_name["a_shifted_ds.xlsx"].position_count, 1)
        self.assertEqual(by_name["c_no_sheet.xlsx"].position_count, 0)
        self.assertEqual(census.rows[0].file_name, "a_shifted_ds.xlsx")
        text = census.copy_tsv()
        self.assertIn("Имя файла\tКол-во тегов\tКол-во позиций", text)
        self.assertIn("a_shifted_ds.xlsx\t0\t1", text)
        self.assertIn("<table", census.copy_html())
        self.assertEqual(census.suspect_count, 1)


if __name__ == "__main__":
    unittest.main()
