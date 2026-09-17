"""Smoke tests for RFP parts sheet skip-list, soft select, and empty-file ERROR."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.rfp_parts.analyze_rfp_parts import (
    PARTS_KIND,
    _extract_records,
    is_expected_sheet_name,
    is_never_read_sheet,
    split_part_sheets,
)

_HEADER = [
    "№ п/п",
    "Титул",
    "Спецификация",
    "Tag",
    "Код 1С",
    "Код РД",
    "Наименование МТР",
    "Технические характеристики",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "Кол-во",
    "Ед. изм",
]
_DATA = [
    1,
    "8525-SKUD",
    "spec",
    "T-1",
    "",
    "BCC0000001",
    "Кабель",
    "mark",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    2,
    "шт",
]


def _write_book(path: Path, sheets: dict[str, list[list[object]]]) -> None:
    wb = Workbook()
    default = wb.active
    names = list(sheets)
    default.title = names[0]
    for title, rows in sheets.items():
        ws = default if title == names[0] else wb.create_sheet(title)
        for row in rows:
            ws.append(row)
    wb.save(path)
    wb.close()


class RfpPartsSheetSelectSmokeTest(unittest.TestCase):
    def test_never_read_sheet_patterns(self) -> None:
        self.assertTrue(is_never_read_sheet("Титульный лист"))
        self.assertTrue(is_never_read_sheet("Титул"))
        self.assertTrue(is_never_read_sheet("Лист регистрации изменений"))
        self.assertTrue(is_never_read_sheet("Содержание"))
        self.assertTrue(is_never_read_sheet("Легенда"))
        self.assertFalse(is_never_read_sheet("Перечень материалов"))
        self.assertFalse(is_never_read_sheet("перечень поставки"))
        self.assertFalse(is_never_read_sheet("Лист1"))
        self.assertTrue(is_expected_sheet_name("перечень материалов"))

    def test_split_skips_cover_keeps_data(self) -> None:
        skipped, candidates = split_part_sheets(
            ["Титульный лист", "Перечень материалов", "Лист2"]
        )
        self.assertEqual(skipped, ("Титульный лист",))
        self.assertEqual(candidates, ("Перечень материалов", "Лист2"))

    def test_single_sheet_any_name_is_read(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "ДС7_ГФ.xlsx"
            _write_book(path, {"перечень поставки": [_HEADER, _DATA]})
            warnings: list[tuple[str, str, str]] = []
            records, stats = _extract_records(path, PARTS_KIND, warnings)
            self.assertEqual(stats.sheet, "перечень поставки")
            self.assertEqual(len(records), 1)
            self.assertEqual(records[0].code, "BCC0000001")
            self.assertFalse(any(level == "ERROR" for level, _, _ in warnings))
            self.assertFalse(
                any("рабочих листов" in msg for _, _, msg in warnings)
            )

    def test_cover_plus_expected_sheet_no_unread_warn(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "ДС13.xlsx"
            _write_book(
                path,
                {
                    "Титульный лист": [["Cover"]],
                    "Перечень материалов": [_HEADER, _DATA],
                },
            )
            warnings: list[tuple[str, str, str]] = []
            records, stats = _extract_records(path, PARTS_KIND, warnings)
            self.assertEqual(stats.sheet, "Перечень материалов")
            self.assertEqual(len(records), 1)
            self.assertFalse(
                any("не читались" in msg for _, _, msg in warnings)
            )

    def test_two_working_sheets_unread_warn(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "ДС3_ГФ.xlsx"
            _write_book(
                path,
                {
                    "Лист1": [_HEADER, _DATA],
                    "Лист2": [["noise"]],
                },
            )
            warnings: list[tuple[str, str, str]] = []
            records, _stats = _extract_records(path, PARTS_KIND, warnings)
            self.assertEqual(len(records), 1)
            warn_msgs = [msg for level, _, msg in warnings if level == "WARN"]
            self.assertTrue(any("не читались: Лист2" in msg for msg in warn_msgs))

    def test_empty_file_without_header_is_error(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "empty.xlsx"
            _write_book(path, {"Лист1": [["нет шапки RFP"]]})
            warnings: list[tuple[str, str, str]] = []
            records, stats = _extract_records(path, PARTS_KIND, warnings)
            self.assertEqual(records, [])
            self.assertEqual(stats.rows, 0)
            errors = [msg for level, _, msg in warnings if level == "ERROR"]
            self.assertTrue(errors)
            self.assertTrue(
                any("позиции не прочитаны" in msg or "шапка RFP не найдена" in msg for msg in errors)
            )

    def test_missing_required_column_is_error(self) -> None:
        broken = list(_HEADER)
        broken[6] = "Наименование"  # not «Наименование МТР»
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "shift.xlsx"
            _write_book(path, {"Перечень материалов": [broken, _DATA]})
            warnings: list[tuple[str, str, str]] = []
            records, stats = _extract_records(path, PARTS_KIND, warnings)
            self.assertEqual(records, [])
            self.assertEqual(stats.rows, 0)
            errors = [msg for level, _, msg in warnings if level == "ERROR"]
            self.assertTrue(any("не найдены обязательные столбцы" in msg for msg in errors))

    def test_rfq_substring_in_tag_is_error(self) -> None:
        rfq_row = list(_DATA)
        rfq_row[3] = "AGCC.287-0000-12.4.1-RFQ-0028"
        ok_row = list(_DATA)
        ok_row[0] = 2
        ok_row[3] = "8529-SS-01-S-UZ-0711"
        ok_row[5] = "BCC0000002"
        with tempfile.TemporaryDirectory() as raw:
            path = Path(raw) / "ДС101.xlsx"
            _write_book(
                path,
                {"Перечень материалов": [_HEADER, rfq_row, ok_row]},
            )
            warnings: list[tuple[str, str, str]] = []
            records, stats = _extract_records(path, PARTS_KIND, warnings)
            self.assertEqual([item.code for item in records], ["BCC0000002"])
            errors = [msg for level, _, msg in warnings if level == "ERROR"]
            self.assertTrue(any("подстрока RFQ" in msg for msg in errors))
            self.assertEqual(stats.rows, 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
