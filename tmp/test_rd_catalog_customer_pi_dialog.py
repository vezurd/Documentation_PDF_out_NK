"""Offscreen smoke for the customer PI dialog table (filter, sort, preview)."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QApplication

from rd_catalog.customer_pi import (
    COLUMN_KEYS,
    IDX_CODE_RD,
    IDX_MARK,
    IDX_NAME,
    IDX_QTY_RD,
    IDX_RD_REV,
    IDX_SPEC,
    IDX_SUPPLIER,
    IDX_TAGS,
    IDX_TITLE,
    IDX_TYPE_MARK,
    IDX_UNITS,
    N_COLS,
    CustomerPiStore,
)
from rd_catalog.customer_pi_auto_mto import (
    HEADER_ROW,
    PAINT_OUTSIDE_KITS,
    STATUS_COLORS,
    STATUS_QUEUED,
)
from rd_catalog.customer_pi_dialog import (
    CustomerPiDialog,
    MtoPreviewDialog,
    _COL_MARK,
    _COL_ROWS,
    _COL_SPEC,
    _COL_STATUS,
    _COL_TITLE,
)
from rd_catalog.kits import kit_identity_key


def _raw(**values: str) -> tuple[str, ...]:
    cells = [""] * N_COLS
    cells[IDX_SUPPLIER] = values.get("supplier", "БИ.СИ.СИ., ООО")
    cells[IDX_TITLE] = values.get("title", "8445")
    cells[IDX_MARK] = values.get("mark", "SOT")
    cells[IDX_SPEC] = values.get("spec", "AGCC.287-8445-SOT.MTO-0001")
    cells[IDX_RD_REV] = values.get("rd_revision", "04-AN01")
    cells[IDX_TAGS] = values.get("tags", "8445-S-SX-1002")
    cells[IDX_CODE_RD] = values.get("code_rd", "BCC0001")
    cells[IDX_NAME] = values.get("name", "Кабель силовой")
    cells[IDX_TYPE_MARK] = values.get("type_mark", "ВВГнг")
    cells[IDX_UNITS] = values.get("units", "м")
    cells[IDX_QTY_RD] = values.get("qty_rd", "12")
    return tuple(cells)


def _store(*rows: tuple[tuple[str, ...], tuple[str, str, str]]) -> CustomerPiStore:
    raw_rows = [item[0] for item in rows]
    parsed = [item[1] for item in rows]
    return CustomerPiStore(
        meta={
            "source_path": r"\\example\report.xlsb",
            "pickle_path": "memory",
            "parsed_at": "2026-09-10T00:00:00Z",
            "counts": {
                "specs": len(raw_rows),
                "discipline": {"MTO": 1, "BOM": 0, "DS": 0},
            },
        },
        columns=list(COLUMN_KEYS),
        rows=raw_rows,
        excel_rows=list(range(3, 3 + len(raw_rows))),
        parsed=parsed,
    )


class CustomerPiDialogTableSmoke(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_filter_sort_preview_and_kits_paint(self) -> None:
        store = _store(
            (
                _raw(),
                ("8445", "SOT", "MTO"),
            ),
            (
                _raw(
                    spec="AGCC.287-0000-94A-0009",
                    title="0000",
                    mark="94A",
                    rd_revision="0",
                    tags="",
                    code_rd="ZZ1",
                    name="Пакет",
                ),
                ("0000", "94A", "other"),
            ),
        )
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            dialog = CustomerPiDialog(
                pickle_path=root / "missing.pkl",
                dest_dir=root / "АвтоМто",
                store=store,
                kit_keys={kit_identity_key("8445", "SOT")},
            )
            dialog.show()
            self.app.processEvents()
            table = dialog._spec_table
            self.assertTrue(table.isSortingEnabled())
            self.assertEqual(table.rowCount(), 2)
            self.assertIn("Комплекты: 1", dialog._kits_summary_label.text())
            self.assertIn("получено МТО: 1", dialog._kits_summary_label.text())
            self.assertIn("спеки вне комплектов: 1", dialog._kits_summary_label.text())

            by_spec = {}
            for visual in range(table.rowCount()):
                spec = table.item(visual, _COL_SPEC).text()
                by_spec[spec] = visual
            sot_row = by_spec["AGCC.287-8445-SOT.MTO-0001"]
            other_row = by_spec["AGCC.287-0000-94A-0009"]
            self.assertEqual(table.item(sot_row, _COL_TITLE).text(), "8445")
            self.assertEqual(table.item(sot_row, _COL_MARK).text(), "SOT")
            queued = QColor(STATUS_COLORS[STATUS_QUEUED])
            faded = QColor(STATUS_COLORS[PAINT_OUTSIDE_KITS])
            self.assertEqual(
                table.item(sot_row, _COL_STATUS).background().color().name(),
                queued.name(),
            )
            self.assertEqual(
                table.item(other_row, _COL_SPEC).background().color().name(),
                faded.name(),
            )

            dialog._spec_filter.setText("8445-SOT")
            self.app.processEvents()
            self.assertFalse(table.isRowHidden(sot_row))
            self.assertTrue(table.isRowHidden(other_row))
            self.assertIn("показано 1", dialog._spec_summary_label.text())

            dialog._spec_filter.setText("")
            self.app.processEvents()
            table.sortItems(_COL_ROWS, Qt.SortOrder.DescendingOrder)
            self.app.processEvents()
            self.assertEqual(table.rowCount(), 2)

            first_spec = table.item(0, _COL_SPEC).text()
            dialog._on_spec_double_clicked(0, 0)
            self.app.processEvents()
            self.assertEqual(len(dialog._preview_windows), 1)
            preview = dialog._preview_windows[0]
            self.assertIsInstance(preview, MtoPreviewDialog)
            self.assertEqual(preview._table.columnCount(), len(HEADER_ROW))
            self.assertEqual(preview._table.horizontalHeaderItem(0).text(), HEADER_ROW[0])
            self.assertIn(first_spec, preview.windowTitle())
            preview.close()
            self.app.processEvents()
            dialog.close()
            self.app.processEvents()


if __name__ == "__main__":
    unittest.main()
