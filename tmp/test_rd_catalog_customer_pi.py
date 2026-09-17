"""Smoke: packaged customer PI pickle loads and matches the BCC filter contract."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.customer_pi import (
    COLUMN_CAPTIONS,
    COLUMN_KEYS,
    FORMAT_ID,
    IDENTITY_CAPTIONS,
    IDX_QTY_RD,
    IDX_RD_REV,
    IDX_SPEC,
    IDX_SUPPLIER,
    LAYOUT_FULL,
    LAYOUT_IDENTITY,
    N_COLS,
    SCHEMA_VERSION,
    canonical_pi_row,
    combine_headers,
    default_pickle_path,
    detect_pi_layout,
    load_customer_pi,
    _row_texts,
)


def _uncombined_full_headers() -> tuple[list[str], list[str]]:
    """Rebuild the two-row TDSheet header from canonical captions."""

    row0 = [""] * N_COLS
    row1 = [""] * N_COLS
    for i, cap in enumerate(COLUMN_CAPTIONS):
        if " / " in cap:
            left, right = cap.split(" / ", 1)
            row0[i] = left
            row1[i] = right
        else:
            row0[i] = cap
    return row0, row1


class CustomerPiLayoutTest(unittest.TestCase):
    def test_full_two_row_header_maps_1_to_1(self) -> None:
        row0, row1 = _uncombined_full_headers()
        layout = detect_pi_layout(row0, row1)
        self.assertEqual(layout.kind, LAYOUT_FULL)
        self.assertEqual(layout.header_row_count, 2)
        self.assertEqual(layout.data_start_row0, 2)
        self.assertEqual(layout.source_n_cols, N_COLS)
        self.assertEqual(layout.src_for_dest, tuple(range(N_COLS)))
        combined = combine_headers(row0, row1, N_COLS)
        self.assertEqual(combined, list(COLUMN_CAPTIONS))

    def test_identity_one_row_pads_to_49(self) -> None:
        data0 = [""] * 23
        data0[0] = "ОЗХ"
        data0[2] = "БИ.СИ.СИ., ООО"
        data0[7] = "1513"
        data0[8] = "POS"
        data0[9] = "AGCC.287-1513-POS.BOM-0001"
        data0[13] = "BCC0000659"
        data0[15] = "01-AN02"
        data0[21] = "шт"
        data0[22] = 1.0
        layout = detect_pi_layout(IDENTITY_CAPTIONS, ["ОЗХ", "", "БИ.СИ.СИ., ООО"])
        self.assertEqual(layout.kind, LAYOUT_IDENTITY)
        self.assertEqual(layout.header_row_count, 1)
        self.assertEqual(layout.data_start_row0, 1)
        self.assertEqual(layout.source_n_cols, 23)
        cells = {i: v for i, v in enumerate(data0)}
        row = canonical_pi_row(cells, layout.src_for_dest)
        self.assertEqual(len(row), N_COLS)
        self.assertEqual(row[IDX_SPEC], "AGCC.287-1513-POS.BOM-0001")
        self.assertEqual(row[IDX_RD_REV], "01-AN02")
        self.assertEqual(row[IDX_SUPPLIER], "БИ.СИ.СИ., ООО")
        self.assertEqual(row[IDX_QTY_RD], "1")
        self.assertEqual(row[1], "")  # Центр закупки
        self.assertEqual(row[24], "")  # ПИ
        self.assertEqual(row[48], "")  # ЗИП

    def test_identity_missing_qty_raises(self) -> None:
        headers = list(IDENTITY_CAPTIONS[:-1])
        with self.assertRaises(ValueError) as ctx:
            detect_pi_layout(headers, ["ОЗХ"])
        self.assertIn("По РД", str(ctx.exception))


class CustomerPiDumpSmoke(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        path = default_pickle_path()
        if not path.is_file():
            raise unittest.SkipTest(f"missing dump {path}")
        cls.store = load_customer_pi(path)

    def test_schema_and_width(self) -> None:
        self.assertEqual(self.store.meta.get("format"), FORMAT_ID)
        self.assertEqual(self.store.meta.get("schema_version"), SCHEMA_VERSION)
        self.assertEqual(len(self.store.columns), N_COLS)
        self.assertEqual(len(COLUMN_KEYS), N_COLS)
        self.assertEqual(self.store.columns[10], "Спецификация")
        self.assertEqual(self.store.columns[16], "№ ревизии РД")

    def test_bcc_filter(self) -> None:
        self.assertEqual(len(self.store), 27217)
        rec = self.store.record(0)
        self.assertIn("БИ.СИ.СИ.", rec.supplier)
        for raw in self.store.rows:
            self.assertIn("БИ.СИ.СИ.", raw[3])

    def test_spec_rev_unique_and_mto_map(self) -> None:
        self.assertEqual(self.store.meta["counts"]["specs_with_multiple_rd_rev"], 0)
        rec = self.store.record(1000)
        fields = rec.as_mto_fields()
        self.assertEqual(set(fields), {"CODE", "UNITS", "VALUES", "TAGS", "NAME", "VENDOR", "TYPE_MARK"})
        self.assertEqual(fields["VENDOR"], "")
        self.assertTrue(rec.spec.startswith("AGCC."))
        self.assertEqual(rec.parsed_title, rec.title)
        self.assertEqual(rec.parsed_mark, rec.mark)

    def test_indexes(self) -> None:
        spec = "AGCC.287-8445-SOT.MTO-0001"
        self.assertIn(spec, self.store.by_spec())
        self.assertGreater(len(self.store.by_code_rd()), 100)
        records = self.store.records_for_spec(spec)
        self.assertEqual(len(records), len(self.store.by_spec()[spec]))
        self.assertTrue(all(item.spec == spec for item in records))


class CustomerPiUncHeaderSmoke(unittest.TestCase):
    """Open live EDMS dumps when the BCC share is mounted; skip otherwise."""

    _PI_DIR = Path(
        r"\\bcc\root\CurProjects\di_manegers\Пятницкий_П"
        r"\Амурский ГХК\Поставки\Отчет систем ПИ"
    )
    _FULL = _PI_DIR / "Отчет для заказчика АГХК 07.09.2026.xlsb"
    _IDENTITY = _PI_DIR / "Отчет для заказчика АГХК 14.09.2026_ВСС.xlsb"

    def _layout_of(self, path: Path):
        try:
            from pyxlsb import open_workbook
        except ImportError as exc:
            raise unittest.SkipTest("pyxlsb is not installed") from exc
        if not path.is_file():
            raise unittest.SkipTest(f"missing {path}")
        with open_workbook(str(path)) as wb:
            with wb.get_sheet(wb.sheets[0]) as sh:
                it = sh.rows()
                r0 = {c.c: c.v for c in next(it)}
                r1 = {c.c: c.v for c in next(it)}
        n = 0
        if r0:
            n = max(n, max(r0) + 1)
        if r1:
            n = max(n, max(r1) + 1)
        return detect_pi_layout(_row_texts(r0), _row_texts(r1)), n, r1

    def test_full_dump_header(self) -> None:
        layout, n, _r1 = self._layout_of(self._FULL)
        self.assertEqual(layout.kind, LAYOUT_FULL)
        self.assertEqual(n, N_COLS)
        self.assertEqual(layout.src_for_dest[IDX_SPEC], IDX_SPEC)

    def test_identity_dump_header_and_first_row(self) -> None:
        layout, n, r1 = self._layout_of(self._IDENTITY)
        self.assertEqual(layout.kind, LAYOUT_IDENTITY)
        self.assertEqual(n, 23)
        row = canonical_pi_row(r1, layout.src_for_dest)
        self.assertEqual(row[IDX_SPEC], "AGCC.287-0000-99A-0001")
        self.assertEqual(row[IDX_QTY_RD], "1")
        self.assertIn("БИ.СИ.СИ.", row[IDX_SUPPLIER])


if __name__ == "__main__":
    unittest.main()
