"""Offline tests for sequential vs actual DS identity parser (no UNC)."""

from __future__ import annotations

import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.rfp_parts.ds_id_coverage import (
    check_ds_id_coverage,
    run_ds_id_coverage_job,
)
from RFQ.rfp_parts.ds_identity import (
    parse_rfp_ds_identity,
    parse_ul_folder_ds_identity,
)


class DsIdentityParserSmokeTest(unittest.TestCase):
    def test_rfp_correction_ds92_24b(self) -> None:
        ident = parse_rfp_ds_identity(
            "ДС92_24Б. AGCC.287-0000-12.4.1-RFP-0014_01_RU.xlsx"
        )
        self.assertEqual(ident.sequential, 24)
        self.assertEqual(ident.actual, 92)
        self.assertEqual(ident.letter, "Б")
        self.assertTrue(ident.compound)
        self.assertEqual(ident.match_key, 92)
        self.assertEqual(ident.label, "ДС92_24Б")

    def test_rfp_simple_and_copy_suffix_not_compound(self) -> None:
        simple = parse_rfp_ds_identity("ДС23. AGCC.287-0000-12.4.1-RFP-0016_0_RU.xlsx")
        self.assertEqual((simple.sequential, simple.actual, simple.compound), (23, 23, False))
        copy_file = parse_rfp_ds_identity("ДС4905_1.xlsx")
        self.assertEqual(copy_file.sequential, 4905)
        self.assertEqual(copy_file.actual, 4905)
        self.assertFalse(copy_file.compound)
        part = parse_rfp_ds_identity("ДС13_ (5)_7230-SOT.xlsx")
        self.assertEqual(part.actual, 13)
        self.assertFalse(part.compound)
        appendix = parse_rfp_ds_identity(
            "ДС47_Приложение №1 – «Детализированная спецификация №3»_6600.xlsx"
        )
        self.assertEqual(appendix.actual, 47)
        self.assertFalse(appendix.compound)

    def test_ul_folder_variants(self) -> None:
        folder = parse_ul_folder_ds_identity("согл УЛ ДС24")
        self.assertEqual(folder.actual, 24)
        self.assertEqual(folder.kind, "simple")
        bare = parse_ul_folder_ds_identity("согл УЛ 4905")
        self.assertEqual(bare.actual, 4905)
        self.assertEqual(bare.kind, "bare")
        gf = parse_ul_folder_ds_identity("согл УЛ ГФ 5титулов")
        self.assertEqual(gf.kind, "gf")
        self.assertIsNone(gf.match_key)
        gf_ds1 = parse_ul_folder_ds_identity("согл УЛ ДС1 ГФ 5титулов")
        self.assertEqual(gf_ds1.kind, "simple")
        self.assertEqual(gf_ds1.actual, 1)
        gf_ds7 = parse_ul_folder_ds_identity("согл УЛ ДС7 ГФ 2 тит")
        self.assertEqual(gf_ds7.kind, "simple")
        self.assertEqual(gf_ds7.actual, 7)
        compound = parse_ul_folder_ds_identity("согл УЛ ДС4_11")
        self.assertEqual(compound.actual, 4)
        self.assertEqual(compound.sequential, 11)
        self.assertEqual(compound.match_keys, [4])
        self.assertTrue(compound.compound)

    def test_coverage_join_on_actual_number(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            root = Path(raw)
            ul = root / "ul"
            rfp = root / "rfp"
            (ul / "согл УЛ ДС92").mkdir(parents=True)
            (ul / "согл УЛ ДС92" / "pl.xlsx").write_bytes(b"x")
            (ul / "согл УЛ 4905").mkdir()
            (ul / "согл УЛ ДС1 ГФ 5титулов").mkdir()
            (ul / "согл УЛ ГФ без номера").mkdir()
            rfp.mkdir()
            (rfp / "ДС92_24Б. AGCC.xlsx").write_bytes(b"x")
            (rfp / "ДС4905.xlsx").write_bytes(b"x")
            (rfp / "ДС95. AGCC.xlsx").write_bytes(b"x")
            (rfp / "ДС1_Госфин.xlsx").write_bytes(b"x")
            result = check_ds_id_coverage(parts_dir=rfp, tsd_root=ul)
            by_label = {row.actual_label: row for row in result.rows}
            self.assertEqual(by_label["ДС92"].status, "both")
            self.assertEqual(by_label["ДС92"].rfp_labels, ("ДС92_24Б",))
            self.assertEqual(by_label["ДС4905"].status, "both")
            self.assertEqual(by_label["ДС1"].status, "both")
            self.assertEqual(by_label["ДС1"].rfp_labels, ("ДС1",))
            self.assertEqual(by_label["ДС95"].status, "rfp_only")
            self.assertEqual(by_label["ГФ"].status, "gf_special")
            self.assertFalse(result.is_ok)
            self.assertIn("только RFP ДС95", result.summary_line())
            self.assertIn("ГФ", result.summary_line())
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                job = run_ds_id_coverage_job(parts_dir=rfp, tsd_root=ul)
            self.assertTrue(job.success)
            self.assertIsNone(job.result_path)
            self.assertIn("только RFP ДС95", job.message)
            self.assertIn("только RFP ДС95", buf.getvalue())


if __name__ == "__main__":
    unittest.main()
