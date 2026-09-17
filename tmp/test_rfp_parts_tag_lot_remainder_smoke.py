"""Offline smoke: tagged parts counters keep leftover integer lot qty on <NO_TAG>."""

from __future__ import annotations

import sys
import unittest
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.rfp_parts.analyze_rfp_parts import (
    NO_TAG_KEY,
    RfpRecord,
    _build_no_tags_summary_rows,
    _build_summary_rows,
    _build_unit_counters,
    _collect_tag_remarks,
    _compute_coarse_collisions,
    _record_tag_keys,
    _unit_key,
)


def _record(
    *,
    code: str = "BCC0000642",
    units: str = "шт",
    values: Decimal = Decimal("3"),
    tags: str = "",
    file_name: str = "ДС67.xlsx",
    excel_row: int = 24,
    ds_name: str = "ДС67",
    ds_title: str = "2265-KSB",
) -> RfpRecord:
    return RfpRecord(
        kind="parts",
        file_name=file_name,
        sheet="Перечень материалов",
        excel_row=excel_row,
        ds_name=ds_name,
        ds_number="67",
        ds_title=ds_title,
        ds_specification=ds_title,
        tags=tags,
        ds_code_1c="",
        code=code,
        name="муфта",
        type_mark="",
        values=values,
        units=units,
    )


class PartsTagLotRemainderSmokeTest(unittest.TestCase):
    def test_lot_gt_tags_merges_remainder_with_untagged_neighbor(self) -> None:
        tagged = _record(
            tags="2265-S-FV-0711",
            values=Decimal("3"),
            excel_row=142,
        )
        untagged = _record(
            tags="",
            values=Decimal("9"),
            excel_row=150,
        )
        aggregation = _build_unit_counters(
            [tagged, untagged], [], emit_warnings=False
        )
        parsed = _record_tag_keys(tagged, [])
        self.assertEqual(parsed, ["2265-S-FV-0711"])
        tagged_key = _unit_key(tagged, parsed[0])
        no_tag_key = _unit_key(untagged, "")
        self.assertEqual(no_tag_key[3], NO_TAG_KEY)
        self.assertEqual(aggregation.units[tagged_key], 1)
        self.assertEqual(aggregation.units[no_tag_key], 11)
        remarks = _collect_tag_remarks([tagged, untagged])
        self.assertIn("lot_gt_tags", {item["kind"] for item in remarks})

    def test_tags_equal_lot_no_remainder(self) -> None:
        rec = _record(tags="TAG-A;TAG-B", values=Decimal("2"))
        aggregation = _build_unit_counters([rec], [], emit_warnings=False)
        parsed = _record_tag_keys(rec, [])
        self.assertEqual(len(parsed), 2)
        for tag in parsed:
            self.assertEqual(aggregation.units[_unit_key(rec, tag)], 1)
        self.assertNotIn(_unit_key(rec, ""), aggregation.units)

    def test_tags_gt_lot_no_negative_extra(self) -> None:
        rec = _record(tags="T-1;T-2;T-3", values=Decimal("1"))
        warnings: list[tuple[str, str, str]] = []
        aggregation = _build_unit_counters([rec], warnings, emit_warnings=True)
        parsed = _record_tag_keys(rec, [])
        self.assertEqual(len(parsed), 3)
        for tag in parsed:
            self.assertEqual(aggregation.units[_unit_key(rec, tag)], 1)
        self.assertNotIn(_unit_key(rec, ""), aggregation.units)
        remarks = _collect_tag_remarks([rec])
        self.assertIn("tags_gt_lot", {item["kind"] for item in remarks})
        self.assertTrue(warnings)

    def test_summary_rows_collapse_no_tag_qty(self) -> None:
        tagged = _record(
            tags="2265-S-FV-0711",
            values=Decimal("3"),
            excel_row=142,
        )
        untagged = _record(
            tags="",
            values=Decimal("9"),
            excel_row=150,
        )
        records = [tagged, untagged]
        aggregation = _build_unit_counters(records, [], emit_warnings=False)
        rows = _build_summary_rows(aggregation, _compute_coarse_collisions(records))
        tagged_rows = [item for item in rows if item.record.tags]
        untagged_rows = [item for item in rows if not item.record.tags]
        self.assertEqual(len(tagged_rows), 1)
        self.assertEqual(tagged_rows[0].record.values, Decimal("1"))
        self.assertEqual(len(untagged_rows), 1)
        self.assertEqual(untagged_rows[0].record.values, Decimal("11"))

    def test_no_tags_net_keeps_source_rows_and_lot_qty(self) -> None:
        kits = [
            _record(
                tags="RX, TX",
                values=Decimal("1"),
                excel_row=20 + index,
                ds_name="ДС45",
                ds_title="8630-KSB3",
                code="BCC0002659",
            )
            for index in range(5)
        ]
        rows = _build_no_tags_summary_rows(
            kits, _compute_coarse_collisions(kits)
        )
        self.assertEqual(len(rows), 5)
        self.assertEqual([item.record.values for item in rows], [Decimal("1")] * 5)
        self.assertTrue(all(item.record.tags == "" for item in rows))
        self.assertEqual([item.record.ds_number for item in rows], ["1", "2", "3", "4", "5"])

    def test_no_tags_net_does_not_sum_same_code_or_expand_tags(self) -> None:
        tagged = _record(
            tags="2265-S-FV-0711",
            values=Decimal("3"),
            excel_row=142,
        )
        untagged = _record(
            tags="",
            values=Decimal("9"),
            excel_row=150,
        )
        records = [tagged, untagged]
        rows = _build_no_tags_summary_rows(
            records, _compute_coarse_collisions(records)
        )
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0].record.values, Decimal("3"))
        self.assertEqual(rows[1].record.values, Decimal("9"))
        self.assertEqual(rows[0].record.tags, "")
        self.assertEqual(rows[1].record.tags, "")


if __name__ == "__main__":
    unittest.main(verbosity=2)
