"""Regression tests for structured RFP progress emission and parsing."""

from __future__ import annotations

import contextlib
import io
import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ds_compare_center.rfp_progress_parser import (
    MILESTONE_IDS,
    MilestoneEvent,
    RfpProgressParser,
)
from RFQ.tags_rfp_compare.rfp_progress import (
    MILESTONE_PREFIX,
    VALID_MILESTONE_IDS,
    emit_milestone,
)


def _line(milestone_id: str, state: str, detail: str = "") -> str:
    payload = {"milestone_id": milestone_id, "state": state, "detail": detail}
    return MILESTONE_PREFIX + json.dumps(
        payload, ensure_ascii=False, separators=(",", ":")
    )


class RfpProgressMarkersSmokeTest(unittest.TestCase):
    def test_emitter_exact_prefix_compact_json_and_validation(self) -> None:
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            emit_milestone("packing_lists", "Skipped", "УЛ выключены")

        expected = (
            '@@RFP_MILESTONE {"milestone_id":"packing_lists",'
            '"state":"Skipped","detail":"УЛ выключены"}\n'
        )
        self.assertEqual(output.getvalue(), expected)
        with self.assertRaises(ValueError):
            emit_milestone("unknown", "Done")
        with self.assertRaises(ValueError):
            emit_milestone("prepare", "done")

    def test_emitter_and_parser_share_canonical_ids(self) -> None:
        self.assertEqual(set(MILESTONE_IDS), set(VALID_MILESTONE_IDS))
        self.assertIn("ds_id_check", MILESTONE_IDS)
        self.assertIn("ds_mp_check", MILESTONE_IDS)
        self.assertIn("ul_preflight", MILESTONE_IDS)
        self.assertEqual(
            MILESTONE_IDS[MILESTONE_IDS.index("prepare") + 1],
            "ds_id_check",
        )
        self.assertEqual(
            MILESTONE_IDS[MILESTONE_IDS.index("ds_id_check") + 1],
            "ds_mp_check",
        )
        self.assertEqual(
            MILESTONE_IDS[MILESTONE_IDS.index("parts_preflight") + 1],
            "ul_preflight",
        )
        self.assertEqual(
            MILESTONE_IDS[MILESTONE_IDS.index("ul_preflight") + 1],
            "rfp_load",
        )

    def test_marker_survives_every_single_split_point(self) -> None:
        raw = (_line("rfp_load", "Done", "строк=17") + "\r\n").encode("utf-8")
        expected = [MilestoneEvent("rfp_load", "Done", "строк=17")]

        for split_at in range(1, len(raw)):
            parser = RfpProgressParser()
            events = parser.feed(raw[:split_at])
            events.extend(parser.feed(raw[split_at:]))
            events.extend(parser.flush())
            self.assertEqual(events, expected, f"split_at={split_at}")

    def test_glued_complete_lines_invalid_lines_and_unterminated_flush(self) -> None:
        parser = RfpProgressParser()
        text = "\n".join(
            (
                "ordinary process output",
                _line("prepare", "Running"),
                "@@RFP_MILESTONE not-json",
                '@@RFP_MILESTONE {"milestone_id":"bad","state":"Done"}',
                '@@RFP_MILESTONE {"milestone_id":"vo_load","state":"Impossible"}',
                _line("parts_preflight", "Done", "ok"),
            )
        )
        complete, unterminated = text.rsplit("\n", 1)

        events = parser.feed(complete + "\n")
        self.assertEqual(
            events,
            [MilestoneEvent("prepare", "Running", "")],
        )
        self.assertEqual(parser.feed(unterminated), [])
        self.assertEqual(
            parser.flush(),
            [MilestoneEvent("parts_preflight", "Done", "ok")],
        )

    def test_representative_lifecycle_includes_ul_skipped_and_error(self) -> None:
        lifecycle = (
            ("prepare", "Running", ""),
            ("prepare", "Done", "result_dir=tmp"),
            ("parts_preflight", "Done", "checklist ok"),
            ("ul_preflight", "Done", "unchanged; кэш совпадает с исходной папкой; position_rows=1"),
            ("rfp_load", "Done", "rows=2"),
            ("mto_google_load", "Done", "rows=3"),
            ("vo_load", "Done", "title_systems=1"),
            ("units_gate", "Done", "RFP: проверено 0, преобразовано 0, без Google 0; MTO: проверено 0, преобразовано 0, без Google 0; УЛ: проверено 0, преобразовано 0, без Google 0"),
            ("step4_match", "Done", "rows=2"),
            ("global_checks", "Done", ""),
            ("packing_lists", "Skipped", "include_packing_lists=false"),
            ("save_excel", "Error", "disk full"),
            ("complete", "Error", "export failed"),
        )
        stream = "\n".join(_line(*event) for event in lifecycle) + "\n"
        parser = RfpProgressParser()

        parsed = parser.feed(stream)

        self.assertEqual(
            parsed,
            [MilestoneEvent(*event) for event in lifecycle],
        )
        self.assertIn(MilestoneEvent("packing_lists", "Skipped", "include_packing_lists=false"), parsed)
        self.assertEqual(parsed[-1], MilestoneEvent("complete", "Error", "export failed"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
