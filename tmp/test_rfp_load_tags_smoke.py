"""Smoke: load_tags=false blanks TAGS before RFP split (VALUES=1, two tags)."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from base.base_classes import RowStd, RowType
from base.tables_columns import CODE, TAGS, UNITS, VALUES
from RFQ.tags_rfp_compare.column_optimization import (
    apply_load_tags_mode,
    strip_loaded_tags,
)
from RFQ.tags_rfp_compare.rfp_tags_utils import (
    get_default_asbuild_config,
    get_default_config,
)
from RFQ.tags_rfp_compare.step1_load_rfp import split_rfp_rows


def _position_row(*, code: str, qty: object, tags: object) -> RowStd:
    row = RowStd()
    row.row_type = RowType.position_row
    row.el[CODE].value = code
    row.el[UNITS].value = "шт"
    row.el[VALUES].value = qty
    row.el[TAGS].value = tags
    return row


def _ok_ds_id_coverage() -> mock.Mock:
    fake = mock.Mock()
    fake.load_error = None
    fake.summary_line.return_value = "ОК: пар 0, замечаний нет"
    return fake


class RfpLoadTagsSmokeTest(unittest.TestCase):
    def test_defaults_keep_tags_enabled(self) -> None:
        self.assertTrue(get_default_config()["load_tags"])
        self.assertTrue(get_default_asbuild_config()["load_tags"])

    def test_strip_clears_list_and_string_tags(self) -> None:
        rows = [
            _position_row(code="A", qty=1, tags=["T1", "T2"]),
            _position_row(code="B", qty=2, tags="X, Y"),
        ]
        cleared = strip_loaded_tags(rows)
        self.assertEqual(cleared, 2)
        self.assertEqual(rows[0].get_tags_list(), [])
        self.assertEqual(rows[1].get_tags_list(), [])
        self.assertEqual(rows[0].get_value(VALUES), 1)

    def test_split_keeps_values_1_two_tags_as_one_row_after_strip(self) -> None:
        tagged = _position_row(code="BCC0001", qty=1, tags=["TAG-A", "TAG-B"])
        expanded_on, split_tags_on, _split_no_on = split_rfp_rows(
            [tagged],
            units_ban=[],
            code_ban=set(),
        )
        self.assertEqual(split_tags_on, 1)
        self.assertEqual(len(expanded_on), 2)
        self.assertEqual([row.get_value(VALUES) for row in expanded_on], [1, 0])
        self.assertEqual(
            [row.get_tags_list() for row in expanded_on],
            [["TAG-A"], ["TAG-B"]],
        )

        skipped = _position_row(code="BCC0001", qty=1, tags=["TAG-A", "TAG-B"])
        strip_loaded_tags([skipped])
        expanded_off, split_tags_off, split_no_off = split_rfp_rows(
            [skipped],
            units_ban=[],
            code_ban=set(),
        )
        self.assertEqual(split_tags_off, 0)
        self.assertEqual(split_no_off, 0)
        self.assertEqual(len(expanded_off), 1)
        self.assertEqual(expanded_off[0].get_value(VALUES), 1)
        self.assertEqual(expanded_off[0].get_tags_list(), [])

    def test_apply_load_tags_mode_true_is_noop(self) -> None:
        row = _position_row(code="BCC0001", qty=1, tags=["T1", "T2"])
        counts = apply_load_tags_mode(
            load_tags=True,
            rfp_rows=[row],
            mto_data={"tm": [row]},
            vo_data=None,
            packing_rows=None,
        )
        self.assertEqual(counts, {"rfp": 0, "mto": 0, "vo": 0, "ul": 0})
        self.assertEqual(row.get_tags_list(), ["T1", "T2"])

    def test_apply_load_tags_mode_false_clears_all_contours(self) -> None:
        rfp = _position_row(code="R", qty=1, tags=["R1", "R2"])
        mto = _position_row(code="M", qty=2, tags=["M1"])
        vo = _position_row(code="V", qty=1, tags="V1")
        ul = _position_row(code="U", qty=1, tags=["U1", "U2"])
        counts = apply_load_tags_mode(
            load_tags=False,
            rfp_rows=[rfp],
            mto_data={"tm": [mto]},
            vo_data={"vo": [vo]},
            packing_rows=[ul],
        )
        self.assertEqual(counts, {"rfp": 1, "mto": 1, "vo": 1, "ul": 1})
        self.assertEqual(rfp.get_tags_list(), [])
        self.assertEqual(mto.get_tags_list(), [])
        self.assertEqual(vo.get_tags_list(), [])
        self.assertEqual(ul.get_tags_list(), [])

    def test_orchestrator_strips_before_gate_when_disabled(self) -> None:
        captured: dict[str, object] = {}
        rfp_row = _position_row(code="BCC0001", qty=1, tags=["TAG-A", "TAG-B"])
        mto_row = _position_row(code="BCC0001", qty=1, tags=["MTO-1"])

        def _gate(*, rfp_rows, mto_data, packing_dataset, **_kwargs):
            captured["gate_rfp_tags"] = list(rfp_rows[0].get_tags_list())
            captured["gate_mto_tags"] = list(mto_data["tm"][0].get_tags_list())
            captured["packing_dataset"] = packing_dataset
            return mock.Mock(summary="ok")

        def _finalize(rfp_data, *_args, **_kwargs):
            captured["finalize_tags"] = list(rfp_data[0].get_tags_list())
            captured["finalize_values"] = rfp_data[0].get_value(VALUES)
            return rfp_data, []

        tmpdir = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(tmpdir.cleanup)
        config = {
            "load_tags": False,
            "memory_log": False,
            "paths": {
                "rfp_path": "rfp.xlsx",
                "mto_path": "",
                "vo_path": "",
                "code_ban_file": "",
                "replacement_table_file": "",
                "result_dir_base": tmpdir.name,
                "units_convert_matrix": str(Path(tmpdir.name) / "matrix.xlsx"),
            },
            "step1": {"rfp_pipeline_mode": "standard", "skip_split": False},
            "step2": {"export_positions_database_excel": False},
            "step3": {},
            "step4": {
                "include_packing_lists": False,
                "export_bcc_accum_matrix": False,
            },
            "rfp_parts": {"use_latest_net": False},
            "rfp_tags_utils": {"save_input_fingerprints": False},
        }
        os.environ.setdefault("PYTHONUTF8", "1")
        with (
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.load_config",
                return_value=config,
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.resolve_rfp_pipeline_mode",
                return_value="standard",
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.resolve_effective_rfp_path",
                return_value=mock.Mock(path="rfp.xlsx", detail="test"),
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.step1_load_rfp_raw",
                return_value=[rfp_row],
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.step2_load_mto_data",
                return_value={"tm": [mto_row]},
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.step3_load_vo_data",
                return_value={},
            ),
            mock.patch("RFQ.tags_rfp_compare.agregate_tags.load_base", return_value=[]),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.run_units_gate",
                side_effect=_gate,
            ),
            mock.patch("RFQ.tags_rfp_compare.agregate_tags.emit_milestone"),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.step1_finalize_rfp_data",
                side_effect=_finalize,
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.step4_analyze_and_match",
                return_value=[],
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.ensure_rfp_parts_net_current",
            ) as ensure_net,
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.check_ds_id_coverage",
                return_value=_ok_ds_id_coverage(),
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.collect_parts_ds_paths",
                return_value={},
            ),
            mock.patch("RFQ.tags_rfp_compare.agregate_tags.open_dir"),
            mock.patch("RFQ.tags_rfp_compare.agregate_tags.finalize_timing_log"),
        ):
            from RFQ.tags_rfp_compare.agregate_tags import main

            main(config_override=config)

        ensure_net.assert_not_called()
        self.assertEqual(captured["gate_rfp_tags"], [])
        self.assertEqual(captured["gate_mto_tags"], [])
        self.assertIsNone(captured["packing_dataset"])
        self.assertEqual(captured["finalize_tags"], [])
        self.assertEqual(captured["finalize_values"], 1)

    def test_orchestrator_preflight_passes_load_tags_false(self) -> None:
        captured: dict[str, object] = {}
        rfp_row = _position_row(code="BCC0001", qty=1, tags=["TAG-A", "TAG-B"])
        mto_row = _position_row(code="BCC0001", qty=1, tags=["MTO-1"])

        def _gate(*, rfp_rows, **_kwargs):
            captured["gate_rfp_tags"] = list(rfp_rows[0].get_tags_list())
            return mock.Mock(summary="ok")

        def _finalize(rfp_data, *_args, **_kwargs):
            return rfp_data, []

        tmpdir = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(tmpdir.cleanup)
        config = {
            "load_tags": False,
            "memory_log": False,
            "paths": {
                "rfp_path": "rfp.xlsx",
                "mto_path": "",
                "vo_path": "",
                "code_ban_file": "",
                "replacement_table_file": "",
                "result_dir_base": tmpdir.name,
                "units_convert_matrix": str(Path(tmpdir.name) / "matrix.xlsx"),
            },
            "step1": {"rfp_pipeline_mode": "standard", "skip_split": False},
            "step2": {"export_positions_database_excel": False},
            "step3": {},
            "step4": {
                "include_packing_lists": False,
                "export_bcc_accum_matrix": False,
            },
            "rfp_parts": {"use_latest_net": True},
            "rfp_tags_utils": {"save_input_fingerprints": False},
        }
        freshness = mock.Mock(needs_rebuild=True, reason="нет rfp_parts_net_no_tags.xlsx")
        os.environ.setdefault("PYTHONUTF8", "1")
        with (
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.load_config",
                return_value=config,
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.resolve_rfp_pipeline_mode",
                return_value="standard",
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.resolve_effective_rfp_path",
                return_value=mock.Mock(
                    path="rfp_parts_net_no_tags.xlsx",
                    detail="latest rfp_parts_net_no_tags.xlsx",
                ),
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.step1_load_rfp_raw",
                return_value=[rfp_row],
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.step2_load_mto_data",
                return_value={"tm": [mto_row]},
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.step3_load_vo_data",
                return_value={},
            ),
            mock.patch("RFQ.tags_rfp_compare.agregate_tags.load_base", return_value=[]),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.run_units_gate",
                side_effect=_gate,
            ),
            mock.patch("RFQ.tags_rfp_compare.agregate_tags.emit_milestone"),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.step1_finalize_rfp_data",
                side_effect=_finalize,
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.step4_analyze_and_match",
                return_value=[],
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.ensure_rfp_parts_net_current",
                return_value=(
                    Path("rfp_parts_net_no_tags.xlsx"),
                    freshness,
                ),
            ) as ensure_net,
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.check_ds_id_coverage",
                return_value=_ok_ds_id_coverage(),
            ),
            mock.patch(
                "RFQ.tags_rfp_compare.agregate_tags.collect_parts_ds_paths",
                return_value={},
            ),
            mock.patch("RFQ.tags_rfp_compare.agregate_tags.open_dir"),
            mock.patch("RFQ.tags_rfp_compare.agregate_tags.finalize_timing_log"),
        ):
            from RFQ.tags_rfp_compare.agregate_tags import main

            main(config_override=config)

        ensure_net.assert_called_once()
        self.assertFalse(ensure_net.call_args.kwargs["load_tags"])
        self.assertEqual(captured["gate_rfp_tags"], [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
