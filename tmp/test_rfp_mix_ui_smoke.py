"""Smoke for independent RFP mix-mode JSON keys and button labels."""

from __future__ import annotations

import copy
import os
import sys
import unittest
from pathlib import Path
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.rfp_parts.ds_hybrid_preflight import (
    INPUT_MODE_DS_ONLY,
    INPUT_MODE_HYBRID,
    INPUT_MODE_LEGACY_NET,
    resolve_launch_mix_mode,
)
from RFQ.rfp_parts.ds_rfp_tag_placement import MIX_MIXED, MIX_SEPARATE
from ds_compare_center.rfp_mix_settings import (
    COLLECT_JOB_DS_ONLY,
    COLLECT_JOB_HYBRID,
    COLLECT_JOB_PARTS,
    KEY_COLLECT_JOB,
    KEY_COLLECT_MIX_MODE,
    KEY_LAUNCH_MIX_MODE,
    format_collect_run_button,
    format_launch_rfp_button,
    read_collect_job,
    read_collect_mix_mode,
    read_launch_mix_mode,
    set_rfp_parts_key,
)


class RfpMixSettingsSmokeTest(unittest.TestCase):
    def test_collect_button_label_from_job_and_mix(self) -> None:
        self.assertEqual(
            format_collect_run_button(COLLECT_JOB_HYBRID, MIX_SEPARATE),
            "Запуск (ДС+RFP · без смешения)",
        )
        self.assertEqual(
            format_collect_run_button(COLLECT_JOB_HYBRID, MIX_MIXED),
            "Запуск (ДС+RFP · со смешением)",
        )
        self.assertEqual(
            format_collect_run_button(COLLECT_JOB_PARTS, MIX_MIXED),
            "Запуск (части RFP)",
        )
        self.assertEqual(
            format_collect_run_button(COLLECT_JOB_DS_ONLY, MIX_MIXED),
            "Запуск (только ДС)",
        )
        self.assertEqual(
            format_collect_run_button("nope", None),
            "Запуск (ДС+RFP · без смешения)",
        )

    def test_launch_button_mix_suffix_only_for_hybrid(self) -> None:
        packing = format_launch_rfp_button(
            include_packing=True,
            input_mode=INPUT_MODE_HYBRID,
            mix=MIX_SEPARATE,
        )
        self.assertEqual(packing, "Сравнить RFP ↔ MTO ↔ РКД ↔ УЛ (без смешения)")
        mixed = format_launch_rfp_button(
            include_packing=False,
            input_mode=INPUT_MODE_HYBRID,
            mix=MIX_MIXED,
        )
        self.assertEqual(mixed, "Сравнить RFP ↔ MTO ↔ РКД (со смешением)")
        parts = format_launch_rfp_button(
            include_packing=True,
            input_mode=INPUT_MODE_LEGACY_NET,
            mix=MIX_MIXED,
        )
        self.assertEqual(parts, "Сравнить RFP ↔ MTO ↔ РКД ↔ УЛ")
        ds_only = format_launch_rfp_button(
            include_packing=False,
            input_mode=INPUT_MODE_DS_ONLY,
            mix=MIX_MIXED,
        )
        self.assertEqual(ds_only, "Сравнить RFP ↔ MTO ↔ РКД")

    def test_two_json_keys_read_independently_without_cross_write(self) -> None:
        config = {
            "rfp_parts": {
                KEY_COLLECT_MIX_MODE: MIX_MIXED,
                KEY_LAUNCH_MIX_MODE: MIX_SEPARATE,
                KEY_COLLECT_JOB: COLLECT_JOB_PARTS,
                "input_mode": INPUT_MODE_HYBRID,
            }
        }
        self.assertEqual(read_collect_mix_mode(config), MIX_MIXED)
        self.assertEqual(read_launch_mix_mode(config), MIX_SEPARATE)
        self.assertEqual(read_collect_job(config), COLLECT_JOB_PARTS)
        self.assertEqual(resolve_launch_mix_mode(config), MIX_SEPARATE)

        set_rfp_parts_key(config, KEY_COLLECT_MIX_MODE, MIX_SEPARATE)
        self.assertEqual(config["rfp_parts"][KEY_COLLECT_MIX_MODE], MIX_SEPARATE)
        self.assertEqual(config["rfp_parts"][KEY_LAUNCH_MIX_MODE], MIX_SEPARATE)
        self.assertEqual(config["rfp_parts"][KEY_COLLECT_JOB], COLLECT_JOB_PARTS)
        self.assertEqual(config["rfp_parts"]["input_mode"], INPUT_MODE_HYBRID)

        set_rfp_parts_key(config, KEY_LAUNCH_MIX_MODE, MIX_MIXED)
        self.assertEqual(config["rfp_parts"][KEY_LAUNCH_MIX_MODE], MIX_MIXED)
        self.assertEqual(config["rfp_parts"][KEY_COLLECT_MIX_MODE], MIX_SEPARATE)
        self.assertEqual(read_collect_mix_mode(config), MIX_SEPARATE)
        self.assertEqual(read_launch_mix_mode(config), MIX_MIXED)

    def test_missing_keys_default_without_inventing_the_other(self) -> None:
        self.assertEqual(read_collect_mix_mode({}), MIX_SEPARATE)
        self.assertEqual(read_launch_mix_mode({}), MIX_SEPARATE)
        self.assertEqual(read_collect_job({}), COLLECT_JOB_HYBRID)
        self.assertEqual(read_collect_mix_mode({"rfp_parts": {}}), MIX_SEPARATE)
        self.assertEqual(
            resolve_launch_mix_mode(
                {"rfp_parts": {KEY_COLLECT_MIX_MODE: MIX_MIXED}}
            ),
            MIX_SEPARATE,
        )
        only_collect = {"rfp_parts": {KEY_COLLECT_MIX_MODE: MIX_MIXED}}
        set_rfp_parts_key(only_collect, KEY_COLLECT_JOB, COLLECT_JOB_DS_ONLY)
        self.assertEqual(only_collect["rfp_parts"][KEY_COLLECT_MIX_MODE], MIX_MIXED)
        self.assertNotIn(KEY_LAUNCH_MIX_MODE, only_collect["rfp_parts"])

    def test_launch_resolver_ignores_collect_mix_mode(self) -> None:
        config = {
            "rfp_parts": {
                KEY_COLLECT_MIX_MODE: MIX_MIXED,
                KEY_LAUNCH_MIX_MODE: MIX_SEPARATE,
            }
        }
        self.assertEqual(resolve_launch_mix_mode(config), MIX_SEPARATE)
        self.assertEqual(resolve_launch_mix_mode({}), MIX_SEPARATE)
        self.assertEqual(
            resolve_launch_mix_mode({"rfp_parts": {"launch_mix_mode": "nope"}}),
            MIX_SEPARATE,
        )


class RfpMixWidgetSmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        from PySide6.QtWidgets import QApplication

        cls.app = QApplication.instance() or QApplication([])

    def test_collect_block_disables_mix_when_job_is_not_hybrid(self) -> None:
        from ds_compare_center import rfp_mix_blocks as mix_module
        from ds_compare_center.rfp_mix_blocks import RfpCollectJobBlock

        config = {
            "rfp_parts": {
                KEY_COLLECT_JOB: COLLECT_JOB_HYBRID,
                KEY_COLLECT_MIX_MODE: MIX_SEPARATE,
                KEY_LAUNCH_MIX_MODE: MIX_MIXED,
                "input_mode": INPUT_MODE_LEGACY_NET,
            }
        }
        saved: list[dict] = []

        def load() -> dict:
            return copy.deepcopy(config)

        def save(cfg: dict) -> bool:
            saved.append(copy.deepcopy(cfg))
            config.clear()
            config.update(copy.deepcopy(cfg))
            return True

        with (
            mock.patch.object(mix_module, "load_config", side_effect=load),
            mock.patch.object(mix_module, "save_config", side_effect=save),
        ):
            block = RfpCollectJobBlock()
            block.apply_config(config)
            self.app.processEvents()
            self.assertTrue(block._mix_radios[MIX_SEPARATE].isEnabled())
            block._job_radios[COLLECT_JOB_PARTS].click()
            self.app.processEvents()
            self.assertFalse(block._mix_radios[MIX_SEPARATE].isEnabled())
            self.assertFalse(block._mix_radios[MIX_MIXED].isEnabled())
            self.assertEqual(block.current_job(), COLLECT_JOB_PARTS)
            self.assertEqual(config["rfp_parts"][KEY_COLLECT_JOB], COLLECT_JOB_PARTS)
            self.assertEqual(config["rfp_parts"][KEY_LAUNCH_MIX_MODE], MIX_MIXED)
            self.assertEqual(config["rfp_parts"][KEY_COLLECT_MIX_MODE], MIX_SEPARATE)
            self.assertEqual(config["rfp_parts"]["input_mode"], INPUT_MODE_LEGACY_NET)
            block.close()
            block.deleteLater()
            self.app.processEvents()

    def test_launch_mix_block_writes_only_launch_key(self) -> None:
        from ds_compare_center import rfp_mix_blocks as mix_module
        from ds_compare_center.rfp_mix_blocks import RfpMixModeBlock

        config = {
            "rfp_parts": {
                KEY_COLLECT_MIX_MODE: MIX_MIXED,
                KEY_LAUNCH_MIX_MODE: MIX_SEPARATE,
                KEY_COLLECT_JOB: COLLECT_JOB_HYBRID,
            }
        }

        def load() -> dict:
            return copy.deepcopy(config)

        def save(cfg: dict) -> bool:
            config.clear()
            config.update(copy.deepcopy(cfg))
            return True

        with (
            mock.patch.object(mix_module, "load_config", side_effect=load),
            mock.patch.object(mix_module, "save_config", side_effect=save),
        ):
            block = RfpMixModeBlock()
            block.apply_config(config)
            self.app.processEvents()
            block._radios[MIX_MIXED].click()
            self.app.processEvents()
            self.assertEqual(config["rfp_parts"][KEY_LAUNCH_MIX_MODE], MIX_MIXED)
            self.assertEqual(config["rfp_parts"][KEY_COLLECT_MIX_MODE], MIX_MIXED)
            self.assertEqual(config["rfp_parts"][KEY_COLLECT_JOB], COLLECT_JOB_HYBRID)
            block.close()
            block.deleteLater()
            self.app.processEvents()


if __name__ == "__main__":
    unittest.main(verbosity=2)
