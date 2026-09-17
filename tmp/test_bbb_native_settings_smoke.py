"""Headless regression tests for the native MTO/BBB settings panel."""

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

from PySide6.QtWidgets import QApplication, QPushButton

from base.bbb_config import get_default_config
import ds_compare_center.bbb_settings_panel as panel_module


_EXPECTED_KEYS = (
    "folder_rules.search_only_in_dwg",
    "mto_vs_code_base.enabled",
    "mto_vs_code_base.check_by_code",
    "mto_vs_code_base.check_equipment_codes",
    "mto_vs_code_base.check_tags_value",
    "mto_vs_code_base.check_value",
    "mto_vs_code_base.check_duplicate_tags",
    "mto_vs_code_base.check_mass",
    "mto_vs_code_base.check_prohibition",
    "mto_vs_code_base.check_tags_4_2_4_4",
    "mto_vs_code_base.check_position_numeration",
    "mto_vs_code_base_correction.enabled",
    "mto_vs_code_base_correction.check_by_code",
    "mto_vs_code_base_correction.check_tags_value",
    "mto_vs_code_base_correction.check_value",
    "mto_vs_code_base_correction.check_duplicate_tags",
    "mto_vs_code_base_correction.check_tags_4_2_4_4",
    "mto_vs_code_base_correction.check_position_numeration",
    "bbb_vs_code_base.enabled",
    "bbb_vs_code_base.check_by_code",
    "bbb_vs_code_base.check_equipment_codes",
    "bbb_vs_code_base.check_tags_value",
    "bbb_vs_code_base.check_value",
    "bbb_vs_code_base.check_duplicate_tags",
    "bbb_vs_code_base.check_mass",
    "bbb_vs_code_base.check_prohibition",
    "bbb_vs_code_base.check_work_code_and_mtr_group",
    "bbb_vs_mto.enabled",
    "bbb_vs_mto.compare_fields",
    "bbb_vs_mto.compare_values",
    "bbb_vs_mto.compare_tags",
    "bbb_vs_mto.compare_title_marka_revision",
    "bbb_vs_mto.compare_bom_annotation_with_mto",
    "bbb_vs_mto.check_missing_mto_codes",
)


class BbbNativeSettingsSmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self) -> None:
        self.profile = get_default_config()
        self.profile["unknown_root"] = {"preserve": [1, 2, 3], "future": True}
        self.profile["mto_vs_code_base"]["hidden_custom"] = {"keep": True}
        self.saved_calls: list[str] = []

        def load_cfg() -> dict:
            return copy.deepcopy(self.profile)

        def save_cfg(config: dict) -> bool:
            self.profile = copy.deepcopy(config)
            return True

        self.patchers = (
            mock.patch.object(panel_module, "load_config", side_effect=load_cfg),
            mock.patch.object(panel_module, "save_config", side_effect=save_cfg),
        )
        for patcher in self.patchers:
            patcher.start()
        self.panel = panel_module.BbbSettingsPanel(
            on_saved=lambda: self.saved_calls.append("saved"),
        )

    def tearDown(self) -> None:
        self.panel.close()
        self.panel.deleteLater()
        self.app.processEvents()
        for patcher in reversed(self.patchers):
            patcher.stop()

    def test_checkbox_keys_match_ctk_schema(self) -> None:
        self.assertEqual(set(self.panel._widgets), set(_EXPECTED_KEYS))

    def test_round_trip_preserves_unknown_and_writes_flags(self) -> None:
        collected = self.panel.collect_config()
        self.assertEqual(
            collected["unknown_root"],
            {"preserve": [1, 2, 3], "future": True},
        )
        self.assertEqual(
            collected["mto_vs_code_base"]["hidden_custom"],
            {"keep": True},
        )
        self.assertFalse(collected["mto_vs_code_base"]["check_prohibition"])
        self.assertTrue(collected["folder_rules"]["search_only_in_dwg"])

        widget = self.panel._widgets["mto_vs_code_base.check_prohibition"]
        widget.setChecked(True)
        dwg = self.panel._widgets["folder_rules.search_only_in_dwg"]
        dwg.setChecked(False)
        self.assertTrue(self.panel.save_to_disk())

        self.assertEqual(self.saved_calls, ["saved"])
        self.assertTrue(self.profile["mto_vs_code_base"]["check_prohibition"])
        self.assertFalse(self.profile["folder_rules"]["search_only_in_dwg"])
        self.assertEqual(self.profile["unknown_root"]["preserve"], [1, 2, 3])
        self.assertEqual(
            self.profile["mto_vs_code_base"]["hidden_custom"],
            {"keep": True},
        )

        widget.setChecked(False)
        dwg.setChecked(True)
        self.panel.reload_from_disk()
        self.assertTrue(
            self.panel._widgets["mto_vs_code_base.check_prohibition"].isChecked()
        )
        self.assertFalse(
            self.panel._widgets["folder_rules.search_only_in_dwg"].isChecked()
        )

    def test_section_types_button_present(self) -> None:
        buttons = [
            btn
            for btn in self.panel.findChildren(QPushButton)
            if btn.text() == "Открыть типы секций MTO"
        ]
        self.assertEqual(len(buttons), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
