"""Headless regression tests for the native RFP settings panel."""

from __future__ import annotations

import copy
import os
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PySide6.QtWidgets import QApplication, QPushButton

from RFQ.tags_rfp_compare.rfp_tags_utils import (
    RFP_PIPELINE_MODE_WITH_ORPHAN_MTO_VO,
    get_default_asbuild_config,
    get_default_config,
)
import ds_compare_center.rfp_settings_panel as panel_module

_UNITS_CONVERT_MATRIX_DEFAULT = (
    r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\матрица_ед_изм.xlsx"
)
_DS_MANAGER_MATRIX_DEFAULT = (
    r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\Список ДС - Фамилии МП.xlsx"
)
_GEM_SUPPLY_CODES_DEFAULT = (
    r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_RFP\Коды_Поставок_ГЭМ_8950.xlsx"
)


class RfpNativeSettingsSmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self) -> None:
        self.profiles = {
            "main": get_default_config(),
            "asbuild": get_default_asbuild_config(),
        }
        for name, config in self.profiles.items():
            config["unknown_root"] = {"profile": name, "preserve": [1, 2, 3]}
            config.setdefault("column_optimization", {}).setdefault("rfp", {})[
                "hidden_custom"
            ] = {"load": False, "truncate": 777, "future": {"keep": True}}
            config["column_optimization"].setdefault("mto", {})[
                "hidden_custom"
            ] = {"load": True, "future": "keep"}
            config["rfp_parts"]["summary_path"] = f"{name}-summary.xlsx"

        self.saved_calls: list[str] = []
        self.goto_calls: list[str] = []

        def load_main() -> dict:
            return copy.deepcopy(self.profiles["main"])

        def load_asbuild() -> dict:
            return copy.deepcopy(self.profiles["asbuild"])

        def save_main(config: dict) -> bool:
            self.profiles["main"] = copy.deepcopy(config)
            return True

        def save_asbuild(config: dict) -> bool:
            self.profiles["asbuild"] = copy.deepcopy(config)
            return True

        dataset = SimpleNamespace(
            quality=SimpleNamespace(value="ok"),
            issues=[],
            format_short=lambda: "УЛ: test cache",
        )
        self.patchers = (
            mock.patch.object(panel_module, "load_config", side_effect=load_main),
            mock.patch.object(
                panel_module, "load_asbuild_config", side_effect=load_asbuild
            ),
            mock.patch.object(panel_module, "save_config", side_effect=save_main),
            mock.patch.object(
                panel_module, "save_asbuild_config", side_effect=save_asbuild
            ),
            mock.patch.object(panel_module, "load_packing_dataset", return_value=dataset),
        )
        for patcher in self.patchers:
            patcher.start()
        self.panel = panel_module.RfpSettingsPanel(
            on_saved=lambda: self.saved_calls.append("saved"),
            on_goto_packing=lambda: self.goto_calls.append("goto"),
        )

    def tearDown(self) -> None:
        thread = getattr(self.panel, "_packing_thread", None)
        if thread is not None and thread.isRunning():
            thread.quit()
            thread.wait(5000)
        self.panel.close()
        self.panel.deleteLater()
        self.app.processEvents()
        for patcher in reversed(self.patchers):
            patcher.stop()

    def _select_profile(self, profile: str) -> None:
        index = self.panel._profile.findData(profile)
        self.assertGreaterEqual(index, 0)
        self.panel._profile.setCurrentIndex(index)
        self.app.processEvents()

    def test_both_profiles_round_trip_unknown_and_hidden_config(self) -> None:
        for profile in ("main", "asbuild"):
            with self.subTest(profile=profile):
                self._select_profile(profile)
                collected = self.panel.collect_config()
                self.assertEqual(
                    collected["unknown_root"],
                    {"profile": profile, "preserve": [1, 2, 3]},
                )
                self.assertEqual(
                    collected["column_optimization"]["rfp"]["hidden_custom"],
                    {"load": False, "truncate": 777, "future": {"keep": True}},
                )
                self.assertEqual(
                    collected["column_optimization"]["mto"]["hidden_custom"],
                    {"load": True, "future": "keep"},
                )

                self.assertNotIn("rfp_parts.summary_path", self.panel._widgets)
                self.assertEqual(
                    collected["rfp_parts"]["summary_path"],
                    f"{profile}-summary.xlsx",
                )
                lot = self.panel._widgets["paths.rfp_registr_lot_path"]
                lot.setText(f"{profile}-lot.xlsx")
                self.assertTrue(self.panel.save_to_disk())
                lot.setText("unsaved-value.xlsx")
                self.panel.reload_from_disk()
                self.assertEqual(lot.text(), f"{profile}-lot.xlsx")
                self.assertEqual(
                    self.profiles[profile]["rfp_parts"]["summary_path"],
                    f"{profile}-summary.xlsx",
                )
                self.assertEqual(
                    self.profiles[profile]["unknown_root"]["profile"], profile
                )

        self.assertEqual(self.saved_calls, ["saved", "saved"])

    def test_pipeline_dependency_asbuild_flags_and_callbacks(self) -> None:
        self._select_profile("main")
        nonstandard = self.panel._pipeline.findData(
            RFP_PIPELINE_MODE_WITH_ORPHAN_MTO_VO
        )
        self.panel._pipeline.setCurrentIndex(nonstandard)
        orphan = self.panel._widgets[
            "step4.include_mto_vo_without_rfp_anchor"
        ]
        self.assertTrue(orphan.isChecked())
        self.assertFalse(orphan.isEnabled())
        self.assertTrue(
            self.panel.collect_config()["step4"]["include_mto_vo_without_rfp_anchor"]
        )

        self._select_profile("asbuild")
        config = self.panel.collect_config()
        self.assertTrue(config["step2"]["flat_mto_structure"])
        self.assertTrue(config["step4"]["filter_to_mto_titles"])
        self.assertFalse(config["step4"]["include_packing_lists"])
        self.assertFalse(config["step4"]["assign_rfp_mto_code_compare_colors"])
        self.assertFalse(config["rfp_parts"]["auto_update_checklist"])
        self.assertFalse(config["rfp_parts"]["use_latest_net"])
        latest_net = self.panel._widgets["rfp_parts.use_latest_net"]
        self.assertFalse(latest_net.isChecked())
        self.assertFalse(latest_net.isEnabled())

        goto = next(
            button
            for button in self.panel.findChildren(QPushButton)
            if button.text() == "Перейти к упаковочным листам"
        )
        goto.click()
        self.app.processEvents()
        self.assertEqual(self.goto_calls, ["goto"])

    def test_latest_net_widget_maps_to_rfp_parts_use_latest_net(self) -> None:
        self.assertIn("rfp_parts.use_latest_net", self.panel._widgets)
        self.assertIn("paths.rfp_path", self.panel._widgets)
        widget = self.panel._widgets["rfp_parts.use_latest_net"]
        self.assertTrue(widget.isChecked())
        widget.setChecked(False)
        collected = self.panel.collect_config()
        self.assertFalse(collected["rfp_parts"]["use_latest_net"])
        self.assertTrue(self.panel.save_to_disk())
        widget.setChecked(True)
        self.panel.reload_from_disk()
        self.assertFalse(widget.isChecked())
        self.assertFalse(self.profiles["main"]["rfp_parts"]["use_latest_net"])

    def test_units_convert_matrix_default_main_asbuild_and_widget(self) -> None:
        main_default = get_default_config()
        asbuild_default = get_default_asbuild_config()
        self.assertEqual(
            main_default["paths"]["units_convert_matrix"],
            _UNITS_CONVERT_MATRIX_DEFAULT,
        )
        self.assertEqual(
            asbuild_default["paths"]["units_convert_matrix"],
            _UNITS_CONVERT_MATRIX_DEFAULT,
        )

        self.assertIn("paths.units_convert_matrix", self.panel._widgets)
        widget = self.panel._widgets["paths.units_convert_matrix"]
        self.assertEqual(widget.text(), _UNITS_CONVERT_MATRIX_DEFAULT)

        for profile in ("main", "asbuild"):
            with self.subTest(profile=profile):
                self._select_profile(profile)
                collected = self.panel.collect_config()
                self.assertEqual(
                    collected["paths"]["units_convert_matrix"],
                    _UNITS_CONVERT_MATRIX_DEFAULT,
                )
                widget.setText(f"{profile}-matrix.xlsx")
                self.assertTrue(self.panel.save_to_disk())
                widget.setText("unsaved-matrix.xlsx")
                self.panel.reload_from_disk()
                self.assertEqual(widget.text(), f"{profile}-matrix.xlsx")
                self.assertEqual(
                    self.profiles[profile]["paths"]["units_convert_matrix"],
                    f"{profile}-matrix.xlsx",
                )

        field_keys = {
            spec.key
            for section in panel_module._SECTIONS
            for spec in section[1]
        }
        self.assertIn("paths.units_convert_matrix", field_keys)
        self.assertIn("step4.ul_match_use_mto_tags", field_keys)
        self.assertTrue(main_default["step4"]["ul_match_use_mto_tags"])
        self.assertTrue(asbuild_default["step4"]["ul_match_use_mto_tags"])
        self.assertIn("step4.ul_match_use_mto_tags", self.panel._widgets)
        self.assertTrue(
            self.panel._widgets["step4.ul_match_use_mto_tags"].isChecked()
        )
        self.assertTrue(main_default["load_tags"])
        self.assertTrue(asbuild_default["load_tags"])
        self.assertIn("load_tags", field_keys)
        self.assertIn("load_tags", self.panel._widgets)
        self.assertTrue(self.panel._widgets["load_tags"].isChecked())

    def test_ds_manager_matrix_default_main_asbuild_and_widget(self) -> None:
        main_default = get_default_config()
        asbuild_default = get_default_asbuild_config()
        self.assertEqual(
            main_default["paths"]["ds_manager_matrix"],
            _DS_MANAGER_MATRIX_DEFAULT,
        )
        self.assertEqual(
            asbuild_default["paths"]["ds_manager_matrix"],
            _DS_MANAGER_MATRIX_DEFAULT,
        )

        self.assertIn("paths.ds_manager_matrix", self.panel._widgets)
        widget = self.panel._widgets["paths.ds_manager_matrix"]
        self.assertEqual(widget.text(), _DS_MANAGER_MATRIX_DEFAULT)

        field_keys = {
            spec.key
            for section in panel_module._SECTIONS
            for spec in section[1]
        }
        self.assertIn("paths.ds_manager_matrix", field_keys)

        for profile in ("main", "asbuild"):
            with self.subTest(profile=profile):
                self._select_profile(profile)
                collected = self.panel.collect_config()
                self.assertEqual(
                    collected["paths"]["ds_manager_matrix"],
                    _DS_MANAGER_MATRIX_DEFAULT,
                )

    def test_gem_supply_codes_default_main_asbuild_and_widget(self) -> None:
        main_default = get_default_config()
        asbuild_default = get_default_asbuild_config()
        self.assertEqual(
            main_default["paths"]["gem_supply_codes"],
            _GEM_SUPPLY_CODES_DEFAULT,
        )
        self.assertEqual(
            asbuild_default["paths"]["gem_supply_codes"],
            _GEM_SUPPLY_CODES_DEFAULT,
        )

        self.assertIn("paths.gem_supply_codes", self.panel._widgets)
        widget = self.panel._widgets["paths.gem_supply_codes"]
        self.assertEqual(widget.text(), _GEM_SUPPLY_CODES_DEFAULT)

        field_keys = {
            spec.key
            for section in panel_module._SECTIONS
            for spec in section[1]
        }
        self.assertIn("paths.gem_supply_codes", field_keys)

        for profile in ("main", "asbuild"):
            with self.subTest(profile=profile):
                self._select_profile(profile)
                collected = self.panel.collect_config()
                self.assertEqual(
                    collected["paths"]["gem_supply_codes"],
                    _GEM_SUPPLY_CODES_DEFAULT,
                )

    def test_summary_path_has_no_widget_and_is_preserved(self) -> None:
        self.assertNotIn("rfp_parts.summary_path", self.panel._widgets)
        self.assertNotIn("rfp_parts.auto_update_checklist", self.panel._widgets)
        self.assertNotIn("paths.summary_path", self.panel._widgets)

        collected = self.panel.collect_config()

        self.assertEqual(
            collected["rfp_parts"]["summary_path"], "main-summary.xlsx"
        )
        self.assertNotIn("summary_path", collected["paths"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
