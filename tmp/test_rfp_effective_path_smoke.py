"""Offline tests for Step1 RFP source: latest parts net vs paths.rfp_path."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.tags_rfp_compare.rfp_tags_utils import (
    get_default_asbuild_config,
    get_default_config,
    resolve_effective_rfp_path,
)


class EffectiveRfpPathSmokeTest(unittest.TestCase):
    def test_defaults_main_on_asbuild_off(self) -> None:
        default = get_default_config()
        self.assertTrue(default["rfp_parts"]["use_latest_net"])
        self.assertEqual(default["rfp_parts"]["input_mode"], "legacy_net")
        self.assertEqual(default["rfp_parts"]["ds_source_dir"], "")
        self.assertTrue(default["rfp_parts"]["ds_registry_path"])
        self.assertFalse(get_default_asbuild_config()["rfp_parts"]["use_latest_net"])
        self.assertEqual(
            get_default_asbuild_config()["rfp_parts"]["input_mode"], "legacy_net"
        )

    def test_uses_paths_when_flag_off(self) -> None:
        resolved = resolve_effective_rfp_path(
            {
                "paths": {"rfp_path": r"C:\fixed.xlsx"},
                "rfp_parts": {"use_latest_net": False},
            }
        )
        self.assertFalse(resolved.used_latest_net)
        self.assertEqual(resolved.path, r"C:\fixed.xlsx")
        self.assertEqual(resolved.detail, "paths.rfp_path")

    def test_missing_fallback_raises(self) -> None:
        with self.assertRaises(FileNotFoundError):
            resolve_effective_rfp_path(
                {
                    "paths": {"rfp_path": "  "},
                    "rfp_parts": {"use_latest_net": False},
                }
            )

    @mock.patch(
        "RFQ.rfp_parts.analyze_rfp_parts.resolve_latest_rfp_parts_net_xlsx"
    )
    def test_uses_latest_net_when_flag_on(self, mocked) -> None:
        latest = Path(r"\\bcc\eng\RFP сводный файл\2026.08.18_14.00") / (
            "rfp_parts_net.xlsx"
        )
        mocked.return_value = latest
        resolved = resolve_effective_rfp_path(
            {
                "paths": {"rfp_path": r"C:\old.xlsx"},
                "rfp_parts": {"use_latest_net": True},
            }
        )
        self.assertTrue(resolved.used_latest_net)
        self.assertEqual(resolved.path, str(latest))

    @mock.patch(
        "RFQ.rfp_parts.analyze_rfp_parts.resolve_latest_rfp_parts_net_xlsx"
    )
    def test_missing_latest_raises(self, mocked) -> None:
        mocked.return_value = None
        with self.assertRaises(FileNotFoundError):
            resolve_effective_rfp_path({"rfp_parts": {"use_latest_net": True}})

    @mock.patch(
        "RFQ.rfp_parts.analyze_rfp_parts.resolve_latest_rfp_parts_net_no_tags_xlsx"
    )
    def test_load_tags_false_uses_no_tags_net(self, mocked) -> None:
        latest = Path(r"\\bcc\eng\RFP сводный файл\2026.09.04_16.00") / (
            "rfp_parts_net_no_tags.xlsx"
        )
        mocked.return_value = latest
        resolved = resolve_effective_rfp_path(
            {
                "load_tags": False,
                "paths": {"rfp_path": r"C:\old.xlsx"},
                "rfp_parts": {"use_latest_net": True},
            }
        )
        self.assertTrue(resolved.used_latest_net)
        self.assertEqual(resolved.path, str(latest))
        self.assertEqual(resolved.detail, "latest rfp_parts_net_no_tags.xlsx")

    @mock.patch(
        "RFQ.rfp_parts.analyze_rfp_parts.resolve_latest_rfp_parts_net_no_tags_xlsx"
    )
    def test_load_tags_false_missing_no_tags_raises(self, mocked) -> None:
        mocked.return_value = None
        with self.assertRaises(FileNotFoundError) as ctx:
            resolve_effective_rfp_path(
                {
                    "load_tags": False,
                    "rfp_parts": {"use_latest_net": True},
                }
            )
        self.assertIn("rfp_parts_net_no_tags.xlsx", str(ctx.exception))

    def test_flag_off_ignores_hybrid_input_mode(self) -> None:
        resolved = resolve_effective_rfp_path(
            {
                "paths": {"rfp_path": r"C:\fixed.xlsx"},
                "rfp_parts": {
                    "use_latest_net": False,
                    "input_mode": "hybrid",
                },
            }
        )
        self.assertFalse(resolved.used_latest_net)
        self.assertEqual(resolved.path, r"C:\fixed.xlsx")

    @mock.patch("RFQ.rfp_parts.ds_hybrid_preflight.resolve_ds_baseline_xlsx")
    def test_ds_only_uses_baseline_stamp(self, mocked) -> None:
        latest = Path(r"\\bcc\eng\RFP сводный файл\_ds_baseline") / (
            "Свод ДС для запуска.xlsx"
        )
        mocked.return_value = latest
        resolved = resolve_effective_rfp_path(
            {
                "paths": {"rfp_path": r"C:\old.xlsx"},
                "rfp_parts": {
                    "use_latest_net": True,
                    "input_mode": "ds_only",
                },
            }
        )
        self.assertTrue(resolved.used_latest_net)
        self.assertEqual(resolved.path, str(latest))
        self.assertEqual(resolved.detail, "latest Свод ДС для запуска.xlsx")

    @mock.patch(
        "RFQ.rfp_parts.analyze_rfp_parts.resolve_latest_rfp_parts_net_xlsx"
    )
    @mock.patch(
        "RFQ.rfp_parts.ds_hybrid_preflight.resolve_ds_hybrid_xlsx",
        return_value=None,
    )
    def test_hybrid_missing_does_not_fallback_to_parts_net(
        self, _hybrid, parts_net
    ) -> None:
        with self.assertRaises(FileNotFoundError) as ctx:
            resolve_effective_rfp_path(
                {
                    "paths": {"rfp_path": r"C:\old.xlsx"},
                    "rfp_parts": {
                        "use_latest_net": True,
                        "input_mode": "hybrid",
                    },
                }
            )
        message = str(ctx.exception)
        self.assertIn("Свод ДС-RFP для запуска.xlsx", message)
        self.assertIn("rfp_parts_net.xlsx", message)
        parts_net.assert_not_called()

    @mock.patch("RFQ.rfp_parts.ds_hybrid_preflight.resolve_ds_hybrid_xlsx")
    def test_hybrid_uses_hybrid_stamp(self, mocked) -> None:
        latest = Path(r"\\bcc\eng\RFP сводный файл\_ds_hybrid") / (
            "Свод ДС-RFP для запуска.xlsx"
        )
        mocked.return_value = latest
        resolved = resolve_effective_rfp_path(
            {
                "paths": {"rfp_path": r"C:\old.xlsx"},
                "rfp_parts": {
                    "use_latest_net": True,
                    "input_mode": "hybrid",
                },
            }
        )
        self.assertTrue(resolved.used_latest_net)
        self.assertEqual(resolved.path, str(latest))
        self.assertEqual(resolved.detail, "latest Свод ДС-RFP для запуска.xlsx")

    @mock.patch(
        "RFQ.rfp_parts.analyze_rfp_parts.resolve_latest_rfp_parts_net_xlsx"
    )
    def test_unknown_input_mode_stays_legacy(self, mocked) -> None:
        latest = Path(r"\\bcc\eng\RFP сводный файл\2026.08.18_14.00") / (
            "rfp_parts_net.xlsx"
        )
        mocked.return_value = latest
        resolved = resolve_effective_rfp_path(
            {
                "paths": {"rfp_path": r"C:\old.xlsx"},
                "rfp_parts": {
                    "use_latest_net": True,
                    "input_mode": "not_a_mode",
                },
            }
        )
        self.assertTrue(resolved.used_latest_net)
        self.assertEqual(resolved.path, str(latest))
        self.assertEqual(resolved.detail, "latest rfp_parts_net.xlsx")


if __name__ == "__main__":
    unittest.main()
