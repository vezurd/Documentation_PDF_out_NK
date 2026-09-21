"""Headless smoke for CursorBind VPN status panel."""

from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PySide6.QtWidgets import QApplication

from ds_compare_center.vpn_panel import VpnPanel
from ds_compare_center.vpn_status import (
    CURSORBIND_IP,
    evaluate_vpn_raw,
    ips_look_correct,
)


class VpnStatusSmokeTest(unittest.TestCase):
    def test_evaluate_ok_stack(self) -> None:
        snap = evaluate_vpn_raw(
            {
                "addrs": [
                    {"InterfaceAlias": "CursorBind", "IPAddress": CURSORBIND_IP},
                    {"InterfaceAlias": "Ethernet", "IPAddress": "172.16.3.235"},
                ],
                "defaults": [
                    {
                        "InterfaceAlias": "Ethernet",
                        "NextHop": "172.16.3.1",
                        "RouteMetric": 5,
                    },
                ],
                "lan": [
                    {
                        "InterfaceAlias": "Ethernet",
                        "NextHop": "172.16.3.1",
                        "RouteMetric": 1,
                    },
                ],
                "socks": True,
                "proc3proxy": True,
                "procProxiFyre": True,
                "unc": True,
                "bat_exists": True,
            }
        )
        self.assertTrue(snap.overall_ok)
        self.assertTrue(all(row.ok for row in snap.checks))

    def test_evaluate_h10_fails(self) -> None:
        snap = evaluate_vpn_raw(
            {
                "addrs": [
                    {
                        "InterfaceAlias": "NetherlandsAmsterdamH10",
                        "IPAddress": "10.73.104.145",
                    },
                ],
                "defaults": [
                    {
                        "InterfaceAlias": "NetherlandsAmsterdamH10",
                        "NextHop": "0.0.0.0",
                        "RouteMetric": 0,
                    },
                    {
                        "InterfaceAlias": "Ethernet",
                        "NextHop": "172.16.3.1",
                        "RouteMetric": 5,
                    },
                ],
                "lan": [],
                "socks": False,
                "proc3proxy": False,
                "procProxiFyre": False,
                "unc": False,
                "bat_exists": True,
            }
        )
        self.assertFalse(snap.overall_ok)
        h10 = next(row for row in snap.checks if row.name == "H10 выключен")
        self.assertFalse(h10.ok)

    def test_ips_look_correct(self) -> None:
        ok, _ = ips_look_correct("84.204.145.165", "31.58.51.202")
        self.assertTrue(ok)
        bad, _ = ips_look_correct("195.218.230.62", "31.58.51.202")
        self.assertFalse(bad)


class VpnPanelSmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_panel_builds(self) -> None:
        panel = VpnPanel()
        self.app.processEvents()
        self.assertIsNotNone(panel._table)
        self.assertGreaterEqual(panel._table.columnCount(), 3)
        panel.close()
        panel.deleteLater()
        self.app.processEvents()


if __name__ == "__main__":
    unittest.main(verbosity=2)
