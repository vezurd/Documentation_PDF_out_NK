"""Offline smoke for DS/hybrid Launch preflight (no UNC writes)."""

from __future__ import annotations

import io
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.rfp_parts.ds_baseline import (
    BASELINE_XLSX_NAME,
    QUALITY_REPORT_PREFIX,
    STRUCTURE_REPORT_PREFIX,
    IdentityDsUnitsConverter,
)
from RFQ.rfp_parts.ds_hybrid_preflight import (
    DS_BASELINE_DIR_NAME,
    DS_HYBRID_DIR_NAME,
    INPUT_MODE_DS_ONLY,
    INPUT_MODE_HYBRID,
    INPUT_MODE_LEGACY_NET,
    STATE_JSON_NAME,
    DsBaselineBlockedError,
    copy_ds_sidecars_to_result_dir,
    ds_baseline_output_dir,
    ds_hybrid_output_dir,
    ensure_ds_baseline_current,
    ensure_ds_hybrid_current,
    resolve_ds_baseline_xlsx,
    resolve_ds_hybrid_xlsx,
    resolve_input_mode,
)
from RFQ.rfp_parts.ds_registry import (
    MODE_WHOLE,
    DsRegistryRelation,
    DsRegistryRow,
    STATUS_ACTIVE,
    write_registry_workbook,
)
from RFQ.rfp_parts.ds_rfp_hybrid import HYBRID_XLSX_NAME, IdentityRfpUnitsConverter
from RFQ.units_convert.models import GoogleUnitsIndex

DS_HEADER = [
    "№ п/п",
    "Титул",
    "Раздел",
    "Спецификация",
    "RFQ",
    "Наименование Позиций Товара по РД",
    "Код 1С СОУ",
    "Код РД",
    "Наименование Позиций Товара Поставщика",
    "Технические требования (ГОСТ/ ТУ и др.)",
    "Ед. изм.",
    "Кол-во",
]
CODE = "BCC0000001"


def _google_index() -> GoogleUnitsIndex:
    return GoogleUnitsIndex(
        units_by_code={CODE: "шт"},
        display_units_by_code={CODE: "шт"},
        display_code_by_code={CODE: CODE},
    )


def _write_xlsx(path: Path, rows: list[list[object]], *, sheet: str = "Спека") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = sheet
    for row in rows:
        ws.append(row)
    wb.save(path)
    wb.close()


def _ds_row(*, npp: object = 1, qty: object = 2, code: str = CODE) -> list[object]:
    return [
        npp,
        "8529",
        "SOS",
        "spec-1",
        "RFQ-1",
        "Кабель",
        "1C-1",
        code,
        "Vendor",
        "NYM",
        "шт",
        qty,
    ]


def _write_registry(path: Path) -> None:
    write_registry_workbook(
        path,
        [
            DsRegistryRow(
                status=STATUS_ACTIVE,
                source_id="13",
                relations=(
                    DsRegistryRelation(
                        group_id="ДС13",
                        rfp_key="13",
                        ul_folder="согл УЛ ДС13",
                        mode=MODE_WHOLE,
                    ),
                ),
            )
        ],
    )


def _write_rfp(path: Path, *, qty: object = 2) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    header = [""] * 20
    header[0] = "№ п/п"
    header[1] = "Титул"
    header[2] = "Спецификация"
    header[3] = "Линия, Tag-номер"
    header[5] = "Код РД"
    header[6] = "Наименование МТР"
    header[7] = "Технические характеристики"
    header[16] = "Кол-во"
    header[17] = "Ед. изм"
    data = [""] * 20
    data[0] = 1
    data[1] = "8529-SOS"
    data[2] = "AGCC.287-8529-SOS.MTO-0001"
    data[5] = CODE
    data[6] = "Кабель"
    data[7] = "NYM"
    data[16] = qty
    data[17] = "шт"
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Перечень материалов"
    ws.append(["титульный"])
    ws.append(header)
    ws.append(data)
    wb.save(path)
    wb.close()


def _ensure_kw() -> dict[str, object]:
    return {
        "converter": IdentityDsUnitsConverter(),
        "google_index": _google_index(),
    }


class ResolveInputModeSmokeTest(unittest.TestCase):
    def test_defaults_and_unknown_to_legacy(self) -> None:
        self.assertEqual(resolve_input_mode(None), INPUT_MODE_LEGACY_NET)
        self.assertEqual(resolve_input_mode({}), INPUT_MODE_LEGACY_NET)
        self.assertEqual(
            resolve_input_mode({"rfp_parts": {"input_mode": "hybrid"}}),
            INPUT_MODE_HYBRID,
        )
        self.assertEqual(
            resolve_input_mode({"rfp_parts": {"input_mode": "ds_only"}}),
            INPUT_MODE_DS_ONLY,
        )
        self.assertEqual(
            resolve_input_mode({"rfp_parts": {"input_mode": "nope"}}),
            INPUT_MODE_LEGACY_NET,
        )


class DsBaselinePreflightSmokeTest(unittest.TestCase):
    def test_rebuild_skip_fingerprint_and_blocking(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            root = Path(raw)
            ds_root = root / "ds"
            reports = root / "reports"
            registry = root / "registry.xlsx"
            _write_registry(registry)
            _write_xlsx(ds_root / "ДС13.xlsx", [DS_HEADER, _ds_row()])

            first, first_fresh = ensure_ds_baseline_current(
                source_root=ds_root,
                registry_path=registry,
                reports_base=reports,
                **_ensure_kw(),
            )
            self.assertTrue(first_fresh.needs_rebuild)
            self.assertEqual(first.name, BASELINE_XLSX_NAME)
            self.assertTrue(first.is_file())
            self.assertEqual(first.parent.name, DS_BASELINE_DIR_NAME)
            mtime = first.stat().st_mtime_ns

            second, second_fresh = ensure_ds_baseline_current(
                source_root=ds_root,
                registry_path=registry,
                reports_base=reports,
                **_ensure_kw(),
            )
            self.assertFalse(second_fresh.needs_rebuild)
            self.assertEqual(second, first)
            self.assertEqual(second.stat().st_mtime_ns, mtime)

            _write_xlsx(ds_root / "ДС13.xlsx", [DS_HEADER, _ds_row(qty=5)])
            third, third_fresh = ensure_ds_baseline_current(
                source_root=ds_root,
                registry_path=registry,
                reports_base=reports,
                **_ensure_kw(),
            )
            self.assertTrue(third_fresh.needs_rebuild)
            self.assertTrue(third.is_file())

            _write_xlsx(ds_root / "unmapped.xlsx", [DS_HEADER, _ds_row(npp=9)])
            with self.assertRaises(DsBaselineBlockedError) as ctx:
                ensure_ds_baseline_current(
                    source_root=ds_root,
                    registry_path=registry,
                    reports_base=reports,
                    **_ensure_kw(),
                )
            self.assertIn("BLOCKED", str(ctx.exception))
            self.assertTrue(
                any(
                    path.name.startswith(STRUCTURE_REPORT_PREFIX)
                    for path in ds_baseline_output_dir(reports).rglob("*.xlsx")
                )
            )

    def test_empty_source_dir_raises_russian(self) -> None:
        with mock.patch(
            "RFQ.rfp_parts.ds_hybrid_preflight._gui_ds_paths",
            return_value={},
        ):
            with self.assertRaises(FileNotFoundError) as ctx:
                ensure_ds_baseline_current(
                    config={"rfp_parts": {"ds_source_dir": ""}}
                )
        message = str(ctx.exception)
        self.assertIn("ds_source_dir", message)
        self.assertIn("last_ds_trusted_folder", message)

    def test_source_dir_falls_back_to_gui_paths(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            root = Path(raw)
            ds_root = root / "ds"
            reports = root / "reports"
            registry = root / "registry.xlsx"
            _write_registry(registry)
            _write_xlsx(ds_root / "ДС13.xlsx", [DS_HEADER, _ds_row()])
            with mock.patch(
                "RFQ.rfp_parts.ds_hybrid_preflight._gui_ds_paths",
                return_value={"last_ds_trusted_folder": str(ds_root)},
            ):
                path, freshness = ensure_ds_baseline_current(
                    config={"rfp_parts": {"ds_source_dir": ""}},
                    registry_path=registry,
                    reports_base=reports,
                    **_ensure_kw(),
                )
            self.assertTrue(path.is_file())
            self.assertTrue(freshness.needs_rebuild)


class DsHybridPreflightSmokeTest(unittest.TestCase):
    def test_hybrid_rebuild_no_parts_net_fallback(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            root = Path(raw)
            ds_root = root / "ds"
            rfp_root = root / "rfp"
            reports = root / "reports"
            registry = root / "registry.xlsx"
            _write_registry(registry)
            _write_xlsx(ds_root / "ДС13.xlsx", [DS_HEADER, _ds_row(qty=2)])
            _write_rfp(rfp_root / "ДС13. AGCC.xlsx", qty=2)

            path, freshness = ensure_ds_hybrid_current(
                source_root=ds_root,
                registry_path=registry,
                rfp_root=rfp_root,
                reports_base=reports,
                rfp_converter=IdentityRfpUnitsConverter(),
                **_ensure_kw(),
            )
            self.assertTrue(freshness.needs_rebuild)
            self.assertEqual(path.name, HYBRID_XLSX_NAME)
            self.assertTrue(path.is_file())
            self.assertEqual(path.parent.name, DS_HYBRID_DIR_NAME)
            self.assertIsNotNone(resolve_ds_hybrid_xlsx(reports_base=reports))
            self.assertIsNotNone(resolve_ds_baseline_xlsx(reports_base=reports))
            self.assertFalse((path.parent / "rfp_parts_net.xlsx").exists())
            self.assertFalse((reports / "rfp_parts_net.xlsx").exists())
            self.assertIn("MATCH", freshness.summary)

            again, again_fresh = ensure_ds_hybrid_current(
                source_root=ds_root,
                registry_path=registry,
                rfp_root=rfp_root,
                reports_base=reports,
                rfp_converter=IdentityRfpUnitsConverter(),
                **_ensure_kw(),
            )
            self.assertFalse(again_fresh.needs_rebuild)
            self.assertEqual(again, path)

    def test_source_registry_newer_mtime_stale_reason(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            root = Path(raw)
            ds_root = root / "ds"
            rfp_root = root / "rfp"
            reports = root / "reports"
            registry = root / "registry.xlsx"
            _write_registry(registry)
            _write_xlsx(ds_root / "ДС13.xlsx", [DS_HEADER, _ds_row(qty=2)])
            _write_rfp(rfp_root / "ДС13. AGCC.xlsx", qty=2)

            path, freshness = ensure_ds_hybrid_current(
                source_root=ds_root,
                registry_path=registry,
                rfp_root=rfp_root,
                reports_base=reports,
                rfp_converter=IdentityRfpUnitsConverter(),
                mix_mode="separate",
                **_ensure_kw(),
            )
            self.assertTrue(freshness.needs_rebuild)
            self.assertIsNotNone(freshness.derived_registry_path)
            derived = freshness.derived_registry_path
            assert derived is not None
            self.assertEqual(derived.name, registry.name)
            self.assertTrue(derived.is_file())
            state_path = ds_hybrid_output_dir(reports) / STATE_JSON_NAME
            state = json.loads(state_path.read_text(encoding="utf-8"))
            self.assertEqual(state["derived_registry_name"], registry.name)
            self.assertIn("registry_mtime_ns", state)
            stored_mtime = int(state["registry_mtime_ns"])

            current = registry.stat()
            os.utime(
                registry,
                ns=(current.st_atime_ns, stored_mtime + 1_000_000_000),
            )

            again, again_fresh = ensure_ds_hybrid_current(
                source_root=ds_root,
                registry_path=registry,
                rfp_root=rfp_root,
                reports_base=reports,
                rfp_converter=IdentityRfpUnitsConverter(),
                mix_mode="separate",
                **_ensure_kw(),
            )
            self.assertTrue(again_fresh.needs_rebuild)
            self.assertTrue(
                again_fresh.reason.startswith("реестр ДС/УЛ новее свода")
            )
            self.assertTrue(again.is_file())
            self.assertIsNotNone(again_fresh.derived_registry_path)


class CopyDsSidecarsSmokeTest(unittest.TestCase):
    def test_copy_missing_skip_and_oserror_prints(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            root = Path(raw)
            stamp = root / "stamp"
            result = root / "result"
            stamp.mkdir()
            sidecar = stamp / f"{QUALITY_REPORT_PREFIX}_stamp.xlsx"
            sidecar.write_bytes(b"xlsx")
            (stamp / BASELINE_XLSX_NAME).write_bytes(b"net")
            (stamp / "rfp_parts_net.xlsx").write_bytes(b"parts")
            copied = copy_ds_sidecars_to_result_dir(stamp, result)
            names = {path.name for path in copied}
            self.assertIn(sidecar.name, names)
            self.assertNotIn(BASELINE_XLSX_NAME, names)
            self.assertNotIn("rfp_parts_net.xlsx", names)

            missing = copy_ds_sidecars_to_result_dir(root / "absent", result)
            self.assertEqual(missing, [])

            buffer = io.StringIO()
            with mock.patch(
                "RFQ.rfp_parts.ds_hybrid_preflight.shutil.copy2",
                side_effect=OSError("locked"),
            ):
                with mock.patch("sys.stdout", buffer):
                    failed = copy_ds_sidecars_to_result_dir(stamp, result / "locked")
            self.assertEqual(failed, [])
            self.assertIn("Не удалось скопировать отчёт ДС", buffer.getvalue())


if __name__ == "__main__":
    unittest.main()
