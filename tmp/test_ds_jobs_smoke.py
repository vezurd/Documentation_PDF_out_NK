"""Offline smoke for DS cockpit jobs and the RFP · Сбор частей panel."""

from __future__ import annotations

import ast
import os
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.ds_compare.ds_compare_config import (
    get_default_gui_paths,
    normalize_gui_paths,
)
from RFQ.rfp_parts.ds_baseline import IdentityDsUnitsConverter
from RFQ.rfp_parts.ds_jobs import (
    CockpitRow,
    DsCockpitSnapshot,
    get_last_ds_cockpit,
    run_ds_baseline_job,
    run_ds_coverage_job,
    run_ds_hybrid_job,
    run_ds_registry_check_job,
)
from RFQ.rfp_parts.ds_registry import (
    DEFAULT_REGISTRY_PATH,
    MODE_WHOLE,
    MIGRATION_REPORT_PREFIX,
    STATUS_ACTIVE,
    DsRegistryRelation,
    DsRegistryRow,
    detect_registry_format,
    write_registry_workbook,
)
from RFQ.rfp_parts.ds_rfp_hybrid import (
    HYBRID_REPORT_PREFIX,
    HYBRID_XLSX_NAME,
    IdentityRfpUnitsConverter,
)
from RFQ.units_convert.models import GoogleUnitsIndex

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

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


def _write_legacy(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws["A1"] = "Актуальный ДС"
    ws["B1"] = "Старый ДС"
    ws["C1"] = "УЛ"
    ws["D1"] = "Примечание"
    ws["A2"] = 13
    ws["B2"] = ""
    ws["C2"] = "согл УЛ ДС13"
    ws["D2"] = ""
    wb.save(path)
    wb.close()


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
        spare_rows=1,
    )


def _write_ds_xlsx(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Спека"
    ws.append(DS_HEADER)
    ws.append(
        [1, "8529", "SOS", "spec", "RFQ", "Кабель", "1C", "BCC0000001", "Vendor", "NYM", "шт", 2]
    )
    wb.save(path)
    wb.close()


def _google() -> GoogleUnitsIndex:
    return GoogleUnitsIndex(
        units_by_code={"BCC0000001": "шт"},
        display_units_by_code={"BCC0000001": "шт"},
        display_code_by_code={"BCC0000001": "BCC0000001"},
    )


class DsJobsSmokeTest(unittest.TestCase):
    def test_gui_paths_have_ds_cockpit_keys(self) -> None:
        defaults = get_default_gui_paths()
        self.assertIn("last_ds_trusted_folder", defaults)
        self.assertIn("last_ds_registry_file", defaults)
        self.assertEqual(defaults["last_ds_registry_file"], str(DEFAULT_REGISTRY_PATH))
        normalized = normalize_gui_paths({"last_ds_file": "x.xlsx"})
        self.assertEqual(normalized["last_ds_registry_file"], str(DEFAULT_REGISTRY_PATH))
        self.assertEqual(normalized["last_ds_file"], "x.xlsx")
        emptied = normalize_gui_paths({"last_ds_trusted_folder": ""})
        self.assertEqual(emptied["last_ds_trusted_folder"], "")

    def test_jobs_module_never_calls_unc_replace(self) -> None:
        import RFQ.rfp_parts.ds_jobs as jobs

        tree = ast.parse(Path(jobs.__file__).read_text(encoding="utf-8"))
        imported: list[str] = []
        for node in ast.walk(tree):
            if isinstance(node, ast.ImportFrom):
                imported.extend(alias.name for alias in node.names)
            elif isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
                imported.append(node.func.id)
        self.assertNotIn("backup_and_replace_registry", imported)
        self.assertNotIn("backup_and_replace_registry", jobs.__all__)

    def test_registry_check_new_format(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            root = Path(raw)
            registry = root / "Реестр_ДС_УЛ.xlsx"
            reports = root / "reports"
            _write_registry(registry)
            result = run_ds_registry_check_job(registry, reports, root / "missing-ul")
            self.assertTrue(result.success)
            self.assertIn("реестр", result.message.lower())
            self.assertTrue(registry.is_file())

    def test_registry_legacy_migrates_copy_not_source(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            root = Path(raw)
            legacy = root / "Реестр_ДС_УЛ.xlsx"
            reports = root / "reports"
            _write_legacy(legacy)
            before = legacy.read_bytes()
            self.assertEqual(detect_registry_format(legacy), "legacy")
            result = run_ds_registry_check_job(legacy, reports, None)
            self.assertTrue(result.success, result.message)
            self.assertEqual(legacy.read_bytes(), before)
            migrated = reports / "Реестр_ДС_УЛ_migrated.xlsx"
            self.assertTrue(migrated.is_file())
            self.assertNotEqual(migrated.resolve(), DEFAULT_REGISTRY_PATH)
            self.assertTrue(
                any(path.name.startswith(MIGRATION_REPORT_PREFIX) for path in reports.iterdir())
            )
            self.assertIn("формат=new", result.message)
            self.assertIn("скопирован из старого", result.message)
            cockpit = get_last_ds_cockpit()
            self.assertIsNotNone(cockpit)
            self.assertIn("Канон UNC не заменён", cockpit.next_step)
            self.assertIn(migrated.name, cockpit.next_step)

    def test_coverage_warns_if_rfp_missing(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            root = Path(raw)
            registry = root / "registry.xlsx"
            ds_root = root / "ds"
            ds_root.mkdir()
            _write_registry(registry)
            _write_ds_xlsx(ds_root / "ДС_13" / "spec.xlsx")
            missing_rfp = root / "no_rfp"
            result = run_ds_coverage_job(
                ds_root, registry, None, None, missing_rfp
            )
            self.assertTrue(result.success, result.message)
            self.assertIn("WARN", result.message)

    def test_baseline_and_hybrid_with_temp_dirs(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            root = Path(raw)
            registry = root / "registry.xlsx"
            ds_root = root / "ds"
            reports = root / "out"
            rfp_root = root / "rfp"
            rfp_root.mkdir()
            _write_registry(registry)
            _write_ds_xlsx(ds_root / "ДС_13" / "spec.xlsx")
            conv = IdentityDsUnitsConverter()
            google = _google()
            baseline = run_ds_baseline_job(
                ds_root,
                registry,
                reports,
                None,
                converter=conv,
                google_index=google,
            )
            self.assertTrue(baseline.success, baseline.message)
            self.assertIsNotNone(baseline.result_path)
            hybrid = run_ds_hybrid_job(
                ds_root,
                registry,
                reports,
                None,
                rfp_root,
                converter=conv,
                rfp_converter=IdentityRfpUnitsConverter(),
                google_index=google,
            )
            self.assertTrue(hybrid.success, hybrid.message)
            self.assertTrue((reports / HYBRID_XLSX_NAME).is_file())
            self.assertTrue(
                any(
                    path.name.startswith(HYBRID_REPORT_PREFIX)
                    for path in reports.glob("*.xlsx")
                )
            )


class RfpPartsDsCockpitPanelSmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        from PySide6.QtWidgets import QApplication

        cls.app = QApplication.instance() or QApplication([])

    def test_instantiate_and_fill_snapshot(self) -> None:
        from ds_compare_center.rfp_parts_panel import RfpPartsPanel

        panel = RfpPartsPanel()
        self.app.processEvents()
        self.assertTrue(hasattr(panel, "monitor"))
        self.assertTrue(hasattr(panel, "splitter"))
        from PySide6.QtWidgets import QPushButton

        texts = [btn.text() for btn in panel.findChildren(QPushButton)]
        self.assertIn("Собрать вход Только ДС", texts)
        self.assertIn("Проверить RFP и наложить", texts)
        self.assertIn("Только покрытие", texts)
        self.assertIn("Проверить реестр", texts)
        self.assertIn("Проверить части RFP и сформировать отчёт", texts)
        snapshot = DsCockpitSnapshot(
            kind="coverage",
            summary="Покрытие: групп=1, файлов ДС=1, RFP=0, без RFP=1, ERROR=0, WARN=1",
            registry_path=Path("registry.xlsx"),
            registry_format="new",
            error_count=0,
            warn_count=1,
            match_count=1,
            empty_code=1,
            registry_rows=[
                CockpitRow(
                    cells=("13", "Активен", "ДС13", "13", "согл УЛ ДС13", "Вся ДС", ""),
                    tone="match",
                )
            ],
            group_rows=[
                CockpitRow(
                    cells=("ДС13", "13", "13", "согл УЛ ДС13", "нет", "MATCH"),
                    tone="match",
                )
            ],
            file_rows=[
                CockpitRow(cells=("ДС", "spec.xlsx", "13", "OK"), tone="ok"),
            ],
            coverage_rows=[
                CockpitRow(
                    cells=("ДС13", "spec.xlsx", "нет RFP", "согл УЛ ДС13", "нет файла RFP"),
                    tone="warn",
                )
            ],
        )
        panel.fill_from_snapshot(snapshot)
        self.app.processEvents()
        self.assertEqual(panel._table_registry.rowCount(), 1)
        self.assertEqual(panel._table_coverage.rowCount(), 1)
        self.assertIn("ERROR=0", panel._cockpit_summary.text())
        self.assertIn("MATCH=1", panel._cockpit_summary.text())
        bg = panel._table_coverage.item(0, 0).background().color().name()
        self.assertEqual(bg.lower(), "#fff8c5")
        match_bg = panel._table_registry.item(0, 0).background().color().name()
        self.assertEqual(match_bg.lower(), "#e6f4ea")
        panel.close()
        panel.deleteLater()
        self.app.processEvents()


if __name__ == "__main__":
    unittest.main(verbosity=2)
