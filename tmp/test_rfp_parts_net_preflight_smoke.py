"""Offline tests for RFP parts-net freshness vs RFP_Зиновьев."""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.rfp_parts.analyze_rfp_parts import (
    AGGREGATION_KEY,
    BUILD_DEPS_JSON_NAME,
    NET_NO_TAGS_XLSX_NAME,
    NET_XLSX_NAME,
    build_parts_build_deps_snapshot,
    write_parts_build_deps_snapshot,
)
from RFQ.rfp_parts.parts_net_preflight import (
    SOURCES_JSON_NAME,
    assess_parts_net_freshness,
    ensure_rfp_parts_net_current,
    write_parts_sources_snapshot,
)
from RFQ.units_convert import ALGORITHM_VERSION, ConversionDependency, ConversionPlan
from RFQ.units_convert.matrix import MatrixDocument, MatrixPair, MatrixRow
from base.base_classes import RowStd
from base.tables_columns import CODE, UNITS


def _touch_xlsx(path: Path, mtime: float) -> None:
    path.write_bytes(b"dummy")
    os_utime = __import__("os").utime
    os_utime(path, (mtime, mtime))


def _write_net(run_dir: Path, mtime: float, files: list[Path]) -> Path:
    run_dir.mkdir(parents=True, exist_ok=True)
    net = run_dir / NET_XLSX_NAME
    _touch_xlsx(net, mtime)
    write_parts_sources_snapshot(run_dir, files[0].parent if files else run_dir, files)
    return net


def _sample_deps_snapshot(matrix_path: Path, *, coef: str = "1") -> dict:
    return {
        "ALGORITHM_VERSION": ALGORITHM_VERSION,
        "AGGREGATION_KEY": AGGREGATION_KEY,
        "matrix_path": str(matrix_path.resolve()),
        "dependencies": [
            {
                "code": "BCC0000516",
                "source_unit": "шт",
                "target_unit": "шт",
                "coefficient": coef,
            }
        ],
        "google_units_by_code": {"BCC0000516": "шт"},
    }


def _write_deps(run_dir: Path, snapshot: dict) -> Path:
    return write_parts_build_deps_snapshot(run_dir, snapshot)


def _google_mock(*, unit: str | None = "шт") -> mock.Mock:
    index = mock.Mock()
    index.google_unit.side_effect = lambda code: unit if code == "BCC0000516" else None
    return index


def _empty_matrix_doc() -> MatrixDocument:
    return MatrixDocument(rows=())


def _forbidden_extract(*_args, **_kwargs):
    raise AssertionError("extract_parts_records must not run during preflight")


def _forbidden_plan(*_args, **_kwargs):
    raise AssertionError("build_conversion_plan must not run during preflight")


class PartsNetPreflightSmokeTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.parts = self.root / "parts"
        self.reports = self.root / "reports"
        self.parts.mkdir()
        self.reports.mkdir()
        self._forbidden_patches = [
            mock.patch(
                "RFQ.rfp_parts.analyze_rfp_parts.extract_parts_records",
                side_effect=_forbidden_extract,
            ),
            mock.patch(
                "RFQ.units_convert.build_conversion_plan",
                side_effect=_forbidden_plan,
            ),
        ]
        self._patchers = [item.start() for item in self._forbidden_patches]

    def tearDown(self) -> None:
        for patcher in reversed(self._patchers):
            patcher.stop()
        self.temp_dir.cleanup()

    def test_empty_parts_folder_raises(self) -> None:
        with self.assertRaises(FileNotFoundError):
            assess_parts_net_freshness(
                parts_dir=self.parts, reports_base=self.reports
            )

    def test_missing_net_needs_rebuild(self) -> None:
        _touch_xlsx(self.parts / "ДС1.xlsx", 1_000)
        freshness = assess_parts_net_freshness(
            parts_dir=self.parts, reports_base=self.reports
        )
        self.assertTrue(freshness.needs_rebuild)
        self.assertIsNone(freshness.net_path)
        self.assertIn("нет rfp_parts_net", freshness.reason)

    def test_newer_part_file_needs_rebuild(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 2_000)
        run_dir = self.reports / "2026.08.18_12.00"
        net = _write_net(run_dir, 1_000, [part])
        freshness = assess_parts_net_freshness(
            parts_dir=self.parts, reports_base=self.reports
        )
        self.assertTrue(freshness.needs_rebuild)
        self.assertEqual(freshness.net_path, net)
        self.assertEqual([path.name for path in freshness.newer_than_net], ["ДС1.xlsx"])

    def test_matching_snapshot_is_current(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)
        run_dir = self.reports / "2026.08.18_12.00"
        matrix_path = self.root / "matrix.xlsx"
        matrix_path.write_bytes(b"matrix")
        net = _write_net(run_dir, 2_000, [part])
        snapshot = _sample_deps_snapshot(matrix_path)
        _write_deps(run_dir, snapshot)
        with mock.patch(
            "RFQ.rfp_parts.parts_net_preflight.load_matrix",
            return_value=_empty_matrix_doc(),
        ):
            freshness = assess_parts_net_freshness(
                parts_dir=self.parts,
                reports_base=self.reports,
                units_matrix_path=matrix_path,
                google_index=_google_mock(unit="шт"),
            )
        self.assertFalse(freshness.needs_rebuild)
        self.assertEqual(freshness.net_path, net)
        self.assertIn("актуален", freshness.reason)

    def test_new_filename_needs_rebuild_even_if_older_mtime(self) -> None:
        old = self.parts / "ДС1.xlsx"
        new = self.parts / "ДС2.xlsx"
        _touch_xlsx(old, 1_000)
        _touch_xlsx(new, 500)
        run_dir = self.reports / "2026.08.18_12.00"
        _write_net(run_dir, 2_000, [old])
        freshness = assess_parts_net_freshness(
            parts_dir=self.parts, reports_base=self.reports
        )
        self.assertTrue(freshness.needs_rebuild)
        self.assertEqual(freshness.added_names, ["ДС2.xlsx"])
        self.assertIn("новые файлы", freshness.reason)

    def test_ensure_rebuilds_when_stale(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)

        def fake_run(*, out_dir, **_kwargs):
            net = Path(out_dir) / NET_XLSX_NAME
            net.parent.mkdir(parents=True, exist_ok=True)
            net.write_bytes(b"net")
            return net

        with mock.patch(
            "RFQ.rfp_parts.analyze_rfp_parts.run_rfp_parts_analyze",
            side_effect=fake_run,
        ) as mocked:
            net_path, freshness = ensure_rfp_parts_net_current(
                parts_dir=self.parts,
                reports_base=self.reports,
            )

        self.assertTrue(freshness.needs_rebuild)
        self.assertTrue(net_path.is_file())
        mocked.assert_called_once()
        self.assertTrue(mocked.call_args.kwargs["no_checklist"])

    def test_sources_snapshot_round_trip(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)
        out = self.root / "stamp"
        out.mkdir()
        path = write_parts_sources_snapshot(out, self.parts, [part])
        payload = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(path.name, SOURCES_JSON_NAME)
        self.assertEqual(payload["files"][0]["name"], "ДС1.xlsx")

    def test_missing_build_deps_needs_rebuild(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)
        run_dir = self.reports / "2026.08.18_12.00"
        matrix_path = self.root / "matrix.xlsx"
        matrix_path.write_bytes(b"matrix")
        _write_net(run_dir, 2_000, [part])
        freshness = assess_parts_net_freshness(
            parts_dir=self.parts,
            reports_base=self.reports,
            units_matrix_path=matrix_path,
            google_index=_google_mock(),
        )
        self.assertTrue(freshness.needs_rebuild)
        self.assertIn("rfp_parts_build_deps", freshness.reason)

    def test_algorithm_mismatch_needs_rebuild(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)
        run_dir = self.reports / "2026.08.18_12.00"
        matrix_path = self.root / "matrix.xlsx"
        matrix_path.write_bytes(b"matrix")
        _write_net(run_dir, 2_000, [part])
        stored = _sample_deps_snapshot(matrix_path)
        stored["ALGORITHM_VERSION"] = "legacy"
        _write_deps(run_dir, stored)
        freshness = assess_parts_net_freshness(
            parts_dir=self.parts,
            reports_base=self.reports,
            units_matrix_path=matrix_path,
            google_index=_google_mock(),
        )
        self.assertTrue(freshness.needs_rebuild)
        self.assertIn("версия алгоритма", freshness.reason)

    def test_aggregation_key_mismatch_needs_rebuild(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)
        run_dir = self.reports / "2026.08.18_12.00"
        matrix_path = self.root / "matrix.xlsx"
        matrix_path.write_bytes(b"matrix")
        _write_net(run_dir, 2_000, [part])
        stored = _sample_deps_snapshot(matrix_path)
        stored.pop("AGGREGATION_KEY", None)
        _write_deps(run_dir, stored)
        freshness = assess_parts_net_freshness(
            parts_dir=self.parts,
            reports_base=self.reports,
            units_matrix_path=matrix_path,
            google_index=_google_mock(),
        )
        self.assertTrue(freshness.needs_rebuild)
        self.assertIn("ключ суммирования", freshness.reason)

    def test_relevant_coef_change_needs_rebuild(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)
        run_dir = self.reports / "2026.08.18_12.00"
        matrix_path = self.root / "matrix.xlsx"
        matrix_path.write_bytes(b"matrix")
        _write_net(run_dir, 2_000, [part])
        _write_deps(
            run_dir,
            {
                **_sample_deps_snapshot(matrix_path, coef="1"),
                "dependencies": [
                    {
                        "code": "BCC0000516",
                        "source_unit": "м",
                        "target_unit": "шт",
                        "coefficient": "1",
                    }
                ],
            },
        )
        matrix_doc = MatrixDocument(
            rows=(
                MatrixRow(
                    code_display="BCC0000516",
                    code_normalized="BCC0000516",
                    google_name="",
                    code_status="",
                    google_display="шт",
                    google_normalized="шт",
                    pairs=[
                        MatrixPair(
                            source_display="м",
                            source_normalized="м",
                            coefficient_raw="2",
                            coefficient=Decimal("2"),
                            is_placeholder=False,
                        )
                    ],
                ),
            )
        )
        with mock.patch(
            "RFQ.rfp_parts.parts_net_preflight.load_matrix",
            return_value=matrix_doc,
        ):
            freshness = assess_parts_net_freshness(
                parts_dir=self.parts,
                reports_base=self.reports,
                units_matrix_path=matrix_path,
                google_index=_google_mock(unit="шт"),
            )
        self.assertTrue(freshness.needs_rebuild)
        self.assertIn("зависимости", freshness.reason)

    def test_relevant_google_unit_change_needs_rebuild(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)
        run_dir = self.reports / "2026.08.18_12.00"
        matrix_path = self.root / "matrix.xlsx"
        matrix_path.write_bytes(b"matrix")
        _write_net(run_dir, 2_000, [part])
        _write_deps(
            run_dir,
            {
                **_sample_deps_snapshot(matrix_path),
                "google_units_by_code": {"BCC0000516": None},
            },
        )
        with mock.patch(
            "RFQ.rfp_parts.parts_net_preflight.load_matrix",
            return_value=_empty_matrix_doc(),
        ):
            freshness = assess_parts_net_freshness(
                parts_dir=self.parts,
                reports_base=self.reports,
                units_matrix_path=matrix_path,
                google_index=_google_mock(unit="шт"),
            )
        self.assertTrue(freshness.needs_rebuild)
        self.assertIn("Google UNITS", freshness.reason)

    def test_unrelated_google_change_no_rebuild(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)
        run_dir = self.reports / "2026.08.18_12.00"
        matrix_path = self.root / "matrix.xlsx"
        matrix_path.write_bytes(b"matrix")
        net = _write_net(run_dir, 2_000, [part])
        _write_deps(run_dir, _sample_deps_snapshot(matrix_path))
        google = mock.Mock()
        google.google_unit.side_effect = lambda code: {
            "BCC0000516": "шт",
            "BCC0000999": "кг",
        }.get(code)
        with mock.patch(
            "RFQ.rfp_parts.parts_net_preflight.load_matrix",
            return_value=_empty_matrix_doc(),
        ):
            freshness = assess_parts_net_freshness(
                parts_dir=self.parts,
                reports_base=self.reports,
                units_matrix_path=matrix_path,
                google_index=google,
            )
        self.assertFalse(freshness.needs_rebuild)
        self.assertEqual(freshness.net_path, net)

    def test_invalid_matrix_placeholder_needs_rebuild(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)
        run_dir = self.reports / "2026.08.18_12.00"
        matrix_path = self.root / "matrix.xlsx"
        matrix_path.write_bytes(b"matrix")
        _write_net(run_dir, 2_000, [part])
        _write_deps(
            run_dir,
            {
                **_sample_deps_snapshot(matrix_path, coef="2"),
                "dependencies": [
                    {
                        "code": "BCC0000516",
                        "source_unit": "м",
                        "target_unit": "шт",
                        "coefficient": "2",
                    }
                ],
            },
        )
        matrix_doc = MatrixDocument(
            rows=(
                MatrixRow(
                    code_display="BCC0000516",
                    code_normalized="BCC0000516",
                    google_name="",
                    code_status="",
                    google_display="шт",
                    google_normalized="шт",
                    pairs=[
                        MatrixPair(
                            source_display="м",
                            source_normalized="м",
                            coefficient_raw="?",
                            coefficient=None,
                            is_placeholder=True,
                        )
                    ],
                ),
            )
        )
        with mock.patch(
            "RFQ.rfp_parts.parts_net_preflight.load_matrix",
            return_value=matrix_doc,
        ):
            freshness = assess_parts_net_freshness(
                parts_dir=self.parts,
                reports_base=self.reports,
                units_matrix_path=matrix_path,
                google_index=_google_mock(unit="шт"),
            )
        self.assertTrue(freshness.needs_rebuild)
        self.assertIn("зависимости", freshness.reason)

    def test_build_deps_snapshot_roundtrip(self) -> None:
        matrix_path = self.root / "matrix.xlsx"
        matrix_path.write_bytes(b"matrix")
        plan = ConversionPlan(
            actions=(),
            warnings=(),
            dependencies=(
                ConversionDependency(
                    code="BCC0000516",
                    source_unit="шт",
                    target_unit="шт",
                    coefficient=Decimal("1"),
                ),
            ),
        )
        row = RowStd()
        row.el[CODE].value = "BCC0000516"
        row.el[UNITS].value = "шт"
        from RFQ.units_convert import build_google_units_index

        google = build_google_units_index([row])
        snapshot = build_parts_build_deps_snapshot(
            plan,
            google_index=google,
            matrix_path=matrix_path,
        )
        out = self.root / "stamp"
        out.mkdir()
        path = write_parts_build_deps_snapshot(out, snapshot)
        payload = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(path.name, BUILD_DEPS_JSON_NAME)
        self.assertEqual(payload["AGGREGATION_KEY"], AGGREGATION_KEY)
        self.assertEqual(payload["dependencies"][0]["coefficient"], "1")
        self.assertEqual(payload["google_units_by_code"]["BCC0000516"], "шт")

    def test_ensure_forwards_units_matrix_path(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)
        matrix_path = self.root / "matrix.xlsx"
        matrix_path.write_bytes(b"matrix")

        def fake_run(*, out_dir, units_matrix_path=None, **_kwargs):
            self.assertEqual(units_matrix_path, matrix_path)
            net = Path(out_dir) / NET_XLSX_NAME
            net.parent.mkdir(parents=True, exist_ok=True)
            net.write_bytes(b"net")
            return net

        with mock.patch(
            "RFQ.rfp_parts.analyze_rfp_parts.run_rfp_parts_analyze",
            side_effect=fake_run,
        ) as mocked:
            net_path, freshness = ensure_rfp_parts_net_current(
                parts_dir=self.parts,
                reports_base=self.reports,
                units_matrix_path=matrix_path,
                google_index=_google_mock(),
            )
        self.assertTrue(freshness.needs_rebuild)
        self.assertTrue(net_path.is_file())
        mocked.assert_called_once()

    def test_require_no_tags_rebuilds_when_sibling_missing(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)
        run_dir = self.reports / "2026.08.18_12.00"
        matrix_path = self.root / "matrix.xlsx"
        matrix_path.write_bytes(b"matrix")
        net = _write_net(run_dir, 2_000, [part])
        _write_deps(run_dir, _sample_deps_snapshot(matrix_path))
        with mock.patch(
            "RFQ.rfp_parts.parts_net_preflight.load_matrix",
            return_value=_empty_matrix_doc(),
        ):
            tagged = assess_parts_net_freshness(
                parts_dir=self.parts,
                reports_base=self.reports,
                units_matrix_path=matrix_path,
                google_index=_google_mock(unit="шт"),
            )
            missing = assess_parts_net_freshness(
                parts_dir=self.parts,
                reports_base=self.reports,
                units_matrix_path=matrix_path,
                google_index=_google_mock(unit="шт"),
                require_no_tags=True,
            )
        self.assertFalse(tagged.needs_rebuild)
        self.assertTrue(missing.needs_rebuild)
        self.assertEqual(missing.net_path, net)
        self.assertIn("rfp_parts_net_no_tags", missing.reason)

    def test_require_no_tags_current_when_sibling_present(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)
        run_dir = self.reports / "2026.08.18_12.00"
        matrix_path = self.root / "matrix.xlsx"
        matrix_path.write_bytes(b"matrix")
        net = _write_net(run_dir, 2_000, [part])
        _touch_xlsx(run_dir / NET_NO_TAGS_XLSX_NAME, 2_000)
        _write_deps(run_dir, _sample_deps_snapshot(matrix_path))
        with mock.patch(
            "RFQ.rfp_parts.parts_net_preflight.load_matrix",
            return_value=_empty_matrix_doc(),
        ):
            freshness = assess_parts_net_freshness(
                parts_dir=self.parts,
                reports_base=self.reports,
                units_matrix_path=matrix_path,
                google_index=_google_mock(unit="шт"),
                require_no_tags=True,
            )
            selected, _ = ensure_rfp_parts_net_current(
                parts_dir=self.parts,
                reports_base=self.reports,
                units_matrix_path=matrix_path,
                google_index=_google_mock(unit="шт"),
                load_tags=False,
            )
        self.assertFalse(freshness.needs_rebuild)
        self.assertEqual(freshness.net_path, net)
        self.assertEqual(selected, run_dir / NET_NO_TAGS_XLSX_NAME)

    def test_ensure_load_tags_false_rebuild_writes_no_tags(self) -> None:
        part = self.parts / "ДС1.xlsx"
        _touch_xlsx(part, 1_000)

        def fake_run(*, out_dir, **_kwargs):
            stamp = Path(out_dir)
            stamp.mkdir(parents=True, exist_ok=True)
            tagged = stamp / NET_XLSX_NAME
            tagged.write_bytes(b"net")
            (stamp / NET_NO_TAGS_XLSX_NAME).write_bytes(b"no_tags")
            return tagged

        with mock.patch(
            "RFQ.rfp_parts.analyze_rfp_parts.run_rfp_parts_analyze",
            side_effect=fake_run,
        ) as mocked:
            net_path, freshness = ensure_rfp_parts_net_current(
                parts_dir=self.parts,
                reports_base=self.reports,
                load_tags=False,
            )
        self.assertTrue(freshness.needs_rebuild)
        self.assertEqual(net_path.name, NET_NO_TAGS_XLSX_NAME)
        self.assertTrue(net_path.is_file())
        mocked.assert_called_once()


if __name__ == "__main__":
    unittest.main(verbosity=2)
