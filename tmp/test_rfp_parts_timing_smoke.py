"""Smoke tests for RFP parts live phase timing (stdout + artifact)."""

from __future__ import annotations

import contextlib
import io
import sys
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.rfp_parts.analyze_rfp_parts import (
    PARTS_TIMING_ARTIFACT_NAME,
    PARTS_TIMING_PREFIX,
    PARTS_TIMING_SLOW_THRESHOLD_SEC,
    _PartsTimingSession,
    parts_timing_format_done,
    parts_timing_format_error,
    parts_timing_format_progress,
    parts_timing_format_slow_file,
    parts_timing_format_start,
    parts_timing_format_total,
)


class RfpPartsTimingSmokeTest(unittest.TestCase):
    def test_format_start_done_and_total(self) -> None:
        self.assertEqual(
            parts_timing_format_start("чтение матрицы и построение плана ЕИ"),
            "[parts timing] START: чтение матрицы и построение плана ЕИ",
        )
        done = parts_timing_format_done("список файлов частей", 0.01, 1.23)
        self.assertTrue(done.startswith("[parts timing] DONE: список файлов частей — "))
        self.assertIn("0.01 с; всего 1.23 с", done)
        total = parts_timing_format_total(34.56)
        self.assertEqual(total, "[parts timing] TOTAL: 34.56 с")

    def test_phase_context_manager_done_and_total_elapsed(self) -> None:
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            with mock.patch(
                "RFQ.rfp_parts.analyze_rfp_parts.time.perf_counter",
                side_effect=[100.0, 100.0, 100.05, 100.12, 100.12],
            ):
                session = _PartsTimingSession()
                with session.phase("тестовая фаза"):
                    pass
                session.emit_total()
        lines = output.getvalue().splitlines()
        self.assertEqual(len(lines), 3)
        self.assertEqual(lines[0], "[parts timing] START: тестовая фаза")
        self.assertEqual(
            lines[1],
            "[parts timing] DONE: тестовая фаза — 0.05 с; всего 0.12 с",
        )
        self.assertEqual(lines[2], "[parts timing] TOTAL: 0.12 с")

    def test_phase_error_preserves_exception(self) -> None:
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            with mock.patch(
                "RFQ.rfp_parts.analyze_rfp_parts.time.perf_counter",
                side_effect=[200.0, 200.0, 200.03, 200.03],
            ):
                session = _PartsTimingSession()
                with self.assertRaises(ValueError):
                    with session.phase("фаза с ошибкой"):
                        raise ValueError("boom")
        err_line = output.getvalue().splitlines()[-1]
        self.assertTrue(err_line.startswith("[parts timing] ERROR: фаза с ошибкой — "))
        self.assertIn("ValueError", err_line)

    def test_progress_and_slow_formatters(self) -> None:
        self.assertEqual(
            parts_timing_format_progress(10, 47, 12.34),
            "[parts timing] PROGRESS: файлы 10/47; извлечение 12.34 с",
        )
        self.assertEqual(
            parts_timing_format_slow_file(3, 47, "part.xlsx", 2.01),
            "[parts timing] SLOW: файл 3/47 part.xlsx — 2.01 с",
        )
        self.assertEqual(PARTS_TIMING_SLOW_THRESHOLD_SEC, 2.0)
        self.assertEqual(PARTS_TIMING_PREFIX, "[parts timing]")

    def test_error_formatter_stable(self) -> None:
        line = parts_timing_format_error("план ЕИ", 5.0, 10.0, "UnitsConversionError")
        self.assertEqual(
            line,
            "[parts timing] ERROR: план ЕИ — 5.00 с; всего 10.00 с; UnitsConversionError",
        )

    def test_write_artifact_best_effort_on_success(self) -> None:
        session = _PartsTimingSession()
        with contextlib.redirect_stdout(io.StringIO()):
            with session.phase("короткая фаза"):
                pass
        out_dir = ROOT / "tmp" / "_rfp_parts_timing_smoke_out"
        out_dir.mkdir(parents=True, exist_ok=True)
        artifact = out_dir / PARTS_TIMING_ARTIFACT_NAME
        if artifact.exists():
            artifact.unlink()
        session.write_artifact(out_dir)
        self.assertTrue(artifact.is_file())
        text = artifact.read_text(encoding="utf-8")
        self.assertIn("[parts timing] START: короткая фаза", text)
        self.assertIn("[parts timing] DONE: короткая фаза", text)

    def test_write_artifact_warns_without_raising(self) -> None:
        session = _PartsTimingSession()
        session._lines.append("[parts timing] START: x")
        output = io.StringIO()
        out_dir = ROOT / "tmp" / "_rfp_parts_timing_warn"
        out_dir.mkdir(parents=True, exist_ok=True)
        with contextlib.redirect_stdout(output):
            with mock.patch.object(
                Path,
                "write_text",
                side_effect=OSError("disk full"),
            ):
                session.write_artifact(out_dir)
        warn = output.getvalue()
        self.assertIn("WARN: cannot write rfp_parts_timing.txt", warn)
        self.assertIn("disk full", warn)


if __name__ == "__main__":
    unittest.main(verbosity=2)
