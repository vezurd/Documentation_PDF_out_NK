"""Local checks for nested RD catalog operation timing."""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.perf_log import (
    PERF_ENV_VAR,
    PERF_LOG_NAME,
    configure_perf_log,
    is_perf_enabled,
    perf_note,
    perf_span,
    reset_perf_log_for_tests,
    set_perf_sink,
    stamp_log_line,
)


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def main() -> None:
    """Verify configure, nest, sink, disable, and exception END lines."""

    reset_perf_log_for_tests()
    os.environ.pop(PERF_ENV_VAR, None)
    assert not is_perf_enabled()
    with perf_span("noop"):
        pass

    assert stamp_log_line("") == ""
    stamped = stamp_log_line("hello")
    assert stamped.endswith(" hello")
    assert stamp_log_line("2026-09-14 12:00:00.001 already") == (
        "2026-09-14 12:00:00.001 already"
    )

    with tempfile.TemporaryDirectory(prefix="rd_catalog_perf_") as temp:
        runtime = Path(temp) / "runtime"
        journal = configure_perf_log(runtime)
        assert journal == runtime / PERF_LOG_NAME
        assert is_perf_enabled()

        sink_lines: list[str] = []
        set_perf_sink(sink_lines.append)
        perf_note("session", app="test")
        with perf_span("outer", kits=2):
            with perf_span("inner"):
                pass
        try:
            with perf_span("boom"):
                raise RuntimeError("x")
        except RuntimeError:
            pass

        text = _read(journal)
        assert "NOTE  session app=test" in text
        assert "BEGIN outer kits=2" in text
        assert "BEGIN inner" in text
        assert "END   inner elapsed_ms=" in text
        assert "END   outer elapsed_ms=" in text
        assert "kits=2" in text
        assert "BEGIN boom" in text
        assert "END   boom elapsed_ms=" in text
        assert "exc=RuntimeError" in text
        assert any("  BEGIN inner" in line for line in text.splitlines())
        assert sink_lines
        assert sink_lines[0].startswith("20")
        assert "NOTE  session" in sink_lines[0]

        set_perf_sink(None)
        os.environ[PERF_ENV_VAR] = "0"
        assert not is_perf_enabled()
        before = journal.stat().st_size
        with perf_span("disabled"):
            perf_note("also_off")
        assert journal.stat().st_size == before
        os.environ.pop(PERF_ENV_VAR, None)

        configure_perf_log(None)
        assert not is_perf_enabled()
        reset_perf_log_for_tests()
        assert not is_perf_enabled()

    print("rd_catalog perf_log: ok")


if __name__ == "__main__":
    main()
