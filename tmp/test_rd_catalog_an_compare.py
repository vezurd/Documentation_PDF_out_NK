"""Local checks for AN pairwise content compare and vs-cell labels."""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.an_compare import (
    AN_VS_CONTENT_MATCH_FILL,
    AN_VS_CONTENT_MISS_FILL,
    AN_VS_REV_DIFF_FILL,
    AN_VS_REV_MATCH_FILL,
    an_vs_cell_fill,
    compare_an_file,
    format_an_vs_cell,
)
from rd_catalog.an_compare_cache import (
    cache_key,
    load_an_content_compare_cache,
    result_from_entry,
    result_to_entry,
    save_an_content_compare_cache,
)
from rd_catalog.customer_pi_auto_mto import compare_mto_pair


def _row(code: str, qty: str, *, tags: tuple[str, ...] = (), units: str = "шт") -> dict[str, object]:
    return {
        "CODE": code,
        "UNITS": units,
        "VALUES": qty,
        "TAGS": list(tags),
        "NAME": "",
        "VENDOR": "",
        "TYPE_MARK": "",
    }


def main() -> None:
    """Run format, pairwise grade, AN load-once, and cache assertions."""

    assert format_an_vs_cell(True, None, has_counterpart=False) == "да"
    assert format_an_vs_cell(True, None, has_counterpart=True) == "да (не сверялось)"
    assert format_an_vs_cell(False, None, has_counterpart=True) == "нет (не сверялось)"
    assert format_an_vs_cell(None, None, has_counterpart=False) == "—"

    exact = {
        "an": [_row("BCC1", "2", tags=("T1", "T2"))],
        "same": [_row("BCC1", "2", tags=("T1", "T2"))],
        "soft": [_row("BCC1", "2", tags=("T1",), units="кг")],
        "diff": [_row("BCC1", "1")],
    }

    def loader(path: str | Path):
        return exact[Path(path).stem]

    matched = compare_mto_pair("an.xlsx", "same.xlsx", loader=loader)
    assert matched.kind == "matched"
    assert matched.paren_label == "четкое"
    assert format_an_vs_cell(True, matched, has_counterpart=True) == "да (четкое)"
    assert an_vs_cell_fill(True, matched) == AN_VS_CONTENT_MATCH_FILL
    assert an_vs_cell_fill(True, None) == AN_VS_REV_MATCH_FILL
    assert an_vs_cell_fill(False, None) == AN_VS_REV_DIFF_FILL
    assert an_vs_cell_fill(None, None) is None

    soft = compare_mto_pair("an.xlsx", "soft.xlsx", loader=loader)
    assert soft.kind == "soft"
    assert soft.paren_label == "ПоКоду и Кол-ву"
    assert format_an_vs_cell(True, soft, has_counterpart=True) == "да (ПоКоду и Кол-ву)"
    assert an_vs_cell_fill(True, soft) == AN_VS_CONTENT_MATCH_FILL

    missing = compare_mto_pair("an.xlsx", "diff.xlsx", loader=loader)
    assert missing.kind == "no_match"
    assert missing.paren_label == "не совпало"
    assert format_an_vs_cell(True, missing, has_counterpart=True) == "да (не совпало)"
    assert an_vs_cell_fill(True, missing) == AN_VS_CONTENT_MISS_FILL
    assert AN_VS_CONTENT_MATCH_FILL != AN_VS_CONTENT_MISS_FILL

    loads: list[str] = []

    def counting_loader(path: str | Path):
        loads.append(Path(path).stem)
        return exact[Path(path).stem]

    vs_auto, vs_rd = compare_an_file(
        "an.xlsx",
        auto_path="diff.xlsx",
        rd_path="same.xlsx",
        loader=counting_loader,
    )
    assert loads == ["an", "diff", "same"]
    assert vs_auto is not None and vs_auto.kind == "no_match"
    assert vs_rd is not None and vs_rd.kind == "matched"

    with tempfile.TemporaryDirectory() as raw:
        runtime = Path(raw)
        key = cache_key("an.xlsx", 1, "pi.xlsx", 2, "rd.xlsx", 3)
        save_an_content_compare_cache(
            runtime,
            {key: {"vs_auto": result_to_entry(missing), "vs_rd": result_to_entry(matched)}},
        )
        loaded = load_an_content_compare_cache(runtime)
        restored_auto = result_from_entry(loaded[key]["vs_auto"])
        restored_rd = result_from_entry(loaded[key]["vs_rd"])
        assert restored_auto is not None and restored_auto.kind == "no_match"
        assert restored_rd is not None and restored_rd.kind == "matched"


if __name__ == "__main__":
    main()
    print("ok")
