"""Smoke: grouped compare cache save/load and rfq_only guard (no UNC DS/MTO)."""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from base.base_classes import RowStd, TableComments
from RFQ.ds_compare.ds_grouped_compare import (
    GroupedDsMtoRfqOnlyError,
    analyze_grouped_ds_specification,
)
from RFQ.ds_compare.grouped_compare_cache import (
    build_grouped_compare_extra_key,
    load_grouped_compare_cache,
    save_grouped_compare_cache,
)


def _minimal_row() -> RowStd:
    row = RowStd.get_std_check_row({}, TableComments())
    return row


def test_cache_roundtrip() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        ds_path = Path(tmp) / "fake_ds.xlsx"
        ds_path.write_bytes(b"x")
        extra = build_grouped_compare_extra_key(
            group_keys=["title", "system", "code"],
            flat_mto_structure=False,
            group_mto_new_positions=True,
            mto_path=str(tmp),
            replacement_table_file="",
        )
        spec_dict: dict = {}
        rows = [_minimal_row()]
        save_grouped_compare_cache(
            str(ds_path),
            extra,
            rows,
            replacement_table_file="",
            spec_dict=spec_dict,
        )
        loaded = load_grouped_compare_cache(
            str(ds_path),
            extra,
            replacement_table_file="",
            spec_dict=spec_dict,
        )
        assert loaded is not None and len(loaded) == 1
        print("cache roundtrip OK")


def test_rfq_only_raises_without_cache_file() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        ds_path = Path(tmp) / "ds.xlsx"
        ds_path.write_bytes(b"x")
        rfq_path = Path(tmp) / "rfq.xlsx"
        rfq_path.write_bytes(b"y")
        try:
            analyze_grouped_ds_specification(
                str(ds_path),
                mto_path=str(tmp),
                rfq_path=str(rfq_path),
                rfq_only=True,
            )
        except GroupedDsMtoRfqOnlyError as exc:
            assert "не найден" in str(exc).lower() or "кэш" in str(exc).lower()
            print("rfq_only without cache file: expected error OK")
            return
        raise AssertionError("expected GroupedDsMtoRfqOnlyError")


if __name__ == "__main__":
    test_cache_roundtrip()
    test_rfq_only_raises_without_cache_file()
    print("all smoke checks passed")
