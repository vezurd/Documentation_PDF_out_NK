"""Local checks for the RD-tree MTO dump walker (tempfile tree only)."""

from __future__ import annotations

import subprocess
import sys
import tempfile
from dataclasses import replace
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.config import CatalogConfig
from rd_catalog.db import CatalogDatabase
from rd_catalog.kits import kit_identity_key
from rd_catalog.rd_dump_index import (
    KIND_OD,
    parse_rd_dump_file,
    rd_dump_is_canonical,
    rd_dump_kind,
    rd_dump_layout_token,
)
from rd_catalog.rd_dump_scan import RdDumpScanOutcome, scan_rd_dump

_VALID_NAME = "AGCC.287-1600-POS.MTO-0001_02-AN01_RU.xlsx"
_LOOSE_NAME = "AGCC.287-1600-POS.MTO-0001_02_RU.xlsx"
_OD_NAME = "AGCC.287-1600-POS.OD-0001_02-AN01_RU.docx"
_OD_LOOSE_NAME = "AGCC.287-1600-POS.OD-0001_02_RU.doc"
_LOCK_NAME = "~$AGCC.287-1600-POS.MTO-0001_02_RU.xlsx"
_OD_LOCK_NAME = "~$AGCC.287-1600-POS.OD-0001_02_RU.docx"
_FORBIDDEN_NAMES = (
    "scan_catalog",
    "execute_scan",
    "store_scan",
    "build_rd_overlays",
)


def _config(root: Path, *, rd_root: Path, db_path: Path) -> CatalogConfig:
    return CatalogConfig(
        rd_root=rd_root,
        sq_root=root / "sq",
        robot_root=root / "robot",
        runtime_dir=root / "runtime",
        db_path=db_path,
        robot_flat_structure=True,
        skip_dirs=("old", "архив"),
        an_root=root / "an",
    )


def _write(path: Path, data: bytes = b"") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)


def _assert_no_scan_catalog_imports() -> None:
    repo = Path(__file__).resolve().parents[1]
    for relative in (
        "rd_catalog/rd_dump_scan.py",
        "rd_catalog/rd_dump_scan_cli.py",
        "rd_catalog/rd_dump_scan_thread.py",
        "rd_catalog/rd_dump_index.py",
        "rd_catalog/rd_dump_tab.py",
    ):
        text = (repo / relative).read_text(encoding="utf-8")
        for name in _FORBIDDEN_NAMES:
            assert name not in text, f"{relative} must not mention {name}"


def _assert_cli_help() -> None:
    repo = Path(__file__).resolve().parents[1]
    completed = subprocess.run(
        [sys.executable, "-m", "rd_catalog.rd_dump_scan_cli", "--help"],
        cwd=repo,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert completed.returncode == 0, completed.stderr
    assert "job" in completed.stdout.casefold()


def main() -> None:
    """Walk a local RD tree, keep non-canonical xlsx/doc, and skip skip-dirs."""

    _assert_no_scan_catalog_imports()
    _assert_cli_help()

    parsed_od = parse_rd_dump_file(
        rf"dummy\{_OD_NAME}",
        size=1,
        mtime_ns=1,
    )
    assert parsed_od is not None
    assert rd_dump_kind(parsed_od) == KIND_OD
    assert parsed_od.revision_text == "02-AN01"
    assert parse_rd_dump_file(
        rf"dummy\{_OD_LOCK_NAME}",
        size=1,
        mtime_ns=1,
    ) is None
    assert parse_rd_dump_file(
        r"dummy\AGCC.287-1600-POS.OD-0001_02_RU.pdf",
        size=1,
        mtime_ns=1,
    ) is None

    with tempfile.TemporaryDirectory(prefix="rd_catalog_rd_dump_") as temp:
        root = Path(temp)
        rd_root = root / "rd"
        junk_root = root / "junk"
        db_path = root / "catalog.sqlite"

        canonical = (
            rd_root
            / "1600"
            / "POS"
            / "Для передачи"
            / "01_рев.02-AN01"
            / "PDF"
            / _VALID_NAME
        )
        loose = rd_root / "1600" / "POS" / _LOOSE_NAME
        canonical_od = (
            rd_root
            / "1600"
            / "POS"
            / "Для передачи"
            / "01_рев.02-AN01"
            / "DWG"
            / _OD_NAME
        )
        loose_od = rd_root / "1600" / "POS" / _OD_LOOSE_NAME
        _write(canonical)
        _write(loose)
        _write(canonical_od)
        _write(loose_od)
        _write(rd_root / "notes.xlsx")
        _write(rd_root / "notes.docx")
        _write(rd_root / _LOCK_NAME)
        _write(rd_root / _OD_LOCK_NAME)
        _write(rd_root / "ignore.pdf", b"%PDF")
        _write(rd_root / "old" / _LOOSE_NAME)
        _write(rd_root / "old" / _OD_LOOSE_NAME)
        _write(rd_root / "архив" / _VALID_NAME)
        _write(junk_root / "notes.xlsx")

        assert rd_dump_is_canonical(str(canonical), rd_root)
        assert not rd_dump_is_canonical(str(loose), rd_root)
        assert rd_dump_layout_token(str(loose), rd_root) == "loose_in_mark"

        config = _config(root, rd_root=rd_root, db_path=db_path)
        database = CatalogDatabase(db_path)
        database.initialize()

        first = scan_rd_dump(config, persist=True)
        assert isinstance(first, RdDumpScanOutcome)
        assert first.failure is None, first.failure
        assert first.cancelled is False
        names = {Path(item.path).name for item in first.files}
        assert names == {_VALID_NAME, _LOOSE_NAME, _OD_NAME, _OD_LOOSE_NAME}, names
        assert first.accepted == 4, first
        # mto valid+loose+notes.xlsx, od canonical+loose+notes.docx; locks/pdf/skip-dirs excluded
        assert first.files_seen == 6, first
        assert first.skipped == 2
        by_kit = database.list_rd_dump_files_by_kit()
        pos_files = by_kit[kit_identity_key("1600", "POS")]
        assert len(pos_files) == 4
        hidden_names = {Path(item.path).name for item in database.list_rd_dump_mto_files()}
        assert _LOCK_NAME not in hidden_names
        assert _OD_LOCK_NAME not in hidden_names
        stored_paths = {item.path for item in database.list_rd_dump_mto_files()}
        assert str(canonical) in stored_paths or any(
            Path(path).name == _VALID_NAME for path in stored_paths
        )
        assert any(Path(path).name == _LOOSE_NAME for path in stored_paths)
        assert any(Path(path).name == _OD_NAME for path in stored_paths)
        assert any(Path(path).name == _OD_LOOSE_NAME for path in stored_paths)
        assert not any("old" in path.casefold() for path in stored_paths)
        an_kept = database.list_an_mto_files()
        assert an_kept == ()

        cancelled = scan_rd_dump(
            config,
            persist=True,
            is_cancelled=lambda: True,
        )
        assert cancelled.cancelled is True
        assert cancelled.failure is None
        assert len(database.list_rd_dump_mto_files()) == 4

        empty_root = scan_rd_dump(replace(config, rd_root=Path("")), persist=True)
        assert empty_root.failure
        assert empty_root.accepted == 0
        kept = database.list_rd_dump_mto_files()
        assert len(kept) == 4

        missing_root = scan_rd_dump(
            replace(config, rd_root=root / "missing_rd"),
            persist=True,
        )
        assert missing_root.failure
        assert len(database.list_rd_dump_mto_files()) == 4

        junk = scan_rd_dump(replace(config, rd_root=junk_root), persist=True)
        assert junk.failure is None, junk.failure
        assert junk.accepted == 0
        assert junk.files_seen == 1
        assert database.list_rd_dump_mto_files() == ()
        assert database.list_rd_dump_files_by_kit() == {}

    print("RD catalog RD dump scan: OK")


if __name__ == "__main__":
    main()
