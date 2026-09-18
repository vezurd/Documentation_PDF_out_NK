"""Local checks for the AN dump walker (tempfile tree only)."""

from __future__ import annotations

import subprocess
import sys
import tempfile
from dataclasses import replace
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.an_scan import AnScanOutcome, scan_an_dump
from rd_catalog.config import CatalogConfig
from rd_catalog.db import CatalogDatabase
from rd_catalog.kits import kit_identity_key

_VALID_NAME = "AGCC.287-1600-POS.MTO-0001_02-AN01_RU.xlsx"
_HIDDEN_VALID_NAME = "AGCC.287-1600-POS.MTO-0001_02_RU.xlsx"
_OD_NAME = "AGCC.287-1600-POS.OD-0001_02-AN01_RU.docx"
_OD_LOCK_NAME = "~$AGCC.287-1600-POS.OD-0001_02_RU.docx"
_LOCK_NAME = "~$AGCC.287-1600-POS.MTO-0001_02_RU.xlsx"
_FORBIDDEN_NAMES = (
    "scan_catalog",
    "execute_scan",
    "store_scan",
    "build_rd_overlays",
)


def _config(root: Path, *, an_root: Path, db_path: Path) -> CatalogConfig:
    return CatalogConfig(
        rd_root=root / "rd",
        sq_root=root / "sq",
        robot_root=root / "robot",
        runtime_dir=root / "runtime",
        db_path=db_path,
        robot_flat_structure=True,
        skip_dirs=("old", "архив"),
        an_root=an_root,
    )


def _write(path: Path, data: bytes = b"") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)


def _assert_no_scan_catalog_imports() -> None:
    repo = Path(__file__).resolve().parents[1]
    for relative in (
        "rd_catalog/an_scan.py",
        "rd_catalog/an_scan_cli.py",
        "rd_catalog/an_scan_thread.py",
    ):
        text = (repo / relative).read_text(encoding="utf-8")
        for name in _FORBIDDEN_NAMES:
            assert name not in text, f"{relative} must not mention {name}"


def _assert_cli_help() -> None:
    repo = Path(__file__).resolve().parents[1]
    completed = subprocess.run(
        [sys.executable, "-m", "rd_catalog.an_scan_cli", "--help"],
        cwd=repo,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert completed.returncode == 0, completed.stderr
    assert "job" in completed.stdout.casefold()


def main() -> None:
    """Walk a local tree, persist snapshots, and keep an empty-root snapshot."""

    _assert_no_scan_catalog_imports()
    _assert_cli_help()

    with tempfile.TemporaryDirectory(prefix="rd_catalog_an_scan_") as temp:
        root = Path(temp)
        an_root = root / "an"
        junk_root = root / "junk"
        db_path = root / "catalog.sqlite"

        _write(an_root / _VALID_NAME)
        _write(an_root / _OD_NAME)
        _write(an_root / "notes.xlsx")
        _write(an_root / "notes.docx")
        _write(an_root / _LOCK_NAME)
        _write(an_root / _OD_LOCK_NAME)
        _write(an_root / "ignore.pdf", b"%PDF")
        _write(an_root / "old" / _HIDDEN_VALID_NAME)
        _write(an_root / "архив" / _HIDDEN_VALID_NAME)
        _write(junk_root / "notes.xlsx")

        config = _config(root, an_root=an_root, db_path=db_path)
        database = CatalogDatabase(db_path)
        database.initialize()

        first = scan_an_dump(config, persist=True)
        assert isinstance(first, AnScanOutcome)
        assert first.failure is None, first.failure
        assert first.cancelled is False
        assert first.accepted == 2, first
        assert first.files_seen == 4, first  # mto + od + notes.xlsx + notes.docx
        assert first.skipped == 2
        assert len(first.files) == 2
        names = {item.name for item in first.files}
        assert _VALID_NAME in names
        assert _OD_NAME in names
        assert first.files[0].title == "1600"
        assert first.files[0].mark == "POS"
        assert first.scanned_at
        by_kit = database.list_an_files_by_kit()
        pos_files = by_kit[kit_identity_key("1600", "POS")]
        assert len(pos_files) == 2
        assert {item.name for item in pos_files} == {_VALID_NAME, _OD_NAME}
        hidden_names = {Path(item.path).name for item in database.list_an_mto_files()}
        assert _HIDDEN_VALID_NAME not in hidden_names

        cancelled = scan_an_dump(
            config,
            persist=True,
            is_cancelled=lambda: True,
        )
        assert cancelled.cancelled is True
        assert cancelled.failure is None
        assert len(database.list_an_mto_files()) == 2

        empty_root = scan_an_dump(replace(config, an_root=Path("")), persist=True)
        assert empty_root.failure
        assert empty_root.accepted == 0
        kept = database.list_an_mto_files()
        assert len(kept) == 2
        assert {item.name for item in kept} == {_VALID_NAME, _OD_NAME}

        missing_root = scan_an_dump(
            replace(config, an_root=root / "missing_an"),
            persist=True,
        )
        assert missing_root.failure
        assert len(database.list_an_mto_files()) == 2

        junk = scan_an_dump(replace(config, an_root=junk_root), persist=True)
        assert junk.failure is None, junk.failure
        assert junk.accepted == 0
        assert junk.files_seen == 1
        assert database.list_an_mto_files() == ()
        assert database.list_an_files_by_kit() == {}

    print("RD catalog AN scan: OK")


if __name__ == "__main__":
    main()
