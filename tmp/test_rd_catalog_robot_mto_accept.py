"""Live/stale robot MTO accept helpers and SQLite CRUD."""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.db import SCHEMA_VERSION, CatalogDatabase
from rd_catalog.models import (
    FileKind,
    FileRecord,
    ReviewState,
    SourceKind,
    make_path_key,
)
from rd_catalog.robot_mto_accept import (
    ACCEPT_LIVE,
    ACCEPT_MISSING_ROBOT,
    ACCEPT_STALE,
    ROBOT_MTO_ACCEPT_FOREGROUND,
    build_robot_mto_accept_view,
    robot_mto_accept_paints_blue,
    robot_mto_accept_paints_green,
    robot_mto_accept_state,
)


def _mto(
    *,
    path: str,
    source: SourceKind,
    size: int,
    mtime_ns: int,
    file_id: int,
    revision: str = "01",
    appendix: str | None = "01",
) -> FileRecord:
    return FileRecord(
        id=file_id,
        path=path,
        path_key=make_path_key(path),
        source=source,
        present=True,
        review_state=ReviewState.PENDING,
        first_seen_run_id=1,
        last_seen_run_id=1,
        data={
            "file_kind": FileKind.MTO_XLSX.value,
            "size": size,
            "mtime_ns": mtime_ns,
            "disk_mtime_ns": mtime_ns,
            "revision": revision,
            "appendix": appendix,
            "name": Path(path).name,
        },
    )


def main() -> None:
    """Verify schema 14 CRUD and live/stale/blue rules."""

    assert SCHEMA_VERSION == 14
    assert ROBOT_MTO_ACCEPT_FOREGROUND == "#1A4FBF"
    robot_path = r"C:\robot\AGCC.287-1600-POS.MTO-0001_02-AN01_RU.xlsx"
    rd_path = r"C:\rd\05_рев.02-AN01\AGCC.287-1600-POS.MTO-0001_02-AN01_RU.xlsx"
    robot = _mto(
        path=robot_path,
        source=SourceKind.ROBOT,
        size=111,
        mtime_ns=10,
        file_id=1,
        revision="02",
        appendix="01",
    )
    rd = _mto(
        path=rd_path,
        source=SourceKind.RD,
        size=222,
        mtime_ns=20,
        file_id=2,
        revision="02",
        appendix="01",
    )
    with tempfile.TemporaryDirectory(prefix="rd_robot_accept_") as temp:
        database = CatalogDatabase(Path(temp) / "catalog.sqlite")
        database.initialize()
        assert database.schema_version() == 14
        stored = database.upsert_robot_mto_accept(
            "1600", "POS", robot=robot, rd=rd, comment="коды"
        )
        assert stored.comment == "коды"
        assert stored.robot_path_key == robot.path_key
        assert stored.rd_size == 222
        assert stored.rd_mtime_ns == 20
        assert stored.robot_revision_text == "02-AN01"
        again = database.upsert_robot_mto_accept(
            "1600", "pos", robot=robot, rd=rd, comment="обновлено"
        )
        assert again.id == stored.id
        assert again.comment == "обновлено"
        assert len(database.list_robot_mto_accepts()) == 1
        assert robot_mto_accept_state(stored, rd=rd, robot=robot) == ACCEPT_LIVE
        edited_robot = _mto(
            path=robot_path,
            source=SourceKind.ROBOT,
            size=333,
            mtime_ns=99,
            file_id=1,
            revision="02",
            appendix="01",
        )
        assert (
            robot_mto_accept_state(stored, rd=rd, robot=edited_robot)
            == ACCEPT_LIVE
        )
        new_rd = _mto(
            path=rd_path,
            source=SourceKind.RD,
            size=222,
            mtime_ns=40,
            file_id=2,
            revision="02",
            appendix="01",
        )
        assert robot_mto_accept_state(stored, rd=new_rd, robot=robot) == ACCEPT_STALE
        other_robot = _mto(
            path=r"C:\robot\other.xlsx",
            source=SourceKind.ROBOT,
            size=111,
            mtime_ns=10,
            file_id=3,
        )
        assert (
            robot_mto_accept_state(stored, rd=rd, robot=other_robot)
            == ACCEPT_MISSING_ROBOT
        )
        assert robot_mto_accept_state(stored, rd=rd, robot=None) == ACCEPT_MISSING_ROBOT
        assert robot_mto_accept_paints_blue(
            ACCEPT_LIVE, content_equal=False
        )
        assert not robot_mto_accept_paints_blue(
            ACCEPT_LIVE, content_equal=True
        )
        assert not robot_mto_accept_paints_blue(
            ACCEPT_STALE, content_equal=False
        )
        assert robot_mto_accept_paints_green(ACCEPT_LIVE)
        assert not robot_mto_accept_paints_green(ACCEPT_STALE)
        assert not robot_mto_accept_paints_green(ACCEPT_MISSING_ROBOT)
        view = build_robot_mto_accept_view(stored, rd=new_rd, robot=robot)
        assert view.status == "устарело"
        live_view = build_robot_mto_accept_view(stored, rd=rd, robot=robot)
        assert live_view.status == "актуально"
        assert "актуально" in live_view.haystack
        database.replace_google_snapshot(
            (),
            (),
            loaded_at="2026-09-21T00:00:00+00:00",
            source="test",
        )
        database.replace_kit_derived((), (), ())
        assert database.get_robot_mto_accept("1600", "POS") is not None
        assert database.delete_robot_mto_accept("1600", "POS")
        assert database.get_robot_mto_accept("1600", "POS") is None
        assert database.list_robot_mto_accepts() == []
    print("RD catalog robot MTO accept: OK")


if __name__ == "__main__":
    main()
