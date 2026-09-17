"""Local smoke checks for deterministic RD delta overlay."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.models import CollisionKind, FileKind, SourceKind
from rd_catalog.overlay import build_overlay, revision_rank
from rd_catalog.parse import parse_catalog_file, parse_transfer_folder


def _pdf(name: str, transfer_name: str, mtime_ns: int):
    transfer = parse_transfer_folder(transfer_name)
    return parse_catalog_file(
        f"C:/fixture/{transfer_name}/{name}",
        SourceKind.RD,
        size=100,
        mtime_ns=mtime_ns,
        transfer=transfer,
    )


def main() -> None:
    """Run overlay inheritance and conflict assertions."""

    assert revision_rank(None, None) < revision_rank("V", None)
    assert revision_rank("V", None) < revision_rank("S", None)
    assert revision_rank("S", None) < revision_rank("0", None)
    assert revision_rank("0", "02") < revision_rank("01", None)

    older_a = _pdf(
        "AGCC.287-2225-KSB.OD-0001_01_RU.pdf",
        "01_рев.01_2225-KSB",
        1,
    )
    older_b = _pdf(
        "AGCC.287-2225-KSB.OD-0002_01_RU.pdf",
        "01_рев.01_2225-KSB",
        1,
    )
    newer_a = _pdf(
        "AGCC.287-2225-KSB.OD-0001_02_RU.pdf",
        "02_рев.02_2225-KSB",
        2,
    )
    conflict_a = _pdf(
        "AGCC.287-2225-KSB.OD-0001_01_RU.pdf",
        "03_рев.01_2225-KSB",
        3,
    )

    overlay = build_overlay(
        [older_a, older_b, newer_a], FileKind.PDF
    )
    assert len(overlay.current) == 2
    assert overlay.current["pdf:agcc.287-2225-ksb.od-0001"] == newer_a
    assert overlay.current["pdf:agcc.287-2225-ksb.od-0002"] == older_b

    conflict_overlay = build_overlay(
        [older_a, newer_a, conflict_a], FileKind.PDF
    )
    assert conflict_overlay.current[
        "pdf:agcc.287-2225-ksb.od-0001"
    ] == conflict_a
    assert any(
        collision.kind is CollisionKind.TRANSFER_ORDER_CONFLICT
        for collision in conflict_overlay.collisions
    )

    rev04 = _pdf(
        "AGCC.287-8950-POS1.OD-0001_04_RU.pdf",
        "10_рев.04_AGCC.287-8950-POS1",
        10,
    )
    rev_an01 = parse_catalog_file(
        "C:/fixture/10_рев.AN01_AGCC.287-8950-POS1/AGCC.287-8950-POS1.OD-0001_04-AN01_RU.pdf",
        SourceKind.RD,
        size=100,
        mtime_ns=11,
        transfer=parse_transfer_folder(
            "10_рев.AN01_AGCC.287-8950-POS1", under_gate=True
        ),
    )
    tied_sequence = build_overlay([rev04, rev_an01], FileKind.PDF)
    assert tied_sequence.current["pdf:agcc.287-8950-pos1.od-0001"] == rev_an01

    older_copy = _pdf(
        "AGCC.287-8950-POS1.OD-0001_04_RU.pdf",
        "10_рев.04_AGCC.287-8950-POS1",
        10,
    )
    newer_copy = parse_catalog_file(
        "C:/fixture/10_повтор_AGCC.287-8950-POS1/AGCC.287-8950-POS1.OD-0001_04_RU.pdf",
        SourceKind.RD,
        size=100,
        mtime_ns=20,
        transfer=parse_transfer_folder(
            "10_повтор_AGCC.287-8950-POS1", under_gate=True
        ),
    )
    same_rev_two_folders = build_overlay([older_copy, newer_copy], FileKind.PDF)
    assert same_rev_two_folders.current[
        "pdf:agcc.287-8950-pos1.od-0001"
    ] == newer_copy
    assert not any(
        collision.kind is CollisionKind.DUP_SAME_REVISION
        for collision in same_rev_two_folders.collisions
    )

    shipped_04 = _pdf(
        "AGCC.287-8950-POS1.OD-0001_04_RU.pdf",
        "10_рев.04_AGCC.287-8950-POS1",
        10,
    )
    unsent_an01 = _pdf(
        "AGCC.287-8950-POS1.OD-0001_04-AN01_RU.pdf",
        "09_рев.04-AN01_AGCC.287-8950-POS1",
        20,
    )
    inverted = build_overlay([shipped_04, unsent_an01], FileKind.PDF)
    assert inverted.current["pdf:agcc.287-8950-pos1.od-0001"] == shipped_04
    inverted_kinds = {collision.kind for collision in inverted.collisions}
    assert CollisionKind.TRANSFER_ORDER_CONFLICT in inverted_kinds
    assert CollisionKind.TRANSFER_MTIME_CONFLICT in inverted_kinds

    transfer15 = "15_рев.02-AN01_AGCC.287-8950-POS5"
    transfer_meta = parse_transfer_folder(transfer15, under_gate=True)

    def _mto(name: str, folder: str, mtime_ns: int):
        return parse_catalog_file(
            f"C:/fixture/{transfer15}/{folder}/{name}",
            SourceKind.RD,
            size=100,
            mtime_ns=mtime_ns,
            transfer=transfer_meta,
        )

    issued_mto = _mto(
        "AGCC.287-8950-POS5.MTO-0001_02-AN01_RU.xlsx", ".", 30
    )
    nested_dup = _mto(
        "AGCC.287-8950-POS5.MTO-0001_02-AN01_RU.xlsx",
        "версия из CP 03072026",
        20,
    )
    nested_old = _mto(
        "AGCC.287-8950-POS5.MTO-0001_02_RU.xlsx",
        "версия из CP 03072026",
        10,
    )
    leftover_current = build_overlay(
        [issued_mto, nested_dup, nested_old], FileKind.MTO_XLSX
    )
    mto_key = "mto:8950-pos5|mto-0001"
    assert leftover_current.current[mto_key] == issued_mto
    assert any(
        collision.kind is CollisionKind.DUP_SAME_REVISION
        for collision in leftover_current.collisions
    )
    print("RD catalog overlay smoke: OK")


if __name__ == "__main__":
    main()
