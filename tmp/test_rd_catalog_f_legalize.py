"""Local checks for RD-folder F legalize (no Google, no Qt)."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.f_journal import build_journal_patch
from rd_catalog.f_legalize import (
    LEGALIZE_APPROVAL_STAGE,
    LEGALIZE_APPROVAL_TOKEN,
    f_line_date_from_mtime_ns,
    job_from_edited_comment,
    legalize_approval_f_line,
    mto_revision_from_records,
)
from rd_catalog.kits import parse_history_line
from rd_catalog.models import FileKind, FileRecord, ReviewState, SourceKind


def _record(
    *,
    file_id: int,
    file_kind: str,
    revision: str,
    appendix: str | None = None,
) -> FileRecord:
    return FileRecord(
        id=file_id,
        path=rf"\\bcc\eng\x\{file_id}",
        path_key=f"rd/{file_id}",
        source=SourceKind.RD,
        present=True,
        review_state=ReviewState.ACKNOWLEDGED,
        first_seen_run_id=1,
        last_seen_run_id=1,
        data={
            "file_kind": file_kind,
            "revision": revision,
            "appendix": appendix,
            "discipline_block": (
                "MTO-0001" if file_kind == FileKind.MTO_XLSX.value else "OD-0001"
            ),
        },
    )


def main() -> None:
    """Assert legalize F line, MTO pick, and edited F после → job."""

    line = legalize_approval_f_line(
        date="17.09.2026",
        revision="01-AN02",
        mto_revision="01",
    )
    assert line == (
        "17.09.2026 код А на рев. 01-AN02 "
        f"{LEGALIZE_APPROVAL_TOKEN} MTO 01 auto"
    )
    parsed = parse_history_line(line)
    assert parsed.stage == LEGALIZE_APPROVAL_STAGE
    assert parsed.revision == "01"
    assert parsed.appendix == "02"
    assert parsed.mto_revision == "01"
    assert parsed.from_robot_auto is True
    assert parsed.transmittals == ()

    no_mto = legalize_approval_f_line(date="17.09.2026", revision="0")
    assert "MTO Нет" in no_mto
    assert parse_history_line(no_mto).mto_absent is True

    records = (
        _record(file_id=1, file_kind="pdf", revision="01", appendix="02"),
        _record(
            file_id=2,
            file_kind=FileKind.MTO_XLSX.value,
            revision="01",
        ),
        _record(
            file_id=3,
            file_kind=FileKind.MTO_XLSX.value,
            revision="0",
        ),
    )
    assert mto_revision_from_records(records) == "01"

    before = "02.12.2024 код А на рев. 04 AGCC-BCC-TRM-000582"
    patch = build_journal_patch(
        before, line, revision="01-AN02", stage=LEGALIZE_APPROVAL_STAGE
    )
    assert patch.action == "inserted"
    assert line in patch.comment_after
    job = job_from_edited_comment(
        title="1600",
        mark="SOT",
        comment_before=patch.comment_before,
        comment_after=patch.comment_after,
        fallback_line=line,
        fallback_revision="01-AN02",
    )
    assert job.f_line == line
    assert job.stage == "code_a"
    assert job.revision == "01-AN02"

    edited = line.replace("17.09.2026", "18.09.2026")
    edited_job = job_from_edited_comment(
        title="1600",
        mark="SOT",
        comment_before=patch.comment_before,
        comment_after=before + "\n" + edited,
        fallback_line=line,
        fallback_revision="01-AN02",
    )
    assert edited_job.f_line == edited
    assert edited_job.f_line.startswith("18.09.2026")

    assert f_line_date_from_mtime_ns(None).count(".") == 2
    print("RD catalog F legalize: OK")


if __name__ == "__main__":
    main()
