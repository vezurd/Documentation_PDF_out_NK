"""Manual accept that a robot MTO corresponds to the official RD MTO.

The row lives in SQLite (``robot_mto_accept``). Live/stale uses the official
folder MTO ``path_key`` + size + disk mtime, not workbook bytes. Origin
green and content-equal stay independent.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path, PureWindowsPath

from rd_catalog.db import RobotMtoAcceptRow, mto_file_stat_signature
from rd_catalog.doc_bundle import format_file_save_date
from rd_catalog.kits import format_revision, kit_identity_key
from rd_catalog.models import FileKind, FileRecord, make_path_key

ROBOT_MTO_ACCEPT_FOREGROUND = "#1A4FBF"

ACCEPT_LIVE = "live"
ACCEPT_STALE = "stale"
ACCEPT_MISSING_ROBOT = "missing_robot"

ACCEPT_STATUS_LABELS = {
    ACCEPT_LIVE: "актуально",
    ACCEPT_STALE: "устарело",
    ACCEPT_MISSING_ROBOT: "нет файла",
}

ACCEPT_TAB_HEADERS = (
    "Титул",
    "Марка",
    "Ревизия робота",
    "MTO · рев.",
    "Статус",
    "Подтверждено",
    "Файл робота",
    "Эталон РД",
    "Путь робота",
)

ACCEPT_LIVE_TOOLTIP = (
    "MTO робота подтверждён вручную как актуальный относительно "
    "текущего MTO РД официальной папки. Содержимое может отличаться "
    "(коды закупки, теги). Это не копия по дате и не сверка байтов."
)


def path_keys_match(left: str, right: str) -> bool:
    """Return whether two catalog path keys name the same file."""

    a = make_path_key(left)
    b = make_path_key(right)
    if a and b:
        return a == b
    return str(left or "").casefold() == str(right or "").casefold()


def record_revision_text(record: FileRecord | None) -> str:
    """Return filename revision of a catalog file, or empty."""

    if record is None:
        return ""
    return format_revision(
        record.data.get("revision"),
        record.data.get("appendix"),
    )


def mto_xlsx_record(
    paths: Sequence[str],
    records_by_path_key: Mapping[str, FileRecord],
) -> FileRecord | None:
    """Return the first MTO xlsx among ``paths`` that is in the index."""

    for path in paths:
        key = make_path_key(path)
        record = records_by_path_key.get(key)
        if record is None:
            record = records_by_path_key.get(str(path).casefold())
        if record is None:
            continue
        kind = str(record.data.get("file_kind") or "").casefold()
        if kind == FileKind.MTO_XLSX.value:
            return record
    return None


def _rd_fingerprint_matches(stored: RobotMtoAcceptRow, rd: FileRecord) -> bool:
    signature = mto_file_stat_signature(rd)
    return (
        path_keys_match(stored.rd_path_key, str(signature.get("path_key") or ""))
        and int(stored.rd_size) == int(signature.get("size") or 0)
        and int(stored.rd_mtime_ns) == int(signature.get("mtime_ns") or 0)
    )


def robot_mto_accept_state(
    stored: RobotMtoAcceptRow | None,
    *,
    rd: FileRecord | None,
    robot: FileRecord | None,
) -> str:
    """Return ``live`` / ``stale`` / ``missing_robot``, or empty.

    Args:
        stored: Persisted accept, if any.
        rd: Current official-folder RD MTO xlsx.
        robot: Current present robot MTO xlsx.

    Returns:
        Status token, or ``""`` when there is no stored row.
    """

    if stored is None:
        return ""
    if robot is None or not path_keys_match(robot.path_key, stored.robot_path_key):
        return ACCEPT_MISSING_ROBOT
    if rd is None or not _rd_fingerprint_matches(stored, rd):
        return ACCEPT_STALE
    return ACCEPT_LIVE


def robot_mto_accept_paints_blue(
    state: str,
    *,
    content_equal: bool | None,
) -> bool:
    """Return whether Комплекты should use blue robot-revision text."""

    return state == ACCEPT_LIVE and content_equal is not True


def robot_mto_accept_tooltip(
    stored: RobotMtoAcceptRow,
    *,
    rd: FileRecord | None,
) -> str:
    """Return the kits-cell note for a live accept."""

    revision = (stored.rd_revision_text or record_revision_text(rd) or "—").strip()
    stamp = ""
    if rd is not None:
        stamp = format_file_save_date(
            int(rd.data.get("mtime_ns") or 0)
        ) or ""
    if not stamp:
        stamp = format_file_save_date(int(stored.rd_mtime_ns or 0)) or ""
    path = stored.rd_path or (rd.path if rd is not None else "")
    date_bit = f", дата {stamp}" if stamp else ""
    lines = [
        ACCEPT_LIVE_TOOLTIP,
        f"Эталон MTO РД {revision}{date_bit}.",
    ]
    if path:
        lines.append(path)
    return "\n".join(lines)


@dataclass(frozen=True, slots=True)
class RobotMtoAcceptView:
    """One tab row: stored accept plus live/stale vs current files."""

    title: str
    mark: str
    status: str
    robot_revision_text: str
    rd_revision_text: str
    decided_at: str
    robot_name: str
    rd_name: str
    robot_path: str
    rd_path: str
    accept: RobotMtoAcceptRow
    haystack: str


def _file_name(path: str) -> str:
    text = str(path or "").strip()
    if not text:
        return ""
    return PureWindowsPath(text).name or Path(text).name


def _decided_label(decided_at: str) -> str:
    raw = str(decided_at or "").strip()
    if not raw:
        return ""
    stamp = raw.replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(stamp)
    except ValueError:
        return raw[:19].replace("T", " ")
    return parsed.strftime("%Y.%m.%d %H:%M")


def build_robot_mto_accept_view(
    stored: RobotMtoAcceptRow,
    *,
    rd: FileRecord | None,
    robot: FileRecord | None,
) -> RobotMtoAcceptView:
    """Join one stored accept with the current official/robot MTO files.

    Args:
        stored: SQLite row.
        rd: Current official-folder RD MTO, if present.
        robot: Current robot MTO, if present.

    Returns:
        Tab row with Russian status and a search haystack.
    """

    state = robot_mto_accept_state(stored, rd=rd, robot=robot)
    status = ACCEPT_STATUS_LABELS.get(state, state or "—")
    robot_rev = (
        record_revision_text(robot)
        or stored.robot_revision_text
        or "—"
    )
    rd_rev = record_revision_text(rd) or stored.rd_revision_text or "—"
    robot_path = (
        robot.path if robot is not None else stored.robot_path
    )
    rd_path = rd.path if rd is not None else stored.rd_path
    decided = _decided_label(stored.decided_at)
    view = RobotMtoAcceptView(
        title=stored.title,
        mark=stored.mark,
        status=status,
        robot_revision_text=robot_rev,
        rd_revision_text=rd_rev,
        decided_at=decided,
        robot_name=_file_name(robot_path),
        rd_name=_file_name(rd_path),
        robot_path=robot_path,
        rd_path=rd_path,
        accept=stored,
        haystack="",
    )
    haystack = " ".join(
        part
        for part in (
            view.title,
            view.mark,
            f"{view.title}-{view.mark}",
            view.status,
            state,
            view.robot_revision_text,
            view.rd_revision_text,
            view.decided_at,
            view.robot_name,
            view.rd_name,
            view.robot_path,
            view.rd_path,
        )
        if part
    ).casefold()
    return RobotMtoAcceptView(
        title=view.title,
        mark=view.mark,
        status=view.status,
        robot_revision_text=view.robot_revision_text,
        rd_revision_text=view.rd_revision_text,
        decided_at=view.decided_at,
        robot_name=view.robot_name,
        rd_name=view.rd_name,
        robot_path=view.robot_path,
        rd_path=view.rd_path,
        accept=view.accept,
        haystack=haystack,
    )


def accepts_by_kit(
    rows: Sequence[RobotMtoAcceptRow],
) -> dict[tuple[str, str], RobotMtoAcceptRow]:
    """Index accepts by kit identity."""

    return {
        kit_identity_key(row.title, row.mark): row for row in rows
    }
