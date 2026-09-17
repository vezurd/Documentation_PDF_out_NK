"""Move an SQ kit folder into a new RD transfer package.

RD and SQ stay read-only except this user-confirmed GUI action. The SQ
title+mark folder is moved (not copied) under
``title / mark / gate / NN_рев.<rev>_от_YYYY.MM.DD``.
"""

from __future__ import annotations

import os
import re
import shutil
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from rd_catalog.kits import format_revision, kit_identity_key
from rd_catalog.models import FileRecord, SourceKind, make_path_key
from rd_catalog.overlay import revision_rank
from rd_catalog.parse import (
    is_transfer_folder_name,
    is_transfer_gate_folder_name,
    normalize_unicode_dashes,
    parse_transfer_folder,
)
from rd_catalog.path_actions import path_is_under
from rd_catalog.perf_log import perf_span

_PDF_DWG_NAMES = frozenset({"pdf", "dwg"})
_MARK_FOLDER_RE = re.compile(
    r"^(?P<seq>\d{1,2})[_.\-\s]+(?P<mark>.+)$",
    re.IGNORECASE,
)
_DEFAULT_GATE_NAME = "Для передачи"


class SqToRdError(Exception):
    """User-visible failure while planning or executing an SQ→RD move."""


@dataclass(frozen=True, slots=True)
class SqToRdPlan:
    """Preview of one kit folder move from SQ into a new RD transfer."""

    title: str
    mark: str
    source_folder: str
    destination_folder: str
    gate_folder: str
    mark_folder: str
    transfer_name: str
    sequence: int
    revision_text: str
    newest_mtime_ns: int
    catalog_file_count: int
    child_names: tuple[str, ...]
    created_gate: bool
    created_mark: bool
    created_title: bool
    rd_root: str
    sq_root: str


@dataclass(frozen=True, slots=True)
class SqToRdResult:
    """Filesystem outcome of :func:`execute_sq_to_rd_transfer`."""

    source_folder: str
    destination_folder: str
    moved_names: tuple[str, ...]
    source_removed: bool


def folder_matches_mark(folder_name: str, mark: str) -> bool:
    """Return whether a directory name is this kit's mark folder.

    Accepts ``POS1`` and ``12_POS1`` / ``04-SOT``. Does not match a different
    mark glued onto the same prefix.

    Args:
        folder_name: One path segment.
        mark: Latin AGCC mark.

    Returns:
        True when the folder belongs to ``mark``.
    """

    folded = normalize_unicode_dashes(folder_name).strip().casefold()
    want = mark.strip().casefold()
    if not folded or not want:
        return False
    if folded == want:
        return True
    match = _MARK_FOLDER_RE.match(folded)
    return bool(match and match.group("mark").casefold() == want)


def next_transfer_sequence(existing: Iterable[int]) -> int:
    """Return the next 1–2 digit transfer sequence.

    Args:
        existing: Sequences already present under the gate.

    Returns:
        ``max(existing) + 1``, or ``1`` when the gate is empty.

    Raises:
        SqToRdError: If the next sequence would exceed 99.
    """

    current = [int(value) for value in existing if int(value) >= 0]
    nxt = (max(current) + 1) if current else 1
    if nxt > 99:
        raise SqToRdError("Следующий номер передачи превысил бы 99.")
    return nxt


def transfer_folder_name(
    sequence: int,
    revision_text: str,
    when: datetime,
) -> str:
    """Build ``NN_рев.<rev>_от_YYYY.MM.DD``.

    Args:
        sequence: Leading transfer number.
        revision_text: Filename revision/AN, e.g. ``04-AN04``.
        when: Local timestamp of the newest SQ file.

    Returns:
        Folder name for the new issued package.
    """

    rev = normalize_unicode_dashes(revision_text).strip()
    date_text = when.strftime("%Y.%m.%d")
    return f"{sequence:02d}_рев.{rev}_от_{date_text}"


def datetime_from_mtime_ns(mtime_ns: int) -> datetime:
    """Convert a nanosecond mtime to local time.

    Args:
        mtime_ns: Filesystem mtime in nanoseconds.

    Returns:
        Local ``datetime``.
    """

    return datetime.fromtimestamp(int(mtime_ns) / 1_000_000_000)


def plan_summary_text(plan: SqToRdPlan) -> str:
    """Build the confirmation-dialog body.

    Args:
        plan: Validated move plan.

    Returns:
        Multi-line Russian description of source and destination.
    """

    when = datetime_from_mtime_ns(plan.newest_mtime_ns)
    created: list[str] = []
    if plan.created_title:
        created.append(f"титул {plan.title}")
    if plan.created_mark:
        created.append(f"марку {Path(plan.mark_folder).name}")
    if plan.created_gate:
        created.append(f"шлюз «{_DEFAULT_GATE_NAME}»")
    created_line = (
        f"Будут созданы: {', '.join(created)}."
        if created
        else "Папка марки и шлюз «Для передачи» уже есть."
    )
    children = ", ".join(plan.child_names) or "—"
    return "\n".join(
        [
            f"Комплект: {plan.title}-{plan.mark}",
            "",
            "Откуда перемещаем (SQ)",
            f"  Папка: {plan.source_folder}",
            f"  Содержимое: {children}",
            f"  Файлов комплекта в каталоге: {plan.catalog_file_count}",
            f"  Ревизия SQ: {plan.revision_text}",
            f"  Самый свежий файл: {when.strftime('%Y.%m.%d %H:%M')}",
            "",
            "Куда в РД",
            f"  Марка: {plan.mark_folder}",
            f"  Шлюз: {plan.gate_folder}",
            f"  Новая передача: {plan.transfer_name}",
            f"  Полный путь: {plan.destination_folder}",
            "",
            created_line,
            "Папка SQ будет перемещена целиком (не копия). "
            "После переноса запустится сканирование РД и SQ.",
        ]
    )


def _kit_records(
    records: Iterable[FileRecord],
    source: SourceKind,
    title: str,
    mark: str,
) -> list[FileRecord]:
    key = kit_identity_key(title, mark)
    matched: list[FileRecord] = []
    for record in records:
        if record.source is not source or not record.present:
            continue
        rec_title = str(record.data.get("title") or "").strip()
        rec_mark = str(record.data.get("mark") or "").strip()
        if kit_identity_key(rec_title, rec_mark) != key:
            continue
        matched.append(record)
    return matched


def _package_dir_for_file(path: Path) -> Path:
    parent = path.parent
    if parent.name.casefold() in _PDF_DWG_NAMES:
        return parent.parent
    return parent


def _sq_kit_folder(
    paths: Iterable[str],
    sq_root: str,
    *,
    records: Iterable[FileRecord],
    title: str,
    mark: str,
) -> Path:
    files = [Path(path) for path in paths if path]
    if not files:
        raise SqToRdError(f"Нет файлов SQ для {title}-{mark}.")
    packages = [_package_dir_for_file(path) for path in files]
    try:
        common = Path(os.path.commonpath([str(item) for item in packages]))
    except ValueError as exc:
        raise SqToRdError(
            f"Файлы SQ комплекта {title}-{mark} лежат на разных дисках."
        ) from exc
    if not path_is_under(common, sq_root):
        raise SqToRdError(
            f"Папка SQ не находится в корне SQ:\n{common}"
        )
    if make_path_key(common) == make_path_key(sq_root):
        raise SqToRdError(
            "Файлы SQ лежат в корне «Ответы на SQ запросы» — "
            "нельзя перемещать весь каталог SQ."
        )
    foreign = _foreign_kit_in_folder(common, records, title, mark, SourceKind.SQ)
    if foreign:
        raise SqToRdError(
            "В папке SQ есть файлы другого комплекта "
            f"({foreign}). Перенос отменён."
        )
    return common


def _foreign_kit_in_folder(
    folder: Path,
    records: Iterable[FileRecord],
    title: str,
    mark: str,
    source: SourceKind,
) -> str | None:
    key = kit_identity_key(title, mark)
    for record in records:
        if record.source is not source or not record.present:
            continue
        if not path_is_under(record.path, folder):
            continue
        rec_title = str(record.data.get("title") or "").strip()
        rec_mark = str(record.data.get("mark") or "").strip()
        if kit_identity_key(rec_title, rec_mark) != key:
            return f"{rec_title}-{rec_mark}"
    return None


def _layout_from_rd_path(
    path: str,
    rd_root: str,
    title: str,
    mark: str,
) -> tuple[Path | None, Path | None]:
    """Return ``(mark_dir, gate_dir)`` parsed from one RD file path."""

    try:
        relative = Path(path).relative_to(Path(rd_root))
    except ValueError:
        return None, None
    dirs = list(relative.parts[:-1])
    title_index = next(
        (
            index
            for index, part in enumerate(dirs)
            if part.strip() == title
        ),
        None,
    )
    if title_index is None:
        return None, None
    mark_dir: Path | None = None
    gate_dir: Path | None = None
    title_dir = Path(rd_root, *dirs[: title_index + 1])
    for offset, part in enumerate(dirs[title_index + 1 :], start=title_index + 1):
        current = Path(rd_root, *dirs[: offset + 1])
        if mark_dir is None and folder_matches_mark(part, mark):
            mark_dir = current
            continue
        if mark_dir is not None and is_transfer_gate_folder_name(part):
            gate_dir = current
            break
    if mark_dir is None:
        mark_dir = title_dir
    return mark_dir, gate_dir


def _list_mark_candidates(title_dir: Path, mark: str) -> list[Path]:
    if not title_dir.is_dir():
        return []
    matched: list[Path] = []
    try:
        children = list(title_dir.iterdir())
    except OSError:
        return []
    for child in children:
        if child.is_dir() and folder_matches_mark(child.name, mark):
            matched.append(child)
    return matched


def _list_gates(mark_dir: Path) -> list[Path]:
    if not mark_dir.is_dir():
        return []
    gates: list[Path] = []
    try:
        children = list(mark_dir.iterdir())
    except OSError:
        return []
    for child in children:
        if child.is_dir() and is_transfer_gate_folder_name(child.name):
            gates.append(child)
    return gates


def _sequences_in_gate(gate_dir: Path) -> list[int]:
    if not gate_dir.is_dir():
        return []
    sequences: list[int] = []
    try:
        children = list(gate_dir.iterdir())
    except OSError:
        return []
    for child in children:
        if not child.is_dir():
            continue
        if not is_transfer_folder_name(child.name, under_gate=True):
            continue
        parsed = parse_transfer_folder(child.name, under_gate=True)
        if parsed.sequence is not None:
            sequences.append(int(parsed.sequence))
    return sequences


def _pick_gate(gates: list[Path]) -> Path | None:
    if not gates:
        return None
    if len(gates) == 1:
        return gates[0]

    def sort_key(folder: Path) -> tuple[int, int, str]:
        seqs = _sequences_in_gate(folder)
        max_seq = max(seqs) if seqs else -1
        prefers_transfer = 1 if "передач" in folder.name.casefold() else 0
        return (max_seq, prefers_transfer, folder.name.casefold())

    return max(gates, key=sort_key)


def _resolve_rd_destination(
    *,
    title: str,
    mark: str,
    rd_root: str,
    rd_paths: Iterable[str],
) -> tuple[Path, Path, bool, bool, bool]:
    """Return mark dir, gate dir, and whether title/mark/gate will be created."""

    rd_root_path = Path(rd_root)
    mark_dirs: list[Path] = []
    gate_dirs: list[Path] = []
    for path in rd_paths:
        mark_dir, gate_dir = _layout_from_rd_path(path, rd_root, title, mark)
        if mark_dir is not None:
            mark_dirs.append(mark_dir)
        if gate_dir is not None:
            gate_dirs.append(gate_dir)

    unique_marks = {
        make_path_key(item): item for item in mark_dirs if item is not None
    }
    if len(unique_marks) > 1:
        names = ", ".join(sorted(str(item) for item in unique_marks.values()))
        raise SqToRdError(
            f"Для {title}-{mark} в РД несколько папок марки:\n{names}"
        )

    created_title = False
    created_mark = False
    created_gate = False
    title_dir = rd_root_path / title
    mark_dir = next(iter(unique_marks.values()), None)
    if mark_dir is None:
        live = _list_mark_candidates(title_dir, mark)
        if len(live) > 1:
            names = ", ".join(child.name for child in live)
            raise SqToRdError(
                f"В {title_dir} несколько папок марки {mark}: {names}"
            )
        if live:
            mark_dir = live[0]
        else:
            mark_dir = title_dir / mark
            created_mark = True
            created_title = not title_dir.is_dir()

    if make_path_key(mark_dir) == make_path_key(title_dir):
        live = _list_mark_candidates(title_dir, mark)
        if live:
            mark_dir = live[0]
        else:
            mark_dir = title_dir / mark
            created_mark = True

    unique_gates = {
        make_path_key(item): item
        for item in gate_dirs
        if item is not None and path_is_under(item, mark_dir)
    }
    gate_dir = next(iter(unique_gates.values()), None)
    if gate_dir is None:
        live_gates = _list_gates(mark_dir)
        gate_dir = _pick_gate(live_gates)
    if gate_dir is None:
        gate_dir = mark_dir / _DEFAULT_GATE_NAME
        created_gate = True
    return mark_dir, gate_dir, created_title, created_mark, created_gate


def _newest_mtime_ns(folder: Path, fallback: int | None) -> int:
    newest = int(fallback or 0)
    try:
        for current_root, _dirs, names in os.walk(folder):
            for name in names:
                path = os.path.join(current_root, name)
                try:
                    mtime_ns = int(os.stat(path, follow_symlinks=False).st_mtime_ns)
                except OSError:
                    continue
                if mtime_ns > newest:
                    newest = mtime_ns
    except OSError:
        pass
    if newest <= 0:
        raise SqToRdError(
            f"Не удалось определить дату файлов в папке SQ:\n{folder}"
        )
    return newest


def _revision_text(
    records: list[FileRecord],
    fallback: str,
) -> str:
    text = (fallback or "").strip()
    if text:
        return text
    best_rank: tuple[int, int, str] | None = None
    best_text = ""
    for record in records:
        revision = str(record.data.get("revision") or "").strip() or None
        appendix = str(record.data.get("appendix") or "").strip() or None
        formatted = format_revision(revision, appendix)
        if not formatted:
            continue
        rank = revision_rank(revision, appendix)
        if best_rank is None or rank > best_rank:
            best_rank = rank
            best_text = formatted
    if not best_text:
        raise SqToRdError("У файлов SQ нет разобранной ревизии имени.")
    return best_text


def plan_sq_to_rd_transfer(
    *,
    title: str,
    mark: str,
    records: Iterable[FileRecord],
    sq_paths: Iterable[str],
    rd_paths: Iterable[str],
    sq_revision_text: str = "",
    sq_max_mtime_ns: int | None = None,
    rd_root: str | Path,
    sq_root: str | Path,
) -> SqToRdPlan:
    """Build a move plan for one kits-row title+mark.

    Args:
        title: Four-digit title.
        mark: Latin AGCC mark.
        records: Persisted catalog files.
        sq_paths: Known SQ file paths for this kit.
        rd_paths: Known RD file paths for this kit (any overlay state).
        sq_revision_text: Display revision from the kits snapshot.
        sq_max_mtime_ns: Snapshot max mtime, used if the folder walk fails.
        rd_root: Configured RD root.
        sq_root: Configured SQ root.

    Returns:
        A validated plan that writes only under ``rd_root`` and reads SQ.

    Raises:
        SqToRdError: If the move would be unsafe or destinations collide.
    """

    with perf_span("sq_rd.plan", title=title, mark=mark):
        rd_root_text = str(rd_root)
        sq_root_text = str(sq_root)
        materialized = list(records)
        sq_records = _kit_records(materialized, SourceKind.SQ, title, mark)
        rd_records = _kit_records(materialized, SourceKind.RD, title, mark)
        all_sq_paths = list(sq_paths) or [item.path for item in sq_records]
        all_rd_paths = list(rd_paths) or [item.path for item in rd_records]
        if not all_sq_paths:
            raise SqToRdError(f"Нет файлов SQ для {title}-{mark}.")

        source_folder = _sq_kit_folder(
            all_sq_paths,
            sq_root_text,
            records=materialized,
            title=title,
            mark=mark,
        )
        if not source_folder.is_dir():
            raise SqToRdError(f"Папка SQ не найдена:\n{source_folder}")

        (
            mark_dir,
            gate_dir,
            created_title,
            created_mark,
            created_gate,
        ) = _resolve_rd_destination(
            title=title,
            mark=mark,
            rd_root=rd_root_text,
            rd_paths=all_rd_paths,
        )
        if path_is_under(gate_dir, sq_root_text):
            raise SqToRdError(
                "Папка назначения попала бы в каталог SQ — перенос отменён."
            )
        if not path_is_under(gate_dir, rd_root_text):
            raise SqToRdError(
                f"Шлюз передачи не находится в корне РД:\n{gate_dir}"
            )

        revision_text = _revision_text(sq_records, sq_revision_text)
        newest_mtime_ns = _newest_mtime_ns(source_folder, sq_max_mtime_ns)
        when = datetime_from_mtime_ns(newest_mtime_ns)
        sequence = next_transfer_sequence(_sequences_in_gate(gate_dir))
        transfer_name = transfer_folder_name(sequence, revision_text, when)
        destination = gate_dir / transfer_name
        if destination.exists():
            raise SqToRdError(
                f"Папка передачи уже существует:\n{destination}"
            )
        if not path_is_under(destination, rd_root_text):
            raise SqToRdError(
                f"Путь назначения не находится в корне РД:\n{destination}"
            )
        if path_is_under(destination, sq_root_text):
            raise SqToRdError("Путь назначения находится в каталоге SQ.")
        if path_is_under(destination, source_folder) or path_is_under(
            source_folder, destination
        ):
            raise SqToRdError("Папка SQ и папка назначения пересекаются.")

        try:
            children = sorted(child.name for child in source_folder.iterdir())
        except OSError as exc:
            raise SqToRdError(
                f"Не удалось прочитать папку SQ:\n{source_folder}\n{exc}"
            ) from exc
        if not children:
            raise SqToRdError(f"Папка SQ пуста:\n{source_folder}")

        return SqToRdPlan(
            title=title,
            mark=mark,
            source_folder=str(source_folder),
            destination_folder=str(destination),
            gate_folder=str(gate_dir),
            mark_folder=str(mark_dir),
            transfer_name=transfer_name,
            sequence=sequence,
            revision_text=revision_text,
            newest_mtime_ns=newest_mtime_ns,
            catalog_file_count=len(sq_records) or len(all_sq_paths),
            child_names=tuple(children),
            created_gate=created_gate,
            created_mark=created_mark,
            created_title=created_title,
            rd_root=rd_root_text,
            sq_root=sq_root_text,
        )


def execute_sq_to_rd_transfer(plan: SqToRdPlan) -> SqToRdResult:
    """Create the RD transfer folder and move SQ children into it.

    Args:
        plan: Validated move plan.

    Returns:
        Destination path and the names that were moved.

    Raises:
        SqToRdError: If a filesystem check fails or a move is unsafe.
    """

    with perf_span("sq_rd.execute"):
        source = Path(plan.source_folder)
        dest = Path(plan.destination_folder)
        gate = Path(plan.gate_folder)
        if not path_is_under(dest, plan.rd_root):
            raise SqToRdError("Путь назначения не находится в корне РД.")
        if path_is_under(dest, plan.sq_root):
            raise SqToRdError("Путь назначения находится в каталоге SQ.")
        if not path_is_under(source, plan.sq_root):
            raise SqToRdError("Исходная папка не находится в каталоге SQ.")
        if not source.is_dir():
            raise SqToRdError(f"Папка SQ не найдена:\n{source}")
        if dest.exists():
            raise SqToRdError(f"Папка передачи уже существует:\n{dest}")

        try:
            gate.mkdir(parents=True, exist_ok=True)
            dest.mkdir(parents=False, exist_ok=False)
        except OSError as exc:
            raise SqToRdError(
                f"Не удалось создать папку передачи:\n{dest}\n{exc}"
            ) from exc

        moved: list[str] = []
        try:
            children = list(source.iterdir())
        except OSError as exc:
            raise SqToRdError(
                f"Не удалось прочитать папку SQ:\n{source}\n{exc}"
            ) from exc
        for child in children:
            target = dest / child.name
            if target.exists():
                raise SqToRdError(
                    f"В папке назначения уже есть {child.name}:\n{target}"
                )
            try:
                shutil.move(str(child), str(target))
            except OSError as exc:
                raise SqToRdError(
                    f"Не удалось переместить {child.name}:\n{child} → {target}\n{exc}"
                ) from exc
            moved.append(child.name)

        source_removed = False
        try:
            remaining = list(source.iterdir())
        except OSError:
            remaining = [source]
        if not remaining:
            try:
                source.rmdir()
                source_removed = True
            except OSError:
                source_removed = False
        return SqToRdResult(
            source_folder=str(source),
            destination_folder=str(dest),
            moved_names=tuple(moved),
            source_removed=source_removed,
        )
