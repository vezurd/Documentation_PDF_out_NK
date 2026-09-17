"""Read-only impact estimate of a strict RD layout rule.

Strict rule: a file counts only when its path is
``rd_root / <4-digit title> / <MARK> / <gate> / <NN_…> / [ | PDF | DWG | BBB] / file``.
Uses the existing :func:`has_canonical_rd_issued_path`, with ``BBB`` added to the
allowed media folders, so the estimate matches the code that would enforce it.
"""

from __future__ import annotations

import json
import os
import sqlite3
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog import parse as rd_parse  # noqa: E402
from rd_catalog.config import load_config  # noqa: E402

config = load_config()
RD_ROOT = str(config.rd_root)
DB = Path(os.environ["LOCALAPPDATA"]) / "Documentation_PDF_out_NK" / "rd_catalog" / "rd_catalog.sqlite"
OUT = Path(__file__).resolve().parent / "diag_strict_scan.txt"

conn = sqlite3.connect(f"file:{DB}?mode=ro", uri=True)
conn.row_factory = sqlite3.Row

banned: set[tuple[str, str]] = set()
banned_path = DB.parent / "banned_title_marks.json"
if banned_path.exists():
    raw = json.loads(banned_path.read_text(encoding="utf-8"))
    items = raw.get("pairs", raw) if isinstance(raw, dict) else raw
    for item in items:
        if isinstance(item, dict):
            banned.add((str(item.get("title", "")).casefold(), str(item.get("mark", "")).casefold()))
        elif isinstance(item, (list, tuple)) and len(item) == 2:
            banned.add((str(item[0]).casefold(), str(item[1]).casefold()))

rows = conn.execute(
    "SELECT id, path, path_key, file_kind, title, mark, present, transfer_name "
    "FROM file_entry WHERE source='rd' AND present=1"
).fetchall()

kept_by_kit: Counter[tuple[str, str]] = Counter()
dropped_by_kit: Counter[tuple[str, str]] = Counter()
kept_kinds: Counter[str] = Counter()
dropped_kinds: Counter[str] = Counter()
dropped_shapes: Counter[str] = Counter()
kept_ids: set[int] = set()
examples: list[str] = []


def shape_of(path: str) -> str:
    """Rough reason a path fails the strict rule, for the report."""

    relative = rd_parse._relative_windows_parts(path, RD_ROOT)
    if relative is None:
        return "вне rd_root"
    dirs = list(relative[:-1])
    if dirs and dirs[-1].casefold() in rd_parse._PACKAGE_MEDIA_FOLDERS:
        dirs.pop()
    if len(dirs) < 4:
        return f"слишком мелкая вложенность ({len(dirs)} уровня)"
    if len(dirs) > 4:
        return f"лишние подпапки ({len(dirs)} уровня)"
    title, mark, gate, transfer = dirs
    if not rd_parse._TITLE_ONLY_FOLDER_RE.fullmatch(
        rd_parse.normalize_unicode_dashes(title).strip()
    ):
        return "первый уровень не 4-значный титул"
    if not rd_parse.is_transfer_gate_folder_name(gate):
        return "нет папки-шлюза (Для передачи / На_отправку)"
    if not rd_parse.is_transfer_folder_name(transfer, under_gate=True):
        return "папка пакета не начинается с NN"
    return "прочее"


for row in rows:
    kit = (str(row["title"] or "").casefold(), str(row["mark"] or "").casefold())
    if kit in banned:
        continue
    if rd_parse.has_canonical_rd_issued_path(row["path"], RD_ROOT):
        kept_by_kit[kit] += 1
        kept_kinds[row["file_kind"]] += 1
        kept_ids.add(row["id"])
    else:
        dropped_by_kit[kit] += 1
        dropped_kinds[row["file_kind"]] += 1
        reason = shape_of(row["path"])
        dropped_shapes[reason] += 1
        if len(examples) < 30:
            examples.append(f"[{reason}] {row['path']}")

all_kits = set(kept_by_kit) | set(dropped_by_kit)
kits_fully_lost = sorted(kit for kit in all_kits if kept_by_kit[kit] == 0)
kits_partly = sorted(kit for kit in all_kits if kept_by_kit[kit] and dropped_by_kit[kit])
kits_clean = sorted(kit for kit in all_kits if kept_by_kit[kit] and not dropped_by_kit[kit])

total_files = sum(kept_by_kit.values()) + sum(dropped_by_kit.values())

lines = [
    "СТРОГИЙ СКАН: ТИТУЛ / МАРКА / <шлюз> / NN_ / [корень | PDF | DWG | BBB]",
    "",
    f"Присутствующих файлов РД (без забаненных): {total_files}",
    f"  остаются в каталоге : {sum(kept_by_kit.values())} ({100*sum(kept_by_kit.values())/max(total_files,1):.1f}%)",
    f"  отсеиваются         : {sum(dropped_by_kit.values())} ({100*sum(dropped_by_kit.values())/max(total_files,1):.1f}%)",
    "",
    f"Комплектов затронуто: {len(all_kits)}",
    f"  полностью чистые (ничего не теряют)      : {len(kits_clean)}",
    f"  теряют часть файлов                      : {len(kits_partly)}",
    f"  ТЕРЯЮТ ВСЁ (уйдут в «есть в Google, нет в РД»): {len(kits_fully_lost)}",
    "",
    "Что отсеивается, по типам файлов:",
]
for kind, count in dropped_kinds.most_common():
    lines.append(f"  {count:6d}  {kind}")
lines.append("")
lines.append("Причины отсева:")
for reason, count in dropped_shapes.most_common():
    lines.append(f"  {count:6d}  {reason}")
lines.append("")
lines.append("Комплекты, теряющие ВСЕ файлы:")
for kit in kits_fully_lost:
    lines.append(f"  {kit[0]}/{kit[1]}  (отсеяно {dropped_by_kit[kit]} файлов)")
lines.append("")
lines.append("Примеры отсеиваемых путей:")
lines.extend("  " + item for item in examples)

# Does the strict rule reduce the MTO ambiguity we measured earlier?
path_by_key = {row["path_key"]: row["path"] for row in rows}
dup_before = dup_after = 0
for row in conn.execute(
    "SELECT path_keys_json FROM current_collision "
    "WHERE scope='overlay_mto' AND kind='dup_same_revision'"
):
    keys = json.loads(row["path_keys_json"] or "[]")
    paths = [path_by_key.get(key) for key in keys]
    if any(path is None for path in paths):
        continue
    dup_before += 1
    survivors = [p for p in paths if rd_parse.has_canonical_rd_issued_path(p, RD_ROOT)]
    if len(survivors) > 1:
        dup_after += 1

lines.append("")
lines.append("Влияние на неоднозначность MTO (дубли одной ревизии):")
lines.append(f"  сейчас  : {dup_before}")
lines.append(f"  останется: {dup_after}")

OUT.write_text("\n".join(lines), encoding="utf-8")
print("\n".join(lines[: 24 + len(dropped_kinds) + len(dropped_shapes)]))
print(f"\n(полный отчёт: {OUT})")
