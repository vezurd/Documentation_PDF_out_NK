"""Second-pass: unique gate names and non-gate trees after title/mark."""

from __future__ import annotations

import json
import os
import re
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path

from rd_catalog.parse import is_transfer_folder_name, normalize_unicode_dashes

DB_PATH = Path(
    os.path.expandvars(
        r"%LOCALAPPDATA%\Documentation_PDF_out_NK\rd_catalog\rd_catalog.sqlite"
    )
)
OUT_TXT = Path(__file__).with_name("analyze_rd_transfer_gates_v2.txt")
OUT_JSON = Path(__file__).with_name("analyze_rd_transfer_gates_v2.json")

TITLE_RE = re.compile(r"^\d{4}$")
MARK_HINT_RE = re.compile(
    r"(ksb|pos|spp|sos|skud|sot|soo|kbi|sagd)",
    re.IGNORECASE,
)
GATE_RE = re.compile(r"передач|отправк", re.IGNORECASE)
NOISE_RE = re.compile(
    r"(замечан|from|work|работ|исходн|интерфейс|мдз|спек|"
    r"закупк|прочее|типов|оборуд|вспомогат|наработ|"
    r"корректир|загрузк|sent|comment|ид[_ ]|from)",
    re.IGNORECASE,
)
SOURCE_DUMP_RE = re.compile(r"^\d{4}\s*[-_].+", re.IGNORECASE)


def _parts(path: str) -> list[str]:
    return [p for p in Path(path).parts if p not in ("\\", "/")]


def _after_rd(path: str) -> list[str] | None:
    parts = _parts(path)
    for i, part in enumerate(parts):
        if part.casefold() == "рд":
            return parts[i + 1 :]
    return None


def _norm(name: str) -> str:
    return normalize_unicode_dashes(name).strip().casefold()


def _is_title(name: str) -> bool:
    return bool(TITLE_RE.fullmatch(normalize_unicode_dashes(name).strip()))


def _is_mark(name: str) -> bool:
    n = normalize_unicode_dashes(name).strip()
    if _is_title(n) or is_transfer_folder_name(n):
        return False
    return bool(MARK_HINT_RE.search(n))


def _gate_class(name: str) -> str:
    n = _norm(name)
    if GATE_RE.search(n):
        return "gate_synonym"
    if is_transfer_folder_name(name):
        return "direct_transfer"
    if SOURCE_DUMP_RE.match(name) and MARK_HINT_RE.search(name):
        return "source_dump_like_1055_skud"
    if NOISE_RE.search(n):
        return "noise_semantic"
    if n in {"pdf", "dwg", "xlsx"}:
        return "file_container"
    return "unknown"


def main() -> None:
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    rows = list(
        conn.execute(
            """
            SELECT path, file_kind, title, mark
            FROM file_entry
            WHERE source = 'rd' AND present = 1
            """
        )
    )

    top_level = Counter()
    gate_names = Counter()
    gate_by_class = defaultdict(Counter)
    unknown_names = Counter()
    source_dump_examples = defaultdict(list)
    unknown_examples = defaultdict(list)
    no_gate_title_mark = Counter()  # (title, mark_folder, first_after)
    files_8441_skud = []
    unique_dirs_no_gate = set()

    file_totals = Counter()
    dir_seen = set()

    for row in rows:
        parent = str(Path(row["path"]).parent)
        rel = _after_rd(row["path"])
        if not rel:
            continue
        top_level[rel[0]] += 1
        file_totals["files"] += 1

        # Skip SQ-looking trees under RD that aren't the configured sq_root
        # (9000/SQ etc.) — still counted, classified later.

        if len(rel) < 2:
            continue

        title_folder = rel[0] if _is_title(rel[0]) else None
        mark_i = None
        if title_folder is not None:
            # first mark-like after title
            for i, part in enumerate(rel[1:], start=1):
                if _is_mark(part):
                    mark_i = i
                    break
            if mark_i is None and not is_transfer_folder_name(rel[1]) and not GATE_RE.search(rel[1]):
                mark_i = 1
        if mark_i is None:
            continue
        after = rel[mark_i + 1 : -1]  # drop filename
        mark_folder = rel[mark_i]
        first = after[0] if after else None

        if title_folder == "8441" and "skud" in _norm(mark_folder):
            files_8441_skud.append(row["path"])

        if first is None:
            cls = "files_in_mark_folder"
            key = "(files in mark folder)"
        else:
            cls = _gate_class(first)
            key = first
            if cls == "gate_synonym":
                gate_names[first] += 1
            elif cls == "unknown":
                unknown_names[first] += 1
                if len(unknown_examples[first]) < 3:
                    unknown_examples[first].append(parent)
            elif cls == "source_dump_like_1055_skud":
                if len(source_dump_examples[first]) < 4:
                    source_dump_examples[first].append(parent)

        gate_by_class[cls][key] += 1

        if cls not in {"gate_synonym", "direct_transfer"}:
            no_gate_title_mark[(title_folder or "?", mark_folder, first or "(none)")] += 1
            unique_dirs_no_gate.add(parent)

        dir_seen.add(parent)

    # Unique first-after-mark names regardless of class, for catalog
    first_after = Counter()
    first_after_ex = defaultdict(list)
    for row in rows:
        rel = _after_rd(row["path"])
        if not rel or not _is_title(rel[0]):
            continue
        mark_i = None
        for i, part in enumerate(rel[1:], start=1):
            if _is_mark(part):
                mark_i = i
                break
        if mark_i is None:
            continue
        after = rel[mark_i + 1 : -1]
        if not after:
            continue
        first_after[after[0]] += 1
        if len(first_after_ex[after[0]]) < 2:
            first_after_ex[after[0]].append(str(Path(row["path"]).parent))

    non_title_tops = {
        name: n for name, n in top_level.items() if not _is_title(name)
    }

    lines: list[str] = []
    lines.append("RD gate analysis v2 — last SQLite snapshot, source=rd present=1")
    lines.append(f"files={file_totals['files']} unique_dirs={len(dir_seen)}")
    lines.append("")
    lines.append("== 1. Top-level folders under РД that are NOT a 4-digit title ==")
    for name, n in sorted(non_title_tops.items(), key=lambda kv: -kv[1]):
        lines.append(f"  {n:6d}  {name}")
    lines.append("")
    lines.append("== 2. Classification of first folder AFTER title + mark ==")
    for cls, counter in sorted(gate_by_class.items(), key=lambda kv: -sum(kv[1].values())):
        total = sum(counter.values())
        lines.append(f"\n-- {cls}: {total} files, {len(counter)} unique names --")
        for name, n in counter.most_common(40):
            lines.append(f"  {n:6d}  {name}")
    lines.append("")
    lines.append("== 3. Gate synonyms (передача / отправка in name) — KEEP these ==")
    for name, n in gate_names.most_common():
        lines.append(f"  {n:6d}  {name!r}")
    lines.append("")
    lines.append("== 4. Source-dump folders like 1055-SKUD (title-mark as folder after system) ==")
    dump_total = sum(gate_by_class["source_dump_like_1055_skud"].values())
    lines.append(f"  files={dump_total} unique_names={len(source_dump_examples)}")
    for name, n in gate_by_class["source_dump_like_1055_skud"].most_common():
        lines.append(f"  {n:6d}  {name}")
        for ex in source_dump_examples[name]:
            lines.append(f"          {ex}")
    lines.append("")
    lines.append("== 5. Unknown first-after-mark names (need human review) ==")
    for name, n in unknown_names.most_common(80):
        lines.append(f"  {n:6d}  {name}")
        for ex in unknown_examples[name][:2]:
            lines.append(f"          {ex}")
    lines.append("")
    lines.append("== 6. Example 8441 / SKUD paths (user sample) ==")
    lines.append(f"  files={len(files_8441_skud)}")
    # unique parents + first after mark
    parents_8441 = Counter()
    firsts_8441 = Counter()
    for p in files_8441_skud:
        rel = _after_rd(p) or []
        parents_8441[str(Path(p).parent)] += 1
        mark_i = None
        for i, part in enumerate(rel[1:], start=1):
            if _is_mark(part):
                mark_i = i
                break
        after = rel[mark_i + 1 : -1] if mark_i is not None else []
        firsts_8441[after[0] if after else "(in mark folder)"] += 1
    lines.append("  first-after-mark:")
    for name, n in firsts_8441.most_common():
        lines.append(f"    {n:6d}  {name}")
    lines.append("  sample parents:")
    for parent, n in parents_8441.most_common(25):
        lines.append(f"    {n:4d}  {parent}")

    lines.append("")
    lines.append("== 7. Title/mark trees without gate/direct-transfer (top 60 by files) ==")
    for (title, mark, first), n in no_gate_title_mark.most_common(60):
        lines.append(f"  {n:6d}  {title}/{mark}/{first}")

    OUT_TXT.write_text("\n".join(lines), encoding="utf-8")
    payload = {
        "non_title_top_level": non_title_tops,
        "gate_synonyms": dict(gate_names),
        "class_totals": {k: sum(v.values()) for k, v in gate_by_class.items()},
        "source_dumps": dict(gate_by_class["source_dump_like_1055_skud"]),
        "unknown": dict(unknown_names),
        "firsts_8441_skud": dict(firsts_8441),
    }
    OUT_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {OUT_TXT}")
    print("class totals", payload["class_totals"])
    print("gate synonym names", len(gate_names))
    print("source dumps", dump_total)
    print("8441 skud files", len(files_8441_skud))


if __name__ == "__main__":
    main()
