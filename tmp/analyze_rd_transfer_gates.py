"""Read-only analysis of RD catalog paths vs transfer-gate folders.

Looks at the latest local SQLite snapshot (not UNC). Classifies unique
directory trees after title + mark: known gates, synonym candidates,
and trees with no transfer-style parent.
"""

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
OUT_JSON = Path(__file__).with_name("analyze_rd_transfer_gates.json")
OUT_TXT = Path(__file__).with_name("analyze_rd_transfer_gates.txt")

RD_ROOT_MARKERS = ("\\рд\\", "/рд/")
KNOWN_GATES = (
    "для передачи",
    "на отправку",
)
# Leaf-ish folders that are file containers, not gates.
FILE_CONTAINERS = {
    "pdf",
    "dwg",
    "xlsx",
    "xls",
    "doc",
    "docx",
    "zip",
    "7z",
    "cad",
    "ifc",
    "nwd",
    "source",
    "исходники",
    "исходные",
}

TITLE_RE = re.compile(r"^\d{4}$")
MARK_HINT_RE = re.compile(
    r"(ksb|pos|spp|sos|skud|sot|soo|kbi|sagd)",
    re.IGNORECASE,
)


def _parts(path: str) -> list[str]:
    return [p for p in Path(path).parts if p not in ("\\", "/")]


def _relative_after_rd(path: str) -> list[str] | None:
    folded = path.replace("/", "\\").casefold()
    marker = "\\рд\\"
    idx = folded.rfind(marker)
    if idx < 0:
        return None
    # Keep original-cased remainder.
    original = path.replace("/", "\\")
    remainder = original[idx + 4 :]  # after \РД\ (len of \рд\ is 4 in fold; original may differ)
    # Safer: find the RD segment in original parts.
    parts = _parts(path)
    rd_i = None
    for i, part in enumerate(parts):
        if part.casefold() == "рд":
            rd_i = i
    if rd_i is None:
        return None
    return parts[rd_i + 1 :]


def _norm(name: str) -> str:
    return normalize_unicode_dashes(name).strip().casefold()


def _looks_like_title(name: str) -> bool:
    return bool(TITLE_RE.fullmatch(normalize_unicode_dashes(name).strip()))


def _looks_like_mark(name: str) -> bool:
    n = normalize_unicode_dashes(name).strip()
    if _looks_like_title(n):
        return False
    if is_transfer_folder_name(n):
        return False
    return bool(MARK_HINT_RE.search(n))


def _is_known_gate(name: str) -> bool:
    n = _norm(name)
    return any(g in n for g in KNOWN_GATES)


def _is_file_container(name: str) -> bool:
    n = _norm(name)
    return n in FILE_CONTAINERS


def main() -> None:
    if not DB_PATH.exists():
        raise SystemExit(f"DB not found: {DB_PATH}")

    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row

    runs = list(
        conn.execute(
            "SELECT id, started_at, completed_at, status, is_baseline "
            "FROM scan_run ORDER BY id DESC LIMIT 10"
        )
    )
    roots = list(conn.execute("SELECT source, root_path, last_successful_run_id FROM scan_root"))
    counts = list(
        conn.execute(
            "SELECT source, present, COUNT(*) AS n FROM file_entry GROUP BY source, present"
        )
    )

    rows = list(
        conn.execute(
            """
            SELECT path, file_kind, title, mark, transfer_name, transfer_sequence,
                   parse_status, parse_error
            FROM file_entry
            WHERE source = 'rd' AND present = 1
            """
        )
    )

    # Unique parent directories of RD files.
    dir_stats: dict[str, dict] = {}
    for row in rows:
        parent = str(Path(row["path"]).parent)
        rec = dir_stats.get(parent)
        if rec is None:
            rec = {
                "dir": parent,
                "files": 0,
                "kinds": Counter(),
                "titles": set(),
                "marks": set(),
                "transfer_names": Counter(),
                "parse_status": Counter(),
            }
            dir_stats[parent] = rec
        rec["files"] += 1
        rec["kinds"][row["file_kind"]] += 1
        if row["title"]:
            rec["titles"].add(row["title"])
        if row["mark"]:
            rec["marks"].add(row["mark"])
        if row["transfer_name"]:
            rec["transfer_names"][row["transfer_name"]] += 1
        rec["parse_status"][row["parse_status"]] += 1

    # Classify each unique directory.
    known_gate_dirs = []
    synonym_gate_dirs = []
    missing_gate_dirs = []
    gate_name_counter = Counter()
    segment3_counter = Counter()  # 3rd relative segment after RD (typically the gate)
    segment_after_mark_counter = Counter()
    unusual_structures = []

    for parent, rec in dir_stats.items():
        rel = _relative_after_rd(parent)
        rec["rel_parts"] = rel
        if not rel:
            unusual_structures.append(
                {"dir": parent, "reason": "cannot find RD segment", "files": rec["files"]}
            )
            continue

        # Typical: title / mark / GATE / transfer / (PDF|DWG|...)
        title_i = 0 if rel and _looks_like_title(rel[0]) else None
        mark_i = None
        if title_i is not None and len(rel) >= 2 and _looks_like_mark(rel[1]):
            mark_i = 1
        elif title_i is not None:
            # mark folder may not match Latin hint; take next non-title
            if len(rel) >= 2 and not is_transfer_folder_name(rel[1]):
                mark_i = 1

        after_mark = rel[mark_i + 1 :] if mark_i is not None else rel[1:]
        if after_mark:
            segment_after_mark_counter[after_mark[0]] += rec["files"]
        if len(rel) >= 3:
            segment3_counter[rel[2]] += rec["files"]

        gate_name = None
        gate_kind = "missing"
        # Search all segments after mark for known gates / transfer.
        search = after_mark if after_mark is not None else rel
        for part in search:
            if _is_known_gate(part):
                gate_name = part
                gate_kind = "known"
                break
        if gate_name is None:
            # synonym candidate: not transfer, not file container, not title,
            # and a later sibling/child is a numbered transfer folder.
            for i, part in enumerate(search):
                if is_transfer_folder_name(part):
                    # previous segment if any is the gate
                    if i > 0:
                        prev = search[i - 1]
                        if not _is_file_container(prev) and not _looks_like_title(prev):
                            gate_name = prev
                            gate_kind = "synonym_before_transfer"
                    else:
                        gate_kind = "transfer_without_gate"
                    break
            else:
                # no transfer folder in path at all
                if search:
                    first = search[0]
                    if not _is_file_container(first) and not is_transfer_folder_name(first):
                        gate_name = first
                        gate_kind = "no_transfer_folder"
                    else:
                        gate_kind = "no_transfer_folder"

        rec["gate_kind"] = gate_kind
        rec["gate_name"] = gate_name
        rec["title_folder"] = rel[title_i] if title_i is not None else None
        rec["mark_folder"] = rel[mark_i] if mark_i is not None else None
        rec["after_mark"] = after_mark

        if gate_name:
            gate_name_counter[gate_name] += rec["files"]

        payload = {
            "dir": parent,
            "files": rec["files"],
            "kinds": dict(rec["kinds"]),
            "title_folder": rec["title_folder"],
            "mark_folder": rec["mark_folder"],
            "gate_name": gate_name,
            "gate_kind": gate_kind,
            "after_mark": after_mark,
            "titles": sorted(rec["titles"]),
            "marks": sorted(rec["marks"]),
            "transfer_names": rec["transfer_names"].most_common(5),
        }
        if gate_kind == "known":
            known_gate_dirs.append(payload)
        elif gate_kind in {"synonym_before_transfer"}:
            synonym_gate_dirs.append(payload)
        else:
            missing_gate_dirs.append(payload)

    # Aggregate missing by (title, mark_folder, first after_mark)
    missing_groups: dict[tuple, dict] = {}
    for item in missing_gate_dirs:
        after = item["after_mark"] or []
        key = (
            item["title_folder"] or "?",
            item["mark_folder"] or "?",
            after[0] if after else "(files directly in mark folder)",
        )
        g = missing_groups.get(key)
        if g is None:
            g = {
                "title": key[0],
                "mark_folder": key[1],
                "first_after_mark": key[2],
                "gate_kind": item["gate_kind"],
                "dirs": 0,
                "files": 0,
                "kinds": Counter(),
                "examples": [],
            }
            missing_groups[key] = g
        g["dirs"] += 1
        g["files"] += item["files"]
        for k, v in item["kinds"].items():
            g["kinds"][k] += v
        if len(g["examples"]) < 3:
            g["examples"].append(item["dir"])

    synonym_groups: dict[str, dict] = {}
    for item in synonym_gate_dirs:
        name = item["gate_name"] or "?"
        g = synonym_groups.get(name)
        if g is None:
            g = {
                "gate_name": name,
                "dirs": 0,
                "files": 0,
                "titles": set(),
                "examples": [],
            }
            synonym_groups[name] = g
        g["dirs"] += 1
        g["files"] += item["files"]
        if item["title_folder"]:
            g["titles"].add(item["title_folder"])
        if len(g["examples"]) < 4:
            g["examples"].append(item["dir"])

    known_groups: dict[str, dict] = {}
    for item in known_gate_dirs:
        name = item["gate_name"] or "?"
        g = known_groups.get(name)
        if g is None:
            g = {"gate_name": name, "dirs": 0, "files": 0}
            known_groups[name] = g
        g["dirs"] += 1
        g["files"] += item["files"]

    # Unique first-after-mark names with file counts and whether transfer follows
    after_mark_names: dict[str, dict] = {}
    for item in (*known_gate_dirs, *synonym_gate_dirs, *missing_gate_dirs):
        after = item["after_mark"] or []
        if not after:
            name = "(empty — files in mark folder)"
        else:
            name = after[0]
        g = after_mark_names.get(name)
        if g is None:
            g = {
                "name": name,
                "norm": _norm(name) if after else name,
                "files": 0,
                "dirs": 0,
                "has_transfer_child": 0,
                "no_transfer": 0,
                "is_known_gate": _is_known_gate(name) if after else False,
                "is_transfer": is_transfer_folder_name(name) if after else False,
                "examples": [],
            }
            after_mark_names[name] = g
        g["files"] += item["files"]
        g["dirs"] += 1
        later = after[1:] if after else []
        if any(is_transfer_folder_name(p) for p in later) or (
            after and is_transfer_folder_name(after[0])
        ):
            g["has_transfer_child"] += 1
        else:
            g["no_transfer"] += 1
        if len(g["examples"]) < 2:
            g["examples"].append(item["dir"])

    def ser_counter(c: Counter) -> list[list]:
        return [[k, v] for k, v in c.most_common()]

    summary = {
        "db_path": str(DB_PATH),
        "db_mtime": DB_PATH.stat().st_mtime,
        "scan_runs": [dict(r) for r in runs],
        "roots": [dict(r) for r in roots],
        "file_counts": [dict(r) for r in counts],
        "rd_present_files": len(rows),
        "unique_rd_dirs": len(dir_stats),
        "known_gate_dirs": len(known_gate_dirs),
        "synonym_gate_dirs": len(synonym_gate_dirs),
        "missing_gate_dirs": len(missing_gate_dirs),
        "known_gate_names": {
            k: {"dirs": v["dirs"], "files": v["files"]} for k, v in known_groups.items()
        },
        "synonym_gate_names": {
            k: {
                "dirs": v["dirs"],
                "files": v["files"],
                "titles": sorted(v["titles"]),
                "examples": v["examples"],
            }
            for k, v in sorted(
                synonym_groups.items(), key=lambda kv: -kv[1]["files"]
            )
        },
        "missing_groups": [
            {
                **v,
                "kinds": dict(v["kinds"]),
            }
            for v in sorted(missing_groups.values(), key=lambda x: -x["files"])
        ],
        "first_after_mark_names": sorted(
            after_mark_names.values(), key=lambda x: -x["files"]
        ),
        "unusual_structures": unusual_structures[:50],
    }

    OUT_JSON.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )

    lines: list[str] = []
    lines.append("=== RD transfer-gate analysis (from last SQLite snapshot) ===")
    lines.append(f"DB: {DB_PATH}")
    lines.append(f"RD present files: {len(rows)}")
    lines.append(f"Unique parent dirs: {len(dir_stats)}")
    lines.append("")
    lines.append("-- scan_run --")
    for r in runs:
        lines.append(
            f"  id={r['id']} status={r['status']} baseline={r['is_baseline']} "
            f"started={r['started_at']} completed={r['completed_at']}"
        )
    lines.append("")
    lines.append("-- roots --")
    for r in roots:
        lines.append(f"  {r['source']}: last_ok={r['last_successful_run_id']} {r['root_path']}")
    lines.append("")
    lines.append("-- file counts (source, present, n) --")
    for r in counts:
        lines.append(f"  {r['source']} present={r['present']}: {r['n']}")
    lines.append("")
    lines.append("-- KNOWN gates (dirs / files) --")
    for name, g in sorted(known_groups.items(), key=lambda kv: -kv[1]["files"]):
        lines.append(f"  {name!r}: {g['dirs']} dirs, {g['files']} files")
    lines.append("")
    lines.append("-- SYNONYM gates (folder before numbered transfer, not known names) --")
    if not synonym_groups:
        lines.append("  (none)")
    for name, g in sorted(synonym_groups.items(), key=lambda kv: -kv[1]["files"]):
        lines.append(
            f"  {name!r}: {g['dirs']} dirs, {g['files']} files, titles={sorted(g['titles'])[:12]}"
        )
        for ex in g["examples"]:
            lines.append(f"      ex: {ex}")
    lines.append("")
    lines.append("-- MISSING / no transfer-gate (grouped by title / mark / first folder after mark) --")
    lines.append(f"  groups={len(missing_groups)} dirs={len(missing_gate_dirs)}")
    for g in sorted(missing_groups.values(), key=lambda x: -x["files"]):
        lines.append(
            f"  [{g['gate_kind']}] {g['title']}/{g['mark_folder']}/{g['first_after_mark']}: "
            f"{g['dirs']} dirs, {g['files']} files, kinds={dict(g['kinds'])}"
        )
        for ex in g["examples"]:
            lines.append(f"      ex: {ex}")
    lines.append("")
    lines.append("-- ALL unique first-after-mark folder names (by files) --")
    for g in sorted(after_mark_names.values(), key=lambda x: -x["files"]):
        lines.append(
            f"  {g['name']!r}: files={g['files']} dirs={g['dirs']} "
            f"known_gate={g['is_known_gate']} is_transfer={g['is_transfer']} "
            f"with_transfer={g['has_transfer_child']} no_transfer={g['no_transfer']}"
        )
        for ex in g["examples"][:1]:
            lines.append(f"      ex: {ex}")

    OUT_TXT.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {OUT_TXT}")
    print(f"Wrote {OUT_JSON}")
    print(f"known={len(known_gate_dirs)} synonym={len(synonym_gate_dirs)} missing={len(missing_gate_dirs)}")


if __name__ == "__main__":
    main()
