"""List title-mark trees with NN_рев structure but no transfer-gate folder."""

from __future__ import annotations

import os
import re
import sqlite3
from collections import defaultdict
from pathlib import Path

from rd_catalog.parse import (
    _TITLE_ONLY_FOLDER_RE,
    _TRANSFER_SEQUENCE_RE,
    is_transfer_folder_name,
    normalize_unicode_dashes,
)

DB_PATH = Path(
    os.path.expandvars(
        r"%LOCALAPPDATA%\Documentation_PDF_out_NK\rd_catalog\rd_catalog.sqlite"
    )
)
OUT_TXT = Path(__file__).with_name("rd_missing_gate_title_marks.txt")
OUT_TSV = Path(__file__).with_name("rd_missing_gate_title_marks.tsv")

GATE_RE = re.compile(r"передач|отправк", re.IGNORECASE)
MARK_HINT_RE = re.compile(
    r"(ksb|pos|spp|sos|skud|sot|soo|kbi|sagd)",
    re.IGNORECASE,
)
NOISE_PARENT_RE = re.compile(
    r"(замечан|from|work|работ|исходн|интерфейс|мдз|спек|"
    r"закупк|прочее|типов|оборуд|вспомогат|наработ|"
    r"корректир|загрузк|sent|comment|черновик|тест|old|оld)",
    re.IGNORECASE,
)
TITLE_RE = re.compile(r"^\d{4}$")
REV_TOKEN_RE = re.compile(r"рев|rev", re.IGNORECASE)


def _parts(path: str) -> list[str]:
    return [p for p in Path(path).parts if p not in ("\\", "/")]


def _after_rd(path: str) -> list[str] | None:
    parts = _parts(path)
    for i, part in enumerate(parts):
        if part.casefold() == "рд":
            return parts[i + 1 :]
    return None


def _looks_like_transfer(name: str) -> bool:
    if is_transfer_folder_name(name):
        return True
    normalized = normalize_unicode_dashes(name).strip()
    if not normalized or _TITLE_ONLY_FOLDER_RE.fullmatch(normalized):
        return False
    if _TRANSFER_SEQUENCE_RE.match(normalized) is None:
        return False
    return bool(REV_TOKEN_RE.search(normalized))


def _is_title(name: str) -> bool:
    return bool(TITLE_RE.fullmatch(normalize_unicode_dashes(name).strip()))


def _is_mark_folder(name: str) -> bool:
    n = normalize_unicode_dashes(name).strip()
    if _is_title(n) or _looks_like_transfer(n):
        return False
    return bool(MARK_HINT_RE.search(n))


def main() -> None:
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    rows = list(
        conn.execute(
            """
            SELECT path, title, mark, file_kind
            FROM file_entry
            WHERE source = 'rd' AND present = 1
            """
        )
    )

    # parent_dir -> stats
    groups: dict[str, dict] = {}

    for row in rows:
        rel = _after_rd(row["path"])
        if not rel or len(rel) < 2:
            continue
        if not _is_title(rel[0]):
            continue
        if any(GATE_RE.search(part) for part in rel[:-1]):
            continue

        transfer_i = None
        for i, part in enumerate(rel[:-1]):
            if _looks_like_transfer(part):
                transfer_i = i
                break
        if transfer_i is None or transfer_i == 0:
            continue

        parent_rel = rel[:transfer_i]
        parent_parts = _parts(row["path"])
        # Rebuild parent path: everything except filename and transfer subtree
        # path parts include UNC prefix + ... + RD + rel
        drop = len(rel) - transfer_i  # filename + after transfer start
        parent_path = str(Path(*parent_parts[:-drop])) if drop < len(parent_parts) else ""
        if not parent_path:
            continue

        parent_name = parent_rel[-1]
        if NOISE_PARENT_RE.search(parent_name):
            continue
        # Extra noise layers between title and transfer (FROM/WORK etc.)
        if any(NOISE_PARENT_RE.search(p) for p in parent_rel[1:]):
            continue

        rec = groups.get(parent_path)
        if rec is None:
            rec = {
                "parent": parent_path,
                "title_folder": rel[0],
                "parent_rel": "\\".join(parent_rel),
                "titles": set(),
                "marks": set(),
                "mark_folders": set(),
                "transfers": set(),
                "files": 0,
                "kinds": defaultdict(int),
            }
            groups[parent_path] = rec
        rec["files"] += 1
        rec["kinds"][row["file_kind"]] += 1
        rec["transfers"].add(rel[transfer_i])
        if row["title"]:
            rec["titles"].add(row["title"])
        if row["mark"]:
            rec["marks"].add(row["mark"])
        for part in parent_rel[1:]:
            if _is_mark_folder(part):
                rec["mark_folders"].add(part)

    items = sorted(
        groups.values(),
        key=lambda r: (r["title_folder"], "\\".join(sorted(r["marks"])), r["parent"]),
    )

    lines: list[str] = []
    lines.append(
        "Титул–марки без папки «Для передачи» / «На отправку», "
        "но сразу под системой есть NN_рев.* (в т.ч. рев.A/B)."
    )
    lines.append(
        "Путь — папка, куда нужно добавить шлюз и перенести дочерние NN_рев."
    )
    lines.append(f"Записей: {len(items)}")
    lines.append("")

    tsv = [
        "title_folder\ttitle_from_files\tmark_from_files\tmark_folders\t"
        "files\tpdf\tmto\teditable\ttransfer_folders\tparent_path"
    ]

    for rec in items:
        titles = ", ".join(sorted(rec["titles"])) or "—"
        marks = ", ".join(sorted(rec["marks"])) or "—"
        mark_folders = ", ".join(sorted(rec["mark_folders"])) or "—"
        transfers = sorted(rec["transfers"])
        kinds = rec["kinds"]
        lines.append(
            f"{rec['title_folder']}-{marks}  files={rec['files']}  "
            f"pdf={kinds.get('pdf', 0)} mto={kinds.get('mto_xlsx', 0)} "
            f"edit={kinds.get('source_editable', 0)}"
        )
        lines.append(f"  папка марки: {mark_folders}  (title в файлах: {titles})")
        lines.append(f"  путь: {rec['parent']}")
        for name in transfers:
            lines.append(f"    → {name}")
        lines.append("")
        tsv.append(
            "\t".join(
                [
                    rec["title_folder"],
                    titles,
                    marks,
                    mark_folders,
                    str(rec["files"]),
                    str(kinds.get("pdf", 0)),
                    str(kinds.get("mto_xlsx", 0)),
                    str(kinds.get("source_editable", 0)),
                    " | ".join(transfers),
                    rec["parent"],
                ]
            )
        )

    OUT_TXT.write_text("\n".join(lines), encoding="utf-8")
    OUT_TSV.write_text("\n".join(tsv), encoding="utf-8")
    print(f"groups={len(items)}")
    print(f"Wrote {OUT_TXT}")
    print(f"Wrote {OUT_TSV}")
    print("titles:", sorted({r['title_folder'] for r in items}))


if __name__ == "__main__":
    main()
