"""Read-only coverage: UL folders vs RFP part files by DS numbers.

Does not change pipeline code. Writes UTF-8 JSON + markdown under tmp/.
"""
from __future__ import annotations

import json
import os
import re
import sys
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from pathlib import Path

os.environ.setdefault("PYTHONUTF8", "1")

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from RFQ.ds_compare.tsd_packing_load import DEFAULT_TSD_PACKING_ROOT, collect_tsd_files
from RFQ.rfp_parts.ds_checklist import DEFAULT_PARTS_DIR, parse_ds_name_from_file_name

OUT_JSON = ROOT / "tmp" / "ul_rfp_ds_coverage.json"
OUT_MD = ROOT / "tmp" / "ul_rfp_ds_coverage.md"

_EXCEL = {".xlsx", ".xlsm", ".xls"}
_SKIP_DIR_PREFIXES = ("~",)
# DS token: ДС92, ДС92_24Б, ДС 92, дс47/13А, ДС47-13А
_DS_TOKEN_RE = re.compile(
    r"ДС\s*(\d+)(?:\s*[_\-./]\s*(\d+))?([А-ЯA-Zа-яa-z]*)",
    re.IGNORECASE,
)
_CYR_A = str.maketrans({"A": "А", "B": "В", "E": "Е", "K": "К", "M": "М", "H": "Н", "O": "О", "P": "Р", "C": "С", "T": "Т", "X": "Х"})


def _safe_print(msg: str) -> None:
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    sys.stdout.buffer.write((msg + "\n").encode(enc, errors="backslashreplace"))


def _norm_letter(raw: str) -> str:
    return (raw or "").upper().translate(_CYR_A)


def _fmt_ds(seq: int | None, orig: int | None, letter: str) -> str:
    if seq is None:
        return ""
    letter = _norm_letter(letter)
    if orig is not None and orig != seq:
        return f"ДС{seq}_{orig}{letter}"
    return f"ДС{seq}{letter}"


@dataclass(frozen=True)
class DsParse:
    """Two-number DS identity parsed from a folder or file name."""

    raw: str
    sequential: int | None
    original: int | None
    letter: str
    has_original_part: bool
    pattern: str

    @property
    def seq_key(self) -> str:
        return f"ДС{self.sequential}" if self.sequential is not None else ""

    @property
    def orig_key(self) -> str:
        if self.original is None:
            return ""
        return f"ДС{self.original}{self.letter}"

    @property
    def full_key(self) -> str:
        return _fmt_ds(self.sequential, self.original if self.has_original_part else None, self.letter)

    @property
    def kind(self) -> str:
        if self.sequential is None:
            return "unparsed"
        if self.has_original_part:
            return "compound"
        if self.letter:
            return "seq_letter"
        return "simple"


def parse_ds_tokens(text: str) -> list[DsParse]:
    """Extract every DS token from a folder/file name."""
    found: list[DsParse] = []
    seen: set[tuple] = set()
    for match in _DS_TOKEN_RE.finditer(text or ""):
        seq = int(match.group(1))
        orig_raw = match.group(2)
        letter = _norm_letter(match.group(3) or "")
        has_orig = orig_raw is not None
        orig = int(orig_raw) if has_orig else seq
        key = (seq, orig if has_orig else None, letter, has_orig)
        if key in seen:
            continue
        seen.add(key)
        if has_orig:
            pattern = "ДС{seq}_{orig}{letter}"
        elif letter:
            pattern = "ДС{seq}{letter}"
        else:
            pattern = "ДС{seq}"
        found.append(
            DsParse(
                raw=re.sub(r"\s+", "", match.group(0)).upper().translate(_CYR_A),
                sequential=seq,
                original=orig,
                letter=letter,
                has_original_part=has_orig,
                pattern=pattern,
            )
        )
    return found


def primary_parse(text: str) -> DsParse:
    tokens = parse_ds_tokens(text)
    if tokens:
        return tokens[0]
    legacy = parse_ds_name_from_file_name(text)
    return DsParse(
        raw=legacy,
        sequential=None,
        original=None,
        letter="",
        has_original_part=False,
        pattern="unparsed",
    )


@dataclass
class UlFolder:
    name: str
    rel: str
    parse: DsParse
    xlsx_count: int = 0
    files: list[str] = field(default_factory=list)


@dataclass
class RfpFile:
    name: str
    parse: DsParse
    legacy_key: str


def _is_skip_dir(name: str) -> bool:
    n = name.strip()
    return n.startswith("~") or n.startswith(".")


def collect_ul_folders(tsd_root: Path) -> tuple[list[UlFolder], list[dict], int, list[str]]:
    """First-level TSD folders + xlsx files grouped by parent folder."""
    folders: dict[str, UlFolder] = {}
    root_files: list[str] = []
    docs = collect_tsd_files(str(tsd_root))
    file_count = 0
    unparsed_files: list[str] = []

    # Also record empty first-level dirs (no xlsx).
    try:
        for child in tsd_root.iterdir():
            if child.is_dir() and not _is_skip_dir(child.name):
                p = primary_parse(child.name)
                folders[child.name] = UlFolder(
                    name=child.name,
                    rel=child.name,
                    parse=p,
                )
    except OSError as exc:
        _safe_print(f"UL iterdir error: {exc!r}")

    for doc in docs:
        file_count += 1
        path = Path(doc.file_full_path)
        try:
            rel = path.resolve().relative_to(tsd_root.resolve())
        except Exception:
            rel = Path(path.name)
        parts = rel.parts
        if len(parts) == 1:
            root_files.append(path.name)
            continue
        top = parts[0]
        if top not in folders:
            folders[top] = UlFolder(name=top, rel=top, parse=primary_parse(top))
        rec = folders[top]
        rec.xlsx_count += 1
        if len(rec.files) < 8:
            rec.files.append(path.name)

    file_rows: list[dict] = []
    for doc in docs:
        path = Path(doc.file_full_path)
        try:
            rel = str(path.resolve().relative_to(tsd_root.resolve()))
        except Exception:
            rel = path.name
        top = Path(rel).parts[0] if Path(rel).parts else path.name
        folder_parse = primary_parse(top if Path(rel).parent != Path(".") else "")
        file_parse = primary_parse(path.name)
        if folder_parse.sequential is None and file_parse.sequential is None:
            unparsed_files.append(rel)
        file_rows.append(
            {
                "rel": rel,
                "folder": top if Path(rel).parent != Path(".") else "(root)",
                "folder_parse": asdict(folder_parse),
                "file_parse": asdict(file_parse),
            }
        )
    return list(folders.values()), file_rows, file_count, root_files


def collect_rfp_files(parts_dir: Path) -> tuple[list[RfpFile], int]:
    files: list[RfpFile] = []
    total = 0
    if not parts_dir.exists():
        return files, 0
    for path in sorted(parts_dir.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file() or path.suffix.lower() not in _EXCEL:
            continue
        if path.name.startswith("~$"):
            continue
        total += 1
        files.append(
            RfpFile(
                name=path.name,
                parse=primary_parse(path.name),
                legacy_key=parse_ds_name_from_file_name(path.name),
            )
        )
    return files, total


def _sort_num(n: int | None) -> tuple[int, int]:
    return (0, n) if n is not None else (1, 10**9)


def build_index(ul_folders: list[UlFolder], rfp_files: list[RfpFile]) -> dict:
    ul_by_seq: dict[int, list[UlFolder]] = defaultdict(list)
    ul_by_orig: dict[int, list[UlFolder]] = defaultdict(list)
    rfp_by_seq: dict[int, list[RfpFile]] = defaultdict(list)
    rfp_by_orig: dict[int, list[RfpFile]] = defaultdict(list)

    for folder in ul_folders:
        if folder.parse.sequential is not None:
            ul_by_seq[folder.parse.sequential].append(folder)
        if folder.parse.original is not None:
            ul_by_orig[folder.parse.original].append(folder)
    for rec in rfp_files:
        if rec.parse.sequential is not None:
            rfp_by_seq[rec.parse.sequential].append(rec)
        if rec.parse.original is not None:
            rfp_by_orig[rec.parse.original].append(rec)

    all_seq = sorted(set(ul_by_seq) | set(rfp_by_seq))
    all_orig = sorted(set(ul_by_orig) | set(rfp_by_orig))

    def status_seq(n: int) -> str:
        has_ul = n in ul_by_seq
        has_rfp = n in rfp_by_seq
        if has_ul and has_rfp:
            return "both"
        if has_ul:
            return "ul_only"
        return "rfp_only"

    def status_orig(n: int) -> str:
        has_ul = n in ul_by_orig
        has_rfp = n in rfp_by_orig
        if has_ul and has_rfp:
            return "both"
        if has_ul:
            return "ul_only"
        return "rfp_only"

    seq_rows = []
    for n in all_seq:
        uls = ul_by_seq.get(n, [])
        rfps = rfp_by_seq.get(n, [])
        seq_rows.append(
            {
                "number": n,
                "status": status_seq(n),
                "ul_folders": [f.name for f in uls],
                "ul_xlsx": sum(f.xlsx_count for f in uls),
                "ul_kinds": sorted({f.parse.kind for f in uls}),
                "ul_full": sorted({f.parse.full_key for f in uls if f.parse.full_key}),
                "ul_orig": sorted({f.parse.original for f in uls if f.parse.original is not None}),
                "rfp_files": [f.name for f in rfps],
                "rfp_kinds": sorted({f.parse.kind for f in rfps}),
                "rfp_full": sorted({f.parse.full_key for f in rfps if f.parse.full_key}),
                "rfp_orig": sorted({f.parse.original for f in rfps if f.parse.original is not None}),
                "rfp_legacy": sorted({f.legacy_key for f in rfps if f.legacy_key}),
            }
        )

    orig_rows = []
    for n in all_orig:
        uls = ul_by_orig.get(n, [])
        rfps = rfp_by_orig.get(n, [])
        orig_rows.append(
            {
                "number": n,
                "status": status_orig(n),
                "ul_folders": [f.name for f in uls],
                "ul_xlsx": sum(f.xlsx_count for f in uls),
                "ul_seq": sorted({f.parse.sequential for f in uls if f.parse.sequential is not None}),
                "ul_full": sorted({f.parse.full_key for f in uls if f.parse.full_key}),
                "rfp_files": [f.name for f in rfps],
                "rfp_seq": sorted({f.parse.sequential for f in rfps if f.parse.sequential is not None}),
                "rfp_full": sorted({f.parse.full_key for f in rfps if f.parse.full_key}),
            }
        )

    # Pairwise: each unique (seq, orig, letter) identity
    pair_map: dict[tuple, dict] = {}

    def pair_key(p: DsParse) -> tuple:
        orig = p.original if p.has_original_part else p.sequential
        return (p.sequential, orig, p.letter or "")

    for folder in ul_folders:
        if folder.parse.sequential is None:
            continue
        key = pair_key(folder.parse)
        rec = pair_map.setdefault(
            key,
            {
                "sequential": key[0],
                "original": key[1],
                "letter": key[2],
                "full": folder.parse.full_key,
                "kind": folder.parse.kind,
                "ul_folders": [],
                "ul_xlsx": 0,
                "rfp_files": [],
            },
        )
        rec["ul_folders"].append(folder.name)
        rec["ul_xlsx"] += folder.xlsx_count
    for rec in rfp_files:
        if rec.parse.sequential is None:
            continue
        key = pair_key(rec.parse)
        item = pair_map.setdefault(
            key,
            {
                "sequential": key[0],
                "original": key[1],
                "letter": key[2],
                "full": rec.parse.full_key,
                "kind": rec.parse.kind,
                "ul_folders": [],
                "ul_xlsx": 0,
                "rfp_files": [],
            },
        )
        item["rfp_files"].append(rec.name)
        if not item["full"]:
            item["full"] = rec.parse.full_key

    pair_rows = []
    for item in pair_map.values():
        has_ul = bool(item["ul_folders"])
        has_rfp = bool(item["rfp_files"])
        if has_ul and has_rfp:
            item["status"] = "both"
        elif has_ul:
            item["status"] = "ul_only"
        else:
            item["status"] = "rfp_only"
        pair_rows.append(item)
    pair_rows.sort(key=lambda r: (_sort_num(r["sequential"]), _sort_num(r["original"]), r["letter"]))

    return {
        "seq_rows": seq_rows,
        "orig_rows": orig_rows,
        "pair_rows": pair_rows,
        "ul_by_seq": {str(k): [f.name for f in v] for k, v in ul_by_seq.items()},
        "rfp_by_seq": {str(k): [f.name for f in v] for k, v in rfp_by_seq.items()},
    }


def summarize(ul_folders, rfp_files, index, file_count, root_files, unparsed_ul_files):
    ul_parsed = [f for f in ul_folders if f.parse.sequential is not None]
    ul_unparsed = [f for f in ul_folders if f.parse.sequential is None]
    rfp_parsed = [f for f in rfp_files if f.parse.sequential is not None]
    rfp_unparsed = [f for f in rfp_files if f.parse.sequential is None]
    ul_compound = [f for f in ul_parsed if f.parse.has_original_part]
    rfp_compound = [f for f in rfp_parsed if f.parse.has_original_part]

    seq = index["seq_rows"]
    orig = index["orig_rows"]
    pairs = index["pair_rows"]

    def count_status(rows, status):
        return sum(1 for r in rows if r["status"] == status)

    # Cross-match: UL sequential equals some RFP original, or vice versa
    ul_seq = {f.parse.sequential for f in ul_parsed}
    ul_orig = {f.parse.original for f in ul_parsed}
    rfp_seq = {f.parse.sequential for f in rfp_parsed}
    rfp_orig = {f.parse.original for f in rfp_parsed}

    ul_seq_not_in_rfp_seq = sorted(ul_seq - rfp_seq)
    ul_seq_in_rfp_orig = sorted(set(ul_seq_not_in_rfp_seq) & rfp_orig)
    rfp_seq_not_in_ul_seq = sorted(rfp_seq - ul_seq)
    rfp_seq_in_ul_orig = sorted(set(rfp_seq_not_in_ul_seq) & ul_orig)
    rfp_orig_not_in_ul_orig = sorted(rfp_orig - ul_orig)
    ul_orig_not_in_rfp_orig = sorted(ul_orig - rfp_orig)

    return {
        "ul_folders_total": len(ul_folders),
        "ul_folders_parsed": len(ul_parsed),
        "ul_folders_unparsed": len(ul_unparsed),
        "ul_folders_compound": len(ul_compound),
        "ul_xlsx_files": file_count,
        "ul_root_xlsx": len(root_files),
        "ul_unparsed_files_sample": unparsed_ul_files[:20],
        "rfp_files_total": len(rfp_files),
        "rfp_files_parsed": len(rfp_parsed),
        "rfp_files_unparsed": len(rfp_unparsed),
        "rfp_files_compound": len(rfp_compound),
        "ul_kind_counts": dict(Counter(f.parse.kind for f in ul_folders)),
        "rfp_kind_counts": dict(Counter(f.parse.kind for f in rfp_files)),
        "seq_both": count_status(seq, "both"),
        "seq_ul_only": count_status(seq, "ul_only"),
        "seq_rfp_only": count_status(seq, "rfp_only"),
        "orig_both": count_status(orig, "both"),
        "orig_ul_only": count_status(orig, "ul_only"),
        "orig_rfp_only": count_status(orig, "rfp_only"),
        "pair_both": count_status(pairs, "both"),
        "pair_ul_only": count_status(pairs, "ul_only"),
        "pair_rfp_only": count_status(pairs, "rfp_only"),
        "ul_seq_not_in_rfp_seq": ul_seq_not_in_rfp_seq,
        "ul_seq_rescued_by_rfp_orig": ul_seq_in_rfp_orig,
        "rfp_seq_not_in_ul_seq": rfp_seq_not_in_ul_seq,
        "rfp_seq_rescued_by_ul_orig": rfp_seq_in_ul_orig,
        "ul_orig_not_in_rfp_orig": ul_orig_not_in_rfp_orig,
        "rfp_orig_not_in_ul_orig": rfp_orig_not_in_ul_orig,
        "need_parser": bool(ul_compound or rfp_compound),
        "seq_vs_orig_differs": bool(ul_compound or rfp_compound),
    }


def write_md(payload: dict) -> str:
    s = payload["summary"]
    lines = [
        "# Покрытие УЛ ↔ RFP по номерам ДС",
        "",
        f"- УЛ корень: `{payload['ul_root']}`",
        f"- RFP папка: `{payload['rfp_root']}`",
        f"- Папок УЛ: **{s['ul_folders_total']}** (разобрано {s['ul_folders_parsed']}, составных {s['ul_folders_compound']}, без ДС {s['ul_folders_unparsed']})",
        f"- Файлов УЛ xlsx: **{s['ul_xlsx_files']}**",
        f"- Файлов RFP: **{s['rfp_files_total']}** (разобрано {s['rfp_files_parsed']}, составных {s['rfp_files_compound']}, без ДС {s['rfp_files_unparsed']})",
        "",
        "## Сводка сопоставления",
        "",
        "| Ключ | Есть в обоих | Только УЛ | Только RFP |",
        "|---|---:|---:|---:|",
        f"| Порядковый номер (`ДС92` из `ДС92_24Б`) | {s['seq_both']} | {s['seq_ul_only']} | {s['seq_rfp_only']} |",
        f"| Фактический / исходный (`24` из `ДС92_24Б`) | {s['orig_both']} | {s['orig_ul_only']} | {s['orig_rfp_only']} |",
        f"| Полная пара seq+orig+буква | {s['pair_both']} | {s['pair_ul_only']} | {s['pair_rfp_only']} |",
        "",
        f"Нужен парсер двух номеров: **{'да' if s['need_parser'] else 'нет'}** "
        f"(составных имён УЛ={s['ul_folders_compound']}, RFP={s['rfp_files_compound']}).",
        "",
        "## Промахи по порядковому номеру",
        "",
        "### Только УЛ (нет RFP с тем же порядковым)",
        "",
    ]
    ul_only = [r for r in payload["index"]["seq_rows"] if r["status"] == "ul_only"]
    rfp_only = [r for r in payload["index"]["seq_rows"] if r["status"] == "rfp_only"]
    if not ul_only:
        lines.append("_нет_")
    else:
        lines += [
            "| Порядковый | Папки УЛ | xlsx | Фактический в УЛ |",
            "|---:|---|---:|---|",
        ]
        for r in ul_only:
            lines.append(
                f"| {r['number']} | {'; '.join(r['ul_folders'])} | {r['ul_xlsx']} | "
                f"{', '.join(str(x) for x in r['ul_orig'])} |"
            )
    lines += ["", "### Только RFP (нет папки УЛ с тем же порядковым)", ""]
    if not rfp_only:
        lines.append("_нет_")
    else:
        lines += [
            "| Порядковый | Файлы RFP | Фактический в RFP |",
            "|---:|---|---|",
        ]
        for r in rfp_only:
            names = "; ".join(r["rfp_files"][:3])
            extra = f" (+{len(r['rfp_files']) - 3})" if len(r["rfp_files"]) > 3 else ""
            lines.append(
                f"| {r['number']} | {names}{extra} | {', '.join(str(x) for x in r['rfp_orig'])} |"
            )

    lines += ["", "## Составные имена (`ДС{порядковый}_{исходный}{буква}`)", ""]
    compounds = [r for r in payload["index"]["pair_rows"] if r["kind"] == "compound"]
    if not compounds:
        lines.append("_составных имён нет_")
    else:
        lines += [
            "| Полный ключ | Порядковый | Исходный | Буква | Статус | УЛ | RFP |",
            "|---|---:|---:|---|---|---|---|",
        ]
        for r in compounds:
            lines.append(
                f"| {r['full']} | {r['sequential']} | {r['original']} | {r['letter'] or '—'} | "
                f"{r['status']} | {'; '.join(r['ul_folders']) or '—'} | "
                f"{'; '.join(r['rfp_files'][:2]) or '—'} |"
            )

    lines += ["", "## Неразобранные папки УЛ", ""]
    unparsed = payload["ul_unparsed_folders"]
    if not unparsed:
        lines.append("_нет_")
    else:
        for name in unparsed:
            lines.append(f"- `{name}`")

    lines += ["", "## Неразобранные файлы RFP", ""]
    unparsed_rfp = payload["rfp_unparsed_files"]
    if not unparsed_rfp:
        lines.append("_нет_")
    else:
        for name in unparsed_rfp:
            lines.append(f"- `{name}`")

    return "\n".join(lines) + "\n"


def main() -> int:
    tsd_root = Path(DEFAULT_TSD_PACKING_ROOT)
    parts_dir = Path(DEFAULT_PARTS_DIR)
    _safe_print(f"UL root exists={tsd_root.exists()} {ascii(str(tsd_root))}")
    _safe_print(f"RFP root exists={parts_dir.exists()} {ascii(str(parts_dir))}")

    if not tsd_root.exists() and not parts_dir.exists():
        _safe_print("STOP: neither UNC root is reachable")
        OUT_JSON.write_text(
            json.dumps({"error": "unc_unreachable", "ul": str(tsd_root), "rfp": str(parts_dir)}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return 1

    ul_folders, file_rows, file_count, root_files = collect_ul_folders(tsd_root)
    rfp_files, rfp_total = collect_rfp_files(parts_dir)
    _safe_print(f"UL folders={len(ul_folders)} xlsx={file_count} root_xlsx={len(root_files)}")
    _safe_print(f"RFP files={rfp_total} parsed={sum(1 for f in rfp_files if f.parse.sequential)}")

    index = build_index(ul_folders, rfp_files)
    unparsed_ul_files = [
        r["rel"]
        for r in file_rows
        if r["folder_parse"]["sequential"] is None and r["file_parse"]["sequential"] is None
    ]
    summary = summarize(ul_folders, rfp_files, index, file_count, root_files, unparsed_ul_files)

    payload = {
        "ul_root": str(tsd_root),
        "rfp_root": str(parts_dir),
        "summary": summary,
        "index": {
            "seq_rows": index["seq_rows"],
            "orig_rows": index["orig_rows"],
            "pair_rows": index["pair_rows"],
        },
        "ul_folders": [
            {
                "name": f.name,
                "xlsx_count": f.xlsx_count,
                **asdict(f.parse),
                "sample_files": f.files,
            }
            for f in sorted(ul_folders, key=lambda x: (_sort_num(x.parse.sequential), x.name.lower()))
        ],
        "rfp_files": [
            {
                "name": f.name,
                "legacy_key": f.legacy_key,
                **asdict(f.parse),
            }
            for f in rfp_files
        ],
        "ul_unparsed_folders": [f.name for f in ul_folders if f.parse.sequential is None],
        "rfp_unparsed_files": [f.name for f in rfp_files if f.parse.sequential is None],
        "ul_root_xlsx": root_files,
    }
    OUT_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    OUT_MD.write_text(write_md(payload), encoding="utf-8")
    _safe_print(f"wrote {OUT_JSON}")
    _safe_print(f"wrote {OUT_MD}")
    _safe_print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
