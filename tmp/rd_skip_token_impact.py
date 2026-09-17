"""Estimate RD files under skip-token path segments (read-only DB)."""

from __future__ import annotations

import os
import sqlite3
from collections import defaultdict
from pathlib import Path

DB_PATH = Path(
    os.path.expandvars(
        r"%LOCALAPPDATA%\Documentation_PDF_out_NK\rd_catalog\rd_catalog.sqlite"
    )
)
OUT = Path(__file__).with_name("rd_skip_token_impact.txt")

CURRENT = [
    "old",
    "temp",
    "tmp",
    "backup",
    "archive",
    "архив",
    "старые",
    "результат_проверки",
    "adapt_debug",
    "Таблички_графики",
    "пример",
    "для сравнения",
]
REQUESTED = [
    "ePlan",
    "МДЗ",
    "Интерфейс",
    "Таблица оснащения",
    "3D",
    "nanoCad",
]
PROPOSED = [
    "FROM",
    "Замечания",
    "WORK",
    "Объекты в закупку",
    "Наработки",
    "Прочее",
    "Оборудование",
    "Типовые схемы",
    "SENT",
    "Comment Attachments",
    "черновики",
    "сравнение",
    "ЗИП",
    "Цесис",
    "2xxx",
    "тест",
    "output",
    "загрузка СР",
    "Загрузка в СР",
    "Первая величина",
    "вспомогатель",
    "маркап",
    "картоп",
    "оld",
]

DANGEROUS_CHECK = [
    "раб",
    "ид",
    "as-build",
    "as-built",
    "SQ",
    "PDF",
    "DWG",
    "рев",
]


def matching_names(token: str, names: dict[str, int]) -> list[tuple[str, int]]:
    needle = token.casefold()
    hits = [(n, c) for n, c in names.items() if needle in n.casefold()]
    hits.sort(key=lambda x: -x[1])
    return hits


def main() -> None:
    conn = sqlite3.connect(str(DB_PATH))
    paths = [
        r[0]
        for r in conn.execute(
            "SELECT path FROM file_entry WHERE source = 'rd' AND present = 1"
        )
    ]
    total = len(paths)
    # unique dir names with file counts (file counted if that name is any ancestor)
    name_files: dict[str, int] = defaultdict(int)
    for path in paths:
        seen = set()
        for part in Path(path).parts:
            if part in {"\\", "/", "РД", "рд"}:
                continue
            if part not in seen:
                name_files[part] += 1
                seen.add(part)

    def report(title: str, tokens: list[str]) -> list[str]:
        lines = [f"== {title} =="]
        for token in tokens:
            hits = matching_names(token, name_files)
            files = 0
            # recount files actually hitting this token (union of paths)
            needle = token.casefold()
            n_files = sum(
                1
                for p in paths
                if any(needle in part.casefold() for part in Path(p).parts)
            )
            lines.append(
                f"  {n_files:6d} files  token={token!r}  unique_dir_names={len(hits)}"
            )
            for name, _c in hits[:12]:
                lines.append(f"           dir {name!r}")
            if len(hits) > 12:
                lines.append(f"           ... +{len(hits) - 12} names")
            files += n_files
        return lines

    lines = [f"RD present files: {total}", ""]
    lines += report("CURRENT skip_dirs (already in config)", CURRENT)
    lines.append("")
    lines += report("REQUESTED (add now)", REQUESTED)
    lines.append("")
    lines += report("PROPOSED extra", PROPOSED)
    lines.append("")
    lines += report("DANGEROUS / do not use as substring", DANGEROUS_CHECK)

    # union of requested
    req_needles = [t.casefold() for t in REQUESTED]
    req_files = sum(
        1
        for p in paths
        if any(
            any(n in part.casefold() for n in req_needles)
            for part in Path(p).parts
        )
    )
    extra_needles = [t.casefold() for t in PROPOSED]
    extra_files = sum(
        1
        for p in paths
        if any(
            any(n in part.casefold() for n in extra_needles)
            for part in Path(p).parts
        )
    )
    both = sum(
        1
        for p in paths
        if any(
            any(n in part.casefold() for n in req_needles + extra_needles)
            for part in Path(p).parts
        )
    )
    lines.append("")
    lines.append(f"union requested: {req_files}")
    lines.append(f"union proposed extra: {extra_files}")
    lines.append(f"union requested+proposed: {both} of {total}")

    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {OUT}")
    print(f"requested union {req_files}, extra {extra_files}, both {both}, total {total}")


if __name__ == "__main__":
    main()
