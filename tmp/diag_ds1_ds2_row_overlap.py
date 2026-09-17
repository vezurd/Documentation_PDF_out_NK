"""Read-only: check whether DS1/DS2 part rows overlap other DS files.

Uses the same extraction as RFQ.rfp_parts (Lot qty, tag parser, ds_name from
file name). Writes UTF-8 markdown to tmp/diag_ds1_ds2_row_overlap.md.

Run from repo root:
  python -X utf8 tmp/diag_ds1_ds2_row_overlap.py
"""

from __future__ import annotations

import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.rfp_parts.analyze_rfp_parts import (  # noqa: E402
    DEFAULT_PARTS_DIR,
    NO_TAG_KEY,
    PARTS_KIND,
    _extract_records,
    _norm_key_text,
    _record_tag_keys,
    _xlsx_files,
)
from RFQ.rfp_parts.ds_checklist import parse_ds_name_from_file_name  # noqa: E402

OUT_MD = ROOT / "tmp" / "diag_ds1_ds2_row_overlap.md"

TARGET_FILES = (
    "ДС1_Госфин_AGCC.287-0000-12.4.1-RFP-0004_0_RU_2.xlsx",
    "ДС2_AGCC.287-0000-12.4.1-RFP-0008_0_RU.XLSX",
)


def _safe_print(text: str) -> None:
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    sys.stdout.buffer.write((text + "\n").encode(enc, errors="backslashreplace"))


def _file_key(name: str) -> str:
    return name.casefold()


def _loc(record) -> str:
    return f"{record.file_name}·стр.{record.excel_row}"


def _tags_of(record) -> list[str]:
    silent: list[tuple[str, str, str]] = []
    return list(dict.fromkeys(_record_tag_keys(record, silent)))


def _md_cell(value: object) -> str:
    text = "" if value is None else str(value)
    return text.replace("|", "\\|").replace("\n", " ").replace("\r", " ")


def main() -> int:
    parts_dir = DEFAULT_PARTS_DIR
    _safe_print(f"parts_dir exists={parts_dir.exists()}: {parts_dir}")
    if not parts_dir.exists():
        _safe_print("UNC folder is not reachable from this environment.")
        return 2

    files = _xlsx_files(parts_dir)
    _safe_print(f"xlsx files: {len(files)}")
    by_name = {_file_key(path.name): path for path in files}

    missing = [name for name in TARGET_FILES if _file_key(name) not in by_name]
    if missing:
        _safe_print("missing targets:")
        for name in missing:
            _safe_print(f"  {name}")
        _safe_print("files starting with ДС1/ДС2:")
        for path in files:
            if path.name.upper().startswith("ДС1") or path.name.upper().startswith("ДС2"):
                _safe_print(f"  {path.name}")
        return 3

    target_keys = {_file_key(name) for name in TARGET_FILES}
    target_paths = [by_name[key] for key in target_keys]
    other_paths = [path for path in files if _file_key(path.name) not in target_keys]

    _safe_print("--- extract targets ---")
    warnings: list[tuple[str, str, str]] = []
    target_records = []
    for path in target_paths:
        _safe_print(f"  {path.name}")
        recs, stats = _extract_records(path, PARTS_KIND, warnings)
        _safe_print(f"    ds={stats.ds_name} rows={len(recs)} sheet={stats.sheet}")
        target_records.extend(recs)

    _safe_print(f"--- extract others ({len(other_paths)} files) ---")
    other_records = []
    for index, path in enumerate(other_paths, start=1):
        recs, stats = _extract_records(path, PARTS_KIND, warnings)
        other_records.extend(recs)
        if index % 10 == 0 or index == len(other_paths):
            _safe_print(f"  {index}/{len(other_paths)} last={path.name} rows={len(recs)}")

    # Indexes on other files
    other_by_tag: dict[str, list] = defaultdict(list)
    other_by_title_code_tag: dict[tuple[str, str, str], list] = defaultdict(list)
    other_by_code_tag: dict[tuple[str, str], list] = defaultdict(list)
    other_by_title_code: dict[tuple[str, str], list] = defaultdict(list)
    other_no_tag = 0
    for record in other_records:
        tags = _tags_of(record)
        title = _norm_key_text(record.ds_title)
        code = _norm_key_text(record.code)
        if not tags:
            other_no_tag += 1
            other_by_title_code[(title, code)].append(record)
            continue
        for tag in tags:
            tag_n = _norm_key_text(tag)
            other_by_tag[tag_n].append(record)
            other_by_code_tag[(code, tag_n)].append(record)
            other_by_title_code_tag[(title, code, tag_n)].append(record)

    # Resolve actual names from extracted records
    actual_names = sorted({r.file_name for r in target_records}, key=str.casefold)
    recs_by_actual: dict[str, list] = defaultdict(list)
    for record in target_records:
        recs_by_actual[record.file_name].append(record)

    lines: list[str] = []
    lines.append("# Пересечение строк ДС1 / ДС2 с остальными частями RFP_Зиновьев")
    lines.append("")
    lines.append(f"- Папка: `{parts_dir.as_posix()}`")
    lines.append(f"- Файлов xlsx в папке: **{len(files)}**")
    lines.append(
        "- Извлечение: тот же контур, что сбор частей "
        "(`Lot` qty, `get_tag`, `ds_name` из имени файла)."
    )
    lines.append(
        "- Ключ свода `ds_name+title+code+tag` **не** схлопывает разные ДС; "
        "здесь проверяем дубли **содержимого** (тег / title+code+tag / code+tag)."
    )
    lines.append("")
    lines.append("## Целевые файлы")
    lines.append("")
    lines.append("| Файл | ds_name | строк | с тегами | без тега |")
    lines.append("|---|---|---:|---:|---:|")

    target_hits_tag: list[dict] = []
    target_hits_tct: list[dict] = []
    target_hits_ct: list[dict] = []
    target_no_tag_title_code: list[dict] = []

    for file_name in actual_names:
        recs = recs_by_actual[file_name]
        with_tags = 0
        without = 0
        ds_name = recs[0].ds_name if recs else parse_ds_name_from_file_name(file_name)
        for record in recs:
            tags = _tags_of(record)
            title = _norm_key_text(record.ds_title)
            code = _norm_key_text(record.code)
            if not tags:
                without += 1
                others = other_by_title_code.get((title, code), [])
                if others:
                    target_no_tag_title_code.append(
                        {
                            "file": file_name,
                            "row": record.excel_row,
                            "ds": record.ds_name,
                            "title": record.ds_title,
                            "code": record.code,
                            "name": record.name,
                            "qty": record.values,
                            "units": record.units,
                            "others": others,
                        }
                    )
                continue
            with_tags += 1
            for tag in tags:
                tag_n = _norm_key_text(tag)
                hit_tag = other_by_tag.get(tag_n, [])
                hit_tct = other_by_title_code_tag.get((title, code, tag_n), [])
                hit_ct = other_by_code_tag.get((code, tag_n), [])
                payload = {
                    "file": file_name,
                    "row": record.excel_row,
                    "ds": record.ds_name,
                    "title": record.ds_title,
                    "code": record.code,
                    "name": record.name,
                    "tag": tag,
                    "tags_text": record.tags,
                    "qty": record.values,
                    "units": record.units,
                }
                if hit_tag:
                    target_hits_tag.append({**payload, "others": hit_tag})
                if hit_tct:
                    target_hits_tct.append({**payload, "others": hit_tct})
                if hit_ct:
                    target_hits_ct.append({**payload, "others": hit_ct})
        lines.append(
            f"| `{_md_cell(file_name)}` | {ds_name} | {len(recs)} | {with_tags} | {without} |"
        )

    lines.append("")
    lines.append("## Остальные файлы")
    lines.append("")
    lines.append(f"- Файлов (не ДС1/ДС2 целевые): **{len(other_paths)}**")
    lines.append(f"- Строк: **{len(other_records)}**")
    lines.append(f"- Строк без тега: **{other_no_tag}**")
    lines.append("")

    extract_errors = [w for w in warnings if w[0] == "ERROR"]
    if extract_errors:
        lines.append("### Ошибки чтения (не вошли в сравнение)")
        lines.append("")
        for level, fname, msg in extract_errors[:40]:
            lines.append(f"- `{_md_cell(fname)}`: {_md_cell(msg)}")
        if len(extract_errors) > 40:
            lines.append(f"- … ещё {len(extract_errors) - 40}")
        lines.append("")

    def _unique_files(hits: list[dict]) -> int:
        names: set[str] = set()
        for item in hits:
            for rec in item["others"]:
                names.add(rec.file_name)
        return len(names)

    lines.append("## Итог пересечений с остальными ДС")
    lines.append("")
    lines.append("| Правило | Совпадений (тег/строка целевого файла) | Чужих файлов |")
    lines.append("|---|---:|---:|")
    lines.append(
        f"| Одинаковый **тег** (как в `rfp_parts_duplicate_tags.xlsx`) | "
        f"{len(target_hits_tag)} | {_unique_files(target_hits_tag)} |"
    )
    lines.append(
        f"| Одинаковые **title + code + tag** (позиция без ds_name) | "
        f"{len(target_hits_tct)} | {_unique_files(target_hits_tct)} |"
    )
    lines.append(
        f"| Одинаковые **code + tag** | "
        f"{len(target_hits_ct)} | {_unique_files(target_hits_ct)} |"
    )
    lines.append(
        f"| Строка **без тега**, совпал title+code | "
        f"{len(target_no_tag_title_code)} | {_unique_files(target_no_tag_title_code)} |"
    )
    lines.append("")

    if not target_hits_tag and not target_hits_tct and not target_hits_ct:
        lines.append(
            "**Вывод:** строки целевых ДС1/ДС2 **не дублируют** позиции остальных "
            "файлов папки ни по тегу, ни по title+code+tag, ни по code+tag."
        )
        lines.append("")
    else:
        lines.append(
            "**Вывод:** есть пересечение с другими ДС. Ниже — совпадения "
            "`title+code+tag` (самый жёсткий ключ содержимого); если пусто, "
            "смотрите совпадения только по тегу."
        )
        lines.append("")

    def _append_hits(title: str, hits: list[dict], limit: int = 80) -> None:
        lines.append(f"## {title}")
        lines.append("")
        if not hits:
            lines.append("Нет совпадений.")
            lines.append("")
            return
        lines.append(
            "| Целевой файл | стр. | ds | тег | код | qty | Чужие вхождения |"
        )
        lines.append("|---|---:|---|---|---|---:|---|")
        shown = hits[:limit]
        for item in shown:
            others = ", ".join(_loc(rec) for rec in item["others"][:6])
            extra = (
                f" и ещё {len(item['others']) - 6}"
                if len(item["others"]) > 6
                else ""
            )
            lines.append(
                "| "
                + " | ".join(
                    [
                        f"`{_md_cell(item['file'])}`",
                        str(item["row"]),
                        _md_cell(item["ds"]),
                        f"`{_md_cell(item['tag'])}`",
                        f"`{_md_cell(item['code'])}`",
                        _md_cell(item["qty"]),
                        _md_cell(others + extra),
                    ]
                )
                + " |"
            )
        if len(hits) > limit:
            lines.append("")
            lines.append(f"Показаны первые {limit} из {len(hits)}.")
        lines.append("")

        # Summary: which other files collide most
        file_counts: dict[str, int] = defaultdict(int)
        for item in hits:
            seen = {rec.file_name for rec in item["others"]}
            for name in seen:
                file_counts[name] += 1
        lines.append("### Чужие файлы с наибольшим числом совпадений")
        lines.append("")
        lines.append("| Файл | Совпадений |")
        lines.append("|---|---:|")
        for name, count in sorted(file_counts.items(), key=lambda kv: (-kv[1], kv[0].casefold()))[
            :30
        ]:
            lines.append(f"| `{_md_cell(name)}` | {count} |")
        lines.append("")

    _append_hits("Совпадения title + code + tag", target_hits_tct)
    _append_hits("Совпадения только по тегу (код/титул могут отличаться)", target_hits_tag)

    lines.append("## Детали совпадений по тегу (титул/код/имя/qty)")
    lines.append("")
    if not target_hits_tag:
        lines.append("Нет совпадений.")
        lines.append("")
    else:
        lines.append(
            "| Тег | Цель: файл·стр / титул / код / qty | Чужое: файл·стр / титул / код / qty |"
        )
        lines.append("|---|---|---|")
        for item in target_hits_tag:
            target_side = (
                f"{item['file']}·стр.{item['row']} / "
                f"{item['title']} / {item['code']} / {item['qty']}"
            )
            other_bits = []
            for rec in item["others"][:4]:
                other_bits.append(
                    f"{rec.file_name}·стр.{rec.excel_row} / "
                    f"{rec.ds_title} / {rec.code} / {rec.values}"
                )
            extra = (
                f" (+{len(item['others']) - 4})"
                if len(item["others"]) > 4
                else ""
            )
            lines.append(
                f"| `{_md_cell(item['tag'])}` | {_md_cell(target_side)} | "
                f"{_md_cell('; '.join(other_bits) + extra)} |"
            )
        lines.append("")

    # Untagged rows: title+code is a weak key (bulk MTR). Split stronger matches.
    lines.append("## Строки без тега: насколько это те же позиции")
    lines.append("")
    if not target_no_tag_title_code:
        lines.append("Нет совпадений title+code у строк без тега.")
        lines.append("")
    else:
        same_qty = 0
        same_name_qty = 0
        same_full = 0
        file_counts: dict[str, int] = defaultdict(int)
        file_full: dict[str, int] = defaultdict(int)
        samples_full: list[dict] = []
        for item in target_no_tag_title_code:
            name_n = _norm_key_text(item["name"])
            units_n = _norm_key_text(item["units"])
            qty = item["qty"]
            others = item["others"]
            seen_files = {rec.file_name for rec in others}
            for name in seen_files:
                file_counts[name] += 1
            qty_hit = any(rec.values == qty for rec in others)
            name_qty_hit = any(
                rec.values == qty and _norm_key_text(rec.name) == name_n
                for rec in others
            )
            full_hit_recs = [
                rec
                for rec in others
                if rec.values == qty
                and _norm_key_text(rec.name) == name_n
                and _norm_key_text(rec.units) == units_n
            ]
            if qty_hit:
                same_qty += 1
            if name_qty_hit:
                same_name_qty += 1
            if full_hit_recs:
                same_full += 1
                for name in {rec.file_name for rec in full_hit_recs}:
                    file_full[name] += 1
                if len(samples_full) < 25:
                    samples_full.append({**item, "full_others": full_hit_recs})
        no_tag_target_rows = sum(
            1
            for recs in recs_by_actual.values()
            for rec in recs
            if not _tags_of(rec)
        )
        lines.append(
            f"- Строк без тега в ДС1/ДС2: **{no_tag_target_rows}**"
        )
        lines.append(
            f"- Из них совпал **title+code** с чужим файлом: **{len(target_no_tag_title_code)}** "
            "(слабый ключ — одинаковый код МТР в разных ДС обычен)"
        )
        lines.append(
            f"- Из них ещё совпало **qty**: **{same_qty}**"
        )
        lines.append(
            f"- Из них **name + qty**: **{same_name_qty}**"
        )
        lines.append(
            f"- Из них **name + qty + ед.изм.** (почти полная копия строки): **{same_full}**"
        )
        lines.append("")
        lines.append("### Чужие файлы по title+code (без тега)")
        lines.append("")
        lines.append("| Файл | Совпадений title+code | в т.ч. name+qty+ед. |")
        lines.append("|---|---:|---:|")
        for name, count in sorted(
            file_counts.items(), key=lambda kv: (-kv[1], kv[0].casefold())
        ):
            lines.append(
                f"| `{_md_cell(name)}` | {count} | {file_full.get(name, 0)} |"
            )
        lines.append("")
        if samples_full:
            lines.append("### Примеры полных копий (без тега, name+qty+ед.)")
            lines.append("")
            lines.append("| Цель | стр. | код | qty | имя | Чужие |")
            lines.append("|---|---:|---|---:|---|---|")
            for item in samples_full:
                others = ", ".join(_loc(rec) for rec in item["full_others"][:4])
                extra = (
                    f" и ещё {len(item['full_others']) - 4}"
                    if len(item["full_others"]) > 4
                    else ""
                )
                lines.append(
                    "| "
                    + " | ".join(
                        [
                            f"`{_md_cell(item['file'])}`",
                            str(item["row"]),
                            f"`{_md_cell(item['code'])}`",
                            _md_cell(item["qty"]),
                            _md_cell(item["name"])[:80],
                            _md_cell(others + extra),
                        ]
                    )
                    + " |"
                )
            lines.append("")

    # DS1 vs DS2 between themselves
    lines.append("## Пересечение ДС1 ↔ ДС2 между собой")
    lines.append("")
    if len(actual_names) < 2:
        lines.append("Оба целевых файла не прочитаны — сравнение между ними пропущено.")
        lines.append("")
    else:
        a_name, b_name = actual_names[0], actual_names[1]
        a_recs = recs_by_actual[a_name]
        b_recs = recs_by_actual[b_name]
        b_by_tag: dict[str, list] = defaultdict(list)
        b_by_tct: dict[tuple[str, str, str], list] = defaultdict(list)
        for record in b_recs:
            title = _norm_key_text(record.ds_title)
            code = _norm_key_text(record.code)
            for tag in _tags_of(record):
                tag_n = _norm_key_text(tag)
                b_by_tag[tag_n].append(record)
                b_by_tct[(title, code, tag_n)].append(record)
        between_tag = 0
        between_tct = 0
        samples: list[str] = []
        for record in a_recs:
            title = _norm_key_text(record.ds_title)
            code = _norm_key_text(record.code)
            for tag in _tags_of(record):
                tag_n = _norm_key_text(tag)
                if b_by_tct.get((title, code, tag_n)):
                    between_tct += 1
                    if len(samples) < 15:
                        samples.append(
                            f"- `{tag}` / `{record.code}` — "
                            f"{_loc(record)} ↔ "
                            + ", ".join(_loc(x) for x in b_by_tct[(title, code, tag_n)][:3])
                        )
                elif b_by_tag.get(tag_n):
                    between_tag += 1
                    if len(samples) < 15:
                        samples.append(
                            f"- тег `{tag}` (разный код/титул) — "
                            f"{_loc(record)} ↔ "
                            + ", ".join(_loc(x) for x in b_by_tag[tag_n][:3])
                        )
        lines.append(f"- `{_md_cell(a_name)}`: {len(a_recs)} строк")
        lines.append(f"- `{_md_cell(b_name)}`: {len(b_recs)} строк")
        lines.append(f"- Совпадений title+code+tag: **{between_tct}**")
        lines.append(f"- Совпадений только тег: **{between_tag}**")
        lines.append("")
        if samples:
            lines.append("Примеры:")
            lines.append("")
            lines.extend(samples)
            lines.append("")
        elif between_tct == 0 and between_tag == 0:
            lines.append("Между ДС1 и ДС2 пересечений по тегам нет.")
            lines.append("")

    def _fp(record) -> tuple:
        tags = _tags_of(record)
        tag_part = tuple(sorted(_norm_key_text(t) for t in tags)) or (NO_TAG_KEY,)
        return (
            _norm_key_text(record.ds_title),
            _norm_key_text(record.code),
            tag_part,
            _norm_key_text(record.name),
            record.values,
            _norm_key_text(record.units),
        )

    def _pair_cover(left: list, right: list) -> tuple[int, int, int]:
        right_fps: dict[tuple, int] = defaultdict(int)
        for rec in right:
            right_fps[_fp(rec)] += 1
        covered = 0
        for rec in left:
            fp = _fp(rec)
            if right_fps.get(fp, 0) > 0:
                covered += 1
                right_fps[fp] -= 1
        left_tagged = sum(1 for rec in left if _tags_of(rec))
        right_tagged = sum(1 for rec in right if _tags_of(rec))
        return covered, left_tagged, right_tagged

    gf_names = sorted(
        {rec.file_name for rec in other_records if "ГФ" in rec.file_name.upper()},
        key=str.casefold,
    )
    other_by_file: dict[str, list] = defaultdict(list)
    for rec in other_records:
        other_by_file[rec.file_name].append(rec)

    lines.append("## Попарно с файлами «ГФ» в той же папке")
    lines.append("")
    lines.append(
        "Полный отпечаток строки: title + code + набор тегов + имя + qty + ед. изм."
    )
    lines.append("")
    if not gf_names:
        lines.append("Других файлов с «ГФ» в имени нет.")
        lines.append("")
    else:
        lines.append("| Цель | ГФ-файл | строк цели | строк ГФ | полных копий из цели | тегов в цели | тегов в ГФ |")
        lines.append("|---|---|---:|---:|---:|---:|---:|")
        for target_name in actual_names:
            left = recs_by_actual[target_name]
            for gf_name in gf_names:
                right = other_by_file.get(gf_name, [])
                covered, left_tagged, right_tagged = _pair_cover(left, right)
                lines.append(
                    "| "
                    + " | ".join(
                        [
                            f"`{_md_cell(target_name)}`",
                            f"`{_md_cell(gf_name)}`",
                            str(len(left)),
                            str(len(right)),
                            str(covered),
                            str(left_tagged),
                            str(right_tagged),
                        ]
                    )
                    + " |"
                )
        lines.append("")
        # How many GF rows are NOT in the matching target
        lines.append("Обратное покрытие (строки ГФ, которых нет в цели, тот же отпечаток):")
        lines.append("")
        for target_name in actual_names:
            left = recs_by_actual[target_name]
            for gf_name in gf_names:
                right = other_by_file.get(gf_name, [])
                covered_rev, _, _ = _pair_cover(right, left)
                missing = len(right) - covered_rev
                lines.append(
                    f"- `{_md_cell(gf_name)}` → `{_md_cell(target_name)}`: "
                    f"в ГФ {len(right)} строк, из них нет в цели **{missing}**"
                )
        lines.append("")

    # Other DS1/DS2 files in the folder (not the two targets)
    other_ds12 = [
        path.name
        for path in other_paths
        if parse_ds_name_from_file_name(path.name).upper() in {"ДС1", "ДС2"}
    ]
    if other_ds12:
        lines.append("## Другие файлы ДС1/ДС2 в папке (не целевые)")
        lines.append("")
        for name in other_ds12:
            lines.append(f"- `{_md_cell(name)}`")
        lines.append("")

    OUT_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")
    _safe_print(f"wrote {OUT_MD}")
    _safe_print(
        f"hits tag={len(target_hits_tag)} tct={len(target_hits_tct)} "
        f"ct={len(target_hits_ct)} no_tag={len(target_no_tag_title_code)}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
