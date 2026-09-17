"""Examples: correlate issuance history + F events + current RD rev."""

from __future__ import annotations

import json
import os
import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from rd_catalog.config import load_config
from rd_catalog.db import CatalogDatabase
from rd_catalog.google_kits import issuance_cache_paths, kits_cache_paths
from rd_catalog.kits import (
    format_revision,
    kit_identity_key,
    parse_google_matrix,
    parse_history_comment,
    parse_issuance_kit_row,
    parse_issuance_matrix,
)
from rd_catalog.models import SourceKind


def _rows(path: Path) -> list[list[str]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    return payload["rows"]


def main() -> None:
    cfg = load_config()
    out: list[str] = []
    p = out.append

    kits_rows = _rows(kits_cache_paths(cfg.runtime_dir)[0])
    iss_rows = _rows(issuance_cache_paths(cfg.runtime_dir)[0])
    kits, _ = parse_google_matrix(kits_rows)

    all_iss: dict[tuple[str, str], list] = defaultdict(list)
    start = 1
    for offset, row in enumerate(iss_rows[start:], start=start + 1):
        parsed = parse_issuance_kit_row(row, row_index=offset)
        if isinstance(parsed, str):
            continue
        all_iss[kit_identity_key(parsed.title, parsed.mark)].append(parsed)

    n_sends = [len(v) for v in all_iss.values()]
    p(f"issuance kits with >=1 parsed send: {len(all_iss)}")
    p(
        "sends per title-mark: "
        f"min={min(n_sends)} median={sorted(n_sends)[len(n_sends)//2]} "
        f"max={max(n_sends)} mean={sum(n_sends)/len(n_sends):.1f}"
    )
    same_rev_multi = 0
    for items in all_iss.values():
        by_rev = Counter(i.revision_text or "—" for i in items)
        if any(n >= 2 for n in by_rev.values()):
            same_rev_multi += 1
    p(f"kits with >=2 issuance rows of the SAME revision: {same_rev_multi}")

    code_c = 0
    code_a_ev = code_b_ev = 0
    for kit in kits:
        for ev in kit.events:
            raw = (ev.stage_label + " " + ev.raw).casefold()
            if ev.stage == "code_a":
                code_a_ev += 1
            if ev.stage == "code_b":
                code_b_ev += 1
            if "код" in raw and any(tok in raw for tok in ("код c", "код с", "код c ")):
                code_c += 1
            if ev.stage == "other" and ("код с" in raw or "код c" in raw):
                code_c += 1
    p(f"F events classified code_a={code_a_ev} code_b={code_b_ev}")
    p(f"F events looking like code C (incl other): check below")

    c_samples = []
    for kit in kits:
        for ev in kit.events:
            low = ev.raw.casefold()
            if "код" in low and ("код c" in low or "код с" in low or "код c" in low.replace("c", "c")):
                if "код а" in low or "код a" in low:
                    continue
                if ev.stage == "code_b":
                    continue
                c_samples.append((kit.title_system, ev.date, ev.stage, ev.raw[:120]))
    p(f"code-C-like samples: {len(c_samples)}")
    for s in c_samples[:12]:
        p(f"  {s}")

    db = CatalogDatabase(cfg.db_path)
    overlay_ids = {int(r["file_id"]) for r in db.current_overlay()}
    from rd_catalog.kits import build_kit_matrix

    matrix = build_kit_matrix(kits, db.list_files(), overlay_ids, issuance_kits=parse_issuance_matrix(iss_rows)[0])
    by_key = {kit_identity_key(r.title, r.mark): r for r in matrix}

    def dump_kit(title: str, mark: str) -> None:
        p("")
        p(f"===== {title}-{mark} =====")
        row = by_key.get(kit_identity_key(title, mark))
        if row is None:
            p("  not in matrix")
            return
        p(f"  summary={row.summary.value} rd={row.rd.revision_text or '—'} robot={row.robot.revision_text or '—'}")
        if row.google:
            p(f"  E={row.google.status_sheet!r} sheet_rev={row.google.sheet_revision_text}")
            for ev in row.google.events:
                rev = format_revision(ev.revision, ev.appendix) or "—"
                p(f"    F {ev.date or 'no-date':12} {ev.stage:16} {ev.stage_label:28} rev={rev:8} {ev.transmittals}")
        sends = all_iss.get(kit_identity_key(title, mark), [])
        p(f"  issuance rows (ALL, not latest-only): {len(sends)}")
        for item in sends:
            p(
                f"    {item.send_date_sortable or item.send_date:12} "
                f"rev={item.revision_text or '—':8} {item.status:28} "
                f"ctrl={item.incoming_control_date_sortable or '—'} "
                f"send={item.send_transmittal or '—'} conf={item.confirm_transmittal or '—'}"
            )

    dump_kit("9110", "KSB1")

    # last F is code_a AND google/rd rev match that event
    p("")
    p("=== kits: last dated F is code_a ===")
    n = 0
    for kit in kits:
        dated = [e for e in kit.events if e.date]
        if not dated:
            continue
        last = dated[-1]
        if last.stage != "code_a":
            continue
        n += 1
        if n <= 5:
            rev = format_revision(last.revision, last.appendix) or "—"
            row = by_key.get(kit_identity_key(kit.title, kit.mark))
            rd = row.rd.revision_text if row else "?"
            p(f"  {kit.title_system} E={kit.status_sheet!r} lastA date={last.date} ev_rev={rev} sheet={kit.sheet_revision_text} rd={rd}")
    p(f"  total last-F code_a: {n}")

    p("")
    p("=== kits: last dated F is tdo_passed ===")
    n = 0
    for kit in kits:
        dated = [e for e in kit.events if e.date]
        if not dated or dated[-1].stage != "tdo_passed":
            continue
        n += 1
        if n <= 5:
            last = dated[-1]
            rev = format_revision(last.revision, last.appendix) or "—"
            row = by_key.get(kit_identity_key(kit.title, kit.mark))
            rd = row.rd.revision_text if row else "?"
            iss = row.issuance.status if row and row.issuance else "—"
            p(f"  {kit.title_system} E={kit.status_sheet!r} lastTDO {last.date} ev_rev={rev} sheet={kit.sheet_revision_text} rd={rd} iss={iss}")
    p(f"  total last-F tdo_passed: {n}")

    p("")
    p("=== issuance status На рассмотрении — examples ===")
    n = 0
    for key, items in all_iss.items():
        last = max(items, key=lambda i: (i.send_date_sortable, i.row_index))
        if "рассмотрен" not in last.status.casefold():
            continue
        n += 1
        if n <= 6:
            kit = next((k for k in kits if kit_identity_key(k.title, k.mark) == key), None)
            last_f = None
            if kit:
                dated = [e for e in kit.events if e.date]
                last_f = dated[-1] if dated else None
            p(
                f"  {last.title_system} iss={last.status!r} send={last.send_date_sortable} "
                f"rev={last.revision_text} E={kit.status_sheet if kit else '—'!r} "
                f"lastF={last_f.stage_label if last_f else '—'} {last_f.date if last_f else ''}"
            )
    p(f"  total latest-issuance 'на рассмотрении': {n}")

    p("")
    p("=== most issuance sends, same rev repeated ===")
    ranked = sorted(all_iss.items(), key=lambda kv: -len(kv[1]))[:8]
    for key, items in ranked:
        by_rev = Counter(i.revision_text or "—" for i in items)
        p(f"  {items[0].title_system} sends={len(items)} revs={dict(by_rev)}")

    # pick one multi-send same rev
    for key, items in all_iss.items():
        by_rev = Counter(i.revision_text or "—" for i in items)
        rev, n = by_rev.most_common(1)[0]
        if n >= 4:
            dump_kit(items[0].title, items[0].mark)
            break

    dest = Path("tmp") / "analyze_kits_status_examples.out.txt"
    dest.write_text("\n".join(out), encoding="utf-8")
    print(f"wrote {dest} lines={len(out)}")


if __name__ == "__main__":
    os.environ.setdefault("PYTHONUTF8", "1")
    main()
