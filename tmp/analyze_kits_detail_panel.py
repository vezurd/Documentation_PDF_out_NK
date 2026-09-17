"""Read-only analysis of live RD catalog DB + Google caches for the kits detail pane."""

from __future__ import annotations

import os
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from rd_catalog.config import load_config
from rd_catalog.db import CatalogDatabase
from rd_catalog.google_kits import load_cached_google_kits
from rd_catalog.kits import (
    KitSummary,
    build_kit_matrix,
    format_revision,
    kit_identity_key,
    last_event_parts,
    summary_label,
)
from rd_catalog.models import FileKind, SourceKind
from rd_catalog.parse import (
    is_transfer_folder_name,
    is_transfer_gate_folder_name,
)


def _mtime(ns: int | None) -> str:
    if not ns:
        return "—"
    try:
        return datetime.fromtimestamp(int(ns) / 1_000_000_000).strftime("%Y-%m-%d")
    except (OSError, OverflowError, ValueError):
        return "—"


def _parts(path: str) -> list[str]:
    text = path.replace("/", "\\").rstrip("\\")
    if text.startswith("\\\\"):
        rest = text[2:]
        segs = [p for p in rest.split("\\") if p]
        if len(segs) >= 2:
            return ["\\\\" + segs[0], segs[1], *segs[2:]]
        return segs
    return [p for p in text.split("\\") if p]


def package_folder(path: str) -> tuple[str, str]:
    """Return (package_or_parent_path, kind)."""

    segs = _parts(path)
    if not segs:
        return "", "empty"
    # Drop filename.
    dirs = segs[:-1]
    # Drop PDF/DWG leaf folders.
    while dirs and dirs[-1].casefold() in {"pdf", "dwg"}:
        dirs = dirs[:-1]
    if not dirs:
        return "\\".join(segs[:-1]), "parent"
    ancestors = dirs[:-1]
    under_gate = any(is_transfer_gate_folder_name(name) for name in ancestors)
    if is_transfer_folder_name(dirs[-1], under_gate=under_gate):
        return "\\".join(dirs), "transfer"
    if any(is_transfer_gate_folder_name(name) for name in dirs):
        return "\\".join(dirs), "gate_or_mark"
    return "\\".join(dirs), "parent"


def main() -> None:
    cfg = load_config()
    out: list[str] = []

    def p(line: str = "") -> None:
        out.append(line)

    p(f"db: {cfg.db_path}")
    p(f"runtime: {cfg.runtime_dir}")
    p(f"sqlite exists: {Path(cfg.db_path).is_file()} size={Path(cfg.db_path).stat().st_size}")
    p()

    db = CatalogDatabase(cfg.db_path)
    records = db.list_files()
    overlay_rows = db.current_overlay()
    current_ids = {int(row["file_id"]) for row in overlay_rows}
    collisions = db.list_current_collisions()
    mto_rows = db.list_mto_comparisons()

    present = [r for r in records if r.present]
    p("=== SQLite files ===")
    p(f"file_entry total={len(records)} present={len(present)}")
    src_kind = Counter(
        (r.source.value, str(r.data.get("file_kind") or ""), int(r.present))
        for r in records
    )
    for (src, kind, pr), n in sorted(src_kind.items()):
        p(f"  {src:6} kind={kind:16} present={pr} n={n}")
    p(f"overlay current ids={len(current_ids)}")
    p(f"collisions={len(collisions)}")
    col_kinds = Counter(str(c.get("kind") or "") for c in collisions)
    for k, n in col_kinds.most_common():
        p(f"  collision {k}: {n}")
    mto_status = Counter(str(r.get("status") or "") for r in mto_rows)
    p(f"mto_comparisons={len(mto_rows)} {dict(mto_status)}")

    p()
    p("=== Path / transfer packages (present files) ===")
    pkg_kind = Counter()
    pkg_by_source: dict[str, set[str]] = defaultdict(set)
    files_per_pkg: dict[str, int] = Counter()
    sample_by_kind: dict[str, list[str]] = defaultdict(list)
    leaf_folders = Counter()
    for rec in present:
        pkg, kind = package_folder(rec.path)
        pkg_kind[f"{rec.source.value}:{kind}"] += 1
        pkg_by_source[rec.source.value].add(pkg)
        files_per_pkg[pkg] += 1
        segs = _parts(rec.path)
        if len(segs) >= 2:
            leaf_folders[f"{rec.source.value}:{segs[-2]}"] += 1
        if len(sample_by_kind[f"{rec.source.value}:{kind}"]) < 3:
            sample_by_kind[f"{rec.source.value}:{kind}"].append(pkg or rec.path)
    for k, n in sorted(pkg_kind.items()):
        p(f"  {k}: files={n}")
    for src, pkgs in sorted(pkg_by_source.items()):
        p(f"  unique package folders {src}: {len(pkgs)}")
    counts = list(files_per_pkg.values())
    if counts:
        p(
            "  files per package: "
            f"min={min(counts)} median={sorted(counts)[len(counts)//2]} "
            f"max={max(counts)} mean={sum(counts)/len(counts):.1f}"
        )
    p("  samples:")
    for k, samples in sorted(sample_by_kind.items()):
        p(f"    {k}:")
        for s in samples:
            p(f"      {s}")
    p("  top immediate parent folder names:")
    for name, n in leaf_folders.most_common(15):
        p(f"    {n:5} {name}")

    p()
    p("=== RD transfer_name / sequence (present parsed) ===")
    rd_present = [r for r in present if r.source is SourceKind.RD]
    seqs = Counter()
    names = Counter()
    kits_transfers: dict[tuple[str, str], set[str]] = defaultdict(set)
    for r in rd_present:
        title = str(r.data.get("title") or "")
        mark = str(r.data.get("mark") or "")
        tname = str(r.data.get("transfer_name") or "")
        seq = r.data.get("transfer_sequence")
        if tname:
            names[tname] += 1
            if title and mark:
                kits_transfers[(title, mark)].add(tname)
        if seq not in (None, ""):
            seqs[int(seq)] += 1
    p(f"  unique transfer_name values: {len(names)}")
    p(f"  files with transfer_sequence: {sum(seqs.values())} / {len(rd_present)}")
    p(f"  sequence histogram (seq -> files): {dict(sorted(seqs.items())[:20])} ... max_seq={max(seqs) if seqs else None}")
    n_pkg_per_kit = [len(v) for v in kits_transfers.values()]
    if n_pkg_per_kit:
        p(
            "  unique transfer folders per title-mark: "
            f"min={min(n_pkg_per_kit)} median={sorted(n_pkg_per_kit)[len(n_pkg_per_kit)//2]} "
            f"max={max(n_pkg_per_kit)} kits={len(n_pkg_per_kit)}"
        )
        top = sorted(kits_transfers.items(), key=lambda kv: -len(kv[1]))[:8]
        for (title, mark), folders in top:
            p(f"    {title}-{mark}: {len(folders)} pkgs")

    p()
    cached = load_cached_google_kits(cfg.runtime_dir)
    if cached is None:
        p("Google cache: MISSING")
        kits = ()
        issuance = ()
    else:
        kits = cached.kits
        issuance = cached.issuance_kits
        p(
            f"Google cache: kits={len(kits)} skipped={cached.stats.skipped} "
            f"issuance={len(issuance)} skipped_iss={cached.issuance_stats.skipped} "
            f"source={cached.source} fetched={cached.fetched_at}"
        )

    p()
    p("=== Google KSB ИД: status_sheet (col E) ===")
    status_e = Counter((k.status_sheet or "«пусто»") for k in kits)
    for val, n in status_e.most_common(40):
        p(f"  {n:4}  {val}")
    p(f"  unique status_sheet: {len(status_e)}")

    p()
    p("=== Google F events: classified stage ===")
    stage_n = Counter()
    stage_other = Counter()
    events_per_kit = []
    unparsed = 0
    for k in kits:
        events_per_kit.append(len(k.events))
        for ev in k.events:
            stage_n[f"{ev.stage} | {ev.stage_label}"] += 1
            if ev.stage == "other":
                stage_other[ev.stage_label[:80]] += 1
            if not ev.parsed:
                unparsed += 1
    p(f"  kits with events: {sum(1 for n in events_per_kit if n)} / {len(kits)}")
    if events_per_kit:
        p(
            "  events per kit: "
            f"min={min(events_per_kit)} median={sorted(events_per_kit)[len(events_per_kit)//2]} "
            f"max={max(events_per_kit)} mean={sum(events_per_kit)/len(events_per_kit):.1f}"
        )
    p(f"  unparsed events: {unparsed}")
    for val, n in stage_n.most_common():
        p(f"  {n:4}  {val}")
    p("  other leftovers (top 25):")
    for val, n in stage_other.most_common(25):
        p(f"    {n:4}  {val}")

    p()
    p("=== last F event vs status_sheet ===")
    last_stage = Counter()
    mismatch_e_vs_f = 0
    for k in kits:
        _, st, _, _ = last_event_parts(k)
        last_stage[st or "«пусто»"] += 1
        e = (k.status_sheet or "").casefold()
        f = (st or "").casefold()
        if e and f and e not in f and f not in e:
            mismatch_e_vs_f += 1
    p(f"  last F stage unique: {len(last_stage)}")
    for val, n in last_stage.most_common(20):
        p(f"  {n:4}  {val}")
    p(f"  kits where col E text and last F stage look different: {mismatch_e_vs_f}")

    p()
    p("=== Выдача РД ПД: status ===")
    iss_status = Counter((k.status or "«пусто»") for k in issuance)
    p(f"  unique issuance.status: {len(iss_status)}")
    for val, n in iss_status.most_common(30):
        p(f"  {n:4}  {val}")
    has_send = sum(1 for k in issuance if k.send_date)
    has_ctrl = sum(1 for k in issuance if k.incoming_control_date)
    has_confirm = sum(1 for k in issuance if k.confirm_transmittal)
    p(f"  with send_date={has_send} incoming_control={has_ctrl} confirm_trm={has_confirm}")

    p()
    p("=== Kit matrix (union Google + issuance + scan) ===")
    matrix = build_kit_matrix(kits, records, current_ids, issuance_kits=issuance)
    p(f"  rows={len(matrix)}")
    summaries = Counter(summary_label(r.summary) for r in matrix)
    for val, n in summaries.most_common():
        p(f"  {n:4}  {val}")
    flags = Counter()
    for r in matrix:
        for f in r.flags:
            flags[f.value] += 1
    p(f"  flags: {dict(flags)}")
    both_google = sum(1 for r in matrix if r.google and r.issuance)
    only_g = sum(1 for r in matrix if r.google and not r.issuance)
    only_i = sum(1 for r in matrix if r.issuance and not r.google)
    p(f"  google+issuance={both_google} google_only={only_g} issuance_only={only_i}")
    review = sum(1 for r in matrix if r.transfer_review_notes)
    p(f"  with transfer_review_notes={review}")

    file_counts = [r.rd.file_count for r in matrix if r.rd.present]
    if file_counts:
        p(
            "  RD overlay-current files per kit: "
            f"min={min(file_counts)} median={sorted(file_counts)[len(file_counts)//2]} "
            f"max={max(file_counts)} mean={sum(file_counts)/len(file_counts):.1f} "
            f"kits_with_rd={len(file_counts)}"
        )
        p(f"  kits with >=8 current RD files (noisy dump): {sum(1 for n in file_counts if n >= 8)}")

    # Google vs issuance status/rev mismatch
    e_vs_iss = Counter()
    rev_g_vs_iss = 0
    for r in matrix:
        if not r.google or not r.issuance:
            continue
        e = r.google.status_sheet or "—"
        s = r.issuance.status or "—"
        if e.casefold() != s.casefold():
            e_vs_iss[(e[:40], s[:40])] += 1
        if (r.google.sheet_revision_text or "") != (r.issuance.revision_text or ""):
            rev_g_vs_iss += 1
    p(f"  google sheet_rev != issuance rev (when both present): {rev_g_vs_iss}")
    p("  top col-E vs issuance.status mismatches:")
    for (e, s), n in e_vs_iss.most_common(12):
        p(f"    {n:3}  E={e!r}  vs  Выдача={s!r}")

    # Sample requested kit from screenshot
    p()
    p("=== Sample 9110-KSB1 ===")
    for r in matrix:
        if r.title == "9110" and r.mark.casefold() == "ksb1":
            p(f"  summary={r.summary.value} flags={[f.value for f in r.flags]}")
            p(f"  rd rev={r.rd.revision_text} files={r.rd.file_count} transfer={r.rd.transfer_name} mtime={_mtime(r.rd.max_mtime_ns)}")
            p(f"  robot rev={r.robot.revision_text} files={r.robot.file_count}")
            p(f"  sq rev={r.sq.revision_text} files={r.sq.file_count}")
            if r.google:
                p(f"  google sheet_rev={r.google.sheet_revision_text} status_sheet={r.google.status_sheet!r}")
                p(f"  google events={len(r.google.events)}")
                for ev in r.google.events[-8:]:
                    p(
                        f"    {ev.date} {ev.stage}/{ev.stage_label} "
                        f"rev={format_revision(ev.revision, ev.appendix) or '—'} "
                        f"trm={ev.transmittals}"
                    )
            if r.issuance:
                p(
                    f"  issuance rev={r.issuance.revision_text} status={r.issuance.status!r} "
                    f"send={r.issuance.send_date_sortable} ctrl={r.issuance.incoming_control_date_sortable}"
                )
            pkgs = []
            for path in r.rd.paths[:20]:
                pkg, kind = package_folder(path)
                pkgs.append((kind, pkg))
            unique_pkg = list(dict.fromkeys(pkg for _, pkg in pkgs))
            p(f"  unique current RD packages from paths: {len(unique_pkg)}")
            for pkg in unique_pkg[:6]:
                p(f"    {pkg}")
            # All present RD transfers for this kit (history)
            hist = defaultdict(list)
            for rec in present:
                if rec.source is not SourceKind.RD:
                    continue
                if str(rec.data.get("title") or "") != "9110":
                    continue
                if str(rec.data.get("mark") or "").casefold() != "ksb1":
                    continue
                pkg, kind = package_folder(rec.path)
                hist[pkg].append(rec)
            p(f"  all present RD packages (incl superseded): {len(hist)}")
            for pkg, files in sorted(hist.items(), key=lambda kv: kv[0]):
                kinds = Counter(str(f.data.get("file_kind")) for f in files)
                revs = sorted(
                    {
                        format_revision(str(f.data.get("revision") or "") or None, str(f.data.get("appendix") or "") or None)
                        or "—"
                        for f in files
                    }
                )
                mt = max(int(f.data.get("mtime_ns") or 0) for f in files)
                cur = sum(1 for f in files if f.id in current_ids)
                p(
                    f"    seq={files[0].data.get('transfer_sequence')} "
                    f"files={len(files)} current={cur} kinds={dict(kinds)} revs={revs} "
                    f"mtime={_mtime(mt)}"
                )
                p(f"      {pkg}")
            break
    else:
        p("  not found")

    p()
    p("=== Discipline mix in overlay-current RD (top) ===")
    discs = Counter()
    for rec in records:
        if rec.id not in current_ids:
            continue
        discs[str(rec.data.get("discipline_block") or "—")] += 1
    for val, n in discs.most_common(20):
        p(f"  {n:4}  {val}")

    dest = Path("tmp") / "analyze_kits_detail_panel.out.txt"
    dest.write_text("\n".join(out), encoding="utf-8")
    print(f"wrote {dest} lines={len(out)}")


if __name__ == "__main__":
    os.environ.setdefault("PYTHONUTF8", "1")
    main()
