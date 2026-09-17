"""Read-only dump of kit_pipeline review vs approval combinations."""

from __future__ import annotations

import json
import os
import sqlite3
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.config import load_config
from rd_catalog.kits import format_revision, parse_history_line, parse_sheet_revision
from rd_catalog.overlay import revision_rank
from rd_catalog.pipeline import (
    pipeline_approval_label,
    pipeline_review_label,
    pipeline_status_label,
    revision_texts_equivalent,
)


def _rank(text: str) -> tuple:
    rev, app = parse_sheet_revision(text)
    return revision_rank(rev, app)


def _rel(left: str, right: str) -> str:
    a = (left or "").strip()
    b = (right or "").strip()
    if not a and not b:
        return "both_empty"
    if not a:
        return "left_empty"
    if not b:
        return "right_empty"
    if revision_texts_equivalent(a, b):
        return "equal"
    if _rank(a) > _rank(b):
        return "left_ahead"
    if _rank(a) < _rank(b):
        return "left_behind"
    return "incomparable"


def main() -> None:
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    config = load_config()
    db_path = config.db_path
    out = {
        "db_path": str(db_path),
        "exists": db_path.exists(),
    }
    if not db_path.exists():
        print(json.dumps(out, ensure_ascii=False, indent=2))
        return
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row

    pipelines = list(conn.execute("SELECT * FROM kit_pipeline"))
    google_by_key = {
        (r["title"].casefold(), r["mark"].casefold()): r
        for r in conn.execute("SELECT * FROM google_kit")
    }
    sends = list(conn.execute("SELECT * FROM issuance_send"))
    latest_send: dict[tuple[str, str], sqlite3.Row] = {}
    for row in sends:
        key = (row["title"].casefold(), row["mark"].casefold())
        prev = latest_send.get(key)
        if prev is None:
            latest_send[key] = row
            continue
        prev_rank = (prev["send_date_sortable"] or "", prev["row_index"] or 0)
        new_rank = (row["send_date_sortable"] or "", row["row_index"] or 0)
        if new_rank >= prev_rank:
            latest_send[key] = row

    events_by_kit: dict[tuple[str, str], list[sqlite3.Row]] = defaultdict(list)
    for row in conn.execute(
        """
        SELECT k.title, k.mark, e.seq, e.raw, e.event_date, e.stage,
               e.stage_label, e.revision, e.appendix
        FROM google_event AS e
        JOIN google_kit AS k ON k.id = e.kit_id
        ORDER BY k.title, k.mark, e.seq
        """
    ):
        events_by_kit[(row["title"].casefold(), row["mark"].casefold())].append(row)

    packages = list(
        conn.execute(
            """
            SELECT title, mark, source, sequence, transfer_name, revision_text,
                   is_grey, overlay_current_count, package_path
            FROM kit_package
            WHERE source = 'rd'
            """
        )
    )
    pkgs_by_kit: dict[tuple[str, str], list[sqlite3.Row]] = defaultdict(list)
    for row in packages:
        pkgs_by_kit[(row["title"].casefold(), row["mark"].casefold())].append(row)

    combo = Counter()
    rel_official_code = Counter()
    rel_official_google = Counter()
    rel_official_issuance = Counter()
    rel_code_google = Counter()
    confusing = Counter()
    samples: dict[str, list[dict]] = defaultdict(list)

    def add_sample(bucket: str, payload: dict, limit: int = 8) -> None:
        if len(samples[bucket]) < limit:
            samples[bucket].append(payload)

    rows_out: list[dict] = []
    for p in pipelines:
        key = (p["title"].casefold(), p["mark"].casefold())
        g = google_by_key.get(key)
        s = latest_send.get(key)
        official = p["official_revision_text"] or ""
        working = p["working_revision_text"] or ""
        code = p["code"] or ""
        code_rev = p["code_revision_text"] or ""
        stale = bool(p["code_stale"])
        status = p["status"] or ""
        google_rev = (g["sheet_revision_text"] if g is not None else "") or ""
        google_status = (g["status_sheet"] if g is not None else "") or ""
        iss_rev = (s["revision_text"] if s is not None else "") or ""
        iss_status = (s["status"] if s is not None else "") or ""
        oc = _rel(official, code_rev) if code else "no_code"
        og = _rel(official, google_rev) if g is not None else "no_google"
        oi = _rel(official, iss_rev) if s is not None else "no_issuance"
        cg = _rel(code_rev, google_rev) if code and g is not None else "skip"
        combo[(status, code or "—", "stale" if stale else "current")] += 1
        rel_official_code[oc] += 1
        rel_official_google[og] += 1
        rel_official_issuance[oi] += 1
        if cg != "skip":
            rel_code_google[cg] += 1

        pattern = None
        if status == "agreed" and code and stale and oc == "left_behind":
            pattern = "agreed_official_behind_stale_letter"
        elif status == "agreed" and code and stale and oc == "left_ahead":
            pattern = "agreed_official_ahead_stale_letter"
        elif status == "agreed" and code and stale and oc == "equal":
            pattern = "agreed_same_rev_stale_letter"
        elif status == "agreed" and (not code):
            pattern = "agreed_no_letter"
        elif status == "tdo_review" and code and not stale:
            pattern = "tdo_with_current_letter"
        elif status == "tdo_review" and code and stale and oc == "left_behind":
            pattern = "tdo_official_behind_stale_letter"
        elif status == "tdo_review" and code and stale:
            pattern = "tdo_stale_letter"
        elif status == "agreed" and code and not stale and oc == "equal":
            pattern = "happy_agreed_current_a"
        elif status == "agreed" and code and not stale and oc != "equal":
            pattern = "agreed_current_letter_rev_mismatch"
        elif google_rev and official and og == "left_behind" and status == "agreed":
            pattern = "agreed_but_google_ahead"
        if pattern:
            confusing[pattern] += 1

        payload = {
            "title": p["title"],
            "mark": p["mark"],
            "status": status,
            "status_label": pipeline_status_label(status),
            "official": official,
            "working": working,
            "code": code or "—",
            "code_rev": code_rev,
            "code_date": p["code_date"] or "",
            "stale": stale,
            "tdo_date": p["tdo_date"] or "",
            "google_rev": google_rev,
            "google_status": google_status,
            "iss_rev": iss_rev,
            "iss_status": iss_status,
            "oc": oc,
            "og": og,
            "oi": oi,
        }
        if pattern:
            add_sample(pattern, payload)
        rows_out.append(payload)

    focus_keys = [("7130", "SOT"), ("6816", "KSB"), ("9000", "KSB"), ("2000", "KSB")]
    focused = {}
    for title, mark in focus_keys:
        key = (title.casefold(), mark.casefold())
        p = next(
            (
                row
                for row in pipelines
                if row["title"].casefold() == key[0]
                and row["mark"].casefold() == key[1]
            ),
            None,
        )
        if p is None:
            focused[f"{title}-{mark}"] = None
            continue
        g = google_by_key.get(key)
        evs = events_by_kit.get(key, [])
        sends_kit = [
            row
            for row in sends
            if row["title"].casefold() == key[0] and row["mark"].casefold() == key[1]
        ]
        pkgs = pkgs_by_kit.get(key, [])
        from rd_catalog.db import KitPipelineRow

        pipeline_row = KitPipelineRow(
            title=p["title"],
            mark=p["mark"],
            status=p["status"],
            code=p["code"],
            code_origin=p["code_origin"],
            working_revision_text=p["working_revision_text"] or "",
            official_revision_text=p["official_revision_text"] or "",
            suspicious=bool(p["suspicious"]),
            algorithm_version=int(p["algorithm_version"] or 1),
            code_stale=bool(p["code_stale"]),
            code_revision_text=p["code_revision_text"] or "",
            code_date=p["code_date"] or "",
            tdo_date=p["tdo_date"] or "",
            review_as_build=bool(p["review_as_build"]) if "review_as_build" in p.keys() else False,
            working_as_build=bool(p["working_as_build"]) if "working_as_build" in p.keys() else False,
        )
        events = []
        for row in evs:
            events.append(
                parse_history_line(row["raw"])
                if row["raw"]
                else None
            )
        events = [item for item in events if item is not None]
        focused[f"{title}-{mark}"] = {
            "pipeline": dict(p),
            "review_label": pipeline_review_label(pipeline_row),
            "approval_label": pipeline_approval_label(pipeline_row),
            "google": {
                "rev": g["sheet_revision_text"] if g is not None else "",
                "status": g["status_sheet"] if g is not None else "",
                "comment": (g["comment_raw"] if g is not None else "")[:2000],
            }
            if g is not None
            else None,
            "events": [
                {
                    "seq": row["seq"],
                    "date": row["event_date"],
                    "stage": row["stage"],
                    "label": row["stage_label"],
                    "rev": format_revision(row["revision"], row["appendix"]),
                    "raw": row["raw"],
                }
                for row in evs
            ],
            "sends": [
                {
                    "rev": row["revision_text"],
                    "status": row["status"],
                    "send_date": row["send_date"],
                    "incoming": row["incoming_control_date"],
                    "trm": row["send_transmittal"],
                    "confirm": row["confirm_transmittal"],
                    "row_index": row["row_index"],
                }
                for row in sends_kit
            ],
            "packages": [
                {
                    "seq": row["sequence"],
                    "name": row["transfer_name"],
                    "rev": row["revision_text"],
                    "grey": bool(row["is_grey"]),
                    "overlay": row["overlay_current_count"],
                }
                for row in pkgs
            ],
        }

    status_counts = Counter(p["status"] for p in pipelines)
    code_counts = Counter((p["code"] or "—") for p in pipelines)
    stale_counts = Counter(bool(p["code_stale"]) for p in pipelines)
    status_stale = Counter(
        (
            p["status"],
            p["code"] or "—",
            "stale" if p["code_stale"] else "current",
        )
        for p in pipelines
    )

    # agreed + stale split by official vs code rev
    agreed_stale_rel = Counter()
    for p in pipelines:
        if p["status"] != "agreed" or not p["code"] or not p["code_stale"]:
            continue
        agreed_stale_rel[
            _rel(p["official_revision_text"] or "", p["code_revision_text"] or "")
        ] += 1

    report = {
        "db_path": str(db_path),
        "n_pipeline": len(pipelines),
        "n_google": len(google_by_key),
        "n_sends": len(sends),
        "status_counts": dict(status_counts),
        "code_counts": dict(code_counts),
        "stale_true": stale_counts[True],
        "stale_false": stale_counts[False],
        "combo": {f"{a}|{b}|{c}": n for (a, b, c), n in status_stale.most_common()},
        "rel_official_code": dict(rel_official_code),
        "rel_official_google": dict(rel_official_google),
        "rel_official_issuance": dict(rel_official_issuance),
        "rel_code_google": dict(rel_code_google),
        "patterns": dict(confusing),
        "agreed_stale_official_vs_code": dict(agreed_stale_rel),
        "samples": samples,
        "focused": focused,
        "equiv_01_01an01": revision_texts_equivalent("01", "01-AN01"),
        "equiv_0_0an01": revision_texts_equivalent("0", "0-AN01"),
        "equiv_01_1": revision_texts_equivalent("01", "1"),
    }
    out_path = Path(__file__).with_name("analyze_review_approval.json")
    out_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    print(f"wrote {out_path}")
    print("n", len(pipelines), "patterns", dict(confusing))
    print("combo top")
    for key, n in status_stale.most_common(20):
        print(f"  {n:4d}  {key}")


if __name__ == "__main__":
    main()
