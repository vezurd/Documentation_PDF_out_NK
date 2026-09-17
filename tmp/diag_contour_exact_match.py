"""Measure approved-contour exact vs base-only package matches.

Read-only against ``tmp/rd_catalog_copy.sqlite3`` (never the live DB).
Writes a before-snapshot JSON on the first run; later runs print the
before/after blast radius (package and MTO filename changes).
"""

from __future__ import annotations

import json
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog.ban_filter import BanFilterStore  # noqa: E402
from rd_catalog.config import load_config  # noqa: E402
from rd_catalog.db import CatalogDatabase  # noqa: E402
from rd_catalog.kits import kit_identity_key, parse_sheet_revision  # noqa: E402
from rd_catalog.overlay import revision_rank  # noqa: E402
from rd_catalog.pipeline import (  # noqa: E402
    resolve_all_approved_contours,
    revision_texts_equivalent,
)

COPY_DB = ROOT / "tmp" / "rd_catalog_copy.sqlite3"
SNAPSHOT = ROOT / "tmp" / "diag_contour_exact_match_before.json"
REPORT = ROOT / "tmp" / "diag_contour_exact_match_out.txt"
FOCUS_KIT = ("6100", "SOS")


def _out_lines(lines: list[str]) -> None:
    text = "\n".join(lines) + "\n"
    REPORT.write_text(text, encoding="utf-8")
    sys.stdout.buffer.write(text.encode("utf-8", errors="backslashreplace"))


def _folder_name(path: str) -> str:
    return Path(path).name if path else ""


def _file_name(path: str) -> str:
    return Path(path).name if path else ""


def _match_kind(package_revision: str, approved: str) -> str:
    if not approved:
        return "neither" if not package_revision else "neither"
    if revision_texts_equivalent(package_revision, approved):
        return "exact"
    cand_rev, _cand_app = parse_sheet_revision(package_revision)
    appr_rev, _appr_app = parse_sheet_revision(approved)
    if cand_rev and appr_rev:
        cand_base = revision_rank(cand_rev, None)[0]
        appr_base = revision_rank(appr_rev, None)[0]
        if cand_base == appr_base:
            return "base_only"
    return "neither"


def _contour_row(contour) -> dict[str, object]:
    return {
        "title": contour.title,
        "mark": contour.mark,
        "approved": contour.approved_revision_text,
        "package_revision": "",
        "package_name": _folder_name(contour.package_path),
        "package_sequence": contour.package_sequence,
        "package_path": contour.package_path,
        "mto_name": _file_name(contour.mto_path),
        "mto_path": contour.mto_path,
        "mto_source": contour.mto_source,
        "match_reason": contour.match_reason,
        "confidence": contour.confidence,
        "warnings": list(contour.warnings),
    }


def main() -> None:
    if not COPY_DB.is_file():
        raise SystemExit(f"missing copy database: {COPY_DB}")

    config = load_config()
    database = CatalogDatabase(COPY_DB)
    records = database.list_files()
    overlay_ids = {
        int(row["file_id"])
        for row in database.current_overlay()
        if row.get("detected_current")
    }
    banned = {
        pair.identity for pair in BanFilterStore.from_runtime_dir(config.runtime_dir).pairs()
    }
    packages_by_id = {pkg.id: pkg for pkg in database.list_kit_packages() if pkg.id}

    contours = [
        contour
        for contour in resolve_all_approved_contours(
            database,
            records=records,
            detected_current_ids=overlay_ids,
            rd_root=config.rd_root,
        )
        if kit_identity_key(contour.title, contour.mark) not in banned
    ]

    rows: list[dict[str, object]] = []
    kinds: Counter[str] = Counter()
    conf: Counter[str] = Counter()
    reasons: Counter[str] = Counter()
    for contour in contours:
        row = _contour_row(contour)
        package = packages_by_id.get(contour.package_id) if contour.package_id else None
        package_rev = (package.revision_text if package is not None else "") or ""
        row["package_revision"] = package_rev
        kind = (
            "no_package"
            if contour.package_id is None
            else _match_kind(package_rev, contour.approved_revision_text)
        )
        row["match_kind"] = kind
        kinds[kind] += 1
        conf[str(contour.confidence or "")] += 1
        reasons[str(contour.match_reason or "none")] += 1
        rows.append(row)

    lines: list[str] = []
    lines.append(f"database: {COPY_DB}")
    lines.append(f"non-banned kits: {len(rows)}")
    lines.append("")
    lines.append("revision match vs approved_revision_text:")
    for key in ("exact", "base_only", "neither", "no_package"):
        lines.append(f"  {key:12s} {kinds[key]:4d}")
    lines.append("")
    lines.append("confidence:")
    for key in ("high", "medium", "low"):
        share = 100 * conf[key] / max(len(rows), 1)
        lines.append(f"  {key:8s} {conf[key]:4d}  ({share:.0f}%)")
    lines.append("")
    lines.append("match_reason:")
    for reason, count in sorted(reasons.items(), key=lambda item: (-item[1], item[0])):
        lines.append(f"  {count:4d}  {reason}")

    focus_key = kit_identity_key(*FOCUS_KIT)
    focus = next(
        (
            row
            for row in rows
            if kit_identity_key(str(row["title"]), str(row["mark"])) == focus_key
        ),
        None,
    )
    lines.append("")
    lines.append("=== 6100/SOS ===")
    if focus is None:
        lines.append("  not found")
    else:
        for key in (
            "approved",
            "package_name",
            "package_sequence",
            "package_revision",
            "mto_name",
            "match_reason",
            "confidence",
            "match_kind",
        ):
            lines.append(f"  {key}: {focus[key]!r}")
        warnings = focus.get("warnings") or []
        if warnings:
            for warning in warnings:
                lines.append(f"  ! {warning}")
        else:
            lines.append("  warnings: []")

    payload = {
        "n": len(rows),
        "kinds": dict(kinds),
        "confidence": dict(conf),
        "reasons": dict(reasons),
        "rows": rows,
    }

    if not SNAPSHOT.is_file():
        SNAPSHOT.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        lines.append("")
        lines.append(f"wrote before-snapshot: {SNAPSHOT}")
        _out_lines(lines)
        return

    before = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
    before_rows = {
        kit_identity_key(str(row["title"]), str(row["mark"])): row
        for row in before.get("rows") or ()
    }
    after_rows = {
        kit_identity_key(str(row["title"]), str(row["mark"])): row
        for row in rows
    }
    keys = sorted(set(before_rows) | set(after_rows))
    mto_changes: list[tuple[str, dict, dict]] = []
    package_only: list[tuple[str, dict, dict]] = []
    confidence_only = 0
    for key in keys:
        old = before_rows.get(key)
        new = after_rows.get(key)
        if old is None or new is None:
            mto_changes.append((f"{key[0]}/{key[1]}", old or {}, new or {}))
            continue
        old_mto = str(old.get("mto_name") or "")
        new_mto = str(new.get("mto_name") or "")
        old_pkg = str(old.get("package_name") or "")
        new_pkg = str(new.get("package_name") or "")
        label = f"{new.get('title')}/{new.get('mark')}"
        if old_mto != new_mto:
            mto_changes.append((label, old, new))
        elif old_pkg != new_pkg:
            package_only.append((label, old, new))
        elif str(old.get("confidence") or "") != str(new.get("confidence") or ""):
            confidence_only += 1

    lines.append("")
    lines.append("=== before vs after ===")
    lines.append(
        "before kinds: "
        + ", ".join(
            f"{k}={before.get('kinds', {}).get(k, 0)}"
            for k in ("exact", "base_only", "neither", "no_package")
        )
    )
    lines.append(
        "before confidence: "
        + ", ".join(
            f"{k}={before.get('confidence', {}).get(k, 0)}"
            for k in ("high", "medium", "low")
        )
    )
    lines.append(f"kits whose chosen MTO file changed: {len(mto_changes)}")
    lines.append(f"kits whose package changed but MTO name did not: {len(package_only)}")
    lines.append(f"kits with only confidence change: {confidence_only}")

    shown = mto_changes[:40]
    if mto_changes:
        lines.append("")
        lines.append("MTO changes (every kit, first 40 if more than 40):")
        for label, old, new in shown:
            lines.append(f"  {label}")
            lines.append(
                "    before: "
                f"NN={old.get('package_sequence')!r} {old.get('package_name')!r} "
                f"rev={old.get('package_revision')!r} "
                f"MTO={old.get('mto_name')!r} "
                f"{old.get('match_reason')}/{old.get('confidence')}"
            )
            lines.append(
                "    after:  "
                f"NN={new.get('package_sequence')!r} {new.get('package_name')!r} "
                f"rev={new.get('package_revision')!r} "
                f"MTO={new.get('mto_name')!r} "
                f"{new.get('match_reason')}/{new.get('confidence')}"
            )
        if len(mto_changes) > 40:
            lines.append(f"  … {len(mto_changes) - 40} more")

    if package_only:
        lines.append("")
        lines.append("package-only changes (MTO filename unchanged):")
        for label, old, new in package_only[:20]:
            lines.append(
                f"  {label}: {old.get('package_name')!r} → {new.get('package_name')!r} "
                f"MTO={new.get('mto_name')!r}"
            )
        if len(package_only) > 20:
            lines.append(f"  … {len(package_only) - 20} more")

    _out_lines(lines)


if __name__ == "__main__":
    main()
