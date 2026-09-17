"""
Batch probe: first page of each PDF in a folder with a fixed MTO page-1 template.

Calls ``extract_page`` directly (same engine as ``v2_pipeline``). Can be invoked from
the pipeline via ``--mto-regression`` or config ``mto_regression_after_run``.

Run from repo root::

  set PYTHONPATH=.
  python -m pdf_parsing_v2_engine.tools.mto_page1_batch_probe
  python -m pdf_parsing_v2_engine.tools.mto_page1_batch_probe --pdf-dir "..." --template "..."

Default paths match AGCC MTO regression set + ``agcc_287/mto_page1.json``.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any

import fitz

from pdf_parsing_v2_engine.models import FieldCatalog, StampTemplate, merge_stamp_template_with_catalog
from pdf_parsing_v2_engine.stamp_extractor import extract_page
from pdf_parsing_v2.v2_config import get_default_v2_config

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_DEFAULT_PDF_DIR = _PROJECT_ROOT / "pdf_parsing_v2_engine" / "templates" / "test_pdf" / "MTO"
_DEFAULT_TEMPLATE = _PROJECT_ROOT / "pdf_parsing_v2_engine" / "templates" / "agcc_287" / "mto_page1.json"


def default_mto_regression_pdf_dir() -> str:
    return str(_DEFAULT_PDF_DIR)


def default_mto_regression_template_path() -> str:
    return str(_DEFAULT_TEMPLATE)


def _load_template(template_path: str) -> StampTemplate:
    path = os.path.abspath(template_path)
    project_dir = os.path.dirname(path)
    cat_path = os.path.join(project_dir, "catalog.json")
    catalog: FieldCatalog | None = None
    if os.path.isfile(cat_path):
        catalog = FieldCatalog.from_json(cat_path)
    tmpl = StampTemplate.from_json(path)
    return merge_stamp_template_with_catalog(tmpl, catalog)


def _merged_cfg(mode: str, v2_cfg: dict[str, Any] | None) -> dict[str, Any]:
    base = dict(v2_cfg) if v2_cfg else dict(get_default_v2_config())
    base["text_extraction_mode"] = mode
    return base


def _run_extract(
    pdf_path: str,
    template: StampTemplate,
    cfg: dict,
) -> object:
    doc_type = template.doc_types[0] if template.doc_types else "MTO"
    doc = fitz.open(pdf_path)
    try:
        page = doc[0]
        return extract_page(page, doc_type, 1, [template], cfg=cfg)
    finally:
        doc.close()


def _newline_stats(fields: dict) -> tuple[int, list[str]]:
    """(total internal newlines in raw_value, list of field_id with \\n)."""
    flagged: list[str] = []
    total = 0
    for fid, fr in fields.items():
        raw = fr.raw_value or ""
        n = raw.count("\n")
        if n:
            total += n
            flagged.append(f"{fid}({n})")
    return total, flagged


def _diff_raw(
    a: dict,
    b: dict,
) -> list[str]:
    out: list[str] = []
    keys = sorted(set(a.keys()) | set(b.keys()))
    for k in keys:
        ra = (a.get(k).raw_value or "") if k in a else ""
        rb = (b.get(k).raw_value or "") if k in b else ""
        if ra != rb:
            out.append(k)
    return out


def run_mto_page1_batch_probe(
    pdf_dir: str | Path,
    template_path: str,
    *,
    v2_cfg: dict[str, Any] | None = None,
    print_report: bool = True,
    json_out: str | Path | None = None,
) -> dict[str, Any]:
    """Run char_center vs get_textbox extraction on page 1 for each PDF.

    Returns a dict with keys ``rows``, ``pdf_dir``, ``template_path`` (and optional
    ``error`` if setup failed — then ``rows`` may be empty).
    """
    pdf_dir = Path(pdf_dir)
    out: dict[str, Any] = {
        "pdf_dir": str(pdf_dir.resolve()),
        "template_path": os.path.abspath(template_path),
        "rows": [],
    }

    if not pdf_dir.is_dir():
        out["error"] = f"not a directory: {pdf_dir}"
        return out

    if not os.path.isfile(template_path):
        out["error"] = f"template not found: {template_path}"
        return out

    try:
        template = _load_template(template_path)
    except Exception as e:
        out["error"] = str(e)
        return out

    pdfs = sorted(pdf_dir.glob("*.pdf"))
    if not pdfs:
        out["error"] = f"no PDF files in {pdf_dir}"
        return out

    cfg_char = _merged_cfg("char_center", v2_cfg)
    cfg_box = _merged_cfg("get_textbox", v2_cfg)
    rows: list[dict] = []

    if print_report:
        print(
            f"[mto_regression] {len(pdfs)} PDF(s), template={template.name!r}, "
            f"doc_types={template.doc_types!r}\n",
        )
        header = (
            f"{'file':<52} {'rot':>3} {'sc_ch':>5} {'sc_tb':>5} "
            f"{'nl_ch':>5} {'diff':>4}  notes"
        )
        print(header)
        print("-" * len(header))

    for pdf in pdfs:
        path = str(pdf)
        try:
            r_char = _run_extract(path, template, cfg_char)
            r_box = _run_extract(path, template, cfg_box)
        except Exception as e:
            row = {"file": pdf.name, "error": str(e)}
            rows.append(row)
            if print_report:
                print(f"{pdf.name:<52} ERR {e}")
            continue

        rot = r_char.frame.rotation
        nl_total, nl_ids = _newline_stats(r_char.fields)
        diff_ids = _diff_raw(r_char.fields, r_box.fields)
        notes: list[str] = []
        if nl_ids:
            notes.append("newlines_char:" + ",".join(nl_ids[:6]))
            if len(nl_ids) > 6:
                notes.append("...")
        if diff_ids:
            notes.append("raw!=get_textbox:" + ",".join(diff_ids[:8]))
            if len(diff_ids) > 8:
                notes.append("...")

        if print_report:
            print(
                f"{pdf.name:<52} {rot:3d} {r_char.template_score:5.2f} "
                f"{r_box.template_score:5.2f} {nl_total:5d} {len(diff_ids):4d}  "
                f"{' | '.join(notes)}",
            )
        rows.append({
            "file": pdf.name,
            "rotation": rot,
            "score_char_center": r_char.template_score,
            "score_get_textbox": r_box.template_score,
            "newline_count_char_center": nl_total,
            "fields_with_newlines_char_center": nl_ids,
            "raw_diff_field_ids": diff_ids,
            "warnings_char": r_char.warnings[:5],
        })

    out["rows"] = rows

    if json_out:
        jpath = Path(json_out)
        jpath.parent.mkdir(parents=True, exist_ok=True)
        with open(jpath, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)
        if print_report:
            print(f"\n[mto_regression] Wrote {jpath}")

    return out


def check_mto_regression_strict(
    report: dict[str, Any],
    *,
    max_newlines_rot90: int = 280,
    min_score_vs_textbox: float = -0.05,
) -> list[str]:
    """Return human-readable failure reasons. Empty list means strict checks passed."""
    reasons: list[str] = []
    if report.get("error"):
        return [f"probe setup: {report['error']}"]

    for row in report.get("rows", []):
        if "error" in row:
            reasons.append(f"{row.get('file', '?')}: extraction error: {row['error']}")
            continue
        rot = int(row.get("rotation", 0))
        nl = int(row.get("newline_count_char_center", 0))
        sc_ch = float(row.get("score_char_center", 0.0))
        sc_tb = float(row.get("score_get_textbox", 0.0))
        name = row.get("file", "?")

        if rot in (90, 270) and nl > max_newlines_rot90:
            reasons.append(
                f"{name}: rotation={rot}, newline_count_char_center={nl} "
                f"(max allowed {max_newlines_rot90})",
            )
        if sc_ch < sc_tb + min_score_vs_textbox:
            reasons.append(
                f"{name}: score_char_center={sc_ch:.3f} < score_get_textbox+"
                f"{min_score_vs_textbox:.2f} ({sc_tb:.3f})",
            )
    return reasons


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--pdf-dir",
        type=str,
        default=str(_DEFAULT_PDF_DIR),
        help="Directory with PDFs (default: templates/test_pdf/MTO)",
    )
    p.add_argument(
        "--template",
        type=str,
        default=str(_DEFAULT_TEMPLATE),
        help="Path to mto_page1.json (default: agcc_287/mto_page1.json)",
    )
    p.add_argument(
        "--json-out",
        type=str,
        default="",
        help="If set, write machine-readable report to this path",
    )
    p.add_argument(
        "--strict",
        action="store_true",
        help="Exit 1 if rotation 90/270 pages exceed newline budget or score vs get_textbox",
    )
    p.add_argument(
        "--max-nl-rot90",
        type=int,
        default=280,
        help="With --strict: max total newlines in char_center raw on 90/270 pages (default 280)",
    )
    args = p.parse_args(argv)

    report = run_mto_page1_batch_probe(
        args.pdf_dir,
        args.template,
        v2_cfg=None,
        print_report=True,
        json_out=args.json_out or None,
    )
    if report.get("error"):
        print(f"ERROR: {report['error']}", file=sys.stderr)
        return 2

    if args.strict:
        bad = check_mto_regression_strict(
            report,
            max_newlines_rot90=args.max_nl_rot90,
        )
        if bad:
            print("[mto_regression] STRICT FAIL:", file=sys.stderr)
            for line in bad:
                print(f"  {line}", file=sys.stderr)
            return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
