"""Verify the pipeline fix by running extract_page on all MTO PDFs."""
from __future__ import annotations

import os
import sys
from pathlib import Path

import fitz

from pdf_parsing_v2_engine.models import StampTemplate
from pdf_parsing_v2_engine.stamp_extractor import extract_page

_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_MTO_DIR = _PROJECT_ROOT / "pdf_parsing_v2_engine" / "templates" / "test_pdf" / "MTO"
_TEMPLATES_DIR = _PROJECT_ROOT / "pdf_parsing_v2_engine" / "templates"

_BAD = ("4000-KSB", "5631-KSB", "7421-SOS.1", "8525-SOS")
_GOOD = ("7570-SOT", "8441-SKUD", "8630-KSB3")


def _load_mto_templates():
    import json
    out = []
    for f in sorted(os.listdir(_TEMPLATES_DIR)):
        if not f.endswith(".json"):
            continue
        path = os.path.join(_TEMPLATES_DIR, f)
        if not os.path.isfile(path):
            continue
        try:
            t = StampTemplate.from_json(path)
            if "MTO" in t.doc_types:
                out.append(t)
        except (OSError, ValueError, json.JSONDecodeError):
            pass
    return out


def _label(fname):
    for b in _BAD:
        if b in fname:
            return "BAD "
    for g in _GOOD:
        if g in fname:
            return "GOOD"
    return "??? "


def main():
    templates = _load_mto_templates()
    print(f"Loaded {len(templates)} MTO templates")

    mto_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else _MTO_DIR
    pdfs = sorted(mto_dir.glob("*.pdf"))
    print(f"Testing {len(pdfs)} PDFs with extract_page\n")

    hdr = f"{'label':5s} {'rot':>4s} {'score':>6s} {'50_File':>8s} {'26_Rev':>8s} {'6_1_Sh':>8s} {'1_DOC':>8s}  name"
    print(hdr)
    print("-" * len(hdr) + "-" * 40)

    for pdf in pdfs:
        doc = fitz.open(str(pdf))
        page = doc[0]
        rot = page.rotation % 360

        result = extract_page(page, "MTO", 1, templates)
        doc.close()

        def _val(fid):
            fr = result.fields.get(fid)
            if fr and fr.cleaned_value and fr.cleaned_value.strip():
                return "OK"
            return "EMPTY"

        label = _label(pdf.name)
        print(
            f"{label:5s} {rot:4d} {result.template_score:6.2f} "
            f"{_val('50_File_Name_Stamp'):>8s} {_val('26_Document_Revision'):>8s} "
            f"{_val('6_1_Sheet_number'):>8s} {_val('1_DOC_TITLE'):>8s}  {pdf.name}"
        )

    # Show field values for a few interesting cases
    print("\n\n=== Sample extracted values ===")
    for pdf in pdfs:
        doc = fitz.open(str(pdf))
        page = doc[0]
        rot = page.rotation % 360
        result = extract_page(page, "MTO", 1, templates)
        doc.close()

        if rot != 90:
            continue

        print(f"\n--- {_label(pdf.name)} rot={rot} {pdf.name} ---")
        for fid in ("1_DOC_TITLE", "26_Document_Revision", "6_1_Sheet_number", "50_File_Name_Stamp"):
            fr = result.fields.get(fid)
            val = (fr.cleaned_value or "").strip()[:60] if fr else ""
            status = "OK" if val else "EMPTY"
            print(f"  [{status:5s}] {fid:30s} = {val!r}")


if __name__ == "__main__":
    main()
