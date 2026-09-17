"""Simulate which official package 2235-KSB uses."""

from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from rd_catalog.db import CatalogDatabase
from rd_catalog.kits import kit_identity_key
from rd_catalog.models import SourceKind
from rd_catalog.pipeline import _exclusion_from_pipeline, _official_package_files

DB = (
    Path(os.environ["LOCALAPPDATA"])
    / "Documentation_PDF_out_NK"
    / "rd_catalog"
    / "rd_catalog.sqlite"
)


def main() -> None:
    db = CatalogDatabase(DB)
    key = kit_identity_key("2235", "KSB")
    pipe = next(
        row
        for row in db.list_kit_pipelines()
        if kit_identity_key(row.title, row.mark) == key
    )
    records = [
        rec
        for rec in db.list_files(source=SourceKind.RD, present_only=True)
        if kit_identity_key(
            str(rec.data.get("title") or ""),
            str(rec.data.get("mark") or ""),
        )
        == key
    ]
    exclusion = _exclusion_from_pipeline(pipe)
    chosen = _official_package_files(
        records,
        key,
        pipe.official_revision_text,
        pipe.working_revision_text,
        exclusion=exclusion,
    )
    folders = sorted({str(rec.data.get("transfer_name") or "") for rec in chosen})
    kinds = {}
    for rec in chosen:
        rev = f"{rec.data.get('revision')}-{rec.data.get('appendix')}"
        kinds.setdefault(str(rec.data.get("file_kind")), set()).add(rev)
    lines = [
        f"working={pipe.working_revision_text}",
        f"official_expected={pipe.official_revision_text}",
        f"working_seq={pipe.working_sequences}",
        f"working_folders={list(pipe.working_transfer_names)}",
        f"chosen_folders={folders}",
        f"chosen_n={len(chosen)}",
        f"chosen_revs={ {k: sorted(v) for k, v in kinds.items()} }",
    ]
    Path(__file__).with_suffix(".txt").write_text(
        "\n".join(str(item) for item in lines) + "\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
