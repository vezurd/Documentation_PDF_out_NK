"""Local runner for the paired stamp deformation benchmark dataset.

This script is meant to be edited directly and launched as a standalone file:

    python pdf_parsing_v2_engine/templates/stamp_deform_bench/run_stamp_deform_bench.py

It does not read project config files. All important run parameters are defined
below in one place.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

DATASET_DIR = Path(__file__).resolve().parent
REPO_ROOT = DATASET_DIR.parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from pdf_parsing_v2_engine.stamp_deform_generator import GenerationConfig, generate_deformed_stamp_dataset


def _pick_single_file(folder: Path, pattern: str) -> Path:
    """Return the single matching file from a folder."""
    matches = sorted(path for path in folder.glob(pattern) if path.is_file())
    if len(matches) != 1:
        pretty = ", ".join(path.name for path in matches[:10]) or "<none>"
        raise SystemExit(
            f"Expected exactly one file matching {pattern!r} in {folder}, got {len(matches)}: {pretty}"
        )
    return matches[0]


# ---------------------------------------------------------------------------
# Dataset paths
# ---------------------------------------------------------------------------
# You can point to a single PDF file or to a whole directory.
STAMPED_INPUT = DATASET_DIR / "input" / "stamped"
BACKGROUND_INPUT = DATASET_DIR / "input" / "background"

# By default this runner expects exactly one template JSON in template/.
# If you prefer an explicit file path, replace the line below with e.g.:
# TEMPLATE_PATH = DATASET_DIR / "template" / "mto_page1.json"
TEMPLATE_PATH = _pick_single_file(DATASET_DIR / "template", "*.json")

# Outputs.
OUTPUT_DIR = DATASET_DIR / "output" / "generated"
MANIFEST_JSONL = DATASET_DIR / "output" / "manifests" / "manifest.jsonl"
MANIFEST_CSV = DATASET_DIR / "output" / "manifests" / "manifest.csv"


# ---------------------------------------------------------------------------
# Main run parameters
# ---------------------------------------------------------------------------
# 1-based page number containing the stamp in the source/background pair.
PAGE_NUM = 1

# Deformation ranges are delta-percent относительно оригинала.
# Example: -25 -> scale 75%, +10 -> scale 110%.
X_RANGE = (-20, 20)
Y_RANGE = (-20, 20)

# Common choices:
#   1  -> dense benchmark
#   5  -> coarse benchmark
#   10 -> fast smoke run
STEP = 4

# Available modes:
#   "grid"      -> all X x Y combinations
#   "diagonal"  -> sx == sy
#   "axis-x"    -> vary only X, keep Y = 100%
#   "axis-y"    -> vary only Y, keep X = 100%
#   "cross"     -> X-axis + Y-axis quick check
MODE = "diagonal"

# Anchor for anisotropic scaling.
# Usually keep "bottom_right" for GOST-like stamp placement.
ANCHOR = "bottom_right"

# Safe default for first launch. Switch to False for actual PDF generation.
DRY_RUN = False


# ---------------------------------------------------------------------------
# Optional engine overrides
# ---------------------------------------------------------------------------
# This runner intentionally does not load pdf_v2_config.json.
# Add only local overrides that you explicitly want for this dataset.
# Usually empty dict is enough.
CFG: dict = {
    # Example:
    # "find_frame_border_band_left_mm": 20.0,
    # "find_frame_border_band_right_mm": 50.0,
    # "find_frame_border_band_top_mm": 20.0,
    # "find_frame_border_band_bottom_mm": 20.0,
}


def main() -> None:
    """Build config and run the paired deformation generator."""
    config = GenerationConfig(
        input_path=STAMPED_INPUT,
        background_input_path=BACKGROUND_INPUT,
        template_path=TEMPLATE_PATH,
        output_dir=OUTPUT_DIR,
        manifest_jsonl_path=MANIFEST_JSONL,
        manifest_csv_path=MANIFEST_CSV,
        page_num=PAGE_NUM,
        x_range=X_RANGE,
        y_range=Y_RANGE,
        step=STEP,
        mode=MODE,
        anchor=ANCHOR,
        dry_run=DRY_RUN,
        cfg=CFG,
    )
    summary = generate_deformed_stamp_dataset(config)

    print("Stamp deform bench runner")
    print(f"  stamped_input:      {STAMPED_INPUT}")
    print(f"  background_input:   {BACKGROUND_INPUT}")
    print(f"  template:           {TEMPLATE_PATH}")
    print(f"  output_dir:         {OUTPUT_DIR}")
    print(f"  manifest_jsonl:     {MANIFEST_JSONL}")
    print(f"  manifest_csv:       {MANIFEST_CSV}")
    print(f"  page_num:           {PAGE_NUM}")
    print(f"  x_range:            {X_RANGE[0]}..{X_RANGE[1]}")
    print(f"  y_range:            {Y_RANGE[0]}..{Y_RANGE[1]}")
    print(f"  step:               {STEP}")
    print(f"  mode:               {MODE}")
    print(f"  anchor:             {ANCHOR}")
    print(f"  dry_run:            {DRY_RUN}")
    print(f"  jobs:               {summary.jobs_count}")
    print(f"  variants_per_job:   {summary.variants_per_job}")
    print(f"  total_variants:     {summary.total_variants}")

    preview_rows = summary.manifest_rows[:5]
    if preview_rows:
        print("  preview:")
        for row in preview_rows:
            print(f"    - {row.output_pdf}")
            print(f"      {json.dumps(row.to_dict(), ensure_ascii=False)}")


if __name__ == "__main__":
    main()
