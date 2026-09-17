"""CLI for generating paired deformed-stamp PDF fixtures."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from pdf_parsing_v2.v2_config import load_v2_config
from pdf_parsing_v2_engine.stamp_deform_generator import (
    DEFAULT_ANCHOR,
    DEFAULT_BACKGROUND_INPUT_DIR,
    DEFAULT_MANIFEST_CSV,
    DEFAULT_MANIFEST_JSONL,
    DEFAULT_MODE,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_PAGE_NUM,
    DEFAULT_RANGE_MAX,
    DEFAULT_RANGE_MIN,
    DEFAULT_STEP,
    DEFAULT_STAMPED_INPUT_DIR,
    DEFAULT_TEMPLATE_PATH,
    GenerationConfig,
    generate_deformed_stamp_dataset,
)


def build_arg_parser() -> argparse.ArgumentParser:
    """Build the command-line parser."""
    parser = argparse.ArgumentParser(
        description="Generate paired deformed-stamp PDF fixtures for v2 benchmarks.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "--input",
        default=str(DEFAULT_STAMPED_INPUT_DIR),
        help="Stamped source PDF or directory with stamped PDFs.",
    )
    parser.add_argument(
        "--background-input",
        default=str(DEFAULT_BACKGROUND_INPUT_DIR),
        help="Background PDF or directory with matching background PDFs.",
    )
    parser.add_argument(
        "--template",
        default=str(DEFAULT_TEMPLATE_PATH),
        help="Path to the v2 template JSON used to compute stamp bbox.",
    )
    parser.add_argument(
        "--page-num",
        type=int,
        default=DEFAULT_PAGE_NUM,
        help="1-based page number containing the stamp in the source PDF.",
    )
    parser.add_argument(
        "--x-range",
        nargs=2,
        type=int,
        metavar=("MIN", "MAX"),
        default=(DEFAULT_RANGE_MIN, DEFAULT_RANGE_MAX),
        help="Deformation delta percent range for X scaling.",
    )
    parser.add_argument(
        "--y-range",
        nargs=2,
        type=int,
        metavar=("MIN", "MAX"),
        default=(DEFAULT_RANGE_MIN, DEFAULT_RANGE_MAX),
        help="Deformation delta percent range for Y scaling.",
    )
    parser.add_argument(
        "--step",
        type=int,
        default=DEFAULT_STEP,
        help="Step in deformation delta percent.",
    )
    parser.add_argument(
        "--mode",
        choices=["grid", "diagonal", "axis-x", "axis-y", "cross"],
        default=DEFAULT_MODE,
        help="How to expand the deformation combinations.",
    )
    parser.add_argument(
        "--anchor",
        choices=["bottom_right", "bottom_left", "top_right", "top_left", "center"],
        default=DEFAULT_ANCHOR,
        help="Anchor used when scaling the stamp bbox.",
    )
    parser.add_argument(
        "--out-dir",
        default=str(DEFAULT_OUTPUT_DIR),
        help="Directory for generated PDFs.",
    )
    parser.add_argument(
        "--manifest-jsonl",
        default=str(DEFAULT_MANIFEST_JSONL),
        help="Path to JSONL manifest.",
    )
    parser.add_argument(
        "--manifest-csv",
        default=str(DEFAULT_MANIFEST_CSV),
        help="Path to CSV manifest.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Only inspect inputs and planned outputs without writing PDFs.",
    )
    return parser


def main() -> None:
    """Parse CLI args and execute the generator."""
    parser = build_arg_parser()
    args = parser.parse_args()

    config = GenerationConfig(
        input_path=Path(args.input),
        background_input_path=Path(args.background_input),
        template_path=Path(args.template),
        output_dir=Path(args.out_dir),
        manifest_jsonl_path=Path(args.manifest_jsonl),
        manifest_csv_path=Path(args.manifest_csv),
        page_num=int(args.page_num),
        x_range=(int(args.x_range[0]), int(args.x_range[1])),
        y_range=(int(args.y_range[0]), int(args.y_range[1])),
        step=int(args.step),
        mode=str(args.mode),
        anchor=str(args.anchor),
        dry_run=bool(args.dry_run),
        cfg=load_v2_config(),
    )

    try:
        summary = generate_deformed_stamp_dataset(config)
    except ValueError as exc:
        raise SystemExit(str(exc)) from exc

    print("Paired stamp deformation dataset")
    print(f"  input:              {config.input_path}")
    print(f"  background_input:   {config.background_input_path}")
    print(f"  template:           {config.template_path}")
    print(f"  out_dir:            {config.output_dir}")
    print(f"  page_num:           {config.page_num}")
    print(f"  mode:               {config.mode}")
    print(f"  x_range:            {config.x_range[0]}..{config.x_range[1]}")
    print(f"  y_range:            {config.y_range[0]}..{config.y_range[1]}")
    print(f"  step:               {config.step}")
    print(f"  jobs:               {summary.jobs_count}")
    print(f"  variants_per_job:   {summary.variants_per_job}")
    print(f"  total_variants:     {summary.total_variants}")

    if summary.dry_run:
        print("  mode_status:        dry_run")
        preview_rows = summary.manifest_rows[:5]
        if preview_rows:
            print("  preview:")
            for row in preview_rows:
                print(f"    - {row.output_pdf}")
                print(f"      {json.dumps(row.to_dict(), ensure_ascii=False)}")
        return

    print("  mode_status:        generated")
    print(f"  manifest_jsonl:     {config.manifest_jsonl_path}")
    print(f"  manifest_csv:       {config.manifest_csv_path}")


if __name__ == "__main__":
    main()
