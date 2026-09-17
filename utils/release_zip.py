import argparse
import datetime
import fnmatch
import os
from pathlib import Path, PurePosixPath
from typing import Iterable
from zipfile import ZIP_DEFLATED, ZipFile


DEFAULT_EXCLUDE_PATTERNS: tuple[str, ...] = (
    ".git/",
    ".cursor/",
    ".idea/",
    ".vscode/",
    "__pycache__/",
    ".pytest_cache/",
    ".mypy_cache/",
    # Directory basename: any folder named "cache" (e.g. RFQ/ds_compare/cache).
    "cache/",
    "debug_output/",
    "grid_diagnostic_output/",
    "tmp/",
    "tmp_step4_compare/",
    "tmp_step4_compare_v2/",
    "_release/",
    "release/",
    # Sample PDFs / regression fixtures (large).
    "pdf_parsing_v2_engine/templates/test_pdf/",
    # Editor / pipeline debug dumps under test trees.
    "adapt_debug/",
    "*.pyc",
    "*.pyo",
    "*.pyd",
    "*.cache",
    "*.log",
)


def _normalize_rel_path(path: str) -> str:
    normalized = path.replace("\\", "/").strip("/")
    return "" if normalized == "." else normalized


def _load_exclude_patterns(ignore_file: Path) -> list[str]:
    patterns = list(DEFAULT_EXCLUDE_PATTERNS)
    if not ignore_file.exists():
        return patterns

    with ignore_file.open("r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            patterns.append(line)
    return patterns


def _match_pattern(rel_path: str, pattern: str, is_dir: bool) -> bool:
    rel_path = _normalize_rel_path(rel_path)
    pattern = pattern.replace("\\", "/").strip()
    if not pattern:
        return False

    dir_only = pattern.endswith("/")
    pattern = _normalize_rel_path(pattern.rstrip("/"))
    rel_obj = PurePosixPath(rel_path)

    if dir_only and not is_dir:
        return False

    # Pattern for directory prefix, e.g. "cache/" or "RFQ/tmp/".
    if dir_only:
        if "/" in pattern:
            return rel_path == pattern or rel_path.startswith(f"{pattern}/")
        # Single path component: match that folder name at any depth
        # (top-level "cache/" is both "cache" and endswith "/cache" — redundant but clear).
        return rel_path == pattern or rel_path.endswith(f"/{pattern}")

    # Pattern with path components.
    if "/" in pattern:
        return fnmatch.fnmatch(rel_path, pattern) or fnmatch.fnmatch(rel_path, f"**/{pattern}")

    # Basename-only pattern, e.g. "*.pyc".
    return fnmatch.fnmatch(rel_obj.name, pattern)


def _should_skip(rel_path: str, is_dir: bool, patterns: Iterable[str]) -> bool:
    for pattern in patterns:
        if _match_pattern(rel_path, pattern, is_dir):
            return True
    return False


def build_release_zip(
    project_root: str | Path,
    output_dir: str | Path | None = None,
    ignore_file_name: str = ".release_zipignore",
) -> tuple[str, int, int]:
    """
    Build a distributable ZIP archive for colleagues.

    Returns:
        (zip_path, included_files_count, skipped_files_count)
    """
    project_root = Path(project_root).resolve()
    if output_dir is None:
        output_dir = project_root / "_release"
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    ignore_file = project_root / ignore_file_name
    exclude_patterns = _load_exclude_patterns(ignore_file)

    date_prefix = datetime.datetime.now().strftime("%Y.%m.%d")
    zip_name = f"{date_prefix}_{project_root.name}.zip"
    zip_path = output_dir / zip_name

    files_to_add: list[tuple[Path, str]] = []
    skipped_files = 0

    for root, dirs, files in os.walk(project_root):
        root_path = Path(root)
        rel_root = _normalize_rel_path(str(root_path.relative_to(project_root)))

        filtered_dirs: list[str] = []
        for d in dirs:
            rel_dir = _normalize_rel_path(f"{rel_root}/{d}" if rel_root else d)
            if _should_skip(rel_dir, is_dir=True, patterns=exclude_patterns):
                continue
            filtered_dirs.append(d)
        dirs[:] = filtered_dirs

        for file_name in files:
            abs_file = root_path / file_name
            rel_file = _normalize_rel_path(str(abs_file.relative_to(project_root)))

            if abs_file == zip_path:
                continue
            if _should_skip(rel_file, is_dir=False, patterns=exclude_patterns):
                skipped_files += 1
                continue

            files_to_add.append((abs_file, rel_file))

    with ZipFile(zip_path, mode="w", compression=ZIP_DEFLATED, compresslevel=6) as zf:
        for abs_file, rel_file in files_to_add:
            zf.write(abs_file, arcname=rel_file)

    return str(zip_path), len(files_to_add), skipped_files


def _main() -> int:
    parser = argparse.ArgumentParser(description="Build project release ZIP without local cache artifacts.")
    parser.add_argument("--project-root", default=".", help="Project root folder")
    parser.add_argument("--output-dir", default=None, help="Folder where ZIP file will be created")
    args = parser.parse_args()

    zip_path, included, skipped = build_release_zip(
        project_root=args.project_root,
        output_dir=args.output_dir,
    )
    print(f"ZIP создан: {zip_path}")
    print(f"Включено файлов: {included}")
    print(f"Пропущено файлов: {skipped}")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
