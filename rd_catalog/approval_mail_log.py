"""Append-only ingest journal for «Письма о согласовании».

Qt-free. File is ``runtime_dir/approval_mail_ingest.log``. Does not write
Google or UNC sources.
"""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path

INGEST_LOG_NAME = "approval_mail_ingest.log"
_MAX_BYTES = 8 * 1024 * 1024


def ingest_log_path(runtime_dir: str | Path) -> Path:
    """Return the ingest log path under *runtime_dir*.

    Args:
        runtime_dir: Catalog runtime folder.

    Returns:
        ``runtime_dir/approval_mail_ingest.log``.
    """

    return Path(runtime_dir) / INGEST_LOG_NAME


def append_approval_mail_ingest(
    runtime_dir: str | Path | None,
    blocks: Sequence[str],
) -> None:
    """Append letter blocks to the ingest log.

    Missing ``runtime_dir``, empty *blocks*, or an OS error are ignored so
    ingest never fails because of the log.

    Args:
        runtime_dir: Catalog runtime folder, or ``None`` to skip.
        blocks: Already formatted UTF-8 text chunks (one letter each).
    """

    if runtime_dir is None or not blocks:
        return
    path = ingest_log_path(runtime_dir)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        _rotate_if_huge(path)
        with path.open("a", encoding="utf-8") as handle:
            for block in blocks:
                handle.write(block.rstrip() + "\n\n")
    except OSError:
        return


def _rotate_if_huge(path: Path) -> None:
    try:
        if not path.is_file() or path.stat().st_size < _MAX_BYTES:
            return
        backup = path.with_name(path.name + ".old")
        backup.unlink(missing_ok=True)
        path.replace(backup)
    except OSError:
        return
