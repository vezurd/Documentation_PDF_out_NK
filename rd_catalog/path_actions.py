"""Windows and UNC path actions that never read file contents."""

from __future__ import annotations

import os
import subprocess
from collections.abc import Iterable
from pathlib import Path


def normalize_path(path: str | None) -> str:
    """Normalize mixed Windows separators while preserving UNC prefixes."""

    text = str(path or "").strip()
    return os.path.normpath(text) if text else ""


def path_is_under(path: str | Path | None, root: str | Path | None) -> bool:
    """Return whether ``path`` is ``root`` or a descendant.

    Args:
        path: Candidate file or folder.
        root: Directory that must contain ``path``.

    Returns:
        True when the normalized, case-folded path sits under ``root``.
    """

    path_key = os.path.normcase(os.path.normpath(str(path or ""))).casefold()
    root_key = os.path.normcase(os.path.normpath(str(root or ""))).casefold()
    if not path_key or not root_key:
        return False
    if path_key == root_key:
        return True
    prefix = root_key.rstrip("\\/") + os.sep
    return path_key.startswith(prefix)


def common_parent_dir(
    paths: Iterable[str],
    *,
    under_root: str | Path | None = None,
) -> str:
    """Return the longest common parent of ``paths``, optionally under a root.

    Args:
        paths: File or directory paths (local or UNC).
        under_root: When set, the result must sit under this directory.

    Returns:
        Normalized common directory, or empty string when paths do not share
        a parent (or do not sit under ``under_root``).
    """

    normalized = [normalize_path(path) for path in paths if str(path or "").strip()]
    if not normalized:
        return ""
    try:
        common = os.path.commonpath(normalized)
    except ValueError:
        return ""
    common = normalize_path(common)
    if under_root is not None:
        root = normalize_path(str(under_root))
        if not root or not path_is_under(common, root):
            return ""
    return common


def containing_folder(path: str | None) -> str:
    """Return the lexical parent folder without touching the filesystem."""

    normalized = normalize_path(path)
    return os.path.dirname(normalized) if normalized else ""


def _open_with_explorer(path: str) -> bool:
    """Open *path* with ``explorer.exe`` when ShellExecute failed.

    The path is its own argument. ``/root,`` is an Explorer switch: a space
    in ``Для передачи`` makes Python quote the whole switch, and Explorer
    then opens Documents.

    Args:
        path: Normalized local or UNC path.

    Returns:
        True when Explorer was launched.
    """

    if os.name != "nt" or not path:
        return False
    try:
        subprocess.Popen(["explorer.exe", path])
    except OSError:
        return False
    return True


def open_directory(path: str | None) -> tuple[bool, str]:
    """Open a folder without probing the UNC source.

    ShellExecute (``os.startfile``) is the primary call. It does not parse
    ``/`` switches, so a space in the issued folder name stays part of the
    path. ``explorer.exe`` with the bare path is only the fallback.

    Args:
        path: Directory path (already a folder, not a file).

    Returns:
        ``(True, path)`` when the folder was handed to Windows, otherwise
        an error.
    """

    normalized = normalize_path(path)
    if not normalized:
        return False, "Путь не указан."
    try:
        os.startfile(normalized)  # type: ignore[attr-defined]
    except OSError as exc:
        if _open_with_explorer(normalized):
            return True, normalized
        return False, f"Не удалось открыть {normalized}: {exc}"
    return True, normalized


def open_path(path: str | None) -> tuple[bool, str]:
    """Ask Windows to open a path without reading its contents."""

    normalized = normalize_path(path)
    if not normalized:
        return False, "Путь не указан."
    try:
        os.startfile(normalized)  # type: ignore[attr-defined]
    except OSError as exc:
        if _open_with_explorer(normalized):
            return True, normalized
        return False, f"Не удалось открыть {normalized}: {exc}"
    return True, normalized


def open_containing_folder(path: str | None) -> tuple[bool, str]:
    """Open a path's containing folder without probing the UNC source."""

    folder = containing_folder(path)
    if not folder:
        return False, "Не удалось определить содержащую папку."
    return open_directory(folder)
