"""Side-effect-free configuration loading for the RD catalog."""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping

from rd_catalog.skip_dirs import load_runtime_skip_dirs


DEFAULT_CONFIG_PATH = Path(__file__).with_name("default_config.json")


@dataclass(frozen=True, slots=True)
class CatalogConfig:
    """Resolved RD catalog configuration."""

    rd_root: Path
    sq_root: Path
    robot_root: Path
    runtime_dir: Path
    db_path: Path
    robot_flat_structure: bool
    skip_dirs: tuple[str, ...]
    google_kits_spreadsheet_id: str = "1iKF-5tf0LAewib-5_E0R7AurGKChk4-QPtbgI1DN4zk"
    google_kits_sheet_name: str = ""
    google_issuance_spreadsheet_id: str = "1NMRZLWdAYmVlWyExtC9PYcWS7cr3TuQzMwu59HiSYEU"
    google_issuance_sheet_name: str = "Выдача РД ПД"
    an_root: Path = Path("")


def _expand_windows_vars(value: str, environment: Mapping[str, str]) -> str:
    def replace(match: re.Match[str]) -> str:
        key = match.group(1)
        return environment.get(key, environment.get(key.upper(), match.group(0)))

    return os.path.expandvars(re.sub(r"%([^%]+)%", replace, value))


def _read_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as stream:
        data = json.load(stream)
    if not isinstance(data, dict):
        raise ValueError(f"Configuration root must be an object: {path}")
    return data


def load_config(
    user_override_path: str | Path | None = None,
    *,
    environment: Mapping[str, str] | None = None,
) -> CatalogConfig:
    """Load packaged defaults and merge an optional user override.

    Importing this module and calling this function never creates directories,
    configuration files, or the production database.

    Args:
        user_override_path: Optional JSON file whose keys replace defaults.
        environment: Environment mapping used to expand ``%NAME%`` variables.

    Returns:
        A validated, resolved configuration object.

    Raises:
        FileNotFoundError: If an explicitly supplied override does not exist.
        ValueError: If JSON shape or required values are invalid.
    """

    data = _read_json(DEFAULT_CONFIG_PATH)
    if user_override_path is not None:
        data.update(_read_json(Path(user_override_path)))

    required = {
        "rd_root",
        "sq_root",
        "robot_root",
        "runtime_dir",
        "db_path",
        "robot_flat_structure",
        "skip_dirs",
    }
    missing = sorted(required - data.keys())
    if missing:
        raise ValueError(f"Missing required configuration keys: {', '.join(missing)}")

    env = os.environ if environment is None else environment

    def resolved_path(key: str) -> Path:
        value = data[key]
        if not isinstance(value, str) or not value.strip():
            raise ValueError(f"Configuration value {key!r} must be a non-empty string")
        return Path(_expand_windows_vars(value, env))

    skip_dirs = data["skip_dirs"]
    if not isinstance(skip_dirs, list) or not all(
        isinstance(item, str) and item for item in skip_dirs
    ):
        raise ValueError("Configuration value 'skip_dirs' must be a string list")
    if not isinstance(data["robot_flat_structure"], bool):
        raise ValueError("Configuration value 'robot_flat_structure' must be boolean")

    runtime_dir = resolved_path("runtime_dir")
    runtime_skips = load_runtime_skip_dirs(runtime_dir)
    if runtime_skips is not None:
        skip_dirs = list(runtime_skips)

    spreadsheet_id = str(
        env.get("RD_KITS_SPREADSHEET_ID")
        or data.get("google_kits_spreadsheet_id")
        or "1iKF-5tf0LAewib-5_E0R7AurGKChk4-QPtbgI1DN4zk"
    ).strip()
    sheet_name = str(
        env.get("RD_KITS_SHEET_NAME") or data.get("google_kits_sheet_name") or ""
    ).strip()
    issuance_id = str(
        env.get("RD_ISSUANCE_SPREADSHEET_ID")
        or data.get("google_issuance_spreadsheet_id")
        or "1NMRZLWdAYmVlWyExtC9PYcWS7cr3TuQzMwu59HiSYEU"
    ).strip()
    issuance_sheet = str(
        env.get("RD_ISSUANCE_SHEET_NAME")
        or data.get("google_issuance_sheet_name")
        or "Выдача РД ПД"
    ).strip()
    raw_an_root = data.get("an_root", "")
    if isinstance(raw_an_root, str) and raw_an_root.strip():
        an_root = Path(_expand_windows_vars(raw_an_root, env))
    else:
        an_root = Path("")

    return CatalogConfig(
        rd_root=resolved_path("rd_root"),
        sq_root=resolved_path("sq_root"),
        robot_root=resolved_path("robot_root"),
        runtime_dir=resolved_path("runtime_dir"),
        db_path=resolved_path("db_path"),
        robot_flat_structure=data["robot_flat_structure"],
        skip_dirs=tuple(skip_dirs),
        google_kits_spreadsheet_id=spreadsheet_id,
        google_kits_sheet_name=sheet_name,
        google_issuance_spreadsheet_id=issuance_id,
        google_issuance_sheet_name=issuance_sheet,
        an_root=an_root,
    )
