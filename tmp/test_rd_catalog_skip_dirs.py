"""Local checks for skip-dir matching, store, config merge, and DB prune."""

from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.config import CatalogConfig, load_config
from rd_catalog.db import CatalogDatabase
from rd_catalog.models import ReviewState, SourceKind
from rd_catalog.scan import _is_skipped_dir, scan_catalog
from rd_catalog.skip_dirs import (
    SkipDirsStore,
    is_skipped_dir_name,
    load_runtime_skip_dirs,
    path_has_skipped_dir,
    packaged_skip_dirs,
    validate_skip_token,
)


def main() -> None:
    """Run skip-dir unit checks without Qt."""

    assert is_skipped_dir_name("замечания", ["Замечания"])
    assert is_skipped_dir_name("пример", packaged_skip_dirs())
    assert _is_skipped_dir("замечания", ("Замечания",))
    remarks_pdf = (
        r"\\bcc\eng\PrDoc\РД\8950\12_POS1\замечания\01"
        r"\AGCC.287-8950-POS1.OD-0001_0_RU_НК.pdf"
    )
    assert path_has_skipped_dir(remarks_pdf, ["Замечания"])
    assert not path_has_skipped_dir(
        r"C:\rd\Для передачи\01\AGCC.287-8950-POS1.замечания.pdf",
        ["Замечания"],
    )
    try:
        validate_skip_token("передач", [])
        raise AssertionError("gate token must be rejected")
    except ValueError as exc:
        assert "передачи" in str(exc).casefold() or "передач" in str(exc)
    try:
        validate_skip_token("отправк", [])
        raise AssertionError("gate token must be rejected")
    except ValueError:
        pass
    try:
        validate_skip_token("old", ["OLD"])
        raise AssertionError("duplicate must be rejected")
    except ValueError:
        pass
    assert validate_skip_token("  nanoCad  ", []) == "nanoCad"

    with tempfile.TemporaryDirectory(prefix="rd_skip_store_") as temp:
        runtime = Path(temp, "runtime")
        store = SkipDirsStore.from_runtime_dir(runtime, packaged=("factory",))
        assert store.tokens() == ("factory",)
        assert store.load_error is None
        assert load_runtime_skip_dirs(runtime) == ("factory",)
        store.add("custom")
        store.save()
        loaded = load_runtime_skip_dirs(runtime)
        assert loaded == ("factory", "custom")
        again = SkipDirsStore.from_runtime_dir(runtime, packaged=("ignored",))
        assert again.tokens() == ("factory", "custom")
        again.sort_casefold()
        assert again.tokens() == ("custom", "factory")
        again.save()

        broken = Path(temp, "broken")
        (broken / "skip_dirs.json").parent.mkdir(parents=True)
        (broken / "skip_dirs.json").write_text("{not json", encoding="utf-8")
        healed = SkipDirsStore.from_runtime_dir(broken, packaged=("healed",))
        assert healed.tokens() == ("healed",)
        assert healed.load_error
        assert load_runtime_skip_dirs(broken) == ("healed",)

        env_root = Path(temp, "localapp")
        env_root.mkdir()
        first = load_config(environment={"LOCALAPPDATA": str(env_root)})
        assert "Замечания" in first.skip_dirs
        assert not (first.runtime_dir / "skip_dirs.json").exists()
        user_runtime = first.runtime_dir
        user_runtime.mkdir(parents=True, exist_ok=True)
        (user_runtime / "skip_dirs.json").write_text(
            json.dumps({"version": 1, "tokens": ["user-only"]}, ensure_ascii=False),
            encoding="utf-8",
        )
        second = load_config(environment={"LOCALAPPDATA": str(env_root)})
        assert second.skip_dirs == ("user-only",)

    with tempfile.TemporaryDirectory(prefix="rd_skip_db_") as temp:
        rd = Path(temp, "rd")
        kept_dir = rd / "8950" / "12_POS1" / "Для передачи" / "01"
        skip_dir = rd / "8950" / "замечания" / "Для передачи" / "01"
        kept_dir.mkdir(parents=True)
        skip_dir.mkdir(parents=True)
        kept = kept_dir / "AGCC.287-8950-POS1.OD-0001_0_RU.pdf"
        skipped = skip_dir / "AGCC.287-8950-POS1.OD-0002_0_RU.pdf"
        kept.write_bytes(b"kept")
        skipped.write_bytes(b"skip-me")
        sq = Path(temp, "sq")
        sq.mkdir()
        robot = Path(temp, "robot")
        robot.mkdir()
        db_path = Path(temp, "runtime", "catalog.sqlite")
        config = CatalogConfig(
            rd_root=rd,
            sq_root=sq,
            robot_root=robot,
            runtime_dir=db_path.parent,
            db_path=db_path,
            robot_flat_structure=True,
            skip_dirs=(),
        )
        database = CatalogDatabase(config.db_path)
        database.initialize()
        summary = scan_catalog(config, sources=(SourceKind.RD,))
        assert {Path(item.path).name for item in summary.files} == {
            kept.name,
            skipped.name,
        }
        database.store_scan(summary)
        before = {record.path: record for record in database.list_files()}
        assert before[str(kept)].present
        assert before[str(skipped)].present
        skipped_state = before[str(skipped)].review_state
        assert database.count_present_skipped(("Замечания",)) == 1
        removed = database.apply_skip_dirs(("Замечания",))
        assert removed == 1
        after = {record.path: record for record in database.list_files()}
        assert after[str(kept)].present
        assert not after[str(skipped)].present
        assert after[str(skipped)].review_state is skipped_state
        assert after[str(skipped)].review_state is ReviewState.ACKNOWLEDGED
        overlay_paths = {row["path"] for row in database.current_overlay()}
        assert str(kept) in overlay_paths
        assert str(skipped) not in overlay_paths
        assert database.count_present_skipped(("Замечания",)) == 0

    print("RD catalog skip dirs: OK")


if __name__ == "__main__":
    main()
