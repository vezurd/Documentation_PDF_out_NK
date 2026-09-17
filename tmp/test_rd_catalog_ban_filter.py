"""Local checks for the RD catalog title+mark ban filter."""

from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.ban_filter import (
    BanFilterStore,
    parse_title_mark,
    title_mark_from_row,
)


def main() -> None:
    """Verify parse, persist, identity, and reload without Qt or UNC."""

    assert parse_title_mark("5850", "SKUD") == ("5850", "SKUD")
    assert parse_title_mark(" 5850-SKUD ", "") == ("5850", "SKUD")
    assert parse_title_mark("5850\u2010SKUD", "") == ("5850", "SKUD")
    try:
        parse_title_mark("5850", "")
    except ValueError:
        pass
    else:
        raise AssertionError("empty mark must raise")

    assert title_mark_from_row("5850", "SKUD") == ("5850", "SKUD")
    assert title_mark_from_row(None, None, "6400-PD") == ("6400", "PD")
    assert title_mark_from_row("6400", None, "6400-PD") == ("6400", "PD")

    with tempfile.TemporaryDirectory(prefix="rd_catalog_ban_") as temp:
        runtime = Path(temp) / "runtime"
        store = BanFilterStore.from_runtime_dir(runtime)
        assert store.pairs() == ()
        assert not store.contains("5850", "SKUD")
        loaded_revision = store.revision

        pair, added = store.add("5850", "SKUD", "аннулирован")
        assert added
        assert store.revision > loaded_revision
        after_add = store.revision
        assert pair.label == "5850-SKUD"
        assert store.contains("5850", "skud")
        again, added_again = store.add("5850", "SKUD")
        assert not added_again
        assert store.revision == after_add
        assert again.identity == pair.identity
        assert len(store.pairs()) == 1

        path = runtime / "banned_title_marks.json"
        payload = json.loads(path.read_text(encoding="utf-8"))
        assert payload["version"] == 1
        assert payload["pairs"][0]["title"] == "5850"
        assert payload["pairs"][0]["mark"] == "SKUD"
        assert payload["pairs"][0]["comment"] == "аннулирован"

        reloaded = BanFilterStore.from_runtime_dir(runtime)
        assert reloaded.contains("5850", "SKUD")
        assert reloaded.pairs()[0].comment == "аннулирован"

        assert store.remove("5850", "SKUD")
        assert store.revision > after_add
        assert not store.contains("5850", "SKUD")
        assert store.pairs() == ()
        assert not store.remove("5850", "SKUD")

        store.add("6400-PD", "")
        assert store.contains("6400", "PD")

        bad = runtime / "banned_title_marks.json"
        bad.write_text("{not json", encoding="utf-8")
        broken = BanFilterStore(bad)
        assert broken.pairs() == ()
        assert broken.load_error

    print("RD catalog ban filter: OK")


if __name__ == "__main__":
    main()
