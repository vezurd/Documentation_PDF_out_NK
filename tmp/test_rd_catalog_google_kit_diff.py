"""Local checks for KSB ИД F/D/E kit-diff used after letter writes."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.kits import (
    GoogleKit,
    changed_google_kit_keys,
    google_kit_sheet_fingerprint,
    kit_identity_key,
)


def _kit(
    title: str,
    mark: str,
    *,
    comment: str = "F",
    revision: str = "Рев. 01",
    status: str = "ok",
) -> GoogleKit:
    return GoogleKit(
        title=title,
        mark=mark,
        mark_raw=mark,
        title_system=f"{title}-{mark}",
        sheet_revision="01",
        sheet_appendix=None,
        sheet_revision_text=revision,
        status_sheet=status,
        comment_raw=comment,
        events=(),
        last_event=None,
        row_index=2,
    )


def main() -> None:
    """Fingerprint and set-diff of Google kit F/D/E text."""

    first = _kit("4110", "KSB", comment="01.01.2026 код А")
    same = _kit("4110", "KSB", comment="01.01.2026 код А")
    assert google_kit_sheet_fingerprint(first) == google_kit_sheet_fingerprint(same)
    assert changed_google_kit_keys((first,), (same,)) == set()

    updated = _kit("4110", "KSB", comment="01.01.2026 код А\n15.09.2026 код А")
    assert changed_google_kit_keys((first,), (updated,)) == {
        kit_identity_key("4110", "KSB")
    }

    de_changed = _kit("4110", "KSB", comment="01.01.2026 код А", status="РД Согласовано")
    assert changed_google_kit_keys((first,), (de_changed,)) == {
        kit_identity_key("4110", "KSB")
    }

    other = _kit("1600", "SOT")
    assert changed_google_kit_keys((first,), (first, other)) == {
        kit_identity_key("1600", "SOT")
    }
    assert changed_google_kit_keys((first, other), (first,)) == {
        kit_identity_key("1600", "SOT")
    }
    assert changed_google_kit_keys((), (first,)) == {kit_identity_key("4110", "KSB")}


if __name__ == "__main__":
    main()
    print("ok")
