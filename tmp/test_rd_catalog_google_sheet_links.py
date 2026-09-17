"""Local checks for Google Sheets cell deep links (no network)."""

from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.config import CatalogConfig
from rd_catalog.google_sheet_links import (
    GOOGLE_HREF_TIP,
    JOURNAL_ISSUANCE_COLUMNS,
    KITS_GOOGLE_COLUMNS,
    KITS_ISSUANCE_COLUMNS,
    KITS_SHEET_TITLE_FALLBACK,
    SheetLinkContext,
    a1_cell,
    google_sheet_cell_url,
    is_google_sheets_url,
    journal_cell_href,
    kits_google_hrefs,
    save_kits_sheet_pin,
    sheet_link_context_from_config,
)
from rd_catalog.kits import GoogleKit, IssuanceKit, parse_history_line, parse_sheet_revision


def _links(**overrides: object) -> SheetLinkContext:
    payload = dict(
        kits_spreadsheet_id="kits-id",
        kits_sheet_title="Контроль выдачи ",
        kits_sheet_id=77,
        issuance_spreadsheet_id="iss-id",
        issuance_sheet_title="Выдача РД ПД",
        issuance_sheet_id=None,
    )
    payload.update(overrides)
    return SheetLinkContext(**payload)  # type: ignore[arg-type]


def _google(row_index: int = 12) -> GoogleKit:
    events = (parse_history_line("04.12.2025 код А на рев. 01 AGCC-BCC-TRM-1"),)
    revision, appendix = parse_sheet_revision("01")
    return GoogleKit(
        title="2210",
        mark="KSB",
        mark_raw="KSB",
        title_system="2210-KSB",
        sheet_revision=revision,
        sheet_appendix=appendix,
        sheet_revision_text="01",
        status_sheet="ok",
        comment_raw=events[0].raw,
        events=events,
        last_event=events[0],
        row_index=row_index,
    )


def _issuance(row_index: int = 40) -> IssuanceKit:
    revision, appendix = parse_sheet_revision("01-AN01")
    return IssuanceKit(
        title="2210",
        mark="KSB",
        mark_raw="KSB",
        title_system="2210-KSB",
        revision=revision,
        appendix=appendix,
        revision_text="01-AN01",
        status="отправлено",
        send_date="01.12.2025",
        send_date_sortable="2025-12-01",
        send_transmittal="AGCC-BCC-TRM-0001",
        incoming_control_date="",
        incoming_control_date_sortable="",
        confirm_transmittal="AGCC-BCC-TRM-0002",
        note_raw="AGCC-BCC-TRM-0002",
        row_index=row_index,
    )


def main() -> None:
    """URL builder, pin cache, and Комплекты/journal column maps."""

    assert a1_cell("f", 12) == "F12"
    assert a1_cell("Q", 0) == ""
    assert a1_cell("", 3) == ""

    with_gid = google_sheet_cell_url(
        "abc", "F12", sheet_id=77, sheet_title="ignored"
    )
    assert with_gid == (
        "https://docs.google.com/spreadsheets/d/abc/edit?gid=77#gid=77&range=F12"
    ), with_gid
    named = google_sheet_cell_url(
        "abc", "B40", sheet_title="Выдача РД ПД"
    )
    assert "range=" in named
    assert "Выдача" in named or "%D0%92%D1%8B%D0%B4%D0%B0%D1%87%D0%B0" in named
    assert is_google_sheets_url(with_gid)
    assert not is_google_sheets_url("https://example.com/")
    assert google_sheet_cell_url("", "F12") == ""

    links = _links()
    hrefs = kits_google_hrefs(google=_google(), issuance=_issuance(), links=links)
    assert hrefs["Google · TRM F"].endswith("range=F12")
    assert "gid=77" in hrefs["Google · TRM F"]
    assert "B40" in hrefs["Выдача · TRM отпр."]
    assert "Q40" in hrefs["Выдача · TRM подтв."]
    for header in KITS_GOOGLE_COLUMNS:
        assert header in hrefs, header
    for header in KITS_ISSUANCE_COLUMNS:
        assert header in hrefs, header
    empty = kits_google_hrefs(google=None, issuance=None, links=links)
    assert empty == {}
    no_row = kits_google_hrefs(google=_google(0), issuance=_issuance(0), links=links)
    assert no_row == {}

    trm = journal_cell_href("TRM", 40, links)
    assert "B40" in trm
    assert journal_cell_href("Титул", 40, links) == ""
    assert journal_cell_href("TRM", 0, links) == ""
    assert set(JOURNAL_ISSUANCE_COLUMNS) >= {"TRM", "TRM подтв."}
    assert GOOGLE_HREF_TIP

    with tempfile.TemporaryDirectory(prefix="rd_sheet_pins_") as raw:
        root = Path(raw)
        assert save_kits_sheet_pin(
            root,
            spreadsheet_id="kits-id",
            sheet_id=9,
            title="Контроль выдачи ",
        )
        config = CatalogConfig(
            rd_root=root,
            sq_root=root,
            robot_root=root,
            runtime_dir=root,
            db_path=root / "x.sqlite",
            robot_flat_structure=False,
            skip_dirs=(),
            google_kits_spreadsheet_id="kits-id",
            google_kits_sheet_name="",
            google_issuance_spreadsheet_id="iss-id",
            google_issuance_sheet_name="Выдача РД ПД",
        )
        ctx = sheet_link_context_from_config(config)
        assert ctx.kits_sheet_id == 9
        assert ctx.kits_sheet_title == "Контроль выдачи "
        assert ctx.issuance_sheet_title == "Выдача РД ПД"
        payload = json.loads((root / "google_sheet_pins.json").read_text(encoding="utf-8"))
        assert payload["kits"]["sheet_id"] == 9

        other = CatalogConfig(
            rd_root=root,
            sq_root=root,
            robot_root=root,
            runtime_dir=root,
            db_path=root / "x.sqlite",
            robot_flat_structure=False,
            skip_dirs=(),
            google_kits_spreadsheet_id="other-id",
            google_kits_sheet_name="",
        )
        mismatched = sheet_link_context_from_config(other)
        assert mismatched.kits_sheet_id is None
        assert mismatched.kits_sheet_title == KITS_SHEET_TITLE_FALLBACK

    print("RD catalog Google sheet links: OK")


if __name__ == "__main__":
    main()
