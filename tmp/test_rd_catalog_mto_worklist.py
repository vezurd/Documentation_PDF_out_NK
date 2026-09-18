"""Local checks for RD catalog MTO worklist rows and short status labels."""

from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from rd_catalog.models import CollisionKind
from rd_catalog.parse import issued_package_dir
from rd_catalog.kits import kit_identity_key
from rd_catalog.monitor_views import rd_mto_overlay_by_kit
from rd_catalog.pipeline import (
    MtoWorklistRow,
    ingest_google_snapshot,
    iter_mto_files_for_cells,
    list_mto_worklist,
    list_revision_matrix,
    rebuild_pipeline,
)
from rd_catalog.status_colors import status_color_label, status_short_label
from test_rd_catalog_pipeline import (
    _LOADED_AT,
    _event,
    _google,
    _mtime_ns,
    _open_db,
    _rebuild,
    _record,
    _send,
)


def _insert_collision(database, path_key: str, kind: CollisionKind) -> None:
    with database._connection() as connection, connection:
        connection.execute(
            """
            INSERT INTO scan_run(started_at, completed_at, status, is_baseline)
            VALUES ('2026-01-01T00:00:00+00:00', '2026-01-01T00:00:00+00:00',
                    'success', 1)
            """
        )
        run_id = int(connection.execute("SELECT last_insert_rowid()").fetchone()[0])
        connection.execute(
            """
            INSERT INTO current_collision(
                scope, source, kind, message, document_key,
                path_keys_json, scan_run_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "overlay_mto",
                "rd",
                kind.value,
                "fixture order conflict",
                "mto:2245-ksb|mto-0001",
                json.dumps([path_key], ensure_ascii=False),
                run_id,
            ),
        )


def _kit_rows(rows, title: str, mark: str):
    return [row for row in rows if row.title == title and row.mark == mark]


def test_status_short_label() -> None:
    assert status_short_label("code_a") == "код A"
    assert status_short_label("tdo_review") == "ТДО"
    assert status_short_label("problem") == status_color_label("problem")
    unknown = "not_a_real_status_key"
    assert status_short_label(unknown) == status_color_label(unknown) == unknown


def test_code_a_row_with_file(temp: Path) -> None:
    database = _open_db(temp / "code_a")
    folder = "01_рев.01_2245-KSB"
    record = _record(
        11,
        title="2245",
        mark="KSB",
        revision="01",
        appendix=None,
        sequence=1,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 8, 5),
        file_kind="mto_xlsx",
    )
    records = [record]
    _rebuild(
        database,
        (
            _google(
                "2245",
                "KSB",
                (
                    _event(
                        date="01.06.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="01",
                    ),
                ),
                revision="01",
                appendix=None,
            ),
        ),
        (
            _send(
                title="2245",
                mark="KSB",
                revision="01",
                appendix=None,
                status="Принят",
                send_date="20.05.2026",
                incoming="25.05.2026",
                row_index=1,
            ),
        ),
        records,
        {11},
    )
    rows = _kit_rows(list_mto_worklist(database, records=records), "2245", "KSB")
    assert len(rows) == 1
    row = rows[0]
    assert row.status == "code_a"
    assert row.revision_text == "01"
    assert row.mto_path == record.path
    assert row.gap_kind == ""
    assert row.package_path == issued_package_dir(record.path)
    assert row.package_path.endswith(folder)
    assert r"\PDF" not in row.package_path.upper()
    code_a_files = iter_mto_files_for_cells(
        database, records=records, predicate="code_a"
    )
    assert [item.path_key for item in code_a_files] == [record.path_key]


def test_no_mto_file(temp: Path) -> None:
    database = _open_db(temp / "no_mto")
    folder = "02_рев.01_1710-POS"
    record = _record(
        21,
        title="1710",
        mark="POS",
        revision="01",
        appendix=None,
        sequence=2,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 4, 1),
    )
    records = [record]
    _rebuild(
        database,
        (
            _google(
                "1710",
                "POS",
                (
                    _event(
                        date="10.04.2026",
                        stage="tdo_passed",
                        stage_label="прошла ТДО",
                        revision="01",
                    ),
                ),
                revision="01",
                appendix=None,
            ),
        ),
        (
            _send(
                title="1710",
                mark="POS",
                revision="01",
                appendix=None,
                status="Принят",
                send_date="01.04.2026",
                incoming="05.04.2026",
            ),
        ),
        records,
        {21},
    )
    rows = _kit_rows(list_mto_worklist(database, records=records), "1710", "POS")
    assert len(rows) == 1
    row = rows[0]
    assert row.gap_kind == "no_mto_file"
    assert row.mto_path == ""
    assert row.package_path.endswith(folder)
    assert row.package_count == 1


def test_no_package(temp: Path) -> None:
    database = _open_db(temp / "no_pkg")
    folder = "02_рев.01_AGCC.287-1715-SOT"
    record = _record(
        20,
        title="1715",
        mark="SOT",
        revision="01",
        appendix=None,
        sequence=2,
        folder=folder,
        mtime_ns=_mtime_ns(2026, 2, 1),
        file_kind="mto_xlsx",
    )
    records = [record]
    _rebuild(
        database,
        (_google("1715", "SOT", (), revision="03", appendix=None),),
        (
            _send(
                title="1715",
                mark="SOT",
                revision="03",
                appendix=None,
                status="На рассмотрении вх.контроля",
                send_date="01.04.2026",
                transmittal="AGCC-BCC-TRM-000030",
            ),
        ),
        records,
        {20},
    )
    rows = {
        row.revision_text: row
        for row in _kit_rows(
            list_mto_worklist(database, records=records), "1715", "SOT"
        )
    }
    assert "03" in rows
    grey = rows["03"]
    assert grey.gap_kind == "no_package"
    assert grey.mto_path == ""
    assert grey.package_path == ""
    assert grey.package_count == 0


def test_as_build_flag(temp: Path) -> None:
    database = _open_db(temp / "as_build")
    mtime = _mtime_ns(2026, 8, 5)
    rec_ifc = _record(
        12,
        title="2246",
        mark="KSB",
        revision="02",
        appendix=None,
        sequence=2,
        folder="02_рев.02_2246-KSB",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    rec_ab = _record(
        13,
        title="2246",
        mark="KSB",
        revision="03",
        appendix=None,
        sequence=8,
        folder="08_рев.03_2246-KSB_as-built",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
        as_build=True,
    )
    records = [rec_ifc, rec_ab]
    _rebuild(
        database,
        (
            _google(
                "2246",
                "KSB",
                (
                    _event(
                        date="10.07.2026",
                        stage="tdo_passed",
                        stage_label="прошла ТДО",
                        revision="02",
                    ),
                    _event(
                        date="01.08.2026",
                        stage="us_build",
                        stage_label="US-BUILD",
                        revision="03",
                    ),
                ),
                revision="03",
                appendix=None,
            ),
        ),
        (
            _send(
                title="2246",
                mark="KSB",
                revision="02",
                appendix=None,
                status="Принят",
                send_date="01.07.2026",
                incoming="05.07.2026",
                row_index=2,
            ),
        ),
        records,
        {13},
    )
    rows = {
        row.revision_text: row
        for row in _kit_rows(
            list_mto_worklist(database, records=records), "2246", "KSB"
        )
    }
    assert "02" in rows
    assert "03" not in rows
    ifc = rows["02"]
    assert ifc.is_as_build is False
    assert ifc.is_current_ifc is True
    heatmap = {
        cell.revision_text: cell
        for cell in list_revision_matrix(database)
        if cell.title == "2246" and cell.mark == "KSB"
    }
    as_build = heatmap["03"]
    assert as_build.is_as_build is True
    assert as_build.is_current is True
    assert as_build.is_current_ifc is False
    assert as_build.pipeline_status == "working"


def test_annulled_cell_skipped_from_worklist(temp: Path) -> None:
    database = _open_db(temp / "annulled_wl")
    mtime = _mtime_ns(2026, 8, 5)
    rec_ifc = _record(
        22,
        title="2247",
        mark="KSB",
        revision="02",
        appendix=None,
        sequence=2,
        folder="02_рев.02_2247-KSB",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    rec_old = _record(
        23,
        title="2247",
        mark="KSB",
        revision="03",
        appendix=None,
        sequence=8,
        folder="08_рев.03_2247-KSB",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    records = [rec_ifc, rec_old]
    _rebuild(
        database,
        (
            _google(
                "2247",
                "KSB",
                (
                    _event(
                        date="01.07.2026",
                        stage="tdo_passed",
                        stage_label="прошла ТДО",
                        revision="02",
                    ),
                ),
                revision="02",
                appendix=None,
            ),
        ),
        (
            _send(
                title="2247",
                mark="KSB",
                revision="02",
                appendix=None,
                status="Принят",
                send_date="01.07.2026",
                incoming="05.07.2026",
                row_index=2,
            ),
        ),
        records,
        {23},
    )
    database.upsert_annulled_flag(
        "2247", "KSB", "03", transfer_name="08_рев.03_2247-KSB", sequence=8
    )
    _rebuild(
        database,
        (
            _google(
                "2247",
                "KSB",
                (
                    _event(
                        date="01.07.2026",
                        stage="tdo_passed",
                        stage_label="прошла ТДО",
                        revision="02",
                    ),
                ),
                revision="02",
                appendix=None,
            ),
        ),
        (
            _send(
                title="2247",
                mark="KSB",
                revision="02",
                appendix=None,
                status="Принят",
                send_date="01.07.2026",
                incoming="05.07.2026",
                row_index=2,
            ),
        ),
        records,
        {23},
    )
    rows = {
        row.revision_text: row
        for row in _kit_rows(
            list_mto_worklist(database, records=records), "2247", "KSB"
        )
    }
    assert "02" in rows
    assert "03" not in rows
    heatmap = {
        cell.revision_text: cell
        for cell in list_revision_matrix(database)
        if cell.title == "2247" and cell.mark == "KSB"
    }
    assert heatmap["03"].pipeline_status == "annulled"


def test_package_count_and_newest_package(temp: Path) -> None:
    database = _open_db(temp / "count")
    mtime = _mtime_ns(2026, 5, 1)
    rec_low = _record(
        31,
        title="3310",
        mark="PD",
        revision="01",
        appendix=None,
        sequence=1,
        folder="01_рев.01_3310-PD",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    rec_high = _record(
        32,
        title="3310",
        mark="PD",
        revision="01",
        appendix=None,
        sequence=3,
        folder="03_рев.01_3310-PD",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    records = [rec_low, rec_high]
    _rebuild(
        database,
        (
            _google(
                "3310",
                "PD",
                (
                    _event(
                        date="01.05.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="01",
                    ),
                ),
                revision="01",
                appendix=None,
            ),
        ),
        (
            _send(
                title="3310",
                mark="PD",
                revision="01",
                appendix=None,
                status="Принят",
                send_date="20.04.2026",
                incoming="25.04.2026",
            ),
        ),
        records,
        {32},
    )
    rows = _kit_rows(list_mto_worklist(database, records=records), "3310", "PD")
    assert len(rows) == 1
    row = rows[0]
    assert row.package_count == 2
    assert row.package_path == issued_package_dir(rec_high.path)
    assert row.package_path.endswith("03_рев.01_3310-PD")
    assert row.mto_path == rec_high.path


def test_problem_kinds(temp: Path) -> None:
    database = _open_db(temp / "problems")
    record = _record(
        41,
        title="4410",
        mark="KSB",
        revision="01",
        appendix=None,
        sequence=1,
        folder="01_рев.01_4410-KSB",
        mtime_ns=_mtime_ns(2026, 3, 1),
        file_kind="mto_xlsx",
    )
    records = [record]
    _rebuild(
        database,
        (
            _google(
                "4410",
                "KSB",
                (
                    _event(
                        date="01.03.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="01",
                    ),
                ),
                revision="01",
                appendix=None,
            ),
        ),
        (
            _send(
                title="4410",
                mark="KSB",
                revision="01",
                appendix=None,
                status="Принят",
                send_date="20.02.2026",
                incoming="25.02.2026",
            ),
        ),
        records,
        {41},
    )
    _insert_collision(
        database, record.path_key, CollisionKind.TRANSFER_ORDER_CONFLICT
    )
    rebuild_pipeline(database, records=records, detected_current_ids={41})
    rows = _kit_rows(list_mto_worklist(database, records=records), "4410", "KSB")
    assert len(rows) == 1
    assert CollisionKind.TRANSFER_ORDER_CONFLICT.value in rows[0].problem_kinds


def test_worklist_order_matches_revision_matrix(temp: Path) -> None:
    database = _open_db(temp / "order")
    mtime = _mtime_ns(2026, 1, 1)
    rec_late = _record(
        51,
        title="3333",
        mark="AAA",
        revision="01",
        appendix=None,
        sequence=1,
        folder="01_рев.01_3333-AAA",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    rec_late_02 = _record(
        52,
        title="3333",
        mark="AAA",
        revision="02",
        appendix=None,
        sequence=2,
        folder="02_рев.02_3333-AAA",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    rec_early = _record(
        53,
        title="1111",
        mark="BBB",
        revision="02",
        appendix=None,
        sequence=1,
        folder="01_рев.02_1111-BBB",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    records = [rec_late, rec_late_02, rec_early]
    _rebuild(database, (), (), records, {51, 52, 53})
    worklist = list_mto_worklist(database, records=records)
    matrix = list_revision_matrix(database)
    assert [(row.title, row.mark, row.revision_text) for row in worklist] == [
        (cell.title, cell.mark, cell.revision_text) for cell in matrix
    ]
    assert [(row.title, row.mark, row.revision_text) for row in worklist] == [
        ("1111", "BBB", "02"),
        ("3333", "AAA", "01"),
        ("3333", "AAA", "02"),
    ]


def test_google_only_kit_no_rd(temp: Path) -> None:
    database = _open_db(temp / "google_only")
    ingest_google_snapshot(
        database,
        (_google("5555", "KSB", (), revision="02", appendix=None),),
        (),
        loaded_at=_LOADED_AT,
        source="test",
    )
    rows = _kit_rows(list_mto_worklist(database, records=[]), "5555", "KSB")
    assert len(rows) == 1
    row = rows[0]
    assert row.gap_kind == "no_rd"
    assert row.in_google is True
    assert row.in_issuance is False
    assert row.status == ""
    assert row.revision_text == "02"
    assert row.has_mto is False
    assert row.mto_path == ""
    assert row.package_count == 0
    cells = list_revision_matrix(database)
    assert _kit_rows(cells, "5555", "KSB") == []
    files = iter_mto_files_for_cells(
        database, records=[], predicate="code_a"
    )
    assert files == ()


def test_has_f_status_per_revision(temp: Path) -> None:
    database = _open_db(temp / "f_status")
    mtime = _mtime_ns(2026, 6, 1)
    rec_plain = _record(
        61,
        title="6610",
        mark="PD",
        revision="01",
        appendix=None,
        sequence=1,
        folder="01_рев.01_6610-PD",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    rec_f = _record(
        62,
        title="6610",
        mark="PD",
        revision="02",
        appendix=None,
        sequence=2,
        folder="02_рев.02_6610-PD",
        mtime_ns=mtime,
        file_kind="mto_xlsx",
    )
    records = [rec_plain, rec_f]
    _rebuild(
        database,
        (
            _google(
                "6610",
                "PD",
                (
                    _event(
                        date="01.06.2026",
                        stage="code_a",
                        stage_label="код А",
                        revision="02",
                    ),
                ),
                revision="02",
                appendix=None,
            ),
        ),
        (),
        records,
        {62},
    )
    rows = {
        row.revision_text: row
        for row in _kit_rows(
            list_mto_worklist(database, records=records), "6610", "PD"
        )
    }
    assert set(rows) >= {"01", "02"}
    assert rows["01"].has_f_status is False
    assert rows["01"].in_google is True
    assert rows["01"].in_issuance is False
    assert rows["01"].gap_kind == ""
    assert rows["02"].has_f_status is True
    assert rows["02"].in_google is True
    assert rows["02"].status == "code_a"


def test_worklist_pin_column() -> None:
    """Empty / current / stale pin cells; columns after the pin keep their data."""

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtGui import QColor
    from PySide6.QtWidgets import QApplication

    from rd_catalog.mto_export import ExportPin, export_pin_view
    from rd_catalog.mto_worklist_tab import (
        MtoWorklistTab,
        _COL_DATE,
        _COL_FILE,
        _COL_LINK,
        _COL_PACKAGE,
        _COL_PATH,
        _COL_PIN,
    )
    from rd_catalog.status_colors import color_for, default_palette

    app = QApplication.instance() or QApplication([])
    pin_name = "AGCC.287-9192-POS.MTO-0001_01_RU.xlsx"
    pin_path = rf"C:\rd\pack\{pin_name}"
    rule_path = r"C:\rd\pack\rule.xlsx"
    mto_path = r"C:\rd\unique-mto.xlsx"
    package_label = "01_рев.01"
    row = MtoWorklistRow(
        title="9192",
        mark="POS",
        revision_text="01",
        status="tdo_review",
        letters="T",
        is_as_build=False,
        is_current=False,
        is_current_ifc=False,
        has_mto=True,
        mto_path=mto_path,
        mto_mtime_ns=1_704_067_200_000_000_000,
        package_label=package_label,
        package_path=r"C:\rd\pack",
        package_count=1,
        has_f_status=True,
    )
    pin = ExportPin(
        title="9192",
        mark="POS",
        package_path=r"C:\rd\pack",
        file_path=pin_path,
        revision_text="01",
        evidence="abc",
        created_at="2026-09-10T12:00:00+00:00",
    )
    palette = default_palette()
    tab = MtoWorklistTab()
    tab.set_rows((row,), palette=palette, is_banned=lambda _t, _m: False)
    app.processEvents()
    table = tab.table()
    date_text = table.item(0, _COL_DATE).text()
    assert table.columnCount() == 16
    assert table.item(0, _COL_PIN).text() == ""
    assert table.item(0, _COL_LINK).text() == "F✓ · РД✓ · MTO✓"
    assert table.item(0, _COL_FILE).text() == "Да"
    assert table.item(0, _COL_PACKAGE).text() == package_label
    assert table.item(0, _COL_PATH).text() == mto_path

    def _lookup(origin: str):
        return lambda _t, _m: export_pin_view(
            pin, origin=origin, rule_path=rule_path
        )

    tab.set_rows(
        (row,),
        palette=palette,
        is_banned=lambda _t, _m: False,
        pin_view_for=_lookup("pin"),
    )
    app.processEvents()
    current = table.item(0, _COL_PIN)
    assert current.text() == pin_name
    assert r"C:\rd\pack" in current.toolTip()
    assert pin_path in current.toolTip()
    assert "2026-09-10T12:00:00+00:00" in current.toolTip()
    assert table.item(0, _COL_FILE).text() == "Да"
    assert table.item(0, _COL_DATE).text() == date_text
    assert table.item(0, _COL_PACKAGE).text() == package_label
    assert table.item(0, _COL_PATH).text() == mto_path

    tab.set_rows(
        (row,),
        palette=palette,
        is_banned=lambda _t, _m: False,
        pin_view_for=_lookup("pin_stale"),
    )
    app.processEvents()
    stale = table.item(0, _COL_PIN)
    assert stale.text() == f"{pin_name} (устарел)"
    assert pin_path in stale.toolTip()
    assert rule_path in stale.toolTip()
    stale_hex = QColor(
        color_for(palette, "export_pin_stale")
    ).name().casefold()
    assert stale.background().color().name().casefold() == stale_hex
    assert table.item(0, _COL_FILE).text() == "Да"
    assert table.item(0, _COL_DATE).text() == date_text
    assert table.item(0, _COL_PACKAGE).text() == package_label
    assert table.item(0, _COL_PATH).text() == mto_path
    tab.deleteLater()
    app.processEvents()


def test_worklist_revision_link_and_gap() -> None:
    """Headers and gap styling distinguish journal status from RD files."""

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtGui import QColor
    from PySide6.QtWidgets import QApplication

    from rd_catalog.mto_worklist_tab import (
        MtoWorklistTab,
        _COL_DATE,
        _COL_FILE,
        _COL_GAP,
        _COL_GOOGLE,
        _COL_LINK,
        _COL_PACKAGE,
        _COL_PATH,
        _COL_PIN,
        _COL_PROBLEMS,
        _COL_STATUS,
    )
    from rd_catalog.status_colors import color_for, default_palette

    app = QApplication.instance() or QApplication([])
    no_package = MtoWorklistRow(
        title="1715",
        mark="SOT",
        revision_text="03",
        status="code_a",
        letters="A",
        is_as_build=False,
        is_current=True,
        is_current_ifc=True,
        has_mto=False,
        gap_kind="no_package",
        has_f_status=True,
    )
    with_file = MtoWorklistRow(
        title="9192",
        mark="POS",
        revision_text="01",
        status="code_a",
        letters="A",
        is_as_build=False,
        is_current=False,
        is_current_ifc=True,
        has_mto=True,
        mto_path=r"C:\rd\pack\AGCC.287-9192-POS.MTO-0001_01_RU.xlsx",
        package_label="01_рев.01",
        package_path=r"C:\rd\pack",
        package_count=1,
        has_f_status=True,
    )
    palette = default_palette()
    tab = MtoWorklistTab()
    tab.set_rows(
        (no_package, with_file),
        palette=palette,
        is_banned=lambda _t, _m: False,
    )
    app.processEvents()
    table = tab.table()
    headers = [
        table.horizontalHeaderItem(column).text()
        for column in range(table.columnCount())
    ]
    assert headers == [
        "Титул",
        "Марка",
        "Ревизия (F/РД)",
        "Этап ревизии",
        "Метки",
        "Связь F → РД → MTO",
        "Google",
        "AB",
        "Ручной выбор MTO",
        "Сверка Авто МТО",
        "Файл MTO",
        "Дата MTO",
        "Пакет",
        "Чего не хватает",
        "Проблемы",
        "Путь MTO",
    ]
    assert (
        _COL_STATUS,
        _COL_LINK,
        _COL_GOOGLE,
        _COL_PIN,
        _COL_FILE,
        _COL_DATE,
        _COL_PACKAGE,
        _COL_GAP,
        _COL_PROBLEMS,
        _COL_PATH,
    ) == (3, 5, 6, 8, 10, 11, 12, 13, 14, 15)
    no_package_index = next(
        index
        for index in range(table.rowCount())
        if table.item(index, 0).text() == "1715"
    )
    with_file_index = next(
        index
        for index in range(table.rowCount())
        if table.item(index, 0).text() == "9192"
    )

    problem_hex = QColor(color_for(palette, "problem")).name().casefold()
    code_a_hex = QColor(color_for(palette, "code_a")).name().casefold()
    link = table.item(no_package_index, _COL_LINK)
    assert link.text() == "F✓ · РД✗ · MTO✗"
    assert "Google F / «Выдача РД ПД»" in link.toolTip()
    assert "не доказывает наличие папки" in link.toolTip()
    assert link.foreground().color().name().casefold() == problem_hex
    assert table.item(no_package_index, _COL_FILE).text() == "Нет"
    for column in (_COL_FILE, _COL_PACKAGE, _COL_GAP):
        assert (
            table.item(no_package_index, column)
            .foreground()
            .color()
            .name()
            .casefold()
            == problem_hex
        )
    for column in range(table.columnCount()):
        assert table.item(no_package_index, column).font().bold()
        assert table.item(no_package_index, column).font().underline()
    status = table.item(no_package_index, _COL_STATUS)
    assert status.background().color().name().casefold() == code_a_hex
    assert "Разрыв: нет папки." in status.toolTip()
    assert "ревизией имени файла «03»" in status.toolTip()

    with_file_link = table.item(with_file_index, _COL_LINK)
    assert with_file_link.text() == "F✓ · РД✓ · MTO✓"
    assert not with_file_link.font().bold()
    assert r"C:\rd\pack" in with_file_link.toolTip()

    tab._filter.setText("F✓ · РД✗ · MTO✗")
    app.processEvents()
    assert table.isRowHidden(no_package_index) is False
    assert table.isRowHidden(with_file_index) is True
    tab._filter.clear()
    tab._copy_visible_tsv()
    tsv_lines = app.clipboard().text().splitlines()
    assert tsv_lines[0].split("\t") == headers
    assert all(len(line.split("\t")) == 16 for line in tsv_lines)

    tab.deleteLater()
    app.processEvents()


def test_official_folder_uses_lagged_mto_filename(temp: Path) -> None:
    """6100-SOT: official OD 04 package still holds MTO 03; older NN is ignored."""

    database = _open_db(temp / "lagged_mto")
    older = _record(
        31,
        title="6100",
        mark="SOT",
        revision="03",
        appendix=None,
        sequence=11,
        folder="11_рев.03_AGCC.287-6100-SOT",
        mtime_ns=_mtime_ns(2024, 10, 1),
        file_kind="mto_xlsx",
    )
    older_od = _record(
        32,
        title="6100",
        mark="SOT",
        revision="03",
        appendix=None,
        sequence=11,
        folder="11_рев.03_AGCC.287-6100-SOT",
        mtime_ns=_mtime_ns(2024, 10, 1),
    )
    official_mto = _record(
        41,
        title="6100",
        mark="SOT",
        revision="03",
        appendix=None,
        sequence=14,
        folder="14_рев.04_AGCC.287-6100-SOT",
        mtime_ns=_mtime_ns(2024, 11, 19),
        file_kind="mto_xlsx",
    )
    official_od = _record(
        42,
        title="6100",
        mark="SOT",
        revision="04",
        appendix=None,
        sequence=14,
        folder="14_рев.04_AGCC.287-6100-SOT",
        mtime_ns=_mtime_ns(2024, 11, 19),
    )
    working_mto = _record(
        51,
        title="6100",
        mark="SOT",
        revision="03",
        appendix="01",
        sequence=15,
        folder="15_рев.AN01_AGCC.287-6100-SOT",
        mtime_ns=_mtime_ns(2025, 2, 1),
        file_kind="mto_xlsx",
    )
    working_od = _record(
        52,
        title="6100",
        mark="SOT",
        revision="04",
        appendix="01",
        sequence=15,
        folder="15_рев.AN01_AGCC.287-6100-SOT",
        mtime_ns=_mtime_ns(2025, 2, 1),
    )
    records = [
        older,
        older_od,
        official_mto,
        official_od,
        working_mto,
        working_od,
    ]
    _rebuild(
        database,
        (
            _google(
                "6100",
                "SOT",
                (
                    _event(
                        date="24.12.2024",
                        stage="code_a",
                        stage_label="код А",
                        revision="04",
                    ),
                ),
                revision="04",
                appendix=None,
            ),
        ),
        (
            _send(
                title="6100",
                mark="SOT",
                revision="04",
                appendix=None,
                status="Принят",
                send_date="19.11.2024",
                incoming="20.11.2024",
            ),
        ),
        records,
        {42},
    )
    rows = _kit_rows(list_mto_worklist(database, records=records), "6100", "SOT")
    official = next(row for row in rows if row.revision_text == "04")
    assert official.mto_path == official_mto.path
    assert official.gap_kind == ""
    assert official.package_path.endswith("14_рев.04_AGCC.287-6100-SOT")
    pipelines = {
        kit_identity_key(row.title, row.mark): row
        for row in database.list_kit_pipelines()
    }
    overlay = rd_mto_overlay_by_kit(rows, pipelines)
    path, rev = overlay[kit_identity_key("6100", "SOT")]
    assert path == official_mto.path
    assert rev == "03"


def main() -> None:
    """Run MTO worklist fixtures against a temporary SQLite file."""

    test_status_short_label()
    with tempfile.TemporaryDirectory(prefix="rd_catalog_mto_worklist_") as raw:
        root = Path(raw)
        test_code_a_row_with_file(root)
        test_no_mto_file(root)
        test_no_package(root)
        test_as_build_flag(root)
        test_annulled_cell_skipped_from_worklist(root)
        test_package_count_and_newest_package(root)
        test_problem_kinds(root)
        test_worklist_order_matches_revision_matrix(root)
        test_google_only_kit_no_rd(root)
        test_has_f_status_per_revision(root)
        test_official_folder_uses_lagged_mto_filename(root)
    test_worklist_pin_column()
    test_worklist_revision_link_and_gap()
    print("RD catalog MTO worklist: OK")


if __name__ == "__main__":
    main()
