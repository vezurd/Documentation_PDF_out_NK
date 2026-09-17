"""Offscreen tests for the «Ревизии MTO» export cockpit."""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path
from dataclasses import replace
from types import SimpleNamespace

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QApplication, QHeaderView, QMenu, QPlainTextEdit

from rd_catalog.config import CatalogConfig
from rd_catalog.db import CatalogDatabase
from rd_catalog.mto_export import (
    DEFAULT_ROBOT_TARGET_NAME,
    ExportPin,
    ExportPinCandidate,
    ExportSelection,
    ExportTarget,
    export_pin_evidence_for_kit,
    list_export_pin_candidates,
    load_export_pins,
    save_export_pins,
    save_export_targets,
)
from rd_catalog.mto_export_copy import (
    EXPORT_STATE_COLOR_KEY,
    EXPORT_STATE_TEXT,
    copy_selections,
    group_export_preview,
    plan_for_visible_rows,
    preview_copy_count,
)
from rd_catalog.mto_export_dialog import MtoExportPreviewDialog
from rd_catalog.mto_pair_compare import PairPoolStatus
from rd_catalog.customer_pi_auto_mto import (
    AUTO_MTO_COMPARE_STATUS_HEADER,
    AutoMtoCompareResult,
    AutoMtoFile,
)
from rd_catalog.revision_matrix_tab import (
    MtoPinPickDialog,
    RevisionMatrixTab,
    _COL_AUTO_MTO,
    _COL_AUTO_MTO_COMPARE,
    _COL_EXPORT,
    _COL_MARK,
    _COL_RD_REV,
    _COL_REV_FIRST,
    _COL_TITLE,
    _PIN_DIALOG_HEADERS,
    _PIN_RULE_MARK,
    _REV_MATCH_BG,
    _ROLE_CELL,
    _ROLE_EXPORT,
    _ROLE_EXPORT_COLOR,
)
from rd_catalog.status_colors import color_for, default_palette, status_color_label
from test_rd_catalog_mto_export import RD_ROOT, _MARK, _TITLE, _load_3240

import pytest


@pytest.fixture(scope="module")
def app() -> QApplication:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    instance = QApplication.instance()
    if instance is None:
        instance = QApplication([])
    return instance


def _cell(
    title: str,
    mark: str,
    revision_text: str,
    *,
    status: str = "code_a",
    letters: str = "A",
) -> SimpleNamespace:
    return SimpleNamespace(
        title=title,
        mark=mark,
        revision_text=revision_text,
        pipeline_status=status,
        letters=letters,
        problem_kinds_json="[]",
        is_as_build=False,
        is_current=revision_text == "02",
        is_current_ifc=revision_text == "01",
        has_mto=True,
    )


def _selection(
    title: str,
    mark: str,
    state: str,
    *,
    source_path: str = "",
    destination_path: str = "",
    existing_target_path: str = "",
    package_path: str = r"C:\rd\pack",
    origin: str = "rule",
    rule_path: str = "",
    warnings: tuple[str, ...] = (),
) -> ExportSelection:
    if state != "no_source" and not source_path:
        source_path = rf"C:\rd\{title}-{mark}.xlsx"
    if (
        state in {"add", "replace", "same_data", "unknown", "pin_stale"}
        and not destination_path
    ):
        destination_path = rf"C:\robot\{title}-{mark}.xlsx"
    if state in {"replace", "same_data", "unknown"} and not existing_target_path:
        existing_target_path = destination_path
    if state == "pin_stale":
        origin = "pin_stale"
        if not rule_path:
            rule_path = rf"C:\rd\{title}-{mark}_rule.xlsx"
        if not source_path or source_path.endswith(f"{title}-{mark}.xlsx"):
            source_path = rf"C:\rd\{title}-{mark}_pin.xlsx"
    return ExportSelection(
        title=title,
        mark=mark,
        rule="latest_issued",
        source_path=source_path,
        source_revision_text="01",
        package_path=package_path,
        origin=origin,
        state=state,
        destination_path=destination_path,
        existing_target_path=existing_target_path,
        confidence="high",
        warnings=warnings,
        rule_path=rule_path,
    )


def _config(root: Path) -> CatalogConfig:
    return CatalogConfig(
        rd_root=root / "rd",
        sq_root=root / "sq",
        robot_root=root / "robot",
        runtime_dir=root / "runtime",
        db_path=root / "runtime" / "catalog.sqlite",
        robot_flat_structure=True,
        skip_dirs=(),
    )


def _hex(item) -> str:
    return item.background().color().name().casefold()


def test_six_states_and_revision_shift(app: QApplication) -> None:
    palette = default_palette()
    for key in EXPORT_STATE_COLOR_KEY.values():
        assert key in palette
        assert status_color_label(key)
    tab = RevisionMatrixTab()
    states = (
        ("1111", "KSB", "add"),
        ("2222", "POS", "replace"),
        ("3333", "SOT", "same_data"),
        ("4444", "PD", "no_source"),
        ("5555", "SKUD", "unknown"),
        ("6666", "KSB1", "pin_stale"),
    )
    cells = []
    for title, mark, _state in states:
        cells.append(_cell(title, mark, "01", letters="A"))
        cells.append(_cell(title, mark, "02", letters="T", status="tdo_review"))
    tab.set_cells(
        cells,
        palette=palette,
        is_banned=lambda _t, _m: False,
    )
    app.processEvents()
    table = tab.table()
    assert table.columnCount() == 8
    headers = [table.horizontalHeaderItem(index).text() for index in range(8)]
    assert headers[:6] == [
        "Титул",
        "Марка",
        "В папку робота",
        "РД · рев.",
        "Авто МТО",
        AUTO_MTO_COMPARE_STATUS_HEADER,
    ]
    assert headers[6:] == ["01", "02"]
    header = table.horizontalHeader()
    assert header.sectionResizeMode(_COL_TITLE) == QHeaderView.ResizeMode.Interactive
    assert header.sectionResizeMode(_COL_MARK) == QHeaderView.ResizeMode.Interactive
    assert header.sectionResizeMode(_COL_EXPORT) == QHeaderView.ResizeMode.Interactive
    assert header.sectionResizeMode(_COL_RD_REV) == QHeaderView.ResizeMode.Interactive
    assert header.sectionResizeMode(_COL_AUTO_MTO) == QHeaderView.ResizeMode.Interactive
    assert header.sectionResizeMode(_COL_AUTO_MTO_COMPARE) == (
        QHeaderView.ResizeMode.Interactive
    )
    assert header.sectionResizeMode(_COL_REV_FIRST) == QHeaderView.ResizeMode.Stretch
    assert header.sectionResizeMode(_COL_REV_FIRST + 1) == QHeaderView.ResizeMode.Stretch
    assert (
        table.horizontalScrollBarPolicy()
        == Qt.ScrollBarPolicy.ScrollBarAlwaysOff
    )

    selections = [_selection(title, mark, state) for title, mark, state in states]
    tab.apply_export_selections(selections)
    app.processEvents()
    assert table.rowCount() == 6
    for row, (title, mark, state) in enumerate(states):
        export_item = table.item(row, _COL_EXPORT)
        assert export_item is not None
        assert export_item.text() == EXPORT_STATE_TEXT[state]
        color_key = export_item.data(_ROLE_EXPORT_COLOR)
        assert color_key == EXPORT_STATE_COLOR_KEY[state]
        expected = QColor(color_for(palette, color_key)).name().casefold()
        assert _hex(export_item) == expected
        stored = export_item.data(_ROLE_EXPORT)
        assert isinstance(stored, ExportSelection)
        assert stored.state == state
        rev_01 = table.item(row, _COL_REV_FIRST)
        rev_02 = table.item(row, _COL_REV_FIRST + 1)
        assert rev_01 is not None and rev_02 is not None
        cell_01 = rev_01.data(_ROLE_CELL)
        cell_02 = rev_02.data(_ROLE_CELL)
        assert cell_01 is not None
        assert cell_02 is not None
        assert cell_01.revision_text == "01"
        assert cell_02.revision_text == "02"
        assert rev_01.text() == "A"
        assert rev_02.text() == "T"
        assert table.item(row, _COL_TITLE).text() == title
        assert table.item(row, _COL_MARK).text() == mark
        if state == "no_source":
            missing = QColor(color_for(palette, "export_missing")).name().casefold()
            assert _hex(table.item(row, _COL_TITLE)) == missing
            assert _hex(table.item(row, _COL_MARK)) == missing
        if state == "pin_stale":
            tip = export_item.toolTip()
            assert "закреплённый файл" in tip
            assert "файл по правилу" in tip
            assert stored.rule_path in tip
            assert stored.source_path in tip
    tab.deleteLater()
    app.processEvents()


def test_kit_jump_only_from_context_action(app: QApplication) -> None:
    tab = RevisionMatrixTab()
    tab.set_cells(
        [_cell("1111", "KSB", "01")],
        palette=default_palette(),
        is_banned=lambda _t, _m: False,
    )
    emitted: list[tuple[str, str]] = []
    tab.kit_activated.connect(
        lambda title, mark: emitted.append((title, mark))
    )
    table = tab.table()
    table.cellClicked.emit(0, _COL_TITLE)
    table.cellDoubleClicked.emit(0, _COL_TITLE)
    app.processEvents()
    assert emitted == []

    menu = QMenu(tab)
    action = tab._add_show_kit_action(menu, "1111", "KSB")
    assert menu.actions()[0] is action
    assert action.text() == 'Показать в „Комплекты“'
    action.trigger()
    app.processEvents()
    assert emitted == [("1111", "KSB")]
    tab.deleteLater()
    app.processEvents()


def test_apply_auto_mto_compare_paints_composite(app: QApplication) -> None:
    tab = RevisionMatrixTab()
    tab.set_auto_mto_queue_paused(True)
    tab.set_cells(
        [_cell("1111", "KSB", "01"), _cell("1111", "KSB", "02")],
        palette=default_palette(),
        is_banned=lambda _t, _m: False,
    )
    rd_path = r"C:\rd\AGCC.287-1111-KSB.MTO-0001_03_RU.xlsx"
    tab.apply_export_selections(
        [_selection("1111", "KSB", "add", source_path=rd_path)]
    )
    files = (
        AutoMtoFile(
            title="1111",
            mark="KSB",
            spec="AGCC.287-1111-KSB.MTO-0001",
            rd_revision="02",
            relpath="1111/KSB/AGCC.287-1111-KSB.MTO-0001_02_RU.xlsx",
            row_count=1,
            fingerprint="mto",
            source_specs=("AGCC.287-1111-KSB.MTO-0001",),
        ),
        AutoMtoFile(
            title="1111",
            mark="KSB",
            spec="AGCC.287-1111-KSB.MTO-0001",
            rd_revision="0",
            relpath="1111/KSB/AGCC.287-1111-KSB.MTO-0001_0_RU.xlsx",
            row_count=1,
            fingerprint="ds",
            source_specs=("AGCC.287-1111-KSB.DS-0001",),
        ),
    )
    tab.set_auto_mto_index({("1111", "ksb"): files})
    result = AutoMtoCompareResult(
        match_kind="composite",
        members=files,
        rd_rows=2,
        auto_rows=2,
        combinations_checked=1,
    )
    tab.apply_auto_mto_compare_result("1111", "KSB", rd_path, result)
    app.processEvents()
    item = tab.table().item(0, _COL_AUTO_MTO)
    assert item is not None
    assert item.text() == "02 (2)"
    compare = tab.table().item(0, _COL_AUTO_MTO_COMPARE)
    assert compare is not None
    assert compare.text() == "четкое · сумма 02 + 0"
    assert _hex(compare) == _REV_MATCH_BG.name().casefold()
    tooltip = compare.toolTip()
    assert "02 + 0" in tooltip
    assert "суммы 2 файлов" in tooltip
    cached = tab.comparison_for("1111", "KSB", rd_path)
    assert cached is result
    tab.deleteLater()
    app.processEvents()


def test_auto_mto_compare_paints_rev_match(app: QApplication) -> None:
    tab = RevisionMatrixTab()
    tab.set_auto_mto_queue_paused(True)
    tab.set_cells(
        [_cell("1111", "KSB", "02")],
        palette=default_palette(),
        is_banned=lambda _t, _m: False,
    )
    rd_path = r"C:\rd\AGCC.287-1111-KSB.MTO-0001_02_RU.xlsx"
    selection = _selection("1111", "KSB", "add", source_path=rd_path)
    selection = replace(selection, source_revision_text="02")
    tab.apply_export_selections([selection])
    tab.set_auto_mto_index(
        {
            ("1111", "ksb"): (
                AutoMtoFile(
                    title="1111",
                    mark="KSB",
                    spec="AGCC.287-1111-KSB.MTO-0001",
                    rd_revision="02",
                    relpath="1111/KSB/AGCC.287-1111-KSB.MTO-0001_02_RU.xlsx",
                    row_count=1,
                    fingerprint="mto",
                ),
            )
        }
    )
    app.processEvents()
    compare = tab.table().item(0, _COL_AUTO_MTO_COMPARE)
    assert compare is not None
    assert compare.text() == "рев. совпала"
    assert _hex(compare) == _REV_MATCH_BG.name().casefold()
    tab.deleteLater()
    app.processEvents()


def _auto_file(title: str, mark: str, revision: str = "01") -> AutoMtoFile:
    return AutoMtoFile(
        title=title,
        mark=mark,
        spec=f"AGCC.287-{title}-{mark}.MTO-0001",
        rd_revision=revision,
        relpath=f"{title}/{mark}/AGCC.287-{title}-{mark}.MTO-0001_{revision}_RU.xlsx",
        row_count=1,
        fingerprint=f"{title}-{mark}-{revision}",
    )


def test_heatmap_sorts_numeric_titles_and_marks(app: QApplication) -> None:
    tab = RevisionMatrixTab()
    tab.set_cells(
        [_cell("9110", "POS", "01"), _cell("1757", "KSB", "01")],
        palette=default_palette(),
        is_banned=lambda _t, _m: False,
    )
    table = tab.table()
    table.setSortingEnabled(True)
    table.sortByColumn(_COL_TITLE, Qt.SortOrder.AscendingOrder)
    app.processEvents()
    assert table.item(0, _COL_TITLE).text() == "1757"
    table.sortByColumn(_COL_MARK, Qt.SortOrder.AscendingOrder)
    app.processEvents()
    assert table.item(0, _COL_MARK).text() == "KSB"
    tab.deleteLater()
    app.processEvents()


def test_auto_mto_enqueue_skips_cached_and_queues_when_running(
    app: QApplication,
) -> None:
    tab = RevisionMatrixTab()
    tab.set_auto_mto_queue_paused(True)
    tab.set_cells(
        [
            _cell("1111", "KSB", "01"),
            _cell("2222", "POS", "01"),
            _cell("3333", "SOT", "01"),
        ],
        palette=default_palette(),
        is_banned=lambda _t, _m: False,
    )
    files_a = (_auto_file("1111", "KSB"),)
    files_b = (_auto_file("2222", "POS"),)
    files_c = (_auto_file("3333", "SOT"),)
    tab.set_auto_mto_index(
        {
            ("1111", "ksb"): files_a,
            ("2222", "pos"): files_b,
            ("3333", "sot"): files_c,
        }
    )
    rd_a = r"C:\rd\1111-KSB.xlsx"
    rd_b = r"C:\rd\2222-POS.xlsx"
    rd_c = r"C:\rd\3333-SOT.xlsx"
    result = AutoMtoCompareResult(
        match_kind="no_match",
        members=(),
        rd_rows=1,
        auto_rows=0,
        combinations_checked=0,
    )
    tab.apply_auto_mto_compare_result("1111", "KSB", rd_a, result)
    tab.set_auto_mto_index(
        {
            ("1111", "ksb"): files_a,
            ("2222", "pos"): files_b,
            ("3333", "sot"): files_c,
        }
    )
    assert tab.comparison_for("1111", "KSB", rd_a) is result
    tab.set_auto_mto_queue_paused(True)
    tab.enqueue_auto_mto_compares(
        [
            ("1111", "KSB", rd_a),
            ("2222", "POS", rd_b),
        ]
    )
    assert tab._auto_mto_queue == [(("2222", "pos"), rd_b)]
    tab._auto_mto_compare_thread = SimpleNamespace(isRunning=lambda: True)
    started = tab.start_auto_mto_compare("3333", "SOT", rd_c)
    assert started is True
    assert tab._auto_mto_queue[0] == (("3333", "sot"), rd_c)
    assert (("2222", "pos"), rd_b) in tab._auto_mto_queue
    tab._auto_mto_compare_thread = None
    tab._auto_mto_queue.clear()
    tab.prepare_close()
    tab.deleteLater()
    app.processEvents()


def test_stale_export_path_enqueues_current_file(app: QApplication) -> None:
    tab = RevisionMatrixTab()
    tab.set_auto_mto_queue_paused(True)
    tab.set_cells(
        [_cell("1111", "KSB", "01")],
        palette=default_palette(),
        is_banned=lambda _t, _m: False,
    )
    files = (_auto_file("1111", "KSB"),)
    old_path = r"C:\rd\old\1111-KSB.xlsx"
    new_path = r"C:\rd\new\1111-KSB.xlsx"
    tab.set_auto_mto_index({("1111", "ksb"): files})
    tab.apply_auto_mto_compare_result(
        "1111",
        "KSB",
        old_path,
        AutoMtoCompareResult(
            match_kind="no_match",
            members=(),
            rd_rows=1,
            auto_rows=0,
            combinations_checked=0,
        ),
    )
    tab._auto_mto_queue.clear()
    tab.apply_export_selections(
        [_selection("1111", "KSB", "add", source_path=new_path)]
    )
    compare = tab.table().item(0, _COL_AUTO_MTO_COMPARE)
    assert compare is not None
    assert compare.text() == "другой файл"
    assert tab._auto_mto_queue == [(("1111", "ksb"), new_path)]
    tab.prepare_close()
    tab.deleteLater()
    app.processEvents()


def test_same_rd_path_dash_variant_does_not_enqueue(app: QApplication) -> None:
    tab = RevisionMatrixTab()
    tab.set_auto_mto_queue_paused(True)
    tab.set_cells(
        [_cell("1111", "KSB", "01")],
        palette=default_palette(),
        is_banned=lambda _t, _m: False,
    )
    files = (_auto_file("1111", "KSB"),)
    rd_path = r"C:\rd\AGCC.287-1111-KSB.MTO-0001_01_RU.xlsx"
    dashed = rd_path.replace("-1111-", "\u20101111\u2010")
    tab.set_auto_mto_index({("1111", "ksb"): files})
    tab.apply_auto_mto_compare_result(
        "1111",
        "KSB",
        dashed,
        AutoMtoCompareResult(
            match_kind="no_match",
            members=(),
            rd_rows=1,
            auto_rows=0,
            combinations_checked=0,
        ),
    )
    tab._auto_mto_queue.clear()
    tab.apply_export_selections(
        [_selection("1111", "KSB", "add", source_path=rd_path)]
    )
    compare = tab.table().item(0, _COL_AUTO_MTO_COMPARE)
    assert compare is not None
    assert compare.text() == "не совпало"
    assert tab._auto_mto_queue == []
    tab.prepare_close()
    tab.deleteLater()
    app.processEvents()


def test_export_path_change_replaces_queued_old_file(app: QApplication) -> None:
    tab = RevisionMatrixTab()
    tab.set_auto_mto_queue_paused(True)
    tab.set_cells(
        [_cell("1111", "KSB", "01")],
        palette=default_palette(),
        is_banned=lambda _t, _m: False,
    )
    files = (_auto_file("1111", "KSB"),)
    old_path = r"C:\rd\old\1111-KSB.xlsx"
    new_path = r"C:\rd\new\1111-KSB.xlsx"
    tab.set_auto_mto_index({("1111", "ksb"): files})
    tab.enqueue_auto_mto_compares([("1111", "KSB", old_path)])
    assert tab._auto_mto_queue == [(("1111", "ksb"), old_path)]
    tab.apply_export_selections(
        [_selection("1111", "KSB", "add", source_path=new_path)]
    )
    assert tab._auto_mto_queue == [(("1111", "ksb"), new_path)]
    tab.apply_auto_mto_compare_result(
        "1111",
        "KSB",
        old_path,
        AutoMtoCompareResult(
            match_kind="no_match",
            members=(),
            rd_rows=1,
            auto_rows=0,
            combinations_checked=0,
        ),
    )
    assert tab.table().item(0, _COL_AUTO_MTO_COMPARE).text() == "другой файл"
    assert tab._auto_mto_queue == [(("1111", "ksb"), new_path)]
    tab.prepare_close()
    tab.deleteLater()
    app.processEvents()


def test_export_gate(app: QApplication) -> None:
    tab = RevisionMatrixTab()
    palette = default_palette()
    tab.set_cells(
        [_cell("1111", "KSB", "01"), _cell("1111", "KSB", "02")],
        palette=palette,
        is_banned=lambda _t, _m: False,
    )
    tab.apply_export_selections([_selection("1111", "KSB", "add")])
    tab.set_pool_status(PairPoolStatus(total=10, compared=4, pending=6, failed=0))
    app.processEvents()
    assert not tab._export_button.isEnabled()
    assert "не завершено" in tab._export_button.toolTip()
    tab.set_pool_status(PairPoolStatus(total=10, compared=8, pending=0, failed=2))
    app.processEvents()
    assert tab._export_button.isEnabled()
    assert tab._pool_label.text().startswith("сравнение завершено")
    assert "не удалось прочитать: 2" in tab._pool_label.text()
    tab.set_pool_status(
        PairPoolStatus(total=10, compared=7, pending=0, failed=0, unpersisted=3)
    )
    assert "пересчитываются каждый сеанс" in tab._pool_label.text()
    tab.deleteLater()
    app.processEvents()


def test_target_restores_rule_and_filter(app: QApplication) -> None:
    with tempfile.TemporaryDirectory(prefix="rd_export_tab_") as temp:
        root = Path(temp)
        for name in ("rd", "sq", "robot", "runtime"):
            (root / name).mkdir()
        config = _config(root)
        database = CatalogDatabase(config.db_path)
        database.initialize()
        default = ExportTarget(
            name=DEFAULT_ROBOT_TARGET_NAME,
            root=str(config.robot_root),
            flat_structure=True,
            is_default_robot=True,
            rule="latest_tdo",
            filter_text="2225",
        )
        custom = ExportTarget(
            name="Архив подрядчика",
            root=str(root / "custom"),
            flat_structure=False,
            is_default_robot=False,
            rule="approved",
            filter_text="KSB",
        )
        (root / "custom").mkdir()
        save_export_targets(config, (default, custom))
        tab = RevisionMatrixTab()
        tab.bind_catalog(config, database)
        app.processEvents()
        tab.select_target_by_name("Архив подрядчика")
        app.processEvents()
        assert tab.current_rule() == "approved"
        assert tab.filter_text() == "KSB"
        tab.select_target_by_name(DEFAULT_ROBOT_TARGET_NAME)
        app.processEvents()
        assert tab.current_rule() == "latest_tdo"
        assert tab.filter_text() == "2225"
        tab.deleteLater()
        app.processEvents()


def test_visible_batch_and_preview_groups(app: QApplication) -> None:
    tab = RevisionMatrixTab()
    palette = default_palette()
    cells = [
        _cell("1111", "KSB", "01"),
        _cell("1111", "KSB", "02"),
        _cell("2222", "POS", "01"),
        _cell("2222", "POS", "02"),
        _cell("3333", "SOT", "01"),
        _cell("3333", "SOT", "02"),
        _cell("4444", "PD", "01"),
        _cell("4444", "PD", "02"),
    ]
    tab.set_cells(cells, palette=palette, is_banned=lambda _t, _m: False)
    selections = (
        _selection("1111", "KSB", "add"),
        _selection("2222", "POS", "replace"),
        _selection("3333", "SOT", "same_data"),
        _selection("4444", "PD", "no_source"),
    )
    tab.apply_export_selections(selections)
    app.processEvents()
    table = tab.table()
    assert sum(1 for row in range(table.rowCount()) if not table.isRowHidden(row)) == 4
    tab._filter.setText("1111")
    app.processEvents()
    visible = tab.visible_export_selections()
    assert len(visible) == 1
    assert visible[0].title == "1111"
    target = ExportTarget(
        name="test",
        root=r"C:\robot",
        flat_structure=True,
        is_default_robot=False,
        rule="latest_issued",
        filter_text="",
    )
    plan = plan_for_visible_rows(visible, target=target)
    assert len(plan.items) == 1
    assert plan.items[0].selection.title == "1111"

    tab._filter.setText("")
    app.processEvents()
    all_visible = tab.visible_export_selections()
    assert {row.title for row in all_visible} == {"1111", "2222", "3333", "4444"}
    groups = group_export_preview(all_visible)
    counts = {group.state: group.count for group in groups}
    assert counts["add"] == 1
    assert counts["replace"] == 1
    assert counts["same_data"] == 1
    assert counts["no_source"] == 1
    copied = copy_selections(all_visible)
    assert {row.state for row in copied} == {"add", "replace"}
    assert preview_copy_count(all_visible) == 2
    full_plan = plan_for_visible_rows(all_visible, target=target)
    plan_states = {item.selection.state for item in full_plan.items}
    assert plan_states == {"add", "replace"}
    skipped_states = {row.state for row in full_plan.skipped}
    assert "same_data" in skipped_states
    assert "no_source" in skipped_states
    preview = MtoExportPreviewDialog(all_visible, target)
    edits = preview.findChildren(QPlainTextEdit)
    assert edits
    preview_text = edits[0].toPlainText()
    assert "добавится (1)" in preview_text
    assert "заменится (1)" in preview_text
    assert "данные совпадают (1)" in preview_text
    assert "нет файла (1)" in preview_text
    preview.close()
    tab.deleteLater()
    app.processEvents()


def test_pin_pick_dialog_marks_rule(app: QApplication) -> None:
    older = ExportPinCandidate(
        package_path=r"C:\rd\07_pack",
        package_name="07_pack",
        sequence=7,
        file_path=r"C:\rd\07_pack\old.xlsx",
        file_name="old.xlsx",
        revision_text="01-AN01",
        size=111,
        is_rule_choice=False,
    )
    newest = ExportPinCandidate(
        package_path=r"C:\rd\12_pack",
        package_name="12_pack",
        sequence=12,
        file_path=r"C:\rd\12_pack\rule.xlsx",
        file_name="rule.xlsx",
        revision_text="02",
        size=222,
        is_rule_choice=True,
    )
    dialog = MtoPinPickDialog((older, newest), current_path="")
    app.processEvents()
    headers = [
        dialog._table.horizontalHeaderItem(index).text()
        for index in range(dialog._table.columnCount())
    ]
    assert headers == list(_PIN_DIALOG_HEADERS)
    assert dialog._table.rowCount() == 2
    assert dialog._table.item(0, 2).text() == "old.xlsx"
    assert dialog._table.item(1, 5).text() == _PIN_RULE_MARK
    assert dialog._table.item(0, 5).text() == ""
    assert dialog._table.currentRow() == 1
    dialog._table.selectRow(0)
    dialog.accept()
    assert dialog.selected_candidate() is older
    dialog.deleteLater()
    app.processEvents()


def test_pin_assign_clear_refresh_and_stale(app: QApplication) -> None:
    palette = default_palette()
    with tempfile.TemporaryDirectory(prefix="rd_export_pin_") as temp:
        root = Path(temp)
        (root / "runtime").mkdir()
        (root / "robot").mkdir()
        (root / "sq").mkdir()
        database, records, current_ids, files = _load_3240(root / "data")
        config = CatalogConfig(
            rd_root=Path(RD_ROOT),
            sq_root=root / "sq",
            robot_root=root / "robot",
            runtime_dir=root / "runtime",
            db_path=database.path,
            robot_flat_structure=True,
            skip_dirs=(),
        )
        tab = RevisionMatrixTab()
        tab.bind_catalog(config, database)
        tab.set_cells(
            [
                _cell(_TITLE, _MARK, "01", letters="A"),
                _cell(_TITLE, _MARK, "02", letters="AB", status="tdo_review"),
            ],
            palette=palette,
            is_banned=lambda _t, _m: False,
        )
        tab.set_export_records(records, current_ids)
        app.processEvents()
        rd_item = tab.table().item(0, _COL_RD_REV)
        assert rd_item is not None
        assert rd_item.text() == "02 AB"
        assert "Официальная рев. РД" in rd_item.toolTip()
        assert "Последняя IFC: 01" in rd_item.toolTip()
        rule_sel = tab.selection_for(_TITLE, _MARK)
        assert rule_sel is not None
        assert rule_sel.origin == "rule"
        assert rule_sel.source_path == files["nn12"].path
        export_item = tab.table().item(0, _COL_EXPORT)
        assert export_item is not None
        stored = export_item.data(_ROLE_EXPORT)
        assert isinstance(stored, ExportSelection)
        assert stored.source_path == files["nn12"].path

        candidates = list_export_pin_candidates(
            records,
            title=_TITLE,
            mark=_MARK,
            rd_root=RD_ROOT,
            rule_path=rule_sel.rule_path,
        )
        assert [item.sequence for item in candidates] == [12, 11, 9, 7]
        assert candidates[0].is_rule_choice
        assert not candidates[-1].is_rule_choice
        older = candidates[-1]
        assert older.file_path == files["nn07"].path
        assert older.size == files["nn07"].data["size"]

        assert tab.assign_export_pin(_TITLE, _MARK, older)
        app.processEvents()
        pinned = tab.selection_for(_TITLE, _MARK)
        assert pinned is not None
        assert pinned.origin == "pin"
        assert pinned.source_path == files["nn07"].path
        assert pinned.rule_path == files["nn12"].path
        export_item = tab.table().item(0, _COL_EXPORT)
        stored = export_item.data(_ROLE_EXPORT)
        assert stored.source_path == files["nn07"].path
        assert stored.origin == "pin"

        assert tab.clear_export_pin(_TITLE, _MARK)
        app.processEvents()
        cleared = tab.selection_for(_TITLE, _MARK)
        assert cleared is not None
        assert cleared.origin == "rule"
        assert cleared.source_path == files["nn12"].path

        stale_same = ExportPin(
            title=_TITLE,
            mark=_MARK,
            package_path=str(Path(files["nn12"].path).parent.parent),
            file_path=files["nn12"].path,
            revision_text="02",
            evidence="deadbeef",
            created_at="2026-09-01T00:00:00+00:00",
        )
        save_export_pins(config, (stale_same,))
        tab.set_export_records(records, current_ids)
        app.processEvents()
        refreshed = tab.selection_for(_TITLE, _MARK)
        assert refreshed is not None
        assert refreshed.origin == "pin"
        assert refreshed.state != "pin_stale"
        saved = load_export_pins(config)
        assert len(saved) == 1
        expected = export_pin_evidence_for_kit(
            database,
            records,
            title=_TITLE,
            mark=_MARK,
            rd_root=RD_ROOT,
        )
        assert saved[0].evidence == expected
        assert saved[0].evidence != "deadbeef"

        disagree = ExportPin(
            title=_TITLE,
            mark=_MARK,
            package_path=str(Path(files["nn07"].path).parent.parent),
            file_path=files["nn07"].path,
            revision_text="01-AN01",
            evidence="deadbeef",
            created_at="2026-09-01T00:00:00+00:00",
        )
        save_export_pins(config, (disagree,))
        tab.set_export_records(records, current_ids)
        app.processEvents()
        stale = tab.selection_for(_TITLE, _MARK)
        assert stale is not None
        assert stale.origin == "pin_stale"
        assert stale.state == "pin_stale"
        assert stale.source_path == files["nn07"].path
        assert stale.rule_path == files["nn12"].path
        tip = tab.table().item(0, _COL_EXPORT).toolTip()
        assert files["nn07"].path in tip
        assert files["nn12"].path in tip
        assert "закреплённый файл" in tip
        assert "файл по правилу" in tip
        after_stale = load_export_pins(config)
        assert after_stale[0].evidence == "deadbeef"
        assert after_stale[0].file_path == files["nn07"].path
        tab.prepare_close()
        tab.deleteLater()
        app.processEvents()


def main() -> None:
    """Run offscreen cockpit checks."""

    app = QApplication.instance() or QApplication([])
    test_six_states_and_revision_shift(app)
    print("six states + revision shift : OK")
    test_kit_jump_only_from_context_action(app)
    print("kit jump context action only : OK")
    test_apply_auto_mto_compare_paints_composite(app)
    print("auto mto composite cell : OK")
    test_auto_mto_compare_paints_rev_match(app)
    print("auto mto rev match paint : OK")
    test_heatmap_sorts_numeric_titles_and_marks(app)
    print("heatmap title/mark sort : OK")
    test_auto_mto_enqueue_skips_cached_and_queues_when_running(app)
    print("auto mto enqueue/queue : OK")
    test_stale_export_path_enqueues_current_file(app)
    print("stale export path enqueues : OK")
    test_same_rd_path_dash_variant_does_not_enqueue(app)
    print("dash-variant path is same file : OK")
    test_export_path_change_replaces_queued_old_file(app)
    print("export path replaces queued job : OK")
    test_export_gate(app)
    print("export gate : OK")
    test_target_restores_rule_and_filter(app)
    print("target restores rule/filter : OK")
    test_visible_batch_and_preview_groups(app)
    print("visible batch + preview groups : OK")
    test_pin_pick_dialog_marks_rule(app)
    print("pin pick dialog : OK")
    test_pin_assign_clear_refresh_and_stale(app)
    print("pin assign/clear/refresh/stale : OK")
    print("RD catalog export tab: OK")


if __name__ == "__main__":
    main()
