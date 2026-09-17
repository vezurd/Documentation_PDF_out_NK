"""Offscreen smoke test for the RD catalog window with local empty sources."""

from __future__ import annotations

import inspect
import io
import json
import os
import sys
import tempfile
import time
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dataclasses import replace
from types import SimpleNamespace
from PySide6.QtCore import QPoint, QSettings, Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QHeaderView,
        QMessageBox,
        QMenu,
        QPlainTextEdit,
        QPushButton,
    QTableWidget,
    QTableWidgetItem,
)
from PySide6.QtTest import QTest

from rd_catalog.ban_dialog import BannedTitlesDialog
from rd_catalog.google_f_write import JournalWriteJob
from rd_catalog.kits_legend_dialog import KitsPaintLegendDialog
from rd_catalog.sheet_de_sync import SheetDeSyncRow
from rd_catalog.sheet_de_sync_dialog import SHEET_DE_SYNC_BUTTON, SheetDeSyncDialog
from rd_catalog.skip_dirs_dialog import SkipDirsDialog
from rd_catalog.config import load_config
from rd_catalog.context_menu_qt import apply_usage_styles
from rd_catalog.context_menu_usage import MENU_KITS, get_context_menu_usage
from rd_catalog.db import CatalogDatabase, KitPipelineRow
from rd_catalog.kits import (
    KitMatrixRow,
    KitSummary,
    SourceKitSnapshot,
    kit_identity_key,
    mto_content_equal_by_kit,
)
from rd_catalog.layout_report import (
    build_layout_report,
    format_layout_report_summary,
)
from rd_catalog.models import FileKind, FileRecord, ReviewState, SourceKind
from rd_catalog.mto_export import (
    PIN_COLUMN_HEADER,
    ExportPin,
    ExportSelection,
    export_pin_view,
    save_export_pins,
)
from rd_catalog.customer_pi_auto_mto import AUTO_MTO_COMPARE_STATUS_HEADER, AutoMtoFile
from rd_catalog.mto_worklist_tab import (
    MtoWorklistTab,
    _COL_DATE,
    _COL_FILE,
    _COL_GOOGLE,
    _COL_PACKAGE,
    _COL_PATH,
    _COL_PIN,
    _COL_PROBLEMS,
)
from rd_catalog.parse import issued_package_dir
from rd_catalog.pipeline import FolderTreeHint, KitPipelineStatus, MtoWorklistRow
from rd_catalog.status_colors import color_for, default_palette
from rd_catalog.theme import apply_catalog_theme
from rd_catalog import window as rd_window
from rd_catalog.doc_bundle import ANNULLED_MARKER, WORKING_MARKER, format_file_save_date
from rd_catalog.window import (
    CatalogWindow,
    attach_kits_table_tips_ctrl_click,
    scroll_kits_tips_header_to_top,
    _HISTORY_COL_DATE,
    _HISTORY_COL_NAME,
    _HISTORY_COL_PACKAGE,
    _HISTORY_COL_PIN,
    _HISTORY_HEADERS,
    _LATEST_FILE_DATE_BG,
    _KITS_COL_AN,
    _KITS_COL_GOOGLE_REV,
    _KITS_COL_ISSUANCE_REV,
    _KITS_COL_PIN,
    _KITS_COL_RD_REV,
    _KITS_COL_ROBOT_REV,
    _KITS_COL_SQ_REV,
    _KITS_COL_WORKING,
    _KITS_HEADERS,
    _KITS_REV_SOURCE_COLUMNS,
    _LAYOUT_MISMATCH_BG,
    _REV_MATCH_BG,
    _ROBOT_ORIGIN_BG,
    _ReadOnlyCopyDelegate,
    _paint_revision_cell,
    _paint_robot_origin_cell,
)


def _pipeline_rd_path(
    rd_root: Path,
    *,
    title: str,
    mark: str,
    folder: str,
    filename: str,
    media: str = "PDF",
) -> str:
    """Return a canonical issued-transfer path under *rd_root*.

    Same layout as ``tmp/test_rd_catalog_pipeline.py`` ``_record``:
    ``<rd_root>/<title>/06_<mark>/Для передачи/<NN_…>/<PDF|DWG>/<file>``.
    """

    return str(
        Path(rd_root)
        / title
        / f"06_{mark}"
        / "Для передачи"
        / folder
        / media
        / filename
    )


def _issued_nn(
    title: str,
    mark: str,
    *,
    sequence: int = 1,
    revision: str = "01",
) -> str:
    return f"{sequence:02d}_рев.{revision}_AGCC.287-{title}-{mark}"


def _stub_record(
    file_id: int,
    *,
    path: str,
    title: str,
    mark: str,
    revision: str = "01",
    core_stem: str | None = None,
    file_kind: str = FileKind.PDF.value,
    transfer_sequence: int | None = None,
    transfer_name: str | None = None,
    mtime_ns: int = 0,
) -> FileRecord:
    """Build a tiny RD FileRecord for tree/jump assertions."""

    name = Path(path).name
    data: dict[str, object] = {
        "file_kind": file_kind,
        "name": name,
        "parse_status": "parsed",
        "title": title,
        "mark": mark,
        "revision": revision,
        "appendix": None,
        "core_stem": core_stem or f"AGCC.287-{title}-{mark}.OD-0001",
        "discipline_block": "OD-0001",
        "mtime_ns": mtime_ns,
    }
    if transfer_sequence is not None:
        data["transfer_sequence"] = transfer_sequence
    if transfer_name:
        data["transfer_name"] = transfer_name
    return FileRecord(
        id=file_id,
        path=path,
        path_key=path.casefold(),
        source=SourceKind.RD,
        present=True,
        review_state=ReviewState.ACKNOWLEDGED,
        first_seen_run_id=1,
        last_seen_run_id=1,
        data=data,
    )


def _wait_catalog_startup(window, app, timeout_s: float = 15.0) -> None:
    """Pump Qt until the off-GUI first-paint hydrate finishes."""

    deadline = time.monotonic() + timeout_s
    while window._startup_hydrate_busy() and time.monotonic() < deadline:
        app.processEvents()
        thread = getattr(window, "_startup_thread", None)
        if thread is not None:
            thread.wait(50)
    app.processEvents()
    assert window._startup_hydrate_busy() is False


def _dummy_matrix_cell(
    title: str,
    mark: str,
    revision_text: str = "01",
) -> SimpleNamespace:
    return SimpleNamespace(
        title=title,
        mark=mark,
        revision_text=revision_text,
        pipeline_status="code_a",
        letters="A",
        problem_kinds_json="[]",
        is_as_build=False,
        is_current=False,
        is_current_ifc=False,
        has_mto=True,
    )


def _check_scan_keep_view_tree_rebuild_contract() -> None:
    """SQ→RD keep_view must not force a documents-tree rebuild."""

    src = inspect.getsource(CatalogWindow._on_scan_finished)
    assert "sq_and_rd" not in src
    assert "stay_on_documents" in src
    assert "rebuild_document_tree = restore is None or stay_on_documents" in src


def _check_kits_context_menu_popup_cost(window: CatalogWindow) -> None:
    """Комплекты right-click must not walk the whole catalog or stat Auto MTO."""

    popup_src = inspect.getsource(CatalogWindow._popup_kits_context_menu)
    assert "is_file()" not in popup_src
    assert "currentRow()" in popup_src
    pkg_src = inspect.getsource(CatalogWindow._package_folder_records)
    assert "_records_for_kit" in pkg_src
    rec_src = inspect.getsource(CatalogWindow._record_by_path_key)
    assert "_records_by_path_key" in rec_src

    rd_root = Path(window.config.rd_root)
    mto = _stub_record(
        801,
        path=_pipeline_rd_path(
            rd_root,
            title="8181",
            mark="KSB",
            folder=_issued_nn("8181", "KSB"),
            filename="AGCC.287-8181-KSB.MTO-0001_01_RU.xlsx",
        ),
        title="8181",
        mark="KSB",
        file_kind=FileKind.MTO_XLSX.value,
    )
    pdf = _stub_record(
        802,
        path=_pipeline_rd_path(
            rd_root,
            title="8181",
            mark="KSB",
            folder=_issued_nn("8181", "KSB"),
            filename="AGCC.287-8181-KSB.OD-0001_01_RU.pdf",
        ),
        title="8181",
        mark="KSB",
    )
    noise = [
        _stub_record(
            9000 + index,
            path=_pipeline_rd_path(
                rd_root,
                title="9090",
                mark="POS",
                folder=_issued_nn("9090", "POS", sequence=1 + (index % 9)),
                filename=f"AGCC.287-9090-POS.OD-{index:04d}_01_RU.pdf",
            ),
            title="9090",
            mark="POS",
        )
        for index in range(400)
    ]
    saved_records = list(window._all_records)
    saved_by_id = dict(window._record_by_id)
    saved_contour = window._contour_records
    try:
        window._assign_catalog_records([mto, pdf, *noise])
        found = window._record_by_path_key(mto.path)
        assert found is mto
        issued_calls = {"n": 0}
        real_issued = issued_package_dir

        def _count_issued(path: str) -> str:
            issued_calls["n"] += 1
            return real_issued(path)

        with patch("rd_catalog.window.issued_package_dir", side_effect=_count_issued):
            siblings = window._package_folder_records(mto)
        assert mto in siblings
        assert pdf in siblings
        assert all(str(item.data.get("title")) != "9090" for item in siblings)
        assert issued_calls["n"] < 20
        kit_row = KitMatrixRow(
            title="8181",
            mark="KSB",
            title_system="8181-KSB",
            rd=SourceKitSnapshot(present=True, revision_text="01"),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.GAP_ROBOT,
        )
        mixed_src = inspect.getsource(CatalogWindow._mixed_title_open_folders_for_kit)
        assert "_records_for_kit" in mixed_src
        assert window._mixed_title_open_folders_for_kit(kit_row) == ()
        assert len(window._records_for_kit("8181", "KSB")) == 2
    finally:
        window._assign_catalog_records(saved_records, contour=saved_contour)
        window._record_by_id = saved_by_id


def _check_mto_compare_not_planned_on_gui(window: CatalogWindow) -> None:
    """RD↔robot MTO plan must be built in the child job, not on the GUI thread."""

    src = inspect.getsource(CatalogWindow._enqueue_mto_compare)
    assert "official_rd_mto_overlay" not in src
    assert "build_mto_compare_plan" not in src
    assert "list_present_robot_mto_files" not in src
    assert window._mto_compare_thread is None
    assert window._enqueue_mto_compare(frozenset(), ()) is False
    assert window._mto_compare_thread is None


def _check_overlay_auto_mto_enqueue_while_heatmap_deferred(window: CatalogWindow) -> None:
    """«другой файл» / «не сверялось» must queue from overlay before heatmap visit."""

    tab = window._revision_matrix_tab
    tab.set_auto_mto_queue_paused(True)
    tab._auto_mto_queue.clear()
    rd_path = r"C:\rd\AGCC.287-1111-KSB.MTO-0001_01_RU.xlsx"
    files = (
        AutoMtoFile(
            title="1111",
            mark="KSB",
            spec="AGCC.287-1111-KSB.MTO-0001",
            rd_revision="01",
            relpath="1111/KSB/AGCC.287-1111-KSB.MTO-0001_01_RU.xlsx",
            row_count=1,
            fingerprint="mto",
        ),
    )
    previous_index = window._auto_mto_by_kit
    previous_deferred = set(window._deferred_widgets)
    window._auto_mto_by_kit = {("1111", "ksb"): files}
    tab.set_auto_mto_index(window._auto_mto_by_kit)
    window._rd_mto_overlay_cache = (
        (id(window._mto_worklist_rows), id(window._kit_pipelines)),
        {("1111", "ksb"): (rd_path, "01")},
    )
    window._deferred_widgets = set(rd_window._DEFERRED_ALL)
    with patch.object(window, "_catalog_workers_busy", return_value=True):
        window._schedule_auto_mto_compares()
    assert tab._auto_mto_queue == [(("1111", "ksb"), rd_path)]
    tab._auto_mto_queue.clear()
    window._auto_mto_by_kit = previous_index
    tab.set_auto_mto_index(previous_index)
    window._rd_mto_overlay_cache = None
    window._deferred_widgets = previous_deferred
    tab.set_auto_mto_queue_paused(window._catalog_workers_busy())


def _matrix_cell_stub(
    title: str,
    mark: str,
    rev: str,
    *,
    status: str = "code_a",
    letters: str = "A",
) -> SimpleNamespace:
    return SimpleNamespace(
        title=title,
        mark=mark,
        revision_text=rev,
        pipeline_status=status,
        letters=letters,
        problem_kinds_json="[]",
        is_as_build=False,
        is_current=False,
        is_current_ifc=False,
        has_mto=True,
    )


def _heatmap_row_letters(tab, title: str, mark: str, column: int) -> str:
    wanted = kit_identity_key(title, mark)
    for row in range(tab._table.rowCount()):
        kit = tab._kit_at_row(row)
        if kit is None or kit_identity_key(*kit) != wanted:
            continue
        item = tab._table.item(row, column)
        return item.text() if item is not None else ""
    raise AssertionError(f"heatmap row {title}-{mark} not found")


def _check_heatmap_patch_cells(window: CatalogWindow) -> None:
    """In-place heatmap patch must not rebuild columns or drop other kits."""

    tab = window._revision_matrix_tab
    assert tab.patch_cells_for_kits(
        (), {kit_identity_key("2225", "KSB")}
    ) is False
    cells = [
        _matrix_cell_stub("2225", "KSB", "01"),
        _matrix_cell_stub("2225", "KSB", "02"),
        _matrix_cell_stub("9192", "POS", "01"),
    ]
    tab.set_cells(
        cells,
        palette=window._status_colors,
        is_banned=lambda _title, _mark: False,
    )
    assert tab._table.rowCount() == 2
    patched = [
        _matrix_cell_stub("2225", "KSB", "01", status="working", letters="W"),
        _matrix_cell_stub("2225", "KSB", "02", status="working", letters="W"),
    ]
    assert tab.patch_cells_for_kits(
        patched, {kit_identity_key("2225", "KSB")}
    ) is True
    assert tab._table.rowCount() == 2
    assert _heatmap_row_letters(tab, "2225", "KSB", 6) == "W"
    assert _heatmap_row_letters(tab, "9192", "POS", 6) == "A"
    tab.set_cells(
        (),
        palette=window._status_colors,
        is_banned=lambda _title, _mark: False,
    )


def _check_scoped_gui_refresh(window: CatalogWindow) -> None:
    """One-kit writes must relabel the tree and leave a deferred heatmap."""

    rd_root = Path(window.config.rd_root)
    rec_a = _stub_record(
        101,
        path=_pipeline_rd_path(
            rd_root,
            title="9192",
            mark="POS",
            folder=_issued_nn("9192", "POS"),
            filename="AGCC.287-9192-POS.OD-0001_01_RU.pdf",
        ),
        title="9192",
        mark="POS",
    )
    rec_b = _stub_record(
        102,
        path=_pipeline_rd_path(
            rd_root,
            title="9192",
            mark="KSB",
            folder=_issued_nn("9192", "KSB"),
            filename="AGCC.287-9192-KSB.OD-0001_01_RU.pdf",
        ),
        title="9192",
        mark="KSB",
    )
    saved_records = list(window._all_records)
    saved_by_id = dict(window._record_by_id)
    saved_contour = window._contour_records
    saved_deferred = set(window._deferred_widgets)
    saved_kit_rows = window._kit_rows
    saved_pipelines = dict(window._kit_pipelines)
    saved_official_token = window._official_ids_token
    saved_official_value = set(window._official_ids_value)
    saved_hints = dict(window._folder_hints)
    saved_worklist = window._mto_worklist_rows
    saved_folder = window._doc_show_folder.isChecked()
    saved_skip = window.config.skip_dirs
    try:
        window.config = replace(window.config, skip_dirs=())
        window._assign_catalog_records([rec_a, rec_b])
        window._deferred_widgets.discard(rd_window._DEFERRED_TREE)
        window._rebuild_document_tree()
        window._deferred_widgets.add(rd_window._DEFERRED_HEATMAP)
        title_item = window._doc_tree.topLevelItem(0)
        assert title_item is not None
        mark_item = title_item.child(0)
        assert mark_item is not None
        revision_item = mark_item.child(0)
        assert revision_item is not None
        hint = FolderTreeHint(is_working=True, working_origin="manual")
        with patch.object(window.database, "upsert_working_flag"), patch(
            "rd_catalog.window.rebuild_pipeline"
        ), patch.object(window, "_folder_hint_for", return_value=hint):
            window._apply_tree_working_flag(revision_item, remove=False)
        assert window._doc_tree.topLevelItem(0) is title_item
        assert title_item.child(0) is mark_item
        assert mark_item.child(0) is revision_item
        assert rd_window._DEFERRED_HEATMAP in window._deferred_widgets
        assert window._revision_matrix_table.rowCount() == 0
        assert WORKING_MARKER in revision_item.text(0)

        try:
            annulled_hint = FolderTreeHint(is_annulled=True)
        except TypeError:
            annulled_hint = SimpleNamespace(
                review_status="",
                review_full="",
                match_reason="",
                f_label="",
                is_working=False,
                working_origin="",
                mto_status="",
                mto_letters="",
                mto_revision_text="",
                has_mto_file=False,
                is_current_mto=False,
                is_current_ifc=False,
                problem_kinds=(),
                is_annulled=True,
            )
        with patch.object(window.database, "upsert_annulled_flag"), patch(
            "rd_catalog.window.rebuild_pipeline"
        ), patch.object(window, "_folder_hint_for", return_value=annulled_hint):
            window._apply_tree_annulled_flag(revision_item, remove=False)
        assert window._doc_tree.topLevelItem(0) is title_item
        assert title_item.child(0) is mark_item
        assert mark_item.child(0) is revision_item
        assert rd_window._DEFERRED_HEATMAP in window._deferred_widgets
        assert window._revision_matrix_table.rowCount() == 0
        assert ANNULLED_MARKER in revision_item.text(0)

        window._doc_show_folder.setChecked(not saved_folder)
        assert mark_item.child(0) is revision_item

        window._refresh_revision_matrix()
        assert rd_window._DEFERRED_HEATMAP in window._deferred_widgets
        assert window._revision_matrix_table.rowCount() == 0
    finally:
        window._doc_show_folder.blockSignals(True)
        window._doc_show_folder.setChecked(saved_folder)
        window._doc_show_folder.blockSignals(False)
        window._all_records = saved_records
        window._record_by_id = saved_by_id
        window._contour_records = saved_contour
        window._contour_ids_cache = None
        window._tree_kit_identities_cache = None
        window._deferred_widgets = saved_deferred
        window._kit_rows = saved_kit_rows
        window._kit_pipelines = saved_pipelines
        window._official_ids_token = saved_official_token
        window._official_ids_value = saved_official_value
        window._folder_hints = saved_hints
        window._set_mto_worklist_rows(saved_worklist)
        window.config = replace(window.config, skip_dirs=saved_skip)
        window._kits_table.setRowCount(len(saved_kit_rows))
        window._paint_kits_table()
        window._rebuild_document_tree()


def _check_scoped_kits_table_refresh(window: CatalogWindow) -> None:
    """Folder-scoped ``refresh`` must paint Комплекты with ``kit_keys=``."""

    rd_root = Path(window.config.rd_root)
    rec_a = _stub_record(
        101,
        path=_pipeline_rd_path(
            rd_root,
            title="9192",
            mark="POS",
            folder=_issued_nn("9192", "POS"),
            filename="AGCC.287-9192-POS.OD-0001_01_RU.pdf",
        ),
        title="9192",
        mark="POS",
    )
    rec_b = _stub_record(
        102,
        path=_pipeline_rd_path(
            rd_root,
            title="9192",
            mark="KSB",
            folder=_issued_nn("9192", "KSB"),
            filename="AGCC.287-9192-KSB.OD-0001_01_RU.pdf",
        ),
        title="9192",
        mark="KSB",
    )
    pos_key = kit_identity_key("9192", "POS")
    pos_folder = str(rd_root / "9192" / "06_POS")
    pos_row = KitMatrixRow(
        title="9192",
        mark="POS",
        title_system="9192-POS",
        rd=SourceKitSnapshot(present=True, revision_text="01"),
        robot=SourceKitSnapshot(),
        sq=SourceKitSnapshot(),
        google=None,
        issuance=None,
        flags=(),
        summary=KitSummary.GAP_ROBOT,
    )
    ksb_row = KitMatrixRow(
        title="9192",
        mark="KSB",
        title_system="9192-KSB",
        rd=SourceKitSnapshot(present=True, revision_text="01"),
        robot=SourceKitSnapshot(),
        sq=SourceKitSnapshot(),
        google=None,
        issuance=None,
        flags=(),
        summary=KitSummary.GAP_ROBOT,
    )
    saved_records = list(window._all_records)
    saved_by_id = dict(window._record_by_id)
    saved_contour = window._contour_records
    saved_deferred = set(window._deferred_widgets)
    saved_pending = window._secondary_tabs_pending
    saved_kit_rows = window._kit_rows
    saved_pipelines = dict(window._kit_pipelines)
    saved_official_token = window._official_ids_token
    saved_official_value = set(window._official_ids_value)
    saved_detected = set(window._detected_current_ids)
    saved_hints = dict(window._folder_hints)
    saved_worklist = window._mto_worklist_rows
    saved_mto_rows = list(window._mto_rows)
    saved_collisions = list(window._collision_rows)
    try:
        window._kit_rows = (pos_row, ksb_row)
        window._official_ids_token = None
        window._official_ids_value = {101, 102}
        with (
            patch.object(window.database, "list_files", return_value=[rec_a, rec_b]),
            patch.object(window, "_should_rebuild_empty_derived", return_value=False),
            patch.object(window, "_refresh_kits_table") as refresh_kits,
        ):
            window.refresh(
                pipeline_subtrees=(pos_folder,),
                ingest_google=False,
                rebuild_derived=False,
                rebuild_document_tree=False,
                defer_secondary=True,
            )
        refresh_kits.assert_called_once()
        assert refresh_kits.call_args.kwargs.get("kit_keys") == {pos_key}

        window._kit_rows = (pos_row, ksb_row)
        window._kit_pipelines = dict(saved_pipelines)
        window._all_records = [rec_a, rec_b]
        window._contour_records = (rec_a, rec_b)
        window._detected_current_ids = {101, 102}
        window._official_ids_token = None
        window._official_ids_value = {101, 102}
        window._issuance_kits_fresh = True
        with (
            patch(
                "rd_catalog.window.patch_official_detected_current_ids",
                wraps=rd_window.patch_official_detected_current_ids,
            ) as patched_ids,
            patch("rd_catalog.window.official_detected_current_ids") as full_ids,
            patch("rd_catalog.window.build_kit_matrix", return_value=(pos_row,)),
            patch.object(window.database, "list_kit_pipelines", return_value=[]),
        ):
            window._rebuild_kit_rows_for_keys({pos_key})
        patched_ids.assert_called_once()
        assert patched_ids.call_args.kwargs.get("kit_keys") == {pos_key}
        full_ids.assert_not_called()
        assert window._official_ids_token is not None
        assert isinstance(window._official_ids_value, set)
    finally:
        window._all_records = saved_records
        window._record_by_id = saved_by_id
        window._contour_records = saved_contour
        window._contour_ids_cache = None
        window._tree_kit_identities_cache = None
        window._deferred_widgets = saved_deferred
        window._secondary_tabs_pending = saved_pending
        window._kit_rows = saved_kit_rows
        window._kit_pipelines = saved_pipelines
        window._official_ids_token = saved_official_token
        window._official_ids_value = saved_official_value
        window._detected_current_ids = saved_detected
        window._folder_hints = saved_hints
        window._set_mto_worklist_rows(saved_worklist)
        window._mto_rows = saved_mto_rows
        window._collision_rows = saved_collisions
        window._kits_table.setRowCount(len(saved_kit_rows))
        window._paint_kits_table()


def _visible_row_count(table) -> int:
    return sum(
        1 for row in range(table.rowCount()) if not table.isRowHidden(row)
    )


def _synthetic_worklist_row(**overrides) -> MtoWorklistRow:
    payload = {
        "title": "2225",
        "mark": "KSB",
        "revision_text": "01",
        "status": "tdo_review",
        "letters": "T",
        "is_as_build": False,
        "is_current": False,
        "is_current_ifc": False,
        "has_mto": True,
        "mto_path": r"C:\rd\file.xlsx",
        "gap_kind": "",
    }
    payload.update(overrides)
    return MtoWorklistRow(**payload)


def _pin_view_lookup(
    pin: ExportPin | None,
    *,
    origin: str = "",
    rule_path: str = "",
    title: str = "9192",
    mark: str = "POS",
):
    """Return a kit-scoped ``export_pin_view`` callback for worklist tests."""

    def lookup(row_title: str, row_mark: str):
        if kit_identity_key(row_title, row_mark) != kit_identity_key(
            title, mark
        ):
            return export_pin_view(None)
        return export_pin_view(pin, origin=origin, rule_path=rule_path)

    return lookup


def _assert_worklist_pin_and_tail(
    table,
    *,
    pin_text: str,
    mto_path: str,
    package_label: str,
    date_text: str,
    stale: bool = False,
    palette=None,
    tooltip_parts: tuple[str, ...] = (),
) -> None:
    """Check the pin cell and the columns that follow the insertion point."""

    pin_item = table.item(0, _COL_PIN)
    assert pin_item is not None
    assert pin_item.text() == pin_text
    for part in tooltip_parts:
        assert part in pin_item.toolTip()
    if stale:
        assert palette is not None
        stale_hex = QColor(
            color_for(palette, "export_pin_stale")
        ).name().casefold()
        assert pin_item.background().color().name().casefold() == stale_hex
    assert table.item(0, _COL_FILE).text() == ("Да" if mto_path else "Нет")
    assert table.item(0, _COL_DATE).text() == date_text
    assert table.item(0, _COL_PACKAGE).text() == package_label
    assert table.item(0, _COL_PATH).text() == mto_path


def _visible_matrix_kits(table) -> list[tuple[str, str]]:
    kits: list[tuple[str, str]] = []
    for row in range(table.rowCount()):
        if table.isRowHidden(row):
            continue
        title_item = table.item(row, 0)
        mark_item = table.item(row, 1)
        if title_item is None or mark_item is None:
            continue
        kits.append((title_item.text(), mark_item.text()))
    return kits


def _check_layout_report_action(window: CatalogWindow, rd_root: Path) -> None:
    """Assert enablement, summary, save sheets, and the zero-violation text."""

    action = window._layout_report_action
    assert action.text() == "Отчёт о раскладке папок РД…"
    saved_records = list(window._all_records)
    saved_ids = dict(window._record_by_id)
    saved_config = window.config

    window._all_records = []
    window._record_by_id = {}
    window._update_action_states()
    assert not action.isEnabled()
    assert action.toolTip()

    lost = _stub_record(
        9001,
        path=str(
            rd_root
            / "1715"
            / "PD"
            / "Рабочая"
            / "AGCC.287-1715-PD.MTO-0001_B_RU.xlsx"
        ),
        title="1715",
        mark="PD",
        revision="B",
        file_kind=FileKind.MTO_XLSX.value,
    )
    window._all_records = [lost]
    window._record_by_id = {lost.id: lost}
    window.config = replace(window.config, rd_root="")
    window._update_action_states()
    assert not action.isEnabled()
    assert action.toolTip()

    window.config = saved_config
    window._update_action_states()
    assert action.isEnabled()
    assert action.toolTip()
    report = build_layout_report(
        records=window._all_records,
        rd_root=window.config.rd_root,
    )
    summary = format_layout_report_summary(report)
    assert str(report.total_files) in summary
    assert f"MTO вне раскладки: {report.total_mto_files}" in summary
    assert str(len(report.fully_lost_kits)) in summary

    with tempfile.TemporaryDirectory(prefix="rd_layout_gui_") as raw:
        dest = Path(raw) / "layout.xlsx"
        with (
            patch(
                "rd_catalog.window.QMessageBox.question",
                return_value=QMessageBox.StandardButton.Yes,
            ),
            patch(
                "rd_catalog.window.QFileDialog.getSaveFileName",
                return_value=(str(dest), "Excel (*.xlsx)"),
            ),
            patch(
                "rd_catalog.window.open_path",
                return_value=(True, str(dest)),
            ),
        ):
            window._on_layout_report()
        payload = dest.read_bytes()
    import openpyxl

    workbook = openpyxl.load_workbook(io.BytesIO(payload))
    try:
        assert workbook.sheetnames == [
            "Сводка",
            "Нет в каталоге",
            "Что исправить",
        ]
    finally:
        workbook.close()

    canonical = _stub_record(
        9002,
        path=str(
            rd_root
            / "9110"
            / "KSB"
            / "Для передачи"
            / "05_рев.0_AGCC.287-9110-KSB"
            / "PDF"
            / "AGCC.287-9110-KSB.OD-0001_0_RU.pdf"
        ),
        title="9110",
        mark="KSB",
        revision="0",
    )
    zero = build_layout_report(records=[canonical], rd_root=rd_root)
    zero_text = format_layout_report_summary(zero)
    assert zero.total_files == 0
    assert "Нарушений раскладки нет" in zero_text
    assert zero_text.strip()

    window._all_records = saved_records
    window._record_by_id = saved_ids
    window.config = saved_config
    window._update_action_states()


def _check_kits_tips_ctrl_click(app: QApplication) -> None:
    """Ctrl+click on a Комплекты cell puts that column's === block on top."""

    from rd_catalog.monitor_views import (
        MonitorCell,
        format_kits_row_tooltips,
        format_kits_tips_pane,
        kits_tooltip_header_offset,
        kits_tooltip_section_start,
    )

    headers = (
        "Сводка",
        "РД · рев.",
        "Статус рассмотрения",
        "Статус согласования",
        "Google · статус",
        "Выдача · статус",
    )
    cells = {}
    for header in headers:
        cells[header] = MonitorCell(
            text=header,
            tooltip="\n".join(f"{header} строка {index}" for index in range(40)),
        )
    body = format_kits_row_tooltips("6816", "KSB", cells)
    pane = format_kits_tips_pane(body, with_ctrl_click_hint=True)
    edit = QPlainTextEdit()
    edit.setReadOnly(True)
    edit.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
    edit.setPlainText(pane)
    edit.resize(480, 110)
    edit.show()
    table = QTableWidget(1, len(headers))
    table.setHorizontalHeaderLabels(list(headers))
    table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    for column, header in enumerate(headers):
        table.setItem(0, column, QTableWidgetItem(header))
        table.setColumnWidth(column, 140)
    table.resize(900, 160)
    table.show()
    app.processEvents()
    attach_kits_table_tips_ctrl_click(
        table,
        lambda row, column: scroll_kits_tips_header_to_top(
            edit, table.horizontalHeaderItem(column).text()
        ),
    )
    target = pane.find("=== Статус согласования ===")
    assert target > 0
    start = scroll_kits_tips_header_to_top(edit, "Статус согласования")
    assert start == kits_tooltip_header_offset(pane, "Статус согласования")
    assert start == kits_tooltip_section_start(pane, target + 12)
    app.processEvents()
    first = edit.cursorForPosition(QPoint(4, 4))
    line = edit.document().findBlock(first.position()).text()
    assert line.startswith("=== Статус согласования ==="), line
    edit.verticalScrollBar().setValue(0)
    app.processEvents()
    click_col = headers.index("Статус согласования")
    click_at = table.visualItemRect(table.item(0, click_col)).center()
    QTest.mouseClick(
        table.viewport(),
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.ControlModifier,
        click_at,
    )
    app.processEvents()
    first = edit.cursorForPosition(QPoint(4, 4))
    line = edit.document().findBlock(first.position()).text()
    assert line.startswith("=== Статус согласования ==="), line
    table.close()
    table.deleteLater()
    edit.close()
    edit.deleteLater()
    app.processEvents()


def main() -> None:
    """Create, refresh, and close the window without starting a source scan."""

    with tempfile.TemporaryDirectory(prefix="rd_catalog_gui_") as temp:
        root = Path(temp)
        for name in ("rd", "sq", "robot", "runtime", "settings"):
            Path(root, name).mkdir()
        override = root / "config.json"
        override.write_text(
            json.dumps(
                {
                    "rd_root": str(root / "rd"),
                    "sq_root": str(root / "sq"),
                    "robot_root": str(root / "robot"),
                    "runtime_dir": str(root / "runtime"),
                    "db_path": str(root / "runtime" / "catalog.sqlite"),
                    "robot_flat_structure": True,
                    "skip_dirs": ["old", "tmp"],
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        config = load_config(override)
        database = CatalogDatabase(config.db_path)
        database.initialize()

        app = QApplication.instance() or QApplication([])
        QSettings.setDefaultFormat(QSettings.Format.IniFormat)
        QSettings.setPath(
            QSettings.Format.IniFormat,
            QSettings.Scope.UserScope,
            str(root / "settings"),
        )
        rd_window._SETTINGS_APPLICATION = "rd_catalog_gui_smoke_test"
        throwaway_settings = QSettings(
            rd_window._SETTINGS_ORGANIZATION,
            rd_window._SETTINGS_APPLICATION,
        )
        throwaway_settings.clear()
        throwaway_settings.sync()
        apply_catalog_theme(app)
        _check_kits_tips_ctrl_click(app)
        window = CatalogWindow(config, database)
        window.show()
        app.processEvents()
        _wait_catalog_startup(window, app)
        assert hasattr(window, "_kits_load_banner")
        assert window._kits_load_banner.isVisible() is False
        # Off-GUI hydrate on an empty local DB; no pipeline rebuild.
        assert window._busy() is False
        assert window._deferred_widgets == set(rd_window._DEFERRED_ALL)
        assert window._kits_table.rowCount() == 0
        assert window._kits_progress.text() == (
            "Ок: 0 / 0 · Совпадает: 0 / 0 · Код A: 0 · ТДО: 0 · "
            "Нет в РД: 0 · Проблемы MTO: 0"
        )
        assert window._kits_progress.toolTip() == ""
        QTest.mouseClick(window._kits_progress, Qt.MouseButton.LeftButton)
        assert QApplication.clipboard().text() == window._kits_progress.text()
        assert window._web_autostart_cb.isChecked() is False
        assert window._web_server.running is False
        assert "0.0.0.0" not in window._web_server.local_url()
        assert "0.0.0.0" not in window._web_server.share_url()
        assert window._web_url_button.text() == "WEB выкл."
        assert window._web_copy_btn.text() == "Ссылка для коллег"
        rd_root = Path(window.config.rd_root)
        window.refresh()
        assert window._mto_compare_thread is None
        assert window._busy() is False
        assert not window._deferred_widgets
        _check_layout_report_action(window, Path(window.config.rd_root))
        window._mto_compare_thread = object()  # type: ignore[assignment]
        assert window._busy() is False
        window._mto_compare_thread = None
        window._revision_matrix_tab._auto_mto_compare_thread = type(
            "_Running",
            (),
            {"isRunning": staticmethod(lambda: True)},
        )()
        assert window._revision_matrix_tab.is_auto_mto_compare_running() is True
        assert window._catalog_workers_busy() is False
        assert window._busy() is False
        window._revision_matrix_tab._auto_mto_compare_thread = None
        _check_overlay_auto_mto_enqueue_while_heatmap_deferred(window)
        _check_scan_keep_view_tree_rebuild_contract()
        _check_kits_context_menu_popup_cost(window)
        _check_mto_compare_not_planned_on_gui(window)
        _check_scoped_gui_refresh(window)
        _check_scoped_kits_table_refresh(window)
        _check_heatmap_patch_cells(window)
        assert hasattr(window, "_mto_sync_label")
        assert "Сверка" in window._mto_sync_label.text()
        assert (
            mto_content_equal_by_kit(
                [{"title": "2225", "mark": "KSB", "diff": {"content_status": "not_compared"}}]
            )
            == {}
        )
        assert (
            mto_content_equal_by_kit(
                [
                    {
                        "title": "2225",
                        "mark": "KSB",
                        "diff": {"content_status": "content_equal"},
                    }
                ]
            )
            == {("2225", "ksb"): True}
        )
        assert window._tabs.currentWidget() is window._kits_tab
        assert window._tabs.indexOf(window._kits_tab) == 0
        assert window._tabs.tabText(window._tabs.indexOf(window._revision_matrix_tab)) == (
            "Ревизии MTO"
        )
        assert window._tabs.indexOf(window._revision_matrix_tab) == 1
        assert window._tabs.tabText(window._tabs.indexOf(window._mto_worklist_tab)) == (
            "MTO · Перечень"
        )
        assert window._tabs.indexOf(window._mto_worklist_tab) == 2
        assert window._tabs.tabText(window._tabs.indexOf(window._an_tab)) == "АН"
        assert window._tabs.indexOf(window._an_tab) == 3
        assert window._tabs.indexOf(window._an_tab) == (
            window._tabs.indexOf(window._mto_worklist_tab) + 1
        )
        assert window._tabs.tabText(window._tabs.indexOf(window._rd_dump_tab)) == "РД"
        assert window._tabs.indexOf(window._rd_dump_tab) == (
            window._tabs.indexOf(window._an_tab) + 1
        )
        assert window._tabs.indexOf(window._mto_readiness_tab) == (
            window._tabs.indexOf(window._rd_dump_tab) + 1
        )
        assert window._tabs.tabText(
            window._tabs.indexOf(window._approval_mail_tab)
        ) == "Письма о согласовании"
        assert (
            window._approval_mail_tab._de_sync_button.text() == SHEET_DE_SYNC_BUTTON
        )
        assert window._tabs.indexOf(window._approval_mail_tab) == (
            window._tabs.indexOf(window._documents_tab) - 1
        )
        journal_index = window._tabs.indexOf(window._issuance_journal_tab)
        assert window._tabs.tabText(journal_index) == "Выдача · Журнал"
        assert journal_index == window._tabs.indexOf(window._mto_readiness_tab) + 1
        assert journal_index == window._tabs.indexOf(window._approval_mail_tab) - 1
        journal_table = window._issuance_journal_tab.table()
        journal_headers = [
            journal_table.horizontalHeaderItem(index).text()
            for index in range(journal_table.columnCount())
        ]
        assert journal_headers == [
            "Титул",
            "Марка",
            "Источник",
            "Рев.",
            "Дата отпр.",
            "TRM",
            "Статус листа",
            "Наш статус",
            "Комментарий",
            "F",
            "РД",
            "Робот",
            "Авто МТО",
            "Вх.контр.",
            "TRM подтв.",
            "Примечание",
            "Сопоставление",
        ]
        assert journal_table.columnCount() == len(journal_headers)
        assert hasattr(window._issuance_journal_tab, "_decision")
        assert hasattr(window._issuance_journal_tab, "_add_button")
        assert window._issuance_journal_tab._add_button.text() == "Добавить строку…"
        assert hasattr(window._issuance_journal_tab, "open_add_row_dialog")
        assert hasattr(window, "_open_legalize_rd_dialog")
        worklist_headers = [
            window._mto_worklist_table.horizontalHeaderItem(index).text()
            for index in range(window._mto_worklist_table.columnCount())
        ]
        assert worklist_headers == [
            "Титул",
            "Марка",
            "Ревизия (F/РД)",
            "Этап ревизии",
            "Метки",
            "Связь F → РД → MTO",
            "Google",
            "AB",
            PIN_COLUMN_HEADER,
            AUTO_MTO_COMPARE_STATUS_HEADER,
            "Файл MTO",
            "Дата MTO",
            "Пакет",
            "Чего не хватает",
            "Проблемы",
            "Путь MTO",
        ]
        assert window._mto_worklist_table.columnCount() == 16
        assert worklist_headers[_COL_PIN] == PIN_COLUMN_HEADER
        assert worklist_headers[_COL_FILE] == "Файл MTO"
        an_headers = [
            window._an_table.horizontalHeaderItem(index).text()
            for index in range(window._an_table.columnCount())
        ]
        assert an_headers == [
            "Титул",
            "Марка",
            "Ревизия АН",
            "К согл. передаче",
            "vs Авто МТО",
            "vs MTO РД",
            "vs Робот",
            "vs Выдача",
            "vs F",
            "vs SQ",
            "Имя",
            "Дата",
            "Папка",
            "Путь",
        ]
        assert window._scan_thread is None
        assert window._an_scan_thread is None
        assert window._mto_table.rowCount() == 0
        assert window._history.rowCount() == 0
        assert window._collision_table.rowCount() == 0
        assert window._kits_table.rowCount() == 0
        assert window._doc_tree.topLevelItemCount() == 0
        assert "ревизия" in window._doc_tree.headerItem().text(0)
        history_headers = [
            window._history.horizontalHeaderItem(index).text()
            for index in range(window._history.columnCount())
        ]
        assert history_headers == list(_HISTORY_HEADERS)
        assert history_headers[6] == "Имя"
        assert history_headers[7] == "Дата"
        assert history_headers[8] == "Рев."
        assert history_headers[9] == "AB"
        assert history_headers[10] == "Рабочая"
        assert history_headers[11] == "Аннул."
        assert history_headers[12] == PIN_COLUMN_HEADER
        assert history_headers[13] == "Статус MTO"
        assert history_headers[_HISTORY_COL_PACKAGE] == "Путь комплекта"
        assert "Путь PDF" not in history_headers
        assert "Пути ред" not in history_headers
        assert window._history.columnCount() == 15
        assert _HISTORY_COL_PACKAGE == 14
        assert history_headers.index("Дата") == history_headers.index("Имя") + 1
        assert history_headers.index("Рабочая") == history_headers.index("AB") + 1
        assert history_headers.index("Аннул.") == history_headers.index("Рабочая") + 1
        assert history_headers.index(PIN_COLUMN_HEADER) == history_headers.index("Аннул.") + 1
        assert history_headers.index("Статус MTO") == history_headers.index("AB") + 4
        assert window._mto_no_as_build.text() == "Без as-build"
        assert window._kits_code_a.text() == "Код A"
        assert window._kits_tdo.text() == "Прошли ТДО"
        assert window._kits_no_as_build.text() == "Без as-build"
        assert window._kits_only_as_build.text() == "Только as-build"
        assert window._kits_mto_problems.text() == "Проблемы MTO"
        assert window._kits_an_closes.text() == "АН закрывает Авто МТО"
        assert window._kits_legend_button.text() == rd_window.KITS_PAINT_LEGEND_BUTTON
        legend = KitsPaintLegendDialog(window._status_colors, window)
        assert legend.windowTitle()
        legend_tables = legend.findChildren(QTableWidget)
        assert len(legend_tables) >= 5
        assert legend_tables[0].rowCount() == 2
        assert any(table.rowCount() >= 3 for table in legend_tables)
        legend.close()
        assert window._kits_de_sync_button.text() == SHEET_DE_SYNC_BUTTON
        de_sync_row = SheetDeSyncRow(
            title="1600",
            mark="SOT",
            rd_revision_text="04",
            d_now="01",
            d_next="04",
            e_now="старый",
            e_next="РД Согласовано",
            last_f_line="02.12.2024 код А на рев. 04",
            last_f_revision="04",
            last_f_stage="код А",
            write_d=True,
            write_e=True,
            job=JournalWriteJob(
                title="1600",
                mark="SOT",
                f_line="02.12.2024 код А на рев. 04",
                revision="04",
                stage="code_a",
            ),
            d_cell="Рев. 04",
            e_cell="РД Согласовано",
        )
        de_sync = SheetDeSyncDialog((de_sync_row,), window)
        assert de_sync.windowTitle()
        headers = [
            de_sync._table.horizontalHeaderItem(index).text()
            for index in range(de_sync._table.columnCount())
        ]
        assert headers[4] == "D сейчас"
        assert headers[5] == "D будет"
        assert de_sync._table.item(0, 4).text() == "01"
        assert de_sync._table.item(0, 5).text() == "04"
        assert (
            de_sync._table.item(0, 4).background().color().name().lower()
            == "#f3c2c2"
        )
        assert (
            de_sync._table.item(0, 5).background().color().name().lower()
            == "#e2f2e1"
        )
        assert de_sync._table.item(0, 1).background().style() == Qt.BrushStyle.NoBrush
        assert len(de_sync.selected_jobs()) == 1
        buttons = {
            button.text(): button for button in de_sync.findChildren(QPushButton)
        }
        buttons["Снять все"].click()
        assert de_sync.selected_jobs() == ()
        buttons["Выделить все"].click()
        assert de_sync.selected_jobs()[0].mark == "SOT"
        de_sync.close()
        assert window._kits_layout_button.text() == rd_window._KITS_LAYOUT_BUTTON
        assert window._kits_xlsx_button.text() == rd_window._KITS_XLSX_BUTTON
        assert not window._kits_no_as_build.isChecked()
        assert not window._kits_only_as_build.isChecked()
        window._kits_no_as_build.setChecked(True)
        assert not window._kits_only_as_build.isEnabled()
        window._kits_only_as_build.setChecked(True)
        assert not window._kits_no_as_build.isChecked()
        assert not window._kits_no_as_build.isEnabled()
        window._kits_only_as_build.setChecked(False)
        assert window._kits_no_as_build.isEnabled()
        kit_headers = [
            window._kits_table.horizontalHeaderItem(index).text()
            for index in range(window._kits_table.columnCount())
        ]
        assert kit_headers == list(_KITS_HEADERS)
        assert kit_headers[_KITS_COL_PIN] == PIN_COLUMN_HEADER
        assert kit_headers.index(PIN_COLUMN_HEADER) == 3
        assert kit_headers[2] == "Ок"
        assert kit_headers.index("Ок") == kit_headers.index("Марка") + 1
        assert kit_headers[_KITS_COL_ISSUANCE_REV] == "Выдача · рев."
        assert kit_headers[_KITS_COL_GOOGLE_REV] == "Google · рев."
        assert kit_headers[_KITS_COL_ROBOT_REV] == "Робот МТО · рев."
        assert kit_headers[4:14] == [
            "Выдача · рев.",
            "Google · рев.",
            "Робот МТО · рев.",
            "SQ · рев.",
            "РД · рев.",
            "Рабочая рев. РД",
            "MTO · рев.",
            "Авто МТО",
            AUTO_MTO_COMPARE_STATUS_HEADER,
            "АН МТО",
        ]
        kits_header = window._kits_table.horizontalHeader()
        assert kits_header.sortIndicatorSection() == 0
        assert kits_header.sortIndicatorOrder() == Qt.SortOrder.AscendingOrder
        assert kit_headers.index("Рабочая рев. РД") == kit_headers.index("РД · рев.") + 1
        assert kit_headers.index("MTO · рев.") == kit_headers.index("Рабочая рев. РД") + 1
        assert kit_headers.index("Авто МТО") == kit_headers.index("MTO · рев.") + 1
        assert AUTO_MTO_COMPARE_STATUS_HEADER in kit_headers
        assert kit_headers.index(AUTO_MTO_COMPARE_STATUS_HEADER) == (
            kit_headers.index("Авто МТО") + 1
        )
        assert kit_headers.index("АН МТО") == kit_headers.index(
            AUTO_MTO_COMPARE_STATUS_HEADER
        ) + 1
        assert kit_headers[_KITS_COL_AN] == "АН МТО"
        assert _KITS_COL_AN not in _KITS_REV_SOURCE_COLUMNS
        assert _KITS_COL_WORKING not in _KITS_REV_SOURCE_COLUMNS
        header_keys = [key for key, _table in window._header_settings_map()]
        assert "window/kits_header_v10" in header_keys
        assert "window/kits_header_v9" not in header_keys
        assert "window/kits_header_v8" not in header_keys
        assert "window/kits_header_v7" not in header_keys
        assert "window/kits_header_v6" not in header_keys
        assert "window/an_tab_header_v2" in header_keys
        assert "window/an_tab_header_v1" not in header_keys
        assert "window/history_header_v8" in header_keys
        assert "window/history_header_v7" not in header_keys
        kits_menu_src = inspect.getsource(window._popup_kits_context_menu)
        assert 'Показать в „АН“' in kits_menu_src
        assert "Открыть папку · АН" in kits_menu_src
        assert "Открыть смешанные папки" in kits_menu_src
        assert "exec_tracked_menu" in kits_menu_src
        assert "is_file()" not in kits_menu_src
        show_kits_src = inspect.getsource(window._show_kits_context_menu)
        assert "gui.show_kits_context_menu" in show_kits_src
        assert "currentRow()" in kits_menu_src
        an_filter_src = inspect.getsource(window._an_tab.set_kit_filter)
        assert "sortByColumn" in an_filter_src
        assert "_COL_DATE" in an_filter_src
        assert "Показать в Выдача · Журнал" in kits_menu_src
        assert "issuance=row.issuance" in show_kits_src
        tree_menu_src = inspect.getsource(window._show_document_tree_context_menu)
        assert "Пометить папку как рабочую" in tree_menu_src
        assert "Пометить папку как аннулированную" in tree_menu_src
        assert "Снять пометку" in tree_menu_src
        assert "exec_tracked_menu" in tree_menu_src
        usage = get_context_menu_usage()
        assert usage is not None
        hot_label = "Открыть содержащую папку · РД"
        cold_label = "Скрыть титул–марку (бан-фильтр)"
        for _ in range(3):
            usage.note(MENU_KITS, hot_label)
        usage.note(MENU_KITS, cold_label)
        assert usage.flush()
        probe = QMenu()
        hot_action = probe.addAction(hot_label)
        cold_action = probe.addAction(cold_label)
        apply_usage_styles(probe, MENU_KITS)
        assert hot_action.font().bold()
        assert not cold_action.font().bold()
        assert "Google · рев." in kit_headers
        assert "Google · рев. F" in kit_headers
        assert "Статус рассмотрения" in kit_headers
        assert "Статус согласования" in kit_headers
        assert (
            window._kits_table.contextMenuPolicy()
            == Qt.ContextMenuPolicy.CustomContextMenu
        )
        assert (
            window._kits_table.editTriggers()
            == QAbstractItemView.EditTrigger.DoubleClicked
        )
        assert isinstance(
            window._kits_table.itemDelegate(), _ReadOnlyCopyDelegate
        )
        assert (
            window._kits_package_table.editTriggers()
            == QAbstractItemView.EditTrigger.DoubleClicked
        )
        assert isinstance(
            window._kits_package_table.itemDelegate(), _ReadOnlyCopyDelegate
        )
        assert (
            window._mto_table.editTriggers()
            == QAbstractItemView.EditTrigger.NoEditTriggers
        )
        review_index = kit_headers.index("Статус рассмотрения")
        approval_index = kit_headers.index("Статус согласования")
        assert approval_index == review_index + 1
        assert hasattr(window, "_kits_card_approval")
        assert hasattr(window, "_kits_card_mto")
        pkg_header = window._kits_package_table.horizontalHeader()
        stretch_cols = {1, 4, 5}
        for index in range(window._kits_package_table.columnCount()):
            expected = (
                QHeaderView.ResizeMode.Stretch
                if index in stretch_cols
                else QHeaderView.ResizeMode.ResizeToContents
            )
            assert pkg_header.sectionResizeMode(index) == expected
        for attr in (
            "_kits_pkg_open_button",
            "_kits_pkg_copy_button",
            "_kits_pkg_jump_button",
            "_kits_pkg_handoff_button",
        ):
            assert hasattr(window, attr)
        assert hasattr(window, "_kits_detail_tabs")
        assert window._kits_detail_tabs.count() == 3
        assert window._kits_detail_tabs.tabText(0) == "Подсказки"
        assert window._kits_detail_tabs.tabText(1) == "Карточка"
        assert window._kits_detail_tabs.tabText(2) == "Текст"
        assert window._kits_detail_tabs.currentIndex() == 0
        assert (
            window._kits_splitter.orientation() == Qt.Orientation.Vertical
        )
        assert window._kits_splitter.widget(0) is window._kits_table
        assert window._kits_splitter.widget(1) is window._kits_detail_tabs
        assert window._kits_splitter.childrenCollapsible() is False
        assert window._kits_detail_tabs.minimumHeight() == 80
        assert hasattr(window, "_kits_detail_placement_button")
        assert (
            window._kits_detail_tabs.cornerWidget()
            is window._kits_detail_placement_button
        )
        assert window._kits_detail_placement_button.toolTip() == "Панель справа"
        window._kits_detail_placement_button.click()
        assert (
            window._kits_splitter.orientation() == Qt.Orientation.Horizontal
        )
        assert window._kits_detail_tabs.minimumWidth() == 140
        assert window._kits_detail_placement_button.toolTip() == "Панель снизу"
        assert (
            window._settings.value("window/kits_detail_placement") == "right"
        )
        window._kits_detail_placement_button.click()
        assert (
            window._kits_splitter.orientation() == Qt.Orientation.Vertical
        )
        assert (
            window._settings.value("window/kits_detail_placement") == "bottom"
        )
        assert hasattr(window, "_kits_tooltips")
        assert hasattr(window, "_kits_tips_click_filter")
        assert "ячейке таблицы" in window._kits_tooltips.placeholderText()
        assert window._status_colors_action.text() == "Цвета статусов…"
        assert window._customer_pi_action.text() == "База заказчика…"
        assert all("КСБ ИД" not in header for header in kit_headers)
        for table in (
            window._mto_table,
            window._kits_table,
            window._collision_table,
            window._history,
        ):
            header = table.horizontalHeader()
            assert header.sectionsMovable()
            assert header.isFirstSectionMovable()
        matrix_header = window._revision_matrix_table.horizontalHeader()
        assert not matrix_header.sectionsMovable()
        dummy_cells = [
            SimpleNamespace(
                title="2225",
                mark="KSB",
                revision_text=rev,
                pipeline_status="code_a",
                letters="A",
                problem_kinds_json="[]",
                is_as_build=False,
                is_current=rev == "03",
                is_current_ifc=rev == "02",
                has_mto=True,
            )
            for rev in ("01", "02", "03", "04", "0-AN02")
        ]
        window._revision_matrix_tab.set_cells(
            dummy_cells,
            palette=window._status_colors,
            is_banned=lambda _title, _mark: False,
        )
        window._tabs.setCurrentWidget(window._revision_matrix_tab)
        window.resize(720, 480)
        app.processEvents()
        matrix = window._revision_matrix_table
        matrix_header = matrix.horizontalHeader()
        assert matrix.columnCount() == 11
        assert matrix.horizontalHeaderItem(2).text() == "В папку робота"
        assert matrix.horizontalHeaderItem(3).text() == "РД · рев."
        assert matrix.horizontalHeaderItem(4).text() == "Авто МТО"
        assert matrix.horizontalHeaderItem(5).text() == AUTO_MTO_COMPARE_STATUS_HEADER
        assert (
            matrix.horizontalScrollBarPolicy()
            == Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        assert (
            matrix_header.sectionResizeMode(0)
            == QHeaderView.ResizeMode.Interactive
        )
        assert (
            matrix_header.sectionResizeMode(2)
            == QHeaderView.ResizeMode.Interactive
        )
        assert (
            matrix_header.sectionResizeMode(3)
            == QHeaderView.ResizeMode.Interactive
        )
        assert (
            matrix_header.sectionResizeMode(4)
            == QHeaderView.ResizeMode.Interactive
        )
        assert (
            matrix_header.sectionResizeMode(5)
            == QHeaderView.ResizeMode.Interactive
        )
        assert (
            matrix_header.sectionResizeMode(6)
            == QHeaderView.ResizeMode.Stretch
        )
        assert (
            matrix_header.sectionResizeMode(10)
            == QHeaderView.ResizeMode.Stretch
        )
        assert matrix_header.length() <= matrix.viewport().width() + 2
        assert window._revision_matrix_tab._rule_combo.count() == 4
        assert "export_missing" in window._status_colors
        assert "export_add" in window._status_colors
        assert window._revision_matrix_tab._export_button.text().startswith(
            "Копировать видимые"
        )
        assert window._ban_action.text() == "Забаненные титулы"
        assert window._tabs.tabText(window._tabs.indexOf(window._kits_tab)) == (
            "Комплекты (0 компл., 0 тит.)"
        )
        assert window._kits_progress.text() == (
            "Ок: 0 / 0 · Совпадает: 0 / 0 · Код A: 0 · ТДО: 0 · "
            "Нет в РД: 0 · Проблемы MTO: 0"
        )
        assert window._tabs.tabText(window._tabs.indexOf(window._collision_tab)) == (
            "Коллизии (0)"
        )
        kit_row = KitMatrixRow(
            title="5850",
            mark="SKUD",
            title_system="5850-SKUD",
            rd=SourceKitSnapshot(),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.GOOGLE_ONLY,
        )
        window._kit_rows = (kit_row,)
        window._kits_table.insertRow(0)
        item = QTableWidgetItem(kit_row.title)
        item.setData(Qt.ItemDataRole.UserRole, kit_row)
        window._kits_table.setItem(0, 0, item)
        window._kit_pipelines[kit_identity_key("5850", "SKUD")] = KitPipelineRow(
            title="5850",
            mark="SKUD",
            status=KitPipelineStatus.NOT_UPLOADED.value,
            code=None,
        )
        window._apply_kits_filter()
        window._update_kits_tab_label()
        assert not window._kits_table.isRowHidden(0)
        assert window._select_kit_row("5850", "SKUD")
        assert window._kits_table.currentRow() == 0
        assert window._tabs.tabText(window._tabs.indexOf(window._kits_tab)) == (
            "Комплекты (1 компл., 1 тит.)"
        )
        assert window._kits_progress.text() == (
            "Ок: 0 / 1 · Совпадает: 0 / 1 · Код A: 0 · ТДО: 0 · "
            "Нет в РД: 1 · Проблемы MTO: 0"
        )

        pair, added = window._ban_store.add("5850", "SKUD", "аннулирован")
        assert added
        assert window._ban_store.contains("5850", "skud")
        window._apply_ban_visibility()
        assert window._ban_action.text() == "Забаненные титулы (1)"
        assert window._kits_table.isRowHidden(0)
        assert window._tabs.tabText(window._tabs.indexOf(window._kits_tab)) == (
            "Комплекты (0 компл., 0 тит.)"
        )
        assert window._kits_progress.text() == (
            "Ок: 0 / 0 · Совпадает: 0 / 0 · Код A: 0 · ТДО: 0 · "
            "Нет в РД: 0 · Проблемы MTO: 0"
        )
        window._all_records = [
            _stub_record(
                21,
                path=_pipeline_rd_path(
                    rd_root,
                    title="2225",
                    mark="KSB",
                    folder=_issued_nn("2225", "KSB"),
                    filename="AGCC.287-2225-KSB.OD-0001_01_RU.pdf",
                ),
                title="2225",
                mark="KSB",
            ),
            _stub_record(
                22,
                path=_pipeline_rd_path(
                    rd_root,
                    title="5850",
                    mark="SKUD",
                    folder=_issued_nn("5850", "SKUD"),
                    filename="AGCC.287-5850-SKUD.OD-0001_01_RU.pdf",
                ),
                title="5850",
                mark="SKUD",
            ),
            _stub_record(
                23,
                path=_pipeline_rd_path(
                    rd_root,
                    title="8950",
                    mark="POS1",
                    folder="01_рев.0_замечания_AGCC.287-8950-POS1",
                    filename="AGCC.287-8950-POS1.OD-0001.pdf",
                ),
                title="8950",
                mark="POS1",
            ),
        ]
        ksb_kit = KitMatrixRow(
            title="2225",
            mark="KSB",
            title_system="2225-KSB",
            rd=SourceKitSnapshot(present=True),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.GAP_ROBOT,
        )
        assert Path(window._rd_folder_for_kit(ksb_kit)) == rd_root / "2225" / "06_KSB"
        stray_path = _pipeline_rd_path(
            rd_root,
            title="2612",
            mark="KSB",
            folder=_issued_nn(
                "2612", "KSB", sequence=9, revision="01-AN01"
            ),
            filename="AGCC.287-2225-KSB.WIR-0010_01-AN01_RU.pdf",
        )
        window._all_records = [
            *window._all_records,
            _stub_record(
                25,
                path=stray_path,
                title="2225",
                mark="KSB",
                transfer_sequence=9,
                transfer_name="09_рев.01-AN01",
            ),
        ]
        mixed_rescan = [
            Path(item) for item in window._rd_rescan_subtrees_for_kit(ksb_kit)
        ]
        assert rd_root / "2612" / "06_KSB" in mixed_rescan
        assert rd_root not in mixed_rescan
        assert [Path(item) for item in window._mixed_title_folders_for_kit(ksb_kit)] == [
            rd_root / "2612" / "06_KSB"
        ]
        assert [
            Path(item) for item in window._mixed_title_open_folders_for_kit(ksb_kit)
        ] == [Path(stray_path).parent]
        window._all_records = window._all_records[:-1]
        missing_kit = KitMatrixRow(
            title="9999",
            mark="POS",
            title_system="9999-POS",
            rd=SourceKitSnapshot(),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.GOOGLE_ONLY,
        )
        assert Path(window._rd_folder_for_kit(missing_kit)) == rd_root / "9999" / "POS"
        assert Path(window._rd_rescan_subtree_for_kit(missing_kit)) == (
            rd_root / "9999"
        )
        gate_kit = KitMatrixRow(
            title="9000",
            mark="KSB",
            title_system="9000-KSB",
            rd=SourceKitSnapshot(
                present=True,
                paths=(
                    str(
                        rd_root
                        / "9000"
                        / "03_Для передачи"
                        / "05_рев.01"
                        / "PDF"
                        / "AGCC.287-9000-KSB.OD-0001_01_RU.pdf"
                    ),
                ),
            ),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.GAP_RD,
        )
        window._all_records = [
            *window._all_records,
            _stub_record(
                24,
                path=gate_kit.rd.paths[0],
                title="9000",
                mark="KSB",
            ),
        ]
        assert Path(window._rd_folder_for_kit(gate_kit)) == rd_root / "9000" / "KSB"
        assert Path(window._rd_rescan_subtree_for_kit(gate_kit)) == (
            rd_root / "9000"
        )
        window._all_records = window._all_records[:-1]
        window.config = replace(window.config, skip_dirs=("Замечания",))
        window.config = replace(window.config, skip_dirs=("Замечания",))
        window._contour_records = None
        window._assign_catalog_records(list(window._all_records))
        window._revision_matrix_tab.set_cells(
            [
                _dummy_matrix_cell("2225", "KSB"),
                _dummy_matrix_cell("5850", "SKUD"),
                _dummy_matrix_cell("1111", "ZZZ"),
                _dummy_matrix_cell("8950", "POS1"),
            ],
            palette=window._status_colors,
            is_banned=lambda title, mark: window._is_banned_pair(title, mark),
            allowed_kits=window._document_tree_kit_identities(),
        )
        assert _visible_matrix_kits(window._revision_matrix_table) == [
            ("2225", "KSB")
        ]
        # tempfile lives under ...\Temp\...; packaged tokens "tmp"/"temp"
        # would hide every fixture path. Idle skip is empty on purpose.
        window.config = replace(window.config, skip_dirs=())
        window._all_records = []

        dialog = BannedTitlesDialog(window._ban_store, window)
        dialog.show()
        app.processEvents()
        assert dialog._table.rowCount() == 1
        dialog.close()
        app.processEvents()
        assert window._skip_edit_button.text() == "Skip-папки…"
        assert window._skip_prune_button.text() == "Убрать skip из дерева"
        skip_dialog = SkipDirsDialog(window._skip_store, window)
        skip_dialog.show()
        app.processEvents()
        assert skip_dialog._list.count() >= 1
        skip_dialog.close()
        app.processEvents()
        assert window._doc_show_folder.text() == "Папка NN"
        assert window._doc_show_review.text() == "Статус Google"
        assert window._doc_show_mto_status.text() == "Статус MTO"
        assert window._doc_show_mto_status.isChecked()
        assert window._doc_show_as_build.text() == "As-build"
        for box in (
            window._doc_show_mto_status,
            window._doc_show_folder,
            window._doc_show_review,
            window._doc_show_current,
            window._doc_show_mto,
            window._doc_show_working,
            window._doc_show_as_build,
        ):
            box.blockSignals(True)
            box.setChecked(False)
            box.blockSignals(False)
        window._doc_show_date.blockSignals(True)
        window._doc_show_date.setChecked(True)
        window._doc_show_date.blockSignals(False)

        rec_a = _stub_record(
            1,
            path=_pipeline_rd_path(
                rd_root,
                title="9192",
                mark="POS",
                folder=_issued_nn("9192", "POS"),
                filename="AGCC.287-9192-POS.OD-0001_01_RU.pdf",
            ),
            title="9192",
            mark="POS",
        )
        rec_b = _stub_record(
            2,
            path=_pipeline_rd_path(
                rd_root,
                title="9192",
                mark="KSB",
                folder=_issued_nn("9192", "KSB"),
                filename="AGCC.287-9192-KSB.OD-0001_01_RU.pdf",
            ),
            title="9192",
            mark="KSB",
        )
        rec_skip = _stub_record(
            4,
            path=_pipeline_rd_path(
                rd_root,
                title="8950",
                mark="POS1",
                folder="01_рев.0_замечания_AGCC.287-8950-POS1",
                filename="AGCC.287-8950-POS1.OD-0001.pdf",
            ),
            title="8950",
            mark="POS1",
        )
        window.config = replace(window.config, skip_dirs=("Замечания",))
        window._all_records = [rec_a, rec_b, rec_skip]
        window._record_by_id = {
            rec_a.id: rec_a,
            rec_b.id: rec_b,
            rec_skip.id: rec_skip,
        }
        window._rebuild_document_tree()
        titles = [
            window._doc_tree.topLevelItem(index).text(0)
            for index in range(window._doc_tree.topLevelItemCount())
        ]
        assert titles == ["9192"]
        with patch.object(
            rd_window, "bundle_documents", wraps=rd_window.bundle_documents
        ) as mock_bundle:
            window._rebuild_document_tree()
            assert mock_bundle.call_count == 0
            window._assign_catalog_records(list(window._all_records))
            window._rebuild_document_tree()
            assert mock_bundle.call_count == 1
        window.config = replace(window.config, skip_dirs=())
        window._all_records = [rec_a, rec_b]
        window._record_by_id = {rec_a.id: rec_a, rec_b.id: rec_b}
        window._rebuild_document_tree()

        rec_nn4 = _stub_record(
            31,
            path=_pipeline_rd_path(
                rd_root,
                title="8950",
                mark="POS4",
                folder="04_Рев.0_AGCC.287-8950-POS4",
                filename="AGCC.287-8950-POS4.BOE-0001_0_RU.pdf",
            ),
            title="8950",
            mark="POS4",
            revision="0",
            core_stem="AGCC.287-8950-POS4.BOE-0001",
            transfer_sequence=4,
            transfer_name="04_Рев.0_AGCC.287-8950-POS4",
        )
        rec_nn2 = _stub_record(
            32,
            path=_pipeline_rd_path(
                rd_root,
                title="8950",
                mark="POS4",
                folder="02_Рев.0_AGCC.287-8950-POS4",
                filename="AGCC.287-8950-POS4.OD-0001_0_RU.pdf",
            ),
            title="8950",
            mark="POS4",
            revision="0",
            core_stem="AGCC.287-8950-POS4.OD-0001",
            transfer_sequence=2,
            transfer_name="02_Рев.0_AGCC.287-8950-POS4",
        )
        rec_nn1 = _stub_record(
            33,
            path=_pipeline_rd_path(
                rd_root,
                title="8950",
                mark="POS4",
                folder="01_AGCC.287-8950-POS4.MTO-0001_0_RU",
                filename="AGCC.287-8950-POS4.MTO-0001_0_RU.xlsx",
                media="DWG",
            ),
            title="8950",
            mark="POS4",
            revision="0",
            core_stem="AGCC.287-8950-POS4.MTO-0001",
            file_kind=FileKind.MTO_XLSX.value,
            transfer_sequence=1,
            transfer_name="01_AGCC.287-8950-POS4.MTO-0001_0_RU",
        )
        rec_nn_loose = _stub_record(
            34,
            path=str(
                rd_root
                / "8950"
                / "POS4"
                / "DWG"
                / "AGCC.287-8950-POS4.BOM-0001_0_RU.dwg"
            ),
            title="8950",
            mark="POS4",
            revision="0",
            core_stem="AGCC.287-8950-POS4.BOM-0001",
            file_kind=FileKind.SOURCE_EDITABLE.value,
        )
        saved_records = list(window._all_records)
        saved_by_id = dict(window._record_by_id)
        window._all_records = [rec_nn4, rec_nn2, rec_nn1, rec_nn_loose]
        window._record_by_id = {record.id: record for record in window._all_records}
        window._rebuild_document_tree()
        title_8950 = window._find_tree_child(window._doc_tree, "8950")
        assert title_8950 is not None
        mark_pos4 = window._find_tree_child(title_8950, "POS4")
        assert mark_pos4 is not None
        nn_keys = []
        for index in range(mark_pos4.childCount()):
            child = mark_pos4.child(index)
            bundles = child.data(0, Qt.ItemDataRole.UserRole) or ()
            nn_keys.append(bundles[0].folder_key)
        assert nn_keys[0] == "01_agcc.287-8950-pos4.mto-0001_0_ru"
        assert nn_keys[1] == "02_рев.0_agcc.287-8950-pos4"
        assert nn_keys[2] == "04_рев.0_agcc.287-8950-pos4"
        assert nn_keys[3].endswith("\\pos4\\dwg")
        window._all_records = saved_records
        window._record_by_id = saved_by_id
        window._rebuild_document_tree()

        rec_c = _stub_record(
            3,
            path=_pipeline_rd_path(
                rd_root,
                title="5850",
                mark="SKUD",
                folder=_issued_nn("5850", "SKUD", revision="02"),
                filename="AGCC.287-5850-SKUD.OD-0001_02_RU.pdf",
            ),
            title="5850",
            mark="SKUD",
            revision="02",
        )
        window._ban_store.remove("5850", "SKUD")
        window._all_records = [*window._all_records, rec_c]
        window._record_by_id[rec_c.id] = rec_c
        window._rebuild_document_tree()
        original_rebuild = window._rebuild_document_tree
        rebuild_calls: list[object] = []

        def _count_rebuild(*args: object, **kwargs: object) -> None:
            rebuild_calls.append(True)
            original_rebuild(*args, **kwargs)

        window._rebuild_document_tree = _count_rebuild  # type: ignore[method-assign]
        try:
            window._jump_to_kit_documents("5850", "SKUD")
            app.processEvents()
            assert len(rebuild_calls) == 0
            assert window._tabs.currentWidget() is window._documents_tab
            assert window._doc_filter.text() == "5850-SKUD"
            current = window._doc_tree.currentItem()
            assert current is not None
            assert current.text(0) == "02"
            dump = window._documents_tree_handoff_text(current)
            assert "RD Catalog · Передать роботу" in dump
            assert "Уровень: ревизия (папка NN)" in dump
            assert "Лейбл сейчас: 02" in dump
            assert "Титул: 5850" in dump
            assert "Марка: SKUD" in dump
        finally:
            window._rebuild_document_tree = original_rebuild
        rec_c.data["transfer_name"] = "07_рев.02_AGCC.287-5850-SKUD"
        window._doc_show_folder.blockSignals(True)
        window._doc_show_folder.setChecked(True)
        window._doc_show_folder.blockSignals(False)
        window._rebuild_document_tree()
        window._jump_to_kit_documents("5850", "SKUD")
        current = window._doc_tree.currentItem()
        assert current is not None
        assert current.text(0) == "02 · 07_рев.02_AGCC.287-5850-SKUD"
        folder_dump = window._documents_tree_handoff_text(current)
        assert "Лейбл сейчас: 02 · 07_рев.02_AGCC.287-5850-SKUD" in folder_dump
        assert "Папка NN: 07_рев.02_AGCC.287-5850-SKUD" in folder_dump
        window._copy_robot_handoff(folder_dump)
        copied = QApplication.clipboard().text()
        assert copied == folder_dump
        window._doc_show_folder.blockSignals(True)
        window._doc_show_folder.setChecked(False)
        window._doc_show_folder.blockSignals(False)
        window._rebuild_document_tree()

        titles_after_jump = [
            window._doc_tree.topLevelItem(index).text(0)
            for index in range(window._doc_tree.topLevelItemCount())
        ]
        assert titles_after_jump == ["5850", "9192"]
        window._doc_filter.blockSignals(True)
        window._doc_filter.setText("9192")
        window._doc_filter.blockSignals(False)
        assert window._doc_tree.currentItem() is not None
        window._apply_document_tree_filter()
        hidden_by_title = {
            window._doc_tree.topLevelItem(index).text(0): window._doc_tree.topLevelItem(
                index
            ).isHidden()
            for index in range(window._doc_tree.topLevelItemCount())
        }
        assert hidden_by_title == {"5850": True, "9192": False}
        assert window._doc_tree.topLevelItemCount() == 2
        assert window._doc_tree.currentItem() is None
        assert window._history.rowCount() == 0
        window._doc_filter.setText("5850")
        assert window._doc_filter_timer.isActive()
        window._doc_filter_timer.stop()
        window._apply_document_tree_filter()
        hidden_by_title = {
            window._doc_tree.topLevelItem(index).text(0): window._doc_tree.topLevelItem(
                index
            ).isHidden()
            for index in range(window._doc_tree.topLevelItemCount())
        }
        assert hidden_by_title == {"5850": False, "9192": True}
        assert window._doc_tree.currentItem() is None

        yellow = _LAYOUT_MISMATCH_BG.name().casefold()
        window.config = replace(window.config, skip_dirs=())
        canonical = _stub_record(
            20,
            path=str(
                rd_root
                / "9192"
                / "11_POS"
                / "Для передачи"
                / "01_рев.01"
                / "PDF"
                / "AGCC.287-9192-POS.OD-0001_01_RU.pdf"
            ),
            title="9192",
            mark="POS",
            revision="01",
        )
        outside = _stub_record(
            21,
            path=str(
                rd_root
                / "9192"
                / "11_POS"
                / "черновик"
                / "AGCC.287-9192-POS.OD-0001_02_RU.pdf"
            ),
            title="9192",
            mark="POS",
            revision="02",
        )
        window._doc_filter.setText("")
        window._assign_catalog_records([canonical, outside])
        window._rebuild_document_tree()
        title_item = window._doc_tree.topLevelItem(0)
        assert title_item is not None
        assert title_item.text(0) == "9192"
        assert title_item.background(0).color().name().casefold() == yellow
        mark_item = title_item.child(0)
        assert mark_item is not None
        assert mark_item.background(0).color().name().casefold() == yellow
        saw_canonical = saw_outside = False
        for index in range(mark_item.childCount()):
            child = mark_item.child(index)
            assert child is not None
            bundles = child.data(0, Qt.ItemDataRole.UserRole) or ()
            paths = " ".join(
                record.path
                for bundle in bundles
                for record in (
                    ([bundle.pdf] if bundle.pdf is not None else [])
                    + list(bundle.editables)
                )
            )
            bg = child.background(0).color().name().casefold()
            if "Для передачи" in paths:
                assert bg != yellow
                saw_canonical = True
            else:
                assert bg == yellow
                saw_outside = True
        assert saw_canonical and saw_outside

        canonical_item = None
        for index in range(mark_item.childCount()):
            child = mark_item.child(index)
            bundles = child.data(0, Qt.ItemDataRole.UserRole) or ()
            paths = " ".join(
                record.path
                for bundle in bundles
                for record in (
                    ([bundle.pdf] if bundle.pdf is not None else [])
                    + list(bundle.editables)
                )
            )
            if "Для передачи" in paths:
                canonical_item = child
                break
        assert canonical_item is not None
        window._doc_tree.setCurrentItem(canonical_item)
        app.processEvents()
        assert window._history.rowCount() >= 1
        package_cell = window._history.item(0, _HISTORY_COL_PACKAGE)
        assert package_cell is not None
        expected_package = issued_package_dir(canonical.path)
        assert package_cell.text() == expected_package
        assert package_cell.text().endswith("01_рев.01")
        assert r"\PDF" not in package_cell.text()
        assert not package_cell.text().lower().endswith(".pdf")
        assert (
            window._doc_tree.contextMenuPolicy()
            == Qt.ContextMenuPolicy.CustomContextMenu
        )
        assert window._rd_folder_for_tree_item(canonical_item) == expected_package
        mark_folder = window._rd_folder_for_tree_item(mark_item)
        assert mark_folder.endswith("11_POS")
        window._copy_history_name()
        assert QApplication.clipboard().text() == Path(canonical.path).name
        window._copy_history_package_path()
        assert QApplication.clipboard().text() == expected_package
        name_cell = window._history.item(0, _HISTORY_COL_NAME)
        assert name_cell is not None
        assert name_cell.text() == Path(canonical.path).name
        empty_pin = window._history.item(0, _HISTORY_COL_PIN)
        assert empty_pin is not None
        assert empty_pin.text() == ""

        ns_old = 1_704_067_200_000_000_000
        ns_new = 1_723_824_000_000_000_000
        canonical.data["mtime_ns"] = ns_old
        companion = _stub_record(
            22,
            path=str(
                Path(canonical.path).parent
                / "AGCC.287-9192-POS.WIR-0001_01_RU.pdf"
            ),
            title="9192",
            mark="POS",
            revision="01",
            core_stem="AGCC.287-9192-POS.WIR-0001",
            mtime_ns=ns_new,
        )
        window._all_records = [canonical, outside, companion]
        window._record_by_id = {
            record.id: record for record in window._all_records
        }
        window._rebuild_document_tree()
        title_item = window._doc_tree.topLevelItem(0)
        mark_item = title_item.child(0)
        canonical_item = None
        for index in range(mark_item.childCount()):
            child = mark_item.child(index)
            bundles = child.data(0, Qt.ItemDataRole.UserRole) or ()
            paths = " ".join(
                record.path
                for bundle in bundles
                for record in (
                    ([bundle.pdf] if bundle.pdf is not None else [])
                    + list(bundle.editables)
                )
            )
            if "Для передачи" in paths:
                canonical_item = child
                break
        assert canonical_item is not None
        window._doc_tree.setCurrentItem(canonical_item)
        app.processEvents()
        assert window._history.rowCount() >= 2
        latest_hex = _LATEST_FILE_DATE_BG.name().casefold()
        old_text = format_file_save_date(ns_old)
        new_text = format_file_save_date(ns_new)
        painted = 0
        saw_old = False
        for row in range(window._history.rowCount()):
            date_cell = window._history.item(row, _HISTORY_COL_DATE)
            assert date_cell is not None
            bg = date_cell.background().color().name().casefold()
            if date_cell.text() == new_text:
                assert bg == latest_hex
                assert "Самая свежая дата в этой папке." in date_cell.toolTip()
                painted += 1
            elif date_cell.text() == old_text:
                assert bg != latest_hex
                saw_old = True
        assert painted == 1
        assert saw_old
        date_header = window._history.horizontalHeaderItem(_HISTORY_COL_DATE)
        assert date_header is not None
        assert "бледно-голубым" in date_header.toolTip()
        canonical.data["mtime_ns"] = 0
        window._all_records = [canonical, outside]
        window._record_by_id = {canonical.id: canonical, outside.id: outside}
        window._rebuild_document_tree()
        title_item = window._doc_tree.topLevelItem(0)
        mark_item = title_item.child(0)
        canonical_item = None
        for index in range(mark_item.childCount()):
            child = mark_item.child(index)
            bundles = child.data(0, Qt.ItemDataRole.UserRole) or ()
            paths = " ".join(
                record.path
                for bundle in bundles
                for record in (
                    ([bundle.pdf] if bundle.pdf is not None else [])
                    + list(bundle.editables)
                )
            )
            if "Для передачи" in paths:
                canonical_item = child
                break
        assert canonical_item is not None
        window._doc_tree.setCurrentItem(canonical_item)
        app.processEvents()
        assert window._history.item(0, _HISTORY_COL_NAME).text() == Path(
            canonical.path
        ).name

        pin_file = "AGCC.287-9192-POS.MTO-0001_01_RU.xlsx"
        pin_path = rf"C:\rd\pack\{pin_file}"
        rule_path = r"C:\rd\pack\rule.xlsx"
        save_export_pins(
            window.config,
            (
                ExportPin(
                    title="9192",
                    mark="POS",
                    package_path=r"C:\rd\pack",
                    file_path=pin_path,
                    revision_text="01",
                    evidence="abc",
                    created_at="2026-09-10T12:00:00+00:00",
                ),
            ),
        )
        window._revision_matrix_tab.apply_export_selections(
            (
                ExportSelection(
                    title="9192",
                    mark="POS",
                    rule="latest_issued",
                    source_path=pin_path,
                    source_revision_text="01",
                    package_path=r"C:\rd\pack",
                    origin="pin",
                    state="add",
                    destination_path="",
                    existing_target_path="",
                    confidence="high",
                    warnings=(),
                    rule_path=rule_path,
                ),
            )
        )
        window._refresh_history()
        app.processEvents()
        current_pin = window._history.item(0, _HISTORY_COL_PIN)
        assert current_pin is not None
        assert current_pin.text() == pin_file
        assert r"C:\rd\pack" in current_pin.toolTip()
        assert pin_path in current_pin.toolTip()
        assert "2026-09-10T12:00:00+00:00" in current_pin.toolTip()
        assert window._history.item(0, _HISTORY_COL_NAME).text() == Path(
            canonical.path
        ).name
        assert window._history.item(0, _HISTORY_COL_PACKAGE).text() == expected_package

        window._revision_matrix_tab.apply_export_selections(
            (
                ExportSelection(
                    title="9192",
                    mark="POS",
                    rule="latest_issued",
                    source_path=pin_path,
                    source_revision_text="01",
                    package_path=r"C:\rd\pack",
                    origin="pin_stale",
                    state="pin_stale",
                    destination_path="",
                    existing_target_path="",
                    confidence="high",
                    warnings=(),
                    rule_path=rule_path,
                ),
            )
        )
        window._refresh_history()
        app.processEvents()
        stale_pin = window._history.item(0, _HISTORY_COL_PIN)
        assert stale_pin is not None
        assert stale_pin.text() == f"{pin_file} (устарел)"
        assert pin_path in stale_pin.toolTip()
        assert rule_path in stale_pin.toolTip()
        stale_hex = QColor(
            color_for(window._status_colors, "export_pin_stale")
        ).name().casefold()
        assert stale_pin.background().color().name().casefold() == stale_hex
        assert window._history.item(0, _HISTORY_COL_NAME).text() == Path(
            canonical.path
        ).name
        assert window._history.item(0, _HISTORY_COL_PACKAGE).text() == expected_package

        populated = KitMatrixRow(
            title="9192",
            mark="POS",
            title_system="9192-POS",
            rd=SourceKitSnapshot(present=True, revision_text="02"),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(present=True, revision_text="01"),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.GAP_ROBOT,
        )
        original_rebuild = window._rebuild_kit_rows
        window._rebuild_kit_rows = lambda: None
        try:
            window._kit_rows = (populated,)
            window._refresh_kits_table()
            app.processEvents()
            kits_pin = window._kits_table.item(0, _KITS_COL_PIN)
            assert kits_pin is not None
            assert kits_pin.text() == f"{pin_file} (устарел)"
            assert window._kits_table.item(0, _KITS_COL_RD_REV).text().startswith("02")
            assert window._kits_table.item(0, _KITS_COL_SQ_REV).text().startswith("01")
            assert kits_pin.background().color().name().casefold() == stale_hex
            window._revision_matrix_tab.apply_export_selections(
                (
                    ExportSelection(
                        title="9192",
                        mark="POS",
                        rule="latest_issued",
                        source_path=pin_path,
                        source_revision_text="01",
                        package_path=r"C:\rd\pack",
                        origin="pin",
                        state="add",
                        destination_path="",
                        existing_target_path="",
                        confidence="high",
                        warnings=(),
                        rule_path=rule_path,
                    ),
                )
            )
            window._kit_rows = (populated,)
            window._refresh_kits_table()
            app.processEvents()
            assert window._kits_table.item(0, _KITS_COL_PIN).text() == pin_file
            assert window._kits_table.item(0, _KITS_COL_RD_REV).text().startswith("02")
            assert window._kits_table.item(0, _KITS_COL_SQ_REV).text().startswith("01")
            save_export_pins(window.config, ())
            window._revision_matrix_tab.apply_export_selections(())
            window._kit_rows = (populated,)
            window._refresh_kits_table()
            app.processEvents()
            assert window._kits_table.item(0, _KITS_COL_PIN).text() == ""
            assert window._kits_table.item(0, _KITS_COL_RD_REV).text().startswith("02")
            assert window._kits_table.item(0, _KITS_COL_SQ_REV).text().startswith("01")
        finally:
            window._rebuild_kit_rows = original_rebuild

        pin_worklist_row = _synthetic_worklist_row(
            title="9192",
            mark="POS",
            mto_path=r"C:\rd\unique-mto.xlsx",
            mto_mtime_ns=1_704_067_200_000_000_000,
            package_label="01_рев.01",
            package_path=r"C:\rd\pack",
            package_count=1,
        )
        window._mto_worklist_rows = [pin_worklist_row]
        save_export_pins(window.config, ())
        window._revision_matrix_tab.apply_export_selections(())
        window._refresh_mto_worklist(reload=False)
        app.processEvents()
        worklist_table = window._mto_worklist_table
        worklist_date = worklist_table.item(0, _COL_DATE).text()
        _assert_worklist_pin_and_tail(
            worklist_table,
            pin_text="",
            mto_path=r"C:\rd\unique-mto.xlsx",
            package_label="01_рев.01",
            date_text=worklist_date,
        )
        save_export_pins(
            window.config,
            (
                ExportPin(
                    title="9192",
                    mark="POS",
                    package_path=r"C:\rd\pack",
                    file_path=pin_path,
                    revision_text="01",
                    evidence="abc",
                    created_at="2026-09-10T12:00:00+00:00",
                ),
            ),
        )
        window._revision_matrix_tab.apply_export_selections(
            (
                ExportSelection(
                    title="9192",
                    mark="POS",
                    rule="latest_issued",
                    source_path=pin_path,
                    source_revision_text="01",
                    package_path=r"C:\rd\pack",
                    origin="pin",
                    state="add",
                    destination_path="",
                    existing_target_path="",
                    confidence="high",
                    warnings=(),
                    rule_path=rule_path,
                ),
            )
        )
        window._on_export_pins_changed()
        app.processEvents()
        _assert_worklist_pin_and_tail(
            worklist_table,
            pin_text=pin_file,
            mto_path=r"C:\rd\unique-mto.xlsx",
            package_label="01_рев.01",
            date_text=worklist_date,
            tooltip_parts=(r"C:\rd\pack", pin_path, "2026-09-10T12:00:00+00:00"),
        )
        window._revision_matrix_tab.apply_export_selections(
            (
                ExportSelection(
                    title="9192",
                    mark="POS",
                    rule="latest_issued",
                    source_path=pin_path,
                    source_revision_text="01",
                    package_path=r"C:\rd\pack",
                    origin="pin_stale",
                    state="pin_stale",
                    destination_path="",
                    existing_target_path="",
                    confidence="high",
                    warnings=(),
                    rule_path=rule_path,
                ),
            )
        )
        window._on_export_pins_changed()
        app.processEvents()
        _assert_worklist_pin_and_tail(
            worklist_table,
            pin_text=f"{pin_file} (устарел)",
            mto_path=r"C:\rd\unique-mto.xlsx",
            package_label="01_рев.01",
            date_text=worklist_date,
            stale=True,
            palette=window._status_colors,
            tooltip_parts=(pin_path, rule_path),
        )
        save_export_pins(window.config, ())
        window._revision_matrix_tab.apply_export_selections(())
        window._on_export_pins_changed()
        app.processEvents()
        _assert_worklist_pin_and_tail(
            worklist_table,
            pin_text="",
            mto_path=r"C:\rd\unique-mto.xlsx",
            package_label="01_рев.01",
            date_text=worklist_date,
        )

        gone = _stub_record(
            22,
            path=str(rd_root / "9192" / "old_place" / "AGCC.287-9192-POS.OD-0009.pdf"),
            title="9192",
            mark="POS",
            revision="09",
        )
        gone = replace(gone, present=False)
        window._all_records = [canonical, outside, gone]
        window._record_by_id[gone.id] = gone
        window._rebuild_document_tree()
        mark_after = window._doc_tree.topLevelItem(0).child(0)
        revision_texts = [
            mark_after.child(index).text(0)
            for index in range(mark_after.childCount())
        ]
        assert not any("09" in text for text in revision_texts)

        painted = QTableWidgetItem("02-AN01")
        _paint_revision_cell(painted, True)
        assert painted.background().color().name().casefold() == (
            _REV_MATCH_BG.name().casefold()
        )
        _paint_robot_origin_cell(painted, matched=True, hint="робот")
        assert painted.background().color().name().casefold() == (
            _ROBOT_ORIGIN_BG.name().casefold()
        )
        assert painted.toolTip() == "робот"

        package = issued_package_dir(r"C:\rd\9110\05_pack\PDF\file.pdf")
        assert package.endswith("05_pack")
        assert "PDF" not in package.split("\\")[-1]
        user_pdf = (
            r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\8950\22_POS2\Для передачи"
            r"\12_рев.AN01_AGCC.287-8950-POS2\PDF"
            r"\AGCC.287-8950-POS2.OD-0001_03-AN01_RU.pdf"
        )
        user_kit = (
            r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\8950\22_POS2\Для передачи"
            r"\12_рев.AN01_AGCC.287-8950-POS2"
        )
        assert issued_package_dir(user_pdf) == user_kit

        history_header = window._history.horizontalHeader()
        wanted_widths = (72, 48, 44, 52, 61, 73, 111, 90, 47, 40, 56, 48, 140, 72, 317)
        assert history_header.count() == len(wanted_widths)
        for index, width in enumerate(wanted_widths):
            history_header.resizeSection(index, width)
        app.processEvents()
        snapshot = tuple(
            history_header.sectionSize(index)
            for index in range(history_header.count())
        )
        window.refresh()
        app.processEvents()
        after_refresh = tuple(
            window._history.horizontalHeader().sectionSize(index)
            for index in range(window._history.columnCount())
        )
        assert after_refresh == snapshot

        worklist_tab = MtoWorklistTab()
        palette = default_palette()
        pin_file_name = "AGCC.287-9192-POS.MTO-0001_01_RU.xlsx"
        pin_file_path = rf"C:\rd\pack\{pin_file_name}"
        pin_rule_path = r"C:\rd\pack\rule.xlsx"
        pin_row = _synthetic_worklist_row(
            title="9192",
            mark="POS",
            mto_path=r"C:\rd\unique-mto.xlsx",
            mto_mtime_ns=1_704_067_200_000_000_000,
            package_label="01_рев.01",
            package_path=r"C:\rd\pack",
            package_count=1,
        )
        worklist_tab.set_rows(
            (pin_row,),
            palette=palette,
            is_banned=lambda _title, _mark: False,
        )
        app.processEvents()
        pin_table = worklist_tab.table()
        pin_date = pin_table.item(0, _COL_DATE).text()
        assert pin_table.columnCount() == 16
        _assert_worklist_pin_and_tail(
            pin_table,
            pin_text="",
            mto_path=r"C:\rd\unique-mto.xlsx",
            package_label="01_рев.01",
            date_text=pin_date,
        )
        stored_pin = ExportPin(
            title="9192",
            mark="POS",
            package_path=r"C:\rd\pack",
            file_path=pin_file_path,
            revision_text="01",
            evidence="abc",
            created_at="2026-09-10T12:00:00+00:00",
        )
        worklist_tab.set_rows(
            (pin_row,),
            palette=palette,
            is_banned=lambda _title, _mark: False,
            pin_view_for=_pin_view_lookup(
                stored_pin, origin="pin", rule_path=pin_rule_path
            ),
        )
        app.processEvents()
        _assert_worklist_pin_and_tail(
            pin_table,
            pin_text=pin_file_name,
            mto_path=r"C:\rd\unique-mto.xlsx",
            package_label="01_рев.01",
            date_text=pin_date,
            tooltip_parts=(
                r"C:\rd\pack",
                pin_file_path,
                "2026-09-10T12:00:00+00:00",
            ),
        )
        worklist_tab.set_rows(
            (pin_row,),
            palette=palette,
            is_banned=lambda _title, _mark: False,
            pin_view_for=_pin_view_lookup(
                stored_pin,
                origin="pin_stale",
                rule_path=pin_rule_path,
            ),
        )
        app.processEvents()
        _assert_worklist_pin_and_tail(
            pin_table,
            pin_text=f"{pin_file_name} (устарел)",
            mto_path=r"C:\rd\unique-mto.xlsx",
            package_label="01_рев.01",
            date_text=pin_date,
            stale=True,
            palette=palette,
            tooltip_parts=(pin_file_path, pin_rule_path),
        )
        code_a_rows = (
            _synthetic_worklist_row(
                title="1111",
                mark="KSB",
                revision_text="01",
                status="code_a",
                letters="A",
            ),
            _synthetic_worklist_row(
                title="2222",
                mark="POS",
                revision_text="01",
                status="tdo_review",
                letters="T",
            ),
            _synthetic_worklist_row(
                title="3333",
                mark="SOT",
                revision_text="01",
                status="no_mto",
                letters="–",
                has_mto=False,
                mto_path="",
                gap_kind="no_mto_file",
            ),
            _synthetic_worklist_row(
                title="4444",
                mark="PD",
                revision_text="01",
                status="code_a",
                letters="A",
                has_mto=False,
                mto_path="",
                gap_kind="no_package",
            ),
        )
        worklist_tab.set_rows(
            code_a_rows,
            palette=palette,
            is_banned=lambda _title, _mark: False,
        )
        app.processEvents()
        assert _visible_row_count(worklist_tab.table()) == 4
        worklist_tab._code_a.setChecked(True)
        app.processEvents()
        assert _visible_row_count(worklist_tab.table()) == 2
        missing_rows = (
            _synthetic_worklist_row(
                title="5555",
                mark="KSB",
                revision_text="01",
                status="tdo_review",
                letters="T",
            ),
            _synthetic_worklist_row(
                title="6666",
                mark="POS",
                revision_text="01",
                status="no_mto",
                letters="–",
                has_mto=False,
                mto_path="",
                gap_kind="no_mto_file",
            ),
            _synthetic_worklist_row(
                title="7777",
                mark="SOT",
                revision_text="01",
                status="not_uploaded",
                letters="",
                has_mto=False,
                mto_path="",
                gap_kind="no_package",
            ),
            _synthetic_worklist_row(
                title="8888",
                mark="PD",
                revision_text="02",
                status="code_a",
                letters="A",
            ),
        )
        worklist_tab._code_a.setChecked(False)
        worklist_tab.set_rows(
            missing_rows,
            palette=palette,
            is_banned=lambda _title, _mark: False,
        )
        worklist_tab._missing.setChecked(True)
        app.processEvents()
        visible_missing = []
        table = worklist_tab.table()
        for index in range(table.rowCount()):
            if table.isRowHidden(index):
                continue
            payload = worklist_tab._row_at(index)
            assert payload is not None
            visible_missing.append(payload)
        assert len(visible_missing) == 2
        assert {row.gap_kind for row in visible_missing} == {
            "no_mto_file",
            "no_package",
        }
        assert all(row.gap_kind for row in visible_missing)

        worklist_tab._missing.setChecked(False)
        gap_rows = (
            _synthetic_worklist_row(
                title="1010",
                mark="KSB",
                revision_text="01",
                status="tdo_review",
                letters="T",
                has_mto=False,
                mto_path="",
                gap_kind="no_rd",
                in_google=True,
                in_issuance=True,
                has_f_status=True,
            ),
            _synthetic_worklist_row(
                title="2020",
                mark="POS",
                revision_text="01",
                status="tdo_review",
                letters="T",
                in_google=False,
                in_issuance=False,
                has_f_status=True,
            ),
            _synthetic_worklist_row(
                title="3030",
                mark="SOT",
                revision_text="01",
                status="tdo_review",
                letters="T",
                in_google=True,
                in_issuance=False,
                has_f_status=False,
            ),
            _synthetic_worklist_row(
                title="3030",
                mark="PD",
                revision_text="01",
                status="tdo_review",
                letters="T",
                in_google=True,
                in_issuance=True,
                has_f_status=True,
                problem_kinds=("liquidity",),
            ),
        )
        worklist_tab.set_rows(
            gap_rows,
            palette=palette,
            is_banned=lambda _title, _mark: False,
        )
        app.processEvents()
        assert worklist_tab._counts_scope.text() == (
            "Титулы: 3 / 3 · Комплекты: 4 / 4 · Строки: 4 / 4"
        )
        assert worklist_tab._counts_breakdown.text() == (
            "Код A: 0 · ТДО: 4 · Нет MTO: 1 · Нет в РД: 1 · "
            "Нет в Google: 1 · Без статуса F: 1 · Проблемы: 1"
        )
        gap_table = worklist_tab.table()
        google_texts = {
            gap_table.item(index, _COL_GOOGLE).text()
            for index in range(gap_table.rowCount())
            if gap_table.item(index, _COL_GOOGLE) is not None
        }
        assert google_texts == {"нет комплекта", "нет статуса", "есть"}
        problem_texts = {
            gap_table.item(index, _COL_PROBLEMS).text()
            for index in range(gap_table.rowCount())
            if gap_table.item(index, _COL_PROBLEMS) is not None
            and gap_table.item(index, _COL_PROBLEMS).text()
        }
        assert problem_texts == {"Ликвидность"}

        def _visible_payloads() -> list[MtoWorklistRow]:
            payloads: list[MtoWorklistRow] = []
            table = worklist_tab.table()
            for index in range(table.rowCount()):
                if table.isRowHidden(index):
                    continue
                payload = worklist_tab._row_at(index)
                assert payload is not None
                payloads.append(payload)
            return payloads

        worklist_tab._no_rd.setChecked(True)
        app.processEvents()
        visible_no_rd = _visible_payloads()
        assert len(visible_no_rd) == 1
        assert visible_no_rd[0].gap_kind == "no_rd"
        assert visible_no_rd[0].title == "1010"
        worklist_tab._no_rd.setChecked(False)

        worklist_tab._no_google.setChecked(True)
        app.processEvents()
        visible_no_google = _visible_payloads()
        assert len(visible_no_google) == 1
        assert not visible_no_google[0].in_google
        assert not visible_no_google[0].in_issuance
        assert visible_no_google[0].title == "2020"
        worklist_tab._no_google.setChecked(False)

        worklist_tab._no_f_status.setChecked(True)
        app.processEvents()
        visible_no_f = _visible_payloads()
        assert len(visible_no_f) == 1
        assert not visible_no_f[0].has_f_status
        assert visible_no_f[0].title == "3030"
        assert visible_no_f[0].mark == "SOT"
        worklist_tab._no_f_status.setChecked(False)
        worklist_tab._filter.setText("liquidity")
        app.processEvents()
        visible_liq = _visible_payloads()
        assert len(visible_liq) == 1
        assert visible_liq[0].mark == "PD"
        worklist_tab._filter.setText("ликвидность")
        app.processEvents()
        assert len(_visible_payloads()) == 1
        worklist_tab._filter.setText("")
        app.processEvents()
        assert worklist_tab._counts_scope.text() == (
            "Титулы: 3 / 3 · Комплекты: 4 / 4 · Строки: 4 / 4"
        )

        worklist_settings = QSettings(
            str(root / "settings" / "mto_worklist_filters.ini"),
            QSettings.Format.IniFormat,
        )
        worklist_tab._code_a.setChecked(True)
        worklist_tab._tdo.setChecked(True)
        worklist_tab._current.setChecked(False)
        worklist_tab._current_ifc.setChecked(True)
        worklist_tab._no_as_build.setChecked(True)
        worklist_tab._only_as_build.setChecked(False)
        worklist_tab._missing.setChecked(True)
        worklist_tab._no_rd.setChecked(True)
        worklist_tab._no_google.setChecked(True)
        worklist_tab._no_f_status.setChecked(False)
        worklist_tab._problems.setChecked(False)
        worklist_tab._filter.setText("KSB-test")
        worklist_tab.save_filters(worklist_settings)
        worklist_settings.sync()
        worklist_tab._code_a.setChecked(False)
        worklist_tab._tdo.setChecked(False)
        worklist_tab._current_ifc.setChecked(False)
        worklist_tab._no_as_build.setChecked(False)
        worklist_tab._missing.setChecked(False)
        worklist_tab._no_rd.setChecked(False)
        worklist_tab._no_google.setChecked(False)
        worklist_tab._no_f_status.setChecked(True)
        worklist_tab._filter.setText("")
        worklist_tab.restore_filters(worklist_settings)
        assert worklist_tab._code_a.isChecked()
        assert worklist_tab._tdo.isChecked()
        assert not worklist_tab._current.isChecked()
        assert worklist_tab._current_ifc.isChecked()
        assert worklist_tab._no_as_build.isChecked()
        assert not worklist_tab._only_as_build.isChecked()
        assert worklist_tab._missing.isChecked()
        assert worklist_tab._no_rd.isChecked()
        assert worklist_tab._no_google.isChecked()
        assert not worklist_tab._no_f_status.isChecked()
        assert not worklist_tab._problems.isChecked()
        assert worklist_tab._filter.text() == "KSB-test"
        worklist_tab.deleteLater()
        app.processEvents()

        window.close()
        app.processEvents()

        reopened = CatalogWindow(config, database)
        reopened.show()
        app.processEvents()
        _wait_catalog_startup(reopened, app)
        restored = tuple(
            reopened._history.horizontalHeader().sectionSize(index)
            for index in range(reopened._history.columnCount())
        )
        assert restored == snapshot
        assert reopened._history.horizontalHeaderItem(
            _HISTORY_COL_PACKAGE
        ).text() == "Путь комплекта"
        assert (
            reopened._kits_splitter.orientation() == Qt.Orientation.Vertical
        )
        reopened.close()
        app.processEvents()
    print("RD catalog GUI smoke: OK")


if __name__ == "__main__":
    main()
