"""Qt-free checks for rd_catalog.monitor_views join cells."""

from __future__ import annotations

import ast
import json
import sys
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.config import load_config
from rd_catalog.db import CatalogDatabase, KitPipelineRow
from rd_catalog.google_sheet_links import SheetLinkContext
from rd_catalog.issuance_review import IssuanceJournalRow, send_identity_fingerprint
from rd_catalog.customer_pi_auto_mto import AutoMtoFile
from rd_catalog.kits import (
    GoogleKit,
    IssuanceKit,
    KitMatrixRow,
    KitSummary,
    SourceKitSnapshot,
    kit_identity_key,
    last_event_parts,
    parse_history_line,
    parse_sheet_revision,
)
from rd_catalog.monitor_views import (
    AN_AGREED_HEADER,
    AN_HEADERS,
    AUTO_MTO_AHEAD_FILL,
    AUTO_MTO_COMPARE_STATUS_HEADER,
    ISSUANCE_SHEET_LABEL,
    KITS_DE_SHEET_LABEL,
    KITS_F_SHEET_LABEL,
    KITS_OK_HEADER,
    KITS_TDO_STATUSES,
    KITS_TIPS_CTRL_CLICK_HINT,
    KITS_WORKING_REV_TOOLTIP,
    OFFICIAL_FOLDER_MTO_MISSING,
    RD_AB_SUFFIX,
    RD_MISSING_TRANSFER_KEY,
    RD_MISSING_TRANSFER_NOTE,
    MIXED_TITLES_KEY,
    REV_DIFF_FILL,
    REV_MATCH_FILL,
    ROBOT_ORIGIN_FILL,
    TDO_STATUSES,
    CatalogMonitor,
    MonitorCell,
    MtoKitFlags,
    format_kits_progress_stats,
    format_kits_row_tooltips,
    format_kits_tips_pane,
    wrap_kits_tips_text,
    tooltip_has_fs_path,
    join_haystack,
    jsonify_value,
    kit_an_targets,
    kits_paint_legend,
    kits_progress_stats,
    kits_tooltip_header_offset,
    kits_tooltip_section_start,
    load_catalog_monitor,
    WEB_MONITOR_SHEETS,
    monitor_to_json,
    build_kits_monitor_row,
    effective_kit_summary,
    format_kits_google_f_rev,
    kits_official_folder_mto_text,
    official_rd_rev_text,
    pipeline_approval_tooltip,
    pipeline_review_tooltip,
    rd_mto_overlay_by_kit,
    auto_mto_rd_target,
    rd_rev_review_lag_notes,
    revision_or_dash,
    worklist_google_cell,
    _build_journal_row,
    _heatmap_rd_rev_cell,
)
from rd_catalog.pipeline import MtoWorklistRow
from rd_catalog.status_colors import color_for, default_palette


def _empty_config(root: Path):
    runtime = root / "runtime"
    runtime.mkdir()
    override = root / "config.json"
    override.write_text(
        json.dumps(
            {
                "rd_root": str(root / "rd"),
                "sq_root": str(root / "sq"),
                "robot_root": str(root / "robot"),
                "runtime_dir": str(runtime),
                "db_path": str(runtime / "catalog.sqlite"),
                "robot_flat_structure": True,
                "skip_dirs": ["old"],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    return load_config(override)


def _worklist_row(**overrides) -> MtoWorklistRow:
    payload = {
        "title": "9110",
        "mark": "KSB",
        "revision_text": "01",
        "status": "not_uploaded",
        "letters": "",
        "is_as_build": False,
        "is_current": False,
        "is_current_ifc": False,
        "has_mto": False,
        "in_google": False,
        "in_issuance": False,
        "has_f_status": False,
    }
    payload.update(overrides)
    return MtoWorklistRow(**payload)


def _google_kit(
    *,
    title: str = "3140",
    mark: str = "KSB2",
    sheet: str = "01-AN02",
    lines: tuple[str, ...] = (),
) -> GoogleKit:
    events = tuple(parse_history_line(line) for line in lines)
    last = events[-1] if events else None
    revision, appendix = parse_sheet_revision(sheet)
    return GoogleKit(
        title=title,
        mark=mark,
        mark_raw=mark,
        title_system=f"{title}-{mark}",
        sheet_revision=revision,
        sheet_appendix=appendix,
        sheet_revision_text=sheet,
        status_sheet="",
        comment_raw="\n".join(lines),
        events=events,
        last_event=last,
        row_index=1,
    )


def _paint_kits(row: KitMatrixRow, pipeline: KitPipelineRow | None, **kwargs):
    payload = {
        "pipeline": pipeline,
        "palette": default_palette(),
        "pins": {},
        "overlay_mto": {},
        "selection": None,
        "comparison": None,
        "auto_files": (),
        "an_files": (),
        "mto_flags": MtoKitFlags(),
        "mto_content_equal": None,
        "ifc_by_kit": {},
        "excluded_sends": (),
    }
    payload.update(kwargs)
    return build_kits_monitor_row(row, **payload)


def _issuance_kit(
    *,
    title: str = "2210",
    mark: str = "KSB",
    rev: str = "01-AN01",
) -> IssuanceKit:
    revision, appendix = parse_sheet_revision(rev)
    return IssuanceKit(
        title=title,
        mark=mark,
        mark_raw=mark,
        title_system=f"{title}-{mark}",
        revision=revision,
        appendix=appendix,
        revision_text=rev,
        status="",
        send_date="",
        send_date_sortable="",
        send_transmittal="",
        incoming_control_date="",
        incoming_control_date_sortable="",
        confirm_transmittal="",
        note_raw="",
        row_index=1,
    )


def _rd_snap(rev: str = "01-AN01", *, mto: str | None = None) -> SourceKitSnapshot:
    mto_text = rev if mto is None else mto
    revision, appendix = parse_sheet_revision(rev)
    mto_rev, mto_app = parse_sheet_revision(mto_text)
    return SourceKitSnapshot(
        present=True,
        revision=revision,
        appendix=appendix,
        revision_text=rev,
        mto_revision=mto_rev,
        mto_appendix=mto_app,
        mto_revision_text=mto_text,
    )


def _src_snap(rev: str = "01-AN01") -> SourceKitSnapshot:
    revision, appendix = parse_sheet_revision(rev)
    return SourceKitSnapshot(
        present=True,
        revision=revision,
        appendix=appendix,
        revision_text=rev,
    )


def _ok_pipeline(**overrides) -> KitPipelineRow:
    payload = {
        "title": "2210",
        "mark": "KSB",
        "status": "agreed",
        "code": "A",
        "code_stale": False,
        "official_revision_text": "01-AN01",
        "code_revision_text": "01-AN01",
    }
    payload.update(overrides)
    return KitPipelineRow(**payload)


def _auto_file(
    *,
    title: str = "1600",
    mark: str = "SOS",
    rev: str = "02-AN01",
) -> AutoMtoFile:
    stem = f"AGCC.287-{title}-{mark}.MTO-0001"
    return AutoMtoFile(
        title=title,
        mark=mark,
        spec=stem,
        rd_revision=rev,
        relpath=f"{title}/{mark}/{stem}_{rev}_RU.xlsx",
        row_count=1,
        fingerprint="mto",
    )


def _ok_kit_row(**overrides) -> KitMatrixRow:
    payload = {
        "title": "2210",
        "mark": "KSB",
        "title_system": "2210-KSB",
        "rd": _rd_snap(),
        "robot": _src_snap(),
        "sq": SourceKitSnapshot(),
        "google": _google_kit(
            title="2210",
            mark="KSB",
            sheet="01",
            lines=("04.12.2025 код А на рев. 01-AN01",),
        ),
        "issuance": _issuance_kit(),
        "flags": (),
        "summary": KitSummary.ALIGNED,
    }
    payload.update(overrides)
    return KitMatrixRow(**payload)


class MonitorViewsTests(unittest.TestCase):
    def test_import_has_no_pyside(self) -> None:
        from rd_catalog import monitor_views

        source = Path(monitor_views.__file__).read_text(encoding="utf-8")
        ast.parse(source)
        self.assertNotIn("PySide", source)
        self.assertNotIn("PyQt", source)

    def test_empty_initialized_sqlite(self) -> None:
        with tempfile.TemporaryDirectory(prefix="rd_monitor_") as temp:
            root = Path(temp)
            config = _empty_config(root)
            database = CatalogDatabase(config.db_path)
            database.initialize()
            monitor = load_catalog_monitor(database, config)
            self.assertIsInstance(monitor, CatalogMonitor)
            self.assertEqual(monitor.kits, ())
            json.dumps(monitor_to_json(monitor))

    def test_web_sheets_skip_heatmap_and_journal(self) -> None:
        with tempfile.TemporaryDirectory(prefix="rd_monitor_web_") as temp:
            root = Path(temp)
            config = _empty_config(root)
            database = CatalogDatabase(config.db_path)
            database.initialize()
            monitor = load_catalog_monitor(
                database, config, sheets=WEB_MONITOR_SHEETS
            )
            self.assertEqual(monitor.heatmap.rows, ())
            self.assertEqual(monitor.heatmap.revision_columns, ())
            self.assertEqual(monitor.worklist, ())
            self.assertEqual(monitor.journal, ())
            self.assertEqual(monitor.mto_readiness, ())
            self.assertEqual(monitor.collisions, ())
            self.assertEqual(monitor.an_rows, ())

    def test_official_rd_rev_ab_suffix(self) -> None:
        pipeline = KitPipelineRow(
            title="2210",
            mark="KSB",
            status="agreed",
            working_revision_text="01-AN02",
            official_revision_text="01-AN01",
            review_as_build=True,
        )
        row = KitMatrixRow(
            title="2210",
            mark="KSB",
            title_system="2210-KSB",
            rd=SourceKitSnapshot(present=True, revision_text="01", as_build=False),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.ALIGNED,
        )
        text = official_rd_rev_text(row, pipeline)
        self.assertEqual(
            text, f"01 · 01-AN01 {RD_MISSING_TRANSFER_NOTE}"
        )
        painted = _heatmap_rd_rev_cell(
            row.rd, "", pipeline=pipeline, palette=default_palette()
        )
        self.assertEqual(painted.text, text)
        self.assertEqual(painted.palette_key, RD_MISSING_TRANSFER_KEY)
        self.assertTrue(painted.fill)

        missing = KitMatrixRow(
            title="2210",
            mark="KSB",
            title_system="2210-KSB",
            rd=SourceKitSnapshot(),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.GAP_RD,
        )
        self.assertEqual(
            official_rd_rev_text(missing, pipeline),
            f"01-AN01 {RD_MISSING_TRANSFER_NOTE}",
        )
        mismatch = KitMatrixRow(
            title="2210",
            mark="KSB",
            title_system="2210-KSB",
            rd=SourceKitSnapshot(present=True, revision_text="01", as_build=False),
            robot=SourceKitSnapshot(present=True, revision_text="01-AN01"),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.REV_MISMATCH,
        )
        self.assertEqual(
            effective_kit_summary(mismatch, pipeline), KitSummary.GAP_RD
        )
        self.assertEqual(
            effective_kit_summary(
                KitMatrixRow(
                    title="2210",
                    mark="KSB",
                    title_system="2210-KSB",
                    rd=SourceKitSnapshot(),
                    robot=SourceKitSnapshot(),
                    sq=SourceKitSnapshot(),
                    google=None,
                    issuance=None,
                    flags=(),
                    summary=KitSummary.GOOGLE_ONLY,
                ),
                pipeline,
            ),
            KitSummary.GOOGLE_ONLY,
        )
        painted_row = build_kits_monitor_row(
            mismatch,
            pipeline=pipeline,
            palette=default_palette(),
            pins={},
            overlay_mto={},
            selection=None,
            comparison=None,
            auto_files=(),
            an_files=(),
            mto_flags=MtoKitFlags(),
            mto_content_equal=None,
            ifc_by_kit={},
            excluded_sends=(),
        )
        self.assertEqual(painted_row.cells["Сводка"].text, "Нет в РД")
        self.assertEqual(
            painted_row.cells["РД · рев."].text,
            f"01 · 01-AN01 {RD_MISSING_TRANSFER_NOTE}",
        )
        self.assertFalse(painted_row.rd_present)
        self.assertFalse(painted_row.summary_aligned)

        matched = KitPipelineRow(
            title="2210",
            mark="KSB",
            status="agreed",
            official_revision_text="01-AN01",
        )
        on_disk = KitMatrixRow(
            title="2210",
            mark="KSB",
            title_system="2210-KSB",
            rd=SourceKitSnapshot(
                present=True, revision_text="01-AN01", as_build=False
            ),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.ALIGNED,
        )
        self.assertEqual(official_rd_rev_text(on_disk, matched), "01-AN01")
        self.assertEqual(
            effective_kit_summary(on_disk, matched), KitSummary.ALIGNED
        )

        no_working = KitPipelineRow(
            title="9110",
            mark="KSB",
            status="agreed",
            official_revision_text="03",
        )
        as_build_row = KitMatrixRow(
            title="9110",
            mark="KSB",
            title_system="9110-KSB",
            rd=SourceKitSnapshot(present=True, revision_text="03", as_build=True),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.ALIGNED,
        )
        text_as_build = official_rd_rev_text(as_build_row, no_working)
        self.assertTrue(text_as_build.endswith(RD_AB_SUFFIX))
        self.assertIn("03", text_as_build)
        self.assertEqual(revision_or_dash(False, "01"), "—")
        self.assertEqual(revision_or_dash(True, ""), "есть")

        review = KitMatrixRow(
            title="3000",
            mark="KSB",
            title_system="3000-KSB",
            rd=SourceKitSnapshot(
                present=True, revision_text="0-AN01", as_build=False
            ),
            robot=SourceKitSnapshot(present=True, revision_text="0-AN01"),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.TRANSFER_REVIEW,
            transfer_review_notes=(
                "Согласованная передача:",
                "  MTO текущей NN 09 (19.05.2026) позже письма A (27.04.2026)",
            ),
        )
        review_pipeline = KitPipelineRow(
            title="3000",
            mark="KSB",
            status="agreed",
            official_revision_text="0-AN01",
        )
        painted_review = build_kits_monitor_row(
            review,
            pipeline=review_pipeline,
            palette=default_palette(),
            pins={},
            overlay_mto={},
            selection=None,
            comparison=None,
            auto_files=(),
            an_files=(),
            mto_flags=MtoKitFlags(),
            mto_content_equal=None,
            ifc_by_kit={},
            excluded_sends=(),
        )
        self.assertEqual(painted_review.cells["Сводка"].text, "Проверить передачи")
        self.assertEqual(painted_review.cells["РД · рев."].fill, REV_DIFF_FILL)
        self.assertIn("позже письма A", painted_review.cells["РД · рев."].tooltip)
        self.assertIn("позже письма A", painted_review.cells["Сводка"].tooltip)
        tips = painted_review.tooltips_text
        self.assertLess(
            tips.find("=== Сводка ==="),
            tips.find("=== РД · рев. ==="),
        )
        self.assertIn("позже письма A", tips)

    def test_mixed_titles_paints_rd_rev_and_summary(self) -> None:
        row = KitMatrixRow(
            title="2630",
            mark="KSB",
            title_system="2630-KSB",
            rd=SourceKitSnapshot(
                present=True, revision_text="01-AN01", as_build=True
            ),
            robot=SourceKitSnapshot(present=True, revision_text="01-AN01"),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.ALIGNED,
            mixed_title_notes=(
                "Смешанные титулы: файлы комплекта лежат в папке другого титула.",
                "  папка 2612 · 09_рев.AN-01 · WIR-0010.pdf",
            ),
        )
        pipeline = KitPipelineRow(
            title="2630",
            mark="KSB",
            status="agreed",
            official_revision_text="01-AN01",
        )
        painted = build_kits_monitor_row(
            row,
            pipeline=pipeline,
            palette=default_palette(),
            pins={},
            overlay_mto={},
            selection=None,
            comparison=None,
            auto_files=(),
            an_files=(),
            mto_flags=MtoKitFlags(),
            mto_content_equal=None,
            ifc_by_kit={},
            excluded_sends=(),
        )
        self.assertEqual(painted.cells["Сводка"].text, "Смешанные титулы")
        self.assertEqual(
            painted.cells["РД · рев."].palette_key, MIXED_TITLES_KEY
        )
        self.assertEqual(
            painted.cells["РД · рев."].fill,
            default_palette()[MIXED_TITLES_KEY],
        )
        self.assertIn("2612", painted.cells["РД · рев."].tooltip)
        self.assertIn("смешанные титулы", painted.haystack)
        self.assertIn("смешанные марки", painted.haystack)

    def test_rd_rev_yellow_when_disk_lags_letter_a_or_tdo(self) -> None:
        row = KitMatrixRow(
            title="3140",
            mark="KSB2",
            title_system="3140-KSB2",
            rd=SourceKitSnapshot(
                present=True,
                revision="01",
                appendix="01",
                revision_text="01-AN01",
                mto_revision="01",
                mto_appendix="01",
                mto_revision_text="01-AN01",
            ),
            robot=SourceKitSnapshot(
                present=True,
                revision="01",
                appendix="01",
                revision_text="01-AN01",
                mto_revision="01",
                mto_appendix="01",
                mto_revision_text="01-AN01",
            ),
            sq=SourceKitSnapshot(),
            google=_google_kit(sheet="01-AN02"),
            issuance=None,
            flags=(),
            summary=KitSummary.REV_MISMATCH,
        )
        stale_a = KitPipelineRow(
            title="3140",
            mark="KSB2",
            status="agreed",
            code="A",
            code_revision_text="01-AN02",
            code_date="10.09.2026",
            code_stale=True,
            official_revision_text="01-AN01",
            working_revision_text="01-AN02",
        )
        self.assertTrue(rd_rev_review_lag_notes(row, stale_a))
        painted = _paint_kits(row, stale_a, mto_content_equal=True)
        rd_cell = painted.cells["РД · рев."]
        self.assertEqual(rd_cell.text, "01-AN01")
        self.assertEqual(rd_cell.fill, REV_DIFF_FILL)
        self.assertNotEqual(rd_cell.fill, ROBOT_ORIGIN_FILL)
        self.assertIn("Письмо A на 01-AN02 (новее диска)", rd_cell.tooltip)
        self.assertIn("письмо a", painted.haystack)
        self.assertEqual(painted.cells["Робот МТО · рев."].fill, ROBOT_ORIGIN_FILL)
        self.assertTrue(painted.cells["Робот МТО · рев."].bold)

        matched_a = KitPipelineRow(
            title="3140",
            mark="KSB2",
            status="agreed",
            code="A",
            code_revision_text="01-AN01",
            official_revision_text="01-AN01",
        )
        self.assertEqual(rd_rev_review_lag_notes(row, matched_a), ())
        matched_paint = _paint_kits(row, matched_a, mto_content_equal=True)
        self.assertEqual(
            matched_paint.cells["РД · рев."].fill, ROBOT_ORIGIN_FILL
        )
        self.assertTrue(matched_paint.cells["Робот МТО · рев."].bold)

        tdo_row = KitMatrixRow(
            title="3140",
            mark="KSB2",
            title_system="3140-KSB2",
            rd=SourceKitSnapshot(
                present=True, revision_text="01-AN01", as_build=False
            ),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=_google_kit(
                sheet="01-AN02",
                lines=("10.09.2026 прошла ТДО рев.01-AN02",),
            ),
            issuance=None,
            flags=(),
            summary=KitSummary.REV_MISMATCH,
        )
        tdo_pipeline = KitPipelineRow(
            title="3140",
            mark="KSB2",
            status="agreed",
            code="A",
            code_revision_text="01-AN01",
            official_revision_text="01-AN01",
            working_revision_text="01-AN02",
        )
        notes = rd_rev_review_lag_notes(tdo_row, tdo_pipeline)
        self.assertTrue(any("ТДО" in item for item in notes))
        tdo_paint = _paint_kits(tdo_row, tdo_pipeline)
        self.assertEqual(tdo_paint.cells["РД · рев."].fill, REV_DIFF_FILL)
        self.assertIn("01-AN02", tdo_paint.cells["РД · рев."].tooltip)
        self.assertIn("ТДО", tdo_paint.cells["РД · рев."].tooltip)

        missing_row = KitMatrixRow(
            title="3140",
            mark="KSB2",
            title_system="3140-KSB2",
            rd=SourceKitSnapshot(
                present=True, revision_text="01", as_build=False
            ),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=_google_kit(sheet="01-AN02"),
            issuance=None,
            flags=(),
            summary=KitSummary.REV_MISMATCH,
        )
        missing_pipeline = KitPipelineRow(
            title="3140",
            mark="KSB2",
            status="agreed",
            code="A",
            code_revision_text="01-AN02",
            code_stale=True,
            official_revision_text="01-AN01",
        )
        missing_paint = _paint_kits(missing_row, missing_pipeline)
        self.assertEqual(
            missing_paint.cells["РД · рев."].palette_key,
            RD_MISSING_TRANSFER_KEY,
        )
        self.assertNotEqual(missing_paint.cells["РД · рев."].fill, REV_DIFF_FILL)

    def test_working_rev_column_only_when_ahead_of_official(self) -> None:
        row = KitMatrixRow(
            title="3860",
            mark="SOT",
            title_system="3860-SOT",
            rd=SourceKitSnapshot(present=True, revision_text="0-AN01"),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.ALIGNED,
        )
        equal = KitPipelineRow(
            title="3860",
            mark="SOT",
            status="agreed",
            official_revision_text="02",
            working_revision_text="02",
        )
        painted_equal = _paint_kits(row, equal)
        self.assertEqual(painted_equal.cells["Рабочая рев. РД"].text, "—")
        self.assertNotEqual(
            painted_equal.cells["Рабочая рев. РД"].palette_key, "working"
        )

        lower = KitPipelineRow(
            title="3860",
            mark="SOT",
            status="agreed",
            official_revision_text="02",
            working_revision_text="01",
        )
        painted_lower = _paint_kits(row, lower)
        self.assertEqual(painted_lower.cells["Рабочая рев. РД"].text, "—")

        ahead = KitPipelineRow(
            title="3860",
            mark="SOT",
            status="agreed",
            official_revision_text="01-AN01",
            working_revision_text="01-AN02",
        )
        painted_ahead = _paint_kits(row, ahead)
        self.assertEqual(painted_ahead.cells["Рабочая рев. РД"].text, "01-AN02")
        self.assertEqual(
            painted_ahead.cells["Рабочая рев. РД"].palette_key, "working"
        )
        self.assertNotIn("As-build", painted_ahead.cells["Рабочая рев. РД"].tooltip)

        as_build = KitPipelineRow(
            title="3860",
            mark="SOT",
            status="agreed",
            official_revision_text="01-AN01",
            working_revision_text="01-AN02",
            working_as_build=True,
        )
        painted_ab = _paint_kits(row, as_build)
        self.assertEqual(
            painted_ab.cells["Рабочая рев. РД"].text,
            f"01-AN02{RD_AB_SUFFIX}",
        )
        self.assertEqual(
            painted_ab.cells["Рабочая рев. РД"].palette_key, "working"
        )
        self.assertIn("As-build", painted_ab.cells["Рабочая рев. РД"].tooltip)
        self.assertIn(
            f"01-an02{RD_AB_SUFFIX.casefold()}",
            painted_ab.haystack,
        )

    def test_an_headers_include_agreed_column(self) -> None:
        self.assertEqual(AN_HEADERS[3], AN_AGREED_HEADER)
        self.assertEqual(AN_HEADERS[3], "К согл. передаче")
        row = KitMatrixRow(
            title="2230",
            mark="KSB",
            title_system="2230-KSB",
            rd=SourceKitSnapshot(present=True, revision_text="01-AN01"),
            robot=SourceKitSnapshot(present=True, revision_text="01-AN02"),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.GAP_RD,
        )
        pipeline = KitPipelineRow(
            title="2230",
            mark="KSB",
            status="agreed",
            code="A",
            code_revision_text="01-AN01",
            code_date="10.12.2025",
            official_revision_text="01-AN01",
            working_revision_text="01-AN02",
        )
        targets = kit_an_targets(
            row,
            auto_files=(),
            rd_path="",
            rd_mto_rev="",
            pipeline=pipeline,
        )
        self.assertEqual(targets.agreed, "01-AN01")
        self.assertEqual(targets.agreed_date, "10.12.2025")
        self.assertEqual(targets.robot, "01-AN02")

    def test_worklist_google_cell_labels(self) -> None:
        missing = worklist_google_cell(_worklist_row())
        self.assertEqual(missing.text, "нет комплекта")
        no_status = worklist_google_cell(
            _worklist_row(in_google=True, has_f_status=False)
        )
        self.assertEqual(no_status.text, "нет статуса")
        present = worklist_google_cell(
            _worklist_row(in_google=True, has_f_status=True)
        )
        self.assertEqual(present.text, "есть")

    def test_tdo_status_sets_differ(self) -> None:
        self.assertIn("agreed", KITS_TDO_STATUSES)
        self.assertNotIn("agreed", TDO_STATUSES)
        self.assertIn("code_a", TDO_STATUSES)
        self.assertNotIn("code_a", KITS_TDO_STATUSES)

    def test_join_haystack_skips_none(self) -> None:
        self.assertEqual(join_haystack("a", None, "", "b"), "a b")
        self.assertEqual(join_haystack(None, None), "")

    def test_kits_google_trm_href(self) -> None:
        links = SheetLinkContext(
            kits_spreadsheet_id="kits-id",
            kits_sheet_title="Контроль выдачи ",
            kits_sheet_id=77,
            issuance_spreadsheet_id="iss-id",
            issuance_sheet_title="Выдача РД ПД",
            issuance_sheet_id=None,
        )
        google = replace(
            _google_kit(title="2210", mark="KSB", sheet="01"),
            row_index=12,
        )
        issuance = replace(_issuance_kit(), row_index=40, send_transmittal="")
        painted = _paint_kits(
            _ok_kit_row(google=google, issuance=issuance),
            _ok_pipeline(),
            sheet_links=links,
        )
        trm = painted.cells["Google · TRM F"]
        self.assertTrue(trm.href.endswith("#gid=77&range=F12"), trm.href)
        self.assertIn("gid=77", trm.href)
        self.assertIn("Shift+клик", trm.tooltip)
        send = painted.cells["Выдача · TRM отпр."]
        self.assertIn("B40", send.href)
        self.assertNotIn("'", send.href)
        self.assertNotIn("!", send.href)
        self.assertEqual(painted.cells["Титул"].href, "")

    def test_journal_trm_href(self) -> None:
        links = SheetLinkContext(
            kits_spreadsheet_id="kits-id",
            kits_sheet_title="Контроль выдачи ",
            kits_sheet_id=None,
            issuance_spreadsheet_id="iss-id",
            issuance_sheet_title="Выдача РД ПД",
            issuance_sheet_id=None,
        )
        issuance = _issuance_kit(title="2869", mark="SOS", rev="01-AN02")
        issuance = replace(issuance, row_index=40, send_transmittal="TRM-2")
        payload = dict(
            title="2869",
            mark="SOS",
            kind="send",
            source="issuance",
            revision_text="01-AN02",
            send_date="15.09.2026",
            send_date_sortable="2026-09-15",
            send_transmittal="TRM-2",
            incoming_control_date="",
            incoming_control_date_sortable="",
            confirm_transmittal="",
            sheet_status="",
            note="",
            decision="active",
            comment=None,
            match_state="matched",
            identity_fingerprint=send_identity_fingerprint(issuance),
            evidence_fingerprint="",
            source_path="",
            path_key="",
            in_f=True,
            in_rd=True,
            in_robot=True,
            in_auto_mto=False,
            review_id=None,
            issuance_send_id=2,
            sheet_row_index=40,
        )
        painted = _build_journal_row(
            IssuanceJournalRow(**payload),
            issuance=issuance,
            sheet_links=links,
        )
        self.assertIn("B40", painted.cells["TRM"].href)
        self.assertNotIn("'", painted.cells["TRM"].href)
        self.assertNotIn("!", painted.cells["TRM"].href)
        self.assertIn("Shift+клик", painted.cells["TRM"].tooltip)
        self.assertEqual(painted.cells["Титул"].href, "")

    def test_journal_row_paints_kits_issuance(self) -> None:
        issuance = IssuanceKit(
            title="2869",
            mark="SOS",
            mark_raw="SOS",
            title_system="2869-SOS",
            revision="01",
            appendix="02",
            revision_text="01-AN02",
            status="",
            send_date="15.09.2026",
            send_date_sortable="2026-09-15",
            send_transmittal="TRM-2",
            incoming_control_date="",
            incoming_control_date_sortable="",
            confirm_transmittal="",
            note_raw="",
            row_index=20,
        )
        payload = dict(
            title="2869",
            mark="SOS",
            kind="send",
            source="issuance",
            revision_text="01-AN02",
            send_date="15.09.2026",
            send_date_sortable="2026-09-15",
            send_transmittal="TRM-2",
            incoming_control_date="",
            incoming_control_date_sortable="",
            confirm_transmittal="",
            sheet_status="",
            note="",
            decision="active",
            comment=None,
            match_state="matched",
            identity_fingerprint=send_identity_fingerprint(issuance),
            evidence_fingerprint="",
            source_path="",
            path_key="",
            in_f=True,
            in_rd=True,
            in_robot=True,
            in_auto_mto=False,
            review_id=None,
            issuance_send_id=2,
        )
        painted = _build_journal_row(IssuanceJournalRow(**payload), issuance=issuance)
        older = _build_journal_row(
            IssuanceJournalRow(
                **{
                    **payload,
                    "revision_text": "01-AN01",
                    "send_date": "22.12.2025",
                    "send_date_sortable": "2025-12-22",
                    "send_transmittal": "TRM-1",
                    "identity_fingerprint": "other",
                    "issuance_send_id": 1,
                }
            ),
            issuance=issuance,
        )
        self.assertTrue(painted.paints_kits_issuance)
        self.assertFalse(older.paints_kits_issuance)
        self.assertIn("2869-sos", painted.haystack)

    def test_monitor_to_json_serializable(self) -> None:
        with tempfile.TemporaryDirectory(prefix="rd_monitor_json_") as temp:
            root = Path(temp)
            config = _empty_config(root)
            database = CatalogDatabase(config.db_path)
            database.initialize()
            monitor = load_catalog_monitor(database, config)
            payload = monitor_to_json(monitor)
            encoded = json.dumps(payload)
            self.assertIn("kits", payload)
            self.assertTrue(encoded)

    def test_kits_paint_legend_bold_only_on_match_columns(self) -> None:
        sections = kits_paint_legend()
        samples = [item for section in sections for item in section.samples]
        self.assertGreaterEqual(len(sections), 4)
        auto_match = next(
            item
            for item in samples
            if item.column == "Авто МТО" and item.bold
        )
        self.assertEqual(auto_match.fill, REV_MATCH_FILL)
        auto_diff = next(
            item
            for item in samples
            if item.column == "Авто МТО"
            and item.text == "01-AN01"
            and not item.bold
        )
        self.assertEqual(auto_diff.fill, REV_DIFF_FILL)
        auto_ahead = next(
            item
            for item in samples
            if item.column == "Авто МТО" and item.text == "02-AN01"
        )
        self.assertEqual(auto_ahead.fill, AUTO_MTO_AHEAD_FILL)
        self.assertFalse(auto_ahead.bold)
        robot = [item for item in samples if item.column == "Робот МТО · рев."]
        mto = [item for item in samples if item.column == "MTO · рев."]
        self.assertTrue(robot)
        self.assertTrue(mto)
        self.assertTrue(any(item.bold for item in robot))
        self.assertTrue(any(not item.bold for item in robot))
        self.assertFalse(any(item.bold for item in mto))
        compare = [
            item
            for item in samples
            if item.column == AUTO_MTO_COMPARE_STATUS_HEADER
        ]
        self.assertTrue(any(item.bold and item.text.startswith("четкое") for item in compare))
        self.assertTrue(
            any(item.text == "не совпало" and not item.bold for item in compare)
        )
        lag = next(
            item
            for item in samples
            if item.column == "РД · рев."
            and "письма A" in item.meaning
        )
        self.assertEqual(lag.fill, REV_DIFF_FILL)
        self.assertFalse(lag.bold)
        f_od = next(
            item
            for item in samples
            if item.column == "Google · рев. F" and item.text == "01-AN01"
        )
        self.assertEqual(f_od.fill, REV_DIFF_FILL)
        f_mto = next(
            item
            for item in samples
            if item.column == "Google · рев. F" and "MTO 03" in item.text
        )
        self.assertEqual(f_mto.text, "04 · MTO 03")
        self.assertEqual(f_mto.fill, REV_DIFF_FILL)
        self.assertIn("MTO", f_mto.meaning)
        f_absent = next(
            item
            for item in samples
            if item.column == "Google · рев. F" and "MTO Нет" in item.text
        )
        self.assertEqual(f_absent.text, "04 · MTO Нет")
        self.assertEqual(f_absent.fill, REV_MATCH_FILL)
        ok_samples = [item for item in samples if item.column == KITS_OK_HEADER]
        self.assertEqual({item.text for item in ok_samples}, {"да", "нет"})
        self.assertFalse(any(item.bold for item in ok_samples))
        painted = kits_paint_legend({"working": "#123456"})
        working_cell = next(
            item
            for section in painted
            for item in section.samples
            if item.column == "Рабочая рев. РД" and item.text != "—"
        )
        self.assertEqual(working_cell.fill, "#123456")
        self.assertTrue(
            any(
                item.column == "Рабочая рев. РД"
                and item.text.endswith("AB")
                for section in painted
                for item in section.samples
            )
        )
        annulled_fill = color_for(default_palette(), "annulled")
        self.assertTrue(
            any(item.fill == annulled_fill for item in samples)
        )
        self.assertTrue(
            any(
                "аннулирована" in f"{item.text} {item.meaning}".casefold()
                for item in samples
            )
        )
        payload = jsonify_value(sections)
        self.assertEqual(payload[0]["title"], sections[0].title)
        self.assertTrue(payload[0]["samples"])
        encoded = json.dumps(payload, ensure_ascii=False)
        self.assertIn("Авто МТО", encoded)


class KitsProgressStatsTests(unittest.TestCase):
    def test_counts_aligned_and_gaps(self) -> None:
        from types import SimpleNamespace

        rows = (
            SimpleNamespace(
                title="2815",
                kit_ok=True,
                summary_aligned=True,
                code_a=True,
                kit_tdo_passed=True,
                rd_present=True,
                has_mto_problem=False,
            ),
            SimpleNamespace(
                title="2815",
                kit_ok=False,
                summary_aligned=False,
                code_a=False,
                kit_tdo_passed=True,
                rd_present=False,
                has_mto_problem=True,
            ),
            SimpleNamespace(
                title="3000",
                kit_ok=False,
                summary_aligned=True,
                code_a=False,
                kit_tdo_passed=False,
                rd_present=True,
                has_mto_problem=False,
            ),
        )
        stats = kits_progress_stats(rows)
        self.assertEqual(stats.total, 3)
        self.assertEqual(stats.ok, 1)
        self.assertEqual(stats.aligned, 2)
        self.assertEqual(stats.code_a, 1)
        self.assertEqual(stats.tdo, 2)
        self.assertEqual(stats.no_rd, 1)
        self.assertEqual(stats.mto_problems, 1)
        self.assertEqual(stats.titles, 2)
        line = format_kits_progress_stats(stats)
        self.assertEqual(
            line,
            "Ок: 1 / 3 · Совпадает: 2 / 3 · Код A: 1 · ТДО: 2 · "
            "Нет в РД: 1 · Проблемы MTO: 1",
        )
        self.assertEqual(
            format_kits_progress_stats(stats, visible=1),
            line + " · видно: 1",
        )
        self.assertEqual(format_kits_progress_stats(stats, visible=3), line)


class RobotOriginPaintTests(unittest.TestCase):
    _MTIME = 1_700_000_000_000_000_000

    def _row(self, *, rd_rev: str, robot_rev: str, mto_rev: str) -> KitMatrixRow:
        rd_tokens = parse_sheet_revision(rd_rev)
        mto_tokens = parse_sheet_revision(mto_rev)
        robot_tokens = parse_sheet_revision(robot_rev)
        return KitMatrixRow(
            title="6100",
            mark="SOS",
            title_system="6100-SOS",
            rd=SourceKitSnapshot(
                present=True,
                revision=rd_tokens[0],
                appendix=rd_tokens[1],
                revision_text=rd_rev,
                mto_revision=mto_tokens[0],
                mto_appendix=mto_tokens[1],
                mto_revision_text=mto_rev,
                mto_mtime_ns=self._MTIME,
            ),
            robot=SourceKitSnapshot(
                present=True,
                revision=robot_tokens[0],
                appendix=robot_tokens[1],
                revision_text=robot_rev,
                mto_mtime_ns=self._MTIME,
            ),
            sq=SourceKitSnapshot(),
            google=None,
            issuance=None,
            flags=(),
            summary=KitSummary.ALIGNED,
        )

    def test_same_mto_rev_different_od_is_origin_green(self) -> None:
        painted = _paint_kits(
            self._row(rd_rev="04", robot_rev="03", mto_rev="03"),
            None,
        )
        cell = painted.cells["Робот МТО · рев."]
        self.assertEqual(cell.text, "03")
        self.assertEqual(cell.fill, ROBOT_ORIGIN_FILL)
        self.assertFalse(cell.bold)

    def test_same_mto_rev_content_equal_is_origin_and_bold(self) -> None:
        painted = _paint_kits(
            self._row(rd_rev="04", robot_rev="03", mto_rev="03"),
            None,
            mto_content_equal=True,
        )
        cell = painted.cells["Робот МТО · рев."]
        self.assertEqual(cell.fill, ROBOT_ORIGIN_FILL)
        self.assertTrue(cell.bold)

    def test_different_mto_rev_date_close_is_not_origin_green(self) -> None:
        painted = _paint_kits(
            self._row(rd_rev="04", robot_rev="03", mto_rev="04"),
            None,
        )
        cell = painted.cells["Робот МТО · рев."]
        self.assertEqual(cell.fill, REV_DIFF_FILL)
        self.assertNotEqual(cell.fill, ROBOT_ORIGIN_FILL)
        self.assertFalse(cell.bold)
        self.assertIn("не совпадает с «MTO · рев.»", cell.tooltip)

    def test_different_mto_rev_content_equal_is_bold_not_origin(self) -> None:
        painted = _paint_kits(
            self._row(rd_rev="04", robot_rev="03", mto_rev="04"),
            None,
            mto_content_equal=True,
        )
        cell = painted.cells["Робот МТО · рев."]
        self.assertTrue(cell.bold)
        self.assertNotEqual(cell.fill, ROBOT_ORIGIN_FILL)
        self.assertEqual(cell.fill, REV_DIFF_FILL)
        self.assertIn("содержимое", cell.tooltip)

    def test_same_rd_rev_and_date_is_origin_green(self) -> None:
        painted = _paint_kits(
            self._row(rd_rev="03", robot_rev="03", mto_rev="03"),
            None,
        )
        cell = painted.cells["Робот МТО · рев."]
        self.assertEqual(cell.fill, ROBOT_ORIGIN_FILL)
        self.assertFalse(cell.bold)

    def test_same_rd_rev_content_equal_is_origin_and_bold(self) -> None:
        painted = _paint_kits(
            self._row(rd_rev="03", robot_rev="03", mto_rev="03"),
            None,
            mto_content_equal=True,
        )
        cell = painted.cells["Робот МТО · рев."]
        self.assertEqual(cell.fill, ROBOT_ORIGIN_FILL)
        self.assertTrue(cell.bold)


class AutoMtoAheadPaintTests(unittest.TestCase):
    def _paint(
        self,
        *,
        auto_rev: str,
        rd_rev: str,
        mto_rev: str,
        rd_present: bool = True,
    ):
        rd = (
            _rd_snap(rd_rev, mto=mto_rev)
            if rd_present
            else SourceKitSnapshot()
        )
        row = _ok_kit_row(
            title="1600",
            mark="SOS",
            title_system="1600-SOS",
            rd=rd,
            robot=_src_snap(rd_rev) if rd_present else SourceKitSnapshot(),
            google=_google_kit(title="1600", mark="SOS", sheet=rd_rev),
            issuance=_issuance_kit(title="1600", mark="SOS", rev=rd_rev),
        )
        overlay: dict[tuple[str, str], tuple[str, str]] = {}
        if rd_present:
            overlay[kit_identity_key("1600", "SOS")] = (
                r"C:\rd\AGCC.287-1600-SOS.MTO-0001.xlsx",
                mto_rev,
            )
        return _paint_kits(
            row,
            _ok_pipeline(
                title="1600",
                mark="SOS",
                official_revision_text=rd_rev,
                code_revision_text=rd_rev,
            ),
            auto_files=(_auto_file(rev=auto_rev),),
            overlay_mto=overlay,
        )

    def test_ahead_of_official_rd_is_magenta(self) -> None:
        painted = self._paint(auto_rev="02-AN01", rd_rev="02", mto_rev="02")
        cell = painted.cells["Авто МТО"]
        self.assertEqual(cell.text, "02-AN01")
        self.assertEqual(cell.fill, AUTO_MTO_AHEAD_FILL)
        self.assertFalse(cell.bold)
        self.assertIn("выше официальной", cell.tooltip)

    def test_match_stays_green(self) -> None:
        painted = self._paint(auto_rev="02", rd_rev="02", mto_rev="02")
        cell = painted.cells["Авто МТО"]
        self.assertEqual(cell.fill, REV_MATCH_FILL)
        self.assertTrue(cell.bold)
        self.assertNotIn("выше официальной", cell.tooltip)

    def test_behind_stays_yellow(self) -> None:
        painted = self._paint(auto_rev="01", rd_rev="02", mto_rev="02")
        cell = painted.cells["Авто МТО"]
        self.assertEqual(cell.fill, REV_DIFF_FILL)
        self.assertFalse(cell.bold)
        self.assertNotIn("выше официальной", cell.tooltip)

    def test_ahead_wins_over_mto_filename_match(self) -> None:
        painted = self._paint(
            auto_rev="02-AN01", rd_rev="02", mto_rev="02-AN01"
        )
        cell = painted.cells["Авто МТО"]
        self.assertEqual(cell.fill, AUTO_MTO_AHEAD_FILL)
        self.assertTrue(cell.bold)

    def test_missing_rd_files_not_magenta(self) -> None:
        painted = self._paint(
            auto_rev="02-AN01", rd_rev="02", mto_rev="02", rd_present=False
        )
        cell = painted.cells["Авто МТО"]
        self.assertNotEqual(cell.fill, AUTO_MTO_AHEAD_FILL)

    def test_magenta_does_not_block_kit_ok(self) -> None:
        painted = self._paint(auto_rev="02-AN01", rd_rev="02", mto_rev="02")
        self.assertEqual(painted.cells["Авто МТО"].fill, AUTO_MTO_AHEAD_FILL)
        self.assertTrue(painted.kit_ok)


class KitOkFlagTests(unittest.TestCase):
    def test_ok_when_aligned_revs_agreed_and_code_a(self) -> None:
        painted = _paint_kits(
            _ok_kit_row(),
            _ok_pipeline(),
            mto_flags=MtoKitFlags(has_problem=True),
        )
        cell = painted.cells[KITS_OK_HEADER]
        self.assertTrue(painted.kit_ok)
        self.assertTrue(painted.summary_aligned)
        self.assertEqual(cell.text, "да")
        self.assertEqual(cell.sort_key, 0)
        self.assertEqual(cell.fill, REV_MATCH_FILL)
        self.assertFalse(cell.bold)
        self.assertIn("хороший", painted.haystack)

    def test_ok_ignores_google_de_and_empty_sq(self) -> None:
        painted = _paint_kits(_ok_kit_row(), _ok_pipeline())
        self.assertEqual(painted.cells["Google · рев."].fill, REV_DIFF_FILL)
        self.assertEqual(painted.cells["SQ · рев."].text, "—")
        self.assertTrue(painted.kit_ok)

    def test_not_ok_when_robot_rev_differs(self) -> None:
        painted = _paint_kits(
            _ok_kit_row(robot=_src_snap("01-AN02")),
            _ok_pipeline(),
        )
        cell = painted.cells[KITS_OK_HEADER]
        self.assertFalse(painted.kit_ok)
        self.assertEqual(cell.text, "нет")
        self.assertEqual(cell.sort_key, 1)
        self.assertEqual(cell.fill, REV_DIFF_FILL)
        self.assertIn("Робот МТО", cell.tooltip)

    def test_not_ok_when_review_not_agreed(self) -> None:
        painted = _paint_kits(
            _ok_kit_row(),
            _ok_pipeline(status="tdo_review"),
        )
        self.assertFalse(painted.kit_ok)
        self.assertIn("Согласован", painted.cells[KITS_OK_HEADER].tooltip)

    def test_not_ok_when_rd_lags_letter_a(self) -> None:
        painted = _paint_kits(
            _ok_kit_row(),
            _ok_pipeline(code_revision_text="01-AN02"),
        )
        self.assertFalse(painted.kit_ok)
        self.assertIn("письма A", painted.cells[KITS_OK_HEADER].tooltip)

    def test_not_ok_when_summary_not_aligned(self) -> None:
        painted = _paint_kits(
            _ok_kit_row(summary=KitSummary.GAP_ROBOT, robot=SourceKitSnapshot()),
            _ok_pipeline(),
        )
        self.assertFalse(painted.summary_aligned)
        self.assertFalse(painted.kit_ok)
        self.assertIn("Совпадает", painted.cells[KITS_OK_HEADER].tooltip)


class OfficialFolderMtoTests(unittest.TestCase):
    _official = r"\\rd\6100\SOS\Для передачи\05_рев.04"
    _older = r"\\rd\6100\SOS\Для передачи\04_рев.03"
    _copied = (
        r"\\rd\6100\SOS\Для передачи\05_рев.04"
        r"\PDF\AGCC.287-6100-SOS.MTO-0001_03_RU.xlsx"
    )
    _old_file = (
        r"\\rd\6100\SOS\Для передачи\04_рев.03"
        r"\PDF\AGCC.287-6100-SOS.MTO-0001_03_RU.xlsx"
    )

    def _pipeline(self) -> dict[tuple[str, str], KitPipelineRow]:
        return {
            kit_identity_key("6100", "SOS"): KitPipelineRow(
                title="6100",
                mark="SOS",
                status="agreed",
                code="A",
                official_revision_text="04",
                working_revision_text="04-AN01",
            )
        }

    def test_copied_mto_in_official_folder_is_current(self) -> None:
        rows = (
            _worklist_row(
                title="6100",
                mark="SOS",
                revision_text="04",
                is_current=True,
                package_path=self._official,
            ),
            _worklist_row(
                title="6100",
                mark="SOS",
                revision_text="03",
                package_path=self._official,
                mto_path=self._copied,
                has_mto=True,
            ),
        )
        overlay = rd_mto_overlay_by_kit(rows, self._pipeline())
        path, rev = overlay[kit_identity_key("6100", "SOS")]
        self.assertEqual(path, self._copied)
        self.assertEqual(rev, "03")
        target = auto_mto_rd_target(
            "6100",
            "SOS",
            pins={},
            overlay_by_kit=overlay,
        )
        self.assertEqual(target[0], self._copied)
        self.assertEqual(target[1], "03")
        self.assertFalse(target[2])

    def test_older_nn_mto_is_not_used(self) -> None:
        rows = (
            _worklist_row(
                title="6100",
                mark="SOS",
                revision_text="04",
                is_current=True,
                package_path=self._official,
            ),
            _worklist_row(
                title="6100",
                mark="SOS",
                revision_text="03",
                package_path=self._older,
                mto_path=self._old_file,
                has_mto=True,
            ),
        )
        overlay = rd_mto_overlay_by_kit(rows, self._pipeline())
        self.assertNotIn(kit_identity_key("6100", "SOS"), overlay)

    def test_same_rev_mto_in_official_folder(self) -> None:
        mto = (
            self._official
            + r"\PDF\AGCC.287-6100-SOS.MTO-0001_04_RU.xlsx"
        )
        rows = (
            _worklist_row(
                title="6100",
                mark="SOS",
                revision_text="04",
                is_current=True,
                package_path=self._official,
                mto_path=mto,
                has_mto=True,
            ),
        )
        overlay = rd_mto_overlay_by_kit(rows, self._pipeline())
        self.assertEqual(overlay[kit_identity_key("6100", "SOS")][1], "04")

    def test_kits_column_shows_net_when_folder_has_no_mto(self) -> None:
        painted = _paint_kits(
            _ok_kit_row(
                title="6100",
                mark="SOS",
                title_system="6100-SOS",
                rd=_rd_snap("04", mto="03"),
            ),
            _ok_pipeline(
                title="6100",
                mark="SOS",
                official_revision_text="04",
                working_revision_text="04-AN01",
            ),
            overlay_mto={},
        )
        cell = painted.cells["MTO · рев."]
        self.assertEqual(cell.text, OFFICIAL_FOLDER_MTO_MISSING)
        self.assertEqual(cell.palette_key, "no_mto")
        self.assertIn("нет файла mto", cell.tooltip.casefold())

    def test_kits_column_shows_copied_filename_rev(self) -> None:
        key = kit_identity_key("6100", "SOS")
        painted = _paint_kits(
            _ok_kit_row(
                title="6100",
                mark="SOS",
                title_system="6100-SOS",
                rd=_rd_snap("04", mto="03"),
            ),
            _ok_pipeline(
                title="6100",
                mark="SOS",
                official_revision_text="04",
                working_revision_text="04-AN01",
            ),
            overlay_mto={key: (self._copied, "03")},
            auto_files=(_auto_file(title="6100", mark="SOS", rev="03"),),
        )
        self.assertEqual(painted.cells["MTO · рев."].text, "03")
        compare = painted.cells[AUTO_MTO_COMPARE_STATUS_HEADER]
        self.assertEqual(compare.text, "рев. совпала")

    def test_folder_mto_text_helpers(self) -> None:
        self.assertEqual(
            kits_official_folder_mto_text(rd_present=True, revision_text="03"),
            "03",
        )
        self.assertEqual(
            kits_official_folder_mto_text(rd_present=True, revision_text=""),
            OFFICIAL_FOLDER_MTO_MISSING,
        )
        self.assertEqual(
            kits_official_folder_mto_text(rd_present=False, revision_text=""),
            "—",
        )


class PipelineStatusTooltipTests(unittest.TestCase):
    def test_review_tooltip_names_f_and_issuance_sheets(self) -> None:
        google = replace(
            _google_kit(
                title="6816",
                mark="KSB",
                sheet="01",
                lines=(
                    "09.10.2023 прошла вх контр AGCC.287-BCC-NPG-TRM-003144",
                    "10.01.2024 код A рев. 01",
                ),
            ),
            status_sheet="Согласовано",
        )
        issuance = replace(
            _issuance_kit(title="6816", mark="KSB", rev="0-AN01"),
            status="На рассмотрении вх.контроля",
            send_date="01.07.2025",
            send_date_sortable="2025-07-01",
            send_transmittal="PGS-AGCC-TRM-011199",
        )
        pipeline = KitPipelineRow(
            title="6816",
            mark="KSB",
            status="tdo_review",
            official_revision_text="0-AN01",
            working_revision_text="01",
            tdo_date="09.10.2023",
            code="A",
            code_stale=True,
            code_revision_text="01",
            code_date="10.01.2024",
        )
        row = KitMatrixRow(
            title="6816",
            mark="KSB",
            title_system="6816-KSB",
            rd=SourceKitSnapshot(present=True, revision_text="0-AN01"),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=google,
            issuance=issuance,
            flags=(),
            summary=KitSummary.ALIGNED,
        )
        painted = _paint_kits(row, pipeline)
        review = painted.cells["Статус рассмотрения"]
        self.assertIn("Прошел ТДО (09.10.2023)", review.text)
        self.assertIn("РД 0-AN01", review.text)
        self.assertIn("отпр. 01.07.2025", review.text)
        self.assertIn(KITS_F_SHEET_LABEL, review.tooltip)
        self.assertIn(KITS_DE_SHEET_LABEL, review.tooltip)
        self.assertIn(ISSUANCE_SHEET_LABEL, review.tooltip)
        self.assertIn("Как читать подпись", review.tooltip)
        self.assertIn("Google-таблицам", review.tooltip)
        self.assertIn("РД на диске", review.tooltip)
        self.assertIn("Внимание: расхождение с D/E", review.tooltip)
        self.assertIn("AGCC.287-BCC-NPG-TRM-003144", review.tooltip)
        self.assertIn("PGS-AGCC-TRM-011199", review.tooltip)
        self.assertIn("не входит в этот статус", review.tooltip)
        self.assertIn("Прохождение 09.10.2023", review.tooltip)
        self.assertIn("Отправка 01.07.2025", review.tooltip)
        self.assertIn(KITS_F_SHEET_LABEL.casefold(), painted.haystack)
        self.assertIn(KITS_DE_SHEET_LABEL.casefold(), painted.haystack)
        self.assertIn(ISSUANCE_SHEET_LABEL.casefold(), painted.haystack)
        self.assertIn("=== Статус рассмотрения ===", painted.tooltips_text)
        approval = painted.cells["Статус согласования"]
        self.assertIn("новее диска", approval.text)
        self.assertIn(KITS_F_SHEET_LABEL, approval.tooltip)
        self.assertIn(KITS_DE_SHEET_LABEL, approval.tooltip)
        self.assertIn("не задаёт", approval.tooltip)
        self.assertIn("отношение:", approval.tooltip)
        self.assertIn("новее диска", approval.tooltip)
        self.assertNotIn("устар", approval.tooltip)
        self.assertNotIn("из kit_pipeline:", approval.tooltip)
        direct = pipeline_review_tooltip(
            pipeline, google=google, issuance=issuance
        )
        self.assertEqual(review.tooltip, direct)
        self.assertEqual(
            approval.tooltip,
            pipeline_approval_tooltip(
                pipeline, google=google, issuance=issuance
            ),
        )

    def test_v2_current_b_lives_on_review_column(self) -> None:
        google = _google_kit(
            title="7700",
            mark="POS",
            sheet="01",
            lines=(
                "10.06.2026 прошла ТДО AGCC-BCC-TRM-000200",
                "12.06.2026 код B рев. 01 AGCC-BCC-TRM-000200",
            ),
        )
        issuance = replace(
            _issuance_kit(title="7700", mark="POS", rev="01"),
            send_date="01.06.2026",
            send_transmittal="AGCC-BCC-TRM-000200",
        )
        pipeline = KitPipelineRow(
            title="7700",
            mark="POS",
            status="tdo_review",
            official_revision_text="01",
            tdo_date="10.06.2026",
            code="B",
            code_stale=False,
            code_revision_text="01",
            code_date="12.06.2026",
        )
        row = KitMatrixRow(
            title="7700",
            mark="POS",
            title_system="7700-POS",
            rd=SourceKitSnapshot(present=True, revision_text="01"),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=google,
            issuance=issuance,
            flags=(),
            summary=KitSummary.ALIGNED,
        )
        painted = _paint_kits(row, pipeline)
        review = painted.cells["Статус рассмотрения"]
        approval = painted.cells["Статус согласования"]
        self.assertIn(" · B · ", review.text)
        self.assertEqual(approval.text, "—")
        self.assertFalse(approval.fill)
        self.assertIn("этого цикла", review.tooltip)
        self.assertIn("Статус рассмотрения", approval.tooltip)

    def test_v3_google_face_ahead_puts_b_on_review(self) -> None:
        google = _google_kit(
            title="6550",
            mark="SKUD",
            sheet="0-AN02",
            lines=(
                "01.06.2026 отправлена на ТДО рев. 0-AN02 AGCC-BCC-TRM-000300",
                "10.06.2026 прошла ТДО рев. 0-AN02 AGCC-BCC-TRM-000300",
                "12.06.2026 код B рев. 0-AN02 AGCC-BCC-TRM-000300",
            ),
        )
        issuance = replace(
            _issuance_kit(title="6550", mark="SKUD", rev="0-AN02"),
            send_date="01.06.2026",
            send_transmittal="AGCC-BCC-TRM-000300",
            incoming_control_date="10.06.2026",
        )
        pipeline = KitPipelineRow(
            title="6550",
            mark="SKUD",
            status="tdo_review",
            official_revision_text="0-AN01",
            tdo_date="09.10.2023",
            code="B",
            code_stale=True,
            code_revision_text="0-AN02",
            code_date="12.06.2026",
        )
        row = KitMatrixRow(
            title="6550",
            mark="SKUD",
            title_system="6550-SKUD",
            rd=SourceKitSnapshot(present=True, revision_text="0-AN01"),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=google,
            issuance=issuance,
            flags=(),
            summary=KitSummary.ALIGNED,
        )
        painted = _paint_kits(row, pipeline)
        review = painted.cells["Статус рассмотрения"]
        approval = painted.cells["Статус согласования"]
        self.assertIn(" · B · ", review.text)
        self.assertIn("выдача 0-AN02", review.text)
        self.assertNotIn("0-AN01", review.text)
        self.assertNotIn("09.10.2023", review.text)
        self.assertIn("10.06.2026", review.text)
        self.assertEqual(approval.text, "—")
        self.assertFalse(approval.fill)
        self.assertIn("Google-таблицам", review.tooltip)
        self.assertTrue(painted.kit_tdo_passed)
        self.assertEqual(review.palette_key, "tdo_review")

    def test_tooltip_section_start_and_ctrl_hint(self) -> None:
        cells = {
            "Сводка": MonitorCell(text="Совпадает", tooltip="сводка подробно"),
            "РД · рев.": MonitorCell(text="01", tooltip="официальная рев."),
            "Статус рассмотрения": MonitorCell(
                text="Прошел ТДО",
                tooltip="лист F",
            ),
        }
        body = format_kits_row_tooltips("6816", "KSB", cells)
        review_at = body.find("=== Статус рассмотрения ===")
        self.assertGreater(review_at, 0)
        self.assertEqual(kits_tooltip_section_start(body, review_at + 8), review_at)
        between = body.find("=== РД · рев. ===")
        self.assertEqual(
            kits_tooltip_section_start(body, review_at - 1),
            between,
        )
        self.assertEqual(kits_tooltip_section_start(body, 0), 0)
        pane = format_kits_tips_pane(body, with_ctrl_click_hint=True)
        self.assertTrue(pane.startswith(KITS_TIPS_CTRL_CLICK_HINT))
        hinted_review = pane.find("=== Статус рассмотрения ===")
        self.assertEqual(
            kits_tooltip_section_start(pane, hinted_review + 3),
            hinted_review,
        )
        self.assertEqual(
            kits_tooltip_header_offset(pane, "Статус рассмотрения"),
            hinted_review,
        )
        self.assertEqual(
            kits_tooltip_header_offset(body, "РД · рев."),
            between,
        )
        self.assertIsNone(kits_tooltip_header_offset(body, "Марка"))
        self.assertIsNone(kits_tooltip_header_offset("", "Сводка"))
        self.assertEqual(format_kits_tips_pane(""), "")

    def test_wrap_kits_tips_text_prose_not_tables(self) -> None:
        runon = " ".join(["слово"] * 40)
        table = "NN\tроль\tпуть\\\\bcc\\eng\\very\\long\\folder\\file.xlsx"
        unc = "\\\\bcc\\eng\\" + ("dir\\" * 20) + "file.xlsx"
        body = f"=== РД · рев. ===\n{runon}\n{table}\n{unc}\n"
        wrapped = wrap_kits_tips_text(body)
        self.assertIn("=== РД · рев. ===", wrapped)
        self.assertIn(runon, wrapped)
        self.assertIn(table, wrapped)
        self.assertIn(unc, wrapped)

    def test_wrap_kits_tips_at_sentence_ends(self) -> None:
        wrapped = wrap_kits_tips_text(KITS_WORKING_REV_TOOLTIP)
        self.assertNotIn("помечена\n", wrapped)
        self.assertNotIn("\nвручную", wrapped)
        self.assertNotIn("спецификаций и\n", wrapped)
        self.assertNotIn("и\nсверке", wrapped)
        lines = [line.strip() for line in wrapped.splitlines() if line.strip()]
        self.assertEqual(
            lines,
            [
                (
                    "Ревизия строго выше официальной «РД · рев.» "
                    "(на диске выше последней выдачи или папка помечена вручную)."
                ),
                "Суффикс AB — рабочая папка as-build.",
                "Не участвует в комплектах, выгрузке спецификаций и сверке MTO.",
            ],
        )
        mixed = (
            "Эта колонка — рев. РД актуального пакета, не рабочая папка. "
            "Вторая фраза начинается с заглавной и остаётся целиком на строке."
        )
        mixed_wrapped = wrap_kits_tips_text(mixed)
        self.assertIn("рев. РД актуального", mixed_wrapped)
        self.assertNotIn("рев.\n", mixed_wrapped)
        mixed_lines = [
            line.strip()
            for line in mixed_wrapped.splitlines()
            if line.strip()
        ]
        self.assertEqual(len(mixed_lines), 2)
        self.assertTrue(mixed_lines[0].endswith("папка."))
        self.assertTrue(mixed_lines[1].startswith("Вторая фраза"))

    def test_wrap_kits_tips_does_not_break_transfer_review_mid_phrase(self) -> None:
        from rd_catalog.kits import KitSummary, _SUMMARY_HINTS

        hint = _SUMMARY_HINTS[KitSummary.TRANSFER_REVIEW]
        wrapped = wrap_kits_tips_text(hint)
        self.assertNotIn("или\n", wrapped)
        self.assertNotIn("\nпо дате", wrapped)
        self.assertNotIn("редактируемые\n", wrapped)
        self.assertNotIn("согл.\n", wrapped)
        self.assertNotIn("\nпередаче", wrapped)
        self.assertNotIn("состав\n", wrapped)
        pane = format_kits_row_tooltips(
            "7416",
            "SKUD.1",
            {
                "Сводка": MonitorCell(
                    text="Проверить передачи",
                    tooltip=hint,
                )
            },
        )
        self.assertIn("=== Сводка ===", pane)
        self.assertNotIn("или\nпо дате", pane)
        self.assertTrue(
            any(
                line.startswith("Текущий состав")
                for line in pane.splitlines()
            )
        )

    def test_wrap_kits_tips_does_not_split_unc_with_spaces(self) -> None:
        identity = (
            "  0-AN01  AGCC.287-7416-SKUD.1.MTO-0001_0-AN01_RU.xlsx  "
            "2025-07-23"
        )
        send_path = (
            r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\АН_RFQ\_АН\7590\На "
            r"отправку\SKUD\01_рев_AN01_AGCC.287‐7590‐SKUD\DWG"
        )
        sq_path = (
            r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\АН_RFQ\_АН\Ответы на "
            r"SQ-запросы\SQ-UIO-AGCC-SOT-20763-0\DWG\7416\SKUD"
        )
        body = (
            "Файлы АН:\n"
            f"{identity}  {send_path}\n"
            f"{identity}  {sq_path}\n"
        )
        wrapped = wrap_kits_tips_text(body)
        self.assertIn(send_path, wrapped)
        self.assertIn(sq_path, wrapped)
        self.assertNotIn("На\n", wrapped)
        self.assertNotIn("на\n", wrapped)
        path_lines = [
            line.strip()
            for line in wrapped.splitlines()
            if tooltip_has_fs_path(line)
        ]
        self.assertEqual(path_lines, [send_path, sq_path])
        for line in wrapped.splitlines():
            if tooltip_has_fs_path(line):
                self.assertNotIn("0-AN01", line)


class GoogleFMtoPaintTests(unittest.TestCase):
    """Комплекты F-cell MTO suffix vs official-folder MTO."""

    def _paint(
        self,
        *,
        line: str,
        rd_rev: str,
        disk_mto: str | None,
        snapshot_mto: str | None = None,
        pipeline: KitPipelineRow | None = None,
    ):
        if snapshot_mto is not None:
            snap_mto = snapshot_mto
        elif disk_mto is not None:
            snap_mto = disk_mto
        else:
            snap_mto = rd_rev
        google = _google_kit(
            title="2210",
            mark="KSB",
            sheet=rd_rev,
            lines=(line,),
        )
        row = KitMatrixRow(
            title="2210",
            mark="KSB",
            title_system="2210-KSB",
            rd=_rd_snap(rd_rev, mto=snap_mto),
            robot=SourceKitSnapshot(),
            sq=SourceKitSnapshot(),
            google=google,
            issuance=None,
            flags=(),
            summary=KitSummary.ALIGNED,
        )
        overlay: dict[tuple[str, str], tuple[str, str]] = {}
        if disk_mto is not None:
            overlay[kit_identity_key("2210", "KSB")] = (
                r"C:\rd\AGCC.287-2210-KSB.MTO-0001.xlsx",
                disk_mto,
            )
        return _paint_kits(row, pipeline, overlay_mto=overlay)

    def test_f_mto_match_keeps_od_fill(self) -> None:
        painted = self._paint(
            line="09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 MTO 03",
            rd_rev="04",
            disk_mto="03",
        )
        f_cell = painted.cells["Google · рев. F"]
        self.assertIn("MTO 03", f_cell.text)
        self.assertEqual(f_cell.text, "04 · MTO 03")
        self.assertEqual(f_cell.fill, REV_MATCH_FILL)
        self.assertIn("F · рев. 04 совпала с РД · рев.", f_cell.tooltip)
        self.assertIn("F MTO 03 совпала с MTO · рев.", f_cell.tooltip)
        self.assertIn("mto 03", painted.haystack)
        mismatched_od = self._paint(
            line="09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 MTO 03",
            rd_rev="01-AN01",
            disk_mto="03",
        )
        self.assertEqual(
            mismatched_od.cells["Google · рев. F"].fill, REV_DIFF_FILL
        )
        self.assertIn("MTO 03", mismatched_od.cells["Google · рев. F"].text)

    def test_f_mto_mismatch_yellow_wins(self) -> None:
        painted = self._paint(
            line="09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 MTO 03",
            rd_rev="04",
            disk_mto="04",
        )
        f_cell = painted.cells["Google · рев. F"]
        self.assertEqual(f_cell.text, "04 · MTO 03")
        self.assertEqual(f_cell.fill, REV_DIFF_FILL)
        self.assertIn("расходится с MTO · рев. 04", f_cell.tooltip)
        mto_cell = painted.cells["MTO · рев."]
        self.assertEqual(mto_cell.text, "04")
        self.assertEqual(mto_cell.fill, REV_DIFF_FILL)
        self.assertIn("MTO на диске 04", mto_cell.tooltip)
        self.assertIn("в F указано MTO 03", mto_cell.tooltip)
        ok_row = _ok_kit_row(
            google=_google_kit(
                title="2210",
                mark="KSB",
                sheet="01",
                lines=("04.12.2025 код А на рев. 01-AN01 MTO 03",),
            ),
        )
        key = kit_identity_key("2210", "KSB")
        ok_painted = _paint_kits(
            ok_row,
            _ok_pipeline(),
            overlay_mto={key: (r"C:\rd\mto.xlsx", "04")},
        )
        self.assertEqual(ok_painted.cells["Google · рев. F"].fill, REV_DIFF_FILL)
        self.assertTrue(ok_painted.kit_ok)

    def test_f_without_suffix_is_event_rev_only(self) -> None:
        painted = self._paint(
            line="09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999",
            rd_rev="04",
            disk_mto="03",
        )
        parts = last_event_parts(painted.matrix_row.google)
        self.assertEqual(len(parts), 4)
        event_rev = parts[2]
        f_cell = painted.cells["Google · рев. F"]
        self.assertEqual(event_rev, "04")
        self.assertEqual(f_cell.text, event_rev)
        self.assertEqual(
            format_kits_google_f_rev(painted.matrix_row.google.last_event),
            event_rev,
        )
        self.assertEqual(f_cell.fill, REV_MATCH_FILL)
        self.assertNotIn("MTO", f_cell.text)
        self.assertFalse(f_cell.tooltip)

    def test_auto_in_text_does_not_yellow(self) -> None:
        painted = self._paint(
            line="09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 auto",
            rd_rev="04",
            disk_mto="03",
        )
        f_cell = painted.cells["Google · рев. F"]
        self.assertEqual(f_cell.text, "04")
        self.assertNotIn("auto", f_cell.text)
        self.assertEqual(f_cell.fill, REV_MATCH_FILL)
        self.assertFalse(f_cell.tooltip)
        with_mto = self._paint(
            line=(
                "09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 MTO 03 auto"
            ),
            rd_rev="04",
            disk_mto="03",
        )
        auto_cell = with_mto.cells["Google · рев. F"]
        self.assertEqual(auto_cell.text, "04 · MTO 03")
        self.assertNotIn("auto", auto_cell.text)
        self.assertEqual(auto_cell.fill, REV_MATCH_FILL)
        self.assertNotIn("автоматом", auto_cell.tooltip)

    def test_f_mto_when_disk_missing_is_yellow(self) -> None:
        painted = self._paint(
            line="09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 MTO 03",
            rd_rev="04",
            disk_mto=None,
        )
        f_cell = painted.cells["Google · рев. F"]
        self.assertEqual(f_cell.fill, REV_DIFF_FILL)
        self.assertIn("нет файла MTO", f_cell.tooltip)
        mto_cell = painted.cells["MTO · рев."]
        self.assertEqual(mto_cell.text, OFFICIAL_FOLDER_MTO_MISSING)
        self.assertEqual(mto_cell.palette_key, "no_mto")
        self.assertIn("F указано MTO 03", mto_cell.tooltip)

    def test_f_mto_absent_matches_disk_missing(self) -> None:
        painted = self._paint(
            line=(
                "09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 MTO Нет auto"
            ),
            rd_rev="04",
            disk_mto=None,
        )
        f_cell = painted.cells["Google · рев. F"]
        self.assertEqual(f_cell.text, "04 · MTO Нет")
        self.assertNotIn("auto", f_cell.text)
        self.assertEqual(f_cell.fill, REV_MATCH_FILL)
        self.assertIn("MTO Нет", f_cell.tooltip)
        self.assertNotIn("автоматом", f_cell.tooltip)
        self.assertIn("mto нет", painted.haystack)
        mto_cell = painted.cells["MTO · рев."]
        self.assertEqual(mto_cell.text, OFFICIAL_FOLDER_MTO_MISSING)
        self.assertEqual(mto_cell.palette_key, "no_mto")
        self.assertNotEqual(mto_cell.fill, REV_DIFF_FILL)

    def test_f_mto_absent_yellow_when_disk_has_file(self) -> None:
        painted = self._paint(
            line=(
                "09.09.2026 код А на рев. 04 AGCC-BCC-TRM-000999 MTO Нет auto"
            ),
            rd_rev="04",
            disk_mto="03",
        )
        f_cell = painted.cells["Google · рев. F"]
        self.assertEqual(f_cell.text, "04 · MTO Нет")
        self.assertEqual(f_cell.fill, REV_DIFF_FILL)
        self.assertIn("MTO Нет", f_cell.tooltip)
        mto_cell = painted.cells["MTO · рев."]
        self.assertEqual(mto_cell.text, "03")
        self.assertEqual(mto_cell.fill, REV_DIFF_FILL)
        self.assertIn("MTO Нет", mto_cell.tooltip)


if __name__ == "__main__":
    unittest.main()
