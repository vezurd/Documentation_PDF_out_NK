"""Offline smoke for the canonical DS registry (no UNC writes)."""

from __future__ import annotations

import sys
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from RFQ.rfp_parts.ds_registry import (
    HEADER_CONDITIONAL_MARK,
    HEADER_REQUIRED_MARK,
    HDR_SOURCE_ID,
    HDR_STATUS,
    ISSUE_DUPLICATE_SOURCE_ID,
    ISSUE_GAPPED_BLOCK,
    ISSUE_GROUP_MISMATCH,
    ISSUE_MULTI_LINK_NEEDS_SPLIT,
    ISSUE_PARTIAL_BLOCK,
    ISSUE_UL_IDENTITY_MISMATCH,
    ISSUE_UL_UNPARSED,
    MIGRATION_REPORT_PREFIX,
    MODE_NEEDS_SPLIT,
    MODE_NO_UL,
    MODE_WHOLE,
    REGISTRY_SHEET_NAME,
    REGISTRY_TABLE_NAME,
    STATUS_ACTIVE,
    STATUS_HISTORY,
    DsRegistryRelation,
    DsRegistryReplaceError,
    DsRegistryRow,
    backup_and_replace_registry,
    canonical_supply_group_id,
    detect_registry_format,
    load_registry,
    migrate_registry,
    no_ul_group_id,
    validate_registry_rows,
    write_registry_workbook,
)


def _load_xlsx(path: Path):
    return load_workbook(BytesIO(path.read_bytes()))


def _fill_rgb(cell) -> str:
    color = getattr(cell.fill, "fgColor", None)
    rgb = getattr(color, "rgb", None) if color is not None else None
    return str(rgb or "").upper()


def _write_legacy(path: Path, rows: list[tuple[object, object, object, object]]) -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws["A1"] = "Актуальный ДС"
    ws["B1"] = "Старый ДС"
    ws["C1"] = "УЛ"
    ws["D1"] = "Примечание"
    for excel_row, (current, old, ul_text, note) in enumerate(rows, start=2):
        ws.cell(row=excel_row, column=1, value=current)
        ws.cell(row=excel_row, column=2, value=old)
        ws.cell(row=excel_row, column=3, value=ul_text)
        ws.cell(row=excel_row, column=4, value=note)
    wb.save(path)
    wb.close()


def _legacy_53_style_rows() -> list[tuple[object, object, object, object]]:
    return [
        (13, "старый 13", "согл УЛ ДС13", ""),
        (47, "", "согл УЛ ДС13", ""),
        (8, "", "согл УЛ ДС8;согл УЛ ДС81", "два УЛ"),
        (101, "", "", "без упаковки"),
        ("4905_1", "4905", "согл УЛ 4905", ""),
        (4905, "", "согл УЛ 4905", ""),
        ("", "99", "", "архивная запись без актуального ДС"),
    ]


class DsRegistrySmokeTest(unittest.TestCase):
    def test_legacy_migration_53_style_patterns(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            root = Path(raw)
            legacy_path = root / "Реестр_ДС_УЛ.xlsx"
            output_path = root / "Реестр_ДС_новый.xlsx"
            report_path = root / f"{MIGRATION_REPORT_PREFIX}_smoke.xlsx"
            _write_legacy(legacy_path, _legacy_53_style_rows())
            self.assertEqual(detect_registry_format(legacy_path), "legacy")

            result = migrate_registry(
                legacy_path, output_path, report_path=report_path
            )
            self.assertTrue(output_path.is_file())
            self.assertTrue(report_path.is_file())
            self.assertEqual(result.source_path, legacy_path)
            self.assertEqual(result.output_path, output_path)
            self.assertEqual(result.report_path, report_path)
            self.assertEqual(result.active_count, 6)
            self.assertEqual(result.history_count, 1)
            self.assertTrue(result.validation.is_ok)

            by_id = {row.source_id: row for row in result.rows if row.source_id}
            self.assertEqual(by_id["13"].status, STATUS_ACTIVE)
            self.assertEqual(by_id["13"].previous_ds, "старый 13")
            self.assertEqual(by_id["13"].revision, "")
            self.assertEqual(by_id["13"].relations[0].group_id, "ДС13")
            self.assertEqual(by_id["13"].relations[0].rfp_key, "13")
            self.assertEqual(by_id["13"].relations[0].ul_folder, "согл УЛ ДС13")
            self.assertEqual(by_id["13"].relations[0].mode, MODE_WHOLE)

            self.assertEqual(by_id["47"].relations[0].group_id, "ДС13")
            self.assertEqual(by_id["47"].relations[0].rfp_key, "13")
            self.assertEqual(by_id["47"].relations[0].ul_folder, "согл УЛ ДС13")

            row8 = by_id["8"]
            self.assertEqual(len(row8.relations), 2)
            self.assertEqual(row8.relations[0].group_id, "ДС8")
            self.assertEqual(row8.relations[0].rfp_key, "8")
            self.assertEqual(row8.relations[0].ul_folder, "согл УЛ ДС8")
            self.assertEqual(row8.relations[1].group_id, "ДС81")
            self.assertEqual(row8.relations[1].rfp_key, "81")
            self.assertEqual(row8.relations[1].ul_folder, "согл УЛ ДС81")
            self.assertEqual(row8.relations[0].mode, MODE_NEEDS_SPLIT)
            self.assertEqual(row8.relations[1].mode, MODE_NEEDS_SPLIT)
            self.assertTrue(
                any(
                    item.code == ISSUE_MULTI_LINK_NEEDS_SPLIT and item.blocks_overlay
                    for item in result.validation.issues
                )
            )
            self.assertTrue(result.validation.blocks_overlay)

            row101 = by_id["101"]
            self.assertEqual(row101.relations[0].group_id, no_ul_group_id("101"))
            self.assertEqual(row101.relations[0].rfp_key, "101")
            self.assertEqual(row101.relations[0].ul_folder, "")
            self.assertEqual(row101.relations[0].mode, MODE_NO_UL)

            self.assertIn("4905_1", by_id)
            self.assertEqual(by_id["4905_1"].source_id, "4905_1")
            self.assertEqual(by_id["4905_1"].relations[0].group_id, "ДС4905")
            self.assertEqual(by_id["4905"].relations[0].group_id, "ДС4905")
            self.assertEqual(by_id["4905"].relations[0].rfp_key, "4905")
            self.assertEqual(by_id["4905"].relations[0].ul_folder, "согл УЛ 4905")

            history = [row for row in result.rows if row.is_history]
            self.assertEqual(len(history), 1)
            self.assertEqual(history[0].status, STATUS_HISTORY)
            self.assertEqual(history[0].source_id, "")
            self.assertEqual(history[0].previous_ds, "99")
            self.assertEqual(history[0].note, "архивная запись без актуального ДС")
            self.assertEqual(history[0].relations, ())
            self.assertEqual(history[0].original_excel_row, 8)
            self.assertEqual(history[0].revision, "")

            report_wb = _load_xlsx(report_path)
            try:
                self.assertEqual(
                    set(report_wb.sheetnames),
                    {"Сводка", "Соответствие строк", "Проблемы"},
                )
                mapping = report_wb["Соответствие строк"]
                mapped_ids = [
                    str(row[3] or "")
                    for row in mapping.iter_rows(min_row=2, values_only=True)
                ]
                self.assertIn("13", mapped_ids)
                self.assertIn("4905_1", mapped_ids)
                self.assertIn("", mapped_ids)
            finally:
                report_wb.close()

    def test_writer_headers_table_validations_colors_comments(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            path = Path(raw) / "registry.xlsx"
            write_registry_workbook(
                path,
                [
                    DsRegistryRow(
                        status=STATUS_ACTIVE,
                        source_id="23",
                        relations=(
                            DsRegistryRelation(
                                group_id="ДС23",
                                rfp_key="23",
                                ul_folder="согл УЛ ДС23",
                                mode=MODE_WHOLE,
                                block_index=1,
                            ),
                        ),
                        original_excel_row=2,
                    )
                ],
            )
            wb = _load_xlsx(path)
            try:
                self.assertEqual(wb.sheetnames, [REGISTRY_SHEET_NAME])
                ws = wb[REGISTRY_SHEET_NAME]
                self.assertEqual(ws.freeze_panes, "A2")
                self.assertIn(REGISTRY_TABLE_NAME, ws.tables)
                table = ws.tables[REGISTRY_TABLE_NAME]
                self.assertTrue(str(table.ref).startswith("A1:"))
                self.assertTrue(ws["A1"].alignment.wrap_text)
                self.assertEqual(ws["A1"].value, HDR_STATUS)
                self.assertEqual(ws["B1"].value, HDR_SOURCE_ID)
                self.assertIn(HEADER_REQUIRED_MARK, str(ws["A1"].value))
                self.assertIn(HEADER_CONDITIONAL_MARK, str(ws["I1"].value))
                self.assertTrue(_fill_rgb(ws["A1"]).endswith("1B4F72"))
                self.assertTrue(_fill_rgb(ws["B1"]).endswith("1B4F72"))
                self.assertTrue(_fill_rgb(ws["I1"]).endswith("FFC000"))
                self.assertTrue(_fill_rgb(ws["C1"]).endswith("D9D9D9"))
                self.assertIsNotNone(ws["A1"].comment)
                self.assertIn("Активен", ws["A1"].comment.text)
                self.assertIn("*", ws["A1"].comment.text)
                self.assertIsNotNone(ws["I1"].comment)
                self.assertIn(MODE_NO_UL, ws["I1"].comment.text)
                self.assertEqual(ws["G1"].border.left.style, "medium")
                self.assertEqual(ws["L1"].border.right.style, "medium")
                formulas = [
                    str(dv.formula1) for dv in ws.data_validations.dataValidation
                ]
                joined = " ".join(formulas)
                self.assertIn(STATUS_ACTIVE, joined)
                self.assertIn(STATUS_HISTORY, joined)
                self.assertIn(MODE_WHOLE, joined)
                self.assertIn(MODE_NEEDS_SPLIT, joined)
                self.assertTrue(ws.conditional_formatting._cf_rules)
            finally:
                wb.close()

    def test_parser_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            path = Path(raw) / "roundtrip.xlsx"
            rows = [
                DsRegistryRow(
                    status=STATUS_ACTIVE,
                    source_id="13",
                    previous_ds="old",
                    note="keep",
                    relations=(
                        DsRegistryRelation(
                            group_id="ДС13",
                            rfp_key="13",
                            ul_folder="согл УЛ ДС13",
                            mode=MODE_WHOLE,
                            block_index=1,
                        ),
                    ),
                    original_excel_row=10,
                ),
                DsRegistryRow(
                    status=STATUS_HISTORY,
                    source_id="",
                    previous_ds="88",
                    note="архив",
                    relations=(),
                    original_excel_row=11,
                ),
                DsRegistryRow(
                    status=STATUS_ACTIVE,
                    source_id="4905_1",
                    relations=(
                        DsRegistryRelation(
                            group_id="ДС4905",
                            rfp_key="4905",
                            ul_folder="согл УЛ 4905",
                            mode=MODE_WHOLE,
                            block_index=1,
                        ),
                    ),
                    original_excel_row=12,
                ),
            ]
            write_registry_workbook(path, rows)
            loaded = load_registry(path)
            self.assertEqual(detect_registry_format(path), "new")
            self.assertTrue(loaded.validation.is_ok)
            self.assertEqual(len(loaded.rows), 3)
            self.assertEqual(loaded.rows[0].source_id, "13")
            self.assertEqual(loaded.rows[0].previous_ds, "old")
            self.assertEqual(loaded.rows[0].note, "keep")
            self.assertEqual(loaded.rows[0].original_excel_row, 10)
            self.assertEqual(loaded.rows[0].relations[0].group_id, "ДС13")
            self.assertEqual(loaded.rows[0].relations[0].rfp_key, "13")
            self.assertEqual(loaded.rows[1].status, STATUS_HISTORY)
            self.assertEqual(loaded.rows[1].relations, ())
            self.assertEqual(loaded.rows[2].source_id, "4905_1")
            self.assertEqual(loaded.rows[2].relations[0].group_id, "ДС4905")
            self.assertEqual(loaded.rows[2].relations[0].rfp_key, "4905")

    def test_validate_duplicate_partial_gap_mismatch(self) -> None:
        duplicate = [
            DsRegistryRow(
                status=STATUS_ACTIVE,
                source_id="13",
                excel_row=2,
                relations=(
                    DsRegistryRelation(
                        group_id="ДС13",
                        rfp_key="13",
                        ul_folder="согл УЛ ДС13",
                        mode=MODE_WHOLE,
                        block_index=1,
                    ),
                ),
            ),
            DsRegistryRow(
                status=STATUS_ACTIVE,
                source_id="13",
                excel_row=3,
                relations=(
                    DsRegistryRelation(
                        group_id="ДС13",
                        rfp_key="13",
                        ul_folder="согл УЛ ДС13",
                        mode=MODE_WHOLE,
                        block_index=1,
                    ),
                ),
            ),
        ]
        dup_codes = {item.code for item in validate_registry_rows(duplicate).issues}
        self.assertIn(ISSUE_DUPLICATE_SOURCE_ID, dup_codes)

        partial = [
            DsRegistryRow(
                status=STATUS_ACTIVE,
                source_id="1",
                excel_row=2,
                relations=(
                    DsRegistryRelation(
                        group_id="",
                        rfp_key="",
                        ul_folder="согл УЛ ДС1",
                        mode=MODE_WHOLE,
                        block_index=1,
                    ),
                ),
            )
        ]
        partial_codes = {item.code for item in validate_registry_rows(partial).issues}
        self.assertIn(ISSUE_PARTIAL_BLOCK, partial_codes)

        gap = [
            DsRegistryRow(
                status=STATUS_ACTIVE,
                source_id="8",
                excel_row=2,
                relations=(
                    DsRegistryRelation(
                        group_id="",
                        rfp_key="",
                        ul_folder="",
                        mode="",
                        block_index=1,
                    ),
                    DsRegistryRelation(
                        group_id="ДС81",
                        rfp_key="81",
                        ul_folder="согл УЛ ДС81",
                        mode=MODE_WHOLE,
                        block_index=2,
                    ),
                ),
            )
        ]
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            path = Path(raw) / "gap.xlsx"
            write_registry_workbook(path, gap)
            loaded = load_registry(path)
            self.assertFalse(loaded.validation.is_ok)
            self.assertIn(
                ISSUE_GAPPED_BLOCK,
                {item.code for item in loaded.validation.issues},
            )

        mismatch = [
            DsRegistryRow(
                status=STATUS_ACTIVE,
                source_id="13",
                excel_row=2,
                relations=(
                    DsRegistryRelation(
                        group_id="ДС13",
                        rfp_key="13",
                        ul_folder="согл УЛ ДС13",
                        mode=MODE_WHOLE,
                        block_index=1,
                    ),
                ),
            ),
            DsRegistryRow(
                status=STATUS_ACTIVE,
                source_id="47",
                excel_row=3,
                relations=(
                    DsRegistryRelation(
                        group_id="ДС13",
                        rfp_key="13",
                        ul_folder="согл УЛ ДС13 доп",
                        mode=MODE_WHOLE,
                        block_index=1,
                    ),
                ),
            ),
        ]
        mismatch_codes = {item.code for item in validate_registry_rows(mismatch).issues}
        self.assertIn(ISSUE_GROUP_MISMATCH, mismatch_codes)

    def test_ul_identity_roundtrip_and_mismatch(self) -> None:
        self.assertEqual(canonical_supply_group_id(13), "ДС13")
        self.assertEqual(canonical_supply_group_id(4905), "ДС4905")
        self.assertEqual(canonical_supply_group_id("4905"), "ДС4905")
        self.assertEqual(no_ul_group_id("101"), "NO_UL:101")

        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            path = Path(raw) / "identity_roundtrip.xlsx"
            write_registry_workbook(
                path,
                [
                    DsRegistryRow(
                        status=STATUS_ACTIVE,
                        source_id="13",
                        relations=(
                            DsRegistryRelation(
                                group_id="ДС13",
                                rfp_key="13",
                                ul_folder="согл УЛ ДС13",
                                mode=MODE_WHOLE,
                                block_index=1,
                            ),
                        ),
                    ),
                    DsRegistryRow(
                        status=STATUS_ACTIVE,
                        source_id="4905_1",
                        relations=(
                            DsRegistryRelation(
                                group_id="ДС4905",
                                rfp_key="4905",
                                ul_folder="согл УЛ 4905",
                                mode=MODE_WHOLE,
                                block_index=1,
                            ),
                        ),
                    ),
                    DsRegistryRow(
                        status=STATUS_ACTIVE,
                        source_id="101",
                        relations=(
                            DsRegistryRelation(
                                group_id=no_ul_group_id("101"),
                                rfp_key="101",
                                ul_folder="",
                                mode=MODE_NO_UL,
                                block_index=1,
                            ),
                        ),
                    ),
                ],
            )
            loaded = load_registry(path)
            self.assertTrue(loaded.validation.is_ok)
            self.assertEqual(loaded.rows[0].relations[0].group_id, "ДС13")
            self.assertEqual(loaded.rows[0].relations[0].rfp_key, "13")
            self.assertEqual(loaded.rows[1].relations[0].group_id, "ДС4905")
            self.assertEqual(loaded.rows[1].relations[0].rfp_key, "4905")
            self.assertEqual(loaded.rows[2].relations[0].group_id, "NO_UL:101")

        bare_group = [
            DsRegistryRow(
                status=STATUS_ACTIVE,
                source_id="13",
                excel_row=2,
                relations=(
                    DsRegistryRelation(
                        group_id="13",
                        rfp_key="13",
                        ul_folder="согл УЛ ДС13",
                        mode=MODE_WHOLE,
                        block_index=1,
                    ),
                ),
            )
        ]
        bare_codes = {item.code for item in validate_registry_rows(bare_group).issues}
        self.assertIn(ISSUE_UL_IDENTITY_MISMATCH, bare_codes)

        wrong_key = [
            DsRegistryRow(
                status=STATUS_ACTIVE,
                source_id="13",
                excel_row=2,
                relations=(
                    DsRegistryRelation(
                        group_id="ДС13",
                        rfp_key="47",
                        ul_folder="согл УЛ ДС13",
                        mode=MODE_WHOLE,
                        block_index=1,
                    ),
                ),
            )
        ]
        key_codes = {item.code for item in validate_registry_rows(wrong_key).issues}
        self.assertIn(ISSUE_UL_IDENTITY_MISMATCH, key_codes)

        unparsed_filled = [
            DsRegistryRow(
                status=STATUS_ACTIVE,
                source_id="13",
                excel_row=2,
                relations=(
                    DsRegistryRelation(
                        group_id="ДС13",
                        rfp_key="13",
                        ul_folder="папка без номера ДС",
                        mode=MODE_WHOLE,
                        block_index=1,
                    ),
                ),
            )
        ]
        unparsed_codes = {
            item.code for item in validate_registry_rows(unparsed_filled).issues
        }
        self.assertIn(ISSUE_UL_UNPARSED, unparsed_codes)
        self.assertNotIn(ISSUE_UL_IDENTITY_MISMATCH, unparsed_codes)

        no_ul = [
            DsRegistryRow(
                status=STATUS_ACTIVE,
                source_id="101",
                excel_row=2,
                relations=(
                    DsRegistryRelation(
                        group_id=no_ul_group_id("101"),
                        rfp_key="101",
                        ul_folder="",
                        mode=MODE_NO_UL,
                        block_index=1,
                    ),
                ),
            )
        ]
        no_ul_codes = {item.code for item in validate_registry_rows(no_ul).issues}
        self.assertNotIn(ISSUE_UL_IDENTITY_MISMATCH, no_ul_codes)
        self.assertNotIn(ISSUE_UL_UNPARSED, no_ul_codes)

    def test_backup_atomic_replace_and_confirm_false(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as raw:
            root = Path(raw)
            target = root / "Реестр_ДС_УЛ.xlsx"
            incoming = root / "new.xlsx"
            write_registry_workbook(
                target,
                [
                    DsRegistryRow(
                        status=STATUS_ACTIVE,
                        source_id="1",
                        relations=(
                            DsRegistryRelation(
                                group_id="ДС1",
                                rfp_key="1",
                                ul_folder="согл УЛ ДС1",
                                mode=MODE_WHOLE,
                                block_index=1,
                            ),
                        ),
                    )
                ],
            )
            write_registry_workbook(
                incoming,
                [
                    DsRegistryRow(
                        status=STATUS_ACTIVE,
                        source_id="2",
                        relations=(
                            DsRegistryRelation(
                                group_id="ДС2",
                                rfp_key="2",
                                ul_folder="согл УЛ ДС2",
                                mode=MODE_WHOLE,
                                block_index=1,
                            ),
                        ),
                    )
                ],
            )
            before = target.read_bytes()
            with self.assertRaises(DsRegistryReplaceError) as refused:
                backup_and_replace_registry(target, incoming, confirm=False)
            self.assertIn("confirm=True", str(refused.exception))
            self.assertEqual(target.read_bytes(), before)
            self.assertEqual(list(root.glob("*_backup_*")), [])

            replaced = backup_and_replace_registry(target, incoming, confirm=True)
            self.assertTrue(replaced.replaced)
            self.assertTrue(replaced.backup_path.is_file())
            self.assertEqual(load_registry(target).rows[0].source_id, "2")
            self.assertEqual(load_registry(replaced.backup_path).rows[0].source_id, "1")

            original_after = target.read_bytes()
            with patch(
                "RFQ.rfp_parts.ds_registry.os.replace",
                side_effect=PermissionError("locked"),
            ):
                with self.assertRaises(DsRegistryReplaceError) as locked:
                    backup_and_replace_registry(target, incoming, confirm=True)
            self.assertIn("занят", str(locked.exception))
            self.assertEqual(target.read_bytes(), original_after)
            self.assertEqual(load_registry(target).rows[0].source_id, "2")


if __name__ == "__main__":
    unittest.main()
