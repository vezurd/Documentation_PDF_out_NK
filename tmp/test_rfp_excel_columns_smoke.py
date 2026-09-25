"""Smoke tests for Step4 Excel column templates (no real config writes)."""

from __future__ import annotations

import copy
import os
import sys
import unittest
from pathlib import Path
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("PYTHONUTF8", "1")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from base.tables_columns import (
    DS_ACTUAL,
    DS_MANAGER,
    DS_NAME,
    DS_NUMBER,
    EQUIPMENT_TYPE_STATUS,
    POSITION_STATUS,
    RFP_SUPPLY_STATUS,
    TAG_EFFECTIVE,
    TYPE_MARK,
)
from RFQ.tags_rfp_compare.step4.step4_6_save_match_result_to_excel import (
    OUTPUT_COLUMNS_CONFIG,
)
from RFQ.tags_rfp_compare.step4 import step4_excel_columns as cols
from RFQ.tags_rfp_compare.step4.step4_excel_columns import (
    DEFAULT_TEMPLATE_ID,
    Step4ColumnTemplateError,
    builtin_column_settings,
    create_template,
    delete_template,
    normalize_template_columns,
    outline_options_for_visible,
    resolve_active_column_defs,
    settings_to_column_defs,
    update_template_columns,
)


def _by_name(settings: list[dict]) -> dict[str, dict]:
    return {item["col_name"]: item for item in settings}


def _visible_index(defs, col_name: str) -> int:
    visible = [item.col_name for item in defs if item.output]
    return visible.index(col_name)


class _ConfigSandbox:
    """In-memory load_config/save_config that never touches the real JSON."""

    def __init__(self) -> None:
        self.data: dict = {
            "paths": {"rfp_path": "KEEP_ME"},
            "step4": {"debug": True},
            "step4_excel_columns": {
                "active_id": DEFAULT_TEMPLATE_ID,
                "templates": [],
            },
        }
        self.save_calls = 0

    def load(self) -> dict:
        return copy.deepcopy(self.data)

    def save(self, config: dict) -> bool:
        self.save_calls += 1
        self.data = copy.deepcopy(config)
        return True


class Step4ExcelColumnsSmokeTest(unittest.TestCase):
    def test_builtin_status_group_excludes_equipment_type(self) -> None:
        settings = builtin_column_settings()
        by_name = _by_name(settings)
        self.assertEqual(by_name[EQUIPMENT_TYPE_STATUS]["group_id"], "")
        status_ids = {
            by_name[TAG_EFFECTIVE]["group_id"],
            by_name[POSITION_STATUS]["group_id"],
            by_name[RFP_SUPPLY_STATUS]["group_id"],
        }
        self.assertEqual(len(status_ids), 1)
        gid = next(iter(status_ids))
        self.assertTrue(gid)
        self.assertTrue(by_name[TAG_EFFECTIVE]["group_collapsed"])
        self.assertTrue(by_name[POSITION_STATUS]["group_collapsed"])
        self.assertTrue(by_name[RFP_SUPPLY_STATUS]["group_collapsed"])
        visible = [item["col_name"] for item in settings if item["output"]]
        supply_at = visible.index(RFP_SUPPLY_STATUS)
        self.assertEqual(visible[supply_at + 1], EQUIPMENT_TYPE_STATUS)
        type_mark = next(d for d in OUTPUT_COLUMNS_CONFIG if d.col_name == TYPE_MARK)
        self.assertEqual(type_mark.width, 20)

    def test_outline_status_and_ds_name_groups(self) -> None:
        defs = settings_to_column_defs(builtin_column_settings())
        opts = outline_options_for_visible(defs)
        for name in (TAG_EFFECTIVE, POSITION_STATUS, RFP_SUPPLY_STATUS):
            idx = _visible_index(defs, name)
            self.assertEqual(opts[idx].get("level"), 1, name)
            self.assertTrue(opts[idx].get("hidden"), name)
            self.assertNotIn("collapsed", opts[idx], name)
        eq_idx = _visible_index(defs, EQUIPMENT_TYPE_STATUS)
        self.assertEqual(opts[eq_idx].get("level"), 0)
        self.assertTrue(opts[eq_idx].get("collapsed"))
        self.assertNotEqual(opts[eq_idx].get("level"), 1)

        for name in (DS_NAME, DS_ACTUAL, DS_NUMBER):
            idx = _visible_index(defs, name)
            self.assertEqual(opts[idx].get("level"), 1, name)
            self.assertTrue(opts[idx].get("hidden"), name)
            self.assertNotIn("collapsed", opts[idx], name)
        mgr_idx = _visible_index(defs, DS_MANAGER)
        self.assertEqual(opts[mgr_idx].get("level"), 0)
        self.assertTrue(opts[mgr_idx].get("collapsed"))

    def test_reorder_round_trip(self) -> None:
        settings = builtin_column_settings()
        visible = [item for item in settings if item["output"]]
        a = visible[4]["col_name"]
        b = visible[5]["col_name"]
        names = [item["col_name"] for item in settings]
        i, j = names.index(a), names.index(b)
        names[i], names[j] = names[j], names[i]
        by_name = _by_name(settings)
        reordered = [copy.deepcopy(by_name[name]) for name in names]
        normalized = normalize_template_columns(reordered)
        self.assertEqual([item["col_name"] for item in normalized], names)
        defs = settings_to_column_defs(normalized)
        self.assertEqual([defn.col_name for defn in defs], names)
        vis_defs = [defn.col_name for defn in defs if defn.output]
        self.assertEqual(vis_defs.index(b), vis_defs.index(a) - 1)

    def test_unknown_dropped_and_missing_appended(self) -> None:
        settings = builtin_column_settings()
        drop_name = EQUIPMENT_TYPE_STATUS
        saved = [item for item in settings if item["col_name"] != drop_name]
        saved.append(
            {
                "col_name": "NO_SUCH_COLUMN",
                "output": True,
                "width": 10,
                "header_label": "нет",
                "group_id": "",
                "group_collapsed": False,
            }
        )
        normalized = normalize_template_columns(saved)
        names = [item["col_name"] for item in normalized]
        self.assertNotIn("NO_SUCH_COLUMN", names)
        self.assertIn(drop_name, names)
        self.assertEqual(names[-1], drop_name)
        builtin_names = [item["col_name"] for item in settings]
        self.assertEqual(sorted(names), sorted(builtin_names))

    def test_missing_active_id_resolves_to_builtin(self) -> None:
        sandbox = _ConfigSandbox()
        sandbox.data["step4_excel_columns"]["active_id"] = "t_does_not_exist"
        with mock.patch.object(cols, "load_config", sandbox.load), mock.patch.object(
            cols, "save_config", sandbox.save
        ):
            defs = resolve_active_column_defs()
        builtin = settings_to_column_defs(builtin_column_settings())
        self.assertEqual([d.col_name for d in defs], [d.col_name for d in builtin])
        self.assertEqual(
            [d.header_label for d in defs],
            [d.header_label for d in builtin],
        )

    def test_delete_and_update_default_refused(self) -> None:
        sandbox = _ConfigSandbox()
        with mock.patch.object(cols, "load_config", sandbox.load), mock.patch.object(
            cols, "save_config", sandbox.save
        ):
            with self.assertRaises(Step4ColumnTemplateError):
                delete_template(DEFAULT_TEMPLATE_ID)
            with self.assertRaises(Step4ColumnTemplateError):
                update_template_columns(DEFAULT_TEMPLATE_ID, builtin_column_settings())
            stored = sandbox.data["step4_excel_columns"]["templates"]
            self.assertFalse(
                any(item.get("id") == DEFAULT_TEMPLATE_ID for item in stored)
            )

    def test_persistence_does_not_drop_other_keys(self) -> None:
        sandbox = _ConfigSandbox()
        with mock.patch.object(cols, "load_config", sandbox.load), mock.patch.object(
            cols, "save_config", sandbox.save
        ):
            tid = create_template("Мой вид", builtin_column_settings())
            self.assertTrue(tid.startswith("t_"))
            self.assertEqual(sandbox.data["paths"]["rfp_path"], "KEEP_ME")
            self.assertTrue(sandbox.data["step4"]["debug"])
            stored = sandbox.data["step4_excel_columns"]
            self.assertEqual(stored["active_id"], tid)
            self.assertEqual(len(stored["templates"]), 1)
            self.assertNotEqual(stored["templates"][0]["id"], DEFAULT_TEMPLATE_ID)
            self.assertGreater(sandbox.save_calls, 0)


class RfpColumnsPanelSmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_default_template_disables_save_and_lists_builtin(self) -> None:
        sandbox = _ConfigSandbox()
        with mock.patch.object(cols, "load_config", sandbox.load), mock.patch.object(
            cols, "save_config", sandbox.save
        ):
            from PySide6.QtWidgets import QLabel, QPushButton

            from ds_compare_center.rfp_columns_panel import (
                RfpColumnsPanel,
                _GroupBox,
                heal_column_groups,
                plan_sheet_insertion,
            )

            panel = RfpColumnsPanel()
            panel.show()
            self.app.processEvents()
            try:
                self.assertFalse(panel._btn_save.isEnabled())
                labels = [panel._combo.itemText(i) for i in range(panel._combo.count())]
                self.assertIn("По умолчанию", labels)
                self.assertEqual(panel._combo.itemData(0), DEFAULT_TEMPLATE_ID)
                self.assertIn(_DEFAULT_HINT_SNIPPET, panel._hint.text())
                self.assertIsNotNone(panel.findChild(QLabel, "sheet-row") or panel.findChild(type(panel._sheet_row)))
                toggles = panel.findChildren(QPushButton, "group-toggle")
                self.assertTrue(any(btn.text() == "+" for btn in toggles))
                stays = [
                    label.text()
                    for label in panel.findChildren(QLabel, "stays-badge")
                    if not label.isHidden()
                ]
                self.assertIn("останется", stays)
                sheet_names = {
                    card.col_name()
                    for card in panel._cards
                    if card.parent() is panel._sheet_row or card.parent().objectName() == "group-box"
                }
                unused_names = {
                    card.col_name()
                    for card in panel._cards
                    if card.parent() is panel._unused_row
                }
                self.assertTrue(sheet_names)
                self.assertTrue(unused_names)
                self.assertFalse(sheet_names & unused_names)
                self.assertEqual(
                    panel._sheet_scroll.horizontalScrollBarPolicy(),
                    Qt.ScrollBarPolicy.ScrollBarAsNeeded,
                )
                self.assertEqual(
                    panel._sheet_scroll.verticalScrollBarPolicy(),
                    Qt.ScrollBarPolicy.ScrollBarAlwaysOff,
                )
                self.assertEqual(
                    panel._unused_scroll.horizontalScrollBarPolicy(),
                    Qt.ScrollBarPolicy.ScrollBarAlwaysOff,
                )
                self.assertEqual(
                    panel._unused_scroll.verticalScrollBarPolicy(),
                    Qt.ScrollBarPolicy.ScrollBarAsNeeded,
                )
                panel.resize(900, 720)
                self.app.processEvents()
                unused_cards = [
                    card for card in panel._cards if card.parent() is panel._unused_row
                ]
                self.assertGreater(len(unused_cards), 1)
                limit = panel._unused_scroll.viewport().width()
                self.assertGreater(limit, 0)
                self.assertLessEqual(panel._unused_row.width(), limit + 2)
                for card in unused_cards:
                    self.assertLessEqual(card.geometry().right(), limit)
                    self.assertLessEqual(card.geometry().bottom(), panel._unused_row.minimumHeight())
                ys = {card.geometry().top() for card in unused_cards}
                self.assertGreater(len(ys), 1)
                for card in unused_cards:
                    self.assertTrue(card.isVisible())
                    self.assertGreaterEqual(card.height(), 40)
                toggles = panel.findChildren(QPushButton, "group-toggle")
                self.assertTrue(toggles)
                for btn in toggles:
                    box = btn.parentWidget()
                    caption = box.findChild(QLabel, "group-caption")
                    self.assertIsNotNone(caption)
                    self.assertEqual(caption.text(), "группа")
                    self.assertLessEqual(btn.y(), 12)
                    self.assertGreaterEqual(caption.geometry().left(), btn.geometry().right() - 2)
                    self.assertFalse(btn.geometry().intersects(caption.geometry()))
                tid = create_template("проверка ширины", panel._columns_payload())
                panel.reload_from_disk()
                panel._load_columns_for_id(tid)
                panel._apply_readonly()
                self.app.processEvents()
                self.assertFalse(panel._is_readonly())
                boxes = panel._sheet_row.findChildren(_GroupBox)
                self.assertTrue(boxes)
                for box in boxes:
                    self.assertGreaterEqual(box.width() + 8, box.minimumSizeHint().width())
                unused_after = [
                    card for card in panel._cards if card.parent() is panel._unused_row
                ]
                self.assertGreater(len(unused_after), 1)
                self.assertGreaterEqual(
                    panel._sheet_row.minimumWidth(),
                    panel._sheet_layout.totalMinimumSize().width() - 4,
                )
                bar = panel._sheet_scroll.horizontalScrollBar()
                self.app.processEvents()
                target = min(400, bar.maximum())
                self.assertGreater(bar.maximum(), 50)
                bar.setValue(target)
                panel._toggle_group(boxes[0].group_id)
                self.app.processEvents()
                self.assertEqual(bar.value(), target)
                name = str(panel._columns[0].get("col_name") or "")
                panel._move_names([name], visible_index=4, unused_at=None, group_id="")
                self.app.processEvents()
                self.assertEqual(bar.value(), target)
            finally:
                panel.close()
                panel.deleteLater()
                self.app.processEvents()


class DropPlanSmokeTest(unittest.TestCase):
    def test_cursor_near_group_joins_and_far_stays_out(self) -> None:
        from ds_compare_center.rfp_columns_panel import plan_sheet_insertion

        units = [
            {"left": 100, "right": 180, "start": 0, "end": 3, "group_id": "g1"},
            {"left": 190, "right": 270, "start": 3, "end": 4, "group_id": ""},
        ]
        insert_at, gid = plan_sheet_insertion(units, 160)
        self.assertEqual((insert_at, gid), (3, "g1"))
        insert_at, gid = plan_sheet_insertion(units, 220)
        self.assertEqual((insert_at, gid), (3, ""))

    def test_heal_drops_singleton_and_keeps_collapsed_run(self) -> None:
        from ds_compare_center.rfp_columns_panel import heal_column_groups

        columns = [
            {"output": True, "group_id": "g1", "group_collapsed": True},
            {"output": True, "group_id": "g1", "group_collapsed": False},
            {"output": True, "group_id": "", "group_collapsed": False},
            {"output": True, "group_id": "g9", "group_collapsed": True},
            {"output": False, "group_id": "g1", "group_collapsed": True},
        ]
        heal_column_groups(columns)
        self.assertEqual(columns[0]["group_id"], "g1")
        self.assertTrue(columns[0]["group_collapsed"])
        self.assertTrue(columns[1]["group_collapsed"])
        self.assertEqual(columns[3]["group_id"], "")
        self.assertEqual(columns[4]["group_id"], "")


_DEFAULT_HINT_SNIPPET = "Шаблон по умолчанию только для просмотра"


if __name__ == "__main__":
    unittest.main(verbosity=2)
