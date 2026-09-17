"""Checks for Комплекты default column-order template."""

from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication, QTableWidget

from rd_catalog.config import load_config
from rd_catalog.db import CatalogDatabase
from rd_catalog.kits_table_layout import (
    KitsLayoutSaveResult,
    KitsTableLayout,
    load_default_kits_table_layout,
    load_kits_table_layout,
    merge_header_order,
    parse_kits_table_layout,
    resolve_kits_table_layout,
    save_default_kits_table_layout,
    write_kits_table_layout,
)
from rd_catalog import window as rd_window
from rd_catalog.window import (
    CatalogWindow,
    _KITS_HEADERS,
    apply_kits_header_layout,
    capture_kits_header_layout,
)


def _visual_names(table: QTableWidget) -> list[str]:
    header = table.horizontalHeader()
    return [
        table.horizontalHeaderItem(header.logicalIndex(visual)).text()
        for visual in range(header.count())
    ]


def _check_parse_and_merge() -> None:
    assert parse_kits_table_layout(None) is None
    assert parse_kits_table_layout({"order": []}) is None
    assert parse_kits_table_layout({"order": [""]}) is None
    assert parse_kits_table_layout({"order": ["A", 1]}) is None
    parsed = parse_kits_table_layout(
        {
            "version": 1,
            "order": ["Марка", "Титул", "Марка", "Нет такой"],
            "widths": {"Титул": 80, "Марка": "12", "bad": -1, "skip": "x"},
        }
    )
    assert parsed is not None
    assert parsed.order == ("Марка", "Титул", "Нет такой")
    assert parsed.widths == {"Титул": 80, "Марка": 12}
    assert merge_header_order(
        ("Марка", "Титул", "Нет такой"),
        ("Титул", "Марка", "Сводка"),
    ) == ("Марка", "Титул", "Сводка")

    packaged = load_kits_table_layout(
        Path(__file__).resolve().parents[1]
        / "rd_catalog"
        / "kits_table_layout.json"
    )
    assert packaged is not None
    assert "АН МТО" in packaged.order
    assert "MTO · рев." in packaged.order
    assert "Ок" in packaged.order
    assert "Рабочая рев. РД" in packaged.order
    assert packaged.widths.get("АН МТО", 0) >= 80
    assert packaged.widths.get("Рабочая рев. РД", 0) >= 80
    assert packaged.widths.get("Ок", 0) >= 40
    assert set(packaged.order) <= set(_KITS_HEADERS)
    assert _KITS_HEADERS[_KITS_HEADERS.index("Марка") + 1] == "Ок"
    assert _KITS_HEADERS[_KITS_HEADERS.index("Сверка Авто МТО") + 1] == "АН МТО"
    assert merge_header_order(
        ("Робот · рев.", "АН", "Титул"),
        ("Титул", "Робот МТО · рев.", "АН МТО"),
    ) == ("Робот МТО · рев.", "АН МТО", "Титул")
    assert merge_header_order(
        ("Титул", "Марка", "РД · рев."),
        ("Титул", "Марка", "Ок", "РД · рев."),
    ) == ("Титул", "Марка", "Ок", "РД · рев.")


def _check_runtime_io(temp: Path) -> None:
    runtime = temp / "runtime"
    layout = KitsTableLayout(
        order=("Марка", "Титул"),
        widths={"Титул": 90, "Марка": 70},
    )
    saved = save_default_kits_table_layout(
        layout,
        runtime,
        packaged_path=temp / "packaged.json",
        write_packaged=True,
    )
    runtime_path = saved.runtime_path
    packaged_path = saved.packaged_path
    assert saved.ok
    assert runtime_path.is_file()
    assert packaged_path is not None and packaged_path.is_file()
    loaded = load_default_kits_table_layout(
        runtime, packaged_path=temp / "missing.json"
    )
    assert loaded is not None
    assert loaded.order == ("Марка", "Титул")
    assert loaded.widths == {"Титул": 90, "Марка": 70}
    only_factory = load_default_kits_table_layout(
        temp / "empty_runtime",
        packaged_path=packaged_path,
    )
    assert only_factory is not None
    assert only_factory.order == ("Марка", "Титул")
    write_kits_table_layout(runtime_path, layout)
    raw = json.loads(runtime_path.read_text(encoding="utf-8"))
    assert raw["version"] == 1
    assert raw["order"] == ["Марка", "Титул"]
    resolved = resolve_kits_table_layout(
        runtime,
        ("Титул", "Марка", "Сводка"),
        packaged_path=temp / "missing.json",
    )
    assert resolved.order == ("Марка", "Титул", "Сводка")
    assert resolved.widths == {"Титул": 90, "Марка": 70}
    empty = resolve_kits_table_layout(
        temp / "no_json",
        ("Титул", "Марка"),
        packaged_path=temp / "missing.json",
    )
    assert empty.order == ("Титул", "Марка")
    assert empty.widths == {}

    blocked = temp / "blocked.json"
    blocked.mkdir()
    failed = save_default_kits_table_layout(
        layout,
        temp / "runtime_partial",
        packaged_path=blocked,
        write_packaged=True,
    )
    assert failed.runtime_path.is_file()
    assert failed.packaged_path is None
    assert failed.packaged_error
    assert not failed.ok

    fallback_target = temp / "replace_busy.json"
    write_kits_table_layout(
        fallback_target,
        KitsTableLayout(order=("Титул",), widths={"Титул": 11}),
    )
    with mock.patch(
        "rd_catalog.kits_table_layout.os.replace",
        side_effect=OSError("sharing violation"),
    ):
        write_kits_table_layout(
            fallback_target,
            KitsTableLayout(order=("Марка",), widths={"Марка": 12}),
        )
    fallback = load_kits_table_layout(fallback_target)
    assert fallback is not None
    assert fallback.order == ("Марка",)
    assert fallback.widths == {"Марка": 12}


def _check_header_apply(app: QApplication) -> None:
    del app
    table = QTableWidget(0, len(_KITS_HEADERS))
    table.setHorizontalHeaderLabels(list(_KITS_HEADERS))
    header = table.horizontalHeader()
    header.setSectionsMovable(True)
    custom = KitsTableLayout(
        order=("РД · рев.", "Титул", "Марка", "нет-колонки"),
        widths={"Титул": 64, "РД · рев.": 88},
    )
    apply_kits_header_layout(header, list(_KITS_HEADERS), custom)
    visual = _visual_names(table)
    assert visual[0] == "РД · рев."
    assert visual[1] == "Титул"
    assert visual[2] == "Марка"
    assert visual[-1] == "РД · дата файла"
    assert set(visual) == set(_KITS_HEADERS)
    assert header.sectionSize(_KITS_HEADERS.index("Титул")) == 64
    assert header.sectionSize(_KITS_HEADERS.index("РД · рев.")) == 88
    captured = capture_kits_header_layout(header, list(_KITS_HEADERS))
    assert captured.order[0] == "РД · рев."
    assert captured.widths["Титул"] == 64
    table.deleteLater()


def _make_window(root: Path) -> CatalogWindow:
    override = root / "config.json"
    if not override.is_file():
        for name in ("rd", "sq", "robot", "runtime"):
            (root / name).mkdir(exist_ok=True)
        override.write_text(
            json.dumps(
                {
                    "rd_root": str(root / "rd"),
                    "sq_root": str(root / "sq"),
                    "robot_root": str(root / "robot"),
                    "runtime_dir": str(root / "runtime"),
                    "db_path": str(root / "runtime" / "catalog.sqlite"),
                    "robot_flat_structure": True,
                    "skip_dirs": ["old"],
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
    config = load_config(override)
    database = CatalogDatabase(config.db_path)
    database.initialize()
    return CatalogWindow(config, database)


def _check_window_fallback(app: QApplication, root: Path) -> None:
    settings_dir = root / "settings"
    settings_dir.mkdir(exist_ok=True)
    QSettings.setDefaultFormat(QSettings.Format.IniFormat)
    QSettings.setPath(
        QSettings.Format.IniFormat,
        QSettings.Scope.UserScope,
        str(settings_dir),
    )
    rd_window._SETTINGS_APPLICATION = "rd_catalog_kits_layout_test"
    throwaway = QSettings(
        rd_window._SETTINGS_ORGANIZATION,
        rd_window._SETTINGS_APPLICATION,
    )
    throwaway.clear()
    throwaway.sync()

    custom_order = (
        "Титул",
        "Марка",
        "РД · рев.",
        "Авто МТО",
        "Сверка Авто МТО",
    )
    write_kits_table_layout(
        root / "runtime" / "kits_table_layout.json",
        KitsTableLayout(
            order=custom_order,
            widths={"Титул": 77, "РД · рев.": 91},
        ),
    )
    window = _make_window(root)
    window.show()
    app.processEvents()
    visual = _visual_names(window._kits_table)
    assert visual[:6] == [
        "Титул",
        "Марка",
        "Ок",
        "РД · рев.",
        "Авто МТО",
        "Сверка Авто МТО",
    ]
    header = window._kits_table.horizontalHeader()
    assert header.sectionSize(_KITS_HEADERS.index("Титул")) == 77
    assert header.sectionSize(_KITS_HEADERS.index("РД · рев.")) == 91

    last = header.count() - 1
    header.moveSection(header.visualIndex(last), 3)
    header.resizeSection(0, 55)
    written = window._write_kits_column_template(write_packaged=False)
    runtime_path = written.runtime_path
    assert written.ok
    assert runtime_path.is_file()
    assert written.packaged_path is None
    saved = load_kits_table_layout(runtime_path)
    assert saved is not None
    assert saved.order[3] == _KITS_HEADERS[last]
    assert saved.widths["Титул"] == 55

    def _boom(*_args, **_kwargs):
        raise AssertionError("QMessageBox must not be used for the layout button")

    original_info = rd_window.QMessageBox.information
    original_warn = rd_window.QMessageBox.warning
    rd_window.QMessageBox.information = _boom
    rd_window.QMessageBox.warning = _boom
    try:
        window._write_kits_column_template = lambda **_kw: KitsLayoutSaveResult(
            runtime_path=runtime_path,
            packaged_path=root / "runtime" / "kits_table_layout.json",
        )
        window._save_kits_column_template()
        assert rd_window._KITS_LAYOUT_OK_FILL in window._kits_layout_button.styleSheet()
        assert "Шаблон колонок сохранён" in window._log.toPlainText()

        def _fail_write(**_kw):
            raise OSError("locked")

        window._write_kits_column_template = _fail_write
        window._save_kits_column_template()
        assert (
            rd_window._KITS_LAYOUT_PROBLEM_FILL
            in window._kits_layout_button.styleSheet()
        )
        assert "не записан" in window._log.toPlainText()
        window._write_kits_column_template = lambda **_kw: KitsLayoutSaveResult(
            runtime_path=runtime_path,
            packaged_error="PermissionError: sharing violation",
        )
        window._save_kits_column_template()
        assert (
            rd_window._KITS_LAYOUT_PROBLEM_FILL
            in window._kits_layout_button.styleSheet()
        )
        assert "заводской не записан" in window._log.toPlainText()
    finally:
        rd_window.QMessageBox.information = original_info
        rd_window.QMessageBox.warning = original_warn

    window.close()
    app.processEvents()
    throwaway.setValue("window/kits_header_v10_count", 1)
    throwaway.remove("window/kits_header_v10")
    throwaway.sync()

    reopened = _make_window(root)
    reopened.show()
    app.processEvents()
    assert _visual_names(reopened._kits_table)[3] == _KITS_HEADERS[last]
    assert (
        reopened._kits_table.horizontalHeader().sectionSize(0) == 55
    )
    reopened.close()
    app.processEvents()


def main() -> None:
    """Run parse, IO, header apply, and QSettings-fallback checks."""

    _check_parse_and_merge()
    app = QApplication.instance() or QApplication([])
    _check_header_apply(app)
    with tempfile.TemporaryDirectory(prefix="rd_kits_layout_") as temp:
        root = Path(temp)
        _check_runtime_io(root)
        _check_window_fallback(app, root)
    print("RD catalog kits table layout: OK")


if __name__ == "__main__":
    main()
