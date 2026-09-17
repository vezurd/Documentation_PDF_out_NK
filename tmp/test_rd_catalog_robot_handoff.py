"""Local checks for RD catalog chat handoff dumps."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.doc_bundle import DocumentTreeLabelOptions
from rd_catalog.robot_handoff import (
    HANDOFF_HEADER,
    format_label_options,
    format_robot_handoff,
    node_kind_label,
    yes_no,
)


def main() -> None:
    """Check handoff formatting without Qt."""

    text = format_robot_handoff(
        "дерево «Все документы» (титул → марка → ревизия)",
        [
            ("Уровень", node_kind_label("revision")),
            ("Лейбл сейчас", "01-AN02 (2024.08.16) · ТДО"),
            ("Пустое", ""),
            ("Файлы", "pdf\tOD-0001.pdf\tC:\\rd\\file.pdf\ndwg\tOD-0001.dwg"),
        ],
    )
    assert text.startswith(HANDOFF_HEADER)
    assert "Где: дерево «Все документы»" in text
    assert "Уровень: ревизия (папка NN)" in text
    assert "Лейбл сейчас: 01-AN02 (2024.08.16) · ТДО" in text
    assert "Пустое:" not in text
    assert "  pdf\tOD-0001.pdf" in text
    assert text.endswith("\n")
    assert node_kind_label("title") == "титул"
    assert node_kind_label("mark") == "марка"
    assert yes_no(True) == "да"
    assert yes_no(False) == "нет"
    options = format_label_options(
        DocumentTreeLabelOptions(show_date=True, show_folder=True, show_review=False)
    )
    assert "дата=да" in options
    assert "папка NN=да" in options
    assert "статус Google=нет" in options
    assert "as-build=нет" in options
    print("RD catalog robot handoff: OK")


if __name__ == "__main__":
    main()
