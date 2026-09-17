"""
Генерация JSON шаблонов этапа C (dwg_page1/2, bbb_mto_page1/2) из правил v1.

Запуск из корня проекта:
  python pdf_parsing_v2/_build_stamp_templates.py

Координаты bbox_mm: [left_from_right, bottom_from_bottom, right_from_right, top_from_bottom] мм
от правого нижнего угла рамки (как elements_coordinates.py, без borders_margin в v2).
"""
from __future__ import annotations

import json
import os


W = 524.28  # ширина штампа мм для перевода ячеек cell_lists → bbox_mm (DWG стр.1)


def bb_cell(cell: list[float]) -> list[float]:
    x1, y1, x2, y2 = cell
    return [round(W - x2, 2), round(y1, 2), round(W - x1, 2), round(y2, 2)]


def add_field(
    fields: list[dict],
    fid: str,
    label: str,
    bbox: list[float],
    clean: str | None,
    expected: str = "optional",
) -> None:
    o: dict = {"id": fid, "label": label, "bbox_mm": bbox, "expected": expected}
    if clean is not None:
        o["clean"] = clean
    fields.append(o)


def split18(
    fields: list[dict],
    prefix: str,
    id_suffix: str,
    label_prefix: str,
    x0: float,
    x1: float,
    yb: float,
    yt: float,
    clean_key: str,
) -> None:
    step = (yt - yb) / 3.0
    for r in range(3):
        y0 = yb + r * step
        y1 = yb + (r + 1) * step
        add_field(
            fields,
            f"{prefix}_{r + 1}_{id_suffix}",
            f"{label_prefix} {r + 1}",
            [x0, round(y0, 2), x1, round(y1, 2)],
            clean_key,
            "optional",
        )


def build_dwg_page1_fields() -> list[dict]:
    fields: list[dict] = []
    add_field(fields, "1_DOC_TITLE", "Обозначение документа (шифр)", [120, 45, 0, 55], "doc_title", "required")
    add_field(fields, "2_Facility_name", "Наименование объекта", [90, 33, 0, 45], "facility_name", "optional")
    add_field(fields, "3_Unit_title_name", "Наименование установки", [100, 15, 55, 30], "unit_title", "optional")
    add_field(fields, "4_Document_name", "Наименование документа", [100, 0, 55, 15], "document_name", "optional")
    add_field(fields, "5_Documentation_type", "Вид документации", [45, 15, 35, 25], "documentation_type", "optional")
    add_field(fields, "6_1_Sheet_number", "Номер листа", [35, 20, 20, 25], "sheet_number_6_1", "required")
    add_field(fields, "6_2_Quantity_of_sheets", "Листов в документе", [35, 15, 20, 20], "sheet_number_6_2", "optional")
    add_field(fields, "7_Total_number_of_sheets", "Всего листов", [17, 15, 0, 25], "total_number_of_sheets", "optional")
    add_field(fields, "26_Document_Revision", "Текущая ревизия документа", [17, 55, 0, 70], "revision", "required")

    col0 = [
        [155.88, 226.8, 184.2, 240.96],
        [155.88, 240.96, 184.2, 255.12],
        [155.88, 255.12, 184.2, 269.28],
        [155.88, 269.28, 184.2, 283.44],
        [155.88, 283.44, 184.2, 297.6],
    ]
    col2 = [
        [113.39, 226.8, 155.88, 240.96],
        [113.39, 240.96, 155.88, 255.12],
        [113.39, 255.12, 155.88, 269.28],
        [113.39, 269.28, 155.88, 283.44],
        [113.39, 297.6, 155.88, 311.88],
    ]
    col5 = [
        [382.56, 226.8, 425.16, 240.96],
        [382.56, 240.96, 425.16, 255.12],
        [382.56, 255.12, 425.16, 269.28],
        [382.56, 269.28, 425.16, 283.44],
        [382.56, 283.44, 425.16, 297.6],
    ]
    for i, c in enumerate(col0):
        add_field(fields, f"Position_{i + 1}", f"Должность {i + 1}", bb_cell(c), None, "optional")
    for i, c in enumerate(col2):
        add_field(fields, f"Surname_{i + 1}", f"Фамилия {i + 1}", bb_cell(c), None, "optional")
    for i, c in enumerate(col5):
        add_field(
            fields,
            f"10_{i + 1}_Signatures_Date",
            f"Дата подписи {i + 1}",
            bb_cell(c),
            "signatures_date_10",
            "optional",
        )

    split18(fields, "18_1", "Current_Revision", "Текущая ревизия (блок 18.1)", 185, 170, 90.0, 111.0, "current_revision_18_1")
    split18(fields, "18_2", "Revision_Date", "Дата ревизии (блок 18.2)", 170, 155, 90.0, 111.0, "revision_date_18_2")
    split18(fields, "18_3", "Purpose_of_issue", "Назначение выпуска (блок 18.3)", 145, 60, 90.0, 111.0, "purpose_of_issue_18_3")

    add_field(fields, "50_File_Name_Stamp", "Имя файла под штампом", [80, -50, 285, -1], "file_name_stamp", "required")
    add_field(fields, "51_Page_Format", "Формат страницы", [-50, -50, 80, -1], "page_format", "optional")
    return fields


def build_bbb_page1_fields() -> list[dict]:
    """Те же поля, что DWG стр.1, но блоки 18.x и даты подписей — ветки list_of_BBB из elements_coordinates.py."""
    fields = build_dwg_page1_fields()
    by_id = {f["id"]: f for f in fields}

    # Signatures_Date_10: база 130,23,112,30 + расширение ±5 мм (BBB)
    for i in range(1, 6):
        k = f"10_{i}_Signatures_Date"
        if k in by_id:
            by_id[k]["bbox_mm"] = [135, 18, 107, 30]

    # Удалить старые 18_* и вставить BBB-версии
    fields[:] = [f for f in fields if not f["id"].startswith(("18_1_", "18_2_", "18_3_"))]
    insert_at = next(i for i, f in enumerate(fields) if f["id"] == "50_File_Name_Stamp")
    sub: list[dict] = []
    split18(sub, "18_1", "Current_Revision", "Текущая ревизия (блок 18.1)", 190, 156, 85.0, 111.0, "current_revision_18_1")
    split18(sub, "18_2", "Revision_Date", "Дата ревизии (блок 18.2)", 170, 144, 85.0, 111.0, "revision_date_18_2")
    split18(sub, "18_3", "Purpose_of_issue", "Назначение выпуска (блок 18.3)", 145, 60, 85.0, 111.0, "purpose_of_issue_18_3")
    for j, item in enumerate(sub):
        fields.insert(insert_at + j, item)

    return fields


def main() -> None:
    root = os.path.dirname(os.path.abspath(__file__))
    tpl = os.path.join(root, "templates")
    os.makedirs(tpl, exist_ok=True)

    templates = [
        {
            "path": os.path.join(tpl, "dwg_page1.json"),
            "name": "DWG стр.1 (ГОСТ 21.1101, координаты elements_coordinates + cell_lists)",
            "doc_types": ["WIR", "LAY", "CAE", "GA", "PL", "DW", "NI"],
            "page_selector": "first",
            "fields": build_dwg_page1_fields(),
        },
        {
            "path": os.path.join(tpl, "dwg_page2.json"),
            "name": "DWG стр.2+ (малый штамп, pageNum>1, не BBB)",
            "doc_types": ["WIR", "LAY", "CAE", "GA", "PL", "DW", "NI"],
            "page_selector": "rest",
            "fields": [
                {
                    "id": "1_DOC_TITLE",
                    "label": "Обозначение документа (шифр)",
                    "bbox_mm": [110, 0, 25, 15],
                    "expected": "required",
                    "clean": "doc_title",
                },
                {
                    "id": "6_1_Sheet_number",
                    "label": "Номер листа",
                    "bbox_mm": [8, 0, 0, 8],
                    "expected": "required",
                    "clean": "sheet_number_6_1",
                },
                {
                    "id": "6_2_Quantity_of_sheets",
                    "label": "Листов в документе",
                    "bbox_mm": [35, 15, 20, 20],
                    "expected": "optional",
                    "clean": "sheet_number_6_2",
                },
                {
                    "id": "26_Document_Revision",
                    "label": "Текущая ревизия",
                    "bbox_mm": [20, 0, 12, 8],
                    "expected": "required",
                    "clean": "revision",
                },
            ],
        },
        {
            "path": os.path.join(tpl, "bbb_mto_page1.json"),
            "name": "BOE/BOM/BOQ/MTO/OD/CJ стр.1 (ветки BBB в elements_coordinates)",
            "doc_types": ["BOE", "BOM", "BOQ", "MTO", "OD", "CJ"],
            "page_selector": "first",
            "fields": build_bbb_page1_fields(),
        },
        {
            "path": os.path.join(tpl, "bbb_mto_page2.json"),
            "name": "BOE/BOM/BOQ/MTO/OD/CJ стр.2+ (как get_cur_page_MTO_BBB: зоны pdfminer)",
            "doc_types": ["BOE", "BOM", "BOQ", "MTO", "OD", "CJ"],
            "page_selector": "rest",
            "fields": [
                {
                    "id": "1_DOC_TITLE",
                    "label": "Обозначение документа (шифр, верхняя полоса справа)",
                    "bbox_mm": [120, 250, 0, 297],
                    "expected": "required",
                    "clean": "doc_title",
                },
                {
                    "id": "6_1_Sheet_number",
                    "label": "Номер листа (та же зона, что шифр в v1 BBB)",
                    "bbox_mm": [120, 250, 0, 297],
                    "expected": "required",
                    "clean": "sheet_number_6_1",
                },
                {
                    "id": "6_2_Quantity_of_sheets",
                    "label": "Листов в документе",
                    "bbox_mm": [120, 250, 0, 297],
                    "expected": "optional",
                    "clean": "sheet_number_6_2",
                },
                {
                    "id": "50_File_Name_Stamp",
                    "label": "Имя файла (полная ширина листа, MTO/BOE ветка)",
                    "bbox_mm": [297, -50, 0, -1],
                    "expected": "required",
                    "clean": "file_name_stamp",
                },
            ],
        },
    ]

    for spec in templates:
        doc = {
            "schema_version": 2,
            "name": spec["name"],
            "doc_types": spec["doc_types"],
            "page_selector": spec["page_selector"],
            "priority": 10,
            "origin": "frame_bottom_right",
            "padding_mm": 0.5,
            "fields": spec["fields"],
        }
        with open(spec["path"], "w", encoding="utf-8") as f:
            json.dump(doc, f, ensure_ascii=False, indent=2)
        print(spec["path"], "fields", len(spec["fields"]))


if __name__ == "__main__":
    main()
