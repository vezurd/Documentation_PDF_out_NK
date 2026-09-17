"""Legacy: только dwg_page1. Полный набор шаблонов — pdf_parsing_v2/_build_stamp_templates.py."""
import json
import os

W = 524.28


def bb(cell: list[float]) -> list[float]:
    x1, y1, x2, y2 = cell
    return [round(W - x2, 2), round(y1, 2), round(W - x1, 2), round(y2, 2)]


def main() -> None:
    fields: list[dict] = []

    def add(fid: str, label: str, bbox: list, clean: str | None, exp: str = "optional") -> None:
        o: dict = {"id": fid, "label": label, "bbox_mm": bbox, "expected": exp}
        if clean is not None:
            o["clean"] = clean
        fields.append(o)

    add("1_DOC_TITLE", "Обозначение документа (шифр)", [120, 45, 0, 55], "doc_title", "required")
    add("2_Facility_name", "Наименование объекта", [90, 33, 0, 45], "facility_name", "optional")
    add("3_Unit_title_name", "Наименование установки", [100, 15, 55, 30], "unit_title", "optional")
    add("4_Document_name", "Наименование документа", [100, 0, 55, 15], "document_name", "optional")
    add("5_Documentation_type", "Вид документации", [45, 15, 35, 25], "documentation_type", "optional")
    add("6_1_Sheet_number", "Номер листа", [35, 20, 20, 25], "sheet_number_6_1", "required")
    add("6_2_Quantity_of_sheets", "Листов в документе", [35, 15, 20, 20], "sheet_number_6_2", "optional")
    add("7_Total_number_of_sheets", "Всего листов", [17, 15, 0, 25], "total_number_of_sheets", "optional")
    add("26_Document_Revision", "Текущая ревизия документа", [17, 55, 0, 70], "revision", "required")

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
        add(f"Position_{i + 1}", f"Должность {i + 1}", bb(c), None, "optional")
    for i, c in enumerate(col2):
        add(f"Surname_{i + 1}", f"Фамилия {i + 1}", bb(c), None, "optional")
    for i, c in enumerate(col5):
        add(f"10_{i + 1}_Signatures_Date", f"Дата подписи {i + 1}", bb(c), "signatures_date_10", "optional")

    def split18(
        prefix: str, id_suffix: str, label_prefix: str, x0: float, x1: float, clean_key: str
    ) -> None:
        yb, yt = 90.0, 111.0
        step = (yt - yb) / 3.0
        for r in range(3):
            y0 = yb + r * step
            y1 = yb + (r + 1) * step
            add(
                f"{prefix}_{r + 1}_{id_suffix}",
                f"{label_prefix} {r + 1}",
                [x0, round(y0, 2), x1, round(y1, 2)],
                clean_key,
                "optional",
            )

    split18("18_1", "Current_Revision", "Текущая ревизия (блок 18.1)", 185, 170, "current_revision_18_1")
    split18("18_2", "Revision_Date", "Дата ревизии (блок 18.2)", 170, 155, "revision_date_18_2")
    split18("18_3", "Purpose_of_issue", "Назначение выпуска (блок 18.3)", 145, 60, "purpose_of_issue_18_3")

    add("50_File_Name_Stamp", "Имя файла под штампом", [80, -50, 285, -1], "file_name_stamp", "required")
    add("51_Page_Format", "Формат страницы", [-50, -50, 80, -1], "page_format", "optional")

    doc = {
        "schema_version": 2,
        "name": "DWG стр.1 (заготовка v1→JSON; уточнить на этапе C.1)",
        "doc_types": ["WIR", "LAY", "CAE", "GA", "PL", "DW", "NI"],
        "page_selector": "first",
        "priority": 10,
        "origin": "frame_bottom_right",
        "padding_mm": 0.5,
        "fields": fields,
    }

    root = os.path.dirname(os.path.abspath(__file__))
    tpl = os.path.join(root, "templates")
    os.makedirs(tpl, exist_ok=True)
    path = os.path.join(tpl, "dwg_page1.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)
    print(path, "fields", len(fields))


if __name__ == "__main__":
    main()
