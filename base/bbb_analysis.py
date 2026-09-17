import utils.path
from base.base_google import load_base
from base.base_classes import RowStd, RowType
from base.bbb_load import load_all_from_dir
from base.bbb_checks import bbb_vs_code_base, bbb_vs_mto_combined
from base.bbb_excel_out import save_bbb_to_excel
from base.bbb_config import load_config
from base.bbb_section_types_excel import load_section_types


def start_bbb_analysis(dir_path: str, out_dir: str | None = None):
    print(f"=== BBB Analysis: {dir_path} ===")

    config = load_config()
    cb_cfg = config.get("bbb_vs_code_base", {})
    section_types = load_section_types()
    mto_cfg = {**config.get("bbb_vs_mto", {}), "section_types": section_types}

    loaded = load_all_from_dir(dir_path)
    if not loaded:
        print("Не удалось загрузить данные. Завершение.")
        return

    path_out_dir = out_dir or utils.path.get_path_out_dir(dir_path, dir_result_prefix="/__результат_BBB_")
    print(f"Результаты будут сохранены в: {path_out_dir}")

    need_code_base = cb_cfg.get("enabled", True)
    code_base_data_std = None
    if need_code_base:
        print("Загрузка code_base...")
        code_base_data_std = load_base()

    # 1) Подготовка и локальные проверки каждого документа.
    for doc_type in ("BOE", "BOM", "BOQ"):
        if doc_type not in loaded:
            continue

        data = loaded[doc_type]
        print(f"\n--- Проверка {doc_type} ({len(data)} строк) ---")

        if doc_type in ("BOE", "BOM"):
            if cb_cfg.get("enabled", True) and code_base_data_std is not None:
                bbb_vs_code_base(data, code_base_data_std, cfg=cb_cfg)
            else:
                RowStd.set_color_by_row_type_in_list(data)
        else:
            RowStd.set_color_by_row_type_in_list(data)

        loaded[doc_type] = data

    # 2) Объединённое сравнение BOE+BOM с MTO.
    if mto_cfg.get("enabled", True) and "MTO" in loaded and ("BOE" in loaded or "BOM" in loaded):
        print("\n--- Сравнение BOE+BOM с MTO ---")
        boe_data, bom_data = bbb_vs_mto_combined(
            loaded.get("BOE"),
            loaded.get("BOM"),
            loaded["MTO"],
            cfg=mto_cfg,
        )
        if "BOE" in loaded:
            loaded["BOE"] = boe_data
        if "BOM" in loaded:
            loaded["BOM"] = bom_data

    # 3) Экспорт каждого документа в свой Excel.
    for doc_type in ("BOE", "BOM", "BOQ"):
        if doc_type not in loaded:
            continue

        data = loaded[doc_type]
        file_prefix = f"check_{doc_type}_vs_base"
        save_bbb_to_excel(data, doc_type, path_out_dir, file_prefix)
        print(f"  -> {doc_type}: готово")

    print("\n=== BBB Analysis завершён ===")
    return path_out_dir


if __name__ == "__main__":
    start_bbb_analysis(r"C:\YandexDisk\темп\BBB\DWG")
