"""
Окно объединённых настроек MTO + BBB.
Позволяет управлять запуском MTO/BBB и детальными проверками MTO/BBB.
"""

import os
import subprocess
import customtkinter
from typing import Any, Callable, Dict, Optional

from base.bbb_config import load_config, save_config, get_default_config


def show_bbb_settings(parent=None, on_save: Optional[Callable[[Dict], None]] = None) -> None:
    config = load_config()

    win = customtkinter.CTkToplevel(parent)
    win.title("Настройки: Проверка MTO + BBB")
    win.geometry("640x760")
    win.resizable(True, True)

    main_frame = customtkinter.CTkScrollableFrame(win)
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)

    vars_: Dict[str, customtkinter.BooleanVar] = {}

    def add_section(title: str, row: int) -> int:
        lbl = customtkinter.CTkLabel(main_frame, text=title, font=("Arial", 14, "bold"))
        lbl.grid(row=row, column=0, columnspan=2, sticky="w", pady=(15, 5))
        return row + 1

    def add_check(key: str, text: str, row: int, value: bool, indent: int = 20) -> int:
        v = customtkinter.BooleanVar(value=value)
        cb = customtkinter.CTkCheckBox(main_frame, text=text, variable=v)
        cb.grid(row=row, column=0, columnspan=2, sticky="w", padx=(indent, 0), pady=2)
        vars_[key] = v
        return row + 1

    row = 0

    # --- folder_rules ---
    folder_rules_cfg = config.get("folder_rules", {})
    row = add_section("Ограничения выбора папки", row)
    row = add_check(
        "folder_rules.search_only_in_dwg",
        "Искать только в папке DWG",
        row,
        folder_rules_cfg.get("search_only_in_dwg", True),
        indent=10,
    )

    # --- mto_vs_code_base ---
    mto_code_cfg = config.get("mto_vs_code_base", {})
    row = add_section("Проверка MTO по code_base (mto_vs_code_base)", row)
    row = add_check("mto_cb.enabled", "Включить mto_vs_code_base", row, mto_code_cfg.get("enabled", True), indent=10)
    row = add_check("mto_cb.check_by_code", "check_by_code — сравнение NAME/TYPE_MARK/UNITS/VENDOR/CODE", row, mto_code_cfg.get("check_by_code", True))
    row = add_check("mto_cb.check_equipment_codes", "check_equipment_codes — проверка кодов оборудования", row, mto_code_cfg.get("check_equipment_codes", True))
    row = add_check("mto_cb.check_tags_value", "check_tags_value — проверка тегов и значений", row, mto_code_cfg.get("check_tags_value", True))
    row = add_check("mto_cb.check_value", "check_value — проверка количеств (VALUES)", row, mto_code_cfg.get("check_value", True))
    row = add_check("mto_cb.check_duplicate_tags", "check_duplicate_tags — проверка дублей тегов", row, mto_code_cfg.get("check_duplicate_tags", True))
    row = add_check("mto_cb.check_mass", "check_mass — проверка массы", row, mto_code_cfg.get("check_mass", True))
    row = add_check("mto_cb.check_prohibition", "check_prohibition — проверка запрещённых кодов", row, mto_code_cfg.get("check_prohibition", False))
    row = add_check("mto_cb.check_tags_4_2_4_4", "check_tags_4_2_4_4 — проверка 94S 4.2 / 4.4", row, mto_code_cfg.get("check_tags_4_2_4_4", True))
    row = add_check("mto_cb.check_position_numeration", "check_position_numeration — проверка нумерации позиций", row, mto_code_cfg.get("check_position_numeration", True))

    # --- mto_vs_code_base_correction ---
    mto_corr_cfg = config.get("mto_vs_code_base_correction", {})
    row = add_section("Коррекция MTO (mto_vs_code_base_correction)", row)
    row = add_check("mto_corr.enabled", "Включить mto_vs_code_base_correction", row, mto_corr_cfg.get("enabled", True), indent=10)
    row = add_check("mto_corr.check_by_code", "check_by_code — коррекция по code_base", row, mto_corr_cfg.get("check_by_code", True))
    row = add_check("mto_corr.check_tags_value", "check_tags_value — проверка тегов и значений", row, mto_corr_cfg.get("check_tags_value", True))
    row = add_check("mto_corr.check_value", "check_value — проверка количеств (VALUES)", row, mto_corr_cfg.get("check_value", True))
    row = add_check("mto_corr.check_duplicate_tags", "check_duplicate_tags — проверка дублей тегов", row, mto_corr_cfg.get("check_duplicate_tags", True))
    row = add_check("mto_corr.check_tags_4_2_4_4", "check_tags_4_2_4_4 — проверка 94S 4.2 / 4.4", row, mto_corr_cfg.get("check_tags_4_2_4_4", True))
    row = add_check("mto_corr.check_position_numeration", "check_position_numeration — проверка нумерации позиций", row, mto_corr_cfg.get("check_position_numeration", True))

    # --- bbb_vs_code_base ---
    cb_cfg = config.get("bbb_vs_code_base", {})
    row = add_section("Проверка по code_base (bbb_vs_code_base)", row)
    row = add_check("cb.enabled", "Включить bbb_vs_code_base", row, cb_cfg.get("enabled", True), indent=10)
    row = add_check("cb.check_by_code", "check_by_code — сравнение NAME/TYPE_MARK/UNITS/VENDOR/CODE", row, cb_cfg.get("check_by_code", True))
    row = add_check("cb.check_equipment_codes", "check_equipment_codes — проверка кодов оборудования", row, cb_cfg.get("check_equipment_codes", True))
    row = add_check("cb.check_tags_value", "check_tags_value — проверка тегов и значений", row, cb_cfg.get("check_tags_value", True))
    row = add_check("cb.check_value", "check_value — проверка количеств (VALUES)", row, cb_cfg.get("check_value", True))
    row = add_check("cb.check_duplicate_tags", "check_duplicate_tags — проверка дублей тегов", row, cb_cfg.get("check_duplicate_tags", True))
    row = add_check("cb.check_mass", "check_mass — проверка массы", row, cb_cfg.get("check_mass", True))
    row = add_check("cb.check_prohibition", "check_prohibition — проверка запрещённых кодов", row, cb_cfg.get("check_prohibition", False))
    row = add_check(
        "cb.check_work_code_and_mtr_group",
        "check_work_code_and_mtr_group — проверка BBB_WORK_CODE и BBB_MTR_GROUP по GoogleBase",
        row,
        cb_cfg.get("check_work_code_and_mtr_group", True),
    )

    # --- bbb_vs_mto ---
    mto_cfg = config.get("bbb_vs_mto", {})

    # Заголовок секции + кнопка открытия Excel файла типов секций
    lbl_mto = customtkinter.CTkLabel(
        main_frame, text="Сравнение с MTO (bbb_vs_mto)", font=("Arial", 14, "bold")
    )
    lbl_mto.grid(row=row, column=0, sticky="w", pady=(15, 5))

    def _open_section_types_excel():
        from base.bbb_section_types_excel import get_excel_path
        excel_path = get_excel_path()
        try:
            os.startfile(excel_path)
        except FileNotFoundError:
            customtkinter.CTkLabel(
                main_frame,
                text=f"Файл не найден:\n{excel_path}",
                text_color="red",
                wraplength=400,
            ).grid(row=row, column=0, columnspan=2, sticky="w", padx=(20, 0))
        except Exception as e:
            print(f"[WARN] Не удалось открыть файл типов секций: {e}")

    btn_sections = customtkinter.CTkButton(
        main_frame,
        text="Открыть типы секций MTO",
        command=_open_section_types_excel,
        width=200,
        height=26,
    )
    btn_sections.grid(row=row, column=1, sticky="e", padx=(5, 0), pady=(15, 5))
    row += 1
    row = add_check("mto.enabled", "Включить bbb_vs_mto", row, mto_cfg.get("enabled", True), indent=10)
    row = add_check("mto.compare_fields", "compare_fields — сравнение NAME/TYPE_MARK/VENDOR/UNITS с MTO", row, mto_cfg.get("compare_fields", True))
    row = add_check("mto.compare_values", "compare_values — сравнение сумм VALUES BBB↔MTO", row, mto_cfg.get("compare_values", True))
    row = add_check("mto.compare_tags", "compare_tags — сравнение тегов BOE↔MTO", row, mto_cfg.get("compare_tags", True))
    row = add_check(
        "mto.compare_title_marka_revision",
        "compare_title_marka_revision — сравнение TITLE/BBB_MARKA/BBB_REVISION с именем файла MTO",
        row,
        mto_cfg.get("compare_title_marka_revision", True),
    )
    row = add_check(
        "mto.compare_bom_annotation_with_mto",
        "compare_bom_annotation_with_mto — проверка BOM.ANNOTATION по сумме слагаемых MTO",
        row,
        mto_cfg.get("compare_bom_annotation_with_mto", True),
    )
    row = add_check(
        "mto.check_missing_mto_codes",
        "check_missing_mto_codes — добавлять строки для кодов MTO, отсутствующих в BOE и BOM",
        row,
        mto_cfg.get("check_missing_mto_codes", True),
    )

    main_frame.grid_columnconfigure(0, weight=1)
    main_frame.grid_columnconfigure(1, weight=0)

    def collect_config() -> Dict[str, Any]:
        out = get_default_config()
        for k, sec, field in [
            ("folder_rules.search_only_in_dwg", "folder_rules", "search_only_in_dwg"),
            ("mto_cb.enabled",              "mto_vs_code_base", "enabled"),
            ("mto_cb.check_by_code",        "mto_vs_code_base", "check_by_code"),
            ("mto_cb.check_equipment_codes","mto_vs_code_base", "check_equipment_codes"),
            ("mto_cb.check_tags_value",     "mto_vs_code_base", "check_tags_value"),
            ("mto_cb.check_value",          "mto_vs_code_base", "check_value"),
            ("mto_cb.check_duplicate_tags", "mto_vs_code_base", "check_duplicate_tags"),
            ("mto_cb.check_mass",           "mto_vs_code_base", "check_mass"),
            ("mto_cb.check_prohibition",    "mto_vs_code_base", "check_prohibition"),
            ("mto_cb.check_tags_4_2_4_4",   "mto_vs_code_base", "check_tags_4_2_4_4"),
            ("mto_cb.check_position_numeration", "mto_vs_code_base", "check_position_numeration"),
            ("mto_corr.enabled",            "mto_vs_code_base_correction", "enabled"),
            ("mto_corr.check_by_code",      "mto_vs_code_base_correction", "check_by_code"),
            ("mto_corr.check_tags_value",   "mto_vs_code_base_correction", "check_tags_value"),
            ("mto_corr.check_value",        "mto_vs_code_base_correction", "check_value"),
            ("mto_corr.check_duplicate_tags","mto_vs_code_base_correction", "check_duplicate_tags"),
            ("mto_corr.check_tags_4_2_4_4", "mto_vs_code_base_correction", "check_tags_4_2_4_4"),
            ("mto_corr.check_position_numeration", "mto_vs_code_base_correction", "check_position_numeration"),
            ("cb.enabled",              "bbb_vs_code_base", "enabled"),
            ("cb.check_by_code",        "bbb_vs_code_base", "check_by_code"),
            ("cb.check_equipment_codes", "bbb_vs_code_base", "check_equipment_codes"),
            ("cb.check_tags_value",     "bbb_vs_code_base", "check_tags_value"),
            ("cb.check_value",          "bbb_vs_code_base", "check_value"),
            ("cb.check_duplicate_tags", "bbb_vs_code_base", "check_duplicate_tags"),
            ("cb.check_mass",           "bbb_vs_code_base", "check_mass"),
            ("cb.check_prohibition",    "bbb_vs_code_base", "check_prohibition"),
            ("cb.check_work_code_and_mtr_group", "bbb_vs_code_base", "check_work_code_and_mtr_group"),
            ("mto.enabled",             "bbb_vs_mto",       "enabled"),
            ("mto.compare_fields",      "bbb_vs_mto",       "compare_fields"),
            ("mto.compare_values",      "bbb_vs_mto",       "compare_values"),
            ("mto.compare_tags",        "bbb_vs_mto",       "compare_tags"),
            ("mto.compare_title_marka_revision", "bbb_vs_mto", "compare_title_marka_revision"),
            ("mto.compare_bom_annotation_with_mto", "bbb_vs_mto", "compare_bom_annotation_with_mto"),
            ("mto.check_missing_mto_codes",          "bbb_vs_mto", "check_missing_mto_codes"),
        ]:
            if k in vars_:
                out[sec][field] = vars_[k].get()
        return out

    def do_save():
        cfg = collect_config()
        if save_config(cfg):
            if on_save:
                on_save(cfg)
            win.destroy()

    btn_frame = customtkinter.CTkFrame(win, fg_color="transparent")
    btn_frame.pack(fill="x", padx=20, pady=(0, 20))
    customtkinter.CTkButton(btn_frame, text="Сохранить", command=do_save, width=120).pack(side="left", padx=(0, 10))
    customtkinter.CTkButton(btn_frame, text="Отмена", command=win.destroy, width=100).pack(side="left")

    win.transient(parent)
    win.grab_set()
