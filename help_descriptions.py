# -*- coding: utf-8 -*-
"""
Справка по кнопкам GUI: цепочки вызовов (модуль.функция), без дублирования
устаревших маркетинговых формулировок. Редактировать вручную при смене wiring.
"""

# Словарь с описаниями функций
HELP_DESCRIPTIONS = {
    "open_pdf_folder": {
        "title": "Центр управления PDF v2",
        "description": """
Цепочка вызовов

└─ subprocess.Popen([sys.executable, "-m", "pdf_v2_monitor"], cwd=<корень проекта>)
   └─ пакет pdf_v2_monitor (отдельный процесс; полный конвейер v2 в своём окне)

Настройки: pdf_v2_config.json в корне и UI монитора.
        """,
        "icon": "📄",
    },

    "mto_dir": {
        "title": "Проверка MTO + BBB (папка DWG)",
        "description": """
Отдельной кнопки «i» сейчас нет; та же логика, что у основной кнопки папки MTO в блоке «ПРОВЕРКА MTO».

Цепочка вызовов

└─ main.open_mto_dwg_with_optional_bbb
   ├─ base.bbb_config.load_config
   ├─ при включённой проверке MTO: base.google_sheets.start(..., out_dir=…, cfg / correction_cfg)
   └─ при включённой проверке BBB: base.bbb_analysis.start_bbb_analysis(dir_path, out_dir=…)
        """,
        "icon": "📊",
    },

    "mto_file": {
        "title": "Проверка MTO (файл)",
        "description": """
Цепочка вызовов

└─ main.open_file_att(MTO_FILE)  →  base.google_sheets.start(
       dir_path, GuiConst.MTO_FILE, file_path,
       cfg=mto_vs_code_base, correction_cfg=mto_vs_code_base_correction)
   └─ дальше внутри google_sheets.start — загрузка/сверка с code base (см. модуль)

Проверка может быть отключена в настройках (mto_vs_code_base.enabled).
        """,
        "icon": "📊",
    },

    "boot_file": {
        "title": "Проверка вывода NanoCAD (BOOT)",
        "description": """
Цепочка вызовов

└─ main.open_file_att(BOOT_FILE)  →  base.google_sheets.start(dir_path, GuiConst.BOOT_FILE, file_path)
        """,
        "icon": "🏗️",
    },

    "open_rfq_folder": {
        "title": "Проверка RFQ файла",
        "description": """
Цепочка вызовов

└─ main.open_rfq_folder  →  RFQ.RFQ_compare.rfq_file_start(file_path, dir_path)
   ├─ base.base_mto.get_std_from_excel_file (лист RFQ)
   ├─ base.base_google.load_base()  (MTO эталон из настроенных путей)
   ├─ base.base_cheks.base_vs_base_std(...)
   └─ base.base_excel_out.check_color_out(..., префикс "RFQ_vs_MTO", …)
        """,
        "icon": "📋",
    },

    "output_file": {
        "title": "Проверка Output файла",
        "description": """
Цепочка вызовов

└─ main.open_file_att(OUTPUT_FILE)  →  base.google_sheets.start(dir_path, GuiConst.OUTPUT_FILE, file_path)
        """,
        "icon": "⚡",
    },

    "button_open_cj_dir": {
        "title": "Кабельные журналы (папка)",
        "description": """
Цепочка вызовов

└─ main.button_open_cj_dir  →  cable_mapping.cab_mapping.cab_mapping_start(dir_path)
        """,
        "icon": "🔌",
    },

    "open_mto_mto_folder": {
        "title": "MTO vs MTO (2 шт.)",
        "description": """
Цепочка вызовов

└─ main.open_mto_mto_folder  →  RFQ.mto_compare.mto_compare_start(pdf_path)
   ├─ utils.path.get_files_single(..., .xlsx)
   ├─ base.base_mto.get_mto_std_from_file / get_std_from_excel_file (по имени файла)
   ├─ RFQ.Value_Compare.start(all_bases)  →  all_base_unique
   └─ RFQ.Value_Compare.two_mto_compare(all_base_unique, path_out_dir)
      └─ base.base_excel_out.check_color_out(...)
        """,
        "icon": "🔄",
    },

    "open_mto_mto_multi_folder": {
        "title": "MTO vs MTO (мульти)",
        "description": """
Цепочка вызовов

└─ main.open_mto_mto_multi_folder  →  RFQ.mto_compare.mto_multi_compare_start(pdf_path)
   ├─ utils.path.get_files_single(...)
   ├─ загрузка баз: MTO / RFQ / output / ZIP (ветвление по имени файла)
   ├─ RFQ.Value_Compare.start(all_bases, check_position_row=0)
   └─ RFQ.Value_Compare.multi_mto_compare(all_base_unique, path_out_dir)
      └─ base.base_excel_out.check_color_out(...)
        """,
        "icon": "🔄",
    },

    "open_mto_mto_chain_folder": {
        "title": "MTO: цепочка ревизий",
        "description": """
Цепочка вызовов

└─ main.open_mto_mto_chain_folder  →  RFQ.mto_compare.mto_chain_compare_start(pdf_path)
   ├─ группировка по титул–марка, сортировка по ревизии; пары соседей
   ├─ base.base_mto.get_mto_std_from_file + RFQ.func_get_unique_row_list.get_unique_row_list
   └─ для каждой пары: RFQ.Value_Compare.two_mto_compare([u_a, u_b], pair_out_dir, …)
      └─ base.base_excel_out.check_color_out(...)
        """,
        "icon": "🔄",
    },

    "open_ds_mto_folder": {
        "title": "Список ДС vs MTO (файл)",
        "description": """
Цепочка вызовов

└─ main.open_ds_mto_folder  →  RFQ.ds_compare.ds_start_init.analyze_ds_specification(ds_file, mto_path=…, summ_ds=True)

Папка МТО: ComboBox справа от кнопки (``mto_paths``, ``mto_path_selected_index`` в
``RFQ/ds_compare/ds_compare_config.json``). Узкая кнопка «⚙️» — список пресетов, добавление и правка.
        """,
        "icon": "📋",
    },

    "open_grouped_ds_mto_folder": {
        "title": "Grouped ДС vs MTO (файл)",
        "description": """
Цепочка вызовов

└─ main.open_grouped_ds_mto_folder  →  RFQ.ds_compare.ds_grouped_compare.analyze_grouped_ds_specification(...)
   ├─ загружает ДС
   ├─ группирует строки по ключам из настроек ДС vs MTO
   ├─ передаёт сгруппированные строки в analyze_ds_specification(...)
   └─ сохраняет отдельный Excel РОБОТ_СРАВНЕНИЕ_GROUPED_*_xw.xlsx

Настройки ключей группировки находятся в окне ⚙️ рядом с выбором папки МТО.
RFQ-колонки в этом MVP ещё не добавляются.
        """,
        "icon": "📋",
    },

    "open_merge_ds_folder": {
        "title": "Объединить ДС из папки",
        "description": """
Цепочка вызовов

└─ main.open_merge_ds_folder  →  RFQ.ds_compare.ds_merge_folder.merge_ds_folder(...)

Примечание: в текущем main.py путь к папке может быть зашит константой, а не диалогом выбора — см. код open_merge_ds_folder.
        """,
        "icon": "📋",
    },

    "open_mto_ds_file": {
        "title": "MTO vs список DS (файл)",
        "description": """
Цепочка вызовов

└─ main.open_mto_ds_file  →  RFQ.ds_compare.mto_ds_compare.mto_ds_list_compare(ds_file)
        """,
        "icon": "🔍",
    },

    "open_agregate_tags": {
        "title": "Проверка МТО с РКД и RFP",
        "description": """
Цепочка вызовов

└─ main.wire_agregate_tags_button → main._agregate_tags_subprocess_work (worker)
   └─ subprocess.Popen([sys.executable, "-u", "RFQ/tags_rfp_compare/agregate_tags.py"], env PYTHONPATH=корень)
      └─ RFQ.tags_rfp_compare.agregate_tags.main()  (точка входа скрипта)

Вторичная кнопка «📊»: main.open_latest_step4_result → RFQ.tags_rfp_compare.rfp_tags_utils.find_latest_step4_result_file + os.startfile
        """,
        "icon": "📋",
    },

    "open_agregate_tags_asbuild": {
        "title": "Проверка as-build",
        "description": """
Цепочка вызовов

└─ main.wire_agregate_tags_asbuild_button → main._agregate_tags_asbuild_subprocess_work
   ├─ при отсутствии файла: save_asbuild_config(load_asbuild_config()) под путь get_asbuild_config_path()
   └─ subprocess.Popen(..., "agregate_tags.py", <путь к конфигу as-build>)
      └─ RFQ.tags_rfp_compare.agregate_tags.main()

Вторичная кнопка «📊»: как у основной проверки — main.open_latest_step4_result.
        """,
        "icon": "🏗️",
    },

    "open_nanocad_db": {
        "title": "Открыть NanoCAD.db",
        "description": """
Цепочка вызовов

└─ main.open_nanocad_db  →  nano_cad.nc_start.load_db(file_path)
        """,
        "icon": "🗄️",
    },

    "open_project_release_zip": {
        "title": "Собрать ZIP для коллег",
        "description": """
Цепочка вызовов

└─ main.open_project_release_zip
   ├─ utils.release_zip.build_release_zip(project_root)  (правила: .release_zipignore)
   └─ utils.path.open_dir(каталог с архивом)
        """,
        "icon": "📦",
    },
}


def get_help_description(key):
    """Возвращает описание функции по ключу"""
    return HELP_DESCRIPTIONS.get(
        key,
        {
            "title": "Справка недоступна",
            "description": "Описание для этой функции не найдено.",
            "icon": "❓",
        },
    )


def get_all_keys():
    """Возвращает список всех доступных ключей"""
    return list(HELP_DESCRIPTIONS.keys())
