"""
Окно настроек для RFP tags compare
Параметры step1, step2, step3, step4 и rfp_tags_utils
"""

from __future__ import annotations

import copy
import tkinter
import customtkinter
from typing import Dict, Any, Callable, Optional

from RFQ.tags_rfp_compare.rfp_tags_utils import (
    DEFAULT_UNITS_SPLIT_BAN,
    load_config,
    save_config,
    resolve_units_split_ban_config,
    resolve_rfp_pipeline_mode,
    RFP_PIPELINE_MODE_MTO_VO_ONLY,
    RFP_PIPELINE_MODE_STANDARD,
    RFP_PIPELINE_MODE_WITH_ORPHAN_MTO_VO,
)


_RFP_PIPELINE_MENU: tuple[tuple[str, str], ...] = (
    (
        RFP_PIPELINE_MODE_STANDARD,
        "Стандарт: RFP из файла; MTO/VO без якоря — чекбокс Step4 ниже",
    ),
    (
        RFP_PIPELINE_MODE_WITH_ORPHAN_MTO_VO,
        "RFP + добавлять MTO/VO без секции RFP (вкл. автоматически)",
    ),
    (
        RFP_PIPELINE_MODE_MTO_VO_ONLY,
        "Только MTO/VO (файл RFP не читается)",
    ),
)
_RFP_PIPELINE_LABEL_TO_MODE = {lbl: key for key, lbl in _RFP_PIPELINE_MENU}
_RFP_PIPELINE_MODE_TO_LABEL = {key: lbl for key, lbl in _RFP_PIPELINE_MENU}


def _bind_paste_to_entry(entry: customtkinter.CTkEntry) -> None:
    """
    Привязывает вставку из буфера к Entry.
    Исправляет проблему: Ctrl+V не работает в CTkEntry при русской раскладке (и др.).
    CTkEntry — это Frame, события клавиш идут во внутренний _entry.
    """
    inner = getattr(entry, "_entry", entry)

    def do_paste(event=None):
        try:
            root = entry.winfo_toplevel()
            text = root.clipboard_get()
        except Exception:
            return
        if text:
            try:
                inner.delete("sel.first", "sel.last")
            except Exception:
                pass
            inner.insert("insert", text)
        if event:
            return "break"

    def on_ctrl_key(event):
        # keycode 86 = физическая клавиша V (одинакова для англ. и рус. раскладки)
        if event.keycode == 86 and event.state & 0x4:  # 0x4 = Control
            return do_paste(event)

    inner.bind("<Control-KeyPress>", on_ctrl_key)

    # Контекстное меню по правому клику
    ctx_menu = tkinter.Menu(entry, tearoff=0)
    ctx_menu.add_command(label="Вставить", command=do_paste)

    def show_context_menu(event):
        try:
            ctx_menu.tk_popup(event.x_root, event.y_root)
        finally:
            ctx_menu.grab_release()

    inner.bind("<Button-3>", show_context_menu)


def _parse_list_str(s: str) -> list:
    """Парсит строку в список (через запятую, пробелы убираются)"""
    if not s or not str(s).strip():
        return []
    return [x.strip() for x in str(s).split(",") if x.strip()]


def show_rfp_tags_settings(
    parent=None,
    on_save: Optional[Callable[[Dict], None]] = None,
    load_fn: Optional[Callable[[], Dict]] = None,
    save_fn: Optional[Callable[[Dict], bool]] = None,
    title_suffix: str = "",
) -> None:
    """
    Открывает окно настроек RFP tags compare.
    При сохранении вызывает on_save(config) если передан.
    load_fn / save_fn — переопределяют функции загрузки/сохранения конфига (для as-build и др.).
    title_suffix — добавляется к заголовку окна (например ' (as-build)').
    """
    _load = load_fn if load_fn is not None else load_config
    _save = save_fn if save_fn is not None else save_config
    config = _load()

    win = customtkinter.CTkToplevel(parent)
    win.title(f"Настройки: Проверка МТО с РКД и RFP{title_suffix}")
    win.geometry("1015x1000")
    win.resizable(True, True)

    main_frame = customtkinter.CTkScrollableFrame(win)
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)

    # Переменные для виджетов
    vars_: Dict[str, Any] = {}

    def add_section(title: str, row_start: int) -> int:
        lbl = customtkinter.CTkLabel(main_frame, text=title, font=("Arial", 14, "bold"))
        lbl.grid(row=row_start, column=0, columnspan=2, sticky="w", pady=(15, 5))
        return row_start + 1

    def add_check(master, key: str, text: str, row: int, value: bool):
        v = customtkinter.BooleanVar(value=value)
        cb = customtkinter.CTkCheckBox(master, text=text, variable=v)
        cb.grid(row=row, column=0, columnspan=2, sticky="w", padx=(20, 0), pady=2)
        vars_[key] = v
        return row + 1

    def add_entry(master, key: str, label: str, row: int, value: str, width: int = 60):
        customtkinter.CTkLabel(master, text=label, anchor="w").grid(row=row, column=0, sticky="w", padx=(20, 5), pady=2)
        v = customtkinter.StringVar(value=str(value) if value is not None else "")
        e = customtkinter.CTkEntry(master, textvariable=v, width=width)
        e.grid(row=row, column=1, sticky="ew", pady=2)
        _bind_paste_to_entry(e)
        vars_[key] = v
        return row + 1

    def add_spin(master, key: str, label: str, row: int, value: int, from_: int = 1, to_: int = 50):
        customtkinter.CTkLabel(master, text=label, anchor="w").grid(row=row, column=0, sticky="w", padx=(20, 5), pady=2)
        v = customtkinter.StringVar(value=str(value))
        sp = customtkinter.CTkEntry(master, textvariable=v, width=80)
        sp.grid(row=row, column=1, sticky="w", pady=2)
        _bind_paste_to_entry(sp)
        vars_[key] = v
        return row + 1

    row = 0

    # === ПРОФИЛИРОВАНИЕ ПАМЯТИ ===
    # === ОПТИМИЗАЦИЯ СТОЛБЦОВ ===
    row = add_section("📐 Оптимизация столбцов (обрезка, load/skip)", row)
    col_opt = config.get("column_optimization", {})
    rfp_col = col_opt.get("rfp", {})
    mto_col = col_opt.get("mto", {})
    row = add_spin(main_frame, "column_optimization.rfp.name.truncate", "RFP name truncate:", row, rfp_col.get("name", {}).get("truncate", 200), 0, 1000)
    row = add_spin(main_frame, "column_optimization.rfp.type_mark.truncate", "RFP type_mark truncate:", row, rfp_col.get("type_mark", {}).get("truncate", 100), 0, 500)
    row = add_check(main_frame, "column_optimization.rfp.DS_SPECIFICATION.load", "RFP DS_SPECIFICATION load (снять = обнулить)", row, rfp_col.get("DS_SPECIFICATION", {}).get("load", False))
    row = add_spin(main_frame, "column_optimization.mto.name.truncate", "MTO name truncate:", row, mto_col.get("name", {}).get("truncate", 200), 0, 1000)
    row = add_spin(main_frame, "column_optimization.mto.type_mark.truncate", "MTO type_mark truncate:", row, mto_col.get("type_mark", {}).get("truncate", 100), 0, 500)

    # === ПРОФИЛИРОВАНИЕ ПАМЯТИ ===
    row = add_section("📊 Профилирование памяти", row)
    row = add_check(main_frame, "memory_log", "memory_log (memory_log.txt на всех этапах)", row, config.get("memory_log", True))
    s4 = config.get("step4", {})
    row = add_check(main_frame, "step4.memory_top_stats", "memory_top_stats (топ аллокаций tracemalloc, медленно)", row, s4.get("memory_top_stats", False))
    row = add_spin(main_frame, "step4.memory_top_n", "memory_top_n:", row, s4.get("memory_top_n", 15), 1, 50)

    # === ПУТИ ===
    row = add_section("📁 Пути к файлам и папкам", row)
    paths = config.get("paths", {})
    for pkey, plabel in [
        ("rfp_path", "RFP файл:"),
        ("rfp_registr_lot_path", "Реестр лотов:"),
        ("mto_path", "Папка MTO:"),
        ("vo_path", "Папка VO:"),
        ("code_ban_file", "Файл code_ban:"),
        ("replacement_table_file", "Таблица замен кодов:"),
        ("result_dir_base", "Базовая папка результатов:"),
    ]:
        row = add_entry(main_frame, f"paths.{pkey}", plabel, row, paths.get(pkey, ""), width=55)

    # === STEP1 ===
    row = add_section("📄 Step1: Загрузка RFP", row)
    s1 = config.get("step1", {})
    row = add_check(main_frame, "step1.debug", "debug", row, s1.get("debug", False))
    row = add_check(
        main_frame,
        "load_tags",
        "load_tags: читать теги RFP/MTO/VO/УЛ (выкл. = свод rfp_parts_net_no_tags.xlsx, сопоставление по коду, без раскладки VALUES=1 на несколько тегов)",
        row,
        config.get("load_tags", True),
    )
    row = add_check(main_frame, "step1.skip_split", "skip_split (split RFP в worker, для parallel_by_title_mark)", row, s1.get("skip_split", False))
    row = add_check(
        main_frame,
        "step1.use_rfp_code_ban",
        "use_rfp_code_ban: читать/обновлять code_ban.txt (выкл. = полное раскрытие RFP по кодам, файл не трогается)",
        row,
        s1.get("use_rfp_code_ban", True),
    )

    _cur_mode = resolve_rfp_pipeline_mode(s1)
    _init_label = _RFP_PIPELINE_MODE_TO_LABEL.get(_cur_mode, _RFP_PIPELINE_MENU[0][1])
    customtkinter.CTkLabel(
        main_frame,
        text="Режим участия RFP (step1.rfp_pipeline_mode):",
        anchor="w",
    ).grid(row=row, column=0, sticky="w", padx=(20, 5), pady=2)
    _pm_var = tkinter.StringVar(value=_init_label)
    _pm_menu = customtkinter.CTkOptionMenu(
        main_frame,
        variable=_pm_var,
        values=[lbl for _, lbl in _RFP_PIPELINE_MENU],
        width=520,
    )
    _pm_menu.grid(row=row, column=1, sticky="ew", pady=2)
    vars_["step1.rfp_pipeline_mode_display"] = _pm_var
    row += 1

    row = add_section("📏 Раскладка RFP/MTO по количеству (ед. изм.)", row)
    usb_cfg = config.get("units_split_ban") or {}
    row = add_check(
        main_frame,
        "units_split_ban.enabled",
        "Бан единиц: не дробить в строки «по 1» позиции с указанными ед. изм. (выкл. = раскладывать всех, кого позволяет логика)",
        row,
        usb_cfg.get("enabled", True),
    )
    _usb_units = usb_cfg.get("units") if isinstance(usb_cfg.get("units"), list) else None
    if not _usb_units:
        _usb_units_str = ", ".join(DEFAULT_UNITS_SPLIT_BAN)
    else:
        _usb_units_str = ", ".join(str(x) for x in _usb_units if str(x).strip())
    row = add_entry(
        main_frame,
        "units_split_ban.units",
        "Единицы в бане (через запятую), при включённом пункте выше:",
        row,
        _usb_units_str,
        width=55,
    )

    def _effective_units_summary() -> str:
        en = bool(vars_["units_split_ban.enabled"].get())
        parsed = _parse_list_str(vars_["units_split_ban.units"].get())
        mini_cfg = {"units_split_ban": {"enabled": en, "units": parsed}}
        eff = resolve_units_split_ban_config(mini_cfg)
        prefix = "Фактически для split (resolve_units_split_ban_config): "
        if not eff:
            return prefix + "нет бана единиц"
        if en and not parsed:
            return prefix + "дефолтный список — " + ", ".join(eff)
        return prefix + ", ".join(eff)

    def _refresh_units_effective_label(*_args) -> None:
        try:
            units_effective_lbl.configure(text=_effective_units_summary())
        except Exception:
            pass

    units_effective_lbl = customtkinter.CTkLabel(
        main_frame,
        text="",
        anchor="w",
        justify="left",
        wraplength=920,
    )
    units_effective_lbl.grid(row=row, column=0, columnspan=2, sticky="ew", padx=(40, 10), pady=(2, 8))
    row += 1
    _refresh_units_effective_label()
    try:
        vars_["units_split_ban.enabled"].trace_add("write", _refresh_units_effective_label)
    except Exception:
        vars_["units_split_ban.enabled"].trace("w", _refresh_units_effective_label)
    try:
        vars_["units_split_ban.units"].trace_add("write", _refresh_units_effective_label)
    except Exception:
        vars_["units_split_ban.units"].trace("w", _refresh_units_effective_label)

    # === STEP2 ===
    row = add_section("📊 Step2: Загрузка MTO", row)
    s2 = config.get("step2", {})
    row = add_check(main_frame, "step2.debug", "debug", row, s2.get("debug", False))
    row = add_check(
        main_frame,
        "step2.export_load_results_excel",
        "export_load_results_excel (Шаг2_MTO_результаты_загрузки_*.xlsx)",
        row,
        s2.get("export_load_results_excel", True),
    )
    row = add_check(
        main_frame,
        "step2.export_positions_database_excel",
        "export_positions_database_excel (Шаг2_MTO_база_position_row_*.xlsx)",
        row,
        s2.get("export_positions_database_excel", False),
    )
    row = add_check(
        main_frame,
        "step2.flat_mto_structure",
        "flat_mto_structure (файлы MTO лежат прямо в корне папки, без подпапок)",
        row,
        s2.get("flat_mto_structure", False),
    )

    # === STEP3 ===
    row = add_section("📋 Step3: Загрузка VO", row)
    s3 = config.get("step3", {})
    row = add_check(main_frame, "step3.debug", "debug", row, s3.get("debug", False))
    row = add_check(main_frame, "step3.export_to_excel", "export_to_excel (выгрузка отчётов VO в Excel)", row, s3.get("export_to_excel", False))

    # === STEP4 ===
    row = add_section("🔍 Step4: Анализ и сопоставление", row)
    s4 = config.get("step4", {})
    row = add_check(main_frame, "step4.debug", "debug (общая отладка этапа, fallback)", row, s4.get("debug", True))
    row = add_check(main_frame, "step4.debug_step4_1", "debug_step4_1 (валидация MTO/VO, split, дубли тегов)", row, s4.get("debug_step4_1", True))
    row = add_check(main_frame, "step4.debug_step4_2", "debug_step4_2 (основной match RFP↔MTO/VO)", row, s4.get("debug_step4_2", True))
    row = add_check(main_frame, "step4.debug_step4_3", "debug_step4_3 (назначение POSITION_STATUS)", row, s4.get("debug_step4_3", True))
    row = add_check(main_frame, "step4.debug_step4_4", "debug_step4_4 (добавление несопоставленных строк и доп. match)", row, s4.get("debug_step4_4", True))
    row = add_check(main_frame, "step4.verbose_progress_messages", "verbose_progress_messages (подробные служебные сообщения 4_3/4_4 в консоли)", row, s4.get("verbose_progress_messages", True))
    row = add_check(main_frame, "step4.collapse_debug", "collapse_debug (детальный лог схлопывания строк в txt)", row, s4.get("collapse_debug", False))
    row = add_check(main_frame, "step4.unified_debug", "unified_debug (единый debug-лог step4_unified_debug_*.txt)", row, s4.get("unified_debug", True))
    row = add_check(main_frame, "step4.timing_log", "timing_log (тайминги этапов в timing_log.xlsx)", row, s4.get("timing_log", True))
    row = add_check(main_frame, "step4.use_multiprocessing", "use_multiprocessing (доп. параллелизм внутри шагов match_*)", row, s4.get("use_multiprocessing", False))
    row = add_check(main_frame, "step4.parallel_by_title_mark", "parallel_by_title_mark (обработка DS/title_mark в нескольких процессах)", row, s4.get("parallel_by_title_mark", False))
    row = add_spin(main_frame, "step4.max_workers", "max_workers (0=авто: cpu_count-1):", row, s4.get("max_workers", 4), 0, 64)
    row = add_check(main_frame, "step4.print_title_systems_table", "print_title_systems_table (печать таблицы RFP/MTO/VO в консоль)", row, s4.get("print_title_systems_table", False))
    row = add_check(
        main_frame,
        "step4.export_title_systems_comparison_excel",
        "Шаг4_RFP_MTO_VO_сравнение_title_system_*.xlsx (Excel-отчёт сравнения title_system)",
        row,
        s4.get("export_title_systems_comparison_excel", True),
    )
    row = add_check(
        main_frame,
        "step4.export_bcc_accum_matrix",
        "Шаг4_Матрица_BCC_*.xlsx (накопительная матрица по закупочному коду BCC)",
        row,
        s4.get("export_bcc_accum_matrix", True),
    )
    row = add_check(
        main_frame,
        "step4.assign_rfp_mto_code_compare_colors",
        "assign_rfp_mto_code_compare_colors (раскраска CODE RFP vs CODE MTO)",
        row,
        s4.get("assign_rfp_mto_code_compare_colors", True),
    )
    row = add_check(
        main_frame,
        "step4.filter_to_mto_titles",
        "filter_to_mto_titles (фильтр выхода: только title_system из MTO папки)",
        row,
        s4.get("filter_to_mto_titles", False),
    )
    row = add_check(
        main_frame,
        "step4.include_packing_lists",
        "include_packing_lists (сравнивать итоговый RFP с упаковочными листами)",
        row,
        s4.get("include_packing_lists", True),
    )
    row = add_check(
        main_frame,
        "step4.ul_match_use_mto_tags",
        "ul_match_use_mto_tags (учитывать теги МТО при посадке УЛ)",
        row,
        s4.get("ul_match_use_mto_tags", True),
    )
    v_include_anchor = tkinter.BooleanVar(value=s4.get("include_mto_vo_without_rfp_anchor", False))
    cb_include_anchor = customtkinter.CTkCheckBox(
        main_frame,
        text=(
            "include_mto_vo_without_rfp_anchor (добавлять MTO/VO без секции RFP с тем же title_system; "
            "в не-стандартном режиме Step1 — принудительно вкл.)"
        ),
        variable=v_include_anchor,
    )
    cb_include_anchor.grid(row=row, column=0, columnspan=2, sticky="w", padx=(20, 0), pady=2)
    vars_["step4.include_mto_vo_without_rfp_anchor"] = v_include_anchor
    row += 1

    def _sync_include_anchor_from_pipeline(_choice: Optional[str] = None) -> None:
        lbl = vars_["step1.rfp_pipeline_mode_display"].get()
        key = _RFP_PIPELINE_LABEL_TO_MODE.get(lbl, RFP_PIPELINE_MODE_STANDARD)
        if key == RFP_PIPELINE_MODE_STANDARD:
            cb_include_anchor.configure(state="normal")
        else:
            cb_include_anchor.configure(state="disabled")
            v_include_anchor.set(True)

    _pm_menu.configure(command=lambda v: _sync_include_anchor_from_pipeline(v))
    _sync_include_anchor_from_pipeline()

    # debug_tag, debug_code, debug_title_system — списки, храним как строки через запятую
    debug_tag_val = s4.get("debug_tag", "")
    debug_code_val = s4.get("debug_code", "BCC0003424")
    debug_ts_val = s4.get("debug_title_system", "8445-SOT1")
    if isinstance(debug_tag_val, list):
        debug_tag_val = ", ".join(debug_tag_val)
    if isinstance(debug_code_val, list):
        debug_code_val = ", ".join(debug_code_val)
    if isinstance(debug_ts_val, list):
        debug_ts_val = ", ".join(debug_ts_val)
    row = add_entry(main_frame, "step4.debug_tag", "debug_tag (теги, через запятую):", row, debug_tag_val, width=45)
    row = add_entry(main_frame, "step4.debug_code", "debug_code (коды, через запятую):", row, debug_code_val, width=45)
    row = add_entry(main_frame, "step4.debug_title_system", "debug_title_system (через запятую):", row, debug_ts_val, width=45)

    # === rfp_tags_utils ===
    row = add_section("⏱ rfp_tags_utils", row)
    rtu = config.get("rfp_tags_utils", {})
    row = add_spin(main_frame, "rfp_tags_utils.finalize_timing_top_n", "finalize_timing_top_n:", row, rtu.get("finalize_timing_top_n", 5), 1, 50)
    row = add_check(
        main_frame,
        "rfp_tags_utils.save_input_fingerprints",
        "save_input_fingerprints (сохранять input_fingerprints.json для MTO/VO)",
        row,
        rtu.get("save_input_fingerprints", True),
    )

    main_frame.grid_columnconfigure(1, weight=1)

    def collect_config() -> Dict[str, Any]:
        # Patch the loaded config so legacy UI saves do not discard unknown
        # sections or unexposed column_optimization options.
        out = copy.deepcopy(config)
        # column_optimization
        if "column_optimization.rfp.name.truncate" in vars_:
            try:
                n = max(0, min(1000, int(vars_["column_optimization.rfp.name.truncate"].get())))
                out.setdefault("column_optimization", {}).setdefault("rfp", {}).setdefault("name", {})["truncate"] = n
                out["column_optimization"]["rfp"]["name"]["load"] = True
            except (ValueError, TypeError):
                pass
        if "column_optimization.rfp.type_mark.truncate" in vars_:
            try:
                n = max(0, min(500, int(vars_["column_optimization.rfp.type_mark.truncate"].get())))
                out.setdefault("column_optimization", {}).setdefault("rfp", {}).setdefault("type_mark", {})["truncate"] = n
                out["column_optimization"]["rfp"]["type_mark"]["load"] = True
            except (ValueError, TypeError):
                pass
        if "column_optimization.rfp.DS_SPECIFICATION.load" in vars_:
            out.setdefault("column_optimization", {}).setdefault("rfp", {}).setdefault("DS_SPECIFICATION", {})["load"] = vars_["column_optimization.rfp.DS_SPECIFICATION.load"].get()
        if "column_optimization.mto.name.truncate" in vars_:
            try:
                n = max(0, min(1000, int(vars_["column_optimization.mto.name.truncate"].get())))
                out.setdefault("column_optimization", {}).setdefault("mto", {}).setdefault("name", {})["truncate"] = n
                out["column_optimization"]["mto"]["name"]["load"] = True
            except (ValueError, TypeError):
                pass
        if "column_optimization.mto.type_mark.truncate" in vars_:
            try:
                n = max(0, min(500, int(vars_["column_optimization.mto.type_mark.truncate"].get())))
                out.setdefault("column_optimization", {}).setdefault("mto", {}).setdefault("type_mark", {})["truncate"] = n
                out["column_optimization"]["mto"]["type_mark"]["load"] = True
            except (ValueError, TypeError):
                pass
        # memory_log (root)
        if "memory_log" in vars_:
            out["memory_log"] = vars_["memory_log"].get()
        if "load_tags" in vars_:
            out["load_tags"] = vars_["load_tags"].get()
        # paths
        for pkey in out["paths"]:
            k = f"paths.{pkey}"
            if k in vars_:
                out["paths"][pkey] = vars_[k].get()
        # step1
        if "step1.debug" in vars_:
            out["step1"]["debug"] = vars_["step1.debug"].get()
        if "step1.skip_split" in vars_:
            out["step1"]["skip_split"] = vars_["step1.skip_split"].get()
        if "step1.use_rfp_code_ban" in vars_:
            out["step1"]["use_rfp_code_ban"] = vars_["step1.use_rfp_code_ban"].get()
        if "step1.rfp_pipeline_mode_display" in vars_:
            lbl = vars_["step1.rfp_pipeline_mode_display"].get()
            mode_key = _RFP_PIPELINE_LABEL_TO_MODE.get(lbl, RFP_PIPELINE_MODE_STANDARD)
            out.setdefault("step1", {})["rfp_pipeline_mode"] = mode_key
            out["step1"].pop("skip_rfp_load", None)
        if "units_split_ban.enabled" in vars_:
            out.setdefault("units_split_ban", {})["enabled"] = vars_["units_split_ban.enabled"].get()
        if "units_split_ban.units" in vars_:
            parsed = _parse_list_str(vars_["units_split_ban.units"].get())
            out.setdefault("units_split_ban", {})["units"] = (
                parsed if parsed else list(DEFAULT_UNITS_SPLIT_BAN)
            )
        # step2
        if "step2.debug" in vars_:
            out["step2"]["debug"] = vars_["step2.debug"].get()
        if "step2.export_load_results_excel" in vars_:
            out["step2"]["export_load_results_excel"] = vars_["step2.export_load_results_excel"].get()
        if "step2.export_positions_database_excel" in vars_:
            out["step2"]["export_positions_database_excel"] = vars_["step2.export_positions_database_excel"].get()
        if "step2.flat_mto_structure" in vars_:
            out["step2"]["flat_mto_structure"] = vars_["step2.flat_mto_structure"].get()
        # step3
        if "step3.debug" in vars_:
            out["step3"]["debug"] = vars_["step3.debug"].get()
        if "step3.export_to_excel" in vars_:
            out["step3"]["export_to_excel"] = vars_["step3.export_to_excel"].get()
        # step4
        if "step4.debug" in vars_:
            out["step4"]["debug"] = vars_["step4.debug"].get()
        if "step4.debug_step4_1" in vars_:
            out["step4"]["debug_step4_1"] = vars_["step4.debug_step4_1"].get()
        if "step4.debug_step4_2" in vars_:
            out["step4"]["debug_step4_2"] = vars_["step4.debug_step4_2"].get()
        if "step4.debug_step4_3" in vars_:
            out["step4"]["debug_step4_3"] = vars_["step4.debug_step4_3"].get()
        if "step4.debug_step4_4" in vars_:
            out["step4"]["debug_step4_4"] = vars_["step4.debug_step4_4"].get()
        if "step4.verbose_progress_messages" in vars_:
            out["step4"]["verbose_progress_messages"] = vars_["step4.verbose_progress_messages"].get()
        if "step4.collapse_debug" in vars_:
            out["step4"]["collapse_debug"] = vars_["step4.collapse_debug"].get()
        if "step4.unified_debug" in vars_:
            out["step4"]["unified_debug"] = vars_["step4.unified_debug"].get()
        if "step4.timing_log" in vars_:
            out["step4"]["timing_log"] = vars_["step4.timing_log"].get()
        if "step4.use_multiprocessing" in vars_:
            out["step4"]["use_multiprocessing"] = vars_["step4.use_multiprocessing"].get()
        if "step4.parallel_by_title_mark" in vars_:
            out["step4"]["parallel_by_title_mark"] = vars_["step4.parallel_by_title_mark"].get()
        if "step4.max_workers" in vars_:
            try:
                val = vars_["step4.max_workers"].get()
                out["step4"]["max_workers"] = max(0, min(64, int(val)))
            except (ValueError, TypeError):
                pass
        if "step4.memory_top_stats" in vars_:
            out["step4"]["memory_top_stats"] = vars_["step4.memory_top_stats"].get()
        if "step4.memory_top_n" in vars_:
            try:
                val = vars_["step4.memory_top_n"].get()
                out["step4"]["memory_top_n"] = max(1, min(50, int(val)))
            except (ValueError, TypeError):
                pass
        if "step4.print_title_systems_table" in vars_:
            out["step4"]["print_title_systems_table"] = vars_["step4.print_title_systems_table"].get()
        if "step4.export_title_systems_comparison_excel" in vars_:
            out["step4"]["export_title_systems_comparison_excel"] = vars_["step4.export_title_systems_comparison_excel"].get()
        if "step4.export_bcc_accum_matrix" in vars_:
            out["step4"]["export_bcc_accum_matrix"] = vars_["step4.export_bcc_accum_matrix"].get()
        if "step4.assign_rfp_mto_code_compare_colors" in vars_:
            out["step4"]["assign_rfp_mto_code_compare_colors"] = vars_["step4.assign_rfp_mto_code_compare_colors"].get()
        if "step4.filter_to_mto_titles" in vars_:
            out["step4"]["filter_to_mto_titles"] = vars_["step4.filter_to_mto_titles"].get()
        if "step4.include_packing_lists" in vars_:
            out["step4"]["include_packing_lists"] = vars_["step4.include_packing_lists"].get()
        if "step4.ul_match_use_mto_tags" in vars_:
            out["step4"]["ul_match_use_mto_tags"] = vars_["step4.ul_match_use_mto_tags"].get()
        if "step4.include_mto_vo_without_rfp_anchor" in vars_:
            out["step4"]["include_mto_vo_without_rfp_anchor"] = vars_["step4.include_mto_vo_without_rfp_anchor"].get()
        if "step4.debug_tag" in vars_:
            out["step4"]["debug_tag"] = _parse_list_str(vars_["step4.debug_tag"].get())
        if "step4.debug_code" in vars_:
            out["step4"]["debug_code"] = _parse_list_str(vars_["step4.debug_code"].get())
        if "step4.debug_title_system" in vars_:
            out["step4"]["debug_title_system"] = _parse_list_str(vars_["step4.debug_title_system"].get())
        # rfp_tags_utils
        if "rfp_tags_utils.finalize_timing_top_n" in vars_:
            try:
                val = vars_["rfp_tags_utils.finalize_timing_top_n"].get()
                out["rfp_tags_utils"]["finalize_timing_top_n"] = max(1, min(50, int(val)))
            except (ValueError, TypeError):
                pass
        if "rfp_tags_utils.save_input_fingerprints" in vars_:
            out["rfp_tags_utils"]["save_input_fingerprints"] = vars_["rfp_tags_utils.save_input_fingerprints"].get()
        return out

    def do_save():
        cfg = collect_config()
        if _save(cfg):
            if on_save:
                on_save(cfg)
            win.destroy()

    btn_frame = customtkinter.CTkFrame(win, fg_color="transparent")
    btn_frame.pack(fill="x", padx=20, pady=(0, 20))
    customtkinter.CTkButton(btn_frame, text="Сохранить", command=do_save, width=120).pack(side="left", padx=(0, 10))
    customtkinter.CTkButton(btn_frame, text="Отмена", command=win.destroy, width=100).pack(side="left")

    size_label = customtkinter.CTkLabel(btn_frame, text="Размер: —", font=("Arial", 10))
    size_label.pack(side="right", padx=(20, 0))

    def update_size_text(event=None):
        w = win.winfo_width()
        h = win.winfo_height()
        if w > 1 and h > 1:
            size_label.configure(text=f"Размер: {w} × {h}")

    win.bind("<Configure>", update_size_text)
    win.after(100, update_size_text)  # обновить после отображения окна

    win.transient(parent)
    win.grab_set()
