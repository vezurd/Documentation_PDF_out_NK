import os
import os.path
import platform
import subprocess
import sys
import tkinter
from typing import Any

import utils.path
from RFQ.DS_compare_MTO import ds_compare_start, ds_compare_folder_start
from RFQ.mto_compare import (
    mto_chain_compare_start,
    mto_compare_start,
    mto_multi_compare_start,
)
from cable_mapping import cab_mapping
from nano_cad import nc_start

from utils import folder_select
from RFQ.RFQ_compare import rfq_file_start
from RFQ.tags_rfp_compare.agregate_tags import main as agregate_tags_main
from RFQ.tags_rfp_compare.rfp_tags_settings_gui import show_rfp_tags_settings
from RFQ.tags_rfp_compare.rfp_tags_utils import (
    load_asbuild_config,
    save_asbuild_config,
    get_asbuild_config_path,
    find_latest_step4_result_file,
)
from base.base_xlsx_load import (
    FormulaCacheMissingError,
    backup_xlsx_to_old_subfolder,
    remove_strikethrough_from_xlsx,
    replace_formulas_with_cached_values,
    remove_hidden_sheets_for_1c,
)
from base.bbb_analysis import start_bbb_analysis
from base.bbb_load import find_bbb_files
from base.bbb_config import load_config as load_bbb_config
from base.excel_recalc_xlwings import recalculate_workbook_save
from base.bbb_settings_gui import show_bbb_settings
from base import google_sheets
import customtkinter
from utils.path import open_dir, get_path_from_file_path, is_file_path
from utils.release_zip import build_release_zip
from GUI.gui_constants import GuiConst
from help_descriptions import get_help_description
from gui_subprocess_job import WorkResult, attach_busy_aware_command, run_subprocess_job_on_button


def _get_mto_run_config():
    cfg = load_bbb_config()
    mto_cfg = cfg.get("mto_vs_code_base", {})
    mto_corr_cfg = cfg.get("mto_vs_code_base_correction", {})
    run_mto_check = mto_cfg.get("enabled", True)
    return run_mto_check, mto_cfg, mto_corr_cfg



# excel_template_file_name = r"templates\out_template.xlsx"
# excel_template_check_list = r"templates\out_template_check_list.xlsx"

# Пути и префиксы файлов:

# att_parsing_suffix_path_out_dir = r"debug/"
# od_parsing_suffix = "_od_parsing.xlsx"
# od_parsing_suffix_path_out_dir = r"debug/"


def open_nanocad_db():
    print("button OPEN_NANOCAD_DB")
    # Окно выбора ФАЙЛА
    file_path = folder_select.getFile()
    # Обработка исключения
    if is_file_path(file_path) is False:
        return None
    dir_path = get_path_from_file_path(file_path)
    print(dir_path, "<open_file_att>")
    print(file_path, "<open_file_att>")
    nc_start.load_db(file_path)
    return True


def open_rfq_folder():
    # Окно выбора ФАЙЛА
    file_path = folder_select.getFile()
    # Обработка исключения
    if is_file_path(file_path) is False:
        return None
    dir_path = get_path_from_file_path(file_path)
    print(dir_path, "<open_file_att>")
    print(file_path, "<open_file_att>")
    rfq_file_start(file_path, dir_path)
    return True


def open_mto_mto_folder():
    print("Запуск проверки MTO vs MTO папки")
    # Окно выбора ПАПКИ
    dir_path = folder_select.getDirectory()
    print(dir_path, "<open_pdf_folder>")
    # Обработка исключения
    if is_file_path(dir_path) is False:
        return None
    dir_name = os.path.basename(dir_path)
    print(dir_name)
    mto_compare_start(pdf_path=dir_path)


def open_mto_mto_chain_folder():
    """MTO revision chain: adjacent rev pairs per title–mark group."""
    print("Запуск проверки MTO vs MTO — цепочка ревизий")
    dir_path = folder_select.getDirectory()
    print(dir_path, "<open_mto_mto_chain_folder>")
    if is_file_path(dir_path) is False:
        return None
    try:
        err = mto_chain_compare_start(pdf_path=dir_path)
    except Exception as e:
        tkinter.messagebox.showerror(
            "MTO — цепочка ревизий",
            f"Неожиданная ошибка:\n{e!s}",
        )
        return None
    if err:
        tkinter.messagebox.showerror("MTO — цепочка ревизий", err)
        return None
    return True


def _launch_ds_compare_center() -> None:
    """Open DS vs MTO control center in a separate process."""
    cmd = [sys.executable, "-m", "ds_compare_center"]
    subprocess.Popen(
        cmd,
        cwd=os.path.dirname(os.path.abspath(__file__)),
    )


def open_remove_strikethrough():
    """Выбрать файл MTO и убрать зачёркнутый текст из листа 'Спецификация'."""
    file_path = folder_select.getFile()
    if is_file_path(file_path) is False:
        return None
    print(f"Убираем зачёркивания: {file_path}")
    try:
        backup = remove_strikethrough_from_xlsx(file_path, sheet_name="Спецификация")
        dir_path = get_path_from_file_path(file_path)
        open_dir(dir_path)
    except Exception as e:
        print(f"Ошибка: {e}")


def open_remove_strikethrough_bbb():
    """Выбрать папку с BOE/BOM/BOQ файлами и убрать зачёркнутый текст."""
    dir_path = folder_select.getDirectory()
    if is_file_path(dir_path) is False:
        return None
    found = find_bbb_files(dir_path)
    if not found:
        print("BOE/BOM/BOQ файлы не найдены в выбранной папке.")
        return None
    for doc_type in ("BOE", "BOM", "BOQ"):
        if doc_type in found:
            doc = found[doc_type]
            print(f"Убираем зачёркивания в {doc_type}: {doc.file_full_path}")
            try:
                remove_strikethrough_from_xlsx(doc.file_full_path, sheet_name=doc_type)
            except Exception as e:
                print(f"Ошибка при обработке {doc_type}: {e}")
    open_dir(dir_path)


def open_prepare_bbb_for_1c():
    """
    Папка с BOE/BOM/BOQ: резервная копия, снятие зачёркивания, пересчёт в Excel
    (кэш формул), замена формул на значения, удаление скрытых листов, снова
    пересчёт в Excel (кэш для последующих роботов). Требуются Windows, Excel, xlwings.
    """
    dir_path = folder_select.getDirectory()
    if is_file_path(dir_path) is False:
        return None
    found = find_bbb_files(dir_path)
    if not found:
        print("BOE/BOM/BOQ файлы не найдены в выбранной папке.")
        return None
    for doc_type in ("BOE", "BOM", "BOQ"):
        if doc_type not in found:
            continue
        doc = found[doc_type]
        path = doc.file_full_path
        print(f"Подготовка для 1C — {doc_type}: {path}")
        try:
            backup_xlsx_to_old_subfolder(path)
            remove_strikethrough_from_xlsx(
                path, sheet_name=doc_type, create_backup=False
            )
            recalculate_workbook_save(path, label="до замены формул")
            replace_formulas_with_cached_values(path, sheet_name=doc_type)
            remove_hidden_sheets_for_1c(path, doc_type)
            recalculate_workbook_save(path, label="после скрытых листов")
        except FormulaCacheMissingError as e:
            print(str(e))
            return None
        except Exception as e:
            print(f"Ошибка при подготовке {doc_type} для 1C: {e}")
    open_dir(dir_path)


def open_project_release_zip():
    """Собрать ZIP проекта без локального кэша/временных файлов."""
    print("Сборка релизного ZIP архива проекта")
    project_root = os.path.dirname(os.path.abspath(__file__))
    try:
        zip_path, included_files, skipped_files = build_release_zip(project_root)
        print(f"Архив создан: {zip_path}")
        print(f"Включено файлов: {included_files}; пропущено: {skipped_files}")
        open_dir(zip_path)
        return True
    except Exception as e:
        print(f"Ошибка сборки ZIP: {e}")
        return None


_agregate_tags_running = False


def _set_agregate_tags_running(v: bool) -> None:
    global _agregate_tags_running
    _agregate_tags_running = v


def _try_show_quantity_balance_fatal_dialog(project_root: str) -> None:
    """После ненулевого кода subprocess — окно с таблицей баланса из папки результатов."""
    try:
        from RFQ.tags_rfp_compare.rfp_tags_utils import load_config
        from RFQ.tags_rfp_compare.step4.step4_quantity_balance import (
            find_latest_fatal_result_dir,
            try_show_fatal_dialog_for_result_dir,
        )

        cfg = load_config()
        base = (cfg.get("paths") or {}).get("result_dir_base", "")
        if not base:
            return
        result_dir = find_latest_fatal_result_dir(base)
        if result_dir:
            try_show_fatal_dialog_for_result_dir(result_dir)
    except Exception as exc:
        print(f"Не удалось показать диалог баланса количеств: {exc}")


def _agregate_tags_subprocess_work() -> tuple:
    """Выполняется в worker-потоке: Popen + wait для проверки МТО с РКД и RFP."""
    print("Запуск проверки МТО с РКД и RFP (отдельный процесс)")
    project_root = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(project_root, "RFQ", "tags_rfp_compare", "agregate_tags.py")
    env = os.environ.copy()
    env["PYTHONPATH"] = project_root
    try:
        proc = subprocess.Popen(
            [sys.executable, "-u", script_path],
            cwd=project_root,
            env=env,
        )
        proc.wait()
        print(f"\nПроверка МТО с РКД и RFP завершена (код: {proc.returncode})")
        if proc.returncode != 0:
            _try_show_quantity_balance_fatal_dialog(project_root)
        return (proc.returncode, None)
    except Exception as e:
        print(f"\nОшибка при запуске проверки: {e}")
        return (None, e)


def wire_agregate_tags_button(app, button) -> None:
    def start_job() -> None:
        run_subprocess_job_on_button(
            app,
            button,
            base_text="Проверка МТО с РКД и RFP",
            running_label="Проверка МТО с РКД и RFP",
            work_fn=_agregate_tags_subprocess_work,
            is_running=lambda: _agregate_tags_running,
            set_running=_set_agregate_tags_running,
        )

    attach_busy_aware_command(
        app,
        button,
        is_running=lambda: _agregate_tags_running,
        start_job=start_job,
    )


_agregate_tags_asbuild_running = False


def _set_agregate_tags_asbuild_running(v: bool) -> None:
    global _agregate_tags_asbuild_running
    _agregate_tags_asbuild_running = v


def _agregate_tags_asbuild_subprocess_work() -> tuple:
    """Выполняется в worker-потоке: as-build конфиг + Popen + wait."""
    print("Запуск проверки as-build (отдельный процесс)")
    project_root = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(project_root, "RFQ", "tags_rfp_compare", "agregate_tags.py")
    asbuild_config_path = get_asbuild_config_path()
    if not os.path.isfile(asbuild_config_path):
        save_asbuild_config(load_asbuild_config())
        print(f"Создан конфиг as-build: {asbuild_config_path}")
    env = os.environ.copy()
    env["PYTHONPATH"] = project_root
    try:
        proc = subprocess.Popen(
            [sys.executable, "-u", script_path, asbuild_config_path],
            cwd=project_root,
            env=env,
        )
        proc.wait()
        print(f"\nПроверка as-build завершена (код: {proc.returncode})")
        return (proc.returncode, None)
    except Exception as e:
        print(f"\nОшибка при запуске проверки as-build: {e}")
        return (None, e)


def wire_agregate_tags_asbuild_button(app, button) -> None:
    def start_job() -> None:
        run_subprocess_job_on_button(
            app,
            button,
            base_text="Проверка as-build",
            running_label="Проверка as-build",
            work_fn=_agregate_tags_asbuild_subprocess_work,
            is_running=lambda: _agregate_tags_asbuild_running,
            set_running=_set_agregate_tags_asbuild_running,
        )

    attach_busy_aware_command(
        app,
        button,
        is_running=lambda: _agregate_tags_asbuild_running,
        start_job=start_job,
    )


def open_latest_step4_result():
    """Открывает последний по дате/времени файл Шаг4_Сопоставление_RFP_MTO_*.xlsx в Excel."""
    path = find_latest_step4_result_file()
    if not path:
        print("Файл Шаг4_Сопоставление_RFP_MTO_*.xlsx не найден в папках результатов.")
        return
    try:
        os.startfile(path)
        print(f"Открыт: {path}")
    except Exception as e:
        print(f"Ошибка при открытии файла: {e}")


def open_bbb_folder():
    """Выбрать папку с BBB (BOE/BOM/BOQ) и MTO файлами и запустить анализ."""
    print("Запуск проверки BBB")
    dir_path = folder_select.getDirectory()
    if is_file_path(dir_path) is False:
        return None
    print(f"BBB Analysis: {dir_path}")
    try:
        result_dir = start_bbb_analysis(dir_path)
        if result_dir:
            open_dir(result_dir)
    except Exception as e:
        print(f"Ошибка BBB Analysis: {e}")


def open_mto_dwg_with_optional_bbb():
    """Выбрать папку DWG с MTO и запустить MTO/BBB проверки по флагам в настройках."""
    print("Запуск проверки MTO DWG (+BBB по настройкам)")
    dir_path = folder_select.getDirectory()
    if is_file_path(dir_path) is False:
        return None

    cfg = load_bbb_config()
    folder_rules_cfg = cfg.get("folder_rules", {})
    search_only_in_dwg = folder_rules_cfg.get("search_only_in_dwg", True)
    if search_only_in_dwg:
        dir_name = os.path.basename(os.path.normpath(dir_path))
        if dir_name.upper() != "DWG":
            print(
                "Ограничение включено: выберите папку с именем DWG "
                "(или отключите флаг 'Искать только в папке DWG' в настройках)."
            )
            return None

    mto_cfg = cfg.get("mto_vs_code_base", {})
    mto_corr_cfg = cfg.get("mto_vs_code_base_correction", {})
    bbb_cb_cfg = cfg.get("bbb_vs_code_base", {})
    bbb_mto_cfg = cfg.get("bbb_vs_mto", {})
    run_mto_check = mto_cfg.get("enabled", True)
    run_bbb_check = bbb_cb_cfg.get("enabled", True) or bbb_mto_cfg.get("enabled", True)

    if not run_mto_check and not run_bbb_check:
        print("Обе проверки отключены в настройках (MTO и BBB).")
        return None

    result_dir = utils.path.get_path_out_dir(dir_path, dir_result_prefix="/__результат_MTO_BBB_")
    print(f"Папка результатов: {result_dir}")

    try:
        docs_in_dir = utils.path.get_files_single(dir_path, [".xlsx"])
        has_mto_in_dir = False
        if docs_in_dir != -1:
            has_mto_in_dir = any(getattr(doc, "doc_Type", "") == "MTO" for doc in docs_in_dir)

        if run_mto_check:
            if has_mto_in_dir:
                print("Запуск MTO-проверки...")
                google_sheets.start(
                    dir_path,
                    GuiConst.MTO_DIR,
                    out_dir=result_dir,
                    cfg=mto_cfg,
                    correction_cfg=mto_corr_cfg,
                )
            else:
                print("MTO-файл в выбранной папке не найден. MTO-проверка пропущена.")
                print("Статус MTO: skipped")
        if run_bbb_check:
            print("Запуск BBB-проверки...")
            start_bbb_analysis(dir_path, out_dir=result_dir)
        open_dir(result_dir)
    except Exception as e:
        print(f"Ошибка объединённой проверки MTO/BBB: {e}")


def open_mto_mto_multi_folder():
    print("Запуск проверки MTO vs MTO мульти, выбор папки")
    # Окно выбора ПАПКИ
    dir_path = folder_select.getDirectory()
    print(dir_path, "<open_pdf_folder>")
    # Обработка исключения
    if is_file_path(dir_path) is False:
        return None
    dir_name = os.path.basename(dir_path)
    print(dir_name)
    mto_multi_compare_start(pdf_path=dir_path)

def button_open_cj_dir():
    print("Запуск проверки ТПК папки")
    # Окно выбора ПАПКИ
    dir_path = folder_select.getDirectory()
    print(dir_path, "<open_pdf_folder>")
    # Обработка исключения
    if is_file_path(dir_path) is False:
        exit(0)
    cab_mapping.cab_mapping_start(dir_path)


def open_file_att(att):
    if att in GuiConst.dict:
        print(GuiConst.dict[att][0])
    # Окно выбора ФАЙЛА
    file_path = folder_select.getFile()
    # Обработка исключения
    if is_file_path(file_path) is False:
        return None
    dir_path = get_path_from_file_path(file_path)
    print(dir_path, "<open_file_att>")
    print(file_path, "<open_file_att>")
    if att == GuiConst.MTO_FILE:
        run_mto_check, mto_cfg, mto_corr_cfg = _get_mto_run_config()
        if not run_mto_check:
            print("Проверка MTO отключена в настройках.")
            return None
        result = google_sheets.start(
            dir_path,
            att,
            file_path,
            cfg=mto_cfg,
            correction_cfg=mto_corr_cfg,
        )
    else:
        result = google_sheets.start(dir_path, att, file_path)
    open_dir(result)
    return True


def open_dir_att(att, out_dir=None):
    if att in GuiConst.dict:
        print(GuiConst.dict[att][0])
    # Окно выбора ПАПКИ
    dir_path = folder_select.getDirectory()
    print(dir_path, "<open_dir_att>")
    # Обработка исключения
    if is_file_path(dir_path) is False:
        return None
    if att == GuiConst.MTO_DIR:
        run_mto_check, mto_cfg, mto_corr_cfg = _get_mto_run_config()
        if not run_mto_check:
            print("Проверка MTO отключена в настройках.")
            return None
        result = google_sheets.start(
            dir_path,
            att,
            out_dir=out_dir,
            cfg=mto_cfg,
            correction_cfg=mto_corr_cfg,
        )
    else:
        result = google_sheets.start(dir_path, att, out_dir=out_dir)
    open_dir(result)
    return True


def _bind_readonly_clipboard_to_ctk_textbox(textbox: customtkinter.CTkTextbox) -> None:
    """Copy / select-all on a disabled CTkTextbox (Ctrl+C works with any keyboard layout)."""
    inner = getattr(textbox, "_textbox", textbox)

    def _top() -> tkinter.Misc:
        return textbox.winfo_toplevel()

    def do_copy(_event: tkinter.Event | None = None) -> str | None:
        try:
            if not inner.tag_ranges("sel"):
                return None
            text = inner.get("sel.first", "sel.last")
        except tkinter.TclError:
            return None
        top = _top()
        top.clipboard_clear()
        top.clipboard_append(text)
        return "break" if _event is not None else None

    def do_select_all(_event: tkinter.Event | None = None) -> str | None:
        try:
            inner.tag_add("sel", "1.0", "end-1c")
        except tkinter.TclError:
            pass
        return "break" if _event is not None else None

    def on_ctrl_keypress(event: tkinter.Event) -> str | None:
        if not (event.state & 0x4):
            return None
        k = event.keycode
        if k == 67:
            return do_copy(event)
        if k == 65:
            return do_select_all(event)
        return None

    inner.bind("<Control-KeyPress>", on_ctrl_keypress)

    ctx = tkinter.Menu(textbox, tearoff=0)
    ctx.add_command(label="Копировать", command=lambda: do_copy())
    ctx.add_command(label="Выделить всё", command=lambda: do_select_all())

    def show_context_menu(event: tkinter.Event) -> None:
        try:
            ctx.tk_popup(event.x_root, event.y_root)
        finally:
            ctx.grab_release()

    inner.bind("<Button-3>", show_context_menu)


def show_help(help_key):
    """Показывает справку по функции"""
    help_info = get_help_description(help_key)
    
    # Создаем окно справки
    help_window = customtkinter.CTkToplevel()
    help_window.title(f"Справка: {help_info['title']}")
    help_window.geometry("760x580")
    help_window.resizable(True, True)
    
    # Делаем окно модальным
    help_window.transient()
    help_window.grab_set()
    
    # Создаем фрейм для содержимого
    main_frame = customtkinter.CTkFrame(help_window)
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)
    
    # Заголовок
    title_label = customtkinter.CTkLabel(
        main_frame, 
        text=f"{help_info['icon']} {help_info['title']}", 
        font=("Arial", 16, "bold")
    )
    title_label.pack(pady=(0, 20))
    
    _mono = "Consolas" if platform.system() == "Windows" else "DejaVu Sans Mono"
    # Текст справки (дерево цепочек — без переноса по словам)
    help_text = customtkinter.CTkTextbox(
        main_frame,
        font=(_mono, 11),
        wrap="none",
    )
    help_text.pack(fill="both", expand=True, pady=(0, 20))
    
    # Вставляем текст справки
    help_text.insert("1.0", help_info['description'])
    help_text.configure(state="disabled")  # только чтение; копирование — через _bind_…
    _bind_readonly_clipboard_to_ctk_textbox(help_text)

    # Кнопка закрытия
    close_button = customtkinter.CTkButton(
        main_frame,
        text="Закрыть",
        command=help_window.destroy,
        width=100
    )
    close_button.pack(pady=(0, 10))
    
    # Центрируем окно
    help_window.update_idletasks()
    x = (help_window.winfo_screenwidth() // 2) - (help_window.winfo_width() // 2)
    y = (help_window.winfo_screenheight() // 2) - (help_window.winfo_height() // 2)
    help_window.geometry(f"+{x}+{y}")


class SetButton(customtkinter.CTkButton):
    def __init__(self, master, key, command: Any, row, column, padx=20, pady=20, width=200):
        super().__init__(master)
        self.button = customtkinter.CTkButton(master, text=GuiConst.dict[key][1], command=lambda: command(key),
                                              width=width)
        self.button.grid(row=row, column=column, padx=padx, pady=pady, sticky="ew")


class ButtonWithHelp:
    """Класс для создания кнопки с иконкой справки (или альтернативной вторичной кнопкой)."""
    def __init__(self, master, text, command, help_key, row, column, padx=20, pady=12, width=200,
                 secondary_command=None, secondary_icon=None, columnspan=1,
                 frame_sticky: str = "ew", button_fill: bool = True):
        self.master = master
        self.help_key = help_key

        # Создаем фрейм для кнопки и иконки (прозрачный фон)
        self.frame = customtkinter.CTkFrame(master, fg_color="transparent")
        self.frame.grid(row=row, column=column, columnspan=columnspan, padx=padx, pady=pady, sticky=frame_sticky)

        # Настраиваем grid weights для фрейма
        btn_col_weight = 1 if button_fill else 0
        self.frame.grid_columnconfigure(0, weight=btn_col_weight)
        self.frame.grid_columnconfigure(1, weight=0)

        # Создаем кнопку
        self.button = customtkinter.CTkButton(
            self.frame,
            text=text,
            command=command,
            width=width
        )
        btn_sticky = "ew" if button_fill else "w"
        self.button.grid(row=0, column=0, padx=(0, 5), pady=5, sticky=btn_sticky)

        # Вторичная кнопка: открытие файла или справка
        if secondary_command is not None and secondary_icon is not None:
            btn_text = secondary_icon
            btn_cmd = secondary_command
            btn_kw = {"width": 30, "height": 30, "corner_radius": 15}
        else:
            btn_text = "i"
            btn_cmd = lambda: show_help(help_key)
            btn_kw = {"width": 30, "height": 30, "font": ("Arial", 14, "bold"), "corner_radius": 15}
        self.help_button = customtkinter.CTkButton(
            self.frame, text=btn_text, command=btn_cmd, **btn_kw
        )
        self.help_button.grid(row=0, column=1, padx=(0, 5), pady=5, sticky="")


class SetLabel(customtkinter.CTkLabel):
    def __init__(self, master, text, row, column, padx: int = 10, pady: int = 10):
        """

        :type
        :type padx: int
        :type pady: int
        """
        super().__init__(master)
        self.label = customtkinter.CTkLabel(master=master, text=text)  # создаем текстовую метку
        self.label.grid(row=row, column=column, padx=padx, pady=pady, sticky="w")


if __name__ == "__main__":
    print("main start")
    # создаем окно main
    # customtkinter.set_widget_scaling(1.5)  # widget dimensions and check_text size
    # customtkinter.set_window_scaling(1.5)  # window geometry dimensions
    app = customtkinter.CTk()
    folder_select.set_main_window(app)  # диалоги выбора — дочерние окна, выводятся на передний план
    app.title("АГХК проверка РД (2026.03.10)")
    # Разрешаем изменение размера окна
    app.resizable(True, True)
    # Убираем фиксированный размер - окно будет подстраиваться под контент
    # app.geometry("820x700")

    # notebook = customtkinter.

    # Минимальная ширина контента (canvas); уменьшена вместе с узким ComboBox MTO в блоке ДС.
    frame_width = 480

    # Прокрутка: полоса только если высота окна меньше контента (без увеличения стартового размера)
    scroll_outer = customtkinter.CTkFrame(app)
    scroll_outer.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
    app.grid_rowconfigure(0, weight=1)
    app.grid_columnconfigure(0, weight=1)
    scroll_outer.grid_rowconfigure(0, weight=1)
    scroll_outer.grid_columnconfigure(0, weight=1)

    _canvas_bg = scroll_outer._apply_appearance_mode(scroll_outer.cget("fg_color"))
    canvas = tkinter.Canvas(
        scroll_outer,
        highlightthickness=0,
        borderwidth=0,
        background=_canvas_bg,
    )
    scrollbar = customtkinter.CTkScrollbar(
        scroll_outer, orientation="vertical", command=canvas.yview
    )

    def _sync_main_scrollbar() -> None:
        canvas.update_idletasks()
        bbox = canvas.bbox("all")
        if not bbox:
            scrollbar.grid_remove()
            return
        content_h = bbox[3] - bbox[1]
        view_h = canvas.winfo_height()
        if content_h > view_h:
            scrollbar.grid(row=0, column=1, sticky="ns", padx=(4, 0), pady=0)
        else:
            scrollbar.grid_remove()
            canvas.yview_moveto(0)

    def _main_yscrollcommand(*args):
        scrollbar.set(*args)

    canvas.configure(yscrollcommand=_main_yscrollcommand)
    canvas.grid(row=0, column=0, sticky="nsew")

    frame1 = customtkinter.CTkFrame(canvas, fg_color="transparent")
    inner_win = canvas.create_window((0, 0), window=frame1, anchor="nw")

    def _on_main_inner_configure(_event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        _sync_main_scrollbar()

    def _on_main_canvas_configure(event):
        canvas.itemconfigure(inner_win, width=event.width)
        _on_main_inner_configure()

    frame1.bind("<Configure>", _on_main_inner_configure)
    canvas.bind("<Configure>", _on_main_canvas_configure)

    def _main_mousewheel_filtered(event):
        if event.widget.winfo_toplevel() != app:
            return
        if not scrollbar.winfo_ismapped():
            return
        w = event.widget
        while w is not None:
            if w == canvas or w == scroll_outer:
                if sys.platform.startswith("win"):
                    canvas.yview_scroll(int(-event.delta / 120), "units")
                elif getattr(event, "num", None) == 4:
                    canvas.yview_scroll(-1, "units")
                elif getattr(event, "num", None) == 5:
                    canvas.yview_scroll(1, "units")
                return "break"
            w = getattr(w, "master", None)

    app.bind_all("<MouseWheel>", _main_mousewheel_filtered, add="+")
    if sys.platform != "darwin":
        app.bind_all("<Button-4>", _main_mousewheel_filtered, add="+")
        app.bind_all("<Button-5>", _main_mousewheel_filtered, add="+")

    def _initial_main_window_fit():
        app.update_idletasks()
        tw = max(frame1.winfo_reqwidth(), frame_width, 1)
        th = max(frame1.winfo_reqheight(), 1)
        canvas.itemconfigure(inner_win, width=tw)
        canvas.configure(scrollregion=(0, 0, tw, th))
        _sync_main_scrollbar()
        # padx/pady=20 на scroll_outer относительно app → +40 к ширине и высоте клиентской области
        margin = 40
        app.geometry(f"{tw + margin}x{th + margin}")

    frame1.grid_columnconfigure(0, weight=1)
    frame1.grid_columnconfigure(1, weight=1)
    frame1.grid_columnconfigure(2, weight=1)
    _SEP_PADY = 6

    ####################################################################################
    # 01. ПРОВЕРКА PDF ДОКУМЕНТОВ                                                      #
    ####################################################################################
    SetLabel(master=frame1, text="📄 ПРОВЕРКА PDF", row=0, column=0)
    pdf_dir_frame = customtkinter.CTkFrame(frame1, fg_color="transparent")
    pdf_dir_frame.grid(row=0, column=1, padx=20, pady=12, sticky="ew")
    pdf_dir_frame.grid_columnconfigure(0, weight=1)

    def _launch_v2_control_center() -> None:
        import subprocess as _sp

        cmd = [sys.executable, "-m", "pdf_v2_monitor"]
        _sp.Popen(
            cmd,
            cwd=os.path.dirname(os.path.abspath(__file__)),
        )

    btn_v2_monitor = customtkinter.CTkButton(
        pdf_dir_frame,
        text="Центр управления PDF v2",
        width=260,
        command=_launch_v2_control_center,
    )
    btn_v2_monitor.grid(row=0, column=0, padx=(0, 5), pady=5, sticky="ew")
    customtkinter.CTkButton(
        pdf_dir_frame,
        text="i",
        width=30,
        height=30,
        font=("Arial", 14, "bold"),
        corner_radius=15,
        command=lambda: show_help("open_pdf_folder"),
    ).grid(row=0, column=1, padx=(0, 5), pady=5, sticky="")
    customtkinter.CTkLabel(
        pdf_dir_frame,
        text="Полный прогон папки PDF: шаблоны, извлечение, теги, ОД, НК — только здесь.",
        text_color="gray",
        wraplength=420,
        justify="left",
    ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 0))

    customtkinter.CTkFrame(frame1, height=2, fg_color="gray").grid(
        row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=_SEP_PADY)

    ####################################################################################
    # 02. ПРОВЕРКА MTO ФАЙЛОВ                                                          #
    ####################################################################################
    SetLabel(master=frame1, text="📊 ПРОВЕРКА MTO", row=2, column=0)
    mto_dir_frame = customtkinter.CTkFrame(frame1, fg_color="transparent")
    mto_dir_frame.grid(row=2, column=1, padx=20, pady=12, sticky="ew")
    mto_dir_frame.grid_columnconfigure(0, weight=1)
    customtkinter.CTkButton(
        mto_dir_frame,
        text=GuiConst.dict[GuiConst.MTO_DIR][1],
        command=open_mto_dwg_with_optional_bbb,
        width=200
    ).grid(row=0, column=0, padx=(0, 5), pady=5, sticky="ew")
    customtkinter.CTkButton(
        mto_dir_frame,
        text="⚙",
        width=36,
        command=lambda: show_bbb_settings(parent=app)
    ).grid(row=0, column=1, padx=(0, 5), pady=5, sticky="")
    ButtonWithHelp(master=frame1, text=GuiConst.dict[GuiConst.MTO_FILE][1], command=lambda: open_file_att(GuiConst.MTO_FILE), 
                   help_key="mto_file", row=2, column=2)

    ButtonWithHelp(master=frame1, text=GuiConst.dict[GuiConst.BOOT_FILE][1], command=lambda: open_file_att(GuiConst.BOOT_FILE), 
                   help_key="boot_file", row=3, column=1)
    ButtonWithHelp(master=frame1, text="Открыть файл RFQ", command=open_rfq_folder, 
                   help_key="open_rfq_folder", row=3, column=2)
    ButtonWithHelp(master=frame1, text=GuiConst.dict[GuiConst.OUTPUT_FILE][1], command=lambda: open_file_att(GuiConst.OUTPUT_FILE), 
                   help_key="output_file", row=4, column=1)

    customtkinter.CTkButton(
        frame1, text="Убрать зачеркивания",
        command=open_remove_strikethrough, width=200
    ).grid(row=4, column=2, padx=20, pady=12, sticky="ew")

    customtkinter.CTkLabel(frame1, text="Синхронизация с ГуглТаблицей — автоматическая",
                           text_color="gray").grid(row=5, column=1, padx=20, pady=12, sticky="w")

    bbb_strike_1c_frame = customtkinter.CTkFrame(frame1, fg_color="transparent")
    bbb_strike_1c_frame.grid(row=5, column=2, padx=20, pady=12, sticky="ew")
    bbb_strike_1c_frame.grid_columnconfigure(0, weight=1)
    customtkinter.CTkButton(
        bbb_strike_1c_frame,
        text="Убрать зачеркивания (BOE/BOM/BOQ)",
        command=open_remove_strikethrough_bbb,
        width=200,
    ).grid(row=0, column=0, pady=(0, 6), sticky="ew")
    customtkinter.CTkButton(
        bbb_strike_1c_frame,
        text="Для 1C",
        command=open_prepare_bbb_for_1c,
        width=200,
    ).grid(row=1, column=0, sticky="ew")

    customtkinter.CTkFrame(frame1, height=2, fg_color="gray").grid(
        row=6, column=0, columnspan=3, sticky="ew", padx=10, pady=_SEP_PADY)

    ####################################################################################
    # 03. ПРОВЕРКА КАБЕЛЬНЫХ ЖУРНАЛОВ                                                 #
    ####################################################################################
    SetLabel(master=frame1, text="⚡ КАБЕЛЬНЫЕ ЖУРНАЛЫ", row=7, column=0)
    ButtonWithHelp(master=frame1, text="КЖ: выбрать папку", command=button_open_cj_dir, 
                   help_key="button_open_cj_dir", row=7, column=1)

    customtkinter.CTkFrame(frame1, height=2, fg_color="gray").grid(
        row=8, column=0, columnspan=3, sticky="ew", padx=10, pady=_SEP_PADY)

    ####################################################################################
    # 04. СРАВНЕНИЕ MTO ФАЙЛОВ                                                         #
    ####################################################################################
    SetLabel(master=frame1, text="🔄 СРАВНЕНИЕ MTO", row=9, column=0)
    ButtonWithHelp(master=frame1, text="MTO vs MTO (2 шт.)", command=open_mto_mto_folder, 
                   help_key="open_mto_mto_folder", row=9, column=1)
    ButtonWithHelp(master=frame1, text="MTO vs MTO (мульти)", command=open_mto_mto_multi_folder, 
                   help_key="open_mto_mto_multi_folder", row=9, column=2)
    ButtonWithHelp(
        master=frame1,
        text="MTO: цепочка ревизий",
        command=open_mto_mto_chain_folder,
        help_key="open_mto_mto_chain_folder",
        row=10,
        column=1,
    )

    ####################################################################################
    # 05. ПРОВЕРКА ДС ФАЙЛОВ                                                           #
    ####################################################################################
    SetLabel(master=frame1, text="📋 ПРОВЕРКА ДС", row=11, column=0)
    ds_center_frame = customtkinter.CTkFrame(frame1, fg_color="transparent")
    ds_center_frame.grid(row=11, column=1, columnspan=2, padx=12, pady=12, sticky="ew")
    ds_center_frame.grid_columnconfigure(0, weight=1)
    customtkinter.CTkButton(
        ds_center_frame,
        text="Центр управления ДС vs MTO",
        width=260,
        command=_launch_ds_compare_center,
    ).grid(row=0, column=0, sticky="ew")
    customtkinter.CTkLabel(
        ds_center_frame,
        text="Merge ДС, ДС vs MTO, grouped, MTO vs ДС — отдельное окно с настройками и сохранением путей.",
        text_color="gray",
        wraplength=520,
        justify="left",
    ).grid(row=1, column=0, sticky="w", pady=(4, 0))

    customtkinter.CTkFrame(frame1, height=2, fg_color="gray").grid(
        row=12, column=0, columnspan=3, sticky="ew", padx=10, pady=_SEP_PADY)

    ####################################################################################
    # 07. ПРОВЕРКА МТО С РКД И RFP                                                      #
    ####################################################################################
    SetLabel(master=frame1, text="Проверка МТО с РКД и RFP", row=13, column=0)
    btn_rfp_mto_check = ButtonWithHelp(
        master=frame1,
        text="Проверка МТО с РКД и RFP",
        command=lambda: None,
        help_key="open_agregate_tags",
        row=13,
        column=1,
        secondary_command=open_latest_step4_result,
        secondary_icon="📊",
    )
    wire_agregate_tags_button(app, btn_rfp_mto_check.button)
    customtkinter.CTkButton(
        frame1, text="⚙️ Настройки", width=120,
        command=lambda: show_rfp_tags_settings(parent=app)
    ).grid(row=13, column=2, padx=20, pady=12, sticky="w")

    ####################################################################################
    # 08. ПРОВЕРКА as-build                                                            #
    ####################################################################################
    SetLabel(master=frame1, text="Проверка as-build", row=14, column=0)
    btn_asbuild_check = ButtonWithHelp(
        master=frame1,
        text="Проверка as-build",
        command=lambda: None,
        help_key="open_agregate_tags_asbuild",
        row=14,
        column=1,
        secondary_command=open_latest_step4_result,
        secondary_icon="📊",
    )
    wire_agregate_tags_asbuild_button(app, btn_asbuild_check.button)
    customtkinter.CTkButton(
        frame1, text="⚙️ Настройки", width=120,
        command=lambda: show_rfp_tags_settings(
            parent=app,
            load_fn=load_asbuild_config,
            save_fn=save_asbuild_config,
            title_suffix=" (as-build)",
        )
    ).grid(row=14, column=2, padx=20, pady=12, sticky="w")

    customtkinter.CTkFrame(frame1, height=2, fg_color="gray").grid(
        row=15, column=0, columnspan=3, sticky="ew", padx=10, pady=_SEP_PADY)

    ####################################################################################
    # 09. РАБОТА С БАЗАМИ ДАННЫХ                                                        #
    ####################################################################################
    SetLabel(master=frame1, text="🗄️ БАЗЫ ДАННЫХ", row=16, column=0)
    ButtonWithHelp(master=frame1, text="Открыть NanoCAD.db", command=open_nanocad_db, 
                   help_key="open_nanocad_db", row=16, column=1)

    ####################################################################################
    # 10. ЭКСПОРТ ПРОЕКТА                                                              #
    ####################################################################################
    SetLabel(master=frame1, text="📦 ЭКСПОРТ ПРОЕКТА", row=17, column=0)
    ButtonWithHelp(master=frame1, text="Собрать ZIP для коллег", command=open_project_release_zip,
                   help_key="open_project_release_zip", row=17, column=1)

    app.update_idletasks()
    _initial_main_window_fit()
    app.mainloop()
    print("zaloop")
