"""GUI: edit MTO folder presets for DS vs MTO (stored in ds_compare_config.json)."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog as fd
from tkinter import messagebox
from typing import Callable

import customtkinter

from RFQ.ds_compare.ds_compare_config import (
    MtoPathPresetDict,
    load_ds_compare_config,
    normalize_ds_vs_mto_output,
    normalize_grouped_compare,
    normalize_in_cabinet_debug,
    normalize_mto_path_entries,
    save_ds_compare_config,
)


def _bind_clipboard_shortcuts_to_ctk_entry(entry: customtkinter.CTkEntry) -> None:
    """Copy/cut/paste/select-all on the inner tk.Entry (works with RU keyboard layout).

    CTkEntry is a Frame; key events go to ``_entry``. Physical keycodes match Ctrl+V/C/X/A
    regardless of the active layout (same approach as ``rfp_tags_settings_gui``).
    """
    inner = getattr(entry, "_entry", entry)

    def _top():
        return entry.winfo_toplevel()

    def do_copy(_event=None):
        try:
            text = inner.selection_get()
        except tk.TclError:
            return None
        top = _top()
        top.clipboard_clear()
        top.clipboard_append(text)
        return "break" if _event else None

    def do_cut(_event=None):
        try:
            text = inner.selection_get()
        except tk.TclError:
            return None
        top = _top()
        top.clipboard_clear()
        top.clipboard_append(text)
        try:
            inner.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        return "break" if _event else None

    def do_paste(_event=None):
        try:
            text = _top().clipboard_get()
        except tk.TclError:
            return None
        if text:
            try:
                inner.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            inner.insert("insert", text)
        return "break" if _event else None

    def do_select_all(_event=None):
        inner.selection_range(0, tk.END)
        inner.icursor(tk.END)
        return "break" if _event else None

    def on_ctrl_keypress(event: tk.Event) -> str | None:
        if not (event.state & 0x4):  # Control
            return None
        k = event.keycode
        if k == 86:  # V
            return do_paste(event)
        if k == 67:  # C
            return do_copy(event)
        if k == 88:  # X
            return do_cut(event)
        if k == 65:  # A
            return do_select_all(event)
        return None

    inner.bind("<Control-KeyPress>", on_ctrl_keypress)

    ctx = tk.Menu(entry, tearoff=0)
    ctx.add_command(label="Копировать", command=lambda: do_copy())
    ctx.add_command(label="Вставить", command=lambda: do_paste())
    ctx.add_command(label="Вырезать", command=lambda: do_cut())
    ctx.add_command(label="Выделить всё", command=lambda: do_select_all())

    def show_context_menu(event: tk.Event) -> None:
        try:
            ctx.tk_popup(event.x_root, event.y_root)
        finally:
            ctx.grab_release()

    inner.bind("<Button-3>", show_context_menu)


def show_ds_mto_path_settings(
    parent: customtkinter.CTk,
    *,
    on_saved: Callable[[], None] | None = None,
) -> None:
    """Modal-like editor for ``mto_paths`` and ``mto_path_selected_index``."""
    win = customtkinter.CTkToplevel(parent)
    win.title("Настройки: ДС vs MTO")
    win.transient(parent)
    win.grab_set()

    cfg = load_ds_compare_config()
    items: list[MtoPathPresetDict] = normalize_mto_path_entries(cfg.get("mto_paths"))
    sel_idx: list[int] = [int(cfg.get("mto_path_selected_index", 0))]
    sel_idx[0] = max(0, min(sel_idx[0], max(len(items) - 1, 0)))

    outer = customtkinter.CTkFrame(win, fg_color="transparent")
    outer.pack(fill="both", expand=True, padx=12, pady=12)

    list_frame = customtkinter.CTkFrame(outer)
    list_frame.pack(fill="both", expand=True, pady=(0, 8))

    lb = tk.Listbox(list_frame, height=12, width=72, exportselection=False)
    sb = tk.Scrollbar(list_frame, orient="vertical", command=lb.yview)
    lb.configure(yscrollcommand=sb.set)
    lb.grid(row=0, column=0, sticky="nsew")
    sb.grid(row=0, column=1, sticky="ns")
    list_frame.grid_rowconfigure(0, weight=1)
    list_frame.grid_columnconfigure(0, weight=1)

    def refresh_listbox():
        lb.delete(0, tk.END)
        for i, it in enumerate(items):
            lb.insert(tk.END, f"{i + 1}. {it['label']}  →  {it['path']}")
        if items:
            pick = min(sel_idx[0], len(items) - 1)
            lb.selection_clear(0, tk.END)
            lb.selection_set(pick)
            lb.activate(pick)
            lb.see(pick)

    lbl_entry = customtkinter.CTkLabel(outer, text="Краткое имя (в списке выбора):")
    lbl_entry.pack(anchor="w")
    entry_label = customtkinter.CTkEntry(outer, width=640)
    entry_label.pack(fill="x", pady=(0, 6))

    path_lbl = customtkinter.CTkLabel(outer, text="Путь к папке МТО:")
    path_lbl.pack(anchor="w")
    entry_path = customtkinter.CTkEntry(outer, width=640)
    entry_path.pack(fill="x", pady=(0, 8))
    _bind_clipboard_shortcuts_to_ctk_entry(entry_label)
    _bind_clipboard_shortcuts_to_ctk_entry(entry_path)

    def load_selection_to_entries():
        sel = lb.curselection()
        if not sel or not items:
            entry_label.delete(0, tk.END)
            entry_path.delete(0, tk.END)
            return
        i = int(sel[0])
        entry_label.delete(0, tk.END)
        entry_label.insert(0, items[i]["label"])
        entry_path.delete(0, tk.END)
        entry_path.insert(0, items[i]["path"])
        sel_idx[0] = i

    def on_list_select(_event=None):
        sel = lb.curselection()
        if sel:
            sel_idx[0] = int(sel[0])
        load_selection_to_entries()

    lb.bind("<<ListboxSelect>>", on_list_select)

    def browse_folder():
        try:
            win.lift()
            win.focus_force()
        except tk.TclError:
            pass
        p = fd.askdirectory(master=win, mustexist=True)
        if p:
            entry_path.delete(0, tk.END)
            entry_path.insert(0, p)

    btn_row = customtkinter.CTkFrame(outer, fg_color="transparent")
    btn_row.pack(fill="x", pady=6)

    def do_add():
        items.append({"label": "Новая папка", "path": ""})
        sel_idx[0] = len(items) - 1
        refresh_listbox()
        lb.selection_clear(0, tk.END)
        lb.selection_set(sel_idx[0])
        load_selection_to_entries()

    def do_remove():
        sel = lb.curselection()
        if not sel:
            messagebox.showwarning("Удаление", "Выберите строку в списке.", parent=win)
            return
        i = int(sel[0])
        del items[i]
        if not items:
            sel_idx[0] = 0
        else:
            sel_idx[0] = min(i, len(items) - 1)
        refresh_listbox()
        load_selection_to_entries()

    def do_apply_row():
        sel = lb.curselection()
        if not sel:
            messagebox.showwarning("Применить", "Выберите строку в списке.", parent=win)
            return
        i = int(sel[0])
        lab = entry_label.get().strip()
        path = entry_path.get().strip()
        if not path:
            messagebox.showerror("Применить", "Путь к папке не может быть пустым.", parent=win)
            return
        items[i] = {"label": lab or path, "path": path}
        refresh_listbox()
        lb.selection_set(i)
        load_selection_to_entries()

    customtkinter.CTkButton(btn_row, text="Добавить", width=100, command=do_add).pack(
        side="left", padx=(0, 6)
    )
    customtkinter.CTkButton(btn_row, text="Удалить", width=100, command=do_remove).pack(
        side="left", padx=(0, 6)
    )
    customtkinter.CTkButton(btn_row, text="Обновить строку", width=130, command=do_apply_row).pack(
        side="left", padx=(0, 6)
    )
    customtkinter.CTkButton(btn_row, text="Обзор…", width=90, command=browse_folder).pack(
        side="left", padx=(0, 6)
    )

    # --- IN_CABINET debug trace (ds_in_cabinet_trace.txt on compare) ---
    dbg_block = customtkinter.CTkFrame(outer)
    dbg_block.pack(fill="x", pady=(14, 0))

    customtkinter.CTkLabel(
        dbg_block,
        text="Отладка IN_CABINET (трассировка при «Список ДС vs MTO»)",
        font=customtkinter.CTkFont(weight="bold"),
    ).pack(anchor="w", padx=8, pady=(8, 4))

    dbg_inner = customtkinter.CTkFrame(dbg_block, fg_color="transparent")
    dbg_inner.pack(fill="x", padx=8, pady=(0, 8))

    ic_dbg = normalize_in_cabinet_debug(cfg.get("in_cabinet_debug"))
    var_debug_enabled = tk.BooleanVar(value=bool(ic_dbg.get("enabled", False)))

    chk_debug = customtkinter.CTkCheckBox(
        dbg_inner,
        text="Включить отладку (файлы ds_in_cabinet_trace.txt и ds_mto_compare_debug.log в папке Excel)",
        variable=var_debug_enabled,
    )
    chk_debug.pack(anchor="w", pady=(0, 8))

    customtkinter.CTkLabel(
        dbg_inner,
        text=(
            "Фильтры трассировки (пустое поле = все значения по этой оси). "
            "Несколько значений — через запятую. Ctrl+C / Ctrl+V в полях."
        ),
        wraplength=620,
        justify="left",
        font=customtkinter.CTkFont(size=12),
    ).pack(anchor="w", pady=(0, 6))

    def _add_dbg_field(parent, label: str, initial: str) -> customtkinter.CTkEntry:
        customtkinter.CTkLabel(parent, text=label).pack(anchor="w")
        ent = customtkinter.CTkEntry(parent, width=640)
        ent.pack(fill="x", pady=(0, 6))
        if initial:
            ent.insert(0, initial)
        _bind_clipboard_shortcuts_to_ctk_entry(ent)
        return ent

    entry_dbg_title = _add_dbg_field(
        dbg_inner, "Титул (подстрока в DS_TITLE или ключе спецификации):", ic_dbg.get("watch_title", "")
    )
    entry_dbg_mark = _add_dbg_field(
        dbg_inner,
        "Марка / раздел (TYPE_MARK, DS_SYSTEM или ключ спецификации, напр. SOS в 8529-SOS):",
        ic_dbg.get("watch_mark", ""),
    )
    entry_dbg_code = _add_dbg_field(
        dbg_inner, "Код закупочный / CODE (и CODE_2 в MTO):", ic_dbg.get("watch_code", "")
    )
    entry_dbg_cabinet = _add_dbg_field(
        dbg_inner,
        "Тег шкафа / IN_CABINET (подстрока; только строки с таким шкафом в trace):",
        ic_dbg.get("watch_cabinet", ""),
    )

    # --- Grouped DS vs MTO MVP settings ---
    grouped_block = customtkinter.CTkFrame(outer)
    grouped_block.pack(fill="x", pady=(14, 0))

    customtkinter.CTkLabel(
        grouped_block,
        text="Grouped DS vs MTO (MVP): ключи группировки",
        font=customtkinter.CTkFont(weight="bold"),
    ).pack(anchor="w", padx=8, pady=(8, 4))

    grouped_inner = customtkinter.CTkFrame(grouped_block, fg_color="transparent")
    grouped_inner.pack(fill="x", padx=8, pady=(0, 8))

    grouped_cfg = normalize_grouped_compare(cfg.get("grouped_compare"))
    var_group_code = tk.BooleanVar(value=bool(grouped_cfg.get("group_by_code", True)))
    var_group_title = tk.BooleanVar(value=bool(grouped_cfg.get("group_by_title", True)))
    var_group_system = tk.BooleanVar(value=bool(grouped_cfg.get("group_by_system", True)))
    var_group_ds_name = tk.BooleanVar(value=bool(grouped_cfg.get("group_by_ds_name", False)))
    var_group_mto_new = tk.BooleanVar(value=bool(grouped_cfg.get("group_mto_new_positions", True)))

    customtkinter.CTkLabel(
        grouped_inner,
        text="Если снять все галочки, будет использован дефолт: Титул + Марка/раздел + Код.",
        wraplength=620,
        justify="left",
        font=customtkinter.CTkFont(size=12),
    ).pack(anchor="w", pady=(0, 6))
    customtkinter.CTkCheckBox(grouped_inner, text="Код (CODE)", variable=var_group_code).pack(anchor="w")
    customtkinter.CTkCheckBox(grouped_inner, text="Титул (DS_TITLE)", variable=var_group_title).pack(anchor="w")
    customtkinter.CTkCheckBox(
        grouped_inner,
        text="Марка / раздел (DS_SYSTEM)",
        variable=var_group_system,
    ).pack(anchor="w")
    customtkinter.CTkCheckBox(
        grouped_inner,
        text="ДС (DS_NAME) — сохранить отдельные строки по ДС",
        variable=var_group_ds_name,
    ).pack(anchor="w")
    customtkinter.CTkCheckBox(
        grouped_inner,
        text="Группировать новые MTO-позиции теми же ключами (правая часть)",
        variable=var_group_mto_new,
    ).pack(anchor="w", pady=(6, 0))

    # --- DS vs MTO Excel output debug settings ---
    output_block = customtkinter.CTkFrame(outer)
    output_block.pack(fill="x", pady=(14, 0))

    customtkinter.CTkLabel(
        output_block,
        text="Excel-выгрузка DS vs MTO",
        font=customtkinter.CTkFont(weight="bold"),
    ).pack(anchor="w", padx=8, pady=(8, 4))

    output_inner = customtkinter.CTkFrame(output_block, fg_color="transparent")
    output_inner.pack(fill="x", padx=8, pady=(0, 8))

    output_cfg = normalize_ds_vs_mto_output(cfg.get("ds_vs_mto_output"))
    var_show_internal_columns = tk.BooleanVar(
        value=bool(output_cfg.get("show_internal_column_names", False))
    )
    customtkinter.CTkCheckBox(
        output_inner,
        text="Показывать технические имена колонок во второй строке заголовка (отладка)",
        variable=var_show_internal_columns,
    ).pack(anchor="w")

    bottom = customtkinter.CTkFrame(outer, fg_color="transparent")
    bottom.pack(fill="x", pady=(12, 0))

    def do_save():
        sel = lb.curselection()
        if sel and items:
            i = int(sel[0])
            lab = entry_label.get().strip()
            path = entry_path.get().strip()
            if path:
                items[i] = {"label": lab or path, "path": path}
        normalized = normalize_mto_path_entries(items)
        cfg2 = load_ds_compare_config()
        cfg2["mto_paths"] = normalized
        si = int(sel_idx[0]) if items else 0
        if normalized:
            si = max(0, min(si, len(normalized) - 1))
        else:
            si = 0
        cfg2["mto_path_selected_index"] = si
        cfg2["in_cabinet_debug"] = {
            "enabled": bool(var_debug_enabled.get()),
            "watch_title": entry_dbg_title.get().strip(),
            "watch_mark": entry_dbg_mark.get().strip(),
            "watch_code": entry_dbg_code.get().strip(),
            "watch_cabinet": entry_dbg_cabinet.get().strip(),
        }
        cfg2["grouped_compare"] = {
            "group_by_code": bool(var_group_code.get()),
            "group_by_title": bool(var_group_title.get()),
            "group_by_system": bool(var_group_system.get()),
            "group_by_ds_name": bool(var_group_ds_name.get()),
            "group_mto_new_positions": bool(var_group_mto_new.get()),
        }
        cfg2["ds_vs_mto_output"] = {
            "show_internal_column_names": bool(var_show_internal_columns.get()),
            "expand_aggregated_replacement_codes": bool(
                output_cfg.get("expand_aggregated_replacement_codes", True)
            ),
            "excel_export_mode": output_cfg.get("excel_export_mode", "both"),
            "columns": output_cfg.get("columns", []),
        }
        if save_ds_compare_config(cfg2):
            messagebox.showinfo("Сохранено", "Настройки записаны в ds_compare_config.json.", parent=win)
            if on_saved:
                on_saved()
            win.destroy()
        else:
            messagebox.showerror("Ошибка", "Не удалось сохранить файл конфигурации.", parent=win)

    customtkinter.CTkButton(bottom, text="Сохранить и закрыть", command=do_save).pack(side="right", padx=(8, 0))
    customtkinter.CTkButton(bottom, text="Отмена", command=win.destroy).pack(side="right")

    refresh_listbox()
    load_selection_to_entries()

    win.update_idletasks()
    w = min(win.winfo_reqwidth() + 24, parent.winfo_screenwidth() - 80)
    h = min(win.winfo_reqheight() + 24, parent.winfo_screenheight() - 80)
    win.geometry(f"{int(w)}x{int(h)}")
    try:
        px = parent.winfo_rootx() + 40
        py = parent.winfo_rooty() + 40
        win.geometry(f"+{px}+{py}")
    except tk.TclError:
        pass
