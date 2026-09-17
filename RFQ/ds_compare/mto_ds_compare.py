from RFQ.ds_compare.DsMtoTable import DsMtoTable
from RFQ.ds_compare.mto_check import MtoCheck
from RFQ.ds_compare.support_finctions import load_ds_data_list
from base import t_comm_initial_classes
from base.base_cheks import check_position_row
from base.base_class_std_table import STDTable
from base.base_classes import TableComments, RowStd
from base.base_mto import get_mto_std_from_file
from pdf_parsing_v2_engine.document import V2Document
from base.tables_columns import *
from colorama import Fore, Back, Style, init

from utils.colors import Color
from utils.path import get_path_out_dir, open_dir
import customtkinter as ctk
import tkinter as tk

# Инициализация colorama для Windows
init(autoreset=True)

def mto_ds_list_compare (mto_file_path, force_ds_update: bool = False):
    """
        Задача: Проверка MTO рабочей документации по закупочным DS спецификациям
        
    Args:
        mto_file_path: Путь к MTO файлу
        force_ds_update: Принудительное обновление данных DS из файла (игнорирует кэш)
    """
    print("Сравнение МТО с списком строк из ДС\n", mto_file_path)

    ds_list_file_path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\Спецификации к ДС в Excel\Резальтат автоматической сверки ДС с МТО\РОБОТ_СРАВНЕНИЕ__ИСХОДНАЯ СУММА_ДС13-ДС50.xlsx"

    # Загрузка данных с поддержкой кэширования
    ds_data = load_ds_data_list(ds_list_file_path, force_update=force_ds_update)
    ds_data_obj = DsMtoTable(ds_data)
    ds_data_obj.debug = False
    ds_data_obj.get_ds_mto_list()
    

    mto_obj = V2Document.from_file_path(mto_file_path)
    mto_base = get_mto_std_from_file(t_com=TableComments(file_full_path=mto_file_path,
                                                                       dir_path="-1",
                                                                       tabel_class=t_comm_initial_classes.MTO,
                                                                       )
                                                   , dbg=0)


    mto_base_positions = [row for row in mto_base if check_position_row(row)]

    """
    Проверки МТО
    """
    # Проверка на Теги
    spec_dict = {mto_obj.doc_Short_Title:mto_base_positions}
    mto_check_obj = MtoCheck(spec_dict)
    mto_check_obj.check()

    #Сравнение количества в МТО и DS
    print(f"\n{Fore.CYAN}{'='*60}")
    print(f"{Fore.CYAN}📊 СРАВНЕНИЕ КОЛИЧЕСТВА В МТО И DS")
    print(f"{Fore.CYAN}{'='*60}")
    
    str_title_system = mto_obj.doc_Short_Title
    ds_mto_title_system = ds_data_obj.get_rows_by_ds_title_system(str_title_system)
    ds_value = ds_data_obj.calculate_sum_by_column(ds_mto_title_system, VALUES_2)
    mto_value = ds_data_obj.calculate_sum_by_column(mto_base_positions, VALUES)
    diff = abs(ds_value - mto_value)
    
    print(f"\n{Fore.YELLOW}📋 Система: {Style.BRIGHT}{str_title_system}")
    print(f"{Fore.BLUE}📈 DS сумма: {Style.BRIGHT}{Fore.GREEN}{ds_value:,.2f}")
    print(f"{Fore.BLUE}📈 МТО сумма: {Style.BRIGHT}{Fore.GREEN}{mto_value:,.2f}")
    
    if diff == 0:
        print(f"\n{Fore.GREEN}✅ {Style.BRIGHT}СУММЫ СОВПАДАЮТ!")        
    else:
        print(f"\n{Fore.RED}❌ {Style.BRIGHT}СУММЫ НЕ СОВПАДАЮТ!")
        print(f"{Fore.RED}⚠️  Разность: {Style.BRIGHT}{Fore.WHITE}{Back.RED} {diff:,.2f} {Style.RESET_ALL}")
        print(f"{Fore.RED}📊 Детали: {ds_value:,.2f} - {mto_value:,.2f} = {diff:,.2f}")
    
    print(f"\n{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
    # Сравнение кол-ва по кодам в МТО и DS
    print(f"\n{Fore.CYAN}{'='*60}")
    print(f"{Fore.CYAN}📊 СРАВНЕНИЕ КОЛИЧЕСТВА ПО КОДАМ В МТО И DS")
    print(f"{Fore.CYAN}{'='*60}")
    ds_code_dict = ds_data_obj.calculate_sum_by_code(ds_mto_title_system, VALUES_2, CODE_2)
    mto_code_dict = ds_data_obj.calculate_sum_by_code(mto_base_positions, VALUES, CODE)
    codes_comparison_result = ds_data_obj.compare_dictionaries(ds_code_dict, mto_code_dict)
    if codes_comparison_result['identical']:
        print(f"\n{Fore.GREEN}✅ {Style.BRIGHT}СУММЫ ПО КОДАМ СОВПАДАЮТ!")
    else:
        print(f"\n{Fore.RED}❌ {Style.BRIGHT}СУММЫ ПО КОДАМ НЕ СОВПАДАЮТ!")
        print(f"{Fore.RED}⚠️  Разность: {Style.BRIGHT}{Fore.WHITE}{Back.RED} {codes_comparison_result['summary']['total_difference']:,.2f} {Style.RESET_ALL}")
        print(f"Кодов в ДС словаре: {codes_comparison_result['summary']['total_codes_first']}")
        print(f"Кодов во МТО словаре: {codes_comparison_result['summary']['total_codes_second']}")
        print(f"Общих кодов: {codes_comparison_result['summary']['common_codes']}")
        print(f"Кодов только в ДС: {len(codes_comparison_result['only_in_first'])}")
        print(f"Кодов только во МТО: {len(codes_comparison_result['only_in_second'])}")
        print(f"Кодов с разными значениями: {len(codes_comparison_result['different_values'])}")
        print(f"Общая разность: {codes_comparison_result['summary']['total_difference']}")
        if codes_comparison_result['only_in_first']:
            print("\nКоды только в ДС словаре:")
            for key, value in codes_comparison_result['only_in_first'].items():
                print(f"  {key}: {value}")

        if codes_comparison_result['only_in_second']:
            print("\nКоды только во МТО словаре:")
            for key, value in codes_comparison_result['only_in_second'].items():
                print(f"  {key}: {value}")

        if codes_comparison_result['different_values']:
            print("\nКоды с разными значениями в ДС и МТО:")
            for key, diff_info in codes_comparison_result['different_values'].items():
                print(
                    f"  {key}: {diff_info['first_value']} vs {diff_info['second_value']} (разность: {diff_info['difference']:.6f})")

    print(f"\n{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
    """
    Конец проверок
    """
    # Берем DS-строки для анализа замен
    str_title_system = mto_obj.doc_Short_Title
    ds_mto_title_system = ds_data_obj.get_rows_by_ds_title_system(str_title_system)

    """
    Перераспределение количеств в МТО: перенос части количества с нового кода (CODE_2)
    на старый код из ДС (CODE) по строкам со статусом замены.
    Если строки со старым кодом нет в МТО — создаем рядом со строкой-источником.
    """

    # Составляем агрегированные потребности переноса: new_code -> {old_code: qty}
    transfers_by_new_code: dict[str, dict[str, float]] = {}
    # Резерв по коду: сколько ДС требует оставить на самом new_code (без замены)
    direct_ds_demand_by_code: dict[str, float] = {}
    for ds_row in ds_mto_title_system:
        old_code = ds_row.get_value(CODE)
        new_code = ds_row.get_value(CODE_2)
        qty = ds_row.get_value(VALUES)
        try:
            qty_val = float(qty) if qty is not None else 0.0
        except Exception:
            qty_val = 0.0
        if not new_code or qty_val <= 0:
            continue
        # Если нет замены (old_code пустой/None) или код не меняется — это прямой спрос по new_code
        if not old_code or old_code == "" or old_code == new_code:
            direct_ds_demand_by_code[str(new_code)] = direct_ds_demand_by_code.get(str(new_code), 0.0) + qty_val
            continue
        # Иначе это потребность переноса с new_code -> old_code
        transfers_by_new_code.setdefault(str(new_code), {}).setdefault(str(old_code), 0.0)
        transfers_by_new_code[str(new_code)][str(old_code)] += qty_val

    if transfers_by_new_code:
        print(f"\n{Fore.CYAN}{'='*60}")
        print(f"{Fore.CYAN}🔁 ПЕРЕРАСПРЕДЕЛЕНИЕ КОЛИЧЕСТВ В МТО ПО ЗАМЕНАМ")
        print(f"{Fore.CYAN}{'='*60}")

    # Индекс помогающий быстро находить строки по коду (игнорируем комментарий "Заменить на")
    def _normalize_code(val) -> str:
        try:
            text = str(val) if val is not None else ""
            # Берем первую строку до перевода строки
            return text.splitlines()[0].strip()
        except Exception:
            return str(val)

    def get_rows_by_code(code: str) -> list:
        target = _normalize_code(code)
        return [
            row for row in mto_base
            if check_position_row(row) and _normalize_code(row.get_value(CODE)) == target
        ]

    # Для постподсветки копим сообщения по строкам источников
    highlight_notes_by_row_id: dict[int, dict] = {}

    for new_code, needs_map in transfers_by_new_code.items():
        # Источники: строки МТО с new_code
        source_rows = get_rows_by_code(new_code)
        if not source_rows:
            print(f"{Fore.YELLOW}⚠️ Нет строк в МТО для кода-источника {new_code} — пропуск")
            continue

        # Общий доступный объем у источников
        total_available = sum((row.get_value(VALUES) or 0) for row in source_rows)
        total_needed = sum(needs_map.values())
        if total_available <= 0:
            print(f"{Fore.YELLOW}⚠️ Нулевой остаток по {new_code} — пропуск")
            continue

        # Резервируем под прямой спрос ДС по этому коду
        reserve = direct_ds_demand_by_code.get(str(new_code), 0.0)
        movable_available = max(total_available - reserve, 0.0)
        if total_needed > movable_available:
            print(f"{Fore.YELLOW}⚠️ Требуемый перенос {total_needed} превышает доступный для переноса {movable_available} (резерв {reserve}) по {new_code}. Переносим частично.")

        # Перебираем старые коды, которым нужно добавить количество
        for old_code, need_qty in needs_map.items():
            remaining = float(need_qty)
            if remaining <= 0:
                continue

            # Получатели: существующие строки со старым кодом (если есть)
            target_rows = get_rows_by_code(old_code)

            # Идем по источникам в том порядке, в каком они в МТО, стараясь не дробить лишнего
            allowed_to_move = movable_available
            for idx, src_row in enumerate(source_rows):
                if remaining <= 0:
                    break
                src_available = src_row.get_value(VALUES) or 0
                if allowed_to_move <= 0:
                    break
                
                # Проверяем, есть ли в ds_mto_title_system строки, которые требуют незамененного кода new_code
                direct_demand_rows = []
                for ds_row in ds_mto_title_system:
                    ds_new_code = ds_row.get_value(CODE_2)
                    ds_old_code = ds_row.get_value(CODE)
                    ds_qty = ds_row.get_value(VALUES)
                    try:
                        ds_qty_val = float(ds_qty) if ds_qty is not None else 0.0
                    except Exception:
                        ds_qty_val = 0.0
                    
                    # Если это строка с прямым спросом по new_code (без замены)
                    if (ds_new_code == new_code and 
                        (not ds_old_code or ds_old_code == "" or ds_old_code == new_code) and 
                        ds_qty_val > 0):
                        direct_demand_rows.append(ds_row)
                
                # Если есть строки с прямым спросом, добавляем комментарий к источнику
                if direct_demand_rows:
                    total_direct_demand = sum(float(row.get_value(VALUES) or 0) for row in direct_demand_rows)
                    comment_text = f"Прямой спрос ДС: {total_direct_demand:,.2f} (строк: {len(direct_demand_rows)})"
                    
                    # Добавляем комментарий к источнику
                    src_annotation = src_row.get_value(ANNOTATION) or ""
                    sep = "\n" if src_annotation else ""
                    src_row.el[ANNOTATION].value = f"{src_annotation}{sep}{comment_text}"
                    
                    # Добавляем строку с прямым спросом в highlight_notes_by_row_id для последующей подсветки
                    rid = id(src_row)
                    note = highlight_notes_by_row_id.get(rid)
                    if not note:
                        note = {"row": src_row, "messages": [], "is_direct_demand": True}
                        highlight_notes_by_row_id[rid] = note
                    note["messages"].append(comment_text)
                
                # Если это последняя строка-источник и перенос не покрыт, допускаем разделение строки
                is_last_source = (idx == len(source_rows) - 1)
                move_qty = min(src_available, remaining, allowed_to_move)
                if move_qty <= 0:
                    continue

                # Уменьшаем у источника
                new_src_val = (src_available - move_qty)
                src_row.el[VALUES].value = new_src_val
                allowed_to_move -= move_qty

                # Поддержка переносов TAGS: переносим целочисленное количество тегов
                tags_available = src_row.get_tags_list()
                tags_to_take = 0
                try:
                    tags_to_take = min(len(tags_available), int(round(move_qty)))
                except Exception:
                    tags_to_take = 0
                moved_tags = []
                if tags_to_take > 0:
                    moved_tags = src_row.pop_tag(tags_to_take)

                # Всегда создаём новый ряд-назначение рядом с источником, чтобы сохранять структуру МТО
                # Если источник обнулился — это будет просто перенос всей строки
                new_row = RowStd.get_row_copy(src_row, t_com=src_row.t_com)
                # Записываем комментарий прямо в CODE, что и на что заменили
                new_row.el[CODE].value = (f"{new_code}"
                                          f"\nЗаменить на:\n"
                                          f"{old_code}")
                new_row.el[VALUES].value = move_qty
                # Задаём теги для новой строки
                if moved_tags:
                    new_row.set_tags_list(moved_tags)
                try:
                    # Вставляем ПЕРЕД источником, чтобы все заменённые строки шли выше остатка нового кода
                    insert_idx = mto_base.index(src_row)
                except ValueError:
                    insert_idx = len(mto_base)
                mto_base.insert(insert_idx, new_row)
                # Подсветка всей новой строки мягким голубым и дополнительно CODE жёлтым
                # for att in ColNames.column_list:
                #     try:
                #         if hasattr(new_row.el.get(att, None), 'color'):
                #             Color.set_el_color(new_row.el[att], Color.soft_cyan)
                #     except Exception:
                #         pass
                Color.set_el_color(new_row.el[CODE], Color.yellow)

                # Если источник обнулился — удаляем строку, чтобы не оставлять нулевые количества
                try:
                    if (src_row.get_value(VALUES) or 0) == 0:
                        if src_row in mto_base:
                            mto_base.remove(src_row)
                except Exception:
                    pass

                # Копим заметку для подсветки источника
                rid = id(src_row)
                note = highlight_notes_by_row_id.get(rid)
                if not note:
                    note = {"row": src_row, "messages": [], "is_direct_demand": False}
                    highlight_notes_by_row_id[rid] = note
                note["messages"].append(f"Заменить на: {old_code} (перенос {move_qty})")

                remaining -= move_qty

            moved = need_qty - max(remaining, 0)
            if moved > 0:
                print(f"{Fore.GREEN}✔ Перенос по {new_code} → {old_code}: {moved}")
            if remaining > 0:
                print(f"{Fore.YELLOW}↪ Недоперенос по {new_code} → {old_code}: {remaining} (не хватило у источника)")

    # Контрольная проверка сохранения общей суммы VALUES
    def _sum_values(rows: list) -> float:
        return sum([(r.get_value(VALUES) or 0.0) for r in rows if check_position_row(r)])

    try:
        # Пересчитываем сумму по модифицированному MTO
        modified_sum = _sum_values(mto_base)
        # Заново загружаем исходный MTO с диска и считаем сумму
        mto_base_check = get_mto_std_from_file(t_com=TableComments(file_full_path=mto_file_path,
                                                                   dir_path="-1",
                                                                   tabel_class=t_comm_initial_classes.MTO),
                                               dbg=0)
        original_sum = _sum_values(mto_base_check)
        print(f"\n{Fore.CYAN}{'='*60}")
        print(f"{Fore.CYAN}🧮 КОНТРОЛЬ СУММЫ ПОСЛЕ ПЕРЕНОСОВ")
        print(f"{Fore.CYAN}{'='*60}")
        print(f"Изначальная сумма: {original_sum:,.2f}")
        print(f"Текущая сумма:     {modified_sum:,.2f}")
        if abs(modified_sum - original_sum) < 1e-6:
            print(f"{Fore.GREEN}✅ Суммы совпадают")
        else:
            diff_total = modified_sum - original_sum
            print(f"{Fore.RED}❌ Суммы НЕ совпадают. Разница: {diff_total:+,.6f}")
    except Exception as e:
        print(f"{Fore.YELLOW}⚠️ Контроль суммы не выполнен: {e}")

    # Подсветка и пометки ПОСЛЕ перераспределения: подсвечиваем только изменённые коды
    if highlight_notes_by_row_id:
        print(f"\n{Fore.CYAN}{'='*60}")
        print(f"{Fore.CYAN}📊 ПОДСВЕТКА СТРОК В МТО ПО ФАКТУ ПЕРЕНОСОВ")
        print(f"{Fore.CYAN}{'='*60}")
        for note in highlight_notes_by_row_id.values():
            row = note["row"]
            msg = "\n".join(note["messages"])
            is_direct_demand = note.get("is_direct_demand", False)
            
            # Подсвечиваем только если реально был перенос из этой строки
            if msg:
                if is_direct_demand:
                    # Красный цвет для строк с прямым спросом
                    Color.set_el_color(row.el[CODE], Color.red)
                else:
                    # Желтый цвет для строк с заменами
                    Color.set_el_color(row.el[CODE], Color.yellow)
                
                ann_val = row.get_value(ANNOTATION) or ""
                sep = "\n" if ann_val else ""
                row.el[ANNOTATION].value = f"{ann_val}{sep}{msg}"
        print(f"{Fore.GREEN}Готово: отмечено {len(highlight_notes_by_row_id)} строк(и)")
    """
    Вывод в EXCEL файл
    """
    # 00 Получаем путь для сохранения результатов работы
    path_out_dir = get_path_out_dir(mto_obj.file_full_path,dir_result_prefix=mto_obj.doc_Short_Title+"_")
    result = STDTable.to_excel(mto_base,
                               ColNames.MTO.column_dict,
                               path_out_dir,
                               "rep_",
                               row_not_print_list=[])
    if result:
        open_dir(result)
        print(f"Успешно завершена проверка: <{mto_obj.file_full_path}>")
    
    # Показываем всплывающее окно с результатами проверок
    show_check_results_popup(
        system_title=str_title_system,
        ds_value=ds_value,
        mto_value=mto_value,
        diff=diff,
        codes_identical=codes_comparison_result['identical'],
        codes_difference=codes_comparison_result['summary']['total_difference'] if not codes_comparison_result['identical'] else 0,
        codes_only_in_ds=codes_comparison_result.get('only_in_first', {}),
        codes_only_in_mto=codes_comparison_result.get('only_in_second', {}),
        codes_different_values=codes_comparison_result.get('different_values', {}),
        transfers_count=len(transfers_by_new_code) if 'transfers_by_new_code' in locals() else 0,
        modified_sum=modified_sum if 'modified_sum' in locals() else 0,
        original_sum=original_sum if 'original_sum' in locals() else 0,
        highlighted_rows=len(highlight_notes_by_row_id) if 'highlight_notes_by_row_id' in locals() else 0
    )


def show_check_results_popup(system_title, ds_value, mto_value, diff, codes_identical, 
                           codes_difference, codes_only_in_ds, codes_only_in_mto, codes_different_values,
                           transfers_count, modified_sum, original_sum, highlighted_rows):
    """
    Показывает всплывающее окно с результатами проверок МТО и DS
    
    Args:
        system_title: Название системы
        ds_value: Сумма по DS
        mto_value: Сумма по МТО
        diff: Разность между DS и МТО
        codes_identical: Совпадают ли суммы по кодам
        codes_difference: Разность по кодам
        codes_only_in_ds: Коды только в ДС словаре
        codes_only_in_mto: Коды только во МТО словаре
        codes_different_values: Коды с разными значениями
        transfers_count: Количество замен
        modified_sum: Сумма после переносов
        original_sum: Изначальная сумма
        highlighted_rows: Количество подсвеченных строк
    """
    # Создаем главное окно результатов
    result_window = ctk.CTkToplevel()
    result_window.title(f"📊 Результаты проверки МТО и DS - {system_title}")
    result_window.geometry("1000x1000")
    result_window.resizable(True, True)
    
    # Делаем окно модальным
    result_window.transient()
    result_window.grab_set()
    
    # Создаем основной фрейм
    main_frame = ctk.CTkFrame(result_window)
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)
    
    # Заголовок
    title_label = ctk.CTkLabel(
        main_frame, 
        text=f"📊 РЕЗУЛЬТАТЫ ПРОВЕРКИ МТО И DS",
        font=("Arial", 18, "bold")
    )
    title_label.pack(pady=(0, 10))
    
    system_label = ctk.CTkLabel(
        main_frame,
        text=f"Система: {system_title}",
        font=("Arial", 14, "bold")
    )
    system_label.pack(pady=(0, 20))
    
    # Создаем фрейм для результатов
    results_frame = ctk.CTkFrame(main_frame)
    results_frame.pack(fill="both", expand=True, pady=(0, 20))
    
    # Основные проверки (акцент)
    main_checks_frame = ctk.CTkFrame(results_frame)
    main_checks_frame.pack(fill="x", padx=10, pady=10)
    
    main_title = ctk.CTkLabel(
        main_checks_frame,
        text="🔍 ОСНОВНЫЕ ПРОВЕРКИ",
        font=("Arial", 16, "bold")
    )
    main_title.pack(pady=(10, 15))
    
    # Сравнение количества по кодам в МТО и DS (главный акцент)
    codes_frame = ctk.CTkFrame(main_checks_frame)
    codes_frame.pack(fill="x", padx=10, pady=5)
    
    codes_title = ctk.CTkLabel(
        codes_frame,
        text="📊 СРАВНЕНИЕ КОЛИЧЕСТВА ПО КОДАМ В МТО И DS",
        font=("Arial", 14, "bold")
    )
    codes_title.pack(pady=(10, 5))
    
    if codes_identical:
        codes_status = ctk.CTkLabel(
            codes_frame,
            text="✅ СУММЫ ПО КОДАМ СОВПАДАЮТ!",
            font=("Arial", 12, "bold"),
            text_color="green"
        )
        codes_status.pack(pady=(0, 10))
    else:
        codes_status = ctk.CTkLabel(
            codes_frame,
            text=f"❌ СУММЫ ПО КОДАМ НЕ СОВПАДАЮТ!\nРазность: {codes_difference:,.2f}",
            font=("Arial", 12, "bold"),
            text_color="red"
        )
        codes_status.pack(pady=(0, 10))
        
        # Создаем текстовое поле для детальной информации о кодах
        codes_details_text = ctk.CTkTextbox(
            codes_frame,
            height=300,
            font=("Consolas", 15)
        )
        codes_details_text.pack(fill="x", padx=10, pady=(0, 10))
        
        # Формируем детальную информацию
        details_text = ""
        
        if codes_only_in_ds:
            details_text += "Коды только в ДС словаре:\n"
            for key, value in codes_only_in_ds.items():
                details_text += f"  {key}: {value}\n"
            details_text += "\n"
        
        if codes_only_in_mto:
            details_text += "Коды только во МТО словаре:\n"
            for key, value in codes_only_in_mto.items():
                details_text += f"  {key}: {value}\n"
            details_text += "\n"
        
        if codes_different_values:
            details_text += "Коды с разными значениями в ДС и МТО:\n"
            for key, diff_info in codes_different_values.items():
                details_text += f"  {key}: {diff_info['first_value']} vs {diff_info['second_value']} (разность: {diff_info['difference']:.6f})\n"
        
        if details_text:
            codes_details_text.insert("1.0", details_text)
            codes_details_text.configure(state="disabled")  # Делаем текст только для чтения
    
    # Перераспределение количеств в МТО по заменам (второй акцент)
    transfers_frame = ctk.CTkFrame(main_checks_frame)
    transfers_frame.pack(fill="x", padx=10, pady=5)
    
    transfers_title = ctk.CTkLabel(
        transfers_frame,
        text="🔁 ПЕРЕРАСПРЕДЕЛЕНИЕ КОЛИЧЕСТВ В МТО ПО ЗАМЕНАМ",
        font=("Arial", 14, "bold")
    )
    transfers_title.pack(pady=(10, 5))
    
    if transfers_count > 0:
        transfers_status = ctk.CTkLabel(
            transfers_frame,
            text=f"✅ Выполнено {transfers_count} замен\nПодсвечено строк: {highlighted_rows}",
            font=("Arial", 12, "bold"),
            text_color="green"
        )
    else:
        transfers_status = ctk.CTkLabel(
            transfers_frame,
            text="ℹ️ Замены не требуются",
            font=("Arial", 12, "bold"),
            text_color="blue"
        )
    transfers_status.pack(pady=(0, 10))
    
    # Дополнительные проверки (второстепенные)
    additional_frame = ctk.CTkFrame(results_frame)
    additional_frame.pack(fill="x", padx=10, pady=10)
    
    additional_title = ctk.CTkLabel(
        additional_frame,
        text="📋 ДОПОЛНИТЕЛЬНЫЕ ПРОВЕРКИ",
        font=("Arial", 14, "bold")
    )
    additional_title.pack(pady=(10, 15))
    
    # Контроль суммы после переносов
    sum_control_frame = ctk.CTkFrame(additional_frame)
    sum_control_frame.pack(fill="x", padx=10, pady=5)
    
    sum_title = ctk.CTkLabel(
        sum_control_frame,
        text="🧮 КОНТРОЛЬ СУММЫ ПОСЛЕ ПЕРЕНОСОВ",
        font=("Arial", 12, "bold")
    )
    sum_title.pack(pady=(5, 5))
    
    sum_diff = abs(modified_sum - original_sum)
    if sum_diff < 1e-6:
        sum_status = ctk.CTkLabel(
            sum_control_frame,
            text=f"✅ Суммы совпадают\nИзначальная: {original_sum:,.2f}\nТекущая: {modified_sum:,.2f}",
            font=("Arial", 11),
            text_color="green"
        )
    else:
        sum_status = ctk.CTkLabel(
            sum_control_frame,
            text=f"❌ Суммы НЕ совпадают\nРазница: {sum_diff:+,.6f}",
            font=("Arial", 11),
            text_color="red"
        )
    sum_status.pack(pady=(0, 5))
    
    # Сравнение количества в МТО и DS
    comparison_frame = ctk.CTkFrame(additional_frame)
    comparison_frame.pack(fill="x", padx=10, pady=5)
    
    comparison_title = ctk.CTkLabel(
        comparison_frame,
        text="📈 СРАВНЕНИЕ КОЛИЧЕСТВА В МТО И DS",
        font=("Arial", 12, "bold")
    )
    comparison_title.pack(pady=(5, 5))
    
    if diff == 0:
        comparison_status = ctk.CTkLabel(
            comparison_frame,
            text=f"✅ СУММЫ СОВПАДАЮТ!\nDS: {ds_value:,.2f}\nМТО: {mto_value:,.2f}",
            font=("Arial", 11),
            text_color="green"
        )
    else:
        comparison_status = ctk.CTkLabel(
            comparison_frame,
            text=f"❌ СУММЫ НЕ СОВПАДАЮТ!\nDS: {ds_value:,.2f}\nМТО: {mto_value:,.2f}\nРазность: {diff:,.2f}",
            font=("Arial", 11),
            text_color="red"
        )
    comparison_status.pack(pady=(0, 5))
    
    # Проверка МТО на теги
    tags_frame = ctk.CTkFrame(additional_frame)
    tags_frame.pack(fill="x", padx=10, pady=5)
    
    tags_title = ctk.CTkLabel(
        tags_frame,
        text="🏷️ ПРОВЕРКА МТО НА ТЕГИ",
        font=("Arial", 12, "bold")
    )
    tags_title.pack(pady=(5, 5))
    
    tags_status = ctk.CTkLabel(
        tags_frame,
        text="✅ ПРОВЕРКА МТО НА ТЕГИ ЗАПУЩЕНА",
        font=("Arial", 11),
        text_color="green"
    )
    tags_status.pack(pady=(0, 5))
    
    # Кнопка закрытия
    close_button = ctk.CTkButton(
        main_frame,
        text="Закрыть",
        command=result_window.destroy,
        width=120,
        height=40,
        font=("Arial", 12, "bold")
    )
    close_button.pack(pady=(10, 0))
    
    # Центрируем окно и адаптируем размер
    result_window.update_idletasks()
    
    # Получаем размеры экрана
    screen_width = result_window.winfo_screenwidth()
    screen_height = result_window.winfo_screenheight()
    
    # Получаем размеры окна после размещения всех элементов
    window_width = result_window.winfo_reqwidth()
    window_height = result_window.winfo_reqheight()
    
    # Ограничиваем размер окна размером экрана (с отступами)
    max_width = min(1000, screen_width - 100)
    max_height = min(1000, screen_height - 100)
    
    # Устанавливаем оптимальный размер
    optimal_width = min(window_width + 50, max_width)
    optimal_height = min(window_height + 50, max_height)
    
    result_window.geometry(f"{optimal_width}x{optimal_height}")
    
    # Центрируем окно
    x = (screen_width - optimal_width) // 2
    y = (screen_height - optimal_height) // 2
    result_window.geometry(f"{optimal_width}x{optimal_height}+{x}+{y}")
    
    # Выводим окно на передний план
    result_window.lift()  # Поднимаем окно поверх других
    result_window.attributes('-topmost', True)  # Делаем окно всегда поверх всех
    result_window.focus_force()  # Принудительно устанавливаем фокус
    
    # Убираем topmost после небольшой задержки, чтобы окно осталось поверх, но не блокировало другие окна
    result_window.after(100, lambda: result_window.attributes('-topmost', False))
