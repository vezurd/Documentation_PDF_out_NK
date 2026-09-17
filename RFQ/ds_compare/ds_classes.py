from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Dict, List, Optional, TextIO

from base.base_excel_out import *
from pdf_parsing_v2_engine.document import V2Document
from RFQ.ds_compare.ds_in_cabinet_alloc import (
    CabinetAllocationResult,
    allocate_mto_slices,
    cabinet_alloc_from_used_rows,
    norm_in_cabinet,
)
from RFQ.ds_compare.ds_in_cabinet_trace import CabinetInCabinetTrace


class ComparisonResult(Enum):
    MATCH = "match"  # Количества совпадают
    INSUFFICIENT = "insufficient"  # В MTO меньше чем в DS
    EXCESS = "excess"  # В MTO больше чем в DS
    NOT_FOUND = "not_found"  # Не найдено в MTO
    NEW_POSITION = "new_position"  # Новая позиция в MTO
    VENDOR_MISMATCH = "vendor_mismatch"  # Несовпадение вендоров
    UNIT_MISMATCH = "unit_mismatch"  # Несовпадение единиц измерения


class ComparisonStatus(Enum):
    MATCH = "Совпадает кол-во"  # Количества совпадаютf
    INSUFFICIENT = "Уменьшилось кол-во"  # В MTO меньше чем в DS
    EXCESS = "Увеличилось кол-во"  # В MTO больше чем в DS
    NOT_FOUND = "Нет в МТО"  # Не найдено в MTO
    NEW_POSITION = "Новая позиция"  # Новая позиция в MTO
    VENDOR_MISMATCH = "vendor_mismatch"  # Несовпадение вендоров
    UNIT_MISMATCH = "unit_mismatch"  # Несовпадение единиц измерения
    MTO_LESS_DS_BY_ROW = "Уменьшилось кол-во (не хватило строк в МТО)"


@dataclass
class ComparisonRule:
    """Базовый класс для правил сравнения"""
    priority: int = 0

    def apply(self, ds_row: RowStd, mto_rows: List[RowStd]) -> Optional[ComparisonResult]:
        raise NotImplementedError


@dataclass
class QuantityComparisonRule(ComparisonRule):
    """Правило сравнения количеств"""
    tolerance: float = 0.0  # Допустимое отклонение

    def apply(self, ds_row: RowStd, mto_rows: List[RowStd]) -> Optional[ComparisonResult]:
        try:
            ds_amount = float(ds_row.get_value(VALUES) or 0)
            total_mto = sum(float(row.get_value(VALUES) or 0) for row in mto_rows)

            if abs(total_mto - ds_amount) <= self.tolerance:
                return ComparisonResult.MATCH
            elif total_mto < ds_amount:
                return ComparisonResult.INSUFFICIENT
            else:
                return ComparisonResult.EXCESS
        except (ValueError, TypeError):
            return ComparisonResult.NOT_FOUND




class DsMtoComparator:
    """

    """

    def __init__(self):
        self.debug = False
        self.debug_code = "EPAPL000005"  # Код для отладочного вывода
        self.replacement_table = {}
        self.debug_log_path: str | Path | None = None
        self._debug_log_fp: TextIO | None = None
        self.cabinet_trace: CabinetInCabinetTrace | None = None

    def _debug_out(self, *args: object, sep: str = " ", end: str = "\n", flush: bool = True) -> None:
        """Print to stdout and mirror to the debug log file while ``compare_specifications`` runs."""
        print(*args, sep=sep, end=end, flush=flush)
        fp = self._debug_log_fp
        if fp is not None:
            fp.write(sep.join(str(a) for a in args))
            fp.write(end)
            if flush:
                fp.flush()

    def _debug_values_2_write(self, ds_row: RowStd, old_value, new_value, operation: str, code: str = None):
        """Отладочный вывод для записи в VALUES_2"""
        if not code:
            code = self.debug_code

        if code and code == self.debug_code:
            current_ds_amount = float(ds_row.get_value(VALUES) or 0)
            current_mto_amount = float(ds_row.get_value(VALUES_2) or 0)
            self._debug_out(f"[VALUES_2_TRACE] {operation}: {old_value} -> {new_value} | "
                            f"DS: {current_ds_amount}, MTO: {current_mto_amount} | "
                            f"Row: {self._get_row_debug_info(ds_row)}")

    def _debug_print(self, message: str, code: str = None):
        """Вывод отладочной информации для заданного кода"""
        if not code:
            code = self.debug_code
        if code and code == self.debug_code:
            self._debug_out(f"[DEBUG {code}] {message}")

    def compare_specifications(self,
                               ds_base_full: List[RowStd],
                               mto_dict: Dict[str, List[RowStd]],
                               replacement_table) -> None:
        """Основной метод сравнения с последовательным распределением"""
        log_fp: TextIO | None = None
        try:
            if self.debug_log_path:
                log_path = Path(self.debug_log_path)
                log_path.parent.mkdir(parents=True, exist_ok=True)
                log_fp = open(log_path, "w", encoding="utf-8", newline="\n")
                self._debug_log_fp = log_fp
                self._debug_out(f"--- DsMtoComparator debug log: {log_path.resolve()} ---")

            self.replacement_table = replacement_table
            ds_grouped = self._group_ds_rows(ds_base_full)

            self._debug_print(f"Начало сравнения. Групп DS: {len(ds_grouped)}", "")

            # Обрабатываем каждую спецификацию
            for spec_name, codes_dict in ds_grouped.items():
                if spec_name not in mto_dict:
                    self._debug_print(f"Спецификация {spec_name} не найдена в MTO", "")
                    self._handle_missing_specification(codes_dict, spec_name)
                    continue

                mto_base = mto_dict[spec_name]
                self._debug_print(f"Обработка спецификации {spec_name}, кодов: {len(codes_dict)}", "")
                self._process_specification(spec_name, codes_dict, mto_base)

            # ПОСТ-ПРОХОД: Обрабатываем новые строки MTO с совпадающими кодами
            self._process_matching_new_positions(ds_base_full, mto_dict, ds_grouped)

            # Добавляем оставшиеся новые позиции из MTO
            self._add_new_positions(ds_base_full, mto_dict, ds_grouped)
        finally:
            self._debug_log_fp = None
            if log_fp is not None:
                log_fp.close()

    def _process_matching_new_positions(self, ds_base_full: List[RowStd], mto_dict: Dict[str, List[RowStd]], ds_grouped: Dict[str, Dict[str, List[RowStd]]]):
        """Пост-проход: обрабатывает новые строки MTO с совпадающими кодами"""
        if self.debug:
            self._debug_out(f"\n[POST-PASS] Обработка новых строк MTO с совпадающими кодами...")
        
        # Создаем группировку DS по CODE_2 для этой функции
        ds_grouped_by_code_2 = self._group_ds_rows_by_code_2(ds_base_full)
        
        for spec_name, mto_base in mto_dict.items():
            if spec_name not in ds_grouped_by_code_2:
                continue
                
            if self.debug:
                self._debug_out(f"[POST-PASS] Обработка спецификации {spec_name}")

            # Получаем все коды из DS для этой спецификации (по CODE_2), исключая пустые
            ds_codes_2 = set()
            for code_2 in ds_grouped_by_code_2[spec_name].keys():
                if code_2 and str(code_2).strip():  # Проверяем, что код не пустой и не состоит из пробелов
                    ds_codes_2.add(code_2)

            if self.debug:
                self._debug_out(
                    f"[POST-PASS] Обработка спецификации {spec_name}, найдено {len(ds_codes_2)} непустых кодов CODE_2")
                if len(ds_codes_2) < len(ds_grouped_by_code_2[spec_name].keys()):
                    empty_count = len(ds_grouped_by_code_2[spec_name].keys()) - len(ds_codes_2)
                    self._debug_out(f"[POST-PASS] Игнорировано {empty_count} пустых кодов")
            
            # Ищем новые строки MTO с совпадающими кодами
            matching_new_positions = []
            for mto_row in mto_base:
                if (isinstance(mto_row, RowStd) and 
                    mto_row.row_type == RowType.position_row):
                    
                    mto_code = mto_row.get_value(CODE)
                    mto_values = float(mto_row.get_value(VALUES) or 0)
                    
                    if mto_values > 0 and mto_code in ds_codes_2:
                        if mto_code == self.debug_code:
                            self._debug_print(f"[POST-PASS] Найдена новая строка с совпадающим кодом: {self._get_row_debug_info(mto_row)}", mto_code)
                        matching_new_positions.append(mto_row)
            
            if matching_new_positions:
                if self.debug:
                    self._debug_out(f"[POST-PASS] Найдено {len(matching_new_positions)} новых строк с совпадающими кодами")
                    for i, position in enumerate(matching_new_positions):
                        mto_code = position.get_value(CODE)
                        mto_code_2 = position.get_value(CODE_2)
                        mto_values = float(position.get_value(VALUES) or 0)
                        mto_tags = position.get_tags_list()
                        mto_number = position.get_value(NUMBERS) or "нет"

                        self._debug_out(f"  Строка #{i + 1}:")
                        self._debug_out(f"    CODE: {mto_code}")
                        self._debug_out(f"    CODE_2: {mto_code_2}")
                        self._debug_out(f"    VALUES: {mto_values}")
                        self._debug_out(f"    NUMBERS: {mto_number}")
                        self._debug_out(f"    TAGS: {mto_tags}")
                        self._debug_out(f"    Спецификация: {spec_name}")

                        # Проверяем соответствие с DS строками
                        if mto_code in ds_grouped_by_code_2[spec_name]:
                            ds_rows_for_code = ds_grouped_by_code_2[spec_name][mto_code]
                            self._debug_out(f"    Соответствующие DS строки: {len(ds_rows_for_code)}")
                            for j, ds_row in enumerate(ds_rows_for_code):
                                ds_values = float(ds_row.get_value(VALUES) or 0)
                                ds_values_2 = float(ds_row.get_value(VALUES_2) or 0)
                                ds_number = ds_row.get_value(NUMBERS) or ds_row.get_value(DS_NUMBER) or "нет"
                                self._debug_out(
                                    f"      DS #{j + 1}: NUMBERS={ds_number}, VALUES={ds_values}, VALUES_2={ds_values_2}")
                        else:
                            self._debug_out(f"    Соответствующие DS строки: не найдено")

                        self._debug_out("")  # Пустая строка для разделения
                
                # Обрабатываем совпадающие новые позиции
                self._distribute_matching_new_positions(matching_new_positions, ds_grouped_by_code_2[spec_name], spec_name)
    
    def _distribute_matching_new_positions(self,
                                           matching_positions: List[RowStd],
                                           ds_codes_dict: Dict[str, List[RowStd]],
                                           spec_name: str):
        """Распределяет новые позиции MTO с совпадающими кодами по DS строкам"""
        if self.debug:
            self._debug_out(f"[POST-PASS] Распределение {len(matching_positions)} совпадающих позиций")
        
        # Группируем новые позиции по кодам (CODE из MTO)
        positions_by_code = {}
        for position in matching_positions:
            code = position.get_value(CODE)
            if code not in positions_by_code:
                positions_by_code[code] = []
            positions_by_code[code].append(position)
        
        # Обрабатываем каждый код
        for code, positions in positions_by_code.items():
            if code not in ds_codes_dict:
                continue
                
            ds_rows = ds_codes_dict[code]
            if not ds_rows:
                continue
            
            if code == self.debug_code:
                self._debug_print(f"[POST-PASS] Обработка кода {code}: {len(positions)} новых позиций", code)
            
            # Сортируем новые позиции по приоритету
            sorted_positions = self._sort_new_positions_by_priority(positions, ds_rows[0])
            
            # Распределяем по DS строкам в порядке приоритета
            self._distribute_to_ds_rows_by_code_2(sorted_positions, ds_rows, spec_name, code)
    
    def _sort_new_positions_by_priority(self, positions: List[RowStd], ds_row: RowStd) -> List[RowStd]:
        """Сортирует новые позиции по приоритету (точное совпадение, затем по таблице замен)"""
        ds_code_2 = ds_row.get_value(CODE_2)
        ds_tag = ds_row.get_value(TAGS)
        ds_spec = self._get_ds_specification_context(ds_row)
        
        def get_priority(position):
            pos_code = position.get_value(CODE)
            pos_tag = position.get_value(TAGS)
            pos_spec = self._get_mto_specification_context(position)
            
            # Приоритет 0: Точное совпадение кода с CODE_2
            if pos_code == ds_code_2:
                return 0
            # Приоритет 1: Совпадение по таблице замен
            elif self._is_replacement_code(pos_code, ds_code_2):
                return 1
            # Приоритет 2: Совпадение TAG
            elif pos_tag == ds_tag:
                return 2
            # Приоритет 3: Совпадение спецификации
            elif pos_spec == ds_spec:
                return 3
            # Приоритет 4: Все остальные
            else:
                return 4
        
        return sorted(positions, key=lambda x: (get_priority(x), float(x.get_value(VALUES) or 0)))
    
    def _is_replacement_code(self, mto_code: str, ds_code: str) -> bool:
        """Проверяет, является ли MTO код заменой для DS кода"""
        if ds_code in self.replacement_table:
            for new_code, status in self.replacement_table[ds_code]:
                if new_code == mto_code:
                    return True
        return False
    
    def _distribute_to_ds_rows_by_code_2(self,
                                         sorted_positions: List[RowStd],
                                         ds_rows: List[RowStd],
                                         spec_name: str,
                                         code: str):
        """Распределяет отсортированные позиции по DS строкам (работает с группировкой по CODE_2)"""
        # Группируем позиции по кодам для суммирования
        positions_by_code = {}
        for position in sorted_positions:
            mto_code = position.get_value(CODE)
            if mto_code not in positions_by_code:
                positions_by_code[mto_code] = []
            positions_by_code[mto_code].append(position)
        
        # Обрабатываем каждый код
        for mto_code, positions in positions_by_code.items():
            # Суммируем все количества для этого кода
            total_amount = sum(float(pos.get_value(VALUES) or 0) for pos in positions)
            
            if total_amount <= 0:
                continue
                
            # Ищем последнюю DS строку с совпадающим CODE_2
            target_ds_row = None
            for ds_row in reversed(ds_rows):  # Ищем с конца списка
                ds_code_2 = ds_row.get_value(CODE_2)
                if ds_code_2 == mto_code:
                    target_ds_row = ds_row
                    break
            
            if not target_ds_row:
                if code == self.debug_code:
                    self._debug_print(f"[POST-PASS] Не найдена DS строка с CODE_2={mto_code}", code)
                continue
            
            if code == self.debug_code:
                self._debug_print(f"[POST-PASS] Распределение в ПОСЛЕДНЮЮ DS строку с CODE_2={mto_code}: "
                                  f"{self._get_row_debug_info(target_ds_row)}", code)
            
            # Обновляем количество в MTO части DS строки (VALUES_2)
            current_mto_amount = float(target_ds_row.get_value(VALUES_2) or 0)
            new_mto_amount = current_mto_amount + total_amount

            # ОТЛАДОЧНЫЙ ВЫВОД
            self._debug_values_2_write(target_ds_row, current_mto_amount, new_mto_amount,
                                       f"POST_PASS_ADD to {mto_code}", code)

            target_ds_row.el[VALUES_2].value = new_mto_amount
            
            # Обновляем комментарий об агрегированных кодах, если он есть
            self._update_aggregated_comment_on_addition(target_ds_row, total_amount, code)
            
            # Обновляем статус строки
            self._update_ds_row_status_after_addition(target_ds_row, total_amount, code)
            
            if code == self.debug_code:
                self._debug_print(f"[POST-PASS] Добавлено {total_amount} к MTO части ПОСЛЕДНЕЙ DS строки. Новое MTO количество: {new_mto_amount}", code)
            
            # Обнуляем количество у всех использованных позиций
            for position in positions:
                position.el[VALUES].value = 0  # Обнуляем количество

    def _distribute_to_ds_rows(self, sorted_positions: List[RowStd], ds_rows: List[RowStd], spec_name: str, code: str):
        """Распределяет отсортированные позиции по DS строкам (старая версия для совместимости)"""
        # Группируем позиции по CODE_2 (из MTO) для правильного распределения
        positions_by_code_2 = {}
        for position in sorted_positions:
            mto_code_2 = position.get_value(CODE_2)
            if mto_code_2 not in positions_by_code_2:
                positions_by_code_2[mto_code_2] = []
            positions_by_code_2[mto_code_2].append(position)
        
        # Распределяем по DS строкам с совпадающим CODE_2
        for mto_code_2, positions in positions_by_code_2.items():
            # Ищем DS строку с совпадающим CODE_2
            target_ds_row = None
            for ds_row in ds_rows:
                ds_code_2 = ds_row.get_value(CODE_2)
                if ds_code_2 == mto_code_2:
                    target_ds_row = ds_row
                    break
            
            if not target_ds_row:
                if code == self.debug_code:
                    self._debug_print(f"[POST-PASS] Не найдена DS строка с CODE_2={mto_code_2}", code)
                continue
            
            if code == self.debug_code:
                self._debug_print(f"[POST-PASS] Распределение в DS строку с CODE_2={mto_code_2}: {self._get_row_debug_info(target_ds_row)}", code)
            
            # Суммируем все количества из новых позиций с этим CODE_2
            total_additional_amount = sum(float(pos.get_value(VALUES) or 0) for pos in positions)
            
            if total_additional_amount > 0:
                # Обновляем количество в MTO части DS строки (VALUES_2)
                current_mto_amount = float(target_ds_row.get_value(VALUES_2) or 0)
                new_mto_amount = current_mto_amount + total_additional_amount

                # ОТЛАДОЧНЫЙ ВЫВОД
                self._debug_values_2_write(target_ds_row, current_mto_amount, new_mto_amount,
                                           f"POST_PASS_ADD_OLD to {mto_code_2}", code)

                target_ds_row.el[VALUES_2].value = new_mto_amount
                
                # Обновляем комментарий об агрегированных кодах, если он есть
                self._update_aggregated_comment_on_addition(target_ds_row, total_additional_amount, code)
                
                # Обновляем статус строки
                self._update_ds_row_status_after_addition(target_ds_row, total_additional_amount, code)
                
                if code == self.debug_code:
                    self._debug_print(f"[POST-PASS] Добавлено {total_additional_amount} к MTO части DS строки. Новое MTO количество: {new_mto_amount}", code)
                
                # Обнуляем количество у использованных позиций
                for position in positions:
                    position.el[VALUES].value = 0  # Обнуляем количество
    
    def _update_ds_row_status_after_addition(self, ds_row: RowStd, added_amount: float, code: str):
        """Обновляет статус DS строки после добавления количества"""
        # Добавляем комментарий о добавлении количества
        comment = ds_row.el[VALUES].comment or ""
        addition_comment = f"[POST_PASS_ADDED: {added_amount}]"
        if addition_comment not in comment:
            ds_row.el[VALUES].comment = comment + addition_comment
        
        # Определяем текущее состояние строки (работаем с MTO частью)
        current_mto_amount = float(ds_row.get_value(VALUES_2) or 0)
        original_mto_amount = current_mto_amount - added_amount
        ds_amount = float(ds_row.get_value(VALUES) or 0)  # DS количество не меняется
        
        # Используем стандартные обработчики в зависимости от ситуации
        # Сравниваем MTO количество с DS количеством
        if "[PRELIMINARY_PROCESSED]" in comment:
            # Если была полностью обработана, теперь избыток
            self._handle_excess(ds_row, current_mto_amount, ds_amount, code)
        elif "[PRELIMINARY_PARTIAL]" in comment:
            # Если была частично обработана, проверяем результат
            if current_mto_amount == ds_amount:
                self._handle_match(ds_row, current_mto_amount, ds_amount, code)
            elif current_mto_amount > ds_amount:
                self._handle_excess(ds_row, current_mto_amount, ds_amount, code)
            else:
                self._handle_insufficient(ds_row, current_mto_amount, ds_amount, code)
        else:
            # Обычная строка - избыток
            self._handle_excess(ds_row, current_mto_amount, ds_amount, code)

    def _group_ds_rows(self, ds_base_full: List[RowStd]) -> Dict[str, Dict[str, List[RowStd]]]:
        """Группирует строки DS по спецификациям и кодам"""
        grouped = {}
        for ds_row in ds_base_full:
            if isinstance(ds_row, RowStd) and ds_row.row_type == RowType.position_row:
                spec_name = V2Document.from_file_path(ds_row.get_value(DS_SPECIFICATION)).doc_Short_Title
                code = ds_row.get_value(CODE)

                if code == self.debug_code:
                    self._debug_print(f"Найдена DS строка: {self._get_row_debug_info(ds_row)}", code)

                if spec_name not in grouped:
                    grouped[spec_name] = {}
                if code not in grouped[spec_name]:
                    grouped[spec_name][code] = []

                grouped[spec_name][code].append(ds_row)

        return grouped

    def _group_ds_rows_by_code_2(self, ds_base_full: List[RowStd]) -> Dict[str, Dict[str, List[RowStd]]]:
        """Группирует строки DS по спецификациям и кодам CODE_2"""
        grouped = {}
        for ds_row in ds_base_full:
            if isinstance(ds_row, RowStd) and ds_row.row_type == RowType.position_row:
                spec_name = V2Document.from_file_path(ds_row.get_value(DS_SPECIFICATION)).doc_Short_Title
                code_2 = ds_row.get_value(CODE_2)

                if code_2 == self.debug_code:
                    self._debug_print(f"Найдена DS строка по CODE_2: {self._get_row_debug_info(ds_row)}", code_2)

                if spec_name not in grouped:
                    grouped[spec_name] = {}
                if code_2 not in grouped[spec_name]:
                    grouped[spec_name][code_2] = []

                grouped[spec_name][code_2].append(ds_row)

        return grouped

    def _get_row_debug_info(self, row: RowStd) -> str:
        """Возвращает отладочную информацию о строке"""
        return (f"CODE: {row.get_value(CODE)}, "                
                f"VALUES: {row.get_value(VALUES)}, "
                f"IN_CABINET: {row.get_value(IN_CABINET)}, "
                f"Ном.строки: {row.get_value(NUMBERS) or row.get_value(DS_NUMBER)}, "
                f"TAGS: {row.get_value(TAGS)}")

    def _handle_missing_specification(self, codes_dict: Dict[str, List[RowStd]],
                                      spec_name: str):
        """Обрабатывает отсутствующую спецификацию в MTO"""
        for code, ds_rows in codes_dict.items():
            for ds_row in ds_rows:
                if code == self.debug_code:
                    self._debug_print(f"Спецификация не найдена: {spec_name}", code)

                Color.set_el_color(ds_row.el[VALUES], Color.yellow)
                Color.set_el_color(ds_row.el[ANNOTATION_2], Color.yellow)
                comment_text = "Спецификация не найдена\n"
                ds_row.el[ANNOTATION_2].value = ds_row.el[ANNOTATION_2].value + comment_text
                comment_text = f"Спецификация {spec_name} не найдена в MTO\n"
                ds_row.el[VALUES].comment = ds_row.el[VALUES].comment + comment_text



    def _process_specification(self,
                               spec_name: str,
                               codes_dict: Dict[str, List[RowStd]],
                               mto_base: List[RowStd]):
        """Обрабатывает одну спецификацию с учетом таблицы замен"""
        
        # ПРЕДВАРИТЕЛЬНЫЙ ПРОХОД: Обрабатываем точные совпадения кодов
        self._process_exact_code_matches(spec_name, codes_dict, mto_base)
        # print("Очищаем строки в МТО с нулевым кол-вом.")
        # mto_base = [mto_row for mto_row in mto_base if mto_row.get_value(VALUES) > 0]

        # ОСНОВНОЙ ПРОХОД: Обрабатываем все коды с заменами (включая частично обработанные)
        for code, ds_rows in codes_dict.items():
                
            if code == self.debug_code:
                self._debug_print(f"Обработка кода, DS строк: {len(ds_rows)}", code)

            # Ищем прямые совпадения
            mto_rows = RowStd.get_row_list_by_code(code, mto_base)

            if code == self.debug_code and mto_rows:
                for mto_row in mto_rows:
                    self._debug_print(f"Найдена MTO строка: {self._get_row_debug_info(mto_row)}", code)

            # Ищем замены через таблицу
            replacement_mto_rows = self._find_replacement_mto_rows(code, mto_base)

            if code == self.debug_code and replacement_mto_rows:
                for mto_row in replacement_mto_rows:
                    self._debug_print(f"Найдена замена MTO: {self._get_row_debug_info(mto_row)}", code)

            # Объединяем результаты
            all_mto_rows = mto_rows + replacement_mto_rows

            if not all_mto_rows:
                if code == self.debug_code:
                    self._debug_print("MTO строки не найдены", code)
                self._handle_code_not_found(ds_rows, spec_name, code)
                continue

            if code == self.debug_code:
                self._debug_print(f"Всего MTO строк: {len(all_mto_rows)}", code)

            # Последовательное распределение количеств
            self._distribute_amounts_sequentially(ds_rows, all_mto_rows, spec_name, code, codes_dict)

    def _process_exact_code_matches(self, spec_name: str, codes_dict: Dict[str, List[RowStd]], mto_base: List[RowStd]):
        """Предварительный проход: обрабатывает точные совпадения кодов"""
        if self.debug:
            self._debug_out(f"[PRELIMINARY] Обработка точных совпадений для спецификации {spec_name}")
        
        for code, ds_rows in codes_dict.items():
            if code == self.debug_code:
                self._debug_print(f"[PRELIMINARY] Обработка точного совпадения кода {code}", code)
            
            # Ищем точные совпадения кодов
            exact_mto_rows = RowStd.get_row_list_by_code(code, mto_base)
            
            if not exact_mto_rows:
                if code == self.debug_code:
                    self._debug_print("[PRELIMINARY] Точных совпадений не найдено", code)
                continue
                
            if code == self.debug_code:
                self._debug_print(f"[PRELIMINARY] Найдено точных совпадений: {len(exact_mto_rows)}", code)
                for i, mto_row in enumerate(exact_mto_rows):
                    self._debug_print(f"[PRELIMINARY] MTO #{i}: {self._get_row_debug_info(mto_row)}", code)
            
            # Обрабатываем точные совпадения с ограничением по количеству
            self._distribute_exact_matches(ds_rows, exact_mto_rows, spec_name, code)

    def _distribute_exact_matches(self, ds_rows: List[RowStd], mto_rows: List[RowStd], spec_name: str, code: str):
        """Распределяет точные совпадения кодов, используя только необходимое количество"""
        if code == self.debug_code:
            self._debug_print(f"[PRELIMINARY] Распределение точных совпадений для кода {code}", code)

        sorted_mto_rows = list(mto_rows)
        
        # Суммируем общее требуемое количество в DS
        total_ds_amount = sum(float(ds_row.get_value(VALUES) or 0) for ds_row in ds_rows)
        
        if code == self.debug_code:
            self._debug_print(f"[PRELIMINARY] DS требует: {total_ds_amount}", code)
            for i, mto_row in enumerate(sorted_mto_rows):
                mto_amount = float(mto_row.get_value(VALUES) or 0)
                self._debug_print(f"[PRELIMINARY] MTO #{i}: amount={mto_amount}", code)
        
        # Распределяем только необходимое количество
        remaining_ds_amount = total_ds_amount
        
        for ds_row in ds_rows:
            if remaining_ds_amount <= 0:
                break
                
            ds_amount = float(ds_row.get_value(VALUES) or 0)
            if ds_amount <= 0:
                continue
                
            if code == self.debug_code:
                self._debug_print(f"[PRELIMINARY] Обработка DS строки: требуется {ds_amount} "
                                  f"({self._get_row_debug_info(ds_row)})", code)
            
            cabinet_alloc = allocate_mto_slices(sorted_mto_rows, ds_amount, ds_row)
            if cabinet_alloc.total_taken >= ds_amount and cabinet_alloc.primary_row is not None:
                allocated_amount = ds_amount
                mto_row_used = cabinet_alloc.primary_row
                if code == self.debug_code:
                    self._debug_print(
                        f"[PRELIMINARY] Выделено {allocated_amount} из {len(cabinet_alloc.slices)} MTO строк",
                        code,
                    )
                self._consume_allocation_slices(cabinet_alloc, code)
                self._process_allocation_result(
                    ds_row, mto_row_used, ds_amount, allocated_amount, code,
                    cabinet_alloc=cabinet_alloc,
                )
                comment = ds_row.el[VALUES].comment or ""
                if "[PRELIMINARY_PROCESSED]" not in comment:
                    ds_row.el[VALUES].comment = comment + "[PRELIMINARY_PROCESSED]"
                remaining_ds_amount -= allocated_amount
                self._remove_exhausted_mto_rows(sorted_mto_rows, code)
            else:
                if code == self.debug_code:
                    self._debug_print(f"[PRELIMINARY] Не найдена MTO строка с достаточным количеством для {ds_amount}", code)
                # Помечаем как частично обработанную для дальнейшей обработки
                comment = ds_row.el[VALUES].comment or ""
                if "[PRELIMINARY_PARTIAL]" not in comment:
                    ds_row.el[VALUES].comment = comment + "[PRELIMINARY_PARTIAL]"

    def _is_code_processed(self, ds_rows: List[RowStd]) -> bool:
        """Проверяет, был ли код уже обработан в предварительном проходе"""
        # Проверяем, есть ли у DS строк комментарии о том, что они обработаны
        for ds_row in ds_rows:
            comment = ds_row.el[VALUES].comment or ""
            if "[PRELIMINARY_PROCESSED]" in comment:
                return True
        return False

    def _find_replacement_mto_rows(self, code: str, mto_base: List[RowStd]) -> List[RowStd]:
        """
        Ищет строки MTO через таблицу замен
        """
        replacement_rows = []

        if code in self.replacement_table.keys():
            for new_code, status in self.replacement_table[code]:
                found_rows = RowStd.get_row_list_by_code(new_code, mto_base)
                for row in found_rows:
                    # Добавляем информацию о замене
                    row.replacement_info = {
                        'original_code': code,
                        'new_code': new_code,
                        'status': status
                    }
                    replacement_rows.append(row)

                    if code == self.debug_code or new_code == self.debug_code:
                        self._debug_print(f"Добавлена замена: {new_code} -> {code}, статус: {status}", code)

        return replacement_rows

    def _has_new_codes_in_specification(self, old_code: str, codes_dict: Dict[str, List[RowStd]]) -> bool:
        """
        Проверяет, есть ли в спецификации строки с новыми кодами из таблицы замен
        """
        if old_code not in self.replacement_table:
            return False
        
        # Получаем все новые коды для данного старого кода
        new_codes = [new_code for new_code, _ in self.replacement_table[old_code]]
        
        # Проверяем, есть ли в спецификации строки с новыми кодами
        for new_code in new_codes:
            if new_code in codes_dict:
                return True
        
        return False

    def _distribute_amounts_sequentially(self,
                                         ds_rows: List[RowStd],
                                         mto_rows: List[RowStd],
                                         spec_name: str,
                                         code: str,
                                         codes_dict: Dict[str, List[RowStd]],
                                         is_preliminary: bool = False):
        """
        Распределяет количества последовательно по строкам DS с обновлением MTO
        Именно в этой функции определяется куда распределиться кол-во.
        """

        # Суммируем общее количество в MTO
        total_mto_amount = sum(float(row.get_value(VALUES) or 0) for row in mto_rows)

        # Суммируем общее требуемое количество в DS
        total_ds_amount = sum(float(ds_row.get_value(VALUES) or 0) for ds_row in ds_rows)

        if code == self.debug_code:
            self._debug_print(f"Распределение: DS требует {total_ds_amount}, MTO имеет {total_mto_amount}", code)
            for i, mto_row in enumerate(mto_rows):
                self._debug_print(f"MTO #{i}: {self._get_row_debug_info(mto_row)}", code)

        # Последовательно распределяем MTO по строкам DS
        for i, ds_row in enumerate(ds_rows):
            # Проверяем, была ли строка уже обработана в предварительном проходе
            comment = ds_row.el[VALUES].comment or ""
            is_preliminary_processed = "[PRELIMINARY_PROCESSED]" in comment
            is_preliminary_partial = "[PRELIMINARY_PARTIAL]" in comment
            
            if is_preliminary_processed and not is_preliminary:
                if code == self.debug_code:
                    self._debug_print(f"DS строка #{i} уже полностью обработана в предварительном проходе, пропускаем", code)
                continue
                
            if is_preliminary_partial and not is_preliminary:
                if code == self.debug_code:
                    self._debug_print(f"DS строка #{i} частично обработана в предварительном проходе, продолжаем обработку", code)
                
            ds_amount = float(ds_row.get_value(VALUES) or 0)

            if code == self.debug_code:
                self._debug_print(f"Обработка DS строки #{i}: {self._get_row_debug_info(ds_row)}", code)

            if not mto_rows:  # MTO закончились
                if code == self.debug_code:
                    self._debug_print("MTO закончились", code)
                self._handle_insufficient_mto(ds_row, ds_amount, 0, spec_name, code)
                continue

            if i != len(ds_rows) - 1:
                mto_row, cabinet_alloc = self._allocate_ds_row(ds_row, mto_rows, ds_amount, code)
                if not mto_row:
                    if code == self.debug_code:
                        self._debug_print("Не найдена подходящая MTO строка", code)
                    self._handle_insufficient_mto(ds_row, ds_amount, 0, spec_name, code)
                    continue
                allocated_amount = min(ds_amount, cabinet_alloc.total_taken)
                if code == self.debug_code:
                    self._debug_print(
                        f"Строка #{i} - берем: {allocated_amount} "
                        f"({len(cabinet_alloc.slices)} MTO источников)",
                        code,
                    )
                self._process_allocation_result(
                    ds_row, mto_row, ds_amount, allocated_amount, code,
                    cabinet_alloc=cabinet_alloc,
                )
                if is_preliminary:
                    comment = ds_row.el[VALUES].comment or ""
                    if "[PRELIMINARY_PROCESSED]" not in comment:
                        ds_row.el[VALUES].comment = comment + "[PRELIMINARY_PROCESSED]"
                self._remove_exhausted_mto_rows(mto_rows, code)
                continue

            mto_row = self._find_mto_row_for_amount(mto_rows, ds_amount, ds_row, code)

            if not mto_row:
                if code == self.debug_code:
                    self._debug_print("Не найдена подходящая MTO строка", code)
                self._handle_insufficient_mto(ds_row, ds_amount, 0, spec_name, code)
                continue

            is_aggregated = hasattr(mto_row, 'is_aggregated') and mto_row.is_aggregated

            if is_aggregated and code == self.debug_code:
                self._debug_print("Работаем с агрегированной строкой", code)

            mto_amount = float(mto_row.get_value(VALUES) or 0)

            if code == self.debug_code:
                self._debug_print(f"Выбрана MTO строка: {self._get_row_debug_info(mto_row)}", code)

            if i == len(ds_rows) - 1:
                # Проверяем, является ли это кодом из таблицы замен и есть ли в спецификации строки с новыми кодами
                is_replacement_code = code in self.replacement_table
                has_new_codes = self._has_new_codes_in_specification(code, codes_dict)
                
                if is_replacement_code and has_new_codes:
                    # Для кодов из таблицы замен с новыми кодами в спецификации - берем только требуемое количество
                    allocated_amount = min(ds_amount, mto_amount)
                    if code == self.debug_code:
                        self._debug_print(f"Последняя строка (код замены с новыми кодами) - берем только требуемое: {allocated_amount}", code)
                else:
                    # Для обычных кодов или кодов замен без новых кодов - берем весь остаток
                    total_remaining = self._get_all_remaining_mto_amount(mto_rows, code)
                    allocated_amount = total_remaining
                    if code == self.debug_code:
                        self._debug_print(f"Последняя строка - берем весь остаток из всех MTO: {allocated_amount}", code)
                    
                    # Создаем агрегированную строку для корректного отображения информации обо всех использованных строках
                    mto_row = self._create_aggregated_mto_row_for_remaining(mto_rows, allocated_amount, ds_row, code)
            else:
                allocated_amount = min(ds_amount, mto_amount)
                if code == self.debug_code:
                    self._debug_print(f"Последняя строка - берем: {allocated_amount}", code)

            cabinet_alloc = getattr(mto_row, 'cabinet_allocation', None)
            self._process_allocation_result(
                ds_row, mto_row, ds_amount, allocated_amount, code,
                cabinet_alloc=cabinet_alloc,
            )
            
            # Если это предварительный проход, помечаем DS строку как обработанную
            if is_preliminary:
                comment = ds_row.el[VALUES].comment or ""
                if "[PRELIMINARY_PROCESSED]" not in comment:
                    ds_row.el[VALUES].comment = comment + "[PRELIMINARY_PROCESSED]"

            if i == len(ds_rows) - 1:
                # Проверяем, нужно ли обновлять все оставшиеся MTO строки
                is_replacement_code = code in self.replacement_table
                has_new_codes = self._has_new_codes_in_specification(code, codes_dict)
                
                if is_replacement_code and has_new_codes:
                    # Для кодов замен с новыми кодами - обновляем только найденную строку
                    self._update_mto_row_after_allocation(mto_row, allocated_amount, code)
                    # Удаляем исчерпанную строку из списка
                    new_mto_amount = float(mto_row.get_value(VALUES) or 0)
                    if new_mto_amount <= 0:
                        mto_rows.remove(mto_row)
                        if code == self.debug_code:
                            self._debug_print("MTO строка исчерпана, удаляем из обработки", code)
                else:
                    # Для обычных кодов - обновляем все оставшиеся MTO строки
                    self._update_all_remaining_mto_rows(mto_rows, allocated_amount, code)
                    # Очищаем список, так как все строки исчерпаны
                    mto_rows.clear()
                    if code == self.debug_code:
                        self._debug_print("Последняя строка - все MTO строки исчерпаны", code)
            else:
                # Обновляем только найденную MTO строку
                self._update_mto_row_after_allocation(mto_row, allocated_amount, code)

                # Если MTO строка исчерпана, удаляем её из списка
                new_mto_amount = float(mto_row.get_value(VALUES) or 0)
                if new_mto_amount <= 0:
                    mto_rows.remove(mto_row)
                    if code == self.debug_code:
                        self._debug_print("MTO строка исчерпана, удаляем из обработки", code)

    def _update_all_remaining_mto_rows(self, mto_rows: List[RowStd], total_allocated_amount: float, code: str):
        """Обновляет все оставшиеся MTO строки для последней DS строки"""
        if code == self.debug_code:
            self._debug_print(f"Обновление всех оставшихся MTO строк, общее количество: {total_allocated_amount}", code)
        
        remaining_amount = total_allocated_amount
        rows_to_remove = []

        def _mto_take_order(row: RowStd) -> tuple:
            has_ic = 0 if norm_in_cabinet(row.get_value(IN_CABINET)) else 1
            return (has_ic, float(row.get_value(VALUES) or 0))

        for mto_row in sorted(mto_rows, key=_mto_take_order):
            if remaining_amount <= 0:
                break
                
            current_amount = float(mto_row.get_value(VALUES) or 0)
            if current_amount <= 0:
                continue
                
            # Берем из этой строки столько, сколько можем
            take_amount = min(remaining_amount, current_amount)
            
            if code == self.debug_code:
                self._debug_print(f"Обновляем MTO строку: {self._get_row_debug_info(mto_row)}, берем: {take_amount}", code)
            
            # Обновляем VALUES
            new_amount = current_amount - take_amount
            mto_row.el[VALUES].value = max(0, new_amount)
            
            # Обновляем TAGS
            self._update_mto_tags(mto_row, take_amount, code)
            
            remaining_amount -= take_amount
            
            if code == self.debug_code:
                self._debug_print(f"MTO VALUES: {current_amount} -> {new_amount}", code)
            
            # Если строка исчерпана, помечаем для удаления
            if new_amount <= 0:
                rows_to_remove.append(mto_row)
                if code == self.debug_code:
                    self._debug_print("MTO строка исчерпана, помечена для удаления", code)
        
        # Удаляем исчерпанные строки
        for row_to_remove in rows_to_remove:
            if row_to_remove in mto_rows:
                mto_rows.remove(row_to_remove)
                if code == self.debug_code:
                    self._debug_print("MTO строка удалена из списка", code)

    def _update_mto_row_after_allocation(self, mto_row: RowStd, allocated_amount: float, code: str):
        """Обновляет MTO строку после распределения количества"""
        # Для агрегированных строк просто снимаем признак (данные уже обновлены)
        if hasattr(mto_row, 'is_aggregated') and mto_row.is_aggregated:
            if code == self.debug_code:
                self._debug_print("Снимаем признак агрегации", code)
            delattr(mto_row, 'is_aggregated')

        # Уменьшаем VALUES
        current_amount = float(mto_row.get_value(VALUES) or 0)
        new_amount = current_amount - allocated_amount
        mto_row.el[VALUES].value = max(0, new_amount)

        if code == self.debug_code:
            self._debug_print(f"MTO VALUES: {current_amount} -> {new_amount}")

        # Обрабатываем TAGS
        self._update_mto_tags(mto_row, allocated_amount, code)

    def _update_mto_tags(self, mto_row: RowStd, allocated_amount: float, code: str):
        """Обновляет TAGS в MTO строке после распределения"""
        try:
            tags_list = mto_row.get_tags_list()
            current_amount = float(mto_row.get_value(VALUES) or 0)

            # Если количество стало нулевым, очищаем TAGS
            if current_amount <= 0:
                mto_row.set_tags_list([])
                if code == self.debug_code:
                    self._debug_print("Очищены TAGS (VALUES=0)", code)
                return

            if code == self.debug_code:
                self._debug_print(f"Обновление TAGS: было {tags_list}, allocated: {allocated_amount}", code)

            if not tags_list:
                return

            # Проверяем соответствие количества тегов и значения
            if len(tags_list) != current_amount + allocated_amount:  # + allocated потому что VALUES уже уменьшено
                # Несоответствие - логируем ошибку
                if code == self.debug_code:
                    self._debug_print(
                        f"Несоответствие TAGS и VALUES: TAGS={len(tags_list)}, VALUES было={current_amount + allocated_amount}",
                        code)

                # Берем пропорциональное количество тегов
                tags_to_remove = int(allocated_amount)
                if tags_to_remove > 0 and tags_to_remove <= len(tags_list):
                    # Удаляем первые N тегов
                    remaining_tags = tags_list[tags_to_remove:]
                    mto_row.set_tags_list(remaining_tags)
                    if code == self.debug_code:
                        self._debug_print(f"TAGS обновлены (пропорционально): {remaining_tags}", code)
            else:
                # Корректное соответствие - удаляем точное количество
                tags_to_remove = int(allocated_amount)
                if tags_to_remove > 0:
                    remaining_tags = tags_list[tags_to_remove:]
                    mto_row.set_tags_list(remaining_tags)
                    if code == self.debug_code:
                        self._debug_print(f"TAGS обновлены (точно): {remaining_tags}", code)

        except (ValueError, TypeError) as e:
            if code == self.debug_code:
                self._debug_print(f"Ошибка обработки TAGS: {e}", code)

    def _get_all_remaining_mto_amount(self, mto_rows: List[RowStd], code: str) -> float:
        """Возвращает общее количество из всех оставшихся MTO строк"""
        total_amount = sum(float(row.get_value(VALUES) or 0) for row in mto_rows)
        if code == self.debug_code:
            self._debug_print(f"Общее количество в оставшихся MTO строках: {total_amount}", code)
        return total_amount

    def _consume_allocation_slices(self, cabinet_alloc: CabinetAllocationResult, code: str) -> None:
        """Decrement VALUES/TAGS on each MTO row in a cabinet-aware allocation."""
        for sl in cabinet_alloc.slices:
            row = sl.row
            take = sl.take_amount
            current_amount = float(row.get_value(VALUES) or 0)
            row.el[VALUES].value = max(0.0, current_amount - take)
            self._update_mto_tags(row, take, code)

    def _remove_exhausted_mto_rows(self, mto_rows: List[RowStd], code: str) -> None:
        exhausted = [r for r in mto_rows if float(r.get_value(VALUES) or 0) <= 0]
        for row in exhausted:
            if row in mto_rows:
                mto_rows.remove(row)
                if code == self.debug_code:
                    self._debug_print("MTO строка исчерпана, удаляем из обработки", code)

    def _apply_in_cabinet_to_ds(
        self,
        ds_row: RowStd,
        cabinet_alloc: CabinetAllocationResult | None,
    ) -> None:
        if cabinet_alloc is None or not cabinet_alloc.slices:
            return
        if IN_CABINET not in ds_row.el:
            return
        ic_value = cabinet_alloc.in_cabinet_value
        ds_row.el[IN_CABINET].value = ic_value
        note = cabinet_alloc.in_cabinet_comment
        if note:
            prev = ds_row.el[IN_CABINET].comment or ""
            if note not in prev:
                ds_row.el[IN_CABINET].comment = (prev + note + "\n") if prev else (note + "\n")

    def _get_tags_for_cabinet_alloc(
        self,
        cabinet_alloc: CabinetAllocationResult,
        code: str,
    ) -> List[str]:
        tags: List[str] = []
        for sl in cabinet_alloc.slices:
            tags.extend(self._get_tags_for_allocation(sl.row, sl.take_amount, code))
        return tags

    def _allocate_ds_row(
        self,
        ds_row: RowStd,
        mto_rows: List[RowStd],
        need_amount: float,
        code: str,
    ) -> tuple[Optional[RowStd], CabinetAllocationResult]:
        """Take ``need_amount`` from MTO rows (cabinet lines first)."""
        cabinet_alloc = allocate_mto_slices(mto_rows, need_amount, ds_row)
        if cabinet_alloc.total_taken <= 0 or cabinet_alloc.primary_row is None:
            return None, cabinet_alloc
        self._consume_allocation_slices(cabinet_alloc, code)
        return cabinet_alloc.primary_row, cabinet_alloc

    def _set_aggregated_in_cabinet(
        self,
        aggregated_row: RowStd,
        used_rows_info: List[dict],
    ) -> CabinetAllocationResult:
        cabinet_alloc = cabinet_alloc_from_used_rows(used_rows_info, aggregated_row)
        aggregated_row.cabinet_allocation = cabinet_alloc
        if IN_CABINET in aggregated_row.el:
            aggregated_row.el[IN_CABINET].value = cabinet_alloc.in_cabinet_value
        return cabinet_alloc

    def _find_mto_row_for_amount(self,
                                 mto_rows: List[RowStd],
                                 amount: float,
                                 ds_row: RowStd,
                                 code: str) -> Optional[RowStd]:
        """Находит строку MTO с достаточным количеством или создает агрегированную строку"""
        available_rows = []

        if code == self.debug_code:
            for i, mto_row in enumerate(mto_rows):
                self._debug_print(f"Поиск строк МТО для подстановки: MTO #{i}: {self._get_row_debug_info(mto_row)}",
                                  code)

        # Сначала сортируем все строки по приоритету, затем ищем подходящие
        ds_tag = ds_row.get_value(TAGS)
        ds_spec = self._get_ds_specification_context(ds_row)
        ds_code = ds_row.get_value(CODE)

        def get_priority(mto_row):
            mto_code = mto_row.get_value(CODE)
            mto_tag = mto_row.get_value(TAGS)
            mto_spec = self._get_mto_specification_context(mto_row)

            if mto_code == ds_code:
                return 0  # Высший приоритет - совпадение CODE
            elif mto_tag == ds_tag:
                return 1  # Высокий приоритет - совпадение TAG
            elif mto_spec == ds_spec:
                return 2  # Средний приоритет - совпадение спецификации
            else:
                return 3  # Низкий приоритет

        # Сортируем все MTO строки по приоритету
        sorted_mto_rows = sorted(mto_rows, key=lambda x: (get_priority(x), float(x.get_value(VALUES) or 0)))

        if code == self.debug_code:
            self._debug_print(f"Сортировка по приоритету. DS: CODE={ds_code}, TAG={ds_tag}, SPEC={ds_spec}", code)
            for i, mto_row in enumerate(sorted_mto_rows):
                priority = get_priority(mto_row)
                mto_code = mto_row.get_value(CODE)
                mto_tag = mto_row.get_value(TAGS)
                mto_spec = self._get_mto_specification_context(mto_row)
                mto_amount = float(mto_row.get_value(VALUES) or 0)
                self._debug_print(f"SORTED #{i}: priority={priority}, amount={mto_amount} (CODE={mto_code}, TAG={mto_tag}, SPEC={mto_spec})", code)

        # Ищем строки с достаточным количеством в порядке приоритета
        for mto_row in sorted_mto_rows:
            mto_amount = float(mto_row.get_value(VALUES) or 0)
            if code == self.debug_code:
                self._debug_print(f"Проверка: mto_amount={mto_amount}, amount={amount}", code)

            if mto_amount >= amount and mto_amount > 0:
                available_rows.append((mto_row, mto_amount))

        # Проверяем, есть ли строка с точным совпадением кода, но недостаточным количеством
        exact_code_found = False
        for mto_row in sorted_mto_rows:
            mto_code = mto_row.get_value(CODE)
            if mto_code == ds_code and mto_row.get_value(VALUES) > 0 :
                exact_code_found = True
                break

        if available_rows:
            if code == self.debug_code:
                self._debug_print(f"Найдены строки с достаточным количеством: {len(available_rows)}", code)
                for i, (mto_row, mto_amount) in enumerate(available_rows):
                    self._debug_print(f"AVAILABLE #{i}: {self._get_row_debug_info(mto_row)}", code)

            def _pick_key(item: tuple) -> tuple:
                mto_row, mto_amount = item
                has_ic = 0 if norm_in_cabinet(mto_row.get_value(IN_CABINET)) else 1
                return (get_priority(mto_row), has_ic, -mto_amount)

            available_rows.sort(key=_pick_key)
            return available_rows[0][0]

        # Если есть строка с точным совпадением кода, но недостаточным количеством,
        # создаем агрегированную строку с приоритетом по кодам
        if exact_code_found:
            if code == self.debug_code:
                self._debug_print("Найдена строка с точным совпадением кода, но недостаточным количеством. Создаем агрегированную строку с приоритетом по кодам", code)
            return self._create_aggregated_mto_row(sorted_mto_rows, amount, ds_row, code)

        # Если не нашли строку с достаточным количеством, ВСЕГДА создаем агрегированную строку
        total_available = sum(float(row.get_value(VALUES) or 0) for row in mto_rows)
        if code == self.debug_code:
            self._debug_print(
                f"Не найдено строк с достаточным количеством. Суммарно доступно: {total_available}, требуется: {amount}",
                code)

        # ВСЕГДА создаем агрегированную строку, даже если суммарно не хватает
        # Программа сама решит как пометить недостаток в процессе распределения
        if code == self.debug_code:
            self._debug_print("Создаем агрегированную строку (даже если не хватает)", code)

        return self._create_aggregated_mto_row(sorted_mto_rows, amount, ds_row, code)

    def _create_aggregated_mto_row(self,
                                   mto_rows: List[RowStd],
                                   required_amount: float,
                                   ds_row: RowStd,
                                   code: str) -> RowStd:
        """Создает агрегированную MTO строку из нескольких строк (даже если не хватает)"""
        if code == self.debug_code:
            self._debug_print(f"Создание агрегированной строки для amount={required_amount}", code)

        # Рассчитываем сколько фактически можем взять
        total_available = sum(float(row.get_value(VALUES) or 0) for row in mto_rows)
        actual_amount = min(required_amount, total_available)

        if code == self.debug_code:
            self._debug_print(f"Фактически можем взять: {actual_amount} из {required_amount}", code)

        # Собираем информацию об использованных строках
        source_numbers = []
        source_annotations: List[str] = []
        source_tags = []
        used_rows_info = []

        alloc_preview = allocate_mto_slices(mto_rows, actual_amount, ds_row)
        for sl in alloc_preview.slices:
            mto_row = sl.row
            take_amount = sl.take_amount
            mto_amount = float(mto_row.get_value(VALUES) or 0)
            if mto_amount <= 0:
                continue

            used_rows_info.append({
                'row': mto_row,
                'take_amount': take_amount,
                'original_amount': mto_amount,
                'original_tags': mto_row.get_tags_list() if mto_row.get_tags_list() else []
            })

            source_number = mto_row.get_value(NUMBERS) or ""
            if source_number and str(source_number) not in source_numbers:
                source_numbers.append(str(source_number))

            ann_raw = mto_row.get_value(ANNOTATION)
            ann_s = str(ann_raw).strip() if ann_raw is not None else ""
            if ann_s and ann_s not in source_annotations:
                source_annotations.append(ann_s)

            mto_tags = mto_row.get_tags_list()
            if mto_tags:
                tags_to_take = min(len(mto_tags), int(take_amount))
                source_tags.extend(mto_tags[:tags_to_take])

        # Определяем коды источников и отличающиеся коды от обрабатываемого DS кода
        try:
            source_codes_used = []
            source_amounts_used = []
            for info in used_rows_info:
                if info.get('take_amount', 0) > 0:
                    src_code = info['row'].get_value(CODE)
                    if src_code is not None:
                        source_codes_used.append(src_code)
                        source_amounts_used.append(info['take_amount'])
            
            # Список уникальных кодов в порядке появления с суммарными количествами
            seen_codes = {}
            source_codes_unique = []
            source_amounts_unique = []
            for i, src_code in enumerate(source_codes_used):
                if src_code not in seen_codes:
                    seen_codes[src_code] = len(source_codes_unique)
                    source_codes_unique.append(src_code)
                    source_amounts_unique.append(source_amounts_used[i])
                else:
                    # Суммируем количество для уже существующего кода
                    idx = seen_codes[src_code]
                    source_amounts_unique[idx] += source_amounts_used[i]
            
            differing_codes = [c for c in source_codes_unique if c != code]
        except Exception:
            source_codes_unique = []
            source_amounts_unique = []
            differing_codes = []

        # Теперь обновляем исходные строки
        for info in used_rows_info:
            mto_row = info['row']
            take_amount = info['take_amount']
            original_amount = info['original_amount']
            original_tags = info['original_tags']

            # ОБНОВЛЯЕМ VALUES в исходной строке
            new_amount = original_amount - take_amount
            mto_row.el[VALUES].value = max(0, new_amount)

            # ОБНОВЛЯЕМ TAGS в исходной строке
            if original_tags:
                tags_to_take = min(len(original_tags), int(take_amount))
                remaining_tags = original_tags[tags_to_take:]
                mto_row.set_tags_list(remaining_tags)

            if code == self.debug_code:
                self._debug_print(f"Обновлена строка: VALUES {original_amount} -> {new_amount}", code)

        aggregated_row = alloc_preview.primary_row or mto_rows[0]

        aggregated_row.el[VALUES].value = actual_amount

        if source_numbers:
            current_number = aggregated_row.get_value(NUMBERS) or ""
            if current_number and current_number not in source_numbers:
                source_numbers.insert(0, str(current_number))
            aggregated_row.el[NUMBERS].value = ", ".join(source_numbers)
            if code == self.debug_code:
                self._debug_print(f"Агрегированные NUMBERS: {aggregated_row.el[NUMBERS].value}", code)

        if source_annotations:
            cur_ann_raw = aggregated_row.get_value(ANNOTATION)
            cur_ann = str(cur_ann_raw).strip() if cur_ann_raw is not None else ""
            if cur_ann and cur_ann not in source_annotations:
                source_annotations.insert(0, cur_ann)
            aggregated_row.el[ANNOTATION].value = ", ".join(source_annotations)
            if code == self.debug_code:
                self._debug_print(f"Агрегированные ANNOTATION: {aggregated_row.el[ANNOTATION].value}", code)

        if source_tags:
            aggregated_row.set_tags_list(source_tags)
            if code == self.debug_code:
                self._debug_print(f"Агрегированные TAGS: {source_tags}", code)

        self._set_aggregated_in_cabinet(aggregated_row, used_rows_info)

        aggregated_row.is_aggregated = True
        aggregated_row.aggregated_source_codes = source_codes_unique
        aggregated_row.aggregated_source_amounts = source_amounts_unique
        if differing_codes:
            aggregated_row.aggregated_replacement_codes = differing_codes

        # Обновляем аннотацию
        aggregated_row.el[ANNOTATION_3].value = "Агрегировано из нескольких строк"
        # Color.set_el_color(aggregated_row.el[ANNOTATION_3], Color.soft_blue)

        if self.cabinet_trace is not None:
            self.cabinet_trace.log_aggregation_sources(
                code=code,
                ds_row=ds_row,
                used_rows_info=used_rows_info,
                aggregated_row=aggregated_row,
            )

        if code == self.debug_code:
            self._debug_print(f"Создана агрегированная строка: {self._get_row_debug_info(aggregated_row)}", code)
            if actual_amount < required_amount:
                self._debug_print(
                    f"ВНИМАНИЕ: Не хватает количества! Взято: {actual_amount}, требуется: {required_amount}", code)

        return aggregated_row

    def _create_aggregated_mto_row_for_remaining(self,
                                                mto_rows: List[RowStd],
                                                total_amount: float,
                                                ds_row: RowStd,
                                                code: str) -> RowStd:
        """Создает агрегированную MTO строку из всех оставшихся строк для последней DS строки"""
        if code == self.debug_code:
            self._debug_print(f"Создание агрегированной строки для всего остатка: {total_amount}", code)

        # Собираем информацию об использованных строках
        source_numbers = []
        source_annotations: List[str] = []
        source_tags = []
        remaining_amount = total_amount
        used_rows_info = []
        source_codes_used = []
        source_amounts_used = []

        alloc_preview = allocate_mto_slices(mto_rows, total_amount, ds_row)
        for sl in alloc_preview.slices:
            mto_row = sl.row
            take_amount = sl.take_amount
            mto_amount = float(mto_row.get_value(VALUES) or 0)
            if mto_amount <= 0:
                continue

            used_rows_info.append({
                'row': mto_row,
                'take_amount': take_amount,
                'original_amount': mto_amount,
                'original_tags': mto_row.get_tags_list() if mto_row.get_tags_list() else []
            })

            src_code = mto_row.get_value(CODE)
            if src_code is not None:
                source_codes_used.append(src_code)
                source_amounts_used.append(take_amount)

            source_number = mto_row.get_value(NUMBERS) or ""
            if source_number and str(source_number) not in source_numbers:
                source_numbers.append(str(source_number))

            ann_raw = mto_row.get_value(ANNOTATION)
            ann_s = str(ann_raw).strip() if ann_raw is not None else ""
            if ann_s and ann_s not in source_annotations:
                source_annotations.append(ann_s)

            mto_tags = mto_row.get_tags_list()
            if mto_tags:
                tags_to_take = min(len(mto_tags), int(take_amount))
                source_tags.extend(mto_tags[:tags_to_take])

        for info in used_rows_info:
            mto_row = info['row']
            take_amount = info['take_amount']
            original_amount = info['original_amount']
            original_tags = info['original_tags']

            new_amount = original_amount - take_amount
            mto_row.el[VALUES].value = max(0, new_amount)

            if original_tags:
                tags_to_take = min(len(original_tags), int(take_amount))
                remaining_tags = original_tags[tags_to_take:]
                mto_row.set_tags_list(remaining_tags)

            if code == self.debug_code:
                self._debug_print(f"Обновлена строка: VALUES {original_amount} -> {new_amount}", code)

        aggregated_row = alloc_preview.primary_row or mto_rows[0]

        aggregated_row.el[VALUES].value = total_amount

        # Формируем NUMBERS как перечисление источников
        if source_numbers:
            current_number = aggregated_row.get_value(NUMBERS) or ""
            if current_number and current_number not in source_numbers:
                source_numbers.insert(0, str(current_number))
            aggregated_row.el[NUMBERS].value = ", ".join(source_numbers)
            if code == self.debug_code:
                self._debug_print(f"Агрегированные NUMBERS: {aggregated_row.el[NUMBERS].value}", code)

        if source_annotations:
            cur_ann_raw = aggregated_row.get_value(ANNOTATION)
            cur_ann = str(cur_ann_raw).strip() if cur_ann_raw is not None else ""
            if cur_ann and cur_ann not in source_annotations:
                source_annotations.insert(0, cur_ann)
            aggregated_row.el[ANNOTATION].value = ", ".join(source_annotations)
            if code == self.debug_code:
                self._debug_print(f"Агрегированные ANNOTATION: {aggregated_row.el[ANNOTATION].value}", code)

        if source_tags:
            aggregated_row.set_tags_list(source_tags)
            if code == self.debug_code:
                self._debug_print(f"Агрегированные TAGS: {source_tags}", code)

        self._set_aggregated_in_cabinet(aggregated_row, used_rows_info)

        aggregated_row.is_aggregated = True
        aggregated_row.aggregated_source_codes = source_codes_used
        aggregated_row.aggregated_source_amounts = source_amounts_used

        aggregated_row.el[ANNOTATION_3].value = "Агрегировано из всех оставшихся строк"

        if self.cabinet_trace is not None:
            self.cabinet_trace.log_aggregation_sources(
                code=code,
                ds_row=ds_row,
                used_rows_info=used_rows_info,
                aggregated_row=aggregated_row,
            )

        if code == self.debug_code:
            self._debug_print(f"Создана агрегированная строка для остатка: {self._get_row_debug_info(aggregated_row)}", code)

        return aggregated_row

    def _get_ds_specification_context(self, ds_row: RowStd):
        return V2Document.from_file_path(ds_row.get_value(DS_SPECIFICATION)).doc_Short_Title

    def _update_aggregated_comment_on_addition(self, ds_row: RowStd, added_amount: float, code: str):
        """Обновляет комментарий об агрегированных кодах при добавлении количества"""
        values_2_comment = ds_row.el[VALUES_2].comment or ""
        
        # Ищем существующий комментарий об агрегированных кодах
        import re
        pattern = r'\[AGG_CODES:\s*([^\]]+)\]'
        match = re.search(pattern, values_2_comment)
        
        if match:
            # Парсим существующие коды и количества
            codes_data_str = match.group(1)
            aggregated_info = []
            
            try:
                for code_amount_pair in codes_data_str.split(','):
                    code_amount_pair = code_amount_pair.strip()
                    if ':' in code_amount_pair:
                        source_code, amount_str = code_amount_pair.split(':', 1)
                        source_code = source_code.strip()
                        amount_str = amount_str.strip()
                        
                        try:
                            amount = float(amount_str)
                            # Распределяем добавленное количество пропорционально между существующими кодами
                            # Находим общее количество для пропорционального распределения
                            total_existing = sum(float(pair.split(':')[1].strip()) for pair in codes_data_str.split(',') if ':' in pair)
                            if total_existing > 0:
                                proportion = amount / total_existing
                                new_amount = amount + (added_amount * proportion)
                            else:
                                new_amount = amount + added_amount
                            
                            aggregated_info.append(f"{source_code}:{new_amount:.2f}")
                        except ValueError:
                            continue
            except Exception:
                # В случае ошибки оставляем как есть
                return
            
            # Создаем новый комментарий
            new_aggregated_comment = f"[AGG_CODES: {','.join(aggregated_info)}]"
            
            # Заменяем старый комментарий новым
            new_comment = re.sub(pattern, new_aggregated_comment, values_2_comment)
            ds_row.el[VALUES_2].comment = new_comment
            
            if code == self.debug_code:
                self._debug_print(f"[COMMENT_UPDATE] Обновлен комментарий агрегации: {new_aggregated_comment}", code)

    def _get_mto_specification_context(self, mto_row: RowStd) -> str:
        """Определяет контекст спецификации для MTO строки"""
        mto_spec = V2Document.from_file_path(mto_row.t_com.file_name).doc_Short_Title
        return mto_spec

    def _process_allocation_result(self,
                                   ds_row: RowStd,
                                   mto_row: RowStd,
                                   ds_amount: float,
                                   allocated_amount: float,
                                   code: str,
                                   cabinet_alloc: CabinetAllocationResult | None = None):
        """Обрабатывает результат распределения с учетом замен"""
        if cabinet_alloc is None:
            cabinet_alloc = getattr(mto_row, 'cabinet_allocation', None)
        if self.cabinet_trace is not None:
            self.cabinet_trace.log_allocation(
                phase="before_transfer",
                code=code,
                ds_row=ds_row,
                mto_row=mto_row,
                allocated=allocated_amount,
            )

        if code == self.debug_code:
            self._debug_print(f"Обработка распределения: DS={ds_amount}, allocated={allocated_amount}")
            self._debug_print(f"MTO до обработки: {self._get_row_debug_info(mto_row)}")

        if cabinet_alloc and len(cabinet_alloc.slices) > 1:
            tags_to_transfer = self._get_tags_for_cabinet_alloc(cabinet_alloc, code)
        else:
            tags_to_transfer = self._get_tags_for_allocation(mto_row, allocated_amount, code)

        if code == self.debug_code and tags_to_transfer:
            self._debug_print(f"Теги для переноса: {tags_to_transfer}", code)

        # Подготовим маркеры замен ДО переноса значений, а применим их ПОСЛЕ, чтобы перенос не стер форматирование
        has_explicit_replacement = hasattr(mto_row, 'replacement_info') and ds_row.get_value(CODE) != mto_row.get_value(CODE)
        replacement_verified = False
        if has_explicit_replacement:
            replacement_info = mto_row.replacement_info
            replacement_verified = (replacement_info.get('status') == 'ПРОВЕРЕН')
            if code == self.debug_code:
                self._debug_print(f"Обнаружена замена: {replacement_info}", code)

        has_aggregated_replacements = hasattr(mto_row, 'aggregated_replacement_codes') and bool(mto_row.aggregated_replacement_codes)
        aggregated_codes_marker = None
        if has_aggregated_replacements:
            codes_list = ", ".join(mto_row.aggregated_replacement_codes)
            aggregated_codes_marker = f"[AGG_REPLACED_FROM: {codes_list}]"
            if code == self.debug_code:
                self._debug_print(f"Обнаружены отличающиеся коды в агрегировании: {mto_row.aggregated_replacement_codes}", code)

        self._transfer_mto_to_ds_with_tags(ds_row, mto_row, tags_to_transfer, allocated_amount, code)
        self._apply_in_cabinet_to_ds(ds_row, cabinet_alloc)

        # ОТЛАДОЧНЫЙ ВЫВОД
        current_values_2 = float(ds_row.get_value(VALUES_2) or 0)
        new_values_2 = allocated_amount  # Обратите внимание: здесь перезаписывается, а не добавляется!
        self._debug_values_2_write(ds_row, current_values_2, new_values_2,
                                   "ALLOCATION_RESULT", code)
        ds_row.el[VALUES_2].value = allocated_amount

        # Сохраняем информацию об агрегированных кодах в комментарии для check_comparison_results
        if hasattr(mto_row, 'aggregated_source_codes') and mto_row.aggregated_source_codes:
            # Создаем детальный комментарий с точными кодами и количествами из каждой использованной строки
            aggregated_info = []
            for i, source_code in enumerate(mto_row.aggregated_source_codes):
                # Используем точную информацию об использованном количестве из каждой строки
                if hasattr(mto_row, 'aggregated_source_amounts') and i < len(mto_row.aggregated_source_amounts):
                    amount = mto_row.aggregated_source_amounts[i]
                    aggregated_info.append(f"{source_code}:{amount:.2f}")
                else:
                    # Если нет точной информации, это ошибка - логируем и пропускаем
                    if code == self.debug_code:
                        self._debug_print(f"[ERROR] Нет точной информации о количестве для кода {source_code}", code)
            
            if aggregated_info:  # Добавляем комментарий только если есть точная информация
                aggregated_comment = f"[AGG_CODES: {','.join(aggregated_info)}]"
                
                # Добавляем к комментарию VALUES_2
                values_2_comment = ds_row.el[VALUES_2].comment or ""
                if aggregated_comment not in values_2_comment:
                    ds_row.el[VALUES_2].comment = values_2_comment + aggregated_comment

        # Применяем визуальные маркеры замен ПОСЛЕ переноса, чтобы их не перетёрли
        if has_explicit_replacement or has_aggregated_replacements:
            ds_row.el[ANNOTATION_3].value = "Заменен"
            if replacement_verified:
                Color.set_el_color(ds_row.el[ANNOTATION_3], Color.soft_pink)
                Color.set_el_color(ds_row.el[CODE], Color.soft_pink)
                # Подсветим также CODE_2 для наглядности
                Color.set_el_color(ds_row.el[CODE_2], Color.soft_pink)
            else:
                Color.set_el_color(ds_row.el[ANNOTATION_3], Color.yellow)
                Color.set_el_color(ds_row.el[CODE], Color.yellow)
                Color.set_el_color(ds_row.el[CODE_2], Color.yellow)

            # Добавим маркер источников к комментарию CODE_2, если он есть
            if aggregated_codes_marker:
                code2_comment_after = ds_row.el[CODE_2].comment or ""
                if aggregated_codes_marker not in code2_comment_after:
                    ds_row.el[CODE_2].comment = code2_comment_after + aggregated_codes_marker

        # Определяем статус и устанавливаем цвета
        if allocated_amount == ds_amount:
            self._handle_match(ds_row, allocated_amount, ds_amount, code)
        elif allocated_amount < ds_amount:
            self._handle_insufficient(ds_row, allocated_amount, ds_amount, code)
        else:
            self._handle_excess(ds_row, allocated_amount, ds_amount, code)

        if self.cabinet_trace is not None:
            self.cabinet_trace.log_allocation(
                phase="after_transfer",
                code=code,
                ds_row=ds_row,
                mto_row=mto_row,
                allocated=allocated_amount,
            )

        if code == self.debug_code:
            self._debug_print(f"MTO после обработки: {self._get_row_debug_info(mto_row)}", code)
            self._debug_print(f"DS после обработки: {self._get_row_debug_info(ds_row)}", code)

    def _get_tags_for_allocation(self, mto_row: RowStd, allocated_amount: float, code: str) -> List[str]:
        """Возвращает теги для распределения в DS строку"""
        try:
            tags_list = mto_row.get_tags_list()
            current_amount = float(mto_row.get_value(VALUES) or 0)

            if not tags_list:
                return []

            # Определяем сколько тегов взять
            if len(tags_list) == current_amount + allocated_amount:  # + allocated потому что VALUES уже уменьшено
                # Корректное соответствие - берем пропорционально
                tags_to_take = int(allocated_amount)
                result = tags_list[:tags_to_take] if tags_to_take <= len(tags_list) else tags_list.copy()
                if code == self.debug_code:
                    self._debug_print(f"Теги взяты пропорционально: {result}", code)
                return result
            else:
                # Несоответствие - берем первый тег или пустой список
                result = [tags_list[0]] if tags_list else []
                if code == self.debug_code:
                    self._debug_print(f"Теги взяты по умолчанию (несоответствие): {result}", code)
                return result

        except (ValueError, TypeError, IndexError) as e:
            if code == self.debug_code:
                self._debug_print(f"Ошибка получения тегов: {e}", code)
            return []

    def _transfer_mto_to_ds_with_tags(self,
                                      ds_row: RowStd,
                                      mto_row: RowStd,
                                      tags_to_transfer: List[str],
                                      allocated_amount: float,
                                      code: str):
        """Переносит данные из MTO с учетом тегов"""
        # Стандартный перенос атрибутов
        self._transfer_mto_to_ds(ds_row, mto_row)


        # Специальная обработка TAGS
        if tags_to_transfer:
            # Устанавливаем теги в TAGS_2
            ds_row.el[TAGS_2].value = tags_to_transfer
            if code == self.debug_code:
                self._debug_print(f"Перенесены теги в TAGS_2: {tags_to_transfer}", code)
        else:
            # Если тегов нет, используем оригинальное значение
            ds_row.el[TAGS_2].value = mto_row.get_value(TAGS)
            if code == self.debug_code:
                self._debug_print(f"Перенесены оригинальные теги: {mto_row.get_value(TAGS)}", code)

    def _handle_match(self, ds_row: RowStd, allocated: float, required: float, code: str):
        """Обрабатывает полное совпадение количеств"""
        if code == self.debug_code:
            self._debug_print("Статус: СОВПАДЕНИЕ", code)
        Color.set_el_color(ds_row.el[VALUES], Color.green)
        Color.set_el_color(ds_row.el[VALUES_2], Color.green)
        ds_row.el[ANNOTATION_2].value = ComparisonStatus.MATCH.value 

    def _handle_insufficient(self, ds_row: RowStd, allocated: float, required: float, code: str):
        """Обрабатывает недостаточное количество"""
        if code == self.debug_code:
            self._debug_print(f"Статус: НЕДОСТАТОК ({allocated} < {required})", code)
        Color.set_el_color(ds_row.el[VALUES], Color.red)
        Color.set_el_color(ds_row.el[VALUES_2], Color.yellow)
        Color.set_el_color(ds_row.el[ANNOTATION_2], Color.yellow)
        ds_row.el[ANNOTATION_2].value = ComparisonStatus.INSUFFICIENT.value 

    def _handle_excess(self, ds_row: RowStd, allocated: float, required: float, code: str):
        """Обрабатывает избыточное количество"""
        if code == self.debug_code:
            self._debug_print(f"Статус: ИЗБЫТОК ({allocated} > {required})", code)
        Color.set_el_color(ds_row.el[VALUES], Color.yellow)
        Color.set_el_color(ds_row.el[VALUES_2], Color.red)
        Color.set_el_color(ds_row.el[ANNOTATION_2], Color.yellow)
        ds_row.el[ANNOTATION_2].value = ComparisonStatus.EXCESS.value 

    def _handle_insufficient_mto(self,
                                 ds_row: RowStd,
                                 required: float,
                                 allocated: float,
                                 spec_name: str,
                                 code: str):
        """Обрабатывает полное отсутствие MTO"""
        if code == self.debug_code:
            self._debug_print("Статус: MTO ЗАКОНЧИЛИСЬ", code)
        Color.set_el_color(ds_row.el[VALUES], Color.red)
        Color.set_el_color(ds_row.el[ANNOTATION_2], Color.red)
        ds_row.el[ANNOTATION_2].value = ComparisonStatus.MTO_LESS_DS_BY_ROW.value 
        comment_text = f"Недостаточно MTO для кода {code} в {spec_name}\n"
        ds_row.el[VALUES].comment = ds_row.el[VALUES].comment + comment_text



    def _handle_code_not_found(self, ds_rows: List[RowStd], spec_name: str,
                               code: str):
        """Обрабатывает отсутствующий код в MTO"""
        for ds_row in ds_rows:
            if code == self.debug_code:
                self._debug_print("Статус: КОД НЕ НАЙДЕН В MTO", code)
            Color.set_el_color(ds_row.el[VALUES], Color.yellow)
            Color.set_el_color(ds_row.el[ANNOTATION_2], Color.yellow)
            comment_text = ComparisonStatus.NOT_FOUND.value + "\n"
            ds_row.el[ANNOTATION_2].value = comment_text
            comment_text = f"Код {code} не найден в MTO {spec_name}\n"
            ds_row.el[VALUES].comment = ds_row.el[VALUES].comment + comment_text



    def _transfer_mto_to_ds(self, ds_row: RowStd, mto_row: RowStd):
        """Переносит данные из MTO строки в правые столбцы DS строки"""
        mapping = {
            TAGS_2: TAGS,
            NUMBERS_2: NUMBERS,
            ANNOTATION_MTO: ANNOTATION,
            NAME_2: NAME,
            TYPE_MARK_2: TYPE_MARK,
            CODE_2: CODE,
            VENDOR_2: VENDOR,
            UNITS_2: UNITS,
            IN_CABINET: IN_CABINET,
        }
        for ds_attr, mto_attr in mapping.items():
            if mto_attr in mto_row.el:
                ds_row.el[ds_attr] = mto_row.el[mto_attr]
                # Color.set_el_color(ds_row.el[ds_attr], Color.no)



    def _add_new_positions(self,
                           ds_base_full: List[RowStd],
                           mto_dict: Dict[str, List[RowStd]],
                           ds_grouped: Dict[str, Dict[str, List[RowStd]]]):
        """
        Добавляет новые позиции из MTO, которых нет в DS (оптимизированная версия)
        
        ОПТИМИЗАЦИЯ:
        - Старая версия: O(n*m) - для каждой новой позиции ищем место вставки O(n) и вставляем O(n)
        - Новая версия: O(n + m*log(m)) - один проход для поиска индексов + сортировка + batch вставка
        - Использует batch-вставку в обратном порядке для избежания сдвига индексов
        - Значительно быстрее для больших объемов данных
        """
        if self.debug:
            self._debug_out(f"\nДобавляем новые позиции из MTO, которых нет в DS...")

        # Собираем все новые позиции по спецификациям
        new_positions_by_spec = {}
        
        for spec_name, mto_base in mto_dict.items():
            if spec_name not in ds_grouped:
                continue

            new_positions = []
            for mto_row in mto_base:
                if (isinstance(mto_row, RowStd) and
                        mto_row.row_type == RowType.position_row):

                    mto_values = float(mto_row.get_value(VALUES) or 0)
                    mto_code = mto_row.get_value(CODE)

                    # Добавляем только если количество больше нуля
                    if mto_values > 0:
                        if mto_code == self.debug_code:
                            self._debug_print(f"Найдена новая позиция: {self._get_row_debug_info(mto_row)}", mto_code)
                        new_positions.append(mto_row)
                    elif mto_code == self.debug_code:
                        self._debug_print(f"Пропускаем MTO (VALUES=0): {self._get_row_debug_info(mto_row)}", mto_code)
            
            if new_positions:
                new_positions_by_spec[spec_name] = new_positions

        # Выполняем batch-вставку для каждой спецификации
        self._batch_insert_new_positions(ds_base_full, new_positions_by_spec)

    def _batch_insert_new_positions(self, ds_base_full: List[RowStd], 
                                   new_positions_by_spec: Dict[str, List[RowStd]]):
        """Оптимизированная batch-вставка новых позиций"""
        if not new_positions_by_spec:
            return

        # Находим индексы вставки для каждой спецификации (оптимизированно)
        spec_insert_indices = self._find_spec_insert_indices(ds_base_full, new_positions_by_spec)

        # Создаем список всех вставок с индексами
        insertions = []
        for spec_name, new_positions in new_positions_by_spec.items():
            if spec_name not in spec_insert_indices:
                continue
                
            insert_index = spec_insert_indices[spec_name]
            for mto_row in new_positions:
                new_row = self._create_new_position_row(mto_row, spec_name)
                insertions.append((insert_index, new_row))
                
                if mto_row.get_value(CODE) == self.debug_code:
                    self._debug_print(f"Подготовлена новая позиция для вставки на индекс {insert_index}", 
                                    mto_row.get_value(CODE))

        # Сортируем вставки по индексу в обратном порядке (от большего к меньшему)
        # Это позволяет вставлять элементы без сдвига индексов
        insertions.sort(key=lambda x: x[0], reverse=True)
        
        # Выполняем все вставки
        for insert_index, new_row in insertions:
            ds_base_full.insert(insert_index, new_row)

    def _find_spec_insert_indices(self, ds_base_full: List[RowStd], 
                                 new_positions_by_spec: Dict[str, List[RowStd]]) -> Dict[str, int]:
        """Находит индексы вставки для каждой спецификации (оптимизированно)"""
        spec_insert_indices = {}
        specs_needed = set(new_positions_by_spec.keys())
        
        # Проходим по списку один раз и запоминаем последние индексы для нужных спецификаций
        for i, row in enumerate(ds_base_full):
            if (isinstance(row, RowStd) and
                    row.row_type == RowType.position_row):
                spec_name = V2Document.from_file_path(row.get_value(DS_SPECIFICATION)).doc_Short_Title
                if spec_name in specs_needed:
                    spec_insert_indices[spec_name] = i + 1  # Вставляем после последней строки
                    
        return spec_insert_indices

    def _insert_new_position(self, ds_base_full: List[RowStd], spec_name: str,
                             mto_row: RowStd):
        """Вставляет новую позицию в DS базу"""
        # Находим позицию для вставки (последняя строка спецификации)
        insert_index = -1
        for i, row in enumerate(ds_base_full):
            if (isinstance(row, RowStd) and
                    row.row_type == RowType.position_row and
                    V2Document.from_file_path(row.get_value(DS_SPECIFICATION)).doc_Short_Title == spec_name):
                insert_index = i

        if insert_index == -1:
            return

        # Вставляем после последней строки спецификации
        insert_index += 1
        new_row = self._create_new_position_row(mto_row, spec_name)
        ds_base_full.insert(insert_index, new_row)

        if mto_row.get_value(CODE) == self.debug_code:
            self._debug_print(f"Новая позиция вставлена на индекс {insert_index}", mto_row.get_value(CODE))

    def _create_new_position_row(self, mto_row: RowStd, spec_name: str) -> RowStd:
        """Создает новую строку для DS из MTO строки"""
        new_row = RowStd.get_std_check_row({}, TableComments())
        new_row.row_type = RowType.position_row
        new_row.el[ROW_TYPE].value = RowType.position_row

        # Устанавливаем признак новой позиции
        new_row.el[ANNOTATION_2].value = "Новая позиция"
        Color.set_el_color(new_row.el[ANNOTATION_2], Color.soft_cyan)
        # new_row.el[DS_SPECIFICATION].value = spec_name

        # Добавляем информацию о документе
        spec_title_system_list = spec_name.split("-")
        new_row.el[DS_TITLE].value = spec_title_system_list[0]
        new_row.el[DS_SYSTEM].value = spec_title_system_list[1]

        # Переносим данные из MTO
        mapping = {
            TAGS_2: TAGS,
            NUMBERS_2: NUMBERS,
            ANNOTATION_MTO: ANNOTATION,
            NAME_2: NAME,
            TYPE_MARK_2: TYPE_MARK,
            CODE_2: CODE,
            VENDOR_2: VENDOR,
            UNITS_2: UNITS,
            VALUES_2: VALUES,
            IN_CABINET: IN_CABINET,
        }

        for new_place, old_place in mapping.items():
            if old_place in mto_row.el:
                new_row.el[new_place].value = mto_row.el[old_place].value
                Color.set_el_color(new_row.el[new_place], Color.soft_cyan)

        return new_row





class RuleFactory:
    """Фабрика для создания правил сравнения"""

    @staticmethod
    def create_quantity_rule(tolerance=0.0) -> QuantityComparisonRule:
        return QuantityComparisonRule(tolerance=tolerance, priority=10)

    @staticmethod
    def create_vendor_rule() -> ComparisonRule:
        """Правило проверки совпадения вендоров"""

        class VendorComparisonRule(ComparisonRule):
            def apply(self, ds_row: RowStd, mto_rows: List[RowStd]) -> Optional[ComparisonResult]:
                ds_vendor = ds_row.get_value(VENDOR)
                mto_vendors = {row.get_value(VENDOR) for row in mto_rows}

                if ds_vendor not in mto_vendors:
                    return ComparisonResult.VENDOR_MISMATCH
                return None

        return VendorComparisonRule(priority=5)

    @staticmethod
    def create_unit_rule() -> ComparisonRule:
        """Правило проверки совпадения единиц измерения"""

        class UnitComparisonRule(ComparisonRule):
            def apply(self, ds_row: RowStd, mto_rows: List[RowStd]) -> Optional[ComparisonResult]:
                ds_unit = ds_row.get_value(UNITS)
                mto_units = {row.get_value(UNITS) for row in mto_rows}

                if ds_unit not in mto_units:
                    return ComparisonResult.UNIT_MISMATCH
                return None

        return UnitComparisonRule(priority=3)