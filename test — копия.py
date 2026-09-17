def _find_mto_row_for_amount(self,
                             mto_rows: List[RowStd],
                             amount: float,
                             ds_row: RowStd,
                             code: str) -> Optional[RowStd]:
    """Находит строку MTO с достаточным количеством или создает агрегированную строку"""
    available_rows = []

    if code == self.debug_code:
        for i, mto_row in enumerate(mto_rows):
            self._debug_print(f"Поиск строк МТО для подстановки: MTO #{i}: {self._get_row_debug_info(mto_row)}", code)

    # Сначала ищем строки с достаточным количеством
    for mto_row in mto_rows:
        mto_amount = float(mto_row.get_value(VALUES) or 0)
        if code == self.debug_code:
            self._debug_print(f"Проверка: mto_amount={mto_amount}, amount={amount}", code)

        if mto_amount >= amount:
            available_rows.append((mto_row, mto_amount))

    if available_rows:
        if code == self.debug_code:
            self._debug_print(f"Найдены строки с достаточным количеством: {len(available_rows)}", code)
            for i, (mto_row, mto_amount) in enumerate(available_rows):
                self._debug_print(f"AVA #{i}: {self._get_row_debug_info(mto_row)}", code)

        # Сортируем по приоритету
        ds_tag = ds_row.get_value(TAGS)
        ds_spec = self._get_ds_specification_context(ds_row)

        def get_priority(mto_row):
            mto_tag = mto_row.get_value(TAGS)
            mto_spec = self._get_mto_specification_context(mto_row)

            if mto_tag == ds_tag:
                return 0  # Высший приоритет - совпадение TAG
            elif mto_spec == ds_spec:
                return 1  # Средний приоритет - совпадение спецификации
            else:
                return 2  # Низкий приоритет

        available_rows.sort(key=lambda x: (get_priority(x[0]), x[1]))
        return available_rows[0][0]

    # Если не нашли строку с достаточным количеством, проверяем суммарно
    total_available = sum(float(row.get_value(VALUES) or 0) for row in mto_rows)
    if code == self.debug_code:
        self._debug_print(f"Не найдено строк с достаточным количеством. Суммарно доступно: {total_available}", code)

    if total_available >= amount:
        # Суммарно хватает - создаем агрегированную строку
        if code == self.debug_code:
            self._debug_print("Суммарно хватает, создаем агрегированную строку", code)

        return self._create_aggregated_mto_row(mto_rows, amount, ds_row, code)
    else:
        # Суммарно не хватает
        if code == self.debug_code:
            self._debug_print("Суммарно не хватает", code)
        return None


def _create_aggregated_mto_row(self,
                               mto_rows: List[RowStd],
                               required_amount: float,
                               ds_row: RowStd,
                               code: str) -> RowStd:
    """Создает агрегированную MTO строку из нескольких строк"""
    if code == self.debug_code:
        self._debug_print(f"Создание агрегированной строки для amount={required_amount}", code)

    # Создаем новую строку на основе первой подходящей
    aggregated_row = copy.deepcopy(mto_rows[0])

    # Собираем информацию об источниках
    source_numbers = []
    source_tags = []
    remaining_amount = required_amount

    for mto_row in mto_rows:
        if remaining_amount <= 0:
            break

        mto_amount = float(mto_row.get_value(VALUES) or 0)
        if mto_amount <= 0:
            continue

        # Берем сколько нужно или сколько есть
        take_amount = min(remaining_amount, mto_amount)

        # Добавляем информацию об источнике
        source_number = mto_row.get_value(NUMBERS) or ""
        if source_number and source_number not in source_numbers:
            source_numbers.append(str(source_number))

        # Добавляем теги (если есть)
        mto_tags = mto_row.get_tags_list()
        if mto_tags:
            tags_to_take = min(len(mto_tags), int(take_amount))
            source_tags.extend(mto_tags[:tags_to_take])

        remaining_amount -= take_amount

    # Устанавливаем агрегированные данные
    aggregated_row.el[VALUES].value = required_amount

    # Формируем NUMBERS как перечисление источников
    if source_numbers:
        aggregated_row.el[NUMBERS].value = ", ".join(source_numbers)
        if code == self.debug_code:
            self._debug_print(f"Агрегированные NUMBERS: {aggregated_row.el[NUMBERS].value}", code)

    # Формируем TAGS как объединение тегов
    if source_tags:
        aggregated_row.set_tags_list(source_tags)
        if code == self.debug_code:
            self._debug_print(f"Агрегированные TAGS: {source_tags}", code)
    else:
        # Если тегов нет, используем теги из первой строки
        first_tags = mto_rows[0].get_tags_list()
        if first_tags:
            aggregated_row.set_tags_list(first_tags[:int(required_amount)])

    # Обновляем другие атрибуты для ясности
    aggregated_row.el[ANNOTATION_3].value = "Агрегировано из нескольких строк"
    Color.set_el_color(aggregated_row.el[ANNOTATION_3], Color.soft_blue)

    if code == self.debug_code:
        self._debug_print(f"Создана агрегированная строка: {self._get_row_debug_info(aggregated_row)}", code)

    return aggregated_row







