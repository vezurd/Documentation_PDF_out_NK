# Handoff: статус RFP «Исключен из поставки»

Дата источника: **2026-09-03**. Реализация: **2026-09-03** (этот чат). Живой UNC DS29 в CI не гонялся.

Правила обновлены: `.cursor/rules/AI_rfp_parts.mdc`, `.cursor/rules/AI_rfq_context.mdc`, `.cursor/rules/AI_rfp_step4_excel.mdc`.

---

## Сделано

- Константа `RFP_SUPPLY_STATUS` (`rfp_supply_status`), net слот **15**; хелпер [rfp_supply_status.py](RFQ/tags_rfp_compare/rfp_supply_status.py).
- Части: плавающая шапка `статус\s*позиц` (не T/19). Канон. **«Исключен из поставки»** → net P. Иное непустое → WARN, Step4 не считает исключённой.
- `_COPY_COLS` обоих light-copy; `rfp_supply_status.load=true`.
- Step4: строка остаётся в Excel; skip match MTO/VO, `assign_position_status`, якорь unmatched MTO, очередь УЛ, Σ ДС BCC. `UL_COMPARE_STATUS` тот же канон., жёлтый. Overlay ГЭМ/ЗИП не перебивают. `POSITION_STATUS` **не** пишется этим текстом.
- Collapse: флаг `is_excluded_from_supply` в ключе.
- Баланс in/out **не** вычитает исключённые `VALUES` (консервация).
- Смоук: `tmp/test_rfp_supply_status_smoke.py`.

## Не делать снова

- Не мапить на `POSITION_STATUS` / не хардкодить T/19 у **частей**.
- Не вставлять колонку в core 0–10 net (`VALUES` @9).
- Не путать net T (`UNITS_CONVERSION_TRACE` @19) со статусом поставки.

## Опционально

Если UNC доступен: 13 строк ДС29 `8950-POS5` / `BCC0002946`, Excel 1358–1370.
