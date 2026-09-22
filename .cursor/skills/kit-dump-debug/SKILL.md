---
name: kit-dump-debug
description: >-
  Diagnoses an RD Catalog GUI handoff that starts with «RD Catalog · Передать роботу».
  Use when the user pastes that block, a Комплекты / Все документы / MTO cell,
  title and mark, or asks why a catalog cell, tooltip, folder, or revision disagrees
  with the disk or Google. Explain the cell first; edit code only if the user asked
  for a change.
---

# Kit dump debug

The paste is a snapshot of one GUI item, not a spec. Domain facts stay in the zone rules. This skill is only the order of work.

## Parse

Read the block from `rd_catalog/robot_handoff.py` format:

- `Где`, `Вкладка`, table or tree level, `Ячейка`
- `Титул`, `Марка`, `Сводка`, review/approval, revision columns
- paths, tooltips, and the human sentence after the block

The question is that sentence (or the screenshot caption). The dump is evidence.

## Which rule

Start with `.cursor/rules/AI_rd_catalog.mdc` if it is not already in context. Then only the sibling that matches the cell:

| Signal | Rule |
|---|---|
| Статус рассмотрения / согласования, Подсказки, Сводка | `AI_rd_catalog_pipeline_status.mdc` |
| Робот МТО, origin, accept, SQLite, pipeline | `AI_rd_catalog_db.mdc` |
| Авто МТО, сверка, база заказчика | `AI_rd_catalog_customer_pi.mdc` |
| Зависание, полный refresh, «открыть папку» ничего не делает | `AI_rd_catalog_perf.mdc` |
| Официальный пакет, выгрузка комплектов | `AI_rd_catalog_handoff_export.mdc` |
| WEB-монитор | `AI_rd_catalog_web.mdc` |

Do not load the rest «на всякий случай».

## Answer

1. Say what the cell is showing and which rule makes it so.
2. If the user points at a folder, date, or Google cell, compare that evidence to the snapshot. UNC paths: check `repr` for hyphen U+2010 versus ASCII `-`.
3. A «почему» question ends with the explanation. Do not edit files, rules, or git.
4. If the user asked to change behavior: smallest diff, then update the matching rule in the same change. One-kit GUI actions stay on the scoped refresh path in `AI_rd_catalog_perf.mdc`.

## Example

Dump says `MTO · рев. = нет`, user says the xlsx is in the issued folder. Answer whether the scan skipped it (name, canonical path, `skip_dirs`) or the matrix is looking at another package. Change the scanner only if they asked to start showing that file.
