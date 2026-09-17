# Handoff: GUI для сбора RFP из частей (`ds_compare_center`)

Дата: 2026-07-23.

> Домен: [`.cursor/rules/AI_rfp_parts.mdc`](../.cursor/rules/AI_rfp_parts.mdc).

## Запуск

| | |
|---|---|
| Приложение | `python -m ds_compare_center --tab rfp_parts` |
| Analyze | `python -X utf8 -m RFQ.rfp_parts --out-dir …` |

## UI

- Read-only пути increase/decrease/summary + `[OK]`/`[NO]`
- Проанализировать / Открыть report.txt / Открыть папку
- **Без** manifest / strict / write-manifest (пауза → позже список ДС)

## Verification

```powershell
python -m py_compile ds_compare_center/rfp_parts_panel.py ds_compare_center/center_window.py
python -m ds_compare_center --tab rfp_parts
```
