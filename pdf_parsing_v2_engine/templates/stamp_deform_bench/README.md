# Paired Stamp Deform Benchmark

Рабочая папка для генерации деформированных PDF-штампов в `paired mode`:
- `input/stamped/` — PDF со штампом
- `input/background/` — matching PDF без штампа
- `template/` — один JSON-шаблон v2
- `output/generated/` — сгенерированные PDF
- `output/manifests/` — `manifest.jsonl` и `manifest.csv`

Что сейчас лежит в папке:
- пара MTO-файлов `AGCC.287-2869-SOS.MTO-0001_01-AN02_RU.pdf`
- шаблон `template/mto_page1.json`

Основной локальный запуск:

```bash
python pdf_parsing_v2_engine/templates/stamp_deform_bench/run_stamp_deform_bench.py
```

Прогон benchmark по уже сгенерированным PDF:

```bash
python pdf_parsing_v2_engine/templates/stamp_deform_bench/run_alignment_benchmark.py
```

Чтобы не перезаписывать прошлый прогон, можно задать label через env:

```bash
set STAMP_ALIGN_RUN_LABEL=after_top_pref
python pdf_parsing_v2_engine/templates/stamp_deform_bench/run_alignment_benchmark.py
```

В `run_stamp_deform_bench.py` можно менять все основные параметры:
- `STAMPED_INPUT`, `BACKGROUND_INPUT`, `TEMPLATE_PATH`
- `PAGE_NUM`
- `X_RANGE`, `Y_RANGE`
- `STEP`
- `MODE`
- `ANCHOR`
- `DRY_RUN`
- `CFG`

Рекомендуемый первый запуск:
- оставить `DRY_RUN = True`
- `MODE = "cross"` или `MODE = "diagonal"`
- `STEP = 5` или `STEP = 10`

После проверки поменять:
- `DRY_RUN = False`

Правило сопоставления при folder-mode:
- относительные пути внутри `input/stamped/` и `input/background/` должны совпадать.
- пример: `input/stamped/DWG/a.pdf` <-> `input/background/DWG/a.pdf`

Прямой CLI тоже доступен:

```bash
python pdf_parsing_v2_engine/tools/generate_deformed_stamp_pdfs.py ^
  --input "pdf_parsing_v2_engine/templates/stamp_deform_bench/input/stamped" ^
  --background-input "pdf_parsing_v2_engine/templates/stamp_deform_bench/input/background" ^
  --template "pdf_parsing_v2_engine/templates/stamp_deform_bench/template/mto_page1.json" ^
  --mode cross ^
  --step 5 ^
  --dry-run
```

Что пишет `run_alignment_benchmark.py`:
- `output/benchmark_runs/<run_label>/summary.csv` — одна строка на generated PDF
- `output/benchmark_runs/<run_label>/summary.jsonl` — тот же summary для diff/скриптов
- `output/benchmark_runs/<run_label>/decision_summary.md` — краткая сводка по verdict и гипотезе `h_top`
- `output/benchmark_runs/<run_label>/alignment_logs/` — alignment log `.txt` по каждому PDF
- `output/benchmark_runs/<run_label>/cases/` — raw JSON с transform, detected lines и diagnostics

Минимальные метрики в summary:
- `sx_percent`, `sy_percent`
- `alignment_verdict`
- `boundary_matched`
- `template_skipped`
- `avg_match_score`
- `h_top_matched_to`
- `h_top_window_top_candidate_id`
- `likely_top_boundary_issue`
