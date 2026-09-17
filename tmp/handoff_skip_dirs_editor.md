# Spec: skip-dirs editor + fast DB prune (RD catalog)

Curator spec. Implement exactly this. Do not write UNC sources. Do not edit `rd_catalog/default_config.json` from GUI. Do not create a fake scan_run. Do not mark pruned files PENDING.

## Goal

On tab **Все документы**:

1. Button **Skip-папки…** opens a dialog: list all skip tokens, add / edit / delete / move up-down / sort A→Я / restore factory list.
2. Button **Убрать skip из дерева** walks **SQLite paths only** (no `os.walk`, no UNC). Marks matching RD+SQ present files absent, rebuilds overlay, refreshes GUI.

Tree must hide skip-matching paths immediately after the skip list is saved (GUI filter), even before prune, because `present=0` rows still appear in the tree today.

## Shared skip matching

New module `rd_catalog/skip_dirs.py` (Qt-free). Move matching here so scan and prune cannot diverge.

```python
def is_skipped_dir_name(name: str, skip_dirs: Iterable[str]) -> bool:
    folded = name.casefold()
    return any(skip.casefold() in folded for skip in skip_dirs)

def path_has_skipped_dir(path: str | Path, skip_dirs: Iterable[str]) -> bool:
    """True if any *parent directory* name matches skip_dirs.

    Do not test the filename itself. Empty skip_dirs → False.
    """
```

`rd_catalog/scan.py` `_is_skipped_dir` must call `is_skipped_dir_name` (thin wrapper or replace call sites). Keep the name `_is_skipped_dir` if tests import it.

Gate folders must never be skippable. Reject a token if any of these names would match:

- `Для передачи`
- `Для_передачи`
- `На_отправку`
- `На_передачу`

`validate_skip_token(token: str, existing: Iterable[str]) -> str` returns stripped token or raises `ValueError` with a Russian message. Reject empty. Reject duplicate (casefold). Do not reject short tokens (`3D`, `old` are factory).

## Persistence

File: `{runtime_dir}/skip_dirs.json` (same dir as `banned_title_marks.json`). Never UNC.

```json
{"version": 1, "tokens": ["old", "temp", "замечания"]}
```

`SkipDirsStore` (mirror `ban_filter.BanFilterStore` style):

- `from_runtime_dir(runtime_dir, packaged: tuple[str, ...])`
- If file missing: in-memory tokens = `packaged` (from current `CatalogConfig.skip_dirs` / packaged defaults). Do not write until user saves.
- `save()` writes UTF-8 JSON. Preserve order.
- `set_tokens(tokens)`, `add`, `replace_at`, `remove_at`, `move(index, delta)`, `sort_casefold()`
- Missing/corrupt `skip_dirs.json` is rewritten from packaged `default_config.json` (no GUI factory reset).
- `packaged_skip_dirs()` reads `rd_catalog/default_config.json` `skip_dirs`

`load_config` after resolving `runtime_dir`:

- If `{runtime_dir}/skip_dirs.json` exists and has a non-empty valid `tokens` list, **replace** `skip_dirs` with that list.
- If file missing, keep packaged / `--config` override list.
- `config_from_job_dict` unchanged (job already has `skip_dirs`).

`CatalogConfig` stays frozen. Window updates via `dataclasses.replace(self.config, skip_dirs=tuple(store.tokens()))`.

Factory defaults stay in `rd_catalog/default_config.json`. GUI never writes that file.

## DB prune (no UNC)

`CatalogDatabase.apply_skip_dirs(skip_dirs: Iterable[str]) -> int` in `rd_catalog/db.py`:

1. One transaction.
2. SELECT `id, path, source` FROM `file_entry` WHERE `present = 1` AND `source IN ('rd', 'sq')`.
3. If `path_has_skipped_dir(path, skip_dirs)`: `UPDATE file_entry SET present = 0 WHERE id = ?`. **Do not** change `review_state`. **Do not** insert `review_event`.
4. Rebuild overlay from remaining present RD parsed files, same as subtree overlay path: `build_rd_overlays(_present_rd_parsed_files(...), ignored_path_keys=_ignored_path_keys(...))`, then replace `overlay_state` rows. Use last successful `scan_run.id` if any, else `0`.
5. Return count of rows marked absent.

Robot files untouched.

Also add `count_present_skipped(skip_dirs) -> int` for the confirm dialog (same match, no write).

## GUI

Pattern: `rd_catalog/ban_dialog.py` + documents-tab filter row in `rd_catalog/window.py` (`_build_documents_tab`).

`rd_catalog/skip_dirs_dialog.py`:

- Title: `Skip-папки`
- Hint: substring match, case-insensitive, like scan prune; factory list is a seed; save writes `%LOCALAPPDATA%\...\skip_dirs.json`; prune button does not walk UNC.
- `QListWidget` of tokens, single-column, internal drag optional but **must** have Вверх / Вниз.
- Buttons: Добавить, Изменить, Удалить, Вверх, Вниз, А→Я.
- Add/edit: `QInputDialog.getText`. On `ValueError` show `QMessageBox.warning`.
- Reset factory: confirm, then `reset_to_packaged`.
- `changed = Signal()` after save.
- OK/Save writes JSON and emits `changed`. Cancel discards unsaved edits (reload from store on open, edit a copy in the dialog).

Documents tab filter row (right of checkboxes):

- `Skip-папки…` → open dialog. On `changed`: `replace` config skip_dirs, rebuild tree.
- `Убрать skip из дерева` → if `_busy()`: same info box as other workers. Else count matches; if 0: information «Нет файлов в базе под текущий skip». Else confirm `Убрать N файлов из каталога без скана сети? Файлы на диске не трогаем. Вернуть можно только полным сканом РД/SQ.` Then `database.apply_skip_dirs`, `refresh()`, log line with N.

Tree `_rebuild_document_tree`: skip any record where `path_has_skipped_dir(record.path, self.config.skip_dirs)`. Do this before grouping so skipped files never form a node.

`_set_workers_enabled`: also disable the prune button while busy.

## Tests

`tmp/test_rd_catalog_skip_dirs.py` (no Qt):

- `is_skipped_dir_name("замечания", ["Замечания"])` True
- `path_has_skipped_dir` True for a PDF under `...\замечания\01\file.pdf`, False when only the filename contains the token
- `validate_skip_token("передач", [])` raises
- store save/load roundtrip in temp dir
- `load_config` with temp LOCALAPPDATA + written `skip_dirs.json` picks user tokens
- temp sqlite: insert two present RD files (one under `замечания`, one under `Для передачи`); `apply_skip_dirs(["Замечания"])` returns 1; skipped `present=0`; kept present; overlay current is the kept file if both are PDF with keys

Keep existing `tmp/test_rd_catalog_parse_smoke.py` green (`_is_skipped_dir` still works).

`tmp/test_rd_catalog_gui_smoke.py`: after window create, both new buttons exist; opening `SkipDirsDialog` offscreen and closing does not crash. Optional: stub a record under `\замечания\` and assert tree hides it when `config.skip_dirs` contains `Замечания`.

## Out of scope

- Editing `default_config.json` from UI
- Auto-merge new factory tokens into an existing user JSON
- Restoring pruned files without a full RD/SQ scan
- Robot skip list
- Changing scan missing-detection / PENDING semantics

## Verification

```
python tmp/test_rd_catalog_skip_dirs.py
python tmp/test_rd_catalog_parse_smoke.py
python tmp/test_rd_catalog_gui_smoke.py
python tmp/test_rd_catalog_source_files.py
```

All print OK, exit 0.

Style: `from __future__ import annotations`, Google docstrings on public functions, `str | None`, English in `rd_catalog/`.
