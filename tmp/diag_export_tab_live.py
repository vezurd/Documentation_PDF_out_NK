"""Drive the real «Ревизии MTO» cockpit with real catalog data, offscreen.

No worker process is started: selections and pool status are injected, which
exercises the third column, the gate and the visible-row batch exactly as the
window would after a comparison pass.
"""

from __future__ import annotations

import os
import sys
from collections import Counter
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from PySide6.QtWidgets import QApplication  # noqa: E402

from rd_catalog.config import load_config  # noqa: E402
from rd_catalog.db import CatalogDatabase  # noqa: E402
from rd_catalog.models import SourceKind  # noqa: E402
from rd_catalog.mto_export import (  # noqa: E402
    default_robot_export_target,
    resolve_export_selections,
    target_files_from_records,
)
from rd_catalog.mto_pair_compare import (  # noqa: E402
    PairPoolStatus,
    pair_pool_status,
    pairs_from_export_selections,
)
from rd_catalog.pipeline import list_revision_matrix  # noqa: E402
from rd_catalog.status_colors import load_status_colors  # noqa: E402
from rd_catalog.revision_matrix_tab import RevisionMatrixTab  # noqa: E402

config = load_config()
database = CatalogDatabase(ROOT / "tmp" / "rd_catalog_copy.sqlite3")
database.initialize()
records = database.list_files()
overlay_ids = {
    int(row["file_id"]) for row in database.current_overlay() if row.get("detected_current")
}

app = QApplication.instance() or QApplication([])
tab = RevisionMatrixTab()
tab.bind_catalog(config, database)
tab.set_cells(
    list_revision_matrix(database),
    palette=load_status_colors(Path(config.runtime_dir) / "status_colors.json"),
    is_banned=lambda _title, _mark: False,
)

target = default_robot_export_target(config)
target_files = target_files_from_records(
    [r for r in records if r.source is SourceKind.ROBOT and r.present]
)
verdict_rows = database.list_mto_pair_comparisons  # noqa: F841  (kept for clarity)

selections = resolve_export_selections(
    database,
    records=records,
    detected_current_ids=overlay_ids,
    rule=tab.current_rule(),
    target=target,
    target_files=target_files,
    pins=(),
    rd_root=config.rd_root,
)
tab.apply_export_selections(selections)

pairs = pairs_from_export_selections(selections)
status = pair_pool_status(database, pairs=pairs, records=records)
tab.set_pool_status(status)

table = tab.table()
print(f"правило по умолчанию : {tab.current_rule()}")
print(f"папка                : {target.name}")
print(f"строк в таблице      : {table.rowCount()}, столбцов: {table.columnCount()}")
print(f"заголовки 0..3       : " + " | ".join(
    table.horizontalHeaderItem(i).text() for i in range(4)
))
print()

texts = Counter()
for row in range(table.rowCount()):
    item = table.item(row, 2)
    texts[item.text() if item is not None else "<пусто>"] += 1
print("третий столбец:")
for text, count in texts.most_common():
    print(f"   {count:4d}  {text}")
print()

print("счётчик пула :", tab._pool_label.text())
print("кнопка       :", tab._export_button.text())
print("доступна     :", tab._export_button.isEnabled())
print("подсказка    :", (tab._export_button.toolTip() or "").strip()[:110])
print()

visible = tab.visible_export_selections()
print(f"видимых строк {tab._visible_row_count()}, выборок в партии {len(visible)}")

tab._filter.setText("1600")
tab._apply_row_visibility()
tab._update_export_button()
visible2 = tab.visible_export_selections()
kits = sorted({f"{s.title}/{s.mark}" for s in visible2})
print(f"после фильтра «1600»: видимых {tab._visible_row_count()}, в партии {len(visible2)}")
print("   ", ", ".join(kits[:12]))
print()

drained = PairPoolStatus(
    total=status.total,
    compared=status.total - 2,
    pending=0,
    failed=2,
    unpersisted=0,
)
tab.set_pool_status(drained)
print("при полностью посчитанном пуле с 2 ошибками:")
print("   счётчик :", tab._pool_label.text())
print("   доступна:", tab._export_button.isEnabled())
