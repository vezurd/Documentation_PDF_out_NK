"""Smoke: critical remarks linkify relative xlsx paths inside brackets."""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from ds_compare_center.tsd_packing_panel import (  # noqa: E402
    _critical_detail_to_html,
    _find_xlsx_path_spans,
    _resolve_xlsx_path,
)


def main() -> None:
    root = r"\\bcc\eng\PrDoc\TSD"
    rel = (
        r"согл УЛ ГФ 2 тит\AGCC.323-2000 - "
        r"2064203.2-M15-AN006580FR_ГОС.ФИН_01_EN.xlsx"
    )
    line = (
        "1. ERROR source_required_fields "
        f"[{rel} · SO - PL · строка 2] field=name: Похоже на позицию"
    )
    spans = _find_xlsx_path_spans(line)
    assert spans, "expected bracket path span"
    assert spans[0][2] == rel, spans[0][2]

    resolved = _resolve_xlsx_path(rel, root)
    assert resolved.endswith(rel.replace("/", "\\")) or rel in resolved, resolved

    html = _critical_detail_to_html(line, root=root)
    assert "<a href=" in html, html
    match = re.search(r'<a href="([^"]+)">([^<]+)</a>', html)
    assert match is not None, html
    assert "file:" in match.group(1), match.group(1)
    assert match.group(2) == rel, match.group(2)

    abs_line = r"\\bcc\eng\a\b.xlsx"
    assert "<a href=" in _critical_detail_to_html(abs_line, root=root)

    with_root = _critical_detail_to_html(f"root: {root}\n\n{line}", root="")
    assert with_root.count("<a href=") >= 2, with_root

    print("OK critical links smoke")


if __name__ == "__main__":
    main()
