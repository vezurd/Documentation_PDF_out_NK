"""Normcontrol rows built from ОД postprocess ``warnings`` (``od_table['warnings']``)."""

from __future__ import annotations

import re
from os.path import basename

from pdf_parsing_v2_rules.output import CheckRow

_OD_CODE_DESCRIPTIONS: dict[int, str] = {
    3001: "ОД: в поле «Формат» разделители — пробелы/переносы без запятой (нормализовано)",
    3002: "ОД: неизвестный ключ формата листа",
    3099: "ОД: предупреждение при разборе ведомости",
}

_WARN_LINE = re.compile(r"^(\d{4}):\s*(.*)\Z", re.DOTALL)
_LEGACY_OD_TIER = re.compile(r"^OD_(\d{4})\s+tier=[^:]+:\s*(.*)\Z", re.DOTALL)


def build_od_table_warning_rows(
    od_warnings: list[str] | None,
    *,
    od_pdf_path: str = "",
) -> list[CheckRow]:
    """Map ``od_table['warnings']`` strings (``\"3001: …\"``) to ``CheckRow`` for НК / xlsx."""
    doc = basename(str(od_pdf_path).strip()) if str(od_pdf_path).strip() else "ОД (ведомость)"
    pdf = str(od_pdf_path).strip()
    rows: list[CheckRow] = []
    for raw in od_warnings or []:
        line = str(raw).strip()
        if not line:
            continue
        m = _WARN_LINE.match(line)
        if m:
            code = int(m.group(1))
            text = (m.group(2) or "").strip()
        else:
            leg = _LEGACY_OD_TIER.match(line)
            if leg:
                code = int(leg.group(1))
                text = (leg.group(2) or "").strip()
            else:
                code = 3099
                text = line
        c_desc = _OD_CODE_DESCRIPTIONS.get(code, _OD_CODE_DESCRIPTIONS[3099])
        rows.append(
            CheckRow(
                result=False,
                c_code=code,
                c_description=c_desc,
                doc_name=doc,
                page_num="нет",
                text=text,
                pdf_path=pdf,
            )
        )
    return rows
