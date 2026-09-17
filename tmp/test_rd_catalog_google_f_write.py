"""Local checks for KSB ИД F/D/E write planning (no network)."""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from rd_catalog.config import CatalogConfig
from rd_catalog.f_journal import format_history_line
from rd_catalog.google_f_write import (
    GoogleWriteError,
    JournalWriteJob,
    _oauth_error_text,
    _requests_session,
    a1_range,
    assert_pin_matches,
    execute_journal_write,
    execute_journal_writes,
    is_ksb_id_header,
    locate_kit_row,
    pin_ksb_id_worksheet,
    pin_worksheet,
    plan_journal_write,
)
from rd_catalog.google_kits import kits_tls_relaxed

_SOT_F = (
    "03.05.2024 код А на рев. 02 AGCC-BCC-TRM-000339\n"
    "21.10.2024 код B\n"
    "02.12.2024 код А на рев. 04 AGCC-BCC-TRM-000582"
)

_HEADER = ["титул", "ИД/ТО/МДЗ/РД", "Наименование", "Тип/ Стадия", "Статус", "Комментарий"]


class MemorySheets:
    """In-memory SheetsClient for locate/pin/write tests."""

    def __init__(
        self,
        rows: list[list[str]],
        *,
        title: str = "Лист1",
        sheet_id: int = 42,
        spreadsheet_id: str = "sheet-id",
        rename_after: int = 0,
    ) -> None:
        self.rows = [list(row) for row in rows]
        self.title = title
        self.sheet_id = sheet_id
        self.spreadsheet_id = spreadsheet_id
        self.rename_after = rename_after
        self.get_count = 0
        self.update_calls = 0
        self.updates: list[tuple[str, list[list[str]]]] = []

    def _row(self, row_index: int) -> list[str]:
        while len(self.rows) < row_index:
            self.rows.append([])
        return self.rows[row_index - 1]

    def get_spreadsheet(self, spreadsheet_id: str) -> dict[str, Any]:
        self.get_count += 1
        title = self.title
        if self.rename_after and self.get_count > self.rename_after:
            title = self.title + " (копия)"
        return {
            "spreadsheetId": spreadsheet_id,
            "sheets": [
                {
                    "properties": {
                        "sheetId": self.sheet_id,
                        "title": title,
                        "index": 0,
                    }
                }
            ],
        }

    def get_values(self, spreadsheet_id: str, a1: str) -> list[list[str]]:
        del spreadsheet_id
        column, row_index = _parse_a1_cell(a1)
        if column is None:
            return [list(row) for row in self.rows]
        row = self._row(row_index)
        col = _COLUMN_INDEX[column]
        value = row[col] if col < len(row) else ""
        return [[value]]

    def update_values(
        self,
        spreadsheet_id: str,
        data: list[tuple[str, list[list[str]]]],
    ) -> None:
        del spreadsheet_id
        self.update_calls += 1
        for a1, values in data:
            self.updates.append((a1, values))
            column, row_index = _parse_a1_cell(a1)
            if column is None:
                continue
            row = self._row(row_index)
            col = _COLUMN_INDEX[column]
            while len(row) <= col:
                row.append("")
            row[col] = values[0][0] if values and values[0] else ""


_COLUMN_INDEX = {"A": 0, "B": 1, "C": 2, "D": 3, "E": 4, "F": 5}


def _parse_a1_cell(a1: str) -> tuple[str | None, int]:
    match = re.search(r"!([A-Z]+)(\d+)$", a1)
    if not match:
        return None, 0
    return match.group(1), int(match.group(2))


def _config() -> CatalogConfig:
    root = Path("unused")
    return CatalogConfig(
        rd_root=root,
        sq_root=root,
        robot_root=root,
        runtime_dir=root,
        db_path=root / "x.sqlite",
        robot_flat_structure=False,
        skip_dirs=(),
        google_kits_spreadsheet_id="sheet-id",
        google_kits_sheet_name="",
    )


def main() -> None:
    """Plan and execute F writes against an in-memory sheet."""

    rows = [
        _HEADER,
        ["1600", "SOT", "x", "Рев. 03", "старый", _SOT_F],
        ["2225", "KSB", "x", "Рев. 01", "ok", "01.01.2024 код А"],
        ["1600", "SOT_IFC", "x", "Рев. 04", "РД Согласовано", _SOT_F],
    ]
    located = locate_kit_row(rows, "1600", "SOT")
    assert located is not None
    assert located.row_index == 4, located.row_index
    assert locate_kit_row(rows, "9999", "SOT") is None

    pin = pin_worksheet(
        {
            "spreadsheetId": "sheet-id",
            "sheets": [
                {"properties": {"sheetId": 7, "title": "Архив", "index": 1}},
                {"properties": {"sheetId": 1, "title": "КСБ ИД", "index": 0}},
            ],
        }
    )
    assert pin.title == "КСБ ИД"
    assert pin.sheet_id == 1
    named = pin_worksheet(
        {
            "spreadsheetId": "sheet-id",
            "sheets": [
                {"properties": {"sheetId": 1, "title": "КСБ ИД", "index": 0}},
                {"properties": {"sheetId": 2, "title": "Нужный", "index": 1}},
            ],
        },
        "Нужный",
    )
    assert named.sheet_id == 2
    try:
        pin_worksheet(
            {
                "spreadsheetId": "sheet-id",
                "sheets": [
                    {"properties": {"sheetId": 1, "title": "КСБ ИД", "index": 0}},
                ],
            },
            "Нет такого",
        )
        raise AssertionError("missing sheet must fail")
    except GoogleWriteError:
        pass

    assert is_ksb_id_header(["титул", "ИД/ТО/МДЗ/РД"]) is True
    assert is_ksb_id_header(["ТИТУЛ", "АКТИВНОСТЬ"]) is False
    assert is_ksb_id_header(["", "№ п/п"]) is False

    spaced_meta = {
        "spreadsheetId": "sheet-id",
        "sheets": [
            {
                "properties": {
                    "sheetId": 9,
                    "title": "Контроль выдачи ",
                    "index": 0,
                }
            }
        ],
    }
    spaced = pin_worksheet(spaced_meta)
    assert spaced.title == "Контроль выдачи "
    by_stripped_name = pin_worksheet(spaced_meta, "Контроль выдачи")
    assert by_stripped_name.sheet_id == 9
    assert by_stripped_name.title == "Контроль выдачи "

    class TwoSheetClient:
        """First tab is a schedule; kits live on the second tab."""

        def __init__(self) -> None:
            self.schedule = [["", "№ п/п"], ["", "1"]]
            self.kits = [list(row) for row in rows]
            self.updates: list[Any] = []

        def get_spreadsheet(self, spreadsheet_id: str) -> dict[str, Any]:
            return {
                "spreadsheetId": spreadsheet_id,
                "sheets": [
                    {
                        "properties": {
                            "sheetId": 1,
                            "title": "График2",
                            "index": 0,
                        }
                    },
                    {
                        "properties": {
                            "sheetId": 2,
                            "title": "Контроль выдачи ",
                            "index": 1,
                        }
                    },
                ],
            }

        def get_values(self, spreadsheet_id: str, a1: str) -> list[list[str]]:
            del spreadsheet_id
            table = self.schedule if "График2" in a1 else self.kits
            column, row_index = _parse_a1_cell(a1)
            if column is None:
                return [list(row) for row in table]
            row = table[row_index - 1]
            col = _COLUMN_INDEX[column]
            value = row[col] if col < len(row) else ""
            return [[value]]

        def update_values(
            self,
            spreadsheet_id: str,
            data: list[tuple[str, list[list[str]]]],
        ) -> None:
            del spreadsheet_id
            self.updates.extend(data)

    two = TwoSheetClient()
    discovered = pin_ksb_id_worksheet(
        two, two.get_spreadsheet("sheet-id"), spreadsheet_id="sheet-id"
    )
    assert discovered.title == "Контроль выдачи "
    assert discovered.sheet_id == 2

    job = JournalWriteJob(
        title="1600",
        mark="SOT",
        f_line=format_history_line(
            date="09.09.2026",
            stage="code_a",
            revision="04",
            transmittal="AGCC-BCC-TRM-000999",
        ),
        revision="04",
        stage="code_a",
    )
    client = MemorySheets(rows)
    plan = plan_journal_write(
        client.rows,
        pin_worksheet(client.get_spreadsheet("sheet-id"), spreadsheet_id="sheet-id"),
        job,
    )
    assert plan.located.row_index == 4
    assert plan.patch.update_de is True
    assert plan.updates[0][0] == a1_range("Лист1", "F4")
    assert len(plan.updates) == 1

    old_job = JournalWriteJob(
        title="1600",
        mark="SOT",
        f_line=format_history_line(
            date="09.10.2023",
            stage="code_a",
            revision="0",
            transmittal="AGCC-BCC-TRM-000136",
        ),
        revision="0",
        stage="code_a",
    )
    old_plan = plan_journal_write(
        rows,
        pin_worksheet(client.get_spreadsheet("sheet-id"), spreadsheet_id="sheet-id"),
        old_job,
    )
    assert old_plan.patch.update_de is False
    assert len(old_plan.updates) == 1
    assert old_plan.updates[0][0].endswith("F4")

    result = execute_journal_write(client, _config(), job)
    assert result.error == "", result.error
    assert result.row_index == 4
    assert result.update_de is False
    assert client.rows[3][5].split("\n")[-1] == job.f_line
    assert client.rows[3][3] == "Рев. 04"
    assert client.rows[3][4] == "РД Согласовано"
    assert client.rows[1][5] == _SOT_F

    already_plan = plan_journal_write(
        client.rows,
        pin_worksheet(client.get_spreadsheet("sheet-id"), spreadsheet_id="sheet-id"),
        job,
    )
    assert already_plan.patch.action == "unchanged"
    assert already_plan.updates == ()
    wrote_n = len(client.updates)
    again = execute_journal_write(client, _config(), job)
    assert again.error == ""
    assert again.update_de is False
    assert len(client.updates) == wrote_n

    f_after = client.rows[3][5]
    stale = MemorySheets(
        [_HEADER, ["1600", "SOT", "x", "Рев. 03", "старый", f_after]]
    )
    stale_pin = pin_worksheet(
        stale.get_spreadsheet("sheet-id"), spreadsheet_id="sheet-id"
    )
    stale_plan = plan_journal_write(stale.rows, stale_pin, job)
    assert stale_plan.patch.action == "unchanged"
    assert not any(item[0].endswith("F2") for item in stale_plan.updates)
    assert any(item[0].endswith("D2") for item in stale_plan.updates)
    assert any(item[0].endswith("E2") for item in stale_plan.updates)
    stale_result = execute_journal_write(stale, _config(), job)
    assert stale_result.error == ""
    assert stale_result.update_de is True
    assert stale.rows[1][5] == f_after
    assert stale.rows[1][3] == "Рев. 04"
    assert stale.rows[1][4] == "РД Согласовано"

    missing_client = MemorySheets(rows)
    missing = execute_journal_write(
        missing_client,
        _config(),
        JournalWriteJob(
            title="8888",
            mark="SOT",
            f_line=job.f_line,
            revision="04",
            stage="code_a",
        ),
    )
    assert "Нет строки комплекта" in missing.error
    assert not missing_client.updates

    renamed = MemorySheets(rows, rename_after=1)
    renamed_ok = execute_journal_write(renamed, _config(), job)
    assert renamed_ok.error == "", renamed_ok.error
    assert renamed.get_count == 1
    assert renamed.updates

    job_ksb = JournalWriteJob(
        title="2225",
        mark="KSB",
        f_line=job.f_line,
        revision="04",
        stage="code_a",
    )
    batch_client = MemorySheets(rows)
    batch_results = execute_journal_writes(batch_client, _config(), (job, job_ksb))
    assert batch_client.get_count == 1
    assert batch_client.update_calls == 2
    assert len(batch_results) == 2
    assert all(not item.error for item in batch_results), batch_results

    same_kit_rows = [
        _HEADER,
        ["1600", "SOT", "x", "Рев. 03", "старый", _SOT_F],
    ]
    same_kit_pin = pin_worksheet(
        {
            "spreadsheetId": "sheet-id",
            "sheets": [
                {"properties": {"sheetId": 42, "title": "Лист1", "index": 0}}
            ],
        },
        spreadsheet_id="sheet-id",
    )
    first_plan = plan_journal_write(same_kit_rows, same_kit_pin, job)
    first_d = next(value for a1, value in first_plan.updates if a1.endswith("D2"))
    first_e = next(value for a1, value in first_plan.updates if a1.endswith("E2"))
    job_sot_2 = JournalWriteJob(
        title="1600",
        mark="SOT",
        f_line=format_history_line(
            date="10.09.2026",
            stage="code_a",
            revision="04",
            transmittal="AGCC-BCC-TRM-001000",
        ),
        revision="04",
        stage="code_a",
    )
    job_sot_3 = JournalWriteJob(
        title="1600",
        mark="SOT",
        f_line=format_history_line(
            date="11.09.2026",
            stage="code_a",
            revision="04",
            transmittal="AGCC-BCC-TRM-001001",
        ),
        revision="04",
        stage="code_a",
    )
    job_sot_4 = JournalWriteJob(
        title="1600",
        mark="SOT",
        f_line=format_history_line(
            date="12.09.2026",
            stage="code_a",
            revision="04",
            transmittal="AGCC-BCC-TRM-001002",
        ),
        revision="04",
        stage="code_a",
    )
    same_kit_jobs = (job, job_sot_2, job_sot_3, job_sot_4)
    same_kit_client = MemorySheets(same_kit_rows)
    same_kit_results = execute_journal_writes(
        same_kit_client, _config(), same_kit_jobs
    )
    assert same_kit_client.get_count == 1
    assert same_kit_client.update_calls == 1
    assert len(same_kit_results) == len(same_kit_jobs)
    assert all(item.error == "" for item in same_kit_results), same_kit_results
    assert same_kit_client.rows[1][5] == same_kit_results[-1].comment_after
    assert same_kit_client.rows[1][3] == first_d
    assert same_kit_client.rows[1][4] == first_e
    assert same_kit_results[0].update_de is True
    assert same_kit_results[0].comment_after != same_kit_results[-1].comment_after

    meta = {
        "spreadsheetId": "sheet-id",
        "sheets": [{"properties": {"sheetId": 42, "title": "Лист1", "index": 0}}],
    }
    good_pin = pin_worksheet(meta, spreadsheet_id="sheet-id")
    assert_pin_matches(meta, good_pin)
    try:
        assert_pin_matches(
            {
                "sheets": [
                    {"properties": {"sheetId": 99, "title": "Лист1", "index": 0}}
                ]
            },
            good_pin,
        )
        raise AssertionError("gid mismatch must fail")
    except GoogleWriteError:
        pass

    saved_strict = os.environ.pop("RD_KITS_SSL_STRICT", None)
    try:
        assert kits_tls_relaxed() is True
        relaxed = _requests_session(verify=not kits_tls_relaxed())
        assert relaxed.verify is False
        relaxed.close()
        os.environ["RD_KITS_SSL_STRICT"] = "1"
        assert kits_tls_relaxed() is False
        strict = _requests_session(verify=not kits_tls_relaxed())
        assert strict.verify is True
        strict.close()
    finally:
        if saved_strict is None:
            os.environ.pop("RD_KITS_SSL_STRICT", None)
        else:
            os.environ["RD_KITS_SSL_STRICT"] = saved_strict

    ssl_msg = _oauth_error_text(
        RuntimeError(
            "SSLCertVerificationError: certificate verify failed: "
            "unable to get local issuer certificate"
        )
    )
    assert ssl_msg.startswith("OAuth Google: не прошла проверка TLS")
    assert _oauth_error_text(TimeoutError("timed out")).startswith("OAuth Google:")

    print("google_f_write ok")


if __name__ == "__main__":
    main()
