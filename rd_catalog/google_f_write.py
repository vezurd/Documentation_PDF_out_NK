"""Write KSB ИД columns F (and optionally D/E) via the Sheets API.

Qt-free. Locate the kit by live columns A+B, never by a cached ``row_index``.
Pin the kits worksheet (header ``титул``, not the leftmost tab) before
``values.batchUpdate``. Persist gid+title to ``google_sheet_pins.json``
when ``runtime_dir`` already exists (cell-jump URLs). Does not create
rows and does not touch «Выдача РД ПД».
"""

from __future__ import annotations

import json
import os
import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Protocol, Sequence
from urllib.parse import quote

from rd_catalog.config import CatalogConfig
from rd_catalog.f_journal import JournalPatch, build_journal_patch, journal_write_needed
from rd_catalog.google_kits import kits_tls_relaxed
from rd_catalog.google_sheet_links import save_kits_sheet_pin
from rd_catalog.kits import GoogleKit, kit_identity_key, parse_google_kit_row

DEFAULT_SERVICE_ACCOUNT_PATH = (
    Path(__file__).resolve().parents[1] / "base_check" / "client_secret.json"
)
_SHEETS_API = "https://sheets.googleapis.com/v4"
_SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets"
_HEADER_TITLES = {"титул", "title"}
_A1_CELL_RE = re.compile(r"!([A-Z]+)(\d+)$")


class GoogleWriteError(RuntimeError):
    """User-facing failure while planning or writing a KSB ИД cell."""


class SheetsClient(Protocol):
    """Minimal Sheets API surface used by :func:`execute_journal_write`."""

    def get_spreadsheet(self, spreadsheet_id: str) -> dict[str, Any]:
        """Return spreadsheet metadata including sheet gid/title."""

    def get_values(self, spreadsheet_id: str, a1_range: str) -> list[list[str]]:
        """Return a value range as a string matrix."""

    def update_values(
        self,
        spreadsheet_id: str,
        data: Sequence[tuple[str, list[list[str]]]],
    ) -> None:
        """Write ranges with RAW input (no formula/date coercion)."""


@dataclass(frozen=True, slots=True)
class WorksheetPin:
    """Gid + title of the worksheet that will receive the write."""

    spreadsheet_id: str
    sheet_id: int
    title: str


@dataclass(frozen=True, slots=True)
class LocatedKitRow:
    """One KSB ИД row found by live title+mark."""

    row_index: int
    kit: GoogleKit


@dataclass(frozen=True, slots=True)
class JournalWriteJob:
    """One F-line write for a title+mark kit."""

    title: str
    mark: str
    f_line: str
    revision: str | None
    stage: str


@dataclass(frozen=True, slots=True)
class GoogleWritePlan:
    """Previewable cell updates computed from the live sheet."""

    pin: WorksheetPin
    located: LocatedKitRow
    patch: JournalPatch
    updates: tuple[tuple[str, str], ...]


@dataclass(frozen=True, slots=True)
class GoogleWriteResult:
    """Outcome of one job after the API write (or a locate/verify error)."""

    title: str
    mark: str
    row_index: int
    sheet_title: str
    comment_before: str
    comment_after: str
    update_de: bool
    error: str = ""


def service_account_path() -> Path:
    """Return the service-account JSON path.

    Returns:
        ``RD_KITS_SERVICE_ACCOUNT`` when set, else the packaged
        ``base_check/client_secret.json``.
    """

    override = os.environ.get("RD_KITS_SERVICE_ACCOUNT", "").strip()
    if override:
        return Path(override)
    return DEFAULT_SERVICE_ACCOUNT_PATH


def a1_range(sheet_title: str, range_body: str) -> str:
    """Build a quoted A1 range on ``sheet_title``.

    Args:
        sheet_title: Worksheet title (may contain spaces or quotes).
        range_body: Range without the sheet prefix, e.g. ``F12`` or ``A:F``.

    Returns:
        ``'Title'!F12`` with quotes escaped.
    """

    escaped = (sheet_title or "").replace("'", "''")
    return f"'{escaped}'!{range_body}"


def pin_worksheet(
    meta: Mapping[str, Any],
    preferred_title: str = "",
    *,
    spreadsheet_id: str = "",
) -> WorksheetPin:
    """Pin the target worksheet by title, else the first sheet (index 0).

    Args:
        meta: ``spreadsheets.get`` payload.
        preferred_title: Configured KSB ИД sheet name; empty means first tab.
            Comparison ignores surrounding whitespace; the pin keeps the
            live title (a trailing space in Google is significant for A1).
        spreadsheet_id: Fallback id when ``meta`` omits ``spreadsheetId``.

    Returns:
        A :class:`WorksheetPin`.

    Raises:
        GoogleWriteError: No sheets, or the named sheet is missing.
    """

    sheets = list(meta.get("sheets") or [])
    if not sheets:
        raise GoogleWriteError("В таблице КСБ ИД нет листов.")
    sid = str(meta.get("spreadsheetId") or spreadsheet_id or "")
    wanted = preferred_title.strip()
    chosen: Mapping[str, Any] | None = None
    if wanted:
        for sheet in sheets:
            props = sheet.get("properties") or {}
            if str(props.get("title") or "").strip() == wanted:
                chosen = props
                break
        if chosen is None:
            raise GoogleWriteError(f"Лист {wanted!r} не найден в таблице КСБ ИД.")
    else:
        for sheet in sheets:
            props = sheet.get("properties") or {}
            if int(props.get("index") or 0) == 0:
                chosen = props
                break
        if chosen is None:
            chosen = sheets[0].get("properties") or {}
    try:
        sheet_id = int(chosen.get("sheetId"))
    except (TypeError, ValueError) as exc:
        raise GoogleWriteError("У листа КСБ ИД нет gid.") from exc
    title = str(chosen.get("title") or "")
    if not title.strip():
        raise GoogleWriteError("У листа КСБ ИД пустое имя.")
    if not sid:
        raise GoogleWriteError("Не задан идентификатор таблицы КСБ ИД.")
    return WorksheetPin(spreadsheet_id=sid, sheet_id=sheet_id, title=title)


def is_ksb_id_header(cells: Sequence[Any]) -> bool:
    """Return whether row 1 looks like the KSB ИД kits sheet.

    Args:
        cells: A1 and B1 (extra cells ignored).

    Returns:
        True when A is ``титул``/``title`` and B looks like the mark column.
    """

    if not cells:
        return False
    title = str(cells[0] or "").strip().casefold()
    if title not in _HEADER_TITLES:
        return False
    mark = str(cells[1] or "").strip().casefold() if len(cells) > 1 else ""
    if not mark:
        return True
    return any(token in mark for token in ("ид", "мдз", "mark", "марка"))


def _sheet_props_by_index(meta: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    items: list[tuple[int, Mapping[str, Any]]] = []
    for sheet in meta.get("sheets") or []:
        props = sheet.get("properties") or {}
        try:
            index = int(props.get("index") or 0)
        except (TypeError, ValueError):
            index = 0
        items.append((index, props))
    items.sort(key=lambda item: item[0])
    return [props for _index, props in items]


def pin_ksb_id_worksheet(
    client: SheetsClient,
    meta: Mapping[str, Any],
    preferred_title: str = "",
    *,
    spreadsheet_id: str = "",
) -> WorksheetPin:
    """Pin the KSB ИД kits tab, not whichever sheet happens to be leftmost.

    Empty ``preferred_title`` probes A1/B1 in tab order and keeps the first
    sheet whose header is the kits table (``титул`` + mark column). The
    workbook's first tab is a schedule (``График2``), not комплекты.

    Args:
        client: Live Sheets client (A1/B1 reads).
        meta: ``spreadsheets.get`` payload.
        preferred_title: Configured sheet name; empty means discover.
        spreadsheet_id: Fallback spreadsheet id.

    Returns:
        A :class:`WorksheetPin` with the live title (including trailing space).

    Raises:
        GoogleWriteError: Named sheet missing, or no kits header found.
    """

    if preferred_title.strip():
        return pin_worksheet(
            meta, preferred_title, spreadsheet_id=spreadsheet_id
        )
    sid = str(meta.get("spreadsheetId") or spreadsheet_id or "")
    props_list = _sheet_props_by_index(meta)
    if not props_list:
        raise GoogleWriteError("В таблице КСБ ИД нет листов.")
    titles = [
        str(props.get("title") or "")
        for props in props_list
        if str(props.get("title") or "").strip()
    ]
    headers = _header_cells_for_sheets(client, sid, titles)
    for title, cells in zip(titles, headers, strict=False):
        if is_ksb_id_header(cells):
            return pin_worksheet(meta, title, spreadsheet_id=sid)
    raise GoogleWriteError(
        "Не найден лист комплектов КСБ ИД (A1=«титул»). "
        "Первый ярлык таблицы — не комплекты. Задайте google_kits_sheet_name."
    )


def assert_pin_matches(meta: Mapping[str, Any], pin: WorksheetPin) -> None:
    """Fail if the pinned gid/title no longer match the live spreadsheet.

    Args:
        meta: Fresh ``spreadsheets.get`` payload.
        pin: Previously captured pin.

    Raises:
        GoogleWriteError: The sheet was renamed, moved, or deleted.
    """

    for sheet in meta.get("sheets") or []:
        props = sheet.get("properties") or {}
        try:
            sheet_id = int(props.get("sheetId"))
        except (TypeError, ValueError):
            continue
        if sheet_id != pin.sheet_id:
            continue
        title = str(props.get("title") or "")
        if title != pin.title:
            raise GoogleWriteError(
                f"Лист gid={pin.sheet_id} переименован: {pin.title!r} → {title!r}."
            )
        return
    raise GoogleWriteError(
        f"Лист «{pin.title}» (gid={pin.sheet_id}) исчез из таблицы КСБ ИД."
    )


def locate_kit_row(
    rows: Sequence[Sequence[Any]],
    title: str,
    mark: str,
) -> LocatedKitRow | None:
    """Find the kit by live A+B; the last duplicate wins.

    Args:
        rows: Sheet values including the header row.
        title: Four-digit title from the letter.
        mark: Latin AGCC mark from the letter.

    Returns:
        The matching row, or ``None`` when the kit is absent (do not create).
    """

    needle = kit_identity_key(title, mark)
    found: LocatedKitRow | None = None
    matrix = [list(row) for row in rows]
    start = 0
    if matrix:
        header = [str(cell or "").strip().casefold() for cell in matrix[0]]
        if header and header[0] in _HEADER_TITLES:
            start = 1
    for offset, row in enumerate(matrix[start:], start=start + 1):
        parsed = parse_google_kit_row([str(cell or "") for cell in row], row_index=offset)
        if isinstance(parsed, str):
            continue
        if kit_identity_key(parsed.title, parsed.mark) == needle:
            found = LocatedKitRow(row_index=offset, kit=parsed)
    return found


def plan_journal_write(
    rows: Sequence[Sequence[Any]],
    pin: WorksheetPin,
    job: JournalWriteJob,
) -> GoogleWritePlan:
    """Build F/D/E cell updates from the live matrix.

    Args:
        rows: Current A:F values.
        pin: Pinned worksheet.
        job: Confirmed letter line.

    Returns:
        A :class:`GoogleWritePlan`.

    Raises:
        GoogleWriteError: Kit row missing or the F line cannot be applied.
    """

    located = locate_kit_row(rows, job.title, job.mark)
    if located is None:
        raise GoogleWriteError(
            f"Нет строки комплекта {job.title}-{job.mark} "
            f"на листе «{pin.title}» КСБ ИД. Новые строки не создаются."
        )
    try:
        patch = build_journal_patch(
            located.kit.comment_raw,
            job.f_line,
            revision=job.revision,
            stage=job.stage,
        )
    except ValueError as exc:
        raise GoogleWriteError(str(exc)) from exc
    write_f, write_d, write_e = journal_write_needed(
        patch,
        live_d=_row_cell(rows, located.row_index, 3),
        live_e=_row_cell(rows, located.row_index, 4),
    )
    updates: list[tuple[str, str]] = []
    if write_f:
        updates.append(
            (a1_range(pin.title, f"F{located.row_index}"), patch.comment_after)
        )
    if write_d and patch.sheet_revision:
        updates.append(
            (a1_range(pin.title, f"D{located.row_index}"), patch.sheet_revision)
        )
    if write_e and patch.status_sheet:
        updates.append(
            (a1_range(pin.title, f"E{located.row_index}"), patch.status_sheet)
        )
    return GoogleWritePlan(pin=pin, located=located, patch=patch, updates=tuple(updates))


def execute_journal_write(
    client: SheetsClient,
    config: CatalogConfig,
    job: JournalWriteJob,
) -> GoogleWriteResult:
    """Write one job. See :func:`execute_journal_writes` for the batch path."""

    results = execute_journal_writes(client, config, (job,))
    if results:
        return results[0]
    return _failed_write_result(job, "Пустой результат записи.")


def execute_journal_writes(
    client: SheetsClient,
    config: CatalogConfig,
    jobs: Sequence[JournalWriteJob],
) -> tuple[GoogleWriteResult, ...]:
    """Write jobs with one HTTP ``values:batchUpdate`` per kit.

    Discovers the kits tab and reads A:F once. Jobs for the same kit are
    planned in input order against an in-memory A:F matrix so chained F
    lines see the previous letter. The sheet receives only the coalesced
    final cells (last value wins per A1 range). Different kits stay
    separate HTTP calls. Does not create rows.

    Args:
        client: Sheets API adapter.
        config: Catalog config.
        jobs: Writes in order (same kit may appear more than once).

    Returns:
        One result per job, aligned with ``jobs``, including per-row errors.
    """

    if not jobs:
        return ()
    try:
        meta = client.get_spreadsheet(config.google_kits_spreadsheet_id)
        pin = pin_ksb_id_worksheet(
            client,
            meta,
            config.google_kits_sheet_name,
            spreadsheet_id=config.google_kits_spreadsheet_id,
        )
        rows = [
            list(row)
            for row in client.get_values(
                pin.spreadsheet_id, a1_range(pin.title, "A:F")
            )
        ]
    except GoogleWriteError as exc:
        return tuple(_failed_write_result(job, str(exc)) for job in jobs)
    runtime = Path(config.runtime_dir)
    if runtime.is_dir():
        save_kits_sheet_pin(
            runtime,
            spreadsheet_id=pin.spreadsheet_id,
            sheet_id=pin.sheet_id,
            title=pin.title,
        )
    groups: dict[tuple[str, str], list[int]] = {}
    kit_order: list[tuple[str, str]] = []
    for index, job in enumerate(jobs):
        key = kit_identity_key(job.title, job.mark)
        if key not in groups:
            groups[key] = []
            kit_order.append(key)
        groups[key].append(index)
    results: list[GoogleWriteResult | None] = [None] * len(jobs)
    for key in kit_order:
        working = [list(row) for row in rows]
        planned_ok: list[int] = []
        pending: list[tuple[str, str]] = []
        for job_index in groups[key]:
            job = jobs[job_index]
            try:
                plan = plan_journal_write(working, pin, job)
            except GoogleWriteError as exc:
                results[job_index] = _failed_write_result(job, str(exc))
                continue
            _apply_a1_updates(working, plan.updates)
            pending.extend(plan.updates)
            planned_ok.append(job_index)
            patch = plan.patch
            wrote_de = any(
                _a1_cell(a1)[0] in {"D", "E"} for a1, _value in plan.updates
            )
            results[job_index] = GoogleWriteResult(
                title=job.title,
                mark=job.mark,
                row_index=plan.located.row_index,
                sheet_title=pin.title,
                comment_before=patch.comment_before,
                comment_after=patch.comment_after,
                update_de=wrote_de,
            )
        coalesced = _coalesce_a1_updates(pending)
        if coalesced:
            payload = tuple((a1, [[value]]) for a1, value in coalesced)
            try:
                client.update_values(pin.spreadsheet_id, payload)
            except GoogleWriteError as exc:
                error_text = str(exc)
                for job_index in planned_ok:
                    results[job_index] = _failed_write_result(
                        jobs[job_index], error_text
                    )
                continue
        rows[:] = working
    return tuple(
        result
        if result is not None
        else _failed_write_result(job, "Пустой результат записи.")
        for job, result in zip(jobs, results, strict=True)
    )


class GoogleSheetsRestClient:
    """Sheets API v4 over an authorized HTTP session (no Drive API)."""

    def __init__(
        self,
        credentials_path: str | Path | None = None,
        *,
        timeout_sec: float | None = None,
    ) -> None:
        """Authorize the service account.

        Args:
            credentials_path: JSON key; default :func:`service_account_path`.
            timeout_sec: HTTP timeout; ``RD_KITS_WRITE_TIMEOUT_SEC`` or 30.

        Raises:
            GoogleWriteError: Missing key file, google-auth, or OAuth/TLS.
        """

        path = Path(credentials_path) if credentials_path else service_account_path()
        if not path.is_file():
            raise GoogleWriteError(
                f"Нет ключа сервисного аккаунта: {path}. "
                "Нужен доступ python-agcc@… к таблице КСБ ИД."
            )
        try:
            from google.auth.transport.requests import AuthorizedSession, Request
            from google.oauth2.service_account import Credentials
        except ImportError as exc:
            raise GoogleWriteError(
                "Пакет google-auth не установлен — запись в Google недоступна."
            ) from exc
        creds = Credentials.from_service_account_file(
            str(path), scopes=[_SHEETS_SCOPE]
        )
        verify = not kits_tls_relaxed()
        token_session = _requests_session(verify=verify)
        try:
            creds.refresh(Request(session=token_session))
        except Exception as exc:
            raise GoogleWriteError(_oauth_error_text(exc)) from exc
        finally:
            token_session.close()
        auth_http = _requests_session(verify=verify)
        self._auth_http = auth_http
        self._session = AuthorizedSession(
            creds, auth_request=Request(session=auth_http)
        )
        self._session.verify = verify
        if timeout_sec is None:
            timeout_sec = float(os.environ.get("RD_KITS_WRITE_TIMEOUT_SEC", "30"))
        self._timeout = timeout_sec

    def get_spreadsheet(self, spreadsheet_id: str) -> dict[str, Any]:
        url = f"{_SHEETS_API}/spreadsheets/{quote(spreadsheet_id, safe='')}"
        params = {"fields": "spreadsheetId,sheets.properties(sheetId,title,index)"}
        payload = self._request("GET", url, params=params)
        if not isinstance(payload, dict):
            raise GoogleWriteError("Sheets API вернул неожиданный метаданные.")
        return payload

    def get_values(self, spreadsheet_id: str, a1_range: str) -> list[list[str]]:
        encoded = quote(a1_range, safe="!'")
        url = (
            f"{_SHEETS_API}/spreadsheets/{quote(spreadsheet_id, safe='')}"
            f"/values/{encoded}"
        )
        payload = self._request("GET", url)
        values = payload.get("values") if isinstance(payload, dict) else None
        if not values:
            return []
        return [
            ["" if cell is None else str(cell) for cell in row] for row in values
        ]

    def batch_get_values(
        self,
        spreadsheet_id: str,
        ranges: Sequence[str],
    ) -> list[list[list[str]]]:
        """Read several A1 ranges in one Sheets API call.

        Args:
            spreadsheet_id: Spreadsheet id.
            ranges: A1 ranges, same order as the returned matrices.

        Returns:
            One string matrix per range (empty matrix when the range is blank).
        """

        if not ranges:
            return []
        url = (
            f"{_SHEETS_API}/spreadsheets/{quote(spreadsheet_id, safe='')}"
            "/values:batchGet"
        )
        payload = self._request(
            "GET", url, params=[("ranges", item) for item in ranges]
        )
        value_ranges = payload.get("valueRanges") if isinstance(payload, dict) else None
        matrices: list[list[list[str]]] = []
        for item in value_ranges or []:
            values = item.get("values") if isinstance(item, dict) else None
            if not values:
                matrices.append([])
                continue
            matrices.append(
                [
                    ["" if cell is None else str(cell) for cell in row]
                    for row in values
                ]
            )
        while len(matrices) < len(ranges):
            matrices.append([])
        return matrices[: len(ranges)]

    def update_values(
        self,
        spreadsheet_id: str,
        data: Sequence[tuple[str, list[list[str]]]],
    ) -> None:
        url = (
            f"{_SHEETS_API}/spreadsheets/{quote(spreadsheet_id, safe='')}"
            "/values:batchUpdate"
        )
        body = {
            "valueInputOption": "RAW",
            "data": [
                {"range": a1, "majorDimension": "ROWS", "values": values}
                for a1, values in data
            ],
        }
        self._request("POST", url, json_body=body)

    def _request(
        self,
        method: str,
        url: str,
        *,
        params: Any = None,
        json_body: Mapping[str, Any] | None = None,
    ) -> Any:
        retries = max(0, int(os.environ.get("RD_KITS_WRITE_QUOTA_RETRIES", "5")))
        attempts = 1 + retries
        response: Any = None
        for attempt in range(attempts):
            try:
                response = self._session.request(
                    method,
                    url,
                    params=_request_params(params),
                    json=json_body,
                    timeout=self._timeout,
                    verify=self._session.verify,
                )
            except Exception as exc:
                raise GoogleWriteError(
                    f"Сеть Sheets API: {type(exc).__name__}: {exc}"
                ) from exc
            if response.status_code != 429 or attempt + 1 >= attempts:
                break
            time.sleep(_quota_retry_delay_sec(response, attempt))
        if response is None:
            raise GoogleWriteError("Сеть Sheets API: пустой ответ.")
        if response.status_code >= 400:
            detail = _api_error_text(response)
            raise GoogleWriteError(
                f"Sheets API HTTP {response.status_code}: {detail}"
            )
        if not response.content:
            return {}
        try:
            return response.json()
        except json.JSONDecodeError as exc:
            raise GoogleWriteError("Sheets API вернул не JSON.") from exc


def _request_params(params: Any) -> Any:
    if not params:
        return None
    if isinstance(params, Mapping):
        return dict(params)
    return list(params)


def _quota_retry_delay_sec(response: Any, attempt: int) -> float:
    raw = ""
    headers = getattr(response, "headers", None)
    if headers:
        raw = str(headers.get("Retry-After") or headers.get("retry-after") or "")
    try:
        if raw.strip():
            return min(90.0, max(1.0, float(raw)))
    except ValueError:
        pass
    return min(60.0, 8.0 * (attempt + 1))


def _header_cells_for_sheets(
    client: SheetsClient,
    spreadsheet_id: str,
    titles: Sequence[str],
) -> list[tuple[str, str]]:
    """Return A1/B1 for each worksheet title (one batchGet when available)."""

    if not titles:
        return []
    batch = getattr(client, "batch_get_values", None)
    if callable(batch):
        matrices = batch(
            spreadsheet_id, [a1_range(title, "A1:B1") for title in titles]
        )
        pairs: list[tuple[str, str]] = []
        for matrix in matrices:
            row = matrix[0] if matrix else []
            a_cell = str(row[0] or "") if row else ""
            b_cell = str(row[1] or "") if len(row) > 1 else ""
            pairs.append((a_cell, b_cell))
        while len(pairs) < len(titles):
            pairs.append(("", ""))
        return pairs[: len(titles)]
    pairs = []
    for title in titles:
        a_cell = _cell_text(client.get_values(spreadsheet_id, a1_range(title, "A1")))
        b_cell = _cell_text(client.get_values(spreadsheet_id, a1_range(title, "B1")))
        pairs.append((a_cell, b_cell))
    return pairs


def _row_cell(rows: Sequence[Sequence[Any]], row_index: int, column: int) -> str:
    """Return a 0-based column from a 1-based sheet row."""

    if row_index < 1 or column < 0:
        return ""
    row = rows[row_index - 1]
    if column >= len(row):
        return ""
    return str(row[column] or "").strip()


def _a1_cell(a1: str) -> tuple[str | None, int]:
    match = _A1_CELL_RE.search(a1 or "")
    if not match:
        return None, 0
    return match.group(1), int(match.group(2))


def _a1_column_index(letters: str) -> int:
    n = 0
    for char in letters.upper():
        if not "A" <= char <= "Z":
            return -1
        n = n * 26 + (ord(char) - 64)
    return n - 1


def _coalesce_a1_updates(
    updates: Sequence[tuple[str, str]],
) -> tuple[tuple[str, str], ...]:
    """Keep the last value per A1 range, in first-seen range order.

    Args:
        updates: Planned cell writes in job order.

    Returns:
        Deduplicated ``(A1, value)`` pairs for one ``values:batchUpdate``.
    """

    order: list[str] = []
    last: dict[str, str] = {}
    for a1, value in updates:
        if a1 not in last:
            order.append(a1)
        last[a1] = value
    return tuple((a1, last[a1]) for a1 in order)


def _apply_a1_updates(
    rows: list[list[str]],
    updates: Sequence[tuple[str, str]],
) -> None:
    """Write planned A1 cells into an in-memory A:F matrix."""

    for a1, value in updates:
        column, row_index = _a1_cell(a1)
        if column is None or row_index < 1:
            continue
        col = _a1_column_index(column)
        if col < 0:
            continue
        idx = row_index - 1
        while len(rows) <= idx:
            rows.append([])
        row = rows[idx]
        while len(row) <= col:
            row.append("")
        row[col] = value


def _failed_write_result(job: JournalWriteJob, error: str) -> GoogleWriteResult:
    return GoogleWriteResult(
        title=job.title,
        mark=job.mark,
        row_index=0,
        sheet_title="",
        comment_before="",
        comment_after="",
        update_de=False,
        error=error,
    )


def _requests_session(*, verify: bool) -> Any:
    """Build a ``requests.Session`` with the kits CSV TLS policy."""

    import requests

    session = requests.Session()
    session.verify = verify
    if not verify:
        import urllib3

        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    return session


def _oauth_error_text(exc: BaseException) -> str:
    """Turn an OAuth/transport failure into a short GUI message."""

    text = f"{type(exc).__name__}: {exc}"
    low = text.lower()
    if "certificate" in low or "ssl" in low:
        return (
            "OAuth Google: не прошла проверка TLS-сертификата "
            f"(часто корпоративный SSL-прокси). {text}"
        )
    return f"OAuth Google: {text}"


def _api_error_text(response: Any) -> str:
    try:
        payload = response.json()
        error = payload.get("error") if isinstance(payload, dict) else None
        if isinstance(error, dict) and error.get("message"):
            return str(error["message"])
    except Exception:
        pass
    text = getattr(response, "text", "") or ""
    return text[:400] or "без текста"


def _cell_text(values: Sequence[Sequence[str]]) -> str:
    if not values or not values[0]:
        return ""
    return values[0][0]


def _normalize_cell(text: str) -> str:
    return (text or "").replace("\r\n", "\n").replace("\r", "\n")
