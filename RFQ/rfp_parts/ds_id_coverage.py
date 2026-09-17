"""Compare RFP part workbooks to UL (TSD) folders by actual DS number.

Read-only. Used by the GUI tab and the ``ds_id_check`` RFP-run milestone.
Does not load xlsx contents — names only.
"""

from __future__ import annotations

import os
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path

from RFQ.ds_compare.tsd_packing_load import DEFAULT_TSD_PACKING_ROOT
from RFQ.rfp_parts.ds_checklist import DEFAULT_PARTS_DIR
from RFQ.rfp_parts.ds_identity import (
    DsIdentity,
    parse_rfp_ds_identity,
    parse_ul_folder_ds_identity,
)

_EXCEL = {".xlsx", ".xlsm", ".xls"}
_SKIP_DIR_PREFIXES = (".", "~")


@dataclass(frozen=True, slots=True)
class DsIdCoverRow:
    """One actual-DS row in the UL ↔ RFP coverage table."""

    actual_label: str
    status: str
    status_ru: str
    ul_folder: str
    ul_xlsx: int
    rfp_labels: tuple[str, ...]
    rfp_files: tuple[str, ...]
    rfp_sequential: tuple[int, ...]
    note: str


@dataclass
class DsIdCoverageResult:
    """Outcome of comparing UL folders to RFP part files."""

    ul_root: Path
    rfp_root: Path
    load_error: str | None = None
    rows: list[DsIdCoverRow] = field(default_factory=list)
    both: int = 0
    ul_only: int = 0
    rfp_only: int = 0
    gf: int = 0
    unparsed: int = 0
    ul_folders: int = 0
    rfp_files: int = 0

    @property
    def is_ok(self) -> bool:
        """True when every numbered side has a pair and nothing is unparsed.

        Госфин folders without a DS number remain a remark (not ok).
        """
        return (
            self.load_error is None
            and self.ul_only == 0
            and self.rfp_only == 0
            and self.gf == 0
            and self.unparsed == 0
        )

    def summary_line(self) -> str:
        """Short Russian line for the RFP milestone / GUI banner."""
        if self.load_error:
            return f"Ошибка: {self.load_error}"
        if self.is_ok:
            return f"ОК: пар {self.both}, замечаний нет"
        bits: list[str] = []
        rfp_only_labels = [
            row.actual_label for row in self.rows if row.status == "rfp_only"
        ]
        ul_only_labels = [
            row.actual_label for row in self.rows if row.status == "ul_only"
        ]
        if rfp_only_labels:
            bits.append("только RFP " + ", ".join(rfp_only_labels))
        if ul_only_labels:
            bits.append("только УЛ " + ", ".join(ul_only_labels))
        if self.gf:
            bits.append(f"ГФ без номера ДС ({self.gf} пап.)")
        unparsed_names = [
            row.ul_folder or row.actual_label
            for row in self.rows
            if row.status == "unparsed"
        ]
        if unparsed_names:
            bits.append("без номера: " + ", ".join(unparsed_names[:4]))
        extra = "; ".join(bits) if bits else "см. таблицу"
        return f"Замечания: {extra}"


def _is_excel(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in _EXCEL and not path.name.startswith("~$")


def _xlsx_in_folder(folder: Path) -> int:
    try:
        return sum(1 for path in folder.iterdir() if _is_excel(path))
    except OSError:
        return 0


def _list_ul_folders(tsd_root: Path) -> list[tuple[str, int, DsIdentity]]:
    items: list[tuple[str, int, DsIdentity]] = []
    for child in sorted(tsd_root.iterdir(), key=lambda p: p.name.lower()):
        if not child.is_dir() or child.name.startswith(_SKIP_DIR_PREFIXES):
            continue
        ident = parse_ul_folder_ds_identity(child.name)
        items.append((child.name, _xlsx_in_folder(child), ident))
    return items


def _list_rfp_files(parts_dir: Path) -> list[tuple[str, DsIdentity]]:
    items: list[tuple[str, DsIdentity]] = []
    for path in sorted(parts_dir.iterdir(), key=lambda p: p.name.lower()):
        if not _is_excel(path):
            continue
        items.append((path.name, parse_rfp_ds_identity(path.name)))
    return items


def check_ds_id_coverage(
    parts_dir: str | Path | None = None,
    tsd_root: str | Path | None = None,
) -> DsIdCoverageResult:
    """Scan RFP parts and UL folders; join on actual DS number.

    Args:
        parts_dir: ``RFP_Зиновьев``. Defaults to ``DEFAULT_PARTS_DIR``.
        tsd_root: TSD packing root. Defaults to ``DEFAULT_TSD_PACKING_ROOT``.

    Returns:
        Coverage result. Missing UNC roots set ``load_error``; rows may be empty.
    """
    rfp_root = Path(parts_dir or DEFAULT_PARTS_DIR)
    ul_root = Path(tsd_root or DEFAULT_TSD_PACKING_ROOT)
    result = DsIdCoverageResult(ul_root=ul_root, rfp_root=rfp_root)

    missing: list[str] = []
    if not ul_root.exists():
        missing.append(f"нет папки УЛ ({ul_root})")
    if not rfp_root.exists():
        missing.append(f"нет папки RFP ({rfp_root})")
    if missing:
        result.load_error = "; ".join(missing)
        return result

    try:
        ul_items = _list_ul_folders(ul_root)
        rfp_items = _list_rfp_files(rfp_root)
    except OSError as exc:
        result.load_error = str(exc)
        return result

    result.ul_folders = len(ul_items)
    result.rfp_files = len(rfp_items)

    rfp_by_actual: dict[int, list[tuple[str, DsIdentity]]] = defaultdict(list)
    for name, ident in rfp_items:
        if ident.actual is not None:
            rfp_by_actual[ident.actual].append((name, ident))

    used_rfp: set[str] = set()
    rows: list[DsIdCoverRow] = []

    def _sort_ul(item: tuple[str, int, DsIdentity]) -> tuple:
        ident = item[2]
        return (ident.actual is None, ident.actual or 10**9, item[0].lower())

    for folder, xlsx_count, ident in sorted(ul_items, key=_sort_ul):
        if ident.kind == "gf":
            gf_rfp = [
                item
                for item in rfp_items
                if item[0] not in used_rfp
                and (
                    item[1].sequential in {1, 2, 7}
                    or "ГФ" in item[0].upper()
                    or "ГОСФИН" in item[0].upper()
                )
            ]
            for name, _ident in gf_rfp:
                used_rfp.add(name)
            result.gf += 1
            rows.append(
                DsIdCoverRow(
                    actual_label="ГФ",
                    status="gf_special",
                    status_ru="Особая группа ГФ",
                    ul_folder=folder,
                    ul_xlsx=xlsx_count,
                    rfp_labels=tuple(dict.fromkeys(i.label for _, i in gf_rfp)),
                    rfp_files=tuple(name for name, _ident in gf_rfp),
                    rfp_sequential=tuple(
                        sorted({i.sequential for _, i in gf_rfp if i.sequential})
                    ),
                    note="Папка госфинансирования, не номер ДС. Нужна раскладка или соответствие.",
                )
            )
            continue
        if ident.kind == "unparsed" or not ident.match_keys:
            result.unparsed += 1
            rows.append(
                DsIdCoverRow(
                    actual_label="?",
                    status="unparsed",
                    status_ru="УЛ без номера",
                    ul_folder=folder,
                    ul_xlsx=xlsx_count,
                    rfp_labels=(),
                    rfp_files=(),
                    rfp_sequential=(),
                    note="Номер ДС из имени папки не извлечён.",
                )
            )
            continue

        hits: list[tuple[str, DsIdentity]] = []
        seen: set[str] = set()
        for key in ident.match_keys:
            for name, rident in rfp_by_actual.get(key, []):
                if name not in seen:
                    hits.append((name, rident))
                    seen.add(name)

        for name, _ident in hits:
            used_rfp.add(name)

        if hits:
            result.both += 1
            note = "Связка по фактическому номеру ДС."
            if ident.compound:
                note = (
                    f"Папка составная {ident.label}: фактический {ident.actual}, "
                    f"порядковый хвост {ident.sequential} (информационный)."
                )
            corr = [i.label for _, i in hits if i.compound]
            if corr:
                note += " RFP-корректировка: " + ", ".join(corr) + "."
            if "ГФ" in folder.upper() or "ГОСФИН" in folder.upper():
                note += " Пометка ГФ в имени папки; номер взят из ДС."
            rows.append(
                DsIdCoverRow(
                    actual_label=ident.label,
                    status="both",
                    status_ru="Есть в УЛ и RFP",
                    ul_folder=folder,
                    ul_xlsx=xlsx_count,
                    rfp_labels=tuple(dict.fromkeys(i.label for _, i in hits)),
                    rfp_files=tuple(name for name, _ident in hits),
                    rfp_sequential=tuple(
                        sorted({i.sequential for _, i in hits if i.sequential})
                    ),
                    note=note,
                )
            )
        else:
            result.ul_only += 1
            rows.append(
                DsIdCoverRow(
                    actual_label=ident.label,
                    status="ul_only",
                    status_ru="Только УЛ",
                    ul_folder=folder,
                    ul_xlsx=xlsx_count,
                    rfp_labels=(),
                    rfp_files=(),
                    rfp_sequential=(),
                    note="Папка УЛ есть, файла RFP с этим фактическим номером нет.",
                )
            )

    leftover: dict[int, list[tuple[str, DsIdentity]]] = defaultdict(list)
    for name, ident in rfp_items:
        if name in used_rfp:
            continue
        key = ident.match_key if ident.match_key is not None else -1
        leftover[key].append((name, ident))

    for key in sorted(leftover, key=lambda n: (n < 0, n)):
        items = leftover[key]
        labels = tuple(dict.fromkeys(i.label for _, i in items))
        result.rfp_only += 1
        note = "Файл RFP есть, папки УЛ с этим фактическим номером нет."
        if any("ПРИЛОЖЕНИЕ" in name.upper() for name, _ident in items):
            note += " Несколько приложений одной ДС."
        rows.append(
            DsIdCoverRow(
                actual_label=labels[0] if labels else "?",
                status="rfp_only",
                status_ru="Только RFP",
                ul_folder="",
                ul_xlsx=0,
                rfp_labels=labels,
                rfp_files=tuple(name for name, _ident in items),
                rfp_sequential=tuple(
                    sorted({i.sequential for _, i in items if i.sequential})
                ),
                note=note,
            )
        )

    result.rows = rows
    return result


_last_coverage: DsIdCoverageResult | None = None


def get_last_ds_id_coverage() -> DsIdCoverageResult | None:
    """Return the last result from ``run_ds_id_coverage_job``, if any."""
    return _last_coverage


@dataclass(frozen=True, slots=True)
class DsIdCoverageJobResult:
    """Duck-typed ``FunctionJobRunner`` result (no GUI import)."""

    success: bool
    message: str
    result_path: str | None = None


def _print_coverage(result: DsIdCoverageResult) -> None:
    """Write the summary and TSV table to stdout.

    Uses ``write()``, not ``stdout.buffer``: the GUI job runner replaces
    stdout with a text stream that has no ``.buffer``.
    """

    def _out(msg: str) -> None:
        text = msg + "\n"
        stream = sys.stdout
        try:
            stream.write(text)
        except UnicodeEncodeError:
            encoding = getattr(stream, "encoding", None) or "utf-8"
            payload = text.encode(encoding, errors="backslashreplace")
            buf = getattr(stream, "buffer", None)
            if buf is not None:
                buf.write(payload)
            else:
                stream.write(payload.decode(encoding, errors="replace"))

    _out(result.summary_line())
    _out(f"УЛ: {result.ul_root}")
    _out(f"RFP: {result.rfp_root}")
    if result.load_error:
        return
    _out(
        "Фактический ДС\tСтатус\tПапка УЛ\tУЛ файлов\tRFP\tПорядковый RFP\tКомментарий"
    )
    for row in result.rows:
        rfp = ", ".join(row.rfp_labels) or "—"
        seq = ", ".join(str(n) for n in row.rfp_sequential) or "—"
        _out(
            f"{row.actual_label}\t{row.status_ru}\t{row.ul_folder or '—'}\t"
            f"{row.ul_xlsx}\t{rfp}\t{seq}\t{row.note}"
        )


def run_ds_id_coverage_job(
    parts_dir: str | Path | None = None,
    tsd_root: str | Path | None = None,
) -> DsIdCoverageJobResult:
    """Scan coverage, print the table, remember the result for the GUI tab.

    Args:
        parts_dir: Optional RFP parts folder override.
        tsd_root: Optional UL root override.

    Returns:
        Job result. ``success`` is false only on ``load_error`` (missing
        roots). Remarks still succeed so the table can be shown. Does not raise.
    """
    global _last_coverage
    os.environ.setdefault("PYTHONUTF8", "1")
    result = check_ds_id_coverage(parts_dir=parts_dir, tsd_root=tsd_root)
    _last_coverage = result
    _print_coverage(result)
    summary = result.summary_line()
    return DsIdCoverageJobResult(
        success=result.load_error is None,
        message=summary,
        result_path=None,
    )


def run_ds_id_coverage_cli(
    parts_dir: str | Path | None = None,
    tsd_root: str | Path | None = None,
) -> int:
    """Print a UTF-8 table to stdout. Returns 0 if ok, 1 if remarks/error."""
    os.environ.setdefault("PYTHONUTF8", "1")
    result = check_ds_id_coverage(parts_dir=parts_dir, tsd_root=tsd_root)
    _print_coverage(result)
    if result.load_error or not result.is_ok:
        return 1
    return 0


def main() -> int:
    return run_ds_id_coverage_cli()


if __name__ == "__main__":
    raise SystemExit(main())
