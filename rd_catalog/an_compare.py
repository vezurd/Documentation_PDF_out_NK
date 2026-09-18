"""Pairwise AN content compare against AutoMTO (PI) and RD MTO.

Qt-free. The AN dump scan still does not open workbooks; this module is
used by the АН tab after the name index is already built.
"""

from __future__ import annotations

from pathlib import Path

from rd_catalog.an_index import AnAgreedScore, KIND_HEADER
from rd_catalog.customer_pi_auto_mto import (
    MtoPairCompareResult,
    compare_document_to_path,
)
from rd_catalog.mto_diff import RowLoader, load_canonical_mto

AN_VS_REV_MATCH_FILL = "#E2F2E1"
AN_VS_CONTENT_MATCH_FILL = "#C5E8C4"
AN_VS_CONTENT_MISS_FILL = "#D4E6B5"
AN_VS_REV_DIFF_FILL = "#F7E8BE"
AN_AGREED_BEST_FILL = "#C5E8C4"
AN_AGREED_BEST_ROW_FILL = "#E2F2E1"
AN_AGREED_HEADER = "К согл. передаче"
AN_AGREED_HEADER_TIP = (
    "Насколько эта папка АН похожа на комплект, переданный "
    "по последней согласованной ревизии (код A / статус «согласован»)."
)
AN_AGREED_ROW_HEADERS = (
    "Титул",
    "Марка",
    "Ревизия АН",
    KIND_HEADER,
    AN_AGREED_HEADER,
    "Имя",
    "Дата",
    "Папка",
    "Путь",
)
AN_HEADERS = (
    "Титул",
    "Марка",
    "Ревизия АН",
    KIND_HEADER,
    AN_AGREED_HEADER,
    "vs Авто МТО",
    "vs MTO РД",
    "vs Робот",
    "vs Выдача",
    "vs F",
    "vs SQ",
    "Имя",
    "Дата",
    "Папка",
    "Путь",
)


def format_an_agreed_cell(score: AnAgreedScore | None) -> str:
    """Return the АН «К согл. передаче» cell text.

    Args:
        score: Per-file score from ``score_an_files_for_agreed``.

    Returns:
        ``—`` without an agreed target, otherwise ``96%`` or
        ``96% · лучше`` for the highlighted winner.
    """

    if score is None or score.percent is None:
        return "—"
    text = f"{score.percent}%"
    if score.is_best:
        return f"{text} · лучше"
    return text


def an_agreed_cell_fill(score: AnAgreedScore | None) -> str | None:
    """Return the hex fill for the agreed-transfer probability cell.

    Args:
        score: Per-file score, if any.

    Returns:
        Rich green for the winner, otherwise ``None``.
    """

    if score is not None and score.is_best:
        return AN_AGREED_BEST_FILL
    return None


def an_agreed_row_fill(score: AnAgreedScore | None) -> str | None:
    """Return a pale row tint for the winning agreed-transfer folder.

    vs-* cells keep their own fills; callers apply this only to identity
    / path columns.

    Args:
        score: Per-file score, if any.

    Returns:
        Pale green for the winner, otherwise ``None``.
    """

    if score is not None and score.is_best:
        return AN_AGREED_BEST_ROW_FILL
    return None


def an_agreed_cell_tooltip(score: AnAgreedScore | None) -> str:
    """Return the tooltip for the agreed-transfer probability cell.

    Args:
        score: Per-file score, if any.

    Returns:
        Multiline Russian explanation.
    """

    if score is None:
        return AN_AGREED_HEADER_TIP
    if not score.reasons:
        return AN_AGREED_HEADER_TIP
    return "\n".join(score.reasons)


def format_an_vs_cell(
    rev_match: bool | None,
    content: MtoPairCompareResult | None,
    *,
    has_counterpart: bool,
) -> str:
    """Return ``да/нет/—`` plus an optional content-grade parenthesis.

    Args:
        rev_match: Filename-revision compare, or ``None`` when no target.
        content: Cached pairwise result, if any.
        has_counterpart: True when a counterpart workbook path exists.

    Returns:
        Table cell text. Parentheses appear only when a counterpart file
        is available to compare.
    """

    if rev_match is True:
        base = "да"
    elif rev_match is False:
        base = "нет"
    else:
        base = "—"
    if not has_counterpart:
        return base
    label = content.paren_label if content is not None else "не сверялось"
    return f"{base} ({label})"


def an_vs_cell_fill(
    rev_match: bool | None,
    content: MtoPairCompareResult | None,
) -> str | None:
    """Return the hex fill for an АН vs-Авто-МТО / vs-MTO-РД cell.

    ``да`` uses two greens: richer when content matches, yellow-green
    when the revision matches but the workbooks do not. Pending compare
    keeps the pale revision-match green. ``нет`` stays yellow.

    Args:
        rev_match: Filename-revision compare, or ``None`` when no target.
        content: Cached pairwise result, if any.

    Returns:
        ``#RRGGBB``, or ``None`` to leave the cell unpainted.
    """

    if rev_match is True:
        if content is not None and content.kind in {"matched", "soft"}:
            return AN_VS_CONTENT_MATCH_FILL
        if content is not None and content.kind == "no_match":
            return AN_VS_CONTENT_MISS_FILL
        return AN_VS_REV_MATCH_FILL
    if rev_match is False:
        return AN_VS_REV_DIFF_FILL
    return None


def an_vs_cell_tooltip(
    *,
    rev_match: bool | None,
    content: MtoPairCompareResult | None,
    an_path: str,
    other_path: str,
    other_label: str,
) -> str:
    """Return a tooltip for one vs-Авто-МТО / vs-MTO-РД cell.

    Args:
        rev_match: Filename-revision compare.
        content: Cached pairwise result, if any.
        an_path: AN workbook path.
        other_path: Counterpart workbook path.
        other_label: Human name of the counterpart.

    Returns:
        Multiline Russian tooltip.
    """

    if rev_match is True:
        rev_text = "да"
    elif rev_match is False:
        rev_text = "нет"
    else:
        rev_text = "—"
    lines = [
        f"Ревизия в имени: {rev_text}.",
        f"АН: {Path(an_path).name if an_path else '—'}",
        f"{other_label}: {Path(other_path).name if other_path else 'нет файла'}",
    ]
    if content is not None:
        lines.extend(content.tooltip_lines(other_label=other_label))
    elif other_path:
        lines.append(f"Содержимое: сверка с {other_label} ещё не завершена.")
    return "\n".join(lines)


def compare_an_file(
    an_path: str | Path,
    *,
    auto_path: str = "",
    rd_path: str = "",
    loader: RowLoader | None = None,
) -> tuple[MtoPairCompareResult | None, MtoPairCompareResult | None]:
    """Load the AN workbook once and grade it against PI and/or RD.

    Args:
        an_path: AN MTO workbook.
        auto_path: Local АвтоМто xlsx from the customer PI catalog.
        rd_path: Current or pinned RD MTO workbook.
        loader: Optional side-effect-free row loader, primarily for tests.

    Returns:
        ``(vs_auto, vs_rd)``. A side is ``None`` when its path is empty.
    """

    try:
        an_document = load_canonical_mto(an_path, loader=loader)
    except Exception as exc:
        failed = MtoPairCompareResult(
            kind="not_compared",
            error=str(exc).splitlines()[0].strip() or type(exc).__name__,
        )
        return (
            failed if auto_path else None,
            failed if rd_path else None,
        )
    vs_auto = (
        compare_document_to_path(an_document, auto_path, loader=loader)
        if auto_path
        else None
    )
    vs_rd = (
        compare_document_to_path(an_document, rd_path, loader=loader)
        if rd_path
        else None
    )
    return vs_auto, vs_rd
