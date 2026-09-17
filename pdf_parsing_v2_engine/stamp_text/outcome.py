"""Structured cleaning result (internal to v2 extraction; ``FieldResult`` stays str-based)."""

from __future__ import annotations

from dataclasses import dataclass, field

# ASCII rule between ``CleanOutcome`` blocks (console-safe).
_RULE_WIDTH = 56
_RULE = "-" * _RULE_WIDTH


@dataclass(frozen=True)
class ParseWarning:
    """Single diagnostic from stamp text parsing."""

    code: str
    message: str
    tier: str | None = None
    """``\"strict\"`` or ``\"relaxed\"`` when applicable."""


def _append_text_block(lines: list[str], label: str, text: str) -> None:
    """One labelled block; each physical line prefixed with ``|`` (ASCII; empty → ``(empty)``)."""
    if not text:
        lines.append(f"  {label}:")
        lines.append("    | (empty)")
        return
    parts = text.split("\n")
    suffix = f" ({len(parts)} lines)" if len(parts) > 1 else ""
    lines.append(f"  {label}{suffix}:")
    for part in parts:
        lines.append(f"    | {part}")


def _append_warnings(lines: list[str], warnings: list[ParseWarning]) -> None:
    if not warnings:
        return
    lines.append("  warnings:")
    for i, w in enumerate(warnings):
        tier_part = f" (tier={w.tier})" if w.tier is not None else ""
        msg_lines = w.message.split("\n")
        lines.append(f"    | [{i}] {w.code}{tier_part}  {msg_lines[0]}")
        for ml in msg_lines[1:]:
            lines.append(f"    | {ml}")


@dataclass
class CleanOutcome:
    """Result of a cleaner pipeline before mapping to ``FieldResult``."""

    value: str
    raw_input: str
    warnings: list[ParseWarning] = field(default_factory=list)
    via: str | None = None
    clean_tier: str | None = None
    """``\"strict\"`` / ``\"relaxed\"`` summary when pipeline uses fallback."""

    def _format_meta_line(self) -> str:
        via = self.via if self.via is not None else "(none)"
        tier = self.clean_tier if self.clean_tier is not None else "(none)"
        return f"CleanOutcome  via={via}  tier={tier}  warnings={len(self.warnings)}"

    def __str__(self) -> str:
        """Human-readable multi-line summary (used by ``print(outcome)``)."""
        lines: list[str] = [_RULE, self._format_meta_line()]
        same = self.value == self.raw_input
        single_line = "\n" not in self.value
        no_warn = not self.warnings

        if same and no_warn and single_line:
            if self.value:
                lines.append(f"  {self.value}")
            else:
                lines.append("  | (empty)")
        elif same and single_line:
            if self.value:
                lines.append(f"  {self.value}")
            else:
                lines.append("  | (empty)")
            _append_warnings(lines, self.warnings)
        elif same:
            if not self.value:
                lines.append("  value / raw_input:")
                lines.append("    | (empty)")
            else:
                parts = self.value.split("\n")
                suffix = f" ({len(parts)} lines)" if len(parts) > 1 else ""
                lines.append(f"  value / raw_input{suffix}:")
                for part in parts:
                    lines.append(f"    | {part}")
            _append_warnings(lines, self.warnings)
        else:
            _append_text_block(lines, "value", self.value)
            _append_text_block(lines, "raw_input", self.raw_input)
            _append_warnings(lines, self.warnings)

        lines.append(_RULE)
        return "\n".join(lines)


def strict_clean_outcome(
    raw_input: str,
    value: object,
    *,
    via: str | None = None,
) -> CleanOutcome:
    """Build a strict-tier outcome with no warnings (single-path cleaners).

    Args:
        raw_input: Original field text before cleaning.
        value: Cleaned value; non-strings are coerced like the former ``_wrap``.
        via: Optional trace label (typically the JSON ``clean`` key).

    Returns:
        ``CleanOutcome`` with empty warnings and ``clean_tier=\"strict\"``.
    """
    if not isinstance(value, str):
        value = str(value) if value is not None else ""
    return CleanOutcome(
        value=value,
        raw_input=raw_input,
        warnings=[],
        via=via,
        clean_tier="strict",
    )
