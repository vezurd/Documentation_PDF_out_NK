"""Read-only: diff UL phase timings and counters between two RFP run folders."""

from __future__ import annotations

import re
import sys
from pathlib import Path

BASE_DIR = Path(
    r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\RFP_MTO_VO\_РЕЗУЛЬТАТА_ПРОВЕРКИ"
)
BASELINE = BASE_DIR / "_результат_проверки_2026.09.09.12.57"
CONTROL = BASE_DIR / "_результат_проверки_2026.09.09.13.40"

SECTION_RE = re.compile(r"^([A-Z][A-Z ]+?)(?: \([a-z]+\))?:$")
ENTRY_RE = re.compile(r"^\s{2}(\S[^:]*):\s*(.+)$")


def parse_report(run: Path) -> tuple[str, dict[str, dict[str, str]]]:
    reports = sorted(run.glob("ul_compare_report_*.txt"))
    if not reports:
        return "", {}
    path = reports[-1]
    sections: dict[str, dict[str, str]] = {}
    current: dict[str, str] | None = None
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        section = SECTION_RE.match(line.strip()) if line.strip().endswith(":") else None
        if section:
            current = sections.setdefault(section.group(1).strip(), {})
            continue
        entry = ENTRY_RE.match(line)
        if entry and current is not None:
            current[entry.group(1).strip()] = entry.group(2).strip()
    return path.name, sections


def as_float(value: str) -> float | None:
    try:
        return float(value)
    except ValueError:
        return None


def main() -> None:
    for run in (BASELINE, CONTROL):
        print(f"{'baseline' if run is BASELINE else 'control '}: {run.name}  exists={run.is_dir()}")
    base_name, base = parse_report(BASELINE)
    ctrl_name, ctrl = parse_report(CONTROL)
    print(f"\nbaseline report: {base_name}\ncontrol  report: {ctrl_name}")

    print("\n=== PHASES (seconds) ===")
    base_phases = base.get("PHASES", {})
    ctrl_phases = ctrl.get("PHASES", {})
    names = sorted(
        set(base_phases) | set(ctrl_phases),
        key=lambda n: -(as_float(base_phases.get(n, "0")) or 0.0),
    )
    total_b = total_c = 0.0
    print(f"{'фаза':<22}{'baseline':>12}{'control':>12}{'дельта':>12}")
    for name in names:
        b = as_float(base_phases.get(name, "")) or 0.0
        c = as_float(ctrl_phases.get(name, "")) or 0.0
        total_b += b
        total_c += c
        print(f"{name:<22}{b:>12.3f}{c:>12.3f}{c - b:>12.3f}")
    print(f"{'ИТОГО':<22}{total_b:>12.3f}{total_c:>12.3f}{total_c - total_b:>12.3f}")
    if total_c:
        print(f"ускорение: {total_b / total_c:.1f}x")

    print("\n=== COUNTERS diff ===")
    mismatches = 0
    for section in ("COUNTERS", "LOADER COUNTERS"):
        base_counters = base.get(section, {})
        ctrl_counters = ctrl.get(section, {})
        for name in sorted(set(base_counters) | set(ctrl_counters)):
            b = base_counters.get(name, "<нет>")
            c = ctrl_counters.get(name, "<нет>")
            if b != c:
                mismatches += 1
                print(f"  РАСХОЖДЕНИЕ [{section}] {name}: baseline={b} control={c}")
    print(f"  расхождений: {mismatches}" if mismatches else "  все счётчики совпали")

    print("\n=== xlsx ===")
    for run in (BASELINE, CONTROL):
        for path in sorted(run.glob("Шаг4_Сопоставление_RFP_MTO_*.xlsx")):
            print(f"  {run.name}\\{path.name}  ({path.stat().st_size:,} B)")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")
    main()
