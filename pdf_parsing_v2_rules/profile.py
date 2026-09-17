from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class RulesProfile:
    id: str
    project_name: str
    list_of_bbb: tuple[str, ...]
    rev_literal: tuple[str, ...]
    rev_numeral: tuple[str, ...]
    text_ifc: str = "IFC - Выпущен для строительства"
    text_ifr: str = "IFR - Выпущено для рассмотрения"
    text_sup: str = "SUP - Заменен"
    text_can: str = "CAN - Аннулирован"
    enabled_checks: frozenset[int] | None = None


_AGCC_PROFILE = RulesProfile(
    id="AGCC.287",
    project_name="AGCC",
    list_of_bbb=("BOM", "BOE", "BOQ", "MTO"),
    rev_literal=("A", "B", "C", "D", "E", "F", "G"),
    rev_numeral=("0", "01", "02", "03", "04", "05", "06", "07", "08", "09",
                 "AN01", "AN02", "AN03", "AN04", "AN05"),
)

_DEFAULT_PROFILE = _AGCC_PROFILE

_REGISTRY: dict[str, RulesProfile] = {
    "AGCC.287": _AGCC_PROFILE,
    "AGCC": _AGCC_PROFILE,
}


def get_profile(project: str | None) -> RulesProfile:
    """Lookup profile; unknown/None → default with warning."""
    if not project:
        return _DEFAULT_PROFILE
    key = project.strip()
    if key in _REGISTRY:
        return _REGISTRY[key]
    for k, v in _REGISTRY.items():
        if k.lower() == key.lower():
            return v
    print(f"[rules] Unknown project {project!r}, using default profile {_DEFAULT_PROFILE.id!r}")
    return _DEFAULT_PROFILE
