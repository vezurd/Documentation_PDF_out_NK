"""Machine-readable milestone output for the RFP pipeline."""

from __future__ import annotations

import json


MILESTONE_PREFIX = "@@RFP_MILESTONE "
VALID_MILESTONE_IDS = frozenset(
    {
        "prepare",
        "ds_id_check",
        "ds_mp_check",
        "parts_preflight",
        "ul_preflight",
        "rfp_load",
        "mto_google_load",
        "vo_load",
        "units_gate",
        "step4_match",
        "global_checks",
        "packing_lists",
        "save_excel",
        "bcc_accum_matrix",
        "complete",
    }
)
VALID_MILESTONE_STATES = frozenset(
    {"Waiting", "Running", "Done", "Skipped", "Error"}
)


def emit_milestone(milestone_id: str, state: str, detail: str = "") -> None:
    """Emit one validated, machine-readable pipeline milestone.

    Args:
        milestone_id: Canonical pipeline milestone identifier.
        state: Canonical milestone state.
        detail: Optional human-readable detail.

    Raises:
        ValueError: If the milestone identifier or state is unsupported.
    """
    if milestone_id not in VALID_MILESTONE_IDS:
        raise ValueError(f"Unknown RFP milestone id: {milestone_id!r}")
    if state not in VALID_MILESTONE_STATES:
        raise ValueError(f"Unknown RFP milestone state: {state!r}")

    payload = {
        "milestone_id": milestone_id,
        "state": state,
        "detail": detail,
    }
    print(
        MILESTONE_PREFIX
        + json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
        flush=True,
    )
