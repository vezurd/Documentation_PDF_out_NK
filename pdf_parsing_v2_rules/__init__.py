"""pdf_parsing_v2_rules — normcontrol checks for v2 pipeline."""

from __future__ import annotations

from pdf_parsing_v2_rules.profile import RulesProfile, get_profile

__all__ = ["RulesProfile", "get_profile", "rules_check_start_v2"]


def rules_check_start_v2(
    curr_proj,
    proj_od_list,
    pdf_path: str,
    path_out_dir: str,
    *,
    profile: RulesProfile | None = None,
    effective_project: str | None = None,
    sanitize_illegal_chars: bool = True,
    od_warnings: list[str] | None = None,
    od_pdf_path: str = "",
) -> dict:
    """Run normcontrol checks, write Excel report, and return structured data."""
    if profile is None:
        profile = get_profile(effective_project)
    from pdf_parsing_v2_rules.checks import run_all_checks
    from pdf_parsing_v2_rules.extraction_checks import build_extraction_check_rows
    from pdf_parsing_v2_rules.od_table_checks import build_od_table_warning_rows
    from pdf_parsing_v2_rules.output import (
        build_normcontrol_report,
        excel_check_list_out,
        print_check_list,
    )

    rules_rows = run_all_checks(
        curr_proj, proj_od_list, profile, od_pdf_path=str(od_pdf_path or "").strip()
    )
    extraction_rows = build_extraction_check_rows(curr_proj)
    od_rows = build_od_table_warning_rows(od_warnings, od_pdf_path=od_pdf_path)
    check_list = rules_rows + extraction_rows + od_rows
    flagged = print_check_list(check_list, pdf_path, 0, 1005)
    report = build_normcontrol_report(check_list, flagged, curr_proj)
    report["xlsx_path"] = excel_check_list_out(
        flagged,
        check_list,
        curr_proj,
        path_out_dir,
        sanitize_illegal_chars=sanitize_illegal_chars,
    )
    return report
