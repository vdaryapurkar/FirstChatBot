"""Build the downloadable output workbook: deterministic comparison/triage
tables (computed in xlsx_ingest.py) plus Claude's root-cause narrative
(from claude_client.py), formatted as a professional report.
"""

from __future__ import annotations

from datetime import datetime, timezone

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from core.xlsx_ingest import (
    CATEGORY_BUY_SELL,
    CATEGORY_OTHER,
    ordered_mismatchtypes,
)

FONT_NAME = "Arial"
HEADER_FONT = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="305496")
FALSE_FILL = PatternFill("solid", fgColor="E2EFDA")
GROUP_FILL = PatternFill("solid", fgColor="D9D9D9")

# One fill per mismatchtype so the different "issues" within a column's sheet
# are visually distinguishable, not just labeled. Unrecognized mismatchtype
# values fall back to the "mismatch" fill.
MISMATCHTYPE_FILLS = {
    "mismatch": PatternFill("solid", fgColor="FCE4E4"),           # red-ish: genuine value diff
    "new_post": PatternFill("solid", fgColor="DCE6F7"),           # blue-ish: new position in post
    "missing_position_post": PatternFill("solid", fgColor="FDEBD3"),  # amber-ish: dropped position
}
DEFAULT_MISMATCHTYPE_FILL = MISMATCHTYPE_FILLS["mismatch"]
MISMATCHTYPE_LABELS = {
    "mismatch": "Mismatch",
    "new_post": "New position (post only)",
    "missing_position_post": "Missing position (post)",
}
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True)
SUBTITLE_FONT = Font(name=FONT_NAME, size=11, italic=True, color="595959")
BODY_FONT = Font(name=FONT_NAME, size=10)
BOLD_BODY_FONT = Font(name=FONT_NAME, size=10, bold=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(vertical="top", wrap_text=True)


def _style_header_row(ws, row: int, n_cols: int):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _autosize(ws, widths: list[int]):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_report(triage: dict[str, dict], structural: dict[str, dict], analysis: dict,
                  sources: list[str], output_path: str,
                  process_types_by_file: dict[str, str] | None = None):
    """triage: the FULL, uncapped per-column output of xlsx_ingest.compute_triage
    (every row, not the token-budget-limited sample that was sent to Claude).
    structural: the FULL, uncapped {mismatchtype: {category: [rows]}} output of
    compute_triage -- the pooled/deduped "new_post"/"missing_position_post" rows.
    analysis: the parsed tool-call dict from claude_client.run_analysis (the
    'result' key: summary/triage_categories/root_causes/recommendations).
    sources: list of original uploaded file names included in this report.
    process_types_by_file: {source_file: "Valuation"|"Settlement"|"NetValuation"|"Credit"|"Unknown"},
    from xlsx_ingest.build_data_digest()'s "process_types_by_file".
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # A column with zero TRUE rows reconciled cleanly (e.g. exposureqty_mismatch,
    # expprice_mismatch on a run where those never differ) -- no issue to
    # report, so skip its detail sheet and its Summary row entirely rather
    # than padding the workbook with tabs that only ever show "No issue".
    # "true_groups" (not "true_rows") drives this: it already excludes the
    # structural new_post/missing_position_post rows, which get their own
    # pooled sheets below instead of a per-column one.
    flagged = {col: issue for col, issue in triage.items() if issue["true_groups"]}
    clean_column_count = sum(1 for issue in triage.values() if not issue["true_rows"])

    _build_summary_sheet(wb, triage, structural, analysis, sources,
                          process_types_by_file or {}, clean_column_count)
    _build_root_cause_sheet(wb, analysis)
    for col, issue in flagged.items():
        _build_issue_sheet(wb, col, issue)
    for mismatchtype in _ordered_structural_types(structural):
        _build_structural_sheet(wb, mismatchtype, structural[mismatchtype])

    wb.save(output_path)


def _ordered_structural_types(structural: dict[str, dict]) -> list[str]:
    """new_post, then missing_position_post, then anything else -- only the
    ones actually present in this run's data."""
    present = set(structural.keys())
    ordered = [mt for mt in ("new_post", "missing_position_post") if mt in present]
    ordered += sorted(present - set(ordered))
    return ordered


def _ordered_categories(categories: dict[str, list]) -> list[str]:
    """BUY/SELL first, then Other, then any other category value found."""
    present = set(categories.keys())
    ordered = [c for c in (CATEGORY_BUY_SELL, CATEGORY_OTHER) if c in present]
    ordered += sorted(present - set(ordered))
    return ordered


def _build_summary_sheet(wb, triage, structural, analysis, sources, process_types_by_file, clean_column_count=0):
    ws = wb.create_sheet("Summary")
    ws.sheet_view.showGridLines = False

    ws["A1"] = "Recon-CI Triage Analysis - Summary"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Generated {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}  |  Sources: {', '.join(sources)}"
    ws["A2"].font = SUBTITLE_FONT

    r = 4
    ws.cell(row=r, column=1, value="Source file")
    ws.cell(row=r, column=2, value="Process type")
    _style_header_row(ws, r, 2)
    r += 1
    for src in sources:
        ws.cell(row=r, column=1, value=src).font = BODY_FONT
        ws.cell(row=r, column=2, value=process_types_by_file.get(src, "Unknown")).font = BODY_FONT
        for j in (1, 2):
            ws.cell(row=r, column=j).border = BORDER
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="Executive Summary").font = BOLD_BODY_FONT
    r += 1
    summary_start = r
    summary_end = summary_start + 4
    ws.merge_cells(start_row=summary_start, start_column=1, end_row=summary_end, end_column=6)
    cell = ws.cell(row=summary_start, column=1)
    cell.value = analysis.get("summary", "")
    cell.font = BODY_FONT
    cell.alignment = WRAP
    r = summary_end + 1

    issue_count = sum(len(issue["true_groups"]) for issue in triage.values())
    issue_count += sum(len(categories) for categories in structural.values())
    bug_report = analysis.get("bug_report") or {}
    if issue_count > 1 and (bug_report.get("synopsis") or bug_report.get("description")):
        r += 1
        ws.cell(row=r, column=1,
                value="Bug Report (copy into TFS to file an investigation ticket)").font = BOLD_BODY_FONT
        r += 1
        ws.cell(row=r, column=1, value="Synopsis").font = BOLD_BODY_FONT
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        c = ws.cell(row=r, column=1, value=bug_report.get("synopsis", ""))
        c.font = BODY_FONT
        c.alignment = WRAP
        r += 1
        ws.cell(row=r, column=1, value="Description").font = BOLD_BODY_FONT
        r += 1
        desc_start = r
        desc_end = desc_start + 9
        ws.merge_cells(start_row=desc_start, start_column=1, end_row=desc_end, end_column=6)
        c = ws.cell(row=desc_start, column=1, value=bug_report.get("description", ""))
        c.font = BODY_FONT
        c.alignment = WRAP
        r = desc_end + 1

    if clean_column_count:
        r += 1
        plural = "s" if clean_column_count != 1 else ""
        ws.cell(row=r, column=1,
                value=(f"{clean_column_count} other mismatch column{plural} reconciled cleanly "
                       "(no differences found) and are omitted from this report.")).font = SUBTITLE_FONT
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        r += 1

    headers = ["Issue (column(s))", "Mismatch Type", "Category", "Count", "Process Type(s)", "Description"]
    header_row = r + 1
    for j, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=j, value=h)
    _style_header_row(ws, header_row, len(headers))

    cat_by_key = {
        (c.get("column"), c.get("mismatchtype")): c for c in analysis.get("triage_categories", [])
        if c.get("column")
    }
    struct_cat_by_key = {
        (c.get("mismatchtype"), c.get("category")): c for c in analysis.get("triage_categories", [])
        if c.get("category") and not c.get("column")
    }

    # Column-specific issues (mostly "mismatch") stay one row per (column,
    # mismatchtype). Rows with no issue at all ("No issue" / all-False) are
    # not reconciliation breaks and are left out of this table entirely.
    r = header_row + 1
    for col, issue in triage.items():
        for mt in ordered_mismatchtypes(issue["true_groups"]):
            rows = issue["true_groups"][mt]
            cat = cat_by_key.get((col, mt), {})
            process_types = sorted({row.get("_process_type", "Unknown") for row in rows})
            row_vals = [
                col,
                MISMATCHTYPE_LABELS.get(mt, mt),
                "-",
                len(rows),
                ", ".join(process_types),
                cat.get("description", ""),
            ]
            for j, v in enumerate(row_vals, start=1):
                c = ws.cell(row=r, column=j, value=v)
                c.font = BODY_FONT
                c.border = BORDER
                c.alignment = WRAP
            r += 1

    # Structural issues (new_post / missing_position_post) are pooled across
    # every column they triggered and shown once per (mismatchtype,
    # category) instead of once per column -- see xlsx_ingest.compute_triage.
    for mt in _ordered_structural_types(structural):
        categories = structural[mt]
        for category in _ordered_categories(categories):
            rows = categories[category]
            triggered_columns = sorted({c for row in rows for c in row.get("_triggered_columns", [])})
            cat = struct_cat_by_key.get((mt, category), {})
            process_types = sorted({row.get("_process_type", "Unknown") for row in rows})
            row_vals = [
                ", ".join(triggered_columns),
                MISMATCHTYPE_LABELS.get(mt, mt),
                category,
                len(rows),
                ", ".join(process_types),
                cat.get("description", ""),
            ]
            for j, v in enumerate(row_vals, start=1):
                c = ws.cell(row=r, column=j, value=v)
                c.font = BODY_FONT
                c.border = BORDER
                c.alignment = WRAP
            r += 1

    _autosize(ws, [24, 26, 12, 10, 24, 60])
    ws.row_dimensions[1].height = 22


def _build_root_cause_sheet(wb, analysis):
    ws = wb.create_sheet("Root Cause Analysis")
    ws.sheet_view.showGridLines = False

    ws["A1"] = "Root Cause Analysis"
    ws["A1"].font = TITLE_FONT

    headers = ["Issue", "Column", "Mismatch Type", "Category", "Process Type", "Explanation", "Evidence", "Affected Scope", "Confidence"]
    header_row = 3
    for j, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=j, value=h)
    _style_header_row(ws, header_row, len(headers))

    r = header_row + 1
    for rc in analysis.get("root_causes", []):
        evidence = "\n".join(f"- {e}" for e in rc.get("evidence", []))
        row_vals = [
            rc.get("issue", ""),
            rc.get("column", ""),
            MISMATCHTYPE_LABELS.get(rc.get("mismatchtype", ""), rc.get("mismatchtype", "")),
            rc.get("category", "") or "-",
            rc.get("process_type", ""),
            rc.get("explanation", ""),
            evidence,
            rc.get("affected_scope", ""),
            rc.get("confidence", ""),
        ]
        for j, v in enumerate(row_vals, start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = BODY_FONT
            c.border = BORDER
            c.alignment = WRAP
        ws.row_dimensions[r].height = 90
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="Recommendations").font = BOLD_BODY_FONT
    r += 1
    for rec in analysis.get("recommendations", []):
        ws.cell(row=r, column=1, value=f"- {rec}").font = BODY_FONT
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
        ws.cell(row=r, column=1).alignment = WRAP
        r += 1

    _autosize(ws, [28, 16, 22, 12, 16, 45, 40, 22, 12])


def _build_issue_sheet(wb, col: str, issue: dict):
    sheet_name = col[:31] if len(col) <= 31 else col[:28] + "..."
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"

    pre_col, post_col = issue["pre_col"], issue["post_col"]

    ws["A1"] = f"Triage: {col}"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Rows grouped by mismatch type (each is a distinct issue), then FALSE (no issue)."
    ws["A2"].font = SUBTITLE_FONT

    true_rows = issue["true_rows"]
    false_rows = issue["false_rows"]
    mismatchtypes = ordered_mismatchtypes(issue["true_groups"])

    id_cols: list[str] = []
    seen = set()
    for row in (true_rows + false_rows):
        for k in row.keys():
            if k in (col, pre_col, post_col, "_source_file", "_sheet_name", "_process_type",
                      "_mismatchtype", "_category", "_triggered_columns", "_difference"):
                continue
            if k.startswith("pre_") or k.startswith("post_") or k.lower().endswith("_mismatch"):
                continue
            if k not in seen:
                seen.add(k)
                id_cols.append(k)

    headers = ["source_file", "sheet_name", "process_type"] + id_cols
    if pre_col:
        headers += [pre_col, post_col, "difference (post - pre)"]
    headers += [col]

    header_row = 4
    for j, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=j, value=h)
    _style_header_row(ws, header_row, len(headers))

    r = header_row + 1
    groups = [(issue["true_groups"][mt], MISMATCHTYPE_FILLS.get(mt, DEFAULT_MISMATCHTYPE_FILL),
               f"ISSUE - {MISMATCHTYPE_LABELS.get(mt, mt)}") for mt in mismatchtypes]
    groups.append((false_rows, FALSE_FILL, "FALSE - no issue"))

    for group_rows, fill, label in groups:
        if not group_rows:
            continue
        ws.cell(row=r, column=1, value=f"{label}  ({len(group_rows)} rows)").font = BOLD_BODY_FONT
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(headers))
        ws.cell(row=r, column=1).fill = GROUP_FILL
        r += 1
        for row in group_rows:
            values = [row.get("_source_file", ""), row.get("_sheet_name", ""), row.get("_process_type", "Unknown")]
            values += [row.get(k, "") for k in id_cols]
            pre_idx = post_idx = diff_idx = None
            if pre_col:
                pre_idx = len(values) + 1
                post_idx = pre_idx + 1
                diff_idx = post_idx + 1
                values += [row.get(pre_col), row.get(post_col), None]
            values += [row.get(col)]
            for j, v in enumerate(values, start=1):
                c = ws.cell(row=r, column=j, value=v)
                c.font = BODY_FONT
                c.border = BORDER
                c.fill = fill
            if pre_col:
                c = ws.cell(row=r, column=diff_idx,
                             value=f"={get_column_letter(post_idx)}{r}-{get_column_letter(pre_idx)}{r}")
                c.font = BODY_FONT
                c.border = BORDER
                c.fill = fill
                c.number_format = "0.0000"
            r += 1

    _autosize(ws, [16, 14, 14] + [14] * len(id_cols) + ([16, 16, 16] if pre_col else []) + [10])


def _build_structural_sheet(wb, mismatchtype: str, categories: dict[str, list]):
    """One sheet per structural mismatchtype (new_post / missing_position_post),
    with rows grouped by position-type category (BUY/SELL vs Other) instead of
    duplicated once per mismatch column -- see xlsx_ingest.compute_triage."""
    label = MISMATCHTYPE_LABELS.get(mismatchtype, mismatchtype)
    ws = wb.create_sheet(label[:31])
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"

    ws["A1"] = f"Triage: {label}"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Each position appears once, pooled across every mismatch column it "
                "triggered, grouped by position type.")
    ws["A2"].font = SUBTITLE_FONT

    all_rows = [row for rows in categories.values() for row in rows]
    id_cols: list[str] = []
    seen = set()
    for row in all_rows:
        for k in row.keys():
            if k in ("_source_file", "_sheet_name", "_process_type", "_mismatchtype",
                      "_category", "_triggered_columns", "_difference"):
                continue
            if k.lower().endswith("_mismatch"):
                continue
            if k not in seen:
                seen.add(k)
                id_cols.append(k)

    headers = ["source_file", "sheet_name", "process_type", "triggered_columns"] + id_cols
    header_row = 4
    for j, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=j, value=h)
    _style_header_row(ws, header_row, len(headers))

    fill = MISMATCHTYPE_FILLS.get(mismatchtype, DEFAULT_MISMATCHTYPE_FILL)
    r = header_row + 1
    for category in _ordered_categories(categories):
        rows = categories[category]
        if not rows:
            continue
        ws.cell(row=r, column=1, value=f"CATEGORY - {category}  ({len(rows)} rows)").font = BOLD_BODY_FONT
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(headers))
        ws.cell(row=r, column=1).fill = GROUP_FILL
        r += 1
        for row in rows:
            values = [
                row.get("_source_file", ""),
                row.get("_sheet_name", ""),
                row.get("_process_type", "Unknown"),
                ", ".join(row.get("_triggered_columns", [])),
            ]
            values += [row.get(k, "") for k in id_cols]
            for j, v in enumerate(values, start=1):
                c = ws.cell(row=r, column=j, value=v)
                c.font = BODY_FONT
                c.border = BORDER
                c.fill = fill
            r += 1

    _autosize(ws, [16, 14, 14, 30] + [14] * len(id_cols))
