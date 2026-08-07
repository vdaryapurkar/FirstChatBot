"""Build the downloadable output workbook: deterministic comparison/triage
tables (computed in xlsx_ingest.py) plus Claude's root-cause narrative
(from claude_client.py), formatted as a professional report.
"""

from __future__ import annotations

from datetime import datetime, timezone

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT_NAME = "Arial"
HEADER_FONT = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="305496")
TRUE_FILL = PatternFill("solid", fgColor="FCE4E4")
FALSE_FILL = PatternFill("solid", fgColor="E2EFDA")
GROUP_FILL = PatternFill("solid", fgColor="D9D9D9")
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


def build_report(triage: dict[str, dict], analysis: dict, sources: list[str], output_path: str):
    """triage: the FULL, uncapped output of xlsx_ingest.compute_triage (every
    row, not the token-budget-limited sample that was sent to Claude).
    analysis: the parsed tool-call dict from claude_client.run_analysis (the
    'result' key: summary/triage_categories/root_causes/recommendations).
    sources: list of original uploaded file names included in this report.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _build_summary_sheet(wb, triage, analysis, sources)
    _build_root_cause_sheet(wb, analysis)
    for col, issue in triage.items():
        _build_issue_sheet(wb, col, issue)

    wb.save(output_path)


def _build_summary_sheet(wb, triage, analysis, sources):
    ws = wb.create_sheet("Summary")
    ws.sheet_view.showGridLines = False

    ws["A1"] = "Recon-Ci Triage Analysis - Summary"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Generated {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}  |  Sources: {', '.join(sources)}"
    ws["A2"].font = SUBTITLE_FONT

    ws["A4"] = "Executive Summary"
    ws["A4"].font = BOLD_BODY_FONT
    ws.merge_cells("A5:F9")
    cell = ws["A5"]
    cell.value = analysis.get("summary", "")
    cell.font = BODY_FONT
    cell.alignment = WRAP

    headers = ["Issue (column)", "TRUE (issue present)", "FALSE (no issue)", "Total", "Description"]
    header_row = 11
    for j, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=j, value=h)
    _style_header_row(ws, header_row, len(headers))

    cat_by_col = {c["column"]: c for c in analysis.get("triage_categories", [])}
    r = header_row + 1
    for col, issue in triage.items():
        cat = cat_by_col.get(col, {})
        true_n, false_n = len(issue["true_rows"]), len(issue["false_rows"])
        row_vals = [
            col,
            true_n,
            false_n,
            true_n + false_n,
            cat.get("description", ""),
        ]
        for j, v in enumerate(row_vals, start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = BODY_FONT
            c.border = BORDER
            c.alignment = WRAP
        r += 1

    _autosize(ws, [22, 18, 16, 10, 70])
    ws.row_dimensions[1].height = 22


def _build_root_cause_sheet(wb, analysis):
    ws = wb.create_sheet("Root Cause Analysis")
    ws.sheet_view.showGridLines = False

    ws["A1"] = "Root Cause Analysis"
    ws["A1"].font = TITLE_FONT

    headers = ["Issue", "Explanation", "Evidence", "Affected Scope", "Confidence"]
    header_row = 3
    for j, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=j, value=h)
    _style_header_row(ws, header_row, len(headers))

    r = header_row + 1
    for rc in analysis.get("root_causes", []):
        evidence = "\n".join(f"- {e}" for e in rc.get("evidence", []))
        row_vals = [
            rc.get("issue", ""),
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
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.cell(row=r, column=1).alignment = WRAP
        r += 1

    _autosize(ws, [28, 45, 40, 22, 12])


def _build_issue_sheet(wb, col: str, issue: dict):
    sheet_name = col[:31] if len(col) <= 31 else col[:28] + "..."
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"

    pre_col, post_col = issue["pre_col"], issue["post_col"]

    ws["A1"] = f"Triage: {col}"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "Rows grouped by TRUE (issue present) then FALSE (no issue)."
    ws["A2"].font = SUBTITLE_FONT

    true_rows = issue["true_rows"]
    false_rows = issue["false_rows"]

    id_cols: list[str] = []
    seen = set()
    for row in (true_rows + false_rows):
        for k in row.keys():
            if k in (col, pre_col, post_col, "_source_file", "_sheet_name", "_difference"):
                continue
            if k.startswith("pre_") or k.startswith("post_") or k.lower().endswith("_mismatch"):
                continue
            if k not in seen:
                seen.add(k)
                id_cols.append(k)

    headers = ["source_file", "sheet_name"] + id_cols
    if pre_col:
        headers += [pre_col, post_col, "difference (post - pre)"]
    headers += [col]

    header_row = 4
    for j, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=j, value=h)
    _style_header_row(ws, header_row, len(headers))

    r = header_row + 1
    for group_rows, fill, label in ((true_rows, TRUE_FILL, "TRUE - issue present"),
                                     (false_rows, FALSE_FILL, "FALSE - no issue")):
        if not group_rows:
            continue
        ws.cell(row=r, column=1, value=f"{label}  ({len(group_rows)} rows)").font = BOLD_BODY_FONT
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(headers))
        ws.cell(row=r, column=1).fill = GROUP_FILL
        r += 1
        for row in group_rows:
            values = [row.get("_source_file", ""), row.get("_sheet_name", "")]
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

    _autosize(ws, [16, 14] + [14] * len(id_cols) + ([16, 16, 16] if pre_col else []) + [10])
