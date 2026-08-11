"""Ingest one or more uploaded .xlsx files (including files pulled from a
folder upload) and compute deterministic comparison/triage statistics.

Design note: all arithmetic (row counts, True/False splits, numeric
pre/post differences) is computed here in Python over the FULL dataset, not
by Claude. Claude is only asked to explain/interpret this already-correct
data -- see core/claude_client.py.
"""

from __future__ import annotations

import math
from pathlib import Path
from typing import Any

import pandas as pd

MISMATCH_SUFFIX = "_mismatch"
MISMATCHTYPE_COLUMN = "mismatchtype"

# mismatchtype values (from the "mismatchtype" column, if present):
# - "mismatch": a genuine value difference between pre and post for a
#   position present on both sides -- the real data-quality issue.
# - "new_post": the position is newly added in post and absent from pre;
#   pre_* columns are expected to read 0/blank as a consequence, not as an
#   error in their own right.
# - "missing_position_post" (also written "missing_post" in some exports,
#   treated identically): the position was present in pre but is absent
#   from post; post_* columns are expected to read 0/blank as a
#   consequence.
# Rows with no recognizable mismatchtype default to "mismatch".
MISMATCHTYPE_ALIASES = {"missing_post": "missing_position_post"}
KNOWN_MISMATCHTYPES = ["mismatch", "new_post", "missing_position_post"]

# "new_post" and "missing_position_post" are structural facts about a
# position (it was added or dropped between pre and post), not a per-column
# data-quality issue -- the same row trips every "*_mismatch" column at once
# (pre_* or post_* all read 0 by construction), so bucketing them per column
# just repeats the same position N times under N different column headings.
# These two types are pooled across all columns, deduped down to one entry
# per physical row, and categorized by position type -- see compute_triage().
STRUCTURAL_MISMATCHTYPES = {"new_post", "missing_position_post"}

POSITIONTYPE_COLUMN = "positiontype"
BUY_SELL_POSITIONTYPES = {"buy", "sell"}
CATEGORY_BUY_SELL = "BUY/SELL"
CATEGORY_OTHER = "Other"

# "mismatch" (and any other/unrecognized mismatchtype) rows are ALSO pooled
# across columns the same way -- a single row's quantity difference cascades
# into any column derived from quantity (e.g. pricequantitysum), and a row's
# row-count change (source-side split/merge of detail rows) cascades into
# any column that aggregates across rows (e.g. debit/credit/grossvalue), even
# when the row's own quantity didn't change. When several "*_mismatch"
# columns are True together for the same physical row, that's one issue, not
# one per column -- report it under whichever of these two "driving" columns
# is present, since that's the one actually worth investigating; the rest
# are listed as "also triggered" for traceability. Falls back to the
# alphabetically-first triggered column when neither driver is present, so a
# genuinely standalone column mismatch still gets its own issue as before.
MISMATCH_PRIMARY_COLUMN_PRIORITY = ["qty_mismatch", "rowcountsum_mismatch"]

# Process-type classification rules. Every uploaded workbook is the result of
# one of four reconciliation processes, identified from its filename and
# (for Settlement) its columns. See detect_process_type() below and
# config/rules.py for how this is explained to Claude.
SETTLEMENT_COLUMNS = {
    "pre_debitsum", "post_debitsum", "debit_mismatch",
    "pre_creditsum", "post_creditsum", "credit_mismatch",
}
PROCESS_TYPES = ("Settlement", "Valuation", "NetValuation", "Credit")


def detect_process_type(filename: str, columns: list[str]) -> str:
    """Classify which of the four reconciliation processes a file's data
    belongs to, purely from its filename and column names:

    - "findetail" in filename AND the debit/credit sum + mismatch columns
      are present -> Settlement
    - "valdetail" in filename -> Valuation
    - "netval" in filename -> NetValuation
    - "credit" in filename -> Credit
    - otherwise -> "Unknown"
    """
    name = filename.lower()
    cols = {str(c).lower() for c in columns}
    if "findetail" in name and SETTLEMENT_COLUMNS.issubset(cols):
        return "Settlement"
    if "valdetail" in name:
        return "Valuation"
    if "netval" in name:
        return "NetValuation"
    if "credit" in name:
        return "Credit"
    return "Unknown"


def read_workbook(path: str | Path) -> dict[str, pd.DataFrame]:
    """Read every sheet of an xlsx file into a dict of {sheet_name: DataFrame}."""
    return pd.read_excel(path, sheet_name=None, dtype=object)


def ingest_files(file_paths: list[tuple[str, str]]) -> list[dict]:
    """file_paths: list of (original_name, stored_path). Returns a flat list of
    {source_file, sheet_name, df, process_type} for every sheet in every file."""
    tables = []
    for original_name, stored_path in file_paths:
        sheets = read_workbook(stored_path)
        for sheet_name, df in sheets.items():
            if df.empty:
                continue
            df = df.dropna(how="all")
            process_type = detect_process_type(original_name, [str(c) for c in df.columns])
            tables.append({
                "source_file": original_name,
                "sheet_name": sheet_name,
                "df": df,
                "process_type": process_type,
            })
    return tables


def _normalize_mismatchtype(value: Any) -> str:
    if value is None:
        return "mismatch"
    s = str(value).strip().lower()
    if not s or s == "nan":
        return "mismatch"
    return MISMATCHTYPE_ALIASES.get(s, s)


def _find_column_ci(columns: list[str], name: str) -> str | None:
    for c in columns:
        if c.lower() == name:
            return c
    return None


def _find_mismatchtype_column(columns: list[str]) -> str | None:
    return _find_column_ci(columns, MISMATCHTYPE_COLUMN)


def _positiontype_category(row, positiontype_col: str | None) -> str:
    """BUY/SELL positiontype -> "BUY/SELL"; anything else (including
    missing/unrecognized) -> "Other". This is the categorization rule for
    structural (new_post/missing_position_post) issues -- see
    STRUCTURAL_MISMATCHTYPES above."""
    if not positiontype_col:
        return CATEGORY_OTHER
    value = row[positiontype_col]
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return CATEGORY_OTHER
    if str(value).strip().lower() in BUY_SELL_POSITIONTYPES:
        return CATEGORY_BUY_SELL
    return CATEGORY_OTHER


def _primary_column(triggered_columns: list[str]) -> str:
    """Which column a non-structural (mismatch/other) issue should be
    reported under when several columns triggered together on the same row
    -- see MISMATCH_PRIMARY_COLUMN_PRIORITY above."""
    for c in MISMATCH_PRIMARY_COLUMN_PRIORITY:
        if c in triggered_columns:
            return c
    return sorted(triggered_columns)[0]


def _is_true(value: Any) -> bool | None:
    if isinstance(value, bool):
        return value
    if value is None:
        return None
    s = str(value).strip().lower()
    if s in ("true", "1", "yes"):
        return True
    if s in ("false", "0", "no"):
        return False
    return None


def _base_variants(base: str) -> list[str]:
    """Common abbreviation swaps seen between a '*_mismatch' column name and
    its paired 'pre_'/'post_' column (e.g. "qty" vs "quantity")."""
    variants = {base}
    for short, long in (("qty", "quantity"), ("amt", "amount"), ("val", "value")):
        variants.add(base.replace(short, long))
        variants.add(base.replace(long, short))
    return list(variants)


def _find_pre_post_pair(columns: list[str], base: str) -> tuple[str | None, str | None]:
    variants = [v.lower() for v in _base_variants(base)]
    pre = [c for c in columns if c.lower().startswith("pre_") and any(v in c.lower() for v in variants)]
    post = [c for c in columns if c.lower().startswith("post_") and any(v in c.lower() for v in variants)]
    if pre and post:
        return sorted(pre, key=len)[0], sorted(post, key=len)[0]
    return None, None


def find_mismatch_columns(tables: list[dict]) -> list[str]:
    cols: set[str] = set()
    for t in tables:
        cols.update(str(c) for c in t["df"].columns)
    return sorted(c for c in cols if c.lower().endswith(MISMATCH_SUFFIX))


def compute_triage(tables: list[dict]) -> tuple[dict[str, dict], dict[str, dict]]:
    """Returns (triage, grouped).

    triage: per mismatch-flag column, the matching pre/post pair (if found),
    FALSE (no issue) rows, and an unclassified count. "true_rows" is every
    row that tripped this column, for a raw per-column count -- but every
    row that tripped ANY column is pooled/deduped across columns (see
    "grouped" below), since the same physical row commonly trips several
    "*_mismatch" columns at once for one underlying reason. Use "grouped",
    not a per-column split, for the actual issue breakdown.

    grouped: {mismatchtype: {group_key: [rows]}}, one entry per distinct
    issue, each row appearing exactly once even if it tripped multiple
    columns. group_key means different things by mismatchtype:

    - "new_post" / "missing_position_post": group_key is a position-type
      category, "BUY/SELL" or "Other" (see _positiontype_category()) --
      these are structural facts about a position, not a per-column
      data-quality issue, so every column they trip is pooled together.
    - anything else (usually "mismatch"): group_key is the "primary" column
      for that row -- usually just the one column that tripped, but when
      several trip together for the same row (e.g. a quantity difference
      also moving a derived column, or a row-count change also moving every
      aggregated column) they're pooled under one driving column instead of
      reported as separate issues (see MISMATCH_PRIMARY_COLUMN_PRIORITY).

    Every row in "grouped" carries "_triggered_columns" (every column it
    tripped) and "_differences" ({column: post-pre}, for whichever triggered
    columns have a resolvable pre/post pair), for traceability.
    """
    mismatch_cols = find_mismatch_columns(tables)
    result: dict[str, dict] = {
        col: {"pre_col": None, "post_col": None, "true_rows": [], "false_rows": [], "unclassified": 0}
        for col in mismatch_cols
    }
    pooled: dict[tuple, dict] = {}

    for t in tables:
        df, source_file, sheet_name = t["df"], t["source_file"], t["sheet_name"]
        process_type = t.get("process_type", "Unknown")
        columns = [str(c) for c in df.columns]
        mismatchtype_col = _find_mismatchtype_column(columns)
        positiontype_col = _find_column_ci(columns, POSITIONTYPE_COLUMN)

        for col in mismatch_cols:
            if col not in df.columns:
                continue
            base = col[: -len(MISMATCH_SUFFIX)]
            pre_col, post_col = _find_pre_post_pair(columns, base)
            if pre_col and result[col]["pre_col"] is None:
                result[col]["pre_col"] = pre_col
                result[col]["post_col"] = post_col

            for idx, row in df.iterrows():
                verdict = _is_true(row[col])
                mismatchtype = _normalize_mismatchtype(
                    row[mismatchtype_col] if mismatchtype_col else None
                )

                difference = None
                if pre_col and post_col:
                    try:
                        difference = float(row[post_col]) - float(row[pre_col])
                    except (TypeError, ValueError):
                        difference = None

                if verdict is True:
                    row_key = (source_file, sheet_name, idx)
                    if row_key in pooled:
                        record = pooled[row_key]
                        record["_triggered_columns"].append(col)
                        record["_differences"][col] = difference
                    else:
                        record = row.to_dict()
                        record["_source_file"] = source_file
                        record["_sheet_name"] = sheet_name
                        record["_process_type"] = process_type
                        record["_mismatchtype"] = mismatchtype
                        record["_positiontype_category"] = _positiontype_category(row, positiontype_col)
                        record["_triggered_columns"] = [col]
                        record["_differences"] = {col: difference}
                        pooled[row_key] = record
                    result[col]["true_rows"].append(pooled[row_key])
                    continue

                record = row.to_dict()
                record["_source_file"] = source_file
                record["_sheet_name"] = sheet_name
                record["_process_type"] = process_type
                record["_mismatchtype"] = mismatchtype
                record["_difference"] = difference
                if verdict is False:
                    result[col]["false_rows"].append(record)
                else:
                    result[col]["unclassified"] += 1

    grouped: dict[str, dict[str, list[dict]]] = {}
    for record in pooled.values():
        mismatchtype = record["_mismatchtype"]
        if mismatchtype in STRUCTURAL_MISMATCHTYPES:
            group_key = record["_positiontype_category"]
        else:
            group_key = _primary_column(record["_triggered_columns"])
        grouped.setdefault(mismatchtype, {}).setdefault(group_key, []).append(record)

    return result, grouped


def ordered_mismatchtypes(true_groups: dict[str, list]) -> list[str]:
    """KNOWN_MISMATCHTYPES first (only the ones present), then any other
    values found in the data, alphabetically -- a stable, readable order for
    sheets/tables without hardcoding every value a source system might use."""
    present = set(true_groups.keys())
    ordered = [mt for mt in KNOWN_MISMATCHTYPES if mt in present]
    ordered += sorted(present - set(ordered))
    return ordered


def _diff_shape(rows: list[dict], column: str | None = None) -> dict:
    """Shape (min/max/mean/constant-offset) of the pre/post differences for
    these rows. column: for grouped/pooled rows (which carry a
    "_differences" dict keyed by every column they triggered), which
    column's differences to summarize; omit for plain per-column FALSE/
    unclassified rows, which carry a single scalar "_difference" instead."""
    if column is not None:
        diffs = [r["_differences"].get(column) for r in rows if r.get("_differences", {}).get(column) is not None]
    else:
        diffs = [r["_difference"] for r in rows if r.get("_difference") is not None]
    if not diffs:
        return {}
    constant = len({round(d, 6) for d in diffs}) == 1
    return {
        "min": min(diffs),
        "max": max(diffs),
        "mean": sum(diffs) / len(diffs),
        "constant_offset": constant,
        "constant_value": diffs[0] if constant else None,
    }


def build_data_digest(
    tables: list[dict], triage: dict[str, dict], grouped: dict[str, dict],
    max_sample_rows: int,
) -> dict:
    """A compact, token-budget-aware summary of the ingested data + the
    already-computed triage stats, meant to be embedded in the Claude
    prompt. Every number in here is exact (computed over the full dataset);
    only the row-level samples are capped.

    "issues" covers non-structural (usually "mismatch") issues, one entry
    per (mismatchtype, primary column) -- see compute_triage(). "column" is
    that primary column; "triggered_columns" lists every column (including
    the primary one) that tripped together on these rows, when more than
    one moved together for the same underlying reason.

    "structural_issues" covers the pooled/deduped "new_post" and
    "missing_position_post" issues, one entry per (mismatchtype, category)
    -- each listing every mismatch column it triggered, since a single
    structural row can trip several at once.
    """
    files = sorted({t["source_file"] for t in tables})
    process_types_by_file = {t["source_file"]: t.get("process_type", "Unknown") for t in tables}
    sheets = [{"source_file": t["source_file"], "sheet_name": t["sheet_name"],
               "process_type": t.get("process_type", "Unknown"),
               "rows": len(t["df"]), "columns": [str(c) for c in t["df"].columns]}
              for t in tables]

    issues = []
    structural_issues = []
    for mismatchtype in ordered_mismatchtypes(grouped):
        groups = grouped[mismatchtype]
        is_structural = mismatchtype in STRUCTURAL_MISMATCHTYPES
        for group_key in sorted(groups.keys()):
            rows = groups[group_key]
            triggered_columns = sorted({c for r in rows for c in r.get("_triggered_columns", [])})
            process_types = sorted({r.get("_process_type", "Unknown") for r in rows})
            if is_structural:
                structural_issues.append({
                    "mismatchtype": mismatchtype,
                    "category": group_key,
                    "count": len(rows),
                    "triggered_columns": triggered_columns,
                    "process_types": process_types,
                    "sample_rows": rows[:max_sample_rows],
                })
            else:
                data = triage.get(group_key, {})
                issues.append({
                    "column": group_key,
                    "mismatchtype": mismatchtype,
                    "triggered_columns": triggered_columns,
                    "pre_column": data.get("pre_col"),
                    "post_column": data.get("post_col"),
                    "count": len(rows),
                    "column_true_total": len(data.get("true_rows", [])),
                    "column_false_total": len(data.get("false_rows", [])),
                    "process_types": process_types,
                    "difference_shape": _diff_shape(rows, column=group_key),
                    "sample_rows": rows[:max_sample_rows],
                })

    return {
        "files": files,
        "process_types_by_file": process_types_by_file,
        "sheets": sheets,
        "issues": issues,
        "structural_issues": structural_issues,
    }


def _clean(v):
    if isinstance(v, float) and math.isnan(v):
        return None
    return v


def json_safe(obj):
    """Recursively replace NaN/pandas Timestamp-like values with JSON-safe types."""
    if isinstance(obj, dict):
        return {str(k): json_safe(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [json_safe(v) for v in obj]
    if isinstance(obj, float):
        return _clean(obj)
    if hasattr(obj, "isoformat"):
        return obj.isoformat()
    return obj
