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
# These two types are instead pooled across all columns, deduped down to one
# entry per physical row, and categorized by position type. "mismatch" (and
# any other/unrecognized mismatchtype value) stays column-specific, since a
# genuine value difference is still meaningfully tied to which column it's
# on -- see compute_triage().
STRUCTURAL_MISMATCHTYPES = {"new_post", "missing_position_post"}

POSITIONTYPE_COLUMN = "positiontype"
BUY_SELL_POSITIONTYPES = {"buy", "sell"}
CATEGORY_BUY_SELL = "BUY/SELL"
CATEGORY_OTHER = "Other"

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
    """Returns (triage, structural).

    triage: per mismatch-flag column, TRUE (issue present) / FALSE (no
    issue) buckets, with the matching pre/post pair + numeric difference
    when one can be found. TRUE rows with mismatchtype "mismatch" (or any
    other non-structural value) are further split into "true_groups" keyed
    by that mismatchtype -- each (column, mismatchtype) combination is a
    distinct issue. "true_rows" is the flattened union of all groups
    *including* structural rows, for callers that just want a raw count of
    what tripped this column; use "true_groups" for the actual per-issue
    breakdown, since that's the one structural rows are deliberately
    excluded from.

    structural: {mismatchtype: {category: [rows]}} for "new_post" and
    "missing_position_post" -- pooled across every column that tripped on
    the same physical row (deduped, since one new/dropped position trips
    every "*_mismatch" column on that row at once) and bucketed by position
    type: "BUY/SELL" vs "Other" (see _positiontype_category()). Each row
    carries "_triggered_columns", the sorted list of mismatch columns that
    were True for it, for traceability.
    """
    mismatch_cols = find_mismatch_columns(tables)
    result: dict[str, dict] = {
        col: {"pre_col": None, "post_col": None, "true_rows": [], "true_groups": {},
              "false_rows": [], "unclassified": 0}
        for col in mismatch_cols
    }
    structural: dict[str, dict[str, dict[tuple, dict]]] = {}

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

                if verdict is True and mismatchtype in STRUCTURAL_MISMATCHTYPES:
                    row_key = (source_file, sheet_name, idx)
                    category = _positiontype_category(row, positiontype_col)
                    bucket = structural.setdefault(mismatchtype, {}).setdefault(category, {})
                    if row_key in bucket:
                        bucket[row_key]["_triggered_columns"].append(col)
                    else:
                        record = row.to_dict()
                        record["_source_file"] = source_file
                        record["_sheet_name"] = sheet_name
                        record["_process_type"] = process_type
                        record["_mismatchtype"] = mismatchtype
                        record["_category"] = category
                        record["_triggered_columns"] = [col]
                        bucket[row_key] = record
                    result[col]["true_rows"].append(bucket[row_key])
                    continue

                record = row.to_dict()
                record["_source_file"] = source_file
                record["_sheet_name"] = sheet_name
                record["_process_type"] = process_type
                record["_mismatchtype"] = mismatchtype
                if pre_col and post_col:
                    try:
                        pre_v = float(row[pre_col])
                        post_v = float(row[post_col])
                        record["_difference"] = post_v - pre_v
                    except (TypeError, ValueError):
                        record["_difference"] = None
                if verdict is True:
                    result[col]["true_rows"].append(record)
                    result[col]["true_groups"].setdefault(mismatchtype, []).append(record)
                elif verdict is False:
                    result[col]["false_rows"].append(record)
                else:
                    result[col]["unclassified"] += 1

    structural_final = {
        mismatchtype: {category: list(rows.values()) for category, rows in categories.items()}
        for mismatchtype, categories in structural.items()
    }

    return result, structural_final


def ordered_mismatchtypes(true_groups: dict[str, list]) -> list[str]:
    """KNOWN_MISMATCHTYPES first (only the ones present), then any other
    values found in the data, alphabetically -- a stable, readable order for
    sheets/tables without hardcoding every value a source system might use."""
    present = set(true_groups.keys())
    ordered = [mt for mt in KNOWN_MISMATCHTYPES if mt in present]
    ordered += sorted(present - set(ordered))
    return ordered


def _diff_shape(rows: list[dict]) -> dict:
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
    tables: list[dict], triage: dict[str, dict], structural: dict[str, dict],
    max_sample_rows: int,
) -> dict:
    """A compact, token-budget-aware summary of the ingested data + the
    already-computed triage stats, meant to be embedded in the Claude
    prompt. Every number in here is exact (computed over the full dataset);
    only the row-level samples are capped.

    "issues" covers per-column, non-structural (mostly "mismatch") issues.
    "structural_issues" covers the pooled/deduped "new_post" and
    "missing_position_post" issues, one entry per (mismatchtype, category)
    -- see compute_triage() -- each listing every mismatch column it
    triggered, since a single structural row can trip several at once.
    """
    files = sorted({t["source_file"] for t in tables})
    process_types_by_file = {t["source_file"]: t.get("process_type", "Unknown") for t in tables}
    sheets = [{"source_file": t["source_file"], "sheet_name": t["sheet_name"],
               "process_type": t.get("process_type", "Unknown"),
               "rows": len(t["df"]), "columns": [str(c) for c in t["df"].columns]}
              for t in tables]

    issues = []
    for col, data in triage.items():
        true_n, false_n = len(data["true_rows"]), len(data["false_rows"])
        for mt in ordered_mismatchtypes(data["true_groups"]):
            rows = data["true_groups"][mt]
            process_types = sorted({r.get("_process_type", "Unknown") for r in rows})
            issues.append({
                "column": col,
                "mismatchtype": mt,
                "pre_column": data["pre_col"],
                "post_column": data["post_col"],
                "count": len(rows),
                "column_true_total": true_n,
                "column_false_total": false_n,
                "process_types": process_types,
                "difference_shape": _diff_shape(rows),
                "sample_rows": rows[:max_sample_rows],
            })

    structural_issues = []
    for mismatchtype in sorted(structural.keys()):
        categories = structural[mismatchtype]
        for category in sorted(categories.keys()):
            rows = categories[category]
            triggered_columns = sorted({c for r in rows for c in r.get("_triggered_columns", [])})
            process_types = sorted({r.get("_process_type", "Unknown") for r in rows})
            structural_issues.append({
                "mismatchtype": mismatchtype,
                "category": category,
                "count": len(rows),
                "triggered_columns": triggered_columns,
                "process_types": process_types,
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
