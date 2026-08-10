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


def _find_mismatchtype_column(columns: list[str]) -> str | None:
    for c in columns:
        if c.lower() == MISMATCHTYPE_COLUMN:
            return c
    return None


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


def compute_triage(tables: list[dict]) -> dict[str, dict]:
    """Group every row, for every mismatch-flag column present, into TRUE
    (issue present) / FALSE (no issue) buckets, and attach the matching
    pre/post pair + numeric difference when one can be found.

    TRUE rows are further split into "true_groups" keyed by mismatchtype
    ("mismatch" / "new_post" / "missing_position_post") -- each (column,
    mismatchtype) combination is a distinct issue: a genuine value mismatch,
    a newly-added position, and a dropped position have different root
    causes and shouldn't be lumped into one bucket or one diff-shape
    calculation (new/missing rows have a pre or post side of 0, which would
    otherwise swamp the "is this a constant offset" check for real
    mismatches). "true_rows" stays as the flattened union of all groups for
    callers that just want a total count.
    """
    mismatch_cols = find_mismatch_columns(tables)
    result: dict[str, dict] = {
        col: {"pre_col": None, "post_col": None, "true_rows": [], "true_groups": {},
              "false_rows": [], "unclassified": 0}
        for col in mismatch_cols
    }

    for t in tables:
        df, source_file, sheet_name = t["df"], t["source_file"], t["sheet_name"]
        process_type = t.get("process_type", "Unknown")
        columns = [str(c) for c in df.columns]
        mismatchtype_col = _find_mismatchtype_column(columns)
        for col in mismatch_cols:
            if col not in df.columns:
                continue
            base = col[: -len(MISMATCH_SUFFIX)]
            pre_col, post_col = _find_pre_post_pair(columns, base)
            if pre_col and result[col]["pre_col"] is None:
                result[col]["pre_col"] = pre_col
                result[col]["post_col"] = post_col

            for _, row in df.iterrows():
                verdict = _is_true(row[col])
                record = row.to_dict()
                record["_source_file"] = source_file
                record["_sheet_name"] = sheet_name
                record["_process_type"] = process_type
                record["_mismatchtype"] = _normalize_mismatchtype(
                    row[mismatchtype_col] if mismatchtype_col else None
                )
                if pre_col and post_col:
                    try:
                        pre_v = float(row[pre_col])
                        post_v = float(row[post_col])
                        record["_difference"] = post_v - pre_v
                    except (TypeError, ValueError):
                        record["_difference"] = None
                if verdict is True:
                    result[col]["true_rows"].append(record)
                    result[col]["true_groups"].setdefault(record["_mismatchtype"], []).append(record)
                elif verdict is False:
                    result[col]["false_rows"].append(record)
                else:
                    result[col]["unclassified"] += 1
    return result


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


def build_data_digest(tables: list[dict], triage: dict[str, dict], max_sample_rows: int) -> dict:
    """A compact, token-budget-aware summary of the ingested data + the
    already-computed triage stats, meant to be embedded in the Claude
    prompt. Every number in here is exact (computed over the full dataset);
    only the row-level samples are capped."""
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

    return {
        "files": files,
        "process_types_by_file": process_types_by_file,
        "sheets": sheets,
        "issues": issues,
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
