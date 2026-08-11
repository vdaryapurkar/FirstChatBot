"""Shared structured-output schema for the triage analysis call.

Both provider clients (core/claude_client.py for Claude, core/openai_client.py
for any OpenAI-compatible server -- Ollama, vLLM, LM Studio, OpenRouter,
Together, Groq, or plain OpenAI) submit this exact schema as a forced tool/
function call, so switching providers never changes what comes back. Each
wraps ANALYSIS_PARAMETERS in its own tool-definition envelope (Anthropic's
"input_schema" vs. OpenAI's "parameters"), but the schema itself -- and the
mismatchtype / process_type vocabulary in it -- lives here once.
"""

from __future__ import annotations

TOOL_NAME = "submit_triage_analysis"
TOOL_DESCRIPTION = "Submit the completed triage analysis for the uploaded reconciliation data."

PROCESS_TYPES = ["Settlement", "Valuation", "NetValuation", "Credit", "Unknown"]
MISMATCHTYPES = ["mismatch", "new_post", "missing_position_post", "other"]
CATEGORIES = ["BUY/SELL", "Other"]

ANALYSIS_PARAMETERS = {
    "type": "object",
    "properties": {
        "summary": {
            "type": "string",
            "description": "3-6 sentence executive summary across all uploaded files.",
        },
        "triage_categories": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "column": {
                        "type": "string",
                        "description": (
                            "The single mismatch column this issue is specific to. Set this for "
                            "'mismatch' (and 'other') mismatchtype issues, which stay one issue "
                            "per column. Leave unset for 'new_post'/'missing_position_post' "
                            "issues -- those are pooled across every column they triggered, so "
                            "set 'category' instead, not 'column'."
                        ),
                    },
                    "mismatchtype": {
                        "type": "string",
                        "enum": MISMATCHTYPES,
                        "description": (
                            "Which row-level mismatchtype this issue covers -- 'mismatch' "
                            "(genuine value difference for a position present pre and post), "
                            "'new_post' (position newly added in post, absent from pre), "
                            "'missing_position_post' (position present in pre but absent from "
                            "post), or 'other' for any value not in that set."
                        ),
                    },
                    "category": {
                        "type": "string",
                        "enum": CATEGORIES,
                        "description": (
                            "Only for 'new_post'/'missing_position_post' issues: the position-"
                            "type category this issue covers -- 'BUY/SELL' if positiontype is "
                            "BUY or SELL, else 'Other'. Each (mismatchtype, category) pair is "
                            "exactly one issue, already pooled across every column it triggered "
                            "in Python -- do not also set 'column' for these, and do not split or "
                            "merge these pairs further. Leave unset for 'mismatch'/'other' issues."
                        ),
                    },
                    "description": {
                        "type": "string",
                        "description": (
                            "For 'new_post'/'missing_position_post' issues, lead with the "
                            "category and what it means structurally (e.g. a new BUY/SELL "
                            "position added in post) rather than restating pre/post values -- "
                            "those are 0 on one side by construction, not informative on their "
                            "own. Only describe a value-based pattern if one actually exists "
                            "across the pooled rows (e.g. all new positions share a market or a "
                            "size range). For 'mismatch' issues, values remain the primary "
                            "evidence as usual."
                        ),
                    },
                    "count": {"type": "integer"},
                    "process_types": {
                        "type": "array",
                        "items": {"type": "string", "enum": PROCESS_TYPES},
                        "description": "Which process type(s) this issue was observed in.",
                    },
                    "correlated_with": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Other mismatch columns whose True/False status always matches this one.",
                    },
                },
                "required": ["mismatchtype", "description"],
            },
        },
        "root_causes": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "issue": {"type": "string"},
                    "column": {
                        "type": "string",
                        "description": "Set for 'mismatch'/'other' issues. Leave unset for 'new_post'/'missing_position_post' -- use 'category' instead.",
                    },
                    "mismatchtype": {"type": "string", "enum": MISMATCHTYPES},
                    "category": {
                        "type": "string",
                        "enum": CATEGORIES,
                        "description": "Only for 'new_post'/'missing_position_post' root causes: which position-type category ('BUY/SELL' or 'Other') this root cause explains.",
                    },
                    "process_type": {"type": "string", "enum": PROCESS_TYPES},
                    "explanation": {"type": "string"},
                    "evidence": {"type": "array", "items": {"type": "string"}},
                    "affected_scope": {"type": "string"},
                    "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
                },
                "required": ["issue", "explanation", "confidence"],
            },
        },
        "recommendations": {"type": "array", "items": {"type": "string"}},
        "bug_report": {
            "type": "object",
            "description": (
                "A synopsis and description written for a bug tracker (e.g. TFS/Azure DevOps), "
                "ready to paste as-is into a new bug's title and description fields to file one "
                "ticket covering everything found this run. Always populate both fields -- even "
                "for a single issue -- but they matter most when there are multiple distinct "
                "(column, mismatchtype) issues to investigate together."
            ),
            "properties": {
                "synopsis": {
                    "type": "string",
                    "description": "One-line bug title summarizing the issue(s) found.",
                },
                "description": {
                    "type": "string",
                    "description": (
                        "Plain-text bug body covering every issue found this run: for each, name "
                        "the column, mismatchtype, process type(s), affected row count, root "
                        "cause, and evidence. Use plain line breaks and '- ' bullets -- no "
                        "markdown syntax (#, **, _, etc.), since this is pasted directly into a "
                        "bug tracker's plain/rich-text description field."
                    ),
                },
            },
            "required": ["synopsis", "description"],
        },
    },
    "required": ["summary", "triage_categories", "root_causes", "recommendations", "bug_report"],
}


def history_to_messages(history: list[dict]) -> list[dict]:
    messages = []
    for m in history:
        if m["role"] not in ("user", "assistant"):
            continue
        messages.append({"role": m["role"], "content": m["content"]})
    return messages
