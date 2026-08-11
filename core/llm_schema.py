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
                    "column": {"type": "string"},
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
                    "description": {"type": "string"},
                    "count": {"type": "integer"},
                    "process_types": {
                        "type": "array",
                        "items": {"type": "string", "enum": PROCESS_TYPES},
                        "description": "Which process type(s) this (column, mismatchtype) issue was observed in.",
                    },
                    "correlated_with": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Other mismatch columns whose True/False status always matches this one.",
                    },
                },
                "required": ["column", "mismatchtype", "description"],
            },
        },
        "root_causes": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "issue": {"type": "string"},
                    "column": {"type": "string"},
                    "mismatchtype": {"type": "string", "enum": MISMATCHTYPES},
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
