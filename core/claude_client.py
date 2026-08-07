"""Thin wrapper around the Anthropic SDK for the triage analysis call.

Structured output is obtained via a forced tool call (submit_triage_analysis)
rather than asking Claude to emit raw JSON in text -- this is the reliable
way to get a parseable result back from the API.
"""

from __future__ import annotations

import json

import anthropic

from config.rules import ANALYSIS_SYSTEM_PROMPT, MODEL_NAME

TOOL_NAME = "submit_triage_analysis"

ANALYSIS_TOOL = {
    "name": TOOL_NAME,
    "description": "Submit the completed triage analysis for the uploaded reconciliation data.",
    "input_schema": {
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
                        "description": {"type": "string"},
                        "true_count": {"type": "integer"},
                        "false_count": {"type": "integer"},
                        "correlated_with": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Other mismatch columns whose True/False status always matches this one.",
                        },
                    },
                    "required": ["column", "description"],
                },
            },
            "root_causes": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "issue": {"type": "string"},
                        "explanation": {"type": "string"},
                        "evidence": {"type": "array", "items": {"type": "string"}},
                        "affected_scope": {"type": "string"},
                        "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
                    },
                    "required": ["issue", "explanation", "confidence"],
                },
            },
            "recommendations": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["summary", "triage_categories", "root_causes", "recommendations"],
    },
}


class ClaudeAnalysisError(RuntimeError):
    pass


def _history_to_messages(history: list[dict]) -> list[dict]:
    messages = []
    for m in history:
        if m["role"] not in ("user", "assistant"):
            continue
        messages.append({"role": m["role"], "content": m["content"]})
    return messages


def run_analysis(
    api_key: str,
    conversation_history: list[dict],
    digest: dict,
    extra_instructions: str | None = None,
) -> dict:
    """conversation_history: prior messages for this session (role/content),
    already stored in order. digest: output of xlsx_ingest.build_data_digest.
    Returns the parsed tool-call input dict."""

    if not api_key or not api_key.strip():
        raise ClaudeAnalysisError("No Claude API key set for this browser session.")

    user_parts = [
        "Here is the newly uploaded data and its computed comparison/triage statistics "
        "(as JSON). Analyze it per the rules in your system prompt.",
        "```json",
        json.dumps(digest, indent=2, default=str),
        "```",
    ]
    if extra_instructions:
        user_parts.append(f"\nAdditional instructions for this analysis: {extra_instructions}")
    user_message = {"role": "user", "content": "\n".join(user_parts)}

    messages = _history_to_messages(conversation_history) + [user_message]

    client = anthropic.Anthropic(api_key=api_key)
    try:
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=4096,
            system=ANALYSIS_SYSTEM_PROMPT,
            messages=messages,
            tools=[ANALYSIS_TOOL],
            tool_choice={"type": "tool", "name": TOOL_NAME},
        )
    except anthropic.AuthenticationError as e:
        raise ClaudeAnalysisError("Claude rejected the API key (authentication error).") from e
    except anthropic.APIConnectionError as e:
        raise ClaudeAnalysisError(f"Could not reach the Claude API: {e}") from e
    except anthropic.APIStatusError as e:
        raise ClaudeAnalysisError(f"Claude API error ({e.status_code}): {e.message}") from e
    except Exception as e:  # noqa: BLE001 - surface any other SDK error to the UI
        raise ClaudeAnalysisError(f"Claude API call failed: {e}") from e

    for block in response.content:
        if block.type == "tool_use" and block.name == TOOL_NAME:
            return {
                "result": block.input,
                "user_message_content": user_message["content"],
            }

    raise ClaudeAnalysisError("Claude did not return a structured analysis (no tool call in response).")
