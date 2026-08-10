"""Thin wrapper around the Anthropic SDK for the triage analysis call.

Structured output is obtained via a forced tool call (submit_triage_analysis)
rather than asking Claude to emit raw JSON in text -- this is the reliable
way to get a parseable result back from the API. The tool schema itself is
shared with core/openai_client.py -- see core/llm_schema.py.
"""

from __future__ import annotations

import json

import anthropic

from config.rules import ANALYSIS_SYSTEM_PROMPT, MODEL_NAME
from core.llm_errors import LLMAnalysisError as ClaudeAnalysisError
from core.llm_schema import ANALYSIS_PARAMETERS, TOOL_DESCRIPTION, TOOL_NAME, history_to_messages

ANALYSIS_TOOL = {
    "name": TOOL_NAME,
    "description": TOOL_DESCRIPTION,
    "input_schema": ANALYSIS_PARAMETERS,
}


def run_analysis(
    api_key: str,
    conversation_history: list[dict],
    digest: dict,
    extra_instructions: str | None = None,
    model: str | None = None,
) -> dict:
    """conversation_history: prior messages for this session (role/content),
    already stored in order. digest: output of xlsx_ingest.build_data_digest.
    model: Claude model ID to use for this call; defaults to config.rules.MODEL_NAME.
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

    messages = history_to_messages(conversation_history) + [user_message]

    client = anthropic.Anthropic(api_key=api_key)
    try:
        response = client.messages.create(
            model=model or MODEL_NAME,
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
