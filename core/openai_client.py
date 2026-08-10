"""OpenAI-compatible chat-completions client for the triage analysis call.

Works with any server implementing the OpenAI /v1/chat/completions API with
tool calling: Ollama (http://localhost:11434/v1), vLLM's OpenAI-compatible
server, LM Studio, text-generation-webui, OpenRouter, Together, Groq, or the
real OpenAI API. This is the path for testing with an open-source model --
point base_url at your server, put in whatever model name it exposes, and an
API key only if the server requires one (local servers usually don't).
Switching back to Claude is just picking "Claude" as the provider again in
the UI; nothing else about the app changes, since both clients submit the
same schema (core/llm_schema.py) as a forced tool call and return the same
shape.
"""

from __future__ import annotations

import json

import openai

from config.rules import ANALYSIS_SYSTEM_PROMPT
from core.llm_errors import LLMAnalysisError
from core.llm_schema import ANALYSIS_PARAMETERS, TOOL_DESCRIPTION, TOOL_NAME, history_to_messages

ANALYSIS_TOOL = {
    "type": "function",
    "function": {
        "name": TOOL_NAME,
        "description": TOOL_DESCRIPTION,
        "parameters": ANALYSIS_PARAMETERS,
    },
}


def run_analysis(
    api_key: str | None,
    base_url: str,
    model: str,
    conversation_history: list[dict],
    digest: dict,
    extra_instructions: str | None = None,
) -> dict:
    """api_key: may be blank/None -- many local model servers don't require one.
    base_url: the server's OpenAI-compatible API root, e.g. http://localhost:11434/v1.
    model: the exact model name the server exposes (e.g. "llama3.1:70b").
    Returns the same {"result": ..., "user_message_content": ...} shape as
    core.claude_client.run_analysis."""

    if not base_url or not base_url.strip():
        raise LLMAnalysisError("No server URL set for the open-source/OpenAI-compatible provider.")
    if not model or not model.strip():
        raise LLMAnalysisError("No model name set for the open-source/OpenAI-compatible provider.")

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

    messages = [{"role": "system", "content": ANALYSIS_SYSTEM_PROMPT}]
    messages += history_to_messages(conversation_history)
    messages.append(user_message)

    client = openai.OpenAI(api_key=(api_key or "not-needed"), base_url=base_url)
    try:
        response = client.chat.completions.create(
            model=model,
            messages=messages,
            tools=[ANALYSIS_TOOL],
            tool_choice={"type": "function", "function": {"name": TOOL_NAME}},
        )
    except openai.AuthenticationError as e:
        raise LLMAnalysisError("The model server rejected the API key (authentication error).") from e
    except openai.APIConnectionError as e:
        raise LLMAnalysisError(f"Could not reach the model server at {base_url}: {e}") from e
    except openai.APIStatusError as e:
        raise LLMAnalysisError(f"Model server error ({e.status_code}): {e.message}") from e
    except Exception as e:  # noqa: BLE001 - surface any other SDK/server error to the UI
        raise LLMAnalysisError(f"Model call failed: {e}") from e

    message = response.choices[0].message
    for tool_call in (message.tool_calls or []):
        if tool_call.function.name == TOOL_NAME:
            try:
                parsed = json.loads(tool_call.function.arguments)
            except json.JSONDecodeError as e:
                raise LLMAnalysisError(f"Model returned malformed JSON for its tool call: {e}") from e
            return {"result": parsed, "user_message_content": user_message["content"]}

    raise LLMAnalysisError(
        "The model did not return a structured tool call. Some open-source models or servers "
        "need tool calling explicitly enabled, or don't support a forced tool_choice -- check "
        "your server's docs (e.g. Ollama needs a model built/tagged with tool-calling support)."
    )
