"""Shared error type raised by any configured LLM provider client
(core/claude_client.py, core/openai_client.py) when it fails to produce a
usable structured analysis -- lets app.py catch one exception type
regardless of which provider is active.
"""


class LLMAnalysisError(RuntimeError):
    pass
