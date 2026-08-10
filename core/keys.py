"""In-memory per-browser-session LLM connection settings: which provider
(Claude or an OpenAI-compatible server), the API key, the model, and (for
non-Claude providers) the server's base URL.

Kept in server process memory only -- never written to the SQLite database,
a file, or logs. Restarting the app clears everything and each browser
session must re-configure (or picks up the .env defaults again -- see
below). This trades convenience for not persisting a secret to disk inside
a repo-adjacent app. The provider/model/base_url carry no secrecy
requirement but are stored the same way for simplicity -- they reset on
restart too, which is fine since they're just a UI preference, unlike the
API key.

Gateway defaults: if AI_API_KEY and AI_GATEWAY_ENDPOINT are both set (e.g.
via a .env file -- see .env.example), every new browser session defaults to
the "openai_compatible" provider pointed at that gateway instead of empty
Claude fields. This is for internal AI gateways (e.g. a LiteLLM proxy) that
front Claude models behind an OpenAI-compatible /chat/completions API --
point base_url at the gateway and it's indistinguishable from any other
OpenAI-compatible server as far as this app is concerned. The UI can still
override any of it per session.
"""

from __future__ import annotations

import os
import threading

from config.rules import MODEL_NAME

PROVIDER_ANTHROPIC = "anthropic"
PROVIDER_OPENAI_COMPATIBLE = "openai_compatible"
PROVIDERS = (PROVIDER_ANTHROPIC, PROVIDER_OPENAI_COMPATIBLE)

_lock = threading.Lock()
_settings: dict[str, dict] = {}


def _defaults() -> dict:
    gateway_key = os.environ.get("AI_API_KEY", "").strip()
    gateway_url = os.environ.get("AI_GATEWAY_ENDPOINT", "").strip()
    if gateway_key and gateway_url:
        return {
            "api_key": gateway_key,
            "provider": PROVIDER_OPENAI_COMPATIBLE,
            "model": os.environ.get("AI_GATEWAY_MODEL", "").strip() or "claude-haiku-4-5",
            "base_url": gateway_url,
        }
    return {
        "api_key": None,
        "provider": PROVIDER_ANTHROPIC,
        "model": MODEL_NAME,
        "base_url": "",
    }


def get_settings(browser_session_id: str) -> dict:
    """Returns a copy of {api_key, provider, model, base_url} for this
    browser session, defaulting to Claude/MODEL_NAME/no key if unset."""
    with _lock:
        return dict(_settings.setdefault(browser_session_id, _defaults()))


def update_settings(browser_session_id: str, fields: dict):
    """fields: any subset of {api_key, provider, model, base_url}. Only keys
    actually present in `fields` are changed -- pass an explicit "" to clear
    a value, omit the key entirely to leave it untouched."""
    with _lock:
        s = _settings.setdefault(browser_session_id, _defaults())
        s.update(fields)


def clear_key(browser_session_id: str):
    """Clears only the API key, leaving provider/model/base_url as-is."""
    with _lock:
        if browser_session_id in _settings:
            _settings[browser_session_id]["api_key"] = None


def has_key(browser_session_id: str) -> bool:
    with _lock:
        return bool(_settings.get(browser_session_id, {}).get("api_key"))
