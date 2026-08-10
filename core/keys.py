"""In-memory Claude API key + model choice, keyed by browser (Flask) session id.

The key is deliberately kept in server process memory only -- never written
to the SQLite database, a file, or logs. Restarting the app clears every key
and each browser session must re-enter it. This trades convenience for not
persisting a secret to disk inside a repo-adjacent app. The model choice
carries no secrecy requirement but is stored the same way for simplicity --
it also resets on restart, which is fine since it's just a UI preference.
"""

import threading

from config.rules import MODEL_NAME

_lock = threading.Lock()
_keys: dict[str, str] = {}
_models: dict[str, str] = {}


def set_key(browser_session_id: str, api_key: str):
    with _lock:
        _keys[browser_session_id] = api_key


def get_key(browser_session_id: str) -> str | None:
    with _lock:
        return _keys.get(browser_session_id)


def clear_key(browser_session_id: str):
    with _lock:
        _keys.pop(browser_session_id, None)


def has_key(browser_session_id: str) -> bool:
    with _lock:
        return browser_session_id in _keys


def set_model(browser_session_id: str, model: str):
    with _lock:
        _models[browser_session_id] = model


def get_model(browser_session_id: str) -> str:
    with _lock:
        return _models.get(browser_session_id, MODEL_NAME)
