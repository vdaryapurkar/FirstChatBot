"""SQLite persistence for triage sessions.

Stores session metadata, chat-like message history (the running "context"
for a session), uploaded-file records, and generated report records.

The Claude API key is deliberately NOT stored here -- see core/keys.py. This
database only holds session context, so restarting the app keeps your
session history but always requires re-entering the API key.
"""

import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent.parent / "data" / "triage.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS sessions (
    id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    carried_from_session_id TEXT
);

CREATE TABLE IF NOT EXISTS messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL,
    role TEXT NOT NULL,
    content TEXT NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY (session_id) REFERENCES sessions(id)
);

CREATE TABLE IF NOT EXISTS uploads (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL,
    original_name TEXT NOT NULL,
    stored_path TEXT NOT NULL,
    sheet_count INTEGER NOT NULL,
    row_count INTEGER NOT NULL,
    uploaded_at TEXT NOT NULL,
    FOREIGN KEY (session_id) REFERENCES sessions(id)
);

CREATE TABLE IF NOT EXISTS reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL,
    stored_path TEXT NOT NULL,
    summary TEXT NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY (session_id) REFERENCES sessions(id)
);
"""


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


@contextmanager
def get_conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    with get_conn() as conn:
        conn.executescript(SCHEMA)


def create_session(title: str, carried_from_session_id: str | None = None) -> str:
    session_id = uuid.uuid4().hex[:12]
    now = _now()
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO sessions (id, title, created_at, updated_at, carried_from_session_id) "
            "VALUES (?, ?, ?, ?, ?)",
            (session_id, title, now, now, carried_from_session_id),
        )
    return session_id


def touch_session(session_id: str):
    with get_conn() as conn:
        conn.execute("UPDATE sessions SET updated_at = ? WHERE id = ?", (_now(), session_id))


def list_sessions() -> list[dict]:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT s.id, s.title, s.created_at, s.updated_at, s.carried_from_session_id, "
            "  (SELECT COUNT(*) FROM messages m WHERE m.session_id = s.id) AS message_count, "
            "  (SELECT COUNT(*) FROM uploads u WHERE u.session_id = s.id) AS upload_count, "
            "  (SELECT COUNT(*) FROM reports r WHERE r.session_id = s.id) AS report_count "
            "FROM sessions s ORDER BY s.updated_at DESC"
        ).fetchall()
        return [dict(r) for r in rows]


def get_session(session_id: str) -> dict | None:
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM sessions WHERE id = ?", (session_id,)).fetchone()
        return dict(row) if row else None


def add_message(session_id: str, role: str, content: str):
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO messages (session_id, role, content, created_at) VALUES (?, ?, ?, ?)",
            (session_id, role, content, _now()),
        )
        conn.execute("UPDATE sessions SET updated_at = ? WHERE id = ?", (_now(), session_id))


def get_messages(session_id: str) -> list[dict]:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT role, content, created_at FROM messages WHERE session_id = ? ORDER BY id ASC",
            (session_id,),
        ).fetchall()
        return [dict(r) for r in rows]


def add_upload(session_id: str, original_name: str, stored_path: str, sheet_count: int, row_count: int) -> int:
    with get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO uploads (session_id, original_name, stored_path, sheet_count, row_count, uploaded_at) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (session_id, original_name, stored_path, sheet_count, row_count, _now()),
        )
        conn.execute("UPDATE sessions SET updated_at = ? WHERE id = ?", (_now(), session_id))
        return cur.lastrowid


def get_uploads(session_id: str) -> list[dict]:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM uploads WHERE session_id = ? ORDER BY id ASC", (session_id,)
        ).fetchall()
        return [dict(r) for r in rows]


def add_report(session_id: str, stored_path: str, summary: str) -> int:
    with get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO reports (session_id, stored_path, summary, created_at) VALUES (?, ?, ?, ?)",
            (session_id, stored_path, summary, _now()),
        )
        conn.execute("UPDATE sessions SET updated_at = ? WHERE id = ?", (_now(), session_id))
        return cur.lastrowid


def get_reports(session_id: str) -> list[dict]:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM reports WHERE session_id = ? ORDER BY id DESC", (session_id,)
        ).fetchall()
        return [dict(r) for r in rows]


def get_report(report_id: int) -> dict | None:
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM reports WHERE id = ?", (report_id,)).fetchone()
        return dict(row) if row else None


def last_assistant_summary(session_id: str) -> str | None:
    """Latest assistant analysis summary for a session, used to carry context into a new session."""
    with get_conn() as conn:
        row = conn.execute(
            "SELECT content FROM messages WHERE session_id = ? AND role = 'assistant' "
            "ORDER BY id DESC LIMIT 1",
            (session_id,),
        ).fetchone()
        return row["content"] if row else None
