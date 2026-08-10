"""Recon-Ci Triage: reconciliation triage interface powered by Claude or any
OpenAI-compatible model server (Ollama, vLLM, LM Studio, OpenRouter, ...).

Run with:
    pip install -r requirements.txt
    python app.py

Then open http://localhost:5000 and pick a provider in the sidebar: Claude
(enter an API key + pick a model) or an open-source/OpenAI-compatible server
(enter its base URL + exact model name, and an API key only if the server
requires one). The connection settings are kept in server memory for your
browser session only -- see core/keys.py -- and never written to disk or the
database. Switch providers any time; session context carries over either way.

Edit config/rules.py to change the triage rules / system prompt sent to
whichever model is active.
"""

from __future__ import annotations

import json
import os
import uuid
from pathlib import Path

from flask import Flask, jsonify, request, send_file, session, render_template

from core import claude_client, db, keys, openai_client, xlsx_ingest, report_builder
from core.llm_errors import LLMAnalysisError
from config.rules import MAX_SAMPLE_ROWS_PER_SHEET, AVAILABLE_MODELS

AVAILABLE_MODEL_IDS = {m["id"] for m in AVAILABLE_MODELS}

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "data" / "uploads"
REPORT_DIR = BASE_DIR / "data" / "reports"
ALLOWED_EXT = {".xlsx"}

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", os.urandom(32))
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200MB total upload cap


def _browser_id() -> str:
    if "browser_id" not in session:
        session["browser_id"] = uuid.uuid4().hex
    return session["browser_id"]


@app.route("/")
def index():
    return render_template("index.html")


# ---------------------------------------------------------------- API key --

@app.route("/api/key", methods=["POST"])
def set_key():
    data = request.get_json(silent=True) or {}
    if not data:
        return jsonify({"error": "no fields provided"}), 400

    browser_id = _browser_id()
    current = keys.get_settings(browser_id)
    fields = {}

    if "provider" in data:
        provider = (data.get("provider") or "").strip()
        if provider not in keys.PROVIDERS:
            return jsonify({"error": f"Unknown provider '{provider}'. Choose one of: {list(keys.PROVIDERS)}"}), 400
        fields["provider"] = provider

    if "api_key" in data:
        fields["api_key"] = (data.get("api_key") or "").strip() or None

    if "model" in data:
        model = (data.get("model") or "").strip()
        effective_provider = fields.get("provider", current["provider"])
        if effective_provider == keys.PROVIDER_ANTHROPIC and model not in AVAILABLE_MODEL_IDS:
            return jsonify({"error": f"Unknown Claude model '{model}'. Choose one of: {sorted(AVAILABLE_MODEL_IDS)}"}), 400
        fields["model"] = model

    if "base_url" in data:
        fields["base_url"] = (data.get("base_url") or "").strip()

    keys.update_settings(browser_id, fields)
    return jsonify({"status": "ok"})


@app.route("/api/key/status")
def key_status():
    s = keys.get_settings(_browser_id())
    return jsonify({
        "has_key": bool(s["api_key"]),
        "provider": s["provider"],
        "model": s["model"],
        "base_url": s["base_url"],
        "available_models": AVAILABLE_MODELS,
        "providers": [
            {"id": keys.PROVIDER_ANTHROPIC, "label": "Claude (Anthropic)"},
            {"id": keys.PROVIDER_OPENAI_COMPATIBLE, "label": "Open-source / OpenAI-compatible (Ollama, vLLM, LM Studio, OpenRouter, ...)"},
        ],
    })


@app.route("/api/key", methods=["DELETE"])
def clear_key():
    keys.clear_key(_browser_id())
    return jsonify({"status": "ok"})


# ---------------------------------------------------------------- sessions --

@app.route("/api/sessions", methods=["GET"])
def list_sessions():
    return jsonify(db.list_sessions())


@app.route("/api/sessions", methods=["POST"])
def create_session():
    data = request.get_json(silent=True) or {}
    title = (data.get("title") or "").strip() or "Untitled session"
    carry_from = data.get("carry_from_session_id")

    carried_summary = None
    if carry_from:
        carried_summary = db.last_assistant_summary(carry_from)
        if carried_summary is None:
            return jsonify({"error": f"No prior analysis found on session {carry_from} to carry over"}), 400

    session_id = db.create_session(title, carried_from_session_id=carry_from)

    if carried_summary:
        carried = db.get_session(carry_from)
        prior_title = carried["title"] if carried else carry_from
        db.add_message(
            session_id, "user",
            f"[Carried-over context from previous session '{prior_title}' ({carry_from})]\n{carried_summary}",
        )
        db.add_message(
            session_id, "assistant",
            "Acknowledged - I will treat that prior analysis as established context for this session.",
        )

    return jsonify(db.get_session(session_id)), 201


@app.route("/api/sessions/<session_id>")
def get_session(session_id):
    s = db.get_session(session_id)
    if not s:
        return jsonify({"error": "session not found"}), 404
    return jsonify({
        "session": s,
        "messages": db.get_messages(session_id),
        "uploads": db.get_uploads(session_id),
        "reports": db.get_reports(session_id),
    })


# ------------------------------------------------------------------ upload --

@app.route("/api/sessions/<session_id>/upload", methods=["POST"])
def upload_files(session_id):
    if not db.get_session(session_id):
        return jsonify({"error": "session not found"}), 404

    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "no files provided"}), 400

    session_upload_dir = UPLOAD_DIR / session_id
    session_upload_dir.mkdir(parents=True, exist_ok=True)

    saved, skipped = [], []
    for f in files:
        original_name = f.filename or "unnamed"
        ext = Path(original_name).suffix.lower()
        if ext not in ALLOWED_EXT:
            skipped.append(original_name)
            continue
        stored_name = f"{uuid.uuid4().hex[:8]}_{Path(original_name).name}"
        stored_path = session_upload_dir / stored_name
        f.save(stored_path)

        try:
            sheets = xlsx_ingest.read_workbook(stored_path)
            sheet_count = len(sheets)
            row_count = sum(len(df) for df in sheets.values())
        except Exception as e:  # noqa: BLE001
            stored_path.unlink(missing_ok=True)
            skipped.append(f"{original_name} (could not be read as .xlsx: {e})")
            continue

        upload_id = db.add_upload(session_id, original_name, str(stored_path), sheet_count, row_count)
        saved.append({
            "id": upload_id, "original_name": original_name,
            "sheet_count": sheet_count, "row_count": row_count,
        })

    return jsonify({"saved": saved, "skipped": skipped})


# ----------------------------------------------------------------- analyze --

@app.route("/api/sessions/<session_id>/analyze", methods=["POST"])
def analyze(session_id):
    if not db.get_session(session_id):
        return jsonify({"error": "session not found"}), 404

    settings = keys.get_settings(_browser_id())
    provider = settings["provider"]
    if provider == keys.PROVIDER_ANTHROPIC and not settings["api_key"]:
        return jsonify({"error": "No Claude API key set. Enter one in the sidebar first."}), 401
    if provider == keys.PROVIDER_OPENAI_COMPATIBLE and not settings["base_url"]:
        return jsonify({"error": "No server URL set. Enter one in the sidebar (e.g. http://localhost:11434/v1 for Ollama)."}), 400
    if provider == keys.PROVIDER_OPENAI_COMPATIBLE and not settings["model"]:
        return jsonify({"error": "No model name set. Enter the exact model name your server exposes (e.g. llama3.1:70b)."}), 400

    data = request.get_json(silent=True) or {}
    extra_instructions = (data.get("extra_instructions") or "").strip() or None

    uploads = db.get_uploads(session_id)
    if not uploads:
        return jsonify({"error": "No files uploaded to this session yet."}), 400

    file_paths = [(u["original_name"], u["stored_path"]) for u in uploads]
    try:
        tables = xlsx_ingest.ingest_files(file_paths)
    except Exception as e:  # noqa: BLE001
        return jsonify({"error": f"Failed to read uploaded files: {e}"}), 400

    if not tables:
        return jsonify({"error": "Uploaded files contained no readable data."}), 400

    triage = xlsx_ingest.compute_triage(tables)
    digest = xlsx_ingest.json_safe(
        xlsx_ingest.build_data_digest(tables, triage, MAX_SAMPLE_ROWS_PER_SHEET)
    )

    history = db.get_messages(session_id)
    try:
        if provider == keys.PROVIDER_ANTHROPIC:
            outcome = claude_client.run_analysis(
                settings["api_key"], history, digest, extra_instructions, model=settings["model"]
            )
        else:
            outcome = openai_client.run_analysis(
                settings["api_key"], settings["base_url"], settings["model"],
                history, digest, extra_instructions,
            )
    except LLMAnalysisError as e:
        return jsonify({"error": str(e)}), 502

    result = outcome["result"]
    db.add_message(session_id, "user", outcome["user_message_content"])
    db.add_message(session_id, "assistant", json.dumps(result))

    session_report_dir = REPORT_DIR / session_id
    session_report_dir.mkdir(parents=True, exist_ok=True)
    report_filename = f"triage_report_{uuid.uuid4().hex[:8]}.xlsx"
    report_path = session_report_dir / report_filename

    sources = sorted({u["original_name"] for u in uploads})
    report_builder.build_report(triage, result, sources, str(report_path),
                                 process_types_by_file=digest.get("process_types_by_file"))

    report_id = db.add_report(session_id, str(report_path), result.get("summary", ""))

    return jsonify({"analysis": result, "report_id": report_id})


@app.route("/api/sessions/<session_id>/reports/<int:report_id>/download")
def download_report(session_id, report_id):
    report = db.get_report(report_id)
    if not report or report["session_id"] != session_id:
        return jsonify({"error": "report not found"}), 404
    return send_file(report["stored_path"], as_attachment=True,
                      download_name=Path(report["stored_path"]).name)


if __name__ == "__main__":
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    db.init_db()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
