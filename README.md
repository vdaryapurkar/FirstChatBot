# Recon-Ci Triage

A local web interface that uses the Claude API to analyze and triage
reconciliation/"unexplained result" spreadsheets, and produces a downloadable
`.xlsx` report with the comparison data, a triage summary, and Claude's
root-cause analysis.

## What it does

- You enter your own Claude API key in the UI (kept in server memory for
  your browser session only -- see [Security notes](#security-notes)).
- Upload one or more `.xlsx` files, or an entire folder of them.
- All row-level comparison math (True/False triage per mismatch-flag column,
  numeric pre/post differences) is computed deterministically in Python over
  the full dataset -- Claude is never asked to do arithmetic, only to
  interpret it and explain root causes, following the rules you configure in
  [`config/rules.py`](config/rules.py).
- Each browser session's uploads, analysis conversation, and generated
  reports are kept as a **session** in a local SQLite database, like a
  Claude chat: you can revisit old sessions, and a new session can carry
  forward the prior session's summary as context (see the "Carry context
  from a previous session" option when creating a session).
- Every analysis run produces a downloadable `.xlsx` with a `Summary` sheet,
  a `Root Cause Analysis` sheet, and one detail sheet per mismatch-flag
  column found in your data (rows grouped TRUE/FALSE, with a live
  `=post-pre` difference formula).

## Setup

```bash
pip install -r requirements.txt
python app.py
```

Then open http://localhost:5000, paste your Claude API key into the sidebar,
and click "Save key".

## Configuring the triage rules

Edit [`config/rules.py`](config/rules.py):

- `ANALYSIS_SYSTEM_PROMPT` is the "Claude context" -- the system prompt sent
  on every analysis call. It ships with the reconciliation triage rules
  worked out earlier in this project (group by `*_mismatch` columns, look
  for constant vs. proportional offsets, etc.) as a working example. Replace
  it with your own domain rules.
- `MODEL_NAME` -- which Claude model to call.
- `MAX_SAMPLE_ROWS_PER_SHEET` -- caps how many example rows per sheet are
  included in the prompt sent to Claude (token budget guard). This never
  affects the output report, which always includes every row -- it only
  limits how many raw rows Claude sees as supporting evidence.

## How data flows

1. **Upload** -- `.xlsx` files (individually or via a folder picker /
   drag-and-drop of a folder) are parsed with `core/xlsx_ingest.py`. Every
   column ending in `_mismatch` is treated as a triage flag; the module
   tries to find its paired `pre_*`/`post_*` columns automatically.
2. **Triage** -- `xlsx_ingest.compute_triage()` groups every row into
   TRUE/FALSE buckets per flag column, over the complete dataset.
3. **Digest** -- `xlsx_ingest.build_data_digest()` builds a compact,
   token-budget-aware JSON summary (exact counts/stats + a capped row
   sample) to send to Claude.
4. **Analyze** -- `core/claude_client.py` sends the system prompt from
   `config/rules.py`, the session's prior conversation (for continuity), and
   the digest to Claude, and gets back a structured result via a forced tool
   call (`submit_triage_analysis`) -- this is what makes the JSON response
   reliable rather than asking Claude to free-text a JSON blob.
5. **Report** -- `core/report_builder.py` merges the full (uncapped) triage
   tables with Claude's narrative into the downloadable workbook.

## Security notes

- Your Claude API key is **never written to disk or the database**. It's
  held in an in-memory dict in the Flask process, keyed by a signed session
  cookie (`core/keys.py`). Restarting the server clears every key; each
  browser session re-enters its key once per server run.
- Session **context** (conversation history, upload metadata, generated
  report paths) *is* persisted to `data/triage.db` so sessions survive
  restarts and can be revisited like chat history -- it contains no
  secrets, only your uploaded data and Claude's analysis text.
- This is built for local/trusted-network use. There's no authentication
  beyond the browser session cookie, uploaded files and reports are stored
  unencrypted under `data/`, and the dev server (`app.run(debug=True)`)
  should not be exposed to the internet as-is.

## Project layout

```
app.py                   Flask routes
config/rules.py          Claude system prompt / triage rules (edit this)
core/db.py                Session/message/upload/report persistence (SQLite)
core/keys.py               In-memory, per-browser-session API key store
core/xlsx_ingest.py         Multi-file/folder ingestion + deterministic triage math
core/claude_client.py        Anthropic SDK wrapper (forced tool-call for structured output)
core/report_builder.py        Builds the downloadable .xlsx report
templates/index.html, static/  Frontend (vanilla JS, no build step)
```
