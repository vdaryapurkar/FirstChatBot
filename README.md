# Recon-Ci Triage

A local web interface that uses an LLM -- Claude, or any OpenAI-compatible
model server (Ollama, vLLM, LM Studio, OpenRouter, Together, Groq, ...) -- to
analyze and triage reconciliation/"unexplained result" spreadsheets, and
produces a downloadable `.xlsx` report with the comparison data, a triage
summary, and the model's root-cause analysis.

## What it does

- Pick a **provider** in the sidebar: Claude (enter an API key, pick a model
  from the dropdown) or an open-source/OpenAI-compatible server (enter its
  base URL and the exact model name it exposes, plus an API key only if the
  server requires one). Everything is kept in server memory for your browser
  session only -- see [Security notes](#security-notes). This is the
  intended path for testing against a local open-source model before paying
  for Claude calls -- switch back to Claude any time by picking it again;
  session context carries over either way, since both providers submit the
  exact same schema and prompt (`core/llm_schema.py`).
- Upload one or more `.xlsx` files, or an entire folder of them.
- All row-level comparison math (True/False triage per mismatch-flag column,
  numeric pre/post differences) is computed deterministically in Python over
  the full dataset -- the model is never asked to do arithmetic, only to
  interpret it and explain root causes, following the rules you configure in
  [`config/rules.py`](config/rules.py).
- Rows are also split by **mismatchtype** (from a `mismatchtype` column, if
  present): `mismatch` (a genuine value difference for a position that
  exists pre and post), `new_post` (a position newly added in post, absent
  from pre), or `missing_position_post` / `missing_post` (a position present
  in pre but dropped from post). Each `(mismatch column, mismatchtype)`
  combination -- e.g. `qty_mismatch` / `new_post` vs. `qty_mismatch` /
  `mismatch` -- is treated as its own distinct issue with its own row group,
  its own diff-shape stats, and its own root-cause explanation, since a new
  or dropped position has a different cause than a genuine value diff.
- Each browser session's uploads, analysis conversation, and generated
  reports are kept as a **session** in a local SQLite database, like a
  Claude chat: you can revisit old sessions, and a new session can carry
  forward the prior session's summary as context (see the "Carry context
  from a previous session" option when creating a session).
- Every analysis run produces a downloadable `.xlsx` with a `Summary` sheet
  (one row per issue: column, mismatch type, count, process type(s),
  description), a `Root Cause Analysis` sheet, and one detail sheet per
  mismatch-flag column found in your data (rows grouped by mismatch type,
  each visually distinct, then FALSE/no-issue, with a live `=post-pre`
  difference formula).

## Setup

```bash
pip install -r requirements.txt
python app.py
```

Then open http://localhost:5000 and configure a provider in the sidebar:
- **Claude**: paste your API key, pick a model, click "Save".
- **Open-source/OpenAI-compatible**: pick the provider from the dropdown,
  enter your server's base URL (e.g. `http://localhost:11434/v1` for
  [Ollama](https://ollama.com)) and the exact model name it serves (e.g.
  `llama3.1:70b`), leave the API key blank unless your server requires one,
  click "Save". The model must support OpenAI-style tool/function calling
  with a forced `tool_choice` for structured output to work -- most modern
  local-serving stacks (recent Ollama, vLLM, LM Studio) do.

### Connecting through an internal AI gateway (e.g. a LiteLLM proxy)

Some organizations front Claude behind an internal gateway that speaks the
OpenAI `/chat/completions` API rather than the native Anthropic API (a
common LiteLLM proxy setup) -- `model` is still a Claude model name like
`claude-haiku-4-5`, but the wire format and client are OpenAI's. That's
exactly what the **Open-source/OpenAI-compatible** provider above talks to,
so no code changes are needed -- just point it at the gateway:

1. `cp .env.example .env`
2. Fill in the three values from your gateway dashboard:
   ```
   AI_API_KEY=<key from the gateway dashboard>
   AI_GATEWAY_ENDPOINT=<gateway base URL, e.g. https://ai-gateway-litellm.your-org.example>
   AI_GATEWAY_MODEL=claude-haiku-4-5
   ```
3. `pip install -r requirements.txt` (pulls in `python-dotenv`, already listed)
4. `python app.py`

Every new browser session now defaults to the Open-source/OpenAI-compatible
provider pre-filled with the gateway's URL/model/key (see `core/keys.py` --
`.env` is gitignored, so the real key never gets committed). The sidebar
still lets you override any of it per session, or switch back to the plain
Claude provider.

## Configuring the triage rules

Edit [`config/rules.py`](config/rules.py):

- `ANALYSIS_SYSTEM_PROMPT` is the "LLM context" -- the system prompt sent on
  every analysis call, to whichever provider is active. It ships with the
  reconciliation triage rules worked out earlier in this project (group by
  `*_mismatch` columns and by `mismatchtype`, look for constant vs.
  proportional offsets, the four process types, etc.) as a working example.
  Replace it with your own domain rules.
- `MODEL_NAME` -- the default Claude model (only used for the Claude
  provider; the open-source provider always uses whatever model name you
  typed in the sidebar).
- `AVAILABLE_MODELS` -- the Claude model presets offered in the sidebar
  dropdown.
- `MAX_SAMPLE_ROWS_PER_SHEET` -- caps how many example rows per issue are
  included in the prompt sent to the model (token budget guard). This never
  affects the output report, which always includes every row -- it only
  limits how many raw rows the model sees as supporting evidence.

## How data flows

1. **Upload** -- `.xlsx` files (individually or via a folder picker /
   drag-and-drop of a folder) are parsed with `core/xlsx_ingest.py`. Every
   column ending in `_mismatch` is treated as a triage flag; the module
   tries to find its paired `pre_*`/`post_*` columns automatically, and
   classifies each file's process type (Valuation/Settlement/NetValuation/
   Credit) from its filename and columns.
2. **Triage** -- `xlsx_ingest.compute_triage()` groups every row into
   TRUE/FALSE buckets per flag column, over the complete dataset, and
   further splits TRUE rows by `mismatchtype` (`mismatch` / `new_post` /
   `missing_position_post`) so each combination is tracked -- and its
   diff-shape computed -- separately.
3. **Digest** -- `xlsx_ingest.build_data_digest()` builds a compact,
   token-budget-aware JSON summary (one entry per `(column, mismatchtype)`
   issue, with exact counts/stats + a capped row sample) to send to the
   model.
4. **Analyze** -- `core/claude_client.py` (Claude) or `core/openai_client.py`
   (any OpenAI-compatible server) sends the system prompt from
   `config/rules.py`, the session's prior conversation (for continuity), and
   the digest to the active provider, and gets back a structured result via
   a forced tool/function call (`submit_triage_analysis`, schema shared in
   `core/llm_schema.py`) -- this is what makes the JSON response reliable
   rather than asking the model to free-text a JSON blob.
5. **Report** -- `core/report_builder.py` merges the full (uncapped) triage
   tables with the model's narrative into the downloadable workbook.

## Security notes

- Your API key (Claude or the open-source server's, if it has one) is
  **never written to disk or the database**. It's held in an in-memory dict
  in the Flask process, keyed by a signed session cookie (`core/keys.py`),
  along with the rest of the connection settings (provider, model, base
  URL). Restarting the server clears everything; each browser session
  re-configures once per server run.
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
app.py                     Flask routes
config/rules.py            LLM system prompt / triage rules (edit this)
core/db.py                  Session/message/upload/report persistence (SQLite)
core/keys.py                 In-memory, per-browser-session provider/key/model/base_url store
core/xlsx_ingest.py           Multi-file/folder ingestion + deterministic triage math + process-type detection
core/llm_schema.py             Shared structured-output schema used by both provider clients
core/llm_errors.py              Shared LLMAnalysisError exception
core/claude_client.py            Anthropic SDK wrapper (forced tool-call for structured output)
core/openai_client.py             OpenAI-compatible client (Ollama/vLLM/LM Studio/OpenRouter/...)
core/report_builder.py             Builds the downloadable .xlsx report
templates/index.html, static/       Frontend (vanilla JS, no build step)
```
