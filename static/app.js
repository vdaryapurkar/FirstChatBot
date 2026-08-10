const state = {
  sessionId: null,
  sessions: [],
};

async function api(path, opts = {}) {
  const res = await fetch(path, {
    headers: opts.body instanceof FormData ? {} : { "Content-Type": "application/json" },
    ...opts,
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data.error || `Request failed (${res.status})`);
  }
  return data;
}

// ---------------------------------------------------------------- API key --

async function refreshKeyStatus() {
  const { has_key, model, available_models } = await api("/api/key/status");

  const select = document.getElementById("modelSelect");
  if (!select.dataset.populated) {
    select.innerHTML = "";
    for (const m of available_models) {
      const opt = document.createElement("option");
      opt.value = m.id;
      opt.textContent = m.label;
      select.appendChild(opt);
    }
    select.dataset.populated = "true";
  }
  select.value = model;

  const el = document.getElementById("keyStatus");
  el.textContent = (has_key ? "Key set for this session." : "No key set.") + ` Model: ${model}`;
  el.className = "key-status " + (has_key ? "ok" : "missing");
}

document.getElementById("saveKeyBtn").addEventListener("click", async () => {
  const input = document.getElementById("apiKeyInput");
  const api_key = input.value.trim();
  const model = document.getElementById("modelSelect").value;
  try {
    await api("/api/key", { method: "POST", body: JSON.stringify({ api_key, model }) });
    input.value = "";
    await refreshKeyStatus();
  } catch (e) {
    alert(e.message);
  }
});

document.getElementById("clearKeyBtn").addEventListener("click", async () => {
  await api("/api/key", { method: "DELETE" });
  await refreshKeyStatus();
});

// ---------------------------------------------------------------- sessions --

async function loadSessions() {
  state.sessions = await api("/api/sessions");
  renderSessionList();
}

function renderSessionList() {
  const list = document.getElementById("sessionList");
  list.innerHTML = "";
  for (const s of state.sessions) {
    const item = document.createElement("div");
    item.className = "session-item" + (s.id === state.sessionId ? " active" : "");
    item.innerHTML = `<div class="title">${escapeHtml(s.title)}</div>
      <div class="sub">${s.message_count} msgs · ${s.upload_count} files · ${s.report_count} reports</div>`;
    item.addEventListener("click", () => selectSession(s.id));
    list.appendChild(item);
  }
}

async function selectSession(id) {
  state.sessionId = id;
  renderSessionList();
  const detail = await api(`/api/sessions/${id}`);
  document.getElementById("noSession").hidden = true;
  document.getElementById("sessionView").hidden = false;
  document.getElementById("sessionTitle").textContent = detail.session.title;
  document.getElementById("sessionMeta").textContent =
    detail.session.carried_from_session_id
      ? `Continues context from session ${detail.session.carried_from_session_id}`
      : "";
  renderUploads(detail.uploads);
  renderConversation(detail.messages);
  renderReports(detail.reports);
  document.getElementById("uploadStatus").textContent = "";
  document.getElementById("analyzeStatus").textContent = "";
}

function renderUploads(uploads) {
  const list = document.getElementById("uploadList");
  list.innerHTML = "";
  for (const u of uploads) {
    const li = document.createElement("li");
    li.textContent = `${u.original_name} — ${u.sheet_count} sheet(s), ${u.row_count} row(s)`;
    list.appendChild(li);
  }
}

function renderConversation(messages) {
  const box = document.getElementById("conversation");
  box.innerHTML = "";
  if (!messages.length) {
    box.innerHTML = '<div class="muted">No analysis run yet.</div>';
    return;
  }
  for (const m of messages) {
    const div = document.createElement("div");
    div.className = "msg " + m.role;
    let body = m.content;
    if (m.role === "assistant") {
      try {
        const parsed = JSON.parse(m.content);
        body = formatAssistantMessage(parsed);
      } catch (_) { /* not JSON, show raw */ }
    }
    div.innerHTML = `<div class="role">${m.role}</div><div>${body}</div>`;
    box.appendChild(div);
  }
  box.scrollTop = box.scrollHeight;
}

function formatAssistantMessage(parsed) {
  let html = `<strong>Summary:</strong> ${escapeHtml(parsed.summary || "")}`;
  if (parsed.root_causes && parsed.root_causes.length) {
    html += "<br><br><strong>Root causes:</strong><ul>";
    for (const rc of parsed.root_causes) {
      html += `<li><strong>${escapeHtml(rc.issue || "")}</strong> (${escapeHtml(rc.confidence || "")} confidence): ${escapeHtml(rc.explanation || "")}</li>`;
    }
    html += "</ul>";
  }
  return html;
}

function renderReports(reports) {
  const list = document.getElementById("reportList");
  list.innerHTML = "";
  if (!reports.length) {
    list.innerHTML = '<li class="muted">No reports generated yet.</li>';
    return;
  }
  for (const r of reports) {
    const li = document.createElement("li");
    const created = new Date(r.created_at).toLocaleString();
    li.innerHTML = `<span>${created}</span>`;
    const a = document.createElement("a");
    a.href = `/api/sessions/${state.sessionId}/reports/${r.id}/download`;
    a.className = "btn btn-secondary btn-sm";
    a.textContent = "Download .xlsx";
    li.appendChild(a);
    list.appendChild(li);
  }
}

// ------------------------------------------------------------ new session --

const modal = document.getElementById("newSessionModal");

document.getElementById("newSessionBtn").addEventListener("click", () => {
  const select = document.getElementById("carrySelect");
  select.innerHTML = '<option value="">None</option>';
  for (const s of state.sessions) {
    const opt = document.createElement("option");
    opt.value = s.id;
    opt.textContent = s.title;
    select.appendChild(opt);
  }
  document.getElementById("newSessionTitle").value = "";
  modal.hidden = false;
});

document.getElementById("cancelSessionBtn").addEventListener("click", () => { modal.hidden = true; });

document.getElementById("createSessionBtn").addEventListener("click", async () => {
  const title = document.getElementById("newSessionTitle").value.trim();
  const carry_from_session_id = document.getElementById("carrySelect").value || null;
  try {
    const created = await api("/api/sessions", {
      method: "POST",
      body: JSON.stringify({ title, carry_from_session_id }),
    });
    modal.hidden = true;
    await loadSessions();
    await selectSession(created.id);
  } catch (e) {
    alert(e.message);
  }
});

// -------------------------------------------------------------- uploading --

async function uploadFiles(fileList) {
  const files = Array.from(fileList).filter(f => f.name.toLowerCase().endsWith(".xlsx"));
  if (!files.length) {
    document.getElementById("uploadStatus").textContent = "No .xlsx files found in selection.";
    return;
  }
  const formData = new FormData();
  for (const f of files) {
    formData.append("files", f, f.webkitRelativePath || f.name);
  }
  document.getElementById("uploadStatus").textContent = `Uploading ${files.length} file(s)...`;
  try {
    const result = await api(`/api/sessions/${state.sessionId}/upload`, { method: "POST", body: formData });
    document.getElementById("uploadStatus").textContent =
      `Uploaded ${result.saved.length} file(s).` + (result.skipped.length ? ` Skipped: ${result.skipped.join(", ")}` : "");
    const detail = await api(`/api/sessions/${state.sessionId}`);
    renderUploads(detail.uploads);
    await loadSessions();
  } catch (e) {
    document.getElementById("uploadStatus").textContent = "Upload failed: " + e.message;
  }
}

document.getElementById("fileInput").addEventListener("change", (e) => uploadFiles(e.target.files));
document.getElementById("folderInput").addEventListener("change", (e) => uploadFiles(e.target.files));

const dropZone = document.getElementById("dropZone");
dropZone.addEventListener("dragover", (e) => { e.preventDefault(); dropZone.classList.add("dragover"); });
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));
dropZone.addEventListener("drop", async (e) => {
  e.preventDefault();
  dropZone.classList.remove("dragover");
  if (!state.sessionId) { alert("Select or create a session first."); return; }
  const files = await collectFilesFromDataTransfer(e.dataTransfer);
  uploadFiles(files);
});

async function collectFilesFromDataTransfer(dataTransfer) {
  const items = dataTransfer.items;
  if (!items || !items[0] || !items[0].webkitGetAsEntry) {
    return dataTransfer.files;
  }
  const files = [];
  const entries = Array.from(items).map(it => it.webkitGetAsEntry()).filter(Boolean);
  async function walk(entry) {
    if (entry.isFile) {
      await new Promise((resolve) => entry.file((file) => { files.push(file); resolve(); }));
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      const entries = await new Promise((resolve) => reader.readEntries(resolve));
      for (const e of entries) await walk(e);
    }
  }
  for (const entry of entries) await walk(entry);
  return files;
}

// -------------------------------------------------------------- analyzing --

document.getElementById("analyzeBtn").addEventListener("click", async () => {
  if (!state.sessionId) return;
  const extra_instructions = document.getElementById("extraInstructions").value.trim();
  const btn = document.getElementById("analyzeBtn");
  const status = document.getElementById("analyzeStatus");
  btn.disabled = true;
  status.textContent = "Analyzing with Claude... this can take up to a minute.";
  try {
    await api(`/api/sessions/${state.sessionId}/analyze`, {
      method: "POST",
      body: JSON.stringify({ extra_instructions }),
    });
    status.textContent = "Done.";
    const detail = await api(`/api/sessions/${state.sessionId}`);
    renderConversation(detail.messages);
    renderReports(detail.reports);
    await loadSessions();
  } catch (e) {
    status.textContent = "Failed: " + e.message;
  } finally {
    btn.disabled = false;
  }
});

// ------------------------------------------------------------------ utils --

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

// -------------------------------------------------------------------- init --

(async function init() {
  await refreshKeyStatus();
  await loadSessions();
})();
