// FUEL Response Parser — browser UI. All parsing happens in worker.js (Pyodide) on this computer.
"use strict";

const worker = new Worker("worker.js");
let seq = 0;
const pending = new Map();
worker.onmessage = (ev) => {
  const { id, ok, data, error } = ev.data;
  const p = pending.get(id);
  pending.delete(id);
  ok ? p.resolve(data) : p.reject(new Error(error));
};
const call = (op, args) => new Promise((resolve, reject) => {
  const id = ++seq;
  pending.set(id, { resolve, reject });
  worker.postMessage({ id, op, args });
});

const $ = (s) => document.querySelector(s);
const el = (tag, attrs = {}, ...kids) => {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") n.className = v; else if (k === "text") n.textContent = v; else if (k.startsWith("on")) n.addEventListener(k.slice(2), v); else n.setAttribute(k, v);
  }
  for (const k of kids) n.append(k);
  return n;
};
const fmt = (n) => Number(n || 0).toLocaleString();
const store = {
  get(k, d) { try { const v = localStorage.getItem(k); return v === null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* storage unavailable */ } },
};

const state = {
  files: { sources: [], template: null, sellers: null },
  flags: { sheet: "", accept_suggestions: false, ignore_unmapped: false },
  result: null,
  assignments: store.get("fuel.teamAssignments", {}), // team key -> slot (1|2|3)
};
let engineReady;
const engineReadyPromise = new Promise((r) => (engineReady = r));

// ---------- engine ----------
(async () => {
  try {
    const info = await call("init");
    $("#engine").classList.add("ready");
    $("#engine-text").textContent = `Parser engine ready · ${info.version}`;
    engineReady();
    updateCheckButton();
  } catch (e) {
    $("#engine").classList.add("failed");
    $("#engine-text").textContent = "Engine failed to load";
    showAlert("error", "The parser engine could not load. Check your internet connection (the engine is downloaded once from cdn.jsdelivr.net), then reload. " + e.message);
  }
})();

// ---------- navigation ----------
const TITLES = { import: "Import files", summary: "Check results", teams: "Team seller columns", review: "Review", export: "Export", history: "Import history" };
function show(view) {
  document.querySelectorAll(".view").forEach((v) => v.classList.add("hidden"));
  $("#view-" + view).classList.remove("hidden");
  document.querySelectorAll(".nav-item").forEach((b) => b.classList.toggle("active", b.dataset.view === view));
  $("#view-title").textContent = TITLES[view];
  if (view === "history") renderHistory();
}
document.querySelectorAll(".nav-item").forEach((b) => b.addEventListener("click", () => show(b.dataset.view)));

function showAlert(kind, message, extra) {
  const a = $("#alert");
  a.className = "alert " + kind;
  a.replaceChildren(el("div", { text: message }));
  if (extra) a.append(extra);
}
const clearAlert = () => $("#alert").classList.add("hidden");
function busy(on, text) {
  $("#busy").classList.toggle("hidden", !on);
  if (text) $("#busy-text").textContent = text;
}

// ---------- files ----------
const dz = $("#dropzone");
dz.addEventListener("dragover", (e) => { e.preventDefault(); dz.classList.add("over"); });
dz.addEventListener("dragleave", () => dz.classList.remove("over"));
dz.addEventListener("drop", (e) => { e.preventDefault(); dz.classList.remove("over"); addFiles([...e.dataTransfer.files]); });
$("#file-input").addEventListener("change", (e) => { const files = [...e.target.files]; e.target.value = ""; addFiles(files); }); // copy: FileList is live

// Each row is also a drop target and picker, so a file can be placed explicitly.
document.querySelectorAll(".file-row").forEach((row) => {
  const role = row.dataset.role;
  const input = row.querySelector("input[type=file]");
  row.querySelector(".choose").addEventListener("click", () => input.click());
  input.addEventListener("change", () => { const files = [...input.files]; input.value = ""; if (files.length) addFiles(files, role); });
  row.addEventListener("dragover", (e) => { e.preventDefault(); e.stopPropagation(); row.classList.add("over"); });
  row.addEventListener("dragleave", () => row.classList.remove("over"));
  row.addEventListener("drop", (e) => { e.preventDefault(); e.stopPropagation(); row.classList.remove("over"); addFiles([...e.dataTransfer.files], role); });
  const rm = row.querySelector(".remove");
  if (rm) rm.addEventListener("click", () => { state.files[role] = null; renderFiles(); });
});

async function readAndClassify(f) {
  const file = { name: f.name, buffer: await f.arrayBuffer() };
  file.info = JSON.parse(await call("classify", { file }));
  return file;
}

// targetRole: "sources" | "template" | "sellers" when dropped on a row; undefined for the large drop area.
async function addFiles(list, targetRole) {
  clearAlert();
  await engineReadyPromise;
  busy(true, `Reading ${list.length} file${list.length > 1 ? "s" : ""}…`);
  const notes = [];
  try {
    for (const f of list) {
      if (!/\.(xlsx|csv)$/i.test(f.name)) { notes.push(`${f.name}: only .xlsx and .csv files are supported.`); continue; }
      const file = await readAndClassify(f);
      if (file.info.kind === "error") { notes.push(`${f.name} could not be read (${file.info.message}).`); continue; }
      let role = targetRole;
      if (!role) role = file.info.kind === "sellers" ? "sellers" : file.info.kind === "template" ? "template" : "sources";
      if (role === "sources") {
        if (state.files.sources.some((s) => s.name === file.name)) {
          notes.push(`${file.name} is already in the list; the newer copy replaced it.`);
          state.files.sources = state.files.sources.filter((s) => s.name !== file.name);
        }
        state.files.sources.push(file);
      } else {
        if (list.length > 1 && targetRole) notes.push(`Only one ${role === "template" ? "template" : "seller directory"} can be used: ${list[list.length - 1].name === f.name ? "the last file dropped" : f.name} was kept.`);
        state.files[role] = file;
      }
    }
  } finally { busy(false); }
  state.flags = { sheet: "", accept_suggestions: false, ignore_unmapped: false };
  renderFiles();
  if (notes.length) showAlert("warn", notes.join(" "));
}

function describe(role, file) {
  const i = file.info || {};
  if (role === "sellers") return i.kind === "sellers" ? ["ok", "Seller directory CSV recognised"] : ["warn", "Does not look like a seller directory (first column should be seller_name)."];
  if (role === "sources") {
    const others = i.other_sheets ? ` · ${i.other_sheets} other sheet${i.other_sheets > 1 ? "s" : ""} ignored` : "";
    if (i.source_sheet) return ["ok", `Response export recognised · using sheet “${i.source_sheet}”${others}`];
    if (i.template_sheet) return ["err", "This is a MASTERLIST template, not a response export. Remove it (the layout is built in)."];
    return ["err", "No response sheet found (Timestamp, SALES LEADER, Lead Source, Seller Name). Remove this file or check it."];
  }
  if (i.template_sheet && i.template_exact)
    return ["ok", i.template_reordered ? `Template recognised · sheet “${i.template_sheet}” · columns re-arranged: the output will follow this template`
                                       : `Template recognised · sheet “${i.template_sheet}” · same arrangement as the built-in layout`];
  if (i.template_sheet) return ["err", `Sheet “${i.template_sheet}” is missing ${i.template_missing} of the 35 MASTERLIST headers or has extra ones. Use the built-in layout or a complete template.`];
  if (i.source_sheet) return ["err", "This is a response export, not a template. Add it to Response exports instead."];
  return ["err", "No MASTERLIST header found in any sheet."];
}

function renderFiles() {
  // Response exports: a list.
  const list = $("#source-list");
  list.replaceChildren(...state.files.sources.map((f, idx) => {
    const [cls, text] = describe("sources", f);
    return el("li", {}, el("span", { class: "fname", text: `${f.name} · ${(f.buffer.byteLength / 1024).toFixed(0)} KB` }),
      el("button", { class: "link", text: "Remove", onclick: () => { state.files.sources.splice(idx, 1); renderFiles(); } }),
      el("span", { class: "fnote " + cls, text }));
  }));
  $("#source-empty").classList.toggle("hidden", state.files.sources.length > 0);
  // Template: built-in unless a file is supplied.
  const trow = document.querySelector(".file-row[data-role=template]");
  const t = state.files.template;
  trow.querySelector(".fname").textContent = t ? `${t.name} · ${(t.buffer.byteLength / 1024).toFixed(0)} KB` : "Built-in MASTERLIST layout · 35 fixed columns";
  const [tcls, ttext] = t ? describe("template", t) : ["ok", "No template file needed. Add one only if the MASTERLIST columns were re-arranged."];
  trow.querySelector(".fnote").className = "fnote " + tcls;
  trow.querySelector(".fnote").textContent = ttext;
  trow.querySelector(".remove").classList.toggle("hidden", !t);
  trow.querySelector(".choose").textContent = t ? "Replace template" : "Use a template file";
  // Seller directory.
  const srow = document.querySelector(".file-row[data-role=sellers]");
  const s = state.files.sellers;
  srow.querySelector(".fname").textContent = s ? `${s.name} · ${(s.buffer.byteLength / 1024).toFixed(0)} KB` : "Drop a file on this row or choose one";
  srow.querySelector(".fname").classList.toggle("muted", !s);
  const [scls, stext] = s ? describe("sellers", s) : ["", ""];
  srow.querySelector(".fnote").className = "fnote " + scls;
  srow.querySelector(".fnote").textContent = stext;
  srow.querySelector(".remove").classList.toggle("hidden", !s);
  srow.querySelector(".choose").textContent = s ? "Replace" : "Choose file";
  updateCheckButton();
}

function updateCheckButton() {
  const n = state.files.sources.length;
  const wrong = state.files.sources.some((f) => describe("sources", f)[0] === "err")
    || (state.files.template && describe("template", state.files.template)[0] === "err");
  $("#btn-check").disabled = !n || wrong || !$("#engine").classList.contains("ready");
  $("#check-hint").textContent = !n ? "Add at least one response export."
    : wrong ? "Remove or replace the file marked in red first."
    : `Validation only: nothing is exported yet.${n > 1 ? ` The ${n} exports will be combined into one import file.` : ""}`;
}

// ---------- running ----------
function overrides() {
  const by = $("#approved-by").value.trim();
  const today = new Date().toISOString().slice(0, 10);
  const out = {};
  for (const [key, slot] of Object.entries(state.assignments)) if (slot) out[key] = { slot: Number(slot), approved_by: by, approved_on: today };
  return out;
}
function params(mode, extra = {}) {
  return {
    mode, sheet: state.flags.sheet || $("#opt-sheet").value.trim() || null,
    timestamp_format: $("#opt-tsfmt").value.trim() || null,
    accept_suggestions: state.flags.accept_suggestions, ignore_unmapped: state.flags.ignore_unmapped,
    overrides: overrides(), ...extra,
  };
}
async function runParser(mode, extra) {
  busy(true, mode === "export" ? "Creating the import file…" : "Checking every response… (large exports take up to a minute)");
  try {
    return await call("run", { params: params(mode, extra), files: state.files });
  } finally { busy(false); }
}

$("#btn-check").addEventListener("click", check);
async function check() {
  clearAlert();
  const { result } = await runParser("dry");
  if (result.status === "needs_confirmation") return confirmPanel(result);
  if (result.status === "error") return showAlert("error", result.message);
  state.result = result;
  document.querySelectorAll(".nav-item[disabled]").forEach((b) => (b.disabled = false));
  renderSummary(); renderTeams(); renderReview(); renderExport();
  show("summary");
}

function confirmPanel(res) {
  const box = el("div");
  if (res.sheet && !res.items.length) {
    box.append(el("p", { text: `Proposed response sheet: "${res.sheet}". Check that it holds the raw form responses.` }),
      el("button", { class: "btn primary", text: `Use sheet "${res.sheet}"`, onclick: () => { state.flags.sheet = res.sheet; check(); } }));
  } else {
    const ul = el("ul");
    for (const i of res.items) ul.append(el("li", { text: `Column ${i.column} “${i.header}”: ${i.kind === "SUGGESTED" ? "looks like " + i.target : "not recognised"} (${i.note})` }));
    box.append(ul);
    const acts = el("div", { class: "actions" });
    if (res.items.some((i) => i.kind === "SUGGESTED")) acts.append(el("button", { class: "btn primary", text: "Confirm the suggested matches", onclick: () => { state.flags.accept_suggestions = true; check(); } }));
    if (res.items.some((i) => i.kind === "UNMAPPED")) acts.append(el("button", { class: "btn", text: "Ignore the unrecognised columns", onclick: () => { state.flags.ignore_unmapped = true; check(); } }));
    box.append(acts);
  }
  showAlert("warn", "The file layout needs your confirmation before anything can be exported. " + res.message, box);
  show("import");
}

// ---------- rendering ----------
function renderSummary() {
  const r = state.result, c = r.counts;
  const m = $("#metrics");
  m.replaceChildren();
  const metric = (label, value, cls = "") => m.append(el("div", { class: "metric " + cls }, el("div", { class: "label", text: label }), el("div", { class: "value", text: fmt(value) })));
  metric("Responses", c.responses);
  metric("Ready to export", c.records_exported, c.records_exported ? "good" : "");
  metric("Held for review", c.records_review, c.records_review ? "hold" : "");
  metric("Duplicates held", c.records_duplicate_held);
  metric("Already imported", c.records_already_imported);
  const ul = $("#checks");
  ul.replaceChildren(...r.checks.map((k) => el("li", { text: `${k.ok ? "✅" : "❌"} ${k.name}` })));
  const files = r.sources.map((s) => `${s.name} (sheet “${s.sheet}”, ${fmt(s.responses)} responses` +
    (s.ignored.length ? `, ${s.ignored.length} other sheet${s.ignored.length > 1 ? "s" : ""} ignored)` : ")")).join(" · ");
  $("#headline").textContent = `${r.sources.length} response export${r.sources.length > 1 ? "s" : ""}: ${files}. ` +
    `MASTERLIST layout: ${r.layout.origin === "built-in" ? "built-in" : r.layout.origin + (r.layout.reordered ? " (re-arranged)" : "")}. ` +
    `Exported: ${fmt(r.exported_verified)} verified · ${fmt(r.exported_owner_approved)} owner-approved (not verified).` +
    (r.previous_batches.length ? ` Some of these files were exported before (${r.previous_batches.map((b) => b.batch_id).join(", ")}).` : "");
  const tb = $("#issues");
  tb.replaceChildren(...r.issues.map((i) => el("tr", {}, el("td", { text: i.code }), el("td", {}, el("span", { class: "sev " + i.severity, text: i.severity })),
    el("td", { class: "num", text: fmt(i.count) }), el("td", { class: "muted", text: i.code === "SLOT_UNRESOLVED" ? "Assign seller columns for the team on the “Team seller columns” screen, or leave the records in review." : i.action }))));
  if (!c.records_exported && c.records_review) {
    showAlert("warn", "No records are ready to export yet. Most are held because their team has no seller-column assignment. Open “Team seller columns” to decide, or export now to download the audit and review file.");
  }
}

function renderTeams() {
  const tb = $("#teams");
  tb.replaceChildren();
  $("#approved-by").value = $("#approved-by").value || store.get("fuel.approvedBy", "");
  $("#teams-sheet").textContent = state.result.sources.map((s) => `${s.name}: ${s.sheet}`).join("; ");
  for (const t of state.result.leaders) {
    const locked = t.status === "VERIFIED";
    const sel = el("select", { "data-key": t.key });
    [["", "Unresolved — keep in review"], ["1", "G/H · Seller Name"], ["2", "I/J · Seller Name1"], ["3", "K/L · Seller Name2"]]
      .forEach(([v, label]) => sel.append(el("option", { value: v, text: label })));
    sel.value = locked ? String(t.slot) : String(state.assignments[t.key] || "");
    sel.disabled = locked;
    sel.addEventListener("change", () => { state.assignments[t.key] = sel.value ? Number(sel.value) : 0; });
    const statusText = t.status === "OWNER_APPROVED_UNVERIFIED" ? "Owner-approved (not verified)" : t.status === "VERIFIED" ? "Verified" : "Unresolved";
    tb.append(el("tr", {}, el("td", { text: t.labels.join(" / ") }), el("td", { class: "num", text: fmt(t.records) }),
      el("td", {}, el("span", { class: "status-pill", text: statusText })), el("td", {}, sel)));
  }
}
$("#btn-apply").addEventListener("click", async () => {
  const assigned = Object.values(state.assignments).some(Boolean);
  if (assigned && !$("#approved-by").value.trim()) return showAlert("error", "Enter your name in “Approved by” before assigning seller columns.");
  for (const k of Object.keys(state.assignments)) if (!state.assignments[k]) delete state.assignments[k];
  store.set("fuel.teamAssignments", state.assignments);
  store.set("fuel.approvedBy", $("#approved-by").value.trim());
  await check();
});
$("#btn-clear-teams").addEventListener("click", () => { state.assignments = {}; store.set("fuel.teamAssignments", {}); renderTeams(); });

function renderReview() {
  const r = state.result;
  $("#review-count").textContent = `· ${fmt(r.review_total)}` + (r.review_total > r.review.length ? ` (first ${r.review.length} shown)` : "");
  const draw = () => {
    const q = $("#review-filter").value.trim().toLowerCase();
    const rows = r.review.filter((x) => !q || `${x.reasons} ${x.leader} ${x.lead_source} ${x.status}`.toLowerCase().includes(q));
    $("#review").replaceChildren(...rows.map((x) => el("tr", {}, el("td", { class: "num", text: x.row }), el("td", { text: x.status }),
      el("td", { text: x.reasons }), el("td", { text: x.leader }), el("td", { text: x.lead_source }))));
  };
  $("#review-filter").oninput = draw;
  draw();
}

function renderExport() {
  const n = state.result.counts.records_exported;
  $("#export-ready").textContent = n
    ? `${fmt(n)} record(s) are ready. Enter the last Id currently in the live MASTERLIST so new Ids continue from it.`
    : "No records are ready. You can still create the files to download the full audit and review workbook.";
  $("#last-id").value = "";
  $("#last-id").placeholder = n ? "Required" : "Not needed";
}
$("#btn-export").addEventListener("click", async () => {
  clearAlert();
  const n = state.result.counts.records_exported;
  const raw = $("#last-id").value.trim();
  if (n && !/^\d+$/.test(raw)) return showAlert("error", "Enter the last Id in the live MASTERLIST (a whole number).");
  const { result, outputs } = await runParser("export", { last_id: raw ? Number(raw) : null, leads_log: $("#leads-log").checked });
  if (result.status !== "ok") return showAlert("error", result.message || "Export needs confirmation: run Check files again.");
  state.result = result;
  renderSummary(); renderTeams(); renderReview();
  const list = $("#download-list");
  list.replaceChildren();
  const labels = { import: "Import file (MASTERLIST)", audit: "Audit & review workbook", leads_log: "Leads Log (17 columns)" };
  for (const [kind, f] of Object.entries(outputs)) {
    const url = URL.createObjectURL(new Blob([f.bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    list.append(el("a", { class: "btn", href: url, download: f.name, text: `⬇ ${labels[kind] || kind} — ${f.name}` }));
  }
  $("#downloads").classList.remove("hidden");
  showAlert("ok", `Batch ${result.batch_id}: ${fmt(result.counts.records_exported)} exported, ${fmt(result.counts.records_review)} held for review. Append the import file to the live MASTERLIST yourself.`);
  show("export");
});

async function renderHistory() {
  await engineReadyPromise;
  const rows = await call("batches");
  const tb = $("#history");
  tb.replaceChildren(...(rows.length ? rows.slice().reverse().map((b) => el("tr", {}, el("td", { text: b.batch_id }), el("td", { text: b.created_at }),
    el("td", { text: b.status }), el("td", { class: "num", text: fmt(b.records_exported) }), el("td", { text: b.source_file_name }),
    el("td", {}, b.status === "exported" ? el("button", { class: "link", text: "Void", onclick: async () => {
      if (!confirm(`Void batch ${b.batch_id}? Its records become exportable again. Only do this if you did not append it to the live MASTERLIST.`)) return;
      await call("voidBatch", { batchId: b.batch_id });
      renderHistory();
    } }) : ""))) : [el("tr", {}, el("td", { colspan: "6", class: "muted", text: "No batches yet." }))]));
}
