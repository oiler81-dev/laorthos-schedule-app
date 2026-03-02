const DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday"];
const $ = (sel) => document.querySelector(sel);

const LOCATIONS = [
  { name: "Wilshire", code: "WIL" },
  { name: "Santa Fe Springs", code: "SFS" },
  { name: "East LA", code: "ELA" },
  { name: "Glendale", code: "GLE" },
  { name: "Tarzana", code: "TZ" },
  { name: "Encino", code: "ENC" },
  { name: "Valencia", code: "VAL" },
  { name: "Thousand Oaks", code: "TO" },
];

const state = {
  auth: { email: null, editor: false },
  view: "my",

  weeks: [],
  currentWeekOf: null,

  published: null,
  draft: null,
};

init().catch(err => {
  console.error(err);
  toast("Init failed. Check console.");
});

async function init(){
  wireTabs();
  wireButtons();

  await loadAuth();
  await loadWeeks();

  state.currentWeekOf = state.weeks[0] || mondayOf(new Date());

  await loadPublished(state.currentWeekOf);
  renderWeekPickers();
  renderAll();
}

/* ---------------- Auth ---------------- */

async function loadAuth(){
  const res = await fetch("./api/me", { cache: "no-store" });
  const data = await res.json();
  state.auth.email = data.email;
  state.auth.editor = !!data.editor;

  $("#authLine").textContent = state.auth.email
    ? `${state.auth.email}${state.auth.editor ? " • Editor" : ""}`
    : `Not logged in`;

  $("#builderTab").style.display = state.auth.editor ? "" : "none";

  $("#btnLogin").style.display = state.auth.email ? "none" : "";
  $("#btnLogout").style.display = state.auth.email ? "" : "none";
}

function login(){ window.location.href = "/.auth/login/aad"; }
function logout(){ window.location.href = "/.auth/logout"; }

/* ---------------- API ---------------- */

async function loadWeeks(){
  try{
    const res = await fetch("./api/weeks", { cache: "no-store" });
    if(!res.ok) throw new Error(await res.text());
    const data = await res.json();
    state.weeks = (data.weeks || []).slice();
    state.weeks.sort((a,b)=> a < b ? 1 : -1);
  }catch(e){
    console.warn("weeks load failed", e);
    state.weeks = [];
  }
}

async function loadPublished(weekOf){
  try{
    const res = await fetch(`./api/schedule?weekOf=${encodeURIComponent(weekOf)}&mode=published`, { cache: "no-store" });
    if(!res.ok){ state.published = null; return; }
    const data = await res.json();
    state.published = data.data;
  }catch{
    state.published = null;
  }
}

async function loadDraft(weekOf){
  try{
    const res = await fetch(`./api/schedule?weekOf=${encodeURIComponent(weekOf)}&mode=draft`, { cache: "no-store" });
    if(!res.ok){ state.draft = null; return; }
    const data = await res.json();
    state.draft = data.data;
  }catch{
    state.draft = null;
  }
}

async function saveSchedule(mode, payload){
  const res = await fetch("./api/schedule", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ weekOf: payload.weekOf, mode, data: payload })
  });
  if(!res.ok) throw new Error(await res.text());
}

/* ---- Default roster/providers endpoints (Option C) ---- */

async function getDefaultRoster(){
  const res = await fetch("./api/roster", { cache: "no-store" });
  if(!res.ok) throw new Error(await res.text());
  const data = await res.json();
  return data.roster || null;
}

async function saveDefaultRoster(roster){
  const res = await fetch("./api/roster", {
    method: "POST",
    headers: { "Content-Type":"application/json" },
    body: JSON.stringify({ roster })
  });
  if(!res.ok) throw new Error(await res.text());
}

async function getDefaultProviders(){
  const res = await fetch("./api/providers", { cache: "no-store" });
  if(!res.ok) throw new Error(await res.text());
  const data = await res.json();
  return data.providers || null;
}

async function saveDefaultProviders(providers){
  const res = await fetch("./api/providers", {
    method: "POST",
    headers: { "Content-Type":"application/json" },
    body: JSON.stringify({ providers })
  });
  if(!res.ok) throw new Error(await res.text());
}

/* ---------------- Tabs ---------------- */

function wireTabs(){
  document.querySelectorAll(".tab").forEach(btn=>{
    btn.addEventListener("click", ()=> setView(btn.dataset.view));
  });
}

function setView(view){
  state.view = view;
  document.querySelectorAll(".tab").forEach(t=>{
    const is = t.dataset.view === view;
    t.classList.toggle("is-active", is);
    t.setAttribute("aria-selected", is ? "true" : "false");
  });
  ["my","everyone","builder"].forEach(v=> $(`#view-${v}`).hidden = (v !== view));
}

/* ---------------- Buttons ---------------- */

function wireButtons(){
  $("#btnPrint").addEventListener("click", ()=> window.print());
  $("#btnExportCSV").addEventListener("click", exportPublishedToCSV);

  $("#btnLogin").addEventListener("click", login);
  $("#btnLogout").addEventListener("click", logout);

  // Import schedule CSV (existing)
  $("#btnImportCSV").addEventListener("click", ()=> $("#csvFile").click());
  $("#csvFile").addEventListener("change", async (e)=>{
    const f = e.target.files?.[0];
    if(!f) return;
    const text = await f.text();
    try{
      const parsed = parseScheduleCSV(text);
      state.draft = parsed;
      state.currentWeekOf = parsed.weekOf;

      if(!state.weeks.includes(parsed.weekOf)){
        state.weeks.unshift(parsed.weekOf);
        state.weeks = Array.from(new Set(state.weeks)).sort((a,b)=> a < b ? 1 : -1);
      }

      renderWeekPickers();
      setAllWeekPickers(parsed.weekOf);

      toast("Schedule CSV imported into Draft.");
      $("#csvFile").value = "";
      renderAll();
    }catch(err){
      console.error(err);
      toast("CSV import failed. Make sure it’s the schedule export.");
    }
  });

  $("#btnLoadDraft").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    await loadDraft(state.currentWeekOf);
    toast(state.draft ? "Draft loaded." : "No draft found for this week.");
    renderAll();
  });

  $("#btnLoadPublished").addEventListener("click", async ()=>{
    await loadPublished(state.currentWeekOf);
    toast(state.published ? "Published loaded." : "No published schedule found.");
    renderAll();
  });

  $("#btnSaveDraft").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;

    try{
      await saveSchedule("draft", payload);
      toast("Draft saved.");
    }catch(e){
      console.error(e);
      toast("Draft save failed. Check API + env vars.");
    }
  });

  $("#btnPublish").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;

    const issues = validate(payload);
    if(issues.length){
      openModal({
        title: `Fix before publish (${issues.length})`,
        body: `<div class="muted" style="line-height:1.45">${issues.map(x=>`• ${escapeHtml(x)}`).join("<br>")}</div>`,
        foot: `<button class="btn btnGhost" data-close="1">Close</button>`
      });
      return;
    }

    try{
      await saveSchedule("published", payload);
      toast("Published.");
      await loadPublished(payload.weekOf);
      renderAll();
    }catch(e){
      console.error(e);
      toast("Publish failed. Check API + permissions.");
    }
  });

  $("#btnValidate").addEventListener("click", ()=>{
    const payload = ensureDraftWeek();
    if(!payload) return;
    const issues = validate(payload);
    if(!issues.length) return toast("Validation passed.");
    openModal({
      title: `Validation issues (${issues.length})`,
      body: `<div class="muted" style="line-height:1.45">${issues.map(x=>`• ${escapeHtml(x)}`).join("<br>")}</div>`,
      foot: `<button class="btn btnGhost" data-close="1">Close</button>`
    });
  });

  $("#btnAddTemplate").addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    openAddTemplateModal();
  });

  $("#btnAutoGuessEmails").addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;

    payload.roster.forEach(r=>{
      if(r.email) return;
      const parts = r.name.toLowerCase().split(/\s+/).filter(Boolean);
      if(parts.length < 2) return;
      const first = parts[0];
      const last = parts[parts.length-1];
      r.email = `${first[0]}${last}@unitymsk.com`;
    });

    toast("Auto-guess applied. Review and correct.");
    renderBuilder();
  });

  // Default roster/providers buttons (Option C)
  $("#btnLoadDefaultRoster").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;

    try{
      const roster = await getDefaultRoster();
      if(!roster?.length) return toast("No default roster saved yet.");
      payload.roster = roster.map(r=>({ id:r.id, name:r.name, role:r.role, email:r.email || "" }));
      ensureEntriesForRoster(payload);
      toast("Default roster loaded into draft.");
      renderBuilder();
    }catch(e){
      console.error(e);
      toast("Failed to load default roster.");
    }
  });

  $("#btnSaveDefaultRoster").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;

    if(!payload.roster?.length) return toast("Draft roster is empty.");

    try{
      await saveDefaultRoster(payload.roster);
      toast("Default roster saved.");
    }catch(e){
      console.error(e);
      toast("Failed to save default roster.");
    }
  });

  $("#btnLoadDefaultProviders").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;

    try{
      const providers = await getDefaultProviders();
      if(!providers?.length) return toast("No default providers saved yet.");
      payload.providers = normalizeProviders(providers);
      payload.providerAssignments ||= {};
      payload.xrLocationAssignments ||= {};
      toast("Default providers loaded into draft.");
      renderBuilder();
    }catch(e){
      console.error(e);
      toast("Failed to load default providers.");
    }
  });

  $("#btnSaveDefaultProviders").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;

    const list = (payload.providers || []).map(p => (typeof p === "string" ? p : p?.name)).map(s=> (s||"").trim()).filter(Boolean);
    if(!list.length) return toast("No providers to save.");

    try{
      await saveDefaultProviders(list);
      toast("Default providers saved.");
    }catch(e){
      console.error(e);
      toast("Failed to save default providers.");
    }
  });

  // New blank week (auto loads default roster + providers)
  $("#btnNewBlankWeek").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");

    const week = state.currentWeekOf;
    const base = (state.published && state.published.weekOf === week) ? state.published : null;

    state.draft = await createBlankDraftFromDefaults(week, base);
    toast(state.draft.roster.length ? "Blank draft created (defaults loaded)." : "Blank draft created (no defaults found).");
    renderAll();
  });

  // Provider Builder buttons
  $("#btnAddProvider").addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;

    const name = prompt("Provider name (ex: Dr. Pelton)");
    if(!name) return;
    payload.providers ||= [];
    payload.providers.push({ id: providerId(name), name: name.trim() });

    payload.providerAssignments ||= {};
    toast("Provider added.");
    renderProviderBuilder(payload);
  });

  $("#btnAddProviderBulk").addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;

    const raw = ($("#providerBulk").value || "").trim();
    if(!raw) return toast("Paste provider names first.");

    const lines = raw.split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
    payload.providers ||= [];
    const existing = new Set(payload.providers.map(p=>p.id));

    let added = 0;
    lines.forEach(n=>{
      const id = providerId(n);
      if(existing.has(id)) return;
      payload.providers.push({ id, name: n });
      existing.add(id);
      added++;
    });

    $("#providerBulk").value = "";
    payload.providerAssignments ||= {};
    payload.xrLocationAssignments ||= {};
    toast(`Added ${added} provider(s).`);
    renderProviderBuilder(payload);
  });

  // Filters
  $("#roleFilter").addEventListener("change", renderEveryone);
  $("#searchFilter").addEventListener("input", renderEveryone);

  // Modal close
  $("#modal").addEventListener("click", (e)=>{
    if(e.target?.dataset?.close) closeModal();
  });
}

/* ---------------- Week pickers ---------------- */

function renderWeekPickers(){
  const weeks = state.weeks.length ? state.weeks : [state.currentWeekOf];
  const opts = weeks.map(w => `<option value="${w}">${w}</option>`).join("");

  $("#weekPickMy").innerHTML = opts;
  $("#weekPickEveryone").innerHTML = opts;
  $("#weekPickBuilder").innerHTML = opts;

  setAllWeekPickers(state.currentWeekOf);

  $("#weekPickMy").onchange = async (e)=>{ await changeWeek(e.target.value); };
  $("#weekPickEveryone").onchange = async (e)=>{ await changeWeek(e.target.value); };
  $("#weekPickBuilder").onchange = async (e)=>{ await changeWeek(e.target.value); };
}

function setAllWeekPickers(week){
  $("#weekPickMy").value = week;
  $("#weekPickEveryone").value = week;
  $("#weekPickBuilder").value = week;
}

async function changeWeek(week){
  state.currentWeekOf = week;
  setAllWeekPickers(week);
  await loadPublished(week);
  renderAll();
}

/* ---------------- Rendering ---------------- */

function renderAll(){
  renderMy();
  renderEveryone();
  renderBuilder();
}

function getActiveForViewing(){
  return state.published;
}

/* ---------------- MY VIEW ---------------- */

function renderMy(){
  const sched = getActiveForViewing();
  const meEmail = (state.auth.email || "").toLowerCase();

  if(!state.auth.email){
    $("#myMatchPill").textContent = "Not logged in";
    $("#todayCard").innerHTML = `<div class="muted">Login required to view your schedule.</div>`;
    $("#weekCards").innerHTML = `<div class="muted">Tap “Login”.</div>`;
    return;
  }

  if(!sched?.roster?.length){
    $("#myMatchPill").textContent = "—";
    $("#todayCard").innerHTML = `<div class="muted">No published schedule found for ${escapeHtml(state.currentWeekOf)}.</div>`;
    $("#weekCards").innerHTML = `<div class="muted">If you’re an editor, publish this week.</div>`;
    return;
  }

  const me = sched.roster.find(r => (r.email || "").toLowerCase() === meEmail) || null;

  if(!me){
    $("#myMatchPill").textContent = "No match";
    $("#todayCard").innerHTML = `<div class="muted">Your email (${escapeHtml(state.auth.email)}) isn’t mapped to the roster for this week.</div>`;
    $("#weekCards").innerHTML = `<div class="muted">Ask an editor to add your email in Builder → Roster Emails.</div>`;
    return;
  }

  $("#myMatchPill").textContent = `${me.name} (${me.role})`;

  const entries = sched.entries?.[me.id] || {};
  const todayName = dayName(new Date());
  const todayVal = (todayName && entries[todayName]) ? templateTextById(sched, entries[todayName]) : null;

  $("#todayCard").innerHTML = `
    <div style="font-weight:900; color:var(--navy)">${formatLongDate(new Date())}</div>
    <div style="margin-top:6px">${todayName ? `<span class="badge">${todayName}</span>` : ""}</div>
    <div style="margin-top:10px; font-size:16px; font-weight:900">${escapeHtml(todayVal || "No assignment today")}</div>
  `;

  $("#weekCards").innerHTML = DAYS.map(d=>{
    const tid = entries[d];
    const txt = tid ? templateTextById(sched, tid) : "—";
    return `
      <div class="dayCard">
        <div>
          <div class="dayName">${d}</div>
          <div class="badge">${me.role}</div>
        </div>
        <div class="dayValue">${escapeHtml(txt)}</div>
      </div>
    `;
  }).join("");
}

/* ---------------- EVERYONE VIEW ---------------- */

function renderEveryone(){
  const sched = getActiveForViewing();
  const table = $("#everyoneGrid");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  if(!sched?.roster?.length){
    thead.innerHTML = "";
    tbody.innerHTML = `<tr><td class="muted" style="padding:14px">No published schedule for ${escapeHtml(state.currentWeekOf)}.</td></tr>`;
    return;
  }

  const role = $("#roleFilter").value;
  const q = ($("#searchFilter").value || "").trim().toLowerCase();

  const rows = sched.roster
    .filter(r => role === "ALL" ? true : r.role === role)
    .filter(r => !q ? true : r.name.toLowerCase().includes(q));

  thead.innerHTML = `
    <tr>
      <th style="min-width:240px;text-align:left;">Team Member</th>
      ${DAYS.map(d=>`<th style="min-width:210px;text-align:left;">${d}</th>`).join("")}
    </tr>
  `;

  tbody.innerHTML = rows.map(r=>{
    const e = sched.entries?.[r.id] || {};
    return `
      <tr>
        <td class="nameCell">
          ${escapeHtml(r.name)}
          <span class="roleTag">${r.role}</span>
        </td>
        ${DAYS.map(d=>{
          const tid = e[d];
          const txt = tid ? templateTextById(sched, tid) : "";
          return `<td>${escapeHtml(txt)}</td>`;
        }).join("")}
      </tr>
    `;
  }).join("");
}

/* ---------------- BUILDER VIEW ---------------- */

function renderBuilder(){
  if(!state.auth.editor){
    $("#builderStatus").textContent = "Editors only";
    return;
  }

  const week = state.currentWeekOf;
  const draft = state.draft && state.draft.weekOf === week ? state.draft : null;
  const pub = state.published && state.published.weekOf === week ? state.published : null;

  $("#builderStatus").textContent =
    draft ? "Draft loaded (editable)" :
    pub ? "No draft loaded (published exists)" :
    "No data loaded for this week";

  // Keep builder clean: no auto-clone anymore. Alex uses New Blank Week.
  const payload = ensureDraftWeek(false);
  if(!payload){
    renderEmptyBuilder();
    return;
  }

  renderRosterEmails(payload);
  renderTemplateList(payload);
  renderBuilderGrid(payload);
  renderProviderBuilder(payload);
}

function renderEmptyBuilder(){
  $("#rosterEmailList").innerHTML = `<div class="muted">Click “New Blank Week”, import schedule CSV, or load a draft.</div>`;
  $("#templateList").innerHTML = `<div class="muted">—</div>`;
  $("#builderGrid").querySelector("thead").innerHTML = "";
  $("#builderGrid").querySelector("tbody").innerHTML = `<tr><td class="muted" style="padding:14px">—</td></tr>`;

  $("#providerList").innerHTML = `<div class="muted">No draft loaded.</div>`;
  $("#xrByLocation").innerHTML = `<div class="muted">—</div>`;
  $("#providerAssignGrid").querySelector("thead").innerHTML = "";
  $("#providerAssignGrid").querySelector("tbody").innerHTML = `<tr><td class="muted" style="padding:14px">—</td></tr>`;
}

function ensureDraftWeek(showToastOnFail=true){
  if(!state.draft){
    if(showToastOnFail) toast("No draft loaded. Click New Blank Week or Import Schedule CSV.");
    return null;
  }
  if(!state.draft.weekOf) state.draft.weekOf = state.currentWeekOf;
  state.draft.templates ||= [];
  state.draft.entries ||= {};
  state.draft.roster ||= [];

  // Provider builder state
  state.draft.providers ||= [];
  state.draft.providerAssignments ||= {};
  state.draft.xrLocationAssignments ||= {};
  state.draft.locations ||= LOCATIONS;

  return state.draft;
}

/* ---------------- Provider Builder ---------------- */

function renderProviderBuilder(payload){
  payload.locations ||= LOCATIONS;
  payload.providers = normalizeProviders(payload.providers || []);
  payload.providerAssignments ||= {};
  payload.xrLocationAssignments ||= {};

  renderProviderList(payload);
  renderXRByLocation(payload);
  renderProviderAssignGrid(payload);
}

function normalizeProviders(list){
  // Accept either strings or objects, normalize to {id,name}
  const out = [];
  const seen = new Set();
  list.forEach(p=>{
    if(!p) return;
    const name = (typeof p === "string" ? p : (p.name || "")).trim();
    if(!name) return;
    const id = (typeof p === "string" ? providerId(name) : (p.id || providerId(name)));
    if(seen.has(id)) return;
    seen.add(id);
    out.push({ id, name });
  });
  return out.sort((a,b)=>a.name.localeCompare(b.name));
}

function renderProviderList(payload){
  const wrap = $("#providerList");
  if(!payload.providers.length){
    wrap.innerHTML = `<div class="muted">No providers yet. Add or bulk add.</div>`;
    return;
  }

  wrap.innerHTML = payload.providers.map(p=>`
    <div class="providerRow">
      <input data-prov-name="${p.id}" type="text" value="${escapeHtml(p.name)}" />
      <button class="iconBtn" title="Delete" data-prov-del="${p.id}">🗑</button>
    </div>
  `).join("");

  wrap.querySelectorAll("input[data-prov-name]").forEach(inp=>{
    inp.addEventListener("input", ()=>{
      const id = inp.dataset.provName;
      const val = (inp.value || "").trim();
      const prov = payload.providers.find(x=>x.id===id);
      if(prov) prov.name = val;
      renderProviderAssignGrid(payload);
    });
  });

  wrap.querySelectorAll("[data-prov-del]").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      const id = btn.dataset.provDel;
      payload.providers = payload.providers.filter(x=>x.id!==id);
      delete payload.providerAssignments[id];
      toast("Provider deleted.");
      renderProviderBuilder(payload);
    });
  });
}

function renderXRByLocation(payload){
  const wrap = $("#xrByLocation");

  const xrOptions = rosterOptions(payload, "XR", true);

  wrap.innerHTML = payload.locations.map(loc=>{
    const cur = payload.xrLocationAssignments?.[loc.code] || "";
    return `
      <div class="locCard">
        <div class="locTitle">${escapeHtml(loc.name)} <span class="roleTag">${escapeHtml(loc.code)}</span></div>
        <div class="field" style="min-width:100%">
          <label>XR (location-only)</label>
          <select data-xr-loc="${loc.code}">
            ${xrOptions(cur)}
          </select>
        </div>
      </div>
    `;
  }).join("");

  wrap.querySelectorAll("select[data-xr-loc]").forEach(sel=>{
    sel.addEventListener("change", ()=>{
      const code = sel.dataset.xrLoc;
      const val = sel.value || "";
      if(!payload.xrLocationAssignments) payload.xrLocationAssignments = {};
      if(!val) delete payload.xrLocationAssignments[code];
      else payload.xrLocationAssignments[code] = val;
    });
  });
}

function renderProviderAssignGrid(payload){
  const table = $("#providerAssignGrid");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  if(!payload.providers.length){
    thead.innerHTML = "";
    tbody.innerHTML = `<tr><td class="muted" style="padding:14px">No providers yet.</td></tr>`;
    return;
  }

  thead.innerHTML = `
    <tr>
      <th style="min-width:260px;text-align:left;">Provider</th>
      <th style="min-width:220px;text-align:left;">Location</th>
      <th style="min-width:260px;text-align:left;">MA</th>
      <th style="min-width:260px;text-align:left;">XR (per provider)</th>
    </tr>
  `;

  const locOptions = (cur)=> [
    `<option value="">— Select —</option>`,
    ...payload.locations.map(l=>`<option value="${l.code}" ${l.code===cur?"selected":""}>${escapeHtml(l.name)} (${escapeHtml(l.code)})</option>`)
  ].join("");

  const maOptions = rosterOptions(payload, "MA", true);
  const xrOptions = rosterOptions(payload, "XR", true);

  tbody.innerHTML = payload.providers.map(p=>{
    const a = payload.providerAssignments?.[p.id] || {};
    return `
      <tr>
        <td class="nameCell">${escapeHtml(p.name)}</td>
        <td>
          <select data-assign-loc="${p.id}">${locOptions(a.locationCode || "")}</select>
        </td>
        <td>
          <select data-assign-ma="${p.id}">
            ${maOptions(a.maStaffId || "")}
          </select>
        </td>
        <td>
          <select data-assign-xr="${p.id}">
            ${xrOptions(a.xrStaffId || "")}
          </select>
          <div class="muted small" style="margin-top:6px">
            Falls back to Location XR if blank.
          </div>
        </td>
      </tr>
    `;
  }).join("");

  tbody.querySelectorAll("select[data-assign-loc]").forEach(sel=>{
    sel.addEventListener("change", ()=>{
      const pid = sel.dataset.assignLoc;
      payload.providerAssignments ||= {};
      payload.providerAssignments[pid] ||= {};
      payload.providerAssignments[pid].locationCode = sel.value || "";
    });
  });

  tbody.querySelectorAll("select[data-assign-ma]").forEach(sel=>{
    sel.addEventListener("change", ()=>{
      const pid = sel.dataset.assignMa;
      payload.providerAssignments ||= {};
      payload.providerAssignments[pid] ||= {};
      payload.providerAssignments[pid].maStaffId = sel.value || "";
    });
  });

  tbody.querySelectorAll("select[data-assign-xr]").forEach(sel=>{
    sel.addEventListener("change", ()=>{
      const pid = sel.dataset.assignXr;
      payload.providerAssignments ||= {};
      payload.providerAssignments[pid] ||= {};
      payload.providerAssignments[pid].xrStaffId = sel.value || "";
    });
  });
}

function rosterOptions(payload, role, allowBlank){
  const list = (payload.roster || []).filter(r=>r.role===role).slice().sort((a,b)=>a.name.localeCompare(b.name));
  return (current)=>{
    const base = allowBlank ? [`<option value="">— None —</option>`] : [];
    const opts = list.map(r=>`<option value="${r.id}" ${r.id===current?"selected":""}>${escapeHtml(r.name)}</option>`);
    return base.concat(opts).join("");
  };
}

/* ---------------- Builder subpanels ---------------- */

function ensureEntriesForRoster(payload){
  payload.entries ||= {};
  payload.roster.forEach(r=>{
    if(!payload.entries[r.id]) payload.entries[r.id] = {};
    DAYS.forEach(d=>{
      if(payload.entries[r.id][d] === undefined) payload.entries[r.id][d] = null;
    });
  });
}

function renderRosterEmails(payload){
  const wrap = $("#rosterEmailList");
  if(!payload.roster.length){
    wrap.innerHTML = `<div class="muted">No roster loaded. Save/Load default roster or import schedule.</div>`;
    return;
  }

  wrap.innerHTML = payload.roster
    .slice()
    .sort((a,b)=>a.name.localeCompare(b.name))
    .map(r => `
      <div class="templateItem">
        <div class="templateTop">
          <div>
            <div class="templateName">${escapeHtml(r.name)} <span class="roleTag">${r.role}</span></div>
            <div class="templateMeta">${escapeHtml(r.id)}</div>
          </div>
        </div>
        <div style="margin-top:10px" class="field">
          <label>Email</label>
          <input data-email-for="${r.id}" type="email" placeholder="name@domain.com" value="${escapeHtml(r.email || "")}">
        </div>
      </div>
    `).join("");

  wrap.querySelectorAll("input[data-email-for]").forEach(inp => {
    inp.addEventListener("input", () => {
      const pid = inp.dataset.emailFor;
      const val = (inp.value || "").trim();
      const person = payload.roster.find(x => x.id === pid);
      if (person) person.email = val;
    });
  });
}

function renderTemplateList(payload){
  const list = $("#templateList");
  payload.templates = normalizeTemplates(payload.templates);

  list.innerHTML = payload.templates.map(t=>{
    const p = t.parsed;
    const meta = p.type === "OFF"
      ? "OFF"
      : `${(p.site ? p.site : "—")} • ${(p.start||"")}–${(p.end||"")}`.replace("–", "–");

    return `
      <div class="templateItem">
        <div class="templateTop">
          <div>
            <div class="templateName">${escapeHtml(p.label || t.raw)}</div>
            <div class="templateMeta">${escapeHtml(meta)}</div>
          </div>
          <div>
            <button class="iconBtn" title="Delete" data-del="${t.id}">🗑</button>
          </div>
        </div>
      </div>
    `;
  }).join("");

  list.querySelectorAll("[data-del]").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      const id = btn.dataset.del;
      payload.templates = payload.templates.filter(x=>x.id !== id);
      Object.keys(payload.entries).forEach(pid=>{
        DAYS.forEach(d=>{
          if(payload.entries[pid]?.[d] === id) payload.entries[pid][d] = null;
        });
      });
      toast("Template deleted.");
      renderBuilder();
    });
  });
}

function renderBuilderGrid(payload){
  const table = $("#builderGrid");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  if(!payload.roster.length){
    thead.innerHTML = "";
    tbody.innerHTML = `<tr><td class="muted" style="padding:14px">No roster loaded.</td></tr>`;
    return;
  }

  ensureEntriesForRoster(payload);

  thead.innerHTML = `
    <tr>
      <th style="min-width:240px;text-align:left;">Team Member</th>
      ${DAYS.map(d=>`<th style="min-width:210px;text-align:left;">${d}</th>`).join("")}
    </tr>
  `;

  tbody.innerHTML = payload.roster.map(r=>{
    const e = payload.entries?.[r.id] || {};
    return `
      <tr>
        <td class="nameCell">
          ${escapeHtml(r.name)} <span class="roleTag">${r.role}</span>
        </td>
        ${DAYS.map(d=>{
          const tid = e[d];
          const txt = tid ? templateTextById(payload, tid) : "";
          return `
            <td>
              <button class="cellBtn" data-pid="${r.id}" data-day="${d}">
                ${escapeHtml(txt || "—")}
              </button>
            </td>
          `;
        }).join("")}
      </tr>
    `;
  }).join("");

  tbody.querySelectorAll(".cellBtn").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      openCellPicker(payload, btn.dataset.pid, btn.dataset.day);
    });
  });
}

function openCellPicker(payload, personId, day){
  const person = payload.roster.find(r=>r.id===personId);
  payload.templates = normalizeTemplates(payload.templates);

  const currentId = payload.entries?.[personId]?.[day] || "";

  const options = [
    `<option value="">— Clear —</option>`,
    ...payload.templates.map(t=>{
      const text = t.parsed.type === "OFF"
        ? "OFF"
        : `${t.parsed.label}${t.parsed.site ? ` (${t.parsed.site})` : ""} ${t.parsed.start ? `${t.parsed.start}-${t.parsed.end}` : ""}`.trim();
      return `<option value="${t.id}" ${t.id===currentId?"selected":""}>${escapeHtml(text)}</option>`;
    })
  ].join("");

  openModal({
    title: `Assign: ${person?.name || "Staff"} • ${day}`,
    body: `
      <div class="field" style="min-width:100%">
        <label>Template</label>
        <select id="cellPick" style="height:46px">${options}</select>
      </div>
    `,
    foot: `
      <button class="btn btnGhost" data-close="1">Cancel</button>
      <button class="btn" id="cellSave">Save</button>
    `,
    onAfter(){
      $("#cellSave").addEventListener("click", ()=>{
        const val = $("#cellPick").value || null;
        if(!payload.entries[personId]) payload.entries[personId] = {};
        payload.entries[personId][day] = val;
        closeModal();
        renderBuilderGrid(payload);
      });
    }
  });
}

function openAddTemplateModal(){
  const payload = ensureDraftWeek();
  if(!payload) return;

  openModal({
    title: "Add template",
    body: `
      <div class="field" style="min-width:100%">
        <label>Template text</label>
        <input id="tplRaw" type="text" placeholder='Example: XR (SFS) 07:50- 04:20' />
        <div class="hint">You can add OFF, TRAINING 08:00- 04:30, FLOAT (Val) 08:00- 04:30</div>
      </div>
    `,
    foot: `
      <button class="btn btnGhost" data-close="1">Cancel</button>
      <button class="btn" id="tplSave">Add</button>
    `,
    onAfter(){
      $("#tplSave").addEventListener("click", ()=>{
        const raw = ($("#tplRaw").value || "").trim();
        if(!raw) return toast("Enter template text.");
        payload.templates.push(normalizeTemplate({ id: idForRaw(raw), raw }));
        closeModal();
        toast("Template added.");
        renderBuilder();
      });
    }
  });
}

/* ---------------- Validation ---------------- */

function validate(payload){
  const issues = [];

  // roster emails
  payload.roster.forEach(r=>{
    if(!r.email) issues.push(`${r.name} missing email`);
  });

  // assignments grid
  payload.roster.forEach(r=>{
    const e = payload.entries?.[r.id] || {};
    DAYS.forEach(d=>{
      const v = e[d];
      if(v === undefined || v === null || v === ""){
        issues.push(`${r.name} missing ${d}`);
      }
    });
  });

  // provider builder basics (soft validation)
  if(payload.providers?.length){
    payload.providers.forEach(p=>{
      const a = payload.providerAssignments?.[p.id];
      if(!a?.locationCode) issues.push(`Provider "${p.name}" missing location`);
      if(!a?.maStaffId) issues.push(`Provider "${p.name}" missing MA`);
      // XR can be provider-specific OR location-only, so we only warn if neither exists
      const locXR = payload.xrLocationAssignments?.[a?.locationCode] || "";
      const anyXR = (a?.xrStaffId || locXR);
      if(!anyXR) issues.push(`Provider "${p.name}" missing XR (set per provider or location)`);
    });
  }

  return issues;
}

/* ---------------- Export CSV ---------------- */

function exportPublishedToCSV(){
  const sched = state.published;
  if(!sched?.roster?.length) return toast("No published schedule to export.");

  const header = ["TEAM MEMBER", ...DAYS];
  const lines = [header];

  sched.roster.forEach(r=>{
    const e = sched.entries?.[r.id] || {};
    const teamLabel = r.role === "XR" ? `(XR) ${r.name}` : r.name;
    const row = [teamLabel, ...DAYS.map(d=>{
      const tid = e[d];
      return tid ? templateTextById(sched, tid) : "";
    })];
    lines.push(row);
  });

  const csv = lines.map(arr => arr.map(csvEscape).join(",")).join("\n");
  downloadTextFile(`Schedule_${state.currentWeekOf}.csv`, csv, "text/csv");
  toast("CSV exported.");
}

/* ---------------- CSV Import parsing ---------------- */

function parseScheduleCSV(text){
  const rows = csvToRows(text);

  const headerIdx = rows.findIndex(r => (r[0]||"").trim().toUpperCase() === "TEAM MEMBER");
  if(headerIdx === -1) throw new Error("No TEAM MEMBER header found.");

  const header = rows[headerIdx].map(x => (x||"").trim());
  const dayIdx = {
    TEAM: 0,
    Monday: header.findIndex(x=>x.toLowerCase()==="monday"),
    Tuesday: header.findIndex(x=>x.toLowerCase()==="tuesday"),
    Wednesday: header.findIndex(x=>x.toLowerCase()==="wednesday"),
    Thursday: header.findIndex(x=>x.toLowerCase()==="thursday"),
    Friday: header.findIndex(x=>x.toLowerCase()==="friday"),
  };
  if(Object.values(dayIdx).some(i=>i<0)) throw new Error("Missing weekday headers.");

  // Find first date above header like "03/02/2026- 03/06/2026"
  let weekOf = null;
  for(let i=headerIdx-1; i>=0; i--){
    const s = (rows[i][0]||"").trim();
    const m = s.match(/(\d{2}\/\d{2}\/\d{4})/);
    if(m){ weekOf = toISODate(m[1]); break; }
  }
  if(!weekOf) weekOf = mondayOf(new Date());

  const roster = [];
  const entries = {};
  const templateRawSet = new Set();

  let emptyCount = 0;
  for(let i=headerIdx+1; i<rows.length; i++){
    const teamCell = (rows[i][dayIdx.TEAM]||"").trim();
    const anyDay = DAYS.some(d => ((rows[i][dayIdx[d]]||"").trim() !== ""));
    if(!teamCell && !anyDay){
      emptyCount++;
      if(emptyCount >= 3) break;
      continue;
    }
    emptyCount = 0;
    if(!teamCell) continue;

    const role = teamCell.trim().startsWith("(XR)") ? "XR" : "MA";
    const name = teamCell.replace(/^\(XR\)\s*/i, "").trim();
    const id = slug(name);

    roster.push({ id, name, role, email: "" });
    entries[id] = entries[id] || {};

    DAYS.forEach(d=>{
      const raw = (rows[i][dayIdx[d]]||"").trim();
      if(!raw) { entries[id][d] = null; return; }

      if(raw.toUpperCase() === "OFF"){
        templateRawSet.add("OFF");
        entries[id][d] = idForRaw("OFF");
        return;
      }
      templateRawSet.add(raw);
      entries[id][d] = idForRaw(raw);
    });
  }

  const templates = Array.from(templateRawSet)
    .filter(Boolean)
    .map(raw => normalizeTemplate({ id: idForRaw(raw), raw }));

  defaultTemplates().forEach(t=>{
    if(!templates.some(x=>x.id===t.id)) templates.push(normalizeTemplate(t));
  });

  return {
    weekOf,
    roster,
    templates,
    entries,

    // Provider Builder defaults (kept for week)
    locations: LOCATIONS,
    providers: [],
    providerAssignments: {},
    xrLocationAssignments: {}
  };
}

/* ---------------- Create Blank Draft from defaults (Option C) ---------------- */

async function createBlankDraftFromDefaults(weekOf, seedFromPublished){
  const base = seedFromPublished ? deepClone(seedFromPublished) : null;

  // Roster
  let roster = base?.roster?.length ? base.roster.map(r=>({
    id: r.id, name: r.name, role: r.role, email: r.email || ""
  })) : [];

  if(!roster.length){
    try{
      const defaultRoster = await getDefaultRoster();
      if(defaultRoster?.length){
        roster = defaultRoster.map(r=>({ id:r.id, name:r.name, role:r.role, email:r.email || "" }));
      }
    }catch(e){
      console.warn("default roster load failed", e);
    }
  }

  // Providers
  let providers = base?.providers?.length ? normalizeProviders(base.providers) : [];
  if(!providers.length){
    try{
      const defaultProviders = await getDefaultProviders();
      if(defaultProviders?.length){
        providers = normalizeProviders(defaultProviders);
      }
    }catch(e){
      console.warn("default providers load failed", e);
    }
  }

  // Templates
  const templates = base?.templates?.length
    ? base.templates.map(t=> normalizeTemplate({ id: t.id, raw: t.raw }))
    : defaultTemplates().map(t=> normalizeTemplate(t));

  // Entries
  const entries = {};
  roster.forEach(r=>{
    entries[r.id] = {};
    DAYS.forEach(d => entries[r.id][d] = null);
  });

  const providerAssignments = base?.providerAssignments ? deepClone(base.providerAssignments) : {};
  const xrLocationAssignments = base?.xrLocationAssignments ? deepClone(base.xrLocationAssignments) : {};

  return {
    weekOf,
    roster,
    templates,
    entries,

    // Provider builder state
    locations: LOCATIONS,
    providers,
    providerAssignments,
    xrLocationAssignments
  };
}

/* ---------------- Templates ---------------- */

function defaultTemplates(){
  return [
    { id: idForRaw("OFF"), raw:"OFF" },
    { id: idForRaw("TRAINING 08:00- 04:30"), raw:"TRAINING 08:00- 04:30" },
    { id: idForRaw("FLOAT (Val) 08:00- 04:30"), raw:"FLOAT (Val) 08:00- 04:30" },
  ];
}

function normalizeTemplates(arr){
  return (arr || []).map(t => normalizeTemplate(t));
}

function normalizeTemplate(t){
  const raw = (t.raw || "").trim();
  const parsed = parseCell(raw);
  return { id: t.id || idForRaw(raw), raw, parsed };
}

function parseCell(raw){
  const s = (raw||"").trim();
  if(!s) return { type:"EMPTY", label:"", site:"", start:"", end:"" };
  if(s.toUpperCase() === "OFF") return { type:"OFF", label:"OFF", site:"", start:"", end:"" };

  let label = s;
  let site = "";

  const siteMatch = s.match(/\(([^\)]+)\)/);
  if(siteMatch) site = siteMatch[1].trim();

  const timeMatch = s.match(/(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})(?:\s*(AM|PM))?/i);
  let start = "", end = "";
  if(timeMatch){
    start = (timeMatch[1]||"").trim();
    end = (timeMatch[2]||"").trim();
  }

  if(siteMatch){
    label = s.slice(0, siteMatch.index).trim();
  }else if(timeMatch){
    label = s.slice(0, timeMatch.index).trim();
  }
  if(!label && site) label = site;

  const upper = label.toUpperCase();
  const type =
    upper === "TRAINING" ? "TRAINING" :
    upper.startsWith("FLOAT") ? "FLOAT" :
    upper === "XR" ? "XR" :
    "ASSIGN";

  return { type, label, site, start, end };
}

function templateTextById(sched, id){
  const templates = (sched.templates || []);
  const t = templates.find(x=>x.id===id);
  if(!t) return "";
  const p = t.parsed || parseCell(t.raw);
  if(p.type === "OFF") return "OFF";
  const sitePart = p.site ? ` (${p.site})` : "";
  const timePart = (p.start && p.end) ? ` ${p.start}- ${p.end}` : "";
  return `${p.label}${sitePart}${timePart}`.trim();
}

/* ---------------- Helpers ---------------- */

function csvToRows(text){
  const lines = text.replace(/\r\n/g,"\n").replace(/\r/g,"\n").split("\n");
  return lines.map(parseCSVLine);
}
function parseCSVLine(line){
  const out = [];
  let cur = "";
  let inQ = false;
  for(let i=0; i<line.length; i++){
    const ch = line[i];
    if(ch === '"'){
      if(inQ && line[i+1] === '"'){ cur += '"'; i++; }
      else inQ = !inQ;
      continue;
    }
    if(ch === "," && !inQ){
      out.push(cur);
      cur = "";
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out;
}
function csvEscape(v){
  const s = (v ?? "").toString();
  if(/[,"\n]/.test(s)) return `"${s.replace(/"/g,'""')}"`;
  return s;
}
function toISODate(mmddyyyy){
  const [mm,dd,yyyy] = mmddyyyy.split("/");
  return `${yyyy}-${mm.padStart(2,"0")}-${dd.padStart(2,"0")}`;
}

function mondayOf(date){
  const d = new Date(date);
  const day = d.getDay();
  const diff = (day === 0 ? -6 : 1) - day;
  d.setDate(d.getDate() + diff);
  return d.toISOString().slice(0,10);
}

function dayName(date){
  const idx = date.getDay();
  const map = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  const name = map[idx];
  return DAYS.includes(name) ? name : null;
}

function formatLongDate(d){
  return d.toLocaleDateString(undefined, { weekday:"long", year:"numeric", month:"long", day:"numeric" });
}

function slug(name){
  return (name||"")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g,"-")
    .replace(/^-+|-+$/g,"")
    .slice(0,60) || "staff";
}

function providerId(name){
  return "p_" + slug(name);
}

function idForRaw(raw){
  const s = (raw||"").trim().toLowerCase();
  let h = 2166136261;
  for(let i=0;i<s.length;i++){
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return "t_" + (h>>>0).toString(16);
}

function deepClone(obj){
  return JSON.parse(JSON.stringify(obj));
}

function toast(msg){
  const el = $("#toast");
  el.textContent = msg;
  el.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(()=>{ el.hidden = true; }, 2600);
}

function downloadTextFile(filename, content, mime){
  const blob = new Blob([content], {type: mime || "text/plain"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function escapeHtml(s){
  return (s ?? "").toString()
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

/* ---------------- Modal ---------------- */

function openModal({title, body, foot, onAfter}){
  $("#modalTitle").textContent = title || "Modal";
  $("#modalBody").innerHTML = body || "";
  $("#modalFoot").innerHTML = foot || `<button class="btn btnGhost" data-close="1">Close</button>`;
  $("#modal").hidden = false;
  setTimeout(()=>{ onAfter?.(); }, 0);
}
function closeModal(){
  $("#modal").hidden = true;
  $("#modalBody").innerHTML = "";
  $("#modalFoot").innerHTML = "";
}
