const DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday"];
const $ = (sel) => document.querySelector(sel);

// Location mapping (display + code)
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

  // Builder modes: "provider" | "staff"
  builderMode: "provider",

  // Builder helpers (staff grid)
  lastPickedTemplateId: null,
  builder: {
    active: null,   // { pid, day, row, col }
    range: null     // { r1,c1,r2,c2 }
  }
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

async function fetchPublished(weekOf){
  try{
    const res = await fetch(`./api/schedule?weekOf=${encodeURIComponent(weekOf)}&mode=published`, { cache: "no-store" });
    if(!res.ok) return null;
    const data = await res.json();
    return data.data || null;
  }catch{
    return null;
  }
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
  ["my","everyone","builder"].forEach(v=>{
    $(`#view-${v}`).hidden = (v !== view);
  });
}

/* ---------------- Buttons ---------------- */

function wireButtons(){
  $("#btnPrint").addEventListener("click", ()=> window.print());
  $("#btnExportCSV").addEventListener("click", exportPublishedToCSV);

  $("#btnLogin").addEventListener("click", login);
  $("#btnLogout").addEventListener("click", logout);

  // Builder mode toggles
  $("#btnModeProvider")?.addEventListener("click", ()=>{
    state.builderMode = "provider";
    renderBuilder();
  });
  $("#btnModeStaff")?.addEventListener("click", ()=>{
    state.builderMode = "staff";
    renderBuilder();
  });

  $("#btnManageProviders")?.addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;
    openProvidersModal(payload);
  });

  // CSV import remains (fallback)
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

      toast("CSV imported into Draft.");
      $("#csvFile").value = "";
      clearSelection();
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
    clearSelection();
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

  // Workflow buttons
  $("#btnNewBlankWeek")?.addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    state.draft = createBlankDraft(state.currentWeekOf, state.published || null);
    toast("Blank draft created.");
    clearSelection();
    renderBuilder();
  });

  $("#btnStartFromPublished")?.addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const pub = state.published && state.published.weekOf === state.currentWeekOf ? state.published : null;
    if(!pub) return toast("No published schedule to start from.");
    state.draft = deepClone(pub);
    state.draft.weekOf = state.currentWeekOf;
    // ensure providerAssignments exists (optional)
    if(!state.draft.providerAssignments) state.draft.providerAssignments = {};
    toast("Draft created from Published.");
    clearSelection();
    renderBuilder();
  });

  $("#btnClonePrevPublished")?.addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");

    const idx = state.weeks.indexOf(state.currentWeekOf);
    const prevWeek = (idx >= 0) ? state.weeks[idx + 1] : null;
    if(!prevWeek) return toast("No previous week found to clone.");

    const prevPub = await fetchPublished(prevWeek);
    if(!prevPub) return toast(`No published schedule for ${prevWeek}.`);

    const cloned = deepClone(prevPub);
    cloned.weekOf = state.currentWeekOf;

    // keep emails from current if we have them
    const currentEmailMap = rosterEmailMap(state.published || state.draft);
    cloned.roster.forEach(r=>{
      const email = currentEmailMap.get(r.id);
      if(email) r.email = email;
    });

    if(!cloned.providerAssignments) cloned.providerAssignments = {};

    state.draft = cloned;
    toast(`Cloned ${prevWeek} (published) → ${state.currentWeekOf} (draft).`);
    clearSelection();
    renderBuilder();
  });

  // Bulk tools (staff grid)
  $("#btnApplySelection")?.addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;
    const tid = $("#bulkTemplate")?.value || "";
    if(!tid) return toast("Pick a bulk template first.");
    if(!state.builder.range) return toast("No selection. Shift+Click to select a range.");
    applyTemplateToRange(payload, tid, state.builder.range);
    state.lastPickedTemplateId = tid;
    toast("Applied to selection.");
    renderBuilderGrid(payload);
    updateSelectionPill(payload);
  });

  $("#btnApplyRow")?.addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;
    const tid = $("#bulkTemplate")?.value || "";
    if(!tid) return toast("Pick a bulk template first.");
    if(!state.builder.active) return toast("Click a cell in the row first.");
    const r = state.builder.active.row;
    applyTemplateToRange(payload, tid, { r1:r, r2:r, c1:0, c2:DAYS.length-1 });
    state.lastPickedTemplateId = tid;
    toast("Applied to row.");
    renderBuilderGrid(payload);
    updateSelectionPill(payload);
  });

  $("#btnApplyDay")?.addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;
    const tid = $("#bulkTemplate")?.value || "";
    if(!tid) return toast("Pick a bulk template first.");
    if(!state.builder.active) return toast("Click a cell in the day column first.");
    const c = state.builder.active.col;
    applyTemplateToRange(payload, tid, { r1:0, r2:payload.roster.length-1, c1:c, c2:c });
    state.lastPickedTemplateId = tid;
    toast("Applied to day.");
    renderBuilderGrid(payload);
    updateSelectionPill(payload);
  });

  $("#btnClearSelection")?.addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;
    if(!state.builder.range) return toast("No selection to clear.");
    applyTemplateToRange(payload, null, state.builder.range);
    toast("Cleared selection.");
    renderBuilderGrid(payload);
    updateSelectionPill(payload);
  });

  $("#providerFilter")?.addEventListener("input", ()=>{
    renderProviderGrid(ensureDraftWeek(false));
  });

  // Everyone filters
  $("#roleFilter").addEventListener("change", renderEveryone);
  $("#searchFilter").addEventListener("input", renderEveryone);

  // Modal close
  $("#modal").addEventListener("click", (e)=>{
    if(e.target?.dataset?.close) closeModal();
  });

  // Keyboard: Enter repeats last template on active staff-grid cell
  document.addEventListener("keydown", (e)=>{
    if(state.view !== "builder") return;
    if(!state.auth.editor) return;
    if(state.builderMode !== "staff") return;
    if(e.key !== "Enter") return;

    const payload = ensureDraftWeek(false);
    if(!payload) return;
    if(!state.builder.active) return;
    if(!state.lastPickedTemplateId) return;

    e.preventDefault();
    const { pid, day } = state.builder.active;
    if(!payload.entries[pid]) payload.entries[pid] = {};
    payload.entries[pid][day] = state.lastPickedTemplateId;
    toast("Applied last template.");
    renderBuilderGrid(payload);
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
  clearSelection();
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

/* ---------------- My Schedule ---------------- */

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

/* ---------------- Everyone ---------------- */

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

/* ---------------- Builder ---------------- */

function renderBuilder(){
  if(!state.auth.editor){
    $("#builderStatus").textContent = "Editors only";
    return;
  }

  // show correct inner mode
  $("#builderProviderWrap").hidden = (state.builderMode !== "provider");
  $("#builderStaffWrap").hidden = (state.builderMode !== "staff");

  const week = state.currentWeekOf;
  const draft = state.draft && state.draft.weekOf === week ? state.draft : null;
  const pub = state.published && state.published.weekOf === week ? state.published : null;

  $("#builderStatus").textContent =
    draft ? "Draft loaded (editable)" :
    pub ? "No draft loaded (use Start From Published or Clone Previous)" :
    "No data loaded for this week (use New Blank)";

  const payload = ensureDraftWeek(false);
  if(!payload){
    renderEmptyBuilder();
    return;
  }

  // Normalize and ensure defaults exist
  payload.templates = normalizeTemplates(payload.templates);
  defaultTemplates().forEach(t=>{
    if(!payload.templates.some(x=>x.id===t.id)) payload.templates.push(normalizeTemplate(t));
  });

  // Ensure provider assignments exists
  if(!payload.providerAssignments) payload.providerAssignments = {};
  if(!payload.providers) payload.providers = deriveProviders(payload);

  renderRosterEmails(payload);
  renderTemplateList(payload);

  if(state.builderMode === "provider"){
    renderProviderGrid(payload);
  }else{
    renderBulkTemplateSelect(payload);
    updateSelectionPill(payload);
    renderBuilderGrid(payload);
  }
}

function renderEmptyBuilder(){
  $("#rosterEmailList").innerHTML = `<div class="muted">Load/create a draft to edit.</div>`;
  $("#templateList").innerHTML = `<div class="muted">—</div>`;
  $("#builderGrid").querySelector("thead").innerHTML = "";
  $("#builderGrid").querySelector("tbody").innerHTML = `<tr><td class="muted" style="padding:14px">—</td></tr>`;

  $("#providerGrid").querySelector("thead").innerHTML = "";
  $("#providerGrid").querySelector("tbody").innerHTML = `<tr><td class="muted" style="padding:14px">—</td></tr>`;
}

function ensureDraftWeek(showToastOnFail=true){
  if(!state.draft){
    if(showToastOnFail) toast("No draft loaded. Use New Blank, Start From Published, or Clone Previous.");
    return null;
  }
  if(!state.draft.weekOf) state.draft.weekOf = state.currentWeekOf;
  if(!state.draft.templates) state.draft.templates = [];
  if(!state.draft.entries) state.draft.entries = {};
  if(!state.draft.roster) state.draft.roster = [];
  return state.draft;
}

/* ---------------- Provider Builder ---------------- */

function deriveProviders(payload){
  // Pull unique "providers" from templates that look like "Provider (LOC) HH:MM- HH:MM"
  // and from existing providerAssignments keys
  const set = new Set();

  Object.keys(payload.providerAssignments || {}).forEach(p => set.add(p));

  (payload.templates || []).forEach(t=>{
    const p = parseProviderish(t.raw);
    if(p?.provider && p.provider.toUpperCase() !== "XR" && p.provider.toUpperCase() !== "OFF"){
      set.add(p.provider);
    }
  });

  return Array.from(set).sort((a,b)=>a.localeCompare(b));
}

function renderProviderGrid(payload){
  const grid = $("#providerGrid");
  const thead = grid.querySelector("thead");
  const tbody = grid.querySelector("tbody");
  const q = ($("#providerFilter")?.value || "").trim().toLowerCase();

  const providers = (payload.providers || []).filter(p => !q || p.toLowerCase().includes(q));

  thead.innerHTML = `
    <tr>
      <th style="min-width:220px;text-align:left;">Provider</th>
      ${DAYS.map(d=>`<th style="min-width:220px;text-align:left;">${d}</th>`).join("")}
    </tr>
  `;

  if(!providers.length){
    tbody.innerHTML = `<tr><td class="muted" style="padding:14px">No providers. Click Manage → Add providers.</td></tr>`;
    return;
  }

  tbody.innerHTML = providers.map(provider=>{
    return `
      <tr>
        <td class="nameCell">${escapeHtml(provider)}</td>
        ${DAYS.map(day=>{
          const cell = getProviderCell(payload, provider, day);
          const maTxt = cell?.ma ? formatAssignSummary("MA", payload, cell.ma) : "MA: —";
          const xrTxt = cell?.xr ? formatAssignSummary("XR", payload, cell.xr) : "XR: —";
          return `
            <td>
              <button class="cellBtn" data-prov="${escapeHtmlAttr(provider)}" data-day="${day}" style="width:100%; text-align:left;">
                <div style="font-weight:800;">${escapeHtml(day)}</div>
                <div class="muted small" style="margin-top:4px;">${escapeHtml(maTxt)}</div>
                <div class="muted small">${escapeHtml(xrTxt)}</div>
              </button>
            </td>
          `;
        }).join("")}
      </tr>
    `;
  }).join("");

  tbody.querySelectorAll(".cellBtn").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      const provider = btn.dataset.prov;
      const day = btn.dataset.day;
      openProviderDayModal(payload, provider, day);
    });
  });
}

function getProviderCell(payload, provider, day){
  const pa = payload.providerAssignments || {};
  const prov = pa[provider] || {};
  return prov[day] || null;
}

function setProviderCell(payload, provider, day, value){
  if(!payload.providerAssignments) payload.providerAssignments = {};
  if(!payload.providerAssignments[provider]) payload.providerAssignments[provider] = {};
  payload.providerAssignments[provider][day] = value;
}

function openProviderDayModal(payload, provider, day){
  const existing = getProviderCell(payload, provider, day) || {};
  const ma = existing.ma || {};
  const xr = existing.xr || {};

  const maStaffOpts = rosterOptions(payload, "MA", ma.staffId || "");
  const xrStaffOpts = rosterOptions(payload, "XR", xr.staffId || "");

  const locOpts = locationOptions(ma.location || "");
  const locOptsXR = locationOptions(xr.location || "");

  const xrLocOnly = !!xr.locationOnly;

  openModal({
    title: `${provider} • ${day}`,
    body: `
      <div class="card" style="border:1px solid rgba(15,23,42,.12); padding:12px; border-radius:12px;">
        <div style="font-weight:900; margin-bottom:8px;">MA Assignment</div>

        <div class="row">
          <div class="field grow">
            <label>MA Staff</label>
            <select id="pa_ma_staff">${maStaffOpts}</select>
          </div>
          <div class="field">
            <label>Location</label>
            <select id="pa_ma_loc">${locOpts}</select>
          </div>
        </div>

        <div class="row">
          <div class="field">
            <label>Start</label>
            <input id="pa_ma_start" type="time" value="${escapeHtml(ma.start || "")}">
          </div>
          <div class="field">
            <label>End</label>
            <input id="pa_ma_end" type="time" value="${escapeHtml(ma.end || "")}">
          </div>
        </div>

        <div class="row">
          <button class="btn btnGhost" id="pa_ma_clear">Clear MA</button>
        </div>
      </div>

      <div style="height:10px"></div>

      <div class="card" style="border:1px solid rgba(15,23,42,.12); padding:12px; border-radius:12px;">
        <div style="font-weight:900; margin-bottom:8px;">XR Assignment</div>

        <div class="row">
          <div class="field grow">
            <label>XR Staff (optional)</label>
            <select id="pa_xr_staff">${xrStaffOpts}</select>
            <div class="hint">If blank + location-only, it won’t assign any staff cell.</div>
          </div>
          <div class="field">
            <label>Location</label>
            <select id="pa_xr_loc">${locOptsXR}</select>
          </div>
        </div>

        <div class="row" style="align-items:center;">
          <div class="field grow" style="margin:0;">
            <label style="display:flex; gap:10px; align-items:center;">
              <input id="pa_xr_loconly" type="checkbox" ${xrLocOnly ? "checked" : ""}>
              Location-only XR (no provider name)
            </label>
          </div>
        </div>

        <div class="row">
          <div class="field">
            <label>Start</label>
            <input id="pa_xr_start" type="time" value="${escapeHtml(xr.start || "")}">
          </div>
          <div class="field">
            <label>End</label>
            <input id="pa_xr_end" type="time" value="${escapeHtml(xr.end || "")}">
          </div>
        </div>

        <div class="row">
          <button class="btn btnGhost" id="pa_xr_clear">Clear XR</button>
        </div>
      </div>
    `,
    foot: `
      <button class="btn btnGhost" data-close="1">Cancel</button>
      <button class="btn" id="pa_save">Save</button>
    `,
    onAfter(){
      $("#pa_ma_clear").addEventListener("click", ()=>{
        $("#pa_ma_staff").value = "";
        $("#pa_ma_loc").value = "";
        $("#pa_ma_start").value = "";
        $("#pa_ma_end").value = "";
      });

      $("#pa_xr_clear").addEventListener("click", ()=>{
        $("#pa_xr_staff").value = "";
        $("#pa_xr_loc").value = "";
        $("#pa_xr_start").value = "";
        $("#pa_xr_end").value = "";
        $("#pa_xr_loconly").checked = false;
      });

      $("#pa_save").addEventListener("click", ()=>{
        const maStaff = ($("#pa_ma_staff").value || "").trim();
        const maLoc = ($("#pa_ma_loc").value || "").trim();
        const maStart = ($("#pa_ma_start").value || "").trim();
        const maEnd = ($("#pa_ma_end").value || "").trim();

        const xrStaff = ($("#pa_xr_staff").value || "").trim();
        const xrLoc = ($("#pa_xr_loc").value || "").trim();
        const xrStart = ($("#pa_xr_start").value || "").trim();
        const xrEnd = ($("#pa_xr_end").value || "").trim();
        const xrLocOnly = !!$("#pa_xr_loconly").checked;

        const cell = { };

        // MA: require staff + loc + time (since it varies)
        if(maStaff && maLoc && maStart && maEnd){
          cell.ma = { staffId: maStaff, location: maLoc, start: maStart, end: maEnd };
        }

        // XR: allow staff optional, but require loc + time to record something useful
        if(xrLoc && xrStart && xrEnd){
          cell.xr = { staffId: xrStaff || "", location: xrLoc, start: xrStart, end: xrEnd, locationOnly: xrLocOnly };
        }

        setProviderCell(payload, provider, day, cell);

        // Now project provider assignments into staff grid
        syncProviderDayToStaff(payload, provider, day);

        toast("Saved provider assignment.");
        closeModal();
        renderProviderGrid(payload);
        // keep staff grid in sync if user switches modes
        renderTemplateList(payload);
      });
    }
  });
}

function rosterOptions(payload, role, selectedId){
  const people = (payload.roster || []).filter(r => r.role === role);
  const opts = [
    `<option value="">— None —</option>`,
    ...people.map(p=> `<option value="${escapeHtmlAttr(p.id)}" ${p.id===selectedId?"selected":""}>${escapeHtml(p.name)}</option>`)
  ];
  return opts.join("");
}

function locationOptions(selected){
  const opts = [
    `<option value="">— Select —</option>`,
    ...LOCATIONS.map(l => `<option value="${escapeHtmlAttr(l.code)}" ${l.code===selected?"selected":""}>${escapeHtml(l.name)} (${escapeHtml(l.code)})</option>`)
  ];
  return opts.join("");
}

function formatAssignSummary(role, payload, a){
  if(!a) return `${role}: —`;
  const person = (payload.roster || []).find(r => r.id === a.staffId);
  const who = person ? person.name : (a.staffId ? a.staffId : "—");
  const locName = LOCATIONS.find(x=>x.code===a.location)?.name || a.location || "—";
  const time = (a.start && a.end) ? `${a.start}-${a.end}` : "";
  if(role === "XR" && a.locationOnly){
    return `XR: ${who || "—"} • ${locName} • ${time}`.trim();
  }
  return `${role}: ${who || "—"} • ${locName} • ${time}`.trim();
}

/**
 * syncProviderDayToStaff
 * Writes the MA/XR assignment into staff grid entries by generating/using templates.
 */
function syncProviderDayToStaff(payload, provider, day){
  // Clear any existing staff entries caused by old provider/day assignments:
  // We do it safely by removing entries that match templates generated for this provider/day.
  // (This keeps it predictable: provider cell is the truth.)
  const pa = getProviderCell(payload, provider, day) || {};
  const oldKeys = possibleProviderTemplateKeys(provider); // used to detect
  clearEntriesMatchingProvider(payload, day, oldKeys);

  // Apply MA
  if(pa.ma){
    const raw = `${provider} (${pa.ma.location}) ${fmtTime(pa.ma.start)}- ${fmtTime(pa.ma.end)}`;
    const tid = ensureTemplate(payload, raw);
    ensureEntry(payload, pa.ma.staffId, day, tid);
  }

  // Apply XR
  if(pa.xr){
    // If no staff selected, we don't write a staff-grid assignment (still recorded in providerAssignments)
    if(pa.xr.staffId){
      const label = pa.xr.locationOnly ? `XR (${pa.xr.location}) ${fmtTime(pa.xr.start)}- ${fmtTime(pa.xr.end)}`
        : `XR ${provider} (${pa.xr.location}) ${fmtTime(pa.xr.start)}- ${fmtTime(pa.xr.end)}`;
      const tid = ensureTemplate(payload, label);
      ensureEntry(payload, pa.xr.staffId, day, tid);
    }
  }
}

function possibleProviderTemplateKeys(provider){
  // basic fragments we can detect in template raw
  // we look for provider name, and for XR provider variant
  const p = provider.toLowerCase();
  return [
    `${p} (`,
    `xr ${p} (`
  ];
}

function clearEntriesMatchingProvider(payload, day, providerKeys){
  const templatesById = new Map((payload.templates||[]).map(t => [t.id, t.raw || ""]));
  (payload.roster||[]).forEach(r=>{
    const e = payload.entries?.[r.id];
    if(!e) return;
    const tid = e[day];
    if(!tid) return;
    const raw = (templatesById.get(tid) || "").toLowerCase();
    if(providerKeys.some(k => raw.includes(k))){
      e[day] = null;
    }
  });
}

function ensureEntry(payload, staffId, day, templateId){
  if(!payload.entries[staffId]) payload.entries[staffId] = {};
  payload.entries[staffId][day] = templateId;
}

function ensureTemplate(payload, raw){
  const id = idForRaw(raw);
  if(!payload.templates.some(t=>t.id===id)){
    payload.templates.push(normalizeTemplate({ id, raw }));
  }
  return id;
}

function fmtTime(t){
  // time input returns HH:MM (24-hour). Keep it as is.
  return t || "";
}

function openProvidersModal(payload){
  const current = (payload.providers || []).slice().sort((a,b)=>a.localeCompare(b));
  openModal({
    title: "Manage Providers",
    body: `
      <div class="field" style="min-width:100%">
        <label>Add provider</label>
        <div style="display:flex; gap:8px;">
          <input id="provNew" type="text" placeholder="Dr. Lastname" />
          <button class="btn" id="provAdd">Add</button>
        </div>
        <div class="hint">Providers drive the Provider Builder grid.</div>
      </div>

      <div style="margin-top:12px;">
        ${current.length ? current.map(p=>`
          <div class="templateItem">
            <div class="templateTop">
              <div>
                <div class="templateName">${escapeHtml(p)}</div>
                <div class="templateMeta muted">Provider</div>
              </div>
              <div>
                <button class="iconBtn" title="Remove" data-prov-del="${escapeHtmlAttr(p)}">🗑</button>
              </div>
            </div>
          </div>
        `).join("") : `<div class="muted">No providers yet.</div>`}
      </div>
    `,
    foot: `<button class="btn btnGhost" data-close="1">Close</button>`,
    onAfter(){
      $("#provAdd").addEventListener("click", ()=>{
        const p = ($("#provNew").value || "").trim();
        if(!p) return toast("Enter a provider name.");
        if(!payload.providers) payload.providers = [];
        if(payload.providers.includes(p)) return toast("Already exists.");
        payload.providers.push(p);
        toast("Provider added.");
        closeModal();
        renderBuilder();
      });

      document.querySelectorAll("[data-prov-del]").forEach(btn=>{
        btn.addEventListener("click", ()=>{
          const p = btn.dataset.provDel;
          payload.providers = (payload.providers || []).filter(x => x !== p);
          // keep providerAssignments (optional) but you can remove if you want
          toast("Provider removed.");
          closeModal();
          renderBuilder();
        });
      });
    }
  });
}

/* ---------------- Staff Grid: bulk + picker ---------------- */

function renderBulkTemplateSelect(payload){
  const sel = $("#bulkTemplate");
  const templates = payload.templates || [];
  const prev = sel.value || state.lastPickedTemplateId || "";

  const opts = [
    `<option value="">(pick template)</option>`,
    ...templates.map(t=>{
      const label = templateTextById(payload, t.id) || t.raw;
      return `<option value="${t.id}">${escapeHtml(label)}</option>`;
    })
  ].join("");

  sel.innerHTML = opts;

  if(prev && templates.some(t=>t.id===prev)){
    sel.value = prev;
  }
}

function updateSelectionPill(payload){
  const pill = $("#selectionPill");
  if(!state.builder.range){
    pill.textContent = state.builder.active ? `${state.builder.active.pid} • ${state.builder.active.day}` : "None";
    return;
  }
  const { r1,c1,r2,c2 } = normalizeRange(state.builder.range);
  const rows = payload?.roster?.length || 0;
  const cols = DAYS.length;
  const rr1 = clamp(r1, 0, Math.max(0, rows-1));
  const rr2 = clamp(r2, 0, Math.max(0, rows-1));
  const cc1 = clamp(c1, 0, cols-1);
  const cc2 = clamp(c2, 0, cols-1);
  const cellCount = (Math.abs(rr2-rr1)+1) * (Math.abs(cc2-cc1)+1);
  pill.textContent = `${cellCount} cells`;
}

function applyTemplateToRange(payload, templateIdOrNull, range){
  const { r1,c1,r2,c2 } = normalizeRange(range);
  const roster = payload.roster || [];
  const rows = roster.length;

  for(let r = r1; r <= r2; r++){
    if(r < 0 || r >= rows) continue;
    const pid = roster[r].id;
    if(!payload.entries[pid]) payload.entries[pid] = {};
    for(let c = c1; c <= c2; c++){
      if(c < 0 || c >= DAYS.length) continue;
      const day = DAYS[c];
      payload.entries[pid][day] = templateIdOrNull ? templateIdOrNull : null;
    }
  }
}

function clearSelection(){
  state.builder.active = null;
  state.builder.range = null;
}

/* ---------------- Builder: roster + templates + staff grid ---------------- */

function renderRosterEmails(payload){
  const wrap = $("#rosterEmailList");
  if(!payload.roster.length){
    wrap.innerHTML = `<div class="muted">No roster loaded.</div>`;
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
      : `${(p.site ? p.site : "—")} • ${(p.start||"")}–${(p.end||"")}`;

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

  thead.innerHTML = `
    <tr>
      <th style="min-width:240px;text-align:left;">Team Member</th>
      ${DAYS.map(d=>`<th style="min-width:210px;text-align:left;">${d}</th>`).join("")}
    </tr>
  `;

  tbody.innerHTML = payload.roster.map((r,rowIdx)=>{
    const e = payload.entries?.[r.id] || {};
    return `
      <tr>
        <td class="nameCell">
          ${escapeHtml(r.name)} <span class="roleTag">${r.role}</span>
        </td>
        ${DAYS.map((d,colIdx)=>{
          const tid = e[d];
          const txt = tid ? templateTextById(payload, tid) : "";
          const isActive = state.builder.active && state.builder.active.pid===r.id && state.builder.active.day===d;
          const isInRange = cellIsInRange(rowIdx, colIdx, state.builder.range);

          const cls = [
            "cellBtn",
            isActive ? "is-active" : "",
            isInRange ? "is-selected" : ""
          ].filter(Boolean).join(" ");

          return `
            <td>
              <button class="${cls}" data-pid="${r.id}" data-day="${d}" data-row="${rowIdx}" data-col="${colIdx}">
                ${escapeHtml(txt || "—")}
              </button>
            </td>
          `;
        }).join("")}
      </tr>
    `;
  }).join("");

  tbody.querySelectorAll(".cellBtn").forEach(btn=>{
    btn.addEventListener("click", (ev)=>{
      const pid = btn.dataset.pid;
      const day = btn.dataset.day;
      const row = parseInt(btn.dataset.row, 10);
      const col = parseInt(btn.dataset.col, 10);

      state.builder.active = { pid, day, row, col };

      if(ev.shiftKey){
        if(!state.builder.range){
          state.builder.range = { r1: row, c1: col, r2: row, c2: col };
        }else{
          state.builder.range.r2 = row;
          state.builder.range.c2 = col;
        }
        updateSelectionPill(payload);
        renderBuilderGrid(payload);
        return;
      }

      state.builder.range = null;
      updateSelectionPill(payload);
      renderBuilderGrid(payload);
      openCellPicker(payload, pid, day);
    });
  });
}

function openCellPicker(payload, personId, day){
  const person = payload.roster.find(r=>r.id===personId);
  payload.templates = normalizeTemplates(payload.templates);

  const currentId = payload.entries?.[personId]?.[day] || "";
  const templates = payload.templates.slice();

  const listHtml = templates.map(t=>{
    const label = templateTextById(payload, t.id) || t.raw;
    const is = t.id === currentId;
    return `
      <button class="btn btnGhost" style="width:100%;justify-content:flex-start;margin:6px 0;${is ? "border-color: var(--navy);" : ""}"
        data-tpl="${t.id}">
        ${escapeHtml(label)}
      </button>
    `;
  }).join("");

  openModal({
    title: `Assign: ${person?.name || "Staff"} • ${day}`,
    body: `
      <div class="field" style="min-width:100%">
        <label>Search templates</label>
        <input id="tplSearch" type="text" placeholder="Type: XR, OFF, Provider…" />
      </div>
      <div id="tplList" style="margin-top:10px; max-height: 46vh; overflow:auto;">
        <button class="btn btnGhost" style="width:100%;justify-content:flex-start;margin:6px 0;" data-tpl="">— Clear —</button>
        ${listHtml}
      </div>
    `,
    foot: `<button class="btn btnGhost" data-close="1">Close</button>`,
    onAfter(){
      const search = $("#tplSearch");
      const list = $("#tplList");

      const bindButtons = ()=>{
        list.querySelectorAll("[data-tpl]").forEach(b=>{
          b.addEventListener("click", ()=>{
            const tid = b.dataset.tpl || null;
            if(!payload.entries[personId]) payload.entries[personId] = {};
            payload.entries[personId][day] = tid;

            if(tid){
              state.lastPickedTemplateId = tid;
              $("#bulkTemplate").value = tid;
            }

            closeModal();
            renderBuilderGrid(payload);
            updateSelectionPill(payload);
          });
        });
      };

      const filterList = ()=>{
        const q = (search.value || "").trim().toLowerCase();
        const buttons = Array.from(list.querySelectorAll("[data-tpl]"));
        buttons.forEach(b=>{
          const tid = b.dataset.tpl || "";
          if(!tid){ b.style.display = ""; return; }
          const label = (templateTextById(payload, tid) || "").toLowerCase();
          b.style.display = !q || label.includes(q) ? "" : "none";
        });
      };

      search.addEventListener("input", filterList);
      bindButtons();
      search.focus();
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
        <input id="tplRaw" type="text" placeholder='Example: Dworsky (SFS) 07:50- 04:20' />
        <div class="hint">OFF • TRAINING 08:00- 04:30 • XR Pelton (WIL) 08:00- 04:30</div>
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
        toast("Template added.");
        closeModal();
        renderBuilder();
      });
    }
  });
}

/* ---------------- Validation ---------------- */

function validate(payload){
  const issues = [];

  payload.roster.forEach(r=>{
    if(!r.email) issues.push(`${r.name} missing email`);
  });

  payload.roster.forEach(r=>{
    const e = payload.entries?.[r.id] || {};
    DAYS.forEach(d=>{
      const v = e[d];
      if(v === undefined || v === null || v === ""){
        issues.push(`${r.name} missing ${d}`);
      }
    });
  });

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

  // optional initial providers derived from templates
  const providerAssignments = {};
  const providers = [];

  return { weekOf, roster, templates, entries, providerAssignments, providers };
}

/* ---------------- Draft creation helpers ---------------- */

function createBlankDraft(weekOf, seedFromPublished){
  const base = seedFromPublished ? deepClone(seedFromPublished) : null;

  const roster = base?.roster?.length ? base.roster.map(r=>({
    id: r.id, name: r.name, role: r.role, email: r.email || ""
  })) : [];

  const templates = base?.templates?.length
    ? base.templates.map(t=> normalizeTemplate({ id: t.id, raw: t.raw }))
    : defaultTemplates().map(t=> normalizeTemplate(t));

  const entries = {};
  roster.forEach(r=>{
    entries[r.id] = {};
    DAYS.forEach(d => entries[r.id][d] = null);
  });

  const providerAssignments = base?.providerAssignments ? deepClone(base.providerAssignments) : {};
  const providers = base?.providers?.length ? base.providers.slice() : [];

  return { weekOf, roster, templates, entries, providerAssignments, providers };
}

function rosterEmailMap(sched){
  const m = new Map();
  if(!sched?.roster?.length) return m;
  sched.roster.forEach(r=>{
    if(r?.id && r?.email) m.set(r.id, r.email);
  });
  return m;
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
    upper.startsWith("XR") ? "XR" :
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

// Tries to interpret "Provider (LOC) HH:MM- HH:MM" and "XR Provider (LOC) ..."
function parseProviderish(raw){
  const s = (raw||"").trim();
  if(!s) return null;
  const m = s.match(/^(.+?)\s*\(([A-Za-z0-9\- ]+)\)\s+(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})/);
  if(!m) return null;

  let label = m[1].trim();
  const loc = (m[2]||"").trim();
  const start = (m[3]||"").trim();
  const end = (m[4]||"").trim();

  // normalize "XR Provider"
  if(label.toUpperCase().startsWith("XR ")){
    label = label.slice(3).trim();
    return { provider: label || null, xr: true, loc, start, end };
  }
  if(label.toUpperCase() === "XR"){
    return { provider: null, xr: true, loc, start, end };
  }
  return { provider: label, xr: false, loc, start, end };
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

function escapeHtmlAttr(s){
  return escapeHtml(s).replaceAll('"', "&quot;");
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

/* ---------------- Range helpers ---------------- */

function normalizeRange(range){
  const r1 = Math.min(range.r1, range.r2);
  const r2 = Math.max(range.r1, range.r2);
  const c1 = Math.min(range.c1, range.c2);
  const c2 = Math.max(range.c1, range.c2);
  return { r1,c1,r2,c2 };
}

function clamp(n, a, b){
  return Math.max(a, Math.min(b, n));
}

function cellIsInRange(r, c, range){
  if(!range) return false;
  const x = normalizeRange(range);
  return r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2;
}
