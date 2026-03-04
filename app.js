const DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday"];
const DEFAULTS_PK = "__defaults__";
const DEFAULT_ROSTER_RK = "roster";
const DEFAULT_PROVIDERS_RK = "providers";
const DEFAULT_PRESETS_RK = "timepresets";

const LOCATIONS = [
  { code:"WIL", name:"Wilshire" },
  { code:"SFS", name:"Santa Fe Springs" },
  { code:"ELA", name:"East LA" },
  { code:"GLE", name:"Glendale" },
  { code:"TZ",  name:"Tarzana" },
  { code:"ENC", name:"Encino" },
  { code:"VAL", name:"Valencia" },
  { code:"TO",  name:"Thousand Oaks" },
];

const DEFAULT_TIME_PRESETS = [
  { id:"t1", label:"7:50am - 4:20pm", start:"07:50", end:"16:20" },
  { id:"t2", label:"8:00am - 4:30pm", start:"08:00", end:"16:30" },
  { id:"t3", label:"8:30am - 5:00pm", start:"08:30", end:"17:00" },
  { id:"t4", label:"8:00am - 5:00pm", start:"08:00", end:"17:00" },
];

const XR_ROOMS = [
  { id:"R1", label:"XR Room 1" },
  { id:"R2", label:"XR Room 2" },
  { id:"R3", label:"XR Room 3" },
];

const $ = (sel) => document.querySelector(sel);

const state = {
  auth: { email: null, editor: false },
  view: "my",

  weeks: [],
  currentWeekOf: null,

  published: null,
  draft: null,

  // Builder selection
  activeProviderId: null,
  providerSearch: "",
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
    state.published = data.data || null;
  }catch{
    state.published = null;
  }
}

async function loadDraft(weekOf){
  try{
    const res = await fetch(`./api/schedule?weekOf=${encodeURIComponent(weekOf)}&mode=draft`, { cache: "no-store" });
    if(!res.ok){ state.draft = null; return; }
    const data = await res.json();
    state.draft = data.data || null;
  }catch{
    state.draft = null;
  }
}

async function loadDefaults(kind){
  const rk =
    kind === "roster" ? DEFAULT_ROSTER_RK :
    kind === "providers" ? DEFAULT_PROVIDERS_RK :
    kind === "timepresets" ? DEFAULT_PRESETS_RK :
    null;
  if(!rk) return null;

  try{
    const res = await fetch(`./api/schedule?weekOf=${encodeURIComponent(DEFAULTS_PK)}&mode=${encodeURIComponent(rk)}`, { cache: "no-store" });
    if(!res.ok) return null;
    const data = await res.json();
    return data.data || null;
  }catch{
    return null;
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

/* ---------------- Week pickers ---------------- */

function renderWeekPickers(){
  const weeks = state.weeks.length ? state.weeks : [state.currentWeekOf];
  const opts = weeks.map(w => `<option value="${w}">${w}</option>`).join("");

  $("#weekPickMy").innerHTML = opts;
  $("#weekPickEveryone").innerHTML = opts;
  $("#weekPickBuilder").innerHTML = opts;

  $("#weekPickMy").value = state.currentWeekOf;
  $("#weekPickEveryone").value = state.currentWeekOf;
  $("#weekPickBuilder").value = state.currentWeekOf;

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

  // keep draft week aligned if already loaded
  if(state.draft && state.draft.weekOf !== week){
    state.draft = null;
    state.activeProviderId = null;
  }

  renderAll();
}

/* ---------------- Buttons ---------------- */

function wireButtons(){
  $("#btnPrint").addEventListener("click", ()=> window.print());
  $("#btnExportCSV").addEventListener("click", exportPublishedToCSV);

  $("#btnLogin").addEventListener("click", login);
  $("#btnLogout").addEventListener("click", logout);

  // Staff import
  $("#btnImportStaffCSV").addEventListener("click", ()=> $("#staffFile").click());
  $("#staffFile").addEventListener("change", async (e)=>{
    const f = e.target.files?.[0];
    if(!f) return;
    const text = await f.text();
    try{
      const staff = parseStaffCSV(text);
      const payload = ensureDraftForWeek(state.currentWeekOf);
      mergeStaffIntoRoster(payload, staff, { defaultRole:"MA" });
      toast(`Staff imported (${staff.length}).`);
      $("#staffFile").value = "";
      renderBuilder();
      renderMy();
      renderEveryone();
    }catch(err){
      console.error(err);
      toast("Staff import failed.");
    }
  });

  $("#btnLoadStaffFromRepo").addEventListener("click", async ()=>{
    try{
      const res = await fetch("./data/staff.csv", { cache:"no-store" });
      if(!res.ok) throw new Error("Missing ./data/staff.csv");
      const text = await res.text();
      const staff = parseStaffCSV(text);
      const payload = ensureDraftForWeek(state.currentWeekOf);
      mergeStaffIntoRoster(payload, staff, { defaultRole:"MA" });
      toast(`Staff loaded from repo (${staff.length}).`);
      renderBuilder();
      renderMy();
      renderEveryone();
    }catch(e){
      console.error(e);
      toast("Load staff from repo failed. Put file at /data/staff.csv");
    }
  });

  $("#btnAddStaff").addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftForWeek(state.currentWeekOf);
    openAddStaffModal(payload);
  });

  // Providers import
  $("#btnImportProvidersCSV").addEventListener("click", ()=> $("#providersFile").click());
  $("#providersFile").addEventListener("change", async (e)=>{
    const f = e.target.files?.[0];
    if(!f) return;
    const text = await f.text();
    try{
      const providers = parseProvidersCSV(text);
      const payload = ensureDraftForWeek(state.currentWeekOf);
      payload.providers = mergeProviders(payload.providers || [], providers);
      ensureProviderAssignmentsScaffold(payload);
      toast(`Providers imported (${providers.length}).`);
      $("#providersFile").value = "";
      renderBuilder();
    }catch(err){
      console.error(err);
      toast("Providers import failed.");
    }
  });

  $("#btnLoadProvidersFromRepo").addEventListener("click", async ()=>{
    try{
      const res = await fetch("./data/providers.csv", { cache:"no-store" });
      if(!res.ok) throw new Error("Missing ./data/providers.csv");
      const text = await res.text();
      const providers = parseProvidersCSV(text);
      const payload = ensureDraftForWeek(state.currentWeekOf);
      payload.providers = mergeProviders(payload.providers || [], providers);
      ensureProviderAssignmentsScaffold(payload);
      toast(`Providers loaded from repo (${providers.length}).`);
      renderBuilder();
    }catch(e){
      console.error(e);
      toast("Load providers from repo failed. Put file at /data/providers.csv");
    }
  });

  $("#btnAddProvider").addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftForWeek(state.currentWeekOf);
    openAddProviderModal(payload);
  });

  // Time preset add
  $("#btnAddTimePreset").addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftForWeek(state.currentWeekOf);
    openAddTimePresetModal(payload);
  });

  // Draft/Published
  $("#btnLoadDraft").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    await loadDraft(state.currentWeekOf);
    if(state.draft){
      normalizeDraft(state.draft);
      toast("Draft loaded.");
    }else{
      toast("No draft found for this week.");
    }
    renderAll();
  });

  $("#btnLoadPublished").addEventListener("click", async ()=>{
    await loadPublished(state.currentWeekOf);
    toast(state.published ? "Published loaded." : "No published schedule found.");
    renderAll();
  });

  $("#btnSaveDraft").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftForWeek(state.currentWeekOf);

    // generate staff-view from provider-driven before saving
    rebuildStaffViewFromProviderDriven(payload);

    try{
      await saveSchedule("draft", payload);
      toast("Draft saved.");
    }catch(e){
      console.error(e);
      toast("Draft save failed.");
    }
  });

  $("#btnPublish").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftForWeek(state.currentWeekOf);

    rebuildStaffViewFromProviderDriven(payload);

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
      toast("Publish failed.");
    }
  });

  $("#btnValidate").addEventListener("click", ()=>{
    const payload = ensureDraftForWeek(state.currentWeekOf);
    rebuildStaffViewFromProviderDriven(payload);
    const issues = validate(payload);
    if(!issues.length) return toast("Validation passed.");
    openModal({
      title: `Validation issues (${issues.length})`,
      body: `<div class="muted" style="line-height:1.45">${issues.map(x=>`• ${escapeHtml(x)}`).join("<br>")}</div>`,
      foot: `<button class="btn btnGhost" data-close="1">Close</button>`
    });
  });

  $("#btnClearWeek").addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftForWeek(state.currentWeekOf);

    payload.providerAssignments = {};
    payload.floatAssignments = defaultFloatAssignments();
    payload.providerDone = {};
    ensureProviderAssignmentsScaffold(payload);

    toast("Cleared Draft Week.");
    renderBuilder();
  });

  // Provider done toggles
  $("#btnProviderMarkDone").addEventListener("click", ()=>{
    const payload = ensureDraftForWeek(state.currentWeekOf);
    if(!state.activeProviderId) return toast("Select a provider.");
    payload.providerDone[state.activeProviderId] = true;
    renderProviderList(payload);
    toast("Marked done.");
  });
  $("#btnProviderMarkNotDone").addEventListener("click", ()=>{
    const payload = ensureDraftForWeek(state.currentWeekOf);
    if(!state.activeProviderId) return toast("Select a provider.");
    payload.providerDone[state.activeProviderId] = false;
    renderProviderList(payload);
    toast("Marked not done.");
  });

  // Everyone filters
  $("#roleFilter").addEventListener("change", renderEveryone);
  $("#searchFilter").addEventListener("input", renderEveryone);

  // Provider search
  $("#providerSearch").addEventListener("input", (e)=>{
    state.providerSearch = (e.target.value || "").trim().toLowerCase();
    const payload = ensureDraftForWeek(state.currentWeekOf);
    renderProviderList(payload);
  });

  // Modal close
  $("#modal").addEventListener("click", (e)=>{
    if(e.target?.dataset?.close) closeModal();
  });
}

/* ---------------- Data model normalization ---------------- */

function normalizeDraft(d){
  d.weekOf = d.weekOf || state.currentWeekOf;
  d.roster = d.roster || [];
  d.providers = d.providers || [];
  d.templates = d.templates || [];
  d.entries = d.entries || {};
  d.assignmentMeta = d.assignmentMeta || {};
  d.timePresets = d.timePresets || DEFAULT_TIME_PRESETS.map(x=>({...x}));
  d.providerAssignments = d.providerAssignments || {};
  d.providerDone = d.providerDone || {};
  d.floatAssignments = d.floatAssignments || defaultFloatAssignments();
  ensureProviderAssignmentsScaffold(d);
}

function defaultFloatAssignments(){
  // per day: floatMA staffId + location + preset, floatXR staffId + location + preset + room
  const o = {};
  DAYS.forEach(d=>{
    o[d] = {
      ma: { staffId:null, location:"", presetId:"" },
      xr: { staffId:null, location:"", presetId:"", roomId:"" }
    };
  });
  return o;
}

function ensureProviderAssignmentsScaffold(payload){
  payload.providerAssignments = payload.providerAssignments || {};
  payload.providerDone = payload.providerDone || {};
  payload.floatAssignments = payload.floatAssignments || defaultFloatAssignments();
  payload.timePresets = payload.timePresets || DEFAULT_TIME_PRESETS.map(x=>({...x}));

  (payload.providers||[]).forEach(p=>{
    if(!payload.providerAssignments[p.id]){
      payload.providerAssignments[p.id] = {};
      DAYS.forEach(d=>{
        payload.providerAssignments[p.id][d] = {
          off: false,
          maStaffId: null,
          // segment A always exists
          segA: { location:"", presetId:"" },
          // segment B optional
          segBEnabled: false,
          segB: { location:"", presetId:"" },
          xrStaffId: null,
          xrLocation: "",
          xrRoomId: "",
          xrPresetId: ""
        };
      });
    }
    if(payload.providerDone[p.id] === undefined) payload.providerDone[p.id] = false;
  });

  // keep activeProviderId valid
  if(state.activeProviderId && !(payload.providers||[]).some(p=>p.id===state.activeProviderId)){
    state.activeProviderId = null;
  }
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

/* ---------------- My / Everyone (staff-driven view) ---------------- */

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
    $("#todayCard").innerHTML = `<div class="muted">Your email (${escapeHtml(state.auth.email)}) isn’t mapped to roster.</div>`;
    $("#weekCards").innerHTML = `<div class="muted">Ask an editor to add your email in Builder.</div>`;
    return;
  }

  $("#myMatchPill").textContent = `${me.name} (${me.role})`;

  const entries = sched.entries?.[me.id] || {};
  const todayName = dayName(new Date());
  const todayVal = (todayName && entries[todayName]) ? (templateTextById(sched, entries[todayName]) || "") : "";

  $("#todayCard").innerHTML = `
    <div style="font-weight:900">${formatLongDate(new Date())}</div>
    <div style="margin-top:10px; font-size:16px; font-weight:900">${escapeHtml(todayVal || "No assignment today")}</div>
  `;

  $("#weekCards").innerHTML = DAYS.map(d=>{
    const tid = entries[d];
    const txt = tid ? templateTextById(sched, tid) : "—";
    return `
      <div class="dayCard">
        <div>
          <div class="dayName">${d}</div>
          <div class="muted small">${escapeHtml(me.role)}</div>
        </div>
        <div class="dayValue">${escapeHtml(txt)}</div>
      </div>
    `;
  }).join("");
}

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
    .filter(r => !q ? true : r.name.toLowerCase().includes(q))
    .slice()
    .sort((a,b)=>a.name.localeCompare(b.name));

  thead.innerHTML = `
    <tr>
      <th style="min-width:220px;text-align:left;">Team Member</th>
      ${DAYS.map(d=>`<th style="min-width:240px;text-align:left;">${d}</th>`).join("")}
    </tr>
  `;

  tbody.innerHTML = rows.map(r=>{
    const e = sched.entries?.[r.id] || {};
    return `
      <tr>
        <td class="nameCell">
          ${escapeHtml(r.name)}
          <span class="roleTag">${escapeHtml(r.role)}</span>
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

/* ---------------- Builder (provider-driven) ---------------- */

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

  if(!draft && pub){
    state.draft = deepClone(pub);
    normalizeDraft(state.draft);
  }
  if(!state.draft){
    // start empty but scaffolded
    state.draft = {
      weekOf: week,
      roster: [],
      providers: [],
      timePresets: DEFAULT_TIME_PRESETS.map(x=>({...x})),
      providerAssignments: {},
      providerDone: {},
      floatAssignments: defaultFloatAssignments(),

      // staff-view fields (generated on save/publish)
      templates: [],
      entries: {},
      assignmentMeta: {}
    };
    normalizeDraft(state.draft);
  }

  const payload = state.draft;
  normalizeDraft(payload);

  renderRoster(payload);
  renderTimePresets(payload);
  renderLocations();

  renderProviderList(payload);
  renderActiveProvider(payload);
  renderFloatPool(payload);
}

/* Provider list left */

function renderProviderList(payload){
  const list = $("#providerList");
  const providers = (payload.providers||[]).slice().sort((a,b)=>a.name.localeCompare(b.name));
  const q = (state.providerSearch||"").toLowerCase();

  const filtered = providers.filter(p => !q ? true : p.name.toLowerCase().includes(q));

  if(!filtered.length){
    list.innerHTML = `<div class="muted" style="padding:10px">No providers loaded.</div>`;
    return;
  }

  list.innerHTML = filtered.map(p=>{
    const done = isProviderComplete(payload, p.id) || !!payload.providerDone[p.id];
    const isActive = state.activeProviderId === p.id;
    const dotClass = done ? "good" : "bad";
    return `
      <div class="providerItem ${isActive ? "is-active":""}" data-provider="${escapeHtml(p.id)}">
        <div>
          <div class="providerName">${escapeHtml(p.name)}</div>
          <div class="providerMeta">${done ? "Complete" : "Incomplete"}</div>
        </div>
        <div class="dot ${dotClass}"></div>
      </div>
    `;
  }).join("");

  list.querySelectorAll("[data-provider]").forEach(el=>{
    el.addEventListener("click", ()=>{
      state.activeProviderId = el.dataset.provider;
      renderProviderList(payload);
      renderActiveProvider(payload);
    });
  });
}

/* Active provider editor */

function renderActiveProvider(payload){
  const title = $("#activeProviderTitle");
  const sub = $("#activeProviderSub");
  const table = $("#providerWeekGrid");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  if(!state.activeProviderId){
    title.textContent = "Select a provider";
    sub.textContent = "Then assign Mon–Fri";
    thead.innerHTML = "";
    tbody.innerHTML = `<tr><td class="muted" style="padding:14px">Pick a provider from the left.</td></tr>`;
    return;
  }

  const provider = (payload.providers||[]).find(p=>p.id===state.activeProviderId);
  if(!provider){
    state.activeProviderId = null;
    return renderActiveProvider(payload);
  }

  title.textContent = provider.name;
  sub.textContent = `Assign MA + Location(s) + Time(s), plus XR room if needed.`;

  thead.innerHTML = `
    <tr>
      <th style="min-width:120px">Day</th>
      <th style="min-width:210px">MA</th>
      <th style="min-width:260px">Segment A (Location + Time)</th>
      <th style="min-width:310px">Segment B (optional)</th>
      <th style="min-width:280px">XR Room (optional)</th>
      <th style="min-width:140px">Actions</th>
    </tr>
  `;

  const locks = computeDayLocks(payload); // {day:{MA:Set, XR:Set}}

  tbody.innerHTML = DAYS.map(day=>{
    const cell = payload.providerAssignments[provider.id][day];

    const maName = cell.off ? "OFF" : staffLabel(payload, cell.maStaffId, "MA");
    const segA = cell.off ? "—" : segmentLabel(payload, cell.segA);
    const segB = cell.off ? "—" : (cell.segBEnabled ? segmentLabel(payload, cell.segB) : "—");
    const xr = cell.off ? "—" : xrLabel(payload, cell);

    const maLockedText = cell.maStaffId ? lockTextForStaffDay(payload, cell.maStaffId, day, provider.id, "MA") : "";
    const xrLockedText = cell.xrStaffId ? lockTextForStaffDay(payload, cell.xrStaffId, day, provider.id, "XR") : "";

    const done = isProviderCompleteDay(payload, provider.id, day);

    return `
      <tr>
        <td style="font-weight:900">${day} ${done ? `<span class="roleTag" style="margin-left:8px; border-color: rgba(31,157,85,.7); background: rgba(31,157,85,.18)">Done</span>` : ""}</td>

        <td>
          <div style="font-weight:900">${escapeHtml(maName)}</div>
          ${maLockedText ? `<div class="muted small" style="margin-top:6px">${escapeHtml(maLockedText)}</div>` : ""}
        </td>

        <td>
          <div style="font-weight:900">${escapeHtml(segA)}</div>
        </td>

        <td>
          <div style="font-weight:900">${escapeHtml(segB)}</div>
        </td>

        <td>
          <div style="font-weight:900">${escapeHtml(xr)}</div>
          ${xrLockedText ? `<div class="muted small" style="margin-top:6px">${escapeHtml(xrLockedText)}</div>` : ""}
        </td>

        <td>
          <button class="btn btnSmall btnGhost" data-edit="${provider.id}" data-day="${day}">Edit</button>
          <button class="btn btnSmall btnGhost" data-clear="${provider.id}" data-day="${day}">Clear</button>
        </td>
      </tr>
    `;
  }).join("");

  tbody.querySelectorAll("[data-edit]").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      openProviderDayModal(payload, btn.dataset.edit, btn.dataset.day);
    });
  });

  tbody.querySelectorAll("[data-clear]").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      const pid = btn.dataset.clear;
      const day = btn.dataset.day;
      payload.providerAssignments[pid][day] = defaultProviderDayCell();
      payload.providerDone[pid] = false;
      renderProviderList(payload);
      renderActiveProvider(payload);
      renderFloatPool(payload);
    });
  });
}

function defaultProviderDayCell(){
  return {
    off:false,
    maStaffId:null,
    segA:{ location:"", presetId:"" },
    segBEnabled:false,
    segB:{ location:"", presetId:"" },
    xrStaffId:null,
    xrLocation:"",
    xrRoomId:"",
    xrPresetId:""
  };
}

/* Float pool */

function renderFloatPool(payload){
  const box = $("#floatBox");
  const locks = computeDayLocks(payload);
  const allProvidersCoveredByDay = computeAllProvidersCoveredByDay(payload); // {day:boolean}

  box.innerHTML = DAYS.map(day=>{
    const ok = allProvidersCoveredByDay[day];
    const float = payload.floatAssignments[day];
    const maLabel = staffLabel(payload, float.ma.staffId, "MA");
    const xrLabelTxt = staffLabel(payload, float.xr.staffId, "XR");

    return `
      <div class="floatRow">
        <div class="floatRowTitle">${day}</div>
        <div class="muted small">${ok ? "All providers covered: Float allowed" : "Providers still missing: Float blocked"}</div>

        <div class="row" style="padding:0; gap:10px; margin-top:8px">
          <div class="field grow" style="min-width:220px">
            <label>Float MA</label>
            <select data-float-ma="${day}" ${ok ? "" : "disabled"}>
              ${staffOptions(payload, "MA", float.ma.staffId, locks[day]?.MA, /*allowSelf*/ true)}
            </select>
          </div>
          <div class="field grow" style="min-width:220px">
            <label>Location</label>
            <select data-float-ma-loc="${day}" ${ok ? "" : "disabled"}>
              ${locationOptionsHTML(float.ma.location)}
            </select>
          </div>
          <div class="field" style="min-width:200px">
            <label>Time</label>
            <select data-float-ma-time="${day}" ${ok ? "" : "disabled"}>
              ${presetOptions(payload, float.ma.presetId)}
            </select>
          </div>
        </div>

        <div class="row" style="padding:0; gap:10px; margin-top:10px">
          <div class="field grow" style="min-width:220px">
            <label>Float XR</label>
            <select data-float-xr="${day}" ${ok ? "" : "disabled"}>
              ${staffOptions(payload, "XR", float.xr.staffId, locks[day]?.XR, true)}
            </select>
          </div>
          <div class="field grow" style="min-width:220px">
            <label>Location</label>
            <select data-float-xr-loc="${day}" ${ok ? "" : "disabled"}>
              ${locationOptionsHTML(float.xr.location)}
            </select>
          </div>
          <div class="field" style="min-width:200px">
            <label>Time</label>
            <select data-float-xr-time="${day}" ${ok ? "" : "disabled"}>
              ${presetOptions(payload, float.xr.presetId)}
            </select>
          </div>
          <div class="field" style="min-width:200px">
            <label>XR Room</label>
            <select data-float-xr-room="${day}" ${ok ? "" : "disabled"}>
              ${xrRoomOptions(float.xr.roomId)}
            </select>
          </div>
        </div>
      </div>
    `;
  }).join("");

  // wire float inputs
  DAYS.forEach(day=>{
    const ok = computeAllProvidersCoveredByDay(payload)[day];
    if(!ok) return;

    const selMA = box.querySelector(`[data-float-ma="${day}"]`);
    const selMALoc = box.querySelector(`[data-float-ma-loc="${day}"]`);
    const selMATime = box.querySelector(`[data-float-ma-time="${day}"]`);

    const selXR = box.querySelector(`[data-float-xr="${day}"]`);
    const selXRLoc = box.querySelector(`[data-float-xr-loc="${day}"]`);
    const selXRTime = box.querySelector(`[data-float-xr-time="${day}"]`);
    const selXRRoom = box.querySelector(`[data-float-xr-room="${day}"]`);

    selMA?.addEventListener("change", ()=>{
      payload.floatAssignments[day].ma.staffId = selMA.value || null;
      renderFloatPool(payload);
      renderActiveProvider(payload);
    });
    selMALoc?.addEventListener("change", ()=>{
      payload.floatAssignments[day].ma.location = selMALoc.value || "";
    });
    selMATime?.addEventListener("change", ()=>{
      payload.floatAssignments[day].ma.presetId = selMATime.value || "";
    });

    selXR?.addEventListener("change", ()=>{
      payload.floatAssignments[day].xr.staffId = selXR.value || null;
      renderFloatPool(payload);
      renderActiveProvider(payload);
    });
    selXRLoc?.addEventListener("change", ()=>{
      payload.floatAssignments[day].xr.location = selXRLoc.value || "";
    });
    selXRTime?.addEventListener("change", ()=>{
      payload.floatAssignments[day].xr.presetId = selXRTime.value || "";
    });
    selXRRoom?.addEventListener("change", ()=>{
      payload.floatAssignments[day].xr.roomId = selXRRoom.value || "";
    });
  });
}

/* Roster list (with role editing) */

function renderRoster(payload){
  const wrap = $("#rosterEmailList");
  const roster = (payload.roster||[]).slice().sort((a,b)=>a.name.localeCompare(b.name));
  if(!roster.length){
    wrap.innerHTML = `<div class="muted" style="padding:10px">No staff loaded. Import staff.</div>`;
    return;
  }

  const locks = computeDayLocks(payload); // day => {MA:Set, XR:Set}

  wrap.innerHTML = roster.map(r=>{
    const lockedDays = DAYS.filter(d => (locks[d]?.[r.role] || new Set()).has(r.id));
    const lockMsg = lockedDays.length ? `Locked: ${lockedDays.join(", ")}` : `Not locked this week`;
    return `
      <div class="rosterItem">
        <div class="rosterTop">
          <div>
            <div class="rosterName">${escapeHtml(r.name)} <span class="roleTag">${escapeHtml(r.role)}</span></div>
            <div class="muted small" style="margin-top:4px">${escapeHtml(r.email || "")}</div>
            <div class="lockPill">🔒 ${escapeHtml(lockMsg)}</div>
          </div>
          <div class="field" style="min-width:120px; margin:0">
            <label>Role</label>
            <select data-role-for="${escapeHtml(r.id)}">
              <option value="MA" ${r.role==="MA"?"selected":""}>MA</option>
              <option value="XR" ${r.role==="XR"?"selected":""}>XR</option>
            </select>
          </div>
        </div>

        <div class="field" style="margin-top:10px">
          <label>Email</label>
          <input data-email-for="${escapeHtml(r.id)}" type="email" placeholder="name@laorthos.com" value="${escapeHtml(r.email || "")}" />
        </div>
      </div>
    `;
  }).join("");

  wrap.querySelectorAll("input[data-email-for]").forEach(inp=>{
    inp.addEventListener("input", ()=>{
      const id = inp.dataset.emailFor;
      const person = payload.roster.find(x=>x.id===id);
      if(person) person.email = (inp.value||"").trim();
    });
  });

  wrap.querySelectorAll("select[data-role-for]").forEach(sel=>{
    sel.addEventListener("change", ()=>{
      const id = sel.dataset.roleFor;
      const person = payload.roster.find(x=>x.id===id);
      if(!person) return;
      person.role = sel.value;
      toast(`${person.name} role set to ${person.role}`);
      renderRoster(payload);
      renderActiveProvider(payload);
      renderFloatPool(payload);
    });
  });
}

/* Time presets list */

function renderTimePresets(payload){
  const list = $("#presetList");
  payload.timePresets = payload.timePresets || DEFAULT_TIME_PRESETS.map(x=>({...x}));

  list.innerHTML = payload.timePresets.map(p=>`
    <div class="presetItem">
      <div>
        <div class="presetName">${escapeHtml(p.label)}</div>
        <div class="muted small">${escapeHtml(p.start)} → ${escapeHtml(p.end)}</div>
      </div>
      <button class="iconBtn" title="Delete" data-del-preset="${escapeHtml(p.id)}">🗑</button>
    </div>
  `).join("");

  list.querySelectorAll("[data-del-preset]").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      const id = btn.dataset.delPreset;
      // block deleting if used
      if(isPresetUsed(payload, id)){
        toast("Preset is used in schedule. Clear assignments first.");
        return;
      }
      payload.timePresets = payload.timePresets.filter(x=>x.id !== id);
      toast("Preset deleted.");
      renderTimePresets(payload);
      renderActiveProvider(payload);
      renderFloatPool(payload);
    });
  });
}

function isPresetUsed(payload, presetId){
  for(const pid of Object.keys(payload.providerAssignments||{})){
    for(const day of DAYS){
      const c = payload.providerAssignments[pid][day];
      if(c.segA?.presetId===presetId) return true;
      if(c.segBEnabled && c.segB?.presetId===presetId) return true;
      if(c.xrPresetId===presetId) return true;
    }
  }
  for(const day of DAYS){
    const f = payload.floatAssignments?.[day];
    if(f?.ma?.presetId===presetId) return true;
    if(f?.xr?.presetId===presetId) return true;
  }
  return false;
}

/* Locations */

function renderLocations(){
  $("#locationList").innerHTML = LOCATIONS.map(l=>`
    <div class="locationChip">${escapeHtml(l.name)} <span>${escapeHtml(l.code)}</span></div>
  `).join("");
}

/* ---------------- Provider day modal ---------------- */

function openProviderDayModal(payload, providerId, day){
  const provider = payload.providers.find(p=>p.id===providerId);
  if(!provider) return;

  const locks = computeDayLocks(payload);
  const dayLocksMA = locks[day]?.MA || new Set();
  const dayLocksXR = locks[day]?.XR || new Set();

  const cell = payload.providerAssignments[providerId][day];
  const staffMA = payload.roster.filter(r=>r.role==="MA");
  const staffXR = payload.roster.filter(r=>r.role==="XR");

  openModal({
    title: `${provider.name} • ${day}`,
    body: `
      <div class="field" style="min-width:100%">
        <label>Status</label>
        <select id="pStatus">
          <option value="WORK" ${!cell.off?"selected":""}>Working</option>
          <option value="OFF" ${cell.off?"selected":""}>OFF</option>
        </select>
        <div class="hint">OFF clears MA/XR for this provider/day.</div>
      </div>

      <div id="pWork">
        <div class="row" style="padding:0; margin-top:12px; gap:12px">
          <div class="field grow" style="min-width:240px">
            <label>MA (required)</label>
            <select id="pMA">
              ${staffOptions(payload, "MA", cell.maStaffId, dayLocksMA, /*allowSelf*/ true, providerId, day)}
            </select>
            <div class="hint">Safeguard: MA can only be assigned to one provider per day.</div>
          </div>

          <div class="field grow" style="min-width:240px">
            <label>Segment A Location</label>
            <select id="pLocA">${locationOptionsHTML(cell.segA.location)}</select>
          </div>

          <div class="field" style="min-width:220px">
            <label>Segment A Time</label>
            <select id="pTimeA">${presetOptions(payload, cell.segA.presetId)}</select>
          </div>
        </div>

        <div class="field" style="min-width:100%; margin-top:12px">
          <label>Two locations same day?</label>
          <select id="pSegBEnabled">
            <option value="NO" ${!cell.segBEnabled?"selected":""}>No</option>
            <option value="YES" ${cell.segBEnabled?"selected":""}>Yes (Segment B)</option>
          </select>
          <div class="hint">Use this when provider moves locations in the same day. MA follows provider.</div>
        </div>

        <div id="pSegBRow" class="row" style="padding:0; margin-top:12px; gap:12px">
          <div class="field grow" style="min-width:240px">
            <label>Segment B Location</label>
            <select id="pLocB">${locationOptionsHTML(cell.segB.location)}</select>
          </div>
          <div class="field" style="min-width:220px">
            <label>Segment B Time</label>
            <select id="pTimeB">${presetOptions(payload, cell.segB.presetId)}</select>
          </div>
        </div>

        <div class="divider"></div>

        <div class="row" style="padding:0; margin-top:12px; gap:12px">
          <div class="field grow" style="min-width:240px">
            <label>XR Staff (optional)</label>
            <select id="pXR">
              ${staffOptions(payload, "XR", cell.xrStaffId, dayLocksXR, true, providerId, day)}
            </select>
            <div class="hint">Safeguard: XR staff can only be assigned to one provider per day.</div>
          </div>

          <div class="field grow" style="min-width:240px">
            <label>XR Location</label>
            <select id="pXRLoc">${locationOptionsHTML(cell.xrLocation)}</select>
          </div>

          <div class="field" style="min-width:200px">
            <label>XR Room</label>
            <select id="pXRRoom">${xrRoomOptions(cell.xrRoomId)}</select>
          </div>

          <div class="field" style="min-width:220px">
            <label>XR Time</label>
            <select id="pXRTime">${presetOptions(payload, cell.xrPresetId)}</select>
          </div>
        </div>
      </div>
    `,
    foot: `
      <button class="btn btnGhost" data-close="1">Cancel</button>
      <button class="btn btnGhost" id="pClear">Clear Day</button>
      <button class="btn" id="pSave">Save</button>
    `,
    onAfter(){
      const workWrap = $("#pWork");
      const statusSel = $("#pStatus");
      const segBEnabledSel = $("#pSegBEnabled");
      const segBRow = $("#pSegBRow");

      const sync = ()=>{
        const off = statusSel.value === "OFF";
        workWrap.style.display = off ? "none" : "";
        segBRow.style.display = (segBEnabledSel.value === "YES") ? "" : "none";
      };
      statusSel.addEventListener("change", sync);
      segBEnabledSel.addEventListener("change", sync);
      sync();

      $("#pClear").addEventListener("click", ()=>{
        payload.providerAssignments[providerId][day] = defaultProviderDayCell();
        payload.providerDone[providerId] = false;
        closeModal();
        renderProviderList(payload);
        renderActiveProvider(payload);
        renderFloatPool(payload);
      });

      $("#pSave").addEventListener("click", ()=>{
        const off = ($("#pStatus").value === "OFF");
        if(off){
          payload.providerAssignments[providerId][day] = defaultProviderDayCell();
          payload.providerAssignments[providerId][day].off = true;
          payload.providerDone[providerId] = false;
          closeModal();
          renderProviderList(payload);
          renderActiveProvider(payload);
          renderFloatPool(payload);
          return;
        }

        const maStaffId = ($("#pMA").value || "").trim() || null;
        if(!maStaffId) return toast("MA is required.");

        // enforce lock (unless it’s currently assigned to this same provider/day)
        if(isStaffLockedForDay(payload, maStaffId, day, "MA", providerId)){
          return toast("That MA is already assigned to another provider for this day.");
        }

        const locA = ($("#pLocA").value||"").trim();
        const timeA = ($("#pTimeA").value||"").trim();
        if(!locA || !timeA) return toast("Segment A location and time are required.");

        const segBEnabled = ($("#pSegBEnabled").value === "YES");
        let locB = "", timeB = "";
        if(segBEnabled){
          locB = ($("#pLocB").value||"").trim();
          timeB = ($("#pTimeB").value||"").trim();
          if(!locB || !timeB) return toast("Segment B location and time are required.");
        }

        const xrStaffId = ($("#pXR").value||"").trim() || null;
        const xrLoc = ($("#pXRLoc").value||"").trim();
        const xrRoomId = ($("#pXRRoom").value||"").trim();
        const xrTime = ($("#pXRTime").value||"").trim();

        if(xrStaffId){
          if(isStaffLockedForDay(payload, xrStaffId, day, "XR", providerId)){
            return toast("That XR staff is already assigned to another provider for this day.");
          }
          if(!xrLoc || !xrRoomId || !xrTime) return toast("XR requires location + room + time.");
        }

        payload.providerAssignments[providerId][day] = {
          off:false,
          maStaffId,
          segA:{ location: locA, presetId: timeA },
          segBEnabled,
          segB:{ location: locB, presetId: timeB },
          xrStaffId,
          xrLocation: xrStaffId ? xrLoc : "",
          xrRoomId: xrStaffId ? xrRoomId : "",
          xrPresetId: xrStaffId ? xrTime : ""
        };

        // auto-done logic: complete if MA assigned + segA complete; segB if enabled complete; XR if set complete
        payload.providerDone[providerId] = payload.providerDone[providerId] || false;

        closeModal();
        renderProviderList(payload);
        renderActiveProvider(payload);
        renderFloatPool(payload);
      });
    }
  });
}

/* ---------------- Completion rules ---------------- */

function isProviderComplete(payload, providerId){
  // complete if each weekday has either OFF or valid MA + segA (+ segB if enabled) and XR is either empty or valid
  return DAYS.every(d => isProviderCompleteDay(payload, providerId, d));
}

function isProviderCompleteDay(payload, providerId, day){
  const c = payload.providerAssignments?.[providerId]?.[day];
  if(!c) return false;
  if(c.off) return true;

  if(!c.maStaffId) return false;
  if(!c.segA?.location || !c.segA?.presetId) return false;
  if(c.segBEnabled){
    if(!c.segB?.location || !c.segB?.presetId) return false;
  }
  if(c.xrStaffId){
    if(!c.xrLocation || !c.xrRoomId || !c.xrPresetId) return false;
  }
  return true;
}

function computeAllProvidersCoveredByDay(payload){
  const out = {};
  const providers = payload.providers || [];
  DAYS.forEach(day=>{
    out[day] = providers.every(p => isProviderCompleteDay(payload, p.id, day));
  });
  return out;
}

/* ---------------- Locks (Safeguard) ---------------- */

function computeDayLocks(payload){
  const locks = {};
  DAYS.forEach(d=>{
    locks[d] = { MA: new Set(), XR: new Set() };
  });

  for(const pid of Object.keys(payload.providerAssignments||{})){
    for(const day of DAYS){
      const c = payload.providerAssignments[pid][day];
      if(!c || c.off) continue;

      if(c.maStaffId) locks[day].MA.add(c.maStaffId);
      if(c.xrStaffId) locks[day].XR.add(c.xrStaffId);
    }
  }

  // float locks too (count as lock)
  for(const day of DAYS){
    const f = payload.floatAssignments?.[day];
    if(f?.ma?.staffId) locks[day].MA.add(f.ma.staffId);
    if(f?.xr?.staffId) locks[day].XR.add(f.xr.staffId);
  }

  return locks;
}

function isStaffLockedForDay(payload, staffId, day, role, currentProviderId){
  // locked if assigned to a different provider (or float) on same day
  for(const pid of Object.keys(payload.providerAssignments||{})){
    const c = payload.providerAssignments[pid][day];
    if(!c || c.off) continue;

    if(role === "MA" && c.maStaffId === staffId){
      if(pid !== currentProviderId) return true;
    }
    if(role === "XR" && c.xrStaffId === staffId){
      if(pid !== currentProviderId) return true;
    }
  }

  const f = payload.floatAssignments?.[day];
  if(role==="MA" && f?.ma?.staffId === staffId) return true;
  if(role==="XR" && f?.xr?.staffId === staffId) return true;

  return false;
}

function lockTextForStaffDay(payload, staffId, day, currentProviderId, role){
  // show where they are assigned (provider name or FLOAT)
  for(const pid of Object.keys(payload.providerAssignments||{})){
    const c = payload.providerAssignments[pid][day];
    if(!c || c.off) continue;
    if(role==="MA" && c.maStaffId===staffId && pid!==currentProviderId){
      const p = payload.providers.find(x=>x.id===pid);
      return `Locked to ${p?.name || "another provider"} (${day})`;
    }
    if(role==="XR" && c.xrStaffId===staffId && pid!==currentProviderId){
      const p = payload.providers.find(x=>x.id===pid);
      return `Locked to ${p?.name || "another provider"} (${day})`;
    }
  }
  const f = payload.floatAssignments?.[day];
  if(role==="MA" && f?.ma?.staffId===staffId) return `Locked to FLOAT (${day})`;
  if(role==="XR" && f?.xr?.staffId===staffId) return `Locked to FLOAT (${day})`;
  return "";
}

/* ---------------- Staff/Provider labels/options ---------------- */

function staffLabel(payload, staffId, role){
  if(!staffId) return role === "MA" ? "Unassigned" : "None";
  const s = (payload.roster||[]).find(r=>r.id===staffId);
  return s ? s.name : "Unknown";
}

function segmentLabel(payload, seg){
  if(!seg?.location || !seg?.presetId) return "—";
  const loc = seg.location;
  const preset = (payload.timePresets||[]).find(x=>x.id===seg.presetId);
  return `${loc} • ${preset ? preset.label : seg.presetId}`;
}

function xrLabel(payload, cell){
  if(!cell.xrStaffId) return "—";
  const staff = staffLabel(payload, cell.xrStaffId, "XR");
  const preset = (payload.timePresets||[]).find(x=>x.id===cell.xrPresetId);
  const time = preset ? preset.label : "";
  return `${staff} • ${cell.xrLocation} • ${cell.xrRoomId} • ${time}`.trim();
}

function staffOptions(payload, role, selectedId, lockedSet, allowSelf=true, currentProviderId=null, day=null){
  const roster = (payload.roster||[]).filter(r => r.role===role).slice().sort((a,b)=>a.name.localeCompare(b.name));
  const opts = [
    `<option value="">— ${role==="MA" ? "Select MA" : "None"} —</option>`
  ];

  roster.forEach(r=>{
    let disabled = false;
    let tag = "";

    if(day){
      // more accurate: check if locked elsewhere
      const lockedElsewhere = isStaffLockedForDay(payload, r.id, day, role, currentProviderId);
      if(lockedElsewhere && r.id !== selectedId){
        disabled = true;
        tag = " (locked)";
      }
    }else{
      // fallback
      if(lockedSet && lockedSet.has(r.id) && r.id !== selectedId) disabled = true;
    }

    opts.push(`<option value="${escapeHtml(r.id)}" ${r.id===selectedId?"selected":""} ${disabled?"disabled":""}>${escapeHtml(r.name)}${tag}</option>`);
  });

  return opts.join("");
}

function locationOptionsHTML(selectedCode){
  const opts = [`<option value="">— Select —</option>`]
    .concat(LOCATIONS.map(l=>`<option value="${l.code}" ${l.code===selectedCode?"selected":""}>${l.name} (${l.code})</option>`));
  return opts.join("");
}

function presetOptions(payload, selectedId){
  const presets = (payload.timePresets||[]).slice();
  const opts = [`<option value="">— Select —</option>`]
    .concat(presets.map(p => `<option value="${p.id}" ${p.id===selectedId?"selected":""}>${escapeHtml(p.label)}</option>`));
  return opts.join("");
}

function xrRoomOptions(selectedId){
  const opts = [`<option value="">— Select —</option>`]
    .concat(XR_ROOMS.map(r => `<option value="${r.id}" ${r.id===selectedId?"selected":""}>${escapeHtml(r.label)}</option>`));
  return opts.join("");
}

/* ---------------- Modals: Add staff/provider/preset ---------------- */

function openAddStaffModal(payload){
  openModal({
    title: "Add Staff",
    body: `
      <div class="field" style="min-width:100%">
        <label>Full name</label>
        <input id="addStaffName" type="text" placeholder="First Last" />
      </div>
      <div class="field" style="min-width:100%; margin-top:10px">
        <label>Email</label>
        <input id="addStaffEmail" type="email" placeholder="name@laorthos.com" />
      </div>
      <div class="field" style="min-width:100%; margin-top:10px">
        <label>Role</label>
        <select id="addStaffRole">
          <option value="MA">MA</option>
          <option value="XR">XR</option>
        </select>
      </div>
    `,
    foot: `
      <button class="btn btnGhost" data-close="1">Cancel</button>
      <button class="btn" id="addStaffSave">Add</button>
    `,
    onAfter(){
      $("#addStaffSave").addEventListener("click", ()=>{
        const name = ($("#addStaffName").value||"").trim();
        const email = ($("#addStaffEmail").value||"").trim();
        const role = $("#addStaffRole").value;

        if(!name) return toast("Name required.");
        const idBase = slug(name);
        const id = payload.roster.some(r=>r.id===idBase) ? `${idBase}-${Math.random().toString(16).slice(2,6)}` : idBase;

        payload.roster.push({ id, name, role, email });
        closeModal();
        toast("Staff added.");
        renderRoster(payload);
        renderActiveProvider(payload);
        renderFloatPool(payload);
      });
    }
  });
}

function openAddProviderModal(payload){
  openModal({
    title: "Add Provider",
    body: `
      <div class="field" style="min-width:100%">
        <label>Provider name</label>
        <input id="addProvName" type="text" placeholder="Dr. Lastname" />
      </div>
    `,
    foot: `
      <button class="btn btnGhost" data-close="1">Cancel</button>
      <button class="btn" id="addProvSave">Add</button>
    `,
    onAfter(){
      $("#addProvSave").addEventListener("click", ()=>{
        const name = ($("#addProvName").value||"").trim();
        if(!name) return toast("Provider name required.");
        payload.providers = mergeProviders(payload.providers || [], [{ id: slug(name), name }]);
        ensureProviderAssignmentsScaffold(payload);
        closeModal();
        toast("Provider added.");
        renderProviderList(payload);
      });
    }
  });
}

function openAddTimePresetModal(payload){
  openModal({
    title: "Add Time Preset",
    body: `
      <div class="field" style="min-width:100%">
        <label>Label</label>
        <input id="tpLabel" type="text" placeholder="Example: 9:00am - 5:30pm" />
      </div>
      <div class="row" style="padding:0; margin-top:12px; gap:12px">
        <div class="field grow" style="min-width:220px">
          <label>Start</label>
          <input id="tpStart" type="time" />
        </div>
        <div class="field grow" style="min-width:220px">
          <label>End</label>
          <input id="tpEnd" type="time" />
        </div>
      </div>
      <div class="hint">This will become selectable everywhere (Provider days + Float).</div>
    `,
    foot: `
      <button class="btn btnGhost" data-close="1">Cancel</button>
      <button class="btn" id="tpSave">Add</button>
    `,
    onAfter(){
      $("#tpSave").addEventListener("click", ()=>{
        const label = ($("#tpLabel").value||"").trim();
        const start = ($("#tpStart").value||"").trim();
        const end = ($("#tpEnd").value||"").trim();
        if(!label) return toast("Label required.");
        if(!start || !end) return toast("Start and end required.");

        const id = `tp_${Math.random().toString(16).slice(2,10)}`;
        payload.timePresets.push({ id, label, start, end });
        closeModal();
        toast("Preset added.");
        renderTimePresets(payload);
        renderActiveProvider(payload);
        renderFloatPool(payload);
      });
    }
  });
}

/* ---------------- Staff/prov CSV parsing ---------------- */

function parseStaffCSV(text){
  const rows = csvToRows(text).filter(r => r.some(x => (x||"").trim() !== ""));
  if(!rows.length) return [];

  const header = rows[0].map(x => (x||"").trim().toLowerCase());
  const idxFirst = header.findIndex(h => ["first name","firstname","first"].includes(h));
  const idxLast = header.findIndex(h => ["last name","lastname","last"].includes(h));
  const idxEmail = header.findIndex(h => ["email","e-mail","user principal name","upn"].includes(h));
  const idxRole = header.findIndex(h => ["role","staff role","type"].includes(h));

  if(idxEmail < 0) throw new Error("No Email column found.");

  const out = [];
  for(let i=1;i<rows.length;i++){
    const r = rows[i];
    const email = (r[idxEmail]||"").trim();
    if(!email) continue;

    const first = idxFirst >= 0 ? (r[idxFirst]||"").trim() : "";
    const last = idxLast >= 0 ? (r[idxLast]||"").trim() : "";
    const name = (first || last) ? `${first} ${last}`.trim() : email.split("@")[0];

    let role = "";
    if(idxRole >= 0){
      role = (r[idxRole]||"").trim().toUpperCase();
      if(role === "XRT") role = "XR";
      if(role !== "MA" && role !== "XR") role = "";
    }

    out.push({ name, email, role });
  }
  return out;
}

function mergeStaffIntoRoster(payload, staff, { defaultRole="MA" } = {}){
  payload.roster = payload.roster || [];
  const existingByEmail = new Map(payload.roster.map(r => [(r.email||"").toLowerCase(), r]));

  staff.forEach(s=>{
    const email = (s.email||"").trim();
    if(!email) return;
    const key = email.toLowerCase();

    if(existingByEmail.has(key)){
      const person = existingByEmail.get(key);
      if(s.name) person.name = s.name;
      if(s.role) person.role = s.role;
      if(!person.role) person.role = defaultRole;
      return;
    }

    const name = (s.name||"").trim() || email.split("@")[0];
    const idBase = slug(name);
    const id = payload.roster.some(r=>r.id===idBase) ? `${idBase}-${Math.random().toString(16).slice(2,6)}` : idBase;

    payload.roster.push({ id, name, role: s.role || defaultRole, email });
    existingByEmail.set(key, payload.roster[payload.roster.length-1]);
  });
}

function parseProvidersCSV(text){
  const rows = csvToRows(text).map(r => (r[0]||"").trim()).filter(Boolean);
  if(!rows.length) return [];
  const first = rows[0].toLowerCase();
  const startIdx = (first.includes("provider")) ? 1 : 0;

  const out = [];
  for(let i=startIdx;i<rows.length;i++){
    const name = rows[i].trim();
    if(!name) continue;
    out.push({ id: slug(name), name });
  }
  return out;
}

function mergeProviders(current, incoming){
  const byName = new Map((current||[]).map(p => [(p.name||"").toLowerCase(), p]));
  incoming.forEach(p=>{
    const name = (p.name||"").trim();
    if(!name) return;
    const key = name.toLowerCase();
    if(byName.has(key)) return;
    byName.set(key, { id: p.id || slug(name), name });
  });
  return Array.from(byName.values()).sort((a,b)=>a.name.localeCompare(b.name));
}

/* ---------------- Provider-driven -> Staff-driven generation ---------------- */

function rebuildStaffViewFromProviderDriven(payload){
  // Creates payload.entries + payload.templates from providerAssignments + floatAssignments
  payload.templates = [];
  payload.entries = {};
  payload.assignmentMeta = {};

  const roster = payload.roster || [];
  roster.forEach(r=>{
    payload.entries[r.id] = {};
    DAYS.forEach(d=> payload.entries[r.id][d] = null);
    payload.assignmentMeta[r.id] = {};
  });

  // helper to add template + assign
  const put = (staffId, day, text, meta) => {
    const tid = ensureTemplate(payload, text);
    payload.entries[staffId][day] = tid;
    payload.assignmentMeta[staffId][day] = meta || {};
  };

  // provider assignments
  for(const pid of Object.keys(payload.providerAssignments||{})){
    const provider = (payload.providers||[]).find(p=>p.id===pid);
    for(const day of DAYS){
      const c = payload.providerAssignments[pid][day];
      if(!c) continue;

      // OFF: we leave staff assignment empty (provider off doesn't automatically mean MA off)
      // If you want "provider OFF day" to show, we can add a provider-level report later.
      if(c.off) continue;

      // MA assignment text
      if(c.maStaffId){
        const maPresetA = findPreset(payload, c.segA.presetId);
        const segAText = `${c.segA.location} ${maPresetA?.label || ""}`.trim();

        let segBText = "";
        if(c.segBEnabled){
          const maPresetB = findPreset(payload, c.segB.presetId);
          segBText = ` + ${c.segB.location} ${maPresetB?.label || ""}`.trim();
        }

        const text = `MA ${provider?.name || ""} • ${segAText}${segBText}`.trim();
        put(c.maStaffId, day, text, {
          kind:"MA",
          providerId: pid,
          providerName: provider?.name || "",
          segA: { ...c.segA },
          segBEnabled: !!c.segBEnabled,
          segB: { ...c.segB },
        });
      }

      // XR assignment text
      if(c.xrStaffId){
        const xrPreset = findPreset(payload, c.xrPresetId);
        const text = `XR ${provider?.name || ""} • ${c.xrLocation} • ${c.xrRoomId} • ${xrPreset?.label || ""}`.trim();
        put(c.xrStaffId, day, text, {
          kind:"XR",
          providerId: pid,
          providerName: provider?.name || "",
          location: c.xrLocation,
          room: c.xrRoomId,
          presetId: c.xrPresetId
        });
      }
    }
  }

  // Float assignments
  for(const day of DAYS){
    const f = payload.floatAssignments?.[day];
    if(!f) continue;

    if(f.ma?.staffId){
      const preset = findPreset(payload, f.ma.presetId);
      const text = `MA FLOAT • ${f.ma.location} • ${preset?.label || ""}`.trim();
      put(f.ma.staffId, day, text, { kind:"MA", float:true, location:f.ma.location, presetId:f.ma.presetId });
    }

    if(f.xr?.staffId){
      const preset = findPreset(payload, f.xr.presetId);
      const text = `XR FLOAT • ${f.xr.location} • ${f.xr.roomId} • ${preset?.label || ""}`.trim();
      put(f.xr.staffId, day, text, { kind:"XR", float:true, location:f.xr.location, room:f.xr.roomId, presetId:f.xr.presetId });
    }
  }
}

function ensureTemplate(payload, raw){
  const id = idForRaw(raw);
  if(!payload.templates.some(t=>t.id===id)){
    payload.templates.push({ id, raw });
  }
  return id;
}

function findPreset(payload, presetId){
  return (payload.timePresets||[]).find(p=>p.id===presetId) || null;
}

/* ---------------- Validation ---------------- */

function validate(payload){
  const issues = [];

  // roster email required for My Schedule
  (payload.roster||[]).forEach(r=>{
    if(!r.email) issues.push(`${r.name} missing email`);
  });

  // Provider schedule must have MA assigned + segA complete each day (unless OFF)
  (payload.providers||[]).forEach(p=>{
    for(const day of DAYS){
      const c = payload.providerAssignments?.[p.id]?.[day];
      if(!c) { issues.push(`${p.name} missing ${day}`); continue; }
      if(c.off) continue;
      if(!c.maStaffId) issues.push(`${p.name} missing MA on ${day}`);
      if(!c.segA?.location || !c.segA?.presetId) issues.push(`${p.name} missing Segment A on ${day}`);
      if(c.segBEnabled){
        if(!c.segB?.location || !c.segB?.presetId) issues.push(`${p.name} missing Segment B on ${day}`);
      }
      if(c.xrStaffId){
        if(!c.xrLocation || !c.xrRoomId || !c.xrPresetId) issues.push(`${p.name} XR incomplete on ${day}`);
      }
    }
  });

  // lock enforcement check (after generation)
  const seen = {};
  for(const r of (payload.roster||[])){
    for(const day of DAYS){
      const tid = payload.entries?.[r.id]?.[day];
      if(!tid) continue;
      const key = `${r.role}:${r.id}:${day}`;
      if(seen[key]) issues.push(`${r.name} double-booked on ${day}`);
      seen[key] = true;
    }
  }

  return issues;
}

/* ---------------- Export CSV ---------------- */

function exportPublishedToCSV(){
  const sched = state.published;
  if(!sched?.roster?.length) return toast("No published schedule to export.");

  const header = ["TEAM MEMBER", ...DAYS];
  const lines = [header];

  const roster = sched.roster.slice().sort((a,b)=>a.name.localeCompare(b.name));
  roster.forEach(r=>{
    const e = sched.entries?.[r.id] || {};
    const teamLabel = r.role === "XR" ? `(XR) ${r.name}` : r.name;
    const row = [teamLabel, ...DAYS.map(d=>{
      const tid = e[d];
      return tid ? (templateTextById(sched, tid) || "") : "";
    })];
    lines.push(row);
  });

  const csv = lines.map(arr => arr.map(csvEscape).join(",")).join("\n");
  downloadTextFile(`Schedule_${state.currentWeekOf}.csv`, csv, "text/csv");
  toast("CSV exported.");
}

/* ---------------- Template lookup for viewing ---------------- */

function templateTextById(sched, id){
  const t = (sched.templates||[]).find(x=>x.id===id);
  return t ? (t.raw || "") : "";
}

/* ---------------- Helpers ---------------- */

function ensureDraftForWeek(weekOf){
  if(!state.draft || state.draft.weekOf !== weekOf){
    state.draft = {
      weekOf,
      roster: [],
      providers: [],
      timePresets: DEFAULT_TIME_PRESETS.map(x=>({...x})),
      providerAssignments: {},
      providerDone: {},
      floatAssignments: defaultFloatAssignments(),
      templates: [],
      entries: {},
      assignmentMeta: {}
    };
  }
  normalizeDraft(state.draft);
  return state.draft;
}

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
    .slice(0,60) || "item";
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
