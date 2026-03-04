const DAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday"];
const DEFAULTS_PK = "__defaults__";
const DEFAULT_ROSTER_RK = "roster";
const DEFAULT_PROVIDERS_RK = "providers";

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

const $ = (sel) => document.querySelector(sel);

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
  const rk = kind === "roster" ? DEFAULT_ROSTER_RK : DEFAULT_PROVIDERS_RK;
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

/* ---------------- Buttons ---------------- */

function wireButtons(){
  $("#btnPrint").addEventListener("click", ()=> window.print());
  $("#btnExportCSV").addEventListener("click", exportPublishedToCSV);

  $("#btnLogin").addEventListener("click", login);
  $("#btnLogout").addEventListener("click", logout);

  // Schedule CSV import
  $("#btnImportScheduleCSV").addEventListener("click", ()=> $("#scheduleFile").click());
  $("#scheduleFile").addEventListener("change", async (e)=>{
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
      $("#scheduleFile").value = "";
      renderAll();
    }catch(err){
      console.error(err);
      toast("Schedule CSV import failed.");
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
      toast("Draft save failed.");
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
      toast("Publish failed.");
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

  $("#btnClearWeek").addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;

    payload.roster.forEach(r=>{
      payload.entries[r.id] = blankWeekEntries();
    });
    if(payload.assignmentMeta) payload.assignmentMeta = {};
    toast("Cleared all assignments for this week (Draft).");
    renderBuilderGrid(payload);
  });

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
          const id = slug(name);

          const emailLower = (email||"").toLowerCase();
          const existing = payload.roster.find(r => (r.email||"").toLowerCase() === emailLower && emailLower);
          if(existing){
            toast("That email already exists in roster.");
            return;
          }

          const idFinal = payload.roster.some(r=>r.id===id) ? `${id}-${Math.random().toString(16).slice(2,6)}` : id;

          payload.roster.push({ id: idFinal, name, role, email });
          payload.entries[idFinal] = payload.entries[idFinal] || blankWeekEntries();

          closeModal();
          toast("Staff added.");
          renderBuilder();
        });
      }
    });
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
      toast(`Providers imported (${providers.length}).`);
      $("#providersFile").value = "";
      renderProviders(payload);
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
      toast(`Providers loaded from repo (${providers.length}).`);
      renderProviders(payload);
    }catch(e){
      console.error(e);
      toast("Load providers from repo failed. Put file at /data/providers.csv");
    }
  });

  $("#btnAddProvider").addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftForWeek(state.currentWeekOf);

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
          payload.providers = payload.providers || [];
          payload.providers = mergeProviders(payload.providers, [{ id: slug(name), name }]);
          closeModal();
          toast("Provider added.");
          renderProviders(payload);
        });
      }
    });
  });

  $("#btnAddProviderBulk").addEventListener("click", ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftForWeek(state.currentWeekOf);
    const raw = ($("#providerBulk").value||"").trim();
    if(!raw) return toast("Paste provider names first.");
    const lines = raw.split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
    const providers = lines.map(name => ({ id: slug(name), name }));
    payload.providers = mergeProviders(payload.providers || [], providers);
    $("#providerBulk").value = "";
    toast(`Added ${providers.length} provider(s).`);
    renderProviders(payload);
  });

  // Defaults roster/providers
  $("#btnSaveDefaultRoster").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftWeek();
    if(!payload) return;

    const rosterOnly = { weekOf: DEFAULTS_PK, roster: payload.roster };
    try{ await saveSchedule(DEFAULT_ROSTER_RK, rosterOnly); toast("Default roster saved."); }
    catch(e){ console.error(e); toast("Default roster save failed."); }
  });

  $("#btnLoadDefaultRoster").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const data = await loadDefaults("roster");
    if(!data?.roster?.length) return toast("No default roster found.");
    const payload = ensureDraftForWeek(state.currentWeekOf);
    const defaults = data.roster.map(r => ({
      id: r.id || slug(r.name),
      name: r.name,
      role: r.role || "MA",
      email: r.email || ""
    }));
    mergeRoster(payload, defaults);
    toast("Default roster loaded into Draft.");
    renderBuilder();
  });

  $("#btnSaveDefaultProviders").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const payload = ensureDraftForWeek(state.currentWeekOf);
    const providers = payload.providers || [];
    const p = { weekOf: DEFAULTS_PK, providers };
    try{ await saveSchedule(DEFAULT_PROVIDERS_RK, p); toast("Default providers saved."); }
    catch(e){ console.error(e); toast("Default providers save failed."); }
  });

  $("#btnLoadDefaultProviders").addEventListener("click", async ()=>{
    if(!state.auth.editor) return toast("Not authorized.");
    const data = await loadDefaults("providers");
    if(!data?.providers?.length) return toast("No default providers found.");
    const payload = ensureDraftForWeek(state.currentWeekOf);
    payload.providers = mergeProviders(payload.providers || [], data.providers);
    toast("Default providers loaded into Draft.");
    renderProviders(payload);
  });

  // Everyone filters
  $("#roleFilter").addEventListener("change", renderEveryone);
  $("#searchFilter").addEventListener("input", renderEveryone);

  // Modal close
  $("#modal").addEventListener("click", (e)=>{
    if(e.target?.dataset?.close) closeModal();
  });

  // Auto-guess emails still works
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
      r.email = `${first[0]}${last}@laorthos.com`;
    });

    toast("Auto-guess applied. Review and correct.");
    renderRoster(payload);
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
    $("#weekCards").innerHTML = `<div class="muted">Ask an editor to add your email in Builder → Roster.</div>`;
    return;
  }

  $("#myMatchPill").textContent = `${me.name} (${me.role})`;

  const entries = sched.entries?.[me.id] || {};
  const todayName = dayName(new Date());
  const todayVal = (todayName && entries[todayName]) ? templateTextById(sched, entries[todayName]) : null;

  $("#todayCard").innerHTML = `
    <div style="font-weight:900; color:var(--char)">${formatLongDate(new Date())}</div>
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
    .filter(r => !q ? true : r.name.toLowerCase().includes(q));

  thead.innerHTML = `
    <tr>
      <th style="min-width:220px;text-align:left;">Team Member</th>
      ${DAYS.map(d=>`<th style="min-width:200px;text-align:left;">${d}</th>`).join("")}
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
  }

  const payload = ensureDraftWeek(false);
  if(!payload){
    $("#rosterEmailList").innerHTML = `<div class="muted">Import schedule / load published / load defaults.</div>`;
    $("#providerList").innerHTML = `<div class="muted">Import providers or load defaults.</div>`;
    $("#locationList").innerHTML = renderLocationBadges();
    $("#builderGrid").querySelector("thead").innerHTML = "";
    $("#builderGrid").querySelector("tbody").innerHTML = `<tr><td class="muted" style="padding:14px">—</td></tr>`;
    return;
  }

  renderRoster(payload);
  renderProviders(payload);
  renderLocations();
  renderBuilderGrid(payload);
}

/* ---------------- Draft helpers ---------------- */

function ensureDraftWeek(showToastOnFail=true){
  if(!state.draft){
    if(showToastOnFail) toast("No draft loaded. Import schedule or load published/defaults.");
    return null;
  }
  if(!state.draft.weekOf) state.draft.weekOf = state.currentWeekOf;
  if(!state.draft.templates) state.draft.templates = [];
  if(!state.draft.entries) state.draft.entries = {};
  if(!state.draft.roster) state.draft.roster = [];
  if(!state.draft.providers) state.draft.providers = [];
  if(!state.draft.assignmentMeta) state.draft.assignmentMeta = {};
  return state.draft;
}

function ensureDraftForWeek(weekOf){
  if(!state.draft || state.draft.weekOf !== weekOf){
    state.draft = {
      weekOf,
      roster: [],
      templates: [],
      entries: {},
      providers: [],
      assignmentMeta: {}
    };
  }
  ensureDraftWeek(false);
  return state.draft;
}

function blankWeekEntries(){
  const o = {};
  DAYS.forEach(d => o[d] = null);
  return o;
}

/* ---------------- Locations ---------------- */

function renderLocations(){
  $("#locationList").innerHTML = LOCATIONS.map(l => `
    <div class="locationRow">
      <div style="font-weight:900;color:rgba(31,41,55,.95)">${escapeHtml(l.name)}</div>
      <div class="roleTag">${escapeHtml(l.code)}</div>
    </div>
  `).join("");
}

function renderLocationBadges(){
  return `<div class="muted">Loaded</div>`;
}

function locationOptionsHTML(selectedCode){
  return LOCATIONS.map(l => `<option value="${l.code}" ${l.code===selectedCode?"selected":""}>${escapeHtml(l.name)} (${l.code})</option>`).join("");
}

/* ---------------- Roster (RESTORED ROLE EDIT) ---------------- */

function renderRoster(payload){
  const wrap = $("#rosterEmailList");
  if(!payload.roster.length){
    wrap.innerHTML = `<div class="muted">No roster loaded. Import staff or load default roster.</div>`;
    return;
  }

  wrap.innerHTML = payload.roster
    .slice()
    .sort((a,b)=>a.name.localeCompare(b.name))
    .map(r => `
      <div class="templateItem">
        <div class="templateTop">
          <div>
            <div class="templateName">${escapeHtml(r.name)}</div>
            <div class="templateMeta">${escapeHtml(r.id)}</div>
          </div>
          <div class="field" style="min-width:120px; margin:0">
            <label>Role</label>
            <select data-role-for="${r.id}">
              <option value="MA" ${r.role==="MA"?"selected":""}>MA</option>
              <option value="XR" ${r.role==="XR"?"selected":""}>XR</option>
            </select>
          </div>
        </div>

        <div style="margin-top:10px" class="field">
          <label>Email</label>
          <input data-email-for="${r.id}" type="email" placeholder="name@laorthos.com" value="${escapeHtml(r.email || "")}">
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

  wrap.querySelectorAll("select[data-role-for]").forEach(sel=>{
    sel.addEventListener("change", ()=>{
      const pid = sel.dataset.roleFor;
      const person = payload.roster.find(x => x.id === pid);
      if(!person) return;
      const oldRole = person.role;
      person.role = sel.value;

      // If role changed, keep existing assignments but rebuild display for that day
      if(oldRole !== person.role){
        rebuildStaffWeekDisplay(payload, person.id);
        toast(`${person.name} role updated to ${person.role}.`);
        renderBuilderGrid(payload);
      }
    });
  });
}

function rebuildStaffWeekDisplay(payload, staffId){
  const metaByDay = payload.assignmentMeta?.[staffId] || {};
  DAYS.forEach(day=>{
    const meta = metaByDay[day];
    if(!meta) return;

    // If OFF, keep OFF
    if(meta.off){
      const tid = ensureTemplate(payload, "OFF");
      payload.entries[staffId] = payload.entries[staffId] || blankWeekEntries();
      payload.entries[staffId][day] = tid;
      return;
    }

    // update role + rebuild text
    meta.role = payload.roster.find(r=>r.id===staffId)?.role || meta.role;

    const providerName = meta.providerId
      ? (payload.providers||[]).find(p=>p.id===meta.providerId)?.name || ""
      : "";

    const display = buildAssignmentText({
      role: meta.role,
      locationCode: meta.location,
      providerName,
      start: meta.start,
      end: meta.end
    });

    const tid = ensureTemplate(payload, display);
    payload.entries[staffId] = payload.entries[staffId] || blankWeekEntries();
    payload.entries[staffId][day] = tid;
  });
}

/* ---------------- Staff import ---------------- */

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
  const existingByEmail = new Map(
    (payload.roster||[]).map(r => [(r.email||"").toLowerCase(), r])
  );

  staff.forEach(s=>{
    const email = (s.email||"").trim();
    if(!email) return;
    const emailKey = email.toLowerCase();

    if(existingByEmail.has(emailKey)){
      const person = existingByEmail.get(emailKey);
      if(!person.name && s.name) person.name = s.name;
      if(s.role) person.role = s.role;
      if(!person.role) person.role = defaultRole;
      return;
    }

    const name = (s.name||"").trim() || email.split("@")[0];
    const id = slug(name);
    const idFinal = payload.roster.some(r=>r.id===id) ? `${id}-${Math.random().toString(16).slice(2,6)}` : id;

    const role = s.role || defaultRole;
    const person = { id: idFinal, name, role, email };
    payload.roster.push(person);
    payload.entries[idFinal] = payload.entries[idFinal] || blankWeekEntries();
    existingByEmail.set(emailKey, person);
  });

  payload.roster.forEach(r=>{
    payload.entries[r.id] = payload.entries[r.id] || blankWeekEntries();
  });

  if(!payload.assignmentMeta) payload.assignmentMeta = {};
}

function mergeRoster(payload, incomingRoster){
  payload.roster = payload.roster || [];
  const byId = new Map(payload.roster.map(r=>[r.id, r]));
  incomingRoster.forEach(r=>{
    if(!r.id) r.id = slug(r.name);
    if(byId.has(r.id)){
      const cur = byId.get(r.id);
      cur.name = r.name || cur.name;
      cur.email = r.email || cur.email;
      cur.role = r.role || cur.role || "MA";
    }else{
      payload.roster.push({ id:r.id, name:r.name, email:r.email||"", role:r.role||"MA" });
      payload.entries[r.id] = payload.entries[r.id] || blankWeekEntries();
    }
  });
}

/* ---------------- Providers ---------------- */

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

function renderProviders(payload){
  const list = $("#providerList");
  const providers = payload.providers || [];
  if(!providers.length){
    list.innerHTML = `<div class="muted">No providers loaded.</div>`;
    return;
  }

  list.innerHTML = providers
    .slice()
    .sort((a,b)=>a.name.localeCompare(b.name))
    .map(p => `
      <div class="providerRow">
        <div style="font-weight:900;color:rgba(31,41,55,.95)">${escapeHtml(p.name)}</div>
        <button class="iconBtn" title="Delete" data-prov-del="${escapeHtml(p.id)}">🗑</button>
      </div>
    `).join("");

  list.querySelectorAll("[data-prov-del]").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      const id = btn.dataset.provDel;
      payload.providers = (payload.providers||[]).filter(x=>x.id !== id);
      toast("Provider deleted.");
      renderProviders(payload);
    });
  });
}

function providerOptionsHTML(payload, selectedId){
  const providers = (payload.providers||[]).slice().sort((a,b)=>a.name.localeCompare(b.name));
  const opts = [`<option value="">— Select provider —</option>`]
    .concat(providers.map(p => `<option value="${p.id}" ${p.id===selectedId?"selected":""}>${escapeHtml(p.name)}</option>`));
  return opts.join("");
}

/* ---------------- Builder grid + Assignment modal (OFF ADDED) ---------------- */

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
      <th style="min-width:220px;text-align:left;">Team Member</th>
      ${DAYS.map(d=>`<th style="min-width:200px;text-align:left;">${d}</th>`).join("")}
    </tr>
  `;

  const sortedRoster = payload.roster.slice().sort((a,b)=>a.name.localeCompare(b.name));

  tbody.innerHTML = sortedRoster.map(r=>{
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
                ${escapeHtml(txt || "— Click to assign —")}
              </button>
            </td>
          `;
        }).join("")}
      </tr>
    `;
  }).join("");

  tbody.querySelectorAll(".cellBtn").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      openAssignmentModal(payload, btn.dataset.pid, btn.dataset.day);
    });
  });
}

function openAssignmentModal(payload, staffId, day){
  const staff = payload.roster.find(r=>r.id===staffId);
  if(!staff) return;

  const role = staff.role || "MA";
  const meta = (payload.assignmentMeta?.[staffId]?.[day]) || {};

  const isOff = !!meta.off;
  const defaultLoc = meta.location || LOCATIONS[0]?.code || "SFS";
  const defaultProv = meta.providerId || "";
  const defaultStart = meta.start || "";
  const defaultEnd = meta.end || "";

  const isMA = role === "MA";

  openModal({
    title: `${staff.name} • ${day} • ${role}`,
    body: `
      <div class="field" style="min-width:100%">
        <label>Status</label>
        <select id="asStatus">
          <option value="WORK" ${!isOff?"selected":""}>Working</option>
          <option value="OFF" ${isOff?"selected":""}>OFF</option>
        </select>
        <div class="hint">If OFF is selected, location/provider/time are not required.</div>
      </div>

      <div id="asWorkFields">
        <div class="row" style="padding:0; margin-top:12px; gap:12px">
          <div class="field grow" style="min-width:240px">
            <label>Location</label>
            <select id="asLoc">
              ${locationOptionsHTML(defaultLoc)}
            </select>
          </div>

          <div class="field grow" style="min-width:240px">
            <label>Provider ${isMA ? "(required)" : "(optional)"}</label>
            <select id="asProv">
              ${providerOptionsHTML(payload, defaultProv)}
            </select>
          </div>
        </div>

        <div class="row" style="padding:0; margin-top:12px; gap:12px">
          <div class="field" style="min-width:180px">
            <label>Start time</label>
            <input id="asStart" type="time" value="${escapeHtml(defaultStart)}" />
          </div>
          <div class="field" style="min-width:180px">
            <label>End time</label>
            <input id="asEnd" type="time" value="${escapeHtml(defaultEnd)}" />
          </div>
        </div>
      </div>
    `,
    foot: `
      <button class="btn btnGhost" data-close="1">Cancel</button>
      <button class="btn btnGhost" id="asClear">Clear</button>
      <button class="btn" id="asSave">Save</button>
    `,
    onAfter(){
      const workFields = $("#asWorkFields");
      const statusSel = $("#asStatus");

      const sync = ()=>{
        const off = statusSel.value === "OFF";
        workFields.style.display = off ? "none" : "";
      };
      statusSel.addEventListener("change", sync);
      sync();

      $("#asClear").addEventListener("click", ()=>{
        if(!payload.entries[staffId]) payload.entries[staffId] = blankWeekEntries();
        payload.entries[staffId][day] = null;

        if(!payload.assignmentMeta) payload.assignmentMeta = {};
        payload.assignmentMeta[staffId] = payload.assignmentMeta[staffId] || {};
        delete payload.assignmentMeta[staffId][day];

        closeModal();
        renderBuilderGrid(payload);
      });

      $("#asSave").addEventListener("click", ()=>{
        const isOffNow = statusSel.value === "OFF";

        if(isOffNow){
          const tid = ensureTemplate(payload, "OFF");
          if(!payload.entries[staffId]) payload.entries[staffId] = blankWeekEntries();
          payload.entries[staffId][day] = tid;

          if(!payload.assignmentMeta) payload.assignmentMeta = {};
          payload.assignmentMeta[staffId] = payload.assignmentMeta[staffId] || {};
          payload.assignmentMeta[staffId][day] = { role, off:true };

          closeModal();
          renderBuilderGrid(payload);
          return;
        }

        const location = ($("#asLoc").value||"").trim();
        const providerId = ($("#asProv").value||"").trim();
        const start = ($("#asStart").value||"").trim();
        const end = ($("#asEnd").value||"").trim();

        if(!location) return toast("Location is required.");
        if(!start || !end) return toast("Start and End time are required.");
        if(isMA && !providerId) return toast("Provider is required for MA.");

        const locObj = LOCATIONS.find(l=>l.code===location);
        const locLabel = locObj ? locObj.code : location;

        const providerName = providerId
          ? (payload.providers||[]).find(p=>p.id===providerId)?.name || ""
          : "";

        const display = buildAssignmentText({
          role,
          locationCode: locLabel,
          providerName,
          start,
          end
        });

        const tid = ensureTemplate(payload, display);

        if(!payload.entries[staffId]) payload.entries[staffId] = blankWeekEntries();
        payload.entries[staffId][day] = tid;

        if(!payload.assignmentMeta) payload.assignmentMeta = {};
        payload.assignmentMeta[staffId] = payload.assignmentMeta[staffId] || {};
        payload.assignmentMeta[staffId][day] = {
          role,
          off:false,
          location: locLabel,
          providerId: providerId || null,
          start,
          end
        };

        closeModal();
        renderBuilderGrid(payload);
      });
    }
  });
}

function buildAssignmentText({ role, locationCode, providerName, start, end }){
  const time = `${start}- ${end}`;
  if(role === "MA"){
    return `MA (${locationCode}) ${providerName} ${time}`.trim();
  }
  const prov = providerName ? `${providerName} ` : "";
  return `XR (${locationCode}) ${prov}${time}`.trim();
}

function ensureTemplate(payload, raw){
  payload.templates = payload.templates || [];
  const id = idForRaw(raw);
  if(!payload.templates.some(t=>t.id===id)){
    payload.templates.push({ id, raw, parsed: parseCell(raw) });
  }
  return id;
}

/* ---------------- Validation ---------------- */

function validate(payload){
  const issues = [];

  payload.roster.forEach(r=>{
    if(!r.email) issues.push(`${r.name} missing email`);
  });

  payload.roster.forEach(r=>{
    const e = payload.entries?.[r.id] || {};
    const metaByDay = payload.assignmentMeta?.[r.id] || {};

    DAYS.forEach(d=>{
      // OFF is valid, working must have assignment
      const meta = metaByDay[d];
      if(meta?.off) return;

      const v = e[d];
      if(v === undefined || v === null || v === ""){
        issues.push(`${r.name} missing ${d}`);
      }
    });
  });

  // MA provider required when not off
  payload.roster.forEach(r=>{
    if((r.role||"MA") !== "MA") return;
    const metaByDay = payload.assignmentMeta?.[r.id] || {};
    DAYS.forEach(d=>{
      const meta = metaByDay[d];
      if(!meta || meta.off) return;
      if(!meta.providerId){
        issues.push(`${r.name} missing Provider on ${d}`);
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

/* ---------------- Schedule CSV parsing ---------------- */

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
    entries[id] = entries[id] || blankWeekEntries();

    DAYS.forEach(d=>{
      const raw = (rows[i][dayIdx[d]]||"").trim();
      if(!raw) { entries[id][d] = null; return; }
      templateRawSet.add(raw);
      entries[id][d] = idForRaw(raw);
    });
  }

  const templates = Array.from(templateRawSet)
    .filter(Boolean)
    .map(raw => ({ id: idForRaw(raw), raw, parsed: parseCell(raw) }));

  // ensure OFF exists
  if(!templates.some(t=>t.id===idForRaw("OFF"))){
    templates.push({ id:idForRaw("OFF"), raw:"OFF", parsed: parseCell("OFF") });
  }

  return { weekOf, roster, templates, entries, providers: [], assignmentMeta: {} };
}

/* ---------------- Template display ---------------- */

function parseCell(raw){
  const s = (raw||"").trim();
  if(!s) return { type:"EMPTY", label:"", site:"", start:"", end:"" };
  if(s.toUpperCase() === "OFF") return { type:"OFF", label:"OFF", site:"", start:"", end:"" };

  const locMatch = s.match(/\(([A-Z]{2,4})\)/);
  const site = locMatch ? locMatch[1].trim() : "";

  const timeMatch = s.match(/(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})/);
  const start = timeMatch ? timeMatch[1] : "";
  const end = timeMatch ? timeMatch[2] : "";

  const upper = s.toUpperCase();
  const type = upper.startsWith("MA ") ? "MA" : upper.startsWith("XR ") ? "XR" : "ASSIGN";

  return { type, label: s, site, start, end };
}

function templateTextById(sched, id){
  const templates = (sched.templates || []);
  const t = templates.find(x=>x.id===id);
  return t ? (t.raw || "") : "";
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
