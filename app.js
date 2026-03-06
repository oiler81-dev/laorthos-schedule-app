const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const DAY_STATUSES = ["Scheduled", "OFF", "PTO", "Holiday", "Admin", "Call Only"];
const NON_SCHEDULED_STATUSES = new Set(["OFF", "PTO", "Holiday", "Admin", "Call Only"]);

const DEFAULT_TIME_PRESETS = [
  "7:50am - 4:20pm",
  "8:00am - 4:30pm",
  "8:30am - 5:00pm",
  "8:00am - 5:00pm"
];

const state = {
  weekOf: "",
  providers: [],
  staff: [],
  schedule: {},
  selectedProviderId: "",
  timePresets: [...DEFAULT_TIME_PRESETS],
  filters: {
    providerSearch: ""
  },
  modal: {
    providerId: "",
    day: ""
  }
};

const $ = (id) => document.getElementById(id);

function startOfWeek(dateStr) {
  const date = dateStr ? new Date(dateStr + "T12:00:00") : new Date();
  const day = date.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + diff);
  return date.toISOString().slice(0, 10);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function keyFor(providerId, day) {
  return `${providerId}__${day}`;
}

function toId(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "");
}

function getProviderById(providerId) {
  return state.providers.find(p => p.id === providerId) || null;
}

function getLocations(provider) {
  const set = new Set();
  if (!provider) return [];
  if (provider.location) set.add(provider.location);
  if (provider.locations) {
    provider.locations.split("|").map(v => v.trim()).filter(Boolean).forEach(v => set.add(v));
  }
  return [...set];
}

function getEntry(providerId, day) {
  const key = keyFor(providerId, day);
  if (!state.schedule[key]) {
    state.schedule[key] = {
      providerId,
      day,
      status: "Scheduled",
      maId: "",
      maName: "",
      xrtId: "",
      xrtName: "",
      location: "",
      secondaryLocation: "",
      time: state.timePresets[0] || "",
      xrRoom: "",
      notes: ""
    };
  }
  return state.schedule[key];
}

function clearStaffingFields(entry) {
  entry.maId = "";
  entry.maName = "";
  entry.xrtId = "";
  entry.xrtName = "";
  entry.location = "";
  entry.secondaryLocation = "";
  entry.time = state.timePresets[0] || "";
  entry.xrRoom = "";
  entry.notes = "";
}

function getStaffById(staffId) {
  return state.staff.find(s => s.id === staffId);
}

function isFloatStaff(staff) {
  return !!staff?.isFloat;
}

function isMAConflict(maId, day, providerId) {
  if (!maId) return false;
  const staff = getStaffById(maId);
  if (isFloatStaff(staff)) return false;

  for (const entry of Object.values(state.schedule)) {
    if (
      entry.providerId !== providerId &&
      entry.day === day &&
      entry.status === "Scheduled" &&
      entry.maId === maId
    ) {
      return true;
    }
  }
  return false;
}

function isDayComplete(entry) {
  if (!entry) return false;
  if (NON_SCHEDULED_STATUSES.has(entry.status)) return true;
  if (entry.status !== "Scheduled") return false;
  return !!(entry.maId && entry.location && entry.time);
}

function providerCompletion(providerId) {
  let complete = 0;
  let partial = 0;
  let nonScheduled = 0;

  for (const day of DAYS) {
    const entry = getEntry(providerId, day);

    if (NON_SCHEDULED_STATUSES.has(entry.status)) {
      complete++;
      nonScheduled++;
      continue;
    }

    if (isDayComplete(entry)) {
      complete++;
    } else if (entry.maId || entry.location || entry.time || entry.xrtId || entry.secondaryLocation || entry.xrRoom || entry.notes) {
      partial++;
    }
  }

  if (complete === 5) return { cls: "green", text: `Complete (${complete}/5)` };
  if (partial > 0 || complete > 0) return { cls: "yellow", text: `Partial (${complete}/5)` };
  return { cls: "red", text: "Not Started" };
}

function getWeekPayload() {
  return DAYS.flatMap(day => {
    return state.providers.map(provider => {
      const e = getEntry(provider.id, day);
      return {
        weekOf: state.weekOf,
        providerId: provider.id,
        providerName: provider.name,
        day,
        status: e.status || "Scheduled",
        maId: e.maId || "",
        maName: e.maName || "",
        xrtId: e.xrtId || "",
        xrtName: e.xrtName || "",
        location: e.location || "",
        secondaryLocation: e.secondaryLocation || "",
        time: e.time || "",
        xrRoom: e.xrRoom || "",
        notes: e.notes || ""
      };
    });
  });
}

function setSaveStatus(text, isGood = false) {
  const el = $("saveStatus");
  el.textContent = text;
  el.style.color = isGood ? "var(--good)" : "var(--muted)";
}

async function apiGet(url) {
  const res = await fetch(url, { credentials: "same-origin" });
  if (!res.ok) throw new Error(`Request failed: ${res.status}`);
  return res.json();
}

async function apiPost(url, body) {
  const res = await fetch(url, {
    method: "POST",
    credentials: "same-origin",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `Request failed: ${res.status}`);
  return data;
}

async function loadProviders() {
  const data = await apiGet("/api/providers");
  state.providers = Array.isArray(data.providers) ? data.providers : [];
  if (!state.selectedProviderId && state.providers.length) {
    state.selectedProviderId = state.providers[0].id;
  }
}

async function loadStaff() {
  const data = await apiGet("/api/roster");
  state.staff = Array.isArray(data.staff) ? data.staff : [];
}

async function loadWeek() {
  const data = await apiGet(`/api/weeks?weekOf=${encodeURIComponent(state.weekOf)}`);
  state.schedule = {};

  const items = Array.isArray(data.items) ? data.items : [];
  for (const item of items) {
    state.schedule[keyFor(item.providerId, item.day)] = {
      providerId: item.providerId,
      day: item.day,
      status: item.status || "Scheduled",
      maId: item.maId || "",
      maName: item.maName || "",
      xrtId: item.xrtId || "",
      xrtName: item.xrtName || "",
      location: item.location || "",
      secondaryLocation: item.secondaryLocation || "",
      time: item.time || state.timePresets[0] || "",
      xrRoom: item.xrRoom || "",
      notes: item.notes || ""
    };
  }

  renderAll();
  setSaveStatus(`Loaded week of ${state.weekOf}`, true);
}

function createProviderStatusDot(providerId) {
  const status = providerCompletion(providerId);
  if (status.cls === "green") return "green";
  if (status.cls === "yellow") return "yellow";
  return "red";
}

function renderProviderList() {
  const wrap = $("providerList");
  wrap.innerHTML = "";

  const q = state.filters.providerSearch.toLowerCase().trim();

  const providers = state.providers.filter(p => {
    if (!q) return true;
    return `${p.name} ${p.location || ""} ${p.specialty || ""}`.toLowerCase().includes(q);
  });

  for (const provider of providers) {
    const item = document.createElement("button");
    item.type = "button";
    item.className = `provider-item ${provider.id === state.selectedProviderId ? "active" : ""}`;

    const dot = createProviderStatusDot(provider.id);

    item.innerHTML = `
      <div class="provider-item-main">
        <div class="provider-item-name">${escapeHtml(provider.name)}</div>
        <div class="provider-item-meta">${escapeHtml([provider.location, provider.specialty].filter(Boolean).join(" • "))}</div>
      </div>
      <span class="status-dot ${dot}"></span>
    `;

    item.addEventListener("click", () => {
      state.selectedProviderId = provider.id;
      renderAll();
    });

    wrap.appendChild(item);
  }
}

function renderStaffPools() {
  const maPool = $("maPool");
  const xrtPool = $("xrtPool");
  const floatPool = $("floatPool");

  maPool.innerHTML = "";
  xrtPool.innerHTML = "";
  floatPool.innerHTML = "";

  for (const staff of state.staff) {
    const chip = document.createElement("div");
    chip.className = `staff-chip ${staff.role === "XRT" ? "xrt" : "ma"} ${staff.isFloat ? "float" : ""}`;
    chip.textContent = staff.name;

    if (staff.isFloat) {
      floatPool.appendChild(chip);
    } else if (staff.role === "XRT") {
      xrtPool.appendChild(chip);
    } else {
      maPool.appendChild(chip);
    }
  }
}

function getStatusBadgeClass(status) {
  const map = {
    "Scheduled": "badge-scheduled",
    "OFF": "badge-off",
    "PTO": "badge-pto",
    "Holiday": "badge-holiday",
    "Admin": "badge-admin",
    "Call Only": "badge-call-only"
  };
  return map[status] || "badge-scheduled";
}

function renderSelectedProvider() {
  const provider = getProviderById(state.selectedProviderId);
  const nameEl = $("selectedProviderName");
  const metaEl = $("selectedProviderMeta");
  const infoEl = $("providerCompletionInfo");
  const grid = $("selectedWeekGrid");

  if (!provider) {
    nameEl.textContent = "Select a provider";
    metaEl.textContent = "";
    infoEl.textContent = "Select a provider";
    grid.innerHTML = "";
    return;
  }

  const status = providerCompletion(provider.id);

  nameEl.textContent = provider.name;
  metaEl.textContent = [provider.location, provider.specialty].filter(Boolean).join(" • ");
  infoEl.textContent = status.text;

  grid.innerHTML = "";

  for (const day of DAYS) {
    const entry = getEntry(provider.id, day);
    const card = document.createElement("button");
    card.type = "button";
    card.className = "day-card";

    const statusClass = getStatusBadgeClass(entry.status);

    const bodyLines = [];

    if (NON_SCHEDULED_STATUSES.has(entry.status)) {
      bodyLines.push(`<div class="day-line"><strong>Status</strong>${escapeHtml(entry.status)}</div>`);
      if (entry.notes) bodyLines.push(`<div class="day-line"><strong>Notes</strong>${escapeHtml(entry.notes)}</div>`);
    } else {
      bodyLines.push(`<div class="day-line"><strong>MA</strong>${escapeHtml(entry.maName || "Not assigned")}</div>`);
      bodyLines.push(`<div class="day-line"><strong>Location</strong>${escapeHtml(entry.location || "Not selected")}</div>`);
      bodyLines.push(`<div class="day-line"><strong>Shift</strong>${escapeHtml(entry.time || "Not selected")}</div>`);
      bodyLines.push(`<div class="day-line"><strong>XRT</strong>${escapeHtml(entry.xrtName || "Optional")}</div>`);
      if (entry.xrRoom) bodyLines.push(`<div class="day-line"><strong>XR Room</strong>${escapeHtml(entry.xrRoom)}</div>`);
      if (entry.secondaryLocation) bodyLines.push(`<div class="day-line"><strong>2nd Location</strong>${escapeHtml(entry.secondaryLocation)}</div>`);
    }

    card.innerHTML = `
      <div class="day-card-head">
        <div class="day-card-name">${escapeHtml(day)}</div>
        <div class="day-card-status ${statusClass}">${escapeHtml(entry.status)}</div>
      </div>
      <div class="day-card-body">${bodyLines.join("")}</div>
    `;

    card.addEventListener("click", () => openDayModal(provider.id, day));
    grid.appendChild(card);
  }
}

function renderReviewGrid() {
  const wrap = $("reviewGridWrap");

  const html = `
    <table class="review-grid">
      <thead>
        <tr>
          <th>Provider</th>
          ${DAYS.map(day => `<th>${escapeHtml(day)}</th>`).join("")}
        </tr>
      </thead>
      <tbody>
        ${state.providers.map(provider => `
          <tr>
            <td class="review-provider">${escapeHtml(provider.name)}</td>
            ${DAYS.map(day => {
              const e = getEntry(provider.id, day);
              const badgeClass = getStatusBadgeClass(e.status);
              return `
                <td class="review-cell">
                  <div class="review-cell-box">
                    <span class="review-mini-badge ${badgeClass}">${escapeHtml(e.status)}</span>
                    ${
                      NON_SCHEDULED_STATUSES.has(e.status)
                        ? `${e.notes ? `<div>${escapeHtml(e.notes)}</div>` : `<div>—</div>`}`
                        : `
                          <div><strong>MA:</strong> ${escapeHtml(e.maName || "—")}</div>
                          <div><strong>Loc:</strong> ${escapeHtml(e.location || "—")}</div>
                          <div><strong>Time:</strong> ${escapeHtml(e.time || "—")}</div>
                          <div><strong>XRT:</strong> ${escapeHtml(e.xrtName || "—")}</div>
                        `
                    }
                  </div>
                </td>
              `;
            }).join("")}
          </tr>
        `).join("")}
      </tbody>
    </table>
  `;

  wrap.innerHTML = html;
}

function renderAll() {
  renderProviderList();
  renderStaffPools();
  renderSelectedProvider();
  renderReviewGrid();
}

function createTimeOptions(selectedValue) {
  const options = [...new Set(state.timePresets)];
  return options.map(v => `<option value="${escapeHtml(v)}" ${v === selectedValue ? "selected" : ""}>${escapeHtml(v)}</option>`).join("");
}

function createLocationOptions(provider, selectedValue) {
  const locations = getLocations(provider);
  const merged = [...new Set(["", ...locations])];
  return merged.map(v => {
    const label = v || "Select location";
    return `<option value="${escapeHtml(v)}" ${v === selectedValue ? "selected" : ""}>${escapeHtml(label)}</option>`;
  }).join("");
}

function createMAOptions(day, providerId, selectedId) {
  const options = ['<option value="">Select MA</option>'];

  const items = state.staff.filter(s => s.role === "MA");
  for (const staff of items) {
    const conflict = isMAConflict(staff.id, day, providerId);
    const disabled = conflict && staff.id !== selectedId;
    const label = disabled ? `${staff.name} — Already assigned ${day}` : staff.name;
    options.push(`<option value="${escapeHtml(staff.id)}" ${staff.id === selectedId ? "selected" : ""} ${disabled ? "disabled" : ""}>${escapeHtml(label)}</option>`);
  }

  return options.join("");
}

function createXRTOptions(selectedId) {
  const options = ['<option value="">Select XRT</option>'];
  const items = state.staff.filter(s => s.role === "XRT");
  for (const staff of items) {
    options.push(`<option value="${escapeHtml(staff.id)}" ${staff.id === selectedId ? "selected" : ""}>${escapeHtml(staff.name)}</option>`);
  }
  return options.join("");
}

function openDayModal(providerId, day) {
  state.modal.providerId = providerId;
  state.modal.day = day;

  const provider = getProviderById(providerId);
  const entry = getEntry(providerId, day);

  $("modalDayTitle").textContent = `${provider?.name || "Provider"} — ${day}`;
  $("modalDaySub").textContent = `Week of ${state.weekOf}`;

  $("modalStatus").value = entry.status || "Scheduled";
  $("modalLocation").innerHTML = createLocationOptions(provider, entry.location);
  $("modalMa").innerHTML = createMAOptions(day, providerId, entry.maId);
  $("modalTime").innerHTML = createTimeOptions(entry.time);
  $("modalXrt").innerHTML = createXRTOptions(entry.xrtId);
  $("modalXrRoom").value = entry.xrRoom || "";
  $("modalSecondaryLocation").value = entry.secondaryLocation || "";
  $("modalNotes").value = entry.notes || "";

  updateModalFieldState(entry.status);

  $("dayModal").classList.remove("hidden");
  $("dayModal").setAttribute("aria-hidden", "false");
}

function closeDayModal() {
  $("dayModal").classList.add("hidden");
  $("dayModal").setAttribute("aria-hidden", "true");
}

function updateModalFieldState(status) {
  const disabled = NON_SCHEDULED_STATUSES.has(status);

  $("modalLocation").disabled = disabled;
  $("modalMa").disabled = disabled;
  $("modalTime").disabled = disabled;
  $("modalXrt").disabled = disabled;
  $("modalXrRoom").disabled = disabled;
  $("modalSecondaryLocation").disabled = disabled;

  $("nonScheduledMessage").classList.toggle("hidden", !disabled);
}

function saveModalDay() {
  const providerId = state.modal.providerId;
  const day = state.modal.day;
  const entry = getEntry(providerId, day);

  const newStatus = $("modalStatus").value;
  const wasScheduled = entry.status === "Scheduled";
  const willBeScheduled = newStatus === "Scheduled";

  entry.status = newStatus;
  entry.notes = $("modalNotes").value.trim();

  if (!willBeScheduled) {
    clearStaffingFields(entry);
    entry.status = newStatus;
    entry.notes = $("modalNotes").value.trim();
  } else {
    const maId = $("modalMa").value;
    const ma = getStaffById(maId);
    const xrtId = $("modalXrt").value;
    const xrt = getStaffById(xrtId);

    if (maId && isMAConflict(maId, day, providerId)) {
      alert(`${ma?.name || maId} is already assigned on ${day}.`);
      return false;
    }

    entry.location = $("modalLocation").value;
    entry.maId = maId || "";
    entry.maName = ma?.name || "";
    entry.time = $("modalTime").value;
    entry.xrtId = xrtId || "";
    entry.xrtName = xrt?.name || "";
    entry.xrRoom = $("modalXrRoom").value.trim();
    entry.secondaryLocation = $("modalSecondaryLocation").value.trim();
    entry.notes = $("modalNotes").value.trim();
  }

  if (wasScheduled && !willBeScheduled) {
    // already cleared above
  }

  renderAll();
  return true;
}

function moveModalDay(step) {
  const idx = DAYS.indexOf(state.modal.day);
  if (idx === -1) return;

  const nextIdx = idx + step;
  if (nextIdx < 0 || nextIdx >= DAYS.length) {
    closeDayModal();
    return;
  }

  const ok = saveModalDay();
  if (!ok) return;

  openDayModal(state.modal.providerId, DAYS[nextIdx]);
}

function addPreset() {
  const input = $("timePresetBuilder");
  const value = input.value.trim();
  if (!value) return;
  if (!state.timePresets.includes(value)) state.timePresets.push(value);
  input.value = "";
}

function setWholeWeekStatus(providerId, status) {
  for (const day of DAYS) {
    const entry = getEntry(providerId, day);
    entry.status = status;

    if (status !== "Scheduled") {
      clearStaffingFields(entry);
      entry.status = status;
    }
  }
  renderAll();
}

function clearWholeWeek(providerId) {
  for (const day of DAYS) {
    const entry = getEntry(providerId, day);
    entry.status = "Scheduled";
    clearStaffingFields(entry);
    entry.status = "Scheduled";
  }
  renderAll();
}

async function saveWeek() {
  const payload = getWeekPayload();
  await apiPost("/api/schedule", {
    weekOf: state.weekOf,
    items: payload
  });
  setSaveStatus(`Saved week of ${state.weekOf}`, true);
}

function wireEvents() {
  $("weekPicker").addEventListener("change", async (e) => {
    state.weekOf = startOfWeek(e.target.value);
    $("weekPicker").value = state.weekOf;
    await loadWeek();
  });

  $("providerSearch").addEventListener("input", (e) => {
    state.filters.providerSearch = e.target.value;
    renderProviderList();
  });

  $("btnRefresh").addEventListener("click", async () => {
    await Promise.all([loadProviders(), loadStaff()]);
    await loadWeek();
  });

  $("btnSave").addEventListener("click", async () => {
    try {
      setSaveStatus("Saving...");
      await saveWeek();
    } catch (err) {
      console.error(err);
      setSaveStatus("Save failed");
      alert(err.message || "Unable to save schedule.");
    }
  });

  $("btnAddPreset").addEventListener("click", () => {
    addPreset();
  });

  $("timePresetBuilder").addEventListener("keydown", (e) => {
    if (e.key === "Enter") addPreset();
  });

  $("btnWeekScheduled").addEventListener("click", () => {
    if (!state.selectedProviderId) return;
    setWholeWeekStatus(state.selectedProviderId, "Scheduled");
  });

  $("btnWeekOff").addEventListener("click", () => {
    if (!state.selectedProviderId) return;
    setWholeWeekStatus(state.selectedProviderId, "OFF");
  });

  $("btnWeekPto").addEventListener("click", () => {
    if (!state.selectedProviderId) return;
    setWholeWeekStatus(state.selectedProviderId, "PTO");
  });

  $("btnWeekHoliday").addEventListener("click", () => {
    if (!state.selectedProviderId) return;
    setWholeWeekStatus(state.selectedProviderId, "Holiday");
  });

  $("btnWeekAdmin").addEventListener("click", () => {
    if (!state.selectedProviderId) return;
    setWholeWeekStatus(state.selectedProviderId, "Admin");
  });

  $("btnWeekClear").addEventListener("click", () => {
    if (!state.selectedProviderId) return;
    clearWholeWeek(state.selectedProviderId);
  });

  $("btnCloseModal").addEventListener("click", closeDayModal);
  $("modalBackdrop").addEventListener("click", closeDayModal);

  $("modalStatus").addEventListener("change", (e) => {
    updateModalFieldState(e.target.value);
  });

  $("btnSaveDay").addEventListener("click", () => {
    const ok = saveModalDay();
    if (!ok) return;
    closeDayModal();
  });

  $("btnPrevDay").addEventListener("click", () => moveModalDay(-1));
  $("btnNextDay").addEventListener("click", () => moveModalDay(1));

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeDayModal();
  });
}

async function init() {
  state.weekOf = startOfWeek();
  $("weekPicker").value = state.weekOf;

  wireEvents();

  try {
    await Promise.all([loadProviders(), loadStaff()]);
    await loadWeek();
  } catch (err) {
    console.error(err);
    setSaveStatus("Failed to load");
    alert("Could not load scheduler data. Check API settings and storage connection.");
  }
}

init();
