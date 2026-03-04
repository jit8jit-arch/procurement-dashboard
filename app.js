/* International-standard improvements:
   - Accessible UI (aria, keyboard-friendly modal, skip link support)
   - Better filtering:
       * Search, approval, multi-section, date range, overdue toggle
   - Column chooser
   - Project details saved (more fields + metadata)
   - Backup/Restore JSON for portability
   - Exports respect current filtered view (Excel + PDF)
   - Double-click row to edit
*/

const LS_KEY = "proc_dash_v2";

// Auto colors for sections
const COLOR_PALETTE = [
  "#22c55e","#3b82f6","#f59e0b","#ef4444","#a855f7","#14b8a6",
  "#e11d48","#84cc16","#06b6d4","#f97316","#8b5cf6","#10b981"
];

// Suggested terms (general, no project-specific samples)
const GENERAL_DESC_TERMS = [
  "Pipes & Fittings",
  "Valves",
  "Pressure Gauge",
  "Pumps",
  "Equipment",
  "Mechanical Equipment",
  "Electrical Panel",
  "Cables",
  "Fixtures",
  "Civil Works"
];

const GENERAL_SECTION_TERMS = [
  "Procurement",
  "Mechanical",
  "Electrical",
  "Civil",
  "Equipment"
];

const DEFAULT_COLUMNS = [
  { key:"sno",      label:"#", show:true },
  { key:"ref",      label:"Submittal Ref", show:true },
  { key:"rev",      label:"Rev", show:true },
  { key:"desc",     label:"Description of Material", show:true, wrap:true },
  { key:"mfg",      label:"Manufacturer / Supplier", show:true },
  { key:"approval", label:"Approval Status", show:true },
  { key:"planned",  label:"Planned Submission Date", show:true },
  { key:"estQty",   label:"Est Qty", show:true },
  { key:"ordQty",   label:"Ordered Qty", show:true },
  { key:"prNo",     label:"PR Number", show:true },
  { key:"prDate",   label:"PR Date", show:true },
  { key:"prStatus", label:"PR Status", show:true },
  { key:"lpoDate",  label:"LPO Issue Date", show:true },
  { key:"lpoNo",    label:"LPO Number", show:true },
  { key:"payment",  label:"Payment Status", show:true },
  { key:"section",  label:"Section", show:true },
  { key:"actions",  label:"", show:true }
];

const $ = (id) => document.getElementById(id);

let state = loadState();
let activeProjectId = state.activeProjectId || null;
let activeSectionChip = "All"; // chip selection
let projectSortMode = state.projectSortMode || "recent"; // recent | name

init();

function init(){
  populateSuggestions();
  bindUI();
  closeModal();
  render();
}

function populateSuggestions(){
  $("descSuggestions").innerHTML = GENERAL_DESC_TERMS.map(t => `<option value="${escAttr(t)}"></option>`).join("");
  $("sectionSuggestions").innerHTML = GENERAL_SECTION_TERMS.map(t => `<option value="${escAttr(t)}"></option>`).join("");
}

function bindUI(){
  $("btnAddProject").onclick = () => openProjectModal();
  $("btnAddProject2").onclick = () => openProjectModal();

  $("projectSearch").oninput = renderProjectList;
  $("btnSortProjects").onclick = () => {
    projectSortMode = (projectSortMode === "recent") ? "name" : "recent";
    state.projectSortMode = projectSortMode;
    saveState();
    renderProjectList();
    toast("Sort", projectSortMode === "recent" ? "Sorted by recent." : "Sorted by name.");
  };

  $("btnClearAll").onclick = () => confirmClearAll();

  $("btnEditProject").onclick = () => {
    const p = getActiveProject();
    if (p) openProjectModal(p);
  };

  $("btnDeleteProject").onclick = () => deleteActiveProject();
  $("btnAddSection").onclick = () => openSectionModal();
  $("btnAddItem").onclick = () => openItemModal();

  // Filters
  $("itemSearch").oninput = () => renderItems();
  $("statusFilter").onchange = () => renderItems();
  $("sectionMulti").onchange = () => renderItems();
  $("plannedFrom").onchange = () => renderItems();
  $("plannedTo").onchange = () => renderItems();
  $("onlyOverdue").onchange = () => renderItems();
  $("sortBy").onchange = () => renderItems();

  $("btnResetFilters").onclick = () => resetFilters();
  $("btnColumns").onclick = () => openColumnsModal();

  $("btnExportExcel").onclick = exportExcel;
  $("btnExportPDF").onclick = exportPDF;

  $("excelImport").addEventListener("change", handleExcelImport);

  // Backup / restore
  $("btnBackup").onclick = backupJSON;
  $("backupImport").addEventListener("change", restoreJSON);

  // Notifications
  $("btnEnableNoti").onclick = enableNotifications;

  // Modal close
  $("modalClose").onclick = closeModal;
  $("modalBackdrop").onclick = (e) => { if (e.target === $("modalBackdrop")) closeModal(); };
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !$("modalBackdrop").classList.contains("hidden")) closeModal();
  });
}

/* -------------------- Render -------------------- */

function render(){
  renderProjectList();
  renderProjectView();
  renderNotifications();
}

function renderProjectList(){
  const list = $("projectList");
  list.innerHTML = "";
  const q = ($("projectSearch").value || "").trim().toLowerCase();

  let projects = state.projects.slice();
  projects.sort((a,b) => {
    if (projectSortMode === "name"){
      return (a.name||"").localeCompare((b.name||""), undefined, { sensitivity:"base" });
    }
    return (b.createdAt||"").localeCompare(a.createdAt||"");
  });

  if (q){
    projects = projects.filter(p =>
      (p.name||"").toLowerCase().includes(q) ||
      (p.refNo||"").toLowerCase().includes(q)
    );
  }

  for (const p of projects){
    const div = document.createElement("div");
    div.className = "card" + (p.id === activeProjectId ? " active" : "");
    div.setAttribute("role","listitem");
    div.tabIndex = 0;

    div.onclick = () => selectProject(p.id);
    div.onkeydown = (e) => { if (e.key === "Enter" || e.key === " ") selectProject(p.id); };

    div.innerHTML = `
      <div class="cardTitle">${esc(p.name || "Untitled")}</div>
      <div class="cardMeta">
        <span class="tag">Ref: ${esc(p.refNo || "—")}</span>
        ${p.client ? `<span class="tag">Client: ${esc(p.client)}</span>` : ""}
        ${p.consultant ? `<span class="tag">Consultant: ${esc(p.consultant)}</span>` : ""}
      </div>
    `;
    list.appendChild(div);
  }
}

function selectProject(id){
  activeProjectId = id;
  state.activeProjectId = id;
  activeSectionChip = "All";
  saveState();
  render();
}

function renderProjectView(){
  const p = getActiveProject();
  const hasActive = !!p;

  $("emptyState").classList.toggle("hidden", hasActive);
  $("projectView").classList.toggle("hidden", !hasActive);
  if (!hasActive) return;

  $("pName").textContent = p.name || "—";
  $("pRef").textContent = `Ref: ${p.refNo || "—"}`;
  $("pClient").textContent = `Client: ${p.client || "—"}`;
  $("pConsultant").textContent = `Consultant: ${p.consultant || "—"}`;
  $("pContractor").textContent = `Main Contractor: ${p.contractor || "—"}`;
  $("pLocation").textContent = `Location: ${p.location || "—"}`;

  renderSections();
  renderSectionMultiFilter();
  renderTableHeader();
  renderItems();
}

function renderSections(){
  const p = getActiveProject();
  const wrap = $("sectionChips");
  wrap.innerHTML = "";

  const sections = normalizeSections(p);
  for (const s of sections){
    const chip = document.createElement("div");
    chip.className = "chip" + (s.name === activeSectionChip ? " active" : "");
    chip.onclick = () => {
      activeSectionChip = s.name;
      // set multi-select to that single section for ease
      setSectionMultiSelection(s.name === "All" ? [] : [s.name]);
      renderItems();
      renderNotifications();
    };

    chip.innerHTML = `
      <span class="dot" style="background:${s.color}"></span>
      <span>${esc(s.name)}</span>
      ${(!s.locked && s.name !== "All")
        ? `<button class="iconBtn" title="Remove section" style="width:28px;height:28px" data-sec="${escAttr(s.name)}">🗑️</button>`
        : ""}
    `;
    wrap.appendChild(chip);

    const delBtn = chip.querySelector("button[data-sec]");
    if (delBtn){
      delBtn.onclick = (e) => {
        e.stopPropagation();
        removeSection(delBtn.getAttribute("data-sec"));
      };
    }
  }
}

function renderSectionMultiFilter(){
  const p = getActiveProject();
  const sel = $("sectionMulti");
  const sections = normalizeSections(p).filter(s => s.name !== "All").map(s => s.name);

  const existing = getSelectedMulti(sel);

  sel.innerHTML = sections.map(s => `<option value="${escAttr(s)}">${esc(s)}</option>`).join("");

  // keep selection if possible
  setSectionMultiSelection(existing.filter(x => sections.includes(x)));
}

function renderTableHeader(){
  const p = getActiveProject();
  if (!p.columns) p.columns = structuredClone(DEFAULT_COLUMNS);

  const headRow = $("tableHeadRow");
  headRow.innerHTML = "";

  for (const col of p.columns){
    if (!col.show) continue;
    const th = document.createElement("th");
    th.textContent = col.label;
    headRow.appendChild(th);
  }
}

function renderItems(){
  const p = getActiveProject();
  if (!p) return;

  const items = getFilteredItemsForView(p);
  $("resultCount").textContent = `${items.length} item(s)`;

  const sections = normalizeSections(p);
  const secColor = new Map(sections.map(s => [s.name, s.color]));

  const body = $("itemsBody");
  body.innerHTML = "";

  const cols = (p.columns || DEFAULT_COLUMNS).filter(c => c.show);

  items.forEach((it, idx) => {
    const tr = document.createElement("tr");
    tr.ondblclick = () => openItemModal(it);

    for (const col of cols){
      const td = document.createElement("td");

      if (col.key === "actions"){
        td.innerHTML = `
          <div class="rowActions">
            <button class="iconBtn" title="Edit" data-edit="${it.id}">✏️</button>
            <button class="iconBtn" title="Delete" data-del="${it.id}">🗑️</button>
          </div>
        `;
      } else if (col.key === "approval"){
        td.innerHTML = statusBadge(it.approval);
      } else if (col.key === "planned"){
        td.innerHTML = plannedBadgeInfo(it.planned, it.approval);
      } else if (col.key === "section"){
        const sname = it.section || "All";
        const color = secColor.get(sname) || "#94a3b8";
        td.innerHTML = `
          <span class="badge" title="Section">
            <span class="dot" style="background:${color}"></span>
            ${esc(sname)}
          </span>
        `;
      } else if (col.key === "desc"){
        td.className = "wrap";
        td.textContent = it.desc || "";
      } else if (col.key === "sno"){
        td.textContent = it.sno ?? (idx + 1);
      } else {
        td.textContent = (it[col.key] ?? "");
      }

      tr.appendChild(td);
    }

    body.appendChild(tr);
  });

  // row action buttons
  body.querySelectorAll("button[data-edit]").forEach(btn => {
    btn.onclick = () => {
      const id = btn.getAttribute("data-edit");
      const item = p.items.find(x => x.id === id);
      if (item) openItemModal(item);
    };
  });

  body.querySelectorAll("button[data-del]").forEach(btn => {
    btn.onclick = () => deleteItem(btn.getAttribute("data-del"));
  });

  renderNotifications();
}

/* -------------------- Filtering -------------------- */

function getFilteredItemsForView(p){
  const q = ($("itemSearch").value || "").trim().toLowerCase();
  const status = ($("statusFilter").value || "").toUpperCase();
  const sortMode = $("sortBy").value || "desc_asc";
  const sectionsSelected = getSelectedMulti($("sectionMulti"));
  const from = $("plannedFrom").value || "";
  const to = $("plannedTo").value || "";
  const onlyOverdue = $("onlyOverdue").checked;

  let items = (p.items || []).slice();

  // chip selection as a quick filter helper (does not override multi-select; it complements)
  if (activeSectionChip && activeSectionChip !== "All"){
    // if multi is empty, follow chip; if multi has selection, keep multi as the primary
    if (sectionsSelected.length === 0){
      items = items.filter(it => (it.section || "All") === activeSectionChip);
    }
  }

  // multi section filter
  if (sectionsSelected.length){
    items = items.filter(it => sectionsSelected.includes((it.section || "All")));
  }

  // text search
  if (q){
    items = items.filter(it => {
      const hay = [
        it.desc, it.ref, it.mfg, it.approval,
        it.prNo, it.lpoNo, it.payment, it.prStatus
      ].join(" ").toLowerCase();
      return hay.includes(q);
    });
  }

  // status filter
  if (status){
    items = items.filter(it => (it.approval || "").toUpperCase() === status);
  }

  // planned date range filter
  if (from){
    items = items.filter(it => (it.planned || "") >= from);
  }
  if (to){
    items = items.filter(it => (it.planned || "") <= to);
  }

  // overdue filter (only when not approved)
  if (onlyOverdue){
    items = items.filter(it => isOverdue(it.planned, it.approval));
  }

  // sorting
  items.sort((a,b) => sortCompare(a,b,sortMode));
  return items;
}

function resetFilters(){
  $("itemSearch").value = "";
  $("statusFilter").value = "";
  $("plannedFrom").value = "";
  $("plannedTo").value = "";
  $("onlyOverdue").checked = false;
  setSectionMultiSelection([]);
  activeSectionChip = "All";
  renderSections();
  renderItems();
  toast("Filters", "Filters reset.");
}

/* -------------------- Columns -------------------- */

function openColumnsModal(){
  const p = getActiveProject();
  if (!p) return;
  if (!p.columns) p.columns = structuredClone(DEFAULT_COLUMNS);

  openModal("Choose Columns");
  const rows = p.columns
    .filter(c => c.key !== "actions")
    .map(c => `
      <label style="display:flex;align-items:center;gap:10px;padding:6px 0">
        <input type="checkbox" data-col="${escAttr(c.key)}" ${c.show ? "checked" : ""} />
        <span>${esc(c.label)}</span>
      </label>
    `).join("");

  $("modalBody").innerHTML = `
    <div class="muted" style="margin-bottom:10px">Choose which columns to display (saved per project).</div>
    <div>${rows}</div>
  `;
  $("modalFoot").innerHTML = `
    <button class="btn" id="m_cancel">Cancel</button>
    <button class="btn primary" id="m_save">Save</button>
  `;

  $("m_cancel").onclick = closeModal;
  $("m_save").onclick = () => {
    const checks = $("modalBody").querySelectorAll("input[type=checkbox][data-col]");
    const wanted = new Map();
    checks.forEach(ch => wanted.set(ch.getAttribute("data-col"), ch.checked));

    p.columns = p.columns.map(c => {
      if (wanted.has(c.key)) return { ...c, show: wanted.get(c.key) };
      return c;
    });

    saveState();
    closeModal();
    renderTableHeader();
    renderItems();
    toast("Saved", "Columns updated.");
  };
}

/* -------------------- Project Modals -------------------- */

function openProjectModal(existing=null){
  openModal(existing ? "Edit Project" : "New Project");

  const p = existing || {
    name:"",
    refNo:"",
    client:"",
    consultant:"",
    contractor:"",
    location:"",
    notes:""
  };

  $("modalBody").innerHTML = `
    <div class="grid">
      <div class="field full">
        <label>Project Name (required)</label>
        <input id="m_pName" value="${escAttr(p.name)}" />
      </div>

      <div class="field">
        <label>Project Reference No. (required)</label>
        <input id="m_pRef" value="${escAttr(p.refNo)}" />
      </div>

      <div class="field">
        <label>Client (optional)</label>
        <input id="m_pClient" value="${escAttr(p.client)}" />
      </div>

      <div class="field">
        <label>Consultant (optional)</label>
        <input id="m_pConsultant" value="${escAttr(p.consultant)}" />
      </div>

      <div class="field">
        <label>Main Contractor (optional)</label>
        <input id="m_pContractor" value="${escAttr(p.contractor)}" />
      </div>

      <div class="field">
        <label>Location (optional)</label>
        <input id="m_pLocation" value="${escAttr(p.location)}" />
      </div>

      <div class="field full">
        <label>Notes (optional)</label>
        <input id="m_pNotes" value="${escAttr(p.notes)}" />
      </div>
    </div>
  `;

  $("modalFoot").innerHTML = `
    <button class="btn" id="m_cancel">Cancel</button>
    <button class="btn primary" id="m_save">${existing ? "Save Changes" : "Create Project"}</button>
  `;

  $("m_cancel").onclick = closeModal;
  $("m_save").onclick = () => {
    const name = $("m_pName").value.trim();
    const refNo = $("m_pRef").value.trim();

    if (!name) return toast("Missing", "Project Name is required.");
    if (!refNo) return toast("Missing", "Project Reference No. is required.");

    if (existing){
      existing.name = name;
      existing.refNo = refNo;
      existing.client = $("m_pClient").value.trim();
      existing.consultant = $("m_pConsultant").value.trim();
      existing.contractor = $("m_pContractor").value.trim();
      existing.location = $("m_pLocation").value.trim();
      existing.notes = $("m_pNotes").value.trim();
      existing.updatedAt = nowISO();
    } else {
      const np = {
        id: uid(),
        name,
        refNo,
        client: $("m_pClient").value.trim(),
        consultant: $("m_pConsultant").value.trim(),
        contractor: $("m_pContractor").value.trim(),
        location: $("m_pLocation").value.trim(),
        notes: $("m_pNotes").value.trim(),
        createdAt: nowISO(),
        updatedAt: nowISO(),
        sections: [{ name:"All", color:"#94a3b8", locked:true }],
        items: [],
        columns: structuredClone(DEFAULT_COLUMNS)
      };
      state.projects.push(np);
      activeProjectId = np.id;
      state.activeProjectId = activeProjectId;
      activeSectionChip = "All";
    }

    saveState();
    closeModal();
    render();
    toast("Saved", "Project saved.");
  };
}

function deleteActiveProject(){
  const p = getActiveProject();
  if (!p) return;

  openModal("Delete Project?");
  $("modalBody").innerHTML = `
    <div style="line-height:1.5">
      This will permanently remove:
      <ul>
        <li><b>${esc(p.name)}</b></li>
        <li>${(p.items||[]).length} item(s)</li>
        <li>All sections/categories</li>
      </ul>
      <div class="muted">This only affects your browser data.</div>
    </div>
  `;
  $("modalFoot").innerHTML = `
    <button class="btn" id="m_cancel">Cancel</button>
    <button class="btn danger" id="m_del">Delete</button>
  `;
  $("m_cancel").onclick = closeModal;
  $("m_del").onclick = () => {
    state.projects = state.projects.filter(x => x.id !== p.id);
    activeProjectId = state.projects[0]?.id || null;
    state.activeProjectId = activeProjectId;
    activeSectionChip = "All";
    saveState();
    closeModal();
    render();
    toast("Deleted", "Project removed.");
  };
}

function confirmClearAll(){
  openModal("Clear All Data?");
  $("modalBody").innerHTML = `
    <div style="line-height:1.5">
      This will delete <b>all projects</b> saved in this browser.
      <div class="muted" style="margin-top:10px">Use Backup first if you want to keep a copy.</div>
    </div>
  `;
  $("modalFoot").innerHTML = `
    <button class="btn" id="m_cancel">Cancel</button>
    <button class="btn danger" id="m_clear">Clear</button>
  `;
  $("m_cancel").onclick = closeModal;
  $("m_clear").onclick = () => {
    localStorage.removeItem(LS_KEY);
    state = { projects: [], activeProjectId:null, projectSortMode:"recent" };
    activeProjectId = null;
    activeSectionChip = "All";
    closeModal();
    render();
    toast("Cleared", "All local data removed.");
  };
}

/* -------------------- Sections -------------------- */

function openSectionModal(){
  const p = getActiveProject();
  if (!p) return;

  openModal("Add Section / Category");
  $("modalBody").innerHTML = `
    <div class="grid">
      <div class="field full">
        <label>Section Name</label>
        <input id="m_sName" list="sectionSuggestions" placeholder="e.g., Mechanical, Electrical, Procurement..." />
      </div>
    </div>
    <div class="muted" style="margin-top:10px;font-size:12px">Color will be assigned automatically.</div>
  `;

  $("modalFoot").innerHTML = `
    <button class="btn" id="m_cancel">Cancel</button>
    <button class="btn primary" id="m_add">Add Section</button>
  `;

  $("m_cancel").onclick = closeModal;
  $("m_add").onclick = () => {
    const name = $("m_sName").value.trim();
    if (!name) return toast("Missing", "Section name is required.");

    const sections = normalizeSections(p);
    if (sections.some(s => s.name.toLowerCase() === name.toLowerCase())){
      return toast("Duplicate", "This section already exists.");
    }

    p.sections = sections.filter(s => s.name !== "All");
    p.sections.push({ name, color: colorFor(name), locked:false });
    p.updatedAt = nowISO();

    saveState();
    closeModal();
    activeSectionChip = name;
    setSectionMultiSelection([name]);
    render();
    toast("Added", `Section "${name}" created.`);
  };
}

function removeSection(sectionName){
  const p = getActiveProject();
  if (!p) return;

  const sections = normalizeSections(p);
  const s = sections.find(x => x.name === sectionName);
  if (!s || s.locked) return;

  (p.items || []).forEach(it => {
    if ((it.section||"") === sectionName) it.section = "All";
  });

  p.sections = sections.filter(x => x.name !== "All" && x.name !== sectionName);
  p.updatedAt = nowISO();

  if (activeSectionChip === sectionName) activeSectionChip = "All";
  setSectionMultiSelection([]);

  saveState();
  render();
  toast("Updated", `Section "${sectionName}" removed (items moved to All).`);
}

function normalizeSections(p){
  const raw = (p.sections || []).slice();
  const hasAll = raw.some(s => s.name === "All");
  const sections = hasAll ? raw : [{ name:"All", color:"#94a3b8", locked:true }, ...raw];

  sections.forEach(s => {
    if (!s.color) s.color = s.name === "All" ? "#94a3b8" : colorFor(s.name);
  });

  const all = sections.find(s => s.name === "All");
  if (all) all.locked = true;

  const seen = new Set();
  return sections.filter(s => {
    const key = (s.name || "").toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function ensureSection(p, name){
  if (!name || name === "All") return;
  const sections = normalizeSections(p);
  if (sections.some(s => s.name === name)) return;
  p.sections = sections.filter(s => s.name !== "All");
  p.sections.push({ name, color: colorFor(name), locked:false });
}

/* -------------------- Items -------------------- */

function openItemModal(existing=null){
  const p = getActiveProject();
  if (!p) return;

  const isEdit = !!existing;
  const it = existing || {
    id: uid(),
    sno:"",
    ref:"",
    rev:"",
    desc:"",
    mfg:"",
    approval:"PENDING",
    planned:"",
    estQty:"",
    ordQty:"",
    prNo:"",
    prDate:"",
    prStatus:"",
    lpoDate:"",
    lpoNo:"",
    payment:"",
    section: pickDefaultSection(p)
  };

  openModal(isEdit ? "Edit Item" : "Add Item");

  $("modalBody").innerHTML = `
    <div class="grid">
      <div class="field">
        <label>S. No</label>
        <input id="m_sno" value="${escAttr(it.sno ?? "")}" />
      </div>

      <div class="field">
        <label>Material Submittal Ref No.</label>
        <input id="m_ref" value="${escAttr(it.ref ?? "")}" />
      </div>

      <div class="field">
        <label>Rev</label>
        <input id="m_rev" value="${escAttr(it.rev ?? "")}" />
      </div>

      <div class="field">
        <label>Approval Status</label>
        <select id="m_approval">
          ${["APPROVED","PENDING","REJECTED","RESUBMIT"].map(x => `<option ${String(it.approval||"").toUpperCase()===x?"selected":""}>${x}</option>`).join("")}
        </select>
      </div>

      <div class="field full">
        <label>Description of Material</label>
        <input id="m_desc" list="descSuggestions" value="${escAttr(it.desc ?? "")}" placeholder="e.g., Pipes & Fittings" />
      </div>

      <div class="field full">
        <label>Manufacturer / Supplier</label>
        <input id="m_mfg" value="${escAttr(it.mfg ?? "")}" />
      </div>

      <div class="field">
        <label>Planned Date of Submission</label>
        <input id="m_planned" type="date" value="${escAttr(it.planned ?? "")}" />
      </div>

      <div class="field">
        <label>Section / Category</label>
        <select id="m_section">
          <option ${it.section==="All"?"selected":""}>All</option>
          ${normalizeSections(p).filter(s => s.name!=="All").map(s => `<option ${it.section===s.name?"selected":""}>${esc(s.name)}</option>`).join("")}
        </select>
      </div>

      <div class="field">
        <label>Est Qty</label>
        <input id="m_est" value="${escAttr(it.estQty ?? "")}" />
      </div>

      <div class="field">
        <label>Ordered Qty</label>
        <input id="m_ord" value="${escAttr(it.ordQty ?? "")}" />
      </div>

      <div class="field">
        <label>PR Number</label>
        <input id="m_prno" value="${escAttr(it.prNo ?? "")}" />
      </div>

      <div class="field">
        <label>PR Date</label>
        <input id="m_prdate" type="date" value="${escAttr(it.prDate ?? "")}" />
      </div>

      <div class="field full">
        <label>PR Status</label>
        <input id="m_prstatus" value="${escAttr(it.prStatus ?? "")}" />
      </div>

      <div class="field">
        <label>LPO Issue Date</label>
        <input id="m_lpodate" type="date" value="${escAttr(it.lpoDate ?? "")}" />
      </div>

      <div class="field">
        <label>LPO Number</label>
        <input id="m_lpono" value="${escAttr(it.lpoNo ?? "")}" />
      </div>

      <div class="field full">
        <label>Payment Status</label>
        <input id="m_payment" value="${escAttr(it.payment ?? "")}" />
      </div>
    </div>
  `;

  $("modalFoot").innerHTML = `
    <button class="btn" id="m_cancel">Cancel</button>
    <button class="btn primary" id="m_save">${isEdit ? "Save Changes" : "Add Item"}</button>
  `;

  $("m_cancel").onclick = closeModal;
  $("m_save").onclick = () => {
    const desc = $("m_desc").value.trim();
    if (!desc) return toast("Missing", "Description is required.");

    it.sno = $("m_sno").value.trim();
    it.ref = $("m_ref").value.trim();
    it.rev = $("m_rev").value.trim();
    it.desc = desc;
    it.mfg = $("m_mfg").value.trim();
    it.approval = $("m_approval").value.trim();
    it.planned = $("m_planned").value;
    it.estQty = $("m_est").value.trim();
    it.ordQty = $("m_ord").value.trim();
    it.prNo = $("m_prno").value.trim();
    it.prDate = $("m_prdate").value;
    it.prStatus = $("m_prstatus").value.trim();
    it.lpoDate = $("m_lpodate").value;
    it.lpoNo = $("m_lpono").value.trim();
    it.payment = $("m_payment").value.trim();
    it.section = $("m_section").value;

    ensureSection(p, it.section);

    if (isEdit){
      const idx = p.items.findIndex(x => x.id === it.id);
      if (idx >= 0) p.items[idx] = it;
      toast("Saved", "Item updated.");
    } else {
      p.items.push(it);
      toast("Added", "Item created.");
    }

    p.updatedAt = nowISO();
    saveState();
    closeModal();
    renderSectionMultiFilter();
    renderSections();
    renderTableHeader();
    renderItems();
  };
}

function pickDefaultSection(p){
  const selected = getSelectedMulti($("sectionMulti"));
  if (selected.length === 1) return selected[0];
  if (activeSectionChip && activeSectionChip !== "All") return activeSectionChip;
  return "All";
}

function deleteItem(itemId){
  const p = getActiveProject();
  if (!p) return;

  p.items = (p.items || []).filter(x => x.id !== itemId);
  p.updatedAt = nowISO();
  saveState();
  renderItems();
  toast("Deleted", "Item removed.");
}

/* -------------------- Notifications -------------------- */

async function enableNotifications(){
  if (!("Notification" in window)) {
    toast("Notifications", "Your browser doesn’t support notifications.");
    return;
  }
  const res = await Notification.requestPermission();
  toast("Notifications", res === "granted" ? "Enabled." : "Not enabled.");
  renderNotifications();
}

function renderNotifications(){
  const p = getActiveProject();
  const panel = $("notiPanel");
  if (!p){ panel.classList.add("hidden"); return; }

  const upcoming = getUpcoming(p);
  if (!upcoming.length){
    panel.classList.add("hidden");
    return;
  }

  panel.classList.remove("hidden");
  panel.innerHTML = `
    <div class="notiTitle">Upcoming planned submissions (next 7 days)</div>
    <div class="notiList">
      ${upcoming.slice(0,6).map(u => `
        <div class="notiItem">
          <b>${esc(u.item.desc || "Item")}</b>
          <span class="muted"> — ${esc(u.item.planned)} (${u.daysLeft} day(s))</span>
          <span class="muted"> • ${esc(u.item.section || "—")}</span>
        </div>
      `).join("")}
      ${upcoming.length > 6 ? `<div class="muted">+${upcoming.length-6} more…</div>` : ""}
    </div>
  `;

  if ("Notification" in window && Notification.permission === "granted"){
    upcoming.forEach(u => {
      if (u.daysLeft <= 1 && (u.item.approval || "").toUpperCase() !== "APPROVED"){
        maybeNotifyOnce(`due_${u.item.id}`, `${p.name}: Due soon`, `${u.item.desc} planned on ${u.item.planned}`);
      }
    });
  }
}

function getUpcoming(p){
  const items = (p.items || []).filter(it => it.planned && (it.approval||"").toUpperCase() !== "APPROVED");
  const today = new Date(); today.setHours(0,0,0,0);

  return items.map(it => {
    const d = new Date(it.planned + "T00:00:00");
    const daysLeft = Math.round((d - today) / (1000*60*60*24));
    return { item: it, daysLeft };
  }).filter(x => x.daysLeft <= 7).sort((a,b)=> a.daysLeft - b.daysLeft);
}

function maybeNotifyOnce(key, title, body){
  const k = `noti_once_${key}`;
  if (sessionStorage.getItem(k)) return;
  sessionStorage.setItem(k, "1");
  new Notification(title, { body });
}

/* -------------------- Import/Export -------------------- */

async function handleExcelImport(e){
  const file = e.target.files?.[0];
  e.target.value = "";
  if (!file) return;

  try{
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type:"array" });
    const sheetName = wb.SheetNames.includes("Log") ? "Log" : wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });

    const headerRowIndex = rows.findIndex(r =>
      r.some(cell => String(cell||"").toUpperCase().includes("DESCRIPTION OF MATERIAL"))
    );
    if (headerRowIndex < 0){
      toast("Import failed", "Couldn’t find the procurement header row (needs 'Description of Material').");
      return;
    }

    const p = getActiveProject();
    if (!p) return;

    const start = headerRowIndex + 1;
    const imported = [];

    for (let i=start; i<rows.length; i++){
      const r = rows[i];
      const desc = (r?.[3] || "").toString().trim();
      if (!desc) continue;

      imported.push({
        id: uid(),
        sno: r?.[0] ?? "",
        ref: r?.[1] ?? "",
        rev: r?.[2] ?? "",
        desc,
        mfg: r?.[4] ?? "",
        approval: (r?.[5] ?? "PENDING").toString().toUpperCase(),
        planned: toISODate(r?.[6]),
        estQty: r?.[7] ?? "",
        ordQty: r?.[8] ?? "",
        prNo: r?.[9] ?? "",
        prDate: toISODate(r?.[10]),
        prStatus: r?.[11] ?? "",
        lpoDate: toISODate(r?.[12]),
        lpoNo: r?.[13] ?? "",
        payment: r?.[14] ?? "",
        section: pickDefaultSection(p) || "All"
      });
    }

    if (!imported.length){
      toast("Import", "No item rows found to import.");
      return;
    }

    imported.forEach(it => ensureSection(p, it.section));
    p.items = (p.items || []).concat(imported);
    p.updatedAt = nowISO();

    saveState();
    renderSectionMultiFilter();
    renderSections();
    renderTableHeader();
    renderItems();
    toast("Imported", `${imported.length} item(s) imported.`);
  } catch(err){
    console.error(err);
    toast("Import failed", "Please try again with the Excel file.");
  }
}

function exportExcel(){
  const p = getActiveProject();
  if (!p) return;

  const items = getFilteredItemsForView(p);
  const cols = (p.columns || DEFAULT_COLUMNS).filter(c => c.show && c.key !== "actions");

  const header = cols.map(c => c.label);
  const aoa = [header];

  items.forEach((it, idx) => {
    aoa.push(cols.map(c => {
      if (c.key === "sno") return it.sno ?? (idx + 1);
      return it[c.key] ?? "";
    }));
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Export");

  const filename = safeFileName(`${p.name}_${p.refNo}_export.xlsx`);
  XLSX.writeFile(wb, filename);
  toast("Export", "Excel downloaded.");
}

function exportPDF(){
  const p = getActiveProject();
  if (!p) return;

  const items = getFilteredItemsForView(p);
  const cols = (p.columns || DEFAULT_COLUMNS).filter(c => c.show && c.key !== "actions");

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation:"landscape", unit:"pt", format:"a4" });

  doc.setFontSize(14);
  doc.text(`${p.name} — Procurement Log`, 40, 40);

  doc.setFontSize(10);
  const meta = `Ref: ${p.refNo}   Client: ${p.client || "—"}   Consultant: ${p.consultant || "—"}   Generated: ${new Date().toLocaleString()}`;
  doc.text(meta, 40, 58);

  const head = [cols.map(c => c.label)];
  const body = items.map((it, idx) => cols.map(c => {
    if (c.key === "sno") return it.sno ?? (idx + 1);
    return it[c.key] ?? "";
  }));

  doc.autoTable({
    head,
    body,
    startY: 75,
    styles: { fontSize: 8, cellPadding: 4 },
    headStyles: { fillColor: [20, 184, 166] },
    theme: "grid",
    margin: { left: 40, right: 40 }
  });

  const filename = safeFileName(`${p.name}_${p.refNo}_export.pdf`);
  doc.save(filename);
  toast("Export", "PDF downloaded.");
}

/* -------------------- Backup / Restore -------------------- */

function backupJSON(){
  const blob = new Blob([JSON.stringify(state, null, 2)], { type:"application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = safeFileName(`procurement_dashboard_backup_${new Date().toISOString().slice(0,10)}.json`);
  a.click();
  URL.revokeObjectURL(url);
  toast("Backup", "Backup downloaded.");
}

async function restoreJSON(e){
  const file = e.target.files?.[0];
  e.target.value = "";
  if (!file) return;

  try{
    const text = await file.text();
    const restored = JSON.parse(text);

    if (!restored || !Array.isArray(restored.projects)){
      return toast("Restore failed", "Invalid backup format.");
    }

    // basic normalization
    restored.projects.forEach(p => {
      if (!p.id) p.id = uid();
      if (!p.sections) p.sections = [{ name:"All", color:"#94a3b8", locked:true }];
      if (!p.items) p.items = [];
      if (!p.columns) p.columns = structuredClone(DEFAULT_COLUMNS);
    });

    state = restored;
    activeProjectId = state.activeProjectId || state.projects[0]?.id || null;
    state.activeProjectId = activeProjectId;
    projectSortMode = state.projectSortMode || "recent";

    saveState();
    render();
    toast("Restored", "Backup restored successfully.");
  } catch(err){
    console.error(err);
    toast("Restore failed", "Could not read the backup file.");
  }
}

/* -------------------- Utilities -------------------- */

function getActiveProject(){
  return state.projects.find(p => p.id === activeProjectId) || null;
}

function loadState(){
  try{
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return { projects:[], activeProjectId:null, projectSortMode:"recent" };
    const parsed = JSON.parse(raw);
    if (!parsed.projects) parsed.projects = [];
    if (!parsed.projectSortMode) parsed.projectSortMode = "recent";
    return parsed;
  }catch{
    return { projects:[], activeProjectId:null, projectSortMode:"recent" };
  }
}

function saveState(){
  localStorage.setItem(LS_KEY, JSON.stringify(state));
}

function openModal(title){
  $("modalTitle").textContent = title;
  $("modalBackdrop").classList.remove("hidden");
}

function closeModal(){
  $("modalBackdrop").classList.add("hidden");
  $("modalBody").innerHTML = "";
  $("modalFoot").innerHTML = "";
}

function toast(title, text){
  const wrap = $("toastWrap");
  const el = document.createElement("div");
  el.className = "toast";
  el.innerHTML = `<div class="toastTitle">${esc(title)}</div><div class="toastText">${esc(text)}</div>`;
  wrap.appendChild(el);
  setTimeout(()=>{ el.style.opacity="0"; el.style.transform="translateY(6px)"; }, 2600);
  setTimeout(()=>{ el.remove(); }, 3200);
}

function uid(){
  return "id_" + Math.random().toString(16).slice(2) + "_" + Date.now().toString(16);
}

function nowISO(){
  return new Date().toISOString();
}

function colorFor(name){
  const n = (name || "").toLowerCase();
  let hash = 0;
  for (let i=0; i<n.length; i++) hash = (hash*31 + n.charCodeAt(i)) >>> 0;
  return COLOR_PALETTE[hash % COLOR_PALETTE.length];
}

function safeFileName(s){
  return (s || "export")
    .toString()
    .replace(/[^\w\-]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function esc(str){
  return String(str ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

function escAttr(str){
  return esc(str).replaceAll("\n"," ");
}

function toISODate(v){
  if (!v) return "";
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  if (!isNaN(d.getTime())){
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth()+1).padStart(2,"0");
    const dd = String(d.getDate()).padStart(2,"0");
    return `${yyyy}-${mm}-${dd}`;
  }
  return "";
}

function sortCompare(a,b,mode){
  const ad = (a.desc||"").toLowerCase();
  const bd = (b.desc||"").toLowerCase();
  const ap = (a.planned||"");
  const bp = (b.planned||"");
  const apr = (a.prDate||"");
  const bpr = (b.prDate||"");

  switch(mode){
    case "desc_desc": return bd.localeCompare(ad);
    case "planned_asc": return ap.localeCompare(bp);
    case "planned_desc": return bp.localeCompare(ap);
    case "prdate_asc": return apr.localeCompare(bpr);
    case "prdate_desc": return bpr.localeCompare(apr);
    case "desc_asc":
    default: return ad.localeCompare(bd);
  }
}

function statusBadge(status){
  const s = (status || "").toUpperCase();
  if (s === "APPROVED") return `<span class="badge ok">APPROVED</span>`;
  if (s === "REJECTED") return `<span class="badge bad">REJECTED</span>`;
  if (s === "RESUBMIT") return `<span class="badge warn">RESUBMIT</span>`;
  return `<span class="badge">PENDING</span>`;
}

function plannedBadgeInfo(planned, approval){
  if (!planned) return `<span class="badge">—</span>`;
  const info = plannedInfo(planned, approval);
  return `<span class="badge ${info.cls}">${esc(planned)} ${info.suffix}</span>`;
}

function plannedInfo(planned, approval){
  const s = (approval||"").toUpperCase();
  if (s === "APPROVED") return { cls:"ok", suffix:"" };

  const today = new Date(); today.setHours(0,0,0,0);
  const d = new Date(planned + "T00:00:00");
  const diffDays = Math.round((d - today) / (1000*60*60*24));

  if (diffDays < 0) return { cls:"bad", suffix:`• overdue ${Math.abs(diffDays)}d` };
  if (diffDays <= 3) return { cls:"warn", suffix:`• due ${diffDays}d` };
  return { cls:"", suffix:`• ${diffDays}d` };
}

function isOverdue(planned, approval){
  if (!planned) return false;
  if ((approval||"").toUpperCase() === "APPROVED") return false;
  const today = new Date(); today.setHours(0,0,0,0);
  const d = new Date(planned + "T00:00:00");
  return d < today;
}

function getSelectedMulti(selectEl){
  return Array.from(selectEl.selectedOptions || []).map(o => o.value);
}

function setSectionMultiSelection(values){
  const sel = $("sectionMulti");
  const set = new Set(values);
  Array.from(sel.options).forEach(o => o.selected = set.has(o.value));
}
