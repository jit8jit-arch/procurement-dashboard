/* Procurement Dashboard (Generalized to Excel Template)
   Columns modeled from your sample:
   - S. No
   - Material Submittal Ref No
   - Rev
   - Description of Material
   - Manufacturer/Supplier
   - Approval Status (dropdown)
   - Planned Date of Submission
   - Est Qty, Ordered Qty
   - PR Number, Date Raised, PR Status (dropdown)
   - LPO Issue Date, LPO Number
   - Payment Status (dropdown)
   - Lead Time
   - Expected Date on Site
   - Qty Delivered, Qty Pending Delivery
   - Status / Remarks

   Includes:
   - Project CRUD (editable anytime)
   - Sections (auto colors)
   - Filters (approval/pr/payment/sections/date range/overdue/search)
   - Pagination
   - Import Excel (header detection)
   - Export Excel/CSV/PDF (filtered view)
   - Backup/Restore JSON
   - Fixed modal (no stuck popup) + mobile UX fixes
*/

const LS_KEY = "proc_dash_general_v1";

const COLOR_PALETTE = [
  "#22c55e","#3b82f6","#f59e0b","#ef4444","#a855f7","#14b8a6",
  "#e11d48","#84cc16","#06b6d4","#f97316","#8b5cf6","#10b981"
];

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

const APPROVAL_OPTIONS = ["APPROVED","PENDING","REJECTED","RESUBMIT"];

const PR_STATUS_OPTIONS = [
  "Pending with Store Keeper",
  "Pending with the QS",
  "Pending",
  "Approved",
  "Rejected",
  "On Hold",
  "Resubmitted",
  "Cancelled"
];

const PAYMENT_STATUS_OPTIONS = [
  "Not Started",
  "Pending",
  "Partially Paid",
  "Paid",
  "On Hold",
  "Cancelled"
];

const DEFAULT_COLUMNS = [
  { key:"sno", label:"#", show:true },
  { key:"subRef", label:"Material Submittal Ref No.", show:true },
  { key:"rev", label:"Rev", show:true },
  { key:"desc", label:"Description of Material", show:true, wrap:true },
  { key:"supplier", label:"Manufacturer / Supplier", show:true },
  { key:"approval", label:"Approval Status", show:true },
  { key:"planned", label:"Planned Date of Submission", show:true },

  { key:"estQty", label:"Est Qty", show:true },
  { key:"ordQty", label:"Ordered Qty", show:true },

  { key:"prNo", label:"PR Number", show:true },
  { key:"prDate", label:"PR Date (Date Raised)", show:true },
  { key:"prStatus", label:"PR Status", show:true },

  { key:"lpoDate", label:"LPO Issue Date", show:true },
  { key:"lpoNo", label:"LPO Number", show:true },

  { key:"payment", label:"Payment Status", show:true },

  { key:"leadTime", label:"Lead Time", show:true },
  { key:"expectedOnSite", label:"Expected Date on Site", show:true },

  { key:"qtyDelivered", label:"Qty Delivered", show:true },
  { key:"qtyPendingDelivery", label:"Qty Pending Delivery", show:true },

  { key:"remarks", label:"Status / Remarks", show:true, wrap:true },

  { key:"section", label:"Section", show:true },
  { key:"actions", label:"", show:true }
];

const $ = (id) => document.getElementById(id);

let state = loadState();
let activeProjectId = state.activeProjectId || null;
let activeSectionChip = "All";
let projectSortMode = state.projectSortMode || "recent";

let currentPage = 1;
let pageSize = 25;

let lastFocusedBeforeModal = null;

init();

function init(){
  populateSuggestions();
  bindUI();
  closeModal(true);
  render();
}

function populateSuggestions(){
  const descDL = $("descSuggestions");
  const secDL = $("sectionSuggestions");
  if (descDL) descDL.innerHTML = GENERAL_DESC_TERMS.map(t => `<option value="${escAttr(t)}"></option>`).join("");
  if (secDL) secDL.innerHTML = GENERAL_SECTION_TERMS.map(t => `<option value="${escAttr(t)}"></option>`).join("");
}

function bindUI(){
  $("btnAddProject").onclick = () => openProjectModal();
  $("btnAddProject2").onclick = () => openProjectModal();

  const btnToggleSidebar = $("btnToggleSidebar");
  if (btnToggleSidebar){
    btnToggleSidebar.onclick = () => $("sidebar").classList.toggle("open");
  }

  $("projectSearch").oninput = renderProjectList;

  $("btnSortProjects").onclick = () => {
    projectSortMode = (projectSortMode === "recent") ? "name" : "recent";
    state.projectSortMode = projectSortMode;
    saveState();
    renderProjectList();
  };

  $("btnClearAll").onclick = confirmClearAll;

  $("btnEditProject").onclick = () => {
    const p = getActiveProject();
    if (p) openProjectModal(p);
  };

  $("btnDeleteProject").onclick = deleteActiveProject;

  $("btnAddSection").onclick = openSectionModal;
  $("btnAddItem").onclick = () => openItemModal();

  // Filters
  $("itemSearch").oninput = () => { currentPage = 1; renderItems(); };
  $("approvalFilter").onchange = () => { currentPage = 1; renderItems(); };
  $("prStatusFilter").onchange = () => { currentPage = 1; renderItems(); };
  $("paymentFilter").onchange = () => { currentPage = 1; renderItems(); };
  $("sectionMulti").onchange = () => { currentPage = 1; renderItems(); };
  $("plannedFrom").onchange = () => { currentPage = 1; renderItems(); };
  $("plannedTo").onchange = () => { currentPage = 1; renderItems(); };
  $("onlyOverdue").onchange = () => { currentPage = 1; renderItems(); };
  $("sortBy").onchange = () => { currentPage = 1; renderItems(); };

  $("btnResetFilters").onclick = resetFilters;
  $("btnColumns").onclick = openColumnsModal;

  // Import/export
  $("excelImport").addEventListener("change", handleExcelImport);
  $("btnExportExcel").onclick = exportExcel;
  $("btnExportCSV").onclick = exportCSV;
  $("btnExportPDF").onclick = exportPDF;

  // Backup/Restore
  $("btnBackup").onclick = backupJSON;
  $("backupImport").addEventListener("change", restoreJSON);

  // Pagination
  $("btnPrevPage").onclick = () => { if (currentPage > 1){ currentPage--; renderItems(); } };
  $("btnNextPage").onclick = () => { currentPage++; renderItems(); };
  $("pageSize").onchange = () => {
    pageSize = parseInt($("pageSize").value, 10) || 25;
    currentPage = 1;
    renderItems();
  };

  // Modal close
  $("modalClose").onclick = () => closeModal();
  $("modalBackdrop").onclick = (e) => { if (e.target === $("modalBackdrop")) closeModal(); };

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && isModalOpen()) closeModal();
    if (e.key === "Tab" && isModalOpen()) trapFocus(e);
  });
}

function render(){
  renderProjectList();
  renderProjectView();
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
    div.tabIndex = 0;
    div.onclick = () => selectProject(p.id);
    div.onkeydown = (e) => { if (e.key === "Enter" || e.key === " ") selectProject(p.id); };

    div.innerHTML = `
      <div class="cardTitle">${esc(p.name || "Untitled")}</div>
      <div class="cardMeta">
        <span class="tag">Ref: ${esc(p.refNo || "—")}</span>
        ${p.client ? `<span class="tag">Client: ${esc(p.client)}</span>` : ""}
      </div>
    `;
    list.appendChild(div);
  }
}

function selectProject(id){
  activeProjectId = id;
  state.activeProjectId = id;
  activeSectionChip = "All";
  currentPage = 1;
  saveState();

  // close sidebar on mobile after selection
  const sb = $("sidebar");
  if (sb) sb.classList.remove("open");

  render();
}

function renderProjectView(){
  const p = getActiveProject();
  const hasActive = !!p;

  $("emptyState").classList.toggle("hidden", hasActive);
  $("projectView").classList.toggle("hidden", !hasActive);
  if (!hasActive) return;

  ensureProjectDefaults(p);

  $("pName").textContent = p.name || "—";
  $("pRef").textContent = `Ref: ${p.refNo || "—"}`;
  $("pClient").textContent = `Client: ${p.client || "—"}`;
  $("pConsultant").textContent = `Consultant: ${p.consultant || "—"}`;
  $("pContractor").textContent = `Main Contractor: ${p.contractor || "—"}`;
  $("pLocation").textContent = `Location: ${p.location || "—"}`;
  $("lastUpdated").textContent = p.updatedAt ? `Last updated: ${new Date(p.updatedAt).toLocaleString()}` : "";

  renderSections();
  renderSectionMultiFilter();
  renderTableHeader();
  renderItems();
}

function ensureProjectDefaults(p){
  if (!p.sections) p.sections = [{ name:"All", color:"#94a3b8", locked:true }];
  if (!p.items) p.items = [];
  if (!p.columns) p.columns = structuredClone(DEFAULT_COLUMNS);
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
      currentPage = 1;
      setSectionMultiSelection(s.name === "All" ? [] : [s.name]);
      renderItems();
    };

    chip.innerHTML = `
      <span class="dot" style="background:${s.color}"></span>
      <span>${esc(s.name)}</span>
      ${(!s.locked && s.name !== "All") ? `<button class="iconBtn" title="Remove section" style="width:28px;height:28px" data-sec="${escAttr(s.name)}">🗑️</button>` : ""}
    `;
    wrap.appendChild(chip);

    const btn = chip.querySelector("button[data-sec]");
    if (btn){
      btn.onclick = (e) => {
        e.stopPropagation();
        removeSection(btn.getAttribute("data-sec"));
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
  setSectionMultiSelection(existing.filter(x => sections.includes(x)));
}

function renderTableHeader(){
  const p = getActiveProject();
  const headRow = $("tableHeadRow");
  headRow.innerHTML = "";
  for (const col of p.columns){
    if (!col.show) continue;
    const th = document.createElement("th");
    th.textContent = col.label;
    headRow.appendChild(th);
  }
}

function getFilteredItemsForView(p){
  const q = ($("itemSearch").value || "").trim().toLowerCase();
  const approval = ($("approvalFilter").value || "").toUpperCase();
  const prStatus = ($("prStatusFilter").value || "");
  const payment = ($("paymentFilter").value || "");
  const sortMode = $("sortBy").value || "desc_asc";
  const sectionsSelected = getSelectedMulti($("sectionMulti"));
  const from = $("plannedFrom").value || "";
  const to = $("plannedTo").value || "";
  const onlyOverdue = $("onlyOverdue").checked;

  let items = (p.items || []).slice();

  if (activeSectionChip && activeSectionChip !== "All"){
    if (sectionsSelected.length === 0){
      items = items.filter(it => (it.section || "All") === activeSectionChip);
    }
  }

  if (sectionsSelected.length){
    items = items.filter(it => sectionsSelected.includes((it.section || "All")));
  }

  if (q){
    items = items.filter(it => {
      const hay = [
        it.desc, it.subRef, it.supplier, it.approval,
        it.prNo, it.prStatus, it.lpoNo, it.payment,
        it.remarks
      ].join(" ").toLowerCase();
      return hay.includes(q);
    });
  }

  if (approval){
    items = items.filter(it => (it.approval || "").toUpperCase() === approval);
  }

  if (prStatus){
    items = items.filter(it => (it.prStatus || "") === prStatus);
  }

  if (payment){
    items = items.filter(it => (it.payment || "") === payment);
  }

  if (from){
    items = items.filter(it => (it.planned || "") >= from);
  }
  if (to){
    items = items.filter(it => (it.planned || "") <= to);
  }

  if (onlyOverdue){
    items = items.filter(it => isOverdue(it.planned, it.approval));
  }

  items.sort((a,b)=> sortCompare(a,b,sortMode));
  return items;
}

function renderItems(){
  const p = getActiveProject();
  if (!p) return;

  const all = getFilteredItemsForView(p);
  updateStats(p);

  const total = all.length;
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  if (currentPage > totalPages) currentPage = totalPages;

  const start = (currentPage - 1) * pageSize;
  const pageItems = all.slice(start, start + pageSize);

  $("resultCount").textContent = `${total} item(s)`;
  $("pageInfo").textContent = `Page ${currentPage} of ${totalPages}`;

  $("btnPrevPage").disabled = currentPage <= 1;
  $("btnNextPage").disabled = currentPage >= totalPages;

  const sections = normalizeSections(p);
  const secColor = new Map(sections.map(s => [s.name, s.color]));

  const body = $("itemsBody");
  body.innerHTML = "";

  const cols = p.columns.filter(c => c.show);

  pageItems.forEach((it, idx) => {
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
        td.innerHTML = `<span class="badge"><span class="dot" style="background:${color}"></span>${esc(sname)}</span>`;
      } else if (col.wrap){
        td.className = "wrap";
        td.textContent = (it[col.key] ?? "");
      } else if (col.key === "sno"){
        td.textContent = it.sno || (start + idx + 1);
      } else {
        td.textContent = (it[col.key] ?? "");
      }

      tr.appendChild(td);
    }

    body.appendChild(tr);
  });

  body.querySelectorAll("button[data-edit]").forEach(btn => {
    btn.onclick = () => {
      const id = btn.getAttribute("data-edit");
      const item = p.items.find(x => x.id === id);
      if (item) openItemModal(item);
    };
  });

  body.querySelectorAll("button[data-del]").forEach(btn => {
    btn.onclick = () => confirmDeleteItem(btn.getAttribute("data-del"));
  });
}

function updateStats(p){
  const items = p.items || [];
  const approved = items.filter(i => (i.approval||"").toUpperCase() === "APPROVED").length;
  const pending = items.filter(i => (i.approval||"").toUpperCase() === "PENDING").length;
  const overdue = items.filter(i => isOverdue(i.planned, i.approval)).length;

  $("stTotal").textContent = items.length;
  $("stApproved").textContent = approved;
  $("stPending").textContent = pending;
  $("stOverdue").textContent = overdue;
}

/* ---------- Project CRUD ---------- */

function openProjectModal(existing=null){
  openModal(existing ? "Edit Project" : "New Project");

  const p = existing || { name:"", refNo:"", client:"", consultant:"", contractor:"", location:"", notes:"" };

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

  $("m_cancel").onclick = () => closeModal();
  $("m_save").onclick = () => {
    const name = $("m_pName").value.trim();
    const refNo = $("m_pRef").value.trim();
    if (!name) return toast("Missing", "Project Name is required.");
    if (!refNo) return toast("Missing", "Project Reference No. is required.");

    setSaving(true);

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
      currentPage = 1;
    }

    saveState();
    closeModal();
    setSaving(false);
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
      Delete project <b>${esc(p.name)}</b>? This will remove all its items.
      <div class="muted" style="margin-top:8px">This only affects your browser data.</div>
    </div>
  `;
  $("modalFoot").innerHTML = `
    <button class="btn" id="m_cancel">Cancel</button>
    <button class="btn danger" id="m_del">Delete</button>
  `;
  $("m_cancel").onclick = () => closeModal();
  $("m_del").onclick = () => {
    setSaving(true);
    state.projects = state.projects.filter(x => x.id !== p.id);
    activeProjectId = state.projects[0]?.id || null;
    state.activeProjectId = activeProjectId;
    activeSectionChip = "All";
    currentPage = 1;
    saveState();
    closeModal();
    setSaving(false);
    render();
  };
}

/* ---------- Section CRUD ---------- */

function openSectionModal(){
  const p = getActiveProject();
  if (!p) return;

  openModal("Add Section / Category");
  $("modalBody").innerHTML = `
    <div class="grid">
      <div class="field full">
        <label>Section Name</label>
        <input id="m_sName" list="sectionSuggestions" placeholder="Type or pick from suggestions…" />
      </div>
    </div>
    <div class="muted" style="margin-top:10px;font-size:12px">Color will be assigned automatically.</div>
  `;
  $("modalFoot").innerHTML = `
    <button class="btn" id="m_cancel">Cancel</button>
    <button class="btn primary" id="m_add">Add Section</button>
  `;
  $("m_cancel").onclick = () => closeModal();
  $("m_add").onclick = () => {
    const name = $("m_sName").value.trim();
    if (!name) return toast("Missing", "Section name is required.");

    const sections = normalizeSections(p);
    if (sections.some(s => s.name.toLowerCase() === name.toLowerCase())){
      return toast("Duplicate", "This section already exists.");
    }

    setSaving(true);
    p.sections = sections.filter(s => s.name !== "All");
    p.sections.push({ name, color: colorFor(name), locked:false });
    p.updatedAt = nowISO();
    saveState();
    setSaving(false);

    closeModal();
    activeSectionChip = name;
    setSectionMultiSelection([name]);
    currentPage = 1;
    render();
  };
}

function removeSection(sectionName){
  const p = getActiveProject();
  if (!p) return;

  const sections = normalizeSections(p);
  const s = sections.find(x => x.name === sectionName);
  if (!s || s.locked) return;

  setSaving(true);
  (p.items || []).forEach(it => { if ((it.section||"") === sectionName) it.section = "All"; });
  p.sections = sections.filter(x => x.name !== "All" && x.name !== sectionName);
  p.updatedAt = nowISO();
  saveState();
  setSaving(false);

  if (activeSectionChip === sectionName) activeSectionChip = "All";
  setSectionMultiSelection([]);
  currentPage = 1;
  render();
}

function normalizeSections(p){
  const raw = (p.sections || []).slice();
  const hasAll = raw.some(s => s.name === "All");
  const sections = hasAll ? raw : [{ name:"All", color:"#94a3b8", locked:true }, ...raw];

  sections.forEach(s => { if (!s.color) s.color = s.name === "All" ? "#94a3b8" : colorFor(s.name); });

  const all = sections.find(s => s.name === "All");
  if (all) all.locked = true;

  const seen = new Set();
  return sections.filter(s => {
    const k = (s.name||"").toLowerCase();
    if (seen.has(k)) return false;
    seen.add(k);
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

/* ---------- Item CRUD (Excel template) ---------- */

function openItemModal(existing=null){
  const p = getActiveProject();
  if (!p) return;

  const isEdit = !!existing;
  const it = existing || {
    id: uid(),
    sno:"",
    subRef:"",
    rev:"",
    desc:"",
    supplier:"",
    approval:"PENDING",
    planned:"",
    estQty:"",
    ordQty:"",
    prNo:"",
    prDate:"",
    prStatus:"Pending",
    lpoDate:"",
    lpoNo:"",
    payment:"Pending",
    leadTime:"",
    expectedOnSite:"",
    qtyDelivered:"",
    qtyPendingDelivery:"",
    remarks:"",
    section: pickDefaultSection()
  };

  openModal(isEdit ? "Edit Item" : "Add Item");

  $("modalBody").innerHTML = `
    <div class="grid">
      <div class="field">
        <label>S. No</label>
        <input id="m_sno" value="${escAttr(it.sno)}" />
      </div>

      <div class="field">
        <label>Material Submittal Ref No.</label>
        <input id="m_subref" value="${escAttr(it.subRef)}" />
      </div>

      <div class="field">
        <label>Rev</label>
        <input id="m_rev" value="${escAttr(it.rev)}" />
      </div>

      <div class="field">
        <label>Approval Status</label>
        <select id="m_approval">
          ${APPROVAL_OPTIONS.map(x => `<option ${String(it.approval||"").toUpperCase()===x?"selected":""}>${x}</option>`).join("")}
        </select>
      </div>

      <div class="field full">
        <label>Description of Material</label>
        <input id="m_desc" list="descSuggestions" value="${escAttr(it.desc)}" placeholder="Type or pick a term…" />
      </div>

      <div class="field full">
        <label>Manufacturer / Supplier</label>
        <input id="m_supplier" value="${escAttr(it.supplier)}" />
      </div>

      <div class="field">
        <label>Planned Date of Submission</label>
        <input id="m_planned" type="date" value="${escAttr(it.planned)}" />
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
        <input id="m_est" value="${escAttr(it.estQty)}" />
      </div>

      <div class="field">
        <label>Ordered Qty</label>
        <input id="m_ord" value="${escAttr(it.ordQty)}" />
      </div>

      <div class="field">
        <label>PR Number</label>
        <input id="m_prno" value="${escAttr(it.prNo)}" />
      </div>

      <div class="field">
        <label>PR Date (Date Raised)</label>
        <input id="m_prdate" type="date" value="${escAttr(it.prDate)}" />
      </div>

      <div class="field">
        <label>PR Status</label>
        <select id="m_prstatus">
          ${PR_STATUS_OPTIONS.map(s => `<option ${it.prStatus===s?"selected":""}>${esc(s)}</option>`).join("")}
        </select>
      </div>

      <div class="field">
        <label>LPO Issue Date</label>
        <input id="m_lpodate" type="date" value="${escAttr(it.lpoDate)}" />
      </div>

      <div class="field">
        <label>LPO Number</label>
        <input id="m_lpono" value="${escAttr(it.lpoNo)}" />
      </div>

      <div class="field">
        <label>Payment Status</label>
        <select id="m_payment">
          ${PAYMENT_STATUS_OPTIONS.map(s => `<option ${it.payment===s?"selected":""}>${esc(s)}</option>`).join("")}
        </select>
      </div>

      <div class="field">
        <label>Lead Time</label>
        <input id="m_lead" value="${escAttr(it.leadTime)}" placeholder="e.g., 2 weeks" />
      </div>

      <div class="field">
        <label>Expected Date on Site</label>
        <input id="m_expected" type="date" value="${escAttr(it.expectedOnSite)}" />
      </div>

      <div class="field">
        <label>Qty Delivered</label>
        <input id="m_delivered" value="${escAttr(it.qtyDelivered)}" />
      </div>

      <div class="field">
        <label>Qty Pending Delivery</label>
        <input id="m_pendingdel" value="${escAttr(it.qtyPendingDelivery)}" />
      </div>

      <div class="field full">
        <label>Status / Remarks</label>
        <input id="m_remarks" value="${escAttr(it.remarks)}" />
      </div>
    </div>
  `;

  $("modalFoot").innerHTML = `
    <button class="btn" id="m_cancel">Cancel</button>
    <button class="btn primary" id="m_save">${isEdit ? "Save Changes" : "Add Item"}</button>
  `;

  $("m_cancel").onclick = () => closeModal();
  $("m_save").onclick = () => {
    const desc = $("m_desc").value.trim();
    if (!desc) return toast("Missing", "Description of Material is required.");

    setSaving(true);

    it.sno = $("m_sno").value.trim();
    it.subRef = $("m_subref").value.trim();
    it.rev = $("m_rev").value.trim();
    it.desc = desc;
    it.supplier = $("m_supplier").value.trim();
    it.approval = $("m_approval").value.trim();
    it.planned = $("m_planned").value;

    it.estQty = $("m_est").value.trim();
    it.ordQty = $("m_ord").value.trim();

    it.prNo = $("m_prno").value.trim();
    it.prDate = $("m_prdate").value;
    it.prStatus = $("m_prstatus").value;

    it.lpoDate = $("m_lpodate").value;
    it.lpoNo = $("m_lpono").value.trim();

    it.payment = $("m_payment").value;

    it.leadTime = $("m_lead").value.trim();
    it.expectedOnSite = $("m_expected").value;
    it.qtyDelivered = $("m_delivered").value.trim();
    it.qtyPendingDelivery = $("m_pendingdel").value.trim();

    it.remarks = $("m_remarks").value.trim();
    it.section = $("m_section").value;

    ensureSection(p, it.section);

    if (isEdit){
      const idx = p.items.findIndex(x => x.id === it.id);
      if (idx >= 0) p.items[idx] = it;
    } else {
      p.items.push(it);
    }

    p.updatedAt = nowISO();
    saveState();

    setSaving(false);
    closeModal();
    renderSectionMultiFilter();
    renderSections();
    renderTableHeader();
    renderItems();
    toast("Saved", isEdit ? "Item updated." : "Item added.");
  };
}

function confirmDeleteItem(itemId){
  const p = getActiveProject();
  if (!p) return;
  const it = p.items.find(x => x.id === itemId);
  if (!it) return;

  openModal("Delete Item?");
  $("modalBody").innerHTML = `<div style="line-height:1.5">Delete item <b>${esc(it.desc)}</b>?</div>`;
  $("modalFoot").innerHTML = `
    <button class="btn" id="m_cancel">Cancel</button>
    <button class="btn danger" id="m_del">Delete</button>
  `;
  $("m_cancel").onclick = () => closeModal();
  $("m_del").onclick = () => {
    setSaving(true);
    p.items = p.items.filter(x => x.id !== itemId);
    p.updatedAt = nowISO();
    saveState();
    setSaving(false);
    closeModal();
    renderItems();
  };
}

function pickDefaultSection(){
  const sel = getSelectedMulti($("sectionMulti"));
  if (sel.length === 1) return sel[0];
  if (activeSectionChip && activeSectionChip !== "All") return activeSectionChip;
  return "All";
}

/* ---------- Columns ---------- */
function openColumnsModal(){
  const p = getActiveProject();
  if (!p) return;

  openModal("Choose Columns");
  const rows = p.columns.filter(c => c.key !== "actions").map(c => `
    <label style="display:flex;align-items:center;gap:10px;padding:6px 0">
      <input type="checkbox" data-col="${escAttr(c.key)}" ${c.show ? "checked" : ""} />
      <span>${esc(c.label)}</span>
    </label>
  `).join("");

  $("modalBody").innerHTML = `<div class="muted" style="margin-bottom:10px">Columns are saved per project.</div>${rows}`;
  $("modalFoot").innerHTML = `
    <button class="btn" id="m_cancel">Cancel</button>
    <button class="btn primary" id="m_save">Save</button>
  `;

  $("m_cancel").onclick = () => closeModal();
  $("m_save").onclick = () => {
    const checks = $("modalBody").querySelectorAll("input[type=checkbox][data-col]");
    const wanted = new Map();
    checks.forEach(ch => wanted.set(ch.getAttribute("data-col"), ch.checked));
    setSaving(true);
    p.columns = p.columns.map(c => wanted.has(c.key) ? { ...c, show: wanted.get(c.key) } : c);
    p.updatedAt = nowISO();
    saveState();
    setSaving(false);
    closeModal();
    renderTableHeader();
    renderItems();
  };
}

/* ---------- Filters Reset ---------- */
function resetFilters(){
  $("itemSearch").value = "";
  $("approvalFilter").value = "";
  $("prStatusFilter").value = "";
  $("paymentFilter").value = "";
  $("plannedFrom").value = "";
  $("plannedTo").value = "";
  $("onlyOverdue").checked = false;
  setSectionMultiSelection([]);
  activeSectionChip = "All";
  currentPage = 1;
  renderSections();
  renderItems();
}

/* ---------- Import Excel (header detection + safe mapping) ---------- */
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

    const headerRow = rows.findIndex(r =>
      (r || []).some(cell => String(cell||"").toUpperCase().includes("DESCRIPTION OF MATERIAL"))
    );
    if (headerRow < 0){
      toast("Import failed", "Could not find 'Description of Material' header row.");
      return;
    }

    const p = getActiveProject();
    if (!p) return;

    // Best guess start row: after header block (often +3)
    let startRow = headerRow + 1;
    for (let i = headerRow + 1; i < Math.min(rows.length, headerRow + 6); i++){
      const r = rows[i] || [];
      const joined = r.map(x=>String(x||"").toUpperCase()).join(" ");
      if (joined.includes("EST") && joined.includes("ORDERED")) startRow = i + 1;
    }

    const imported = [];
    for (let i = startRow; i < rows.length; i++){
      const r = rows[i] || [];
      const desc = String(r[3] || "").trim();
      if (!desc) continue;

      imported.push({
        id: uid(),
        sno: String(r[0] ?? "").trim(),
        subRef: String(r[1] ?? "").trim(),
        rev: String(r[2] ?? "").trim(),
        desc: desc,
        supplier: String(r[4] ?? "").trim(),
        approval: String(r[5] ?? "PENDING").toUpperCase().trim(),
        planned: toISODate(r[6]),
        estQty: String(r[7] ?? "").trim(),
        ordQty: String(r[8] ?? "").trim(),
        prNo: String(r[9] ?? "").trim(),
        prDate: toISODate(r[10]),
        prStatus: String(r[11] ?? "").trim(),
        lpoDate: toISODate(r[12]),
        lpoNo: String(r[13] ?? "").trim(),
        payment: String(r[14] ?? "").trim(),
        leadTime: String(r[15] ?? "").trim(),
        expectedOnSite: toISODate(r[16]),
        qtyDelivered: String(r[17] ?? "").trim(),
        qtyPendingDelivery: String(r[18] ?? "").trim(),
        remarks: String(r[19] ?? "").trim(),
        section: pickDefaultSection()
      });
    }

    if (!imported.length){
      toast("Import", "No items found to import.");
      return;
    }

    setSaving(true);
    imported.forEach(it => ensureSection(p, it.section));
    p.items = (p.items || []).concat(imported);
    p.updatedAt = nowISO();
    saveState();
    setSaving(false);

    currentPage = 1;
    renderItems();
    toast("Imported", `${imported.length} item(s) imported.`);
  } catch(err){
    console.error(err);
    toast("Import failed", "Please try again with the Excel file.");
  }
}

/* ---------- Export (filtered view) ---------- */
function exportExcel(){
  const p = getActiveProject();
  if (!p) return;

  const items = getFilteredItemsForView(p);
  const cols = p.columns.filter(c => c.show && c.key !== "actions");

  const aoa = [cols.map(c => c.label)];
  items.forEach((it, idx) => {
    aoa.push(cols.map(c => c.key === "sno" ? (it.sno || (idx+1)) : (it[c.key] ?? "")));
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Export");

  XLSX.writeFile(wb, safeFileName(`${p.name}_${p.refNo}_export.xlsx`));
}

function exportCSV(){
  const p = getActiveProject();
  if (!p) return;

  const items = getFilteredItemsForView(p);
  const cols = p.columns.filter(c => c.show && c.key !== "actions");

  const header = cols.map(c => `"${String(c.label).replaceAll('"','""')}"`).join(",");
  const rows = items.map((it, idx) => cols.map(c => {
    const v = (c.key === "sno") ? (it.sno || (idx+1)) : (it[c.key] ?? "");
    return `"${String(v).replaceAll('"','""')}"`;
  }).join(","));

  const csv = [header, ...rows].join("\n");
  const blob = new Blob([csv], { type:"text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = safeFileName(`${p.name}_${p.refNo}_export.csv`);
  a.click();
  URL.revokeObjectURL(url);
}

function exportPDF(){
  const p = getActiveProject();
  if (!p) return;

  const items = getFilteredItemsForView(p);
  const cols = p.columns.filter(c => c.show && c.key !== "actions");

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation:"landscape", unit:"pt", format:"a4" });

  doc.setFontSize(14);
  doc.text(`${p.name} — Procurement Log`, 40, 40);
  doc.setFontSize(10);
  doc.text(`Ref: ${p.refNo}   Generated: ${new Date().toLocaleString()}`, 40, 58);

  doc.autoTable({
    head: [cols.map(c => c.label)],
    body: items.map((it, idx) => cols.map(c => c.key === "sno" ? (it.sno || (idx+1)) : (it[c.key] ?? ""))),
    startY: 75,
    styles: { fontSize: 8, cellPadding: 4 },
    theme: "grid",
    margin: { left: 40, right: 40 }
  });

  doc.save(safeFileName(`${p.name}_${p.refNo}_export.pdf`));
}

/* ---------- Clear all ---------- */
function confirmClearAll(){
  openModal("Clear All Data?");
  $("modalBody").innerHTML = `<div style="line-height:1.5">This will delete all projects saved in this browser.</div>`;
  $("modalFoot").innerHTML = `
    <button class="btn" id="m_cancel">Cancel</button>
    <button class="btn danger" id="m_clear">Clear</button>
  `;
  $("m_cancel").onclick = () => closeModal();
  $("m_clear").onclick = () => {
    localStorage.removeItem(LS_KEY);
    state = { projects:[], activeProjectId:null, projectSortMode:"recent" };
    activeProjectId = null;
    currentPage = 1;
    closeModal();
    render();
  };
}

/* ---------- Backup / Restore ---------- */
function backupJSON(){
  const blob = new Blob([JSON.stringify(state, null, 2)], { type:"application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = safeFileName(`procurement_backup_${new Date().toISOString().slice(0,10)}.json`);
  a.click();
  URL.revokeObjectURL(url);
}

async function restoreJSON(e){
  const file = e.target.files?.[0];
  e.target.value = "";
  if (!file) return;

  try{
    const text = await file.text();
    const restored = JSON.parse(text);
    if (!restored || !Array.isArray(restored.projects)) return;

    restored.projects.forEach(p => {
      if (!p.id) p.id = uid();
      if (!p.sections) p.sections = [{ name:"All", color:"#94a3b8", locked:true }];
      if (!p.items) p.items = [];
      if (!p.columns) p.columns = structuredClone(DEFAULT_COLUMNS);
    });

    state = restored;
    activeProjectId = state.activeProjectId || state.projects[0]?.id || null;
    state.activeProjectId = activeProjectId;
    saveState();
    currentPage = 1;
    render();
  } catch {
    toast("Restore failed", "Invalid backup file.");
  }
}

/* ---------- Modal core (fixed + mobile drawer close) ---------- */
function isModalOpen(){ return $("modalBackdrop").classList.contains("show"); }

function openModal(title){
  lastFocusedBeforeModal = document.activeElement;

  // ✅ close sidebar drawer on mobile so it doesn’t cover the modal
  const sb = $("sidebar");
  if (sb) sb.classList.remove("open");

  $("modalTitle").textContent = title;
  $("modalBackdrop").classList.add("show");
  document.body.classList.add("modalOpen");

  setTimeout(() => {
    const f = getModalFocusable();
    if (f.length) f[0].focus();
  }, 0);
}

function closeModal(force=false){
  $("modalBackdrop").classList.remove("show");
  document.body.classList.remove("modalOpen");
  $("modalBody").innerHTML = "";
  $("modalFoot").innerHTML = "";
  if (!force && lastFocusedBeforeModal?.focus) lastFocusedBeforeModal.focus();
  lastFocusedBeforeModal = null;
}

function getModalFocusable(){
  const modal = $("modalBackdrop");
  return Array.from(modal.querySelectorAll('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'))
    .filter(el => !el.disabled && el.offsetParent !== null);
}

function trapFocus(e){
  const focusable = getModalFocusable();
  if (!focusable.length) return;
  const first = focusable[0];
  const last = focusable[focusable.length - 1];
  if (e.shiftKey && document.activeElement === first){
    e.preventDefault(); last.focus();
  } else if (!e.shiftKey && document.activeElement === last){
    e.preventDefault(); first.focus();
  }
}

/* ---------- Saving indicator ---------- */
function setSaving(isSaving){
  const el = $("saveState");
  if (!el) return;
  el.textContent = isSaving ? "Saving…" : "Saved";
  el.classList.toggle("saving", isSaving);
}

/* ---------- Helpers ---------- */
function getActiveProject(){ return state.projects.find(p => p.id === activeProjectId) || null; }

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

function saveState(){ localStorage.setItem(LS_KEY, JSON.stringify(state)); }
function uid(){ return "id_" + Math.random().toString(16).slice(2) + "_" + Date.now().toString(16); }
function nowISO(){ return new Date().toISOString(); }

function colorFor(name){
  const n = (name || "").toLowerCase();
  let hash = 0;
  for (let i=0; i<n.length; i++) hash = (hash*31 + n.charCodeAt(i)) >>> 0;
  return COLOR_PALETTE[hash % COLOR_PALETTE.length];
}

function safeFileName(s){
  return (s || "export").toString().replace(/[^\w\-]+/g, "_").replace(/_+/g, "_").replace(/^_+|_+$/g, "");
}

function esc(str){
  return String(str ?? "").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;");
}
function escAttr(str){ return esc(str).replaceAll("\n"," "); }

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

function toast(title, text){
  const wrap = $("toastWrap");
  const el = document.createElement("div");
  el.className = "toast";
  el.innerHTML = `<div class="toastTitle">${esc(title)}</div><div class="toastText">${esc(text)}</div>`;
  wrap.appendChild(el);
  setTimeout(()=>{ el.style.opacity="0"; el.style.transform="translateY(6px)"; }, 2600);
  setTimeout(()=>{ el.remove(); }, 3200);
}
