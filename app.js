/* Procurement Projects Dashboard
   - Projects (Project Name required, Ref required; client/consultant optional)
   - Sections per project (auto color)
   - Items based on procurement columns
   - Import XLSX / Export XLSX / Export PDF
   - In-app notifications + optional browser notifications
   - Suggested general terms for common materials and sections
*/

const LS_KEY = "proc_dash_v1";

const COLOR_PALETTE = [
  "#22c55e","#3b82f6","#f59e0b","#ef4444","#a855f7","#14b8a6",
  "#e11d48","#84cc16","#06b6d4","#f97316","#8b5cf6","#10b981"
];

// Suggested descriptions and sections based on typical procurement items
const GENERAL_DESC_TERMS = [
  "Pipes & Fittings",
  "Pressure Gauge",
  "Valves",
  "Equipment",
  "Pumps",
  "Electrical Panel",
  "Mechanical Equipment",
  "Civil Works",
  "Cables",
  "Fixtures"
];
const GENERAL_SECTION_TERMS = [
  "Procurement",
  "Mechanical",
  "Electrical",
  "Civil",
  "Equipment"
];

const $ = (id) => document.getElementById(id);

// Load persisted state or initialise empty state
let state = loadState();
let activeProjectId = state.activeProjectId || null;
let activeSection = "All";

init();

function init(){
  bindUI();
  populateSuggestions();
  // ensure modal is hidden on load
  closeModal();
  // render after loading suggestions
  render();
}

function populateSuggestions(){
  // Populate datalist options for descriptions and sections
  const descDL = $("descSuggestions");
  const secDL = $("sectionSuggestions");
  if (descDL){
    descDL.innerHTML = GENERAL_DESC_TERMS.map(t => `<option value="${escapeHtmlAttr(t)}"></option>`).join("");
  }
  if (secDL){
    secDL.innerHTML = GENERAL_SECTION_TERMS.map(t => `<option value="${escapeHtmlAttr(t)}"></option>`).join("");
  }
}

function bindUI(){
  $("btnAddProject").onclick = () => openProjectModal();
  $("btnAddProject2").onclick = () => openProjectModal();

  $("projectSearch").oninput = renderProjectList;

  $("btnEditProject").onclick = () => {
    const p = getActiveProject();
    if (p) openProjectModal(p);
  };
  $("btnDeleteProject").onclick = () => deleteActiveProject();

  $("btnAddSection").onclick = () => openSectionModal();
  $("btnAddItem").onclick = () => openItemModal();

  $("itemSearch").oninput = renderItems;
  $("statusFilter").onchange = renderItems;
  $("sortBy").onchange = renderItems;

  $("btnExportExcel").onclick = exportExcel;
  $("btnExportPDF").onclick = exportPDF;

  $("excelImport").addEventListener("change", handleExcelImport);

  $("modalClose").onclick = closeModal;
  $("modalBackdrop").onclick = (e) => { if (e.target === $("modalBackdrop")) closeModal(); };

  $("btnEnableNoti").onclick = async () => {
    if (!("Notification" in window)) {
      toast("Notifications", "Your browser doesn’t support notifications.");
      return;
    }
    const res = await Notification.requestPermission();
    toast("Notifications", res === "granted" ? "Enabled." : "Not enabled.");
    renderNotifications();
  };
}

function render(){
  renderProjectList();
  renderProjectView();
  renderNotifications();
}

/* -------------------- Project list / view -------------------- */

function renderProjectList(){
  const list = $("projectList");
  list.innerHTML = "";
  const q = ($("projectSearch").value || "").trim().toLowerCase();

  const filtered = state.projects
    .slice()
    .sort((a,b)=> (b.createdAt||"").localeCompare(a.createdAt||""))
    .filter(p => {
      if (!q) return true;
      return (p.name||"").toLowerCase().includes(q) || (p.refNo||"").toLowerCase().includes(q);
    });

  for (const p of filtered){
    const div = document.createElement("div");
    div.className = "card" + (p.id === activeProjectId ? " active" : "");
    div.onclick = () => {
      activeProjectId = p.id;
      state.activeProjectId = p.id;
      activeSection = "All";
      saveState();
      render();
    };

    div.innerHTML = `
      <div class="cardTitle">${escapeHtml(p.name || "Untitled")}</div>
      <div class="cardMeta">
        <span class="tag">Ref: ${escapeHtml(p.refNo || "—")}</span>
        ${p.client ? `<span class="tag">Client: ${escapeHtml(p.client)}</span>` : ""}
        ${p.consultant ? `<span class="tag">Consultant: ${escapeHtml(p.consultant)}</span>` : ""}
      </div>
    `;
    list.appendChild(div);
  }
}

function renderProjectView(){
  const hasActive = !!getActiveProject();
  $("emptyState").classList.toggle("hidden", hasActive);
  $("projectView").classList.toggle("hidden", !hasActive);
  if (!hasActive) return;

  const p = getActiveProject();
  $("pName").textContent = p.name || "—";
  $("pRef").textContent = `Ref: ${p.refNo || "—"}`;
  $("pClient").textContent = `Client: ${p.client || "—"}`;
  $("pConsultant").textContent = `Consultant: ${p.consultant || "—"}`;
  $("pContractor").textContent = `Main Contractor: ${p.contractor || "—"}`;

  renderSections();
  renderItems();
}

function getActiveProject(){
  return state.projects.find(p => p.id === activeProjectId) || null;
}

/* -------------------- Sections -------------------- */

function renderSections(){
  const p = getActiveProject();
  const wrap = $("sectionChips");
  wrap.innerHTML = "";
  const sections = normalizeSections(p);
  for (const s of sections){
    const chip = document.createElement("div");
    chip.className = "chip" + (s.name === activeSection ? " active" : "");
    chip.onclick = () => { activeSection = s.name; renderItems(); renderNotifications(); };

    chip.innerHTML = `
      <span class="dot" style="background:${s.color}"></span>
      <span>${escapeHtml(s.name)}</span>
      ${(!s.locked && s.name !== "All") ? `<button class="iconBtn" title="Remove section" style="width:28px;height:28px" data-sec="${escapeHtmlAttr(s.name)}">🗑️</button>` : ""}
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
    <div class="muted" style="margin-top:10px;font-size:12px">
      Color will be assigned automatically.
    </div>
  `;
  $("modalFoot").innerHTML = `
    <button class="btn" id="m_cancel">Cancel</button>
    <button class="btn primary" id="m_add">Add Section</button>
  `;
  $("m_cancel").onclick = closeModal;
  $("m_add").onclick = () => {
    const name = $("m_sName").value.trim();
    if (!name){ toast("Missing", "Section name is required."); return; }
    const sections = normalizeSections(p);
    if (sections.some(s => s.name.toLowerCase() === name.toLowerCase())){
      toast("Duplicate", "This section already exists.");
      return;
    }
    p.sections = sections.filter(s => s.name !== "All");
    p.sections.push({ name, color: colorFor(name), locked:false });
    saveState();
    closeModal();
    activeSection = name;
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
  if (activeSection === sectionName) activeSection = "All";
  saveState();
  render();
  toast("Updated", `Section "${sectionName}" removed (items moved to All).`);
}

/* -------------------- Items -------------------- */

function renderItems(){
  const p = getActiveProject();
  if (!p) return;
  const q = ($("itemSearch").value || "").trim().toLowerCase();
  const status = $("statusFilter").value || "";
  const sortMode = $("sortBy").value || "desc_asc";
  let items = (p.items || []).slice();
  if (activeSection && activeSection !== "All"){
    items = items.filter(it => (it.section || "") === activeSection);
  }
  if (q){
    items = items.filter(it => {
      const hay = [it.desc, it.ref, it.mfg, it.approval, it.prNo, it.lpoNo, it.payment, it.prStatus].join(" ").toLowerCase();
      return hay.includes(q);
    });
  }
  if (status){
    items = items.filter(it => (it.approval || "").toUpperCase() === status);
  }
  items.sort((a,b)=> sortCompare(a,b,sortMode));
  const body = $("itemsBody");
  body.innerHTML = "";
  const sections = normalizeSections(p);
  const secColor = new Map(sections.map(s => [s.name, s.color]));
  items.forEach((it, idx) => {
    const tr = document.createElement("tr");
    const approvalBadge = statusBadge(it.approval);
    const plannedBadge = plannedBadgeInfo(it.planned, it.approval);
    tr.innerHTML = `
      <td>${it.sno ?? (idx+1)}</td>
      <td>${escapeHtml(it.ref || "")}</td>
      <td>${escapeHtml(it.rev || "")}</td>
      <td class="wrap">${escapeHtml(it.desc || "")}</td>
      <td>${escapeHtml(it.mfg || "")}</td>
      <td>${approvalBadge}</td>
      <td>${plannedBadge}</td>
      <td>${escapeHtml(it.estQty ?? "")}</td>
      <td>${escapeHtml(it.ordQty ?? "")}</td>
      <td>${escapeHtml(it.prNo || "")}</td>
      <td>${escapeHtml(it.prDate || "")}</td>
      <td>${escapeHtml(it.prStatus || "")}</td>
      <td>${escapeHtml(it.lpoDate || "")}</td>
      <td>${escapeHtml(it.lpoNo || "")}</td>
      <td>${escapeHtml(it.payment || "")}</td>
      <td>
        <span class="badge" title="Section">
          <span class="dot" style="background:${secColor.get(it.section || 'All') || '#94a3b8'}"></span>
          ${escapeHtml(it.section || "—")}
        </span>
      </td>
      <td>
        <div class="rowActions">
          <button class="iconBtn" title="Edit" data-edit="${it.id}">✏️</button>
          <button class="iconBtn" title="Delete" data-del="${it.id}">🗑️</button>
        </div>
      </td>
    `;
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
    btn.onclick = () => {
      const id = btn.getAttribute("data-del");
      deleteItem(id);
    };
  });
  renderNotifications();
}

function openItemModal(existing=null){
  const p = getActiveProject();
  if (!p) return;
  const isEdit = !!existing;
  const it = existing || {
    id: uid(),
    sno: "",
    ref: "",
    rev: "",
    desc: "",
    mfg: "",
    approval: "PENDING",
    planned: "",
    estQty: "",
    ordQty: "",
    prNo: "",
    prDate: "",
    prStatus: "",
    lpoDate: "",
    lpoNo: "",
    payment: "",
    section: (activeSection && activeSection !== "All") ? activeSection : "All"
  };
  const sections = normalizeSections(p).filter(s => s.name !== "All").map(s => s.name);
  if (!sections.length) sections.push("Procurement");
  openModal(isEdit ? "Edit Item" : "Add Item");
  $("modalBody").innerHTML = `
    <div class="grid">
      <div class="field">
        <label>S. No</label>
        <input id="m_sno" value="${escapeHtmlAttr(it.sno ?? "")}" placeholder="auto/optional" />
      </div>
      <div class="field">
        <label>Material Submittal Ref No.</label>
        <input id="m_ref" value="${escapeHtmlAttr(it.ref ?? "")}" />
      </div>
      <div class="field">
        <label>Rev</label>
        <input id="m_rev" value="${escapeHtmlAttr(it.rev ?? "")}" />
      </div>
      <div class="field">
        <label>Approval Status</label>
        <select id="m_approval">
          ${["APPROVED","PENDING","REJECTED","RESUBMIT"].map(x => `<option ${String(it.approval||"").toUpperCase()===x?"selected":""}>${x}</option>`).join("")}
        </select>
      </div>
      <div class="field full">
        <label>Description of Material</label>
        <input id="m_desc" list="descSuggestions" value="${escapeHtmlAttr(it.desc ?? "")}" placeholder="e.g., Pipes & Fittings" />
      </div>
      <div class="field full">
        <label>Manufacturer / Supplier</label>
        <input id="m_mfg" value="${escapeHtmlAttr(it.mfg ?? "")}" />
      </div>
      <div class="field">
        <label>Planned Date of Submission</label>
        <input id="m_planned" type="date" value="${escapeHtmlAttr(it.planned ?? "")}" />
      </div>
      <div class="field">
        <label>Section / Category</label>
        <select id="m_section">
          <option ${it.section==="All"?"selected":""}>All</option>
          ${normalizeSections(p).filter(s => s.name!=="All").map(s => `<option ${it.section===s.name?"selected":""}>${escapeHtml(s.name)}</option>`).join("")}
        </select>
      </div>
      <div class="field">
        <label>Est Qty</label>
        <input id="m_est" value="${escapeHtmlAttr(it.estQty ?? "")}" />
      </div>
      <div class="field">
        <label>Ordered Qty</label>
        <input id="m_ord" value="${escapeHtmlAttr(it.ordQty ?? "")}" />
      </div>
      <div class="field">
        <label>PR Number</label>
        <input id="m_prno" value="${escapeHtmlAttr(it.prNo ?? "")}" />
      </div>
      <div class="field">
        <label>PR Date</label>
        <input id="m_prdate" type="date" value="${escapeHtmlAttr(it.prDate ?? "")}" />
      </div>
      <div class="field full">
        <label>PR Status</label>
        <input id="m_prstatus" value="${escapeHtmlAttr(it.prStatus ?? "")}" />
      </div>
      <div class="field">
        <label>LPO Issue Date</label>
        <input id="m_lpodate" type="date" value="${escapeHtmlAttr(it.lpoDate ?? "")}" />
      </div>
      <div class="field">
        <label>LPO Number</label>
        <input id="m_lpono" value="${escapeHtmlAttr(it.lpoNo ?? "")}" />
      </div>
      <div class="field full">
        <label>Payment Status</label>
        <input id="m_payment" value="${escapeHtmlAttr(it.payment ?? "")}" />
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
    if (!desc){ toast("Missing", "Description is required."); return; }
    it.sno = valueOrNull($("m_sno").value);
    it.ref = $("m_ref").value.trim();
    it.rev = $("m_rev").value.trim();
    it.desc = desc;
    it.mfg = $("m_mfg").value.trim();
    it.approval = $("m_approval").value.trim();
    it.planned = $("m_planned").value;
    it.estQty = valueOrNull($("m_est").value);
    it.ordQty = valueOrNull($("m_ord").value);
    it.prNo = $("m_prno").value.trim();
    it.prDate = $("m_prdate").value;
    it.prStatus = $("m_prstatus").value.trim();
    it.lpoDate = $("m_lpodate").value;
    it.lpoNo = $("m_lpono").value.trim();
    it.payment = $("m_payment").value.trim();
    it.section = $("m_section").value;
    if (isEdit){
      const idx = p.items.findIndex(x => x.id === it.id);
      if (idx >= 0) p.items[idx] = it;
      toast("Saved", "Item updated.");
    } else {
      p.items.push(it);
      toast("Added", "Item created.");
    }
    // ensure section exists
    if (it.section && it.section !== "All"){
      ensureSection(p, it.section);
    }
    saveState();
    closeModal();
    render();
  };
}

function deleteItem(itemId){
  const p = getActiveProject();
  if (!p) return;
  p.items = (p.items || []).filter(x => x.id !== itemId);
  saveState();
  renderItems();
  toast("Deleted", "Item removed.");
}

/* -------------------- Project modals -------------------- */

function openProjectModal(existing=null){
  openModal(existing ? "Edit Project" : "New Project");
  const p = existing || { name:"", refNo:"", client:"", consultant:"", contractor:"" };
  $("modalBody").innerHTML = `
    <div class="grid">
      <div class="field full">
        <label>Project Name (required)</label>
        <input id="m_pName" value="${escapeHtmlAttr(p.name)}" placeholder="e.g., Project A" />
      </div>
      <div class="field">
        <label>Project Reference No. (required)</label>
        <input id="m_pRef" value="${escapeHtmlAttr(p.refNo)}" placeholder="e.g., PRJ-001" />
      </div>
      <div class="field">
        <label>Client (optional)</label>
        <input id="m_pClient" value="${escapeHtmlAttr(p.client)}" placeholder="e.g., Client Name" />
      </div>
      <div class="field">
        <label>Consultant (optional)</label>
        <input id="m_pConsultant" value="${escapeHtmlAttr(p.consultant)}" placeholder="e.g., Consultant" />
      </div>
      <div class="field">
        <label>Main Contractor (optional)</label>
        <input id="m_pContractor" value="${escapeHtmlAttr(p.contractor)}" placeholder="e.g., Contractor" />
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
    if (!name){ toast("Missing", "Project Name is required."); return; }
    if (!refNo){ toast("Missing", "Project Reference No. is required."); return; }
    if (existing){
      existing.name = name;
      existing.refNo = refNo;
      existing.client = $("m_pClient").value.trim();
      existing.consultant = $("m_pConsultant").value.trim();
      existing.contractor = $("m_pContractor").value.trim();
    } else {
      const np = {
        id: uid(),
        name, refNo,
        client: $("m_pClient").value.trim(),
        consultant: $("m_pConsultant").value.trim(),
        contractor: $("m_pContractor").value.trim(),
        createdAt: nowISO(),
        sections: [{ name:"All", color:"#94a3b8", locked:true }],
        items: []
      };
      state.projects.push(np);
      activeProjectId = np.id;
      state.activeProjectId = activeProjectId;
      activeSection = "All";
    }
    saveState();
    closeModal();
    render();
    toast("Saved", "Project updated.");
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
        <li><b>${escapeHtml(p.name)}</b></li>
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
    activeSection = "All";
    saveState();
    closeModal();
    render();
    toast("Deleted", "Project removed.");
  };
}

/* -------------------- Notifications -------------------- */

function renderNotifications(){
  const p = getActiveProject();
  const panel = $("notiPanel");
  if (!p){ panel.classList.add("hidden"); return; }
  const upcoming = getUpcoming(p);
  const filtered = (activeSection && activeSection !== "All")
    ? upcoming.filter(u => u.item.section === activeSection)
    : upcoming;
  if (!filtered.length){
    panel.classList.add("hidden");
    return;
  }
  panel.classList.remove("hidden");
  panel.innerHTML = `
    <div class="notiTitle">Upcoming planned submissions</div>
    <div class="notiList">
      ${filtered.slice(0,6).map(u => `
        <div class="notiItem">
          <b>${escapeHtml(u.item.desc || "Item")}</b>
          <span class="muted"> — ${escapeHtml(u.item.planned)} (${u.daysLeft} day(s))</span>
          <span class="muted"> • ${escapeHtml(u.item.section || "—")}</span>
        </div>
      `).join("")}
      ${filtered.length > 6 ? `<div class="muted">+${filtered.length-6} more…</div>` : ""}
    </div>
  `;
  if ("Notification" in window && Notification.permission === "granted"){
    filtered.forEach(u => {
      if (u.daysLeft <= 1 && (u.item.approval || "").toUpperCase() !== "APPROVED"){
        maybeNotifyOnce(`due_${u.item.id}`, `${p.name}: Due soon`, `${u.item.desc} planned on ${u.item.planned}`);
      }
    });
  }
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
    const headerRowIndex = rows.findIndex(r => r.some(cell => String(cell||"").toUpperCase().includes("DESCRIPTION OF MATERIAL")));
    if (headerRowIndex < 0){
      toast("Import failed", "Couldn’t find the procurement header row.");
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
      const it = {
        id: uid(),
        sno: r?.[0] ?? "",
        ref: r?.[1] ?? "",
        rev: r?.[2] ?? "",
        desc: desc,
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
        section: activeSection !== "All" ? activeSection : "All"
      };
      imported.push(it);
    }
    if (!imported.length){
      toast("Import", "No item rows found to import.");
      return;
    }
    p.items = (p.items || []).concat(imported);
    saveState();
    render();
    toast("Imported", `${imported.length} item(s) imported.`);
  } catch(err){
    console.error(err);
    toast("Import failed", "Please try again with the Excel file.");
  }
}

function exportExcel(){
  const p = getActiveProject();
  if (!p) return;
  const items = getFilteredItemsForExport(p);
  const header = [
    "S. NO",
    "MATERIAL SUBMITTAL REF NO.",
    "REV",
    "DESCRIPTION OF MATERIAL",
    "MANUFACTURER / SUPPLIER",
    "APPROVAL STATUS",
    "PLANNED DATE OF SUBMISSION",
    "EST QTY",
    "ORDERED QTY",
    "PR NUMBER",
    "DATE RAISED",
    "PR STATUS",
    "LPO ISSUE DATE",
    "LPO NUMBER",
    "PAYMENT STATUS",
    "SECTION"
  ];
  const aoa = [header].concat(items.map(it => ([
    it.sno ?? "",
    it.ref ?? "",
    it.rev ?? "",
    it.desc ?? "",
    it.mfg ?? "",
    it.approval ?? "",
    it.planned ?? "",
    it.estQty ?? "",
    it.ordQty ?? "",
    it.prNo ?? "",
    it.prDate ?? "",
    it.prStatus ?? "",
    it.lpoDate ?? "",
    it.lpoNo ?? "",
    it.payment ?? "",
    it.section ?? ""
  ])));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Export");
  const filename = safeFileName(`${p.name}_${activeSection || "All"}_export.xlsx`);
  XLSX.writeFile(wb, filename);
  toast("Export", "Excel downloaded.");
}

function exportPDF(){
  const p = getActiveProject();
  if (!p) return;
  const items = getFilteredItemsForExport(p);
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation:"landscape", unit:"pt", format:"a4" });
  doc.setFontSize(14);
  doc.text(`${p.name} — Procurement Log`, 40, 40);
  doc.setFontSize(10);
  doc.text(`Section: ${activeSection || "All"}    Generated: ${new Date().toLocaleString()}`, 40, 58);
  const head = [[
    "#","Submittal Ref","Rev","Description","Supplier","Approval","Planned","Est","Ord","PR No","PR Date","PR Status","LPO Date","LPO No","Payment","Section"
  ]];
  const body = items.map((it, idx) => ([
    it.sno ?? (idx+1),
    it.ref ?? "",
    it.rev ?? "",
    it.desc ?? "",
    it.mfg ?? "",
    it.approval ?? "",
    it.planned ?? "",
    it.estQty ?? "",
    it.ordQty ?? "",
    it.prNo ?? "",
    it.prDate ?? "",
    it.prStatus ?? "",
    it.lpoDate ?? "",
    it.lpoNo ?? "",
    it.payment ?? "",
    it.section ?? ""
  ]));
  doc.autoTable({
    head,
    body,
    startY: 75,
    styles: { fontSize: 8, cellPadding: 4 },
    headStyles: { fillColor: [20, 184, 166] },
    theme: "grid",
    margin: { left: 40, right: 40 }
  });
  const filename = safeFileName(`${p.name}_${activeSection || "All"}_export.pdf`);
  doc.save(filename);
  toast("Export", "PDF downloaded.");
}

/* -------------------- Helpers -------------------- */

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
  const sections = normalizeSections(p);
  if (sections.some(s => s.name === name)) return;
  p.sections = sections.filter(s => s.name !== "All");
  p.sections.push({ name, color: colorFor(name), locked:false });
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
  return `<span class="badge ${info.cls}">${escapeHtml(planned)} ${info.suffix}</span>`;
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
  const already = sessionStorage.getItem(k);
  if (already) return;
  sessionStorage.setItem(k, "1");
  new Notification(title, { body });
}

function getFilteredItemsForExport(p){
  const q = ($("itemSearch").value || "").trim().toLowerCase();
  const status = $("statusFilter").value || "";
  const sortMode = $("sortBy").value || "desc_asc";
  let items = (p.items || []).slice();
  if (activeSection && activeSection !== "All"){
    items = items.filter(it => (it.section || "") === activeSection);
  }
  if (q){
    items = items.filter(it => {
      const hay = [it.desc, it.ref, it.mfg, it.approval, it.prNo, it.lpoNo, it.payment, it.prStatus].join(" ").toLowerCase();
      return hay.includes(q);
    });
  }
  if (status){
    items = items.filter(it => (it.approval || "").toUpperCase() === status);
  }
  items.sort((a,b)=> sortCompare(a,b,sortMode));
  return items;
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
  el.innerHTML = `<div class="toastTitle">${escapeHtml(title)}</div><div class="toastText">${escapeHtml(text)}</div>`;
  wrap.appendChild(el);
  setTimeout(()=>{ el.style.opacity="0"; el.style.transform="translateY(6px)"; }, 2600);
  setTimeout(()=>{ el.remove(); }, 3200);
}

function loadState(){
  try{
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return { projects:[], activeProjectId:null };
    const parsed = JSON.parse(raw);
    if (!parsed.projects) parsed.projects = [];
    return parsed;
  }catch{
    return { projects:[], activeProjectId:null };
  }
}
function saveState(){
  localStorage.setItem(LS_KEY, JSON.stringify(state));
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
  return (s || "export").toString().replace(/[^\w\-]+/g, "_").replace(/_+/g, "_").replace(/^_+|_+$/g, "");
}
function escapeHtml(str){
  return String(str ?? "").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;");
}
function escapeHtmlAttr(str){ return escapeHtml(str).replaceAll("\n"," "); }
function valueOrNull(v){ const t = String(v ?? "").trim(); return t === "" ? "" : t; }
function toISODate(v){ if (!v) return ""; const s = String(v).trim(); if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s; const d = new Date(s); if (!isNaN(d.getTime())){ const yyyy = d.getFullYear(); const mm = String(d.getMonth()+1).padStart(2,"0"); const dd = String(d.getDate()).padStart(2,"0"); return `${yyyy}-${mm}-${dd}`; } return ""; }
function itemSeed(x){ return { id: uid(), sno: x.sno ?? "", ref: x.ref ?? "", rev: x.rev ?? "", desc: x.desc ?? "", mfg: x.mfg ?? "", approval: (x.approval ?? "PENDING").toString().toUpperCase(), planned: x.planned ?? "", estQty: x.estQty ?? "", ordQty: x.ordQty ?? "", prNo: x.prNo ?? "", prDate: x.prDate ?? "", prStatus: x.prStatus ?? "", lpoDate: x.lpoDate ?? "", lpoNo: x.lpoNo ?? "", payment: x.payment ?? "", section: x.section ?? "All" }; }