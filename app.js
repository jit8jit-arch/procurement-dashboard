/* Professional Procurement Dashboard (v3)
   FIXES:
   - Modal uses .show class (no stuck popup)
   - ESC closes, click outside closes
   - Focus trap in modal
   - Body scroll lock while modal open

   FEATURES:
   - Projects + per-project saved settings (columns)
   - Sections with auto colors
   - Advanced filters (search, status, multi-section, date range, overdue)
   - Pagination
   - Export Excel + PDF (exports filtered view)
   - Import Excel
   - Backup/Restore JSON
   - Row edit/delete + confirm delete + double-click row to edit
   - No sample project details
*/

const LS_KEY = "proc_dash_v3";

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

const DEFAULT_COLUMNS = [
  { key:"sno",      label:"#", show:true },
  { key:"ref",      label:"Submittal Ref", show:true },
  { key:"rev",      label:"Rev", show:true },
  { key:"desc",     label:"Description of Material", show:true, wrap:true
