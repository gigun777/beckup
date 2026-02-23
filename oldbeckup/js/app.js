// js/app.js
// BUILD: v12.6.14.1 (device env: PPI/DPR/type + nested subjournals)
import { openDB, cfgGet, cfgSet, getRows, addRow, putRow, deleteRow, clearRows, clearAllRows,
         getAllCases, addCase, getCaseRows, addCaseRow, putCaseRow, deleteCaseRow, clearAllCasesAndRows } from "./db.js";
import { DEFAULT_SHEETS, CASE_DESC_COLUMNS, uaDateToday, parseUAdate, excelSerialToUAdate, isIntegerString, nowStamp, safeName, normalizeForExport } from "./schema.js";
import { $, el, showMenu, hideMenu, modalOpen, btn, confirmDeleteNumber, downloadBlob } from "./ui.js";
import { unzipStoreEntries, unzipEntries } from "./zip.js";
import { exportDOCXTable, exportXLSXTable } from "./office.js";
import { exportPDFTable } from "./pdfgen.js";
import { exportAllZipJSON, exportJournalAsJSON, exportJournalAsDOCX, exportJournalAsXLSX, exportJournalAsPDF, makeJournalExportFileName, makeCaseExportFileName } from "./export.js";
import { ensureDefaultTransferRules, ensureDefaultTransferTemplates, getTransferTemplates } from "./transfer.js?v=12.6.11";
import { getAllSheets, saveUserSheets, saveAllSheets, getSheetSettings, saveSheetSettings, getAddFieldConfig, saveAddFieldConfig, buildSettingsUI } from "./settings.js?v=12.6.11";
import { initDeviceEnv, getDeviceEnv } from "./device.js";

// Device environment (PPI / type / DPR) for adaptive UX
const deviceEnv = initDeviceEnv();

await openDB();
await ensureDefaultTransferRules();

const state = {
  deviceEnv, // runtime screen/device info (PPI/DPR/type)
  uiPrefs:{ btnScale: 1 },
  // Stage 2 (minimal): fixed Level="Адміністратор" + one root space ("Простір 1").
  level:"admin",
  spaceId:"space1", // single root space
  spaces:[],          // cached spaces list
  // Journal tree (nested subjournals)
  jtree:null,         // { nodes: {id:node}, topIds:[id...] }
  journalPath:[],     // [topJournalId, childId, childId, ...]
  mode:"sheet",
  sheetKey: DEFAULT_SHEETS[0].key,
  caseId: null,
  search:"",
  sort: { col:null, dir:1 },
  sheets: [],
  sheetSettings: {},
  addFieldsCfg: {},
  settingsDirty:false,
  settingsTab:"sheets",
  selectionMode:false,
  selectedRowIds:new Set(),
};

const levelSelect = $("#levelSelect");
const spaceSelect = $("#spaceSelect");
const btnAddSpace = $("#btnAddSpace");
const sheetSelect = $("#sheetSelect");
const btnAddChild0 = $("#btnAddChild0");
const subjournalChain = $("#subjournalChain");
const caseSelect = $("#caseSelect");
const btnCaseBack = $("#btnCaseBack");
const table = $("#table");
const cards = $("#cards");
const menu = $("#menu");
const sideHint = $("#sideHint");

// Keep device env in state (for future adaptive sizing)
window.addEventListener("deviceenv:change", (e)=>{ state.deviceEnv = e.detail; });

function isCustomKey(key){ return key.startsWith("custom_"); }

function toggleRowSelection(rowId){
  if(!state.selectionMode) return;
  const set = state.selectedRowIds;
  if(set.has(rowId)) set.delete(rowId);
  else set.add(rowId);
  state.selectedRowIds = new Set(set);
  // update transfer button visibility
  $("#btnTransferSelected").style.display = (state.selectionMode?"inline-block":"none");
  render();
}

async function loadConfig(){
  state.sheets = await getAllSheets();
  state.sheetSettings = await getSheetSettings();
  state.addFieldsCfg = await getAddFieldConfig();

  // UI preferences
  {
    const rawScale = await cfgGet("ui_btn_scale");
    const s = Number(rawScale);
    const btnScale = (Number.isFinite(s) && s>=0.6 && s<=2.2) ? s : 1;
    state.uiPrefs.btnScale = btnScale;
    document.documentElement.style.setProperty("--btn-scale", String(btnScale));
  }

  const saved = await cfgGet("last_view");
  if(saved){
    state.mode = "sheet";
    state.spaceId = saved.spaceId || state.spaceId || "space1";
    state.journalPath = Array.isArray(saved.journalPath) ? saved.journalPath : [];
    state.sheetKey = saved.sheetKey || state.sheetKey;
    state.caseId = null;
  }
  state.spaces = await ensureSpaces();
  // Load per-space journal tree (each space has its own independent hierarchy)
  state.jtree = await ensureJournalTree(state.spaceId, state.sheets);
  // ensure there is always a valid top journal selected
  if(!state.journalPath.length){
    const topId = state.jtree?.topIds?.[0];
    if(topId) state.journalPath = [topId];
  }
}
async function saveView(){
	await cfgSet("last_view", {spaceId:state.spaceId, journalPath:state.journalPath, mode:"sheet", sheetKey:state.sheetKey, caseId:null});
}
function currentSheet(){ return state.sheets.find(s=>s.key===activeSheetKey()); }

// --- Spaces & Journal instances (Stage 2: nested subjournals) ---
// Space is fixed ("Простір 1") for now. Journals/subjournals are *instances*.
function currentInstanceId(){ return state.journalPath[state.journalPath.length-1] || null; }
function currentDataKey(){
  const id = currentInstanceId();
  return id ? `${state.spaceId}::${id}` : `${state.spaceId}::root`;
}
function activeSheetKey(){
  const id = currentInstanceId();
  const n = id ? state.jtree?.nodes?.[id] : null;
  return n?.sheetKey || state.sheetKey;
}

// Data key for a top-level (space) journal by sheetKey (used by transfer/import/export routines).
function journalKeyForSheet(sheetKey){
  return `${state.spaceId}::root:${sheetKey}`;
}

async function ensureSpaces(){
  let spaces = await cfgGet("spaces_v1");
  if(!Array.isArray(spaces) || !spaces.length){
    spaces = [
      {id:"space1", name:"Простір 1", parentId:null, kind:"space", meta:{}},
    ];
    await cfgSet("spaces_v1", spaces);
  }
  // Only root spaces for now (folders-in-folders later)
  return spaces.filter(s=>s && s.kind==="space");
}

function norm(s){ return (s||"").toString().trim().toLowerCase(); }

function nodeById(id){ return state.jtree?.nodes?.[id] || null; }
function childrenOf(id){
  const n = nodeById(id);
  return (n?.children||[]).map(cid=>nodeById(cid)).filter(Boolean);
}

async function saveJournalTree(spaceId){
  if(state.jtree) await cfgSet(`journal_tree_v1:${spaceId}`, state.jtree);
}

// Journal tree (nested subjournals) is stored globally for the fixed space.
// Each node is a *journal instance* with its own isolated data.
async function ensureJournalTree(spaceId, sheets){
  let tree = await cfgGet(`journal_tree_v1:${spaceId}`);
  if(!tree || typeof tree!=="object" || !tree.nodes){
    tree = { nodes:{}, topIds:[] };
  }
  if(!Array.isArray(tree.topIds)) tree.topIds = [];
  if(!tree.nodes || typeof tree.nodes!=="object") tree.nodes = {};

  // Ensure deterministic top-level journals for all existing sheets (6 defaults + admin added)
  const sheetOrder = (sheets||[]).map(s=>s.key);
  for(let i=0;i<sheetOrder.length;i++){
    const key = sheetOrder[i];
    const id = `root:${key}`;
    if(!tree.nodes[id]){
      tree.nodes[id] = {
        id,
        parentId:null,
        sheetKey:key,
        numPath:[i+1],
        title: (sheets||[]).find(s=>s.key===key)?.name || key,
        children:[],
      };
    }
    if(!tree.topIds.includes(id)) tree.topIds.push(id);
  }
  // Keep topIds order consistent with sheets order
  tree.topIds = sheetOrder.map(k=>`root:${k}`).filter(id=>tree.nodes[id]);

  await cfgSet(`journal_tree_v1:${spaceId}`, tree);
  return tree;
}

function nodeTitle(n){
  const num = Array.isArray(n?.numPath) ? n.numPath.join('.') : '';
  const base = n?.title || '';
  return num ? `${num} ${base}`.trim() : base;
}



function ensureSimplifiedConfig(entity){
  if(!entity) return;
  if(!entity.simplified) entity.simplified = { enabled:false, on:false, activeTemplateId:null, templates:[] };
  if(typeof entity.simplified.enabled!=="boolean") entity.simplified.enabled = false;
  if(typeof entity.simplified.on!=="boolean") entity.simplified.on = false;
  if(!Array.isArray(entity.simplified.templates)) entity.simplified.templates = [];
  // No default template: normal table view is the default when simplified view is OFF.
  // User creates simplified templates explicitly in the constructor.
  if(entity.simplified.enabled && !entity.simplified.activeTemplateId && entity.simplified.templates.length){
    entity.simplified.activeTemplateId = entity.simplified.templates[0].id;
  }
}
function currentSimplifiedEntity(){
  if(state.mode==="case"){
    // Stage 1: cases do not yet have simplified settings; return null.
    return null;
  }
  return currentSheet();
}
function updateSimplifiedToggle(){
  const btn = document.getElementById("btnSimpleView");
  const sel = document.getElementById("simpleViewTemplate");
  if(!btn) return;
  const ent = currentSimplifiedEntity();
  if(!ent){
    btn.disabled=true; btn.classList.remove("btn-toggle-on"); btn.title="Спрощений перегляд (немає профілю)";
    if(sel){ sel.style.display="none"; sel.innerHTML=""; }
    return;
  }
  ensureSimplifiedConfig(ent);
  const cfg = ent.simplified;
  const hasProfile = cfg.enabled && (cfg.templates||[]).length>0;
  btn.disabled = !hasProfile;
  btn.classList.toggle("btn-toggle-on", !!(hasProfile && cfg.on));
  btn.textContent = cfg.on ? "☰ Спрощено: ON" : "☰ Спрощено";
  btn.title = hasProfile ? "Спрощений перегляд" : "Спрощений перегляд (не налаштовано)";

  // Stage 2: template switcher (only when profile exists and more than 1 template)
  if(sel){
    if(hasProfile && (cfg.templates||[]).length>1){
      sel.style.display = "";
      sel.innerHTML = "";
      for(const t of cfg.templates){
        sel.appendChild(el("option",{value:t.id, textContent:t.name||t.id}));
      }
      sel.value = cfg.activeTemplateId || cfg.templates[0].id;
      sel.onchange = async ()=>{
        cfg.activeTemplateId = sel.value;
        await saveUserSheets(state.sheets);
      };
    } else {
      sel.style.display = "none";
      sel.innerHTML = "";
      sel.onchange = null;
    }
  }
}
async function toggleSimplifiedView(){
  const ent = currentSimplifiedEntity();
  if(!ent) return;
  ensureSimplifiedConfig(ent);
  if(!(ent.simplified.enabled && ent.simplified.templates.length)) return;
  ent.simplified.on = !ent.simplified.on;
  await saveUserSheets(state.sheets); // persist
  updateSimplifiedToggle();
  try{ if(state.mode==='sheet') await renderSheet(); }catch(e){}
}



function getActiveSimplifiedTemplate(entity){
  if(!entity?.simplified) return null;
  const cfg = entity.simplified;
  const id = cfg.activeTemplateId || (cfg.templates?.[0]?.id || null);
  if(!id) return null;
  return (cfg.templates||[]).find(t=>t.id===id) || null;
}
function hexToRgba(hex, alpha){
  try{
    if(!hex) return `rgba(0,0,0,${alpha})`;
    let h = hex.trim();
    if(h.startsWith("#")) h=h.slice(1);
    if(h.length===3) h=h.split("").map(c=>c+c).join("");
    const r=parseInt(h.slice(0,2),16), g=parseInt(h.slice(2,4),16), b=parseInt(h.slice(4,6),16);
    const a=Math.max(0,Math.min(1,Number(alpha)));
    return `rgba(${r},${g},${b},${a})`;
  }catch(e){
    const a=Math.max(0,Math.min(1,Number(alpha)));
    return `rgba(0,0,0,${a})`;
  }
}
function ensureSimplifiedTheme(entity){
  ensureSimplifiedConfig(entity);
  const t = entity.simplified.theme || (entity.simplified.theme = {});
  if(!Number.isFinite(t.radius)) t.radius = 16;
  if(typeof t.showBorders!=="boolean") t.showBorders = false;
  if(typeof t.glass!=="boolean") t.glass = true;
  if(typeof t.gradient!=="boolean") t.gradient = false;
  if(typeof t.cardColor!=="string") t.cardColor = "#ff3b30";
  if(typeof t.cardOpacity!=="number") t.cardOpacity = 0.92;
  if(typeof t.gradFrom!=="string") t.gradFrom = "#ff3b30";
  if(typeof t.gradTo!=="string") t.gradTo = "#ff9500";
  if(typeof t.bgColor!=="string") t.bgColor = "";
  if(typeof t.borderColor!=="string") t.borderColor = "rgba(255,255,255,0.30)";
  if(!Number.isFinite(t.blur)) t.blur = 18;
  if(typeof t.customCss!=="string") t.customCss = "";
  // Conditional card background rules (per-row)
  if(typeof t.cardBgRulesEnabled!=="boolean") t.cardBgRulesEnabled = false;
  if(!Array.isArray(t.cardBgRules)) t.cardBgRules = [];
  return t;
}
function applySimplifiedTheme(entity){
  const th = ensureSimplifiedTheme(entity);
  // background for the cards area
  if(cards){
    cards.classList.toggle("sv-bg", !!th.bgColor);
    cards.style.setProperty("--sv-bg", th.bgColor || "transparent");
    cards.style.setProperty("--sv-radius", (th.radius||16)+"px");
    cards.style.setProperty("--sv-border", th.borderColor || "rgba(255,255,255,0.30)");
    cards.style.setProperty("--sv-blur", (th.blur||18)+"px");
    // card bg with opacity
    const bg = hexToRgba(th.cardColor || "#ff3b30", (typeof th.cardOpacity==="number") ? th.cardOpacity : 0.92);
    cards.style.setProperty("--sv-card-bg", bg);
    const g1 = hexToRgba(th.gradFrom || th.cardColor || "#ff3b30", (typeof th.cardOpacity==="number") ? th.cardOpacity : 0.92);
    const g2 = hexToRgba(th.gradTo || "#ff9500", (typeof th.cardOpacity==="number") ? th.cardOpacity : 0.92);
    cards.style.setProperty("--sv-grad-from", g1);
    cards.style.setProperty("--sv-grad-to", g2);
  }
  // inject custom css
  const id="svCustomStyle";
  let st=document.getElementById(id);
  if(th.customCss && th.customCss.trim()){
    if(!st){ st=document.createElement("style"); st.id=id; document.head.appendChild(st); }
    st.textContent = th.customCss;
  } else if(st){ st.remove(); }
  return th;
}

function computeBlockValue(sheet, row, block){
  const srcs = Array.isArray(block?.sources) ? block.sources : [];
  const vals = srcs.map(i=>{
    const idx = Number(i);
    const col = sheet.columns?.[idx];
    const name = col?.name;
    const v = name ? (row?.data?.[name] ?? "") : "";
    return (v===null || v===undefined) ? "" : String(v);
  });
  const op = block?.op || "concat";
  const delim = (typeof block?.delimiter==="string") ? block.delimiter : " ";
  if(op==="newline") return vals.join("\n");
  if(op==="seq") return vals.join("");
  return vals.join(delim);
}

function computeCellValue(sheet, row, cellCfg){
  const blocks = Array.isArray(cellCfg?.blocks) ? cellCfg.blocks : (Array.isArray(cellCfg) ? cellCfg : []);
  if(!blocks.length) return "";
  const joinAll = cellCfg?.joinAll || {op:"newline", delimiter:""};
  const joins = Array.isArray(cellCfg?.joins) ? cellCfg.joins : [];
  let out = computeBlockValue(sheet,row,blocks[0]);
  for(let i=1;i<blocks.length;i++){
    const j = joins[i-1] || joinAll || {op:"newline", delimiter:""};
    const op = j.op || "newline";
    const delim = (typeof j.delimiter==="string") ? j.delimiter : ((joinAll?.delimiter)||" ");
    if(op==="seq") out += "";
    else if(op==="concat") out += delim;
    else out += "\n";
    out += computeBlockValue(sheet,row,blocks[i]);
  }
  return out;
}


function parseDateDMY(s){
  // expects DD.MM.YY or DD.MM.YYYY
  if(typeof s!=="string") return null;
  const m = s.trim().match(/^([0-3]?\d)\.([01]?\d)\.(\d{2}|\d{4})$/);
  if(!m) return null;
  const d = Number(m[1]), mo = Number(m[2]), yRaw = Number(m[3]);
  const y = (m[3].length===2) ? (2000 + yRaw) : yRaw;
  if(!(d>=1&&d<=31&&mo>=1&&mo<=12&&y>=1900&&y<=2100)) return null;
  // basic date validity
  const dt = new Date(y, mo-1, d);
  if(dt.getFullYear()!==y || (dt.getMonth()+1)!==mo || dt.getDate()!==d) return null;
  return dt;
}
function isNumericStr(s){
  if(typeof s!=="string") return false;
  const t = s.trim();
  if(!t) return false;
  return /^-?\d+(\.\d+)?$/.test(t);
}
function getRowValueByColIndex(sheet, row, colIndex){
  const idx = Number(colIndex);
  const col = sheet.columns?.[idx];
  const name = col?.name;
  const v = name ? (row?.data?.[name] ?? "") : "";
  return (v===null || v===undefined) ? "" : String(v);
}
function resolveCardBgHex(sheet, row, th){
  // default
  let base = th.cardColor || "#ff3b30";
  if(!th.cardBgRulesEnabled) return base;
  const rules = Array.isArray(th.cardBgRules) ? th.cardBgRules : [];
  for(const r of rules){
    if(!r || typeof r!=="object") continue;
    const v = getRowValueByColIndex(sheet, row, r.col);
    const vv = (v ?? "").toString();
    const test = r.test || "notempty";
    let ok=false;
    if(test==="empty") ok = vv.trim()==="";
    else if(test==="notempty") ok = vv.trim()!=="";
    else if(test==="isnumber") ok = isNumericStr(vv);
    else if(test==="isdate") ok = !!parseDateDMY(vv);
    else if(test==="equals") ok = vv.trim() === String(r.value ?? "").trim();
    else if(test==="contains") ok = vv.toLowerCase().includes(String(r.value ?? "").toLowerCase());
    if(ok){
      const col = (typeof r.color==="string" && r.color.trim()) ? r.color.trim() : base;
      return col;
    }
  }
  return base;
}

function renderSimplifiedCardsForSheet(sheet, rows){
  const t = getActiveSimplifiedTemplate(sheet);
  if(!t?.layout?.rows || !t?.layout?.cols) return false;
  applySimplifiedTheme(sheet);
  // Toggle displays
  if(cards) cards.style.display="";
  if(table) table.style.display="none";
  cards.innerHTML="";
  const th = ensureSimplifiedTheme(sheet);
  const showBorders = !!th.showBorders;
  for(const row of rows){
    const card = el("div",{className:"sv-card"});
    const bgHex = resolveCardBgHex(sheet, row, th);
    const op = (typeof th.cardOpacity==="number") ? th.cardOpacity : 0.92;
    const bgRgba = hexToRgba(bgHex, op);
    card.style.setProperty("--sv-card-bg", bgRgba);
    if(th.gradient){
      // If a rule matched (different from default), make gradient a solid color.
      const g1 = hexToRgba(bgHex, op);
      const g2 = g1;
      card.style.setProperty("--sv-grad-from", g1);
      card.style.setProperty("--sv-grad-to", g2);
    }
    if(th.glass) card.classList.add("sv-glass");
    if(th.gradient) card.classList.add("sv-gradient");
    if(showBorders) card.classList.add("sv-show-borders");
    const grid = el("div",{className:"sv-card-grid"});
    grid.style.gridTemplateColumns = `repeat(${t.layout.cols}, minmax(0, 1fr))`;
    // build cells row-major
    for(let r=0;r<t.layout.rows;r++){
      for(let c=0;c<t.layout.cols;c++){
        const key = `${r}-${c}`;
        const cfg = t.layout.cells?.[key] || {blocks:[]};
        const val = computeCellValue(sheet,row,cfg);
        const cell = el("div",{className:"sv-card-cell",textContent: val});
        // For border rendering in grid: add data attributes
        cell.dataset.r=String(r); cell.dataset.c=String(c);
        grid.appendChild(cell);
      }
    }
    card.appendChild(grid);
    cards.appendChild(card);
  }
  // Apply background on main wrap (optional)
  return true;
}

function applyDefaultSortForSheet(sheet){
  // Reset sort to sheet defaults. User can override by clicking headers.
  state.sort = { col:null, dir:1 };
  if(!sheet) return;

  const ds = sheet.defaultSort;
  if(ds && ds.col){
    const exists = (sheet.columns || []).some(c => c.name === ds.col);
    if(exists){
      state.sort.col = ds.col;
      state.sort.dir = (ds.dir === 'desc') ? -1 : 1;
      return;
    }
  }
  // Fallback to legacy orderColumn behavior handled in render() when state.sort.col is null.
}


function visibleColumns(sheet){
  const cfg = state.sheetSettings[sheet.key] || {};
  const hidden = new Set(cfg.hiddenCols || []);
  return sheet.columns.map(c=>c.name).filter(n=>!hidden.has(n));
}
function setSideHint(text){ sideHint.textContent = text || ""; }

$("#btnSettings").onclick=(e)=>{
  e.stopPropagation();
  openSettings();
};

$("#btnImportExport").onclick=(e)=>{
  e.stopPropagation();
  openImportExportWindow();
};
document.addEventListener("click",(e)=>{
  if(menu.style.display==="block" && !menu.contains(e.target) && e.target!==$("#btnSettings")) hideMenu(menu);
});
menu.addEventListener("click", async (e)=>{
  const b = e.target.closest("button");
  if(!b) return;
  const act = b.dataset.action;
  hideMenu(menu);
  if(act==="settingsPanel") return openSettings();
  if(act==="exportCurrent") return exportCurrentFlow();
  if(act==="exportAllZip") return exportAllFlow();
  if(act==="importJson") return $("#fileImportJson").click();
  if(act==="importZip") return $("#fileImportZip").click();
  if(act==="importXlsx") return $("#fileImportXlsx").click();
  if(act==="print") return window.print();
  if(act==="clearCurrent") return clearCurrent();
  if(act==="clearAll") return clearAll();
});

$("#btnSimpleView").onclick=()=>toggleSimplifiedView();
$("#searchInput").addEventListener("input",(e)=>{ state.search=e.target.value||""; render(); });
$("#btnAdd").onclick=()=>addFlow();


// Jump scroll buttons (top/bottom)
const scrollJump = document.getElementById("scrollJump");
const btnScrollTop = document.getElementById("btnScrollTop");
const btnScrollBottom = document.getElementById("btnScrollBottom");
function getScrollEl(){ return document.scrollingElement || document.documentElement; }
function updateScrollJump(){
  if(!scrollJump) return;
  const se = getScrollEl();
  const y = window.scrollY || (se ? se.scrollTop : 0) || 0;
  // show after a small scroll so it doesn't clutter the UI
  if(y > 180) scrollJump.classList.remove("hidden");
  else scrollJump.classList.add("hidden");
}
if(btnScrollTop){
  btnScrollTop.addEventListener("click", ()=>window.scrollTo({top:0, behavior:"smooth"}));
}
if(btnScrollBottom){
  btnScrollBottom.addEventListener("click", ()=>{
    const se = getScrollEl();
    const max = Math.max(se.scrollHeight - se.clientHeight, 0);
    window.scrollTo({top:max, behavior:"smooth"});
  });
}
window.addEventListener("scroll", updateScrollJump, {passive:true});
window.addEventListener("resize", updateScrollJump);
setTimeout(updateScrollJump, 0);

// Selection mode (multi-row transfer)
$("#btnSelect").onclick=async ()=>{
  if(state.mode!=="sheet" && state.mode!=="case") return alert("Режим вибору працює тільки в листі або в описі справи.");
  if(!state.selectionMode){
    state.selectionMode=true;
    state.selectedRowIds=new Set();
    $("#btnTransferSelected").style.display="inline-block";
    $("#btnSelect").textContent="☑ Вибір*";
    render();
    return;
  }
  const op = await modalOpen({
    title:"Режим вибору",
    bodyNodes:[el("div",{className:"muted",textContent:`Вибрано: ${state.selectedRowIds.size}`})],
    actions:[
      btn("Скасувати","cancel","btn"),
      btn("Вибрати всі","all","btn btn-primary"),
      btn("Зняти всі","none","btn"),
      btn("Вийти","exit","btn")
    ]
  });
  if(op.type==="all"){
    if(state.mode==="sheet"){
      const rows=await getRows(currentDataKey());
      state.selectedRowIds=new Set(rows.map(r=>r.id));
    }else{
      const rows=await getCaseRows(state.caseId);
      state.selectedRowIds=new Set(rows.map(r=>r.id));
    }
    render();
  }else if(op.type==="none"){
    state.selectedRowIds=new Set();
    render();
  }else if(op.type==="exit"){
    state.selectionMode=false;
    state.selectedRowIds=new Set();
    $("#btnTransferSelected").style.display="none";
    $("#btnSelect").textContent="☑ Вибір";
    render();
  }
};

$("#btnTransferSelected").onclick=async ()=>{
  if(!state.selectionMode || !state.selectedRowIds.size) return alert("Немає вибраних строк.");
  if(state.mode==="sheet"){
    const sheet=currentSheet();
    const all=await getRows(currentDataKey());
    const selected = all.filter(r=>state.selectedRowIds.has(r.id));
    if(!selected.length) return alert("Немає вибраних строк.");
    await transferMultipleFlow(sheet, selected);
  }else{
    const all=await getCaseRows(state.caseId);
    const selected = all.filter(r=>state.selectedRowIds.has(r.id));
    if(!selected.length) return alert("Немає вибраних строк.");
    await transferMultipleCaseFlow(state.caseId, selected);
  }
};
$("#fileImportJson").addEventListener("change",(e)=>importJsonFile(e.target));
$("#fileImportZip").addEventListener("change",(e)=>importZipFile(e.target));
$("#fileImportXlsx").addEventListener("change",(e)=>importXlsxFile(e.target));
$("#fileImportXlsx").addEventListener("change",(e)=>importXlsxFile(e.target));

// Space is fixed for now (no editor), so no change handler.

async function openAddChildModal(parentPath, parentDepth){
  // Adds a child journal instance under the node at parentDepth.
  // parentPath includes IDs up to that parent (inclusive).
  const parentId = parentPath[parentDepth];
  const parentNode = nodeById(parentId);
  if(!parentNode) return;

  const siblings = (parentNode.children||[]).filter(cid=>!!nodeById(cid));
  const nextIdx = siblings.length + 1;
  const numPath = (parentNode.numPath||[]).concat([nextIdx]);

  // --- Combined picker (single field + ▾) ---
  let selectedSheetKey = (state.sheets?.[0]?.key) || parentNode.sheetKey;
  const allSheets = (state.sheets||[]).slice();

  const showAllBtn = el('button', {className:'btn', textContent:'▾', title:'Показати повний список журналів', style:'width:42px; padding:0;'});
  const comboInput = el('input', {className:'input', placeholder:'Пошук / вибір журналу…', value:''});

  const list = el('div', {className:'combo-list', style:'display:none;'});
  let currentItems = [];

  function renderList(filterText){
    const q = String(filterText||'').trim().toLowerCase();
    currentItems = allSheets.filter(s=>{
      if(!q) return true;
      return String(s.title||'').toLowerCase().includes(q) || String(s.key||'').toLowerCase().includes(q);
    });
    list.innerHTML='';
    if(currentItems.length===0){
      list.appendChild(el('div', {className:'combo-item muted', textContent:'Нічого не знайдено'}));
      return;
    }
    for(const s of currentItems){
      const item = el('div', {className:'combo-item', textContent:s.title, title:s.key});
      item.onclick = ()=>{
        selectedSheetKey = s.key;
        comboInput.value = s.title;
        closeList();
      };
      list.appendChild(item);
    }
  }

  function openList(filterText){
    renderList(filterText);
    list.style.display='block';
  }
  function closeList(){ list.style.display='none'; }

  comboInput.oninput = ()=> openList(comboInput.value);
  comboInput.onfocus = ()=> openList(comboInput.value);
  comboInput.onkeydown = (e)=>{
    if(e.key==='Enter'){
      e.preventDefault();
      if(currentItems[0]){
        selectedSheetKey = currentItems[0].key;
        comboInput.value = currentItems[0].title;
        closeList();
      }
    } else if(e.key==='Escape'){
      closeList();
    }
  };
  showAllBtn.onclick = ()=>{ comboInput.value=''; openList(''); comboInput.focus(); };

  const outsideClose = (ev)=>{
    if(!list.contains(ev.target) && ev.target!==comboInput && ev.target!==showAllBtn){
      closeList();
      document.removeEventListener('mousedown', outsideClose, true);
    }
  };
  comboInput.addEventListener('focus', ()=>{
    document.addEventListener('mousedown', outsideClose, true);
  });

  const comboWrap = el('div', {className:'combo-wrap'},
    el('div', {style:'display:flex; gap:8px;'}, showAllBtn, comboInput),
    list
  );

  const nameInput = el('input', {className:'input', value:'', placeholder:'Назва (необов’язково)'});

  const res = await modalOpen({
    title: `Додати піджурнал у: ${nodeTitle(parentNode)}`,
    bodyNodes:[
      el('div', {className:'muted', textContent:'Оберіть, який саме журнал буде піджурналом. Дані піджурналу ізольовані.'}),
      el('div', {style:'height:10px'}),
      comboWrap,
      el('div', {style:'height:10px'}),
      el('div', {className:'muted', textContent:'Назва піджурналу'}),
      nameInput
    ],
    actions:[
      btn('Скасувати','cancel','btn'),
      btn('Створити','ok','btn btn-primary')
    ]
  });
  if(res?.type!=='ok') return;

  // Resolve exact match if typed
  const typed = String(comboInput.value||'').trim();
  if(typed){
    const m = (state.sheets||[]).find(s=>String(s.title||'').toLowerCase()===typed.toLowerCase());
    if(m) selectedSheetKey = m.key;
  }

  const childSheet = (state.sheets||[]).find(s=>s.key===selectedSheetKey) || (state.sheets||[])[0];
  const name = String(nameInput.value||'').trim() || (childSheet?.title || childSheet?.name || childSheet?.key || 'Піджурнал');

  const id = 'sj_' + Date.now() + '_' + Math.random().toString(16).slice(2);
  state.jtree.nodes[id] = {
    id,
    parentId: parentId,
    sheetKey: childSheet.key,
    numPath,
    title: name,
    children:[]
  };
  if(!Array.isArray(parentNode.children)) parentNode.children = [];
  parentNode.children.push(id);
  await saveJournalTree(state.spaceId);

  // Switch view to the newly created node
  state.journalPath = parentPath.slice(0, parentDepth+1).concat([id]);
  await saveView();
  renderJournalNav();
  render();
}

caseSelect.addEventListener("change", async ()=>{
  const v=caseSelect.value;
  if(!v){ state.mode="sheet"; state.caseId=null; btnCaseBack.style.display="none"; await saveView(); render(); return; }
  state.mode="case"; state.caseId=parseInt(v,10); btnCaseBack.style.display="inline-block"; await saveView(); render();
});
btnCaseBack.onclick=async ()=>{ state.mode="sheet"; state.caseId=null; caseSelect.value=""; btnCaseBack.style.display="none"; await saveView(); render(); };

$("#settingsClose").onclick=()=>closeSettings();
$("#settingsCancel").onclick=()=>closeSettings();
$("#settingsSave").onclick=()=>saveSettings();
document.querySelectorAll(".tab").forEach(t=>{
  t.onclick=()=>{
    document.querySelectorAll(".tab").forEach(x=>x.classList.remove("active"));
    t.classList.add("active");
    state.settingsTab=t.dataset.tab;
    renderSettings();
  };
});
function markDirty(){ state.settingsDirty=true; $("#settingsSave").textContent="Зберегти*"; }
function clearDirty(){ state.settingsDirty=false; $("#settingsSave").textContent="Зберегти"; }
async function openImportExportWindow(){
  const backdrop = $("#modalBackdrop");
  const t = $("#modalTitle");
  const b = $("#modalBody");
  const a = $("#modalActions");

  const close = ()=>{
    backdrop.style.display = "none";
    backdrop.onclick = null;
  };

  t.textContent = "Імпорт / Експорт";
  b.innerHTML = "";
  a.innerHTML = "";

  // Tabs
  const tabs = el("div", {className:"tabs"});
  const tabImport = el("button", {className:"tab active", textContent:"Імпорт"});
  const tabExport = el("button", {className:"tab", textContent:"Експорт"});
  const tabService = el("button", {className:"tab", textContent:"Сервіс"});
  tabs.append(tabImport, tabExport, tabService);

  const content = el("div", {className:"settings-content", style:"padding:0"});

  const setActive = (which)=>{
    [tabImport,tabExport,tabService].forEach(x=>x.classList.remove("active"));
    which.classList.add("active");
  };

  const sectionTitle = (text)=>el("div", {className:"muted", textContent:text, style:"margin:6px 0 10px"});

  const renderImport = ()=>{
    content.innerHTML = "";
    content.appendChild(sectionTitle("Імпорт у систему"));
    const row = el("div", {style:"display:flex; gap:10px; flex-wrap:wrap"});
    const bJson = btn("📥 Імпорт JSON", "imp_json", "btn btn-primary");
    const bZip = btn("📥 Імпорт ZIP", "imp_zip", "btn btn-primary");
    const bXlsx = btn("📥 Імпорт XLSX (у поточний)", "imp_xlsx", "btn");
    row.append(bJson,bZip,bXlsx);
    content.appendChild(row);

    // helper hint
    content.appendChild(el("div", {className:"muted", textContent:"Порада: ZIP — це повний імпорт, JSON — імпорт одного журналу/опису справи.", style:"margin-top:10px"}));

    bJson.onclick = ()=>{ close(); $("#fileImportJson").click(); };
    bZip.onclick  = ()=>{ close(); $("#fileImportZip").click(); };
    bXlsx.onclick = ()=>{ close(); $("#fileImportXlsx").click(); };
  };

  const renderExport = ()=>{
    content.innerHTML = "";
    content.appendChild(sectionTitle("Експорт/друк"));
    const row = el("div", {style:"display:flex; gap:10px; flex-wrap:wrap"});
    const bCur = btn("📤 Експорт поточного", "exp_cur", "btn btn-primary");
    const bAll = btn("📦 Повний експорт ВСЬОГО (ZIP→JSON)", "exp_all", "btn");
    const bPrint = btn("🖨 Друк / PDF (через друк)", "print", "btn");
    row.append(bCur,bAll,bPrint);
    content.appendChild(row);

    bCur.onclick = async ()=>{ close(); await exportCurrentFlow(); };
    bAll.onclick = async ()=>{ close(); await exportAllFlow(); };
    bPrint.onclick = ()=>{ close(); window.print(); };
  };

  const renderService = ()=>{
    content.innerHTML = "";
    content.appendChild(sectionTitle("Обережно: дії видалення"));
    const row = el("div", {style:"display:flex; gap:10px; flex-wrap:wrap"});
    const bCC = btn("🧹 Очистити поточний", "clr_cur", "btn");
    const bCA = btn("🧨 Очистити ВСЕ", "clr_all", "btn");
    bCC.classList.add("danger");
    bCA.classList.add("danger");
    row.append(bCC,bCA);
    content.appendChild(row);

    bCC.onclick = async ()=>{ close(); await clearCurrent(); };
    bCA.onclick = async ()=>{ close(); await clearAll(); };
  };

  tabImport.onclick = ()=>{ setActive(tabImport); renderImport(); };
  tabExport.onclick = ()=>{ setActive(tabExport); renderExport(); };
  tabService.onclick = ()=>{ setActive(tabService); renderService(); };

  b.appendChild(tabs);
  b.appendChild(content);
  renderImport();

  const closeBtn = btn("Закрити", "close", "btn");
  closeBtn.onclick = ()=>close();
  a.appendChild(closeBtn);

  backdrop.style.display = "flex";
  backdrop.onclick = (e)=>{ if(e.target===backdrop) close(); };
}

function openSettings(){ $("#settingsBackdrop").style.display="flex"; clearDirty(); renderSettings(); }
function closeSettings(){ $("#settingsBackdrop").style.display="none"; }
async function saveSettings(){
  const userSheets = state.sheets.filter(s=>isCustomKey(s.key));
  await saveUserSheets(userSheets);
  // Persist full schema (including renamed/changed default sheets)
  await saveAllSheets(state.sheets);
  await saveSheetSettings(state.sheetSettings);
  await saveAddFieldConfig(state.addFieldsCfg);
  clearDirty();
  await loadConfig();
  applyDefaultSortForSheet(currentSheet());
  const spaces = await ensureSpaces();
  state.spaces = spaces;
  fillSpaceSelect(spaces);
  state.jtree = await ensureJournalTree(state.spaceId, state.sheets);
  ensureValidJournalPath();
  renderJournalNav();
  render();
}
function renderSettings(){
  const root = buildSettingsUI({tab:state.settingsTab,sheets:state.sheets,settings:state.sheetSettings,addFieldsCfg:state.addFieldsCfg,uiPrefs:state.uiPrefs,onDirty:markDirty});
  const box=$("#settingsContent"); box.innerHTML=""; box.appendChild(root);
}


function fillSpaceSelect(spaces){
  if(!spaceSelect) return;
  const list = (spaces||[]).filter(s=>s && s.kind==="space" && !s.parentId);
  if(!list.length) return;
  const current = list.find(s=>s.id===state.spaceId) || list[0];
  spaceSelect.innerHTML="";
  for(const s of list){
    spaceSelect.appendChild(el("option",{value:s.id, textContent:`📁 ${s.name}`}));
  }
  spaceSelect.value = current.id;
  spaceSelect.disabled = false;
}

function ensureValidJournalPath(){
  const topIds = state.jtree?.topIds || [];
  if(!state.journalPath.length || !state.jtree?.nodes?.[state.journalPath[0]]){
    if(topIds[0]) state.journalPath = [topIds[0]];
  }
  // Trim to existing nodes
  state.journalPath = state.journalPath.filter(id=>!!state.jtree?.nodes?.[id]);
  if(!state.journalPath.length && topIds[0]) state.journalPath = [topIds[0]];
}

function renderJournalNav(){
  // Top-level journal select
  if(!sheetSelect) return;
  ensureValidJournalPath();

  const topNodes = (state.jtree?.topIds||[]).map(id=>state.jtree.nodes[id]).filter(Boolean);
  sheetSelect.innerHTML="";
  for(const n of topNodes){
    sheetSelect.appendChild(el("option",{value:n.id, textContent:`📄 ${nodeTitle(n)}`}));
  }
  sheetSelect.value = state.journalPath[0] || topNodes[0]?.id;

  // Active highlight: only the currently displayed node (deepest)
  sheetSelect.classList.toggle("nav-active", state.journalPath.length===1);

  sheetSelect.onchange = async ()=>{
    state.journalPath = [sheetSelect.value];
    await saveView();
    renderJournalNav();
    render();
  };

  // Add subjournal under top journal
  btnAddChild0.style.display = "inline-block";
  btnAddChild0.onclick = ()=> openAddChildModal(state.journalPath.slice(0,1), 0);

  // Render nested chain (select + plus) for children of the currently selected path
  subjournalChain.innerHTML="";
  for(let depth=0; depth<state.journalPath.length; depth++){
    const parentId = state.journalPath[depth];
    const kids = childrenOf(parentId);
    if(!kids.length) break;

    const nextDepth = depth+1;

    // Child selector with a "stay here" option.
    // This prevents auto-jumping into the first grandchild when a node has children.
    const sel = el("select",{className:"select", style:"max-width:220px", ariaLabel:"Піджурнал"});
    sel.appendChild(el("option",{value:"", textContent:"↩ У цьому журналі"}));
    for(const k of kids){
      sel.appendChild(el("option",{value:k.id, textContent:`📄 ${nodeTitle(k)}`}));
    }

    const existingSelected = state.journalPath[nextDepth];
    // If the user hasn't explicitly navigated deeper, keep selection empty (stay on current node).
    sel.value = (existingSelected && kids.find(k=>k.id===existingSelected)) ? existingSelected : "";

    // Keep path consistent with selection.
    if(sel.value===""){
      if(state.journalPath.length>nextDepth) state.journalPath = state.journalPath.slice(0, nextDepth);
    }else if(state.journalPath[nextDepth] !== sel.value){
      state.journalPath = state.journalPath.slice(0, nextDepth).concat([sel.value]);
    }

    sel.classList.toggle("nav-active", (state.journalPath.length===nextDepth+1));
    sel.onchange = async ()=>{
      if(sel.value===""){
        state.journalPath = state.journalPath.slice(0, nextDepth);
      }else{
        state.journalPath = state.journalPath.slice(0, nextDepth).concat([sel.value]);
      }
      await saveView();
      renderJournalNav();
      render();
    };

    const plus = el("button",{className:"btn btn-ghost", textContent:"＋", title:"Додати піджурнал"});
    // Add a child under the currently selected node at this level.
    // If no child is selected ("stay here"), add under the current parent node.
    plus.onclick = ()=>{
      const basePath = state.journalPath.slice(0, nextDepth); // includes parentId
      const targetPath = (sel.value && sel.value!=="") ? basePath.concat([sel.value]) : basePath;
      const targetDepth = targetPath.length-1;
      openAddChildModal(targetPath, targetDepth);
    };

    subjournalChain.appendChild(sel);
    subjournalChain.appendChild(plus);
  }
}
async function fillCaseSelect(){
  // Legacy feature (Номенклатура → Опис справ) is disabled.
  // Subjournals replace this functionality for ALL journals.
  caseSelect.style.display="none";
  btnCaseBack.style.display="none";
  return;
}

function matchesSearch(row, sheet){
  // Search must work regardless of hidden/visible columns.
  const q=(state.search||"").trim().toLowerCase();
  if(!q) return true;
  const parts=[];
  for(const colDef of (sheet.columns||[])){
    const c = colDef.name;
    if(!c) continue;
    // main cell value
    parts.push(String(row.data?.[c] ?? ""));
    // subrows values (if any)
    if(colDef.subrows){
      for(const sr of (row.subrows||[])) parts.push(String(sr?.[c] ?? ""));
    }
  }
  return parts.join(" ").toLowerCase().includes(q);
}

function matchesSearchCase(row){
  const q=(state.search||"").trim().toLowerCase();
  if(!q) return true;
  const parts=[];
  for(const col of (CASE_DESC_COLUMNS||[])){
    parts.push(String(row?.[col.name] ?? ""));
  }
  return parts.join(" ").toLowerCase().includes(q);
}
function sortRows(rows, sheet){
  const {col,dir}=state.sort; if(!col) return rows;
  const def=sheet.columns.find(c=>c.name===col);
  const getVal=(r)=>{
    if(def?.subrows) return String((r.subrows?.[0]?.[col])??"").toLowerCase();
    const v=r.data?.[col]??"";
    if(def?.type==="int" && isIntegerString(v)) return parseInt(v,10);
    if(def?.type==="date"){
      const p=parseUAdate(v); if(!p) return 0;
      const m=/^(\d{2})\.(\d{2})\.(\d{4})$/.exec(p);
      if(m) return new Date(+m[3],+m[2]-1,+m[1]).getTime();
      return 0;
    }
    return String(v).toLowerCase();
  };
  return rows.slice().sort((a,b)=>{ const va=getVal(a), vb=getVal(b); if(va<vb) return -1*dir; if(va>vb) return 1*dir; return 0; });
}
function toggleSort(sheet,col){ if(state.sort.col===col) state.sort.dir*=-1; else {state.sort.col=col; state.sort.dir=1;} render(); }
function nextOrder(rows,col){ let max=0; for(const r of rows){ const v=parseInt(r.data?.[col]??0,10); if(!Number.isNaN(v)&&v>max) max=v;} return max+1; }

function updateStickyOffsets(){
  try{
    const topbar=document.querySelector(".topbar");
    if(topbar){
      const topH=Math.round(topbar.getBoundingClientRect().height);
      document.documentElement.style.setProperty("--topbar-h", topH+"px");
    }
    const thead=table?.querySelector?.("thead");
    const tr1=thead?.querySelector?.("tr");
    if(tr1){
      const h1=Math.round(tr1.getBoundingClientRect().height);
      document.documentElement.style.setProperty("--head1-h", h1+"px");
    }
  }catch(e){}
}

async function render(){
  // Do not reload config on every render (it would override current UI state).
  // Config is loaded at startup, after settings save, and for explicit refresh.
  const spaces = await ensureSpaces();
  fillSpaceSelect(spaces);
  if(!state.jtree) state.jtree = await ensureJournalTree(state.spaceId, state.sheets);
  ensureValidJournalPath();
  renderJournalNav();
  if(levelSelect) levelSelect.value = "admin";
  applyDefaultSortForSheet(currentSheet());
  await fillCaseSelect();
  if(state.mode==="case" && state.caseId) return renderCase(state.caseId);
  return renderSheet();
}
async function renderSheet(){
  const sheet=currentSheet(); if(!sheet) return;
  ensureSimplifiedConfig(sheet);
  updateSimplifiedToggle();
  ensureSimplifiedConfig(sheet);
  updateSimplifiedToggle();
  setSideHint(sheet.key.startsWith("custom_")?"Користувацький лист":"");
  const colsVisible=visibleColumns(sheet);
  let rows=await getRows(currentDataKey());
  rows=rows.filter(r=>matchesSearch(r,sheet));
  if(sheet.orderColumn && !state.sort.col){
    rows=rows.slice().sort((a,b)=>parseInt(a.data?.[sheet.orderColumn]??0,10)-parseInt(b.data?.[sheet.orderColumn]??0,10));
  } else rows=sortRows(rows,sheet);


// Simplified view (cards) — Stage 4
if(sheet.simplified?.enabled && sheet.simplified?.on && sheet.simplified?.templates?.length){
  const ok = renderSimplifiedCardsForSheet(sheet, rows);
  if(ok){
    updateStickyOffsets();
    return;
  }
}
// fallback: normal table
if(cards){ cards.style.display="none"; cards.innerHTML=""; }
if(table){ table.style.display=""; }
table.innerHTML="";

  // col widths (persisted per sheet)
  const colgroup = el("colgroup");
  const ss = state.sheetSettings || (state.sheetSettings = {});
  const sheetSS = ss[sheet.key] || (ss[sheet.key] = {});
  const colWidths = sheetSS.colWidths || (sheetSS.colWidths = {});
  const colEls = [];
  colsVisible.forEach((name)=>{
    const col = el("col");
    const w = colWidths[name];
    if(typeof w === "number" && w > 0) col.style.width = w + "px";
    colEls.push(col);
    colgroup.appendChild(col);
  });
  // action cols (fixed)
  colgroup.appendChild(el("col",{className:"col-action"}));
  colgroup.appendChild(el("col",{className:"col-action"}));
  table.appendChild(colgroup);

  const thead=el("thead");
  const tr1=el("tr"), tr2=el("tr");
  colsVisible.forEach((name,i)=>{
    const th=el("th",{textContent:name});
    th.classList.add("sortable","col-resizable");
    th.onclick=()=>toggleSort(sheet,name);
    // column resize (drag near right edge) — pointer-events based (Edge-safe)
    let resizing=false;
    const RESIZE_ZONE=12; // px from right edge
    const minW = 60;

    const isNearRightEdge = (ev)=>{
      const x = ev.clientX;
      const r = th.getBoundingClientRect();
      return x >= (r.right - RESIZE_ZONE);
    };

    const setHover = (on)=>{
      if(on) th.classList.add("col-resize-hover");
      else th.classList.remove("col-resize-hover");
    };

    const startResize = (ev)=>{
      ev.preventDefault(); ev.stopPropagation();
      resizing=true;
      document.body.classList.add("col-resize-active");
      setHover(true);

      const startX = ev.clientX;
      const startW = th.getBoundingClientRect().width;

      // Capture pointer so Edge keeps sending move events even if cursor leaves header
      if(ev.pointerId != null && th.setPointerCapture){
        try{ th.setPointerCapture(ev.pointerId); }catch(e){}
      }

      const onMove = (mv)=>{
        const dx = mv.clientX - startX;
        const w = Math.max(minW, Math.round(startW + dx));
        colEls[i].style.width = w + "px";
        colWidths[name] = w;
      };

      const finish = async ()=>{
        document.removeEventListener("pointermove", onMove);
        document.removeEventListener("pointerup", onUp);
        document.removeEventListener("pointercancel", onUp);
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
        document.body.classList.remove("col-resize-active");
        resizing=false;
        setHover(false);
        try{ await saveSheetSettings(ss); }catch(e){}
      };

      const onUp = ()=>{ finish(); };

      // Prefer pointer events (Edge), fall back to mouse
      document.addEventListener("pointermove", onMove);
      document.addEventListener("pointerup", onUp, {once:true});
      document.addEventListener("pointercancel", onUp, {once:true});
      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp, {once:true});
    };

    // hover detection (pointer + mouse)
    const onHoverMove = (ev)=>{ if(resizing) return; setHover(isNearRightEdge(ev)); };
    th.addEventListener("pointermove", onHoverMove);
    th.addEventListener("mousemove", onHoverMove);
    th.addEventListener("mouseleave",()=>{ if(!resizing) setHover(false); });
    th.addEventListener("pointerleave",()=>{ if(!resizing) setHover(false); });

    // start resize (pointer + mouse)
    th.addEventListener("pointerdown",(ev)=>{ if(isNearRightEdge(ev)) startResize(ev); });
    th.addEventListener("mousedown",(ev)=>{ if(isNearRightEdge(ev)) startResize(ev); });

    // prevent sort click after resizing / near-edge interaction
    th.addEventListener("click",(ev)=>{
      if(resizing || th.classList.contains("col-resize-hover")){ ev.stopPropagation(); }
    }, true);
    if(state.sort.col===name) th.appendChild(el("span",{className:"sort-ind",textContent: state.sort.dir===1?"▲":"▼"}));
    tr1.appendChild(th);
    tr2.appendChild(el("th",{textContent:String(i+1)}));
  });
  tr1.appendChild(el("th",{className:"th-action",textContent:"↪"}));
  tr1.appendChild(el("th",{className:"th-action",textContent:"🗑"}));
  tr2.appendChild(el("th",{className:"th-action",textContent:String(colsVisible.length+1)}));
  tr2.appendChild(el("th",{className:"th-action",textContent:String(colsVisible.length+2)}));
  thead.appendChild(tr1); thead.appendChild(tr2); table.appendChild(thead);
  // after header exists, measure exact offsets for sticky stacking
  updateStickyOffsets();
  const tbody=el("tbody"); table.appendChild(tbody);

  const hasSubCols = sheet.columns.some(c=>c.subrows);
  const actionDeleteCell=(row)=>{
    const td=el("td",{className:"td-action"});
    const b=el("button",{className:"icon danger",textContent:"🗑"});
    b.onclick=async (ev)=>{ev.stopPropagation(); await deleteFlow(sheet,row);};
    td.appendChild(b); return td;
  };
  const actionTransferCell=(row)=>{
    const td=el("td",{className:"td-action"});
    const b=el("button",{className:"icon",textContent:"↪"});
    b.onclick=async (ev)=>{ev.stopPropagation(); await transferFlow(sheet,row);};
    td.appendChild(b); return td;
  };

  // Render: main row + 0..N subrows. Columns with subrows=false are shared (rowSpan).
  if(!hasSubCols){
    for(const r of rows){
      const tr=el("tr");
      if(state.selectionMode){
        tr.classList.add("sel-row");
        if(state.selectedRowIds.has(r.id)) tr.classList.add("selected");
        tr.onclick=(ev)=>{ ev.preventDefault(); ev.stopPropagation(); toggleRowSelection(r.id); };
      }
      for(const cn of colsVisible){
        const td=el("td",{textContent:String(r.data?.[cn]??"")});
        td.onclick=(ev)=>{ if(state.selectionMode){ ev.stopPropagation(); toggleRowSelection(r.id); return; } editCell(sheet,r,cn,null); };
        tr.appendChild(td);
      }
      tr.appendChild(actionTransferCell(r));
      tr.appendChild(actionDeleteCell(r));
      tbody.appendChild(tr);
    }
    return;
  }
  for(const r of rows){
    const subs = r.subrows || [];
    const totalRows = 1 + subs.length; // 1 main + subrows
    const firstSubCol = colsVisible.find(cn=>{
      const d=sheet.columns.find(x=>x.name===cn);
      return !!d?.subrows;
    });
    for(let i=0;i<totalRows;i++){
      const tr=el("tr");
      if(state.selectionMode){
        tr.classList.add("sel-row");
        if(state.selectedRowIds.has(r.id)) tr.classList.add("selected");
        tr.onclick=(ev)=>{ ev.preventDefault(); ev.stopPropagation(); toggleRowSelection(r.id); };
      }
      for(const colName of colsVisible){
        const def=sheet.columns.find(x=>x.name===colName);
        const allowSub = !!def?.subrows;
        if(!allowSub){
          if(i!==0) continue;
          const td=el("td",{textContent:String(r.data?.[colName]??"")});
          td.rowSpan = totalRows;
          td.onclick=(ev)=>{ if(state.selectionMode){ ev.stopPropagation(); toggleRowSelection(r.id); return; } editCell(sheet,r,colName,null); };
          tr.appendChild(td);
          continue;
        }

        // allowSub === true
        let txt="";
        let subIndex=null;
        const hasSubs = (subs.length>0);
        if(i===0){
          // Main row acts as Subrow #1 (no separate "main" vs numbered subrows)
          txt = (hasSubs && colName==="Номер примірника") ? "1" : String(r.data?.[colName] ?? "");
          subIndex = null;
        } else {
          const sr = subs[i-1] || {};
          txt = (colName==="Номер примірника") ? String(i+1) : String(sr[colName] ?? "");
          subIndex = i-1;
        }
        const td=el("td",{});
        // Show subrow ordinal for ALL subrows (including the first one), but only when row actually has subrows.
        if(hasSubs && firstSubCol && colName===firstSubCol){
          td.appendChild(el("span",{className:"subrow-idx",textContent:String(i+1)}));
          td.appendChild(el("span",{textContent:" "}));
        }
        td.appendChild(el("span",{textContent:txt}));
        td.onclick=(ev)=>{ if(state.selectionMode){ ev.stopPropagation(); toggleRowSelection(r.id); return; } editCell(sheet,r,colName,subIndex); };

        tr.appendChild(td);
      }
      if(i===0){
        const tdT=actionTransferCell(r); tdT.rowSpan=totalRows; tr.appendChild(tdT);
        const tdD=actionDeleteCell(r); tdD.rowSpan=totalRows; tr.appendChild(tdD);
      }
      tbody.appendChild(tr);
    }
  }
}
async function renderCase(caseId){
  updateSimplifiedToggle();
  updateSimplifiedToggle();
  setSideHint("Внутрішній опис справи");
  table.innerHTML="";
  let rows=await getCaseRows(caseId);
  // Apply search filter for case descriptions as well
  rows = rows.filter(matchesSearchCase);
  rows.sort((a,b)=>parseInt(a["№ з/п"]??0,10)-parseInt(b["№ з/п"]??0,10));
  const thead=el("thead"); const tr1=el("tr"), tr2=el("tr");
  CASE_DESC_COLUMNS.forEach((col,i)=>{ tr1.appendChild(el("th",{textContent:col.name})); tr2.appendChild(el("th",{textContent:String(i+1)})); });
  tr1.appendChild(el("th",{className:"th-action",textContent:"↪"}));
  tr2.appendChild(el("th",{className:"th-action",textContent:String(CASE_DESC_COLUMNS.length+1)}));
  tr1.appendChild(el("th",{className:"th-action",textContent:"🗑"}));
  tr2.appendChild(el("th",{className:"th-action",textContent:String(CASE_DESC_COLUMNS.length+2)}));
  thead.appendChild(tr1); thead.appendChild(tr2); table.appendChild(thead);
  const tbody=el("tbody"); table.appendChild(tbody);
  for(const r of rows){
    const tr=el("tr");
    if(state.selectionMode && state.selectedRowIds?.has(r.id)) tr.classList.add("row-selected");
    CASE_DESC_COLUMNS.forEach(col=>{
      const td=el("td",{textContent:String(r[col.name]??"")});
      td.onclick=async (ev)=>{
        if(state.selectionMode){ ev.stopPropagation(); toggleRowSelection(r.id); return; }
        if(col.editable===false) return;
        const v=prompt(`Редагування\n${col.name}:`,String(r[col.name]??""));
        if(v===null) return;
        r[col.name]=String(v);
        await putCaseRow(r);
        render();
      };
      tr.appendChild(td);
    });
    // transfer button
    const tdT=el("td",{className:"td-action"});
    const bT=el("button",{className:"icon",textContent:"↪"});
    bT.onclick=async (ev)=>{ ev.stopPropagation(); await transferCaseFlow(caseId, r); };
    tdT.appendChild(bT); tr.appendChild(tdT);
    const tdD=el("td",{className:"td-action"});
    const b=el("button",{className:"icon danger",textContent:"🗑"});
    b.onclick=async (ev)=>{ev.stopPropagation(); const ok=await confirmDeleteNumber("Видалити рядок?"); if(!ok) return; await deleteCaseRow(r.id); render();};
    tdD.appendChild(b); tr.appendChild(tdD);
    tbody.appendChild(tr);
  }
}
async function validateValue(def, raw){
  const s=String(raw??"").trim();
  if(def?.required && !s){ alert(`Поле «${def.name}» є обовʼязковим.`); return null; }
  if(!s) return "";
  if(def?.type==="int"){ if(!isIntegerString(s)){ alert("Потрібно число (лише цифри)."); return null; } return s; }
  if(def?.type==="date"){ const p=parseUAdate(s); if(!p){ alert("Дата: ДД.ММ.РР"); return null; } return p; }
  return s;
}
async function editCell(sheet,row,colName,subIndex){
  const def=sheet.columns.find(c=>c.name===colName);
  if(def?.editable===false) return;
  if(def?.subrows){
    const lineLabel = (subIndex===null)
      ? "Підстрочка № 1"
      : `Підстрочка № ${subIndex+2}`;
    const actions=[
      btn("Скасувати","cancel","btn"),
      btn("Редагувати","edit","btn btn-primary"),
      btn("Додати підстрочку","add","btn btn-primary"),
    ];
    if(subIndex!==null) actions.push(btn("Видалити підстрочку","del","btn"));
    const op = await modalOpen({
      title:"Оберіть дію",
      bodyNodes:[el("div",{className:"muted",textContent:`${sheet.title}\n${colName}\n${lineLabel}`})],
      actions
    });
    if(op.type==="cancel") return;
    if(op.type==="edit"){
      if(subIndex===null){
        const current=String(row.data?.[colName]??"");
        const v=prompt(`${sheet.title}\n\nРедагування: ${colName}`, current);
        if(v===null) return;
        const val=await validateValue(def,v); if(val===null) return;
        row.data=row.data||{}; row.data[colName]=val;
        await putRow(row); render();
        return;
      }
      return editSubCell(sheet,row,colName,subIndex);
    }
    if(op.type==="add") return addSubRow(sheet,row,subIndex);
    if(op.type==="del") return deleteSubRow(sheet,row,subIndex);
    return;
  }
  const current=String(row.data?.[colName]??"");
  const v=prompt(`${sheet.title}\n\nРедагування: ${colName}`, current);
  if(v===null) return;
  const val=await validateValue(def,v); if(val===null) return;
  row.data=row.data||{}; row.data[colName]=val;
  await putRow(row); render();
}
async function editSubCell(sheet,row,colName,subIndex){
  const subs=row.subrows||[]; const sr=subs[subIndex]||{};
  if(colName==="Номер примірника") return;
  const v=prompt(`${sheet.title}\n\nРедагування підрядка #${subIndex+1}\n${colName}:`, String(sr[colName]??"")); if(v===null) return;
  const def=sheet.columns.find(c=>c.name===colName);
  const val=await validateValue(def,v); if(val===null) return;
  sr[colName]=val; subs[subIndex]=sr; row.subrows=subs;
  if(sheet.columns.some(c=>c.name==="Кількість примірників")){ row.data=row.data||{}; row.data["Кількість примірників"]=String(subs.length); }
  await putRow(row); render();
}
async function addSubRow(sheet,row,afterIndex){
  const subs=row.subrows||[];
  const insertAt=(afterIndex===null||afterIndex===undefined)?subs.length:Math.min(afterIndex+1,subs.length);
  const newSr={};
  const subCols=sheet.columns.filter(c=>c.subrows && !c.computed && c.editable!==false && c.name!=="Номер примірника");
  if(!subCols.length){
    alert("У цьому листі підстрочки вимкнені для всіх колонок (або всі колонки службові).\n\nУвімкніть підстрочки в конструкторі колонок.");
    return;
  }
  for(const sc of subCols){
    const v=prompt(`${sheet.title}\n\nНовий підрядок\n${sc.name}:`,""); if(v===null) return;
    const vv=await validateValue(sc,v); if(vv===null) return;
    newSr[sc.name]=vv;
  }
  subs.splice(insertAt,0,newSr); row.subrows=subs;
  if(sheet.columns.some(c=>c.name==="Кількість примірників")){ row.data=row.data||{}; row.data["Кількість примірників"]=String(subs.length); }
  await putRow(row); render();
}
async function deleteSubRow(sheet,row,subIndex){
  const subs=row.subrows||[];
  if(subIndex===null||subIndex===undefined) return;
  const ok=await confirmDeleteNumber(`${sheet.title}\nВидалити підрядок № ${subIndex+1}?`); if(!ok) return;
  subs.splice(subIndex,1); row.subrows=subs;
  if(sheet.columns.some(c=>c.name==="Кількість примірників")){ row.data=row.data||{}; row.data["Кількість примірників"]=String(subs.length); }
  await putRow(row); render();
}
async function addFlow(){
  if(state.mode==="case"){ alert("Додавання у опис справи — через перенесення ↪."); return; }
  const sheet=currentSheet(); if(!sheet) return;
  const rows=await getRows(currentDataKey());
  const record={data:{}, subrows:[]};
  if(sheet.orderColumn) record.data[sheet.orderColumn]=String(nextOrder(rows,sheet.orderColumn));
  for(const c of sheet.columns){
    if(c.defaultValue && record.data[c.name]==null) record.data[c.name]=c.defaultValue;
    if(c.type==="date" && c.defaultToday && record.data[c.name]==null) record.data[c.name]=uaDateToday();
  }
  const addCfg = state.addFieldsCfg[sheet.key] || sheet.addFields || sheet.columns.map(c=>c.name);
  for(const name of addCfg){
    const def=sheet.columns.find(c=>c.name===name); if(!def) continue;
    if(sheet.orderColumn && name===sheet.orderColumn) continue;
    if(def.computed) continue;
    // Main row value (even if column allows subrows)
    const v=prompt(`${sheet.title}\n\nВведіть: ${name}${def.required?" (обовʼязково)":""}`, String(record.data[name]??""));
    if(v===null) return;
    const vv=await validateValue(def,v); if(vv===null) return;
    record.data[name]=vv;
  }

  // Optional subrows (per-row)
  const subCols=sheet.columns.filter(c=>c.subrows && !c.computed && c.editable!==false && c.name!=="Номер примірника");
  if(subCols.length){
    const want = confirm("Додати підстрочки для цієї строки?\n\nOK = Так\nCancel = Ні");
    if(want){
      while(true){
        const sr={};
        for(const sc of subCols){
          const v=prompt(`${sheet.title}\n\nПідстрочка ${record.subrows.length+1}\n${sc.name}:`,"");
          if(v===null) return;
          const vv=await validateValue(sc,v); if(vv===null) return;
          sr[sc.name]=vv;
        }
        record.subrows.push(sr);
        const more=confirm("Додати ще одну підстрочку?\n\nOK = Так\nCancel = Ні");
        if(!more) break;
      }
      if(sheet.columns.some(c=>c.name==="Кількість примірників")) record.data["Кількість примірників"]=String(record.subrows.length);
    }
  }
  await addRow(currentDataKey(), record);
  if(sheet.key==="nomenklatura") await ensureCaseFromNomenRecord(record);
  await fillCaseSelect();
  render();
}
async function ensureCaseFromNomenRecord(row){
  const idx=String(row.data?.["Індекс справи"]??"").trim();
  const title=String(row.data?.["Заголовок справи (тому, частини)"]??"").trim();
  if(!idx && !title) return null;
  const cases=await getAllCases();
  const existing=cases.find(c=>String(c.caseIndex||"").trim()===idx && String(c.caseTitle||"").trim()===title);
  if(existing) return existing;
  const c={caseIndex:idx, caseTitle:title, createdAt:new Date().toISOString(), createdFrom:"nomenklatura"};
  const id=await addCase(c); c.id=id; return c;
}
async function deleteFlow(sheet,row){
  const hasSub = sheet.columns.some(c=>c.subrows);
  if(!hasSub || !(row.subrows && row.subrows.length)){
    const ok=await confirmDeleteNumber(`${sheet.title}\nВидалити всю строку?`); if(!ok) return;
    await deleteRow(row.id); render(); return;
  }
  const op = await modalOpen({
    title:"Видалення",
    bodyNodes:[el("div",{className:"muted",textContent:`${sheet.title}\nОберіть що видалити`})],
    actions:[btn("Скасувати","cancel","btn"),btn("Вся строка","row","btn btn-primary"),btn("Підстрока","sub","btn btn-primary")]
  });
  if(op.type==="cancel") return;
  if(op.type==="row"){ const ok=await confirmDeleteNumber(`${sheet.title}\nВидалити всю строку?`); if(!ok) return; await deleteRow(row.id); render(); return; }
  if(op.type==="sub"){
    const v=prompt(`Виберіть номер підстроки (1..${row.subrows.length})`,"1"); if(v===null) return;
    const n=parseInt(v,10); if(Number.isNaN(n)||n<1||n>row.subrows.length) return alert("Некоректний номер.");
    const ok=await confirmDeleteNumber(`${sheet.title}\nВидалити підстроку № ${n}?`); if(!ok) return;
    row.subrows.splice(n-1,1);
    if(sheet.columns.some(c=>c.name==="Кількість примірників")){ row.data=row.data||{}; row.data["Кількість примірників"]=String(row.subrows.length); }
    await putRow(row); render();
  }
}
async function transferFlow(sheet,row){
  const tpls=await getTransferTemplates();
  const forSheet=tpls.filter(t=>t.fromSheetKey===sheet.key);
  if(!forSheet.length){ alert("Немає шаблонів перенесення для цього листа."); return; }

  const selTpl = el("select",{className:"select"});
  forSheet.forEach((t,i)=>selTpl.appendChild(el("option",{value:t.id,textContent:`${i+1}) ${t.name||"(без назви)"}`})));
  selTpl.value = forSheet[0].id;

  const info = el("div",{className:"muted",style:"margin-top:6px"});

  const modeWrap = el("div",{style:"margin-top:10px"});
  const rbAll = el("input",{type:"radio", name:"submode", checked:true});
  const rbPick = el("input",{type:"radio", name:"submode"});
  const lblAll = el("label",{style:"display:flex; gap:8px; align-items:center;"});
  lblAll.appendChild(rbAll); lblAll.appendChild(el("span",{textContent:"Переносити всі підстрочки (1,2,3... по індексу)"}));
  const lblPick = el("label",{style:"display:flex; gap:8px; align-items:center; margin-top:6px;"});
  lblPick.appendChild(rbPick); lblPick.appendChild(el("span",{textContent:"Обрати підстрочки"}));
  modeWrap.appendChild(lblAll);
  modeWrap.appendChild(lblPick);

  const subsBox = el("div",{style:"margin-top:8px; padding-left:22px"});
  modeWrap.appendChild(subsBox);

  const render = ()=>{
    const t = forSheet.find(x=>x.id===selTpl.value) || forSheet[0];
    if(t.toSheetKey==="__case__"){
      info.textContent = "Ціль: Опис справи";
    } else {
      const dest = state.sheets.find(s=>s.key===t.toSheetKey);
      info.textContent = dest ? `Ціль: Лист: ${dest.title}` : "";
    }
    subsBox.innerHTML="";
    const subs=row.subrows||[];
    const total = 1 + subs.length;
    if(total===1){
      // only one subrow (№1)
      rbPick.disabled=true;
      rbAll.checked=true;
      subsBox.appendChild(el("div",{className:"muted",textContent:"У цієї строки є тільки підстрочка №1."}));
      return;
    }
    rbPick.disabled=false;
    subsBox.appendChild(el("div",{className:"muted",textContent:"Оберіть підстрочки для перенесення (можна декілька):"}));
    const tools=el("div",{className:"row",style:"gap:8px; margin-top:6px"});
    const bAll=el("button",{className:"btn",textContent:"Всі"});
    const bNone=el("button",{className:"btn",textContent:"Жодної"});
    tools.appendChild(bAll); tools.appendChild(bNone);
    subsBox.appendChild(tools);
    const list=el("div",{style:"margin-top:6px"});
    // Показуємо ВСІ підстрочки 1..N, де 1 — це row.data, а 2..N — row.subrows
    for(let i=0;i<total;i++){
      const lab=el("label",{style:"display:flex; gap:8px; align-items:center; margin:2px 0"});
      const ch=el("input",{type:"checkbox"});
      // subIndex тут у "загальній" шкалі: 0 => перша підстрочка (row.data), 1.. => row.subrows[subIndex-1]
      ch.dataset.subIndex=String(i);
      lab.appendChild(ch);
      lab.appendChild(el("span",{textContent:`Підстрочка ${i+1}`}));
      list.appendChild(lab);
    }
    bAll.onclick=()=>{ list.querySelectorAll("input[type=checkbox]").forEach(x=>x.checked=true); };
    bNone.onclick=()=>{ list.querySelectorAll("input[type=checkbox]").forEach(x=>x.checked=false); };
    subsBox.appendChild(list);
  };
  selTpl.onchange=render;
  render();

  const op = await modalOpen({
    title:"Перенесення",
    bodyNodes:[
      el("div",{className:"muted",textContent:sheet.title}),
      selTpl,
      info,
      modeWrap
    ],
    actions:[btn("Скасувати","cancel","btn"),btn("Перенести","go","btn btn-primary")]
  });
  if(op.type!=="go") return;

  const tpl = forSheet.find(x=>x.id===selTpl.value) || forSheet[0];
  let subMode="all";
  let selectedSubIdx=[];
  if(rbPick.checked){
    subMode="selected";
    selectedSubIdx = Array.from(subsBox.querySelectorAll("input[type=checkbox]")).filter(ch=>ch.checked).map(ch=>parseInt(ch.dataset.subIndex,10)).filter(n=>Number.isFinite(n));
    if(!selectedSubIdx.length) return alert("Оберіть хоча б одну підстрочку.");
  }
  await runTransferForRows(sheet,[row],tpl,{subMode, subIndexes:selectedSubIdx});
}

function casePseudoSheet(){
  return { key:"__case__", title:"Опис справи", columns: CASE_DESC_COLUMNS.map(c=>({name:c.name, subrows:false})) };
}
function wrapCaseRow(r){
  // normalize case row to the same shape as journal rows
  return { id:r.id, data:{...r}, subrows:[] };
}

async function transferCaseFlow(caseId, caseRow){
  const tpls=await getTransferTemplates();
  const forCase=tpls.filter(t=>t.fromSheetKey==="__case__");
  if(!forCase.length){ alert("Немає шаблонів перенесення для опису справи."); return; }

  const selTpl = el("select",{className:"select"});
  forCase.forEach((t,i)=>selTpl.appendChild(el("option",{value:t.id,textContent:`${i+1}) ${t.name||"(без назви)"}`})));
  selTpl.value = forCase[0].id;

  const info = el("div",{className:"muted",style:"margin-top:6px"});
  const renderInfo=()=>{
    const t = forCase.find(x=>x.id===selTpl.value) || forCase[0];
    if(t.toSheetKey==="__case__") info.textContent = "Ціль: Опис справи";
    else {
      const dest = state.sheets.find(s=>s.key===t.toSheetKey);
      info.textContent = dest ? `Ціль: Лист: ${dest.title}` : "";
    }
  };
  selTpl.onchange=renderInfo; renderInfo();

  const op = await modalOpen({
    title:"Перенесення",
    bodyNodes:[
      el("div",{className:"muted",textContent:"Опис справи"}),
      selTpl,
      info,
      el("div",{className:"muted",textContent:"У описі справи підстрочок немає — переноситься один рядок."})
    ],
    actions:[btn("Скасувати","cancel","btn"),btn("Перенести","go","btn btn-primary")]
  });
  if(op.type!=="go") return;
  const tpl = forCase.find(x=>x.id===selTpl.value) || forCase[0];
  await runTransferForRows(casePseudoSheet(), [wrapCaseRow(caseRow)], tpl, {subMode:"main", subIndexes:[]}, {caseId});
}

async function transferMultipleCaseFlow(caseId, caseRows){
  const tpls=await getTransferTemplates();
  const forCase=tpls.filter(t=>t.fromSheetKey==="__case__");
  if(!forCase.length){ alert("Немає шаблонів перенесення для опису справи."); return; }

  const selTpl = el("select",{className:"select"});
  forCase.forEach((t,i)=>selTpl.appendChild(el("option",{value:t.id,textContent:`${i+1}) ${t.name||"(без назви)"}`})));
  selTpl.value=forCase[0].id;
  const info=el("div",{className:"muted",style:"margin-top:6px"});
  const renderInfo=()=>{
    const t = forCase.find(x=>x.id===selTpl.value) || forCase[0];
    if(t.toSheetKey==="__case__") info.textContent = "Ціль: Опис справи";
    else {
      const dest = state.sheets.find(s=>s.key===t.toSheetKey);
      info.textContent = dest ? `Ціль: Лист: ${dest.title}` : "";
    }
  };
  selTpl.onchange=renderInfo; renderInfo();

  const op = await modalOpen({
    title:"Перенесення вибраних",
    bodyNodes:[
      el("div",{className:"muted",textContent:`Опис справи — вибрано рядків: ${caseRows.length}`}),
      selTpl,
      info
    ],
    actions:[btn("Скасувати","cancel","btn"),btn("Перенести","go","btn btn-primary")]
  });
  if(op.type!=="go") return;
  const tpl = forCase.find(x=>x.id===selTpl.value) || forCase[0];
  const wrapped = caseRows.map(wrapCaseRow);
  await runTransferForRows(casePseudoSheet(), wrapped, tpl, {subMode:"main", subIndexes:[]}, {caseId});
}


async function transferMultipleFlow(sheet, rows){
  const tpls=await getTransferTemplates();
  const forSheet=tpls.filter(t=>t.fromSheetKey===sheet.key);
  if(!forSheet.length){ alert("Немає шаблонів перенесення для цього листа."); return; }

  const selTpl = el("select",{className:"select"});
  forSheet.forEach((t,i)=>selTpl.appendChild(el("option",{value:t.id,textContent:`${i+1}) ${t.name||"(без назви)"}`})));
  selTpl.value=forSheet[0].id;

  const info=el("div",{className:"muted",style:"margin-top:6px"});
  const modeSel=el("select",{className:"select"});
  [{v:"main",t:"Тільки основні строки"},{v:"all",t:"Основні + всі підстрочки"}].forEach(x=>modeSel.appendChild(el("option",{value:x.v,textContent:x.t})));

  const render=()=>{
    const t=forSheet.find(x=>x.id===selTpl.value) || forSheet[0];
    if(t.toSheetKey==="__case__"){
      info.textContent = "Ціль: Опис справи";
    } else {
      const dest=state.sheets.find(s=>s.key===t.toSheetKey);
      info.textContent = dest ? `Ціль: Лист: ${dest.title}` : "";
    }
  };
  selTpl.onchange=render; render();

  const op = await modalOpen({
    title:"Перенесення вибраних",
    bodyNodes:[
      el("div",{className:"muted",textContent:`${sheet.title} — вибрано строк: ${rows.length}`}),
      selTpl,
      info,
      modeSel,
      el("div",{className:"muted",textContent:"Для вибраних строк детальний вибір підстрочок по кожній строкі не робимо — або тільки основні, або всі підстрочки."})
    ],
    actions:[btn("Скасувати","cancel","btn"),btn("Перенести","go","btn btn-primary")]
  });
  if(op.type!=="go") return;
  const tpl = forSheet.find(x=>x.id===selTpl.value) || forSheet[0];
  const subMode = modeSel.value || "main";
  await runTransferForRows(sheet, rows, tpl, {subMode, subIndexes:[]});
}


async function runTransferForRows(sourceSheet, rows, tpl, {subMode, subIndexes}, ctx={}){
  const toIsCase = (tpl.toSheetKey==="__case__");
  const toSheet = toIsCase ? casePseudoSheet() : (state.sheets.find(s=>s.key===tpl.toSheetKey) || null);
  if(!toSheet) return alert("Цільовий лист не знайдено.");
  let targetCaseId = null;
  if(toIsCase){
    // If we are already inside a case view and caller provided caseId, default to it.
    targetCaseId = ctx.caseId || (state.mode==="case" ? state.caseId : null);
    if(!targetCaseId){
      const id = await pickOrCreateCase();
      if(!id) return;
      targetCaseId = id;
    }
  }

  // Determine which routes require writing to subrows
  let routes = (tpl.routes||[]).map(r=>({...r, sources:[...(r.sources||[])]}));
  const needEnable = new Set();
  if(!toIsCase && subMode!=="main"){
    for(const r of routes){
      const colName = (toSheet.columns?.[r.targetCol]?.name);
      if(!colName) continue;
      const def = toSheet.columns.find(c=>c.name===colName);
      if(def && def.subrows===false) needEnable.add(colName);
    }
  }
  if(needEnable.size){
    const cols=[...needEnable].join(", ");
    const op=await modalOpen({
      title:"Підстрочки заборонені",
      bodyNodes:[
        el("div",{className:"muted",textContent:`Для перенесення потрібні підстрочки у колонках: ${cols}`}),
        el("div",{className:"muted",textContent:"Дозволити підстрочки у цих колонках і продовжити?"})
      ],
      actions:[btn("Скасувати","cancel","btn"),btn("Пропустити ці колонки","skip","btn"),btn("Дозволити і продовжити","allow","btn btn-primary")]
    });
    if(op.type==="cancel") return;
    if(op.type==="allow"){
      const all=await getAllSheets();
      const sh=all.find(s=>s.key===toSheet.key);
      if(sh){
        for(const cn of needEnable){
          const c=sh.columns.find(x=>x.name===cn);
          if(c) c.subrows=true;
        }
        await saveAllSheets(all);
        state.sheets=await getAllSheets();
      }
    }
    if(op.type==="skip"){
      routes = routes.filter(r=>{
        const colName = (toSheet.columns?.[r.targetCol]?.name);
        return colName ? !needEnable.has(colName) : true;
      });
    }
  }

  const getVal = (row, colIdx, subIdx)=>{
    const colName = sourceSheet.columns?.[colIdx]?.name;
    if(!colName) return "";
    // subIdx is 0-based: 0 => first subrow (main/data), 1.. => row.subrows[subIdx-1]
    if(subIdx!=null){
      if(subIdx===0){
        const v = row.data?.[colName];
        return String(v ?? "");
      }
      const sr = (row.subrows||[])[subIdx-1];
      if(sr && sr[colName]!=null) return String(sr[colName]);
      return "";
    }
    return String(row.data?.[colName] ?? "");
  };

  const compute = (row, subIdx, route)=>{
    const vals = (route.sources||[]).map(i=>getVal(row,i,subIdx));
    if(route.op==="sum"){
      let sum=0;
      for(const v of vals){
        const n = parseFloat(String(v).replace(",", "."));
        if(!Number.isNaN(n)) sum += n;
      }
      const s = (Math.round(sum)===sum) ? String(sum) : String(sum);
      return s;
    }
    if(route.op==="seq"){
      return vals.map(v=>String(v)).join("");
    }
    if(route.op==="newline"){
      const parts = vals.map(v=>String(v)).filter(v=>v.trim()!=="");
      return parts.join("\n");
    }
    // concat
    const delim = (route.delimiter==null) ? " " : String(route.delimiter);
    const parts = vals.map(v=>String(v)).filter(v=>v.trim()!=="");
    return parts.join(delim);
  };

  for(const row of rows){
    // When destination is Case Description, each selected subrow becomes its own case row (no subrows in case table).
    if(toIsCase){
      let idxes=[0];
      const total = 1 + (row.subrows||[]).length;
      if(subMode==="all") idxes = Array.from({length: total}, (_,i)=>i);
      else if(subMode==="selected"){
        idxes = (subIndexes||[]).filter(n=>Number.isFinite(n));
        if(!idxes.length) idxes=[0];
      } else idxes=[0];
      idxes = [...idxes].sort((a,b)=>a-b);

      for(const subIdx of idxes){
        const mapped = {};
        for(const r of routes){
          const tgtName = toSheet.columns?.[r.targetCol]?.name;
          if(!tgtName) continue;
          mapped[tgtName] = compute(row, subIdx, r);
        }
        await appendCaseRow(targetCaseId, mapped);
      }
      continue;
    }

    const out = { data:{}, subrows:[] };

    // Build list of source subrow indexes:
    // 0 => first subrow (row.data), 1.. => row.subrows[subIdx-1]
    let idxes=[0];
    const total = 1 + (row.subrows||[]).length;
    if(subMode==="all"){
      idxes = Array.from({length: total}, (_,i)=>i);
    } else if(subMode==="selected"){
      idxes = (subIndexes||[]).filter(n=>Number.isFinite(n));
      if(!idxes.length) idxes=[0];
    } else if(subMode==="main"){
      idxes=[0];
    }
    idxes = [...idxes].sort((a,b)=>a-b);

    const firstIdx = idxes[0] ?? 0;

    // Compute destination main/data from the first selected subrow
    for(const r of routes){
      const tgtName = toSheet.columns?.[r.targetCol]?.name;
      if(!tgtName) continue;
      out.data[tgtName] = compute(row, firstIdx, r);
    }

    // Additional selected subrows become destination subrows (by index order)
    const extraIdxes = idxes.slice(1);
    if(extraIdxes.length){
      for(let j=0;j<extraIdxes.length;j++) out.subrows.push({});
      for(const r of routes){
        const tgtName = toSheet.columns?.[r.targetCol]?.name;
        if(!tgtName) continue;
        const def = toSheet.columns.find(c=>c.name===tgtName);
        if(def && def.subrows){
          for(let j=0;j<extraIdxes.length;j++){
            out.subrows[j][tgtName] = compute(row, extraIdxes[j], r);
          }
        }
      }
    }
    // auto-order for destination
    if(toSheet.orderColumn){
      const cur=String(out.data[toSheet.orderColumn]??"").trim();
      if(!cur){
        const existing=await getRows(journalKeyForSheet(toSheet.key));
        const maxN = existing.reduce((m,r)=>Math.max(m, parseInt(r.data?.[toSheet.orderColumn]??0,10)||0),0);
        out.data[toSheet.orderColumn]=String(maxN+1);
      }
    }
    await addRow(journalKeyForSheet(toSheet.key), out);
  }
  alert("Перенесено.");
  render();
}

function computeMappedRow(target,row,subIndex){
  const evalExpr=(expr)=>{
    const field=(col,from)=>{ if(from==="sub"){ const sr=(row.subrows||[])[subIndex]||{}; return sr[col]??""; } return row.data?.[col]??""; };
    if(expr.op==="field") return String(field(expr.col, expr.from||"data"));
    if(expr.op==="concat"){ const j=expr.joiner??" "; return (expr.parts||[]).map(evalExpr).filter(v=>v!=="").join(j); }
    if(expr.op==="sum"){ let sum=0; for(const p of (expr.parts||[])){ const n=parseInt(String(evalExpr(p)).trim(),10); if(!Number.isNaN(n)) sum+=n; } return String(sum); }
    return "";
  };
  const mapped={};
  for(const m of target.map||[]) mapped[m.destCol]=evalExpr(m.expr);
  return mapped;
}
async function pickOrCreateCase(){
  const cases=await getAllCases();
  const list=cases.map((c,i)=>`${i+1}) ${c.caseIndex||"(без індексу)"} — ${c.caseTitle||"(без заголовка)"}`).join("\n");
  const v=prompt(`Виберіть справу (номер) або введіть індекс вручну.\n\nІснуючі:\n${list}`,"");
  if(v===null) return null;
  const t=v.trim(); if(!t) return null;
  const n=parseInt(t,10);
  if(!Number.isNaN(n)&&n>=1&&n<=cases.length) return cases[n-1].id;
  const index=t;
  const existing=cases.find(c=>String(c.caseIndex||"").trim()===index);
  if(existing) return existing.id;
  const title=prompt("Заголовок справи (необовʼязково):",""); if(title===null) return null;
  const id=await addCase({caseIndex:index, caseTitle:String(title||"").trim(), createdAt:new Date().toISOString(), createdFrom:"manual"});
  return id;
}
async function appendCaseRow(caseId,mapped){
  const rows=await getCaseRows(caseId);
  let max=0; for(const r of rows){ const v=parseInt(r["№ з/п"]??0,10); if(!Number.isNaN(v)&&v>max) max=v; }
  await addCaseRow(caseId,{...mapped,"№ з/п":String(max+1)});
}
async function exportCurrentFlow(){
  // Determine export profile for current sheet
  const getExportProfileForSheet = (sheetKey)=>{
    const cfg = state.sheetSettings[sheetKey] || {};
    return cfg.export || {pageSize:"A4",orientation:"portrait",exportHiddenCols:[],rowFilters:[]};
  };

  if(state.mode==="case" && state.caseId){
    const cases=await getAllCases();
    const c=cases.find(x=>x.id===state.caseId) || {id:state.caseId, caseIndex:"", caseTitle:""};
    const rows=await getCaseRows(state.caseId);
    rows.sort((a,b)=>parseInt(a["№ з/п"]??0,10)-parseInt(b["№ з/п"]??0,10));

    const op=await modalOpen({
      title:"Експорт поточного (опис справи)",
      bodyNodes:[el("div",{className:"muted",textContent:`${c.caseIndex||""} — ${c.caseTitle||""}`})],
      actions:[
        btn("Скасувати","cancel","btn"),
        btn("JSON","json","btn btn-primary"),
        btn("DOCX","docx","btn btn-primary"),
        btn("XLSX","xlsx","btn btn-primary"),
        btn("PDF","pdf","btn btn-primary"),
      ]
    });
    if(op.type==="cancel") return;

    const stamp=nowStamp();
    if(op.type==="json"){
      const payload={meta:{type:"case_description",exportedAt:new Date().toISOString(),case:c},rows};
      downloadBlob(new Blob([JSON.stringify(payload,null,2)],{type:"application/json"}), makeCaseExportFileName(c.caseIndex,c.caseTitle,stamp));
      return;
    }

    const cols=CASE_DESC_COLUMNS.map(x=>x.name);
    const flatRows = rows.map(r=>{
      const o={}; cols.forEach(k=>o[k]=r[k]??""); return o;
    });

    const title = "Внутрішній опис справи";
    const subtitle = `Справа: ${c.caseIndex||""} — ${c.caseTitle||""}\nЕкспорт: ${new Date().toLocaleString()}`;
    const filenameBase = `Opis_spravy_${c.caseIndex||""}_${c.caseTitle||""}`;

    if(op.type==="docx"){
      exportDOCXTable({title, subtitle, columns:cols, rows:flatRows, filenameBase});
      return;
    }
    if(op.type==="xlsx"){
      exportXLSXTable({title, columns:cols, rows:flatRows, filenameBase});
      return;
    }
    if(op.type==="pdf"){
      exportPDFTable({title, subtitle, columns:cols, rows:flatRows, filenameBase, pageSize:"A4", orientation:"portrait"});
      return;
    }
    return;
  }

  const sheet=currentSheet(); 
  const rows=await getRows(currentDataKey());

  // columns to export: visible-for-export = (not exportHiddenCols)
  const viewVisible=visibleColumns(sheet);
  const exportProfile = getExportProfileForSheet(sheet.key);
  const exportHidden = new Set(exportProfile.exportHiddenCols || []);
  const exportCols = sheet.columns.map(c=>c.name).filter(n=>!exportHidden.has(n));

  const op=await modalOpen({
    title:"Експорт поточного листа",
    bodyNodes:[
      el("div",{className:"muted",textContent:sheet.title}),
      el("div",{className:"muted",textContent:`Профіль: ${exportProfile.pageSize||"A4"} / ${exportProfile.orientation||"portrait"}; колонки (експорт): ${exportCols.length}`})
    ],
    actions:[
      btn("Скасувати","cancel","btn"),
      btn("JSON","json","btn btn-primary"),
      btn("DOCX","docx","btn btn-primary"),
      btn("XLSX","xlsx","btn btn-primary"),
      btn("PDF","pdf","btn btn-primary"),
    ]
  });
  if(op.type==="cancel") return;

  if(op.type==="json") return exportJournalAsJSON({sheet, rows, sheetExportProfile:exportProfile, visibleColumnsForView:viewVisible});
  if(op.type==="docx") return exportJournalAsDOCX({sheet, rows, columns:exportCols, sheetExportProfile:exportProfile});
  if(op.type==="xlsx") return exportJournalAsXLSX({sheet, rows, columns:exportCols, sheetExportProfile:exportProfile});
  if(op.type==="pdf")  return exportJournalAsPDF({sheet, rows, columns:exportCols, sheetExportProfile:exportProfile});
}
async function exportAllFlow(){
  // Export ALL data as ZIP(JSON set). If spaces/subjournals exist,
  // export virtual journals: (space/subjournal) × (sheet)
  const baseSheets = state.sheets;
  let sheets = baseSheets;

  try {
    const spaces = await ensureSpaces();
    const nodes = (spaces || []).filter(s => s && (s.kind === "space" || s.kind === "subjournal"));

    if (nodes.length) {
      const virt = [];
      for (const node of nodes) {
        for (const sh of baseSheets) {
          virt.push({
            ...JSON.parse(JSON.stringify(sh)),
            key: journalKey(node.id, sh.key),
            title: `${node.name} / ${sh.title}`,
          });
        }
      }
      sheets = virt;
    }
  } catch (e) {
    // If spaces subsystem isn't available, fall back to base sheets
    sheets = baseSheets;
  }

  const allRowsBySheet = new Map();
  for (const sh of sheets) {
    allRowsBySheet.set(sh.key, await getRows(sh.key));
  }

  const cases = await getAllCases();
  const caseRowsByCaseId = new Map();
  for (const c of cases) {
    caseRowsByCaseId.set(c.id, await getCaseRows(c.id));
  }

  await exportAllZipJSON({ sheets, allRowsBySheet, cases, caseRowsByCaseId });
}

async function importJsonFile(input){
  const file=input.files && input.files[0]; if(!file) return;
  const text=await file.text();
  try{
    const parsed=JSON.parse(text);
    await importJsonWizard(parsed);
  }catch(e){
    alert("Помилка імпорту JSON: "+e.message);
  }finally{
    input.value="";
    render();
  }
}
async function importZipFile(input){
  const file=input.files && input.files[0]; if(!file) return;
  const buf=await file.arrayBuffer();
  let entries;
  try{ entries=unzipStoreEntries(buf).filter(x=>x.name.toLowerCase().endsWith(".json")); }
  catch(e){ alert("Не вдалося прочитати ZIP: "+e.message); input.value=""; return; }
  if(!entries.length){ alert("У ZIP немає JSON."); input.value=""; return; }
  const ok=confirm(`Знайдено ${entries.length} JSON файлів.\n\nІмпорт замінить поточні дані.\nПродовжити?`);
  if(!ok){ input.value=""; return; }
  await clearAllRows(); await clearAllCasesAndRows();
  for(const ent of entries){
    try{ const parsed=JSON.parse(new TextDecoder().decode(ent.data)); await importPayload(parsed,true); }
    catch(e){ console.warn("ZIP import fail",ent.name,e); }
  }
  input.value=""; alert("Імпорт ZIP завершено."); render();
}

function colLettersToIndex(ref){
  // e.g. "A1" -> 1, "AA10" -> 27
  const m=/^([A-Z]+)\d+$/.exec(ref||"");
  if(!m) return null;
  const s=m[1];
  let n=0;
  for(let i=0;i<s.length;i++){ n = n*26 + (s.charCodeAt(i)-64); }
  return n;
}
function getCellTextFromXml(cellEl, sharedStrings){
  if(!cellEl) return "";
  const t = cellEl.getAttribute("t") || "";
  if(t==="inlineStr"){
    const tEl=cellEl.getElementsByTagName("t")[0];
    return tEl ? (tEl.textContent||"") : "";
  }
  const vEl=cellEl.getElementsByTagName("v")[0];
  const v = vEl ? (vEl.textContent||"") : "";
  if(t==="s"){
    const idx=parseInt(v,10);
    return Number.isFinite(idx) && sharedStrings[idx]!=null ? sharedStrings[idx] : "";
  }
  return v;
}
function parseSharedStringsXml(xmlText){
  const out=[];
  try{
    const doc=new DOMParser().parseFromString(xmlText,"application/xml");
    const si = Array.from(doc.getElementsByTagName("si"));
    for(const node of si){
      // shared string may be rich text; concatenate all <t>
      const ts = Array.from(node.getElementsByTagName("t"));
      out.push(ts.map(x=>x.textContent||"").join(""));
    }
  }catch(_e){}
  return out;
}

async function importXlsxFile(input){
  const file=input.files && input.files[0];
  if(!file) return;
  const sheet=currentSheet();
  if(!sheet){ input.value=""; return; }

  const op=await modalOpen({
    title:"Імпорт XLSX у поточний лист",
    bodyNodes:[
      el("div",{className:"muted",textContent:`Лист: ${sheet.title}`}),
      el("div",{className:"muted",textContent:"Файл має починатися з даних (без заголовків). Колонки — в тому ж порядку, що і в журналі."}),
    ],
    actions:[
      btn("Скасувати","cancel","btn"),
      btn("Додати рядки","append","btn btn-primary"),
      btn("Замінити дані","replace","btn btn-primary"),
    ]
  });
  if(op.type==="cancel"){ input.value=""; return; }

  try{
    const buf=await file.arrayBuffer();
    const entries=await unzipEntries(buf);
    const findEntry=(name)=>entries.find(e=>e.name===name);
    const sheetEntry = findEntry("xl/worksheets/sheet1.xml") || entries.find(e=>e.name.startsWith("xl/worksheets/") && e.name.endsWith(".xml"));
    if(!sheetEntry) throw new Error("XLSX: не знайдено xl/worksheets/sheet1.xml");

    const ssEntry = findEntry("xl/sharedStrings.xml");
    const sharedStrings = ssEntry ? parseSharedStringsXml(new TextDecoder().decode(ssEntry.data)) : [];

    const sheetXml = new TextDecoder().decode(sheetEntry.data);
    const doc=new DOMParser().parseFromString(sheetXml,"application/xml");
    const rows = Array.from(doc.getElementsByTagName("row"));
    const colDefs = sheet.columns;

    if(op.type==="replace"){
      const ok=confirm(`Замінити ВСІ дані листа «${sheet.title}» імпортом з Excel?`);
      if(!ok){ input.value=""; return; }
      await clearRows(currentDataKey());
    }

    const existing = await getRows(currentDataKey());
    const orderCol = sheet.orderColumn;
    let nextNum = orderCol ? nextOrder(existing, orderCol) : null;

    let imported=0;
    const errors=[];
    for(const rEl of rows){
      const cells = Array.from(rEl.getElementsByTagName("c"));
      if(!cells.length) continue;
      const byIndex = new Map();
      for(const cEl of cells){
        const ref=cEl.getAttribute("r")||"";
        const idx=colLettersToIndex(ref);
        if(!idx) continue;
        byIndex.set(idx, cEl);
      }

      const data={};
      let any=false;
      for(let ci=0; ci<colDefs.length; ci++){
        const def=colDefs[ci];
        const cellEl = byIndex.get(ci+1);
        let val = getCellTextFromXml(cellEl, sharedStrings);
        val = String(val??"").trim();

        if(def.type==="int"){
          if(val===""){
            // ok
          } else if(/^\d+$/.test(val)){
            // ok
          } else if(/^\d+(?:\.0+)?$/.test(val)){
            // Excel numeric like 12.0
            val = String(parseInt(val,10));
          } else {
            errors.push(`Рядок ${rEl.getAttribute("r")||"?"}: поле «${def.name}» має бути числом`);
            val = "";
          }
        }

        if(def.type==="date"){
          if(val===""){
            // ok
          } else {
            const p=parseUAdate(val);
            if(p) val=p;
            else {
              // try excel serial
              const serial = Number(val);
              const ex = excelSerialToUAdate(serial);
              if(ex) val=ex;
              else {
                errors.push(`Рядок ${rEl.getAttribute("r")||"?"}: поле «${def.name}» має дату ДД.ММ.РР`);
                val="";
              }
            }
          }
        }

        if(val!=="") any=true;
        data[def.name]=val;
      }
      if(!any) continue;

      if(orderCol && (!data[orderCol] || !/^\d+$/.test(String(data[orderCol]).trim()))){
        data[orderCol] = String(nextNum++);
      }

      await addRow(currentDataKey(),{data, subrows:[]});
      imported++;
    }
    input.value="";
    render();
    const msg = `Імпорт XLSX завершено. Додано рядків: ${imported}.`;
    if(errors.length){
      alert(msg + "\n\nПопередження (перші 5):\n" + errors.slice(0,5).join("\n"));
    } else {
      alert(msg);
    }
  }catch(e){
    console.error(e);
    alert("Помилка імпорту XLSX: " + e.message);
    input.value="";
  }
}

// ------------------------
// JSON import wizard (column-index mapping)

function journalSourceFromParsed(parsed){
  const metaTitle = parsed?.meta?.title || parsed?.sheet?.title || "";
  const sourceSheet = parsed?.sheet || null;
  // v2 preferred
  if(Array.isArray(parsed?.rowsV2)){
    const rows = parsed.rowsV2.map(r=>({ cells: Array.isArray(r.cells)?r.cells:[], subrows:r.subrows||[] }));
    const colsCount = Number.isFinite(parsed.columnsCount) ? parsed.columnsCount : Math.max(0, ...rows.map(x=>x.cells.length));
    const sourceCols = Math.max(colsCount, Math.max(0,...rows.map(x=>x.cells.length)));
    return { title: metaTitle, key: parsed?.meta?.key||"", sourceSheet, sourceCols, rows };
  }

  // legacy: rows with data/exportData objects
  const colNames = Array.isArray(sourceSheet?.columns) ? sourceSheet.columns.map(c=>c.name) : [];
  const legacyRows = Array.isArray(parsed?.rows) ? parsed.rows : [];
  const rows = legacyRows.map(r=>{
    const obj = r.exportData || r.data || {};
    const cells = colNames.map(n=>String(obj?.[n] ?? ""));
    return { cells, subrows: r.subrows||[] };
  });
  const sourceCols = colNames.length || Math.max(0,...rows.map(x=>x.cells.length));
  return { title: metaTitle, key: parsed?.meta?.key||"", sourceSheet, sourceCols, rows };
}

function buildMappingUI({targetCols, sourceCols}){
  const wrap = el("div",{className:"import-map"});
  const info = el("div",{className:"muted",textContent:`В журналі ${targetCols} колонок. В імпорті ${sourceCols} колонок.`});
  const table = el("table",{className:"map-table"});
  const trTop = el("tr",{});
  const trBottom = el("tr",{});
  const mappingInputs = [];

  for(let i=0;i<targetCols;i++){
    trTop.appendChild(el("td",{textContent:String(i+1)}));
    const td = el("td",{});
    const inp = el("input",{type:"number",min:"0",max:String(sourceCols),value: String(Math.min(i+1, sourceCols||0))});
    inp.style.width = "100%";
    mappingInputs.push(inp);
    td.appendChild(inp);
    trBottom.appendChild(td);
  }
  table.appendChild(trTop);
  table.appendChild(trBottom);

  const hint = el("div",{className:"muted",textContent:"Нижній рядок — номер колонки з імпорту (0 = пропустити)."});
  wrap.appendChild(info);
  wrap.appendChild(table);
  wrap.appendChild(hint);

  const getMapping = ()=> mappingInputs.map(inp=>{
    const v = parseInt(inp.value,10);
    return Number.isFinite(v) ? v : 0;
  });
  return { node: wrap, getMapping };
}

async function importJsonWizard(parsed){
  // Optional: full import bundle (single JSON containing multiple journals)
  if(Array.isArray(parsed?.journals)){
    const pre = await modalOpen({
      title:"Майстер повного імпорту JSON",
      bodyNodes:[
        el("div",{className:"muted",textContent:`Журналів у файлі: ${parsed.journals.length}`}),
        el("div",{className:"muted",textContent:"Імпорт виконується по черзі для кожного журналу."})
      ],
      actions:[
        btn("Скасувати","cancel","btn"),
        btn("Додати","append","btn btn-primary"),
        btn("Замінити","replace","btn btn-danger"),
      ]
    });
    if(pre.type==="cancel") return;
    const replace = (pre.type==="replace");
    // cases (if any) - import after journals
    for(const j of parsed.journals){
      // reuse single-journal wizard but force mode
      if(replace){
        // only clear once per target sheet inside applyJournalImport
      }
      await importJsonWizard({ ...j, __bundleMode:true, __bundleReplace:replace });
    }
    if(Array.isArray(parsed?.cases)){
      for(const c of parsed.cases){ await importPayload(c); }
    }
    return;
  }
  // case descriptions: keep as-is
  if(parsed?.meta?.type==="case_description"){
    await importPayload(parsed);
    alert("Імпорт JSON завершено.");
    return;
  }

  if(parsed?.meta?.type!=="journal"){
    throw new Error("Невідомий формат JSON (очікується journal або case_description).")
  }

  const source = journalSourceFromParsed(parsed);
  const targetDefault = (state.mode==="sheet") ? currentSheet() : state.sheets[0];
  if(!targetDefault) throw new Error("Немає доступних листів для імпорту.");

  // choose target sheet
  const sel = el("select",{className:"input"});
  for(const sh of state.sheets){
    const opt = el("option",{value:sh.key,textContent:`${sh.title} (${sh.key})`});
    if(sh.key===targetDefault.key) opt.selected=true;
    sel.appendChild(opt);
  }

  let preType = "append";
  if(parsed?.__bundleMode){
    preType = parsed.__bundleReplace ? "replace" : "append";
  } else {
    const pre = await modalOpen({
      title:"Майстер імпорту JSON",
      bodyNodes:[
        el("div",{className:"muted",textContent:`Файл: ${source.title || "(без назви)"}`}),
        el("div",{textContent:"Оберіть лист, в який імпортувати:"}),
        sel,
      ],
      actions:[
        btn("Скасувати","cancel","btn"),
        btn("Додати","append","btn btn-primary"),
        btn("Замінити","replace","btn btn-danger"),
      ]
    });
    if(pre.type==="cancel") return;
    preType = pre.type;
  }

  const targetSheet = state.sheets.find(s=>s.key===sel.value) || targetDefault;

  // sheet title mismatch warning
  if(source.title && targetSheet.title && source.title !== targetSheet.title){
    const ok = await modalOpen({
      title:"Попередження",
      bodyNodes:[
        el("div",{textContent:`Назва листа в СЕДО: «${targetSheet.title}»`} ),
        el("div",{textContent:`Назва листа в JSON: «${source.title}»`} ),
        el("div",{className:"muted",textContent:"Назви відрізняються. Продовжити імпорт?"})
      ],
      actions:[ btn("Скасувати","cancel","btn"), btn("Продовжити","go","btn btn-primary") ]
    });
    if(ok.type!=="go") return;
  }

  const targetCols = (targetSheet.columns||[]).length;
  const sourceCols = Math.max(0, source.sourceCols||0);
  const {node, getMapping} = buildMappingUI({targetCols, sourceCols});
  const step = await modalOpen({
    title:"Зіставлення колонок",
    bodyNodes:[ node ],
    actions:[ btn("Скасувати","cancel","btn"), btn("Імпортувати","do","btn btn-primary") ]
  });
  if(step.type!=="do") return;

  const mapping = getMapping(); // 1-based source col numbers or 0
  const replace = (preType==="replace");
  await applyJournalImport({targetSheet, source, mapping, replace});
  if(!parsed?.__bundleMode) alert("Імпорт JSON завершено.");
}

function normalizeDateToDDMMYY(s){
  const t = String(s||"").trim();
  if(!t) return "";
  // dd.mm.yy or dd.mm.yyyy
  let m = /^([0-3]\d)\.([01]\d)\.(\d{2}|\d{4})$/.exec(t);
  if(m){
    const dd=m[1], mm=m[2];
    const yy = m[3].length===4 ? m[3].slice(-2) : m[3];
    return `${dd}.${mm}.${yy}`;
  }
  // yyyy-mm-dd
  m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(t);
  if(m){
    return `${m[3]}.${m[2]}.${m[1].slice(-2)}`;
  }
  return t;
}

async function applyJournalImport({targetSheet, source, mapping, replace}){
  if(replace){
    await clearRows(journalKeyForSheet(targetSheet.key));
  }
  const cols = targetSheet.columns||[];
  const errors=[];
  let imported=0;
  // determine next auto-number if needed
  let nextNum=null;
  const orderCol = targetSheet.orderColumn;
  if(orderCol){
    const existing = await getRows(journalKeyForSheet(targetSheet.key));
    let max=0;
    for(const r of existing){
      const v=parseInt(r?.data?.[orderCol]||"",10);
      if(Number.isFinite(v)) max=Math.max(max,v);
    }
    nextNum=max+1;
  }

  for(let i=0;i<source.rows.length;i++){
    const srcRow = source.rows[i];
    const cells = Array.isArray(srcRow.cells) ? srcRow.cells : [];
    const data={};
    for(let tIdx=0;tIdx<cols.length;tIdx++){
      const col = cols[tIdx];
      const srcColNum = parseInt(mapping[tIdx]||0,10);
      const srcIdx = srcColNum ? (srcColNum-1) : -1;
      let v = (srcIdx>=0 && srcIdx<cells.length) ? String(cells[srcIdx] ?? "") : "";
      if(col.type==="date") v = normalizeDateToDDMMYY(v);
      // basic int sanitize
      if(col.type==="int"){
        v = String(v||"").trim();
        if(v!=="" && !/^\d+$/.test(v)){
          errors.push(`Рядок ${i+1}, колонка ${tIdx+1}: очікується число`);
          v = "";
        }
      }
      data[col.name]=v;
    }
    // auto-fill order column if empty
    if(orderCol && (!data[orderCol] || String(data[orderCol]).trim()==="") && nextNum!=null){
      data[orderCol]=String(nextNum++);
    }
    await addRow(journalKeyForSheet(targetSheet.key),{data, subrows: srcRow.subrows||[]});
    imported++;
  }
  if(errors.length){
    alert(`Імпорт завершено. Додано: ${imported}.\n\nПопередження (перші 10):\n`+errors.slice(0,10).join("\n"));
  }
  render();
}
async function importPayload(parsed){
  if(parsed?.meta?.type==="journal"){
    const key=parsed.meta.key;
    let sheet=state.sheets.find(s=>s.key===key);
    if(!sheet){
      const sh=parsed.sheet;
      const custom={key,title:parsed.meta.title||sh?.title||key,orderColumn:sh?.orderColumn||null,columns:sh?.columns||[],addFields:sh?.addFields||(sh?.columns||[]).map(c=>c.name),subrows:sh?.subrows||null,export:sh?.export||{pageSize:"A4",orientation:"portrait"}};
      state.sheets.push(custom);
      await saveUserSheets(state.sheets.filter(s=>isCustomKey(s.key)));
      await loadConfig();
    }
    // v2: rowsV2 (cells array) preferred
    if(Array.isArray(parsed.rowsV2)){
      const colNames = (sheet?.columns||[]).map(c=>c.name);
      for(const r of parsed.rowsV2){
        const cells = Array.isArray(r.cells)?r.cells:[];
        const data={};
        for(let i=0;i<colNames.length;i++) data[colNames[i]] = String(cells[i] ?? "");
        await addRow(key,{data, subrows:r.subrows||[]});
      }
    } else if(Array.isArray(parsed.rows)){
      for(const r of parsed.rows){
        await addRow(key,{data:r.data||{}, subrows:r.subrows||[]});
      }
    }
    return;
  }
  if(parsed?.meta?.type==="case_description"){
    const c=parsed.meta.case||{};
    const id=await addCase({caseIndex:c.caseIndex||"", caseTitle:c.caseTitle||"", createdFrom:c.createdFrom||"import", createdAt:c.createdAt||new Date().toISOString()});
    if(Array.isArray(parsed.rows)){
      for(const r of parsed.rows) await addCaseRow(id,{...r});
    }
  }
}
async function clearCurrent(){
  if(state.mode==="case"){ alert("Очистку опису справи додамо наступним кроком."); return; }
  const sh=currentSheet(); const ok=confirm(`Очистити лист «${sh.title}» повністю?`); if(!ok) return;
  await clearRows(currentDataKey()); render();
}
async function clearAll(){
  const ok=confirm("Очистити ВСІ листи, всі справи і всі описи?"); if(!ok) return;
  await clearAllRows(); await clearAllCasesAndRows(); render();
}
await loadConfig();
await ensureDefaultTransferTemplates(state.sheets);
const __spaces = await ensureSpaces();
fillSpaceSelect(__spaces);

// Space switching (isolated per-space hierarchy)
if(spaceSelect){
  spaceSelect.onchange = async ()=>{
    state.spaceId = spaceSelect.value || state.spaceId;
    state.jtree = await ensureJournalTree(state.spaceId, state.sheets);
    state.journalPath = [];
    ensureValidJournalPath();
    await saveView();
    renderJournalNav();
    render();
  };
}

// Add new root space (no editor yet)
if(btnAddSpace){
  btnAddSpace.onclick = async ()=>{
    const name = (prompt("Назва нового простору:", "Новий простір")||"").trim();
    if(!name) return;
    const spaces = await cfgGet("spaces_v1") || [];
    const id = `space_${Date.now().toString(36)}`;
    spaces.push({id, name, parentId:null, kind:"space", meta:{}});
    await cfgSet("spaces_v1", spaces);
    state.spaces = await ensureSpaces();
    fillSpaceSelect(state.spaces);
    // initialize tree for the new space
    state.spaceId = id;
    state.jtree = await ensureJournalTree(state.spaceId, state.sheets);
    state.journalPath = [];
    ensureValidJournalPath();
    await saveView();
    renderJournalNav();
    render();
  };
}
ensureValidJournalPath();
renderJournalNav();
await fillCaseSelect();
window.addEventListener("resize", ()=>{ updateStickyOffsets(); });
render();
