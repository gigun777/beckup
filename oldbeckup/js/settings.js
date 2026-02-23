// js/settings.js
import { cfgGet, cfgSet } from "./db.js";
import { DEFAULT_SHEETS } from "./schema.js";
import { el, modalOpen, btn } from "./ui.js";
import { buildTransferUI } from "./transfer_ui.js?v=12.5.1.3";
import { BackupCore } from "./backup_core.js?v=12.6.10";

async function openTabBackupDialog(tabId){
  // For settings tabs we use SETTINGS_PARTIAL, for transfer tab use TRANSFER_RULES_PARTIAL.
  const isTransfer = tabId === "transfer";
  const expectedType = isTransfer ? BackupCore.TYPES.TRANSFER_RULES_PARTIAL : BackupCore.TYPES.SETTINGS_PARTIAL;
  const opts = BackupCore.getBackupOptionsForTab(tabId);

  const wrap = el("div",{style:"min-width:560px; max-width:760px"});
  wrap.appendChild(el("div",{className:"muted",textContent:"Оберіть, які саме параметри потраплять до бекапу. Можна експортувати або імпортувати вибірково."}));

  const list = el("div",{style:"margin-top:10px; display:flex; flex-direction:column; gap:8px"});
  const checks = new Map();
  for(const o of opts){
    const row = el("label",{className:"row",style:"gap:10px; align-items:center; cursor:pointer; justify-content:flex-start"});
    const cb = el("input",{type:"checkbox"});
    cb.checked = !!o.checked;
    checks.set(o.id, cb);
    row.appendChild(cb);
    row.appendChild(el("div",{textContent:o.label}));
    list.appendChild(row);
  }
  wrap.appendChild(list);

  const tools = el("div",{className:"row",style:"flex-wrap:wrap; gap:8px; margin-top:10px"});
  const btnAll = el("button",{className:"btn",textContent:"Вибрати все"});
  const btnNone = el("button",{className:"btn",textContent:"Зняти все"});
  btnAll.onclick=()=>{ for(const cb of checks.values()) cb.checked=true; };
  btnNone.onclick=()=>{ for(const cb of checks.values()) cb.checked=false; };
  tools.appendChild(btnAll);
  tools.appendChild(btnNone);
  wrap.appendChild(tools);

  const res = await modalOpen({
    title: isTransfer ? "Бекап: Перенесення" : "Бекап: Налаштування вкладки",
    bodyNodes:[wrap],
    actions:[
      btn("Експорт","export","btn btn-primary"),
      btn("Імпорт","import","btn"),
      btn("Закрити","cancel","btn"),
    ]
  });

  if(res.type==="export"){
    const optionIds = Array.from(checks.entries()).filter(([,cb])=>cb.checked).map(([id])=>id);
    try{
      if(isTransfer){
        await BackupCore.exportTransferRulesPartialBackup();
      } else {
        await BackupCore.exportSettingsPartialBackup({scope:"tab", tabId, optionIds});
      }
    }catch(e){ alert(e?.message||String(e)); }
  }

  if(res.type==="import"){
    try{
      // Import file. Validation is done inside backup core (expectedType).
      const input = document.createElement("input");
      input.type = "file";
      input.accept = "application/json,.json";
      input.onchange = () => {
        const file = input.files && input.files[0];
        if(!file) return;
        const reader = new FileReader();
        reader.onload = async()=>{
          try{
            await BackupCore.importBackupFromText(String(reader.result||""), expectedType);
            location.reload();
          }catch(e){ alert(e?.message||String(e)); }
        };
        reader.readAsText(file);
      };
      input.click();
    }catch(e){ alert(e?.message||String(e)); }
  }
}

export async function getAllSheets(){
  // If user customized even default sheets, we store the full schema in config.
  const storedAll = await cfgGet("all_sheets");
  if(Array.isArray(storedAll) && storedAll.length){
    return storedAll;
  }
  const user = (await cfgGet("user_sheets")) || [];
  // Default behavior (requested): allow subrows for all columns by default.
  // Users can disable per-column in the constructor.
  const base = JSON.parse(JSON.stringify(DEFAULT_SHEETS));
  for(const sh of base){
    for(const c of (sh.columns||[])){
      if(c.editable===false) continue;
      if(c.computed) continue;
      c.subrows = (c.subrows!==false); // default true
    }
  }
  for(const sh of (user||[])){
    for(const c of (sh.columns||[])){
      if(c.subrows===undefined) c.subrows = true;
    }
  }
  return [...base, ...(user||[])];
}
export async function saveUserSheets(userSheets){ await cfgSet("user_sheets", userSheets); }

export async function saveAllSheets(allSheets){ await cfgSet("all_sheets", allSheets); }

export async function getSheetSettings(){ return (await cfgGet("sheet_settings")) || {}; }
export async function saveSheetSettings(settings){ await cfgSet("sheet_settings", settings); }

export async function getAddFieldConfig(){ return (await cfgGet("add_fields")) || {}; }
export async function saveAddFieldConfig(cfg){ await cfgSet("add_fields", cfg); }

export function genKeyFromTitle(title){
  const base = String(title||"").toLowerCase().trim().replace(/\s+/g,"_").replace(/[^\p{L}\p{N}_]+/gu,"");
  return "custom_" + base.slice(0,40) + "_" + Math.floor(Math.random()*10000);
}

function dndList({items, onReorder, renderItem}){
  const wrap=el("div",{});
  let dragIndex=-1;
  const render=()=>{
    wrap.innerHTML="";
    items.forEach((it,idx)=>{
      const row=el("div",{className:"row",style:"justify-content:space-between; border:1px solid #e6e6e6; padding:8px; border-radius:12px; margin:6px 0;"});
      row.draggable=true;
      row.ondragstart=()=>{ dragIndex=idx; row.style.opacity="0.6"; };
      row.ondragend=()=>{ row.style.opacity="1"; };
      row.ondragover=(e)=>{ e.preventDefault(); row.style.background="#f6f6f6"; };
      row.ondragleave=()=>{ row.style.background="#fff"; };
      row.ondrop=(e)=>{ e.preventDefault(); row.style.background="#fff";
        const to=idx;
        if(dragIndex<0 || dragIndex===to) return;
        const [m]=items.splice(dragIndex,1);
        items.splice(to,0,m);
        dragIndex=-1;
        onReorder();
        render();
      };
      row.appendChild(renderItem(it,idx));
      wrap.appendChild(row);
    });
  };
  render();
  return wrap;
}

// Stage 2: Simplified view templates editor (per sheet)

// Stage 3 (constructor-only): Simplified view templates editor (per sheet)
// NOTE: No rendering of cards yet. This only builds and stores the template layout (rows×cols, cells→blocks).
async function openSimplifiedTemplatesEditor({sheet, onDirty}){
  if(!sheet) return;
  if(!sheet.simplified) sheet.simplified = { enabled:false, on:false, activeTemplateId:null, templates:[] };
  if(!Array.isArray(sheet.simplified.templates)) sheet.simplified.templates = [];
  // IMPORTANT: no "default" template is created automatically. Normal table view is the default when simplified view is OFF.
  if(!sheet.simplified.activeTemplateId && sheet.simplified.templates.length){
    sheet.simplified.activeTemplateId = sheet.simplified.templates[0].id;
  }

  const templates = sheet.simplified.templates;
  let activeId = sheet.simplified.activeTemplateId || (templates[0]&&templates[0].id) || null;

  // Make the generic modal wider for the constructor.
  const modalEl = document.querySelector("#modalBackdrop .modal");
  if(modalEl) modalEl.classList.add("modal-wide");

  const wrap = el("div",{style:"display:flex; gap:12px; align-items:stretch; min-width:960px; max-width:1200px"});
  const left = el("div",{style:"width:280px"});
  const right = el("div",{style:"flex:1; min-width:0"});
  wrap.appendChild(left);
  wrap.appendChild(right);

  const listTitle = el("div",{className:"row",style:"justify-content:space-between; align-items:center"});
  listTitle.appendChild(el("b",{textContent:"Шаблони спрощеного перегляду"}));
  const addBtn = el("button",{className:"btn btn-primary",textContent:"＋ Додати шаблон"});
  listTitle.appendChild(addBtn);
  left.appendChild(listTitle);

  const list = el("div",{style:"margin-top:8px"});
  left.appendChild(list);

  const hint = el("div",{className:"muted",style:"margin-top:10px"});
  hint.textContent = "Шаблони — це інструкції для конструктора спрощеного перегляду. Рендеринг карток буде на кроці 4.";
  left.appendChild(hint);

  function getActive(){ return templates.find(t=>t.id===activeId) || null; }

  const ensureTheme = ()=>{
    const simp = sheet.simplified || (sheet.simplified={ enabled:false, on:false, activeTemplateId:null, templates:[] });
    if(!simp.theme || typeof simp.theme!=="object") simp.theme = {};
    const th = simp.theme;
    if(!Number.isFinite(th.radius)) th.radius = 16;
    if(typeof th.showBorders!=="boolean") th.showBorders = false;
    if(typeof th.glass!=="boolean") th.glass = true;
    if(typeof th.gradient!=="boolean") th.gradient = false;
    if(typeof th.cardColor!=="string") th.cardColor = "#ff3b30";
    if(typeof th.cardOpacity!=="number") th.cardOpacity = 0.92;
    if(typeof th.gradFrom!=="string") th.gradFrom = "#ff3b30";
    if(typeof th.gradTo!=="string") th.gradTo = "#ff9500";
    if(typeof th.bgColor!=="string") th.bgColor = "";
    if(typeof th.borderColor!=="string") th.borderColor = "rgba(255,255,255,0.30)";
    if(!Number.isFinite(th.blur)) th.blur = 18;
    if(typeof th.customCss!=="string") th.customCss = "";
    // Conditional card background rules (per-row) — used in simplified cards rendering (Step 4)
    if(typeof th.cardBgRulesEnabled!=="boolean") th.cardBgRulesEnabled = false;
    if(!Array.isArray(th.cardBgRules)) th.cardBgRules = [];
    return th;
  };

  const ensureLayout = (t)=>{
    if(!t.layout || typeof t.layout!=="object") t.layout = {};
    if(!Number.isFinite(t.layout.rows)) t.layout.rows = 2;
    if(!Number.isFinite(t.layout.cols)) t.layout.cols = 2;
    if(!t.layout.cells || typeof t.layout.cells!=="object") t.layout.cells = {};
  };

  // Helper: read/write cell configuration (backward compatible)
  // Old format: cells[key] = [blocks...]
  // New format: cells[key] = { join:{op,delimiter}, blocks:[...] }
  const getCellConfig = (t, key)=>{
    ensureLayout(t);
    const raw = t.layout.cells[key];
    // Backward compat:
    //  - old format: cells[key] = [blocks...]
    //  - v12.4.0.2+: cells[key] = { join:{op,delimiter}, blocks:[...] }  (cell-level join)
    // New format: cells[key] = { joinAll:{op,delimiter}, joins:[{op,delimiter}...], blocks:[...] }
    if(Array.isArray(raw)){
      return { joinAll:{op:"newline", delimiter:""}, joins:[], blocks: raw };
    }
    if(raw && typeof raw==="object"){
      const blocks = Array.isArray(raw.blocks) ? raw.blocks : [];
      const joinAllRaw = (raw.joinAll && typeof raw.joinAll==="object") ? raw.joinAll
                       : (raw.join && typeof raw.join==="object") ? raw.join
                       : {};
      const joinAllOp = (joinAllRaw.op==="concat"||joinAllRaw.op==="seq"||joinAllRaw.op==="newline") ? joinAllRaw.op : "newline";
      const joinAllDelimiter = (joinAllOp==="concat") ? String(joinAllRaw.delimiter ?? " ") : "";
      const joinsRaw = Array.isArray(raw.joins) ? raw.joins : [];
      const joins = joinsRaw.map(j=>{
        const op = (j && (j.op==="concat"||j.op==="seq"||j.op==="newline")) ? j.op : joinAllOp;
        const delimiter = (op==="concat") ? String(j?.delimiter ?? (joinAllDelimiter||" ")) : "";
        return {op, delimiter};
      });
      return { joinAll:{op:joinAllOp, delimiter:joinAllDelimiter}, joins, blocks };
    }
    return { joinAll:{op:"newline", delimiter:""}, joins:[], blocks: [] };
  };

  const setCellConfig = (t, key, cfg)=>{
    ensureLayout(t);
    const blocks = Array.isArray(cfg?.blocks) ? cfg.blocks : [];
    const joinAllOp = (cfg?.joinAll?.op==="concat"||cfg?.joinAll?.op==="seq"||cfg?.joinAll?.op==="newline") ? cfg.joinAll.op : "newline";
    const joinAllDelimiter = (joinAllOp==="concat") ? String(cfg?.joinAll?.delimiter ?? " ") : "";
    const joinsIn = Array.isArray(cfg?.joins) ? cfg.joins : [];
    const joins = [];
    const want = Math.max(0, blocks.length-1);
    for(let i=0;i<want;i++){
      const j = joinsIn[i] || {};
      const op = (j.op==="concat"||j.op==="seq"||j.op==="newline") ? j.op : joinAllOp;
      const delimiter = (op==="concat") ? String(j.delimiter ?? (joinAllDelimiter||" ")) : "";
      joins.push({op, delimiter});
    }
    t.layout.cells[key] = { joinAll:{op:joinAllOp, delimiter:joinAllDelimiter}, joins, blocks };
  };

  const colOptions = ()=> sheet.columns.map((c,i)=>({value:String(i), label:`${i+1}. ${c.name}`}));

  const renderList = ()=>{
    list.innerHTML="";
    if(!templates.length){
      list.appendChild(el("div",{className:"muted",textContent:"Немає шаблонів. Натисніть “＋ Додати шаблон”."}));
      return;
    }
    templates.forEach(t=>{
      const row = el("div",{className:"row",style:`justify-content:space-between; align-items:center; padding:8px; border-radius:12px; border:1px solid #e6e6e6; margin:6px 0; cursor:pointer; ${t.id===activeId?"box-shadow:0 0 0 2px rgba(0,0,0,0.12) inset":""}`});
      const name = el("div",{textContent:t.name||t.id, style:"max-width:180px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap"});
      const badge = el("span",{className:"pill",textContent:(t.id===activeId?"Активний":"")});
      row.onclick=()=>{ activeId=t.id; sheet.simplified.activeTemplateId=activeId; onDirty(); render(); };
      row.appendChild(name);
      row.appendChild(badge);
      list.appendChild(row);
    });
  };

  const openCellEditor = async ({t, r, c})=>{
    ensureLayout(t);
    const key = `${r}-${c}`;
    const cfg = getCellConfig(t, key);
    const blocks = JSON.parse(JSON.stringify(cfg.blocks));

    const body = el("div",{style:"min-width:640px; max-width:860px"});
    body.appendChild(el("div",{className:"muted",textContent:`Комірка: R${r+1} × C${c+1}`}));

    // How to join blocks inside this cell (global + per-gap)
    const joinAll = JSON.parse(JSON.stringify(cfg.joinAll || {op:"newline", delimiter:""}));
    let joins = Array.isArray(cfg.joins) ? JSON.parse(JSON.stringify(cfg.joins)) : [];

    const joinRow = el("div",{className:"row",style:"gap:8px; align-items:center; margin-top:8px; flex-wrap:wrap"});
    joinRow.appendChild(el("div",{className:"muted",textContent:"Поєднання блоків (для всіх):"}));

    const joinAllSel = el("select",{className:"input",style:"min-width:220px"});
    joinAllSel.appendChild(el("option",{value:"newline",textContent:"З нової строки"}));
    joinAllSel.appendChild(el("option",{value:"concat",textContent:"Конкатенація"}));
    joinAllSel.appendChild(el("option",{value:"seq",textContent:"Послідовний запис"}));
    joinAllSel.value = joinAll.op || "newline";
    joinRow.appendChild(joinAllSel);

    const joinAllDelim = el("input",{className:"input",placeholder:"розділювач",value: (joinAllSel.value==="concat") ? (joinAll.delimiter||" ") : " ", style:"width:120px"});
    joinRow.appendChild(joinAllDelim);

    const applyAllBtn = el("button",{className:"btn",textContent:"Застосувати до всіх"});
    joinRow.appendChild(applyAllBtn);

    const refreshJoinAllUI = ()=>{
      const op = joinAllSel.value;
      joinAll.op = op;
      if(op==="concat"){
        joinAllDelim.disabled = false;
        joinAll.delimiter = (joinAllDelim.value ?? " ");
      } else {
        joinAllDelim.disabled = true;
        joinAll.delimiter = "";
      }
    };
    joinAllSel.onchange = ()=>{ refreshJoinAllUI(); };
    joinAllDelim.oninput = ()=>{ if(joinAllSel.value==="concat"){ joinAll.delimiter = joinAllDelim.value; } };
    refreshJoinAllUI();

    const normalizeJoins = ()=>{
      const want = Math.max(0, blocks.length-1);
      if(!Array.isArray(joins)) joins = [];
      while(joins.length < want){
        joins.push({op: joinAll.op, delimiter: joinAll.op==="concat" ? (joinAll.delimiter||" ") : ""});
      }
      joins = joins.slice(0, want);
      joins = joins.map(j=>{
        const op = (j && (j.op==="concat"||j.op==="seq"||j.op==="newline")) ? j.op : joinAll.op;
        const delimiter = (op==="concat") ? String(j?.delimiter ?? (joinAll.delimiter||" ")) : "";
        return {op, delimiter};
      });
    };

    applyAllBtn.onclick = ()=>{
      refreshJoinAllUI();
      normalizeJoins();
      for(let i=0;i<joins.length;i++){
        joins[i] = {op: joinAll.op, delimiter: joinAll.op==="concat" ? (joinAll.delimiter||" ") : ""};
      }
      renderBlocks();
    };

    body.appendChild(joinRow);



    const blocksWrap = el("div",{style:"margin-top:10px; display:flex; flex-direction:column; gap:10px"});
    body.appendChild(blocksWrap);

    const renderBlocks = ()=>{
      blocksWrap.innerHTML="";
      normalizeJoins();
      if(!blocks.length){
        blocksWrap.appendChild(el("div",{className:"muted",textContent:"Блоків немає. Додайте блок, щоб ця комірка відображала дані."}));
      }
      blocks.forEach((b, idx)=>{
        if(!b || typeof b!=="object") b = blocks[idx] = {sources:[], op:"concat", delimiter:" "};
        if(!Array.isArray(b.sources)) b.sources=[];
        if(!b.op) b.op="concat";
        if(b.delimiter==null) b.delimiter=" ";

        const card = el("div",{className:"card"});
        const head = el("div",{className:"row",style:"justify-content:space-between; align-items:center"});
        head.appendChild(el("b",{textContent:`Блок ${idx+1}`}));

        const move = el("div",{className:"row",style:"gap:6px"});
        const up = el("button",{className:"btn btn-ghost",textContent:"↑",title:"Вгору"});
        const down = el("button",{className:"btn btn-ghost",textContent:"↓",title:"Вниз"});
        up.onclick=()=>{ if(idx===0) return; const tmp=blocks[idx-1]; blocks[idx-1]=blocks[idx]; blocks[idx]=tmp; joins=[]; renderBlocks(); };
        down.onclick=()=>{ if(idx===blocks.length-1) return; const tmp=blocks[idx+1]; blocks[idx+1]=blocks[idx]; blocks[idx]=tmp; joins=[]; renderBlocks(); };
        move.appendChild(up); move.appendChild(down);

        const del = el("button",{className:"btn danger",textContent:"🗑",title:"Видалити блок"});
        del.onclick=()=>{ blocks.splice(idx,1); joins=[]; renderBlocks(); };

        const rightBtns = el("div",{className:"row",style:"gap:6px"});
        rightBtns.appendChild(move);
        rightBtns.appendChild(del);

        head.appendChild(rightBtns);
        card.appendChild(head);

        // Sources
        const srcField = el("div",{className:"field"});
        srcField.appendChild(el("div",{className:"label",textContent:"Джерела (колонки поточного листа)"}));
        const srcList = el("div",{style:"display:flex; flex-direction:column; gap:6px"});
        srcField.appendChild(srcList);

        const renderSources = ()=>{
          srcList.innerHTML="";
          if(b.sources.length===0){
            srcList.appendChild(el("div",{className:"muted",textContent:"Джерел немає. Додайте хоча б одну колонку."}));
          }
          b.sources.forEach((s, sidx)=>{
            const row = el("div",{className:"row",style:"align-items:center"});
            const sel = el("select",{className:"input"});
            colOptions().forEach(opt=>{
              const o = el("option",{value:opt.value, textContent:opt.label});
              if(String(s)===String(opt.value)) o.selected = true;
              sel.appendChild(o);
            });
            sel.onchange=()=>{ b.sources[sidx]=Number(sel.value); onDirty(); };
            const rm = el("button",{className:"btn",textContent:"−",title:"Прибрати джерело"});
            rm.onclick=()=>{ b.sources.splice(sidx,1); renderSources(); };
            row.appendChild(sel);
            row.appendChild(rm);
            srcList.appendChild(row);
          });
        };

        const addSrcBtn = el("button",{className:"btn btn-primary",textContent:"＋ Додати джерело"});
        addSrcBtn.onclick=()=>{ b.sources.push(0); renderSources(); };
        srcField.appendChild(addSrcBtn);
        card.appendChild(srcField);
        renderSources();

        // Rule
        const ruleRow = el("div",{className:"row",style:"align-items:center; gap:8px"});
        ruleRow.appendChild(el("div",{className:"label",textContent:"Правило"}));

        const opSel = el("select",{className:"input"});
        const opOpts = [
          {v:"concat", t:"Конкатенація"},
          {v:"seq", t:"Послідовний запис"},
          {v:"newline", t:"З нової строки"},
        ];
        opOpts.forEach(o=>{
          const opt = el("option",{value:o.v, textContent:o.t});
          if(b.op===o.v) opt.selected=true;
          opSel.appendChild(opt);
        });
        opSel.onchange=()=>{ b.op=opSel.value; renderBlocks(); };

        ruleRow.appendChild(opSel);

        // delimiter for concat
        if(b.op==="concat"){
          const delSel = el("select",{className:"input", title:"Розділювач"});
          const dels = [
            {v:" ", t:"(пробіл)"},
            {v:"", t:"(без)"},
            {v:"-", t:"-"},
            {v:"/", t:"/"},
            {v:":", t:":"},
            {v:";", t:";"},
            {v:",", t:","},
            {v:".", t:"."},
          ];
          dels.forEach(d=>{
            const opt = el("option",{value:d.v, textContent:d.t});
            if(String(b.delimiter)===String(d.v)) opt.selected=true;
            delSel.appendChild(opt);
          });
          delSel.onchange=()=>{ b.delimiter = delSel.value; };
          ruleRow.appendChild(el("div",{className:"label",textContent:"Розділювач"}));
          ruleRow.appendChild(delSel);
        } else {
          // no delimiter
        }

        card.appendChild(ruleRow);

        // Note about line separation between blocks (MVP)
        card.appendChild(el("div",{className:"muted",textContent:"Налаштуйте правила між блоками (між Блок 1 і Блок 2, і т.д.). Загальний параметр зверху дозволяє швидко виставити однакове правило для всіх."}));

        blocksWrap.appendChild(card);

        // Join settings between this block and the next one
        if(idx < blocks.length-1){
          const j = joins[idx] || {op: joinAll.op, delimiter: joinAll.op==="concat" ? (joinAll.delimiter||" ") : ""};
          joins[idx] = j;

          const jr = el("div",{className:"row",style:"align-items:center; gap:8px; margin:4px 0 12px 0; padding:8px; border-radius:10px; background:rgba(0,0,0,0.03); flex-wrap:wrap"});
          jr.appendChild(el("div",{className:"muted",textContent:`Між блоком ${idx+1} і ${idx+2}:`}));

          const sel = el("select",{className:"input",style:"min-width:200px"});
          sel.appendChild(el("option",{value:"newline",textContent:"З нової строки"}));
          sel.appendChild(el("option",{value:"concat",textContent:"Конкатенація"}));
          sel.appendChild(el("option",{value:"seq",textContent:"Послідовний запис"}));
          sel.value = j.op || "newline";
          jr.appendChild(sel);

          const delInp = el("input",{className:"input",placeholder:"розділювач",value: (sel.value==="concat") ? (j.delimiter||" ") : " ", style:"width:120px"});
          jr.appendChild(delInp);

          const refresh = ()=>{
            const op = sel.value;
            j.op = op;
            if(op==="concat"){
              delInp.disabled = false;
              j.delimiter = delInp.value ?? " ";
            } else {
              delInp.disabled = true;
              j.delimiter = "";
            }
          };
          sel.onchange = ()=>{ refresh(); onDirty(); };
          delInp.oninput = ()=>{ if(sel.value==="concat"){ j.delimiter = delInp.value; onDirty(); } };
          refresh();

          blocksWrap.appendChild(jr);
        }
      });
    };

    const addBlockBtn = el("button",{className:"btn btn-primary",textContent:"＋ Додати блок"});
    addBlockBtn.onclick=()=>{ blocks.push({sources:[], op:"concat", delimiter:" "}); joins=[]; renderBlocks(); };
    body.appendChild(addBlockBtn);

    renderBlocks();

    const res = await modalOpen({
      title:"Редактор комірки (блоки)",
      bodyNodes:[body],
      actions:[
        btn("Скасувати","cancel","btn"),
        btn("Зберегти","ok","btn btn-primary"),
      ]
    });

    if(res && res.type==="ok"){
      // sanitize and save
      const clean = blocks.map(b=>({
        sources: Array.isArray(b.sources) ? b.sources.map(x=>Number(x)).filter(x=>Number.isFinite(x)) : [],
        op: (b.op==="seq"||b.op==="newline"||b.op==="concat") ? b.op : "concat",
        delimiter: (b.op==="concat") ? String(b.delimiter ?? " ") : ""
      }));
      setCellConfig(t, key, { joinAll, joins, blocks: clean });
      onDirty();
    }
  };

  const renderEditor = ()=>{
    right.innerHTML="";
    const t = getActive();

    if(!t){
      const msg = el("div",{className:"card"});
      msg.appendChild(el("b",{textContent:"Немає шаблону"}));
      msg.appendChild(el("div",{className:"muted",style:"margin-top:6px",textContent:"Створіть шаблон, щоб налаштувати макет та блоки комірок."}));
      right.appendChild(msg);
      return;
    }

    ensureLayout(t);

    right.appendChild(el("h4",{textContent:`Редагування шаблону: ${t.name||t.id}`}));

    // Name
    const nameRow = el("div",{className:"row",style:"align-items:center; gap:8px"});
    nameRow.appendChild(el("div",{className:"muted",textContent:"Назва шаблону"}));
    const nameInput = el("input",{className:"input",value:t.name||"",placeholder:"Напр.: Коротко / Для керівника"});
    nameInput.oninput=()=>{ t.name=nameInput.value; onDirty(); renderList(); };
    nameRow.appendChild(nameInput);
    right.appendChild(nameRow);

    // Top controls
    const ctr = el("div",{className:"card",style:"margin-top:12px"});
    ctr.appendChild(el("b",{textContent:"Конструктор макету (прев’ю)"}));

    const ctrRow = el("div",{className:"row",style:"margin-top:10px; align-items:flex-end"});
    const fCols = el("div",{className:"field", style:"margin:0"});
    fCols.appendChild(el("div",{className:"label",textContent:"Кількість колонок (1–6)"}));
    const inpCols = el("input",{className:"input", type:"number", min:"1", max:"6", value:String(t.layout.cols)});
    fCols.appendChild(inpCols);

    const fRows = el("div",{className:"field", style:"margin:0"});
    fRows.appendChild(el("div",{className:"label",textContent:"Кількість строк (1–6)"}));
    const inpRows = el("input",{className:"input", type:"number", min:"1", max:"6", value:String(t.layout.rows)});
    fRows.appendChild(inpRows);

    const applyBtn = el("button",{className:"btn btn-primary",textContent:"Застосувати макет"});
    const jumpThemeBtn = el("button",{className:"btn",textContent:"Тема ↓", title:"Перейти до налаштувань теми"});
    jumpThemeBtn.onclick=()=>{ const elTheme = document.getElementById('svThemeAnchor'); if(elTheme) elTheme.scrollIntoView({behavior:'smooth', block:'start'}); };
    applyBtn.onclick=()=>{
      const nc = Math.max(1, Math.min(6, Number(inpCols.value||2)));
      const nr = Math.max(1, Math.min(6, Number(inpRows.value||2)));
      const oldRows = t.layout.rows, oldCols = t.layout.cols;
      t.layout.rows = nr; t.layout.cols = nc;
      // keep cells that still fit
      const nextCells = {};
      for(const k of Object.keys(t.layout.cells||{})){
        const [rr,cc] = k.split("-").map(x=>Number(x));
        if(Number.isFinite(rr)&&Number.isFinite(cc) && rr<nr && cc<nc){
          nextCells[`${rr}-${cc}`] = t.layout.cells[k];
        }
      }
      t.layout.cells = nextCells;
      onDirty();
      renderEditor(); // refresh grid
    };

    ctrRow.appendChild(fCols);
    ctrRow.appendChild(fRows);
    ctrRow.appendChild(applyBtn);
    if(typeof jumpThemeBtn!=="undefined") ctrRow.appendChild(jumpThemeBtn);
    ctr.appendChild(ctrRow);

    // Grid preview
    const gridWrap = el("div",{style:"margin-top:12px"});
    const grid = el("div",{className:"sv-grid"});
    grid.style.gridTemplateColumns = `repeat(${t.layout.cols}, minmax(120px, 1fr))`;

    for(let r=0;r<t.layout.rows;r++){
      for(let c=0;c<t.layout.cols;c++){
        const key = `${r}-${c}`;
        const cfg = getCellConfig(t, key);
        const n = cfg.blocks.length;
        const status = n===0 ? "Порожньо" : (n===1 ? "1 блок" : `${n} блоки`);
        const cell = el("div",{className:"sv-cell"});
        cell.appendChild(el("div",{className:"sv-cell-title",textContent:`R${r+1} C${c+1}`}));
        cell.appendChild(el("div",{className:"sv-cell-sub muted",textContent:status}));
        if(n>0) cell.classList.add("sv-cell-filled");
        cell.onclick=async ()=>{ await openCellEditor({t, r, c}); renderEditor(); };
        grid.appendChild(cell);
      }
    }
    gridWrap.appendChild(grid);
    ctr.appendChild(gridWrap);

    ctr.appendChild(el("div",{className:"muted",style:"margin-top:10px",textContent:"Натисніть на комірку макету, щоб налаштувати блоки (джерела + правила поєднання). На цьому кроці це лише структура — рендеринг буде на кроці 4."}));

    right.appendChild(ctr);


// Theme (per sheet)
const themeCard = el("div",{className:"card",style:"margin-top:12px"});
    themeCard.id = "svThemeAnchor";
themeCard.appendChild(el("b",{textContent:"Тема карток (для цього листа)"}));
const th = ensureTheme();

const tRow1 = el("div",{className:"row",style:"margin-top:10px; flex-wrap:wrap; gap:10px; align-items:flex-end"});
// radius
const fRad = el("div",{className:"field",style:"margin:0"});
fRad.appendChild(el("div",{className:"label",textContent:"Заокруглення (px)"}));
const inpRad = el("input",{className:"input", type:"number", min:"0", max:"48", value:String(th.radius)});
inpRad.oninput=()=>{ th.radius = Math.max(0, Math.min(48, Number(inpRad.value||0))); onDirty(); };
fRad.appendChild(inpRad);

// opacity
const fOp = el("div",{className:"field",style:"margin:0"});
fOp.appendChild(el("div",{className:"label",textContent:"Прозорість (0..1)"}));
const inpOp = el("input",{className:"input", type:"number", step:"0.05", min:"0.05", max:"1", value:String(th.cardOpacity)});
inpOp.oninput=()=>{ th.cardOpacity = Math.max(0.05, Math.min(1, Number(inpOp.value||0.92))); onDirty(); };
fOp.appendChild(inpOp);

// card color
const fCol = el("div",{className:"field",style:"margin:0"});
fCol.appendChild(el("div",{className:"label",textContent:"Колір картки"}));
const inpCol = el("input",{className:"input", type:"color", value: th.cardColor || "#ff3b30", style:"width:120px"});
inpCol.oninput=()=>{ th.cardColor = inpCol.value; onDirty(); };
fCol.appendChild(inpCol);

tRow1.appendChild(fRad);
tRow1.appendChild(fOp);
tRow1.appendChild(fCol);
themeCard.appendChild(tRow1);

// Conditional background for cards based on a rule evaluated on a selected column in the row
const condRow = el("div",{className:"row",style:"margin-top:10px; flex-wrap:wrap; gap:12px; align-items:center"});
const cbCond = el("input",{type:"checkbox"}); cbCond.checked = !!th.cardBgRulesEnabled;
cbCond.onchange=()=>{ th.cardBgRulesEnabled = !!cbCond.checked; onDirty(); renderEditor(); };
const lblCond = el("label",{className:"muted",style:"display:flex; gap:8px; align-items:center"});
lblCond.appendChild(cbCond);
lblCond.appendChild(el("span",{textContent:"ФОН карток за правилами (умовне забарвлення)"}));
condRow.appendChild(lblCond);
themeCard.appendChild(condRow);

if(th.cardBgRulesEnabled){
  const rulesWrap = el("div",{style:"margin-top:6px; border:1px dashed #d8d8d8; border-radius:12px; padding:10px"});
  rulesWrap.appendChild(el("div",{className:"muted",textContent:"Порядок важливий: застосовується перше правило, яке спрацювало. Перевірка робиться по значенню вибраної колонки в рядку."}));

  const rules = Array.isArray(th.cardBgRules) ? th.cardBgRules : (th.cardBgRules=[]);

  const rowBtns = el("div",{className:"row",style:"justify-content:space-between; align-items:center; margin-top:8px"});
  const addRuleBtn = el("button",{className:"btn btn-primary",textContent:"＋ Додати правило"});
  rowBtns.appendChild(addRuleBtn);
  rulesWrap.appendChild(rowBtns);

  const listRules = el("div",{style:"margin-top:8px"});
  rulesWrap.appendChild(listRules);

  const tests = [
    {v:"notempty", t:"Є значення (не порожньо)"},
    {v:"empty", t:"Порожньо"},
    {v:"isnumber", t:"Тип: число"},
    {v:"isdate", t:"Тип: дата (ДД.ММ.РР)"},
    {v:"equals", t:"Значення дорівнює…"},
    {v:"contains", t:"Значення містить…"}
  ];

  const renderRules = ()=>{
    listRules.innerHTML="";
    if(!rules.length){
      listRules.appendChild(el("div",{className:"muted",textContent:"Немає правил. Натисніть “＋ Додати правило”."}));
      return;
    }
    rules.forEach((r, idx)=>{
      if(!r || typeof r!=="object") r = rules[idx] = {};
      if(r.col===undefined || r.col===null) r.col = "0";
      if(typeof r.test!=="string") r.test = "notempty";
      if(typeof r.value!=="string") r.value = "";
      if(typeof r.color!=="string") r.color = "#ff3b30";

      const box = el("div",{style:"border:1px solid #e6e6e6; border-radius:12px; padding:10px; margin:8px 0"});
      const head = el("div",{className:"row",style:"justify-content:space-between; align-items:center"});
      head.appendChild(el("b",{textContent:`Правило ${idx+1}`}));
      const del = el("button",{className:"btn danger",textContent:"Видалити"});
      del.onclick=()=>{ rules.splice(idx,1); onDirty(); renderEditor(); };
      head.appendChild(del);
      box.appendChild(head);

      const line = el("div",{className:"row",style:"margin-top:8px; gap:10px; flex-wrap:wrap; align-items:flex-end"});
      // column
      const fCol2 = el("div",{className:"field",style:"margin:0"});
      fCol2.appendChild(el("div",{className:"label",textContent:"Комірка (колонка)"}));
      const selCol = el("select",{className:"input",style:"min-width:260px"});
      colOptions().forEach(o=> selCol.appendChild(el("option",{value:o.value,textContent:o.label})));
      selCol.value = String(r.col);
      selCol.onchange=()=>{ r.col = selCol.value; onDirty(); };
      fCol2.appendChild(selCol);

      // test
      const fTest = el("div",{className:"field",style:"margin:0"});
      fTest.appendChild(el("div",{className:"label",textContent:"Умова"}));
      const selTest = el("select",{className:"input",style:"min-width:220px"});
      tests.forEach(t=> selTest.appendChild(el("option",{value:t.v,textContent:t.t})));
      selTest.value = r.test;
      fTest.appendChild(selTest);

      const fVal = el("div",{className:"field",style:"margin:0"});
      fVal.appendChild(el("div",{className:"label",textContent:"Значення"}));
      const inpVal = el("input",{className:"input",style:"min-width:180px", value:r.value});
      inpVal.oninput=()=>{ r.value = inpVal.value; onDirty(); };
      fVal.appendChild(inpVal);

      const fColor2 = el("div",{className:"field",style:"margin:0"});
      fColor2.appendChild(el("div",{className:"label",textContent:"Колір"}));
      const inpColor2 = el("input",{className:"input", type:"color", value: r.color.startsWith("#")?r.color:"#ff3b30", style:"width:120px"});
      inpColor2.oninput=()=>{ r.color = inpColor2.value; onDirty(); };
      fColor2.appendChild(inpColor2);

      selTest.onchange=()=>{
        r.test = selTest.value;
        const needs = (r.test==="equals"||r.test==="contains");
        inpVal.disabled = !needs;
        if(!needs) r.value = "";
        onDirty();
      };
      // init disabled
      selTest.onchange();

      line.appendChild(fCol2);
      line.appendChild(fTest);
      line.appendChild(fVal);
      line.appendChild(fColor2);

      box.appendChild(line);

      // ordering
      const ord = el("div",{className:"row",style:"justify-content:flex-end; gap:8px; margin-top:8px"});
      const up = el("button",{className:"btn",textContent:"↑"});
      const dn = el("button",{className:"btn",textContent:"↓"});
      up.onclick=()=>{ if(idx<=0) return; const tmp=rules[idx-1]; rules[idx-1]=rules[idx]; rules[idx]=tmp; onDirty(); renderEditor(); };
      dn.onclick=()=>{ if(idx>=rules.length-1) return; const tmp=rules[idx+1]; rules[idx+1]=rules[idx]; rules[idx]=tmp; onDirty(); renderEditor(); };
      ord.appendChild(up); ord.appendChild(dn);
      box.appendChild(ord);

      listRules.appendChild(box);
    });
  };

  addRuleBtn.onclick=()=>{
    rules.push({col:"0", test:"notempty", value:"", color:"#ff3b30"});
    onDirty();
    renderEditor();
  };

  renderRules();
  themeCard.appendChild(rulesWrap);
}


const tRow2 = el("div",{className:"row",style:"margin-top:8px; flex-wrap:wrap; gap:14px; align-items:center"});
const cbBorders = el("input",{type:"checkbox"}); cbBorders.checked = !!th.showBorders;
cbBorders.onchange=()=>{ th.showBorders = !!cbBorders.checked; onDirty(); };
const lblBorders = el("label",{className:"muted",style:"display:flex; gap:8px; align-items:center"});
lblBorders.appendChild(cbBorders); lblBorders.appendChild(el("span",{textContent:"Відображати межі комірок"}));
tRow2.appendChild(lblBorders);

const cbGlass = el("input",{type:"checkbox"}); cbGlass.checked = !!th.glass;
cbGlass.onchange=()=>{ th.glass = !!cbGlass.checked; onDirty(); };
const lblGlass = el("label",{className:"muted",style:"display:flex; gap:8px; align-items:center"});
lblGlass.appendChild(cbGlass); lblGlass.appendChild(el("span",{textContent:"Liquid Glass (blur)"}));
tRow2.appendChild(lblGlass);

const cbGrad = el("input",{type:"checkbox"}); cbGrad.checked = !!th.gradient;
cbGrad.onchange=()=>{ th.gradient = !!cbGrad.checked; onDirty(); renderEditor(); };
const lblGrad = el("label",{className:"muted",style:"display:flex; gap:8px; align-items:center"});
lblGrad.appendChild(cbGrad); lblGrad.appendChild(el("span",{textContent:"Градієнт"}));
tRow2.appendChild(lblGrad);

themeCard.appendChild(tRow2);

if(th.gradient){
  const gRow = el("div",{className:"row",style:"margin-top:8px; flex-wrap:wrap; gap:10px; align-items:flex-end"});
  const fG1 = el("div",{className:"field",style:"margin:0"});
  fG1.appendChild(el("div",{className:"label",textContent:"Градієнт: from"}));
  const inpG1 = el("input",{className:"input", type:"color", value: th.gradFrom || "#ff3b30", style:"width:120px"});
  inpG1.oninput=()=>{ th.gradFrom = inpG1.value; onDirty(); };
  fG1.appendChild(inpG1);

  const fG2 = el("div",{className:"field",style:"margin:0"});
  fG2.appendChild(el("div",{className:"label",textContent:"Градієнт: to"}));
  const inpG2 = el("input",{className:"input", type:"color", value: th.gradTo || "#ff9500", style:"width:120px"});
  inpG2.oninput=()=>{ th.gradTo = inpG2.value; onDirty(); };
  fG2.appendChild(inpG2);

  gRow.appendChild(fG1); gRow.appendChild(fG2);
  themeCard.appendChild(gRow);
}

const bgRow = el("div",{className:"row",style:"margin-top:8px; flex-wrap:wrap; gap:10px; align-items:flex-end"});
const fBg = el("div",{className:"field",style:"margin:0"});
fBg.appendChild(el("div",{className:"label",textContent:"Колір фону (лист)"}));
const inpBg = el("input",{className:"input", type:"color", value: (th.bgColor && th.bgColor.startsWith("#")) ? th.bgColor : "#ffffff", style:"width:120px"});
const cbBg = el("input",{type:"checkbox"}); cbBg.checked = !!th.bgColor;
const lblBg = el("label",{className:"muted",style:"display:flex; gap:8px; align-items:center"});
lblBg.appendChild(cbBg); lblBg.appendChild(el("span",{textContent:"увімкнути фон"}));
cbBg.onchange=()=>{ th.bgColor = cbBg.checked ? inpBg.value : ""; onDirty(); };
inpBg.oninput=()=>{ if(cbBg.checked){ th.bgColor = inpBg.value; onDirty(); } };
fBg.appendChild(inpBg);
bgRow.appendChild(fBg);
bgRow.appendChild(lblBg);
themeCard.appendChild(bgRow);

const cssRow = el("div",{style:"margin-top:10px"});
cssRow.appendChild(el("div",{className:"label",textContent:"Custom CSS (опційно)"}));
const ta = el("textarea",{className:"input",style:"width:100%; min-height:70px", placeholder:".sv-card{...}", value: th.customCss||""});
ta.oninput=()=>{ th.customCss = ta.value; onDirty(); };
cssRow.appendChild(ta);
themeCard.appendChild(cssRow);

right.appendChild(themeCard);

    // Delete template
    const dangerRow = el("div",{className:"row",style:"justify-content:space-between; margin-top:12px; align-items:center"});
    const delBtn = el("button",{className:"btn danger",textContent:"🗑 Видалити шаблон"});
    delBtn.onclick=()=>{
      const ok = confirm(`Видалити шаблон "${t.name||t.id}"?`);
      if(!ok) return;
      const idx = templates.findIndex(x=>x.id===t.id);
      if(idx>=0) templates.splice(idx,1);
      if(templates.length===0){
        sheet.simplified.activeTemplateId=null;
        activeId=null;
      } else if(!templates.some(x=>x.id===activeId)){
        activeId = templates[0].id;
        sheet.simplified.activeTemplateId=activeId;
      }
      onDirty();
      render();
    };
    dangerRow.appendChild(el("div",{className:"muted",textContent:""}));
    dangerRow.appendChild(delBtn);
    right.appendChild(dangerRow);
  };

  const render = ()=>{ renderList(); renderEditor(); };

  addBtn.onclick=()=>{
    const name = prompt("Назва шаблону спрощеного перегляду:", "Шаблон");
    if(name===null) return;
    const nm = name.trim();
    if(!nm) return alert("Назва не може бути порожньою.");
    const id = `sv_${Date.now().toString(36)}_${Math.floor(Math.random()*999)}`;
    templates.push({id, name:nm, layout:{rows:2, cols:2, cells:{}}, version:1});
    activeId=id;
    sheet.simplified.activeTemplateId=activeId;
    onDirty();
    render();
  };

  render();

  try{
    await modalOpen({
      title:`Спрощений перегляд (конструктор): ${sheet.title}`,
      bodyNodes:[wrap],
      actions:[ btn("Закрити","cancel","btn") ]
    });
  } finally {
    if(modalEl) modalEl.classList.remove("modal-wide");
  }
}


export function buildSettingsUI({tab, sheets, settings, addFieldsCfg, uiPrefs, onDirty}){
  const root = document.createElement("div");
  // wrapper block to keep function scope consistent
  {

  // --- UI preferences (applies immediately) ---
  {
    const card = el("div",{className:"card"});
    card.appendChild(el("h4",{textContent:"Інтерфейс"}));
    card.appendChild(el("div",{className:"muted",textContent:"Налаштування масштабу кнопок застосовується одразу та зберігається для цього пристрою/браузера."}));

    const row = el("div",{className:"row",style:"margin-top:10px; align-items:center; gap:10px; flex-wrap:wrap"});
    row.appendChild(el("div",{className:"muted",textContent:"Розмір кнопок:"}));

    const value = el("span",{className:"pill",textContent:"100%"});
    const range = el("input",{type:"range",min:"0.80",max:"1.60",step:"0.05"});
    range.style.flex = "1";
    range.style.minWidth = "240px";

    const init = Number(uiPrefs?.btnScale);
    const scale = (Number.isFinite(init) && init>=0.6 && init<=2.2) ? init : 1;
    range.value = String(scale);
    value.textContent = `${Math.round(scale*100)}%`;

    const apply = async (s)=>{
      const v = (Number.isFinite(s) && s>=0.6 && s<=2.2) ? s : 1;
      document.documentElement.style.setProperty("--btn-scale", String(v));
      value.textContent = `${Math.round(v*100)}%`;
      try{ await cfgSet("ui_btn_scale", v); }catch{ /* ignore */ }
      if(uiPrefs) uiPrefs.btnScale = v;
    };

    range.addEventListener("input", ()=>{
      const s = Number(range.value);
      const v = (Number.isFinite(s) && s>=0.6 && s<=2.2) ? s : 1;
      document.documentElement.style.setProperty("--btn-scale", String(v));
      value.textContent = `${Math.round(v*100)}%`;
    });
    range.addEventListener("change", ()=>apply(Number(range.value)));

    row.appendChild(range);
    row.appendChild(value);
    card.appendChild(row);

    root.appendChild(card);
  }

  if(tab==="sheets"){
    const card = el("div",{className:"card"});
    const head = el("div",{className:"row",style:"justify-content:space-between; align-items:center; gap:8px; flex-wrap:wrap"});
    head.appendChild(el("h4",{textContent:"Листи (стандартні та користувацькі)"}));
    const bk = el("button",{className:"btn",textContent:"💾/📥 Бекап вкладки"});
    bk.onclick=()=>openTabBackupDialog("sheets");
    head.appendChild(bk);
    card.appendChild(head);
    card.appendChild(el("div",{className:"muted",textContent:"Тут можна перейменовувати будь-які листи. Видаляти — лише користувацькі. Структуру колонок налаштовуйте у вкладці “Колонки”."}));
    const btnAdd = el("button",{className:"btn btn-primary",textContent:"＋ Додати лист"});
    btnAdd.onclick=async()=>{
      const wrap = el("div",{className:"vstack",style:"gap:10px; min-width: 320px;"});

      const row1 = el("div",{className:"row",style:"gap:8px; flex-wrap:wrap; align-items:center"});
      row1.appendChild(el("label",{className:"muted",textContent:"Назва листа:"}));
      const inpTitle = el("input",{className:"input",type:"text",placeholder:"Напр., Журнал договорів"});
      inpTitle.style.flex="1";
      row1.appendChild(inpTitle);

      const row2 = el("div",{className:"row",style:"gap:8px; flex-wrap:wrap; align-items:center"});
      row2.appendChild(el("label",{className:"muted",textContent:"Кількість колонок:"}));
      const inpCount = el("input",{className:"input",type:"number",min:"1",max:"60",value:"3",style:"width:120px"});
      row2.appendChild(inpCount);

      const colsWrap = el("div",{className:"vstack",style:"gap:6px"});
      const colsHead = el("div",{className:"muted",textContent:"Назви колонок (обовʼязково):"});
      colsWrap.appendChild(colsHead);
      const colsInputs = el("div",{className:"vstack",style:"gap:6px"});

      const rebuildCols = ()=>{
        colsInputs.innerHTML="";
        let n = parseInt(inpCount.value,10);
        if(!Number.isFinite(n) || n<1) n=1;
        if(n>60) n=60;
        inpCount.value=String(n);
        for(let i=0;i<n;i++){
          const r = el("div",{className:"row",style:"gap:8px; align-items:center"});
          r.appendChild(el("span",{className:"muted",textContent:`${i+1}.`,style:"width:22px; text-align:right"}));
          const inp = el("input",{className:"input",type:"text",placeholder:`Колонка ${i+1}`});
          inp.dataset.idx=String(i);
          inp.style.flex="1";
          colsInputs.appendChild(r);
          r.appendChild(inp);
        }
      };

      inpCount.addEventListener("input", rebuildCols);
      rebuildCols();

      colsWrap.appendChild(colsInputs);

      wrap.appendChild(row1);
      wrap.appendChild(row2);
      wrap.appendChild(colsWrap);

      // open modal
      const res = await modalOpen({
        title:"Додати лист",
        bodyNodes:[wrap],
        actions:[ btn("Скасувати","cancel","btn"), btn("Додати","ok","btn btn-primary") ]
      });
      if(res.type!=="ok") return;

      const t = (inpTitle.value||"").trim();
      if(!t) return alert("Назва листа не може бути порожньою.");

      const cols = Array.from(colsInputs.querySelectorAll("input"))
        .map(i=>String(i.value||"").trim())
        .filter(Boolean);

      if(cols.length<1) return alert("Потрібна хоча б 1 колонка.");
      const nExpected = parseInt(inpCount.value,10);
      if(cols.length !== nExpected){
        return alert("Заповніть назви всіх колонок або зменшіть їх кількість.");
      }

      const key = genKeyFromTitle(t);
      const sheet = {
        key,
        title: t,
        orderColumn: null,
        subrows: null,
        columns: cols.map((name)=>({name, type:"text", required:false, subrows:true})),
        addFields: cols.slice(0, Math.min(cols.length, 8)),
        export:{ pageSize:"A4", orientation:"portrait", exportHiddenCols:[], rowFilters:[] },
        simplified: { enabled:false, on:false, activeTemplateId:null, templates:[], theme:{} }
      };
      sheets.push(sheet);
      onDirty();
    };
    card.appendChild(el("div",{className:"row"})).appendChild(btnAdd);
    if(!sheets.length) card.appendChild(el("div",{className:"muted",textContent:"Немає листів."}));
    else for(const sh of sheets){
      const c=el("div",{className:"card"});
      c.appendChild(el("div",{innerHTML:`<b>${sh.title}</b> <span class="pill">${sh.key}</span>`}));
      const row=el("div",{className:"row"});

      const btnRename=el("button",{className:"btn",textContent:"✏️ Перейменувати"});
      btnRename.onclick=()=>{
        const title = prompt("Нова назва листа:", sh.title||"");
        if(title===null) return;
        const t=title.trim(); if(!t) return alert("Назва не може бути порожньою.");
        sh.title=t; onDirty();
      };
      row.appendChild(btnRename);

      // Default sort per sheet (per-sheet settings)
      const sortBlock = el("div",{className:"row",style:"flex-wrap:wrap; gap:8px; align-items:center"});
      sortBlock.appendChild(el("div",{className:"muted",textContent:"Сортування за замовчуванням:"}));

      const sortCol = el("select",{className:"select"});
      sortCol.appendChild(el("option",{value:"",textContent:"— Не сортувати"}));
      (sh.columns||[]).forEach((col, idx)=>{
        sortCol.appendChild(el("option",{value:col.name,textContent:`${idx+1}. ${col.name}`}));
      });
      const curSortCol = sh.defaultSort?.col || "";
      sortCol.value = curSortCol;

      const sortDir = el("select",{className:"select"});
      sortDir.appendChild(el("option",{value:"asc",textContent:"↑ від меншого до більшого"}));
      sortDir.appendChild(el("option",{value:"desc",textContent:"↓ від більшого до меншого"}));
      sortDir.value = (sh.defaultSort?.dir==="desc") ? "desc" : "asc";

      const applySortDefaults = ()=>{
        const colName = sortCol.value || null;
        if(!colName){
          delete sh.defaultSort;
        } else {
          sh.defaultSort = { col: colName, dir: sortDir.value };
        }
        onDirty();
      };
      sortCol.onchange=applySortDefaults;
      sortDir.onchange=applySortDefaults;

      sortBlock.appendChild(sortCol);
      sortBlock.appendChild(sortDir);
      c.appendChild(sortBlock);


      // Simplified view per sheet (Stage 1)
      if(!sh.simplified) sh.simplified = { enabled:false, on:false, activeTemplateId:null, templates:[] };
      const simpRow = el("div",{className:"row",style:"flex-wrap:wrap; gap:8px; align-items:center"});

      const cb = el("input",{type:"checkbox"});
      cb.checked = !!sh.simplified.enabled;

      const btnCfg = el("button",{className:"btn",textContent:"⚙ Налаштувати спрощений перегляд"});
      btnCfg.onclick=async ()=>{
        await openSimplifiedTemplatesEditor({sheet:sh, onDirty});
      };
      // show/hide dynamically when checkbox toggles
      btnCfg.style.display = sh.simplified.enabled ? "" : "none";

      cb.onchange=()=>{
        sh.simplified.enabled = !!cb.checked;
        if(sh.simplified.enabled){
          if(!Array.isArray(sh.simplified.templates)) sh.simplified.templates=[];
          // No default template here: user creates templates explicitly in the constructor.
          btnCfg.style.display = "";
        } else {
          sh.simplified.on = false;
          btnCfg.style.display = "none";
        }
        onDirty();
      };

      const lbl = el("label",{className:"muted",style:"display:flex;gap:8px;align-items:center"});
      lbl.appendChild(cb);
      lbl.appendChild(el("span",{textContent:"Дозволити спрощений перегляд для цього листа"}));
      simpRow.appendChild(lbl);
      simpRow.appendChild(btnCfg);

      c.appendChild(simpRow);

      if(sh.key.startsWith("custom_")){
        const btnDel=el("button",{className:"btn",textContent:"🗑 Видалити"});
        btnDel.onclick=()=>{
          const ok=confirm(`Видалити лист "${sh.title}" разом з усіма даними?`);
          if(!ok) return;
          const idx=sheets.findIndex(x=>x.key===sh.key);
          if(idx>=0){ sheets.splice(idx,1); onDirty(); }
        };
        row.appendChild(btnDel);
      }
      c.appendChild(row);
      card.appendChild(c);
    }
    root.appendChild(card); 
  }

  if(tab==="columns"){
    const card=el("div",{className:"card"});
    const head = el("div",{className:"row",style:"justify-content:space-between; align-items:center; gap:8px; flex-wrap:wrap"});
    head.appendChild(el("h4",{textContent:"Колонки: порядок (drag&drop / ↑↓), видимість, тип, обовʼязковість, підстроки"}));
    const bk = el("button",{className:"btn",textContent:"💾/📥 Бекап вкладки"});
    bk.onclick=()=>openTabBackupDialog("columns");
    head.appendChild(bk);
    card.appendChild(head);
    const pick = el("select",{className:"select"});
    for(const sh of sheets) pick.appendChild(el("option",{value:sh.key,textContent:sh.title}));
    card.appendChild(el("div",{className:"row"})).appendChild(pick);
    const area=el("div",{style:"margin-top:10px"}); card.appendChild(area);

    const render = ()=>{
      area.innerHTML="";
      const key=pick.value;
      const sh=sheets.find(s=>s.key===key); if(!sh) return;

      const sCfg = settings[key] || (settings[key]={ hiddenCols:[], widths:{}, export: sh.export||{pageSize:"A4",orientation:"portrait",exportHiddenCols:[],rowFilters:[]} });
      const hidden = new Set(sCfg.hiddenCols||[]);

      const list = dndList({
        items: sh.columns,
        onReorder: ()=>{ onDirty(); },
        renderItem: (col, idx)=>{
          const left=el("div",{});
          left.appendChild(el("div",{innerHTML:`<b>${col.name}</b> <span class="pill">drag</span> <span class="pill">${idx+1}</span>`}));
          left.appendChild(el("div",{className:"muted",textContent:`type: ${col.type} | required: ${col.required?"так":"ні"} | subrows: ${col.subrows?"так":"ні"}`}));

          const controls=el("div",{className:"row"});
          const up=el("button",{className:"btn",textContent:"↑"});
          const down=el("button",{className:"btn",textContent:"↓"});
          up.onclick=()=>{ if(idx===0) return; const t=sh.columns[idx-1]; sh.columns[idx-1]=sh.columns[idx]; sh.columns[idx]=t; onDirty(); render(); };
          down.onclick=()=>{ if(idx===sh.columns.length-1) return; const t=sh.columns[idx+1]; sh.columns[idx+1]=sh.columns[idx]; sh.columns[idx]=t; onDirty(); render(); };

          const vis=el("button",{className:"btn",textContent:hidden.has(col.name)?"👁 Показати":"🙈 Сховати"});
          vis.onclick=()=>{ if(hidden.has(col.name)) hidden.delete(col.name); else hidden.add(col.name); sCfg.hiddenCols=Array.from(hidden); onDirty(); render(); };

          const req=el("button",{className:"btn",textContent:col.required?"★ Обовʼязк.":"☆ Необовʼязк."});
          req.onclick=()=>{ col.required=!col.required; onDirty(); render(); };

          const sub=el("button",{className:"btn",textContent:col.subrows?"↵ Підстроки":"— Без підстрок"});
          sub.onclick=()=>{ col.subrows=!col.subrows; onDirty(); render(); };

          const type=el("select",{className:"select"});
          ["text","int","date"].forEach(v=>type.appendChild(el("option",{value:v,textContent:v})));
          type.value=col.type||"text";
          type.onchange=()=>{ col.type=type.value; onDirty(); render(); };

          controls.appendChild(up); controls.appendChild(down); controls.appendChild(vis); controls.appendChild(req); controls.appendChild(sub); controls.appendChild(type);

          const del=el("button",{className:"btn",textContent:"🗑 Колонку"});
          del.onclick=()=>{
            if(sh.columns.length<=1) return alert("Потрібна хоча б 1 колонка.");
            const ok=confirm(`Видалити колонку "${col.name}"?`);
            if(!ok) return;
            sh.columns.splice(idx,1);
            onDirty();
            render();
          };
          controls.appendChild(del);

          const row=el("div",{className:"row",style:"width:100%;justify-content:space-between"});
          row.appendChild(left);
          row.appendChild(controls);
          return row;
        }
      });

      area.appendChild(list);

      const addRow=el("div",{className:"row",style:"margin-top:10px"});
      const btnAdd=el("button",{className:"btn btn-primary",textContent:"＋ Додати колонку"});
      btnAdd.onclick=()=>{
        const name=prompt("Назва колонки:");
        if(name===null) return;
        const n=name.trim();
        if(!n) return;
        if(sh.columns.some(c=>c.name===n)) return alert("Така колонка вже існує.");
        sh.columns.push({name:n,type:"text",required:false,subrows:true});
        onDirty();
        render();
      };
      addRow.appendChild(btnAdd);
      area.appendChild(addRow);
    };

    pick.onchange=render; render();
    root.appendChild(card); 
  }

  if(tab==="addform"){
    const card=el("div",{className:"card"});
    const head = el("div",{className:"row",style:"justify-content:space-between; align-items:center; gap:8px; flex-wrap:wrap"});
    head.appendChild(el("h4",{textContent:"Які поля вводяться при натисканні “Додати”"}));
    const bk = el("button",{className:"btn",textContent:"💾/📥 Бекап вкладки"});
    bk.onclick=()=>openTabBackupDialog("addform");
    head.appendChild(bk);
    card.appendChild(head);
    const pick = el("select",{className:"select"});
    for(const sh of sheets) pick.appendChild(el("option",{value:sh.key,textContent:sh.title}));
    card.appendChild(el("div",{className:"row"})).appendChild(pick);
    const area=el("div",{style:"margin-top:10px"}); card.appendChild(area);

    const render=()=>{
      area.innerHTML="";
      const sh=sheets.find(s=>s.key===pick.value); if(!sh) return;
      const cfg = addFieldsCfg[sh.key] || (addFieldsCfg[sh.key]=(sh.addFields||sh.columns.map(c=>c.name)).slice());
      const inCfg=new Set(cfg);
      sh.columns.forEach((col)=>{
        const row=el("div",{className:"row",style:"justify-content:space-between; border:1px solid #e6e6e6; padding:8px; border-radius:12px; margin:6px 0;"});
        row.appendChild(el("div",{innerHTML:`<b>${col.name}</b> ${col.subrows?'<span class="pill">↵</span>':''}`}));
        const controls=el("div",{className:"row"});
        const chk=el("input",{type:"checkbox"}); chk.checked=inCfg.has(col.name);
        chk.onchange=()=>{ 
          if(chk.checked) { if(!cfg.includes(col.name)) cfg.push(col.name); } 
          else {const i=cfg.indexOf(col.name); if(i>=0) cfg.splice(i,1);} 
          onDirty(); render(); 
        };
        controls.appendChild(chk);
        if(inCfg.has(col.name)){
          const up=el("button",{className:"btn",textContent:"↑"});
          const dn=el("button",{className:"btn",textContent:"↓"});
          up.onclick=()=>{ const i=cfg.indexOf(col.name); if(i<=0) return; [cfg[i-1],cfg[i]]=[cfg[i],cfg[i-1]]; onDirty(); render(); };
          dn.onclick=()=>{ const i=cfg.indexOf(col.name); if(i<0||i>=cfg.length-1) return; [cfg[i+1],cfg[i]]=[cfg[i],cfg[i+1]]; onDirty(); render(); };
          controls.appendChild(up); controls.appendChild(dn);
        }
        row.appendChild(controls);
        area.appendChild(row);
      });
    };

    pick.onchange=render; render();
    root.appendChild(card); 
  }

  if(tab==="transfer"){
    const card=el("div",{className:"card"});
    const head = el("div",{className:"row",style:"justify-content:space-between; align-items:center; gap:8px; flex-wrap:wrap"});
    head.appendChild(el("h4",{textContent:"Перенесення: шаблони та маршрути"}));
    const bk = el("button",{className:"btn",textContent:"💾/📥 Бекап вкладки"});
    bk.onclick=()=>openTabBackupDialog("transfer");
    head.appendChild(bk);
    card.appendChild(head);
    card.appendChild(el("div",{className:"muted",textContent:"Створюйте шаблони перенесення. Кожен шаблон містить один або кілька маршрутів: (колонка 1 + колонка 2 + ...) → (цільова колонка)."}));
    const area=el("div",{style:"margin-top:10px"});
    card.appendChild(area);
    // async UI
    buildTransferUI({sheets, onDirty}).then(ui=>{ area.appendChild(ui); });
    root.appendChild(card);
  }

  if(tab==="backup"){
    const card = el("div",{className:"card"});
    card.appendChild(el("h4",{textContent:"Бекап"}));
    card.appendChild(el("div",{className:"muted",textContent:
      "Можна робити: 1) загальний бекап налаштувань, 2) загальний бекап перенесень, 3) вибіркові бекапи (по вкладках або “все одразу” із вибором параметрів)."}));

    const row0 = el("div",{className:"row",style:"flex-wrap:wrap; gap:8px; margin-top:10px"});
    const btnPartialAll = el("button",{className:"btn",textContent:"🧩 Вибірковий бекап всіх налаштувань (з вибором)"});
    btnPartialAll.onclick = ()=>openTabBackupDialog("all");
    row0.appendChild(btnPartialAll);
    card.appendChild(row0);

    const row1 = el("div",{className:"row",style:"flex-wrap:wrap; gap:8px; margin-top:10px"});
    const btnExpSettings = el("button",{className:"btn btn-primary",textContent:"💾 Експорт налаштувань"});
    btnExpSettings.onclick = async()=>{ try{ await BackupCore.exportSettingsBackup(); }catch(e){ alert(e?.message||String(e)); } };
    const btnImpSettings = el("button",{className:"btn",textContent:"📥 Імпорт налаштувань"});
    btnImpSettings.onclick = ()=>{ try{ BackupCore.pickAndImportSettingsBackup(); }catch(e){ alert(e?.message||String(e)); } };
    row1.appendChild(btnExpSettings);
    row1.appendChild(btnImpSettings);
    card.appendChild(row1);

    const row2 = el("div",{className:"row",style:"flex-wrap:wrap; gap:8px; margin-top:8px"});
    const btnExpRules = el("button",{className:"btn btn-primary",textContent:"💾 Експорт правил перенесення"});
    btnExpRules.onclick = async()=>{ try{ await BackupCore.exportTransferRulesBackup(); }catch(e){ alert(e?.message||String(e)); } };
    const btnImpRules = el("button",{className:"btn",textContent:"📥 Імпорт правил перенесення"});
    btnImpRules.onclick = ()=>{ try{ BackupCore.pickAndImportTransferRulesBackup(); }catch(e){ alert(e?.message||String(e)); } };
    row2.appendChild(btnExpRules);
    row2.appendChild(btnImpRules);
    card.appendChild(row2);

    const warn = el("div",{className:"card",style:"margin-top:12px"});
    warn.appendChild(el("b",{textContent:"Увага"}));
    warn.appendChild(el("div",{className:"muted",style:"margin-top:6px",textContent:
      "Імпорт замінює відповідний набір конфігурацій у поточній базі. Дані рядків (журнали/справи) не змінюються."}));
    card.appendChild(warn);

    root.appendChild(card);
  }


if(tab==="export"){
    const card=el("div",{className:"card"});
    const head = el("div",{className:"row",style:"justify-content:space-between; align-items:center; gap:8px; flex-wrap:wrap"});
    head.appendChild(el("h4",{textContent:"Експортні профілі на лист"}));
    const bk = el("button",{className:"btn",textContent:"💾/📥 Бекап вкладки"});
    bk.onclick=()=>openTabBackupDialog("export");
    head.appendChild(bk);
    card.appendChild(head);
    card.appendChild(el("div",{className:"muted",textContent:"Тут задається: pageSize, orientation, приховані колонки тільки для експорту, фільтр рядків при експорті."}));

    const pick = el("select",{className:"select"});
    for(const sh of sheets) pick.appendChild(el("option",{value:sh.key,textContent:sh.title}));
    card.appendChild(el("div",{className:"row"})).appendChild(pick);

    const area=el("div",{style:"margin-top:10px"});
    card.appendChild(area);

    const render=()=>{
      area.innerHTML="";
      const sh=sheets.find(s=>s.key===pick.value); if(!sh) return;
      const sCfg = settings[sh.key] || (settings[sh.key]={ hiddenCols:[], widths:{}, export: sh.export||{pageSize:"A4",orientation:"portrait",exportHiddenCols:[],rowFilters:[]} });
      sCfg.export = sCfg.export || {pageSize:"A4",orientation:"portrait",exportHiddenCols:[],rowFilters:[]};

      // pageSize/orientation
      const row1=el("div",{className:"row"});
      const ps=el("select",{className:"select"});
      ["A4","A3"].forEach(v=>ps.appendChild(el("option",{value:v,textContent:v})));
      ps.value=sCfg.export.pageSize||"A4";
      ps.onchange=()=>{ sCfg.export.pageSize=ps.value; onDirty(); };

      const ori=el("select",{className:"select"});
      ["portrait","landscape"].forEach(v=>ori.appendChild(el("option",{value:v,textContent:v})));
      ori.value=sCfg.export.orientation||"portrait";
      ori.onchange=()=>{ sCfg.export.orientation=ori.value; onDirty(); };

      row1.appendChild(el("div",{className:"muted",textContent:"pageSize:"})); row1.appendChild(ps);
      row1.appendChild(el("div",{className:"muted",textContent:"orientation:"})); row1.appendChild(ori);
      area.appendChild(row1);

      // export-only hidden cols
      const hiddenSet=new Set(sCfg.export.exportHiddenCols||[]);
      const colsCard=el("div",{className:"card"});
      colsCard.appendChild(el("div",{innerHTML:"<b>Колонки приховані тільки при експорті</b>"}));
      sh.columns.forEach(col=>{
        const r=el("div",{className:"row",style:"justify-content:space-between; padding:6px 0;"});
        r.appendChild(el("div",{textContent:col.name}));
        const chk=el("input",{type:"checkbox"}); chk.checked=hiddenSet.has(col.name);
        chk.onchange=()=>{ if(chk.checked) hiddenSet.add(col.name); else hiddenSet.delete(col.name); sCfg.export.exportHiddenCols=Array.from(hiddenSet); onDirty(); };
        r.appendChild(chk);
        colsCard.appendChild(r);
      });
      area.appendChild(colsCard);

      // rowFilters
      const filters = sCfg.export.rowFilters || (sCfg.export.rowFilters=[]);
      const fCard=el("div",{className:"card"});
      fCard.appendChild(el("div",{innerHTML:"<b>Фільтр рядків при експорті</b> <span class='pill'>поки простий</span>"}));
      fCard.appendChild(el("div",{className:"muted",textContent:"Якщо фільтри задані — експортуються тільки рядки, що проходять умови."}));

      const renderFilters=()=>{
        // clear previous filter UIs except header
        while(fCard.children.length>2) fCard.removeChild(fCard.lastChild);
        if(!filters.length) fCard.appendChild(el("div",{className:"muted",textContent:"Немає фільтрів."}));
        filters.forEach((f,idx)=>{
          const row=el("div",{className:"row"});
          const col=el("select",{className:"select"});
          sh.columns.forEach(c=>col.appendChild(el("option",{value:c.name,textContent:c.name})));
          col.value=f.col||sh.columns[0]?.name||"";
          col.onchange=()=>{ f.col=col.value; onDirty(); };

          const op=el("select",{className:"select"});
          ["contains","equals","not_contains","not_equals"].forEach(v=>op.appendChild(el("option",{value:v,textContent:v})));
          op.value=f.op||"contains";
          op.onchange=()=>{ f.op=op.value; onDirty(); };

          const val=el("input",{className:"input",value:f.value||"",placeholder:"значення"});
          val.oninput=()=>{ f.value=val.value; onDirty(); };

          const del=el("button",{className:"btn",textContent:"🗑"});
          del.onclick=()=>{ filters.splice(idx,1); onDirty(); renderFilters(); };

          row.appendChild(col); row.appendChild(op); row.appendChild(val); row.appendChild(del);
          fCard.appendChild(row);
        });
      };

      const addF=el("button",{className:"btn btn-primary",textContent:"＋ Додати фільтр"});
      addF.onclick=()=>{ filters.push({col:sh.columns[0]?.name||"",op:"contains",value:""}); onDirty(); renderFilters(); };
      fCard.appendChild(el("div",{className:"row"})).appendChild(addF);

      renderFilters();
      area.appendChild(fCard);
    };

    pick.onchange=render; render();
    root.appendChild(card);
  }

  }
  return root;
}
