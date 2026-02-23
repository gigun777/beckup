// js/transfer_ui.js
// UI builder for Transfer Templates v2 (route-based, index-driven).
// A template has multiple routes (paths). Each route maps 1+ source columns from FROM sheet to 1 target column in TO sheet.
// Sources are always from template's fromSheetKey (per product requirement).

import { el } from "./ui.js";
// Versioned import to avoid Edge module caching issues
import { getTransferTemplates, setTransferTemplates } from "./transfer.js?v=12.5.1.3";
import { CASE_DESC_COLUMNS } from "./schema.js";

function jsonClone(x){ return JSON.parse(JSON.stringify(x)); }
function opt(value, label){ return el("option",{value, textContent: label}); }

function sheetByKey(sheets, key){ return sheets.find(s=>s.key===key) || null; }

function buildSheetPicker(sheets, value, onChange){
  const sel = el("select",{className:"select"});
  sel.appendChild(opt("__case__","Опис справи"));
  for(const sh of sheets){
    sel.appendChild(opt(sh.key, sh.title));
  }
  sel.value = value || (sheets[0]?.key||"__case__");
  sel.onchange = ()=>onChange(sel.value);
  return sel;
}

function buildToPicker(sheets, value, onChange){
  // v12.3.0: allow destination to be either a Sheet or the Case Description table.
  const sel = el("select",{className:"select"});
  sel.appendChild(opt("__case__","Опис справи"));
  for(const sh of sheets){
    sel.appendChild(opt(sh.key, sh.title));
  }
  sel.value = value || ("__case__");
  sel.onchange = ()=>onChange(sel.value);
  return sel;
}

function casePseudoSheet(){
  return { key:"__case__", title:"Опис справи", columns: CASE_DESC_COLUMNS.map(c=>({name:c.name})) };
}

function sheetOrCaseByKey(sheets, key){
  if(key==="__case__") return casePseudoSheet();
  return sheetByKey(sheets, key);
}

function colLabel(sheet, idx0){
  const c = sheet?.columns?.[idx0];
  return c ? `${idx0+1}) ${c.name}` : `${idx0+1}`;
}
function buildColPicker(sheet, valueIdx, onChange){
  const sel = el("select",{className:"select"});
  (sheet?.columns||[]).forEach((c,i)=>{
    sel.appendChild(opt(String(i), `${i+1}) ${c.name}`));
  });
  if(valueIdx==null || valueIdx<0) valueIdx = 0;
  sel.value = String(valueIdx);
  sel.onchange = ()=>onChange(parseInt(sel.value,10));
  return sel;
}

function buildOpPicker(value, onChange){
  const sel = el("select",{className:"select"});
  sel.appendChild(opt("concat","Конкатенація"));
  sel.appendChild(opt("sum","Сумування"));
  sel.appendChild(opt("seq","Послідовний запис"));
  sel.appendChild(opt("newline","З нової строки"));
  sel.value = value || "concat";
  sel.onchange = ()=>onChange(sel.value);
  return sel;
}

function buildDelimiterPicker(value, onChange){
  const sel = el("select",{className:"select"});
  const delims = ["", " ", "-", "/", ":", ";", ",", "."];
  for(const d of delims){
    const label = (d==="") ? "(без розділювача)" : (d===" " ? "(пробіл)" : d);
    sel.appendChild(opt(d, label));
  }
  sel.value = (value==null) ? " " : value;
  sel.onchange = ()=>onChange(sel.value);
  return sel;
}

function routeCard({fromSheet, toSheet, route, onChange, onDelete}){
  const wrap = el("div",{className:"card", style:"padding:10px; margin:10px 0;"});

  const hdr = el("div",{className:"row", style:"justify-content:space-between; align-items:center"});
  hdr.appendChild(el("div",{className:"muted", textContent:"Маршрут"}));
  const delBtn = el("button",{className:"btn btn-danger", textContent:"Видалити маршрут"});
  delBtn.onclick = ()=>onDelete();
  hdr.appendChild(delBtn);
  wrap.appendChild(hdr);

  // Sources list
  const srcBlock = el("div",{style:"margin-top:8px"});
  srcBlock.appendChild(el("div",{className:"muted", textContent:"Джерела (колонки листа-джерела)"}));

  const srcList = el("div",{});
  const renderSources = ()=>{
    srcList.innerHTML="";
    (route.sources||[]).forEach((idx, si)=>{
      const rrow = el("div",{className:"row", style:"align-items:center; gap:8px; margin-top:6px"});
      const sel = buildColPicker(fromSheet, idx, (v)=>{
        route.sources[si]=v;
        onChange();
      });
      rrow.appendChild(sel);
      const rm = el("button",{className:"btn", textContent:"−"});
      rm.onclick = ()=>{
        route.sources.splice(si,1);
        if(!route.sources.length) route.sources=[0];
        onChange(); renderSources();
      };
      rrow.appendChild(rm);
      srcList.appendChild(rrow);
    });
  };
  renderSources();
  srcBlock.appendChild(srcList);

  const addSrc = el("button",{className:"btn", textContent:"+ Джерело"});
  addSrc.onclick = ()=>{
    route.sources = route.sources || [];
    route.sources.push(0);
    onChange(); renderSources();
  };
  srcBlock.appendChild(el("div",{style:"margin-top:6px"})).appendChild(addSrc);
  wrap.appendChild(srcBlock);

  // Operation + delimiter
  const opRow = el("div",{className:"row", style:"margin-top:10px; align-items:center; gap:8px"});
  opRow.appendChild(el("div",{className:"muted", textContent:"Правило:"}));
  const opSel = buildOpPicker(route.op, (v)=>{
    route.op=v;
    if(v==="seq") route.delimiter="";
    if(v==="newline") route.delimiter="\n";
    onChange(); renderDelim();
  });
  opRow.appendChild(opSel);

  const delimWrap = el("div",{});
  function renderDelim(){
    delimWrap.innerHTML="";
    if(route.op!=="concat") return;
    delimWrap.appendChild(el("div",{className:"muted", textContent:"Розділювач:"}));
    const dsel = buildDelimiterPicker(route.delimiter, (v)=>{ route.delimiter=v; onChange(); });
    delimWrap.appendChild(dsel);
  }
  renderDelim();
  opRow.appendChild(delimWrap);
  wrap.appendChild(opRow);

  // Target
  const tgtRow = el("div",{className:"row", style:"margin-top:10px; align-items:center; gap:8px"});
  tgtRow.appendChild(el("div",{className:"muted", textContent:"Цільова колонка (лист-призначення):"}));
  const tgtSel = buildColPicker(toSheet, route.targetCol ?? 0, (v)=>{ route.targetCol=v; onChange(); });
  tgtRow.appendChild(tgtSel);
  wrap.appendChild(tgtRow);

  // Preview
  const prev = el("div",{className:"muted", style:"margin-top:8px"});
  const preview = ()=>{
    const srcNames = (route.sources||[]).map(i=>colLabel(fromSheet,i)).join(" + ");
    const opName = route.op==="sum" ? "Σ" : (route.op==="seq" ? "⧺" : "⊕");
    const delim = (route.op==="concat") ? (route.delimiter==="" ? "''" : route.delimiter) : "";
    prev.textContent = `${srcNames}  ${opName}${delim?("("+delim+")"):""}  →  ${colLabel(toSheet, route.targetCol??0)}`;
  };
  preview();
  wrap.appendChild(prev);

  // keep preview updated
  wrap._updatePreview = preview;
  return wrap;
}

export async function buildTransferUI({sheets, onDirty}){
  const root = el("div",{});
  let templates = await getTransferTemplates();
  templates = jsonClone(templates||[]);
  // If empty, create one empty template to start (user can delete).
  if(!templates.length && sheets.length){
    templates.push({
      id: crypto.randomUUID(),
      name: "Новий шаблон",
      fromSheetKey: sheets[0].key,
      toType: "sheet",
      toSheetKey: sheets[0].key,
      routes: [{sources:[0], op:"concat", delimiter:" ", targetCol:0}]
    });
  }

  let activeId = templates[0]?.id || null;

  const layout = el("div",{className:"row", style:"gap:12px; align-items:flex-start"});
  const left = el("div",{style:"min-width:260px; max-width:320px;"});
  const right = el("div",{style:"flex:1;"});
  layout.appendChild(left); layout.appendChild(right);
  root.appendChild(layout);

  const saveBar = el("div",{className:"row", style:"margin-top:10px; gap:8px"});
  const btnSave = el("button",{className:"btn btn-primary", textContent:"Зберегти"});
  const btnReload = el("button",{className:"btn", textContent:"Перезавантажити"});
  saveBar.appendChild(btnSave); saveBar.appendChild(btnReload);
  root.appendChild(saveBar);

  function activeTpl(){ return templates.find(t=>t.id===activeId) || templates[0] || null; }

  function markDirty(){ onDirty && onDirty(); }

  function renderLeft(){
    left.innerHTML="";
    left.appendChild(el("div",{className:"muted", textContent:"Шаблони перенесення"}));
    const list = el("div",{style:"margin-top:8px"});
    templates.forEach((t,i)=>{
      const b = el("button",{className:"btn", style:`display:block; width:100%; text-align:left; margin:4px 0; ${t.id===activeId?"font-weight:700":""}`, textContent:`${i+1}. ${t.name||"(без назви)"}`});
      b.onclick = ()=>{ activeId=t.id; renderRight(); renderLeft(); };
      list.appendChild(b);
    });
    left.appendChild(list);

    const add = el("button",{className:"btn btn-primary", textContent:"+ Додати шаблон", style:"margin-top:8px; width:100%"});
    add.onclick = ()=>{
      const base = activeTpl() || {};
      const fromKey = base.fromSheetKey || sheets[0]?.key;
      const toKey = base.toSheetKey || sheets[0]?.key;
      templates.push({
        id: crypto.randomUUID(),
        name: "Новий шаблон",
        fromSheetKey: fromKey,
        toType:"sheet",
        toSheetKey: toKey,
        routes: [{sources:[0], op:"concat", delimiter:" ", targetCol:0}]
      });
      activeId = templates[templates.length-1].id;
      markDirty(); renderLeft(); renderRight();
    };
    left.appendChild(add);

    const del = el("button",{className:"btn btn-danger", textContent:"Видалити шаблон", style:"margin-top:6px; width:100%"});
    del.onclick = ()=>{
      if(!templates.length) return;
      const t=activeTpl(); if(!t) return;
      if(!confirm(`Видалити шаблон "${t.name||""}"?`)) return;
      templates = templates.filter(x=>x.id!==t.id);
      activeId = templates[0]?.id || null;
      markDirty(); renderLeft(); renderRight();
    };
    left.appendChild(del);
  }

  function renderRight(){
    right.innerHTML="";
    const t = activeTpl();
    if(!t){ right.appendChild(el("div",{className:"muted", textContent:"Немає шаблонів."})); return; }

    const fromSheet = sheetOrCaseByKey(sheets, t.fromSheetKey) || sheets[0];
    const toSheet = sheetOrCaseByKey(sheets, t.toSheetKey) || sheets[0];

    const head = el("div",{className:"card"});
    head.appendChild(el("h4",{textContent:"Редагування шаблону перенесення"}));
    right.appendChild(head);

    const nameRow = el("div",{className:"row", style:"align-items:center; gap:8px; margin-top:6px"});
    nameRow.appendChild(el("div",{className:"muted", textContent:"Назва:"}));
    const nameInput = el("input",{className:"input", value:t.name||"", style:"flex:1"});
    nameInput.oninput = ()=>{ t.name=nameInput.value; markDirty(); renderLeft(); };
    nameRow.appendChild(nameInput);
    head.appendChild(nameRow);

    const sheetRow = el("div",{className:"row", style:"align-items:center; gap:10px; margin-top:8px"});
    sheetRow.appendChild(el("div",{className:"muted", textContent:"З листа:"}));
    const fromPick = buildSheetPicker(sheets, t.fromSheetKey, (v)=>{
      t.fromSheetKey=v;
      // reset routes sources if out of range
      const fs = sheetOrCaseByKey(sheets,v);
      const max = (fs?.columns||[]).length-1;
      (t.routes||[]).forEach(r=>{
        r.sources=(r.sources||[0]).map(i=>Math.min(Math.max(0,i),max>=0?max:0));
      });
      markDirty(); renderRight();
    });
    sheetRow.appendChild(fromPick);

    sheetRow.appendChild(el("div",{className:"muted", textContent:"→ До листа:"}));
    const toPick = buildToPicker(sheets, t.toSheetKey, (v)=>{
      t.toSheetKey=v;
      const ts = sheetOrCaseByKey(sheets,v);
      const max = (ts?.columns||[]).length-1;
      (t.routes||[]).forEach(r=>{
        r.targetCol = Math.min(Math.max(0, r.targetCol ?? 0), max>=0?max:0);
      });
      markDirty(); renderRight();
    });
    sheetRow.appendChild(toPick);
    head.appendChild(sheetRow);

    const routesWrap = el("div",{style:"margin-top:10px"});
    routesWrap.appendChild(el("div",{className:"muted", textContent:"Маршрути перенесення (1 маршрут = 1 цільова комірка)"}));

    const routesList = el("div",{});
    const renderRoutes = ()=>{
      routesList.innerHTML="";
      const fs = sheetOrCaseByKey(sheets, t.fromSheetKey) || sheets[0];
      const ts = sheetOrCaseByKey(sheets, t.toSheetKey) || sheets[0];
      (t.routes||[]).forEach((r,ri)=>{
        const card = routeCard({
          fromSheet: fs,
          toSheet: ts,
          route: r,
          onChange: ()=>{
            // normalize
            r.sources = (r.sources||[]).map(n=>Number.isFinite(n)?n:0);
            if(!r.sources.length) r.sources=[0];
            if(r.op==="seq"){ r.delimiter=""; }
            markDirty();
            // update preview
            card._updatePreview && card._updatePreview();
          },
          onDelete: ()=>{
            t.routes.splice(ri,1);
            if(!t.routes.length) t.routes=[{sources:[0], op:"concat", delimiter:" ", targetCol:0}];
            markDirty(); renderRoutes();
          }
        });
        routesList.appendChild(card);
      });
    };
    renderRoutes();
    routesWrap.appendChild(routesList);

    const addRoute = el("button",{className:"btn btn-primary", textContent:"+ Додати маршрут", style:"margin-top:8px"});
    addRoute.onclick = ()=>{
      t.routes = t.routes || [];
      t.routes.push({sources:[0], op:"concat", delimiter:" ", targetCol:0});
      markDirty(); renderRoutes();
    };
    routesWrap.appendChild(addRoute);

    right.appendChild(routesWrap);
  }

  btnSave.onclick = async ()=>{
    // persist
    await setTransferTemplates(templates);
    alert("Збережено шаблони перенесення.");
  };
  btnReload.onclick = async ()=>{
    templates = jsonClone((await getTransferTemplates())||[]);
    activeId = templates[0]?.id || null;
    renderLeft(); renderRight();
  };

  renderLeft();
  renderRight();
  return root;
}
