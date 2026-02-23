(function(){
  const uid=()=>Math.random().toString(16).slice(2)+Date.now().toString(16);
  const clamp=(n,a,b)=>Math.max(a,Math.min(b,n));
  const toStr=v=>(v===null||v===undefined)?"":String(v);
  const toNum=v=>{const x=typeof v==="number"?v:parseFloat(String(v).replace(",","."));return Number.isFinite(x)?x:NaN;};
  const deepClone=x=>JSON.parse(JSON.stringify(x));
  const escapeHtml=s=>String(s).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;");

  function normalizeTemplates(templates,sheets){
    const sheetKeys=new Set((sheets||[]).map(s=>s.key));
    const safeKey=k=>sheetKeys.has(k)?k:(sheets[0]?.key??"");
    const colCount=k=>(sheets.find(s=>s.key===k)?.columns?.length??0);
    (templates||[]).forEach(t=>{
      if(!t.id) t.id=uid();
      t.name=toStr(t.name)||"Новий шаблон";
      t.fromSheetKey=safeKey(t.fromSheetKey);
      t.toSheetKey=safeKey(t.toSheetKey);
      if(!Array.isArray(t.routes)||!t.routes.length) t.routes=[{sources:[0],op:"concat",delimiter:" ",targetCol:0}];
      t.routes.forEach(r=>{
        if(!Array.isArray(r.sources)||!r.sources.length) r.sources=[0];
        r.sources=r.sources.map(i=>clamp(parseInt(i||0,10)||0,0,Math.max(0,colCount(t.fromSheetKey)-1)));
        r.op=r.op||"concat";
        if(r.op==="concat"&&(r.delimiter===undefined||r.delimiter===null)) r.delimiter=" ";
        if(r.op!=="concat") r.delimiter=r.delimiter??"";
        r.targetCol=clamp(parseInt(r.targetCol||0,10)||0,0,Math.max(0,colCount(t.toSheetKey)-1));
      });
    });
    return templates;
  }
  function routeOpLabel(op,delim){
    if(op==="sum") return "sum";
    if(op==="newline") return "join(\\n)";
    if(op==="seq") return "seq";
    if(op==="concat") return `join(${delim===""?"∅":JSON.stringify(delim)})`;
    return op||"concat";
  }
  function computeRouteResult(route,sourceRow){
    const values=(route.sources||[]).map(i=>sourceRow?.[i]);
    const op=route.op||"concat";
    if(op==="sum"){
      const nums=values.map(toNum).filter(Number.isFinite);
      return nums.length?String(nums.reduce((a,b)=>a+b,0)):"";
    }
    if(op==="newline") return values.map(toStr).filter(s=>s!=="").join("\n");
    if(op==="seq") return values.map(toStr).filter(s=>s!=="").join("");
    const d=route.delimiter??" ";
    return values.map(toStr).filter(s=>s!=="").join(d);
  }
  function makeEmptyRowForSheet(sheet){const n=sheet?.columns?.length??0;return Array.from({length:n},()=> "");}

  const $settings=document.getElementById("dvTransferSettingsOverlay");
  const $exec=document.getElementById("dvTransferExecOverlay");

  function __applyOverlayInlineStyles(ov){
    if(!ov) return;
    // Make overlay work even if CSS wasn't loaded
    ov.style.position = "fixed";
    ov.style.inset = "0";
    ov.style.zIndex = "99999";
    ov.style.background = "rgba(0,0,0,0.35)";
    ov.style.alignItems = "center";
    ov.style.justifyContent = "center";
    ov.style.padding = "12px";
    ov.style.boxSizing = "border-box";
    if(!ov.classList.contains("dv-open")) ov.style.display = "none";
    ov.setAttribute("aria-hidden", ov.classList.contains("dv-open") ? "false" : "true");
    // Ensure inner panel is readable
    const panel = ov.querySelector(".dv-modal") || ov.firstElementChild;
    if(panel){
      panel.style.maxWidth = panel.style.maxWidth || "980px";
      panel.style.width = panel.style.width || "min(980px, 96vw)";
      panel.style.maxHeight = panel.style.maxHeight || "90vh";
      panel.style.overflow = panel.style.overflow || "auto";
      panel.style.background = panel.style.background || "#fff";
      panel.style.borderRadius = panel.style.borderRadius || "10px";
      panel.style.boxShadow = panel.style.boxShadow || "0 10px 30px rgba(0,0,0,0.25)";
    }
  }
  __applyOverlayInlineStyles($settings);
  __applyOverlayInlineStyles($exec);

  const $tplList=document.getElementById("dvTplList");
  const $tplAdd=document.getElementById("dvTplAdd");
  const $tplDel=document.getElementById("dvTplDel");
  const $tplName=document.getElementById("dvTplName");
  const $fromSheet=document.getElementById("dvFromSheet");
  const $toSheet=document.getElementById("dvToSheet");
  const $tplHint=document.getElementById("dvTplHint");
  const $routes=document.getElementById("dvRoutes");
  const $routeAdd=document.getElementById("dvRouteAdd");
  const $settingsSave=document.getElementById("dvSettingsSave");
  const $settingsReload=document.getElementById("dvSettingsReload");
  const $settingsMsg=document.getElementById("dvSettingsMsg");

  const $execTemplate=document.getElementById("dvExecTemplate");
  const $execSubtitle=document.getElementById("dvTransferExecSubtitle");
  const $execInfo=document.getElementById("dvExecInfo");
  const $execPreview=document.getElementById("dvExecPreview");
  const $execApply=document.getElementById("dvExecApply");
  const $execMsg=document.getElementById("dvExecMsg");
  const $execGoToTarget=document.getElementById("dvExecGoToTarget");
  const $execCloseOnSuccess=document.getElementById("dvExecCloseOnSuccess");

  let S={sheets:[],templatesOriginal:[],templatesDraft:[],activeTemplateId:null,onSave:null,onClose:null,exec:{sourceSheetKey:"",sourceRow:[],onApply:null,onClose:null}};
  const sheetByKey=k=>S.sheets.find(s=>s.key===k);
  const colName=(sheetKey,idx)=>sheetByKey(sheetKey)?.columns?.[idx]?.name ?? `Колонка ${idx+1}`;

  function openOverlay(which){
    const ov=which==="settings"?$settings:$exec;
    if(!ov) return;
    ov.classList.add("dv-open"); ov.setAttribute("aria-hidden","false");
    ov.style.display="flex";
    document.body.style.overflow="hidden";
  }
  function closeOverlay(which){
    const ov=which==="settings"?$settings:$exec;
    ov.classList.remove("dv-open"); ov.setAttribute("aria-hidden","true");
    ov.style.display="none";
    if(!$settings.classList.contains("dv-open")&&!$exec.classList.contains("dv-open")) document.body.style.overflow="";
  }
  function showMsg($el,text){$el.style.display=text?"":"none";$el.textContent=text||"";}
  function fillSheetSelect($sel,sheets,selectedKey){
    $sel.innerHTML="";
    for(const sh of sheets){const opt=document.createElement("option"); opt.value=sh.key; opt.textContent=sh.name||sh.key; $sel.appendChild(opt);}
    $sel.value=selectedKey || (sheets[0]?.key??"");
  }

  const activeTemplate=()=>S.templatesDraft.find(t=>t.id===S.activeTemplateId)||null;
  function setActiveTemplate(id){S.activeTemplateId=id; renderTemplateList(); renderEditor();}

  function renderTemplateList(){
    $tplList.innerHTML="";
    const tpls=S.templatesDraft;
    if(!tpls.length){
      const d=document.createElement("div"); d.className="dv-hint"; d.textContent="Шаблонів ще немає. Додайте перший шаблон.";
      $tplList.appendChild(d); $tplDel.disabled=true; return;
    }
    tpls.forEach((t,i)=>{
      const item=document.createElement("div"); item.className="dv-list-item"+(t.id===S.activeTemplateId?" dv-active":""); item.onclick=()=>setActiveTemplate(t.id);
      const left=document.createElement("div"); left.style.display="flex"; left.style.flexDirection="column"; left.style.gap="4px";
      const title=document.createElement("div"); title.style.fontWeight="850"; title.style.fontSize="13px"; title.textContent=`${i+1}. ${t.name||"Без назви"}`;
      const meta=document.createElement("div"); meta.style.fontSize="11.5px"; meta.style.color="rgba(0,0,0,.55)";
      meta.textContent=`${sheetByKey(t.fromSheetKey)?.name||t.fromSheetKey} → ${sheetByKey(t.toSheetKey)?.name||t.toSheetKey}`;
      const badge=document.createElement("div"); badge.className="dv-badge"; badge.textContent=`${t.routes?.length||0} маршрут(ів)`;
      left.appendChild(title); left.appendChild(meta); item.appendChild(left); item.appendChild(badge); $tplList.appendChild(item);
    });
    $tplDel.disabled=false;
  }

  function renderEditor(){
    const t=activeTemplate(); const has=!!t;
    $tplName.disabled=!has; $fromSheet.disabled=!has; $toSheet.disabled=!has; $routeAdd.disabled=!has;
    if(!has){$tplName.value=""; $routes.innerHTML=""; $tplHint.textContent=""; return;}
    $tplName.value=t.name||"";
    fillSheetSelect($fromSheet,S.sheets,t.fromSheetKey);
    fillSheetSelect($toSheet,S.sheets,t.toSheetKey);
    $tplHint.textContent=`З листа: ${sheetByKey(t.fromSheetKey)?.name||t.fromSheetKey} → До листа: ${sheetByKey(t.toSheetKey)?.name||t.toSheetKey} • Маршрутів: ${t.routes?.length||0}`;
    renderRoutes();
  }

  function renderRoutes(){
    const t=activeTemplate(); $routes.innerHTML=""; if(!t) return;
    const fromCols=sheetByKey(t.fromSheetKey)?.columns??[];
    const toCols=sheetByKey(t.toSheetKey)?.columns??[];

    t.routes.forEach((r,idx)=>{
      const card=document.createElement("div"); card.className="dv-route";
      const head=document.createElement("div"); head.className="dv-route-h"; head.innerHTML=`<div>Маршрут №${idx+1}</div>`;
      const delBtn=document.createElement("button"); delBtn.className="dv-btn dv-btn-danger"; delBtn.textContent="🗑 Видалити маршрут";
      delBtn.onclick=()=>{t.routes.splice(idx,1); if(!t.routes.length) t.routes.push({sources:[0],op:"concat",delimiter:" ",targetCol:0}); renderTemplateList(); renderEditor();};
      head.appendChild(delBtn);

      const body=document.createElement("div"); body.className="dv-route-b";

      const sourcesWrap=document.createElement("div"); sourcesWrap.className="dv-col";
      const sLabel=document.createElement("div"); sLabel.className="dv-label"; sLabel.textContent="Джерела (колонки листа-джерела)";
      sourcesWrap.appendChild(sLabel);

      (r.sources||[]).forEach((srcIdx,sIdx)=>{
        const row=document.createElement("div"); row.className="dv-source-row";
        const sel=document.createElement("select"); sel.className="dv-select"; sel.style.minWidth="260px";
        for(let i=0;i<fromCols.length;i++){const opt=document.createElement("option"); opt.value=String(i); opt.textContent=`${i+1}. ${fromCols[i]?.name ?? ("Колонка "+(i+1))}`; sel.appendChild(opt);}
        sel.value=String(clamp(srcIdx,0,Math.max(0,fromCols.length-1)));
        sel.onchange=()=>{r.sources[sIdx]=parseInt(sel.value,10)||0; renderTemplateList();};

        const minus=document.createElement("button"); minus.className="dv-mini-btn dv-mini-danger"; minus.textContent="−";
        minus.onclick=()=>{r.sources.splice(sIdx,1); if(!r.sources.length) r.sources=[0]; renderEditor();};

        row.appendChild(sel); row.appendChild(minus); sourcesWrap.appendChild(row);
      });

      const addSource=document.createElement("button"); addSource.className="dv-btn"; addSource.textContent="+ Додати джерело";
      addSource.onclick=()=>{(r.sources=r.sources||[0]).push(0); renderEditor();};
      sourcesWrap.appendChild(addSource);

      const opWrap=document.createElement("div"); opWrap.className="dv-row";

      const opCol=document.createElement("div"); opCol.className="dv-col";
      const opLabel=document.createElement("div"); opLabel.className="dv-label"; opLabel.textContent="Правило";
      const opSel=document.createElement("select"); opSel.className="dv-select"; opSel.style.minWidth="260px";
      [{v:"concat",t:"Конкатенація"},{v:"seq",t:"Об’єднання без розділювача"},{v:"newline",t:"З нової строки"},{v:"sum",t:"Сумування (числа)"}]
        .forEach(o=>{const opt=document.createElement("option"); opt.value=o.v; opt.textContent=o.t; opSel.appendChild(opt);});
      opSel.value=r.op||"concat";
      opSel.onchange=()=>{r.op=opSel.value; if(r.op!=="concat") r.delimiter=r.delimiter??""; if(r.op==="concat"&&(r.delimiter===undefined||r.delimiter===null)) r.delimiter=" "; renderEditor(); renderTemplateList();};
      opCol.appendChild(opLabel); opCol.appendChild(opSel);

      const delimCol=document.createElement("div"); delimCol.className="dv-col";
      const delimLabel=document.createElement("div"); delimLabel.className="dv-label"; delimLabel.textContent="Розділювач";
      const delimSel=document.createElement("select"); delimSel.className="dv-select"; delimSel.style.minWidth="220px";
      [{v:"",t:"(немає)"},{v:" ",t:"пробіл"},{v:" - ",t:" - "},{v:"-",t:"-"},{v:"/",t:"/"},{v:":",t:":"},{v:"; ",t:"; "},{v:", ",t:", "},{v:". ",t:". "}]
        .forEach(d=>{const opt=document.createElement("option"); opt.value=d.v; opt.textContent=d.t; delimSel.appendChild(opt);});
      delimSel.value=r.delimiter??" ";
      delimSel.onchange=()=>{r.delimiter=delimSel.value; renderTemplateList();};
      delimCol.appendChild(delimLabel); delimCol.appendChild(delimSel);

      const tgtCol=document.createElement("div"); tgtCol.className="dv-col";
      const tgtLabel=document.createElement("div"); tgtLabel.className="dv-label"; tgtLabel.textContent="Цільова колонка (лист-призначення)";
      const tgtSel=document.createElement("select"); tgtSel.className="dv-select"; tgtSel.style.minWidth="260px";
      for(let i=0;i<toCols.length;i++){const opt=document.createElement("option"); opt.value=String(i); opt.textContent=`${i+1}. ${toCols[i]?.name ?? ("Колонка "+(i+1))}`; tgtSel.appendChild(opt);}
      tgtSel.value=String(clamp(r.targetCol,0,Math.max(0,toCols.length-1)));
      tgtSel.onchange=()=>{r.targetCol=parseInt(tgtSel.value,10)||0; renderTemplateList();};
      tgtCol.appendChild(tgtLabel); tgtCol.appendChild(tgtSel);

      opWrap.appendChild(opCol);
      opWrap.appendChild((r.op==="concat")?delimCol:document.createElement("div"));
      opWrap.appendChild(tgtCol);

      const preview=document.createElement("div"); preview.className="dv-route-preview";
      const srcNames=(r.sources||[]).map(i=>`[${colName(t.fromSheetKey,i)}]`).join(" + ");
      const opL=routeOpLabel(r.op,r.delimiter??" ");
      const tgt=`[${colName(t.toSheetKey,r.targetCol)}]`;
      preview.innerHTML=`<div><b>Превʼю:</b> ${srcNames} <span class="dv-mono">(${opL})</span> → ${tgt}</div>`;

      body.appendChild(sourcesWrap); body.appendChild(opWrap); body.appendChild(preview);
      card.appendChild(head); card.appendChild(body); $routes.appendChild(card);
    });
  }

  function settingsAddTemplate(){
    const firstKey=S.sheets[0]?.key??"";
    const t={id:uid(),name:"Новий шаблон",fromSheetKey:firstKey,toSheetKey:firstKey,routes:[{sources:[0],op:"concat",delimiter:" ",targetCol:0}]};
    S.templatesDraft.push(t); setActiveTemplate(t.id);
  }
  function settingsDeleteActive(){
    const t=activeTemplate(); if(!t) return;
    if(!confirm(`Видалити шаблон "${t.name}"?`)) return;
    const i=S.templatesDraft.findIndex(x=>x.id===t.id);
    if(i>=0) S.templatesDraft.splice(i,1);
    S.activeTemplateId=S.templatesDraft[0]?.id??null;
    renderTemplateList(); renderEditor();
  }
  function settingsAddRoute(){const t=activeTemplate(); if(!t) return; t.routes.push({sources:[0],op:"concat",delimiter:" ",targetCol:0}); renderTemplateList(); renderEditor();}
  function settingsNormalizeAfterSheetChange(){const t=activeTemplate(); if(!t) return; normalizeTemplates([t],S.sheets); renderTemplateList(); renderEditor();}

  function renderExecTemplates(){
    const srcKey=S.exec.sourceSheetKey;
    const candidates=(S.templatesDraft||[]).filter(t=>t.fromSheetKey===srcKey);
    $execTemplate.innerHTML="";
    if(!candidates.length){
      const opt=document.createElement("option"); opt.value=""; opt.textContent="Немає шаблонів для цього листа";
      $execTemplate.appendChild(opt); $execTemplate.disabled=true; $execApply.disabled=true;
      showMsg($execMsg,"⚠️ Немає шаблонів перенесення для поточного листа. Створіть шаблон у Налаштуваннях.");
      $execInfo.textContent=""; $execPreview.innerHTML=""; return;
    }
    $execTemplate.disabled=false; $execApply.disabled=false; showMsg($execMsg,"");
    candidates.forEach(t=>{const opt=document.createElement("option"); opt.value=t.id; opt.textContent=t.name||"Без назви"; $execTemplate.appendChild(opt);});
    $execTemplate.value=candidates[0].id; renderExecPreview();
  }

  function renderExecPreview(){
    const tplId=$execTemplate.value;
    const t=S.templatesDraft.find(x=>x.id===tplId);
    if(!t){$execPreview.innerHTML=""; return;}
    const fromName=sheetByKey(t.fromSheetKey)?.name||t.fromSheetKey;
    const toName=sheetByKey(t.toSheetKey)?.name||t.toSheetKey;
    $execInfo.textContent=`З листа: ${fromName} → До листа: ${toName} • Маршрутів: ${t.routes?.length||0}`;
    const srcRow=S.exec.sourceRow||[];
    $execPreview.innerHTML="";
    (t.routes||[]).forEach(r=>{
      const tr=document.createElement("tr");
      const tdSrc=document.createElement("td");
      tdSrc.innerHTML=(r.sources||[]).map(i=>`<div>• ${colName(t.fromSheetKey,i)}: <span class="dv-mono">${escapeHtml(toStr(srcRow[i]))}</span></div>`).join("");
      const tdOp=document.createElement("td"); tdOp.className="dv-mono"; tdOp.textContent=routeOpLabel(r.op,r.delimiter??" ");
      const tdTgt=document.createElement("td"); tdTgt.textContent=colName(t.toSheetKey,r.targetCol);
      const tdRes=document.createElement("td"); tdRes.innerHTML=`<span class="dv-mono">${escapeHtml(computeRouteResult(r,srcRow))}</span>`;
      tr.appendChild(tdSrc); tr.appendChild(tdOp); tr.appendChild(tdTgt); tr.appendChild(tdRes);
      $execPreview.appendChild(tr);
    });
  }

  function execApply(){
    const tplId=$execTemplate.value;
    const t=S.templatesDraft.find(x=>x.id===tplId);
    if(!t){showMsg($execMsg,"⚠️ Шаблон не обрано."); return;}
    const targetSheet=sheetByKey(t.toSheetKey);
    if(!targetSheet){showMsg($execMsg,"⚠️ Лист-призначення не знайдено."); return;}
    const targetRow=makeEmptyRowForSheet(targetSheet);
    const srcRow=S.exec.sourceRow||[];
    for(const r of (t.routes||[])) targetRow[r.targetCol]=computeRouteResult(r,srcRow);
    const actions={goToTarget:!!$execGoToTarget.checked, closeOnSuccess:!!$execCloseOnSuccess.checked};
    try{
      if(typeof S.exec.onApply==="function") S.exec.onApply({template:deepClone(t), targetRow, actions});
      showMsg($execMsg,"✅ Перенесення підготовлено (targetRow сформовано).");
      if(actions.closeOnSuccess){closeOverlay("exec"); if(typeof S.exec.onClose==="function") S.exec.onClose();}
    }catch(e){
      showMsg($execMsg,"❌ Помилка перенесення: "+(e?.message||String(e)));
    }
  }

  document.querySelectorAll("[data-dv-close]").forEach(btn=>{
    btn.addEventListener("click",()=>{
      const which=btn.getAttribute("data-dv-close");
      if(which==="settings"){closeOverlay("settings"); if(typeof S.onClose==="function") S.onClose();}
      else{closeOverlay("exec"); if(typeof S.exec.onClose==="function") S.exec.onClose();}
    });
  });

  $tplAdd.onclick=settingsAddTemplate;
  $tplDel.onclick=settingsDeleteActive;
  $routeAdd.onclick=settingsAddRoute;
  $tplName.oninput=()=>{const t=activeTemplate(); if(!t) return; t.name=$tplName.value; renderTemplateList();};
  $fromSheet.onchange=()=>{const t=activeTemplate(); if(!t) return; t.fromSheetKey=$fromSheet.value; settingsNormalizeAfterSheetChange();};
  $toSheet.onchange=()=>{const t=activeTemplate(); if(!t) return; t.toSheetKey=$toSheet.value; settingsNormalizeAfterSheetChange();};

  $settingsReload.onclick=()=>{
    S.templatesDraft=deepClone(S.templatesOriginal||[]);
    normalizeTemplates(S.templatesDraft,S.sheets);
    S.activeTemplateId=S.templatesDraft[0]?.id??null;
    showMsg($settingsMsg,"↻ Перезавантажено зі сховища (draft скинуто).");
    renderTemplateList(); renderEditor();
  };
  $settingsSave.onclick=()=>{
    try{
      normalizeTemplates(S.templatesDraft,S.sheets);
      const out=deepClone(S.templatesDraft);
      if(typeof S.onSave==="function") S.onSave(out);
      S.templatesOriginal=deepClone(out);
      showMsg($settingsMsg,"✅ Збережено.");
      renderTemplateList(); renderEditor();
    }catch(e){
      showMsg($settingsMsg,"❌ Помилка збереження: "+(e?.message||String(e)));
    }
  };

  $execTemplate.onchange=renderExecPreview;
  $execApply.onclick=execApply;

  [$settings,$exec].forEach(ov=>{
    ov.addEventListener("mousedown",(e)=>{
      if(e.target===ov){
        const which=ov===$settings?"settings":"exec";
        closeOverlay(which);
        if(which==="settings"){if(typeof S.onClose==="function") S.onClose();}
        else{if(typeof S.exec.onClose==="function") S.exec.onClose();}
      }
    });
  });

  window.addEventListener("keydown",(e)=>{
    if(e.key!=="Escape") return;
    if($exec.classList.contains("dv-open")){closeOverlay("exec"); if(typeof S.exec.onClose==="function") S.exec.onClose(); return;}
    if($settings.classList.contains("dv-open")){closeOverlay("settings"); if(typeof S.onClose==="function") S.onClose();}
  });

  window.TransferUI={
    openSettings(opts){
      S.sheets=deepClone(opts.sheets||[]);
      S.templatesOriginal=deepClone(opts.templates||[]);
      S.templatesDraft=deepClone(opts.templates||[]);
      normalizeTemplates(S.templatesDraft,S.sheets);
      S.onSave=typeof opts.onSave==="function"?opts.onSave:null;
      S.onClose=typeof opts.onClose==="function"?opts.onClose:null;
      document.getElementById("dvTransferSettingsTitle").textContent=opts.title||"Налаштування → Перенесення";
      showMsg($settingsMsg,"");
      if(!S.sheets.length) S.sheets=[{key:"default",name:"Default",columns:[{id:"c1",name:"Колонка 1"}]}];
      fillSheetSelect($fromSheet,S.sheets,S.sheets[0].key);
      fillSheetSelect($toSheet,S.sheets,S.sheets[0].key);
      S.activeTemplateId=S.templatesDraft[0]?.id??null;
      renderTemplateList(); renderEditor();
      openOverlay("settings");
    },
    openTransfer(opts){
      S.sheets=deepClone(opts.sheets||[]);
      S.templatesDraft=deepClone(opts.templates||[]);
      normalizeTemplates(S.templatesDraft,S.sheets);
      S.exec.sourceSheetKey=opts.sourceSheetKey || (S.sheets[0]?.key??"");
      S.exec.sourceRow=Array.isArray(opts.sourceRow)?opts.sourceRow:[];
      S.exec.onApply=typeof opts.onApply==="function"?opts.onApply:null;
      S.exec.onClose=typeof opts.onClose==="function"?opts.onClose:null;
      document.getElementById("dvTransferExecTitle").textContent=opts.title||"Перенесення запису";
      const srcName=sheetByKey(S.exec.sourceSheetKey)?.name||S.exec.sourceSheetKey;
      $execSubtitle.textContent=`Джерело: ${srcName}`;
      showMsg($execMsg,"");
      renderExecTemplates();
      openOverlay("exec");
    },
    closeAll(){if($exec.classList.contains("dv-open")) closeOverlay("exec"); if($settings.classList.contains("dv-open")) closeOverlay("settings");}
  };
})();
