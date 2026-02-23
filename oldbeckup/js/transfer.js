// js/transfer.js
import { cfgGet, cfgSet } from "./db.js";
export async function getTransferRules(){ return (await cfgGet("transfer_rules")) || []; }
export async function setTransferRules(rules){ await cfgSet("transfer_rules", rules); }
export async function ensureDefaultTransferRules(){
  let rules = await getTransferRules();
  if(rules.some(r=>r._default)) return;
  rules.push({
    id: crypto.randomUUID(),
    _default:true,
    fromSheetKey:"vkhidni_potochni",
    name:"Перенесення Вхідні → Справа",
    targets:[
      { to:"case", subrowSource:"none",
        map:[
          { destCol:"Вхідний", expr:{op:"concat", joiner:" / ", parts:[
            {op:"field", from:"data", col:"Вхідний номер документа"},
            {op:"field", from:"data", col:"Індекс документа"}
          ]}},
          { destCol:"Найменування документа", expr:{op:"field", from:"data", col:"Короткий зміст"}},
          { destCol:"Номери аркушів", expr:{op:"field", from:"data", col:"Номери аркушів у справі"}},
          { destCol:"Примітки", expr:{op:"field", from:"data", col:"Примітки"}}
        ]
      }
    ]
  });
  rules.push({
    id: crypto.randomUUID(),
    _default:true,
    fromSheetKey:"vykhidni_potochni",
    name:"Перенесення Вихідні → Справа (з підрядка)",
    targets:[
      { to:"case", subrowSource:"ask",
        map:[
          { destCol:"Вихідний", expr:{op:"concat", joiner:" / ", parts:[
            {op:"field", from:"data", col:"Номер вихідного листа"},
            {op:"field", from:"data", col:"Індекс вихідного листа"}
          ]}},
          { destCol:"Вхідний", expr:{op:"field", from:"data", col:"Номер та індекс вхідного документа"}},
          { destCol:"Найменування документа", expr:{op:"field", from:"data", col:"Короткий зміст"}},
          { destCol:"Номери аркушів", expr:{op:"concat", joiner:"; ", parts:[
            {op:"field",from:"sub",col:"Кількість аркушів документа", subIndex:"selected"},
            {op:"field",from:"sub",col:"Кількість аркушів Додатку", subIndex:"selected"}
          ]}},
          { destCol:"Примітки", expr:{op:"field", from:"data", col:"Примітки"}}
        ]
      }
    ]
  });

  // Мінімальні перенесення як у index_single: Вхідні → Інвентарні / Створені інвентарні
  rules.push({
    id: crypto.randomUUID(),
    _default:true,
    fromSheetKey:"vkhidni_potochni",
    name:"Перенесення Вхідні → Інвентарні документи (мінімально)",
    targets:[
      { to:"sheet", toSheetKey:"inventarni_dokumenty", subrowSource:"none",
        map:[
          { destCol:"Назва видання", expr:{op:"field", from:"data", col:"Короткий зміст"}},
          { destCol:"Вхідний реєстраційний індекс супровідного листа", expr:{op:"field", from:"data", col:"Вхідний номер документа"}},
          { destCol:"Дата супровідного листа", expr:{op:"field", from:"data", col:"Дата надходження"}},
          { destCol:"Примітки", expr:{op:"field", from:"data", col:"Примітки"}}
        ]
      }
    ]
  });

  rules.push({
    id: crypto.randomUUID(),
    _default:true,
    fromSheetKey:"vkhidni_potochni",
    name:"Перенесення Вхідні → Створені інвентарні (мінімально)",
    targets:[
      { to:"sheet", toSheetKey:"stvoreni_inventarni", subrowSource:"none",
        map:[
          { destCol:"Індекс документу", expr:{op:"field", from:"data", col:"Індекс документа"}},
          { destCol:"Дата", expr:{op:"field", from:"data", col:"Дата надходження"}},
          { destCol:"Короткий зміст", expr:{op:"field", from:"data", col:"Короткий зміст"}},
          { destCol:"Відмітка про виконання документа", expr:{op:"field", from:"data", col:"Відмітка про виконання"}},
          { destCol:"Примітки", expr:{op:"field", from:"data", col:"Примітки"}}
        ]
      }
    ]
  });
  await setTransferRules(rules);
}


// ---- Templates v2 (route-based, index-driven) ----
export async function getTransferTemplates(){
  return (await cfgGet("transfer_templates_v2")) || [];
}
export async function setTransferTemplates(tpls){
  await cfgSet("transfer_templates_v2", tpls);
}

function _colIndexByName(sheet, name){
  const idx = (sheet?.columns||[]).findIndex(c=>c.name===name);
  return idx>=0 ? idx : null;
}
function _exprCollectSourceIdx(fromSheet, expr, out){
  if(!expr) return;
  if(expr.op==="field"){
    const idx=_colIndexByName(fromSheet, expr.col);
    if(idx!=null) out.push(idx);
    return;
  }
  if(expr.op==="concat" || expr.op==="sum"){
    (expr.parts||[]).forEach(p=>_exprCollectSourceIdx(fromSheet,p,out));
  }
}
function _exprToRoute(fromSheet, toSheet, destColName, expr){
  const sources=[];
  _exprCollectSourceIdx(fromSheet, expr, sources);
  const tgtIdx = _colIndexByName(toSheet, destColName);
  if(tgtIdx==null) return null;
  const op = (expr?.op==="sum") ? "sum" : "concat";
  const delimiter = (op==="concat") ? String(expr?.joiner ?? " ") : "";
  // seq = concat with empty delimiter
  const op2 = (op==="concat" && delimiter==="") ? "seq" : op;
  return { sources, op: op2, delimiter, targetCol: tgtIdx };
}

/**
 * Ensure v2 templates exist. If none, tries to migrate from legacy "transfer_rules"
 * using current sheets schema. If migration yields nothing, creates empty list.
 */
export async function ensureDefaultTransferTemplates(sheets){
  let tpls = await getTransferTemplates();
  if(tpls && tpls.length) return;

  const legacy = await getTransferRules();
  const byKey = new Map((sheets||[]).map(s=>[s.key,s]));
  const migrated = [];

  for(const r of (legacy||[])){
    const tgt = r.targets?.[0];
    if(!tgt) continue;
    const fromSheet = byKey.get(r.fromSheetKey);
    if(!fromSheet) continue;

    if(tgt.to!=="sheet") continue; // v12.2 UI focuses on sheet→sheet templates
    const toSheet = byKey.get(tgt.toSheetKey);
    if(!toSheet) continue;

    const routes=[];
    for(const m of (tgt.map||[])){
      const route=_exprToRoute(fromSheet,toSheet,m.destCol,m.expr);
      if(route && route.sources.length) routes.push(route);
    }
    if(!routes.length) continue;

    migrated.push({
      id: r.id || crypto.randomUUID(),
      _default: !!r._default,
      name: r.name || "Шаблон перенесення",
      fromSheetKey: r.fromSheetKey,
      toType: "sheet",
      toSheetKey: tgt.toSheetKey,
      routes
    });
  }

  // If nothing migrated, keep empty; user can create templates in UI.
  await setTransferTemplates(migrated);
}
