// js/schema.js
export function safeName(s){
  return String(s||"").trim().replace(/[\\\/:*?"<>|]+/g,"_").replace(/\s+/g,"_").slice(0,120) || "export";
}
export function pad2(n){ return String(n).padStart(2,"0"); }
export function nowStamp(){
  const d=new Date();
  return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}_${pad2(d.getHours())}-${pad2(d.getMinutes())}-${pad2(d.getSeconds())}`;
}

// === CASE DESC COLUMNS (moved up to avoid TDZ) ===
export const CASE_DESC_COLUMNS = [
  {name:"№ з/п", type:"int", required:false, subrows:false, editable:false},
  {name:"Вихідний", type:"text", required:false, subrows:false},
  {name:"Вхідний", type:"text", required:false, subrows:false},
  {name:"Найменування документа", type:"text", required:false, subrows:false},
  {name:"Номери аркушів", type:"text", required:false, subrows:false},
  {name:"Примітки", type:"text", required:false, subrows:false}
];

export const DEFAULT_SHEETS = [
  { key:"vkhidni_potochni", title:"Вхідні поточні", orderColumn:"Номер за порядком",
    columns:[
      {name:"Номер за порядком", type:"int", required:false, subrows:false, editable:false},
      {name:"Вхідний номер документа", type:"int", required:true, subrows:false},
      {name:"Індекс документа", type:"text", required:false, subrows:false},
      {name:"Дата надходження", type:"date", required:true, subrows:false, defaultToday:true},
      {name:"Короткий зміст", type:"text", required:true, subrows:false},
      {name:"Кількість аркушів", type:"int", required:true, subrows:false},
      {name:"Кількість аркушів додатку", type:"int", required:false, subrows:false},
      {name:"Відповідальний виконавець", type:"text", required:false, subrows:false},
      {name:"Відмітка про взяття на контроль", type:"text", required:false, subrows:false, defaultValue:"Ні"},
      {name:"Відмітка про виконання", type:"text", required:false, subrows:false, defaultValue:"Не виконано"},
      {name:"Номер справи за номенклатурою", type:"text", required:false, subrows:false},
      {name:"Номери аркушів у справі", type:"text", required:false, subrows:false},
      {name:"Примітки", type:"text", required:false, subrows:false}
    ],
    addFields:["Вхідний номер документа","Індекс документа","Дата надходження","Короткий зміст","Кількість аркушів","Кількість аркушів додатку","Відповідальний виконавець","Примітки"],
    subrows:null, export:{pageSize:"A4", orientation:"portrait", includeColumns:null}
  },
  { key:"vykhidni_potochni", title:"Вихідні поточні", orderColumn:"Номер запису за порядком",
    columns:[
      {name:"Номер запису за порядком", type:"int", required:false, subrows:false, editable:false},
      {name:"Номер вихідного листа", type:"int", required:true, subrows:false},
      {name:"Індекс вихідного листа", type:"text", required:false, subrows:false},
      {name:"Номер та індекс вхідного документа", type:"text", required:false, subrows:false},
      {name:"Короткий зміст", type:"text", required:false, subrows:false},
      {name:"Кому надіслано", type:"text", required:false, subrows:true},
      {name:"Кількість примірників", type:"int", required:false, subrows:false, computed:true},
      {name:"Номер примірника", type:"int", required:false, subrows:true, computed:true},
      {name:"Кількість аркушів документа", type:"int", required:false, subrows:true},
      {name:"Кількість аркушів Додатку", type:"int", required:false, subrows:true},
      {name:"Виконавець", type:"text", required:false, subrows:false},
      {name:"Справа", type:"text", required:false, subrows:false},
      {name:"Аркуш", type:"text", required:false, subrows:false},
      {name:"Примітки", type:"text", required:false, subrows:false}
    ],
    addFields:["Номер вихідного листа","Індекс вихідного листа","Номер та індекс вхідного документа","Короткий зміст","Виконавець","Справа","Аркуш","Примітки","Кому надіслано","Кількість аркушів документа","Кількість аркушів Додатку"],
    subrows:{ key:"recipients", numberedBy:"Кому надіслано", autoNumberColumn:"Номер примірника" },
    export:{pageSize:"A4", orientation:"portrait", includeColumns:null}
  },
  { key:"nomenklatura", title:"Номенклатура", orderColumn:null,
    columns:[
      {name:"Підрозділ", type:"text", required:false, subrows:false},
      {name:"Індекс справи", type:"text", required:false, subrows:false},
      {name:"Індекс справи за номенклатурою справ минулого року", type:"text", required:false, subrows:false},
      {name:"Заголовок справи (тому, частини)", type:"text", required:false, subrows:false},
      {name:"Коли розпочато", type:"date", required:false, subrows:true},
      {name:"Коли закінчено", type:"date", required:false, subrows:true},
      {name:"Кількість аркушів", type:"int", required:false, subrows:true},
      {name:"Строк зберігання справи (тому, частини)і стаття за Переліком", type:"text", required:false, subrows:false},
      {name:"Примітка", type:"text", required:false, subrows:true},
      {name:"Статус", type:"text", required:false, subrows:true},
      {name:"Посилання на електронний архів", type:"text", required:false, subrows:true},
      {name:"Примітки", type:"text", required:false, subrows:false}
    ],
    addFields:["Підрозділ","Індекс справи","Заголовок справи (тому, частини)","Коли розпочато","Коли закінчено","Кількість аркушів","Примітка","Статус","Посилання на електронний архів","Примітки"],
    subrows:{ key:"periods", numberedBy:"Коли розпочато" },
    export:{pageSize:"A4", orientation:"landscape", includeColumns:null}
  }
  ,
  // === Журнал "Опис справ" (загальний, доступний за замовчуванням) ===
  { key:"opys_sprav", title:"Опис справ", orderColumn:"№ з/п",
    columns: CASE_DESC_COLUMNS.map((c, i)=>({
      name: c.name,
      type: (i===0?"int":"text"),
      required:false,
      subrows:false,
      editable: i===0?false:true
    })),
    addFields: CASE_DESC_COLUMNS.slice(1).map(c=>c.name),
    subrows:null,
    export:{pageSize:"A4", orientation:"portrait", includeColumns:null}
  }
  ,
  // === Додаткові базові шаблони з index_single ===
  { key:"arkhiv", title:"Архів", orderColumn:"№ за порядком",
    columns:[
      {name:"№ за порядком", type:"int", required:false, subrows:false, editable:false},
      {name:"Номери справ за номенклатурою справ", type:"text", required:false, subrows:false},
      {name:"Найменування справ", type:"text", required:false, subrows:false},
      {name:"Коли розпочато", type:"date", required:false, subrows:false},
      {name:"Коли закінчено", type:"date", required:false, subrows:false},
      {name:"Кількість аркушів", type:"int", required:false, subrows:false},
      {name:"Кількість аркушів додатка", type:"int", required:false, subrows:false},
      {name:"Строк та місце зберігання справ", type:"text", required:false, subrows:false},
      {name:"Відмітка про знищення", type:"text", required:false, subrows:false},
      {name:"Примітки", type:"text", required:false, subrows:false}
    ],
    addFields:[
      "Номери справ за номенклатурою справ",
      "Найменування справ",
      "Коли розпочато",
      "Коли закінчено",
      "Кількість аркушів",
      "Кількість аркушів додатка",
      "Строк та місце зберігання справ",
      "Відмітка про знищення",
      "Примітки"
    ],
    subrows:null,
    export:{pageSize:"A4", orientation:"landscape", includeColumns:null}
  },
  { key:"inventarni_dokumenty", title:"Інвентарні документи", orderColumn:"Номер за порядком",
    columns:[
      {name:"Номер за порядком", type:"int", required:false, subrows:false, editable:false},
      {name:"Назва видання", type:"text", required:false, subrows:false},
      {name:"Звідки надійшло або де надруковано", type:"text", required:false, subrows:false},
      {name:"Вхідний реєстраційний індекс супровідного листа", type:"text", required:false, subrows:false},
      {name:"Дата супровідного листа", type:"date", required:false, subrows:false},
      {name:"Кількість примірників", type:"int", required:false, subrows:false},
      {name:"Номери примірників", type:"text", required:false, subrows:false},
      {name:"Куди надіслано або кому видано", type:"text", required:false, subrows:false},
      {name:"Реєстраційний індекс вихідного документа або відмітка про отримання", type:"text", required:false, subrows:false},
      {name:"Дата вихідного документа, отримання", type:"date", required:false, subrows:false},
      {name:"Кількість примірників (повернення)", type:"int", required:false, subrows:false},
      {name:"Номери примірників (повернення)", type:"text", required:false, subrows:false},
      {name:"Дата повернення", type:"date", required:false, subrows:false},
      {name:"Номери повернених примірників", type:"text", required:false, subrows:false},
      {name:"Дата знищення", type:"date", required:false, subrows:false},
      {name:"Номер акта знищення", type:"text", required:false, subrows:false},
      {name:"Дата акта знищення", type:"date", required:false, subrows:false},
      {name:"Примітки", type:"text", required:false, subrows:false}
    ],
    addFields:[
      "Назва видання",
      "Звідки надійшло або де надруковано",
      "Вхідний реєстраційний індекс супровідного листа",
      "Дата супровідного листа",
      "Кількість примірників",
      "Номери примірників",
      "Куди надіслано або кому видано",
      "Реєстраційний індекс вихідного документа або відмітка про отримання",
      "Дата вихідного документа, отримання",
      "Кількість примірників (повернення)",
      "Номери примірників (повернення)",
      "Дата повернення",
      "Номери повернених примірників",
      "Дата знищення",
      "Номер акта знищення",
      "Дата акта знищення",
      "Примітки"
    ],
    subrows:null,
    export:{pageSize:"A4", orientation:"portrait", includeColumns:null}
  },
  { key:"stvoreni_inventarni", title:"Створені інвентарні документи", orderColumn:null,
    columns:[
      {name:"Індекс документу", type:"text", required:false, subrows:false},
      {name:"Дата", type:"date", required:false, subrows:false, defaultToday:true},
      {name:"Кореспондент", type:"text", required:false, subrows:false},
      {name:"Короткий зміст", type:"text", required:false, subrows:false},
      {name:"Відмітка про виконання документа", type:"text", required:false, subrows:false},
      {name:"Примітки", type:"text", required:false, subrows:false}
    ],
    addFields:["Індекс документу","Дата","Кореспондент","Короткий зміст","Відмітка про виконання документа","Примітки"],
    subrows:null,
    export:{pageSize:"A4", orientation:"portrait", includeColumns:null}
  }
];

// v12.1: by default every column allows subrows. Users may disable per column
// in the sheet builder. We normalize legacy schema here.
for(const sh of DEFAULT_SHEETS){
  for(const c of (sh.columns||[])){
    if(c.subrows === false) c.subrows = true;
    if(c.subrows === undefined || c.subrows === null) c.subrows = true;
  }
}
export function normalizeForExport(sheet, row){
  const out = {...(row.data||{})};
  // Subrows: treat the "main" row as subrow #1, then additional subrows follow (by index).
  // For export, we join values with newlines. UI-only numbering is NOT exported as a separate field.
  const sub = row.subrows || [];
  if(sub.length){
    for (const col of (sheet.columns||[]).filter(c=>c.subrows)){
      const mainVal = String(out[col.name] ?? "");
      const vals = [mainVal, ...sub.map(sr=>String(sr?.[col.name] ?? ""))];
      if (col.name==="Номер примірника"){
        out[col.name] = vals.map((_,i)=>String(i+1)).join("\n");
      } else if (col.name==="Кількість примірників"){
        out[col.name] = String(vals.length);
      } else {
        out[col.name] = vals.map(v=>String(v ?? "")).join("\n");
      }
    }
    if (sheet.columns.some(c=>c.name==="Кількість примірників")){
      out["Кількість примірників"]=String(1 + sub.length);
    }
  }
  return out;
}
export function uaDateToday(){
  const d=new Date();
  // UA short format: DD.MM.RR
  return `${pad2(d.getDate())}.${pad2(d.getMonth()+1)}.${pad2(d.getFullYear()%100)}`;
}
// NOTE: regex literals must use single backslashes (\d), not double (\\d).
export function isIntegerString(s){
  const t=String(s??"").trim();
  if(!t) return false;
  return /^\d+$/.test(t);
}

// Parses a date and normalizes it to "DD.MM.RR".
// Accepts: "DD.MM.RR", "DD.MM.RRRR", "RRRR-MM-DD".
// For 2-digit years, assumes 2000-2099.
export function parseUAdate(s){
  const t=String(s??"").trim();
  if(!t) return null;

  // DD.MM.RR or DD.MM.RRRR
  const m1=/^(\d{2})\.(\d{2})\.(\d{2}|\d{4})$/.exec(t);
  if(m1){
    const dd=+m1[1], mm=+m1[2];
    const yRaw=m1[3];
    const yyyy = (yRaw.length===2) ? (2000 + (+yRaw)) : (+yRaw);
    const d=new Date(yyyy, mm-1, dd);
    if(d.getFullYear()===yyyy && d.getMonth()===mm-1 && d.getDate()===dd){
      return `${pad2(dd)}.${pad2(mm)}.${pad2(yyyy%100)}`;
    }
    return null;
  }

  // RRRR-MM-DD
  const m2=/^(\d{4})-(\d{2})-(\d{2})$/.exec(t);
  if(m2){
    const yyyy=+m2[1], mm=+m2[2], dd=+m2[3];
    const d=new Date(yyyy, mm-1, dd);
    if(d.getFullYear()===yyyy && d.getMonth()===mm-1 && d.getDate()===dd){
      return `${pad2(dd)}.${pad2(mm)}.${pad2(yyyy%100)}`;
    }
    return null;
  }

  return null;
}

// Excel serial date (1900 date system) -> "DD.MM.RR".
// Excel counts days from 1899-12-30 (because of the 1900 leap-year bug).
export function excelSerialToUAdate(serial){
  const n = Number(serial);
  if(!Number.isFinite(n)) return null;
  // Ignore unrealistically small numbers
  if(n < 1) return null;
  const ms = (n * 86400000);
  const epoch = Date.UTC(1899, 11, 30);
  const d = new Date(epoch + ms);
  if(Number.isNaN(d.getTime())) return null;
  return `${pad2(d.getUTCDate())}.${pad2(d.getUTCMonth()+1)}.${pad2(d.getUTCFullYear()%100)}`;
}
