// js/export.js
import { downloadBlob } from "./ui.js";
import { nowStamp, safeName, normalizeForExport } from "./schema.js";
import { makeZipStore } from "./zip.js";
import { exportDOCXTable, exportXLSXTable } from "./office.js";
import { exportPDFTable } from "./pdfgen.js";

export function makeJournalExportFileName(title, stamp){
  return `${safeName(title)}_${stamp}.json`;
}
export function makeCaseExportFileName(caseIndex, caseTitle, stamp){
  const label = `${caseIndex||"Без_індексу"}_${caseTitle||"Без_заголовка"}`;
  return `${safeName("Opis_spravy")}_${safeName(label)}_${stamp}.json`;
}

function rowsToFlatObjects(sheet, rows){
  return rows.map(r=>normalizeForExport(sheet,r));
}

function applyRowFilters(flatRows, rowFilters){
  if(!rowFilters || !rowFilters.length) return flatRows;
  const pass = (row)=>{
    for(const f of rowFilters){
      const v = String(row[f.col] ?? "");
      const needle = String(f.value ?? "");
      if(f.op==="contains" && !v.includes(needle)) return false;
      if(f.op==="not_contains" && v.includes(needle)) return false;
      if(f.op==="equals" && v !== needle) return false;
      if(f.op==="not_equals" && v === needle) return false;
    }
    return true;
  };
  return flatRows.filter(pass);
}

export function exportJournalAsJSON({sheet, rows, sheetExportProfile, visibleColumnsForView}){
  const stamp=nowStamp();
  // v2: column-index based backup (cells array in the order of sheet.columns)
  const colNames = (sheet?.columns||[]).map(c=>c.name);
  const rowsV2 = rows.map(r=>({
    id: r.id,
    createdAt: r.createdAt,
    updatedAt: r.updatedAt,
    cells: colNames.map(n=>String(r?.data?.[n] ?? "")),
    subrows: r.subrows||[]
  }));
  const payload = {
    meta:{ type:"journal", version:2, key:sheet.key, title:sheet.title, exportedAt:new Date().toISOString() },
    // keep schema snapshot for reference, but imports should not rely on column names
    sheet,
    columnsCount: colNames.length,
    exportProfile: sheetExportProfile || null,
    // v2 format
    rowsV2,
    // legacy format kept for backward compatibility
    rows: rows.map(r=>({ ...r, exportData: normalizeForExport(sheet,r) }))
  };
  const json = JSON.stringify(payload, null, 2);
  downloadBlob(new Blob([json], {type:"application/json"}), makeJournalExportFileName(sheet.title, stamp));
}

export function exportJournalAsDOCX({sheet, rows, columns, sheetExportProfile}){
  const flat = rowsToFlatObjects(sheet, rows);
  const filtered = applyRowFilters(flat, sheetExportProfile?.rowFilters);
  exportDOCXTable({
    title: sheet.title,
    subtitle: `Експорт: ${new Date().toLocaleString()}`,
    columns,
    rows: filtered,
    filenameBase: sheet.title
  });
}

export function exportJournalAsXLSX({sheet, rows, columns, sheetExportProfile}){
  const flat = rowsToFlatObjects(sheet, rows);
  const filtered = applyRowFilters(flat, sheetExportProfile?.rowFilters);
  exportXLSXTable({
    title: sheet.title,
    columns,
    rows: filtered,
    filenameBase: sheet.title
  });
}

export function exportJournalAsPDF({sheet, rows, columns, sheetExportProfile}){
  const flat = rowsToFlatObjects(sheet, rows);
  const filtered = applyRowFilters(flat, sheetExportProfile?.rowFilters);
  exportPDFTable({
    title: sheet.title,
    subtitle: `Експорт: ${new Date().toLocaleString()}`,
    columns,
    rows: filtered,
    filenameBase: sheet.title,
    pageSize: sheetExportProfile?.pageSize || "A4",
    orientation: sheetExportProfile?.orientation || "portrait"
  });
}

export async function exportAllZipJSON({sheets, allRowsBySheet, cases, caseRowsByCaseId}){
  const stamp = nowStamp();
  const files = [];
  for(const sh of sheets){
    const rows = allRowsBySheet.get(sh.key) || [];
    const colNames = (sh?.columns||[]).map(c=>c.name);
    const rowsV2 = rows.map(r=>({
      id: r.id,
      createdAt: r.createdAt,
      updatedAt: r.updatedAt,
      cells: colNames.map(n=>String(r?.data?.[n] ?? "")),
      subrows: r.subrows||[]
    }));
    const payload = {
      meta:{ type:"journal", version:2, key:sh.key, title:sh.title, exportedAt:new Date().toISOString() },
      sheet: sh,
      columnsCount: colNames.length,
      rowsV2,
      rows: rows.map(r=>({ ...r, exportData: normalizeForExport(sh,r) }))
    };
    const json = JSON.stringify(payload, null, 2);
    files.push({ name: makeJournalExportFileName(sh.title, stamp), data: new TextEncoder().encode(json) });
  }
  for(const c of cases){
    const rows = caseRowsByCaseId.get(c.id) || [];
    const payload = { meta:{ type:"case_description", exportedAt:new Date().toISOString(), case: c }, rows };
    const json=JSON.stringify(payload,null,2);
    files.push({ name: makeCaseExportFileName(c.caseIndex,c.caseTitle,stamp), data: new TextEncoder().encode(json) });
  }
  const zipBytes = makeZipStore(files);
  const blob=new Blob([zipBytes],{type:"application/zip"});
  downloadBlob(blob, `dilovodstvo_full_export_${stamp}.zip`);
}
