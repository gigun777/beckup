import { unzipEntries, entriesMap } from './zip_read.js';
import { suggestColumnMapping, buildImportPlan, applyImportPlanToRows } from './import_constructor_core.js';

function decodeUtf8(bytes) {
  return new TextDecoder().decode(bytes);
}

function parseXml(xmlText) {
  if (typeof DOMParser !== 'function') {
    throw new Error('DOMParser is required for XLSX parsing (browser/runtime with DOMParser)');
  }
  return new DOMParser().parseFromString(xmlText, 'application/xml');
}

function parseSharedStrings(sharedStringsXml) {
  if (!sharedStringsXml) return [];
  const doc = parseXml(sharedStringsXml);
  const out = [];
  const sis = Array.from(doc.getElementsByTagName('si'));
  for (const si of sis) {
    const ts = Array.from(si.getElementsByTagName('t'));
    out.push(ts.map((x) => x.textContent || '').join(''));
  }
  return out;
}

function extractWorkbookSheets(workbookDoc) {
  const out = [];
  const sheetsNode = workbookDoc.getElementsByTagName('sheets')[0];
  const sheetEls = Array.from(sheetsNode?.getElementsByTagName('sheet') || []);
  for (const el of sheetEls) {
    out.push({
      name: el.getAttribute('name') || '',
      relId: el.getAttribute('r:id') || el.getAttribute('id') || '',
      sheetId: el.getAttribute('sheetId') || ''
    });
  }
  return out;
}

function extractWorkbookRelationships(relsDoc) {
  const rels = Array.from(relsDoc.getElementsByTagName('Relationship'));
  return rels.map((r) => ({
    id: r.getAttribute('Id') || '',
    target: r.getAttribute('Target') || ''
  }));
}

function normalizeWorksheetPath(target) {
  if (!target) return '';
  return target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`;
}

function resolveWorksheetBySelector(files, selector = {}) {
  const workbookPath = 'xl/workbook.xml';
  const relsPath = 'xl/_rels/workbook.xml.rels';
  if (!files.has(workbookPath) || !files.has(relsPath)) {
    throw new Error('XLSX workbook metadata not found');
  }

  const workbookDoc = parseXml(decodeUtf8(files.get(workbookPath)));
  const relsDoc = parseXml(decodeUtf8(files.get(relsPath)));
  const sheets = extractWorkbookSheets(workbookDoc);
  const rels = extractWorkbookRelationships(relsDoc);

  if (!sheets.length) throw new Error('No sheet entries in workbook.xml');

  let picked;
  if (typeof selector?.name === 'string' && selector.name.trim()) {
    picked = sheets.find((s) => s.name === selector.name.trim());
    if (!picked) throw new Error(`Worksheet not found by name: ${selector.name}`);
  } else if (Number.isFinite(selector?.index)) {
    const idx = Number(selector.index);
    if (idx < 0 || idx >= sheets.length) throw new Error(`Worksheet index out of range: ${idx}`);
    picked = sheets[idx];
  } else {
    picked = sheets[0];
  }

  const rel = rels.find((r) => r.id === picked.relId);
  if (!rel) throw new Error(`Relationship not found for worksheet: ${picked.name || picked.relId}`);

  const worksheetPath = normalizeWorksheetPath(rel.target);
  if (!worksheetPath || !files.has(worksheetPath)) {
    throw new Error(`Worksheet file not found: ${worksheetPath}`);
  }

  return {
    worksheetPath,
    sheetName: picked.name || '',
    sheetIndex: sheets.indexOf(picked),
    sheets: sheets.map((s, i) => ({ name: s.name, index: i }))
  };
}

function colLettersToIndex(ref) {
  const m = /^([A-Z]+)\d+$/.exec(ref || '');
  if (!m) return null;
  const letters = m[1];
  let n = 0;
  for (let i = 0; i < letters.length; i += 1) n = (n * 26) + (letters.charCodeAt(i) - 64);
  return n;
}

function getCellText(cellEl, sharedStrings) {
  const t = cellEl.getAttribute('t') || '';
  if (t === 'inlineStr') {
    const tEl = cellEl.getElementsByTagName('t')[0];
    return tEl ? (tEl.textContent || '') : '';
  }
  const vEl = cellEl.getElementsByTagName('v')[0];
  const v = vEl ? (vEl.textContent || '') : '';
  if (t === 's') {
    const idx = parseInt(v, 10);
    return Number.isFinite(idx) && sharedStrings[idx] != null ? sharedStrings[idx] : '';
  }
  return v;
}

function parseWorksheetRows(worksheetXml, sharedStrings) {
  const doc = parseXml(worksheetXml);
  const rowEls = Array.from(doc.getElementsByTagName('row'));
  const out = [];

  for (const rowEl of rowEls) {
    const cells = Array.from(rowEl.getElementsByTagName('c'));
    if (!cells.length) continue;

    let max = 0;
    const map = new Map();
    for (const c of cells) {
      const ref = c.getAttribute('r') || '';
      const idx1 = colLettersToIndex(ref);
      if (!idx1) continue;
      if (idx1 > max) max = idx1;
      map.set(idx1, getCellText(c, sharedStrings));
    }

    const arr = [];
    for (let i = 1; i <= max; i += 1) arr.push(map.get(i) ?? '');
    out.push(arr);
  }

  return out;
}


/**
 * Parse ANY .xlsx and return rows matrix.
 * Doesn't depend on workbook file naming conventions.
 */
export async function parseAnyXlsx(arrayBuffer, { worksheet = {} } = {}) {
  const entries = await unzipEntries(arrayBuffer);
  const files = entriesMap(entries);

  const sharedStrings = parseSharedStrings(files.has('xl/sharedStrings.xml') ? decodeUtf8(files.get('xl/sharedStrings.xml')) : null);
  const picked = resolveWorksheetBySelector(files, worksheet);

  const rows = parseWorksheetRows(decodeUtf8(files.get(picked.worksheetPath)), sharedStrings);
  return {
    worksheetPath: picked.worksheetPath,
    worksheetName: picked.sheetName,
    worksheetIndex: picked.sheetIndex,
    worksheets: picked.sheets,
    rows,
    rowCount: rows.length,
    colCount: rows.reduce((m, r) => Math.max(m, r.length), 0)
  };
}

/**
 * Import from ANY excel into journal records.
 * - mapping can be manual (source column index -> target key)
 * - if mapping absent, tries auto-map by header names.
 */
export async function importAnyExcelToRecords({
  arrayBuffer,
  targetColumns,
  mapping = null,
  worksheet = {},
  headerRowIndex = 0,
  dataRowStartIndex = 1
} = {}) {
  if (!Array.isArray(targetColumns) || !targetColumns.length) {
    throw new Error('targetColumns is required');
  }

  const parsed = await parseAnyXlsx(arrayBuffer, { worksheet });
  const rows = parsed.rows;
  if (!rows.length) {
    return {
      records: [],
      mappingUsed: [],
      worksheet: {
        name: parsed.worksheetName,
        index: parsed.worksheetIndex,
        path: parsed.worksheetPath
      },
      warnings: ['Empty worksheet']
    };
  }

  const header = rows[headerRowIndex] || [];

  const mappingUsed = (Array.isArray(mapping) && mapping.length)
    ? mapping
    : suggestColumnMapping({ headerRow: header, targetColumns });

  const { plan, warnings: planWarnings } = buildImportPlan({ mapping: mappingUsed, targetColumns });
  const records = applyImportPlanToRows({ rows, plan, dataRowStartIndex });

  return {
    records,
    mappingUsed: plan,
    worksheet: {
      name: parsed.worksheetName,
      index: parsed.worksheetIndex,
      path: parsed.worksheetPath
    },
    warnings: plan.length ? planWarnings : [...planWarnings, 'No columns mapped']
  };
}
