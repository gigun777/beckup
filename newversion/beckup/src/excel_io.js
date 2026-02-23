import { unzipEntries, entriesMap } from './zip_read.js';

function decodeUtf8(bytes) {
  return new TextDecoder().decode(bytes);
}

function parseXml(xmlText) {
  if (typeof DOMParser !== 'function') {
    throw new Error('DOMParser is required for XLSX parsing (browser/runtime with DOMParser)');
  }
  return new DOMParser().parseFromString(xmlText, 'application/xml');
}

function selectFirst(nodes, tag) {
  const arr = Array.from(nodes || []);
  return arr.find((n) => n.nodeName === tag || n.localName === tag) || null;
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

function resolveFirstWorksheetPath(files) {
  // Fully independent from file names: resolve by workbook.xml + relationships.
  const workbookPath = 'xl/workbook.xml';
  const relsPath = 'xl/_rels/workbook.xml.rels';
  if (!files.has(workbookPath) || !files.has(relsPath)) {
    throw new Error('XLSX workbook metadata not found');
  }

  const workbookDoc = parseXml(decodeUtf8(files.get(workbookPath)));
  const relsDoc = parseXml(decodeUtf8(files.get(relsPath)));

  const firstSheet = selectFirst(workbookDoc.getElementsByTagName('sheet'), 'sheet');
  if (!firstSheet) throw new Error('No sheet entries in workbook.xml');

  const relId = firstSheet.getAttribute('r:id') || firstSheet.getAttribute('id');
  if (!relId) throw new Error('Worksheet relationship id not found');

  const rels = Array.from(relsDoc.getElementsByTagName('Relationship'));
  const rel = rels.find((r) => r.getAttribute('Id') === relId);
  if (!rel) throw new Error(`Relationship ${relId} not found`);

  const target = rel.getAttribute('Target');
  if (!target) throw new Error('Worksheet target not found');

  // target can be relative like "worksheets/sheet1.xml"
  const normalized = target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`;
  return normalized;
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

function normalizeHeaderName(v) {
  return String(v || '').trim().toLowerCase();
}

/**
 * Parse ANY .xlsx and return rows matrix. Doesn't depend on file names.
 */
export async function parseAnyXlsx(arrayBuffer) {
  const entries = await unzipEntries(arrayBuffer);
  const files = entriesMap(entries);

  const sharedStrings = parseSharedStrings(files.has('xl/sharedStrings.xml') ? decodeUtf8(files.get('xl/sharedStrings.xml')) : null);
  const worksheetPath = resolveFirstWorksheetPath(files);
  if (!files.has(worksheetPath)) throw new Error(`Worksheet file not found: ${worksheetPath}`);

  const rows = parseWorksheetRows(decodeUtf8(files.get(worksheetPath)), sharedStrings);
  return {
    worksheetPath,
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
  headerRowIndex = 0,
  dataRowStartIndex = 1
} = {}) {
  if (!Array.isArray(targetColumns) || !targetColumns.length) {
    throw new Error('targetColumns is required');
  }

  const parsed = await parseAnyXlsx(arrayBuffer);
  const rows = parsed.rows;
  if (!rows.length) return { records: [], mappingUsed: [], warnings: ['Empty worksheet'] };

  const header = rows[headerRowIndex] || [];
  const targetMeta = targetColumns.map((c, i) => {
    if (typeof c === 'string') return { key: c, name: c, index: i };
    return { key: c.key ?? c.name ?? `col_${i + 1}`, name: c.name ?? c.key ?? `col_${i + 1}`, index: i };
  });

  let mappingUsed = [];
  if (Array.isArray(mapping) && mapping.length) {
    // manual mapping format: [{sourceCol:1,targetKey:'name'}] sourceCol is 1-based
    mappingUsed = mapping
      .filter((m) => Number.isFinite(m?.sourceCol) && m.sourceCol >= 1 && m.targetKey)
      .map((m) => ({ sourceCol: m.sourceCol, targetKey: m.targetKey }));
  } else {
    // auto-map by normalized header names
    const byHeader = new Map();
    for (let i = 0; i < header.length; i += 1) {
      byHeader.set(normalizeHeaderName(header[i]), i + 1);
    }
    mappingUsed = targetMeta
      .map((t) => ({ sourceCol: byHeader.get(normalizeHeaderName(t.name)), targetKey: t.key }))
      .filter((x) => Number.isFinite(x.sourceCol));
  }

  const records = [];
  for (let r = dataRowStartIndex; r < rows.length; r += 1) {
    const src = rows[r] || [];
    const cells = {};
    let hasAny = false;

    for (const m of mappingUsed) {
      const v = src[m.sourceCol - 1] ?? '';
      if (String(v).trim() !== '') hasAny = true;
      cells[m.targetKey] = String(v ?? '');
    }

    if (!hasAny) continue;
    records.push({ id: crypto.randomUUID(), cells, subrows: [] });
  }

  return {
    records,
    mappingUsed,
    warnings: mappingUsed.length ? [] : ['No columns mapped']
  };
}
