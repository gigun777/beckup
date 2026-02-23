function toStringCell(v) {
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  return JSON.stringify(v);
}

function nowStamp(date = new Date()) {
  const pad2 = (n) => String(n).padStart(2, '0');
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}_${pad2(date.getHours())}-${pad2(date.getMinutes())}-${pad2(date.getSeconds())}`;
}

function safeName(s) {
  return String(s || 'export')
    .trim()
    .replace(/[\\/:*?"<>|]+/g, '_')
    .replace(/\s+/g, '_')
    .slice(0, 120) || 'export';
}

function normalizeColumns(sheet) {
  const columns = Array.isArray(sheet?.columns) ? sheet.columns : [];
  return columns.map((c, i) => {
    if (typeof c === 'string') return { key: c, name: c, index: i };
    const name = c?.name ?? c?.label ?? c?.title ?? `col_${i + 1}`;
    const key = c?.key ?? name;
    return { key, name, index: i };
  });
}

function readRecordCell(record, column) {
  if (record?.cells && typeof record.cells === 'object') {
    if (column.key in record.cells) return record.cells[column.key];
    if (column.name in record.cells) return record.cells[column.name];
  }
  if (record?.data && typeof record.data === 'object') {
    if (column.name in record.data) return record.data[column.name];
    if (column.key in record.data) return record.data[column.key];
  }
  if (Array.isArray(record?.cells)) {
    return record.cells[column.index];
  }
  return '';
}

function normalizeRowSubrows(record) {
  if (Array.isArray(record?.subrows)) return record.subrows;
  return [];
}

/**
 * Build old-compatible single-sheet JSON payload (journal v2 + legacy rows).
 */
export function createSheetJsonPayload({ sheet, records = [], exportedAt = new Date().toISOString(), exportProfile = null }) {
  if (!sheet?.key) throw new Error('sheet.key is required');

  const columns = normalizeColumns(sheet);
  const columnNames = columns.map((c) => c.name);

  const rowsV2 = records.map((record) => ({
    id: record?.id ?? crypto.randomUUID(),
    createdAt: record?.createdAt ?? null,
    updatedAt: record?.updatedAt ?? null,
    cells: columns.map((col) => toStringCell(readRecordCell(record, col))),
    subrows: normalizeRowSubrows(record)
  }));

  const legacyRows = records.map((record) => {
    const exportData = {};
    for (const col of columns) exportData[col.name] = toStringCell(readRecordCell(record, col));
    return {
      ...record,
      exportData,
      subrows: normalizeRowSubrows(record)
    };
  });

  return {
    meta: {
      type: 'journal',
      version: 2,
      key: sheet.key,
      title: sheet.title || sheet.name || sheet.key,
      exportedAt
    },
    sheet: {
      ...sheet,
      columns: columnNames.map((name) => ({ name }))
    },
    columnsCount: columns.length,
    exportProfile,
    rowsV2,
    rows: legacyRows
  };
}

/**
 * Excel-oriented JSON for a single sheet.
 * rows are plain objects {colName:value} + matrix for direct tooling usage.
 */
export function createExcelJsonPayload({ sheet, records = [], exportedAt = new Date().toISOString() }) {
  if (!sheet?.key) throw new Error('sheet.key is required');

  const columns = normalizeColumns(sheet);
  const header = columns.map((c) => c.name);

  const rows = records.map((record) => {
    const out = {};
    for (const col of columns) out[col.name] = toStringCell(readRecordCell(record, col));
    return out;
  });

  const matrix = rows.map((row) => header.map((h) => row[h]));

  return {
    format: 'beckup-excel-json',
    formatVersion: 1,
    exportedAt,
    sheet: {
      key: sheet.key,
      title: sheet.title || sheet.name || sheet.key,
      columns: header
    },
    rows,
    matrix
  };
}

export function buildSheetJsonFileName(sheetTitle, stamp = nowStamp()) {
  return `${safeName(sheetTitle || 'sheet')}_${stamp}.json`;
}

export function buildExcelJsonFileName(sheetTitle, stamp = nowStamp()) {
  return `${safeName(sheetTitle || 'sheet')}_${stamp}.excel.json`;
}
