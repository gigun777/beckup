function isObj(x) { return x && typeof x === 'object' && !Array.isArray(x); }

export function normalizeBackupBundle(bundle) {
  if (!isObj(bundle)) throw new Error('Backup must be an object');

  const format = String(bundle.format || 'beckup-full-json');
  const formatVersion = Number(bundle.formatVersion || 1);
  const sections = isObj(bundle.sections) ? bundle.sections : {};

  return {
    format,
    formatVersion,
    createdAt: bundle.createdAt || new Date().toISOString(),
    app: isObj(bundle.app) ? bundle.app : { name: '@beckup/beckup', version: 'unknown' },
    sections
  };
}

export function normalizeJournalPayload(item) {
  if (!isObj(item)) return { ok: false, reason: 'journal item must be object' };
  if (item?.meta?.type !== 'journal') return { ok: false, reason: 'meta.type is not journal' };

  const key = item?.meta?.key || item?.sheet?.key || null;
  if (!key) return { ok: false, reason: 'journal key missing' };

  const sheetColumns = Array.isArray(item?.sheet?.columns)
    ? item.sheet.columns.map((c, i) => {
      if (typeof c === 'string') return { name: c, index: i };
      return { name: c?.name || c?.key || `col_${i + 1}`, index: i };
    })
    : [];

  const rowsV2 = Array.isArray(item.rowsV2) ? item.rowsV2 : [];
  const rowsLegacy = Array.isArray(item.rows) ? item.rows : [];

  return {
    ok: true,
    payload: {
      ...item,
      meta: {
        ...(isObj(item.meta) ? item.meta : {}),
        type: 'journal',
        version: Number(item?.meta?.version || 2),
        key,
        title: item?.meta?.title || item?.sheet?.title || key
      },
      sheet: {
        ...(isObj(item.sheet) ? item.sheet : {}),
        key,
        title: item?.sheet?.title || item?.meta?.title || key,
        columns: sheetColumns
      },
      rowsV2,
      rows: rowsLegacy,
      columnsCount: Number(item?.columnsCount || sheetColumns.length)
    }
  };
}

export function mergeReports(base, next) {
  return {
    applied: Number(base?.applied || 0) + Number(next?.applied || 0),
    skipped: Number(base?.skipped || 0) + Number(next?.skipped || 0),
    warnings: [...(base?.warnings || []), ...(next?.warnings || [])],
    errors: [...(base?.errors || []), ...(next?.errors || [])]
  };
}
