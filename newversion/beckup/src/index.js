import { makeZipStore, makeZipBlob } from './zip_store.js';
import {
  createSheetJsonPayload,
  createExcelJsonPayload,
  buildSheetJsonFileName,
  buildExcelJsonFileName
} from './sheet_json.js';

function toBytes(x) {
  if (x instanceof Uint8Array) return x;
  if (typeof x === 'string') return new TextEncoder().encode(x);
  if (x instanceof ArrayBuffer) return new Uint8Array(x);
  return new TextEncoder().encode(JSON.stringify(x, null, 2));
}

/**
 * Build a generic ZIP backup archive from logical file entries.
 * @param {{name:string,data:any}[]} entries
 */
export function createZipBackup(entries) {
  const files = (entries || []).map((e) => ({ name: e.name, data: toBytes(e.data) }));
  return makeZipStore(files);
}

/**
 * Build ZIP backup blob (browser-oriented).
 */
export function createZipBackupBlob(entries) {
  const files = (entries || []).map((e) => ({ name: e.name, data: toBytes(e.data) }));
  return makeZipBlob(files);
}

/**
 * Build two-file ZIP for a single sheet:
 * - old-compatible journal json (v2 + legacy rows)
 * - excel-oriented json projection
 */
export function createSheetZipBackup({ sheet, records = [], exportProfile = null }) {
  const journal = createSheetJsonPayload({ sheet, records, exportProfile });
  const excel = createExcelJsonPayload({ sheet, records });
  const files = [
    { name: buildSheetJsonFileName(sheet.title || sheet.name || sheet.key), data: JSON.stringify(journal, null, 2) },
    { name: buildExcelJsonFileName(sheet.title || sheet.name || sheet.key), data: JSON.stringify(excel, null, 2) }
  ];

  return {
    zipBytes: createZipBackup(files),
    files,
    payloads: { journal, excel }
  };
}

export {
  makeZipStore,
  makeZipBlob,
  createSheetJsonPayload,
  createExcelJsonPayload,
  buildSheetJsonFileName,
  buildExcelJsonFileName
};
