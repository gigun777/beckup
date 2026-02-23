import { VERSION, BACKUP_FORMAT, DELTA_BACKUP_FORMAT, ENCRYPTED_BACKUP_FORMAT } from './types/public.js';
import { assertStorage, createMemoryStorage, createLocalStorageStorage } from './storage/storage_iface.js';
import { NAV_KEYS, loadNavigationState, saveNavigationState } from './storage/db_nav.js';
import { normalizeLocation } from './core/level_model_core.js';
import { pushHistory } from './core/navigation_core.js';
import { createUIRegistry } from './core/ui_registry_core.js';
import { createSchemaRegistry } from './core/schema_registry_core.js';
import { createCommandsRegistry } from './core/commands_registry_core.js';
import { createSettingsRegistry } from './core/settings_registry_core.js';
import { createJournalTemplatesContainer } from './stores/journal_templates_container.js';
import { createIntegrity, decryptBackup, encryptBackup, verifyIntegrity } from './backup/crypto.js';
import { createTableStoreModule } from './modules/table_store.js';
export { assertStorage, createMemoryStorage, createLocalStorageStorage } from './storage/storage_iface.js';


function deepFreeze(obj) {
  if (!obj || typeof obj !== 'object') return obj;
  Object.freeze(obj);
  for (const value of Object.values(obj)) deepFreeze(value);
  return obj;
}

function toNavPayload(nav) {
  return {
    spaces_nodes_v2: nav.spaces ?? [],
    journals_nodes_v2: nav.journals ?? [],
    nav_last_loc_v2: nav.lastLoc ?? null,
    nav_history_v2: nav.history ?? []
  };
}

function fromNavPayload(payload) {
  return {
    spaces: payload.spaces_nodes_v2 ?? [],
    journals: payload.journals_nodes_v2 ?? [],
    lastLoc: payload.nav_last_loc_v2 ?? null,
    history: payload.nav_history_v2 ?? []
  };
}

// ------------------------------------------------------------------------
// Excel import/export helper functions.
// These functions implement minimal XLSX generation and parsing with ZIP store.
// They are inspired by older SEDO versions but extended to handle multiple sheets.
// The helper functions are defined in the module scope so that createSEDO can
// capture them in closures.

// Escape XML special characters for spreadsheet strings.
function excelXmlEsc(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Convert 1-based column index to Excel column letters (e.g. 1 -> A, 27 -> AA).
function excelColLetter(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Check if a value represents a number (integer or decimal). Strings containing only digits (with optional sign and decimal point) are considered numbers.
function excelIsNumber(v) {
  const s = String(v ?? '').trim();
  return /^-?\d+(?:\.\d+)?$/.test(s);
}

// Build worksheet XML for a single sheet using inline strings.
// sheetName is ignored here; names are defined in workbook.xml.
function excelBuildWorksheetXml(columns, rows) {
  let sheetRows = '';
  // Header row (row 1)
  const headerCells = columns.map((c, ci) => {
    const addr = excelColLetter(ci + 1) + '1';
    return `<c r="${addr}" t="inlineStr"><is><t xml:space="preserve">${excelXmlEsc(c)}</t></is></c>`;
  }).join('');
  sheetRows += `<row r="1">${headerCells}</row>`;
  // Data rows
  for (let ri = 0; ri < rows.length; ri++) {
    const rIndex = ri + 2;
    const row = rows[ri] ?? {};
    const cells = columns.map((c, ci) => {
      const addr = excelColLetter(ci + 1) + String(rIndex);
      const v = row[c] ?? '';
      if (excelIsNumber(v)) {
        return `<c r="${addr}" t="n"><v>${String(v).trim()}</v></c>`;
      }
      return `<c r="${addr}" t="inlineStr"><is><t xml:space="preserve">${excelXmlEsc(v)}</t></is></c>`;
    }).join('');
    sheetRows += `<row r="${rIndex}">${cells}</row>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>${sheetRows}</sheetData></worksheet>`;
}

// ZIP writer helpers derived from the minimal zip store implementation.
function excelU16(n) { return new Uint8Array([n & 255, (n >>> 8) & 255]); }
function excelU32(n) { return new Uint8Array([n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]); }
function excelConcatBytes(chunks) {
  const total = chunks.reduce((s, c) => s + c.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) { out.set(c, off); off += c.length; }
  return out;
}
const EXCEL_CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    table[i] = c >>> 0;
  }
  return table;
})();
function excelCrc32(bytes) {
  let crc = 0 ^ (-1);
  for (let i = 0; i < bytes.length; i++) {
    crc = (crc >>> 8) ^ EXCEL_CRC_TABLE[(crc ^ bytes[i]) & 0xFF];
  }
  return (crc ^ (-1)) >>> 0;
}
function excelDosTimeDate(date) {
  const d = date || new Date();
  let time = 0;
  time |= ((Math.floor(d.getSeconds() / 2)) & 31);
  time |= (d.getMinutes() & 63) << 5;
  time |= (d.getHours() & 31) << 11;
  let dt = 0;
  dt |= (d.getDate() & 31);
  dt |= ((d.getMonth() + 1) & 15) << 5;
  dt |= ((d.getFullYear() - 1980) & 127) << 9;
  return { time: time & 0xFFFF, date: dt & 0xFFFF };
}
function excelMakeZipStore(files) {
  const localParts = [], centralParts = [];
  let offset = 0;
  const { time, date } = excelDosTimeDate(new Date());
  for (const f of files) {
    const nameBytes = new TextEncoder().encode(f.name);
    const dataBytes = f.data instanceof Uint8Array ? f.data : new Uint8Array(f.data);
    const c = excelCrc32(dataBytes);
    const localHeader = excelConcatBytes([
      excelU32(0x04034b50), excelU16(20), excelU16(0), excelU16(0),
      excelU16(time), excelU16(date),
      excelU32(c), excelU32(dataBytes.length), excelU32(dataBytes.length),
      excelU16(nameBytes.length), excelU16(0)
    ]);
    localParts.push(localHeader, nameBytes, dataBytes);
    const centralHeader = excelConcatBytes([
      excelU32(0x02014b50),
      excelU16(20), excelU16(20),
      excelU16(0), excelU16(0),
      excelU16(time), excelU16(date),
      excelU32(c), excelU32(dataBytes.length), excelU32(dataBytes.length),
      excelU16(nameBytes.length),
      excelU16(0), excelU16(0),
      excelU16(0), excelU16(0),
      excelU32(0),
      excelU32(offset)
    ]);
    centralParts.push(centralHeader, nameBytes);
    offset += localHeader.length + nameBytes.length + dataBytes.length;
  }
  const centralDir = excelConcatBytes(centralParts);
  const localData = excelConcatBytes(localParts);
  const end = excelConcatBytes([
    excelU32(0x06054b50),
    excelU16(0), excelU16(0),
    excelU16(files.length), excelU16(files.length),
    excelU32(centralDir.length),
    excelU32(localData.length),
    excelU16(0)
  ]);
  return excelConcatBytes([localData, centralDir, end]);
}

// Minimal unzip (STORE/DEFLATE) to parse XLSX files during import.
function excelReadU16(dv, o) { return dv.getUint16(o, true); }
function excelReadU32(dv, o) { return dv.getUint32(o, true); }
function excelFindEOCD(dv) {
  const sig = 0x06054b50;
  const maxBack = Math.min(dv.byteLength, 22 + 0xFFFF);
  for (let i = dv.byteLength - 22; i >= dv.byteLength - maxBack; i--) {
    if (i < 0) break;
    if (excelReadU32(dv, i) === sig) return i;
  }
  return -1;
}
async function excelInflateRawBytes(u8) {
  if (typeof DecompressionStream !== 'function') {
    throw new Error('ZIP: DecompressionStream missing for DEFLATE');
  }
  const tryAlg = async (alg) => {
    const ds = new DecompressionStream(alg);
    const ab = await new Response(new Blob([u8]).stream().pipeThrough(ds)).arrayBuffer();
    return new Uint8Array(ab);
  };
  try { return await tryAlg('deflate-raw'); }
  catch (_e) { return await tryAlg('deflate'); }
}
async function excelUnzipEntries(arrayBuffer) {
  const dv = new DataView(arrayBuffer);
  const eocdOff = excelFindEOCD(dv);
  if (eocdOff < 0) throw new Error('ZIP: EOCD not found');
  const cdSize = excelReadU32(dv, eocdOff + 12);
  const cdOff = excelReadU32(dv, eocdOff + 16);
  let p = cdOff;
  const files = [];
  while (p < cdOff + cdSize) {
    const sig = excelReadU32(dv, p);
    if (sig !== 0x02014b50) throw new Error('ZIP: Central Directory broken');
    const compMethod = excelReadU16(dv, p + 10);
    const compSize = excelReadU32(dv, p + 20);
    const uncompSize = excelReadU32(dv, p + 24);
    const nameLen = excelReadU16(dv, p + 28);
    const extraLen = excelReadU16(dv, p + 30);
    const commentLen = excelReadU16(dv, p + 32);
    const localOff = excelReadU32(dv, p + 42);
    const nameBytes = new Uint8Array(arrayBuffer, p + 46, nameLen);
    const name = new TextDecoder().decode(nameBytes);
    const lsig = excelReadU32(dv, localOff);
    if (lsig !== 0x04034b50) throw new Error('ZIP: Local Header broken');
    const lNameLen = excelReadU16(dv, localOff + 26);
    const lExtraLen = excelReadU16(dv, localOff + 28);
    const dataOff = localOff + 30 + lNameLen + lExtraLen;
    const compData = new Uint8Array(arrayBuffer, dataOff, compSize);
    let data;
    if (compMethod === 0) data = compData;
    else if (compMethod === 8) {
      data = await excelInflateRawBytes(compData);
      if (uncompSize && data.length !== uncompSize) { /* size mismatch tolerated */ }
    } else {
      throw new Error(`ZIP: unsupported compression method ${compMethod} for ${name}`);
    }
    files.push({ name, data });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return files;
}

// Parse shared strings xml into array.
function excelParseSharedStringsXml(xml) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const siList = doc.getElementsByTagName('si');
  const out = [];
  for (let i = 0; i < siList.length; i++) {
    const tEl = siList[i].getElementsByTagName('t')[0];
    out.push(tEl ? tEl.textContent || '' : '');
  }
  return out;
}

// Extract cell text from <c> element using shared strings array.
function excelGetCellTextFromXml(cellEl, sharedStrings) {
  if (!cellEl) return '';
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

// Convert cell reference like "AA10" to column index (1-based).
function excelColLettersToIndex(ref) {
  const m = /^([A-Z]+)\d+$/.exec(ref || '');
  if (!m) return null;
  const s = m[1];
  let n = 0;
  for (let i = 0; i < s.length; i++) { n = n * 26 + (s.charCodeAt(i) - 64); }
  return n;
}

// Parse workbook from unzipped entries. Returns array of sheets with name, columns, rows.
async function excelParseWorkbook(entries) {
  const map = new Map();
  for (const f of entries) map.set(f.name, f);
  let sharedStrings = [];
  const sstEntry = map.get('xl/sharedStrings.xml');
  if (sstEntry) sharedStrings = excelParseSharedStringsXml(new TextDecoder().decode(sstEntry.data));
  const wbEntry = map.get('xl/workbook.xml');
  if (!wbEntry) throw new Error('XLSX: workbook.xml not found');
  const wbDoc = new DOMParser().parseFromString(new TextDecoder().decode(wbEntry.data), 'application/xml');
  const sheetEls = wbDoc.getElementsByTagName('sheet');
  const sheetsInfo = [];
  for (let i = 0; i < sheetEls.length; i++) {
    const name = sheetEls[i].getAttribute('name') || `Sheet${i + 1}`;
    const rId = sheetEls[i].getAttribute('r:id');
    sheetsInfo.push({ name, rId });
  }
  const relsEntry = map.get('xl/_rels/workbook.xml.rels');
  const relsMap = new Map();
  if (relsEntry) {
    const relsDoc = new DOMParser().parseFromString(new TextDecoder().decode(relsEntry.data), 'application/xml');
    const relEls = relsDoc.getElementsByTagName('Relationship');
    for (let i = 0; i < relEls.length; i++) {
      const id = relEls[i].getAttribute('Id');
      const target = relEls[i].getAttribute('Target');
      if (id && target) relsMap.set(id, target);
    }
  }
  const result = [];
  for (const info of sheetsInfo) {
    const target = relsMap.get(info.rId) || `worksheets/sheet${result.length + 1}.xml`;
    const entryName = `xl/${target.replace(/^\/+/, '')}`;
    const sheetEntry = map.get(entryName);
    if (!sheetEntry) continue;
    const sheetDoc = new DOMParser().parseFromString(new TextDecoder().decode(sheetEntry.data), 'application/xml');
    const rowsEls = sheetDoc.getElementsByTagName('row');
    let columns = [];
    const rows = [];
    for (let i = 0; i < rowsEls.length; i++) {
      const rowEl = rowsEls[i];
      const cellEls = Array.from(rowEl.getElementsByTagName('c'));
      const rowObj = {};
      if (i === 0) {
        for (const cellEl of cellEls) {
          const ref = cellEl.getAttribute('r');
          const colIdx = excelColLettersToIndex(ref);
          const header = excelGetCellTextFromXml(cellEl, sharedStrings);
          while (columns.length < colIdx) columns.push('');
          columns[colIdx - 1] = header;
        }
      } else {
        for (const cellEl of cellEls) {
          const ref = cellEl.getAttribute('r');
          const colIdx = excelColLettersToIndex(ref);
          const value = excelGetCellTextFromXml(cellEl, sharedStrings);
          const colName = columns[colIdx - 1];
          if (colName) rowObj[colName] = value;
        }
        rows.push(rowObj);
      }
    }
    result.push({ name: info.name, columns, rows });
  }
  return result;
}

// Build a complete workbook (ZIP) from multiple sheets.
function excelBuildWorkbook(sheets) {
  const entries = [];
  const te = new TextEncoder();
  const workbookSheets = [];
  const workbookRels = [];
  let sheetId = 1;
  for (const sheet of sheets) {
    const xml = excelBuildWorksheetXml(sheet.columns, sheet.rows);
    const fileName = `xl/worksheets/sheet${sheetId}.xml`;
    entries.push({ name: fileName, data: te.encode(xml) });
    workbookSheets.push({ name: sheet.name, id: sheetId });
    workbookRels.push({ id: `rId${sheetId}`, target: `worksheets/sheet${sheetId}.xml` });
    sheetId++;
  }
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${workbookSheets.map((s) => `<sheet name="${excelXmlEsc(s.name)}" sheetId="${s.id}" r:id="rId${s.id}"/>`).join('')}</sheets></workbook>`;
  entries.push({ name: 'xl/workbook.xml', data: te.encode(workbookXml) });
  const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${workbookRels.map((rel) => `<Relationship Id="${rel.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${rel.target}"/>`).join('')}</Relationships>`;
  entries.push({ name: 'xl/_rels/workbook.xml.rels', data: te.encode(workbookRelsXml) });
  const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  entries.push({ name: '_rels/.rels', data: te.encode(rootRelsXml) });
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${workbookSheets.map((s) => `<Override PartName="/xl/worksheets/sheet${s.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`;
  entries.push({ name: '[Content_Types].xml', data: te.encode(contentTypesXml) });
  return excelMakeZipStore(entries);
}

// ------------------------------------------------------------------------

export function createSEDO(options = {}) {
  const storage = options.storage ?? createMemoryStorage();
  assertStorage(storage);

  const listeners = new Map();
  const modules = new Map();
  const backupProviders = new Map();
  const moduleDisposers = new Map();
  const uiRegistry = createUIRegistry();
  const schemaRegistry = createSchemaRegistry();
  const settingsRegistry = createSettingsRegistry();
  const journalTemplates = createJournalTemplatesContainer(storage);

  const state = {
    spaces: [],
    journals: [],
    history: [],
    activeSpaceId: null,
    activeJournalId: null,
    started: false,
    revision: 0
  };

  function getRuntimeCtx() {
    return { api, storage, sdo: instance };
  }
  const commandsRegistry = createCommandsRegistry(getRuntimeCtx);

  let ui = null;

  function emit(event, payload) {
    for (const fn of listeners.get(event) ?? []) fn(payload);
  }

  async function bumpRevision(changedKeys = []) {
    state.revision += 1;
    await storage.set(NAV_KEYS.revision, state.revision);
    const log = (await storage.get(NAV_KEYS.revisionLog)) ?? [];
    log.push({ rev: state.revision, changedKeys, at: new Date().toISOString() });
    await storage.set(NAV_KEYS.revisionLog, log.slice(-500));
  }

  const api = {
    getState: () => deepFreeze(structuredClone(state)),
    dispatch(action) {
      if (typeof action?.reduce !== 'function') throw new Error('Action must include reduce(state)');
      action.reduce(state);
      emit('state:changed', api.getState());
      return state;
    }
  };

  function createModuleUIApi(moduleId) {
    const disposers = moduleDisposers.get(moduleId) ?? [];
    moduleDisposers.set(moduleId, disposers);
    function track(unregisterFn) {
      disposers.push(unregisterFn);
      return () => {
        unregisterFn();
        const idx = disposers.indexOf(unregisterFn);
        if (idx >= 0) disposers.splice(idx, 1);
      };
    }
    return {
      registerButton(def) { return track(uiRegistry.registerButton({ ...def })); },
      registerPanel(def) { return track(uiRegistry.registerPanel({ ...def })); },
      listButtons(filter) { return uiRegistry.listButtons(filter); },
      listPanels(filter) { return uiRegistry.listPanels(filter); }
    };
  }

  function createModuleCtx(moduleId) {
    const disposers = moduleDisposers.get(moduleId) ?? [];
    moduleDisposers.set(moduleId, disposers);
    const track = (fn) => {
      disposers.push(fn);
      return () => {
        fn();
        const idx = disposers.indexOf(fn);
        if (idx >= 0) disposers.splice(idx, 1);
      };
    };

    return {
      api,
      storage,
      ui: createModuleUIApi(moduleId),
      registerSchema(schemaDef) { return track(schemaRegistry.register(schemaDef)); },
      registerCommands(commandDefs) { return track(commandsRegistry.register(commandDefs)); },
      registerSettings(settingsDef) { return track(settingsRegistry.register(settingsDef)); },
      schemas: {
        get: (id) => schemaRegistry.get(id),
        list: (filter) => schemaRegistry.list(filter),
        resolve: (target) => schemaRegistry.resolve(target)
      },
      commands: {
        run: (id, args) => commandsRegistry.run(id, args),
        list: (filter) => commandsRegistry.list(filter)
      },
      settings: {
        listTabs: () => settingsRegistry.listTabs(),
        getKey: (key) => storage.get(key),
        setKey: (key, value) => storage.set(key, value)
      },
      backup: {
        registerProvider(provider) {
          if (!provider?.id || typeof provider.export !== 'function' || typeof provider.import !== 'function' || typeof provider.describe !== 'function') {
            throw new Error('Backup provider must include id/describe/export/import');
          }
          backupProviders.set(provider.id, provider);
          return track(() => backupProviders.delete(provider.id));
        }
      }
    };
  }

  const instance = {
    version: VERSION,
    api,
    ui: {
      listButtons: (filter) => uiRegistry.listButtons(filter),
      listPanels: (filter) => uiRegistry.listPanels(filter),
      subscribe: (handler) => uiRegistry.subscribe(handler)
    },
    schemas: {
      get: (id) => schemaRegistry.get(id),
      list: (filter) => schemaRegistry.list(filter),
      resolve: (target) => schemaRegistry.resolve(target)
    },
    commands: {
      run: (id, args) => commandsRegistry.run(id, args),
      list: (filter) => commandsRegistry.list(filter)
    },
    settings: {
      listTabs: () => settingsRegistry.listTabs(),
      getKey: (key) => storage.get(key),
      setKey: (key, value) => storage.set(key, value)
    },
    journalTemplates: {
      listTemplates: () => journalTemplates.listTemplates(),
      listTemplateEntities: () => journalTemplates.listTemplateEntities(),
      getTemplate: (id) => journalTemplates.getTemplate(id),
      addTemplate: (template) => journalTemplates.addTemplate(template),
      deleteTemplate: (id) => journalTemplates.deleteTemplate(id),
      exportDelta: (sinceRevision = 0) => journalTemplates.exportDelta(sinceRevision),
      applyDelta: (patch) => journalTemplates.applyDelta(patch)
    },
    use(module) {
      if (!module?.id || typeof module?.init !== 'function') throw new Error('Invalid module');
      if (modules.has(module.id)) return instance;
      const ctx = createModuleCtx(module.id);
      module.init(ctx);
      modules.set(module.id, module);
      emit('module:used', module.id);
      return instance;
    },
    async loadModuleFromUrl(url) {
      const mod = await import(url);
      const plugin = mod.default ?? mod.module ?? mod;
      instance.use(plugin);
      return plugin;
    },
    async start() {
      await journalTemplates.ensureInitialized();
      backupProviders.set('journal-templates', {
        id: 'journal-templates',
        version: '0.1.0',
        describe: () => ({ settings: ['templates:*'], userData: [] }),
        export: async () => ({ templates: await journalTemplates.listTemplateEntities() }),
        import: async (payload) => {
          for (const template of payload.templates ?? []) {
            await journalTemplates.deleteTemplate(template.id);
            await journalTemplates.addTemplate(template);
          }
          return { applied: true, warnings: [] };
        }
      });

// --- Custom backup provider for transfer templates ---
{
  // Dynamically import transfer core to avoid loading overhead on startup.  The transfer core
  // exposes loadTemplates()/saveTemplates() for reading and writing full template arrays.  We
  // always clone the returned array on export to avoid accidental mutation.
  const { createTransferCore } = await import('./core/transfer_core.js');
  const transferCore = createTransferCore({ storage });
  backupProviders.set('transfer-templates', {
    id: 'transfer-templates',
    version: '1.0.0',
    describe: () => ({ settings: [], userData: ['transfer:templates:v1'] }),
    export: async () => {
      const templates = await transferCore.loadTemplates();
      return { templates: Array.isArray(templates) ? [...templates] : [] };
    },
    import: async (payload, opts = {}) => {
      const newTemplates = Array.isArray(payload?.templates) ? payload.templates : [];
      const mode = opts.mode ?? 'merge';
      let existing = await transferCore.loadTemplates();
      if (!Array.isArray(existing)) existing = [];
      if (mode === 'replace') {
        existing = [];
      }
      const byId = new Map();
      for (const tpl of existing) {
        if (tpl && typeof tpl.id === 'string') {
          byId.set(tpl.id, { ...tpl });
        }
      }
      for (const tpl of newTemplates) {
        if (!tpl || typeof tpl.id !== 'string') continue;
        const prev = byId.get(tpl.id) ?? {};
        byId.set(tpl.id, { ...prev, ...tpl });
      }
      const merged = Array.from(byId.values());
      await transferCore.saveTemplates(merged);
      return { applied: true, warnings: [] };
    }
  });
}

      
// --- Backup provider for table column settings ---
{
  const TABLE_SETTINGS_KEY = '@sdo/module-table-renderer:settings';
  backupProviders.set('table-settings', {
    id: 'table-settings',
    version: '1.0.0',
    describe: async () => {
      const nav = await loadNavigationState(storage);
      const journalIds = Array.isArray(nav?.journals) ? nav.journals.map((j) => j.id) : [];
      const settingsKeys = [TABLE_SETTINGS_KEY, ...journalIds.map((id) => `${TABLE_SETTINGS_KEY}:${id}`)];
      return { settings: settingsKeys, userData: [] };
    },
    export: async () => {
      const nav = await loadNavigationState(storage);
      const journalIds = Array.isArray(nav?.journals) ? nav.journals.map((j) => j.id) : [];
      const settingsKeys = [TABLE_SETTINGS_KEY, ...journalIds.map((id) => `${TABLE_SETTINGS_KEY}:${id}`)];
      const data = {};
      for (const key of settingsKeys) {
        const val = await storage.get(key);
        if (val !== undefined) data[key] = val;
      }
      return { settings: data };
    },
    import: async (payload) => {
      const data = payload?.settings ?? {};
      for (const [key, value] of Object.entries(data)) {
        await storage.set(key, value);
      }
      return { applied: true, warnings: [] };
    }
  });
}
const nav = await loadNavigationState(storage);
      state.spaces = nav.spaces;
      state.journals = nav.journals;
      state.history = nav.history;
      
      // MIGRATION: ensure every journal has templateId (old test journals may not have it)
      try {
        const tpls = await journalTemplates.listTemplateEntities();
        const defaultTplId = (tpls.find((t) => t.id === 'test')?.id) || (tpls[0]?.id) || null;
        if (defaultTplId) {
          let changed = false;
          state.journals = (state.journals || []).map((j) => {
            if (j && !j.templateId) { changed = true; return { ...j, templateId: defaultTplId }; }
            return j;
          });
          if (changed) {
            await saveNavigationState(storage, { spaces: state.spaces, journals: state.journals, history: state.history, lastLoc: nav.lastLoc });
          }
        }
      } catch (e) {
        // ignore migration errors
      }

const loc = normalizeLocation({ spaces: state.spaces, journals: state.journals, lastLoc: nav.lastLoc });
      state.activeSpaceId = loc.activeSpaceId;
      state.activeJournalId = loc.activeJournalId;
      state.started = true;
      if (options.mount && typeof options.createUI === 'function') {
        ui = options.createUI({ sdo: instance, mount: options.mount, api });
      }
      emit('started', api.getState());
      return instance;
    },
    async destroy() {
      for (const [moduleId, disposers] of moduleDisposers.entries()) {
        for (const dispose of disposers.splice(0)) dispose();
        const module = modules.get(moduleId);
        if (typeof module?.destroy === 'function') await module.destroy();
      }
      uiRegistry.clear();
      schemaRegistry.clear();
      settingsRegistry.clear();
      commandsRegistry.clear();
      ui?.destroy?.();
      listeners.clear();
      modules.clear();
      backupProviders.clear();
      moduleDisposers.clear();
    },
    getState: api.getState,
    async commit(mutator, changedKeys = []) {
      mutator(state);
      state.history = pushHistory(state.history, {
        activeSpaceId: state.activeSpaceId,
        activeJournalId: state.activeJournalId,
        at: new Date().toISOString()
      });
      await saveNavigationState(storage, {
        spaces: state.spaces,
        journals: state.journals,
        lastLoc: { activeSpaceId: state.activeSpaceId, activeJournalId: state.activeJournalId },
        history: state.history
      });
      await bumpRevision(changedKeys);
      emit('state:changed', api.getState());
    },
    async exportNavigationState() {
      return toNavPayload(await loadNavigationState(storage));
    },
    async importNavigationState(payload) {
      const nav = fromNavPayload(payload);
      await saveNavigationState(storage, nav);
      const loc = normalizeLocation({ spaces: nav.spaces, journals: nav.journals, lastLoc: nav.lastLoc });
      state.spaces = nav.spaces;
      state.journals = nav.journals;
      state.history = nav.history;
      state.activeSpaceId = loc.activeSpaceId;
      state.activeJournalId = loc.activeJournalId;
      return { applied: true, warnings: [] };
    },
    async exportBackup(opts = {}) {
      const scope = opts.scope ?? 'all';
      const backupId = crypto.randomUUID();
      const bundle = {
        format: BACKUP_FORMAT,
        formatVersion: 1,
        backupId,
        createdAt: new Date().toISOString(),
        app: { name: '@sdo/core', version: VERSION },
        scope,
        core: { navigation: null, settings: { coreSettings: (await storage.get(NAV_KEYS.coreSettings)) ?? {} } },
        modules: {},
        userData: {}
      };
      if (opts.includeNavigation !== false && (scope === 'all' || scope === 'userData' || scope === 'modules')) {
        bundle.core.navigation = await instance.exportNavigationState();
      }
      const moduleIds = opts.modules ?? [...backupProviders.keys()];
      for (const id of moduleIds) {
        const provider = backupProviders.get(id);
        if (!provider) continue;
        bundle.modules[id] = { moduleVersion: provider.version, data: await provider.export({ includeUserData: opts.includeUserData !== false, scope }) };
      }
      bundle.integrity = await createIntegrity(bundle);
      return opts.encrypt?.enabled ? encryptBackup(bundle, opts.encrypt.password) : bundle;
    },
    async importBackup(input, opts = {}) {
      const bundle = input?.format === ENCRYPTED_BACKUP_FORMAT ? await decryptBackup(input, opts.decrypt?.password ?? '') : input;
      if (bundle?.format !== BACKUP_FORMAT) throw new Error('Unsupported backup format');
      if (!await verifyIntegrity(bundle)) throw new Error('Backup integrity check failed');

      const report = { core: { applied: false, warnings: [] }, navigation: { applied: false, warnings: [] }, modules: {} };
      if (bundle.core?.settings) {
        await storage.set(NAV_KEYS.coreSettings, bundle.core.settings.coreSettings ?? {});
        report.core.applied = true;
      }
      if (bundle.core?.navigation) {
        report.navigation = await instance.importNavigationState(bundle.core.navigation);
      }
      for (const [id, payload] of Object.entries(bundle.modules ?? {})) {
        const provider = backupProviders.get(id);
        if (!provider?.import) {
          report.modules[id] = { applied: false, warnings: ['provider not found'] };
          continue;
        }
        report.modules[id] = await provider.import(payload.data, { mode: opts.mode ?? 'merge', includeUserData: opts.includeUserData !== false });
      }
      return report;
    },
    async exportDelta({ baseId, baseHashB64, sinceRevision = 0 } = {}) {
      const log = (await storage.get(NAV_KEYS.revisionLog)) ?? [];
      const changes = log.filter((item) => item.rev > sinceRevision);
      return {
        format: DELTA_BACKUP_FORMAT,
        formatVersion: 1,
        base: { baseId, baseHashB64 },
        createdAt: new Date().toISOString(),
        revision: state.revision,
        changes: { core: { set: { revision: state.revision }, del: [] }, navigation: changes, modules: {} }
      };
    },
    async applyDelta(baseBundle, deltaBundle) {
      if (baseBundle.backupId !== deltaBundle.base.baseId) throw new Error('Delta baseId mismatch');
      return { ...baseBundle, deltaAppliedAt: new Date().toISOString(), delta: deltaBundle };
    },

    // Export datasets to an XLSX workbook. Accepts optional journalIds array and filename. If journalIds is omitted, all journals will be exported.
    async exportXlsx({ journalIds, filename } = {}) {
      const tableStore = createTableStoreModule();
      const bundle = await tableStore.exportTableData(storage, { journalIds, includeFormatting: false });
      const sheets = [];
      const journalNameById = {};
      for (const j of state.journals) {
        if (j && j.id) journalNameById[j.id] = j.name || j.title || j.id;
      }
      for (const dataset of bundle.datasets) {
        let columns = [];
        if (dataset.records && dataset.records.length > 0) {
          const first = dataset.records[0];
          columns = Object.keys(first.cells ?? {});
          for (const rec of dataset.records) {
            for (const k of Object.keys(rec.cells ?? {})) {
              if (!columns.includes(k)) columns.push(k);
            }
          }
        }
        if (dataset.schemaId) {
          const schema = schemaRegistry.get(dataset.schemaId);
          if (schema && Array.isArray(schema.columns?.order)) {
            const ordered = [];
            for (const k of schema.columns.order) if (columns.includes(k)) ordered.push(k);
            for (const k of columns) if (!ordered.includes(k)) ordered.push(k);
            columns = ordered;
          }
        }
        const rows = [];
        for (const rec of dataset.records) {
          const rowObj = {};
          for (const col of columns) {
            rowObj[col] = rec.cells?.[col] ?? '';
          }
          rows.push(rowObj);
        }
        const sheetName = journalNameById[dataset.journalId] ?? String(dataset.journalId);
        sheets.push({ name: sheetName, columns, rows });
      }
      const bytes = excelBuildWorkbook(sheets);
      const fname = (filename || 'export') + '_' + new Date().toISOString().replace(/[:\.]/g, '-') + '.xlsx';
      const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fname;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      return { exported: true, sheets: sheets.length };
    },

    // Import records from an XLSX file. Each sheet will be imported into a journal matching either the sheet name or a journal with that name.
    async importXlsx(file, { mode = 'merge' } = {}) {
      if (!file) throw new Error('importXlsx: file is required');
      const ab = await (file.arrayBuffer ? file.arrayBuffer() : new Response(file).arrayBuffer());
      const entries = await excelUnzipEntries(ab);
      const sheets = await excelParseWorkbook(entries);
      const tableStore = createTableStoreModule();
      const journalIdByName = {};
      for (const j of state.journals) {
        const nameKey = (j.name || j.title || '').trim();
        if (nameKey) journalIdByName[nameKey] = j.id;
      }
      const results = [];
      for (const sheet of sheets) {
        const jId = journalIdByName[sheet.name] ?? sheet.name;
        const records = sheet.rows.map((row) => {
          const cells = {};
          for (const [key, value] of Object.entries(row)) {
            const vStr = String(value ?? '').trim();
            if (/^-?\d+(?:\.\d+)?$/.test(vStr)) {
              cells[key] = Number(vStr);
            } else {
              cells[key] = value;
            }
          }
          return {
            id: crypto.randomUUID(),
            cells,
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
          };
        });
        await tableStore.upsertRecords(storage, jId, records, mode);
        results.push({ journalId: jId, imported: records.length });
      }
      return { imported: true, sheets: results };
    },
    on(event, handler) {
      const arr = listeners.get(event) ?? [];
      arr.push(handler);
      listeners.set(event, arr);
      return () => instance.off(event, handler);
    },
    off(event, handler) {
      const arr = listeners.get(event) ?? [];
      listeners.set(event, arr.filter((h) => h !== handler));
    }
  };

  if (Array.isArray(options.modules)) {
    for (const module of options.modules) instance.use(module);
  }
  return instance;
}

export function createNavi(storage) {
  assertStorage(storage);
  return {
    async exportNavigationState() {
      return toNavPayload(await loadNavigationState(storage));
    },
    async importNavigationState(payload) {
      return saveNavigationState(storage, fromNavPayload(payload));
    }
  };
}

export { encryptBackup, decryptBackup, signBackup, verifyBackup, verifyIntegrity } from './backup/crypto.js';
export { VERSION as version };

export { createTableEngine, createTableEngineModule } from './modules/table_engine.js';

export { createTableStoreModule } from './modules/table_store.js';
export { createTableFormatterModule, formatCell, parseInput } from './modules/table_formatter.js';
export { createTableRendererModule, getRenderableCells } from './modules/table_renderer.js';
export { createJournalStore } from './stores/journal_store.js';
export { createJournalTemplatesContainer } from './stores/journal_templates_container.js';
export { createTableSubrowsBridge } from './modules/table_subrows_bridge.js';
