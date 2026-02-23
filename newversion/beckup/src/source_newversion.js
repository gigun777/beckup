import { NAV_KEYS, loadNavigationState } from '../../src/storage/db_nav.js';

function toColumnsFromPayload(payload) {
  if (!Array.isArray(payload?.sheet?.columns)) return [];
  return payload.sheet.columns.map((c) => c?.name || c?.key).filter(Boolean);
}

function recordsFromRowsV2(rowsV2, columns) {
  return (rowsV2 || []).map((r) => {
    const cells = {};
    for (let i = 0; i < columns.length; i += 1) cells[columns[i]] = r.cells?.[i] ?? '';
    return {
      id: r.id || crypto.randomUUID(),
      cells,
      subrows: Array.isArray(r.subrows) ? r.subrows : [],
      createdAt: r.createdAt || null,
      updatedAt: r.updatedAt || null
    };
  });
}

function recordsFromLegacyRows(rows) {
  return (rows || []).map((r) => {
    const src = r?.exportData || r?.data || {};
    return {
      id: r?.id || crypto.randomUUID(),
      cells: { ...src },
      subrows: Array.isArray(r?.subrows) ? r.subrows : [],
      createdAt: r?.createdAt || null,
      updatedAt: r?.updatedAt || null
    };
  });
}

/**
 * Create source/target adapter for newversion storage.
 * This ensures backup/export reads from the primary source (storage), not UI-rendered tables.
 */
export function createNewversionSourceAdapter(storage, { tableDatasetPrefix = 'tableStore:dataset:' } = {}) {
  if (!storage || typeof storage.get !== 'function' || typeof storage.set !== 'function') {
    throw new Error('storage with get/set is required');
  }

  async function listJournals() {
    const nav = await loadNavigationState(storage);
    return Array.isArray(nav.journals) ? nav.journals : [];
  }

  async function resolveJournal(journalKeyOrId) {
    const journals = await listJournals();
    return journals.find((j) => (j.key === journalKeyOrId || j.id === journalKeyOrId)) || null;
  }

  function datasetKey(journalId) {
    return `${tableDatasetPrefix}${journalId}`;
  }

  async function saveRecordsToJournal(journalId, incomingRecords, { mode = 'merge' } = {}) {
    const key = datasetKey(journalId);
    const current = await storage.get(key);
    const currentRecords = Array.isArray(current?.records) ? current.records : [];

    let records;
    if (mode === 'replace') {
      records = incomingRecords;
    } else {
      const byId = new Map(currentRecords.map((r) => [r.id, r]));
      for (const r of incomingRecords) byId.set(r.id, r);
      records = [...byId.values()];
    }

    await storage.set(key, {
      ...(current || {}),
      journalId,
      schema: current?.schema || null,
      records,
      merges: Array.isArray(current?.merges) ? current.merges : []
    });
  }

  return {
    async listJournals() {
      return listJournals();
    },

    async loadJournalSchema(journalId) {
      const dataset = await storage.get(datasetKey(journalId));
  return {
    async listJournals() {
      const nav = await loadNavigationState(storage);
      return Array.isArray(nav.journals) ? nav.journals : [];
    },

    async loadJournalSchema(journalId) {
      const dataset = await storage.get(`${tableDatasetPrefix}${journalId}`);
      // schemaId may point to registry; for now return columns fallback from dataset meta if present.
      return dataset?.schema || { fields: [] };
    },

    async loadJournalRecords(journalId) {
      const dataset = await storage.get(datasetKey(journalId));
      const dataset = await storage.get(`${tableDatasetPrefix}${journalId}`);
      return Array.isArray(dataset?.records) ? dataset.records : [];
    },

    async loadJournalExportProfile(journalId) {
      return await storage.get(`@sdo/module-table-renderer:settings:${journalId}`) || null;
    },

    async loadSettings() {
      return {
        core: await storage.get(NAV_KEYS.coreSettings),
        tableGlobal: await storage.get('@sdo/module-table-renderer:settings')
      };
    },

    async loadNavigation() {
      return await loadNavigationState(storage);
    },

    async loadTransfer() {
      return {
        templates: await storage.get('transfer:templates:v1')
      };
    },

    async saveJournalPayload(journalKey, payload, { mode = 'merge' } = {}) {
      const journal = await resolveJournal(journalKey);
      const journalId = journal?.id || journalKey;

      const rowsV2 = Array.isArray(payload?.rowsV2) ? payload.rowsV2 : [];
      const columns = toColumnsFromPayload(payload);

      let incomingRecords = recordsFromRowsV2(rowsV2, columns);
      if (!incomingRecords.length && Array.isArray(payload?.rows)) {
        incomingRecords = recordsFromLegacyRows(payload.rows);
      }

      await saveRecordsToJournal(journalId, incomingRecords, { mode });
    },

    async saveJournalRecords(journalKeyOrId, records, { mode = 'merge' } = {}) {
      const journal = await resolveJournal(journalKeyOrId);
      const journalId = journal?.id || journalKeyOrId;
      const incoming = Array.isArray(records) ? records : [];
      await saveRecordsToJournal(journalId, incoming, { mode });
      const nav = await loadNavigationState(storage);
      const journal = (nav.journals || []).find((j) => (j.key === journalKey || j.id === journalKey));
      const journalId = journal?.id || journalKey;
      const key = `${tableDatasetPrefix}${journalId}`;

      const rowsV2 = Array.isArray(payload?.rowsV2) ? payload.rowsV2 : [];
      const columns = Array.isArray(payload?.sheet?.columns) ? payload.sheet.columns.map((c) => c.name || c.key) : [];

      let incomingRecords = rowsV2.map((r) => {
      const incomingRecords = rowsV2.map((r) => {
        const cells = {};
        for (let i = 0; i < columns.length; i += 1) cells[columns[i]] = r.cells?.[i] ?? '';
        return {
          id: r.id || crypto.randomUUID(),
          cells,
          subrows: Array.isArray(r.subrows) ? r.subrows : [],
          createdAt: r.createdAt || null,
          updatedAt: r.updatedAt || null
        };
      });

      // Partial-recovery fallback: if rowsV2 is missing, try legacy rows[].exportData / rows[].data.
      if (!incomingRecords.length && Array.isArray(payload?.rows)) {
        incomingRecords = payload.rows.map((r) => {
          const src = r?.exportData || r?.data || {};
          return {
            id: r?.id || crypto.randomUUID(),
            cells: { ...src },
            subrows: Array.isArray(r?.subrows) ? r.subrows : [],
            createdAt: r?.createdAt || null,
            updatedAt: r?.updatedAt || null
          };
        });
      }

      const current = await storage.get(key);
      const currentRecords = Array.isArray(current?.records) ? current.records : [];
      let records;

      if (mode === 'replace') records = incomingRecords;
      else {
        const byId = new Map(currentRecords.map((r) => [r.id, r]));
        for (const r of incomingRecords) byId.set(r.id, r);
        records = [...byId.values()];
      }

      await storage.set(key, {
        ...(current || {}),
        journalId,
        schema: current?.schema || null,
        records,
        merges: Array.isArray(current?.merges) ? current.merges : []
      });
    },

    async saveSettings(payload, { mode = 'merge' } = {}) {
      if (mode === 'replace') {
        await storage.set(NAV_KEYS.coreSettings, payload?.core || {});
        await storage.set('@sdo/module-table-renderer:settings', payload?.tableGlobal || {});
        return;
      }
      const core = (await storage.get(NAV_KEYS.coreSettings)) || {};
      const table = (await storage.get('@sdo/module-table-renderer:settings')) || {};
      await storage.set(NAV_KEYS.coreSettings, { ...core, ...(payload?.core || {}) });
      await storage.set('@sdo/module-table-renderer:settings', { ...table, ...(payload?.tableGlobal || {}) });
    },

    async saveNavigation(payload) {
      if (!payload) return;
      await storage.set(NAV_KEYS.spaces, payload.spaces || []);
      await storage.set(NAV_KEYS.journals, payload.journals || []);
      await storage.set(NAV_KEYS.lastLoc, payload.lastLoc || null);
      await storage.set(NAV_KEYS.history, payload.history || []);
    },

    async saveTransfer(payload, { mode = 'merge' } = {}) {
      const key = 'transfer:templates:v1';
      if (mode === 'replace') {
        await storage.set(key, payload?.templates || []);
        return;
      }
      const cur = (await storage.get(key)) || [];
      const byId = new Map((Array.isArray(cur) ? cur : []).map((x) => [x.id, x]));
      for (const t of (payload?.templates || [])) {
        if (t?.id) byId.set(t.id, { ...(byId.get(t.id) || {}), ...t });
      }
      await storage.set(key, [...byId.values()]);
    }
  };
}
