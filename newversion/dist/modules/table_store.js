const KEYS = {
  index: 'tableStore:index',
  rev: 'tableStore:rev',
  dataset: (journalId) => `tableStore:dataset:${journalId}`,
  chlog: (journalId) => `tableStore:chlog:${journalId}`
};

function nowIso() {
  return new Date().toISOString();
}

function clone(value) {
  return structuredClone(value);
}

function ensureArray(value) {
  return Array.isArray(value) ? value : [];
}

function normalizeRecord(record = {}) {
  return {
    id: record.id ?? crypto.randomUUID(),
    cells: { ...(record.cells ?? {}) },
    // Cell-level subrows: array of { cells: {..} }
    subrows: Array.isArray(record.subrows)
      ? record.subrows.map((s) => ({ cells: { ...((s && s.cells) ? s.cells : {}) } }))
      : undefined,
    fmt: record.fmt ? clone(record.fmt) : undefined,
    rowFmt: record.rowFmt ? clone(record.rowFmt) : undefined,
    tags: Array.isArray(record.tags) ? [...record.tags] : undefined,
    createdAt: record.createdAt ?? nowIso(),
    updatedAt: record.updatedAt ?? nowIso()
  };
}

function normalizeDataset(journalId, input = {}) {
  return {
    journalId,
    schemaId: input.schemaId,
    records: ensureArray(input.records).map((record) => normalizeRecord(record)),
    meta: {
      createdAt: input.meta?.createdAt ?? nowIso(),
      updatedAt: input.meta?.updatedAt ?? nowIso(),
      revision: Number(input.meta?.revision ?? 0)
    }
  };
}

function stripFormatting(dataset) {
  return {
    ...dataset,
    records: dataset.records.map((record) => ({
      ...record,
      fmt: undefined,
      rowFmt: undefined
    }))
  };
}

function validateBundle(bundle) {
  const errors = [];
  if (!bundle || bundle.format !== 'sdo-table-data' || bundle.formatVersion !== 1) {
    errors.push('Unsupported bundle format/version');
    return { valid: false, errors };
  }
  if (!Array.isArray(bundle.datasets)) errors.push('datasets must be an array');
  else {
    for (const dataset of bundle.datasets) {
      if (!dataset?.journalId) errors.push('dataset.journalId is required');
      if (!Array.isArray(dataset?.records)) errors.push(`dataset.records must be array for ${dataset?.journalId ?? '<unknown>'}`);
    }
  }
  return { valid: errors.length === 0, errors };
}

export function createTableStoreModule() {
  async function getIndex(storage) {
    return ensureArray(await storage.get(KEYS.index));
  }

  async function setIndex(storage, index) {
    await storage.set(KEYS.index, index);
  }

  async function bumpGlobalRev(storage) {
    const next = Number((await storage.get(KEYS.rev)) ?? 0) + 1;
    await storage.set(KEYS.rev, next);
    return next;
  }

  async function appendChange(storage, journalId, entry) {
    const key = KEYS.chlog(journalId);
    const current = ensureArray(await storage.get(key));
    current.push(entry);
    await storage.set(key, current.slice(-200));
  }

  async function getDataset(storage, journalId) {
    const current = await storage.get(KEYS.dataset(journalId));
    if (!current) return normalizeDataset(journalId, {});
    return normalizeDataset(journalId, current);
  }

  async function saveDataset(storage, dataset, { trackDelete = false } = {}) {
    const normalized = normalizeDataset(dataset.journalId, dataset);
    normalized.meta.updatedAt = nowIso();
    normalized.meta.revision += 1;

    await storage.set(KEYS.dataset(dataset.journalId), normalized);
    const index = await getIndex(storage);
    const nextIndex = index.filter((item) => item.journalId !== dataset.journalId);
    nextIndex.push({ journalId: dataset.journalId, revision: normalized.meta.revision, updatedAt: normalized.meta.updatedAt });
    await setIndex(storage, nextIndex);

    const rev = await bumpGlobalRev(storage);
    await appendChange(storage, dataset.journalId, {
      baseRev: rev - 1,
      toRev: rev,
      set: { [KEYS.dataset(dataset.journalId)]: normalized },
      del: trackDelete ? [KEYS.dataset(dataset.journalId)] : []
    });

    return normalized;
  }

  async function clearDataset(storage, journalId) {
    await storage.del(KEYS.dataset(journalId));
    const index = await getIndex(storage);
    await setIndex(storage, index.filter((item) => item.journalId !== journalId));
    const rev = await bumpGlobalRev(storage);
    await appendChange(storage, journalId, {
      baseRev: rev - 1,
      toRev: rev,
      set: {},
      del: [KEYS.dataset(journalId)]
    });
  }

  async function upsertRecords(storage, journalId, records, mode = 'merge') {
    const current = await getDataset(storage, journalId);
    const incoming = ensureArray(records).map((record) => normalizeRecord(record));

    let nextRecords;
    if (mode === 'replace') {
      nextRecords = incoming;
    } else {
      const map = new Map(current.records.map((record) => [record.id, record]));
      for (const record of incoming) {
        map.set(record.id, {
          ...map.get(record.id),
          ...record,
          cells: { ...(map.get(record.id)?.cells ?? {}), ...(record.cells ?? {}) },
          // If incoming has subrows defined, take it; otherwise keep existing.
          subrows: record.subrows ?? map.get(record.id)?.subrows,
          fmt: record.fmt ?? map.get(record.id)?.fmt,
          rowFmt: record.rowFmt ?? map.get(record.id)?.rowFmt,
          updatedAt: nowIso(),
          createdAt: map.get(record.id)?.createdAt ?? record.createdAt ?? nowIso()
        });
      }
      nextRecords = [...map.values()];
    }

    return saveDataset(storage, { ...current, records: nextRecords });
  }

  async function exportTableData(storage, { journalIds, includeFormatting = true } = {}) {
    const index = await getIndex(storage);
    const ids = journalIds?.length ? journalIds : index.map((item) => item.journalId);
    const datasets = [];
    for (const id of ids) {
      const dataset = await getDataset(storage, id);
      datasets.push(includeFormatting ? dataset : stripFormatting(dataset));
    }
    return {
      format: 'sdo-table-data',
      formatVersion: 1,
      exportedAt: nowIso(),
      datasets
    };
  }

  async function importTableData(storage, bundle, { mode = 'merge' } = {}) {
    const validation = validateBundle(bundle);
    if (!validation.valid) return { applied: false, errors: validation.errors, datasets: [] };

    const results = [];
    for (const incoming of bundle.datasets) {
      const journalId = incoming.journalId;
      const normalizedIncoming = normalizeDataset(journalId, incoming);
      if (mode === 'replace') {
        const saved = await saveDataset(storage, normalizedIncoming);
        results.push({ journalId, revision: saved.meta.revision, mode: 'replace' });
      } else {
        const current = await getDataset(storage, journalId);
        const map = new Map(current.records.map((record) => [record.id, record]));
        for (const record of normalizedIncoming.records) map.set(record.id, record);
        const saved = await saveDataset(storage, { ...current, schemaId: normalizedIncoming.schemaId ?? current.schemaId, records: [...map.values()] });
        results.push({ journalId, revision: saved.meta.revision, mode: 'merge' });
      }
    }

    return { applied: true, errors: [], datasets: results };
  }

  async function exportDelta(storage, { sinceRev = 0 } = {}) {
    const index = await getIndex(storage);
    const set = {};

    let toRev = Number(await storage.get(KEYS.rev) ?? 0);
    for (const item of index) {
      const log = ensureArray(await storage.get(KEYS.chlog(item.journalId)));
      if (log.some((entry) => Number(entry.toRev) > sinceRev)) {
        set[KEYS.dataset(item.journalId)] = await getDataset(storage, item.journalId);
      }
    }

    return {
      baseRev: sinceRev,
      toRev,
      set,
      del: []
    };
  }

  async function applyDelta(storage, delta, { mode = 'merge' } = {}) {
    if (!delta || typeof delta !== 'object') return { applied: false, errors: ['Invalid delta'] };
    const setEntries = Object.entries(delta.set ?? {});

    for (const [key, dataset] of setEntries) {
      if (!key.startsWith('tableStore:dataset:')) continue;
      const journalId = key.slice('tableStore:dataset:'.length);
      if (mode === 'replace') {
        await saveDataset(storage, normalizeDataset(journalId, dataset));
      } else {
        await importTableData(storage, {
          format: 'sdo-table-data',
          formatVersion: 1,
          exportedAt: nowIso(),
          datasets: [{ ...dataset, journalId }]
        }, { mode: 'merge' });
      }
    }

    for (const key of delta.del ?? []) {
      if (key.startsWith('tableStore:dataset:')) {
        const journalId = key.slice('tableStore:dataset:'.length);
        await clearDataset(storage, journalId);
      }
    }

    return { applied: true, errors: [] };
  }

  return {
    id: '@sdo/module-table-store',
    version: '1.0.0',
    init(ctx) {
      const api = {
        getDataset: (journalId) => getDataset(ctx.storage, journalId),
        listDatasets: () => getIndex(ctx.storage),
        addRecord: async (journalId, recordPartial) => {
          const current = await getDataset(ctx.storage, journalId);
          const record = normalizeRecord(recordPartial ?? {});
          await saveDataset(ctx.storage, { ...current, records: [...current.records, record] });
          return record.id;
        },
        updateRecord: async (journalId, recordId, patch) => {
          const current = await getDataset(ctx.storage, journalId);
          const next = current.records.map((record) => {
            if (record.id !== recordId) return record;
            return {
              ...record,
              ...patch,
              cells: { ...(record.cells ?? {}), ...(patch?.cells ?? {}) },
              fmt: patch?.fmt ? { ...(record.fmt ?? {}), ...patch.fmt } : record.fmt,
              rowFmt: patch?.rowFmt ? { ...(record.rowFmt ?? {}), ...patch.rowFmt } : record.rowFmt,
              updatedAt: nowIso()
            };
          });
          await saveDataset(ctx.storage, { ...current, records: next });
        },
        deleteRecord: async (journalId, recordId) => {
          const current = await getDataset(ctx.storage, journalId);
          await saveDataset(ctx.storage, { ...current, records: current.records.filter((record) => record.id !== recordId) });
        },
        clearDataset: (journalId) => clearDataset(ctx.storage, journalId),
        upsertRecords: async (journalId, records, mode = 'merge') => {
          await upsertRecords(ctx.storage, journalId, records, mode);
        },
        deleteRecords: async (journalId, ids = []) => {
          const remove = new Set(ids);
          const current = await getDataset(ctx.storage, journalId);
          await saveDataset(ctx.storage, { ...current, records: current.records.filter((record) => !remove.has(record.id)) });
        },
        exportTableData: (opts) => exportTableData(ctx.storage, opts),
        importTableData: (bundle, opts) => importTableData(ctx.storage, bundle, opts),
        exportDelta: (opts) => exportDelta(ctx.storage, opts),
        applyDelta: (delta, opts) => applyDelta(ctx.storage, delta, opts)
      };

      ctx.api.tableStore = api;

      ctx.registerCommands([
        {
          id: '@sdo/module-table-store.export',
          title: 'Export table data',
          run: async () => api.exportTableData({ includeFormatting: true })
        }
      ]);

      ctx.backup.registerProvider({
        id: 'tableStore',
        version: '1.0.0',
        describe: async () => ({ settings: [KEYS.index, KEYS.rev], userData: (await getIndex(ctx.storage)).map((item) => KEYS.dataset(item.journalId)) }),
        export: async (opts = {}) => {
          const payload = {
            revision: Number(await ctx.storage.get(KEYS.rev) ?? 0),
            index: await getIndex(ctx.storage)
          };
          if (opts.includeUserData !== false) {
            payload.userData = await api.exportTableData({ includeFormatting: true });
          }
          return payload;
        },
        import: async (payload, opts = {}) => {
          if (opts.includeUserData !== false && payload.userData) {
            return api.importTableData(payload.userData, { mode: opts.mode ?? 'merge' });
          }
          return { applied: true, errors: [] };
        },
        exportDelta: async (sinceRev = 0) => api.exportDelta({ sinceRev }),
        applyDelta: async (patch, opts = {}) => api.applyDelta(patch, { mode: opts.mode ?? 'merge' })
      });
    }
  };
}

export { KEYS as tableStoreKeys };
