function normalizeSchema(schema = {}) {
  const fields = Array.isArray(schema.fields) ? schema.fields : [];
  return {
    ...schema,
    fields,
    fieldMap: new Map(fields.map((field) => [field.key, field]))
  };
}

function normalizeSettings(settings = {}) {
  return {
    columns: {
      order: Array.isArray(settings.columns?.order) ? settings.columns.order : null,
      visibility: settings.columns?.visibility ?? {},
      widths: settings.columns?.widths ?? {}
    },
    sort: settings.sort ?? null,
    filter: {
      global: settings.filter?.global ?? '',
      perColumn: settings.filter?.perColumn ?? {}
    },
    expandedRowIds: new Set(settings.expandedRowIds ?? []),
    selectedRowIds: new Set(settings.selectedRowIds ?? []),
    subrows: {
      columnsSubrowsEnabled: settings?.subrows?.columnsSubrowsEnabled ?? {}
    }
  };
}

function normalizeDataset(dataset = {}) {
  return {
    records: Array.isArray(dataset.records) ? dataset.records : [],
    merges: Array.isArray(dataset.merges) ? dataset.merges : []
  };
}

function toDisplayText(value) {
  if (value == null) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  return JSON.stringify(value);
}

function compareValues(a, b) {
  if (a == null && b == null) return 0;
  if (a == null) return 1;
  if (b == null) return -1;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
}

function defaultByType(type) {
  if (type === 'number' || type === 'money') return 0;
  if (type === 'boolean') return false;
  return '';
}

function buildColumns(schema, settings) {
  const keys = schema.fields.map((field) => field.key);
  const ordered = settings.columns.order ? settings.columns.order.filter((key) => keys.includes(key)) : keys;
  const rest = keys.filter((key) => !ordered.includes(key));
  return [...ordered, ...rest]
    .filter((key) => settings.columns.visibility[key] !== false)
    .map((key) => ({
      columnKey: key,
      width: settings.columns.widths[key] ?? null,
      field: schema.fieldMap.get(key)
    }));
}

function buildDatasetIndex(dataset) {
  const recordById = new Map();
  const childrenById = new Map();
  const parentById = new Map();

  for (const record of dataset.records) {
    recordById.set(record.id, record);
    childrenById.set(record.id, []);
  }

  for (const record of dataset.records) {
    const children = childrenById.get(record.id) ?? [];

    if (Array.isArray(record.childrenIds)) {
      for (const childId of record.childrenIds) {
        if (!recordById.has(childId)) continue;
        children.push(childId);
        parentById.set(childId, record.id);
      }
    }

    if (record.parentId && recordById.has(record.parentId)) {
      const siblings = childrenById.get(record.parentId) ?? [];
      if (!siblings.includes(record.id)) siblings.push(record.id);
      parentById.set(record.id, record.parentId);
    }

    childrenById.set(record.id, children);
  }

  const rootIds = dataset.records.filter((record) => !parentById.has(record.id)).map((record) => record.id);
  return { recordById, childrenById, parentById, rootIds };
}

function collectVisibleIds({ index, columns, settings }) {
  const globalNeedle = settings.filter.global.trim().toLowerCase();
  const visibleSet = new Set();

  const sortIds = createSorter(index, settings.sort);

  const rowMatches = (record) => {
    if (!globalNeedle) return true;
    for (const column of columns) {
      const text = toDisplayText(record.cells?.[column.columnKey]).toLowerCase();
      if (text.includes(globalNeedle)) return true;
    }
    return false;
  };

  const dfs = (rowId) => {
    const record = index.recordById.get(rowId);
    if (!record) return false;
    const childIds = sortIds(index.childrenById.get(rowId) ?? []);
    let childVisible = false;
    for (const childId of childIds) {
      if (dfs(childId)) childVisible = true;
    }
    const keep = rowMatches(record) || childVisible;
    if (keep) visibleSet.add(rowId);
    return keep;
  };

  for (const rootId of sortIds(index.rootIds)) dfs(rootId);
  return { visibleSet, sortIds };
}

function createSorter(index, sortState) {
  return (ids) => {
    if (!sortState?.columnKey || !sortState?.dir) return [...ids];
    const dir = sortState.dir === 'desc' ? -1 : 1;
    return [...ids].sort((a, b) => {
      const av = index.recordById.get(a)?.cells?.[sortState.columnKey];
      const bv = index.recordById.get(b)?.cells?.[sortState.columnKey];
      const cmp = compareValues(av, bv);
      if (cmp !== 0) return cmp * dir;
      return String(a).localeCompare(String(b));
    });
  };
}

function flattenRows({ index, settings, editState, visibleSet, sortIds }) {
  const rows = [];

  const walk = (rowId, depth) => {
    if (!visibleSet.has(rowId)) return;
    const record = index.recordById.get(rowId);
    if (!record) return;

    // UI model: Subrows live under a technical (kind:'group') node and are rendered INSIDE cells,
    // not as separate visible rows. Therefore we skip group nodes in the visible tree traversal.
    const allChildIds = sortIds(index.childrenById.get(rowId) ?? []).filter((childId) => visibleSet.has(childId));
    const displayChildIds = allChildIds.filter((childId) => {
      const childRec = index.recordById.get(childId);
      return childRec?.kind !== 'group';
    });

    const hasChildren = displayChildIds.length > 0;
    const isExpanded = hasChildren && settings.expandedRowIds.has(rowId);

    rows.push({
      rowId,
      record,
      depth,
      isParent: hasChildren,
      hasChildren,
      isExpanded,
      isSelected: settings.selectedRowIds.has(rowId),
      isEditing: editState?.rowId === rowId ? editState : null
    });

    if (!isExpanded) return;
    for (const childId of displayChildIds) {
      walk(childId, depth + 1);
    }
  };

  for (const rootId of sortIds(index.rootIds)) walk(rootId, 0);
  return rows;
}

function buildCellSpanMap(merges, rowIds, columnKeys) {
  const rowIndex = new Map(rowIds.map((id, idx) => [id, idx]));
  const colIndex = new Map(columnKeys.map((key, idx) => [key, idx]));
  const spans = new Map();

  for (const merge of merges) {
    const startRow = rowIndex.get(merge.rowId);
    const startCol = colIndex.get(merge.colKey);
    if (startRow == null || startCol == null) continue;

    const rowSpan = Math.max(1, merge.rowSpan ?? 1);
    const colSpan = Math.max(1, merge.colSpan ?? 1);
    const topLeft = `${merge.rowId}:${merge.colKey}`;
    spans.set(topLeft, { rowSpan, colSpan });

    for (let r = startRow; r < startRow + rowSpan && r < rowIds.length; r += 1) {
      for (let c = startCol; c < startCol + colSpan && c < columnKeys.length; c += 1) {
        const rowId = rowIds[r];
        const colKey = columnKeys[c];
        const key = `${rowId}:${colKey}`;
        if (key === topLeft) continue;
        spans.set(key, { coveredBy: { rowId: merge.rowId, colKey: merge.colKey } });
      }
    }
  }

  return spans;
}

export function createTableEngine({ schema, settings = {} }) {
  let normalizedSchema = normalizeSchema(schema);
  let normalizedSettings = normalizeSettings(settings);
  let dataset = normalizeDataset();
  let datasetIndex = buildDatasetIndex(dataset);
  let editState = null;

  const api = {
    setDataset(nextDataset) {
      dataset = normalizeDataset(nextDataset);
      datasetIndex = buildDatasetIndex(dataset);
      return api;
    },
    setSettings(nextSettings) {
      normalizedSettings = normalizeSettings(nextSettings);
      return api;
    },
    setSort(columnKey, dir = 'asc') {
      normalizedSettings.sort = { columnKey, dir };
      return api;
    },
    setGlobalFilter(text = '') {
      normalizedSettings.filter.global = text;
      return api;
    },
    toggleExpand(rowId) {
      if (normalizedSettings.expandedRowIds.has(rowId)) normalizedSettings.expandedRowIds.delete(rowId);
      else normalizedSettings.expandedRowIds.add(rowId);
      return api;
    },
    toggleSelect(rowId) {
      if (normalizedSettings.selectedRowIds.has(rowId)) normalizedSettings.selectedRowIds.delete(rowId);
      else normalizedSettings.selectedRowIds.add(rowId);
      return api;
    },
    selectAllVisible() {
      const view = api.compute();
      for (const row of view.rows) normalizedSettings.selectedRowIds.add(row.rowId);
      return api;
    },
    beginEdit(rowId, colKey) {
      editState = { rowId, colKey };
      return editState;
    },
    applyEdit(rowId, colKey, newValue) {
      editState = null;
      const record = datasetIndex.recordById.get(rowId);
      return {
        recordId: rowId,
        cellsPatch: { [colKey]: newValue },
        fmtPatch: record?.fmt?.[colKey] ? { [colKey]: record.fmt[colKey] } : undefined
      };
    },
    cancelEdit() {
      editState = null;
      return api;
    },
    getAddFormModel() {
      return normalizedSchema.fields.map((field) => ({
        key: field.key,
        label: field.label,
        type: field.type,
        required: !!field.required,
        default: field.default ?? defaultByType(field.type)
      }));
    },
    validateAddForm(values = {}) {
      const errors = [];
      for (const field of normalizedSchema.fields) {
        const value = values[field.key];
        if (field.required && (value == null || value === '')) {
          errors.push({ key: field.key, code: 'required', message: `${field.key} is required` });
          continue;
        }
        if (value == null || value === '') continue;
        if ((field.type === 'number' || field.type === 'money') && Number.isNaN(Number(value))) {
          errors.push({ key: field.key, code: 'type', message: `${field.key} must be a number` });
        }
        if (field.type === 'boolean' && typeof value !== 'boolean') {
          errors.push({ key: field.key, code: 'type', message: `${field.key} must be a boolean` });
        }
      }
      return { valid: errors.length === 0, errors };
    },
    buildRecordFromForm(values = {}) {
      const cells = {};
      for (const field of normalizedSchema.fields) {
        cells[field.key] = values[field.key] ?? field.default ?? defaultByType(field.type);
      }
      return {
        id: crypto.randomUUID(),
        cells,
        fmt: {},
        childrenIds: []
      };
    },
    addRow(values = {}, inputDataset = dataset) {
      const validation = api.validateAddForm(values);
      if (!validation.valid) {
        return { ok: false, errors: validation.errors, dataset: inputDataset };
      }
      const record = api.buildRecordFromForm(values);
      const nextDataset = {
        ...inputDataset,
        records: [...(inputDataset.records ?? []), record]
      };
      return { ok: true, record, dataset: nextDataset };
    },

    // ------------------------------
// Engine-native subrows API (cell-level)
// ------------------------------

_ensureSubrowsArray(rowId, inputDataset = dataset) {
  const ds = normalizeDataset(inputDataset);
  const idx = buildDatasetIndex(ds);
  const rec = idx.recordById.get(rowId);
  if (!rec) return { dataset: ds, ok: false };
  if (Array.isArray(rec.subrows)) return { dataset: ds, ok: true };
  const nextRecords = (ds.records ?? []).map((r) => (r.id === rowId ? { ...r, subrows: [] } : r));
  return { dataset: { ...ds, records: nextRecords }, ok: true };
},

// Kept for API compatibility (groupId is not used in cell-level model)
ensureSubrowsGroup(rowId, inputDataset = dataset) {
  const { dataset: ds } = api._ensureSubrowsArray(rowId, inputDataset);
  return { dataset: ds, groupId: rowId };
},

listSubrows(rowId, inputDataset = dataset) {
  const ds = normalizeDataset(inputDataset);
  const idx = buildDatasetIndex(ds);
  const rec = idx.recordById.get(rowId);
  const arr = Array.isArray(rec?.subrows) ? rec.subrows : [];
  return arr.map((_, i) => `${rowId}::sub::${i}`);
},

_parseSubrowId(subrowId) {
  const m = typeof subrowId === 'string' ? subrowId.match(/^(.*)::sub::(\d+)$/) : null;
  if (!m) return null;
  return { ownerId: m[1], index: Number(m[2]) };
},

addSubrow(rowId, { insertAfterId, cells } = {}, inputDataset = dataset) {
  let { dataset: ds } = api._ensureSubrowsArray(rowId, inputDataset);
  const idx = buildDatasetIndex(ds);
  const rec = idx.recordById.get(rowId);
  if (!rec) return { dataset: ds, subrowId: null };
  const subrows = Array.isArray(rec.subrows) ? [...rec.subrows] : [];
  const newSub = { cells: { ...(cells ?? {}) } };
  let at = subrows.length;
  if (insertAfterId) {
    const p = api._parseSubrowId(insertAfterId);
    if (p && p.ownerId === rowId && Number.isFinite(p.index)) at = Math.min(subrows.length, p.index + 1);
  }
  subrows.splice(at, 0, newSub);
  const nextRecords = (ds.records ?? []).map((r) => (r.id === rowId ? { ...r, subrows } : r));
  // Recompute id at same index after insert
  const subrowId = `${rowId}::sub::${at}`;
  return { dataset: { ...ds, records: nextRecords }, subrowId };
},

removeSubrow(subrowId, inputDataset = dataset) {
  const ds = normalizeDataset(inputDataset);
  const p = api._parseSubrowId(subrowId);
  if (!p) return { dataset: ds, removed: false };
  const idx = buildDatasetIndex(ds);
  const rec = idx.recordById.get(p.ownerId);
  if (!rec || !Array.isArray(rec.subrows) || p.index < 0 || p.index >= rec.subrows.length) return { dataset: ds, removed: false };
  const subrows = [...rec.subrows];
  subrows.splice(p.index, 1);
  const nextRecords = (ds.records ?? []).map((r) => (r.id === p.ownerId ? { ...r, subrows } : r));
  return { dataset: { ...ds, records: nextRecords }, removed: true };
},

applySubrowEdit(subrowId, colKey, value, inputDataset = dataset) {
  const ds = normalizeDataset(inputDataset);
  const p = api._parseSubrowId(subrowId);
  if (!p) return { dataset: ds, ok: false };
  const idx = buildDatasetIndex(ds);
  const rec = idx.recordById.get(p.ownerId);
  if (!rec || !Array.isArray(rec.subrows) || p.index < 0 || p.index >= rec.subrows.length) return { dataset: ds, ok: false };
  const subrows = rec.subrows.map((s, i) => {
    if (i !== p.index) return s;
    const cells2 = { ...(s?.cells ?? {}) };
    cells2[colKey] = value;
    return { ...s, cells: cells2 };
  });
  const nextRecords = (ds.records ?? []).map((r) => (r.id === p.ownerId ? { ...r, subrows } : r));
  return { dataset: { ...ds, records: nextRecords }, ok: true };
},

resolveCellEditTarget({ rowId, colKey, settings }) {
  const enabled = settings?.subrows?.columnsSubrowsEnabled?.[colKey] !== false;
  if (!enabled) return { editTargetRowId: rowId, needsChoice: false, candidates: [rowId] };
  const subs = api.listSubrows(rowId);
  return { editTargetRowId: rowId, needsChoice: false, candidates: [rowId, ...subs] };
},

getTransferCandidates(rowId) {
  return [rowId];
},

getSubrowLabel(subrowId) {
  const p = api._parseSubrowId(subrowId);
  if (!p) return '';
  return `Підстрока №${p.index + 1}`;
},

    compute() {
      const columns = buildColumns(normalizedSchema, normalizedSettings);
      const { visibleSet, sortIds } = collectVisibleIds({
        index: datasetIndex,
        columns,
        settings: normalizedSettings
      });
      const rows = flattenRows({
        index: datasetIndex,
        settings: normalizedSettings,
        editState,
        visibleSet,
        sortIds
      });

      const rowIds = rows.map((row) => row.rowId);
      const columnKeys = columns.map((column) => column.columnKey);
      const cellSpanMap = buildCellSpanMap(dataset.merges, rowIds, columnKeys);

      return {
        columns,
        rows,
        mergedCells: [...cellSpanMap.entries()].map(([cell, span]) => ({ cell, ...span })),
        cellSpanMap,
        selection: [...normalizedSettings.selectedRowIds],
        sortState: normalizedSettings.sort,
        filterState: {
          global: normalizedSettings.filter.global,
          perColumn: normalizedSettings.filter.perColumn
        }
      };
    }
  };

  return api;
}

export function createTableEngineModule({ schema, initialSettings = {}, dataset = {} }) {
  const engine = createTableEngine({ schema, settings: initialSettings });
  engine.setDataset(dataset);

  return {
    id: '@sdo/module-table-engine',
    version: '1.0.0',
    init(ctx) {
      ctx.registerSchema({
        id: '@sdo/module-table-engine.records',
        version: '1.0.0',
        domain: 'table',
        appliesTo: { any: true },
        fields: schema.fields
      });

      ctx.registerSettings({
        id: '@sdo/module-table-engine.settings',
        tab: { id: 'table-engine', title: 'Table Engine', order: 20 },
        fields: [
          {
            key: '@sdo/module-table-engine:filter.global',
            label: 'Global filter',
            type: 'text',
            default: '',
            read: async () => initialSettings.filter?.global ?? '',
            write: async (_runtime, value) => engine.setGlobalFilter(value)
          }
        ]
      });

      ctx.registerCommands([
        {
          id: '@sdo/module-table-engine.compute',
          title: 'Compute table view',
          run: async () => engine.compute()
        }
      ]);

      ctx.ui.registerButton({
        id: '@sdo/module-table-engine:compute',
        label: 'Compute table',
        location: 'toolbar',
        order: 40,
        onClick: () => ctx.commands.run('@sdo/module-table-engine.compute')
      });
    },
    engine
  };
}
