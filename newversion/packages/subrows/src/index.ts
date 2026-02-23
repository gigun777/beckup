function cloneDataset(ds = {}) {
  return {
    records: { ...(ds.records ?? {}) },
    order: Array.isArray(ds.order) ? [...ds.order] : []
  };
}

function isEnabled(settings, colKey) {
  return settings?.columnsSubrowsEnabled?.[colKey] === true;
}

function createId(prefix = 'subrow') {
  return `${prefix}_${Math.random().toString(36).slice(2, 8)}`;
}

function ensureGroup(ds, rowId) {
  const next = cloneDataset(ds);
  const row = next.records[rowId];
  if (!row) return { dataset: next, value: null };
  if (row.kind === 'group') return { dataset: next, value: rowId };
  if (row.parentId && next.records[row.parentId]?.kind === 'group') return { dataset: next, value: row.parentId };

  let groupId = `${rowId}__group`;
  while (next.records[groupId]) groupId = createId('group');

  next.records[groupId] = {
    id: groupId,
    kind: 'group',
    parentId: row.parentId ?? null,
    childrenIds: [rowId],
    cells: {}
  };

  next.records[rowId] = { ...row, parentId: groupId };

  const idx = next.order.indexOf(rowId);
  if (idx >= 0) next.order.splice(idx, 0, groupId);
  else next.order.push(groupId);

  const parent = row.parentId ? next.records[row.parentId] : null;
  if (parent?.childrenIds) {
    parent.childrenIds = parent.childrenIds.map((id) => (id === rowId ? groupId : id));
  }

  return { dataset: next, value: groupId };
}

function addSubrow(ds, rowId) {
  const grouped = ensureGroup(ds, rowId);
  const next = cloneDataset(grouped.dataset);
  const groupId = grouped.value;
  if (!groupId) return { dataset: next, value: null };

  const group = next.records[groupId];
  const childrenIds = Array.isArray(group.childrenIds) ? [...group.childrenIds] : [];
  let newId = createId();
  while (next.records[newId]) newId = createId();

  next.records[newId] = { id: newId, kind: 'row', parentId: groupId, childrenIds: [], cells: {} };
  group.childrenIds = [...childrenIds, newId];

  const anchor = childrenIds[childrenIds.length - 1] ?? groupId;
  const anchorIdx = next.order.indexOf(anchor);
  if (anchorIdx >= 0) next.order.splice(anchorIdx + 1, 0, newId);
  else next.order.push(newId);

  return { dataset: next, value: newId };
}

function removeSubrow(ds, subrowId) {
  const next = cloneDataset(ds);
  const row = next.records[subrowId];
  if (!row) return { dataset: next, value: false };
  if (row.kind === 'group') return { dataset: next, value: false };

  const parent = row.parentId ? next.records[row.parentId] : null;
  if (parent?.childrenIds) parent.childrenIds = parent.childrenIds.filter((id) => id !== subrowId);
  delete next.records[subrowId];
  next.order = next.order.filter((id) => id !== subrowId);
  return { dataset: next, value: true };
}

function computeVisibleRows(ds) {
  return (ds?.order ?? []).map((id) => ds.records[id]).filter(Boolean);
}

function getTransferCandidates(ds, rowId) {
  const row = ds?.records?.[rowId];
  if (!row) return [];
  if (row.kind === 'group') return [...(row.childrenIds ?? [])];
  if (row.parentId && ds.records[row.parentId]?.kind === 'group') {
    return [...(ds.records[row.parentId].childrenIds ?? [])].filter((id) => ds.records[id]?.kind !== 'group');
  }
  return [];
}

function resolveEditTarget(ds, cellRef, settings) {
  if (!isEnabled(settings, cellRef.colKey)) return { type: 'normalEdit', targetRowId: cellRef.rowId };
  const candidates = getTransferCandidates(ds, cellRef.rowId);
  if (candidates.length > 0) return { type: 'subrowEdit', targetRowId: candidates[0], candidates };
  return { type: 'normalEdit', targetRowId: cellRef.rowId };
}

async function handleCellClickSubrowsFlow({ ds, cellRef, settings, ui }) {
  if (!isEnabled(settings, cellRef.colKey)) {
    return { dataset: ds, addedSubrowId: null, editTargetRowId: cellRef.rowId, highlightSubrows: [], needsChoice: false };
  }

  const candidates = getTransferCandidates(ds, cellRef.rowId);
  if (candidates.length === 0) {
    const add = addSubrow(ds, cellRef.rowId, settings);
    ui?.toast?.('Створено нову підстроку');
    return { dataset: add.dataset, addedSubrowId: add.value, editTargetRowId: add.value, highlightSubrows: [add.value], needsChoice: false };
  }

  const action = (await ui?.askCellAction?.({ cellRef, candidates })) ?? 'editExisting';
  if (action === 'addSubrow') {
    const add = addSubrow(ds, cellRef.rowId, settings);
    ui?.toast?.('Створено нову підстроку');
    return { dataset: add.dataset, addedSubrowId: add.value, editTargetRowId: add.value, highlightSubrows: [add.value], needsChoice: false };
  }

  if (candidates.length === 1) {
    return { dataset: ds, addedSubrowId: null, editTargetRowId: candidates[0], highlightSubrows: candidates, needsChoice: false };
  }

  const chosen = await ui?.pickSubrow?.({ items: candidates.map((id) => ({ id })) });
  return { dataset: ds, addedSubrowId: null, editTargetRowId: chosen ?? null, highlightSubrows: candidates, needsChoice: true };
}

export { ensureGroup, addSubrow, removeSubrow, computeVisibleRows, resolveEditTarget, handleCellClickSubrowsFlow, getTransferCandidates };
