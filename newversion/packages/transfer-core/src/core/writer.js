function cloneDataset(dataset = {}) {
  return {
    ...dataset,
    records: (dataset.records ?? []).map((record) => ({
      ...record,
      cells: { ...(record.cells ?? {}) }
    })),
    meta: { ...(dataset.meta ?? {}), updatedAt: Date.now() }
  };
}

function appendByMode(currentValue, nextValue, writeMode) {
  const current = currentValue ?? '';
  const incoming = nextValue ?? '';

  if (current === '') return incoming;
  if (writeMode?.appendMode === 'space') return `${current} ${incoming}`;
  if (writeMode?.appendMode === 'newline') return `${current}\n${incoming}`;
  if (writeMode?.appendMode === 'separator') return `${current}${writeMode.appendSeparator ?? ''}${incoming}`;
  return `${current}${incoming}`;
}

export function applyWrite(dataset, targetCell, value, writeMode) {
  const record = dataset.records?.find((item) => item.id === targetCell.recordId);
  if (!record || !targetCell?.fieldId) return false;

  if (writeMode?.mode === 'append') {
    record.cells[targetCell.fieldId] = appendByMode(record.cells[targetCell.fieldId], value, writeMode);
  } else {
    record.cells[targetCell.fieldId] = value;
  }

  return true;
}

export function applyWrites(datasets, writes) {
  const sourceNextDataset = cloneDataset(datasets.sourceDataset);
  const targetNextDataset = cloneDataset(datasets.targetDataset);

  for (const write of writes ?? []) {
    const dataset = write.target?.journalId === sourceNextDataset.journalId ? sourceNextDataset : targetNextDataset;
    applyWrite(dataset, write.target, write.value, write.writeMode ?? write.write);
  }

  return { sourceNextDataset, targetNextDataset };
}
