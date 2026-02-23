function toSubrowsDataset(dataset = {}) {
  const recordsArray = Array.isArray(dataset.records) ? dataset.records : [];
  const records = Object.fromEntries(recordsArray.map((record) => [record.id, { ...record }]));
  const order = recordsArray.map((record) => record.id);
  return { records, order };
}

function fromSubrowsDataset(baseDataset = {}, subrowsDataset) {
  const order = Array.isArray(subrowsDataset?.order) ? subrowsDataset.order : [];
  const recordsMap = subrowsDataset?.records ?? {};
  const records = order
    .map((id) => recordsMap[id])
    .filter(Boolean)
    .map((record) => ({ ...record }));

  return {
    ...baseDataset,
    records
  };
}

function assertApi(api) {
  const required = [
    'computeVisibleRows',
    'resolveEditTarget',
    'handleCellClickSubrowsFlow',
    'addSubrow',
    'removeSubrow',
    'getTransferCandidates'
  ];

  for (const key of required) {
    if (typeof api?.[key] !== 'function') {
      throw new Error(`Subrows api is missing function: ${key}`);
    }
  }
}

export function createTableSubrowsBridge(subrowsApi, defaultSettings = { columnsSubrowsEnabled: {} }) {
  assertApi(subrowsApi);

  function resolveSettings(settingsOverride) {
    return settingsOverride ?? defaultSettings;
  }

  function computeVisibleRows(dataset, settingsOverride) {
    const settings = resolveSettings(settingsOverride);
    const ds = toSubrowsDataset(dataset);
    return subrowsApi.computeVisibleRows(ds, settings);
  }

  function resolveEditTarget(dataset, cellRef, settingsOverride) {
    const settings = resolveSettings(settingsOverride);
    const ds = toSubrowsDataset(dataset);
    return subrowsApi.resolveEditTarget(ds, cellRef, settings);
  }

  async function handleCellClickSubrowsFlow({ dataset, cellRef, settings, ui }) {
    const resolvedSettings = resolveSettings(settings);
    const ds = toSubrowsDataset(dataset);
    const result = await subrowsApi.handleCellClickSubrowsFlow({
      ds,
      cellRef,
      settings: resolvedSettings,
      ui
    });

    return {
      ...result,
      dataset: fromSubrowsDataset(dataset, result.dataset)
    };
  }

  function addSubrow(dataset, rowId, settingsOverride) {
    const settings = resolveSettings(settingsOverride);
    const ds = toSubrowsDataset(dataset);
    const result = subrowsApi.addSubrow(ds, rowId, settings);

    return {
      rowId: result.value,
      dataset: fromSubrowsDataset(dataset, result.dataset)
    };
  }

  function removeSubrow(dataset, subrowId) {
    const ds = toSubrowsDataset(dataset);
    const result = subrowsApi.removeSubrow(ds, subrowId);

    return {
      removed: result.value,
      dataset: fromSubrowsDataset(dataset, result.dataset)
    };
  }

  function getTransferCandidates(dataset, rowId) {
    return subrowsApi.getTransferCandidates(toSubrowsDataset(dataset), rowId);
  }

  return {
    computeVisibleRows,
    resolveEditTarget,
    handleCellClickSubrowsFlow,
    addSubrow,
    removeSubrow,
    getTransferCandidates,
    adapters: {
      toSubrowsDataset,
      fromSubrowsDataset
    }
  };
}
