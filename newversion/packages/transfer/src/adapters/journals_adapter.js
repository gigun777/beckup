export function createJournalsAdapter(deps = {}) {
  const {
    listTargets,
    getRow,
    addRow,
    loadDataset,
    saveDataset,
    getSchema,
    listJournals
  } = deps;

  return {
    listTargets: listTargets ?? listJournals ?? (async () => []),
    getRow,
    addRow,
    loadDataset,
    saveDataset,
    getSchema,
    listJournals: listJournals ?? listTargets ?? (async () => [])
  };
}
