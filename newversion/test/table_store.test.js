import test from 'node:test';
import assert from 'node:assert/strict';

import { createSEDO, createMemoryStorage } from '../src/index.js';
import { createTableStoreModule, tableStoreKeys } from '../src/modules/table_store.js';

test('table store CRUD + formatting persistence', async () => {
  const sdo = createSEDO({ storage: createMemoryStorage(), modules: [createTableStoreModule()] });
  await sdo.start();

  const id = await sdo.api.tableStore.addRecord('j1', {
    cells: { a: { t: 'text', v: 'X' } },
    fmt: { a: { color: '#111', ext: { unknown: { deep: true } } } },
    rowFmt: { bg: '#eee', ext: { any: 1 } }
  });

  assert.equal(typeof id, 'string');
  const ds = await sdo.api.tableStore.getDataset('j1');
  assert.equal(ds.records.length, 1);
  assert.deepEqual(ds.records[0].fmt.a.ext, { unknown: { deep: true } });
  assert.deepEqual(ds.records[0].rowFmt.ext, { any: 1 });

  await sdo.api.tableStore.updateRecord('j1', id, { cells: { a: { t: 'text', v: 'Y' } } });
  const updated = await sdo.api.tableStore.getDataset('j1');
  assert.equal(updated.records[0].cells.a.v, 'Y');

  await sdo.api.tableStore.deleteRecord('j1', id);
  const afterDelete = await sdo.api.tableStore.getDataset('j1');
  assert.equal(afterDelete.records.length, 0);
});

test('table store export/import merge replace and includeFormatting flag', async () => {
  const sdo = createSEDO({ storage: createMemoryStorage(), modules: [createTableStoreModule()] });
  await sdo.start();

  await sdo.api.tableStore.addRecord('j1', {
    id: 'r1',
    cells: { a: { t: 'text', v: 'A' } },
    fmt: { a: { color: '#000', ext: { keep: true } } }
  });

  const full = await sdo.api.tableStore.exportTableData({ includeFormatting: true });
  assert.equal(full.format, 'sdo-table-data');
  assert.deepEqual(full.datasets[0].records[0].fmt.a.ext, { keep: true });

  const plain = await sdo.api.tableStore.exportTableData({ includeFormatting: false });
  assert.equal(plain.datasets[0].records[0].fmt, undefined);

  const mergeBundle = {
    format: 'sdo-table-data',
    formatVersion: 1,
    exportedAt: new Date().toISOString(),
    datasets: [{
      journalId: 'j1',
      records: [{ id: 'r2', cells: { a: { t: 'text', v: 'B' } }, fmt: { a: { ext: { z: 1 } } } }],
      meta: { revision: 0 }
    }]
  };
  await sdo.api.tableStore.importTableData(mergeBundle, { mode: 'merge' });
  const merged = await sdo.api.tableStore.getDataset('j1');
  assert.equal(merged.records.length, 2);

  const replaceBundle = {
    format: 'sdo-table-data',
    formatVersion: 1,
    exportedAt: new Date().toISOString(),
    datasets: [{ journalId: 'j1', records: [{ id: 'r3', cells: { a: { t: 'text', v: 'C' } } }], meta: { revision: 0 } }]
  };
  await sdo.api.tableStore.importTableData(replaceBundle, { mode: 'replace' });
  const replaced = await sdo.api.tableStore.getDataset('j1');
  assert.deepEqual(replaced.records.map((r) => r.id), ['r3']);
});

test('table store backup provider and delta', async () => {
  const storage = createMemoryStorage();
  const sdo = createSEDO({ storage, modules: [createTableStoreModule()] });
  await sdo.start();

  await sdo.api.tableStore.addRecord('j1', { id: 'r1', cells: { a: { t: 'num', v: 1 } }, fmt: { a: { ext: { q: 7 } } } });
  const providersExport = await sdo.exportBackup({ modules: ['tableStore'], includeUserData: true });
  assert.equal(providersExport.modules.tableStore.data.userData.format, 'sdo-table-data');

  const delta = await sdo.api.tableStore.exportDelta({ sinceRev: 0 });
  assert.ok(Object.keys(delta.set).includes(tableStoreKeys.dataset('j1')));

  const sdo2 = createSEDO({ storage: createMemoryStorage(), modules: [createTableStoreModule()] });
  await sdo2.start();
  const report = await sdo2.api.tableStore.applyDelta(delta, { mode: 'replace' });
  assert.equal(report.applied, true);
  const imported = await sdo2.api.tableStore.getDataset('j1');
  assert.equal(imported.records[0].fmt.a.ext.q, 7);
});
