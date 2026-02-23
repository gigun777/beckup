import test from 'node:test';
import assert from 'node:assert/strict';

import { createMemoryStorage } from '../../src/storage/storage_iface.js';
import { NAV_KEYS } from '../../src/storage/db_nav.js';
import { createNewversionSourceAdapter } from '../src/index.js';

test('source adapter restores journal from legacy rows fallback when rowsV2 missing', async () => {
  const storage = createMemoryStorage();

  await storage.set(NAV_KEYS.journals, [{ id: 'j1', key: 'incoming', title: 'Incoming' }]);
  const adapter = createNewversionSourceAdapter(storage);

  const payload = {
    meta: { type: 'journal', key: 'incoming' },
    sheet: { columns: [{ name: 'Номер' }, { name: 'Дата' }] },
    rows: [
      { id: 'r1', exportData: { 'Номер': '10', 'Дата': '01.01.25' } },
      { id: 'r2', data: { 'Номер': '11', 'Дата': '02.01.25' } }
    ]
  };

  await adapter.saveJournalPayload('incoming', payload, { mode: 'replace' });
  const ds = await storage.get('tableStore:dataset:j1');

  assert.equal(Array.isArray(ds.records), true);
  assert.equal(ds.records.length, 2);
  assert.equal(ds.records[0].cells['Номер'], '10');
  assert.equal(ds.records[1].cells['Дата'], '02.01.25');
});
