import test from 'node:test';
import assert from 'node:assert/strict';

import { createMemoryStorage } from '../../src/storage/storage_iface.js';
import { NAV_KEYS } from '../../src/storage/db_nav.js';
import { createBeckupProvider } from '../src/index.js';

test('createBeckupProvider exports and imports via db-first adapter', async () => {
  const storage = createMemoryStorage();

  await storage.set(NAV_KEYS.spaces, [{ id: 's1', title: 'Space 1' }]);
  await storage.set(NAV_KEYS.journals, [{ id: 'j1', key: 'incoming', title: 'Incoming' }]);
  await storage.set('tableStore:dataset:j1', {
    journalId: 'j1',
    records: [{ id: 'r1', cells: { number: '10' }, subrows: [] }],
    merges: []
  });
  await storage.set('transfer:templates:v1', [{ id: 't1', name: 'T1' }]);

  const provider = createBeckupProvider({ storage });
  const exported = await provider.export({ scope: 'all' });
  assert.equal(exported.format, 'beckup-full-json');
  assert.equal(exported.sections.journals.count, 1);

  const imported = await provider.import(exported, { mode: 'replace' });
  assert.equal(imported.applied, true);
  assert.ok(Array.isArray(imported.warnings));
});
