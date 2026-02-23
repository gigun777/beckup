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


test('provider import exposes hasErrors and applied=false when import errors exist', async () => {
  const storage = createMemoryStorage();
  const provider = createBeckupProvider({ storage });

  const payload = {
    sections: {
      settings: { payload: { theme: 'dark' } }
    }
  };

  // Force an adapter-level error path.
  const originalSet = storage.set.bind(storage);
  storage.set = async (key, value) => {
    if (key === 'core_settings_v2') throw new Error('settings write failed');
    return originalSet(key, value);
  };

  const imported = await provider.import(payload, { mode: 'merge' });
  assert.equal(imported.applied, false);
  assert.equal(imported.hasErrors, true);
  assert.ok(imported.warnings.some((w) => w.includes('settings write failed')));
});


test('provider import keeps applied=true when only warnings exist', async () => {
  const storage = createMemoryStorage();
  const provider = createBeckupProvider({ storage });

  const payload = {
    sections: {
      journals: {
        items: [{ meta: { type: 'not-journal' } }]
      }
    }
  };

  const imported = await provider.import(payload, { mode: 'merge' });
  assert.equal(imported.applied, true);
  assert.equal(imported.hasErrors, false);
  assert.ok(imported.warnings.some((w) => w.includes('Skipped journal')));
});
