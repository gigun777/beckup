import test from 'node:test';
import assert from 'node:assert/strict';

import { exportFullJsonBackupFromSource, importFullJsonBackupToSource } from '../src/index.js';

test('export/import full json backup tolerates partial sections', async () => {
  const source = {
    async listJournals() { return [{ id: 'j1', key: 'incoming', title: 'Incoming' }]; },
    async loadJournalRecords() { return [{ id: 'r1', cells: { A: '1' } }]; },
    async loadJournalSchema() { return { columns: [{ name: 'A' }] }; },
    async loadSettings() { return { theme: 'dark' }; }
  };

  const backup = await exportFullJsonBackupFromSource({ source, include: { journals: true, settings: true, navigation: false, transfer: false } });
  assert.equal(backup.format, 'beckup-full-json');
  assert.equal(backup.sections.journals.items.length, 1);

  const saved = { journals: [] };
  const target = {
    async saveJournalPayload(journalKey, payload) { saved.journals.push({ journalKey, payload }); },
    async saveSettings(v) { saved.settings = v; }
  };

  const report = await importFullJsonBackupToSource(backup, { target, mode: 'merge' });
  assert.equal(report.journals.applied, 1);
  assert.equal(saved.journals.length, 1);
  assert.equal(saved.settings.theme, 'dark');

  // Partial payload (only journals) should still import journals.
  const partial = { sections: { journals: backup.sections.journals } };
  const report2 = await importFullJsonBackupToSource(partial, { target, mode: 'merge' });
  assert.equal(report2.journals.applied, 1);
});


test('import uses normalized sections object for optional blocks', async () => {
  const target = {
    async saveJournalPayload() {}
  };

  const report = await importFullJsonBackupToSource({ sections: 'invalid-shape' }, { target, mode: 'merge' });
  assert.equal(report.settings.applied, false);
  assert.equal(report.navigation.applied, false);
  assert.equal(report.transfer.applied, false);
  assert.deepEqual(report.settings.errors, []);
});
