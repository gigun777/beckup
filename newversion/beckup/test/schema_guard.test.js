import test from 'node:test';
import assert from 'node:assert/strict';

import {
  normalizeBackupBundle,
  normalizeJournalPayload,
  importFullJsonBackupToSource
} from '../src/index.js';

test('normalizeBackupBundle applies defaults', () => {
  const n = normalizeBackupBundle({ sections: {} });
  assert.equal(n.format, 'beckup-full-json');
  assert.equal(n.formatVersion, 1);
  assert.equal(typeof n.createdAt, 'string');
});

test('normalizeJournalPayload validates and normalizes key/sheet', () => {
  const ok = normalizeJournalPayload({ meta: { type: 'journal', key: 'incoming' }, sheet: { columns: ['A'] } });
  assert.equal(ok.ok, true);
  assert.equal(ok.payload.meta.key, 'incoming');
  assert.equal(ok.payload.sheet.columns[0].name, 'A');

  const bad = normalizeJournalPayload({ meta: { type: 'x' } });
  assert.equal(bad.ok, false);
});

test('import report contains meta + skipped/errors buckets', async () => {
  const target = { async saveJournalPayload() {} };
  const payload = {
    format: 'beckup-full-json',
    formatVersion: 1,
    sections: {
      journals: {
        items: [
          { meta: { type: 'not-journal' } },
          { meta: { type: 'journal', key: 'k1' }, sheet: { columns: [] }, rowsV2: [] }
        ]
      }
    }
  };

  const report = await importFullJsonBackupToSource(payload, { target, mode: 'merge' });
  assert.equal(report.meta.formatVersion, 1);
  assert.equal(report.journals.applied, 1);
  assert.equal(report.journals.skipped, 1);
  assert.equal(Array.isArray(report.journals.errors), true);
});
