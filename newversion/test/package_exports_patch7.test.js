import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const ROOT = process.cwd();
const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));

const PATCH7_EXPORT_KEYS = [
  './ui',
  './ui/theme.css',
  './ui/styles.css',
  './ui/theme',
  './ui/manager',
  './ui/controls',
  './ui/modal',
  './ui/form',
  './ui/toast',
  './ui/backup',
  './ui/bootstrap',
  './ui/core',
  './ui/settings/init',
  './ui/settings/shell',
  './ui/settings/registry',
  './ui/settings/state',
  './ui/settings/features/table',
  './ui/settings/features/uxui',
  './ui/settings/features/backup'
];

test('Patch 7 package exports keys are present', () => {
  for (const key of PATCH7_EXPORT_KEYS) {
    assert.equal(typeof pkg.exports[key], 'string', `Missing exports entry: ${key}`);
  }
});

test('Patch 7 export targets exist on disk', () => {
  for (const key of PATCH7_EXPORT_KEYS) {
    const target = pkg.exports[key];
    const abs = path.join(ROOT, target.replace(/^\.\//, ''));
    assert.equal(fs.existsSync(abs), true, `Export target does not exist for ${key}: ${target}`);
  }
});

test('Patch 7 JS entrypoints are importable from dist', async () => {
  const jsKeys = PATCH7_EXPORT_KEYS.filter((key) => !key.endsWith('.css'));
  for (const key of jsKeys) {
    const target = pkg.exports[key];
    const abs = path.join(ROOT, target.replace(/^\.\//, ''));
    await import(pathToFileURL(abs).href);
  }
  assert.ok(true);
});
