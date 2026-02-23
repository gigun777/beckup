import test from 'node:test';
import assert from 'node:assert/strict';

import { createZipBackup } from '../src/index.js';

test('createZipBackup returns PK zip bytes', () => {
  const bytes = createZipBackup([
    { name: 'a.txt', data: 'hello' },
    { name: 'b.json', data: { x: 1 } }
  ]);

  assert.ok(bytes instanceof Uint8Array);
  assert.equal(bytes[0], 0x50); // P
  assert.equal(bytes[1], 0x4b); // K
  assert.ok(bytes.length > 40);
});
