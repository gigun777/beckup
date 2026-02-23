import test from 'node:test';
import assert from 'node:assert/strict';

import {
  createSheetJsonPayload,
  createExcelJsonPayload,
  createSheetZipBackup
} from '../src/index.js';

const sheet = {
  key: 'incoming',
  title: 'Вхідні',
  columns: [
    { key: 'n', name: 'Номер' },
    { key: 'd', name: 'Дата' }
  ]
};

const records = [
  { id: 'r1', cells: { n: 10, d: '01.01.25' } },
  { id: 'r2', data: { 'Номер': '11', 'Дата': '02.01.25' } }
];

test('createSheetJsonPayload builds journal v2 payload', () => {
  const out = createSheetJsonPayload({ sheet, records });
  assert.equal(out.meta.type, 'journal');
  assert.equal(out.meta.version, 2);
  assert.equal(out.columnsCount, 2);
  assert.equal(out.rowsV2.length, 2);
  assert.equal(out.rowsV2[0].cells[0], '10');
  assert.equal(out.rows[1].exportData['Номер'], '11');
});

test('createExcelJsonPayload builds excel-json projection', () => {
  const out = createExcelJsonPayload({ sheet, records });
  assert.equal(out.format, 'beckup-excel-json');
  assert.equal(out.sheet.columns.length, 2);
  assert.equal(out.rows.length, 2);
  assert.equal(out.matrix[0][0], '10');
});

test('createSheetZipBackup returns zip + payloads', () => {
  const out = createSheetZipBackup({ sheet, records });
  assert.ok(out.zipBytes instanceof Uint8Array);
  assert.equal(out.payloads.journal.meta.type, 'journal');
  assert.equal(out.payloads.excel.format, 'beckup-excel-json');
  assert.equal(out.files.length, 2);
});
