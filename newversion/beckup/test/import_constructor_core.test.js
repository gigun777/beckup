import test from 'node:test';
import assert from 'node:assert/strict';

import {
  suggestColumnMapping,
  buildImportPlan,
  applyImportPlanToRows
} from '../src/index.js';

test('import constructor suggests mapping by header and aliases', () => {
  const header = ['Номер док', 'Дата', 'Тема'];
  const targetColumns = [
    { key: 'number', name: 'Номер' },
    { key: 'date', name: 'Дата документа' },
    { key: 'subject', name: 'Короткий зміст' }
  ];
  const aliases = {
    number: ['Номер док'],
    subject: ['Тема']
  };

  const mapping = suggestColumnMapping({ headerRow: header, targetColumns, aliases });
  assert.equal(mapping.length, 2);
  assert.deepEqual(mapping[0], { sourceCol: 1, targetKey: 'number' });
});

test('import plan applies to rows and builds records', () => {
  const rows = [
    ['Номер', 'Дата'],
    ['10', '01.01.25'],
    ['', ''],
    ['11', '02.01.25']
  ];
  const targetColumns = [{ key: 'number', name: 'Номер' }, { key: 'date', name: 'Дата' }];

  const { plan, warnings } = buildImportPlan({
    mapping: [{ sourceCol: 1, targetKey: 'number' }, { sourceCol: 2, targetKey: 'date' }],
    targetColumns
  });
  assert.equal(warnings.length, 0);

  const records = applyImportPlanToRows({ rows, plan, dataRowStartIndex: 1 });
  assert.equal(records.length, 2);
  assert.equal(records[0].cells.number, '10');
  assert.equal(records[1].cells.date, '02.01.25');
});
