import test from 'node:test';
import assert from 'node:assert/strict';

import { formatCell, parseInput } from '../src/modules/table_formatter.js';

test('formatCell formats values by type', () => {
  const numberCell = formatCell({ t: 'number', v: 12345.6 }, {}, { type: 'number' }, { locale: 'uk-UA' });
  assert.equal(numberCell.text, '12 345,6');
  assert.equal(numberCell.align, 'right');

  const dateCell = formatCell({ t: 'date', v: '2025-02-16' }, {}, { type: 'date' }, { dateFormat: 'DD.MM.YYYY' });
  assert.equal(dateCell.text, '16.02.2025');

  const boolCell = formatCell({ t: 'bool', v: true }, {}, { type: 'bool' });
  assert.equal(boolCell.text, 'Так');

  const enumCell = formatCell({ t: 'enum', v: 'draft' }, {}, {
    type: 'enum',
    options: [{ value: 'draft', label: 'Чернетка' }]
  });
  assert.equal(enumCell.text, 'Чернетка');
});

test('formatCell applies fmt styles and conditional rules, ignores unknown ext', () => {
  const cell = formatCell(
    { t: 'text', v: '' },
    {
      align: 'center',
      bold: true,
      color: '#111',
      bg: '#eee',
      wrap: false,
      ext: { unknownProp: { nested: true } },
      rules: [{ when: { empty: true }, style: { color: 'red' } }]
    },
    {
      type: 'text',
      rules: [{ when: { equals: 'X' }, style: { bg: 'yellow' } }]
    }
  );

  assert.equal(cell.align, 'center');
  assert.equal(cell.style.fontWeight, '700');
  assert.equal(cell.style.backgroundColor, '#eee');
  assert.equal(cell.style.color, 'red');
  assert.equal(cell.style.whiteSpace, 'nowrap');
});

test('parseInput returns valid CellValue by type', () => {
  assert.deepEqual(parseInput('123,4', { type: 'number' }), { t: 'number', v: 123.4 });
  assert.deepEqual(parseInput('так', { type: 'bool' }), { t: 'bool', v: true });
  assert.deepEqual(parseInput('16.02.2025', { type: 'date' }), { t: 'date', v: '2025-02-16' });
  assert.deepEqual(parseInput('Чернетка', {
    type: 'enum',
    options: [{ value: 'draft', label: 'Чернетка' }]
  }), { t: 'enum', v: 'draft' });
  assert.deepEqual(parseInput('hello', { type: 'text' }), { t: 'text', v: 'hello' });
});
