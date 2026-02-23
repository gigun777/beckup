import test from 'node:test';
import assert from 'node:assert/strict';

import { getRenderableCells } from '../src/modules/table_renderer.js';

test('getRenderableCells skips covered cells and keeps top-left span', () => {
  const row = { rowId: 'r1' };
  const columns = [{ columnKey: 'a' }, { columnKey: 'b' }, { columnKey: 'c' }];
  const cellSpanMap = new Map([
    ['r1:a', { rowSpan: 2, colSpan: 2 }],
    ['r1:b', { coveredBy: { rowId: 'r1', colKey: 'a' } }]
  ]);

  const cells = getRenderableCells(row, columns, cellSpanMap);
  assert.deepEqual(cells, [
    { colKey: 'a', span: { rowSpan: 2, colSpan: 2 } },
    { colKey: 'c', span: { rowSpan: 1, colSpan: 1 } }
  ]);
});
