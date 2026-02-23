import test from 'node:test';
import assert from 'node:assert/strict';

import { createTableEngine } from '../src/modules/table_engine.js';

const schema = {
  id: 'table.schema',
  fields: [
    { key: 'name', label: 'Name', type: 'text', required: true },
    { key: 'amount', label: 'Amount', type: 'number' },
    { key: 'status', label: 'Status', type: 'text' }
  ]
};

function buildEngine() {
  const engine = createTableEngine({
    schema,
    settings: {
      columns: {
        order: ['status', 'name', 'amount'],
        visibility: { amount: true },
        widths: { status: 120 }
      },
      expandedRowIds: ['p1'],
      selectedRowIds: []
    }
  });

  engine.setDataset({
    records: [
      { id: 'p1', cells: { name: 'Parent', amount: 10, status: 'group' }, childrenIds: ['c1'] },
      { id: 'c1', cells: { name: 'Child A', amount: 5, status: 'ok' }, parentId: 'p1' },
      { id: 'r2', cells: { name: 'Standalone', amount: 2, status: 'new' } }
    ],
    merges: [{ rowId: 'p1', colKey: 'status', rowSpan: 2, colSpan: 2 }]
  });

  return engine;
}

test('table engine computes hierarchy + merged cells', () => {
  const engine = buildEngine();
  const view = engine.compute();

  assert.deepEqual(view.columns.map((x) => x.columnKey), ['status', 'name', 'amount']);
  assert.equal(view.columns[0].width, 120);

  assert.deepEqual(view.rows.map((x) => x.rowId), ['p1', 'c1', 'r2']);
  assert.equal(view.rows[0].depth, 0);
  assert.equal(view.rows[1].depth, 1);
  assert.equal(view.rows[0].isExpanded, true);

  assert.deepEqual(view.cellSpanMap.get('p1:status'), { rowSpan: 2, colSpan: 2 });
  assert.deepEqual(view.cellSpanMap.get('c1:status'), { coveredBy: { rowId: 'p1', colKey: 'status' } });
  assert.deepEqual(view.cellSpanMap.get('p1:name'), { coveredBy: { rowId: 'p1', colKey: 'status' } });
});

test('table engine supports sorting, filtering and selection', () => {
  const engine = buildEngine();

  engine.setSort('amount', 'desc');
  let view = engine.compute();
  assert.deepEqual(view.rows.map((x) => x.rowId), ['p1', 'c1', 'r2']);

  engine.setGlobalFilter('stand');
  view = engine.compute();
  assert.deepEqual(view.rows.map((x) => x.rowId), ['r2']);

  engine.toggleSelect('p1');
  engine.toggleSelect('c1');
  view = engine.compute();
  assert.deepEqual(view.selection.sort(), ['c1', 'p1']);
});

test('inline edit and add row form API', () => {
  const engine = buildEngine();

  const edit = engine.beginEdit('c1', 'name');
  assert.equal(edit.rowId, 'c1');

  const patch = engine.applyEdit('c1', 'name', 'Child B');
  assert.deepEqual(patch, {
    recordId: 'c1',
    cellsPatch: { name: 'Child B' },
    fmtPatch: undefined
  });

  const model = engine.getAddFormModel();
  assert.equal(model.length, 3);
  assert.equal(model[0].required, true);

  const invalid = engine.validateAddForm({ amount: 'nan' });
  assert.equal(invalid.valid, false);
  assert.equal(invalid.errors.length, 2);

  const valid = engine.validateAddForm({ name: 'N', amount: 10, status: 'ok' });
  assert.equal(valid.valid, true);

  const record = engine.buildRecordFromForm({ name: 'N' });
  assert.equal(typeof record.id, 'string');
  assert.equal(record.cells.name, 'N');
});

test('table engine hides records with kind=group from visible rows', () => {
  const engine = createTableEngine({
    schema,
    settings: { columns: { order: null, visibility: {}, widths: {} }, expandedRowIds: ['g1'], selectedRowIds: [] }
  });

  engine.setDataset({
    records: [
      { id: 'g1', kind: 'group', cells: { name: 'Group' }, childrenIds: ['r1'] },
      { id: 'r1', kind: 'row', parentId: 'g1', cells: { name: 'Child' } }
    ],
    merges: []
  });

  const view = engine.compute();
  assert.deepEqual(view.rows.map((x) => x.rowId), ['r1']);
});
