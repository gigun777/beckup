import test from 'node:test';
import assert from 'node:assert/strict';

import { createTableSubrowsBridge } from '../src/modules/table_subrows_bridge.js';

function createFakeApi() {
  return {
    computeVisibleRows(ds) {
      return ds.order.map((id) => ({ id }));
    },
    resolveEditTarget() {
      return { type: 'normalEdit', targetRowId: 'r1' };
    },
    async handleCellClickSubrowsFlow({ ds }) {
      return {
        dataset: ds,
        addedSubrowId: null,
        editTargetRowId: null,
        highlightSubrows: [],
        needsChoice: false
      };
    },
    addSubrow(ds) {
      const id = 'r2';
      const next = {
        records: {
          ...ds.records,
          r2: { id: 'r2', kind: 'row', parentId: 'g1', cells: {} }
        },
        order: [...ds.order, id]
      };
      return { dataset: next, value: id };
    },
    removeSubrow(ds, subrowId) {
      const next = {
        records: Object.fromEntries(Object.entries(ds.records).filter(([id]) => id !== subrowId)),
        order: ds.order.filter((id) => id !== subrowId)
      };
      return { dataset: next, value: true };
    },
    getTransferCandidates(ds, rowId) {
      return ds.records[rowId] ? [rowId] : [];
    }
  };
}

test('table subrows bridge adapts array dataset to map dataset', () => {
  const bridge = createTableSubrowsBridge(createFakeApi(), { columnsSubrowsEnabled: {} });
  const dataset = {
    records: [{ id: 'g1', kind: 'group', parentId: null, childrenIds: ['r1'], cells: {} }, { id: 'r1', kind: 'row', parentId: 'g1', cells: {} }]
  };

  const visible = bridge.computeVisibleRows(dataset);
  assert.deepEqual(visible, [{ id: 'g1' }, { id: 'r1' }]);

  const added = bridge.addSubrow(dataset, 'g1');
  assert.equal(added.rowId, 'r2');
  assert.equal(added.dataset.records.at(-1).id, 'r2');

  const removed = bridge.removeSubrow(added.dataset, 'r2');
  assert.equal(removed.removed, true);
  assert.equal(removed.dataset.records.some((r) => r.id === 'r2'), false);
});
