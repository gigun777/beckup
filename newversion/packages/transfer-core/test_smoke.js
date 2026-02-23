import assert from 'node:assert/strict';
import { buildTransferPlan, previewTransferPlan, applyTransferPlan } from './src/index.js';

const sourceDataset = {
  journalId: 'A',
  records: [{ id: 'r1', cells: { firstName: 'Іван', lastName: 'Петренко', n1: 10, n2: 2 } }]
};

const targetDataset = {
  journalId: 'B',
  records: [{ id: 't1', cells: { fullName: '', total: 0 } }]
};

const sourceSchema = {
  journalId: 'A',
  fields: [
    { id: 'firstName', title: 'First name', type: 'text' },
    { id: 'lastName', title: 'Last name', type: 'text' },
    { id: 'n1', title: 'N1', type: 'number' },
    { id: 'n2', title: 'N2', type: 'number' }
  ]
};

const targetSchema = {
  journalId: 'B',
  fields: [
    { id: 'fullName', title: 'Full name', type: 'text' },
    { id: 'total', title: 'Total', type: 'number' }
  ]
};

const template = {
  id: 'tpl-1',
  title: 'Transfer sample',
  rules: [
    {
      id: 'rule-1',
      name: 'Concat full name',
      sources: [
        { cell: { journalId: 'A', recordId: 'r1', fieldId: 'firstName' } },
        { cell: { journalId: 'A', recordId: 'r1', fieldId: 'lastName' } }
      ],
      op: 'concat',
      params: { separator: ' ', trim: true, skipEmpty: true },
      targets: [{ cell: { journalId: 'B', recordId: 't1', fieldId: 'fullName' } }],
      write: { mode: 'replace' }
    },
    {
      id: 'rule-2',
      name: 'Math total',
      sources: [
        { cell: { journalId: 'A', recordId: 'r1', fieldId: 'n1' } },
        { cell: { journalId: 'A', recordId: 'r1', fieldId: 'n2' } }
      ],
      op: 'math',
      params: { mathOp: '/', precision: 2, coerceNumeric: 'loose' },
      targets: [{ cell: { journalId: 'B', recordId: 't1', fieldId: 'total' } }],
      write: { mode: 'replace' }
    }
  ]
};

const plan = buildTransferPlan({
  template,
  source: { schema: sourceSchema, dataset: sourceDataset },
  target: { schema: targetSchema, dataset: targetDataset },
  selection: { recordIds: ['r1'] },
  context: { currentRecordId: 'r1', targetRecordId: 't1' }
});

const preview = previewTransferPlan(plan);
assert.equal(preview.errors.length, 0);
assert.equal(preview.rules.length, 2);

const applied = applyTransferPlan(plan);
assert.equal(applied.report.errors.length, 0);
assert.equal(applied.targetNextDataset.records[0].cells.fullName, 'Іван Петренко');
assert.equal(applied.targetNextDataset.records[0].cells.total, 5);

console.log('transfer-core smoke test passed');
