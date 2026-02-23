import { buildTransferPlan, previewTransferPlan, applyTransferPlan } from '../../../transfer-core/src/index.js';
import { loadTemplates } from '../storage/templates_store.js';

function replacePlaceholders(template, { sourceJournalId, sourceRecordId, targetJournalId, targetRecordId }) {
  const next = structuredClone(template);
  for (const rule of next.rules ?? []) {
    for (const source of rule.sources ?? []) {
      if (source?.cell?.recordId === '__CURRENT__') source.cell.recordId = sourceRecordId;
      if (source?.cell?.journalId === '__CURRENT_JOURNAL__') source.cell.journalId = sourceJournalId;
    }
    for (const target of rule.targets ?? []) {
      if (target?.cell?.recordId === '__TARGET__') target.cell.recordId = targetRecordId;
      if (target?.cell?.journalId === '__TARGET_JOURNAL__') target.cell.journalId = targetJournalId;
    }
  }
  return next;
}

export async function openRunTransferModal({ api, sourceJournalId, recordIds }) {
  const templates = await loadTemplates(api.storageAdapter);
  const journals = await api.listJournals();

  const content = document.createElement('div');
  const targetSelect = document.createElement('select');
  const templateSelect = document.createElement('select');
  const previewBox = document.createElement('pre');

  journals
    .filter((journal) => journal.id !== sourceJournalId)
    .forEach((journal) => targetSelect.append(new Option(journal.title ?? journal.id, journal.id)));
  templates.forEach((template) => templateSelect.append(new Option(template.title, template.id)));

  const runPreview = async () => {
    const templateId = templateSelect.value;
    const targetJournalId = targetSelect.value;
    const sourceRecordId = recordIds[0];

    const sourceDataset = await api.loadDataset(sourceJournalId);
    const targetDataset = await api.loadDataset(targetJournalId);
    const sourceSchema = await api.getSchema(sourceJournalId);
    const targetSchema = await api.getSchema(targetJournalId);

    const targetRecordId = targetDataset.records?.[0]?.id;
    const baseTemplate = templates.find((item) => item.id === templateId);
    const template = replacePlaceholders(baseTemplate, { sourceJournalId, sourceRecordId, targetJournalId, targetRecordId });

    const plan = buildTransferPlan({
      template,
      source: { schema: sourceSchema, dataset: sourceDataset },
      target: { schema: targetSchema, dataset: targetDataset },
      selection: { recordIds },
      context: { currentRecordId: sourceRecordId, targetRecordId }
    });

    const preview = previewTransferPlan(plan);
    previewBox.textContent = JSON.stringify(preview, null, 2);

    return { plan, targetJournalId };
  };

  const previewBtn = document.createElement('button');
  previewBtn.textContent = 'Preview';
  previewBtn.onclick = () => runPreview();

  const applyBtn = document.createElement('button');
  applyBtn.textContent = 'Застосувати';
  applyBtn.onclick = async () => {
    const { plan, targetJournalId } = await runPreview();
    const result = applyTransferPlan(plan);
    if (result.report.errors.length) {
      previewBox.textContent = JSON.stringify(result.report, null, 2);
      return;
    }
    await api.saveDataset(targetJournalId, result.targetNextDataset);
    window.UI.modal.close(modalId);
  };

  const cancelBtn = document.createElement('button');
  cancelBtn.textContent = 'Скасувати';
  cancelBtn.onclick = () => window.UI.modal.close(modalId);

  content.append('Цільовий журнал', targetSelect, 'Шаблон', templateSelect, previewBtn, applyBtn, cancelBtn, previewBox);

  const modalId = window.UI.modal.open({ title: 'Перенести', contentNode: content, closeOnOverlay: true });
  return modalId;
}
