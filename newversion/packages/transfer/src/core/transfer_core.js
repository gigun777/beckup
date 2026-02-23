import { buildTransferPlan, previewTransferPlan, applyTransferPlan } from '../../../transfer-core/src/index.js';
import { migrateTemplates, validateTemplate } from './model/template.schema.js';

export function createTransferCore({ storage, journals, logger } = {}) {
  if (!storage) throw new Error('createTransferCore requires storage adapter');
  if (!journals) throw new Error('createTransferCore requires journals adapter');

  async function loadTemplates() {
    return migrateTemplates(await storage.loadTemplates());
  }

  async function saveTemplates(next) {
    await storage.saveTemplates(migrateTemplates(next));
  }

  const templatesApi = {
    async list() {
      return loadTemplates();
    },
    async save(template) {
      validateTemplate(template);
      const current = await loadTemplates();
      const idx = current.findIndex((item) => item.id === template.id);
      if (idx >= 0) current[idx] = { ...template, schemaVersion: 2 };
      else current.push({ ...template, schemaVersion: 2 });
      await saveTemplates(current);
      return template;
    },
    async remove(id) {
      const current = await loadTemplates();
      await saveTemplates(current.filter((item) => item.id !== id));
    }
  };

  async function prepareTransfer({ templateId, sourceRef, rowIds }) {
    const templates = await loadTemplates();
    const template = templates.find((item) => item.id === templateId);
    if (!template) throw new Error(`template not found: ${templateId}`);

    const sourceDataset = await journals.loadDataset(sourceRef.journalId);
    const sourceSchema = await journals.getSchema(sourceRef.journalId);

    return {
      template,
      sourceRef,
      rowIds: Array.isArray(rowIds) ? rowIds : [rowIds],
      sourceDataset,
      sourceSchema
    };
  }

  async function preview(ctx, { targetRef }) {
    const targetDataset = await journals.loadDataset(targetRef.journalId);
    const targetSchema = await journals.getSchema(targetRef.journalId);
    const currentRecordId = ctx.rowIds?.[0] ?? null;
    const targetRecordId = targetDataset.records?.[0]?.id ?? null;

    const plan = buildTransferPlan({
      template: ctx.template,
      source: { schema: ctx.sourceSchema, dataset: ctx.sourceDataset },
      target: { schema: targetSchema, dataset: targetDataset },
      selection: { recordIds: ctx.rowIds },
      context: { currentRecordId, targetRecordId }
    });

    const report = previewTransferPlan(plan);
    return { plan, report, targetDataset, targetSchema };
  }

  async function commit(previewCtx) {
    const applied = applyTransferPlan(previewCtx.plan);
    if (applied.report.errors.length) {
      logger?.error?.('transfer commit blocked', applied.report.errors);
      return applied;
    }

    await journals.saveDataset(previewCtx.plan.target.dataset.journalId, applied.targetNextDataset);
    return applied;
  }

  return {
    templates: templatesApi,
    prepareTransfer,
    preview,
    commit
  };
}
