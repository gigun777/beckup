import { resolveSources } from './extract.js';
import { evaluateRule } from './rules.js';
import { validate } from './guards.js';
import { buildReport } from './result.js';
import { applyWrites } from './writer.js';

function resolveTargetSpec(target, plan) {
  if (target?.cell) return target.cell;
  if (target?.targetRowFieldId) {
    return {
      journalId: plan.target.dataset.journalId,
      recordId: plan.context?.targetRecordId ?? plan.context?.currentRecordId ?? plan.selection?.recordIds?.[0] ?? null,
      fieldId: target.targetRowFieldId
    };
  }
  return null;
}

function createExecution(plan) {
  const ruleResults = new Map();
  const ctx = {
    sourceDataset: plan.source.dataset,
    targetDataset: plan.target.dataset,
    selection: plan.selection,
    context: plan.context,
    ruleResults
  };

  const steps = (plan.template.rules ?? []).map((rule) => {
    const resolvedSources = resolveSources(rule.sources, ctx);
    const result = evaluateRule(rule, resolvedSources);
    ruleResults.set(rule.id, result.value);
    const resolvedTargets = (rule.targets ?? []).map((target) => resolveTargetSpec(target, plan)).filter(Boolean);
    return { rule, resolvedSources, resolvedTargets, result };
  });

  return { steps };
}

export function buildTransferPlan(input) {
  return {
    template: input.template,
    source: input.source,
    target: input.target,
    selection: input.selection ?? { recordIds: [] },
    context: input.context ?? {}
  };
}

export function previewTransferPlan(plan) {
  const execution = createExecution(plan);
  const validation = validate(execution, plan);
  return buildReport(execution, validation);
}

export function applyTransferPlan(plan) {
  const execution = createExecution(plan);
  const validation = validate(execution, plan);
  const report = buildReport(execution, validation);

  if (!validation.allowed) {
    return {
      sourceNextDataset: plan.source.dataset,
      targetNextDataset: plan.target.dataset,
      report
    };
  }

  const writes = report.writes.map((item) => ({
    target: item.target,
    value: item.value,
    writeMode: item.writeMode
  }));

  const next = applyWrites(
    {
      sourceDataset: plan.source.dataset,
      targetDataset: plan.target.dataset
    },
    writes
  );

  return {
    ...next,
    report
  };
}
