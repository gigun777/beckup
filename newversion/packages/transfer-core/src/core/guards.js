function getSchemaByJournal(plan, journalId) {
  if (plan.source.schema?.journalId === journalId) return plan.source.schema;
  if (plan.target.schema?.journalId === journalId) return plan.target.schema;
  return null;
}

function hasField(schema, fieldId) {
  return Boolean(schema?.fields?.some((field) => field.id === fieldId));
}

function isForbiddenTarget(target, ctx) {
  const forbidden = ctx?.policies?.forbiddenTargetFieldIds ?? [];
  if (target?.fieldId) return forbidden.includes(target.fieldId);
  return false;
}

export function validate(execution, plan) {
  const errors = [];
  const warnings = [];
  const ctx = plan.context ?? {};

  for (const step of execution.steps ?? []) {
    if (!step.result?.ok) {
      errors.push({ ruleId: step.rule.id, code: step.result?.error ?? 'rule_evaluation_failed' });
    }

    if (!step.resolvedTargets?.length) {
      errors.push({ ruleId: step.rule.id, code: 'empty_target' });
      continue;
    }

    for (const target of step.resolvedTargets) {
      if (!target?.fieldId) {
        errors.push({ ruleId: step.rule.id, code: 'empty_target' });
        continue;
      }

      const schema = getSchemaByJournal(plan, target.journalId);
      if (!hasField(schema, target.fieldId)) {
        errors.push({ ruleId: step.rule.id, code: 'target_field_not_found', target });
      }
      if (isForbiddenTarget(target, ctx)) {
        errors.push({ ruleId: step.rule.id, code: 'forbidden_target', target });
      }
    }

    if (step.result?.ok && (step.result.value === '' || step.result.value == null)) {
      warnings.push({ ruleId: step.rule.id, code: 'empty_result' });
    }
  }

  return { allowed: errors.length === 0, errors, warnings };
}
