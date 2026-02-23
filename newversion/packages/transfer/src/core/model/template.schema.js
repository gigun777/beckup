export const TEMPLATE_SCHEMA_VERSION = 2;

export function migrateTemplates(input) {
  const list = Array.isArray(input) ? input : [];
  return list.map((tpl) => ({
    schemaVersion: TEMPLATE_SCHEMA_VERSION,
    id: tpl.id,
    title: tpl.title ?? 'Untitled template',
    sourceRef: tpl.sourceRef ?? null,
    targetRef: tpl.targetRef ?? null,
    rules: Array.isArray(tpl.rules) ? tpl.rules : [],
    options: tpl.options ?? {}
  }));
}

export function validateTemplate(template) {
  if (!template || typeof template !== 'object') throw new Error('template must be object');
  if (!template.id) throw new Error('template.id is required');
  if (!Array.isArray(template.rules)) throw new Error('template.rules must be array');
  return true;
}
