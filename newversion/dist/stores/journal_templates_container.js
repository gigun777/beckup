const INDEX_KEY = 'templates:index';
const TEMPLATE_KEY_PREFIX = 'templates:tpl:';

function nowIso() {
  return new Date().toISOString();
}

function toTemplateKey(id) {
  return `${TEMPLATE_KEY_PREFIX}${id}`;
}

function validateTemplate(template) {
  if (!template?.id || typeof template.id !== 'string' || /\s/.test(template.id)) {
    throw new Error('Template id is required and must not contain spaces');
  }
  if (!template?.title || typeof template.title !== 'string' || !template.title.trim()) {
    throw new Error('Template title is required');
  }
  if (!Array.isArray(template.columns) || template.columns.length < 1) {
    throw new Error('Template must include at least one column');
  }

  const keys = new Set();
  for (const column of template.columns) {
    if (!column?.key || /\s/.test(column.key)) throw new Error('Column key is required and must not contain spaces');
    if (keys.has(column.key)) throw new Error('Column keys must be unique within template');
    keys.add(column.key);
    if (!column?.label || !String(column.label).trim()) throw new Error('Column label is required');
  }
}

export function createJournalTemplatesContainer(storage) {
  let initialized = false;

  async function ensureInitialized() {
    if (initialized) return;
    const index = (await storage.get(INDEX_KEY)) ?? [];
    if (index.length === 0) {
      const defaultTemplate = {
        id: 'test',
        title: 'test',
        columns: [
          { key: 'c1', label: '1' },
          { key: 'c2', label: '2' },
          { key: 'c3', label: '3' },
          { key: 'c4', label: '4' },
          { key: 'c5', label: '5' }
        ],
        createdAt: nowIso(),
        updatedAt: nowIso()
      };
      await storage.set(INDEX_KEY, [defaultTemplate.id]);
      await storage.set(toTemplateKey(defaultTemplate.id), defaultTemplate);
    }
    initialized = true;
  }

  async function listTemplates() {
    await ensureInitialized();
    const ids = (await storage.get(INDEX_KEY)) ?? [];
    const templates = [];
    for (const id of ids) {
      const template = await storage.get(toTemplateKey(id));
      if (!template) continue;
      templates.push({ id: template.id, title: template.title, columnCount: template.columns.length });
    }
    return templates;
  }

  async function listTemplateEntities() {
    await ensureInitialized();
    const ids = (await storage.get(INDEX_KEY)) ?? [];
    const out = [];
    for (const id of ids) {
      const tpl = await storage.get(toTemplateKey(id));
      if (tpl) out.push(tpl);
    }
    return out;
  }

  async function getTemplate(id) {
    await ensureInitialized();
    if (!id) return null;
    return (await storage.get(toTemplateKey(id))) ?? null;
  }

  async function addTemplate(template) {
    await ensureInitialized();
    validateTemplate(template);

    const ids = (await storage.get(INDEX_KEY)) ?? [];
    if (ids.includes(template.id)) throw new Error(`Template ${template.id} already exists`);

    const prepared = {
      id: template.id,
      title: template.title,
      columns: template.columns.map((c) => ({ key: c.key, label: c.label })),
      createdAt: nowIso(),
      updatedAt: nowIso()
    };

    await storage.set(toTemplateKey(prepared.id), prepared);
    await storage.set(INDEX_KEY, [...ids, prepared.id]);
  }

  async function deleteTemplate(id) {
    await ensureInitialized();
    const ids = (await storage.get(INDEX_KEY)) ?? [];
    if (!ids.includes(id)) return;
    await storage.del(toTemplateKey(id));
    await storage.set(INDEX_KEY, ids.filter((x) => x !== id));
  }

  async function exportDelta(sinceRevision = 0) {
    await ensureInitialized();
    const changes = (await storage.list('templates:')) ?? [];
    return {
      revision: sinceRevision + 1,
      set: Object.fromEntries(changes.map((entry) => [entry.key, entry.value])),
      del: []
    };
  }

  async function applyDelta(patch) {
    for (const [key, value] of Object.entries(patch?.set ?? {})) {
      await storage.set(key, value);
    }
    for (const key of patch?.del ?? []) {
      await storage.del(key);
    }
    initialized = true;
    return { applied: true, warnings: [] };
  }

  return {
    listTemplates,
    getTemplate,
    listTemplateEntities,
    addTemplate,
    deleteTemplate,
    exportDelta,
    applyDelta,
    ensureInitialized,
    keys: { index: INDEX_KEY, prefix: TEMPLATE_KEY_PREFIX }
  };
}
