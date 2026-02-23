const STORAGE_KEY = 'transfer_templates_v1';

async function adapterGet(storageAdapter, key) {
  if (!storageAdapter) return null;
  if (typeof storageAdapter.get === 'function') return storageAdapter.get(key);
  if (typeof storageAdapter.getItem === 'function') {
    const raw = storageAdapter.getItem(key);
    try {
      return JSON.parse(raw);
    } catch {
      return raw;
    }
  }
  return null;
}

async function adapterSet(storageAdapter, key, value) {
  if (!storageAdapter) return;
  if (typeof storageAdapter.set === 'function') {
    await storageAdapter.set(key, value);
    return;
  }
  if (typeof storageAdapter.setItem === 'function') {
    storageAdapter.setItem(key, JSON.stringify(value));
  }
}

export async function loadTemplates(storageAdapter) {
  const raw = (await adapterGet(storageAdapter, STORAGE_KEY)) ?? [];
  return Array.isArray(raw) ? raw : [];
}

export async function saveTemplates(storageAdapter, templates) {
  await adapterSet(storageAdapter, STORAGE_KEY, Array.isArray(templates) ? templates : []);
}

export function createTemplateDraft() {
  return {
    id: `tpl-${Date.now()}`,
    title: 'Новий шаблон',
    rules: []
  };
}

export async function exportTemplatesBackup(storageAdapter) {
  const templates = await loadTemplates(storageAdapter);
  const payload = {
    version: 'transfer-templates.v1',
    createdAt: new Date().toISOString(),
    section: 'transfer',
    templates
  };
  return JSON.stringify(payload, null, 2);
}

export async function importTemplatesBackup(storageAdapter, input) {
  const parsed = typeof input === 'string' ? JSON.parse(input) : input;
  if (!Array.isArray(parsed?.templates)) return false;
  await saveTemplates(storageAdapter, parsed.templates);
  return true;
}
