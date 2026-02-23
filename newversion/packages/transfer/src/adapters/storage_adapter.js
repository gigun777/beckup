export function createTransferStorage({ storage, key = 'sdo.transfer.templates.v2' }) {
  if (!storage) throw new Error('storage adapter is required');

  async function readRaw() {
    if (typeof storage.get === 'function') return (await storage.get(key)) ?? [];
    if (typeof storage.getItem === 'function') {
      const raw = storage.getItem(key);
      if (!raw) return [];
      try { return JSON.parse(raw); } catch { return []; }
    }
    return [];
  }

  async function writeRaw(value) {
    if (typeof storage.set === 'function') {
      await storage.set(key, value);
      return;
    }
    if (typeof storage.setItem === 'function') {
      storage.setItem(key, JSON.stringify(value));
      return;
    }
    throw new Error('storage adapter does not support set/setItem');
  }

  return {
    async loadTemplates() {
      return readRaw();
    },
    async saveTemplates(templates) {
      await writeRaw(Array.isArray(templates) ? templates : []);
    },
    async clearTemplates() {
      await writeRaw([]);
    }
  };
}
