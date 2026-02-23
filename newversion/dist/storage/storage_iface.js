export function assertStorage(storage) {
  const methods = ['get', 'set', 'del'];
  for (const method of methods) {
    if (typeof storage?.[method] !== 'function') {
      throw new Error(`Storage adapter must implement ${method}(...)`);
    }
  }
}

export function createMemoryStorage(seed = {}) {
  const db = new Map(Object.entries(seed));
  return {
    async get(key) {
      return db.has(key) ? structuredClone(db.get(key)) : null;
    },
    async set(key, value) {
      db.set(key, structuredClone(value));
    },
    async del(key) {
      db.delete(key);
    },
    async list(prefix = '') {
      return [...db.entries()]
        .filter(([k]) => k.startsWith(prefix))
        .map(([key, value]) => ({ key, value: structuredClone(value) }));
    }
  };
}

export function createLocalStorageStorage(namespace = 'sdo') {
  const hasLS = typeof window !== 'undefined' && window.localStorage;
  if (!hasLS) {
    // Fallback for non-browser environments
    return createMemoryStorage();
  }
  const prefix = `${namespace}:`;
  return {
    async get(key) {
      const raw = window.localStorage.getItem(prefix + key);
      if (raw == null) return null;
      try { return JSON.parse(raw); } catch { return null; }
    },
    async set(key, value) {
      window.localStorage.setItem(prefix + key, JSON.stringify(value));
    },
    async del(key) {
      window.localStorage.removeItem(prefix + key);
    },
    async list(prefixFilter = '') {
      const out = [];
      const fullPrefix = prefix + prefixFilter;
      for (let i = 0; i < window.localStorage.length; i++) {
        const k = window.localStorage.key(i);
        if (!k || !k.startsWith(fullPrefix)) continue;
        const shortKey = k.slice(prefix.length);
        try {
          out.push({ key: shortKey, value: JSON.parse(window.localStorage.getItem(k)) });
        } catch {
          out.push({ key: shortKey, value: null });
        }
      }
      return out;
    }
  };
}
