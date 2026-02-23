function sortByOrder(items) {
  return [...items].sort((a, b) => (a.order ?? 1000) - (b.order ?? 1000) || a.id.localeCompare(b.id));
}

function matchesFilter(item, filter) {
  if (!filter) return true;
  if (typeof filter === 'function') return Boolean(filter(item));
  return Object.entries(filter).every(([key, value]) => item[key] === value);
}

function toNumVersion(version) {
  const n = Number(version);
  if (Number.isFinite(n)) return n;
  const main = String(version).split('.').map((x) => Number(x));
  if (main.some((x) => !Number.isFinite(x))) return null;
  return main[0] * 1_000_000 + (main[1] ?? 0) * 1_000 + (main[2] ?? 0);
}

export function createSchemaRegistry() {
  const schemas = new Map();

  return {
    register(def) {
      if (!def?.id || !def?.version || !def?.domain || !Array.isArray(def.fields)) {
        throw new Error('SchemaDef requires id/version/domain/fields');
      }
      const existing = schemas.get(def.id);
      if (existing) {
        const prev = toNumVersion(existing.version);
        const next = toNumVersion(def.version);
        if (prev == null || next == null || next <= prev) {
          throw new Error(`Schema ${def.id} already registered with same or newer version`);
        }
      }
      schemas.set(def.id, def);
      return () => schemas.delete(def.id);
    },
    get(id) {
      return schemas.has(id) ? schemas.get(id) : null;
    },
    list(filter) {
      return sortByOrder([...schemas.values()].filter((x) => matchesFilter(x, filter))).map((x) => x);
    },
    resolve(target) {
      const list = this.list();
      return list.find((schema) => {
        const a = schema.appliesTo ?? { any: true };
        return a.any || a.templateId === target || a.journalId === target || a.spaceId === target;
      }) ?? null;
    },
    clear() {
      schemas.clear();
    }
  };
}
