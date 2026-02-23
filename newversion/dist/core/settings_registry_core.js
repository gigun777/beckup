function sortByOrder(items) {
  return [...items].sort((a, b) => (a.order ?? 1000) - (b.order ?? 1000) || a.id.localeCompare(b.id));
}

export function createSettingsRegistry() {
  const defs = new Map();

  return {
    register(def) {
      if (!def?.id || !def?.tab?.id || !def?.tab?.title || !Array.isArray(def.fields)) {
        throw new Error('SettingsDef requires id/tab/fields');
      }
      if (defs.has(def.id)) throw new Error(`Settings ${def.id} already registered`);
      defs.set(def.id, def);
      return () => defs.delete(def.id);
    },
    list() {
      return [...defs.values()].map((x) => x);
    },
    listTabs() {
      const tabs = new Map();
      for (const def of defs.values()) {
        const tabId = def.tab.id;
        const existing = tabs.get(tabId) ?? { id: tabId, title: def.tab.title, order: def.tab.order ?? 1000, items: [] };
        existing.items.push(def);
        tabs.set(tabId, existing);
      }
      return sortByOrder([...tabs.values()]).map((tab) => ({ ...tab, items: sortByOrder(tab.items) }));
    },
    clear() {
      defs.clear();
    }
  };
}
