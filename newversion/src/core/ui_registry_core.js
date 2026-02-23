function sortByOrder(items) {
  return [...items].sort((a, b) => (a.order ?? 1000) - (b.order ?? 1000) || a.id.localeCompare(b.id));
}

function matchesFilter(item, filter) {
  if (!filter) return true;
  if (typeof filter === 'function') return Boolean(filter(item));
  return Object.entries(filter).every(([key, value]) => item[key] === value);
}

export function createUIRegistry() {
  const buttons = new Map();
  const panels = new Map();
  const subscribers = new Set();

  function notify() {
    for (const fn of subscribers) fn();
  }

  function ensureUnique(map, id, type) {
    if (!id || typeof id !== 'string') throw new Error(`${type} id must be a non-empty string`);
    if (map.has(id)) throw new Error(`${type} with id "${id}" already registered`);
  }

  return {
    registerButton(def) {
      if (!def?.label || typeof def.onClick !== 'function') {
        throw new Error('ButtonDef requires label and onClick(ctx)');
      }
      ensureUnique(buttons, def.id, 'button');
      buttons.set(def.id, { order: 1000, location: 'toolbar', ...def });
      notify();
      return () => {
        if (buttons.delete(def.id)) notify();
      };
    },
    registerPanel(def) {
      if (!def?.title || typeof def.render !== 'function') {
        throw new Error('PanelDef requires title and render(mountEl, ctx)');
      }
      ensureUnique(panels, def.id, 'panel');
      panels.set(def.id, { order: 1000, location: 'settings', ...def });
      notify();
      return () => {
        if (panels.delete(def.id)) notify();
      };
    },
    listButtons(filter) {
      return sortByOrder([...buttons.values()].filter((item) => matchesFilter(item, filter)));
    },
    listPanels(filter) {
      return sortByOrder([...panels.values()].filter((item) => matchesFilter(item, filter)));
    },
    subscribe(handler) {
      subscribers.add(handler);
      return () => subscribers.delete(handler);
    },
    clear() {
      buttons.clear();
      panels.clear();
      notify();
    }
  };
}
