function sortByOrder(items) {
  return [...items].sort((a, b) => (a.order ?? 1000) - (b.order ?? 1000) || a.id.localeCompare(b.id));
}

function matchesFilter(item, filter) {
  if (!filter) return true;
  if (typeof filter === 'function') return Boolean(filter(item));
  return Object.entries(filter).every(([key, value]) => item[key] === value);
}

export function createCommandsRegistry(getRuntimeCtx) {
  const commands = new Map();

  function normalizeDefs(input) {
    const defs = typeof input === 'function' ? input(getRuntimeCtx()) : input;
    return Array.isArray(defs) ? defs : [defs];
  }

  return {
    register(input) {
      const defs = normalizeDefs(input);
      const ids = [];
      for (const def of defs) {
        if (!def?.id || !def?.title || typeof def.run !== 'function') {
          throw new Error('CommandDef requires id/title/run(ctx,args)');
        }
        if (commands.has(def.id)) throw new Error(`Command ${def.id} already registered`);
        commands.set(def.id, { order: 1000, ...def });
        ids.push(def.id);
      }
      return () => {
        for (const id of ids) commands.delete(id);
      };
    },
    async run(commandId, args) {
      const def = commands.get(commandId);
      if (!def) throw new Error(`Command not found: ${commandId}`);
      const runtime = getRuntimeCtx();
      if (typeof def.when === 'function' && !def.when(runtime)) {
        throw new Error(`Command is disabled by when(): ${commandId}`);
      }
      return def.run(runtime, args);
    },
    list(filter) {
      return sortByOrder([...commands.values()].filter((x) => matchesFilter(x, filter))).map((x) => ({ ...x }));
    },
    clear() {
      commands.clear();
    }
  };
}
