export function createJournalStore(templates = []) {
  const byId = new Map(templates.map((tpl) => [tpl.id, tpl]));
  return {
    getTemplate(id) { return byId.get(id) ?? null; },
    list() { return [...byId.values()]; }
  };
}
