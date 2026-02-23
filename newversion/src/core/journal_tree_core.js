function nowIso() { return new Date().toISOString(); }

export function createJournal({ spaceId, parentId, templateId, title, index }) {
  return {
    id: crypto.randomUUID(),
    spaceId,
    parentId,
    templateId,
    title,
    index,
    childCount: 0,
    createdAt: nowIso(),
    updatedAt: nowIso()
  };
}

export function addJournal(nodes, node) {
  const next = [...nodes, node];
  const parent = next.find((n) => n.id === node.parentId);
  if (parent) parent.childCount += 1;
  return next;
}

export function deleteJournalSubtree(nodes, rootJournalId) {
  const ids = new Set([rootJournalId]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const node of nodes) {
      if (ids.has(node.parentId) && !ids.has(node.id)) {
        ids.add(node.id);
        changed = true;
      }
    }
  }
  return {
    removedIds: ids,
    nodes: nodes.filter((n) => !ids.has(n.id)).map((n) => ({ ...n, childCount: 0 }))
  };
}

export function journalsForSpace(nodes, spaceId) {
  return nodes.filter((j) => j.spaceId === spaceId);
}
