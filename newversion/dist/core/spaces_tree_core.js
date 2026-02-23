function nowIso() { return new Date().toISOString(); }

export function createSpace(title, parentId = null) {
  return {
    id: crypto.randomUUID(),
    title,
    parentId,
    childCount: 0,
    createdAt: nowIso(),
    updatedAt: nowIso()
  };
}

export function addSpace(nodes, node) {
  const next = [...nodes, node];
  if (node.parentId) {
    const parent = next.find((n) => n.id === node.parentId);
    if (parent) parent.childCount += 1;
  }
  return next;
}

export function deleteSpaceSubtree(nodes, spaceId) {
  const ids = new Set([spaceId]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const node of nodes) {
      if (node.parentId && ids.has(node.parentId) && !ids.has(node.id)) {
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
