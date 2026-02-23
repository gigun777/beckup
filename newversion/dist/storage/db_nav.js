export const NAV_KEYS = {
  spaces: 'spaces_nodes_v2',
  journals: 'journals_nodes_v2',
  lastLoc: 'nav_last_loc_v2',
  history: 'nav_history_v2',
  coreSettings: 'core_settings_v2',
  revision: 'core_revision_v2',
  revisionLog: 'core_revision_log_v2'
};

export async function loadNavigationState(storage) {
  return {
    spaces: (await storage.get(NAV_KEYS.spaces)) ?? [],
    journals: (await storage.get(NAV_KEYS.journals)) ?? [],
    lastLoc: (await storage.get(NAV_KEYS.lastLoc)) ?? null,
    history: (await storage.get(NAV_KEYS.history)) ?? []
  };
}

export async function saveNavigationState(storage, state) {
  await storage.set(NAV_KEYS.spaces, state.spaces);
  await storage.set(NAV_KEYS.journals, state.journals);
  await storage.set(NAV_KEYS.lastLoc, state.lastLoc);
  await storage.set(NAV_KEYS.history, state.history.slice(-100));
}
