import { VERSION, BACKUP_FORMAT, DELTA_BACKUP_FORMAT, ENCRYPTED_BACKUP_FORMAT } from './types/public.js';
import { assertStorage, createMemoryStorage, createLocalStorageStorage } from './storage/storage_iface.js';
import { NAV_KEYS, loadNavigationState, saveNavigationState } from './storage/db_nav.js';
import { normalizeLocation } from './core/level_model_core.js';
import { pushHistory } from './core/navigation_core.js';
import { createUIRegistry } from './core/ui_registry_core.js';
import { createSchemaRegistry } from './core/schema_registry_core.js';
import { createCommandsRegistry } from './core/commands_registry_core.js';
import { createSettingsRegistry } from './core/settings_registry_core.js';
import { createJournalTemplatesContainer } from './stores/journal_templates_container.js';
import { createIntegrity, decryptBackup, encryptBackup, verifyIntegrity } from './backup/crypto.js';
export { assertStorage, createMemoryStorage, createLocalStorageStorage } from './storage/storage_iface.js';


function deepFreeze(obj) {
  if (!obj || typeof obj !== 'object') return obj;
  Object.freeze(obj);
  for (const value of Object.values(obj)) deepFreeze(value);
  return obj;
}

function toNavPayload(nav) {
  return {
    spaces_nodes_v2: nav.spaces ?? [],
    journals_nodes_v2: nav.journals ?? [],
    nav_last_loc_v2: nav.lastLoc ?? null,
    nav_history_v2: nav.history ?? []
  };
}

function fromNavPayload(payload) {
  return {
    spaces: payload.spaces_nodes_v2 ?? [],
    journals: payload.journals_nodes_v2 ?? [],
    lastLoc: payload.nav_last_loc_v2 ?? null,
    history: payload.nav_history_v2 ?? []
  };
}

export function createSEDO(options = {}) {
  const storage = options.storage ?? createMemoryStorage();
  assertStorage(storage);

  const listeners = new Map();
  const modules = new Map();
  const backupProviders = new Map();
  const moduleDisposers = new Map();
  const uiRegistry = createUIRegistry();
  const schemaRegistry = createSchemaRegistry();
  const settingsRegistry = createSettingsRegistry();
  const journalTemplates = createJournalTemplatesContainer(storage);

  const state = {
    spaces: [],
    journals: [],
    history: [],
    activeSpaceId: null,
    activeJournalId: null,
    started: false,
    revision: 0
  };

  function getRuntimeCtx() {
    return { api, storage, sdo: instance };
  }
  const commandsRegistry = createCommandsRegistry(getRuntimeCtx);

  let ui = null;

  function emit(event, payload) {
    for (const fn of listeners.get(event) ?? []) fn(payload);
  }

  async function bumpRevision(changedKeys = []) {
    state.revision += 1;
    await storage.set(NAV_KEYS.revision, state.revision);
    const log = (await storage.get(NAV_KEYS.revisionLog)) ?? [];
    log.push({ rev: state.revision, changedKeys, at: new Date().toISOString() });
    await storage.set(NAV_KEYS.revisionLog, log.slice(-500));
  }

  const api = {
    getState: () => deepFreeze(structuredClone(state)),
    dispatch(action) {
      if (typeof action?.reduce !== 'function') throw new Error('Action must include reduce(state)');
      action.reduce(state);
      emit('state:changed', api.getState());
      return state;
    }
  };

  function createModuleUIApi(moduleId) {
    const disposers = moduleDisposers.get(moduleId) ?? [];
    moduleDisposers.set(moduleId, disposers);
    function track(unregisterFn) {
      disposers.push(unregisterFn);
      return () => {
        unregisterFn();
        const idx = disposers.indexOf(unregisterFn);
        if (idx >= 0) disposers.splice(idx, 1);
      };
    }
    return {
      registerButton(def) { return track(uiRegistry.registerButton({ ...def })); },
      registerPanel(def) { return track(uiRegistry.registerPanel({ ...def })); },
      listButtons(filter) { return uiRegistry.listButtons(filter); },
      listPanels(filter) { return uiRegistry.listPanels(filter); }
    };
  }

  function createModuleCtx(moduleId) {
    const disposers = moduleDisposers.get(moduleId) ?? [];
    moduleDisposers.set(moduleId, disposers);
    const track = (fn) => {
      disposers.push(fn);
      return () => {
        fn();
        const idx = disposers.indexOf(fn);
        if (idx >= 0) disposers.splice(idx, 1);
      };
    };

    return {
      api,
      storage,
      ui: createModuleUIApi(moduleId),
      registerSchema(schemaDef) { return track(schemaRegistry.register(schemaDef)); },
      registerCommands(commandDefs) { return track(commandsRegistry.register(commandDefs)); },
      registerSettings(settingsDef) { return track(settingsRegistry.register(settingsDef)); },
      schemas: {
        get: (id) => schemaRegistry.get(id),
        list: (filter) => schemaRegistry.list(filter),
        resolve: (target) => schemaRegistry.resolve(target)
      },
      commands: {
        run: (id, args) => commandsRegistry.run(id, args),
        list: (filter) => commandsRegistry.list(filter)
      },
      settings: {
        listTabs: () => settingsRegistry.listTabs(),
        getKey: (key) => storage.get(key),
        setKey: (key, value) => storage.set(key, value)
      },
      backup: {
        registerProvider(provider) {
          if (!provider?.id || typeof provider.export !== 'function' || typeof provider.import !== 'function' || typeof provider.describe !== 'function') {
            throw new Error('Backup provider must include id/describe/export/import');
          }
          backupProviders.set(provider.id, provider);
          return track(() => backupProviders.delete(provider.id));
        }
      }
    };
  }

  const instance = {
    version: VERSION,
    api,
    ui: {
      listButtons: (filter) => uiRegistry.listButtons(filter),
      listPanels: (filter) => uiRegistry.listPanels(filter),
      subscribe: (handler) => uiRegistry.subscribe(handler)
    },
    schemas: {
      get: (id) => schemaRegistry.get(id),
      list: (filter) => schemaRegistry.list(filter),
      resolve: (target) => schemaRegistry.resolve(target)
    },
    commands: {
      run: (id, args) => commandsRegistry.run(id, args),
      list: (filter) => commandsRegistry.list(filter)
    },
    settings: {
      listTabs: () => settingsRegistry.listTabs(),
      getKey: (key) => storage.get(key),
      setKey: (key, value) => storage.set(key, value)
    },
    journalTemplates: {
      listTemplates: () => journalTemplates.listTemplates(),
      listTemplateEntities: () => journalTemplates.listTemplateEntities(),
      getTemplate: (id) => journalTemplates.getTemplate(id),
      addTemplate: (template) => journalTemplates.addTemplate(template),
      deleteTemplate: (id) => journalTemplates.deleteTemplate(id),
      exportDelta: (sinceRevision = 0) => journalTemplates.exportDelta(sinceRevision),
      applyDelta: (patch) => journalTemplates.applyDelta(patch)
    },
    use(module) {
      if (!module?.id || typeof module?.init !== 'function') throw new Error('Invalid module');
      if (modules.has(module.id)) return instance;
      const ctx = createModuleCtx(module.id);
      module.init(ctx);
      modules.set(module.id, module);
      emit('module:used', module.id);
      return instance;
    },
    async loadModuleFromUrl(url) {
      const mod = await import(url);
      const plugin = mod.default ?? mod.module ?? mod;
      instance.use(plugin);
      return plugin;
    },
    async start() {
      await journalTemplates.ensureInitialized();
      backupProviders.set('journal-templates', {
        id: 'journal-templates',
        version: '0.1.0',
        describe: () => ({ settings: ['templates:*'], userData: [] }),
        export: async () => ({ templates: await journalTemplates.listTemplateEntities() }),
        import: async (payload) => {
          for (const template of payload.templates ?? []) {
            await journalTemplates.deleteTemplate(template.id);
            await journalTemplates.addTemplate(template);
          }
          return { applied: true, warnings: [] };
        }
      });

// --- Custom backup provider for transfer templates ---
{
  // Dynamically import transfer core to avoid loading overhead on startup.  The transfer core
  // exposes loadTemplates()/saveTemplates() for reading and writing full template arrays.  We
  // always clone the returned array on export to avoid accidental mutation.
  const { createTransferCore } = await import('./core/transfer_core.js');
  const transferCore = createTransferCore({ storage });
  backupProviders.set('transfer-templates', {
    id: 'transfer-templates',
    version: '1.0.0',
    describe: () => ({ settings: [], userData: ['transfer:templates:v1'] }),
    export: async () => {
      const templates = await transferCore.loadTemplates();
      return { templates: Array.isArray(templates) ? [...templates] : [] };
    },
    import: async (payload, opts = {}) => {
      const newTemplates = Array.isArray(payload?.templates) ? payload.templates : [];
      const mode = opts.mode ?? 'merge';
      let existing = await transferCore.loadTemplates();
      if (!Array.isArray(existing)) existing = [];
      if (mode === 'replace') {
        existing = [];
      }
      const byId = new Map();
      for (const tpl of existing) {
        if (tpl && typeof tpl.id === 'string') {
          byId.set(tpl.id, { ...tpl });
        }
      }
      for (const tpl of newTemplates) {
        if (!tpl || typeof tpl.id !== 'string') continue;
        const prev = byId.get(tpl.id) ?? {};
        byId.set(tpl.id, { ...prev, ...tpl });
      }
      const merged = Array.from(byId.values());
      await transferCore.saveTemplates(merged);
      return { applied: true, warnings: [] };
    }
  });
}

      
// --- Backup provider for table column settings ---
{
  const TABLE_SETTINGS_KEY = '@sdo/module-table-renderer:settings';
  backupProviders.set('table-settings', {
    id: 'table-settings',
    version: '1.0.0',
    describe: async () => {
      const nav = await loadNavigationState(storage);
      const journalIds = Array.isArray(nav?.journals) ? nav.journals.map((j) => j.id) : [];
      const settingsKeys = [TABLE_SETTINGS_KEY, ...journalIds.map((id) => `${TABLE_SETTINGS_KEY}:${id}`)];
      return { settings: settingsKeys, userData: [] };
    },
    export: async () => {
      const nav = await loadNavigationState(storage);
      const journalIds = Array.isArray(nav?.journals) ? nav.journals.map((j) => j.id) : [];
      const settingsKeys = [TABLE_SETTINGS_KEY, ...journalIds.map((id) => `${TABLE_SETTINGS_KEY}:${id}`)];
      const data = {};
      for (const key of settingsKeys) {
        const val = await storage.get(key);
        if (val !== undefined) data[key] = val;
      }
      return { settings: data };
    },
    import: async (payload) => {
      const data = payload?.settings ?? {};
      for (const [key, value] of Object.entries(data)) {
        await storage.set(key, value);
      }
      return { applied: true, warnings: [] };
    }
  });
}
const nav = await loadNavigationState(storage);
      state.spaces = nav.spaces;
      state.journals = nav.journals;
      state.history = nav.history;
      
      // MIGRATION: ensure every journal has templateId (old test journals may not have it)
      try {
        const tpls = await journalTemplates.listTemplateEntities();
        const defaultTplId = (tpls.find((t) => t.id === 'test')?.id) || (tpls[0]?.id) || null;
        if (defaultTplId) {
          let changed = false;
          state.journals = (state.journals || []).map((j) => {
            if (j && !j.templateId) { changed = true; return { ...j, templateId: defaultTplId }; }
            return j;
          });
          if (changed) {
            await saveNavigationState(storage, { spaces: state.spaces, journals: state.journals, history: state.history, lastLoc: nav.lastLoc });
          }
        }
      } catch (e) {
        // ignore migration errors
      }

const loc = normalizeLocation({ spaces: state.spaces, journals: state.journals, lastLoc: nav.lastLoc });
      state.activeSpaceId = loc.activeSpaceId;
      state.activeJournalId = loc.activeJournalId;
      state.started = true;
      if (options.mount && typeof options.createUI === 'function') {
        ui = options.createUI({ sdo: instance, mount: options.mount, api });
      }
      emit('started', api.getState());
      return instance;
    },
    async destroy() {
      for (const [moduleId, disposers] of moduleDisposers.entries()) {
        for (const dispose of disposers.splice(0)) dispose();
        const module = modules.get(moduleId);
        if (typeof module?.destroy === 'function') await module.destroy();
      }
      uiRegistry.clear();
      schemaRegistry.clear();
      settingsRegistry.clear();
      commandsRegistry.clear();
      ui?.destroy?.();
      listeners.clear();
      modules.clear();
      backupProviders.clear();
      moduleDisposers.clear();
    },
    getState: api.getState,
    async commit(mutator, changedKeys = []) {
      mutator(state);
      state.history = pushHistory(state.history, {
        activeSpaceId: state.activeSpaceId,
        activeJournalId: state.activeJournalId,
        at: new Date().toISOString()
      });
      await saveNavigationState(storage, {
        spaces: state.spaces,
        journals: state.journals,
        lastLoc: { activeSpaceId: state.activeSpaceId, activeJournalId: state.activeJournalId },
        history: state.history
      });
      await bumpRevision(changedKeys);
      emit('state:changed', api.getState());
    },
    async exportNavigationState() {
      return toNavPayload(await loadNavigationState(storage));
    },
    async importNavigationState(payload) {
      const nav = fromNavPayload(payload);
      await saveNavigationState(storage, nav);
      const loc = normalizeLocation({ spaces: nav.spaces, journals: nav.journals, lastLoc: nav.lastLoc });
      state.spaces = nav.spaces;
      state.journals = nav.journals;
      state.history = nav.history;
      state.activeSpaceId = loc.activeSpaceId;
      state.activeJournalId = loc.activeJournalId;
      return { applied: true, warnings: [] };
    },
    async exportBackup(opts = {}) {
      const scope = opts.scope ?? 'all';
      const backupId = crypto.randomUUID();
      const bundle = {
        format: BACKUP_FORMAT,
        formatVersion: 1,
        backupId,
        createdAt: new Date().toISOString(),
        app: { name: '@sdo/core', version: VERSION },
        scope,
        core: { navigation: null, settings: { coreSettings: (await storage.get(NAV_KEYS.coreSettings)) ?? {} } },
        modules: {},
        userData: {}
      };
      if (opts.includeNavigation !== false && (scope === 'all' || scope === 'userData' || scope === 'modules')) {
        bundle.core.navigation = await instance.exportNavigationState();
      }
      const moduleIds = opts.modules ?? [...backupProviders.keys()];
      for (const id of moduleIds) {
        const provider = backupProviders.get(id);
        if (!provider) continue;
        bundle.modules[id] = { moduleVersion: provider.version, data: await provider.export({ includeUserData: opts.includeUserData !== false, scope }) };
      }
      bundle.integrity = await createIntegrity(bundle);
      return opts.encrypt?.enabled ? encryptBackup(bundle, opts.encrypt.password) : bundle;
    },
    async importBackup(input, opts = {}) {
      const bundle = input?.format === ENCRYPTED_BACKUP_FORMAT ? await decryptBackup(input, opts.decrypt?.password ?? '') : input;
      if (bundle?.format !== BACKUP_FORMAT) throw new Error('Unsupported backup format');
      if (!await verifyIntegrity(bundle)) throw new Error('Backup integrity check failed');

      const report = { core: { applied: false, warnings: [] }, navigation: { applied: false, warnings: [] }, modules: {} };
      if (bundle.core?.settings) {
        await storage.set(NAV_KEYS.coreSettings, bundle.core.settings.coreSettings ?? {});
        report.core.applied = true;
      }
      if (bundle.core?.navigation) {
        report.navigation = await instance.importNavigationState(bundle.core.navigation);
      }
      for (const [id, payload] of Object.entries(bundle.modules ?? {})) {
        const provider = backupProviders.get(id);
        if (!provider?.import) {
          report.modules[id] = { applied: false, warnings: ['provider not found'] };
          continue;
        }
        report.modules[id] = await provider.import(payload.data, { mode: opts.mode ?? 'merge', includeUserData: opts.includeUserData !== false });
      }
      return report;
    },
    async exportDelta({ baseId, baseHashB64, sinceRevision = 0 } = {}) {
      const log = (await storage.get(NAV_KEYS.revisionLog)) ?? [];
      const changes = log.filter((item) => item.rev > sinceRevision);
      return {
        format: DELTA_BACKUP_FORMAT,
        formatVersion: 1,
        base: { baseId, baseHashB64 },
        createdAt: new Date().toISOString(),
        revision: state.revision,
        changes: { core: { set: { revision: state.revision }, del: [] }, navigation: changes, modules: {} }
      };
    },
    async applyDelta(baseBundle, deltaBundle) {
      if (baseBundle.backupId !== deltaBundle.base.baseId) throw new Error('Delta baseId mismatch');
      return { ...baseBundle, deltaAppliedAt: new Date().toISOString(), delta: deltaBundle };
    },
    on(event, handler) {
      const arr = listeners.get(event) ?? [];
      arr.push(handler);
      listeners.set(event, arr);
      return () => instance.off(event, handler);
    },
    off(event, handler) {
      const arr = listeners.get(event) ?? [];
      listeners.set(event, arr.filter((h) => h !== handler));
    }
  };

  if (Array.isArray(options.modules)) {
    for (const module of options.modules) instance.use(module);
  }
  return instance;
}

export function createNavi(storage) {
  assertStorage(storage);
  return {
    async exportNavigationState() {
      return toNavPayload(await loadNavigationState(storage));
    },
    async importNavigationState(payload) {
      return saveNavigationState(storage, fromNavPayload(payload));
    }
  };
}

export { encryptBackup, decryptBackup, signBackup, verifyBackup, verifyIntegrity } from './backup/crypto.js';
export { VERSION as version };

export { createTableEngine, createTableEngineModule } from './modules/table_engine.js';

export { createTableStoreModule } from './modules/table_store.js';
export { createTableFormatterModule, formatCell, parseInput } from './modules/table_formatter.js';
export { createTableRendererModule, getRenderableCells } from './modules/table_renderer.js';
export { createJournalStore } from './stores/journal_store.js';
export { createJournalTemplatesContainer } from './stores/journal_templates_container.js';
export { createTableSubrowsBridge } from './modules/table_subrows_bridge.js';
