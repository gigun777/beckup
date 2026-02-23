/**
 * UI bootstrap helpers for host application startup.
 *
 * Patch 1 goal:
 * - initialize theme as early as possible
 * - initialize UI settings/state once during app bootstrap
 */
(function attachUIBootstrap(global) {
  const UI = (global.UI = global.UI || {});

  /**
   * Normalizes storage adapter shape for UI modules.
   * Adapter contract: { getItem(key), setItem(key, value) }
   */
  function applyStorageAdapter(storageAdapter) {
    if (!storageAdapter) return;
    if (typeof storageAdapter.getItem !== 'function' || typeof storageAdapter.setItem !== 'function') {
      throw new Error('UI bootstrap: storage adapter must implement getItem/setItem');
    }
    UI.storage = storageAdapter;
  }

  /**
   * Bootstraps theme + UI manager in the correct order.
   *
   * @param {{storage?: {getItem: Function, setItem: Function}, settingsHost?: string|Element}} options
   */
  function bootstrap(options = {}) {
    applyStorageAdapter(options.storage);

    UI.settings?.init?.();

    const appliedTheme = global.UITheme?.initTheme ? global.UITheme.initTheme() : global.initTheme?.();
    const settings = UI.init?.() || null;
    return {
      theme: appliedTheme || global.UITheme?.getTheme?.() || global.getTheme?.() || null,
      settings};
  }

  UI.bootstrap = bootstrap;
  global.UIBootstrap = Object.freeze({ bootstrap });
})(typeof window !== 'undefined' ? window : globalThis);
