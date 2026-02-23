/**
 * UI/UX Manager: centralized settings state + events + settings panel renderer.
 *
 * Public API:
 * - UI.init()
 * - UI.applySettings(partialSettings)
 * - UI.getSettings()
 * - UI.on(eventName, cb)
 * - UI.emit(eventName, payload)
 *
 * Host integration:
 * - Optional external storage adapter can be provided as `UI.storage`
 *   and must implement `getItem(key)` / `setItem(key, value)`.
 */
(function attachUIManager(global) {
  const UI = (global.UI = global.UI || {});
  const STORAGE_KEY = 'ui.settings';

  const DEFAULTS = Object.freeze({
    theme: 'light',
    scale: 1,
    touchMode: false,
    navCircles: true,
    gestures: true,
    tableDensity: 'normal'
  });

  const listeners = new Map();
  let state = { ...DEFAULTS };

  function getStorage() {
    return UI.storage || global.localStorage;
  }

  function emit(eventName, payload) {
    (listeners.get(eventName) || []).forEach((cb) => cb(payload));
  }

  function on(eventName, cb) {
    const list = listeners.get(eventName) || [];
    list.push(cb);
    listeners.set(eventName, list);
    return () => listeners.set(eventName, (listeners.get(eventName) || []).filter((item) => item !== cb));
  }

  function readStorage() {
    const storage = getStorage();
    try {
      const raw = storage.getItem(STORAGE_KEY);
      return raw ? { ...DEFAULTS, ...JSON.parse(raw) } : { ...DEFAULTS };
    } catch {
      return { ...DEFAULTS };
    }
  }

  function writeStorage(nextState) {
    getStorage().setItem(STORAGE_KEY, JSON.stringify(nextState));
  }

  function normalize(next) {
    const normalized = { ...next };

    normalized.scale = Number(normalized.scale);
    if (!Number.isFinite(normalized.scale)) normalized.scale = DEFAULTS.scale;
    normalized.scale = Math.min(1.4, Math.max(1, normalized.scale));

    normalized.touchMode = Boolean(normalized.touchMode);
    normalized.navCircles = Boolean(normalized.navCircles);
    normalized.gestures = Boolean(normalized.gestures);

    if (!['compact', 'normal'].includes(normalized.tableDensity)) {
      normalized.tableDensity = DEFAULTS.tableDensity;
    }

    return normalized;
  }

  function applyRootState() {
    const root = document.documentElement;
    root.style.setProperty('--ui-scale', String(state.scale));
    root.dataset.touchMode = state.touchMode ? 'on' : 'off';
    root.dataset.navCircles = state.navCircles ? 'on' : 'off';
    root.dataset.gestures = state.gestures ? 'on' : 'off';
    root.dataset.tableDensity = state.tableDensity;
  }

  function applySettings(partialSettings = {}) {
    const prev = { ...state };
    const merged = normalize({ ...state, ...partialSettings });

    if ('theme' in partialSettings && global.UITheme?.applyTheme) {
      merged.theme = global.UITheme.applyTheme(partialSettings.theme);
    }

    state = merged;
    applyRootState();
    writeStorage(state);

    // themeChanged is emitted by UITheme.applyTheme() to avoid duplicates.
    if (prev.scale !== state.scale) emit('scaleChanged', state.scale);
    emit('settingsChanged', { ...state });
    return { ...state };
  }

  function getSettings() {
    return { ...state };
  }



  function init() {
    state = normalize(readStorage());

    if (global.UITheme?.initTheme) {
      state.theme = global.UITheme.initTheme();
    }

    applyRootState();
    writeStorage(state);
    emit('settingsChanged', { ...state });
    return { ...state };
  }

  UI.init = init;
  UI.on = on;
  UI.emit = emit;
  UI.applySettings = applySettings;
  UI.getSettings = getSettings;
})(typeof window !== 'undefined' ? window : globalThis);
