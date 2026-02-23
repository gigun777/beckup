/**
 * Theme runtime API (ESM + backward-compatible globals).
 *
 * Canonical contract:
 * - Storage key: 'ui.theme'
 * - Apply theme via: document.documentElement.dataset.theme = 'light'|'dark'
 * - Persist via UI.storage (if present) else localStorage fallback
 *
 * Public API (ESM exports):
 * - initTheme()
 * - applyTheme(themeName)
 * - toggleTheme()
 * - getTheme()
 */
const STORAGE_KEY = 'ui.theme';
const ROOT = document.documentElement;
const FALLBACK_THEME = 'light';

function safeLocalStorageGet(key) {
  try { return window.localStorage ? window.localStorage.getItem(key) : null; } catch { return null; }
}
function safeLocalStorageSet(key, value) {
  try { if (window.localStorage) window.localStorage.setItem(key, value); } catch { /* ignore */ }
}

/**
 * Returns storage adapter used by theme layer.
 * Adapter contract: { getItem(key), setItem(key, value) }
 */
function getStorage() {
  const ui = globalThis.UI;
  if (ui && ui.storage && typeof ui.storage.getItem === 'function' && typeof ui.storage.setItem === 'function') {
    return ui.storage;
  }
  return {
    getItem: (k) => safeLocalStorageGet(k),
    setItem: (k, v) => safeLocalStorageSet(k, v),
  };
}

function normalizeTheme(theme) {
  return (theme === 'dark' || theme === 'light') ? theme : FALLBACK_THEME;
}

export function getTheme() {
  const stored = getStorage().getItem(STORAGE_KEY);
  if (stored === 'dark' || stored === 'light') return stored;
  // fall back to current dataset if valid
  const cur = ROOT.dataset.theme;
  if (cur === 'dark' || cur === 'light') return cur;
  return null;
}

export function applyTheme(themeName) {
  const t = normalizeTheme(themeName);
  ROOT.dataset.theme = t; // core requirement: data-theme on root
  getStorage().setItem(STORAGE_KEY, t);
  return t;
}

export function initTheme() {
  const stored = getTheme();
  if (stored) return applyTheme(stored);

  // prefers-color-scheme fallback
  let preferred = null;
  try {
    preferred = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  } catch {
    preferred = null;
  }
  return applyTheme(preferred || FALLBACK_THEME);
}

export function toggleTheme() {
  const cur = getTheme() || FALLBACK_THEME;
  return applyTheme(cur === 'dark' ? 'light' : 'dark');
}

// Backward-compatible globals for legacy integration
(function attachGlobals(global) {
  global.UITheme = global.UITheme || {};
  global.UITheme.initTheme = initTheme;
  global.UITheme.applyTheme = applyTheme;
  global.UITheme.toggleTheme = toggleTheme;
  global.UITheme.getTheme = getTheme;

  // Also expose simple global functions (used by older app.js)
  global.initTheme = initTheme;
  global.applyTheme = applyTheme;
  global.toggleTheme = toggleTheme;
  global.getTheme = getTheme;
})(globalThis);
