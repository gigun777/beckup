/**
 * ESM bootstrap wrapper (Option 2).
 *
 * Goal:
 * - Provide a single, deterministic startup entrypoint without requiring the host to call globals.
 * - Internally, legacy UI modules may still attach to globalThis.UI via side-effects. This wrapper
 *   normalizes that into an explicit exported `bootstrap()` function.
 *
 * Usage:
 *   import { bootstrap } from "./dist/ui/ui_bootstrap_esm.js";
 *   await bootstrap({ storage: localStorage, settingsHost: document.getElementById("...") });
 */

import { initTheme, applyTheme, toggleTheme, getTheme } from "./theme.js";

// Transfer UI (visual modals) + bridge (UI -> TransferCore)
import "./transfer_modals.js";
import { attachTransferUI } from "./ui_transfer_bridge.js";

// Side-effect imports: attach UI engines to globalThis.UI (legacy IIFE modules in this repo).
import "./ui_manager.js";
import "./ui_modal.js";
import "./ui_form.js";
import "./ui_toast.js";
import "./ui_backup.js";

// Settings subsystem (also legacy IIFE modules).

/**
 * @typedef {{getItem:(k:string)=>string|null, setItem:(k:string,v:string)=>void}} StorageAdapter
 */

/**
 * @param {{storage?: StorageAdapter, settingsHost?: HTMLElement|null, preferSystemTheme?: boolean}} options
 */
export async function bootstrap(options = {}) {
  // Init SettingsWindow v2 (global) once
  try {
    if (window.SettingsWindow && !window.__sws_inited) {
      window.SettingsWindow.init({ theme: 'light' });
      window.__sws_inited = true;
    }
  } catch (e) { console.warn('SettingsWindow.init failed:', e); }

  const global = globalThis;

  // Ensure UI namespace exists for legacy modules
  const UI = (global.UI = global.UI || {});

  const storage = normalizeStorage(options.storage);
  // Expose storage adapter for legacy IIFE modules (ui_manager/settings/etc.)
  UI.storage = storage;

  // Lazy-attach Transfer UI bridge once SEDO api is available.
  const tryAttachTransfer = () => {
    const api = global.sdo?.api || global.SDO?.api || global.__sdo_api || global.__sdo?.api;
    if (!api) return false;
    if (!api.tableStore) return false;
    try {
      attachTransferUI({ UI, api, storage: options.storage || UI.storage || global.localStorage });
      return true;
    } catch (e) {
      console.warn('TransferUI attach failed', e);
      return false;
    }
  };
  if (!tryAttachTransfer()) {
    let attempts = 0;
    const t = setInterval(() => {
      attempts++;
      if (tryAttachTransfer() || attempts > 50) clearInterval(t);
    }, 100);
  }

  // 2) Theme init (ESM theme runtime)
  // Note: current theme.js persists to localStorage by design. We still initialize early here.
  initTheme({ preferSystemTheme: options.preferSystemTheme ?? true });

  // 3) UI state manager init (apply scale/density/etc.)
  if (typeof UI.init === "function") {
    UI.init();
  }

  // 4) Optional: mount UX|UI tab right away if host provided
  return {
    theme: { initTheme, applyTheme, toggleTheme, getTheme },
    settings: {
      getSettings: UI.getSettings ? UI.getSettings.bind(UI) : null,
      applySettings: UI.applySettings ? UI.applySettings.bind(UI) : null,
      on: UI.on ? UI.on.bind(UI) : null
    }};
}

function normalizeStorage(adapter) {
  // Accept browser Storage, or a simple {getItem,setItem} adapter.
  const ls = globalThis.localStorage;
  if (!adapter) return ls;

  // If it's a Web Storage object, it already implements getItem/setItem
  if (typeof adapter.getItem === "function" && typeof adapter.setItem === "function") return adapter;

  // Fallback to localStorage
  return ls;
}
