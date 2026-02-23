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

// Side-effect imports: attach UI engines to globalThis.UI (legacy IIFE modules in this repo).
import "./ui_manager.js";
import "./ui_modal.js";
import "./ui_form.js";
import "./ui_toast.js";
import "./ui_backup.js";

// Settings subsystem (also legacy IIFE modules).
import "./settings/settings_registry.js";
import "./settings/settings_state.js";
// Legacy settings shell modal removed (SWS v2 is the only settings UI)
import "./settings/settings_init.js";

/**
 * @typedef {{getItem:(k:string)=>string|null, setItem:(k:string,v:string)=>void}} StorageAdapter
 */

/**
 * @param {{storage?: StorageAdapter, settingsHost?: HTMLElement|null, preferSystemTheme?: boolean}} options
 */
export async function bootstrap(options = {}) {
  const global = globalThis;

  // Ensure UI namespace exists for legacy modules
  const UI = (global.UI = global.UI || {});
  const storage = normalizeStorage(options.storage);

  // Expose storage adapter for legacy IIFE modules (ui_manager/settings/etc.)
  UI.storage = storage;

  // 1) Settings init (register default features once)
  if (UI.settings && typeof UI.settings.init === "function") {
    UI.settings.init();
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
