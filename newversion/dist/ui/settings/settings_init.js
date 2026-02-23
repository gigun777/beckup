/**
 * Settings integration initializer.
 * Ensures default features are registered once.
 */
(function attachSettingsInit(global) {
  const UI = (global.UI = global.UI || {});
  UI.settings = UI.settings || {};

  let initialized = false;

  function initSettingsRegistry() {
    if (initialized) return;

    if (!UI.settings.registry || typeof UI.settings.registry.registerFeature !== 'function') {
      throw new Error('UI.settings.init: settings_registry.js must be loaded before settings_init.js');
    }

    if (!UI.settings.createState || typeof UI.settings.createState !== 'function') {
      throw new Error('UI.settings.init: settings_state.js must be loaded before settings_init.js');
    }

    initialized = true;

    if (typeof UI.settings.registerTableSettingsFeature === 'function') {
      UI.settings.registerTableSettingsFeature();
    }
    if (typeof UI.settings.registerUxUiSettingsFeature === 'function') {
      UI.settings.registerUxUiSettingsFeature();
    }
    if (typeof UI.settings.registerBackupSettingsFeature === 'function') {
      UI.settings.registerBackupSettingsFeature();
    }
  }

  UI.settings.init = initSettingsRegistry;
})(typeof window !== 'undefined' ? window : globalThis);
