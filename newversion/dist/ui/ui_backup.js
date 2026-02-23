/**
 * UX|UI backup module.
 *
 * Public API:
 * - UI.backup.export(sectionName, options)
 * - UI.backup.import(json)
 * - UI.backup.getManifest()
 */
(function attachBackup(global) {
  const UI = (global.UI = global.UI || {});

  const manifest = Object.freeze({
    uxui: ['theme', 'scale', 'touchMode', 'navCircles', 'gestures', 'tableDensity']
  });

  function pickSettings(keys, source) {
    const selected = {};
    keys.forEach((key) => {
      if (key in source) selected[key] = source[key];
    });
    return selected;
  }

  function exportBackup(sectionName = 'uxui', options = {}) {
    const settings = UI.getSettings?.() || {};
    const sections = {};

    if (sectionName === 'all') {
      Object.keys(manifest).forEach((name) => {
        const keys = options.keys?.[name]?.length ? options.keys[name] : manifest[name];
        sections[name] = pickSettings(keys, settings);
      });
    } else {
      const keys = options.keys?.length ? options.keys : manifest[sectionName] || [];
      sections[sectionName] = pickSettings(keys, settings);
    }

    const payload = {
      version: options.version || '12.x',
      createdAt: new Date().toISOString(),
      sections
    };

    const json = JSON.stringify(payload, null, 2);
    return options.asBlob ? new Blob([json], { type: 'application/json' }) : json;
  }

  function importBackup(input) {
    let parsed;
    try {
      parsed = typeof input === 'string' ? JSON.parse(input) : input;
    } catch {
      return false;
    }

    if (!parsed?.sections || typeof parsed.sections !== 'object') return false;

    let applied = false;
    Object.keys(manifest).forEach((section) => {
      if (parsed.sections[section] && typeof parsed.sections[section] === 'object') {
        UI.applySettings?.(pickSettings(manifest[section], parsed.sections[section]));
        applied = true;
      }
    });

    return applied;
  }

  function getManifest() {
    return JSON.parse(JSON.stringify(manifest));
  }

  UI.backup = {
    export: exportBackup,
    import: importBackup,
    getManifest
  };
})(typeof window !== 'undefined' ? window : globalThis);
