/**
 * Global settings registry.
 * Stores settings features and exposes deterministic ordered lookup.
 */
(function attachSettingsRegistry(global) {
  const UI = (global.UI = global.UI || {});
  UI.settings = UI.settings || {};

  const featureMap = new Map();

  function normalizeFeature(feature) {
    if (!feature || typeof feature !== 'object') {
      throw new Error('Settings registry: feature must be an object');
    }
    if (!feature.id || !feature.title) {
      throw new Error('Settings registry: feature must contain id and title');
    }
    const sections = Array.isArray(feature.sections) ? feature.sections : [];
    return {
      ...feature,
      sections: sections.map((section, index) => ({
        order: index,
        ...section
      }))
    };
  }

  function registerFeature(feature) {
    const normalized = normalizeFeature(feature);
    featureMap.set(normalized.id, normalized);
    return normalized;
  }

  function getFeature(featureId) {
    return featureMap.get(featureId) || null;
  }

  function listFeatures() {
    return [...featureMap.values()].sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  }

  function listSections(featureId) {
    const feature = getFeature(featureId);
    if (!feature) return [];
    return [...feature.sections].sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  }

  function clearFeatures() {
    featureMap.clear();
  }

  UI.settings.registry = {
    registerFeature,
    getFeature,
    listFeatures,
    listSections,
    clearFeatures
  };
})(typeof window !== 'undefined' ? window : globalThis);
