/**
 * Draft/apply/discard state for settings modal sections.
 */
(function attachSettingsState(global) {
  const UI = (global.UI = global.UI || {});
  UI.settings = UI.settings || {};

  function createSettingsState() {
    const drafts = new Map();

    function setDraft(sectionId, patch) {
      const prev = drafts.get(sectionId) || {};
      drafts.set(sectionId, { ...prev, ...patch });
      return drafts.get(sectionId);
    }

    function getDraft(sectionId) {
      return { ...(drafts.get(sectionId) || {}) };
    }

    function hasDraft(sectionId) {
      return drafts.has(sectionId);
    }

    function discard(sectionId) {
      drafts.delete(sectionId);
    }

    function apply(sectionId, applyFn) {
      const payload = getDraft(sectionId);
      const result = applyFn?.(payload);
      drafts.delete(sectionId);
      return result;
    }

    function clear() {
      drafts.clear();
    }

    return {
      setDraft,
      getDraft,
      hasDraft,
      discard,
      apply,
      clear
    };
  }

  UI.settings.createState = createSettingsState;
})(typeof window !== 'undefined' ? window : globalThis);
