// Centralized modal manager (Step 2)
// - One place to control overlay, scroll-lock, ESC close, and consistent styling.

export function createModalManager(layerEl) {
  if (!layerEl) throw new Error('createModalManager: layerEl is required');

  layerEl.classList.add('sdo-modal-layer');
  layerEl.hidden = true;

  let isOpen = false;
  let lastActive = null;
  let cleanupFn = null;

  function lockScroll() {
    const html = document.documentElement;
    html.dataset.sdoScrollLock = '1';
    html.style.overflow = 'hidden';
    document.body.style.overflow = 'hidden';
  }

  function unlockScroll() {
    const html = document.documentElement;
    delete html.dataset.sdoScrollLock;
    html.style.overflow = '';
    document.body.style.overflow = '';
  }

  function close() {
    if (!isOpen) return;
    isOpen = false;
    layerEl.hidden = true;
    layerEl.innerHTML = '';
    unlockScroll();
    cleanupFn?.();
    cleanupFn = null;
    if (lastActive && typeof lastActive.focus === 'function') {
      try { lastActive.focus(); } catch (_) {}
    }
    lastActive = null;
  }

  function open(contentEl, opts = {}) {
    const { closeOnOverlay = true } = opts;

    close();
    lastActive = document.activeElement;
    isOpen = true;

    const overlay = document.createElement('div');
    overlay.className = 'sdo-ui-modal-overlay';

    const win = document.createElement('div');
    win.className = 'sdo-ui-modal-window';

    if (contentEl) win.append(contentEl);
    overlay.append(win);
    layerEl.append(overlay);
    layerEl.hidden = false;
    lockScroll();

    const onKey = (e) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        close();
      }
    };
    document.addEventListener('keydown', onKey);

    const onOverlayClick = (e) => {
      if (!closeOnOverlay) return;
      if (e.target === overlay) close();
    };
    overlay.addEventListener('mousedown', onOverlayClick);

    // Focus first focusable element
    queueMicrotask(() => {
      const focusable = win.querySelector('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])');
      if (focusable && typeof focusable.focus === 'function') focusable.focus();
    });

    cleanupFn = () => {
      document.removeEventListener('keydown', onKey);
      overlay.removeEventListener('mousedown', onOverlayClick);
    };

    return { overlay, win, close };
  }

  return { open, close, get isOpen() { return isOpen; } };
}
