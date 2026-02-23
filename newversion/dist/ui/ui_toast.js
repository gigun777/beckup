/**
 * Toast notifications.
 */
(function attachToast(global) {
  const UI = (global.UI = global.UI || {});
  let host = null;

  function ensureHost() {
    if (host) return host;
    host = document.createElement('div');
    host.className = 'ui-toast-host';
    document.body.appendChild(host);
    return host;
  }

  function show(text, opts = {}) {
    const node = document.createElement('div');
    node.className = `ui-toast type-${opts.type || 'info'}`;
    node.textContent = text;

    ensureHost().appendChild(node);
    global.setTimeout(() => node.remove(), opts.timeout ?? 3200);
    return node;
  }

  function undo(text, opts = {}) {
    const node = document.createElement('div');
    node.className = 'ui-toast type-warn';

    const label = document.createElement('span');
    label.textContent = text;

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn btn-link';
    btn.textContent = opts.undoText || 'Undo';
    btn.addEventListener('click', () => {
      opts.onUndo?.();
      node.remove();
    });

    node.append(label, btn);
    ensureHost().appendChild(node);
    global.setTimeout(() => node.remove(), opts.timeout ?? 5000);
    return node;
  }

  UI.toast = { show, undo };
})(typeof window !== 'undefined' ? window : globalThis);
