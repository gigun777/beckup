/**
 * Unified controls factory.
 */
(function attachControls(global) {
  const UI = (global.UI = global.UI || {});

  function assignBaseProps(node, opts = {}) {
    if (opts.id) node.id = opts.id;
    if (opts.title) node.title = opts.title;
    if (opts.className) node.className = `${node.className} ${opts.className}`.trim();
    if (opts.disabled) node.disabled = true;
    if (typeof opts.onClick === 'function') node.addEventListener('click', opts.onClick);
    return node;
  }

  function button(opts = {}) {
    const node = document.createElement('button');
    node.type = opts.type || 'button';
    node.className = `btn ${opts.variant ? `btn-${opts.variant}` : ''}`.trim();
    node.textContent = opts.text || 'Button';
    return assignBaseProps(node, opts);
  }

  function iconButton(opts = {}) {
    const node = button({ ...opts, text: '' });
    node.classList.add('btn-icon');
    node.setAttribute('aria-label', opts.ariaLabel || opts.title || 'icon button');
    node.innerHTML = opts.iconHtml || '<span aria-hidden="true">◎</span>';
    return node;
  }

  function toggle(opts = {}) {
    const wrapper = document.createElement('label');
    wrapper.className = `toggle ${opts.className || ''}`.trim();

    const input = document.createElement('input');
    input.type = 'checkbox';
    input.checked = Boolean(opts.checked);

    const slider = document.createElement('span');
    slider.className = 'toggle-slider';

    const text = document.createElement('span');
    text.className = 'toggle-label';
    text.textContent = opts.label || '';

    input.addEventListener('change', () => opts.onChange?.(input.checked));
    wrapper.append(input, slider, text);
    return wrapper;
  }

  function select(opts = {}) {
    const node = document.createElement('select');
    node.className = `input ${opts.className || ''}`.trim();

    (opts.options || []).forEach((item) => {
      const option = document.createElement('option');
      option.value = String(item.value);
      option.textContent = item.label;
      if (String(opts.value) === String(item.value)) option.selected = true;
      node.appendChild(option);
    });

    node.addEventListener('change', () => opts.onChange?.(node.value));
    return assignBaseProps(node, opts);
  }

  UI.controls = { button, iconButton, toggle, select };
})(typeof window !== 'undefined' ? window : globalThis);
