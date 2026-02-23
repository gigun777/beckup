/**
 * Schema-based form builder.
 *
 * Supported field types: text, number, date, textarea, select, checkbox
 */
(function attachForm(global) {
  const UI = (global.UI = global.UI || {});

  function createFieldInput(field) {
    if (field.type === 'textarea') return document.createElement('textarea');
    if (field.type === 'select') return document.createElement('select');
    const input = document.createElement('input');
    input.type = field.type || 'text';
    return input;
  }

  function applyCommonConstraints(input, field) {
    if (field.placeholder) input.placeholder = field.placeholder;
    if (field.required) input.required = true;
    if (field.pattern) input.pattern = field.pattern;
    if (field.min !== undefined) input.min = String(field.min);
    if (field.max !== undefined) input.max = String(field.max);
    if (field.step !== undefined) input.step = String(field.step);
    if (field.mask) input.dataset.mask = field.mask;
  }

  function setInitialValue(input, field, value) {
    if (field.type === 'checkbox') input.checked = Boolean(value);
    else if (field.type !== 'select' && value !== undefined && value !== null) input.value = String(value);
  }

  function setupSelectOptions(select, field, value) {
    (field.options || []).forEach((opt) => {
      const option = document.createElement('option');
      option.value = String(opt.value);
      option.textContent = opt.label;
      if (String(value ?? '') === String(opt.value)) option.selected = true;
      select.appendChild(option);
    });
  }

  function validateByType(element) {
    if (!element.name) return true;

    let message = '';
    const value = element.type === 'checkbox' ? String(element.checked) : String(element.value || '');

    if (element.dataset.mask && value) {
      const maskRegex = new RegExp(element.dataset.mask);
      if (!maskRegex.test(value)) message = 'Невірний формат поля.';
    }

    if (!message && element.type === 'number' && value) {
      if (Number.isNaN(Number(value))) message = 'Введіть коректне число.';
    }

    if (!message && element.type === 'date' && value) {
      if (Number.isNaN(new Date(value).getTime())) message = 'Введіть коректну дату.';
    }

    element.setCustomValidity(message);
    return !message;
  }

  function validateForm(form) {
    const fields = [...form.querySelectorAll('input,select,textarea')];
    fields.forEach((el) => validateByType(el));
    return form.reportValidity();
  }

  function getInteractiveNodes(form) {
    return [...form.querySelectorAll('input,select,textarea,button[type="submit"]')]
      .filter((el) => !el.disabled && el.type !== 'hidden');
  }

  function bindKeyboardNavigation(form) {
    form.addEventListener('keydown', (event) => {
      const nodes = getInteractiveNodes(form);
      const index = nodes.indexOf(document.activeElement);
      if (index < 0) return;

      if (event.key === 'Enter' && event.ctrlKey) {
        event.preventDefault();
        form.requestSubmit();
        return;
      }

      if (event.key === 'Enter' && !event.altKey) {
        event.preventDefault();
        const next = event.shiftKey ? nodes[index - 1] : nodes[index + 1];
        if (next) next.focus();
        return;
      }

      if (event.key === 'Escape') {
        const modal = form.closest('.ui-modal');
        if (modal?.dataset.modalId) UI.modal?.close?.(modal.dataset.modalId);
      }
    });
  }

  function build(container, schema = [], opts = {}) {
    container.innerHTML = '';
    const form = document.createElement('form');
    form.className = 'ui-form';

    schema.forEach((field) => {
      const row = document.createElement('label');
      row.className = 'ui-form-row';
      row.dataset.name = field.name;

      const label = document.createElement('span');
      label.className = 'ui-form-label';
      label.textContent = field.label || field.name;

      const input = createFieldInput(field);
      input.name = field.name;
      input.className = 'input';

      const value = opts.values?.[field.name] ?? field.defaultValue;
      applyCommonConstraints(input, field);
      setInitialValue(input, field, value);
      if (field.type === 'select') setupSelectOptions(input, field, value);

      input.addEventListener('input', () => validateByType(input));
      input.addEventListener('change', () => validateByType(input));

      row.append(label, input);
      form.appendChild(row);
    });

    const actions = document.createElement('div');
    actions.className = 'ui-form-actions';

    const submit = document.createElement('button');
    submit.type = 'submit';
    submit.className = 'btn btn-primary';
    submit.textContent = opts.submitText || 'Зберегти';
    actions.appendChild(submit);

    if (opts.quickAdd) {
      const quick = document.createElement('label');
      quick.className = 'ui-form-quick';
      quick.innerHTML = '<input type="checkbox" name="__quickAdd" /> Швидке додавання';
      actions.appendChild(quick);
    }

    form.appendChild(actions);
    bindKeyboardNavigation(form);

    form.addEventListener('submit', (event) => {
      event.preventDefault();
      if (!validateForm(form)) return;

      const values = getValues(form);
      opts.onSubmit?.(values, { quickAdd: values.__quickAdd === true });
      if (values.__quickAdd) form.reset();
    });

    container.appendChild(form);
    return form;
  }

  function getValues(container) {
    const form = container.matches?.('form') ? container : container.querySelector?.('form');
    if (!form) return {};

    const payload = {};
    [...form.elements].forEach((el) => {
      if (!el.name) return;

      if (el.type === 'checkbox') payload[el.name] = el.checked;
      else if (el.type === 'number') payload[el.name] = el.value === '' ? null : Number(el.value);
      else payload[el.name] = el.value;
    });
    return payload;
  }

  UI.form = { build, getValues };
})(typeof window !== 'undefined' ? window : globalThis);
