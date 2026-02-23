function h(tag, props = {}, children = []) {
  const node = document.createElement(tag);
  for (const [key, value] of Object.entries(props)) {
    if (key === 'className') node.className = value;
    else if (key.startsWith('on') && typeof value === 'function') node.addEventListener(key.slice(2).toLowerCase(), value);
    else node[key] = value;
  }
  for (const child of children) node.append(child);
  return node;
}

function ruleDraft() {
  return {
    id: `rule-${Date.now()}`,
    name: 'Нове правило',
    sources: [{ value: '' }],
    op: 'concat',
    params: { separator: ' ', trim: true, skipEmpty: true },
    targets: [],
    write: { mode: 'replace' }
  };
}

export function openTemplateEditorModal({ template, onSave, journals }) {
  const draft = structuredClone(template);
  const host = h('div');

  const titleInput = h('input', { value: draft.title, placeholder: 'Назва шаблону' });
  const rulesHost = h('div');

  const renderRules = () => {
    rulesHost.innerHTML = '';
    draft.rules.forEach((rule, index) => {
      const row = h('div', { className: 'sdo-settings-row' });
      const nameInput = h('input', {
        value: rule.name,
        oninput: (e) => {
          draft.rules[index].name = e.target.value;
        }
      });

      const opSelect = h('select', {
        onchange: (e) => {
          draft.rules[index].op = e.target.value;
        }
      }, [h('option', { value: 'direct', textContent: 'direct' }), h('option', { value: 'concat', textContent: 'concat' }), h('option', { value: 'math', textContent: 'math' })]);
      opSelect.value = rule.op;

      const separatorInput = h('input', {
        value: rule.params?.separator ?? '',
        placeholder: 'separator (напр. /, пробіл, \\n)',
        oninput: (e) => {
          draft.rules[index].params = { ...(draft.rules[index].params ?? {}), separator: e.target.value };
        }
      });

      const sourceJournal = journals[0]?.id;
      const sourceField = journals[0]?.fields?.[0]?.id ?? '';
      const targetJournal = journals[1]?.id ?? journals[0]?.id ?? '';
      const targetField = journals[1]?.fields?.[0]?.id ?? journals[0]?.fields?.[0]?.id ?? '';

      const setSampleMapBtn = h('button', {
        textContent: 'Заповнити приклад мапінгу',
        onclick: () => {
          draft.rules[index].sources = [
            { cell: { journalId: sourceJournal, recordId: '__CURRENT__', fieldId: sourceField } }
          ];
          draft.rules[index].targets = [
            { cell: { journalId: targetJournal, recordId: '__TARGET__', fieldId: targetField } }
          ];
        }
      });

      const removeBtn = h('button', {
        textContent: 'Видалити правило',
        onclick: () => {
          draft.rules.splice(index, 1);
          renderRules();
        }
      });

      row.append(h('label', { textContent: `Правило #${index + 1}` }), nameInput, opSelect, separatorInput, setSampleMapBtn, removeBtn);
      rulesHost.append(row);
    });
  };

  const addRuleBtn = h('button', {
    textContent: 'Додати правило',
    onclick: () => {
      draft.rules.push(ruleDraft());
      renderRules();
    }
  });

  const saveBtn = h('button', {
    textContent: 'Зберегти',
    onclick: () => {
      draft.title = titleInput.value || draft.title;
      onSave(draft);
      window.UI.modal.close(modalId);
    }
  });

  const cancelBtn = h('button', {
    textContent: 'Скасувати',
    onclick: () => window.UI.modal.close(modalId)
  });

  host.append(h('label', { textContent: 'Назва шаблону' }), titleInput, addRuleBtn, rulesHost, saveBtn, cancelBtn);
  renderRules();

  const modalId = window.UI.modal.open({ title: 'Редактор шаблону перенесення', contentNode: host, closeOnOverlay: true });
  return modalId;
}
