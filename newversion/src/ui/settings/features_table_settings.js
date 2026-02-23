/**
 * Table settings feature module.
 */
(function attachTableSettingsFeature(global) {
  const UI = (global.UI = global.UI || {});
  UI.settings = UI.settings || {};

  const TABLE_SETTINGS_KEY = '@sdo/module-table-renderer:settings';

  function sectionContent(title, description) {
    return function render(container) {
      container.innerHTML = '';
      const h = document.createElement('h4');
      h.textContent = title;
      const p = document.createElement('p');
      p.textContent = description;
      container.append(h, p);
    };
  }

  function getSettingsStorage() {
    if (UI.storage && typeof UI.storage.get === 'function' && typeof UI.storage.set === 'function') {
      return UI.storage;
    }

    if (UI.storage && typeof UI.storage.getItem === 'function' && typeof UI.storage.setItem === 'function') {
      return {
        get: async (key) => {
          const raw = UI.storage.getItem(key);
          if (raw == null) return null;
          try { return JSON.parse(raw); } catch { return raw; }
        },
        set: async (key, value) => {
          UI.storage.setItem(key, JSON.stringify(value));
        }
      };
    }

    return {
      get: async (key) => {
        const raw = global.localStorage?.getItem?.(key);
        if (raw == null) return null;
        try {
          return JSON.parse(raw);
        } catch {
          return raw;
        }
      },
      set: async (key, value) => {
        global.localStorage?.setItem?.(key, JSON.stringify(value));
      }
    };
  }

  async function readTableSettings(journalId) {
    const storage = getSettingsStorage();
    try {
      const value = await storage.get(journalId ? `${TABLE_SETTINGS_KEY}:${journalId}` : TABLE_SETTINGS_KEY);
      const legacy = journalId ? await storage.get(TABLE_SETTINGS_KEY) : null;
      
      return (value ?? legacy) ?? { columns: { visibility: {} }, subrows: { columnsSubrowsEnabled: {} } };
    } catch {
      return { columns: { visibility: {} }, subrows: { columnsSubrowsEnabled: {} } };
    }
  }

  function renderColumnsSettingsSection(container) {
    container.innerHTML = '';
    const header = document.createElement('h4');
    header.textContent = 'Колонки';
    const desc = document.createElement('p');
    desc.textContent = 'Увімкніть підстроки для потрібних колонок.';
    container.append(header, desc);

    const list = document.createElement('div');
    list.style.display = 'grid';
    list.style.gap = '8px';
    container.append(list);

    const saveBtn = document.createElement('button');
    saveBtn.textContent = 'Зберегти';
    saveBtn.style.marginTop = '12px';
    container.append(saveBtn);

    let settings = { columns: { visibility: {} }, subrows: { columnsSubrowsEnabled: {} } };

    const run = async () => {
      const state0 = UI.sdo?.getState?.() ?? { journals: [], activeJournalId: null };
      settings = await readTableSettings(state0.activeJournalId);
      const state = UI.sdo?.getState?.() ?? { journals: [], activeJournalId: null };
      const activeJournal = (state.journals ?? []).find((j) => j.id === state.activeJournalId) ?? null;
      const templateId = activeJournal?.templateId;
      const template = templateId ? await UI.sdo?.journalTemplates?.getTemplate?.(templateId) : null;
      const columns = template?.columns ?? [];

      list.innerHTML = '';
      for (const column of columns) {
        const row = document.createElement('label');
        row.style.display = 'flex';
        row.style.alignItems = 'center';
        row.style.gap = '8px';

        const subrows = document.createElement('input');
        subrows.type = 'checkbox';
        subrows.checked = settings.subrows?.columnsSubrowsEnabled?.[column.key] === true;
        subrows.addEventListener('change', () => {
          settings = {
            ...settings,
            subrows: {
              ...(settings.subrows ?? { columnsSubrowsEnabled: {} }),
              columnsSubrowsEnabled: {
                ...((settings.subrows ?? {}).columnsSubrowsEnabled ?? {}),
                [column.key]: subrows.checked
              }
            }
          };
        });

        const text = document.createElement('span');
        text.textContent = `${column.label} (${column.key})`;
        row.append(subrows, text);
        list.append(row);
      }
    };

    saveBtn.addEventListener('click', async () => {
      await UI.storage?.set(TABLE_SETTINGS_KEY, settings);
      UI.toast?.show?.('Налаштування колонок збережено');
    });

    run();
  }

  async function readTableSettings(journalId) {
    const storage = getSettingsStorage();
    try {
      const value = await storage.get(journalId ? `${TABLE_SETTINGS_KEY}:${journalId}` : TABLE_SETTINGS_KEY);
      const legacy = journalId ? await storage.get(TABLE_SETTINGS_KEY) : null;
      
      return (value ?? legacy) ?? { columns: { visibility: {} }, subrows: { columnsSubrowsEnabled: {} } };
    } catch {
      return { columns: { visibility: {} }, subrows: { columnsSubrowsEnabled: {} } };
    }
  }

  function createColumnsSettingsNode(settings, columns, onChange) {
    const wrapper = document.createElement('div');
    wrapper.style.display = 'grid';
    wrapper.style.gap = '8px';

    for (const column of columns) {
      const row = document.createElement('label');
      row.style.display = 'flex';
      row.style.alignItems = 'center';
      row.style.gap = '8px';

      const subrows = document.createElement('input');
      subrows.type = 'checkbox';
      subrows.checked = settings.subrows?.columnsSubrowsEnabled?.[column.key] === true;
      subrows.addEventListener('change', () => {
        const next = {
          ...settings,
          subrows: {
            ...(settings.subrows ?? { columnsSubrowsEnabled: {} }),
            columnsSubrowsEnabled: {
              ...((settings.subrows ?? {}).columnsSubrowsEnabled ?? {}),
              [column.key]: subrows.checked
            }
          }
        };
        onChange(next);
      });

      const text = document.createElement('span');
      text.textContent = `${column.label} (${column.key})`;
      row.append(subrows, text);
      wrapper.append(row);
    }

    return wrapper;
  }

  function renderColumnsSettingsSection(container) {
    container.innerHTML = '';

    const header = document.createElement('h4');
    header.textContent = 'Колонки';
    const desc = document.createElement('p');
    desc.textContent = 'Відкрийте модалку налаштувань і увімкніть підстроки для потрібних колонок.';
    const openBtn = document.createElement('button');
    openBtn.textContent = 'Налаштувати колонки';

    container.append(header, desc, openBtn);

    openBtn.addEventListener('click', async () => {
      const state0 = UI.sdo?.getState?.() ?? { journals: [], activeJournalId: null };
      settings = await readTableSettings(state0.activeJournalId);
      const state = UI.sdo?.getState?.() ?? { journals: [], activeJournalId: null };
      const activeJournal = (state.journals ?? []).find((j) => j.id === state.activeJournalId) ?? null;
      const templateId = activeJournal?.templateId;
      const template = templateId ? await UI.sdo?.journalTemplates?.getTemplate?.(templateId) : null;
      const columns = template?.columns ?? [];

      if (!UI.modal?.open) {
        UI.toast?.show?.('Модалка недоступна в цьому середовищі');
        return;
      }

      const body = document.createElement('div');
      body.style.display = 'grid';
      body.style.gap = '12px';

      const listWrap = document.createElement('div');
      const saveBtn = document.createElement('button');
      saveBtn.textContent = 'Зберегти';

      const rerenderList = () => {
        listWrap.innerHTML = '';
        listWrap.append(createColumnsSettingsNode(settings, columns, (next) => {
          settings = next;
          rerenderList();
        }));
      };
      rerenderList();

      body.append(listWrap, saveBtn);

      const modalId = UI.modal.open({
        title: 'Налаштування колонок',
        contentNode: body,
        closeOnOverlay: true,
        escClose: true
      });

      saveBtn.addEventListener('click', async () => {
        const storage = getSettingsStorage();
        const state1 = UI.sdo?.getState?.() ?? { activeJournalId: null };
        const jid = state1.activeJournalId;
        if(jid) await storage.set(`${TABLE_SETTINGS_KEY}:${jid}`, settings);
        await storage.set(TABLE_SETTINGS_KEY, settings);
        UI.toast?.show?.('Налаштування колонок збережено');
        UI.modal.close(modalId);
      });
    });
  }

  async function readTableSettings(journalId) {
    const storage = getSettingsStorage();
    try {
      const value = await storage.get(journalId ? `${TABLE_SETTINGS_KEY}:${journalId}` : TABLE_SETTINGS_KEY);
      const legacy = journalId ? await storage.get(TABLE_SETTINGS_KEY) : null;
      
      return (value ?? legacy) ?? { columns: { visibility: {} }, subrows: { columnsSubrowsEnabled: {} } };
    } catch {
      return { columns: { visibility: {} }, subrows: { columnsSubrowsEnabled: {} } };
    }
  }

  function createColumnsSettingsNode(settings, columns, onChange) {
    const wrapper = document.createElement('div');
    wrapper.style.display = 'grid';
    wrapper.style.gap = '8px';

    for (const column of columns) {
      const row = document.createElement('label');
      row.style.display = 'flex';
      row.style.alignItems = 'center';
      row.style.gap = '8px';

      const subrows = document.createElement('input');
      subrows.type = 'checkbox';
      subrows.checked = settings.subrows?.columnsSubrowsEnabled?.[column.key] === true;
      subrows.addEventListener('change', () => {
        const next = {
          ...settings,
          subrows: {
            ...(settings.subrows ?? { columnsSubrowsEnabled: {} }),
            columnsSubrowsEnabled: {
              ...((settings.subrows ?? {}).columnsSubrowsEnabled ?? {}),
              [column.key]: subrows.checked
            }
          }
        };
        onChange(next);
      });

      const text = document.createElement('span');
      text.textContent = `${column.label} (${column.key})`;
      row.append(subrows, text);
      wrapper.append(row);
    }

    return wrapper;
  }

  function renderColumnsSettingsSection(container) {
    container.innerHTML = '';

    const header = document.createElement('h4');
    header.textContent = 'Колонки';
    const desc = document.createElement('p');
    desc.textContent = 'Відкрийте модалку налаштувань і увімкніть підстроки для потрібних колонок.';
    const openBtn = document.createElement('button');
    openBtn.textContent = 'Налаштувати колонки';

    container.append(header, desc, openBtn);

    openBtn.addEventListener('click', async () => {
      const state0 = UI.sdo?.getState?.() ?? { journals: [], activeJournalId: null };
      settings = await readTableSettings(state0.activeJournalId);
      const state = UI.sdo?.getState?.() ?? { journals: [], activeJournalId: null };
      const activeJournal = (state.journals ?? []).find((j) => j.id === state.activeJournalId) ?? null;
      const templateId = activeJournal?.templateId;
      const template = templateId ? await UI.sdo?.journalTemplates?.getTemplate?.(templateId) : null;
      const columns = template?.columns ?? [];

      if (!UI.modal?.open) {
        UI.toast?.show?.('Модалка недоступна в цьому середовищі');
        return;
      }

      const body = document.createElement('div');
      body.style.display = 'grid';
      body.style.gap = '12px';

      const listWrap = document.createElement('div');
      const saveBtn = document.createElement('button');
      saveBtn.textContent = 'Зберегти';

      const rerenderList = () => {
        listWrap.innerHTML = '';
        listWrap.append(createColumnsSettingsNode(settings, columns, (next) => {
          settings = next;
          rerenderList();
        }));
      };
      rerenderList();

      body.append(listWrap, saveBtn);

      const modalId = UI.modal.open({
        title: 'Налаштування колонок',
        contentNode: body,
        closeOnOverlay: true,
        escClose: true
      });

      saveBtn.addEventListener('click', async () => {
        const storage = getSettingsStorage();
        const state1 = UI.sdo?.getState?.() ?? { activeJournalId: null };
        const jid = state1.activeJournalId;
        if(jid) await storage.set(`${TABLE_SETTINGS_KEY}:${jid}`, settings);
        await storage.set(TABLE_SETTINGS_KEY, settings);
        UI.toast?.show?.('Налаштування колонок збережено');
        UI.modal.close(modalId);
      });
    });
  }

  async function readTableSettings(journalId) {
    const storage = getSettingsStorage();
    try {
      const value = await storage.get(journalId ? `${TABLE_SETTINGS_KEY}:${journalId}` : TABLE_SETTINGS_KEY);
      const legacy = journalId ? await storage.get(TABLE_SETTINGS_KEY) : null;
      
      return (value ?? legacy) ?? { columns: { visibility: {} }, subrows: { columnsSubrowsEnabled: {} } };
    } catch {
      return { columns: { visibility: {} }, subrows: { columnsSubrowsEnabled: {} } };
    }
  }

  function createColumnsSettingsNode(settings, columns, onChange) {
    const wrapper = document.createElement('div');
    wrapper.style.display = 'grid';
    wrapper.style.gap = '8px';

    for (const column of columns) {
      const row = document.createElement('label');
      row.style.display = 'flex';
      row.style.alignItems = 'center';
      row.style.gap = '8px';

      const subrows = document.createElement('input');
      subrows.type = 'checkbox';
      subrows.checked = settings.subrows?.columnsSubrowsEnabled?.[column.key] === true;
      subrows.addEventListener('change', () => {
        const next = {
          ...settings,
          subrows: {
            ...(settings.subrows ?? { columnsSubrowsEnabled: {} }),
            columnsSubrowsEnabled: {
              ...((settings.subrows ?? {}).columnsSubrowsEnabled ?? {}),
              [column.key]: subrows.checked
            }
          }
        };
        onChange(next);
      });

      const text = document.createElement('span');
      text.textContent = `${column.label} (${column.key})`;
      row.append(subrows, text);
      wrapper.append(row);
    }

    return wrapper;
  }

  function renderColumnsSettingsSection(container) {
    container.innerHTML = '';

    const header = document.createElement('h4');
    header.textContent = 'Колонки';
    const desc = document.createElement('p');
    desc.textContent = 'Відкрийте модалку налаштувань і увімкніть підстроки для потрібних колонок.';
    const openBtn = document.createElement('button');
    openBtn.textContent = 'Налаштувати колонки';

    container.append(header, desc, openBtn);

    openBtn.addEventListener('click', async () => {
      const state0 = UI.sdo?.getState?.() ?? { journals: [], activeJournalId: null };
      settings = await readTableSettings(state0.activeJournalId);
      const state = UI.sdo?.getState?.() ?? { journals: [], activeJournalId: null };
      const activeJournal = (state.journals ?? []).find((j) => j.id === state.activeJournalId) ?? null;
      const templateId = activeJournal?.templateId;
      const template = templateId ? await UI.sdo?.journalTemplates?.getTemplate?.(templateId) : null;
      const columns = template?.columns ?? [];

      if (!UI.modal?.open) {
        UI.toast?.show?.('Модалка недоступна в цьому середовищі');
        return;
      }

      const body = document.createElement('div');
      body.style.display = 'grid';
      body.style.gap = '12px';

      const listWrap = document.createElement('div');
      const saveBtn = document.createElement('button');
      saveBtn.textContent = 'Зберегти';

      const rerenderList = () => {
        listWrap.innerHTML = '';
        listWrap.append(createColumnsSettingsNode(settings, columns, (next) => {
          settings = next;
          rerenderList();
        }));
      };
      rerenderList();

      body.append(listWrap, saveBtn);

      const modalId = UI.modal.open({
        title: 'Налаштування колонок',
        contentNode: body,
        closeOnOverlay: true,
        escClose: true
      });

      saveBtn.addEventListener('click', async () => {
        const storage = getSettingsStorage();
        const state1 = UI.sdo?.getState?.() ?? { activeJournalId: null };
        const jid = state1.activeJournalId;
        if(jid) await storage.set(`${TABLE_SETTINGS_KEY}:${jid}`, settings);
        await storage.set(TABLE_SETTINGS_KEY, settings);
        UI.toast?.show?.('Налаштування колонок збережено');
        UI.modal.close(modalId);
      });
    });
  }

  function createTableSettingsFeature() {
    return {
      id: 'table',
      title: 'Таблиці',
      order: 10,
      sections: [
        {
          id: 'journals',
          title: 'Журнали',
          order: 10,
          renderContent: sectionContent('Журнали', 'Керування журналами та шаблонами журналів.'),
          onConfirm: ({ draft }) => draft
        },
        {
          id: 'columns',
          title: 'Колонки',
          order: 20,
          renderContent: renderColumnsSettingsSection,
          onConfirm: ({ draft }) => draft
        },
        {
          id: 'quickAdd',
          title: 'Поля +Додати',
          order: 30,
          renderContent: sectionContent('Поля +Додати', 'Набір полів для швидкого додавання записів.'),
          onConfirm: ({ draft }) => draft
        },
        {
          id: 'transfer',
          title: 'Перенесення',
          order: 40,
          renderContent: function renderTransferSection(container){
            container.innerHTML='';
            const h=document.createElement('h3'); h.textContent='Перенесення';
            const p=document.createElement('p'); p.textContent='Шаблони перенесення між журналами та правила формування рядка.';
            const btn=document.createElement('button'); btn.textContent='Відкрити налаштування перенесення';
            btn.onclick=()=>{ const tr=(globalThis.UI?.transfer); if(tr?.openSettings) tr.openSettings(); else globalThis.UI?.toast?.warning?.('Transfer UI не готовий'); };
            container.append(h,p,btn);
          },
          onConfirm: ({ draft }) => draft
        }
      ]
    };
  }

  function registerTableSettingsFeature() {
    const feature = createTableSettingsFeature();
    UI.settings.registry?.registerFeature(feature);
    return feature;
  }

  UI.settings.createTableSettingsFeature = createTableSettingsFeature;
  UI.settings.registerTableSettingsFeature = registerTableSettingsFeature;
})(typeof window !== 'undefined' ? window : globalThis);
