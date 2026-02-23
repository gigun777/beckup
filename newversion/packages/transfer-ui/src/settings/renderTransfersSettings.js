import {
  createTemplateDraft,
  loadTemplates,
  saveTemplates,
  exportTemplatesBackup,
  importTemplatesBackup
} from '../storage/templates_store.js';
import { openTemplateEditorModal } from './templateEditorModal.js';

export async function renderTransferSettingsSection(container, api) {
  const { storageAdapter, listJournals, getSchema } = api;
  const templates = await loadTemplates(storageAdapter);

  const journalsRaw = await listJournals();
  const journals = await Promise.all(
    journalsRaw.map(async (journal) => ({
      ...journal,
      fields: (await getSchema(journal.id))?.fields ?? []
    }))
  );

  const render = async () => {
    container.innerHTML = '';

    const actions = document.createElement('div');
    actions.className = 'sdo-settings-row';

    const addBtn = document.createElement('button');
    addBtn.textContent = 'Створити шаблон';
    addBtn.onclick = () =>
      openTemplateEditorModal({
        template: createTemplateDraft(),
        journals,
        onSave: async (nextTemplate) => {
          templates.push(nextTemplate);
          await saveTemplates(storageAdapter, templates);
          await render();
        }
      });

    const exportBtn = document.createElement('button');
    exportBtn.textContent = 'Backup шаблонів';
    exportBtn.onclick = async () => {
      const json = await exportTemplatesBackup(storageAdapter);
      if (window.UI?.modal?.alert) window.UI.modal.alert(json, { title: 'Transfer templates backup JSON' });
      else window.prompt('Скопіюйте backup JSON', json);
    };

    const importBtn = document.createElement('button');
    importBtn.textContent = 'Restore шаблонів';
    importBtn.onclick = async () => {
      const json = window.prompt('Вставте backup JSON для шаблонів');
      if (!json) return;
      try {
        const ok = await importTemplatesBackup(storageAdapter, json);
        if (!ok) throw new Error('invalid payload');
        const next = await loadTemplates(storageAdapter);
        templates.splice(0, templates.length, ...next);
        await render();
      } catch (error) {
        window.UI?.toast?.show?.(`Restore помилка: ${error.message}`);
      }
    };

    actions.append(addBtn, exportBtn, importBtn);

    const list = document.createElement('div');
    templates.forEach((template) => {
      const row = document.createElement('div');
      row.className = 'sdo-settings-row';

      const title = document.createElement('strong');
      title.textContent = template.title;

      const editBtn = document.createElement('button');
      editBtn.textContent = 'Редагувати';
      editBtn.onclick = () =>
        openTemplateEditorModal({
          template,
          journals,
          onSave: async (nextTemplate) => {
            const idx = templates.findIndex((item) => item.id === nextTemplate.id);
            templates[idx] = nextTemplate;
            await saveTemplates(storageAdapter, templates);
            await render();
          }
        });

      const removeBtn = document.createElement('button');
      removeBtn.textContent = 'Видалити';
      removeBtn.onclick = async () => {
        const idx = templates.findIndex((item) => item.id === template.id);
        templates.splice(idx, 1);
        await saveTemplates(storageAdapter, templates);
        await render();
      };

      row.append(title, editBtn, removeBtn);
      list.append(row);
    });

    container.append(actions, list);
  };

  await render();
}
