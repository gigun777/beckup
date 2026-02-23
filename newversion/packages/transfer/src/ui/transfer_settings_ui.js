export async function renderTransferSettings(container, { core, ui }) {
  const templates = await core.templates.list();
  container.innerHTML = '';

  const list = document.createElement('div');
  const add = document.createElement('button');
  add.textContent = 'Створити шаблон';
  add.onclick = async () => {
    const tpl = { id: `tpl-${Date.now()}`, title: 'Новий шаблон', rules: [] };
    await core.templates.save(tpl);
    await renderTransferSettings(container, { core, ui });
  };

  container.append(add, list);

  for (const tpl of templates) {
    const row = document.createElement('div');
    const name = document.createElement('strong');
    name.textContent = tpl.title;

    const edit = document.createElement('button');
    edit.textContent = 'Редагувати';
    edit.onclick = () => {
      const content = document.createElement('div');
      const input = document.createElement('input');
      input.value = tpl.title;
      const save = document.createElement('button');
      save.textContent = 'Зберегти';
      save.onclick = async () => {
        await core.templates.save({ ...tpl, title: input.value });
        ui.closeModal(modalId);
        await renderTransferSettings(container, { core, ui });
      };
      content.append(input, save);
      const modalId = ui.openModal({ title: 'Редактор шаблону', contentNode: content });
    };

    const remove = document.createElement('button');
    remove.textContent = 'Видалити';
    remove.onclick = async () => {
      await core.templates.remove(tpl.id);
      await renderTransferSettings(container, { core, ui });
    };

    row.append(name, edit, remove);
    list.append(row);
  }
}
