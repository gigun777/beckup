import { renderTransferSettings } from './transfer_settings_ui.js';

export function createTransferUI({ core, ui, journals }) {
  if (!core) throw new Error('createTransferUI requires core');
  if (!ui) throw new Error('createTransferUI requires ui adapter');

  return {
    core,

    async openSettings(container) {
      return renderTransferSettings(container, { core, ui });
    },

    async openTransferModal({ sourceRef, rowIds }) {
      const targets = await journals.listTargets();
      const templates = await core.templates.list();

      const content = document.createElement('div');
      const targetSelect = document.createElement('select');
      const templateSelect = document.createElement('select');
      const previewBox = document.createElement('pre');

      for (const target of targets.filter((item) => item.id !== sourceRef.journalId)) {
        targetSelect.append(new Option(target.title ?? target.id, target.id));
      }
      for (const template of templates) {
        templateSelect.append(new Option(template.title, template.id));
      }

      let previewCtx = null;
      const previewBtn = document.createElement('button');
      previewBtn.textContent = 'Preview';
      previewBtn.onclick = async () => {
        const ctx = await core.prepareTransfer({
          templateId: templateSelect.value,
          sourceRef,
          rowIds
        });
        previewCtx = await core.preview(ctx, { targetRef: { journalId: targetSelect.value } });
        previewBox.textContent = JSON.stringify(previewCtx.report, null, 2);
      };

      const applyBtn = document.createElement('button');
      applyBtn.textContent = 'Застосувати';
      applyBtn.onclick = async () => {
        if (!previewCtx) {
          previewBtn.click();
          return;
        }
        const result = await core.commit(previewCtx);
        previewBox.textContent = JSON.stringify(result.report, null, 2);
        if (!result.report.errors.length) ui.closeModal(modalId);
      };

      const cancelBtn = document.createElement('button');
      cancelBtn.textContent = 'Скасувати';
      cancelBtn.onclick = () => ui.closeModal(modalId);

      content.append('Ціль', targetSelect, 'Шаблон', templateSelect, previewBtn, applyBtn, cancelBtn, previewBox);
      const modalId = ui.openModal({ title: 'Перенести', contentNode: content });
      return modalId;
    }
  };
}
