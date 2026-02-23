/**
 * Backup settings feature module.
 */
(function attachBackupSettingsFeature(global) {
  const UI = (global.UI = global.UI || {});
  UI.settings = UI.settings || {};

  function downloadJson(filename, text) {
    const blob = new Blob([text], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  function createBackupSettingsFeature() {
    return {
      id: 'backup',
      title: 'Backup',
      order: 30,
      sections: [
        {
          id: 'uxui-backup',
          title: 'Експорт / Імпорт UX|UI',
          order: 10,
          renderContent(container) {
            container.innerHTML = '';

            const root = document.createElement('div');
            root.className = 'sdo-settings-section';

            const desc = document.createElement('p');
            desc.textContent = 'Експортуйте поточні UX|UI налаштування або імпортуйте backup JSON.';

            const actions = document.createElement('div');
            actions.className = 'sdo-settings-actions';

            const exportUxuiBtn = document.createElement('button');
            exportUxuiBtn.className = 'btn';
            exportUxuiBtn.textContent = 'Export UX|UI';

            const exportAllBtn = document.createElement('button');
            exportAllBtn.className = 'btn';
            exportAllBtn.textContent = 'Export All';

            const importBtn = document.createElement('button');
            importBtn.className = 'btn btn-primary';
            importBtn.textContent = 'Import UX|UI';

            const fileInput = document.createElement('input');
            fileInput.type = 'file';
            fileInput.accept = 'application/json,.json';
            fileInput.hidden = true;

            const status = document.createElement('div');
            status.className = 'sdo-status';

            exportUxuiBtn.addEventListener('click', () => {
              const manifest = UI.backup?.getManifest?.() || { uxui: [] };
              const json = UI.backup?.export?.('uxui', { keys: manifest.uxui }) || '{}';
              downloadJson('uxui-backup.json', json);
              status.textContent = 'UX|UI backup експортовано.';
            });

            exportAllBtn.addEventListener('click', () => {
              const manifest = UI.backup?.getManifest?.() || {};
              const json = UI.backup?.export?.('all', {
                keys: Object.keys(manifest).reduce((acc, key) => {
                  acc[key] = manifest[key];
                  return acc;
                }, {})
              }) || '{}';
              downloadJson('ui-backup-all.json', json);
              status.textContent = 'Повний backup експортовано.';
            });

            importBtn.addEventListener('click', () => fileInput.click());

            fileInput.addEventListener('change', async () => {
              const [file] = fileInput.files || [];
              if (!file) return;
              const text = await file.text();
              const ok = UI.backup?.import?.(text);
              status.textContent = ok ? 'Backup імпортовано.' : 'Не вдалося імпортувати backup.';
              UI.toast?.show?.(status.textContent, { type: ok ? 'success' : 'error' });
              fileInput.value = '';
            });

            actions.append(exportUxuiBtn, exportAllBtn, importBtn, fileInput);
            root.append(desc, actions, status);
            container.appendChild(root);
          },
          onConfirm: ({ changes }) => changes
        }
      ]
    };
  }

  function registerBackupSettingsFeature() {
    const feature = createBackupSettingsFeature();
    UI.settings.registry?.registerFeature(feature);
    return feature;
  }

  UI.settings.createBackupSettingsFeature = createBackupSettingsFeature;
  UI.settings.registerBackupSettingsFeature = registerBackupSettingsFeature;
})(typeof window !== 'undefined' ? window : globalThis);
