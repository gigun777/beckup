/**
 * UX|UI settings feature module.
 */
(function attachUxUiSettingsFeature(global) {
  const UI = (global.UI = global.UI || {});
  UI.settings = UI.settings || {};

  function createUxUiSettingsFeature() {
    return {
      id: 'uxui',
      title: 'UX|UI',
      order: 20,
      sections: [
        {
          id: 'appearance',
          title: 'Зовнішній вигляд',
          order: 10,
          renderContent(container, ctx) {
            const UIX = (typeof window !== 'undefined' ? window.UI : globalThis.UI);
            container.innerHTML = '';

            if (!UIX || typeof UIX.getSettings !== 'function' || typeof UIX.applySettings !== 'function') {
              container.textContent = 'UX|UI: UI engine не ініціалізований (UI.getSettings/applySettings відсутні).';
              return;
            }

            const initial = UIX.getSettings();
            const draft = { ...initial };
            let dirty = false;

            const root = document.createElement('div');
            root.className = 'uxui-settings';

            const header = document.createElement('div');
            header.className = 'uxui-header';
            header.textContent = 'Налаштування UX|UI';
            root.appendChild(header);

            const body = document.createElement('div');
            body.className = 'uxui-body';
            root.appendChild(body);

            const footer = document.createElement('div');
            footer.className = 'uxui-footer';
            root.appendChild(footer);

            function markDirty() {
              dirty = true;
              applyBtn.disabled = false;
              resetBtn.disabled = false;
            }

            function row(labelText, controlEl, hintText) {
              const r = document.createElement('div');
              r.className = 'uxui-row';

              const label = document.createElement('div');
              label.className = 'uxui-label';
              label.textContent = labelText;

              const ctrl = document.createElement('div');
              ctrl.className = 'uxui-control';
              ctrl.appendChild(controlEl);

              r.appendChild(label);
              r.appendChild(ctrl);

              if (hintText) {
                const hint = document.createElement('div');
                hint.className = 'uxui-hint';
                hint.textContent = hintText;
                r.appendChild(hint);
              }
              return r;
            }

            function makeSelect(options, value) {
              const sel = document.createElement('select');
              sel.className = 'uxui-select';
              for (const [val, title] of options) {
                const o = document.createElement('option');
                o.value = val;
                o.textContent = title;
                if (String(val) === String(value)) o.selected = true;
                sel.appendChild(o);
              }
              return sel;
            }

            function makeToggle(checked) {
              const input = document.createElement('input');
              input.type = 'checkbox';
              input.className = 'uxui-toggle';
              input.checked = !!checked;
              return input;
            }

            // Theme
            const themeSel = makeSelect([['light','Денна'], ['dark','Нічна']], draft.theme || 'light');
            themeSel.addEventListener('change', () => { draft.theme = themeSel.value; markDirty(); });
            body.appendChild(row('Тема', themeSel));

            // Scale
            const scaleSel = makeSelect([['1','100%'], ['1.1','110%'], ['1.2','120%'], ['1.3','130%'], ['1.4','140%']], String(draft.scale ?? 1));
            scaleSel.addEventListener('change', () => { draft.scale = Number(scaleSel.value); markDirty(); });
            body.appendChild(row('Масштаб', scaleSel, 'Рекомендовано 110–130% для сенсорних екранів.'));

            // Touch mode
            const touchToggle = makeToggle(draft.touchMode);
            touchToggle.addEventListener('change', () => { draft.touchMode = touchToggle.checked; markDirty(); });
            body.appendChild(row('Touch mode', touchToggle, 'Більші поля/кнопки для сенсора.'));

            // Navigation circles
            const navToggle = makeToggle(draft.navCircles);
            navToggle.addEventListener('change', () => { draft.navCircles = navToggle.checked; markDirty(); });
            body.appendChild(row('Круглі кнопки навігації', navToggle));

            // Gestures
            const gestToggle = makeToggle(draft.gestures);
            gestToggle.addEventListener('change', () => { draft.gestures = gestToggle.checked; markDirty(); });
            body.appendChild(row('Жести', gestToggle));

            // Table density
            const densSel = makeSelect([['normal','Звичайна'], ['compact','Компактна']], draft.tableDensity || 'normal');
            densSel.addEventListener('change', () => { draft.tableDensity = densSel.value; markDirty(); });
            body.appendChild(row('Щільність таблиці', densSel));

            // Buttons
            const applyBtn = document.createElement('button');
            applyBtn.className = 'btn uxui-btn-primary';
            applyBtn.textContent = 'Застосувати';
            applyBtn.disabled = true;

            const resetBtn = document.createElement('button');
            resetBtn.className = 'btn uxui-btn';
            resetBtn.textContent = 'Скасувати';
            resetBtn.disabled = true;

            footer.appendChild(resetBtn);
            footer.appendChild(applyBtn);

            function computePatch(from, to) {
              const patch = {};
              const keys = Object.keys(to);
              for (const k of keys) {
                if (to[k] !== from[k]) patch[k] = to[k];
              }
              return patch;
            }

            applyBtn.addEventListener('click', () => {
              const current = UIX.getSettings();
              const patch = computePatch(current, draft);
              UIX.applySettings(patch);
              dirty = false;
              applyBtn.disabled = true;
              resetBtn.disabled = true;
            });

            resetBtn.addEventListener('click', () => {
              const cur = UIX.getSettings();
              for (const k of Object.keys(draft)) draft[k] = cur[k];
              // sync controls
              themeSel.value = draft.theme || 'light';
              scaleSel.value = String(draft.scale ?? 1);
              touchToggle.checked = !!draft.touchMode;
              navToggle.checked = !!draft.navCircles;
              gestToggle.checked = !!draft.gestures;
              densSel.value = draft.tableDensity || 'normal';
              dirty = false;
              applyBtn.disabled = true;
              resetBtn.disabled = true;
            });

            container.appendChild(root);

            // Live sync (if external changes happen while open)
            const off = UIX.on?.('settingsChanged', (next) => {
              if (dirty) return; // don't override user draft
              try {
                draft.theme = next.theme || 'light';
                draft.scale = next.scale ?? 1;
                draft.touchMode = !!next.touchMode;
                draft.navCircles = !!next.navCircles;
                draft.gestures = !!next.gestures;
                draft.tableDensity = next.tableDensity || 'normal';

                themeSel.value = draft.theme;
                scaleSel.value = String(draft.scale);
                touchToggle.checked = draft.touchMode;
                navToggle.checked = draft.navCircles;
                gestToggle.checked = draft.gestures;
                densSel.value = draft.tableDensity;
              } catch (e) {}
            });

            return () => { try { off && off(); } catch (e) {} };
          },
          onConfirm: ({ changes }) => {
            UI.applySettings?.(changes || {});
          }
        }
      ]
    };
  }

  function registerUxUiSettingsFeature() {
    const feature = createUxUiSettingsFeature();
    UI.settings.registry?.registerFeature(feature);
    return feature;
  }

  UI.settings.createUxUiSettingsFeature = createUxUiSettingsFeature;
  UI.settings.registerUxUiSettingsFeature = registerUxUiSettingsFeature;
})(typeof window !== 'undefined' ? window : globalThis);
