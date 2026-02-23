# UI/UX Manager Engine (Production Notes)

Це централізований файл опису всього `src/ui` модуля для інтеграції в існуючий сайт.

## Репозиторій: короткий аудит

На момент цієї ітерації в репозиторії наявні тільки `src/ui` та `dist/ui`.
Файли існуючого сайту (`index.html`, `app.js`, таблиці модулів) у поточній гілці відсутні,
тому інтеграційні точки нижче описані як контракт для підключення в `main`/production дерево.

---

## 1) Призначення

`UI/UX Manager` ізолює UX/UI-логіку від бізнес-модулів і дає єдині API для:

- theme management (light/dark/gray/contrast),
- UI settings state (scale, touch-mode, gestures, etc.),
- modal/form/control/toast компонентів,
- часткового/повного backup UX|UI налаштувань.

---

## 2) Правильний порядок підключення

1. `dist/ui/theme.css`
2. `dist/ui/theme_tokens.js`
3. `dist/ui/theme.js`
4. `dist/ui/ui_manager.js`
5. `dist/ui/ui_controls.js`
6. `dist/ui/ui_modal.js`
7. `dist/ui/ui_form.js`
8. `dist/ui/ui_toast.js`
9. `dist/ui/ui_backup.js`

10. `dist/ui/settings/settings_registry.js`
11. `dist/ui/settings/settings_state.js`
12. `dist/ui/settings/features_table_settings.js`
13. `dist/ui/settings/features_uxui_settings.js`
14. `dist/ui/settings/features_backup_settings.js`
15. `dist/ui/settings/settings_shell_modal.js`
16. `dist/ui/settings/settings_init.js`

> Важливо: `theme.css` має підключатись раніше інших UI стилів, щоб уникати flicker і локальних перевизначень.

---

## 3) Інтеграція у `app.js` (контракт)

```js
// (опційно) якщо у вашого застосунку є власне storage API:
UI.storage = {
  getItem(key) { return window.localStorage.getItem(key); },
  setItem(key, value) { window.localStorage.setItem(key, value); }
};

// На старті застосунку (рекомендовано):
UI.bootstrap({
  settingsHost: '#ux-ui-settings-container',
  mountOptions: { liveSync: true }
});

// Або вручну:
UI.settings.init();
UI.init();
UI.mountSettingsTab('#ux-ui-settings-container', { liveSync: true });

// Реакція модулів на зміни:
UI.on('themeChanged', (theme) => console.log('Theme:', theme));
UI.on('scaleChanged', (scale) => console.log('Scale:', scale));
```

---

## 4) API модуля тем

### `UITheme.initTheme()`
- зчитує `ui.theme`;
- якщо не знайдено, бере `prefers-color-scheme`;
- застосовує тему через `document.documentElement.dataset.theme`.

### `UITheme.applyTheme(themeName)`
- перевіряє підтримку теми;
- встановлює `data-theme` на root;
- зберігає у сховище.

> Сумісність: також доступні глобальні функції `initTheme/applyTheme/toggleTheme/getTheme`.

### `UITheme.toggleTheme()`
- перемикає `dark <-> light`.

### `UITheme.getTheme()`
- повертає активну тему.

---

## 5) API менеджера стану

### `UI.bootstrap(options)`
- запускає ініціалізацію в порядку: storage adapter → `UI.settings.init()` → theme init → `UI.init()`;
- якщо задано `settingsHost`, використовує `UI.mountSettingsTab()` (fallback: `UI.renderSettingsTab()`);
- повертає `{ theme, settings, settingsMount }` для керованого lifecycle.

### `UI.init()`
- читає `ui.settings`;
- ініціалізує тему;
- застосовує root dataset-флаги (`touchMode`, `navCircles`, `gestures`, `tableDensity`) і `--ui-scale`.

### `UI.applySettings(partialSettings)`
- застосовує часткові зміни без повного перерендеру;
- нормалізує значення (напр. `scale` в діапазоні `1..1.4`);
- пише у `ui.settings`.

### `UI.getSettings()`
- повертає snapshot поточного стану.

### `UI.on(eventName, cb)` / `UI.emit(eventName, payload)`
Події:
- `themeChanged`
- `scaleChanged`
- `settingsChanged`

### `UI.renderSettingsTab(container)`
- одноразово рендерить UI вкладки UX|UI для керування theme/scale/touchMode/navCircles/gestures/tableDensity.

### `UI.mountSettingsTab(container, { liveSync? })`
- монтує UX|UI вкладку в host та повертає handle з `destroy()`;
- при `liveSync !== false` перепромальовує вкладку на `settingsChanged`.

---

## 6) API компонентів

### Модалки (`UI.modal`)
- `open({ title?, contentNode|html, footerButtons?, onClose?, closeOnOverlay?, escClose? })`
- `close(modalId?)`
- `alert(text, { title? })`
- `confirm(text, { title?, okText?, cancelText? }) => Promise<boolean>`

Гарантується:
- один overlay,
- стек модалок,
- scroll-lock,
- focus trap,
- ESC close (конфігуровано).

### Форми (`UI.form`)
- `build(container, schema, { values?, onSubmit?, quickAdd?, submitText? })`
- `getValues(container)`

Підтримка:
- типи `text|number|date|textarea|select|checkbox`,
- required/pattern/min/max/step,
- Enter/Shift+Enter/Ctrl+Enter/Esc навігація.


### `UI.settings.openModal({ initialFeature?, initialSection?, mobilePane?, actor? })`
- відкриває глобальне settings-вікно з feature-level + section-level навігацією;
- confirm payload: `{ featureId, sectionId, changes, source, timestamp, actor }`;
- підтримує mobile pane flow: `feature -> section -> content`.

### Feature-level navigation (Patch 4)
- shell працює з кількома фічами через registry (`table`, `uxui`, `backup`);
- при перемиканні feature/section та при закритті застосовується `canLeave` guard поточної секції;
- draft зберігається namespaced-ключами `featureId:sectionId`;
- для UX|UI зміни йдуть у draft і застосовуються на Confirm через `UI.applySettings` (без auto-apply під час редагування).

### Контроли (`UI.controls`)
- `button(opts)`
- `toggle(opts)`
- `select(opts)`
- `iconButton(opts)`

### Тости (`UI.toast`)
- `show(text, { type, timeout })`
- `undo(text, { timeout, onUndo, undoText })`

---

## 7) Backup API

### `UI.backup.export(sectionName, options)`
- експортує JSON або Blob;
- підтримує частковий export через `options.keys`;
- підтримує `sectionName = 'all'` для експорту всіх секцій manifest.

### `UI.backup.import(json)`
- імпортує backup і застосовує валідні ключі через `UI.applySettings`.

### `UI.backup.getManifest()`
- повертає список доступних ключів для backup.

---

## 8) Практичні правила для команди

1. Не створювати “ручні” модалки: тільки через `UI.modal`.
2. Форми типу “+ Додати” будувати через `UI.form`.
3. Нові UI-елементи фарбувати тільки `var(--token)`.
4. Налаштування UX|UI змінювати через `UI.applySettings()`.
5. Для backup/import не дублювати логіку — тільки `UI.backup`.

---

## 8.1) Package entrypoints (Patch 7)

Для npm/package інтеграції використовуйте ESM entrypoint-и:

- `@sdo/core/ui` → агрегований `dist/ui/index.js`;
- `@sdo/core/ui/theme.css` та `@sdo/core/ui/styles.css` для стилів;
- точкові підмодулі: `@sdo/core/ui/modal`, `@sdo/core/ui/form`, `@sdo/core/ui/manager`,
  `@sdo/core/ui/settings/*`.

Це дозволяє поступово переходити на ESM-first імпорт без зламу backward-compatible
global API в браузерному runtime.


---


## Legacy cleanup status

- Legacy `openSettingsModal()` у `ui_core` відключено від table-specific реалізації й переведено на виклик `UI.settings.openModal(...)`.
- Для сумісності стилі/класи legacy (`.sdo-settings-nav-list`, `.sdo-settings-nav-item.is-selected`) тимчасово залишені, але більше не є джерелом правди для нової shell-модалки.
- Подальший крок: прибрати невикористовувані legacy-стилі після фінального переходу всіх модулів на Patch 4 shell API.

---

## 9) Mobile/WebView поведінка settings shell

- На мобільних shell перемикає панелі у порядку: `feature` → `section` → `content`.
- Кнопка Back йде у зворотному порядку (`content -> section -> feature`), а потім закриває модалку.
- На desktop відображаються всі панелі одночасно.
