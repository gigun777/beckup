# Аналіз інтеграції моделі backup з `oldbeckup` у `newversion`

## Контекст

У `oldbeckup` backup-модель побудована навколо:
- **типізованих backup payload** (`settings`, `transferRules`, `settingsPartial`, `transferRulesPartial`);
- **маніфесту ключів конфігурації** (`__backup_manifest_v1`) для автоматичного підхоплення нових `cfgSet` ключів;
- **таб-орієнтованого partial backup/import** (sheets/columns/export/addform/transfer).

У `newversion` backup вже є модульним і провайдер-орієнтованим:
- ядро збирає bundle формату `sdo-backup`;
- модулі додають свої дані через `backup.registerProvider(...)`;
- є `integrity` + optional encryption + import mode (`merge/replace` для провайдерів).

## Що вже сумісне концептуально

1. **Провайдерна модель `newversion` добре відповідає old-підходу «групами даних»**:
   - у `oldbeckup` групи — це backup type + optionIds;
   - у `newversion` групи — це provider id + payload data.

2. **Transfer backup інтегрується майже напряму**:
   - `oldbeckup` переносив `transfer_rules` + `transfer_templates_v2`;
   - `newversion` уже має `transfer-templates` provider (`transfer:templates:v1`).

3. **Core already has integrity+crypto**, чого в старій моделі не було як стандарту:
   - це підсилює old partial backup при перенесенні в новий формат.

## Основні розбіжності

1. **Рівень гранулярності**
   - `oldbeckup`: key-level backup через manifest + optionIds.
   - `newversion`: provider-level backup без вбудованого key manifest на рівні storage.

2. **Семантика partial backup**
   - `oldbeckup`: «вкладка налаштувань» = набір ключів.
   - `newversion`: немає стандартизованого `scope=tab` профілю, лише загальний `scope` і список `modules`.

3. **Ключі storage різні**
   - `oldbeckup`: `all_sheets`, `sheet_settings`, `add_fields`, `transfer_rules`, ...
   - `newversion`: namespaced ключі (наприклад, `@sdo/module-table-renderer:settings`, `transfer:templates:v1`).

## Висновок по інтеграції

**Інтеграція можлива і доцільна**, але краще робити її не як прямий перенос `__backup_manifest_v1`, а як **provider-backed manifest layer** у `newversion`.

Тобто:
- не відтворювати старий `cfgSet`-скан по всьому storage;
- додати єдиний backup-реєстр capability-полів для провайдерів;
- зберегти UX старого partial backup (вибір галочками), але мапити його на provider payload selectors.

## Рекомендована цільова архітектура (newversion)

### 1) Backup Capability Manifest (новий контракт)

Для кожного provider додати optional метод:

```js
getCapabilities() => {
  groups: [
    { id: 'columns.hidden', label: 'Колонки: hidden', default: true },
    { id: 'columns.widths', label: 'Колонки: widths', default: true },
    { id: 'export.profile', label: 'Експортні профілі', default: true }
  ]
}
```

і розширити:

```js
export({ includeUserData, scope, selectors })
import(payload, { mode, includeUserData, selectors })
```

`selectors` = обрані групи (аналог old `optionIds`).

### 2) Перенести old partial UX у new settings UI

Розширити backup settings feature:
- показувати групи від кожного provider;
- дозволяти export/import по групах;
- сформувати unified backup JSON із `selectors` metadata.

### 3) Для table settings provider

Розбити payload на логічні частини:
- `hiddenCols`
- `colWidths`
- `exportProfile`

і дозволити селективний export/import цих частин.

### 4) Для transfer provider

Підтримати обидва шари (опційно):
- `templates` (вже є);
- `rulesLegacy` (опціонально, лише для міграції old backup JSON).

### 5) Міст сумісності old → new (import only)

Додати `legacyBackupAdapter`:
1. розпізнає `oldbeckup` backup type;
2. мапить old keys на new providers;
3. формує `bundle.modules[providerId].data`;
4. викликає стандартний `importBackup`.

Це дасть міграцію історичних backup файлів без засмічення core API.


## Порівняння бази даних/сховища (oldbeckup vs newversion)

### `oldbeckup` (IndexedDB, multi-store)

- Фізичний DB: `dilovodstvoDB_modular`.
- Object stores:
  - `config` — ключ-значення конфігурації (`cfgGet/cfgSet`);
  - `rows` — записи журналів (індекс `by_journalKey`);
  - `cases` — сутності справ;
  - `case_rows` — записи опису справ (індекс `by_caseId`).
- Є вбудований службовий ключ `__backup_manifest_v1` для класифікації ключів на `settings/transferRules/excluded`.

**Наслідок:** old-модель має сильну привʼязку до конкретної фізичної структури IndexedDB + ключів конфігурації.

### `newversion` (adapter-based key-value storage)

- Немає жорстко прошитої фізичної БД у core.
- Є контракт storage adapter: `get/set/del` (+ optional `list`).
- Поточні реалізації:
  - `createMemoryStorage` (Map);
  - `createLocalStorageStorage` (localStorage з namespace).
- Навігаційний стан зберігається наборами ключів `*_v2` (`spaces_nodes_v2`, `journals_nodes_v2`, `core_settings_v2`, `core_revision_v2`, ...).

**Наслідок:** new-модель абстрагує фізичний storage, тому backup має бути логічно модульним (через providers), а не привʼязаним до сканування фізичних таблиць.

## Ключові розбіжності моделі даних

1. **Store-centric vs Key-space-centric**
   - `oldbeckup`: доменні сутності розкладені по object stores (`rows/cases/case_rows`).
   - `newversion`: сутності лежать у namespaced key-space і контролюються модулями.

2. **Явна schema stores vs еволюційні ключі v2**
   - `oldbeckup`: схема БД фіксується в `onupgradeneeded` IndexedDB.
   - `newversion`: еволюція через нові/версійовані ключі (`*_v2`) та код міграцій у модулях.

3. **Global manifest backup vs provider ownership**
   - `oldbeckup`: один маніфест знає про всі config-ключі.
   - `newversion`: кожен provider відповідає за власний експорт/імпорт.

4. **Високий ризик key-collision у legacy-підході при growth**
   - `oldbeckup`: багато «плоских» ключів у `config`.
   - `newversion`: очікується namespacing (`moduleId:*`, `@sdo/...`).

## Імовірні проблеми при інтеграції

1. **Неповна міграція даних журналів**
   - Якщо переносити лише settings/transfer, але не мапити `rows/cases/case_rows`, частина старих даних залишиться недоступною.

2. **Втрата семантики partial backup**
   - Old `optionIds` (columns/export/addform) можуть не мати 1:1 відображення у new providers без додаткових `selectors`.

3. **Конфлікти merge через різні key-стратегії**
   - У newversion merge відбувається всередині provider; без per-group policy можливе перетирання налаштувань.

4. **Розбіжність версій форматів шаблонів перенесення**
   - `oldbeckup`: `transfer_templates_v2` (legacy shape).
   - `newversion`: `transfer:templates:v1` (інший контракт/зберігання).

5. **Проблеми продуктивності при спробі «емуляції old manifest»**
   - Скан всього key-space в adapter storage суперечить design-цілям newversion і погіршує масштабування.

6. **Складність відкату/аудиту**
   - У old-version немає revision log на рівні core state як у newversion; змішані імпорти старих backup можуть давати складні edge-cases у delta-ланцюжку.

## Що треба закласти, щоб уникнути проблем

- Чіткий **mapping spec**: `old key/store -> new provider/group`.
- `legacy import adapter` з dry-run preview (що буде імпортовано / що пропущено).
- Per-group merge policy (`replace`, `merge`, `skip-if-exists`).
- Обовʼязкова версіонізація payload migration для кожного provider.
- Smoke/regression тести на реальних old backup файлах.

## Поетапний план впровадження

### Phase A (швидка цінність)
1. Додати `selectors` у `exportBackup/importBackup` plumbing.
2. Оновити `table-settings` provider до partial groups.
3. Оновити `features_backup_settings.js` UI для вибору груп.

### Phase B (сумісність)
4. Додати `legacyBackupAdapter` для `oldbeckup` JSON.
5. Додати smoke-tests на import старих backup.

### Phase C (hardening)
6. Валідація схем провайдерів на export/import.
7. Політика merge/replace per-group.
8. Report UI: `applied/warnings/skipped` на групу.

## Ризики і як їх зняти

1. **Розсинхрон payload shape між версіями модулів**
   - Рішення: `moduleVersion` + migration map у provider.

2. **Конфлікти merge при partial import**
   - Рішення: per-group merge strategy (`replace`, `deep-merge`, `append-uniq`).

3. **Непрозорість для користувача**
   - Рішення: показувати preview (що саме буде застосовано) перед import.

## Практична оцінка

- Технічно: **висока сумісність (8/10)**.
- Обсяг робіт: **середній** (ядро + 2-3 провайдери + settings UI).
- Ризик: **помірний**, керований через adapter + versioned migrations.

## Рекомендація

Інтегрувати old-модель **концептуально** (селективність + групи + migration),
але залишити нативний формат `newversion` (`sdo-backup`) як єдиний канонічний transport.
