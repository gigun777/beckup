# TransferUI — модалки “Налаштування → Перенесення” та “Перенести”

Дата пакету: 2026-02-19

Цей пакет містить **готові модальні вікна** (ванільний JS, без залежностей):
1) **Налаштування → Перенесення** (редактор шаблонів)
2) **Перенести** (виконання перенесення з превʼю)

## Файли
- `transfer_modals.css` — стилі
- `transfer_modals.html` — розмітка (вставити 1 раз у DOM)
- `transfer_modals.js` — логіка + `window.TransferUI`
- `README_UA.md` — ця пояснювальна записка
- `demo.html` — демо для перевірки

## Як підключити
1) CSS:
```html
<link rel="stylesheet" href="transfer_modals.css">
```
2) HTML (1 раз, в кінці body):
- вставте вміст `transfer_modals.html`
3) JS:
```html
<script src="transfer_modals.js"></script>
```

## Дані, які треба передати
### sheets
```js
sheets = [{ key, name, columns:[{id,name}, ...] }]
```
- `columns` мають бути в тому ж порядку, що й у вашій таблиці (маршрути працюють через індекси).

### templates
```js
templates = [{ id, name, fromSheetKey, toSheetKey, routes:[{sources, op, delimiter, targetCol}] }]
```
- `sources` та `targetCol` — індекси колонок.

## ВИКЛИКИ З ВАШОГО КОДУ

### Відкрити Налаштування → Перенесення
```js
TransferUI.openSettings({
  sheets: getSheetsForTransferUI(),
  templates: cfgGet("transfer_templates_v2") || [],
  onSave: (newTemplates) => cfgSet("transfer_templates_v2", newTemplates),
});
```

### Відкрити “Перенести” для виділеного рядка
```js
TransferUI.openTransfer({
  sheets: getSheetsForTransferUI(),
  templates: cfgGet("transfer_templates_v2") || [],
  sourceSheetKey: state.currentSheetKey,
  sourceRow: state.selectedRowCells, // масив значень клітинок
  onApply: ({ template, targetRow, actions }) => {
    addRowToSheet(template.toSheetKey, targetRow);
    if(actions.goToTarget) openSheet(template.toSheetKey);
  }
});
```

## Мапінг керування
### Налаштування → Перенесення
- Назва → `template.name`
- З листа → `template.fromSheetKey`
- До листа → `template.toSheetKey`
- Джерела → `route.sources[]`
- Правило → `route.op` (concat | seq | newline | sum)
- Розділювач → `route.delimiter` (тільки concat)
- Цільова колонка → `route.targetCol`
- Зберегти → ваш `cfgSet(...)` через `onSave(newTemplates)`

### Вікно “Перенести”
- Шаблон → фільтр по `fromSheetKey`
- Превʼю → `sourceRow` + routes
- Перенести → формує `targetRow` та викликає `onApply(...)`
- `actions.goToTarget` → перейти до листа-призначення
- `actions.closeOnSuccess` → закрити після успіху
