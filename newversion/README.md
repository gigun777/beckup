# @sdo/core

ESM-first headless SEDO core with module registries: UI, Schema, Commands, Settings, Backup.

## Install
```bash
npm i @sdo/core
```

## Quick start
```js
import { createSEDO, createMemoryStorage } from '@sdo/core';
import { createModuleManagerUI } from '@sdo/core/dist/ui/ui_core.js';

const sdo = createSEDO({
  storage: createMemoryStorage(),
  mount: document.getElementById('app'),
  createUI: createModuleManagerUI
});
await sdo.start();
```

## Public API
- `createSEDO(options)`
- `createNavi(storage)`
- `version`
- backup helpers: `encryptBackup/decryptBackup/signBackup/verifyBackup/verifyIntegrity`

SEDO instance:
- module lifecycle: `use`, `loadModuleFromUrl`, `start`, `destroy`
- state: `getState`, `commit`, `on/off`
- UI registry: `ui.listButtons/listPanels/subscribe`
- Schema registry: `schemas.get/list/resolve`
- Commands registry: `commands.list/run`
- Settings registry: `settings.listTabs/getKey/setKey`
- backup: `exportBackup/importBackup/exportDelta/applyDelta`

## XAMPP run
1. `npm install && npm run build`
2. Copy repo into `htdocs` (e.g. `C:\xampp\htdocs\core`)
3. Open `http://localhost/core/` (repo includes ready `index.html`)

## Module authoring
- use only `ctx` API, never import core internals
- keys must be namespaced (`moduleId:*`)
- register schema/commands/settings/ui in `init(ctx)`
- backup provider must support `describe/export/import` and `includeUserData`
- delta uses keyed patch (`set/del`) with revisions

See full spec: [`docs/MODULES_SPEC.md`](./docs/MODULES_SPEC.md)
Reference module: [`docs/reference-module.mjs`](./docs/reference-module.mjs)


## Journal templates container (@sdo/journal-templates, v0.1.0)
- API: `sdo.journalTemplates.listTemplates/getTemplate/addTemplate/deleteTemplate`
- storage keys: `templates:index`, `templates:tpl:${id}`
- first run default template: `test` with 5 columns
- UI: top-right **Шаблони** button opens templates manager modal
- journal creation now picks template list from container (not hardcoded)


## Table engine module (@sdo/module-table-engine, v1.0.0)
- API export: `createTableEngine`, `createTableEngineModule`
- V1 engine supports: columns order/visibility/width, `sub_rows` hierarchy, merge map (`rowSpan/colSpan/coveredBy`), sort, global filter, selection, inline edit patching, add-form model/validation.
- Integration-ready module registers schema/settings/commands and a toolbar button via `ctx.ui.*`.
