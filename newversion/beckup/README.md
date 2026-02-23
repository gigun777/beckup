# @beckup/beckup

Stage-1 NPM module for backup flows aligned with `newversion` storage/module architecture.

## What is included (Phase 1)

- Generic **ZIP backup** builder (`STORE` method, no compression dependency).
- Single-sheet JSON exports:
  - `journal v2` compatible payload (`rowsV2` + legacy `rows.exportData`).
  - `excel-json` payload for tooling and tabular pipelines.
- Ready-to-integrate API for later wiring into `newversion` backup providers.

## Install (local workspace)

```bash
npm i ./newversion/beckup
```

## API

```js
import {
  createZipBackup,
  createZipBackupBlob,
  createSheetJsonPayload,
  createExcelJsonPayload,
  createSheetZipBackup
} from '@beckup/beckup';
```

### `createZipBackup(entries)`

Build zip bytes from entries:

```js
const zipBytes = createZipBackup([
  { name: 'meta.json', data: { app: 'beckup' } },
  { name: 'data.txt', data: 'hello' }
]);
```

### `createSheetJsonPayload({ sheet, records, exportProfile })`

Creates old-compatible single-sheet payload:
- `meta.type='journal'`
- `version=2`
- `rowsV2`
- legacy `rows[*].exportData`

### `createExcelJsonPayload({ sheet, records })`

Creates Excel-oriented JSON:
- `rows` (array of objects by column names)
- `matrix` (2D rows for tabular pipelines)

### `createSheetZipBackup({ sheet, records, exportProfile })`

Creates a two-file zip archive:
- `<sheet>_<stamp>.json`
- `<sheet>_<stamp>.excel.json`

Returns `{ zipBytes, files, payloads }`.

## Notes

- Module is ESM-only.
- Designed to match `newversion` style (modular APIs + storage-agnostic architecture).
- Next phases should add:
  - provider adapter for `createSEDO().backup.registerProvider(...)`
  - legacy `oldbeckup` import bridge
  - selective/partial export groups.


## Phase 2 (implemented)

### Source-of-truth export/import (DB-first)

New APIs are designed to work from data source adapters (storage/DB), not rendered UI tables:

- `exportFullJsonBackupFromSource({ source, include })`
- `importFullJsonBackupToSource(payload, { target, mode })`
- `createNewversionSourceAdapter(storage)`

### Any-Excel import (filename independent)

- `parseAnyXlsx(arrayBuffer, { worksheet })` resolves workbook/sheet via XML relationships, not by file name conventions (worksheet can be selected by `name` or `index`).
- `importAnyExcelToRecords({ arrayBuffer, targetColumns, mapping?, worksheet? })` supports:
  - manual mapping (`sourceCol -> targetKey`)
  - auto mapping by header names when mapping is omitted.

> Import UI/constructor is intentionally not included yet (per workflow): backend import/export functions first.

- JSON import in newversion adapter supports partial recovery fallback from legacy `rows[].exportData` / `rows[].data` when `rowsV2` is absent.


## Phase 4 (implemented)

- Added provider factory `createBeckupProvider({ storage })` for integration with `newversion` backup registry (`backup.registerProvider(...)`).
- Provider export/import is DB-first and reuses `createNewversionSourceAdapter(storage)` internally.
- This keeps backup logic UI-independent and ready for next steps (manual import constructor UI by command).


## Phase 5 (implemented)

- Added backend import-constructor core (no UI):
  - `suggestColumnMapping({ headerRow, targetColumns, aliases })`
  - `buildImportPlan({ mapping, targetColumns })`
  - `applyImportPlanToRows({ rows, plan, dataRowStartIndex })`
- `importAnyExcelToRecords(...)` now uses the constructor core internally (manual mapping or suggested mapping).
- This prepares the future manual mapping modal without coupling import logic to UI.
