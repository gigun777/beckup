// js/backup_core.js
// Dilovodstvo backup core (modular IndexedDB version).
//
// Backups:
//  - settings: all config keys except sheets schema (all_sheets/user_sheets) and rules
//  - transferRules: transfer_rules + transfer_templates_v2
//
// Key list is driven by a manifest ("__backup_manifest_v1") which is kept up-to-date
// automatically inside db.js (every cfgSet registers the key).

import { cfgGet, cfgSet, cfgGetBackupManifest } from "./db.js";

const APP = "dilovodstvo";

const TYPES = {
  SETTINGS: "settings",
  TRANSFER_RULES: "transferRules",
  SETTINGS_PARTIAL: "settingsPartial",
  TRANSFER_RULES_PARTIAL: "transferRulesPartial",
};

// Backup versions (increment when you change the on-disk backup structure)
const LATEST_VERSION = {
  [TYPES.SETTINGS]: 1,
  [TYPES.TRANSFER_RULES]: 1,
  [TYPES.SETTINGS_PARTIAL]: 1,
  [TYPES.TRANSFER_RULES_PARTIAL]: 1,
};

// Hard protection: never allow sheets/schema keys into any backup.
const DENY_EXACT = new Set(["__backup_manifest_v1"]);
const DENY_PREFIXES = [];

const RULE_KEYS = new Set(["transfer_rules", "transfer_templates_v2", "transferRules"]);
const SHEET_SCHEMA_KEYS = new Set(["all_sheets", "user_sheets"]);
const SHEET_SETTINGS_KEY = "sheet_settings";
const ADD_FIELDS_KEY = "add_fields";

// Logical option ids (used by UI)
const OPT = {
  SHEETS_SCHEMA: "sheets.schema",
  SHEET_HIDDEN: "sheetSettings.hiddenCols",
  SHEET_WIDTHS: "sheetSettings.widths",
  SHEET_EXPORT: "sheetSettings.export",
  ADD_FIELDS: "addFields",
  OTHER_SETTINGS: "otherSettings",
  TRANSFER_RULES: "transfer.rules",
};

// Tab → default option set
const TAB_DEFAULTS = {
  sheets: [OPT.SHEETS_SCHEMA],
  columns: [OPT.SHEETS_SCHEMA, OPT.SHEET_HIDDEN, OPT.SHEET_WIDTHS],
  export: [OPT.SHEET_EXPORT],
  addform: [OPT.ADD_FIELDS],
  transfer: [OPT.TRANSFER_RULES],
  // "all" means everything possible
  all: [OPT.SHEETS_SCHEMA, OPT.SHEET_HIDDEN, OPT.SHEET_WIDTHS, OPT.SHEET_EXPORT, OPT.ADD_FIELDS, OPT.OTHER_SETTINGS],
};

const OPT_LABELS = {
  [OPT.SHEETS_SCHEMA]: "Листи та колонки (назви, порядок, структура, default sort)",
  [OPT.SHEET_HIDDEN]: "Колонки: приховані/видимі (hiddenCols)",
  [OPT.SHEET_WIDTHS]: "Колонки: встановлена ширина (widths)",
  [OPT.SHEET_EXPORT]: "Експортні профілі на лист (pageSize/orientation/експортні фільтри)",
  [OPT.ADD_FIELDS]: "Поля для “Додати” (add_fields)",
  [OPT.OTHER_SETTINGS]: "Інші налаштування (все, що не входить у вкладки вище)",
  [OPT.TRANSFER_RULES]: "Перенесення (правила та шаблони)",
};

function isDeniedKey(key){
  if(!key) return true;
  if(DENY_EXACT.has(key)) return true;
  for(const p of DENY_PREFIXES){
    if(key.startsWith(p)) return true;
  }
  return false;
}

function safeParseJSON(text){
  try{ return JSON.parse(text); }catch{ return null; }
}

function downloadJSON(obj, filename){
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

async function getKeysByType(type){
  const m = await cfgGetBackupManifest();
  const out = [];
  const keysObj = (m && typeof m === "object" && m.keys) ? m.keys : {};
  for(const [k, meta] of Object.entries(keysObj)){
    const t = (meta && typeof meta === "object") ? meta.type : meta;
    if(t !== type) continue;
    if(isDeniedKey(k)) continue;
    out.push(k);
  }
  return out;
}

async function getAllSettingsKeys(){
  return await getKeysByType(TYPES.SETTINGS);
}

async function getOtherSettingsKeys(){
  const keys = await getAllSettingsKeys();
  return keys.filter(k => !SHEET_SCHEMA_KEYS.has(k) && k !== SHEET_SETTINGS_KEY && k !== ADD_FIELDS_KEY && !RULE_KEYS.has(k));
}

function uniq(arr){
  return Array.from(new Set(arr));
}

function pickSheetSettingsParts(sheetSettings, parts){
  // parts: { hiddenCols?:true, widths?:true, export?:true }
  const out = {};
  const src = (sheetSettings && typeof sheetSettings === "object") ? sheetSettings : {};
  for(const [sheetKey, cfg] of Object.entries(src)){
    if(!cfg || typeof cfg !== "object") continue;
    const dst = {};
    if(parts.hiddenCols) dst.hiddenCols = Array.isArray(cfg.hiddenCols) ? cfg.hiddenCols : [];
    if(parts.widths) dst.widths = (cfg.widths && typeof cfg.widths === "object") ? cfg.widths : {};
    if(parts.export) dst.export = (cfg.export && typeof cfg.export === "object") ? cfg.export : {};
    out[sheetKey] = dst;
  }
  return out;
}

function mergeSheetSettingsParts(existing, incoming, parts){
  const base = (existing && typeof existing === "object") ? existing : {};
  const inc = (incoming && typeof incoming === "object") ? incoming : {};
  const out = JSON.parse(JSON.stringify(base));
  for(const [sheetKey, patch] of Object.entries(inc)){
    if(!out[sheetKey] || typeof out[sheetKey] !== "object") out[sheetKey] = {};
    const tgt = out[sheetKey];
    const p = (patch && typeof patch === "object") ? patch : {};
    if(parts.hiddenCols && ("hiddenCols" in p)) tgt.hiddenCols = Array.isArray(p.hiddenCols) ? p.hiddenCols : [];
    if(parts.widths && ("widths" in p)) tgt.widths = (p.widths && typeof p.widths === "object") ? p.widths : {};
    if(parts.export && ("export" in p)) tgt.export = (p.export && typeof p.export === "object") ? p.export : {};
  }
  return out;
}

function selectionFromOptionIds(optionIds){
  const set = new Set(optionIds || []);
  return {
    includeSchema: set.has(OPT.SHEETS_SCHEMA),
    includeAddFields: set.has(OPT.ADD_FIELDS),
    includeOtherSettings: set.has(OPT.OTHER_SETTINGS),
    sheetSettingsParts: {
      hiddenCols: set.has(OPT.SHEET_HIDDEN),
      widths: set.has(OPT.SHEET_WIDTHS),
      export: set.has(OPT.SHEET_EXPORT),
    },
  };
}

export function getBackupOptionsForTab(tabId){
  const defaults = TAB_DEFAULTS[tabId] || [];
  // Transfer tab is special: options only for transfer rules.
  if(tabId === "transfer"){
    return [{ id: OPT.TRANSFER_RULES, label: OPT_LABELS[OPT.TRANSFER_RULES], checked: true }];
  }
  // For settings tabs
  const base = [OPT.SHEETS_SCHEMA, OPT.SHEET_HIDDEN, OPT.SHEET_WIDTHS, OPT.SHEET_EXPORT, OPT.ADD_FIELDS, OPT.OTHER_SETTINGS];
  return base.map(id => ({ id, label: OPT_LABELS[id], checked: defaults.includes(id) }));
}

export async function buildSettingsPartialBackup({ scope="tab", tabId="", optionIds=[] } = {}){
  const sel = selectionFromOptionIds(optionIds);
  const payload = {};

  if(sel.includeSchema){
    for(const k of SHEET_SCHEMA_KEYS){
      payload[k] = await cfgGet(k);
    }
  }

  const parts = sel.sheetSettingsParts || {};
  if(parts.hiddenCols || parts.widths || parts.export){
    const ss = await cfgGet(SHEET_SETTINGS_KEY);
    payload[SHEET_SETTINGS_KEY] = pickSheetSettingsParts(ss, parts);
  }

  if(sel.includeAddFields){
    payload[ADD_FIELDS_KEY] = await cfgGet(ADD_FIELDS_KEY);
  }

  if(sel.includeOtherSettings){
    const others = await getOtherSettingsKeys();
    for(const k of others){
      payload[k] = await cfgGet(k);
    }
  }

  return {
    app: APP,
    backupType: TYPES.SETTINGS_PARTIAL,
    backupVersion: LATEST_VERSION[TYPES.SETTINGS_PARTIAL],
    createdAt: new Date().toISOString(),
    scope,
    tabId,
    optionIds,
    payload,
  };
}

export async function buildTransferRulesPartialBackup(){
  const payload = {};
  for(const k of ["transfer_rules", "transfer_templates_v2"]){
    payload[k] = await cfgGet(k);
  }
  return {
    app: APP,
    backupType: TYPES.TRANSFER_RULES_PARTIAL,
    backupVersion: LATEST_VERSION[TYPES.TRANSFER_RULES_PARTIAL],
    createdAt: new Date().toISOString(),
    scope: "tab",
    tabId: "transfer",
    optionIds: [OPT.TRANSFER_RULES],
    payload,
  };
}

// ---- Migrations ----
// Hook for future upgrades. Each step must mutate `data` and increment data.backupVersion by 1.
const MIGRATIONS = {
  [TYPES.SETTINGS]: {},
  [TYPES.TRANSFER_RULES]: {},
  [TYPES.SETTINGS_PARTIAL]: {},
  [TYPES.TRANSFER_RULES_PARTIAL]: {},
};

function validateHeader(data, expectedType){
  if(!data || typeof data !== "object") throw new Error("Невірний формат бекапу");
  if(data.app !== APP) throw new Error("Це не бекап Dilovodstvo");
  if(data.backupType !== expectedType) throw new Error(`Очікував бекап типу "${expectedType}"`);
  if(typeof data.backupVersion !== "number") throw new Error("Немає backupVersion");
  if(data.backupVersion > LATEST_VERSION[expectedType]){
    throw new Error("Бекап створений новішою версією програми (потрібно оновитись)");
  }
  if(!data.payload || typeof data.payload !== "object") data.payload = {};
}

function migrateForward(type, data){
  const cur = data.backupVersion;
  const step = MIGRATIONS[type]?.[cur];
  if(!step) throw new Error(`Немає міграції для ${type} v${cur} -> v${cur+1}`);
  step(data);
  if(data.backupVersion !== cur + 1) throw new Error("Міграція некоректна (backupVersion не збільшено)");
}

export async function buildSettingsBackup(){
  const keys = await getKeysByType(TYPES.SETTINGS);
  const payload = {};
  for(const k of keys){
    payload[k] = await cfgGet(k);
  }
  return {
    app: APP,
    backupType: TYPES.SETTINGS,
    backupVersion: LATEST_VERSION[TYPES.SETTINGS],
    createdAt: new Date().toISOString(),
    payload,
  };
}

export async function buildTransferRulesBackup(){
  const keys = await getKeysByType(TYPES.TRANSFER_RULES);
  const payload = {};
  for(const k of keys){
    payload[k] = await cfgGet(k);
  }
  return {
    app: APP,
    backupType: TYPES.TRANSFER_RULES,
    backupVersion: LATEST_VERSION[TYPES.TRANSFER_RULES],
    createdAt: new Date().toISOString(),
    payload,
  };
}

export async function exportSettingsBackup(){
  const backup = await buildSettingsBackup();
  downloadJSON(backup, `dilovodstvo_settings_backup_${Date.now()}.json`);
}

export async function exportTransferRulesBackup(){
  const backup = await buildTransferRulesBackup();
  downloadJSON(backup, `dilovodstvo_transferRules_backup_${Date.now()}.json`);
}

export async function importBackupObject(data, expectedType){
  validateHeader(data, expectedType);
  const latest = LATEST_VERSION[expectedType];
  while(data.backupVersion < latest){
    migrateForward(expectedType, data);
  }

  // Write payload to config; keys will be auto-registered in manifest via cfgSet.
  // Partial settings backup merges sheet_settings sub-parts instead of replacing the whole object.
  const isSettingsPartial = expectedType === TYPES.SETTINGS_PARTIAL;
  const isRulesPartial = expectedType === TYPES.TRANSFER_RULES_PARTIAL;
  const optionIds = Array.isArray(data.optionIds) ? data.optionIds : [];
  const sel = selectionFromOptionIds(optionIds);
  const parts = sel.sheetSettingsParts || {};

  for(const [k, v] of Object.entries(data.payload || {})){
    if(isDeniedKey(k)) continue;

    // Enforce separation without relying on manifest (manifest may be stale).
    const isRuleKey = RULE_KEYS.has(k);

    // Enforce separation for full backups
    if(expectedType === TYPES.SETTINGS && isRuleKey) continue;
    if(expectedType === TYPES.TRANSFER_RULES && !isRuleKey) continue;

    // Enforce separation for partial backups
    if(isSettingsPartial && isRuleKey) continue;
    if(isRulesPartial && !isRuleKey) continue;

    // Merge sheet_settings parts for partial settings backup
    if(isSettingsPartial && k === SHEET_SETTINGS_KEY && v && typeof v === "object"){
      const existing = await cfgGet(SHEET_SETTINGS_KEY);
      const merged = mergeSheetSettingsParts(existing, v, parts);
      await cfgSet(SHEET_SETTINGS_KEY, merged);
      continue;
    }

    await cfgSet(k, v);
  }
}

export async function importBackupFromText(jsonText, expectedType){
  const parsed = safeParseJSON(jsonText);
  if(!parsed) throw new Error("Файл не є валідним JSON");
  await importBackupObject(parsed, expectedType);
}

export function pickAndImportFile(expectedType){
  const input = document.createElement("input");
  input.type = "file";
  input.accept = "application/json,.json";

  input.onchange = () => {
    const file = input.files && input.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = async () => {
      try{
        await importBackupFromText(String(reader.result || ""), expectedType);
        location.reload();
      }catch(e){
        alert(e?.message ? e.message : String(e));
      }
    };
    reader.readAsText(file);
  };

  input.click();
}

export async function exportSettingsPartialBackup({scope="tab", tabId="", optionIds=[]} = {}){
  const backup = await buildSettingsPartialBackup({scope, tabId, optionIds});
  downloadJSON(backup, `dilovodstvo_settings_${tabId || 'custom'}_backup_${Date.now()}.json`);
}

export async function exportTransferRulesPartialBackup(){
  const backup = await buildTransferRulesPartialBackup();
  downloadJSON(backup, `dilovodstvo_transferRules_backup_${Date.now()}.json`);
}

export const BackupCore = {
  TYPES,
  MIGRATIONS,
  getBackupOptionsForTab,
  importBackupFromText,
  exportSettingsBackup,
  exportTransferRulesBackup,
  exportSettingsPartialBackup,
  exportTransferRulesPartialBackup,
  pickAndImportSettingsBackup: () => pickAndImportFile(TYPES.SETTINGS),
  pickAndImportTransferRulesBackup: () => pickAndImportFile(TYPES.TRANSFER_RULES),
  pickAndImportSettingsPartialBackup: () => pickAndImportFile(TYPES.SETTINGS_PARTIAL),
  pickAndImportTransferRulesPartialBackup: () => pickAndImportFile(TYPES.TRANSFER_RULES_PARTIAL),
};

// convenient global for inline handlers / debugging
window.BackupCore = BackupCore;
