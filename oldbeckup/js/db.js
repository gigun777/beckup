// js/db.js
const DB_NAME = "dilovodstvoDB_modular";
const DB_VERSION = 1;
export let db = null;

// ---- Backup manifest (registry of config keys) ----
// Config in this app lives in IndexedDB store "config". To make backups automatic and
// future-proof, we keep a manifest of all config keys that exist/are written.
// New features that add new cfgSet("some_key", ...) will automatically be included
// in backups (unless explicitly excluded).
const BACKUP_MANIFEST_KEY = "__backup_manifest_v1";
const BACKUP_TYPES = {
  SETTINGS: "settings",
  TRANSFER_RULES: "transferRules",
  EXCLUDED: "excluded",
};
const BACKUP_DENY_EXACT = new Set([BACKUP_MANIFEST_KEY]);
const BACKUP_DENY_PREFIXES = [];
const BACKUP_TRANSFER_RULE_KEYS = new Set([
  "transfer_rules",
  "transfer_templates_v2",
]);

function _isDeniedCfgKey(key){
  if(!key) return true;
  if(key === BACKUP_MANIFEST_KEY) return true;
  if(BACKUP_DENY_EXACT.has(key)) return true;
  return BACKUP_DENY_PREFIXES.some(p=>key.startsWith(p));
}

function _classifyCfgKey(key){
  if(_isDeniedCfgKey(key)) return BACKUP_TYPES.EXCLUDED;
  if(BACKUP_TRANSFER_RULE_KEYS.has(key)) return BACKUP_TYPES.TRANSFER_RULES;
  return BACKUP_TYPES.SETTINGS;
}

function _cfgStore(mode="readonly"){
  return db.transaction("config", mode).objectStore("config");
}

async function _backupLoadManifest(){
  return new Promise((resolve)=>{
    try {
      const req = _cfgStore("readonly").get(BACKUP_MANIFEST_KEY);
      req.onsuccess = ()=>{
        const v = req.result?.value;
        if(v && typeof v === "object" && typeof v.keys === "object") resolve(v);
        else resolve({ version: 1, keys: {} });
      };
      req.onerror = ()=>resolve({ version: 1, keys: {} });
    } catch {
      resolve({ version: 1, keys: {} });
    }
  });
}

async function _backupSaveManifest(manifest){
  return new Promise((resolve)=>{
    try {
      const req = _cfgStore("readwrite").put({ key: BACKUP_MANIFEST_KEY, value: manifest });
      req.onsuccess = ()=>resolve(true);
      req.onerror = ()=>resolve(false);
    } catch {
      resolve(false);
    }
  });
}

async function _backupEnsureManifestScanned(){
  // Called once after DB open: scans all existing config keys and records them.
  const manifest = await _backupLoadManifest();

  const keys = await new Promise((resolve)=>{
    try {
      const req = _cfgStore("readonly").getAllKeys();
      req.onsuccess = ()=>resolve(req.result || []);
      req.onerror = ()=>resolve([]);
    } catch {
      resolve([]);
    }
  });

  let changed = false;
  for(const k of keys){
    if(!k || k === BACKUP_MANIFEST_KEY) continue;
    if(!manifest.keys[k]){
      manifest.keys[k] = { type: _classifyCfgKey(k) };
      changed = true;
    } else {
      // Keep manifest in sync with current classification rules.
      const t = _classifyCfgKey(k);
      if(manifest.keys[k].type !== t){
        manifest.keys[k].type = t;
        changed = true;
      }
    }
  }

  // Ensure manifest exists even when no keys yet
  if(!manifest.keys) manifest.keys = {};
  await _backupSaveManifest(manifest);
  return true;
}

async function _backupRegisterCfgKey(key){
  if(!key || key === BACKUP_MANIFEST_KEY) return;
  const k = String(key);
  const t = _classifyCfgKey(k);
  // We still record excluded keys in manifest, but they won't be exported.
  const manifest = await _backupLoadManifest();
  if(!manifest.keys) manifest.keys = {};
  if(!manifest.keys[k] || manifest.keys[k].type !== t){
    manifest.keys[k] = { type: t };
    await _backupSaveManifest(manifest);
  }
}

export async function cfgGetBackupManifest(){
  // Public helper for backup core.
  return await _backupLoadManifest();
}

export async function openDB(){
  return new Promise((resolve, reject)=>{
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e)=>{
      const _db = e.target.result;
      if(!_db.objectStoreNames.contains("config")){
        _db.createObjectStore("config", { keyPath:"key" });
      }
      if(!_db.objectStoreNames.contains("rows")){
        const st = _db.createObjectStore("rows", { keyPath:"id", autoIncrement:true });
        st.createIndex("by_journalKey","journalKey",{unique:false});
      }
      if(!_db.objectStoreNames.contains("cases")){
        _db.createObjectStore("cases", { keyPath:"id", autoIncrement:true });
      }
      if(!_db.objectStoreNames.contains("case_rows")){
        const st=_db.createObjectStore("case_rows", { keyPath:"id", autoIncrement:true });
        st.createIndex("by_caseId","caseId",{unique:false});
      }
    };
    req.onsuccess=async()=>{
      db=req.result;
      // Build/update backup manifest once DB is ready.
      try{ await _backupEnsureManifestScanned(); }catch{ /* ignore */ }
      resolve(db);
    };
    req.onerror=()=>reject(req.error);
  });
}
function store(name, mode="readonly"){
  return db.transaction(name, mode).objectStore(name);
}
export function cfgGet(key){
  return new Promise((resolve,reject)=>{
    const req=store("config").get(key);
    req.onsuccess=()=>resolve(req.result ? req.result.value : null);
    req.onerror=()=>reject(req.error);
  });
}
export function cfgSet(key, value){
  return new Promise((resolve,reject)=>{
    const req=store("config","readwrite").put({key,value});
    req.onsuccess=()=>{
      // Register key for automatic backup inclusion.
      _backupRegisterCfgKey(key).catch(()=>{});
      resolve(true);
    };
    req.onerror=()=>reject(req.error);
  });
}
export function getRows(journalKey){
  return new Promise((resolve,reject)=>{
    const st=store("rows");
    const idx=st.index("by_journalKey");
    const req=idx.getAll(journalKey);
    req.onsuccess=()=>resolve(req.result||[]);
    req.onerror=()=>reject(req.error);
  });
}
export function addRow(journalKey, row){
  return new Promise((resolve,reject)=>{
    const st=store("rows","readwrite");
    const rec={ journalKey, ...row };
    const req=st.add(rec);
    req.onsuccess=()=>resolve(req.result);
    req.onerror=()=>reject(req.error);
  });
}
export function putRow(row){
  return new Promise((resolve,reject)=>{
    const req=store("rows","readwrite").put(row);
    req.onsuccess=()=>resolve(true);
    req.onerror=()=>reject(req.error);
  });
}
export function deleteRow(id){
  return new Promise((resolve,reject)=>{
    const req=store("rows","readwrite").delete(id);
    req.onsuccess=()=>resolve(true);
    req.onerror=()=>reject(req.error);
  });
}
export function clearRows(journalKey){
  return getRows(journalKey).then(async rows=>{
    for(const r of rows) await deleteRow(r.id);
    return true;
  });
}
export async function clearAllRows(){
  return new Promise((resolve,reject)=>{
    const req=store("rows","readwrite").clear();
    req.onsuccess=()=>resolve(true);
    req.onerror=()=>reject(req.error);
  });
}
export function getAllCases(){
  return new Promise((resolve,reject)=>{
    const req=store("cases").getAll();
    req.onsuccess=()=>resolve(req.result||[]);
    req.onerror=()=>reject(req.error);
  });
}
export function addCase(obj){
  return new Promise((resolve,reject)=>{
    const req=store("cases","readwrite").add(obj);
    req.onsuccess=()=>resolve(req.result);
    req.onerror=()=>reject(req.error);
  });
}
export function getCaseRows(caseId){
  return new Promise((resolve,reject)=>{
    const st=store("case_rows");
    const idx=st.index("by_caseId");
    const req=idx.getAll(caseId);
    req.onsuccess=()=>resolve(req.result||[]);
    req.onerror=()=>reject(req.error);
  });
}
export function addCaseRow(caseId, row){
  return new Promise((resolve,reject)=>{
    const req=store("case_rows","readwrite").add({caseId, ...row});
    req.onsuccess=()=>resolve(req.result);
    req.onerror=()=>reject(req.error);
  });
}
export function putCaseRow(row){
  return new Promise((resolve,reject)=>{
    const req=store("case_rows","readwrite").put(row);
    req.onsuccess=()=>resolve(true);
    req.onerror=()=>reject(req.error);
  });
}
export function deleteCaseRow(id){
  return new Promise((resolve,reject)=>{
    const req=store("case_rows","readwrite").delete(id);
    req.onsuccess=()=>resolve(true);
    req.onerror=()=>reject(req.error);
  });
}
export async function clearAllCasesAndRows(){
  await new Promise((resolve,reject)=>{
    const req=store("cases","readwrite").clear();
    req.onsuccess=()=>resolve(true);
    req.onerror=()=>reject(req.error);
  });
  await new Promise((resolve,reject)=>{
    const req=store("case_rows","readwrite").clear();
    req.onsuccess=()=>resolve(true);
    req.onerror=()=>reject(req.error);
  });
}
