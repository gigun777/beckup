/**
 * TransferCore (DOM-free).
 * - Stores transfer templates in storage
 * - Applies a template to a source record to produce a target record
 * - Persists to tableStore datasets via provided adapters
 *
 * This file intentionally has NO DOM dependencies.
 */
import { canonicalJsonStringify } from '../utils/canonical_json.js';

const STORAGE_KEY = 'transfer:templates:v1';

function deepClone(x){ return x==null ? x : JSON.parse(JSON.stringify(x)); }
function ensureArray(x){ return Array.isArray(x) ? x : []; }

/**
 * @param {{storage:{get:(k:string)=>Promise<any>|any,set:(k:string,v:any)=>Promise<void>|void}}} deps
 */
export function createTransferCore(deps){
  const storage = deps?.storage;
  if(!storage) throw new Error('TransferCore requires deps.storage');

  async function loadTemplates(){
    const raw = await storage.get(STORAGE_KEY);
    return ensureArray(raw);
  }
  async function saveTemplates(templates){
    // stable-ish serialization (helps diffs/backups)
    const normalized = ensureArray(templates).map(t=>deepClone(t));
    await storage.set(STORAGE_KEY, JSON.parse(canonicalJsonStringify(normalized)));
  }

  /**
   * Apply template routes to srcRow array and return targetRow array.
   * Template format: { fromSheetKey,toSheetKey,routes:[{sources:number[],op,delimiter,targetCol:number}] }
   */
  /**
   * Apply template routes to a source row and produce a target row.
   *
   * Supports both index-based and key-based routes:
   * - sources: number[] (indexes) OR string[] (column keys)
   * - targetCol: number (index) OR string (column key)
   */
  function applyTemplateToRow(template, srcRow, opts){
    const t = template || {};
    const routes = ensureArray(t.routes);
    const optObj = Array.isArray(opts) ? { targetColKeys: opts } : (opts || {});
    const sourceColKeys = ensureArray(optObj.sourceColKeys);
    const targetColKeys = ensureArray(optObj.targetColKeys);

    const resolveSrcIdx = (v)=>{
      if(Number.isFinite(+v)) return (+v);
      if(typeof v === 'string' && sourceColKeys.length){
        const idx = sourceColKeys.indexOf(v);
        return idx >= 0 ? idx : NaN;
      }
      return NaN;
    };
    const resolveTgtIdx = (v)=>{
      if(Number.isFinite(+v)) return (+v);
      if(typeof v === 'string' && targetColKeys.length){
        const idx = targetColKeys.indexOf(v);
        return idx >= 0 ? idx : NaN;
      }
      return NaN;
    };
    const targetRow = [];
    for(const r of routes){
      const srcIdxs = ensureArray(r.sources).map(resolveSrcIdx).filter(n=>Number.isFinite(n) && n>=0);
      const srcVals = srcIdxs.length ? srcIdxs.map(i => srcRow?.[i]) : ensureArray(r.sources).map(i => srcRow?.[i]);
      let out = '';
      const op = r.op || 'concat';
      if(op === 'sum'){
        const nums = srcVals.map(v => {
          const n = typeof v === 'number' ? v : parseFloat(String(v ?? '').replace(',','.'));
          return Number.isFinite(n) ? n : 0;
        });
        out = nums.reduce((a,b)=>a+b,0);
      } else if(op === 'newline'){
        out = srcVals.map(v => v==null?'':String(v)).join('\n');
      } else if(op === 'seq'){
        out = srcVals.map(v => v==null?'':String(v)).join('');
      } else { // concat default
        const delim = r.delimiter ?? ' ';
        out = srcVals.map(v => v==null?'':String(v)).join(String(delim));
      }
      const keyOrIdx = (r.targetColKey ?? r.targetColId ?? r.targetCol);
      let idx = resolveTgtIdx(keyOrIdx);
      if(!Number.isFinite(idx)) idx = 0;
      if(idx < 0) idx = 0;
      targetRow[idx] = out;
    }
    return targetRow;
  }

  /**
   * Creates a new record for destination dataset from a targetRow array and column keys.
   * @param {string[]} targetColKeys
   */
  function buildRecordFromRow(targetColKeys, targetRow){
    const cells = {};
    for(let i=0;i<targetColKeys.length;i++){
      const k = targetColKeys[i];
      if(!k) continue;
      const v = targetRow?.[i];
      if(v !== undefined) cells[k] = v;
    }
    return { cells };
  }

  return {
    storageKey: STORAGE_KEY,
    loadTemplates,
    saveTemplates,
    applyTemplateToRow,
    buildRecordFromRow
  };
}
