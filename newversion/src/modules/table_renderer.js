import { createTableEngine } from './table_engine.js';
import { formatCell as defaultFormatCell, parseInput as defaultParseInput } from './table_formatter.js';

function cellKey(rowId, colKey) {
  return `${rowId}:${colKey}`;
}

export function getRenderableCells(row, columns, cellSpanMap) {
  const cells = [];
  for (const column of columns) {
    const key = cellKey(row.rowId, column.columnKey);
    const span = cellSpanMap.get(key);
    if (span?.coveredBy) continue;
    cells.push({
      colKey: column.columnKey,
      span: span ?? { rowSpan: 1, colSpan: 1 }
    });
  }
  return cells;
}

function normalizeDataset(input = {}) {
  return {
    records: Array.isArray(input.records) ? input.records : [],
    merges: Array.isArray(input.merges) ? input.merges : []
  };
}

function updateDatasetWithPatch(dataset, patch) {
  return {
    ...dataset,
    records: dataset.records.map((record) => {
      if (record.id !== patch.recordId) return record;
      return {
        ...record,
        cells: { ...(record.cells ?? {}), ...(patch.cellsPatch ?? {}) },
        fmt: { ...(record.fmt ?? {}), ...(patch.fmtPatch ?? {}) }
      };
    })
  };
}

function applyColumnSettings(settings, nextColumns) {
  return {
    ...settings,
    columns: {
      ...(settings.columns ?? {}),
      ...nextColumns
    }
  };
}

function buildHeaderTitle(runtime) {
  const state = runtime?.sdo?.getState?.() ?? {};
  const journal = (state.journals ?? []).find((j) => j.id === state.activeJournalId);
  return journal ? `Таблиця: ${journal.title}` : 'Таблиця';
}

function parseSubrowId(subrowId) {
  const m = typeof subrowId === 'string' ? subrowId.match(/^(.*)::sub::(\d+)$/) : null;
  if (!m) return null;
  return { ownerId: m[1], index: Number(m[2]) };
}

function isSubrowsEnabled(settings, colKey) {
  // Enabled by default for all columns unless explicitly set to false.
  return settings?.subrows?.columnsSubrowsEnabled?.[colKey] !== false;
}

function createSubrowsUiAdapter({ engine, dataset }) {
  async function askCellAction() {
    if (!window.UI?.modal?.open) {
      const add = window.confirm('Додати нову підстроку для цієї клітинки? Натисніть Скасувати для редагування існуючої.');
      return add ? 'addSubrow' : 'editExisting';
    }

    return new Promise((resolve) => {
      const box = document.createElement('div');
      box.innerHTML = '<p style="margin:0 0 12px;">Оберіть дію для підстрок у цій клітинці.</p>';
      const controls = document.createElement('div');
      controls.style.display = 'flex';
      controls.style.gap = '8px';
      controls.style.justifyContent = 'flex-end';
      const editBtn = document.createElement('button');
      editBtn.textContent = 'Редагувати існуючу';
      const addBtn = document.createElement('button');
      addBtn.textContent = 'Додати підстроку';
      controls.append(editBtn, addBtn);
      box.append(controls);

      const modalId = window.UI.modal.open({ title: 'Підстроки', contentNode: box, closeOnOverlay: true });
      const close = (action) => {
        window.UI.modal.close(modalId);
        resolve(action);
      };
      editBtn.addEventListener('click', () => close('editExisting'));
      addBtn.addEventListener('click', () => close('addSubrow'));
    });
  }

  async function pickSubrow(opts) {
    const ids = (opts?.items ?? []).map((i) => i.id);
    if (ids.length === 0) return null;

    const labelOf = (id) => {
      try {
        return engine?.getSubrowLabel ? engine.getSubrowLabel(id, dataset) : id;
      } catch {
        return id;
      }
    };

    if (!window.UI?.modal?.open) {
      const choices = ids.map((id) => `${id} (${labelOf(id)})`).join('\n');
      const chosen = window.prompt(`Оберіть підстроку:\n${choices}`, ids[0]);
      return ids.includes(chosen) ? chosen : null;
    }

    return new Promise((resolve) => {
      const box = document.createElement('div');
      const hint = document.createElement('p');
      hint.textContent = 'Оберіть підстроку для редагування:';
      hint.style.margin = '0 0 12px';
      box.append(hint);

      const list = document.createElement('div');
      list.style.display = 'grid';
      list.style.gap = '8px';
      ids.forEach((id) => {
        const btn = document.createElement('button');
        btn.textContent = labelOf(id);
        btn.addEventListener('click', () => {
          window.UI.modal.close(modalId);
          resolve(id);
        });
        list.append(btn);
      });
      box.append(list);
      const modalId = window.UI.modal.open({ title: 'Підстроки', contentNode: box, closeOnOverlay: true });
    });
  }

  return {
    askCellAction,
    pickSubrow,
    toast(msg) {
      if (window.UI?.toast?.show) window.UI.toast.show(msg);
    }
  };
}

export function createTableRendererModule(opts = {}) {
  const {
    // legacy/fallback single-dataset key (used only when tableStore module is not present)
    datasetKey = '@sdo/module-table-renderer:dataset',
    settingsKey = '@sdo/module-table-renderer:settings'
  } = opts;
  const initialSettings = {
    columns: { order: null, visibility: {}, widths: {} },
    sort: null,
    filter: { global: '' },
    expandedRowIds: [],
    selectedRowIds: [],
    subrows: { columnsSubrowsEnabled: {} }
  };

  let engine = null;
  let currentSchemaId = null;
  let selectionMode = false;

  function schemaFromTemplate(template) {
    const cols = Array.isArray(template?.columns) ? template.columns : [];
    return {
      id: template?.id ? `tpl:${template.id}` : 'tpl:__none__',
      fields: cols.map((c) => ({ key: c.key, label: c.label, type: 'text' }))
    };
  }

  async function resolveSchema(runtime) {
    const state = runtime?.api?.getState ? runtime.api.getState() : (runtime?.sdo?.api?.getState ? runtime.sdo.api.getState() : null);
    const journalId = state?.activeJournalId;
    // Auto-select: if no active journal but there are journals in the active space, pick the first root journal.
    if (!journalId && state?.activeSpaceId && Array.isArray(state?.journals) && state.journals.length) {
      const candidate = state.journals.find((j) => j.spaceId === state.activeSpaceId && j.parentId === state.activeSpaceId);
      if (candidate && typeof runtime?.sdo?.commit === 'function') {
        await runtime.sdo.commit((next) => { next.activeJournalId = candidate.id; }, ['nav_last_loc_v2']);
        // refresh state snapshot after commit
        const st2 = runtime?.api?.getState ? runtime.api.getState() : (runtime?.sdo?.api?.getState ? runtime.sdo.api.getState() : null);
        const j2 = (st2?.journals ?? []).find((j) => j.id === st2?.activeJournalId);
        // continue resolving with the updated journal/state
        return await (async () => {
          const journal = j2;
          let templateId = journal?.templateId;
          const jt = runtime?.api?.journalTemplates || runtime?.sdo?.api?.journalTemplates || runtime?.sdo?.journalTemplates;
          if (!jt?.getTemplate) return { schema: { id: 'tpl:__none__', fields: [] }, journal, state: st2 };

          if (journal && !templateId) {
            const list = typeof jt.listTemplateEntities === 'function' ? await jt.listTemplateEntities() : [];
            const defaultTplId = (list.find((t) => t.id === 'test')?.id) || (list[0]?.id) || null;
            if (defaultTplId) {
              templateId = defaultTplId;
              await runtime.sdo.commit((next) => {
                next.journals = (next.journals ?? []).map((j) => (j.id === journal.id ? { ...j, templateId: defaultTplId } : j));
              }, ['journals_nodes_v2']);
            }
          }

          if (!templateId) return { schema: { id: 'tpl:__none__', fields: [] }, journal, state: st2 };
          const template = await jt.getTemplate(templateId);
          return { schema: schemaFromTemplate(template), journal, state: st2 };
        })();
      }
    }

    const journal = (state?.journals ?? []).find((j) => j.id === journalId);
    let templateId = journal?.templateId;

    const jt = runtime?.api?.journalTemplates || runtime?.sdo?.api?.journalTemplates || runtime?.sdo?.journalTemplates;
    if (!jt?.getTemplate) return { schema: { id: 'tpl:__none__', fields: [] }, journal, state };

    // Auto-heal: if journal exists but has no templateId, assign default (prefer "test")
    if (journal && !templateId) {
      const list = typeof jt.listTemplateEntities === 'function' ? await jt.listTemplateEntities() : [];
      const defaultTplId = (list.find((t) => t.id === 'test')?.id) || (list[0]?.id) || null;
      if (defaultTplId) {
        templateId = defaultTplId;
        // Persist into navigation state (best-effort)
        if (typeof runtime?.sdo?.commit === 'function') {
          await runtime.sdo.commit((next) => {
            next.journals = (next.journals ?? []).map((j) => (j.id === journal.id ? { ...j, templateId: defaultTplId } : j));
          }, ['journals_nodes_v2']);
        }
      }
    }

    if (!templateId) return { schema: { id: 'tpl:__none__', fields: [] }, journal, state };

    const template = await jt.getTemplate(templateId);
    return { schema: schemaFromTemplate(template), journal, state };
  }


  function tplSettingsKey(templateId) {
    return templateId ? `${settingsKey}:tpl:${templateId}` : null;
  }

  async function loadSettings(storage, templateId) {
    const globalSettings = { ...((await storage.get(settingsKey)) ?? {}) };
    const merged = { ...initialSettings, ...globalSettings };

    // Subrows enable/disable must be template-scoped (not global across all journals).
    // We store template-specific overrides under `${settingsKey}:tpl:<templateId>`.
    const tKey = tplSettingsKey(templateId);
    if (tKey) {
      const tplPart = (await storage.get(tKey)) ?? null;
      const tplMap = tplPart?.subrows?.columnsSubrowsEnabled ?? null;

      if (tplMap && typeof tplMap === 'object') {
        merged.subrows = { ...(merged.subrows ?? {}), columnsSubrowsEnabled: { ...(tplMap ?? {}) } };
      } else {
        // Migration: if template-scoped settings don't exist yet, seed them from the current global map.
        const legacyMap = merged.subrows?.columnsSubrowsEnabled;
        if (legacyMap && typeof legacyMap === 'object') {
          await storage.set(tKey, { subrows: { columnsSubrowsEnabled: { ...legacyMap } } });
        } else {
          await storage.set(tKey, { subrows: { columnsSubrowsEnabled: {} } });
        }
      }
    }
    return merged;
  }

  async function saveSettings(storage, templateId, settings) {
    // Persist non-template-scoped settings globally.
    // IMPORTANT: do not persist subrows.columnsSubrowsEnabled globally, otherwise it leaks across templates.
    const nextGlobal = { ...settings };
    if (nextGlobal?.subrows && typeof nextGlobal.subrows === 'object') {
      nextGlobal.subrows = { ...nextGlobal.subrows };
      delete nextGlobal.subrows.columnsSubrowsEnabled;
    }
    await storage.set(settingsKey, nextGlobal);

    // Persist template-scoped subrows map.
    const tKey = tplSettingsKey(templateId);
    if (tKey) {
      const map = settings?.subrows?.columnsSubrowsEnabled ?? {};
      await storage.set(tKey, { subrows: { columnsSubrowsEnabled: { ...map } } });
    }
  }

  async function loadDataset(runtime, storage, journalId) {
    const store = runtime?.api?.tableStore || runtime?.sdo?.api?.tableStore;
    if (store?.getDataset && journalId) {
      const ds = await store.getDataset(journalId);
      return normalizeDataset({ records: ds.records ?? [], merges: ds.merges ?? [] });
    }
    // fallback single-dataset storage
    return normalizeDataset((await storage.get(datasetKey)) ?? { records: [], merges: [] });
  }

  async function saveDataset(runtime, storage, journalId, dataset) {
    const store = runtime?.api?.tableStore || runtime?.sdo?.api?.tableStore;
    if (store?.upsertRecords && journalId) {
      // Replace records for now (renderer owns ordering)
      await store.upsertRecords(journalId, dataset.records ?? [], 'replace');
      return;
    }
    await storage.set(datasetKey, dataset);
  }

  function rerender(mount, runtime, renderFn) {
    mount.innerHTML = '';
    const cleanup = renderFn();
    if (typeof cleanup === 'function') return cleanup;
    return () => {};
    return () => {
          cleanupTableToolbar();};
  }

  function createModal() {
    const overlay = document.createElement('div');
    overlay.style.position = 'fixed';
    overlay.style.inset = '0';
    overlay.style.background = 'rgba(0,0,0,.35)';
    overlay.style.display = 'flex';
    overlay.style.alignItems = 'center';
    overlay.style.justifyContent = 'center';

    const modal = document.createElement('div');
    modal.style.background = '#fff';
    modal.style.padding = '12px';
    modal.style.borderRadius = '8px';
    modal.style.minWidth = '360px';

    overlay.append(modal);
    return { overlay, modal };
  }

  function columnSettingsUI(host, schema, settings, onChange) {
    const wrap = document.createElement('div');
    wrap.style.display = 'flex';
    wrap.style.gap = '8px';
    wrap.style.flexWrap = 'wrap';

    const schemaKeys = (schema && Array.isArray(schema.fields)) ? schema.fields.map((f) => f.key) : [];
    const ordered = (settings.columns && Array.isArray(settings.columns.order) && settings.columns.order.length)
      ? settings.columns.order
      : schemaKeys;

    for (const key of ordered) {
      const col = document.createElement('div');
      col.style.border = '1px solid #ddd';
      col.style.padding = '4px';

      const label = document.createElement('span');
      label.textContent = key;
      label.style.marginRight = '6px';

      const visible = document.createElement('input');
      visible.type = 'checkbox';
      visible.checked = settings.columns?.visibility?.[key] !== false;
      visible.addEventListener('change', () => {
        onChange(applyColumnSettings(settings, {
          visibility: { ...(settings.columns?.visibility ?? {}), [key]: visible.checked }
        }));
      });

      const subrows = document.createElement('input');
      subrows.type = 'checkbox';
      subrows.title = 'Підстроки';
      subrows.checked = settings.subrows?.columnsSubrowsEnabled?.[key] !== false;
      subrows.addEventListener('change', () => {
        onChange({
          ...settings,
          subrows: {
            ...(settings.subrows ?? { columnsSubrowsEnabled: {} }),
            columnsSubrowsEnabled: {
              ...((settings.subrows ?? {}).columnsSubrowsEnabled ?? {}),
              [key]: subrows.checked
            }
          }
        });
      });

      const widthInput = document.createElement('input');
      widthInput.type = 'number';
      widthInput.min = '40';
      widthInput.style.width = '72px';
      widthInput.value = settings.columns?.widths?.[key] ?? '';
      widthInput.addEventListener('change', () => {
        onChange(applyColumnSettings(settings, {
          widths: { ...(settings.columns?.widths ?? {}), [key]: Number(widthInput.value) || null }
        }));
      });

      const left = document.createElement('button');
      left.textContent = '←';
      left.addEventListener('click', () => {
        const idx = ordered.indexOf(key);
        if (idx <= 0) return;
        const nextOrder = [...ordered];
        [nextOrder[idx - 1], nextOrder[idx]] = [nextOrder[idx], nextOrder[idx - 1]];
        onChange(applyColumnSettings(settings, { order: nextOrder }));
      });

      const right = document.createElement('button');
      right.textContent = '→';
      right.addEventListener('click', () => {
        const idx = ordered.indexOf(key);
        if (idx < 0 || idx >= ordered.length - 1) return;
        const nextOrder = [...ordered];
        [nextOrder[idx], nextOrder[idx + 1]] = [nextOrder[idx + 1], nextOrder[idx]];
        onChange(applyColumnSettings(settings, { order: nextOrder }));
      });

      col.append(label, visible, subrows, widthInput, left, right);
      wrap.append(col);
    }

    host.append(wrap);
  }

  async function renderPanelFactory(mount, runtime) {
    function cleanupTableToolbar(){
      const host = document.querySelector('.sdo-table-toolbar-host');
      if (host) host.innerHTML = '';
    }

    let cleanup = () => {};

    const doRender = async () => {
      cleanup();
      cleanup = rerender(mount, runtime, () => {
        const container = document.createElement('div');
        const title = document.createElement('h4');
        title.textContent = buildHeaderTitle(runtime);

        const controls = document.createElement('div');
        controls.style.display = 'flex';
        controls.style.gap = '8px';

        const addBtn = document.createElement('button');
        addBtn.textContent = '+ Додати';

        const selectBtn = document.createElement('button');
        selectBtn.textContent = selectionMode ? 'Вибір: ON' : 'Вибір';

        const search = document.createElement('input');
        search.placeholder = 'Пошук';

        // Table must never cause horizontal scroll for the whole page.
        // Horizontal scroll is allowed ONLY inside the table module.
        const tableScroll = document.createElement('div');
        tableScroll.className = 'sdo-table-scroll';

        const table = document.createElement('table');
        table.className = 'sdo-table';
        // Fill the panel width by default; horizontal scroll stays inside tableScroll.
        // Column widths are controlled via <colgroup> so header/body always align.
        table.style.width = 'max-content';
        table.style.minWidth = '100%';
        table.style.borderCollapse = 'separate';
        table.style.borderSpacing = '0';
        tableScroll.append(table);

        container.className = 'sdo-table-panel';
        controls.className = 'sdo-table-controls';

        container.append(title, tableScroll);

        // Mount table controls into global header host (top bar)
        const headerHost = document.querySelector('.sdo-table-toolbar-host');
        if (headerHost) {
          headerHost.innerHTML = '';
          controls.classList.add('sdo-table-controls-inline');
          headerHost.append(controls);
        }
        controls.append(addBtn, selectBtn, search);
        mount.append(container);

        const listeners = [];

        // current journal id for dataset operations
        let currentJournalId = null;
        let currentTemplateId = null;

        const refreshTable = async () => {
          const resolved = await resolveSchema(runtime);
          const schema = resolved.schema;
          currentJournalId = resolved.state?.activeJournalId ?? null;
          currentTemplateId = resolved.journal?.templateId ?? null;

          const settings = await loadSettings(runtime.storage, currentTemplateId);
          const dataset = await loadDataset(runtime, runtime.storage, currentJournalId);
          if (!schema || !Array.isArray(schema.fields) || schema.fields.length === 0) {
            table.innerHTML = '';
            const msg = document.createElement('div');
            msg.style.padding = '8px';
            msg.style.color = '#666';
            msg.textContent = 'Немає колонок: журнал не має шаблону або шаблон не знайдено. Створіть журнал з шаблоном (наприклад, test).';
            table.append(msg);
            return;
          }

          // rebuild engine if schema changed
          if (!engine || currentSchemaId !== schema.id) {
            currentSchemaId = schema.id;
          }
          engine = createTableEngine({ schema, settings });
          engine.setDataset(dataset);
          const view = engine.compute();

          table.innerHTML = '';

          // One table:
          // - <thead> has 2 sticky rows (titles + column numbers)
          // - plus 2 fixed-width action columns on the far right (Transfer / Delete), like v1
          // - <colgroup> defines widths so header/body never drift.
          const colgroup = document.createElement('colgroup');
          const actionsColW = 44;
          const availableW = Math.max(320, tableScroll.getBoundingClientRect().width || 0);
          const nCols = view.columns.length;
          const baseW = Math.max(90, Math.floor((availableW - actionsColW * 2) / Math.max(1, nCols)));

          const thead = document.createElement('thead');
          const titleTr = document.createElement('tr');
          titleTr.className = 'sdo-hdr-title';
          const idxTr = document.createElement('tr');
          idxTr.className = 'sdo-hdr-idx';

          let colIdx = 0;
          for (const col of view.columns) {
            colIdx += 1;

            const w = col.width ? col.width : baseW;
            const colEl = document.createElement('col');
            colEl.style.width = `${w}px`;
            colEl.style.minWidth = `${w}px`;
            colgroup.append(colEl);

            const thTitle = document.createElement('th');
            // Column title + per-column subrows toggle
            const hdrWrap = document.createElement('div');
            hdrWrap.className = 'sdo-hdr-wrap';
            hdrWrap.style.display = 'flex';
            hdrWrap.style.alignItems = 'center';
            hdrWrap.style.justifyContent = 'center';
            hdrWrap.style.gap = '6px';

            const hdrLabel = document.createElement('span');
            hdrLabel.textContent = col.field?.label ?? col.columnKey;

            const toggle = document.createElement('button');
            toggle.type = 'button';
            toggle.className = 'sdo-hdr-subtoggle';
            toggle.style.width = '22px';
            toggle.style.height = '22px';
            toggle.style.borderRadius = '10px';
            toggle.style.border = '1px solid var(--border,#ddd)';
            toggle.style.background = 'var(--panel,#fff)';
            toggle.style.cursor = 'pointer';
            toggle.style.padding = '0';
            toggle.style.lineHeight = '1';
            toggle.style.fontSize = '12px';

            const isEnabled = settings?.subrows?.columnsSubrowsEnabled?.[col.columnKey] !== false;
            toggle.textContent = isEnabled ? '↳' : '×';
            toggle.title = isEnabled ? 'Підстроки: увімкнено (натисніть, щоб вимкнути)' : 'Підстроки: вимкнено (натисніть, щоб увімкнути)';
            toggle.dataset.colKey = col.columnKey;
            toggle.addEventListener('click', async (ev) => {
              ev.preventDefault();
              ev.stopPropagation();
              const colKey = toggle.dataset.colKey;
              const cur = settings?.subrows?.columnsSubrowsEnabled?.[colKey];
              const nextSettings = { ...settings, subrows: { ...(settings.subrows ?? {}), columnsSubrowsEnabled: { ...((settings.subrows ?? {}).columnsSubrowsEnabled ?? {}) } } };
              const enabledNow = cur !== false;
              nextSettings.subrows.columnsSubrowsEnabled[colKey] = enabledNow ? false : true;
              await saveSettings(runtime.storage, currentTemplateId, nextSettings);
              await refreshTable();
            });

            hdrWrap.append(hdrLabel, toggle);
            thTitle.append(hdrWrap);
            titleTr.append(thTitle);

            const thIdx = document.createElement('th');
            thIdx.className = 'sdo-col-idx';
            thIdx.textContent = String(colIdx);
            idxTr.append(thIdx);
          }

          const colTransfer = document.createElement('col');
          colTransfer.style.width = `${actionsColW}px`;
          colTransfer.style.minWidth = `${actionsColW}px`;
          const colDelete = document.createElement('col');
          colDelete.style.width = `${actionsColW}px`;
          colDelete.style.minWidth = `${actionsColW}px`;
          colgroup.append(colTransfer, colDelete);

          const thTransfer = document.createElement('th');
          thTransfer.className = 'sdo-col-actions';
          thTransfer.rowSpan = 2;
          thTransfer.title = 'Перенести';
          thTransfer.textContent = '⇄';

          const thDelete = document.createElement('th');
          thDelete.className = 'sdo-col-actions';
          thDelete.rowSpan = 2;
          thDelete.title = 'Видалити';
          thDelete.textContent = '🗑';

          titleTr.append(thTransfer, thDelete);

          thead.append(titleTr, idxTr);
          table.append(colgroup);
          table.append(thead);

          // Measure the 1st header row height and set CSS var so the 2nd row can sticky under it.
          // (Needed because row height can change with theme/font/2-line labels.)
          const syncHeaderHeights = () => {
            const h = titleTr.getBoundingClientRect().height;
            table.style.setProperty('--sdo-thead-row1-h', `${Math.ceil(h)}px`);
          };
          requestAnimationFrame(syncHeaderHeights);
          window.addEventListener('resize', syncHeaderHeights);

          const tbody = document.createElement('tbody');
          table.append(tbody);

          for (const row of view.rows) {
            const tr = document.createElement('tr');
// Cell-level subrows: subrows are rendered as additional "lines" inside enabled columns.
const rowSubrowIds = engine.listSubrows ? engine.listSubrows(row.rowId, dataset) : [];
const hasAnySubrowsCol = view.columns.some((c) => isSubrowsEnabled(settings, c.columnKey));
const rowLineCount = hasAnySubrowsCol ? (1 + rowSubrowIds.length) : 1;

            const renderableCells = getRenderableCells(row, view.columns, view.cellSpanMap);
            for (const cell of renderableCells) {
              const td = document.createElement('td');
              const span = cell.span;
              if (span.rowSpan) td.rowSpan = span.rowSpan;
              if (span.colSpan) td.colSpan = span.colSpan;

              const formatted = defaultFormatCell(row.record.cells?.[cell.colKey], row.record.fmt?.[cell.colKey] ?? {}, schema.fields.find((f) => f.key === cell.colKey) ?? {}, { locale: 'uk-UA', dateFormat: 'DD.MM.YYYY' });
              const firstColKey = view.columns[0]?.columnKey;
              const isFirstCol = cell.colKey === firstColKey;

// Cell content: for enabled columns we render stacked lines (base + subrows).
// For disabled columns we render a single block (no inner separators), but the cell still stretches
// to the row height automatically because the tallest cell defines the <tr> height.
td.innerHTML = '';
const stack = document.createElement('div');
const colAllowsSubrows = isSubrowsEnabled(settings, cell.colKey);

if (!colAllowsSubrows || rowLineCount <= 1) {
  stack.className = 'sdo-cell-stack sdo-cell-stack--single';
  const line = document.createElement('div');
  line.className = 'sdo-cell-line sdo-cell-line--single';
  line.dataset.rowId = row.rowId;

  // Base value
  const field = schema.fields.find((f) => f.key === cell.colKey) ?? {};
  const rec = (dataset.records ?? []).find((r) => r.id === row.rowId);
  const fmt = rec?.fmt?.[cell.colKey] ?? {};
  const val = rec?.cells?.[cell.colKey];
  const f2 = defaultFormatCell(val, fmt, field, { locale: 'uk-UA', dateFormat: 'DD.MM.YYYY' });
  line.textContent = f2.text ?? '';
  if (f2.align) line.style.textAlign = f2.align;
  if (f2.style) Object.assign(line.style, f2.style);

  stack.append(line);
  td.append(stack);
} else {
  stack.className = 'sdo-cell-stack';
  for (let li = 0; li < rowLineCount; li += 1) {
    const line = document.createElement('div');
    line.className = 'sdo-cell-line';
    let targetRowId = null;

    if (li === 0) {
      targetRowId = row.rowId;
    } else if (colAllowsSubrows) {
      targetRowId = rowSubrowIds[li - 1] ?? null;
    }

    if (targetRowId) {
      line.dataset.rowId = targetRowId;
      const p = parseSubrowId(targetRowId);
      const field = schema.fields.find((f) => f.key === cell.colKey) ?? {};
      let fmt = {};
      let val;
      if (p && p.ownerId === row.rowId) {
        const sub = Array.isArray(row.record.subrows) ? row.record.subrows[p.index] : null;
        val = sub?.cells?.[cell.colKey];
      } else {
        const rec = (dataset.records ?? []).find((r) => r.id === targetRowId);
        fmt = rec?.fmt?.[cell.colKey] ?? {};
        val = rec?.cells?.[cell.colKey];
      }
      const f2 = defaultFormatCell(val, fmt, field, { locale: 'uk-UA', dateFormat: 'DD.MM.YYYY' });
      line.textContent = f2.text ?? '';
      if (f2.align) line.style.textAlign = f2.align;
      if (f2.style) Object.assign(line.style, f2.style);
    } else {
      // Padding line (keeps height aligned across enabled columns)
      line.innerHTML = '&nbsp;';
      line.classList.add('sdo-cell-line--pad');
    }

    stack.append(line);
  }
  td.append(stack);
}

// Indentation only for the first (tree) column.
if (isFirstCol) {
  td.style.paddingLeft = `${row.depth * 16 + 8}px`;
}
if (rowLineCount > 1) td.classList.add('sdo-cell-has-subrows');

              // Actions are rendered as their own fixed-width columns at the far right (see below).

              if (formatted.align) td.style.textAlign = formatted.align;
              if (formatted.style) Object.assign(td.style, formatted.style);

              if (cell.colKey === view.columns[0]?.columnKey && row.hasChildren) {
                const expander = document.createElement('button');
                expander.textContent = row.isExpanded ? '▾' : '▸';
                expander.style.marginRight = '4px';
                expander.addEventListener('click', async (ev) => {
                  ev.stopPropagation();
                  engine.toggleExpand(row.rowId);
                  const next = { ...settings, expandedRowIds: [...engine.compute().rows.filter((r) => r.isExpanded).map((r) => r.rowId)] };
                  await saveSettings(runtime.storage, currentTemplateId, next);
                  await refreshTable();
                });
                td.prepend(expander);
              }

              td.addEventListener('click', async (ev) => {
  const spanInfo = view.cellSpanMap.get(cellKey(row.rowId, cell.colKey));
  if (spanInfo?.coveredBy) return;

  // Determine which "line" was clicked (base row or a specific subrow line).
  const lineEl = ev.target?.closest?.('.sdo-cell-line');
  const targetRowId = lineEl?.dataset?.rowId;
  const colAllowsSubrows = isSubrowsEnabled(settings, cell.colKey);

  // If user clicked on a padding line (no target), ignore.
  if (!targetRowId) return;

  // If subrows are disabled for this column, force edits to the base row.
  const editRowId = colAllowsSubrows ? targetRowId : row.rowId;

  const currentDataset = await loadDataset(runtime, runtime.storage, currentJournalId);
  const pEdit = parseSubrowId(editRowId);
  const baseRec = (currentDataset.records ?? []).find((r) => r.id === row.rowId) ?? row.record;

  let currentValue = '';
  if (pEdit && pEdit.ownerId === row.rowId) {
    const sub = Array.isArray(baseRec.subrows) ? baseRec.subrows[pEdit.index] : null;
    currentValue = sub?.cells?.[cell.colKey] ?? '';
  } else {
    const editRecord = (currentDataset.records ?? []).find((r) => r.id === editRowId) ?? row.record;
    currentValue = editRecord.cells?.[cell.colKey] ?? '';
  }

  // Inline edit (single input) — on save we patch the target record.
  engine.beginEdit(editRowId, cell.colKey);
  const field = schema.fields.find((f) => f.key === cell.colKey) ?? {};
  const inputModel = formatted.editor ?? { type: 'text', props: {} };
  const input = document.createElement('input');
  input.type = inputModel.type === 'number' ? 'number' : inputModel.type === 'date' ? 'date' : 'text';
  input.value = currentValue;

  td.innerHTML = '';
  td.append(input);
  input.focus();

  const save = async () => {
    const parsed = defaultParseInput(input.value, field);
    const dsNow = await loadDataset(runtime, runtime.storage, currentJournalId);

    const pEdit2 = parseSubrowId(editRowId);
    if (pEdit2 && pEdit2.ownerId === row.rowId && typeof engine.applySubrowEdit === 'function') {
      const { dataset: nextDataset } = engine.applySubrowEdit(editRowId, cell.colKey, parsed.v, dsNow);
      await saveDataset(runtime, runtime.storage, currentJournalId, nextDataset);
    } else {
      const patch = engine.applyEdit(editRowId, cell.colKey, parsed.v);
      const nextDataset = updateDatasetWithPatch(dsNow, patch);
      await saveDataset(runtime, runtime.storage, currentJournalId, nextDataset);
    }

    await refreshTable();
  };

  input.addEventListener('keydown', async (e2) => {
    if (e2.key === 'Enter') await save();
    if (e2.key === 'Escape') {
      engine.cancelEdit();
      await refreshTable();
    }
  });
  input.addEventListener('blur', save, { once: true });
});

              tr.append(td);
            }
// Fixed action columns at far right (Transfer / Context)
{
  const tdTransfer = document.createElement('td');
  tdTransfer.className = 'sdo-col-actions';
  const transferBtn = document.createElement('button');
  transferBtn.className = 'sdo-row-transfer';
  transferBtn.textContent = '⇄';
  transferBtn.title = 'Копіювати/перенести';
  transferBtn.addEventListener('click', (ev) => {
    ev.stopPropagation();
    runtime.sdo.commands.run('table.transferRow', { rowId: row.rowId });
  });
  tdTransfer.append(transferBtn);
  tr.append(tdTransfer);

  const tdCtx = document.createElement('td');
  tdCtx.className = 'sdo-col-actions';
  const ctxBtn = document.createElement('button');
  ctxBtn.className = 'sdo-row-context';
  ctxBtn.textContent = '☰';
  ctxBtn.title = 'Контекст';

  ctxBtn.addEventListener('click', async (ev) => {
    ev.stopPropagation();

    const content = document.createElement('div');
    content.className = 'ui-modal-content';

    const btnAddSub = document.createElement('button');
    btnAddSub.className = 'btn btn-primary';
    btnAddSub.textContent = 'Додати підстроку';

    const btnDelete = document.createElement('button');
    btnDelete.className = 'btn';
    btnDelete.textContent = 'Видалення…';

    const footer = document.createElement('div');
    footer.className = 'ui-modal-footer';
    const btnClose = document.createElement('button');
    btnClose.className = 'btn';
    btnClose.textContent = 'Закрити';

    footer.append(btnClose);

    content.append(btnAddSub);
    content.append(document.createElement('hr'));
    content.append(btnDelete);
    content.append(footer);

    const modalId = UI.modal.open({ title: 'Дії', contentNode: content, closeOnOverlay: true });

    btnClose.addEventListener('click', () => UI.modal.close(modalId));

    btnAddSub.addEventListener('click', async () => {
      UI.modal.close(modalId);

      const dsNow = await loadDataset(runtime, runtime.storage, currentJournalId);
      engine.setDataset(dsNow);

      // Create a subrow with empty cells only for columns where subrows are enabled.
      const initCells = {};
      for (const c of view.columns) {
        const k = c.columnKey;
        const enabled = settings?.subrows?.columnsSubrowsEnabled?.[k] !== false;
        if (enabled) initCells[k] = '';
      }

      const { dataset: nextDataset } = engine.addSubrow(row.rowId, { cells: initCells }, dsNow);
      await saveDataset(runtime, runtime.storage, currentJournalId, nextDataset);
      await refreshTable();
    });

    btnDelete.addEventListener('click', async () => {
      UI.modal.close(modalId);

      const node = document.createElement('div');
      node.className = 'ui-modal-content';

      const p = document.createElement('p');
      p.textContent = 'Видалити всю строку чи підстрочку?';
      node.append(p);

      const rowBtn = document.createElement('button');
      rowBtn.className = 'btn btn-danger';
      rowBtn.textContent = 'Всю строку';

      const subWrap = document.createElement('div');
      subWrap.style.display = 'flex';
      subWrap.style.gap = '8px';
      subWrap.style.alignItems = 'center';

      const subBtn = document.createElement('button');
      subBtn.className = 'btn btn-danger';
      subBtn.textContent = 'Підстрочку';

      const subInput = document.createElement('input');
      subInput.type = 'number';
      subInput.min = '1';
      subInput.placeholder = '№';
      subInput.style.width = '80px';

      subWrap.append(subBtn);
      subWrap.append(subInput);

      const actions = document.createElement('div');
      actions.className = 'ui-modal-footer';

      const cancelBtn = document.createElement('button');
      cancelBtn.className = 'btn';
      cancelBtn.textContent = 'Скасувати';

      actions.append(cancelBtn);

      node.append(rowBtn);
      node.append(document.createElement('hr'));
      node.append(subWrap);
      node.append(actions);

      const modalId2 = UI.modal.open({ title: 'Видалення', contentNode: node, closeOnOverlay: true });

      cancelBtn.addEventListener('click', () => UI.modal.close(modalId2));

      rowBtn.addEventListener('click', async () => {
        const ok = await UI.modal.confirm('Видалити всю строку?', { title: 'Підтвердження' });
        if (!ok) return;
        UI.modal.close(modalId2);

        const currentDataset2 = await loadDataset(runtime, runtime.storage, currentJournalId);
        const nextDataset2 = { ...currentDataset2, records: (currentDataset2.records ?? []).filter((r) => r.id !== row.rowId) };
        await saveDataset(runtime, runtime.storage, currentJournalId, nextDataset2);
        await refreshTable();
      });

      subBtn.addEventListener('click', async () => {
        const n = Number(subInput.value);
        if (!Number.isFinite(n) || n < 1) {
          subInput.focus();
          subInput.style.outline = '2px solid #d33';
          return;
        }
        const dsNow2 = await loadDataset(runtime, runtime.storage, currentJournalId);
        engine.setDataset(dsNow2);

        if (n === 1) {
          try { UI.toast?.show?.('Підстрочка №1 — це батьківська строка. Для її видалення оберіть "Всю строку".'); } catch {}
          return;
        }
        const subrowId = `${row.rowId}::sub::${n - 2}`;
        const { dataset: nextDataset2, removed } = engine.removeSubrow(subrowId, dsNow2);
        if (!removed) {
          subInput.style.outline = '2px solid #d33';
          return;
        }
        UI.modal.close(modalId2);
        await saveDataset(runtime, runtime.storage, currentJournalId, nextDataset2);
        await refreshTable();
      });
    });
  });

  tdCtx.append(ctxBtn);
  tr.append(tdCtx);

            if (selectionMode) {
              tr.style.cursor = 'pointer';
              tr.addEventListener('click', async () => {
                engine.toggleSelect(row.rowId);
                const next = { ...settings, selectedRowIds: [...engine.compute().selection] };
                await saveSettings(runtime.storage, currentTemplateId, next);
                await refreshTable();
              });
            }

            tbody.append(tr);
          }
        };
        };

        async function openAddRowFlow() {
          if (!engine) {
            await refreshTable();
            return;
          }

          const model = engine.getAddFormModel();

          if (window.UI?.modal?.open) {
            // Prefer native modal form (avoid chained browser prompts). If UI.form.create exists
            // use it, otherwise build a minimal form locally.
            const schema = model.map((f) => ({
              id: f.key,
              label: f.label,
              type: f.type || 'text',
              required: !!f.required,
              placeholder: f.placeholder || '',
              options: f.options || null
            }));
            let modalId;

            const onSubmit = async (values) => {
              const currentDataset = await loadDataset(runtime, runtime.storage, currentJournalId);
              const addResult = engine.addRow(values, currentDataset);
              if (!addResult.ok) {
                window.UI?.toast?.show?.('Не вдалося додати запис');
                return;
              }
              await saveDataset(runtime, runtime.storage, currentJournalId, addResult.dataset);
              window.UI.modal.close(modalId);
              await refreshTable();
            };

            const onCancel = () => window.UI.modal.close(modalId);

            let formNode;
            if (window.UI?.form?.create) {
              formNode = window.UI.form.create({ schema, onSubmit, onCancel });
            } else {
              // Minimal local form (scrollable)
              formNode = document.createElement('form');
              formNode.className = 'sdo-form sdo-form--add-row';
              formNode.style.maxHeight = '70vh';
              formNode.style.overflow = 'auto';

              const inputs = {};
              for (const field of schema) {
                const row = document.createElement('div');
                row.className = 'sdo-form__row';
                const label = document.createElement('label');
                label.textContent = field.label;
                const input = document.createElement('input');
                input.type = 'text';
                input.placeholder = field.placeholder || '';
                input.required = !!field.required;
                input.value = model.find(m => m.key === field.id)?.default ?? '';
                inputs[field.id] = input;
                row.appendChild(label);
                row.appendChild(input);
                formNode.appendChild(row);
              }

              const actions = document.createElement('div');
              actions.className = 'sdo-form__actions';
              const btnCancel = document.createElement('button');
              btnCancel.type = 'button';
              btnCancel.textContent = 'Скасувати';
              const btnOk = document.createElement('button');
              btnOk.type = 'submit';
              btnOk.textContent = 'Додати';
              actions.appendChild(btnCancel);
              actions.appendChild(btnOk);
              formNode.appendChild(actions);

              btnCancel.addEventListener('click', (e) => {
                e.preventDefault();
                onCancel();
              });

              formNode.addEventListener('submit', async (e) => {
                e.preventDefault();
                const values = {};
                for (const k of Object.keys(inputs)) values[k] = inputs[k].value ?? '';
                await onSubmit(values);
              });
            }

            modalId = window.UI.modal.open({
              title: 'Додати запис',
              contentNode: formNode,
              closeOnOverlay: true,
              escClose: true
            });
            return;
          }

          // Absolute fallback (no modal available)
          const first = model?.[0];
          const v = window.prompt(`Введіть ${first?.label ?? 'значення'}`, first?.default ?? '');
          if (v === null) return;
          const currentDataset = await loadDataset(runtime, runtime.storage, currentJournalId);
          const addResult = engine.addRow({ [first?.key ?? 'title']: v }, currentDataset);
          if (!addResult.ok) {
            window.UI?.toast?.show?.('Не вдалося додати запис');
            return;
          }
          await saveDataset(runtime, runtime.storage, currentJournalId, addResult.dataset);
          await refreshTable();
        }

        async function openAddSubrowFlow(ownerRowId, { insertAfterId } = {}) {
          if (!engine) {
            await refreshTable();
            return;
          }

          const model = engine.getAddFormModel();

          if (window.UI?.modal?.open) {
            const schema = model.map((f) => ({
              id: f.key,
              label: f.label,
              type: f.type || 'text',
              required: !!f.required,
              placeholder: f.placeholder || '',
              options: f.options || null
            }));

            let modalId;
            const onSubmit = async (values) => {
              const currentDataset = await loadDataset(runtime, runtime.storage, currentJournalId);
              const res = engine.addSubrow(ownerRowId, { insertAfterId, cells: values }, currentDataset);
              await saveDataset(runtime, runtime.storage, currentJournalId, res.dataset);
              window.UI.modal.close(modalId);
              await refreshTable();
            };
            const onCancel = () => window.UI.modal.close(modalId);

            let formNode;
            if (window.UI?.form?.create) {
              formNode = window.UI.form.create({ schema, onSubmit, onCancel });
            } else {
              formNode = document.createElement('form');
              formNode.className = 'sdo-form sdo-form--add-subrow';
              formNode.style.maxHeight = '70vh';
              formNode.style.overflow = 'auto';
              const inputs = {};
              for (const field of schema) {
                const row = document.createElement('div');
                row.className = 'sdo-form__row';
                const label = document.createElement('label');
                label.textContent = field.label;
                const input = document.createElement('input');
                input.type = 'text';
                input.placeholder = field.placeholder || '';
                input.required = !!field.required;
                inputs[field.id] = input;
                row.appendChild(label);
                row.appendChild(input);
                formNode.appendChild(row);
              }
              const actions = document.createElement('div');
              actions.className = 'sdo-form__actions';
              const btnCancel = document.createElement('button');
              btnCancel.type = 'button';
              btnCancel.textContent = 'Скасувати';
              const btnOk = document.createElement('button');
              btnOk.type = 'submit';
              btnOk.textContent = 'Додати';
              actions.append(btnCancel, btnOk);
              formNode.appendChild(actions);
              btnCancel.addEventListener('click', (e) => { e.preventDefault(); onCancel(); });
              formNode.addEventListener('submit', async (e) => {
                e.preventDefault();
                const values = {};
                for (const k of Object.keys(inputs)) values[k] = inputs[k].value ?? '';
                await onSubmit(values);
              });
            }

            modalId = window.UI.modal.open({
              title: 'Додати підстроку',
              contentNode: formNode,
              closeOnOverlay: true,
              escClose: true
            });
            return;
          }

          // Safe fallback: add an empty subrow without prompts.
          const currentDataset = await loadDataset(runtime, runtime.storage, currentJournalId);
          const res = engine.addSubrow(ownerRowId, { insertAfterId, cells: {} }, currentDataset);
          await saveDataset(runtime, runtime.storage, currentJournalId, res.dataset);
          await refreshTable();
        }

        addBtn.addEventListener('click', async () => {
          await openAddRowFlow();
        });


        selectBtn.addEventListener('click', async () => {
          selectionMode = !selectionMode;
          await refreshTable();
        });

        search.addEventListener('change', async () => {
          const settings = await loadSettings(runtime.storage, currentTemplateId);
          const next = { ...settings, filter: { ...(settings.filter ?? {}), global: search.value ?? '' } };
          await saveSettings(runtime.storage, currentTemplateId, next);
          await refreshTable();
        });

        refreshTable();

        return () => {
          cleanupTableToolbar();
          for (const [el, type, fn] of listeners) el.removeEventListener(type, fn);
        };
      });
    };

    doRender();
    const off = runtime.sdo.on('state:changed', doRender);
    return () => {
          cleanupTableToolbar();
      off?.();
      cleanup?.();
    };
  }

  return {
    id: '@sdo/module-table-renderer',
    version: '1.0.0',
    init(ctx) {
      ctx.registerCommands([
        {
          id: '@sdo/module-table-renderer.refresh',
          title: 'Refresh table renderer',
          run: async () => true
        },
        {
          id: '@sdo/module-table-renderer.toggle-selection-mode',
          title: 'Toggle table selection mode',
          run: async () => { selectionMode = !selectionMode; }
        },
        {
  id: 'table.transferRow',
  title: 'Transfer row',
  run: async (runtime, args = {}) => {
    const sourceJournalId = args.sourceJournalId ?? runtime?.api?.getState?.()?.activeJournalId;
    const rowId = args.rowId;
    if (!sourceJournalId || !rowId) return false;

    const tr = globalThis.UI?.transfer;
    if (!tr?.openRowModal) {
      globalThis.UI?.toast?.warning?.('Transfer UI не готовий');
      return false;
    }

    // Include subrows automatically (engine-native).
    let recordIds = [rowId];
    try {
      const resolved = await resolveSchema(runtime);
      const tplId = resolved.journal?.templateId ?? null;
      const settings = await loadSettings(runtime.storage, tplId);
      const schema = resolved.schema;
      const dataset = await loadDataset(runtime, runtime.storage, sourceJournalId);
      const eng = createTableEngine({ schema, settings });
      eng.setDataset(dataset);
      if (typeof eng.getTransferCandidates === 'function') {
        const cand = eng.getTransferCandidates(rowId);
        if (Array.isArray(cand) && cand.length) recordIds = cand;
      }
    } catch (e) {
      // fallback to single row
    }

    await tr.openRowModal({ sourceJournalId, recordIds });
    return true;
  }
}
      ]);

      ctx.ui.registerButton({
        id: '@sdo/module-table-renderer:add-row',
        label: '+ Додати',
        location: 'toolbar',
        order: 30,
        onClick: () => ctx.commands.run('@sdo/module-table-renderer.refresh')
      });

      ctx.ui.registerButton({
        id: '@sdo/module-table-renderer:selection',
        label: 'Вибір',
        location: 'toolbar',
        order: 31,
        onClick: () => ctx.commands.run('@sdo/module-table-renderer.toggle-selection-mode')
      });

      ctx.ui.registerPanel({
        id: '@sdo/module-table-renderer:panel',
        title: 'Table',
        location: 'main',
        order: 5,
        render: (mount, runtime) => {
          if (typeof document === 'undefined') return () => {};
          if (!runtime?.storage) runtime.storage = ctx.storage;
          if (!runtime?.sdo) runtime.sdo = runtime?.api?.sdo;
          return renderPanelFactory(mount, runtime);
        }
      });
    }
  };
}