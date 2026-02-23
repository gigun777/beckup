import { createSheetJsonPayload } from './sheet_json.js';
import { normalizeBackupBundle, normalizeJournalPayload } from './schema_guard.js';

/**
 * Export all backup data from source adapters (single source of truth = DB/storage).
 */
export async function exportFullJsonBackupFromSource({
  app = { name: '@beckup/beckup', version: '0.1.0' },
  source,
  include = { journals: true, settings: true, navigation: true, transfer: true }
} = {}) {
  if (!source) throw new Error('source adapter is required');

  const out = {
    format: 'beckup-full-json',
    formatVersion: 1,
    createdAt: new Date().toISOString(),
    app,
    sections: {}
  };

  if (include.journals) {
    const journals = await (source.listJournals?.() || []);
    const journalPayloads = [];
    for (const journal of journals || []) {
      const records = await source.loadJournalRecords?.(journal.id, journal);
      const schema = await source.loadJournalSchema?.(journal.id, journal);
      const sheet = {
        key: journal.key || journal.id,
        title: journal.title || journal.name || journal.id,
        columns: schema?.columns || schema?.fields || []
      };
      journalPayloads.push(createSheetJsonPayload({
        sheet,
        records: Array.isArray(records) ? records : [],
        exportProfile: await source.loadJournalExportProfile?.(journal.id, journal) || null
      }));
    }
    out.sections.journals = { count: journalPayloads.length, items: journalPayloads };
  }

  if (include.settings) {
    out.sections.settings = {
      payload: await source.loadSettings?.() || {}
    };
  }

  if (include.navigation) {
    out.sections.navigation = {
      payload: await source.loadNavigation?.() || null
    };
  }

  if (include.transfer) {
    out.sections.transfer = {
      payload: await source.loadTransfer?.() || {}
    };
  }

  return out;
}

/**
 * Import backup with partial tolerance:
 * - if some sections are missing/corrupt, journals can still be restored.
 */
export async function importFullJsonBackupToSource(payload, { target, mode = 'merge' } = {}) {
  if (!target) throw new Error('target adapter is required');

  const normalized = normalizeBackupBundle(payload);

  const report = {
    journals: { applied: 0, skipped: 0, warnings: [], errors: [] },
    settings: { applied: false, warnings: [], errors: [] },
    navigation: { applied: false, warnings: [], errors: [] },
    transfer: { applied: false, warnings: [], errors: [] },
    meta: { format: normalized.format, formatVersion: normalized.formatVersion }
  };

  const journals = normalized?.sections?.journals?.items;
  if (Array.isArray(journals)) {
    for (const j of journals) {
      try {
        const checked = normalizeJournalPayload(j);
        if (!checked.ok) {
          report.journals.skipped += 1;
          report.journals.warnings.push(`Skipped journal: ${checked.reason}`);
          continue;
        }
        const journalKey = checked.payload.meta.key;
        await target.saveJournalPayload?.(journalKey, checked.payload, { mode });
        report.journals.applied += 1;
      } catch (e) {
        report.journals.errors.push(`Journal import error: ${e?.message || String(e)}`);
  if (!payload || typeof payload !== 'object') throw new Error('Invalid backup payload');
  if (!target) throw new Error('target adapter is required');

  const report = {
    journals: { applied: 0, warnings: [] },
    settings: { applied: false, warnings: [] },
    navigation: { applied: false, warnings: [] },
    transfer: { applied: false, warnings: [] }
  };

  const journals = payload?.sections?.journals?.items;
  if (Array.isArray(journals)) {
    for (const j of journals) {
      try {
        if (j?.meta?.type !== 'journal') {
          report.journals.warnings.push('Skipped non-journal item');
          continue;
        }
        const journalKey = j?.meta?.key || j?.sheet?.key;
        if (!journalKey) {
          report.journals.warnings.push('Skipped journal without key');
          continue;
        }
        await target.saveJournalPayload?.(journalKey, j, { mode });
        report.journals.applied += 1;
      } catch (e) {
        report.journals.warnings.push(`Journal import warning: ${e?.message || String(e)}`);
      }
    }
  }

  try {
    if (payload?.sections?.settings) {
      await target.saveSettings?.(payload.sections.settings.payload || {}, { mode });
      report.settings.applied = true;
    }
  } catch (e) {
    report.settings.errors.push(e?.message || String(e));
    report.settings.warnings.push(e?.message || String(e));
  }

  try {
    if (payload?.sections?.navigation) {
      await target.saveNavigation?.(payload.sections.navigation.payload || null, { mode });
      report.navigation.applied = true;
    }
  } catch (e) {
    report.navigation.errors.push(e?.message || String(e));
    report.navigation.warnings.push(e?.message || String(e));
  }

  try {
    if (payload?.sections?.transfer) {
      await target.saveTransfer?.(payload.sections.transfer.payload || {}, { mode });
      report.transfer.applied = true;
    }
  } catch (e) {
    report.transfer.errors.push(e?.message || String(e));
    report.transfer.warnings.push(e?.message || String(e));
  }

  return report;
}
