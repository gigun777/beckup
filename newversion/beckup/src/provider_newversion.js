import { exportFullJsonBackupFromSource, importFullJsonBackupToSource } from './json_io.js';
import { createNewversionSourceAdapter } from './source_newversion.js';

/**
 * Phase-4: provider factory for newversion createSEDO().backup.registerProvider(...)
 * DB-first by design: uses storage adapter as single source of truth.
 */
export function createBeckupProvider({
  storage,
  id = 'beckup-core',
  version = '0.4.0',
  include = { journals: true, settings: true, navigation: true, transfer: true }
} = {}) {
  if (!storage) throw new Error('storage is required');

  const adapter = createNewversionSourceAdapter(storage);

  return {
    id,
    version,
    describe() {
      return {
        settings: [
          'core_settings_v2',
          '@sdo/module-table-renderer:settings',
          '@sdo/module-table-renderer:settings:*'
        ],
        userData: [
          'tableStore:dataset:*',
          'spaces_nodes_v2',
          'journals_nodes_v2',
          'nav_last_loc_v2',
          'nav_history_v2',
          'transfer:templates:v1'
        ]
      };
    },
    async export({ scope } = {}) {
      return exportFullJsonBackupFromSource({ source: adapter, include, app: { name: '@beckup/beckup', version }, scope });
    },
    async import(payload, { mode = 'merge' } = {}) {
      const report = await importFullJsonBackupToSource(payload, { target: adapter, mode });
      return {
        applied: true,
        warnings: [
          ...report.journals.warnings,
          ...report.settings.warnings,
          ...report.navigation.warnings,
          ...report.transfer.warnings,
          ...report.journals.errors,
          ...report.settings.errors,
          ...report.navigation.errors,
          ...report.transfer.errors
          ...report.transfer.warnings
        ],
        report
      };
    }
  };
}
