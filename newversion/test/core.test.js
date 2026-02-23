import test from 'node:test';
import assert from 'node:assert/strict';

import { computeNumbering, canGoBackJournal, canGoBackSpace, currentJournalLabel, plusCreatesOnlyChildren } from '../src/core/navigation_core.js';
import { deleteSpaceSubtree } from '../src/core/spaces_tree_core.js';
import { deleteJournalSubtree } from '../src/core/journal_tree_core.js';
import { createSEDO, createMemoryStorage, createNavi } from '../src/index.js';

test('numbering supports 1, 1.2, 1.2.3 pattern fragments', () => {
  const numbering = computeNumbering(['1', '2', '3']);
  assert.deepEqual(numbering, ['1', '1.2', '1.2.3']);
});

test('back and plus rules', () => {
  assert.equal(canGoBackSpace({ parentId: null }), false);
  assert.equal(canGoBackSpace({ parentId: 'root' }), true);
  assert.equal(canGoBackJournal({ parentId: 'space-1' }, 'space-1'), false);
  assert.equal(canGoBackJournal({ parentId: 'journal-1' }, 'space-1'), true);
  assert.equal(plusCreatesOnlyChildren(), true);
});

test('delete subtree works for spaces and journals', () => {
  const spaces = [
    { id: 'r', parentId: null, childCount: 2 },
    { id: 'a', parentId: 'r', childCount: 0 },
    { id: 'b', parentId: 'r', childCount: 1 },
    { id: 'c', parentId: 'b', childCount: 0 }
  ];
  const result = deleteSpaceSubtree(spaces, 'b');
  assert.deepEqual([...result.removedIds].sort(), ['b', 'c']);

  const journals = [
    { id: 'j1', parentId: 'space-1', childCount: 1 },
    { id: 'j2', parentId: 'j1', childCount: 0 }
  ];
  const jRes = deleteJournalSubtree(journals, 'j1');
  assert.deepEqual([...jRes.removedIds].sort(), ['j1', 'j2']);
});

test('empty space shows add journal label', () => {
  const label = currentJournalLabel({ journals: [], activeSpaceId: 's1', activeJournalId: null });
  assert.equal(label, 'Додай журнал');
});

test('integration: create space/journal then delete with commit pipeline', async () => {
  const sdo = createSEDO({ storage: createMemoryStorage() });
  await sdo.start();

  await sdo.commit((state) => {
    state.spaces = [{ id: 's-root', title: 'Root', parentId: null, childCount: 0 }];
    state.activeSpaceId = 's-root';
  }, ['spaces_nodes_v2']);

  await sdo.commit((state) => {
    state.journals = [{ id: 'j-root', spaceId: 's-root', parentId: 's-root', templateId: 'base', title: 'J1', childCount: 0 }];
    state.activeJournalId = 'j-root';
  }, ['journals_nodes_v2']);

  let snapshot = sdo.getState();
  assert.equal(snapshot.activeSpaceId, 's-root');
  assert.equal(snapshot.activeJournalId, 'j-root');

  await sdo.commit((state) => {
    state.journals = [];
    state.activeJournalId = null;
  }, ['journals_nodes_v2']);

  snapshot = sdo.getState();
  assert.equal(snapshot.journals.length, 0);
  assert.equal(snapshot.activeJournalId, null);
});

test('createNavi uses v2 payload contract', async () => {
  const storage = createMemoryStorage();
  const navi = createNavi(storage);

  await navi.importNavigationState({
    spaces_nodes_v2: [{ id: 's', title: 'S', parentId: null, childCount: 0 }],
    journals_nodes_v2: [],
    nav_last_loc_v2: { activeSpaceId: 's', activeJournalId: null },
    nav_history_v2: []
  });

  const exported = await navi.exportNavigationState();
  assert.equal(exported.spaces_nodes_v2[0].id, 's');
  assert.equal(exported.nav_last_loc_v2.activeSpaceId, 's');
});

test('importBackup fails on invalid integrity hash', async () => {
  const sdo = createSEDO({ storage: createMemoryStorage() });
  await sdo.start();

  const bundle = await sdo.exportBackup();
  bundle.integrity.payloadHashB64 = 'invalid';

  await assert.rejects(() => sdo.importBackup(bundle), /integrity check failed/);
});

test('module can register ordered UI buttons and panels', async () => {
  const sdo = createSEDO({ storage: createMemoryStorage() });
  sdo.use({
    id: 'ui-module',
    version: '1.0.0',
    init(ctx) {
      ctx.ui.registerButton({ id: 'b2', label: 'B2', location: 'toolbar', order: 20, onClick() {} });
      ctx.ui.registerButton({ id: 'b1', label: 'B1', location: 'toolbar', order: 10, onClick() {} });
      ctx.ui.registerPanel({ id: 'p1', title: 'P1', location: 'settings', order: 5, render() {} });
    }
  });

  const buttons = sdo.ui.listButtons({ location: 'toolbar' });
  assert.deepEqual(buttons.map((b) => b.id), ['b1', 'b2']);
  const panels = sdo.ui.listPanels({ location: 'settings' });
  assert.equal(panels[0].id, 'p1');
});

test('destroy clears module UI registrations via unregister', async () => {
  const sdo = createSEDO({ storage: createMemoryStorage() });
  sdo.use({
    id: 'temp-module',
    version: '1.0.0',
    init(ctx) {
      ctx.ui.registerButton({ id: 'temp:b', label: 'Tmp', location: 'toolbar', onClick() {} });
    }
  });

  assert.equal(sdo.ui.listButtons().length, 1);
  await sdo.destroy();
  assert.equal(sdo.ui.listButtons().length, 0);
});

test('ui registry enforces unique ids', () => {
  const sdo = createSEDO({ storage: createMemoryStorage() });
  assert.throws(() => {
    sdo.use({
      id: 'dup-module',
      version: '1.0.0',
      init(ctx) {
        ctx.ui.registerButton({ id: 'dup', label: 'One', location: 'toolbar', onClick() {} });
        ctx.ui.registerButton({ id: 'dup', label: 'Two', location: 'toolbar', onClick() {} });
      }
    });
  }, /already registered/);
});

test('schema registry works with get/list/resolve and version guard', () => {
  const sdo = createSEDO({ storage: createMemoryStorage() });
  sdo.use({
    id: 'schema-mod',
    version: '1.0.0',
    init(ctx) {
      ctx.registerSchema({
        id: 'schema.a',
        version: '1.0.0',
        domain: 'table',
        appliesTo: { templateId: 'tpl-1' },
        fields: [{ key: 'name', label: 'Name', type: 'text' }]
      });
    }
  });

  assert.equal(sdo.schemas.get('schema.a').domain, 'table');
  assert.equal(sdo.schemas.resolve('tpl-1').id, 'schema.a');
  assert.equal(sdo.schemas.list().length, 1);

  assert.throws(() => {
    sdo.use({
      id: 'schema-mod-2',
      version: '1.0.0',
      init(ctx) {
        ctx.registerSchema({
          id: 'schema.a',
          version: '1.0.0',
          domain: 'table',
          appliesTo: { any: true },
          fields: []
        });
      }
    });
  }, /same or newer version/);
});

test('commands registry run/list supports when guard', async () => {
  const sdo = createSEDO({ storage: createMemoryStorage() });
  let ran = false;

  sdo.use({
    id: 'cmd-mod',
    version: '1.0.0',
    init(ctx) {
      ctx.registerCommands([
        {
          id: 'cmd.ok',
          title: 'OK',
          menu: { location: 'toolbar' },
          run: async () => { ran = true; }
        },
        {
          id: 'cmd.blocked',
          title: 'Blocked',
          when: () => false,
          run: async () => {}
        }
      ]);
    }
  });

  assert.equal(sdo.commands.list().length, 2);
  await sdo.commands.run('cmd.ok');
  assert.equal(ran, true);
  await assert.rejects(() => sdo.commands.run('cmd.blocked'), /disabled by when/);
});

test('settings registry exposes tabs and read/write via storage namespace', async () => {
  const sdo = createSEDO({ storage: createMemoryStorage() });

  sdo.use({
    id: 'settings-mod',
    version: '1.0.0',
    init(ctx) {
      ctx.registerSettings({
        id: 'settings-mod.settings',
        tab: { id: 'settings-mod', title: 'Settings Mod', order: 1 },
        fields: [
          {
            key: 'settings-mod:viewMode',
            label: 'View mode',
            type: 'text',
            read: async () => (await ctx.storage.get('settings-mod:viewMode')) ?? 'table',
            write: async (_runtime, value) => ctx.storage.set('settings-mod:viewMode', value)
          }
        ]
      });
    }
  });

  const tabs = sdo.settings.listTabs();
  assert.equal(tabs[0].id, 'settings-mod');
  const field = tabs[0].items[0].fields[0];
  await field.write({}, 'cards');
  assert.equal(await field.read({}), 'cards');
});

test('journal templates container initializes default template and supports add/delete', async () => {
  const sdo = createSEDO({ storage: createMemoryStorage() });
  await sdo.start();

  const initial = await sdo.journalTemplates.listTemplates();
  assert.ok(initial.some((x) => x.id === 'test'));

  await sdo.journalTemplates.addTemplate({
    id: 'custom',
    title: 'Custom',
    columns: [{ key: 'c1', label: 'A' }]
  });

  const added = await sdo.journalTemplates.getTemplate('custom');
  assert.equal(added.title, 'Custom');

  await sdo.journalTemplates.deleteTemplate('custom');
  const removed = await sdo.journalTemplates.getTemplate('custom');
  assert.equal(removed, null);
});
