'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { createGuiService } = require('../app/gui-service');

function deferred() {
  let resolve;
  const promise = new Promise(r => { resolve = r; });
  return { promise, resolve };
}
function fixture(overrides = {}) {
  const calls = [];
  const context = { close: async () => calls.push('close') };
  const service = createGuiService({
    initialize: async () => true,
    locations: () => ({ workbookDir: '/synthetic', logDir: '/synthetic/Logs', backupDir: '/synthetic/Backups' }),
    discover: async () => ({ valid: [{ fileName: 'A.xlsx' }, { fileName: 'B.xlsx' }], skipped: [] }),
    chooseFolder: async () => ({ changed: false }),
    openDirectory: async folder => calls.push(folder),
    openConnectWise: async () => { calls.push('open'); return context; },
    reconcile: async (_context, filename, options) => {
      calls.push(filename);
      options.onProgress({ stage: 'reconciling', current: 1, total: 2, projectID: 101 });
      return { workbookFilename: filename, status: 'COMPLETED', workbookWritten: true,
        summary: { rowsUpdated: 2, blankRowsFilled: 1, rowsInserted: 0 }, reviewCount: 0 };
    },
    ...overrides,
  });
  return { service, calls, context };
}

test('the UI requires login confirmation and processes selected workbooks sequentially', async () => {
  const gate = deferred(); const running = [];
  const f = fixture({ reconcile: async (_ctx, name, options) => {
    running.push(name); options.onProgress({ stage: 'reconciling', current: 1, total: 2 });
    if (name === 'A.xlsx') await gate.promise;
    return { workbookFilename: name, status: 'COMPLETED' };
  } });
  await f.service.initialize();
  await assert.rejects(f.service.continueRun(), /login/i);
  await f.service.start(['A.xlsx', 'B.xlsx']);
  assert.equal(f.service.snapshot().phase, 'awaiting_login');
  assert.deepEqual(running, []);
  await f.service.continueRun();
  await new Promise(setImmediate);
  assert.deepEqual(running, ['A.xlsx']);
  assert.equal(f.service.snapshot().progress.total, 2);
  await assert.rejects(f.service.start(['B.xlsx']), /progress/i);
  await assert.rejects(f.service.changeFolder(), /progress/i);
  await assert.rejects(f.service.refresh(), /progress/i);
  gate.resolve(); await f.service.waitForRun();
  assert.deepEqual(running, ['A.xlsx', 'B.xlsx']);
  assert.equal(f.service.snapshot().phase, 'complete');
  assert.equal(f.service.snapshot().results.length, 2);
  assert.equal(f.calls.at(-1), 'close');
});

test('only discovered workbook names can be selected and duplicate selection is rejected', async () => {
  const { service, calls } = fixture(); await service.initialize();
  for (const names of [[], ['../A.xlsx'], ['A.xlsx', 'A.xlsx'], ['Unknown.xlsx'], 'A.xlsx']) {
    await assert.rejects(service.start(names), /select/i);
  }
  assert.deepEqual(calls, []);
  await service.close();
});

test('a failed workbook retains its summary and the next workbook still runs', async () => {
  const { service } = fixture({ reconcile: async (_ctx, name) => {
    if (name === 'A.xlsx') throw Object.assign(new Error('Synthetic failure'), {
      auditLogPath: '/synthetic/Logs/A.csv', summaryLogPath: '/synthetic/Logs/A.txt',
      runResult: { workbookWritten: false, summary: { projectsFailed: 1 } },
    });
    return { workbookFilename: name, status: 'COMPLETED_WITH_REVIEW', reviewCount: 2 };
  } });
  await service.initialize(); await service.start(['A.xlsx', 'B.xlsx']);
  await service.continueRun(); await service.waitForRun();
  const state = service.snapshot();
  assert.equal(state.phase, 'complete');
  assert.equal(state.results[0].status, 'FAILED');
  assert.equal(state.results[0].error, 'Synthetic failure');
  assert.equal(state.results[0].summaryLogPath, '/synthetic/Logs/A.txt');
  assert.equal(state.results[1].status, 'COMPLETED_WITH_REVIEW');
});

test('closing during a write waits for the active workbook and selected batch to finish', async () => {
  const gate = deferred(); const f = fixture({ reconcile: async (_ctx, name) => {
    await gate.promise; f.calls.push(name); return { workbookFilename: name, status: 'COMPLETED' };
  } });
  await f.service.initialize(); await f.service.start(['A.xlsx']); await f.service.continueRun();
  let closed = false; const closing = f.service.close().then(() => { closed = true; });
  await new Promise(setImmediate); assert.equal(closed, false); assert(!f.calls.includes('close'));
  gate.resolve(); await closing;
  assert.deepEqual(f.calls, ['open', 'A.xlsx', 'close']);
});

test('cancelled folder selection preserves existing folder and workbook list', async () => {
  const { service } = fixture(); await service.initialize();
  const before = service.snapshot(); await service.changeFolder();
  assert.equal(service.snapshot().workbookDir, before.workbookDir);
  assert.deepEqual(service.snapshot().workbooks, before.workbooks);
  await service.close();
});

test('missing saved folder does not silently scan the application directory', async () => {
  let scanned = false;
  const { service } = fixture({ initialize: async () => false, discover: async () => { scanned = true; } });
  await service.initialize(); assert.equal(scanned, false);
  assert.equal(service.snapshot().workbookDir, null);
  assert.equal(service.snapshot().workbooks.length, 0);
});

test('login can be cancelled without reconciling and another run can be started', async () => {
  const f = fixture(); await f.service.initialize(); await f.service.start(['A.xlsx']);
  await f.service.cancelLogin(); assert.equal(f.service.snapshot().phase, 'ready');
  assert.deepEqual(f.calls, ['open', 'close']);
  await f.service.start(['B.xlsx']); await f.service.continueRun(); await f.service.waitForRun();
  assert.deepEqual(f.calls, ['open', 'close', 'open', 'B.xlsx', 'close']);
});

test('closing ConnectWise at the login gate gives a recoverable error without writing', async () => {
  const { EventEmitter } = require('node:events');
  const context = new EventEmitter(); context.close = async () => {};
  const f = fixture({ openConnectWise: async () => context });
  await f.service.initialize(); await f.service.start(['A.xlsx']); context.emit('close');
  assert.equal(f.service.snapshot().phase, 'error');
  assert.match(f.service.snapshot().error, /closed/i);
  await assert.rejects(f.service.continueRun(), /login/i);
  assert.deepEqual(f.calls, []); await f.service.close();
});

test('a failed folder change shows an error and preserves the previous folder and list', async () => {
  const f = fixture({ chooseFolder: async () => { throw new Error('Synthetic unavailable folder'); } });
  await f.service.initialize();
  await assert.rejects(f.service.changeFolder('/missing'), /unavailable/);
  assert.equal(f.service.snapshot().workbookDir, '/synthetic');
  assert.equal(f.service.snapshot().workbooks.length, 2);
  assert.match(f.service.snapshot().error, /unavailable/); await f.service.close();
});
