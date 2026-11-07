'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const vm = require('node:vm');
const { createRequire } = require('node:module');
const { EventEmitter } = require('node:events');
const paths = require('../app/paths');
const { selectWorkbooks } = require('../app/commission-menu');

// Exercise the real opener with a Windows process adapter, on either CI OS.
// No Explorer windows, browser sessions, or actual employee files are needed.
function setup(t, { error, throwOnSpawn = false } = {}) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'commission-open-folder-'));
  const workbooks = path.join(root, 'Workbooks with spaces — 2026');
  fs.mkdirSync(workbooks);
  const originalPaths = { ...paths };
  paths.configureDirectories(workbooks, path.join(root, 'User data'));
  const calls = [], messages = [];
  for (const method of ['log', 'warn', 'error']) t.mock.method(console, method, (...args) => messages.push(args.join(' ')));
  t.after(() => { Object.assign(paths, originalPaths); fs.rmSync(root, { recursive: true, force: true }); });
  const filename = require.resolve('../app/settings');
  const localRequire = createRequire(filename);
  const processAdapter = {
    spawnSync(command, args, options) {
      calls.push({ command, args, options, synchronous: true });
      return error ? { error, status: null } : { status: 0 };
    },
    spawn(command, args, options) {
      if (throwOnSpawn) throw error;
      const child = new EventEmitter();
      const call = { command, args, options, unreferenced: false };
      calls.push(call);
      child.unref = () => { call.unreferenced = true; };
      queueMicrotask(() => child.emit(error ? 'error' : 'spawn', error));
      return child;
    },
  };
  const context = { module: { exports: {} }, console,
    process: { platform: 'win32', env: {} },
    require: name => name === 'node:child_process' ? processAdapter : localRequire(name) };
  vm.runInNewContext(fs.readFileSync(filename, 'utf8'), context, { filename });
  return { root, workbooks, calls, messages, openFolder: context.module.exports.openFolder };
}

test('Windows folder commands request a visible independent Explorer window', async t => {
  const f = setup(t);
  await f.openFolder(paths.LOG_DIR);
  assert.equal(f.calls.length, 1);
  assert.equal(f.calls[0].command, 'explorer.exe');
  assert.equal(f.calls[0].options.windowsHide, false);
  assert.equal(f.calls[0].options.detached, true);
  assert.equal(f.calls[0].unreferenced, true);
  assert(f.messages.some(message => message.includes(paths.LOG_DIR)));
});

test('folder paths with spaces and shell characters remain a single data argument', async t => {
  const f = setup(t);
  const folder = path.join(f.workbooks, 'Logs & monthly [records] 100%');
  await f.openFolder(folder);
  assert.equal(f.calls[0].args.length, 1);
  assert.equal(f.calls[0].args[0], folder);
  assert.equal(Boolean(f.calls[0].options.shell), false);
  assert(fs.statSync(folder).isDirectory());
});

test('Explorer launch failures include the folder and underlying Windows error', async t => {
  const f = setup(t, { error: new Error('Synthetic Windows policy denied launch') });
  await assert.rejects(Promise.resolve().then(() => f.openFolder(paths.LOG_DIR)), error =>
    error.message.includes(paths.LOG_DIR) && error.message.includes('policy denied launch'));
});

test('synchronous Explorer startup errors are also reported', async t => {
  const f = setup(t, { error: new Error('Synthetic process creation failure'), throwOnSpawn: true });
  await assert.rejects(Promise.resolve().then(() => f.openFolder(paths.BACKUP_DIR)), /process creation failure/);
});

test('a file occupying the Logs location is reported instead of opening it as a folder', async t => {
  const f = setup(t);
  fs.writeFileSync(paths.LOG_DIR, 'Synthetic conflicting file');
  await assert.rejects(Promise.resolve().then(() => f.openFolder(paths.LOG_DIR)), /not a folder/);
  assert.equal(f.calls.length, 0);
});

test('an unavailable workbook parent is not silently recreated by a folder shortcut', async t => {
  const f = setup(t);
  fs.rmdirSync(f.workbooks);
  await assert.rejects(Promise.resolve().then(() => f.openFolder(paths.LOG_DIR)), /unavailable/);
  assert.equal(fs.existsSync(f.workbooks), false);
  assert.equal(f.calls.length, 0);
});

test('both menu shortcuts reach the visible Windows opener and return to the menu', async t => {
  const f = setup(t);
  const answers = ['4', '5', '0'];
  const result = await selectWorkbooks({ ask: async () => answers.shift(),
    discover: async () => ({ valid: [{ fileName: 'Synthetic.xlsx', paidHistorySheetNames: ['Paid to Emp'] }], skipped: [] }),
    openDirectory: f.openFolder });
  assert.equal(result, null);
  assert.deepEqual(f.calls.map(call => call.args[0]), [paths.LOG_DIR, paths.BACKUP_DIR]);
  assert(f.calls.every(call => call.options.windowsHide === false && call.unreferenced));
});
