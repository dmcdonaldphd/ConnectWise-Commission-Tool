'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const paths = require('../app/paths');
const { createSettingsStore, initializeWorkbookFolder, chooseWorkbookFolder,
  pickWindowsFolder } = require('../app/settings');
const { selectWorkbooks, findCommissionWorkbooks } = require('../app/commission-menu');
const ExcelJS = require('exceljs');

function setup(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'commission-folders-'));
  const original = { ...paths };
  const workbooks = path.join(root, 'Workbooks with spaces — 2026');
  const state = path.join(root, 'User data');
  fs.mkdirSync(workbooks);
  const store = createSettingsStore(state);
  const messages = [];
  for (const method of ['log', 'warn', 'error']) t.mock.method(console, method, (...args) => messages.push(args.join(' ')));
  t.after(() => { Object.assign(paths, original); fs.rmSync(root, { recursive: true, force: true }); });
  return { root, workbooks, state, store, messages };
}

test('first launch selects a workbook folder and remembers it outside the application', async t => {
  const f = setup(t);
  let picked = 0;
  assert.equal(await initializeWorkbookFolder({ store: f.store, ask: async () => assert.fail('No text prompt expected'),
    pickFolder: async () => { picked++; return f.workbooks; } }), true);
  assert.equal(picked, 1);
  assert.equal(f.store.load().workbookDir, f.workbooks);
  assert.equal(paths.WORKBOOK_DIR, f.workbooks);
  assert.equal(paths.BACKUP_DIR, path.join(f.workbooks, 'Backups'));
  assert.equal(paths.LOG_DIR, path.join(f.workbooks, 'Logs'));
  assert.equal(paths.PROFILE_DIR, path.join(f.state, 'connectwise-profile'));
  assert.equal(paths.TEMP_DIR, path.join(f.state, 'Temp'));
  assert(fs.statSync(paths.TEMP_DIR).isDirectory());
});

test('future launches use the remembered folder without opening the picker', async t => {
  const f = setup(t);
  f.store.save(f.workbooks);
  assert.equal(await initializeWorkbookFolder({ store: f.store, ask: async () => assert.fail('No prompt expected'),
    pickFolder: async () => assert.fail('Picker must not reopen') }), true);
  assert.equal(paths.WORKBOOK_DIR, f.workbooks);
});

test('changing folders preserves browser state, prior workbooks, and old backups', async t => {
  const f = setup(t);
  f.store.save(f.workbooks);
  await initializeWorkbookFolder({ store: f.store });
  fs.mkdirSync(paths.PROFILE_DIR);
  fs.writeFileSync(path.join(paths.PROFILE_DIR, 'synthetic-profile-marker'), 'keep');
  fs.mkdirSync(paths.BACKUP_DIR);
  fs.writeFileSync(path.join(paths.BACKUP_DIR, 'synthetic-backup.txt'), 'keep');
  const second = path.join(f.root, 'Another folder'); fs.mkdirSync(second);
  assert.equal(await chooseWorkbookFolder({ store: f.store, pickFolder: async () => second }), true);
  assert.equal(paths.WORKBOOK_DIR, second);
  assert.equal(f.store.load().workbookDir, second);
  assert.equal(fs.readFileSync(path.join(paths.PROFILE_DIR, 'synthetic-profile-marker'), 'utf8'), 'keep');
  assert.equal(fs.readFileSync(path.join(f.workbooks, 'Backups', 'synthetic-backup.txt'), 'utf8'), 'keep');
});

test('canceling first selection exits cleanly without saving settings or creating runtime data', async t => {
  const f = setup(t);
  assert.equal(await initializeWorkbookFolder({ store: f.store, pickFolder: async () => null }), false);
  assert.equal(f.store.load(), null);
  assert(!fs.existsSync(f.state));
});

test('canceling a folder change leaves the existing selection and settings untouched', async t => {
  const f = setup(t);
  f.store.save(f.workbooks);
  await initializeWorkbookFolder({ store: f.store });
  const before = fs.readFileSync(f.store.filename);
  assert.equal(await chooseWorkbookFolder({ store: f.store, pickFolder: async () => null }), false);
  assert.equal(paths.WORKBOOK_DIR, f.workbooks);
  assert.deepEqual(fs.readFileSync(f.store.filename), before);
});

test('a missing remembered folder prompts for a replacement instead of silently using the application directory', async t => {
  const f = setup(t);
  f.store.save(f.workbooks); fs.rmdirSync(f.workbooks);
  const replacement = path.join(f.root, 'Replacement'); fs.mkdirSync(replacement);
  assert.equal(await initializeWorkbookFolder({ store: f.store, pickFolder: async () => replacement }), true);
  assert.equal(paths.WORKBOOK_DIR, replacement);
  assert(f.messages.some(message => message.includes('unavailable')));
});

test('corrupt settings can be recovered by an explicit new selection', async t => {
  const f = setup(t);
  fs.mkdirSync(f.state); fs.writeFileSync(f.store.filename, '{invalid json');
  assert.equal(await initializeWorkbookFolder({ store: f.store, pickFolder: async () => f.workbooks }), true);
  assert.equal(f.store.load().workbookDir, f.workbooks);
  assert(f.messages.some(message => message.includes('settings')));
});

test('invalid folder selection does not overwrite valid settings', async t => {
  const f = setup(t);
  f.store.save(f.workbooks);
  const before = fs.readFileSync(f.store.filename);
  assert.throws(() => f.store.save(path.join(f.root, 'Missing')), /folder/i);
  assert.deepEqual(fs.readFileSync(f.store.filename), before);
  assert.throws(() => f.store.save(f.store.filename), /folder/i);
});

test('failed runtime directory preparation leaves the previous selection and settings intact', async t => {
  const f = setup(t);
  f.store.save(f.workbooks);
  paths.configureDirectories(f.workbooks, f.state);
  fs.writeFileSync(path.join(f.state, 'Temp'), 'Synthetic conflicting file');
  const before = fs.readFileSync(f.store.filename);
  const replacement = path.join(f.root, 'Replacement'); fs.mkdirSync(replacement);
  assert.equal(await chooseWorkbookFolder({ store: f.store, pickFolder: async () => replacement }), false);
  assert.equal(paths.WORKBOOK_DIR, f.workbooks);
  assert.deepEqual(fs.readFileSync(f.store.filename), before);
});

test('a failed settings replacement preserves the old settings and removes temporary files', t => {
  const f = setup(t); f.store.save(f.workbooks);
  const before = fs.readFileSync(f.store.filename);
  const rename = fs.renameSync;
  t.mock.method(fs, 'renameSync', (source, target) => {
    if (target === f.store.filename) throw new Error('Synthetic settings write failure');
    return rename(source, target);
  });
  assert.throws(() => f.store.save(f.workbooks), /Synthetic settings write failure/);
  assert.deepEqual(fs.readFileSync(f.store.filename), before);
  assert(!fs.readdirSync(f.state).some(name => name.endsWith('.tmp')));
});

test('canceling recovery leaves an unsupported settings file intact', async t => {
  const f = setup(t); fs.mkdirSync(f.state);
  const raw = JSON.stringify({ version: 999, workbookDir: f.workbooks });
  fs.writeFileSync(f.store.filename, raw);
  assert.equal(await initializeWorkbookFolder({ store: f.store, pickFolder: async () => null }), false);
  assert.equal(fs.readFileSync(f.store.filename, 'utf8'), raw);
});

test('unavailable native folder picker offers a text-path fallback with spaces and quotes', async t => {
  const f = setup(t);
  assert.equal(await chooseWorkbookFolder({ store: f.store,
    pickFolder: async () => { throw new Error('Synthetic picker unavailable'); },
    ask: async () => `"${f.workbooks}"` }), true);
  assert.equal(paths.WORKBOOK_DIR, f.workbooks);
});

test('empty fallback input cancels without changing settings', async t => {
  const f = setup(t);
  assert.equal(await chooseWorkbookFolder({ store: f.store,
    pickFolder: async () => { throw new Error('Synthetic picker unavailable'); },
    ask: async () => '' }), false);
  assert.equal(f.store.load(), null);
});

test('Windows folder picker passes its initial path as data rather than executable PowerShell', () => {
  const initial = 'C:\\Synthetic $(&) folder';
  let call;
  const chosen = pickWindowsFolder(initial, (command, args, options) => {
    call = { command, args, options };
    return { status: 0, stdout: 'C:\\Selected folder\r\n' };
  });
  assert.equal(chosen, 'C:\\Selected folder');
  assert.equal(call.command, 'powershell.exe');
  assert.equal(call.options.env.PHD_COMMISSION_INITIAL_FOLDER, initial);
  assert(!call.args.join(' ').includes(initial));
  assert(!call.args.some(arg => /bypass/i.test(arg)));
});

test('Windows folder picker distinguishes cancellation from process failures', () => {
  assert.equal(pickWindowsFolder('', () => ({ status: 0, stdout: '' })), null);
  assert.throws(() => pickWindowsFolder('', () => ({ status: 1, stderr: 'Unavailable' })), /folder picker/i);
});

test('per-user data defaults are stable across application installation paths', () => {
  assert.equal(paths.getUserDataDirectory('win32', { LOCALAPPDATA: 'C:\\Users\\Synthetic\\AppData\\Local' }),
    'C:\\Users\\Synthetic\\AppData\\Local\\PHDTech\\Commission Tool');
  const base = path.join(os.tmpdir(), 'commission-test-user-data');
  assert.equal(paths.getUserDataDirectory(process.platform, { PHD_COMMISSION_DATA_DIR: base }), base);
});

test('menu preserves individual and process-all choices and makes folder selection visible', async t => {
  const f = setup(t);
  paths.configureDirectories(f.workbooks, f.state);
  const valid = [{ fileName: 'One.xlsx', paidHistorySheetNames: ['Paid to Emp 2026'] },
    { fileName: 'Two.xlsx', paidHistorySheetNames: ['Paid to Emp 2026'] }];
  const discover = async () => ({ valid, skipped: [] });
  assert.deepEqual(await selectWorkbooks({ ask: async () => '1', discover }), [valid[0]]);
  assert.deepEqual(await selectWorkbooks({ ask: async () => '3', discover }), valid);
  assert(f.messages.some(message => message.includes(f.workbooks)));
});

test('menu can change folders before selecting a workbook', async t => {
  const f = setup(t);
  let changed = 0, discoveries = 0;
  const answers = ['3', '1'];
  const valid = [{ fileName: 'One.xlsx', paidHistorySheetNames: ['Paid to Emp'] }];
  const result = await selectWorkbooks({ ask: async () => answers.shift(),
    discover: async () => { discoveries++; return { valid, skipped: [] }; },
    changeFolder: async () => { changed++; return true; } });
  assert.equal(changed, 1); assert.equal(discoveries, 2);
  assert.deepEqual(result, valid);
});

test('menu handles no eligible workbooks and cancellation without opening a browser', async t => {
  setup(t);
  const answers = ['1', '0'];
  let changed = 0;
  const result = await selectWorkbooks({ ask: async () => answers.shift(),
    discover: async () => ({ valid: [], skipped: [] }),
    changeFolder: async () => { changed++; return false; } });
  assert.equal(changed, 1); assert.equal(result, null);
});

test('menu opens selected workbook logs/backups and then allows normal selection', async t => {
  const f = setup(t); paths.configureDirectories(f.workbooks, f.state);
  const answers = ['4', '5', '1'];
  const opened = [];
  const valid = [{ fileName: 'One.xlsx', paidHistorySheetNames: ['Paid to Emp'] }];
  const result = await selectWorkbooks({ ask: async () => answers.shift(),
    discover: async () => ({ valid, skipped: [] }), openDirectory: folder => opened.push(folder) });
  assert.deepEqual(opened, [paths.LOG_DIR, paths.BACKUP_DIR]);
  assert.deepEqual(result, valid);
});

test('invalid menu choices retry safely rather than processing a workbook', async t => {
  setup(t);
  const answers = ['not a number', '99', '0'];
  assert.equal(await selectWorkbooks({ ask: async () => answers.shift(),
    discover: async () => ({ valid: [], skipped: [] }) }), null);
});

test('discovery reads the selected directory rather than workbooks in the application directory', async t => {
  const f = setup(t);
  paths.configureDirectories(f.workbooks, f.state);
  fs.writeFileSync(path.join(f.workbooks, 'Not a commission workbook.txt'), 'Synthetic');
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Commission').getCell('A1').value = 'Synthetic';
  workbook.addWorksheet('Paid to Emp 2026').getCell('A1').value = 'Synthetic';
  const buffer = await workbook.xlsx.writeBuffer();
  fs.writeFileSync(path.join(f.workbooks, 'Selected.xlsx'), buffer);
  fs.writeFileSync(path.join(f.workbooks, '~$Locked.xlsx'), buffer);
  fs.writeFileSync(path.join(f.workbooks, 'Ignored.working.xlsx'), buffer);
  fs.writeFileSync(path.join(f.workbooks, 'Selected - 2026-01-02_03-04-05.xlsx'), buffer);
  fs.writeFileSync(path.join(f.workbooks, 'Selected - 2026-01-02_03-04-05-1234abcd.xlsx'), buffer);
  assert.deepEqual(await findCommissionWorkbooks(), {
    valid: [{ fileName: 'Selected.xlsx', paidHistorySheetNames: ['Paid to Emp 2026'] }], skipped: [],
  });
});
