'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { preparePackage, validatePackage } = require('../installer/package');

function fixture(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'commission-package-'));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  function write(name, data = 'Synthetic fixture') {
    const target = path.join(root, name);
    fs.mkdirSync(path.dirname(target), { recursive: true }); fs.writeFileSync(target, data);
  }
  for (const name of ['app.js', 'audit.js', 'commission-menu.js', 'connectwise.js', 'paths.js', 'settings.js', 'workbook.js', 'xlsx-writer.js',
    'log.js', 'gui.js', 'gui-service.js', 'gui-server.js', 'ui/index.html', 'ui/styles.css', 'ui/ui.js']) write(`app/${name}`, '// Synthetic source');
  write('runtime/node.exe'); write('runtime/LICENSE'); write('Run Commission Tool.cmd');
  write('installer/QUICK_START.txt');
  write('package.json', JSON.stringify({ name: 'synthetic', version: '1.1.0', dependencies: { playwright: '1.62.0', jszip: '3.10.1' } }));
  write('package-lock.json', '{}');
  return { root, write, stage: path.join(root, 'dist', 'package') };
}

function finishFakePackage(f) {
  f.write('dist/package/PHDTech Commission Tool.exe');
  for (const name of ['playwright/package.json', 'exceljs/package.json', 'jszip/package.json']) f.write(`dist/package/node_modules/${name}`, '{}');
  f.write('dist/package/browsers/synthetic-browser/chrome.exe');
}

test('staging copies only distribution files and excludes financial/authentication/developer data', t => {
  const f = fixture(t);
  for (const name of ['Employee.xlsx', 'Backups/Original.xlsx', 'Logs/Private.csv', 'connectwise-profile/Cookies',
    'settings.json', '.env', 'app/private-client.js', 'app/apply-workbook-changes.ps1', 'node_modules/private-marker']) f.write(name);
  const staged = preparePackage(f.root);
  assert.equal(staged, f.stage);
  assert.deepEqual(fs.readdirSync(f.stage).sort(), ['Quick Start.txt', 'Run Commission Tool.cmd', 'app', 'package-lock.json', 'package.json', 'runtime'].sort());
  assert(fs.existsSync(path.join(f.stage, 'app', 'xlsx-writer.js')));
  assert(!fs.existsSync(path.join(f.stage, 'app', 'private-client.js')));
  assert(!fs.existsSync(path.join(f.stage, 'app', 'apply-workbook-changes.ps1')));
  assert(!fs.existsSync(path.join(f.stage, 'node_modules')));
});

test('missing bundled runtime fails before deleting an earlier staging directory', t => {
  const f = fixture(t); f.write('dist/package/old-build-marker.txt', 'keep');
  fs.unlinkSync(path.join(f.root, 'runtime/node.exe'));
  assert.throws(() => preparePackage(f.root), /node.exe/);
  assert.equal(fs.readFileSync(path.join(f.stage, 'old-build-marker.txt'), 'utf8'), 'keep');
});

test('staging refuses to delete user settings left inside its output directory', t => {
  const f = fixture(t); f.write('dist/package/settings.json', 'Preserve this user file');
  assert.throws(() => preparePackage(f.root), /user data/i);
  assert.equal(fs.readFileSync(path.join(f.stage, 'settings.json'), 'utf8'), 'Preserve this user file');
});

test('a complete package requires runtime libraries and a bundled browser', t => {
  const f = fixture(t); preparePackage(f.root);
  assert.throws(() => validatePackage(f.stage), /missing/i);
  finishFakePackage(f);
  assert.doesNotThrow(() => validatePackage(f.stage));
  fs.rmSync(path.join(f.stage, 'browsers'), { recursive: true });
  assert.throws(() => validatePackage(f.stage), /browser/i);
});

test('package validation rejects workbooks and browser profiles even if later steps add them', t => {
  const f = fixture(t); preparePackage(f.root); finishFakePackage(f);
  f.write('dist/package/Employee.xlsx');
  assert.throws(() => validatePackage(f.stage), /user data/i);
  fs.unlinkSync(path.join(f.stage, 'Employee.xlsx'));
  f.write('dist/package/connectwise-profile/Cookies');
  assert.throws(() => validatePackage(f.stage), /user data/i);
});

test('Playwright must be a production dependency for offline installation', t => {
  const f = fixture(t); preparePackage(f.root); finishFakePackage(f);
  f.write('dist/package/package.json', JSON.stringify({ version: '1.1.0', dependencies: { jszip: '3.10.1' }, devDependencies: { playwright: '1.62.0' } }));
  assert.throws(() => validatePackage(f.stage), /production dependency/i);
});

test('staging refuses a linked output directory instead of removing files outside the project', t => {
  const f = fixture(t); const elsewhere = path.join(f.root, 'Keep this directory'); fs.mkdirSync(elsewhere);
  fs.writeFileSync(path.join(elsewhere, 'marker.txt'), 'keep');
  fs.symlinkSync(elsewhere, path.join(f.root, 'dist'), process.platform === 'win32' ? 'junction' : 'dir');
  assert.throws(() => preparePackage(f.root), /link/i);
  assert.equal(fs.readFileSync(path.join(elsewhere, 'marker.txt'), 'utf8'), 'keep');
});
