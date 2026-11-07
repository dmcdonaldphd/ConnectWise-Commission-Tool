'use strict';

const fs = require('node:fs');
const path = require('node:path');

const APP_FILES = ['app.js', 'audit.js', 'commission-menu.js', 'connectwise.js',
  'paths.js', 'settings.js', 'workbook.js', 'xlsx-writer.js'];
const ROOT_FILES = ['package.json', 'package-lock.json', 'Run Commission Tool.cmd', 'Quick Start.txt',
  'app', 'runtime', 'node_modules', 'browsers', 'build-info.json'];

function assertNoUserData(folder, root = folder) {
  if (!fs.existsSync(folder)) return;
  for (const entry of fs.readdirSync(folder, { withFileTypes: true })) {
    const name = entry.name.toLowerCase();
    const target = path.join(folder, entry.name);
    if (name.endsWith('.xlsx') || name === 'connectwise-profile' ||
        /\.commission-write-.*\.tmp$/.test(name) || /^commission-write-.*\.json$/.test(name) ||
        (folder === root && ['settings.json', 'backups', 'logs', '.env', '.npmrc'].includes(name))) {
      throw new Error(`Refusing to package or remove user data: ${target}`);
    }
    if (entry.isDirectory()) assertNoUserData(target, root);
  }
}

function requireFile(filename) {
  if (!fs.existsSync(filename) || !fs.lstatSync(filename).isFile()) throw new Error(`Missing distribution file: ${filename}`);
}

function preparePackage(projectRoot) {
  const root = path.resolve(projectRoot);
  const mappings = [
    ...APP_FILES.map(name => [`app/${name}`, `app/${name}`]),
    ['runtime/node.exe', 'runtime/node.exe'], ['runtime/LICENSE', 'runtime/LICENSE'],
    ['package.json', 'package.json'], ['package-lock.json', 'package-lock.json'],
    ['Run Commission Tool.cmd', 'Run Commission Tool.cmd'], ['installer/QUICK_START.txt', 'Quick Start.txt'],
  ];
  for (const [source] of mappings) requireFile(path.join(root, source));
  const dist = path.join(root, 'dist');
  const stage = path.join(dist, 'package');
  for (const directory of [dist, stage]) {
    if (fs.existsSync(directory) && fs.lstatSync(directory).isSymbolicLink()) throw new Error(`Refusing linked output directory: ${directory}`);
  }
  assertNoUserData(stage);
  fs.rmSync(stage, { recursive: true, force: true });
  for (const [source, destination] of mappings) {
    const target = path.join(stage, destination);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.copyFileSync(path.join(root, source), target);
  }
  return stage;
}

function validatePackage(stage) {
  assertNoUserData(stage);
  for (const entry of fs.readdirSync(stage)) {
    if (!ROOT_FILES.includes(entry)) throw new Error(`Unexpected distribution entry: ${entry}`);
  }
  for (const file of [...APP_FILES.map(name => `app/${name}`), 'runtime/node.exe', 'runtime/LICENSE',
    'Run Commission Tool.cmd', 'Quick Start.txt', 'package.json', 'package-lock.json',
    'node_modules/playwright/package.json', 'node_modules/exceljs/package.json', 'node_modules/jszip/package.json']) {
    requireFile(path.join(stage, file));
  }
  const metadata = JSON.parse(fs.readFileSync(path.join(stage, 'package.json'), 'utf8'));
  if (!metadata.dependencies?.playwright || !metadata.dependencies?.jszip) throw new Error('Playwright and JSZip must be a production dependency.');
  const browsers = path.join(stage, 'browsers');
  if (!fs.existsSync(browsers) || !fs.statSync(browsers).isDirectory() || !fs.readdirSync(browsers).length)
    throw new Error('Missing bundled browser.');
}

module.exports = { preparePackage, validatePackage };
