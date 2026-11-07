'use strict';

const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const assert = require('node:assert/strict');
const { createRequire } = require('node:module');
const { spawnSync } = require('node:child_process');

async function main() {
  const appDir = path.resolve(process.argv[2]);
  assert.equal(process.platform, 'win32');
  assert.equal(process.arch, 'x64');
  assert(Number(process.versions.node.split('.')[0]) >= 24, 'Bundled Node must support the validated Node 24 baseline');
  process.env.PLAYWRIGHT_BROWSERS_PATH = path.join(appDir, 'browsers');
  const requireApp = createRequire(path.join(appDir, 'package.json'));
  for (const name of ['exceljs', 'jszip', 'playwright']) {
    const relative = path.relative(path.join(appDir, 'node_modules'), requireApp.resolve(name));
    assert(relative && !relative.startsWith('..') && !path.isAbsolute(relative), `${name} must be bundled, not resolved from the workstation`);
  }
  const ExcelJS = requireApp('exceljs');
  const { chromium } = requireApp('playwright');
  const { COMMISSION_REQUIRED_HEADERS, loadWorkbook } = requireApp('./app/workbook');
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'commission-installed-smoke-'));
  let browser;
  try {
    const headers = COMMISSION_REQUIRED_HEADERS;
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Commission');
    const values = { '%': 0.1, 'Company': 'Synthetic company', 'Project ID': 101,
      'Project/Ticket': 'Synthetic project', 'Invoice #': 'SYNTHETIC-1',
      'Comm': { formula: '1*0', result: 0 } };
    sheet.addTable({ name: 'Table1', ref: 'A1', headerRow: true, totalsRow: true,
      columns: headers.map(name => ({ name, ...(name === 'Comm' ? { totalsRowFunction: 'sum' } : {}) })),
      rows: [headers.map(name => values[name] ?? null)] });
    const history = workbook.addWorksheet('Paid to Emp 2026');
    history.addRow(['Invoice #', 'Paycheck Date']); history.addRow(['SYNTHETIC-PAID', 'Synthetic marker']);
    const target = path.join(temporary, 'Synthetic.xlsx'); await workbook.xlsx.writeFile(target);
    const backup = path.join(temporary, 'Original.xlsx'); fs.copyFileSync(target, backup);
    const payload = path.join(temporary, 'write.json');
    fs.writeFileSync(payload, JSON.stringify({ workbookPath: target, backupPath: backup,
      existingRowUpdates: [{ rowNumber: 2, projectID: 101, invoice: 'SYNTHETIC-1', changes: [
        { field: 'Inv Amt (Pretax)', newValue: 123.45 }, { field: 'Invoice Date', newValue: '2026-01-02T00:00:00.000Z' },
        { field: 'Paid Amt (Pretax)', newValue: 123.45 }, { field: 'Paid Date', newValue: '2026-01-03T00:00:00.000Z' },
      ] }] }));
    const writer = spawnSync(process.execPath, [path.join(appDir, 'app', 'xlsx-writer.js'), payload], { encoding: 'utf8' });
    assert.equal(writer.status, 0, writer.stderr || writer.stdout);
    const saved = await loadWorkbook(target);
    assert.equal(saved.getWorksheet('Commission').getCell(2, headers.indexOf('Inv Amt (Pretax)') + 1).value, 123.45);
    assert.equal(saved.getWorksheet('Paid to Emp 2026').getCell('A2').value, 'SYNTHETIC-PAID');
    const executable = chromium.executablePath();
    assert(fs.existsSync(executable), 'Compatible Chromium must be present');
    const relative = path.relative(path.join(appDir, 'browsers'), executable);
    assert(relative && !relative.startsWith('..') && !path.isAbsolute(relative), 'Browser must come from the bundled folder');
    browser = await chromium.launch({ headless: true, executablePath: executable });
    const page = await browser.newPage();
    await page.setContent('<h1>Packaged browser ready</h1>');
    assert.equal(await page.textContent('h1'), 'Packaged browser ready');
    console.log(JSON.stringify({ appVersion: requireApp('./package.json').version, node: process.version,
      playwright: requireApp('playwright/package.json').version, chromium: browser.version(), result: 'passed' }));
  } finally {
    if (browser) await browser.close();
    fs.rmSync(temporary, { recursive: true, force: true });
  }
}

main().catch(error => { console.error(error.stack || error.message); process.exitCode = 1; });
