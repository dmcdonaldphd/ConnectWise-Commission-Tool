'use strict';
// Synthetic end-to-end UI acceptance. Uses the same shipped server, controller,
// HTML and scripts, with fake ConnectWise/workbook adapters; never real user data.
const assert = require('node:assert/strict');
const { createRequire } = require('node:module');
const path = require('node:path');

async function smokeUi(browser, requireApp, screenshot) {
  const { createGuiService } = requireApp('./app/gui-service');
  const { createGuiServer } = requireApp('./app/gui-server');
  const { logger } = requireApp('./app/log');
  let release;
  const gate = new Promise(resolve => { release = resolve; });
  const calls = [];
  const names = ['A. Sample - Commission Earning.xlsx', 'B. Sample - Commission Earning.xlsx'];
  const service = createGuiService({
    initialize: async () => true,
    locations: () => ({ workbookDir: 'C:\\PHDTech\\Commission Workbooks', logDir: 'Synthetic Logs', backupDir: 'Synthetic Backups' }),
    discover: async () => ({ valid: names.map(fileName => ({ fileName })), skipped: [] }),
    chooseFolder: async () => ({ changed: false, needsPath: true }),
    openDirectory: async folder => calls.push(folder),
    openConnectWise: async () => ({ close: async () => calls.push('close') }),
    readSummary: () => 'Synthetic run summary\n<svg onload="window.injected=true">\nReview source conflict before payroll.',
    reconcile: async (_ctx, name, { onProgress }) => {
      calls.push(name);
      onProgress({ stage: 'reconciling', current: 2, total: 12, projectID: 101 });
      logger.log('Synthetic activity <img src=x onerror="window.injected=true">');
      await gate;
      if (name === names[1]) throw Object.assign(new Error('Synthetic workbook is locked in Excel.'), { summaryLogPath: 'Synthetic failed summary' });
      return { workbookFilename: name, status: 'COMPLETED_WITH_REVIEW', workbookWritten: true,
        reviewCount: 2, summaryLogPath: 'Synthetic summary',
        summary: { rowsUpdated: 3, blankRowsFilled: 1, rowsInserted: 2 } };
    },
  });
  await service.initialize();
  const server = await createGuiServer(service);
  const page = await browser.newPage({ viewport: { width: 1200, height: 1000 } });
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  try {
    await page.goto(server.url);
    await page.waitForFunction(() => document.getElementById('selection-count').textContent === '2 of 2 selected');
    assert.equal(await page.locator('.workbook-row').count(), 2);
    if (screenshot) await page.screenshot({ path: screenshot, fullPage: true });
    await page.getByRole('button', { name: 'Change', exact: true }).click();
    await page.locator('#folder-dialog').waitFor({ state: 'visible' });
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    assert.match(await page.locator('#folder-path').textContent(), /Commission Workbooks/);
    await page.locator('#select-all').uncheck(); assert(await page.locator('#primary').isDisabled());
    await page.locator('#select-all').check();
    await page.locator('#primary').click();
    await page.locator('#login-notice').waitFor({ state: 'visible' });
    assert.deepEqual(calls, []);
    assert(await page.locator('#change-folder').isDisabled());
    await page.locator('#primary').click();
    await page.waitForFunction(() => document.getElementById('progress-detail').textContent.includes('2 of 12'));
    assert(await page.locator('#primary').isDisabled());
    assert.deepEqual(calls, [names[0]]);
    release();
    await page.waitForFunction(() => document.getElementById('phase-badge').textContent === 'Finished');
    assert.equal(await page.locator('#result-rows tr').count(), 2);
    assert.match(await page.locator('#result-rows').textContent(), /Review needed/);
    assert.match(await page.locator('#result-rows').textContent(), /3 updated · 1 filled · 2 added/);
    assert.match(await page.locator('#result-rows').textContent(), /locked in Excel/);
    await page.getByRole('button', { name: 'Run summary', exact: true }).first().click();
    await page.locator('#summary-dialog').waitFor({ state: 'visible' });
    assert.match(await page.locator('#summary-text').textContent(), /<svg onload=/);
    assert.equal(await page.evaluate(() => window.injected), undefined);
    assert.equal(await page.locator('#summary-text svg').count(), 0);
    await page.locator('#summary-dialog .close-dialog').click();
    await page.getByRole('button', { name: 'Logs folder', exact: true }).click();
    await page.waitForFunction(() => !document.getElementById('primary').disabled);
    assert(calls.includes('Synthetic Logs'));
    await page.setViewportSize({ width: 700, height: 900 });
    assert(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
    assert.deepEqual(errors, []);
    console.log('Desktop UI: selection, login gate, progress, results, summary, folder actions and text rendering passed.');
  } finally {
    release(); await service.close(); await page.close(); await server.close();
  }
}
module.exports = { smokeUi };
if (require.main === module) {
  const requireApp = createRequire(path.resolve(__dirname, '../package.json'));
  const { chromium } = requireApp('playwright');
  (async () => {
    const browser = await chromium.launch({ headless: true, executablePath: process.env.COMMISSION_TEST_BROWSER || chromium.executablePath() });
    try { await smokeUi(browser, requireApp, process.env.COMMISSION_UI_SCREENSHOT); }
    finally { await browser.close(); }
  })().catch(error => { console.error(error); process.exitCode = 1; });
}
