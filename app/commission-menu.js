const fs = require('fs');
const path = require('path');
const readline = require('readline');

const { chromium, BASE_URL } = require('./connectwise');
const { loadWorkbook } = require('./workbook');
const { reconcileWorkbook } = require('./app');
const { describeRunStatus } = require('./audit');
const { createSettingsStore, initializeWorkbookFolder, chooseWorkbookFolder, openFolder } = require('./settings');

const paths = require('./paths');

function ask(message) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return new Promise((resolve) => {
    rl.question(message, (answer) => {
      rl.close();
      resolve(answer.trim());
    });
  });
}

function isCandidateFile(fileName) {
  const lower = fileName.toLowerCase();

  return (
    lower.endsWith('.xlsx') &&
    !lower.startsWith('~$') &&
    !lower.endsWith('.working.xlsx') &&
    !/ - \d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}(?:-[a-f0-9]{8})?\.xlsx$/.test(lower)
  );
}

function normalizeSheetName(name) {
  return String(name || '').trim().toLowerCase();
}

function getCommissionSheet(workbook) {
  return workbook.worksheets.find(
    (sheet) => normalizeSheetName(sheet.name) === 'commission'
  );
}

function getPaidHistorySheets(workbook) {
  return workbook.worksheets.filter((sheet) =>
    normalizeSheetName(sheet.name).startsWith('paid to emp')
  );
}

async function findCommissionWorkbooks() {
  const candidates = fs
    .readdirSync(paths.WORKBOOK_DIR, { withFileTypes: true })
    .filter((entry) => entry.isFile() && isCandidateFile(entry.name))
    .map((entry) => entry.name)
    .sort((a, b) => a.localeCompare(b));

  const valid = [];
  const skipped = [];

  for (const fileName of candidates) {
    try {
      const workbook = await loadWorkbook(path.join(paths.WORKBOOK_DIR, fileName));
      const commissionSheet = getCommissionSheet(workbook);
      const paidHistorySheets = getPaidHistorySheets(workbook);

      if (!commissionSheet) {
        skipped.push({
          fileName,
          reason: 'No Commission worksheet found',
        });
        continue;
      }

      if (paidHistorySheets.length === 0) {
        skipped.push({
          fileName,
          reason: 'No worksheet beginning with "Paid to Emp" found',
        });
        continue;
      }

      valid.push({
        fileName,
        paidHistorySheetNames: paidHistorySheets.map((sheet) => sheet.name),
      });
    } catch (error) {
      skipped.push({
        fileName,
        reason: `Could not read workbook: ${error.message}`,
      });
    }
  }

  return { valid, skipped };
}

async function selectWorkbooks({ ask: askQuestion = ask, discover = findCommissionWorkbooks,
  changeFolder, openDirectory = openFolder } = {}) {
  const chooseAnotherFolder = changeFolder || (() => chooseWorkbookFolder({ ask: askQuestion }));
  while (true) {
    let discovery;
    try { discovery = await discover(); }
    catch (error) {
      console.error(`Could not read the workbook folder: ${error.message}`);
      discovery = { valid: [], skipped: [] };
    }
    const workbooks = discovery.valid;
    console.log('');
    console.log(`Workbook folder: ${paths.WORKBOOK_DIR}`);
    console.log('');
    const options = workbooks.map(info => ({ label: info.fileName, workbook: info }));
    if (workbooks.length) options.push({ label: 'Process all workbooks', action: 'all' });
    else console.log('No eligible commission workbooks found. Choose a folder containing Commission and Paid to Emp worksheets.');
    options.push({ label: 'Change workbook folder', action: 'change' },
      { label: 'Open Logs folder', action: 'logs' }, { label: 'Open Backups folder', action: 'backups' });
    options.forEach((option, index) => {
      console.log(`${index + 1}. ${option.label}`);
      if (option.workbook) console.log(`   Paid history: ${option.workbook.paidHistorySheetNames.join(', ')}`);
    });
    console.log('0. Exit');
    if (discovery.skipped.length) {
      console.log('Other Excel workbooks skipped:');
      for (const item of discovery.skipped) console.log(`- ${item.fileName}: ${item.reason}`);
    }
    const answer = String(await askQuestion('Select an option: ')).trim();
    if (answer === '0') return null;
    const selection = Number(answer);
    if (!answer || !Number.isInteger(selection) || selection < 1 || selection > options.length) {
      console.log('Invalid selection. Please choose one of the listed options.');
      continue;
    }
    const option = options[selection - 1];
    if (option.workbook) return [option.workbook];
    if (option.action === 'all') return workbooks;
    if (option.action === 'change') await chooseAnotherFolder();
    else {
      try { await openDirectory(option.action === 'logs' ? paths.LOG_DIR : paths.BACKUP_DIR); }
      catch (error) { console.error(error.message); }
    }
  }
}

async function main() {
  const store = createSettingsStore();
  if (!await initializeWorkbookFolder({ store, ask })) {
    console.log('The tool has closed without processing any workbooks.');
    return;
  }
  const selected = await selectWorkbooks({ changeFolder: () => chooseWorkbookFolder({ store, ask }) });
  if (!selected) return;

  console.log('');
  console.log('Launching ConnectWise...');

  const context = await chromium.launchPersistentContext(
    paths.PROFILE_DIR,
    {
      headless: false,
      viewport: null,
    }
  );

  try {
    const page = context.pages()[0] || (await context.newPage());

    await page.goto(BASE_URL, {
      waitUntil: 'domcontentloaded',
    });

    console.log('');
    console.log('ConnectWise has opened. Log in if necessary.');

    await ask(
      'Once ConnectWise is fully loaded, return here and press ENTER...'
    );

    const results = await processWorkbooks(context, selected);
    if (results.some(result => ['FAILED', 'COMPLETED_WITH_ERRORS'].includes(result.status))) {
      process.exitCode = 1;
    }
  } finally {
    await context.close();
  }
}

async function processWorkbooks(context, selected, reconcile = reconcileWorkbook) {
  const results = [];
  for (let index = 0; index < selected.length; index += 1) {
    const fileName = selected[index].fileName;

    console.log('');
    console.log('==================================================');
    console.log(`Workbook ${index + 1} of ${selected.length}: ${fileName}`);
    console.log('==================================================');

    try {
      results.push(await reconcile(context, fileName));
    } catch (error) {
      console.error(`${fileName}: FAILED — ${error.message}`);
      results.push({ workbookFilename: fileName, status: 'FAILED', auditLogPath: error.auditLogPath });
    }
  }

  console.log('');
  console.log('Selected workbook results:');
  for (const result of results) {
    console.log(`${result.workbookFilename}: ${describeRunStatus(result.status)}`);
    if (result.auditLogPath) console.log(`  Audit: ${result.auditLogPath}`);
  }
  const needsAttention = results.some(result => result.status !== 'COMPLETED');
  console.log(needsAttention ? 'Processing finished. Review the results above and saved run summaries.' : 'All selected workbooks completed successfully.');
  return results;
}

if (require.main === module) main().catch((error) => {
  console.error('');
  console.error('FATAL ERROR');
  console.error(error.stack || error.message);
  process.exitCode = 1;
});

module.exports = { processWorkbooks, findCommissionWorkbooks, selectWorkbooks, main };
