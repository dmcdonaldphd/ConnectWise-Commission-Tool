const fs = require('fs');
const path = require('path');
const readline = require('readline');

const { chromium, BASE_URL } = require('./connectwise');
const { loadWorkbook } = require('./workbook');
const { reconcileWorkbook } = require('./app');
const { describeRunStatus } = require('./audit');

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
    !lower.endsWith('.working.xlsx')
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

async function main() {
  const discovery = await findCommissionWorkbooks();
  const workbooks = discovery.valid;

  if (workbooks.length === 0) {
    throw new Error(
      'No eligible commission workbooks were found. Each workbook must contain a Commission sheet and at least one sheet whose name begins with "Paid to Emp".'
    );
  }

  console.log('');
  console.log('Commission workbooks found:');
  console.log('');

  workbooks.forEach((workbookInfo, index) => {
    console.log(`${index + 1}. ${workbookInfo.fileName}`);
    console.log(
      `   Paid history: ${workbookInfo.paidHistorySheetNames.join(', ')}`
    );
  });

  if (discovery.skipped.length > 0) {
    console.log('');
    console.log('Other Excel workbooks skipped:');
    discovery.skipped.forEach((item) => {
      console.log(`- ${item.fileName}: ${item.reason}`);
    });
  }

  console.log(`${workbooks.length + 1}. Process all workbooks`);
  console.log('');

  const answer = await ask('Select an option: ');
  const selection = Number(answer);

  if (
    !Number.isInteger(selection) ||
    selection < 1 ||
    selection > workbooks.length + 1
  ) {
    throw new Error('Invalid selection.');
  }

  const selected =
    selection === workbooks.length + 1
      ? workbooks
      : [workbooks[selection - 1]];

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

module.exports = { processWorkbooks, findCommissionWorkbooks, main };
