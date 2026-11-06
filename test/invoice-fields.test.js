'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');
const paths = require('../app/paths');
const { reconcileWorkbook } = require('../app/app');
const { loadWorkbook, getHeaderMap, readProjects } = require('../app/workbook');
const { writeWorkbook } = require('../app/xlsx-writer');
const { getProjectInvoices } = require('../app/connectwise');
const { processWorkbooks } = require('../app/commission-menu');
const { createSettingsStore, initializeWorkbookFolder } = require('../app/settings');

// All workbooks and ConnectWise responses are synthetic. Tests exercise the real
// reconciliation, response parsing, child-process writer, and saved XLSX files.
const HEADERS = ['%', 'Company', 'Project ID', 'Project/Ticket', 'Notes',
  'Invoice Date', 'Invoice #', 'Inv Amt (Pretax)', 'Paid Amt (Pretax)',
  'Paid Date', 'Comm', 'Paycheck Date'];
const REORDERED = ['Notes', 'Invoice #', 'Invoice Date', 'Company', '%',
  'Paycheck Date', 'Inv Amt (Pretax)', 'Project/Ticket', 'Paid Date',
  'Project ID', 'Comm', 'Paid Amt (Pretax)', 'Extra column'];
const OLD_DATE = new Date('2026-09-01T00:00:00.000Z');
const INVOICE_DATE = new Date('2026-10-15T00:00:00.000Z');
// Match the application's existing ConnectWise timestamp/calendar convention.
const SOURCE_DATE = new Date(2026, 9, 15, 12).getTime();

function colName(number) {
  let result = '';
  while (number > 0) {
    number -= 1;
    result = String.fromCharCode(65 + number % 26) + result;
    number = Math.floor(number / 26);
  }
  return result;
}

async function setup(t, records, { headers = HEADERS, history = [], spareRows = 0, dateFormat = 'm/d/yyyy' } = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'commission-invoice-fields-'));
  const originalPaths = { ...paths };
  Object.assign(paths, { WORKBOOK_DIR: dir, BACKUP_DIR: path.join(dir, 'Backups'),
    LOG_DIR: path.join(dir, 'Logs'), TEMP_DIR: dir });
  t.after(() => {
    Object.assign(paths, originalPaths);
    fs.rmSync(dir, { recursive: true, force: true });
  });
  const messages = [];
  for (const method of ['log', 'warn', 'error']) {
    t.mock.method(console, method, (...args) => messages.push(args.join(' ')));
  }
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Commission');
  const all = [...records, ...Array.from({ length: spareRows }, () => ({}))];
  sheet.addTable({ name: 'Table1', ref: 'A1', headerRow: true, totalsRow: true,
    columns: headers.map(name => ({ name, ...(name === 'Comm' ? { totalsRowFunction: 'sum' } : {}) })),
    rows: all.map((record, index) => headers.map(name => {
      if (name === 'Comm') {
        return { formula: `${colName(headers.indexOf('Paid Amt (Pretax)') + 1)}${index + 2}*${colName(headers.indexOf('%') + 1)}${index + 2}`, result: 0 };
      }
      if (name === '%') return record[name] ?? 0.1;
      return record[name] ?? null;
    })) });
  for (let row = 2; row <= all.length + 1; row++) {
    for (const name of ['Invoice Date', 'Paid Date', 'Paycheck Date']) {
      if (headers.includes(name)) sheet.getCell(row, headers.indexOf(name) + 1).numFmt = dateFormat;
    }
    for (const name of ['Inv Amt (Pretax)', 'Paid Amt (Pretax)']) {
      if (headers.includes(name)) sheet.getCell(row, headers.indexOf(name) + 1).numFmt = '$#,##0.00';
    }
  }
  for (const name of ['Paid to Emp 2025', 'Paid to Emp 2026']) {
    const paid = workbook.addWorksheet(name);
    // Deliberately use a different header order from Commission.
    paid.addRow(['Notes', 'Paycheck Date', 'Invoice #']);
    for (const invoice of history) paid.addRow(['Synthetic paid history', OLD_DATE, invoice]);
  }
  const other = workbook.addWorksheet('Unrelated');
  other.getCell('A1').value = 'Preserve unrelated data';
  const target = path.join(dir, 'Synthetic.xlsx');
  await workbook.xlsx.writeFile(target);
  const before = fs.readFileSync(target);
  return { dir, target, before, headers, messages };
}

function record(invoice, project = 101, overrides = {}) {
  return { 'Project ID': project, 'Company': 'Synthetic company',
    'Project/Ticket': 'Synthetic project', 'Notes': 'Keep notes',
    'Invoice #': invoice, ...overrides };
}

function source(invoice, { project = 101, amount = 1234.56, date = SOURCE_DATE,
  status = 'Unpaid' } = {}) {
  return { project, invoice_number: invoice, billing_log_recid: `${project}:${invoice}`,
    invoice_amount: amount * 1.08, invoiceSubTotal: amount,
    date_invoice: date, paid_amount: status === 'Paid' ? amount * 1.08 : status === 'Partially Paid' ? 50 : 0,
    invoiceBalance: status === 'Paid' ? 0 : status === 'Partially Paid' ? amount * 1.08 - 50 : amount * 1.08,
    date_paid: status === 'Paid' ? SOURCE_DATE : null };
}

async function reconcile(fixture, invoices, { failInvoice, failProject } = {}) {
  const requests = [];
  const context = { request: { post: async (url, options) => {
    const outer = JSON.parse(new URLSearchParams(options.data).get('actionMessage'));
    const payload = JSON.parse(outer.payload);
    requests.push({ url, payload });
    let action;
    if (url.includes('GetMultilineDataAction_projectinvoicelist')) {
      const project = Number(payload.multilineUserParams[0].paramValue);
      if (project === failProject) throw new Error('Synthetic project listing failure');
      const start = (payload.activePage - 1) * payload.maxResult;
      action = { multilineRows: invoices.filter(i => i.project === project).slice(start, start + payload.maxResult).map(i => ({ row: JSON.stringify(i) })) };
      assert(payload.columns.includes('date_invoice'));
    } else {
      const invoice = invoices.find(i => i.billing_log_recid === payload.billingLogRecID);
      assert(invoice, 'Unexpected invoice request');
      if (invoice.invoice_number === failInvoice) throw new Error('Synthetic invoice retrieval failure');
      action = url.includes('GetInvoiceCalculationsPodAction')
        ? { invoiceCalculationsPodViewModel: { invoiceSubTotal: invoice.invoiceSubTotal } }
        : { invoiceCreditsAndPaymentsPodViewModel: { invoiceBalance: invoice.invoiceBalance } };
    }
    return { ok: () => true, text: async () => JSON.stringify({ success: true, data: { isSuccess: true, action } }) };
  } } };
  let outcome;
  try {
    outcome = await reconcileWorkbook(context, path.basename(fixture.target));
    fixture.lastAuditPath = outcome.auditLogPath;
  } catch (error) {
    fixture.lastAuditPath = error.auditLogPath;
    throw error;
  }
  const workbook = await loadWorkbook(fixture.target);
  const sheet = workbook.getWorksheet('Commission');
  const columns = getHeaderMap(sheet);
  return { workbook, sheet, requests, outcome,
    cell: (row, name) => sheet.getCell(row, columns.get(name)),
    rows: readProjects(sheet) };
}

function audit(fixture) {
  if (fixture.lastAuditPath) return fs.readFileSync(fixture.lastAuditPath, 'utf8');
  const dir = path.join(fixture.dir, 'Logs');
  const file = fs.readdirSync(dir).find(name => name.endsWith('.csv'));
  assert(file, 'Expected saved reconciliation audit');
  return fs.readFileSync(path.join(dir, file), 'utf8');
}

async function assertPreservedParts(before, after) {
  const oldZip = await JSZip.loadAsync(before);
  const newZip = await JSZip.loadAsync(after, { checkCRC32: true });
  const mutable = new Set(['xl/worksheets/sheet1.xml', 'xl/tables/table1.xml',
    'xl/styles.xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels',
    '[Content_Types].xml', 'xl/calcChain.xml']);
  for (const [name, entry] of Object.entries(oldZip.files)) {
    if (entry.dir || mutable.has(name)) continue;
    assert(newZip.file(name), `Missing package part ${name}`);
    assert.deepEqual(await newZip.file(name).async('nodebuffer'), await entry.async('nodebuffer'), `Changed unrelated package part ${name}`);
  }
}

for (const [name, status] of [['unpaid', 'Unpaid'], ['partially paid', 'Partially Paid'], ['paid', 'Paid']]) {
  test(`fills existing ${name} invoice metadata from pretax subtotal`, async t => {
    const f = await setup(t, [record('INV-1')]);
    const result = await reconcile(f, [source('INV-1', { status })]);
    assert.equal(result.cell(2, 'Inv Amt (Pretax)').value, 1234.56);
    assert.deepEqual(result.cell(2, 'Invoice Date').value, INVOICE_DATE);
    assert.equal(result.cell(2, 'Paid Amt (Pretax)').value ?? null, status === 'Paid' ? 1234.56 : null);
    assert.equal(result.cell(2, 'Paid Date').value instanceof Date, status === 'Paid');
    assert.equal(result.cell(2, 'Inv Amt (Pretax)').fill.fgColor.argb, 'FFFFFF00');
    assert.equal(result.cell(2, 'Invoice Date').fill.fgColor.argb, 'FFFFFF00');
    assert.equal(result.cell(2, 'Invoice Date').numFmt, 'm/d/yyyy');
    assert.equal(result.cell(2, 'Notes').value, 'Keep notes');
    assert.equal(result.cell(2, 'Comm').value.formula, 'I2*A2');
    const backups = fs.readdirSync(path.join(f.dir, 'Backups'));
    assert.equal(backups.length, 1);
    assert.deepEqual(fs.readFileSync(path.join(f.dir, 'Backups', backups[0])), f.before);
    await assertPreservedParts(f.before, fs.readFileSync(f.target));
    const csv = audit(f);
    assert(csv.includes('Inv Amt (Pretax)'));
    assert(csv.includes('Invoice Date,blank,10/15/2026'));
  });
}

test('replaces differing invoice fields by header after columns are reordered', async t => {
  const f = await setup(t, [record('INV-1', 101, {
    'Invoice Date': OLD_DATE, 'Inv Amt (Pretax)': 99, 'Extra column': 'Retain me',
    'Paid Amt (Pretax)': 75, 'Paid Date': OLD_DATE,
  })], { headers: REORDERED });
  const result = await reconcile(f, [source('INV-1')]);
  assert.equal(result.cell(2, 'Inv Amt (Pretax)').value, 1234.56);
  assert.deepEqual(result.cell(2, 'Invoice Date').value, INVOICE_DATE);
  assert.equal(result.cell(2, 'Extra column').value, 'Retain me');
  assert.equal(result.cell(2, 'Paid Amt (Pretax)').value ?? null, null);
  assert.equal(result.cell(2, 'Paid Date').value ?? null, null);
  assert.equal(result.cell(2, 'Comm').value.formula, 'L2*E2');
  assert(audit(f).includes('Invoice Date,9/1/2026,10/15/2026'));
});

test('fills both fields on a blank project row and records overwritten values', async t => {
  const f = await setup(t, [record(null, 101, { 'Invoice Date': OLD_DATE, 'Inv Amt (Pretax)': 99 })], { headers: REORDERED });
  const result = await reconcile(f, [source('INV-NEW')]);
  assert.equal(result.cell(2, 'Invoice #').value, 'INV-NEW');
  assert.equal(result.cell(2, 'Inv Amt (Pretax)').value, 1234.56);
  assert.deepEqual(result.cell(2, 'Invoice Date').value, INVOICE_DATE);
  for (const field of ['Invoice #', 'Invoice Date', 'Inv Amt (Pretax)']) {
    assert.equal(result.cell(2, field).fill.fgColor.argb, 'FFDAF2D0');
  }
  const csv = audit(f);
  assert(csv.includes('FILL_BLANK_ROW'));
  assert(csv.includes('Invoice / Inv Amt (Pretax) / Invoice Date'));
  assert(csv.includes('10/15/2026'));
  assert(csv.includes('blank; $99.00; 9/1/2026; blank; blank'));
});

for (const spareRows of [0, 2]) {
  test(`inserts metadata next to its project ${spareRows ? 'using spare rows' : 'with table expansion'}`, async t => {
    const f = await setup(t, [record('INV-1'), record('INV-2', 202)], { headers: REORDERED, spareRows });
    const result = await reconcile(f, [source('INV-1'), source('INV-NEW', { status: 'Paid' }), source('INV-2', { project: 202 })]);
    assert.deepEqual(result.rows.map(r => r.invoice), ['INV-1', 'INV-NEW', 'INV-2']);
    assert.equal(result.cell(3, 'Project ID').value, 101);
    assert.equal(result.cell(3, 'Inv Amt (Pretax)').value, 1234.56);
    assert.deepEqual(result.cell(3, 'Invoice Date').value, INVOICE_DATE);
    assert.equal(result.cell(3, 'Comm').value.formula, 'L3*E3');
    assert.equal(result.cell(4, 'Comm').value.formula, 'L4*E4');
    assert.equal(result.cell(3, 'Invoice Date').fill.fgColor.argb, 'FFDAF2D0');
    assert.equal(result.cell(3, 'Invoice Date').numFmt, 'm/d/yyyy');
    assert.equal(result.sheet.getTable('Table1').table.tableRef, spareRows ? 'A1:M6' : 'A1:M5');
    assert(audit(f).includes('INSERT_NEW_ROW'));
    await assertPreservedParts(f.before, fs.readFileSync(f.target));
  });
}

test('repeated reconciliation leaves matching invoice dates and amounts untouched', async t => {
  const f = await setup(t, [record('INV-1')]);
  const invoices = [source('INV-1', { status: 'Paid' }), source('INV-NEW')];
  await reconcile(f, invoices);
  const afterFirst = fs.readFileSync(f.target);
  const backups = fs.readdirSync(path.join(f.dir, 'Backups'));
  await reconcile(f, invoices);
  assert.deepEqual(fs.readFileSync(f.target), afterFirst);
  assert.deepEqual(fs.readdirSync(path.join(f.dir, 'Backups')), backups);
  assert(f.messages.some(message => message.includes('No workbook changes are required.')));
});

test('clears the previous invoice-amount highlight when only the date changes next time', async t => {
  const f = await setup(t, [record('INV-1')]);
  await reconcile(f, [source('INV-1')]);
  const result = await reconcile(f, [source('INV-1', { date: new Date(2026, 9, 16, 12).getTime() })]);
  assert.equal(result.cell(2, 'Invoice Date').fill.fgColor.argb, 'FFFFFF00');
  assert.notEqual(result.cell(2, 'Inv Amt (Pretax)').fill.fgColor?.argb, 'FFFFFF00');
  assert.equal(result.cell(2, 'Inv Amt (Pretax)').numFmt, '$#,##0.00');
  assert.equal(result.cell(2, 'Invoice Date').value.toISOString(), '2026-10-16T00:00:00.000Z');
});

test('rejects invalid source dates and holds the whole project without writing', async t => {
  const f = await setup(t, [record('INV-1', 101, { 'Invoice Date': OLD_DATE })]);
  const result = await reconcile(f, [source('INV-1', { date: 'not-a-date' })]);
  assert.deepEqual(result.cell(2, 'Invoice Date').value, OLD_DATE);
  assert.equal(result.cell(2, 'Inv Amt (Pretax)').value ?? null, null);
  assert.equal(result.outcome.status, 'FAILED');
  assert(audit(f).includes('PROJECT_ERROR'));
  assert.deepEqual(fs.readFileSync(f.target), f.before);
});

test('clears an old date on a blank project slot when ConnectWise explicitly has no date', async t => {
  const f = await setup(t, [record(null, 101, { 'Invoice Date': OLD_DATE })]);
  const result = await reconcile(f, [source('INV-NEW', { date: null })]);
  assert.equal(result.cell(2, 'Invoice #').value, 'INV-NEW');
  assert.equal(result.cell(2, 'Inv Amt (Pretax)').value, 1234.56);
  assert.equal(result.cell(2, 'Invoice Date').value ?? null, null);
  assert(audit(f).includes('FILL_BLANK_ROW'));
});

test('accepts numeric timestamp strings at local midnight without repeat date changes', async t => {
  const f = await setup(t, [record('INV-1')]);
  const invoices = [source('INV-1', { date: String(new Date(2026, 9, 15).getTime()) })];
  const result = await reconcile(f, invoices);
  assert.deepEqual(result.cell(2, 'Invoice Date').value, INVOICE_DATE);
  const first = fs.readFileSync(f.target);
  await reconcile(f, invoices);
  assert.deepEqual(fs.readFileSync(f.target), first);
});

test('General-formatted date cells display and round-trip as dates, including new styles', async t => {
  const f = await setup(t, [record('INV-1'), record('INV-2', 202)], { dateFormat: 'General' });
  const invoices = [source('INV-1', { status: 'Paid' }), source('INV-NEW'), source('INV-2', { project: 202 })];
  const result = await reconcile(f, invoices);
  for (const row of [2, 3, 4]) {
    assert.deepEqual(result.cell(row, 'Invoice Date').value, INVOICE_DATE);
    assert.notEqual(result.cell(row, 'Invoice Date').numFmt, 'General');
  }
  assert(result.cell(2, 'Paid Date').value instanceof Date);
  const first = fs.readFileSync(f.target);
  await reconcile(f, invoices);
  assert.deepEqual(fs.readFileSync(f.target), first);
});

test('does not infer a subtotal from missing or invalid ConnectWise data', async t => {
  const f = await setup(t, [record('INV-1', 101, { 'Inv Amt (Pretax)': 99 })]);
  for (const amount of [null, 'invalid']) {
    await reconcile(f, [source('INV-1', { amount })]);
    assert.deepEqual(fs.readFileSync(f.target), f.before);
  }
  assert(f.messages.some(message => message.includes('no valid ConnectWise pretax subtotal')));
});

test('retains already-paid, zero-dollar and duplicate-workbook exclusions', async t => {
  const f = await setup(t, [record('INV-PAID'), record('INV-DUP'), record('INV-DUP')], { history: ['INV-PAID'] });
  await reconcile(f, [source('INV-PAID'), source('INV-ZERO', { amount: 0 }), source('INV-DUP')]);
  assert.deepEqual(fs.readFileSync(f.target), f.before);
  assert(!fs.existsSync(path.join(f.dir, 'Backups')));
});

for (const field of ['Invoice Date', 'Inv Amt (Pretax)']) {
  test(`missing required ${field} header stops before fetching or writing`, async t => {
    const f = await setup(t, [record('INV-1')], { headers: HEADERS.filter(h => h !== field) });
    await assert.rejects(reconcile(f, [source('INV-1')]), /missing required column header/);
    assert.deepEqual(fs.readFileSync(f.target), f.before);
    assert(!fs.existsSync(path.join(f.dir, 'Backups')));
  });
}

test('duplicate invoice metadata header is rejected without modifying the original', async t => {
  const f = await setup(t, [record('INV-1')]);
  const zip = await JSZip.loadAsync(f.before);
  const strings = await zip.file('xl/sharedStrings.xml').async('string');
  zip.file('xl/sharedStrings.xml', strings.replace('<t>Notes</t>', '<t>Invoice Date</t>'));
  const malformed = await zip.generateAsync({ type: 'nodebuffer' });
  fs.writeFileSync(f.target, malformed);
  await assert.rejects(reconcile(f, [source('INV-1')]), /duplicate header/);
  assert.deepEqual(fs.readFileSync(f.target), malformed);
});

test('writer rejects formula replacement in invoice metadata and preserves the original', async t => {
  const f = await setup(t, [record('INV-1', 101, { 'Inv Amt (Pretax)': { formula: '1+1', result: 2 } })]);
  await assert.rejects(writeWorkbook({ workbookPath: f.target, existingRowUpdates: [{ rowNumber: 2,
    projectID: 101, invoice: 'INV-1', changes: [{ field: 'Inv Amt (Pretax)', newValue: 1234.56 }] }] }), /attempted to replace formula/);
  assert.deepEqual(fs.readFileSync(f.target), f.before);
});

async function rewriteFixture(fixture, change) {
  const workbook = await loadWorkbook(fixture.target);
  change(workbook);
  await workbook.xlsx.writeFile(fixture.target);
  fixture.before = fs.readFileSync(fixture.target);
}

test('finds paid invoices after blank history rows, including cached formula values', async t => {
  const f = await setup(t, [record(null)]);
  await rewriteFixture(f, workbook => {
    workbook.getWorksheet('Paid to Emp 2026').getCell('C100').value = { formula: '"INV-PAID"', result: 'INV-PAID' };
  });
  const result = await reconcile(f, [source('INV-PAID')]);
  assert.equal(result.rows.filter(row => row.invoice === 'INV-PAID').length, 0);
  assert.deepEqual(fs.readFileSync(f.target), f.before);
  assert(audit(f).includes('SKIP_ALREADY_PAID'));
});

test('stops safely when a history invoice formula lacks its cached value', async t => {
  const f = await setup(t, [record(null)]);
  await rewriteFixture(f, workbook => {
    workbook.getWorksheet('Paid to Emp 2026').getCell('C100').value = { formula: '"INV-PAID"' };
  });
  await assert.rejects(reconcile(f, [source('INV-PAID')]), /no cached value/);
  assert.deepEqual(fs.readFileSync(f.target), f.before);
  assert(audit(f).includes('RUN_ERROR'));
});

test('does not insert a source invoice already listed under a different project', async t => {
  const f = await setup(t, [record('INV-A'), record('INV-X', 202)]);
  const result = await reconcile(f, [source('INV-X')]);
  assert.equal(result.rows.filter(row => row.invoice === 'INV-X').length, 1);
  assert.equal(result.outcome.status, 'COMPLETED_WITH_ERRORS');
  assert.deepEqual(fs.readFileSync(f.target), f.before);
  assert(audit(f).includes('already exists in Commission row'));
});

test('blocks duplicate insertion even when the occupied row has an invalid Project ID', async t => {
  const f = await setup(t, [record('INV-A'), record('INV-X', 'invalid')]);
  await reconcile(f, [source('INV-X')]);
  assert.deepEqual(fs.readFileSync(f.target), f.before);
  assert(audit(f).includes('another or invalid project'));
});

test('flags existing cross-project duplicates without changing either invoice', async t => {
  const f = await setup(t, [record('INV-DUP'), record('INV-DUP', 202)]);
  const result = await reconcile(f, [source('INV-DUP')]);
  assert.equal(result.outcome.status, 'COMPLETED_WITH_REVIEW');
  assert.deepEqual(fs.readFileSync(f.target), f.before);
  assert(audit(f).includes('REVIEW_DUPLICATE'));
});

test('processes an identical repeated source invoice once', async t => {
  const f = await setup(t, [record(null)]);
  const invoice = source('INV-NEW');
  const result = await reconcile(f, [invoice, { ...invoice }]);
  assert.equal(result.rows.filter(row => row.invoice === 'INV-NEW').length, 1);
  assert(audit(f).includes('SKIP_DUPLICATE_SOURCE'));
});

test('holds conflicting source records with the same invoice number', async t => {
  const f = await setup(t, [record(null)]);
  const first = source('INV-NEW');
  const result = await reconcile(f, [first, { ...first, billing_log_recid: 'different-record' }]);
  assert.equal(result.outcome.status, 'FAILED');
  assert.deepEqual(fs.readFileSync(f.target), f.before);
  assert(audit(f).includes('Conflicting ConnectWise records'));
});

test('detects an invoice returned by two source projects before staging writes', async t => {
  const f = await setup(t, [record('INV-A'), record('INV-B', 202)]);
  const result = await reconcile(f, [source('INV-NEW'), source('INV-NEW', { project: 202 })]);
  assert.equal(result.outcome.summary.projectsFailed, 2);
  assert.equal(result.outcome.workbookWritten, false);
  assert.deepEqual(fs.readFileSync(f.target), f.before);
});

test('discards a failed project plan while writing a different successful project', async t => {
  const f = await setup(t, [record('INV-A'), record('INV-B', 202)]);
  const result = await reconcile(f, [source('INV-NEW'), source('INV-FAIL'), source('INV-B', { project: 202 })], { failInvoice: 'INV-FAIL' });
  assert.equal(result.outcome.status, 'COMPLETED_WITH_ERRORS');
  assert.equal(result.outcome.summary.projectsFailed, 1);
  assert.equal(result.outcome.summary.projectsSucceeded, 1);
  assert.equal(result.outcome.summary.rowsInserted, 0);
  assert.equal(result.rows.some(row => row.invoice === 'INV-NEW'), false);
  assert.equal(result.cell(2, 'Inv Amt (Pretax)').value ?? null, null);
  assert.equal(result.cell(3, 'Inv Amt (Pretax)').value, 1234.56);
  const csv = audit(f);
  assert(csv.includes('NOT_APPLIED'));
  assert(csv.includes('APPLIED'));
  assert(csv.includes('PROJECT_ERROR'));
});

test('an early project listing failure still produces a failed run audit', async t => {
  const f = await setup(t, [record('INV-A')]);
  const result = await reconcile(f, [], { failProject: 101 });
  assert.equal(result.outcome.status, 'FAILED');
  assert.deepEqual(fs.readFileSync(f.target), f.before);
  assert(audit(f).includes('Synthetic project listing failure'));
  assert(fs.readFileSync(result.outcome.summaryLogPath, 'utf8').includes('Projects failed: 1'));
});

for (const status of ['Unpaid', 'Partially Paid']) {
  test(`clears stale paid amount/date when ConnectWise now says ${status}`, async t => {
    const f = await setup(t, [record('INV-A', 101, { 'Paid Amt (Pretax)': 1000, 'Paid Date': OLD_DATE })]);
    const invoice = source('INV-A', { status });
    if (status === 'Partially Paid') invoice.date_paid = SOURCE_DATE;
    const result = await reconcile(f, [invoice]);
    assert.equal(result.cell(2, 'Paid Amt (Pretax)').value ?? null, null);
    assert.equal(result.cell(2, 'Paid Date').value ?? null, null);
    assert.equal(result.cell(2, 'Comm').value.formula, 'I2*A2');
    assert.equal(result.cell(2, 'Paid Amt (Pretax)').fill.fgColor.argb, 'FFFFFF00');
    assert(audit(f).includes('Paid Amt (Pretax)'));
    await assertPreservedParts(f.before, fs.readFileSync(f.target));
  });
}

test('replaces a nonblank paid date and paid amount with corrected ConnectWise data', async t => {
  const f = await setup(t, [record('INV-A', 101, { 'Paid Amt (Pretax)': 1000, 'Paid Date': OLD_DATE })]);
  const invoice = source('INV-A', { status: 'Paid' });
  invoice.date_paid = new Date(2026, 9, 18, 12).getTime();
  const result = await reconcile(f, [invoice]);
  assert.equal(result.cell(2, 'Paid Amt (Pretax)').value, 1234.56);
  assert.equal(result.cell(2, 'Paid Date').value.toISOString(), '2026-10-18T00:00:00.000Z');
  assert(audit(f).includes('Paid Date,9/1/2026,10/18/2026'));
  const first = fs.readFileSync(f.target);
  await reconcile(f, [invoice]);
  assert.deepEqual(fs.readFileSync(f.target), first);
});

test('explicitly empty source dates clear existing dates rather than preserving stale data', async t => {
  const f = await setup(t, [record('INV-A', 101, { 'Invoice Date': OLD_DATE, 'Paid Date': OLD_DATE })]);
  const invoice = source('INV-A', { status: 'Paid', date: null });
  invoice.date_paid = null;
  const result = await reconcile(f, [invoice]);
  assert.equal(result.cell(2, 'Invoice Date').value ?? null, null);
  assert.equal(result.cell(2, 'Paid Date').value ?? null, null);
  assert.equal(result.cell(2, 'Paid Amt (Pretax)').value, 1234.56);
});

test('missing balance or missing date field holds the project instead of inferring Paid or blank', async t => {
  const f = await setup(t, [record('INV-A', 101, { 'Paid Amt (Pretax)': 1000, 'Paid Date': OLD_DATE })]);
  for (const field of ['invoiceBalance', 'date_paid', 'date_invoice', 'paid_amount', 'invoice_amount']) {
    const invoice = source('INV-A', { status: 'Paid' });
    delete invoice[field];
    const result = await reconcile(f, [invoice]);
    assert.equal(result.outcome.status, 'FAILED');
    assert.deepEqual(fs.readFileSync(f.target), f.before);
  }
});

test('accepts an explicit empty paid amount as no payment when the balance is valid', async t => {
  const f = await setup(t, [record('INV-A')]);
  const invoice = source('INV-A');
  invoice.paid_amount = null;
  const result = await reconcile(f, [invoice]);
  assert.equal(result.outcome.status, 'COMPLETED');
  assert.equal(result.cell(2, 'Paid Amt (Pretax)').value ?? null, null);
});

test('updates an existing invoice changed to zero while skipping a new zero-dollar invoice', async t => {
  const f = await setup(t, [record('INV-A', 101, { 'Inv Amt (Pretax)': 1000, 'Paid Amt (Pretax)': 1000, 'Paid Date': OLD_DATE })]);
  const result = await reconcile(f, [source('INV-A', { amount: 0 }), source('INV-ZERO', { amount: 0 })]);
  assert.equal(result.cell(2, 'Inv Amt (Pretax)').value, 0);
  assert.equal(result.cell(2, 'Paid Amt (Pretax)').value ?? null, null);
  assert.equal(result.rows.some(row => row.invoice === 'INV-ZERO'), false);
});

test('saves audits and summaries for matching no-change runs without creating backups', async t => {
  const f = await setup(t, [record('INV-A', 101, { 'Inv Amt (Pretax)': 1234.56, 'Invoice Date': INVOICE_DATE })]);
  const first = await reconcile(f, [source('INV-A')]);
  const second = await reconcile(f, [source('INV-A')]);
  assert.equal(second.outcome.status, 'COMPLETED');
  assert.notEqual(first.outcome.auditLogPath, second.outcome.auditLogPath);
  assert(fs.existsSync(first.outcome.auditLogPath));
  assert(fs.existsSync(second.outcome.summaryLogPath));
  assert(audit(f).includes('RUN_SUMMARY'));
  assert(!fs.existsSync(path.join(f.dir, 'Backups')));
  assert.deepEqual(fs.readFileSync(f.target), f.before);
});

test('writer failures retain a backup and an audit distinguishing proposals from applied changes', async t => {
  const f = await setup(t, [record('INV-A', 101, { 'Inv Amt (Pretax)': { formula: '1+1', result: 2 } })]);
  await assert.rejects(reconcile(f, [source('INV-A')]), /Workbook update failed/);
  assert.deepEqual(fs.readFileSync(f.target), f.before);
  const csv = audit(f);
  assert(csv.includes('NOT_APPLIED'));
  assert(csv.includes('RUN_ERROR'));
  assert(csv.includes('FAILED'));
  const backup = fs.readdirSync(path.join(f.dir, 'Backups'))[0];
  assert.deepEqual(fs.readFileSync(path.join(f.dir, 'Backups', backup)), f.before);
  assert(!fs.readdirSync(f.dir).some(name => name.startsWith('commission-write-')));
});

test('audit records final row numbers after inserting above updated invoices', async t => {
  const f = await setup(t, [record('INV-A'), record('INV-B', 202)]);
  await reconcile(f, [source('INV-A'), source('INV-NEW'), source('INV-B', { project: 202 })]);
  const csv = audit(f);
  const update = csv.split('\n').find(line => line.includes('INV-B,4,Invoice Date'));
  assert(update, 'Expected final row 4 in audit');
  assert(update.trim().endsWith('APPLIED,3'), 'Expected original row 3 to remain available');
});

test('writer independently rejects a duplicate invoice in a fill or insertion payload', async t => {
  const f = await setup(t, [record('INV-A'), record(null)]);
  await assert.rejects(writeWorkbook({ workbookPath: f.target,
    blankRowFills: [{ rowNumber: 3, projectID: 101, invoice: 'inv-a', invoiceAmount: 1234.56, invoiceDate: INVOICE_DATE }] }), /duplicate invoice/);
  assert.deepEqual(fs.readFileSync(f.target), f.before);
});

test('menu reports failures per workbook and continues sequentially to later workbooks', async t => {
  await setup(t, [record('INV-A')]);
  const calls = [];
  const results = await processWorkbooks({}, [{ fileName: 'First.xlsx' }, { fileName: 'Second.xlsx' }], async (_, name) => {
    calls.push(name);
    if (name === 'First.xlsx') throw new Error('Synthetic workbook failure');
    return { workbookFilename: name, status: 'COMPLETED' };
  });
  assert.deepEqual(calls, ['First.xlsx', 'Second.xlsx']);
  assert.deepEqual(results.map(result => result.status), ['FAILED', 'COMPLETED']);
});

test('fetches subsequent ConnectWise pages rather than stopping at 250 invoices', async () => {
  const pages = [];
  const request = { post: async (_, options) => {
    const payload = JSON.parse(JSON.parse(new URLSearchParams(options.data).get('actionMessage')).payload);
    pages.push(payload.activePage);
    const count = payload.activePage === 1 ? 250 : 1;
    const rows = Array.from({ length: count }, (_, index) => ({ row: JSON.stringify({ billing_log_recid: `${payload.activePage}:${index}`, invoice_number: `INV-${payload.activePage}-${index}` }) }));
    return { ok: () => true, text: async () => JSON.stringify({ success: true, data: { isSuccess: true, action: { multilineRows: rows } } }) };
  } };
  const invoices = await getProjectInvoices(request, 101);
  assert.equal(invoices.length, 251);
  assert.deepEqual(pages, [1, 2]);
});

test('rejects a repeated full ConnectWise page instead of silently using an incomplete list', async () => {
  const rows = Array.from({ length: 250 }, (_, index) => ({ row: JSON.stringify({ billing_log_recid: index, invoice_number: `INV-${index}` }) }));
  const request = { post: async () => ({ ok: () => true, text: async () => JSON.stringify({ success: true, data: { isSuccess: true, action: { multilineRows: rows } } }) }) };
  await assert.rejects(getProjectInvoices(request, 101), /repeated an invoice page/);
});

test('selected-folder reconciliation writes only the selected copy, with backups/logs nearby and temp state per user', async t => {
  const f = await setup(t, [record('INV-A')]);
  const original = f.target;
  const selected = path.join(f.dir, 'Selected workbook folder'); fs.mkdirSync(selected);
  f.target = path.join(selected, path.basename(original));
  fs.copyFileSync(original, f.target);
  const state = path.join(f.dir, 'Private user data');
  await initializeWorkbookFolder({ store: createSettingsStore(state), pickFolder: async () => selected });
  const result = await reconcile(f, [source('INV-A')]);
  assert.equal(result.cell(2, 'Inv Amt (Pretax)').value, 1234.56);
  assert.deepEqual(fs.readFileSync(original), f.before);
  assert.equal(path.dirname(result.outcome.auditLogPath), path.join(selected, 'Logs'));
  assert.equal(path.dirname(result.outcome.backupPath), path.join(selected, 'Backups'));
  assert.equal(paths.PROFILE_DIR, path.join(state, 'connectwise-profile'));
  assert.deepEqual(fs.readdirSync(path.join(state, 'Temp')), []);
});
