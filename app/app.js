const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { randomUUID } = require('crypto');
const { CHANGE_ACTIONS, saveRunAudit, describeRunStatus } = require('./audit');
const console = require('./log').logger;

const paths = require('./paths');

const {
  getProjectInvoices,
  getInvoiceCalculations,
  getInvoicePayments,
  determinePaymentStatus,
} = require('./connectwise');

const {
  loadWorkbook,
  getWorksheet,
  getHeaderMap,
  getCellValue,
  PAID_HISTORY_REQUIRED_HEADERS,
  readProjects,
  groupRowsByProject,
} = require('./workbook');


function normalizeInvoiceNumber(value) {
  return String(value || '').trim().toUpperCase();
}


function hasValue(value) {
  return (
    value !== null &&
    value !== undefined &&
    String(value).trim() !== ''
  );
}

function toNumber(value) {
  if (!hasValue(value)) {
    return 0;
  }

  const number = Number(value);
  return Number.isNaN(number) ? 0 : number;
}

function amountsMatch(firstValue, secondValue) {
  return Math.abs(
    toNumber(firstValue) - toNumber(secondValue)
  ) < 0.005;
}

function formatCurrency(value) {
  if (!hasValue(value)) {
    return '';
  }

  const number = Number(value);

  if (Number.isNaN(number)) {
    return String(value);
  }

  return number.toLocaleString('en-US', {
    style: 'currency',
    currency: 'USD',
  });
}

function connectWiseDateToDate(value) {
  if (!value) {
    return null;
  }

  const date = new Date(Number(value));

  if (Number.isNaN(date.getTime())) {
    return null;
  }

  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate()
  );
}

// ExcelJS reads Excel date-only cells as UTC midnight. Keep invoice dates in
// that same representation so comparisons and audit dates survive a second run
// in Windows time zones west of UTC. ConnectWise timestamps follow the existing
// payment-date convention before their calendar date is stored in Excel.
function connectWiseInvoiceDateToDate(value) {
  const date = connectWiseDateToDate(value);
  return date
    ? new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()))
    : null;
}

function formatDate(value, useUtc = true) {
  if (!value) {
    return '';
  }

  const date =
    value instanceof Date
      ? value
      : new Date(value);

  if (Number.isNaN(date.getTime())) {
    return String(value);
  }

  return (
    `${(useUtc ? date.getUTCMonth() : date.getMonth()) + 1}/` +
    `${useUtc ? date.getUTCDate() : date.getDate()}/` +
    `${useUtc ? date.getUTCFullYear() : date.getFullYear()}`
  );
}

function formatTimestamp(date = new Date()) {
  const pad = (value) => String(value).padStart(2, '0');

  return (
    `${date.getFullYear()}-` +
    `${pad(date.getMonth() + 1)}-` +
    `${pad(date.getDate())}_` +
    `${pad(date.getHours())}-` +
    `${pad(date.getMinutes())}-` +
    `${pad(date.getSeconds())}`
  );
}

function csvEscape(value) {
  const text =
    value === null || value === undefined
      ? ''
      : String(value);

  if (
    text.includes(',') ||
    text.includes('"') ||
    text.includes('\n')
  ) {
    return `"${text.replace(/"/g, '""')}"`;
  }

  return text;
}

function cloneValue(value) {
  if (
    value === null ||
    value === undefined ||
    typeof value !== 'object'
  ) {
    return value;
  }

  if (value instanceof Date) {
    return new Date(value.getTime());
  }

  return JSON.parse(JSON.stringify(value));
}

function translateFormulaForNewRow(
  formula,
  sourceRowNumber,
  targetRowNumber
) {
  if (typeof formula !== 'string') {
    return formula;
  }

  const pattern =
    /(\$?[A-Z]{1,3})(\$?)(\d+)/g;

  return formula.replace(
    pattern,
    (match, column, absoluteRow, rowNumber) => {
      if (
        absoluteRow === '$' ||
        Number(rowNumber) !== sourceRowNumber
      ) {
        return match;
      }

      return `${column}${targetRowNumber}`;
    }
  );
}

function copyTemplateRow(
  sheet,
  sourceRowNumber,
  targetRowNumber
) {
  const sourceRow = sheet.getRow(sourceRowNumber);
  const targetRow = sheet.getRow(targetRowNumber);

  targetRow.height = sourceRow.height;
  targetRow.hidden = sourceRow.hidden;
  targetRow.outlineLevel = sourceRow.outlineLevel;

  const columnCount = Math.max(
    sheet.columnCount,
    sourceRow.cellCount
  );

  for (
    let columnNumber = 1;
    columnNumber <= columnCount;
    columnNumber += 1
  ) {
    const sourceCell =
      sourceRow.getCell(columnNumber);
    const targetCell =
      targetRow.getCell(columnNumber);

    const value = cloneValue(sourceCell.value);

    if (
      value &&
      typeof value === 'object' &&
      typeof value.formula === 'string'
    ) {
      value.formula =
        translateFormulaForNewRow(
          value.formula,
          sourceRowNumber,
          targetRowNumber
        );
    }

    targetCell.value = value;
    targetCell.style = cloneValue(
      sourceCell.style
    );

    if (sourceCell.numFmt) {
      targetCell.numFmt = sourceCell.numFmt;
    }
  }
}

function printInvoiceDetails(invoiceData) {
  console.log(`Invoice: ${invoiceData.invoiceNumber}`);
  console.log(`  Invoice Date:  ${formatDate(invoiceData.invoiceDate, true) || 'Unavailable'}`);
  console.log(
    `  Status:        ${invoiceData.status}`
  );
  console.log(
    `  Pretax Amount: ${formatCurrency(
      invoiceData.pretaxAmount
    )}`
  );
  console.log(
    `  Gross Paid:    ${formatCurrency(
      invoiceData.grossPaidAmount
    )}`
  );
  console.log(
    `  Paid Date:     ${
      formatDate(invoiceData.paidDate) || 'None'
    }`
  );
  console.log(
    `  Balance:       ${formatCurrency(
      invoiceData.balance
    )}`
  );
}

function addAuditEntry(auditEntries, entry) {
  auditEntries.push({
    timestamp: new Date().toISOString(),
    result: CHANGE_ACTIONS.has(entry.action) ? 'PLANNED' : entry.action.startsWith('REVIEW') ? 'REVIEW' : 'INFO',
    ...entry,
  });
}

function printSummary(summary) {
  console.log('');
  console.log(
    '=========================================='
  );
  console.log('Proposed Workbook Changes');
  console.log('');
  console.log(
    `Projects processed:          ${summary.projectsProcessed}`
  );
  console.log(`Projects completed / failed: ${summary.projectsSucceeded} / ${summary.projectsFailed}`);
  console.log(
    `Invoices reviewed:           ${summary.invoicesReviewed}`
  );
  console.log(
    `Existing rows updated:       ${summary.rowsUpdated}`
  );
  console.log(
    `Blank rows filled:           ${summary.blankRowsFilled}`
  );
  console.log(
    `New rows inserted:           ${summary.rowsInserted}`
  );
  console.log(
    `Workbook-only invoices:      ${summary.workbookOnlyInvoices}`
  );
  console.log(
    `Zero-dollar invoices ignored:${String(
      summary.zeroDollarInvoices
    ).padStart(7)}`
  );
  console.log(
    `Already paid invoices skipped:${String(
      summary.alreadyPaidSkipped
    ).padStart(6)}`
  );
  console.log(
    `Duplicate invoice warnings:  ${summary.duplicateWarnings}`
  );
  console.log(
    '=========================================='
  );
}

async function reconcileWorkbook(context, workbookFilename, { onProgress } = {}) {
  const progress = update => { try { onProgress?.(update); } catch { /* UI reporting cannot interrupt reconciliation. */ } };
  if (!context || !context.request) {
    throw new Error('An authenticated ConnectWise browser context is required.');
  }

  if (!workbookFilename) {
    throw new Error('No commission workbook was selected.');
  }

  const WORKBOOK_FILENAME = workbookFilename;

  const workbookPath = path.join(
    paths.WORKBOOK_DIR,
    WORKBOOK_FILENAME
  );

  const backupsDirectory = path.join(
    paths.BACKUP_DIR
  );

  const logsDirectory = path.join(
    paths.LOG_DIR
  );

  const runTimestamp = `${formatTimestamp()}-${randomUUID().slice(0, 8)}`;

  const backupPath = path.join(
    backupsDirectory,
    `${path.parse(WORKBOOK_FILENAME).name} - ${runTimestamp}.xlsx`
  );

  const auditLogPath = path.join(
    logsDirectory,
    `${path.parse(WORKBOOK_FILENAME).name} - Reconciliation - ${runTimestamp}.csv`
  );

  const auditEntries = [];

  const summary = {
    projectsProcessed: 0,
    projectsSucceeded: 0,
    projectsFailed: 0,
    invoicesReviewed: 0,
    rowsUpdated: 0,
    blankRowsFilled: 0,
    rowsInserted: 0,
    workbookOnlyInvoices: 0,
    zeroDollarInvoices: 0,
    duplicateWarnings: 0,
    alreadyPaidSkipped: 0,
  };

  let workbookWritten = false;
  let writeUncertain = false;
  let backupCreated = false;
  let fatalError = null;
  let writePayloadPath = null;
  const summaryLogPath = path.join(logsDirectory,
    `${path.parse(WORKBOOK_FILENAME).name} - Run Summary - ${runTimestamp}.txt`);
  const result = () => {
    const reviewCount = auditEntries.filter(entry => entry.result === 'REVIEW').length;
    return { workbookFilename, workbookWritten, writeUncertain, auditLogPath, summaryLogPath,
      backupPath: backupCreated ? backupPath : null, summary: { ...summary }, reviewCount,
      status: fatalError ? 'FAILED' : summary.projectsFailed
        ? (summary.projectsSucceeded ? 'COMPLETED_WITH_ERRORS' : 'FAILED')
        : reviewCount ? 'COMPLETED_WITH_REVIEW' : 'COMPLETED' };
  };
  fs.mkdirSync(logsDirectory, { recursive: true });
  try {
    const workbook = await loadWorkbook(
      workbookPath
    );

    const sheet = getWorksheet(
      workbook,
      'Commission'
    );

    const paidHistorySheets = workbook.worksheets.filter((candidateSheet) =>
      String(candidateSheet.name || '')
        .trim()
        .toLowerCase()
        .startsWith('paid to emp')
    );

    if (paidHistorySheets.length === 0) {
      throw new Error(
        'No worksheet beginning with "Paid to Emp" was found.'
      );
    }

    const alreadyPaidInvoices = new Map();
    const paidHistoryDuplicates = [];

    for (const paidSheet of paidHistorySheets) {
      const paidColumns = getHeaderMap(
        paidSheet,
        PAID_HISTORY_REQUIRED_HEADERS
      );

      for (
        let rowNumber = 2;
        rowNumber <= paidSheet.rowCount;
        rowNumber += 1
      ) {
        const row = paidSheet.getRow(rowNumber);
        const invoiceCell = row.getCell(paidColumns.get('Invoice #'));
        if (invoiceCell.value && typeof invoiceCell.value === 'object' &&
            ('formula' in invoiceCell.value || 'sharedFormula' in invoiceCell.value) &&
            (!Object.prototype.hasOwnProperty.call(invoiceCell.value, 'result') || invoiceCell.value.result === undefined)) {
          throw new Error(`Paid-history invoice formula has no cached value: ${paidSheet.name} row ${rowNumber}. Open and save it in Excel before reconciling.`);
        }
        const invoiceNumber = invoiceCell.text;
        const normalizedInvoice =
          normalizeInvoiceNumber(invoiceNumber);

        if (!normalizedInvoice) {
          continue;
        }

        const record = {
          invoiceNumber: String(invoiceNumber).trim(),
          paycheckDate: row.getCell(paidColumns.get('Paycheck Date')).value,
          rowNumber,
          sheetName: paidSheet.name,
        };

        if (!alreadyPaidInvoices.has(normalizedInvoice)) {
          alreadyPaidInvoices.set(normalizedInvoice, record);
        } else {
          paidHistoryDuplicates.push({
            normalizedInvoice,
            first: alreadyPaidInvoices.get(normalizedInvoice),
            duplicate: record,
          });
        }
      }
    }

    summary.duplicateWarnings = paidHistoryDuplicates.length;
    const workbookRows = readProjects(sheet);
    const projects = groupRowsByProject(workbookRows);
    // Include invoice cells even on rows with an invalid/missing project ID: they
    // still occupy an invoice number and must never be duplicated by insertion.
    const activeInvoices = new Map();
    const commissionColumns = getHeaderMap(sheet);
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const invoice = normalizeInvoiceNumber(row.getCell(commissionColumns.get('Invoice #')).text);
      if (!invoice) return;
      if (!activeInvoices.has(invoice)) activeInvoices.set(invoice, []);
      activeInvoices.get(invoice).push({ rowNumber,
        projectID: Number(getCellValue(row.getCell(commissionColumns.get('Project ID')))) });
    });
    for (const [invoice, locations] of activeInvoices) {
      if (locations.length > 1) {
        summary.duplicateWarnings++;
        addAuditEntry(auditEntries, { action: 'REVIEW_DUPLICATE', invoice,
          notes: `Invoice appears in multiple Commission rows: ${locations.map(r => r.rowNumber).join(', ')}. No automatic updates to this invoice.` });
      }
      if (alreadyPaidInvoices.has(invoice)) {
        addAuditEntry(auditEntries, { action: 'REVIEW_ACTIVE_PAID_OVERLAP', invoice,
          notes: 'Invoice is present in Commission and paid history. It will be skipped; review the active row manually.' });
      }
    }
    // Fetch project lists first so cross-project source conflicts are known before
    // staging any writes. These authenticated requests remain sequential.
    const projectLists = new Map();
    const sourceProjects = new Map();
    for (const project of projects) {
      progress({ stage: 'retrieving', current: projectLists.size + 1, total: projects.length, projectID: project.projectID });
      try {
        const invoices = await getProjectInvoices(context.request, project.projectID);
        projectLists.set(project.projectID, { invoices });
        for (const invoice of invoices) {
          const key = normalizeInvoiceNumber(invoice.invoice_number);
          if (!key) throw new Error('ConnectWise returned a blank invoice number.');
          if (!sourceProjects.has(key)) sourceProjects.set(key, new Set());
          sourceProjects.get(key).add(project.projectID);
        }
      } catch (error) { projectLists.set(project.projectID, { error }); }
    }


    const existingRowUpdates = [];
    const blankRowFills = [];
    const newRowInserts = [];

    console.log('');
    console.log(
      `Paid-history sheets checked: ${paidHistorySheets
        .map((paidSheet) => paidSheet.name)
        .join(', ')}`
    );

    for (const duplicate of paidHistoryDuplicates) {
      addAuditEntry(auditEntries, { action: 'REVIEW_PAID_HISTORY_DUPLICATE',
        invoice: duplicate.duplicate.invoiceNumber,
        notes: `Paid-history duplicate in ${duplicate.first.sheetName} row ${duplicate.first.rowNumber} and ${duplicate.duplicate.sheetName} row ${duplicate.duplicate.rowNumber}; excluded from reconciliation.` });
      console.warn(
        `WARNING: Invoice ${duplicate.duplicate.invoiceNumber} appears more than once in paid history: ` +
          `${duplicate.first.sheetName} row ${duplicate.first.rowNumber} and ` +
          `${duplicate.duplicate.sheetName} row ${duplicate.duplicate.rowNumber}. ` +
          'It will still be treated as already paid and skipped.'
      );
    }
    console.log('');
    console.log(
      `Reconciling ${projects.length} projects...`
    );
    console.log('');

    for (
      let projectIndex = 0;
      projectIndex < projects.length;
      projectIndex += 1
    ) {
      const project = projects[projectIndex];
      progress({ stage: 'reconciling', current: projectIndex + 1, total: projects.length, projectID: project.projectID });
      summary.projectsProcessed += 1;

      console.log(
        '=================================================='
      );
      console.log(
        `[${projectIndex + 1}/${projects.length}] ` +
        `Project ${project.projectID}`
      );
      console.log(project.company);
      console.log(project.description);
      console.log('');

      const checkpoint = { updates: existingRowUpdates.length, fills: blankRowFills.length,
        inserts: newRowInserts.length, audit: auditEntries.length,
        rowsUpdated: summary.rowsUpdated, blankRowsFilled: summary.blankRowsFilled,
        rowsInserted: summary.rowsInserted };
      try {
        const listing = projectLists.get(project.projectID);
        if (listing.error) throw listing.error;
        const connectWiseInvoices = [];
        const seen = new Map();
        for (const invoice of listing.invoices) {
          const key = normalizeInvoiceNumber(invoice.invoice_number);
          if (sourceProjects.get(key).size > 1) {
            throw new Error(`Invoice ${key} was returned for multiple projects; no changes from this project will be written.`);
          }
          const existing = seen.get(key);
          if (existing) {
            if (String(existing.billing_log_recid) !== String(invoice.billing_log_recid) ||
                ['invoice_amount', 'paid_amount', 'date_invoice', 'date_paid'].some(field => existing[field] !== invoice[field])) {
              throw new Error(`Conflicting ConnectWise records for invoice ${key}; no changes from this project will be written.`);
            }
            addAuditEntry(auditEntries, { action: 'SKIP_DUPLICATE_SOURCE', projectID: project.projectID,
              invoice: key, notes: 'Repeated identical ConnectWise invoice record; processed once.' });
            continue;
          }
          seen.set(key, invoice);
          connectWiseInvoices.push(invoice);
        }

        console.log('');

        const workbookInvoiceRows =
          new Map();

        const blankWorkbookRows = [];

        for (const row of project.workbookRows) {
          const normalizedInvoice =
            normalizeInvoiceNumber(
              row.invoice
            );

          if (!normalizedInvoice) {
            blankWorkbookRows.push(row);
            continue;
          }

          if (!workbookInvoiceRows.has(normalizedInvoice)) workbookInvoiceRows.set(normalizedInvoice, row);
        }

        const connectWiseInvoiceNumbers =
          new Set();

        if (connectWiseInvoices.length === 0) {
          console.log(
            'ConnectWise: No invoices found.'
          );
          console.log('');
        }

        for (
          const invoice of connectWiseInvoices
        ) {
          summary.invoicesReviewed += 1;

          const invoiceNumber = String(
            invoice.invoice_number || ''
          ).trim();

          const normalizedInvoice =
            normalizeInvoiceNumber(
              invoiceNumber
            );

          if (normalizedInvoice) {
            connectWiseInvoiceNumbers.add(
              normalizedInvoice
            );
          }

          const alreadyPaidRecord =
            alreadyPaidInvoices.get(
              normalizedInvoice
            );

          if (alreadyPaidRecord) {
            summary.alreadyPaidSkipped += 1;

            console.log(
              `  ACTION: SKIP — invoice already exists on ${alreadyPaidRecord.sheetName}`
            );
            console.log(
              `  Paid-history location: ${alreadyPaidRecord.sheetName} row ${alreadyPaidRecord.rowNumber}`
            );
            console.log(
              `  Paycheck Date: ${
                formatDate(alreadyPaidRecord.paycheckDate) ||
                'Not recorded'
              }`
            );
            console.log('');

            addAuditEntry(auditEntries, {
              action: 'SKIP_ALREADY_PAID',
              projectID: project.projectID,
              company: project.company,
              invoice: invoiceNumber,
              rowNumber: alreadyPaidRecord.rowNumber,
              field: alreadyPaidRecord.sheetName,
              oldValue: '',
              newValue: '',
              notes:
                `Already present on ${alreadyPaidRecord.sheetName}; paycheck date ${
                  formatDate(alreadyPaidRecord.paycheckDate) ||
                  'not recorded'
                }`,
            });

            continue;
          }

          const billingLogRecID =
            invoice.billing_log_recid;

          const [calculations, payments] =
            await Promise.all([
              getInvoiceCalculations(
                context.request,
                billingLogRecID
              ),
              getInvoicePayments(
                context.request,
                billingLogRecID
              ),
            ]);

          const status =
            determinePaymentStatus(
              invoice,
              payments
            );

          const pretaxAmount = Number(
            calculations.invoiceSubTotal || 0
          );

          if (!hasValue(calculations.invoiceSubTotal) ||
              !['number', 'string'].includes(typeof calculations.invoiceSubTotal) || !Number.isFinite(pretaxAmount)) {
            throw new Error(`Invoice ${invoiceNumber} has no valid ConnectWise pretax subtotal; no values will be inferred for it.`);
          }

          const readSourceDate = field => {
            if (!Object.prototype.hasOwnProperty.call(invoice, field))
              throw new Error(`Invoice ${invoiceNumber} is missing the ConnectWise ${field} field; the project is held for review.`);
            const value = invoice[field];
            if (value === null || (typeof value === 'string' && value.trim() === '')) return null;
            const date = (typeof value === 'number' || typeof value === 'string')
              ? connectWiseInvoiceDateToDate(value) : null;
            if (!date) throw new Error(`Invoice ${invoiceNumber} has invalid ConnectWise ${field}; the project is held for review.`);
            return date;
          };
          const invoiceDate = readSourceDate('date_invoice');

          const grossPaidAmount = Number(
            invoice.paid_amount || 0
          );

          const paidDate = readSourceDate('date_paid');
          const desiredPaidAmount = status === 'Paid' ? pretaxAmount : null;
          const desiredPaidDate = status === 'Paid' ? paidDate : null;

          const balance = Number(
            payments.invoiceBalance || 0
          );

          const invoiceData = {
            invoiceNumber,
            normalizedInvoice,
            status,
            pretaxAmount,
            invoiceDate,
            grossPaidAmount,
            paidDate,
            balance,
          };

          printInvoiceDetails(invoiceData);

          if (
            Math.abs(pretaxAmount) < 0.005 && !activeInvoices.has(normalizedInvoice)
          ) {
            summary.zeroDollarInvoices += 1;

            console.log(
              '  ACTION: IGNORE — zero-dollar invoice'
            );
            console.log('');

            addAuditEntry(auditEntries, {
              action: 'IGNORE_ZERO_DOLLAR',
              projectID: project.projectID,
              company: project.company,
              invoice: invoiceNumber,
              rowNumber: '',
              field: '',
              oldValue: '',
              newValue: '',
              notes: 'Zero-dollar invoice',
            });

            continue;
          }

          if (
            (activeInvoices.get(normalizedInvoice) || []).length > 1
          ) {
            console.log(
              '  ACTION: REVIEW — duplicate workbook invoice prevents automatic update'
            );
            console.log('');
            continue;
          }

          const locations = activeInvoices.get(normalizedInvoice) || [];
          if (locations.length && locations[0].projectID !== project.projectID) {
            throw new Error(`Invoice ${invoiceNumber} already exists in Commission row ${locations[0].rowNumber} under another or invalid project; it will not be inserted again.`);
          }
          const existingRow = workbookInvoiceRows.get(normalizedInvoice);

          if (existingRow) {
            const fieldChanges = [];

            if (!hasValue(existingRow.invoiceAmount) || !amountsMatch(existingRow.invoiceAmount, pretaxAmount)) {
              fieldChanges.push({
                field: 'Inv Amt (Pretax)',
                oldValue: existingRow.invoiceAmount,
                newValue: pretaxAmount,
              });
            }

            if (formatDate(existingRow.invoiceDate, true) !== formatDate(invoiceDate, true)) {
              fieldChanges.push({
                field: 'Invoice Date',
                oldValue: existingRow.invoiceDate,
                newValue: invoiceDate,
              });
            }

            if (desiredPaidAmount === null ? hasValue(existingRow.paidAmount)
              : !hasValue(existingRow.paidAmount) || !amountsMatch(existingRow.paidAmount, desiredPaidAmount)) {
              fieldChanges.push({ field: 'Paid Amt (Pretax)', oldValue: existingRow.paidAmount, newValue: desiredPaidAmount });
            }
            if (formatDate(existingRow.paidDate, true) !== formatDate(desiredPaidDate, true)) {
              fieldChanges.push({ field: 'Paid Date', oldValue: existingRow.paidDate, newValue: desiredPaidDate });
            }

            if (fieldChanges.length === 0) {
              console.log(
                `  ACTION: NONE — matches workbook row ${existingRow.rowNumber}`
              );
            } else {
              existingRowUpdates.push({
                projectID:
                  project.projectID,
                company: project.company,
                invoice: invoiceNumber,
                rowNumber:
                  existingRow.rowNumber,
                changes: fieldChanges,
              });

              summary.rowsUpdated += 1;

              console.log(
                `  ACTION: UPDATE workbook row ${existingRow.rowNumber}`
              );

              for (
                const change of fieldChanges
              ) {
                const isDate = change.field === 'Paid Date' || change.field === 'Invoice Date';
                const oldDisplay =
                  isDate
                    ? formatDate(
                        change.oldValue,
                        true
                      ) || 'blank'
                    : formatCurrency(
                        change.oldValue
                      ) || 'blank';

                const newDisplay =
                  isDate
                    ? formatDate(
                        change.newValue,
                        true
                      )
                    : formatCurrency(
                        change.newValue
                      );

                console.log(
                  `    ${change.field}: ${oldDisplay} → ${newDisplay}`
                );

                addAuditEntry(
                  auditEntries,
                  {
                    action:
                      'UPDATE_EXISTING_ROW',
                    projectID:
                      project.projectID,
                    company:
                      project.company,
                    invoice:
                      invoiceNumber,
                    rowNumber:
                      existingRow.rowNumber,
                    field:
                      change.field,
                    oldValue:
                      oldDisplay,
                    newValue:
                      newDisplay,
                    notes: '',
                  }
                );
              }
            }

            console.log('');
            continue;
          }

          const blankRow =
            blankWorkbookRows.shift();

          const paidAmount =
            status === 'Paid'
              ? pretaxAmount
              : null;

          const workbookPaidDate =
            status === 'Paid'
              ? paidDate
              : null;

          if (blankRow) {
            blankRowFills.push({
              projectID:
                project.projectID,
              company: project.company,
              invoice: invoiceNumber,
              rowNumber:
                blankRow.rowNumber,
              invoiceAmount: pretaxAmount,
              invoiceDate,
              paidAmount,
              paidDate:
                workbookPaidDate,
            });

            summary.blankRowsFilled += 1;

            console.log(
              `  ACTION: FILL BLANK workbook row ${blankRow.rowNumber}`
            );
            console.log(
              `    Invoice: blank → ${invoiceNumber}`
            );
            console.log(`    Inv Amt (Pretax): ${formatCurrency(pretaxAmount)}`);
            console.log(`    Invoice Date: ${formatDate(invoiceDate, true) || 'blank'}`);
            console.log(
              `    Paid Amt (Pretax): ${
                paidAmount === null
                  ? 'blank'
                  : formatCurrency(
                      paidAmount
                    )
              }`
            );
            console.log(
              `    Paid Date: ${
                workbookPaidDate
                  ? formatDate(
                      workbookPaidDate, true
                    )
                  : 'blank'
              }`
            );

            addAuditEntry(auditEntries, {
              action: 'FILL_BLANK_ROW',
              projectID:
                project.projectID,
              company: project.company,
              invoice: invoiceNumber,
              rowNumber:
                blankRow.rowNumber,
              field:
                'Invoice / Inv Amt (Pretax) / Invoice Date / Paid Amt / Paid Date',
              oldValue:
                `blank; ${formatCurrency(blankRow.invoiceAmount) || 'blank'}; ` +
                `${formatDate(blankRow.invoiceDate, true) || 'blank'}; ` +
                `${formatCurrency(blankRow.paidAmount) || 'blank'}; ` +
                `${formatDate(blankRow.paidDate, true) || 'blank'}`,
              newValue:
                `${invoiceNumber}; ` +
                `${formatCurrency(pretaxAmount)}; ` +
                `${formatDate(invoiceDate, true) || 'blank'}; ` +
                `${
                  paidAmount === null
                    ? 'blank'
                    : formatCurrency(
                        paidAmount
                      )
                }; ` +
                `${
                  workbookPaidDate
                    ? formatDate(
                        workbookPaidDate
                      )
                    : 'blank'
                }`,
              notes: '',
            });
          } else {
            const lastProjectRowNumber =
              Math.max(
                ...project.workbookRows.map(
                  (row) => row.rowNumber
                )
              );

            newRowInserts.push({
              projectID:
                project.projectID,
              company: project.company,
              description:
                project.description,
              invoice: invoiceNumber,
              invoiceAmount: pretaxAmount,
              invoiceDate,
              paidAmount,
              paidDate:
                workbookPaidDate,
              afterRowNumber:
                lastProjectRowNumber,
            });

            summary.rowsInserted += 1;

            console.log(
              '  ACTION: ADD NEW ROW — no blank project row is available'
            );
            console.log(
              `    Invoice: ${invoiceNumber}`
            );
            console.log(`    Inv Amt (Pretax): ${formatCurrency(pretaxAmount)}`);
            console.log(`    Invoice Date: ${formatDate(invoiceDate, true) || 'blank'}`);
            console.log(
              `    Paid Amt (Pretax): ${
                paidAmount === null
                  ? 'blank'
                  : formatCurrency(
                      paidAmount
                    )
              }`
            );
            console.log(
              `    Paid Date: ${
                workbookPaidDate
                  ? formatDate(
                      workbookPaidDate, true
                    )
                  : 'blank'
              }`
            );

            addAuditEntry(auditEntries, {
              action: 'INSERT_NEW_ROW',
              projectID:
                project.projectID,
              company: project.company,
              invoice: invoiceNumber,
              rowNumber:
                `after ${lastProjectRowNumber}`,
              field:
                'Invoice / Inv Amt (Pretax) / Invoice Date / Paid Amt / Paid Date',
              oldValue: '',
              newValue:
                `${invoiceNumber}; ` +
                `${formatCurrency(pretaxAmount)}; ` +
                `${formatDate(invoiceDate, true) || 'blank'}; ` +
                `${
                  paidAmount === null
                    ? 'blank'
                    : formatCurrency(
                        paidAmount
                      )
                }; ` +
                `${
                  workbookPaidDate
                    ? formatDate(
                        workbookPaidDate
                      )
                    : 'blank'
                }`,
              notes: '',
            });
          }

          console.log('');
        }

        for (const row of project.workbookRows) {
          const normalizedWorkbookInvoice =
            normalizeInvoiceNumber(
              row.invoice
            );

          if (
            normalizedWorkbookInvoice &&
            !connectWiseInvoiceNumbers.has(
              normalizedWorkbookInvoice
            )
          ) {
            summary.workbookOnlyInvoices += 1;

            console.log(
              `Workbook-only invoice: ${row.invoice}`
            );
            console.log(
              '  ACTION: REVIEW — not returned by ConnectWise'
            );
            console.log('');

            addAuditEntry(auditEntries, {
              action: 'REVIEW_WORKBOOK_ONLY',
              projectID:
                project.projectID,
              company: project.company,
              invoice: row.invoice,
              rowNumber: row.rowNumber,
              field: '',
              oldValue: '',
              newValue: '',
              notes:
                'Invoice not returned by ConnectWise',
            });
          }
        }
        summary.projectsSucceeded++;
      } catch (error) {
        existingRowUpdates.length = checkpoint.updates;
        blankRowFills.length = checkpoint.fills;
        newRowInserts.length = checkpoint.inserts;
        summary.rowsUpdated = checkpoint.rowsUpdated;
        summary.blankRowsFilled = checkpoint.blankRowsFilled;
        summary.rowsInserted = checkpoint.rowsInserted;
        summary.projectsFailed++;
        for (const entry of auditEntries.slice(checkpoint.audit)) {
          if (CHANGE_ACTIONS.has(entry.action)) {
            entry.result = 'NOT_APPLIED';
            entry.notes = 'Project failed; this proposed change was discarded.';
          }
        }
        console.error(
          `Project ${project.projectID} failed:`
        );
        console.error(error.message);
        console.log('');

        addAuditEntry(auditEntries, {
          action: 'PROJECT_ERROR',
          projectID:
            project.projectID,
          company: project.company,
          invoice: '',
          rowNumber: '',
          field: '',
          oldValue: '',
          newValue: '',
          notes: error.message,
        });
      }
    }

    printSummary(summary);

    const totalWorkbookChanges =
      summary.rowsUpdated +
      summary.blankRowsFilled +
      summary.rowsInserted;

    if (totalWorkbookChanges === 0) {
      console.log('');
      console.log(
        'No workbook changes are required.'
      );
      if (summary.projectsFailed) console.error('Failed projects need attention; no changes from those projects were written.');
      return result();
    }

    console.log('');
    console.log(
      `${totalWorkbookChanges} workbook row change(s) will be made.`
    );

    console.log(
      'Creating backup, applying changes, and writing the audit log...'
    );

    fs.mkdirSync(backupsDirectory, {
      recursive: true,
    });

    fs.mkdirSync(logsDirectory, {
      recursive: true,
    });

    fs.copyFileSync(
      workbookPath,
      backupPath,
      fs.constants.COPYFILE_EXCL
    );
    backupCreated = true;

    const insertGroups = Array.from(
      newRowInserts.reduce((groups, insert) => {
        if (!groups.has(insert.projectID)) {
          groups.set(insert.projectID, {
            projectID: insert.projectID,
            afterRowNumber: insert.afterRowNumber,
            inserts: [],
          });
        }

        groups.get(insert.projectID).inserts.push(insert);
        return groups;
      }, new Map()).values()
    );

    // Audit the final row addresses after insertion, retaining source addresses
    // so an employee can compare the saved workbook with its original backup.
    for (const entry of auditEntries) {
      if (entry.result !== 'PLANNED' && entry.action !== 'REVIEW_WORKBOOK_ONLY') continue;
      entry.originalRowNumber = entry.rowNumber;
      if (entry.action === 'INSERT_NEW_ROW') {
        const index = newRowInserts.findIndex(insert => insert.projectID === entry.projectID &&
          normalizeInvoiceNumber(insert.invoice) === normalizeInvoiceNumber(entry.invoice));
        const insert = newRowInserts[index];
        entry.rowNumber = insert.afterRowNumber + 1 + newRowInserts.filter((other, otherIndex) =>
          other.afterRowNumber < insert.afterRowNumber ||
          (other.afterRowNumber === insert.afterRowNumber && otherIndex < index)).length;
      } else if (Number.isInteger(entry.rowNumber)) {
        entry.rowNumber += newRowInserts.filter(insert => insert.afterRowNumber < entry.originalRowNumber).length;
      }
    }

    writePayloadPath = path.join(
      paths.TEMP_DIR,
      `commission-write-${runTimestamp}.json`
    );

    const writePayload = {
      workbookPath,
      backupPath,
      existingRowUpdates,
      blankRowFills,
      insertGroups,
    };

    fs.writeFileSync(
      writePayloadPath,
      JSON.stringify(writePayload, null, 2),
      'utf8'
    );

    // Write the XLSX package directly. Microsoft Excel COM automation can
    // block indefinitely on this workbook even when Excel opens successfully.
    // This writer touches only the Commission XML, cell styles, and calculation
    // flags; it preserves historical paid sheets byte-for-byte.
    const writerScriptPath = path.join(__dirname, 'xlsx-writer.js');
    const writerTracePath = path.join(
      logsDirectory,
      `${path.parse(WORKBOOK_FILENAME).name} - Offline Writer - ${runTimestamp}.log`
    );

    console.log('');
    console.log('Applying workbook changes without Microsoft Excel...');

    saveRunAudit({ ...result(), status: 'PLANNED' }, auditEntries);
    progress({ stage: 'saving' });
    const writerResult = spawnSync(
      process.execPath,
      [writerScriptPath, writePayloadPath, writerTracePath],
      {
        encoding: 'utf8',
        windowsHide: true,
        timeout: 120000,
      }
    );

    if (writerResult.error || writerResult.status !== 0) {
      try {
        workbookWritten = !fs.readFileSync(workbookPath).equals(fs.readFileSync(backupPath));
        writeUncertain = workbookWritten;
      } catch {
        writeUncertain = true;
      }
      const details = [
        writerResult.error && writerResult.error.message,
        writerResult.stderr && writerResult.stderr.trim(),
        writerResult.stdout && writerResult.stdout.trim(),
        writerResult.signal && `Signal: ${writerResult.signal}`,
        `Exit status: ${writerResult.status === null ? 'not available' : writerResult.status}`,
        `Offline writer log: ${writerTracePath}`,
      ].filter(Boolean);

      throw new Error(
        `Workbook update failed. Check the workbook and its pre-write backup before retrying.\n${details.join('\n')}`
      );
    }

    workbookWritten = true;
    if (writerResult.stderr && writerResult.stderr.trim()) {
      console.warn(writerResult.stderr.trim());
      addAuditEntry(auditEntries, { action: 'REVIEW_WRITER_LOG', notes: writerResult.stderr.trim() });
    }
    if (writerResult.stdout && writerResult.stdout.trim()) {
      console.log(writerResult.stdout.trim());
    }
    console.log(`Offline writer log: ${writerTracePath}`);

    console.log(`Workbook updated: ${workbookPath}`);
    console.log(`Backup created: ${backupPath}`);
    return result();
  } catch (error) {
    fatalError = error;
    error.auditLogPath = auditLogPath;
    error.summaryLogPath = summaryLogPath;
    error.runResult = result();
    addAuditEntry(auditEntries, { action: 'RUN_ERROR', result: 'FAILED', notes: error.message });
    throw error;
  } finally {
    try {
      if (writePayloadPath && fs.existsSync(writePayloadPath)) fs.unlinkSync(writePayloadPath);
    } catch (cleanupError) {
      addAuditEntry(auditEntries, { action: 'REVIEW_TEMP_FILE', notes: `Could not remove temporary write plan: ${cleanupError.message}` });
    }
    for (const entry of auditEntries) {
      if (entry.result === 'PLANNED') entry.result = writeUncertain ? 'UNCONFIRMED' : workbookWritten ? 'APPLIED' : 'NOT_APPLIED';
    }
    const outcome = result();
    if (fatalError) fatalError.runResult = outcome;
    addAuditEntry(auditEntries, { action: 'RUN_SUMMARY', result: outcome.status,
      notes: `${summary.projectsSucceeded} completed project(s), ${summary.projectsFailed} failed; workbook ${writeUncertain ? 'write outcome unconfirmed; compare with backup' : workbookWritten ? 'updated' : 'unchanged'}.` });
    try {
      saveRunAudit(outcome, auditEntries);
      console.log(`Run result: ${describeRunStatus(outcome.status)}`);
      console.log(`Audit log: ${auditLogPath}`);
      console.log(`Run summary: ${summaryLogPath}`);
    } catch (auditError) {
      console.error(`Could not finalize the audit: ${auditError.message}`);
      if (!fatalError) throw Object.assign(new Error(`Workbook ${workbookWritten ? 'was saved' : 'was not changed'}, but its final audit could not be saved: ${auditError.message}`), { auditLogPath, summaryLogPath, runResult: outcome });
    }
  }
}

module.exports = { reconcileWorkbook };
