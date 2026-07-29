const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const {
  getProjectInvoices,
  getInvoiceCalculations,
  getInvoicePayments,
  determinePaymentStatus,
} = require('./connectwise');

const {
  loadWorkbook,
  getWorksheet,
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

function formatDate(value) {
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
    `${date.getMonth() + 1}/` +
    `${date.getDate()}/` +
    `${date.getFullYear()}`
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

async function reconcileWorkbook(context, workbookFilename) {
  if (!context || !context.request) {
    throw new Error('An authenticated ConnectWise browser context is required.');
  }

  if (!workbookFilename) {
    throw new Error('No commission workbook was selected.');
  }

  const WORKBOOK_FILENAME = workbookFilename;

  const workbookPath = path.join(
    __dirname,
    WORKBOOK_FILENAME
  );

  const backupsDirectory = path.join(
    __dirname,
    'Backups'
  );

  const logsDirectory = path.join(
    __dirname,
    'Logs'
  );

  const runTimestamp = formatTimestamp();

  const backupPath = path.join(
    backupsDirectory,
    `${path.parse(WORKBOOK_FILENAME).name} - ${runTimestamp}.xlsx`
  );

  const auditLogPath = path.join(
    logsDirectory,
    `${path.parse(WORKBOOK_FILENAME).name} - Reconciliation - ${runTimestamp}.csv`
  );

  const temporaryWorkbookPath = path.join(
    __dirname,
    `${path.parse(WORKBOOK_FILENAME).name}.working.xlsx`
  );

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
    for (
      let rowNumber = 2;
      rowNumber <= paidSheet.actualRowCount;
      rowNumber += 1
    ) {
      const row = paidSheet.getRow(rowNumber);
      const invoiceNumber = row.getCell(5).value;
      const normalizedInvoice =
        normalizeInvoiceNumber(invoiceNumber);

      if (!normalizedInvoice) {
        continue;
      }

      const record = {
        invoiceNumber: String(invoiceNumber).trim(),
        paycheckDate: row.getCell(9).value,
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

  const workbookRows = readProjects(sheet);
  const projects = groupRowsByProject(
    workbookRows
  );

  const existingRowUpdates = [];
  const blankRowFills = [];
  const newRowInserts = [];
  const auditEntries = [];

  const summary = {
    projectsProcessed: 0,
    invoicesReviewed: 0,
    rowsUpdated: 0,
    blankRowsFilled: 0,
    rowsInserted: 0,
    workbookOnlyInvoices: 0,
    zeroDollarInvoices: 0,
    duplicateWarnings: paidHistoryDuplicates.length,
    alreadyPaidSkipped: 0,
  };

  try {
    console.log('');
    console.log(
      `Paid-history sheets checked: ${paidHistorySheets
        .map((paidSheet) => paidSheet.name)
        .join(', ')}`
    );

    for (const duplicate of paidHistoryDuplicates) {
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

      try {
        const connectWiseInvoices =
          await getProjectInvoices(
            context.request,
            project.projectID
          );
        console.log('');

        const workbookInvoiceRows =
          new Map();

        const duplicateWorkbookInvoices =
          new Set();

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

          if (
            workbookInvoiceRows.has(
              normalizedInvoice
            )
          ) {
            duplicateWorkbookInvoices.add(
              normalizedInvoice
            );
          } else {
            workbookInvoiceRows.set(
              normalizedInvoice,
              row
            );
          }
        }

        for (
          const duplicateInvoice of
          duplicateWorkbookInvoices
        ) {
          summary.duplicateWarnings += 1;

          console.log(
            `Workbook duplicate invoice: ${duplicateInvoice}`
          );
          console.log(
            '  ACTION: REVIEW — duplicate invoice number'
          );
          console.log('');

          addAuditEntry(auditEntries, {
            action: 'REVIEW_DUPLICATE',
            projectID: project.projectID,
            company: project.company,
            invoice: duplicateInvoice,
            rowNumber: '',
            field: '',
            oldValue: '',
            newValue: '',
            notes:
              'Duplicate invoice number in workbook',
          });
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

          const grossPaidAmount = Number(
            invoice.paid_amount || 0
          );

          const paidDate =
            connectWiseDateToDate(
              invoice.date_paid
            );

          const balance = Number(
            payments.invoiceBalance || 0
          );

          const invoiceData = {
            invoiceNumber,
            normalizedInvoice,
            status,
            pretaxAmount,
            grossPaidAmount,
            paidDate,
            balance,
          };

          printInvoiceDetails(invoiceData);

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

          if (
            Math.abs(pretaxAmount) < 0.005
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
            duplicateWorkbookInvoices.has(
              normalizedInvoice
            )
          ) {
            console.log(
              '  ACTION: REVIEW — duplicate workbook invoice prevents automatic update'
            );
            console.log('');
            continue;
          }

          const existingRow =
            workbookInvoiceRows.get(
              normalizedInvoice
            );

          if (existingRow) {
            const fieldChanges = [];

            if (
              status === 'Paid' &&
              (
                !hasValue(
                  existingRow.paidAmount
                ) ||
                !amountsMatch(
                  existingRow.paidAmount,
                  pretaxAmount
                )
              )
            ) {
              fieldChanges.push({
                column: 7,
                field: 'Paid Amt (Pretax)',
                oldValue:
                  existingRow.paidAmount,
                newValue: pretaxAmount,
              });
            }

            if (
              status === 'Paid' &&
              !hasValue(existingRow.paidDate) &&
              paidDate
            ) {
              fieldChanges.push({
                column: 8,
                field: 'Paid Date',
                oldValue:
                  existingRow.paidDate,
                newValue: paidDate,
              });
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
                const oldDisplay =
                  change.field ===
                  'Paid Date'
                    ? formatDate(
                        change.oldValue
                      ) || 'blank'
                    : formatCurrency(
                        change.oldValue
                      ) || 'blank';

                const newDisplay =
                  change.field ===
                  'Paid Date'
                    ? formatDate(
                        change.newValue
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
                      workbookPaidDate
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
                'Invoice / Paid Amt / Paid Date',
              oldValue: 'blank',
              newValue:
                `${invoiceNumber}; ` +
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
                      workbookPaidDate
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
                'Invoice / Paid Amt / Paid Date',
              oldValue: '',
              newValue:
                `${invoiceNumber}; ` +
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
      } catch (error) {
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
      return;
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
      backupPath
    );

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

    const writePayloadPath = path.join(
      __dirname,
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

    const writerScriptPath = path.join(
      __dirname,
      'apply-workbook-changes.ps1'
    );

    console.log('');
    console.log(
      'Applying workbook changes through Microsoft Excel...'
    );

    const writerResult = spawnSync(
      'powershell.exe',
      [
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        writerScriptPath,
        '-PayloadPath',
        writePayloadPath,
      ],
      {
        encoding: 'utf8',
        windowsHide: true,
      }
    );

    if (fs.existsSync(writePayloadPath)) {
      fs.unlinkSync(writePayloadPath);
    }

    if (writerResult.status !== 0) {
      const writerMessage =
        writerResult.stderr ||
        writerResult.stdout ||
        'Unknown Excel write error.';

      throw new Error(
        `Workbook update failed. The original workbook was not intentionally replaced.\n${writerMessage.trim()}`
      );
    }

    if (writerResult.stdout.trim()) {
      console.log(writerResult.stdout.trim());
    }

    const auditHeaders = [
      'Timestamp',
      'Action',
      'Project ID',
      'Company',
      'Invoice',
      'Workbook Row',
      'Field',
      'Old Value',
      'New Value',
      'Notes',
    ];

    const auditLines = [
      auditHeaders
        .map(csvEscape)
        .join(','),
    ];

    for (const entry of auditEntries) {
      auditLines.push(
        [
          entry.timestamp,
          entry.action,
          entry.projectID,
          entry.company,
          entry.invoice,
          entry.rowNumber,
          entry.field,
          entry.oldValue,
          entry.newValue,
          entry.notes,
        ]
          .map(csvEscape)
          .join(',')
      );
    }

    fs.writeFileSync(
      auditLogPath,
      auditLines.join('\r\n'),
      'utf8'
    );

    console.log('');
    console.log(
      '=========================================='
    );
    console.log(
      'Commission Reconciliation Complete'
    );
    console.log('');
    console.log(
      `Projects processed:          ${summary.projectsProcessed}`
    );
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
    console.log('');
    console.log(
      `Workbook updated:\n${workbookPath}`
    );
    console.log('');
    console.log(
      `Backup created:\n${backupPath}`
    );
    console.log('');
    console.log(
      `Audit log created:\n${auditLogPath}`
    );
    console.log(
      '=========================================='
    );
  } catch (error) {
    if (
      fs.existsSync(
        temporaryWorkbookPath
      )
    ) {
      fs.unlinkSync(
        temporaryWorkbookPath
      );
    }

    throw error;
  }
}

module.exports = { reconcileWorkbook };
