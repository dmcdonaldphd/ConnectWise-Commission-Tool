'use strict';

const fs = require('node:fs');

const CHANGE_ACTIONS = new Set(['UPDATE_EXISTING_ROW', 'FILL_BLANK_ROW', 'INSERT_NEW_ROW']);
const STATUS_LABELS = { COMPLETED: 'Completed successfully', COMPLETED_WITH_REVIEW: 'Completed — review items',
  COMPLETED_WITH_ERRORS: 'Completed with project errors', FAILED: 'Failed', PLANNED: 'Changes planned; write not confirmed' };
const describeRunStatus = status => STATUS_LABELS[status] || status;

function csvEscape(value) {
  const text = value === null || value === undefined ? '' : String(value);
  return /[,"\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function atomicWrite(filename, contents) {
  const temporary = `${filename}.${process.pid}.tmp`;
  try {
    fs.writeFileSync(temporary, contents, { encoding: 'utf8', flag: 'wx' });
    fs.renameSync(temporary, filename);
  } finally {
    if (fs.existsSync(temporary)) fs.unlinkSync(temporary);
  }
}

function saveRunAudit(result, entries) {
  const headers = ['Timestamp', 'Action', 'Project ID', 'Company', 'Invoice',
    'Workbook Row', 'Field', 'Old Value', 'New Value', 'Notes', 'Result', 'Original Workbook Row'];
  const lines = [headers.map(csvEscape).join(',')];
  for (const entry of entries) {
    lines.push([entry.timestamp, entry.action, entry.projectID, entry.company,
      entry.invoice, entry.rowNumber, entry.field, entry.oldValue, entry.newValue,
      entry.notes, entry.result, entry.originalRowNumber].map(csvEscape).join(','));
  }
  atomicWrite(result.auditLogPath, lines.join('\r\n') + '\r\n');
  const { summary } = result;
  const text = [
    `Workbook: ${result.workbookFilename}`,
    `Result: ${describeRunStatus(result.status)}`,
    `Workbook changed: ${result.writeUncertain ? 'Unconfirmed — compare with backup' : result.workbookWritten ? 'Yes' : 'No'}`,
    `Projects attempted: ${summary.projectsProcessed}`,
    `Projects completed: ${summary.projectsSucceeded}`,
    `Projects failed: ${summary.projectsFailed}`,
    `Invoices reviewed: ${summary.invoicesReviewed}`,
    `Applied rows updated / filled / inserted: ${result.workbookWritten && !result.writeUncertain ? `${summary.rowsUpdated} / ${summary.blankRowsFilled} / ${summary.rowsInserted}` : '0 / 0 / 0'}`,
    `Review items: ${result.reviewCount}`,
    `Backup: ${result.backupPath || 'Not needed / not created'}`,
    `Audit: ${result.auditLogPath}`,
    '',
    ...entries.filter(e => /ERROR|REVIEW/.test(e.action)).map(e =>
      `${e.action}${e.projectID ? ` — project ${e.projectID}` : ''}${e.invoice ? `, invoice ${e.invoice}` : ''}: ${e.notes || ''}`),
  ];
  atomicWrite(result.summaryLogPath, text.join('\r\n') + '\r\n');
}

module.exports = { CHANGE_ACTIONS, saveRunAudit, describeRunStatus };
