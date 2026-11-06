const fs = require('fs');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

async function loadWorkbook(filename) {
  const workbook = new ExcelJS.Workbook();

  try {
    await workbook.xlsx.readFile(filename, {
      ignoreNodes: ['autoFilter'],
    });
    return workbook;
  } catch (error) {
    const message = String(error && error.message ? error.message : error);

    // Some valid Excel workbooks contain dynamic date filters such as
    // <dynamicFilter type="thisMonth" .../>. Certain ExcelJS versions fail
    // while parsing these nodes even when autoFilter is listed in ignoreNodes.
    // In that specific case, sanitize an in-memory copy of the XLSX package and
    // load that copy. The original workbook on disk is never modified.
    if (!message.includes('Unexpected xml node in parseOpen') || !message.includes('dynamicFilter')) {
      throw error;
    }

    const sourceBuffer = await fs.promises.readFile(filename);
    const zip = await JSZip.loadAsync(sourceBuffer);
    let changed = false;

    const worksheetFiles = Object.keys(zip.files).filter((name) =>
      /^xl\/worksheets\/sheet\d+\.xml$/i.test(name)
    );

    for (const worksheetName of worksheetFiles) {
      const worksheetFile = zip.file(worksheetName);
      if (!worksheetFile) {
        continue;
      }

      const xml = await worksheetFile.async('string');
      const sanitizedXml = xml
        .replace(/<dynamicFilter\b[^>]*\/>/gi, '')
        .replace(/<dynamicFilter\b[^>]*>[\s\S]*?<\/dynamicFilter>/gi, '');

      if (sanitizedXml !== xml) {
        zip.file(worksheetName, sanitizedXml);
        changed = true;
      }
    }

    if (!changed) {
      throw error;
    }

    const sanitizedBuffer = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
    });

    const sanitizedWorkbook = new ExcelJS.Workbook();
    await sanitizedWorkbook.xlsx.load(sanitizedBuffer, {
      ignoreNodes: ['autoFilter'],
    });

    return sanitizedWorkbook;
  }
}

function getWorksheet(workbook, name) {
  const sheet = workbook.getWorksheet(name);

  if (!sheet) {
    throw new Error(`Worksheet "${name}" not found.`);
  }

  return sheet;
}

function getCellValue(cell) {
  const value = cell.value;

  if (
    value &&
    typeof value === 'object' &&
    Object.prototype.hasOwnProperty.call(value, 'result')
  ) {
    return value.result;
  }

  return value;
}

function cleanText(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
}

function normalizeHeader(value) {
  return cleanText(value)
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function getHeaderMap(sheet, requiredHeaders = []) {
  const headerRow = sheet.getRow(1);
  const headers = new Map();
  const displayNames = new Map();

  const lastColumn = Math.max(sheet.columnCount, headerRow.cellCount);

  for (let columnNumber = 1; columnNumber <= lastColumn; columnNumber += 1) {
    const cell = headerRow.getCell(columnNumber);
    const headerText = cleanText(cell.text || getCellValue(cell));
    const normalized = normalizeHeader(headerText);

    if (!normalized) {
      continue;
    }

    if (headers.has(normalized)) {
      throw new Error(
        `Worksheet "${sheet.name}" has duplicate header "${headerText}" in columns ` +
        `${headers.get(normalized)} and ${columnNumber}.`
      );
    }

    headers.set(normalized, columnNumber);
    displayNames.set(normalized, headerText);
  }

  const missing = [];

  for (const requiredHeader of requiredHeaders) {
    const normalized = normalizeHeader(requiredHeader);
    if (!headers.has(normalized)) {
      missing.push(requiredHeader);
    }
  }

  if (missing.length > 0) {
    throw new Error(
      `Worksheet "${sheet.name}" is missing required column header(s): ${missing.join(', ')}.`
    );
  }

  return {
    get(headerName) {
      const normalized = normalizeHeader(headerName);
      const columnNumber = headers.get(normalized);

      if (!columnNumber) {
        throw new Error(
          `Worksheet "${sheet.name}" does not contain column header "${headerName}".`
        );
      }

      return columnNumber;
    },
    has(headerName) {
      return headers.has(normalizeHeader(headerName));
    },
    entries: headers,
    displayNames,
  };
}

const COMMISSION_REQUIRED_HEADERS = [
  '%',
  'Company',
  'Project ID',
  'Project/Ticket',
  'Notes',
  'Invoice #',
  'Paid Amt (Pretax)',
  'Paid Date',
  'Comm',
  'Paycheck Date',
];

const PAID_HISTORY_REQUIRED_HEADERS = [
  'Invoice #',
  'Paycheck Date',
];

function readProjects(sheet) {
  const rows = [];
  const columns = getHeaderMap(sheet, COMMISSION_REQUIRED_HEADERS);

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      return;
    }

    const rawProjectID = getCellValue(row.getCell(columns.get('Project ID')));

    if (
      rawProjectID === null ||
      rawProjectID === undefined ||
      String(rawProjectID).trim() === ''
    ) {
      return;
    }

    const projectID = Number(rawProjectID);

    if (!Number.isInteger(projectID) || projectID <= 0) {
      console.warn(
        `Skipping row ${rowNumber}: invalid Project ID "${rawProjectID}"`
      );
      return;
    }

    rows.push({
      rowNumber,
      projectID,
      company: cleanText(row.getCell(columns.get('Company')).text),
      description: cleanText(row.getCell(columns.get('Project/Ticket')).text),
      notes: cleanText(row.getCell(columns.get('Notes')).text),
      invoice: cleanText(row.getCell(columns.get('Invoice #')).text),
      paidAmount: getCellValue(row.getCell(columns.get('Paid Amt (Pretax)'))),
      paidDate: getCellValue(row.getCell(columns.get('Paid Date'))),
    });
  });

  return rows;
}

function groupRowsByProject(rows) {
  const projects = new Map();

  for (const row of rows) {
    if (!projects.has(row.projectID)) {
      projects.set(row.projectID, {
        projectID: row.projectID,
        company: row.company,
        description: row.description,
        workbookRows: [],
      });
    }

    projects.get(row.projectID).workbookRows.push(row);
  }

  return Array.from(projects.values());
}

module.exports = {
  loadWorkbook,
  getWorksheet,
  getCellValue,
  getHeaderMap,
  COMMISSION_REQUIRED_HEADERS,
  PAID_HISTORY_REQUIRED_HEADERS,
  readProjects,
  groupRowsByProject,
};
