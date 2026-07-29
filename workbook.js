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
        // Standard self-closing dynamic filter nodes.
        .replace(/<dynamicFilter\b[^>]*\/>/gi, '')
        // Defensive handling in case Excel emits an explicit closing tag.
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

function readProjects(sheet) {
  const rows = [];

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      return;
    }

    const rawProjectID = getCellValue(row.getCell(3));

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
      company: cleanText(row.getCell(2).text),
      description: cleanText(row.getCell(4).text),
      notes: cleanText(row.getCell(5).text),
      invoice: cleanText(row.getCell(6).text),
      paidAmount: getCellValue(row.getCell(7)),
      paidDate: getCellValue(row.getCell(8)),
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
  readProjects,
  groupRowsByProject,
};
