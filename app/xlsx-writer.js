'use strict';

// PHDTech Commission Tool - surgical XLSX writer (no Microsoft Excel automation).
// Edits the Commission worksheet, table metadata, styles, and calculation flags.
// Automatically extends the table at its bottom if existing unused rows run out.
// Does not launch Microsoft Excel or PowerShell, and preserves paid-history sheets.

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const REQUIRED = ['%', 'Company', 'Project ID', 'Project/Ticket', 'Notes',
  'Invoice Date', 'Invoice #', 'Inv Amt (Pretax)', 'Paid Amt (Pretax)', 'Paid Date', 'Comm', 'Paycheck Date'];
const MANAGED = new Set(['%', 'company', 'project id', 'project/ticket',
  'notes', 'invoice date', 'invoice #', 'inv amt (pretax)', 'paid amt (pretax)', 'paid date', 'paycheck date']);
const XML_ESC = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
  .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
const XML_UNESC = s => String(s).replace(/&(?:amp|lt|gt|quot|apos);/g, x => ({
  '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'",
})[x]);
const normalized = s => String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();
const attrs = s => Object.fromEntries(Array.from(s.matchAll(/([:\w.-]+)="([^"]*)"/g), m => [m[1], XML_UNESC(m[2])]));
const colnum = s => [...s.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const colname = n => { let s = ''; while (n > 0) { n--; s = String.fromCharCode(65 + n % 26) + s; n = Math.floor(n / 26); } return s; };
const timestamp = () => new Date().toISOString().replace('T', ' ').replace('Z', ' UTC');
let logPath = null;
function trace(s) {
  const line = `${timestamp()} ${s}`;
  if (logPath) fs.appendFileSync(logPath, line + '\n', 'utf8');
  console.log(s);
}
function fail(s) { throw new Error(s); }
function readZipText(zip, name) {
  const file = zip.file(name);
  if (!file) fail(`Missing XLSX part: ${name}`);
  return file.async('string');
}
function findCommissionSheet(workbookXml, relsXml) {
  const nodes = [...workbookXml.matchAll(/<sheet\s+[^>]*\/?\s*>/g)];
  const sheet = nodes.map(m => attrs(m[0])).find(a => a.name === 'Commission');
  if (!sheet) fail('Worksheet named Commission was not found.');
  const id = sheet['r:id'];
  const relationships = [...relsXml.matchAll(/<Relationship\s+[^>]*\/?\s*>/g)].map(m => attrs(m[0]));
  const rel = relationships.find(r => r.Id === id && /\/worksheet$/.test(r.Type));
  if (!rel) fail('Cannot resolve Commission worksheet path.');
  if (rel.Target.startsWith('/')) return rel.Target.replace(/^\//, '');
  return path.posix.normalize(path.posix.join('xl', rel.Target));
}
async function locateCommissionTable(zip) {
  const found = [];
  for (const name of Object.keys(zip.files).filter(n => /^xl\/tables\/table\d+\.xml$/i.test(n))) {
    const xml = await readZipText(zip, name);
    const root = xml.match(/<table\s+[^>]*>/);
    if (!root) continue;
    const a = attrs(root[0]);
    if (a.name === 'Table1' || a.displayName === 'Table1') found.push({ name, xml, attr: a });
  }
  if (found.length !== 1) fail(`Expected exactly one Table1 definition, got ${found.length}.`);
  const t = found[0];
  const ref = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(t.attr.ref || '');
  if (!ref) fail('Unsupported Commission table range: ' + t.attr.ref);
  const startCol = colnum(ref[1]);
  const startRow = Number(ref[2]);
  const endCol = colnum(ref[3]);
  const endRow = Number(ref[4]);
  if (startRow !== 1 || startCol !== 1) fail('Expected Commission table to start at A1; refusing a potentially incorrect write.');
  const headers = [...t.xml.matchAll(/<tableColumn\s+[^>]*?(?:\/>|>)/g)].map(m => attrs(m[0]).name);
  if (headers.length !== endCol - startCol + 1) fail('Table column metadata does not match its range.');
  const map = new Map();
  headers.forEach((h, i) => {
    const key = normalized(h);
    if (map.has(key)) fail(`Duplicate Commission table column header: ${h}`);
    map.set(key, startCol + i);
  });
  for (const h of REQUIRED) if (!map.has(normalized(h))) fail(`Missing Commission table column: ${h}`);
  if (!Number(t.attr.totalsRowCount)) fail('Expected a totals row in Commission table. The writer will not guess where data ends.');
  return { ...t, headers: map, startRow, startCol, endRow, dataStart: startRow + 1, dataEnd: endRow - Number(t.attr.totalsRowCount), endCol };
}

// Parsing only direct worksheet sheetData rows. Reuse their original XML structure
// to preserve worksheet settings, table definitions, formula and comment metadata.
function loadRows(xml, table) {
  const m = /<sheetData>([\s\S]*?)<\/sheetData>/.exec(xml);
  if (!m) fail('Invalid Commission worksheet: no sheetData.');
  const rows = new Map();
  for (const rowMatch of m[1].matchAll(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g)) {
    const raw = rowMatch[0];
    const opening = /^<row\b[^>]*>/.exec(raw);
    if (!opening) fail('Invalid worksheet row XML.');
    const ra = attrs(opening[0]);
    const rowNumber = Number(ra.r);
    if (!Number.isInteger(rowNumber) || rows.has(rowNumber)) fail('Duplicate/invalid worksheet row.');
    const cells = new Map();
    const inside = raw.slice(opening[0].length, -'</row>'.length);
    for (const c of inside.matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)) {
      const rawCell = c[0];
      const a = attrs(rawCell.slice(0, rawCell.indexOf('>') + 1));
      const cm = /^([A-Z]+)(\d+)$/.exec(a.r || '');
      if (!cm || Number(cm[2]) !== rowNumber) fail(`Unexpected cell reference in row ${rowNumber}.`);
      const col = colnum(cm[1]);
      if (cells.has(col)) fail(`Duplicate cell in row ${rowNumber}, column ${cm[1]}`);
      cells.set(col, rawCell);
    }
    rows.set(rowNumber, { rowNumber, opening: opening[0], cells, raw });
  }
  for (let r = table.dataStart; r <= table.dataEnd; r++) {
    if (!rows.has(r)) rows.set(r, { rowNumber: r, opening: `<row r="${r}">`, cells: new Map(), raw: null });
  }
  return { match: m, rows };
}
function cellAttrs(raw) { return raw ? attrs(raw.slice(0, raw.indexOf('>') + 1)) : {}; }
function cValue(raw) {
  if (!raw) return null;
  const a = cellAttrs(raw);
  if (a.t === 'inlineStr') {
    const m = raw.match(/<t(?:\s+[^>]*)?>([\s\S]*?)<\/t>/);
    return m ? XML_UNESC(m[1]) : '';
  }
  const m = raw.match(/<v>([\s\S]*?)<\/v>/);
  if (!m) return null;
  if (a.t === 's') return { sharedString: Number(m[1]) };
  if (a.t === 'e') return { error: m[1] };
  return Number.isFinite(Number(m[1])) ? Number(m[1]) : XML_UNESC(m[1]);
}
function isPresent(raw) { const v = cValue(raw); return hasFormula(raw) || (v !== null && v !== ''); }
function hasFormula(raw) { return !!raw && /<f(?:\s|>)/.test(raw); }
function ensureRow(table, rows, rowNumber) {
  if (!Number.isSafeInteger(rowNumber) || rowNumber < table.dataStart || rowNumber > table.dataEnd)
    fail(`Row ${rowNumber} is outside the Commission table data range ${table.dataStart}:${table.dataEnd}.`);
  return rows.get(rowNumber);
}
function excelSerial(value) {
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) fail('Invalid commission date: ' + value);
  return Math.floor((Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) - Date.UTC(1899, 11, 30)) / 86400000);
}
function replaceStyle(raw, style) {
  const m = /^<c\b([^>]*?)(?:\/?>)/.exec(raw);
  if (!m) fail('Cannot update XLSX cell style.');
  let open = m[0];
  if (/\bs="\d+"/.test(open)) open = open.replace(/\bs="\d+"/, `s="${style}"`);
  else open = open.replace(/\/?>(?=$)/, x => ` s="${style}"${x}`);
  return open + raw.slice(m[0].length);
}
function replaceCellRef(raw, newRow) {
  const opening = raw.match(/^<c\b[^>]*?(?:\/>|>)/);
  if (!opening) fail('Cannot relocate an XLSX cell.');
  const a = cellAttrs(raw);
  if (!a.r) fail('Cell lacks Excel reference.');
  const newRef = a.r.replace(/\d+$/, String(newRow));
  return raw.replace(/^<c\b[^>]*>/, opening[0].replace(/\br="[^"]+"/, `r="${newRef}"`));
}
function assignCell(row, col, value, isDate, highlight, styles, allowRemovingMovedFormula=false) {
  const old = row.cells.get(col);
  if (hasFormula(old) && !allowRemovingMovedFormula) fail(`Safety stop: attempted to replace formula at ${colname(col)}${row.rowNumber}.`);
  const a = cellAttrs(old);
  if (a.cm || a.vm) fail(`Safety stop: unsupported Excel cell metadata at ${colname(col)}${row.rowNumber}.`);
  let baseStyle = Number(a.s || 0);
  if (isDate && value !== null && value !== undefined && value !== '')
    baseStyle = styles.getDateStyle(baseStyle);
  const style = highlight ? styles.getStyle(baseStyle, highlight) : baseStyle;
  const reference = `${colname(col)}${row.rowNumber}`;
  const prefix = `<c r="${reference}"${style ? ` s="${style}"` : ''}`;
  let raw;
  if (value === null || value === undefined || value === '') raw = prefix + '/>';
  else if (isDate) raw = prefix + `><v>${excelSerial(value)}</v></c>`;
  else if (typeof value === 'number' && Number.isFinite(value)) raw = prefix + `><v>${value}</v></c>`;
  else if (typeof value === 'string') {
    const space = /(^\s|\s$)/.test(value) ? ' xml:space="preserve"' : '';
    raw = prefix + ` t="inlineStr"><is><t${space}>${XML_ESC(value)}</t></is></c>`;
  } else fail('Unsupported commission cell value: ' + JSON.stringify(value));
  row.cells.set(col, raw);
}
function rowValue(row, col, ss) {
  const v = cValue(row.cells.get(col));
  return v && typeof v === 'object' && 'sharedString' in v ? ss[v.sharedString] : v;
}
function readSharedStrings(xml) {
  if (!xml) return [];
  const list = [];
  for (const m of xml.matchAll(/<si(?:\s[^>]*)?>[\s\S]*?<\/si>/g)) {
    const ts = [...m[0].matchAll(/<t(?:\s+[^>]*)?>([\s\S]*?)<\/t>/g)];
    list.push(ts.map(x => XML_UNESC(x[1])).join(''));
  }
  return list;
}
function verifyProject(row, projectId, headers, ss) {
  const actual = rowValue(row, headers.get('project id'), ss);
  if (!Number.isFinite(Number(actual)) || Number(actual) !== Number(projectId))
    fail(`Safety stop: project mismatch at row ${row.rowNumber}: expected ${projectId}, found ${String(actual)}.`);
}
function makeStyles(xml) {
  const fillMatch = /<fills\s+count="(\d+)"[^>]*>([\s\S]*?)<\/fills>/.exec(xml);
  const xfMatch = /<cellXfs\s+count="(\d+)"[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
  if (!fillMatch || !xfMatch) fail('XLSX styles.xml does not have expected fill/style tables.');
  const xfs = [...xfMatch[2].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map(x => x[0]);
  if (xfs.length !== Number(xfMatch[1])) fail('XLSX cell style definitions do not match count.');
  const cache = new Map();
  // Match only the two known colors used by earlier versions of this tool.
  // Restore the closest existing non-highlighted style (font, borders, dates,
  // number formats and alignment stay unchanged). Historical sheets untouched.
  const originalFills = [...fillMatch[2].matchAll(/<fill\b[^>]*?(?:\/>|>[\s\S]*?<\/fill>)/g)].map(x => x[0]);
  if (originalFills.length !== Number(fillMatch[1])) fail('Invalid XLSX fill count.');
  const highlightFillIds = new Map();
  originalFills.forEach((xml, i) => {
    if (/<fgColor\s+rgb="(?:FFFFFF00|FFDAF2D0)"\s*\/>/i.test(xml) &&
        /<patternFill\s+patternType="solid"/.test(xml)) highlightFillIds.set(i, /FFFFFF00/i.test(xml) ? 'yellow' : 'green');
  });
  const xfKey = xml => xml.replace(/\s+fillId="[^"]*"/g, '').replace(/\s+applyFill="[^"]*"/g, '');
  const xfFill = xml => Number((/\bfillId="(\d+)"/.exec(xml) || [,'0'])[1]);
  function clearStyle(id) {
    if (!Number.isSafeInteger(id) || !xfs[id]) return id;
    if (!highlightFillIds.has(xfFill(xfs[id]))) return id;
    const key = xfKey(xfs[id]);
    const candidate = xfs.findIndex(xf => !highlightFillIds.has(xfFill(xf)) && xfKey(xf) === key);
    return candidate >= 0 ? candidate : id;
  }
  const fills = [];
  const newXfs = [];
  const dateStyles = new Map();
  const getXf = id => xfs[id] || newXfs[id - xfs.length];
  // Newly added workbook columns can still have Excel's General format. A
  // numeric date needs a date number format to display and round-trip as a date.
  // Preserve existing custom formats, fonts, borders, alignment and fills.
  function getDateStyle(oldIndex) {
    const xf = getXf(oldIndex);
    if (!xf) fail(`Unsupported date style ID: ${oldIndex}`);
    if (Number((/\bnumFmtId="(\d+)"/.exec(xf) || [, '0'])[1]) !== 0) return oldIndex;
    if (dateStyles.has(oldIndex)) return dateStyles.get(oldIndex);
    let dated = /\bnumFmtId="[^"]*"/.test(xf)
      ? xf.replace(/\bnumFmtId="[^"]*"/, 'numFmtId="14"')
      : xf.replace(/^<xf/, '<xf numFmtId="14"');
    dated = /\bapplyNumberFormat="[^"]*"/.test(dated)
      ? dated.replace(/\bapplyNumberFormat="[^"]*"/, 'applyNumberFormat="1"')
      : dated.replace(/^<xf/, '<xf applyNumberFormat="1"');
    const id = xfs.length + newXfs.length;
    newXfs.push(dated);
    dateStyles.set(oldIndex, id);
    return id;
  }
  const rgb = { yellow: 'FFFFFF00', green: 'FFDAF2D0' };
  const colors = new Map();
  // A style added earlier in this write operation is not in the original
  // cellXfs array. Track generated style IDs so that re-highlighting a cell
  // (for example while coloring an entire newly inserted row) is safe.
  const generatedStyles = new Map();
  function getStyle(oldIndex, color) {
    const previous = generatedStyles.get(oldIndex);
    if (previous) {
      // The requested highlight is already on this cell. Reuse it rather
      // than treating the generated style ID as an unknown workbook style.
      if (previous.color === color) return oldIndex;
      // If a subsequent operation wants a different highlight, apply it to
      // the original style, preserving fonts, borders and number formats.
      oldIndex = previous.baseStyle;
    }
    const key = `${oldIndex}/${color}`;
    if (cache.has(key)) return cache.get(key);
    if (!getXf(oldIndex) || !rgb[color]) fail(`Unsupported style ID or highlight: ${key}`);
    if (!colors.has(color)) {
      colors.set(color, Number(fillMatch[1]) + fills.length);
      fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="${rgb[color]}"/><bgColor indexed="64"/></patternFill></fill>`);
    }
    let xf = getXf(oldIndex);
    if (/\bfillId="[^"]*"/.test(xf)) xf = xf.replace(/\bfillId="[^"]*"/, `fillId="${colors.get(color)}"`);
    else xf = xf.replace(/^<xf/, `<xf fillId="${colors.get(color)}"`);
    if (/\bapplyFill="[^"]*"/.test(xf)) xf = xf.replace(/\bapplyFill="[^"]*"/, 'applyFill="1"');
    else xf = xf.replace(/^<xf/, '<xf applyFill="1"');
    const id = xfs.length + newXfs.length;
    newXfs.push(xf);
    generatedStyles.set(id, { baseStyle: oldIndex, color });
    cache.set(key, id);
    return id;
  }
  function serialize() {
    if (fills.length === 0 && newXfs.length === 0) return xml;
    let ret = xml.replace(fillMatch[0], fillMatch[0].replace(`count="${fillMatch[1]}"`, `count="${Number(fillMatch[1]) + fills.length}"`).replace('</fills>', fills.join('') + '</fills>'));
    ret = ret.replace(xfMatch[0], xfMatch[0].replace(`count="${xfMatch[1]}"`, `count="${Number(xfMatch[1]) + newXfs.length}"`).replace('</cellXfs>', newXfs.join('') + '</cellXfs>'));
    return ret;
  }
  return { getStyle, getDateStyle, clearStyle, serialize };
}
function rebuildRow(row) {
  if (row.rowNumber < 1) fail('Invalid row number.');
  const open = row.opening.replace(/\s*\/?>$/, '>');
  return `${open}${[...row.cells.entries()].sort(([a], [b]) => a-b).map(([, cell]) => cell).join('')}</row>`;
}
function setCalculatedCaches(rows, table, indices) {
  const col = table.headers.get('comm');
  for (const rowNum of indices) {
    if (rowNum < table.dataStart || rowNum > table.dataEnd) continue; // totals and trailing rows are not calculated Commission data
    const row = ensureRow(table, rows, rowNum);
    const raw = row.cells.get(col);
    if (raw && hasFormula(raw)) {
      const rewritten = raw.replace(/<v>[\s\S]*?<\/v>/, '<v>0</v>');
      row.cells.set(col, rewritten);
    }
  }
}
function translateCellFormulaForRow(raw, sourceRow, destinationRow) {
  if (!hasFormula(raw)) return raw;
  if (/<f\b[^>]*\bt="shared"/.test(raw)) fail('Cannot safely relocate shared formulas outside the calculated Comm column.');
  const offset=destinationRow-sourceRow;
  return raw.replace(/(<f(?:\s+[^>]*)?>)([\s\S]*?)(<\/f>)/g, (all, op, formula, cl) => {
    if (/!/.test(formula)) fail('Cannot safely move formula containing a cross-worksheet reference.');
    return op + formula.replace(/(^|[^A-Za-z0-9_.])(\$?[A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])/g,
      (whole, pre, col, rowAbs, rowNum) => pre+col+rowAbs+(rowAbs?rowNum:String(Number(rowNum)+offset))) + cl;
  }).replace(/<v>[\s\S]*?<\/v>/,'<v>0</v>');
}
function copyDataRow(from, to, table) {
  for (let col = 1; col <= table.endCol; col++) {
    if (col === table.headers.get('comm')) continue; // calculated formula remains in its original Excel row
    const raw = from.cells.get(col);
    if (!raw) { to.cells.delete(col); continue; }
    const a = cellAttrs(raw);
    if (a.cm || a.vm) fail(`Cannot safely relocate rich cell metadata in ${colname(col)}${from.rowNumber}.`);
    // Value + original style move as a unit to preserve row-specific formatting.
    to.cells.set(col, translateCellFormulaForRow(replaceCellRef(raw, to.rowNumber), from.rowNumber, to.rowNumber));
  }
}
function insertionRows(payload) {
  const all = [];
  const groups = (payload.insertGroups || []).slice().sort((a, b) => b.afterRowNumber - a.afterRowNumber);
  for (const g of groups) {
    let offset = 0;
    for (const insert of g.inserts || []) all.push({ insert, after: Number(g.afterRowNumber) + offset++ });
  }
  return all;
}
function protectUnsupportedReferences(sheetXml, countInserts) {
  if (!countInserts) return;
  // We move cell contents inside existing rows; objects anchored to specific
  // cells would not move with those contents. Refuse this uncommon case.
  if (/<(?:hyperlinks|mergeCells|legacyDrawing|drawing|oleObjects|controls)\b/.test(sheetXml))
    fail('Commission sheet contains cell-anchored objects or merged cells. Safe offline row insertion is not supported.');
}

// Reverting a tool highlight only changes the style index of existing
// Commission cells. This is in-memory XML work: no Excel per-cell COM calls.
function clearPreviousHighlights(rows, table, styles, affected) {
  let cells = 0;
  for (let r = table.dataStart; r <= table.dataEnd; r++) {
    const row = rows.get(r);
    if (!row) continue;
    for (const [column, raw] of row.cells) {
      const id = Number(cellAttrs(raw).s || 0);
      const base = styles.clearStyle(id);
      if (base !== id) {
        row.cells.set(column, replaceStyle(raw, base));
        affected.add(r);
        cells++;
      }
    }
  }
  trace(`Cleared ${cells} prior automation-highlighted Commission cell(s) using XLSX styles (no Excel).`);
}

// Real worksheet row insertion takes place only at the very end of the
// Commission data table. Existing project rows keep their addresses; totals
// and any worksheet rows below totals are moved, along with their formulas.
function shiftFormulaReferences(formula, startRow, amount, oldDataEnd, newDataEnd, isTotalsRow) {
  if (!formula) return formula;
  if (/<|>/.test(formula)) fail('Unsupported formula XML characters during table extension.');
  // Another-sheet reference could resolve against a different worksheet;
  // reject instead of silently moving that target's row references.
  if (/!/.test(formula) && /\$?[A-Z]{1,3}\$?\d+/i.test(formula))
    fail('Cannot safely move a Commission row containing cross-sheet references.');
  let f = formula.replace(/(^|[^A-Za-z0-9_.])((?:\$?[A-Z]{1,3}\$?)(\d+))/g,
    (whole, pre, ref, rowStr) => {
      const old = Number(rowStr);
      return pre + (old >= startRow ? ref.replace(/\d+$/, String(old + amount)) : ref);
    });
  // The totals-row COUNTA/SUBTOTAL frequently uses plain A2:A188 instead of
  // a table structured reference; extend such ranges to cover appended rows.
  if (isTotalsRow) {
    f = f.replace(/(\$?[A-Z]{1,3}\$?\d+\s*:\s*\$?[A-Z]{1,3}\$?)(\d+)/g,
      (whole, pre, last) => Number(last) === oldDataEnd ? pre + newDataEnd : whole);
  }
  return f;
}
function rewriteFormulaCells(cell, startRow, amount, oldDataEnd, newDataEnd, isTotalsRow) {
  if (!hasFormula(cell)) return cell;
  if (/<f\b[^>]*\bt="shared"/.test(cell)) fail('Cannot safely extend a table with shared formulas.');
  let changed = false;
  const result = cell.replace(/(<f(?:\s+[^>]*)?>)([\s\S]*?)(<\/f>)/g,
    (whole, op, f, cl) => {
      const next = shiftFormulaReferences(f, startRow, amount, oldDataEnd, newDataEnd, isTotalsRow);
      if (next !== f) changed = true;
      return op + next + cl;
    });
  return changed ? result.replace(/<v>[\s\S]*?<\/v>/, '<v>0</v>') : cell;
}
function extendTableRange(table, growth) {
  const oldEnd = table.endRow;
  const oldDataEnd = table.dataEnd;
  // All relevant ranges begin at row 1 or end at the old last data row.
  function incrementRef(attr, xml) {
    return xml.replace(new RegExp(`(<${attr}\\b[^>]*?\\bref=")([A-Z]+)(\\d+):([A-Z]+)(\\d+)(")`, 'g'),
      (whole, prefix, leftCol, leftRow, rightCol, rightRow, suffix) => {
        const end = Number(rightRow);
        if (end !== oldEnd && end !== oldDataEnd) return whole;
        return `${prefix}${leftCol}${leftRow}:${rightCol}${end + growth}${suffix}`;
      });
  }
  let xml = table.xml;
  for (const tag of ['table','autoFilter','sortState','sortCondition']) xml = incrementRef(tag, xml);
  table.xml = xml;
  table.endRow += growth;
  table.dataEnd += growth;
}
// Excel sometimes stores the Commission calculated column as a shared formula
// (one master formula + empty followers). Growing the table invalidates the
// shared formula range. Convert its Commission data cells to equivalent
// per-row formulas before inserting rows, preserving the same calculation.
function materializeCommissionSharedFormula(rows, table, affected) {
  const col = table.headers.get('comm');
  const groups = new Map();
  for (let r=table.dataStart; r<=table.dataEnd; r++) {
    const raw = rows.get(r).cells.get(col);
    const master = raw && /<f\b([^>]*\bt="shared"[^>]*)>([\s\S]*?)<\/f>/.exec(raw);
    if (!master || !master[2]) continue;
    const a = attrs('<f '+master[1]+'>');
    if (!a.si || !/^[0-9]+$/.test(a.si)) fail('Malformed Excel shared formula ID.');
    if (groups.has(a.si)) fail('Duplicate Excel shared formula master ID in Commission data.');
    groups.set(a.si, { row:r, formula:master[2] });
  }
  let converted=0;
  for (let r=table.dataStart; r<=table.dataEnd; r++) {
    const row=rows.get(r);
    const raw=row.cells.get(col);
    if (!raw || !/<f\b[^>]*\bt="shared"/.test(raw)) continue;
    const tag=/<f\b[^>]*\bt="shared"[^>]*(?:\/>|>[\s\S]*?<\/f>)/.exec(raw);
    if (!tag) fail('Malformed shared formula in Commission data.');
    const a=attrs(tag[0]);
    const master=groups.get(a.si);
    if (!master) fail(`Shared Commission formula has no master for row ${r}.`);
    const offset=r-master.row;
    if (/!/.test(master.formula)) fail('Cannot safely materialize shared Commission formula referencing another worksheet.');
    const formula=master.formula.replace(/(^|[^A-Za-z0-9_.])(\$?[A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_])/g,
      (whole, pre, colref, rowAbs, rowref) => pre+colref+rowAbs+(rowAbs?rowref:String(Number(rowref)+offset)));
    let updated=raw.replace(tag[0], `<f>${formula}</f>`).replace(/<v>[\s\S]*?<\/v>/, '<v>0</v>');
    row.cells.set(col,updated);
    affected.add(r);
    converted++;
  }
  if (converted) trace(`Converted ${converted} shared Commission formula cells into equivalent row-specific formulas for safe table growth.`);
}

function growCommissionTable(rows, table, sheetXml, amount, affected) {
  if (amount <= 0) return { sheetXml, growth: 0 };
  if (!Number.isInteger(amount) || amount > 10000) fail('Unsafe Commission table growth requested.');
  materializeCommissionSharedFormula(rows, table, affected);
  const oldTotalsRow = table.dataEnd + 1;
  const oldDataEnd = table.dataEnd;
  const newDataEnd = oldDataEnd + amount;
  const template = rows.get(oldDataEnd);
  if (!template) fail('Cannot extend Commission table without a formula/style template row.');
  const templateStyles = new Map([...template.cells].map(([col, raw]) => [col, cellAttrs(raw).s]));
  const templateFormula = template.cells.get(table.headers.get('comm'));
  if (!hasFormula(templateFormula)) fail('Cannot extend Commission table without an existing Comm calculated formula.');
  // Whole-column references and off-sheet anchors require Excel-specific
  // shifting semantics; these are deliberately refused, not guessed.
  if (/<(?:mergeCells|hyperlinks|drawing|legacyDrawing|oleObjects|controls|conditionalFormatting|dataValidations)\b/.test(sheetXml))
    fail('Commission worksheet has anchored objects/ranges; safe table extension is unsupported.');
  const shifted = [...rows.keys()].filter(r => r >= oldTotalsRow).sort((a,b) => b-a);
  for (const oldRow of shifted) {
    const row = rows.get(oldRow);
    rows.delete(oldRow);
    row.rowNumber = oldRow + amount;
    row.opening = row.opening.replace(/\br="\d+"/, `r="${row.rowNumber}"`);
    const changed = new Map();
    for (const [col, cell] of row.cells) {
      let raw = replaceCellRef(cell, row.rowNumber);
      raw = rewriteFormulaCells(raw, oldTotalsRow, amount, oldDataEnd, newDataEnd, oldRow === oldTotalsRow);
      changed.set(col, raw);
    }
    row.cells = changed;
    row.raw = null;
    rows.set(row.rowNumber, row);
    affected.add(row.rowNumber);
  }
  // The newly added rows are genuine worksheet/table rows, not previously
  // allocated blanks. Clone only presentation and the calculated Comm formula.
  for (let nr = oldTotalsRow; nr < oldTotalsRow + amount; nr++) {
    const cells = new Map();
    for (let col = 1; col <= table.endCol; col++) {
      const s = templateStyles.get(col);
      const style = s ? ` s="${s}"` : '';
      const ref = `${colname(col)}${nr}`;
      if (col === table.headers.get('comm')) {
        let cell = replaceCellRef(templateFormula, nr);
        // Only relative coordinate references in the template shift to this row.
        const offset = nr - oldDataEnd;
        if (offset && hasFormula(cell)) cell = rewriteFormulaCells(cell, oldDataEnd, offset, oldDataEnd, nr, false);
        cells.set(col, cell);
      } else {
        cells.set(col, `<c r="${ref}"${style}/>`);
      }
    }
    rows.set(nr, { rowNumber:nr, opening:`<row r="${nr}">`, cells, raw:null });
    affected.add(nr);
  }
  // If other existing formulas on the Commission worksheet reference the
  // old totals row or below, move those references too (same-sheet only).
  for (const [r, row] of rows) {
    if (r >= oldTotalsRow) continue;
    for (const [col, cell] of row.cells) {
      if (!hasFormula(cell)) continue;
      const rewritten = rewriteFormulaCells(cell, oldTotalsRow, amount, oldDataEnd, newDataEnd, false);
      if (rewritten !== cell) {
        row.cells.set(col, rewritten);
        affected.add(r);
      }
    }
  }
  extendTableRange(table, amount);
  sheetXml = sheetXml.replace(/(<dimension\s+ref="[A-Z]+\d+:[A-Z]+)(\d+)(")/,
    (whole, pre, end, suffix) => `${pre}${Number(end) >= oldTotalsRow ? Number(end) + amount : Number(end)}${suffix}`);
  trace(`Physically added ${amount} Commission table row(s); table now ends at ${table.endRow}, totals shifted from ${oldTotalsRow} to ${oldTotalsRow+amount}.`);
  return { sheetXml, growth: amount };
}
async function protectExternalCommissionReferences(zip, commissionSheetPath, firstMovedRow) {
  const ref = /(?:'Commission'|Commission)!\$?[A-Z]{1,3}\$?(\d+)/gi;
  for (const name of Object.keys(zip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/i.test(n) && n !== commissionSheetPath)) {
    const xml=await readZipText(zip,name);
    for (const f of xml.matchAll(/<f(?:\s[^>]*)?>([\s\S]*?)<\/f>/g)) {
      for (const match of f[1].matchAll(ref)) {
        if (Number(match[1]) >= firstMovedRow) {
          fail(`Safety stop: formula in ${name} references Commission row ${match[1]} that would move during table expansion.`);
        }
      }
    }
  }
  const workbookXml=await readZipText(zip,'xl/workbook.xml');
  for (const m of workbookXml.matchAll(/<definedName\b[^>]*>([\s\S]*?)<\/definedName>/g)) {
    for (const hit of m[1].matchAll(ref)) {
      if (Number(hit[1]) >= firstMovedRow) fail('Safety stop: a defined name refers to a Commission row that would move.');
    }
  }
}

function discardStaleCalculationChain(zip) {
  // Excel rebuilds calcChain the next time it recalculates. Keeping stale
  // absolute formula coordinates after physically inserting rows is unsafe.
  const chain = 'xl/calcChain.xml';
  if (!zip.file(chain)) return;
  zip.remove(chain);
  return true;
}

async function writeWorkbook(payload) {
  const target = path.resolve(payload.workbookPath || '');
  const buffer = await fs.promises.readFile(target);
  const zip = await JSZip.loadAsync(buffer, { checkCRC32: true });
  const wxml = await readZipText(zip, 'xl/workbook.xml');
  const rels = await readZipText(zip, 'xl/_rels/workbook.xml.rels');
  const sheetName = findCommissionSheet(wxml, rels);
  const table = await locateCommissionTable(zip);
  const originalTableEndRow = table.endRow;
  let sheetXml = await readZipText(zip, sheetName);
  const styleXml = await readZipText(zip, 'xl/styles.xml');
  const ss = zip.file('xl/sharedStrings.xml') ? readSharedStrings(await readZipText(zip, 'xl/sharedStrings.xml')) : [];
  const state = loadRows(sheetXml, table);
  const rows = state.rows;
  const styles = makeStyles(styleXml);
  const h = table.headers;
  const affected = new Set();
  const insertions = insertionRows(payload);
  // Defend the writer boundary too: no fill/insert may reuse an invoice number
  // already in Commission or reserved by an earlier operation in this payload.
  const invoiceKey = value => String(value === null || value === undefined ? '' : value).trim().toUpperCase();
  const occupiedInvoices = new Set();
  for (let r = table.dataStart; r <= table.dataEnd; r++) {
    const key = invoiceKey(rowValue(rows.get(r), h.get('invoice #'), ss));
    if (key) occupiedInvoices.add(key);
  }
  for (const operation of [...(payload.blankRowFills || []), ...insertions.map(item => item.insert)]) {
    const key = invoiceKey(operation.invoice);
    if (!key || occupiedInvoices.has(key)) fail(`Safety stop: blank or duplicate invoice in write plan: ${operation.invoice}.`);
    occupiedInvoices.add(key);
  }
  protectUnsupportedReferences(sheetXml, insertions.length);
  clearPreviousHighlights(rows, table, styles, affected);
  trace(`Offline writer: Commission=${sheetName}, data rows ${table.dataStart}..${table.dataEnd}, update=${(payload.existingRowUpdates||[]).length}, fill=${(payload.blankRowFills||[]).length}, insert=${insertions.length}`);

  for (const u of payload.existingRowUpdates || []) {
    const row = ensureRow(table, rows, Number(u.rowNumber));
    verifyProject(row, u.projectID, h, ss);
    const invoice = rowValue(row, h.get('invoice #'), ss);
    if (normalized(invoice) !== normalized(u.invoice)) fail(`Invoice mismatch at row ${u.rowNumber}: expected ${u.invoice}, found ${invoice}.`);
    for (const change of u.changes || []) {
      const field = normalized(change.field);
      if (!MANAGED.has(field)) fail(`Unexpected field update: ${change.field}`);
      if (!['invoice date', 'inv amt (pretax)', 'paid date', 'paid amt (pretax)'].includes(field)) fail(`Unexpected reconciliation update field: ${change.field}`);
      assignCell(row, h.get(field), change.newValue, field === 'paid date' || field === 'invoice date', 'yellow', styles);
      affected.add(row.rowNumber);
    }
  }
  for (const fill of payload.blankRowFills || []) {
    const row = ensureRow(table, rows, Number(fill.rowNumber));
    verifyProject(row, fill.projectID, h, ss);
    if (isPresent(row.cells.get(h.get('invoice #')))) fail(`Refusing to overwrite existing invoice at row ${fill.rowNumber}.`);
    assignCell(row, h.get('invoice #'), String(fill.invoice), false, 'green', styles);
    assignCell(row, h.get('inv amt (pretax)'), fill.invoiceAmount, false, 'green', styles);
    assignCell(row, h.get('invoice date'), fill.invoiceDate, true, 'green', styles);
    assignCell(row, h.get('paid amt (pretax)'), fill.paidAmount, false, 'green', styles);
    assignCell(row, h.get('paid date'), fill.paidDate, true, 'green', styles);
    affected.add(row.rowNumber);
  }
  if (insertions.length) {
    // Do not physically insert worksheet rows: use existing blank rows within
    // the table, shifting table *values* down (not formulas or row addresses).
    const checkCols = [...new Set(Array.from({length:table.endCol}, (_, i) => i+1).filter(i => i !== h.get('%') && i !== h.get('comm')))];
    let lastUsed = table.dataStart-1;
    for (let r = table.dataStart; r <= table.dataEnd; r++) {
      if (checkCols.some(c => isPresent(rows.get(r).cells.get(c)))) lastUsed = r;
    }
    const shortage = Math.max(0, lastUsed + insertions.length - table.dataEnd);
    if (shortage) {
      const expanded = growCommissionTable(rows, table, sheetXml, shortage, affected);
      sheetXml = expanded.sheetXml;
    }
    trace(`New invoice allocation: ${table.dataEnd - lastUsed} row slot(s) available after optional automatic expansion.`);
    for (const { insert, after } of insertions) {
      const desired = after + 1;
      if (desired <= table.dataStart || desired > lastUsed+1) fail(`Unsupported insertion position after row ${after}.`);
      const predecessor = ensureRow(table, rows, after);
      verifyProject(predecessor, insert.projectID, h, ss);
      const pctRaw = predecessor.cells.get(h.get('%'));
      if (!pctRaw || hasFormula(pctRaw)) fail(`Cannot safely copy commission percentage from row ${after}.`);
      for (let r=lastUsed; r>=desired; r--) {
        copyDataRow(rows.get(r), rows.get(r+1), table);
        affected.add(r+1);
      }
      lastUsed++;
      const row = ensureRow(table, rows, desired);
      // Retain calculated Commission cell/formula; clear every other cell's
      // value to avoid carrying over previous data into this new invoice row.
      for (const col of checkCols) assignCell(row, col, null, false, null, styles, true);
      // Copy the original percentage and populate only meaningful fields.
      const pct = cValue(pctRaw);
      if (pct && typeof pct === 'object') fail('Unsupported rich percentage cell.');
      assignCell(row, h.get('%'), pct, false, null, styles);
      assignCell(row, h.get('company'), insert.company, false, null, styles);
      assignCell(row, h.get('project id'), Number(insert.projectID), false, null, styles);
      assignCell(row, h.get('project/ticket'), insert.description, false, null, styles);
      assignCell(row, h.get('invoice #'), insert.invoice, false, null, styles);
      assignCell(row, h.get('inv amt (pretax)'), insert.invoiceAmount, false, null, styles);
      assignCell(row, h.get('invoice date'), insert.invoiceDate, true, null, styles);
      assignCell(row, h.get('paid amt (pretax)'), insert.paidAmount, false, null, styles);
      assignCell(row, h.get('paid date'), insert.paidDate, true, null, styles);
      // Match the old Excel writer: green highlight across the entire new row.
      // This changes only cell style, leaving Commission formulas untouched.
      for (let col=1; col<=table.endCol; col++) {
        const oldXml=row.cells.get(col);
        if (!oldXml) continue;
        const oldStyle=Number(cellAttrs(oldXml).s || 0);
        row.cells.set(col, replaceStyle(oldXml, styles.getStyle(oldStyle, 'green')));
      }
      affected.add(desired);
      trace(`Allocated Commission row ${desired} for invoice ${insert.invoice} (project ${insert.projectID}).`);
    }
  }
  setCalculatedCaches(rows, table, affected);
  // Preserve untouched rows byte-for-byte while serializing new/moved rows.
  const missingModified = [...affected].filter(r => !rows.has(r));
  if (missingModified.length) fail('Modified worksheet rows were missing from in-memory XLSX data.');
  // A non-expanded old data row missing in source XML must not be manufactured
  // inadvertently; the user's workbooks have materialized table rows.
  let changedRows = 0;
  const sheetDataBody = [...rows.entries()].sort(([a], [b]) => a-b).map(([r, row]) => {
    if (!affected.has(r) && row.raw) return row.raw;
    if (!affected.has(r) && !row.raw) return '';
    changedRows++;
    return rebuildRow(row);
  }).join('');
  if (changedRows !== affected.size) fail('Not all Commission modifications were serialized.');
  sheetXml = sheetXml.replace(state.match[0], `<sheetData>${sheetDataBody}</sheetData>`);
  zip.file(table.name, table.xml);
  // When table structure changes, stale formula addresses in calcChain must
  // be removed along with its metadata; Excel will rebuild it on next open.
  if (table.endRow !== originalTableEndRow) {
    discardStaleCalculationChain(zip);
    const relsXml = await readZipText(zip, 'xl/_rels/workbook.xml.rels');
    zip.file('xl/_rels/workbook.xml.rels', relsXml.replace(/<Relationship\b(?=[^>]*Type="[^"]*\/calcChain")[^>]*\/>/g, ''));
    const typesXml = await readZipText(zip, '[Content_Types].xml');
    zip.file('[Content_Types].xml', typesXml.replace(/<Override\b(?=[^>]*PartName="\/xl\/calcChain.xml")[^>]*\/>/g, ''));
  }
  zip.file(sheetName, sheetXml);
  const newStyles = styles.serialize();
  if (newStyles !== styleXml) zip.file('xl/styles.xml', newStyles);
  // Date/numeric changes invalidate cached Comm formulas, so request Excel
  // refresh cached calculated results when the user next opens the workbook.
  let newWorkbookXml = wxml;
  if (/<calcPr\b[^>]*\/>/.test(wxml)) {
    newWorkbookXml = wxml.replace(/<calcPr\b[^>]*\/>/, m => {
      let n = m;
      for (const [key, val] of [['calcMode','auto'],['fullCalcOnLoad','1'],['forceFullCalc','1']]) {
        if (new RegExp('\\b'+key+'="[^"]*"').test(n)) n=n.replace(new RegExp('\\b'+key+'="[^"]*"'), `${key}="${val}"`);
        else n=n.replace(/\/>$/, ` ${key}="${val}"/>`);
      }
      return n;
    });
  } else newWorkbookXml = /<extLst\b/.test(wxml)
    ? wxml.replace(/<extLst\b/, '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/><extLst')
    : wxml.replace('</workbook>', '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>');
  if (newWorkbookXml !== wxml) zip.file('xl/workbook.xml', newWorkbookXml);
  const out = await zip.generateAsync({ type:'nodebuffer', compression:'DEFLATE', compressionOptions:{level:6} });
  const verify = await JSZip.loadAsync(out, { checkCRC32: true });
  if (!verify.file(sheetName)) fail('Generated workbook failed integrity check.');
  // Atomic replacement: original stays intact until generated workbook is ready.
  const temp = `${target}.commission-write-${process.pid}.tmp`;
  try {
    await fs.promises.writeFile(temp, out, { flag:'wx' });
    await fs.promises.rename(temp, target);
  } finally {
    if (fs.existsSync(temp)) await fs.promises.unlink(temp).catch(()=>{});
  }
  // The replacement has committed. A trace-file failure must not report that
  // the workbook write itself failed; the caller records this warning in audit.
  try {
    trace(`Workbook saved successfully without Microsoft Excel. ${affected.size} affected row(s); historical sheets unchanged.`);
  } catch (error) {
    console.warn(`Workbook was saved, but its writer trace could not be finalized: ${error.message}`);
  }
}
async function main() {
  if (process.argv.length < 3) fail('Usage: node xlsx-writer.js PAYLOAD_PATH [TRACE_PATH]');
  if (process.argv[3]) {
    logPath=path.resolve(process.argv[3]);
    fs.mkdirSync(path.dirname(logPath), {recursive:true});
  }
  const payload = JSON.parse(await fs.promises.readFile(process.argv[2], 'utf8'));
  await writeWorkbook(payload);
}
if (require.main === module) main().catch(e=>{try{trace('ERROR: '+(e.stack||e.message||String(e)));}catch{}process.exitCode=1;});
module.exports={writeWorkbook};
