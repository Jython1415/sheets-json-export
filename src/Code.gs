/**
 * Runs when the add-on is installed
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Adds a custom menu to Google Sheets when the spreadsheet opens
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem('Export Rows (Markdown)', 'exportRowBasedMarkdown')
    .addItem('Export Columns (Markdown)', 'exportColumnBasedMarkdown')
    .addItem('Export Formulas (XML)', 'exportFormulasXML')
    .addItem('Export All Data (XML)', 'exportGeneralXML')
    .addToUi();
}

/**
 * Exports the selected range as row-based Markdown-KV format
 */
function exportRowBasedMarkdown() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select a range first.');
    return;
  }

  const data = buildRowBasedMarkdown(range);
  showClipboardDialog(data, 'Row-based Data');
}

/**
 * Exports the selected range as column-based Markdown-KV format
 */
function exportColumnBasedMarkdown() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select a range first.');
    return;
  }

  const data = buildColumnBasedMarkdown(range);
  showClipboardDialog(data, 'Column-based Data');
}

/**
 * Exports the selected range formulas as XML
 */
function exportFormulasXML() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select a range first.');
    return;
  }

  const data = buildFormulasXML(range);
  showClipboardDialog(data, 'Formulas (XML)');
}

/**
 * Exports the selected range values as general XML
 */
function exportGeneralXML() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select a range first.');
    return;
  }

  const data = buildGeneralXML(range);
  showClipboardDialog(data, 'All Data (XML)');
}

/**
 * Helper: Extract range metadata
 */
function getRangeMeta(range) {
  return {
    sheetName: range.getSheet().getName(),
    rangeNotation: range.getA1Notation(),
    startRow: range.getRow(),
    startCol: range.getColumn(),
    numRows: range.getNumRows(),
    numCols: range.getNumColumns(),
    totalCells: range.getNumRows() * range.getNumColumns(),
    timestamp: new Date().toISOString()
  };
}

/**
 * Helper: Convert column number to letter (1=A, 27=AA)
 */
function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const remainder = (col - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

/**
 * Helper: Get absolute A1 notation for a cell
 */
function getCellA1Notation(startRow, startCol, row, col) {
  const absoluteRow = startRow + row - 1;
  const absoluteCol = startCol + col - 1;
  return columnToLetter(absoluteCol) + absoluteRow;
}

/**
 * Helper: Get relative INDEX notation
 */
function getRelativeNotation(row, col) {
  return 'INDEX(' + row + ',' + col + ')';
}

/**
 * Helper: Get row header (first column value)
 */
function getRowHeader(values, row) {
  return values[row - 1] ? values[row - 1][0] : '';
}

/**
 * Helper: Get column header (first row value)
 */
function getColHeader(values, col) {
  return values[0] ? values[0][col - 1] : '';
}

/**
 * Helper: Calculate absolute row number
 */
function getAbsoluteRowNumber(startRow, relativeRow) {
  return startRow + relativeRow - 1;
}

/**
 * Helper: Calculate absolute column letter
 */
function getAbsoluteColumnLetter(startCol, relativeCol) {
  return columnToLetter(startCol + relativeCol - 1);
}

/**
 * Helper: Escape XML special characters
 */
function escapeXML(text) {
  if (typeof text !== 'string') {
    text = String(text);
  }
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Helper: Escape Markdown special characters
 */
function escapeMarkdown(text) {
  if (typeof text !== 'string') {
    text = String(text);
  }
  return text
    .replace(/\\/g, '\\\\')
    .replace(/\*/g, '\\*')
    .replace(/_/g, '\\_')
    .replace(/\[/g, '\\[')
    .replace(/\]/g, '\\]');
}

/**
 * Helper: Check if cell is empty
 */
function isEmptyCell(value) {
  return value === null || value === undefined || value === '';
}

/**
 * Build row-based Markdown-KV format
 */
function buildRowBasedMarkdown(range) {
  const meta = getRangeMeta(range);
  const values = range.getValues();

  let output = '---\n';
  output += 'Sheet: ' + meta.sheetName + '\n';
  output += 'Range: ' + meta.rangeNotation + '\n';
  output += 'Selection: ' + meta.numRows + ' rows × ' + meta.numCols + ' columns (' + meta.totalCells + ' cells)\n';

  const headerValues = [];
  for (let col = 0; col < meta.numCols; col++) {
    headerValues.push(values[0][col]);
  }
  output += 'Headers: ' + headerValues.join(', ') + '\n';
  output += 'Export Type: Row-based (Markdown-KV)\n';
  output += 'Timestamp: ' + meta.timestamp + '\n';
  output += 'Note: First row treated as headers. Each subsequent row is a record.\n';
  output += '---\n\n';

  for (let row = 2; row <= meta.numRows; row++) {
    const absRowNum = getAbsoluteRowNumber(meta.startRow, row);
    const rowRange = 'A' + absRowNum + ':' + columnToLetter(meta.startCol + meta.numCols - 1) + absRowNum;
    const startRowRange = getCellA1Notation(meta.startRow, meta.startCol, row, 1);
    const endRowRange = getCellA1Notation(meta.startRow, meta.startCol, row, meta.numCols);
    const actualRange = startRowRange.match(/[A-Z]+/)[0] + absRowNum + ':' + endRowRange.match(/[A-Z]+/)[0] + absRowNum;

    output += '## Row ' + row + ' (Sheet Row ' + absRowNum + ') [' + actualRange + ']\n\n';

    for (let col = 1; col <= meta.numCols; col++) {
      const header = values[0][col - 1];
      const cellValue = values[row - 1][col - 1];
      const absRef = getCellA1Notation(meta.startRow, meta.startCol, row, col);
      const relRef = getRelativeNotation(row, col);

      const displayValue = isEmptyCell(cellValue) ? '' : cellValue;
      output += '**' + escapeMarkdown(String(header)) + '** [' + absRef + ', ' + relRef + ']: ' + escapeMarkdown(String(displayValue)) + '\n';
    }
    output += '\n';
  }

  return output;
}

/**
 * Build column-based Markdown-KV format
 */
function buildColumnBasedMarkdown(range) {
  const meta = getRangeMeta(range);
  const values = range.getValues();

  let output = '---\n';
  output += 'Sheet: ' + meta.sheetName + '\n';
  output += 'Range: ' + meta.rangeNotation + '\n';
  output += 'Selection: ' + meta.numRows + ' rows × ' + meta.numCols + ' columns (' + meta.totalCells + ' cells)\n';

  const headerValues = [];
  for (let row = 0; row < meta.numRows; row++) {
    headerValues.push(values[row][0]);
  }
  output += 'Headers: ' + headerValues.join(', ') + '\n';
  output += 'Export Type: Column-based (Markdown-KV)\n';
  output += 'Timestamp: ' + meta.timestamp + '\n';
  output += 'Note: First column treated as headers. Each subsequent column is a record.\n';
  output += '---\n\n';

  for (let col = 2; col <= meta.numCols; col++) {
    const absColLetter = getAbsoluteColumnLetter(meta.startCol, col);
    const headerValue = values[0][col - 1];
    const colStartRow = meta.startRow;
    const colEndRow = meta.startRow + meta.numRows - 1;
    const actualRange = absColLetter + colStartRow + ':' + absColLetter + colEndRow;

    output += '## Column ' + col + ' (Sheet Column ' + absColLetter + '): ' + escapeMarkdown(String(headerValue)) + ' [' + actualRange + ']\n\n';

    for (let row = 1; row <= meta.numRows; row++) {
      const header = values[row - 1][0];
      const cellValue = values[row - 1][col - 1];
      const absRef = getCellA1Notation(meta.startRow, meta.startCol, row, col);
      const relRef = getRelativeNotation(row, col);

      const displayValue = isEmptyCell(cellValue) ? '' : cellValue;
      output += '**' + escapeMarkdown(String(header)) + '** [' + absRef + ', ' + relRef + ']: ' + escapeMarkdown(String(displayValue)) + '\n';
    }
    output += '\n';
  }

  return output;
}

/**
 * Build formula XML format
 */
function buildFormulasXML(range) {
  const meta = getRangeMeta(range);
  const values = range.getValues();
  const formulas = range.getFormulas();

  let output = '<?xml version="1.0" encoding="UTF-8"?>\n';
  output += '<!--\n';
  output += 'Sheet: ' + meta.sheetName + '\n';
  output += 'Range: ' + meta.rangeNotation + '\n';
  output += 'Selection: ' + meta.numRows + ' rows × ' + meta.numCols + ' columns (' + meta.totalCells + ' cells)\n';

  let formulaCount = 0;
  for (let row = 0; row < meta.numRows; row++) {
    for (let col = 0; col < meta.numCols; col++) {
      if (!isEmptyCell(formulas[row][col])) {
        formulaCount++;
      }
    }
  }

  output += 'Formula Cells: ' + formulaCount + ' (' + (meta.totalCells - formulaCount) + ' empty/value-only cells excluded)\n';
  output += 'Export Type: Formula-only (XML)\n';
  output += 'Timestamp: ' + meta.timestamp + '\n';
  output += '\n';
  output += 'IMPORTANT: This export includes ONLY cells containing formulas.\n';
  output += 'Empty cells and cells with static values are excluded.\n';
  output += '-->\n\n';
  output += '<formulas>\n';

  for (let row = 1; row <= meta.numRows; row++) {
    for (let col = 1; col <= meta.numCols; col++) {
      const formula = formulas[row - 1][col - 1];
      if (!isEmptyCell(formula)) {
        const absRef = getCellA1Notation(meta.startRow, meta.startCol, row, col);
        const relRef = getRelativeNotation(row, col);
        const rowNum = row;
        const colNum = col;

        output += '  <cell abs="' + absRef + '" rel="' + relRef + '" row="' + rowNum + '" col="' + colNum + '">\n';
        output += '    <formula>' + escapeXML(formula) + '</formula>\n';

        const rowHeader = values[0] ? values[0][col - 1] : null;
        const colHeader = values[row - 1] ? values[row - 1][0] : null;

        if (!isEmptyCell(rowHeader) || !isEmptyCell(colHeader)) {
          output += '    <context>\n';
          if (!isEmptyCell(rowHeader)) {
            output += '      <rowHeader>' + escapeXML(String(rowHeader)) + '</rowHeader>\n';
          }
          if (!isEmptyCell(colHeader)) {
            output += '      <colHeader>' + escapeXML(String(colHeader)) + '</colHeader>\n';
          }
          output += '    </context>\n';
        }

        output += '  </cell>\n';
      }
    }
  }

  output += '</formulas>\n';
  return output;
}

/**
 * Build general XML format (values only)
 */
function buildGeneralXML(range) {
  const meta = getRangeMeta(range);
  const values = range.getValues();

  let nonEmptyCount = 0;
  for (let row = 0; row < meta.numRows; row++) {
    for (let col = 0; col < meta.numCols; col++) {
      if (!isEmptyCell(values[row][col])) {
        nonEmptyCount++;
      }
    }
  }

  let output = '<?xml version="1.0" encoding="UTF-8"?>\n';
  output += '<!--\n';
  output += 'Sheet: ' + meta.sheetName + '\n';
  output += 'Range: ' + meta.rangeNotation + '\n';
  output += 'Selection: ' + meta.numRows + ' rows × ' + meta.numCols + ' columns (' + meta.totalCells + ' cells)\n';
  output += 'Non-Empty Cells: ' + nonEmptyCount + ' (' + (meta.totalCells - nonEmptyCount) + ' empty cells excluded)\n';
  output += 'Export Type: General (XML, values only)\n';
  output += 'Timestamp: ' + meta.timestamp + '\n';
  output += '\n';
  output += 'IMPORTANT: This export includes cell VALUES only (formulas are evaluated).\n';
  output += 'Empty cells are excluded to optimize token usage.\n';
  output += '-->\n\n';
  output += '<data>\n';

  for (let row = 1; row <= meta.numRows; row++) {
    for (let col = 1; col <= meta.numCols; col++) {
      const cellValue = values[row - 1][col - 1];
      if (!isEmptyCell(cellValue)) {
        const absRef = getCellA1Notation(meta.startRow, meta.startCol, row, col);
        const relRef = getRelativeNotation(row, col);

        output += '  <cell abs="' + absRef + '" rel="' + relRef + '" row="' + row + '" col="' + col + '">\n';
        output += '    <value>' + escapeXML(String(cellValue)) + '</value>\n';
        output += '  </cell>\n';
      }
    }
  }

  output += '</data>\n';
  return output;
}

/**
 * Shows a dialog that displays data for copying to clipboard
 */
function showClipboardDialog(data, dataType) {
  const htmlTemplate = HtmlService.createTemplateFromFile('ClipboardDialog');
  htmlTemplate.jsonData = data;
  htmlTemplate.dataType = dataType;

  const html = htmlTemplate.evaluate()
    .setWidth(500)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, 'Export Data');
}
