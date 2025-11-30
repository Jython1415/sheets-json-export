/**
 * Adds a custom menu to Google Sheets when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Export to JSON')
    .addItem('Copy Values (JSON)', 'exportValuesJSON')
    .addItem('Copy Formulas (JSON)', 'exportFormulasJSON')
    .addItem('Copy Both (JSON)', 'exportBothJSON')
    .addToUi();
}

/**
 * Exports the selected range values as JSON to clipboard
 */
function exportValuesJSON() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select a range first.');
    return;
  }
  
  const data = buildValuesJSON(range);
  showClipboardDialog(data, 'Values');
}

/**
 * Exports the selected range formulas as JSON to clipboard
 */
function exportFormulasJSON() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select a range first.');
    return;
  }
  
  const data = buildFormulasJSON(range);
  showClipboardDialog(data, 'Formulas');
}

/**
 * Exports both values and formulas for the selected range as JSON to clipboard
 */
function exportBothJSON() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select a range first.');
    return;
  }
  
  const data = buildCombinedJSON(range);
  showClipboardDialog(data, 'Values and Formulas');
}

/**
 * Builds JSON structure for values only
 * @param {Range} range - The active range
 * @return {string} JSON string
 */
function buildValuesJSON(range) {
  const sheet = range.getSheet();
  const values = range.getValues();
  
  const output = {
    metadata: {
      sheetName: sheet.getName(),
      range: range.getA1Notation(),
      dimensions: {
        rows: range.getNumRows(),
        columns: range.getNumColumns()
      },
      exportedAt: new Date().toISOString(),
      exportType: 'values'
    },
    data: values
  };
  
  return JSON.stringify(output, null, 2);
}

/**
 * Builds JSON structure for formulas only
 * @param {Range} range - The active range
 * @return {string} JSON string
 */
function buildFormulasJSON(range) {
  const sheet = range.getSheet();
  const formulas = range.getFormulas();
  
  const output = {
    metadata: {
      sheetName: sheet.getName(),
      range: range.getA1Notation(),
      dimensions: {
        rows: range.getNumRows(),
        columns: range.getNumColumns()
      },
      exportedAt: new Date().toISOString(),
      exportType: 'formulas'
    },
    data: formulas
  };
  
  return JSON.stringify(output, null, 2);
}

/**
 * Builds JSON structure for both values and formulas
 * @param {Range} range - The active range
 * @return {string} JSON string
 */
function buildCombinedJSON(range) {
  const sheet = range.getSheet();
  const values = range.getValues();
  const formulas = range.getFormulas();
  
  const output = {
    metadata: {
      sheetName: sheet.getName(),
      range: range.getA1Notation(),
      dimensions: {
        rows: range.getNumRows(),
        columns: range.getNumColumns()
      },
      exportedAt: new Date().toISOString(),
      exportType: 'combined'
    },
    values: values,
    formulas: formulas
  };
  
  return JSON.stringify(output, null, 2);
}

/**
 * Shows a dialog that copies the JSON data to clipboard
 * @param {string} jsonData - The JSON data to copy
 * @param {string} dataType - Description of what type of data is being copied
 */
function showClipboardDialog(jsonData, dataType) {
  const htmlTemplate = HtmlService.createTemplateFromFile('ClipboardDialog');
  htmlTemplate.jsonData = jsonData;
  htmlTemplate.dataType = dataType;
  
  const html = htmlTemplate.evaluate()
    .setWidth(500)
    .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Export to JSON');
}
