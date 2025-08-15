const SOURCE_SHEET = 'Active';
const TARGET_SHEET = 'Archived';
const HEADER_ROWS  = 1;
const DATE_COL     = 7;
const STATUS_COL   = 5;
const DONE_VALUE   = 'false';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡Row Tools')
    .addItem('Move overdue rows now', 'moveOverdueRows')
    .addToUi();
}

function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== SOURCE_SHEET) return;
    if (e.range.getRow() <= HEADER_ROWS) return;
    const row = e.range.getRow();
    const lastCol = sh.getLastColumn();
    const rowValues = sh.getRange(row, 1, 1, lastCol).getValues()[0];
    const status = String(rowValues[STATUS_COL - 1] || '').trim();
    const dateVal = rowValues[DATE_COL - 1];
    const today = new Date(); today.setHours(0,0,0,0);
    const isDate = dateVal instanceof Date;
    const due = isDate && dateVal.setHours(0,0,0,0) <= today.getTime();
    const isDone = status.toLowerCase() === DONE_VALUE.toLowerCase();
    if (!(due || isDone)) return;
    moveRow_(rowValues);
    sh.deleteRow(row);
  } catch (err) {
    console.error(err);
  }
}

function moveOverdueRows() {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SOURCE_SHEET);
  const dst = ss.getSheetByName(TARGET_SHEET);
  if (!src || !dst) throw new Error('Check SOURCE_SHEET and TARGET_SHEET names.');

  const lastRow = src.getLastRow();
  const lastCol = src.getLastColumn();
  if (lastRow <= HEADER_ROWS) return;
  const range = src.getRange(HEADER_ROWS + 1, 1, lastRow - HEADER_ROWS, lastCol);
  const values = range.getValues();
  const today = new Date(); today.setHours(0,0,0,0);
  for (let i = values.length - 1; i >= 0; i--) {
    const rowValues = values[i];
    const status = String(rowValues[STATUS_COL - 1] || '').trim();
    const dateVal = rowValues[DATE_COL - 1];
    const isDate = dateVal instanceof Date;
    const due = isDate && dateVal.setHours(0,0,0,0) <= today.getTime();
    const isDone = status.toLowerCase() === DONE_VALUE.toLowerCase();
    if (due || isDone) {
      moveRow_(rowValues, dst);
      src.deleteRow(HEADER_ROWS + 1 + i);
    }
  }
}

function moveRow_(rowValues, dstSheet) {
  const ss = SpreadsheetApp.getActive();
  const dst = dstSheet || ss.getSheetByName(TARGET_SHEET);
  const lastCol = ss.getSheetByName(SOURCE_SHEET).getLastColumn();
  if (dst.getLastRow() === 0) dst.insertRows(1);
  if (dst.getLastRow() < HEADER_ROWS) {
    const src = ss.getSheetByName(SOURCE_SHEET);
    const headers = src.getRange(1, 1, HEADER_ROWS, lastCol).getValues();
    if (dst.getLastRow() === 0) dst.insertRows(1, HEADER_ROWS);
    dst.getRange(1, 1, HEADER_ROWS, lastCol).setValues(headers);
  }
  const targetRow = dst.getLastRow() + 1;
  dst.getRange(targetRow, 1, 1, lastCol).setValues([rowValues]);
}
