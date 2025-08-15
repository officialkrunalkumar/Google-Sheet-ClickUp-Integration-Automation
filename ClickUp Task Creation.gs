const SHEET_NAME = 'Active';
const CLICKUP_API_TOKEN = 'pk_87861544_WD49ESGN1U1EA4WYZR08TIDV9SWBYABV';
const LIST_ID = '6-901211526001-1';
const STATUS_HEADER = 'Status';
const TASK_ID_HEADER = 'Task ID (ClickUp)';
const TASK_URL_HEADER = 'Task Link (ClickUp)';
const PURSUING_VALUES = ['Moved to Clickup'];

const REQUIRED_HEADERS = [
  'Portal',
  'Tender Title',
  'Buyer/Issuing Authority',
  'Short Description',
  'Closing Date',
  'Estimated Value',
  'Link to Tender'
];

const DUE_TIME = { hour: 17, minute: 0, second: 0 };
const DEFAULT_STATUS = 'open';
const DEFAULT_PRIORITY = 2;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ClickUp')
    .addItem('Create task for current row', 'createTaskForActiveRow_')
    .addItem('Create tasks for all “Pursuing”', 'bulkCreateTasksForPursuing_')
    .addToUi();
}

function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== SHEET_NAME) return;
    const row = e.range.getRow();
    if (row === 1) return; // header

    const headers = getHeaderMap_(sh);
    ensureResultColumns_(sh, headers);

    const statusCol = headers[STATUS_HEADER];
    if (!statusCol) return;

    if (e.range.getColumn() !== statusCol) return;

    const values = getRowObject_(sh, row, headers);
    const newStatus = String(values[STATUS_HEADER] || '').trim().toLowerCase();

    if (!PURSUING_VALUES.includes(newStatus)) return;

    if (values[TASK_ID_HEADER] || values[TASK_URL_HEADER]) {
      sh.getParent().toast('Task already exists for this row.', 'ClickUp', 3);
      return;
    }

    const missing = REQUIRED_HEADERS.filter(h => values[h] === '' || values[h] == null);
    if (missing.length) {
      sh.getParent().toast('Missing fields: ' + missing.join(', '), 'ClickUp', 5);
      return;
    }

    const task = createClickUpTask_(values);

    sh.getRange(row, headers[TASK_ID_HEADER]).setValue(task.id || '');
    sh.getRange(row, headers[TASK_URL_HEADER]).setValue(task.url || '');
    sh.getParent().toast('ClickUp task created ✔', 'ClickUp', 3);
  } catch (err) {
    console.error(err);
    SpreadsheetApp.getActive().toast('Error creating task: ' + err.message, 'ClickUp', 8);
  }
}

function createTaskForActiveRow_() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_NAME) throw new Error('Switch to sheet: ' + SHEET_NAME);
  const row = sh.getActiveCell().getRow();
  if (row === 1) throw new Error('Select a data row, not the header.');

  const headers = getHeaderMap_(sh);
  ensureResultColumns_(sh, headers);
  const values = getRowObject_(sh, row, headers);

  const st = String(values[STATUS_HEADER] || '').trim().toLowerCase();
  if (!PURSUING_VALUES.includes(st)) throw new Error('Row status is not “Pursuing”.');

  if (values[TASK_ID_HEADER] || values[TASK_URL_HEADER]) {
    throw new Error('Task already exists for this row.');
  }

  const missing = REQUIRED_HEADERS.filter(h => values[h] === '' || values[h] == null);
  if (missing.length) throw new Error('Missing fields: ' + missing.join(', '));

  const task = createClickUpTask_(values);
  sh.getRange(row, headers[TASK_ID_HEADER]).setValue(task.id || '');
  sh.getRange(row, headers[TASK_URL_HEADER]).setValue(task.url || '');
  SpreadsheetApp.getActive().toast('ClickUp task created ✔', 'ClickUp', 3);
}

function bulkCreateTasksForPursuing_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const headers = getHeaderMap_(sh);
  ensureResultColumns_(sh, headers);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return;

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const statusColIdx = headers[STATUS_HEADER] - 1;
  const taskIdIdx = headers[TASK_ID_HEADER] - 1;
  const taskUrlIdx = headers[TASK_URL_HEADER] - 1;

  let created = 0, skipped = 0, errors = 0;

  for (let r = 0; r < data.length; r++) {
    try {
      const rowVals = data[r];
      const status = String(rowVals[statusColIdx] || '').trim().toLowerCase();
      const hasTask = !!(rowVals[taskIdIdx] || rowVals[taskUrlIdx]);
      if (!PURSUING_VALUES.includes(status) || hasTask) { skipped++; continue; }

      const values = getRowObjectFromArray_(rowVals, headers);
      const missing = REQUIRED_HEADERS.filter(h => values[h] === '' || values[h] == null);
      if (missing.length) { skipped++; continue; }

      const task = createClickUpTask_(values);
      sh.getRange(r + 2, taskIdIdx + 1).setValue(task.id || '');
      sh.getRange(r + 2, taskUrlIdx + 1).setValue(task.url || '');
      created++;
    } catch (e) {
      console.error('Row ' + (r + 2) + ': ' + e);
      errors++;
    }
  }

  SpreadsheetApp.getActive().toast(`Created: ${created}, Skipped: ${skipped}, Errors: ${errors}`, 'ClickUp', 6);
}

function createClickUpTask_(values) {
  if (!CLICKUP_API_TOKEN || CLICKUP_API_TOKEN === 'YOUR_API_TOKEN') {
    throw new Error('Set CLICKUP_API_TOKEN in the script CONFIG.');
  }
  if (!LIST_ID || LIST_ID === 'YOUR_LIST_ID') {
    throw new Error('Set LIST_ID in the script CONFIG.');
  }

  const title = buildTaskName_(values);
  const description = buildDescription_(values);
  const due = buildDueDateMs_(values['Closing Date']);

  const url = `https://api.clickup.com/api/v2/list/${encodeURIComponent(LIST_ID)}/task`;
  const payload = {
    name: title,
    description: description,
    status: DEFAULT_STATUS,
    priority: DEFAULT_PRIORITY,
    due_date: due,
    due_date_time: !!due
  };

  const options = {
    method: 'post',
    headers: {
      'Authorization': CLICKUP_API_TOKEN,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('ClickUp API error ' + code + ': ' + body);
  }
  const json = JSON.parse(body);
  return { id: json.id, url: json.url || `https://app.clickup.com/t/${json.id}` };
}

function buildTaskName_(values) {
  const title = values['Tender Title'] || 'Untitled Tender';
  const portal = values['Portal'] ? ` — ${values['Portal']}` : '';
  const closing = values['Closing Date'] ? ` (Closes: ${formatDate_(values['Closing Date'])})` : '';
  return title + portal + closing;
}

function buildDescription_(values) {
  const lines = [];
  lines.push(`Portal : ${values['Portal'] || '-'}`);
  lines.push(`Tender Title : ${values['Tender Title'] || '-'}`);
  lines.push(`Buyer/Issuing Authority : ${values['Buyer/Issuing Authority'] || '-'}`);
  lines.push(`Closing Date : ${values['Closing Date'] ? formatDate_(values['Closing Date']) : '-'}`);
  lines.push(`Estimated Value : ${values['Estimated Value'] || '-'}`);
  if (values['Link to Tender']) {
    lines.push(`Link : ${values['Link to Tender']}`);
  }
  lines.push('');
  lines.push(`Summary`);
  lines.push(values['Short Description'] || '-');
  return lines.join('\n');
}

function buildDueDateMs_(dateVal) {
  if (!(dateVal instanceof Date)) return undefined;
  const d = new Date(dateVal);
  d.setHours(DUE_TIME.hour || 17, DUE_TIME.minute || 0, DUE_TIME.second || 0, 0);
  return d.getTime();
}

function getHeaderMap_(sh) {
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { if (h) map[String(h).trim()] = i + 1; });

  if (!map[TASK_ID_HEADER]) {
    sh.getRange(1, lastCol + 1).setValue(TASK_ID_HEADER);
    map[TASK_ID_HEADER] = lastCol + 1;
  }
  if (!map[TASK_URL_HEADER]) {
    sh.getRange(1, sh.getLastColumn() + 1).setValue(TASK_URL_HEADER);
    map[TASK_URL_HEADER] = sh.getLastColumn();
  }
  return map;
}

function ensureResultColumns_(sh, headers) {
  return headers;
}

function getRowObject_(sh, row, headers) {
  const lastCol = sh.getLastColumn();
  const rowVals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  return getRowObjectFromArray_(rowVals, headers);
}
function getRowObjectFromArray_(rowVals, headers) {
  const obj = {};
  Object.keys(headers).forEach(h => {
    const col = headers[h] - 1;
    obj[h] = rowVals[col];
  });
  return obj;
}

function formatDate_(d) {
  if (!(d instanceof Date)) return String(d);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
