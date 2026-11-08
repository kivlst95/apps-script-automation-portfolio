/**
 * Budget payment requests.
 *
 * Every department works in its own spreadsheet (this script is bound to it).
 * Requests are submitted to one central register owned by the budget holder,
 * get a department-scoped ID (e.g. MKT_42) and are reviewed there. Review
 * statuses are mirrored back to the department spreadsheet.
 *
 * Department spreadsheet:
 *   "Add Request"      B1 = submit checkbox, A3:I = new request rows
 *   "Payment Request"  A  = delete checkbox, B:M = mirror of the register (read-only)
 * Central register:
 *   "Payments Requests" A = ID, B:J = request fields, K:L = review columns
 *   filled by the budget owner (K = approval decision)
 *
 * Triggers (run installTriggers once, as a user with edit access to the register):
 *   onBudgetSheetEdit  installable on-edit trigger
 *   refreshStatuses    time-driven, keeps the mirror up to date
 * Because the triggers run as their owner, department users never need
 * access to the central register.
 */

const SHEET_ADD_REQUEST = 'Add Request';
const SHEET_MY_REQUESTS = 'Payment Request';
const SHEET_REGISTER = 'Payments Requests';

const SUBMIT_CHECKBOX_A1 = 'B1';
const INPUT_FIRST_ROW = 3;
const INPUT_WIDTH = 9;               // columns A:I on "Add Request"
const REGISTER_WIDTH = 12;           // ID + 9 request fields + 2 review columns (A:L)
const REGISTER_APPROVAL_INDEX = 10;  // column K: once filled, the request is locked
const DEFAULT_REQUIRED_FIELDS = [0, 1, 3, 5, 6, 7];
const STATUS_REFRESH_MINUTES = 15;
const LOCK_TIMEOUT_MS = 30000;

/** Installable on-edit trigger: routes checkbox clicks to submit or delete. */
function onBudgetSheetEdit(e) {
  if (!e || !e.range || e.value !== 'TRUE') return;
  const range = e.range;
  const sheetName = range.getSheet().getName();
  const isSubmit = sheetName === SHEET_ADD_REQUEST && range.getA1Notation() === SUBMIT_CHECKBOX_A1;
  const isDelete = sheetName === SHEET_MY_REQUESTS && range.getRow() > 1 && range.getColumn() === 1;
  if (!isSubmit && !isDelete) return;

  try {
    if (isSubmit) submitRequests_(range.getSheet());
    else deleteRequest_(range);
  } catch (err) {
    Log.error('Budget request action failed', err, { sheet: sheetName, cell: range.getA1Notation() });
    alert_('Something went wrong and nothing was changed. Please try again or contact the budget owner.');
  } finally {
    range.setValue(false);
  }
}

/** Copies the department's rows from the register into "Payment Request". */
function refreshStatuses() {
  const config = getBudgetConfig_();
  const rows = readRegister_(openRegister_(config)).filter(function (row) {
    return String(row[0]).indexOf(config.prefix) === 0;
  });

  const target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MY_REQUESTS);
  const maxRows = target.getMaxRows();
  if (maxRows > 1) target.getRange(2, 2, maxRows - 1, REGISTER_WIDTH).clearContent();
  if (rows.length) target.getRange(2, 2, rows.length, REGISTER_WIDTH).setValues(rows);
  Log.debug('Statuses refreshed', { department: config.department, requests: rows.length });
}

/** Creates the on-edit and status-refresh triggers. Safe to run again. */
function installTriggers() {
  const handlers = ['onBudgetSheetEdit', 'refreshStatuses'];
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return handlers.indexOf(t.getHandlerFunction()) >= 0; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('onBudgetSheetEdit').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  ScriptApp.newTrigger('refreshStatuses').timeBased().everyMinutes(STATUS_REFRESH_MINUTES).create();
  Log.info('Triggers installed', { refreshMinutes: STATUS_REFRESH_MINUTES });
}

/** Validates the new rows, writes them to the register with IDs and clears the form. */
function submitRequests_(sheet) {
  const inputRange = sheet.getRange(INPUT_FIRST_ROW, 1, sheet.getMaxRows() - INPUT_FIRST_ROW + 1, INPUT_WIDTH);
  const entries = inputRange.getValues()
    .map(function (values, i) { return { values: values, sheetRow: INPUT_FIRST_ROW + i }; })
    .filter(function (entry) { return entry.values[0] !== ''; });

  if (!entries.length) return alert_('Nothing to submit: start each request in column A.');

  const required = getJsonProperty_('BUDGET_REQUIRED_FIELDS', DEFAULT_REQUIRED_FIELDS);
  const incomplete = entries
    .filter(function (entry) {
      return required.some(function (i) { return entry.values[i] === '' || entry.values[i] === null; });
    })
    .map(function (entry) { return entry.sheetRow; });
  if (incomplete.length) {
    return alert_('Not enough information in row(s) ' + incomplete.join(', ') + '. Nothing was submitted.');
  }

  const config = getBudgetConfig_();
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);
  let ids;
  try {
    const register = openRegister_(config);
    let lastId = getLastDepartmentId_(register, config.prefix);
    ids = entries.map(function (entry) {
      lastId += 1;
      const id = config.prefix + lastId;
      register.appendRow([id].concat(entry.values)); // appendRow is atomic per row
      return id;
    });
    SpreadsheetApp.flush();
    inputRange.clearContent();
  } finally {
    lock.releaseLock();
  }

  Log.info('Requests submitted', { department: config.department, ids: ids });
  refreshStatuses();
  SpreadsheetApp.getActiveSpreadsheet().toast('Submitted: ' + ids.join(', '), 'Budget requests');
}

/** Deletes a request from the register, but only before the budget owner reviewed it. */
function deleteRequest_(checkbox) {
  const requestId = String(checkbox.offset(0, 1).getValue()).trim();
  if (!requestId) return alert_('Found nothing to delete in this row.');

  const config = getBudgetConfig_();
  if (requestId.indexOf(config.prefix) !== 0) return alert_('Request ' + requestId + ' belongs to another department.');

  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);
  try {
    const register = openRegister_(config);
    const rows = readRegister_(register);
    const index = rows.findIndex(function (row) { return String(row[0]) === requestId; });
    if (index < 0) return alert_('Request ' + requestId + ' was not found in the register.');
    if (rows[index][REGISTER_APPROVAL_INDEX] !== '') {
      return alert_('Request ' + requestId + ' has already been reviewed and cannot be deleted.');
    }

    // Re-check the exact row right before deleting: another department may have changed the register.
    const rowNumber = index + 2;
    if (String(register.getRange(rowNumber, 1).getValue()) !== requestId) {
      throw new Error('Register changed during delete of ' + requestId);
    }
    register.deleteRow(rowNumber);
  } finally {
    lock.releaseLock();
  }

  Log.info('Request deleted', { department: config.department, id: requestId });
  refreshStatuses();
}

/** Highest numeric ID already used by this department, 0 if none. */
function getLastDepartmentId_(register, prefix) {
  return readRegister_(register).reduce(function (max, row) {
    const id = String(row[0]);
    if (id.indexOf(prefix) !== 0) return max;
    const n = parseInt(id.slice(prefix.length), 10);
    return Number.isFinite(n) && n > max ? n : max;
  }, 0);
}

function readRegister_(register) {
  const lastRow = register.getLastRow();
  return lastRow < 2 ? [] : register.getRange(2, 1, lastRow - 1, REGISTER_WIDTH).getValues();
}

function openRegister_(config) {
  const sheet = SpreadsheetApp.openById(config.registerId).getSheetByName(SHEET_REGISTER);
  if (!sheet) throw new Error('Sheet "' + SHEET_REGISTER + '" not found in the register spreadsheet');
  return sheet;
}

function getBudgetConfig_() {
  const department = getRequiredProperty_('BUDGET_DEPARTMENT_CODE');
  return {
    registerId: getRequiredProperty_('BUDGET_REGISTER_SPREADSHEET_ID'),
    department: department,
    prefix: department + '_'
  };
}

function alert_(message) {
  SpreadsheetApp.getUi().alert(message);
}
