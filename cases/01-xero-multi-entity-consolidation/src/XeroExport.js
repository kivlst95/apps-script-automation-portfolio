/**
 * ================================================================
 *  XERO -> GOOGLE SHEETS: MULTI-ENTITY EXPORT AND REPORTS
 *  Second file of the project. The first one is the OAuth module (XeroAuth.js).
 * ================================================================
 *
 *  Uses from the auth module:
 *      XERO_COMPANIES, xeroFetch(), getXeroService_(),
 *      fetchXeroConnections_(), syncXeroTenants(),
 *      readXeroTenants_(), resetXero()
 *
 *  REQUIRED SCOPES (XERO_SCOPES in the auth module)
 *      accounting.invoices.read                  invoices, bills, credit notes
 *      accounting.payments.read                  payments, prepayments, overpayments
 *      accounting.banktransactions.read          bank transactions and transfers
 *      accounting.contacts.read                  contacts and their balances
 *      accounting.settings.read                  chart of accounts
 *      accounting.reports.trialbalance.read      balances at a date
 *      accounting.reports.profitandloss.read     P&L sections (otherwise mapToPnl: false)
 *      accounting.manualjournals.read            manual journals
 *      accounting.attachments.read               document attachments
 *
 *  The /Journals endpoint requires the Xero Advanced plan, so
 *  includeJournals is false by default. Missing P&L items (FX differences,
 *  depreciation, rounding) are taken from Reports/ProfitAndLoss instead.
 *  A 401 AuthorizationUnsuccessful means a missing scope, not a bad token.
 *
 *  PAYMENT SPLIT BY LINE ITEM
 *  A payment of 500 against an invoice with lines "Ads 300 / IT 200"
 *  becomes two rows (300 Ads, 200 IT) in Payments and BankTransactions,
 *  so cash lands in the same P&L sections as the document.
 *  Diagnostics: diagnosePaymentSplit(), countUnsplitPayments().
 *
 *  SHEETS
 *    Documents, Lines, BankTransactions, Payments, Balances, Bank_Balances,
 *    Aging, AR_Aged_Report, AP_Aged_Report, AR_AP_Summary,
 *    Cashflow_By_Category, PnL_Report (source of truth), PnL_Adjustments,
 *    PnL (transaction layer), Balance_Sheet, Balance_Sheet_Pivot,
 *    Bank_Overview, Bank_Unreconciled, Attachments, _Export_Log, checks
 *
 *  THREE RUN MODES
 *    testExport(7)        end-to-end check for a short period, writes to
 *                         TEST_ sheets only, one entity, no triggers
 *    startExport()        full reload from CONFIG.fromDate, runs on triggers
 *                         and survives the daily API quota
 *    incrementalExport()  daily update: only records changed since the
 *                         last sync, updated in place
 *
 *  XERO LIMITS: 60 requests/minute and 5000/day per organisation (rolling
 *  24h window). The initial history load does not fit: the export pauses
 *  when the quota runs low and resumes by trigger. State is persisted, so
 *  the load continues over several days instead of restarting.
 *
 *  IF THE EXPORT IS STUCK
 *    exportStatus()       which job stopped and with which error
 *    runExport()          continue from the same point
 *    resetExportState()   reset, then start again with startExport()
 *
 *  Note: inline comments are partly in Russian (original working language).
 * ================================================================
 */

// ================================================================
//  1. CONFIGURATION
// ================================================================

const CONFIG = {
  /** Spreadsheet ID. Empty = active spreadsheet. */
  spreadsheetId: '',

  /** Date for balances and aging. Empty = today. */
  balanceDate: '',

  /** Start date for documents and transactions. */
  fromDate: '2026-01-01',

  // ---- export scope ----

  /** Document line items to the Lines sheet. Needed to split payments by category. */
  includeLineItems: true,

  /** Credit notes, prepayments, overpayments. Without them AR/AP will not reconcile. */
  includeCreditNotes: true,
  includePrepayments: true,

  /** Payments against documents. */
  includePayments: true,

  /**
   * General ledger /Journals. Needed for a correct P&L: FX
   * differences, manual journals, depreciation and rounding have no document
   * lines and never reach the Lines sheet.
   * Requires scope accounting.journals.read.
   */
  includeJournals: false,

  /**
   * Journals do not support a date filter: pagination runs by offset
   * from the start of history. Entries before fromDate are dropped after
   * retrieval, so on a long history the first pages are wasted.
   * Speeds up repeated runs: journal number to start from.
   */
  journalsStartOffset: 0,

  /**
   * Export only journal entries that are missing from other sheets.
   * From payment journals the FX difference and rounding lines are kept:
   * a realised FX difference has no journal type of its own,
   * it arrives as an extra line inside ACCRECPAYMENT/ACCPAYPAYMENT.
   * false = export the whole general ledger.
   */
  journalsOnlyMissing: true,

  /**
   * Missing P&L items from the Reports/ProfitAndLoss report.
   * Calculates the difference between the account total in the Xero report and
   * what was collected from documents; the remainder is what is missing in Lines:
   * FX differences, depreciation, rounding, payroll.
   * Monthly, without transaction detail: this is how the report works.
   */
  includePnlAdjustments: true,

  /**
   * standardLayout returns a SUMMARY layout: accounts are collapsed
   * into groups and rows carry no AccountID. That hurts twice:
   * account detail is lost and adjustments cannot be matched.
   * Keep false so the report comes back by account.
   */
  reportsStandardLayout: false,

  /** Balance sheet. How many periods back and at which step. Xero: 1 to 11. */
  includeBalanceSheet: true,
  balanceSheetPeriods: 11,
  balanceSheetTimeframe: 'MONTH',      // MONTH | QUARTER | YEAR

  /**
   * Payments as rows in the BankTransactions sheet as well.
   * Otherwise the sheet does not show all account movements: invoice payments
   * go through /Payments, not through /BankTransactions.
   */
  paymentsIntoBankSheet: true,

  /**
   * Split a payment by the line items of the paid document.
   * Requires includeLineItems: true.
   */
  splitPaymentsByCategory: true,

  /**
   * If the document lines are not in the Lines sheet (document older than fromDate),
   * fetch the document. +1 request per document, result is cached.
   */
  resolvePaymentCategories: true,

  /** Maximum extra document requests per run and entity. */
  resolveCategoriesLimit: 500,

  /** Internal transfers between own accounts (two rows). */
  includeBankTransfers: true,

  /** Do not duplicate transfers with SPEND-TRANSFER / RECEIVE-TRANSFER legs. */
  skipTransferTypes: true,

  // ---- extra requests (cost quota) ----

  /** Fetch a transaction again if Xero returned no lines. +1 request per transaction. */
  resolveContraAccounts: true,

  /** Both legs of a transfer: exact amounts and currencies. +2 requests per transfer. */
  resolveTransferAmounts: true,

  /** Who reconciled. Expensive: +1 request per reconciled transaction. */
  fetchReconciledBy: false,

  /** Maximum history requests per run and entity. */
  reconciledByLimit: 300,

  /** Fetch history only for transactions not older than this date. Empty = all. */
  reconciledByFrom: '',

  // ---- reports and checks ----

  /** Tag transactions with P&L sections. Requires scope profitandloss.read. */
  mapToPnl: true,

  /** Rebuild reports right after the export. */
  buildReportsAfterExport: true,

  /** Verify record counts after the export. */
  verifyAfterExport: true,

  // ---- technical ----

  /** Records per page. Responses with line items are heavy, so use fewer. */
  pageSize: 500,
  pageSizeWithLines: 100,

  /** Guard against endless pagination. */
  maxPages: 200,

  /** Pause before the 6-minute limit. */
  maxRuntimeMs: 4.5 * 60 * 1000,

  /**
   * Stop when less than this reserve of the daily quota is left.
   * Better to stop ourselves than hit a 429 in the middle of a page:
   * the position is not lost and no requests are wasted.
   */
  dailyQuotaFloor: 150,

  /**
   * Hours to wait before resuming after hitting the daily limit.
   * The window is rolling: requests are released gradually, so
   * there is no need to wait until midnight.
   */
  resumeAfterQuotaHours: 6,

  /**
   * How far back to look by TRANSACTION DATE for entities
   * that cannot be filtered by modification date (transfers,
   * prepayments). In incremental mode they use a date window
   * and changes are filtered on the script side.
   * Edits to a transfer older than this window are not picked up.
   */
  incrementalLookbackDays: 180,

  // ---- attachments ----

  /**
   * Google Drive folder for attachments. Empty = a folder
   * Xero Attachments is created in the root, with a subfolder per entity.
   */
  attachmentsFolderId: getProperty_('XERO_ATTACHMENTS_FOLDER_ID', ''),

  /** Document types to fetch attachments for. */
  attachmentDocTypes: {
    Bill: 1,
    Invoice: 1
  },

  /**
   * Maximum documents per run. Each document needs at least
   * two requests (list plus file), and downloads are slower
   * than regular responses.
   */
  attachmentsPerRun: 200,

  /** Skip files larger than this, MB. */
  attachmentMaxSizeMb: 20,

  /**
   * Time reserve for attachments, smaller than the overall maxRuntimeMs.
   * A file download takes much longer than a regular response, and one
   * heavy file at the end of a run can easily exceed the execution limit:
   * then everything not yet written is lost.
   */
  attachmentsRuntimeMs: 3.5 * 60 * 1000,

  /** Fetch attachments at the end of the incremental update. */
  attachmentsInIncremental: true,

  /**
   * Remove the sheet filter before writing.
   *
   * A basic filter in Google Sheets is more than a view: sorting
   * through it physically reorders rows, and its range does not grow
   * with the data. During bulk rewrites this corrupts the content.
   *
   * If filters are needed permanently, use Filter views
   * (Data > Change filter views): they live separately from the data,
   * do not reorder rows and are not affected by this setting.
   */
  dropFiltersBeforeWrite: true,

  /** Aging buckets in days. */
  agingBuckets: [30, 60, 90]
};

const PROP_STATE = 'XERO_EXPORT_STATE';
const PROP_SEEN  = 'XERO_EXPORT_SEEN';
const PROP_SYNC  = 'XERO_LAST_SYNC';

/**
 * Sheet name prefix. The test run sets TEST_ so that
 * production data is not touched. Empty in other modes.
 */
var SHEET_PREFIX = '';

/**
 * Time of the last sync. When set, the export filters
 * by UpdatedDateUTC instead of the document date, i.e. fetches only
 * what changed. Set by the incremental mode.
 */
var SYNC_SINCE = null;

/** Collect rows in memory instead of writing to the sheet; needed for upsert. */
var WRITE_BUFFER = null;

// ================================================================
//  2. MENU
// ================================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Xero')
    .addItem('Authorize', 'menuAuth')
    .addItem('Connection status', 'menuStatus')
    .addSeparator()
    .addItem('Show organisations', 'menuConnections')
    .addItem('Link organisations', 'menuSyncTenants')
    .addSeparator()
    .addItem('One-week test', 'menuTestExport')
    .addItem('Full export', 'menuFullExport')
    .addItem('Update changes', 'incrementalExport')
    .addItem('Balances only', 'exportBalancesOnly')
    .addItem('Bank only', 'exportBankOnly')
    .addSeparator()
    .addItem('Export status', 'menuExportStatus')
    .addItem('Resume export', 'runExport')
    .addItem('Reset export', 'menuResetExport')
    .addSeparator()
    .addItem('Check completeness', 'verifyExport')
    .addItem('Check payment split', 'verifyLineAllocation')
    .addItem('Reconcile account movements', 'verifyBankMovement')
    .addItem('Unsplit payments', 'countUnsplitPayments')
    .addItem('Set attachment flags', 'backfillAttachmentFlags')
    .addItem('Download attachments', 'syncAttachments')
    .addItem('Pending attachments', 'attachmentsPending')
    .addItem('Account balances', 'buildBankOverview')
    .addItem('P&L transactions', 'buildPnl')
    .addItem('Missing P&L items', 'buildPnlAdjustments')
    .addItem('P&L report only', 'exportPnlReportOnly')
    .addItem('Monthly balance sheet', 'buildBalanceSheetPivot')
    .addItem('Balance sheet only', 'exportBalanceSheetOnly')
    .addItem('Rebuild reports', 'buildAllReports')
    .addSeparator()
    .addItem('Remaining Xero quota', 'showApiQuota')
    .addItem('Diagnostics', 'menuDiagnose')
    .addItem('Redirect URI', 'menuRedirectUri')
    .addItem('Reset authorization', 'menuReset')
    .addToUi();
}

function menuAuth() {
  showLink_('Xero authorization',
    'Open the link and choose an organisation. For three organisations repeat three times.',
    getXeroService_().getAuthorizationUrl());
}

function menuRedirectUri() { alert_('Redirect URI', getXeroService_().getRedirectUri()); }

function menuStatus() {
  const service = getXeroService_();
  if (!service.hasAccess()) {
    alert_('No access', service.getLastError() || 'Authorization required.');
    return;
  }
  const map = readXeroTenants_();
  alert_('Status', 'Access granted.\n\n' + Object.keys(XERO_COMPANIES).map(function (a) {
    return a + ' — ' + (map[a] ? map[a].tenantName + ' ✓' : 'NOT LINKED');
  }).join('\n'));
}

function menuConnections() {
  alert_('App organisations',
    fetchXeroConnections_().map(function (c) { return c.tenantName; }).join('\n') +
    '\n\nCompare with the XERO_COMPANIES constant in the auth module.');
}

function menuSyncTenants() {
  const map = syncXeroTenants();
  alert_('Linking', Object.keys(XERO_COMPANIES).map(function (a) {
    return a + ' → ' + (map[a] ? map[a].tenantName : 'NOT FOUND');
  }).join('\n'));
}

function menuFullExport() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_STATE)) {
    const ui = SpreadsheetApp.getUi();
    const answer = ui.alert('An unfinished export exists',
      'Start over? All sheets will be overwritten.\n\n' +
      'NO = continue the previous export from where it stopped.',
      ui.ButtonSet.YES_NO_CANCEL);
    if (answer === ui.Button.CANCEL) return;
    if (answer === ui.Button.NO) { runExport(); return; }
  }
  startExport();
  alert_('Export', 'Started. If it does not finish within 6 minutes, it continues by trigger.');
}

function menuTestExport() {
  testExport(7);
  alert_('Test', 'Done. Results are in the TEST_* sheets, details in the log.');
}

function menuExportStatus() { alert_('Export status', exportStatusText_()); }

function menuResetExport() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Reset export', 'Clear the state and remove resume triggers?',
      ui.ButtonSet.YES_NO) === ui.Button.YES) {
    resetExportState();
    alert_('Reset', 'Done. Next: Full export.');
  }
}

function menuDiagnose() {
  xeroDiagnose();
  alert_('Diagnostics', 'Result in the log: Extensions > Apps Script > Executions.');
}

function menuReset() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Reset', 'Delete tokens and links? You will need to authorize again.',
      ui.ButtonSet.YES_NO) === ui.Button.YES) {
    resetXero();
    alert_('Reset', 'Done.');
  }
}

function alert_(title, msg) {
  SpreadsheetApp.getUi().alert(title, msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function showLink_(title, text, url) {
  const html = HtmlService.createHtmlOutput(
    '<div style="font:14px/1.6 -apple-system,Segoe UI,sans-serif;padding:16px">' +
    text + '<br><br><a href="' + url + '" target="_blank">Open Xero</a></div>'
  ).setWidth(420).setHeight(160);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

// ================================================================
//  3. EXPORT STATE
// ================================================================

/** Reset a stuck export: state, duplicate registry, triggers. */
function resetExportState() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_STATE);
  props.deleteProperty(PROP_SEEN);
  stopExport();
  Log.info('Export state cleared, resume triggers removed.');
}

function exportStatus() { Log.info(exportStatusText_()); }

function exportStatusText_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(PROP_STATE);
  const lines = [];

  if (!raw) {
    lines.push('No active export.');
  } else {
    const s = JSON.parse(raw);
    const job = s.jobs[s.index] || {};
    lines.push('Job ' + (s.index + 1) + ' of ' + s.jobs.length +
               ': ' + (job.company || '—') + ' / ' + (job.task || '—') +
               ', page ' + s.page);
    lines.push('Started: ' + new Date(s.startedAt).toISOString());
    if (s.quotaPauses) lines.push('Pauses due to daily quota: ' + s.quotaPauses);
    if (s.lastError) lines.push('Last error: ' + s.lastError);
  }

  const triggers = ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'runExport'; });
  lines.push('Resume triggers: ' + triggers.length +
    (triggers.length > 1 ? '  <- EXTRA, parallel runs break the refresh token' : ''));

  return lines.join('\n');
}

// ================================================================
//  4. DEDUPLICATION
// ================================================================

/**
 * With date-filtered pagination Xero can shift the result set between
 * pages: a record on the boundary arrives twice or gets lost.
 * Duplicates are removed here, losses are caught by verifyExport().
 *
 * The registry lives ONLY in the memory of one execution. It must not go
 * to Script Properties: one value is limited to 9 KB, about 200
 * IDs. When exceeded, setProperty throws outside
 * the try/catch around the job: the execution fails entirely, the state
 * remains, no resume trigger is set and the last jobs
 * in the queue simply never run.
 *
 * Duplicates that slip through at a pause boundary are removed by
 * dedupeSheets() at the end of the export.
 */
function seenLoad_() { return {}; }
function seenSave_() { /* intentionally does nothing */ }

function seenFilter_(seen, bucket, items, idField) {
  if (!seen[bucket]) seen[bucket] = {};
  const store = seen[bucket];
  const out = [];
  items.forEach(function (item) {
    const id = item[idField];
    if (!id || store[id]) return;
    store[id] = 1;
    out.push(item);
  });
  return out;
}

// ================================================================
//  5. CHART OF ACCOUNTS
// ================================================================

const ACCOUNTS_CACHE = {};

function accounts_(company) {
  if (ACCOUNTS_CACHE[company]) return ACCOUNTS_CACHE[company];
  const map = {};
  (xeroFetch(company, 'Accounts').Accounts || []).forEach(function (a) {
    const entry = { name: a.Name, code: a.Code, class: a.Class || '', type: a.Type || '',
                    system: a.SystemAccount || '' };
    if (a.Code) map[String(a.Code)] = entry;
    if (a.AccountID) map[a.AccountID] = entry;
  });
  ACCOUNTS_CACHE[company] = map;
  return map;
}

// ================================================================
//  6. EXPORT ENGINE
// ================================================================

function startExport() {
  stopExport();

  // Order matters: payments after invoices, otherwise the payment split
  // by category finds no document lines
  const jobs = [];
  Object.keys(XERO_COMPANIES).forEach(function (company) {
    jobs.push({ company: company, task: 'balances' });
    jobs.push({ company: company, task: 'invoices' });
    if (CONFIG.includeCreditNotes) jobs.push({ company: company, task: 'creditNotes' });
    if (CONFIG.includePrepayments) {
      jobs.push({ company: company, task: 'prepayments' });
      jobs.push({ company: company, task: 'overpayments' });
    }
    jobs.push({ company: company, task: 'bankTransactions' });
    if (CONFIG.includeBankTransfers) jobs.push({ company: company, task: 'bankTransfers' });
    if (CONFIG.includePayments) jobs.push({ company: company, task: 'payments' });
    if (CONFIG.includeJournals) jobs.push({ company: company, task: 'journals' });
    if (CONFIG.includeBalanceSheet) jobs.push({ company: company, task: 'balanceSheet' });
    if (CONFIG.includePnlAdjustments) jobs.push({ company: company, task: 'pnlReport' });
  });

  Object.keys(SCHEMA).forEach(resetSheet_);

  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_STATE, JSON.stringify({
    jobs: jobs, index: 0, page: 1, startedAt: Date.now(), balanceDate: balanceDate_()
  }));

  runExport();
}

/**
 * Runs jobs while there is time left.
 *
 * LockService is mandatory: two parallel runs rotate the Xero refresh token
 * and break the connection; the old token dies right after the exchange.
 * An error inside a job does not hang the export: the reason is written
 * to the state, triggers are removed and a human decides what next.
 */
function runExport() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Log.info('Another run is in progress; exiting to avoid breaking the token.');
    return;
  }

  try {
    const started = Date.now();
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(PROP_STATE);
    if (!raw) { Log.info('No active export. Run startExport().'); return; }

    const state = JSON.parse(raw);
    const seen = seenLoad_();

    while (state.index < state.jobs.length) {
      if (Date.now() - started > CONFIG.maxRuntimeMs) {
        props.setProperty(PROP_STATE, JSON.stringify(state));
        scheduleContinuation_();
        Log.info('Pause: job %s/%s.', state.index + 1, state.jobs.length);
        return;
      }

      const job = state.jobs[state.index];
      const before = countSheetTotal_();
      let done;

      try {
        switch (job.task) {
          case 'balances':         done = exportBalances_(job.company, state.balanceDate); break;
          case 'invoices':         done = exportInvoices_(job.company, state, seen);       break;
          case 'creditNotes':      done = exportCreditNotes_(job.company, state, seen);    break;
          case 'prepayments':      done = exportPrepay_(job.company, state, seen, 'Prepayments');  break;
          case 'overpayments':     done = exportPrepay_(job.company, state, seen, 'Overpayments'); break;
          case 'bankTransactions': done = exportBankTx_(job.company, state, seen);         break;
          case 'bankTransfers':    done = exportTransfers_(job.company, seen);             break;
          case 'payments':         done = exportPayments_(job.company, state, seen);       break;
          case 'journals':         done = exportJournals_(job.company, state, seen);       break;
          case 'balanceSheet':     done = exportBalanceSheet_(job.company, state.balanceDate); break;
          case 'pnlReport':        done = exportPnlReport_(job.company, state.balanceDate); break;
          default: throw new Error('Unknown job: ' + job.task);
        }
      } catch (err) {
        // The daily quota is not an error but an expected pause during the initial
        // load: save the position and continue in a few hours
        if (err.quotaExhausted) {
          state.quotaPauses = (state.quotaPauses || 0) + 1;
          props.setProperty(PROP_STATE, JSON.stringify(state));
          scheduleResume_(CONFIG.resumeAfterQuotaHours);
          Log.warn('Daily quota exhausted at %s / %s page %s. ' +
                     'Resuming in %s h. Quota pauses: %s.',
            job.company, job.task, state.page,
            CONFIG.resumeAfterQuotaHours, state.quotaPauses);
          return;
        }

        // Other errors are not retried automatically: if the cause is the token
        // or a scope, retries would burn the quota without result
        state.lastError = job.company + '/' + job.task + ' page ' + state.page + ': ' + err.message;
        props.setProperty(PROP_STATE, JSON.stringify(state));
        stopExport();

        Log.error('STOPPED. %s', state.lastError);
        Log.info('Continue: runExport(). Start over: resetExportState(), then startExport().');
        if (/401|Unauthorized|AuthorizationUnsuccessful/i.test(err.message)) {
          Log.info('If previous jobs succeeded, the cause is the scope for this endpoint, ' +
                     'not the token. Check XERO_SCOPES and authorize again.');
        }
        return;
      }

      // Shows how many rows each job returned: an empty job
      // is immediately visible
      Log.info('%s / %s page %s -> +%s rows%s',
        job.company, job.task, state.page,
        countSheetTotal_() - before, done ? '  [job completed]' : '');

      if (done) { state.index++; state.page = 1; state.offset = 0; }
      props.setProperty(PROP_STATE, JSON.stringify(state));

      // Stop early instead of waiting for a 429
      if (quotaLow_()) {
        state.quotaPauses = (state.quotaPauses || 0) + 1;
        props.setProperty(PROP_STATE, JSON.stringify(state));
        scheduleResume_(CONFIG.resumeAfterQuotaHours);
        Log.warn('Remaining daily quota %s (threshold %s): pausing for %s h. ' +
                   'Job %s/%s. Quota pauses: %s.',
          LAST_QUOTA.day, CONFIG.dailyQuotaFloor, CONFIG.resumeAfterQuotaHours,
          state.index + 1, state.jobs.length, state.quotaPauses);
        return;
      }
    }

    props.deleteProperty(PROP_STATE);
    stopExport();
    Log.info('All jobs completed in %s s.', Math.round((Date.now() - state.startedAt) / 1000));

    dedupeSheets();
    markSynced();                     // from now on incremental updates are possible
    if (CONFIG.buildReportsAfterExport) buildAllReports();
    if (CONFIG.verifyAfterExport) { verifyExport(); countUnsplitPayments(); }

  } finally {
    try { lock.releaseLock(); } catch (err) { /* already released */ }
  }
}

function stopExport() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runExport') ScriptApp.deleteTrigger(t);
  });
}

function scheduleContinuation_() {
  stopExport();
  ScriptApp.newTrigger('runExport').timeBased().after(60 * 1000).create();
}

/** Resume in a few hours after hitting the daily quota. */
function scheduleResume_(hours) {
  stopExport();
  ScriptApp.newTrigger('runExport').timeBased()
    .after(Math.round(hours * 3600 * 1000)).create();
}

/**
 * Whether to stop because of the quota.
 * LAST_QUOTA is set in xeroFetch from the Xero response headers.
 */
function quotaLow_() {
  return typeof LAST_QUOTA !== 'undefined' &&
         LAST_QUOTA.day !== null && LAST_QUOTA.day !== undefined &&
         LAST_QUOTA.day < CONFIG.dailyQuotaFloor;
}

/** Total row count across all data sheets, for the log counter. */
function countSheetTotal_() {
  let total = 0;
  Object.keys(SCHEMA).forEach(function (name) {
    const sh = sheet_(name);
    if (sh) total += Math.max(0, sh.getLastRow() - 1);
  });
  return total;
}

/**
 * Removes duplicates by ID. A duplicate can appear if
 * the export paused in the middle of pagination: the registry lives in
 * execution memory and starts over after a restart.
 */
function dedupeSheets() {
  const targets = [
    { sheet: 'Documents',        idCol: 'DocumentID' },
    { sheet: 'BankTransactions', idCol: 'TransactionID' },
    { sheet: 'Lines',            idCol: null },   // composite key
    { sheet: 'Payments',         idCol: null }
  ];

  targets.forEach(function (t) {
    const sh = sheet_(t.sheet);
    if (!sh || sh.getLastRow() < 3) return;

    const schema = SCHEMA[t.sheet];
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, schema.length).getValues();

    let keyOf;
    if (t.idCol) {
      const i = schema.indexOf(t.idCol);
      keyOf = function (r) { return String(r[i]); };
    } else if (t.sheet === 'Lines') {
      const iDoc = schema.indexOf('DocumentID'), iLine = schema.indexOf('LineItemID');
      const iDesc = schema.indexOf('Description'), iAmt = schema.indexOf('LineAmount');
      keyOf = function (r) {
        return r[iLine] ? String(r[iDoc]) + '|' + String(r[iLine])
                        : String(r[iDoc]) + '|' + String(r[iDesc]) + '|' + String(r[iAmt]);
      };
    } else {
      const iPay = schema.indexOf('PaymentID'), iAcc = schema.indexOf('AccountCode');
      const iAmt = schema.indexOf('AllocatedAmount');
      keyOf = function (r) {
        return String(r[iPay]) + '|' + String(r[iAcc]) + '|' + String(r[iAmt]);
      };
    }

    const seen = {}, kept = [];
    data.forEach(function (r) {
      const key = keyOf(r);
      if (!key || key === 'undefined') { kept.push(r); return; }
      if (seen[key]) return;
      seen[key] = 1;
      kept.push(r);
    });

    const removed = data.length - kept.length;
    if (!removed) return;

    dropFilter_(sh);
    sh.getRange(2, 1, data.length, schema.length).clearContent();
    if (kept.length) {
      ensureRows_(sh, kept.length + 1);
      sh.getRange(2, 1, kept.length, schema.length).setValues(kept);
    }
    Log.info('%s: duplicates removed: %s', t.sheet, removed);
  });
}

function pageSize_() {
  return CONFIG.includeLineItems ? CONFIG.pageSizeWithLines : CONFIG.pageSize;
}

/** Xero returns a pagination object, which we rely on. */
function lastPage_(res, received, state, size) {
  const p = res.pagination;
  if (p && p.pageCount) {
    if (state.page >= p.pageCount || state.page >= CONFIG.maxPages) return true;
    state.page++;
    return false;
  }
  if (received < size || state.page >= CONFIG.maxPages) return true;
  state.page++;
  return false;
}

// ================================================================
//  7. DOCUMENT LINES
// ================================================================

/**
 * Lines of one document with allocated payment.
 *
 * In Xero a payment is linked to the whole document, not to a line.
 * LinePaid / LineCredited / LineDue are CALCULATED: the document amount
 * is allocated in proportion to the line share of the total INCLUDING TAX
 * (the payment includes tax, otherwise lines with different rates
 * would get skewed amounts).
 *
 * The rounding remainder goes to the last line, so
 * the sum of the lines matches the document exactly.
 */
function docLines_(company, docType, doc, accMap) {
  const items = doc.lineItems || [];
  if (!items.length) return [];

  const docTotal    = num_(doc.total);
  const docPaid     = num_(doc.paid);
  const docCredited = num_(doc.credited);
  const docDue      = num_(doc.due);

  const gross = items.map(function (li) {
    return num_(li.LineAmount) + num_(li.TaxAmount);
  });
  const grossSum = gross.reduce(function (a, b) { return a + b; }, 0);

  let paidLeft = docPaid, creditedLeft = docCredited, dueLeft = docDue;

  return items.map(function (li, idx) {
    const isLast = idx === items.length - 1;
    const share = grossSum !== 0 ? gross[idx] / grossSum : 0;

    let linePaid, lineCredited, lineDue;
    if (isLast) {
      linePaid = round2_(paidLeft);
      lineCredited = round2_(creditedLeft);
      lineDue = round2_(dueLeft);
    } else {
      linePaid = round2_(docPaid * share);
      lineCredited = round2_(docCredited * share);
      lineDue = round2_(docDue * share);
      paidLeft -= linePaid;
      creditedLeft -= lineCredited;
      dueLeft -= lineDue;
    }

    const acc = accMap[String(li.AccountCode)] || {};

    return [company, docType, doc.id, doc.number || '', doc.date,
            doc.contact || '', doc.contactId || '', doc.status || '', doc.currency || '',
            li.LineItemID || '', li.Description || '',
            li.AccountCode || '', acc.name || '', acc.class || '', acc.type || '',
            li.ItemCode || '',
            num_(li.Quantity), num_(li.UnitAmount), num_(li.DiscountRate),
            li.TaxType || '', num_(li.TaxAmount),
            num_(li.LineAmount), round2_(gross[idx]), round2_(share * 100),
            linePaid, lineCredited, lineDue,
            docTotal, docPaid, docDue,
            (li.Tracking || []).map(function (t) { return t.Name + ': ' + t.Option; }).join('; ')];
  });
}

// ================================================================
//  8. DOCUMENTS: ONE SHEET
// ================================================================

/**
 * Row for the Documents sheet.
 * DocType: Invoice | Bill | Credit Note (AR) | Credit Note (AP) | Prepayment | Overpayment
 * Side: AR (receivables) | AP (payables)
 */
function docRow_(company, o) {
  return [company, o.docType, o.side, o.id, o.number || '', o.xeroType, o.status,
          o.date, o.dueDate || '', o.contactName || '', o.contactId || '', o.reference || '',
          o.currency || '', num_(o.rate),
          num_(o.subTotal), num_(o.totalTax), num_(o.total),
          num_(o.paid), num_(o.credited), num_(o.due),
          o.fullyPaidOn || '', o.updated || '',
          o.hasAttachments === true, ''];   // links are set by syncAttachments()
}

/** Invoices and bills in one pass: no type filter. */
function exportInvoices_(company, state, seen) {
  const size = pageSize_();
  const accMap = accounts_(company);

  const res = xeroFetch(company, 'Invoices?' + [
    'where=' + encodeURIComponent(dateFilter_('Invoices')),
    'order=Date',
    'page=' + state.page,
    'pageSize=' + size,
    'summaryOnly=' + (CONFIG.includeLineItems ? 'false' : 'true')
  ].join('&'));

  const all = res.Invoices || [];
  const docs = seenFilter_(seen, company + ':Invoices', all, 'InvoiceID');

  const rows = [], aging = [];
  let lines = [];

  docs.forEach(function (d) {
    const isAR = d.Type === 'ACCREC';
    const docType = isAR ? 'Invoice' : 'Bill';
    const date = parseDate_(d.Date), due = parseDate_(d.DueDate);
    const contact = d.Contact || {};

    rows.push(docRow_(company, {
      docType: docType, side: isAR ? 'AR' : 'AP',
      id: d.InvoiceID, number: d.InvoiceNumber, xeroType: d.Type, status: d.Status,
      date: date, dueDate: due,
      contactName: contact.Name, contactId: contact.ContactID, reference: d.Reference,
      // For a document in base currency Xero sends no rate, so it is 1
      currency: d.CurrencyCode, rate: d.CurrencyRate === undefined ? 1 : d.CurrencyRate,
      subTotal: d.SubTotal, totalTax: d.TotalTax, total: d.Total,
      paid: d.AmountPaid, credited: d.AmountCredited, due: d.AmountDue,
      fullyPaidOn: parseDate_(d.FullyPaidOnDate), updated: parseDate_(d.UpdatedDateUTC),
      hasAttachments: d.HasAttachments
    }));

    if (num_(d.AmountDue) > 0 && d.Status === 'AUTHORISED') {
      const days = daysOverdue_(due, state.balanceDate);
      aging.push([company, state.balanceDate, isAR ? 'AR' : 'AP',
                  contact.Name || '', contact.ContactID || '',
                  d.InvoiceNumber || '', date, due, days, bucket_(days),
                  d.CurrencyCode || '', num_(d.AmountDue)]);
    }

    if (CONFIG.includeLineItems) {
      lines = lines.concat(docLines_(company, docType, {
        id: d.InvoiceID, number: d.InvoiceNumber, date: date,
        contact: contact.Name, contactId: contact.ContactID,
        status: d.Status, currency: d.CurrencyCode,
        total: d.Total, paid: d.AmountPaid,
        credited: d.AmountCredited, due: d.AmountDue,
        lineItems: d.LineItems
      }, accMap));
    }
  });

  appendRows_('Documents', rows);
  appendRows_('Aging', aging);
  appendRows_('Lines', lines);

  return lastPage_(res, all.length, state, size);
}

function exportCreditNotes_(company, state, seen) {
  const size = pageSize_();
  const accMap = accounts_(company);

  const res = xeroFetch(company, 'CreditNotes?' + [
    'where=' + encodeURIComponent(dateFilter_('CreditNotes')),
    'order=Date',
    'page=' + state.page,
    'pageSize=' + size
  ].join('&'));

  const all = res.CreditNotes || [];
  const notes = seenFilter_(seen, company + ':CreditNotes', all, 'CreditNoteID');

  const rows = [];
  let lines = [];

  notes.forEach(function (n) {
    const isAR = n.Type === 'ACCRECCREDIT';
    const docType = isAR ? 'Credit Note (AR)' : 'Credit Note (AP)';
    const date = parseDate_(n.Date);
    const contact = n.Contact || {};

    rows.push(docRow_(company, {
      docType: docType, side: isAR ? 'AR' : 'AP',
      id: n.CreditNoteID, number: n.CreditNoteNumber, xeroType: n.Type, status: n.Status,
      date: date, dueDate: parseDate_(n.DueDate),
      contactName: contact.Name, contactId: contact.ContactID, reference: n.Reference,
      currency: n.CurrencyCode, rate: n.CurrencyRate === undefined ? 1 : n.CurrencyRate,
      subTotal: n.SubTotal, totalTax: n.TotalTax,
      // A credit note reduces the balance, so it is written as negative
      total: -num_(n.Total), paid: 0, credited: 0, due: -num_(n.RemainingCredit),
      fullyPaidOn: parseDate_(n.FullyPaidOnDate), updated: parseDate_(n.UpdatedDateUTC),
      hasAttachments: n.HasAttachments
    }));

    if (CONFIG.includeLineItems) {
      lines = lines.concat(docLines_(company, docType, {
        id: n.CreditNoteID, number: n.CreditNoteNumber, date: date,
        contact: contact.Name, contactId: contact.ContactID,
        status: n.Status, currency: n.CurrencyCode,
        total: -num_(n.Total), paid: 0, credited: 0,
        due: -num_(n.RemainingCredit),
        lineItems: n.LineItems
      }, accMap));
    }
  });

  appendRows_('Documents', rows);
  appendRows_('Lines', lines);

  return lastPage_(res, all.length, state, size);
}

/** Prepayments and overpayments also sit in receivables/payables. */
function exportPrepay_(company, state, seen, endpoint) {
  const size = pageSize_();
  const idField = endpoint === 'Prepayments' ? 'PrepaymentID' : 'OverpaymentID';
  const label = endpoint === 'Prepayments' ? 'Prepayment' : 'Overpayment';

  const res = xeroFetch(company, endpoint + '?' + [
    'where=' + encodeURIComponent(dateFilter_(endpoint)),
    'order=Date',
    'page=' + state.page,
    'pageSize=' + size
  ].join('&'));

  const all = (res[endpoint] || []).filter(changedSince_);
  const items = seenFilter_(seen, company + ':' + endpoint, all, idField);

  const rows = items.map(function (p) {
    const isAR = /RECEIVE/.test(p.Type || '');
    const contact = p.Contact || {};
    return docRow_(company, {
      docType: label, side: isAR ? 'AR' : 'AP',
      id: p[idField], number: p.InvoiceNumber || '', xeroType: p.Type, status: p.Status,
      date: parseDate_(p.Date), dueDate: '',
      contactName: contact.Name, contactId: contact.ContactID, reference: p.Reference,
      currency: p.CurrencyCode, rate: p.CurrencyRate,
      subTotal: p.SubTotal, totalTax: p.TotalTax, total: p.Total,
      paid: 0, credited: 0, due: p.RemainingCredit,
      fullyPaidOn: '', updated: parseDate_(p.UpdatedDateUTC)
    });
  });

  appendRows_('Documents', rows);
  return lastPage_(res, all.length, state, size);
}

// ================================================================
//  9. DOCUMENT LINE ITEMS FOR PAYMENT SPLIT
// ================================================================

let CATEGORY_INDEX = null;

/**
 * Reset the index. Required before each payments job:
 * between jobs, rows of the next entity are appended to Lines,
 * and the old index does not contain them, so payments would fall back.
 */
function resetCategoryIndex_() { CATEGORY_INDEX = null; }

/**
 * DocumentID -> document line items, from the already exported Lines sheet.
 * Makes no API calls.
 *
 * @return {Object} { docId: {accounts: [{code,name,class,type,gross}], grossSum} }
 */
function categoryIndex_() {
  if (CATEGORY_INDEX) return CATEGORY_INDEX;

  CATEGORY_INDEX = {};
  const sh = sheet_('Lines');
  if (!sh || sh.getLastRow() < 2) {
    Log.warn('categoryIndex_: the Lines sheet is empty, the payment split will not work.');
    return CATEGORY_INDEX;
  }

  const s = SCHEMA.Lines;
  const i = {
    docId:   s.indexOf('DocumentID'),
    type:    s.indexOf('DocType'),
    code:    s.indexOf('AccountCode'),
    name:    s.indexOf('AccountName'),
    cls:     s.indexOf('AccountClass'),
    accType: s.indexOf('AccountType'),
    gross:   s.indexOf('LineGross')
  };

  sh.getRange(2, 1, sh.getLastRow() - 1, s.length).getValues().forEach(function (r) {
    const id = r[i.docId];
    if (!id || r[i.type] === 'BankTransaction') return;

    if (!CATEGORY_INDEX[id]) CATEGORY_INDEX[id] = { byCode: {}, grossSum: 0 };
    const entry = CATEGORY_INDEX[id];
    const code = String(r[i.code] || '(no account)');
    const gross = num_(r[i.gross]);

    if (!entry.byCode[code]) {
      entry.byCode[code] = { code: code, name: r[i.name] || '', class: r[i.cls] || '',
                             type: r[i.accType] || '', gross: 0 };
    }
    entry.byCode[code].gross += gross;
    entry.grossSum += gross;
  });

  // Expand into arrays, sorted by amount descending
  Object.keys(CATEGORY_INDEX).forEach(function (id) {
    const e = CATEGORY_INDEX[id];
    e.accounts = Object.keys(e.byCode).map(function (c) { return e.byCode[c]; })
      .sort(function (a, b) { return b.gross - a.gross; });
    delete e.byCode;
  });

  Log.info('categoryIndex_: documents with line items: %s', Object.keys(CATEGORY_INDEX).length);
  return CATEGORY_INDEX;
}

/** Budget of extra document requests per run, per entity. */
const CATEGORY_BUDGET = {};
const CATEGORY_FETCHED = {};

/**
 * Line items of a document missing from Lines, e.g. an invoice
 * issued before fromDate but paid in the exported period.
 */
function fetchDocCategories_(company, invoiceId) {
  if (!CONFIG.resolvePaymentCategories || !invoiceId) return null;
  if (CATEGORY_FETCHED[invoiceId] !== undefined) return CATEGORY_FETCHED[invoiceId];

  if (CATEGORY_BUDGET[company] === undefined) CATEGORY_BUDGET[company] = CONFIG.resolveCategoriesLimit;
  if (CATEGORY_BUDGET[company] <= 0) return null;
  CATEGORY_BUDGET[company]--;

  let doc;
  try {
    const res = xeroFetch(company, 'Invoices/' + invoiceId);
    doc = (res.Invoices || [])[0];
  } catch (err) {
    Log.error('Could not fetch document %s: %s', invoiceId, err.message);
    CATEGORY_FETCHED[invoiceId] = null;
    return null;
  }
  if (!doc || !doc.LineItems || !doc.LineItems.length) {
    CATEGORY_FETCHED[invoiceId] = null;
    return null;
  }

  const accMap = accounts_(company);
  const byCode = {};
  let grossSum = 0;

  doc.LineItems.forEach(function (li) {
    const code = String(li.AccountCode || '(no account)');
    const gross = num_(li.LineAmount) + num_(li.TaxAmount);
    const acc = accMap[code] || {};
    if (!byCode[code]) {
      byCode[code] = { code: code, name: acc.name || '', class: acc.class || '',
                       type: acc.type || '', gross: 0 };
    }
    byCode[code].gross += gross;
    grossSum += gross;
  });

  const result = {
    accounts: Object.keys(byCode).map(function (c) { return byCode[c]; })
      .sort(function (a, b) { return b.gross - a.gross; }),
    grossSum: grossSum
  };

  CATEGORY_FETCHED[invoiceId] = result;
  return result;
}

/**
 * Allocates the payment amount across the document line items by their share.
 * The rounding remainder goes to the last item, so the sum
 * of the allocated parts equals the payment amount exactly.
 */
function allocatePayment_(amount, categories) {
  if (!categories || !categories.accounts.length || categories.grossSum === 0) return null;

  const accounts = categories.accounts;
  let left = amount;

  return accounts.map(function (a, idx) {
    const isLast = idx === accounts.length - 1;
    const share = a.gross / categories.grossSum;
    const part = isLast ? round2_(left) : round2_(amount * share);
    if (!isLast) left -= part;
    return { code: a.code, name: a.name, class: a.class, type: a.type,
             share: round2_(share * 100), amount: part };
  });
}

// ================================================================
//  10. BANK
// ================================================================

/** SPEND* = money leaves the account, RECEIVE* = money comes in. */
function direction_(type) {
  if (/^SPEND/.test(type))   return 'Out';
  if (/^RECEIVE/.test(type)) return 'In';
  return '';
}

/**
 * Direction of a document payment.
 *   ACCRECPAYMENT   customer pays an invoice      -> inflow
 *   ACCPAYPAYMENT   bill paid to a supplier       -> outflow
 *   ARCREDITPAYMENT refund on an AR credit note   -> outflow
 *   APCREDITPAYMENT refund on an AP credit note   -> inflow
 */
function paymentDirection_(type) {
  switch (type) {
    case 'ACCRECPAYMENT':   return 'In';
    case 'ACCPAYPAYMENT':   return 'Out';
    case 'ARCREDITPAYMENT': return 'Out';
    case 'APCREDITPAYMENT': return 'In';
    default:
      if (/^AR/.test(type)) return 'In';
      if (/^AP/.test(type)) return 'Out';
      return '';
  }
}

/** Contra accounts from transaction lines: where exactly the money landed. */
function contra_(lineItems, accMap) {
  const codes = [], names = [];
  (lineItems || []).forEach(function (li) {
    if (!li.AccountCode || codes.indexOf(li.AccountCode) >= 0) return;
    codes.push(li.AccountCode);
    names.push((accMap[li.AccountCode] || {}).name || '');
  });
  return { codes: codes.join(' | '), names: names.join(' | ') };
}

/** Budget of history requests per run, per entity. */
const HISTORY_BUDGET = {};

/**
 * Who reconciled and when. The transaction itself does not contain this,
 * only /History does, with User, Changes, Details, DateUTC fields.
 *
 * Limitations:
 *   - actions via the API are logged as System Generated, not as a person
 *   - reconciliation by a bank rule has no user
 *   - the Changes wording depends on the locale, so several variants are matched
 */
function reconciledBy_(company, bankTransactionId) {
  const empty = { user: '', at: '' };
  if (!bankTransactionId) return empty;

  if (HISTORY_BUDGET[company] === undefined) HISTORY_BUDGET[company] = CONFIG.reconciledByLimit;
  if (HISTORY_BUDGET[company] <= 0) return { user: 'request limit exhausted', at: '' };
  HISTORY_BUDGET[company]--;

  let records;
  try {
    const res = xeroFetch(company, 'BankTransactions/' + bankTransactionId + '/History');
    records = res.HistoryRecords || [];
  } catch (err) {
    Log.error('reconciledBy_ failed', err);
    return { user: 'error: ' + err.message, at: '' };
  }

  const rx = /reconcil/i;
  let best = null;

  records.forEach(function (h) {
    const text = String(h.Changes || '') + ' ' + String(h.Details || '');
    if (!rx.test(text)) return;
    const when = parseDate_(h.DateUTC);
    if (!best || (when && best.when && when > best.when)) {
      best = { user: h.User || '', when: when };
    }
  });

  return best ? { user: best.user, at: best.when || '' } : empty;
}

function exportBankTx_(company, state, seen) {
  const size = pageSize_();
  const accMap = accounts_(company);

  const res = xeroFetch(company, 'BankTransactions?' + [
    'where=' + encodeURIComponent(dateFilter_('BankTransactions')),
    'order=Date',
    'page=' + state.page,
    'pageSize=' + size
  ].join('&'));

  const all = res.BankTransactions || [];
  const txs = seenFilter_(seen, company + ':BankTransactions', all, 'BankTransactionID');

  const rows = [];
  let lines = [];
  const recFrom = CONFIG.reconciledByFrom ? new Date(CONFIG.reconciledByFrom + 'T00:00:00Z') : null;

  txs.forEach(function (t) {
    if (CONFIG.skipTransferTypes && /TRANSFER/.test(t.Type)) return;

    const bank = t.BankAccount || {};
    const bankName = bank.Name || (accMap[bank.AccountID] || {}).name || '';
    const bankCode = bank.Code || (accMap[bank.AccountID] || {}).code || '';
    const contact = t.Contact || {};
    const dir = direction_(t.Type);
    const date = parseDate_(t.Date);

    let lineItems = t.LineItems;
    if ((!lineItems || !lineItems.length) && CONFIG.resolveContraAccounts) {
      const single = xeroFetch(company, 'BankTransactions/' + t.BankTransactionID);
      lineItems = ((single.BankTransactions || [])[0] || {}).LineItems || [];
    }
    const c = contra_(lineItems, accMap);

    const wantHistory = CONFIG.fetchReconciledBy &&
                        t.IsReconciled === true &&
                        (!recFrom || !(date instanceof Date) || date >= recFrom);
    const rec = wantHistory ? reconciledBy_(company, t.BankTransactionID) : { user: '', at: '' };

    rows.push([company, 'BankTransaction', t.BankTransactionID, t.Type, t.Status,
               date, dir, bankName, bankCode,
               dir === 'Out' ? bankName : c.names,
               dir === 'Out' ? c.names : bankName,
               c.codes, c.names,
               contact.Name || '', t.Reference || '', t.IsReconciled === true,
               t.CurrencyCode || '', num_(t.CurrencyRate),
               num_(t.SubTotal), num_(t.TotalTax), num_(t.Total),
               parseDate_(t.UpdatedDateUTC),
               rec.user, rec.at]);

    if (CONFIG.includeLineItems) {
      // A bank transaction is itself a payment: paid = total, due = 0
      lines = lines.concat(docLines_(company, 'BankTransaction', {
        id: t.BankTransactionID, number: '', date: date,
        contact: contact.Name, contactId: contact.ContactID,
        status: t.Status, currency: t.CurrencyCode,
        total: t.Total, paid: t.Total, credited: 0, due: 0,
        lineItems: lineItems
      }, accMap));
    }
  });

  appendRows_('BankTransactions', rows);
  appendRows_('Lines', lines);

  return lastPage_(res, all.length, state, size);
}

/**
 * Journal entries on both sides of a transfer. For a multi-currency transfer, Amount
 * in /BankTransfers is in the source account currency, which is wrong
 * for the incoming leg. The entries themselves carry the correct Total and CurrencyCode.
 */
function transferLegs_(company, transfer) {
  if (!CONFIG.resolveTransferAmounts) return { out: {}, in: {} };

  function fetchTx(id) {
    if (!id) return {};
    try {
      const res = xeroFetch(company, 'BankTransactions/' + id);
      return (res.BankTransactions || [])[0] || {};
    } catch (err) {
      Log.error('Could not fetch transaction %s: %s', id, err.message);
      return {};
    }
  }

  return { out: fetchTx(transfer.FromBankTransactionID),
           in:  fetchTx(transfer.ToBankTransactionID) };
}

/**
 * Transfers between own accounts: TWO rows per transfer, Out and In.
 * BankTransfers returns no pagination object and sends everything at once,
 * so one request without page (page-by-page would loop forever).
 */
function exportTransfers_(company, seen) {
  const accMap = accounts_(company);
  const res = xeroFetch(company, 'BankTransfers?where=' +
                        encodeURIComponent(dateFilter_('BankTransfers')));

  // BankTransfer cannot be filtered by modification date, so we filter here
  const all = (res.BankTransfers || []).filter(changedSince_);
  const transfers = seenFilter_(seen, company + ':BankTransfers', all, 'BankTransferID');

  const rows = [];

  transfers.forEach(function (t) {
    const from = t.FromBankAccount || {}, to = t.ToBankAccount || {};

    const fromName = from.Name || (accMap[from.AccountID] || {}).name || '';
    const toName   = to.Name   || (accMap[to.AccountID]   || {}).name || '';
    const fromCode = from.Code || (accMap[from.AccountID] || {}).code || '';
    const toCode   = to.Code   || (accMap[to.AccountID]   || {}).code || '';

    const date = parseDate_(t.Date);
    const updated = parseDate_(t.UpdatedDateUTC);
    const legs = transferLegs_(company, t);

    function legRow(dir, leg, txId, recFlag) {
      const isOut = dir === 'Out';
      const amount = leg.Total !== undefined ? num_(leg.Total) : num_(t.Amount);
      const currency = leg.CurrencyCode || '';
      const rate = leg.CurrencyRate !== undefined ? num_(leg.CurrencyRate) : num_(t.CurrencyRate);
      const reconciled = leg.IsReconciled !== undefined ? leg.IsReconciled === true : recFlag === true;

      const rec = (CONFIG.fetchReconciledBy && reconciled && txId)
        ? reconciledBy_(company, txId)
        : { user: '', at: '' };

      return [company, 'BankTransfer',
              t.BankTransferID + (isOut ? ':out' : ':in'),
              isOut ? 'TRANSFER-OUT' : 'TRANSFER-IN',
              leg.Status || '',
              date, dir,
              isOut ? fromName : toName,
              isOut ? fromCode : toCode,
              fromName, toName,                    // From/To are the same in both rows
              isOut ? toCode : fromCode,           // contra account of this leg
              isOut ? toName : fromName,
              '',                                  // Contact
              t.Reference || '',
              reconciled,
              currency, rate,
              amount, 0, amount,
              updated,
              rec.user, rec.at];
    }

    rows.push(legRow('Out', legs.out, t.FromBankTransactionID, t.FromIsReconciled));
    rows.push(legRow('In',  legs.in,  t.ToBankTransactionID,   t.ToIsReconciled));
  });

  appendRows_('BankTransactions', rows);
  return true;
}

/**
 * Document payments split by line item.
 *
 * A payment of 500 against an invoice with lines Ads 300 / IT 200 becomes two
 * rows: 300 to Ads, 200 to IT. This way money lands in the same P&L sections
 * as the document itself, not in receivables.
 *
 * Line items come from the already exported Lines sheet; if the document is
 * missing there (issued before fromDate), it is fetched, within a limit.
 * When no line items can be found at all, the payment is written as one row
 * with the contra account Accounts Receivable / Payable.
 */
function exportPayments_(company, state, seen) {
  const size = CONFIG.pageSize;
  const accMap = accounts_(company);

  // Before the first page of the job the index is rebuilt: Lines already
  // contains rows of this entity that were missing in the previous job
  if (state.page === 1) resetCategoryIndex_();
  const index = CONFIG.splitPaymentsByCategory ? categoryIndex_() : {};

  const res = xeroFetch(company, 'Payments?' + [
    'where=' + encodeURIComponent(dateFilter_('Payments')),
    'order=Date',
    'page=' + state.page,
    'pageSize=' + size
  ].join('&'));

  const all = res.Payments || [];
  const payments = seenFilter_(seen, company + ':Payments', all, 'PaymentID');

  const payRows = [], bankRows = [];

  payments.forEach(function (p) {
    const inv = p.Invoice || {};
    const acc = p.Account || {};
    const contact = inv.Contact || {};
    const date = parseDate_(p.Date);
    const amount = num_(p.Amount);
    const updated = parseDate_(p.UpdatedDateUTC);
    const rate = p.CurrencyRate !== undefined ? num_(p.CurrencyRate) : '';

    const bankName = acc.Name || (accMap[acc.AccountID] || {}).name || '';
    const bankCode = acc.Code || (accMap[acc.AccountID] || {}).code || '';
    const dir = paymentDirection_(p.PaymentType);

    // Document line items: first from the Lines sheet, then an extra request
    let categories = null;
    if (CONFIG.splitPaymentsByCategory && inv.InvoiceID) {
      categories = index[inv.InvoiceID] || fetchDocCategories_(company, inv.InvoiceID);
    }
    const parts = categories ? allocatePayment_(amount, categories) : null;

    const fallbackLabel = (inv.Type === 'ACCPAY' ? 'Accounts Payable' : 'Accounts Receivable') +
                          (inv.InvoiceNumber ? ' / ' + inv.InvoiceNumber : '');

    if (parts) {
      parts.forEach(function (part) {
        payRows.push([company, p.PaymentID, p.PaymentType || '', p.Status || '',
                      date,
                      inv.InvoiceID || '', inv.InvoiceNumber || '', inv.Type || '',
                      contact.Name || '', bankName, bankCode,
                      p.Reference || '', p.IsReconciled === true, rate,
                      amount,
                      part.code, part.name, part.class, part.type,
                      part.share, part.amount,
                      updated]);
      });
    } else {
      payRows.push([company, p.PaymentID, p.PaymentType || '', p.Status || '',
                    date,
                    inv.InvoiceID || '', inv.InvoiceNumber || '', inv.Type || '',
                    contact.Name || '', bankName, bankCode,
                    p.Reference || '', p.IsReconciled === true, rate,
                    amount,
                    '', fallbackLabel, '', '', 100, amount,
                    updated]);
    }

    if (!CONFIG.paymentsIntoBankSheet) return;
    if (p.Status === 'DELETED') return;          // deleted payments do not move money

    if (parts) {
      // One row per line item: the ID suffix tells them apart
      parts.forEach(function (part, idx) {
        bankRows.push([company, 'Payment',
                       p.PaymentID + ':' + (idx + 1),
                       p.PaymentType || '', p.Status || '',
                       date, dir, bankName, bankCode,
                       dir === 'Out' ? bankName : part.name,
                       dir === 'Out' ? part.name : bankName,
                       part.code, part.name,
                       contact.Name || '',
                       p.Reference || (inv.InvoiceNumber ? 'Inv. ' + inv.InvoiceNumber : ''),
                       p.IsReconciled === true,
                       '', rate,
                       part.amount, 0, part.amount,
                       updated, '', '']);
      });
    } else {
      bankRows.push([company, 'Payment', p.PaymentID, p.PaymentType || '', p.Status || '',
                     date, dir, bankName, bankCode,
                     dir === 'Out' ? bankName : fallbackLabel,
                     dir === 'Out' ? fallbackLabel : bankName,
                     inv.InvoiceID || '', fallbackLabel,
                     contact.Name || '', p.Reference || '', p.IsReconciled === true,
                     '', rate,
                     amount, 0, amount,
                     updated, '', '']);
    }
  });

  appendRows_('Payments', payRows);
  if (bankRows.length) appendRows_('BankTransactions', bankRows);

  return lastPage_(res, all.length, state, size);
}

// ================================================================
//  10a. GENERAL LEDGER
// ================================================================

/**
 * Journal types already covered by the Documents, Lines,
 * BankTransactions and Payments sheets. Skipped when journalsOnlyMissing is set.
 */
const COVERED_SOURCE_TYPES = {
  ACCREC: 1, ACCPAY: 1, ACCRECCREDIT: 1, ACCPAYCREDIT: 1,
  ACCRECPAYMENT: 1, ACCPAYPAYMENT: 1, ARCREDITPAYMENT: 1, APCREDITPAYMENT: 1,
  CASHREC: 1, CASHPAID: 1, TRANSFER: 1,
  ARPREPAYMENT: 1, APPREPAYMENT: 1, AROVERPAYMENT: 1, APOVERPAYMENT: 1
};

/**
 * System accounts whose lines are needed EVEN inside covered journals.
 * A realised FX difference has no journal type of its own: Xero
 * adds it as a separate line to the payment journal. If covered types
 * were dropped entirely, forex would be lost with them.
 */
const KEEP_SYSTEM_ACCOUNTS = {
  REALISEDCURRENCYGAIN: 'FX difference (realised)',
  UNREALISEDCURRENCYGAIN: 'FX difference (unrealised)',
  BANKCURRENCYGAIN: 'FX difference by account',
  ROUNDING: 'Rounding'
};

/** A clear reason why the row is in the sheet. */
function journalReason_(sourceType, systemAccount) {
  if (KEEP_SYSTEM_ACCOUNTS[systemAccount]) return KEEP_SYSTEM_ACCOUNTS[systemAccount];
  if (sourceType === 'MANJOURNAL') return 'Manual journal';
  if (/PAYROLL|PAYSLIP|WAGE/.test(sourceType)) return 'Payroll';
  if (/EXPCLAIM|EXPPAYMENT/.test(sourceType)) return 'Expense claim';
  return sourceType || 'Other';
}

/**
 * Journal entries that are missing from other sheets.
 *
 * What ends up here: FX differences, manual journals, depreciation,
 * rounding, payroll. They have no LineItems, so they never appear
 * in the Lines sheet. This is how Xero works, not a gap in the export.
 *
 * Pagination is special: not page but offset by journal number, 100 records
 * per request. The endpoint has no date filter, so records
 * before fromDate are dropped after retrieval.
 */
function exportJournals_(company, state, seen) {
  const offset = state.offset || CONFIG.journalsStartOffset || 0;
  const res = xeroFetch(company, 'Journals?offset=' + offset);

  const all = res.Journals || [];
  if (!all.length) return true;

  const journals = seenFilter_(seen, company + ':Journals', all, 'JournalID');
  const accMap = accounts_(company);
  const from = new Date(CONFIG.fromDate + 'T00:00:00Z');
  const rows = [];
  let maxNumber = offset;

  journals.forEach(function (j) {
    const number = num_(j.JournalNumber);
    if (number > maxNumber) maxNumber = number;

    const date = parseDate_(j.JournalDate);
    if (date instanceof Date && date < from) return;      // outside the period

    const sourceType = j.SourceType || '';
    const covered = CONFIG.journalsOnlyMissing && COVERED_SOURCE_TYPES[sourceType];

    (j.JournalLines || []).forEach(function (l) {
      const acc = accMap[String(l.AccountCode)] || {};
      const system = acc.system || '';

      // From covered journals only FX differences and rounding are taken
      if (covered && !KEEP_SYSTEM_ACCOUNTS[system]) return;

      rows.push([company, j.JournalID, date, number,
                 sourceType, j.SourceID || '', j.Reference || '',
                 journalReason_(sourceType, system),
                 l.JournalLineID || '',
                 l.AccountCode || '', l.AccountName || '', l.AccountType || '',
                 acc.class || '', system,
                 l.Description || '',
                 num_(l.NetAmount), num_(l.GrossAmount), num_(l.TaxAmount),
                 l.TaxType || '',
                 (l.TrackingCategories || []).map(function (t) {
                   return t.Name + ': ' + t.Option;
                 }).join('; '),
                 parseDate_(j.CreatedDateUTC)]);
    });
  });

  appendRows_('Journals', rows);

  // Journals come 100 at a time; the next call starts from the highest number
  all.forEach(function (j) {
    const n = num_(j.JournalNumber);
    if (n > maxNumber) maxNumber = n;
  });
  state.offset = maxNumber;

  return all.length < 100;
}

// ================================================================
//  10b. P&L FROM THE XERO REPORT
// ================================================================

/**
 * Monthly P&L from Reports/ProfitAndLoss is the source of truth for accounts
 * that are not in documents: FX differences, depreciation, rounding,
 * payroll. The /Journals endpoint, which has these entries line by line,
 * is only available on the Advanced plan since 2026.
 *
 * The report is aggregated: total per account per month, no transactions.
 * Main limitation: a line cannot be matched to a specific
 * payment from here, only the total can be seen.
 *
 * Sign: Xero returns revenue and expenses as positive in their sections,
 * so no extra sign handling is needed here.
 */
function exportPnlReport_(company, toDate) {
  const res = xeroFetch(company, 'Reports/ProfitAndLoss?' + [
    'toDate=' + reportEndDate_(),
    'periods=' + CONFIG.balanceSheetPeriods,
    'timeframe=' + CONFIG.balanceSheetTimeframe,
    'standardLayout=' + (CONFIG.reportsStandardLayout ? 'true' : 'false')
  ].join('&'));

  const parsed = flattenReportPeriods_(res);
  if (!parsed.rows.length) {
    Log.warn('%s: P&L is empty, check scope accounting.reports.profitandloss.read.', company);
    return true;
  }

  const rows = [];
  parsed.rows.forEach(function (r) {
    r.values.forEach(function (v, idx) {
      if (!v) return;                       // empty cells do not clutter the sheet
      rows.push([company, parsed.periods[idx] || ('period ' + (idx + 1)),
                 r.summary ? 'TOTAL ' + r.section : r.section,
                 r.accountId || '', r.name, round2_(v)]);
    });
  });

  appendRows_('PnL_Report', rows);

  const withId = parsed.rows.filter(function (r) { return r.accountId; }).length;
  Log.info('%s / P&L: report rows %s (with AccountID %s), periods %s%s',
    company, parsed.rows.length, withId, parsed.periods.length,
    withId ? '' : '  <- NO AccountID: adjustments will not match');

  return true;
}

/**
 * What from the Xero report did not reach the documents.
 *
 * Compares totals per account and month from PnL_Report with what
 * was collected from the Lines sheet. The difference is what is missing:
 * FX differences, depreciation, rounding. One row per account and month.
 *
 * Built after the export, makes no API calls.
 */
function buildPnlAdjustments() {
  const repSh = sheet_('PnL_Report');
  if (!repSh || repSh.getLastRow() < 2) {
    Log.warn('The PnL_Report sheet is empty.');
    return;
  }

  // 1. What is already in documents: account and month -> amount
  const fromDocs = {};
  const pnlSh = sheet_('PnL');
  if (pnlSh && pnlSh.getLastRow() > 1) {
    const headers = SCHEMA.Lines.concat(PNL_SHEET_EXTRA);
    const i = {
      company: headers.indexOf('Company'), date: headers.indexOf('Date'),
      name: headers.indexOf('AccountName'), amount: headers.indexOf('PnL Amount')
    };
    pnlSh.getRange(2, 1, pnlSh.getLastRow() - 1, headers.length).getValues().forEach(function (r) {
      if (!r[i.company] || !(r[i.date] instanceof Date)) return;
      const key = [r[i.company], Utilities.formatDate(r[i.date], 'UTC', 'MMM yyyy'),
                   String(r[i.name]).trim()].join('||');
      fromDocs[key] = (fromDocs[key] || 0) + num_(r[i.amount]);
    });
  }

  // 2. Compare with the report
  const s = SCHEMA.PnL_Report;
  const i = {
    company: s.indexOf('Company'), period: s.indexOf('PeriodDate'),
    section: s.indexOf('Section'), account: s.indexOf('Account'),
    amount: s.indexOf('Amount')
  };

  const rows = [];
  repSh.getRange(2, 1, repSh.getLastRow() - 1, s.length).getValues().forEach(function (r) {
    if (!r[i.company] || /^TOTAL/.test(String(r[i.section]))) return;

    const account = String(r[i.account]).trim();
    const period = String(r[i.period]).trim();
    const inXero = num_(r[i.amount]);
    const key = [r[i.company], period, account].join('||');
    const inDocs = round2_(fromDocs[key] || 0);
    const diff = round2_(inXero - inDocs);

    if (Math.abs(diff) < 0.01) return;      // matches, not shown

    rows.push([r[i.company], period, r[i.section], account,
               inXero, inDocs, diff,
               inDocs === 0 ? 'not in documents' : 'partly in documents']);
  });

  writeReport_('PnL_Adjustments',
    ['Company', 'Period', 'PnL Section', 'Account',
     'In Xero report', 'From documents', 'Missing', 'Comment'], rows, 4);

  Log.info('PnL_Adjustments: rows %s', rows.length);
}

/**
 * What exactly the Xero report returned. Writes nothing to sheets,
 * only logs the structure: periods, sections, presence of AccountID.
 *
 *   diagnoseReport('entity_cy', 'ProfitAndLoss')
 *   diagnoseReport('entity_cy', 'BalanceSheet')
 */
function diagnoseReport(company, reportName) {
  const name = reportName || 'ProfitAndLoss';
  const dateParam = name === 'BalanceSheet' ? 'date=' : 'toDate=';

  [false, true].forEach(function (standard) {
    Log.info('=== %s, standardLayout=%s ===', name, standard);

    let res;
    try {
      res = xeroFetch(company, 'Reports/' + name + '?' + [
        dateParam + balanceDate_(),
        'periods=' + CONFIG.balanceSheetPeriods,
        'timeframe=' + CONFIG.balanceSheetTimeframe,
        'standardLayout=' + standard
      ].join('&'));
    } catch (err) {
      Log.error('request failed: %s', err.message);
      return;
    }

    const parsed = flattenReportPeriods_(res);
    const withId = parsed.rows.filter(function (r) { return r.accountId; }).length;

    Log.info('periods (%s): %s', parsed.periods.length, parsed.periods.join(' | '));
    Log.info('rows %s, of which with AccountID %s', parsed.rows.length, withId);

    const sections = [];
    parsed.rows.forEach(function (r) {
      if (sections.indexOf(r.section) < 0) sections.push(r.section);
    });
    Log.info('sections: %s', sections.join(' | '));

    parsed.rows.slice(0, 8).forEach(function (r) {
      Log.info('   [%s] %s | AccountID: %s | %s',
        r.section, r.name, r.accountId || 'NO', r.values.slice(0, 3).join(', '));
    });
  });
}

// ================================================================
//  10c. BALANCE SHEET
// ================================================================

/**
 * Parses a Xero report with SEVERAL periods in columns.
 * The regular flattenReport_ does not fit: there columns become
 * object keys, while here their number depends on the settings.
 *
 * @return {{periods: Array<string>, rows: Array<{section,accountId,name,values}>}}
 */
function flattenReportPeriods_(payload) {
  const report = (payload.Reports || [])[0];
  if (!report) return { periods: [], rows: [] };

  let periods = [];
  const rows = [];

  (report.Rows || []).forEach(function (row) {
    if (row.RowType === 'Header') {
      // The first cell is the row label, then the periods
      periods = (row.Cells || []).slice(1).map(function (c) {
        return String(c.Value || '').trim();
      });
      return;
    }

    const section = row.Title || '';
    (row.Rows || []).forEach(function (r) {
      if (r.RowType !== 'Row' && r.RowType !== 'SummaryRow') return;
      const cells = r.Cells || [];
      if (!cells.length) return;

      let accountId = '';
      (cells[0].Attributes || []).forEach(function (a) {
        if (a.Id === 'account' || a.Id === 'accountID') accountId = a.Value;
      });

      rows.push({
        section: section,
        accountId: accountId,
        name: String(cells[0].Value || '').trim(),
        summary: r.RowType === 'SummaryRow',
        values: cells.slice(1).map(function (c) { return num_(c.Value); })
      });
    });
  });

  return { periods: periods, rows: rows };
}

/**
 * Balance sheet by period. One request returns several columns,
 * via the periods and timeframe parameters. Section total rows are marked
 * in the Section column with the prefix TOTAL so they can be filtered out.
 */
function exportBalanceSheet_(company, date) {
  const res = xeroFetch(company, 'Reports/BalanceSheet?' + [
    'date=' + reportEndDate_(),
    'periods=' + CONFIG.balanceSheetPeriods,
    'timeframe=' + CONFIG.balanceSheetTimeframe,
    'standardLayout=' + (CONFIG.reportsStandardLayout ? 'true' : 'false')
  ].join('&'));

  const parsed = flattenReportPeriods_(res);
  if (!parsed.rows.length) {
    Log.warn('%s: balance sheet is empty, check scope accounting.reports.balancesheet.read.', company);
    return true;
  }

  const rows = [];
  parsed.rows.forEach(function (r) {
    r.values.forEach(function (v, idx) {
      const period = parsed.periods[idx] || ('period ' + (idx + 1));
      rows.push([company, period,
                 r.summary ? 'TOTAL ' + r.section : r.section,
                 r.accountId || '', r.name, round2_(v)]);
    });
  });

  appendRows_('Balance_Sheet', rows);
  return true;
}

// ================================================================
//  11. BALANCES
// ================================================================

function exportBalances_(company, date) {
  const tb = xeroFetch(company, 'Reports/TrialBalance?date=' + date + '&paymentsOnly=false');
  appendRows_('Balances', flattenReport_(tb).map(function (r) {
    const debit = num_(r['Debit']), credit = num_(r['Credit']);
    const yDebit = num_(r['YTD Debit']), yCredit = num_(r['YTD Credit']);
    return [company, date, r.section || '', r.accountId || '', r.accountName || '',
            debit, credit, debit - credit, yDebit, yCredit, yDebit - yCredit];
  }));

  const bs = xeroFetch(company, 'Reports/BankSummary?fromDate=' + CONFIG.fromDate + '&toDate=' + date);
  appendRows_('Bank_Balances', flattenReport_(bs).map(function (r) {
    return [company, CONFIG.fromDate, date, r.accountId || '', r.accountName || '',
            num_(r['Opening Balance']), num_(r['Cash Received']),
            num_(r['Cash Spent']), num_(r['Closing Balance'])];
  }));

  return true;
}

// ================================================================
//  12. CHECKS
// ================================================================

/**
 * Asks Xero how many records it sees per endpoint
 * (pageSize=1 -> pagination.itemCount) and compares with the rows in the sheet.
 * A non-zero Diff means pagination lost something.
 */
function verifyExport() {
  const checks = [
    { endpoint: 'Invoices',         skip: false },
    { endpoint: 'CreditNotes',      skip: !CONFIG.includeCreditNotes },
    { endpoint: 'BankTransactions', skip: false },
    { endpoint: 'Payments',         skip: !CONFIG.includePayments }
  ];

  const counted = countSheetRows_();
  const rows = [];

  Object.keys(XERO_COMPANIES).forEach(function (company) {
    checks.forEach(function (c) {
      if (c.skip) return;

      let expected = '';
      try {
        const res = xeroFetch(company, c.endpoint + '?pageSize=1&page=1&where=' +
                              encodeURIComponent('Date>=' + xeroDate_(CONFIG.fromDate)));
        expected = (res.pagination && res.pagination.itemCount !== undefined)
          ? res.pagination.itemCount : '';
      } catch (err) {
    Log.error('verifyExport failed', err);
        expected = 'error: ' + err.message;
      }

      const actual = (counted[company] && counted[company][c.endpoint]) || 0;
      const diff = (expected === '' || isNaN(Number(expected))) ? '' : Number(expected) - actual;

      rows.push([company, c.endpoint, expected, actual, diff,
                 diff === 0 ? 'ok' : (diff === '' ? 'counter not available' : 'CHECK')]);
    });
  });

  writeReport_('_Export_Log', ['Company', 'Endpoint', 'In Xero', 'In sheet', 'Diff', 'Status'], rows);
  Log.info('The check was written to the _Export_Log sheet.');
}

/**
 * Count of exported records per entity and type.
 * Payments are counted by unique PaymentID, otherwise the split
 * by line item would look like extra records.
 */
function countSheetRows_() {
  const out = {};

  function bump(company, key) {
    if (!out[company]) out[company] = {};
    out[company][key] = (out[company][key] || 0) + 1;
  }

  const docs = sheet_('Documents');
  if (docs && docs.getLastRow() > 1) {
    const s = SCHEMA.Documents;
    const iC = s.indexOf('Company'), iT = s.indexOf('DocType');
    docs.getRange(2, 1, docs.getLastRow() - 1, s.length).getValues().forEach(function (r) {
      if (!r[iC]) return;
      if (r[iT] === 'Invoice' || r[iT] === 'Bill') bump(r[iC], 'Invoices');
      else if (/Credit Note/.test(r[iT]))          bump(r[iC], 'CreditNotes');
    });
  }

  const bank = sheet_('BankTransactions');
  if (bank && bank.getLastRow() > 1) {
    const s = SCHEMA.BankTransactions;
    const iC = s.indexOf('Company'), iS = s.indexOf('Source');
    bank.getRange(2, 1, bank.getLastRow() - 1, s.length).getValues().forEach(function (r) {
      if (r[iC] && r[iS] === 'BankTransaction') bump(r[iC], 'BankTransactions');
    });
  }

  const pay = sheet_('Payments');
  if (pay && pay.getLastRow() > 1) {
    const s = SCHEMA.Payments;
    const iC = s.indexOf('Company'), iP = s.indexOf('PaymentID');
    const seenPay = {};
    pay.getRange(2, 1, pay.getLastRow() - 1, s.length).getValues().forEach(function (r) {
      if (!r[iC] || !r[iP]) return;
      const key = r[iC] + '||' + r[iP];
      if (seenPay[key]) return;
      seenPay[key] = 1;
      bump(r[iC], 'Payments');
    });
  }

  return out;
}

/**
 * Compares the sum of LinePaid across lines with the document AmountPaid.
 * A difference means an allocation error and should never happen.
 */
function verifyLineAllocation() {
  const linesSh = sheet_('Lines'), docsSh = sheet_('Documents');
  if (!linesSh || !docsSh || linesSh.getLastRow() < 2 || docsSh.getLastRow() < 2) {
    Log.warn('No data to check.');
    return;
  }

  const ls = SCHEMA.Lines, ds = SCHEMA.Documents;
  const li = { id: ls.indexOf('DocumentID'), paid: ls.indexOf('LinePaid'),
               type: ls.indexOf('DocType') };
  const di = { id: ds.indexOf('DocumentID'), paid: ds.indexOf('AmountPaid'),
               company: ds.indexOf('Company'), number: ds.indexOf('Number') };

  const byDoc = {};
  linesSh.getRange(2, 1, linesSh.getLastRow() - 1, ls.length).getValues().forEach(function (r) {
    if (!r[li.id] || r[li.type] === 'BankTransaction') return;
    if (!byDoc[r[li.id]]) byDoc[r[li.id]] = 0;
    byDoc[r[li.id]] += num_(r[li.paid]);
  });

  const rows = [];
  docsSh.getRange(2, 1, docsSh.getLastRow() - 1, ds.length).getValues().forEach(function (r) {
    const id = r[di.id];
    if (!id || byDoc[id] === undefined) return;
    const diff = round2_(num_(r[di.paid]) - byDoc[id]);
    if (diff !== 0) {
      rows.push([r[di.company], r[di.number], id, num_(r[di.paid]), round2_(byDoc[id]), diff]);
    }
  });

  if (!rows.length) {
    Log.info('Payment allocation matches for all documents.');
    return;
  }
  writeReport_('_Line_Allocation_Check',
               ['Company', 'Number', 'DocumentID', 'DocPaid', 'SumLinePaid', 'Diff'], rows);
  Log.info('Differences: %s. See the _Line_Allocation_Check sheet.', rows.length);
}

/**
 * Compares bank account movements from the BankTransactions sheet
 * with movements from the Xero report (Bank_Balances sheet).
 * A non-zero Diff means a source of movement is not exported.
 * Candidates: batch payments, manual journals on the bank account,
 * expense claims.
 */
function verifyBankMovement() {
  const bankSh = sheet_('BankTransactions'), balSh = sheet_('Bank_Balances');
  if (!bankSh || !balSh || bankSh.getLastRow() < 2 || balSh.getLastRow() < 2) {
    Log.warn('No data to reconcile.');
    return;
  }

  const bs = SCHEMA.BankTransactions;
  const bi = { company: bs.indexOf('Company'), account: bs.indexOf('BankAccount'),
               dir: bs.indexOf('Direction'), total: bs.indexOf('Total') };

  const moved = {};
  bankSh.getRange(2, 1, bankSh.getLastRow() - 1, bs.length).getValues().forEach(function (r) {
    if (!r[bi.company] || !r[bi.account]) return;
    const key = r[bi.company] + '||' + r[bi.account];
    if (!moved[key]) moved[key] = { in: 0, out: 0 };
    const amount = num_(r[bi.total]);
    if (r[bi.dir] === 'In') moved[key].in += amount;
    else if (r[bi.dir] === 'Out') moved[key].out += amount;
  });

  const ls = SCHEMA.Bank_Balances;
  const li = { company: ls.indexOf('Company'), account: ls.indexOf('BankAccount'),
               received: ls.indexOf('CashReceived'), spent: ls.indexOf('CashSpent') };

  const rows = [];
  balSh.getRange(2, 1, balSh.getLastRow() - 1, ls.length).getValues().forEach(function (r) {
    if (!r[li.company] || !r[li.account]) return;
    const key = r[li.company] + '||' + r[li.account];
    const m = moved[key] || { in: 0, out: 0 };
    const diffIn = round2_(num_(r[li.received]) - m.in);
    const diffOut = round2_(num_(r[li.spent]) - m.out);
    rows.push([r[li.company], r[li.account],
               num_(r[li.received]), round2_(m.in), diffIn,
               num_(r[li.spent]), round2_(m.out), diffOut,
               (diffIn === 0 && diffOut === 0) ? 'ok' : 'CHECK']);
  });

  writeReport_('_Bank_Movement_Check',
               ['Company', 'BankAccount',
                'Xero Received', 'Exported In', 'Diff In',
                'Xero Spent', 'Exported Out', 'Diff Out', 'Status'], rows);
  Log.info('Reconciliation in the _Bank_Movement_Check sheet.');
}

// ================================================================
//  13. PAYMENT SPLIT DIAGNOSTICS
// ================================================================

/** How many payments remain unsplit, per entity. */
function countUnsplitPayments() {
  const sh = sheet_('Payments');
  if (!sh || sh.getLastRow() < 2) { Log.warn('The Payments sheet is empty.'); return; }

  const s = SCHEMA.Payments;
  const i = { company: s.indexOf('Company'), code: s.indexOf('AccountCode'),
              name: s.indexOf('AccountName'), total: s.indexOf('PaymentTotal') };

  const stat = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, s.length).getValues().forEach(function (r) {
    const c = r[i.company];
    if (!c) return;
    if (!stat[c]) stat[c] = { split: 0, unsplit: 0, unsplitAmount: 0 };
    // Unsplit rows have an empty AccountCode and Accounts Payable/Receivable as AccountName
    if (!r[i.code] && /Accounts (Payable|Receivable)/.test(String(r[i.name]))) {
      stat[c].unsplit++;
      stat[c].unsplitAmount += num_(r[i.total]);
    } else {
      stat[c].split++;
    }
  });

  Object.keys(stat).forEach(function (c) {
    Log.info('%s: split rows: %s, not split: %s (amount %s)',
      c, stat[c].split, stat[c].unsplit, round2_(stat[c].unsplitAmount));
  });
}

/**
 * Why a specific payment was not split by line item.
 * The document number is taken from the Payments sheet:
 *
 *   diagnosePaymentSplit('entity_cy', 'Aug25')
 */
function diagnosePaymentSplit(company, invoiceNumber) {
  Log.info('=== Diagnosis %s / %s ===', company, invoiceNumber);
  Log.info('splitPaymentsByCategory: %s', CONFIG.splitPaymentsByCategory);
  Log.info('includeLineItems: %s', CONFIG.includeLineItems);
  Log.info('resolvePaymentCategories: %s', CONFIG.resolvePaymentCategories);

  const paySh = sheet_('Payments');
  if (!paySh || paySh.getLastRow() < 2) { Log.warn('The Payments sheet is empty.'); return; }

  const ps = SCHEMA.Payments;
  const pi = { company: ps.indexOf('Company'), number: ps.indexOf('InvoiceNumber'),
               invId: ps.indexOf('InvoiceID'), acc: ps.indexOf('AccountName'),
               total: ps.indexOf('PaymentTotal') };

  let invoiceId = '', found = 0;
  paySh.getRange(2, 1, paySh.getLastRow() - 1, ps.length).getValues().forEach(function (r) {
    if (r[pi.company] !== company || String(r[pi.number]) !== String(invoiceNumber)) return;
    found++;
    invoiceId = r[pi.invId];
    Log.info('row: line item "%s", amount %s', r[pi.acc], r[pi.total]);
  });

  if (!found) { Log.info('No payment with this document number found.'); return; }
  Log.warn('InvoiceID: %s', invoiceId || 'EMPTY <- payment without a document, nothing to split by');
  if (!invoiceId) return;

  const docsSh = sheet_('Documents');
  const ds = SCHEMA.Documents;
  let inDocs = false;
  if (docsSh && docsSh.getLastRow() > 1) {
    docsSh.getRange(2, 1, docsSh.getLastRow() - 1, ds.length).getValues().forEach(function (r) {
      if (r[ds.indexOf('DocumentID')] === invoiceId) inDocs = true;
    });
  }
  Log.info('document in the Documents sheet: %s%s', inDocs,
    inDocs ? '' : ' <- issued outside the fromDate period, an extra request is needed');

  resetCategoryIndex_();
  const hit = categoryIndex_()[invoiceId];
  if (hit) {
    Log.info('rows in Lines: %s line items, amount %s', hit.accounts.length, round2_(hit.grossSum));
    hit.accounts.forEach(function (a) {
      Log.info('   %s — %s: %s', a.code, a.name, round2_(a.gross));
    });
    Log.warn('-> the split should work. If it is missing in the sheet, rerun the export.');
    return;
  }

  Log.warn('NO rows in Lines <- this is where the fallback came from');

  const fetched = fetchDocCategories_(company, invoiceId);
  if (!fetched) {
    Log.warn('the extra request returned nothing: check resolvePaymentCategories, ' +
               'the resolveCategoriesLimit and document availability.');
    return;
  }
  Log.info('the extra request returned %s line items:', fetched.accounts.length);
  fetched.accounts.forEach(function (a) {
    Log.info('   %s — %s: %s', a.code, a.name, round2_(a.gross));
  });
  Log.info('-> the extra request works but did not during the export: probably ' +
             'the resolveCategoriesLimit was exhausted (currently %s).',
             CONFIG.resolveCategoriesLimit);
}

// ================================================================
//  14. CONNECTION DIAGNOSTICS
// ================================================================

/** Full connection check. Changes nothing, read-only. */
function xeroDiagnose() {
  const props = PropertiesService.getScriptProperties();
  const service = getXeroService_();

  Log.info('=== 1. SERVICE STATE ===');
  Log.info('hasAccess: %s', service.hasAccess());
  Log.error('lastError: %s', service.getLastError() || '—');

  const raw = props.getProperty('oauth2.xero');
  if (!raw) { Log.info('No token. Authorization from scratch is required.'); return; }

  const token = JSON.parse(raw);
  Log.warn('refresh_token: %s', token.refresh_token ? 'present' : 'MISSING <- refresh not possible');
  if (token.granted_time && token.expires_in) {
    const minutes = Math.round(((token.granted_time + Number(token.expires_in)) * 1000 - Date.now()) / 60000);
    const idleDays = Math.floor((Date.now() - token.granted_time * 1000) / 86400000);
    Log.warn('access token expires in %s min %s', minutes, minutes < 0 ? '<- EXPIRED' : '');
    Log.warn('refreshed %s days ago %s', idleDays, idleDays > 60 ? '<- >60 days, refresh token is dead' : '');
  }
  Log.info('scope: %s', token.scope || '—');
  Log.warn('<- compare this list with the endpoints you use');

  Log.info('\n=== 2. FORCED REFRESH ===');
  try { service.refresh(); Log.info('refresh succeeded'); }
  catch (err) { Log.error('refresh FAILED: %s\n-> resetXero() and authorize again.', err.message); }

  Log.info('\n=== 3. RAW REQUEST /connections ===');
  let connections = [];
  try {
    const res = UrlFetchApp.fetch('https://api.xero.com/connections', {
      headers: { Authorization: 'Bearer ' + service.getAccessToken(), Accept: 'application/json' },
      muteHttpExceptions: true
    });
    Log.info('HTTP %s | %s', res.getResponseCode(), res.getContentText().slice(0, 400));
    if (res.getResponseCode() === 200) connections = JSON.parse(res.getContentText());
  } catch (err) { Log.error('Request failed: %s', err.message); }

  Log.info('\n=== 4. ORGANISATIONS ===');
  const stored = readXeroTenants_();
  const live = {};
  connections.forEach(function (c) {
    live[c.tenantId] = c.tenantName;
    Log.info('in Xero: "%s"', c.tenantName);
  });
  Object.keys(XERO_COMPANIES).forEach(function (alias) {
    const t = stored[alias];
    if (!t) { Log.warn('%s: NOT LINKED', alias); return; }
    Log.info('%s → %s %s', alias, t.tenantName,
      live[t.tenantId] ? 'ok' : '<- tenantId missing from connections');
  });

  Log.info('\n=== 5. ENDPOINTS PER ENTITY ===');
  const probes = ['Organisation', 'Invoices?pageSize=1', 'Payments?pageSize=1',
                  'Prepayments?pageSize=1', 'BankTransactions?pageSize=1',
                  'Reports/ProfitAndLoss'];
  Object.keys(XERO_COMPANIES).forEach(function (alias) {
    const t = stored[alias];
    if (!t) return;
    probes.forEach(function (path) {
      const res = UrlFetchApp.fetch('https://api.xero.com/api.xro/2.0/' + path, {
        headers: {
          Authorization: 'Bearer ' + service.getAccessToken(),
          'xero-tenant-id': t.tenantId,
          Accept: 'application/json'
        },
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      Log.info('%s / %s → %s %s', alias, path.split('?')[0], code,
        code === 200 ? '' : '<- no scope or no access');
    });
  });

  Log.info('\n=== 6. EXPORT ===');
  Log.info(exportStatusText_());
}

// ================================================================
//  15. REPORTS
// ================================================================

function buildAllReports() {
  if (CONFIG.mapToPnl) { tagWithPnl(); buildCashflow(); }
  buildAgedReport_('AR');
  buildAgedReport_('AP');
  buildArApSummary();
  buildBankOverview();
  buildPnl();
  if (CONFIG.includePnlAdjustments) buildPnlAdjustments();
  buildBalanceSheetPivot();
  Log.info('Reports rebuilt.');
}

// ---------- Final P&L: documents plus journals ----------

/**
 * Service columns of the PnL sheet, placed after the Lines format columns.
 * Attachments at the end: links are taken from the Documents sheet by
 * DocumentID, no separate Xero requests.
 */
const PNL_SHEET_EXTRA = ['Source', 'Reason', 'PnL Section', 'PnL Amount', 'Attachments'];

/**
 * Document statuses that Xero does NOT count in the P&L.
 * Drafts and voided documents sit in Lines like the rest,
 * but are not in the report; if not dropped, documents would
 * add up to more than the report and adjustments would go negative.
 */
const NON_POSTED_STATUSES = {
  DRAFT: 1, SUBMITTED: 1, VOIDED: 1, DELETED: 1
};

/**
 * Row sign in P&L terms.
 *
 * In Xero revenue is a credit, expense is a debit. Document lines
 * carry no debit/credit flag, but it follows clearly from the type:
 *   credit: invoice, supplier credit note, bank inflow
 *   debit: bill, customer credit note, bank outflow
 *
 * For a REVENUE account a credit increases revenue, a debit reduces it.
 * For EXPENSE the opposite. This way refunds and reversals reach the report
 * with the correct sign instead of doubling turnover.
 */
function pnlSignFromDoc_(docType, direction, accountClass) {
  let isCredit;
  if (docType === 'Invoice' || docType === 'Credit Note (AP)') isCredit = true;
  else if (docType === 'Bill' || docType === 'Credit Note (AR)') isCredit = false;
  else if (docType === 'BankTransaction') isCredit = direction === 'In';
  else return 0;                                  // prepayments do not affect the P&L

  return accountClass === 'REVENUE' ? (isCredit ? 1 : -1)
                                    : (isCredit ? -1 : 1);
}

/**
 * A clear label for an adjustment row: by system account,
 * or by name when there is none.
 */
function adjustmentReason_(acc, accountName) {
  const system = acc.system || '';
  if (KEEP_SYSTEM_ACCOUNTS[system]) return KEEP_SYSTEM_ACCOUNTS[system];

  const name = String(accountName || acc.name || '').toLowerCase();
  if (/currency|forex|exchange/.test(name)) return 'FX difference';
  if (/deprecia|amortis|amortiz/.test(name)) return 'Depreciation';
  if (/round/.test(name)) return 'Rounding';
  if (/payroll|wage|salar/.test(name)) return 'Payroll';
  return 'Not in documents';
}

/**
 * TRANSACTIONS for building the P&L, in the same field format as the Lines sheet,
 * plus four service columns at the end:
 *
 *   Source      Documents | Journals: where the row came from
 *   Reason      for journals: FX difference, manual journal, depreciation...
 *   PnL Section Xero report section (Income, Cost of Sales, Operating Expenses)
 *   PnL Amount  signed amount in P&L terms
 *
 * Sources:
 *   Lines    = everything that went through documents and bank
 *   Journals = what is not in documents: FX differences, manual
 *              journals, depreciation, rounding, payroll
 *
 * Only REVENUE and EXPENSE accounts are taken: balance sheet accounts are dropped.
 *
 * PnL Amount: revenue positive, expenses positive.
 * Profit = sum of revenue sections minus sum of expense sections.
 * Refunds and reversals come in negative and reduce turnover instead of doubling it.
 * For a summary, build a regular pivot table on this sheet.
 */
function buildPnl() {
  const headers = SCHEMA.Lines.concat(PNL_SHEET_EXTRA);
  const blank = new Array(SCHEMA.Lines.length).fill('');
  const iDateCol = SCHEMA.Lines.indexOf('Date');
  const maps = {};
  const rows = [];
  const docTotals = {};          // entity||month||account -> amount from documents

  // Attachment links: DocumentID -> links from the Documents sheet
  const attachById = {};
  const docsSh = sheet_('Documents');
  if (docsSh && docsSh.getLastRow() > 1 && docsSh.getLastColumn() >= SCHEMA.Documents.length) {
    const dsc = SCHEMA.Documents;
    const di = { id: dsc.indexOf('DocumentID'), links: dsc.indexOf('Attachments') };
    docsSh.getRange(2, 1, docsSh.getLastRow() - 1, dsc.length).getValues().forEach(function (r) {
      const links = String(r[di.links] || '').trim();
      if (r[di.id] && links && links !== '—') attachById[String(r[di.id])] = links;
    });
  }

  // Directions of bank transactions: they are not in the Lines sheet
  const dirById = {};
  const bankSh = sheet_('BankTransactions');
  if (bankSh && bankSh.getLastRow() > 1) {
    const bs = SCHEMA.BankTransactions;
    const bi = { id: bs.indexOf('TransactionID'), dir: bs.indexOf('Direction') };
    bankSh.getRange(2, 1, bankSh.getLastRow() - 1, bs.length).getValues().forEach(function (r) {
      if (r[bi.id]) dirById[String(r[bi.id]).split(':')[0]] = r[bi.dir];
    });
  }

  // 1. Document lines are carried over as is, columns unchanged
  const linesSh = sheet_('Lines');
  if (linesSh && linesSh.getLastRow() > 1) {
    const ls = SCHEMA.Lines;
    const i = {
      company: ls.indexOf('Company'), docType: ls.indexOf('DocType'),
      docId: ls.indexOf('DocumentID'), code: ls.indexOf('AccountCode'),
      status: ls.indexOf('DocStatus'), amount: ls.indexOf('LineAmount')
    };

    let skippedDraft = 0;

    linesSh.getRange(2, 1, linesSh.getLastRow() - 1, ls.length).getValues().forEach(function (r) {
      const company = r[i.company];
      if (!company) return;

      // Drafts and voided documents are excluded from the P&L
      if (NON_POSTED_STATUSES[String(r[i.status]).toUpperCase()]) { skippedDraft++; return; }

      if (!maps[company]) maps[company] = pnlMap_(company);
      const acc = maps[company][String(r[i.code])] || {};
      if (acc.class !== 'REVENUE' && acc.class !== 'EXPENSE') return;

      const sign = pnlSignFromDoc_(r[i.docType], dirById[String(r[i.docId])], acc.class);
      if (!sign) return;

      const pnlAmount = round2_(sign * num_(r[i.amount]));
      rows.push(r.concat(['Documents', '', acc.section || 'Other P&L', pnlAmount,
                          attachById[String(r[i.docId])] || '']));

      // Collect totals per account and month: they are used to calculate
      // what is missing compared with the Xero report
      if (r[iDateCol] instanceof Date) {
        const key = [company, Utilities.formatDate(r[iDateCol], 'UTC', 'yyyy-MM'),
                     String(r[i.code])].join('||');
        docTotals[key] = (docTotals[key] || 0) + pnlAmount;
      }
    });

    if (skippedDraft) {
      Log.warn('PnL: skipped rows in draft/voided status: %s', skippedDraft);
    }
  }

  // 2. Missing journal entries are laid out in the same columns.
  //    Fields that a journal entry simply does not have stay empty:
  //    quantity, price, line payment
  const jSh = sheet_('Journals');
  if (jSh && jSh.getLastRow() > 1) {
    const js = SCHEMA.Journals;
    const ls = SCHEMA.Lines;
    const j = {
      company: js.indexOf('Company'), id: js.indexOf('JournalID'),
      date: js.indexOf('JournalDate'), number: js.indexOf('JournalNumber'),
      source: js.indexOf('SourceType'), reason: js.indexOf('Reason'),
      ref: js.indexOf('Reference'), lineId: js.indexOf('JournalLineID'),
      code: js.indexOf('AccountCode'), name: js.indexOf('AccountName'),
      type: js.indexOf('AccountType'), desc: js.indexOf('Description'),
      net: js.indexOf('NetAmount'), tax: js.indexOf('TaxAmount'),
      taxType: js.indexOf('TaxType'), tracking: js.indexOf('Tracking')
    };
    const c = {
      company: ls.indexOf('Company'), docType: ls.indexOf('DocType'),
      docId: ls.indexOf('DocumentID'), number: ls.indexOf('Number'),
      date: ls.indexOf('Date'), status: ls.indexOf('DocStatus'),
      lineId: ls.indexOf('LineItemID'), desc: ls.indexOf('Description'),
      code: ls.indexOf('AccountCode'), name: ls.indexOf('AccountName'),
      cls: ls.indexOf('AccountClass'), type: ls.indexOf('AccountType'),
      taxType: ls.indexOf('TaxType'), tax: ls.indexOf('TaxAmount'),
      amount: ls.indexOf('LineAmount'), gross: ls.indexOf('LineGross'),
      tracking: ls.indexOf('Tracking')
    };

    jSh.getRange(2, 1, jSh.getLastRow() - 1, js.length).getValues().forEach(function (r) {
      const company = r[j.company];
      if (!company) return;

      if (!maps[company]) maps[company] = pnlMap_(company);
      const acc = maps[company][String(r[j.code])] || {};
      if (acc.class !== 'REVENUE' && acc.class !== 'EXPENSE') return;

      // In journals debit is positive, credit negative
      const net = num_(r[j.net]);
      const amount = acc.class === 'REVENUE' ? -net : net;

      const row = blank.slice();
      row[c.company]  = company;
      row[c.docType]  = 'Journal';
      row[c.docId]    = r[j.id];
      row[c.number]   = r[j.number];
      row[c.date]     = r[j.date];
      row[c.status]   = r[j.source];
      row[c.lineId]   = r[j.lineId];
      row[c.desc]     = r[j.desc] || r[j.ref] || '';
      row[c.code]     = r[j.code];
      row[c.name]     = r[j.name] || acc.name || '';
      row[c.cls]      = acc.class;
      row[c.type]     = r[j.type];
      row[c.taxType]  = r[j.taxType];
      row[c.tax]      = num_(r[j.tax]);
      row[c.amount]   = round2_(amount);
      row[c.gross]    = round2_(amount + num_(r[j.tax]));
      row[c.tracking] = r[j.tracking];

      rows.push(row.concat(['Journals', r[j.reason] || '',
                            acc.section || 'Other P&L', round2_(amount), '']));

      if (r[j.date] instanceof Date) {
        const key = [company, Utilities.formatDate(r[j.date], 'UTC', 'yyyy-MM'),
                     String(r[j.code])].join('||');
        docTotals[key] = (docTotals[key] || 0) + round2_(amount);
      }
    });
  }

  // 3. What is missing from the Xero report, as separate rows.
  //    This covers FX differences, depreciation, rounding:
  //    they have no documents, and line-level entries live in /Journals,
  //    which is only available on the Advanced plan.
  if (CONFIG.includePnlAdjustments) {
    const repSh = sheet_('PnL_Report');
    if (repSh && repSh.getLastRow() > 1) {
      const rs = SCHEMA.PnL_Report;
      const k = {
        company: rs.indexOf('Company'), period: rs.indexOf('PeriodDate'),
        section: rs.indexOf('Section'), accId: rs.indexOf('AccountID'),
        account: rs.indexOf('Account'), amount: rs.indexOf('Amount')
      };
      const c = {
        company: SCHEMA.Lines.indexOf('Company'), docType: SCHEMA.Lines.indexOf('DocType'),
        docId: SCHEMA.Lines.indexOf('DocumentID'), date: iDateCol,
        status: SCHEMA.Lines.indexOf('DocStatus'), desc: SCHEMA.Lines.indexOf('Description'),
        code: SCHEMA.Lines.indexOf('AccountCode'), name: SCHEMA.Lines.indexOf('AccountName'),
        cls: SCHEMA.Lines.indexOf('AccountClass'), type: SCHEMA.Lines.indexOf('AccountType'),
        amount: SCHEMA.Lines.indexOf('LineAmount'), gross: SCHEMA.Lines.indexOf('LineGross')
      };

      let matched = 0, added = 0, noAccount = 0, badPeriod = 0;

      repSh.getRange(2, 1, repSh.getLastRow() - 1, rs.length).getValues().forEach(function (r) {
        const company = r[k.company];
        if (!company || /^TOTAL/.test(String(r[k.section]))) return;

        const periodEnd = parsePeriodEnd_(r[k.period]);
        if (!(periodEnd instanceof Date)) { badPeriod++; return; }

        if (!maps[company]) maps[company] = pnlMap_(company);
        const acc = maps[company][String(r[k.accId])] || {};
        if (acc.class !== 'REVENUE' && acc.class !== 'EXPENSE') { noAccount++; return; }

        const month = Utilities.formatDate(periodEnd, 'UTC', 'yyyy-MM');
        const key = [company, month, String(acc.code || '')].join('||');
        const diff = round2_(num_(r[k.amount]) - round2_(docTotals[key] || 0));
        if (Math.abs(diff) < 0.01) { matched++; return; }
        added++;

        const row = blank.slice();
        row[c.company] = company;
        row[c.docType] = 'Adjustment';
        row[c.docId]   = company + ':' + month + ':' + (acc.code || r[k.account]);
        row[c.date]    = periodEnd;
        row[c.status]  = 'From Xero report';
        row[c.desc]    = r[k.account];
        row[c.code]    = acc.code || '';
        row[c.name]    = acc.name || r[k.account];
        row[c.cls]     = acc.class;
        row[c.type]    = acc.type || '';
        row[c.amount]  = diff;
        row[c.gross]   = diff;

        rows.push(row.concat(['Xero report', adjustmentReason_(acc, r[k.account]),
                              acc.section || r[k.section], diff, '']));
      });

      Log.info('Adjustments: added %s, matched %s, account not recognised %s, ' +
                 'period not parsed %s', added, matched, noAccount, badPeriod);
      if (noAccount && !added) {
        Log.warn('<- The report seems to have no AccountID. Set ' +
                   'reportsStandardLayout: false and export the report again.');
      }
    } else {
      Log.warn('The PnL_Report sheet is empty, there will be no adjustments.');
    }
  }

  // Chronological order: entity, date, account
  const iDate = SCHEMA.Lines.indexOf('Date');
  const iCompany = SCHEMA.Lines.indexOf('Company');
  const iCode = SCHEMA.Lines.indexOf('AccountCode');
  rows.sort(function (a, b) {
    if (a[iCompany] !== b[iCompany]) return String(a[iCompany]).localeCompare(String(b[iCompany]));
    const da = a[iDate] instanceof Date ? a[iDate].getTime() : 0;
    const db = b[iDate] instanceof Date ? b[iDate].getTime() : 0;
    if (da !== db) return da - db;
    return String(a[iCode]).localeCompare(String(b[iCode]));
  });

  writeReport_('PnL', headers, rows, 3, true);
  Log.info('PnL: transactions %s', rows.length);
}

// ---------- Monthly balance sheet ----------

/**
 * Pivot of the Balance_Sheet sheet: accounts in rows, periods in columns.
 * Section total rows (prefixed TOTAL) are moved to the bottom of their
 * section so they are not added up with regular accounts.
 */
function buildBalanceSheetPivot() {
  const sh = sheet_('Balance_Sheet');
  if (!sh || sh.getLastRow() < 2) {
    Log.warn('The Balance_Sheet sheet is empty.');
    return;
  }

  const s = SCHEMA.Balance_Sheet;
  const i = {
    company: s.indexOf('Company'), period: s.indexOf('PeriodDate'),
    section: s.indexOf('Section'), account: s.indexOf('Account'),
    amount: s.indexOf('Amount')
  };

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, s.length).getValues();
  const periods = [], agg = {};

  data.forEach(function (r) {
    if (!r[i.company] || !r[i.account]) return;
    const period = String(r[i.period]);
    if (periods.indexOf(period) < 0) periods.push(period);

    const key = [r[i.company], r[i.section], r[i.account]].join('||');
    if (!agg[key]) agg[key] = {};
    agg[key][period] = num_(r[i.amount]);
  });

  // Periods come in Xero report order, which is kept:
  // alphabetical sorting would mix up the months

  const rows = Object.keys(agg).sort(function (a, b) {
    const pa = a.split('||'), pb = b.split('||');
    if (pa[0] !== pb[0]) return pa[0].localeCompare(pb[0]);
    // Total rows go to the end of their section
    const ta = /^TOTAL/.test(pa[1]), tb = /^TOTAL/.test(pb[1]);
    const sa = pa[1].replace(/^TOTAL /, ''), sb = pb[1].replace(/^TOTAL /, '');
    if (sa !== sb) return sa.localeCompare(sb);
    if (ta !== tb) return ta ? 1 : -1;
    return pa[2].localeCompare(pb[2]);
  }).map(function (key) {
    const p = key.split('||');
    const row = [p[0], p[1], p[2]];
    periods.forEach(function (m) { row.push(round2_(agg[key][m] || 0)); });
    return row;
  });

  writeReport_('Balance_Sheet_Pivot',
    ['Company', 'Section', 'Account'].concat(periods), rows, 3);

  Log.info('Balance_Sheet_Pivot: rows %s, periods %s', rows.length, periods.length);
}

// ---------- Bank account balances ----------

/**
 * Two sheets in the style of the Xero dashboard:
 *   Bank_Overview      = Xero balance and an ESTIMATE of the statement balance per account
 *   Bank_Unreconciled  = unreconciled transactions that make up the gap
 *
 * IMPORTANT ABOUT STATEMENT BALANCE
 * The real statement balance is not available via the API: Reports/BankStatement
 * has been closed since April 2024 (requires signing an addendum to the
 * developer terms), and the Finance API with Bank Statements Plus is granted on request
 * for lending scenarios.
 *
 * Therefore the StatementBalanceEst column is CALCULATED:
 *     Xero balance minus the net of unreconciled transactions
 * Same logic as Xero itself: the gap between the ledger balance
 * and the statement balance is what is not yet reconciled.
 *
 * Where the estimate differs from reality: statement lines that came
 * from the bank but are not yet matched to any Xero transaction. They
 * raise the statement balance but are not in the ledger; they cannot be seen
 * without access to the statement report.
 */
function buildBankOverview() {
  const balSh = sheet_('Bank_Balances');
  const bankSh = sheet_('BankTransactions');
  if (!balSh || balSh.getLastRow() < 2) {
    Log.warn('No data: export balances first.');
    return;
  }

  // 1. Unreconciled transactions from the BankTransactions sheet
  const unrec = {};          // company||account → {in, out, count}
  const detail = [];

  if (bankSh && bankSh.getLastRow() > 1) {
    const bs = SCHEMA.BankTransactions;
    const i = {
      company:  bs.indexOf('Company'),
      source:   bs.indexOf('Source'),
      id:       bs.indexOf('TransactionID'),
      type:     bs.indexOf('Type'),
      date:     bs.indexOf('Date'),
      dir:      bs.indexOf('Direction'),
      account:  bs.indexOf('BankAccount'),
      contra:   bs.indexOf('ContraAccountNames'),
      contact:  bs.indexOf('Contact'),
      ref:      bs.indexOf('Reference'),
      rec:      bs.indexOf('Reconciled'),
      currency: bs.indexOf('Currency'),
      total:    bs.indexOf('Total')
    };

    bankSh.getRange(2, 1, bankSh.getLastRow() - 1, bs.length).getValues().forEach(function (r) {
      if (!r[i.company] || !r[i.account]) return;
      if (r[i.rec] === true || String(r[i.rec]).toUpperCase() === 'TRUE') return;

      const key = r[i.company] + '||' + r[i.account];
      if (!unrec[key]) unrec[key] = { in: 0, out: 0, count: 0 };

      const amount = num_(r[i.total]);
      if (r[i.dir] === 'In') unrec[key].in += amount;
      else if (r[i.dir] === 'Out') unrec[key].out += amount;
      unrec[key].count++;

      detail.push([r[i.company], r[i.account], r[i.date], r[i.dir],
                   r[i.source], r[i.type], r[i.id],
                   r[i.contact] || '', r[i.contra] || '', r[i.ref] || '',
                   r[i.currency] || '', amount,
                   r[i.dir] === 'Out' ? -amount : amount]);
    });
  }

  // 2. Xero balance from Bank_Balances (BankSummary closing balance)
  const ls = SCHEMA.Bank_Balances;
  const li = {
    company: ls.indexOf('Company'),
    toDate:  ls.indexOf('ToDate'),
    account: ls.indexOf('BankAccount'),
    closing: ls.indexOf('ClosingBalance')
  };

  const rows = [];
  balSh.getRange(2, 1, balSh.getLastRow() - 1, ls.length).getValues().forEach(function (r) {
    if (!r[li.company] || !r[li.account]) return;

    const key = r[li.company] + '||' + r[li.account];
    const u = unrec[key] || { in: 0, out: 0, count: 0 };
    const netUnrec = u.in - u.out;
    const xeroBalance = num_(r[li.closing]);

    rows.push([r[li.company], r[li.account], r[li.toDate],
               round2_(xeroBalance),
               round2_(xeroBalance - netUnrec),      // statement balance estimate
               round2_(netUnrec),
               round2_(u.in), round2_(u.out), u.count]);
  });

  // Sorting: entity, then account
  rows.sort(function (a, b) {
    return a[0] === b[0] ? String(a[1]).localeCompare(String(b[1]))
                         : String(a[0]).localeCompare(String(b[0]));
  });

  writeReport_('Bank_Overview',
    ['Company', 'BankAccount', 'AsOf', 'BalanceInXero', 'StatementBalanceEst',
     'NetUnreconciled', 'Unreconciled In', 'Unreconciled Out', 'UnreconciledCount'],
    rows, 2);

  detail.sort(function (a, b) {
    if (a[0] !== b[0]) return String(a[0]).localeCompare(String(b[0]));
    if (a[1] !== b[1]) return String(a[1]).localeCompare(String(b[1]));
    return (b[2] instanceof Date ? b[2].getTime() : 0) - (a[2] instanceof Date ? a[2].getTime() : 0);
  });

  writeReport_('Bank_Unreconciled',
    ['Company', 'BankAccount', 'Date', 'Direction', 'Source', 'Type', 'TransactionID',
     'Contact', 'ContraAccount', 'Reference', 'Currency', 'Amount', 'SignedAmount'],
    detail, 2, true);

  Log.info('Bank_Overview: %s accounts, unreconciled transactions: %s', rows.length, detail.length);
}

// ---------- 15.1 P&L sections ----------

const PNL_EXTRA = ['PnL Section', 'PnL Account', 'Account Class', 'Account Type', 'Cash Impact'];

/** Account code -> P&L report section. The grouping comes from Xero itself. */
function pnlMap_(company) {
  const map = accounts_(company);

  Object.keys(map).forEach(function (k) {
    const e = map[k];
    if (!e.section) {
      e.section = (e.class === 'REVENUE' || e.class === 'EXPENSE') ? 'Other P&L' : 'Balance Sheet';
    }
  });

  const pnl = xeroFetch(company, 'Reports/ProfitAndLoss?fromDate=' +
                        CONFIG.fromDate + '&toDate=' + balanceDate_());

  flattenReport_(pnl).forEach(function (row) {
    if (!row.accountId || !row.section) return;
    const e = map[row.accountId];
    if (e) e.section = row.section;          // object shared with the entry by code
  });

  return map;
}

/**
 * Tags money movements with P&L sections.
 *
 * Split payments carry a regular account code in ContraAccountCodes,
 * so they are handled with the same logic as bank transactions.
 * Unsplit payments (document not found) are flagged separately.
 */
function tagWithPnl() {
  const sh = sheet_('BankTransactions');
  if (!sh || sh.getLastRow() < 2) return;

  const base = SCHEMA.BankTransactions;

  if (sh.getRange(1, base.length + 1).getValue() !== PNL_EXTRA[0]) {
    sh.getRange(1, base.length + 1, 1, PNL_EXTRA.length).setValues([PNL_EXTRA]).setFontWeight('bold');
  }

  const count = sh.getLastRow() - 1;
  const data = sh.getRange(2, 1, count, base.length).getValues();
  const i = {
    company:   base.indexOf('Company'),
    source:    base.indexOf('Source'),
    direction: base.indexOf('Direction'),
    codes:     base.indexOf('ContraAccountCodes'),
    names:     base.indexOf('ContraAccountNames'),
    total:     base.indexOf('Total')
  };

  const maps = {};

  const out = data.map(function (r) {
    const company = r[i.company];
    if (!company) return ['', '', '', '', ''];

    const impact = r[i.direction] === 'Out' ? -num_(r[i.total]) : num_(r[i.total]);

    if (r[i.source] === 'BankTransfer') {
      // Two legs give -amount and +amount: zero for the entity,
      // but the movement on each account is visible
      return ['Internal Transfer', '', '', '', impact];
    }

    if (!maps[company]) maps[company] = pnlMap_(company);

    const codes = String(r[i.codes] || '').split('|')
      .map(function (c) { return c.trim(); }).filter(String);

    // Unsplit payment: ContraAccountCodes holds an InvoiceID, not an account code
    if (r[i.source] === 'Payment' && (!codes.length || !maps[company][codes[0]])) {
      return ['Unattributed payment', r[i.names] || '', '', '', impact];
    }

    if (!codes.length) return ['Unclassified', '', '', '', impact];

    if (codes.length > 1) {
      const sections = [];
      codes.forEach(function (c) {
        const s = (maps[company][c] || {}).section || 'Unknown';
        if (sections.indexOf(s) < 0) sections.push(s);
      });
      return [sections.length === 1 ? sections[0] : 'Split: ' + sections.join(' / '),
              codes.map(function (c) { return (maps[company][c] || {}).name || c; }).join(' | '),
              '', '', impact];
    }

    const e = maps[company][codes[0]];
    if (!e) return ['Unknown account ' + codes[0], '', '', '', impact];
    return [e.section, e.name, e.class, e.type, impact];
  });

  sh.getRange(2, base.length + 1, count, PNL_EXTRA.length).setValues(out);
  sh.getRange(2, base.length + PNL_EXTRA.length, count, 1).setNumberFormat('#,##0.00');
}

/** Entity x P&L section x month. */
function buildCashflow() {
  const sh = sheet_('BankTransactions');
  const headers = SCHEMA.BankTransactions.concat(PNL_EXTRA);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < headers.length) return;

  const i = {
    company: headers.indexOf('Company'),
    date:    headers.indexOf('Date'),
    section: headers.indexOf('PnL Section'),
    impact:  headers.indexOf('Cash Impact')
  };

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
  const months = [], agg = {};

  data.forEach(function (r) {
    if (!(r[i.date] instanceof Date)) return;
    const month = Utilities.formatDate(r[i.date], 'UTC', 'yyyy-MM');
    if (months.indexOf(month) < 0) months.push(month);

    const key = r[i.company] + '||' + (r[i.section] || 'Unclassified');
    if (!agg[key]) agg[key] = {};
    agg[key][month] = (agg[key][month] || 0) + num_(r[i.impact]);
  });

  months.sort();

  const rows = Object.keys(agg).sort().map(function (key) {
    const parts = key.split('||');
    const row = [parts[0], parts[1]];
    let total = 0;
    months.forEach(function (m) {
      const v = round2_(agg[key][m] || 0);
      row.push(v);
      total += v;
    });
    row.push(round2_(total));
    return row;
  });

  writeReport_('Cashflow_By_Category',
               ['Company', 'PnL Section'].concat(months).concat(['Total']), rows, 2);
}

// ---------- 15.2 Aging ----------

function bucketOrder_() {
  const b = CONFIG.agingBuckets;
  return ['Current', '1-' + b[0], (b[0] + 1) + '-' + b[1], (b[1] + 1) + '-' + b[2], b[2] + '+'];
}

function buildAgedReport_(side) {
  const target = side === 'AR' ? 'AR_Aged_Report' : 'AP_Aged_Report';
  const buckets = bucketOrder_();

  const headers = ['Company', 'Contact', 'ContactID'].concat(buckets)
    .concat(['Total', 'Xero Outstanding', 'Xero Overdue', 'Diff']);

  const pivot = pivotAging_(side);
  const rows = [];

  Object.keys(XERO_COMPANIES).forEach(function (company) {
    const balances = contactBalances_(company, side);
    const contacts = pivot[company] || {};
    const sums = {};

    Object.keys(contacts)
      .sort(function (a, b) { return contacts[b].total - contacts[a].total; })
      .forEach(function (id) {
        const c = contacts[id];
        const bal = balances[id] || { outstanding: '', overdue: '' };
        const row = [company, c.name, id];

        buckets.forEach(function (b) {
          const v = round2_(c.buckets[b] || 0);
          row.push(v);
          sums[b] = (sums[b] || 0) + v;
        });

        row.push(round2_(c.total), bal.outstanding, bal.overdue,
                 bal.outstanding === '' ? '' : round2_(c.total - bal.outstanding));
        sums.total = (sums.total || 0) + c.total;
        rows.push(row);
      });

    const totalRow = [company, '- ENTITY TOTAL -', ''];
    buckets.forEach(function (b) { totalRow.push(round2_(sums[b] || 0)); });
    totalRow.push(round2_(sums.total || 0), '', '', '');
    rows.push(totalRow);
    rows.push(new Array(headers.length).fill(''));
  });

  writeReport_(target, headers, rows, 2);
}

/** The Aging sheet is shared by AR and AP, filtered by the Side column. */
function pivotAging_(side) {
  const sh = sheet_('Aging');
  if (!sh || sh.getLastRow() < 2) return {};

  const s = SCHEMA.Aging;
  const i = {
    company: s.indexOf('Company'), side: s.indexOf('Side'),
    contact: s.indexOf('Contact'), id: s.indexOf('ContactID'),
    bucket: s.indexOf('Bucket'), amount: s.indexOf('AmountDue')
  };

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, s.length).getValues();
  const out = {};

  data.forEach(function (r) {
    if (r[i.side] !== side) return;
    const company = r[i.company], id = r[i.id] || r[i.contact];
    if (!company || !id) return;
    if (!out[company]) out[company] = {};
    if (!out[company][id]) out[company][id] = { name: r[i.contact], buckets: {}, total: 0 };

    const e = out[company][id], amount = num_(r[i.amount]);
    e.buckets[r[i.bucket]] = (e.buckets[r[i.bucket]] || 0) + amount;
    e.total += amount;
  });

  return out;
}

/** Xero figures for reconciliation: balance as of TODAY in base currency. */
function contactBalances_(company, side) {
  const out = {};
  const field = side === 'AR' ? 'AccountsReceivable' : 'AccountsPayable';

  for (var page = 1; page <= 50; page++) {
    const res = xeroFetch(company, 'Contacts?page=' + page + '&pageSize=1000&where=' +
                          encodeURIComponent('ContactStatus=="ACTIVE"'));
    const contacts = res.Contacts || [];

    contacts.forEach(function (c) {
      const b = (c.Balances || {})[field];
      if (!b || (num_(b.Outstanding) === 0 && num_(b.Overdue) === 0)) return;
      out[c.ContactID] = { outstanding: round2_(b.Outstanding), overdue: round2_(b.Overdue) };
    });

    const p = res.pagination;
    if (p && p.pageCount) { if (page >= p.pageCount) break; }
    else if (contacts.length < 1000) break;
  }

  return out;
}

// ---------- 15.3 AR/AP summary ----------

function buildArApSummary() {
  const headers = ['Company', 'BalanceDate', 'Type', 'Bucket', 'Contacts', 'Amount',
                   'ControlAccount', 'ControlBalance', 'Diff'];
  const rows = [];

  Object.keys(XERO_COMPANIES).forEach(function (company) {
    const control = controlBalances_(company);

    [['AR', control.debtors], ['AP', control.creditors]].forEach(function (cfg) {
      const side = cfg[0], ctl = cfg[1];
      const pivot = pivotAging_(side)[company] || {};
      const date = balanceDate_();

      const byBucket = {};
      let total = 0, count = 0;

      Object.keys(pivot).forEach(function (id) {
        Object.keys(pivot[id].buckets).forEach(function (b) {
          byBucket[b] = (byBucket[b] || 0) + pivot[id].buckets[b];
        });
        total += pivot[id].total;
        count++;
      });

      bucketOrder_().forEach(function (b) {
        if (byBucket[b] === undefined) return;
        rows.push([company, date, side, b, '', round2_(byBucket[b]), '', '', '']);
      });

      rows.push([company, date, side, 'TOTAL', count, round2_(total),
                 ctl.name, ctl.balance, round2_(total - ctl.balance)]);
    });
  });

  writeReport_('AR_AP_Summary', headers, rows);
}

/** DEBTORS / CREDITORS control accounts from the trial balance. */
function controlBalances_(company) {
  const date = balanceDate_();
  const list = xeroFetch(company, 'Accounts').Accounts || [];
  const debtors   = list.filter(function (a) { return a.SystemAccount === 'DEBTORS'; })[0];
  const creditors = list.filter(function (a) { return a.SystemAccount === 'CREDITORS'; })[0];

  const byId = {};
  flattenReport_(xeroFetch(company, 'Reports/TrialBalance?date=' + date)).forEach(function (r) {
    if (r.accountId) byId[r.accountId] = r;
  });

  function pick(acc, sign) {
    if (!acc) return { name: '—', balance: 0 };
    const row = byId[acc.AccountID];
    return {
      name: acc.Name + ' (' + acc.Code + ')',
      balance: row ? round2_((num_(row['Debit']) - num_(row['Credit'])) * sign) : 0
    };
  }

  return { debtors: pick(debtors, 1), creditors: pick(creditors, -1) };
}

// ================================================================
//  16. SHEETS
// ================================================================

const SCHEMA = {
  /** Invoices, bills, credit notes, prepayments and overpayments together. */
  Documents: ['Company', 'DocType', 'Side', 'DocumentID', 'Number', 'XeroType', 'Status',
              'Date', 'DueDate', 'Contact', 'ContactID', 'Reference',
              'Currency', 'Rate', 'SubTotal', 'TotalTax', 'Total',
              'AmountPaid', 'AmountCredited', 'AmountDue', 'FullyPaidOn', 'UpdatedUTC',
              'HasAttachments', 'Attachments'],

  /**
   * Attachments uploaded to Google Drive.
   * The link is duplicated in the Attachments column of the Documents sheet.
   */
  Attachments: ['Company', 'DocType', 'DocumentID', 'Number', 'Contact', 'Date',
                'FileName', 'MimeType', 'SizeKb', 'DriveFileId', 'DriveUrl', 'DownloadedAt'],

  /**
   * Lines of documents and bank transactions. Link: DocumentID.
   * LinePaid / LineCredited / LineDue = calculated allocation of the document amount
   * across lines, see the comment on docLines_().
   */
  Lines: ['Company', 'DocType', 'DocumentID', 'Number', 'Date',
          'Contact', 'ContactID', 'DocStatus', 'Currency',
          'LineItemID', 'Description',
          'AccountCode', 'AccountName', 'AccountClass', 'AccountType', 'ItemCode',
          'Quantity', 'UnitAmount', 'DiscountRate', 'TaxType', 'TaxAmount',
          'LineAmount', 'LineGross', 'LineShare',
          'LinePaid', 'LineCredited', 'LineDue',
          'DocTotal', 'DocPaid', 'DocDue', 'Tracking'],

  /**
   * All money movements. Source: BankTransaction | Payment | BankTransfer.
   * A payment is split by document line items: one row per item,
   * TransactionID with suffix :1, :2 and so on.
   */
  BankTransactions: ['Company', 'Source', 'TransactionID', 'Type', 'Status', 'Date',
                     'Direction', 'BankAccount', 'BankAccountCode',
                     'FromAccount', 'ToAccount', 'ContraAccountCodes', 'ContraAccountNames',
                     'Contact', 'Reference', 'Reconciled',
                     'Currency', 'Rate', 'SubTotal', 'TotalTax', 'Total', 'UpdatedUTC',
                     'ReconciledBy', 'ReconciledAt'],

  /**
   * Payments split by the line items of the paid document.
   * PaymentTotal = whole payment amount, AllocatedAmount = share of this item.
   * The sum of AllocatedAmount within one PaymentID equals PaymentTotal.
   */
  Payments: ['Company', 'PaymentID', 'PaymentType', 'Status', 'Date',
             'InvoiceID', 'InvoiceNumber', 'InvoiceType', 'Contact',
             'BankAccount', 'BankAccountCode', 'Reference', 'Reconciled', 'Rate',
             'PaymentTotal',
             'AccountCode', 'AccountName', 'AccountClass', 'AccountType',
             'CategoryShare', 'AllocatedAmount', 'UpdatedUTC'],

  /**
   * Journal entries missing from other sheets. Reason explains why
   * the row is here. NetAmount follows Xero logic: debit positive,
   * credit negative.
   */
  Journals: ['Company', 'JournalID', 'JournalDate', 'JournalNumber',
             'SourceType', 'SourceID', 'Reference', 'Reason', 'JournalLineID',
             'AccountCode', 'AccountName', 'AccountType', 'AccountClass',
             'SystemAccount', 'Description',
             'NetAmount', 'GrossAmount', 'TaxAmount', 'TaxType',
             'Tracking', 'CreatedUTC'],

  /** Balance sheet by period: one row per account and period date. */
  Balance_Sheet: ['Company', 'PeriodDate', 'Section', 'AccountID', 'Account', 'Amount'],

  /** P&L from the Xero report: one row per account and period. */
  PnL_Report: ['Company', 'PeriodDate', 'Section', 'AccountID', 'Account', 'Amount'],

  Balances: ['Company', 'BalanceDate', 'Section', 'AccountID', 'Account',
             'Debit', 'Credit', 'Balance', 'YTD Debit', 'YTD Credit', 'YTD Balance'],

  Bank_Balances: ['Company', 'FromDate', 'ToDate', 'AccountID', 'BankAccount',
                  'OpeningBalance', 'CashReceived', 'CashSpent', 'ClosingBalance'],

  /** Shared aging sheet: AR and AP are distinguished by the Side column. */
  Aging: ['Company', 'BalanceDate', 'Side', 'Contact', 'ContactID', 'Number',
          'Date', 'DueDate', 'DaysOverdue', 'Bucket', 'Currency', 'AmountDue']
};

function spreadsheet_() {
  return CONFIG.spreadsheetId
    ? SpreadsheetApp.openById(CONFIG.spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) { return spreadsheet_().getSheetByName(SHEET_PREFIX + name); }

function resetSheet_(name) {
  const ss = spreadsheet_();
  const full = SHEET_PREFIX + name;
  const sh = ss.getSheetByName(full) || ss.insertSheet(full);
  dropFilter_(sh);
  sh.clear();
  // The sheet may have grown in a previous export: empty rows slow down
  // getLastRow and bloat the file
  if (sh.getMaxRows() > 2000) sh.deleteRows(2001, sh.getMaxRows() - 2000);
  sh.getRange(1, 1, 1, SCHEMA[name].length).setValues([SCHEMA[name]]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

/**
 * Removes a basic filter before a bulk write.
 *
 * Sorting through a filter physically reorders rows, and the filter range
 * does not expand when rows are appended, so overwriting such a
 * sheet corrupts the data. Filter views are not affected:
 * they are stored separately and left untouched.
 */
function dropFilter_(sh) {
  if (!CONFIG.dropFiltersBeforeWrite || !sh) return;
  try {
    const f = sh.getFilter();
    if (f) {
      f.remove();
      Log.info('Removed the filter from sheet %s: it blocks bulk writes.', sh.getName());
    }
  } catch (err) {
    // Not critical: the write will proceed anyway
    Log.error('Could not remove the filter from %s: %s', sh.getName(), err.message);
  }
}

/**
 * Make sure the sheet has enough rows.
 * A new sheet has 1000 rows, and setValues beyond the grid
 * does not expand it but throws an error. Add a buffer to avoid calling
 * insertRowsAfter on every page.
 */
function ensureRows_(sh, needed) {
  const max = sh.getMaxRows();
  if (needed > max) sh.insertRowsAfter(max, needed - max + 500);
}

function appendRows_(name, rows) {
  if (!rows || !rows.length) return;

  // In incremental mode rows are collected in memory: they still
  // have to be merged with already exported rows by ID
  if (WRITE_BUFFER) {
    if (!WRITE_BUFFER[name]) WRITE_BUFFER[name] = [];
    WRITE_BUFFER[name] = WRITE_BUFFER[name].concat(rows);
    return;
  }

  const ss = spreadsheet_();
  const full = SHEET_PREFIX + name;
  let sh = ss.getSheetByName(full);
  // The prefix is needed here too: resetSheet_ already creates the sheet with it
  if (!sh) { resetSheet_(name); sh = ss.getSheetByName(full); }

  const start = sh.getLastRow() + 1;
  if (start <= 2) dropFilter_(sh);      // on the first page of the job
  ensureRows_(sh, start + rows.length);
  sh.getRange(start, 1, rows.length, SCHEMA[name].length).setValues(rows);
}

/**
 * Writes a whole report sheet.
 *
 * @param {boolean} [filter] create a filter on the finished data.
 *   Set only on sheets meant for reading: a basic filter
 *   physically reorders rows when sorting, so before
 *   the next write dropFilter_ always removes it.
 */
function writeReport_(name, headers, rows, frozenCols, filter) {
  const ss = spreadsheet_();
  const full = SHEET_PREFIX + name;
  const sh = ss.getSheetByName(full) || ss.insertSheet(full);

  dropFilter_(sh);
  sh.clear();

  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (frozenCols) sh.setFrozenColumns(frozenCols);

  if (rows.length) {
    ensureRows_(sh, rows.length + 1);
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  if (filter) sh.getDataRange().createFilter();
}

// ================================================================
//  17. UTILITIES
// ================================================================

/** Xero reports come in nested sections; they are flattened into plain objects. */
function flattenReport_(payload) {
  const report = (payload.Reports || [])[0];
  if (!report) return [];

  let headers = [];
  const out = [];

  (report.Rows || []).forEach(function (row) {
    if (row.RowType === 'Header') {
      headers = (row.Cells || []).map(function (c) { return String(c.Value || '').trim(); });
      return;
    }
    const section = row.Title || '';
    (row.Rows || []).forEach(function (r) {
      if (r.RowType !== 'Row' && r.RowType !== 'SummaryRow') return;
      const obj = { section: section };
      (r.Cells || []).forEach(function (cell, i) {
        obj[headers[i] || ('col' + i)] = cell.Value;
        (cell.Attributes || []).forEach(function (a) {
          if (a.Id === 'account' || a.Id === 'accountID') obj.accountId = a.Value;
        });
      });
      obj.accountName = obj[headers[0]] !== undefined ? obj[headers[0]] : '';
      out.push(obj);
    });
  });

  return out;
}

function daysOverdue_(due, balanceDate) {
  if (!due) return 0;
  return Math.floor((new Date(balanceDate + 'T00:00:00Z') - due) / 86400000);
}

function bucket_(days) {
  const b = CONFIG.agingBuckets;
  if (days <= 0)    return 'Current';
  if (days <= b[0]) return '1-' + b[0];
  if (days <= b[1]) return (b[0] + 1) + '-' + b[1];
  if (days <= b[2]) return (b[1] + 1) + '-' + b[2];
  return b[2] + '+';
}

function balanceDate_() {
  return CONFIG.balanceDate || Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
}

/**
 * End date for reports with periods.
 *
 * Xero counts periods BACKWARDS FROM toDate: with toDate = 20 August
 * the periods are 21 July to 20 August and so on. Document totals
 * are grouped by calendar month, so the matching would be off.
 * Use the end of the month so periods match calendar months.
 */
function reportEndDate_() {
  const base = new Date(balanceDate_() + 'T00:00:00Z');
  const monthEnd = new Date(Date.UTC(base.getUTCFullYear(), base.getUTCMonth() + 1, 0));
  return Utilities.formatDate(monthEnd, 'UTC', 'yyyy-MM-dd');
}

/** '/Date(1735689600000+0000)/' or 'YYYY-MM-DD' -> Date */
function parseDate_(v) {
  if (!v) return '';
  const m = /\/Date\((-?\d+)/.exec(String(v));
  if (m) return new Date(Number(m[1]));
  const d = new Date(v);
  return isNaN(d.getTime()) ? '' : d;
}

/**
 * Period label from the Xero report -> last day of that month.
 * Xero returns them differently depending on organisation settings:
 * 31 Mar 2025, Mar 2025, Mar-25. All three are parsed.
 */
const MONTH_NAMES = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5,
                      jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };

function parsePeriodEnd_(label) {
  const text = String(label || '').trim();
  if (!text) return '';

  const m = /([A-Za-z]{3})[a-z]*[\s-]+(\d{2,4})/.exec(text);
  if (m) {
    const month = MONTH_NAMES[m[1].toLowerCase()];
    if (month === undefined) return '';
    let year = Number(m[2]);
    if (year < 100) year += 2000;
    return new Date(Date.UTC(year, month + 1, 0));    // day zero = end of month
  }

  const d = new Date(text);
  return isNaN(d.getTime()) ? '' : d;
}

/** Date -> DateTime(2025,01,31,14,05,00) for the where parameter */
function xeroDateTime_(d) {
  return 'DateTime(' + d.getUTCFullYear() + ',' + (d.getUTCMonth() + 1) + ',' +
         d.getUTCDate() + ',' + d.getUTCHours() + ',' + d.getUTCMinutes() + ',' +
         d.getUTCSeconds() + ')';
}

/**
 * Entities that can be filtered by modification date.
 *
 * BankTransfer has no UpdatedDateUTC field in queries at all: Xero
 * responds with QueryParseException. Prepayments and overpayments are
 * not included either: their behaviour is inconsistent, and an error costs
 * a failed job in the middle of a run.
 */
const SUPPORTS_UPDATED_FILTER = {
  Invoices: 1, CreditNotes: 1, BankTransactions: 1, Payments: 1
};

/**
 * Filter condition for where.
 *
 * Regular mode: by document date. Incremental: by modification
 * date, so old documents that were edited are included too.
 * If the entity does not support this filter, a transaction date window is used
 * and changedSince_() drops the rest.
 *
 * @param {string} entity  Invoices, BankTransfers and so on
 */
function dateFilter_(entity) {
  if (!SYNC_SINCE) return 'Date>=' + xeroDate_(CONFIG.fromDate);

  if (SUPPORTS_UPDATED_FILTER[entity]) {
    return 'UpdatedDateUTC>=' + xeroDateTime_(SYNC_SINCE);
  }

  // Transaction date window, but not earlier than the export start
  const lookback = new Date(SYNC_SINCE.getTime() -
                            CONFIG.incrementalLookbackDays * 86400000);
  const from = new Date(CONFIG.fromDate + 'T00:00:00Z');
  const start = lookback > from ? lookback : from;
  return 'Date>=' + xeroDate_(Utilities.formatDate(start, 'UTC', 'yyyy-MM-dd'));
}

/**
 * Whether the record changed after the last sync.
 * Needed where a modification date filter cannot be sent in the request:
 * Xero returns the whole window and the rest is dropped here.
 */
function changedSince_(item) {
  if (!SYNC_SINCE) return true;
  const updated = parseDate_(item.UpdatedDateUTC);
  return !(updated instanceof Date) || updated >= SYNC_SINCE;
}

/** 'YYYY-MM-DD' -> DateTime(2025,01,01) for the where parameter */
function xeroDate_(iso) {
  const p = iso.split('-');
  return 'DateTime(' + Number(p[0]) + ',' + Number(p[1]) + ',' + Number(p[2]) + ')';
}

function num_(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function round2_(v) { return Math.round(num_(v) * 100) / 100; }

// ================================================================
//  18. TARGETED RUNS
// ================================================================

/** How many Xero requests are left today per entity. */
function showApiQuota() {
  const service = getXeroService_();
  const tenants = readXeroTenants_();

  Object.keys(XERO_COMPANIES).forEach(function (alias) {
    const t = tenants[alias];
    if (!t) { Log.warn('%s: not linked', alias); return; }

    const res = UrlFetchApp.fetch('https://api.xero.com/api.xro/2.0/Organisation', {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken(),
        'xero-tenant-id': t.tenantId,
        Accept: 'application/json'
      },
      muteHttpExceptions: true
    });

    const h = res.getHeaders();
    Log.info('%s: %s of 5000 left today, %s of 60 this minute',
      alias, h['X-DayLimit-Remaining'] || '?', h['X-MinLimit-Remaining'] || '?');
  });
}

/** P&L report only, for all entities. */
function exportPnlReportOnly() {
  const date = balanceDate_();
  resetSheet_('PnL_Report');
  Object.keys(XERO_COMPANIES).forEach(function (c) { exportPnlReport_(c, date); });
  buildPnl();
  Log.info('P&L report as of %s exported, PnL sheet rebuilt.', date);
}

/** Balance sheet only, for all entities. */
function exportBalanceSheetOnly() {
  const date = balanceDate_();
  resetSheet_('Balance_Sheet');
  Object.keys(XERO_COMPANIES).forEach(function (c) { exportBalanceSheet_(c, date); });
  Log.info('Balance sheet as of %s exported.', date);
}

function exportBalancesOnly() {
  const date = balanceDate_();
  resetSheet_('Balances');
  resetSheet_('Bank_Balances');
  Object.keys(XERO_COMPANIES).forEach(function (c) { exportBalances_(c, date); });
}

function exportBankOnly() {
  resetSheet_('BankTransactions');
  if (CONFIG.includePayments) resetSheet_('Payments');

  const seen = {};
  Object.keys(XERO_COMPANIES).forEach(function (company) {
    const st = { page: 1 };
    while (!exportBankTx_(company, st, seen)) {}
    if (CONFIG.includeBankTransfers) exportTransfers_(company, seen);
    if (CONFIG.includePayments) {
      const sp = { page: 1 };
      while (!exportPayments_(company, sp, seen)) {}
    }
  });

  if (CONFIG.mapToPnl) { tagWithPnl(); buildCashflow(); }
  countUnsplitPayments();
}

// ================================================================
//  19. THREE RUN MODES
// ================================================================

/** Jobs of one entity: a shared list for all modes. */
function jobsFor_(company) {
  const jobs = [{ company: company, task: 'balances' },
                { company: company, task: 'invoices' }];
  if (CONFIG.includeCreditNotes) jobs.push({ company: company, task: 'creditNotes' });
  if (CONFIG.includePrepayments) {
    jobs.push({ company: company, task: 'prepayments' });
    jobs.push({ company: company, task: 'overpayments' });
  }
  jobs.push({ company: company, task: 'bankTransactions' });
  if (CONFIG.includeBankTransfers) jobs.push({ company: company, task: 'bankTransfers' });
  if (CONFIG.includePayments) jobs.push({ company: company, task: 'payments' });
  if (CONFIG.includeJournals) jobs.push({ company: company, task: 'journals' });
  if (CONFIG.includeBalanceSheet) jobs.push({ company: company, task: 'balanceSheet' });
  if (CONFIG.includePnlAdjustments) jobs.push({ company: company, task: 'pnlReport' });
  return jobs;
}

/** Runs one job to completion, without pauses or triggers. */
function runJobToEnd_(job, balanceDateValue, seen) {
  const state = { page: 1, offset: 0, balanceDate: balanceDateValue };
  let guard = 0;

  while (guard++ < CONFIG.maxPages) {
    let done;
    switch (job.task) {
      case 'balances':         done = exportBalances_(job.company, state.balanceDate); break;
      case 'invoices':         done = exportInvoices_(job.company, state, seen);       break;
      case 'creditNotes':      done = exportCreditNotes_(job.company, state, seen);    break;
      case 'prepayments':      done = exportPrepay_(job.company, state, seen, 'Prepayments');  break;
      case 'overpayments':     done = exportPrepay_(job.company, state, seen, 'Overpayments'); break;
      case 'bankTransactions': done = exportBankTx_(job.company, state, seen);         break;
      case 'bankTransfers':    done = exportTransfers_(job.company, seen);             break;
      case 'payments':         done = exportPayments_(job.company, state, seen);       break;
      case 'journals':         done = exportJournals_(job.company, state, seen);       break;
      case 'balanceSheet':     done = exportBalanceSheet_(job.company, state.balanceDate); break;
      case 'pnlReport':        done = exportPnlReport_(job.company, state.balanceDate); break;
      default: throw new Error('Unknown job: ' + job.task);
    }
    if (done) return;
  }
  Log.info('%s / %s: hit maxPages', job.company, job.task);
}

// ---------- 1. Test run ----------

/**
 * Checks the whole chain on a short period.
 *
 * Writes to sheets prefixed TEST_; production data is not touched.
 * Runs synchronously without triggers or pauses, so the period must be
 * short: a few days, otherwise it hits the six-minute execution limit.
 *
 *   testExport()             last 7 days, first entity
 *   testExport(14)           last 14 days
 *   testExport(7, 'entity_uk')   specific entity
 *
 * What it checks: availability of all endpoints and scopes, response parsing,
 * payment allocation by line, payment split by line item, P&L build
 * and balance sheet. The log shows how many rows each job returned.
 */
function testExport(days, company) {
  const period = days || 7;
  const target = company || Object.keys(XERO_COMPANIES)[0];
  if (!XERO_COMPANIES[target]) {
    Log.warn('Unknown entity "%s". Available: %s', target, Object.keys(XERO_COMPANIES).join(', '));
    return;
  }

  const savedPrefix = SHEET_PREFIX;
  const savedFrom = CONFIG.fromDate;
  const savedSync = SYNC_SINCE;

  const from = new Date(Date.now() - period * 86400000);
  const started = Date.now();

  try {
    SHEET_PREFIX = 'TEST_';
    SYNC_SINCE = null;
    CONFIG.fromDate = Utilities.formatDate(from, 'UTC', 'yyyy-MM-dd');
    resetCategoryIndex_();

    Log.info('=== TEST: %s, period from %s ===', target, CONFIG.fromDate);
    Object.keys(SCHEMA).forEach(resetSheet_);

    const seen = {};
    jobsFor_(target).forEach(function (job) {
      const before = countSheetTotal_();
      try {
        runJobToEnd_(job, balanceDate_(), seen);
        Log.info('  %-16s +%s rows', job.task, countSheetTotal_() - before);
      } catch (err) {
        Log.error('  %-16s ERROR: %s', job.task, err.message);
      }
    });

    buildAllReports();
    Log.info('=== Done in %s s. See the TEST_* sheets ===',
      Math.round((Date.now() - started) / 1000));

  } finally {
    SHEET_PREFIX = savedPrefix;
    CONFIG.fromDate = savedFrom;
    SYNC_SINCE = savedSync;
    resetCategoryIndex_();
  }
}

/** Delete the test sheets. */
function clearTestSheets() {
  const ss = spreadsheet_();
  let removed = 0;
  ss.getSheets().forEach(function (sh) {
    if (sh.getName().indexOf('TEST_') === 0) { ss.deleteSheet(sh); removed++; }
  });
  Log.info('Test sheets deleted: %s', removed);
}

// ---------- 3. Incremental update ----------

/**
 * Key by which a row belongs to a Xero record.
 * Payments and transfers have an ID with a suffix (:1, :out);
 * the part before the colon is used to replace the whole record.
 */
function ownerKey_(sheetName, row) {
  const schema = SCHEMA[sheetName];
  function at(col) {
    const i = schema.indexOf(col);
    return i < 0 ? '' : String(row[i] || '');
  }

  switch (sheetName) {
    case 'Documents':        return at('DocumentID');
    case 'Lines':            return at('DocumentID');
    case 'BankTransactions': return at('TransactionID').split(':')[0];
    case 'Payments':         return at('PaymentID');
    case 'Aging':            return at('Company') + '|' + at('ContactID') + '|' + at('Number');
    default: return '';
  }
}

/**
 * Replaces all rows in the sheet that belong to updated records
 * and appends new ones. The sheet is rewritten in one operation:
 * faster than deleting rows one by one.
 */
function upsertSheet_(sheetName, newRows, filter) {
  if (!newRows || !newRows.length) return { updated: 0, added: 0 };

  const schema = SCHEMA[sheetName];
  const sh = sheet_(sheetName);
  if (!sh) { resetSheet_(sheetName); appendRows_(sheetName, newRows); return { updated: 0, added: newRows.length }; }

  const touched = {};
  newRows.forEach(function (r) {
    const key = ownerKey_(sheetName, r);
    if (key) touched[key] = 1;
  });

  dropFilter_(sh);

  const existing = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, schema.length).getValues()
    : [];

  let updated = 0;
  const kept = existing.filter(function (r) {
    const key = ownerKey_(sheetName, r);
    if (key && touched[key]) { updated++; return false; }
    return true;
  });

  const merged = kept.concat(newRows);
  ensureRows_(sh, merged.length + 1);
  sh.getRange(2, 1, Math.max(existing.length, merged.length), schema.length).clearContent();
  if (merged.length) sh.getRange(2, 1, merged.length, schema.length).setValues(merged);

  if (filter) sh.getDataRange().createFilter();

  return { updated: updated, added: newRows.length };
}

/**
 * Daily update: only what changed since the last run.
 *
 * UpdatedDateUTC is used instead of the document date filter,
 * so old documents that were edited or paid are included too.
 * Rows are updated in place by ID, sheets are not cleared.
 *
 * Snapshots (balances, balance sheet, P&L report) are re-exported in full:
 * a few requests per entity and always an up-to-date picture.
 *
 * Cost: about 10-30 requests per entity versus several thousand
 * for a full reload.
 */
function incrementalExport() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { Log.info('Another run is in progress.'); return; }

  let completed = false;

  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(PROP_STATE)) {
      Log.warn('A full export is running; skipping the incremental update.');
      return;
    }

    const syncRaw = props.getProperty(PROP_SYNC);
    const sync = syncRaw ? JSON.parse(syncRaw) : {};
    const startedAt = new Date();

    // Snapshots are re-exported in full, so they are cleared once for all entities
    ['Balances', 'Bank_Balances', 'Balance_Sheet', 'PnL_Report'].forEach(resetSheet_);

    Object.keys(XERO_COMPANIES).forEach(function (company) {
      const since = sync[company] ? new Date(sync[company]) : null;
      if (!since) {
        Log.info('%s: no sync timestamp. Run a full export first.', company);
        return;
      }

      // Small overlap backwards: smooths out clock differences
      SYNC_SINCE = new Date(since.getTime() - 60 * 60 * 1000);
      WRITE_BUFFER = {};
      resetCategoryIndex_();

      const seen = {};
      let failed = false;

      jobsFor_(company).forEach(function (job) {
        try {
          runJobToEnd_(job, balanceDate_(), seen);
        } catch (err) {
          failed = true;
          Log.error('%s / %s: error: %s', company, job.task, err.message);
        }
      });

      const buffer = WRITE_BUFFER;
      WRITE_BUFFER = null;
      SYNC_SINCE = null;

      // Snapshots are written as is, changing data is merged
      ['Balances', 'Bank_Balances', 'Balance_Sheet', 'PnL_Report'].forEach(function (name) {
        if (buffer[name]) appendRows_(name, buffer[name]);
      });

      ['Documents', 'Lines', 'BankTransactions', 'Payments', 'Aging'].forEach(function (name) {
        if (!buffer[name]) return;
        const res = upsertSheet_(name, buffer[name], true);
        Log.info('%s / %s: replaced %s, added %s', company, name, res.updated, res.added);
      });

      if (!failed) sync[company] = startedAt.toISOString();
    });

    props.setProperty(PROP_SYNC, JSON.stringify(sync));
    buildAllReports();
    completed = true;
    Log.info('Incremental update completed.');

  } finally {
    WRITE_BUFFER = null;
    SYNC_SINCE = null;
    try { lock.releaseLock(); } catch (err) { /* already released */ }
  }

  // Attachments after releasing the lock: syncAttachments acquires it itself
  if (completed && CONFIG.attachmentsInIncremental) syncAttachments();
}

/**
 * Mark the full export as current as of now.
 * Runs automatically when startExport finishes, but may be
 * needed manually, e.g. after manual edits of the sheets.
 */
function markSynced() {
  const sync = {};
  const now = new Date().toISOString();
  Object.keys(XERO_COMPANIES).forEach(function (c) { sync[c] = now; });
  PropertiesService.getScriptProperties().setProperty(PROP_SYNC, JSON.stringify(sync));
  Log.info('Sync timestamp set to %s', now);
}

/** When each entity was last synced. */
function showSyncStatus() {
  const raw = PropertiesService.getScriptProperties().getProperty(PROP_SYNC);
  if (!raw) { Log.info('No sync yet.'); return; }
  const sync = JSON.parse(raw);
  Object.keys(XERO_COMPANIES).forEach(function (c) {
    if (!sync[c]) { Log.info('%s: not synced', c); return; }
    const hours = Math.round((Date.now() - new Date(sync[c]).getTime()) / 360000) / 10;
    Log.info('%s: %s (%s h ago)', c, sync[c], hours);
  });
}

// ================================================================
//  20. ATTACHMENTS ON GOOGLE DRIVE
// ================================================================

/**
 * Entity folder on Drive. Created on first access.
 * If attachmentsFolderId is empty, the root folder Xero Attachments
 * is created in My Drive.
 */
function attachmentFolder_(company) {
  let root;
  if (CONFIG.attachmentsFolderId) {
    root = DriveApp.getFolderById(CONFIG.attachmentsFolderId);
  } else {
    const found = DriveApp.getFoldersByName('Xero Attachments');
    root = found.hasNext() ? found.next() : DriveApp.createFolder('Xero Attachments');
  }

  const sub = root.getFoldersByName(company);
  return sub.hasNext() ? sub.next() : root.createFolder(company);
}

/**
 * Downloads an attachment file.
 *
 * Does not go through xeroFetch: it parses responses as JSON, while here
 * binary data arrives. The Accept header must carry the MIME type
 * of the file itself, otherwise Xero returns metadata instead of content.
 */
function fetchAttachmentBlob_(company, endpoint, docId, att) {
  const tenant = readXeroTenants_()[company];
  if (!tenant) throw new Error('Entity "' + company + '" is not linked.');

  const url = 'https://api.xero.com/api.xro/2.0/' + endpoint + '/' + docId +
              '/Attachments/' + encodeURIComponent(att.FileName);

  const res = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + getXeroService_().getAccessToken(),
      'xero-tenant-id': tenant.tenantId,
      Accept: att.MimeType || 'application/octet-stream'
    },
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code === 429) {
    const err = new Error('QUOTA_EXHAUSTED while downloading an attachment.');
    err.quotaExhausted = true;
    throw err;
  }
  if (code !== 200) throw new Error('Attachment ' + att.FileName + ': HTTP ' + code);

  return res.getBlob();
}

/**
 * Uploads document attachments to Drive and writes the links
 * to the Attachments column of the Documents sheet.
 *
 * Processes only rows where HasAttachments = true and no links exist yet,
 * so a rerun does not download saved files again but picks up
 * what was not reached last time. A run is limited by
 * attachmentsPerRun documents and execution time; the rest
 * is taken by the next run, for which a trigger is set.
 *
 *   syncAttachments()            all entities, limit from CONFIG
 *   syncAttachments('entity_uk')     one entity
 *   syncAttachments(null, 50)    custom limit
 */
function syncAttachments(company, limit) {
  // A trigger passes an event object as the first argument, not a string.
  // Without this check the entity filter would reject all rows:
  // the run would do nothing, processing zero documents.
  if (company && typeof company !== 'string') company = null;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) { Log.info('Another run is in progress.'); return; }

  try {
    const sh = sheet_('Documents');
    if (!sh || sh.getLastRow() < 2) { Log.warn('The Documents sheet is empty.'); return; }

    const ds = SCHEMA.Documents;
    const i = {
      company: ds.indexOf('Company'), docType: ds.indexOf('DocType'),
      id: ds.indexOf('DocumentID'), number: ds.indexOf('Number'),
      contact: ds.indexOf('Contact'), date: ds.indexOf('Date'),
      has: ds.indexOf('HasAttachments'), links: ds.indexOf('Attachments')
    };

    if (sh.getLastColumn() < ds.length) {
      Log.info('The Documents sheet has no HasAttachments and Attachments columns: ' +
                 'a full re-export on this version is required.');
      return;
    }

    const rowCount = sh.getLastRow() - 1;
    const data = sh.getRange(2, 1, rowCount, ds.length).getValues();
    const linkCol = data.map(function (r) { return [r[i.links]]; });

    const cap = limit || CONFIG.attachmentsPerRun;
    const started = Date.now();
    const folders = {};
    const attachRows = [];

    let processed = 0, files = 0, skipped = 0, failed = 0, paused = false;
    let pauseReason = '';

    for (var idx = 0; idx < data.length; idx++) {
      if (processed >= cap) { paused = true; pauseReason = 'limit'; break; }
      if (Date.now() - started > CONFIG.attachmentsRuntimeMs) {
        paused = true; pauseReason = 'time'; break;
      }

      const r = data[idx];
      if (!r[i.company]) continue;
      if (company && r[i.company] !== company) continue;
      if (!CONFIG.attachmentDocTypes[r[i.docType]]) continue;
      if (r[i.has] !== true && String(r[i.has]).toUpperCase() !== 'TRUE') continue;
      if (String(r[i.links] || '').trim()) continue;      // already uploaded

      processed++;
      const docCompany = r[i.company];
      const docId = r[i.id];

      // Credit notes are in a different endpoint
      const endpoint = /Credit Note/.test(r[i.docType]) ? 'CreditNotes' : 'Invoices';

      let list;
      try {
        list = (xeroFetch(docCompany, endpoint + '/' + docId + '/Attachments')
                || {}).Attachments || [];
      } catch (err) {
        if (err.quotaExhausted) { paused = true; pauseReason = 'quota'; break; }
        Log.info('%s: attachment list not received: %s', r[i.number], err.message);
        failed++;
        continue;
      }

      if (!list.length) { linkCol[idx][0] = '—'; continue; }

      if (!folders[docCompany]) folders[docCompany] = attachmentFolder_(docCompany);
      const folder = folders[docCompany];
      const urls = [];

      for (var a = 0; a < list.length; a++) {
        const att = list[a];
        const sizeKb = Math.round(num_(att.ContentLength) / 1024);

        if (sizeKb > CONFIG.attachmentMaxSizeMb * 1024) {
          Log.warn('%s / %s: skipped, %s MB', r[i.number], att.FileName,
                     Math.round(sizeKb / 1024));
          skipped++;
          continue;
        }

        try {
          const blob = fetchAttachmentBlob_(docCompany, endpoint, docId, att);
          // Name with the document number: files from Xero are often called scan.pdf
          const name = (r[i.number] ? r[i.number] + ' - ' : '') + att.FileName;
          const file = folder.createFile(blob.setName(name));

          urls.push(file.getUrl());
          files++;

          attachRows.push([docCompany, r[i.docType], docId, r[i.number] || '',
                           r[i.contact] || '', r[i.date],
                           att.FileName, att.MimeType || '', sizeKb,
                           file.getId(), file.getUrl(), new Date()]);
        } catch (err) {
          if (err.quotaExhausted) { paused = true; pauseReason = 'quota'; break; }
          Log.info('%s / %s: %s', r[i.number], att.FileName, err.message);
          failed++;
        }
      }

      if (urls.length) linkCol[idx][0] = urls.join('\n');
      if (paused) break;

      // Intermediate save: if the execution is cut off by time
      // or hits the Xero limit, files already downloaded will not be downloaded again
      if (processed % 25 === 0) {
        dropFilter_(sh);
        sh.getRange(2, i.links + 1, rowCount, 1).setValues(linkCol);
        if (attachRows.length) {
          appendRows_('Attachments', attachRows);
          attachRows.length = 0;
        }
      }
    }

    // Links are written in one operation
    dropFilter_(sh);
    sh.getRange(2, i.links + 1, rowCount, 1).setValues(linkCol);
    if (attachRows.length) appendRows_('Attachments', attachRows);

    const reasonText = {
      quota: 'XERO DAILY QUOTA EXHAUSTED',
      time: 'execution time ran out',
      limit: 'attachmentsPerRun limit reached'
    }[pauseReason] || 'reason unknown';

    Log.info('Attachments: documents processed %s, files saved %s, ' +
               'skipped by size %s, errors %s%s',
      processed, files, skipped, failed,
      paused ? '  <- pause: ' + reasonText : '');

    if (paused && pauseReason === 'quota') {
      // The quota will not return in a minute: the window is rolling, 24 hours.
      // Retrying every minute would only waste runs.
      Log.warn('Remaining daily quota: %s. Resuming in %s h.',
        (typeof LAST_QUOTA !== 'undefined' && LAST_QUOTA.day !== null
          ? LAST_QUOTA.day : 'unknown'),
        CONFIG.resumeAfterQuotaHours);
      scheduleAttachments_(CONFIG.resumeAfterQuotaHours * 60);
    } else if (paused) {
      scheduleAttachments_(1);
    } else {
      stopAttachments_();
      Log.info('All attachments uploaded, trigger removed.');
    }

  } finally {
    // The lock may have been released by itself during a long run;
    // releasing it again throws IllegalStateException and fails
    // the execution after everything has been saved
    try { lock.releaseLock(); } catch (err) { /* already released */ }
  }
}

/**
 * Continue uploading attachments after the given number of minutes.
 *
 * After hitting the daily quota it is set to hours, not a minute:
 * otherwise the trigger would run idle until the window recovers.
 * The trigger calls syncAttachments with an event object as the first
 * argument; the function ignores it, see the check at its start.
 */
function scheduleAttachments_(minutes) {
  const delay = Math.round((minutes || 1) * 60 * 1000);
  try {
    stopAttachments_();
    ScriptApp.newTrigger('syncAttachments').timeBased().after(delay).create();
  } catch (err) {
    Log.error('Could not set the resume trigger: %s', err.message);
    Log.info('Run syncAttachments() manually; it continues from the same point.');
  }
}

function stopAttachments_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncAttachments') ScriptApp.deleteTrigger(t);
  });
}

/** How many documents are still waiting for attachment upload. */
function attachmentsPending() {
  const sh = sheet_('Documents');
  if (!sh || sh.getLastRow() < 2) { Log.warn('The Documents sheet is empty.'); return; }

  const ds = SCHEMA.Documents;
  if (sh.getLastColumn() < ds.length) {
    Log.info('No HasAttachments and Attachments columns: a re-export is required.');
    return;
  }

  const i = {
    company: ds.indexOf('Company'), docType: ds.indexOf('DocType'),
    has: ds.indexOf('HasAttachments'), links: ds.indexOf('Attachments')
  };

  const stat = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, ds.length).getValues().forEach(function (r) {
    if (!r[i.company] || !CONFIG.attachmentDocTypes[r[i.docType]]) return;
    const withAtt = r[i.has] === true || String(r[i.has]).toUpperCase() === 'TRUE';
    if (!withAtt) return;

    if (!stat[r[i.company]]) stat[r[i.company]] = { done: 0, pending: 0 };
    if (String(r[i.links] || '').trim()) stat[r[i.company]].done++;
    else stat[r[i.company]].pending++;
  });

  Object.keys(stat).forEach(function (c) {
    Log.info('%s - uploaded %s, waiting %s', c, stat[c].done, stat[c].pending);
  });
  if (!Object.keys(stat).length) Log.info('No documents with attachments found.');
}

/**
 * Adds two columns introduced in the new version to an existing Documents
 * sheet: HasAttachments and Attachments, WITHOUT a full re-export.
 *
 * The attachment flag is taken from Xero in one cheap pass:
 * with summaryOnly=true a request returns up to 1000 documents, so
 * an entity takes a few calls instead of thousands. Matching
 * is done by DocumentID; amounts and other fields are not touched.
 *
 * Run once after updating the script. After that the export
 * sets the flag itself.
 */
function backfillAttachmentFlags() {
  const sh = sheet_('Documents');
  if (!sh || sh.getLastRow() < 2) { Log.warn('The Documents sheet is empty.'); return; }

  const ds = SCHEMA.Documents;
  const iHas = ds.indexOf('HasAttachments');
  const iLinks = ds.indexOf('Attachments');
  const iId = ds.indexOf('DocumentID');

  // 1. Headers: appended if the sheet is from an older version
  if (sh.getLastColumn() < ds.length) {
    sh.getRange(1, iHas + 1, 1, 2)
      .setValues([[ds[iHas], ds[iLinks]]])
      .setFontWeight('bold');
    Log.info('Added the HasAttachments and Attachments columns.');
  }

  // 2. Ask Xero for documents WITH attachments ONLY.
  //
  //    The HasAttachments==true filter is more reliable than reading the field:
  //    with summaryOnly=true Xero returns a light object without heavy
  //    fields and the flag is missing, which is why the previous version
  //    marked almost everything as without attachments. The result set is also smaller:
  //    only the needed documents arrive.
  const withAttachments = {};
  let total = 0;

  Object.keys(XERO_COMPANIES).forEach(function (company) {
    ['Invoices', 'CreditNotes'].forEach(function (endpoint) {
      const idField = endpoint === 'Invoices' ? 'InvoiceID' : 'CreditNoteID';
      let page = 1, got = 0;

      while (page <= CONFIG.maxPages) {
        let res;
        try {
          res = xeroFetch(company, endpoint + '?' + [
            'where=' + encodeURIComponent('HasAttachments==true AND Date>=' +
                                          xeroDate_(CONFIG.fromDate)),
            'order=Date',
            'page=' + page,
            'pageSize=' + CONFIG.pageSize
          ].join('&'));
        } catch (err) {
          Log.info('%s / %s page %s: %s', company, endpoint, page, err.message);
          if (/401|Unauthorized/i.test(err.message)) {
            Log.warn('<- scope accounting.attachments.read seems to be missing');
          }
          break;
        }

        const items = res[endpoint] || [];
        items.forEach(function (d) { withAttachments[String(d[idField])] = 1; });
        got += items.length;

        const p = res.pagination;
        if (p && p.pageCount) { if (page >= p.pageCount) break; }
        else if (items.length < CONFIG.pageSize) break;
        page++;
      }

      total += got;
      Log.info('%s / %s: with attachments %s', company, endpoint, got);
    });
  });

  // 3. Set the flag: a document in the result set has attachments
  const rowCount = sh.getLastRow() - 1;
  const data = sh.getRange(2, 1, rowCount, ds.length).getValues();
  const hasCol = [];
  let marked = 0;

  data.forEach(function (r) {
    const id = String(r[iId] || '');
    if (!id) { hasCol.push(['']); return; }
    const has = withAttachments[id] === 1;
    if (has) marked++;
    hasCol.push([has]);
  });

  dropFilter_(sh);
  sh.getRange(2, iHas + 1, rowCount, 1).setValues(hasCol);

  Log.info('Documents with attachments in Xero: %s, flagged in the sheet: %s of %s rows.',
    total, marked, rowCount);
  if (total > marked) {
    Log.info('Difference %s: documents outside the fromDate period or of another type: ' +
               'they are simply not in the sheet.', total - marked);
  }
  Log.info('Next: syncAttachments() uploads the files to Drive.');
}