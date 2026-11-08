/**
 * ClickUp client lists from templates.
 *
 * When a new client is signed, Zapier calls this web app with the client name,
 * the business unit and the ID of the related ClickUp task. The script:
 *   1. checks whether the client already exists in a ClickUp registry List
 *   2. picks the List templates for "<unit>_<exists>" from CLICKUP_LIST_MAPPING
 *      (one request can create several Lists, e.g. RACI and Pipeline)
 *   3. creates each List from its ClickUp template in the target Folder or Space
 *   4. writes a description with a link to the source task
 *   5. records every created List in a spreadsheet and returns the List IDs
 *
 * Request (single):  { "secret": "...", "client": "Client X", "unit": "Unit A", "linkedTask": "869abc123" }
 * Request (batch):   { "secret": "...", "jobs": [ { client, unit, linkedTask }, ... ] }
 * Response:          { ok, batchId, created, failed, results, rejected, deferred }
 */

const CLICKUP_BASE_URL = 'https://api.clickup.com/api/v2';
const TIME_ZONE = 'Europe/Berlin';
const TIME_GUARD_MS = 25 * 1000;      // answer before Zapier's 30 s webhook timeout
const MAX_ATTEMPTS = 3;
const RATE_LIMIT_WAIT_MS = 8000;
const DESCRIPTION_RETRY_WAIT_MS = 3000;
const LOCK_TIMEOUT_MS = 20000;

/** Written into the description of every created List ('' to skip). A spec may override it. */
const DESCRIPTION_TEMPLATE = 'Client: {client}\nSource task: {linkedTaskUrl}\nID: {linkedTask}\nCreated: {date}';
const APPEND_DESCRIPTION = true;      // true: add below the template's description, false: replace it

const DEFAULT_REGISTRY = { listId: '', matchBy: 'name', fieldId: '', includeClosed: true, cacheSeconds: 120 };
const DEFAULT_RECORDS = { tab: 'Subscriptions', columns: ['listId', 'client', 'linkedTask', 'unit'] };
const DEFAULT_RECORDS_FIRST_COLUMN = 4;   // column D

// ============================ Web app ============================

function doPost(e) {
  const start = Date.now();
  try {
    const request = JSON.parse(e.postData.contents);
    if (!request || request.secret !== getRequiredProperty_('WEBHOOK_SECRET')) {
      Log.warn('Rejected request with invalid secret');
      return jsonOutput_({ ok: false, error: 'unauthorized' });
    }

    const jobs = Array.isArray(request.jobs) ? request.jobs : [request];
    const mapping = getJsonProperty_('CLICKUP_LIST_MAPPING', {});
    const batchId = Utilities.getUuid().slice(0, 8);
    const summary = { results: [], rejected: [], deferred: [], created: 0, failed: 0 };

    jobs.forEach(function (job, index) {
      if (!job.client || !job.unit || !job.linkedTask) {
        summary.rejected.push({ index: index, error: 'client, unit and linkedTask are required' });
        return;
      }
      if (Date.now() - start > TIME_GUARD_MS) {
        summary.deferred.push({ index: index, client: job.client, error: 'time guard, send again' });
        return;
      }
      processJob_(job, index, mapping, batchId, summary);
    });

    Log.info('Batch finished', {
      batchId: batchId, jobs: jobs.length, created: summary.created, failed: summary.failed,
      rejected: summary.rejected.length, deferred: summary.deferred.length, ms: Date.now() - start
    });
    return jsonOutput_({
      ok: !summary.failed && !summary.rejected.length && !summary.deferred.length,
      batchId: batchId,
      created: summary.created,
      failed: summary.failed,
      results: summary.results,
      rejected: summary.rejected,
      deferred: summary.deferred
    });
  } catch (err) {
    Log.error('Request failed', err);
    return jsonOutput_({ ok: false, error: String(err.message || err) });
  }
}

/** Creates every List configured for one client and records the outcome in summary. */
function processJob_(job, index, mapping, batchId, summary) {
  const isExisting = isExistingClient_(job.client);
  const key = job.unit + '_' + isExisting;
  const specs = mapping[key];
  if (!specs) {
    summary.rejected.push({ index: index, client: job.client, key: key, error: 'no mapping for key' });
    return;
  }

  const lists = specs.map(function (spec) {
    const ctx = {
      batchId: batchId,
      client: String(job.client).trim(),
      unit: job.unit,
      linkedTask: String(job.linkedTask).trim(),
      isExisting: isExisting,
      templateId: spec.templateId,
      date: Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd')
    };
    ctx.linkedTaskUrl = 'https://app.clickup.com/t/' + ctx.linkedTask;
    ctx.targetName = spec.namePattern ? fillTemplate_(spec.namePattern, ctx) : ctx.client;

    try {
      ctx.listId = createListFromTemplate_(spec, ctx.targetName);
      ctx.listUrl = 'https://app.clickup.com/v/li/' + ctx.listId;
      setListDescription_(ctx.listId, spec, ctx);
      recordList_(ctx);
      summary.created++;
      Log.info('List created', { batchId: batchId, client: ctx.client, list: ctx.targetName, listId: ctx.listId });
      return { name: ctx.targetName, listId: ctx.listId, url: ctx.listUrl };
    } catch (err) {
      summary.failed++;
      Log.error('List creation failed', err, { batchId: batchId, client: ctx.client, list: ctx.targetName, templateId: spec.templateId });
      return { name: ctx.targetName, error: String(err.message || err) };
    }
  });

  summary.results.push({ client: job.client, key: key, isExisting: isExisting, lists: lists });
}

// ============================ ClickUp ============================

/**
 * Creates a List from a template and returns its ID.
 * 400 usually means the name is already taken in that Folder (which also stops
 * duplicates on webhook retries); 404 means template, Folder or Space not found.
 */
function createListFromTemplate_(spec, name) {
  if (!spec.templateId) throw new Error('templateId is missing in the mapping');
  if (!spec.folderTo === !spec.spaceTo) throw new Error('set exactly one of folderTo or spaceTo');

  const path = spec.folderTo
    ? '/folder/' + spec.folderTo + '/list_template/' + spec.templateId
    : '/space/' + spec.spaceTo + '/list_template/' + spec.templateId;
  const body = { name: name };
  if (spec.options && Object.keys(spec.options).length) body.options = spec.options;

  const response = clickUpFetch_(path, 'post', body);
  const listId = response.id || (response.list && response.list.id);
  if (!listId) throw new Error('No list ID in response: ' + JSON.stringify(response));
  return String(listId);
}

/**
 * Templates do not carry a description through, so it is set afterwards.
 * Large templates keep filling in after the List ID is returned, hence the retries.
 * A failed description never fails the request: the List exists either way.
 */
function setListDescription_(listId, spec, ctx) {
  const template = spec.descriptionTemplate !== undefined ? spec.descriptionTemplate : DESCRIPTION_TEMPLATE;
  if (!template) return;
  const text = fillTemplate_(template, Object.assign({}, ctx, {
    date: Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm')
  }));

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      let content = text;
      if (APPEND_DESCRIPTION) {
        const current = clickUpFetch_('/list/' + listId).content || '';
        content = current ? current + '\n\n---\n' + text : text;
      }
      clickUpFetch_('/list/' + listId, 'put', { content: content });
      return;
    } catch (err) {
      if (attempt === MAX_ATTEMPTS) {
        Log.warn('Description not set', err, { listId: listId });
        return;
      }
      Utilities.sleep(DESCRIPTION_RETRY_WAIT_MS);
    }
  }
}

/**
 * Whether the client already exists in the registry List.
 * On any error it returns false: creating a List is better than dropping a real
 * request over a transient API error. The error is logged.
 */
function isExistingClient_(client) {
  const registry = getRegistryConfig_();
  if (!registry.listId) return false;
  try {
    return getRegistryValues_(registry).indexOf(String(client).trim().toLowerCase()) >= 0;
  } catch (err) {
    Log.warn('Registry lookup failed, client treated as new', err, { client: client });
    return false;
  }
}

/** Lowercased names (or field values) of the registry List, cached so a batch costs one read. */
function getRegistryValues_(registry) {
  const cache = CacheService.getScriptCache();
  const cacheKey = ['registry', registry.listId, registry.matchBy, registry.fieldId].join('_');
  const hit = cache.get(cacheKey);
  if (hit) return JSON.parse(hit);

  const values = [];
  for (let page = 0; ; page++) {
    const tasks = clickUpFetch_('/list/' + registry.listId + '/task?archived=false&subtasks=false' +
      '&page=' + page + '&include_closed=' + (registry.includeClosed !== false)).tasks || [];
    tasks.forEach(function (task) {
      let value = task.name;
      if (registry.matchBy === 'field') {
        const field = (task.custom_fields || []).filter(function (f) { return f.id === registry.fieldId; })[0];
        value = field ? customFieldText_(field) : '';
      }
      if (value) values.push(String(value).trim().toLowerCase());
    });
    if (tasks.length < 100) break;
  }

  try {
    cache.put(cacheKey, JSON.stringify(values), registry.cacheSeconds || 120);
  } catch (err) {
    Log.debug('Registry too large for cache, not cached', { entries: values.length });
  }
  return values;
}

/** Readable text of a custom field value of any type. */
function customFieldText_(field) {
  const value = field.value;
  if (value === null || value === undefined) return '';
  if (Array.isArray(value)) return value.map(function (x) { return x.name || x.id || x; }).join(', ');
  if (typeof value === 'object') return value.name || value.id || '';
  if (field.type === 'drop_down' && field.type_config && field.type_config.options) {
    const option = field.type_config.options[value];   // dropdowns return the option index
    if (option) return option.name || '';
  }
  return String(value);
}

/** ClickUp API call with rate-limit retries. */
function clickUpFetch_(path, method, payload) {
  const options = {
    method: method || 'get',
    headers: { Authorization: getRequiredProperty_('CLICKUP_API_TOKEN') },
    contentType: 'application/json',
    muteHttpExceptions: true
  };
  if (payload) options.payload = JSON.stringify(payload);

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    const response = UrlFetchApp.fetch(CLICKUP_BASE_URL + path, options);
    const code = response.getResponseCode();
    if (code === 429) {
      Log.warn('ClickUp rate limit, waiting', { path: path, attempt: attempt });
      Utilities.sleep(RATE_LIMIT_WAIT_MS);
      continue;
    }
    if (code >= 400) throw new Error(path + ' -> ' + code + ': ' + response.getContentText());
    const body = response.getContentText();
    return body ? JSON.parse(body) : {};
  }
  throw new Error('ClickUp rate limit not cleared: ' + path);
}

// ============================ Records sheet ============================

/** One row per created List. Writes only its own columns; never fails list creation. */
function recordList_(ctx) {
  const spreadsheetId = getProperty_('RECORDS_SPREADSHEET_ID', null);
  if (!spreadsheetId) return;
  const records = getJsonProperty_('RECORDS_SHEET', DEFAULT_RECORDS);
  const firstColumn = getRecordsFirstColumn_();

  const lock = LockService.getScriptLock();   // parallel webhooks must not write the same row
  try {
    lock.waitLock(LOCK_TIMEOUT_MS);
    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(records.tab);
    if (!sheet) throw new Error('Records tab not found: ' + records.tab);

    const row = nextEmptyRowInColumn_(sheet, firstColumn);
    if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 20);
    const values = records.columns.map(function (key) { return ctx[key] !== undefined ? ctx[key] : ''; });
    sheet.getRange(row, firstColumn, 1, values.length).setValues([values]);
    SpreadsheetApp.flush();
  } catch (err) {
    Log.error('Record not written', err, { listId: ctx.listId, client: ctx.client });
  } finally {
    lock.releaseLock();
  }
}

/** Row after the last filled cell of one column, ignoring other columns. Row 1 is the header. */
function nextEmptyRowInColumn_(sheet, column) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = sheet.getRange(1, column, lastRow, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() !== '') return i + 2;
  }
  return 2;
}

// ============================ Admin tools (run from the editor) ============================

/** Lists the workspace List templates with their IDs, to fill in CLICKUP_LIST_MAPPING. */
function listTemplates() {
  const workspaceId = getRequiredProperty_('CLICKUP_WORKSPACE_ID');
  for (let page = 0; ; page++) {
    const templates = clickUpFetch_('/team/' + workspaceId + '/listTemplate?page=' + page).templates || [];
    templates.forEach(function (t) { Log.info('%s | %s', t.id, t.name); });
    if (templates.length < 100) break;
  }
}

/** Checks the mapping, the registry List and the records sheet after a config change. */
function validateConfig() {
  const mapping = getJsonProperty_('CLICKUP_LIST_MAPPING', {});
  Object.keys(mapping).forEach(function (key) {
    const specs = mapping[key] || [];
    if (!specs.length) return Log.info('%s | empty, nothing will be created', key);
    specs.forEach(function (spec, i) {
      const problems = [];
      if (!spec.templateId) problems.push('no templateId');
      else if (spec.templateId.indexOf('t-') !== 0) problems.push('templateId should start with t-');
      if (!spec.folderTo === !spec.spaceTo) problems.push('set exactly one of folderTo or spaceTo');
      Log[problems.length ? 'warn' : 'info']('%s [%s] | %s | name: %s | %s', key, i, spec.templateId || '?',
        spec.namePattern || '{client}', problems.length ? 'PROBLEM: ' + problems.join(', ') : 'ok');
    });
  });

  const registry = getRegistryConfig_();
  try {
    const list = clickUpFetch_('/list/' + registry.listId);
    Log.info('Registry | %s | entries: %s | match by: %s', list.name, getRegistryValues_(registry).length, registry.matchBy);
  } catch (err) {
    Log.error('Registry is not reachable', err);
  }

  const spreadsheetId = getProperty_('RECORDS_SPREADSHEET_ID', null);
  if (!spreadsheetId) return Log.info('Records sheet not configured');
  const records = getJsonProperty_('RECORDS_SHEET', DEFAULT_RECORDS);
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(records.tab);
  Log[sheet ? 'info' : 'error']('Records tab "%s": %s, first column: %s', records.tab, sheet ? 'found' : 'MISSING', getRecordsFirstColumn_());
}

// ============================ Helpers ============================

/** First column of the log row (1 = A). Script Property RECORDS_FIRST_COLUMN, default 4 (D). */
function getRecordsFirstColumn_() {
  const column = parseInt(getProperty_('RECORDS_FIRST_COLUMN', DEFAULT_RECORDS_FIRST_COLUMN), 10);
  if (!Number.isFinite(column) || column < 1) throw new Error('RECORDS_FIRST_COLUMN must be a column number (1 = A)');
  return column;
}

function getRegistryConfig_() {
  return Object.assign({}, DEFAULT_REGISTRY, getJsonProperty_('CLICKUP_CLIENT_REGISTRY', {}));
}

/** Replaces {key} placeholders; unknown keys become ''. */
function fillTemplate_(template, ctx) {
  return String(template).replace(/{(\w+)}/g, function (match, key) {
    return ctx[key] !== undefined && ctx[key] !== null ? ctx[key] : '';
  });
}

function jsonOutput_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
