/**
 * Xero API client (single organisation) for Google Apps Script.
 *
 * OAuth 2.0 via the apps-script-oauth2 library (see appsscript.json).
 *
 * Script Properties:
 *   XERO_CLIENT_ID, XERO_CLIENT_SECRET   Xero app credentials
 *   XERO_TENANT_ID                       (optional) organisation to use; default = first connected
 *   XERO_HANDBOOK_SPREADSHEET_ID         (optional) spreadsheet for reference data
 *   XERO_HANDBOOK_SHEET_GIDS             (optional) JSON {"contacts": 0, "items": 123, "accounts": 456}
 *
 * Setup: run logXeroRedirectUri(), register the URI in the Xero app,
 * run xeroApi_('/Organisation') once and open the authorization link from the log.
 */

const XERO_API_BASE = 'https://api.xero.com/api.xro/2.0';
const XERO_SCOPES = [
  'offline_access',
  'accounting.invoices',
  'accounting.payments',
  'accounting.banktransactions',
  'accounting.contacts',
  'accounting.manualjournals',
  'accounting.settings'
].join(' ');

/** Configured OAuth2 service. Tokens are stored in Script Properties. */
function getXeroService_() {
  return OAuth2.createService('Xero')
    .setAuthorizationBaseUrl('https://login.xero.com/identity/connect/authorize')
    .setTokenUrl('https://identity.xero.com/connect/token')
    .setClientId(getRequiredProperty_('XERO_CLIENT_ID'))
    .setClientSecret(getRequiredProperty_('XERO_CLIENT_SECRET'))
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope(XERO_SCOPES);
}

/** OAuth callback: stores the tenant of the first connected organisation. */
function authCallback(request) {
  const service = getXeroService_();
  if (!service.handleCallback(request)) {
    Log.warn('Xero authorization denied');
    return HtmlService.createHtmlOutput('Denied.');
  }
  const response = UrlFetchApp.fetch('https://api.xero.com/connections', {
    headers: { Authorization: 'Bearer ' + service.getAccessToken() }
  });
  const connections = JSON.parse(response.getContentText());
  service.getStorage().setValue('tenantId', connections[0].tenantId);
  Log.info('Xero authorized', { tenantName: connections[0].tenantName, connections: connections.length });
  return HtmlService.createHtmlOutput('Success! You can close this tab.');
}

/** Tenant to send in the xero-tenant-id header. */
function getXeroTenantId_() {
  return getProperty_('XERO_TENANT_ID', null) || getXeroService_().getStorage().getValue('tenantId');
}

/**
 * Calls the Xero Accounting API.
 * @param {string} urlTail  path after /api.xro/2.0, e.g. '/Invoices?page=1'
 * @param {string} [method] HTTP method, default GET
 * @param {Object} [payload] request body for POST/PUT
 * @param {boolean} [muteHttp] return the parsed error body instead of throwing
 * @return {Object|null} parsed JSON response, or null when authorization is required
 */
function xeroApi_(urlTail, method, payload, muteHttp) {
  method = String(method || 'get').toUpperCase();
  const service = getXeroService_();
  if (!service.hasAccess()) {
    Log.warn('Xero authorization required: open %s and run again', service.getAuthorizationUrl());
    return null;
  }

  const options = {
    method: method,
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + service.getAccessToken(),
      'xero-tenant-id': getXeroTenantId_(),
      Accept: 'application/json'
    }
  };
  if (payload && method !== 'GET') {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }

  const response = UrlFetchApp.fetch(XERO_API_BASE + urlTail, options);
  const code = response.getResponseCode();
  const text = response.getContentText();
  const body = text ? JSON.parse(text) : {};

  if (code >= 300) {
    Log.error('Xero %s %s failed with HTTP %s', method, urlTail, code, { response: text.slice(0, 500) });
    if (!muteHttp) throw new Error('Xero ' + method + ' ' + urlTail + ' failed with HTTP ' + code);
  } else {
    Log.debug('Xero %s %s: HTTP %s', method, urlTail, code);
  }
  return body;
}

/** Clears stored tokens so the authorization can be repeated. */
function resetXeroAuth() {
  getXeroService_().reset();
  Log.info('Xero authorization reset');
}

/** Logs the redirect URI to register in the Xero app settings. */
function logXeroRedirectUri() {
  Log.info('Xero redirect URI: %s', OAuth2.getRedirectUri());
}

/* ---------------- Reference data (handbook) ---------------- */

function handbookSheet_(name) {
  const gids = getJsonProperty_('XERO_HANDBOOK_SHEET_GIDS', {});
  if (gids[name] === undefined) throw new Error('XERO_HANDBOOK_SHEET_GIDS has no entry for ' + name);
  return SpreadsheetApp.openById(getRequiredProperty_('XERO_HANDBOOK_SPREADSHEET_ID'))
    .getSheetById(Number(gids[name]));
}

function writeHandbook_(sheet, header, rows) {
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), header.length).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}

/** Writes all Xero contacts (ID, name) to the handbook, page by page. */
function syncXeroContactsToHandbook() {
  const rows = [];
  for (let page = 1; ; page++) {
    const contacts = (xeroApi_('/Contacts?page=' + page, 'get', null, true) || {}).Contacts || [];
    if (!contacts.length) break;
    contacts.forEach(function (c) { rows.push([c.ContactID, c.Name]); });
  }
  writeHandbook_(handbookSheet_('contacts'), ['ContactID', 'Name'], rows);
  Log.info('Handbook: %s contacts written', rows.length);
}

/** Writes all Xero items to the handbook. */
function syncXeroItemsToHandbook() {
  const items = (xeroApi_('/Items', 'get', null, true) || {}).Items || [];
  const rows = items.map(function (i) { return [i.ItemID, i.Code, i.Description, i.Name]; });
  writeHandbook_(handbookSheet_('items'), ['ItemID', 'Code', 'Description', 'Name'], rows);
  Log.info('Handbook: %s items written', rows.length);
}

/** Writes the chart of accounts to the handbook. */
function syncXeroAccountsToHandbook() {
  const accounts = (xeroApi_('/Accounts', 'get', null, true) || {}).Accounts || [];
  const rows = accounts.map(function (a) { return [a.AccountID, a.Code, a.Name, a.Type, a.Status, a.TaxType]; });
  writeHandbook_(handbookSheet_('accounts'), ['AccountID', 'Code', 'Name', 'Type', 'Status', 'TaxType'], rows);
  Log.info('Handbook: %s accounts written', rows.length);
}
