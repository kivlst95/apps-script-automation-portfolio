/**
 * Xero OAuth 2.0 for several organisations (tenants) with one Xero app.
 *
 * Provides to XeroExport.js:
 *   XERO_COMPANIES, XERO_SCOPES, LAST_QUOTA, xeroFetch(), getXeroService_(),
 *   fetchXeroConnections_(), syncXeroTenants(), readXeroTenants_(), resetXero()
 *
 * Script Properties:
 *   XERO_CLIENT_ID, XERO_CLIENT_SECRET   Xero app credentials
 *   XERO_COMPANIES                       JSON: alias -> organisation name in Xero,
 *                                        e.g. {"entity_cy": "Example Holding Ltd", "entity_uk": "Example UK Ltd"}
 *
 * Authorization: open the link from the menu once per organisation, then
 * "Link organisations" maps every alias to its tenantId by organisation name.
 */

const XERO_SCOPES = [
  'offline_access',
  'accounting.invoices.read',
  'accounting.payments.read',
  'accounting.banktransactions.read',
  'accounting.contacts.read',
  'accounting.settings.read',
  'accounting.reports.trialbalance.read',
  'accounting.reports.profitandloss.read',
  'accounting.reports.balancesheet.read',
  'accounting.reports.aged.read',
  'accounting.reports.banksummary.read',
  'accounting.manualjournals.read',
  'accounting.attachments.read'
].join(' ');

/** Alias -> organisation name, from Script Property XERO_COMPANIES. */
const XERO_COMPANIES = getJsonProperty_('XERO_COMPANIES', {});

/** Internal state: alias -> {tenantId, tenantName}. Written by syncXeroTenants(). */
const XERO_TENANTS_KEY = 'XERO_TENANTS_STATE';

/** Remaining API quota from the last response headers (used to pause before the limit). */
var LAST_QUOTA = { day: null, minute: null };

function getXeroService_() {
  return OAuth2.createService('XeroMultiTenant')
    .setAuthorizationBaseUrl('https://login.xero.com/identity/connect/authorize')
    .setTokenUrl('https://identity.xero.com/connect/token')
    .setClientId(getRequiredProperty_('XERO_CLIENT_ID'))
    .setClientSecret(getRequiredProperty_('XERO_CLIENT_SECRET'))
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setLock(LockService.getScriptLock())
    .setScope(XERO_SCOPES);
}

function authCallback(request) {
  const ok = getXeroService_().handleCallback(request);
  if (!ok) {
    Log.warn('Xero authorization denied');
    return HtmlService.createHtmlOutput('Denied.');
  }
  const map = syncXeroTenants();
  Log.info('Xero authorized', { linked: Object.keys(map) });
  return HtmlService.createHtmlOutput('Success! Linked: ' + Object.keys(map).join(', ') + '. You can close this tab.');
}

/** All organisations the app is connected to. */
function fetchXeroConnections_() {
  const res = UrlFetchApp.fetch('https://api.xero.com/connections', {
    headers: { Authorization: 'Bearer ' + getXeroService_().getAccessToken(), Accept: 'application/json' },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Xero /connections failed with HTTP ' + res.getResponseCode() + ': ' + res.getContentText());
  }
  return JSON.parse(res.getContentText());
}

/** Maps every alias in XERO_COMPANIES to a connected organisation by name. */
function syncXeroTenants() {
  const connections = fetchXeroConnections_();
  const map = {};
  Object.keys(XERO_COMPANIES).forEach(function (alias) {
    const wanted = String(XERO_COMPANIES[alias]).trim().toLowerCase();
    const hit = connections.filter(function (c) {
      return String(c.tenantName).trim().toLowerCase() === wanted;
    })[0];
    if (hit) map[alias] = { tenantId: hit.tenantId, tenantName: hit.tenantName };
    else Log.warn('Xero organisation not found for alias %s (%s)', alias, XERO_COMPANIES[alias]);
  });
  PropertiesService.getScriptProperties().setProperty(XERO_TENANTS_KEY, JSON.stringify(map));
  return map;
}

/** Alias -> {tenantId, tenantName} as linked last time. */
function readXeroTenants_() {
  return getJsonProperty_(XERO_TENANTS_KEY, {});
}

/** Deletes tokens and tenant links. */
function resetXero() {
  getXeroService_().reset();
  PropertiesService.getScriptProperties().deleteProperty(XERO_TENANTS_KEY);
  Log.info('Xero authorization and tenant links reset');
}

/**
 * GET request to the Xero Accounting API for one organisation.
 * Handles the per-minute limit (waits and retries) and throws an error with
 * quotaExhausted = true when the daily limit is reached, so the export can pause.
 *
 * @param {string} company alias from XERO_COMPANIES
 * @param {string} path    path after /api.xro/2.0/, e.g. 'Invoices?page=1'
 * @return {Object|null}   parsed JSON, or null for 404
 */
function xeroFetch(company, path) {
  const tenant = readXeroTenants_()[company];
  if (!tenant) throw new Error('Xero organisation "' + company + '" is not linked. Run syncXeroTenants().');

  for (let attempt = 1; attempt <= 4; attempt++) {
    const res = UrlFetchApp.fetch('https://api.xero.com/api.xro/2.0/' + path, {
      headers: {
        Authorization: 'Bearer ' + getXeroService_().getAccessToken(),
        'xero-tenant-id': tenant.tenantId,
        Accept: 'application/json'
      },
      muteHttpExceptions: true
    });

    const headers = lowerCaseKeys_(res.getAllHeaders());
    if (headers['x-daylimit-remaining'] !== undefined) LAST_QUOTA.day = Number(headers['x-daylimit-remaining']);
    if (headers['x-minlimit-remaining'] !== undefined) LAST_QUOTA.minute = Number(headers['x-minlimit-remaining']);

    const code = res.getResponseCode();
    if (code === 200) return JSON.parse(res.getContentText());
    if (code === 404) return null;

    if (code === 429) {
      if (String(headers['x-rate-limit-problem'] || '').toLowerCase() === 'day') {
        const err = new Error('QUOTA_EXHAUSTED: Xero daily limit reached for ' + company);
        err.quotaExhausted = true;
        throw err;
      }
      const waitSec = Math.min(Number(headers['retry-after']) || 5, 60);
      Log.warn('Xero minute limit for %s, waiting %s s (attempt %s)', company, waitSec, attempt);
      Utilities.sleep(waitSec * 1000);
      continue;
    }

    if (code >= 500 && attempt < 4) {
      Log.warn('Xero HTTP %s on %s, retrying (attempt %s)', code, path, attempt);
      Utilities.sleep(2000 * attempt);
      continue;
    }

    throw new Error('Xero HTTP ' + code + ' on ' + company + ' / ' + path + ': ' + res.getContentText().slice(0, 300));
  }
  throw new Error('Xero request failed after retries: ' + company + ' / ' + path);
}

function lowerCaseKeys_(obj) {
  const out = {};
  Object.keys(obj || {}).forEach(function (k) { out[k.toLowerCase()] = obj[k]; });
  return out;
}
