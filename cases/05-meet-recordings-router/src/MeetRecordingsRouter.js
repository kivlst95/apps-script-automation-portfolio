/**
 * Meet Recordings Router
 * Google Meet recordings -> project folder (matched by keyword) -> Notion page.
 *
 * Script Properties (Project Settings -> Script properties):
 *   GOOGLE_SERVICE_ACCOUNT_KEY                 - service account JSON key
 *   NOTION_API_TOKEN           - Notion internal integration secret
 *   ROUTING_SPREADSHEET_ID              - spreadsheet with Routing / Hosts / Log sheets
 *   FALLBACK_DRIVE_FOLDER_ID     - (optional) folder for unmatched meetings
 *   FALLBACK_NOTION_DATABASE_ID  - (optional) Notion database for unmatched meetings
 *   Without fallback, unmatched meetings stay in place and are logged as NO_MATCH
 *   ALERT_EMAIL            - where errors are sent
 *   WEBHOOK_SECRET            - shared secret for the web app (keyword lookup from other Zaps)
 *
 * Run order: setup() -> fill sheets -> testAuth() -> run() with DRY_RUN=true -> DRY_RUN=false -> installTrigger()
 * Note: inline comments are partly in Ukrainian (original working language).
 */

const CONFIG = {
  DRY_RUN: false,                 // true = only writes to Log what it would do. Switch to false after checking
  LOOKBACK_HOURS: 24,            // period for reviewing finished meetings
  GRACE_HOURS: 3,                // after how many hours a meeting without a recording is treated as not recorded
  MAX_RUNTIME_MS: 5 * 60 * 1000, // buffer before the 6-minute Apps Script limit
  RENAME_FILES: true,            // "2026-10-06 14:00 - Meeting title"
  CLEANUP_EMPTY_FOLDERS: true,   // move empty meeting subfolders to trash after moving files (except recurring meetings)
  MEET_ROOT_FOLDER_NAME: 'Google Meet', // name of the root folder Meet creates in the host My Drive
  TIMEZONE: 'Europe/Berlin',
  NOTION_VERSION: '2022-06-28',
  // Property names in Notion databases. null = do not fill this property
  // Notion columns are looked up by name from the list, and if not found, by type (when the database has only one column of that type).
  // The API does not return column order, so binding is by type, not by position.
  // So databases with "Title/State/When/URL" and with "Name/Status/Date/Meeting Link" work without setup.
  NOTION_COLUMNS: {
    status: ['Status', 'Status', 'Tags', 'Tags'],          // Status, Select or Multi-select
    date: ['Date', 'Date'],                                // Date
    recording: ['Link', 'Meeting Link', 'Link', 'Recording'], // URL
    // Optional: filled only if such a column exists in the database
    attendees: ['Participants', 'Attendees'],                  // Text
    keyword: ['Keyword', 'Keyword'],                 // Text
    host: ['Host', 'Host'],                                // Email
  },
  // Status value: the first one that exists in the database is used (case-insensitive: "completed" = "Completed").
  // If none exists: Select / Multi-select with one option uses that option; Status uses an option from the Complete group;
  // Select with several options creates the first value from the list; Multi-select with several options is not filled
  NOTION_STATUS_VALUES: ['Done', 'Held', 'Finished', 'Done', 'Completed', 'Done'],
  // 'auto' = like existing database records (with time if they have time), true = always with time, false = date only
  NOTION_DATE_WITH_TIME: 'auto',
  SHEETS: { routing: 'KEYWORDS', hosts: 'USERS', log: 'Log' },
  SCOPES: [
    'https://www.googleapis.com/auth/meetings.space.readonly',
    'https://www.googleapis.com/auth/calendar.readonly',
    'https://www.googleapis.com/auth/drive',
  ],
};

const MEET = 'https://meet.googleapis.com/v2';
const DRIVE = 'https://www.googleapis.com/drive/v3';
const CAL = 'https://www.googleapis.com/calendar/v3';

/* ======================= MAIN ======================= */

function run() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // the previous run is still in progress
  const started = Date.now();
  try {
    const ss = SpreadsheetApp.openById(prop_('ROUTING_SPREADSHEET_ID'));
    const hosts = getHosts_(ss);
    const rules = getRules_(ss);
    const processed = getProcessed_(ss);
    const since = new Date(Date.now() - CONFIG.LOOKBACK_HOURS * 3600e3).toISOString();
    const filter = encodeURIComponent(`end_time>="${since}"`);

    for (const host of hosts) {
      if (Date.now() - started > CONFIG.MAX_RUNTIME_MS) break;
      let records;
      try {
        records = listAll_(host, `${MEET}/conferenceRecords?filter=${filter}`, 'conferenceRecords');
      } catch (e) {
        alert_(`Could not get meetings for ${host}`, e);
        continue;
      }
      for (const rec of records) {
        if (Date.now() - started > CONFIG.MAX_RUNTIME_MS) break;
        if (processed.records.has(rec.name) || isSkipped_(rec.name)) continue;
        try {
          processRecord_(ss, host, rec, rules, processed);
        } catch (e) {
          logRow_(ss, { record: rec.name, host, status: 'ERROR', error: String(e.message || e).slice(0, 1000) });
          alert_(`Error processing ${rec.name} (${host})`, e);
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function processRecord_(ss, host, rec, rules, processed) {
  const hoursSinceEnd = (Date.now() - new Date(rec.endTime).getTime()) / 3600e3;
  const recordings = listAll_(host, `${MEET}/${rec.name}/recordings`, 'recordings');

  if (!recordings.length) {
    if (hoursSinceEnd > CONFIG.GRACE_HOURS) markSkipped_(rec.name); // meeting without recording
    return;
  }
  if (recordings.some(r => r.state !== 'FILE_GENERATED')) return; // still processing

  const space = gapi_(host, `${MEET}/${rec.space}`);
  const event = findEvent_(host, space.meetingCode, rec.startTime);

  const videoIds = recordings.map(r => r.driveDestination.file);

  // One meeting can be visible to several people from Hosts (co-organisers, several managers on a call).
  // If these videos are already processed, skip them to avoid duplicates in Log and Notion.
  const markDone = () => { processed.records.add(rec.name); videoIds.forEach(id => processed.files.add(id)); };
  if (videoIds.every(id => processed.files.has(id))) { markDone(); return; }

  // Only the owner can move a file, so we act on behalf of the recording owner,
  // not the host through whom the meeting was found
  const firstFile = gapi_(host, `${DRIVE}/files/${videoIds[0]}?fields=name,owners(emailAddress)&supportsAllDrives=true`);
  const firstFileName = firstFile.name;
  const owner = ((firstFile.owners || [])[0] || {}).emailAddress;
  const domain = host.split('@')[1];
  const actor = owner && owner.toLowerCase().endsWith('@' + domain) ? owner.toLowerCase() : host;

  const title = (event && event.summary) || firstFileName;
  const route = routeProject_(rules, title, firstFileName);
  if (!route) {
    // No keyword and no fallback configured: files are left untouched
    logRow_(ss, { record: rec.name, host: actor, title, fileIds: videoIds.join(','), status: 'NO_MATCH' });
    markDone();
    return;
  }
  const attendees = event ? (event.attendees || []).filter(a => !a.resource).map(a => a.email).join(', ') : '';

  if (CONFIG.DRY_RUN) {
    logRow_(ss, {
      record: rec.name, host: actor, title, keyword: route.keyword, folderId: route.folderId,
      notionDbId: route.notionDbId, fileIds: videoIds.join(','), status: 'DRY_RUN',
    });
    markDone(); // in memory for this run only: DRY_RUN is not treated as processed next time
    return;
  }

  const stamp = Utilities.formatDate(new Date(rec.startTime), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  const base = `${stamp} — ${title}`;
  const sourceFolders = [];
  videoIds.forEach((id, i) => sourceFolders.push(
    ...moveFile_(actor, id, route.folderId, CONFIG.RENAME_FILES ? base + (videoIds.length > 1 ? ` (part ${i + 1})` : '') : null)));
  cleanupSourceFolders_(actor, sourceFolders, event);

  const link = id => `https://drive.google.com/file/d/${id}/view`;
  const notionUrl = createNotionPage_(route.notionDbId, {
    title,
    start: rec.startTime,
    end: rec.endTime,
    recordingUrl: link(videoIds[0]),
    attendees,
    keyword: route.keyword,
    host: actor,
    extraLinks: videoIds.slice(1).map(link), // if the recording is split into several parts
  });

  logRow_(ss, {
    record: rec.name, host: actor, title, keyword: route.keyword, folderId: route.folderId, notionDbId: route.notionDbId,
    fileIds: videoIds.join(','), notionUrl, status: 'DONE',
  });
  markDone();
}

/* ======================= ROUTING ======================= */

// All rules whose keyword occurs in the text, from the highest priority
function findMatches_(rules, ...texts) {
  const haystack = texts.filter(Boolean).join(' | ').toLowerCase();
  return rules
    .filter(r => haystack.includes(r.keyword.toLowerCase()))
    .sort((a, b) => b.priority - a.priority || b.keyword.length - a.keyword.length);
}

function routeProject_(rules, ...texts) {
  // Moving recordings needs rules with a folder and a Notion database; rows with only direction/meet_type are skipped here
  const hits = findMatches_(rules, ...texts).filter(r => r.folderId && r.notionDbId);
  if (!hits.length) {
    const folderId = prop_('FALLBACK_DRIVE_FOLDER_ID', false);
    const notionDbId = prop_('FALLBACK_NOTION_DATABASE_ID', false);
    if (!folderId || !notionDbId) return null; // fallback not configured
    return { keyword: null, folderId: cleanFolderId_(folderId), notionDbId: cleanNotionId_(notionDbId) };
  }
  return hits[0];
}

// The Routing sheet is read by header names, so the column order does not matter.
// All columns except service ones go to r.extra (e.g. direction, meet_type) and are returned via the Web App.
// The value - is treated as empty.
const ROUTING_FIELDS = ['keyword', 'folder_id', 'notion_db_id', 'priority', 'active', 'comment'];

function getRules_(ss) {
  const values = sheet_(ss, CONFIG.SHEETS.routing).getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const idx = name => headers.indexOf(name);
  if (idx('keyword') < 0) throw new Error('The Routing sheet has no keyword column');
  const clean = v => { const t = String(v == null ? '' : v).trim(); return t === '-' ? '' : t; };

  return values.slice(1).map(r => {
    const get = name => (idx(name) >= 0 ? clean(r[idx(name)]) : '');
    if (String(get('active')).toUpperCase() === 'FALSE') return null;
    const extra = {};
    headers.forEach((h, i) => { if (h && !ROUTING_FIELDS.includes(h)) extra[h] = clean(r[i]); });
    return {
      keyword: get('keyword'),
      folderId: get('folder_id') ? cleanFolderId_(get('folder_id')) : '',
      notionDbId: get('notion_db_id') ? cleanNotionId_(get('notion_db_id')) : '',
      priority: Number(get('priority')) || 0,
      extra,
    };
  }).filter(r => r && r.keyword);
}

function getHosts_(ss) {
  return sheet_(ss, CONFIG.SHEETS.hosts).getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[1] !== false && String(r[1]).toUpperCase() !== 'FALSE')
    .map(r => String(r[0]).trim().toLowerCase());
}

// Accepts both an ID and a full folder link
function cleanFolderId_(v) {
  const s = String(v).trim();
  const m = s.match(/folders\/([\w-]+)/);
  return m ? m[1] : s;
}

// Accepts both an ID and a Notion database link (takes the ID before ?v=, not the view ID)
function cleanNotionId_(v) {
  const s = String(v).trim().split('?')[0].replace(/-/g, '');
  const m = s.match(/([0-9a-f]{32})(?!.*[0-9a-f]{32})/i);
  return m ? m[1] : s;
}

/* ======================= WEB APP FOR ZAPIER ======================= */
// The same keyword matching logic as in run(), available to other Zaps.
// GET  ?token=...&text=...           (recommended for Webhooks by Zapier)
// POST {"token": "...", "text": "..."}
// Response: { ok, matched, keyword, folder_id, notion_db_id, priority, <all extra Routing columns>, matches: [...] }
// Extra columns (direction, meet_type, user_id, ...) are picked up automatically by header

function doGet(e) {
  Log.info('doGet: request received', { hasBody: !!(e && e.postData), params: (e && e.parameter) ? Object.keys(e.parameter).filter(function (k) { return k !== 'token'; }) : [] });
  return routeResponse_((e && e.parameter) || {});
}

function doPost(e) {
  Log.info('doPost: request received', { hasBody: !!(e && e.postData), params: (e && e.parameter) ? Object.keys(e.parameter).filter(function (k) { return k !== 'token'; }) : [] });
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { /* empty body */ }
  return routeResponse_(body);
}

function routeResponse_(params) {
  const out = obj => ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  try {
    if (!params.token || params.token !== prop_('WEBHOOK_SECRET')) return out({ ok: false, error: 'unauthorized' });
    const ss = SpreadsheetApp.openById(prop_('ROUTING_SPREADSHEET_ID'));
    const hits = findMatches_(getRules_(ss), String(params.text || ''));
    const toJson = r => ({ ...r.extra, keyword: r.keyword, folder_id: r.folderId, notion_db_id: r.notionDbId, priority: r.priority });
    if (!hits.length) return out({ ok: true, matched: false, matches: [] });
    return out({ ok: true, matched: true, ...toJson(hits[0]), matches: hits.map(toJson) });
  } catch (err) {
    Log.error('routeResponse_ failed', err);
    return out({ ok: false, error: String(err.message || err) });
  }
}

/* ======================= GOOGLE APIs ======================= */

function findEvent_(user, meetingCode, startTime) {
  if (!meetingCode) return null;
  const t = new Date(startTime).getTime();
  const q = [
    'timeMin=' + encodeURIComponent(new Date(t - 3 * 3600e3).toISOString()),
    'timeMax=' + encodeURIComponent(new Date(t + 3 * 3600e3).toISOString()),
    'singleEvents=true',
    'maxResults=250',
  ].join('&');
  const items = gapi_(user, `${CAL}/calendars/primary/events?${q}`).items || [];
  const code = meetingCode.toLowerCase();
  return items.find(e =>
    ((e.conferenceData && e.conferenceData.conferenceId) || '').toLowerCase() === code ||
    (e.hangoutLink || '').toLowerCase().includes(code)) || null;
}

// Returns the IDs of folders the file was removed from
function moveFile_(user, fileId, folderId, newName) {
  const f = gapi_(user, `${DRIVE}/files/${fileId}?fields=name,parents&supportsAllDrives=true`);
  const parents = f.parents || [];
  const params = ['supportsAllDrives=true', 'fields=id'];
  let removed = [];
  if (!parents.includes(folderId)) {
    params.push('addParents=' + folderId);
    removed = parents.filter(p => p !== folderId);
    if (removed.length) params.push('removeParents=' + removed.join(','));
  }
  const body = {};
  if (newName) {
    const ext = (f.name.match(/\.(mp4|webm|m4a|mov)$/i) || [''])[0];
    if (f.name !== newName + ext) body.name = newName + ext;
  }
  if (params.length === 2 && !body.name) return removed; // already in place with the right name
  gapi_(user, `${DRIVE}/files/${fileId}?${params.join('&')}`, { method: 'patch', body });
  return removed;
}

// Moves meeting subfolders that became empty to trash. Safety checks:
// - only one-time meetings (recurring ones share one folder that Meet keeps writing to)
// - only folders owned by the host and only direct subfolders of Google Meet (not Legacy Meet Recordings)
// - only if no file is left in the folder (transcripts, Gemini notes and chat stay, and so does the folder)
// Trash, not deletion: 30 days to restore
function cleanupSourceFolders_(user, folderIds, event) {
  if (!CONFIG.CLEANUP_EMPTY_FOLDERS) return;
  if (!event || event.recurringEventId || event.recurrence) return;
  [...new Set(folderIds)].forEach(id => {
    try {
      const folder = gapi_(user, `${DRIVE}/files/${id}?fields=name,mimeType,parents,ownedByMe&supportsAllDrives=true`);
      if (folder.mimeType !== 'application/vnd.google-apps.folder' || !folder.ownedByMe) return;
      if (/meet recordings/i.test(folder.name)) return;
      const parentId = (folder.parents || [])[0];
      if (!parentId) return;
      const parent = gapi_(user, `${DRIVE}/files/${parentId}?fields=name&supportsAllDrives=true`);
      if (parent.name !== CONFIG.MEET_ROOT_FOLDER_NAME) return;
      const q = encodeURIComponent(`'${id}' in parents and trashed=false`);
      const children = gapi_(user, `${DRIVE}/files?q=${q}&pageSize=1&fields=files(id)&supportsAllDrives=true&includeItemsFromAllDrives=true`).files || [];
      if (children.length) return;
      gapi_(user, `${DRIVE}/files/${id}?supportsAllDrives=true&fields=id`, { method: 'patch', body: { trashed: true } });
    } catch (e) {
      Log.warn(`Folder ${id} not cleaned up: ${e.message}`);
    }
  });
}

function listAll_(user, url, key) {
  const out = [];
  let pageToken = '';
  do {
    const sep = url.includes('?') ? '&' : '?';
    const res = gapi_(user, url + (pageToken ? `${sep}pageToken=${encodeURIComponent(pageToken)}` : ''));
    out.push(...(res[key] || []));
    pageToken = res.nextPageToken;
  } while (pageToken);
  return out;
}

function gapi_(user, url, opt = {}) {
  let lastError;
  for (let attempt = 0; attempt < 4; attempt++) {
    const params = {
      method: opt.method || 'get',
      headers: { Authorization: 'Bearer ' + getToken_(user) },
      muteHttpExceptions: true,
    };
    if (opt.body) {
      params.contentType = 'application/json';
      params.payload = JSON.stringify(opt.body);
    }
    const res = UrlFetchApp.fetch(url, params);
    const code = res.getResponseCode();
    if (code < 300) return JSON.parse(res.getContentText() || '{}');
    lastError = `${code} ${url}\n${res.getContentText()}`;
    if (code === 429 || code >= 500) { Utilities.sleep(1000 * Math.pow(2, attempt)); continue; }
    break;
  }
  throw new Error(lastError);
}

// Service account token on behalf of a user (Domain-Wide Delegation), no external libraries
function getToken_(user) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'tok_' + Utilities.base64EncodeWebSafe(user);
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const sa = JSON.parse(prop_('GOOGLE_SERVICE_ACCOUNT_KEY'));
  const now = Math.floor(Date.now() / 1000);
  const enc = obj => Utilities.base64EncodeWebSafe(JSON.stringify(obj)).replace(/=+$/, '');
  const unsigned = enc({ alg: 'RS256', typ: 'JWT' }) + '.' + enc({
    iss: sa.client_email,
    sub: user,
    scope: CONFIG.SCOPES.join(' '),
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600,
  });
  const signature = Utilities.base64EncodeWebSafe(
    Utilities.computeRsaSha256Signature(unsigned, sa.private_key)).replace(/=+$/, '');

  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: `${unsigned}.${signature}` },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error(`Token for ${user}: ${res.getContentText()}`);
  const token = JSON.parse(res.getContentText()).access_token;
  cache.put(cacheKey, token, 3000); // 50 minutes
  return token;
}

/* ======================= NOTION ======================= */

const _notionSchemas = {}; // database schema cache for one run
function getNotionSchema_(databaseId) {
  if (_notionSchemas[databaseId]) return _notionSchemas[databaseId];
  const res = UrlFetchApp.fetch(`https://api.notion.com/v1/databases/${databaseId}`, {
    headers: { Authorization: 'Bearer ' + prop_('NOTION_API_TOKEN'), 'Notion-Version': CONFIG.NOTION_VERSION },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error(`Notion schema ${databaseId}: ${res.getContentText()}`);
  _notionSchemas[databaseId] = JSON.parse(res.getContentText()).properties;
  return _notionSchemas[databaseId];
}

// Determines which columns of a specific database to write to and which status value to set
function resolveNotionColumns_(schema) {
  const C = CONFIG.NOTION_COLUMNS;
  const entries = Object.entries(schema);
  // Name lookup is case-insensitive and ignores extra spaces: "Tags" = "tags"
  const norm = x => String(x).trim().toLowerCase();
  const byName = (names, types) => {
    for (const n of names) {
      const hit = entries.find(([key, p]) => norm(key) === norm(n) && types.includes(p.type));
      if (hit) return hit[0];
    }
    return null;
  };
  const byType = types => {
    const found = entries.filter(([, p]) => types.includes(p.type));
    return found.length === 1 ? found[0][0] : null; // by type only if there is a single column of that type
  };
  const cols = {
    title: (entries.find(([, p]) => p.type === 'title') || [null])[0],
    status: byName(C.status, ['status', 'select', 'multi_select']) || byType(['status', 'select', 'multi_select']),
    date: byName(C.date, ['date']) || byType(['date']),
    recording: byName(C.recording, ['url']) || byType(['url']),
    attendees: byName(C.attendees, ['rich_text']),
    keyword: byName(C.keyword, ['rich_text']),
    host: byName(C.host, ['email', 'rich_text']),
  };

  let statusValue = null;
  if (cols.status) {
    const p = schema[cols.status];
    const options = (p[p.type].options || []);
    for (const v of CONFIG.NOTION_STATUS_VALUES) {
      const opt = options.find(o => o.name.trim().toLowerCase() === v.toLowerCase());
      if (opt) { statusValue = opt.name; break; } // use the name exactly as written in the database
    }
    // Select / Multi-select with a single option: use it, whatever its name
    if (!statusValue && (p.type === 'multi_select' || p.type === 'select') && options.length === 1) {
      statusValue = options[0].name;
    }
    if (!statusValue && p.type === 'status') {
      const complete = (p.status.groups || []).find(g => /complete/i.test(g.name));
      const opt = complete && options.find(o => (complete.option_ids || []).includes(o.id));
      statusValue = opt ? opt.name : null;
    }
    if (!statusValue && p.type === 'select') statusValue = CONFIG.NOTION_STATUS_VALUES[0]; // Select creates the option itself
  }
  return { cols, statusValue };
}

// Look at the last few database records: if the date has a time, write with time
const _dateModes = {};
function notionDateHasTime_(databaseId, dateProp) {
  const key = databaseId + '|' + dateProp;
  if (key in _dateModes) return _dateModes[key];
  const res = UrlFetchApp.fetch(`https://api.notion.com/v1/databases/${databaseId}/query`, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + prop_('NOTION_API_TOKEN'), 'Notion-Version': CONFIG.NOTION_VERSION },
    payload: JSON.stringify({
      page_size: 5,
      filter: { property: dateProp, date: { is_not_empty: true } },
      sorts: [{ timestamp: 'created_time', direction: 'descending' }],
    }),
    muteHttpExceptions: true,
  });
  let hasTime = false;
  if (res.getResponseCode() === 200) {
    hasTime = (JSON.parse(res.getContentText()).results || [])
      .some(pg => ((((pg.properties || {})[dateProp] || {}).date || {}).start || '').includes('T'));
  }
  _dateModes[key] = hasTime;
  return hasTime;
}

function createNotionPage_(databaseId, d) {
  const schema = getNotionSchema_(databaseId);
  const { cols, statusValue } = resolveNotionColumns_(schema);
  if (!cols.title) throw new Error(`Notion ${databaseId}: title column not found`);
  if (!cols.recording) throw new Error(`Notion ${databaseId}: link column (URL) not found`);

  const text = s => [{ text: { content: String(s || '—').slice(0, 2000) } }];
  const props = {};
  props[cols.title] = { title: text(d.title) };
  props[cols.recording] = { url: d.recordingUrl };
  if (cols.status && statusValue) {
    const type = schema[cols.status].type;
    props[cols.status] = type === 'multi_select' ? { multi_select: [{ name: statusValue }] }
      : type === 'select' ? { select: { name: statusValue } }
        : { status: { name: statusValue } };
  }
  if (cols.date) {
    const withTime = CONFIG.NOTION_DATE_WITH_TIME === 'auto'
      ? notionDateHasTime_(databaseId, cols.date)
      : CONFIG.NOTION_DATE_WITH_TIME;
    props[cols.date] = withTime
      ? { date: { start: d.start, end: d.end } }
      : { date: { start: Utilities.formatDate(new Date(d.start), CONFIG.TIMEZONE, 'yyyy-MM-dd') } };
  }
  if (cols.attendees) props[cols.attendees] = { rich_text: text(d.attendees) };
  if (cols.keyword) props[cols.keyword] = { rich_text: text(d.keyword || '—') };
  if (cols.host) props[cols.host] = schema[cols.host].type === 'email' ? { email: d.host } : { rich_text: text(d.host) };

  const children = (d.extraLinks || []).map(url => ({ object: 'block', type: 'bookmark', bookmark: { url } }));
  const body = { parent: { database_id: databaseId }, properties: props };
  if (children.length) body.children = children;

  for (let attempt = 0; attempt < 4; attempt++) {
    const res = UrlFetchApp.fetch('https://api.notion.com/v1/pages', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + prop_('NOTION_API_TOKEN'), 'Notion-Version': CONFIG.NOTION_VERSION },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code < 300) return JSON.parse(res.getContentText()).url;
    if (code === 429 || code >= 500) {
      const retryAfter = Number(res.getHeaders()['Retry-After'] || res.getHeaders()['retry-after']) || Math.pow(2, attempt);
      Utilities.sleep(retryAfter * 1000);
      continue;
    }
    throw new Error(`Notion ${code} (database ${databaseId}): ${res.getContentText()}`);
  }
  throw new Error(`Notion: retry limit exceeded (database ${databaseId})`);
}

/* ======================= STATE / LOG ======================= */

function getProcessed_(ss) {
  const values = sheet_(ss, CONFIG.SHEETS.log).getDataRange().getValues().slice(1);
  // NO_MATCH also counts as processed. To process a meeting again (e.g. after adding a rule), delete its row from Log
  const done = values.filter(r => r[9] === 'DONE' || r[9] === 'NO_MATCH');
  return {
    records: new Set(done.map(r => r[1])),
    // Video IDs are remembered too: the same meeting is not processed twice even through another host
    files: new Set(done.flatMap(r => String(r[7] || '').split(',').map(x => x.trim()).filter(Boolean))),
  };
}

function logRow_(ss, d) {
  sheet_(ss, CONFIG.SHEETS.log).appendRow([
    new Date(), d.record || '', d.host || '', d.title || '', d.keyword || '', d.folderId || '',
    d.notionDbId || '', d.fileIds || '', d.notionUrl || '', d.status || '', d.error || '',
  ]);
}

function isSkipped_(recordName) {
  return CacheService.getScriptCache().get('skip_' + recordName) !== null;
}

function markSkipped_(recordName) {
  CacheService.getScriptCache().put('skip_' + recordName, '1', 21600); // 6 hours (maximum cache)
}

function alert_(subject, err) {
  const to = prop_('ALERT_EMAIL', false);
  const message = String((err && err.stack) || err);
  Log.error(subject, message);
  if (!to) return;
  const digest = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, subject + message.slice(0, 300)));
  const cache = CacheService.getScriptCache();
  if (cache.get('alert_' + digest)) return; // do not spam with the same error
  cache.put('alert_' + digest, '1', 21600);
  MailApp.sendEmail(to, '[Meet Router] ' + subject, message);
}

/* ======================= HELPERS ======================= */

function prop_(key, required = true) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (required && !v) throw new Error('Script Property not set: ' + key);
  return v;
}

function sheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet "${name}" is missing: run setup()`);
  return sh;
}

/* ======================= SETUP / DIAGNOSTICS ======================= */

function setup() {
  const ss = SpreadsheetApp.openById(prop_('ROUTING_SPREADSHEET_ID'));
  const headers = {
    [CONFIG.SHEETS.routing]: ['keyword', 'folder_id', 'notion_db_id', 'priority', 'active', 'direction', 'meet_type', 'user_id', 'comment'],
    [CONFIG.SHEETS.hosts]: ['email', 'active', 'comment'],
    [CONFIG.SHEETS.log]: ['timestamp', 'conference_record', 'host', 'title', 'keyword', 'folder_id',
      'notion_db_id', 'file_ids', 'notion_url', 'status', 'error'],
  };
  Object.entries(headers).forEach(([name, row]) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(row);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, row.length).setFontWeight('bold');
    }
  });
  Log.info('Sheets are ready');
}

// Check: token for each host, access to Meet/Calendar and to each Notion database
function testAuth() {
  const ss = SpreadsheetApp.openById(prop_('ROUTING_SPREADSHEET_ID'));
  getHosts_(ss).forEach(host => {
    try {
      const since = new Date(Date.now() - 7 * 24 * 3600e3).toISOString();
      const recs = gapi_(host, `${MEET}/conferenceRecords?pageSize=10&filter=${encodeURIComponent(`end_time>="${since}"`)}`);
      gapi_(host, `${CAL}/calendars/primary/events?maxResults=1`);
      Log.info(`✅ ${host}: meetings in 7 days (first page): ${(recs.conferenceRecords || []).length}`);
    } catch (e) {
      Log.info(`❌ ${host}: ${e.message}`);
    }
  });
  const dbs = new Set(getRules_(ss).map(r => r.notionDbId).filter(Boolean));
  const fallbackDb = prop_('FALLBACK_NOTION_DATABASE_ID', false);
  if (fallbackDb) dbs.add(cleanNotionId_(fallbackDb));
  dbs.forEach(id => {
    const res = UrlFetchApp.fetch(`https://api.notion.com/v1/databases/${id}`, {
      headers: { Authorization: 'Bearer ' + prop_('NOTION_API_TOKEN'), 'Notion-Version': CONFIG.NOTION_VERSION },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) { Log.info(`❌ Notion ${id}: ${res.getContentText()}`); return; }
    const schema = JSON.parse(res.getContentText()).properties;
    const { cols, statusValue } = resolveNotionColumns_(schema);
    const title = JSON.parse(res.getContentText()).title.map(t => t.plain_text).join('');
    const problems = [];
    if (!cols.recording) problems.push('no URL column for the link');
    if (!cols.date) problems.push('no date column');
    if (!cols.status) problems.push('no status column');
    else if (!statusValue) problems.push(`status has none of: ${CONFIG.NOTION_STATUS_VALUES.join(', ')}`);
    const dateMode = !cols.date ? '' : CONFIG.NOTION_DATE_WITH_TIME === 'auto'
      ? (notionDateHasTime_(id, cols.date) ? ' (with time)' : ' (date only)')
      : (CONFIG.NOTION_DATE_WITH_TIME ? ' (with time)' : ' (date only)');
    const map = `${cols.title} | ${cols.status || '—'}=${statusValue || '—'} | ${cols.date || '—'}${dateMode} | ${cols.recording || '—'}`;
    Log.info(`${problems.length ? '⚠️' : '✅'} Notion "${title}" (${id}): ${map}${problems.length ? ' — ' + problems.join('; ') : ''}`);
  });
}

function backfill(days = 5, dryRun = true) {
  const saved = { lookback: CONFIG.LOOKBACK_HOURS, dry: CONFIG.DRY_RUN };
  CONFIG.LOOKBACK_HOURS = days * 24;
  CONFIG.DRY_RUN = dryRun;
  try { run(); } finally { CONFIG.LOOKBACK_HOURS = saved.lookback; CONFIG.DRY_RUN = saved.dry; }
}
function backfillDry() { backfill(5, true); }
function backfillReal() { backfill(5, false); }
