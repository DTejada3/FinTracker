/**
 * Finance Tracker backend and web app host. See README.md for setup.
 */

const IMAGE_FOLDER = 'FinanceTracker Images';
const PIN_RESET_MINUTES = 10;     // how long an emailed reset code stays valid
const PIN_RESET_MAX_TRIES = 5;    // wrong codes allowed before the code is cancelled

const SHEETS = {
  Accounts:      ['id', 'name', 'type', 'initial', 'creditLimit', 'excluded', 'order', 'archived', 'createdAt', 'color'],
  Categories:    ['id', 'name', 'type', 'order'],
  SubCategories: ['id', 'categoryId', 'name', 'order'],
  Records:       ['id', 'date', 'type', 'accountId', 'toAccountId', 'qty', 'item', 'categoryId', 'subCategoryId',
                  'unitAmount', 'amount', 'payee', 'note', 'image', 'plannedId', 'debtId', 'createdAt',
                  'expiryDate', 'expiryStatus', 'remindOn', 'excludeFromReport'],
  Planned:       ['id', 'name', 'type', 'accountId', 'categoryId', 'subCategoryId', 'amount', 'payee', 'startDate',
                  'nextDue', 'repeat', 'customDays', 'autoConfirm', 'active', 'note', 'toAccountId'],
  Debts:         ['id', 'person', 'direction', 'amount', 'paid', 'date', 'dueDate', 'note', 'accountId', 'status'],
  Templates:     ['id', 'name', 'type', 'accountId', 'toAccountId', 'item', 'categoryId', 'subCategoryId',
                  'qty', 'amount', 'payee', 'note', 'order'],
  Receipts:      ['id', 'title', 'date', 'amount', 'note', 'image', 'recordId', 'createdAt'],
  Settings:      ['id', 'value']
};
const KEYS = {
  accounts: 'Accounts', categories: 'Categories', subcategories: 'SubCategories',
  records: 'Records', planned: 'Planned', debts: 'Debts', templates: 'Templates', receipts: 'Receipts'
};
const NUMS  = new Set(['initial', 'creditLimit', 'order', 'qty', 'unitAmount', 'amount', 'customDays', 'paid']);
const BOOLS = new Set(['excluded', 'archived', 'autoConfirm', 'active', 'excludeFromReport']);
const DATES = new Set(['date', 'startDate', 'nextDue', 'dueDate', 'expiryDate', 'remindOn']);

/** Serves the app itself. Open the /exec link and you get the Finance Tracker page. */
function doGet(e) {
  const props = PropertiesService.getScriptProperties();
  if (e && e.parameter && e.parameter.owner === '1') {
    try {
      requireTrackerOwner_(props);
      if (!trackerOwnerReady_(props)) throw new Error('Managed tracker setup is required.');
      return HtmlService.createHtmlOutputFromFile('Owner').setTitle('Finance Tracker · User management')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err) {
      return HtmlService.createHtmlOutput('<h1>System owner sign-in</h1><p>Open this page in your browser while signed in to the Google account that deployed the tracker. If the dashboard is still unavailable, finish the managed tracker setup in Apps Script.</p>')
        .setTitle('Finance Tracker · System owner');
    }
  }
  if (e && e.parameter && e.parameter.ping) {
    return out_({ ok: true, managedAuth: managedTrackerEnabled_(props), message: 'Finance Tracker is running.' });
  }
  let html;
  try { html = HtmlService.createHtmlOutputFromFile('index'); }
  catch (err) { html = HtmlService.createHtmlOutputFromFile('Index'); }
  return html
    .setTitle('Finance Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Called by the page through google.script.run. Takes and returns a JSON string. */
function rpc(json) {
  let res;
  try { res = handle_(JSON.parse(json)); }
  catch (err) { res = { ok: false, error: String(err && err.message || err) }; }
  return JSON.stringify(res);
}

/** Kept so an old copy of the app that still posts to the URL keeps working. */
function doPost(e) {
  let res;
  try { res = handle_(JSON.parse(e.postData.contents)); }
  catch (err) { res = { ok: false, error: String(err && err.message || err) }; }
  return out_(res);
}

function handle_(req) {
  const timings = {};
  const started = Date.now();
  const props = PropertiesService.getScriptProperties();
  const managed = managedTrackerEnabled_(props);
  if (req?.action === 'accessMode') {
    return { ok: true, managedAuth: managed, ownerUrl: managed ? ScriptApp.getService().getUrl() : null };
  }
  if (!req || !['login', 'resume', 'logout', 'ping', 'getAll', 'batch', 'sync', 'pinResetRequest', 'pinResetVerify', 'migrationImage', 'managedImage'].includes(req.action)) {
    return { ok: false, error: 'Unknown action' };
  }
  let spreadsheetId = props.getProperty('SPREADSHEET_ID');
  if (!managed) {
    if (['login', 'resume', 'logout', 'managedImage'].includes(req.action)) return { ok: false, error: 'Managed sign-in is not available on this deployment.', code: 'setup' };
    const apiKey = props.getProperty('API_KEY');
    if (!apiKey || !apiKey.trim()) return { ok: false, error: 'Set API_KEY in Apps Script Project Settings.' };
    if (req.key !== apiKey) return { ok: false, error: 'Invalid API key' };
    if (!spreadsheetId) return { ok: false, error: 'Run setup from the Sheet’s Apps Script editor.' };
  }
  function finish(result) {
    timings.total = Date.now() - started;
    return { ...result, capabilities: { syncBatch: true, reportExclusion: true }, timings };
  }
  if (!managed && req.action === 'ping') {
    try {
      timed_(timings, 'open', () => SpreadsheetApp.openById(spreadsheetId));
      return finish({ ok: true, pong: true });
    } catch (err) {
      return finish({ ok: false, error: String(err && err.message || err) });
    }
  }
  if (!managed && req.action === 'migrationImage') {
    try { return finish(exportTrackerImage_(req.id)); }
    catch (err) { return finish({ ok: false, error: String(err && err.message || err) }); }
  }
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    timed_(timings, 'lock', () => lock.waitLock(25000));
    locked = true;
    const tokenAuth = managed && (req.sessionToken !== undefined || ['resume', 'logout'].includes(req.action));
    const session = tokenAuth ? authenticateTrackerSession_(props, req.email, req.sessionToken) : null;
    const user = tokenAuth ? session?.user : managed ? authenticateTrackerUser_(props, req.email, req.key) : null;
    if (managed && !user) return finish({ ok: false, error: 'Email address or API key was not accepted.', code: 'auth' });
    if (user) {
      spreadsheetId = user.sheetId;
      if (req.action === 'logout') {
        saveTrackerSessions_(props, user, session.entries.filter(entry => entry !== session.entry));
        return finish({ ok: true });
      }
      if (session && session.entry.expires - Date.now() < TRACKER_SESSION_LIFETIME - 86400000) {
        session.entry.expires = Date.now() + TRACKER_SESSION_LIFETIME;
        saveTrackerSessions_(props, user, session.entries);
      }
      if (req.action === 'ping') return finish({ ok: true, pong: true });
      if (req.action === 'migrationImage') return finish({ ok: false, error: 'Legacy image export is unavailable in managed mode.' });
      if (req.action === 'managedImage') return finish(readManagedImage_(props, user, req.id));
    }
    const ss = timed_(timings, 'open', () => SpreadsheetApp.openById(spreadsheetId));
    if (req.action === 'login' || req.action === 'resume') {
      const result = { ok: true, accountId: sha_(user.sheetId), data: timed_(timings, 'read', () => getAll_(ss)) };
      if (req.action === 'login' && !tokenAuth && req.remember === true) {
        result.sessionToken = issueTrackerSession_(props, user);
      }
      return finish(result);
    }
    if (req.action === 'pinResetRequest') return finish(pinResetRequest_(ss, user?.email));
    if (req.action === 'pinResetVerify') {
      const result = pinResetVerify_(ss, String(req.code || ''), user);
      if (result.ok) SpreadsheetApp.flush();
      return finish(result);
    }
    const ops = req.action === 'getAll' ? [] : (req.ops || []);
    if (!Array.isArray(ops)) return finish({ ok: false, error: 'Invalid batch' });
    const result = ops.length ? timed_(timings, 'write', () => {
      const batch = runBatch_(ss, ops, timings, user);
      SpreadsheetApp.flush();
      return batch;
    }) : { ok: true, errors: [], failedOps: [] };
    if (req.action !== 'batch') {
      try {
        result.data = timed_(timings, 'read', () => getAll_(ss));
      } catch (err) {
        if (req.action !== 'sync') throw err;
        // The writes committed. Let the client acknowledge them even if the read failed.
        result.snapshotError = String(err && err.message || err);
      }
    }
    return finish(result);
  } catch (err) {
    return finish({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (locked) lock.releaseLock();
  }
}

function timed_(timings, key, fn) {
  const started = Date.now();
  try { return fn(); }
  finally { timings[key] = (timings[key] || 0) + Date.now() - started; }
}

/* ---------- PIN reset by email ---------- */

function sha_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

function ownerEmail_(ss) {
  let email = '';
  try { const o = ss.getOwner(); email = o ? o.getEmail() : ''; } catch (e) {}
  if (!email) { try { email = Session.getEffectiveUser().getEmail(); } catch (e) {} }
  return email;
}

function maskEmail_(email) {
  const p = String(email).split('@');
  if (p.length !== 2) return email;
  return p[0].charAt(0) + '***@' + p[1];
}

/** Emails a one-time 6 digit code to the owner of the sheet. */
function pinResetRequest_(ss, emailAddress) {
  const props = PropertiesService.getScriptProperties();
  const resetKey = emailAddress ? 'PIN_RESET_' + ss.getId() : 'PIN_RESET';
  const now = Date.now();
  let prev = null;
  try { prev = JSON.parse(props.getProperty(resetKey) || 'null'); } catch (e) {}
  if (prev && prev.sent && now - prev.sent < 60 * 1000) {
    return { ok: false, error: 'A code was just sent. Please wait a minute before asking for another one.' };
  }
  const email = emailAddress || ownerEmail_(ss);
  if (!email) return { ok: false, error: 'Could not find the email address of the sheet owner.' };
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + now);
  let code = '';
  for (let i = 0; i < 6; i++) code += String((bytes[i] & 255) % 10);
  props.setProperty(resetKey, JSON.stringify({ h: sha_('ft-reset|' + code), exp: now + PIN_RESET_MINUTES * 60 * 1000, tries: 0, sent: now }));
  MailApp.sendEmail({
    to: email,
    subject: 'Finance Tracker PIN reset code',
    body: 'Your Finance Tracker PIN reset code is: ' + code + '\n\n' +
          'It expires in ' + PIN_RESET_MINUTES + ' minutes. After you enter it, your PIN goes back to 000000 and you will be asked to choose a new one.\n\n' +
          'If you did not ask for this, you can ignore this email. Your PIN has not been changed.'
  });
  return { ok: true, sentTo: maskEmail_(email), minutes: PIN_RESET_MINUTES };
}

/** Checks the emailed code. On success the PIN goes back to the default and must be changed. */
function pinResetVerify_(ss, code, user) {
  const props = PropertiesService.getScriptProperties();
  const resetKey = user ? 'PIN_RESET_' + ss.getId() : 'PIN_RESET';
  let st = null;
  try { st = JSON.parse(props.getProperty(resetKey) || 'null'); } catch (e) {}
  if (!st) return { ok: false, error: 'No reset code is active. Ask for a new one.' };
  if (Date.now() > st.exp) { props.deleteProperty(resetKey); return { ok: false, error: 'That code has expired. Ask for a new one.' }; }
  if (sha_('ft-reset|' + code.replace(/\D/g, '')) !== st.h) {
    st.tries = (st.tries || 0) + 1;
    if (st.tries >= PIN_RESET_MAX_TRIES) { props.deleteProperty(resetKey); return { ok: false, error: 'Too many wrong codes. Ask for a new one.' }; }
    props.setProperty(resetKey, JSON.stringify(st));
    return { ok: false, error: 'Wrong code. ' + (PIN_RESET_MAX_TRIES - st.tries) + ' tries left.' };
  }
  props.deleteProperty(resetKey);
  // an empty pinHash means "default PIN 000000", which the app forces you to change before use
  runBatch_(ss, [{ op: 'upsert', sheet: 'Settings', row: { id: 'pinHash', value: '' } }]);
  return { ok: true };
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the bound script editor to register the Sheet and prepare its tabs. */
function setup() {
  requireTrackerOwner_(PropertiesService.getScriptProperties());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the target Sheet and run setup from Extensions > Apps Script.');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    ensureSheets_(ss);
    SpreadsheetApp.flush();
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  } finally {
    lock.releaseLock();
  }
}

function ensureSheets_(ss) {
  Object.keys(SHEETS).forEach(name => sheetFor_(ss, name, true));
  const dflt = ss.getSheetByName('Sheet1');
  if (dflt && ss.getSheets().length > 1 && dflt.getLastRow() === 0) ss.deleteSheet(dflt);
}

function sheetFor_(ss, name, validateHeader) {
  const cols = SHEETS[name];
  const existing = ss.getSheetByName(name);
  if (existing) {
    const width = existing.getMaxColumns();
    if (width < cols.length) existing.insertColumnsAfter(width, cols.length - width);
    if (validateHeader) migrateHeader_(existing, name, existing.getRange(1, 1, 1, cols.length).getValues()[0]);
    return existing;
  }
  const sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#e8f0fe');
  sh.setFrozenRows(1);
  const rows = sh.getMaxRows() - 1;
  cols.forEach(function (c, i) {
    if (!NUMS.has(c) && !BOOLS.has(c)) sh.getRange(2, i + 1, rows, 1).setNumberFormat('@');
  });
  return sh;
}

/** Adds any newly introduced columns (e.g. SubCategories.order) to sheets created by an older version. */
function migrateHeader_(sh, name, header) {
  const cols = SHEETS[name];
  const hdr = header.map(String);
  if (cols.every((col, i) => hdr[i] === col)) return;
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#e8f0fe');
  const rows = Math.max(1, sh.getMaxRows() - 1);
  cols.forEach(function (c, i) {
    if (hdr[i] !== c && !NUMS.has(c) && !BOOLS.has(c)) sh.getRange(2, i + 1, rows, 1).setNumberFormat('@');
  });
}

function readSheet_(ss, name, tz) {
  const sh = sheetFor_(ss, name, false);
  const cols = SHEETS[name];
  const vals = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), cols.length).getValues();
  migrateHeader_(sh, name, vals.shift());
  const rows = [];
  vals.forEach(function (v) {
    if (v[0] === '' || v[0] === null) return;
    const o = {};
    cols.forEach(function (c, i) {
      let x = v[i];
      if (DATES.has(c)) x = (x instanceof Date) ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : String(x || '');
      else if (NUMS.has(c)) x = Number(x) || 0;
      else if (BOOLS.has(c)) x = (x === true || String(x).toLowerCase() === 'true');
      else x = (x instanceof Date) ? x.toISOString() : String(x === null || x === undefined ? '' : x);
      if (c === 'excludeFromReport' && !x) return;
      o[c] = x;
    });
    rows.push(o);
  });
  return rows;
}

function getAll_(ss) {
  const tz = Session.getScriptTimeZone();
  const data = {};
  Object.keys(KEYS).forEach(function (k) { data[k] = readSheet_(ss, KEYS[k], tz); });
  const settings = {};
  readSheet_(ss, 'Settings', tz).forEach(function (r) { settings[r.id] = r.value; });
  data.settings = settings;
  return data;
}

function runBatch_(ss, ops, timings = {}, user) {
  const cache = Object.create(null); // sheet name -> { sh, ids, last }
  const errors = [];
  const failedOps = [];
  let imageFolder;

  function info(name) {
    if (cache[name]) return cache[name];
    const sh = sheetFor_(ss, name, true);
    const last = sh.getLastRow();
    const ids = Object.create(null);
    if (last >= 2) {
      sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { if (r[0] !== '') ids[String(r[0])] = i + 2; });
    }
    cache[name] = { sh: sh, ids: ids, last: Math.max(last, 1) };
    return cache[name];
  }

  ops.forEach(function (op, index) {
    try {
      if (!Object.prototype.hasOwnProperty.call(SHEETS, op.sheet)) throw new Error('Unknown sheet ' + op.sheet);
      const c = info(op.sheet);
      const cols = SHEETS[op.sheet];
      if (op.op === 'upsert') {
        const row = op.row;
        const image = row.image && ['Records', 'Receipts'].includes(op.sheet) && String(row.image).indexOf('data:') === 0
          ? timed_(timings, 'image', () => {
            if (user) {
              if (!imageFolder) imageFolder = managedImageFolder_(PropertiesService.getScriptProperties(), user, true);
              return 'managed:' + imageFolder.createFile(trackerImageBlob_(row.image, row.id, op.sheet === 'Receipts' ? 2000 : 1000)).getId();
            }
            if (!imageFolder) {
              const folders = DriveApp.getFoldersByName(IMAGE_FOLDER);
              imageFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(IMAGE_FOLDER);
            }
            return saveImage_(imageFolder, row.image, row.id, op.sheet === 'Receipts' ? 2000 : 1000);
          })
          : row.image;
        const values = cols.map(function (col) {
          const v = col === 'image' ? image : row[col];
          if (v === undefined || v === null) return '';
          if (BOOLS.has(col)) return !!v;
          // Sheets interprets an initial "=" as a formula, even for text cells.
          if (user && typeof v === 'string' && v.startsWith('=')) return "'" + v;
          return v;
        });
        const id = String(row.id);
        if (c.ids[id]) {
          // Older clients omit the report flag. Preserve the existing column on their edits.
          if (op.sheet === 'Records' && typeof row.excludeFromReport !== 'boolean') values.pop();
          c.sh.getRange(c.ids[id], 1, 1, values.length).setValues([values]);
        } else {
          c.last += 1;
          c.sh.getRange(c.last, 1, 1, cols.length).setValues([values]);
          c.ids[id] = c.last;
        }
      } else if (op.op === 'delete') {
        const id = String(op.id);
        if (c.ids[id]) {
          c.sh.deleteRow(c.ids[id]);
          delete cache[op.sheet];   // row numbers shifted, rebuild next time
        }
      }
    } catch (err) {
      failedOps.push(index);
      errors.push((op && op.sheet) + ': ' + (err && err.message || err));
    }
  });
  return { ok: true, errors, failedOps };
}

function saveImage_(folder, dataUrl, id, size) {
  const file = folder.createFile(trackerImageBlob_(dataUrl, id, size));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w' + (size || 1000);
}

function trackerImageBlob_(dataUrl, id, size) {
  const m = /^data:(.+?);base64,(.+)$/.exec(dataUrl);
  if (!m || !/^image\/(?:jpeg|png|webp|gif)$/.test(m[1]) || m[2].length > 13333336) throw new Error('Invalid receipt image.');
  return Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (size > 1000 ? 'receipt_' : 'item_') + id + '.jpg');
}

function exportTrackerImage_(id) {
  if (typeof id !== 'string' || !/^[-\w]+$/.test(id)) throw new Error('Invalid legacy image reference.');
  const folders = DriveApp.getFoldersByName(IMAGE_FOLDER);
  if (!folders.hasNext()) throw new Error('The legacy receipt folder could not be found.');
  const folderId = folders.next().getId();
  const file = DriveApp.getFileById(id);
  const parents = file.getParents();
  let belongsToTracker = false;
  while (parents.hasNext()) if (parents.next().getId() === folderId) { belongsToTracker = true; break; }
  const type = file.getMimeType();
  if (!belongsToTracker || file.isTrashed() || !/^image\/(?:jpeg|png|webp|gif)$/.test(type) || file.getSize() > 10000000) {
    throw new Error('This image could not be exported from the legacy tracker folder.');
  }
  return { ok: true, contentType: type, base64: Utilities.base64Encode(file.getBlob().getBytes()) };
}

const TRACKER_USERS = 'Users';
const TRACKER_FIELDS = {
  email: ['Email', 'Email Address'],
  sheetId: ['Spreadsheet ID', 'Sheet ID', 'Spreadsheet URL', 'Sheet URL'],
  nonce: ['Key nonce'],
  state: ['Ledger status']
};

function managedTrackerEnabled_(props) {
  return props.getProperty('TRACKER_MANAGED') === 'true';
}

function trackerOwnerReady_(props) {
  return !!(props.getProperty('TRACKER_OWNER_EMAIL') && props.getProperty('TRACKER_REGISTRY_ID') &&
    props.getProperty('TRACKER_KEY_SECRET') && props.getProperty('TRACKER_FOLDER_ID'));
}

function trackerEmail_(value) {
  const email = typeof value === 'string' ? value.trim().toLowerCase() : '';
  return email.length <= 254 && /^[a-z0-9][a-z0-9._%+'-]*@[a-z0-9][a-z0-9.-]*\.[a-z]{2,}$/i.test(email) ? email : '';
}

function requireTrackerOwner_(props) {
  const effective = trackerEmail_(Session.getEffectiveUser().getEmail());
  const active = trackerEmail_(Session.getActiveUser().getEmail());
  const owner = trackerEmail_(props.getProperty('TRACKER_OWNER_EMAIL')) || effective;
  if (!active || active !== effective || active !== owner) throw new Error('The deployment owner must sign in with Google.');
  return owner;
}

function trackerFolder_(props) {
  const id = props.getProperty('TRACKER_FOLDER_ID');
  if (!id) throw new Error('The tracker folder is not configured.');
  return DriveApp.getFolderById(id);
}

function requirePrivateDriveItem_(item, owner) {
  if (trackerEmail_(item.getOwner()?.getEmail()) !== owner || item.getSharingAccess() !== DriveApp.Access.PRIVATE ||
      item.getEditors().some(u => trackerEmail_(u.getEmail()) !== owner) ||
      item.getViewers().some(u => trackerEmail_(u.getEmail()) !== owner)) {
    throw new Error('Use a private Google Drive folder and registry owned by the deployment owner.');
  }
}

function trackerSheetId_(value) {
  const raw = String(value || '').trim();
  const id = raw.match(/^https:\/\/docs\.google\.com\/spreadsheets\/d\/([-\w]+)(?:[/?#]|$)/)?.[1] || raw;
  return /^[-\w]{10,}$/.test(id) ? id : '';
}

function trackerRegistry_(props, initialize) {
  const id = props.getProperty('TRACKER_REGISTRY_ID');
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(TRACKER_USERS);
  if (!sh && initialize) sh = ss.insertSheet(TRACKER_USERS);
  if (!sh) throw new Error('The registry must have a Users tab.');
  let header = sh.getLastRow() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  if (!header.length && initialize) {
    header = Object.values(TRACKER_FIELDS).map(names => names[0]);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);
  }
  const normalize = v => String(v).toLowerCase().replace(/[^a-z0-9]/g, '');
  const cols = {};
  for (const [field, names] of Object.entries(TRACKER_FIELDS)) {
    const allowed = names.map(normalize);
    const matches = header.flatMap((name, i) => allowed.includes(normalize(name)) ? [i] : []);
    if (matches.length > 1) throw new Error('The Users tab has duplicate ' + names[0] + ' columns.');
    if (!matches.length) {
      if (!initialize || field === 'email' || field === 'sheetId') throw new Error('The Users tab needs a ' + names[0] + ' column.');
      if (header.length >= sh.getMaxColumns()) sh.insertColumnAfter(header.length);
      sh.getRange(1, header.length + 1).setValue(names[0]);
      matches.push(header.length);
      header.push(names[0]);
    }
    cols[field] = matches[0];
  }
  const users = new Map(), sheetIds = new Set();
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, header.length).getValues() : [];
  rows.forEach((cells, i) => {
    if (cells.every(v => v === '')) return;
    const email = trackerEmail_(cells[cols.email]), sheetId = trackerSheetId_(cells[cols.sheetId]);
    const nonce = String(cells[cols.nonce] || '');
    if (!email || !sheetId || sheetId === id || users.has(email) || sheetIds.has(sheetId) || (nonce && !/^[-\w]{16,80}$/.test(nonce))) {
      throw new Error('Invalid or duplicate user or spreadsheet on Users row ' + (i + 2) + '.');
    }
    sheetIds.add(sheetId);
    users.set(email, { email, sheetId, nonce, state: cells[cols.state] === 'new' ? 'new' : 'existing', row: i + 2 });
  });
  return { sh, cols, users, sheetIds, width: header.length };
}

function appendTrackerUser_(registry, user) {
  const row = Array(registry.width).fill('');
  for (const field of Object.keys(TRACKER_FIELDS)) row[registry.cols[field]] = user[field];
  registry.sh.appendRow(row);
}

function trackerUserKey_(props, user) {
  const secret = props.getProperty('TRACKER_KEY_SECRET');
  if (!secret || !user.nonce) throw new Error('Managed tracker keys need setup.');
  const bytes = Utilities.computeHmacSha256Signature(JSON.stringify(['tracker-user-v1', user.email, user.nonce]), secret);
  return 'ftk_' + Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function authenticateTrackerUser_(props, email, key) {
  const normalized = trackerEmail_(email);
  if (!normalized || typeof key !== 'string' || !/^ftk_[-\w]{43}$/.test(key)) return null;
  const user = trackerRegistry_(props, false).users.get(normalized);
  if (!user?.nonce) return null;
  const expected = trackerUserKey_(props, user);
  let mismatched = 0;
  for (let i = 0; i < expected.length; i++) mismatched |= expected.charCodeAt(i) ^ key.charCodeAt(i);
  return mismatched ? null : user;
}

const TRACKER_SESSION_LIFETIME = 30 * 86400000;
const TRACKER_MAX_SESSIONS = 8;
const trackerSessionsKey_ = user => 'TRACKER_SESSIONS_' + sha_(user.email);

// Session reads and writes run under the same lock as ledger writes and key rotation.
function trackerSessions_(props, user) {
  let saved;
  try { saved = JSON.parse(props.getProperty(trackerSessionsKey_(user)) || 'null'); }
  catch (err) { return []; }
  if (saved?.nonce !== user.nonce || saved?.sheetId !== user.sheetId || !Array.isArray(saved.entries)) return [];
  return saved.entries.filter(entry => entry.expires > Date.now());
}

function saveTrackerSessions_(props, user, entries) {
  const key = trackerSessionsKey_(user);
  if (!entries.length) props.deleteProperty(key);
  else props.setProperty(key, JSON.stringify({ nonce: user.nonce, sheetId: user.sheetId, entries }));
}

function issueTrackerSession_(props, user) {
  const token = 'fts_' + (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').toLowerCase();
  const active = trackerSessions_(props, user).sort((a, b) => b.expires - a.expires);
  const entry = { digest: sha_(token), expires: Date.now() + TRACKER_SESSION_LIFETIME };
  saveTrackerSessions_(props, user, [entry, ...active.slice(0, TRACKER_MAX_SESSIONS - 1)]);
  return token;
}

function authenticateTrackerSession_(props, email, token) {
  const normalized = trackerEmail_(email);
  if (!normalized || typeof token !== 'string' || !/^fts_[a-f0-9]{64}$/.test(token)) return null;
  const user = trackerRegistry_(props, false).users.get(normalized);
  if (!user?.nonce) return null;
  const entries = trackerSessions_(props, user), digest = sha_(token);
  const entry = entries.find(item => item.digest === digest);
  return entry ? { user, entry, entries } : null;
}

/** Run from the Apps Script editor as the account that deploys the tracker. */
function setupManagedTracker() {
  const props = PropertiesService.getScriptProperties();
  const owner = requireTrackerOwner_(props);
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const legacyId = props.getProperty('SPREADSHEET_ID');
    const legacyEmail = trackerEmail_(props.getProperty('TRACKER_LEGACY_EMAIL'));
    if (legacyId && !props.getProperty('TRACKER_REGISTRY_ID') && !legacyEmail) {
      throw new Error('Set TRACKER_LEGACY_EMAIL to the email that owns the existing tracker ledger.');
    }
    if (!props.getProperty('TRACKER_FOLDER_ID')) {
      props.setProperty('TRACKER_FOLDER_ID', DriveApp.createFolder('Financial Tracker Account').getId());
    }
    const folder = trackerFolder_(props);
    requirePrivateDriveItem_(folder, owner);
    if (!props.getProperty('TRACKER_REGISTRY_ID')) {
      const registry = SpreadsheetApp.create('Finance Tracker User Registry');
      const file = DriveApp.getFileById(registry.getId());
      file.moveTo(folder);
      props.setProperty('TRACKER_REGISTRY_ID', registry.getId());
    }
    requirePrivateDriveItem_(DriveApp.getFileById(props.getProperty('TRACKER_REGISTRY_ID')), owner);
    const registry = trackerRegistry_(props, true);
    if (legacyId && !registry.sheetIds.has(legacyId) && !legacyEmail) {
      throw new Error('Register the current tracker spreadsheet in Users or set TRACKER_LEGACY_EMAIL.');
    }
    const existingLedgers = new Set(registry.sheetIds);
    if (legacyId) existingLedgers.add(legacyId);
    existingLedgers.forEach(id => {
      const ss = SpreadsheetApp.openById(id);
      if (!ss.getSheetByName('Records') || !ss.getSheetByName('Accounts')) {
        throw new Error('Every existing ledger must have the tracker Records and Accounts tabs.');
      }
    });
    if (!props.getProperty('TRACKER_KEY_SECRET')) {
      if ([...registry.users.values()].some(user => user.nonce)) throw new Error('Restore TRACKER_KEY_SECRET before using an existing managed registry.');
      props.setProperty('TRACKER_KEY_SECRET', Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid());
    }
    if (legacyId && legacyEmail && registry.users.has(legacyEmail) && registry.users.get(legacyEmail).sheetId !== legacyId) {
      throw new Error('The legacy email is already assigned to a different sheet.');
    }
    if (legacyId && legacyEmail && !registry.sheetIds.has(legacyId)) {
      appendTrackerUser_(registry, { email: legacyEmail, sheetId: legacyId, nonce: Utilities.getUuid(), state: 'existing' });
    }
    registry.users.forEach(user => {
      if (!user.nonce) registry.sh.getRange(user.row, registry.cols.nonce + 1).setValue(Utilities.getUuid());
    });
    SpreadsheetApp.flush();
    props.setProperty('TRACKER_OWNER_EMAIL', owner);
  } finally {
    lock.releaseLock();
  }
}

/** Run from the editor after the updated deployment and new user keys are ready. */
function enableManagedTracker() {
  const props = PropertiesService.getScriptProperties();
  requireTrackerOwner_(props);
  if (!trackerOwnerReady_(props)) throw new Error('Run setupManagedTracker first.');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const { users } = trackerRegistry_(props, false);
    if (!users.size || [...users.values()].some(user => !user.nonce)) throw new Error('Register users and run setupManagedTracker first.');
    props.setProperty('TRACKER_MANAGED', 'true');
  } finally {
    lock.releaseLock();
  }
}

/** Only available to the deployment owner's verified Google session. */
function ownerRpc(action, email) {
  const props = PropertiesService.getScriptProperties();
  try { requireTrackerOwner_(props); }
  catch (err) { return { ok: false, error: 'The deployment owner must sign in with Google.' }; }
  if (!trackerOwnerReady_(props)) return { ok: false, error: 'Managed tracker setup is required.' };
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    lock.waitLock(25000);
    locked = true;
    const registry = trackerRegistry_(props, false);
    if (action === 'list') return { ok: true, users: Array.from(registry.users.values(), user => ({
      email: user.email, state: user.state
    })).sort((a, b) => a.email.localeCompare(b.email)) };
    email = trackerEmail_(email);
    if (!email) return { ok: false, error: 'Enter a valid email address.' };
    let user = registry.users.get(email);
    if (action === 'add') {
      if (user) return { ok: false, error: 'This email is already registered.' };
      const folder = trackerFolder_(props);
      requirePrivateDriveItem_(folder, props.getProperty('TRACKER_OWNER_EMAIL'));
      const provisioning = 'TRACKER_PROVISION_' + sha_(email);
      const pending = props.getProperty(provisioning);
      const ss = pending ? SpreadsheetApp.openById(pending) : SpreadsheetApp.create('Finance Tracker Ledger');
      if (!pending) props.setProperty(provisioning, ss.getId());
      const file = DriveApp.getFileById(ss.getId());
      file.moveTo(folder);
      ensureSheets_(ss);
      user = { email, sheetId: ss.getId(), nonce: Utilities.getUuid(), state: 'new' };
      appendTrackerUser_(registry, user);
      SpreadsheetApp.flush();
      props.deleteProperty(provisioning);
    } else {
      if (!user) return { ok: false, error: 'This email is no longer registered.' };
      if (action === 'regenerate') {
        user.nonce = Utilities.getUuid();
        registry.sh.getRange(user.row, registry.cols.nonce + 1).setValue(user.nonce);
        props.deleteProperty(trackerSessionsKey_(user));
      } else if (action !== 'key') return { ok: false, error: 'Unknown action.' };
    }
    SpreadsheetApp.flush();
    return { ok: true, key: trackerUserKey_(props, user) };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  } finally {
    if (locked) lock.releaseLock();
  }
}

function managedImageFolder_(props, user, create) {
  const root = trackerFolder_(props);
  const name = 'receipts-' + user.sheetId;
  const matches = root.getFoldersByName(name);
  if (matches.hasNext()) return matches.next();
  if (!create) throw new Error('Receipt image is unavailable.');
  return root.createFolder(name);
}

function readManagedImage_(props, user, id) {
  if (typeof id !== 'string' || !/^[-\w]{10,}$/.test(id)) throw new Error('Receipt image is unavailable.');
  const folderId = managedImageFolder_(props, user, false).getId();
  const file = DriveApp.getFileById(id);
  const parents = file.getParents();
  const type = file.getMimeType();
  while (parents.hasNext()) {
    if (parents.next().getId() === folderId && !file.isTrashed() &&
        /^image\/(?:jpeg|png|webp|gif)$/.test(type) && file.getSize() <= 10000000) {
      return { ok: true, contentType: type, base64: Utilities.base64Encode(file.getBlob().getBytes()) };
    }
  }
  throw new Error('Receipt image is unavailable.');
}
