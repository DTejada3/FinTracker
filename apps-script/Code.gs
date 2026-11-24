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
  if (e && e.parameter && e.parameter.ping) {
    return out_({ ok: true, message: 'Finance Tracker is running.' });
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
  const apiKey = props.getProperty('API_KEY');
  if (!apiKey || !apiKey.trim()) return { ok: false, error: 'Set API_KEY in Apps Script Project Settings.' };
  if (!req || req.key !== apiKey) return { ok: false, error: 'Invalid API key' };
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) return { ok: false, error: 'Run setup from the Sheet’s Apps Script editor.' };
  if (!['ping', 'getAll', 'batch', 'sync', 'pinResetRequest', 'pinResetVerify', 'migrationImage'].includes(req.action)) {
    return { ok: false, error: 'Unknown action' };
  }
  function finish(result) {
    timings.total = Date.now() - started;
    return { ...result, capabilities: { syncBatch: true, reportExclusion: true }, timings };
  }
  if (req.action === 'ping') {
    try {
      timed_(timings, 'open', () => SpreadsheetApp.openById(spreadsheetId));
      return finish({ ok: true, pong: true });
    } catch (err) {
      return finish({ ok: false, error: String(err && err.message || err) });
    }
  }
  if (req.action === 'migrationImage') {
    try { return finish(exportTrackerImage_(req.id)); }
    catch (err) { return finish({ ok: false, error: String(err && err.message || err) }); }
  }
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    timed_(timings, 'lock', () => lock.waitLock(25000));
    locked = true;
    const ss = timed_(timings, 'open', () => SpreadsheetApp.openById(spreadsheetId));
    if (req.action === 'pinResetRequest') return finish(pinResetRequest_(ss));
    if (req.action === 'pinResetVerify') {
      const result = pinResetVerify_(ss, String(req.code || ''));
      if (result.ok) SpreadsheetApp.flush();
      return finish(result);
    }
    const ops = req.action === 'getAll' ? [] : (req.ops || []);
    if (!Array.isArray(ops)) return finish({ ok: false, error: 'Invalid batch' });
    const result = ops.length ? timed_(timings, 'write', () => {
      const batch = runBatch_(ss, ops, timings);
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
function pinResetRequest_(ss) {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  let prev = null;
  try { prev = JSON.parse(props.getProperty('PIN_RESET') || 'null'); } catch (e) {}
  if (prev && prev.sent && now - prev.sent < 60 * 1000) {
    return { ok: false, error: 'A code was just sent. Please wait a minute before asking for another one.' };
  }
  const email = ownerEmail_(ss);
  if (!email) return { ok: false, error: 'Could not find the email address of the sheet owner.' };
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + now);
  let code = '';
  for (let i = 0; i < 6; i++) code += String((bytes[i] & 255) % 10);
  props.setProperty('PIN_RESET', JSON.stringify({ h: sha_('ft-reset|' + code), exp: now + PIN_RESET_MINUTES * 60 * 1000, tries: 0, sent: now }));
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
function pinResetVerify_(ss, code) {
  const props = PropertiesService.getScriptProperties();
  let st = null;
  try { st = JSON.parse(props.getProperty('PIN_RESET') || 'null'); } catch (e) {}
  if (!st) return { ok: false, error: 'No reset code is active. Ask for a new one.' };
  if (Date.now() > st.exp) { props.deleteProperty('PIN_RESET'); return { ok: false, error: 'That code has expired. Ask for a new one.' }; }
  if (sha_('ft-reset|' + code.replace(/\D/g, '')) !== st.h) {
    st.tries = (st.tries || 0) + 1;
    if (st.tries >= PIN_RESET_MAX_TRIES) { props.deleteProperty('PIN_RESET'); return { ok: false, error: 'Too many wrong codes. Ask for a new one.' }; }
    props.setProperty('PIN_RESET', JSON.stringify(st));
    return { ok: false, error: 'Wrong code. ' + (PIN_RESET_MAX_TRIES - st.tries) + ' tries left.' };
  }
  props.deleteProperty('PIN_RESET');
  // an empty pinHash means "default PIN 000000", which the app forces you to change before use
  runBatch_(ss, [{ op: 'upsert', sheet: 'Settings', row: { id: 'pinHash', value: '' } }]);
  return { ok: true };
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the bound script editor to register the Sheet and prepare its tabs. */
function setup() {
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

function runBatch_(ss, ops, timings = {}) {
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
        const image = row.image && String(row.image).indexOf('data:') === 0
          ? timed_(timings, 'image', () => {
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
  const m = /^data:(.+?);base64,(.+)$/.exec(dataUrl);
  if (!m) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (size > 1000 ? 'receipt_' : 'item_') + id + '.jpg');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w' + (size || 1000);
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
