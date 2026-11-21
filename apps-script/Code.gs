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
                  'expiryDate', 'expiryStatus', 'remindOn'],
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
const BOOLS = new Set(['excluded', 'archived', 'autoConfirm', 'active']);
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
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('API_KEY');
  if (!apiKey || !apiKey.trim()) return { ok: false, error: 'Set API_KEY in Apps Script Project Settings.' };
  if (!req || req.key !== apiKey) return { ok: false, error: 'Invalid API key' };
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) return { ok: false, error: 'Run setup from the Sheet’s Apps Script editor.' };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const ss = SpreadsheetApp.openById(spreadsheetId);
    ensureSheets_(ss);
    if (req.action === 'ping')            return { ok: true, pong: true };
    if (req.action === 'getAll')          return { ok: true, data: getAll_(ss) };
    if (req.action === 'batch')           return runBatch_(ss, req.ops || []);
    if (req.action === 'pinResetRequest') return pinResetRequest_(ss);
    if (req.action === 'pinResetVerify')  return pinResetVerify_(ss, String(req.code || ''));
    return { ok: false, error: 'Unknown action' };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
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
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (sh) { migrateHeader_(sh, name); return; }
    sh = ss.insertSheet(name);
    const cols = SHEETS[name];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#e8f0fe');
    sh.setFrozenRows(1);
    // keep ids, dates and text as plain text so Sheets never converts them
    cols.forEach(function (c, i) {
      if (!NUMS.has(c) && !BOOLS.has(c)) sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    });
  });
  const dflt = ss.getSheetByName('Sheet1');
  if (dflt && ss.getSheets().length > 1 && dflt.getLastRow() === 0) ss.deleteSheet(dflt);
}

/** Adds any newly introduced columns (e.g. SubCategories.order) to sheets created by an older version. */
function migrateHeader_(sh, name) {
  const cols = SHEETS[name];
  if (sh.getMaxColumns() < cols.length) sh.insertColumnsAfter(sh.getMaxColumns(), cols.length - sh.getMaxColumns());
  const hdr = sh.getRange(1, 1, 1, cols.length).getValues()[0].map(String);
  if (hdr.join('|') === cols.join('|')) return;
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#e8f0fe');
  cols.forEach(function (c, i) {
    if (hdr[i] !== c && !NUMS.has(c) && !BOOLS.has(c)) sh.getRange(2, i + 1, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat('@');
  });
}

function readSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  const cols = SHEETS[name];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const tz = Session.getScriptTimeZone();
  const vals = sh.getRange(2, 1, last - 1, cols.length).getValues();
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
      o[c] = x;
    });
    rows.push(o);
  });
  return rows;
}

function getAll_(ss) {
  const data = {};
  Object.keys(KEYS).forEach(function (k) { data[k] = readSheet_(ss, KEYS[k]); });
  const settings = {};
  readSheet_(ss, 'Settings').forEach(function (r) { settings[r.id] = r.value; });
  data.settings = settings;
  return data;
}

function runBatch_(ss, ops) {
  const cache = {};   // sheet name -> { sh, ids: {id: rowNumber}, last }
  const errors = [];

  function info(name) {
    if (cache[name]) return cache[name];
    const sh = ss.getSheetByName(name);
    const last = sh.getLastRow();
    const ids = {};
    if (last >= 2) {
      sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { if (r[0] !== '') ids[String(r[0])] = i + 2; });
    }
    cache[name] = { sh: sh, ids: ids, last: Math.max(last, 1) };
    return cache[name];
  }

  ops.forEach(function (op) {
    try {
      if (!SHEETS[op.sheet]) throw new Error('Unknown sheet ' + op.sheet);
      const c = info(op.sheet);
      const cols = SHEETS[op.sheet];
      if (op.op === 'upsert') {
        const row = op.row;
        if (row.image && String(row.image).indexOf('data:') === 0) row.image = saveImage_(row.image, row.id, op.sheet === 'Receipts' ? 2000 : 1000);
        const values = cols.map(function (col) {
          const v = row[col];
          if (v === undefined || v === null) return '';
          if (BOOLS.has(col)) return !!v;
          return v;
        });
        const id = String(row.id);
        if (c.ids[id]) {
          c.sh.getRange(c.ids[id], 1, 1, cols.length).setValues([values]);
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
      errors.push(op.sheet + ': ' + (err && err.message || err));
    }
  });
  return { ok: true, errors: errors };
}

function saveImage_(dataUrl, id, size) {
  const m = /^data:(.+?);base64,(.+)$/.exec(dataUrl);
  if (!m) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (size > 1000 ? 'receipt_' : 'item_') + id + '.jpg');
  const it = DriveApp.getFoldersByName(IMAGE_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(IMAGE_FOLDER);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w' + (size || 1000);
}
