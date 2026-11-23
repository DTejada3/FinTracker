(function (root) {
  'use strict';

  const DRIVE = 'https://www.googleapis.com/drive/v3';
  const SHEETS = 'https://sheets.googleapis.com/v4/spreadsheets';
  const SCOPE = 'https://www.googleapis.com/auth/drive.file';
  const FOLDER = 'Financial Tracker Account';
  const TRACKER = 'Finance Tracker Ledger';
  const FOLDER_TYPE = 'application/vnd.google-apps.folder';
  const SHEET_TYPE = 'application/vnd.google-apps.spreadsheet';
  const TABS = { Accounts: 'accounts', Categories: 'categories', SubCategories: 'subcategories', Records: 'records',
    Planned: 'planned', Debts: 'debts', Templates: 'templates', Receipts: 'receipts' };
  const HEADERS = ['Change ID', 'Table', 'Record ID', 'Action', 'Data'];
  const RANGE = "'Ledger'!A:E";
  const quote = value => "'" + String(value).replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'";

  class GoogleError extends Error {
    constructor(message, status) { super(message); this.name = 'GoogleError'; this.status = status; }
  }

  async function digest(text) {
    const bytes = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
    return Array.from(new Uint8Array(bytes), value => value.toString(16).padStart(2, '0')).join('');
  }

  function canonical(value) {
    if (Array.isArray(value)) return value.map(canonical);
    if (value && typeof value === 'object') {
      return Object.fromEntries(Object.keys(value).sort().map(key => [key, canonical(value[key])]));
    }
    return value;
  }
  const same = (left, right) => JSON.stringify(canonical(left)) === JSON.stringify(canonical(right));

  function decode(rows, ledger) {
    if (!HEADERS.every((value, index) => rows[0]?.[index] === value)) {
      throw new Error('The Finance Tracker ledger has an unexpected header.');
    }
    const seen = new Set(), latest = new Map();
    for (const [event, table, id, action, ...payload] of rows.slice(1)) {
      if (!event) continue;
      if (seen.has(event)) continue;
      if (!id || !['upsert', 'delete'].includes(action) ||
          !(Object.prototype.hasOwnProperty.call(TABS, table) || table === 'Settings' || table === 'Migration')) {
        throw new Error('The Finance Tracker ledger contains an invalid entry.');
      }
      const value = action === 'upsert' ? JSON.parse(payload.join('')) : null;
      if (action === 'upsert' && (!value || typeof value !== 'object' || String(value.id) !== String(id))) {
        throw new Error('The Finance Tracker ledger contains an invalid record ID.');
      }
      seen.add(event);
      latest.set(table + '\0' + id, { table, id, action, value });
    }
    const data = Object.fromEntries(Object.values(TABS).map(key => [key, []]));
    data.settings = Object.create(null);
    for (const {table, id, action, value} of latest.values()) {
      if (action === 'delete') continue;
      if (table === 'Settings') data.settings[id] = value.value;
      else if (table !== 'Migration') data[TABS[table]].push(value);
    }
    return { data, seen, latest, ledger };
  }

  function rowFor(op) {
    const value = op.op === 'upsert' ? JSON.stringify(op.row) : '';
    // Sheets permits 50,000 characters per cell. Split long notes without breaking Unicode.
    const parts = value.match(/[\s\S]{1,22000}/gu) || [];
    return [op.event, op.sheet, String(op.row?.id ?? op.id), op.op, ...parts];
  }

  class LedgerStore {
    constructor(token, onRequest = () => {}) {
      this.token = token;
      this.onRequest = onRequest;
    }

    async request(url, options = {}, type = 'json') {
      const started = performance.now();
      let ok = false;
      try {
        const result = await fetch(url, {
          ...options, credentials: 'omit', redirect: 'error', cache: 'no-store',
          signal: AbortSignal.timeout(45000),
          headers: { ...options.headers, Authorization: 'Bearer ' + this.token }
        });
        if (result.status === 401) throw new GoogleError('Your Google session expired. Sign in again to sync.', 401);
        if (!result.ok) throw new GoogleError('Google could not complete the request (HTTP ' + result.status + ').', result.status);
        ok = true;
        return type === 'blob' ? await result.blob() : await result.json();
      } finally {
        this.onRequest(url.startsWith(DRIVE) || url.includes('/upload/drive/') ? 'Google Drive' : 'Google Sheets',
          performance.now() - started, ok ? 'ok' : 'failed');
      }
    }

    json(url, method, data) {
      return this.request(url, { method, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) });
    }

    async list(query) {
      const result = await this.request(DRIVE + '/files?' + new URLSearchParams({
        q: "trashed = false and 'me' in owners and " + query,
        fields: 'nextPageToken,files(id,name,parents)', orderBy: 'createdTime', pageSize: '2'
      }));
      if (result.files.length > 1 || result.nextPageToken) throw new Error('More than one matching Finance Tracker file exists in Drive. Resolve the duplicate before syncing.');
      return result.files[0];
    }

    async open(expected) {
      const { user } = await this.request(DRIVE + '/about?fields=user(permissionId,emailAddress,displayName)');
      if (!user?.permissionId) throw new Error('Google did not return an account identifier.');
      if (expected && expected.accountId !== user.permissionId) {
        throw new Error('Choose the Google account associated with this tracker on this device.');
      }
      const folderQuery = 'mimeType = ' + quote(FOLDER_TYPE) + ' and name = ' + quote(FOLDER);
      let folder = await this.list(folderQuery);
      if (expected && (!folder || folder.id !== expected.folderId)) {
        throw new Error('The saved tracker folder is unavailable. Restore it in Google Drive and try again.');
      }
      if (!folder) folder = await this.json(DRIVE + '/files?fields=id', 'POST', { name: FOLDER, mimeType: FOLDER_TYPE });
      const fileQuery = 'mimeType = ' + quote(SHEET_TYPE) + ' and name = ' + quote(TRACKER) + ' and ' + quote(folder.id) + ' in parents';
      let file = await this.list(fileQuery);
      if (expected && (!file || file.id !== expected.spreadsheetId)) {
        throw new Error('The saved tracker spreadsheet is unavailable. Restore it in Google Drive and try again.');
      }
      if (!file) {
        file = await this.json(DRIVE + '/files?fields=id', 'POST', { name: TRACKER, mimeType: SHEET_TYPE, parents: [folder.id] });
      }
      this.folderId = folder.id;
      this.sheetURL = SHEETS + '/' + encodeURIComponent(file.id);
      await this.prepare();
      return { accountId: user.permissionId, email: user.emailAddress || '', spreadsheetId: file.id, folderId: folder.id };
    }

    async prepare() {
      const info = await this.request(this.sheetURL + '?fields=sheets(properties(sheetId,title)');
      const tabs = info.sheets.map(sheet => sheet.properties);
      if (!tabs.some(tab => tab.title === 'Ledger')) {
        try {
          await this.json(this.sheetURL + ':batchUpdate', 'POST', { requests: [{ addSheet: { properties: { title: 'Ledger', gridProperties: { frozenRowCount: 1 } } } }] });
        } catch (error) {
          // A second device may have initialized the same spreadsheet.
          const fresh = await this.request(this.sheetURL + '?fields=sheets(properties(title))');
          if (!fresh.sheets.some(sheet => sheet.properties.title === 'Ledger')) throw error;
        }
      }
      const range = this.sheetURL + '/values/' + encodeURIComponent("'Ledger'!A1:E1");
      const result = await this.request(range);
      if (!result.values?.length) await this.json(range + '?valueInputOption=RAW', 'PUT', { values: [HEADERS] });
      else decode(result.values);
    }

    async read() {
      const result = await this.request(this.sheetURL + '/values/' + encodeURIComponent("'Ledger'") + '?valueRenderOption=UNFORMATTED_VALUE');
      return decode(result.values || []);
    }

    async append(ops) {
      if (!ops.length) return;
      await this.json(this.sheetURL + '/values/' + encodeURIComponent(RANGE) +
        ':append?valueInputOption=RAW&insertDataOption=INSERT_ROWS', 'POST', { values: ops.map(rowFor) });
    }

    async image(ref, readLegacy) {
      if (!ref || ref.startsWith('drive:')) return ref;
      const key = await digest(ref);
      const stored = await this.list('appProperties has { key = ' + quote('financeImage') + ' and value = ' + quote(key) + ' } and ' + quote(this.folderId) + ' in parents');
      if (stored) return 'drive:' + stored.id;
      const blob = ref.startsWith('data:image/') ? await (await fetch(ref)).blob() : await readLegacy(ref);
      if (!blob.type.startsWith('image/') || !blob.size) throw new Error('A ledger image could not be copied.');
      const boundary = 'fintracker-' + crypto.randomUUID();
      const metadata = { name: 'Ledger image.' + (blob.type === 'image/png' ? 'png' : 'jpg'),
        mimeType: blob.type, parents: [this.folderId], appProperties: { financeImage: key } };
      const body = new Blob(['--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n',
        JSON.stringify(metadata), '\r\n--' + boundary + '\r\nContent-Type: ' + blob.type + '\r\n\r\n', blob,
        '\r\n--' + boundary + '--'], { type: 'multipart/related; boundary=' + boundary });
      const uploaded = await this.request('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', { method: 'POST', body });
      return 'drive:' + uploaded.id;
    }

    readImage(ref) {
      const id = ref.slice('drive:'.length);
      if (!/^[-\w]+$/.test(id)) throw new Error('Invalid image reference.');
      return this.request(DRIVE + '/files/' + encodeURIComponent(id) + '?alt=media', {}, 'blob');
    }

    async sync(ops) {
      const prepared = [];
      for (const op of ops) {
        const next = { ...op };
        if (next.row?.image?.startsWith('data:')) next.row = { ...next.row, image: await this.image(next.row.image) };
        prepared.push(next);
      }
      await this.append(prepared);
      const snapshot = await this.read();
      if (prepared.some(op => !snapshot.seen.has(op.event))) throw new Error('The Sheet has not confirmed all pending changes. Try syncing again.');
      return snapshot.data;
    }

    async migrate(data, pending, source, readLegacy, baseline) {
      const marker = await digest('migration:' + source);
      const previous = await this.read();
      const migrated = previous.latest.get('Migration\0' + marker);
      if (migrated) {
        if (migrated.value.baseline !== baseline) {
          throw new Error('The old Sheet changed after it was migrated. Your legacy connection was kept for review.');
        }
        const changes = [];
        for (const op of pending) {
          const row = op.row?.image ? { ...op.row, image: await this.image(op.row.image, readLegacy) } : op.row;
          changes.push({ ...op, row });
        }
        return this.sync(changes);
      }
      const entries = [];
      for (const [table, name] of Object.entries(TABS)) {
        for (const original of data[name]) {
          const row = { ...original };
          if (row.image) row.image = await this.image(row.image, readLegacy);
          entries.push({ sheet: table, row, op: 'upsert' });
        }
      }
      for (const [id, value] of Object.entries(data.settings)) entries.push({ sheet: 'Settings', row: { id, value }, op: 'upsert' });
      for (const op of pending) if (op.op === 'delete') entries.push({ sheet: op.sheet, id: op.id, op: 'delete' });
      const batch = [];
      for (const op of entries) {
        const id = String(op.row?.id ?? op.id);
        const prior = previous.latest.get(op.sheet + '\0' + id);
        if (prior) {
          if (prior.action !== op.op || (op.op === 'upsert' && !same(prior.value, op.row))) {
            throw new Error('The Drive ledger already has a different version of a legacy entry. The existing connection was kept.');
          }
        } else {
          batch.push({ ...op, event: await digest(marker + ':' + op.sheet + ':' + id + ':' + JSON.stringify(canonical(op.row || op.op))) });
        }
      }
      for (let offset = 0; offset < batch.length; offset += 100) await this.append(batch.slice(offset, offset + 100));
      const verified = await this.read();
      for (const op of entries) {
        const actual = verified.latest.get(op.sheet + '\0' + String(op.row?.id ?? op.id));
        if (!actual || actual.action !== op.op || (op.op === 'upsert' && !same(actual.value, op.row))) {
          throw new Error('The imported ledger could not be verified. The existing connection was kept.');
        }
      }
      await this.append([{ sheet: 'Migration', row: { id: marker, baseline }, op: 'upsert', event: marker }]);
      const completed = await this.read();
      if (!completed.seen.has(marker)) throw new Error('Google Sheets has not confirmed the migration.');
      return completed.data;
    }
  }

  function browserToken(clientId, hint) {
    return new Promise((resolve, reject) => {
      const client = google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: SCOPE,
        include_granted_scopes: false,
        login_hint: hint || '',
        callback(response) {
          if (response.error || !response.access_token ||
              !google.accounts.oauth2.hasGrantedAllScopes(response, SCOPE)) {
            reject(new Error('Google Drive access was not granted.'));
          } else resolve(response.access_token);
        },
        error_callback() { reject(new Error('Google sign-in did not finish. Try again.')); }
      });
      client.requestAccessToken({ prompt: hint ? '' : 'select_account' });
    });
  }

  root.FinanceDrive = { LedgerStore, browserToken, digest, decode, rowFor };
})(globalThis);
