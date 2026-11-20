/**
 * GO TO TARGET — Internal Client Workspace
 * Google Apps Script backend for index.html
 *
 * Database sheets: Clients, Workspaces, Users
 * Each team member has an individual username and salted password hash. Admin chooses the password from the Google Sheet menu.
 *
 * Deployment:
 * SIMPLE SETUP:
 * 1) If this script is attached to your Google Sheet, leave SPREADSHEET_URL_OR_ID blank.
 *    Otherwise paste the full Google Sheet URL (or Sheet ID) in GTT_CONFIG below.
 * 2) Run setupDatabaseSheets() once.
 * 3) In the Google Sheet menu, choose 'Set / reset Admin password' and choose your password.
 * 4) Deploy as Web app: Execute as Me, Who has access: Anyone.
 * 5) Paste the Web App URL ending in /exec into GTT_APPS_SCRIPT_URL in index.html.
 *
 * Security notes:
 * - All data operations except login/health checks require a short-lived session.
 * - Client/workspace writes and user administration require an ADMIN session.
 * - Each user has an individual salted password hash; plaintext passwords are never stored.
 * - The browser does not contain any password or API secret. All write actions require a signed-in session.
 */

const GTT_CONFIG = Object.freeze({
  // Leave blank when this script is attached to the target Google Sheet.
  // If this is a standalone Apps Script project, paste the full Sheet URL or ID here.
  SPREADSHEET_URL_OR_ID: '',
  SHEETS: {
    CLIENTS: 'Clients',
    WORKSPACES: 'Workspaces',
    USERS: 'Users'
  },
  SESSION_TTL_SECONDS: 21600, // 6 hours
  DEFAULT_ADMIN: {
    user_id: 'usr_admin_01',
    name: 'GO TO TARGET Admin',
    username: 'admin',
    email: 'admin@gototarget.co',
    role: 'ADMIN',
    avatar_url: '',
    active: true,
    password_salt: '',
    password_hash: '',
    session_version: 1
  }
});

const GTT_SCHEMAS = [
  {
    name: GTT_CONFIG.SHEETS.CLIENTS,
    headers: ['client_id', 'client_name', 'logo_url', 'description', 'archived', 'created_at'],
    widths: [140, 220, 260, 340, 100, 130]
  },
  {
    name: GTT_CONFIG.SHEETS.WORKSPACES,
    headers: ['workspace_id', 'client_id', 'workspace_name', 'description', 'icon', 'color', 'external_url', 'created_at', 'updated_at'],
    widths: [140, 140, 160, 320, 120, 110, 340, 130, 130]
  },
  {
    name: GTT_CONFIG.SHEETS.USERS,
    headers: ['user_id', 'name', 'username', 'email', 'role', 'avatar_url', 'active', 'password_salt', 'password_hash', 'session_version'],
    widths: [140, 200, 150, 240, 110, 240, 100, 160, 300, 120]
  }
];

// -----------------------------------------------------------------------------
// SHEET SETUP
// -----------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 GO TO TARGET')
    .addItem('1. Setup / repair database sheets', 'setupDatabaseSheets')
    .addItem('2. Set / reset Admin password', 'resetAdminPasswordFromSheet')
    .addItem('3. Add demo client and workspaces (optional)', 'seedInitialData')
    .addToUi();
}

/** Run once from the target Google Sheet or Apps Script editor. */
function setupDatabaseSheets() {
  const ss = getDatabaseSpreadsheet_();
  ensureDatabaseSheets_(ss);
  ensureDefaultUsers_();
  ss.toast('Database is ready: Clients, Workspaces, Users', 'GO TO TARGET', 6);
  Logger.log('GO TO TARGET database is ready: ' + ss.getUrl());
  return 'Database ready';
}

function seedInitialData() {
  const ss = getDatabaseSpreadsheet_();
  ensureDatabaseSheets_(ss);
  ensureDefaultUsers_();

  const clientsSheet = ss.getSheetByName(GTT_CONFIG.SHEETS.CLIENTS);
  const workspacesSheet = ss.getSheetByName(GTT_CONFIG.SHEETS.WORKSPACES);
  if (clientsSheet.getLastRow() > 1) {
    ss.toast('Client data already exists. No demo rows were added.', 'GO TO TARGET', 6);
    return;
  }

  const today = today_();
  const clientId = 'cli_demo_01';
  clientsSheet.appendRow([clientId, 'ABC Company', '', 'Demo client — safe to edit or delete', false, today]);
  [
    ['ws_demo_01', clientId, 'VIDEO', 'Video production and video working files', 'video', 'blue', 'https://drive.google.com', today, today],
    ['ws_demo_02', clientId, 'GRAPHIC', 'Graphic design and artwork files', 'image', 'cyan', '', today, today],
    ['ws_demo_03', clientId, 'CONTENT', 'Content, captions, copywriting and content planning', 'file-text', 'mint', 'https://docs.google.com/spreadsheets', today, today],
    ['ws_demo_04', clientId, 'BRIEF', 'Brief, references and client information', 'clipboard-list', 'orange', '', today, today]
  ].forEach(row => workspacesSheet.appendRow(row));

  ss.toast('Demo client and four workspaces added.', 'GO TO TARGET', 6);
}

/**
 * Set your own Admin password from the Google Sheet menu.
 * This avoids temporary passwords, execution-log hunting, and blocking alert dialogs.
 */
function resetAdminPasswordFromSheet() {
  const ss = getDatabaseSpreadsheet_();
  ensureDatabaseSheets_(ss);
  ensureDefaultUsers_();

  const admin = getUsers_().find(user =>
    normalizeUsername_(user.username) === 'admin' ||
    normalizeEmail_(user.email) === GTT_CONFIG.DEFAULT_ADMIN.email
  );
  if (!admin) throw new Error('Admin account not found. Run Setup / repair database sheets first.');

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Set Admin password',
    'Choose a password for username: admin\nUse at least 8 characters. The password will not be stored as plain text.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    ss.toast('Password was not changed.', 'GO TO TARGET', 5);
    return 'Cancelled';
  }

  const password = String(response.getResponseText() || '');
  if (password.length < 8) {
    ui.alert('Password must be at least 8 characters. Please run this menu item again.');
    return 'Password too short';
  }

  setUserPassword_(admin.user_id, password);
  updateObject_(GTT_CONFIG.SHEETS.USERS, 'user_id', admin.user_id, {
    active: true,
    role: 'ADMIN',
    username: 'admin'
  });
  ss.toast('Admin password saved. Sign in with username: admin', 'GO TO TARGET', 8);
  Logger.log('Admin password updated for user_id=' + admin.user_id);
  return 'Admin password saved';
}

function ensureDatabaseSheets_(ss) {
  GTT_SCHEMAS.forEach(schema => {
    let sheet = ss.getSheetByName(schema.name);
    if (!sheet) sheet = ss.insertSheet(schema.name);

    const existingLastColumn = Math.max(sheet.getLastColumn(), 1);
    const existingHeaders = sheet.getRange(1, 1, 1, existingLastColumn).getDisplayValues()[0]
      .map(value => String(value || '').trim());
    const isBlank = existingHeaders.every(value => !value);

    if (isBlank) {
      sheet.getRange(1, 1, 1, schema.headers.length).setValues([schema.headers]);
    } else {
      // Add missing columns at the right without deleting or reordering existing data.
      const missing = schema.headers.filter(header => existingHeaders.indexOf(header) === -1);
      if (missing.length) {
        sheet.getRange(1, existingLastColumn + 1, 1, missing.length).setValues([missing]);
      }
    }

    const headerWidth = Math.max(sheet.getLastColumn(), schema.headers.length);
    sheet.getRange(1, 1, 1, headerWidth)
      .setBackground('#061D42')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(10)
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 34);
    sheet.setFrozenRows(1);
    schema.widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));
  });
}

function ensureDefaultUsers_() {
  const sheet = getDatabaseSpreadsheet_().getSheetByName(GTT_CONFIG.SHEETS.USERS);
  const admin = GTT_CONFIG.DEFAULT_ADMIN;
  const existingUsers = sheetRows_(GTT_CONFIG.SHEETS.USERS);
  const adminExists = existingUsers.some(user => normalizeEmail_(user.email) === admin.email || normalizeUsername_(user.username) === 'admin');
  if (!adminExists) appendObject_(GTT_CONFIG.SHEETS.USERS, admin);

  // Backfill usernames for accounts created by the previous email-based version.
  const refreshedUsers = sheetRows_(GTT_CONFIG.SHEETS.USERS).sort((a, b) => {
    const aAdmin = normalizeEmail_(a.email) === admin.email || normalizeUsername_(a.username) === 'admin';
    const bAdmin = normalizeEmail_(b.email) === admin.email || normalizeUsername_(b.username) === 'admin';
    return Number(bAdmin) - Number(aAdmin);
  });
  const used = new Set();
  refreshedUsers.forEach((user, index) => {
    let username = normalizeUsername_(user.username);
    if (!username && normalizeEmail_(user.email) === admin.email) username = 'admin';
    if (!username) username = usernameFromEmail_(user.email) || ('user' + (index + 1));
    if (!/^[a-z0-9][a-z0-9._-]{2,29}$/.test(username)) username = 'user' + (index + 1);
    const base = username.slice(0, 25);
    let candidate = username;
    let suffix = 2;
    while (used.has(candidate)) {
      candidate = base + suffix;
      suffix++;
    }
    used.add(candidate);
    if (normalizeUsername_(user.username) !== candidate) {
      updateObject_(GTT_CONFIG.SHEETS.USERS, 'user_id', user.user_id, { username: candidate });
    }
  });
}

function getDatabaseSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const configured = String(GTT_CONFIG.SPREADSHEET_URL_OR_ID || '').trim();
  let configuredId = '';

  if (configured) {
    const match = configured.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    configuredId = match ? match[1] : configured;
    if (!/^[a-zA-Z0-9_-]{20,}$/.test(configuredId)) {
      throw new Error('SPREADSHEET_URL_OR_ID is not a valid Google Sheet URL or ID.');
    }
    props.setProperty('GTT_SPREADSHEET_ID', configuredId);
    return SpreadsheetApp.openById(configuredId);
  }

  const savedId = props.getProperty('GTT_SPREADSHEET_ID');
  if (savedId) return SpreadsheetApp.openById(savedId);

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    props.setProperty('GTT_SPREADSHEET_ID', active.getId());
    return active;
  }

  throw new Error('Google Sheet not configured. Paste its URL/ID into GTT_CONFIG.SPREADSHEET_URL_OR_ID, then run setupDatabaseSheets().');
}

function ensureReady_() {
  const ss = getDatabaseSpreadsheet_();
  const needsRepair = GTT_SCHEMAS.some(schema => {
    const sheet = ss.getSheetByName(schema.name);
    if (!sheet) return true;
    const width = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(value => String(value || '').trim());
    return schema.headers.some(header => headers.indexOf(header) === -1);
  });
  if (needsRepair) ensureDatabaseSheets_(ss);
  ensureDefaultUsers_();
  return ss;
}

// -----------------------------------------------------------------------------
// HTTP API
// -----------------------------------------------------------------------------
function doGet() {
  try {
    const ss = ensureReady_();
    return jsonResponse_(true, {
      status: 'online',
      connected: true,
      spreadsheet: ss.getName(),
      sheets: GTT_SCHEMAS.map(schema => schema.name),
      service: 'GO TO TARGET Internal Workspace API',
      timestamp: new Date().toISOString()
    });
  } catch (err) {
    return jsonResponse_(false, { connected: false }, err && err.message ? err.message : String(err));
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Request body is empty. Send a JSON POST request.');
    }

    const body = JSON.parse(e.postData.contents);
    const action = String(body.action || '').trim();
    const payload = body.payload || {};
    const sessionToken = String(body.session_token || '');

    if (!action) throw new Error('Missing action.');

    if (action === 'ping') {
      return jsonResponse_(true, { status: 'online', timestamp: new Date().toISOString() });
    }

    if (action === 'authenticateUser') {
      ensureReady_();
      return jsonResponse_(true, authenticateUser_(payload.username, payload.password || payload.passcode));
    }

    const actor = requireSession_(sessionToken);
    const ss = ensureReady_();
    let data;

    switch (action) {
      // READ: any authenticated member
      case 'validateSession':
        data = publicUser_(actor);
        break;
      case 'getCurrentUser':
        data = publicUser_(actor);
        break;
      case 'getClients':
        data = getClients_();
        break;
      case 'getClientById':
        data = getClientById_(payload.client_id);
        break;
      case 'getAllWorkspaces':
        data = getAllWorkspaces_();
        break;
      case 'getWorkspacesByClient':
        data = getWorkspacesByClient_(payload.client_id);
        break;
      case 'getUsers':
        requireAdmin_(actor);
        data = getUsers_().map(publicUser_);
        break;
      case 'importLocalDataIfEmpty':
        requireAdmin_(actor);
        data = withLock_(() => importLocalDataIfEmpty_(payload));
        break;

      // WRITE: Admin only
      case 'createClient':
        requireAdmin_(actor);
        data = withLock_(() => createClient_(payload));
        break;
      case 'updateClient':
        requireAdmin_(actor);
        data = withLock_(() => updateClient_(payload));
        break;
      case 'deleteClient':
        requireAdmin_(actor);
        data = withLock_(() => deleteClient_(payload.client_id));
        break;
      case 'createWorkspace':
        requireAdmin_(actor);
        data = withLock_(() => createWorkspace_(payload));
        break;
      case 'updateWorkspace':
        requireAdmin_(actor);
        data = withLock_(() => updateWorkspace_(payload));
        break;
      case 'deleteWorkspace':
        requireAdmin_(actor);
        data = withLock_(() => deleteWorkspace_(payload.workspace_id));
        break;
      case 'createUser':
        requireAdmin_(actor);
        data = withLock_(() => createUser_(payload));
        break;
      case 'updateUser':
        requireAdmin_(actor);
        data = withLock_(() => updateUser_(payload, actor));
        break;
      case 'deleteUser':
        requireAdmin_(actor);
        data = withLock_(() => deleteUser_(payload.user_id, actor));
        break;
      case 'resetUserPassword':
        requireAdmin_(actor);
        data = withLock_(() => resetUserPassword_(payload.user_id));
        break;
      default:
        throw new Error('Unknown action: ' + action);
    }

    return jsonResponse_(true, data);
  } catch (err) {
    return jsonResponse_(false, null, err && err.message ? err.message : String(err));
  }
}

function jsonResponse_(success, data, error) {
  return ContentService.createTextOutput(JSON.stringify({
    success: Boolean(success),
    data: data === undefined ? null : data,
    error: error || null
  })).setMimeType(ContentService.MimeType.JSON);
}

// -----------------------------------------------------------------------------
// SESSIONS & AUTHENTICATION
// -----------------------------------------------------------------------------
function authenticateUser_(username, password) {
  const normalizedUsername = normalizeUsername_(username);
  if (!normalizedUsername || !password) throw new Error('Enter your username and password.');
  if (!/^[a-z0-9][a-z0-9._-]{2,29}$/.test(normalizedUsername)) {
    throw new Error('Username format is invalid. Use 3–30 letters, numbers, dots, underscores or hyphens.');
  }

  assertLoginNotBlocked_(normalizedUsername);
  const user = getUsers_().find(item => normalizeUsername_(item.username) === normalizedUsername && item.active === true);
  if (!user) {
    recordLoginFailure_(normalizedUsername);
    throw new Error('Username not found or this account has been disabled.');
  }

  if (!user.password_hash || !user.password_salt) {
    recordLoginFailure_(normalizedUsername);
    throw new Error('Password is not set for this account. In the Google Sheet, open GO TO TARGET > Set / reset Admin password.');
  }

  if (!verifyPassword_(String(password), user)) {
    recordLoginFailure_(normalizedUsername);
    throw new Error('Invalid username or password.');
  }

  clearLoginFailures_(normalizedUsername);
  const sessionToken = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const sessionData = { user_id: user.user_id, session_version: Number(user.session_version || 1), issued_at: Date.now() };
  CacheService.getScriptCache().put('GTT_SESSION_' + sessionToken, JSON.stringify(sessionData), GTT_CONFIG.SESSION_TTL_SECONDS);
  return { user: publicUser_(user), session_token: sessionToken, expires_in: GTT_CONFIG.SESSION_TTL_SECONDS };
}

function loginRateKey_(email) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(email), Utilities.Charset.UTF_8);
  return 'GTT_LOGIN_FAIL_' + Utilities.base64Encode(digest).replace(/[^A-Za-z0-9]/g, '');
}

function assertLoginNotBlocked_(email) {
  const key = loginRateKey_(email);
  const count = Number(CacheService.getScriptCache().get(key) || 0);
  if (count >= 8) throw new Error('Too many sign-in attempts. Please try again in 15 minutes.');
}

function recordLoginFailure_(email) {
  const cache = CacheService.getScriptCache();
  const key = loginRateKey_(email);
  const count = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(count), 900);
}

function clearLoginFailures_(email) {
  CacheService.getScriptCache().remove(loginRateKey_(email));
}

// New hashes are stored as 'sha256$<hex>'. The old base64 format could start with '+', which the
// sheet may read as a formula (and the formula guard would now reject it). Hex never starts with
// = + - @. The prefix also leaves room for a stronger hash format later.
const GTT_HASH_PREFIX = 'sha256$';

function sha256Bytes_(password, salt) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + ':' + String(password),
    Utilities.Charset.UTF_8
  );
}

function hashPassword_(password, salt) {
  const hex = sha256Bytes_(password, salt).map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  return GTT_HASH_PREFIX + hex;
}

// Accepts both the new format and hashes saved by the first version (plain base64),
// so passwords that were already set keep working.
function verifyPassword_(password, user) {
  const stored = String(user.password_hash || '');
  const salt = String(user.password_salt || '');
  const expected = stored.indexOf(GTT_HASH_PREFIX) === 0
    ? hashPassword_(password, salt)
    : Utilities.base64Encode(sha256Bytes_(password, salt));
  return constantTimeEquals_(expected, stored);
}

function generateTemporaryPassword_() {
  return 'GTT-' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}

function setUserPassword_(userId, password) {
  const salt = Utilities.getUuid().replace(/-/g, '');
  const existing = getUsers_().find(user => String(user.user_id) === String(userId));
  const nextVersion = Number(existing && existing.session_version || 1) + 1;
  updateObject_(GTT_CONFIG.SHEETS.USERS, 'user_id', userId, {
    password_salt: salt,
    password_hash: hashPassword_(password, salt),
    session_version: nextVersion
  });
}

function constantTimeEquals_(a, b) {
  if (a.length !== b.length) return false;
  let mismatch = 0;
  for (let i = 0; i < a.length; i++) mismatch |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return mismatch === 0;
}

function requireSession_(sessionToken) {
  if (!sessionToken) throw new Error('Your session has expired. Please sign in again.');
  const raw = CacheService.getScriptCache().get('GTT_SESSION_' + sessionToken);
  if (!raw) throw new Error('Your session has expired. Please sign in again.');
  const session = JSON.parse(raw);
  const user = getUsers_().find(item => String(item.user_id) === String(session.user_id) && item.active === true);
  if (!user || Number(user.session_version || 1) !== Number(session.session_version || 1)) {
    CacheService.getScriptCache().remove('GTT_SESSION_' + sessionToken);
    throw new Error('Your session is no longer valid. Please sign in again.');
  }
  return user;
}

function requireAdmin_(user) {
  if (!user || String(user.role).toUpperCase() !== 'ADMIN') {
    throw new Error('Permission denied. Administrator access is required.');
  }
}

function publicUser_(user) {
  return {
    user_id: user.user_id,
    name: user.name,
    username: normalizeUsername_(user.username) || usernameFromEmail_(user.email),
    email: user.email || '',
    role: user.role,
    avatar_url: user.avatar_url || '',
    active: user.active === true
  };
}

// -----------------------------------------------------------------------------
// SHEET HELPERS
// -----------------------------------------------------------------------------
function getSheet_(name) {
  const sheet = getDatabaseSpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name + '. Run setupDatabaseSheets() first.');
  return sheet;
}

function sheetRows_(sheetName) {
  const sheet = getSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0].map(value => String(value || '').trim());
  return values.slice(1).filter(row => !row.every(value => value === '' || value === null)).map(row => {
    const item = {};
    headers.forEach((header, index) => {
      if (!header) return;
      let value = row[index];
      if (value instanceof Date) value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      item[header] = value;
    });
    return item;
  });
}

function headerMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};
  headers.forEach((header, index) => { if (header) map[String(header).trim()] = index + 1; });
  return map;
}

function findRowById_(sheetName, idColumn, id) {
  const sheet = getSheet_(sheetName);
  const map = headerMap_(sheet);
  const col = map[idColumn];
  if (!col) throw new Error('Missing column: ' + idColumn);
  if (sheet.getLastRow() <= 1) return -1;
  const values = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getDisplayValues();
  const index = values.findIndex(row => String(row[0]) === String(id));
  return index === -1 ? -1 : index + 2;
}

// Final safety net for every write to the sheet.
// Google Sheets treats text starting with = + - @ as a formula. A formula such as =IMPORTXML(...)
// could send sheet data (including password hashes) to an outside website, so it is never written.
function toCellValue_(value) {
  if (value === undefined || value === null) return '';
  if (isFormulaLike_(value)) throw new Error('Text cannot start with = + - or @.');
  return value;
}

// Text cells are set to '@' (plain text) BEFORE writing, so Sheets keeps the value exactly as sent
// (no auto-conversion to dates or numbers, leading zeros kept). Numbers use '0'.
function cellFormat_(value) {
  return typeof value === 'number' ? '0' : '@';
}

function appendObject_(sheetName, object) {
  const sheet = getSheet_(sheetName);
  const map = headerMap_(sheet);
  const width = sheet.getLastColumn();
  const row = new Array(width).fill('');
  Object.keys(object).forEach(key => {
    if (map[key]) row[map[key] - 1] = toCellValue_(object[key]);
  });
  const rowNumber = sheet.getLastRow() + 1;
  // appendRow() added rows automatically; getRange() does not, so grow the sheet when it is full.
  if (rowNumber > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 100);
  const target = sheet.getRange(rowNumber, 1, 1, width);
  target.setNumberFormats([row.map(cellFormat_)]);
  target.setValues([row]);
}

function updateObject_(sheetName, idColumn, id, patch) {
  const sheet = getSheet_(sheetName);
  const rowNumber = findRowById_(sheetName, idColumn, id);
  if (rowNumber < 0) throw new Error('Record not found: ' + id);
  const map = headerMap_(sheet);
  // Check every value first so a rejected field never leaves the row half-updated.
  const values = {};
  Object.keys(patch).forEach(key => {
    if (map[key] && key !== idColumn) values[key] = toCellValue_(patch[key]);
  });
  Object.keys(values).forEach(key => {
    sheet.getRange(rowNumber, map[key]).setNumberFormat(cellFormat_(values[key])).setValue(values[key]);
  });
  return rowNumber;
}

function withLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return callback(); } finally { lock.releaseLock(); }
}

function bool_(value) {
  return value === true || String(value).toLowerCase() === 'true' || String(value) === '1';
}
function normalizeEmail_(email) { return String(email || '').trim().toLowerCase(); }
function normalizeUsername_(username) { return String(username || '').trim().toLowerCase(); }
function usernameFromEmail_(email) { return normalizeEmail_(email).split('@')[0].replace(/[^a-z0-9._-]/g, ''); }
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function makeId_(prefix) { return prefix + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 16); }

// -----------------------------------------------------------------------------
// INPUT VALIDATION — never trust values sent from the browser.
// Anyone can call the Web App URL directly, so every field is checked here again.
// -----------------------------------------------------------------------------
const GTT_ALLOWED_ICONS = ['video', 'image', 'file-text', 'clipboard-list', 'folder-open', 'file-spreadsheet', 'link-2'];
const GTT_ALLOWED_COLORS = ['blue', 'cyan', 'mint', 'orange', 'navy'];

function isFormulaLike_(value) {
  return typeof value === 'string' && /^[=+\-@]/.test(value);
}

function cleanText_(value, label, maxLength, required) {
  const text = String(value == null ? '' : value).trim();
  if (required && !text) throw new Error(label + ' is required.');
  if (text.length > maxLength) throw new Error(label + ' must be ' + maxLength + ' characters or fewer.');
  if (isFormulaLike_(text)) throw new Error(label + ' cannot start with = + - or @.');
  return text;
}

// Links must be full web links. Blocks javascript: and other links that could run code in the page.
function cleanUrl_(value, label) {
  const url = String(value == null ? '' : value).trim();
  if (!url) return '';
  if (url.length > 2000 || !/^https?:\/\/[^\s<>"'`]+$/i.test(url)) {
    throw new Error(label + ' must be a full link starting with https://');
  }
  return url;
}

// IDs are placed into the page's click handlers, so only a simple, safe format is accepted.
function cleanId_(value, prefix) {
  const id = String(value == null ? '' : value).trim();
  if (!id) return makeId_(prefix);
  if (!/^[A-Za-z0-9_-]{1,64}$/.test(id)) throw new Error('Invalid ID format.');
  return id;
}

function cleanDate_(value) {
  const text = String(value == null ? '' : value).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(text) ? text : today_();
}

function cleanChoice_(value, allowed, fallback) {
  const choice = String(value == null ? '' : value).trim();
  return allowed.indexOf(choice) >= 0 ? choice : fallback;
}

function cleanEmail_(value) {
  const email = normalizeEmail_(value);
  if (!email) return '';
  if (email.length > 254 || !/^[^\s@<>"']+@[^\s@<>"']+\.[^\s@<>"']+$/.test(email) || isFormulaLike_(email)) {
    throw new Error('Email format is invalid.');
  }
  return email;
}

// -----------------------------------------------------------------------------
// CLIENTS
// -----------------------------------------------------------------------------
function getClients_() {
  return sheetRows_(GTT_CONFIG.SHEETS.CLIENTS).map(client => {
    client.archived = bool_(client.archived);
    return client;
  });
}

function getClientById_(clientId) {
  return getClients_().find(client => String(client.client_id) === String(clientId)) || null;
}

function cleanClientFields_(client) {
  return {
    client_name: cleanText_(client.client_name, 'Client name', 120, true),
    logo_url: cleanUrl_(client.logo_url, 'Client logo URL'),
    description: cleanText_(client.description, 'Description', 1000, false),
    archived: bool_(client.archived)
  };
}

function createClient_(payload) {
  const client = payload && payload.client;
  if (!client) throw new Error('Client name is required.');
  const id = cleanId_(client.client_id, 'cli');
  if (findRowById_(GTT_CONFIG.SHEETS.CLIENTS, 'client_id', id) >= 0) throw new Error('Client ID already exists.');

  const savedClient = Object.assign({ client_id: id }, cleanClientFields_(client), {
    created_at: cleanDate_(client.created_at)
  });
  // Check the default workspaces before writing anything, so a bad workspace
  // cannot leave a client saved without its workspaces.
  const workspaces = (Array.isArray(payload.defaultWorkspaces) ? payload.defaultWorkspaces : [])
    .map(workspace => buildWorkspace_(Object.assign({}, workspace, { client_id: id })));
  const workspaceIds = new Set();
  workspaces.forEach(ws => {
    if (workspaceIds.has(ws.workspace_id) || findRowById_(GTT_CONFIG.SHEETS.WORKSPACES, 'workspace_id', ws.workspace_id) >= 0) {
      throw new Error('Workspace ID already exists.');
    }
    workspaceIds.add(ws.workspace_id);
  });

  appendObject_(GTT_CONFIG.SHEETS.CLIENTS, savedClient);
  workspaces.forEach(ws => appendObject_(GTT_CONFIG.SHEETS.WORKSPACES, ws));
  return savedClient;
}

function updateClient_(client) {
  if (!client || !client.client_id) throw new Error('Client ID is required.');
  const patch = cleanClientFields_(client);
  updateObject_(GTT_CONFIG.SHEETS.CLIENTS, 'client_id', client.client_id, patch);
  return Object.assign({}, client, patch);
}

function deleteClient_(clientId) {
  if (!clientId) throw new Error('Client ID is required.');
  const clientsSheet = getSheet_(GTT_CONFIG.SHEETS.CLIENTS);
  const clientRow = findRowById_(GTT_CONFIG.SHEETS.CLIENTS, 'client_id', clientId);
  if (clientRow < 0) throw new Error('Client not found: ' + clientId);
  clientsSheet.deleteRow(clientRow);

  const wsSheet = getSheet_(GTT_CONFIG.SHEETS.WORKSPACES);
  const wsMap = headerMap_(wsSheet);
  const clientIdCol = wsMap.client_id;
  for (let row = wsSheet.getLastRow(); row >= 2; row--) {
    if (String(wsSheet.getRange(row, clientIdCol).getDisplayValue()) === String(clientId)) wsSheet.deleteRow(row);
  }
  return { deleted: clientId };
}

// -----------------------------------------------------------------------------
// WORKSPACES
// -----------------------------------------------------------------------------
function getAllWorkspaces_() { return sheetRows_(GTT_CONFIG.SHEETS.WORKSPACES); }
function getWorkspacesByClient_(clientId) {
  return getAllWorkspaces_().filter(workspace => String(workspace.client_id) === String(clientId));
}

function cleanWorkspaceFields_(workspace) {
  return {
    workspace_name: cleanText_(workspace.workspace_name, 'Workspace name', 60, true),
    description: cleanText_(workspace.description, 'Description', 500, false),
    icon: cleanChoice_(workspace.icon, GTT_ALLOWED_ICONS, 'folder-open'),
    color: cleanChoice_(workspace.color, GTT_ALLOWED_COLORS, 'blue'),
    external_url: cleanUrl_(workspace.external_url, 'Workspace link')
  };
}

// Validates a workspace and returns the row to save (does not write or check for duplicates).
function buildWorkspace_(workspace) {
  if (!workspace || !String(workspace.client_id || '').trim()) throw new Error('Workspace client_id is required.');
  return Object.assign({
    workspace_id: cleanId_(workspace.workspace_id, 'ws'),
    client_id: cleanId_(workspace.client_id, 'cli')
  }, cleanWorkspaceFields_(workspace), {
    created_at: cleanDate_(workspace.created_at),
    updated_at: cleanDate_(workspace.updated_at)
  });
}

function createWorkspace_(workspace) {
  const saved = buildWorkspace_(workspace);
  if (findRowById_(GTT_CONFIG.SHEETS.CLIENTS, 'client_id', saved.client_id) < 0) throw new Error('The selected client does not exist.');
  if (findRowById_(GTT_CONFIG.SHEETS.WORKSPACES, 'workspace_id', saved.workspace_id) >= 0) throw new Error('Workspace ID already exists.');
  appendObject_(GTT_CONFIG.SHEETS.WORKSPACES, saved);
  return saved;
}

function updateWorkspace_(workspace) {
  if (!workspace || !workspace.workspace_id) throw new Error('Workspace ID is required.');
  const patch = Object.assign(cleanWorkspaceFields_(workspace), { updated_at: today_() });
  updateObject_(GTT_CONFIG.SHEETS.WORKSPACES, 'workspace_id', workspace.workspace_id, patch);
  return Object.assign({}, workspace, patch);
}

function deleteWorkspace_(workspaceId) {
  if (!workspaceId) throw new Error('Workspace ID is required.');
  const sheet = getSheet_(GTT_CONFIG.SHEETS.WORKSPACES);
  const row = findRowById_(GTT_CONFIG.SHEETS.WORKSPACES, 'workspace_id', workspaceId);
  if (row < 0) throw new Error('Workspace not found: ' + workspaceId);
  sheet.deleteRow(row);
  return { deleted: workspaceId };
}

// -----------------------------------------------------------------------------
// ONE-TIME LOCAL DATA IMPORT
// Only imports when both cloud data sheets are empty. Existing cloud data is never overwritten.
// -----------------------------------------------------------------------------
function importLocalDataIfEmpty_(payload) {
  const clientsSheet = getSheet_(GTT_CONFIG.SHEETS.CLIENTS);
  const workspacesSheet = getSheet_(GTT_CONFIG.SHEETS.WORKSPACES);
  if (clientsSheet.getLastRow() > 1 || workspacesSheet.getLastRow() > 1) {
    return { skipped: true, reason: 'Cloud database already contains records.' };
  }

  const importedClientIds = new Set();
  let clientsImported = 0;
  let workspacesImported = 0;
  let usersImported = 0;
  const userCredentials = [];

  (Array.isArray(payload.clients) ? payload.clients : []).forEach(client => {
    if (!client || !String(client.client_id || '').trim() || !String(client.client_name || '').trim()) return;
    const saved = Object.assign({ client_id: cleanId_(client.client_id, 'cli') }, cleanClientFields_(client), {
      created_at: cleanDate_(client.created_at)
    });
    appendObject_(GTT_CONFIG.SHEETS.CLIENTS, saved);
    importedClientIds.add(saved.client_id);
    clientsImported++;
  });

  (Array.isArray(payload.workspaces) ? payload.workspaces : []).forEach(workspace => {
    if (!workspace || !importedClientIds.has(String(workspace.client_id))) return;
    const saved = Object.assign({}, workspace, { workspace_id: String(workspace.workspace_id || makeId_('ws')) });
    createWorkspace_(saved);
    workspacesImported++;
  });

  const knownUsernames = new Set(getUsers_().map(user => normalizeUsername_(user.username)));
  (Array.isArray(payload.users) ? payload.users : []).forEach(user => {
    if (!user) return;
    const legacyEmail = normalizeEmail_(user.email);
    const username = normalizeUsername_(user.username) || usernameFromEmail_(legacyEmail);
    if (!username || knownUsernames.has(username)) return;
    const created = createUser_({
      user_id: user.user_id,
      name: user.name,
      username: username,
      email: legacyEmail,
      role: user.role,
      avatar_url: user.avatar_url || '',
      active: user.active !== false
    });
    knownUsernames.add(username);
    usersImported++;
    userCredentials.push({ username: username, password: created.temporary_password });
  });

  return {
    imported: true,
    clients: clientsImported,
    workspaces: workspacesImported,
    users: usersImported,
    user_credentials: userCredentials
  };
}

// -----------------------------------------------------------------------------
// USERS — managed centrally in the Users sheet
// -----------------------------------------------------------------------------
function getUsers_() {
  return sheetRows_(GTT_CONFIG.SHEETS.USERS).map(user => {
    user.active = bool_(user.active);
    user.role = String(user.role || 'STAFF').toUpperCase() === 'ADMIN' ? 'ADMIN' : 'STAFF';
    user.username = normalizeUsername_(user.username) || usernameFromEmail_(user.email);
    user.email = normalizeEmail_(user.email);
    return user;
  });
}

function createUser_(user) {
  const name = cleanText_(user.name, 'Team member name', 100, true);
  const username = normalizeUsername_(user.username) || usernameFromEmail_(user.email);
  const email = cleanEmail_(user.email);
  const role = String(user.role || 'STAFF').toUpperCase();
  if (!/^[a-z0-9][a-z0-9._-]{2,29}$/.test(username)) {
    throw new Error('Username must be 3–30 characters and use only letters, numbers, dots, underscores or hyphens.');
  }
  if (role !== 'ADMIN' && role !== 'STAFF') throw new Error('Role must be ADMIN or STAFF.');
  if (getUsers_().some(existing => normalizeUsername_(existing.username) === username)) {
    throw new Error('This username is already registered.');
  }
  if (email && !/^\S+@\S+\.\S+$/.test(email)) throw new Error('Email format is invalid.');
  if (email && getUsers_().some(existing => normalizeEmail_(existing.email) === email)) throw new Error('This email is already registered.');

  let userId = cleanId_(user.user_id, 'usr');
  if (findRowById_(GTT_CONFIG.SHEETS.USERS, 'user_id', userId) >= 0) userId = makeId_('usr');
  const temporaryPassword = generateTemporaryPassword_();
  const passwordSalt = Utilities.getUuid().replace(/-/g, '');
  const saved = {
    user_id: userId,
    name,
    username,
    email,
    role,
    avatar_url: cleanUrl_(user.avatar_url, 'Avatar URL'),
    active: user.active === undefined ? true : bool_(user.active),
    session_version: 1,
    password_salt: passwordSalt,
    password_hash: hashPassword_(temporaryPassword, passwordSalt)
  };
  appendObject_(GTT_CONFIG.SHEETS.USERS, saved);
  return { user: publicUser_(saved), temporary_password: temporaryPassword };
}

function updateUser_(user, actor) {
  if (!user || !user.user_id) throw new Error('User ID is required.');
  const existing = getUsers_().find(item => String(item.user_id) === String(user.user_id));
  if (!existing) throw new Error('User not found.');

  const patch = {};
  if (user.name !== undefined) patch.name = cleanText_(user.name, 'Name', 100, false);
  if (user.username !== undefined) patch.username = normalizeUsername_(user.username);
  if (user.email !== undefined) patch.email = cleanEmail_(user.email);
  if (user.role !== undefined) {
    const role = String(user.role).toUpperCase();
    if (role !== 'ADMIN' && role !== 'STAFF') throw new Error('Role must be ADMIN or STAFF.');
    patch.role = role;
  }
  if (user.avatar_url !== undefined) patch.avatar_url = cleanUrl_(user.avatar_url, 'Avatar URL');
  if (user.active !== undefined) patch.active = bool_(user.active);
  if (patch.name !== undefined && !patch.name) throw new Error('Name cannot be blank.');
  if (patch.username !== undefined && !/^[a-z0-9][a-z0-9._-]{2,29}$/.test(patch.username)) {
    throw new Error('Username must be 3–30 characters and use only letters, numbers, dots, underscores or hyphens.');
  }
  if (patch.username && getUsers_().some(item => item.user_id !== existing.user_id && normalizeUsername_(item.username) === patch.username)) {
    throw new Error('This username is already registered.');
  }
  if (patch.email && !/^\S+@\S+\.\S+$/.test(patch.email)) throw new Error('Email format is invalid.');
  if (patch.email && getUsers_().some(item => item.user_id !== existing.user_id && normalizeEmail_(item.email) === patch.email)) {
    throw new Error('This email is already registered.');
  }

  const resultingRole = patch.role || existing.role;
  const resultingActive = patch.active === undefined ? existing.active : patch.active;
  if (existing.user_id === actor.user_id && (resultingRole !== 'ADMIN' || !resultingActive)) {
    throw new Error('You cannot disable or demote your own administrator account.');
  }
  ensureAtLeastOneAdmin_(existing.user_id, resultingRole, resultingActive);
  updateObject_(GTT_CONFIG.SHEETS.USERS, 'user_id', existing.user_id, patch);
  return publicUser_(Object.assign({}, existing, patch));
}

function deleteUser_(userId, actor) {
  if (!userId) throw new Error('User ID is required.');
  if (String(userId) === String(actor.user_id)) throw new Error('You cannot remove your own account.');
  const users = getUsers_();
  const existing = users.find(user => String(user.user_id) === String(userId));
  if (!existing) throw new Error('User not found.');
  if (existing.active && existing.role === 'ADMIN') {
    const otherAdmins = users.filter(user => user.active && user.role === 'ADMIN' && user.user_id !== existing.user_id);
    if (!otherAdmins.length) throw new Error('At least one active administrator must remain.');
  }
  const sheet = getSheet_(GTT_CONFIG.SHEETS.USERS);
  const row = findRowById_(GTT_CONFIG.SHEETS.USERS, 'user_id', userId);
  sheet.deleteRow(row);
  return { deleted: userId };
}

function resetUserPassword_(userId) {
  if (!userId) throw new Error('User ID is required.');
  const user = getUsers_().find(item => String(item.user_id) === String(userId));
  if (!user) throw new Error('User not found.');
  const temporaryPassword = generateTemporaryPassword_();
  setUserPassword_(user.user_id, temporaryPassword);
  return { user: publicUser_(user), temporary_password: temporaryPassword };
}

function ensureAtLeastOneAdmin_(changingUserId, resultingRole, resultingActive) {
  const users = getUsers_();
  const activeAdminsAfter = users.filter(user => user.active && user.role === 'ADMIN' && user.user_id !== changingUserId).length
    + ((resultingActive && resultingRole === 'ADMIN') ? 1 : 0);
  if (activeAdminsAfter < 1) throw new Error('At least one active administrator must remain.');
}
