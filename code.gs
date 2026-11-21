/**
 * GO TO TARGET — Internal Client Workspace
 * Google Apps Script backend for index.html
 *
 * Database sheets: Clients, Workspaces, Users, Deliverables
 * Each team member has an individual username and salted password hash. Admin sets passwords from the web page.
 *
 * Deployment:
 * SIMPLE SETUP:
 * 1) If this script is attached to your Google Sheet, leave SPREADSHEET_URL_OR_ID blank.
 *    Otherwise paste the full Google Sheet URL (or Sheet ID) in GTT_CONFIG below.
 * 2) Run testSetup() once (grants permissions for Sheets + Drive), then setupDatabaseSheets().
 * 3) In the Google Sheet menu, choose 'Set / reset Admin password' and choose your password.
 * 4) Deploy as Web app: Execute as Me, Who has access: Anyone.
 *    To update later: Deploy > Manage deployments > edit > Version: New version (keeps the same URL).
 * 5) Paste the Web App URL ending in /exec into GTT_APPS_SCRIPT_URL in index.html.
 *
 * Security notes:
 * - All data operations except login/health checks require a short-lived session.
 * - Client/workspace/deliverable management and user administration require an ADMIN session.
 * - STAFF only see the clients whose team they are in (Clients.team_ids), with those clients'
 *   workspaces and their own pieces of work. The filter runs here, not only in the page.
 * - STAFF can only read, and submit a result link for deliverables assigned to them.
 *   Sending the link finishes the piece (no separate review step — decided 21 Nov 2026).
 * - Each user has an individual salted, iterated password hash; plaintext passwords are never stored.
 */

const GTT_CONFIG = Object.freeze({
  // Leave blank when this script is attached to the target Google Sheet.
  // If this is a standalone Apps Script project, paste the full Sheet URL or ID here.
  SPREADSHEET_URL_OR_ID: '',
  // Company Google Drive folder for client logos.
  // The Google account that deploys the Web App ("Execute as: Me") must be able to edit this folder.
  LOGO_FOLDER_ID: '1olYQzFMs7WsuGqSiA3DNytDCgkKmhxC_',
  SHEETS: {
    CLIENTS: 'Clients',
    WORKSPACES: 'Workspaces',
    USERS: 'Users',
    DELIVERABLES: 'Deliverables',
    PACKAGES: 'Packages'
  },
  SESSION_TTL_SECONDS: 21600, // 6 hours
  PASSWORD_MIN_LENGTH: 6,
  // Each extra round makes guessing a leaked hash slower. testSetup() logs how long one hash takes.
  HASH_ITERATIONS: 1000,
  // Sheet rows are cached so most requests skip reading the sheet. Writes made through this script
  // clear the cache immediately; edits typed directly into the sheet are picked up by onEdit(),
  // or after this many seconds at most.
  CACHE_SECONDS: 300,
  // Bump this when GTT_SCHEMAS changes so every deployment re-checks the sheet columns once.
  SCHEMA_VERSION: '2026-11-21e',
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
    // team_ids = user_ids of the STAFF who look after this client, comma-separated.
    headers: ['client_id', 'client_name', 'logo_url', 'description', 'archived', 'created_at', 'team_ids'],
    widths: [140, 220, 260, 340, 100, 130, 220]
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
  },
  {
    // One row = one set of work for one client in one month, e.g. "Banner x5, Nov 2026".
    // The team finishes the whole set and sends ONE link (usually a Drive folder); counts add up
    // `quantity`. Extra work later is added as a new row. Older rows without quantity count as 1.
    name: GTT_CONFIG.SHEETS.DELIVERABLES,
    headers: ['deliverable_id', 'client_id', 'month', 'type', 'title', 'assignee_id', 'status', 'result_url',
      'review_note', 'submitted_by', 'submitted_at', 'reviewed_at', 'created_at', 'updated_at', 'quantity', 'package_id'],
    widths: [150, 140, 90, 110, 220, 140, 100, 320, 280, 140, 140, 140, 140, 140, 90, 150]
  },
  {
    // A client's contract: from start_month for <months> months, the same work every month.
    // items_json = [{type, quantity, assignee_id, title}]. last_generated_month remembers which month
    // was already created, so deleting a created piece never brings it back.
    name: GTT_CONFIG.SHEETS.PACKAGES,
    headers: ['package_id', 'client_id', 'start_month', 'months', 'items_json', 'active', 'last_generated_month', 'created_at', 'updated_at'],
    widths: [150, 140, 100, 80, 420, 80, 150, 140, 140]
  }
];

// Work types and statuses are fixed lists so the sheet can never hold a value the page cannot show.
const GTT_DELIVERABLE_TYPES = ['banner', 'video', 'influencer', 'content', 'other'];
// A piece is either not sent yet, or done (its result link was sent). The earlier review statuses
// (REVIEW / REVISE / APPROVED) may still sit in old rows; getDeliverables_() reads them as these two.
const GTT_STATUS = Object.freeze({ TODO: 'TODO', DONE: 'DONE' });
const GTT_MAX_PIECES_PER_ADD = 30;
const GTT_MAX_PACKAGE_MONTHS = 36;
const GTT_MAX_PACKAGE_ITEMS = 20;

// -----------------------------------------------------------------------------
// SHEET SETUP
// -----------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 GO TO TARGET')
    .addItem('1. Setup / repair database sheets', 'setupDatabaseSheets')
    .addItem('2. Set / reset Admin password', 'resetAdminPasswordFromSheet')
    .addItem('3. Add demo client and workspaces (optional)', 'seedInitialData')
    .addItem('4. Refresh web data (after editing the sheet by hand)', 'clearAllCaches')
    .addToUi();
}

// Simple trigger: when someone types into the sheet directly, drop the cached copy of that sheet
// so the website shows the change on the next request instead of after CACHE_SECONDS.
function onEdit(e) {
  try {
    if (e && e.range) invalidateRows_(e.range.getSheet().getName());
  } catch (err) {
    // Never block the person editing the sheet.
  }
}

// -----------------------------------------------------------------------------
// KEEP WARM — a time trigger refreshes the cached sheet rows every 5 minutes, so people opening
// the website rarely wait for the sheet to be read. Install once: Run installWarmUpTrigger().
// -----------------------------------------------------------------------------
function warmUp() {
  // Never fail loudly: a failing trigger would email the owner every time.
  try {
    ensureReady_();
    const lock = LockService.getScriptLock();
    // Under the lock so a save cannot happen between reading the sheet and storing the copy
    // (that could put old data back into the cache). If someone is saving right now, skip this round.
    if (!lock.tryLock(5000)) return 'busy';
    try {
      try { generateDueWork_(); markMonthlyWorkDone_(); } catch (err) { /* never block the warm-up */ }
      const cache = CacheService.getScriptCache();
      GTT_SCHEMAS.forEach(schema => {
        try {
          // Slightly longer than the 5-minute trigger so the copy never runs out between rounds.
          cache.put(rowsCacheKey_(schema.name), JSON.stringify(readSheetRows_(schema.name)), GTT_CONFIG.CACHE_SECONDS + 120);
        } catch (err) { /* too big for the cache: that sheet is just read on demand */ }
      });
    } finally {
      lock.releaseLock();
    }
    jsSha256Ok_();
    return 'ok';
  } catch (err) {
    return 'error: ' + (err && err.message ? err.message : err);
  }
}

/** Run once from the Apps Script editor (asks for permission to manage triggers). Safe to run again. */
function installWarmUpTrigger() {
  removeWarmUpTrigger();
  ScriptApp.newTrigger('warmUp').timeBased().everyMinutes(5).create();
  const result = warmUp();
  Logger.log('Warm-up trigger installed (every 5 minutes). First run: ' + result);
  return 'Installed';
}

function removeWarmUpTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'warmUp')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
  return 'Removed';
}

function clearAllCaches() {
  GTT_SCHEMAS.forEach(schema => invalidateRows_(schema.name));
  CacheService.getScriptCache().remove(readyCacheKey_());
  try {
    getDatabaseSpreadsheet_().toast('Website data will reload from the sheet on the next request.', 'GO TO TARGET', 5);
  } catch (err) {}
  return 'Cache cleared';
}

/** Run once from the target Google Sheet or Apps Script editor. */
function setupDatabaseSheets() {
  const ss = getDatabaseSpreadsheet_();
  ensureDatabaseSheets_(ss);
  ensureDefaultUsers_();
  CacheService.getScriptCache().put(readyCacheKey_(), '1', 21600);
  ss.toast('Database is ready: Clients, Workspaces, Users, Deliverables', 'GO TO TARGET', 6);
  Logger.log('GO TO TARGET database is ready: ' + ss.getUrl());
  return 'Database ready';
}

/**
 * Press Run on this function first: it triggers the permission prompt (Sheets + Drive) and
 * checks that the script can reach the Sheet and the logo folder. Read-only — it never writes,
 * so it is safe to run on the live Sheet. Results appear in the Execution log.
 */
function testSetup() {
  const lines = [];
  let ss;
  try {
    ss = getDatabaseSpreadsheet_();
    lines.push('OK  Spreadsheet: ' + ss.getName() + ' (' + ss.getUrl() + ')');
  } catch (err) {
    Logger.log('FAIL Cannot open spreadsheet: ' + (err && err.message ? err.message : err));
    return 'FAIL';
  }

  let allSheetsOk = true;
  GTT_SCHEMAS.forEach(schema => {
    const sheet = ss.getSheetByName(schema.name);
    if (!sheet) {
      allSheetsOk = false;
      lines.push('MISSING sheet "' + schema.name + '" -> run setupDatabaseSheets()');
      return;
    }
    const width = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(value => String(value || '').trim());
    const missing = schema.headers.filter(header => headers.indexOf(header) === -1);
    if (missing.length) {
      allSheetsOk = false;
      lines.push('MISSING columns in "' + schema.name + '": ' + missing.join(', ') + ' -> run setupDatabaseSheets()');
    } else {
      lines.push('OK  Sheet "' + schema.name + '": ' + Math.max(sheet.getLastRow() - 1, 0) + ' rows');
    }
  });

  // Only report whether a password exists; never log the hash itself.
  if (allSheetsOk) {
    const admins = readSheetRows_(GTT_CONFIG.SHEETS.USERS).filter(user => String(user.role) === 'ADMIN' && bool_(user.active));
    const withPassword = admins.filter(user => String(user.password_hash || '').trim());
    if (!admins.length) lines.push('MISSING active ADMIN user -> run setupDatabaseSheets()');
    else if (!withPassword.length) lines.push('TODO Admin password not set -> Sheet menu 🚀 GO TO TARGET > 2. Set / reset Admin password');
    else lines.push('OK  Admin accounts with password: ' + withPassword.map(user => user.username).join(', '));
  }

  try {
    const folder = DriveApp.getFolderById(GTT_CONFIG.LOGO_FOLDER_ID);
    const me = Session.getEffectiveUser().getEmail();
    const owner = folder.getOwner() ? folder.getOwner().getEmail() : '';
    const canEdit = owner === me || folder.getEditors().some(editor => editor.getEmail() === me);
    lines.push((canEdit ? 'OK  ' : 'FAIL ') + 'Logo folder "' + folder.getName() + '"' +
      (canEdit ? ' (this account can upload)' : ' -> share the folder with ' + me + ' as Editor'));
  } catch (err) {
    lines.push('FAIL Logo folder: ' + (err && err.message ? err.message : err) + ' -> check LOGO_FOLDER_ID and sharing');
  }

  const started = Date.now();
  hashPassword_('timing-test', 'salt');
  lines.push('INFO One password hash takes ' + (Date.now() - started) + ' ms (' + GTT_CONFIG.HASH_ITERATIONS + ' rounds, ' +
    (jsSha256Ok_() ? 'fast JavaScript SHA-256' : 'SLOW: JavaScript SHA-256 self-check failed, using Utilities') + ')');

  Logger.log(lines.join('\n'));
  return allSheetsOk ? 'OK' : 'NEEDS SETUP';
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
  invalidateRows_(GTT_CONFIG.SHEETS.CLIENTS);
  invalidateRows_(GTT_CONFIG.SHEETS.WORKSPACES);

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
    'Choose a password for username: admin\nUse at least ' + GTT_CONFIG.PASSWORD_MIN_LENGTH + ' characters. The password will not be stored as plain text.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    ss.toast('Password was not changed.', 'GO TO TARGET', 5);
    return 'Cancelled';
  }

  const password = String(response.getResponseText() || '');
  if (password.length < GTT_CONFIG.PASSWORD_MIN_LENGTH) {
    ui.alert('Password must be at least ' + GTT_CONFIG.PASSWORD_MIN_LENGTH + ' characters. Please run this menu item again.');
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
      // First time the team column appears: fill it from the work already handed out,
      // so nobody loses sight of a client they already have work for.
      if (schema.name === GTT_CONFIG.SHEETS.CLIENTS && missing.indexOf('team_ids') >= 0) {
        invalidateRows_(schema.name);
        backfillClientTeams_();
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
    invalidateRows_(schema.name);
  });
}

function ensureDefaultUsers_() {
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

// Opened once per request and reused: openById() is one of the slowest calls in Apps Script.
let gttSpreadsheet_ = null;

function getDatabaseSpreadsheet_() {
  if (gttSpreadsheet_) return gttSpreadsheet_;
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
    gttSpreadsheet_ = SpreadsheetApp.openById(configuredId);
    return gttSpreadsheet_;
  }

  const savedId = props.getProperty('GTT_SPREADSHEET_ID');
  if (savedId) {
    gttSpreadsheet_ = SpreadsheetApp.openById(savedId);
    return gttSpreadsheet_;
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    props.setProperty('GTT_SPREADSHEET_ID', active.getId());
    gttSpreadsheet_ = active;
    return active;
  }

  throw new Error('Google Sheet not configured. Paste its URL/ID into GTT_CONFIG.SPREADSHEET_URL_OR_ID, then run setupDatabaseSheets().');
}

function readyCacheKey_() { return 'GTT_READY_' + GTT_CONFIG.SCHEMA_VERSION; }

// Checking every sheet's columns on every request was slow, so it now runs once (under the lock,
// because it may write) and is remembered for 6 hours.
function ensureReady_() {
  const cache = CacheService.getScriptCache();
  if (cache.get(readyCacheKey_())) return getDatabaseSpreadsheet_();
  return withLock_(() => {
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
    cache.put(readyCacheKey_(), '1', 21600);
    return ss;
  });
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
      ensureMonthlyWork_();
      const auth = authenticateUser_(payload.username, payload.password || payload.passcode);
      // Send the first page of data with the login reply so the page needs one round trip, not two.
      auth.bootstrap = buildBootstrap_(getUsers_().find(user => user.user_id === auth.user.user_id), payload.month);
      return jsonResponse_(true, auth);
    }

    ensureReady_();
    ensureMonthlyWork_();
    const actor = requireSession_(sessionToken);
    let data;

    switch (action) {
      // READ: any authenticated member
      case 'bootstrap':
        data = buildBootstrap_(actor, payload.month);
        break;
      case 'validateSession':
      case 'getCurrentUser':
        data = publicUser_(actor);
        break;
      case 'getClients':
        data = visibleClients_(actor);
        break;
      case 'getAllWorkspaces':
        data = visibleWorkspaces_(actor);
        break;
      case 'getDeliverables':
        data = deliverablesForUser_(actor, cleanMonth_(payload.month));
        break;

      // STAFF + ADMIN: send a finished link for a piece of work (own pieces only for STAFF)
      case 'submitDeliverable':
        data = withLock_(() => submitDeliverable_(payload, actor));
        break;

      // ADMIN only
      case 'getUsers':
        requireAdmin_(actor);
        data = getUsers_().map(publicUser_);
        break;
      case 'uploadClientLogo':
        requireAdmin_(actor);
        // Drive upload is slow, so it runs outside the lock.
        data = uploadClientLogo_(payload);
        break;
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
      case 'createDeliverables':
        requireAdmin_(actor);
        data = withLock_(() => createDeliverables_(payload));
        break;
      case 'updateDeliverable':
        requireAdmin_(actor);
        data = withLock_(() => updateDeliverable_(payload));
        break;
      case 'deleteDeliverable':
        requireAdmin_(actor);
        data = withLock_(() => deleteDeliverable_(payload.deliverable_id));
        break;
      case 'copyDeliverables':
        requireAdmin_(actor);
        data = withLock_(() => copyDeliverables_(payload));
        break;
      case 'getPackages':
        requireAdmin_(actor);
        data = getPackages_();
        break;
      case 'savePackage':
        requireAdmin_(actor);
        data = withLock_(() => savePackage_(payload));
        break;
      case 'deletePackage':
        requireAdmin_(actor);
        data = withLock_(() => deletePackage_(payload.package_id));
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
      case 'setUserPassword':
        requireAdmin_(actor);
        data = withLock_(() => adminSetUserPassword_(payload));
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

// Everything the page needs after sign-in, in one response.
function buildBootstrap_(actor, requestedMonth) {
  const isAdmin = actor.role === 'ADMIN';
  const users = getUsers_();
  // A bad month from the browser must not break sign-in; fall back to this month.
  const month = /^\d{4}-(0[1-9]|1[0-2])$/.test(String(requestedMonth || '')) ? String(requestedMonth) : thisMonth_();
  return {
    user: publicUser_(actor),
    month: month,
    clients: visibleClients_(actor),
    workspaces: visibleWorkspaces_(actor),
    // Names only, so pages can show who is responsible for each piece of work.
    team: users.filter(user => user.active).map(user => ({ user_id: user.user_id, name: user.name, role: user.role })),
    users: isAdmin ? users.map(publicUser_) : [],
    // Contracts are for admins only (staff simply see the work that gets created each month).
    packages: isAdmin ? getPackages_() : [],
    deliverables: deliverablesForUser_(actor, month)
  };
}

// -----------------------------------------------------------------------------
// SESSIONS & AUTHENTICATION
// -----------------------------------------------------------------------------
function authenticateUser_(username, password) {
  const normalizedUsername = normalizeUsername_(username);
  if (!normalizedUsername || !password) throw new Error('Enter your username and password.');
  // Same message for every failure, so the reply never reveals whether a username exists.
  const failMessage = 'Invalid username or password.';
  if (!/^[a-z0-9][a-z0-9._-]{2,29}$/.test(normalizedUsername)) throw new Error(failMessage);

  assertLoginNotBlocked_(normalizedUsername);
  const user = getUsers_().find(item => normalizeUsername_(item.username) === normalizedUsername && item.active === true);
  if (!user || !user.password_hash || !user.password_salt || !verifyPassword_(String(password), user)) {
    recordLoginFailure_(normalizedUsername);
    throw new Error(failMessage);
  }

  clearLoginFailures_(normalizedUsername);
  // Upgrade hashes saved in an older or weaker format now that we know the password is correct.
  if (needsRehash_(user)) {
    try {
      withLock_(() => {
        const salt = Utilities.getUuid().replace(/-/g, '');
        updateObject_(GTT_CONFIG.SHEETS.USERS, 'user_id', user.user_id, {
          password_salt: salt,
          password_hash: hashPassword_(String(password), salt)
        });
      });
    } catch (err) {
      // Sign-in still succeeds; the upgrade is retried next time.
    }
  }

  const sessionToken = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const sessionData = { user_id: user.user_id, session_version: Number(user.session_version || 1), issued_at: Date.now() };
  CacheService.getScriptCache().put('GTT_SESSION_' + sessionToken, JSON.stringify(sessionData), GTT_CONFIG.SESSION_TTL_SECONDS);
  return { user: publicUser_(user), session_token: sessionToken, expires_in: GTT_CONFIG.SESSION_TTL_SECONDS };
}

function loginRateKey_(username) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(username), Utilities.Charset.UTF_8);
  return 'GTT_LOGIN_FAIL_' + Utilities.base64Encode(digest).replace(/[^A-Za-z0-9]/g, '');
}

function assertLoginNotBlocked_(username) {
  const key = loginRateKey_(username);
  const count = Number(CacheService.getScriptCache().get(key) || 0);
  if (count >= 8) throw new Error('Too many sign-in attempts. Please try again in 15 minutes.');
}

function recordLoginFailure_(username) {
  const cache = CacheService.getScriptCache();
  const key = loginRateKey_(username);
  const count = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(count), 900);
}

function clearLoginFailures_(username) {
  CacheService.getScriptCache().remove(loginRateKey_(username));
}

// Hash formats (all hex, so none can start with = + - @ in the sheet):
//   'sha256i$<rounds>$<hex>' — current: salted SHA-256 repeated <rounds> times
//   'sha256$<hex>'           — previous single-round format (still accepted, upgraded at next sign-in)
//   plain base64             — first version (still accepted, upgraded at next sign-in)
const GTT_HASH_PREFIX = 'sha256$';
const GTT_HASH_ITER_PREFIX = 'sha256i$';

function sha256Bytes_(password, salt) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + ':' + String(password),
    Utilities.Charset.UTF_8
  );
}

function toHex_(bytes) {
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

// Round 1 uses Utilities (it handles the UTF-8 text). Rounds 2..N hash the previous 32 bytes.
// Each Utilities.computeDigest call costs about 1.2 ms in Apps Script, so 1000 rounds took ~1.2 s
// on every sign-in (measured 21 Nov 2026). The same SHA-256 written in plain JavaScript gives
// byte-for-byte the same result in a few ms, so stored hashes stay valid. If the JavaScript version
// ever disagrees with Utilities (self-check below), the slow Utilities loop is used instead,
// so a bug here can only make sign-in slow — never lock anyone out.
function iteratedHashHex_(password, salt, rounds) {
  let bytes = sha256Bytes_(password, salt);
  if (rounds > 1 && jsSha256Ok_()) {
    let digest = bytes.map(b => b & 0xff);
    for (let i = 1; i < rounds; i++) digest = sha256Js_(digest);
    return toHex_(digest);
  }
  for (let i = 1; i < rounds; i++) {
    bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  }
  return toHex_(bytes);
}

let gttJsSha256Ok_ = null;

// Compares the JavaScript SHA-256 with Utilities once per request (two known inputs).
function jsSha256Ok_() {
  if (gttJsSha256Ok_ !== null) return gttJsSha256Ok_;
  try {
    const samples = [
      [],
      Array.from({ length: 32 }, (_, i) => (i * 37 + 11) & 0xff),
      Array.from({ length: 100 }, (_, i) => (i * 101 + 7) & 0xff)
    ];
    gttJsSha256Ok_ = samples.every(input => {
      const expected = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input.map(b => (b > 127 ? b - 256 : b)));
      return toHex_(sha256Js_(input)) === toHex_(expected);
    });
  } catch (err) {
    gttJsSha256Ok_ = false;
  }
  return gttJsSha256Ok_;
}

const GTT_SHA256_K = [
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
];

// Standard SHA-256 (FIPS 180-4). Input: array of byte values 0..255. Output: 32 values 0..255.
function sha256Js_(input) {
  const len = input.length;
  const blocks = Math.ceil((len + 9) / 64);
  const msg = new Array(blocks * 64).fill(0);
  for (let i = 0; i < len; i++) msg[i] = input[i];
  msg[len] = 0x80;
  const bitLen = len * 8;
  msg[msg.length - 4] = (bitLen >>> 24) & 0xff;
  msg[msg.length - 3] = (bitLen >>> 16) & 0xff;
  msg[msg.length - 2] = (bitLen >>> 8) & 0xff;
  msg[msg.length - 1] = bitLen & 0xff;

  const h = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
  const w = new Array(64);
  for (let b = 0; b < blocks; b++) {
    for (let t = 0; t < 16; t++) {
      const o = b * 64 + t * 4;
      w[t] = ((msg[o] << 24) | (msg[o + 1] << 16) | (msg[o + 2] << 8) | msg[o + 3]) | 0;
    }
    for (let t = 16; t < 64; t++) {
      const x = w[t - 15], y = w[t - 2];
      const s0 = ((x >>> 7) | (x << 25)) ^ ((x >>> 18) | (x << 14)) ^ (x >>> 3);
      const s1 = ((y >>> 17) | (y << 15)) ^ ((y >>> 19) | (y << 13)) ^ (y >>> 10);
      w[t] = (w[t - 16] + s0 + w[t - 7] + s1) | 0;
    }
    let a = h[0], c1 = h[1], c = h[2], d = h[3], e = h[4], f = h[5], g = h[6], hh = h[7];
    for (let t = 0; t < 64; t++) {
      const S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
      const ch = (e & f) ^ (~e & g);
      const t1 = (hh + S1 + ch + GTT_SHA256_K[t] + w[t]) | 0;
      const S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
      const maj = (a & c1) ^ (a & c) ^ (c1 & c);
      const t2 = (S0 + maj) | 0;
      hh = g; g = f; f = e; e = (d + t1) | 0; d = c; c = c1; c1 = a; a = (t1 + t2) | 0;
    }
    h[0] = (h[0] + a) | 0; h[1] = (h[1] + c1) | 0; h[2] = (h[2] + c) | 0; h[3] = (h[3] + d) | 0;
    h[4] = (h[4] + e) | 0; h[5] = (h[5] + f) | 0; h[6] = (h[6] + g) | 0; h[7] = (h[7] + hh) | 0;
  }
  const out = [];
  h.forEach(v => out.push((v >>> 24) & 0xff, (v >>> 16) & 0xff, (v >>> 8) & 0xff, v & 0xff));
  return out;
}

function hashPassword_(password, salt) {
  const rounds = GTT_CONFIG.HASH_ITERATIONS;
  return GTT_HASH_ITER_PREFIX + rounds + '$' + iteratedHashHex_(password, salt, rounds);
}

function verifyPassword_(password, user) {
  const stored = String(user.password_hash || '');
  const salt = String(user.password_salt || '');
  let expected;
  if (stored.indexOf(GTT_HASH_ITER_PREFIX) === 0) {
    const rounds = Number(stored.split('$')[1]);
    // Bound the work so a hand-edited cell cannot make sign-in hang.
    if (!(rounds >= 1 && rounds <= 20000)) return false;
    expected = GTT_HASH_ITER_PREFIX + rounds + '$' + iteratedHashHex_(password, salt, rounds);
  } else if (stored.indexOf(GTT_HASH_PREFIX) === 0) {
    expected = GTT_HASH_PREFIX + toHex_(sha256Bytes_(password, salt));
  } else {
    expected = Utilities.base64Encode(sha256Bytes_(password, salt));
  }
  return constantTimeEquals_(expected, stored);
}

function needsRehash_(user) {
  return String(user.password_hash || '').indexOf(GTT_HASH_ITER_PREFIX + GTT_CONFIG.HASH_ITERATIONS + '$') !== 0;
}

function cleanPassword_(value) {
  const password = String(value == null ? '' : value);
  if (password.length < GTT_CONFIG.PASSWORD_MIN_LENGTH) {
    throw new Error('รหัสผ่านต้องมีอย่างน้อย ' + GTT_CONFIG.PASSWORD_MIN_LENGTH + ' ตัวอักษร');
  }
  if (password.length > 100) throw new Error('รหัสผ่านยาวเกินไป (ไม่เกิน 100 ตัวอักษร)');
  return password;
}

// Changing a password also bumps session_version, which signs that user out everywhere.
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

function rowsCacheKey_(sheetName) { return 'GTT_ROWS_' + sheetName; }

function invalidateRows_(sheetName) {
  CacheService.getScriptCache().remove(rowsCacheKey_(sheetName));
}

// Cached read. Every write in this file calls invalidateRows_() so the next read is fresh.
function sheetRows_(sheetName) {
  const cache = CacheService.getScriptCache();
  const key = rowsCacheKey_(sheetName);
  const hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch (err) { /* fall through to a fresh read */ }
  }
  const rows = readSheetRows_(sheetName);
  try {
    cache.put(key, JSON.stringify(rows), GTT_CONFIG.CACHE_SECONDS);
  } catch (err) {
    // A cache entry is limited to 100 KB; very large sheets are simply read every time.
  }
  return rows;
}

function readSheetRows_(sheetName) {
  const sheet = getSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const tz = Session.getScriptTimeZone();
  const headers = values[0].map(value => String(value || '').trim());
  return values.slice(1).filter(row => !row.every(value => value === '' || value === null)).map(row => {
    const item = {};
    headers.forEach((header, index) => {
      if (!header) return;
      let value = row[index];
      if (value instanceof Date) value = Utilities.formatDate(value, tz, 'yyyy-MM-dd');
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
// (no auto-conversion to dates or numbers, e.g. '2026-11' stays text, leading zeros kept). Numbers use '0'.
function cellFormat_(value) {
  return typeof value === 'number' ? '0' : '@';
}

// Writes many rows in one call (much faster than one call per row).
function appendObjects_(sheetName, objects) {
  if (!objects.length) return;
  const sheet = getSheet_(sheetName);
  const map = headerMap_(sheet);
  const width = sheet.getLastColumn();
  const rows = objects.map(object => {
    const row = new Array(width).fill('');
    Object.keys(object).forEach(key => {
      if (map[key]) row[map[key] - 1] = toCellValue_(object[key]);
    });
    return row;
  });
  const firstRow = sheet.getLastRow() + 1;
  const neededLastRow = firstRow + rows.length - 1;
  // appendRow() added rows automatically; getRange() does not, so grow the sheet when it is full.
  if (neededLastRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), neededLastRow - sheet.getMaxRows() + 100);
  }
  const target = sheet.getRange(firstRow, 1, rows.length, width);
  target.setNumberFormats(rows.map(row => row.map(cellFormat_)));
  target.setValues(rows);
  invalidateRows_(sheetName);
}

function appendObject_(sheetName, object) {
  appendObjects_(sheetName, [object]);
}

// Reads the row once and writes it back in one call (one call per cell was noticeably slower).
// Only the cells being changed get a new number format, so untouched cells keep theirs.
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
  const keys = Object.keys(values);
  if (keys.length) {
    const cols = keys.map(key => map[key]);
    const first = Math.min.apply(null, cols), last = Math.max.apply(null, cols);
    const range = sheet.getRange(rowNumber, first, 1, last - first + 1);
    const row = range.getValues()[0];
    const formats = range.getNumberFormats()[0];
    keys.forEach(key => {
      row[map[key] - first] = values[key];
      formats[map[key] - first] = cellFormat_(values[key]);
    });
    range.setNumberFormats([formats]);
    range.setValues([row]);
  }
  invalidateRows_(sheetName);
  return rowNumber;
}

// Deletes every row whose <column> equals <value>. Reads the column once, then deletes bottom-up
// so row numbers stay valid while deleting.
function deleteRowsWhere_(sheetName, column, value) {
  const sheet = getSheet_(sheetName);
  const col = headerMap_(sheet)[column];
  if (!col || sheet.getLastRow() <= 1) return 0;
  const values = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getDisplayValues();
  let deleted = 0;
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]) === String(value)) {
      sheet.deleteRow(i + 2);
      deleted++;
    }
  }
  if (deleted) invalidateRows_(sheetName);
  return deleted;
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
function nowText_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }
function thisMonth_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM'); }
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

// Like cleanId_, but the ID must be sent (used when it points at an existing record).
function requireId_(value, label) {
  const id = String(value == null ? '' : value).trim();
  if (!id || !/^[A-Za-z0-9_-]{1,64}$/.test(id)) throw new Error(label + ' is required.');
  return id;
}

function cleanDate_(value) {
  const text = String(value == null ? '' : value).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(text) ? text : today_();
}

function cleanMonth_(value) {
  const month = String(value == null ? '' : value).trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error('เดือนต้องอยู่ในรูปแบบ YYYY-MM');
  return month;
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

function cleanClientFields_(client) {
  const fields = {
    client_name: cleanText_(client.client_name, 'Client name', 120, true),
    logo_url: cleanUrl_(client.logo_url, 'Client logo URL'),
    description: cleanText_(client.description, 'Description', 1000, false),
    archived: bool_(client.archived)
  };
  // Only touch the team when the page sends it, so other saves never wipe it by accident.
  if (client.team_ids !== undefined) fields.team_ids = cleanTeamIds_(client.team_ids);
  return fields;
}

// -----------------------------------------------------------------------------
// CLIENT TEAMS — which STAFF may see which client
// -----------------------------------------------------------------------------
function parseTeamIds_(value) {
  return String(value || '').split(',').map(id => id.trim()).filter(Boolean);
}

// Accepts an array or a comma-separated string; keeps only IDs of users that exist.
function cleanTeamIds_(value) {
  const wanted = Array.isArray(value) ? value.map(String) : parseTeamIds_(value);
  const known = new Set(getUsers_().map(user => String(user.user_id)));
  const ids = [];
  wanted.forEach(id => {
    const clean = String(id).trim();
    if (/^[A-Za-z0-9_-]{1,64}$/.test(clean) && known.has(clean) && ids.indexOf(clean) === -1) ids.push(clean);
  });
  return ids.join(',');
}

function isClientVisible_(actor, client) {
  return actor.role === 'ADMIN' || parseTeamIds_(client.team_ids).indexOf(String(actor.user_id)) >= 0;
}

function visibleClients_(actor) {
  return getClients_().filter(client => isClientVisible_(actor, client));
}

function visibleClientIdSet_(actor) {
  return new Set(visibleClients_(actor).map(client => String(client.client_id)));
}

function visibleWorkspaces_(actor) {
  if (actor.role === 'ADMIN') return getAllWorkspaces_();
  const ids = visibleClientIdSet_(actor);
  return getAllWorkspaces_().filter(ws => ids.has(String(ws.client_id)));
}

// Handing a piece of work to someone puts them on that client's team, so they can always
// open the client's brief and links. Takes a map { client_id: [user_id, ...] } to write each
// client at most once.
function addToClientTeams_(byClient) {
  const clients = getClients_();
  Object.keys(byClient).forEach(clientId => {
    const client = clients.find(item => String(item.client_id) === String(clientId));
    if (!client) return;
    const team = parseTeamIds_(client.team_ids);
    const before = team.length;
    byClient[clientId].forEach(userId => {
      if (userId && team.indexOf(String(userId)) === -1) team.push(String(userId));
    });
    if (team.length !== before) {
      updateObject_(GTT_CONFIG.SHEETS.CLIENTS, 'client_id', client.client_id, { team_ids: team.join(',') });
    }
  });
}

function teamAdditionsFrom_(rows) {
  const byClient = {};
  rows.forEach(row => {
    if (!row.assignee_id) return;
    (byClient[row.client_id] = byClient[row.client_id] || []).push(row.assignee_id);
  });
  return byClient;
}

function backfillClientTeams_() {
  // On a brand-new sheet the Deliverables sheet does not exist yet, so there is nothing to copy.
  if (!getDatabaseSpreadsheet_().getSheetByName(GTT_CONFIG.SHEETS.DELIVERABLES)) return;
  addToClientTeams_(teamAdditionsFrom_(getDeliverables_()));
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
  const existingIds = new Set(getAllWorkspaces_().map(ws => String(ws.workspace_id)));
  const workspaceIds = new Set();
  workspaces.forEach(ws => {
    if (workspaceIds.has(ws.workspace_id) || existingIds.has(ws.workspace_id)) {
      throw new Error('Workspace ID already exists.');
    }
    workspaceIds.add(ws.workspace_id);
  });

  appendObject_(GTT_CONFIG.SHEETS.CLIENTS, savedClient);
  appendObjects_(GTT_CONFIG.SHEETS.WORKSPACES, workspaces);
  return savedClient;
}

function updateClient_(client) {
  if (!client || !client.client_id) throw new Error('Client ID is required.');
  const patch = cleanClientFields_(client);
  updateObject_(GTT_CONFIG.SHEETS.CLIENTS, 'client_id', client.client_id, patch);
  return Object.assign({}, client, patch);
}

function deleteClient_(clientId) {
  const id = requireId_(clientId, 'Client ID');
  if (!deleteRowsWhere_(GTT_CONFIG.SHEETS.CLIENTS, 'client_id', id)) throw new Error('Client not found: ' + id);
  deleteRowsWhere_(GTT_CONFIG.SHEETS.WORKSPACES, 'client_id', id);
  deleteRowsWhere_(GTT_CONFIG.SHEETS.DELIVERABLES, 'client_id', id);
  deleteRowsWhere_(GTT_CONFIG.SHEETS.PACKAGES, 'client_id', id);
  return { deleted: id };
}

// Saves an uploaded logo into the company Drive folder and returns a link the page can show.
// The browser shrinks the image first, so a logo is normally well under 200 KB.
function uploadClientLogo_(payload) {
  const dataUrl = String(payload && payload.data_url || '');
  if (dataUrl.length > 3000000) throw new Error('รูปใหญ่เกินไป กรุณาใช้รูปที่เล็กกว่า 2 MB');
  const match = dataUrl.match(/^data:image\/(png|jpeg|webp);base64,([A-Za-z0-9+/=]+)$/);
  if (!match) throw new Error('โลโก้ต้องเป็นไฟล์รูป PNG, JPG หรือ WebP');

  const bytes = Utilities.base64Decode(match[2]);
  const extension = match[1] === 'jpeg' ? 'jpg' : match[1];
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const blob = Utilities.newBlob(bytes, 'image/' + match[1], 'client-logo-' + stamp + '-' + makeId_('img').slice(4, 12) + '.' + extension);

  const file = DriveApp.getFolderById(GTT_CONFIG.LOGO_FOLDER_ID).createFile(blob);
  try {
    // The page shows the logo with a plain <img>, which only works for files anyone with the link can view.
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    file.setTrashed(true);
    throw new Error('อัปโหลดได้แต่ตั้งค่าแชร์ลิงก์ไม่ได้ (Drive ขององค์กรอาจห้ามแชร์): ' + (err && err.message ? err.message : err));
  }
  return { file_id: file.getId(), url: 'https://lh3.googleusercontent.com/d/' + file.getId() };
}

// -----------------------------------------------------------------------------
// WORKSPACES
// -----------------------------------------------------------------------------
function getAllWorkspaces_() { return sheetRows_(GTT_CONFIG.SHEETS.WORKSPACES); }

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
  const id = requireId_(workspaceId, 'Workspace ID');
  if (!deleteRowsWhere_(GTT_CONFIG.SHEETS.WORKSPACES, 'workspace_id', id)) throw new Error('Workspace not found: ' + id);
  return { deleted: id };
}

// -----------------------------------------------------------------------------
// DELIVERABLES — pieces of work per client per month
// -----------------------------------------------------------------------------
function getDeliverables_() {
  return sheetRows_(GTT_CONFIG.SHEETS.DELIVERABLES).map(item => {
    // If someone typed a month by hand, Sheets may have turned it into a date ('2026-11-01').
    item.month = String(item.month || '').slice(0, 7);
    item.status = String(item.result_url || '').trim() ? GTT_STATUS.DONE : GTT_STATUS.TODO;
    item.quantity = quantityOf_(item.quantity);
    return item;
  });
}

function findDeliverable_(deliverableId) {
  const id = requireId_(deliverableId, 'Deliverable ID');
  const item = getDeliverables_().find(row => String(row.deliverable_id) === id);
  if (!item) throw new Error('ไม่พบงานชิ้นนี้ (อาจถูกลบไปแล้ว)');
  return item;
}

// ADMIN sees every piece in the month. STAFF sees only their own pieces of clients they are on
// the team of: the chosen month, plus anything older that is not sent yet, so unfinished work
// is never hidden. Taking someone off a client's team hides that client's work from them too.
function deliverablesForUser_(actor, month) {
  const all = getDeliverables_();
  if (actor.role === 'ADMIN') return all.filter(item => item.month === month);
  const visible = visibleClientIdSet_(actor);
  return all.filter(item => String(item.assignee_id) === String(actor.user_id) && visible.has(String(item.client_id)) &&
    (item.month === month || (item.month < month && item.status !== GTT_STATUS.DONE)));
}

// Blank or broken cells (older one-piece-per-row data) count as 1 piece.
function quantityOf_(value) {
  const n = Math.floor(Number(value));
  return n >= 1 && n <= 999 ? n : 1;
}

function cleanQuantity_(value) {
  const n = Number(value);
  if (!(Math.floor(n) === n && n >= 1 && n <= GTT_MAX_PIECES_PER_ADD)) {
    throw new Error('จำนวนชิ้นต้องอยู่ระหว่าง 1–' + GTT_MAX_PIECES_PER_ADD);
  }
  return n;
}

function cleanDeliverableType_(value) {
  const type = String(value == null ? '' : value).trim();
  if (GTT_DELIVERABLE_TYPES.indexOf(type) === -1) throw new Error('ประเภทงานไม่ถูกต้อง');
  return type;
}

// Blank = not assigned yet. Otherwise it must be an active team member.
function cleanAssignee_(value) {
  const id = String(value == null ? '' : value).trim();
  if (!id) return '';
  const user = getUsers_().find(item => String(item.user_id) === id && item.active);
  if (!user) throw new Error('ไม่พบผู้รับผิดชอบที่เลือก หรือบัญชีถูกปิดใช้งาน');
  return user.user_id;
}

function requireClientExists_(clientId) {
  const id = requireId_(clientId, 'Client ID');
  if (findRowById_(GTT_CONFIG.SHEETS.CLIENTS, 'client_id', id) < 0) throw new Error('The selected client does not exist.');
  return id;
}

function newDeliverable_(fields) {
  const now = nowText_();
  return Object.assign({
    deliverable_id: makeId_('dlv'),
    status: GTT_STATUS.TODO,
    result_url: '',
    review_note: '',
    submitted_by: '',
    submitted_at: '',
    reviewed_at: '',
    created_at: now,
    updated_at: now
  }, fields);
}

// Adds one set of <count> pieces of one type, e.g. "VDO x 4", as a single row with one result link.
// Returns an array (one row) so the page handles it the same way as copied rows.
function createDeliverables_(payload) {
  const clientId = requireClientExists_(payload.client_id);
  const month = cleanMonth_(payload.month);
  const type = cleanDeliverableType_(payload.type);
  const quantity = cleanQuantity_(payload.count);
  const assigneeId = cleanAssignee_(payload.assignee_id);
  const title = cleanText_(payload.title, 'ชื่องาน', 120, false);

  const rows = [newDeliverable_({
    client_id: clientId,
    month: month,
    type: type,
    title: title,
    assignee_id: assigneeId,
    quantity: quantity
  })];
  appendObjects_(GTT_CONFIG.SHEETS.DELIVERABLES, rows);
  addToClientTeams_(teamAdditionsFrom_(rows));
  return rows;
}

function updateDeliverable_(payload) {
  const existing = findDeliverable_(payload.deliverable_id);
  const patch = { updated_at: nowText_() };
  if (payload.title !== undefined) patch.title = cleanText_(payload.title, 'ชื่องาน', 120, false);
  if (payload.type !== undefined) patch.type = cleanDeliverableType_(payload.type);
  if (payload.quantity !== undefined) patch.quantity = cleanQuantity_(payload.quantity);
  if (payload.assignee_id !== undefined) patch.assignee_id = cleanAssignee_(payload.assignee_id);
  updateObject_(GTT_CONFIG.SHEETS.DELIVERABLES, 'deliverable_id', existing.deliverable_id, patch);
  if (patch.assignee_id) addToClientTeams_({ [existing.client_id]: [patch.assignee_id] });
  return Object.assign({}, existing, patch);
}

function deleteDeliverable_(deliverableId) {
  const id = requireId_(deliverableId, 'Deliverable ID');
  if (!deleteRowsWhere_(GTT_CONFIG.SHEETS.DELIVERABLES, 'deliverable_id', id)) throw new Error('ไม่พบงานชิ้นนี้ (อาจถูกลบไปแล้ว)');
  return { deleted: id };
}

// Monthly clients usually get the same set of work every month. Copies the sets (type, title,
// number of pieces, person responsible) from one month to another, for one client or for every active client.
// A client that already has work in the target month is skipped, so pressing twice never doubles it.
function copyDeliverables_(payload) {
  const fromMonth = cleanMonth_(payload.from_month);
  const toMonth = cleanMonth_(payload.to_month);
  if (fromMonth === toMonth) throw new Error('เดือนต้นทางและปลายทางต้องไม่ใช่เดือนเดียวกัน');
  const onlyClientId = payload.client_id ? requireClientExists_(payload.client_id) : '';

  const activeClientIds = new Set(getClients_().filter(client => !client.archived).map(client => String(client.client_id)));
  const activeUserIds = new Set(getUsers_().filter(user => user.active).map(user => String(user.user_id)));
  const all = getDeliverables_();
  const clientsWithTarget = new Set(all.filter(item => item.month === toMonth).map(item => String(item.client_id)));

  const rows = [];
  const skipped = [];
  all.filter(item => item.month === fromMonth).forEach(item => {
    const clientId = String(item.client_id);
    if (onlyClientId && clientId !== onlyClientId) return;
    if (!onlyClientId && !activeClientIds.has(clientId)) return;
    if (clientsWithTarget.has(clientId)) {
      if (skipped.indexOf(clientId) === -1) skipped.push(clientId);
      return;
    }
    rows.push(newDeliverable_({
      client_id: clientId,
      month: toMonth,
      type: cleanChoice_(item.type, GTT_DELIVERABLE_TYPES, 'other'),
      title: isFormulaLike_(String(item.title || '')) ? '' : String(item.title || ''),
      assignee_id: activeUserIds.has(String(item.assignee_id)) ? String(item.assignee_id) : '',
      quantity: quantityOf_(item.quantity)
    }));
  });
  appendObjects_(GTT_CONFIG.SHEETS.DELIVERABLES, rows);
  addToClientTeams_(teamAdditionsFrom_(rows));
  return { created: rows, skipped_client_ids: skipped };
}

// The only write STAFF can do: attach the finished link to a piece assigned to them.
// Sending the link marks the whole set done. Sending again replaces the link (e.g. a wrong link was pasted).
function submitDeliverable_(payload, actor) {
  const item = findDeliverable_(payload.deliverable_id);
  const isAdmin = actor.role === 'ADMIN';
  if (!isAdmin && (String(item.assignee_id) !== String(actor.user_id) || !visibleClientIdSet_(actor).has(String(item.client_id)))) {
    throw new Error('งานชิ้นนี้ไม่ได้มอบหมายให้คุณ');
  }
  const url = cleanUrl_(payload.result_url, 'ลิงก์ผลงาน');
  if (!url) throw new Error('กรุณาวางลิงก์ผลงาน');

  const now = nowText_();
  const patch = {
    status: GTT_STATUS.DONE,
    result_url: url,
    submitted_by: actor.user_id,
    submitted_at: now,
    updated_at: now
  };
  updateObject_(GTT_CONFIG.SHEETS.DELIVERABLES, 'deliverable_id', item.deliverable_id, patch);
  return Object.assign({}, item, patch);
}

// -----------------------------------------------------------------------------
// PACKAGES — contracts that create the same work every month for N months
// -----------------------------------------------------------------------------
function addMonths_(month, n) {
  const parts = String(month).split('-');
  const total = Number(parts[0]) * 12 + (Number(parts[1]) - 1) + n;
  return Math.floor(total / 12) + '-' + ('0' + (total % 12 + 1)).slice(-2);
}

function packageEndMonth_(pkg) {
  return addMonths_(pkg.start_month, Number(pkg.months) - 1);
}

function getPackages_() {
  return sheetRows_(GTT_CONFIG.SHEETS.PACKAGES).map(row => {
    let items = [];
    try { items = JSON.parse(String(row.items_json || '[]')); } catch (err) { items = []; }
    return {
      package_id: String(row.package_id),
      client_id: String(row.client_id),
      start_month: String(row.start_month || '').slice(0, 7),
      months: Number(row.months) || 1,
      items: Array.isArray(items) ? items : [],
      active: bool_(row.active),
      last_generated_month: String(row.last_generated_month || '').slice(0, 7),
      created_at: row.created_at,
      updated_at: row.updated_at
    };
  });
}

function cleanPackageItems_(value) {
  const list = Array.isArray(value) ? value : [];
  if (!list.length) throw new Error('แพ็กเกจต้องมีงานอย่างน้อย 1 รายการ');
  if (list.length > GTT_MAX_PACKAGE_ITEMS) throw new Error('งานในแพ็กเกจมีได้ไม่เกิน ' + GTT_MAX_PACKAGE_ITEMS + ' รายการ');
  return list.map(item => ({
    type: cleanDeliverableType_(item && item.type),
    quantity: cleanQuantity_(item && item.quantity),
    assignee_id: cleanAssignee_(item && item.assignee_id),
    title: cleanText_(item && item.title, 'ชื่องาน', 120, false)
  }));
}

function cleanPackageMonths_(value) {
  const n = Number(value);
  if (!(Math.floor(n) === n && n >= 1 && n <= GTT_MAX_PACKAGE_MONTHS)) {
    throw new Error('จำนวนเดือนต้องอยู่ระหว่าง 1–' + GTT_MAX_PACKAGE_MONTHS);
  }
  return n;
}

// Create (no package_id) or edit a package. Edits apply to months not created yet; work already
// created for this month stays as it is (the admin can still edit those pieces directly).
// If the package covers this month and this month was not created yet, it is created right away.
function savePackage_(payload) {
  const start = cleanMonth_(payload.start_month);
  const months = cleanPackageMonths_(payload.months);
  const items = cleanPackageItems_(payload.items);
  const active = payload.active === undefined ? true : bool_(payload.active);
  const now = nowText_();
  let pkgId = String(payload.package_id || '').trim();

  if (pkgId) {
    pkgId = requireId_(pkgId, 'Package ID');
    const existing = getPackages_().find(pkg => pkg.package_id === pkgId);
    if (!existing) throw new Error('ไม่พบแพ็กเกจนี้ (อาจถูกลบไปแล้ว)');
    updateObject_(GTT_CONFIG.SHEETS.PACKAGES, 'package_id', pkgId, {
      start_month: start, months: months, items_json: JSON.stringify(items), active: active, updated_at: now
    });
  } else {
    const clientId = requireClientExists_(payload.client_id);
    pkgId = makeId_('pkg');
    appendObject_(GTT_CONFIG.SHEETS.PACKAGES, {
      package_id: pkgId, client_id: clientId, start_month: start, months: months,
      items_json: JSON.stringify(items), active: active, last_generated_month: '', created_at: now, updated_at: now
    });
  }
  const created = generateDueWork_(pkgId);
  return { package: getPackages_().find(pkg => pkg.package_id === pkgId), created: created };
}

function deletePackage_(packageId) {
  const id = requireId_(packageId, 'Package ID');
  if (!deleteRowsWhere_(GTT_CONFIG.SHEETS.PACKAGES, 'package_id', id)) throw new Error('ไม่พบแพ็กเกจนี้ (อาจถูกลบไปแล้ว)');
  // Work already created from the package stays: it was real work for those months.
  return { deleted: id };
}

// Creates this month's work for every running package that has not been created yet.
// Must run under the script lock. Safe to call many times: last_generated_month stops repeats.
function generateDueWork_(onlyPackageId) {
  const month = thisMonth_();
  const activeClients = new Set(getClients_().filter(client => !client.archived).map(client => String(client.client_id)));
  const activeUsers = new Set(getUsers_().filter(user => user.active).map(user => String(user.user_id)));
  const rows = [];
  getPackages_().forEach(pkg => {
    if (onlyPackageId && pkg.package_id !== onlyPackageId) return;
    if (!pkg.active || !activeClients.has(pkg.client_id)) return;
    if (!/^\d{4}-\d{2}$/.test(pkg.start_month)) return;
    if (month < pkg.start_month || month > packageEndMonth_(pkg)) return;
    if (pkg.last_generated_month && pkg.last_generated_month >= month) return;
    const pkgRows = pkg.items.map(item => newDeliverable_({
      client_id: pkg.client_id,
      month: month,
      type: cleanChoice_(item.type, GTT_DELIVERABLE_TYPES, 'other'),
      title: isFormulaLike_(String(item.title || '')) ? '' : String(item.title || '').slice(0, 120),
      // Someone who left (account disabled) is not given new work; the admin sees it unassigned.
      assignee_id: activeUsers.has(String(item.assignee_id)) ? String(item.assignee_id) : '',
      quantity: quantityOf_(item.quantity),
      package_id: pkg.package_id
    }));
    // Mark first, then write: if writing fails half-way the month is still not doubled next time.
    updateObject_(GTT_CONFIG.SHEETS.PACKAGES, 'package_id', pkg.package_id, { last_generated_month: month });
    rows.push.apply(rows, pkgRows);
  });
  if (rows.length) {
    appendObjects_(GTT_CONFIG.SHEETS.DELIVERABLES, rows);
    addToClientTeams_(teamAdditionsFrom_(rows));
  }
  return rows;
}

function monthlyWorkKey_() { return 'GTT_PKG_DONE_' + thisMonth_(); }

function markMonthlyWorkDone_() {
  CacheService.getScriptCache().put(monthlyWorkKey_(), '1', 21600);
}

// Called on every request: a cheap cache check, and once per month (or every 6 hours at most)
// it creates the due work. The 5-minute warm-up normally does this first, so people rarely wait.
function ensureMonthlyWork_() {
  if (CacheService.getScriptCache().get(monthlyWorkKey_())) return;
  try {
    withLock_(() => { generateDueWork_(); markMonthlyWorkDone_(); });
  } catch (err) {
    // Creating package work must never block sign-in or other requests; it is retried next request.
  }
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

// Admin types the username and password for the new member on the Settings page.
function createUser_(user) {
  const name = cleanText_(user.name, 'Team member name', 100, true);
  const username = normalizeUsername_(user.username);
  const email = cleanEmail_(user.email);
  const role = String(user.role || 'STAFF').toUpperCase();
  const password = cleanPassword_(user.password);
  if (!/^[a-z0-9][a-z0-9._-]{2,29}$/.test(username)) {
    throw new Error('Username must be 3–30 characters and use only letters, numbers, dots, underscores or hyphens.');
  }
  if (role !== 'ADMIN' && role !== 'STAFF') throw new Error('Role must be ADMIN or STAFF.');
  const users = getUsers_();
  if (users.some(existing => normalizeUsername_(existing.username) === username)) {
    throw new Error('This username is already registered.');
  }
  if (email && users.some(existing => normalizeEmail_(existing.email) === email)) throw new Error('This email is already registered.');

  let userId = cleanId_(user.user_id, 'usr');
  if (findRowById_(GTT_CONFIG.SHEETS.USERS, 'user_id', userId) >= 0) userId = makeId_('usr');
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
    password_hash: hashPassword_(password, passwordSalt)
  };
  appendObject_(GTT_CONFIG.SHEETS.USERS, saved);
  return publicUser_(saved);
}

function adminSetUserPassword_(payload) {
  const userId = requireId_(payload.user_id, 'User ID');
  const user = getUsers_().find(item => String(item.user_id) === userId);
  if (!user) throw new Error('User not found.');
  setUserPassword_(user.user_id, cleanPassword_(payload.password));
  return publicUser_(user);
}

function updateUser_(user, actor) {
  if (!user || !user.user_id) throw new Error('User ID is required.');
  const users = getUsers_();
  const existing = users.find(item => String(item.user_id) === String(user.user_id));
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
  if (patch.username && users.some(item => item.user_id !== existing.user_id && normalizeUsername_(item.username) === patch.username)) {
    throw new Error('This username is already registered.');
  }
  if (patch.email && users.some(item => item.user_id !== existing.user_id && normalizeEmail_(item.email) === patch.email)) {
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
  const id = requireId_(userId, 'User ID');
  if (id === String(actor.user_id)) throw new Error('You cannot remove your own account.');
  const users = getUsers_();
  const existing = users.find(user => String(user.user_id) === id);
  if (!existing) throw new Error('User not found.');
  if (existing.active && existing.role === 'ADMIN') {
    const otherAdmins = users.filter(user => user.active && user.role === 'ADMIN' && user.user_id !== existing.user_id);
    if (!otherAdmins.length) throw new Error('At least one active administrator must remain.');
  }
  deleteRowsWhere_(GTT_CONFIG.SHEETS.USERS, 'user_id', id);
  return { deleted: id };
}

function ensureAtLeastOneAdmin_(changingUserId, resultingRole, resultingActive) {
  const users = getUsers_();
  const activeAdminsAfter = users.filter(user => user.active && user.role === 'ADMIN' && user.user_id !== changingUserId).length
    + ((resultingActive && resultingRole === 'ADMIN') ? 1 : 0);
  if (activeAdminsAfter < 1) throw new Error('At least one active administrator must remain.');
}
