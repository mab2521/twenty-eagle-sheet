/**
 * Twenty Eagle case pipeline — Google Sheet side.
 *
 *   1. doPost()                 receives rows from sheet_sync.py (scraper -> Sheet)
 *      Each case is filed on the tab for its session (e.g. "September Session 25");
 *      cases with no session yet stay on the "Cases" tab.
 *   2. A human reviews the row, fills the yellow columns, ticks "Approved"
 *   3. createWordPressDrafts()  turns approved rows into WordPress DRAFTS
 *   4. A human reviews the draft in WordPress and publishes it there
 *
 * This script never publishes anything: post status is hard-coded to "draft".
 *
 * Script Properties (Project Settings > Script Properties):
 *   SYNC_TOKEN        any long random string; the same value goes in SHEET_SYNC_TOKEN
 *   WP_URL            e.g. https://example.com   (no trailing slash)
 *   WP_USER           WordPress username
 *   WP_APP_PASSWORD   from WordPress > Users > Profile > Application Passwords
 *   WP_POST_TYPE      optional; leave out to create Pages (the default), or "posts" for Posts
 */

const SHEET_NAME = 'Cases';   // catch-all tab for cases with no session yet

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];
// Session tabs are named like the Session column: "September Session 25"
const SESSION_TAB = new RegExp('^(' + MONTH_NAMES.join('|') + ') Session (\\d{2})$');

// group: scraper = overwritten on every sync
//        human   = never overwritten by sync (blank cells may be pre-filled, see SEEDED)
//        wp      = written by createWordPressDrafts()
// verbatim: true = stored character for character as received (see writeVerbatim_)
// keepIfBlank: true = a sync with no value leaves the existing cell alone
const COLUMNS = [
  { key: 'case_key',          header: 'Case key',              group: 'scraper', width: 160 },
  { key: 'title',             header: 'Case title',            group: 'scraper', width: 240 },
  { key: 'decision_date',     header: 'Decision date',         group: 'scraper', width: 100, keepIfBlank: true },
  { key: 'decision_type',     header: 'Decision type',         group: 'scraper', width: 100, keepIfBlank: true },
  { key: 'argument_date',     header: 'Argument date',         group: 'scraper', width: 100, keepIfBlank: true },
  { key: 'argument_no',       header: 'Argument no.',          group: 'scraper', width: 90, keepIfBlank: true },
  { key: 'session',           header: 'Session',               group: 'scraper', width: 150, keepIfBlank: true },
  { key: 'apl_number',        header: 'APL no.',               group: 'scraper', width: 120 },
  { key: 'category',          header: 'Category',              group: 'scraper', width: 180 },
  { key: 'issues_courtpass',  header: 'Issues (CourtPASS)',    group: 'scraper', width: 380, verbatim: true },
  { key: 'ji_status',         header: 'JI status',             group: 'scraper', width: 100 },
  { key: 'ji_filings',        header: 'JI docket entries',     group: 'scraper', width: 220 },
  { key: 'files_available',   header: 'Files available',       group: 'scraper', width: 280 },
  { key: 'files_unavailable', header: 'Files not available',   group: 'scraper', width: 220 },
  { key: 'url_check',         header: 'Link check',            group: 'scraper', width: 200 },
  { key: 'last_synced',       header: 'Last synced',           group: 'scraper', width: 130 },

  { key: 'blog_summary',      header: 'Blog summary (edit me)', group: 'human', width: 380 },
  { key: 'decision_url',      header: 'Decision URL',          group: 'human', width: 200 },
  { key: 'video_url',         header: 'Argument video URL',    group: 'human', width: 200 },
  { key: 'transcript_url',    header: 'Transcript URL',        group: 'human', width: 200 },
  { key: 'merits_briefs',         header: 'NYCA Merits Briefs',    group: 'human', width: 280 },
  { key: 'other_materials',       header: 'NYCA Other Materials',  group: 'human', width: 280 },
  { key: 'lower_court_decision',  header: 'Lower Court Decision',  group: 'human', width: 280 },
  { key: 'lower_court_materials', header: 'Lower Court Materials', group: 'human', width: 280 },
  { key: 'request_jl',        header: 'Request JL?',           group: 'human', width: 90,  checkbox: true },
  { key: 'notes',             header: 'Reviewer notes',        group: 'human', width: 240 },
  { key: 'reviewed_by',       header: 'Reviewed by',           group: 'human', width: 100 },
  { key: 'approved',          header: 'Approved',              group: 'human', width: 80,  checkbox: true },

  { key: 'wp_post_id',        header: 'WP post ID',            group: 'wp', width: 90 },
  { key: 'wp_edit_link',      header: 'WP edit link',          group: 'wp', width: 240 },
  { key: 'wp_status',         header: 'WP status',             group: 'wp', width: 220 },
];

// Human columns the sync may pre-fill, but only while the cell is still blank
const SEEDED = ['blog_summary', 'decision_url', 'video_url', 'transcript_url',
  'merits_briefs', 'other_materials', 'lower_court_decision', 'lower_court_materials'];

// The "Select one..." dropdown on each case page, in template order.
// Each column holds one document per line: "Label | URL" (no URL = plain text).
const DROPDOWN_SECTIONS = [
  { key: 'merits_briefs',         value: 'option1', label: 'NYCA Merits Briefs' },
  { key: 'other_materials',       value: 'option2', label: 'NYCA Other Materials' },
  { key: 'lower_court_decision',  value: 'option3', label: 'Lower Court Decision' },
  { key: 'lower_court_materials', value: 'option4', label: 'Lower Court Materials' },
];

const GROUP_COLORS = { scraper: '#d9e2f3', human: '#fff2cc', wp: '#e2efda' };

// ───────────────────────── Menu & setup ─────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Twenty Eagle')
    .addItem('Create WordPress drafts for approved rows', 'createWordPressDrafts')
    .addSeparator()
    .addItem('Move cases into session tabs', 'organizeSessionTabs')
    .addItem('Test WordPress connection', 'testWordPressConnection')
    .addItem('Set up sheet', 'setupSheet')
    .addToUi();
}

/** Creates the "Cases" tab, and adds any columns a newer version of this script
 *  introduced to it and every session tab. Safe to re-run: data and edits are left alone. */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME, 0);
  prepareCaseSheet_(sheet);
  caseSheets_().forEach(s => { if (s.getName() !== SHEET_NAME) prepareCaseSheet_(s); });
  return sheet;
}

/** Headers, colours, widths, text formatting and warnings for one case tab. */
function prepareCaseSheet_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS.map(c => c.header)]);
  } else {
    COLUMNS.forEach((c, i) => {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
      if (headers.indexOf(c.header) !== -1) return;
      if (i < sheet.getLastColumn()) sheet.insertColumnBefore(i + 1);
      sheet.getRange(1, i + 1).setValue(c.header);
    });
  }

  const col = columnMap_(sheet);
  const rows = sheet.getMaxRows();
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  COLUMNS.forEach(c => {
    const n = col[c.key];
    sheet.getRange(1, n).setFontWeight('bold').setWrap(true)
      .setVerticalAlignment('middle').setBackground(GROUP_COLORS[c.group]);
    sheet.setColumnWidth(n, c.width);
    const body = sheet.getRange(2, n, rows - 1, 1).setWrap(true).setVerticalAlignment('top');
    // Plain text everywhere except checkboxes, so dates stay exactly as scraped
    if (!c.checkbox) body.setNumberFormat('@');
    // Soft warning on scraper-owned columns (edits there are overwritten at the next sync)
    if (c.group === 'scraper') {
      sheet.getRange(1, n, rows, 1).protect().setDescription('Filled by the scraper').setWarningOnly(true);
    }
  });
  return sheet;
}

function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return setupSheet();
  return ensureReady_(sheet);
}

/** Adds any missing columns to a case tab before it is used. */
function ensureReady_(sheet) {
  if (sheet.getLastRow() < 1) return prepareCaseSheet_(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  return COLUMNS.every(c => headers.indexOf(c.header) !== -1) ? sheet : prepareCaseSheet_(sheet);
}

/** The "Cases" tab plus every session tab. */
function caseSheets_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(s => s.getName() === SHEET_NAME || SESSION_TAB.test(s.getName()));
}

/** "September Session 25" -> 202508, for keeping tabs in date order. */
function sessionOrder_(name) {
  const m = SESSION_TAB.exec(name);
  return m ? (2000 + Number(m[2])) * 100 + MONTH_NAMES.indexOf(m[1]) : -1;
}

/** The tab for a session, created in date order after "Cases" if it doesn't exist.
 *  A blank or unrecognised session means the "Cases" tab. */
function sessionSheet_(session) {
  if (!SESSION_TAB.test(session || '')) return getSheet_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName(session);
  if (existing) return ensureReady_(existing);
  const order = sessionOrder_(session);
  let pos = 0;
  ss.getSheets().forEach((x, i) => {
    const name = x.getName();
    if (name === SHEET_NAME || (SESSION_TAB.test(name) && sessionOrder_(name) < order)) pos = i + 1;
  });
  return prepareCaseSheet_(ss.insertSheet(session, pos));
}

/** Where every case currently is: case key -> { sheet, row }. */
function caseIndex_(colsOf) {
  const index = {};
  caseSheets_().forEach(sheet => {
    ensureReady_(sheet);
    const last = sheet.getLastRow();
    if (last < 2) return;
    sheet.getRange(2, colsOf(sheet).case_key, last - 1, 1).getDisplayValues()
      .forEach((r, i) => { if (r[0]) index[r[0]] = { sheet: sheet, row: i + 2 }; });
  });
  return index;
}

/** Moves one case row to another tab, keeping every value, tick and reviewer edit.
 *  Updates row numbers in the index for rows below it on the old tab. */
function moveRow_(at, target, colsOf, index) {
  const from = at.sheet, fromCol = colsOf(from), toCol = colsOf(target);
  const values = from.getRange(at.row, 1, 1, from.getLastColumn()).getValues()[0];
  const row = target.getLastRow() + 1;
  COLUMNS.filter(c => c.checkbox).forEach(c => target.getRange(row, toCol[c.key]).insertCheckboxes());
  const out = new Array(target.getLastColumn()).fill('');
  COLUMNS.forEach(c => { out[toCol[c.key] - 1] = values[fromCol[c.key] - 1]; });
  target.getRange(row, 1, 1, out.length).setValues([out]);
  COLUMNS.filter(c => c.verbatim)
    .forEach(c => writeVerbatim_(target.getRange(row, toCol[c.key]), values[fromCol[c.key] - 1]));
  from.deleteRow(at.row);
  Object.keys(index).forEach(k => {
    const x = index[k];
    if (x.sheet.getName() === from.getName() && x.row > at.row) x.row--;
  });
  return { sheet: target, row: row };
}

/** Menu item: moves every row on "Cases" that has a session onto its session tab.
 *  Safe to re-run; if it stops partway on a large sheet, run it again. */
function organizeSessionTabs() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const colCache = {};
    const colsOf = s => colCache[s.getName()] || (colCache[s.getName()] = columnMap_(s));
    const cases = getSheet_();
    const last = cases.getLastRow();
    if (last < 2) { ui.alert('No cases on the "Cases" tab to move.'); return; }
    const sessions = cases.getRange(2, colsOf(cases).session, last - 1, 1).getDisplayValues();
    let moved = 0, kept = 0;
    sessions.forEach((r, i) => {
      const session = r[0].trim();
      if (!SESSION_TAB.test(session)) { kept++; return; }
      moveRow_({ sheet: cases, row: i + 2 - moved }, sessionSheet_(session), colsOf, {});
      moved++;
    });
    ui.alert('Moved ' + moved + ' case(s) to their session tabs.' +
      (kept ? '\n' + kept + ' case(s) have no session yet and stay on "Cases".' : ''));
  } finally {
    lock.releaseLock();
  }
}

/** Maps column key -> 1-based column number by reading the header row,
 *  so columns can be reordered in the Sheet without breaking the script. */
function columnMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};
  COLUMNS.forEach(c => {
    const i = headers.indexOf(c.header);
    if (i === -1) throw new Error('Missing column "' + c.header + '". Run Twenty Eagle > Set up sheet.');
    map[c.key] = i + 1;
  });
  return map;
}

// ───────────────────────── 1. Scraper -> Sheet ─────────────────────────

/** Lets you confirm in a browser that the deployed version is the current one. */
function doGet() {
  return ContentService.createTextOutput('Twenty Eagle sync endpoint is live (v6: session tabs)');
}

function doPost(e) {
  let out;
  try {
    const body = JSON.parse(e.postData.contents);
    const token = PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN');
    if (!token || body.token !== token) throw new Error('Bad or missing token');
    out = upsertCases_(body.cases || []);
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/** New case -> new row. Known case -> refresh scraper columns, and pre-fill
 *  any SEEDED human cell that is still blank. Human entries are never replaced. */
function upsertCases_(cases) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const colCache = {};
    const colsOf = s => colCache[s.getName()] || (colCache[s.getName()] = columnMap_(s));
    const scraperCols = COLUMNS.filter(c => c.group === 'scraper');
    const index = caseIndex_(colsOf);

    const added = [], updated = [], issueMismatches = [];
    cases.forEach(c => {
      if (!c.case_key) return;
      let at = index[c.case_key];
      // A blank session in the sync keeps the case's existing session
      const session = c.session ||
        (at ? at.sheet.getRange(at.row, colsOf(at.sheet).session).getDisplayValue().trim() : '');
      const target = sessionSheet_(session);
      if (at && at.sheet.getName() !== target.getName()) at = moveRow_(at, target, colsOf, index);
      const isNew = !at;
      if (isNew) {
        at = { sheet: target, row: target.getLastRow() + 1 };
        index[c.case_key] = at;
      }
      const sheet = at.sheet, row = at.row, col = colsOf(sheet);
      scraperCols.forEach(sc => {
        if (sc.keepIfBlank && !c[sc.key]) return;
        const cell = sheet.getRange(row, col[sc.key]);
        if (!sc.verbatim) { cell.setValue(safeText_(c[sc.key])); return; }
        if (!writeVerbatim_(cell, c[sc.key])) {
          issueMismatches.push(c.title || c.case_key);
          console.error('Sheet text differs from scraped text for ' + sc.key + ': ' + c.case_key);
        }
      });

      SEEDED.forEach(key => {
        const cell = sheet.getRange(row, col[key]);
        if (c[key] && !cell.getDisplayValue().trim()) cell.setValue(safeText_(c[key]));
      });

      if (isNew) {
        COLUMNS.filter(x => x.checkbox)
          .forEach(x => sheet.getRange(row, col[x.key]).insertCheckboxes());
        added.push(c.title);
      } else {
        updated.push(c.title);
      }
    });
    return { ok: true, added: added, updated: updated, issue_mismatches: issueMismatches };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Stores text exactly as received: no trimming, no added apostrophe, and no
 * conversion to a formula, number or date. Reads the cell back afterwards and
 * returns false if what the sheet holds is not identical to what was sent.
 */
function writeVerbatim_(cell, value) {
  const text = value == null ? '' : String(value);
  cell.setNumberFormat('@');
  if (text === '') {
    cell.clearContent();
  } else {
    cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(text).build());
  }
  return String(cell.getValue()) === text;
}

/** Stops scraped text from ever being interpreted as a spreadsheet formula. */
function safeText_(v) {
  const s = v == null ? '' : String(v);
  return /^[=+]/.test(s) ? "'" + s : s;
}

// ───────────────────────── 3. Sheet -> WordPress draft ─────────────────────────

function wpConfig_() {
  const p = PropertiesService.getScriptProperties();
  const cfg = {
    url: (p.getProperty('WP_URL') || '').replace(/\/+$/, ''),
    user: p.getProperty('WP_USER'),
    pass: p.getProperty('WP_APP_PASSWORD'),
    type: p.getProperty('WP_POST_TYPE') || 'pages',
  };
  if (!cfg.url || !cfg.user || !cfg.pass) {
    throw new Error('Add WP_URL, WP_USER and WP_APP_PASSWORD under Project Settings > Script Properties.');
  }
  cfg.auth = 'Basic ' + Utilities.base64Encode(cfg.user + ':' + cfg.pass);
  return cfg;
}

/**
 * Finds where this site's REST API lives instead of assuming /wp-json/.
 * WordPress announces the address in a "Link" header on its pages; if that is
 * missing, the two standard addresses are tried in turn. Returns e.g.
 * "https://example.com/wp-json/" or "https://example.com/?rest_route=/".
 */
function wpApiRoot_(cfg) {
  if (cfg.root) return cfg.root;
  const looksLikeApi = res => res.getResponseCode() === 200 &&
    /"namespaces"\s*:/.test(res.getContentText().slice(0, 5000));
  const get = url => UrlFetchApp.fetch(url, { muteHttpExceptions: true });

  let announced = '';
  try {
    const headers = get(cfg.url + '/').getAllHeaders();
    const key = Object.keys(headers).filter(k => k.toLowerCase() === 'link')[0];
    const m = /<([^>]+)>\s*;\s*rel="https:\/\/api\.w\.org\/"/.exec([].concat(headers[key] || []).join(', '));
    if (m) announced = m[1];
  } catch (err) { /* fall through to the standard addresses */ }

  if (/public-api\.wordpress\.com/.test(announced)) {
    throw new Error('This site is hosted on WordPress.com, which does not accept Application ' +
      'Passwords at this address. It needs a different connection method.');
  }
  const candidates = [announced, cfg.url + '/wp-json/', cfg.url + '/?rest_route=/'].filter(String);
  for (let i = 0; i < candidates.length; i++) {
    try {
      if (looksLikeApi(get(candidates[i]))) { cfg.root = candidates[i]; return cfg.root; }
    } catch (err) { /* try the next one */ }
  }
  throw new Error('Could not find the WordPress REST API at ' + cfg.url + '. Check that WP_URL is ' +
    'the address of the WordPress site itself, and that no security plugin has the REST API switched off.');
}

/** The wp-admin address for editing a post (handles WordPress in a subfolder). */
function wpEditLink_(cfg, id) {
  const root = wpApiRoot_(cfg);
  const base = root.indexOf('/wp-json/') !== -1 ? root.split('/wp-json/')[0] : root.split('/?rest_route=')[0];
  return base + '/wp-admin/post.php?post=' + id + '&action=edit';
}

/** Turns a WordPress error reply into one short readable line. */
function wpError_(res) {
  const body = res.getContentText();
  try {
    const j = JSON.parse(body);
    if (j && j.message) return 'WordPress ' + res.getResponseCode() + ': ' + j.message + (j.code ? ' (' + j.code + ')' : '');
  } catch (err) { /* not JSON */ }
  return 'WordPress ' + res.getResponseCode() + ': ' +
    body.replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 160);
}

function testWordPressConnection() {
  const ui = SpreadsheetApp.getUi();
  try {
    const cfg = wpConfig_();
    const root = wpApiRoot_(cfg);
    const res = UrlFetchApp.fetch(root + 'wp/v2/users/me', {
      headers: { Authorization: cfg.auth }, muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) { ui.alert('API found at ' + root + '\n\n' + wpError_(res)); return; }
    const types = UrlFetchApp.fetch(root + 'wp/v2/' + cfg.type, {
      headers: { Authorization: cfg.auth }, muteHttpExceptions: true,
    });
    ui.alert('Connected as "' + JSON.parse(res.getContentText()).name + '".\nAPI address: ' + root + '\n' +
      (types.getResponseCode() === 200
        ? 'Drafts will be created as: ' + cfg.type
        : 'But WP_POST_TYPE "' + cfg.type + '" was not recognised. ' + wpError_(types)));
  } catch (err) {
    ui.alert(String(err.message || err));
  }
}

function createWordPressDrafts() {
  const ui = SpreadsheetApp.getUi();
  let cfg;
  try { cfg = wpConfig_(); wpApiRoot_(cfg); } catch (err) { ui.alert(err.message); return; }

  // Approved rows with no draft yet, from "Cases" and every session tab
  const ready = [];
  caseSheets_().forEach(sheet => {
    ensureReady_(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const col = columnMap_(sheet);
    const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    const raw = range.getValues();            // real booleans for checkboxes
    const text = range.getDisplayValues();    // everything else as shown
    for (let r = 0; r < raw.length; r++) {
      const get = key => text[r][col[key] - 1].trim();
      if (raw[r][col.approved - 1] === true && !get('wp_post_id')) {
        ready.push({ sheet: sheet, col: col, row: r + 2, get: get });
      }
    }
  });
  if (!ready.length) {
    ui.alert('Nothing to do. Tick "Approved" on a row that has no WP post ID yet.');
    return;
  }
  const names = ready.map(x => '• ' + x.get('title') + ' (' + x.sheet.getName() + ')').join('\n');
  const answer = ui.alert('Create ' + ready.length + ' WordPress draft(s)?',
    names + '\n\nThese are created as drafts. Nothing is published.', ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;

  let made = 0;
  ready.forEach(x => {
    const sheet = x.sheet, col = x.col, row = x.row;
    const status = sheet.getRange(row, col.wp_status);
    try {
      const c = {};
      COLUMNS.forEach(k => { c[k.key] = x.get(k.key); });
      if (!c.title) throw new Error('Case title is empty');
      if (!c.blog_summary) throw new Error('Blog summary is empty');
      if (!c.reviewed_by) throw new Error('Add your initials under "Reviewed by" first');

      const res = UrlFetchApp.fetch(wpApiRoot_(cfg) + 'wp/v2/' + cfg.type, {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: cfg.auth },
        payload: JSON.stringify({ title: c.title, content: buildPostContent_(c), status: 'draft' }),
        muteHttpExceptions: true,
      });
      if (res.getResponseCode() !== 201) throw new Error(wpError_(res));
      const post = JSON.parse(res.getContentText());
      sheet.getRange(row, col.wp_post_id).setValue(String(post.id));
      sheet.getRange(row, col.wp_edit_link)
        .setValue(wpEditLink_(cfg, post.id));
      status.setValue('Draft created ' +
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
      made++;
    } catch (err) {
      status.setValue('ERROR: ' + String(err.message || err));
    }
  });
  ui.alert(made + ' of ' + ready.length + ' draft(s) created. Open the "WP edit link" to review and publish.');
}

/**
 * The body of the draft, laid out like the existing case pages: the issues
 * summary in a wide left column, a bordered Decision / Argument box on the
 * right, and the "Select one..." document dropdown underneath. Written as
 * standard WordPress blocks so the reviewer can edit every piece in the editor.
 */
function buildPostContent_(c) {
  const link = (url, label) => isHttpUrl_(url) ? '<a href="' + esc_(url) + '">' + label + '</a>' : '';
  const rows = [];

  if (c.decision_date) {
    const d = esc_(blogDate_(c.decision_date));
    rows.push('Decision: ' + (link(c.decision_url, d) || d));
  }
  if (c.argument_date) {
    const extras = [link(c.video_url, 'vid.'), link(c.transcript_url, 'tr.')].filter(String).join(' ');
    rows.push('Argument: ' + esc_(blogDate_(c.argument_date)) + (extras ? ' ' + extras : ''));
  }

  const summary = '<!-- wp:paragraph -->\n<p>' + esc_(c.blog_summary) + '</p>\n<!-- /wp:paragraph -->';
  const dropdown = buildDropdown_(c);
  if (!rows.length) return summary + '\n\n' + dropdown;

  const box = '<!-- wp:table {"hasFixedLayout":false} -->\n' +
    '<figure class="wp-block-table"><table><tbody>' +
    rows.map(r => '<tr><td>' + r + '</td></tr>').join('') +
    '</tbody></table></figure>\n<!-- /wp:table -->';

  return '<!-- wp:columns -->\n<div class="wp-block-columns">' +
    '<!-- wp:column {"width":"66.66%"} -->\n<div class="wp-block-column" style="flex-basis:66.66%">' +
    summary + '</div>\n<!-- /wp:column -->\n\n' +
    '<!-- wp:column {"width":"33.33%"} -->\n<div class="wp-block-column" style="flex-basis:33.33%">' +
    box + '</div>\n<!-- /wp:column --></div>\n<!-- /wp:columns -->\n\n' + dropdown;
}

/** The [dropdown] and [dropdown-content] shortcodes from the page template. */
function buildDropdown_(c) {
  const options = DROPDOWN_SECTIONS.map(s =>
    '[dropdown-option value="' + s.value + '"]' + s.label + '[/dropdown-option]').join('');
  const contents = DROPDOWN_SECTIONS.map(s =>
    '[dropdown-content value="' + s.value + '"]' + documentLines_(c[s.key], s.label) + '[/dropdown-content]').join('');
  return '<!-- wp:shortcode -->\n' +
    '[dropdown class="Court of Appeals Briefs"][dropdown-option]Select one...[/dropdown-option]' +
    options + '[/dropdown]\n' + contents +
    '\n<!-- /wp:shortcode -->';
}

/** Turns "Label | URL" lines into links. A line with no URL stays plain text.
 *  A malformed line stops the draft, so a broken link can't slip through. */
function documentLines_(cell, section) {
  return String(cell || '').split(/\r?\n/).map(l => l.trim()).filter(String).map(line => {
    const bar = line.lastIndexOf('|');
    const label = (bar === -1 ? line : line.slice(0, bar)).trim();
    const url = bar === -1 ? '' : line.slice(bar + 1).trim();
    if (!label || isHttpUrl_(label) || (bar !== -1 && !isHttpUrl_(url))) {
      throw new Error(section + ': use "Label | URL" on this line: ' + line);
    }
    const text = esc_(label).replace(/\[/g, '&#91;').replace(/\]/g, '&#93;');
    return url ? '<a href="' + esc_(url) + '">' + text + '</a>' : text;
  }).join('\n');
}

/** "05/26/2026" -> "2026.05.26", the date style used on the blog. */
function blogDate_(s) {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(s).trim());
  return m ? m[3] + '.' + ('0' + m[1]).slice(-2) + '.' + ('0' + m[2]).slice(-2) : String(s);
}

function isHttpUrl_(u) { return /^https?:\/\/\S+$/i.test(u || ''); }

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
