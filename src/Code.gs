/**
 * Twenty Eagle case pipeline — Google Sheet side.
 *
 *   1. doPost()                 receives rows from sheet_sync.py (scraper -> Sheet)
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
 *   WP_POST_TYPE      "pages" if each case is a WordPress Page; leave out for Posts
 */

const SHEET_NAME = 'Cases';

// group: scraper = overwritten on every sync
//        human   = never overwritten by sync (blank cells may be pre-filled, see SEEDED)
//        wp      = written by createWordPressDrafts()
const COLUMNS = [
  { key: 'case_key',          header: 'Case key',              group: 'scraper', width: 160 },
  { key: 'title',             header: 'Case title',            group: 'scraper', width: 240 },
  { key: 'decision_date',     header: 'Decision date',         group: 'scraper', width: 100 },
  { key: 'decision_type',     header: 'Decision type',         group: 'scraper', width: 100 },
  { key: 'argument_date',     header: 'Argument date',         group: 'scraper', width: 100 },
  { key: 'argument_no',       header: 'Argument no.',          group: 'scraper', width: 90 },
  { key: 'apl_number',        header: 'APL no.',               group: 'scraper', width: 120 },
  { key: 'category',          header: 'Category',              group: 'scraper', width: 180 },
  { key: 'issues_courtpass',  header: 'Issues (CourtPASS)',    group: 'scraper', width: 380 },
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
  { key: 'request_jl',        header: 'Request JL?',           group: 'human', width: 90,  checkbox: true },
  { key: 'notes',             header: 'Reviewer notes',        group: 'human', width: 240 },
  { key: 'reviewed_by',       header: 'Reviewed by',           group: 'human', width: 100 },
  { key: 'approved',          header: 'Approved',              group: 'human', width: 80,  checkbox: true },

  { key: 'wp_post_id',        header: 'WP post ID',            group: 'wp', width: 90 },
  { key: 'wp_edit_link',      header: 'WP edit link',          group: 'wp', width: 240 },
  { key: 'wp_status',         header: 'WP status',             group: 'wp', width: 220 },
];

// Human columns the sync may pre-fill, but only while the cell is still blank
const SEEDED = ['blog_summary', 'decision_url', 'video_url', 'transcript_url'];

const GROUP_COLORS = { scraper: '#d9e2f3', human: '#fff2cc', wp: '#e2efda' };

// ───────────────────────── Menu & setup ─────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Twenty Eagle')
    .addItem('Create WordPress drafts for approved rows', 'createWordPressDrafts')
    .addSeparator()
    .addItem('Test WordPress connection', 'testWordPressConnection')
    .addItem('Set up sheet', 'setupSheet')
    .addToUi();
}

/** Creates the "Cases" tab, or adds any columns a newer version of this script
 *  introduced. Safe to re-run: existing data and edits are left alone. */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

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
  if (!sheet || sheet.getLastRow() < 1) return setupSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  return COLUMNS.every(c => headers.indexOf(c.header) !== -1) ? sheet : setupSheet();
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
  return ContentService.createTextOutput('Twenty Eagle sync endpoint is live (v2: link lookup)');
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
    const sheet = getSheet_();
    const col = columnMap_(sheet);
    const scraperCols = COLUMNS.filter(c => c.group === 'scraper');

    const rowByKey = {};
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, col.case_key, lastRow - 1, 1).getDisplayValues()
        .forEach((r, i) => { if (r[0]) rowByKey[r[0]] = i + 2; });
    }

    const added = [], updated = [];
    cases.forEach(c => {
      if (!c.case_key) return;
      let row = rowByKey[c.case_key];
      const isNew = !row;
      if (isNew) {
        row = sheet.getLastRow() + 1;
        rowByKey[c.case_key] = row;
      }
      scraperCols.forEach(sc => sheet.getRange(row, col[sc.key]).setValue(safeText_(c[sc.key])));

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
    return { ok: true, added: added, updated: updated };
  } finally {
    lock.releaseLock();
  }
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
    type: p.getProperty('WP_POST_TYPE') || 'posts',
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

  const sheet = getSheet_();
  const col = columnMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('No cases in the sheet yet.'); return; }

  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const raw = range.getValues();            // real booleans for checkboxes
  const text = range.getDisplayValues();    // everything else as shown
  const get = (r, key) => text[r][col[key] - 1].trim();

  const ready = [];
  for (let r = 0; r < raw.length; r++) {
    if (raw[r][col.approved - 1] === true && !get(r, 'wp_post_id')) ready.push(r);
  }
  if (!ready.length) {
    ui.alert('Nothing to do. Tick "Approved" on a row that has no WP post ID yet.');
    return;
  }
  const names = ready.map(r => '• ' + get(r, 'title')).join('\n');
  const answer = ui.alert('Create ' + ready.length + ' WordPress draft(s)?',
    names + '\n\nThese are created as drafts. Nothing is published.', ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;

  let made = 0;
  ready.forEach(r => {
    const row = r + 2;
    const status = sheet.getRange(row, col.wp_status);
    try {
      const c = {};
      COLUMNS.forEach(x => { c[x.key] = get(r, x.key); });
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
 * summary in a wide left column and a bordered Decision / Argument box on the
 * right. Written as standard WordPress blocks so the reviewer can edit every
 * piece in the block editor. (The "Select one…" dropdown on the live pages is
 * not generated here.)
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
  if (!rows.length) return summary;

  const box = '<!-- wp:table {"hasFixedLayout":false} -->\n' +
    '<figure class="wp-block-table"><table><tbody>' +
    rows.map(r => '<tr><td>' + r + '</td></tr>').join('') +
    '</tbody></table></figure>\n<!-- /wp:table -->';

  return '<!-- wp:columns -->\n<div class="wp-block-columns">' +
    '<!-- wp:column {"width":"66.66%"} -->\n<div class="wp-block-column" style="flex-basis:66.66%">' +
    summary + '</div>\n<!-- /wp:column -->\n\n' +
    '<!-- wp:column {"width":"33.33%"} -->\n<div class="wp-block-column" style="flex-basis:33.33%">' +
    box + '</div>\n<!-- /wp:column --></div>\n<!-- /wp:columns -->';
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
