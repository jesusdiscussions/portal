/**
 * Blog Platform Engine — V2 (Phase 1: schema, User IDs, migration).
 * Spreadsheet = database; GitHub Pages = public UI; Worker = Phase 2+.
 */

var SHEET_POST_LOG = 'PostLog';
var SHEET_PAGE_CONTENT = 'PageContent';
var SHEET_EMAIL_QUEUE = 'EmailQueue';
var SHEET_ACTION_TOKENS = 'ActionTokens';
var SHEET_COMMENTS = 'CommentsLog';
/** Never delete or overwrite during setup/migration. */
var SHEET_NOTES_WHITELIST = 'Notes';

var USER_ID_START = 1000;
var PROP_NEXT_USER_ID = 'NEXT_USER_ID';

var HEADER_POST_ID = 'Post ID';
var HEADER_STATUS = 'STATUS';
var HEADER_SENT = 'Sent Timestamp';
var HEADER_DAYS_OPEN = 'Comments Duration Open';
var HEADER_FIRST = 'First Name';
var HEADER_LAST = 'Last Name';
var HEADER_EMAIL = 'Email Address';
var HEADER_SUBJECT = 'Email Subject';
var HEADER_MESSAGE = 'Email Message Body';
var HEADER_FILE_IDS = 'Google Drive File IDs';

var MASTER_MARKER = 'post id';

var STATUS_SUBSCRIBED = 'SUBSCRIBED';
var STATUS_CONFIRMED = 'CONFIRMED';
var STATUS_UNSUBSCRIBED = 'UNSUBSCRIBED';

var ATTACH_ERROR_BG = '#ffcccc';

var GITHUB_PAGES_BASE_URL_DEFAULT = 'https://jesusdiscussions.github.io/portal';
var WEB_APP_URL_DEFAULT =
  'https://script.google.com/macros/s/AKfycbxa3oj2BTcEQladm9llk7yyMYU2vcy7GzQ_dqbnPVGAUiSY9Imf_VZqjIJAHSDD-Q_wbw/exec';

var POST_LOG_HEADERS = [
  HEADER_POST_ID, HEADER_STATUS, HEADER_SENT, HEADER_DAYS_OPEN,
  HEADER_FIRST, HEADER_LAST, HEADER_EMAIL, HEADER_SUBJECT,
  HEADER_MESSAGE, HEADER_FILE_IDS
];

var COMMENTS_LOG_HEADERS = [
  'Comment ID', 'Parent ID', HEADER_POST_ID, 'Timestamp',
  'User Email', 'User Display Name', 'Comment Text'
];

var PAGE_CONTENT_HEADERS = [
  'Post ID', 'Live', 'Page Headline', 'Page Body HTML', 'Google Drive File IDs',
  'Comments Duration Open', 'Discussion Start At', 'Force Closed', 'Last Updated', 'Publish Version'
];

var EMAIL_QUEUE_HEADERS = [
  'User ID', 'Post ID', 'Email Address', 'First Name', 'Last Name', 'STATUS',
  'Email Subject', 'Email Body', 'Send?', 'Sent Timestamp'
];

var ACTION_TOKEN_HEADERS = [
  'Token', 'Post ID', 'Email Address', 'Action', 'Expires At', 'Used'
];

var PC = {
  POST: 0, LIVE: 1, HEADLINE: 2, BODY: 3, FILES: 4, DAYS: 5, START: 6,
  FORCE: 7, UPDATED: 8, VERSION: 9
};

var EQ = {
  UID: 0, POST: 1, EMAIL: 2, FIRST: 3, LAST: 4, STATUS: 5,
  SUBJECT: 6, BODY: 7, SEND: 8, SENT: 9
};

var CMT = { ID: 0, PARENT: 1, POST: 2, TIME: 3, EMAIL: 4, NAME: 5, TEXT: 6 };

// ——— Menu ———

function onOpen() {
  rememberSpreadsheetId_();
  SpreadsheetApp.getUi()
    .createMenu('Blog Platform Engine')
    .addItem('Setup V2 sheets (one-shot)', 'setupV2Sheets')
    .addItem('Migrate PostLog → V2…', 'migratePostLogToV2')
    .addItem('Assign missing User IDs', 'assignMissingUserIds')
    .addSeparator()
    .addItem('Set GitHub Pages base URL…', 'promptGitHubPagesUrl')
    .addItem('Test API (open in browser)…', 'openApiTestLink')
    .addSeparator()
    .addItem('Run Post Send & Sync (legacy)', 'runPostEngine')
    .addToUi();
}

/** Test API: shows URL in a native alert (modal HTML is blocked in Sheets for this permission set). */
function openApiTestLink() {
  var url = getWebAppUrl() + '?api=json&cmd=ping';
  var ui = SpreadsheetApp.getUi();
  /** showModalDialog often fails with "permissions not sufficient" in Sheets; alert always works. */
  ui.alert(
    'API access test',
    'Open this URL in a new browser tab.\n\n' +
      'You should see JSON starting with {"ok":true ...\n\n' +
      'If Google asks you to sign in, redeploy the web app with access Anyone (anonymous).\n\n' +
      url,
    ui.ButtonSet.OK
  );
}

function promptGitHubPagesUrl() {
  var ui = SpreadsheetApp.getUi();
  var current = PropertiesService.getScriptProperties().getProperty('GITHUB_PAGES_BASE_URL') ||
    GITHUB_PAGES_BASE_URL_DEFAULT;
  var result = ui.prompt(
    'GitHub Pages base URL',
    'No trailing slash.\nCurrent: ' + current,
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;
  var url = result.getResponseText().trim().replace(/\/$/, '');
  if (!url) {
    ui.alert('URL was not saved (empty input).');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('GITHUB_PAGES_BASE_URL', url);
  ui.alert('Saved:\n' + url);
}

/** @deprecated Use setupV2Sheets */
function setupSheets() {
  setupV2Sheets();
}

function setupV2Sheets() {
  rememberSpreadsheetId_();
  var ss = openSpreadsheet_();
  if (!ss) {
    SpreadsheetApp.getUi().alert('Open the spreadsheet bound to this project first.');
    return;
  }
  var results = ensureV2Schema_(ss, {});
  initUserIdSequenceFromSheet_();
  SpreadsheetApp.getUi().alert(
    'Blog Platform Engine V2 sheets ready.\n\n' +
      results.join('\n') +
      '\n\nNotes tab: left unchanged.\n' +
      'Next: run "Migrate PostLog → V2" if you have data in PostLog.\n' +
      'User IDs start at ' + USER_ID_START + ' (menu: Assign missing User IDs).'
  );
}

// ——— V2 schema, User IDs, migration ———

function isWhitelistedSheet_(name) {
  return String(name || '').trim() === SHEET_NOTES_WHITELIST;
}

function ensureV2Schema_(ss, opts) {
  opts = opts || {};
  var clearTabs = opts.clearTabs || {};
  var lines = [];
  var tabs = [
    { name: SHEET_PAGE_CONTENT, headers: PAGE_CONTENT_HEADERS },
    { name: SHEET_EMAIL_QUEUE, headers: EMAIL_QUEUE_HEADERS },
    { name: SHEET_ACTION_TOKENS, headers: ACTION_TOKEN_HEADERS },
    { name: SHEET_COMMENTS, headers: COMMENTS_LOG_HEADERS }
  ];
  for (var i = 0; i < tabs.length; i++) {
    var t = tabs[i];
    var r = ensureSheetWithHeaders(ss, t.name, t.headers, !!clearTabs[t.name]);
    lines.push('• ' + t.name + (r.created ? ' (created)' : r.headersReset ? ' (headers reset)' : ''));
  }
  if (ss.getSheetByName(SHEET_NOTES_WHITELIST)) {
    lines.push('• ' + SHEET_NOTES_WHITELIST + ' (preserved)');
  }
  return lines;
}

function getPageContentSheet() {
  var ss = openSpreadsheet_();
  if (!ss) return null;
  return ss.getSheetByName(SHEET_PAGE_CONTENT);
}

function getEmailQueueSheet() {
  var ss = openSpreadsheet_();
  if (!ss) return null;
  return ss.getSheetByName(SHEET_EMAIL_QUEUE);
}

function getActionTokensSheet() {
  var ss = openSpreadsheet_();
  if (!ss) return null;
  return ss.getSheetByName(SHEET_ACTION_TOKENS);
}

function initUserIdSequenceFromSheet_() {
  var props = PropertiesService.getScriptProperties();
  var sheet = getEmailQueueSheet();
  var max = USER_ID_START - 1;
  if (sheet && sheet.getLastRow() >= 2) {
    var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var n = parseInt(String(vals[i][0] || '').replace(/\D/g, ''), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  var next = Math.max(USER_ID_START, max + 1);
  props.setProperty(PROP_NEXT_USER_ID, String(next));
  return next;
}

function allocateUserIds_(count) {
  var props = PropertiesService.getScriptProperties();
  var next = parseInt(props.getProperty(PROP_NEXT_USER_ID) || '', 10);
  if (isNaN(next) || next < USER_ID_START) next = initUserIdSequenceFromSheet_();
  var out = [];
  for (var i = 0; i < count; i++) {
    out.push(String(next));
    next++;
  }
  props.setProperty(PROP_NEXT_USER_ID, String(next));
  return out;
}

function assignMissingUserIds() {
  rememberSpreadsheetId_();
  var sheet = getEmailQueueSheet();
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Run "Setup V2 sheets" first.');
    return;
  }
  initUserIdSequenceFromSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('EmailQueue has no data rows.');
    return;
  }
  var uids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var need = 0;
  for (var i = 0; i < uids.length; i++) {
    if (!String(uids[i][0] || '').trim()) need++;
  }
  if (!need) {
    SpreadsheetApp.getUi().alert('Every EmailQueue row already has a User ID.');
    return;
  }
  var newIds = allocateUserIds_(need);
  var idx = 0;
  for (var r = 0; r < uids.length; r++) {
    if (!String(uids[r][0] || '').trim()) {
      sheet.getRange(r + 2, 1).setValue(newIds[idx++]);
    }
  }
  SpreadsheetApp.getUi().alert('Assigned ' + need + ' User ID(s). Next available: ' + propsNextUserId_());
}

function propsNextUserId_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_NEXT_USER_ID) || String(USER_ID_START);
}

function migratePostLogToV2() {
  rememberSpreadsheetId_();
  var ui = SpreadsheetApp.getUi();
  var ss = openSpreadsheet_();
  if (!ss) {
    ui.alert('Open the spreadsheet bound to this project first.');
    return;
  }
  var legacy = getPostLogSheet();
  if (!legacy || legacy.getLastRow() < 2) {
    ui.alert('No PostLog (or CampaignLog) data found to migrate. V2 sheets will still be created.');
    ensureV2Schema_(ss, {});
    return;
  }

  var confirm = ui.alert(
    'Migrate to V2',
    'This will:\n' +
      '• Copy master rows → PageContent\n' +
      '• Copy subscriber rows → EmailQueue (new User IDs from ' + USER_ID_START + ')\n' +
      '• Delete the PostLog / CampaignLog tab\n' +
      '• Leave the Notes tab untouched\n\n' +
      'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var lastRow = legacy.getLastRow();
  var lastCol = Math.max(legacy.getLastColumn(), POST_LOG_HEADERS.length);
  var headers = legacy.getRange(1, 1, 1, lastCol).getValues()[0];
  var colMap = buildColumnMap(headers);
  var clusters = buildPostClusters(legacy, headers, lastRow, lastCol);

  ensureV2Schema_(ss, {
    clearTabs: {}
  });
  ensureSheetWithHeaders(ss, SHEET_PAGE_CONTENT, PAGE_CONTENT_HEADERS, true);
  ensureSheetWithHeaders(ss, SHEET_EMAIL_QUEUE, EMAIL_QUEUE_HEADERS, true);
  ensureSheetWithHeaders(ss, SHEET_ACTION_TOKENS, ACTION_TOKEN_HEADERS, false);
  ensureSheetWithHeaders(ss, SHEET_COMMENTS, COMMENTS_LOG_HEADERS, false);
  initUserIdSequenceFromSheet_();

  var pageSheet = getPageContentSheet();
  var emailSheet = getEmailQueueSheet();
  var pageByWebId = {};
  var emailRows = [];
  var skipped = [];
  var postCol = colIndex(colMap, HEADER_POST_ID, 0);
  var firstCol = colIndex(colMap, HEADER_FIRST, 4);
  var lastColIdx = colIndex(colMap, HEADER_LAST, 5);
  var subjectCol = colIndex(colMap, HEADER_SUBJECT, 7);
  var messageCol = colIndex(colMap, HEADER_MESSAGE, 8);
  var fileCol = colIndex(colMap, HEADER_FILE_IDS, 9);
  var daysCol = colIndex(colMap, HEADER_DAYS_OPEN, 3);
  var sentCol = colIndex(colMap, HEADER_SENT, 2);
  var statusCol = colIndex(colMap, HEADER_STATUS, 1);
  var emailCol = colIndex(colMap, HEADER_EMAIL, 6);

  for (var c = 0; c < clusters.length; c++) {
    var cluster = clusters[c];
    var master = cluster.master;
    var rowNum = cluster.masterSheetRow;
    var first = String(master[firstCol] || '').trim();
    if (first.toLowerCase() !== MASTER_MARKER) {
      skipped.push('Row ' + rowNum + ' (Post ID "' + cluster.postId + '"): missing master blueprint');
      continue;
    }
    var routing = getRouting(master, colMap);
    var templateSubject = String(master[subjectCol] || '').trim();
    var templateBody = String(master[messageCol] || '');
    var templateFiles = String(master[fileCol] || '').trim();
    var daysOpen = master[daysCol];
    var now = new Date();

    pageByWebId[routing.webPostId] = [
      routing.webPostId,
      'Y',
      templateSubject,
      templateBody,
      templateFiles,
      daysOpen === '' || daysOpen == null ? '' : daysOpen,
      '',
      'N',
      now,
      1
    ];

    for (var i = 0; i < cluster.rows.length; i++) {
      var row = cluster.rows[i].row;
      if (!isSubscriberRow(row, colMap)) continue;
      emailRows.push([
        '',
        String(row[postCol] || '').trim() || cluster.postId,
        String(row[emailCol] || '').trim(),
        String(row[firstCol] || '').trim(),
        String(row[lastColIdx] || '').trim(),
        String(row[statusCol] || '').trim() || STATUS_SUBSCRIBED,
        templateSubject,
        templateBody,
        'N',
        formatTs(row[sentCol])
      ]);
    }
  }

  var pageKeys = Object.keys(pageByWebId);
  var pageRows = [];
  for (var pk = 0; pk < pageKeys.length; pk++) pageRows.push(pageByWebId[pageKeys[pk]]);
  if (pageRows.length) {
    pageSheet.getRange(2, 1, pageRows.length + 1, PAGE_CONTENT_HEADERS.length).setValues(pageRows);
  }
  if (emailRows.length) {
    var ids = allocateUserIds_(emailRows.length);
    for (var e = 0; e < emailRows.length; e++) {
      emailRows[e][EQ.UID] = ids[e];
    }
    emailSheet.getRange(2, 1, emailRows.length + 1, EMAIL_QUEUE_HEADERS.length).setValues(emailRows);
  }

  var legacyName = legacy.getName();
  if (!isWhitelistedSheet_(legacyName)) {
    ss.deleteSheet(legacy);
  }

  var summary =
    'V2 migration complete.\n\n' +
    '• PageContent rows: ' + pageRows.length + '\n' +
    '• EmailQueue rows: ' + emailRows.length + ' (User IDs assigned)\n' +
    '• Removed tab: ' + legacyName + '\n' +
    '• Notes tab: preserved\n' +
    '• Next User ID: ' + propsNextUserId_();
  if (skipped.length) {
    summary += '\n\nSkipped clusters:\n' + skipped.join('\n');
  }
  ui.alert(summary);
}

// ——— Post engine ———

function runPostEngine() {
  rememberSpreadsheetId_();
  var sheet = getPostLogSheet();
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'Legacy PostLog is not present (V2 migration may have removed it).\n\n' +
        'V2 batch email send ships in Phase 2. Use EmailQueue + PageContent for now.'
    );
    return;
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var lastCol = Math.max(sheet.getLastColumn(), POST_LOG_HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colMap = buildColumnMap(headers);
  var clusters = buildPostClusters(sheet, headers, lastRow, lastCol);
  if (!clusters.length) return;

  for (var v = 0; v < clusters.length; v++) {
    var err = validateMasterBlueprint(clusters[v], colMap);
    if (err) {
      SpreadsheetApp.getUi().alert(err);
      return;
    }
  }

  var driveCache = {};
  for (var c = 0; c < clusters.length; c++) {
    processPostCluster(clusters[c], sheet, headers, colMap, driveCache);
  }
}

function buildColumnMap(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var name = String(headers[i] || '').trim();
    if (name) map[name] = i;
  }
  return map;
}

function colIndex(colMap, name, fallback) {
  return colMap[name] !== undefined ? colMap[name] : fallback;
}

function buildPostClusters(sheet, headers, lastRow, lastCol) {
  var postCol = colIndex(buildColumnMap(headers), HEADER_POST_ID, 0);
  var rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var clusterMap = {};
  var order = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!rowHasContent(row)) continue;
    var key = String(row[postCol] || '').trim();
    if (!key) continue;

    if (!clusterMap[key]) {
      clusterMap[key] = {
        postId: key,
        master: row,
        masterSheetRow: i + 2,
        rows: []
      };
      order.push(key);
    } else {
      clusterMap[key].rows.push({ row: row, sheetRow: i + 2 });
    }
  }

  var out = [];
  for (var k = 0; k < order.length; k++) out.push(clusterMap[order[k]]);
  return out;
}

function rowHasContent(row) {
  for (var i = 0; i < row.length; i++) {
    if (String(row[i] || '').trim() !== '') return true;
  }
  return false;
}

function validateMasterBlueprint(cluster, colMap) {
  var postCol = colIndex(colMap, HEADER_POST_ID, 0);
  var firstCol = colIndex(colMap, HEADER_FIRST, 4);
  var lastCol = colIndex(colMap, HEADER_LAST, 5);
  var master = cluster.master;
  var rowNum = cluster.masterSheetRow;
  var postId = String(master[postCol] || '').trim();
  var first = String(master[firstCol] || '').trim();
  var last = String(master[lastCol] || '').trim();
  var label = postId || '(blank Post ID)';

  if (!first) {
    return 'HALTED — Row ' + rowNum + ' (Post ID cluster "' + label + '"):\n\n' +
      '• First Name is blank.\n• It must be "Post ID" (case-insensitive).';
  }
  if (first.toLowerCase() !== MASTER_MARKER) {
    return 'HALTED — Row ' + rowNum + ' (Post ID cluster "' + label + '"):\n\n' +
      '• First Name is "' + first + '".\n• It must be "Post ID" (case-insensitive).';
  }
  if (!postId) {
    return 'HALTED — Row ' + rowNum + ':\n\n• Post ID (column A) is blank.';
  }
  if (!last) {
    return 'HALTED — Row ' + rowNum + ' (Post ID cluster "' + label + '"):\n\n' +
      '• Last Name is blank.\n• Set web target ID (Launch: same as Post ID; Modify: target Post ID).';
  }
  return null;
}

function getRouting(master, colMap) {
  var postCol = colIndex(colMap, HEADER_POST_ID, 0);
  var lastCol = colIndex(colMap, HEADER_LAST, 5);
  var tracking = String(master[postCol] || '').trim();
  var webTarget = String(master[lastCol] || '').trim();
  var launch = tracking === webTarget;
  return {
    trackingPostId: tracking,
    webPostId: launch ? tracking : webTarget,
    isLaunch: launch
  };
}

function processPostCluster(cluster, sheet, headers, colMap, driveCache) {
  var routing = getRouting(cluster.master, colMap);
  var master = cluster.master;
  var subjectCol = colIndex(colMap, HEADER_SUBJECT, 7);
  var messageCol = colIndex(colMap, HEADER_MESSAGE, 8);
  var fileCol = colIndex(colMap, HEADER_FILE_IDS, 9);
  var daysCol = colIndex(colMap, HEADER_DAYS_OPEN, 3);
  var sentCol = colIndex(colMap, HEADER_SENT, 2);
  var statusCol = colIndex(colMap, HEADER_STATUS, 1);
  var emailCol = colIndex(colMap, HEADER_EMAIL, 6);
  var firstCol = colIndex(colMap, HEADER_FIRST, 4);
  var lastCol = colIndex(colMap, HEADER_LAST, 5);

  var templateSubject = String(master[subjectCol] || '').trim();
  var templateBody = String(master[messageCol] || '');
  var templateFiles = master[fileCol];
  var daysOpen = parseFloat(master[daysCol]) || 0;
  var launchedAt = new Date();

  syncWebPostToGitHub({
    postId: routing.webPostId,
    trackingPostId: routing.trackingPostId,
    subject: templateSubject,
    messageHtml: templateBody,
    fileIdsRaw: String(templateFiles || ''),
    daysOpen: daysOpen,
    launchedAt: launchedAt.toISOString(),
    mode: routing.isLaunch ? 'launch' : 'modify'
  }, routing.isLaunch);

  for (var i = 0; i < cluster.rows.length; i++) {
    var entry = cluster.rows[i];
    var row = entry.row;
    var sheetRow = entry.sheetRow;

    if (!isSubscriberRow(row, colMap)) continue;

    var status = String(row[statusCol] || '').trim().toUpperCase();
    if (status !== STATUS_SUBSCRIBED && status !== STATUS_CONFIRMED) continue;
    if (!isBlankTimestamp(row[sentCol])) continue;

    var recipient = String(row[emailCol] || '').trim();
    if (!recipient) continue;

    var rowFiles = String(row[fileCol] || '').trim() ? row[fileCol] : templateFiles;
    var fileCell = sheet.getRange(sheetRow, fileCol + 1);
    var audit = auditDriveFiles(rowFiles, driveCache);

    if (!audit.valid) {
      fileCell.setBackground(ATTACH_ERROR_BG);
      fileCell.setNote(audit.errorNote);
      continue;
    }
    fileCell.setBackground(null);
    fileCell.clearNote();

    var mergedSubject = applyMergeTags(templateSubject, headers, row);
    var mergedBody = applyMergeTags(templateBody, headers, row);
    var fname = String(row[firstCol] || '').trim();
    var lname = String(row[lastCol] || '').trim();

    try {
      sendPostEmail(
        mergedSubject,
        mergedBody,
        routing.webPostId,
        recipient,
        audit.files,
        recipient,
        fname,
        lname,
        daysOpen,
        launchedAt
      );
      sheet.getRange(sheetRow, sentCol + 1).setValue(new Date());
    } catch (e) {
      Logger.log('Send failed row ' + sheetRow + ': ' + e.message);
    }
  }
}

function isSubscriberRow(row, colMap) {
  var firstCol = colIndex(colMap, HEADER_FIRST, 4);
  var emailCol = colIndex(colMap, HEADER_EMAIL, 6);
  var first = String(row[firstCol] || '').trim().toLowerCase();
  if (first === MASTER_MARKER) return false;
  return String(row[emailCol] || '').trim() !== '';
}

function applyMergeTags(template, headers, row) {
  var out = String(template || '');
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').trim();
    if (!h) continue;
    var val = row[i];
    if (val instanceof Date) val = val.toString();
    if (val === null || val === undefined) val = '';
    out = out.split('{{' + h + '}}').join(String(val));
  }
  return out;
}

function isBlankTimestamp(v) {
  return v === '' || v === null || v === undefined;
}

function auditDriveFiles(raw, cache) {
  var ids = parseIdList(raw);
  if (!ids.length) return { valid: true, files: [], errorNote: '' };
  var files = [];
  for (var i = 0; i < ids.length; i++) {
    var id = ids[i];
    if (!cache[id]) cache[id] = verifyDriveFile(id);
    if (!cache[id].ok) {
      return { valid: false, files: [], errorNote: 'Drive ID "' + id + '": ' + cache[id].error };
    }
    files.push({ id: id, name: cache[id].name, url: cache[id].url });
  }
  return { valid: true, files: files, errorNote: '' };
}

function parseIdList(raw) {
  if (!raw) return [];
  return String(raw).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
}

function verifyDriveFile(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var access = file.getSharingAccess();
    var ok = access === DriveApp.Access.ANYONE || access === DriveApp.Access.ANYONE_WITH_LINK;
    if (!ok) {
      return { ok: false, error: 'File must be public or anyone-with-link.' };
    }
    return {
      ok: true,
      name: file.getName(),
      url: 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(fileId)
    };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : 'Invalid file ID.' };
  }
}

function sendPostEmail(subject, bodyHtml, webPostId, to, files, email, fname, lname, daysOpen, sentAt) {
  var subj = String(subject || '').trim() || 'Blog Post';
  var deadlineIso = getDeadlineIso(sentAt, daysOpen);
  var html = bodyHtml + buildEmailFooter(webPostId, email, fname, lname, deadlineIso);
  var opts = { htmlBody: html };
  var blobs = [];
  for (var i = 0; i < files.length; i++) {
    blobs.push(DriveApp.getFileById(files[i].id).getBlob());
  }
  if (blobs.length) opts.attachments = blobs;
  MailApp.sendEmail(to, subj, stripHtml(html), opts);
}

// ——— GitHub Pages links (email + portal) ———

function getGitHubPagesBaseUrl() {
  var base = PropertiesService.getScriptProperties().getProperty('GITHUB_PAGES_BASE_URL') ||
    GITHUB_PAGES_BASE_URL_DEFAULT;
  return String(base).trim().replace(/\/$/, '');
}

function getWebAppUrl() {
  var prop = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL');
  if (prop) return prop.trim();
  try {
    var url = ScriptApp.getService().getUrl();
    if (url) return url;
  } catch (e) { /* not deployed */ }
  return WEB_APP_URL_DEFAULT;
}

function resolvePostId(params) {
  if (!params) return '';
  var post = String(params.post || '').trim();
  if (post) return post;
  return String(params.camp || '').trim();
}

function buildGitHubPageUrl(page, queryMap, hash) {
  var parts = [];
  var keys = Object.keys(queryMap || {});
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var v = queryMap[k];
    if (v !== undefined && v !== null && String(v) !== '') {
      parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(String(v)));
    }
  }
  var url = getGitHubPagesBaseUrl() + '/' + page;
  if (parts.length) url += '?' + parts.join('&');
  if (hash) url += hash;
  return url;
}

/** Email links stay short: camp + email (+ optional names). Deadline comes from posts/*.json or API. */
function buildEmailFooter(webPostId, email, fname, lname, deadlineIso) {
  var q = { camp: webPostId, email: email };
  if (fname) q.fname = fname;
  if (lname) q.lname = lname;
  var confirm = buildGitHubPageUrl('index.html', Object.assign({ action: 'confirm' }, q));
  var unsub = buildGitHubPageUrl('index.html', Object.assign({ action: 'unsub' }, q));
  var comments = buildGitHubPageUrl('CommentsPage.html', q, '#join-comments');
  return '<hr style="margin-top:24px;border:none;border-top:1px solid #ddd;">' +
    '<p style="font-size:13px;color:#555;">' +
    '<a href="' + confirm + '">Confirm subscription</a> &nbsp;|&nbsp; ' +
    '<a href="' + unsub + '">Unsubscribe</a> &nbsp;|&nbsp; ' +
    '<a href="' + comments + '">Join Comments</a></p>';
}

function getDeadlineIso(sentAt, daysOpen) {
  var days = parseFloat(daysOpen);
  if (!sentAt || isNaN(days) || days <= 0) return '';
  return new Date(sentAt.getTime() + days * 86400000).toISOString();
}

function isExpired(deadlineIso, sentAt, daysOpen) {
  if (deadlineIso) {
    var d = new Date(deadlineIso);
    if (!isNaN(d.getTime()) && Date.now() >= d.getTime()) return true;
  }
  var days = parseFloat(daysOpen);
  if (!sentAt || isNaN(days) || days <= 0) return false;
  return Date.now() >= sentAt.getTime() + days * 86400000;
}

function stripHtml(html) {
  return String(html).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

// ——— GitHub repo sync (posts/*.json) ———

function syncWebPostToGitHub(payload, isLaunch) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('GITHUB_TOKEN');
  var repo = props.getProperty('GITHUB_REPO');
  if (!token || !repo) {
    Logger.log('GitHub sync skipped (set GITHUB_TOKEN + GITHUB_REPO). Post ' + payload.postId);
    return;
  }

  var attachments = resolveAttachmentList(payload.fileIdsRaw);
  var body = {
    postId: payload.postId,
    trackingPostId: payload.trackingPostId,
    subject: payload.subject,
    messageHtml: payload.messageHtml,
    fileIdsRaw: payload.fileIdsRaw,
    daysOpen: payload.daysOpen,
    launchedAt: payload.launchedAt,
    mode: payload.mode,
    attachments: attachments,
    updatedAt: new Date().toISOString()
  };

  var path = 'posts/' + payload.postId + '.json';
  var content = Utilities.base64Encode(JSON.stringify(body, null, 2), Utilities.Charset.UTF_8);
  var apiUrl = 'https://api.github.com/repos/' + repo + '/contents/' + path;
  var sha = null;

  if (!isLaunch) {
    var getResp = UrlFetchApp.fetch(apiUrl, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
      muteHttpExceptions: true
    });
    if (getResp.getResponseCode() === 200) {
      sha = JSON.parse(getResp.getContentText()).sha;
    }
  }

  var putBody = {
    message: (isLaunch ? 'Launch ' : 'Modify ') + payload.postId,
    content: content,
    branch: props.getProperty('GITHUB_BRANCH') || 'main'
  };
  if (sha) putBody.sha = sha;

  var resp = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/vnd.github+json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(putBody),
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
    Logger.log('GitHub sync failed: ' + resp.getContentText());
  }
}

// ——— Web app API (text/plain POST, JSON GET reads) ———

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (String(params.api || '').toLowerCase() === 'json') {
    var merged = mergeParams(params, {});
    var result = handleApiRequest(merged);
    var json = serializeApiResult_(result);

    /** Iframe bridge for mail clients (e.g. Yahoo) that block dynamic script injection. */
    if (String(params.embed || '').trim() === '1') {
      var b64 = Utilities.base64Encode(json, Utilities.Charset.UTF_8);
      var rid = String(params.rid || '').replace(/[^a-zA-Z0-9_-]/g, '');
      var html =
        '<!DOCTYPE html><html><head><meta charset="utf-8"><base target="_self">' +
        '<script>(function(){' +
        'var rid=' + JSON.stringify(rid) + ';' +
        'try{' +
        'var s=atob(' + JSON.stringify(b64) + ');' +
        'var p=JSON.parse(s);' +
        'var t=null;try{t=window.top;}catch(e1){}' +
        'if(!t||t===window){try{t=window.parent;}catch(e2){}}' +
        'if(t&&t!==window){t.postMessage({source:"bpe_api",rid:rid,payload:p},"*");}' +
        '}catch(err){' +
        'var t2=null;try{t2=window.top;}catch(e3){}' +
        'if(!t2||t2===window){try{t2=window.parent;}catch(e4){}}' +
        'if(t2&&t2!==window){t2.postMessage({source:"bpe_api",rid:rid,payload:{ok:false,message:String(err)}}, "*");}' +
        '}' +
        '})();</script></head><body></body></html>';
      return ContentService.createTextOutput(html).setMimeType(ContentService.MimeType.HTML);
    }

    var callback = String(params.callback || '').trim();
    if (callback && /^[a-zA-Z_$][\w.$]*$/.test(callback)) {
      return ContentService.createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return textOut(json);
  }
  return textOut(
    'Blog Platform Engine API. Use GET ?api=json&cmd=... or JSONP with &callback=...'
  );
}

function serializeApiResult_(result) {
  if (result._rawJson) return result._rawJson;
  if (result.ok === false) {
    return JSON.stringify({ ok: false, message: result.message || 'Request failed' });
  }
  return JSON.stringify(result);
}

function doPost(e) {
  var payload = parseRequestBody(e);
  var result = handleApiRequest(payload);
  return textOut(serializeApiResult_(result));
}

function mergeParams(params, into) {
  var out = into || {};
  var keys = Object.keys(params || {});
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (out[k] === undefined || out[k] === '') out[k] = params[k];
  }
  return out;
}

function parseRequestBody(e) {
  var payload = {};
  if (e && e.postData && e.postData.contents) {
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (err) {
      payload = {};
    }
  }
  if (e && e.parameter) mergeParams(e.parameter, payload);
  return payload;
}

function handleApiRequest(payload) {
  var cmd = String(payload.cmd || payload.action || '').trim().toLowerCase();
  if (cmd === 'unsubscribe') cmd = 'unsub';

  if (cmd === 'ping') {
    var sheet = getPostLogSheet();
    return {
      ok: true,
      _rawJson: JSON.stringify({
        ok: true,
        message: 'API reachable',
        spreadsheetLinked: !!sheet,
        spreadsheetId: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || ''
      })
    };
  }

  if (cmd === 'confirm' || cmd === 'unsub') {
    return apiStatusUpdate(payload, cmd === 'confirm' ? STATUS_CONFIRMED : STATUS_UNSUBSCRIBED);
  }

  if (cmd === 'comments' || cmd === 'getcomments') {
    return apiGetComments(resolvePostId(payload));
  }

  if (cmd === 'portal' || cmd === 'getportal') {
    return apiGetPortal(resolvePostId(payload));
  }

  if (cmd === 'comment' || payload.comment) {
    return apiPostComment(payload);
  }

  return { ok: false, message: 'Unknown cmd. Use confirm, unsub, getComments, getPortal, or comment.' };
}

function apiStatusUpdate(payload, newStatus) {
  var postId = resolvePostId(payload);
  var email = String(payload.email || '').trim();
  if (!postId) return { ok: false, message: 'Missing post parameter.' };
  if (!email) return { ok: false, message: 'Missing email parameter.' };
  if (!getPostLogSheet()) {
    return {
      ok: false,
      message: 'Spreadsheet not linked. Open the PostLog sheet once, then run Setup.'
    };
  }

  var updated = updateSubscriberStatus(postId, email, newStatus);
  if (!updated) {
    return { ok: false, message: 'No matching subscription row for this post.' };
  }

  var fname = String(payload.fname || '').trim();
  var lname = String(payload.lname || '').trim();
  var body = {
    ok: true,
    status: 'success',
    message: newStatus === STATUS_CONFIRMED
      ? 'Your subscription is confirmed.'
      : 'You have successfully unsubscribed from this post.'
  };
  if (newStatus === STATUS_UNSUBSCRIBED) {
    var rq = { action: 'confirm', camp: postId, email: email };
    if (fname) rq.fname = fname;
    if (lname) rq.lname = lname;
    body.resubscribeUrl = buildGitHubPageUrl('index.html', rq);
  }
  return { ok: true, _rawJson: JSON.stringify(body) };
}

function apiGetComments(postId) {
  if (!postId) return { ok: false, message: 'Missing post parameter.' };
  return {
    ok: true,
    _rawJson: JSON.stringify({ ok: true, comments: loadComments(postId) })
  };
}

function apiGetPortal(postId) {
  if (!postId) return { ok: false, message: 'Missing post parameter.' };
  var meta = findPostMaster(postId);
  if (!meta.found) {
    return { ok: false, _rawJson: JSON.stringify({ ok: false, message: 'Post not found.' }) };
  }
  var deadlineIso = getDeadlineIso(meta.sentAt, meta.daysOpen);
  var expired = isExpired(deadlineIso, meta.sentAt, meta.daysOpen);
  return {
    ok: true,
    _rawJson: JSON.stringify({
      ok: true,
      postId: postId,
      emailSubject: meta.subject,
      emailBodyHtml: meta.messageHtml,
      attachments: resolveAttachmentList(meta.fileIdsRaw),
      daysOpen: meta.daysOpen,
      deadlineIso: deadlineIso,
      expired: expired,
      comments: loadComments(postId)
    })
  };
}

function apiPostComment(payload) {
  var postId = resolvePostId(payload);
  var text = String(payload.comment || '').trim();
  var email = String(payload.email || '').trim();
  var displayName = String(payload.displayName || '').trim();
  if (!displayName) {
    var fn = String(payload.fname || '').trim();
    var ln = String(payload.lname || '').trim();
    displayName = [fn, ln].filter(Boolean).join(' ').trim();
  }
  var parentId = String(payload.parentId || '').trim();

  if (!postId) return { ok: false, message: 'Post ID required.' };
  if (!text) return { ok: false, message: 'Comment cannot be empty.' };
  if (!email) return { ok: false, message: 'Email required.' };

  var meta = findPostMaster(postId);
  if (!meta.found) return { ok: false, message: 'Post not found.' };

  var deadlineIso = getDeadlineIso(meta.sentAt, meta.daysOpen);
  if (isExpired(deadlineIso, meta.sentAt, meta.daysOpen)) {
    return { ok: false, message: 'Discussion period has ended.', expired: true };
  }

  try {
    saveComment(postId, parentId, email, displayName || 'Guest', text);
    return {
      ok: true,
      _rawJson: JSON.stringify({
        ok: true,
        message: 'Comment posted.',
        comments: loadComments(postId)
      })
    };
  } catch (e) {
    return { ok: false, message: (e && e.message) ? e.message : 'Could not save comment.' };
  }
}

function textOut(msg) {
  return ContentService.createTextOutput(String(msg))
    .setMimeType(ContentService.MimeType.TEXT);
}

// ——— Post / comment data ———

function findPostMaster(webPostId) {
  var empty = {
    found: false, sentAt: null, daysOpen: 0,
    subject: '', messageHtml: '', fileIdsRaw: ''
  };
  var sheet = getPostLogSheet();
  var target = String(webPostId).trim();
  if (!sheet || !target) return empty;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return empty;

  var lastCol = Math.max(sheet.getLastColumn(), POST_LOG_HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colMap = buildColumnMap(headers);
  var clusters = buildPostClusters(sheet, headers, lastRow, lastCol);

  for (var i = clusters.length - 1; i >= 0; i--) {
    var routing = getRouting(clusters[i].master, colMap);
    if (routing.webPostId !== target) continue;
    var m = clusters[i].master;
    return {
      found: true,
      sentAt: parseDate(m[colIndex(colMap, HEADER_SENT, 2)]),
      daysOpen: parseFloat(m[colIndex(colMap, HEADER_DAYS_OPEN, 3)]) || 0,
      subject: String(m[colIndex(colMap, HEADER_SUBJECT, 7)] || ''),
      messageHtml: String(m[colIndex(colMap, HEADER_MESSAGE, 8)] || ''),
      fileIdsRaw: String(m[colIndex(colMap, HEADER_FILE_IDS, 9)] || '')
    };
  }
  return empty;
}

function updateSubscriberStatus(webPostId, email, newStatus) {
  var sheet = getPostLogSheet();
  if (!sheet) return false;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  var lastCol = Math.max(sheet.getLastColumn(), POST_LOG_HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colMap = buildColumnMap(headers);
  var emailCol = colIndex(colMap, HEADER_EMAIL, 6);
  var statusCol = colIndex(colMap, HEADER_STATUS, 1);
  var targetPost = String(webPostId).trim();
  var targetEmail = normalizeEmail_(email);
  var clusters = buildPostClusters(sheet, headers, lastRow, lastCol);

  for (var c = 0; c < clusters.length; c++) {
    if (getRouting(clusters[c].master, colMap).webPostId !== targetPost) continue;
    for (var r = 0; r < clusters[c].rows.length; r++) {
      var entry = clusters[c].rows[r];
      if (normalizeEmail_(entry.row[emailCol]) === targetEmail) {
        sheet.getRange(entry.sheetRow, statusCol + 1).setValue(newStatus);
        return true;
      }
    }
  }
  return false;
}

function normalizeEmail_(val) {
  return String(val || '').trim().replace(/\s+/g, '').toLowerCase();
}

function loadComments(postId) {
  var sheet = getCommentsSheet();
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var target = String(postId).trim();
  var list = [];

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][CMT.POST]).trim() !== target) continue;
    list.push({
      id: String(data[i][CMT.ID] || ''),
      parentId: String(data[i][CMT.PARENT] || ''),
      postId: target,
      timestamp: formatTs(data[i][CMT.TIME]),
      email: String(data[i][CMT.EMAIL] || ''),
      displayName: String(data[i][CMT.NAME] || ''),
      text: String(data[i][CMT.TEXT] || '')
    });
  }

  list.sort(function (a, b) {
    return new Date(a.timestamp) - new Date(b.timestamp);
  });
  return list;
}

function saveComment(postId, parentId, email, displayName, text) {
  if (parentId && !commentExists(parentId, postId)) {
    throw new Error('Parent comment not found.');
  }
  var sheet = ensureCommentsSheet();
  sheet.appendRow([
    Utilities.getUuid(),
    parentId || '',
    postId,
    new Date(),
    email,
    displayName,
    text
  ]);
}

function commentExists(commentId, postId) {
  var sheet = getCommentsSheet();
  if (!sheet) return false;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][CMT.ID]).trim() === String(commentId).trim() &&
      String(data[i][CMT.POST]).trim() === String(postId).trim()) {
      return true;
    }
  }
  return false;
}

function resolveAttachmentList(raw) {
  var audit = auditDriveFiles(raw, {});
  return audit.valid ? audit.files : [];
}

// ——— Sheet helpers ———

function rememberSpreadsheetId_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
    }
  } catch (err) {
    Logger.log('rememberSpreadsheetId_: ' + err.message);
  }
}

function openSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (err) {
      Logger.log('openById failed: ' + err.message);
    }
  }
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      props.setProperty('SPREADSHEET_ID', active.getId());
      return active;
    }
  } catch (err) {
    Logger.log('getActiveSpreadsheet failed: ' + err.message);
  }
  return null;
}

function getPostLogSheet() {
  var ss = openSpreadsheet_();
  if (!ss) return null;
  return ss.getSheetByName(SHEET_POST_LOG) || ss.getSheetByName('CampaignLog');
}

function getCommentsSheet() {
  var ss = openSpreadsheet_();
  if (!ss) return null;
  return ss.getSheetByName(SHEET_COMMENTS);
}

function ensureCommentsSheet() {
  var ss = openSpreadsheet_();
  if (!ss) throw new Error('Spreadsheet not linked. Open the sheet and run Setup.');
  return ensureSheetWithHeaders(ss, SHEET_COMMENTS, COMMENTS_LOG_HEADERS).sheet;
}

function ensureSheetWithHeaders(ss, name, headers, resetBody) {
  if (isWhitelistedSheet_(name)) {
    var existing = ss.getSheetByName(name);
    return { sheet: existing, created: false, headersReset: false };
  }
  var sheet = ss.getSheetByName(name);
  var created = false;
  if (!sheet) {
    sheet = ss.insertSheet(name);
    created = true;
  }
  var headerMismatch = sheet.getLastColumn() < headers.length;
  if (!created && sheet.getLastRow() >= 1) {
    var current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
    for (var h = 0; h < headers.length; h++) {
      if (String(current[h] || '').trim() !== headers[h]) {
        headerMismatch = true;
        break;
      }
    }
  }
  if (created || !sheet.getRange(1, 1).getValue() || headerMismatch) {
    if (resetBody && sheet.getMaxRows() > 1) {
      var lr = Math.max(sheet.getLastRow(), 2);
      if (lr > 1) {
        sheet.getRange(2, 1, lr - 1, headers.length).clearContent();
      }
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
  return { sheet: sheet, created: created, headersReset: headerMismatch && !created };
}

function parseDate(val) {
  if (isBlankTimestamp(val)) return null;
  if (val instanceof Date) return val;
  var d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function formatTs(val) {
  if (!val) return '';
  if (val instanceof Date) return val.toISOString();
  return String(val);
}
