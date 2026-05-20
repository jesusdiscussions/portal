/**
 * Blog Platform Engine — Google Apps Script backend
 * CampaignLog columns: A=Post ID … (row 1 = dynamic headers)
 * CommentsLog columns: A=Comment ID … C=Post ID (header C1)
 */

var SHEET_POST_LOG = 'CampaignLog';
var SHEET_COMMENTS = 'CommentsLog';

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

var MASTER_ROUTING_MARKER = 'post id';

var CMT = {
  ID: 0,
  PARENT: 1,
  POST: 2,
  TIMESTAMP: 3,
  EMAIL: 4,
  NAME: 5,
  TEXT: 6
};

var STATUS_SUBSCRIBED = 'SUBSCRIBED';
var STATUS_CONFIRMED = 'CONFIRMED';
var STATUS_UNSUBSCRIBED = 'UNSUBSCRIBED';
var ATTACH_ERROR_BG = '#ffcccc';

/**
 * GitHub Pages landing page (no trailing slash).
 * Override via Script property GITHUB_PAGES_BASE_URL, or edit the default below.
 */
var GITHUB_PAGES_BASE_URL_DEFAULT = 'https://jesusdiscussions.github.io/portal';

/**
 * Deployed web app /exec URL (Emailer Backend @16).
 * Keep in sync with WEBAPP_URL in index.html and CommentsPage.html.
 * Override via Script property WEB_APP_URL.
 */
var WEB_APP_URL_DEFAULT =
  'https://script.google.com/macros/s/AKfycbxa3oj2BTcEQladm9llk7yyMYU2vcy7GzQ_dqbnPVGAUiSY9Imf_VZqjIJAHSDD-Q_wbw/exec';

var POST_LOG_HEADERS = [
  HEADER_POST_ID,
  HEADER_STATUS,
  HEADER_SENT,
  HEADER_DAYS_OPEN,
  HEADER_FIRST,
  HEADER_LAST,
  HEADER_EMAIL,
  HEADER_SUBJECT,
  HEADER_MESSAGE,
  HEADER_FILE_IDS
];

var COMMENTS_LOG_HEADERS = [
  'Comment ID',
  'Parent ID',
  HEADER_POST_ID,
  'Timestamp',
  'User Email',
  'User Display Name',
  'Comment Text'
];

// ——— Spreadsheet menu ———

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Blog Platform Engine')
    .addItem('Setup sheets (one-shot)', 'setupSheets')
    .addItem('Set GitHub Pages base URL…', 'promptGitHubPagesUrl')
    .addItem('Open web app deployments…', 'openWebAppDeployments')
    .addSeparator()
    .addItem('Run Post Send & Sync', 'runPostEngine')
    .addToUi();
}

/** Opens Apps Script deployment manager (set access to Anyone, then New version). */
function openWebAppDeployments() {
  var scriptId = ScriptApp.getScriptId();
  var url = 'https://script.google.com/home/projects/' + scriptId + '/deployments';
  var html = HtmlService.createHtmlOutput(
    '<p>Update <b>Emailer Backend</b>: Edit → New version → Who has access: <b>Anyone</b></p>' +
    '<p><a href="' + url + '" target="_blank">Open Deployments</a></p>' +
    '<script>window.open(' + JSON.stringify(url) + ');</script>'
  ).setWidth(420).setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, 'Web app access');
}

/** Saves GITHUB_PAGES_BASE_URL script property (used by email link factory). */
function promptGitHubPagesUrl() {
  var ui = SpreadsheetApp.getUi();
  var current = PropertiesService.getScriptProperties().getProperty('GITHUB_PAGES_BASE_URL') ||
    GITHUB_PAGES_BASE_URL_DEFAULT;
  var result = ui.prompt(
    'GitHub Pages base URL',
    'Enter the URL with no trailing slash.\nCurrent: ' + current,
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  var url = result.getResponseText().trim().replace(/\/$/, '');
  if (!url) {
    ui.alert('URL was not saved (empty input).');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('GITHUB_PAGES_BASE_URL', url);
  ui.alert('GitHub Pages URL saved:\n' + url);
}

/**
 * One-shot initializer: creates CampaignLog / CommentsLog tabs, writes header rows,
 * freezes row 1, and adds a sample subscriber row when CampaignLog is new.
 */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var postResult = ensureSheetWithHeaders(ss, SHEET_POST_LOG, POST_LOG_HEADERS);
  ensureSheetWithHeaders(ss, SHEET_COMMENTS, COMMENTS_LOG_HEADERS);

  if (postResult.created) {
    addSamplePostRow(postResult.sheet);
  }

  SpreadsheetApp.getUi().alert(
    'Blog Platform Engine sheets are ready.\n\n' +
    '• ' + SHEET_POST_LOG + (postResult.created ? ' (created + sample row)' : '') + '\n' +
    '• ' + SHEET_COMMENTS + '\n\n' +
    'Deploy the script as a web app, then use Run Post Send & Sync when rows are ready.'
  );
}

// ——— Post broadcast (grouped by Post ID) ———

function runPostEngine() {
  var sheet = getPostLogSheet();
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Missing tab: "' + SHEET_POST_LOG + '".');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  var lastCol = Math.max(sheet.getLastColumn(), POST_LOG_HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colMap = buildColumnMap(headers);
  var clusters = buildPostClusters(sheet, headers, lastRow, lastCol);

  if (clusters.length === 0) {
    return;
  }

  for (var v = 0; v < clusters.length; v++) {
    var validationError = validateMasterBlueprint(clusters[v].master, colMap);
    if (validationError) {
      SpreadsheetApp.getUi().alert(validationError);
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
    if (name) {
      map[name] = i;
    }
  }
  return map;
}

function colIndex(colMap, headerName, fallback) {
  if (colMap[headerName] !== undefined) {
    return colMap[headerName];
  }
  return fallback;
}

function buildPostClusters(sheet, headers, lastRow, lastCol) {
  var rows = getRangeByRowSpan(sheet, 2, 1, lastRow, lastCol).getValues();
  var postCol = colIndex(buildColumnMap(headers), HEADER_POST_ID, 0);
  var clusters = [];
  var clusterMap = {};
  var clusterOrder = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!rowHasAnyContent(row)) {
      continue;
    }

    var postKey = String(row[postCol] || '').trim();
    if (!postKey) {
      continue;
    }

    if (!clusterMap[postKey]) {
      clusterMap[postKey] = {
        postId: postKey,
        master: row,
        masterSheetRow: i + 2,
        rows: []
      };
      clusterOrder.push(postKey);
    }

    clusterMap[postKey].rows.push({
      row: row,
      sheetRow: i + 2
    });
  }

  for (var k = 0; k < clusterOrder.length; k++) {
    clusters.push(clusterMap[clusterOrder[k]]);
  }

  return clusters;
}

function rowHasAnyContent(row) {
  for (var i = 0; i < row.length; i++) {
    if (String(row[i] || '').trim() !== '') {
      return true;
    }
  }
  return false;
}

/**
 * Master Blueprint gate: First Name must be "Post ID", Post ID and Last Name required.
 * @return {string|null} Error message or null if valid.
 */
function validateMasterBlueprint(masterRow, colMap) {
  var postCol = colIndex(colMap, HEADER_POST_ID, 0);
  var firstCol = colIndex(colMap, HEADER_FIRST, 4);
  var lastCol = colIndex(colMap, HEADER_LAST, 5);

  var postId = String(masterRow[postCol] || '').trim();
  var firstName = String(masterRow[firstCol] || '').trim();
  var lastName = String(masterRow[lastCol] || '').trim();
  var clusterLabel = postId || '(unknown cluster)';

  if (!firstName) {
    return 'HALTED — Master Blueprint for Post ID cluster "' + clusterLabel + '":\n\n' +
      '• First Name is blank.\n' +
      '• It must be exactly "Post ID" (case-insensitive) to authorize routing.';
  }

  if (firstName.toLowerCase() !== MASTER_ROUTING_MARKER) {
    return 'HALTED — Master Blueprint for Post ID cluster "' + clusterLabel + '":\n\n' +
      '• First Name is misspelled: found "' + firstName + '".\n' +
      '• It must be exactly "Post ID" (case-insensitive).';
  }

  if (!postId) {
    return 'HALTED — Master Blueprint row:\n\n' +
      '• The leftmost Post ID field is blank.\n' +
      '• Enter a Post ID before running.';
  }

  if (!lastName) {
    return 'HALTED — Master Blueprint for Post ID cluster "' + clusterLabel + '":\n\n' +
      '• Last Name is blank.\n' +
      '• Enter the web destination Post number (Launch: same as Post ID; Modify: target Post ID).';
  }

  return null;
}

function getRoutingFromMaster(masterRow, colMap) {
  var postCol = colIndex(colMap, HEADER_POST_ID, 0);
  var lastCol = colIndex(colMap, HEADER_LAST, 5);
  var trackingPostId = String(masterRow[postCol] || '').trim();
  var lastName = String(masterRow[lastCol] || '').trim();
  var isLaunch = trackingPostId === lastName;

  return {
    trackingPostId: trackingPostId,
    webPostId: isLaunch ? trackingPostId : lastName,
    isLaunch: isLaunch,
    isModify: !isLaunch
  };
}

function processPostCluster(cluster, sheet, headers, colMap, driveCache) {
  var routing = getRoutingFromMaster(cluster.master, colMap);
  var master = cluster.master;
  var subjectCol = colIndex(colMap, HEADER_SUBJECT, 7);
  var messageCol = colIndex(colMap, HEADER_MESSAGE, 8);
  var fileIdsCol = colIndex(colMap, HEADER_FILE_IDS, 9);
  var statusCol = colIndex(colMap, HEADER_STATUS, 1);
  var sentCol = colIndex(colMap, HEADER_SENT, 2);
  var daysCol = colIndex(colMap, HEADER_DAYS_OPEN, 3);
  var emailCol = colIndex(colMap, HEADER_EMAIL, 6);
  var firstCol = colIndex(colMap, HEADER_FIRST, 4);
  var lastCol = colIndex(colMap, HEADER_LAST, 5);

  var templateSubject = String(master[subjectCol] || '').trim();
  var templateMessage = String(master[messageCol] || '');
  var templateFileIds = master[fileIdsCol];
  var daysOpen = parseFloat(master[daysCol]) || 0;

  var portalPayload = {
    postId: routing.webPostId,
    trackingPostId: routing.trackingPostId,
    subject: templateSubject,
    messageHtml: templateMessage,
    fileIdsRaw: String(templateFileIds || ''),
    daysOpen: daysOpen,
    mode: routing.isLaunch ? 'launch' : 'modify'
  };

  syncWebPostToGitHub(portalPayload, routing.isLaunch);

  for (var i = 0; i < cluster.rows.length; i++) {
    var entry = cluster.rows[i];
    var row = entry.row;
    var sheetRow = entry.sheetRow;

    if (!rowHasSendableContent(row, colMap)) {
      continue;
    }

    var status = String(row[statusCol] || '').trim().toUpperCase();
    if (status !== STATUS_SUBSCRIBED && status !== STATUS_CONFIRMED) {
      continue;
    }

    if (!isSentTimestampBlank(row[sentCol])) {
      continue;
    }

    var recipient = String(row[emailCol] || '').trim();
    if (!recipient) {
      continue;
    }

    var fileIdsCell = sheet.getRange(sheetRow, fileIdsCol + 1);
    var rowFileIds = String(row[fileIdsCol] || '').trim() ? row[fileIdsCol] : templateFileIds;

    var audit = auditDriveFileIds(rowFileIds, driveCache);

    if (!audit.valid) {
      fileIdsCell.setBackground(ATTACH_ERROR_BG);
      fileIdsCell.setNote(audit.errorNote);
      continue;
    }

    fileIdsCell.setBackground(null);
    fileIdsCell.clearNote();

    var mergedSubject = applyMergeTags(templateSubject, headers, row);
    var mergedMessage = applyMergeTags(templateMessage, headers, row);
    var first = String(row[firstCol] || '').trim();
    var last = String(row[lastCol] || '').trim();
    var email = String(row[emailCol] || '').trim();

    try {
      sendPostEmail(
        mergedSubject,
        mergedMessage,
        routing.webPostId,
        recipient,
        audit.files,
        email,
        first,
        last,
        daysOpen
      );
      sheet.getRange(sheetRow, sentCol + 1).setValue(new Date());
    } catch (err) {
      Logger.log('Row ' + sheetRow + ' send failed: ' + err.message);
    }
  }
}

function rowHasSendableContent(row, colMap) {
  var emailCol = colIndex(colMap, HEADER_EMAIL, 6);
  var subjectCol = colIndex(colMap, HEADER_SUBJECT, 7);
  var messageCol = colIndex(colMap, HEADER_MESSAGE, 8);
  return String(row[emailCol] || '').trim() !== '' ||
    String(row[subjectCol] || '').trim() !== '' ||
    String(row[messageCol] || '').trim() !== '';
}

function applyMergeTags(template, headers, row) {
  var result = String(template || '');
  for (var i = 0; i < headers.length; i++) {
    var header = String(headers[i] || '').trim();
    if (!header) {
      continue;
    }
    var tag = '{{' + header + '}}';
    var value = row[i];
    if (value instanceof Date) {
      value = value.toString();
    }
    if (value === null || value === undefined) {
      value = '';
    }
    result = result.split(tag).join(String(value));
  }
  return result;
}

function isSentTimestampBlank(val) {
  return val === '' || val === null || val === undefined;
}

/**
 * Validates comma-separated Drive IDs using a per-run cache.
 * Files must exist and be shared with anyone (link or full public).
 */
function auditDriveFileIds(rawIds, cache) {
  var ids = parseCommaSeparatedIds(rawIds);
  if (ids.length === 0) {
    return { valid: true, files: [], errorNote: '' };
  }

  var files = [];
  for (var i = 0; i < ids.length; i++) {
    var id = ids[i];
    var cached = cache[id];

    if (!cached) {
      cached = verifyDriveFile(id);
      cache[id] = cached;
    }

    if (!cached.ok) {
      return {
        valid: false,
        files: [],
        errorNote: 'Drive file ID "' + id + '": ' + cached.error
      };
    }

    files.push({
      id: id,
      name: cached.name,
      url: cached.url
    });
  }

  return { valid: true, files: files, errorNote: '' };
}

function parseCommaSeparatedIds(raw) {
  if (!raw) {
    return [];
  }
  return String(raw)
    .split(',')
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s.length > 0; });
}

function verifyDriveFile(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var access = file.getSharingAccess();
    var isPublic = access === DriveApp.Access.ANYONE ||
      access === DriveApp.Access.ANYONE_WITH_LINK;

    if (!isPublic) {
      return {
        ok: false,
        error: 'Sharing is restricted (file must be public or anyone-with-link).'
      };
    }

    return {
      ok: true,
      name: file.getName(),
      url: 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(fileId)
    };
  } catch (e) {
    return {
      ok: false,
      error: (e && e.message) ? e.message : 'Invalid or inaccessible file ID.'
    };
  }
}

function sendPostEmail(subject, bodyHtml, webPostId, recipient, verifiedFiles, email, first, last, daysOpen) {
  var subjectLine = String(subject || '').trim() || 'Blog Post Message';
  var sentNow = new Date();
  var deadlineIso = getDeadlineIsoFromSentAndDays(sentNow, daysOpen);
  var footer = buildTrackingFooter(webPostId, email, first, last, deadlineIso);
  var html = bodyHtml + footer;

  var blobs = [];
  for (var i = 0; i < verifiedFiles.length; i++) {
    blobs.push(DriveApp.getFileById(verifiedFiles[i].id).getBlob());
  }

  var options = { htmlBody: html };
  if (blobs.length > 0) {
    options.attachments = blobs;
  }

  var plain = stripHtml(html);
  MailApp.sendEmail(recipient, subjectLine, plain, options);
}

// ——— GitHub web post sync (optional; requires script properties) ———

/**
 * Pushes portal payload JSON to GitHub (posts/{webPostId}.json).
 * Launch: create file if missing. Modify: overwrite existing post file.
 * Set GITHUB_TOKEN and GITHUB_REPO (e.g. "owner/repo") script properties to enable.
 */
function syncWebPostToGitHub(portalPayload, isLaunch) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('GITHUB_TOKEN');
  var repo = props.getProperty('GITHUB_REPO');
  if (!token || !repo) {
    Logger.log('GitHub sync skipped (GITHUB_TOKEN or GITHUB_REPO not set). Post ' + portalPayload.postId);
    return;
  }

  var path = 'posts/' + portalPayload.postId + '.json';
  var content = JSON.stringify({
    postId: portalPayload.postId,
    trackingPostId: portalPayload.trackingPostId,
    subject: portalPayload.subject,
    messageHtml: portalPayload.messageHtml,
    fileIdsRaw: portalPayload.fileIdsRaw,
    daysOpen: portalPayload.daysOpen,
    mode: portalPayload.mode,
    updatedAt: new Date().toISOString()
  }, null, 2);

  var encoded = Utilities.base64Encode(content, Utilities.Charset.UTF_8);
  var apiUrl = 'https://api.github.com/repos/' + repo + '/contents/' + path;
  var existingSha = null;

  if (!isLaunch) {
    try {
      var getResp = UrlFetchApp.fetch(apiUrl, {
        method: 'get',
        headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
        muteHttpExceptions: true
      });
      if (getResp.getResponseCode() === 200) {
        existingSha = JSON.parse(getResp.getContentText()).sha;
      }
    } catch (e) {
      Logger.log('GitHub GET for modify: ' + e.message);
    }
  }

  var body = {
    message: (isLaunch ? 'Launch post ' : 'Modify post ') + portalPayload.postId,
    content: encoded,
    branch: props.getProperty('GITHUB_BRANCH') || 'main'
  };
  if (existingSha) {
    body.sha = existingSha;
  }

  var resp = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/vnd.github+json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) {
    Logger.log('GitHub sync failed for post ' + portalPayload.postId + ': ' + resp.getContentText());
  }
}

// ——— Dynamic URL factory (GitHub Pages) ———

function getGitHubPagesBaseUrl() {
  var fromProps = PropertiesService.getScriptProperties().getProperty('GITHUB_PAGES_BASE_URL');
  var base = (fromProps || GITHUB_PAGES_BASE_URL_DEFAULT).trim();
  return base.replace(/\/$/, '');
}

/** Active Apps Script web app /exec URL (used by static HTML fetch targets). */
function getWebAppUrl() {
  var fromProps = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL');
  if (fromProps) {
    return fromProps.trim();
  }
  try {
    var url = ScriptApp.getService().getUrl();
    if (url) {
      return url;
    }
  } catch (err) {
    // Not deployed as a web app in this context.
  }
  return WEB_APP_URL_DEFAULT;
}

/**
 * Resolves Post ID from URL/query map (canonical post, legacy camp).
 */
function resolvePostIdFromUrlParameters(params) {
  if (!params) {
    return '';
  }
  var post = String(params.post || '').trim();
  if (post) {
    return post;
  }
  return String(params.camp || '').trim();
}

/** Subscriber comment portal (served by this web app doGet). */
function buildCommentsPageLink(postId, email, fname, lname, deadlineIso) {
  var params = [
    'page=comments',
    'post=' + encodeURIComponent(postId),
    'email=' + encodeURIComponent(email),
    'fname=' + encodeURIComponent(fname || ''),
    'lname=' + encodeURIComponent(lname || '')
  ];
  if (deadlineIso) {
    params.push('deadline=' + encodeURIComponent(deadlineIso));
  }
  return getWebAppUrl() + '?' + params.join('&');
}

function buildPostLink(action, postId, email, fname, lname, deadlineIso) {
  var normalized = normalizeAction(action);
  if (normalized === 'comments') {
    return buildCommentsPageLink(postId, email, fname, lname, deadlineIso);
  }
  var params = [
    'action=' + encodeURIComponent(normalized),
    'post=' + encodeURIComponent(postId),
    'email=' + encodeURIComponent(email),
    'fname=' + encodeURIComponent(fname || ''),
    'lname=' + encodeURIComponent(lname || '')
  ];
  return getWebAppUrl() + '?' + params.join('&');
}

function getDeadlineIsoFromSentAndDays(sentAt, daysOpen) {
  var days = parseFloat(daysOpen);
  if (!sentAt || isNaN(days) || days <= 0) {
    return '';
  }
  return new Date(sentAt.getTime() + days * 24 * 60 * 60 * 1000).toISOString();
}

function isExpiredByDeadline(deadlineIso) {
  if (!deadlineIso) {
    return false;
  }
  var expiry = new Date(deadlineIso);
  return !isNaN(expiry.getTime()) && Date.now() >= expiry.getTime();
}

function normalizeAction(action) {
  var a = String(action || '').trim().toLowerCase();
  if (a === 'unsubscribe') {
    return 'unsub';
  }
  if (a === 'portal') {
    return 'comments';
  }
  return a;
}

function buildTrackingFooter(postId, email, fname, lname, deadlineIso) {
  return '<hr style="margin-top:24px;border:none;border-top:1px solid #ddd;">' +
    '<p style="font-size:13px;color:#555;">' +
    '<a href="' + buildPostLink('confirm', postId, email, fname, lname) + '">Confirm subscription</a> &nbsp;|&nbsp; ' +
    '<a href="' + buildPostLink('unsub', postId, email, fname, lname) + '">Unsubscribe</a> &nbsp;|&nbsp; ' +
    '<a href="' + buildPostLink('comments', postId, email, fname, lname, deadlineIso) + '">Join Comments</a>' +
    '</p>';
}

function stripHtml(html) {
  return String(html).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

// ——— Web app routing ———

/**
 * GET: serves subscriber HTML (index / CommentsPage) or JSON API (?api=json).
 * Email links and legacy GitHub Pages redirects should land here (not cross-origin fetch).
 */
function doGet(e) {
  var params = (e && e.parameter) || {};
  var postId = resolvePostIdFromUrlParameters(params);
  var action = normalizeAction(params.action || '');
  var page = String(params.page || '').trim().toLowerCase();

  if (String(params.api || '').trim().toLowerCase() === 'json') {
    var payload = parseIncomingPayload({ parameter: params });
    if (payload.comment) {
      return corsJson(handleCommentSubmissionData(payload));
    }
    return corsJson(handleExternalActionData(payload));
  }

  if (page === 'comments' || action === 'comments') {
    return HtmlService.createHtmlOutputFromFile('CommentsPage')
      .setTitle('Post — Community')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (action === 'confirm' || action === 'unsub') {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Post')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (postId && String(params.email || '').trim()) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Post')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return textPage(
    'Blog Platform Engine — open a link from your post email, or use ' +
      getWebAppUrl() + ' with post, action, and email parameters.',
    'Blog Platform Engine'
  );
}

/**
 * POST: GitHub Pages actions (JSON body) and CommentsPage comment submissions (form fields).
 */
function doPost(e) {
  var payload = parseIncomingPayload(e);

  if (payload.comment) {
    return handleCommentSubmission(payload);
  }

  return handleExternalAction(payload);
}

function parseIncomingPayload(e) {
  var payload = {};

  if (e && e.postData && e.postData.contents) {
    var type = String(e.postData.type || '').toLowerCase();
    var raw = e.postData.contents;
    if (type.indexOf('application/json') !== -1 || type.indexOf('text/plain') !== -1) {
      try {
        payload = JSON.parse(raw);
      } catch (err) {
        payload = {};
      }
    }
  }

  if (e && e.parameter) {
    var keys = Object.keys(e.parameter);
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      if (payload[key] === undefined || payload[key] === '') {
        payload[key] = e.parameter[key];
      }
    }
  }

  return payload;
}

/** Resolves post ID from JSON/form payload (post or legacy camp). */
function resolvePostIdFromPayload(payload) {
  return resolvePostIdFromUrlParameters(payload);
}

/** Called from HtmlService via google.script.run (no cross-origin fetch). */
function apiExternalAction(payload) {
  return handleExternalActionData(payload);
}

/** Called from HtmlService via google.script.run. */
function apiCommentSubmit(payload) {
  return handleCommentSubmissionData(payload);
}

function handleExternalAction(payload) {
  return corsJson(handleExternalActionData(payload));
}

function handleExternalActionData(payload) {
  var action = normalizeAction(payload.action);
  var postId = resolvePostIdFromPayload(payload);
  var email = String(payload.email || '').trim();
  var fname = String(payload.fname || '').trim();
  var lname = String(payload.lname || '').trim();

  if (!action) {
    return { status: 'error', message: 'Missing action parameter.' };
  }
  if (!postId) {
    return { status: 'error', message: 'Missing post (Post ID) parameter.' };
  }

  if (action === 'confirm') {
    if (!email) {
      return { status: 'error', message: 'Missing email parameter.' };
    }
    var confirmed = updatePostStatus(postId, email, STATUS_CONFIRMED);
    if (!confirmed) {
      return {
        status: 'error',
        message: 'We could not find a matching subscription for this post.'
      };
    }
    return {
      status: 'success',
      message: 'Your subscription is confirmed.'
    };
  }

  if (action === 'unsub') {
    if (!email) {
      return { status: 'error', message: 'Missing email parameter.' };
    }
    var unsubscribed = updatePostStatus(postId, email, STATUS_UNSUBSCRIBED);
    if (!unsubscribed) {
      return {
        status: 'error',
        message: 'We could not find a matching subscription for this post.'
      };
    }
    return {
      status: 'success',
      message: 'You have successfully unsubscribed from this post.',
      resubscribeUrl: buildPostLink('confirm', postId, email, fname, lname)
    };
  }

  if (action === 'comments') {
    return getCommentsPortalPayload(postId, email, fname, lname);
  }

  return { status: 'error', message: 'Unknown action: ' + action };
}

/** Portal data bundle for GitHub Pages comments view. */
function getCommentsPortalPayload(postId, email, fname, lname) {
  if (!email) {
    return { status: 'error', message: 'Missing email parameter.' };
  }

  var meta = findPostByWebId(postId);
  if (!meta.found) {
    return { status: 'error', message: 'Post not found.' };
  }

  var deadlineIso = getDeadlineIsoFromSentAndDays(meta.sentAt, meta.daysOpen);
  var expired = isExpiredByDeadline(deadlineIso) ||
    isPostExpired(meta.sentAt, meta.daysOpen);

  return {
    status: 'success',
    postId: postId,
    email: email,
    fname: fname,
    lname: lname,
    emailSubject: meta.subject,
    emailBodyHtml: meta.messageHtml,
    attachments: resolveAttachmentList(meta.fileIdsRaw),
    comments: getCommentsForPost(postId),
    deadlineIso: deadlineIso,
    expired: expired
  };
}

function handleCommentSubmission(payload) {
  return corsJson(handleCommentSubmissionData(payload));
}

function handleCommentSubmissionData(payload) {
  var postId = resolvePostIdFromPayload(payload);
  var text = String(payload.comment || '').trim();
  var email = String(payload.email || '').trim();
  var fname = String(payload.fname || '').trim();
  var lname = String(payload.lname || '').trim();
  var parentId = String(payload.parentId || '').trim();

  if (!postId) {
    return { status: 'error', message: 'Post ID is required.', success: false };
  }
  if (!text) {
    return { status: 'error', message: 'Comment cannot be empty.', success: false };
  }

  var meta = findPostByWebId(postId);
  if (!meta.found) {
    return { status: 'error', message: 'Post not found.', success: false };
  }
  var deadlineIso = getDeadlineIsoFromSentAndDays(meta.sentAt, meta.daysOpen);
  if (isExpiredByDeadline(deadlineIso) || isPostExpired(meta.sentAt, meta.daysOpen)) {
    return {
      status: 'error',
      message: 'The discussion period has ended.',
      expired: true,
      success: false
    };
  }

  var displayName = [fname, lname].filter(Boolean).join(' ').trim() || 'Anonymous';

  try {
    appendComment(postId, parentId, email, displayName, text);
    return {
      status: 'success',
      success: true,
      message: 'Comment posted.',
      comments: getCommentsForPost(postId)
    };
  } catch (err) {
    return {
      status: 'error',
      success: false,
      message: (err && err.message) ? err.message : 'Could not save comment.'
    };
  }
}

function textPage(message, title) {
  return HtmlService.createHtmlOutput('<p>' + escapeHtml(message) + '</p>').setTitle(title);
}

/**
 * JSON API response for cross-origin fetch (GitHub Pages, CommentsPage).
 * Deploy web app as "Anyone" so browsers can read the response body.
 */
function corsJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ——— Post & comment data ———

/**
 * Finds portal content by web Post ID (CommentsLog column C / URL post param).
 * Launch masters: web ID = Post ID. Modify masters: web ID = Last Name on master row.
 */
function findPostByWebId(webPostId) {
  var sheet = getPostLogSheet();
  var target = String(webPostId).trim();
  var empty = {
    found: false,
    sentAt: null,
    daysOpen: 0,
    subject: '',
    messageHtml: '',
    fileIdsRaw: ''
  };

  if (!sheet || !target) {
    return empty;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return empty;
  }

  var lastCol = Math.max(sheet.getLastColumn(), POST_LOG_HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colMap = buildColumnMap(headers);
  var clusters = buildPostClusters(sheet, headers, lastRow, lastCol);

  for (var i = clusters.length - 1; i >= 0; i--) {
    var cluster = clusters[i];
    var routing = getRoutingFromMaster(cluster.master, colMap);
    if (routing.webPostId !== target) {
      continue;
    }

    var master = cluster.master;
    var subjectCol = colIndex(colMap, HEADER_SUBJECT, 7);
    var messageCol = colIndex(colMap, HEADER_MESSAGE, 8);
    var fileIdsCol = colIndex(colMap, HEADER_FILE_IDS, 9);
    var sentCol = colIndex(colMap, HEADER_SENT, 2);
    var daysCol = colIndex(colMap, HEADER_DAYS_OPEN, 3);

    return {
      found: true,
      sentAt: parseSentTimestamp(master[sentCol]),
      daysOpen: parseFloat(master[daysCol]) || 0,
      subject: String(master[subjectCol] || ''),
      messageHtml: String(master[messageCol] || ''),
      fileIdsRaw: String(master[fileIdsCol] || ''),
      rowIndex: cluster.masterSheetRow
    };
  }

  return empty;
}

function updatePostStatus(webPostId, email, newStatus) {
  var sheet = getPostLogSheet();
  if (!sheet) {
    return false;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return false;
  }

  var lastCol = Math.max(sheet.getLastColumn(), POST_LOG_HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colMap = buildColumnMap(headers);
  var emailCol = colIndex(colMap, HEADER_EMAIL, 6);
  var statusCol = colIndex(colMap, HEADER_STATUS, 1);
  var targetPost = String(webPostId).trim();
  var targetEmail = String(email).trim().toLowerCase();
  var clusters = buildPostClusters(sheet, headers, lastRow, lastCol);

  for (var c = 0; c < clusters.length; c++) {
    var routing = getRoutingFromMaster(clusters[c].master, colMap);
    if (routing.webPostId !== targetPost) {
      continue;
    }
    for (var i = 0; i < clusters[c].rows.length; i++) {
      var entry = clusters[c].rows[i];
      var rowEmail = String(entry.row[emailCol] || '').trim().toLowerCase();
      if (rowEmail === targetEmail) {
        sheet.getRange(entry.sheetRow, statusCol + 1).setValue(newStatus);
        return true;
      }
    }
  }

  return false;
}

function getCommentsForPost(postId) {
  var sheet = getCommentsLogSheet();
  if (!sheet) {
    return [];
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  var data = getRangeByRowSpan(sheet, 2, 1, lastRow, 7).getValues();
  var target = String(postId).trim();
  var list = [];

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][CMT.POST]).trim() !== target) {
      continue;
    }
    list.push({
      id: String(data[i][CMT.ID] || ''),
      parentId: String(data[i][CMT.PARENT] || ''),
      postId: target,
      timestamp: formatTimestamp(data[i][CMT.TIMESTAMP]),
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

function appendComment(postId, parentId, userEmail, displayName, text) {
  var sheet = getOrCreateCommentsLogSheet();

  if (parentId) {
    var exists = commentExists(parentId, postId);
    if (!exists) {
      throw new Error('Parent comment not found.');
    }
  }

  var commentId = Utilities.getUuid();
  sheet.appendRow([
    commentId,
    parentId || '',
    postId,
    new Date(),
    userEmail || '',
    displayName,
    text
  ]);

  return commentId;
}

function commentExists(commentId, postId) {
  var sheet = getCommentsLogSheet();
  if (!sheet) {
    return false;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return false;
  }

  var data = getRangeByRowSpan(sheet, 2, 1, lastRow, 7).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][CMT.ID]).trim() === String(commentId).trim() &&
      String(data[i][CMT.POST]).trim() === String(postId).trim()) {
      return true;
    }
  }

  return false;
}

function resolveAttachmentList(rawIds) {
  var cache = {};
  var audit = auditDriveFileIds(rawIds, cache);
  return audit.valid ? audit.files : [];
}

// ——— Sheet helpers ———

/**
 * Sheet.getRange(row, column, numRows, numColumns) — the third argument is a row
 * count, not the last row index. This helper spans startRow through endRow inclusive.
 */
function getRangeByRowSpan(sheet, startRow, startCol, endRow, numCols) {
  var numRows = endRow - startRow + 1;
  if (numRows < 1) {
    throw new Error('Invalid row span: endRow must be >= startRow.');
  }
  return sheet.getRange(startRow, startCol, numRows, numCols);
}

function getPostLogSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_POST_LOG);
}

function getCommentsLogSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COMMENTS);
}

function getOrCreateCommentsLogSheet() {
  return ensureSheetWithHeaders(
    SpreadsheetApp.getActiveSpreadsheet(),
    SHEET_COMMENTS,
    COMMENTS_LOG_HEADERS
  ).sheet;
}

function ensureSheetWithHeaders(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  var created = false;

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    created = true;
  }

  var needsHeaders = created || isHeaderRowEmpty(sheet);
  if (needsHeaders) {
    applyHeaderRow(sheet, headers);
  }

  return { sheet: sheet, created: created };
}

function isHeaderRowEmpty(sheet) {
  var first = sheet.getRange(1, 1).getValue();
  return first === '' || first === null;
}

function applyHeaderRow(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#f3f3f3');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function addSamplePostRow(sheet) {
  if (sheet.getLastRow() > 1) {
    return;
  }
  getRangeByRowSpan(sheet, 2, 1, 2, POST_LOG_HEADERS.length).setValues([[
    '',
    STATUS_SUBSCRIBED,
    '',
    7,
    MASTER_ROUTING_MARKER,
    '',
    'jane.doe@example.com',
    'Welcome to the community',
    '<p>Thanks for joining, {{First Name}}. Share your thoughts in the portal after launch.</p>',
    ''
  ]]);
}

function parseSentTimestamp(val) {
  if (isSentTimestampBlank(val)) {
    return null;
  }
  if (val instanceof Date) {
    return val;
  }
  var parsed = new Date(val);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function isPostExpired(sentAt, daysOpen) {
  var days = parseFloat(daysOpen);
  if (!sentAt || isNaN(days) || days <= 0) {
    return false;
  }
  var expiryMs = sentAt.getTime() + days * 24 * 60 * 60 * 1000;
  return Date.now() >= expiryMs;
}

function formatTimestamp(val) {
  if (!val) {
    return '';
  }
  if (val instanceof Date) {
    return val.toISOString();
  }
  return String(val);
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
