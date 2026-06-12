/**
 * @fileoverview GAS Project Auditor — Main Orchestration File
 *
 * Phases:
 *   1. Discovery       — getAllScriptProjects()      — Drive.Files.list() across all drives
 *   2. Deployments     — getWebAppDeployments()      — Apps Script REST API, Bearer auth
 *   3. Parent Resolve  — resolveParentFile()         — Drive.Files.get(), mimeType mapping
 *   4. Orchestration   — runFullAudit()              — batch writes, 5-min guard, pause/resume
 *
 * Companion files expected in the same clasp project:
 *   Utils.js    — retryWithBackoff(fn, maxRetries, baseDelayMs)
 *   Reporter.js — initializeSheet(ss), initializeLogSheet(ss),
 *                 flushBatch(sheet, rows), finalizeSheet(ss)
 *
 * Required OAuth scopes (appsscript.json):
 *   https://www.googleapis.com/auth/drive.readonly
 *   https://www.googleapis.com/auth/spreadsheets
 *   https://www.googleapis.com/auth/script.deployments.readonly
 *   https://www.googleapis.com/auth/script.external_request
 *
 * Required Advanced Service:
 *   Drive API v3 (identifier: Drive)
 *
 * @author  GAS Architect
 * @version 1.0.0
 */

// ---------------------------------------------------------------------------
// CONSTANTS
// ---------------------------------------------------------------------------

/** Maximum execution duration guard in milliseconds (5 minutes out of 6-min cap). */
var MAX_EXEC_MS = 5 * 60 * 1000;

/** Minimum remaining time threshold — pause if below this value (ms). */
var PAUSE_THRESHOLD_MS = 60 * 1000;

/** Milliseconds to sleep between project iterations for quota protection. */
var INTER_PROJECT_SLEEP_MS = 150;

/** Number of rows to accumulate before flushing a batch to the sheet. */
var BATCH_FLUSH_SIZE = 10;

/** PropertiesService keys for pause/resume state. */
var PROP_LAST_INDEX   = 'GAS_AUDIT_LAST_INDEX';
var PROP_SHEET_ID     = 'GAS_AUDIT_SHEET_ID';
var PROP_ALL_PROJECTS = 'GAS_AUDIT_PROJECTS_JSON';

/** Apps Script deployments REST endpoint base. */
var DEPLOYMENTS_BASE = 'https://script.googleapis.com/v1/projects/{scriptId}/deployments';

// ---------------------------------------------------------------------------
// PHASE 1 — DISCOVERY
// ---------------------------------------------------------------------------

/**
 * Discovers all Google Apps Script projects accessible to the running user,
 * including those in shared drives.  Paginates automatically until all pages
 * are exhausted.
 *
 * Uses Drive.Files.list() with a structured params object (NOT a hand-built
 * query string) to satisfy the Apps Script Advanced Drive Service signature.
 *
 * @return {Array<Object>} Flat array of Drive file metadata objects.
 *   Each object contains: id, name, createdTime, modifiedTime, owners,
 *   parents, shared.
 */
function getAllScriptProjects() {
  var allFiles = [];
  var pageToken = null;

  try {
    do {
      var params = {
        q: "mimeType='application/vnd.google-apps.script' and trashed=false",
        fields: 'nextPageToken, files(id, name, createdTime, modifiedTime, owners, parents, shared)',
        includeItemsFromAllDrives: true,
        supportsAllDrives: true,
        corpora: 'allDrives',
        pageSize: 1000
      };

      if (pageToken) {
        params.pageToken = pageToken;
      }

      var response = Drive.Files.list(params);

      if (response && response.files && response.files.length > 0) {
        allFiles = allFiles.concat(response.files);
      }

      pageToken = response ? response.nextPageToken : null;

    } while (pageToken);

    Logger.log('[Discovery] Total script projects found: %s', allFiles.length);
    return allFiles;

  } catch (err) {
    Logger.log('[Discovery] FATAL ERROR during Drive.Files.list(): %s\nStack: %s',
               err.message, err.stack);
    throw new Error('[getAllScriptProjects] ' + err.message);
  }
}

// ---------------------------------------------------------------------------
// PHASE 2 — WEB APP DEPLOYMENTS
// ---------------------------------------------------------------------------

/**
 * Retrieves all non-HEAD web app deployments for a given Apps Script project.
 *
 * Calls the Apps Script REST API with a Bearer token from ScriptApp.getOAuthToken().
 * Filters out the @HEAD pseudo-deployment and only returns entryPoints of
 * type WEB_APP.  Uses retryWithBackoff() from Utils.js for transient failures.
 *
 * @param  {string} scriptId The Apps Script project ID (Drive file ID).
 * @return {Array<Object>} Array of web app deployment descriptors:
 *   {deploymentId, description, webAppUrl, accessLevel, executeAs, versionNumber}
 *   Returns [] when the project has no web app deployments or on error.
 */
function getWebAppDeployments(scriptId) {
  if (!scriptId) {
    Logger.log('[Deployments] Called with empty scriptId — skipping.');
    return [];
  }

  var url = DEPLOYMENTS_BASE.replace('{scriptId}', encodeURIComponent(scriptId));
  var token = ScriptApp.getOAuthToken();

  var fetchOptions = {
    method: 'GET',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  };

  var rawResponse;
  try {
    rawResponse = retryWithBackoff(function () {
      return UrlFetchApp.fetch(url, fetchOptions);
    }, 4, 500);
  } catch (err) {
    Logger.log('[Deployments] retryWithBackoff exhausted for scriptId=%s: %s', scriptId, err.message);
    return [];
  }

  var responseCode = rawResponse.getResponseCode();
  if (responseCode !== 200) {
    Logger.log(
      '[Deployments] Non-200 response for scriptId=%s — HTTP %s: %s',
      scriptId,
      responseCode,
      rawResponse.getContentText().substring(0, 300)
    );
    return [];
  }

  var parsed;
  try {
    parsed = JSON.parse(rawResponse.getContentText());
  } catch (parseErr) {
    Logger.log('[Deployments] JSON parse error for scriptId=%s: %s', scriptId, parseErr.message);
    return [];
  }

  var deployments = parsed.deployments || [];
  var webApps = [];

  for (var i = 0; i < deployments.length; i++) {
    var dep = deployments[i];

    // Filter out the @HEAD pseudo-deployment.
    if (dep.deploymentId === '@HEAD') continue;

    var entryPoints = dep.entryPoints || [];
    for (var j = 0; j < entryPoints.length; j++) {
      var ep = entryPoints[j];
      if (ep.entryPointType !== 'WEB_APP') continue;

      var webAppConfig = ep.webApp || {};
      var deployConfig = dep.deploymentConfig || {};

      webApps.push({
        deploymentId:  dep.deploymentId || '',
        description:   deployConfig.description || '',
        webAppUrl:     webAppConfig.url || '',
        accessLevel:   webAppConfig.access || '',
        executeAs:     webAppConfig.executeAs || '',
        versionNumber: deployConfig.versionNumber || 0
      });
    }
  }

  return webApps;
}

// ---------------------------------------------------------------------------
// PHASE 3 — PARENT RESOLUTION
// ---------------------------------------------------------------------------

/**
 * Resolves a Drive parent ID to human-readable metadata.
 *
 * Fetches file metadata via Drive.Files.get() and maps the mimeType to a
 * friendly type label.  isContainer is true when the parent is a Workspace
 * document (Sheet, Doc, Form, Slides) — i.e. a container script host —
 * rather than a plain folder or a nested script project.
 *
 * @param  {string} parentId The Drive file/folder ID to resolve.
 * @return {Object} Shape:
 *   {parentId, parentName, parentType, parentUrl, isContainer, error?}
 *   On any error returns a safe object with error property set.
 */
function resolveParentFile(parentId) {
  if (!parentId) {
    return {
      parentId:   '',
      parentName: 'Unknown',
      parentType: 'Unknown',
      parentUrl:  '',
      isContainer: false,
      error:      'No parentId provided'
    };
  }

  try {
    var file = Drive.Files.get(parentId, {
      fields: 'id, name, mimeType, webViewLink',
      supportsAllDrives: true
    });

    var mimeType   = file.mimeType || '';
    var parentType = _mimeTypeToLabel(mimeType);
    var isContainer = (
      parentType !== 'Drive Folder' &&
      parentType !== 'Nested Script' &&
      parentType !== 'Unknown'
    );

    return {
      parentId:    file.id   || parentId,
      parentName:  file.name || 'Unnamed',
      parentType:  parentType,
      parentUrl:   file.webViewLink || '',
      isContainer: isContainer
    };

  } catch (err) {
    Logger.log('[ParentResolve] Error resolving parentId=%s: %s', parentId, err.message);
    return {
      parentId:    parentId,
      parentName:  'Error',
      parentType:  'Unknown',
      parentUrl:   '',
      isContainer: false,
      error:       err.message
    };
  }
}

/**
 * Maps a Google Drive MIME type string to a human-readable label.
 *
 * @private
 * @param  {string} mimeType The MIME type string from the Drive API.
 * @return {string} Human-readable type label.
 */
function _mimeTypeToLabel(mimeType) {
  var map = {
    'application/vnd.google-apps.spreadsheet':  'Google Sheet',
    'application/vnd.google-apps.document':     'Google Doc',
    'application/vnd.google-apps.form':         'Google Form',
    'application/vnd.google-apps.presentation': 'Google Slides',
    'application/vnd.google-apps.folder':       'Drive Folder',
    'application/vnd.google-apps.script':       'Nested Script'
  };
  return map[mimeType] || 'Other (' + mimeType + ')';
}

// ---------------------------------------------------------------------------
// PHASE 4 — MASTER ORCHESTRATOR
// ---------------------------------------------------------------------------

/**
 * Runs the full GAS Project Audit end-to-end.
 *
 * Creates a new Google Sheet, discovers all Apps Script projects in the
 * account (including shared drives), then iterates through each project to
 * collect deployment and parent metadata.  Rows are written in batches of
 * BATCH_FLUSH_SIZE via Reporter.flushBatch() to minimise Sheets API calls.
 *
 * Execution is guarded against the Apps Script 6-minute cap: if remaining
 * time drops below PAUSE_THRESHOLD_MS the run is paused and state is saved
 * to PropertiesService.  Call resumeAudit() to continue.
 *
 * On successful completion, PropertiesService state is cleared.
 *
 * Depends on:
 *   Utils.js    — retryWithBackoff()
 *   Reporter.js — initializeSheet(), initializeLogSheet(),
 *                 flushBatch(), finalizeSheet()
 */
function runFullAudit() {
  var execStart   = Date.now();
  var scriptProps = PropertiesService.getScriptProperties();

  // ---- Create output spreadsheet -----------------------------------------
  var timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm'
  );
  var ssName = 'GAS Audit \u2014 ' + timestamp;

  var ss;
  try {
    ss = SpreadsheetApp.create(ssName);
    Logger.log('[Audit] Created spreadsheet: %s | URL: %s', ssName, ss.getUrl());
  } catch (createErr) {
    Logger.log('[Audit] FATAL: Could not create spreadsheet: %s', createErr.message);
    throw createErr;
  }

  var auditSheet = initializeSheet(ss);
  var logSheet   = initializeLogSheet(ss);

  _appendLog(logSheet, 'INFO', 'Audit started — spreadsheet: ' + ss.getUrl());

  // ---- Phase 1: Discovery --------------------------------------------------
  var allProjects;
  try {
    allProjects = getAllScriptProjects();
  } catch (discoverErr) {
    _appendLog(logSheet, 'ERROR', 'Discovery failed: ' + discoverErr.message);
    Logger.log('[Audit] Discovery failed: %s', discoverErr.message);
    return;
  }

  var totalCount = allProjects.length;
  Logger.log('[Audit] Discovered %s script projects.', totalCount);
  _appendLog(logSheet, 'INFO', 'Discovered ' + totalCount + ' script projects.');

  // Persist project list and spreadsheet ID so resumeAudit() can pick up.
  scriptProps.setProperty(PROP_SHEET_ID, ss.getId());
  scriptProps.setProperty(PROP_ALL_PROJECTS, JSON.stringify(allProjects));
  scriptProps.setProperty(PROP_LAST_INDEX, '0');

  // ---- Phase 2-3: Iterate projects ----------------------------------------
  var batchRows  = [];
  var startIndex = 0;

  _processProjects(
    allProjects,
    startIndex,
    auditSheet,
    logSheet,
    batchRows,
    execStart,
    scriptProps,
    ss
  );
}

/**
 * Resumes a paused audit run using state stored in PropertiesService.
 *
 * Reads lastProcessedIndex and spreadsheetId saved by runFullAudit() or a
 * prior resumeAudit() call.  If no state is found logs a warning and exits.
 */
function resumeAudit() {
  var execStart   = Date.now();
  var scriptProps = PropertiesService.getScriptProperties();

  var sheetId      = scriptProps.getProperty(PROP_SHEET_ID);
  var lastIndex    = parseInt(scriptProps.getProperty(PROP_LAST_INDEX) || '0', 10);
  var projectsJson = scriptProps.getProperty(PROP_ALL_PROJECTS);

  if (!sheetId || !projectsJson) {
    Logger.log('[ResumeAudit] No saved audit state found. Run runFullAudit() first.');
    return;
  }

  var allProjects;
  try {
    allProjects = JSON.parse(projectsJson);
  } catch (parseErr) {
    Logger.log('[ResumeAudit] Could not parse saved project list: %s', parseErr.message);
    return;
  }

  var ss;
  try {
    ss = SpreadsheetApp.openById(sheetId);
  } catch (openErr) {
    Logger.log('[ResumeAudit] Could not open spreadsheet id=%s: %s', sheetId, openErr.message);
    return;
  }

  // Locate the audit data sheet (first sheet) and the log sheet by name.
  var auditSheet = ss.getSheets()[0];
  var logSheet   = null;
  var sheets     = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    if (sheets[s].getName() === 'Execution Log') {
      logSheet = sheets[s];
      break;
    }
  }
  if (!logSheet) {
    logSheet = initializeLogSheet(ss);
  }

  Logger.log('[ResumeAudit] Resuming from index %s of %s', lastIndex, allProjects.length);
  _appendLog(logSheet, 'INFO', 'RESUME — starting from index ' + lastIndex + ' of ' + allProjects.length);

  var batchRows = [];
  _processProjects(
    allProjects,
    lastIndex,
    auditSheet,
    logSheet,
    batchRows,
    execStart,
    scriptProps,
    ss
  );
}

/**
 * Core iteration loop shared by runFullAudit() and resumeAudit().
 *
 * Iterates through allProjects starting at startIndex, collecting web app
 * deployments and parent metadata for each project, then writing rows in
 * batches.  Pauses gracefully and saves state when execution time is low.
 *
 * @private
 * @param {Array<Object>} allProjects  Full array of discovered script projects.
 * @param {number}        startIndex   Index to begin processing from (0 = fresh run).
 * @param {Sheet}         auditSheet   The main data output sheet.
 * @param {Sheet}         logSheet     The execution log sheet.
 * @param {Array<Array>}  batchRows    Accumulator array for pending rows.
 * @param {number}        execStart    Date.now() timestamp when execution began.
 * @param {Properties}    scriptProps  Script-level PropertiesService reference.
 * @param {Spreadsheet}   ss           The output spreadsheet.
 */
function _processProjects(
  allProjects, startIndex, auditSheet, logSheet, batchRows, execStart, scriptProps, ss
) {
  var totalCount   = allProjects.length;
  // nextWriteRow tracks the 1-based sheet row for the next batch flush.
  // Row 1 = header; data starts at row 2. We advance it by batchRows.length after each flush.
  var nextWriteRow = auditSheet.getLastRow() + 1;
  if (nextWriteRow < 2) nextWriteRow = 2;

  // Stats counters — passed to finalizeSheet() for the Summary dashboard.
  var stats = {
    total:      0,
    bound:      0,
    standalone: 0,
    withWebApp: 0,
    errors:     0
  };

  for (var i = startIndex; i < totalCount; i++) {

    // ---- Execution time guard --------------------------------------------
    var elapsed   = Date.now() - execStart;
    var remaining = MAX_EXEC_MS - elapsed;

    if (remaining < PAUSE_THRESHOLD_MS) {
      // Flush whatever is in the buffer before pausing.
      if (batchRows.length > 0) {
        try {
          flushBatch(auditSheet, batchRows, nextWriteRow);
          nextWriteRow += batchRows.length;
          batchRows.length = 0; // in-place reset
        } catch (flushErr) {
          Logger.log('[Audit] Flush error before pause: %s', flushErr.message);
        }
      }

      scriptProps.setProperty(PROP_LAST_INDEX, String(i));
      var pauseMsg = 'PAUSED at index ' + i + '/' + totalCount +
                     ' — resume with resumeAudit()';
      Logger.log('[Audit] ' + pauseMsg);
      _appendLog(logSheet, 'PAUSE', pauseMsg);
      return; // Exit without finalizing.
    }

    // ---- Per-project processing ------------------------------------------
    var project  = allProjects[i];
    var scriptId = project.id;
    var name     = project.name || 'Unnamed';

    // Phase 2: Web app deployments.
    var deployments = [];
    try {
      deployments = getWebAppDeployments(scriptId);
    } catch (depErr) {
      Logger.log('[Audit] Deployment error for %s (%s): %s', name, scriptId, depErr.message);
    }

    // Phase 3: Parent resolution.
    var parentInfo = {
      parentId:    '',
      parentName:  'No Parent',
      parentType:  'Standalone',
      parentUrl:   '',
      isContainer: false
    };
    var parents = project.parents;
    if (parents && parents.length > 0) {
      try {
        parentInfo = resolveParentFile(parents[0]);
      } catch (parentErr) {
        Logger.log('[Audit] Parent error for %s (%s): %s', name, scriptId, parentErr.message);
      }
    }

    // ---- Build row(s) ---------------------------------------------------
    // One row per web app deployment; one summary row if no deployments found.
    if (deployments.length === 0) {
      batchRows.push(_buildRow(project, parentInfo, null));
    }
    for (var d = 0; d < deployments.length; d++) {
      batchRows.push(_buildRow(project, parentInfo, deployments[d]));
    }

    // ---- Update stats counters -------------------------------------------
    stats.total++;
    if (parentInfo.isContainer) { stats.bound++; } else { stats.standalone++; }
    if (deployments.length > 0) { stats.withWebApp++; }

    // ---- Batch flush -----------------------------------------------------
    if (batchRows.length >= BATCH_FLUSH_SIZE) {
      try {
        flushBatch(auditSheet, batchRows, nextWriteRow);
        nextWriteRow += batchRows.length;
        batchRows.length = 0; // in-place reset (matches Reporter.flushBatch contract)
      } catch (flushErr) {
        Logger.log('[Audit] Flush error at index %s: %s', i, flushErr.message);
        _appendLog(logSheet, 'ERROR', 'Flush error at index ' + i + ': ' + flushErr.message);
        stats.errors++;
      }
    }

    // ---- Quota protection -----------------------------------------------
    Utilities.sleep(INTER_PROJECT_SLEEP_MS);
  }

  // ---- Final flush -------------------------------------------------------
  if (batchRows.length > 0) {
    try {
      flushBatch(auditSheet, batchRows, nextWriteRow);
    } catch (finalFlushErr) {
      Logger.log('[Audit] Final flush error: %s', finalFlushErr.message);
      _appendLog(logSheet, 'ERROR', 'Final flush error: ' + finalFlushErr.message);
      stats.errors++;
    }
  }

  // ---- Finalize (pass stats for Summary dashboard) -----------------------
  try {
    finalizeSheet(ss, stats);
  } catch (finalizeErr) {
    Logger.log('[Audit] finalizeSheet error: %s', finalizeErr.message);
  }

  var completeMsg = 'Audit complete — ' + totalCount + ' projects processed. ' + ss.getUrl();
  Logger.log('[Audit] ' + completeMsg);
  _appendLog(logSheet, 'COMPLETE', completeMsg);

  // Clear saved state on success.
  scriptProps.deleteProperty(PROP_LAST_INDEX);
  scriptProps.deleteProperty(PROP_SHEET_ID);
  scriptProps.deleteProperty(PROP_ALL_PROJECTS);

  // Show completion alert (only fires in a UI context — silently skipped otherwise).
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      'GAS Audit Complete',
      'All ' + totalCount + ' projects processed.\n\nSheet URL:\n' + ss.getUrl(),
      ui.ButtonSet.OK
    );
  } catch (_) {
    // No UI available (trigger / standalone context) — silently skip.
  }
}

/**
 * Builds a flat row array for a single script project / deployment pairing.
 *
 * Column order MUST match the header row defined in Reporter.initializeSheet().
 * Default 18-column layout:
 *   A  Script ID           B  Script Name         C  Owner Email
 *   D  Created Time        E  Last Modified        F  Shared?
 *   G  Parent ID           H  Parent Name          I  Parent Type
 *   J  Is Container?       K  Parent URL           L  Has Web App?
 *   M  Deployment ID       N  Deployment Desc      O  Web App URL
 *   P  Access Level        Q  Execute As           R  Version Number
 *
 * @private
 * @param  {Object}      project    Drive file metadata object from Phase 1.
 * @param  {Object}      parentInfo Resolved parent metadata from Phase 3.
 * @param  {Object|null} deployment Single web app deployment descriptor, or null.
 * @return {Array}       Flat array of cell values (strings / numbers / booleans).
 */
function _buildRow(project, parentInfo, deployment) {
  var owners     = project.owners || [];
  var ownerEmail = owners.length > 0 ? (owners[0].emailAddress || '') : '';
  var hasWebApp  = !!deployment;

  return [
    project.id           || '',                       // A: Script ID
    project.name         || 'Unnamed',                // B: Script Name
    ownerEmail,                                       // C: Owner Email
    project.createdTime  || '',                       // D: Created Time
    project.modifiedTime || '',                       // E: Last Modified
    project.shared       ? 'Yes' : 'No',              // F: Shared?
    parentInfo.parentId   || '',                      // G: Parent ID
    parentInfo.parentName || '',                      // H: Parent Name
    parentInfo.parentType || '',                      // I: Parent Type
    parentInfo.isContainer ? 'Yes' : 'No',            // J: Is Container?
    parentInfo.parentUrl  || '',                      // K: Parent URL
    hasWebApp ? 'Yes' : 'No',                         // L: Has Web App?
    hasWebApp ? deployment.deploymentId  : '',        // M: Deployment ID
    hasWebApp ? deployment.description   : '',        // N: Deployment Desc
    hasWebApp ? deployment.webAppUrl     : '',        // O: Web App URL
    hasWebApp ? deployment.accessLevel   : '',        // P: Access Level
    hasWebApp ? deployment.executeAs     : '',        // Q: Execute As
    hasWebApp ? deployment.versionNumber : ''         // R: Version Number
  ];
}

/**
 * Appends a timestamped entry to the Execution Log sheet.
 *
 * This private helper avoids a hard dependency on Reporter.js for log writes
 * originating inside Code.js.  Uses appendRow() — acceptable here because
 * log writes are infrequent (not in the tight project loop).
 *
 * @private
 * @param {Sheet}  logSheet  The execution log sheet.  No-op if null/undefined.
 * @param {string} level     Severity label: INFO | PAUSE | COMPLETE | ERROR.
 * @param {string} message   Human-readable log message.
 */
function _appendLog(logSheet, level, message) {
  if (!logSheet) return;
  try {
    var ts = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
    logSheet.appendRow([ts, level, message]);
  } catch (logErr) {
    Logger.log('[AppendLog] Could not write to log sheet: %s', logErr.message);
  }
}

// ---------------------------------------------------------------------------
// VALIDATION HELPERS
// ---------------------------------------------------------------------------

/**
 * Runs all three discovery phases against a single Script ID and logs every
 * result object to Logger.  Use this before a full audit to verify that
 * authentication, Drive API access, and Apps Script API access are all working.
 *
 * @param {string} scriptId A valid Apps Script project ID (Google Drive file ID).
 * @example
 *   // In the Apps Script editor — Run > testSingleProject
 *   // Or call with a hard-coded ID:
 *   function myTest() { testSingleProject('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms'); }
 */
function testSingleProject(scriptId) {
  if (!scriptId) {
    Logger.log('[TestSingle] No scriptId provided. Pass a valid Script ID as the argument.');
    return;
  }

  Logger.log('=== testSingleProject: %s ===', scriptId);

  // Phase 2 — Deployments
  Logger.log('--- Phase 2: Web App Deployments ---');
  var deployments;
  try {
    deployments = getWebAppDeployments(scriptId);
    Logger.log('Deployments found: %s', deployments.length);
    for (var d = 0; d < deployments.length; d++) {
      Logger.log('  [%s] %s', d, JSON.stringify(deployments[d]));
    }
  } catch (depErr) {
    Logger.log('  ERROR: %s', depErr.message);
    deployments = [];
  }

  // Phase 3 — Parent Resolution
  // First, fetch the script's own parents list via Drive.Files.get.
  Logger.log('--- Phase 3: Parent Resolution ---');
  try {
    var fileMeta = Drive.Files.get(scriptId, {
      fields: 'id, name, parents',
      supportsAllDrives: true
    });
    var parents = fileMeta.parents || [];
    Logger.log('Script name: %s | Parents: %s', fileMeta.name, JSON.stringify(parents));

    if (parents.length > 0) {
      var parentInfo = resolveParentFile(parents[0]);
      Logger.log('Parent info: %s', JSON.stringify(parentInfo));
    } else {
      Logger.log('Script has no parent — standalone project.');
    }
  } catch (parentErr) {
    Logger.log('  ERROR resolving parent: %s', parentErr.message);
  }

  Logger.log('=== testSingleProject complete ===');
}

/**
 * Executes Phase 1 only and logs every discovered project to Logger.
 *
 * Use this as a quick sanity-check to confirm Drive API access and see the
 * total project inventory before committing to a full audit run.
 *
 * Log format per project:
 *   [index] <id> | <name> | modified=<date> | shared=<Yes|No>
 */
function listAllProjectIds() {
  Logger.log('=== listAllProjectIds — Phase 1 Discovery ===');

  var projects;
  try {
    projects = getAllScriptProjects();
  } catch (err) {
    Logger.log('FATAL: Discovery failed: %s', err.message);
    return;
  }

  Logger.log('Total projects discovered: %s', projects.length);
  Logger.log('--------------------------------------------------');

  for (var i = 0; i < projects.length; i++) {
    var p = projects[i];
    Logger.log(
      '[%s] %s | %s | modified=%s | shared=%s',
      i,
      p.id,
      p.name         || 'Unnamed',
      p.modifiedTime || 'N/A',
      p.shared       ? 'Yes' : 'No'
    );
  }

  Logger.log('=== listAllProjectIds complete ===');
}
