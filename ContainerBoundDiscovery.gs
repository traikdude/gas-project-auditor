/**
 * @fileoverview ContainerBoundDiscovery.gs
 * Enhanced discovery of container-bound Google Apps Scripts.
 *
 * The standard Drive API search for mimeType='application/vnd.google-apps.script'
 * only returns STANDALONE scripts. Scripts bound to Sheets, Docs, Forms, or Slides
 * are often completely invisible to that search. This module uses a 3-method sweep
 * to surface those hidden scripts and merge them with the Drive MIME results.
 *
 * Public entry points:
 *   DISCOVERY_findContainerBoundScripts() — returns discovered script objects
 *   DISCOVERY_auditAllSources()           — full audit + Sheet output
 *
 * @author  GAS Project Auditor
 * @version 1.0.0
 */

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** Base URL for the Apps Script REST API. */
var DISCOVERY_SCRIPT_API_BASE = 'https://script.googleapis.com/v1';

/** Maximum files to fetch per Drive page (hard cap 1 000; 100 is safe). */
var DISCOVERY_DRIVE_PAGE_SIZE = 100;

/** Maximum projects per Apps Script API page. */
var DISCOVERY_API_PAGE_SIZE = 50;

/** Milliseconds to sleep between per-file container checks (rate-limit guard). */
var DISCOVERY_SLEEP_MS = 100;

// ---------------------------------------------------------------------------
// Public: DISCOVERY_findContainerBoundScripts
// ---------------------------------------------------------------------------

/**
 * Orchestrates a 3-method sweep to discover container-bound scripts that are
 * invisible to the standard Drive MIME type search.
 *
 * Methods used:
 *  1. Apps Script REST API projects.list  (broadest, may require extra OAuth)
 *  2. Drive scan of Sheets/Docs/Forms/Slides → check each for a bound script
 *  3. Deduplication merge of both result sets
 *
 * @return {Array.<{scriptId: string,
 *                  title: string,
 *                  parentId: string,
 *                  parentTitle: string,
 *                  parentMimeType: string,
 *                  discoveryMethod: string}>}
 *   Merged, deduplicated array of discovered container-bound script objects.
 */
function DISCOVERY_findContainerBoundScripts() {
  try {
    Logger.log('[ContainerDiscovery] Starting 3-method container-bound script sweep...');

    // Method 1 — Apps Script REST API projects.list
    var apiResults = DISCOVERY_tryProjectsList_();
    Logger.log('[ContainerDiscovery] projects.list returned ' + apiResults.length + ' script(s).');

    // Method 2 — Scan container files in Drive
    var scanResults = DISCOVERY_scanContainerFiles_();
    Logger.log('[ContainerDiscovery] Container file scan returned ' + scanResults.length + ' script(s).');

    // Method 3 — Merge and deduplicate
    var merged = DISCOVERY_mergeAndDedup_(apiResults, scanResults);

    Logger.log('[ContainerDiscovery] Container-bound discovery found ' + merged.length + ' scripts.');
    return merged;

  } catch (e) {
    Logger.log('[ContainerDiscovery] DISCOVERY_findContainerBoundScripts() ERROR: ' + e.message);
    return [];
  }
}

// ---------------------------------------------------------------------------
// Private: DISCOVERY_tryProjectsList_
// ---------------------------------------------------------------------------

/**
 * Attempts to list all Apps Script projects via the REST API projects.list
 * endpoint. Handles pagination via nextPageToken. Fails gracefully — if the
 * endpoint is unavailable (403/404) this returns an empty array and logs the
 * reason so the caller can fall back to the Drive scan.
 *
 * Endpoint: GET https://script.googleapis.com/v1/projects?pageSize=50
 *
 * @return {Array.<{scriptId: string,
 *                  title: string,
 *                  parentId: string,
 *                  discoveryMethod: string}>}
 *   Projects returned by the API, or [] on failure.
 * @private
 */
function DISCOVERY_tryProjectsList_() {
  var results = [];
  var pageToken = null;
  var token = ScriptApp.getOAuthToken();

  try {
    do {
      var url = DISCOVERY_SCRIPT_API_BASE + '/projects?pageSize=' + DISCOVERY_API_PAGE_SIZE;
      if (pageToken) {
        url += '&pageToken=' + encodeURIComponent(pageToken);
      }

      var options = {
        method: 'GET',
        headers: {
          'Authorization': 'Bearer ' + token,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };

      var response = UrlFetchApp.fetch(url, options);
      var code = response.getResponseCode();

      if (code !== 200) {
        Logger.log('[ContainerDiscovery] projects.list returned HTTP ' + code +
                   ' — endpoint not available, falling back.');
        return [];
      }

      var payload = JSON.parse(response.getContentText());
      var projects = payload.projects || [];

      for (var i = 0; i < projects.length; i++) {
        var p = projects[i];
        results.push({
          scriptId: p.scriptId || '',
          title: p.title || '(Untitled)',
          parentId: (p.parentId && p.parentId.fileId) ? p.parentId.fileId : '',
          parentTitle: '',
          parentMimeType: '',
          discoveryMethod: 'api_list'
        });
      }

      pageToken = payload.nextPageToken || null;

    } while (pageToken);

  } catch (e) {
    Logger.log('[ContainerDiscovery] DISCOVERY_tryProjectsList_() ERROR: ' + e.message);
    return [];
  }

  return results;
}

// ---------------------------------------------------------------------------
// Private: DISCOVERY_scanContainerFiles_
// ---------------------------------------------------------------------------

/**
 * Searches Drive for the four container MIME types (Sheets, Docs, Forms,
 * Presentations) and checks each file for a bound Apps Script project.
 *
 * Drive queries use allDrives corpus so Shared Drives are included.
 *
 * @return {Array.<{scriptId: string,
 *                  title: string,
 *                  parentId: string,
 *                  parentTitle: string,
 *                  parentMimeType: string,
 *                  discoveryMethod: string}>}
 *   All bound script objects found across all container files.
 * @private
 */
function DISCOVERY_scanContainerFiles_() {
  var containerMimeTypes = [
    'application/vnd.google-apps.spreadsheet',
    'application/vnd.google-apps.document',
    'application/vnd.google-apps.form',
    'application/vnd.google-apps.presentation'
  ];

  var allContainerFiles = [];

  // Collect files for all 4 MIME types
  for (var m = 0; m < containerMimeTypes.length; m++) {
    var mimeType = containerMimeTypes[m];
    var files = DISCOVERY_listDriveFilesForMime_(mimeType);
    allContainerFiles = allContainerFiles.concat(files);
  }

  Logger.log('[ContainerDiscovery] Scanning ' + allContainerFiles.length +
             ' Sheets/Docs/Forms for bound scripts...');

  var boundScripts = [];

  for (var f = 0; f < allContainerFiles.length; f++) {
    var file = allContainerFiles[f];
    var found = DISCOVERY_checkFileForBoundScript_(file);
    if (found.length > 0) {
      boundScripts = boundScripts.concat(found);
    }
    // Rate-limit guard between per-file API calls
    Utilities.sleep(DISCOVERY_SLEEP_MS);
  }

  return boundScripts;
}

// ---------------------------------------------------------------------------
// Private: DISCOVERY_listDriveFilesForMime_
// ---------------------------------------------------------------------------

/**
 * Returns all Drive files matching a given MIME type, including items from
 * Shared Drives. Handles pagination automatically.
 *
 * @param {string} mimeType  The Google Drive MIME type to search for.
 * @return {Array.<{id: string, name: string, mimeType: string}>}
 *   Files found in Drive matching the MIME type.
 * @private
 */
function DISCOVERY_listDriveFilesForMime_(mimeType) {
  var files = [];
  var pageToken = null;

  try {
    do {
      var params = {
        q: "mimeType='" + mimeType + "' and trashed=false",
        includeItemsFromAllDrives: true,
        supportsAllDrives: true,
        corpora: 'allDrives',
        pageSize: DISCOVERY_DRIVE_PAGE_SIZE,
        fields: 'nextPageToken, files(id, name, mimeType, owners, parents)'
      };

      if (pageToken) {
        params.pageToken = pageToken;
      }

      var result = Drive.Files.list(params);
      var items = (result && result.files) ? result.files : [];
      files = files.concat(items);
      pageToken = result.nextPageToken || null;

    } while (pageToken);

  } catch (e) {
    Logger.log('[ContainerDiscovery] DISCOVERY_listDriveFilesForMime_(' +
               mimeType + ') ERROR: ' + e.message);
  }

  return files;
}

// ---------------------------------------------------------------------------
// Private: DISCOVERY_checkFileForBoundScript_
// ---------------------------------------------------------------------------

/**
 * Checks whether a single Google file (Sheet, Doc, Form, or Presentation)
 * has a container-bound Apps Script project attached to it.
 *
 * Uses the Apps Script REST API endpoint:
 *   GET https://script.googleapis.com/v1/projects?parent=containers/{fileId}
 *
 * A 404 or empty projects array means no bound script. Errors are caught and
 * logged; an empty array is returned so the caller continues gracefully.
 *
 * NOTE: The caller MUST add Utilities.sleep() between calls to this function
 * to avoid hitting Apps Script API rate limits.
 *
 * @param {{id: string, name: string, mimeType: string}} file
 *   A Drive file object containing at minimum id, name, and mimeType.
 * @return {Array.<{scriptId: string,
 *                  title: string,
 *                  parentId: string,
 *                  parentTitle: string,
 *                  parentMimeType: string,
 *                  discoveryMethod: string}>}
 *   Bound script objects found for this file, or [] if none.
 * @private
 */
function DISCOVERY_checkFileForBoundScript_(file) {
  var results = [];

  try {
    var token = ScriptApp.getOAuthToken();
    var url = DISCOVERY_SCRIPT_API_BASE +
              '/projects?parent=containers/' + encodeURIComponent(file.id) +
              '&pageSize=' + DISCOVERY_API_PAGE_SIZE;

    var options = {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();

    if (code !== 200) {
      // 404 = no bound script on this file; other codes = access/quota issue
      return [];
    }

    var payload = JSON.parse(response.getContentText());
    var projects = payload.projects || [];

    for (var i = 0; i < projects.length; i++) {
      var p = projects[i];
      results.push({
        scriptId: p.scriptId || '',
        title: p.title || '(Untitled)',
        parentId: file.id,
        parentTitle: file.name || '(Untitled)',
        parentMimeType: file.mimeType || '',
        discoveryMethod: 'container_scan'
      });
    }

  } catch (e) {
    Logger.log('[ContainerDiscovery] DISCOVERY_checkFileForBoundScript_(' +
               file.id + ') ERROR: ' + e.message);
  }

  return results;
}

// ---------------------------------------------------------------------------
// Private: DISCOVERY_mergeAndDedup_
// ---------------------------------------------------------------------------

/**
 * Merges two arrays of script objects and removes duplicates based on scriptId.
 * When the same script appears in both arrays, the entry with richer metadata
 * (parentTitle, parentMimeType) is preferred; otherwise the first occurrence wins.
 *
 * Logs a summary line showing how many scripts were newly discovered via the
 * container scan vs. the primary array.
 *
 * @param {Array.<Object>} primaryScripts
 *   First array (e.g. from Apps Script projects.list API). Used as the baseline.
 * @param {Array.<Object>} discoveredScripts
 *   Second array (e.g. from the Drive container file scan).
 * @return {Array.<Object>}
 *   Merged, deduplicated array with all unique script objects.
 * @private
 */
function DISCOVERY_mergeAndDedup_(primaryScripts, discoveredScripts) {
  var seen = {};
  var merged = [];
  var newCount = 0;

  // Index primary scripts first
  for (var i = 0; i < primaryScripts.length; i++) {
    var s = primaryScripts[i];
    if (s.scriptId && !seen[s.scriptId]) {
      seen[s.scriptId] = true;
      merged.push(s);
    }
  }

  // Add discovered scripts that are not already indexed
  for (var j = 0; j < discoveredScripts.length; j++) {
    var d = discoveredScripts[j];
    if (d.scriptId && !seen[d.scriptId]) {
      seen[d.scriptId] = true;
      merged.push(d);
      newCount++;
    }
  }

  Logger.log('[ContainerDiscovery] Merged: ' + merged.length +
             ' unique scripts (' + newCount + ' newly discovered via container scan).');

  return merged;
}

// ---------------------------------------------------------------------------
// Public: DISCOVERY_auditAllSources
// ---------------------------------------------------------------------------

/**
 * Full audit entry point. Combines the standard Drive MIME search results
 * (from getAllScriptProjects() in Code.gs) with container-bound discovery,
 * writes everything to a new Google Sheet, and shows a completion toast.
 *
 * Sheet columns:
 *   Script ID | Title | Parent ID | Parent Title | Parent Type | Discovery Method
 *
 * This function can be run directly from the Apps Script editor or a custom
 * menu item.
 *
 * @return {{total: number,
 *            fromDriveSearch: number,
 *            fromContainerScan: number,
 *            fromApiList: number}}
 *   Summary statistics for the audit run.
 */
function DISCOVERY_auditAllSources() {
  try {
    Logger.log('[ContainerDiscovery] DISCOVERY_auditAllSources() — starting full audit...');

    // ---- Step 1: Standard Drive MIME search --------------------------------
    var driveScripts = [];
    try {
      driveScripts = getAllScriptProjects(); // defined in Code.gs
      Logger.log('[ContainerDiscovery] getAllScriptProjects() returned ' +
                 driveScripts.length + ' script(s) from Drive MIME search.');
    } catch (e) {
      Logger.log('[ContainerDiscovery] getAllScriptProjects() unavailable: ' + e.message);
    }

    // Normalise Drive results to the shared schema if they are plain objects
    var normalisedDrive = [];
    for (var d = 0; d < driveScripts.length; d++) {
      var raw = driveScripts[d];
      normalisedDrive.push({
        scriptId: raw.scriptId || raw.id || '',
        title: raw.title || raw.name || '(Untitled)',
        parentId: raw.parentId || '',
        parentTitle: raw.parentTitle || '',
        parentMimeType: raw.parentMimeType || '',
        discoveryMethod: raw.discoveryMethod || 'drive_mime_search'
      });
    }

    // ---- Step 2: Container-bound discovery ---------------------------------
    var containerScripts = DISCOVERY_findContainerBoundScripts();

    // ---- Step 3: Merge & dedup ---------------------------------------------
    var allScripts = DISCOVERY_mergeAndDedup_(normalisedDrive, containerScripts);

    // ---- Step 4: Compute stats ---------------------------------------------
    var stats = {
      total: allScripts.length,
      fromDriveSearch: 0,
      fromContainerScan: 0,
      fromApiList: 0
    };

    for (var i = 0; i < allScripts.length; i++) {
      var method = allScripts[i].discoveryMethod;
      if (method === 'drive_mime_search') {
        stats.fromDriveSearch++;
      } else if (method === 'container_scan') {
        stats.fromContainerScan++;
      } else if (method === 'api_list') {
        stats.fromApiList++;
      }
    }

    // ---- Step 5: Write to a new Google Sheet --------------------------------
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var sheetName = 'Container Discovery Audit \u2014 ' + today;
    var ss = SpreadsheetApp.create(sheetName);
    var sheet = ss.getActiveSheet();
    sheet.setName('Results');

    // Headers
    var headers = [
      'Script ID',
      'Title',
      'Parent ID',
      'Parent Title',
      'Parent Type',
      'Discovery Method'
    ];
    sheet.appendRow(headers);

    // Freeze header row and bold it
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

    // Data rows
    var rows = [];
    for (var r = 0; r < allScripts.length; r++) {
      var sc = allScripts[r];
      rows.push([
        sc.scriptId || '',
        sc.title || '',
        sc.parentId || '',
        sc.parentTitle || '',
        sc.parentMimeType || '',
        sc.discoveryMethod || ''
      ]);
    }

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    // Auto-resize columns for readability
    sheet.autoResizeColumns(1, headers.length);

    var sheetUrl = ss.getUrl();
    Logger.log('[ContainerDiscovery] Audit sheet created: ' + sheetUrl);
    Logger.log('[ContainerDiscovery] Stats: ' + JSON.stringify(stats));

    // ---- Step 6: Toast notification ----------------------------------------
    SpreadsheetApp.getActiveSpreadsheet &&
      SpreadsheetApp.getActiveSpreadsheet() &&
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Found ' + stats.total + ' scripts total (' +
        stats.fromDriveSearch + ' Drive / ' +
        stats.fromContainerScan + ' container scan / ' +
        stats.fromApiList + ' API list). Sheet: ' + sheetUrl,
        'Container Discovery Audit Complete',
        10
      );

    return stats;

  } catch (e) {
    Logger.log('[ContainerDiscovery] DISCOVERY_auditAllSources() ERROR: ' + e.message);
    return { total: 0, fromDriveSearch: 0, fromContainerScan: 0, fromApiList: 0 };
  }
}
