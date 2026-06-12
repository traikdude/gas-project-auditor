/**
 * =============================================================================
 * GAS PROJECT AUDITOR — PREFLIGHT VALIDATION HARNESS
 * File: Tests.js
 * =============================================================================
 * Run `runAllPreflightChecks()` BEFORE executing the full audit.
 * All checks are self-contained, return {pass, message} objects, and log
 * to Logger so you can inspect output in the Execution log panel.
 *
 * Execution flow:
 *   1. checkDriveApiEnabled()        — Drive Advanced Service bound?
 *   2. checkAppsScriptApiEnabled()   — Apps Script REST API reachable?
 *   3. checkOAuthScopes()            — Token covers script.projects.readonly?
 *   4. checkExecutionEnvironment()   — Session, Properties, Sheets all live?
 *   5. testSingleScriptDiscovery()   — Can we actually find GAS files?
 *   6. testDeploymentFetch()         — Deployments endpoint accessible?
 *   7. runDryRun()                   — 3-project simulation without writes
 *
 * Dependencies (must match appsscript.json):
 *   Advanced Services : Drive API v3 (Drive)
 *   OAuth Scopes      : https://www.googleapis.com/auth/drive.readonly
 *                       https://www.googleapis.com/auth/script.projects.readonly
 *                       https://www.googleapis.com/auth/spreadsheets
 *                       https://www.googleapis.com/auth/userinfo.email
 * =============================================================================
 */

// ---------------------------------------------------------------------------
// CONSTANTS (mirror or import from Code.js)
// ---------------------------------------------------------------------------
var TEST_SCRIPT_MIME = "application/vnd.google-apps.script";
var TEST_GAS_API_BASE = "https://script.googleapis.com/v1";
var TEST_TIMEOUT_MS = 10000; // 10s UrlFetch timeout for preflight
var TEST_MAX_SAMPLE = 3;     // how many projects to pull in discovery test

// ---------------------------------------------------------------------------
// 1. ORCHESTRATOR
// ---------------------------------------------------------------------------

/**
 * runAllPreflightChecks()
 * -----------------------
 * Runs every check in sequence. Logs a pass/fail summary table and shows
 * a SpreadsheetApp toast. Returns true only when every check passes.
 *
 * Trigger from the Apps Script editor: Run -> runAllPreflightChecks
 *
 * @return {boolean} true if ALL checks passed, false otherwise
 */
function runAllPreflightChecks() {
  var start = Date.now();
  Logger.log("=================================================================");
  Logger.log("  GAS Project Auditor -- Preflight Validation Harness");
  Logger.log("  Started: " + new Date().toISOString());
  Logger.log("=================================================================");

  var results = [];

  // -- Check 1: Drive Advanced Service --------------------------------------
  Logger.log("\n[1/6] Checking Drive API Advanced Service...");
  var r1 = checkDriveApiEnabled();
  results.push({ name: "Drive API Enabled", pass: r1.pass, message: r1.message });
  _logResult(r1);

  // -- Check 2: Apps Script REST API ----------------------------------------
  Logger.log("\n[2/6] Checking Apps Script API reachability...");
  var r2 = checkAppsScriptApiEnabled();
  results.push({ name: "Apps Script API Reachable", pass: r2.pass, message: r2.message + " (HTTP " + r2.responseCode + ")" });
  _logResult(r2);

  // -- Check 3: OAuth Scopes ------------------------------------------------
  Logger.log("\n[3/6] Checking OAuth token + scope coverage...");
  var r3 = checkOAuthScopes();
  results.push({ name: "OAuth Scopes", pass: r3.pass, message: r3.message });
  _logResult(r3);

  // -- Check 4: Execution Environment ---------------------------------------
  Logger.log("\n[4/6] Checking execution environment...");
  var r4 = checkExecutionEnvironment();
  results.push({ name: "Execution Environment", pass: r4.pass, message: r4.message });
  _logResult(r4);

  // -- Check 5: Script Discovery --------------------------------------------
  Logger.log("\n[5/6] Testing single script discovery (first " + TEST_MAX_SAMPLE + " results)...");
  var r5 = testSingleScriptDiscovery();
  results.push({ name: "Script Discovery", pass: r5.pass, message: "Found " + r5.count + " scripts. Samples: " + JSON.stringify(r5.samples) });
  _logResult({ pass: r5.pass, message: "Discovered " + r5.count + " script(s) in sample." });

  // -- Check 6: Deployment Fetch --------------------------------------------
  var sampleId = (r5.samples && r5.samples.length > 0) ? r5.samples[0].id : null;
  Logger.log("\n[6/6] Testing deployment fetch" + (sampleId ? " on: " + sampleId : " (no sample ID, skipping)") + "...");
  var r6 = testDeploymentFetch(sampleId);
  results.push({ name: "Deployment Fetch", pass: r6.pass, message: r6.message });
  _logResult(r6);

  // -- Summary Table --------------------------------------------------------
  var allPassed = results.every(function(r) { return r.pass; });
  var elapsed = ((Date.now() - start) / 1000).toFixed(2);

  Logger.log("\n=================================================================");
  Logger.log("  PREFLIGHT SUMMARY  (" + elapsed + "s)");
  Logger.log("=================================================================");
  Logger.log(_padRight("CHECK", 30) + " | " + _padRight("STATUS", 8) + " | DETAILS");
  Logger.log(_repeat("-", 90));
  results.forEach(function(r) {
    Logger.log(
      _padRight(r.name, 30) + " | " +
      _padRight(r.pass ? "PASS" : "FAIL", 8) + " | " +
      r.message
    );
  });
  Logger.log(_repeat("-", 90));
  Logger.log("OVERALL: " + (allPassed ? "ALL CHECKS PASSED -- safe to run full audit." : "ONE OR MORE CHECKS FAILED -- fix issues before auditing."));
  Logger.log("=================================================================\n");

  // -- Toast ----------------------------------------------------------------
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      ss.toast(
        allPassed
          ? "All " + results.length + " preflight checks passed in " + elapsed + "s."
          : results.filter(function(r) { return !r.pass; }).length + " check(s) failed. See Logger.",
        "GAS Auditor Preflight",
        8
      );
    }
  } catch (toastErr) {
    Logger.log("(Toast skipped -- no active spreadsheet bound)");
  }

  return allPassed;
}


// ---------------------------------------------------------------------------
// 2. INDIVIDUAL CHECKS
// ---------------------------------------------------------------------------

/**
 * checkDriveApiEnabled()
 * ----------------------
 * Attempts a minimal Drive.Files.list() call using the Advanced Service.
 * If Drive is not enabled in Services, this throws immediately.
 *
 * @return {{ pass: boolean, message: string }}
 */
function checkDriveApiEnabled() {
  try {
    if (typeof Drive === "undefined" || typeof Drive.Files === "undefined") {
      return {
        pass: false,
        message: "Drive Advanced Service is NOT enabled. Go to Services -> Add Drive API."
      };
    }

    var result = Drive.Files.list({
      pageSize: 1,
      q: "mimeType='" + TEST_SCRIPT_MIME + "'",
      fields: "files(id,name)"
    });

    return {
      pass: true,
      message: "Drive.Files.list() succeeded. Files returned in sample: " +
               (result.files ? result.files.length : 0) + "."
    };
  } catch (e) {
    return {
      pass: false,
      message: "Drive.Files.list() threw: " + e.message +
               ". Ensure Drive API v3 is added under Services and drive.readonly scope is in manifest."
    };
  }
}


/**
 * checkAppsScriptApiEnabled()
 * ---------------------------
 * Hits the Apps Script REST API projects.list endpoint with pageSize=1
 * to confirm the API is accessible and the token carries the right scope.
 *
 * @return {{ pass: boolean, message: string, responseCode: number }}
 */
function checkAppsScriptApiEnabled() {
  var responseCode = -1;
  try {
    var token = ScriptApp.getOAuthToken();
    if (!token) {
      return {
        pass: false,
        message: "ScriptApp.getOAuthToken() returned null. OAuth not configured.",
        responseCode: -1
      };
    }

    var url = TEST_GAS_API_BASE + "/projects?pageSize=1";
    var response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true,
      followRedirects: true
    });

    responseCode = response.getResponseCode();

    if (responseCode === 200) {
      var body = JSON.parse(response.getContentText());
      return {
        pass: true,
        message: "Apps Script API responded 200 OK. Projects in sample: " +
                 (body.projects ? body.projects.length : 0) + ".",
        responseCode: 200
      };
    }

    if (responseCode === 403) {
      return {
        pass: false,
        message: "403 Forbidden -- Apps Script API may not be enabled in Google Cloud Console, " +
                 "or the OAuth token lacks script.projects.readonly scope.",
        responseCode: 403
      };
    }

    if (responseCode === 401) {
      return {
        pass: false,
        message: "401 Unauthorized -- token is invalid or expired. " +
                 "Ensure appsscript.json includes the script.projects.readonly scope.",
        responseCode: 401
      };
    }

    return {
      pass: false,
      message: "Unexpected HTTP " + responseCode + ". Body: " +
               response.getContentText().substring(0, 300),
      responseCode: responseCode
    };

  } catch (e) {
    return {
      pass: false,
      message: "UrlFetchApp.fetch threw: " + e.message,
      responseCode: responseCode
    };
  }
}


/**
 * checkOAuthScopes()
 * ------------------
 * Confirms the token is non-null AND actually works against a
 * scope-gated endpoint (script.projects.readonly). A valid token that
 * lacks the scope will produce a 403 on the API call even if it is
 * non-null -- this double-checks both conditions.
 *
 * @return {{ pass: boolean, message: string }}
 */
function checkOAuthScopes() {
  try {
    var token = ScriptApp.getOAuthToken();
    if (!token) {
      return { pass: false, message: "Token is null -- OAuth not initialised." };
    }

    var tokenInfoUrl = "https://oauth2.googleapis.com/tokeninfo?access_token=" +
                       encodeURIComponent(token);
    var resp = UrlFetchApp.fetch(tokenInfoUrl, { muteHttpExceptions: true });
    var code = resp.getResponseCode();

    if (code !== 200) {
      return {
        pass: false,
        message: "Token validation failed -- tokeninfo returned HTTP " + code +
                 ". Token may be expired or invalid."
      };
    }

    var info = JSON.parse(resp.getContentText());
    var scope = info.scope || "";
    var requiredScopes = [
      "https://www.googleapis.com/auth/drive.readonly",
      "https://www.googleapis.com/auth/script.projects.readonly"
    ];

    var missing = requiredScopes.filter(function(s) {
      return scope.indexOf(s) === -1;
    });

    if (missing.length > 0) {
      return {
        pass: false,
        message: "Token is valid but MISSING required scope(s): " + missing.join(", ") +
                 ". Add them to appsscript.json oauthScopes."
      };
    }

    return {
      pass: true,
      message: "Token valid. All required scopes confirmed: " + requiredScopes.join(", ") + "."
    };

  } catch (e) {
    return { pass: false, message: "Scope check threw: " + e.message };
  }
}


/**
 * checkExecutionEnvironment()
 * ---------------------------
 * Verifies:
 *   1. Session.getActiveUser().getEmail() -> non-empty
 *   2. PropertiesService.getScriptProperties() -> accessible
 *   3. SpreadsheetApp.create() can be called (confirms Sheets scope;
 *      the sheet is immediately deleted to avoid clutter)
 *
 * @return {{ pass: boolean, message: string }}
 */
function checkExecutionEnvironment() {
  var issues = [];
  var details = [];

  // 1. Active User
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email || email.trim() === "") {
      issues.push("Session.getActiveUser().getEmail() returned empty.");
    } else {
      details.push("User: " + email);
    }
  } catch (e) {
    issues.push("Session.getActiveUser() threw: " + e.message);
  }

  // 2. Script Properties
  try {
    var props = PropertiesService.getScriptProperties();
    var testKey = "_preflight_test_" + Date.now();
    props.setProperty(testKey, "ok");
    var readBack = props.getProperty(testKey);
    props.deleteProperty(testKey);
    if (readBack !== "ok") {
      issues.push("PropertiesService round-trip mismatch.");
    } else {
      details.push("PropertiesService: OK");
    }
  } catch (e) {
    issues.push("PropertiesService threw: " + e.message);
  }

  // 3. Spreadsheet Create/Delete (Sheets scope)
  var tempId = null;
  try {
    var tempName = '__GAS_Auditor_Preflight_' + Date.now() + '__';
    var tempSheet = SpreadsheetApp.create(tempName);
    tempId = tempSheet.getId();
    var cleanedUp = false;

    // Try permanent delete first (Drive.Files.remove bypasses Trash)
    try {
      Drive.Files.remove(tempId);
      cleanedUp = true;
      details.push('SpreadsheetApp.create() + cleanup: OK (permanently deleted)');
    } catch (removeErr) {
      // Fallback: move to Trash — better than leaving it in root
      try {
        Drive.Files.update({ trashed: true }, tempId);
        cleanedUp = true;
        details.push('SpreadsheetApp.create() + cleanup: OK (moved to Trash)');
      } catch (trashErr) {
        details.push('SpreadsheetApp.create(): OK (WARNING: cleanup failed — temp sheet "' + tempName + '" may remain in Drive. Delete it manually.)');
      }
    }
  } catch (e) {
    issues.push('SpreadsheetApp.create() threw: ' + e.message +
                ' -- ensure https://www.googleapis.com/auth/spreadsheets scope is in manifest.');
  }

  var pass = issues.length === 0;
  return {
    pass: pass,
    message: pass
      ? details.join(" | ")
      : "ISSUES: " + issues.join("; ") + " | OK: " + details.join(", ")
  };
}


// ---------------------------------------------------------------------------
// 3. DISCOVERY + DEPLOYMENT TESTS
// ---------------------------------------------------------------------------

/**
 * testSingleScriptDiscovery()
 * ---------------------------
 * Calls Drive.Files.list() to find up to TEST_MAX_SAMPLE script projects.
 * Returns names and IDs in `samples` array. This is a canary test --
 * if Drive API works in checkDriveApiEnabled() but this returns 0 results,
 * the account likely has only container-bound scripts (see LIMITATIONS.md).
 *
 * @return {{ pass: boolean, count: number, samples: Array<{id:string, name:string}> }}
 */
function testSingleScriptDiscovery() {
  try {
    var result = Drive.Files.list({
      pageSize: TEST_MAX_SAMPLE,
      q: "mimeType='" + TEST_SCRIPT_MIME + "' and trashed=false",
      fields: "files(id,name,parents,createdTime)",
      orderBy: "modifiedTime desc"
    });

    var files = result.files || [];
    var samples = files.map(function(f) {
      return {
        id: f.id,
        name: f.name,
        parents: f.parents || [],
        created: f.createdTime || "unknown"
      };
    });

    if (samples.length === 0) {
      Logger.log(
        "WARNING: Drive search returned 0 standalone scripts.\n" +
        "   This can mean:\n" +
        "   (a) Your account has ONLY container-bound scripts (bound to Sheets/Docs/Forms)\n" +
        "   (b) The drive.readonly scope is missing\n" +
        "   (c) All scripts are in Shared Drives (need allDrives params)\n" +
        "   See LIMITATIONS.md -- Container-Bound Script Discovery Gap."
      );
    } else {
      Logger.log("Sample scripts found:");
      samples.forEach(function(s, i) {
        Logger.log(
          "  [" + (i + 1) + "] " + s.name +
          "\n       ID: " + s.id +
          "\n       Parents: " + (s.parents.length > 0 ? s.parents.join(", ") : "WARNING: NONE (orphaned)") +
          "\n       Created: " + s.created
        );
      });
    }

    return { pass: true, count: samples.length, samples: samples };

  } catch (e) {
    Logger.log("testSingleScriptDiscovery ERROR: " + e.message);
    return { pass: false, count: 0, samples: [], message: e.message };
  }
}


/**
 * testDeploymentFetch(scriptId)
 * -----------------------------
 * Calls the Apps Script API deployments.list endpoint for the given
 * scriptId. Logs the full deployment list. If scriptId is null/undefined,
 * returns a graceful skip result.
 *
 * Surfaces the @HEAD deployment filtering issue:
 *   deployments where entryPoints is empty or description === "@HEAD"
 *   are not real web app deployments and should be excluded.
 *
 * @param  {string|null} scriptId  -- Apps Script project ID to query
 * @return {{ pass: boolean, message: string }}
 */
function testDeploymentFetch(scriptId) {
  if (!scriptId) {
    return {
      pass: true,
      message: "SKIPPED -- no scriptId provided. Pass a valid project ID to test deployment fetch."
    };
  }

  try {
    var token = ScriptApp.getOAuthToken();
    var url = TEST_GAS_API_BASE + "/projects/" + scriptId + "/deployments?pageSize=50";

    var response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();

    if (code === 404) {
      return {
        pass: false,
        message: "404 -- scriptId '" + scriptId + "' not found or you lack access. " +
                 "Ensure the Apps Script project exists and the token has script.projects.readonly."
      };
    }

    if (code === 403) {
      return {
        pass: false,
        message: "403 Forbidden -- Apps Script API not enabled for this project's Cloud Console, " +
                 "or token lacks script.deployments.readonly scope."
      };
    }

    if (code !== 200) {
      return {
        pass: false,
        message: "Unexpected HTTP " + code + " from deployments endpoint. Body: " +
                 response.getContentText().substring(0, 200)
      };
    }

    var body = JSON.parse(response.getContentText());
    var deployments = body.deployments || [];

    Logger.log("\nDeployments for script ID: " + scriptId);
    Logger.log(_repeat("-", 70));

    var webApps = 0;
    var headCount = 0;

    deployments.forEach(function(d, i) {
      var isHead = (d.deploymentConfig && d.deploymentConfig.description === "@HEAD") ||
                   !d.entryPoints || d.entryPoints.length === 0;
      if (isHead) {
        headCount++;
      } else {
        webApps++;
      }

      Logger.log(
        "[" + (i + 1) + "] " + (d.deploymentConfig ? d.deploymentConfig.description : "(no description)") +
        "\n    deploymentId : " + d.deploymentId +
        "\n    versionNumber: " + (d.deploymentConfig ? d.deploymentConfig.versionNumber : "?") +
        "\n    isHead       : " + isHead +
        "\n    entryPoints  : " + (d.entryPoints ? JSON.stringify(d.entryPoints).substring(0, 120) : "none")
      );
    });

    Logger.log(_repeat("-", 70));
    Logger.log(
      "Total deployments: " + deployments.length +
      " | @HEAD (filtered): " + headCount +
      " | Real web apps: " + webApps
    );

    return {
      pass: true,
      message: "Deployments endpoint OK. Total: " + deployments.length +
               ", @HEAD entries: " + headCount +
               ", real deployments: " + webApps + "."
    };

  } catch (e) {
    return { pass: false, message: "testDeploymentFetch threw: " + e.message };
  }
}


// ---------------------------------------------------------------------------
// 4. DRY RUN
// ---------------------------------------------------------------------------

/**
 * runDryRun()
 * -----------
 * Processes the first TEST_MAX_SAMPLE projects through all audit phases
 * WITHOUT writing anything to a Google Sheet. Produces a formatted
 * preview table in Logger showing exactly what the full audit would output.
 *
 * Phases simulated:
 *   1. Discover projects (Drive.Files.list)
 *   2. Fetch deployment list per project (Apps Script API)
 *   3. Resolve parent folder name (Drive.Files.get)
 *   4. Format output row
 *
 * @return {{ pass: boolean, projectsProcessed: number, elapsedMs: number }}
 */
function runDryRun() {
  var start = Date.now();
  Logger.log("\n" + _repeat("=", 70));
  Logger.log("  DRY RUN -- First " + TEST_MAX_SAMPLE + " Projects (no writes)");
  Logger.log(_repeat("=", 70));

  var discovery = testSingleScriptDiscovery();
  if (!discovery.pass || discovery.count === 0) {
    Logger.log("ABORT: Dry run aborted -- script discovery failed or returned 0 results.");
    return { pass: false, projectsProcessed: 0, elapsedMs: Date.now() - start };
  }

  var token = ScriptApp.getOAuthToken();
  var rows = [];

  discovery.samples.forEach(function(script, idx) {
    Logger.log("\nProcessing [" + (idx + 1) + "/" + discovery.samples.length + "]: " + script.name);

    // Phase 1: Resolve parent folder
    var parentName = "(unknown)";
    var parentId = "(none)";
    if (script.parents && script.parents.length > 0) {
      parentId = script.parents[0];
      try {
        var parentFile = Drive.Files.get(parentId, { fields: "id,name" });
        parentName = parentFile.name;
      } catch (parentErr) {
        parentName = "404 (orphaned or no access)";
      }
    } else {
      parentName = "No parents (orphaned script)";
    }

    // Phase 2: Fetch deployments
    var deploymentSummary = "(none)";
    var webAppUrl = "(none)";
    var versionNumber = "-";
    var accessType = "-";
    var headCount = 0;
    var realDepCount = 0;

    try {
      var url = TEST_GAS_API_BASE + "/projects/" + script.id + "/deployments?pageSize=50";
      var resp = UrlFetchApp.fetch(url, {
        method: "GET",
        headers: { Authorization: "Bearer " + token },
        muteHttpExceptions: true
      });

      if (resp.getResponseCode() === 200) {
        var body = JSON.parse(resp.getContentText());
        var deps = body.deployments || [];

        deps.forEach(function(d) {
          var isHead = (d.deploymentConfig && d.deploymentConfig.description === "@HEAD") ||
                       !d.entryPoints || d.entryPoints.length === 0;
          if (isHead) {
            headCount++;
            return;
          }
          realDepCount++;
          if (d.entryPoints && d.entryPoints.length > 0) {
            var ep = d.entryPoints[0];
            if (ep.webApp) {
              webAppUrl = ep.webApp.url || "(no URL)";
              accessType = ep.webApp.access || "UNKNOWN";
            }
          }
          versionNumber = (d.deploymentConfig && d.deploymentConfig.versionNumber != null)
            ? String(d.deploymentConfig.versionNumber) : "?";
          deploymentSummary = (d.deploymentConfig && d.deploymentConfig.description)
            ? d.deploymentConfig.description : "(no description)";
        });

        if (realDepCount === 0) {
          deploymentSummary = "No deployments (only @HEAD)";
        }
      } else {
        deploymentSummary = "HTTP " + resp.getResponseCode() + " error";
      }
    } catch (depErr) {
      deploymentSummary = "ERROR: " + depErr.message;
    }

    // Phase 3: Accumulate preview row
    rows.push({
      name: script.name,
      id: script.id,
      created: script.created,
      parentId: parentId,
      parentName: parentName,
      deploymentSummary: deploymentSummary,
      versionNumber: versionNumber,
      webAppUrl: webAppUrl,
      accessType: accessType,
      realDeployments: realDepCount,
      headDeployments: headCount
    });
  });

  // Render Preview Table
  var elapsed = ((Date.now() - start) / 1000).toFixed(2);
  Logger.log("\n" + _repeat("=", 70));
  Logger.log("  DRY RUN PREVIEW TABLE  (" + rows.length + " projects, " + elapsed + "s)");
  Logger.log(_repeat("=", 70));

  var colW = [28, 18, 22, 14, 10, 6];
  var header =
    _padRight("PROJECT NAME", colW[0]) + " | " +
    _padRight("PARENT FOLDER", colW[1]) + " | " +
    _padRight("DEPLOYMENT", colW[2]) + " | " +
    _padRight("WEB APP URL", colW[3]) + " | " +
    _padRight("ACCESS", colW[4]) + " | " +
    _padRight("VER", colW[5]);
  Logger.log(header);
  Logger.log(_repeat("-", 110));

  rows.forEach(function(r) {
    Logger.log(
      _padRight(_truncate(r.name, colW[0] - 1), colW[0]) + " | " +
      _padRight(_truncate(r.parentName, colW[1] - 1), colW[1]) + " | " +
      _padRight(_truncate(r.deploymentSummary, colW[2] - 1), colW[2]) + " | " +
      _padRight(_truncate(r.webAppUrl, colW[3] - 1), colW[3]) + " | " +
      _padRight(_truncate(r.accessType, colW[4] - 1), colW[4]) + " | " +
      _padRight(r.versionNumber, colW[5])
    );
  });

  Logger.log(_repeat("-", 110));
  Logger.log("Dry run complete. " + rows.length + " project(s) processed in " + elapsed + "s.");
  Logger.log("No data was written to any spreadsheet.\n");

  return { pass: true, projectsProcessed: rows.length, elapsedMs: Date.now() - start };
}


// ---------------------------------------------------------------------------
// PRIVATE UTILITY HELPERS
// ---------------------------------------------------------------------------

/** Log a {pass, message} result object with a visual indicator. */
function _logResult(result) {
  Logger.log((result.pass ? "  PASS -- " : "  FAIL -- ") + result.message);
}

/** Right-pad a string to width characters. */
function _padRight(str, width) {
  str = String(str || "");
  while (str.length < width) { str += " "; }
  return str;
}

/** Repeat a character n times. */
function _repeat(char, n) {
  var s = "";
  for (var i = 0; i < n; i++) { s += char; }
  return s;
}

/** Truncate a string to maxLen characters. */
function _truncate(str, maxLen) {
  str = String(str || "");
  return str.length > maxLen ? str.substring(0, maxLen - 1) + "~" : str;
}
