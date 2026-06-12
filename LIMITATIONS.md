# GAS Project Auditor — Complete Limitations Reference

> **Last Updated:** 2026-06-12
> This document covers every known limitation, gap, and edge case in the
> GAS Project Auditor. Read this **before** interpreting audit output.

---

## Table of Contents

1. [Execution Time Limit (6-Minute Cap)](#1-execution-time-limit)
2. [Container-Bound Script Discovery Gap](#2-container-bound-script-discovery-gap)
3. [The 18 vs 57 Discrepancy Explained](#3-the-18-vs-57-discrepancy-explained)
4. [Shared Drive Scripts](#4-shared-drive-scripts)
5. [API Quota Limits](#5-api-quota-limits)
6. [Deleted / Orphaned Scripts](#6-deleted--orphaned-scripts)
7. [@HEAD Deployment Entries](#7-head-deployment-entries)
8. [Multiple Deployments per Project](#8-multiple-deployments-per-project)
9. [Scripts You Do Not Own](#9-scripts-you-do-not-own)
10. [Large Accounts (>500 Scripts)](#10-large-accounts-500-scripts)
11. [Rate Limiting and Exponential Backoff](#11-rate-limiting-and-exponential-backoff)
12. [appsscript.json Manifest Scope Mismatch](#12-appsscriptjson-manifest-scope-mismatch)

---

## 1. Execution Time Limit

| Property | Value |
|---|---|
| **Hard cap** | 6 minutes (360 seconds) per GAS execution |
| **Applies to** | Any single function run, including triggered runs |
| **Symptom** | `Exceeded maximum execution time` exception; audit output truncated |
| **Safe threshold** | Design for ≤ 5 minutes; leave 60s buffer for sheet writes |

### Why It Matters
Google Apps Script enforces a hard 6-minute execution wall on all consumer
and Workspace accounts. The auditor makes at least **two HTTP calls per
script** (one to the Apps Script API for deployments, one to Drive for
parent folder resolution). At 100 scripts, that is roughly 200+ round trips.
Even at 200ms per call, 100 scripts = ~40 seconds in pure fetch time, which
is fine. At 500 scripts, execution time becomes the critical constraint.

### Workaround: `resumeAudit()` via Time-Driven Trigger

The auditor implements a continuation pattern using Script Properties:

```javascript
// At the top of processNextBatch():
var props = PropertiesService.getScriptProperties();
var lastProcessedIndex = parseInt(props.getProperty('AUDIT_CURSOR') || '0');

// After each batch:
props.setProperty('AUDIT_CURSOR', String(lastProcessedIndex + BATCH_SIZE));

// Time guard — stop if < 90 seconds remain:
var elapsed = Date.now() - START_MS;
if (elapsed > 270000) {          // 4.5 minutes
  _scheduleResume();
  return;
}

function _scheduleResume() {
  ScriptApp.newTrigger('resumeAudit')
    .timeBased()
    .after(30 * 1000)            // 30 seconds later
    .create();
}

function resumeAudit() {
  // Delete the trigger that fired this
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'resumeAudit') {
      ScriptApp.deleteTrigger(t);
    }
  });
  processNextBatch();            // re-enters with cursor restored
}
```

> **Important:** `AUDIT_CURSOR` must be reset to `'0'` before each fresh
> audit run, or the auditor will resume from a stale position.

---

## 2. Container-Bound Script Discovery Gap

| Property | Detail |
|---|---|
| **Affected method** | `Drive.Files.list()` with `mimeType='application/vnd.google-apps.script'` |
| **What IS found** | Standalone script projects (created directly at script.google.com) |
| **What is MISSED** | Scripts bound to Sheets, Docs, Forms, Slides, or Sites |
| **Severity** | HIGH — potentially the majority of an account's scripts are missed |

### Why Container-Bound Scripts Are Invisible to Drive Search

Container-bound scripts are **not independent Drive files**. They are
embedded metadata of the parent container (the Sheet, Doc, etc.). When you
open the Drive file list with MIME type filter `application/vnd.google-apps.script`,
Google only returns files whose *primary* MIME type is that of a script
project — which means standalone scripts only.

The container-bound script shares its Drive file ID with the parent container
but has a different internal script ID. The Drive API returns the parent's
MIME type (e.g., `application/vnd.google-apps.spreadsheet`), not the script type.

### How to Discover Container-Bound Scripts

**Method A — Apps Script API `projects.list`:**
```
GET https://script.googleapis.com/v1/projects
```
This endpoint returns *all* projects the authenticated user has access to,
including container-bound ones. However:
- It requires `script.projects.readonly` OAuth scope
- Pagination uses `pageToken`, not Drive's standard cursor
- It does NOT return deployment URLs directly

**Method B — Enumerate parent files, then check for bound scripts:**
```javascript
// For each Sheet/Doc/Form found in Drive:
var parentId = sheet.id;
var scriptUrl = "https://script.googleapis.com/v1/projects?parentId=" + parentId;
// A container-bound script will have this parentId in its projectId field
```

**Method C — script.google.com dashboard:**
Navigate to https://script.google.com/home — this is the only UI that shows
ALL scripts (standalone + container-bound) in a single list.

### Audit Impact

| Script Type | Drive.Files.list | Apps Script API | Dashboard |
|---|---|---|---|
| Standalone (My Drive) | YES | YES | YES |
| Standalone (Shared Drive) | Only with allDrives | YES | YES |
| Container-Bound | **NO** | YES | YES |
| Container-Bound (Shared Drive) | **NO** | Maybe | YES |

---

## 3. The 18 vs 57 Discrepancy Explained

> **Scenario:** `clasp list` shows **18** scripts. The Apps Script dashboard
> at script.google.com shows **57** scripts. Why?

### Source-by-Source Breakdown

| Source | Count Shown | What It Includes | What It Excludes |
|---|---|---|---|
| `clasp list` | ~18 | Standalone scripts clasprc-linked in local `.clasprc.json`, typically owned by you in My Drive | Container-bound scripts, Shared Drive scripts, scripts not explicitly cloned |
| `script.google.com/home` | ~57 | ALL scripts you can edit — standalone + container-bound, My Drive + Shared Drives | Scripts shared with you that you only have view access to (may or may not appear) |
| `Drive.Files.list` (default) | ~18–30 | Standalone scripts in My Drive you own | Container-bound, Shared Drive, scripts owned by others |
| `Drive.Files.list` (+ allDrives) | ~30–45 | Standalone scripts across all drives | Container-bound scripts |
| Apps Script API `projects.list` | ~50–57 | All projects including container-bound | Scripts you have only view-level access to |

### The Core Reasons for the Gap

**Reason 1 — clasp only tracks standalone scripts**
`clasp list` reads from the Apps Script API `projects.list` but filters to
projects that appear to have a cloned local directory (i.e., have a
`.clasp.json` in a local folder). If you never ran `clasp clone` for a
container-bound script, it will not appear.

**Reason 2 — Container-bound scripts inflate the dashboard count**
Every Sheet, Doc, or Form with an attached script adds an entry in the
dashboard. If your organization uses 39 such files, those 39 container-bound
scripts fill the gap between 18 and 57.

**Reason 3 — Shared Drive scripts**
Scripts in Shared Drives appear in the dashboard but require
`supportsAllDrives=true` + `includeItemsFromAllDrives=true` in Drive API
calls. Without these parameters, they are invisible to the auditor.

**Reason 4 — clasp scope is project-local**
`clasp list` shows projects you have explicitly cloned or created via clasp
on the current machine. It is not a comprehensive account inventory.

### Recommended Fix for the Auditor

Use the Apps Script API `projects.list` as the **primary discovery source**,
then supplement with Drive.Files.list for metadata enrichment:

```javascript
// Primary: Apps Script API (finds everything including container-bound)
var token = ScriptApp.getOAuthToken();
var url = "https://script.googleapis.com/v1/projects?pageSize=50";
// paginate with nextPageToken until exhausted

// Secondary enrichment: Drive.Files.get(scriptId) for folder/parent info
// Note: this will 404 on container-bound scripts — handle gracefully
```

---

## 4. Shared Drive Scripts

| Property | Detail |
|---|---|
| **Required params** | `supportsAllDrives: true`, `includeItemsFromAllDrives: true` |
| **Permission to access** | Must have at least Viewer role on the Shared Drive |
| **Common error** | `403 cannotShareTeamDriveTopFolderWithAnyoneOrDomains` or generic 403 |
| **API behavior** | Without the params, Shared Drive files are silently excluded — no error |

### Required Drive.Files.list Parameters

```javascript
var result = Drive.Files.list({
  pageSize: 100,
  q: "mimeType='application/vnd.google-apps.script' and trashed=false",
  fields: "nextPageToken,files(id,name,parents,createdTime,modifiedTime,owners)",
  orderBy: "modifiedTime desc",
  // REQUIRED for Shared Drive inclusion:
  supportsAllDrives: true,
  includeItemsFromAllDrives: true,
  corpora: "allDrives"           // or "user" for My Drive only
});
```

### Without These Parameters

The Drive API silently returns only My Drive files. No error is thrown.
This is the most common reason for undercount — the `queryString` bug in
the original `getAllScriptProjects()` also caused this: passing a URL
query string as a params object means the extra fields were interpreted as
URL-encoded string keys, which the Drive Advanced Service ignores silently.

### Handling 403 on Shared Drive Items

```javascript
// Wrap per-file metadata calls in a try/catch:
try {
  var meta = Drive.Files.get(fileId, {
    fields: "id,name,parents",
    supportsAllDrives: true
  });
} catch (e) {
  if (e.message.indexOf('403') !== -1) {
    Logger.log('No access to file: ' + fileId + ' -- skipping');
  }
}
```

---

## 5. API Quota Limits

### Drive API (v3)

| Operation | Quota | Notes |
|---|---|---|
| Read requests | 1,000 / 100 seconds / user | `Files.list`, `Files.get` |
| Write requests | 300 / 100 seconds / user | Not used by auditor |
| Daily total | 1,000,000,000 units / day | Unlikely to hit |
| Per-project | 1,000 / 100 seconds | Shared across all services |

### Apps Script API (v1)

| Operation | Quota | Notes |
|---|---|---|
| `projects.list` | 100 / 100 seconds / user | Primary discovery |
| `projects.deployments.list` | 100 / 100 seconds / user | Called per project |
| Daily execution quota | Per-project limit applies | See Cloud Console |

### UrlFetchApp

| Property | Limit |
|---|---|
| Calls per day | 20,000 (consumer) / 100,000 (Workspace) |
| Response size | 50 MB per response |
| Concurrent connections | 10 simultaneous |
| Timeout | 60 seconds per request (not configurable) |

### Retry Logic (Recommended Pattern)

```javascript
function fetchWithRetry(url, options, maxRetries) {
  maxRetries = maxRetries || 3;
  var delay = 1000;  // start at 1 second
  for (var attempt = 0; attempt < maxRetries; attempt++) {
    var resp = UrlFetchApp.fetch(url, options);
    var code = resp.getResponseCode();
    if (code === 200) return resp;
    if (code === 429 || code === 503) {
      // Respect Retry-After header if present
      var retryAfter = parseInt(resp.getHeaders()['Retry-After'] || '0') * 1000;
      Utilities.sleep(Math.max(retryAfter, delay));
      delay *= 2;  // exponential backoff
      continue;
    }
    // Non-retryable error
    throw new Error('HTTP ' + code + ': ' + resp.getContentText().substring(0, 200));
  }
  throw new Error('Max retries exceeded for: ' + url);
}
```

> **Note:** The original spec's `getAllScriptProjects()` had **no retry logic**.
> A single 429 from the Drive API would cause the entire audit to fail silently.

---

## 6. Deleted / Orphaned Scripts

| Condition | Observable Symptom | Auditor Behavior |
|---|---|---|
| Script in Trash | `trashed=true` in Drive metadata | Filtered by `trashed=false` query |
| Orphaned script (parents empty) | `parents: []` in API response | Currently causes null-ref in parent lookup |
| Parent folder deleted | `Drive.Files.get(parentId)` returns 404 | Unhandled exception in original code |
| Script owned by deleted user | Metadata incomplete, name may be blank | Row written with empty fields |

### The `parents[]` Empty Array Bug

The original `getAllScriptProjects()` accesses `file.parents[0]` directly
without checking if the array is empty. Orphaned scripts (scripts whose
parent folder was deleted, or root-level scripts that lost their path) have
`parents: []`. Accessing index 0 on an empty array returns `undefined`, which
then causes `Drive.Files.get(undefined)` to throw.

**Fix:**
```javascript
var parentId = (file.parents && file.parents.length > 0) ? file.parents[0] : null;
var parentName = parentId ? _resolveParentName(parentId) : "(orphaned / root)";
```

---

## 7. @HEAD Deployment Entries

| Property | Detail |
|---|---|
| **What it is** | A synthetic "deployment" representing the live editor HEAD state |
| **Always present** | Yes — every Apps Script project has exactly one @HEAD entry |
| **Has a URL** | No — `entryPoints` array is empty for @HEAD |
| **Should appear in audit** | No — it is not a real deployment |

### How to Identify and Filter @HEAD

```javascript
function isHeadDeployment(deployment) {
  // Method 1: Check description
  if (deployment.deploymentConfig &&
      deployment.deploymentConfig.description === "@HEAD") {
    return true;
  }
  // Method 2: Check for empty or missing entryPoints
  if (!deployment.entryPoints || deployment.entryPoints.length === 0) {
    return true;
  }
  // Method 3: Check for null/undefined versionNumber (HEAD has no pinned version)
  if (deployment.deploymentConfig &&
      deployment.deploymentConfig.versionNumber == null) {
    return true;
  }
  return false;
}

// Usage:
var realDeployments = deployments.filter(function(d) {
  return !isHeadDeployment(d);
});
```

> **Note:** The original spec did not filter @HEAD deployments, causing
> every project to appear to have at least one "deployment" even if it had
> never been deployed.

---

## 8. Multiple Deployments per Project

A single Apps Script project can have multiple active deployments (e.g.,
a "Production" deployment at version 10 and a "Staging" deployment at version 11).

### Serialization in Audit Output

The auditor serializes multiple deployments into a single spreadsheet row
using a pipe-delimited format:

```
DEPLOYMENT NAME    : Production | Staging | Beta
VERSION            : 10 | 11 | 12
ACCESS TYPE        : ANYONE | DOMAIN | ANYONE_ANONYMOUS
WEB APP URL        : https://script.google.com/... | https://... | (none)
```

### Recommended Alternative: One Row Per Deployment

For accounts with complex multi-deployment scripts, consider expanding the
output to one row per deployment rather than one row per project. This
preserves full fidelity but increases sheet row count proportionally.

---

## 9. Scripts You Do Not Own

| Condition | API Behavior | Auditor Impact |
|---|---|---|
| Script shared with you (Editor) | Appears in `projects.list`, full metadata | Works correctly |
| Script shared with you (Viewer) | Appears in `projects.list`, limited metadata | Deployment fetch may 403 |
| Script in a group/org Drive | Appears with `includeItemsFromAllDrives` | Owner field may show org email |
| Script owned by deleted user | Metadata may be incomplete | Name may be empty |

### Owner Field Caveat

The `owners` array in Drive metadata returns the file owner's email and
display name. For scripts owned by other users (shared-with-you scenarios),
the owner will show the other user's email — this is expected. The auditor
should record this accurately rather than assuming all scripts are self-owned.

---

## 10. Large Accounts (>500 Scripts)

| Script Count | Estimated Audit Time | Risk Level |
|---|---|---|
| 1–50 | < 30 seconds | None |
| 51–200 | 30–120 seconds | Low |
| 201–500 | 2–5 minutes | Medium — approach 6-min limit |
| 500+ | 5–15 minutes (continuation required) | HIGH — must use resumeAudit() |

### Pagination Behavior

Drive.Files.list returns a maximum of 1,000 files per page. Scripts
beyond 1,000 require cursor-based pagination via `nextPageToken`:

```javascript
var allFiles = [];
var pageToken = null;
do {
  var params = {
    pageSize: 200,
    q: "mimeType='application/vnd.google-apps.script' and trashed=false",
    fields: "nextPageToken,files(id,name,parents,createdTime)",
    supportsAllDrives: true,
    includeItemsFromAllDrives: true
  };
  if (pageToken) params.pageToken = pageToken;
  var page = Drive.Files.list(params);
  allFiles = allFiles.concat(page.files || []);
  pageToken = page.nextPageToken;
} while (pageToken);
```

### appendRow() Performance Issue

The original code called `sheet.appendRow()` inside a per-script loop.
At 500 scripts, this is 500 individual Sheets write calls, each costing
~200–400ms and consuming Sheets quota. The correct pattern is to batch
all rows into a 2D array and write once with `setValues()`:

```javascript
// WRONG (original spec):
scripts.forEach(function(s) {
  sheet.appendRow([s.name, s.id, s.url]);  // 1 write per script
});

// CORRECT (batch write):
var rows = scripts.map(function(s) {
  return [s.name, s.id, s.url];
});
sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
```

---

## 11. Rate Limiting and Exponential Backoff

### HTTP 429 — Too Many Requests

Both the Drive API and Apps Script API return HTTP 429 when quota is
exceeded. The response includes a `Retry-After` header (in seconds).

| API | Typical Trigger | Retry-After |
|---|---|---|
| Drive API | >1000 reads/100s | 10–60 seconds |
| Apps Script API | >100 reads/100s | 30–120 seconds |
| UrlFetchApp quota | >20,000 calls/day | Next UTC midnight |

### HTTP 503 — Service Unavailable

Transient GCP outages return 503. These are safe to retry immediately
with a short backoff. Do not confuse with 429 (quota exhaustion).

### Backoff Strategy

```
Attempt 1: Immediate
Attempt 2: Wait 1 second
Attempt 3: Wait 2 seconds
Attempt 4: Wait 4 seconds (exponential)
Give up after 4 attempts — log the failure, continue to next script
```

Do **not** use `Utilities.sleep()` for longer than 30 seconds per
invocation — you will consume your 6-minute execution budget waiting.
If a 429 Retry-After is >30 seconds, log and skip that script; write
a continuation marker to Script Properties and re-queue via trigger.

---

## 12. appsscript.json Manifest Scope Mismatch

### The Problem

If the `oauthScopes` array in `appsscript.json` does not include every
scope that the code actually uses, one of two things happens:

1. **Missing scope in manifest but used in code:** Google automatically
   adds the scope at runtime — BUT only if the automatic scope detection
   catches it. Complex `UrlFetchApp` calls to external APIs are NOT
   automatically detected. The result is a 403 on the external API call
   with no clear error message.

2. **Declared scope but not actually used:** The OAuth consent screen
   asks the user for permissions that the script does not need. This is
   a privacy/security issue that will trigger consent screen warnings.

### Required Scopes for the GAS Project Auditor

```json
{
  "timeZone": "America/New_York",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Drive",
        "version": "v3",
        "serviceId": "drive"
      }
    ]
  },
  "oauthScopes": [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/script.projects.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/script.external_request"
  ],
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

### Scope Verification at Runtime

`checkOAuthScopes()` in Tests.js validates this by hitting the Google
tokeninfo endpoint and inspecting the `scope` field of the live token.
Run it after any changes to `appsscript.json` to confirm the new scopes
take effect (may require re-authorization).

> **Important:** After adding new scopes to `appsscript.json`, you MUST
> revoke existing authorization and re-run the script to trigger the
> new consent prompt. Simply editing the manifest is not enough — the
> existing OAuth token continues to be used until it expires or is revoked.
> Revoke at: https://myaccount.google.com/permissions

---

## Summary Matrix

| # | Limitation | Severity | Has Workaround? |
|---|---|---|---|
| 1 | 6-minute execution cap | HIGH | YES — resumeAudit() + trigger |
| 2 | Container-bound scripts invisible to Drive | HIGH | YES — Apps Script API projects.list |
| 3 | 18 vs 57 count discrepancy | HIGH | YES — see source breakdown above |
| 4 | Shared Drive scripts silently excluded | HIGH | YES — allDrives params |
| 5 | API quota limits (429) | MEDIUM | YES — retry with backoff |
| 6 | Orphaned scripts (parents[]) crash | MEDIUM | YES — null-guard the parents array |
| 7 | @HEAD appears as fake deployment | MEDIUM | YES — filter by description/@HEAD |
| 8 | Multi-deployment serialization | LOW | YES — pipe-delimited or expand rows |
| 9 | Shared-with-me scripts have limited metadata | LOW | PARTIAL — log owner accurately |
| 10 | >500 scripts exceed time limit | HIGH (at scale) | YES — batch + continuation cursor |
| 11 | Rate limiting (429 / 503) | MEDIUM | YES — exponential backoff |
| 12 | Manifest scope mismatch | MEDIUM | YES — checkOAuthScopes() preflight |

---

*Document maintained alongside `Code.js` and `Tests.js` in the
00-GAS-Project-Auditor project directory.*
