# GAS Project Auditor

> **Automatically inventories every Google Apps Script project in your Google account — standalone and container-bound — and writes a fully formatted audit report directly to a Google Sheet.**

![Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=google&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-green)
![Status](https://img.shields.io/badge/status-production--ready-brightgreen)

---

## Table of Contents

1. [What It Does](#what-it-does)
2. [Prerequisites Checklist](#prerequisites-checklist)
3. [Quick Start](#quick-start)
4. [Configuration Reference](#configuration-reference)
5. [Output Schema](#output-schema)
6. [Resuming a Paused Audit](#resuming-a-paused-audit)
7. [Limitations](#limitations)
8. [Security Notes](#security-notes)
9. [File Structure](#file-structure)

---

## What It Does

The **GAS Project Auditor** crawls your Google Drive (and optionally Shared Drives) using the Apps Script API and Drive API, then for each script project it discovers:

| Discovery step | Detail |
|---|---|
| **Lists all script projects** | Uses `script.projects.list` with pagination |
| **Resolves bound containers** | Detects whether each script is bound to a Sheet, Doc, Form, Slide, or Site, and fetches the container's name and URL |
| **Enumerates deployments** | Calls `script.projects.deployments.list` to count active deployments and capture any Web App URLs |
| **Checks sharing status** | Reads the Drive file metadata to determine if the script is shared |

Results are streamed into a Google Sheet in configurable batches — so even accounts with hundreds of scripts complete without hitting memory limits. Three tabs are produced:

- **Summary** — at-a-glance statistics dashboard (always tab 1)
- **Audit Results** — full inventory table with clickable hyperlinks
- **Execution Log** — timestamped INFO / WARN / ERROR entries for every step

---

## Prerequisites Checklist

Complete **all five steps** before running the auditor. Missing any step is the most common cause of `Exception: You do not have permission` errors.

- [ ] **1. Apps Script API enabled**
  In the target GCP project, go to **APIs & Services → Library** and enable:
  `Apps Script API` (`script.googleapis.com`)

- [ ] **2. Drive API enabled**
  Enable `Drive API` (`drive.googleapis.com`) in the same GCP project.

- [ ] **3. Script linked to a GCP project with OAuth consent screen**
  In the Apps Script editor open **Project Settings → Google Cloud Platform (GCP) Project** and paste your GCP project number.
  The OAuth consent screen must be configured (Internal is fine for Workspace accounts).

- [ ] **4. Advanced service: Drive API v3 added inside Apps Script**
  In the editor: **Services (+) → Drive API → v3 → Add**.
  This exposes `Drive.Files.list()` and `Drive.Drives.list()` inside the script.

- [ ] **5. Apps Script API enabled in your account settings**
  Visit [script.google.com/home/usersettings](https://script.google.com/home/usersettings) and toggle
  **"Google Apps Script API" → On**.

---

## Quick Start

### Step 1 — Create the output spreadsheet

Open [sheets.google.com](https://sheets.google.com), create a blank spreadsheet, and copy its **Spreadsheet ID** from the URL:
```
https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit
```

### Step 2 — Deploy the script

Either:
- **Copy/paste** all `.js` files into a new Apps Script project at [script.google.com](https://script.google.com), **or**
- **Use clasp** to push the project from this directory:
  ```bash
  clasp login
  clasp create --type standalone --title "GAS Project Auditor"
  clasp push
  ```

### Step 3 — Configure

Open `Config.js` and update the required values:

```js
var CONFIG = {
  SPREADSHEET_ID:      'YOUR_SPREADSHEET_ID_HERE',  // from Step 1
  BATCH_SIZE:          50,                           // rows per flush
  MAX_SCRIPTS:         0,                            // 0 = no limit
  INCLUDE_SHARED_DRIVES: false,                      // toggle Shared Drive scan
};
```

### Step 4 — Run the audit

In the Apps Script editor, select the function `runAudit` from the dropdown and click ▶ **Run**.
Grant the required OAuth scopes when prompted (one-time per account).

### Step 5 — View results

The script opens / activates the output spreadsheet automatically.
Navigate to the **Summary** tab for the dashboard, or **Audit Results** for the full inventory.

---

## Configuration Reference

All settings live in `Config.js` under the `CONFIG` object.

| Key | Type | Default | Description |
|---|---|---|---|
| `SPREADSHEET_ID` | `string` | `''` | **Required.** ID of the Google Sheet to write results into. |
| `RESULTS_SHEET_NAME` | `string` | `'Audit Results'` | Name of the main inventory tab. Will be created or cleared on each run. |
| `LOG_SHEET_NAME` | `string` | `'Execution Log'` | Name of the execution log tab. |
| `DASHBOARD_SHEET_NAME` | `string` | `'Summary'` | Name of the summary dashboard tab. Always moved to position 1. |
| `BATCH_SIZE` | `number` | `50` | Number of script rows to accumulate before flushing to the sheet. Lower values reduce memory usage; higher values reduce API write calls. Recommended range: 25–100. |
| `MAX_SCRIPTS` | `number` | `0` | Maximum scripts to audit. `0` = unlimited (audit everything). Useful for dry-run testing. |
| `INCLUDE_SHARED_DRIVES` | `boolean` | `false` | When `true`, also scans all Shared Drives (Team Drives) visible to the running user. Requires `Drive API` to be enabled. |
| `PAGE_SIZE` | `number` | `50` | Number of results per page when calling `script.projects.list`. Max allowed by the API is 50. |
| `RESUME_KEY` | `string` | `'GAS_AUDIT_RESUME_TOKEN'` | Property key used to store the pagination token in `ScriptProperties` when an execution is interrupted. |
| `TIMEZONE` | `string` | `Session.getScriptTimeZone()` | Timezone string for date formatting (e.g., `'America/New_York'`). |

---

## Output Schema

The **Audit Results** sheet contains 13 columns:

| Col | Header | Type | Description |
|---|---|---|---|
| A | `#` | Number | Sequential row number (1-based). |
| B | `Script Name` | String | Display name of the Apps Script project as shown in the editor. |
| C | `Script ID` | String | The unique `scriptId` from the Apps Script API (same as the Drive file ID for standalone scripts). |
| D | `Script Editor URL` | Hyperlink | Clickable link that opens the project directly in the Apps Script editor (`script.google.com/home/projects/<scriptId>/edit`). |
| E | `Bound?` | `Yes` / `No` | Whether the script is container-bound (attached to a Sheet, Doc, Form, etc.) rather than standalone. |
| F | `Container Name` | String | Display name of the bound container file. Empty for standalone scripts. |
| G | `Container Type` | String | MIME type label of the container (e.g., `Google Sheet`, `Google Doc`, `Google Form`, `Google Slide`). Empty for standalone. |
| H | `Container ID` | String | Drive file ID of the container. Can be used to construct a direct URL. |
| I | `Container URL` | Hyperlink | Clickable link that opens the container file (Sheet, Doc, etc.) directly in Google Drive. |
| J | `Web App URL(s)` | Hyperlink | The deployment URL of the most recent Web App deployment, if any. Multiple deployments are separated by newlines. |
| K | `Deploy Count` | Number | Total number of active deployments for this script project. Includes Web Apps, API Executables, and Add-ons. |
| L | `Last Modified` | DateTime | ISO-formatted date/time when the script file was last modified according to Drive metadata. |
| M | `Shared?` | `Yes` / `No` | Whether the script file has been shared with anyone other than the owner. |

---

## Resuming a Paused Audit

Apps Script functions are subject to a **6-minute execution time limit**. For accounts with many scripts the audit may be interrupted. The auditor handles this automatically:

### How resume works

1. Before each page fetch, the current **pagination token** is saved to `ScriptProperties` under the key defined by `CONFIG.RESUME_KEY`.
2. If execution times out, the token is preserved.
3. Call `resumeAudit()` to pick up exactly where the run stopped:
   ```js
   // In the Apps Script editor, run this function manually:
   resumeAudit();
   ```
4. The resume function reads the saved token, skips already-processed rows, and continues appending to the same sheet — it does **not** clear existing data.
5. After a successful full completion, the stored token is deleted automatically.

### Manual reset

To force a fresh audit from scratch (clearing all stored state):
```js
resetAuditState();
```

This deletes the resume token and clears the output sheets before re-running.

---

## Limitations

### ⏱ Execution time (6-minute limit)
Apps Script enforces a hard 6-minute execution timeout. Accounts with **200+ scripts** may require multiple resume cycles. Each resume call picks up from the last saved page token.
**Mitigation:** Use `resumeAudit()` or set up a time-driven trigger to call it repeatedly until `ScriptProperties.getProperty(CONFIG.RESUME_KEY)` is `null`.

### 📦 Quota limits
The Apps Script API allows **~100 requests per 100 seconds** per user. The auditor includes a `Utilities.sleep(200)` between deployment list calls to stay within quota.
If you see `Error 429: Too many requests`, increase the sleep interval in `Config.js → SLEEP_MS`.

### 🗂 Shared Drives / Team Drives
By default, `INCLUDE_SHARED_DRIVES: false`. Enabling it adds a separate `Drive.Drives.list()` pass. Note:
- Only Shared Drives where the running user has **at least Commenter** access are scanned.
- Scripts **inside** Shared Drive containers can only be discovered if the runner has the Apps Script API scope for those files.

### 🔗 Container-bound discovery gap
The Apps Script API's `projects.list` returns scripts the authenticated user **owns**. Scripts bound to containers owned by other users (e.g., shared Sheets you can edit but don't own) are **not returned** — even if you have editor access. This is an API-level restriction with no workaround.

### 📛 Project names vs. file names
The `Script Name` column reflects the **Apps Script project title**, not the Drive file name. For bound scripts, these are usually the same as the container name — but can diverge if renamed independently in the editor.

---

## Security Notes

- **OAuth scopes** requested by this script:
  - `https://www.googleapis.com/auth/script.projects.readonly` — reads project metadata and deployments
  - `https://www.googleapis.com/auth/drive.metadata.readonly` — reads file names, sharing status, and container info
  - `https://www.googleapis.com/auth/spreadsheets` — writes results to the output sheet
  - `https://www.googleapis.com/auth/script.scriptapp` — reads script properties for resume state

- The auditor is **read-only** with respect to your Apps Script projects and Drive files. It creates and modifies only the designated output spreadsheet.

- The output spreadsheet will contain **all your script IDs**. Treat it as a sensitive internal document — avoid sharing it publicly. Script IDs can be used to call API Executable deployments if the correct OAuth token is obtained.

- If run under a **Google Workspace** account, the audit covers scripts owned by that specific user. It does **not** enumerate scripts owned by other users in the domain (requires Admin SDK / domain-wide delegation, which is out of scope).

---

## File Structure

```
00-GAS-Project-Auditor/
├── appsscript.json     — Manifest (scopes, runtime version)
├── Config.js           — All tunable settings (CONFIG object)
├── Auditor.js          — Core audit loop, pagination, Drive/API calls
├── Reporter.js         — ALL Google Sheets output logic (this module)
└── README.md           — This file
```

### Reporter.js public API

| Function | Signature | Purpose |
|---|---|---|
| `initializeSheet` | `(ss) → Sheet` | Create/clear Audit Results tab, write formatted header |
| `initializeLogSheet` | `(ss) → Sheet` | Create/clear Execution Log tab |
| `initializeDashboardSheet` | `(ss) → Sheet` | Create/clear Summary tab, move to position 0 |
| `flushBatch` | `(sheet, batchRows[], startRow)` | Batch-write rows + apply colors + inject HYPERLINK formulas |
| `finalizeSheet` | `(ss, stats)` | Conditional formatting, auto-resize, sort, activate dashboard |
| `updateDashboard` | `(ss, stats)` | Write stats pairs to Summary tab with color-coded values |
| `appendLogRow` | `(logSheet, level, msg)` | Append one color-coded row to the Execution Log |
| `applyHeaderProtection` | `(ss)` | Optional: lock header rows with warning-only protection |

---

*Generated by GAS Project Auditor v1.0.0 — 2026*
