![Script files pass under a brass magnifying glass into a green spreadsheet ledger.](assets/banner/banner.png)

# GAS Project Auditor

Which Apps Script files can you find in Drive, and which ones expose a web app? GAS Project Auditor builds a Google Sheets inventory from the files visible to the account running it. Each row pairs a project with a returned web app entry point, so two deployments of the same project can appear side by side when you filter by script ID.

**Contents** · [Features](#features) · [Quick start](#quick-start) · [Configuration](#configuration) · [Output schema](#output-schema) · [Resume and limitations](#resume-and-limitations) · [File structure](#file-structure)

<a id="features"></a>
## 🔎 What it does

| Feature | What you get |
|---|---|
| Discover script files | Pages through non-trashed Apps Script files with the Drive v3 all-drives flags. Captures Drive names, timestamps, sharing status and available owner metadata. [Discovery loop](Code.js) |
| Inspect deployments | Reads the Apps Script deployments endpoint, skips the literal `@HEAD` ID and extracts `WEB_APP` entries. A project with no returned entry still gets a row. [Deployment lookup and row builder](Code.js) |
| Resolve parent files | Looks up the first Drive parent when present and records its name, type and link. A folder parent describes where a file lives. [Parent lookup](Code.js) |
| Read the report in Sheets | Writes rows in batches, links parent and web app URLs, formats the inventory and timestamps the execution log. A pre-existing Summary tab can show run statistics. [Sheet writer](Reporter.js) |

<a id="quick-start"></a>
## 🚀 Prerequisites and quick start

1. Create a standalone Apps Script project. Copy the five JavaScript modules and the optional container-discovery module listed below into editor script files. Enable manifest editing in Project Settings and use the supplied [V8 manifest](appsscript.json). If you maintain the project with clasp, sync these same source files and manifest to your own script project. Run the auditor from the editor.
2. In the linked Google Cloud project, enable the Drive and Apps Script APIs and configure OAuth consent for your account. Enable the Apps Script API in your account settings. Check that **Drive v3** appears under Services. The manifest declares the advanced service and the scopes used for Drive reads, deployment reads, external requests, account identity and spreadsheet writes. [Manifest](appsscript.json)
3. **Choose your output before the first run.** In your copy of [Config.js](Config.js), replace `MASTER_SPREADSHEET_ID` with the ID of a dedicated spreadsheet you control, or set it to an empty string to create a new spreadsheet. A fresh run clears the results and log tabs it initializes. For a dashboard on the first run, create a tab named **Summary** in your dedicated workbook.
4. Select **runFullAudit** in the editor, click **Run**, and authorize the requested scopes. Follow the execution log for the output spreadsheet URL. If opening the configured destination fails, the script attempts to create a new timestamped workbook. [Run entry point](Code.js)
5. Open **Audit Results** and filter on **Has Web App?** or a script ID. Keep the workbook private to its intended readers. It can contain owner addresses, file IDs, deployment IDs and web app URLs.

For a smaller diagnostic, `runDryRun()` samples discovery and logs its findings. The broader `runAllPreflightChecks()` includes a write check that creates a temporary spreadsheet and attempts to remove it. [Diagnostic routines](Tests.js)

<a id="configuration"></a>
## ⚙️ Configuration reference

| Setting | Effective behavior |
|---|---|
| Output and tab names | `MASTER_SPREADSHEET_ID` controls the destination. The reporter defaults to **Audit Results**, **Execution Log** and **Summary**. Keep the default log name when resuming. [Config](Config.js) · [Reporter](Reporter.js) |
| Discovery | The main loop requests up to 1,000 files per Drive page and sets all-drives flags directly. [Drive request](Code.js) |
| Timing and batching | The loop uses a 10-row flush threshold, waits 150 ms between projects and checks for a pause when less than one minute remains in its five-minute budget. These local constants take precedence over similarly named configuration fields. [Loop constants](Code.js) |

<a id="output-schema"></a>
## 📊 Output schema

The results sheet has **18 columns**. Use **Script ID** to group rows. Use the execution log to investigate missing deployment data. [Headers and formatting](Reporter.js) · [Row construction](Code.js)

| Columns | Fields, in sheet order |
|---|---|
| A–F · Project | Script ID, Script Name, Owner Email, Created Time, Last Modified, Shared? |
| G–K · Parent | Parent ID, Parent Name, Parent Type, Is Container?, Parent URL |
| L–R · Deployment | Has Web App?, Deployment ID, Deploy Description, Web App URL, Access Level, Execute As, Version # |

<a id="resume-and-limitations"></a>
## ⏸️ Resuming a paused audit and interpreting gaps

- **After a planned pause**, move **Audit Results** to the first tab before running `resumeAudit()`. Resume reopens the saved spreadsheet, selects its first tab and continues from a stored index in the saved project list. The log records the pause. Statistics are initialized per processing invocation, so a resumed summary can cover only the final segment. [Resume and processing loop](Code.js)
- **Discovery coverage depends on Drive visibility.** Container-bound scripts can be absent from the MIME-type search. The separate `DISCOVERY_auditAllSources()` routine explores additional candidates. It runs independently of the main audit and its coverage is unverified. [Supplemental discovery](ContainerBoundDiscovery.gs)
- **“No” under Has Web App? can also mean the lookup failed.** The lookup returns an empty list on several errors and reads only the first deployments response. Inspect the Apps Script execution logs before treating a missing entry as conclusive. [Deployment lookup](Code.js)
- **A completion message alone cannot confirm every row was written.** The loop logs certain write and finalization errors and continues. Large inventories also depend on the project list fitting in Script Properties. Inspect the rows and logs after a long run. [Processing loop](Code.js)

<a id="file-structure"></a>
## 🗂️ Tech stack and file structure

| File | Role |
|---|---|
| [appsscript.json](appsscript.json) | Apps Script V8 runtime, Drive advanced service and OAuth scopes |
| [Config.js](Config.js) | Output destination and configuration defaults |
| [Code.js](Code.js) | Drive discovery, REST deployment lookups, audit and resume entry points |
| [Reporter.js](Reporter.js) | Spreadsheet tabs, header definitions, batch writes and dashboard helpers |
| [Utils.js](Utils.js) | Retry, timestamp, logging and cell-value helpers |
| [Tests.js](Tests.js) | Editor-run diagnostics and sample discovery |
| [ContainerBoundDiscovery.gs](ContainerBoundDiscovery.gs) | Supplemental container discovery experiment |
| [LIMITATIONS.md](LIMITATIONS.md) | Background and proposed workarounds, including examples outside the main audit |
