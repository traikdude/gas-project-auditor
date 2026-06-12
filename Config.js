/**
 * Config.js — GAS Project Auditor
 * Centralized configuration for all tunable parameters.
 * Edit values here; do NOT hardcode them in core logic files.
 *
 * @module Config
 */

// ---------------------------------------------------------------------------
// PRIMARY CONFIGURATION OBJECT
// ---------------------------------------------------------------------------

/** @type {Object} CONFIG — Single source of truth for auditor settings. */
var CONFIG = {

  // ── Project Identity (PERMANENT — do NOT remove or change) ───────────────

  /**
   * Google Cloud Platform project number linked to this Apps Script project.
   * Set in GAS Editor → Project Settings → GCP Project: 825046261103
   * This MUST remain consistent across all clasp pushes and deployments.
   */
  GCP_PROJECT_NUMBER: '825046261103',

  /**
   * The Apps Script Script ID for the GAS Project Auditor.
   * Used for self-referencing, API calls, and deployment verification.
   */
  SCRIPT_ID: '1TSn9NbYuw1Vt3Qod1d9iccX0t_YL_4oWnTniG2dCoRywyQHRxiKjUL94',

  // ── Output / Sheet identity ──────────────────────────────────────────────

  /** Prefix applied to every generated audit sheet tab name. */
  AUDIT_SHEET_PREFIX: 'GAS Audit',

  /** Name of the tab used for structured audit result rows. */
  RESULTS_SHEET_NAME: 'Audit Results',

  /** Name of the tab used for summary / dashboard metrics. */
  DASHBOARD_SHEET_NAME: 'Summary',

  /** Name of the tab used for timestamped execution log entries. */
  LOG_SHEET_NAME: 'Execution Log',

  // ── Quota & throughput controls ──────────────────────────────────────────

  /**
   * Milliseconds to pause between processing individual script projects.
   * Keeps the auditor well inside Apps Script's URL-fetch / Drive API quotas.
   * Increase if you see quota-exceeded errors.
   */
  SLEEP_BETWEEN_PROJECTS_MS: 150,

  /**
   * Number of result rows to accumulate in memory before flushing to the
   * sheet. Larger values reduce SpreadsheetApp write calls but consume more
   * heap. Tune against your project count.
   */
  BATCH_WRITE_SIZE: 20,

  /**
   * Safety wall in seconds. The auditor checks elapsed time and checkpoints
   * before reaching the hard 6-minute GAS execution cap (360 s).
   * 300 s gives a comfortable 60-second buffer for sheet writes and cleanup.
   */
  MAX_EXECUTION_SECONDS: 300,

  // ── Retry / backoff ──────────────────────────────────────────────────────

  /** Maximum number of API call attempts before propagating the error. */
  RETRY_MAX_ATTEMPTS: 3,

  /**
   * Base delay in milliseconds for the first retry interval.
   * Subsequent intervals are computed as:
   *   delay = MIN(BASE * 2^attempt + jitter, MAX)
   */
  RETRY_BASE_DELAY_MS: 1000,

  /**
   * Hard ceiling on any single retry delay.
   * Prevents runaway waits when the exponential factor grows large.
   */
  RETRY_MAX_DELAY_MS: 8000,

  // ── Drive traversal ──────────────────────────────────────────────────────

  /**
   * When true, the auditor enumerates Shared Drives (Team Drives) in addition
   * to My Drive. Requires the drive.readonly scope on a domain account or an
   * account with Shared Drive membership.
   */
  INCLUDE_ALL_DRIVES: true,

  // ── Deployment filtering ─────────────────────────────────────────────────

  /**
   * When true, rows whose deployment ID equals '@HEAD' are suppressed from
   * the results sheet. @HEAD entries are ephemeral development pseudo-
   * deployments and are rarely useful in a production audit.
   */
  FILTER_HEAD_DEPLOYMENTS: true,

  // ── PropertiesService resume keys ────────────────────────────────────────

  /**
   * Key used to persist the last-processed project index into PropertiesService.
   * Allows the auditor to resume a mid-flight run after a timeout.
   */
  PROPERTIES_KEY_INDEX: 'AUDIT_LAST_INDEX',

  /**
   * Key used to persist the output spreadsheet ID so that a resumed run
   * appends to the same sheet rather than creating a new one.
   */
  PROPERTIES_KEY_SHEET_ID: 'AUDIT_SHEET_ID'
};


// ---------------------------------------------------------------------------
// MIME TYPE MAP
// Maps Google Drive / Apps Script MIME type strings to human-readable labels.
// Used when rendering the "Project Type" column in audit output.
// ---------------------------------------------------------------------------

/**
 * @type {Object.<string, string>}
 * Keys are full MIME type strings as returned by the Drive API.
 * Values are short display labels suitable for a spreadsheet cell.
 */
var MIME_TYPE_MAP = {
  // Google Apps Script container-bound / standalone
  'application/vnd.google-apps.script':                'Apps Script',

  // Google Workspace document types (container-bound scripts live inside these)
  'application/vnd.google-apps.spreadsheet':           'Spreadsheet',
  'application/vnd.google-apps.document':              'Document',
  'application/vnd.google-apps.presentation':          'Presentation',
  'application/vnd.google-apps.form':                  'Form',
  'application/vnd.google-apps.site':                  'Sites',

  // Other Drive types that may surface during enumeration
  'application/vnd.google-apps.folder':                'Folder',
  'application/vnd.google-apps.drive-sdk':             'Drive SDK App',
  'application/vnd.google-apps.shortcut':              'Shortcut',
  'application/vnd.google-apps.unknown':               'Unknown',

  // Script project JSON (as seen in Drive exports)
  'application/vnd.google-apps.script+json':           'Script JSON',

  // Plain / binary fallback
  'application/json':                                  'JSON File',
  'text/plain':                                        'Plain Text',
  'application/octet-stream':                          'Binary'
};
