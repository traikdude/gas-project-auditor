/**
 * Utils.js — GAS Project Auditor
 * Shared utility helpers consumed across all auditor modules.
 * All functions are global (no ES-module syntax) for GAS V8 compatibility.
 *
 * Dependencies: Config.js must be evaluated before this file.
 *
 * @module Utils
 */

// ---------------------------------------------------------------------------
// 1. retryWithBackoff
// ---------------------------------------------------------------------------

/**
 * Executes `fn` up to CONFIG.RETRY_MAX_ATTEMPTS times using exponential
 * backoff with full jitter. Re-throws the last error if all attempts fail.
 *
 * Backoff formula (per attempt index `i`, 0-based):
 *   baseInterval = CONFIG.RETRY_BASE_DELAY_MS * Math.pow(2, i)
 *   capped        = Math.min(baseInterval, CONFIG.RETRY_MAX_DELAY_MS)
 *   delay         = Math.random() * capped   // full jitter
 *
 * @param {function(): *} fn       - Zero-argument function to attempt.
 * @param {string}       [context] - Optional label for log messages.
 * @returns {*} The return value of `fn` on success.
 * @throws {Error} The last caught error after all attempts are exhausted.
 */
function retryWithBackoff(fn, context) {
  var label = context || 'retryWithBackoff';
  var lastError;

  for (var attempt = 0; attempt < CONFIG.RETRY_MAX_ATTEMPTS; attempt++) {
    try {
      return fn();
    } catch (err) {
      lastError = err;

      var isLastAttempt = (attempt === CONFIG.RETRY_MAX_ATTEMPTS - 1);
      if (isLastAttempt) {
        break; // fall through to throw
      }

      // Exponential backoff with full jitter
      var baseInterval = CONFIG.RETRY_BASE_DELAY_MS * Math.pow(2, attempt);
      var cappedInterval = Math.min(baseInterval, CONFIG.RETRY_MAX_DELAY_MS);
      var jitteredDelay = Math.floor(Math.random() * cappedInterval);

      Logger.log(
        '[' + label + '] Attempt ' + (attempt + 1) + ' failed: ' + err.message +
        '. Retrying in ' + jitteredDelay + ' ms…'
      );

      Utilities.sleep(jitteredDelay);
    }
  }

  Logger.log(
    '[' + label + '] All ' + CONFIG.RETRY_MAX_ATTEMPTS + ' attempts failed. ' +
    'Last error: ' + (lastError ? lastError.message : 'unknown')
  );
  throw lastError;
}


// ---------------------------------------------------------------------------
// 2. formatTimestamp
// ---------------------------------------------------------------------------

/**
 * Converts an ISO 8601 date string to a human-readable 'YYYY-MM-DD HH:mm'
 * string rendered in the script's configured timezone (America/New_York).
 *
 * Returns 'N/A' for falsy input, and the raw input string if parsing fails.
 *
 * @param {string} isoString - ISO date string, e.g. '2024-03-15T18:30:00.000Z'.
 * @returns {string} Formatted timestamp or fallback label.
 */
function formatTimestamp(isoString) {
  if (!isoString) {
    return 'N/A';
  }

  try {
    var date = new Date(isoString);
    if (isNaN(date.getTime())) {
      return String(isoString);
    }
    // Utilities.formatDate uses the script's Apps Script timezone setting
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  } catch (e) {
    Logger.log('[formatTimestamp] Parse error for "' + isoString + '": ' + e.message);
    return String(isoString);
  }
}


// ---------------------------------------------------------------------------
// 3. sanitizeForSheet
// ---------------------------------------------------------------------------

/**
 * Coerces a value into a form safe for writing to a Google Sheets cell.
 *
 * Rules applied:
 *  - null / undefined → empty string ''
 *  - boolean          → kept as-is (Sheets renders TRUE/FALSE natively)
 *  - number           → kept as-is
 *  - object / array   → JSON-serialised string (for debugging traceability)
 *  - string           → trimmed to 50 000 characters (Sheets cell limit)
 *  - anything else    → String() coercion
 *
 * @param {*} value - The raw value to sanitize.
 * @returns {string|number|boolean} A Sheets-safe scalar.
 */
function sanitizeForSheet(value) {
  var CELL_CHAR_LIMIT = 50000;

  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'boolean' || typeof value === 'number') {
    return value;
  }

  if (typeof value === 'object') {
    try {
      var serialized = JSON.stringify(value);
      return serialized.length > CELL_CHAR_LIMIT
        ? serialized.slice(0, CELL_CHAR_LIMIT)
        : serialized;
    } catch (e) {
      return '[unserializable object]';
    }
  }

  // String path
  var str = String(value);
  return str.length > CELL_CHAR_LIMIT ? str.slice(0, CELL_CHAR_LIMIT) : str;
}


// ---------------------------------------------------------------------------
// 4. getExecutionElapsedMs
// ---------------------------------------------------------------------------

/**
 * Returns the number of milliseconds elapsed since `startTime`.
 *
 * @param {Date} startTime - The Date object captured at execution start.
 * @returns {number} Elapsed milliseconds (always >= 0).
 */
function getExecutionElapsedMs(startTime) {
  return Date.now() - startTime.getTime();
}


// ---------------------------------------------------------------------------
// 5. getRemainingExecutionMs
// ---------------------------------------------------------------------------

/**
 * Returns the number of milliseconds remaining within the safe execution
 * budget defined by CONFIG.MAX_EXECUTION_SECONDS.
 *
 * A negative return value means the budget has already been exceeded —
 * callers should checkpoint / persist state and exit immediately.
 *
 * @param {Date} startTime - The Date object captured at execution start.
 * @returns {number} Remaining milliseconds. Negative = over budget.
 */
function getRemainingExecutionMs(startTime) {
  var budgetMs = CONFIG.MAX_EXECUTION_SECONDS * 1000;
  return budgetMs - getExecutionElapsedMs(startTime);
}


// ---------------------------------------------------------------------------
// 6. logToSheet
// ---------------------------------------------------------------------------

/**
 * Appends a single log row to the designated execution log sheet.
 *
 * Row format: [ISO timestamp, level, message]
 *
 * Levels are conventionally 'INFO', 'WARN', or 'ERROR'.
 * The function is intentionally tolerant — it swallows its own errors so that
 * a sheet-write failure during error-handling never masks the original problem.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} logSheet - Target sheet object.
 * @param {'INFO'|'WARN'|'ERROR'} level                 - Severity label.
 * @param {string} message                              - Human-readable message.
 * @returns {void}
 */
function logToSheet(logSheet, level, message) {
  if (!logSheet) {
    Logger.log('[logToSheet] logSheet is null — cannot write: [' + level + '] ' + message);
    return;
  }

  try {
    var timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
    logSheet.appendRow([
      timestamp,
      sanitizeForSheet(level),
      sanitizeForSheet(message)
    ]);
  } catch (e) {
    // Do NOT re-throw — logging must never crash the caller.
    Logger.log('[logToSheet] Failed to write to sheet: ' + e.message);
  }
}


// ---------------------------------------------------------------------------
// 7. buildScriptEditorUrl
// ---------------------------------------------------------------------------

/**
 * Constructs the direct Apps Script IDE URL for a given script project.
 *
 * @param {string} scriptId - The Apps Script project ID (not the Drive file ID).
 * @returns {string} A fully-qualified URL to the script editor.
 * @example
 *   buildScriptEditorUrl('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms')
 *   // → 'https://script.google.com/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms/edit'
 */
function buildScriptEditorUrl(scriptId) {
  if (!scriptId) {
    return '';
  }
  return 'https://script.google.com/d/' + scriptId + '/edit';
}
