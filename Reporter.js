/**
 * Reporter.js — Google Sheets Output Module
 * GAS Project Auditor
 *
 * Handles ALL Google Sheet interactions:
 *   - Sheet initialization (Audit Results, Execution Log, Summary Dashboard)
 *   - Batch row writes with HYPERLINK formula injection
 *   - Alternating row colors
 *   - Conditional formatting rules
 *   - Column auto-resize and freeze
 *   - Summary dashboard stats population
 *
 * All functions are globally scoped (no ES module syntax).
 * Depends on CONFIG object defined in Config.js (or equivalent).
 *
 * @author  GAS Project Auditor
 * @version 1.0.0
 */

// ---------------------------------------------------------------------------
// Column index constants (1-based) — must match _buildRow() in Code.js
// Col: A            B             C              D               E
//      Script ID     Script Name   Owner Email    Created Time    Last Modified
// Col: F             G             H              I               J
//      Shared?       Parent ID     Parent Name    Parent Type     Is Container?
// Col: K             L             M              N               O
//      Parent URL    Has Web App?  Deployment ID  Deploy Desc     Web App URL
// Col: P             Q             R
//      Access Level  Execute As    Version Number
// ---------------------------------------------------------------------------
var COL_SCRIPT_ID        = 1;   // A: Script ID
var COL_SCRIPT_NAME      = 2;   // B: Script Name
var COL_OWNER_EMAIL      = 3;   // C: Owner Email
var COL_CREATED          = 4;   // D: Created Time
var COL_LAST_MODIFIED    = 5;   // E: Last Modified
var COL_SHARED           = 6;   // F: Shared?
var COL_PARENT_ID        = 7;   // G: Parent ID
var COL_PARENT_NAME      = 8;   // H: Parent Name
var COL_PARENT_TYPE      = 9;   // I: Parent Type
var COL_BOUND            = 10;  // J: Is Container?  <- conditional format (Yes = green)
var COL_PARENT_URL       = 11;  // K: Parent URL     <- HYPERLINK
var COL_HAS_WEBAPP       = 12;  // L: Has Web App?
var COL_DEPLOY_ID        = 13;  // M: Deployment ID
var COL_DEPLOY_DESC      = 14;  // N: Deploy Description
var COL_WEBAPP_URL       = 15;  // O: Web App URL    <- HYPERLINK  conditional format
var COL_ACCESS_LEVEL     = 16;  // P: Access Level
var COL_EXECUTE_AS       = 17;  // Q: Execute As
var COL_VERSION_NUM      = 18;  // R: Version Number

var HEADER_ROW = [
  'Script ID',
  'Script Name',
  'Owner Email',
  'Created Time',
  'Last Modified',
  'Shared?',
  'Parent ID',
  'Parent Name',
  'Parent Type',
  'Is Container?',
  'Parent URL',
  'Has Web App?',
  'Deployment ID',
  'Deploy Description',
  'Web App URL',
  'Access Level',
  'Execute As',
  'Version #'
];

// Column widths in pixels (index 0 = col A = Script ID)
var COLUMN_WIDTHS = [160, 200, 180, 130, 130, 60, 130, 200, 130, 90, 200, 80, 160, 180, 300, 120, 140, 70];

// Sheet name constants - falls back to CONFIG if defined, else uses literals
var RESULTS_SHEET_NAME   = (typeof CONFIG !== 'undefined' && CONFIG.RESULTS_SHEET_NAME)   ? CONFIG.RESULTS_SHEET_NAME   : 'Audit Results';
var LOG_SHEET_NAME       = (typeof CONFIG !== 'undefined' && CONFIG.LOG_SHEET_NAME)       ? CONFIG.LOG_SHEET_NAME       : 'Execution Log';
var DASHBOARD_SHEET_NAME = (typeof CONFIG !== 'undefined' && CONFIG.DASHBOARD_SHEET_NAME) ? CONFIG.DASHBOARD_SHEET_NAME : 'Summary';

// Row colors for alternating stripes
var COLOR_ROW_EVEN = '#f8f9fa';
var COLOR_ROW_ODD  = '#ffffff';

// Header style colors
var COLOR_RESULTS_HEADER   = '#1a73e8';  // Google Blue
var COLOR_LOG_HEADER       = '#424242';  // Charcoal
var COLOR_DASHBOARD_TITLE  = '#1a73e8';

// Conditional format colors
var COLOR_BOUND_BG    = '#e6f4ea';  // Light green - bound scripts
var COLOR_WEBAPP_BG   = '#e8f0fe';  // Light blue  - scripts with deployments

// ---------------------------------------------------------------------------
// Helper: get or create a sheet by name, clearing it if it already exists
// ---------------------------------------------------------------------------
function _getOrCreateSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clearContents();
    sheet.clearFormats();
    sheet.clearConditionalFormatRules();
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
  } else {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// Helper: apply HYPERLINK formula to a cell if value is a URL
// ---------------------------------------------------------------------------
function _applyHyperlinkIfUrl_(cell, rawValue, labelOverride) {
  if (typeof rawValue === 'string' && rawValue.indexOf('http') === 0) {
    var label = labelOverride || rawValue;
    var escapedUrl   = rawValue.replace(/"/g, '""');
    var escapedLabel = label.replace(/"/g, '""');
    cell.setFormula('=HYPERLINK("' + escapedUrl + '","' + escapedLabel + '")');
    cell.setFontColor('#1155cc');
  } else if (rawValue !== '' && rawValue !== null && rawValue !== undefined) {
    cell.setValue(rawValue);
  }
}

// ---------------------------------------------------------------------------
// 1. initializeSheet(ss)
//    Creates/clears the 'Audit Results' tab and writes the formatted header.
// ---------------------------------------------------------------------------
function initializeSheet(ss) {
  var sheet = _getOrCreateSheet_(ss, RESULTS_SHEET_NAME);

  // --- Header row ---
  var headerRange = sheet.getRange(1, 1, 1, HEADER_ROW.length);
  headerRange.setValues([HEADER_ROW]);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setBackground(COLOR_RESULTS_HEADER);
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setWrap(false);

  // Freeze row 1
  sheet.setFrozenRows(1);

  // --- Column widths ---
  for (var i = 0; i < COLUMN_WIDTHS.length; i++) {
    sheet.setColumnWidth(i + 1, COLUMN_WIDTHS[i]);
  }

  sheet.setRowHeight(1, 36);

  Logger.log('[Reporter] Audit Results sheet initialized.');
  return sheet;
}

// ---------------------------------------------------------------------------
// 2. initializeLogSheet(ss)
//    Creates/clears the 'Execution Log' tab.
// ---------------------------------------------------------------------------
function initializeLogSheet(ss) {
  var sheet = _getOrCreateSheet_(ss, LOG_SHEET_NAME);

  var logHeaders = ['Timestamp', 'Level', 'Message'];
  var headerRange = sheet.getRange(1, 1, 1, logHeaders.length);
  headerRange.setValues([logHeaders]);
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setBackground(COLOR_LOG_HEADER);
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');
  headerRange.setWrap(false);

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 180);  // Timestamp
  sheet.setColumnWidth(2, 80);   // Level
  sheet.setColumnWidth(3, 600);  // Message
  sheet.setRowHeight(1, 36);

  Logger.log('[Reporter] Execution Log sheet initialized.');
  return sheet;
}

// ---------------------------------------------------------------------------
// 3. initializeDashboardSheet(ss)
//    Creates/clears the 'Summary' tab, moves it to index 0, writes the title.
// ---------------------------------------------------------------------------
function initializeDashboardSheet(ss) {
  var sheet = _getOrCreateSheet_(ss, DASHBOARD_SHEET_NAME);

  // Move Summary to be the FIRST tab
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);

  // --- Title cell: merge A1:D1 ---
  var titleRange = sheet.getRange('A1:D1');
  titleRange.merge();
  titleRange.setValue('GAS Project Audit Summary');
  titleRange.setFontSize(16);
  titleRange.setFontWeight('bold');
  titleRange.setHorizontalAlignment('center');
  titleRange.setVerticalAlignment('middle');
  titleRange.setBackground(COLOR_DASHBOARD_TITLE);
  titleRange.setFontColor('#ffffff');
  sheet.setRowHeight(1, 48);

  // --- Spacer row 2 ---
  sheet.setRowHeight(2, 12);

  // --- Column widths for the dashboard ---
  sheet.setColumnWidth(1, 240);  // Label column
  sheet.setColumnWidth(2, 160);  // Value column
  sheet.setColumnWidth(3, 40);   // Spacer
  sheet.setColumnWidth(4, 40);   // Spacer

  Logger.log('[Reporter] Summary dashboard sheet initialized.');
  return sheet;
}

// ---------------------------------------------------------------------------
// 4. flushBatch(sheet, batchRows, startRowIndex)
//    Writes all rows at once then applies alternating colors + HYPERLINK formulas.
// ---------------------------------------------------------------------------
function flushBatch(sheet, batchRows, startRowIndex) {
  if (!batchRows || batchRows.length === 0) return;

  var numRows = batchRows.length;
  var numCols = batchRows[0].length;

  // --- Bulk write all values ---
  sheet.getRange(startRowIndex, 1, numRows, numCols).setValues(batchRows);

  // --- Row-by-row post-processing: colors + hyperlinks ---
  for (var i = 0; i < numRows; i++) {
    var sheetRowIndex = startRowIndex + i;
    var rowData       = batchRows[i];

    // Alternating row background (row 2 = first data row)
    var rowColor = (sheetRowIndex % 2 === 0) ? COLOR_ROW_EVEN : COLOR_ROW_ODD;
    sheet.getRange(sheetRowIndex, 1, 1, numCols).setBackground(rowColor);

    // Column D -- Script Editor URL (col 4)
    var editorUrl = rowData[COL_EDITOR_URL - 1];
    if (editorUrl && typeof editorUrl === 'string' && editorUrl.indexOf('http') === 0) {
      var editorCell = sheet.getRange(sheetRowIndex, COL_EDITOR_URL);
      _applyHyperlinkIfUrl_(editorCell, editorUrl, 'Open Editor');
    }

    // Column I -- Container URL (col 9)
    var containerUrl = rowData[COL_CONTAINER_URL - 1];
    if (containerUrl && typeof containerUrl === 'string' && containerUrl.indexOf('http') === 0) {
      var containerCell = sheet.getRange(sheetRowIndex, COL_CONTAINER_URL);
      _applyHyperlinkIfUrl_(containerCell, containerUrl, 'Open Container');
    }

    // Column J -- Web App URL(s) (col 10)
    var webAppUrl = rowData[COL_WEBAPP_URLS - 1];
    if (webAppUrl && typeof webAppUrl === 'string' && webAppUrl.indexOf('http') === 0) {
      var webAppCell = sheet.getRange(sheetRowIndex, COL_WEBAPP_URLS);
      _applyHyperlinkIfUrl_(webAppCell, webAppUrl, 'Open Web App');
    }
  }

  // --- Vertical alignment + font for data rows ---
  sheet.getRange(startRowIndex, 1, numRows, numCols)
    .setVerticalAlignment('middle')
    .setFontSize(10);

  // Clear the batch array in-place (mutates caller reference)
  batchRows.length = 0;

  Logger.log('[Reporter] Flushed ' + numRows + ' rows starting at row ' + startRowIndex + '.');
}

// ---------------------------------------------------------------------------
// 5. finalizeSheet(ss, stats)
//    stats = { total, bound, standalone, withWebApp, errors }
// ---------------------------------------------------------------------------
function finalizeSheet(ss, stats) {
  // 5a. Populate the dashboard
  updateDashboard(ss, stats);

  // 5b. Get Audit Results sheet
  var auditSheet = ss.getSheetByName(RESULTS_SHEET_NAME);
  if (!auditSheet) {
    Logger.log('[Reporter] WARNING: Audit Results sheet not found during finalize.');
    return;
  }

  var lastRow = auditSheet.getLastRow();
  var lastCol = auditSheet.getLastColumn();

  // 5c. Auto-resize all columns
  if (lastCol > 0) {
    auditSheet.autoResizeColumns(1, lastCol);
  }

  // 5d. Conditional formatting
  if (lastRow > 1) {
    // Col J (Is Container? / Bound) = 'Yes' -> light green
    var boundRange = auditSheet.getRange(2, COL_BOUND, lastRow - 1, 1);
    var boundRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Yes')
      .setBackground(COLOR_BOUND_BG)
      .setRanges([boundRange])
      .build();

    // Col L (Has Web App?) = 'Yes' -> light blue
    var webAppRange = auditSheet.getRange(2, COL_HAS_WEBAPP, lastRow - 1, 1);
    var deployRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Yes')
      .setBackground(COLOR_WEBAPP_BG)
      .setRanges([webAppRange])
      .build();

    var rules = auditSheet.getConditionalFormatRules();
    rules.push(boundRule);
    rules.push(deployRule);
    auditSheet.setConditionalFormatRules(rules);
  }

  // 5e. Sort by Script Name (col B = COL_SCRIPT_NAME = 2) ascending, skipping header
  if (lastRow > 2) {
    var sortRange = auditSheet.getRange(2, 1, lastRow - 1, lastCol);
    sortRange.sort({ column: COL_SCRIPT_NAME, ascending: true });
  }

  // 5f. Activate Summary sheet
  var dashSheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (dashSheet) {
    ss.setActiveSheet(dashSheet);
  }

  Logger.log('[Reporter] Sheet finalized. Total scripts: ' + (stats ? stats.total : 'N/A'));
}

// ---------------------------------------------------------------------------
// 6. updateDashboard(ss, stats)
//    Writes stats pairs to the Summary sheet starting at row 3.
// ---------------------------------------------------------------------------
function updateDashboard(ss, stats) {
  var sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet) {
    Logger.log('[Reporter] WARNING: Summary sheet not found. Cannot update dashboard.');
    return;
  }

  stats = stats || {};
  var total      = stats.total      || 0;
  var bound      = stats.bound      || 0;
  var standalone = stats.standalone || 0;
  var withWebApp = stats.withWebApp || 0;
  var errors     = stats.errors     || 0;

  var runDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );

  var runBy = '';
  try {
    runBy = Session.getActiveUser().getEmail() || 'unknown';
  } catch (e) {
    runBy = 'unknown (no permission)';
  }

  var statRows = [
    ['Total Scripts Found',       total],
    ['Container-Bound Scripts',   bound],
    ['Standalone Scripts',        standalone],
    ['Scripts with Web App',      withWebApp],
    ['Scripts with No Web App',   total - withWebApp],
    ['Errors Encountered',        errors],
    ['Audit Run Date',            runDate],
    ['Audit Run By',              runBy]
  ];

  var startRow = 3;
  var numStats = statRows.length;

  // Write labels + values
  sheet.getRange(startRow, 1, numStats, 2).setValues(statRows);

  // Format label column (A)
  var labelRange = sheet.getRange(startRow, 1, numStats, 1);
  labelRange.setFontWeight('bold');
  labelRange.setFontSize(11);
  labelRange.setHorizontalAlignment('left');
  labelRange.setVerticalAlignment('middle');
  labelRange.setBackground('#f1f3f4');

  // Format value column (B)
  var valueRange = sheet.getRange(startRow, 2, numStats, 1);
  valueRange.setFontWeight('bold');
  valueRange.setFontSize(11);
  valueRange.setHorizontalAlignment('center');
  valueRange.setVerticalAlignment('middle');
  valueRange.setBackground('#ffffff');

  // Color-code specific value cells
  // Errors -> red tint if errors > 0
  if (errors > 0) {
    sheet.getRange(startRow + 5, 2).setBackground('#fce8e6').setFontColor('#c5221f');
  }
  // Web App count -> blue tint
  sheet.getRange(startRow + 3, 2).setBackground('#e8f0fe').setFontColor('#1967d2');
  // Bound scripts -> green tint
  sheet.getRange(startRow + 1, 2).setBackground('#e6f4ea').setFontColor('#137333');

  // Row heights
  for (var r = 0; r < numStats; r++) {
    sheet.setRowHeight(startRow + r, 32);
  }

  // Border around the stats block
  sheet.getRange(startRow, 1, numStats, 2).setBorder(
    true, true, true, true, true, true,
    '#bdc1c6',
    SpreadsheetApp.BorderStyle.SOLID
  );

  // Divider row between title and stats
  sheet.getRange(2, 1, 1, 4).setBackground('#1a73e8');

  Logger.log('[Reporter] Dashboard updated with ' + numStats + ' stat rows.');
}

// ---------------------------------------------------------------------------
// 7. appendLogRow(logSheet, level, message)
//    Appends a single row to the Execution Log sheet.
//    level: 'INFO' | 'WARN' | 'ERROR'
// ---------------------------------------------------------------------------
function appendLogRow(logSheet, level, message) {
  if (!logSheet) return;

  var timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );

  var newRow = logSheet.getLastRow() + 1;
  logSheet.getRange(newRow, 1, 1, 3).setValues([[timestamp, level, message]]);

  // Color-code by level
  var levelCell = logSheet.getRange(newRow, 2);
  switch (level) {
    case 'ERROR':
      levelCell.setBackground('#fce8e6').setFontColor('#c5221f').setFontWeight('bold');
      break;
    case 'WARN':
      levelCell.setBackground('#fef7e0').setFontColor('#b06000').setFontWeight('bold');
      break;
    case 'INFO':
    default:
      levelCell.setBackground('#e6f4ea').setFontColor('#137333');
      break;
  }

  // Alternating background for timestamp and message
  var rowColor = (newRow % 2 === 0) ? '#f8f9fa' : '#ffffff';
  logSheet.getRange(newRow, 1).setBackground(rowColor);
  logSheet.getRange(newRow, 3).setBackground(rowColor);
}

// ---------------------------------------------------------------------------
// 8. applyHeaderProtection(ss)
//    Optional: Protect header rows with warning-only protection.
//    Call after finalizeSheet() if desired.
// ---------------------------------------------------------------------------
function applyHeaderProtection(ss) {
  var sheetsToProtect = [
    { name: RESULTS_SHEET_NAME,   rows: 1 },
    { name: LOG_SHEET_NAME,       rows: 1 },
    { name: DASHBOARD_SHEET_NAME, rows: 2 }
  ];

  sheetsToProtect.forEach(function(cfg) {
    var sheet = ss.getSheetByName(cfg.name);
    if (!sheet) return;
    var protection = sheet.getRange(1, 1, cfg.rows, sheet.getMaxColumns()).protect();
    protection.setDescription('Header row -- do not edit');
    protection.setWarningOnly(true);
  });

  Logger.log('[Reporter] Header protection applied to all sheets.');
}
