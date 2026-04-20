// ======================================================================
// File: Import_FromExcel_PreserveFormatting.gs
// Purpose:
//   - Always treat SOURCE as an Excel (.xlsx) in Drive.
//   - Convert → read the given worksheet ("aggregatecsv csv").
//   - Filter Region == 10000002 and split by BuyOrder TRUE-ish / FALSE-ish.
//   - Write INTO existing tabs in MASTER, preserving formatting & filters.
//   - Maintain/refresh Named Ranges: Buy_TRUE_Data / Buy_FALSE_Data.
//
// Why this style:
//   - You asked to skip file-type checks: we always convert via Drive API.
//   - You want formatting + named ranges preserved: we never delete tabs;
//     we only clear and overwrite the data area and then update named ranges.
//
// Prereqs (one-time):
//   - In Apps Script → left sidebar “Services” → + → add **Drive API**.
//   - If prompted in Cloud Console, enable **Google Drive API** there.
//
// How to run:
//   - Hook Import_FromExcel_PreserveFormatting() in your existing menu script,
//     or run it directly from the editor after selecting that function.
//
// Logging / Debugging:
//   - Uses console + Logger at each step.
//   - If anything fails, errors are explicit about the failing step.
//
// ======================================================================

/** ----------------------- USER SETTINGS (EDIT ME) -------------------- */
const SETTINGS = {
  // Excel file (.xlsx) in Drive
  SOURCE_XLSX_FILE_ID: '1mjNjb_f7ryicL1cEjX5Igf6MTc98dxmP',
  SOURCE_WORKSHEET_NAME: 'aggregatecsv csv',   // worksheet name inside the Excel file

  // Master Google Sheet (destination)
  MASTER_SHEET_ID: '1Q5KS6B5h2mtFNGKOmOjX_1j5fbbnY02OBni2A8kzxe4',

  // Output tabs (must already exist if you want to keep their formatting)
  TRUE_TAB_NAME: 'Buy_TRUE',
  FALSE_TAB_NAME: 'Buy_FALSE',

  // Named ranges we will maintain on each write (data only, excludes header row)
  TRUE_NAMED_RANGE: 'Buy_TRUE_Data',
  FALSE_NAMED_RANGE: 'Buy_FALSE_Data',

  // Filter
  TARGET_REGION: '10000002',

  // Behavior
  DELETE_TEMP_CONVERTED_FILE: true,  // clean up the converted temp sheet
  SHOW_TOASTS: true                  // toasts if running in a bound project
};
/** --------------------- END USER SETTINGS (EDIT ME) ------------------ */

// -------------------- Logging helpers --------------------
function log_(level, msg) {
  const line = `[${level}] ${msg}`;
  Logger.log(line);
  console.log(line);
}
function toast_(msg, seconds = 5) {
  if (!SETTINGS.SHOW_TOASTS) return;
  try { const ss = SpreadsheetApp.getActiveSpreadsheet(); if (ss) ss.toast(String(msg), 'Import', seconds); }
  catch (e) { /* ignore when running unbound */ }
}

// -------------------- Main entry ------------------------
function Import_FromExcel_PreserveFormatting() {
  const t0 = Date.now();
  log_('INFO', '=== START Import_FromExcel_PreserveFormatting ===');

  // --- 1) Get the source Excel file & convert to a temporary Google Sheet ---
  const srcFile = DriveApp.getFileById(SETTINGS.SOURCE_XLSX_FILE_ID);
  const srcName = srcFile.getName();
  log_('INFO', `Source Excel: "${srcName}"`);

  const blob = srcFile.getBlob();

  // Try to put the converted file in the same folder (nice-to-have)
  let parentId = null;
  try {
    const parents = srcFile.getParents();
    if (parents.hasNext()) parentId = parents.next().getId();
  } catch (e) {
    log_('WARN', `Could not get parent folder; converted file goes to Drive root. ${e}`);
  }

  // v3 vs v2 adaptive conversion:
  let convertedTempId = null;
  if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.create === 'function') {
    // Drive API v3
    const resourceV3 = { name: `[Converted] ${srcName}`, mimeType: MimeType.GOOGLE_SHEETS };
    if (parentId) resourceV3.parents = [parentId];
    const converted = Drive.Files.create(resourceV3, blob, { supportsAllDrives: true });
    convertedTempId = converted.id;
    log_('INFO', `Converted (v3) → temp file: ${convertedTempId}`);
  } else if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.insert === 'function') {
    // Drive API v2
    const resourceV2 = { title: `[Converted] ${srcName}`, mimeType: MimeType.GOOGLE_SHEETS };
    if (parentId) resourceV2.parents = [{ id: parentId }];
    const converted = Drive.Files.insert(resourceV2, blob, { convert: true, supportsAllDrives: true });
    convertedTempId = converted.id;
    log_('INFO', `Converted (v2) → temp file: ${convertedTempId}`);
  } else {
    throw new Error('Advanced Drive Service not enabled. Add "Drive API" in Services, then retry.');
  }

  // --- 2) Open the converted sheet and read the worksheet values ---
  const workingSS = SpreadsheetApp.openById(convertedTempId);
  const srcTab = workingSS.getSheetByName(SETTINGS.SOURCE_WORKSHEET_NAME);
  if (!srcTab) throw new Error(`Worksheet "${SETTINGS.SOURCE_WORKSHEET_NAME}" not found in converted file.`);
  const values = srcTab.getDataRange().getValues();
  if (!values || values.length === 0) throw new Error('Source worksheet has no data.');
  log_('INFO', `Read ${values.length} rows from worksheet "${SETTINGS.SOURCE_WORKSHEET_NAME}"`);

  // --- 3) Resolve headers & indexes (why: safer than positions) ---
  const header = values[0];
  const need = ['Region','ItemID','BuyOrder','weightedaverage','maxval','minval','stddev','median','volume','numorders','fivepercent','orderSet'];
  const idx = {};
  need.forEach((name) => {
    const i = header.indexOf(name);
    if (i === -1) throw new Error(`Required column "${name}" not found. Headers: ${JSON.stringify(header)}`);
    idx[name] = i;
  });
  log_('INFO', `Header indexes: ${JSON.stringify(idx)}`);

  // --- 4) Filter Region and split by BuyOrder (TRUE-ish / FALSE-ish) ---
  const outTrue = [header];
  const outFalse = [header];

  let scanned = 0, keptT = 0, keptF = 0, skippedBO = 0;
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    scanned++;

    const region = String(row[idx.Region]).trim();
    if (region !== SETTINGS.TARGET_REGION) continue;

    const bo = String(row[idx.BuyOrder]).trim().toLowerCase();
    const isTrue  = (bo === 'true' || bo === '1' || bo === 'yes');
    const isFalse = (bo === 'false' || bo === '0' || bo === 'no');

    if (isTrue) { outTrue.push(row); keptT++; }
    else if (isFalse) { outFalse.push(row); keptF++; }
    else { skippedBO++; }
  }

  log_('INFO', `Filter summary: Region=${SETTINGS.TARGET_REGION}, scanned=${scanned}, TRUE=${keptT}, FALSE=${keptF}, skipped=${skippedBO}`);
  toast_(`Filtered: TRUE=${keptT}, FALSE=${keptF}`);

  // --- 5) Open MASTER and write while preserving formatting ---
  const masterSS = SpreadsheetApp.openById(SETTINGS.MASTER_SHEET_ID);

  // We DO NOT delete sheets; we only clear & overwrite the data block.
  // This preserves existing formatting, filters, column widths, notes, etc.
  upsertValuesPreserveFormatting_(
    masterSS,
    SETTINGS.TRUE_TAB_NAME,
    outTrue,
    SETTINGS.TRUE_NAMED_RANGE
  );

  upsertValuesPreserveFormatting_(
    masterSS,
    SETTINGS.FALSE_TAB_NAME,
    outFalse,
    SETTINGS.FALSE_NAMED_RANGE
  );

  // --- 6) Cleanup converted temp file if requested ---
  if (SETTINGS.DELETE_TEMP_CONVERTED_FILE && convertedTempId) {
    try {
      DriveApp.getFileById(convertedTempId).setTrashed(true);
      log_('INFO', `Deleted temporary converted file ${convertedTempId}`);
    } catch (e) {
      log_('WARN', `Could not delete temporary converted file ${convertedTempId}: ${e}`);
    }
  }

  log_('INFO', `=== END Import_FromExcel_PreserveFormatting in ${((Date.now() - t0) / 1000).toFixed(1)}s ===`);
  toast_('Import done.');
}

// -------------------- Writer (preserve formatting) ---------------------
/**
 * Write values into an existing or new tab, preserving formatting:
 * - If the sheet exists: we clear only the target data area and write values.
 *   We DO NOT call clear() or delete the sheet, to keep formatting/filters.
 * - If the sheet does not exist: we create it and write values.
 * - We refresh a Named Range to point at the data area excluding the header.
 *
 * Why exclude the header in the named range:
 * - Formulas/pivots/charts often want just the data rows; header is stable.
 */
function upsertValuesPreserveFormatting_(ss, sheetName, values, namedRangeName) {
  if (!values || values.length === 0) {
    log_('WARN', `No rows to write for "${sheetName}"`);
    return;
  }

  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    // Create if missing (first run). Formatting starts blank; user can style once.
    log_('INFO', `Sheet "${sheetName}" not found → creating`);
    sh = ss.insertSheet(sheetName);
  } else {
    log_('INFO', `Sheet "${sheetName}" exists → preserving formatting`);
  }

  const rows = values.length;
  const cols = values[0].length;

  // Ensure the sheet has enough rows/cols; we only INSERT (never delete)
  // to avoid shifting/removing existing formatting/named ranges.
  if (sh.getMaxRows() < rows) {
    const toAdd = rows - sh.getMaxRows();
    sh.insertRowsAfter(sh.getMaxRows(), toAdd);
    log_('DEBUG', `Inserted ${toAdd} row(s) to fit data in "${sheetName}"`);
  }
  if (sh.getMaxColumns() < cols) {
    const toAdd = cols - sh.getMaxColumns();
    sh.insertColumnsAfter(sh.getMaxColumns(), toAdd);
    log_('DEBUG', `Inserted ${toAdd} column(s) to fit data in "${sheetName}"`);
  }

  // Clear contents ONLY in the target write region (keeps formatting & filters)
  const target = sh.getRange(1, 1, rows, cols);
  target.clearContent();

  // Write the new data in a single bulk setValues() call
  target.setValues(values);

  // OPTIONAL: If there are stale contents outside the new block (old longer run),
  // we can clear the remainder below (contents only) for cleanliness, without
  // touching formatting. This is safe and keeps your layouts intact.
  if (sh.getMaxRows() > rows) {
    const extraRows = sh.getMaxRows() - rows;
    sh.getRange(rows + 1, 1, extraRows, sh.getMaxColumns()).clearContent();
  }
  if (sh.getMaxColumns() > cols) {
    const extraCols = sh.getMaxColumns() - cols;
    sh.getRange(1, cols + 1, rows, extraCols).clearContent();
  }

  // Maintain/refresh Named Range to data rows ONLY (exclude header row):
  // Data area = A2 : (rows x cols) starting at row 2
  if (namedRangeName) {
    const dataRowCount = Math.max(0, rows - 1);
    const dataRange = dataRowCount > 0
      ? sh.getRange(2, 1, dataRowCount, cols)
      : sh.getRange(2, 1, 1, cols); // minimal valid range even if empty

    upsertNamedRange_(ss, namedRangeName, dataRange);
    log_('INFO', `Named Range "${namedRangeName}" set to ${dataRange.getA1Notation()} on "${sheetName}"`);
  }

  log_('INFO', `Wrote ${rows} rows (incl header) into "${sheetName}" while preserving formatting`);
}

// -------------------- Named Range helper -------------------------------
/**
 * Create or update a named range on the given spreadsheet.
 * - If the named range exists → setRange(...) to the new range.
 * - Else → addNamedRange(...).
 */
function upsertNamedRange_(ss, name, range) {
  const existing = ss.getNamedRanges().filter(nr => nr.getName() === name);
  if (existing.length > 0) {
    existing[0].setRange(range);
  } else {
    ss.addNamedRange(name, range);
  }
}
