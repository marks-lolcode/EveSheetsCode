/**
 * Web App entrypoints with shared dispatcher and robust logging.
 * WHAT: Accepts either JSON POST or URL query GET.
 * WHY: Easier testing (GET in browser) and reliable headless calls (POST).
 *
 * SECURITY: Uses a shared secret token. Keep it private.
 * IMPORTANT: TOKEN_VALUE must match what your PowerShell script sends.
 */

function doPost(e) {
  // POST from PowerShell or other clients
  return _handleRequest_(e, 'POST');
}

function doGet(e) {
  // GET from a browser for quick "ping" tests
  return _handleRequest_(e, 'GET');
}

function _handleRequest_(e, method) {
  Logger.log('[%s] Start. Has e=%s', method, !!e);

  // 1) Parse payload safely
  // If POST with application/json: read and parse JSON body.
  // Else: read query params (so ?fn=ping&token=... works in a browser).
  let payload = {};
  try {
    if (method === 'POST' &&
        e && e.postData &&
        e.postData.type &&
        String(e.postData.type).indexOf('application/json') !== -1) {

      const raw = (e.postData.contents != null)
        ? e.postData.contents
        : (e.postData.getDataAsString ? e.postData.getDataAsString() : '');

      payload = raw ? JSON.parse(raw) : {};
      Logger.log('[%s] Parsed JSON payload keys: %s', method, Object.keys(payload).join(','));
    } else {
      payload = e && e.parameter
        ? Object.keys(e.parameter).reduce((o, k) => (o[k] = e.parameter[k], o), {})
        : {};
      Logger.log('[%s] Parsed query payload keys: %s', method, Object.keys(payload).join(','));
    }
  } catch (err) {
    Logger.log('[%s] ERROR parsing payload: %s', method, err && err.message);
    return _json({ ok: false, error: 'Invalid JSON/params' });
  }

  // 2) Auth: shared-secret token
  // Replace with your exact token, and ensure your PowerShell script sends the same string.
  const TOKEN_VALUE = 'albatross-dreamland-oxidant-abstract';  // <<< SET THIS EXACTLY
  const incoming = String((payload.token || payload.Token || '')).trim();
  const expected = String(TOKEN_VALUE).trim();
  const match = (incoming === expected);

  Logger.log('[%s] Token check: incoming.len=%s expected.len=%s equal=%s',
             method, incoming.length, expected.length, match);

  if (!match) {
    return _json({ ok: false, error: 'Forbidden' });
  }

  // 3) Dispatch to function
  const fn = String(payload.fn || 'updatePIFullness');
  const t0 = Date.now();
  Logger.log('[%s] Dispatching fn=%s', method, fn);

  try {
    let result;

    if (fn === 'ping') {
      // Lightweight health check
      result = { pong: true, now: new Date().toISOString(), method };
    } else if (fn === 'updatePIFullness') {
      // Your two-step pipeline:
      // Step 1 writes the sheet, Step 2 reads/derives; flush ensures reads see writes.
      const a = importCSVFromDrive();          // Must exist in your project
      SpreadsheetApp.flush();                  // Ensure all writes are committed
      const b = filterInventoryByTypeFirst();  // Must exist in your project
      result = {
        importedRows: (a && a.length) || 0,
        processedRows: (b && b.length) || 0
      };
    } else {
      return _json({ ok: false, error: 'Unknown function' });
    }

    const ms = Date.now() - t0;
    Logger.log('[%s] Success fn=%s in %s ms', method, fn, ms);
    return _json({ ok: true, fn, ms, result });

  } catch (err) {
    Logger.log('[%s] ERROR running fn=%s: %s\nStack: %s', method, fn,
               err && err.message, err && err.stack);
    return _json({ ok: false, error: String(err && err.message || err) });
  }
}

/**
 * Small helper to return JSON consistently.
 * WHAT: Avoids repeating boilerplate; sets correct MIME type.
 */
function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}









/**
 * updatePIFullness(options)
 *
 * WHAT this does (high level):
 *   1) Runs your import step (importCSVFromDrive) to populate the source sheet.
 *   2) Forces pending writes to apply (SpreadsheetApp.flush) so the next step reads fresh data.
 *   3) Runs your processing step (filterInventoryByTypeFirst) to build the derived/output sheet.
 *   4) Logs each step with timings and row/column counts for easier debugging.
 *
 * WHY this design:
 *   - Web Apps and triggers run "headless" (no UI). Calls like SpreadsheetApp.getUi() will fail.
 *   - Inline logging shows WHAT is happening and WHY at every step, which is crucial when you
 *     cannot see the spreadsheet while the script runs.
 *   - Returning a small summary object gives your caller (PowerShell) a quick success snapshot.
 *
 * REQUIREMENTS:
 *   - importCSVFromDrive()   must exist and perform the import.
 *   - filterInventoryByTypeFirst() must exist and perform the processing.
 *   - Neither function should use UI APIs (alerts/prompts/toasts).
 *
 * @param {Object=} options                    // Optional, safe to omit.
 * @param {string=} options.importSheetName    // Sheet to inspect for import counts (default 'AllInventory').
 * @param {string=} options.processSheetName   // If provided, we will also inspect this sheet for final counts.
 * @param {boolean=} options.dryRun            // If true, skip the processing step (useful for testing import only).
 * @param {number=} options.flushWaitMs        // Optional extra wait after flush if you read very large ranges.
 * @return {Object} summary                    // { ok, importedRows, importedCols, processedRows, processedCols, msImport, msFlush, msProcess, msTotal, note? }
 */
function updatePIFullness(options) {
  // ---------------------------
  // 0) Read options with defaults
  // ---------------------------
  options = options || {};
  const importSheetName  = options.importSheetName  || 'AllInventory';
  const processSheetName = options.processSheetName || null;         // Leave null if unknown.
  const dryRun           = options.dryRun === true ? true : false;   // Default false.
  const flushWaitMs      = Number.isFinite(options.flushWaitMs) ? options.flushWaitMs : 0;

  // ---------------------------
  // 1) Start: set up timers and summary we will return to the caller
  // ---------------------------
  const tStart = Date.now();
  Logger.log('[updatePIFullness] Start. Options=%s', JSON.stringify({
    importSheetName, processSheetName, dryRun, flushWaitMs
  }));

  // We collect metrics as we go so the caller can see what happened without opening logs.
  let summary = {
    ok: true,
    importedRows: 0,
    importedCols: 0,
    processedRows: 0,
    processedCols: 0,
    msImport: 0,
    msFlush: 0,
    msProcess: 0,
    msTotal: 0,
    note: '' // we append short flags like 'import-fallback' or 'process-fallback' when we infer sizes from sheets
  };

  // Small helper used a few times: compute elapsed milliseconds since a timestamp
  const elapsed = (t0) => Date.now() - t0;

  try {
    // -------------------------------------------------------
    // 2) IMPORT STEP: write fresh data into the import sheet
    // -------------------------------------------------------
    // WHY: The processing step should always work off a known, up-to-date source.
    Logger.log('[updatePIFullness] Step 1/3: importCSVFromDrive() starting.');
    const tImport = Date.now();

    let importedReturn;  // Whatever the import function returns (could be undefined, 2D array, or a Range)
    try {
      importedReturn = importCSVFromDrive();
    } catch (err) {
      // BEGINNER NOTE: Catching here lets us log a very specific failure location.
      Logger.log('[updatePIFullness] ERROR in importCSVFromDrive(): %s', err && err.message);
      throw err; // Re-throw so the outer catch sets ok=false and returns a clear error to the caller.
    }

    summary.msImport = elapsed(tImport);
    Logger.log('[updatePIFullness] Step 1/3: import finished in %d ms.', summary.msImport);

    // Determine how many rows/columns the import produced so we can log and return it.
    // We try, in order:
    //   A) The function returned a 2D array of values (fastest to count).
    //   B) The function returned a Range (use getNumRows/getNumColumns).
    //   C) Fallback: inspect the sheet named importSheetName (getLastRow/getLastColumn).
    // The fallback is useful when the import function itself returns nothing.
    if (Array.isArray(importedReturn) && importedReturn.length > 0 && Array.isArray(importedReturn[0])) {
      // Case A: 2D array
      summary.importedRows = importedReturn.length;
      summary.importedCols = importedReturn[0].length;
      Logger.log('[updatePIFullness] Import size (from return array): %d rows x %d cols.',
                 summary.importedRows, summary.importedCols);
    } else if (importedReturn && typeof importedReturn.getNumRows === 'function' && typeof importedReturn.getNumColumns === 'function') {
      // Case B: Range
      summary.importedRows = Number(importedReturn.getNumRows()) || 0;
      summary.importedCols = Number(importedReturn.getNumColumns()) || 0;
      Logger.log('[updatePIFullness] Import size (from return Range): %d rows x %d cols.',
                 summary.importedRows, summary.importedCols);
    } else {
      // Case C: Fallback to sheet inspection
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss ? ss.getSheetByName(importSheetName) : null;
      if (sh) {
        summary.importedRows = Number(sh.getLastRow()) || 0;
        summary.importedCols = Number(sh.getLastColumn()) || 0;
        summary.note += 'import-fallback; ';
        Logger.log('[updatePIFullness] Import size (inferred from sheet "%s"): %d rows x %d cols.',
                   importSheetName, summary.importedRows, summary.importedCols);
      } else {
        Logger.log('[updatePIFullness] WARNING: Could not find sheet "%s" to infer import size.', importSheetName);
      }
    }

    // -------------------------------------------------------
    // 3) FLUSH: force pending writes to be applied before reading
    // -------------------------------------------------------
    // WHY: Apps Script batches writes for performance. flush() makes sure the data written
    // in the import step is actually committed before we read in the next step.
    const tFlush = Date.now();
    SpreadsheetApp.flush();
    if (flushWaitMs > 0) {
      // Optional small buffer if your processing step reads very large ranges immediately.
      Utilities.sleep(flushWaitMs);
    }
    summary.msFlush = elapsed(tFlush);
    Logger.log('[updatePIFullness] Flush complete in %d ms. Extra wait applied: %d ms.',
               summary.msFlush, flushWaitMs);

    // -------------------------------------------------------
    // 4) PROCESS STEP: transform imported data into derived output
    // -------------------------------------------------------
    if (dryRun) {
      // WHY: Useful when you want to test the import path without touching downstream sheets.
      Logger.log('[updatePIFullness] Step 2/3 skipped (dryRun=true). Returning after import + flush.');
      summary.msTotal = elapsed(tStart);
      Logger.log('[updatePIFullness] End (dryRun). Total: %d ms.', summary.msTotal);
      return summary;
    }

    Logger.log('[updatePIFullness] Step 2/3: filterInventoryByTypeFirst() starting.');
    const tProcess = Date.now();

    let processedReturn; // Whatever the processing function returns
    try {
      processedReturn = filterInventoryByTypeFirst();
    } catch (err) {
      Logger.log('[updatePIFullness] ERROR in filterInventoryByTypeFirst(): %s', err && err.message);
      throw err;
    }

    summary.msProcess = elapsed(tProcess);
    Logger.log('[updatePIFullness] Step 2/3: processing finished in %d ms.', summary.msProcess);

    // Determine the size of the processed result using the same logic as import.
    if (Array.isArray(processedReturn) && processedReturn.length > 0 && Array.isArray(processedReturn[0])) {
      summary.processedRows = processedReturn.length;
      summary.processedCols = processedReturn[0].length;
      Logger.log('[updatePIFullness] Processed size (from return array): %d rows x %d cols.',
                 summary.processedRows, summary.processedCols);
    } else if (processedReturn && typeof processedReturn.getNumRows === 'function' && typeof processedReturn.getNumColumns === 'function') {
      summary.processedRows = Number(processedReturn.getNumRows()) || 0;
      summary.processedCols = Number(processedReturn.getNumColumns()) || 0;
      Logger.log('[updatePIFullness] Processed size (from return Range): %d rows x %d cols.',
                 summary.processedRows, summary.processedCols);
    } else if (processSheetName) {
      // Fallback only if you tell us which sheet the processing step wrote to.
      const ss2 = SpreadsheetApp.getActiveSpreadsheet();
      const sh2 = ss2 ? ss2.getSheetByName(processSheetName) : null;
      if (sh2) {
        summary.processedRows = Number(sh2.getLastRow()) || 0;
        summary.processedCols = Number(sh2.getLastColumn()) || 0;
        summary.note += 'process-fallback; ';
        Logger.log('[updatePIFullness] Processed size (inferred from sheet "%s"): %d rows x %d cols.',
                   processSheetName, summary.processedRows, summary.processedCols);
      } else {
        Logger.log('[updatePIFullness] NOTE: processSheetName not found, unable to infer processed size.');
      }
    } else {
      Logger.log('[updatePIFullness] NOTE: processing function returned no size and no processSheetName was provided.');
    }

    // -------------------------------------------------------
    // 5) FINISH: record total runtime and return a concise summary
    // -------------------------------------------------------
    summary.msTotal = elapsed(tStart);
    Logger.log('[updatePIFullness] End. Total=%d ms | Imported=%dx%d | Processed=%dx%d',
               summary.msTotal,
               summary.importedRows, summary.importedCols,
               summary.processedRows, summary.processedCols);

    return summary;

  } catch (err) {
    // Centralized fail-safe: ensure the caller still receives a structured JSON-friendly object.
    summary.ok = false;
    summary.error = String(err && err.message || err);
    summary.msTotal = elapsed(tStart);
    Logger.log('[updatePIFullness] FATAL: %s', summary.error);
    Logger.log('[updatePIFullness] Aborted after %d ms. Imported=%dx%d Processed=%dx%d',
               summary.msTotal,
               summary.importedRows, summary.importedCols,
               summary.processedRows, summary.processedCols);
    return summary;
  }
}





function importCSVFromDrive() {
  // ------------------------------------------------------------
  // Simple logging helpers (WHAT: print to logs; WHY: easier debugging)
  // We keep these tiny to avoid changing behavior—just messages + timestamps.
  // ------------------------------------------------------------
  var ts  = function () { return new Date().toISOString(); };
  var log = function (msg) { console.log("[INFO  " + ts() + "] " + msg); Logger.log("[INFO ] " + msg); };
  var warn = function (msg) { console.log("[WARN  " + ts() + "] " + msg); Logger.log("[WARN ] " + msg); };
  var err = function (msg) { console.error("[ERROR " + ts() + "] " + msg); Logger.log("[ERROR] " + msg); };

  // ------------------------------------------------------------
  // ** User-defined variables **
  // WHAT: names that locate the folder/file and the target sheet.
  // WHY: keeping them up here makes the script easy to configure.
  // ------------------------------------------------------------
  var folderName = "jeveassets_exports"; // The name of the folder in your Google Drive
  var csvFileName = "assets_export.csv"; // The name of the exported CSV file
  var targetSheetName = "AllInventory"; // The name of the sheet to import into

  // Start timing so we can see total duration in logs (WHY: performance insight)
  var startedAtMs = Date.now();
  log("importCSVFromDrive: START");
  log("Config → folderName=\"" + folderName + "\", csvFileName=\"" + csvFileName + "\", targetSheetName=\"" + targetSheetName + "\"");

  // ------------------------------------------------------------
  // WHAT: get the Drive folder iterator by name and take the first match.
  // WHY: script expects a single folder with the given name.
  // ------------------------------------------------------------
  var folder = DriveApp.getFoldersByName(folderName).next();
  log('Folder located: name="' + folder.getName() + '", id=' + folder.getId());

  // ------------------------------------------------------------
  // WHAT: search for a file named csvFileName in that folder.
  // WHY: we only import if that exact file is present.
  // ------------------------------------------------------------
  var files = folder.getFilesByName(csvFileName);
  log('Searching for file "' + csvFileName + '" in folder "' + folderName + '"…');

  if (files.hasNext()) {
    var file = files.next();
    log('File found: name="' + file.getName() + '", id=' + file.getId() + ", size=" + file.getSize() + " bytes, lastUpdated=" + file.getLastUpdated());

    // ------------------------------------------------------------
    // WHAT: read CSV text and parse it into a 2D array.
    // WHY: Utilities.parseCsv produces the rows/columns we can write to the sheet.
    // ------------------------------------------------------------
    var tReadStart = Date.now();
    var csvText = file.getBlob().getDataAsString(); // default charset; unchanged to keep original behavior
    var csvData = Utilities.parseCsv(csvText);
    var readMs = Date.now() - tReadStart;

    // Defensive logging (WHY: quick check for empty/malformed CSV without changing behavior)
    if (!csvData || !csvData.length || !csvData[0]) {
      warn("Parsed CSV appears empty or malformed (no rows or columns). Continuing without data changes.");
    } else {
      log("CSV parsed: rows=" + csvData.length + ", cols=" + csvData[0].length + " (read+parse " + readMs + "ms)");
    }

    // ------------------------------------------------------------
    // WHAT: get the target sheet.
    // WHY: this is where we’ll write the CSV.
    // ------------------------------------------------------------
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
    if (!sheet) {
      err('Target sheet "' + targetSheetName + '" not found. Aborting before data write.');
      // We do not throw to avoid changing control flow; the original code assumed it exists.
      // Returning gracefully preserves original behavior expectations.
      log("importCSVFromDrive: END (no target sheet). Total " + (Date.now() - startedAtMs) + "ms");
      return;
    }
    log('Target sheet ready: name="' + sheet.getName() + '", sheetId=' + sheet.getSheetId());

    // ------------------------------------------------------------
    // WHAT: clear the existing data before import.
    // WHY: ensures the sheet exactly mirrors the CSV content.
    // ------------------------------------------------------------
    var tClearStart = Date.now();
    sheet.clearContents();
    var clearMs = Date.now() - tClearStart;
    log("Cleared existing sheet contents in " + clearMs + "ms");

    // ------------------------------------------------------------
    // WHAT: write the entire CSV data range starting at A1.
    // WHY: single bulk write is fastest and maintains structure.
    // ------------------------------------------------------------
    if (csvData && csvData.length && csvData[0]) {
      var tWriteStart = Date.now();
      sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
      var writeMs = Date.now() - tWriteStart;
      log("Wrote " + csvData.length + " rows × " + csvData[0].length + " cols to sheet in " + writeMs + "ms");
    } else {
      warn("Skipped write: csvData was empty or malformed.");
    }
  } else {
    // If we got here, the file wasn't present
    warn('File "' + csvFileName + '" not found in folder "' + folderName + '". Nothing imported.');
  }

  log("importCSVFromDrive: END (success path reached). Total " + (Date.now() - startedAtMs) + "ms");
}
