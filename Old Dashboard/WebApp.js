// ==========================================================
// FILE: WebApp_Core.gs
// WHAT: Router + endpoints the PowerShell script calls,
//        plus core helpers used by both web app and UI.
// WHY:  Keep all server-facing logic in one place.
// ==========================================================

/* =========================
   === CONFIG CONSTANTS ===
   ========================= */

// Bump this when you deploy so you can see which version is live.
const BUILD_TAG = 'upload-rows-2025-09-20T';

// Shared secret your client (PowerShell) must include in POST bodies.
const TOKEN = 'albatross-dreamland-oxidant-abstract';

// If you still use the server-side (Drive) importer:
const CSV_NAME = 'assets_export.csv';
const CSV_FOLDER_ID = '1IR1yvlvx0ZjIpVXmcKcubruG0n4ChlcS';
const CSV_FILE_ID = ''; // leave blank to auto-find by folder+name

// Target spreadsheet and primary/production sheet.
const TARGET_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const TARGET_SHEET_NAME  = 'AllInventory';

// NEW: Hidden staging sheet used to avoid expensive recalcs while importing.
const STAGING_SHEET_NAME = 'AllInventory__staging'; // create/hide on demand

// Include buffered logs in JSON responses (handy for PowerShell).
const RETURN_LOGS = true;

// Default rows per chunk for legacy server-side chunking.
const DEFAULT_CHUNK_ROWS = 5000;

/* =========================
   === LOGGING HELPERS   ===
   ========================= */

function newLogger() {
  // WHAT: Capture logs into a buffer we can also return in JSON.
  // WHY:  Lets the client view step-by-step detail without opening Apps Script logs.
  const buf = [];
  return {
    log: (msg, data) => {
      try {
        const line = (data === undefined) ? String(msg) : `${msg} ${JSON.stringify(data)}`;
        buf.push(line);
        console.log(line); // also to Cloud logs
      } catch (e) {
        const line = `${msg} [unserializable-data]`;
        buf.push(line);
        console.log(line);
      }
    },
    dump: () => buf.slice(),
  };
}

function json_(ok, fn, result, L, startedMs) {
  const ms = startedMs ? (Date.now() - startedMs) : 0;
  const out = { ok: !!ok, fn: String(fn || ''), ms, result: result || {} };
  if (RETURN_LOGS && L) out.logs = L.dump();
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =========================
   === DRIVE HELPERS     ===
   ========================= */

function findCsvFile_(L) {
  // WHAT: Prefer an explicit file ID; otherwise choose newest by folder+name.
  if (CSV_FILE_ID && CSV_FILE_ID.trim()) {
    try {
      const f = DriveApp.getFileById(CSV_FILE_ID.trim());
      if (f && f.getName() === CSV_NAME) {
        L.log('[findCsvFile_] using CSV_FILE_ID', { id: CSV_FILE_ID, name: f.getName(), updated: f.getLastUpdated() });
        return f;
      }
      L.log('[findCsvFile_] CSV_FILE_ID name mismatch; falling back', { expected: CSV_NAME, got: f && f.getName() });
    } catch (err) {
      L.log('[findCsvFile_] CSV_FILE_ID lookup failed; falling back', { error: String(err) });
    }
  }

  const folder = DriveApp.getFolderById(CSV_FOLDER_ID);
  const files = folder.getFilesByName(CSV_NAME);
  let newest = null;
  while (files.hasNext()) {
    const f = files.next();
    if (!newest || f.getLastUpdated() > newest.getLastUpdated()) newest = f;
  }
  if (!newest) throw new Error('CSV not found by ID or folder/name.');
  L.log('[findCsvFile_] found by folder+name', { id: newest.getId(), updated: newest.getLastUpdated() });
  return newest;
}

function readCsvRows_(file, L) {
  // WHAT: Read entire CSV → 2D array for setValues.
  const txt  = file.getBlob().getDataAsString();
  const rows = Utilities.parseCsv(txt);
  L.log('[readCsvRows_]', { rows: rows.length, cols: rows[0] ? rows[0].length : 0 });
  return rows;
}

/* =========================
   === SHEET HELPERS     ===
   ========================= */

/**
 * Get or create a sheet by name (optionally keep it hidden).
 * BEGINNER NOTE: We create the staging sheet once and keep it hidden so that
 * import writes do not trigger visible recalculations on the live tab.
 */
function getOrCreateSheet_(ss, name, makeHidden, L) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    L && L.log('[getOrCreateSheet_]', { created: name });
  }
  if (makeHidden && !sh.isSheetHidden()) {
    sh.hideSheet();
    L && L.log('[getOrCreateSheet_]', { hid: name });
  }
  return sh;
}

function clearSheet_(sheet, L) {
  sheet.clearContents();
  L && L.log('[clearSheet_]', { sheet: sheet.getName() });
}

function ensureCapacity_(sheet, neededRows, neededCols, L) {
  // WHAT: Pre-sizing prevents implicit grid growth during writes, which is slower.
  const curRows = sheet.getMaxRows();
  const curCols = sheet.getMaxColumns();
  if (curRows < neededRows) {
    sheet.insertRowsAfter(curRows, neededRows - curRows);
    L && L.log('[ensureCapacity_] rows expanded', { from: curRows, to: neededRows });
  }
  if (curCols < neededCols) {
    sheet.insertColumnsAfter(curCols, neededCols - curCols);
    L && L.log('[ensureCapacity_] cols expanded', { from: curCols, to: neededCols });
  }
}

/**
 * Resize a sheet’s grid to match exactly (no extra rows/cols left).
 * WHY: Commit with an exact size to avoid downstream recalcs across large blank areas.
 */
function resizeGridExactly_(sheet, rows, cols, L) {
  // Grow first (if needed), then shrink by deleting extras.
  ensureCapacity_(sheet, rows, cols, L);

  const curMaxRows = sheet.getMaxRows();
  const curMaxCols = sheet.getMaxColumns();

  if (curMaxRows > rows) {
    const del = curMaxRows - rows;
    sheet.deleteRows(rows + 1, del);
    L && L.log('[resizeGridExactly_] rows shrunk', { from: curMaxRows, to: rows });
  }
  if (curMaxCols > cols) {
    const del = curMaxCols - cols;
    sheet.deleteColumns(cols + 1, del);
    L && L.log('[resizeGridExactly_] cols shrunk', { from: curMaxCols, to: cols });
  }
}

/**
 * Write a 2D block starting at (1-based) row/col.
 */
function writeBlock_(sheet, values, startRow, startCol, L) {
  if (!values || !values.length) return { rowsWritten: 0 };
  const numRows = values.length;
  const numCols = values[0].length || 1;
  sheet.getRange(startRow, startCol, numRows, numCols).setValues(values);
  L && L.log('[writeBlock_]', { startRow, startCol, numRows, numCols });
  return { rowsWritten: numRows };
}

/**
 * Compute the used size (last row/col) for a sheet quickly.
 */
function usedSize_(sheet) {
  const rows = Math.max(1, sheet.getLastRow());
  const cols = Math.max(1, sheet.getLastColumn());
  return { rows, cols };
}

/* =========================
   === IMPORT: STAGING   ===
   ========================= */

/**
 * NEW: Clear the staging sheet before a fresh import.
 */
function beginImportStaging_(L) {
  const staging = getOrCreateSheet_(TARGET_SPREADSHEET, STAGING_SHEET_NAME, true, L);
  clearSheet_(staging, L);
  SpreadsheetApp.flush();
  return { ok: true, staging: STAGING_SHEET_NAME };
}

/**
 * NEW: Write rows into staging (client provides a 2D array).
 * BEGINNER NOTE: We pre-size the grid once, then write. This avoids repeated resizing.
 */
function importRowsToStaging_(L, body) {
  const staging = getOrCreateSheet_(TARGET_SPREADSHEET, STAGING_SHEET_NAME, true, L);
  const startIndex = Math.max(0, Number(body.startIndex || 0)); // 0-based
  const values = body && body.values;
  if (!values || !values.length || !Array.isArray(values[0])) {
    throw new Error('Missing/invalid "values"; expected 2D array.');
  }
  const neededRows = startIndex + values.length;
  const neededCols = values[0].length || 1;
  ensureCapacity_(staging, neededRows, neededCols, L);

  const wrote = writeBlock_(staging, values, startIndex + 1, 1, L).rowsWritten;
  SpreadsheetApp.flush();
  return { ok: true, wrote, nextIndex: startIndex + wrote };
}

/**
 * NEW: Single-call fast path: upload all rows to staging, then (optionally) swap.
 * Pass { values: [...], blockRows?: number, swap?: boolean }
 */
function importAllToStaging_(L, body) {
  const staging = getOrCreateSheet_(TARGET_SPREADSHEET, STAGING_SHEET_NAME, true, L);

  const values = body && body.values;
  if (!values || !values.length || !Array.isArray(values[0])) {
    throw new Error('Missing/invalid "values" (2D array).');
  }

  const BLOCK_ROWS = Number(body.blockRows || 2000);
  const totalRows  = values.length;
  const totalCols  = values[0].length || 1;

  // Start clean
  clearSheet_(staging, L);
  ensureCapacity_(staging, totalRows, totalCols, L);
  L.log('[importAllToStaging_] begin', { totalRows, totalCols, blockRows: BLOCK_ROWS });

  let written = 0;
  while (written < totalRows) {
    const remaining = totalRows - written;
    const take = Math.min(BLOCK_ROWS, remaining);
    const chunk = values.slice(written, written + take);
    const r1 = 1 + written;
    staging.getRange(r1, 1, chunk.length, totalCols).setValues(chunk);
    written += take;
    L.log('[importAllToStaging_] wrote block', { r1, rows: take, written, remaining: totalRows - written });
  }

  SpreadsheetApp.flush();
  L.log('[importAllToStaging_] done', { written });

  if (body.swap === true) {
    const swapped = swapStagingToAllInv_(L);
    return { ok: true, rows: written, swapped };
  }
  return { ok: true, rows: written, swapped: false };
}

/**
 * NEW: Atomic swap — commit the staging sheet into AllInventory in one go.
 * STEPS:
 *  1) Read used block from staging.
 *  2) Clear AllInventory.
 *  3) Resize AllInventory’s grid to exactly match staging’s used size.
 *  4) Write the block once.
 */
function swapStagingToAllInv_(L) {
  const staging = getOrCreateSheet_(TARGET_SPREADSHEET, STAGING_SHEET_NAME, true, L);
  const live    = TARGET_SPREADSHEET.getSheetByName(TARGET_SHEET_NAME);
  if (!live) throw new Error('Target sheet not found: ' + TARGET_SHEET_NAME);

  const used = usedSize_(staging);
  L.log('[swapStagingToAllInv_] staging used size', used);

  // Read values once
  const values = (used.rows > 0 && used.cols > 0)
    ? staging.getRange(1, 1, used.rows, used.cols).getValues()
    : [[]];

  // Prepare live sheet
  clearSheet_(live, L);
  resizeGridExactly_(live, used.rows, used.cols, L);

  // Commit in one shot
  if (used.rows > 0 && used.cols > 0) {
    live.getRange(1, 1, used.rows, used.cols).setValues(values);
  }
  SpreadsheetApp.flush();

  L.log('[swapStagingToAllInv_] committed', { rows: used.rows, cols: used.cols });
  return { ok: true, rows: used.rows, cols: used.cols };
}

/* =========================
   === IMPORT: LEGACY    ===
   =========================
   To preserve compatibility:
   - updatePIFullness_ now imports CSV → STAGING → SWAP
   - updatePIFullnessChunked_ still exists; also writes to staging then swap when done.
*/

function updatePIFullness_(L) {
  // CSV → memory → staging → swap
  const file  = findCsvFile_(L);
  const rows  = readCsvRows_(file, L);

  beginImportStaging_(L);
  // Write all to staging in blocks (reuse the importer).
  importAllToStaging_(L, { values: rows, blockRows: 2000, swap: true });

  return { ok: true, rows: rows.length, via: 'staging+swap' };
}

function updatePIFullnessChunked_(L, body) {
  // NOTE: This is kept for UI chunk demo. Each call appends to staging.
  const startIndex = Math.max(0, Number(body.startIndex || 0));
  const chunkRows  = Math.max(1, Number(body.chunkRows || DEFAULT_CHUNK_ROWS));
  const clear      = !!body.clear;

  const file      = findCsvFile_(L);
  const rows      = readCsvRows_(file, L);
  const totalRows = rows.length;

  if (clear || startIndex === 0) beginImportStaging_(L);

  const endExclusive = Math.min(totalRows, startIndex + chunkRows);
  const slice        = rows.slice(startIndex, endExclusive);
  if (slice.length) importRowsToStaging_(L, { startIndex, values: slice });

  const done      = endExclusive >= totalRows;
  const nextIndex = done ? totalRows : endExclusive;

  if (done) swapStagingToAllInv_(L);

  SpreadsheetApp.flush();
  return { ok: true, totalRows, wrote: slice.length, startIndex, nextIndex, done, via: 'staging+swap' };
}

/* =========================
   === HTTP HANDLERS     ===
   ========================= */

function doGet(e) {
  const L = newLogger();
  const fn = (e && e.parameter && (e.parameter.fn || e.parameter.op)) || 'ping';
  if (fn === 'ping') {
    return json_(true, 'ping', { pong: true, build: BUILD_TAG, now: new Date().toISOString(), method: 'GET' }, L);
  }
  return json_(true, 'echo', { params: e ? e.parameter : {} }, L);
}

function doPost(e) {
  const L = newLogger();
  const started = Date.now();

  // Parse JSON
  let body = {};
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    body = JSON.parse(raw);
  } catch (err) {
    return json_(false, 'parseError', { error: 'Invalid JSON body' }, L, started);
  }

  // Token
  if (!body || body.token !== TOKEN) {
    return json_(false, 'forbidden', { error: 'Forbidden' }, L, started);
  }

  const fn = body.fn || body.op || 'ping';

  try {
    switch (fn) {
      // Diagnostics
      case 'ping': return json_(true, fn, { pong: true, build: BUILD_TAG, now: new Date().toISOString(), method: 'POST' }, L, started);
      case 'debugFindCsv': {
        const f = findCsvFile_(L);
        return json_(true, fn, { id: f.getId(), name: f.getName(), size: f.getSize(), lastUpdated: f.getLastUpdated(), url: 'https://drive.google.com/open?id=' + f.getId() }, L, started);
      }

      // Timestamps
      case 'setNamedNow':
      case 'set_named_now': {
        const name = String(body.name || '').trim();
        if (!name) return json_(false, fn, { error: 'missing_name' }, L, started);
        const r  = TARGET_SPREADSHEET.getRangeByName(name);
        if (!r) return json_(false, fn, { error: 'named_range_not_found', name }, L, started);
        r.setValue(new Date());
        return json_(true, fn, { ok: true, name }, L, started);
      }

      // Legacy CSV importers (now route through staging+swap)
      case 'updatePIFullness': {
        const r = updatePIFullness_(L);
        return json_(true, fn, r, L, started);
      }
      case 'updatePIFullnessChunked': {
        const r = updatePIFullnessChunked_(L, body);
        return json_(true, fn, r, L, started);
      }

      // Fast upload path (client-parsed CSV)
      case 'beginImport': {
        const r = beginImportStaging_(L);
        return json_(true, fn, r, L, started);
      }
      case 'importRows': {
        const r = importRowsToStaging_(L, body);
        return json_(true, fn, r, L, started);
      }
      case 'importAll': {
        const r = importAllToStaging_(L, body);
        return json_(true, fn, r, L, started);
      }

      // NEW: explicit finalize step if you want a separate call
      case 'swapStagingToAllInv': {
        const r = swapStagingToAllInv_(L);
        return json_(true, fn, r, L, started);
      }

// In doPost(e) switch(fn) { ... } add/replace this case:
case 'refreshAllInvImport': {
  // BEGINNER: Some runs need to poke a “middle” sheet in another file.
  // If the function isn’t present in this project, we just no-op and return ok=true
  // so the pipeline can continue to the move/swap step.
  let ok = true, detail = { called: false };
  try {
    if (typeof RefreshAllInvImport === 'function') {
      console.log('[refreshAllInvImport] calling project function RefreshAllInvImport()');
      RefreshAllInvImport(); // your optional custom routine
      detail.called = true;
    } else {
      console.log('[refreshAllInvImport] no-op; function not present');
    }
    SpreadsheetApp.flush();
  } catch (err) {
    ok = false;
    detail.error = String(err);
    console.error('[refreshAllInvImport] error', err);
  }
  return json_(ok, fn, { ok, detail }, L, started);
}

case 'allinv_move': {
  // BEGINNER NOTE: We now write directly into AllInventory (no renames, no named-range changes).
  // We also return a 'detail' object so your PS logs show rows/cols and method used.
  let moved = false;
  let detail = null;

  try {
    // Call the new in-place writer (make sure AllInv_Move_WriteInPlace is defined elsewhere).
    detail = AllInv_Move_WriteInPlace(); // <-- the new function you added
    moved = !!(detail && detail.ok);
    L.log('[allinv_move] completed', { moved, detail });
  } catch (err) {
    // WHY: Bubble errors into logs and response so your scripts can see what happened.
    const msg = String(err && err.stack || err);
    L.log('[allinv_move] error', { error: msg });
    detail = { ok: false, error: msg };
  }

  // Flush any pending sheet operations for safety.
  SpreadsheetApp.flush();

  // Optional: stamp a named range so you can see last-run time in the sheet
  try {
    const rng = TARGET_SPREADSHEET.getRangeByName('allinv_move');
    if (rng) rng.setValue(new Date());
  } catch (err) {
    L.log('[allinv_move] stamp failed', { error: String(err) });
  }

  // Return richer info so your PowerShell logs are more helpful.
  return json_(true, fn, { moved, detail }, L, started);
}


      default: return json_(false, fn, { error: 'Unknown function' }, L, started);
    }
  } catch (err) {
    return json_(false, fn, { error: String(err) }, L, started);
  }
}
