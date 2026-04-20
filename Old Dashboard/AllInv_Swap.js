/**
 * AllInv_Move
 *
 * WHAT:
 *   Atomically swap the latest uploaded data from a hidden staging sheet
 *   (AllInventory__STAGING) into the live AllInventory sheet with minimal
 *   recalculation churn.
 *
 * WHY:
 *   Writing once into the live sheet (after it’s cleared and right-sized)
 *   avoids triggering formulas thousands of times during a row-by-row copy.
 *
 * HOW:
 *   1) Read the full used range from STAGING into memory.
 *   2) Clear AllInventory, resize grid to match the staging data’s size.
 *   3) Write all values in a single setValues() call.
 *   4) Flush, then (optionally) trim the grid to keep workbook size tidy.
 *
 * LOGGING:
 *   Uses the project’s newLogger() if available; otherwise falls back to console.
 */
function AllInv_Move() {
  // --- logger setup (beginner-friendly) ---
  // We try to reuse newLogger() from WebApp.gs so logs show up in web responses.
  // If not available, simple console logs still give you visibility in Executions → Logs.
  const L = (typeof newLogger === 'function')
    ? newLogger()
    : { log: (m, d) => console.log(d === undefined ? String(m) : (String(m) + ' ' + JSON.stringify(d))) };

  const t0 = Date.now();
  L.log('[Move] Start');

  // --- configuration (what/why) ---
  // We explicitly name the staging and live tabs. Using fixed names keeps the
  // web-app/PowerShell pipeline simple and predictable.
  const STAGING_SHEET_NAME = 'AllInventory__STAGING'; // hidden temp tab populated by import
  const LIVE_SHEET_NAME    = 'AllInventory';          // visible, formula-bearing tab

  // --- get sheets (why: fail fast with clear messages if something is misnamed) ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staging = ss.getSheetByName(STAGING_SHEET_NAME);
  if (!staging) {
    L.log('[Move] ERROR: staging sheet not found', { sheet: STAGING_SHEET_NAME });
    throw new Error('Staging sheet not found: ' + STAGING_SHEET_NAME);
  }
  const live = ss.getSheetByName(LIVE_SHEET_NAME) || ss.insertSheet(LIVE_SHEET_NAME);

  // --- read used range from staging (what/why) ---
  // We only read the area that actually has content (getLastRow/Column),
  // which is faster and prevents growing the live sheet unnecessarily.
  const usedRows = Math.max(1, staging.getLastRow());
  const usedCols = Math.max(1, staging.getLastColumn());
  if (usedRows === 1 && usedCols === 1 && String(staging.getRange(1, 1).getValue()).trim() === '') {
    // Staging appears empty. We avoid blanking the live sheet in this edge case.
    L.log('[Move] Staging appears empty; aborting to protect live sheet.', { usedRows, usedCols });
    return { ok: false, reason: 'empty_staging' };
  }

  const values = staging.getRange(1, 1, usedRows, usedCols).getValues();
  L.log('[Move] Read staging values', { rows: usedRows, cols: usedCols });

  // --- clear live + right-size once (why: one resize minimizes recalc/overhead) ---
  live.clearContents(); // clears cells but preserves formatting; deliberate to keep existing formats
  // If you prefer to *remove* existing formats, use clear() instead of clearContents().

  // Resize grid to exactly match used data. This prevents incremental auto-expansion,
  // which can be slow and cause repeated recalculations of dependent formulas.
  const curMaxRows = live.getMaxRows();
  const curMaxCols = live.getMaxColumns();
  if (curMaxRows < usedRows) live.insertRowsAfter(curMaxRows, usedRows - curMaxRows);
  if (curMaxCols < usedCols) live.insertColumnsAfter(curMaxCols, usedCols - curMaxCols);
  if (curMaxRows > usedRows) live.deleteRows(usedRows + 1, curMaxRows - usedRows);
  if (curMaxCols > usedCols) live.deleteColumns(usedCols + 1, curMaxCols - usedCols);
  L.log('[Move] Resized live grid', { rows: usedRows, cols: usedCols });

  // --- single write (what/why) ---
  // One setValues() → one big recalc instead of thousands of tiny ones.
  live.getRange(1, 1, usedRows, usedCols).setValues(values);
  SpreadsheetApp.flush();
  L.log('[Move] Wrote live values', { rows: usedRows, cols: usedCols, elapsed_ms: Date.now() - t0 });

  // Optional: keep staging hidden so accidental edits don’t happen.
  try { staging.hideSheet(); } catch (e) { /* ignore if already hidden */ }

  // Optional: very small safety buffer so users can type a bit at the bottom/right
  // without growing the grid immediately (feel free to adjust or remove).
  const safetyRows = Math.max(usedRows + 50, 200);
  const safetyCols = Math.max(usedCols + 5, 26);
  const finalRows = Math.max(safetyRows, live.getMaxRows());
  const finalCols = Math.max(safetyCols, live.getMaxColumns());
  if (finalRows > live.getMaxRows()) live.insertRowsAfter(live.getMaxRows(), finalRows - live.getMaxRows());
  if (finalCols > live.getMaxColumns()) live.insertColumnsAfter(live.getMaxColumns(), finalCols - live.getMaxColumns());
  L.log('[Move] Added tiny safety buffer', { rows: finalRows, cols: finalCols });

  L.log('[Move] Done', { total_ms: Date.now() - t0 });
  return { ok: true, rows: usedRows, cols: usedCols };
}


function _test_AllInv_Move() {
  const result = AllInv_Move();
  Logger.log(JSON.stringify(result));
  console.log(JSON.stringify(result));
}