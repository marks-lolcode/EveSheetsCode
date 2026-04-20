/**
 * AllInv_Move_WriteInPlace
 *
 * WHAT:
 *   Write data from "AllInventory__STAGING" directly into "AllInventory"
 *   with minimal recalculation churn. No sheet renames, no named-range edits.
 *
 * WHY this order:
 *   - Clear contents first (keeps your formats/filters).
 *   - Grow only (insert rows/cols if needed) so we avoid costly pre-delete recalcs.
 *   - Single setValues() write → one main recalc.
 *   - Shrink extras after the write (cheaper than before).
 */
function AllInv_Move_WriteInPlace() {
  // Beginner-friendly logger: shows up in Executions/Logs for web calls.
  const log = (msg, data) => {
    try { console.log(msg, data === undefined ? '' : JSON.stringify(data)); } catch (_) {}
  };

  const STAGING = 'AllInventory__STAGING';
  const LIVE    = 'AllInventory';

  log('[AllInv] Move start (write-in-place)');

  // 1) Find sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const staging = ss.getSheetByName(STAGING);
  if (!staging) {
    log('[AllInv] ERROR: staging sheet missing', { STAGING });
    throw new Error('Staging sheet not found: ' + STAGING);
  }
  const live = ss.getSheetByName(LIVE) || ss.insertSheet(LIVE);

  // 2) Read used block from staging (WHAT: data to copy; WHY: one big write later)
  const rows = staging.getLastRow();
  const cols = staging.getLastColumn();
  if (!rows || !cols) {
    log('[AllInv] Staging is empty; aborting to protect live sheet.');
    return { ok: false, reason: 'empty_staging' };
  }
  const values = staging.getRange(1, 1, rows, cols).getValues();
  log('[AllInv] Read staging', { rows, cols });

  // 3) Clear ONLY contents on live (WHY: keep formats/filters/views)
  live.clearContents();
  log('[AllInv] Cleared live contents');

  // 4) Grow capacity as needed (WHY: avoid pre-write deletes which trigger heavy recalcs)
  ensureCapacityGrowOnly_(live, rows, cols, log);
  log('[AllInv] Capacity ensured', { maxRows: live.getMaxRows(), maxCols: live.getMaxColumns() });

  // 5) Single write (WHY: one big recalc instead of many small ones)
  live.getRange(1, 1, rows, cols).setValues(values);
  SpreadsheetApp.flush();
  log('[AllInv] Wrote values', { rows, cols });

  // 6) Shrink extras AFTER the write (WHY: cheaper than doing it before)
  shrinkExtra_(live, rows, cols, log);
  log('[AllInv] Shrunk grid (post-write)', { maxRows: live.getMaxRows(), maxCols: live.getMaxColumns() });

  // Keep staging hidden so users don’t edit it by accident
  try { staging.hideSheet(); } catch (_) {}

  log('[AllInv] Move done (write-in-place)');
  return { ok: true, rows, cols, method: 'write_in_place' };
}

/**
 * Grow capacity only (insert rows/cols if needed). Do NOT delete here.
 * BEGINNER: Deleting before writing causes big recalculations. We shrink later.
 */
function ensureCapacityGrowOnly_(sheet, neededRows, neededCols, log) {
  const curRows = sheet.getMaxRows();
  const curCols = sheet.getMaxColumns();
  if (curRows < neededRows) {
    sheet.insertRowsAfter(curRows, neededRows - curRows);
    if (log) log('[AllInv] Rows expanded', { from: curRows, to: neededRows });
  }
  if (curCols < neededCols) {
    sheet.insertColumnsAfter(curCols, neededCols - curCols);
    if (log) log('[AllInv] Cols expanded', { from: curCols, to: neededCols });
  }
}

/**
 * Shrink extra rows/cols AFTER the data is written (cheaper overall).
 */
function shrinkExtra_(sheet, rows, cols, log) {
  const mr = sheet.getMaxRows();
  const mc = sheet.getMaxColumns();
  if (mr > rows) {
    sheet.deleteRows(rows + 1, mr - rows);
    if (log) log('[AllInv] Rows shrunk', { from: mr, to: rows });
  }
  if (mc > cols) {
    sheet.deleteColumns(cols + 1, mc - cols);
    if (log) log('[AllInv] Cols shrunk', { from: mc, to: cols });
  }
}
